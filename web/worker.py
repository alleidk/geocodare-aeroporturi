"""Un singur fir de execuție care preia joburile din coadă, pe rând."""

import logging
import threading
import time
import traceback

from . import config, db, geocoding

logger = logging.getLogger("geocodare_web")

_wake = threading.Event()


def job_dir(job_id: str):
    return config.JOBS_DIR / job_id


def input_path(job_id: str):
    return job_dir(job_id) / "input.xlsx"


def output_path(job_id: str):
    return job_dir(job_id) / "output.xlsx"


def notify():
    _wake.set()


def start():
    # Joburile întrerupte de o repornire a serverului sunt reluate (cache-ul păstrează progresul)
    db.execute("UPDATE jobs SET status = 'queued', phase = NULL WHERE status = 'running'")
    threading.Thread(target=_loop, name="geocode-worker", daemon=True).start()


def _next_job():
    return db.query("SELECT * FROM jobs WHERE status = 'queued' ORDER BY updated_at LIMIT 1", one=True)


def _loop():
    last_purge = 0.0
    while True:
        if time.time() - last_purge > 86400:
            geocoding.purge_expired_cache()
            last_purge = time.time()

        job = _next_job()
        if job is None:
            _wake.wait(timeout=30)
            _wake.clear()
            continue
        _run(job)


def _run(job):
    job_id = job["id"]
    db.update_job(job_id, status="running", message=None, cancel_requested=0, done=0, total=0,
                  addresses=0, found=0)

    def should_cancel():
        row = db.query("SELECT cancel_requested FROM jobs WHERE id = ?", (job_id,), one=True)
        return row is None or bool(row["cancel_requested"])

    try:
        errors, summary = geocoding.run_job(job, input_path(job_id), output_path(job_id), should_cancel)
        db.update_job(job_id, status="done", phase=None, errors=errors, message=summary)
    except geocoding.JobCancelled:
        db.update_job(job_id, status="cancelled", phase=None,
                      message="Oprit la cerere. Adresele deja geocodate rămân în cache.")
    except geocoding.JobFailed as e:
        db.update_job(job_id, status="failed", phase=None, message=str(e))
    except Exception as e:
        logger.error("Job %s eșuat:\n%s", job_id, traceback.format_exc())
        db.update_job(job_id, status="failed", phase=None, message=f"Eroare neașteptată: {e}")
