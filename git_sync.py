"""
git_sync.py

Commits and pushes the Excel workbook to GitHub after a local sync, so
the Streamlit Community Cloud deployment -- which reads the workbook
straight out of its repo checkout and auto-redeploys on every push to
the watched branch -- shows data from your last local run.

Only ever called when config.IS_LOCAL_INSTANCE is true (the Mac running
the actual automation and holding real git push credentials). A
deployed cloud instance never calls this: its filesystem is ephemeral
and it has no git credentials, so a push attempt there would just fail
-- callers are expected to check IS_LOCAL_INSTANCE themselves before
calling, the same way they gate which UI controls are even shown.
"""

import logging
import subprocess

from config import BASE_DIR, EXCEL_FILE_PATH

logger = logging.getLogger(__name__)


def push_workbook_if_changed(commit_message: str = "Data sync") -> bool:
    """Commit and push the workbook if it changed since the last commit.

    Returns True if a push happened, False otherwise (nothing changed,
    or the push failed). Never raises -- a failed push (no network, no
    credentials, not a git checkout) is logged as a warning and the
    caller continues normally; the local workbook itself is unaffected
    either way, only the deployed copy would stay stale until the next
    successful push.
    """
    try:
        rel_path = str(EXCEL_FILE_PATH.relative_to(BASE_DIR))
    except ValueError:
        logger.warning("git_sync: workbook is outside the repo (%s) -- skipping push", EXCEL_FILE_PATH)
        return False

    try:
        status = subprocess.run(
            ["git", "status", "--porcelain", rel_path],
            cwd=BASE_DIR, capture_output=True, text=True, check=True,
        )
        if not status.stdout.strip():
            logger.info("git_sync: workbook unchanged, nothing to push")
            return False

        subprocess.run(["git", "add", rel_path], cwd=BASE_DIR, check=True, capture_output=True)
        subprocess.run(["git", "commit", "-m", commit_message], cwd=BASE_DIR, check=True, capture_output=True)
        subprocess.run(["git", "push"], cwd=BASE_DIR, check=True, capture_output=True, text=True)
        logger.info("git_sync: pushed updated workbook to GitHub")
        return True
    except subprocess.CalledProcessError as exc:
        stderr = exc.stderr.strip() if isinstance(exc.stderr, str) else exc.stderr
        logger.warning(
            "git_sync: failed to push workbook (%s) -- the deployed dashboard will show stale "
            "data until this is resolved (check `git push` works manually in this repo)",
            stderr or exc,
        )
        return False
    except FileNotFoundError:
        logger.warning("git_sync: git executable not found -- skipping push")
        return False
