"""
git_sync.py

Keeps the Excel workbook and the tracked-ticker list (data/tickers.json)
in sync with GitHub in both directions:

  - pull_latest(): fast-forwards the local checkout before a write
    action, so a long-running process (e.g. a cloud container that's
    been up for a while) mutates on top of the freshest known state
    instead of whatever it happened to start with -- otherwise a stale
    container could push a commit that clobbers a change made elsewhere
    (your Mac's scheduled sync, or another session) since it last booted.
  - push_state_if_changed(): commits and pushes the workbook + tickers
    file if either changed, so the Streamlit Community Cloud deployment
    (which reads them straight out of its repo checkout and auto-
    redeploys on every push) picks up fresh data.

Two push/pull paths, chosen automatically:
  - Locally, GITHUB_TOKEN is unset, so this uses a plain `git push`/
    `git pull`, relying on your already-cached local git credentials
    (set up once via a manual `git push` -- see the project README).
  - When GITHUB_TOKEN is set (only meaningful for a deployed instance,
    which has no cached local credentials at all), git operations go
    through an ephemeral token-authenticated remote URL instead, built
    fresh on every call rather than stored anywhere on disk.
"""

import logging
import subprocess
from typing import Optional

from config import BASE_DIR, EXCEL_FILE_PATH, GITHUB_REPO, GITHUB_TOKEN, TICKERS_FILE

logger = logging.getLogger(__name__)

# Every write action (add/remove company, edit Watchlist) touches at
# most these two files -- simplest to always stage both rather than
# track precisely which one a given caller changed.
TRACKED_PATHS = [EXCEL_FILE_PATH, TICKERS_FILE]

# A fresh cloud container has no git identity configured at all (unlike
# your Mac, which already has one globally) -- set it inline on every
# commit so this works regardless of environment.
_GIT_IDENTITY = ["-c", "user.name=Investment Dashboard Bot", "-c", "user.email=noreply@investment-dashboard.local"]


def _remote_url() -> str:
    if GITHUB_TOKEN:
        return f"https://x-access-token:{GITHUB_TOKEN}@github.com/{GITHUB_REPO}.git"
    return "origin"


def _rel_paths() -> list:
    paths = []
    for p in TRACKED_PATHS:
        try:
            paths.append(str(p.relative_to(BASE_DIR)))
        except ValueError:
            logger.warning("git_sync: %s is outside the repo -- skipping", p)
    return paths


def pull_latest() -> bool:
    """Fast-forward the local checkout to origin/main. Best-effort: a
    failed pull (offline, diverged history, not a git checkout) is
    logged and otherwise ignored -- the subsequent push will fail
    loudly on its own if we were genuinely out of sync, which is a
    clearer signal than trying to auto-resolve a conflict here.
    """
    try:
        subprocess.run(
            ["git", "pull", "--ff-only", _remote_url(), "main"],
            cwd=BASE_DIR, check=True, capture_output=True, text=True,
        )
        return True
    except subprocess.CalledProcessError as exc:
        logger.warning("git_sync: pull failed (%s) -- continuing with the current checkout", exc.stderr)
        return False
    except FileNotFoundError:
        logger.warning("git_sync: git executable not found -- skipping pull")
        return False


def push_state_if_changed(commit_message: str = "Data sync") -> bool:
    """Commit and push the workbook + tickers.json if either changed.

    Returns True if a push happened. Never raises -- a failed push (no
    network, no credentials, not a git checkout, rejected as
    non-fast-forward) is logged as a warning and the caller continues
    normally; the caller's own in-memory/on-disk state is unaffected
    either way, only the deployed copy stays stale until the next
    successful push.
    """
    rel_paths = _rel_paths()
    if not rel_paths:
        return False

    try:
        status = subprocess.run(
            ["git", "status", "--porcelain", *rel_paths],
            cwd=BASE_DIR, capture_output=True, text=True, check=True,
        )
        if not status.stdout.strip():
            logger.info("git_sync: nothing changed, nothing to push")
            return False

        subprocess.run(["git", "add", *rel_paths], cwd=BASE_DIR, check=True, capture_output=True)
        subprocess.run(
            ["git", *_GIT_IDENTITY, "commit", "-m", commit_message],
            cwd=BASE_DIR, check=True, capture_output=True,
        )
        subprocess.run(
            ["git", "push", _remote_url(), "HEAD:main"],
            cwd=BASE_DIR, check=True, capture_output=True, text=True,
        )
        logger.info("git_sync: pushed updated state to GitHub")
        return True
    except subprocess.CalledProcessError as exc:
        stderr = exc.stderr.strip() if isinstance(exc.stderr, str) else exc.stderr
        logger.warning(
            "git_sync: failed to push (%s) -- the deployed dashboard will show stale data "
            "until this is resolved",
            stderr or exc,
        )
        return False
    except FileNotFoundError:
        logger.warning("git_sync: git executable not found -- skipping push")
        return False


def sync_before_write() -> None:
    """Call before any write action that will be followed by
    push_state_if_changed(): pulls the latest state and drops the
    in-memory workbook cache (excel_workbook._workbook) so the write
    starts from what's actually on GitHub, not a possibly-stale copy
    this process loaded when it started.
    """
    pull_latest()
    import excel_workbook  # deferred to avoid a module-level import cycle

    excel_workbook.reload_workbook()
