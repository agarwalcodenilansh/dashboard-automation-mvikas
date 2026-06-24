from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    # python-dotenv is listed in requirements.txt, but keeping this optional makes
    # the CLI usable in lightweight environments that only pass real env vars.
    pass


@dataclass(frozen=True)
class Settings:
    source_url: str = os.getenv("MVIKAS_SOURCE_URL", "").strip()
    static_dir: Path = Path(os.getenv("MVIKAS_STATIC_DIR", "static")).resolve()
    update_time: str = os.getenv("MVIKAS_UPDATE_TIME", "09:30").strip()
    timezone: str = os.getenv("MVIKAS_TIMEZONE", "Asia/Kolkata").strip()
    auto_start: bool = os.getenv("MVIKAS_AUTO_START", "true").lower() in {"1", "true", "yes", "on"}
    commit_to_git: bool = os.getenv("MVIKAS_COMMIT_TO_GIT", "false").lower() in {"1", "true", "yes", "on"}
    git_repo_dir: str = os.getenv("MVIKAS_GIT_REPO_DIR", "").strip()
    git_branch: str = os.getenv("MVIKAS_GIT_BRANCH", "main").strip()
    git_author_name: str = os.getenv("MVIKAS_GIT_AUTHOR_NAME", "MVIKAS Bot").strip()
    git_author_email: str = os.getenv("MVIKAS_GIT_AUTHOR_EMAIL", "mvikas-bot@example.com").strip()

    @property
    def script_path(self) -> Path:
        return self.static_dir / "script.js"

    @property
    def data_path(self) -> Path:
        return self.static_dir / "latest_data.json"


settings = Settings()
