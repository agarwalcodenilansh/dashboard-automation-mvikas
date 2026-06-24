from __future__ import annotations

import subprocess
from datetime import datetime
from pathlib import Path
from typing import Any

from .config import Settings
from .processor import parse_excel, write_json
from .renderer import write_script


def update_dashboard(source: str | Path, static_dir: str | Path, report_date: str | None = None) -> dict[str, Any]:
    """Parse source Excel/Google Sheet and write script.js + latest_data.json."""
    static = Path(static_dir).resolve()
    static.mkdir(parents=True, exist_ok=True)
    data = parse_excel(source, report_date=report_date)
    write_script(data, static / "script.js")
    write_json(data, static / "latest_data.json")
    return {
        "ok": True,
        "reportDate": data.get("reportDate"),
        "generatedAt": data.get("generatedAt"),
        "scriptPath": str(static / "script.js"),
        "dataPath": str(static / "latest_data.json"),
        "totals": {
            "clients": len(data.get("clients", [])),
            "open": sum(int(x.get("count", 0)) for x in data.get("openData", [])),
            "edd": sum(int(x.get("count", 0)) for x in data.get("eddData", [])),
            "due": sum(int(x.get("count", 0)) for x in data.get("dueData", [])),
            "booked": sum(int(x.get("count", 0)) for x in data.get("bookedData", [])),
            "dailyKg": round(sum(float(x.get("kg", 0)) for x in data.get("dailyTonnageData", [])), 2),
            "monthKg": round(sum(float(x.get("achieved", 0)) for x in data.get("clients", [])), 2),
        },
    }


def commit_outputs(settings: Settings, message: str | None = None) -> None:
    """Optionally commit generated assets when the backend runs inside a git clone."""
    if not settings.commit_to_git:
        return
    repo = Path(settings.git_repo_dir or settings.static_dir).resolve()
    if not (repo / ".git").exists() and not (repo.parent / ".git").exists():
        raise RuntimeError(f"Git repository not found for commit: {repo}")
    msg = message or f"chore: update MVIKAS dashboard data {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    commands = [
        ["git", "config", "user.name", settings.git_author_name],
        ["git", "config", "user.email", settings.git_author_email],
        ["git", "add", "script.js", "latest_data.json"],
        ["git", "commit", "-m", msg],
        ["git", "push", "origin", settings.git_branch],
    ]
    for cmd in commands:
        result = subprocess.run(cmd, cwd=repo, text=True, capture_output=True)
        # No changes to commit is not a failure for scheduled jobs.
        if result.returncode != 0:
            combined = (result.stdout + "\n" + result.stderr).lower()
            if "nothing to commit" in combined or "no changes added" in combined:
                return
            raise RuntimeError(f"Git command failed: {' '.join(cmd)}\n{result.stdout}\n{result.stderr}")
