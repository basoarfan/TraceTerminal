"""JSON logging for shareholder analysis results."""

import json
from datetime import datetime
from pathlib import Path
from typing import Any


class DataLogger:
    """Save individual analyses and complete sessions as JSON."""

    def __init__(self, logs_dir: Path) -> None:
        self.logs_dir = logs_dir
        self.logs_dir.mkdir(parents=True, exist_ok=True)

    def log_analysis(
        self,
        analysis: dict[str, Any],
        query: dict[str, Any],
    ) -> Path:
        log_file = self.logs_dir / f"analysis_{self._timestamp()}.json"
        self._write(
            log_file,
            {
                "timestamp": datetime.now().isoformat(),
                "analysis": analysis,
                "query": query,
            },
        )
        return log_file

    def log_session(self, entries: list[dict[str, Any]]) -> Path:
        log_file = self.logs_dir / f"session_{self._timestamp()}.json"
        self._write(
            log_file,
            {
                "timestamp": datetime.now().isoformat(),
                "total_analyses": len(entries),
                "entries": entries,
            },
        )
        return log_file

    def _write(self, path: Path, data: dict[str, Any]) -> None:
        with path.open("w", encoding="utf-8") as file:
            json.dump(data, file, indent=2, ensure_ascii=False, default=str)

    def _timestamp(self) -> str:
        return datetime.now().strftime("%Y%m%d_%H%M%S_%f")
