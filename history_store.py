"""Local per-document history for ActGeneratorPC.

Every generated act gets its own JSON metadata file.  There is deliberately no
shared mutable index: one interrupted write cannot damage the remaining history.
"""

from __future__ import annotations

import datetime as dt
import json
import os
import time
import uuid
from pathlib import Path
from typing import Mapping


def _stored_path(path: str | Path, app_dir: str | Path) -> str:
    """Handle stored path."""
    absolute = Path(path).resolve()
    root = Path(app_dir).resolve()
    try:
        return absolute.relative_to(root).as_posix()
    except ValueError:
        return str(absolute)


def _resolved_path(value: str, app_dir: str | Path) -> str:
    """Handle resolved path."""
    candidate = Path(value)
    if not candidate.is_absolute():
        candidate = Path(app_dir) / candidate
    return str(candidate.resolve())


def save_history_record(
    # COLLEAGUE EDIT POINT: history schema is assembled in this function.
    history_dir: str | Path,
    app_dir: str | Path,
    document_path: str | Path,
    act_data: Mapping[str, object],
) -> str:
    """Save history record for this workflow."""
    directory = Path(history_dir)
    directory.mkdir(parents=True, exist_ok=True)
    record_id = uuid.uuid4().hex
    document = Path(document_path).resolve()
    data = dict(act_data)
    if data.get("template"):
        data["template"] = _stored_path(str(data["template"]), app_dir)
    now = dt.datetime.now(dt.timezone.utc).isoformat()
    record = {
        "id": record_id,
        "created_at": now,
        "modified_at": now,
        "document": _stored_path(document, app_dir),
        "document_name": document.name,
        "data": data,
    }
    target = directory / f"{document.stem}-{record_id}.json"
    temporary = target.with_suffix(".json.tmp")
    temporary.write_text(
        json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    os.replace(temporary, target)
    return str(target)


def load_history_records(
    history_dir: str | Path,
    app_dir: str | Path,
    output_dir: str | Path,
) -> list[dict[str, object]]:
    """Load history records for this workflow."""
    directory = Path(history_dir)
    directory.mkdir(parents=True, exist_ok=True)
    records: list[dict[str, object]] = []
    indexed_documents: set[str] = set()
    clear_marker = directory / ".cleared-before"
    try:
        cleared_before = float(clear_marker.read_text(encoding="ascii").strip())
    except (OSError, ValueError):
        cleared_before = 0.0

    for path in directory.glob("*.json"):
        try:
            record = json.loads(path.read_text(encoding="utf-8-sig"))
            data = record.get("data")
            if not isinstance(record, dict) or not isinstance(data, dict):
                continue
            document = _resolved_path(str(record.get("document", "")), app_dir)
            if not document:
                continue
            # The DOCX file is the source of truth for whether an act still
            # exists.  Keep the sidecar metadata on disk so a restored file can
            # reappear, but do not show stale rows in the history table.
            if not Path(document).is_file():
                continue
            if data.get("template"):
                data["template"] = _resolved_path(str(data["template"]), app_dir)
            record["document"] = document
            record["metadata_path"] = str(path.resolve())
            record["data"] = data
            created_at = str(record.get("created_at", ""))
            try:
                created_timestamp = dt.datetime.fromisoformat(created_at).timestamp()
            except ValueError:
                created_timestamp = path.stat().st_mtime
            if created_timestamp <= cleared_before:
                indexed_documents.add(os.path.normcase(os.path.abspath(document)))
                continue
            records.append(record)
            indexed_documents.add(os.path.normcase(os.path.abspath(document)))
        except (OSError, ValueError, TypeError):
            continue

    output = Path(output_dir)
    if output.exists():
        for document in output.glob("*.docx"):
            absolute = str(document.resolve())
            if os.path.normcase(absolute) in indexed_documents:
                continue
            if document.stat().st_mtime <= cleared_before:
                continue
            modified = dt.datetime.fromtimestamp(
                document.stat().st_mtime, tz=dt.timezone.utc
            ).isoformat()
            records.append(
                {
                    "id": f"legacy-{document.stem}",
                    "created_at": modified,
                    "modified_at": modified,
                    "document": absolute,
                    "document_name": document.name,
                    "metadata_path": "",
                    "data": {
                        "date": dt.datetime.fromtimestamp(
                            document.stat().st_mtime
                        ).strftime("%d.%m.%Y"),
                        "executor": "",
                        "location": "",
                        "model": "",
                        "serial": "",
                        "work": "",
                        "issues": "",
                        "done_work": "",
                        "materials": "",
                        "qty": 0,
                        "template": "",
                    },
                }
            )

    records.sort(key=lambda item: str(item.get("created_at", "")), reverse=True)
    return records


def clear_history_records(history_dir: str | Path) -> int:
    """Clear history metadata without deleting any generated DOCX files."""
    directory = Path(history_dir)
    directory.mkdir(parents=True, exist_ok=True)
    marker = directory / ".cleared-before"
    temporary = marker.with_suffix(".tmp")
    temporary.write_text(str(time.time()), encoding="ascii")
    os.replace(temporary, marker)
    removed = 0
    for path in directory.glob("*.json"):
        path.unlink()
        removed += 1
    return removed
