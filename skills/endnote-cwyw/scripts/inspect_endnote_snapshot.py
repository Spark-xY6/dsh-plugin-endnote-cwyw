# /// script
# requires-python = ">=3.10"
# dependencies = []
# ///

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sqlite3
import struct
import sys
from pathlib import Path


FIELDS = [
    "id",
    "author",
    "year",
    "title",
    "secondary_title",
    "volume",
    "pages",
    "number",
    "accession_number",
    "reference_type",
]

OPTIONAL_FIELDS = [
    "electronic_resource_number",
    "abstract",
    "url",
]


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest().upper()


def text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    return str(value)


def normalize_doi(value: object) -> str:
    raw = text(value).strip()
    if not raw:
        return ""
    cleaned = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", raw, flags=re.IGNORECASE)
    cleaned = re.sub(r"^doi:\s*", "", cleaned, flags=re.IGNORECASE).strip()
    match = re.search(r"\b(10\.\d{4,9}/[^\s<>\"'?&#]+)", cleaned)
    if match:
        return match.group(1).rstrip(".,;)]}").casefold()
    return ""


def normalize_group(value: str) -> str:
    return " ".join(value.replace("_", " ").casefold().split())


def group_name(spec: object) -> str:
    match = re.search(r"<name>(.*?)</name>", text(spec), flags=re.DOTALL)
    return match.group(1).strip() if match else ""


def parse_members(blob: bytes | None) -> list[int]:
    if blob is None or len(blob) < 8:
        raise ValueError("group members BLOB is missing or shorter than eight bytes")
    payload = len(blob) - 8
    if payload % 4:
        raise ValueError("group members BLOB length is not 8 + 4N")
    count = struct.unpack_from("<I", blob, 4)[0]
    if count != payload // 4:
        raise ValueError(
            f"group member count mismatch: header={count}, payload={payload // 4}"
        )
    ids = list(struct.unpack_from(f"<{count}I", blob, 8)) if count else []
    if len(ids) != len(set(ids)):
        raise ValueError("group members BLOB contains duplicate record IDs")
    return ids


def parse_ids(value: str | None) -> list[int]:
    if not value:
        return []
    return sorted({int(item.strip()) for item in value.split(",") if item.strip()})


def record_dict(row: tuple[object, ...], selected_fields: list[str]) -> dict[str, object]:
    result: dict[str, object] = {
        name: text(value) for name, value in zip(selected_fields, row)
    }
    result["id"] = int(row[0])
    result["doi"] = normalize_doi(result.get("electronic_resource_number", ""))
    return result


def inspect(
    database: Path,
    group: str | None,
    requested_ids: list[int],
    *,
    include_trashed: bool = False,
    all_records: bool = False,
) -> dict[str, object]:
    uri = f"file:{database.as_posix()}?mode=ro&immutable=1"
    connection = sqlite3.connect(uri, uri=True)
    try:
        cursor = connection.cursor()
        tables = {
            row[0]
            for row in cursor.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()
        }
        missing_tables = sorted({"refs", "groups"} - tables)
        if missing_tables:
            raise RuntimeError(f"Required EndNote tables missing: {missing_tables}")
        ref_columns = {
            row[1] for row in cursor.execute('PRAGMA table_info("refs")').fetchall()
        }
        missing_columns = sorted(set(FIELDS) - ref_columns)
        if missing_columns:
            raise RuntimeError(f"Required refs columns missing: {missing_columns}")

        selected_fields = list(FIELDS) + [
            col for col in OPTIONAL_FIELDS if col in ref_columns
        ]
        has_trash_state = "trash_state" in ref_columns
        trashed_count = (
            int(
                cursor.execute(
                    "SELECT COUNT(*) FROM refs WHERE COALESCE(trash_state, 0) != 0"
                ).fetchone()[0]
            )
            if has_trash_state
            else 0
        )
        where_clause = (
            " WHERE COALESCE(trash_state, 0) = 0"
            if has_trash_state and not include_trashed
            else ""
        )
        column_sql = ", ".join(f'"{name}"' for name in selected_fields)
        rows = cursor.execute(f"SELECT {column_sql} FROM refs{where_clause}").fetchall()
        records = {
            int(row[0]): record_dict(row, selected_fields) for row in rows
        }
        if len(records) != len(rows):
            raise RuntimeError("refs contains duplicate record IDs")

        all_ref_ids = (
            {int(r[0]) for r in cursor.execute("SELECT id FROM refs").fetchall()}
            if has_trash_state and not include_trashed
            else set(records)
        )

        group_result: dict[str, object] | None = None
        group_ids: set[int] | None = None
        if group:
            target = normalize_group(group)
            matches = []
            for group_id, spec, members in cursor.execute(
                "SELECT group_id, spec, members FROM groups"
            ).fetchall():
                name = group_name(spec)
                if normalize_group(name) == target:
                    matches.append((int(group_id), name, parse_members(members)))
            if len(matches) != 1:
                raise RuntimeError(
                    f"Expected exactly one group matching {group!r}; found {len(matches)}"
                )
            group_id, name, members = matches[0]
            missing_group_records = sorted(set(members) - all_ref_ids)
            if missing_group_records:
                raise RuntimeError(
                    f"Group contains IDs absent from refs: {missing_group_records}"
                )
            active_members = [m for m in members if m in records]
            group_ids = set(active_members)
            group_result = {
                "group_id": group_id,
                "group_name": name,
                "group_count": len(active_members),
                "unique_member_ids": len(group_ids),
            }

        if all_records:
            selected = [records[record_id] for record_id in sorted(records)]
        else:
            selected = [
                records[record_id]
                for record_id in requested_ids
                if record_id in records
            ]
        missing_requested = sorted(set(requested_ids) - set(records))
        outside_group = (
            sorted(set(requested_ids) - group_ids)
            if requested_ids and group_ids is not None
            else []
        )
        checks = {
            "schema_readable": True,
            "library_ids_unique": len(records) == len(rows),
            "requested_records_exist": not missing_requested,
        }
        if group is not None:
            checks["requested_records_in_group"] = not outside_group
        return {
            "status": "PASS" if all(checks.values()) else "FAIL",
            "database": str(database),
            "database_sha256": sha256_file(database),
            "database_size": database.stat().st_size,
            "library_count": len(records),
            "trashed_count": trashed_count,
            "include_trashed": include_trashed,
            "full_integrity_check": "NOT_VERIFIED_PROPRIETARY_COLLATION_MAY_BE_UNAVAILABLE",
            "group": group_result,
            "requested_record_ids": requested_ids,
            "selected_records": selected,
            "missing_requested_record_ids": missing_requested,
            "requested_record_ids_outside_group": outside_group,
            "checks": checks,
        }
    finally:
        connection.close()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Inspect a frozen EndNote sdb.eni snapshot without modifying it."
    )
    parser.add_argument("--database", required=True, type=Path)
    parser.add_argument("--group")
    parser.add_argument("--record-ids")
    parser.add_argument("--include-trashed", action="store_true")
    parser.add_argument("--all-records", action="store_true")
    parser.add_argument("--report", type=Path)
    return parser


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    args = build_parser().parse_args()
    database = args.database.resolve()
    if not database.is_file():
        raise FileNotFoundError(database)
    if database.parent.name.casefold() == "sdb" and database.parent.parent.name.casefold().endswith(
        ".data"
    ):
        raise RuntimeError(
            "Refusing the live EndNote database path. Close EndNote and copy sdb.eni into the work area first."
        )
    result = inspect(
        database,
        args.group,
        parse_ids(args.record_ids),
        include_trashed=args.include_trashed,
        all_records=args.all_records,
    )
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if args.report:
        report = args.report.resolve()
        if report.exists():
            raise FileExistsError(f"Refusing to overwrite report: {report}")
        report.parent.mkdir(parents=True, exist_ok=True)
        report.write_text(rendered + "\n", encoding="utf-8")
    print(rendered)
    return 0 if result["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())

