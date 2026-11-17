from __future__ import annotations

import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
import json
import logging
import os
from pathlib import Path
import re
import time
from urllib.parse import urlencode, urljoin, urlsplit
from zoneinfo import ZoneInfo

from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import requests
from tqdm import tqdm

from .rules import build_row, parse_date, rate
from .publishing import publish_report

CREDITORS = ("45244782", "60197609")
COLUMNS = ["SPIS_ZNACKA", "DATUM_UPADKU", "STAV", "PREDP_USP", "PREDP_USP_AKT", "NOTE"]
APPROVAL_EVENT_TYPES = {174, 175, 1065}


def write_json(path, value):
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(value, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    temp.replace(path)


class API:
    def __init__(self):
        self.root = os.getenv("REKENBER_API_ROOT", "https://api.rekenber.cz").rstrip("/")
        self.token = os.getenv("REKENBER_API_TOKEN")
        self.timeout = int(os.getenv("REKENBER_TIMEOUT_SECONDS", "60"))
        if not self.token:
            raise ValueError("Set REKENBER_API_TOKEN in .env or the environment")

    def get(self, url):
        url = urljoin(self.root + "/", url)
        if urlsplit(url).netloc != urlsplit(self.root).netloc or urlsplit(url).scheme != urlsplit(self.root).scheme:
            raise ValueError("Refusing API link outside configured API origin")
        for attempt in range(6):
            try:
                response = requests.get(url, headers={"token": self.token}, timeout=self.timeout, allow_redirects=False)
                if 300 <= response.status_code < 400:
                    raise ValueError("Unexpected API redirect")
                response.raise_for_status()
                return response.json()
            except requests.RequestException as exc:
                status = exc.response.status_code if exc.response is not None else None
                if attempt == 5 or (status is not None and status < 500 and status != 429):
                    raise
                delay = 0.7 * 2 ** attempt
                logging.warning(
                    "API request %s failed (%s: %s); retry %d/5 in %.1fs",
                    url,
                    type(exc).__name__,
                    exc,
                    attempt + 1,
                    delay,
                )
                time.sleep(delay)

    def pages(self, path, *, progress_label=None, progress_position=None, **params):
        url = path + "?" + urlencode(params, doseq=True)
        result, seen = [], set()
        with tqdm(desc=progress_label, unit="page", position=progress_position, dynamic_ncols=True, disable=None if progress_label else True) as progress:
            while url:
                if url in seen:
                    raise ValueError("API pagination cycle")
                seen.add(url)
                progress.set_postfix_str(f"requesting page {len(seen)}; {len(result):,} rows loaded")
                page = self.get(url)
                result.extend(page["results"])
                pagination = page.get("pagination") or {}
                if pagination.get("total_pages") is not None:
                    progress.total = int(pagination["total_pages"])
                progress.update(1)
                progress.set_postfix_str(f"{len(result):,} rows loaded")
                url = pagination.get("next")
        return result


def cached_fetch(workspace, name, keys, fetch, workers=8):
    """Checkpoint completed requests even if another request fails."""
    path = workspace / f"{name}_cache.json"
    cache = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    keys = sorted(set(keys))
    missing = [key for key in keys if key not in cache]
    logging.info("%s: %d cached, %d to fetch", name, len(keys) - len(missing), len(missing))
    with tqdm(total=len(keys), initial=len(keys) - len(missing), desc=name.replace("_", " "), unit="request", position=0, dynamic_ncols=True, disable=None) as progress, ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(fetch, key): key for key in missing}
        try:
            for count, future in enumerate(as_completed(futures), 1):
                cache[futures[future]] = future.result()
                progress.update(1)
                if count % 50 == 0:
                    write_json(path, cache)
                    logging.info("%s: fetched %d/%d", name, count, len(missing))
        finally:
            write_json(path, cache)
    return {key: cache[key] for key in keys}


def latest_event(events, event_type=None):
    candidates = [e for e in events if event_type is None or any(str(t["id"]) == str(event_type) for t in e.get("event_types", []))]
    return max(candidates, key=lambda e: (e["publication_date"], e.get("number_within_section", 0), e.get("subnumber") or 0), default=None)


def document_id(event):
    return str((event.get("document") or event.get("isir_document") or {})["id"]) if event and (event.get("document") or event.get("isir_document")) else None


def public_document(identifier):
    return f"https://isir.justice.cz:8443/isir_public_ws/doc/Document?idDokument={identifier}" if identifier else None


def review_documents(insolvency, events):
    zpo_event = latest_event(events, 1036) or latest_event(events, 1037)
    zpo_url = public_document(document_id(zpo_event))
    proposal = parse_date(insolvency.get("proposal_date"))
    if proposal is not None and proposal >= date(2025, 1, 1):
        approvals = [event for event in events if any(int(t["id"]) in APPROVAL_EVENT_TYPES for t in event.get("event_types", []))]
        return public_document(document_id(latest_event(approvals))), zpo_url
    return zpo_url, zpo_url


def fetch_extracts(connection_string, ids):
    import psycopg
    from psycopg.rows import dict_row

    result = {}
    with psycopg.connect(connection_string, row_factory=dict_row, connect_timeout=15) as connection:
        with connection.cursor() as cursor:
            cursor.execute("SET TRANSACTION READ ONLY")
            cursor.execute("SET LOCAL statement_timeout = '120s'")
            for start in tqdm(range(0, len(ids), 1000), desc="ZOPO database", unit="batch", dynamic_ncols=True, disable=None):
                cursor.execute("""
                    SELECT DISTINCT ON (ws_file_id) ws_file_id, extract
                    FROM annotations.extract
                    WHERE ws_file_id = ANY(%s) AND extract_type = %s
                    ORDER BY ws_file_id, created_at DESC, id DESC
                """, ([int(i) for i in ids[start:start + 1000]], "Zpráva o plnění oddlužení"))
                for row in cursor:
                    result[str(row["ws_file_id"])] = row["extract"]
    return {key: result.get(key) for key in ids}


def zopo_rates(workspace, events, skip=False):
    """Latest published ZOPO only; never substitute a stale older return rate."""
    latest = {key: document_id(latest_event(value, 39)) for key, value in events.items()}
    ids = sorted({i for i in latest.values() if i})
    path = workspace / "zopo_extracts_cache.json"
    cache = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    missing = [key for key in ids if key not in cache]
    warning = None
    if skip:
        warning = "ZOPO lookup explicitly skipped"
        return {key: None for key in latest}, warning
    if missing:
        connection_string = os.getenv("REKENBER_ANNOTATION_DB_URL")
        if not connection_string:
            raise ValueError("Set REKENBER_ANNOTATION_DB_URL or explicitly pass --skip-zopo")
        logging.info("ZOPO extracts: %d cached, %d documents to query", len(ids) - len(missing), len(missing))
        # Connection/query failures must be visible, not cached as empty matches.
        cache.update(fetch_extracts(connection_string, missing))
        write_json(path, cache)
    rates = {}
    for key, identifier in latest.items():
        extract = cache.get(identifier) or {}
        try:
            rates[key] = rate((extract.get("obligations_summary") or {}).get("current_return_rate"))
        except (ValueError, TypeError, AttributeError):
            logging.warning("Invalid current_return_rate in latest ZOPO for insolvency %s; leaving blank", key)
            rates[key] = None
    return rates, warning


def load_annotations(path):
    if not path.exists():
        return {}
    result = {}
    with path.open(encoding="utf-8") as handle:
        for line_number, line in enumerate(handle, 1):
            if not line.strip():
                continue
            try:
                item = json.loads(line)
            except json.JSONDecodeError as exc:
                raise ValueError(f"Invalid JSON in {path} at line {line_number}: {exc.msg}") from exc
            if not isinstance(item, dict):
                raise ValueError(f"Expected annotation object in {path} at line {line_number}")
            identifier = str(item.get("insolvency_id") or "").strip()
            if not identifier.isdigit() or identifier in result:
                raise ValueError(f"Invalid or duplicate annotation insolvency_id: {identifier}")
            rate(item.get("expected_return_rate"))
            result[identifier] = item
    return result


def append_annotation_templates(path, missing, existing):
    new = [item for item in missing if item["insolvency_id"] not in existing]
    if not new:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    needs_newline = False
    if path.exists() and path.stat().st_size:
        with path.open("rb") as handle:
            handle.seek(-1, 2)
            needs_newline = handle.read(1) != b"\n"
    with path.open("a", encoding="utf-8") as handle:
        if needs_newline:
            handle.write("\n")
        for item in new:
            handle.write(json.dumps({"expected_return_rate": None, "note": None, **item}, ensure_ascii=False) + "\n")


def write_workbook(path, columns, rows, title, widths):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = title
    sheet.append(columns)
    for row in rows:
        sheet.append([row.get(column) for column in columns])
        for cell in sheet[sheet.max_row]:
            # Notes/identifiers are literal strings, including user-entered leading '='.
            if isinstance(cell.value, str):
                cell.data_type = "s"
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if columns[cell.column - 1] == "DATUM_UPADKU":
                cell.number_format = "dd.mm.yyyy"
            elif columns[cell.column - 1] in {"PREDP_USP", "PREDP_USP_AKT", "expected_return_rate"}:
                cell.number_format = '0.00"%"'
            elif columns[cell.column - 1] == "document_url" and cell.value:
                cell.hyperlink = cell.value
                cell.style = "Hyperlink"
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="155C50")
    sheet.row_dimensions[1].height = 26
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column, width in zip(sheet.iter_cols(max_row=1), widths):
        sheet.column_dimensions[column[0].column_letter].width = width
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(".tmp.xlsx")
    workbook.save(temp)
    workbook.close()
    temp.replace(path)


def run(args):
    workspace = Path("workspace") / args.report_name
    workspace.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", handlers=[logging.StreamHandler(), logging.FileHandler(workspace / "report.log", encoding="utf-8")], force=True)
    api = API()
    manifest = {"api_root": api.root, "creditors": list(CREDITORS), "schema_version": 1}
    manifest_path = workspace / "manifest.json"
    if manifest_path.exists() and json.loads(manifest_path.read_text()) != manifest:
        raise ValueError("Workspace configuration differs; choose a new report name")
    write_json(manifest_path, manifest)
    if args.refresh:
        for path in workspace.glob("*_cache.json"):
            path.unlink()
    # A previous export must not masquerade as a successful rerun after a failure.
    output = workspace / "final_report.xlsx"
    if output.exists():
        output.replace(workspace / "previous_final_report.xlsx")
    write_json(workspace / "summary.json", {"status": "running"})
    fetch = lambda name, keys, fn: cached_fetch(workspace, name, keys, fn, args.max_workers)
    portfolios = fetch("creditor_insolvencies", CREDITORS, lambda key: api.pages("/v1/isir/insolvencies", progress_label=f"Creditor {key}", progress_position=CREDITORS.index(key) + 1, creditor_subject_id=key, page_size=100))
    candidates = {str(item["self"]["id"]) for items in portfolios.values() for item in items if item["state"] == "ODDLUŽENÍ"}
    details = fetch("insolvencies", candidates, lambda key: api.get(f"/v1/isir/insolvencies/{key}"))
    details = {key: item for key, item in details.items() if item["state"] == "ODDLUŽENÍ"}
    discharges = fetch("discharges", details, lambda key: api.get(details[key]["discharge"]["href"]) if details[key].get("discharge") else None)
    debtor_links = {d["href"] for ins in details.values() for d in ins.get("debtors", []) if d.get("href")}
    debtors = fetch("debtors", debtor_links, api.get)
    events = fetch("events_v4", details, lambda key: api.pages("/v1/isir/events", insolvency_id=key, event_type_id=[39, 1036, 1037, 174, 175, 1065], page_size=100))
    current, warning = zopo_rates(workspace, events, args.skip_zopo)
    annotations_path = Path(os.getenv("REKENBER_ANNOTATIONS_PATH", "data/discharge_annotations.jsonl"))
    annotations = load_annotations(annotations_path)
    rows, missing, audit = [], [], []
    for key, ins in sorted(details.items(), key=lambda pair: (pair[1]["reference_number"]["year"], pair[1]["reference_number"]["case_number"])):
        correction = annotations.get(key, {})
        if str(correction.get("note") or "").strip().casefold() == "exclude":
            continue
        people = [debtors.get(d.get("href"), d) for d in ins.get("debtors", [])]
        row, reasons = build_row(ins, discharges[key] or {}, people, current[key], correction)
        rows.append(row)
        audit.append({"insolvency_id": key, **row, "latest_zopo_document_id": document_id(latest_event(events[key], 39))})
        if reasons:
            item = dict(annotations.get(key, {"insolvency_id": key}))
            item.update(reference_number=row["SPIS_ZNACKA"], reason="; ".join(reasons))
            item["document_url"], item["zpo_document_url"] = review_documents(ins, events[key])
            missing.append(dict(item))
            logging.warning("Missing data: %s (insolvency_id=%s) | %s | source: %s | ZPO (Zpráva pro oddlužení): %s", row["SPIS_ZNACKA"], key, item["reason"], item["document_url"] or "no source document available", item["zpo_document_url"] or "no ZPO document available")
    write_json(workspace / "rows.json", audit)
    write_json(workspace / "missing_data.json", missing)
    summary = {"status": "needs_annotations" if missing else "exporting", "created_at": datetime.now(ZoneInfo("Europe/Prague")).isoformat(), "rows": len(rows), "missing_data": len(missing), "updated_rates_present": sum(r["PREDP_USP_AKT"] is not None for r in rows), "warning": warning}
    write_json(workspace / "summary.json", summary)
    if warning:
        logging.warning(warning)
    if missing:
        append_annotation_templates(annotations_path, missing, annotations)
        raise SystemExit(f"Export stopped: {len(missing)} insolvencies need a rate or note in {annotations_path}. Edit the JSONL rows, then rerun the same command.")
    write_workbook(output, COLUMNS, rows, "Report", [25, 22, 20, 20, 22, 90])
    summary["status"] = "complete"
    write_json(workspace / "summary.json", summary)
    logging.info("Exported %d insolvencies to %s", len(rows), output)
    if os.getenv("REKENBER_PUBLISH_REPORT", "false").lower() == "true":
        publish_saved_report(args.report_name)


def publish_saved_report(report_name):
    workspace = Path("workspace") / report_name
    summary = json.loads((workspace / "summary.json").read_text(encoding="utf-8"))
    if summary.get("status") != "complete":
        raise ValueError("Only a successfully completed report can be published")
    try:
        summary["publication"] = publish_report(workspace / "final_report.xlsx", report_name, summary["rows"])
        summary["publication_status"] = "published"
    except Exception:
        summary["publication_status"] = "failed"
        write_json(workspace / "summary.json", summary)
        raise
    write_json(workspace / "summary.json", summary)


def main():
    load_dotenv()
    parser = argparse.ArgumentParser(description="Monthly ČS / SSČS insolvency return report")
    subparsers = parser.add_subparsers(dest="command", required=True)
    command = subparsers.add_parser("run")
    command.add_argument("--report-name", required=True, help="New name each month; rerun same name to reuse snapshot caches")
    command.add_argument("--max-workers", type=int, default=int(os.getenv("REKENBER_MAX_WORKERS", "8")))
    command.add_argument("--refresh", action="store_true", help="Invalidate all snapshot caches before fetching")
    command.add_argument("--skip-zopo", action="store_true", help="Explicitly leave PREDP_USP_AKT blank, without database access")
    publish = subparsers.add_parser("publish", help="Publish an already completed local report without fetching API data")
    publish.add_argument("--report-name", required=True)
    args = parser.parse_args()
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_.-]*", args.report_name):
        parser.error("report-name must be a simple name without path separators")
    if args.command == "publish":
        logging.basicConfig(level=logging.INFO)
        publish_saved_report(args.report_name)
        return
    if args.max_workers < 1:
        parser.error("max-workers must be positive")
    run(args)


if __name__ == "__main__":
    main()
