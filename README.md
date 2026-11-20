# ČS retail portfolio returns report

Monthly Excel report for Česká spořitelna (`45244782`) and Stavební spořitelna
České spořitelny (`60197609`). Uses the same ISIR HTTP API and snapshot/annotation
workflow as the adjacent Zaplo and Provident projects. No input portfolio is needed.

```sh
uv sync
cp .env.example .env
# Fill REKENBER_API_TOKEN and REKENBER_ANNOTATION_DB_URL in .env
uv run cs-returns-report run --report-name 2026-11
```

Output: `workspace/2026-11/final_report.xlsx`. Use a **new report name each month**.
The same name resumes a snapshot, including cached empty responses and missing
ZOPO extracts. `--refresh` clears that snapshot's caches, including DB extracts;
manual corrections are retained. This is a current-state snapshot collected when
run, not a reconstruction of historical state at a month-end date.
`--max-workers 8` controls API parallelism. `.env` is loaded automatically.
Terminal progress bars show completed/cached requests for each stage, separate
page counts and loaded-row counts for both creditor lists, and ZOPO database
batches. While waiting for a page, the creditor bar shows which page is being
requested. API retries log the error type and backoff delay.

All list endpoints are paginated. The pipeline unions both creditors' insolvency
lists, filters exact current state `ODDLUŽENÍ`, fetches details and checks state
again, and deduplicates by insolvency ID. The API creditor filter identifies
insolvencies with creditor receivables; neither creditor must be present twice.
Caches are atomic JSON checkpoints scoped to each report. A failed request stops
the run and preserves successful requests. Treat the workspace as sensitive data.

The report contains exactly these columns:

| Column | Meaning |
| --- | --- |
| SPIS_ZNACKA | `INS {case_number}/{year}` from insolvency detail |
| DATUM_UPADKU | **Discharge allowed date**, as requested (not a separate insolvency-decision date) |
| STAV | Current API state |
| PREDP_USP | Expected return selected using the rules below |
| PREDP_USP_AKT | Latest ZOPO's `obligations_summary.current_return_rate`; nullable |
| NOTE | Manual edge-case explanation |
| DATUM_EXPORTU | Export date in Europe/Prague, identical for every row |

Rates use **percentage points** in both API/annotation values and Excel: enter
`60` for 60%, `0.5` for 0.5%. Excel displays a literal percent sign without
rescaling the number. Zero is valid. Dates are native Excel dates.

## Expected return rules

- Proposal **on or after 2025-01-01**: `approval.approved_expected_return_rate` on discharge.
  The cutoff date is included to avoid a gap with the preceding interval.
- Proposal in **[2021-06-01, 2025-01-01)**: use the satisfaction estimate with scope
  `splátkový kalendář se zpeněžením majetkové podstaty`. Prefer scope
  `splátkový kalendář se zpeněžením majetkové podstaty za 3 roky` if its rate is
  at least 60, or if a debtor reached ordinary retirement age at proposal.
  A missing three-year estimate falls back to the standard estimate.
- Earlier proposal: use only the standard scope above.

Retirement is deliberately best effort: use the **higher ordinary Czech age
limit for the birth year**, irrespective of gender/children raised. This is 60
years before birth year 1936, then 60 years 2 months plus 2 months per birth
year, capped at 65. This pre-2025 schedule applies to the proposal interval
where retirement affects selection. Source: [ČSSZ retirement ages](https://www.cssz.gov.cz/starobni-duchod-podrobne).
For joint debtors, one qualifying debtor suffices. Missing birth date defaults
to not retired. These cases do **not** create retirement review tasks. Early
retirement, actual pension receipt and special occupational limits are not modeled.

After approval (`approval.discharge_approved_date`), a missing selected expected
rate blocks export until a rate **or an explanatory note** is provided. Before
approval, an absent rate is expected and stays blank. Missing proposal dates,
conflicting estimates, missing discharge resources or allowance dates also need
a correction/note so missing data is not silently treated as unapproved.

The API must expose `approval.approved_expected_return_rate`. After deploying
this field, rerun with `--refresh` to replace cached responses that lack it;
otherwise approved post-cutoff cases enter the annotation loop.

## Manual annotation loop

The first run with missing required data stops export, writes
`workspace/<name>/missing_data.json`, and creates/appends
`data/discharge_annotations.jsonl` (override with `REKENBER_ANNOTATIONS_PATH`).
Only new missing insolvency IDs are appended; existing rows are never rewritten
or appended again on reruns. Edit `expected_return_rate` or `note` in each row,
then rerun the **same command**. For example:

```json
{"insolvency_id": "12023", "expected_return_rate": 42, "note": null}
```

Generated templates also include the reference, reason, `document_url`, and
`zpo_document_url`. Missing-data diagnostics print these links to the console
and report log. For proposals on or after 2025-01-01, the source is the latest
approval event among types **174, 175, 1065**. For earlier proposals, it is the
latest **ZPO / Zpráva pro oddlužení (type 1036)**, falling back to the latest
type **1037** only when there is no 1036 event. The ZPO link is also printed
separately for newer cases. Missing source documents are reported explicitly.
The event cache uses a new name so existing snapshots fetch these additional types.

Numeric corrections override the selected API rate; notes appear in NOTE and
allow edge-case blanks. Corrections persist across months, are keyed by stable
insolvency ID, and never modify raw caches. Duplicate IDs, malformed JSON and
invalid rates stop the run. An existing numeric correction remains authoritative
on future snapshots until removed.

Export happens only after all required cases pass. A previous final workbook is
renamed `previous_final_report.xlsx` when rerunning, so an unsuccessful rerun has
no misleading current final output. `summary.json` reports completion/review
status, row count and updated-rate coverage; `rows.json` preserves per-row IDs
and latest ZOPO document provenance; `report.log` records progress.

## Latest ZOPO lookup

Uses paginated ISIR events of type **39**, chooses the latest published event
(publication date, then section number/subnumber), and queries
`annotations.extract` by its `document.id` = `ws_file_id` and extract type
`Zpráva o plnění oddlužení`. Within a document, the newest extract wins by
`created_at`, then `id`. Only `obligations_summary.current_return_rate` is used.
No older ZOPO is substituted if the latest document lacks an extract or rate.
This applies to all proposal dates and has no inherited 2025 event-date filter.

`REKENBER_ANNOTATION_DB_URL` is a PostgreSQL connection URI, with the credentials/SSH
tunnel you use locally. Only read queries are issued, in batches with a timeout.
Missing extracts/rates remain blank and never create review tasks. Database
connection/query failures stop visibly and are not cached as empty data; if DB
access is unavailable, `--skip-zopo` explicitly produces blank PREDP_USP_AKT and
records that choice in the summary. An app database connection is unnecessary.

## Checks

```sh
uv run python -m unittest discover -s tests -v
```

Tests cover date/60% boundaries, conservative retirement and joint debtors,
zero versus missing, annotation gates and overrides, pagination, cache reuse,
latest-document selection, Excel round-trips, and a complete stopped/reviewed/
resumed pipeline with creditor deduplication. No production API or DB is needed.

## Publish to Reporty in the app

The publisher connects **directly** to the app PostgreSQL database and the app's
S3-compatible storage over your VPN. There is no report upload/registration API.
Local Excel output remains available even if publication fails.

Before the first publication:

1. Apply `storage/rekenber_app_db/migration/V423__generated_reports.sql` in the
   `rekenber` repository, then deploy its GraphQL backend and frontend changes.
   This creates `generated_report` and the `organization_settings.reports_enabled`
   flag (disabled by default). Enable it for umbrella organization **2** in the
   target instance's organization settings.
2. Create the **`rekenber-generated-reports`** bucket in the app's S3 storage.
   The existing app file-storage path uses AWS S3 in the checked-in production
   configuration. A Hetzner or Seaweed **S3 endpoint** can be configured too.
3. Configure the backend's `AWS_REPORTS_S3_BUCKET_NAME=rekenber-generated-reports`
   (also the default). Its existing AWS credentials need `GetObject` on this bucket.
   `AWS_S3_ENDPOINT` is optional, and `AWS_S3_REGION` defaults to `us-east-1`.
4. Configure this project's `.env`: `REKENBER_APP_DB_URL`,
   `REKENBER_REPORT_S3_BUCKET`, `REKENBER_REPORT_S3_ROOT`, region/optional endpoint,
   and report S3 access/secret keys. The publisher needs `PutObject` on the bucket
   and SELECT on organization/settings plus INSERT/UPDATE on `generated_report`
   (and sequence access). The S3 endpoint/region must match the app backend.

Set `REKENBER_PUBLISH_REPORT=true` to publish automatically after each successful
`run`. The default `false` preserves local-only runs until connections are set up.
To publish an already completed report without refetching or regenerating:

```sh
uv run cs-returns-report publish --report-name 2026-11
```

The target umbrella ID defaults to **2** (`REKENBER_REPORT_ORGANIZATION_ID`).
The publisher refuses disabled organizations. Reports are ordinary **unencrypted
XLSX files**, independent of the app's Vault setting.

Object layout:

```text
<root>/reports/organizations/2/cs-retail-portfolio-returns/<upload-date>/<revision>/cs-retail-portfolio-returns-<upload-date>.xlsx
```

Each upload uses a new immutable revision. After upload succeeds, a transaction
upserts one registry row per organization/report type/report name. The title, period and filename use the upload date in Europe/Prague (YYYY-MM-DD).
Publishing again on the same date updates that entry. A different upload date creates a new row.
Failure leaves the previous registered revision intact. Older or unregistered
objects are retained; any future cleanup must preserve currently referenced keys.
`summary.json` records `publication_status`, report ID and checksum.

In the app, **Reporty** is visible only in enabled organization workspaces. It
shows a paginated table with name, period, generation time, row count and size,
and a **Stáhnout XLSX** button. List and download requests enforce umbrella
organization membership on the server; admins/analysts may use the selected
workspace. Disabling the settings flag blocks downloads as well as hiding the
menu. The app verifies downloaded bytes against the stored size/SHA-256.

Availability, listing and downloads use the app's configured **instance GraphQL
connection** (`VITE_GRAPHQL_URL`): `reportAccess`, `generatedReports`, and
`generatedReportContent`. No `/file/reports` REST routes or generic API base URL
are used. GraphQL resolves the current user through the existing `@CurrentUser`
decorator; ordinary users send no organization selector, while admins/analysts
may pass the selected workspace. Downloads request uncached base64 content and
save it as a plain XLSX file. The report registry and storage configuration
belong to the instance serving those queries.

To change visibility later, update the existing organization settings:

```sql
UPDATE public.organization_settings
SET reports_enabled = false
WHERE organization_id = 2;
```
