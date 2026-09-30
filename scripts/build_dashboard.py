#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Rebuilds the Life Extension GMV dashboard HTML from Haley's OneDrive ops
fill-in workbook, and writes the finished page to OUT_PATH.

Usage:
    python3 build_dashboard.py <template.html> <out.html>

The template is the checked-in demo.html-derived page that contains the
placeholder line:
    var BUILD_META = {"generatedAt":"__PENDING__","sourceNote":"__PENDING__"};
    var REAL_RECORDS = [];
This script replaces exactly that line with the real, freshly-fetched data.

Only GMV is read and only GMV ever reaches the output. Net sales
(净销售) and target sheets in the workbook are never parsed or exposed.
"""
import os
import sys
import json
import subprocess
import datetime as dt

import openpyxl

# The OneDrive share link is read from an environment variable (ONEDRIVE_URL),
# never hardcoded here, since this file lives in a public repo. In GitHub
# Actions it's injected from a repository secret; the file itself can never
# leak the sheet's private net-sales/target data.
SHARE_URL = os.environ.get("ONEDRIVE_URL")
SHEET_NAME = "每日填报"
DATE_ROW = 4
GMV_FIRST_ROW = 5
GMV_LAST_ROW = 14
FIRST_DATE_COL = 2  # column B

# sheet channel label (col A) -> (subchannel key, group key) -- must match
# the SUBCHANNELS list in the dashboard HTML.
CHANNEL_MAP = {
    "天猫旗舰店": "tmall_flagship",
    "AliHealth": "tmall_alihealth",
    "TDI": "tmall_tdi",
    "京东自营": "jd_self",
    "京东POP": "jd_pop",
    "抖音": "douyin",
    "小红书": "red",
    "唯品会": "vip",
    "拼多多": "pdd",
    "分销": "b2b",
}

EXCEL_EPOCH = dt.datetime(1899, 12, 30)


def excel_serial_to_iso(serial):
    return (EXCEL_EPOCH + dt.timedelta(days=int(serial))).strftime("%Y-%m-%d")


def cell_to_iso_date(value):
    """Return an ISO date string if `value` is a date/datetime or a legacy
    Excel serial number, else None (e.g. a "9月汇总" summary-column label)."""
    if isinstance(value, (dt.datetime, dt.date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, (int, float)):
        return excel_serial_to_iso(value)
    return None


def download_workbook(dest_path):
    if not SHARE_URL:
        raise RuntimeError("ONEDRIVE_URL environment variable is not set")
    # NOTE: SharePoint's anonymous-link flow issues a "tenantanon" auth cookie
    # on the first redirect and requires it to be replayed on the next hop,
    # so the curl cookie engine must be explicitly enabled (-b/-c) even
    # though we don't care about persisting cookies across runs.
    url = SHARE_URL + "&download=1"
    cookie_jar = dest_path + ".cookies"
    subprocess.run(
        ["curl", "-sS", "-L", "-b", cookie_jar, "-c", cookie_jar, "-o", dest_path, url],
        check=True,
    )


def parse_records(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[SHEET_NAME]

    # channel label -> row index, from column A of the GMV block
    row_for_channel = {}
    for r in range(GMV_FIRST_ROW, GMV_LAST_ROW + 1):
        label = ws.cell(row=r, column=1).value
        if label in CHANNEL_MAP:
            row_for_channel[label] = r

    # Date columns live in row 4 from column B onward, but the sheet now
    # interleaves a "9月汇总" / "10月汇总" / ... text column after each
    # month's days -- those must be SKIPPED, not treated as the end of the
    # data (that was a real bug: it used to silently drop every month after
    # the first summary column). We scan the full row and stop only after a
    # run of genuinely empty cells (past the last real column, "12月汇总").
    date_cols = []  # list of (col_idx, iso_date)
    c = FIRST_DATE_COL
    empty_run = 0
    max_col = ws.max_column
    while c <= max_col and empty_run <= 10:
        v = ws.cell(row=DATE_ROW, column=c).value
        iso = cell_to_iso_date(v)
        if iso is not None:
            date_cols.append((c, iso))
            empty_run = 0
        elif v is None:
            empty_run += 1
        # else: a text summary column like "9月汇总" -- skip without counting
        # it as empty, since real date columns still follow it.
        c += 1

    records = []
    for col, iso_date in date_cols:
        rec = {"date": iso_date}
        any_value = False
        for label, key in CHANNEL_MAP.items():
            row = row_for_channel.get(label)
            val = ws.cell(row=row, column=col).value if row else None
            val = val or 0
            if val:
                any_value = True
            rec[key] = val
        if any_value:
            records.append(rec)

    records.sort(key=lambda r: r["date"])
    return records


def inject(template_path, out_path, records):
    with open(template_path, "r", encoding="utf-8") as f:
        html = f.read()

    marker = (
        'var BUILD_META = {"generatedAt":"__PENDING__","sourceNote":"__PENDING__"};\n'
        '  var REAL_RECORDS = [];'
    )
    if marker not in html:
        raise RuntimeError("template marker not found -- did the template change?")

    generated_at = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    meta = {
        "generatedAt": generated_at,
        "sourceNote": "运营渠道数据填报表 · 每日填报 · GMV",
    }
    replacement = (
        "var BUILD_META = " + json.dumps(meta, ensure_ascii=False) + ";\n"
        "  var REAL_RECORDS = " + json.dumps(records, ensure_ascii=False) + ";"
    )
    html = html.replace(marker, replacement)

    # The template is an Artifact-style fragment (no doctype/html/head/body).
    # GitHub Pages needs a real standalone document, so wrap it here.
    full_doc = (
        "<!doctype html>\n<html lang=\"zh-CN\"><head>\n"
        "<meta charset=\"utf-8\">\n"
        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n"
        "</head><body>\n" + html + "\n</body></html>\n"
    )

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(full_doc)


def main():
    if len(sys.argv) != 3:
        print("usage: build_dashboard.py <template.html> <out.html>", file=sys.stderr)
        sys.exit(1)
    template_path, out_path = sys.argv[1], sys.argv[2]

    xlsx_tmp = "/tmp/_le_ops_fetch.xlsx"
    download_workbook(xlsx_tmp)
    records = parse_records(xlsx_tmp)
    inject(template_path, out_path, records)
    print(f"wrote {out_path} with {len(records)} day-records")


if __name__ == "__main__":
    main()
