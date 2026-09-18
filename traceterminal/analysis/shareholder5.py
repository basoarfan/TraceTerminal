"""Shareholder ownership >=5% Excel parser and analyzer."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from functools import lru_cache
from pathlib import Path
from typing import Any


UNKNOWN = "Data tidak tersedia"
NOT_IN_DOCUMENT = "Tidak tercantum pada dokumen"
UNCONFIRMED = "Belum dapat dipastikan"


# Broker identities, including aliases used by historical 5% reports.
# Sources checked 2026-09-05:
# https://www.idx.co.id/id/anggota-bursa-dan-partisipan/profil-anggota-bursa/
# https://web.ksei.co.id/services/participants/brokers?setLocale=id-ID
# Bank names below include explicit user-requested account-code aliases.
# Other bank custodians retain their source names unless mapped explicitly.
BROKER_CODE_MAP = {
    # User-requested mappings for Rekening Efek.
    "PT BANK DBS INDONESIA": "DP",
    "BANK CENTRAL ASIA TBK": "SQ",
    "BANK RAKYAT INDONESIA (PERSERO) PT": "OD",
    "PT. AJAIB SEKURITAS ASIA": "XC",
    "PT AJAIB SEKURITAS ASIA": "XC",
    "PT. STOCKBIT SEKURITAS DIGITAL": "XL",
    "PT STOCKBIT SEKURITAS DIGITAL": "XL",
    "PT. TRIMEGAH SEKURITAS INDONESIA TBK": "LG",
    "PT TRIMEGAH SEKURITAS INDONESIA TBK": "LG",
    "PT INDO PREMIER SEKURITAS": "PD",
    "INDO PREMIER SEKURITAS, PT": "PD",
    "MANDIRI SEKURITAS, PT": "CC",
    "MANDIRI SEKURITAS": "CC",
    "PT BCA SEKURITAS": "SQ",
    "ERDIKHA ELIT, PT": "AO",
    "PT ERDIKHA ELIT SEKURITAS": "AO",
    "PT MIRAE ASSET SEKURITAS INDONESIA": "YP",
    "MIRAE ASSET SEKURITAS INDONESIA": "YP",
    "PT ARTHA SEKURITAS INDONESIA": "SH",
    "ARTHA SEKURITAS INDONESIA": "SH",
    "PT BUANA CAPITAL SEKURITAS": "RF",
    "BUANA CAPITAL SEKURITAS": "RF",
    "BNI SEKURITAS, PT": "NI",
    "PT BNI SEKURITAS": "NI",
    "PT MNC SEKURITAS": "EP",
    "MNC SEKURITAS, PT": "EP",
    "SINARMAS SEKURITAS, PT": "DH",
    "PT SINARMAS SEKURITAS": "DH",
    "PT RHB SEKURITAS INDONESIA": "DR",
    "PT CGS INTERNATIONAL SEKURITAS INDONESIA": "YU",
    "PT CGS INTERNATIONAL SEKURITAS": "YU",
    "CGS INTERNATIONAL SEKURITAS INDONESIA, PT": "YU",
    "PT SUCOR SEKURITAS": "AZ",
    "SUCOR SEKURITAS, PT": "AZ",
    "PT KOREA INVESTMENT AND SEKURITAS INDONESIA": "BQ",
    "PT KIWOOM SEKURITAS INDONESIA": "AG",
    "PT MAYBANK SEKURITAS INDONESIA": "ZP",
    "PT NH KORINDO SEKURITAS INDONESIA": "XA",
    "PT HENAN PUTIHRAI SEKURITAS": "HP",
    "SAMUEL SEKURITAS INDONESIA, PT": "IF",
    "PANIN SEKURITAS TBK, PT": "GR",
    "PANIN SEKURITAS": "GR",
    "PT VERDHANA SEKURITAS INDONESIA": "BB",
    "VERDHANA SEKURITAS INDONESIA, PT": "BB",
    "PT DBS VICKERS SEKURITAS INDONESIA": "DP",
    "DBS VICKERS SEKURITAS INDONESIA, PT": "DP",
    "PT HSBC SECURITIES INDONESIA": "GW",
    "HSBC SECURITIES INDONESIA, PT": "GW",
    "PT BRI DANAREKSA SEKURITAS": "OD",
    "BRI DANAREKSA SEKURITAS, PT": "OD",
    "PT BAHANA SEKURITAS": "DX",
    "BAHANA SEKURITAS, PT": "DX",
    "PT UOB KAY HIAN SEKURITAS": "AI",
    "UOB KAY HIAN SEKURITAS, PT": "AI",
    "PT J.P. MORGAN SEKURITAS INDONESIA": "BK",
    "PT JP MORGAN SEKURITAS INDONESIA": "BK",
    "J.P. MORGAN SEKURITAS INDONESIA, PT": "BK",
    "PT UBS SEKURITAS INDONESIA": "AK",
    "UBS SEKURITAS INDONESIA, PT": "AK",
    "PT MACQUARIE SEKURITAS INDONESIA": "RX",
    "MACQUARIE SEKURITAS INDONESIA, PT": "RX",
    "PT CLSA SEKURITAS INDONESIA": "KZ",
    "CLSA SEKURITAS INDONESIA, PT": "KZ",
    "PT KGI SEKURITAS INDONESIA": "HD",
    "KGI SEKURITAS INDONESIA, PT": "HD",
    "PT VALBURY SEKURITAS INDONESIA": "CP",
    "VALBURY SEKURITAS INDONESIA, PT": "CP",
    "PT EVERBRIGHT SEKURITAS INDONESIA": "EL",
    "EVERBRIGHT SEKURITAS INDONESIA, PT": "EL",
    "PT OCBC SEKURITAS INDONESIA": "TP",
    "OCBC SEKURITAS INDONESIA, PT": "TP",
    "PT PHILLIP SEKURITAS INDONESIA": "KK",
    "PHILLIP SEKURITAS INDONESIA, PT": "KK",
    "PT YUANTA SEKURITAS INDONESIA": "FS",
    "YUANTA SEKURITAS INDONESIA, PT": "FS",
    "PT VICTORIA SEKURITAS INDONESIA": "MI",
    "VICTORIA SEKURITAS INDONESIA, PT": "MI",
    "PT RELIANCE SEKURITAS INDONESIA TBK": "LS",
    "PT RELIANCE SEKURITAS INDONESIA": "LS",
    "RELIANCE SEKURITAS INDONESIA, PT": "LS",
    "PT PHINTRACO SEKURITAS": "AT",
    "PHINTRACO SEKURITAS, PT": "AT",
    "PT WATERFRONT SEKURITAS INDONESIA": "FZ",
    "WATERFRONT SEKURITAS INDONESIA, PT": "FZ",
    "PT INVESTINDO NUSANTARA SEKURITAS": "IN",
    "INVESTINDO NUSANTARA SEKURITAS, PT": "IN",
    "PT JASA UTAMA CAPITAL SEKURITAS": "YB",
    "JASA UTAMA CAPITAL SEKURITAS, PT": "YB",
    "PT FAC SEKURITAS INDONESIA": "PC",
    "FAC SEKURITAS INDONESIA, PT": "PC",
    "PT SEMESTA INDOVEST SEKURITAS": "MG",
    "SEMESTA INDOVEST SEKURITAS, PT": "MG",
    "PT ELIT SUKSES SEKURITAS": "SA",
    "ELIT SUKSES SEKURITAS, PT": "SA",
    "ALDIRACITA SEKURITAS INDONESIA": "PP",
    "AMANTARA SEKURITAS INDONESIA": "YO",
    "ANUGERAH SEKURITAS INDONESIA": "ID",
    "BINAARTHA SEKURITAS": "AR",
    "BNC SEKURITAS INDONESIA": "GA",
    "BUMIPUTERA SEKURITAS": "ZR",
    "CIPTADANA SEKURITAS ASIA": "KI",
    "DANASAKTI SEKURITAS INDONESIA": "PF",
    "DANATAMA MAKMUR SEKURITAS": "II",
    "DWIDANA SAKTI SEKURITAS": "TS",
    "EKOKAPITAL SEKURITAS": "ES",
    "EKUATOR SWARNA SEKURITAS": "MK",
    "EQUITY SEKURITAS INDONESIA": "BS",
    "EVERGREEN SEKURITAS INDONESIA": "EL",
    "FORTE GLOBAL SEKURITAS": "FO",
    "HARITA KENCANA SEKURITAS": "AF",
    "INA SEKURITAS INDONESIA": "RB",
    "INDO CAPITAL SEKURITAS": "IU",
    "INDO HARVEST SEKURITAS": "IH",
    "INTEGRITY CAPITAL SEKURITAS": "IC",
    "INTI FIKASA SEKURITAS": "BF",
    "INTI TELADAN SEKURITAS": "IT",
    "KAF SEKURITAS INDONESIA": "DU",
    "KAY HIAN SEKURITAS": "AI",
    "KB VALBURY SEKURITAS": "CP",
    "KRESNA SEKURITAS": "KS",
    "LABA SEKURITAS INDONESIA": "TF",
    "LOTUS ANDALAN SEKURITAS": "YJ",
    "MAGENTA KAPITAL SEKURITAS INDONESIA": "PI",
    "MAKINDO SEKURITAS": "DD",
    "MASINDO ARTHA SEKURITAS": "DM",
    "MEGA CAPITAL SEKURITAS": "CD",
    "MINNA PADI INVESTAMA SEKURITAS": "MU",
    "NET SEKURITAS": "OK",
    "PACIFIC SEKURITAS INDONESIA": "AP",
    "PANCA GLOBAL SEKURITAS": "PG",
    "PARAMITRA ALFA SEKURITAS": "PS",
    "PILARMAS INVESTINDO SEKURITAS": "PO",
    "PLUANG MAJU SEKURITAS": "RO",
    "PROFINDO SEKURITAS INDONESIA": "RG",
    "SHINHAN SEKURITAS INDONESIA": "AH",
    "SUKADANA PRIMA SEKURITAS": "AD",
    "SUPRA SEKURITAS INDONESIA": "SS",
    "SURYA FAJAR SEKURITAS": "SF",
    "TRUST SEKURITAS": "BR",
    "TUNTUN SEKURITAS INDONESIA": "QA",
    "WANTEG SEKURITAS": "AN",
    "WEBULL SEKURITAS INDONESIA": "GI",
    "YAKIN BERTUMBUH SEKURITAS": "YB",
    "YULIE SEKURITAS INDONESIA": "RS",
}


@dataclass(frozen=True)
class FivePercentRow:
    """One row extracted from a >=5% shareholder report."""

    source_file: str
    report_date: str
    data: dict[str, Any]


def analyze_five_percent_documents(
    documents_dir: Path,
    security_code: str,
    period_months: int | None = None,
    broker_code_map: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Read >=5% Excel files of the last N months and analyze the security code.

    ``period_months`` selects the newest N calendar months available (e.g. 2
    means the newest month plus the previous month).  When omitted, only the
    newest month is used, keeping the original behavior.
    """
    target = security_code.strip().upper()
    if not target:
        raise ValueError("Kode Efek wajib diisi.")
    if not documents_dir.exists():
        raise FileNotFoundError(f"Folder dokumen tidak ditemukan: {documents_dir}")

    xlsx_files = sorted(documents_dir.glob("*.xlsx"), key=_file_date_key)
    if not xlsx_files:
        return _empty_result(target, [f"Tidak ada file Excel di {documents_dir}."])

    latest_month = _file_month(xlsx_files[-1])
    month_files = [xlsx for xlsx in xlsx_files if _file_month(xlsx) == latest_month]
    if period_months and period_months > 1:
        earliest_month = _shift_month(latest_month, -(period_months - 1))
        files_to_parse = [
            xlsx for xlsx in xlsx_files
            if earliest_month <= _file_month(xlsx) <= latest_month
        ]
    else:
        files_to_parse = month_files
    if not files_to_parse:
        files_to_parse = [xlsx_files[-1]]

    rows_by_file: dict[str, list[FivePercentRow]] = {}
    errors: list[str] = []
    issuer_map: dict[str, str] = {}
    broker_map = _merged_broker_code_map(broker_code_map)
    for xlsx_file in files_to_parse:
        try:
            rows = _parse_excel_for_code(xlsx_file, target, broker_map)
            rows_by_file[xlsx_file.name] = rows
            for row in rows:
                code = str(row.data.get("KODE_EFEK", "")).upper()
                issuer = row.data.get("NAMA_EMITEN")
                if code and issuer and issuer != NOT_IN_DOCUMENT:
                    issuer_map.setdefault(code, str(issuer))
        except ImportError as exc:
            raise ImportError(
                "Library openpyxl belum tersedia. Jalankan `pip install -r requirements.txt`."
            ) from exc
        except Exception as exc:
            errors.append(f"{xlsx_file.name}: {exc}")

    files_with_rows = [
        xlsx for xlsx in files_to_parse
        if rows_by_file.get(xlsx.name)
    ]
    latest_file = max(files_with_rows, key=_file_date_key) if files_with_rows else max(month_files, key=_file_date_key)
    latest_rows = [
        _with_issuer_fallback(row, issuer_map)
        for row in rows_by_file.get(latest_file.name, [])
        if str(row.data.get("KODE_EFEK", "")).upper() == target
    ]
    broker_sets = _known_broker_sets_by_holder(rows_by_file, target)
    preferred_brokers = _single_known_brokers(broker_sets)
    crossing_holders = _crossing_holders(broker_sets)
    latest_rows = _combine_unavailable_broker_rows(
        latest_rows,
        preferred_brokers,
        crossing_holders,
    )
    latest_rows.sort(
        key=lambda row: (
            str(row.data.get("NAMA_PEMEGANG_SAHAM", "")),
            str(row.data.get("NAMA_PEMEGANG_REKENING_EFEK", "")),
        )
    )

    latest_report_date = _latest_report_date(latest_rows) or _date_from_filename(latest_file)
    changes = _build_latest_changes(rows_by_file, target, issuer_map)

    if period_months and period_months > 1:
        period_label = f"{period_months} bulan terakhir"
        comparison_note = (
            "Analisa akumulasi/distribusi membandingkan kepemilikan pada tanggal "
            "laporan paling awal di periode terpilih (awal bulan sebelumnya) "
            "dengan tanggal laporan terbaru."
        )
    else:
        period_label = "bulan terbaru"
        comparison_note = (
            "Analisa akumulasi/distribusi membandingkan tanggal laporan paling "
            "awal dengan tanggal laporan terbaru pada bulan tersebut."
        )

    notes = [
        f"Data utama memakai laporan terbaru dari {period_label} yang tersedia.",
        comparison_note,
        "Pencarian hanya memakai Kode Efek.",
        "Parser Excel dipakai untuk membaca baris file terbaru agar baris tidak tergabung dengan emiten lain.",
    ]
    if errors:
        notes.append("Sebagian dokumen gagal dibaca: " + "; ".join(errors))
    if not latest_rows:
        notes.append(f"Tidak ada data untuk Kode Efek tersebut pada laporan terbaru {period_label}.")

    return {
        "security_code": target,
        "period_months": period_months or "ALL",
        "latest_month": latest_month,
        "latest_file": latest_file.name,
        "latest_report_date": latest_report_date,
        "rows": [row.data | {"SOURCE_FILE": row.source_file} for row in latest_rows],
        "changes": changes,
        "documents_read": [xlsx.name for xlsx in files_to_parse],
        "notes": notes,
    }


def scan_five_percent_documents(
    documents_dir: Path,
    period_months: int,
    broker_code_map: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Scan every >=5% issuer and return only ownership movements."""
    if period_months not in {1, 2, 3, 6, 12}:
        raise ValueError("Periode scanner harus 1, 2, 3, 6, atau 12.")
    if not documents_dir.exists():
        raise FileNotFoundError(f"Folder dokumen tidak ditemukan: {documents_dir}")

    xlsx_files = sorted(documents_dir.glob("*.xlsx"), key=_file_date_key)
    if not xlsx_files:
        return {
            "period_months": period_months,
            "changes": [],
            "documents_read": [],
            "notes": [f"Tidak ada file Excel di {documents_dir}."],
            "total_rows_read": 0,
            "total_issuers_scanned": 0,
            "total_changes": 0,
        }

    latest_month = _file_month(xlsx_files[-1])
    if period_months == 1:
        # The newest report already contains its immediately preceding
        # ownership snapshot, so one file is sufficient for the 2-day scan.
        files_to_parse = [xlsx_files[-1]]
    else:
        earliest_month = _shift_month(latest_month, -(period_months - 1))
        files_to_parse = [
            xlsx for xlsx in xlsx_files
            if earliest_month <= _file_month(xlsx) <= latest_month
        ]
    if not files_to_parse:
        files_to_parse = [xlsx_files[-1]]

    rows_by_code: dict[str, dict[str, list[FivePercentRow]]] = {}
    issuer_map: dict[str, str] = {}
    errors: list[str] = []
    total_rows_read = 0
    broker_map = _merged_broker_code_map(broker_code_map)

    for xlsx_file in files_to_parse:
        try:
            rows = _parse_excel_for_code(xlsx_file, None, broker_map)
        except ImportError as exc:
            raise ImportError(
                "Library openpyxl belum tersedia. Jalankan `pip install -r requirements.txt`."
            ) from exc
        except Exception as exc:
            errors.append(f"{xlsx_file.name}: {exc}")
            continue

        total_rows_read += len(rows)
        for row in rows:
            code = str(row.data.get("KODE_EFEK") or "").upper()
            if not code:
                continue
            rows_by_code.setdefault(code, {}).setdefault(xlsx_file.name, []).append(row)
            issuer = row.data.get("NAMA_EMITEN")
            if issuer not in (None, "", NOT_IN_DOCUMENT):
                issuer_map.setdefault(code, str(issuer))

    changes: list[dict[str, Any]] = []
    for code in sorted(rows_by_code):
        code_changes = _build_latest_changes(rows_by_code[code], code, issuer_map)
        for change in code_changes:
            delta = change.get("CHANGE_SHARES")
            if not isinstance(delta, (int, float)) or delta == 0:
                continue
            changes.append(
                {
                    "KODE_EFEK": code,
                    "NAMA_EMITEN": issuer_map.get(code, NOT_IN_DOCUMENT),
                    **change,
                }
            )

    changes.sort(
        key=lambda change: (
            str(change.get("KODE_EFEK", "")),
            str(change.get("NAMA_PEMEGANG_SAHAM", "")),
            str(change.get("NAMA_PEMEGANG_REKENING_EFEK", "")),
        )
    )

    available_months = sorted({_file_month(path) for path in files_to_parse})
    dates_scanned = sorted(
        {
            row.report_date
            for files in rows_by_code.values()
            for rows in files.values()
            for row in rows
        }
        | {
            str(change.get(date_key))
            for change in changes
            for date_key in ("FROM_DATE", "TO_DATE")
            if change.get(date_key) not in (None, "", UNKNOWN, UNCONFIRMED)
        },
        key=_date_key,
    )
    notes = [
        "Scanner membaca seluruh kode emiten tanpa meminta Kode Efek.",
        "Hanya perubahan kepemilikan yang ditampilkan; posisi netral disembunyikan.",
    ]
    if period_months == 1:
        notes.append(
            "Periode 2 Hari membandingkan tanggal laporan terbaru dengan tanggal sebelumnya."
        )
    else:
        notes.append(
            "Perubahan membandingkan posisi paling awal dengan posisi terbaru dalam periode terpilih."
        )
    if period_months > 1 and len(available_months) < period_months:
        notes.append(
            f"Periode {period_months} bulan dipilih, tetapi hanya "
            f"{len(available_months)} bulan laporan yang tersedia."
        )
    if errors:
        notes.append("Sebagian dokumen gagal dibaca: " + "; ".join(errors))

    return {
        "period_months": period_months,
        "latest_month": latest_month,
        "dates_scanned": dates_scanned,
        "changes": changes,
        "documents_read": [xlsx.name for xlsx in files_to_parse],
        "notes": notes,
        "total_rows_read": total_rows_read,
        "total_issuers_scanned": len(rows_by_code),
        "total_changes": len(changes),
    }


def _empty_result(security_code: str, notes: list[str]) -> dict[str, Any]:
    return {
        "security_code": security_code,
        "period_months": "ALL",
        "latest_month": UNKNOWN,
        "latest_file": UNKNOWN,
        "latest_report_date": UNKNOWN,
        "rows": [],
        "changes": [],
        "documents_read": [],
        "notes": notes,
    }


def _parse_excel_for_code(
    excel_file: Path,
    security_code: str | None,
    broker_code_map: dict[str, str] | None = None,
) -> list[FivePercentRow]:
    """Parse the KSEI >=5% Excel report, optionally filtering Kode Efek.

    The report uses real table columns (No, Kode Efek, Nama Emiten, Nama
    Pemegang Rekening Efek, Nama Pemegang Saham, Nama Rekening Efek, ...,
    Kepemilikan Per <tanggal-awal> x3, Kepemilikan Per <tanggal-akhir> x3,
    Perubahan), so each row maps directly to the analysis fields.
    """
    target = security_code.strip().upper() if security_code else None
    try:
        import openpyxl  # type: ignore
    except ImportError:
        raise ImportError(
            "Library openpyxl belum tersedia. Jalankan `pip install -r requirements.txt`."
        ) from None

    wb = openpyxl.load_workbook(str(excel_file), read_only=True, data_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        rows: list[FivePercentRow] = []
        previous_report_date: str | None = None
        report_date: str | None = None
        current_code = ""
        current_issuer = ""
        current_shareholder = ""
        issuer_by_code: dict[str, str] = {}
        data_started = False

        for raw_row in ws.iter_rows(values_only=True):
            row = [_excel_cell(cell) for cell in raw_row]
            if not data_started:
                if _table_cell(row, 1).upper() == "KODE EFEK":
                    previous_report_date, report_date = _excel_report_dates(row)
                    data_started = True
                continue

            if _is_footer_row(row):
                continue
            if not any(row):
                continue

            code = _table_cell(row, 1).upper() or current_code
            if _table_cell(row, 1):
                current_code = code

            issuer_name = _table_cell(row, 2) or current_issuer or NOT_IN_DOCUMENT
            broker_name = _table_cell(row, 3) or NOT_IN_DOCUMENT
            shareholder_name = _table_cell(row, 4) or current_shareholder
            account_name = _table_cell(row, 5) or shareholder_name or NOT_IN_DOCUMENT
            domicile = _table_cell(row, 9) or NOT_IN_DOCUMENT
            local_foreign = _local_foreign_label(_table_cell(row, 10))
            previous_holding = _table_int_or_unknown(row, 11)
            holding = _table_int_or_unknown(row, 14)
            combined_holding = _table_int_or_unknown(row, 15)
            percentage = _table_cell(row, 16) or NOT_IN_DOCUMENT
            change_shares = _to_int(_table_cell(row, 17))

            if issuer_name != NOT_IN_DOCUMENT:
                current_issuer = issuer_name
                issuer_by_code.setdefault(code, issuer_name)
            if shareholder_name:
                current_shareholder = shareholder_name

            if not code or not shareholder_name:
                continue
            if target and code != target:
                continue
            if not isinstance(holding, int):
                continue

            broker_code = _broker_code(broker_name, broker_code_map)
            data = {
                "KODE_EFEK": code,
                "NAMA_EMITEN": issuer_name,
                "NAMA_PEMEGANG_REKENING_EFEK": broker_code,
                "NAMA_PEMEGANG_REKENING_EFEK_ASLI": broker_name,
                "NAMA_PEMEGANG_SAHAM": shareholder_name,
                "NAMA_REKENING_EFEK": account_name,
                "BROKER_NAME": account_name,
                "BROKER_CODE": broker_code,
                "DOMISILI": domicile,
                "LOCAL_FOREIGN": local_foreign,
                "KEPEMILIKAN": holding,
                "JUMLAH_SAHAM": holding,
                "SAHAM_GABUNGAN_PER_INVESTOR": combined_holding,
                "PERSENTASE": percentage,
                "STATUS_KEPEMILIKAN": NOT_IN_DOCUMENT,
                "BLOCKING_REASON": NOT_IN_DOCUMENT,
                "PREVIOUS_HOLDING": previous_holding,
                "PREVIOUS_REPORT_DATE": previous_report_date or UNCONFIRMED,
                "CHANGE_SHARES": change_shares,
                "REPORT_DATE": report_date or _date_from_filename(excel_file),
                "ROW_LEVEL": "AGGREGATE" if _table_cell(row, 2) else "DETAIL",
            }
            rows.append(
                FivePercentRow(
                    excel_file.name,
                    report_date or _date_from_filename(excel_file),
                    data,
                )
            )
        return rows
    finally:
        wb.close()


def _excel_cell(value: Any) -> str:
    """Convert an Excel cell value to a clean string (dates to YYYY-MM-DD)."""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    return _clean(value)


def _excel_report_dates(header_row: list[str]) -> tuple[str | None, str | None]:
    """Extract previous/latest report dates from the 'Kepemilikan Per' header cells."""
    found: list[str] = []
    for cell in header_row:
        if "KEPEMILIKAN PER" in cell.upper():
            date = _report_date_from_text(cell)
            if date:
                found.append(date)
    found = sorted(set(found), key=_date_key)
    if len(found) >= 2:
        return found[-2], found[-1]
    if found:
        return None, found[-1]
    return None, None


def _is_footer_row(row: list[str]) -> bool:
    first = _table_cell(row, 0).upper()
    return first.startswith("*") or first.startswith("HITAM") or first.startswith("BIRU")


def _table_int_or_unknown(row: list[str], idx: int) -> int | str:
    return _to_int(_table_cell(row, idx))


def _local_foreign_label(value: str) -> str:
    upper = value.upper()
    if upper == "L":
        return "LOCAL"
    if upper == "A":
        return "ASING"
    return value or NOT_IN_DOCUMENT


def _table_cell(row: list[str], idx: int) -> str:
    return row[idx] if idx < len(row) else ""


def _build_latest_changes(
    rows_by_file: dict[str, list[FivePercentRow]],
    security_code: str,
    issuer_map: dict[str, str],
) -> list[dict[str, Any]]:
    dated_rows: list[tuple[str, list[FivePercentRow]]] = []
    for rows in rows_by_file.values():
        filtered = [
            _with_issuer_fallback(row, issuer_map)
            for row in rows
            if str(row.data.get("KODE_EFEK", "")).upper() == security_code
        ]
        if filtered:
            dated_rows.append((_latest_report_date(filtered) or filtered[0].report_date, filtered))
    dated_rows.sort(key=lambda item: _date_key(item[0]))
    if dated_rows:
        earliest_snapshot = _earliest_previous_snapshot(dated_rows[0][1])
        if earliest_snapshot is not None and _date_key(earliest_snapshot[0]) < _date_key(dated_rows[0][0]):
            dated_rows.insert(0, earliest_snapshot)
    if len(dated_rows) < 2:
        if not dated_rows:
            return []
        latest_date, latest_rows = dated_rows[-1]
        all_dates = sorted(
            {row.report_date for rows in rows_by_file.values() for row in rows},
            key=_date_key,
        )
        previous_date = all_dates[0] if all_dates else UNCONFIRMED
        latest_map = _aggregate_by_holder_broker(latest_rows)
        changes = []
        for item in latest_map.values():
            holding = int(item.get("KEPEMILIKAN") or 0)
            changes.append(
                {
                    "NAMA_PEMEGANG_SAHAM": item.get("NAMA_PEMEGANG_SAHAM", UNKNOWN),
                    "NAMA_PEMEGANG_REKENING_EFEK": item.get("NAMA_PEMEGANG_REKENING_EFEK", UNKNOWN),
                    "FROM_DATE": previous_date,
                    "TO_DATE": latest_date,
                    "PREVIOUS_HOLDING": 0,
                    "LATEST_HOLDING": holding,
                    "CHANGE_SHARES": holding,
                    "ESTIMATED_LOTS": holding / 100,
                    "STATUS": "Akumulasi",
                }
            )
        return sorted(_merge_equal_holding_unknown_broker_changes(changes), key=_change_sort_key)

    previous_date, previous_rows = dated_rows[0]
    latest_date, latest_rows = dated_rows[-1]
    broker_sets = _known_broker_sets_from_dated_rows(dated_rows)
    preferred_brokers = _single_known_brokers(broker_sets)
    crossing_holders = _crossing_holders(broker_sets)
    previous_map = _aggregate_by_holder_broker(previous_rows, crossing_holders)
    latest_map = _aggregate_by_holder_broker(latest_rows, crossing_holders)
    previous_map, latest_map = _collapse_unavailable_broker_counterparts(
        previous_map,
        latest_map,
        preferred_brokers,
        crossing_holders,
    )

    changes: list[dict[str, Any]] = []
    for key in sorted(set(previous_map) | set(latest_map)):
        prev = previous_map.get(key)
        latest = latest_map.get(key)
        prev_holding = int(prev["KEPEMILIKAN"]) if prev else 0
        latest_holding = int(latest["KEPEMILIKAN"]) if latest else 0
        delta = latest_holding - prev_holding
        status = (
            "Akumulasi" if delta > 0 else
            "Distribusi" if delta < 0 else
            "Netral"
        )
        sample = latest or prev or {}
        changes.append(
            {
                "NAMA_PEMEGANG_SAHAM": sample.get("NAMA_PEMEGANG_SAHAM", UNKNOWN),
                "NAMA_PEMEGANG_REKENING_EFEK": sample.get("NAMA_PEMEGANG_REKENING_EFEK", UNKNOWN),
                "FROM_DATE": previous_date,
                "TO_DATE": latest_date,
                "PREVIOUS_HOLDING": prev_holding,
                "LATEST_HOLDING": latest_holding,
                "CHANGE_SHARES": delta,
                "ESTIMATED_LOTS": delta / 100,
                "STATUS": status,
            }
        )
    return sorted(_merge_equal_holding_unknown_broker_changes(changes), key=_change_sort_key)


def _merge_equal_holding_unknown_broker_changes(
    changes: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Merge offsetting rows when broker is missing on one side only.

    If holder A has 10,000 shares under broker AO in the previous date and the
    latest date only exposes the same 10,000 shares without a broker code, keep
    one neutral AO row.  If both sides have clear but different brokers, leave
    them split because that indicates broker movement/crossing.
    """
    used: set[int] = set()
    merged: list[dict[str, Any]] = []

    for i, left in enumerate(changes):
        if i in used:
            continue
        match_index = _find_equal_holding_unknown_broker_pair(i, left, changes, used)
        if match_index is None:
            continue

        right = changes[match_index]
        used.add(i)
        used.add(match_index)
        merged.append(_merge_equal_holding_change(left, right))

    merged.extend(item for idx, item in enumerate(changes) if idx not in used)
    return merged


def _find_equal_holding_unknown_broker_pair(
    left_index: int,
    left: dict[str, Any],
    changes: list[dict[str, Any]],
    used: set[int],
) -> int | None:
    left_holder = str(left.get("NAMA_PEMEGANG_SAHAM", "")).strip().upper()
    left_prev = left.get("PREVIOUS_HOLDING")
    left_latest = left.get("LATEST_HOLDING")
    left_broker = str(left.get("NAMA_PEMEGANG_REKENING_EFEK", "")).strip().upper()

    for right_index in range(left_index + 1, len(changes)):
        if right_index in used:
            continue
        right = changes[right_index]
        right_holder = str(right.get("NAMA_PEMEGANG_SAHAM", "")).strip().upper()
        if right_holder != left_holder:
            continue

        right_prev = right.get("PREVIOUS_HOLDING")
        right_latest = right.get("LATEST_HOLDING")
        right_broker = str(right.get("NAMA_PEMEGANG_REKENING_EFEK", "")).strip().upper()

        if not _is_offsetting_equal_holding(left_prev, left_latest, right_prev, right_latest):
            continue
        if _brokers_conflict(left_broker, right_broker):
            continue
        return right_index
    return None


def _is_offsetting_equal_holding(
    left_prev: Any,
    left_latest: Any,
    right_prev: Any,
    right_latest: Any,
) -> bool:
    return (
        isinstance(left_prev, int)
        and isinstance(left_latest, int)
        and isinstance(right_prev, int)
        and isinstance(right_latest, int)
        and (
            (left_prev > 0 and left_latest == 0 and right_prev == 0 and right_latest == left_prev)
            or (right_prev > 0 and right_latest == 0 and left_prev == 0 and left_latest == right_prev)
        )
    )


def _brokers_conflict(left_broker: str, right_broker: str) -> bool:
    return _is_known_broker(left_broker) and _is_known_broker(right_broker) and left_broker != right_broker


def _merge_equal_holding_change(
    left: dict[str, Any],
    right: dict[str, Any],
) -> dict[str, Any]:
    known_broker = _known_broker_from_pair(left, right)
    holding = max(
        int(left.get("PREVIOUS_HOLDING") or 0),
        int(left.get("LATEST_HOLDING") or 0),
        int(right.get("PREVIOUS_HOLDING") or 0),
        int(right.get("LATEST_HOLDING") or 0),
    )
    sample = left if _is_known_broker(str(left.get("NAMA_PEMEGANG_REKENING_EFEK", ""))) else right
    return {
        "NAMA_PEMEGANG_SAHAM": sample.get("NAMA_PEMEGANG_SAHAM", UNKNOWN),
        "NAMA_PEMEGANG_REKENING_EFEK": known_broker,
        "FROM_DATE": left.get("FROM_DATE") or right.get("FROM_DATE"),
        "TO_DATE": left.get("TO_DATE") or right.get("TO_DATE"),
        "PREVIOUS_HOLDING": holding,
        "LATEST_HOLDING": holding,
        "CHANGE_SHARES": 0,
        "ESTIMATED_LOTS": 0,
        "STATUS": "Netral",
    }


def _known_broker_from_pair(left: dict[str, Any], right: dict[str, Any]) -> str:
    for item in (left, right):
        broker = str(item.get("NAMA_PEMEGANG_REKENING_EFEK", "")).strip().upper()
        if _is_known_broker(broker):
            return broker
    return UNCONFIRMED


def _aggregate_by_holder_broker(
    rows: list[FivePercentRow],
    crossing_holders: set[str] | None = None,
) -> dict[tuple[str, str], dict[str, Any]]:
    crossing_holders = crossing_holders or set()
    result: dict[tuple[str, str], dict[str, Any]] = {}
    for index, row in enumerate(rows):
        holder = str(row.data.get("NAMA_PEMEGANG_SAHAM", UNKNOWN)).strip().upper()
        account = str(row.data.get("NAMA_PEMEGANG_REKENING_EFEK", UNKNOWN)).strip().upper()
        if holder in crossing_holders and _is_unavailable_broker(account):
            key = (holder, f"{account}#{index}")
        else:
            key = _holder_broker_key(holder, account)
        if key not in result:
            result[key] = dict(row.data)
            if _is_unavailable_broker(account) and holder not in crossing_holders:
                result[key]["NAMA_PEMEGANG_REKENING_EFEK"] = UNCONFIRMED
        else:
            existing = result[key]
            existing["KEPEMILIKAN"] = int(existing.get("KEPEMILIKAN") or 0) + int(row.data.get("KEPEMILIKAN") or 0)
    return result


def _collapse_unavailable_broker_counterparts(
    previous_map: dict[tuple[str, str], dict[str, Any]],
    latest_map: dict[tuple[str, str], dict[str, Any]],
    preferred_brokers: dict[str, str] | None = None,
    crossing_holders: set[str] | None = None,
) -> tuple[dict[tuple[str, str], dict[str, Any]], dict[tuple[str, str], dict[str, Any]]]:
    crossing_holders = crossing_holders or set()
    holders_to_collapse = {
        holder
        for holder, broker in set(previous_map) | set(latest_map)
        if holder not in crossing_holders
        and (broker == UNCONFIRMED or _is_unavailable_broker(broker))
    }
    if not holders_to_collapse:
        return previous_map, latest_map
    return (
        _collapse_map_by_holder(previous_map, holders_to_collapse, preferred_brokers or {}),
        _collapse_map_by_holder(latest_map, holders_to_collapse, preferred_brokers or {}),
    )


def _collapse_map_by_holder(
    source: dict[tuple[str, str], dict[str, Any]],
    holders_to_collapse: set[str],
    preferred_brokers: dict[str, str],
) -> dict[tuple[str, str], dict[str, Any]]:
    result: dict[tuple[str, str], dict[str, Any]] = {}
    for (holder, broker), item in source.items():
        preferred_broker = preferred_brokers.get(holder, UNCONFIRMED)
        key = (holder, preferred_broker) if holder in holders_to_collapse else (holder, broker)
        if key not in result:
            result[key] = dict(item)
            if holder in holders_to_collapse:
                result[key]["NAMA_PEMEGANG_REKENING_EFEK"] = preferred_broker
            continue
        result[key]["KEPEMILIKAN"] = int(result[key].get("KEPEMILIKAN") or 0) + int(item.get("KEPEMILIKAN") or 0)
    return result


def _combine_unavailable_broker_rows(
    rows: list[FivePercentRow],
    preferred_brokers: dict[str, str] | None = None,
    crossing_holders: set[str] | None = None,
) -> list[FivePercentRow]:
    preferred_brokers = preferred_brokers or {}
    crossing_holders = crossing_holders or set()
    combined: dict[tuple[str, str], FivePercentRow] = {}
    passthrough: list[FivePercentRow] = []
    for row in rows:
        holder = str(row.data.get("NAMA_PEMEGANG_SAHAM", UNKNOWN)).strip().upper()
        broker = str(row.data.get("NAMA_PEMEGANG_REKENING_EFEK", UNKNOWN)).strip().upper()
        if not _is_unavailable_broker(broker) or holder in crossing_holders:
            passthrough.append(row)
            continue

        key = (holder, row.report_date)
        if key not in combined:
            data = dict(row.data)
            preferred_broker = preferred_brokers.get(holder, UNCONFIRMED)
            data["NAMA_PEMEGANG_REKENING_EFEK"] = preferred_broker
            if preferred_broker != UNCONFIRMED:
                data["NAMA_REKENING_EFEK"] = preferred_broker
            else:
                data["NAMA_REKENING_EFEK"] = UNCONFIRMED
            combined[key] = FivePercentRow(row.source_file, row.report_date, data)
            continue

        existing = combined[key]
        data = dict(existing.data)
        _sum_int_field(data, row.data, "KEPEMILIKAN")
        _sum_int_field(data, row.data, "JUMLAH_SAHAM")
        _sum_int_field(data, row.data, "PREVIOUS_HOLDING")
        _sum_int_field(data, row.data, "CHANGE_SHARES")
        if data.get("SAHAM_GABUNGAN_PER_INVESTOR") == NOT_IN_DOCUMENT:
            data["SAHAM_GABUNGAN_PER_INVESTOR"] = row.data.get("SAHAM_GABUNGAN_PER_INVESTOR", NOT_IN_DOCUMENT)
        if data.get("PERSENTASE") == NOT_IN_DOCUMENT:
            data["PERSENTASE"] = row.data.get("PERSENTASE", NOT_IN_DOCUMENT)
        combined[key] = FivePercentRow(existing.source_file, existing.report_date, data)

    return passthrough + list(combined.values())


def _known_broker_sets_by_holder(
    rows_by_file: dict[str, list[FivePercentRow]],
    security_code: str,
) -> dict[str, set[str]]:
    dated_rows: list[tuple[str, list[FivePercentRow]]] = []
    for rows in rows_by_file.values():
        filtered = [
            row for row in rows
            if str(row.data.get("KODE_EFEK", "")).upper() == security_code
        ]
        if filtered:
            dated_rows.append((_latest_report_date(filtered) or filtered[0].report_date, filtered))
    dated_rows.sort(key=lambda item: _date_key(item[0]))
    return _known_broker_sets_from_dated_rows(dated_rows)


def _known_broker_sets_from_dated_rows(
    dated_rows: list[tuple[str, list[FivePercentRow]]],
) -> dict[str, set[str]]:
    brokers: dict[str, set[str]] = {}
    for _, rows in dated_rows:
        for row in rows:
            holder = str(row.data.get("NAMA_PEMEGANG_SAHAM", UNKNOWN)).strip().upper()
            broker = str(row.data.get("NAMA_PEMEGANG_REKENING_EFEK", "")).strip().upper()
            if holder and _is_known_broker(broker):
                brokers.setdefault(holder, set()).add(broker)
    return brokers


def _single_known_brokers(broker_sets: dict[str, set[str]]) -> dict[str, str]:
    return {
        holder: next(iter(brokers))
        for holder, brokers in broker_sets.items()
        if len(brokers) == 1
    }


def _crossing_holders(broker_sets: dict[str, set[str]]) -> set[str]:
    return {
        holder
        for holder, brokers in broker_sets.items()
        if len(brokers) > 1
    }


def _sum_int_field(target: dict[str, Any], source: dict[str, Any], field: str) -> None:
    left = target.get(field)
    right = source.get(field)
    if isinstance(left, int) and isinstance(right, int):
        target[field] = left + right


def _holder_broker_key(holder: str, account: str) -> tuple[str, str]:
    if _is_unavailable_broker(account):
        return holder, UNCONFIRMED
    return holder, account


def _is_unavailable_broker(value: str) -> bool:
    upper = str(value or "").strip().upper()
    return upper == UNCONFIRMED.upper() or upper.startswith("KODE BROKER TIDAK TERSEDIA")


def _is_known_broker(value: str) -> bool:
    # An unmapped account-holder name still identifies a distinct account.
    # Keep it separate from other names/codes during ownership comparisons.
    broker = str(value or "").strip().upper()
    return bool(broker) and broker not in {UNKNOWN.upper(), NOT_IN_DOCUMENT.upper(), UNCONFIRMED.upper()} and not _is_unavailable_broker(broker)


def _change_sort_key(item: dict[str, Any]) -> tuple[int, float, str]:
    status_order = {"Akumulasi": 0, "Distribusi": 1, "Netral": 2}
    return (
        status_order.get(str(item.get("STATUS")), 99),
        -abs(float(item.get("CHANGE_SHARES") or 0)),
        str(item.get("NAMA_PEMEGANG_SAHAM", "")),
    )


def _earliest_previous_snapshot(
    rows: list[FivePercentRow],
) -> tuple[str, list[FivePercentRow]] | None:
    """Re-anchor rows of the earliest file to its previous snapshot date.

    Every >=5% report carries two snapshots ("Kepemilikan Per <awal>" and
    "Kepemilikan Per <akhir>").  For the comparison baseline the earliest file
    of the period is anchored to its first snapshot so the analysis starts from
    the beginning of the previous month rather than from zero.
    """
    anchored: list[FivePercentRow] = []
    snapshot_date: str | None = None
    for row in rows:
        previous_holding = row.data.get("PREVIOUS_HOLDING")
        previous_date = row.data.get("PREVIOUS_REPORT_DATE")
        if not isinstance(previous_date, str):
            continue
        if _date_key(previous_date) == datetime.min:
            continue
        snapshot_date = previous_date
        if not isinstance(previous_holding, int):
            # A blank previous holding means the holder was not yet present.
            # Keep the snapshot date; absence from ``anchored`` represents 0.
            continue
        data = dict(row.data)
        data["KEPEMILIKAN"] = previous_holding
        data["JUMLAH_SAHAM"] = previous_holding
        data["REPORT_DATE"] = previous_date
        anchored.append(FivePercentRow(row.source_file, previous_date, data))
    if snapshot_date is None:
        return None
    return snapshot_date, anchored


def _with_issuer_fallback(
    row: FivePercentRow,
    issuer_map: dict[str, str],
) -> FivePercentRow:
    code = str(row.data.get("KODE_EFEK", "")).upper()
    if row.data.get("NAMA_EMITEN") != NOT_IN_DOCUMENT or code not in issuer_map:
        return row
    data = dict(row.data)
    data["NAMA_EMITEN"] = issuer_map[code]
    return FivePercentRow(row.source_file, row.report_date, data)


def _merged_broker_code_map(
    broker_code_map: dict[str, str] | None = None,
) -> dict[str, str]:
    return dict(_normalized_broker_map(tuple((broker_code_map or {}).items())))


def _broker_code(
    broker_name: str,
    broker_code_map: dict[str, str] | None = None,
) -> str:
    broker_map = _normalized_broker_map(tuple((broker_code_map or {}).items()))
    return broker_map.get(_normalize_broker_name(broker_name), broker_name)


def _normalize_broker_name(name: str) -> str:
    """Ignore legal affixes, punctuation and whitespace, not entity names."""
    name = str(name).upper().replace(".", "")
    tokens = re.sub(r"[^\w\s]", " ", name).split()
    while tokens and tokens[0] == "PT":
        tokens.pop(0)
    while tokens and tokens[-1] in {"PT", "TBK"}:
        tokens.pop()
    return " ".join(tokens)


@lru_cache(maxsize=16)
def _normalized_broker_map(overrides: tuple[tuple[str, str], ...]) -> dict[str, str]:
    """Cache the lookup per configuration; explicit aliases override defaults."""
    merged = {
        _normalize_broker_name(name): code for name, code in BROKER_CODE_MAP.items()
    }
    for name, code in overrides:
        clean_name = _normalize_broker_name(name)
        clean_code = _clean(code).upper()
        if clean_name and clean_code:
            merged[clean_name] = clean_code
    return merged


def _clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("\n", " ")).strip()


def _to_int(value: Any) -> int | str:
    cleaned = re.sub(r"[^\d-]", "", str(value or ""))
    if not cleaned or cleaned == "-":
        return NOT_IN_DOCUMENT
    return int(cleaned)


def _file_month(path: Path) -> str:
    date_text = _date_from_filename(path)
    return date_text[:7] if date_text != UNKNOWN else UNKNOWN


def _shift_month(month: str, delta: int) -> str:
    """Return the month ``delta`` months away from an 'YYYY-MM' string."""
    try:
        year, month_number = int(month[:4]), int(month[5:7])
    except (ValueError, TypeError):
        return month
    total = year * 12 + (month_number - 1) + delta
    return f"{total // 12:04d}-{total % 12 + 1:02d}"


def _file_date_key(path: Path) -> datetime:
    return _date_key(_date_from_filename(path))


def _date_from_filename(path: Path) -> str:
    return _date_from_text(path.stem) or UNKNOWN


def _report_date_from_text(text: str) -> str | None:
    match = re.search(r"per(?:\s+tanggal)?\s+(\d{1,2})\s*-?\s*([A-Za-z]+)\s*-?\s*(\d{4})", text, flags=re.IGNORECASE)
    if match:
        return _date_from_parts(match.group(3), match.group(2), match.group(1))
    return _date_from_text(text)


def _date_from_text(text: str) -> str | None:
    match = re.search(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", text)
    if match:
        year, month, day = map(int, match.groups())
        return f"{year:04d}-{month:02d}-{day:02d}"
    match = re.search(r"(\d{1,2})[-\s/]([A-Za-z]+)[-\s/](\d{4})", text)
    if match:
        return _date_from_parts(match.group(3), match.group(2), match.group(1))
    return None


def _date_from_parts(year_text: str, month_text: str, day_text: str) -> str | None:
    month_map = {
        "jan": 1, "january": 1, "januari": 1,
        "feb": 2, "february": 2, "februari": 2,
        "mar": 3, "march": 3, "maret": 3,
        "apr": 4, "april": 4,
        "may": 5, "mei": 5,
        "jun": 6, "june": 6, "juni": 6,
        "jul": 7, "july": 7, "juli": 7,
        "aug": 8, "august": 8, "agu": 8, "agustus": 8,
        "sep": 9, "sept": 9, "september": 9,
        "oct": 10, "october": 10, "okt": 10, "oktober": 10,
        "nov": 11, "november": 11,
        "dec": 12, "december": 12, "des": 12, "desember": 12,
    }
    month = month_map.get(month_text.lower())
    if not month:
        return None
    return f"{int(year_text):04d}-{month:02d}-{int(day_text):02d}"


def _date_key(value: Any) -> datetime:
    try:
        return datetime.strptime(str(value), "%Y-%m-%d")
    except ValueError:
        return datetime.min


def _latest_report_date(rows: list[FivePercentRow]) -> str | None:
    dates = sorted({row.report_date for row in rows}, key=_date_key)
    return dates[-1] if dates else None
