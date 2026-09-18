import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from traceterminal.analysis.shareholder5 import (
    FivePercentRow,
    _broker_code,
    _build_latest_changes,
    _combine_unavailable_broker_rows,
    _merged_broker_code_map,
    _parse_excel_for_code,
)


class BrokerAccountTests(unittest.TestCase):
    def test_report_name_variants_resolve_to_verified_codes(self):
        examples = {
            "PT. KGI Sekuritas Indonesia": "HD",
            "KGI SEKURITAS INDONESIA, PT.": "HD",
            "  pt.  KGI\nSekuritas Indonesia  ": "HD",
            "PT KB VALBURY SEKURITAS": "CP",
            "PT KAY HIAN SEKURITAS": "AI",
            "PT UOB KAY HIAN SEKURITAS": "AI",
            "PT CIPTADANA SEKURITAS ASIA": "KI",
            "PT Anugerah Sekuritas Indonesia": "ID",
            "PT ARTHA SEKURITAS INDONESIA": "SH",
            "BUMIPUTERA SEKURITAS, PT": "ZR",
            "PT Yakin Bertumbuh Sekuritas": "YB",
            "PT JASA UTAMA CAPITAL SEKURITAS": "YB",
            "PT. Laba Sekuritas Indonesia": "TF",
            "PT MINNA PADI INVESTAMA SEKURITAS Tbk": "MU",
            "PT. J.P. MORGAN SEKURITAS INDONESIA": "BK",
        }
        for name, expected in examples.items():
            with self.subTest(name=name):
                self.assertEqual(_broker_code(name), expected)

    def test_unmapped_names_and_custodians_keep_source_text(self):
        names = [
            "BANK MANDIRI, PT - CUSTODY",
            "PT BANK HSBC INDONESIA",
            "PT. BANK PERMATA, TBK",
            "CITIBANK, N. A",
            "BUT DEUTSCHE BANK AG",
            "PT CORPUS SEKURITAS INDONESIA",
            "PT Contoh Pemegang Rekening Baru",
            "PT ONIX SEKURITAS (REKENING TAMPUNGAN KSEI UNTUK CLOSED MEMBER-FM001)",
            "PT POOL ADVISTA SEKURITAS (TAMP QA001)",
            "PT MANDIRI SEKURITAS INTERNATIONAL",
        ]
        for name in names:
            with self.subTest(name=name):
                self.assertEqual(_broker_code(name), name)

    def test_requested_bank_account_aliases(self):
        examples = {
            "PT BANK DBS INDONESIA": "DP",
            "pt. bank dbs indonesia": "DP",
            "BANK CENTRAL ASIA Tbk": "SQ",
            "BANK CENTRAL ASIA Tbk, PT": "SQ",
            "bank rakyat indonesia (persero) pt": "OD",
            "PT. BANK RAKYAT INDONESIA (PERSERO) Tbk": "OD",
        }
        for name, expected in examples.items():
            with self.subTest(name=name):
                self.assertEqual(_broker_code(name), expected)

    def test_custom_alias_overrides_all_equivalent_default_spellings(self):
        overrides = {"PT ARTHA SEKURITAS INDONESIA": "ZZ"}
        for mapping in (overrides, _merged_broker_code_map(overrides)):
            self.assertEqual(_broker_code("Artha Sekuritas Indonesia, PT", mapping), "ZZ")
            self.assertEqual(_broker_code("PT. KGI Sekuritas Indonesia", mapping), "HD")

    def test_excel_uses_account_holder_column_for_fallback(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "2026-09-03.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["No", "Kode Efek"])
            sheet.append([
                1, "TEST", "Emiten Uji", "Bank Kustodian Uji, PT",
                "Investor Uji", "Nama Rekening Berbeda", "", "", "", "ID", "L",
                100, 100, "10", 150, 150, "15", 50,
            ])
            workbook.save(path)
            workbook.close()
            parsed = _parse_excel_for_code(path, "TEST")
        self.assertEqual(len(parsed), 1)
        data = parsed[0].data
        self.assertEqual(data["NAMA_PEMEGANG_REKENING_EFEK"], "Bank Kustodian Uji, PT")
        self.assertEqual(data["NAMA_PEMEGANG_REKENING_EFEK_ASLI"], "Bank Kustodian Uji, PT")
        self.assertEqual(data["NAMA_REKENING_EFEK"], "Nama Rekening Berbeda")

    def test_unmapped_accounts_remain_distinct_in_ownership_changes(self):
        def row(name, holding, previous):
            return FivePercentRow("2026-09-03.xlsx", "2026-09-03", {
                "KODE_EFEK": "TEST",
                "NAMA_PEMEGANG_SAHAM": "Investor Uji",
                "NAMA_PEMEGANG_REKENING_EFEK": _broker_code(name),
                "KEPEMILIKAN": holding,
                "PREVIOUS_HOLDING": previous,
                "PREVIOUS_REPORT_DATE": "2026-09-02",
                "REPORT_DATE": "2026-09-03",
            })

        rows = [row("Bank Kustodian A, PT", 0, 100),
                row("Bank Kustodian B, PT", 100, 0)]
        self.assertEqual(_combine_unavailable_broker_rows(rows), rows)
        changes = _build_latest_changes({"2026-09-03.xlsx": rows}, "TEST", {})
        by_account = {item["NAMA_PEMEGANG_REKENING_EFEK"]: item for item in changes}
        self.assertEqual(set(by_account), {"Bank Kustodian A, PT", "Bank Kustodian B, PT"})
        self.assertEqual(by_account["Bank Kustodian A, PT"]["CHANGE_SHARES"], -100)
        self.assertEqual(by_account["Bank Kustodian B, PT"]["CHANGE_SHARES"], 100)


if __name__ == "__main__":
    unittest.main()
