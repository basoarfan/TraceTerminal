import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from openpyxl import Workbook

from traceterminal.analysis.shareholder import scan_shareholder_by_investor
from traceterminal.analysis.shareholder5 import scan_five_percent_documents
from traceterminal.app import TraceTerminalApp


class InvestorScannerTests(unittest.TestCase):
    def test_five_percent_name_search_includes_neutral_and_filters_other_names(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "2026-09-03.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["No", "Kode Efek", "Kepemilikan Per 2026-09-02",
                          "Kepemilikan Per 2026-09-03"])
            for number, name, previous, current in (
                (1, "Investor Uji", 100, 150),
                (2, "Investor Uji Netral", 100, 100),
                (3, "Pemegang Lain", 100, 200),
            ):
                sheet.append([
                    number, "TEST", "Emiten Uji", "Bank Kustodian Uji, PT",
                    name, name, "", "", "", "ID", "L",
                    previous, previous, "10", current, current, "15",
                    current - previous,
                ])
            workbook.save(path)
            workbook.close()

            result = scan_shareholder_by_investor(
                Path(directory), "investor uji", 2, source="5%",
            )
            default = scan_five_percent_documents(Path(directory), 2)

        self.assertEqual(result["total_changes"], 2)
        self.assertEqual(result["total_issuers_matched"], 1)
        self.assertEqual({row["STATUS"] for row in result["changes"]},
                         {"Akumulasi", "Netral"})
        self.assertEqual(result["changes"][0]["CURRENT_TOTAL"], 150)
        self.assertTrue(all(row["CHANGE_SHARES"] != 0 for row in default["changes"]))

    def test_menu_six_scans_both_sources_even_when_one_percent_is_empty_or_missing(self):
        for first_result in ({"changes": []}, FileNotFoundError("missing")):
            with self.subTest(first_result=first_result):
                app = TraceTerminalApp.__new__(TraceTerminalApp)
                app.ui = Mock()
                app.ui.ask_investor_name.return_value = "Investor Uji"
                app.ui.ask_shareholder_period.return_value = 2
                app._record = Mock()
                with patch("traceterminal.app.scan_shareholder_by_investor",
                           side_effect=[first_result, {"changes": []}]) as scan:
                    app._handle_choice("6")

                self.assertEqual([call.kwargs["source"] for call in scan.call_args_list],
                                 ["1%", "5%"])
                self.assertTrue(all(call.args[1] == "Investor Uji"
                                    for call in scan.call_args_list))
                app.ui.press_enter.assert_not_called()
                entry = app._record.call_args.args[0]
                self.assertIn("5%", entry["analysis"])


if __name__ == "__main__":
    unittest.main()
