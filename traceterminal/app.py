"""Main application for shareholder document analysis."""

import sys
from pathlib import Path

from traceterminal.analysis.shareholder import (
    analyze_shareholder_documents,
    scan_shareholder_by_investor,
    scan_shareholder_documents,
)
from traceterminal.analysis.shareholder5 import (
    analyze_five_percent_documents,
    scan_five_percent_documents,
)
from traceterminal.analysis.shareholder_name import analyze_shareholder_name_documents
from traceterminal.logger.data_logger import DataLogger
from traceterminal.ui.terminal import TerminalUI


class TraceTerminalApp:
    """Interactive terminal for the shareholder analysis workflows."""

    def __init__(self) -> None:
        self.ui = TerminalUI()
        self.logger = DataLogger(Path.cwd() / "output" / "logs")
        self.session_entries: list[dict] = []
        self.running = True

    def run(self) -> None:
        """Run the main menu until the user exits."""
        self.ui.show_banner()

        while self.running:
            try:
                self._handle_choice(self.ui.show_menu())
            except KeyboardInterrupt:
                self._shutdown()
            except Exception as exc:
                self.ui.show_error(str(exc))
                self.ui.press_enter()

    def _handle_choice(self, choice: str) -> None:
        actions = {
            "1": self._shareholder_ownership_analyze,
            "2": self._five_percent_ownership_analyze,
            "3": self._owner_name_ownership_analyze,
            "4": self._shareholder_scanner,
            "5": self._five_percent_scanner,
            "6": self._investor_name_scanner,
            "0": self._shutdown,
        }
        actions[choice]()

    def _shareholder_ownership_analyze(self) -> None:
        share_code = self.ui.ask_share_code()
        period_months = self.ui.ask_shareholder_period()
        documents_dir = Path.cwd() / "documents" / "pemegang_saham_1%"

        self.ui.show_progress(
            f"Membaca Excel pemegang saham {period_months} bulan terakhir dari "
            f"{documents_dir}..."
        )
        try:
            analysis = analyze_shareholder_documents(
                documents_dir,
                share_code,
                period_months=period_months,
            )
        except (FileNotFoundError, ValueError, ImportError) as exc:
            self._show_analysis_error(exc)
            return

        self.ui.show_shareholder_analysis(analysis)
        self._record(
            {
                "source": "shareholder_ownership_excel",
                "symbol": share_code,
                "period_months": period_months,
                "analysis": analysis,
            },
            share_code,
        )

    def _five_percent_ownership_analyze(self) -> None:
        security_code = self.ui.ask_security_code()
        period_months = self.ui.ask_shareholder_period()
        documents_dir = Path.cwd() / "documents" / "pemegang_saham_5%"

        self.ui.show_progress(
            f"Membaca Excel pemegang saham >5% {period_months} bulan terakhir dari "
            f"{documents_dir}..."
        )
        try:
            analysis = analyze_five_percent_documents(
                documents_dir,
                security_code,
                period_months=period_months,
            )
        except (FileNotFoundError, ValueError, ImportError) as exc:
            self._show_analysis_error(exc)
            return

        self.ui.show_five_percent_analysis(analysis)
        self._record(
            {
                "source": "shareholder_5_percent_excel",
                "symbol": security_code,
                "period_months": period_months,
                "analysis": analysis,
            },
            security_code,
        )

    def _owner_name_ownership_analyze(self) -> None:
        share_code = self.ui.ask_share_code()
        period_months = self.ui.ask_shareholder_period()
        documents_dir = Path.cwd() / "documents" / "klasifikasi_pemegang_saham"

        self.ui.show_progress(
            f"Membaca Excel klasifikasi pemegang saham {period_months} bulan terakhir dari "
            f"{documents_dir}..."
        )
        try:
            analysis = analyze_shareholder_name_documents(
                documents_dir,
                share_code,
                period_months=period_months,
            )
        except (FileNotFoundError, ValueError, ImportError) as exc:
            self._show_analysis_error(exc)
            return

        self.ui.show_shareholder_classification_analysis(analysis)
        self._record(
            {
                "source": "shareholder_classification_excel",
                "share_code": share_code,
                "period_months": period_months,
                "analysis": analysis,
            },
            share_code,
        )

    def _shareholder_scanner(self) -> None:
        period_months = self.ui.ask_shareholder_period()
        documents_dir = Path.cwd() / "documents" / "pemegang_saham_1%"

        self.ui.show_progress(
            f"Memindai seluruh emiten pemegang saham 1% untuk "
            f"{period_months} bulan terakhir dari {documents_dir}..."
        )
        try:
            analysis = scan_shareholder_documents(
                documents_dir,
                period_months=period_months,
            )
        except (FileNotFoundError, ValueError, ImportError) as exc:
            self._show_analysis_error(exc)
            return

        self.ui.show_shareholder_scanner(analysis)
        self._record(
            {
                "source": "shareholder_1_percent_scanner",
                "period_months": period_months,
                "analysis": analysis,
            },
            "ALL",
        )

    def _five_percent_scanner(self) -> None:
        period_months = self.ui.ask_five_percent_scanner_period()
        documents_dir = Path.cwd() / "documents" / "pemegang_saham_5%"

        period_label = (
            "dua tanggal laporan terbaru"
            if period_months == 1
            else f"{period_months} bulan terakhir"
        )
        self.ui.show_progress(
            f"Memindai seluruh emiten pemegang saham >5% untuk "
            f"{period_label} dari {documents_dir}..."
        )
        try:
            analysis = scan_five_percent_documents(
                documents_dir,
                period_months=period_months,
            )
        except (FileNotFoundError, ValueError, ImportError) as exc:
            self._show_analysis_error(exc)
            return

        self.ui.show_five_percent_scanner(analysis)
        self._record(
            {
                "source": "shareholder_5_percent_scanner",
                "period_months": period_months,
                "analysis": analysis,
            },
            "ALL",
        )

    def _investor_name_scanner(self) -> None:
        investor_name = self.ui.ask_investor_name()
        period_months = self.ui.ask_shareholder_period()
        analyses = {}
        errors = {}
        for source in ("1%", "5%"):
            documents_dir = Path.cwd() / "documents" / f"pemegang_saham_{source}"
            self.ui.show_progress(
                f"Mencari pemegang saham '{investor_name}' pada dokumen {source} "
                f"untuk {period_months} bulan terakhir dari {documents_dir}..."
            )
            try:
                analysis = scan_shareholder_by_investor(
                    documents_dir,
                    investor_name,
                    period_months=period_months,
                    source=source,
                )
            except (FileNotFoundError, ValueError, ImportError) as exc:
                errors[source] = str(exc)
                self.ui.show_error(f"Dokumen {source}: {exc}")
                continue

            analyses[source] = analysis
            self.ui.show_investor_name_scanner(analysis)

        if not analyses:
            self.ui.press_enter()
            return

        self._record(
            {
                "source": "shareholder_name_scanner",
                "investor_name": investor_name,
                "period_months": period_months,
                "analysis": analyses,
                "errors": errors,
            },
            "ALL",
        )

    def _record(self, entry: dict, symbol: str) -> None:
        self.logger.log_analysis(entry, {"symbol": symbol})
        self.session_entries.append(entry)
        self.ui.press_enter()

    def _show_analysis_error(self, exc: Exception) -> None:
        self.ui.show_error(str(exc))
        self.ui.press_enter()

    def _shutdown(self) -> None:
        if not self.running:
            return
        if self.session_entries:
            session_log = self.logger.log_session(self.session_entries)
            self.ui.show_progress(f"Session log tersimpan: {session_log}")
        self.ui.show_success("Trace Terminal - Selesai.")
        self.running = False


def run() -> None:
    """Convenience entry point."""
    try:
        TraceTerminalApp().run()
    except KeyboardInterrupt:
        sys.exit(0)
