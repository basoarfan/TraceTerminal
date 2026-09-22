"""Rich terminal UI using a shared RGB palette on a black backdrop.

Gray headers frame white data, orange accents, and yellow navigation.
Bright green/red highlight changes; darker tones style brokers and errors.
Rich emits 24-bit colors on supported terminals.
"""

from datetime import datetime
from typing import Any

from rich import box
from rich.console import Console, Group
from rich.panel import Panel
from rich.prompt import Prompt
from rich.table import Table
from rich.text import Text
from rich.theme import Theme

from traceterminal.analysis.shareholder5 import BROKER_CODE_MAP

# Shared RGB palette for all terminal screens.
GREEN = "#51C77A"         # accumulation / success
GREEN_NORMAL = "#23823E"  # green broker group
RED = "#C8463A"           # distribution / red broker group
RED_DARK = "#83160F"      # error badge background
ORANGE = "#D78C66"        # securities / accents / other brokers
YELLOW = "#D6D86B"        # navigation / warnings
BASE = "#E3E8E8"          # body text
HEADER = "#7A7E87"        # headers / borders / secondary text
BG = "#000000"           # backdrop

# Broker colors requested for the Stockbit-style account column. Other
# recognized broker codes use orange; unmapped account names keep their style.
RED_BROKERS = frozenset(
    "LS DR KZ BK AK ZP DP AI AG YP YU KK HD RX XA RB FS BQ CP DU TP".split()
)
GREEN_BROKERS = frozenset("CC NI OD DX".split())
KNOWN_BROKERS = frozenset(BROKER_CODE_MAP.values())
BROKER_KEYS = frozenset({"NAMA_PEMEGANG_REKENING_EFEK", "BROKER_CODE"})

THEME = Theme(
    {
        "amber": f"bold {YELLOW}",
        "accent": f"bold {ORANGE}",
        "success": f"bold {GREEN}",
        "up": f"bold {GREEN}",
        "down": f"bold {RED}",
        "danger": f"bold {BASE} on {RED_DARK}",
        "warning": f"bold {YELLOW}",
        "info": ORANGE,
        "muted": HEADER,
        "header": f"bold {HEADER}",
        "prompt": f"bold {YELLOW}",
        "prompt.choices": HEADER,
        "prompt.default": ORANGE,
        "prompt.invalid": f"bold {RED}",
        "prompt.invalid.choice": f"bold {RED}",
        "flat": BASE,
    }
)

# Keys whose numeric content should be right-aligned (Bloomberg reads
# numbers in tight columns from the right edge).
NUMERIC_RIGHT = frozenset(
    {
        "TOTAL_HOLDING_SHARES",
        "HOLDINGS_SCRIPLESS",
        "HOLDINGS_SCRIP",
        "PERCENTAGE",
        "PREVIOUS_TOTAL",
        "CURRENT_TOTAL",
        "PREVIOUS_HOLDING",
        "LATEST_HOLDING",
        "CHANGE_SHARES",
        "ESTIMATED_LOTS",
        "KEPEMILIKAN",
        "GOVERNMENT",
        "POLITICAL_PARTIES",
        "STATE_OWNED_COMPANY",
        "CORPORATE",
        "INDIVIDUAL",
        "TOTAL_SCRIPLESS",
        "PREVIOUS_VALUE",
        "LATEST_VALUE",
    }
)

# Keys where the sign of the number is meaningful and should be colored.
DELTA_KEYS = frozenset({"CHANGE_SHARES", "ESTIMATED_LOTS"})

# Semantic column groups: each group gets one readable color, Bloomberg-style.
SECURITY_KEYS = frozenset({"SHARE_CODE", "KODE_EFEK"})
DATE_KEYS = frozenset({"DATE", "REPORT_DATE", "FROM_DATE", "TO_DATE"})
NAME_KEYS = frozenset(
    {
        "ISSUER_NAME",
        "NAMA_EMITEN",
        "INVESTOR_NAME",
        "NAMA_PEMEGANG_SAHAM",
        "NAMA_PEMEGANG_REKENING_EFEK",
        "NAMA_REKENING_EFEK",
    }
)
FIELD_KEYS = frozenset(
    {
        "INVESTOR_CLASSIFICATION",
        "NATIONALITY",
        "DOMICILE",
        "DOMISILI",
        "LOCAL_FOREIGN",
        "KOLOM",
    }
)


class TerminalUI:
    """Display the menu, prompts, and shareholder analysis tables."""

    def __init__(self) -> None:
        self.console = Console(
            theme=THEME,
            style=f"{BASE} on {BG}",
        )

    # ------------------------------------------------------------------
    # Header / chrome
    # ------------------------------------------------------------------
    def show_banner(self) -> None:
        self.console.clear()
        now = datetime.now().strftime("%d %b %Y  %H:%M:%S")
        header = (
            "[accent]Analisis Data Kepemilikan Saham Indonesia di Atas 1% & 5%[/accent]\n"
            "[muted]▸ Github : [link=https://github.com/basoarfan]https://github.com/basoarfan[/link][/muted]\n"
            "[muted]▸ Sociabuzz : [link=https://sociabuzz.com/basoarfan/tribe]https://sociabuzz.com/basoarfan/tribe[/link][/muted]"
        )
        self.console.print(
            Panel(
                Text.from_markup(header),
                title="[header]Trace Terminal[/header]",
                subtitle=Text(now, style=HEADER),
                border_style=HEADER,
                box=box.DOUBLE,
                padding=(1, 2),
            )
        )
        self.console.print(
            "[muted]Pilih nomor menu lalu tekan <ENTER>. "
            "Tekan <CTRL+C> kapan saja untuk keluar.[/muted]\n"
        )

    def show_menu(self) -> str:
        items = [
            ("1", "Analisa Pemegang Saham 1%"),
            ("2", "Analisa Pemegang Saham >5%"),
            ("3", "Cari Klasifikasi Pemegang Saham"),
            ("4", "Scanner Pemegang Saham 1% (Seluruh Emiten)"),
            ("5", "Scanner Pemegang Saham 5% (Seluruh Emiten)"),
            ("6", "Scanner Nama Pemegang Saham"),
            ("0", "Keluar"),
        ]
        menu = Table(show_header=False, box=None, pad_edge=False, padding=(0, 2))
        menu.add_column(style=f"bold {YELLOW}", width=7, justify="center")
        menu.add_column(style=BASE)
        for key, action in items:
            menu.add_row(f"[{key}]", action)

        footer = (
            "\n[muted]NAVIGASI:[/muted] "
            "[amber]\\[1]-\\[6][/amber] jalankan modul      "
            "[amber]\\[0][/amber] keluar      "
            "[muted]<CTRL+C>[/muted] batalkan"
        )
        self.console.print(
            Panel(
                Group(menu, Text.from_markup(footer)),
                title="[header]MENU UTAMA[/header]",
                border_style=HEADER,
                box=box.SQUARE,
                padding=(1, 2),
            )
        )
        return Prompt.ask(
            console=self.console,
            prompt="\n[amber]PILIH MENU > [/amber]",
            choices=["0", "1", "2", "3", "4", "5", "6"],
            default="1",
        )

    # ------------------------------------------------------------------
    # Prompts
    # ------------------------------------------------------------------
    def ask_share_code(self) -> str:
        return Prompt.ask(
            console=self.console,
            prompt="[accent]KODE EMITEN / SHARE_CODE > [/accent]"
        ).strip().upper()

    def ask_security_code(self) -> str:
        return Prompt.ask(
            console=self.console,
            prompt="[accent]KODE EFEK > [/accent]"
        ).strip().upper()

    def ask_investor_name(self) -> str:
        return Prompt.ask(
            console=self.console,
            prompt="[amber]NAMA PEMEGANG SAHAM > [/amber]"
        ).strip()

    def ask_shareholder_period(self) -> int:
        return int(
            Prompt.ask(
                console=self.console,
                prompt="[amber]PERIODE DATA (2/3/6/12 BULAN) > [/amber]",
                choices=["2", "3", "6", "12"],
                default="2",
            )
        )

    def ask_five_percent_scanner_period(self) -> int:
        return int(
            Prompt.ask(
                console=self.console,
                prompt="[amber]PERIODE ([1] 2 HARI; [2/3/6/12] BULAN) > [/amber]",
                choices=["1", "2", "3", "6", "12"],
                default="1",
            )
        )

    # ------------------------------------------------------------------
    # Status / feedback lines
    # ------------------------------------------------------------------
    def show_progress(self, message: str) -> None:
        self.console.print(f"  [amber]»[/amber] {message}")

    def show_error(self, message: str) -> None:
        self.console.print(f"\n[danger] ✗ ERROR: [/danger] {message}")

    def show_success(self, message: str) -> None:
        self.console.print(f"\n[success]✓ {message}[/success]")

    def press_enter(self) -> None:
        self.console.input(
            "\n[muted]Tekan <ENTER> untuk kembali ke menu...[/muted]"
        )

    # ------------------------------------------------------------------
    # Report screens
    # ------------------------------------------------------------------
    def show_shareholder_analysis(self, analysis: dict[str, Any]) -> None:
        share_code = analysis.get("share_code", "UNKNOWN")
        period = analysis.get("period_months", "ALL")
        rows = analysis.get("rows", [])

        table = self._table(f"DATA KEPEMILIKAN SAHAM — {share_code} ({period} BULAN)")
        columns = [
            ("Tanggal", "DATE"),
            ("Kode", "SHARE_CODE"),
            ("Emiten", "ISSUER_NAME"),
            ("Investor", "INVESTOR_NAME"),
            ("Klasifikasi", "INVESTOR_CLASSIFICATION"),
            ("L/F", "LOCAL_FOREIGN"),
            ("Kebangsaan", "NATIONALITY"),
            ("Domisili", "DOMICILE"),
            ("Scripless", "HOLDINGS_SCRIPLESS"),
            ("Scrip", "HOLDINGS_SCRIP"),
            ("Total Saham", "TOTAL_HOLDING_SHARES"),
            ("Persentase", "PERCENTAGE"),
        ]
        self._add_data(table, columns, rows)
        self.console.print(table)

        changes = analysis.get("changes", [])
        change_table = self._table("ANALISA AKUMULASI / DISTRIBUSI", HEADER)
        change_columns = [
            ("Investor", "INVESTOR_NAME"),
            ("Dari", "FROM_DATE"),
            ("Ke", "TO_DATE"),
            ("Saham Sebelumnya", "PREVIOUS_TOTAL"),
            ("Saham Terbaru", "CURRENT_TOTAL"),
            ("Perubahan Saham", "CHANGE_SHARES"),
            ("Estimasi Lot", "ESTIMATED_LOTS"),
            ("Status", "STATUS"),
        ]
        self._add_data(change_table, change_columns, changes)
        self.console.print(change_table)
        self._show_notes(analysis.get("notes", []))

    def show_shareholder_scanner(self, analysis: dict[str, Any]) -> None:
        period = analysis.get("period_months", "ALL")
        changes = analysis.get("changes", [])
        total_issuers = analysis.get("total_issuers_scanned", 0)

        table = self._table(
            f"SCANNER PEMEGANG SAHAM 1% — {period} BULAN "
            f"({total_issuers} EMITEN DIPINDAI)",
            HEADER,
        )
        columns = [
            ("Emiten", "SHARE_CODE"),
            ("Investor", "INVESTOR_NAME"),
            ("Dari", "FROM_DATE"),
            ("Ke", "TO_DATE"),
            ("Saham Sebelumnya", "PREVIOUS_TOTAL"),
            ("Saham Terbaru", "CURRENT_TOTAL"),
            ("Estimasi Lot", "ESTIMATED_LOTS"),
            ("Status", "STATUS"),
        ]
        self._add_data(table, columns, changes)
        self.console.print(table)
        self._show_notes(analysis.get("notes", []))

    def show_five_percent_analysis(self, analysis: dict[str, Any]) -> None:
        security_code = analysis.get("security_code", "UNKNOWN")
        period = analysis.get("period_months") or "ALL"
        report_date = analysis.get("latest_report_date", "Data tidak tersedia")
        rows = analysis.get("rows", [])

        if period != "ALL":
            title = (
                f"DATA PEMEGANG SAHAM >5% — {security_code} "
                f"({period} BULAN TERAKHIR, LAPORAN S.D. {report_date})"
            )
        else:
            title = f"DATA PEMEGANG SAHAM >5% — {security_code} ({report_date})"
        table = self._table(title)
        columns = [
            ("Emiten", "NAMA_EMITEN"),
            ("Pemegang Rekening Efek", "NAMA_PEMEGANG_REKENING_EFEK"),
            ("Pemegang Saham", "NAMA_PEMEGANG_SAHAM"),
            ("Rekening Efek", "NAMA_REKENING_EFEK"),
            ("Domisili", "DOMISILI"),
            ("Jumlah Saham", "KEPEMILIKAN"),
        ]
        self._add_data(table, columns, rows)
        self.console.print(table)

        changes = analysis.get("changes", [])
        change_table = self._table("ANALISA AKUMULASI / DISTRIBUSI >5%", HEADER)
        change_columns = [
            ("Pemegang Saham", "NAMA_PEMEGANG_SAHAM"),
            ("Rekening Efek", "NAMA_PEMEGANG_REKENING_EFEK"),
            ("Dari", "FROM_DATE"),
            ("Ke", "TO_DATE"),
            ("Saham Sebelumnya", "PREVIOUS_HOLDING"),
            ("Saham Terbaru", "LATEST_HOLDING"),
            ("Perubahan Saham", "CHANGE_SHARES"),
            ("Estimasi Lot", "ESTIMATED_LOTS"),
            ("Status", "STATUS"),
        ]
        self._add_data(change_table, change_columns, changes)
        self.console.print(change_table)
        self._show_notes(analysis.get("notes", []))

    def show_five_percent_scanner(self, analysis: dict[str, Any]) -> None:
        period = analysis.get("period_months", "ALL")
        changes = analysis.get("changes", [])
        total_issuers = analysis.get("total_issuers_scanned", 0)
        period_label = "2 HARI" if period == 1 else f"{period} BULAN"

        table = self._table(
            f"SCANNER PEMEGANG SAHAM 5% — {period_label} "
            f"({total_issuers} EMITEN DIPINDAI)",
            HEADER,
        )
        columns = [
            ("Emiten", "KODE_EFEK"),
            ("Pemegang Saham", "NAMA_PEMEGANG_SAHAM"),
            ("Rekening Efek", "NAMA_PEMEGANG_REKENING_EFEK"),
            ("Dari", "FROM_DATE"),
            ("Ke", "TO_DATE"),
            ("Saham Sebelumnya", "PREVIOUS_HOLDING"),
            ("Saham Terbaru", "LATEST_HOLDING"),
            ("Estimasi Lot", "ESTIMATED_LOTS"),
            ("Status", "STATUS"),
        ]
        self._add_data(table, columns, changes)
        self.console.print(table)
        self._show_notes(analysis.get("notes", []))

    def show_investor_name_scanner(self, analysis: dict[str, Any]) -> None:
        investor_query = analysis.get("investor_query", "UNKNOWN")
        source = analysis.get("source", "1%")
        period = analysis.get("period_months", "ALL")
        changes = analysis.get("changes", [])
        total_issuers = analysis.get("total_issuers_matched", 0)

        table = self._table(
            f"SCANNER NAMA PEMEGANG SAHAM {source} — {investor_query} — "
            f"{period} BULAN ({total_issuers} EMITEN)",
            HEADER,
        )
        columns = [
            ("Emiten", "SHARE_CODE"),
            ("Pemegang Saham", "INVESTOR_NAME"),
            ("Dari", "FROM_DATE"),
            ("Ke", "TO_DATE"),
            ("Saham Sebelumnya", "PREVIOUS_TOTAL"),
            ("Saham Terbaru", "CURRENT_TOTAL"),
            ("Estimasi Lot", "ESTIMATED_LOTS"),
            ("Status", "STATUS"),
        ]
        self._add_data(table, columns, changes)
        self.console.print(table)
        self._show_notes(analysis.get("notes", []))

    def show_shareholder_classification_analysis(self, analysis: dict[str, Any]) -> None:
        share_code = analysis.get("share_code", "UNKNOWN")
        period = analysis.get("period_months", "ALL")
        rows = analysis.get("rows", [])

        table = self._table(
            f"KLASIFIKASI PEMEGANG SAHAM — {share_code} ({period} BULAN)"
        )
        columns = [
            ("Tanggal", "REPORT_DATE"),
            ("SHARE CODE", "SHARE_CODE"),
            ("Emiten", "ISSUER_NAME"),
            ("Government", "GOVERNMENT"),
            ("Political Parties", "POLITICAL_PARTIES"),
            ("State Owned Company", "STATE_OWNED_COMPANY"),
            ("Corporate", "CORPORATE"),
            ("Individual", "INDIVIDUAL"),
            ("Total Scripless", "TOTAL_SCRIPLESS"),
        ]
        self._add_data(table, columns, rows)
        self.console.print(table)

        changes = analysis.get("changes", [])
        change_table = self._table(
            "PERUBAHAN KLASIFIKASI (TERAWAL vs TERBARU)",
            HEADER,
        )
        change_columns = [
            ("Kolom", "KOLOM"),
            ("Dari", "FROM_DATE"),
            ("Ke", "TO_DATE"),
            ("Sebelumnya", "PREVIOUS_VALUE"),
            ("Terbaru", "LATEST_VALUE"),
            ("Perubahan Saham", "CHANGE_SHARES"),
            ("Estimasi Lot", "ESTIMATED_LOTS"),
            ("Status", "STATUS"),
        ]
        self._add_data(change_table, change_columns, changes)
        self.console.print(change_table)
        self._show_notes(analysis.get("notes", []))

    # ------------------------------------------------------------------
    # Table building helpers
    # ------------------------------------------------------------------
    def _table(self, title: str, color: str = HEADER) -> Table:
        return Table(
            title=title,
            title_style=f"bold {ORANGE}",
            header_style=f"bold {HEADER}",
            title_justify="left",
            border_style=color,
            box=box.SQUARE,
            show_lines=True,
            pad_edge=False,
            padding=(0, 1),
        )

    def _add_data(
        self,
        table: Table,
        columns: list[tuple[str, str]],
        rows: list[dict[str, Any]],
    ) -> None:
        for title, key in columns:
            table.add_column(
                title.upper(),
                justify="right" if key in NUMERIC_RIGHT else "left",
                overflow="fold",
                style=self._column_style(key),
            )
        if not rows:
            placeholder = Text("Data tidak tersedia", style=f"italic {HEADER}")
            table.add_row(*([placeholder] * len(columns)))
            return
        for row in rows:
            table.add_row(*[self._value(row.get(key), key) for _, key in columns])

    def _column_style(self, key: str) -> str | None:
        """Color a column by its semantic group (Bloomberg-ish palette)."""
        if key in SECURITY_KEYS:
            return f"bold {ORANGE}"   # tickers / security codes
        if key in DATE_KEYS:
            return HEADER              # dates & periods
        if key in NAME_KEYS:
            return f"bold {BASE}"     # issuer / holder names
        if key in FIELD_KEYS:
            return ORANGE             # classification, nationality, domicile
        return None

    def _value(self, value: Any, key: str = "") -> str | Text:
        if value is None or value == "":
            return "[muted]Data tidak tersedia[/muted]"
        if isinstance(value, bool):
            return "Ya" if value else "Tidak"
        if key in BROKER_KEYS:
            code = str(value).strip().upper()
            if code in KNOWN_BROKERS:
                color = (
                    RED if code in RED_BROKERS else
                    GREEN_NORMAL if code in GREEN_BROKERS else ORANGE
                )
                return Text(str(value), style=f"bold {color}")
        if key == "STATUS":
            return self._status(str(value))
        if isinstance(value, int):
            return self._signed(f"{value:,}", value, key)
        if isinstance(value, float):
            decimals = 2 if key == "ESTIMATED_LOTS" else 4
            return self._signed(f"{value:,.{decimals}f}", value, key)
        return str(value)

    def _signed(self, formatted: str, value: float, key: str) -> str:
        if key not in DELTA_KEYS:
            return formatted
        if value > 0:
            return f"[up]+{formatted}[/up]"
        if value < 0:
            return f"[down]{formatted}[/down]"
        return formatted

    def _status(self, text: str) -> str:
        lowered = text.lower()
        if "akumulasi" in lowered or "masuk" in lowered:
            return f"[up]▲ {text}[/up]"
        if "distribusi" in lowered or "keluar" in lowered:
            return f"[down]▼ {text}[/down]"
        if "netral" in lowered:
            return f"[flat]{text}[/flat]"
        return f"[warning]{text}[/warning]"

    def _show_notes(self, notes: list[str]) -> None:
        if not notes:
            return
        content = Text()
        for note in notes:
            content.append("- ", style=f"bold {YELLOW}")
            content.append(note + "\n")
        self.console.print(
            Panel(
                content,
                title="[header]CATATAN TRANSPARANSI[/header]",
                border_style=HEADER,
                box=box.SQUARE,
                padding=(1, 2),
            )
        )
