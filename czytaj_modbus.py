import asyncio
import argparse
import re
import shutil
import subprocess
import time
from datetime import datetime
from pathlib import Path

try:
    from rich.console import Console, Group
    from rich.table import Table
    from rich.text import Text
    from textual.app import App, ComposeResult
    from textual.widgets import DataTable, Footer, Header, Static
except ImportError as error:
    raise SystemExit(
        "Brak wymaganej biblioteki Rich lub Textual. Zainstaluj je poleceniem: "
        "python3 -m pip install rich textual"
    ) from error


def load_register_info(path):
    try:
        from openpyxl import load_workbook
    except ImportError as error:
        raise SystemExit(
            "Brak biblioteki openpyxl. Zainstaluj ja poleceniem: "
            "python -m pip install openpyxl"
        ) from error

    workbook_path = Path(path)
    if not workbook_path.is_file():
        raise SystemExit(f"Nie znaleziono pliku z mapa zmiennych: {workbook_path}")

    workbook = load_workbook(workbook_path, data_only=True, read_only=True)
    try:
        sheet = workbook["Status_var"]
    except KeyError as error:
        workbook.close()
        raise SystemExit(
            f"Arkusz 'Status_var' nie istnieje w pliku {workbook_path}"
        ) from error

    rows = sheet.iter_rows(values_only=True)
    headers = next(rows, None)
    if headers is None:
        workbook.close()
        raise SystemExit(f"Arkusz 'Status_var' w {workbook_path} jest pusty.")

    columns = {str(name).strip(): index for index, name in enumerate(headers) if name}
    required = {"Address", "Name", "IEC type"}
    missing = required - columns.keys()
    if missing:
        workbook.close()
        raise SystemExit(
            "Brak wymaganych kolumn w arkuszu Status_var: "
            + ", ".join(sorted(missing))
        )

    metadata_columns = {
        "name": "Name",
        "description": "Description",
        "iec_type": "IEC type",
        "min": "Min",
        "max": "Max",
        "scale": "Scale",
        "offset": "Offset",
        "unit": "Unit",
    }

    def cell_text(row, column_name):
        index = columns.get(column_name)
        if index is None or index >= len(row) or row[index] is None:
            return "-"
        return str(row[index]).strip() or "-"

    register_info = {}
    for row_number, row in enumerate(rows, start=2):
        if not row or row[columns["Address"]] is None:
            continue
        try:
            address = int(row[columns["Address"]])
        except (TypeError, ValueError) as error:
            workbook.close()
            raise SystemExit(
                f"Nieprawidlowy adres w arkuszu Status_var, wiersz {row_number}."
            ) from error

        register_info[address] = {
            key: cell_text(row, column_name)
            for key, column_name in metadata_columns.items()
        }

    workbook.close()
    if not register_info:
        raise SystemExit(f"Nie znaleziono adresow w arkuszu Status_var: {workbook_path}")
    return register_info


def parse_register_values(output):
    return {
        int(address): value.strip()
        for address, value in re.findall(
            r"^\s*\[(\d+)\]\s*:\s*(.*?)\s*$", output, re.MULTILINE
        )
    }


def contiguous_ranges(addresses):
    sorted_addresses = sorted(set(addresses))
    if not sorted_addresses:
        return

    range_start = previous = sorted_addresses[0]
    for address in sorted_addresses[1:]:
        if address != previous + 1:
            yield range_start, previous - range_start + 1
            range_start = address
        previous = address
    yield range_start, previous - range_start + 1


def read_register_range(command, start, count, timeout):
    range_command = command.copy()
    range_command[range_command.index("-r") + 1] = str(start)
    range_command[range_command.index("-c") + 1] = str(count)

    try:
        result = subprocess.run(
            range_command,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired:
        return {}, {
            address: "Timeout"
            for address in range(start, start + count)
        }

    output = result.stdout + result.stderr
    values = parse_register_values(output)
    if result.returncode == 0 and values:
        return values, {}

    if "Illegal Data Address" in output and count > 1:
        left_count = count // 2
        left_values, left_errors = read_register_range(
            command, start, left_count, timeout
        )
        right_values, right_errors = read_register_range(
            command, start + left_count, count - left_count, timeout
        )
        return (
            left_values | right_values,
            left_errors | right_errors,
        )

    error = output.strip() or f"modpoll exit code {result.returncode}"
    return {}, {
        address: error
        for address in range(start, start + count)
        if address not in values
    }


def read_configured_registers(
    command, register_info, start, count, timeout, on_update=None
):
    last_address = start + count - 1
    addresses = []
    for address, info in register_info.items():
        if start <= address <= last_address:
            addresses.append(address)
            if info["iec_type"] == "UDINT" and address < last_address:
                addresses.append(address + 1)

    values = {}
    errors = {}
    for block_start, block_count in contiguous_ranges(addresses):
        block_values, block_errors = read_register_range(
            command, block_start, block_count, timeout
        )
        values.update(block_values)
        errors.update(block_errors)
        if on_update is not None:
            on_update(block_values, block_errors)
    return values, errors


def build_dashboard(start, count, values, register_info, errors=None):
    errors = errors or {}
    table = Table(title="Odczyt rejestrow Modbus RTU", expand=True)
    for column in (
        "Adres",
        "Name",
        "Description",
        "IEC type",
        "Min",
        "Max",
        "Scale",
        "Offset",
        "Unit",
        "Wartosc",
    ):
        table.add_column(column, no_wrap=column in {"Adres", "IEC type"})

    failed_reads = 0
    for address in range(start, start + count):
        info = register_info.get(address)
        if info is None:
            continue

        value = values.get(address, "Brak odczytu")
        if address in errors:
            value = f"BLAD: {errors[address].splitlines()[-1]}"
            value = value.replace("\n", " ")
            failed_reads += 1
        if info["iec_type"] == "BOOL" and value in ("0", "1"):
            value = f"{value} ({'FALSE' if value == '0' else 'TRUE'})"

        if value.endswith("(TRUE)"):
            value_style = "bold green"
        elif value.endswith("(FALSE)"):
            value_style = "bold red"
        elif value.startswith("BLAD:"):
            value_style = "bold white on red"
        elif value == "Brak odczytu":
            value_style = "dim"
        elif value.lstrip("-").isdigit():
            value_style = "bold cyan"
        else:
            value_style = ""

        table.add_row(
            str(address),
            info["name"],
            info["description"],
            info["iec_type"],
            info["min"],
            info["max"],
            info["scale"],
            info["offset"],
            info["unit"],
            value,
            style=value_style,
        )

    status = Text(
        f"{datetime.now():%Y-%m-%d %H:%M:%S} | "
        f"zakres {start}-{start + count - 1} | "
        f"odczyty bledne: {failed_reads}"
    )
    return Group(status, table)


def styled_value(value, iec_type):
    if iec_type == "BOOL" and value in ("0", "1"):
        state = "TRUE" if value == "1" else "FALSE"
        style = "bold green" if value == "1" else "bold red"
        return Text(f"{value} ({state})", style=style)
    if value.startswith("BLAD:"):
        return Text(value, style="bold white on red")
    if value == "Brak odczytu":
        return Text(value, style="dim")
    if value.lstrip("-").isdigit():
        return Text(value, style="bold cyan")
    return Text(value)


class RegisterMonitor(App):
    CSS = """
    #status {
        height: 2;
        padding: 0 1;
    }
    DataTable {
        height: 1fr;
    }
    """
    BINDINGS = [
        ("q", "quit", "Koniec"),
        ("home", "scroll_top", "Poczatek tabeli"),
        ("end", "scroll_bottom", "Koniec tabeli"),
    ]

    def __init__(
        self,
        command,
        register_info,
        start,
        count,
        timeout,
        interval,
        no_color=False,
    ):
        super().__init__(ansi_color=not no_color)
        self.command = command
        self.register_info = register_info
        self.start = start
        self.count = count
        self.timeout = timeout
        self.interval = interval
        self.value_column = None

    def compose(self) -> ComposeResult:
        yield Header(show_clock=False)
        yield Static("Ladowanie odczytow...", id="status")
        yield DataTable(id="registers", zebra_stripes=True)
        yield Footer()

    def on_mount(self):
        table = self.query_one("#registers", DataTable)
        for label, key in (
            ("Adres", "address"),
            ("Name", "name"),
            ("Description", "description"),
            ("IEC type", "iec_type"),
            ("Min", "min"),
            ("Max", "max"),
            ("Scale", "scale"),
            ("Offset", "offset"),
            ("Unit", "unit"),
            ("Wartosc", "value"),
        ):
            column_key = table.add_column(label, key=key)
            if key == "value":
                self.value_column = column_key

        for address, info in sorted(self.register_info.items()):
            if not self.start <= address < self.start + self.count:
                continue
            table.add_row(
                str(address),
                info["name"],
                info["description"],
                info["iec_type"],
                info["min"],
                info["max"],
                info["scale"],
                info["offset"],
                info["unit"],
                Text("Oczekiwanie...", style="dim"),
                key=str(address),
            )

        self.run_worker(self.poll_loop(), name="modbus-poll", exclusive=True)

    async def poll_loop(self):
        while True:
            values, errors = await asyncio.to_thread(
                read_configured_registers,
                self.command,
                self.register_info,
                self.start,
                self.count,
                self.timeout,
            )

            table = self.query_one("#registers", DataTable)
            for address in sorted(self.register_info):
                if not self.start <= address < self.start + self.count:
                    continue
                info = self.register_info[address]
                value = values.get(address, "Brak odczytu")
                if address in errors:
                    value = f"BLAD: {errors[address].splitlines()[-1]}"
                table.update_cell(
                    str(address),
                    self.value_column,
                    styled_value(value, info["iec_type"]),
                )

            self.query_one("#status", Static).update(
                f"{datetime.now():%Y-%m-%d %H:%M:%S} | "
                f"{len(self.register_info)} zmiennych | "
                f"bledne odczyty: {len(errors)} | "
                "Przewijaj tabela; q = wyjscie"
            )
            await asyncio.sleep(self.interval)

    def action_scroll_top(self):
        self.query_one("#registers", DataTable).scroll_home(animate=False)

    def action_scroll_bottom(self):
        self.query_one("#registers", DataTable).scroll_end(animate=False)


def main():
    parser = argparse.ArgumentParser(
        description="Cykliczny odczyt rejestrow Modbus RTU przez modpoll"
    )
    parser.add_argument("--port", default="/dev/ttyUSB0", help="Port, np. /dev/ttyUSB0 lub COM3")
    parser.add_argument("--slave", type=int, default=1, help="Adres urzadzenia Modbus")
    parser.add_argument("--baud", type=int, default=38400, help="Predkosc transmisji")
    parser.add_argument("--start", type=int, default=8960, help="Pierwszy adres rejestru")
    parser.add_argument("--count", type=int, default=103, help="Liczba rejestrow")
    parser.add_argument("--interval", type=float, default=1.0, help="Przerwa miedzy odczytami w sekundach")
    parser.add_argument("--timeout", type=float, default=3.0, help="Limit czasu pojedynczego odczytu w sekundach")
    parser.add_argument("--modpoll", default="modpoll", help="Sciezka do programu modpoll")
    parser.add_argument(
        "--no-clear",
        action="store_true",
        help="Nie czysc terminala; dopisuj kolejne odczyty (przydatne do logowania)",
    )
    parser.add_argument(
        "--no-color",
        action="store_true",
        help="Wylacz kolorowanie wartosci",
    )
    parser.add_argument(
        "--metadata",
        default=str(Path(__file__).parent / "export" / "status_var.xlsx"),
        help="Plik XLSX z arkuszem Status_var",
    )
    args = parser.parse_args()

    register_info = load_register_info(args.metadata)
    console = Console(no_color=args.no_color)
    if shutil.which(args.modpoll) is None:
        raise SystemExit(
            f"Nie znaleziono programu {args.modpoll!r}. "
            "Zainstaluj modpoll albo podaj jego sciezke przez --modpoll."
        )

    configured_addresses = [
        address
        for address in register_info
        if args.start <= address < args.start + args.count
    ]
    if not configured_addresses:
        raise SystemExit("Zakres nie zawiera zadnego adresu z arkusza Status_var.")

    command = [
        args.modpoll,
        "-1",
        "-m", "rtu",
        "-b", str(args.baud),
        "-p", "even",
        "-d", "8",
        "-s", "1",
        "-a", str(args.slave),
        "-r", str(args.start),
        "-c", str(args.count),
        args.port,
    ]

    try:
        if args.no_clear:
            while True:
                values, errors = read_configured_registers(
                    command,
                    register_info,
                    args.start,
                    args.count,
                    args.timeout,
                )
                console.print(
                    build_dashboard(
                        args.start,
                        args.count,
                        values,
                        register_info,
                        errors,
                    )
                )
                time.sleep(args.interval)
        else:
            RegisterMonitor(
                command=command,
                register_info=register_info,
                start=args.start,
                count=args.count,
                timeout=args.timeout,
                interval=args.interval,
                no_color=args.no_color,
            ).run()

    except KeyboardInterrupt:
        console.print("\nZatrzymano odczyt.")


if __name__ == "__main__":
    main()
