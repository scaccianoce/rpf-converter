from __future__ import annotations

import argparse
import csv
import re
import sys
import tkinter as tk
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog
from typing import Callable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

FIELD_WIDTH = 16
DEFAULT_RECORD_PREFIX = "RB"
SHEET_NAME = "Campi RPF"
MAX_FIELD_NUMBER = 20
DEFAULT_MAX_IMPORT_FIELD = 12
NON_IMPORTABLE_RECORDS = frozenset({10, 11})
NUMERIC_VALUE = re.compile(r"^[+-]?\d+(?:,\d+)?$")


class RpfError(ValueError):
    pass


def metadata_columns(record_prefix: str) -> tuple[str, str, str]:
    return ("Modulo", "Riga file", f"Record {record_prefix}")


def field_columns(record_prefix: str) -> tuple[str, ...]:
    return tuple(f"{record_prefix}{i:03d}" for i in range(1, MAX_FIELD_NUMBER + 1))


def read_rpf(path: Path) -> list[str]:
    try:
        with path.open("r", encoding="ascii", newline="") as file:
            return file.readlines()
    except UnicodeDecodeError as error:
        raise RpfError("Il file RPF deve essere codificato in ASCII.") from error


def validate_reader_settings(record_prefix: str, field_width: int = FIELD_WIDTH) -> tuple[str, int]:
    record_prefix = record_prefix.strip().upper()
    if not re.fullmatch(r"[A-Z]{2}", record_prefix):
        raise RpfError("Il prefisso del record deve contenere esattamente due lettere.")
    if field_width <= 0:
        raise RpfError("La lunghezza dei campi deve essere un numero positivo.")
    return record_prefix, field_width


C_PAYLOAD_START = 89
C_PAYLOAD_END = 1889
C_PAYLOAD_LENGTH = C_PAYLOAD_END - C_PAYLOAD_START
C_CHUNK_SIZE = 8 + FIELD_WIDTH
C_CHUNKS_PER_RECORD = C_PAYLOAD_LENGTH // C_CHUNK_SIZE
C_BODY_LENGTH = 1898
C_MODULE_START = 17
C_MODULE_END = 25
Z_C_COUNT_START = 24
Z_C_COUNT_END = 33
CODE_RE = re.compile(r"^[A-Z0-9]{8}$")


def _split_line_ending(line: str) -> tuple[str, str]:
    if line.endswith("\r\n"):
        return line[:-2], "\r\n"
    if line.endswith("\n"):
        return line[:-1], "\n"
    if line.endswith("\r"):
        return line[:-1], "\r"
    return line, ""


def parse_c_modules(lines: list[str]) -> dict[int, dict[str, object]]:
    """Legge i record C usando il Progressivo modulo ufficiale (pos. 18-25)."""
    modules: dict[int, dict[str, object]] = {}
    for line_index, line in enumerate(lines):
        body, ending = _split_line_ending(line)
        if not body.startswith("C"):
            continue
        if len(body) != C_BODY_LENGTH:
            raise RpfError(
                f"Il record C alla riga {line_index + 1} e' lungo {len(body)} caratteri; "
                f"ne sono attesi {C_BODY_LENGTH} prima di CR/LF."
            )
        module_text = body[C_MODULE_START:C_MODULE_END]
        if not module_text.isdigit() or int(module_text) <= 0:
            raise RpfError(
                f"Progressivo modulo non valido nel record C alla riga {line_index + 1}: "
                f"'{module_text}'."
            )
        module = int(module_text)
        info = modules.setdefault(
            module,
            {"line_indexes": [], "templates": [], "chunks": []},
        )
        info["line_indexes"].append(line_index)
        info["templates"].append((body, ending))

        payload = body[C_PAYLOAD_START:C_PAYLOAD_END]
        chunks = info["chunks"]
        for offset in range(0, C_PAYLOAD_LENGTH, C_CHUNK_SIZE):
            code = payload[offset : offset + 8]
            value = payload[offset + 8 : offset + C_CHUNK_SIZE]
            if not code.strip():
                if value.strip():
                    raise RpfError(
                        f"Valore senza codice nel record C alla riga {line_index + 1}, "
                        f"slot {offset // C_CHUNK_SIZE + 1}."
                    )
                continue
            if not CODE_RE.fullmatch(code):
                raise RpfError(
                    f"Codice non valido '{code}' nel record C alla riga {line_index + 1}."
                )
            chunks.append({"code": code, "value": value, "source_line": line_index + 1})

    if not modules:
        raise RpfError("Nel file RPF non sono presenti record di tipo C.")
    return modules


def find_rpf_fields(
    lines: list[str],
    record_prefix: str = DEFAULT_RECORD_PREFIX,
    field_width: int = FIELD_WIDTH,
) -> list[dict[str, object]]:
    """Trova i campi del quadro usando il Progressivo modulo dei record C."""
    record_prefix, field_width = validate_reader_settings(record_prefix, field_width)
    if field_width != FIELD_WIDTH:
        raise RpfError("Per il formato RPF dei record C la lunghezza campo deve essere 16.")
    modules = parse_c_modules(lines)
    pattern = re.compile(
        rf"^{re.escape(record_prefix)}(?P<record>\d{{3}})(?P<field>\d{{3}})$"
    )
    fields: list[dict[str, object]] = []
    for module in sorted(modules):
        for chunk in modules[module]["chunks"]:
            match = pattern.fullmatch(chunk["code"])
            if not match:
                continue
            fields.append(
                {
                    "module": module,
                    "line_number": int(chunk["source_line"]),
                    "record": int(match["record"]),
                    "field": int(match["field"]),
                    "value": chunk["value"],
                }
            )
    return fields


def annotate_rb_chunks(
    chunks: list[dict[str, object]], record_prefix: str
) -> tuple[dict[tuple[int, int], int], set[int]]:
    pattern = re.compile(
        rf"^{re.escape(record_prefix)}(?P<record>\d{{3}})(?P<field>\d{{3}})$"
    )
    index: dict[tuple[int, int], int] = {}
    records: set[int] = set()
    for idx, chunk in enumerate(chunks):
        match = pattern.fullmatch(str(chunk["code"]))
        if not match:
            continue
        record = int(match["record"])
        field = int(match["field"])
        index[(record, field)] = idx
        records.add(record)
    return index, records


def insert_rb_chunk(
    chunks: list[dict[str, object]],
    record_prefix: str,
    record: int,
    field: int,
    formatted_value: str,
) -> None:
    """Inserisce un campo RB mantenendo l'ordine logico dei campi RB del modulo."""
    pattern = re.compile(
        rf"^{re.escape(record_prefix)}(?P<record>\d{{3}})(?P<field>\d{{3}})$"
    )
    _index, records = annotate_rb_chunks(chunks, record_prefix)
    if record not in records:
        raise RpfError(f"Il record {record_prefix}{record:03d} non esiste nel modulo RPF.")

    target = (record, field)
    rb_positions: list[tuple[int, tuple[int, int]]] = []
    for idx, chunk in enumerate(chunks):
        match = pattern.fullmatch(str(chunk["code"]))
        if match:
            rb_positions.append((idx, (int(match["record"]), int(match["field"]))))

    insertion_index: int | None = None
    last_rb_index: int | None = None
    for idx, pair in rb_positions:
        last_rb_index = idx
        if pair > target:
            insertion_index = idx
            break
    if insertion_index is None:
        if last_rb_index is None:
            raise RpfError("Nessun campo RB presente nel modulo.")
        insertion_index = last_rb_index + 1

    chunks.insert(
        insertion_index,
        {"code": f"{record_prefix}{record:03d}{field:03d}", "value": formatted_value, "source_line": 0},
    )


def repack_c_modules(lines: list[str], modules: dict[int, dict[str, object]]) -> list[str]:
    """Ricostruisce i record C per modulo, creando record C aggiuntivi se necessari.

    Ogni record C conserva 89 caratteri posizionali, 75 coppie codice/valore da
    24 caratteri e la coda originale. Gli eventuali record aggiunti mantengono
    lo stesso Progressivo modulo e sono inseriti subito dopo l'ultimo record C
    di quel modulo. Il conteggio dei record C nel record Z viene aggiornato.
    """
    replacements: dict[int, str] = {}
    additions_after: dict[int, list[str]] = {}
    total_c_records = 0

    for module, info in modules.items():
        line_indexes: list[int] = list(info["line_indexes"])
        templates: list[tuple[str, str]] = list(info["templates"])
        chunks: list[dict[str, object]] = list(info["chunks"])
        needed = max(1, (len(chunks) + C_CHUNKS_PER_RECORD - 1) // C_CHUNKS_PER_RECORD)
        record_count = max(len(line_indexes), needed)
        total_c_records += record_count

        for record_no in range(record_count):
            template_body, template_ending = templates[min(record_no, len(templates) - 1)]
            start = record_no * C_CHUNKS_PER_RECORD
            selected = chunks[start : start + C_CHUNKS_PER_RECORD]
            payload = "".join(str(c["code"]) + str(c["value"]) for c in selected)
            payload = payload.ljust(C_PAYLOAD_LENGTH)
            new_body = (
                template_body[:C_PAYLOAD_START]
                + payload
                + template_body[C_PAYLOAD_END:]
            )
            new_line = new_body + template_ending
            if record_no < len(line_indexes):
                replacements[line_indexes[record_no]] = new_line
            else:
                additions_after.setdefault(line_indexes[-1], []).append(new_line)

    rebuilt: list[str] = []
    for line_index, line in enumerate(lines):
        current = replacements.get(line_index, line)
        rebuilt.append(current)
        rebuilt.extend(additions_after.get(line_index, []))

    z_indexes = [i for i, line in enumerate(rebuilt) if _split_line_ending(line)[0].startswith("Z")]
    if len(z_indexes) != 1:
        raise RpfError("Il file RPF deve contenere un solo record Z.")
    z_index = z_indexes[0]
    z_body, z_ending = _split_line_ending(rebuilt[z_index])
    if len(z_body) != C_BODY_LENGTH:
        raise RpfError("Il record Z non ha la lunghezza prevista di 1898 caratteri prima di CR/LF.")
    z_body = z_body[:Z_C_COUNT_START] + f"{total_c_records:09d}" + z_body[Z_C_COUNT_END:]
    rebuilt[z_index] = z_body + z_ending
    return rebuilt


def is_numeric_field(value: str) -> bool:
    return bool(value[:1].isspace() and NUMERIC_VALUE.fullmatch(value.strip()))


def excel_value(value: str) -> str | int | float:
    stripped = value.strip()
    if not is_numeric_field(value):
        return stripped
    normalized = stripped.replace(",", ".")
    return float(normalized) if "," in stripped else int(normalized)


def grouped_records(
    fields: list[dict[str, object]],
    record_prefix: str,
    convert_value: Callable[[str], object] = excel_value,
) -> list[dict[str, object]]:
    metadata = metadata_columns(record_prefix)
    columns = field_columns(record_prefix)
    records: dict[tuple[int, int], dict[str, object]] = {}

    for item in fields:
        key = (int(item["module"]), int(item["record"]))
        record = records.setdefault(
            key,
            {
                metadata[0]: key[0],
                metadata[1]: int(item["line_number"]),
                metadata[2]: f"{key[1]:03d}",
                **{column: "" for column in columns},
            },
        )
        record[metadata[1]] = min(int(record[metadata[1]]), int(item["line_number"]))

        field_number = int(item["field"])
        if 1 <= field_number <= MAX_FIELD_NUMBER:
            record[f"{record_prefix}{field_number:03d}"] = convert_value(str(item["value"]))

    return list(records.values())


def export_excel(
    rpf_path: Path,
    excel_path: Path,
    record_prefix: str = DEFAULT_RECORD_PREFIX,
    field_width: int = FIELD_WIDTH,
) -> int:
    record_prefix, field_width = validate_reader_settings(record_prefix, field_width)
    fields = find_rpf_fields(read_rpf(rpf_path), record_prefix, field_width)
    records = grouped_records(fields, record_prefix)
    if not records:
        raise RpfError(f"Nel file non sono stati trovati campi {record_prefix}.")

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = SHEET_NAME
    headers = (*metadata_columns(record_prefix), *field_columns(record_prefix))
    sheet.append(headers)

    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")

    sheet.freeze_panes = "D2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(records) + 1}"

    for record in records:
        sheet.append([record[column] for column in headers])

    for column in sheet.iter_cols(min_col=4, max_col=len(headers), min_row=2):
        for cell in column:
            if isinstance(cell.value, str):
                cell.number_format = "@"

    sheet.column_dimensions["A"].width = 10
    sheet.column_dimensions["B"].width = 12
    sheet.column_dimensions["C"].width = 12
    for column_index in range(4, len(headers) + 1):
        sheet.column_dimensions[get_column_letter(column_index)].width = 15

    instructions = workbook.create_sheet("Istruzioni")
    instructions["A1"] = "Istruzioni"
    instructions["A1"].font = Font(bold=True, size=14)
    instructions["A3"] = (
        f"Modificare soltanto le colonne da {record_prefix}001 a "
        f"{record_prefix}{MAX_FIELD_NUMBER:03d}. Le colonne Modulo, Riga file e Record "
        f"{record_prefix} identificano la posizione originale e non devono essere modificate."
    )
    instructions["A4"] = (
        f"Ogni valore occupa {field_width} caratteri nel file RPF. Non aggiungere o eliminare righe. "
        "Se un campo non e presente nel RPF di destinazione ma contiene un valore nel file importato, viene aggiunto automaticamente."
    )
    instructions["A5"] = (
        f"In importazione i record {record_prefix}010 e {record_prefix}011 non vengono mai modificati: "
        "sono considerati campi di calcolo e restano quelli del file RPF originale."
    )
    instructions.column_dimensions["A"].width = 120

    workbook.save(excel_path)
    return len(records)


def export_csv(
    rpf_path: Path,
    csv_path: Path,
    record_prefix: str = DEFAULT_RECORD_PREFIX,
    field_width: int = FIELD_WIDTH,
) -> int:
    record_prefix, field_width = validate_reader_settings(record_prefix, field_width)
    fields = find_rpf_fields(read_rpf(rpf_path), record_prefix, field_width)
    records = grouped_records(fields, record_prefix, lambda value: value.strip())
    if not records:
        raise RpfError(f"Nel file non sono stati trovati campi {record_prefix}.")

    headers = (*metadata_columns(record_prefix), *field_columns(record_prefix))
    with csv_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=headers, delimiter=";", extrasaction="raise")
        writer.writeheader()
        writer.writerows(records)
    return len(records)


def align_value(value: str, previous: str, field_width: int) -> str:
    value = value.strip()
    if len(value) > field_width:
        raise RpfError(f"Il valore '{value}' supera i {field_width} caratteri consentiti.")
    if not value:
        return " " * field_width
    return value.rjust(field_width) if is_numeric_field(previous) else value.ljust(field_width)


def value_for_rpf(value: object) -> str:
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float):
        return format(value, "g").replace(".", ",")
    return str(value)


def has_same_numeric_value(value: str, previous: str) -> bool:
    if not NUMERIC_VALUE.fullmatch(value.strip()):
        return False
    try:
        return Decimal(value.strip().replace(",", ".")) == Decimal(
            previous.strip().replace(",", ".")
        )
    except InvalidOperation:
        return False


def required_column_indexes(sheet, headers: tuple[str, ...]) -> dict[str, int]:
    headers_by_name = {
        str(cell.value).strip(): cell.column
        for cell in sheet[1]
        if cell.value is not None
    }
    missing = [column for column in headers if column not in headers_by_name]
    if missing:
        raise RpfError(f"Nel foglio mancano le colonne: {', '.join(missing)}.")
    return headers_by_name


def import_metadata_columns_from_names(names: set[str], record_prefix: str) -> tuple[str, str, str]:
    """Accetta sia la nuova intestazione 'Modulo' sia la vecchia 'Pagina'."""
    module_column = "Modulo" if "Modulo" in names else "Pagina" if "Pagina" in names else "Modulo"
    return (module_column, "Riga file", f"Record {record_prefix}")




def format_new_field_value(value: object, field_width: int) -> str:
    """Formatta un valore per un campo non ancora presente nel RPF."""
    text = value_for_rpf(value).strip()
    if len(text) > field_width:
        raise RpfError(f"Il valore '{text}' supera i {field_width} caratteri consentiti.")
    if not text:
        return " " * field_width
    return text.rjust(field_width) if NUMERIC_VALUE.fullmatch(text) else text.ljust(field_width)

def import_excel(
    rpf_path: Path,
    excel_path: Path,
    output_path: Path,
    record_prefix: str = DEFAULT_RECORD_PREFIX,
    field_width: int = FIELD_WIDTH,
    max_field: int = DEFAULT_MAX_IMPORT_FIELD,
) -> int:
    record_prefix, field_width = validate_reader_settings(record_prefix, field_width)
    if field_width != FIELD_WIDTH:
        raise RpfError("Per il formato RPF dei record C la lunghezza campo deve essere 16.")
    if not 1 <= max_field <= MAX_FIELD_NUMBER:
        raise RpfError("L'ultimo campo da importare deve essere compreso tra 1 e 20.")

    fields_columns = field_columns(record_prefix)[:max_field]
    lines = read_rpf(rpf_path)
    modules = parse_c_modules(lines)

    workbook = load_workbook(excel_path, data_only=False)
    if SHEET_NAME not in workbook.sheetnames:
        raise RpfError(f"Il file Excel deve contenere il foglio '{SHEET_NAME}'.")
    sheet = workbook[SHEET_NAME]
    header_names = {str(cell.value).strip() for cell in sheet[1] if cell.value is not None}
    metadata = import_metadata_columns_from_names(header_names, record_prefix)
    columns = required_column_indexes(sheet, (*metadata, *fields_columns))

    seen_records: set[tuple[int, int]] = set()
    imported_values = 0

    for row_number in range(2, sheet.max_row + 1):
        identifiers = [sheet.cell(row_number, columns[column]).value for column in metadata]
        if all(value is None for value in identifiers):
            continue
        if any(value is None for value in identifiers):
            raise RpfError(f"Identificativi incompleti nella riga Excel {row_number}.")

        try:
            module = int(identifiers[0])
            _informative_line = int(identifiers[1])
            record = int(identifiers[2])
        except (TypeError, ValueError) as error:
            raise RpfError(f"Identificativi non validi nella riga Excel {row_number}.") from error

        # Questi record sono risultati di calcolo: non vengono mai importati.
        if record in NON_IMPORTABLE_RECORDS:
            continue

        record_key = (module, record)
        if record_key in seen_records:
            raise RpfError(f"Il record nella riga Excel {row_number} e duplicato.")
        seen_records.add(record_key)

        module_info = modules.get(module)
        if module_info is None:
            raise RpfError(f"Il Modulo {module} non esiste nel file RPF di destinazione.")
        chunks = module_info["chunks"]
        rb_index, target_records = annotate_rb_chunks(chunks, record_prefix)
        if record not in target_records:
            if any(
                sheet.cell(row_number, columns[column]).value is not None
                and str(sheet.cell(row_number, columns[column]).value).strip()
                for column in fields_columns
            ):
                raise RpfError(
                    f"Il record {record_prefix}{record:03d} del Modulo {module} "
                    f"(riga Excel {row_number}) non esiste nel file RPF di destinazione."
                )
            continue

        for field_number, column_name in enumerate(fields_columns, start=1):
            cell_value = sheet.cell(row_number, columns[column_name]).value
            has_value = cell_value is not None and str(cell_value).strip() != ""

            rb_index, _target_records = annotate_rb_chunks(chunks, record_prefix)
            chunk_index = rb_index.get((record, field_number))

            if not has_value:
                if chunk_index is not None:
                    del chunks[chunk_index]
                    imported_values += 1
                continue

            new_text = value_for_rpf(cell_value)
            if chunk_index is not None:
                previous = chunks[chunk_index]["value"]
                formatted = (
                    previous
                    if is_numeric_field(previous) and has_same_numeric_value(new_text, previous)
                    else align_value(new_text, previous, field_width)
                )
                if formatted != previous:
                    chunks[chunk_index]["value"] = formatted
                    imported_values += 1
            else:
                formatted = format_new_field_value(cell_value, field_width)
                insert_rb_chunk(chunks, record_prefix, record, field_number, formatted)
                imported_values += 1

    lines = repack_c_modules(lines, modules)
    with output_path.open("w", encoding="ascii", newline="") as file:
        file.writelines(lines)
    return imported_values


def import_csv(
    rpf_path: Path,
    csv_path: Path,
    output_path: Path,
    record_prefix: str = DEFAULT_RECORD_PREFIX,
    field_width: int = FIELD_WIDTH,
    max_field: int = DEFAULT_MAX_IMPORT_FIELD,
) -> int:
    record_prefix, field_width = validate_reader_settings(record_prefix, field_width)
    if field_width != FIELD_WIDTH:
        raise RpfError("Per il formato RPF dei record C la lunghezza campo deve essere 16.")
    if not 1 <= max_field <= MAX_FIELD_NUMBER:
        raise RpfError("L'ultimo campo da importare deve essere compreso tra 1 e 20.")

    fields_columns = field_columns(record_prefix)[:max_field]
    lines = read_rpf(rpf_path)
    modules = parse_c_modules(lines)

    seen_records: set[tuple[int, int]] = set()
    imported_values = 0

    with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file, delimiter=";")
        if reader.fieldnames is None:
            raise RpfError("Il file CSV e vuoto.")
        metadata = import_metadata_columns_from_names(set(reader.fieldnames), record_prefix)
        headers = (*metadata, *fields_columns)
        missing = [column for column in headers if column not in reader.fieldnames]
        if missing:
            raise RpfError(f"Nel CSV mancano le colonne: {', '.join(missing)}.")

        for row_number, row in enumerate(reader, start=2):
            identifiers = [row[column] for column in metadata]
            if any(value is None or not value.strip() for value in identifiers):
                raise RpfError(f"Identificativi incompleti nella riga CSV {row_number}.")

            try:
                module = int(identifiers[0])
                _informative_line = int(identifiers[1])
                record = int(identifiers[2])
            except ValueError as error:
                raise RpfError(f"Identificativi non validi nella riga CSV {row_number}.") from error

            if record in NON_IMPORTABLE_RECORDS:
                continue

            record_key = (module, record)
            if record_key in seen_records:
                raise RpfError(f"Il record nella riga CSV {row_number} e duplicato.")
            seen_records.add(record_key)

            module_info = modules.get(module)
            if module_info is None:
                raise RpfError(f"Il Modulo {module} non esiste nel file RPF di destinazione.")
            chunks = module_info["chunks"]
            rb_index, target_records = annotate_rb_chunks(chunks, record_prefix)
            if record not in target_records:
                if any((row[column] or "").strip() for column in fields_columns):
                    raise RpfError(
                        f"Il record {record_prefix}{record:03d} del Modulo {module} "
                        f"(riga CSV {row_number}) non esiste nel file RPF di destinazione."
                    )
                continue

            for field_number, column_name in enumerate(fields_columns, start=1):
                raw_value = row[column_name] or ""
                has_value = raw_value.strip() != ""
                rb_index, _target_records = annotate_rb_chunks(chunks, record_prefix)
                chunk_index = rb_index.get((record, field_number))

                if not has_value:
                    if chunk_index is not None:
                        del chunks[chunk_index]
                        imported_values += 1
                    continue

                if chunk_index is not None:
                    previous = chunks[chunk_index]["value"]
                    formatted = (
                        previous
                        if is_numeric_field(previous) and has_same_numeric_value(raw_value, previous)
                        else align_value(raw_value, previous, field_width)
                    )
                    if formatted != previous:
                        chunks[chunk_index]["value"] = formatted
                        imported_values += 1
                else:
                    formatted = format_new_field_value(raw_value, field_width)
                    insert_rb_chunk(chunks, record_prefix, record, field_number, formatted)
                    imported_values += 1

    lines = repack_c_modules(lines, modules)
    with output_path.open("w", encoding="ascii", newline="") as file:
        file.writelines(lines)
    return imported_values

class Application(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Convertitore campi RPF")
        self.resizable(False, False)

        frame = tk.Frame(self, padx=24, pady=24)
        frame.pack()

        self.record_prefix = tk.StringVar(value=DEFAULT_RECORD_PREFIX)
        self.field_width = tk.StringVar(value=str(FIELD_WIDTH))

        tk.Label(frame, text="Convertitore RPF - Excel e CSV", font=("Arial", 16, "bold")).pack(pady=(0, 12))
        tk.Label(
            frame,
            text=(
                "Esporta i campi RPF in Excel o CSV e riporta le modifiche nel file originale.\n"
                "Ogni riga fisica contenente record del prefisso selezionato viene trattata come un modulo.\n"
                "I record RB010 e RB011 sono esclusi dall'importazione."
            ),
            wraplength=520,
            justify="center",
        ).pack(pady=(0, 18))

        settings = tk.LabelFrame(frame, text="Lettura record", padx=10, pady=8)
        settings.pack(fill="x", pady=(0, 14))
        tk.Label(settings, text="Prefisso (2 lettere):").grid(row=0, column=0, sticky="w")
        tk.Entry(settings, textvariable=self.record_prefix, width=6).grid(row=0, column=1, padx=(8, 20))
        tk.Label(settings, text="Lunghezza campo:").grid(row=0, column=2, sticky="w")
        tk.Entry(settings, textvariable=self.field_width, width=6).grid(row=0, column=3, padx=(8, 0))

        tk.Button(frame, text="1. Esporta RPF in Excel", width=32, command=self.export).pack(pady=4)
        tk.Button(frame, text="2. Importa Excel in RPF", width=32, command=self.import_changes).pack(pady=4)
        tk.Button(frame, text="3. Esporta RPF in CSV", width=32, command=self.export_csv_action).pack(pady=4)
        tk.Button(frame, text="4. Importa CSV in RPF", width=32, command=self.import_csv_action).pack(pady=4)

    def reader_settings(self) -> tuple[str, int] | None:
        try:
            return validate_reader_settings(self.record_prefix.get(), int(self.field_width.get()))
        except ValueError:
            messagebox.showerror("Configurazione non valida", "La lunghezza del campo deve essere un numero intero positivo.")
        except RpfError as error:
            messagebox.showerror("Configurazione non valida", str(error))
        return None

    def ask_max_field(self) -> int | None:
        return simpledialog.askinteger(
            "Campi da importare",
            "Fino a quale campo vuoi importare?\n\n"
            "Inserisci un numero da 1 a 20.\n"
            "Esempio: 8 importa i campi da RB001 a RB008.\n"
            "I record RB010 e RB011 vengono sempre esclusi dall'importazione.",
            parent=self,
            initialvalue=DEFAULT_MAX_IMPORT_FIELD,
            minvalue=1,
            maxvalue=20,
        )

    def export(self) -> None:
        settings = self.reader_settings()
        if settings is None:
            return
        source = filedialog.askopenfilename(title="Seleziona il file RPF", filetypes=[("File RPF", "*.rpf"), ("Tutti i file", "*.*")])
        if not source:
            return
        destination = filedialog.asksaveasfilename(title="Salva il file Excel", defaultextension=".xlsx", filetypes=[("File Excel", "*.xlsx")])
        if not destination:
            return
        try:
            count = export_excel(Path(source), Path(destination), *settings)
            messagebox.showinfo("Esportazione completata", f"Creati {count} record nel file Excel.")
        except (OSError, RpfError) as error:
            messagebox.showerror("Esportazione non riuscita", str(error))

    def export_csv_action(self) -> None:
        settings = self.reader_settings()
        if settings is None:
            return
        source = filedialog.askopenfilename(title="Seleziona il file RPF", filetypes=[("File RPF", "*.rpf"), ("Tutti i file", "*.*")])
        if not source:
            return
        destination = filedialog.asksaveasfilename(title="Salva il file CSV", defaultextension=".csv", filetypes=[("File CSV", "*.csv")])
        if not destination:
            return
        try:
            count = export_csv(Path(source), Path(destination), *settings)
            messagebox.showinfo("Esportazione completata", f"Creati {count} record nel file CSV.")
        except (OSError, RpfError) as error:
            messagebox.showerror("Esportazione non riuscita", str(error))

    def import_changes(self) -> None:
        settings = self.reader_settings()
        if settings is None:
            return
        max_field = self.ask_max_field()
        if max_field is None:
            return
        source = filedialog.askopenfilename(title="Seleziona il file RPF originale", filetypes=[("File RPF", "*.rpf"), ("Tutti i file", "*.*")])
        if not source:
            return
        excel = filedialog.askopenfilename(title="Seleziona il file Excel modificato", filetypes=[("File Excel", "*.xlsx")])
        if not excel:
            return
        destination = filedialog.asksaveasfilename(title="Salva il file RPF modificato", defaultextension=".rpf", filetypes=[("File RPF", "*.rpf")])
        if not destination:
            return
        try:
            count = import_excel(Path(source), Path(excel), Path(destination), *settings, max_field)
            messagebox.showinfo(
                "Importazione completata",
                f"Aggiornati {count} campi (da {settings[0]}001 a {settings[0]}{max_field:03d}).\n"
                f"I record {settings[0]}010 e {settings[0]}011 sono stati esclusi.",
            )
        except (OSError, RpfError) as error:
            messagebox.showerror("Importazione non riuscita", str(error))

    def import_csv_action(self) -> None:
        settings = self.reader_settings()
        if settings is None:
            return
        max_field = self.ask_max_field()
        if max_field is None:
            return
        source = filedialog.askopenfilename(title="Seleziona il file RPF originale", filetypes=[("File RPF", "*.rpf"), ("Tutti i file", "*.*")])
        if not source:
            return
        csv_file = filedialog.askopenfilename(title="Seleziona il file CSV modificato", filetypes=[("File CSV", "*.csv")])
        if not csv_file:
            return
        destination = filedialog.asksaveasfilename(title="Salva il file RPF modificato", defaultextension=".rpf", filetypes=[("File RPF", "*.rpf")])
        if not destination:
            return
        try:
            count = import_csv(Path(source), Path(csv_file), Path(destination), *settings, max_field)
            messagebox.showinfo(
                "Importazione completata",
                f"Aggiornati {count} campi (da {settings[0]}001 a {settings[0]}{max_field:03d}).\n"
                f"I record {settings[0]}010 e {settings[0]}011 sono stati esclusi.",
            )
        except (OSError, RpfError) as error:
            messagebox.showerror("Importazione non riuscita", str(error))


def add_common_options(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--prefix", default=DEFAULT_RECORD_PREFIX, help="Prefisso di due lettere, predefinito RB")
    parser.add_argument("--field-width", type=int, default=FIELD_WIDTH, help="Lunghezza del valore, predefinita 16")


def main() -> int:
    parser = argparse.ArgumentParser(description="Convertitore di campi RPF tra RPF, Excel e CSV.")
    subparsers = parser.add_subparsers(dest="command")

    export_parser = subparsers.add_parser("export", help="Esporta un RPF in Excel")
    export_parser.add_argument("rpf", type=Path)
    export_parser.add_argument("excel", type=Path)
    add_common_options(export_parser)

    import_parser = subparsers.add_parser("import", help="Importa le modifiche Excel in un RPF")
    import_parser.add_argument("rpf", type=Path)
    import_parser.add_argument("excel", type=Path)
    import_parser.add_argument("output", type=Path)
    add_common_options(import_parser)
    import_parser.add_argument(
        "--max-field", type=int, default=DEFAULT_MAX_IMPORT_FIELD, metavar="N",
        help="Importa solo i campi da 001 fino a N (1-20, predefinito 12); i record RB010 e RB011 sono sempre esclusi",
    )

    export_csv_parser = subparsers.add_parser("export-csv", help="Esporta un RPF in CSV")
    export_csv_parser.add_argument("rpf", type=Path)
    export_csv_parser.add_argument("csv", type=Path)
    add_common_options(export_csv_parser)

    import_csv_parser = subparsers.add_parser("import-csv", help="Importa le modifiche CSV in un RPF")
    import_csv_parser.add_argument("rpf", type=Path)
    import_csv_parser.add_argument("csv", type=Path)
    import_csv_parser.add_argument("output", type=Path)
    add_common_options(import_csv_parser)
    import_csv_parser.add_argument(
        "--max-field", type=int, default=DEFAULT_MAX_IMPORT_FIELD, metavar="N",
        help="Importa solo i campi da 001 fino a N (1-20, predefinito 12); i record RB010 e RB011 sono sempre esclusi",
    )

    args = parser.parse_args()

    try:
        if args.command == "export":
            print(f"Esportati {export_excel(args.rpf, args.excel, args.prefix, args.field_width)} record.")
        elif args.command == "import":
            print(f"Aggiornati {import_excel(args.rpf, args.excel, args.output, args.prefix, args.field_width, args.max_field)} campi.")
        elif args.command == "export-csv":
            print(f"Esportati {export_csv(args.rpf, args.csv, args.prefix, args.field_width)} record.")
        elif args.command == "import-csv":
            print(f"Aggiornati {import_csv(args.rpf, args.csv, args.output, args.prefix, args.field_width, args.max_field)} campi.")
        else:
            Application().mainloop()
    except (OSError, RpfError) as error:
        print(f"Errore: {error}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
