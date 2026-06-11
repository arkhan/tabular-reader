#!/usr/bin/env python
import csv
import io
import os
from types import SimpleNamespace


def get_file_format(filename):
    if isinstance(filename, (io.BytesIO, bytes)):
        return None
    _, ext = os.path.splitext(filename)
    return ext.lower().lstrip(".")


def read_csv(source, **kwargs):
    csv_kwargs = {
        k: v for k, v in kwargs.items() if k in ["delimiter", "quotechar", "encoding"]
    }

    if isinstance(source, bytes):
        source = io.BytesIO(source)
    if isinstance(source, io.BytesIO):
        source.seek(0)
        f = io.TextIOWrapper(source, encoding=csv_kwargs.pop("encoding", "utf-8-sig"))
    else:
        f = open(source, "r", encoding=csv_kwargs.pop("encoding", "utf-8-sig"))

    try:
        reader = csv.reader(f, **csv_kwargs)
        total = list(reader)
    finally:
        if not isinstance(source, io.BytesIO):
            f.close()

    header = total[0] if total else []
    filtered_indices = [
        i for i, val in enumerate(header) if val is not None and str(val).strip() != ""
    ]
    return [
        [row[i] if i < len(row) else None for i in filtered_indices] for row in total
    ]


def read_xls(source, worksheet="", skip_rows=0, **kwargs):
    import xlrd

    if isinstance(source, bytes):
        source = io.BytesIO(source)
    if isinstance(source, io.BytesIO):
        source.seek(0)
        wb = xlrd.open_workbook(file_contents=source.read())
    else:
        wb = xlrd.open_workbook(source)

    ws = wb.sheet_by_name(worksheet) if worksheet else wb.sheet_by_index(0)

    total = [[cell.value for cell in row] for row in ws.get_rows()]
    header_idx = skip_rows if skip_rows < len(total) else 0
    header = total[header_idx] if total else []
    filtered_indices = [
        i for i, val in enumerate(header) if val is not None and str(val).strip() != ""
    ]
    return [
        [row[i] if i < len(row) else None for i in filtered_indices] for row in total
    ]


def read_xlsx(source, worksheet="", skip_rows=0, **kwargs):
    from openpyxl import load_workbook

    _OPENPYXL_KWARGS = {"read_only", "keep_vba", "data_only", "keep_links", "rich_text"}
    excel_kwargs = {k: v for k, v in kwargs.items() if k in _OPENPYXL_KWARGS}

    if isinstance(source, bytes):
        source = io.BytesIO(source)
    if isinstance(source, io.BytesIO):
        source.seek(0)

    wb = load_workbook(source, **excel_kwargs)
    ws = worksheet and wb[worksheet] or wb.active

    total = [[col.value for col in row] for row in ws]
    header_idx = skip_rows if skip_rows < len(total) else 0
    header = total[header_idx] if total else []
    filtered_indices = [
        i for i, val in enumerate(header) if val is not None and str(val).strip() != ""
    ]
    return [
        [row[i] if i < len(row) else None for i in filtered_indices] for row in total
    ]


class TabularReader:
    def __init__(
        self,
        source,
        worksheet="",
        fieldnames=None,
        restval=None,
        restkey=None,
        skip_blank_lines=False,
        skip_rows=0,
        skip_row=None,
        *args,
        **kwargs,
    ):
        if skip_row is not None and not skip_rows:
            skip_rows = skip_row
        file_format = get_file_format(source)

        if file_format is None:
            filtered_data = self._detect_and_read_bytes(source, worksheet, skip_rows=skip_rows, **kwargs)
        elif file_format == "csv":
            filtered_data = read_csv(source, **kwargs)
        elif file_format == "xlsx":
            filtered_data = read_xlsx(source, worksheet, skip_rows=skip_rows, **kwargs)
        elif file_format == "xls":
            filtered_data = read_xls(source, worksheet, skip_rows=skip_rows, **kwargs)
        else:
            raise ValueError(f"Unsupported format: {file_format}")

        if skip_rows:
            filtered_data = filtered_data[skip_rows:]

        self.reader = iter(filtered_data)
        self._fieldnames = fieldnames
        self.restkey = restkey
        self.restval = restval
        self.skip_blank_lines = skip_blank_lines
        self.line_num = 0

    def _detect_and_read_bytes(self, source, worksheet, skip_rows=0, **kwargs):
        errors = []
        if isinstance(source, bytes):
            source = io.BytesIO(source)

        for fmt, reader_func in [
            ("xlsx", read_xlsx),
            ("xls", read_xls),
            ("csv", read_csv),
        ]:
            try:
                if isinstance(source, io.BytesIO):
                    source.seek(0)
                return reader_func(source, worksheet, skip_rows=skip_rows, **kwargs)
            except Exception as e:
                errors.append((fmt, str(e)))
                if isinstance(source, io.BytesIO):
                    source.seek(0)

        details = "; ".join(f"{fmt}: {err}" for fmt, err in errors)
        raise ValueError(f"Could not detect file format. Tried: {details}")

    @property
    def fieldnames(self):
        if self._fieldnames is None:
            try:
                self._fieldnames = next(self.reader)
            except StopIteration:
                pass
        self.line_num += 1
        return self._fieldnames

    @fieldnames.setter
    def fieldnames(self, value):
        self._fieldnames = value

    def __iter__(self):
        return self

    def __next__(self):
        if self.line_num == 0:
            self.fieldnames
        row = next(self.reader)
        self.line_num += 1
        while self.skip_blank_lines and all(cell is None for cell in row):
            row = next(self.reader)
        record = SimpleNamespace(**dict(zip(self.fieldnames, row)))
        return record

    next = __next__
