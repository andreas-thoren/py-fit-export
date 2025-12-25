import json
from copy import copy
from pathlib import Path
from openpyxl.utils import range_boundaries, get_column_letter
from openpyxl.formula.translate import Translator
from garmin_fit_sdk import Decoder, Stream

class FitExporter:
    def __init__(self, file_path: Path):
        stream = Stream.from_file(file_path)
        decoder = Decoder(stream)
        self.messages, self.errors = decoder.read()

        if self.errors:
            print(self.errors)

    def extract_key_info(self):
        # Maybe later allow for passing in which info should be extracted.
        # Would need to define structure for message dict in this case.
        session = self.messages["session_mesgs"][0]
        key_info = {}
        key_info["wrk_sport"] = session["sport"]
        key_info["wrk_date"] = str(session["start_time"].date())
        key_info["wrk_name"] = self.messages["workout_mesgs"][0]["wkt_name"]
        key_info["wrk_length"] = session["total_distance"]
        key_info["wrk_load"] = session["training_load_peak"]
        return key_info

    def export_to_json(self, export_path: Path, info_dct: dict | None = None):
        info_dct = info_dct if info_dct is not None else self.messages
        safe_messages = self._make_json_safe(info_dct)
        export_path.write_text(
            json.dumps(safe_messages, indent=2, ensure_ascii=False),
            encoding="utf-8"
        )

    def _make_json_safe(self, obj):
        if isinstance(obj, dict):
            return {k: self._make_json_safe(v) for k, v in obj.items()}
        if isinstance(obj, list):
            return [self._make_json_safe(v) for v in obj]
        if hasattr(obj, "isoformat"):  # datetime
            return obj.isoformat()
        if isinstance(obj, bytes):
            return obj.hex()
        return obj

    @staticmethod
    def _make_ref(min_col, min_row, max_col, max_row):
        return f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"    

    @staticmethod
    def append_table_values(ws, table, row_values: dict):
        # 1. Raise early if tabel has totals row since function will not work!
        if table.totalsRowShown:
            raise RuntimeError(
                "append_table_values() cannot be used on tables with a totals row; "
                "this requires row insertion logic."
            )

        # 2. Construct initial required metadata about table and new row
        columns = { tc.name: i for i, tc in enumerate(table.tableColumns) }
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        row_i = max_row + 1
        column_index_range = range(min_col, max_col+1)

        # 3. Error checks. Check row_values type and content. Also check new row is empty!
        if not isinstance(row_values, dict):
            raise TypeError(f"row_values has to be a dict not a {type(row_values)}!")

        faulty_keys = row_values.keys() - columns.keys()
        if faulty_keys:
            bad = ", ".join(sorted(map(str, faulty_keys)))
            raise ValueError(f"row_values contains invalid column names: {bad}")

        if any(ws.cell(row=row_i, column=c).value is not None for c in column_index_range):
            raise ValueError(f"Row {row_i} is not empty in the table area; would overwrite data!")

        # 4. Set cell values for new row
        for col, val in row_values.items():
            col_i = min_col + columns[col]
            ws.cell(row=row_i, column=col_i, value=val)

        # 5. Fix formating of new row and its cells. Also expand formulas to new rows!
        if max_row > min_row:
            ws.row_dimensions[row_i].height = ws.row_dimensions[max_row].height

            for col_i in column_index_range:
                src = ws.cell(row=max_row, column=col_i)
                dst = ws.cell(row=row_i, column=col_i)

                if src.has_style:
                    dst._style = copy(src._style)
                if dst.value is None and (isinstance(src.value, str) and src.value.startswith("=")):
                    dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)

        # 6. Resize table to include new row
        new_ref = FitExporter._make_ref(min_col, min_row, max_col, row_i)
        table.ref = new_ref
