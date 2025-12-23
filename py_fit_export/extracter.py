import json
from pathlib import Path
from garmin_fit_sdk import Decoder, Stream

class FitExtractor:
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
