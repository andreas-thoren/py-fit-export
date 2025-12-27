from datetime import datetime
from pathlib import Path
from typing import Any, Iterable
from garmin_fit_sdk import Decoder, Stream
from py_fit_export.utils import excel_safe_datetime


class FitInfoExtractor:

    def __init__(self, activity_path: Path):
        self.fit = self._extract_fit_info(activity_path)
        self._workout = self._extract_info_dict("workout_mesgs")
        self._session = self._extract_info_dict("session_mesgs")


    def _extract_fit_info(self, activity_path: Path) -> dict[str, Any]:
        stream = Stream.from_file(activity_path)
        decoder = Decoder(stream)
        fit_info, fit_errors = decoder.read()

        if fit_errors:
            print(fit_errors)

        return fit_info


    def _extract_info_dict(self, fit_info_key: str) -> dict[str, Any]:
        list_container = self.fit.get(fit_info_key)
        if isinstance(list_container, list) and list_container:
            info_dict = list_container[0]
            return info_dict if isinstance(info_dict, dict) else {}
        return {}


    # --- fields ---------------------------------------------------------------
    # To add a field:
    # - Add both a method and the name of method to class attribute FIELDS.

    FIELDS = frozenset(
        (
            "wrk_sport",
            "wrk_start_time",
            "wrk_name",
            "wrk_length",
            "wrk_load",
        )
    )
    def wrk_sport(self) -> str | None:
        return self._session.get("sport")

    def wrk_start_time(self) -> datetime | None:
        v = self._session.get("start_time")
        return excel_safe_datetime(v) if isinstance(v, datetime) else None

    def wrk_name(self) -> str | None:
        return self._workout.get("wkt_name")

    def wrk_length(self) -> float | None:
        return self._session.get("total_distance")

    def wrk_load(self) -> float | None:
        return self._session.get("training_load_peak")

    # --- extraction ---
    def extract(self, fields: Iterable[str] | None = None) -> dict[str, Any]:
        out: dict[str, Any] = {}
        for name in fields or self.FIELDS:
            if name not in self.FIELDS:
                raise KeyError(f"Unknown field: {name}") from None
            attr = getattr(self, name)
            out[name] = attr()

        return out
