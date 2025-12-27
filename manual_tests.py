from pathlib import Path
from py_fit_export.wrk_info_export import (
    export_to_json,
    export_activity_to_excel,
    export_activities_to_excel,
)
from py_fit_export.fit_info_extractor import FitInfoExtractor


def print_key_info(activity):
    extractor = FitInfoExtractor(activity)
    print(extractor.extract())


def test_export_to_json(activity, out_path):
    extractor = FitInfoExtractor(activity)
    export_to_json(out_path, extractor.fit)


def test_export_activity(activity, out_path, column_map):
    export_activity_to_excel(
        activity, out_path, column_map, "Running sessions", "tblRun"
    )


def test_export_activities(activities, out_path, column_map):
    export_activities_to_excel(
        activities, out_path, column_map, "Running sessions", "tblRun"
    )


if __name__ == "__main__":
    test_activity = Path("ACTIVITY.fit")
    test_activity2 = Path("ACTIVITY2.fit")
    test_activities = [test_activity, test_activity2]
    out = Path("test_workbook.xlsx")
    assert test_activity.is_file()
    columns = {
        "wrk_start_time": "Date",
        "wrk_name": "Name",
        "wrk_length": "Distance",
        "wrk_load": "Load",
    }

    # print_key_info(test_activity)
    # test_export_to_json(test_activity, Path("activity.json"))
    test_export_activity(test_activity, out, columns)
    # test_export_activities(test_activities, out, columns)
