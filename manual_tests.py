from pathlib import Path
from py_fit_export.core import (
    extract_fit_info,
    extract_key_info,
    export_to_json,
    export_activity_to_excel,
    export_activities_to_excel,
)


def print_key_info(activity):
    fit_info = extract_fit_info(activity)
    print(extract_key_info(fit_info))


def test_export_to_json(activity, out_path):
    fit_info = extract_fit_info(activity)
    export_to_json(out_path, fit_info)


def test_export_activity(activity, out_path, column_map):
    export_activity_to_excel(
        activity, column_map, out_path, "Running sessions", "tblRun"
    )


def test_export_activities(activities, out_path, column_map):
    export_activities_to_excel(
        activities, column_map, out_path, "Running sessions", "tblRun"
    )


if __name__ == "__main__":
    test_activity = Path("ACTIVITY.fit")
    test_activity2 = Path("ACTIVITY2.fit")
    test_activities = [test_activity, test_activity2]
    out = Path("test_workbook.xlsx")
    assert test_activity.is_file()
    columns = {
        "wrk_date": "Date",
        "wrk_name": "Name",
        "wrk_length": "Distance",
        "wrk_load": "Load",
    }

    # append_table_records(test_activity, Path("test_workbook.xlsx"), "Running sessions", "tblRun")
    # print_key_info(test_activity)
    # test_export_to_json(test_activity, Path("activity.json"))
    # test_export_activity(test_activity, out, columns)
    # test_export_activities(test_activities, out, columns)
