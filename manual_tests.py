from pathlib import Path
from py_fit_export.extracter import FitExtractor

test_activity = Path("ACTIVITY.fit")
test_out = Path("activity.json")
assert test_activity.is_file()
extractor = FitExtractor(test_activity)

print(extractor.extract_key_info())
# extractor.export_to_json(test_out)
