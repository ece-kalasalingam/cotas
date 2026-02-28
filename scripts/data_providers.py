import json
from pathlib import Path

from scripts.constants import ID_COURSE_SETUP


class UniversalDataProvider:
    @staticmethod
    def _default_setup_data() -> dict:
        return {
            "Course_Metadata": [
                ["Course_Code", ""],
                ["Section", ""],
                ["Semester", ""],
                ["Academic_Year", ""],
                ["Total_COs", ""],
            ],
            "Assessment_Config": [],
            "Question_Map": [],
            "Students": [],
        }

    @staticmethod
    def _load_setup_sample_from_assets() -> dict:
        project_root = Path(__file__).resolve().parent.parent
        sample_path = project_root / "assets" / "sample_setup_data.json"
        if not sample_path.exists():
            return UniversalDataProvider._default_setup_data()

        try:
            payload = json.loads(sample_path.read_text(encoding="utf-8"))
        except Exception:
            return UniversalDataProvider._default_setup_data()

        if not isinstance(payload, dict):
            return UniversalDataProvider._default_setup_data()

        defaults = UniversalDataProvider._default_setup_data()
        cleaned: dict = {}
        for key in defaults.keys():
            value = payload.get(key, defaults[key])
            cleaned[key] = value if isinstance(value, list) else defaults[key]

        return cleaned

    @staticmethod
    def get_data_for_template(type_id: str) -> dict:
        if type_id != ID_COURSE_SETUP:
            return {}
        return UniversalDataProvider._load_setup_sample_from_assets()

