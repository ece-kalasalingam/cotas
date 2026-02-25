from typing import Dict, List, Any, Optional
from scripts.blueprints import COURSE_SETUP_BP #

class UniversalDataProvider:
    """
    Unit: Content Generation.
    Generates pre-filled data maps internally. 
    Supports multiple template IDs to ensure the correct data structure is provided.
    """

    @staticmethod
    def get_data_for_template(template_id: str) -> Dict[str, List[List[Any]]]:
        """
        Orchestrator to fetch pre-filled data based on the Template ID.
        """
        if template_id == "COURSE_SETUP_V1":
            return CourseSetupDataProvider.get_prefilled_data()
        # Future-proofing for other templates
        # elif template_id == "ATTAINMENT_V1":
        #     return AttainmentDataProvider.get_prefilled_data()
        
        return {}

class CourseSetupDataProvider:
    """
    Sub-Unit: Course Setup Data.
    Contains the specific sample values you provided, structured for the Engine.
    """

    @staticmethod
    def get_prefilled_data() -> Dict[str, List[List[Any]]]:
        """
        Generates the internal pre-filled data map using your sample values.
        """
        return {
            "Course_Metadata": [
                ["Course_Code", "ECE000"],
                ["Course_Name", "SAMPLE COURSE"],
                ["Section", "A"],
                ["Semester", "III"],
                ["Academic_Year", "2025-26"],
                ["Faculty_Name", "ABCCE"],
                ["Total_Outcomes", 5]
            ],
            "Assessment_Config": [
                ["S1", 17.5, "YES", "YES", "YES"],
                ["S2", 17.5, "YES", "YES", "YES"],
                ["MSP", 10, "YES", "YES", "YES"],
                ["RLP", 5, "YES", "YES", "YES"],
                ["ESP", 20, "NO", "NO", "YES"],
                ["ESE", 30, "NO", "NO", "YES"],
                ["CSURVEY", 100, "NO", "YES", "NO"]
            ],
            "Question_Map": [
                ["S1", 1, 2, 1],
                ["S1", 2, 2, 1],
                ["S1", 3, 2, 2],
                ["S2", 6, 16, 3],
                ["ESP", 1, 100, "1,2,3,4,5"]
            ],
            "Students": [
                ["R101", "STUD1"],
                ["R1032", "STUD2"]
            ]
        }