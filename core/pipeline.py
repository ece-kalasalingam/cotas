from core.loader import SetupLoader
from core.setup_validator import SetupValidator
from core.excel_adapter import ExcelAdapter
from core.filled_validator import FilledMarksValidator
from core.calculator import COCalculator


class COPipeline:
    """
    Orchestrates full CO computation flow.
    Only this layer handles file paths.
    Logic layers remain pure.
    """

    def __init__(self, setup_path: str, filled_path: str):
        self.setup_path = setup_path
        self.filled_path = filled_path

    # =====================================================
    # RUN PIPELINE
    # =====================================================

    def run(self):
        # -----------------------------
        # 1. Load Setup
        # -----------------------------
        loader = SetupLoader(self.setup_path)

        metadata = loader.load_metadata()
        config = loader.load_config()
        students = loader.load_students()
        qmap = loader.load_question_map()

        validator = SetupValidator(metadata, config, students, qmap)
        validated = validator.validate()

        # -----------------------------
        # 2. Load Filled Sheets
        # -----------------------------
        adapter = ExcelAdapter(self.filled_path)
        all_sheets = adapter.load_all()

        # Separate direct and indirect sheets
        direct_sheets = {
            name: all_sheets[name]
            for name, comp in validated.components.items()
            if comp.direct
        }

        indirect_sheets = {
            tool.name: all_sheets[f"{tool.name}_INDIRECT"]
            for tool in validated.indirect_tools
        }

        # -----------------------------
        # 3. Validate Filled Structure
        # -----------------------------
        filled_validator = FilledMarksValidator(validated, all_sheets)
        filled_validator.validate()

        # -----------------------------
        # 4. Compute
        # -----------------------------
        calculator = COCalculator(
            validated,
            direct_sheets,
            indirect_sheets
        )

        result = calculator.run()

        return result