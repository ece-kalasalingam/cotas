import os
import tempfile
from openpyxl.workbook import Workbook
from core.exceptions import ValidationError


class AtomicWorkbookWriter:
    """
    Handles safe, atomic workbook saving.
    No Excel logic.
    No business logic.
    """

    @staticmethod
    def save(workbook: Workbook, output_path: str) -> None:
        # Check if file is open (Windows-safe check)
        if os.path.exists(output_path):
            try:
                os.rename(output_path, output_path)
            except OSError:
                raise ValidationError(
                    f"File '{os.path.basename(output_path)}' "
                    f"is open in Excel."
                )

        temp_dir = os.path.dirname(output_path)
        fd, temp_path = tempfile.mkstemp(
            suffix=".xlsx",
            dir=temp_dir,
        )
        os.close(fd)

        try:
            workbook.save(temp_path)

            if os.path.exists(output_path):
                os.remove(output_path)

            os.replace(temp_path, output_path)

        except Exception:
            if os.path.exists(temp_path):
                os.remove(temp_path)
            raise