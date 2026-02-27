import os
import tempfile
import xlsxwriter
from typing import cast
from scripts.exceptions import ValidationError

class AtomicWorkbookWriter:
    """
    Handles safe, atomic workbook saving for XlsxWriter.
    Ensures the final file is only created/overwritten if generation succeeds.
    """

    @staticmethod
    def create_temp_path(output_path: str) -> str:
        """Generates a safe temp path in the same directory as the final output."""
        # Ensure the directory exists
        target_dir = os.path.dirname(os.path.abspath(output_path))
        if not os.path.exists(target_dir):
            os.makedirs(target_dir)

        fd, temp_path = tempfile.mkstemp(
            suffix=".xlsx",
            dir=target_dir,
        )
        os.close(fd)
        return temp_path

    @staticmethod
    def finalize(workbook: xlsxwriter.Workbook, final_path: str) -> None:
        temp_path = cast(str, workbook.filename)
        try:
            workbook.close()

            # Improved Atomic Swap with Retry Logic
            if os.path.exists(final_path):
                # Attempting a direct replacement; Windows will raise PermissionError if locked
                try:
                    os.replace(temp_path, final_path)
                except PermissionError:
                    raise ValidationError(f"File '{os.path.basename(final_path)}' is open. Please close it.")
            else:
                os.replace(temp_path, final_path)

        except Exception as e:
            # Clean up the temp file if the final swap or close failed
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)
            # Re-wrap or re-raise for the Renderer to catch
            if isinstance(e, ValidationError):
                raise e
            raise ValidationError(f"Atomic move failed: {str(e)}")