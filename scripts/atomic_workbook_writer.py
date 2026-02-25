import os
import tempfile
import xlsxwriter
from typing import cast
from core.exceptions import ValidationError

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
        """
        Closes the workbook and performs the atomic swap to the final destination.
        This is the 'Point of No Return'.
        """
        # 1. Retrieve the temp path used during initialization
        temp_path = cast(str, workbook.filename)

        try:
            # 2. Close handles and flush data to the temp file
            workbook.close()

            # 3. Windows Lock Check: Try to rename the file to itself
            if os.path.exists(final_path):
                try:
                    os.rename(final_path, final_path)
                except OSError:
                    raise ValidationError(
                        f"File '{os.path.basename(final_path)}' is open in Excel. Please close it."
                    )
                # Remove the old version to make room for os.replace
                os.remove(final_path)

            # 4. Perform the final atomic swap
            os.replace(temp_path, final_path)

        except Exception as e:
            # Clean up the temp file if the final swap or close failed
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)
            # Re-wrap or re-raise for the Renderer to catch
            if isinstance(e, ValidationError):
                raise e
            raise ValidationError(f"Atomic move failed: {str(e)}")