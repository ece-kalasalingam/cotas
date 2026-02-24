import os
import tempfile
import xlsxwriter
from core.exceptions import ValidationError
from typing import cast

class AtomicWorkbookWriter:
    """
    Handles safe, atomic workbook saving for XlsxWriter.
    Ensures a temporary file is used during generation to prevent corruption.
    """

    @staticmethod
    def save(workbook: xlsxwriter.Workbook, output_path: str) -> None:
        """
        Finalizes and moves the workbook to the final destination.
        Note: In XlsxWriter, the workbook is already associated with a temp path.
        This method ensures the hand-off to the final output_path is safe.
        """
        if os.path.exists(output_path):
            try:
                os.rename(output_path, output_path)
            except OSError:
                raise ValidationError(
                    f"File '{os.path.basename(output_path)}' is open in Excel."
                )

        # Use 'cast' to tell Pylance that we KNOW this is a string
        # XlsxWriter stores the path here, and since we provided a string 
        # in create_temp_path, this is safe.
        temp_path = cast(str, workbook.filename)

        try:
            workbook.close()

            if os.path.exists(output_path):
                os.remove(output_path)

            # Now Pylance will be happy because temp_path is explicitly 'str'
            os.replace(temp_path, output_path)

        except Exception as e:
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)
            raise e

    @staticmethod
    def create_temp_path(output_path: str) -> str:
        """
        Helper to generate a safe temp path in the same directory.
        Use this when initializing xlsxwriter.Workbook().
        """
        temp_dir = os.path.dirname(os.path.abspath(output_path))
        fd, temp_path = tempfile.mkstemp(
            suffix=".xlsx",
            dir=temp_dir,
        )
        os.close(fd)
        return temp_path
    @staticmethod
    def save_final_from_temp(temp_path: str, final_path: str) -> None:
        """Moves a generated temp file to the final destination safely."""
        if os.path.exists(final_path):
            try:
                # Check for Windows file lock
                os.rename(final_path, final_path)
            except OSError:
                raise Exception(f"File '{os.path.basename(final_path)}' is open in Excel. Please close it.")

        try:
            if os.path.exists(final_path):
                os.remove(final_path)
            os.replace(temp_path, final_path)
        except Exception as e:
            if os.path.exists(temp_path):
                os.remove(temp_path)
            raise e