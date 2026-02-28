from scripts.exceptions import ValidationError


class AtomicWorkbookWriter:
    @staticmethod
    def finalize(workbook, final_path: str) -> None:
        try:
            workbook.close()
        except Exception as exc:
            raise ValidationError(f"Failed to finalize workbook: {exc}") from exc
        _ = final_path
