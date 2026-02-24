import pandas as pd


class ExcelAdapter:
    """
    Centralized Excel I/O adapter.
    This is the ONLY layer allowed to touch pandas Excel I/O.
    """

    def __init__(self, path: str):
        self.path = path
        self._excel = pd.ExcelFile(path)

    def sheet_names(self):
        return list(self._excel.sheet_names)

    def load_all(self) -> dict:
        sheets = {}

        for raw_name in self._excel.sheet_names:
            name = str(raw_name) # Explicitly cast for type safety
            if (
            not name.endswith("_INDIRECT")
            and name not in ("Course_Info")
            ):
                sheets[name] = pd.read_excel(
                self._excel,
                sheet_name=name,
                header=None
                )
            else:
                sheets[name] = pd.read_excel(
                self._excel,
                sheet_name=name
                )

        return sheets

    def load_direct_sheet(self, sheet_name: str):
        """
        Used when header=None is required (direct sheets).
        """
        return pd.read_excel(self._excel, sheet_name=sheet_name, header=None)