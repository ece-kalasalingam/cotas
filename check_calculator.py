import pandas as pd
import numpy as np

from core.co_calculator import COCalculator
from core.models import (
    Question, ComponentInfo, Student, ValidatedSetup, IndirectToolInfo
)

def make_validated():
    students = (
        Student("R001", "A"),
        Student("R002", "B"),
    )
    questions = (
        Question("S1", "Q1", 10.0, ("1",)),
        Question("S1", "Q2", 10.0, ("2",)),
        Question("S1", "Q3", 10.0, ("1",)),  
    )
    questions1 = (
        Question("S2", "Q1", 10.0, (" 1 , 2 ",)),  
    )
    components = {
        "S1": ComponentInfo(
            weight=50.0,
            cia=True,
            co_split=True,   # toggle True/False to test both behaviors
            direct=True,
            questions=questions,
        ),
        "S2": ComponentInfo(
            weight=50.0,
            cia=True,
            co_split=False,   # toggle True/False to test both behaviors
            direct=True,
            questions=questions1,
        ),
    }
    indirect_tools = (IndirectToolInfo("Survey", 100.0),)
    return ValidatedSetup(components=components, students=students, indirect_tools=indirect_tools)

def make_direct_sheet():
    # IMPORTANT: direct sheet is read with header=None → template-like rows
    return pd.DataFrame([
        ["RegNo", "Student_Name", "Q1", "Q2", "Q3", "Total"],
        ["CO",    "",             "1",  "2",  "1", ""],
        ["Max",   "",             10,   10,   10,    ""],
        ["R001",  "A",            8,    6,    10,    ""],
        ["R002",  "B",            "A",  7,    8,     ""],  # absent in Q1
    ])
def make_direct_sheet2():
    # IMPORTANT: direct sheet is read with header=None → template-like rows
    return pd.DataFrame([
        ["RegNo", "Student_Name", "Q1", "Total"],
        ["CO",    "",             "1, 2",  ""],
        ["Max",   "",             10,   ""],
        ["R001",  "A",            8,    ""],
        ["R002",  "B",            "A",  ""],  # absent in Q1
    ])

def make_indirect_sheet():
    return pd.DataFrame({
        "RegNo": ["R001", "R002"],
        "Student_Name": ["A", "B"],
        "CO1": [5, 4],
        "CO2": [3, "A"],
    })

if __name__ == "__main__":
    validated = make_validated()

    direct_sheets = {"S1": make_direct_sheet(), "S2": make_direct_sheet2()}
    indirect_sheets = {"Survey": make_indirect_sheet()}

    calc = COCalculator(validated, direct_sheets, indirect_sheets)
    out = calc.run()
    pd.set_option("display.max_rows", None)
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", None)
    pd.set_option("display.max_colwidth", None)

    print(out["co_sheets"]["2_Direct"])