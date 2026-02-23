from dataclasses import dataclass
from typing import  Dict, Tuple
import pandas as pd


# =====================================================
# Raw Sheet Models (IO Layer)
# =====================================================

@dataclass
class CourseMetadata:
    course_code: str
    course_name: str
    section: str
    semester: str
    academic_year: str
    faculty_name: str
    total_cos: int

@dataclass
class AssessmentConfig:
    dataframe: pd.DataFrame


@dataclass
class StudentList:
    dataframe: pd.DataFrame


@dataclass
class QuestionMap:
    dataframe: pd.DataFrame


# =====================================================
# Domain Models (Validated & Optimized Layer)
# =====================================================

@dataclass(frozen=True)
class Question:
    identifier: str
    max_marks: float
    co_list: Tuple[str, ...]


@dataclass(frozen=True)
class ComponentInfo:
    weight: float
    cia: bool
    co_split: bool
    direct: bool
    questions: Tuple[Question, ...]


@dataclass(frozen=True)
class Student:
    reg_no: str
    name: str

@dataclass(frozen=True)
class IndirectToolInfo:
    name: str
    weight: float  # e.g., 0.1 for 10%

@dataclass(frozen=True)
class ValidatedSetup:
    components: Dict[str, ComponentInfo]
    students: Tuple[Student, ...]
    indirect_tools: Tuple[IndirectToolInfo, ...]
