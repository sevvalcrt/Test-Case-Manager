from dataclasses import dataclass
from typing import List


@dataclass
class TestStep:
    step_no: int
    action: str
    test_data: str
    expected_result: str


@dataclass
class TestData:
    key: str
    value: str


@dataclass
class TestCase:

    tc_id: str
    name: str

    priority: str
    application: str

    version: int

    creator: str

    create_date: str

    automation: bool

    pre_conditions: List[str]

    test_data: List[TestData]

    steps: List[TestStep]

    post_conditions: List[str]