import re

from junitparser import JUnitXml
from junitparser.junitparser import Failure, TestCase


class Result:
    def __init__(
        self, name: str, result: str, file_path: str, time: float, message: str = ""
    ) -> None:
        _QF_RESULT_PASS = 1
        _QF_RESULT_FAIL = 2
        _QF_RESULT_SKIP = 3
        _QF_RESULT_NA = 6

        if result == "pass":
            self.result = _QF_RESULT_PASS
        elif result == "failure" or result == "error":
            self.result = _QF_RESULT_FAIL
        elif result == "skipped":
            self.result = _QF_RESULT_SKIP
        else:
            self.result = _QF_RESULT_NA

        self.name = name
        self.file_path = file_path
        self.time = time
        self.message = message

    def to_dict(self) -> dict:
        return {
            "test_method_name": self.name,
            "result": self.result,
            "message": self.message,
            "file_path": self.file_path,
            "time": self.time,
        }


def _get_file_path(case: TestCase) -> str:
    if "file=" in str(case.tostring()):
        return re.sub(
            'file=|"',
            "",
            re.search(r'file=".*?"', str(case.tostring())).group(0),
        )
    else:
        return ""


def parse_xml(xml_path: str) -> list[dict]:
    results = []

    for suite in JUnitXml.fromfile(xml_path):
        for case in suite:
            if len(list(case)) == 0:
                # Passed Cases
                result = Result(
                    name=case.name,
                    result="pass",
                    file_path=_get_file_path(case),
                    time=case.time,
                )
            else:
                # Failure, Error, Skipped Cases
                for detail in case:  # the length of the case must be 1.
                    result = Result(
                        name=case.name,
                        result=detail._tag,
                        file_path=_get_file_path(case),
                        time=case.time,
                        message=detail.message,
                    )
            results.append(result.to_dict())

    return results
