import configparser
import json
import time
from datetime import datetime

import requests


class API:
    def __init__(self, config_path: str, config_section: str) -> None:
        self._STATUS_OK: int = 200
        self._STATUS_CREATED: int = 201
        self._CASE_NOT_FOUND: int = -1

        config = configparser.ConfigParser()
        config.read(config_path, encoding="utf-8")
        config = config[config_section]

        self.base_url: str = config["base_url"]
        self.api_key: str = config["api_key"]
        self.user_id: int = int(config["user_id"])
        self.test_suites_id: int = int(config["test_suites_id"])
        self.test_suites_version_id: int = int(config["test_suites_version_id"])
        self.test_phase_id: int = int(config["test_phase_id"])
        self.test_suite_assignment_id: int = int(config["test_suite_assignment_id"])

    def get_users(self) -> dict:
        qf_api_url = self.base_url + "users" + "?api_key=" + self.api_key

        res = requests.get(qf_api_url, verify=False)
        time.sleep(1)
        if res.status_code == 200:
            return json.loads(res.text)
        else:
            raise Exception("ユーザ取得に失敗しました")

    def create_test_cycle(self) -> int:
        qf_api_url = (
            self.base_url
            + "test_phases/"
            + str(self.test_phase_id)
            + "/test_suite_assignments/"
            + str(self.test_suite_assignment_id)
            + "/test_cycles"
            + "?api_key="
            + self.api_key
        )

        now = datetime.now()

        data_cycle = {
            "test_cycle[name]": now.strftime("自動テストサイクル %Y%m%d %H-%M-%S"),
            "test_cycle[target_priorities][]": "A",
            "test_cycle[start_on]": now,
            "test_cycle[end_on]": now,
            "test_cycle[status]": "waiting_for_review",
        }

        res = requests.post(
            qf_api_url,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            data=data_cycle,
            verify=False,
        )
        time.sleep(1)

        if res.status_code == self._STATUS_CREATED:
            print("テストサイクルを作成しました: " + data_cycle["test_cycle[name]"])
            return json.loads(res.text)["id"]
        else:
            raise Exception("テストサイクル作成に失敗しました")

    def get_test_cases(self) -> dict:
        qf_api_url = (
            self.base_url
            + "test_suites/"
            + str(self.test_suites_id)
            + "/test_suite_versions/"
            + str(self.test_suites_version_id)
            + "/test_cases"
            + "?api_key="
            + self.api_key
        )

        res = requests.get(qf_api_url, verify=False)
        time.sleep(1)

        if res.status_code == self._STATUS_OK:
            return json.loads(res.text)["test_cases"]
        else:
            raise Exception("テストケース取得に失敗しました")

    def get_test_case_no_from_id(self, test_cases: dict, test_id: str) -> int:
        for case in test_cases:
            if test_id in case.values():
                return case["no"]
        else:
            return self._CASE_NOT_FOUND

    def post_test_results(self, results: dict, test_cycle_id: int) -> None:
        qf_api_url = (
            self.base_url
            + "test_phases/"
            + str(self.test_phase_id)
            + "/test_suite_assignments/"
            + str(self.test_suite_assignment_id)
            + "/test_cycles/"
            + str(test_cycle_id)
            + "/test_results?api_key="
            + self.api_key
        )

        test_cases = self.get_test_cases()

        for i, result in enumerate(results):
            test_case_no = self.get_test_case_no_from_id(
                test_cases=test_cases,
                test_id=result["test_method_name"],
            )

            if test_case_no != self._CASE_NOT_FOUND:
                data = {
                    "test_result[test_case_no]": test_case_no,
                    "test_result[result]": result["result"],
                    "test_result[user_id]": self.user_id,
                    "test_result[executed_at]": datetime.now(),
                    "test_result[content1]": result["file_path"],
                    "test_result[content2]": result["time"],
                    "test_result[content3]": result["message"],
                }

                res = requests.post(
                    qf_api_url,
                    headers={"Content-Type": "application/x-www-form-urlencoded"},
                    data=data,
                    verify=False,
                )
                time.sleep(1)

                if res.status_code == self._STATUS_CREATED:
                    print("テスト結果投入完了 : " + str(i + 1) + " / " + str(len(results)))
                if res.status_code != self._STATUS_CREATED:
                    raise Exception("テスト結果作成に失敗しました")

        print("Done")
