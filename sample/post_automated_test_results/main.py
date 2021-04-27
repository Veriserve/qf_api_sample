import sys

from requests.packages import urllib3
from requests.packages.urllib3.exceptions import InsecureRequestWarning

from modules import api, result

if __name__ == "__main__":
    urllib3.disable_warnings(InsecureRequestWarning)

    try:
        qf_api = api.API(config_path="config.ini", config_section="QF_API")

        test_cycle_id = qf_api.create_test_cycle()
        results = result.parse_xml("results/result.xml")
        qf_api.post_test_results(results=results, test_cycle_id=test_cycle_id)

    except Exception as e:
        print(e)
        sys.exit(1)
