# coding: utf-8
""" Copyright (c) 2020 VeriServe Corporation """

import os
import re
import json
import requests
import csv
import time

#【定数】
base_api_url = 'https://cloud.veriserve.co.jp/' # production
SUFFIX_OF_OUT_FILE = '_test_cycle_urls.csv'
SLEEP_TIME = 0.4

def strip_quotes(src):
    ret = src.strip()
    ret = ret.strip('\'')
    ret = ret.strip('\"')
    return ret

def get_request_pages(mid_url, content_name, response_key = None):
    ret = None
    url = base_api_url + 'api/v2/' + mid_url + '?api_key=' + _api_key
    while True:
        response = requests.get(url)
        time.sleep(SLEEP_TIME)
        if response.status_code != 200:
            return None
        content = json.loads(response.content)
        if response_key is None:
            if ret is None:
                ret = content[content_name]
            else:
                ret.extend(content[content_name])
        else:
            if ret is None:
                ret = []
            ret.extend(content[response_key])
        if 'next_url' not in content or content['next_url'] is None or content['total_pages'] == 0:
            break
        url = content['next_url']
    return ret

def get_test_cycles(test_suite_assignments):
    cycles_dict = {}
    for tsa in test_suite_assignments:
        mid_url = 'test_phases/' + test_phase_id + '/test_suite_assignments/' + str(tsa['id']) + '/test_cycles'
        test_cycles = get_request_pages(mid_url, 'test_cycles')
        if test_cycles is None:
            print('"test_cycles"の取得に失敗しました。')
            exit()
        for cycle in test_cycles:
            cycles_dict[cycle['name']] = [str(tsa['id']),tsa['test_suite_name'],str(cycle['id']),cycle['name']]
    # テストサイクル名でソートしないといけない
    sorted_list = sorted(cycles_dict.items(), key=lambda x:x[0])
    test_cycle_urls = []
    for l in sorted_list:
        tp_url = base_api_url + 'projects/' + project_id + '/test_phases/' + test_phase_id + '/test_suite_assignments/' + l[1][0] + '/test_cycles'
        tp_name = test_phase_name + 'の' + l[1][1]
        tc_id = l[1][2]
        #tc_id = '/' + l[1][2] + '#'
        dl_url = base_api_url + 'api/v2/test_phases/' + test_phase_id + '/test_suite_assignments/' + l[1][0] + '/test_cycles/' + l[1][2] + '.csv'
        test_cycle_urls.append([tp_url,tp_name,l[1][3],tc_id,dl_url])
    return test_cycle_urls

# 禁則文字を_に置き換える
def ensure_no_kinsoku_chars(filepath):
    return re.sub(r'[\\|/|:|?|"|<|>|\|]', '_', filepath)

if __name__ == '__main__':
    start = time.time()
    # コマンドラインオプション処理
    import argparse
    ap = argparse.ArgumentParser(description='プロジェクト下の全テストフェーズの設定画面のURLを取得する。')
    ap.add_argument("-a", "--api_key", action='store', help="APIキー", required=True)
    ap.add_argument("-o", "--output_dir", action='store', help="txtファイルを出力するディレクトリ。指定しなければカレントディレクトリ", default="", required=False)
    args = ap.parse_args()
    _api_key = strip_quotes(args.api_key)
    _output_dir = strip_quotes(args.output_dir)

    test_phases = get_request_pages('test_phases', 'test_phases')
    if test_phases is None:
        print('APIキーが間違っています。')
        exit()
    for phase in test_phases:
        project_id = str(phase['project_id'])
        test_phase_id = str(phase['id'])
        test_phase_name = phase['name']
        test_cycle_urls = get_test_cycles(phase['test_suite_assignments'])
        output_path = os.path.join(_output_dir,ensure_no_kinsoku_chars(phase['name']) + '_' + str(phase['id']) + SUFFIX_OF_OUT_FILE)
        print(output_path + " へ書き出し中...")
        os.makedirs(os.path.dirname(output_path), exist_ok=True)    # Ensure exist dir
        with open(output_path, 'w') as f:
            writer = csv.writer(f, lineterminator='\n') # 改行コード（\n）を指定しておく
            for url in test_cycle_urls:
                writer.writerow([ s.encode('cp932', 'ignore').decode("CP932") if (type(s) is str) else s for s in url])

    elapsed_time = time.time() - start
    print ("elapsed_time:{0}".format(elapsed_time) + "[sec]")

