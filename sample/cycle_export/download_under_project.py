# coding: utf-8
""" Copyright (c) 2020 VeriServe Corporation """

import os
import io
import csv
import re

import json
import requests
from urllib.request import urlretrieve
from urllib.error import URLError,HTTPError
import pandas as pd
from pyexcelerate import Workbook

import time

#【定数】
BASE_API_URL = 'https://cloud.veriserve.co.jp/'
SLEEP_TIME = 1.0


def strip_quotes(src):
    ret = src.strip()
    ret = ret.strip('\'')
    ret = ret.strip('\"')
    return ret

def get_request_pages(mid_url, content_name, response_key = None):
    ret = None
    url = BASE_API_URL + 'api/v2/' + mid_url + '?api_key=' + _api_key
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


def download_csv(target_url,file_name):
    url = target_url + '?api_key=' + _api_key

    file_path = ensure_no_kinsoku_chars(file_name + '.csv')
    try:
        ret = urlretrieve(url,"{0}".format(file_path))
        time.sleep(SLEEP_TIME)  #【スリープ】
    except HTTPError as e:
        print('The server couldn\'t fulfill the request. Error code: ', e.code)
        exit()
    except URLError as e:
        print('We failed to reach a server. Reason: ', e.reason)
        exit()
    else:
        with io.open(ret[0],'r',encoding='utf_8_sig') as f:
            reader = csv.reader(f)
            download_data = [row for row in reader]
            f.close()
            os.remove(ret[0])
        return download_data

def save_vals(merged_data,output_path):
    df = pd.DataFrame(merged_data)
    wb = Workbook()
    wb.new_sheet("Sheet1", data=df.values.tolist())
    wb.save(output_path)

# 禁則文字を_に置き換える
def ensure_no_kinsoku_chars(filepath):
    return re.sub(r'[\\|/|:|?|"|<|>|\| |]', '_', filepath)

def download_under_project(test_phases):
    current_path = os.getcwd()
    for phase in test_phases:
        #フォルダ名は「テストフェーズ名」とする
        output_dir = current_path + '/' + ensure_no_kinsoku_chars(phase['name'])+ '/'
        os.makedirs(os.path.dirname(output_dir), exist_ok=True)    # Ensure exist dir
    
        for tsa in phase['test_suite_assignments']:
            mid_url = 'test_phases/' + str(phase['id']) + '/test_suite_assignments/' + str(tsa['id']) + '/test_cycles'
            test_cycles = get_request_pages(mid_url, 'test_cycles')
            if test_cycles is None:
                print('"test_cycles"の取得に失敗しました。')
                exit()
            for cycle in test_cycles:
                target_url = BASE_API_URL + 'api/v2/' + mid_url + '/' + str(cycle['id']) + '.csv'
                download_data = download_csv(target_url,str(cycle['id']))
                if download_data is None:
                    print(target_url[2] + "のCSVデータ取得に失敗しました。")
                    exit()
                #出力ファイル名は「テストサイクル名」とする
                output_path = output_dir + ensure_no_kinsoku_chars(cycle['name']) +'.xlsx' 
                print("output_path: ",output_path,)
                save_vals(download_data,output_path)

if __name__ == '__main__':
    # コマンドラインオプション処理
    import argparse
    ap = argparse.ArgumentParser(description='プロジェクト下の全テストフェーズの設定画面のURLを取得する。')
    ap.add_argument("-a", "--api_key", action='store', help="APIキー", required=True)
    args = ap.parse_args()
    _api_key = strip_quotes(args.api_key)

    test_phases = get_request_pages('test_phases', 'test_phases')
    if test_phases is None:
        print('APIキーが間違っています。')
        exit()

    # メイン処理
    download_under_project(test_phases)

    print('Done')