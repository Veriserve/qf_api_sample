
# coding: utf-8
""" Copyright (c) 2020 VeriServe Corporation """
try:
    import json
    import requests
    import datetime
    import os
    import settings
    import time
    import logging
except ModuleNotFoundError as err:
    print('「' + err.name + '」のモジュールが見つかりません。以下のコマンドラインでインストールしてください。')
    print('>> pip install ' + err.name)
    exit()

#QFサービス規約対応のため
SLEEP_TIME = 1
#ログの設定
logging.basicConfig(filename='app_{0}.log'.format(datetime.datetime.now().strftime('%Y%m%d')), 
                    filemode='a', 
                    format='%(asctime)s - %(levelname)s - %(message)s', 
                    level=logging.INFO, datefmt='%Y-%m-%d %H:%M:%S')
# -----------------------------------------------------------
# 文字列から「'」と「"」の文字を削除する
#
# Parameters:
#  src（str）：対象文字列
#
# Returns:
#  ret（str）：「'」と「"」の文字を削除した文字列
# -----------------------------------------------------------
def strip_quotes(src):
    if src is None:
        return None
    ret = src.strip()
    #「'」を削除する
    ret = ret.strip('\'')
     #「"」を削除する
    ret = ret.strip('\"')
    return ret

# -----------------------------------------------------------
# APIのURLを作成する
#
# Parameters:
#  mid_url（str）：APIのURLの中間文字列
#
# Returns:
#  APIのURL
# -----------------------------------------------------------
def build_url_api(mid_url):
    if settings.BASE_API_URL.endswith('/'):
        return settings.BASE_API_URL + mid_url + '?api_key=' + settings.API_KEY
    else:
        return settings.BASE_API_URL + '/' + mid_url + '?api_key=' + settings.API_KEY

# -----------------------------------------------------------
# APIから情報を取得する
#
# Parameters:
#  mid_url（str）：APIのURLの中間文字列
#  content_name（str）：コンテンツのキー名
#
# Returns:
#  ret（array）：APIから応答された結果のリスト
# -----------------------------------------------------------
def get_request_pages(mid_url, content_name):
    ret = None
    url = build_url_api(mid_url)
    while True:
        #APIにリクエストを送信する
        response = requests.get(url)
        time.sleep(SLEEP_TIME)
        #応答結果の確認
        if response.status_code != 200:
            return None
        content = json.loads(response.content)
        #リストに結果を格納する
        if ret is None:
            ret = content[content_name]
        else:
            ret.extend(content[content_name])

        #次のページがあるかどうかをチェックする
        if 'next_url' not in content or content['next_url'] is None or content['total_pages'] == 0:
            break
        url = content['next_url']
    return ret

# -----------------------------------------------------------
# APIからプロジェクトの情報を取得する
#
# Returns:
#  project_id（int）：プロジェクトのid
#  default_label_content（dict）：ラベルコンテンツのデファルト値の一覧
# -----------------------------------------------------------
def get_project_info():
    project_id = None
    url = build_url_api('current_project')

    #APIにリクエストを送信する
    response = requests.get(url)
    time.sleep(SLEEP_TIME)
    #応答結果の確認
    if response.status_code != 200:
        return None
    content = json.loads(response.content)
    #リストに結果を格納する
    project_id = content['id']

    return project_id

# -----------------------------------------------------------
# APIに登録リクエストを送信する
#
# Parameters:
#  mid_url（str）：APIのURLの中間文字列
#  payload（str）：APIに送信されたデータ
#  headers（dict）：ヘッダの設定
#
# Returns:
#  content（object）：APIに登録されたレコード
#
# Exception：リクエストの情報で例外をスローする
# -----------------------------------------------------------
def post_request(mid_url, payload, headers):
    url = build_url_api(mid_url)
    #データのエンコード
    payload = payload.encode('utf-8')
    response = requests.post(url, data=payload, headers=headers)
    time.sleep(SLEEP_TIME)
    #応答結果の確認
    if response.status_code == 201:
        content = json.loads(response.content)
    else:
        raise Exception("requests.post エラー！: status_code: {}, URL: {}, payload: {}, headers: {}".format(response.status_code, url, payload, headers))
    return content

# -----------------------------------------------------------
# APIに更新リクエストを送信する
#
# Parameters:
#  mid_url（str）：APIのURLの中間文字列
#  payload（str）：APIに送信されたデータ
#  headers（dict）：ヘッダの設定
#
# Returns:
#  content（object）：APIに更新されたレコード
#
# Exception：リクエストの情報で例外をスローする
# -----------------------------------------------------------
def update_request(mid_url, payload, headers):
    url = build_url_api(mid_url)
    #データのエンコード
    payload = payload.encode('utf-8')
    response = requests.patch(url, data=payload, headers=headers)
    time.sleep(SLEEP_TIME)
    #応答結果の確認
    if response.status_code == 200:
        return json.loads(response.content)
    else:
        raise Exception("requests.patch エラー！: status_code: {}, URL: {}, payload: {}, headers: {}".format(response.status_code, url, payload, headers))

# -----------------------------------------------------------
# APIに削除リクエストを送信する
#
# Parameters:
#  mid_url（str）：APIのURLの中間文字列
#  headers（dict）：ヘッダの設定
#
# Returns:
#  True（bool）：削除が成功した
#
# Exception：リクエストの情報で例外をスローする
# -----------------------------------------------------------
def delete_request(mid_url, headers):
    url = build_url_api(mid_url)
    response = requests.delete(url, headers=headers)
    time.sleep(SLEEP_TIME)
    if response.status_code == 204:
        return True
    else:
        raise Exception("requests.delete エラー！: status_code: {}, URL: {}, headers: {}".format(response.status_code, url, headers))

# -----------------------------------------------------------
# フォルダ内のすべてのファイルをロードする
#
# Returns:
#  folders（dict）：該当フォルダーのパスとファイル名のリスト
# -----------------------------------------------------------
def load_files_in_folder():
    try:
        folders = os.walk(settings.FOLDER_PATH)
    except FileNotFoundError:
        return None
    return folders

# -----------------------------------------------------------
# 同ラベル名が存在するテストスイートの項目Nをテストブロッカー列に指定する
#
# Parameters:
#  col_block_index（int）：同ラベルインデックス
#  test_suite_id（int）：テストスイート番号
#
# Returns:
#  True（bool）：更新が成功した場合
#  False（bool）：更新に失敗した場合
# -----------------------------------------------------------
def update_test_block(col_block_index, test_suite_id):
    mid_url = 'test_suites/' + str(test_suite_id)
    headers = {'content-type':'application/x-www-form-urlencoded'}
    payload = 'test_suite[test_blocker_column]=' + str(col_block_index)

    # テストスイートを更新する
    if update_request(mid_url, payload, headers):
        return True
    else:
        #更新に失敗した時に処理を中断する
        print('テストブロッカー列に指定に失敗しました。')
        exit()

# -----------------------------------------------------------
# マイン関数
# -----------------------------------------------------------
if __name__ == '__main__':
    global _target_index

    # コマンドラインオプション処理
    import argparse
    ap = argparse.ArgumentParser(description='Quality Forwardのテスト結果文字列を置換する。')
    ap.add_argument("-n", "--label_name", action='store', help="ラベル名を指定する", required=True)
    args = ap.parse_args()
    _label_name = strip_quotes(args.label_name)

    #project_idを取得する
    project_id = get_project_info()
    if project_id is None:
        print('APIキーが不正です。')
        exit()

    #QFから全てのテストスイートを取得する
    _test_suite_list = get_request_pages('test_suites', 'test_suites')

    for ts in _test_suite_list:
        if ts['project_id'] != project_id:
            continue
        isTarget = False
        #QFから全てのテストスイートバージョンを取得する
        _test_suite_vs_list = get_request_pages('test_suites/' + str(ts['id']) + '/test_suite_versions', 'test_suite_versions')
        #「利用可」のテストスイートバージョンが存在する

        for tsv in _test_suite_vs_list:
            if tsv['status'] == settings.TSV_STATUS:
                isTarget = True
                break
        if isTarget == False:
            continue

        _target_index = 0

        keys = ts.keys()
        label_content_list = list(filter(lambda x: x.find('label_content') != -1, keys))

        #同ラベル名が存在するテストスイートの項目Nを見つける
        for label in label_content_list:
            if ts[label] == _label_name:
                _target_index = int(label.replace('label_content', ''))
                break

        if _target_index > 0:
            update_test_block(_target_index, ts['id'])
            logging.info('設定を有効にしたテストスイート名：' + ts['name'])
            print('設定を有効にしたテストスイート名：' + ts['name'])

    print('テストブロッカー指定が完了しました。')

    exit()