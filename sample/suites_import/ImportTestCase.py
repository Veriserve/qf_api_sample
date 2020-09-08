
# coding: utf-8
""" Copyright (c) 2020 VeriServe Corporation """
try:
    import json
    import requests
    import datetime
    import os
    import xlrd
    import settings
    import urllib.parse
    import time
    from openpyxl import load_workbook
    from format import format_value
except ModuleNotFoundError as err:
    print('「' + err.name + '」のモジュールが見つかりません。以下のコマンドラインでインストールしてください。')
    print('>> pip install ' + err.name)
    exit()

#【定数】
NOW = datetime.datetime.now().strftime('%Y-%m-%d')
EXCEL_EXTENSION = '.xlsx'
#QFサービス規約対応のため
SLEEP_TIME = 1

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
# 名前でテストスイートの存在をチェックします
#
# Parameters:
#  test_suite_name（str）：テストスイート名
#
# Returns:
#  test_suite_id（int）：テストスイート番号
# -----------------------------------------------------------
def check_exist_test_suite(test_suite_name):
    test_suite_id = None
    #テストスイートが存在しない場合
    if _test_suite_list is None:
        return None
    #テストスイートが存在する場合
    for x in _test_suite_list:
        if x['name'] == test_suite_name:
            test_suite_id = x['id']

    return test_suite_id

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
# エクセルファイルのデータを全てロードする
#
# Parameters:
#  file_path（str）：エクセルファイルのパス
#
# Returns:
#  ws_rows_list（array）：各シートのデータのリスト
#  sheet_name_list（array）：対象シート名のリスト
#
# Exception：ファイルが見つかった時に例外をスローする
# -----------------------------------------------------------
def load_excel(file_path, is_correct_data):
    #各シートのデータのリスト
    ws_rows_list = []
    #対象シート名のリスト
    sheet_name_list = []
    try:
        #データ抽出用のスプレッドシートファイルを開きます
        wb = xlrd.open_workbook(filename=file_path)

        for ws in wb._sheet_list:
            #表示されているシートのみを取得します
            if ws.visibility == 0:
                if is_correct_data == False:
                    ws_rows_list.append(list(ws._cell_values))
                sheet_name_list.append(ws.name)

        if is_correct_data:
            wb = load_workbook(filename=file_path, read_only=True, data_only=True, keep_links=False)
            for s in sheet_name_list:
                ws_rows_list.append(list(wb[s].rows))
    except FileNotFoundError:
        print('excelインポートするファイルが見つかりません。')
        return None

    return ws_rows_list, sheet_name_list

# -----------------------------------------------------------
# テストスイートを番号で削除する
#
# Parameters:
#  test_suite_id（int）：テストスイート番号
#
# Returns:
#  True（bool）：削除が成功した
#  ※削除に失敗した時に処理を中断する
# -----------------------------------------------------------
def delete_test_suite(test_suite_id):
    mid_url = 'test_suites/' + str(test_suite_id)
    headers = {'Content-Type' : 'application/x-www-form-urlencoded'}
    if delete_request(mid_url, headers) == True:
        return True
    else:
        print('テストスイート削除に失敗しました。')
        exit()

# -----------------------------------------------------------
# 新規テストスイートを作成する
#
# Parameters:
#  test_suite_name（str）：テストスイート名
#  ws_rows（array）：対象ヘッダデータのリスト
#
# Returns:
#  test_suite_id（int）：登録されたテストスイート番号
#  ※作成に失敗した時に処理を中断する
# -----------------------------------------------------------
def create_new_suite(test_suite_name, ws_rows):
    #テストスイートの存在を確認する
    test_suite_id = check_exist_test_suite(test_suite_name)
    if test_suite_id != None:
        if settings.TEST_SUITE_DELETE_FLG:
            #存在する場合は古いテストスイートを削除する
            delete_test_suite(test_suite_id)
        else:
            print('以下のテストスイートが存在しています。')
            print('  ▪ ' + test_suite_name)
            exit()

    mid_url = 'test_suites'
    headers = {'content-type':'application/x-www-form-urlencoded'}
    payload = 'test_suite[project_id]=' + str(project_id) + '&test_suite[name]=' + test_suite_name

    #ヘッダデータの設定
    j = _col_start_index    #読み取りを開始する列を指定します
    no = 1                      #ラベルカテゴリのインデックス

    #各列のデータを可能範囲で読み取る
    while j <= _col_end_index:
        val = ws_rows[_header_row_index][j].value
        if val is not None and str(val).strip() != '' and str(val).strip() != settings.COL_TITLE_START:
            #空のセルではない場合
            payload = payload + "&test_suite[label_category" + str(no) + "]=" + str(val)
            payload = payload + "&test_suite[use_category" + str(no) + "]=true"
            no += 1
        j += 1

    #作成時間を設定する
    payload = payload + '&test_suite[created_at]=' + NOW
    
    # 新規テストスイートを作成する
    content = post_request(mid_url, payload, headers)

    if content != None:
        #応答結果からテストスイート番号を取得する
        test_suite_id = content['id']
    else:
        #作成に失敗した場合に処理を中断する
        print('テストスイートの作成に失敗しました。')
        exit()
    return test_suite_id

# -----------------------------------------------------------
# 新規テストスイートバージョンを作成する
#
# Parameters:
#  test_suite_id（int）：テストスイート番号
#
# Returns:
#  test_suite_vs_id（int）：登録されたテストスイートバージョン番号
#  ※作成に失敗した時に処理を中断する
# -----------------------------------------------------------
def create_new_tsv(test_suite_id):
    mid_url = 'test_suites/' + str(test_suite_id) + '/test_suite_versions'
    headers = {'content-type':'application/x-www-form-urlencoded'}
    payload = 'test_suite_version[project_id]=' + str(project_id) + '&test_suite_version[name]=' + settings.TSV_NAME + '&test_suite_version[created_at]=' + NOW

    content = post_request(mid_url, payload, headers)
    if content != None:
        #応答結果からテストスイートバージョン番号を取得する
        test_suite_vs_id = content['id']
    else:
        #作成に失敗した場合に処理を中断する
        print('テストスイートバージョンの作成に失敗しました。')
        exit()
    return test_suite_vs_id

# -----------------------------------------------------------
# 新規テストケースを作成する
#
# Parameters:
#  test_suite_id（int）：テストスイート番号
#  test_suite_version_id（int）：テストスイートバージョン番号
#  data_row（array）：対象データのリスト
#
# Returns:
#  True（bool）：テストケースが成功した場合
#  False（bool）：テストケースの作成に失敗した場合
# -----------------------------------------------------------
def create_new_tc(test_suite_id,test_suite_version_id,data_row):
    mid_url = 'test_suites/' + str(test_suite_id) + '/test_suite_versions/' + str(test_suite_version_id) + '/test_cases'
    headers = {'content-type':'application/x-www-form-urlencoded'}
    payload = 'test_case[no]=' + str(tc_no)
    #優先度
    if data_row[_col_priority_index].value is not None:
        payload = payload + '&test_case[priority]=' + data_row[_col_priority_index].value

    #テストケースのデータの設定
    i = _col_start_index    #読み取りを開始する列を指定します
    no = 1                      #カテゴリのインデックス

    #各列のデータを可能範囲で読み取る
    while i <= _col_end_index:
        if i != _col_priority_index:
            if data_row[i].value is not None:
                if data_row[i].data_type == 's':
                    val = str(data_row[i].value)
                else:
                    val = format_value(data_row[i].value, data_row[i].data_type, data_row[i].number_format)
                payload = payload + '&test_case[category' + str(no) + ']=' + urllib.parse.quote(val)
            no += 1
        i +=1
    #作成時間を設定する
    payload = payload + '&test_case[created_at]=' + NOW
    
    #テストケースを作成する
    content = post_request(mid_url, payload, headers)
    if content != None:
        return True
    else:
        return False

# -----------------------------------------------------------
# テストスイートバージョンのステータスを更新する
#
# Parameters:
#  test_suite_id（int）：テストスイート番号
#  test_suite_version_id（int）：テストスイートバージョン番号
#
# Returns:
#  True（bool）：更新が成功した
#  ※更新に失敗した時に処理を中断する
# -----------------------------------------------------------
def update_test_suite_version(test_suite_id,test_suite_version_id):
    mid_url = 'test_suites/' + str(test_suite_id) + '/test_suite_versions/' + str(test_suite_version_id)
    headers = {'content-type':'application/x-www-form-urlencoded'}

    #「利用可」にステータスを設定する
    payload = 'test_suite_version[status]=' + settings.TSV_STATUS
    if update_request(mid_url, payload, headers):
        return True
    else:
        #更新に失敗した時に処理を中断する
        print('テストスイート更新に失敗しました。')
        exit()

# -----------------------------------------------------------
# 重複ファイル名をチェックする
#
# Returns:
#  重複ファイル名がある場合に処理を中断する
#  重複ファイル名一覧を表示する
# -----------------------------------------------------------
def check_duplicate_file():
    target_files = []
    duplicate_files = []
     #フォルダーのパスとファイルの一覧を取得する
    folders = load_files_in_folder()

    #フォルダ内のすべてのファイルを取得する
    for path, subdirs, files in folders:
        for f in files:
            target_files.append(f)

    #各ファイルをループして重複をチェックします
    for f in target_files:
        if target_files.count(f) > 1 and duplicate_files.count(f) == 0:
            #リストに重複ファイル名を格納する
            duplicate_files.append(f)

    #重複ファイル名がある場合に処理を中断する
    if len(duplicate_files) > 0:
        print('=====以下の重複ファイル名をご修正下さい。=====')

        #重複ファイル名一覧を表示する
        for f in duplicate_files:
            print(f)
        exit()

# -----------------------------------------------------------
# ファイルのフォーマットをチェックする
#
# Returns:
#  ひとつでも読み込み不可のExcelシートがある場合は処理を中断し、エラーメッセージを表示する
#  列の定義が設定可能範囲より多い時
#  優先度の列が定義されていない
# -----------------------------------------------------------
def check_format_files():
    no_priority_files = []  #「優先度」無しファイルの一覧
    max_column_files = []   # QFで取り込める列の最大より大きいファイルの一覧
    test_suites_exist = []  # 存在しているテストスイートの一覧
     #フォルダーのパスとファイルの一覧を取得する
    folders = load_files_in_folder()

    for path, subdirs, files in folders:
        for f in files:
            #Excel以外のファイル拡張子
            if EXCEL_EXTENSION not in f:
                continue

            #エクセルファイルのデータを全てロードする
            ws_rows_list, _sheet_name_list = load_excel(path + '\\' + f, False)
            sheet_index = 0
            #各シートをループする
            for ws_rows in ws_rows_list:
                _header_row_index = 0
                
                #対象シートのフラグ
                isTarget = False

                #列インデックスの開始とヘッダー行インデックスを見つける
                #最初の10列と10行だけを見つける
                for r in ws_rows:
                    #列の開始インデックスをリセットする
                    _col_priority_index = 0
                    for c in r:
                        #最初の10列が終わったら、停止します
                        if _col_priority_index > 20:
                            break
                        #「優先度」の列を見つけた場合
                        if str(c).strip() == settings.COL_TITLE_START:
                            #対象シート
                            isTarget = True
                            break
                        _col_priority_index += 1

                    #対象シートだ、または最初の10行が終わったら、停止します
                    if isTarget or _header_row_index > 20:
                        break
                    _header_row_index += 1

                #シート対象ではない場合、処理を中断する
                if isTarget == False:
                    no_priority_files.append('  ▪ ' + path + '\\' + f)
                
                #列の定義が設定可能範囲より多いかどうかチェックする
                _col_start_index = _col_priority_index + 1
                _col_end_index = -1
                _col_import_num = 0
                i = _col_start_index

                #列インポートの数を数える
                while i < len(ws_rows[_header_row_index]):
                    #セルが空の場合は、停止します
                    if _col_start_index != -1 and str(ws_rows[_header_row_index][i]).strip() == '':
                        _col_end_index = i - 1
                        break
                    elif i == len(ws_rows[_header_row_index]) - 1:
                        _col_end_index = i
                    i += 1
                _col_import_num = _col_end_index - _col_start_index + 1

                #列の定義が設定可能範囲より多い時に処理を中断し、エラーメッセージを表示する
                if _col_import_num >= settings.QF_COLUMN_MAX:
                    max_column_files.append('  ▪ ' + path + '\\' + f)

                #存在しているテストスイートがあるかどうかチェックする
                if settings.TEST_SUITE_DELETE_FLG == 0:
                    test_suite_name = f[0:f.rindex('.')] + '-' + _sheet_name_list[sheet_index]
                    test_suite_id = check_exist_test_suite(test_suite_name)
                    #存在の場合
                    if test_suite_id != None:
                        test_suites_exist.append('  ▪ ' + test_suite_name)
                #シートの次
                sheet_index += 1
    isExit = False
    #優先度の列が定義されていない場合にメセージエラーを表示する
    if len(no_priority_files) > 0:
        isExit = True
        print('■優先度の列が定義されていません。')
        for x in no_priority_files:
            print(x)
    #QFで取り込める列の最大より大きい場合にメセージエラーを表示する
    if len(max_column_files) > 0:
        isExit = True
        print('\n■QFで取り込める列の最大より大きい。')
        for x in max_column_files:
            print(x)
    #存在テストスイートがある場合にメセージエラーを表示する
    if len(test_suites_exist) > 0:
        isExit = True
        print('\n■以下のテストスイートが存在しています。')
        for x in test_suites_exist:
            print(x)

    #処理が終了する
    if isExit:
        exit()

# -----------------------------------------------------------
# マイン関数
# -----------------------------------------------------------
if __name__ == '__main__':
    global _header_row_index
    global _col_start_index
    global _col_end_index
    global _col_priority_index

    #project_idを取得する
    project_id = get_request_pages('current_project','id')
    if project_id is None:
        print('APIキーが不正です。')
        exit()

    #QFから全てのテストスイートを取得する
    _test_suite_list = get_request_pages('test_suites', 'test_suites')

    #重複ファイル名をチェックする
    check_duplicate_file()
    #ひとつでも読み込み不可のExcelシートがある場合は処理を中断し、エラーメッセージを表示する
    check_format_files()
    #フォルダーのパスとファイルの一覧を取得する
    folders = load_files_in_folder()

    #対象ファイルがない場合
    if folders == None:
        print('対象ファイルが見つかりません。')
        exit()

    #各ファイルをループする
    for path, subdirs, files in folders:
        for f in files:
            #Excel以外のファイル拡張子
            if EXCEL_EXTENSION not in f:
                continue

            #エクセルファイルのデータを全てロードする
            ws_rows_list, _sheet_name_list = load_excel(path + '\\' + f, True)
            sheet_index = 0

            #各シートをループする
            for ws_rows in ws_rows_list:
                _header_row_index = 0
                
                #対象シートのフラグ
                isTarget = False

                #列インデックスの開始とヘッダー行インデックスを見つける
                #最初の10列と10行だけを見つける
                for r in ws_rows:
                    #列の開始インデックスをリセットする
                    _col_priority_index = 0
                    for c in r:
                        #最初の10列が終わったら、停止します
                        if _col_priority_index > 20:
                            break
                        #「優先度」の列を見つけた場合
                        if str(c.value).strip() == settings.COL_TITLE_START:
                            #対象シート
                            isTarget = True
                            break
                        _col_priority_index += 1

                    #対象シートだ、または最初の10行が終わったら、停止します
                    if isTarget or _header_row_index > 20:
                        break
                    _header_row_index += 1
                
                #列の定義が設定可能範囲より多いかどうかチェックする
                _col_start_index = _col_priority_index + 1
                _col_end_index = -1
                i = _col_start_index

                #列インポートの数を数える
                while i < len(ws_rows[_header_row_index]):
                    #セルが空の場合は、停止します
                    if _col_start_index != -1 and (ws_rows[_header_row_index][i].value is None or str(ws_rows[_header_row_index][i].value).strip() == ''):
                        _col_end_index = i - 1
                        break
                    elif i == len(ws_rows[_header_row_index]) - 1:
                        _col_end_index = i
                    i += 1
                #新規テストスイートを作成する
                _test_suite_name = f[0:f.rindex('.')] + '-' + _sheet_name_list[sheet_index]
                test_suite_id = create_new_suite(_test_suite_name, ws_rows)

                #新規テストスイートバージョンを作成する
                test_suite_vs_id = create_new_tsv(test_suite_id)

                #QFにテストケースを登録する
                i = _header_row_index + 1   #読み取りを開始する行を指定します（ヘッダの行の次の行から）
                tc_no = 1                   #テストケースNo

                #各行をループして、QFに登録する
                while i < len(ws_rows):
                    #テストケースのフラグ
                    isTestCase = False
                    #テストケースかどうかをチェックする
                    for c in ws_rows[i]:
                        #空でないセルが少なくとも1つある場合、テストケースになります。
                        if c.value is not None and str(c.value).strip() != '':
                            isTestCase = True
                            break
                    #対象テストケースの場合
                    if isTestCase:
                        #テストケースを作成する
                        if create_new_tc(test_suite_id, test_suite_vs_id, ws_rows[i]):
                            print(str(i - _header_row_index) + '番目のテストケースの作成が合格しました。')
                            tc_no += 1
                        else:
                            print(str(i - _header_row_index) + '番目のテストケースのインポートに失敗しました。')
                    else:
                        #対象外テストケースの場合、終了する
                        break

                    i += 1

                #availableにstatusを編集する
                if update_test_suite_version(test_suite_id, test_suite_vs_id):
                    print('「' + _test_suite_name + '」テストスイートの' + str(i - _header_row_index - 1) + '件のテストケース作成が完了しました。')
                
                #シートの次
                sheet_index += 1
            
    print('テストケースのインポートが完了しました。')
    exit()