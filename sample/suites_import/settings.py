BASE_API_URL = 'https://cloud.veriserve.co.jp/api/v2/'          #APIのパス
API_KEY = 'XXXX'                                                #APIのキー
FOLDER_PATH = 'XXXX'                                            #対象ファイルを格納するフォルダーのパス
TSV_NAME = '1.0'                                                #テストスイートバージョン名
QF_COLUMN_MAX = 21                                              #QFで取り込める列の最大（テスト定義の自由項目数 +「優先度」列）
TSV_STATUS = 'available'                                        #テストスイートバージョンのステータスの値（利用可）
COL_TITLE_START  = '優先度'                                     #ヘッダーの最初の列のタイトル
TEST_SUITE_DELETE_FLG = 1                                       #1：削除する、0：削除しない