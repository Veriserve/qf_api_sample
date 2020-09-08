try:
    from fractions import Fraction
    import datetime
    import locale
    import platform
    import math
except ModuleNotFoundError as err:
    print('「' + err.name + '」のモジュールが見つかりません。以下のコマンドラインでインストールしてください。')
    print('>> pip install ' + err.name)
    exit()

#組み込みフォーマットの分数の定義
BUILTIN_FORMATS_FRACTION = {
    '# ?/?' : 10,               #Up to one digit
    '# ??/??' : 100,            #Up to two digits
    '#\\ ???/???' : 1000,       #Up to three digits
    '#\\ ?/2' : 2,              #As halves (1/2)
    '#\\ ?/4' : 4,              #As quarters (2/4)
    '#\\ ?/8' : 8,              #As eighths (4/8)
    '#\\ ??/16' : 16,           #As sixteenths (8/16)
    '#\\ ?/10' : 10,            #As tenths (3/10)
    '#\\ ??/100' : 100          #As hundredths (30/100)
    }
#データ型は数値です
TYPE_NUMERIC = 'n'
#データ型は日付です
TYPE_DATE = 'd'

# -----------------------------------------------------------
# 値をパーセント形式に変換する
#
# Parameters:
#  val（float）：Excelから読み取る値
#  format（str）：Excelから読み取る形式
#
# Returns:
#  変換した値
# -----------------------------------------------------------
def format_percent(val, format):
    num = len(format[format.rindex('.') + 1 : format.rindex('%')])
    return ('{:.' + str(num) + '%}').format(val)

# -----------------------------------------------------------
# 値を分数形式に変換する
#
# Parameters:
#  val（float）：Excelから読み取る値
#  format（str）：Excelから読み取る形式
#
# Returns:
#  変換した値
# -----------------------------------------------------------
def format_fraction(val, format):
    whole = math.floor(Fraction(val))
    frac = val - whole
    fmt = BUILTIN_FORMATS_FRACTION[format]
    if format.find('?/?') > -1:
        return str(whole) + ' ' + str(Fraction(frac).limit_denominator(fmt))
    else:
        if (frac * fmt) - math.floor(frac * fmt) < 0.5:
            #端数を切り捨てる
            numerator = math.floor(frac * fmt)
        else:
            #端数を切り上げする
            numerator = math.ceil(frac * fmt)
        denominator = fmt
        return str(whole) + ' ' + str(numerator) + '/' + str(denominator)

# -----------------------------------------------------------
# 値を数値形式に変換する
#
# Parameters:
#  val（float）：Excelから読み取る値
#  format（str）：Excelから読み取る形式
#
# Returns:
#  変換した値
# -----------------------------------------------------------
def format_numberic(val, format):
    comma_index = format.find('.')
    digits = 0
    i = comma_index + 1
    #ドットの後の数を数える
    while i < len(format):
        if format[i] == '0':
            digits += 1
        else:
            break
        i += 1
    return ('{:.' + str(digits) + 'f}').format(val)

# -----------------------------------------------------------
# 値を科学形式に変換する
#
# Parameters:
#  val（float）：Excelから読み取る値
#  format（str）：Excelから読み取る形式
#
# Returns:
#  変換した値
# -----------------------------------------------------------
def format_scientific(val, format):
    comma_index = format.find('.')
    digits = 0
    i = comma_index + 1
    #ドットの後の数を数える
    while i < len(format):
        if format[i] == '0':
            digits += 1
        else:
            break
        i += 1
    return ('{:.' + str(digits) + 'E}').format(val)

# -----------------------------------------------------------
# 値を通貨形式に変換する
#
# Parameters:
#  val（float）：Excelから読み取る値
#  format（str）：Excelから読み取る形式
#
# Returns:
#  変換した値
# -----------------------------------------------------------
def format_currency(val, format):
    comma_index = format.find('.')
    digits = 0
    i = comma_index + 1
    #ドットの後の数を数える
    while i < len(format):
        if format[i] == '0':
            digits += 1
        else:
            break
        i += 1

    #通貨のユニット（例：￥）
    prex = ''

    #通貨のユニットを取得する
    if format.find('"') > -1:
        j = format.find('"') + 1
        while j < len(format):
            if format[j] == '"':
                break
            prex = format[j]
            j += 1
    return (prex + '{:,.' + str(digits) + 'f}').format(val)

# -----------------------------------------------------------
# 値を日付形式に変換する
#
# Parameters:
#  val（float）：Excelから読み取る値
#  format（str）：Excelから読み取る形式
#
# Returns:
#  変換した値
# -----------------------------------------------------------
def format_datetime(val, format):
    locale.setlocale(locale.LC_CTYPE, "Japanese_Japan.932")
    if platform.system() == 'Windows':
        padding = '#'
    else:
        padding = '-'

    fmt = format
    #フォーマットに「\」を全て削除する
    if fmt.find('\\'):
        fmt = fmt.replace('\\', '')
    #フォーマットに「;@」を全て削除する
    if fmt.find(';@'):
        fmt = fmt.replace(';@', '')
    #フォーマットに「”」を全て削除する
    if fmt.find('"'):
        fmt = fmt.replace('"', '')
    #フォーマットに年の形式を取り替える
    if fmt.find('yyyy') > -1:
        fmt = fmt.replace('yyyy', '%Y')
    elif fmt.find('yy') > -1:
        fmt = fmt.replace('yy', '%y')
    #フォーマットに分の形式を取り替える
    if fmt.find(':mm') > -1:
        fmt = fmt.replace('mm', '%M')
    #フォーマットに月の形式を取り替える
    if fmt.find('mmmmm') > -1:
        fmt = fmt.replace('mmmmm', '%B', 1)
    elif fmt.find('mmmm') > -1:
        fmt = fmt.replace('mmmm', '%B', 1)
    elif fmt.find('mmm') > -1:
        fmt = fmt.replace('mmm', '%b', 1)
    elif fmt.find('mm') > -1:
        fmt = fmt.replace('mm', '%m', 1)
    elif fmt.find('m') > -1:
        fmt = fmt.replace('m', '%{}m'.format(padding), 1)
    #フォーマットに日の形式を取り替える
    if fmt.find('dd') > -1:
        fmt = fmt.replace('dd', '%d')
    elif fmt.find('d') > -1:
        fmt = fmt.replace('d', '%{}d'.format(padding))
    #フォーマットに時の形式を取り替える
    if fmt.find('H') > -1:
        fmt = fmt.replace('H', '%H')
    elif fmt.find('h') > -1:
        fmt = fmt.replace('h', '%{}H'.format(padding))
    #フォーマットに秒の形式を取り替える
    if fmt.find('ss') > -1:
        fmt = fmt.replace('ss', '%S')
    #AM/PMの形式を取り替える
    if fmt.find('AM/PM') > -1:
        fmt = fmt.replace('AM/PM', '%p')
    #フォーマットに[$-409]を削除る
    if fmt.find(']') > -1:
        fmt = fmt[fmt.index(']') + 1 : len(fmt)]

    return val.strftime(fmt)

# -----------------------------------------------------------
# 値をExcelの形式に変換する
#
# Parameters:
#  val（float）：Excelから読み取る値
#  type（str）：データタイプ
#  format（str）：Excelから読み取る形式
#
# Returns:
#  変換した値
# -----------------------------------------------------------
def format_value(val, type, format):
    data = val
    #数値の場合
    if type == TYPE_NUMERIC:
        if format.find('#,#') > -1:
            #通貨の形式の場合
            data = format_currency(val, format)
        elif format.find('0%') > -1:
            #パーセントの形式の場合
            data = format_percent(val, format)
        elif format.find('?/') > -1:
            #分数の形式の場合
            data = format_fraction(val, format)
        elif format.find('E+0') > -1:
            #科学の形式の場合
            data = format_scientific(val, format)
        elif format.find('0.0') > -1:
            #数値の形式の場合
            data = format_numberic(val, format)
    #日付の場合
    elif type == TYPE_DATE:
        data = format_datetime(val, format)
    return str(data)