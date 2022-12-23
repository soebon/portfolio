#ライブラリインポート
import pyautogui as pgui
import time,win32gui,win32con,ctypes
import os,string,subprocess
import pyocr
from PIL import Image, ImageEnhance
from openpyxl import Workbook

#プログラムを作製したパソコンの画面サイズ
DEF_SCR_SIZE = [1920,1080]
#適用するパソコンのサイズ
CUR_SCR_SIZE = [1920,1080]
def foreground():
    hwnd = ctypes.windll.user32.FindWindowW(0,"BlueStacks")
    win32gui.SetWindowPos(hwnd,win32con.HWND_TOP,0,0,0,0,win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
    #2つ目の要素の内容
    #HWND_TOPMOST:ウィンドウを常に最前面にする。
    #HWND_BOTTOM:ウィンドウを最後に置く。
    #HWND_NOTOPMOST:ウィンドウの最前面解除。
    #HWND_TOP:ウィンドウを先頭に置く。

    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    win32gui.SetForegroundWindow(hwnd)
    pgui.moveTo(left+60, top + 10)
    pgui.click()
def coordinate(x,y):
    x_cur = CUR_SCR_SIZE[0]*x/DEF_SCR_SIZE[0]
    y_cur = CUR_SCR_SIZE[1]*y/DEF_SCR_SIZE[1]
    return [x_cur,y_cur]
def scroll():
    #台データ（数値）が全部見れるように下に移動
    time.sleep(2)
    count = 23
    while count >0:
        pgui.press(['down'])
        count -= 1
def screenshot_1():
    #1日前
    pgui.screenshot('.\\screenshot\\bonus_1.png',region=coordinate(1065,277)+coordinate(36,30))
    pgui.screenshot('.\\screenshot\\big_1.png',region=coordinate(1065,355)+coordinate(36,30))
    pgui.screenshot('.\\screenshot\\sum_game_1.png',region=coordinate(1037,743)+coordinate(62,30))
    pgui.screenshot('.\\screenshot\\last_game_1.png',region=coordinate(1050,820)+coordinate(50,30))
def ocr_function(ocr_datafile):
    #Pah設定
    TESSERACT_PATH = 'C:\\Users\\htska\\AppData\\Local\\Programs\\Tesseract-OCR' #インストールしたTesseract-OCRのpath
    TESSDATA_PATH = 'C:\\Users\\htska\\AppData\\Local\\Programs\\Tesseract-OCR\\tessdata' #tessdataのpath

    os.environ["PATH"] += os.pathsep + TESSERACT_PATH
    os.environ["TESSDATA_PREFIX"] = TESSDATA_PATH

    #OCRエンジン取得
    tools = pyocr.get_available_tools()
    tool = tools[0]

    #OCRの設定 ※tesseract_layout=6が精度には重要。デフォルトは3
    builder = pyocr.builders.TextBuilder(tesseract_layout=7)

    #解析画像読み込み(雨ニモマケズ)
    img = Image.open(ocr_datafile) #他の拡張子でもOK

    #画像からOCRで日本語を読んで、文字列として取り出す
    txt_pyocr = tool.image_to_string(img , builder=builder)
    try:
      return(int(txt_pyocr))
    except ValueError:
      return(-1)
def data_get_input(column_number,kishu,kishudaisuu):#column_numberは本日:1、1日前:2、2日前:3、kishuはアイム:0、マイ:1、ファンキー:2、kishudaisuuは台数-1

    #先頭番号を押す
    time.sleep(5)
    pgui.click(coordinate(726,877),duration=3)
    #台のデータが見えるようにスクロール
    time.sleep(5)
    scroll()
    #スクリーンショット(column_numberと対応づけ)
    screenshot_1()
    #値を格納する
    bonus_1 = ocr_function('.\\screenshot\\bonus_1.png')
    big_1 = ocr_function('.\\screenshot\\big_1.png')
    sum_game_1 = ocr_function('.\\screenshot\\sum_game_1.png')
    last_game_1 = ocr_function('.\\screenshot\\last_game_1.png')
    datas_1 = [bonus_1,big_1,sum_game_1,last_game_1]
    #シート番号を指定する
    sheet = wb.worksheets[kishu]#アイム:0、マイ:1、ファンキー:2
    #アルファベットのリストを作製
    alphabet =list(string.ascii_uppercase)
    #入力したときの日付を入力する
    import datetime
    dt_now = datetime.datetime.now()
    sheet['A1']=dt_now.strftime('%Y年%m月%d日 %H')
    #「本日」、「1日前」、「2日前」を入力する
    sheet[alphabet[column_number]+str(1)] = '1日前'
    #データを入力する
    for i,data  in zip([2,3,8,9],datas_1):
        sheet[alphabet[column_number]+str(i)].value = data
    #埋まってない値に数式を書き込む
    sheet[alphabet[column_number]+str(4)].value = '='+alphabet[column_number]+str(2)+'-'+alphabet[column_number]+str(3)
    sheet[alphabet[column_number]+str(5)].value = '='+alphabet[column_number]+str(8)+'/'+alphabet[column_number]+str(3)
    sheet[alphabet[column_number]+str(6)].value = '='+alphabet[column_number]+str(8)+'/'+alphabet[column_number]+str(4)
    sheet[alphabet[column_number]+str(7)].value = '='+alphabet[column_number]+str(8)+'/'+alphabet[column_number]+str(2)
    #上の結果を繰り返す('台数-1'をrange()の中に入力する)
    for repeat in range(kishudaisuu):
        #次台を押す
        pgui.click(1211,82)
        #台のデータが見えるようにスクロール
        scroll()
        #スクリーンショット
        screenshot_1()
        #値を格納する
        bonus_1 = ocr_function('.\\screenshot\\bonus_1.png')
        big_1 = ocr_function('.\\screenshot\\big_1.png')
        sum_game_1 = ocr_function('.\\screenshot\\sum_game_1.png')
        last_game_1 = ocr_function('.\\screenshot\\last_game_1.png')
        datas_1 = [bonus_1,big_1,sum_game_1,last_game_1]
        #データを入力する
        for i,data  in zip([2,3,8,9],datas_1):
            sheet[alphabet[column_number]+str(i+10*(repeat+1))].value = data
        #埋まってない値に数式を書き込む
        sheet[alphabet[column_number]+str(4+10*(repeat+1))].value = '='+alphabet[column_number]+str(2+10*(repeat+1))+'-'+alphabet[column_number]+str(3+10*(repeat+1))
        sheet[alphabet[column_number]+str(5+10*(repeat+1))].value = '='+alphabet[column_number]+str(8+10*(repeat+1))+'/'+alphabet[column_number]+str(3+10*(repeat+1))
        sheet[alphabet[column_number]+str(6+10*(repeat+1))].value = '='+alphabet[column_number]+str(8+10*(repeat+1))+'/'+alphabet[column_number]+str(4+10*(repeat+1))
        sheet[alphabet[column_number]+str(7+10*(repeat+1))].value = '='+alphabet[column_number]+str(8+10*(repeat+1))+'/'+alphabet[column_number]+str(2+10*(repeat+1))
#ぱちタウンを起動する
#cp= subprocess.run(r'"C:\Program Files\BlueStacks_nxt\HD-Player.exe" --instance Nougat64 --cmd launchApp --package "com.dmm.ptown"')
#time.sleep(2)
#全画面表示にする
foreground()
pgui.press('F11')

#ぱちタウンのホーム画面
#マイページを押す
pgui.click(coordinate(677,45),duration=1)
#まるみつ霧島店を押す
pgui.click(coordinate(969,383),duration=2)
#データ公開を押す
pgui.click(coordinate(966,566),duration=3)
#「機種名」検索バーを押す
pgui.click(coordinate(954,236),duration=3)
#「あいむ」と入力して検索する
pgui.click()
pgui.press('hanja')
pgui.typewrite('aimu')
pgui.press('enter')
pgui.press('hanja')
pgui.press('enter')
time.sleep(3)
#「台番号から台を探す」を押す
pgui.click(coordinate(939,530),duration=3)
#'1'を入力
pgui.typewrite('1')
#Sアイムジャグラーを押す
pgui.click(coordinate(932,722),duration=3)
#screenshotのディレクトリ作製
path = '.\\screenshot'
os.mkdir(path)
#エクセルのファイルをopenpyxlで開く
from openpyxl import Workbook, load_workbook
wb = load_workbook("マルミツデータ一時保存.xlsx")
#データ入力(column_numberは本日:1、1日前:2、2日前:3、kishuはアイム:0、マイ:1、ファンキー:2、kishudaisuuは台数-1)
data_get_input(2,0,28)
#上スクロール
time.sleep(2)
count = 5
while count >0:
    pgui.press(['up'])
    count -= 1
#検索バーを押す
pgui.click(966,97)
#'1'を削除する
pgui.press('backspace')
#「まい」と入力して検索する
pgui.press('hanja')
pgui.typewrite('mai')
pgui.press('enter')
pgui.press('hanja')
pgui.press('enter')
time.sleep(3)
#「台番号から台を探す」を押す
pgui.click(coordinate(939,530),duration=2)
#「まい」を削除する
pgui.press('backspace')
pgui.press('backspace')
#'2'を入力
pgui.typewrite('2')
#SマイジャグラーVを押す
pgui.click(coordinate(932,722),duration=3)
#データ入力(column_numberは本日:1、1日前:2、2日前:3、kishuはアイム:0、マイ:1、ファンキー:2、kishudaisuuは台数-1)
data_get_input(2,1,19)
#上スクロール
time.sleep(2)
count = 5
while count >0:
    pgui.press(['up'])
    count -= 1
#検索バーを押す
pgui.click(966,97)
#'2'を削除する
pgui.press('backspace')
#「ふぁんきー」と入力して検索する
pgui.press('hanja')
pgui.typewrite('fanki-')
pgui.press('enter')
pgui.press('hanja')
pgui.press('enter')
time.sleep(3)
#「台番号から台を探す」を押す
pgui.click(coordinate(939,530),duration=2)
#「ふぁんきー」を削除する
pgui.press('backspace')
pgui.press('backspace')
pgui.press('backspace')
pgui.press('backspace')
pgui.press('backspace')
#'3'を入力
pgui.typewrite('3')
#Sファンキージャグラー2を押す
pgui.click(coordinate(932,722),duration=3)
#データ入力(column_numberは本日:1、1日前:2、2日前:3、kishuはアイム:0、マイ:1、ファンキー:2、kishudaisuuは台数-1)
data_get_input(2,2,4)
#上スクロール
time.sleep(2)
count = 5
while count >0:
    pgui.press(['up'])
    count -= 1
#検索バーを押す
pgui.click(966,97)
#'3'を削除する
pgui.press('backspace')
#上スクロール
time.sleep(2)
count = 10
while count >0:
    pgui.press(['up'])
    count -= 1
#「TOP」を押す
pgui.press('enter')
#最大化をやめる
pgui.press('F11')
#エクセルを保存する
wb.save("マルミツデータ一時保存.xlsx")
#スクリーンショットのディレクトリを削除する
import shutil
shutil.rmtree(path)

