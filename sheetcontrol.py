import json
import csv
import math
import ConfigController as cc
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# -----------------------------------
# gcpにスプレッドシート操作用のAPI立てて、そのAPIに共有する必要あり
# -----------------------------------
# コンフィグファイル
configFile = './temtool/config.ini'
# 鍵は外だししているので取得
key = cc.Call(configFile).getPropertiesC("GoogleSheets","GDA_SECRET_KEY","pleseSetKey",r"Google Drive API")
# シート名
cini = cc.Call(configFile)
# シートのキーを指定
sheetName = cini.getPropertiesC("loadLanguageData","SHEET_NAME","1E-r3fWr4jGlIWGagdUoh1K_39OlZG5GyU6kUBSJvKTg","作業スプレッドシート名")

#認証情報を使ってスプレッドシートの操作権を取得
c = ServiceAccountCredentials._from_parsed_json_keyfile(json.loads(key), ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive'])
gs = gspread.authorize(c)


def loadcsv(filename):
    csv_file = open(filename, "r", encoding="utf-8", errors="", newline="" )
    f = csv.reader(csv_file, delimiter=",", doublequote=True, lineterminator="\r\n", quotechar='"', skipinitialspace=True)
    csv_file.close
    return f

def addNewLine(worksheet,headsize,kousinLen,width):
    """
    key値管理のデータに対して、未登録のkeyがある場合に
    既存に影響を出さずに追加を行う
    worksheet: 作業を行うシート
    headsize: 項目名など説明文を上に追加する行数
    kousinLen: 今回追加するデータ
            [[key,anyValue,anyValue],[key,anyValue,anyValue]]
            →この場合widthは3
    """
    size = len(kousinLen)
    cell_list = worksheet.range(1 + headsize, 1, size + headsize, width)
    nowNameList = []
    addNameList = []
    # 最後の行を調べるために初期値を最大で設定
    endpoint = size
    # スプレッドシートの行をすべて確認
    for i in range(math.floor(len(cell_list)/width)):
        # 登録されている中で終端に来たら行を記録
        if(cell_list[i*width].value == ""):
            endpoint = i + 1
            break
        nowNameList.append(cell_list[i*width].value)
    # CSVをすべて確認
    for k in kousinLen:
        # キーがあるかを確認してないものを追加
        if ((k[0] in nowNameList) != True):
            # 場合は追加用のものを取得
            #for j in range(width):
            #    addNameList.append(k[j])
            addNameList.append(k[0])
            addNameList.append(k[1])
    print(addNameList)
    # 新規追加処理を行う
    add_list = worksheet.range(endpoint + headsize, 1, endpoint + headsize + len(addNameList)/width + headsize + 1, width)
    # 追加候補を全部ぶっこむ
    for i in range(len(addNameList)):
        add_list[i].value = addNameList[i]
    # シートの反映
    worksheet.update_cells(add_list) 
    return 1

def getHonyakuLine(worksheet,headsize,maxsize,defaultline,dataline):
    cell_list = worksheet.range(1 + headsize, 1, maxsize + headsize, 1)
    cell_list2 = worksheet.range(1 + headsize, dataline, maxsize + headsize, dataline)
    cell_list3 = worksheet.range(1 + headsize, defaultline, maxsize + headsize, defaultline)
    honyakuList = []
    for i in range(len(cell_list)):
        if(cell_list[i].value != ""):
            if(cell_list2[i].value != ""):
                # 翻訳データがある場合
                honyakuList.append([cell_list[i].value,cell_list2[i].value])
            else:
                # 翻訳データがない場合はdefaultをいれる
                honyakuList.append([cell_list[i].value,cell_list3[i].value])
    return honyakuList

def writeCsv(filename, data):
    # ファイル出力処理
    f = open(filename, 'w', newline="", encoding='utf-8')
    try:
        writer = csv.writer(f)
        writer.writerows(data)
    except Exception as e:
        print(e)
    f.close()

def updateZukan():
    """
    図鑑のアップデートを行う
    """
    #共有したスプレッドシートのキー（後述）を使ってシートの情報を取得
    worksheet = gs.open_by_key(sheetName).worksheet('名前')
    # 図鑑データの読み込み
    f = loadcsv("./data/temtemstatus.csv");
    # 最大値の取得
    maxNo = 0
    hage = {}
    for row in f:
        hage[row[0]] = row[1]
        maxNo = maxNo if (maxNo > int(row[0])) else int(row[0])

    # 取得するぜ
    cell_list = worksheet.range(1, 1, maxNo, 2)

    # 取得したcell_listを編集
    for k in range(maxNo):
        # 図鑑番号
        cell_list[k*2].value = k+1
        # 名前(ない場合はNone)
        if str(k+1) in hage.keys():
            cell_list[k*2+1].value = hage.get(str(k+1))
        else:
            cell_list[k*2+1].value = "None"
    # 編集したcell_listをシートに反映
    worksheet.update_cells(cell_list) 


# 図鑑名称のアップデート
# キーに対する日本語訳を取得。訳されていない場合は英語をいれる
def loadJpName():
    #共有したスプレッドシートのキー（後述）を使ってシートの情報を取得
    worksheet = gs.open_by_key(sheetName).worksheet('名前')
    headsize = 0
    maxsize = 200
    defaultline = 2
    dataline = 3
    nameList = getHonyakuLine(worksheet,headsize,maxsize,defaultline,dataline)
    # ファイル出力
    writeCsv('./data/temtemJpName.csv', nameList)

def updatetechnic():
    """
    技のアップデートを行う
    """
    #共有したスプレッドシートのキー（後述）を使ってシートの情報を取得
    worksheet = gs.open_by_key(sheetName).worksheet('技名')
    # 図鑑データの読み込み
    f = loadcsv("./data/temtemtechniques.csv")
    # 新しい個数を確認 
    kousinLen = []
    for i in f:
        kousinLen.append(i)
    # 既存のサイズを取得
    headsize = 1
    addNewLine(worksheet,headsize,kousinLen,2)

def loadJptechnic():
    worksheet = gs.open_by_key(sheetName).worksheet('技名')
    headsize = 1
    f = loadcsv("./data/temtemtechniques.csv")
    # 新しい個数を確認 
    csvline = []
    for i in f:
        csvline.append(i)
    maxsize = len(csvline)
    defaultline = 2
    dataline = 3
    line = getHonyakuLine(worksheet,headsize,maxsize,defaultline,dataline)
    writeCsv('./data/temtemJptechnic.csv', line)

def updatetrate():
    """
    個性のアップデートを行う
    """
    #共有したスプレッドシートのキー（後述）を使ってシートの情報を取得
    worksheet = gs.open_by_key(sheetName).worksheet('個性')
    # 図鑑データの読み込み
    f = loadcsv("./data/temtemTrateDetail.csv")
    # 新しい個数を確認 
    kousinLen = []
    for i in f:
        kousinLen.append(i)
    # 既存のサイズを取得
    headsize = 1
    addNewLine(worksheet,headsize,kousinLen,2)

def loadJptrate():
    worksheet = gs.open_by_key(sheetName).worksheet('個性')
    headsize = 1
    f = loadcsv("./data/temtemTrateDetail.csv")
    # 新しい個数を確認 
    csvline = []
    for i in f:
        csvline.append(i)
    maxsize = len(csvline)
    defaultline = 2
    dataline = 3
    line = getHonyakuLine(worksheet,headsize,maxsize,defaultline,dataline)
    writeCsv('./data/temtemJptrate.csv', line)
def updateitem():
    """
    道具のアップデートを行う
    """
    #共有したスプレッドシートのキー（後述）を使ってシートの情報を取得
    worksheet = gs.open_by_key(sheetName).worksheet('道具')
    # 図鑑データの読み込み
    f = loadcsv("./data/temtemitem.csv")
    # 新しい個数を確認 
    kousinLen = []
    for i in f:
        kousinLen.append(i)
    # 既存のサイズを取得
    headsize = 1
    addNewLine(worksheet,headsize,kousinLen,2)

def loadJpitem():
    worksheet = gs.open_by_key(sheetName).worksheet('道具')
    headsize = 1
    f = loadcsv("./data/temtemitem.csv")
    # 新しい個数を確認 
    csvline = []
    for i in f:
        csvline.append(i)
    maxsize = len(csvline)
    defaultline = 2
    dataline = 3
    line = getHonyakuLine(worksheet,headsize,maxsize,defaultline,dataline)
    writeCsv('./data/temtemJpitem.csv', line)
# updateZukan()
#loadJpName()
#updatetechnic()
loadJptechnic()
#updatetrate()
#loadJptrate()
#updateitem()
#loadJpitem()

