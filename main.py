import os
import glob
import openpyxl

# pip install openpyxlをやってExcelを開ける環境を作ること。

#  出力先の中身を空にする。
def delPDF ():
    print('ファイル消去メソッド呼び出し')
    # for file in glob.glob('./OutputPDF/*.pdf'):
        # os.remove(file)


# ファイルを開きExcel内部でIDが書いてあるセルを特定する。
def getPassIDText():
    print('IDセル特定呼出')
    f = open('./ImportPasswordText/passAddressList.txt')
    lines = f.readlines()
    count = 0
    splitWord = '=' # 区切り文字を定義
    rowLine = ""
    colLine = ""
    for line in lines:
        count = count + 1
        print(line)
        if line.find("Column") !=-1 : 
            print('Column行' + line)
            colLine = line
        elif line.find("Row")  != -1: 
            print('Row行' + line)
            rowLine = line
        else:
            print('関係ない行が含まれています   : ' + str(count) + '行目')

    return getSplit(colLine,splitWord),getSplit(rowLine,splitWord)

# 文字列を分割し、改ページがあったら削除する
def getSplit(word,splitWord):
    retWord = word.split(splitWord)[1].replace('\n','')
    return retWord


# パスリストを取得する
def getPassList():
    f = open('./getPassIDText/passList.txt')
    lines = f.readlines()
    passList = []
    for line in lines:
        passList.append(line.replace('\n',''))
    return passList

# 元のExcelのファイル名を取得する。
def getTagetFileNameList():
    retList = []
    f_path = 'ImportExcel'
    for f in glob.glob('ImportExcel/*.xlsx'):
        retList.append(os.path.split(f)[1])
    return retList

#Excelファイルを開く。
def getExcel(fileName,targetAddress,sheetName) :
    wb=openpyxl.load_workbook('ImportExcel/' + fileName)
    sheet = wb.get_sheet_by_name(sheetName)
    # 対象のセルのアドレスを取得
    x = sheet[targetAddress].value
    print(x)




# メインメソッド
print('スタート')
# 出力先の削除を行う。
delPDF()

# 指定座標を取得する。
col,row = getPassIDText()
# 座標生成
targetAddress = col + str(row)

# パスリストを取得する。
passList = getPassList()

# 対象のファイル名を取得する。
tagetFileNameList = getTagetFileNameList()




# 対象のファイルを開く
for name in tagetFileNameList :
    getExcel(name,targetAddress,'Sheet1')



