import os
import glob



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


# メインメソッド
print('スタート')
# 出力先の削除を行う。
delPDF()

# 指定座標を取得する。
col,row = getPassIDText()

print(col + ":"+ row)