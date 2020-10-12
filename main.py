import os
import glob




#  出力先の中身を空にする。
def delPDF ():
    for file in glob.glob('./OutputPDF/*.pdf'):
        os.remove(file)

# 該当のUserIdが表記されている箇所をメモ帳から取得する。


#
# メインメソッド

#出力先の削除を行う。
delPDF()