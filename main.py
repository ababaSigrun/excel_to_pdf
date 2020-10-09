import os
import glob


#  1.まずは出力先の中身を空にする。
for file in glob.glob('./OutputPDF/*.pdf'):
    os.remove(file)

# 該当のUserIdが表記されている箇所をメモ帳から取得する。


# 