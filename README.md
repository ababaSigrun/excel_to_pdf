###############################################################

使い方




###############################################################

設計書替わり。
大雑把な概要。

大量のExcelファイルをpdf（パス付）にして保存するシステム。


詳細設計
開始時にOutputPDFの中身を全部削除する。
importExcelフォルダに対象のExcelを入れる。
Excelファイルのファイル名を取得する。
getPassIDTextにExcel内のセル座標を入手し、IDを入手
ImportPasswordにて該当のパスワードと一致するIDを探し、パスワードを入手。

該当のパスワードを使いImportExcelのExcelをOutputPDFに吐き出す。


