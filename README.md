###############################################################

使い方




###############################################################

設計書替わり。
大雑把な概要。

大量のExcelファイルをpdf（パス付）にして保存するシステム。


ゆるふわ設計


1.開始時にOutputPDFの中身を全部削除する。

2.getPassIDTextを開きExcel内のセル座標のどこがIDを示している箇所を入手する

3.ImportExcelフォルダの対象Excelを開く。

4.Excelファイルのファイル名を取得する。

5.ImportPasswordにて該当のパスワードと一致するIDを探し、パスワードを入手。

6.該当のパスワードを使いImportExcelのExcelをOutputPDFに吐き出す。


