###############################################################

使い方

###############################################################
社員の給与明細とか社員IDのあるものをExcelからpdfに一括でパスをつけて変換するためのプログラムです。

1.対象のExcelファイルの社員IDにあたるセル座標ををpassAddressList.txtに記載する。
　→Columnに列を入力してください。 ROWに列を入れてください。
　例）社員番号がK6の場合
  Column=K
  Row=6
  となります。

2.パスワードのリストを作成します。
　passList.txtに社員IDごとにパスワードを設定します。
  例）社員ID=password
      ABC1234=passTanaka
      ABC5656=passYamada

3.pdfにするExcelファイルをImportExcelファイルに格納します。

4.pip OpenPyXLを実行します。


###############################################################

基本設計

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


###############################################################
参考資料
給与明細テンプレート
https://bizroute.net/download/kyuuyo01

