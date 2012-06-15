復旧・復興支援制度データベースAPIを操作するVBScript
======================
[復旧・復興支援制度データベースAPI][fukkouAPI]を使用したサンプルをVBScriptで作ってみました。
[fukkouAPI]: http://www.r-assistance.go.jp/about_api.aspx

サンプルの内容
------
## GetSummaries2Csv ##
復旧・復興支援制度の全概要を取得し、CSVファイルに出力するVBScriptです。

### 使い方 ###
1. [GetSummaries2Csv.vbs][VBSFile]を、ご自分のPCの適当なフォルダにコピーします。

2.コマンドプロンプトを起動し、コマンドプロンプトから『cscript //nologo GetSummaries2Csv.vbs』と入力して実行します。

3.『GetSummaries2Csv.vbs』が存在するフォルダに、以下のファイルが出力されます。
* 処理結果CSVファイル 『SupportSummaries.csv』[Link][CSVFile]
* 地方公共団体（コードおよび名称）が記述されたXMLファイル 『Municipalities_(都道府県コード、01～47).xml』
* 制度のIDが記述されたXMLファイル 『SupportInformations_(連番).xml』
* 制度の概要が記述されたXMLファイル 『SupportSummaries_(連番).xml』

3.『終了しました』と表示されたら、完了です。![処理結果][CapFile]

4.出力されたCSVファイルは、Excelで開くときに警告は出ますが正常に表示されるはずです。

[VBSFile]: https://github.com/kamatamadai/r-assistance-api-vbscript/blob/master/GetSummaries2Csv/GetSummaries2Csv.vbs
[CSVFile]: https://github.com/kamatamadai/r-assistance-api-vbscript/blob/master/GetSummaries2Csv/SupportSummaries.csv
[CapFile]: https://github.com/kamatamadai/r-assistance-api-vbscript/blob/master/GetSummaries2Csv/captureResults.PNG

ライセンス
----------
許諾なしで、ソースの改修や再配布を行えます。

