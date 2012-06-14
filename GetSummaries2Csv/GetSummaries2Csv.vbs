Option Explicit
' ===========================================================
' GetSummaries2Csv.vbs
' 復旧・復興支援制度データベースの全制度の概要をCSVファイルに出力する
' ===========================================================
'************************************************************
'定数
'************************************************************
'' 保存先のCSVファイル名
Const CSV_FILE_NAME = "SupportSummaries.csv"

'' 制度情報の複数IDを取得するURL
Const INFORMATIONS_URL = "http://api.r-assistance.go.jp/v1/api.svc/searchSupportInformations?appkey=0&maxcount=100&pageno="
'' 制度情報の複数IDが記述されたXMLファイルにつける接頭辞
Const INFORMATIONS_FILE_PREFIX = "SupportInformations_"

'' 制度情報の概要を取得するURL
Const SUMMARIES_URL = "http://api.r-assistance.go.jp/v1/api.svc/getSupportInformationSummaries?appkey=0&ids="
'' 制度情報の概要が記述されたXMLファイルにつける接頭辞
Const SUMMARIES_FILE_PREFIX = "SupportSummaries_"

'' 制度詳細情報のURL
Const DETAIL_URL = "http://www.r-assistance.go.jp/contentdetail_k.aspx?ContentID="

'' 地方公共団体（コードおよび名称）の情報を取得するURL
Const MUNICIPALITIES_URL = "http://api.r-assistance.go.jp/v1/api.svc/getMunicipalities?appkey=0&prefcode="
'' 地方公共団体（コードおよび名称）のXMLファイルにつける接頭辞
Const MUNICIPALITIES_PREFIX = "Municipalities_"

'************************************************************
' 初期処理、終了処理
'************************************************************
Dim g_objXMLHTTP    ' XMLHTTPオブジェクト
Dim g_objXML        ' XMLオブジェクト
Dim g_objFSO        ' ファイルシステムオブジェクト
Dim g_objDictionary ' 辞書オブジェクト

Set g_objXMLHTTP    = CreateObject("MSXML2.XMLHTTP.3.0")
Set g_objXML        = CreateObject("MSXML2.DOMDocument")
Set g_objFSO        = WScript.CreateObject("Scripting.FileSystemObject")
Set g_objDictionary = WScript.CreateObject("Scripting.Dictionary")

Call Main()
WScript.Echo "終了しました"

g_objDictionary.RemoveAll() '辞書の全要素削除
Set g_objDictionary = Nothing
Set g_objXMLHTTP = Nothing
Set g_objXML   = Nothing
Set g_objFSO   = Nothing

'************************************************************
' メイン処理
'************************************************************
Sub Main()
	Dim objCsvFile         ' 出力対象のCSVファイル
	Dim objCsvTextStream   ' 出力対象のCSVファイルの出力ストリーム

	Dim infosCount         ' 制度情報のIDが記述されたXMLの数
	Dim infosFileName      ' 制度情報のIDが記述されたXMLのファイル名
	Dim summariesFileName  ' 制度情報の概要が記述されたXMLのファイル名

	Dim i
	Dim ids

	' 地方公共団体（コードおよび名称）が設定された連想配列を作成
	If SetLocalGovCodeDic() <> 0 Then
		WScript.Echo "地方公共団体（コードおよび名称）の取得に失敗しました"
		Exit Sub
	End If

	' 既存のCSVファイルが存在すれば、削除
	if g_objFSO.FileExists(CSV_FILE_NAME) Then
		g_objFSO.DeleteFile CSV_FILE_NAME
	End if

	' CSVファイルを作成
	g_objFSO.CreateTextFile CSV_FILE_NAME
	Set objCsvFile = g_objFSO.GetFile(CSV_FILE_NAME)
	' 8:追記モード, -2:システムデフォルト文字コード SJISで
	Set objCsvTextStream = objCsvFile.OpenAsTextStream(8, -2)
	objCsvTextStream.WriteLine("ID,制度のタイトル,お問い合わせ先,対象地域,概要,詳細情報URL")

	' 制度情報のIDが記述されたXMLをローカルに保存
	infosCount = GetSupportInformations()

	If infosCount = 0 Then
		WScript.Echo "制度情報のIDの取得に失敗しました"
	Else
		For i = 1 To infosCount
			infosFileName = INFORMATIONS_FILE_PREFIX & i & ".xml"
			summariesFileName = SUMMARIES_FILE_PREFIX & i & ".xml"

			' 制度情報のIDが記述されたXMLから、ID部分を取得
			ids = ReadXmlTag(infosFileName, "IDs")
			' カンマ区切りをパイプ区切りに変換
			ids = Replace(ids, ",", "|")

			' 制度情報の概要が記述されたXMLをローカルに保存
			If GetSupportSummaries(ids, summariesFileName) = 0 Then
				' CSVファイルに出力
				Call ConvertSummaries2Csv(summariesFileName, objCsvTextStream)
			Else
				WScript.Echo "制度情報の概要の取得に失敗しました"
				Exit For
			End If
		Next
	End If

	objCsvTextStream.Close()
	Set objCsvTextStream = Nothing
	Set objCsvFile = Nothing
End Sub

'************************************************************
'関数ID          ：SetLocalGovCodeDic
'説明            ：連想配列に、地方公共団体（コードおよび名称）を設定する
'引数            ：なし
'戻り値          ：実行結果 0:成功 -1:失敗
'************************************************************
Function SetLocalGovCodeDic()
	Dim i
	Dim result
	Dim prefcode
	Dim saveFileName
	Dim targetUrl
	Dim nodeRoot
	Dim nodeItems
	Dim nodeItem
	Dim code
	Dim name
	Dim prefectureName
	Dim prefectureLocalGovCode

	SetLocalGovCodeDic = -1

	' 都道府県の情報を設定
	Call SetPrefectureLocalGovCodeDic(g_objDictionary)

	' 47都道府県の地方公共団体の情報を設定
	For i = 1 To 47
		' 2桁の数字に変換 ※VBScriptではFormat関数使えない
		If i < 10 Then
			prefcode = 0 & i
		Else
			prefcode = i
		End If

		saveFileName = MUNICIPALITIES_PREFIX & prefcode & ".xml"
		targetUrl = MUNICIPALITIES_URL & prefcode

		If SaveResponse(targetUrl, saveFileName) <> 0 Then
			WScript.Echo "地方公共団体（コードおよび名称）情報の取得に失敗しました"
			Exit For
		End If

		'非同期
		g_objXml.async=False
		' XMLファイルの読み込み
		result = g_objXml.Load(saveFileName)
		If Not result Then
			WScript.Echo g_objXml.parseError.errorCode
			WScript.Echo g_objXml.parseError.reason
			Exit Function
		End If

		Set nodeRoot = g_objXml.documentElement
		Set nodeItems = nodeRoot.childNodes

		prefectureLocalGovCode = prefcode & "000"
		prefectureName = g_objDictionary(prefectureLocalGovCode)

		For Each nodeItem In nodeItems
			code = nodeItem.getAttribute("LocalGovCode")
			name = prefectureName & nodeItem.getElementsByTagName("LocalGovName").Item(0).Text

			'連想配列に格納
			g_objDictionary.add code, name
		Next
	Next

	SetLocalGovCodeDic = 0
End Function

'************************************************************
'関数ID          ：getSupportInformations
'説明            ：制度情報の複数IDが記述されたXMLをローカルに保存する
'引数            ：なし
'戻り値          ：XMLファイルのファイル数
'                ：記述されていないものはカウントしない
'************************************************************
Function GetSupportInformations()
	Dim pageNo
	Dim targetUrl
	Dim saveFileName
	Dim resData

	pageNo = 1
	Do
		targetUrl = INFORMATIONS_URL & pageNo
		saveFileName = INFORMATIONS_FILE_PREFIX & pageNo & ".xml"

		If -1 = SaveResponse(targetUrl, saveFileName) Then
			Exit Do
		End If

		resData = ReadTextAll(saveFileName)

		' IDが記述されていないことを判定 (簡易判定)
		If InStr(resData, "<IDs/>") > 0 Then
			Exit Do
		End If

		pageNo = pageNo + 1
	LOOP

	GetSupportInformations = pageNo - 1
End Function

'************************************************************
'関数ID          ：GetSupportSummaries
'説明            ：制度情報の概要が記述されたXMLをローカルに保存する
'引数            ：ids … 対象となる制度のID、パイプ区切り、最大100個
'引数            ：saveFileName … 保存ファイル名
'戻り値          ：実行結果 0:成功 -1:失敗
'************************************************************
Function GetSupportSummaries(ids, saveFileName)
	Dim targetUrl

	GetSupportSummaries = -1

	targetUrl = SUMMARIES_URL & ids

	If SaveResponse(targetUrl, saveFileName) = 0 Then
		GetSupportSummaries = 0
	End If
End Function

'************************************************************
'関数ID          ：ConvertSummaries2Csv
'説明            ：制度情報の概要のXMLファイルをCSVファイルに保存
'引数            ：xmlFilePath … XMLファイルパス
'引数            ：objtextStream …出力ファイルの出力ストリーム
'戻り値          ：なし
'************************************************************
Sub ConvertSummaries2Csv(xmlFilePath, objtextStream)
	Dim result
	Dim nodeRoot
	Dim nodeItems
	Dim nodeItem
	Dim id
	Dim title
	Dim contact
	Dim intentArea
	Dim summary
	Dim lineData ' 書き込む行データ

	'非同期
	g_objXml.async=False
	' XMLファイルの読み込み
	result = g_objXml.Load(xmlFilePath)
	If Not result Then
		WScript.Echo g_objXml.parseError.errorCode
		WScript.Echo g_objXml.parseError.reason
		Exit Sub
	End If

	Set nodeRoot = g_objXml.documentElement
	' SupportInformationSummaryタグのリストを取得
	Set nodeItems = nodeRoot.getElementsByTagName("SupportInformationSummary")

	' SupportInformationSummaryタグを1つずつ取り出す。
	For Each nodeItem In nodeItems
		id = nodeItem.getAttribute("ID")
		title = nodeItem.getElementsByTagName("Title").Item(0).Text
		contact = nodeItem.getElementsByTagName("Contact").Item(0).Text
		intentArea = nodeItem.getElementsByTagName("IntendedArea").Item(0).Text
		summary = nodeItem.getElementsByTagName("Summary").Item(0).Text

		intentArea = ConvertLocalGovCode2LocalGovName(intentArea)

		lineData = id & ",""" & title & """,""" & contact & """,""" & intentArea & """,""" & summary & """," & DETAIL_URL & id
		objtextStream.WriteLine(lineData)
	Next
End Sub

'************************************************************
'関数ID          ：ConvertLocalGovCode2LocalGovName
'説明            ：地方公共団体コードを地方公共団体名に変換する
'引数            ：codes … 地方公共団体コード、カンマ区切り
'戻り値          ：地方公共団体名に変換された文字列、//区切り
'************************************************************
Function ConvertLocalGovCode2LocalGovName(codes)
	Dim array
	Dim i
	Dim names
	Dim code

	array = Split(codes, ",")
	names = ""

	For i = 0 To UBound(array)
		code = array(i)

		If i > 0 Then
			names = names & "//"
		End If

		If g_objDictionary.Exists(code) Then
			names = names & g_objDictionary(code)
		Else
			names = names & "[" & code & "]"
		End If
	Next

	ConvertLocalGovCode2LocalGovName = names

End Function

' ===========================================================
' 共通処理
' ===========================================================
'************************************************************
'関数ID          ：SaveResponse
'説明            ：指定されたURLの取得内容をファイルに保存する
'引数            ：url      … 取得先URL
'引数            ：filePath … ファイルパス
'戻り値          ：実行結果 0:成功 -1:失敗
'************************************************************
Function SaveResponse(url, filePath)
	Dim objAdodbStream    ' ADODB.Streamオブジェクト

	SaveResponse = -1

	' 既存のファイルを削除
	if g_objFSO.FileExists(filePath) Then
		g_objFSO.DeleteFile filePath
	End if

	' 同期処理
	g_objXMLHTTP.Open "GET", url, False
	g_objXMLHTTP.Send

	If g_objXMLHTTP.Status = 200 Then
		Set objAdodbStream = CreateObject("ADODB.Stream")
		objAdodbStream.Open
		objAdodbStream.Type = 1
		objAdodbStream.Write g_objXMLHTTP.responseBody
		objAdodbStream.SaveToFile filePath
		objAdodbStream.Close
		Set objAdodbStream = Nothing
		SaveResponse = 0
	Else
		WScript.Echo "Error returnCode:" & g_objXMLHTTP.Status
	End If
End Function

'************************************************************
'関数ID          ：ReadTextAll
'説明            ：テキストファイルの内容をすべて読み込みます
'引数            ：filePath ･･･ テキストファイルパス
'戻り値          ：ファイルの内容のテキスト
'************************************************************
Function ReadTextAll(filePath)
	Dim objTextStream  ' テキストストリームオブジェクト
	Dim resData        ' テキスト内容
	
	Set objTextStream = g_objFSO.OpenTextFile(filePath, 1)
	resData = objTextStream.ReadAll

	objTextStream.Close
	Set objTextStream = Nothing
	ReadTextAll = resData
End Function

'************************************************************
'関数ID          ：ReadXmlTag
'説明            ：指定されたXMLファイルの指定タグの情報を返す
'引数            ：xmlFilePath … XMLファイルパス
'引数            ：tagName … タグ名
'戻り値          ：指定タグの情報
'************************************************************
Function ReadXmlTag(xmlFilePath, tagName)
	Dim nodeRoot
	Dim nodeItems
	Dim nodeItem
	Dim result

	'非同期
	g_objXml.async=False
	' XMLファイルの読み込み
	result = g_objXml.Load(xmlFilePath)
	If Not result Then
		WScript.Echo g_objXml.parseError.errorCode
		WScript.Echo g_objXml.parseError.reason
		Exit Function
	End If

	Set nodeRoot = g_objXml.documentElement
	Set nodeItems = nodeRoot.getElementsByTagName(tagName)

	ReadXmlTag = nodeItems.Item(0).Text
End Function

' ===========================================================
' その他
' ===========================================================
'************************************************************
'関数ID          ：SetPrefectureLocalGovCodeDic
'説明            ：連想配列に、都道府県の地方公共団体（コードおよび名称）を設定する
'引数            ：dic … 連想配列格納用のDictionaryオブジェクト
'戻り値          ：なし
'************************************************************
Sub SetPrefectureLocalGovCodeDic(dic)
	'連想配列に格納 下記はJISで決められたコード体系

	dic.add "01000","北海道"
	dic.add "02000","青森県"
	dic.add "03000","岩手県"
	dic.add "04000","宮城県"
	dic.add "05000","秋田県"
	dic.add "06000","山形県"
	dic.add "07000","福島県"
	dic.add "08000","茨城県"
	dic.add "09000","栃木県"
	dic.add "10000","群馬県"
	dic.add "11000","埼玉県"
	dic.add "12000","千葉県"
	dic.add "13000","東京都"
	dic.add "14000","神奈川県"
	dic.add "15000","新潟県"
	dic.add "16000","富山県"
	dic.add "17000","石川県"
	dic.add "18000","福井県"
	dic.add "19000","山梨県"
	dic.add "20000","長野県"
	dic.add "21000","岐阜県"
	dic.add "22000","静岡県"
	dic.add "23000","愛知県"
	dic.add "24000","三重県"
	dic.add "25000","滋賀県"
	dic.add "26000","京都府"
	dic.add "27000","大阪府"
	dic.add "28000","兵庫県"
	dic.add "29000","奈良県"
	dic.add "30000","和歌山県"
	dic.add "31000","鳥取県"
	dic.add "32000","島根県"
	dic.add "33000","岡山県"
	dic.add "34000","広島県"
	dic.add "35000","山口県"
	dic.add "36000","徳島県"
	dic.add "37000","香川県"
	dic.add "38000","愛媛県"
	dic.add "39000","高知県"
	dic.add "40000","福岡県"
	dic.add "41000","佐賀県"
	dic.add "42000","長崎県"
	dic.add "43000","熊本県"
	dic.add "44000","大分県"
	dic.add "45000","宮崎県"
	dic.add "46000","鹿児島県"
	dic.add "47000","沖縄県"
End Sub

