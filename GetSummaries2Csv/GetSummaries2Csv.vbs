Option Explicit
' ===========================================================
' GetSummaries2Csv.vbs
' �����E�����x�����x�f�[�^�x�[�X�̑S���x�̊T�v��CSV�t�@�C���ɏo�͂���
' ===========================================================
'************************************************************
'�萔
'************************************************************
'' �ۑ����CSV�t�@�C����
Const CSV_FILE_NAME = "SupportSummaries.csv"

'' ���x���̕���ID���擾����URL
Const INFORMATIONS_URL = "http://api.r-assistance.go.jp/v1/api.svc/searchSupportInformations?appkey=0&maxcount=100&pageno="
'' ���x���̕���ID���L�q���ꂽXML�t�@�C���ɂ���ړ���
Const INFORMATIONS_FILE_PREFIX = "SupportInformations_"

'' ���x���̊T�v���擾����URL
Const SUMMARIES_URL = "http://api.r-assistance.go.jp/v1/api.svc/getSupportInformationSummaries?appkey=0&ids="
'' ���x���̊T�v���L�q���ꂽXML�t�@�C���ɂ���ړ���
Const SUMMARIES_FILE_PREFIX = "SupportSummaries_"

'' ���x�ڍ׏���URL
Const DETAIL_URL = "http://www.r-assistance.go.jp/contentdetail_k.aspx?ContentID="

'' �n�������c�́i�R�[�h����і��́j�̏����擾����URL
Const MUNICIPALITIES_URL = "http://api.r-assistance.go.jp/v1/api.svc/getMunicipalities?appkey=0&prefcode="
'' �n�������c�́i�R�[�h����і��́j��XML�t�@�C���ɂ���ړ���
Const MUNICIPALITIES_PREFIX = "Municipalities_"

'************************************************************
' ���������A�I������
'************************************************************
Dim g_objXMLHTTP    ' XMLHTTP�I�u�W�F�N�g
Dim g_objXML        ' XML�I�u�W�F�N�g
Dim g_objFSO        ' �t�@�C���V�X�e���I�u�W�F�N�g
Dim g_objDictionary ' �����I�u�W�F�N�g

Set g_objXMLHTTP    = CreateObject("MSXML2.XMLHTTP.3.0")
Set g_objXML        = CreateObject("MSXML2.DOMDocument")
Set g_objFSO        = WScript.CreateObject("Scripting.FileSystemObject")
Set g_objDictionary = WScript.CreateObject("Scripting.Dictionary")

Call Main()
WScript.Echo "�I�����܂���"

g_objDictionary.RemoveAll() '�����̑S�v�f�폜
Set g_objDictionary = Nothing
Set g_objXMLHTTP = Nothing
Set g_objXML   = Nothing
Set g_objFSO   = Nothing

'************************************************************
' ���C������
'************************************************************
Sub Main()
	Dim objCsvFile         ' �o�͑Ώۂ�CSV�t�@�C��
	Dim objCsvTextStream   ' �o�͑Ώۂ�CSV�t�@�C���̏o�̓X�g���[��

	Dim infosCount         ' ���x����ID���L�q���ꂽXML�̐�
	Dim infosFileName      ' ���x����ID���L�q���ꂽXML�̃t�@�C����
	Dim summariesFileName  ' ���x���̊T�v���L�q���ꂽXML�̃t�@�C����

	Dim i
	Dim ids

	' �n�������c�́i�R�[�h����і��́j���ݒ肳�ꂽ�A�z�z����쐬
	If SetLocalGovCodeDic() <> 0 Then
		WScript.Echo "�n�������c�́i�R�[�h����і��́j�̎擾�Ɏ��s���܂���"
		Exit Sub
	End If

	' ������CSV�t�@�C�������݂���΁A�폜
	if g_objFSO.FileExists(CSV_FILE_NAME) Then
		g_objFSO.DeleteFile CSV_FILE_NAME
	End if

	' CSV�t�@�C�����쐬
	g_objFSO.CreateTextFile CSV_FILE_NAME
	Set objCsvFile = g_objFSO.GetFile(CSV_FILE_NAME)
	' 8:�ǋL���[�h, -2:�V�X�e���f�t�H���g�����R�[�h SJIS��
	Set objCsvTextStream = objCsvFile.OpenAsTextStream(8, -2)
	objCsvTextStream.WriteLine("ID,���x�̃^�C�g��,���₢���킹��,�Ώےn��,�T�v,�ڍ׏��URL")

	' ���x����ID���L�q���ꂽXML�����[�J���ɕۑ�
	infosCount = GetSupportInformations()

	If infosCount = 0 Then
		WScript.Echo "���x����ID�̎擾�Ɏ��s���܂���"
	Else
		For i = 1 To infosCount
			infosFileName = INFORMATIONS_FILE_PREFIX & i & ".xml"
			summariesFileName = SUMMARIES_FILE_PREFIX & i & ".xml"

			' ���x����ID���L�q���ꂽXML����AID�������擾
			ids = ReadXmlTag(infosFileName, "IDs")
			' �J���}��؂���p�C�v��؂�ɕϊ�
			ids = Replace(ids, ",", "|")

			' ���x���̊T�v���L�q���ꂽXML�����[�J���ɕۑ�
			If GetSupportSummaries(ids, summariesFileName) = 0 Then
				' CSV�t�@�C���ɏo��
				Call ConvertSummaries2Csv(summariesFileName, objCsvTextStream)
			Else
				WScript.Echo "���x���̊T�v�̎擾�Ɏ��s���܂���"
				Exit For
			End If
		Next
	End If

	objCsvTextStream.Close()
	Set objCsvTextStream = Nothing
	Set objCsvFile = Nothing
End Sub

'************************************************************
'�֐�ID          �FSetLocalGovCodeDic
'����            �F�A�z�z��ɁA�n�������c�́i�R�[�h����і��́j��ݒ肷��
'����            �F�Ȃ�
'�߂�l          �F���s���� 0:���� -1:���s
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

	' �s���{���̏���ݒ�
	Call SetPrefectureLocalGovCodeDic(g_objDictionary)

	' 47�s���{���̒n�������c�̂̏���ݒ�
	For i = 1 To 47
		' 2���̐����ɕϊ� ��VBScript�ł�Format�֐��g���Ȃ�
		If i < 10 Then
			prefcode = 0 & i
		Else
			prefcode = i
		End If

		saveFileName = MUNICIPALITIES_PREFIX & prefcode & ".xml"
		targetUrl = MUNICIPALITIES_URL & prefcode

		If SaveResponse(targetUrl, saveFileName) <> 0 Then
			WScript.Echo "�n�������c�́i�R�[�h����і��́j���̎擾�Ɏ��s���܂���"
			Exit For
		End If

		'�񓯊�
		g_objXml.async=False
		' XML�t�@�C���̓ǂݍ���
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

			'�A�z�z��Ɋi�[
			g_objDictionary.add code, name
		Next
	Next

	SetLocalGovCodeDic = 0
End Function

'************************************************************
'�֐�ID          �FgetSupportInformations
'����            �F���x���̕���ID���L�q���ꂽXML�����[�J���ɕۑ�����
'����            �F�Ȃ�
'�߂�l          �FXML�t�@�C���̃t�@�C����
'                �F�L�q����Ă��Ȃ����̂̓J�E���g���Ȃ�
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

		' ID���L�q����Ă��Ȃ����Ƃ𔻒� (�ȈՔ���)
		If InStr(resData, "<IDs/>") > 0 Then
			Exit Do
		End If

		pageNo = pageNo + 1
	LOOP

	GetSupportInformations = pageNo - 1
End Function

'************************************************************
'�֐�ID          �FGetSupportSummaries
'����            �F���x���̊T�v���L�q���ꂽXML�����[�J���ɕۑ�����
'����            �Fids �c �ΏۂƂȂ鐧�x��ID�A�p�C�v��؂�A�ő�100��
'����            �FsaveFileName �c �ۑ��t�@�C����
'�߂�l          �F���s���� 0:���� -1:���s
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
'�֐�ID          �FConvertSummaries2Csv
'����            �F���x���̊T�v��XML�t�@�C����CSV�t�@�C���ɕۑ�
'����            �FxmlFilePath �c XML�t�@�C���p�X
'����            �FobjtextStream �c�o�̓t�@�C���̏o�̓X�g���[��
'�߂�l          �F�Ȃ�
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
	Dim lineData ' �������ލs�f�[�^

	'�񓯊�
	g_objXml.async=False
	' XML�t�@�C���̓ǂݍ���
	result = g_objXml.Load(xmlFilePath)
	If Not result Then
		WScript.Echo g_objXml.parseError.errorCode
		WScript.Echo g_objXml.parseError.reason
		Exit Sub
	End If

	Set nodeRoot = g_objXml.documentElement
	' SupportInformationSummary�^�O�̃��X�g���擾
	Set nodeItems = nodeRoot.getElementsByTagName("SupportInformationSummary")

	' SupportInformationSummary�^�O��1�����o���B
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
'�֐�ID          �FConvertLocalGovCode2LocalGovName
'����            �F�n�������c�̃R�[�h��n�������c�̖��ɕϊ�����
'����            �Fcodes �c �n�������c�̃R�[�h�A�J���}��؂�
'�߂�l          �F�n�������c�̖��ɕϊ����ꂽ������A//��؂�
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
' ���ʏ���
' ===========================================================
'************************************************************
'�֐�ID          �FSaveResponse
'����            �F�w�肳�ꂽURL�̎擾���e���t�@�C���ɕۑ�����
'����            �Furl      �c �擾��URL
'����            �FfilePath �c �t�@�C���p�X
'�߂�l          �F���s���� 0:���� -1:���s
'************************************************************
Function SaveResponse(url, filePath)
	Dim objAdodbStream    ' ADODB.Stream�I�u�W�F�N�g

	SaveResponse = -1

	' �����̃t�@�C�����폜
	if g_objFSO.FileExists(filePath) Then
		g_objFSO.DeleteFile filePath
	End if

	' ��������
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
'�֐�ID          �FReadTextAll
'����            �F�e�L�X�g�t�@�C���̓��e�����ׂēǂݍ��݂܂�
'����            �FfilePath ��� �e�L�X�g�t�@�C���p�X
'�߂�l          �F�t�@�C���̓��e�̃e�L�X�g
'************************************************************
Function ReadTextAll(filePath)
	Dim objTextStream  ' �e�L�X�g�X�g���[���I�u�W�F�N�g
	Dim resData        ' �e�L�X�g���e
	
	Set objTextStream = g_objFSO.OpenTextFile(filePath, 1)
	resData = objTextStream.ReadAll

	objTextStream.Close
	Set objTextStream = Nothing
	ReadTextAll = resData
End Function

'************************************************************
'�֐�ID          �FReadXmlTag
'����            �F�w�肳�ꂽXML�t�@�C���̎w��^�O�̏���Ԃ�
'����            �FxmlFilePath �c XML�t�@�C���p�X
'����            �FtagName �c �^�O��
'�߂�l          �F�w��^�O�̏��
'************************************************************
Function ReadXmlTag(xmlFilePath, tagName)
	Dim nodeRoot
	Dim nodeItems
	Dim nodeItem
	Dim result

	'�񓯊�
	g_objXml.async=False
	' XML�t�@�C���̓ǂݍ���
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
' ���̑�
' ===========================================================
'************************************************************
'�֐�ID          �FSetPrefectureLocalGovCodeDic
'����            �F�A�z�z��ɁA�s���{���̒n�������c�́i�R�[�h����і��́j��ݒ肷��
'����            �Fdic �c �A�z�z��i�[�p��Dictionary�I�u�W�F�N�g
'�߂�l          �F�Ȃ�
'************************************************************
Sub SetPrefectureLocalGovCodeDic(dic)
	'�A�z�z��Ɋi�[ ���L��JIS�Ō��߂�ꂽ�R�[�h�̌n

	dic.add "01000","�k�C��"
	dic.add "02000","�X��"
	dic.add "03000","��茧"
	dic.add "04000","�{�錧"
	dic.add "05000","�H�c��"
	dic.add "06000","�R�`��"
	dic.add "07000","������"
	dic.add "08000","��錧"
	dic.add "09000","�Ȗ،�"
	dic.add "10000","�Q�n��"
	dic.add "11000","��ʌ�"
	dic.add "12000","��t��"
	dic.add "13000","�����s"
	dic.add "14000","�_�ސ쌧"
	dic.add "15000","�V����"
	dic.add "16000","�x�R��"
	dic.add "17000","�ΐ쌧"
	dic.add "18000","���䌧"
	dic.add "19000","�R����"
	dic.add "20000","���쌧"
	dic.add "21000","�򕌌�"
	dic.add "22000","�É���"
	dic.add "23000","���m��"
	dic.add "24000","�O�d��"
	dic.add "25000","���ꌧ"
	dic.add "26000","���s�{"
	dic.add "27000","���{"
	dic.add "28000","���Ɍ�"
	dic.add "29000","�ޗǌ�"
	dic.add "30000","�a�̎R��"
	dic.add "31000","���挧"
	dic.add "32000","������"
	dic.add "33000","���R��"
	dic.add "34000","�L����"
	dic.add "35000","�R����"
	dic.add "36000","������"
	dic.add "37000","���쌧"
	dic.add "38000","���Q��"
	dic.add "39000","���m��"
	dic.add "40000","������"
	dic.add "41000","���ꌧ"
	dic.add "42000","���茧"
	dic.add "43000","�F�{��"
	dic.add "44000","�啪��"
	dic.add "45000","�{�茧"
	dic.add "46000","��������"
	dic.add "47000","���ꌧ"
End Sub

