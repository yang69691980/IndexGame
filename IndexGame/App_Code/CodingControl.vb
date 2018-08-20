Imports System
Imports System.Data
Imports Microsoft.VisualBasic

Public Module CodingControl
	Public Function GetGUID() As String
		Return System.Guid.NewGuid.ToString
	End Function

    Public Function GetIsHttps() As Boolean
        Dim RetValue As Boolean = False

        If String.IsNullOrEmpty(HttpContext.Current.Request.Headers("X-Forwarded-Proto")) = False Then
            If CStr(HttpContext.Current.Request.Headers("X-Forwarded-For")).ToUpper = "HTTPS" Then
                RetValue = True
            End If
        Else
            RetValue = HttpContext.Current.Request.IsSecureConnection
        End If

        Return RetValue
    End Function

    Public Function GetUserIP() As String
        Dim RetValue As String = String.Empty

        If String.IsNullOrEmpty(HttpContext.Current.Request.Headers("X-Forwarded-For")) = False Then
            RetValue = HttpContext.Current.Request.Headers("X-Forwarded-For")
        Else
            RetValue = HttpContext.Current.Request.UserHostAddress
        End If

        Return RetValue
    End Function

    Public Function IsEmpty(ByVal CheckVariant As Object) As Boolean
		Dim RetValue As Boolean

		RetValue = True
		If IsNothing(CheckVariant) = False Then
			If IsDBNull(CheckVariant) = False Then
				If VarType(CheckVariant) = VariantType.String Then
					If CheckVariant <> String.Empty Then
						RetValue = False
					End If
				Else
					RetValue = False
				End If
			End If
		End If

		Return RetValue
	End Function

    Public Function RandomPassword(ByVal MaxPasswordChars As Integer) As String
        Dim R2 As New Random(Timer)

        Return RandomPassword(R2, MaxPasswordChars)
    End Function

    Public Function RandomPassword(R As Random, ByVal MaxPasswordChars As Integer) As String
        Dim PasswordString As String

        PasswordString = "1234567890ABCDEFGHJKLMNPQRSTUVWXYZ"

        Return RandomPassword(R, MaxPasswordChars, PasswordString)
    End Function

    Public Function RandomPassword(R As Random, ByVal MaxPasswordChars As Integer, AvailableCharList As String) As String
        Dim I As Integer
        Dim CharIndex As Integer
        Dim PasswordString As String
        Dim RetValue As String

        RetValue = String.Empty
        PasswordString = AvailableCharList
        For I = 1 To MaxPasswordChars
            CharIndex = R.Next(0, PasswordString.Length - 1)

            RetValue = RetValue & PasswordString.Chars(CharIndex)
        Next

        Return RetValue
    End Function

    Public Function GetQueryString() As String
        Dim QueryString As String
        Dim QueryStringIndex As Integer

        QueryStringIndex = Request.RawUrl.IndexOf("?")
        QueryString = String.Empty
        If QueryStringIndex > 0 Then
            QueryString = Request.RawUrl.Substring(QueryStringIndex + 1)
        End If

        Return QueryString
    End Function

	Public Function GetURLFilename(ByVal S)
		' 傳出檔名(無 http:// 開頭的 URL 檔名)
        Dim TempInt As Integer
        Dim TempStr As String
        Dim I As Integer

		TempStr = S

		' 找尋是否有 ? (參數)
		TempInt = InStr(TempStr, "?")
		If TempInt <> 0 Then
			TempStr = Left(TempStr, TempInt - 1)
		End If

		' 找尋是否有 "/" 或 "\"
		For I = Len(TempStr) To 1 Step -1
			If (Mid(TempStr, I, 1) = "/") Or (Mid(TempStr, I, 1) = "\") Then
				TempStr = Mid(TempStr, I + 1)
				Exit For
			End If
		Next

		Return TempStr
	End Function

	Public Function FormSubmit() As Boolean
		If Request.HttpMethod.Trim.ToUpper = "POST" Then
			Return True
		Else
			Return False
		End If
	End Function

	Public Function GetDefaultLanguage() As String
		' 取得使用者的語言
		' 傳回: 字串, 代表使用者預設的語言集
		Dim LangArr() As String
		Dim Temp As String
		Dim TempArr() As String
		Dim RetValue As String

		Temp = Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")
		TempArr = Temp.Split(";")

		LangArr = TempArr(0).Split(",")

		If LangArr(0).Trim = String.Empty Then
			RetValue = "en-us"
		Else
			RetValue = LangArr(0)
		End If

		Return RetValue
	End Function

	Public Function EMailCheck(ByVal EMail As String) As Boolean
		' 驗證 EMail Address 是否允許使用
		Dim atIndex As Integer
		Dim TempEMail As String
		Dim AccountName As String
		Dim DomainName As String
		Dim DomainArray() As String
		Dim CheckAllow As Boolean
		Dim AccountAllowChar As String
		Dim DomainAllowChar As String
		Dim TempChar As String
		Dim NotFound As Boolean
		Dim I As Integer
		Dim J As Integer

		CheckAllow = False
		AccountAllowChar = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ~%&-_+=\/.:"
		DomainAllowChar = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-_"
		TempEMail = EMail

		atIndex = TempEMail.IndexOf("@")
		If atIndex <> -1 Then
			' 解析帳號與 Domain, 同時針對 Domain 分割
			AccountName = TempEMail.Substring(0, atIndex)
			DomainName = TempEMail.Substring(atIndex + 1)
			DomainArray = DomainName.Split(".")

			' 驗證帳號正確性
			NotFound = False
			If AccountName <> String.Empty Then
				For I = 0 To AccountName.Length - 1
					TempChar = AccountName.Chars(I)
					If AccountAllowChar.IndexOf(TempChar) = -1 Then
						NotFound = True
						Exit For
					End If
				Next

				If NotFound = False Then
					' 驗證 Domain 正確性
					If IsNothing(DomainArray) = False Then
						If DomainArray.Length >= 0 Then
							For I = 0 To DomainArray.Length - 1
								If DomainArray(I).Length > 0 Then
									For J = 0 To DomainArray(I).Length - 1
										TempChar = DomainArray(I).Chars(J)
										If DomainAllowChar.IndexOf(TempChar) = -1 Then
											NotFound = True
											Exit For
										End If
									Next
								Else
									NotFound = True
								End If
							Next
						Else
							NotFound = True
						End If
					Else
						NotFound = True
					End If

					If NotFound = False Then
						CheckAllow = True
					End If
				End If
			End If
		End If

		Return CheckAllow
	End Function

	Public Function PIDCheck(ByVal PID As String) As Boolean
		' 使用者身份證驗證
		' 輸入: 身分證 String
		' 輸出: Boolean (True=驗證成功/False=驗證失敗)
		Dim PIDValue() As Integer
		Dim PIDHead() As String
		Dim PIDPosition As Integer
		Dim HeadStr As String
		Dim HeadHigh As Integer
		Dim HeadLow As Integer
		Dim CheckSum As Integer
		Dim SumValue As Integer
		Dim CheckPass As Boolean
		Dim I As Integer

		CheckPass = False
		If PID.Length = 10 Then
			PIDValue = New Integer() {10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35}
			PIDHead = New String() {"A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "L", "M", "N", "P", "Q", "R", "S", "T", "U", "V", "X", "Y", "W", "Z", "I", "O"}

			HeadStr = PID.Chars(0)
			CheckSum = Val(PID.Substring(PID.Length - 1, 1))
			PIDPosition = -1
			For I = 0 To PIDHead.Length - 1
				If PIDHead(I) = HeadStr.ToUpper Then
					PIDPosition = I
					Exit For
				End If
			Next

			If PIDPosition > -1 Then
				HeadHigh = PIDValue(PIDPosition) \ 10
				HeadLow = PIDValue(PIDPosition) Mod 10

				SumValue = HeadHigh + 9 * HeadLow
				For I = 1 To 8
					SumValue = SumValue + ((9 - I) * Val(PID.Chars(I)))
				Next

				SumValue = SumValue + CheckSum
				If (SumValue Mod 10) = 0 Then
					CheckPass = True
				End If
			End If
		End If

		Return CheckPass
	End Function

	Public Function JSEncodeString(ByVal Content As String) As String
		' ASP 產生 JavaScript 字串轉換函式
		' 避免雙引號或反斜線輸出造成前端錯誤
		' 輸入: String
		' 輸出: String
		Dim TEMP As String
		Dim TEMP2 As String
		Dim TEMPStr As String

		If Content <> String.Empty Then
			TEMP = Content
			TEMP2 = ""
			For Each TEMPStr In TEMP
				Select Case TEMPStr
					Case "\"
						TEMPStr = "\x5c"
					Case """"
						TEMPStr = "\x22"
					Case Chr(13)
						TEMPStr = "\n"
					Case Chr(10)
						TEMPStr = ""
				End Select

				TEMP2 = TEMP2 & TEMPStr
			Next
		End If

		Return TEMP2
	End Function

	Public Function LogonDomainUser() As String
		' 取得 NT 登入使用者+Domain
		' 輸出: String
		Return Request.ServerVariables("LOGON_USER")
	End Function

	Public Function LogonUser() As String
		' 取得 NT 登入使用者
		' 輸出: String
		Dim Temp As String
		Dim TempInt As Integer

		Temp = Request.ServerVariables("LOGON_USER")

		TempInt = Temp.IndexOf("\")
		If TempInt <> -1 Then
			Temp = Temp.Substring(TempInt + 1).Trim
		End If

		Return Temp
	End Function

	Public Function LogonDomain() As String
		' 取得 NT 登入 Domain
		' 輸出: String
		Dim Temp As String
		Dim TempInt As Integer

		Temp = Request.ServerVariables("LOGON_USER")

		TempInt = Temp.IndexOf("\")
		If TempInt <> -1 Then
			Temp = Temp.Substring(0, TempInt)
		End If

		Return Temp
	End Function

	Public Function StrUnicode(ByVal S As String) As String
		' 標準字串轉換為網頁 Unicode 輸出
		Dim TempChar As String
		Dim TempStr As String
		Dim I As Integer

		TempStr = String.Empty
		For I = 0 To S.Length - 1
			TempChar = S.Chars(I)

			If (Asc(TempChar) > 255) Or (Asc(TempChar) < 0) Then
				TempStr = TempStr & "&#" & AscW(TempChar) & ";"
			Else
				TempStr = TempStr & TempChar
			End If
		Next

		Return TempStr
	End Function

	Public Function UnicodeStr(ByVal S As String) As Integer
		' 網頁 Unicode 轉換標準字串輸出
		' 輸入: 網頁 Unicode
		' 輸出: String
		Dim TempStr As String
		Dim Temp As String
		Dim TempInt As Integer
		Dim TempInt2 As Integer
		Dim UniStr As String
		Dim UnicodeChar As String

		TempStr = S
		Temp = String.Empty
		Do
			If TempStr = "" Then Exit Do
			TempInt = TempStr.IndexOf("&#")
			If TempInt <> -1 Then
				TempInt2 = TempStr.IndexOf(";", TempInt + 1)
				If TempInt2 <> -1 Then
					UniStr = TempStr.Substring(TempInt + 2, TempInt2 - TempInt - 2).Trim
				Else
					UniStr = TempStr.Substring(TempInt + 2).Trim
				End If

				If CStr(UniStr.Chars(0)).ToUpper = "X" Then
					UnicodeChar = Val("&H" & UniStr.Substring(1))
				Else
					UnicodeChar = Val(UniStr)
				End If

				Temp = Temp & TempStr.Substring(0, TempInt) & ChrW(UnicodeChar)
				TempStr = TempStr.Substring(TempInt2 + 1)
			Else
				Temp = Temp & TempStr
				TempStr = ""
			End If
		Loop

		UnicodeStr = Temp
    End Function

	Public Sub SendMail(ByVal SMTPServer As String, ByVal FromName As String, ByVal ToName As String, ByVal Subject As String, ByVal Body As String, ByVal Charset As String)
		Call SendMail(SMTPServer, New System.Net.Mail.MailAddress(FromName), New System.Net.Mail.MailAddress(ToName), Subject, Body, Charset)
	End Sub

	Public Sub SendMail(ByVal SMTPServer As String, ByVal FromName As System.Net.Mail.MailAddress, ByVal ToName As System.Net.Mail.MailAddress, ByVal Subject As String, ByVal Body As String, ByVal Charset As String)
		Dim SMTP As System.Net.Mail.SmtpClient
		Dim Msg As System.Net.Mail.MailMessage

		SMTP = New System.Net.Mail.SmtpClient(SMTPServer)
		Msg = New System.Net.Mail.MailMessage
		Msg.Body = Body

		Msg.IsBodyHtml = True

		If Charset = String.Empty Then
			Msg.BodyEncoding = System.Text.Encoding.Default
		Else
			Msg.BodyEncoding = System.Text.Encoding.GetEncoding(Charset)
		End If

		Msg.From = FromName
        Msg.To.Add(ToName)
		Msg.Subject = Subject

        Try
            SMTP.Send(Msg)
        Catch ex As Exception

        End Try

		Msg = Nothing
	End Sub

	Public Function GetWebBinaryContent(ByVal URL As String) As Byte()
		Dim HttpContent() As Byte
		Dim HttpClient As System.Net.WebClient

		HttpClient = New System.Net.WebClient
		HttpContent = HttpClient.DownloadData(URL)

		Return HttpContent
	End Function

	Public Function GetWebTextContent(ByVal URL As String, Optional ByVal Method As String = "GET", Optional ByVal SendData As String = "") As String
		Dim HttpClient As System.Net.HttpWebRequest
		Dim HttpResponse As System.Net.HttpWebResponse
		Dim Stm As System.IO.Stream
		Dim SR As System.IO.StreamReader
		Dim RetValue As String
		Dim SendBytes() As Byte

		System.Net.ServicePointManager.CertificatePolicy = New TrustAllCertificatePolicy

		HttpClient = System.Net.HttpWebRequest.Create(URL)
		HttpClient.Method = Method
		HttpClient.Accept = "*/*"
		HttpClient.UserAgent = "Sender"
		HttpClient.KeepAlive = False

		Select Case Method.ToUpper
			Case "POST"
				SendBytes = System.Text.Encoding.Default.GetBytes(SendData)

				HttpClient.ContentType = "application/x-www-form-urlencoded"
				HttpClient.ContentLength = SendBytes.Length
				HttpClient.GetRequestStream.Write(SendBytes, 0, SendBytes.Length)
		End Select

		Try
			HttpResponse = HttpClient.GetResponse()
		Catch ex As System.Net.WebException
			HttpResponse = ex.Response
		End Try

		Stm = HttpResponse.GetResponseStream
		SR = New System.IO.StreamReader(Stm)
		RetValue = SR.ReadToEnd

		Stm.Close()

		Try
			HttpResponse.Close()
		Catch ex As System.Net.WebException
		End Try

		HttpClient = Nothing

		Return RetValue
	End Function

    Public Function MyPageBasicRoot() As String
        Dim Path As String
        Dim RetValue As String

        Path = MyPagePath()

        If Path.Substring(Path.LastIndexOf("/") + 1).ToUpper = "MAINTAIN" Then
            RetValue = Path.Substring(0, Path.LastIndexOf("/"))
        Else
            RetValue = Path
        End If

        Return RetValue
    End Function

    Public Function MyServerName() As String
        ' 取得目前的主機電腦名稱
        MyServerName = Request.ServerVariables("SERVER_NAME")
    End Function

	Public Function MyPageRoot() As String
		Dim RetValue As String

		RetValue = Request.ApplicationPath

		If RetValue.Substring(RetValue.Length - 1, 1) = "/" Then
			If RetValue.Length = 1 Then
				RetValue = String.Empty
			Else
				RetValue = RetValue.Substring(0, RetValue.Length - 1)
			End If
		End If

		Return RetValue
	End Function

	Public Function MyPagePath() As String
		' 取得目前網頁的路徑
		Dim Tmp As String
		Dim I As Integer
		Dim RetValue As String

		Tmp = MyPageURL()
		I = Tmp.LastIndexOf("/")
		If I = -1 Then
			I = Tmp.LastIndexOf("\")
		End If

		If I <> -1 Then
			RetValue = Tmp.Substring(0, I)
		Else
			RetValue = Tmp
		End If

		Return RetValue
	End Function

	Public Function MyPage() As String
		Dim RetValue As String
		Dim TmpIndex As Integer

		' 取得目前的網頁路徑+網頁名稱(不含參數)
		RetValue = MyPageURL()
		TmpIndex = RetValue.IndexOf("?")
		If TmpIndex <> -1 Then
			RetValue = RetValue.Substring(0, TmpIndex)
		End If

		Return RetValue
	End Function

	Public Function MyPageURL() As String
		' 取得目前網頁路徑+網頁名稱+參數
		Return Request.Url.ToString
	End Function

	Public Function MyPageRootURL() As String
		Dim HttpTag As String
		Dim ServerName As String
		Dim Path As String
		Dim RetValue As String

		If Request.IsSecureConnection Then
			HttpTag = "https://"
		Else
			HttpTag = "http://"
		End If

		If (CInt(Request.ServerVariables("SERVER_PORT")) = 80) Or (IsEmpty(Request.ServerVariables("SERVER_PORT"))) Then
			ServerName = MyServerName()
		Else
			ServerName = MyServerName() & ":" & Request.ServerVariables("SERVER_PORT")
		End If

		Path = MyPageRoot()

		RetValue = HttpTag & ServerName & "/" & Path

		Return RetValue
	End Function

	Public Function MyPageParam() As String
        Dim Param As String
        Dim URL As String

		URL = Request.RawUrl
		Param = URL.Substring(URL.IndexOf("?") + 1)

		Return Param
	End Function

    Public Function MyPageName() As String
        Dim PageName As String
        Dim URL As String
        Dim TmpIndex As Integer

        URL = Request.RawUrl
        TmpIndex = URL.LastIndexOf("/")
        If TmpIndex <> -1 Then
            URL = URL.Substring(TmpIndex + 1)
        End If

        TmpIndex = URL.IndexOf("?")
        If TmpIndex <> -1 Then
            PageName = URL.Substring(0, TmpIndex)
        Else
            PageName = URL
        End If

        Return PageName
    End Function

	Public Function ASPString(ByVal SS As String) As String
		' 字串轉換函式
		' 自動轉換 CrLf 成為 HTML 斷行(<br>)
        Return SS.Replace(vbLf, String.Empty).Replace(vbCr, "<br>")
    End Function

    Public Function NarrowConv(ByVal SS As String) As String
        Return StrConv(SS, VbStrConv.Narrow)
    End Function

    Public Function UserIP() As String
        ' 取得使用者的 IP Address
        Return Request.UserHostAddress
    End Function

    Public Function GetFilename(ByVal S As String) As String
        ' 取得檔案名稱
        ' 輸入: 包含路徑的檔名
        ' 輸出: 純檔名(包含副檔名)
        Dim TempStr As String
        Dim TempInt As Integer

        TempStr = ""
        TempInt = S.Replace("/", "\").LastIndexOf("\")
        If TempInt <> -1 Then
            TempStr = S.Substring(TempInt + 1)
        Else
            TempStr = S
        End If

        Return TempStr
    End Function

    Public Function GetFileExtName(ByVal S As String) As String
        ' 取得副檔名名稱
        ' 輸入: 檔案名稱
        ' 輸出: 純副檔名
        Dim TempStr As String
        Dim TempInt As Integer

        TempStr = String.Empty
        TempInt = S.LastIndexOf(".")
        If TempInt <> -1 Then
            TempStr = S.Substring(TempInt + 1)
        End If

        Return TempStr
    End Function

    Public Function GetStringLength(ByVal S As String) As Integer
        Return System.Text.Encoding.Default.GetByteCount(S)
    End Function

    'Public Sub Redirect(ByVal DestURL As String, Optional ByVal ResponseEnd As Boolean = True)
    '    If DestURL = String.Empty Then
    '        DestURL = "/"
    '    End If

    '    If System.EnterpriseServices.ContextUtil.IsInTransaction = True Then
    '        ' In Transaction
    '        ResponseEnd = False
    '    End If

    '    Response.Clear()

    '    Response.Write("<html>")
    '    Response.Write("<head>")
    '    Response.Write("<meta http-equiv=""Refresh"" content=""0; URL=" & DestURL & """>")
    '    Response.Write("</head>")
    '    Response.Write("<script language=javascript>")
    '    Response.Write("function refreshPage() {")
    '    Response.Write("    window.location.href=""" & DestURL & """;")
    '    Response.Write("}")
    '    Response.Write("</script>")
    '    Response.Write("<body onload=""refreshPage();"">")
    '    Response.Write("</body>")
    '    Response.Write("</html>")

    '    If ResponseEnd Then
    '        Response.End()
    '    Else
    '        Response.Flush()
    '        Response.SuppressContent = True
    '        Response.Close()
    '    End If
    'End Sub

    Public Function XMLDeserial(ByVal xmlContent As String, ByVal objType As Type) As Object
        Dim XMLSer As System.Xml.Serialization.XmlSerializer
        Dim Stm As System.IO.MemoryStream
        Dim XMLArray() As Byte
        Dim RetValue As Object = Nothing

        If xmlContent <> String.Empty Then
            XMLArray = System.Text.Encoding.UTF8.GetBytes(xmlContent)

            Stm = New System.IO.MemoryStream
            Stm.Write(XMLArray, 0, XMLArray.Length)
            Stm.Position = 0
            XMLSer = New Xml.Serialization.XmlSerializer(objType)

            RetValue = XMLSer.Deserialize(Stm)

            Stm.Dispose()
            Stm = Nothing
        End If

        Return RetValue
    End Function

    Public Function GetMD5(ByVal DataString As String) As String
        Return GetMD5(System.Text.Encoding.UTF8.GetBytes(DataString))
    End Function

    Public Function GetMD5(ByVal Data() As Byte) As String
        Dim MD5Provider As New System.Security.Cryptography.MD5CryptoServiceProvider
        Dim RetValue() As Byte
        Dim RetValue2 As String = String.Empty

        RetValue = MD5Provider.ComputeHash(Data)
        MD5Provider = Nothing
        RetValue2 = System.Convert.ToBase64String(RetValue)

        Return RetValue2
    End Function

    Private Function Request() As System.Web.HttpRequest
        Return HttpContext.Current.Request
    End Function

    Private Function Response() As System.Web.HttpResponse
        Return HttpContext.Current.Response
    End Function

    Private Function Server() As System.Web.HttpServerUtility
        Return HttpContext.Current.Server
    End Function

    Private Function Application() As System.Web.HttpApplicationState
        Return HttpContext.Current.Application
    End Function

    Private Function Session() As System.Web.SessionState.HttpSessionState
        Return HttpContext.Current.Session
    End Function
End Module

Public Class TrustAllCertificatePolicy
	Implements System.Net.ICertificatePolicy

	Public Function CheckValidationResult(ByVal srvPoint As System.Net.ServicePoint, ByVal certificate As System.Security.Cryptography.X509Certificates.X509Certificate, ByVal request As System.Net.WebRequest, ByVal certificateProblem As Integer) As Boolean Implements System.Net.ICertificatePolicy.CheckValidationResult
		Return True
	End Function
End Class