Imports System
Imports System.Data
Imports System.Security.Cryptography

Public Module Platform
    Public DBConnStr As String = System.Configuration.ConfigurationManager.ConnectionStrings("DBConnStr").ConnectionString
    Public DateTimeNull As DateTime = CDate("1900/1/1")
    Public ServerKey As String = System.Configuration.ConfigurationManager.AppSettings("ServerKey")
    Public MyDomainURL As String = System.Configuration.ConfigurationManager.AppSettings("MyDomainURL")
    Public IsTestSite As Boolean = System.Configuration.ConfigurationManager.AppSettings("IsTestSite")
    Public RedisServerIP As String = System.Configuration.ConfigurationManager.AppSettings("RedisServerIP")
    Public RedisServerPort As Integer = System.Configuration.ConfigurationManager.AppSettings("RedisServerPort")
    Public RedisServerSSL As Boolean = System.Configuration.ConfigurationManager.AppSettings("RedisServerSSL")
    Public RedisServerKey As String = System.Configuration.ConfigurationManager.AppSettings("RedisServerKey")
    Public SharedFolder As String = System.Configuration.ConfigurationManager.AppSettings("SharedFolder")
    Public Key3DES As String = "onoeTs39aHfAATKGxYmyJ3Nf"
    Public DirSplit As String = "\"

    Public RedisServer As StackExchange.Redis.IServer

    Private RedisClient As StackExchange.Redis.ConnectionMultiplexer
    Private Rnd As New Random
    Private SyncRoot As New ArrayList

    Public Function ChkPermission(ByVal PermissionName As String) As Boolean
        Dim RetValue As Boolean = False
        Dim SS As String = String.Empty
        Dim DBCmd As System.Data.SqlClient.SqlCommand
        Dim DT As System.Data.DataTable
        Dim Admin As AdminLoginState = Nothing

        Admin = GetAdminLogin()

        SS = "Select AdminRolePermission.* From AdminRolePermission With (Nolock) " &
        " Left Join AdminTable With (Nolock) On AdminRolePermission.forAdminRoleID = AdminTable.forAdminRoleID" &
        " Where AdminTable.AdminID = @AdminID And AdminRolePermission.forPermissionName = @forPermissionName"
        DBCmd = New System.Data.SqlClient.SqlCommand
        DBCmd.CommandText = SS
        DBCmd.CommandType = Data.CommandType.Text
        DBCmd.Parameters.Add("@AdminID", Data.SqlDbType.Int).Value = Admin.AdminID
        DBCmd.Parameters.Add("@forPermissionName", Data.SqlDbType.VarChar).Value = PermissionName
        DT = GetDB(DBConnStr, DBCmd)

        If DT.Rows.Count > 0 Then
            RetValue = True
        End If

        Return RetValue
    End Function

    Public Function GetAdminLogin() As AdminLoginState
        Dim Admin As AdminLoginState = Nothing

        If IsNothing(Session("PlatformAdmin_Logined")) = False Then
            If CBool(Session("PlatformAdmin_Logined")) Then
                Admin = Session("PlatformAdmin")
            End If
        End If

        Return Admin
    End Function

    Public Sub SetAdminLogin(Admin As AdminLoginState)
        Session("PlatformAdmin_Logined") = True
        Session("PlatformAdmin") = Admin
    End Sub

    Public Function LoginAdmin(AdminAccount As String, AdminPassword As String, CompanyCode As String) As AdminLoginState
        Dim SS As String
        Dim DBCmd As System.Data.SqlClient.SqlCommand
        Dim DT As System.Data.DataTable
        Dim RetValue As AdminLoginState
        Dim GameToken As String = String.Empty

        SS = "SELECT * FROM AdminTable WITH (NOLOCK) WHERE LoginAccount=@LoginAccount" &
            " And forCompanyID=(Select CompanyID From CompanyTable With(Nolock) Where CompanyCode=@CompanyCode)"
        DBCmd = New System.Data.SqlClient.SqlCommand
        DBCmd.CommandText = SS
        DBCmd.CommandType = Data.CommandType.Text
        DBCmd.Parameters.Add("@LoginAccount", Data.SqlDbType.VarChar).Value = AdminAccount
        DBCmd.Parameters.Add("@CompanyCode", Data.SqlDbType.VarChar).Value = CompanyCode
        DT = GetDB(DBConnStr, DBCmd)
        If DT.Rows.Count > 0 Then
            If GetMD5(AdminPassword) = DT.Rows(0)("LoginPassword") Then
                GameToken = GetGameToken(DT.Rows(0)("forCompanyID"))

                RetValue = New AdminLoginState
                RetValue.AdminAccount = DT.Rows(0)("LoginAccount")
                RetValue.AdminID = DT.Rows(0)("AdminID")
                RetValue.RealName = DT.Rows(0)("RealName")
                RetValue.forCompanyID = DT.Rows(0)("forCompanyID")
                RetValue.GameToken = GameToken
            End If
        End If

        Return RetValue
    End Function



    Public Function BuildPersonCode(UserAccountID As Integer) As String
        Dim V As String
        Dim V2 As String

        V2 = New String("0", 8 - CStr(UserAccountID).Length) & CStr(UserAccountID)

        ' yyyyMMddHHmmss+UserID
        V = (Now.Second Mod 10) & V2

        Return Code36(V)
    End Function

    Public Function Code36(Value As UInteger) As String
        Dim RetValue As String = String.Empty
        Dim ModValue As UInteger
        Dim FreeValue As UInteger = Value
        Dim MapChars As String = "0123456789ABCDEFGHJKLMNPQRSTUVWXYZ"

        Do
            If FreeValue <= 0 Then Exit Do

            ModValue = FreeValue Mod 34
            FreeValue = FreeValue \ 34

            RetValue &= MapChars.Chars(ModValue)
        Loop

        Return RetValue
    End Function

    Public Function ChineseNumberCheck(S As String) As String
        Dim TmpString As New StringBuilder

        If String.IsNullOrEmpty(S) = False Then
            For Each EachStr As String In S
                Select Case EachStr
                    Case "＋"
                        TmpString.Append("+")
                    Case "－"
                        TmpString.Append("-")
                    Case "（"
                        TmpString.Append("(")
                    Case "）"
                        TmpString.Append(")")
                    Case "　"
                        TmpString.Append(" ")
                    Case "１", "一"
                        TmpString.Append("1")
                    Case "２", "二"
                        TmpString.Append("2")
                    Case "３", "三"
                        TmpString.Append("3")
                    Case "４", "四"
                        TmpString.Append("4")
                    Case "５", "五"
                        TmpString.Append("5")
                    Case "６", "六"
                        TmpString.Append("6")
                    Case "７", "七"
                        TmpString.Append("7")
                    Case "８", "八"
                        TmpString.Append("8")
                    Case "９", "九"
                        TmpString.Append("9")
                    Case "０", "零"
                        TmpString.Append("0")
                    Case Else
                        TmpString.Append(EachStr)
                End Select
            Next
        End If

        Return TmpString.ToString
    End Function

    Public Function GetRedisServerTime() As DateTime
        Dim RetValue As DateTime
        Dim Success As Boolean = False

        RedisPrepare()

        If IsNothing(RedisServer) = False Then
            For I As Integer = 1 To 3
                Try
                    RetValue = RedisServer.Time()
                    Success = True
                    Exit For
                Catch ex As Exception

                End Try
            Next
        End If

        If Success Then
            Return RetValue.ToLocalTime
        Else
            Return Now
        End If
    End Function

    Public Function GetRedisClient(Optional db As Integer = -1) As StackExchange.Redis.IDatabase
        Dim RetValue As StackExchange.Redis.IDatabase

        RedisPrepare()

        If db = -1 Then
            RetValue = RedisClient.GetDatabase()
        Else
            RetValue = RedisClient.GetDatabase(db)
        End If

        Return RetValue
    End Function

    Public Function DSFindTableByName(DS As System.Data.DataSet, TableName As String) As System.Data.DataTable
        Dim RetValue As System.Data.DataTable = Nothing

        For Each EachTable As System.Data.DataTable In DS.Tables
            If String.IsNullOrEmpty(EachTable.TableName) = False Then
                If EachTable.TableName.Trim.ToUpper = TableName.Trim.ToUpper Then
                    RetValue = EachTable
                    Exit For
                End If
            End If
        Next

        Return RetValue
    End Function

    Public Function RandomCreator(Min As Integer, Max As Integer) As Integer
        Dim RetValue As Integer

        SyncLock SyncRoot
            RetValue = Rnd.Next(Min, Max)
        End SyncLock

        Return RetValue
    End Function

    Public Function GetString(s As Object) As String
        If IsDBNull(s) = False Then
            If String.IsNullOrEmpty(s) = False Then
                Return s
            Else
                Return String.Empty
            End If
        Else
            Return String.Empty
        End If
    End Function

    Public Function Encrypt3DES(Content As String, Key As String) As String
        Return Encrypt3DES(System.Text.Encoding.Default.GetBytes(Content), Key)
    End Function

    Public Function Encrypt3DES(Content() As Byte, Key As String) As String
        Dim DES As New System.Security.Cryptography.TripleDESCryptoServiceProvider
        Dim DESEncrypt As System.Security.Cryptography.ICryptoTransform

        DES.Key = System.Text.Encoding.UTF8.GetBytes(Key)
        DES.Mode = CipherMode.ECB

        DESEncrypt = DES.CreateEncryptor
        Return System.Convert.ToBase64String(DESEncrypt.TransformFinalBlock(Content, 0, Content.Length))
    End Function

    Public Function Decrypt3DES(EncStr As String, Key As String) As Byte()
        Dim DES As New System.Security.Cryptography.TripleDESCryptoServiceProvider
        Dim DESDecrypt As System.Security.Cryptography.ICryptoTransform
        Dim SrcContent() As Byte

        DES.Key = System.Text.Encoding.UTF8.GetBytes(Key)
        DES.Mode = CipherMode.ECB
        DES.Padding = PaddingMode.PKCS7

        DESDecrypt = DES.CreateDecryptor

        SrcContent = System.Convert.FromBase64String(EncStr)
        Return DESDecrypt.TransformFinalBlock(SrcContent, 0, SrcContent.Length)
    End Function
    Function GetGameToken(ByVal CompanyID As Integer) As String
        Dim SS As String = String.Empty
        Dim DBCmd As System.Data.SqlClient.SqlCommand
        Dim DT As System.Data.DataTable
        Dim rnd As New Random()
        Dim chars As Char() = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim myLength As Integer = rnd.Next(4, 4)
        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder
        Dim Rdn As String = String.Empty
        Dim myHash As String = String.Empty
        Dim myToken As String = String.Empty
        Dim oSha1 As System.Security.Cryptography.SHA1Managed = New System.Security.Cryptography.SHA1Managed
        Dim bytData As Byte()
        Dim bytResult As Byte()
        Dim RetValue As String = String.Empty
        Dim CompanyCodeLength As String

        If CompanyID <> 0 Then
            SS = "Select * From CompanyTable With (Nolock) Where CompanyID=@CompanyID"
            DBCmd = New System.Data.SqlClient.SqlCommand
            DBCmd.CommandText = SS
            DBCmd.CommandType = Data.CommandType.Text
            DBCmd.Parameters.Add("@CompanyID", Data.SqlDbType.Int).Value = CompanyID
            DT = GetDB(DBConnStr, DBCmd)

            If DT.Rows.Count > 0 Then
                If String.IsNullOrEmpty(DT.Rows(0)("PrivateKey")) = False Then
                    If String.IsNullOrEmpty(DT.Rows(0)("CompanyCode")) = False Then
                        For i As Integer = 0 To myLength - 1
                            sb.Append(chars(rnd.Next(chars.Length - 1)))
                        Next

                        Rdn = sb.ToString
                    End If

                    myHash = Rdn & ":" & DT.Rows(0)("CompanyCode") & ":" & DT.Rows(0)("PrivateKey")
                    bytData = System.Text.Encoding.Default.GetBytes(myHash)
                    bytResult = oSha1.ComputeHash(bytData)
                    myToken = System.Convert.ToBase64String(bytResult)
                    CompanyCodeLength = DT.Rows(0)("CompanyCode").ToString.Length
                    CompanyCodeLength = CompanyCodeLength.PadLeft(2, "0")
                    RetValue = CompanyCodeLength & Rdn & DT.Rows(0)("CompanyCode") & myToken
                    'Response.Write(Rdn & DT.Rows(0)("CompanyCode") & myToken)
                End If
            End If
        End If

        Return RetValue
    End Function

    Public Function FormatDecimal(s As Decimal) As Decimal
        Dim iValue As Decimal
        Dim LeftValue As Decimal
        Dim i As Integer = 1
        Dim s2 As Decimal
        Dim IsNegative As Boolean = False

        If s < 0 Then
            IsNegative = True
        End If

        s2 = Math.Abs(s)

        iValue = Math.Floor(s2) \ 1
        LeftValue = s2 Mod 1

        Do
            Dim tmpValue As Decimal

            tmpValue = (LeftValue * (10 ^ i) Mod 1)
            If tmpValue = 0 Then
                iValue += (LeftValue * (10 ^ i)) * (10 ^ -i)
                Exit Do
            Else
                i += 1
            End If
        Loop

        If IsNegative Then
            Return 0 - iValue
        Else
            Return iValue
        End If
    End Function

    Public Function IsDateTimeNull(s As Object) As Boolean
        If IsNothing(s) = False Then
            Return IsDateTimeNull(CDate(s))
        Else
            Return True
        End If
    End Function

    Public Function IsDateTimeNull(s As String) As Boolean
        If String.IsNullOrEmpty(s) = False Then
            Return IsDateTimeNull(CDate(s))
        Else
            Return True
        End If
    End Function

    Public Function IsDateTimeNull(s As DateTime) As Boolean
        Dim RetValue As Boolean = False

        If s.Equals(Nothing) = False Then
            If s.Equals(DateTimeNull) Then
                RetValue = True
            End If
        Else
            RetValue = True
        End If

        Return RetValue
    End Function

    Public Sub AlertMessage(ByVal Content As String, ByVal ReturnURL As String)
        Response.ClearContent()
        Response.Write("<Script Language=""JavaScript"">")
        Response.Write("alert(""" & JSEncodeString(Content) & """);")

        If ReturnURL = String.Empty Then
            Response.Write("window.history.back (1);")

        Else
            Response.Write("window.location.href=""" & ReturnURL & """;")
        End If

        Response.Write("</Script>")
        Response.Flush()
        Response.End()
    End Sub

    Public Function XMLSerial(ByVal obj As Object) As String
        Dim XMLSer As System.Xml.Serialization.XmlSerializer
        Dim Stm As System.IO.MemoryStream
        Dim XMLArray() As Byte
        Dim RetValue As String

        XMLSer = New Xml.Serialization.XmlSerializer(obj.GetType)
        Stm = New System.IO.MemoryStream
        XMLSer.Serialize(Stm, obj)

        Stm.Position = 0

        ReDim XMLArray(Stm.Length - 1)
        Stm.Read(XMLArray, 0, XMLArray.Length)
        Stm.Dispose()
        Stm = Nothing

        RetValue = System.Text.Encoding.UTF8.GetString(XMLArray)

        Return RetValue
    End Function

    Private Sub RedisPrepare()
        If IsNothing(RedisClient) Then
            Dim RedisConnectionString As String

            'td888.redis.cache.windows.net:6380,password=NNgtf6X9eX6FJILCUymHFrCZNeUQVxcGhhuhQcJUOsg=,ssl=True,abortConnect=False
            RedisConnectionString = RedisServerIP & ":" & RedisServerPort & ",abortConnect=False"

            If String.IsNullOrEmpty(RedisServerKey) = False Then
                RedisConnectionString &= ",password=" & RedisServerKey
            End If

            If RedisServerSSL Then
                RedisConnectionString &= ",ssl=True"
            Else
                RedisConnectionString &= ",ssl=False"
            End If


            RedisClient = StackExchange.Redis.ConnectionMultiplexer.Connect(RedisConnectionString)
        End If

        If IsNothing(RedisServer) Then
            RedisServer = RedisClient.GetServer(RedisServerIP, RedisServerPort)
        End If
    End Sub

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

    Private Function User() As System.Security.Principal.IPrincipal
        Return HttpContext.Current.User
    End Function

    Public Class PlatformSession
        Public Shared Property Item(Key As String) As Object
            Get
                Dim RetValue As Object = Nothing

                If PlatformRedisCache.GlobalSession.ExistSession(Session.SessionID, Key) Then
                    Dim ObjectTypeString As String
                    Dim ObjectContent As String
                    Dim ObjectType As System.Type

                    ObjectTypeString = PlatformRedisCache.GlobalSession.GetSession(Session.SessionID, Key & ".ObjectType")
                    ObjectContent = PlatformRedisCache.GlobalSession.GetSession(Session.SessionID, Key)

                    Select Case ObjectTypeString.ToUpper
                        Case "System.Data.DataSet".ToUpper
                            RetValue = PlatformRedisCache.DSDeserialize(ObjectContent)
                        Case "System.Data.DataTable".ToUpper
                            RetValue = PlatformRedisCache.DTDeserialize(ObjectContent)
                        Case Else
                            ObjectType = System.Type.GetType(ObjectTypeString)
                            RetValue = XMLDeserial(ObjectContent, ObjectType)
                    End Select
                End If

                Return RetValue
            End Get
            Set(value As Object)
                If IsNothing(value) = False Then
                    Dim ObjectTypeString As String
                    Dim ObjectContent As String

                    ObjectTypeString = value.GetType.ToString

                    Select Case ObjectTypeString.ToUpper
                        Case "System.Data.DataSet".ToUpper
                            ObjectContent = PlatformRedisCache.DSSerialize(value)
                        Case "System.Data.DataTable".ToUpper
                            ObjectContent = PlatformRedisCache.DTSerialize(value)
                        Case Else
                            ObjectContent = XMLSerial(value)
                    End Select

                    PlatformRedisCache.GlobalSession.WriteSession(Session.SessionID, Key & ".ObjectType", ObjectTypeString)
                    PlatformRedisCache.GlobalSession.WriteSession(Session.SessionID, Key, ObjectContent)
                End If
            End Set
        End Property

        Public Shared Sub KeepAlive()
            PlatformRedisCache.GlobalSession.KeepSession(Session.SessionID)
        End Sub

        Public Shared Sub Expired()
            PlatformRedisCache.GlobalSession.ExpireSession(Session.SessionID)
        End Sub

        Public Shared Function KeyExist(Key As String) As Boolean
            Return PlatformRedisCache.GlobalSession.ExistSession(Session.SessionID, Key)
        End Function
    End Class
End Module
