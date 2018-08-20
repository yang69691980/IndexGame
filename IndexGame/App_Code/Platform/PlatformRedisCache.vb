Imports Microsoft.VisualBasic

Public Class PlatformRedisCache
    Public Class GlobalSession
        Private Shared XMLPath As String = "GlobalSession"

        Private Shared iSyncRoot As New ArrayList
        Private Shared iDBIndex As Integer = 0

        Public Shared Function ExistSession(SessionID As String, SessionKey As String) As Boolean
            Dim Client As StackExchange.Redis.IDatabase = Platform.GetRedisClient(iDBIndex)
            Dim Key As String

            Key = XMLPath & ":" & SessionID
            Return Client.HashExists(Key.ToUpper, SessionKey.ToUpper)
        End Function

        Public Shared Sub ExpireSession(SessionID As String)
            Dim Client As StackExchange.Redis.IDatabase = Platform.GetRedisClient(iDBIndex)
            Dim Key As String

            Key = XMLPath & ":" & SessionID
            If Client.KeyExists(Key.ToUpper) Then
                Client.KeyDelete(Key.ToUpper)
            End If
        End Sub

        Public Shared Function GetSession(SessionID As String, SessionKey As String) As String
            Dim Client As StackExchange.Redis.IDatabase = Platform.GetRedisClient(iDBIndex)
            Dim Key As String
            Dim RetValue As String

            Key = XMLPath & ":" & SessionID
            If Client.HashExists(Key.ToUpper, SessionKey.ToUpper) Then
                RetValue = Client.HashGet(Key.ToUpper, SessionKey.ToUpper)
            End If

            Return RetValue
        End Function

        Public Shared Sub WriteSession(SessionID As String, SessionKey As String, Value As String)
            Dim Client As StackExchange.Redis.IDatabase = Platform.GetRedisClient(iDBIndex)
            Dim Key As String

            Key = XMLPath & ":" & SessionID
            Client.HashSet(Key.ToUpper, SessionKey.ToUpper, Value)
            Client.KeyExpire(Key.ToUpper, New TimeSpan(0, 0, 300))
        End Sub

        Public Shared Sub KeepSession(SessionID As String)
            Dim Client As StackExchange.Redis.IDatabase = Platform.GetRedisClient(iDBIndex)
            Dim Key As String

            Key = XMLPath & ":" & SessionID
            If Client.KeyExists(Key.ToUpper) Then
                Client.KeyExpire(Key.ToUpper, New TimeSpan(0, 0, 300))
            End If
        End Sub
    End Class


    Public Shared Function DSSerialize(_ds As System.Data.DataSet) As String
        Dim result As String = String.Empty

        If IsNothing(_ds) = False Then
            Dim writer As New System.IO.StringWriter
            Dim I As Integer = 0

            For Each EachTable As System.Data.DataTable In _ds.Tables
                I += 1

                If String.IsNullOrEmpty(EachTable.TableName) Then
                    EachTable.TableName = "Datatable." & I
                End If
            Next

            _ds.WriteXml(writer, System.Data.XmlWriteMode.WriteSchema)

            result = writer.ToString()
        End If

        Return result
    End Function

    Public Shared Function DTSerialize(_dt As System.Data.DataTable) As String
        Dim result As String = String.Empty

        If IsNothing(_dt) = False Then
            Dim writer As New System.IO.StringWriter

            If String.IsNullOrEmpty(_dt.TableName) Then
                _dt.TableName = "Datatable"
            End If

            _dt.WriteXml(writer, System.Data.XmlWriteMode.WriteSchema)

            result = writer.ToString()
        End If

        Return result
    End Function

    Public Shared Function DTDeserialize(_strDATA As String) As System.Data.DataTable
        Dim DT As New System.Data.DataTable
        Dim StringStream As New System.IO.StringReader(_strDATA)

        DT.ReadXml(StringStream)

        Return DT
    End Function

    Public Shared Function DSDeserialize(_strDATA As String) As System.Data.DataSet
        Dim DS As New System.Data.DataSet
        Dim StringStream As New System.IO.StringReader(_strDATA)

        DS.ReadXml(StringStream)

        Return DS
    End Function

End Class
