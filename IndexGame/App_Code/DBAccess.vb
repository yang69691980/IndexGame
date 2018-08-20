Imports System
Imports System.Data
Imports Microsoft.VisualBasic

Public Module DBAccess
    Private iDBType As EnumDBType = EnumDBType.SqlClient

    Public Enum EnumDBType
        OleDB = 0
        SqlClient = 1
    End Enum

    Public WriteOnly Property DBAccessSetType() As EnumDBType
        Set(ByVal value As EnumDBType)
            iDBType = value
        End Set
    End Property

    Public Function GetDB(ByVal DBConnStr As String, ByVal SS As String) As System.Data.DataTable
        Dim DBCmd As System.Data.Common.DbCommand

        DBCmd = GetDBObjCommand()
        DBCmd.CommandType = System.Data.CommandType.Text
        DBCmd.CommandText = SS

        Return GetDB(DBConnStr, DBCmd)
    End Function

    Public Function GetDB(ByVal DBConnStr As String, ByVal DBCmd As System.Data.Common.DbCommand) As System.Data.DataTable
        Dim DT As System.Data.DataTable
        Dim DA As System.Data.Common.DbDataAdapter
        Dim DC As System.Data.DataColumn
        Dim DBConn As System.Data.Common.DbConnection
        Dim SS As String

        DBConn = GetDBObjConnection()
        DBConn.ConnectionString = DBConnStr
        DBCmd.Connection = DBConn

        SS = DBCmd.CommandText

        DT = New System.Data.DataTable
        DA = GetDBObjDataAdapter()
        DA.SelectCommand = DBCmd

        Try
            DBConn.Open()
            'DA.FillSchema(DT, System.Data.SchemaType.Source)
            DA.Fill(DT)
        Catch ex As Exception
            Throw ex
        Finally
            Try
                DBConn.Close()
            Catch ex As Exception

            End Try
        End Try

        DT.ExtendedProperties.Add("CreateInstance", "DBAccess")
        DT.ExtendedProperties.Add("DBAccess_OleDBString", SS)
        DT.ExtendedProperties.Add("DBAccess_DBCommand", DBCmd)
        DT.ExtendedProperties.Add("DBAccess_AutoNumber", String.Empty)

        For Each DC In DT.Columns
            If DC.AutoIncrement Then
                DT.ExtendedProperties.Item("DBAccess_AutoNumber") = DC.ColumnName
                DC.AutoIncrementSeed = -1
                DC.AutoIncrementStep = -1
                Exit For
            End If
        Next

        Return DT
    End Function

    Public Function GetDBValue(ByVal DBConnStr As String, ByVal SS As String) As Object
        Dim DBCmd As System.Data.Common.DbCommand

        DBCmd = GetDBObjCommand()
        DBCmd.CommandType = System.Data.CommandType.Text
        DBCmd.CommandText = SS

        Return GetDBValue(DBConnStr, DBCmd)
    End Function

    Public Function GetDBValue(ByVal DBConnStr As String, ByVal DBCmd As System.Data.Common.DbCommand) As Object
        Dim DBConn As System.Data.Common.DbConnection
        Dim RetValue As Object
        Dim T As TransactionDB
        Dim DoTrans As Boolean = False
        Dim LDS As System.LocalDataStoreSlot

        LDS = System.Threading.Thread.GetNamedDataSlot("DBAccess_Transaction")
        If IsNothing(LDS) = False Then
            T = System.Threading.Thread.GetData(LDS)
            If IsNothing(T) = False Then
                If T.ConnectionString = DBConnStr Then
                    DoTrans = True
                End If
            End If
        End If

        If DoTrans Then
            RetValue = T.GetDBValue(DBCmd)
        Else
            DBConn = GetDBObjConnection()
            DBConn.ConnectionString = DBConnStr
            DBCmd.Connection = DBConn

            Try
                DBConn.Open()
                RetValue = DBCmd.ExecuteScalar
            Catch ex As Exception
                Throw ex
            Finally
                Try
                    DBConn.Close()
                Catch ex As Exception

                End Try
            End Try

            DBConn.Dispose()
            DBConn = Nothing
        End If

        Return RetValue
    End Function

    Public Function ExecuteDB(ByVal DBConnStr As String, ByVal SS As String) As Integer
        Dim DBCmd As System.Data.Common.DbCommand
        Dim RetValue As Integer

        DBCmd = GetDBObjCommand()

        DBCmd.CommandText = SS
        DBCmd.CommandType = System.Data.CommandType.Text

        RetValue = ExecuteDB(DBConnStr, DBCmd)

        Return RetValue
    End Function

    Public Function ExecuteDB(ByVal DBConnStr As String, ByVal Cmd As System.Data.Common.DbCommand) As Integer
        Dim DBConn As System.Data.Common.DbConnection
        Dim RetValue As Integer
        Dim T As TransactionDB
        Dim DoTrans As Boolean = False
        Dim LDS As System.LocalDataStoreSlot

        LDS = System.Threading.Thread.GetNamedDataSlot("DBAccess_Transaction")
        If IsNothing(LDS) = False Then
            T = System.Threading.Thread.GetData(LDS)
            If IsNothing(T) = False Then
                If T.ConnectionString = DBConnStr Then
                    DoTrans = True
                End If
            End If
        End If

        If DoTrans Then
            RetValue = T.ExecuteDB(Cmd)
        Else
            DBConn = GetDBObjConnection()
            DBConn.ConnectionString = DBConnStr
            Cmd.Connection = DBConn

            Try
                DBConn.Open()
                RetValue = Cmd.ExecuteNonQuery()
            Catch ex As Exception
                Throw ex
            Finally
                Try
                    DBConn.Close()
                Catch ex As Exception

                End Try
            End Try

            DBConn.Dispose()
            DBConn = Nothing
        End If

        Return RetValue
    End Function

    Public Sub ExecuteTransDB(DBConnStr As String, Func As Action(Of TransactionDB))
        Dim TransDB As TransactionDB = Nothing
        Dim ExecuteSuccess As Boolean = False
        Dim LDS As System.LocalDataStoreSlot

        LDS = System.Threading.Thread.GetNamedDataSlot("DBAccess_Transaction")
        If IsNothing(LDS) Then
            LDS = System.Threading.Thread.AllocateNamedDataSlot("DBAccess_Transaction")
        End If

        TransDB = New TransactionDB(DBConnStr)
        System.Threading.Thread.SetData(LDS, TransDB)
        Try
            Func.Invoke(TransDB)
            ExecuteSuccess = True
        Catch ex As Exception
            If IsNothing(TransDB) = False Then
                Try
                    TransDB.Rollback()
                Catch ex2 As Exception

                End Try
            End If

            System.Threading.Thread.FreeNamedDataSlot("DBAccess_Transaction")

            Throw ex
        End Try

        If ExecuteSuccess Then
            If IsNothing(TransDB) = False Then
                Try
                    TransDB.Commit()
                Catch ex As Exception

                End Try
            End If

            System.Threading.Thread.FreeNamedDataSlot("DBAccess_Transaction")
        End If
    End Sub

    Public Function SubmitDB(ByVal DBConnStr As String, ByVal DT As System.Data.DataTable) As System.Data.DataTable
        Dim DBConn As System.Data.Common.DbConnection
        Dim DA As System.Data.Common.DbDataAdapter
        Dim DBCmd As System.Data.Common.DbCommand
        Dim DBCmdBuilder As System.Data.Common.DbCommandBuilder
        Dim DV As System.Data.DataView
        Dim NeedUpdate As Boolean
        Dim QueryString As String

        NeedUpdate = False
        If DT.ExtendedProperties.Item("CreateInstance") = "DBAccess" Then
            QueryString = DT.ExtendedProperties("DBAccess_OleDBString")
            DBCmd = DT.ExtendedProperties("DBAccess_DBCommand")

            DBConn = GetDBObjConnection()
            DBConn.ConnectionString = DBConnStr

            'DBCmd = GetDBObjCommand()
            'DBCmd.CommandText = QueryString
            DBCmd.Connection = DBConn

            DA = GetDBObjDataAdapter()
            DA.SelectCommand = DBCmd

            'DA = DT.ExtendedProperties.Item("DBAccess_DataAdapter")

            DBCmdBuilder = GetDBObjCommandBuilder(DA)

            DV = New System.Data.DataView(DT)

            ' 檢查是否有新增資料
            DV.RowStateFilter = System.Data.DataViewRowState.Added
            If DV.Count > 0 Then
                DA.InsertCommand = DBCmdBuilder.GetInsertCommand
                NeedUpdate = True
            End If

            ' 檢查是否有刪除資料
            DV.RowStateFilter = System.Data.DataViewRowState.Deleted
            If DV.Count > 0 Then
                DA.DeleteCommand = DBCmdBuilder.GetDeleteCommand
                NeedUpdate = True
            End If

            ' 檢查是否有更新資料
            DV.RowStateFilter = System.Data.DataViewRowState.ModifiedCurrent
            If DV.Count > 0 Then
                DA.UpdateCommand = DBCmdBuilder.GetUpdateCommand
                NeedUpdate = True
            End If

            If NeedUpdate Then
                DBConn.Open()
                DBHandleUpdate(DA, DT)
                DBConn.Close()

                DT.AcceptChanges()
            End If

            DV.Dispose()
            DBCmdBuilder.Dispose()
            DA.Dispose()
            DBCmd.Dispose()
            DBConn.Dispose()

            DA = Nothing
            DV = Nothing
            DBCmdBuilder = Nothing
            DBCmd = Nothing
            DBConn = Nothing
        Else
            Err.Raise(1001, , "DataTable 無法識別")
        End If

        Return DT
    End Function

    Private Sub DBHandleUpdate(ByVal DA As System.Data.Common.DbDataAdapter, ByVal DT As System.Data.DataTable)
        Dim SqlDA As System.Data.SqlClient.SqlDataAdapter
        Dim OleDA As System.Data.OleDb.OleDbDataAdapter

        Select Case iDBType
            Case EnumDBType.OleDB
                OleDA = DA

                AddHandler OleDA.RowUpdated, AddressOf DBAccess_onRowUpdate_OleDB
                OleDA.Update(DT)
                RemoveHandler OleDA.RowUpdated, AddressOf DBAccess_onRowUpdate_OleDB
            Case EnumDBType.SqlClient
                SqlDA = DA

                AddHandler SqlDA.RowUpdated, AddressOf DBAccess_onRowUpdate_SqlClient
                SqlDA.Update(DT)
                RemoveHandler SqlDA.RowUpdated, AddressOf DBAccess_onRowUpdate_SqlClient
            Case Else
                DA.Update(DT)
        End Select
    End Sub

    Private Sub DBAccess_onRowUpdate_OleDB(ByVal sender As Object, ByVal args As Data.OleDb.OleDbRowUpdatedEventArgs)
        Dim newID As Object
        Dim IDFieldName As String
        Dim DBCmd As System.Data.OleDb.OleDbCommand
        Dim ColumnReadOnlyValue As Boolean

        If args.StatementType = System.Data.StatementType.Insert Then
            IDFieldName = args.Row.Table.ExtendedProperties.Item("DBAccess_AutoNumber")
            If IDFieldName <> String.Empty Then
                DBCmd = New System.Data.OleDb.OleDbCommand("SELECT @@IDENTITY")
                DBCmd.Connection = args.Command.Connection
                newID = DBCmd.ExecuteScalar()

                If Not (newID Is System.DBNull.Value) Then
                    ColumnReadOnlyValue = args.Row.Table.Columns(IDFieldName).ReadOnly

                    args.Row.Table.Columns(IDFieldName).ReadOnly = False
                    args.Row(IDFieldName) = CInt(newID)
                    args.Row.Table.Columns(IDFieldName).ReadOnly = ColumnReadOnlyValue
                End If
            End If
        End If
    End Sub

    Private Sub DBAccess_onRowUpdate_SqlClient(ByVal sender As Object, ByVal args As Data.SqlClient.SqlRowUpdatedEventArgs)
        Dim newID As Object
        Dim IDFieldName As String
        Dim DBCmd As System.Data.SqlClient.SqlCommand
        Dim ColumnReadOnlyValue As Boolean

        If args.StatementType = System.Data.StatementType.Insert Then
            IDFieldName = args.Row.Table.ExtendedProperties.Item("DBAccess_AutoNumber")
            If IDFieldName <> String.Empty Then
                DBCmd = New System.Data.SqlClient.SqlCommand("SELECT @@IDENTITY")
                DBCmd.Connection = args.Command.Connection
                newID = DBCmd.ExecuteScalar()

                If Not (newID Is System.DBNull.Value) Then
                    ColumnReadOnlyValue = args.Row.Table.Columns(IDFieldName).ReadOnly

                    args.Row.Table.Columns(IDFieldName).ReadOnly = False
                    args.Row(IDFieldName) = CInt(newID)
                    args.Row.Table.Columns(IDFieldName).ReadOnly = ColumnReadOnlyValue
                End If
            End If
        End If
    End Sub

    Private Function GetDBObjConnection() As System.Data.Common.DbConnection
        Dim RetValue As System.Data.Common.DbConnection = Nothing

        Select Case iDBType
            Case EnumDBType.OleDB
                RetValue = New System.Data.OleDb.OleDbConnection
            Case EnumDBType.SqlClient
                RetValue = New System.Data.SqlClient.SqlConnection
        End Select

        Return RetValue
    End Function

    Private Function GetDBObjCommand() As System.Data.Common.DbCommand
        Dim RetValue As System.Data.Common.DbCommand = Nothing

        Select Case iDBType
            Case EnumDBType.OleDB
                RetValue = New System.Data.OleDb.OleDbCommand
            Case EnumDBType.SqlClient
                RetValue = New System.Data.SqlClient.SqlCommand
        End Select

        Return RetValue
    End Function

    Private Function GetDBObjCommandBuilder(ByVal DA As System.Data.Common.DataAdapter) As System.Data.Common.DbCommandBuilder
        Dim RetValue As System.Data.Common.DbCommandBuilder = Nothing

        Select Case iDBType
            Case EnumDBType.OleDB
                RetValue = New System.Data.OleDb.OleDbCommandBuilder(DA)
            Case EnumDBType.SqlClient
                RetValue = New System.Data.SqlClient.SqlCommandBuilder(DA)
        End Select

        Return RetValue
    End Function

    Private Function GetDBObjDataAdapter() As System.Data.Common.DbDataAdapter
        Dim RetValue As System.Data.Common.DbDataAdapter = Nothing

        Select Case iDBType
            Case EnumDBType.OleDB
                RetValue = New System.Data.OleDb.OleDbDataAdapter()
            Case EnumDBType.SqlClient
                RetValue = New System.Data.SqlClient.SqlDataAdapter()
        End Select

        Return RetValue
    End Function

    Public Class TransactionDB
        Implements IDisposable

        Private iConnectionString As String = String.Empty
        Private iDBConn As System.Data.Common.DbConnection = Nothing
        Private iDBTrans As System.Data.Common.DbTransaction = Nothing

        Public ReadOnly Property ConnectionString As String
            Get
                Return iConnectionString
            End Get
        End Property

        Public Sub New(DBConnStr As String)
            Dim Success As Boolean = False

            iConnectionString = DBConnStr
            iDBConn = GetDBObjConnection()
            iDBConn.ConnectionString = DBConnStr

            Try
                iDBConn.Open()
                iDBTrans = iDBConn.BeginTransaction
            Catch ex As Exception
                If IsNothing(iDBTrans) = False Then
                    Try
                        iDBTrans.Dispose()
                    Catch ex2 As Exception

                    End Try
                End If

                If IsNothing(iDBConn) = False Then
                    Try
                        iDBConn.Close()
                    Catch ex2 As Exception

                    End Try

                    Try
                        iDBConn.Dispose()
                    Catch ex2 As Exception

                    End Try
                End If

                iDBConn = Nothing
                iDBTrans = Nothing

                Throw ex
            End Try
        End Sub

        Public Function ExecuteDB(SS As String) As Integer
            Dim DBCmd As System.Data.Common.DbCommand

            DBCmd = GetDBObjCommand()

            DBCmd.CommandText = SS
            DBCmd.CommandType = System.Data.CommandType.Text

            Return ExecuteDB(DBCmd)
        End Function

        Public Function ExecuteDB(Cmd As System.Data.Common.DbCommand) As Integer
            Dim RetValue As Integer

            Cmd.Connection = iDBConn
            Cmd.Transaction = iDBTrans

            Try
                RetValue = Cmd.ExecuteNonQuery()
            Catch ex As Exception
                Throw ex
            End Try

            Return RetValue
        End Function

        Public Function GetDBValue(ByVal SS As String) As Object
            Dim DBCmd As System.Data.Common.DbCommand

            DBCmd = GetDBObjCommand()
            DBCmd.CommandType = System.Data.CommandType.Text
            DBCmd.CommandText = SS

            Return GetDBValue(DBCmd)
        End Function

        Public Function GetDBValue(ByVal DBCmd As System.Data.Common.DbCommand) As Object
            Dim RetValue As Object

            DBCmd.Connection = iDBConn
            DBCmd.Transaction = iDBTrans

            Try
                RetValue = DBCmd.ExecuteScalar
            Catch ex As Exception
                Throw ex
            End Try

            Return RetValue
        End Function

        Public Sub Commit()
            If IsNothing(iDBTrans) = False Then
                iDBTrans.Commit()
            End If

            Me.Dispose()
        End Sub

        Public Sub Rollback()
            If IsNothing(iDBTrans) = False Then
                iDBTrans.Rollback()
            End If

            Me.Dispose()
        End Sub

#Region "IDisposable Support"
        Private disposedValue As Boolean ' 偵測多餘的呼叫

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: 處置 Managed 狀態 (Managed 物件)。
                    If IsNothing(iDBTrans) = False Then
                        Try
                            iDBTrans.Dispose()
                        Catch ex As Exception

                        End Try
                    End If

                    If IsNothing(iDBConn) = False Then
                        Try
                            iDBConn.Close()
                        Catch ex As Exception

                        End Try

                        Try
                            iDBConn.Dispose()
                        Catch ex As Exception

                        End Try
                    End If
                End If

                ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下方的 Finalize()。
                ' TODO: 將大型欄位設為 null。
            End If
            disposedValue = True
        End Sub

        ' TODO: 只有當上方的 Dispose(disposing As Boolean) 具有要釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
        'Protected Overrides Sub Finalize()
        '    ' 請勿變更這個程式碼。請將清除程式碼放在上方的 Dispose(disposing As Boolean) 中。
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' Visual Basic 加入這個程式碼的目的，在於能正確地實作可處置的模式。
        Public Sub Dispose() Implements IDisposable.Dispose
            ' 請勿變更這個程式碼。請將清除程式碼放在上方的 Dispose(disposing As Boolean) 中。
            Dispose(True)
            ' TODO: 覆寫上列 Finalize() 時，取消下行的註解狀態。
            ' GC.SuppressFinalize(Me)
        End Sub
#End Region
    End Class
End Module
