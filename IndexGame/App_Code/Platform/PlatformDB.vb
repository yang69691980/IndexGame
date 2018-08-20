Imports Microsoft.VisualBasic

Public Class PlatformDB
    Public Shared Function DisableUserAccountPoint(UserAccountID As Integer, PointType As String) As Integer
        Dim SS As String
        Dim DBCmd As System.Data.SqlClient.SqlCommand

        SS = "spDisableUserAccountPoint"
        DBCmd = New System.Data.SqlClient.SqlCommand
        DBCmd.CommandText = SS
        DBCmd.CommandType = Data.CommandType.StoredProcedure
        DBCmd.Parameters.Add("@UserAccountID", Data.SqlDbType.Int).Value = UserAccountID
        DBCmd.Parameters.Add("@PointType", Data.SqlDbType.VarChar).Value = PointType
        DBCmd.Parameters.Add("@Return", Data.SqlDbType.Int).Direction = Data.ParameterDirection.ReturnValue
        ExecuteDB(DBConnStr, DBCmd)

        Return DBCmd.Parameters("@Return").Value
    End Function

    Public Shared Function CreateUserAccountPoint(UserAccountID As Integer, PointType As String) As Integer
        Dim SS As String
        Dim DBCmd As System.Data.SqlClient.SqlCommand

        SS = "spCreateUserAccountPoint"
        DBCmd = New System.Data.SqlClient.SqlCommand
        DBCmd.CommandText = SS
        DBCmd.CommandType = Data.CommandType.StoredProcedure
        DBCmd.Parameters.Add("@UserAccountID", Data.SqlDbType.Int).Value = UserAccountID
        DBCmd.Parameters.Add("@PointType", Data.SqlDbType.VarChar).Value = PointType
        DBCmd.Parameters.Add("@Return", Data.SqlDbType.Int).Direction = Data.ParameterDirection.ReturnValue
        ExecuteDB(DBConnStr, DBCmd)

        Return DBCmd.Parameters("@Return").Value
    End Function

    Public Shared Function UserAccountDeposit(ApiKey As String, UserAccountID As Integer, TransactionCode As String, RemoteIP As String, PointType As String, Amount As Decimal) As Integer
        Dim SS As String
        Dim DBCmd As System.Data.SqlClient.SqlCommand

        SS = "spUserAccountDeposit"
        DBCmd = New System.Data.SqlClient.SqlCommand
        DBCmd.CommandText = SS
        DBCmd.CommandType = Data.CommandType.StoredProcedure
        DBCmd.Parameters.Add("@ApiKey", Data.SqlDbType.VarChar).Value = ApiKey
        DBCmd.Parameters.Add("@UserAccountID", Data.SqlDbType.Int).Value = UserAccountID
        DBCmd.Parameters.Add("@TransactionCode", Data.SqlDbType.VarChar).Value = TransactionCode
        DBCmd.Parameters.Add("@RemoteIP", Data.SqlDbType.VarChar).Value = RemoteIP
        DBCmd.Parameters.Add("@PointType", Data.SqlDbType.VarChar).Value = PointType
        DBCmd.Parameters.Add("@Amount", Data.SqlDbType.Decimal).Value = Amount
        DBCmd.Parameters.Add("@Return", Data.SqlDbType.Int).Direction = Data.ParameterDirection.ReturnValue
        ExecuteDB(DBConnStr, DBCmd)

        Return DBCmd.Parameters("@Return").Value
    End Function

    Public Shared Function UserAccountDepositConfirm(ApiKey As String, UserAccountID As Integer, TransactionCode As String, RemoteIP As String, PointType As String, Amount As Decimal, ByRef FreePoint As Decimal) As Integer
        Dim SS As String
        Dim DBCmd As System.Data.SqlClient.SqlCommand
        Dim P As System.Data.SqlClient.SqlParameter

        SS = "spUserAccountDepositConfirm"
        DBCmd = New System.Data.SqlClient.SqlCommand
        DBCmd.CommandText = SS
        DBCmd.CommandType = Data.CommandType.StoredProcedure
        DBCmd.Parameters.Add("@ApiKey", Data.SqlDbType.VarChar).Value = ApiKey
        DBCmd.Parameters.Add("@UserAccountID", Data.SqlDbType.Int).Value = UserAccountID
        DBCmd.Parameters.Add("@TransactionCode", Data.SqlDbType.VarChar).Value = TransactionCode
        DBCmd.Parameters.Add("@RemoteIP", Data.SqlDbType.VarChar).Value = RemoteIP
        DBCmd.Parameters.Add("@PointType", Data.SqlDbType.VarChar).Value = PointType
        DBCmd.Parameters.Add("@Amount", Data.SqlDbType.Decimal).Value = Amount

        P = New System.Data.SqlClient.SqlParameter("@FreePoint", Data.SqlDbType.Decimal)
        P.Direction = Data.ParameterDirection.Output
        P.Precision = 18
        P.Scale = 4

        DBCmd.Parameters.Add(P)
        DBCmd.Parameters.Add("@Return", Data.SqlDbType.Int).Direction = Data.ParameterDirection.ReturnValue
        ExecuteDB(DBConnStr, DBCmd)

        If DBCmd.Parameters("@Return").Value = 0 Then
            FreePoint = DBCmd.Parameters("@FreePoint").Value
        End If

        Return DBCmd.Parameters("@Return").Value
    End Function


    Public Shared Function UserAccountDraw(ApiKey As String, UserAccountID As Integer, TransactionCode As String, RemoteIP As String, PointType As String, Amount As Decimal) As Integer
        Dim SS As String
        Dim DBCmd As System.Data.SqlClient.SqlCommand

        SS = "spUserAccountDraw"
        DBCmd = New System.Data.SqlClient.SqlCommand
        DBCmd.CommandText = SS
        DBCmd.CommandType = Data.CommandType.StoredProcedure
        DBCmd.Parameters.Add("@ApiKey", Data.SqlDbType.VarChar).Value = ApiKey
        DBCmd.Parameters.Add("@UserAccountID", Data.SqlDbType.Int).Value = UserAccountID
        DBCmd.Parameters.Add("@TransactionCode", Data.SqlDbType.VarChar).Value = TransactionCode
        DBCmd.Parameters.Add("@RemoteIP", Data.SqlDbType.VarChar).Value = RemoteIP
        DBCmd.Parameters.Add("@PointType", Data.SqlDbType.VarChar).Value = PointType
        DBCmd.Parameters.Add("@Amount", Data.SqlDbType.Decimal).Value = Amount
        DBCmd.Parameters.Add("@Return", Data.SqlDbType.Int).Direction = Data.ParameterDirection.ReturnValue
        ExecuteDB(DBConnStr, DBCmd)

        Return DBCmd.Parameters("@Return").Value
    End Function

    Public Shared Function UserAccountDrawConfirm(ApiKey As String, UserAccountID As Integer, TransactionCode As String, RemoteIP As String, PointType As String, Amount As Decimal, ByRef FreePoint As Decimal) As Integer
        Dim SS As String
        Dim DBCmd As System.Data.SqlClient.SqlCommand
        Dim P As System.Data.SqlClient.SqlParameter

        SS = "spUserAccountDrawConfirm"
        DBCmd = New System.Data.SqlClient.SqlCommand
        DBCmd.CommandText = SS
        DBCmd.CommandType = Data.CommandType.StoredProcedure
        DBCmd.Parameters.Add("@ApiKey", Data.SqlDbType.VarChar).Value = ApiKey
        DBCmd.Parameters.Add("@UserAccountID", Data.SqlDbType.Int).Value = UserAccountID
        DBCmd.Parameters.Add("@TransactionCode", Data.SqlDbType.VarChar).Value = TransactionCode
        DBCmd.Parameters.Add("@RemoteIP", Data.SqlDbType.VarChar).Value = RemoteIP
        DBCmd.Parameters.Add("@PointType", Data.SqlDbType.VarChar).Value = PointType
        DBCmd.Parameters.Add("@Amount", Data.SqlDbType.Decimal).Value = Amount

        P = New System.Data.SqlClient.SqlParameter("@FreePoint", Data.SqlDbType.Decimal)
        P.Direction = Data.ParameterDirection.Output
        P.Precision = 18
        P.Scale = 4

        DBCmd.Parameters.Add(P)
        DBCmd.Parameters.Add("@Return", Data.SqlDbType.Int).Direction = Data.ParameterDirection.ReturnValue
        ExecuteDB(DBConnStr, DBCmd)

        If DBCmd.Parameters("@Return").Value = 0 Then
            FreePoint = DBCmd.Parameters("@FreePoint").Value
        End If

        Return DBCmd.Parameters("@Return").Value
    End Function


    Public Shared Function GetUserAccountByLoginAccount(LoginAccount As String, CompanyCode As String) As System.Data.DataTable
        Dim DT As System.Data.DataTable
        Dim SS As String
        Dim DBCmd As System.Data.SqlClient.SqlCommand

        SS = "SELECT * FROM UserAccount WITH (NOLOCK) WHERE LoginAccount=@LoginAccount AND CompanyCode=@CompanyCode"
        DBCmd = New System.Data.SqlClient.SqlCommand
        DBCmd.CommandText = SS
        DBCmd.CommandType = Data.CommandType.Text
        DBCmd.Parameters.Add("@LoginAccount", Data.SqlDbType.VarChar).Value = LoginAccount
        DBCmd.Parameters.Add("@CompanyCode", Data.SqlDbType.VarChar).Value = CompanyCode
        DT = GetDB(DBConnStr, DBCmd)

        Return DT
    End Function

    Public Shared Function GetUserAccountByID(UserAccountID As Integer) As System.Data.DataTable
        Dim DT As System.Data.DataTable
        Dim SS As String
        Dim DBCmd As System.Data.SqlClient.SqlCommand

        SS = "SELECT * FROM UserAccount WITH (NOLOCK) WHERE UserAccountID=@UserAccountID"
        DBCmd = New System.Data.SqlClient.SqlCommand
        DBCmd.CommandText = SS
        DBCmd.CommandType = Data.CommandType.Text
        DBCmd.Parameters.Add("@UserAccount", Data.SqlDbType.Int).Value = UserAccountID
        DT = GetDB(DBConnStr, DBCmd)

        Return DT
    End Function

End Class
