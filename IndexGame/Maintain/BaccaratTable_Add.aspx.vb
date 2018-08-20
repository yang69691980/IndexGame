Partial Class Admin_Add
    Inherits System.Web.UI.Page

    Protected Sub CheckLoginAccountExist(sender As Object, e As Ext.Net.RemoteValidationEventArgs)
        Dim Text As Ext.Net.TextField = sender
        Dim SS As String
        Dim DBCmd As System.Data.SqlClient.SqlCommand
        Dim Admin As AdminLoginState

        Admin = GetAdminLogin()

        SS = "SELECT COUNT(*) FROM AdminTable WITH (NOLOCK) WHERE LoginAccount=@LoginAccount" &
            " And forCompanyID=@forCompanyID"
        DBCmd = New System.Data.SqlClient.SqlCommand
        DBCmd.CommandText = SS
        DBCmd.CommandType = Data.CommandType.Text
        DBCmd.Parameters.Add("@LoginAccount", Data.SqlDbType.VarChar).Value = Text.Value
        DBCmd.Parameters.Add("@forCompanyID", Data.SqlDbType.Int).Value = selectCompany.SelectedItem.Value
        If GetDBValue(DBConnStr, DBCmd) <= 0 Then
            e.Success = True
        Else
            e.Success = False
            e.ErrorMessage = "登入帳號已經存在"
        End If
    End Sub

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim SS As String
        Dim DBCmd As System.Data.SqlClient.SqlCommand
        Dim AdminRoleDT As System.Data.DataTable
        Dim CompanyTableDT As System.Data.DataTable
        Dim Admin As AdminLoginState
        Dim mySelectItem As Ext.Net.ListItem

        If IsNothing(GetAdminLogin) Then Response.Redirect("RefreshParent.aspx?Login.aspx", True)

        'If CheckSessionStatePermission("Admin") <> enumSessionStatePermission.AccessSuccess Then AlertMessage("您沒有存取這個項目的權限", "javascript:window.parent.closeActiveTab();")
        Admin = GetAdminLogin()

        If Ext.Net.X.IsAjaxRequest = False Then
            SS = "Select * From AdminRole With(Nolock)" &
            " Where forCompanyID = @forCompanyID And AdminRoleState='0'"
            DBCmd = New System.Data.SqlClient.SqlCommand
            DBCmd.CommandText = SS
            DBCmd.CommandType = Data.CommandType.Text
            DBCmd.Parameters.Add("@forCompanyID", Data.SqlDbType.Int).Value = Admin.forCompanyID
            AdminRoleDT = GetDB(DBConnStr, DBCmd)

            For Each AdminRoleDRV As System.Data.DataRowView In AdminRoleDT.DefaultView
                mySelectItem = New Ext.Net.ListItem
                mySelectItem.Text = AdminRoleDRV("AdminRoleName")
                mySelectItem.Value = AdminRoleDRV("AdminRoleID")
                selectAdminRole.Items.Add(mySelectItem)
            Next

            If ChkPermission("OtherCompanyMaintain") = True Then
                selectCompany.Visible = True

                SS = "SELECT * FROM CompanyTable WITH (NOLOCK) Where CompanyState = 0"
                CompanyTableDT = GetDB(DBConnStr, SS)


                For Each EachDRV As System.Data.DataRowView In CompanyTableDT.DefaultView
                    mySelectItem = New Ext.Net.ListItem
                    mySelectItem.Text = EachDRV("CompanyName")
                    mySelectItem.Value = EachDRV("CompanyID")
                    selectCompany.Items.Add(mySelectItem)
                Next

                If CompanyTableDT.DefaultView.Count > 0 Then
                    selectCompany.SetValueAndFireSelect(CompanyTableDT.DefaultView(0)("CompanyID"))
                End If

            End If
        End If
    End Sub

    Protected Sub btnSave_DirectClick(sender As Object, e As Ext.Net.DirectEventArgs) Handles btnSave.DirectClick
        Dim SS As String
        Dim DBCmd As System.Data.SqlClient.SqlCommand
        Dim Admin As AdminLoginState
        Dim AllowUpdate As Boolean = False
        Dim forCompanyID As Integer = 0

        Admin = GetAdminLogin()

        If String.IsNullOrEmpty(txtLoginPassword.Text) = False Then
            If txtLoginPassword.Text <> txtLoginPassword2.Text Then
                ShowMsg("Exception", "密碼驗證失敗")
            Else
                AllowUpdate = True
            End If
        Else
            AllowUpdate = True
        End If

        If ChkPermission("OtherCompanyMaintain") = True Then
            forCompanyID = selectCompany.SelectedItem.Value
        Else
            forCompanyID = Admin.forCompanyID
        End If

        If AllowUpdate Then
            SS = "SELECT COUNT(*) FROM AdminTable WITH (NOLOCK) WHERE LoginAccount=@LoginAccount And forCompanyID=@forCompanyID"
            DBCmd = New System.Data.SqlClient.SqlCommand
            DBCmd.CommandText = SS
            DBCmd.CommandType = Data.CommandType.Text
            DBCmd.Parameters.Add("@LoginAccount", Data.SqlDbType.VarChar).Value = txtLoginAccount.Text
            DBCmd.Parameters.Add("@forCompanyID", Data.SqlDbType.Int).Value = selectCompany.SelectedItem.Value
            If GetDBValue(DBConnStr, DBCmd) <= 0 Then
                SS = "INSERT INTO AdminTable (LoginAccount, LoginPassword, RealName, Description, forAdminRoleID, forCompanyID) " &
                     "                        VALUES (@LoginAccount, @LoginPassword, @RealName, @Description, @forAdminRoleID, @forCompanyID)"
                DBCmd = New System.Data.SqlClient.SqlCommand
                DBCmd.CommandText = SS
                DBCmd.CommandType = Data.CommandType.Text
                DBCmd.Parameters.Add("@LoginAccount", Data.SqlDbType.VarChar).Value = txtLoginAccount.Text
                DBCmd.Parameters.Add("@LoginPassword", Data.SqlDbType.VarChar).Value = GetMD5(txtLoginPassword.Text)
                DBCmd.Parameters.Add("@RealName", Data.SqlDbType.NVarChar).Value = txtRealname.Text
                DBCmd.Parameters.Add("@Description", Data.SqlDbType.NVarChar).Value = txtDescription.Text
                DBCmd.Parameters.Add("@forAdminRoleID", Data.SqlDbType.Int).Value = selectAdminRole.SelectedItem.Value
                DBCmd.Parameters.Add("@forCompanyID", Data.SqlDbType.Int).Value = selectCompany.SelectedItem.Value
                ExecuteDB(DBConnStr, DBCmd)

                CloseActiveTab("Admin_Maint.aspx", "Message", "儲存成功")
            Else
                ShowMsg("Exception", "帳號已經存在")
            End If
        End If
    End Sub

    Protected Sub btnClose_DirectClick(sender As Object, e As Ext.Net.DirectEventArgs) Handles btnClose.DirectClick
        CloseActiveTab()
    End Sub

    Private Sub selectCompany_DirectSelect(sender As Object, e As DirectEventArgs) Handles selectCompany.DirectSelect
        Dim SS As String
        Dim DBCmd As System.Data.SqlClient.SqlCommand
        Dim AdminRoleDT As System.Data.DataTable


        SS = "Select * From AdminRole With(Nolock)" &
        " Where forCompanyID = @forCompanyID And AdminRoleState='0'"
        DBCmd = New System.Data.SqlClient.SqlCommand
        DBCmd.CommandText = SS
        DBCmd.CommandType = Data.CommandType.Text
        DBCmd.Parameters.Add("@forCompanyID", Data.SqlDbType.Int).Value = selectCompany.SelectedItem.Value
        AdminRoleDT = GetDB(DBConnStr, DBCmd)

        selectAdminRole.GetStore.RemoveAll(True)
        selectAdminRole.ClearValue()

        For Each AdminRoleDRV As System.Data.DataRowView In AdminRoleDT.DefaultView
            selectAdminRole.AddItem(AdminRoleDRV("AdminRoleName"), AdminRoleDRV("AdminRoleID"))
        Next

    End Sub

End Class
