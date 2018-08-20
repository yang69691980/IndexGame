Partial Class Admin_Edit
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        If IsNothing(GetAdminLogin) Then Response.Redirect("RefreshParent.aspx?Login.aspx", True)

        'If CheckSessionStatePermission("Admin") <> enumSessionStatePermission.AccessSuccess Then AlertMessage("您沒有存取這個項目的權限", "javascript:window.parent.closeActiveTab();")

        If Ext.Net.X.IsAjaxRequest = False Then
            Dim DT As System.Data.DataTable
            Dim AdminID As Integer
            Dim SS As String
            Dim DBCmd As System.Data.SqlClient.SqlCommand
            Dim Admin As AdminLoginState
            Dim CompanyTableDT As System.Data.DataTable
            Dim mySelectItem As Ext.Net.ListItem
            Dim AdminRoleDT As System.Data.DataTable


            AdminID = Request("ID")
            Admin = GetAdminLogin()

            SS = "SELECT * FROM AdminTable WITH (NOLOCK) WHERE AdminID=@AdminID"
            DBCmd = New System.Data.SqlClient.SqlCommand
            DBCmd.CommandText = SS
            DBCmd.CommandType = Data.CommandType.Text
            DBCmd.Parameters.Add("@AdminID", Data.SqlDbType.Int).Value = Request("ID")
            DT = GetDB(DBConnStr, DBCmd)

            If DT.Rows.Count > 0 Then
                SS = "Select * From AdminRole With(Nolock)" &
                " Where forCompanyID = @forCompanyID And AdminRoleState='0'"
                DBCmd = New System.Data.SqlClient.SqlCommand
                DBCmd.CommandText = SS
                DBCmd.CommandType = Data.CommandType.Text
                DBCmd.Parameters.Add("@forCompanyID", Data.SqlDbType.Int).Value = DT.Rows(0)("forCompanyID")
                AdminRoleDT = GetDB(DBConnStr, DBCmd)

                lblLoginAccount.Text = DT.Rows(0)("LoginAccount")

                Select Case DT.Rows(0)("AdminState")
                    Case 0
                        rdoAdminStateNormal.Checked = True
                    Case 1
                        rdoAdminStateDisable.Checked = True
                End Select

                txtRealname.Text = DT.Rows(0)("RealName")
                txtDescription.Text = DT.Rows(0)("Description")

                For Each AdminRoleDRV As System.Data.DataRowView In AdminRoleDT.DefaultView
                    mySelectItem = New Ext.Net.ListItem
                    mySelectItem.Text = AdminRoleDRV("AdminRoleName")
                    mySelectItem.Value = AdminRoleDRV("AdminRoleID")
                    selectAdminRole.Items.Add(mySelectItem)
                Next

                If AdminRoleDT.Rows.Count > 0 Then
                    selectAdminRole.SetValueAndFireSelect(DT.Rows(0)("forAdminRoleID"))
                End If

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
                        selectCompany.SetValue(DT.Rows(0)("forCompanyID"))
                        'selectCompany.SetValueAndFireSelect(DT.Rows(0)("forCompanyID"))
                    End If
                End If


            Else
                CloseActiveTab(Nothing, "Exception", "帳號不存在")
            End If
        End If
    End Sub

    Protected Sub btnSave_DirectClick(sender As Object, e As Ext.Net.DirectEventArgs) Handles btnSave.DirectClick
        Dim SS As String
        Dim DBCmd As System.Data.SqlClient.SqlCommand
        Dim DT As System.Data.DataTable
        Dim AdminID As Integer
        Dim Admin As AdminLoginState
        Dim WhereString As String
        Dim AllowUpdate As Boolean = False
        Dim forCompanyID As Integer = 0

        AdminID = Request("ID")
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


        If AllowUpdate = True Then
            If String.IsNullOrEmpty(selectAdminRole.SelectedItem.Value) = True Then
                ShowMsg("Exception", "請選擇權限")
                AllowUpdate = False
            End If
        End If

        If ChkPermission("OtherCompanyMaintain") = True Then
            forCompanyID = selectCompany.SelectedItem.Value
        Else
            forCompanyID = Admin.forCompanyID
        End If

        If AllowUpdate = True Then
            'ShowMsg("", Admin.AdminAccount & "," & selectCompany.SelectedItem.Value & "," & Admin.AdminID)
            SS = "SELECT COUNT(*) FROM AdminTable WITH (NOLOCK)" &
            " WHERE LoginAccount=(Select LoginAccount From AdminTable WITH (Nolock) Where AdminID=@AdminID)" &
            " And forCompanyID=@forCompanyID" &
            " And AdminID <> @AdminID"
            DBCmd = New System.Data.SqlClient.SqlCommand
            DBCmd.CommandText = SS
            DBCmd.CommandType = Data.CommandType.Text
            DBCmd.Parameters.Add("@forCompanyID", Data.SqlDbType.Int).Value = selectCompany.SelectedItem.Value
            DBCmd.Parameters.Add("@AdminID", Data.SqlDbType.Int).Value = AdminID
            If GetDBValue(DBConnStr, DBCmd) > 0 Then
                AllowUpdate = False
                ShowMsg("Exception", "帳號已經存在")
            End If
        End If

        If AllowUpdate Then
            DBCmd = New System.Data.SqlClient.SqlCommand
            DBCmd.CommandType = Data.CommandType.Text

            WhereString = "RealName=@RealName, Description=@Description, AdminState=@AdminState, forAdminRoleID=@forAdminRoleID, forCompanyID=@forCompanyID"
            DBCmd.Parameters.Add("@RealName", Data.SqlDbType.NVarChar).Value = txtRealname.Text
            DBCmd.Parameters.Add("@Description", Data.SqlDbType.NVarChar).Value = txtDescription.Text
            DBCmd.Parameters.Add("@forAdminRoleID", Data.SqlDbType.Int).Value = selectAdminRole.SelectedItem.Value
            DBCmd.Parameters.Add("@forCompanyID", Data.SqlDbType.Int).Value = forCompanyID

            If rdoAdminStateNormal.Checked Then
                DBCmd.Parameters.Add("@AdminState", Data.SqlDbType.Int).Value = 0
            ElseIf rdoAdminStateDisable.Checked Then
                DBCmd.Parameters.Add("@AdminState", Data.SqlDbType.Int).Value = 1
            End If

            If String.IsNullOrEmpty(txtLoginPassword.Text) = False Then
                If String.IsNullOrEmpty(WhereString) = False Then
                    WhereString &= ","
                End If
                WhereString &= "LoginPassword=@LoginPassword"
                DBCmd.Parameters.Add("@LoginPassword", Data.SqlDbType.VarChar).Value = GetMD5(txtLoginPassword.Text)
            End If

            SS = "UPDATE AdminTable SET " & WhereString & " WHERE AdminID=@AdminID"
            DBCmd.CommandText = SS
            DBCmd.Parameters.Add("@AdminID", Data.SqlDbType.Int).Value = AdminID
            ExecuteDB(DBConnStr, DBCmd)

            CloseActiveTab("Admin_Maint.aspx", "Message", "儲存成功")
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
