
Partial Class Admin_Maint
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        If IsNothing(GetAdminLogin) Then Response.Redirect("RefreshParent.aspx?Login.aspx", True)

        'If CheckSessionStatePermission("Admin") <> enumSessionStatePermission.AccessSuccess Then AlertMessage("您沒有存取這個項目的權限", "javascript:window.parent.closeActiveTab();")

        If Ext.Net.X.IsAjaxRequest = False Then
            Store1_SetDataBind()
        End If
    End Sub

    Public Sub Store1_RefreshData(sender As Object, e As Ext.Net.StoreReadDataEventArgs)
        Store1_SetDataBind()
    End Sub

    Public Sub Store1_SetDataBind()
        Dim DT As System.Data.DataTable
        Dim SS As String
        Dim DBCmd As System.Data.SqlClient.SqlCommand
        Dim Admin As AdminLoginState
        Dim Param As String = String.Empty

        Admin = GetAdminLogin()

        If ChkPermission("OtherCompanyMaintain") = False Then
            Param = " Where forCompanyID = @forCompanyID"
        End If

        SS = "SELECT * FROM AdminTable WITH (NOLOCK) " & Param
        DBCmd = New System.Data.SqlClient.SqlCommand
        DBCmd.CommandText = SS
        DBCmd.CommandType = Data.CommandType.Text
        If ChkPermission("OtherCompanyMaintain") = False Then
            DBCmd.Parameters.Add("@forCompanyID", Data.SqlDbType.Int).Value = Admin.forCompanyID
        End If
        DT = GetDB(DBConnStr, DBCmd)

        Store1.DataSource = DT
        Store1.DataBind()
    End Sub

    Public Sub Store1_Submit(sender As Object, e As Ext.Net.StoreSubmitDataEventArgs)
        Dim xml As System.Xml.XmlNode

        xml = e.Xml
        Response.Clear()

        Select Case FormatType.Value.ToString
            Case "xml"
                Dim strXml As String = xml.OuterXml

                Response.AddHeader("Content-Disposition", "attachment; filename=submittedData.xml")
                Response.AddHeader("Content-Length", strXml.Length)
                Response.ContentType = "application/xml"
                Response.Write(strXml)
            Case "xls"
                Dim xtExcel As System.Xml.Xsl.XslCompiledTransform

                Response.ContentType = "application/vnd.ms-excel"
                Response.AddHeader("Content-Disposition", "attachment; filename=submittedData.xls")

                xtExcel = New System.Xml.Xsl.XslCompiledTransform
                xtExcel.Load(Server.MapPath("../Files/Excel.xsl"))
                xtExcel.Transform(xml, Nothing, Response.OutputStream)
            Case "csv"
                Dim xtCsv As System.Xml.Xsl.XslCompiledTransform

                Response.ContentType = "application/octet-stream"
                Response.AddHeader("Content-Disposition", "attachment; filename=submittedData.csv")

                xtCsv = New System.Xml.Xsl.XslCompiledTransform
                xtCsv.Load(Server.MapPath("../Files/Csv.xsl"))
                xtCsv.Transform(xml, Nothing, Response.OutputStream)
        End Select

        Response.Flush()
        Response.End()
    End Sub

    Protected Sub btnNew_DirectClick(sender As Object, e As Ext.Net.DirectEventArgs) Handles btnNew.DirectClick
        NewTabToURL("後端帳號新增", "Admin_Add.aspx")
    End Sub

    Protected Sub btnEdit_DirectClick(sender As Object, e As Ext.Net.DirectEventArgs) Handles btnEdit.DirectClick
        Dim AdminID As Integer

        If RowSelectionModel1.SelectedRows.Count > 0 Then
            AdminID = RowSelectionModel1.SelectedRows(0).RecordID

            NewTabToURL("後端帳號編輯", "Admin_Edit.aspx?ID=" & AdminID)
        End If
    End Sub

    Protected Sub btnDelete_DirectClick(sender As Object, e As Ext.Net.DirectEventArgs) Handles btnDelete.DirectClick
        If RowSelectionModel1.SelectedRows.Count > 0 Then
            For Each EachRow As Ext.Net.SelectedRow In RowSelectionModel1.SelectedRows
                Dim SS As String
                Dim DBCmd As System.Data.SqlClient.SqlCommand
                Dim Admin As AdminLoginState

                Admin = GetAdminLogin()

                SS = "UPDATE AdminTable SET AdminState=1 WHERE AdminID=@AdminID"
                DBCmd = New Data.SqlClient.SqlCommand
                DBCmd.CommandText = SS
                DBCmd.CommandType = Data.CommandType.Text
                DBCmd.Parameters.Add("@AdminID", Data.SqlDbType.Int).Value = EachRow.RecordID
                ExecuteDB(DBConnStr, DBCmd)

                ReloadActiveTab()
            Next
        End If
    End Sub
End Class
