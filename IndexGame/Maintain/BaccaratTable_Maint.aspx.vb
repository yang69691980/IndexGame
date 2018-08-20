
Partial Class Admin_Maint
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim DT As System.Data.DataTable
        Dim SS As String
        Dim mySelectItem As New Ext.Net.ListItem

        If Ext.Net.X.IsAjaxRequest = False Then
            SS = "SELECT * FROM GameArea WITH (NOLOCK) Where GameAreaState = 0"
            DT = GetDB(DBConnStr, SS)

            For Each EachDRV As System.Data.DataRowView In DT.DefaultView
                mySelectItem = New Ext.Net.ListItem
                mySelectItem.Text = EachDRV("Description")
                mySelectItem.Value = EachDRV("GameAreaCode")
                selectGamArea.Items.Add(mySelectItem)
            Next

            If DT.DefaultView.Count > 0 And IsNothing(selectGamArea.SelectedItem.Value) Then
                selectGamArea.SetValueAndFireSelect(DT.DefaultView(0)("GameAreaCode"))
            Else
                selectGamArea.SetValueAndFireSelect(selectGamArea.SelectedItem.Value)
            End If

            'Store1_SetDataBind()
        End If
    End Sub

    Public Sub Store1_RefreshData(sender As Object, e As Ext.Net.StoreReadDataEventArgs)
        Store1_SetDataBind()
    End Sub

    Public Sub Store1_SetDataBind()
        Dim DT As System.Data.DataTable
        Dim SS As String
        Dim DBCmd As System.Data.SqlClient.SqlCommand
        Dim Param As String = String.Empty
        Dim GameTokenResult As Boolean = False

        GameTokenResult = CheckGameToken(Request("GameToken"))

        If GameTokenResult = True Then
            If IsNothing(selectGamArea.SelectedItem.Value) = False Then
                Param = " Where BaccaratTable.GameAreaCode=@GameAreaCode"
            End If

            SS = "Select BaccaratTable.*, GameArea.Description As GameAreaName From BaccaratTable With(Nolock)" &
            " Left Join GameArea On BaccaratTable.GameAreaCode = GameArea.GameAreaCode " & Param
            DBCmd = New System.Data.SqlClient.SqlCommand
            DBCmd.CommandType = Data.CommandType.Text
            DBCmd.CommandText = SS
            If IsNothing(selectGamArea.SelectedItem.Value) = False Then
                DBCmd.Parameters.Add("@GameAreaCode", Data.SqlDbType.VarChar).Value = selectGamArea.SelectedItem.Value
            End If
            DT = GetDB(DBConnStr, DBCmd)

            Store1.DataSource = DT
            Store1.DataBind()
        End If
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
        NewTabToURL("桌新增", "BaccaratTable_Add.aspx")
    End Sub

    Protected Sub btnEdit_DirectClick(sender As Object, e As Ext.Net.DirectEventArgs) Handles btnEdit.DirectClick
        Dim EditRowID As Integer

        If RowSelectionModel1.SelectedRows.Count > 0 Then
            EditRowID = RowSelectionModel1.SelectedRows(0).RecordID

            NewTabToURL("桌編輯", "BaccaratTable_Add.aspx?ID=" & EditRowID)
        End If
    End Sub

    Protected Sub btnDelete_DirectClick(sender As Object, e As Ext.Net.DirectEventArgs) Handles btnDelete.DirectClick
        If RowSelectionModel1.SelectedRows.Count > 0 Then
            For Each EachRow As Ext.Net.SelectedRow In RowSelectionModel1.SelectedRows
                Dim SS As String
                Dim DBCmd As System.Data.SqlClient.SqlCommand

                SS = "UPDATE BaccaratTable SET TableState=3 WHERE TableID=@TableID"
                DBCmd = New Data.SqlClient.SqlCommand
                DBCmd.CommandText = SS
                DBCmd.CommandType = Data.CommandType.Text
                DBCmd.Parameters.Add("@TableID", Data.SqlDbType.Int).Value = EachRow.RecordID
                ExecuteDB(DBConnStr, DBCmd)

                ReloadActiveTab()
            Next
        End If
    End Sub

    Private Sub selectGamArea_DirectSelect(sender As Object, e As Ext.Net.DirectEventArgs) Handles selectGamArea.DirectSelect
        Store1_SetDataBind()
    End Sub
End Class
