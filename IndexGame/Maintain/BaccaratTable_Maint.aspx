<%@ Page Language="VB" AutoEventWireup="false" CodeFile="BaccaratTable_Maint.aspx.vb" Inherits="Admin_Maint" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
    <style type="text/css">
         .my-grid .x-grid-cell-inner {
            padding-top: 20px;
            padding-bottom: 20px;
            padding-left: 20px;
            padding-right: 20px;
        }
    </style>
</head>
<script language="javascript">
    var tableState = function (value) {
        var retValue = "";

        switch (value) {
            case 0:
                retValue = "<font color=blue>關閉</font>";
                break;
            case 1:
                retValue = "<font color=>正常</font>";
                break;
            case 2:
                retValue = "<font color=red>包桌</font>";
                break;
            case 3:
                retValue = "<font color=red>停用</font>";
                break;
        }
        return retValue;
    };

    var exportData = function (format) {
        App.FormatType.setValue(format);
        var store = App.GridPanel1.store;

        store.submitData(null, { isUpload: true });
    };
</script>
<body>
    <form id="form1" runat="server">
        <div>
            <ext:ResourceManager ID="ResourceManager1" runat="server" StateProvider="LocalStorage" />
            <ext:Store
                ID="Store1"
                runat="server"
                OnReadData="Store1_RefreshData"
                OnSubmitData="Store1_Submit">
                <Model>
                    <ext:Model ID="Model1" IDProperty="TableID" runat="server">
                        <Fields>
                            <ext:ModelField Name="TableID" />
                            <ext:ModelField Name="TableNumber" />
                            <ext:ModelField Name="TableState" />
                            <ext:ModelField Name="APIID" />
                            <ext:ModelField Name="VideoURL" />
                            <ext:ModelField Name="MobileVideoURL" />
                        </Fields>
                    </ext:Model>
                </Model>
            </ext:Store>

            <ext:Hidden ID="FormatType" runat="server" />
            <ext:Viewport ID="Viewport1" runat="server" Layout="BorderLayout">
                <Items>
                    <ext:GridPanel
                        ID="GridPanel1"
                        runat="server"
                        StoreID="Store1"
                        Stateful="true"
                        StateID="Admin_Maint"
                        Region="Center">
                        <ColumnModel ID="ColumnModel1" runat="server">
                            <Columns>
                                <ext:Column ID="Column1" runat="server" Text="路單編號" DataIndex="TableNumber"></ext:Column>
                                <ext:Column ID="Column3" runat="server" Text="狀態" DataIndex="TableState">
                                    <Renderer Fn="tableState" />
                                </ext:Column>
                                <ext:Column ID="Column6" runat="server" Text="API對應" DataIndex="APIID"></ext:Column>
                                <ext:Column ID="Column7" runat="server" Text="VideoURL" DataIndex="CreateDate"></ext:Column>
                                <ext:Column ID="Column8" runat="server" Text="MobileVideoURL" DataIndex="CreateDate"></ext:Column>
                            </Columns>
                        </ColumnModel>
                        <SelectionModel>
                            <ext:RowSelectionModel ID="RowSelectionModel1" runat="server" Mode="Multi" />
                        </SelectionModel>
                        <TopBar>
                            <ext:Toolbar ID="Toolbar1" runat="server">
                                <Items>
                                    <ext:SelectBox
                                        ID="selectGamArea"
                                        runat="server">
                                        <Items>
                                        </Items>
                                    </ext:SelectBox>

                                    <ext:Button ID="btnNew" runat="server" Icon="ApplicationAdd" Text="新增" />
                                    <ext:ToolbarSeparator />
                                    <ext:Button ID="btnEdit" runat="server" Icon="ApplicationEdit" Text="編輯" />
                                    <ext:ToolbarSeparator />
                                    <ext:Button ID="btnDelete" runat="server" Icon="ApplicationDelete" Text="停用">
                                        <Listeners>
                                            <Click Handler="return confirm('確定停用選擇的項目?');"></Click>
                                        </Listeners>
                                        </ext:Button> 
                                    <ext:ToolbarSeparator />
                                    <ext:ToolbarFill ID="ToolbarFill1" runat="server" />
                                    <ext:Button ID="Button4" runat="server" Text="Print" Icon="Printer" OnClientClick="window.print();" />
                                    <ext:Button ID="Button1" runat="server" Text="To XML" Icon="PageCode">
                                        <Listeners>
                                            <Click Handler="exportData('xml');" />
                                        </Listeners>
                                    </ext:Button>
                                    <ext:Button ID="Button2" runat="server" Text="To Excel" Icon="PageExcel">
                                        <Listeners>
                                            <Click Handler="exportData('xls');" />
                                        </Listeners>
                                    </ext:Button>

                                    <ext:Button ID="Button3" runat="server" Text="To CSV" Icon="PageAttach">
                                        <Listeners>
                                            <Click Handler="exportData('csv');" />
                                        </Listeners>
                                    </ext:Button>
                                </Items>
                            </ext:Toolbar>
                        </TopBar>
                    </ext:GridPanel>
                </Items>
            </ext:Viewport>

        </div>
    </form>
</body>
</html>
