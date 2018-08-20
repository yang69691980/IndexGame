<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Admin_Add.aspx.vb" Inherits="Admin_Add" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
</head>
    <script language="javascript">
    </script>
<body>
    <form id="form1" runat="server">
        <div>
            <ext:ResourceManager ID="ResourceManager1" runat="server" StateProvider="LocalStorage" />
            <ext:FormPanel
                ID="Panel1"
                runat="server"
                Title="後端帳號新增"
                BodyPaddingSummary="5 5 0"
                Width="450"
                Frame="true"
                Resizable ="true"
                ButtonAlign="Center"
                Layout="FormLayout">
                <FieldDefaults MsgTarget="Side" LabelWidth="75" />
                <Plugins>
                    <ext:DataTip ID="DataTip1" runat="server" />
                </Plugins>
                <Items>
                    <ext:FieldSet ID="FieldSet1" runat="server"
                        Flex="1"
                        Title="基本資料"
                        Layout="AnchorLayout"
                        >
                        <Items>
                            <ext:SelectBox ID="selectCompany" Name="selectCompanyName" FieldLabel="公司選擇" Visible="false" runat="server">
                                <RemoteValidation OnValidation="CheckLoginAccountExist" />
                            </ext:SelectBox>

                            <ext:TextField ID="txtLoginAccount" runat="server" FieldLabel="登入帳號" AllowBlank="false" AutoFocus="true" IsRemoteValidation="true">
                                <RemoteValidation OnValidation="CheckLoginAccountExist" />
                            </ext:TextField>
                            <ext:TextField 
                                ID="txtLoginPassword" 
                                runat="server" 
                                FieldLabel="登入密碼" 
                                AllowBlank="false" 
                                InputType="Password" >
                                <Listeners>
                                    <ValidityChange Handler="this.next().validate();" />
                                    <Blur Handler="this.next().validate();" />
                                </Listeners>
                            </ext:TextField>
                            <ext:TextField 
                                ID="txtLoginPassword2" 
                                runat="server" 
                                MsgTarget="Side" 
                                Vtype="password" 
                                FieldLabel="確認密碼" 
                                AllowBlank="false" 
                                InputType="Password">
                                <CustomConfig>
                                    <ext:ConfigItem Name="initialPassField" Value="txtLoginPassword" Mode="Value" />
                                </CustomConfig>
                            </ext:TextField>
                            <ext:TextField ID="txtRealname" runat="server" FieldLabel="姓名" AllowBlank="true"></ext:TextField>
                            <ext:TextField ID="txtDescription" runat="server" FieldLabel="描述" AllowBlank="true"></ext:TextField>
                            <ext:SelectBox
                                FieldLabel="對應角色"
                                ID="selectAdminRole"
                                runat="server">
                            </ext:SelectBox>

                        </Items>
                    </ext:FieldSet>
                </Items>
                <Buttons>
                    <ext:Button ID="btnSave" runat="server" Text="送出" Icon="Disk">
                        <DirectEvents>
                            <Click>
                                <EventMask ShowMask="true" Msg="Wait..." MinDelay="500" />
                            </Click>
                        </DirectEvents>
                    </ext:Button>
                    <ext:Button ID="btnClose" runat="server" Text="關閉" />
                </Buttons>
            </ext:FormPanel>

        </div>
    </form>
</body>
</html>
