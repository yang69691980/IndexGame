<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Admin_Edit.aspx.vb" Inherits="Admin_Edit" %>

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
                Title="後端帳號編輯"
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
                            </ext:SelectBox>
                            <ext:FieldContainer 
                                runat="server" 
                                FieldLabel="登入帳號" 
                                AnchorHorizontal="100%" 
                                Layout="HBoxLayout">                                       
                                <Items>
                                    <ext:Label ID="lblLoginAccount" runat="server" Text=""></ext:Label>
                                </Items>
                            </ext:FieldContainer>
                            <ext:TextField 
                                ID="txtLoginPassword" 
                                runat="server" 
                                FieldLabel="登入密碼" 
                                AllowBlank="true" 
                                InputType="Password" 
                                Note="(不修改請保留空白)">
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
                                AllowBlank="true" 
                                InputType="Password">
                                <CustomConfig>
                                    <ext:ConfigItem Name="initialPassField" Value="txtLoginPassword" Mode="Value" />
                                </CustomConfig>
                            </ext:TextField>
                            <ext:RadioGroup ID="RadioGroup3" runat="server" FieldLabel="帳戶狀態" ColumnsNumber="3" AutomaticGrouping="false">
                                <Items>
                                    <ext:Radio ID="rdoAdminStateNormal" runat="server" Name="rdoAdminState" InputValue="0" BoxLabel="正常" />
                                    <ext:Radio ID="rdoAdminStateDisable" runat="server" Name="rdoAdminState" InputValue="1" BoxLabel="停用" />
                                </Items>
                            </ext:RadioGroup>
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
