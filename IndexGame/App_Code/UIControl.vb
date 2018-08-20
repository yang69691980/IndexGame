Imports System
Imports Microsoft.VisualBasic

Public Module UIControl
    Public Const MainFrameNameSpace = "99Backend.MainFrame"

    Public Enum enumPanelType
        StatusPanel
        MenuPanel
        MainTabPanel
    End Enum

    Public Function GetMyPageId() As String
        Return MyPageName.Trim.ToUpper
    End Function

    Public Sub CallRootScript(Fn As String)
        Dim EM As New EventMessage

        EM.Cmd = EventMessage.enumCmd.CallRootScript
        EM.Set("Fn", Fn)

        Ext.Net.MessageBus.Default.Publish(MainFrameNameSpace, HttpContext.Current.Server.UrlEncode(XMLSerial(EM)))
    End Sub

    Public Sub CallPanelScript(Panel As enumPanelType, FnName As String)
        Dim EM As New EventMessage

        EM.Cmd = EventMessage.enumCmd.CallPanelScript
        EM.Set("Panel", Panel.ToString)
        EM.Set("Fn", FnName)

        Ext.Net.MessageBus.Default.Publish(MainFrameNameSpace, HttpContext.Current.Server.UrlEncode(XMLSerial(EM)))
    End Sub

    Public Sub NewTabToURL(Title As String, URL As String)
        Dim EM As New EventMessage
        Dim PageId As String
        Dim TmpIndex As Integer

        TmpIndex = URL.IndexOf("?")
        If TmpIndex <> -1 Then
            PageId = URL.Substring(0, TmpIndex)
        Else
            PageId = URL
        End If

        EM.Cmd = EventMessage.enumCmd.NewTabToURL
        EM.Set("Id", PageId.Trim.ToUpper)
        EM.Set("URL", URL)
        EM.Set("Title", Title)

        Ext.Net.MessageBus.Default.Publish(MainFrameNameSpace, HttpContext.Current.Server.UrlEncode(XMLSerial(EM)))
    End Sub

    Public Sub NewTabToURL2(Title As String, URL As String)
        Dim EM As New EventMessage
        Dim PageId As String
        Dim TmpIndex As Integer

        TmpIndex = URL.IndexOf("?")
        If TmpIndex <> -1 Then
            PageId = URL.Substring(0, TmpIndex)
        Else
            PageId = URL
        End If

        EM.Cmd = EventMessage.enumCmd.NewTabToURL
        EM.Set("Id", System.Guid.NewGuid.ToString)
        EM.Set("URL", URL)
        EM.Set("Title", Title)

        Ext.Net.MessageBus.Default.Publish(MainFrameNameSpace, HttpContext.Current.Server.UrlEncode(XMLSerial(EM)))
    End Sub

    Public Sub CloseActiveTab(Optional ReloadId As String = Nothing, Optional ShowMessageTitle As String = Nothing, Optional ShowMessage As String = Nothing)
        Dim EM As New EventMessage

        EM.Cmd = EventMessage.enumCmd.CloseActiveTab
        'EM.Set("CloseId", GetMyPageId)

        If String.IsNullOrEmpty(ReloadId) = False Then
            Dim PageId As String
            Dim TmpIndex As Integer

            TmpIndex = ReloadId.IndexOf("?")
            If TmpIndex <> -1 Then
                PageId = ReloadId.Substring(0, TmpIndex)
            Else
                PageId = ReloadId
            End If

            EM.Set("ReloadTab", PageId.Trim.ToUpper)
            'EM.Set("ReloadURL", ReloadId.Trim.ToUpper)
            EM.Set("ReloadURL", String.Empty)
        End If

        EM.Set("Title", ShowMessageTitle)
        EM.Set("Message", ShowMessage)

        Ext.Net.MessageBus.Default.Publish(MainFrameNameSpace, HttpContext.Current.Server.UrlEncode(XMLSerial(EM)))
    End Sub

    Public Sub CloseTab(TabId As String, Optional ReloadId As String = Nothing, Optional ShowMessageTitle As String = Nothing, Optional ShowMessage As String = Nothing)
        Dim EM As New EventMessage

        EM.Cmd = EventMessage.enumCmd.CloseTab
        EM.Set("CloseId", TabId.Trim.ToUpper)

        If String.IsNullOrEmpty(ReloadId) = False Then
            Dim PageId As String
            Dim TmpIndex As Integer

            TmpIndex = ReloadId.IndexOf("?")
            If TmpIndex <> -1 Then
                PageId = ReloadId.Substring(0, TmpIndex)
            Else
                PageId = ReloadId
            End If

            EM.Set("ReloadTab", PageId.Trim.ToUpper)
            'EM.Set("ReloadURL", ReloadId.Trim.ToUpper)
            EM.Set("ReloadURL", String.Empty)
        End If

        EM.Set("Title", ShowMessageTitle)
        EM.Set("Message", ShowMessage)

        Ext.Net.MessageBus.Default.Publish(MainFrameNameSpace, HttpContext.Current.Server.UrlEncode(XMLSerial(EM)))
    End Sub

    Public Sub ReloadTab(TabId As String)
        Dim EM As New EventMessage

        EM.Cmd = EventMessage.enumCmd.ReloadTab

        If String.IsNullOrEmpty(TabId) = False Then
            Dim PageId As String
            Dim TmpIndex As Integer

            TmpIndex = TabId.IndexOf("?")
            If TmpIndex <> -1 Then
                PageId = TabId.Substring(0, TmpIndex)
            Else
                PageId = TabId
            End If

            EM.Set("ReloadTab", PageId.Trim.ToUpper)
            'EM.Set("ReloadURL", TabId.Trim.ToUpper)
            EM.Set("ReloadURL", String.Empty)
        End If

        EM.Set("ReloadTab", TabId.Trim.ToUpper)
        Ext.Net.MessageBus.Default.Publish(MainFrameNameSpace, HttpContext.Current.Server.UrlEncode(XMLSerial(EM)))
    End Sub

    Public Sub ReloadTab(TabId As String, URL As String)
        Dim EM As New EventMessage

        EM.Cmd = EventMessage.enumCmd.ReloadTab

        If String.IsNullOrEmpty(TabId) = False Then
            Dim PageId As String
            Dim TmpIndex As Integer

            TmpIndex = TabId.IndexOf("?")
            If TmpIndex <> -1 Then
                PageId = TabId.Substring(0, TmpIndex)
            Else
                PageId = TabId
            End If

            EM.Set("ReloadTab", PageId.Trim.ToUpper)
            EM.Set("ReloadURL", URL)
        End If

        EM.Set("ReloadTab", TabId.Trim.ToUpper)
        Ext.Net.MessageBus.Default.Publish(MainFrameNameSpace, HttpContext.Current.Server.UrlEncode(XMLSerial(EM)))
    End Sub

    Public Sub PlaySound(MediaURL As String)
        Dim EM As New EventMessage

        EM.Cmd = EventMessage.enumCmd.PlaySound
        EM.Set("MediaURL", MediaURL)
        Ext.Net.MessageBus.Default.Publish(MainFrameNameSpace, HttpContext.Current.Server.UrlEncode(XMLSerial(EM)))
    End Sub

    Public Sub ReloadActiveTab()
        Dim EM As New EventMessage

        EM.Cmd = EventMessage.enumCmd.ReloadActiveTab
        Ext.Net.MessageBus.Default.Publish(MainFrameNameSpace, HttpContext.Current.Server.UrlEncode(XMLSerial(EM)))
    End Sub

    Public Sub ReloadActiveTab(URL As String)
        Dim EM As New EventMessage

        EM.Cmd = EventMessage.enumCmd.ReloadActiveTab
        EM.Set("ReloadURL", URL)
        Ext.Net.MessageBus.Default.Publish(MainFrameNameSpace, HttpContext.Current.Server.UrlEncode(XMLSerial(EM)))
    End Sub

    Public Sub Confirm(Title As String, Message As String, PageNameSpace As String, OKHandle As String, Optional CancelHandle As String = "")
        Dim _handleOK As String = ""
        Dim _handleCancel As String = ""

        If String.IsNullOrEmpty(OKHandle) = False Then
            _handleOK = PageNameSpace & "." & OKHandle & "()"
        End If

        If String.IsNullOrEmpty(CancelHandle) = False Then
            _handleCancel = PageNameSpace & "." & CancelHandle & "()"
        End If

        Ext.Net.X.Msg.Confirm(Title, Message, New Ext.Net.MessageBoxButtonsConfig() _
                                                    With {.Yes = New Ext.Net.MessageBoxButtonConfig _
                                                                With {.Handler = _handleOK, .Text = "OK"},
                                                           .No = New Ext.Net.MessageBoxButtonConfig _
                                                                With {.Handler = _handleCancel, .Text = "Cancel"}}).Show()
    End Sub

    Public Sub ShowMsg(Title As String, Message As String)
        Dim EM As New EventMessage

        EM.Cmd = EventMessage.enumCmd.ShowMsg
        EM.Set("Title", Title)
        EM.Set("Message", Message)
        Ext.Net.MessageBus.Default.Publish(MainFrameNameSpace, HttpContext.Current.Server.UrlEncode(XMLSerial(EM)))
    End Sub

    Public Sub AlertMsg(Title As String, Message As String)
        Ext.Net.X.Msg.Alert(Title, Message).Show()
    End Sub

    Public Sub Notification(Title As String, Message As String)
        Dim EM As New EventMessage

        EM.Cmd = EventMessage.enumCmd.Notification
        EM.Set("Title", Title)
        EM.Set("Message", Message)
        Ext.Net.MessageBus.Default.Publish(MainFrameNameSpace, HttpContext.Current.Server.UrlEncode(XMLSerial(EM)))
    End Sub

    Public Sub Redirect(URL As String, Optional MaskMsg As String = "Loading...")
        Ext.Net.X.Redirect(URL, MaskMsg)
        'Ext.Net.X.Js.Call("parent.Ext.net.DirectMethods.Redirect", URL)
    End Sub

    Private Function Request() As System.Web.HttpRequest
        Return HttpContext.Current.Request
    End Function

    Private Function Response() As System.Web.HttpResponse
        Return HttpContext.Current.Response
    End Function

    Private Function Server() As System.Web.HttpServerUtility
        Return HttpContext.Current.Server
    End Function

    Private Function Application() As System.Web.HttpApplicationState
        Return HttpContext.Current.Application
    End Function

    Private Function Session() As System.Web.SessionState.HttpSessionState
        Return HttpContext.Current.Session
    End Function

    <Serializable>
    Public Class EventMessage
        Public Enum enumCmd
            NewTabToURL
            NewTabToURL2
            CloseTab
            CloseActiveTab
            ReloadTab
            ReloadActiveTab
            ShowMsg
            CallPanelScript
            CallRootScript
            Notification
            PlaySound
            StopSound
        End Enum

        Public Cmd As enumCmd
        Public Params As New List(Of HeaderItemClass)

        Public Function FindKey(name As String) As HeaderItemClass
            Dim RetValue As HeaderItemClass = Nothing

            SyncLock Params
                For Each EachHHI As HeaderItemClass In Params
                    If EachHHI.Name.Trim.ToUpper = name.Trim.ToUpper Then
                        RetValue = EachHHI
                        Exit For
                    End If
                Next
            End SyncLock

            Return RetValue
        End Function

        Public Sub [Set](name As String, value As String)
            Dim HHI As HeaderItemClass

            HHI = FindKey(name)

            If IsNothing(HHI) Then
                HHI = New HeaderItemClass
                HHI.Name = name
                HHI.Value = value

                SyncLock Params
                    Params.Add(HHI)
                End SyncLock
            Else
                HHI.Value = value
            End If
        End Sub

        Public ReadOnly Property GetValue(name As String) As String
            Get
                Dim RetValue As String = Nothing
                Dim HHI As HeaderItemClass

                SyncLock Params
                    HHI = FindKey(name)
                End SyncLock

                If IsNothing(HHI) = False Then
                    RetValue = HHI.Value
                End If

                Return RetValue
            End Get
        End Property

        <Serializable()>
        Public Class HeaderItemClass
            Public Name As String
            Public Value As String
        End Class
    End Class
End Module
