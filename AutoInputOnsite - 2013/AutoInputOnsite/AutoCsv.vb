Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading.Thread

Public Class AutoCsv

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function FindWindowEx( _
        ByVal parentHandle As IntPtr, _
        ByVal childAfter As IntPtr, _
        ByVal lclassName As String, _
        ByVal windowTitle As String) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function SendMessage( _
        ByVal hWnd As IntPtr, _
        ByVal Msg As Integer, _
        ByVal wParam As Integer, _
        ByVal lParam As StringBuilder) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function IsWindowVisible( _
        ByVal hWnd As IntPtr) As Boolean
    End Function

    Private Enum WM As Integer
        SETTEXT = &HC
    End Enum

    Private Enum BM As Integer
        CLICK = &HF5
    End Enum

    'Private WithEvents WebBrowser As New WebBrowser
    Private WithEvents BackgroundWorker As New BackgroundWorker

    Public Enum AcStepType As Integer
        ErrPage = 0
        PasswordInput = 1
        MainMenu = 2
        MitumoriChoose = 3
        sonota = 999
    End Enum

    Public AcStep As Integer = 0

    Private Sub AutoCsv_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim authData() As Byte = System.Text.UnicodeEncoding.UTF8.GetBytes(String.Format("{0}:{1}", "china\shil2", "asdf\@123"))
        Dim authHeader As String = "Authorization: Basic " + Convert.ToBase64String(authData) + "\r\n"
        'Dim uri As System.Uri = New Uri("http://ons-ap-01d/ONS/top/scripts/Sougou_Menu.asp")
        'Me.WebBrowser1.Navigate(uri, "", Nothing, authHeader)
        'Me.WebBrowser1.Navigate(Uri)
        'http://ons-ap-01d/ONS/top/scripts/Sougou_Menu.asp


        'Dim uri As System.Uri = New Uri("http://ons-ap-01d/ONS/top/scripts/Sougou_Menu.asp")
        'Dim uri As System.Uri = New Uri("http://bpn-ap-01d/ONS/Top/Scripts/Sougou_Senimenu.asp?strBpnFlg=1&strSenisaki=http://bpn-ap-01d/ons/mit/Scripts/mitMainFr.asp")

        'Dim uri As System.Uri = New Uri("http://bpn-ap-01d/ons/mit/Scripts/mitMainFr.asp?strServerName=&strQualtyTokCd=&strQualtyJgyCd=&strBpnFlg=1")
        'Me.WebBrowser1.Navigate(uri)

        Dim uri As System.Uri = New Uri("http://bpn-ap-01d/ons/mit/Scripts/midasiSinkiNyuuryoku.asp?strMitKbn=&strServerName=&strBpnFlg=1&datNoCacheDummy=2018%2F12%2F26+16%3A07%3A46")
        Me.WebBrowser1.Navigate(uri, "", Nothing, authHeader)
        'POPUP
        Dim ActiveX As SHDocVw.WebBrowser = Me.WebBrowser1.ActiveXInstance
        AddHandler ActiveX.NewWindow2, AddressOf WebBrowser_ActiveX_NewWindow2

    End Sub

    Private Sub WebBrowser_ActiveX_NewWindow2(ByRef ppDisp As Object, ByRef Cancel As Boolean)
        Dim popup As RawBrowserPopup = New RawBrowserPopup()
        popup.Visible = True
        ppDisp = popup.WebBrowser.ActiveXInstance
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted

        ''入力エラー
        If AcStep = AcStepType.ErrPage Then


            ''fraMitBody
            'Dim JgyCdText As HtmlElement = GetElementBy(WebBrowser1, "fraMitBody", "input", "name", "strJgyCdText")
            'JgyCdText.InnerText = "TPP5"


            'Dim TokMeiText As HtmlElement = GetElementBy(WebBrowser1, "fraMitBody", "input", "name", "strTokMeiText")
            'TokMeiText.InnerText = "420159"

            '    'OnSiteパスワード入力へ
            '    Dim btn As HtmlElement = GetElementByInputTxt(WebBrowser1, "input", "value", "OnSiteパスワード入力へ")
            '    btn.InvokeMember("click")
            '    AcStep = AcStepType.PasswordInput

            'ElseIf AcStep = AcStepType.PasswordInput Then 'OnSiteパスワード入力
            '    '
            '    Dim password As HtmlElement = GetElementByInputTxt(WebBrowser1, "input", "name", "strPassWord")
            '    password.InnerText = "kaihatu"
            '    Dim btn As HtmlElement = GetElementByInputTxt(WebBrowser1, "input", "value", "ログオン")
            '    btn.InvokeMember("click")
            '    AcStep = AcStepType.MainMenu

            'ElseIf AcStep = AcStepType.MainMenu Then        'MENU
            '    Dim aMitumori As HtmlElement = GetElementBy(WebBrowser1, "SubHeader", "a", "innertext", "[見積]")
            '    aMitumori.InvokeMember("click")
            '    AcStep = AcStepType.MitumoriChoose

            'ElseIf AcStep = AcStepType.MitumoriChoose Then  'MENU

            '    Dim btn As HtmlElement = GetElementBy(WebBrowser1, "Main", "input", "value", "物販明細")
            '    btn.InvokeMember("click")
            '    AcStep = AcStepType.sonota


        End If

    End Sub

    Public Function GetElementByInputTxt(ByRef webApp As WebBrowser, ByVal tagName As String, ByVal keyName As String, ByVal keyTxt As String) As HtmlElement
        For Each ele As HtmlElement In webApp.Document.GetElementsByTagName(tagName)
            Try
                If ele.GetAttribute(keyName) = keyTxt Then
                    Return ele
                End If
            Catch ex As Exception
            End Try
        Next
        Return Nothing
    End Function

    Public Function GetElementBy(ByRef webApp As WebBrowser, ByVal fraName As String, ByVal tagName As String, ByVal keyName As String, ByVal keyTxt As String) As HtmlElement

        Dim eles As HtmlElementCollection

        If fraName = "" Then
            eles = webApp.Document.GetElementsByTagName(tagName)
        Else

            Dim fra As HtmlWindow = GetFrameByName(webApp, fraName)

            eles = fra.Document.GetElementsByTagName(tagName)
        End If


        For Each ele As HtmlElement In eles
            Try

                If keyName = "innertext" Then
                    If ele.InnerText = keyTxt Then
                        Return ele
                    End If
                Else
                    If ele.GetAttribute(keyName) = keyTxt Then
                        Return ele
                    End If
                End If

            Catch ex As Exception

            End Try

        Next

        Return Nothing

    End Function

    Public Function GetFrameByName(ByRef webApp As WebBrowser, ByVal name As String) As HtmlWindow

        For Each fra As HtmlWindow In webApp.Document.Window.Frames
            If fra.Name = name Then
                Return fra
            Else
                If GetFrameByName(fra, name) IsNot Nothing Then
                    Return GetFrameByName(fra, name)
                End If
            End If
        Next

        Return Nothing

    End Function


    Public Function GetFrameByName(ByRef webApp As HtmlWindow, ByVal name As String) As HtmlWindow

        For Each fra As HtmlWindow In webApp.Document.Window.Frames
            If fra.Name = name Then
                Return fra
            Else
                If GetFrameByName(fra, name) IsNot Nothing Then
                    Return GetFrameByName(fra, name)
                End If
            End If
        Next

        Return Nothing

    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim JgyCdText As HtmlElement = GetElementBy(WebBrowser1, "fraMitBody", "input", "name", "strJgyCdText")
        'JgyCdText.InnerText = "TPP5"


        'Dim TokMeiText As HtmlElement = GetElementBy(WebBrowser1, "fraMitBody", "input", "name", "strTokMeiText")
        'TokMeiText.InnerText = "420159"


        'Dim aryKijyunSyouhinBunrui As HtmlElement = GetElementBy(WebBrowser1, "fraMitBody", "select", "name", "aryKijyunSyouhinBunrui")
        'aryKijyunSyouhinBunrui.SetAttribute("value", "A0001,サッシ,L90000")

        Dim JgyCdText As HtmlElement = GetElementBy(WebBrowser1, "", "input", "name", "strJgyCdText")
        JgyCdText.InnerText = "TPP5"


        Dim TokMeiText As HtmlElement = GetElementBy(WebBrowser1, "", "input", "name", "strTokMeiText")
        TokMeiText.InnerText = "420159"


        Dim aryKijyunSyouhinBunrui As HtmlElement = GetElementBy(WebBrowser1, "", "select", "name", "aryKijyunSyouhinBunrui")
        aryKijyunSyouhinBunrui.SetAttribute("value", "A0001,サッシ,L90000")

    End Sub
End Class
