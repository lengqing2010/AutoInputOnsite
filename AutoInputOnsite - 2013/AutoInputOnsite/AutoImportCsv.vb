Option Explicit On
Option Strict On

Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading.Thread

Public Class AutoImportCsv

    'File Choose use dll
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> Public Shared Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> Private Shared Function FindWindowEx( _
        ByVal parentHandle As IntPtr, _
        ByVal childAfter As IntPtr, _
        ByVal lclassName As String, _
        ByVal windowTitle As String) As IntPtr
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> Private Shared Function SendMessage( _
        ByVal hWnd As IntPtr, _
        ByVal Msg As Integer, _
        ByVal wParam As Integer, _
        ByVal lParam As StringBuilder) As Integer
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> Private Shared Function IsWindowVisible( _
        ByVal hWnd As IntPtr) As Boolean
    End Function
    Private Enum WM As Integer
        SETTEXT = &HC
    End Enum
    Private Enum BM As Integer
        CLICK = &HF5
    End Enum
    Private WithEvents BackgroundWorker As New BackgroundWorker

    Private Sub BackgroundWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker.DoWork
        Dim FileName As String = DirectCast(e.Argument, String)
        Dim hWnd As IntPtr
        Do While hWnd = IntPtr.Zero
            hWnd = FindWindow("#32770", "アップロードするファイルの選択")
            Sleep(1)
        Loop
        Dim hComboBoxEx As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "ComboBoxEx32", String.Empty)
        Dim hComboBox As IntPtr = FindWindowEx(hComboBoxEx, IntPtr.Zero, "ComboBox", String.Empty)
        Dim hEdit As IntPtr = FindWindowEx(hComboBox, IntPtr.Zero, "Edit", String.Empty)
        Do Until IsWindowVisible(hEdit)
            Sleep(1)
        Loop


        SendMessage(hEdit, WM.SETTEXT, 0, New StringBuilder(FileName))
        Dim hButton As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "Button", "開く(&O)")
        SendMessage(hButton, BM.CLICK, 0, Nothing)
    End Sub


    'IE
    Dim Ie As SHDocVw.InternetExplorerMedium

    Sub AddMsg(ByVal txt As String)
        Me.rtbxMsg.Text = Now.ToString("yyyy/MM/dd HH:mm:ss") & ":" & txt & vbNewLine & Me.rtbxMsg.Text
    End Sub



    ''' <summary>
    ''' Wait For Complete
    ''' </summary>
    ''' <param name="webApp"></param>
    ''' <remarks></remarks>
    Sub WaitComplete(ByRef webApp As SHDocVw.InternetExplorerMedium)
        System.Threading.Thread.Sleep(20)
        For i As Integer = 0 To 10
            Do Until webApp.ReadyState = WebBrowserReadyState.Complete AndAlso Not webApp.Busy
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(10)
            Loop
            System.Threading.Thread.Sleep(10)
        Next
    End Sub

    ''' <summary>
    ''' Get Element
    ''' </summary>
    ''' <param name="webApp"></param>
    ''' <param name="fraName"></param>
    ''' <param name="tagName"></param>
    ''' <param name="keyName"></param>
    ''' <param name="keyTxt"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetElementBy(ByRef webApp As SHDocVw.InternetExplorerMedium, ByVal fraName As String, ByVal tagName As String, ByVal keyName As String, ByVal keyTxt As String) As mshtml.IHTMLElement

        Dim ele As mshtml.IHTMLElement

        AddMsg("[" & keyTxt & "]--[frame：[" & fraName & "]/tagName:[" & tagName & "]/keyName:[" & keyName & "]")
        Try
            ele = GetElementByDo(webApp, fraName, tagName, keyName, keyTxt)

            'While GetElementByDo(webApp, fraName, tagName, keyName, keyTxt) Is Nothing
            '    System.Windows.Forms.Application.DoEvents()
            '    System.Threading.Thread.Sleep(1)
            'End While
            Return ele
        Catch ex As Exception
            AddMsg(ex.Message)
            'ele = GetElementByDo(webApp, fraName, tagName, keyName, keyTxt)
            Return Nothing
        End Try
    End Function

    Public Function ElementByClick(ByRef webApp As SHDocVw.InternetExplorerMedium, ByVal fraName As String, ByVal tagName As String, ByVal keyName As String, ByVal keyTxt As String) As Boolean
        AddMsg("{Click}[" & keyTxt & "]--[frame：[" & fraName & "]/tagName:[" & tagName & "]/keyName:[" & keyName & "]")
        Dim ele As mshtml.IHTMLElement = GetElementBy(webApp, fraName, tagName, keyName, keyTxt)
        If ele Is Nothing Then
            Return False
        Else
            ele.click()
            Return True
        End If
    End Function

    Public Function ElementInnerText(ByRef webApp As SHDocVw.InternetExplorerMedium, ByVal fraName As String, ByVal tagName As String, ByVal keyName As String, ByVal keyTxt As String, ByVal value As String) As Boolean
        AddMsg("{InnerText}[" & keyTxt & "]--[frame：[" & fraName & "]/tagName:[" & tagName & "]/keyName:[" & keyName & "]")
        Dim ele As mshtml.IHTMLElement = GetElementBy(webApp, fraName, tagName, keyName, keyTxt)
        If ele Is Nothing Then
            Return False
        Else
            ele.innerText = value
            Return True
        End If
    End Function

    ''' <summary>
    ''' Get Element
    ''' </summary>
    ''' <param name="webApp"></param>
    ''' <param name="fraName"></param>
    ''' <param name="tagName"></param>
    ''' <param name="keyName"></param>
    ''' <param name="keyTxt"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetElementByDo(ByRef webApp As SHDocVw.InternetExplorerMedium, ByVal fraName As String, ByVal tagName As String, ByVal keyName As String, ByVal keyTxt As String) As mshtml.IHTMLElement



        WaitComplete(webApp)



        Dim Doc As mshtml.HTMLDocument = CType(webApp.Document, mshtml.HTMLDocument)
        Dim eles As mshtml.IHTMLElementCollection

        If fraName = "" Then
            eles = Doc.getElementsByTagName(tagName)
        Else
            Dim fra As mshtml.HTMLWindow2 = GetFrameByName(webApp, fraName)
            Doc = CType(fra.document, mshtml.HTMLDocument)
            eles = Doc.getElementsByTagName(tagName)
        End If


        For Each ele As mshtml.IHTMLElement In eles
            Try

                If keyName = "innertext" Then
                    If ele.innerText = keyTxt Then
                        Return ele
                    End If
                Else
                    If ele.getAttribute(keyName).ToString = keyTxt Then
                        Return ele
                    End If
                End If

            Catch ex As Exception

            End Try

        Next

        Return Nothing

    End Function

    ''' <summary>
    ''' Get Frame
    ''' </summary>
    ''' <param name="webApp"></param>
    ''' <param name="name"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFrameByName(ByRef webApp As SHDocVw.InternetExplorerMedium, ByVal name As String) As mshtml.HTMLWindow2
        Dim Doc As mshtml.HTMLDocument = CType(webApp.Document, mshtml.HTMLDocument)
        Dim length As Integer = Doc.frames.length
        Dim frames As mshtml.FramesCollection = Doc.frames
        Dim i As Object
        For i = 0 To length - 1
            Dim frm As mshtml.HTMLWindow2 = CType(frames.item(i), mshtml.HTMLWindow2)
            If frm.name = name Then
                Return frm
            End If
        Next

        For i = 0 To length - 1
            Dim frm As mshtml.HTMLWindow2 = CType(frames.item(i), mshtml.HTMLWindow2)
            Dim wd As Object = GetFrameByName(frm, name)
            If wd IsNot Nothing Then
                Return CType(wd, mshtml.HTMLWindow2)
            End If
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' Get Frame
    ''' </summary>
    ''' <param name="webApp"></param>
    ''' <param name="name"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFrameByName(ByRef webApp As mshtml.HTMLWindow2, ByVal name As String) As mshtml.HTMLWindow2
        Dim Doc As mshtml.HTMLDocument = CType(webApp.document, mshtml.HTMLDocument)
        Dim length As Integer = Doc.frames.length
        Dim frames As mshtml.FramesCollection = Doc.frames
        Dim i As Object
        For i = 0 To length - 1
            Dim frm As mshtml.HTMLWindow2 = CType(frames.item(i), mshtml.HTMLWindow2)
            If frm.name = name Then
                Return frm
            End If
        Next

        For i = 0 To length - 1
            Dim frm As mshtml.HTMLWindow2 = CType(frames.item(i), mshtml.HTMLWindow2)
            Dim wd As Object = GetFrameByName(frm, name)
            If wd IsNot Nothing Then
                Return CType(wd, mshtml.HTMLWindow2)
            End If
        Next

        Return Nothing

    End Function


    ''' <summary>
    ''' 実行
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        Ie = New SHDocVw.InternetExplorerMedium
        'Dim authHeader As Object = "Authorization: Basic " + Convert.ToBase64String(System.Text.UnicodeEncoding.UTF8.GetBytes(String.Format("{0}:{1}", "china1\shil2", "asdf\@123"))) + "\r\n"
        Dim url As String = _
            "http://ons-ap-01d/ONS/top/scripts/Sougou_Menu.asp"

        Try
            'Ie.Navigate(url, , , , authHeader)
            Ie.Navigate(url)
            Ie.Silent = True
            Ie.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ElementByClick(Ie, "", "input", "value", "OnSiteパスワード入力へ")
        'OnSiteパスワード入力
        ElementInnerText(Ie, "", "input", "name", "strPassWord", "kaihatu")
        ElementByClick(Ie, "", "input", "value", "ログオン")
        '業務別総合メニュー
        ElementByClick(Ie, "SubHeader", "a", "innertext", "[見積]")
        '物販明細
        ElementByClick(Ie, "Main", "input", "value", "物販明細")

        '新規見積もり
SinkiMiturori:

        Dim ShellWindows As New SHDocVw.ShellWindows
        For Each childIe As SHDocVw.InternetExplorerMedium In ShellWindows
            Dim filename As String = System.IO.Path.GetFileNameWithoutExtension(Ie.FullName).ToLower()
            If filename = "iexplore" Then
                If CType(childIe, SHDocVw.InternetExplorerMedium).LocationURL.Contains("mitSearch.asp") Then
                    If CType(childIe.Document, mshtml.HTMLDocument).title = "OnSite" Then
                        If CType(childIe.Document, mshtml.HTMLDocument).url.Contains("mitSearch.asp") Then
                            WaitComplete(childIe)

                            Try
                                ElementByClick(childIe, "", "input", "value", "新規見積")
                                GoTo SinkiMituroriOK

                            Catch ex As Exception
                                GoTo SinkiMiturori
                            End Try

                        End If
                    End If
                End If
            End If
        Next
        GoTo SinkiMiturori


SinkiMituroriOK:

        System.Threading.Thread.Sleep(500)
        WaitComplete(Ie)
        '見積見出入力
        ElementInnerText(Ie, "fraMitBody", "input", "name", "strJgyCdText", "TPP5")
        ElementInnerText(Ie, "fraMitBody", "input", "name", "strTokMeiText", "420159")
        GetElementBy(Ie, "fraMitBody", "select", "name", "aryKijyunSyouhinBunrui").setAttribute("value", "A0001,サッシ,L90000")


        ElementByClick(Ie, "fraMitBody", "input", "name", "btnUtiwake")
        'System.Threading.Thread.Sleep(500)
        WaitComplete(Ie)
        'System.Threading.Thread.Sleep(500)
        WaitComplete(Ie)
        'System.Threading.Thread.Sleep(500)
        WaitComplete(Ie)
        'System.Threading.Thread.Sleep(500)
        WaitComplete(Ie)
        '見積内訳入力
WaitNY:
        Try
            Dim ele As mshtml.IHTMLElement = GetElementBy(Ie, "fraMitBody", "DIV", "classname", "ttl")
            If ele.innerText <> "見積内訳入力" Then
                GoTo WaitNY
            End If
            ElementByClick(Ie, "fraMitBody", "input", "value", "CSV取込")
        Catch ex As Exception
            GoTo WaitNY
        End Try


        WaitComplete(Ie)
        System.Threading.Thread.Sleep(500)
        Me.BackgroundWorker.RunWorkerAsync(Application.ExecutablePath)
        For Each childIe As SHDocVw.InternetExplorerMedium In ShellWindows
            Dim filename As String = System.IO.Path.GetFileNameWithoutExtension(Ie.FullName).ToLower()
            If filename = "iexplore" Then
                If CType(childIe, SHDocVw.InternetExplorerMedium).LocationURL.Contains("fileYomikomiSiji.asp") Then
                    If CType(childIe.Document, mshtml.HTMLDocument).title = "OnSite" Then
                        If CType(childIe.Document, mshtml.HTMLDocument).url.Contains("fileYomikomiSiji.asp") Then
                            WaitComplete(childIe)
                            'System.Threading.Thread.Sleep(500)
                            Try
                                WaitComplete(childIe)
                                System.Threading.Thread.Sleep(100)
                                ElementByClick(childIe, "", "input", "value", "参　照")
                                System.Threading.Thread.Sleep(100)
                            Catch ex As Exception
                                GoTo WaitNY
                            End Try
                            'WaitComplete(childIe)
                        End If
                    End If
                End If

            End If
        Next

    End Sub

    Private Sub AutoImportCsv_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim hWnd As IntPtr
        Do While hWnd = IntPtr.Zero
            hWnd = FindWindow("#32770", "windows セキュリティ")
            Sleep(1)
        Loop

   

        Dim hComboBoxEx As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "DirectUIHWND", String.Empty)
        ' Dim hComboBox As IntPtr = FindWindowEx(hComboBoxEx, IntPtr.Zero, "ComboBox", String.Empty)
        Dim CtrlNotifySink As IntPtr = FindWindowEx(hComboBoxEx, IntPtr.Zero, "CtrlNotifySink", String.Empty)

        Dim OK As IntPtr = FindWindowEx(CtrlNotifySink, IntPtr.Zero, "msctls_progress32", String.Empty)


        ' SendMessage(hEdit, WM.SETTEXT, 0, New StringBuilder("fff"))

        ' FindWindowEx(hComboBox, IntPtr.Zero, "OK", String.Empty)

    End Sub
End Class