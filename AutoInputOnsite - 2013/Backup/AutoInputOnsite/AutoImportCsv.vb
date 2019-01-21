Option Explicit On
Option Strict On

Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading.Thread

Public Class AutoImportCsv


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

    Dim Ie As New SHDocVw.InternetExplorer

    Private Sub AutoImportCsv_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        Ie.Quit()
    End Sub

    Private Sub AutoImportCsv_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Ie.Quit()
    End Sub

    Private Sub AutoImportCsv_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim authHeader As Object = "Authorization: Basic " + Convert.ToBase64String(System.Text.UnicodeEncoding.UTF8.GetBytes(String.Format("{0}:{1}", "china1\shil2", "asdf\@123"))) + "\r\n"
        Dim url As String = _
            "http://ons-ap-01d/ONS/top/scripts/Sougou_Menu.asp"
        Ie.Navigate(url, , , , authHeader)
        'Ie.Navigate(url)
        Ie.Silent = True
        Ie.Visible = True

        'WaitComplete(Ie)

        GetElementBy(Ie, "", "input", "value", "OnSiteパスワード入力へ").click()
        'WaitComplete(Ie)

        'OnSiteパスワード入力
        GetElementBy(Ie, "", "input", "name", "strPassWord").innerText = "kaihatu"
        GetElementBy(Ie, "", "input", "value", "ログオン").click()
        'WaitComplete(Ie)

        '業務別総合メニュー
        GetElementBy(Ie, "SubHeader", "a", "innertext", "[見積]").click()
        ' WaitComplete(Ie)

        '物販明細
        GetElementBy(Ie, "Main", "input", "value", "物販明細").click()
        'WaitComplete(Ie)
        ' System.Threading.Thread.Sleep(1500)


        '新規見積もり
SinkiMiturori:

        Dim ShellWindows As New SHDocVw.ShellWindows
        For Each childIe As SHDocVw.InternetExplorer In ShellWindows
            Dim filename As String = System.IO.Path.GetFileNameWithoutExtension(Ie.FullName).ToLower()
            If filename = "iexplore" Then
                If CType(childIe, SHDocVw.InternetExplorer).LocationURL.Contains("mitSearch.asp") Then
                    If CType(childIe.Document, mshtml.HTMLDocument).title = "OnSite" Then
                        If CType(childIe.Document, mshtml.HTMLDocument).url.Contains("mitSearch.asp") Then
                            WaitComplete(childIe)
                            ' System.Threading.Thread.Sleep(500)
                            Try
                                GetElementBy(childIe, "", "input", "value", "新規見積").click()
                                GoTo SinkiMituroriOK
                                'System.Threading.Thread.Sleep(100)
                            Catch ex As Exception
                                GoTo SinkiMiturori
                            End Try
                            'WaitComplete(childIe)
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
        GetElementBy(Ie, "fraMitBody", "input", "name", "strJgyCdText").innerText = "TPP5"
        GetElementBy(Ie, "fraMitBody", "input", "name", "strTokMeiText").innerText = "420159"
        GetElementBy(Ie, "fraMitBody", "select", "name", "aryKijyunSyouhinBunrui").setAttribute("value", "A0001,サッシ,L90000")


        GetElementBy(Ie, "fraMitBody", "input", "name", "btnUtiwake").click()
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
            GetElementBy(Ie, "fraMitBody", "input", "value", "CSV取込").click()
        Catch ex As Exception
            GoTo WaitNY
        End Try


        WaitComplete(Ie)
        System.Threading.Thread.Sleep(500)
        Me.BackgroundWorker.RunWorkerAsync(Application.ExecutablePath)
        For Each childIe As SHDocVw.InternetExplorer In ShellWindows
            Dim filename As String = System.IO.Path.GetFileNameWithoutExtension(Ie.FullName).ToLower()
            If filename = "iexplore" Then
                If CType(childIe, SHDocVw.InternetExplorer).LocationURL.Contains("fileYomikomiSiji.asp") Then
                    If CType(childIe.Document, mshtml.HTMLDocument).title = "OnSite" Then
                        If CType(childIe.Document, mshtml.HTMLDocument).url.Contains("fileYomikomiSiji.asp") Then
                            WaitComplete(childIe)
                            'System.Threading.Thread.Sleep(500)
                            Try
                                WaitComplete(childIe)
                                System.Threading.Thread.Sleep(100)
                                GetElementBy(childIe, "", "input", "value", "参　照").click()
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


    ''' <summary>
    ''' Wait For Complete
    ''' </summary>
    ''' <param name="webApp"></param>
    ''' <remarks></remarks>
    Sub WaitComplete(ByRef webApp As SHDocVw.InternetExplorer)
        System.Threading.Thread.Sleep(100)
        For i As Integer = 0 To 10
            Do Until webApp.ReadyState = WebBrowserReadyState.Complete AndAlso Not webApp.Busy
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(10)
            Loop
            System.Threading.Thread.Sleep(10)
        Next

    End Sub

    Public Function GetElementBy(ByRef webApp As SHDocVw.InternetExplorer, ByVal fraName As String, ByVal tagName As String, ByVal keyName As String, ByVal keyTxt As String) As mshtml.IHTMLElement
        Try
            While GetElementByDo(webApp, fraName, tagName, keyName, keyTxt) Is Nothing
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(1)
            End While

            Return GetElementByDo(webApp, fraName, tagName, keyName, keyTxt)
        Catch ex As Exception
            Return GetElementByDo(webApp, fraName, tagName, keyName, keyTxt)
        End Try


    End Function


    Public Function GetElementByDo(ByRef webApp As SHDocVw.InternetExplorer, ByVal fraName As String, ByVal tagName As String, ByVal keyName As String, ByVal keyTxt As String) As mshtml.IHTMLElement

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


    Public Function GetFrameByName(ByRef webApp As SHDocVw.InternetExplorer, ByVal name As String) As mshtml.HTMLWindow2
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

End Class