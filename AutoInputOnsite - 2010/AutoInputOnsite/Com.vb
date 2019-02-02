Imports System.Configuration
Imports System.Text
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Threading.Thread

Public Class Com
    Public Declare Function timeGetTime Lib "winmm.dll" () As Long
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Region "Windows DLL"
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

#End Region

    Public folder_Hattyuu As String = ConfigurationManager.AppSettings("Folder_Hattyuu").ToString()
    Public folder_Hattyuu_kanryou As String = ConfigurationManager.AppSettings("Folder_Hattyuu_kanryou").ToString()
    Public folder_Nouki As String = ConfigurationManager.AppSettings("Folder_Nouki").ToString()
    Public folder_Nouki_kanryou As String = ConfigurationManager.AppSettings("Folder_Nouki_kanryou").ToString()
    Public user As String = ConfigurationManager.AppSettings("User").ToString()
    Public password As String = ConfigurationManager.AppSettings("Password").ToString()
    Public url As String = ConfigurationManager.AppSettings("PasswordNyuuryoku").ToString()

    Public logFileName As String
    Public file_list_hattyuu As List(Of String)
    Public file_list_Nouki As List(Of String)

    Sub StartWK(ByRef bw As BackgroundWorker)
        If bw Is Nothing Then
            bw = New BackgroundWorker
        End If
        bw.WorkerSupportsCancellation = True
        bw.RunWorkerAsync()
    End Sub

    Sub StopWK(ByRef bw As BackgroundWorker)
        bw.CancelAsync()
        bw.Dispose()
        bw = Nothing
    End Sub

    Public Shared Sub Sleep5(ByVal Interval As Integer)
        Dim __time As DateTime = DateTime.Now
        Dim __Span As Int64 = Interval * 10000
        While (DateTime.Now.Ticks - __time.Ticks < __Span)
            Application.DoEvents()
        End While
    End Sub

    ''' <summary>
    ''' 認証情報
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Ninnsyou(Optional ByVal sec As Integer = 2)

        Dim user As String = ConfigurationManager.AppSettings("User").ToString()
        Dim password As String = ConfigurationManager.AppSettings("Password").ToString()

        If user = "" Then user = "china\shil2"
        If password = "" Then password = "asdf@123"

        Dim idx As Integer = 0

        Dim hWnd As IntPtr
        Do While hWnd = IntPtr.Zero
            hWnd = FindWindow("#32770", "windows セキュリティ")
            Sleep(10)
            If idx = sec * 100 Then
                Exit Sub
            Else
                idx += 1
            End If
        Loop

        Dim DirectUIHWND As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "DirectUIHWND", String.Empty)
        Dim CtrlNotifySink1 As IntPtr = FindWindowEx(DirectUIHWND, IntPtr.Zero, "CtrlNotifySink", String.Empty)
        Dim CtrlNotifySink2 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink1, "CtrlNotifySink", String.Empty)
        Dim CtrlNotifySink3 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink2, "CtrlNotifySink", String.Empty)
        Dim CtrlNotifySink4 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink3, "CtrlNotifySink", String.Empty)
        Dim CtrlNotifySink5 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink4, "CtrlNotifySink", String.Empty)
        Dim CtrlNotifySink6 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink5, "CtrlNotifySink", String.Empty)
        Dim CtrlNotifySink7 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink6, "CtrlNotifySink", String.Empty)
        Dim CtrlNotifySink8 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink7, "CtrlNotifySink", String.Empty)

        Dim hEdit As IntPtr = FindWindowEx(CtrlNotifySink7, IntPtr.Zero, "Edit", String.Empty)
        SendMessage(hEdit, WM.SETTEXT, 0, New StringBuilder(user))
        Dim hEdit1 As IntPtr = FindWindowEx(CtrlNotifySink8, IntPtr.Zero, "Edit", String.Empty)
        SendMessage(hEdit1, WM.SETTEXT, 0, New StringBuilder(password))
        Dim hEdit2 As IntPtr = FindWindowEx(CtrlNotifySink3, IntPtr.Zero, "Button", "OK")
        SendMessage(hEdit2, BM.CLICK, 0, Nothing)

        Ninnsyou(sec)

    End Sub

    Public Sub New(ByVal inLogFileName As String)

        logFileName = inLogFileName

        If user = "" Then user = "china\shil2"
        If password = "" Then password = "asdf@123"

        'フォルダ読み
        If Not System.IO.Directory.Exists(folder_Hattyuu) Then
            My.Computer.FileSystem.CreateDirectory(folder_Hattyuu)
        End If

        If Not System.IO.Directory.Exists(folder_Hattyuu_kanryou) Then
            My.Computer.FileSystem.CreateDirectory(folder_Hattyuu_kanryou)
        End If

        If Not System.IO.Directory.Exists(folder_Nouki) Then
            My.Computer.FileSystem.CreateDirectory(folder_Nouki)
        End If

        If Not System.IO.Directory.Exists(folder_Nouki_kanryou) Then
            My.Computer.FileSystem.CreateDirectory(folder_Nouki_kanryou)
        End If

        file_list_hattyuu = GetAllFiles(folder_Hattyuu, "*.csv")
        file_list_Nouki = GetAllFiles(folder_Nouki, "*.csv")

    End Sub

    'MESSAGE 追加
    Public Sub AddMsg(ByVal msg As String)
        'Me.RichTextBox1.Text = Now.ToString("yy/MM/dd HH:mm:ss") & "⇒:" & msg & vbNewLine & Me.RichTextBox1.Text
        If ConfigurationManager.AppSettings("Debug").ToString().ToUpper = "TRUE" Then
            WriteLog(Now.ToString("yy/MM/dd HH:mm:ss") & "⇒:" & msg, logFileName)
        End If
    End Sub

    'ログ出力
    Private Sub WriteLog(ByVal Msg As String, ByVal tmpfileName As String)
        Dim varAppPath As String
        varAppPath = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "log"
        System.IO.Directory.CreateDirectory(varAppPath)
        'TextBox1.Text = varAppPath
        Dim head As String
        head = System.DateTime.Now.Hour.ToString() + ":" + System.DateTime.Now.Minute.ToString()
        Msg = head + System.Environment.NewLine + Msg + System.Environment.NewLine
        'Dim strDate As String
        'strDate = System.DateTime.Now.ToString("yyyyMMdd")
        Dim strFile As String
        strFile = varAppPath + "\" + tmpfileName + ".log"
        Dim SW As System.IO.StreamWriter
        SW = New System.IO.StreamWriter(strFile, True)
        SW.WriteLine(Msg)
        SW.Flush()
        SW.Close()
    End Sub

    'ファイル取得
    Private Function GetAllFiles(ByVal strDirect As String, ByVal exName As String) As List(Of String)

        Dim flList As New List(Of String)

        If Not (strDirect Is Nothing) Then
            Dim mFileInfo As System.IO.FileInfo
            'Dim mDir As System.IO.DirectoryInfo
            Dim mDirInfo As New System.IO.DirectoryInfo(strDirect)
            Try
                For Each mFileInfo In mDirInfo.GetFiles(exName)
                    'Debug.Print(mFileInfo.FullName)
                    flList.Add(mFileInfo.Name)
                Next
            Catch ex As System.IO.DirectoryNotFoundException

            End Try
        End If

        Return flList

    End Function

    ''' <summary>
    ''' Wait For Complete
    ''' </summary>
    ''' <param name="webApp"></param>
    ''' <remarks></remarks>
    'Public Sub WaitComplete(ByRef webApp As SHDocVw.InternetExplorerMedium)
    '    Sleep5(100)
    '    For i As Integer = 0 To 10
    '        Do Until webApp.ReadyState = WebBrowserReadyState.Complete AndAlso Not webApp.Busy
    '            System.Windows.Forms.Application.DoEvents()
    '            Sleep5(10)
    '        Loop
    '        Sleep5(10)
    '    Next

    'End Sub

    'Get Frame
    Public Function GetFrameByName(ByRef webApp As SHDocVw.InternetExplorerMedium, ByVal name As String) As mshtml.HTMLWindow2

        SleepAndWaitComplete(webApp)

        Try
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
        Catch ex As Exception

            Return Nothing
        End Try

    End Function

    'Get Frame
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

    'Sleep App
    Sub SleepAndWaitComplete(ByRef webApp As SHDocVw.InternetExplorerMedium, Optional ByVal tmOut As Integer = 100)

        Sleep5(100)

        For i As Integer = 0 To 10
            Do Until webApp.ReadyState = WebBrowserReadyState.Complete AndAlso Not webApp.Busy
                System.Windows.Forms.Application.DoEvents()
                Sleep5(tmOut / 10)
                i = 0
            Loop

            Do Until CType(webApp.Document, mshtml.HTMLDocument).readyState.ToLower = "complete"
                System.Windows.Forms.Application.DoEvents()
                Sleep5(tmOut / 10)
                i = 0
            Loop


            Try
                Dim Doc As mshtml.HTMLDocument = CType(webApp.Document, mshtml.HTMLDocument)
                Dim length As Integer = Doc.frames.length
                Dim frames As mshtml.FramesCollection = Doc.frames
                For j As Integer = 0 To length - 1
                    Dim frm As mshtml.HTMLWindow2 = CType(frames.item(j), mshtml.HTMLWindow2)
                    Do Until frm.document.readyState.ToLower = "complete"
                        System.Windows.Forms.Application.DoEvents()
                        Sleep5(1)
                    Loop
                Next
            Catch ex As Exception

            End Try

        Next







        'Try
        '    'For i As Integer = 1 To tmOut
        '    '    Sleep5(1)
        '    '    System.Windows.Forms.Application.DoEvents()
        '    'Next

        '    For i As Integer = 0 To 100
        '        Do Until webApp.ReadyState = WebBrowserReadyState.Complete AndAlso Not webApp.Busy
        '            System.Windows.Forms.Application.DoEvents()
        '            Sleep5(1)
        '            i = 0
        '        Loop
        '        Sleep5(1)
        '    Next
        'Catch ex As Exception
        'End Try
    End Sub

    'Get element and do soming
    Public Function GetElementByDo(ByRef webApp As SHDocVw.InternetExplorerMedium, ByVal fraName As String, ByVal tagName As String, ByVal keyName As String, ByVal keyTxt As String) As mshtml.IHTMLElement

        If webApp Is Nothing Then
            Return Nothing
        End If
        SleepAndWaitComplete(webApp)
        Dim Doc As mshtml.HTMLDocument = CType(webApp.Document, mshtml.HTMLDocument)
        Dim eles As mshtml.IHTMLElementCollection
        If fraName = "" Then
            eles = Doc.getElementsByTagName(tagName)
        Else
            Dim fra As mshtml.HTMLWindow2 = GetFrameWait(webApp, fraName)
            Doc = CType(fra.document, mshtml.HTMLDocument)
            eles = Doc.getElementsByTagName(tagName)
        End If

        SleepAndWaitComplete(webApp)
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

    'Get Frame
    Public Function GetFrameWait(ByRef webApp As SHDocVw.InternetExplorerMedium, ByVal name As String) As mshtml.HTMLWindow2

        Dim fra As mshtml.HTMLWindow2

        For i As Integer = 1 To 10
            fra = GetFrameByName(webApp, name)
            If fra IsNot Nothing Then
                Return fra
            End If
            Sleep5(200)
        Next

        Return Nothing

    End Function

    'Get Element
    Public Function GetElementBy(ByRef webApp As SHDocVw.InternetExplorerMedium, ByVal fraName As String, ByVal tagName As String, ByVal keyName As String, ByVal keyTxt As String) As mshtml.IHTMLElement

        SleepAndWaitComplete(webApp)

        Try
            While GetElementByDo(webApp, fraName, tagName, keyName, keyTxt) Is Nothing
                System.Windows.Forms.Application.DoEvents()
                Sleep5(1)
            End While

            Return GetElementByDo(webApp, fraName, tagName, keyName, keyTxt)
        Catch ex As Exception
            Return GetElementByDo(webApp, fraName, tagName, keyName, keyTxt)
        End Try

    End Function


    Public Function GetElement(ByVal fra As mshtml.HTMLWindow2, ByVal tagName As String, ByVal keyName As String, ByVal keyTxt As String) As mshtml.IHTMLElement

        Dim eles As mshtml.IHTMLElementCollection
        Dim Doc As mshtml.HTMLDocument
        Doc = CType(fra.document, mshtml.HTMLDocument)
        eles = Doc.getElementsByTagName(tagName)
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

    Public Sub Sleep5(ByVal Interval As Double)
        'Dim __time As DateTime = DateTime.Now
        'Dim __Span As Int64 = Interval * 10000
        'While (DateTime.Now.Ticks - __time.Ticks < __Span)
        '    Application.DoEvents()
        'End While
        'Timer1.Interval = Interval
        'Timer1.Enabled = True
        'Timer1.Start()

        Dim Start As Long
        Start = timeGetTime
        Do While (timeGetTime < Start + CLng(Interval))
            Windows.Forms.Application.DoEvents()
            Sleep(1)
        Loop

    End Sub
End Class
