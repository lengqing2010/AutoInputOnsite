Imports System.Configuration
Imports System.Text
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Threading.Thread
'Imports Windows.Forms.Application

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


    Private Declare Function URLDownloadToFile Lib "urlmon" _
       Alias "URLDownloadToFileA" _
      (ByVal pCaller As Long, _
       ByVal szURL As String, _
       ByVal szFileName As String, _
       ByVal dwReserved As Long, _
       ByVal lpfnCB As Long) As Long

    Private Const ERROR_SUCCESS As Long = 0
    Private Const BINDF_GETNEWESTVERSION As Long = &H10
    Private Const INTERNET_FLAG_RELOAD As Long = &H80000000

    Public Function DownloadFile(ByVal sSourceUrl As String, _
    ByVal sLocalFile As String) As Boolean

        'Download the file. BINDF_GETNEWESTVERSION forces 
        'the API to download from the specified source. 
        'Passing 0& as dwReserved causes the locally-cached 
        'copy to be downloaded, if available. If the API 
        'returns ERROR_SUCCESS (0), DownloadFile returns True.
        DownloadFile = URLDownloadToFile(0&, _
                                         sSourceUrl, _
                                         sLocalFile, _
                                         BINDF_GETNEWESTVERSION, _
                                         0&) = ERROR_SUCCESS

    End Function

    'Declare Function URLDownloadToFile Lib "urlmon" Alias _
    '"URLDownloadToFileA" (ByVal pCaller As Long, _
    'ByVal szURL As String, _
    'ByVal szFileName As String, _
    'ByVal dwReserved As Long, _
    'ByVal lpfnCB As Long) As Long

    'Declare Function DeleteUrlCacheEntry Lib "wininet" _
    '    Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long

    'Sub fileDL(ByVal objIE As SHDocVw.InternetExplorerMedium, _
    'ByVal tagStr As String, _
    'ByVal saveFolder As String, _
    'Optional ByVal fileDLType As String = "pdf")

    '    Dim fileURL As String, fileName As String, savePath As String
    '    Dim cacheDel As Long, result As Long


    '    'ファイルチェック
    '    If fileURL = "" Then
    '        MsgBox("ファイルURLを取得できませんでした。")
    '        Exit Sub
    '    End If
    '    fileURL = tagStr
    '    'ファイル名
    '    fileName = "a.pdf"

    '    'ファイル保存先(+画像ファイル名）
    '    savePath = "D:\" & fileName

    '    'キャッシュクリア
    '    cacheDel = DeleteUrlCacheEntry(fileURL)

    '    'ファイルをダウンロード
    '    result = URLDownloadToFile(0, fileURL, savePath, 0, 0)

    '    If result <> 0 Then
    '        MsgBox("ダウンロードに失敗しました。")
    '    End If

    'End Sub


    Public folder_Hattyuu As String = ConfigurationManager.AppSettings("Folder_Hattyuu").ToString()
    Public folder_Hattyuu_kanryou As String = ConfigurationManager.AppSettings("Folder_Hattyuu_kanryou").ToString()
    Public folder_Nouki As String = ConfigurationManager.AppSettings("Folder_Nouki").ToString()
    Public folder_Nouki_kanryou As String = ConfigurationManager.AppSettings("Folder_Nouki_kanryou").ToString()
    Public user As String = ConfigurationManager.AppSettings("User").ToString()
    Public password As String = ConfigurationManager.AppSettings("Password").ToString()
    Public url As String = ConfigurationManager.AppSettings("PasswordNyuuryoku").ToString()
    Public pdfPath As String = ConfigurationManager.AppSettings("Pdf_Path").ToString()

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

    ' Public hatyuu_list As List(Of String)
    ' Public nouki_list As List(Of String)

    Public Sub NewWindowsCom()

        user = ConfigurationManager.AppSettings("User").ToString()
        password = ConfigurationManager.AppSettings("Password").ToString()
        If user = "" Then user = "china\shil2"
        If password = "" Then password = "asdf@123"

        Dim hatyuu_list_idx As Integer = 0
        Dim nouki_list_idx As Integer = 0

        Dim hWnd As IntPtr
        Do While hWnd = IntPtr.Zero

            '認証情報
            If hWnd = IntPtr.Zero Then
                hWnd = FindWindow("#32770", "windows セキュリティ")
                If hWnd <> IntPtr.Zero Then
                    NewNinnsyou(hWnd)
                End If
                hWnd = IntPtr.Zero
            End If

            'ファイルアップロード
            If hWnd = IntPtr.Zero Then
                hWnd = FindWindow("#32770", "アップロードするファイルの選択")
                If hWnd <> IntPtr.Zero Then
                    NewFileSantaku(hWnd, file_list_hattyuu(hatyuu_list_idx))
                    hatyuu_list_idx += 1
                End If

                hWnd = IntPtr.Zero
            End If

            'SAVE PDF
            If PdfSavePageCtrlS(file_list_Nouki(nouki_list_idx)) Then
                nouki_list_idx += 1
                hWnd = IntPtr.Zero
            End If

            Sleep5(10)

        Loop

    End Sub

    Public Sub NewNinnsyou(ByVal hWnd As IntPtr)
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
    End Sub

    Public Sub NewFileSantaku(ByVal hWnd As IntPtr, ByVal fileName As String)
        Dim hComboBoxEx As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "ComboBoxEx32", String.Empty)
        Dim hComboBox As IntPtr = FindWindowEx(hComboBoxEx, IntPtr.Zero, "ComboBox", String.Empty)
        Dim hEdit As IntPtr = FindWindowEx(hComboBox, IntPtr.Zero, "Edit", String.Empty)
        Do Until IsWindowVisible(hEdit)
            Sleep5(1)
        Loop
        SendMessage(hEdit, WM.SETTEXT, 0, New StringBuilder(fileName))
        Dim hButton As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "Button", "開く(&O)")
        SendMessage(hButton, BM.CLICK, 0, Nothing)
    End Sub

    Public Sub NewSavePdf(ByVal hWnd As IntPtr, ByVal csvFileName As String)

        If System.IO.File.Exists(pdfPath & csvFileName & ".pdf") Then
            FileSystem.Rename(pdfPath & csvFileName & ".pdf", pdfPath & csvFileName & ".pdf" & ".bk." & Now.ToString("yyyyMMddHHmmss"))
        End If

        Dim filePathComboBoxEx32 As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "ComboBoxEx32", String.Empty)
        Dim filePathComboBox As IntPtr = FindWindowEx(filePathComboBoxEx32, IntPtr.Zero, "ComboBox", String.Empty)
        Dim hEditfilePath As IntPtr = FindWindowEx(filePathComboBox, IntPtr.Zero, "Edit", String.Empty)
        SendMessage(hEditfilePath, WM.SETTEXT, 0, New StringBuilder(pdfPath & csvFileName & ".pdf"))
        Dim saveButton As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "Button", "保存")
        SendMessage(saveButton, BM.CLICK, 0, Nothing)

    End Sub

    Public Sub NewSavePdf2(ByVal hWnd As IntPtr, ByVal csvFileName As String)

        If System.IO.File.Exists(pdfPath & csvFileName & ".pdf") Then
            FileSystem.Rename(pdfPath & csvFileName & ".pdf", pdfPath & csvFileName & ".pdf" & ".bk." & Now.ToString("yyyyMMddHHmmss"))
        End If

        Dim DUIViewWndClassName As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "DUIViewWndClassName", String.Empty)
        Dim DirectUIHWND As IntPtr = FindWindowEx(DUIViewWndClassName, IntPtr.Zero, "DirectUIHWND", String.Empty)
        Dim FloatNotifySink As IntPtr = FindWindowEx(DirectUIHWND, IntPtr.Zero, "FloatNotifySink", String.Empty)

        Dim filePathComboBox As IntPtr = FindWindowEx(FloatNotifySink, IntPtr.Zero, "ComboBox", String.Empty)

        Dim hEditfilePath As IntPtr = FindWindowEx(filePathComboBox, IntPtr.Zero, "Edit", String.Empty)


        SendMessage(hEditfilePath, WM.SETTEXT, 0, New StringBuilder(pdfPath & csvFileName & ".pdf"))
        Dim saveButton As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "Button", "保存(&S)")
        SendMessage(saveButton, BM.CLICK, 0, Nothing)

    End Sub

    Public Function PdfSavePageCtrlS(ByVal csvFileName As String) As Boolean

        csvFileName = csvFileName.Split("\"c)(csvFileName.Split("\"c).Length - 1)

        Dim ShellWindows As SHDocVw.ShellWindows = New SHDocVw.ShellWindows
        For Each childIe As SHDocVw.InternetExplorerMedium In (ShellWindows)
            System.Windows.Forms.Application.DoEvents()
            Dim filename As String = System.IO.Path.GetFileNameWithoutExtension(childIe.FullName).ToLower()
            If filename = "iexplore" Then
                If CType(childIe, SHDocVw.InternetExplorerMedium).LocationURL.Contains("servlet") Then
                    SleepAndWaitComplete(childIe, 100)
                    CType(childIe, SHDocVw.InternetExplorerMedium).ExecWB(SHDocVw.OLECMDID.OLECMDID_SAVEAS, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_PROMPTUSER, "E:\aa.pdf", "E:\aa.pdf")
                    CType(childIe, SHDocVw.InternetExplorerMedium).ExecWB(SHDocVw.OLECMDID.OLECMDID_SAVE, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DODEFAULT, "E:\aa.pdf", "E:\aa.pdf")
                    Sleep5(300)

                    Dim AcroPDF As SHDocVw.InternetExplorer
                    Dim Store As Object
                    For Each AcroPDF In ShellWindows
                        If TypeName(AcroPDF.Document) = "AcroPDF" Then
                            AcroPDF.Visible = True




                            Store = AcroPDF.LocationURL




                        End If
                    Next

                    ' DownloadFile("", "")

                    Dim myDoc As mshtml.HTMLDocument = CType(childIe, SHDocVw.InternetExplorerMedium).Document

                    Dim frm As mshtml.HTMLFrameElement = CType(childIe.Document, mshtml.HTMLDocument).body.children(1)
                    DownloadFile(frm.src, "E:\aa.pdf")
                    Dim hWnd As IntPtr
                    hWnd = FindWindow("#32770", "コピーを保存")

                    If hWnd <> IntPtr.Zero Then
                        NewSavePdf(hWnd, csvFileName)
                    End If
                    'fileDL(childIe, CType(childIe, SHDocVw.InternetExplorerMedium).LocationURL, "D:\")

                    If hWnd = IntPtr.Zero Then
                        hWnd = FindWindow("#32770", "名前を付けて保存")
                        NewSavePdf2(hWnd, csvFileName)
                    End If



                    childIe.Quit()

                    Return True

                End If
            End If
        Next
        ShellWindows = Nothing
        Return False

    End Function


    '''' <summary>
    '''' 認証情報
    '''' </summary>
    '''' <remarks></remarks>
    'Public Sub Ninnsyou(Optional ByVal sec As Integer = 2)

    '    Dim user As String = ConfigurationManager.AppSettings("User").ToString()
    '    Dim password As String = ConfigurationManager.AppSettings("Password").ToString()

    '    If user = "" Then user = "china\shil2"
    '    If password = "" Then password = "asdf@123"

    '    Dim idx As Integer = 0

    '    Dim hWnd As IntPtr
    '    Do While hWnd = IntPtr.Zero
    '        hWnd = FindWindow("#32770", "windows セキュリティ")
    '        Sleep5(10)
    '        If idx = sec * 100 Then
    '            Exit Sub
    '        Else
    '            idx += 1
    '        End If
    '    Loop

    '    Dim DirectUIHWND As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "DirectUIHWND", String.Empty)
    '    Dim CtrlNotifySink1 As IntPtr = FindWindowEx(DirectUIHWND, IntPtr.Zero, "CtrlNotifySink", String.Empty)
    '    Dim CtrlNotifySink2 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink1, "CtrlNotifySink", String.Empty)
    '    Dim CtrlNotifySink3 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink2, "CtrlNotifySink", String.Empty)
    '    Dim CtrlNotifySink4 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink3, "CtrlNotifySink", String.Empty)
    '    Dim CtrlNotifySink5 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink4, "CtrlNotifySink", String.Empty)
    '    Dim CtrlNotifySink6 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink5, "CtrlNotifySink", String.Empty)
    '    Dim CtrlNotifySink7 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink6, "CtrlNotifySink", String.Empty)
    '    Dim CtrlNotifySink8 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink7, "CtrlNotifySink", String.Empty)

    '    Dim hEdit As IntPtr = FindWindowEx(CtrlNotifySink7, IntPtr.Zero, "Edit", String.Empty)
    '    SendMessage(hEdit, WM.SETTEXT, 0, New StringBuilder(user))
    '    Dim hEdit1 As IntPtr = FindWindowEx(CtrlNotifySink8, IntPtr.Zero, "Edit", String.Empty)
    '    SendMessage(hEdit1, WM.SETTEXT, 0, New StringBuilder(password))
    '    Dim hEdit2 As IntPtr = FindWindowEx(CtrlNotifySink3, IntPtr.Zero, "Button", "OK")
    '    SendMessage(hEdit2, BM.CLICK, 0, Nothing)

    '    Ninnsyou(sec)

    'End Sub

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

    Public Shared Sub Sleep5(ByVal Interval As Double)
        Dim Start As Long
        Start = timeGetTime
        Do While (timeGetTime < Start + CLng(Interval))
            Windows.Forms.Application.DoEvents()
            Sleep(1)
        Loop
    End Sub



    '    Private Structure FILETIME

    '        Dim dwLowDateTime As Long
    '        Dim dwHighDateTime As Long
    '    End Structure

    '    Private Const WAIT_ABANDONED& = &H80&
    '    Private Const WAIT_ABANDONED_0& = &H80&
    '    Private Const WAIT_FAILED& = -1&
    '    Private Const WAIT_IO_COMPLETION& = &HC0&
    '    Private Const WAIT_OBJECT_0& = 0
    '    Private Const WAIT_OBJECT_1& = 1
    '    Private Const WAIT_TIMEOUT& = &H102&
    '    Private Const INFINITE = &HFFFF
    '    Private Const ERROR_ALREADY_EXISTS = 183&
    '    Private Const QS_HOTKEY& = &H80
    '    Private Const QS_KEY& = &H1
    '    Private Const QS_MOUSEBUTTON& = &H4
    '    Private Const QS_MOUSEMOVE& = &H2
    '    Private Const QS_PAINT& = &H20
    '    Private Const QS_POSTMESSAGE& = &H8
    '    Private Const QS_SENDMESSAGE& = &H40
    '    Private Const QS_TIMER& = &H10
    '    Private Const QS_MOUSE& = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
    '    Private Const QS_INPUT& = (QS_MOUSE Or QS_KEY)
    '    Private Const QS_ALLEVENTS& = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
    '    Private Const QS_ALLINPUT& = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
    '    Private Const UNITS = 4294967296.0#
    '    Private Const MAX_LONG = -2147483648.0#
    '    Private Declare Function CreateWaitableTimer _
    '    Lib "kernel32" _
    '    Alias "CreateWaitableTimerA" (ByVal lpSemaphoreAttributes As Long, _
    '    ByVal bManualReset As Long, _
    '    ByVal lpName As String) As Long


    '    Private Declare Function OpenWaitableTimer _
    '    Lib "kernel32" _
    '    Alias "OpenWaitableTimerA" (ByVal dwDesiredAccess As Long, _
    '    ByVal bInheritHandle As Long, _
    '    ByVal lpName As String) As Long


    '    Private Declare Function SetWaitableTimer _
    '    Lib "kernel32" (ByVal hTimer As Long, _
    '    ByVal lpDueTime As FILETIME, _
    '    ByVal lPeriod As Long, _
    '    ByVal pfnCompletionRoutine As Long, _
    '    ByVal lpArgToCompletionRoutine As Long, _
    '    ByVal fResume As Long) As Long


    '    Private Declare Function CancelWaitableTimer Lib "kernel32" (ByVal hTimer As Long)


    '    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


    '    Private Declare Function WaitForSingleObject _
    '    Lib "kernel32" (ByVal hHandle As Long, _
    '    ByVal dwMilliseconds As Long) As Long


    '    Private Declare Function MsgWaitForMultipleObjects _
    '    Lib "user32" (ByVal nCount As Long, _
    '    ByVal pHandles As Long, _
    '    ByVal fWaitAll As Long, _
    '    ByVal dwMilliseconds As Long, _
    '    ByVal dwWakeMask As Long) As Long


    '    Private mlTimer As Long


    '    Private Sub Class_Terminate()
    '        On Error Resume Next
    '        If mlTimer <> 0 Then CloseHandle(mlTimer)
    '    End Sub


    '    Public Sub Wait(ByVal MilliSeconds As Long)
    '        On Error GoTo ErrHandler
    '        Dim ft As FILETIME
    '        Dim lBusy As Long
    '        Dim lRet As Long
    '        Dim dblDelay As Double
    '        Dim dblDelayLow As Double
    '        mlTimer = CreateWaitableTimer(0, True, Application.EXEName & "Timer" & Format$(Now(), "NNSS"))
    '        If Err.LastDllError <> ERROR_ALREADY_EXISTS Then
    '            ft.dwLowDateTime = -1
    '            ft.dwHighDateTime = -1
    '            lRet = SetWaitableTimer(mlTimer, ft, 0, 0, 0, 0)
    '        End If
    '        dblDelay = CDbl(MilliSeconds) * 10000.0#
    '        ft.dwHighDateTime = -CLng(dblDelay / UNITS) - 1
    '        dblDelayLow = -UNITS * (dblDelay / UNITS - Fix(CStr(dblDelay / UNITS)))
    '        If dblDelayLow < MAX_LONG Then dblDelayLow = UNITS + dblDelayLow
    '        ft.dwLowDateTime = CLng(dblDelayLow)
    '        lRet = SetWaitableTimer(mlTimer, ft, 0, 0, 0, False)
    '        Do
    '            lBusy = MsgWaitForMultipleObjects(1, mlTimer, False, INFINITE, QS_ALLINPUT&)
    '            Windows.Forms.Application.DoEvents()
    '        Loop Until lBusy = WAIT_OBJECT_0
    '        CloseHandle(mlTimer)
    '        mlTimer = 0
    '        Exit Sub
    'ErrHandler:
    '        Err.Raise(Err.Number, Err.Source, "[clsWaitableTimer.Wait]" & Err.Description)
    '    End Sub

End Class
