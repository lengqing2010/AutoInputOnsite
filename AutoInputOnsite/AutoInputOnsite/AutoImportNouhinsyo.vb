Option Explicit On
Option Strict On

Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading.Thread
Imports System.Configuration

Public Class AutoImportNouhinsyo





    Public _ProBar As Decimal = 0
    Public lv1 As Decimal
    Public lv2 As Decimal
    Public Property ProBar() As Integer
        Get
            If _ProBar < 100 Then
                Return CInt(Int(_ProBar))
            Else
                Return 100
            End If

        End Get
        Set(ByVal value As Integer)
            _ProBar = value
        End Set

    End Property

    Public Sub AddProBar(ByVal x As Decimal)
        _ProBar += x
    End Sub

    Public insatu As Boolean = False


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

#Region "閉じる"
    Private Sub AutoImportCsv_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        Try
            Ie.Quit()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub AutoImportCsv_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            Ie.Quit()
        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "認証情報"



    'Private Sub NinnsyouBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles NinnsyouBackgroundWorker.DoWork
    '    Pub_Com.Ninnsyou()
    'End Sub

    '''' <summary>
    '''' 認証情報
    '''' </summary>
    '''' <remarks></remarks>
    'Private Sub Ninnsyou()
    '    Dim user As String = ConfigurationManager.AppSettings("User").ToString()
    '    Dim password As String = ConfigurationManager.AppSettings("Password").ToString()
    '    If user = "" Then
    '        user = "china\shil2"
    '    End If
    '    If password = "" Then
    '        password = "asdf@123"
    '    End If

    '    Dim hWnd As IntPtr
    '    Do While hWnd = IntPtr.Zero
    '        hWnd = FindWindow("#32770", "windows セキュリティ")
    '        Sleep(1)
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

    '    Ninnsyou()
    'End Sub
    ''' <summary>
    ''' 認証情報
    ''' </summary>
    ''' <remarks></remarks>
    'Public Sub Ninnsyou()

    '    Dim user As String = ConfigurationManager.AppSettings("User").ToString()
    '    Dim password As String = ConfigurationManager.AppSettings("Password").ToString()

    '    If user = "" Then user = "china\shil2"
    '    If password = "" Then password = "asdf@123"

    '    Dim idx As Integer = 0

    '    Dim hWnd As IntPtr
    '    Do While hWnd = IntPtr.Zero
    '        hWnd = FindWindow("#32770", "windows セキュリティ")
    '        Sleep(10)
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

    '    Ninnsyou()

    'End Sub



#End Region

#Region "自動実行"

    ' Private WithEvents BackgroundWorker As BackgroundWorker
    ' Private WithEvents NinnsyouBackgroundWorker As BackgroundWorker

    'Private WithEvents PDFBackgroundWorker As BackgroundWorker
    'Private WithEvents PDFMessageBackgroundWorker As BackgroundWorker


    Public Ie As SHDocVw.InternetExplorerMedium
    Public Pub_Com As Com

    '実行
    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
        DoAll()
    End Sub



    '実行 MAIN
    Public Sub DoAll()

        'SavePdf("aaaa.csv")
        Pub_Com = New Com("納期指定発注" & Now.ToString("yyyyMMddHHmmss"))
        If Pub_Com.file_list_Nouki.Count = 0 Then
            ProBar = 100
            Exit Sub
        End If

        lv1 = CDec(90 / Pub_Com.file_list_Nouki.Count)

        '一回目
        Dim firsOpenKbn As Boolean = True

        '＊＊＊New Object
        ' NinnsyouBackgroundWorker = New BackgroundWorker
        'BackgroundWorker = New BackgroundWorker

        Ie = New SHDocVw.InternetExplorerMedium

        Dim authHeader As Object = "Authorization: Basic " + _
        Convert.ToBase64String(System.Text.UnicodeEncoding.UTF8.GetBytes(String.Format("{0}:{1}", Pub_Com.user, Pub_Com.password))) + "\r\n"

        '＊＊＊ OnSiteパスワード入力画面
        Ie.Navigate(Pub_Com.url, , , , authHeader)
        Ie.Silent = True
        Ie.Visible = True

        '認証
        'Dim thr As New Threading.Thread(AddressOf Ninnsyou)
        'thr.Start()
        ProBar = 5
        '＊＊＊ログイン
        DoStep1_Login()


        'thr.Abort()
        ProBar = 10
        '  Pub_Com.StopWK(NinnsyouBackgroundWorker)

        'Dim idx As Integer = 1

        'CSV ファイルｓ 取込
        For zzz As Integer = 0 To Pub_Com.file_list_Nouki.Count - 1

            lv2 = lv1 / 15

            Dim csvFileName As String = Pub_Com.file_list_Nouki(zzz).ToString.Trim
            Dim csvFileNames() As String = Pub_Com.file_list_Nouki(zzz).ToString.Trim.Split("-"c)
            Dim 事業所, 得意先, 下店, 現場名, 備考, 日付連番 As String
            事業所 = csvFileNames(0)
            得意先 = csvFileNames(1)
            下店 = csvFileNames(2)
            現場名 = csvFileNames(3)
            備考 = csvFileNames(4)
            日付連番 = csvFileNames(5)


            If firsOpenKbn = False Then
                Pub_Com.GetElementBy(Ie, "fraHead", "input", "value", "絞込検索").click()
                Pub_Com.SleepAndWaitComplete(Ie)
            End If
            firsOpenKbn = False
            AddProBar(lv2) '1


            Pub_Com.AddMsg("取込：" & Pub_Com.file_list_Nouki(zzz).ToString.Trim)

            'Dim thr2 As New Threading.Thread(AddressOf Ninnsyou)
            'thr2.Start()
            '見積検索
            Pub_Com.AddMsg("見積検索")
            DoStep1_PoupuSentaku(事業所, 得意先, 下店, 現場名, 備考, 日付連番, csvFileName)
            Pub_Com.SleepAndWaitComplete(Ie)
            AddProBar(lv2) '2
            'thr2.Abort()

            '納期日設定
            If Not DoStep2_Set() Then
                Continue For
            End If
            Pub_Com.SleepAndWaitComplete(Ie)
            Pub_Com.SleepAndWaitComplete(Ie)
            Pub_Com.SleepAndWaitComplete(Ie)

            '該当データがありません NEXT
            Dim fraTmp As mshtml.HTMLWindow2 = Pub_Com.GetFrameByName(Ie, "fraHyou")
            If fraTmp IsNot Nothing Then
                If fraTmp.document.body.innerText.IndexOf("該当データがありません") >= 0 Then
                    Continue For
                End If
            End If


            'CSVファイル内容取込
            Dim strData As String() = System.IO.File.ReadAllLines(Pub_Com.folder_Nouki & csvFileName)
            Dim code As String = ""
            Dim nouki As String = ""

            AddProBar(lv2) '3

            'CSV LINES
            For jjj As Integer = 0 To strData.Length - 1

                If strData(jjj).Trim <> "" Then

                    'コード 納期
                    code = strData(jjj).Split(","c)(1).Trim
                    nouki = CDate(strData(jjj).Split(","c)(2).Trim).ToString("yyyy/MM/dd")


                    Dim fra As mshtml.HTMLWindow2 = Pub_Com.GetFrameWait(Ie, "fraMitBody")
                    Dim Doc As mshtml.HTMLDocument = CType(fra.document, mshtml.HTMLDocument)
                    Dim eles As mshtml.IHTMLElementCollection = Doc.getElementsByTagName("input")

                    Dim cbEles As mshtml.IHTMLElementCollection = Doc.getElementsByName("strMeisaiKey")
                    Dim nouhinDateEles As mshtml.IHTMLElementCollection = Doc.getElementsByName("strSiteiNouhinDate")

                    For i As Integer = 0 To cbEles.length - 1
                        Dim tr As mshtml.IHTMLTableRow = CType(CType(cbEles.item(i), mshtml.IHTMLElement).parentElement.parentElement, mshtml.IHTMLTableRow)
                        Dim td As mshtml.HTMLTableCell = CType(tr.cells.item(1), mshtml.HTMLTableCell)
                        Dim table As mshtml.IHTMLTable = CType(CType(cbEles.item(i), mshtml.IHTMLElement).parentElement.parentElement.parentElement.parentElement, mshtml.IHTMLTable)

                        Dim isHaveDate As Boolean = False

                        If td.innerText = code Then
                            Dim sel As mshtml.IHTMLSelectElement = CType(nouhinDateEles.item(i), mshtml.IHTMLSelectElement)

                            For j As Integer = 0 To sel.length
                                If CType(sel.item(j), mshtml.IHTMLOptionElement).value.IndexOf(nouki) > 0 Then
                                    CType(sel.item(j), mshtml.IHTMLOptionElement).selected = True
                                    isHaveDate = True
                                    Exit For
                                End If
                            Next


                            If Not isHaveDate Then
                                MsgBox("納品希望日：[" & nouki & "]がありません")
                                Exit Sub
                            End If
                        End If

                    Next

                End If

            Next

            AddProBar(lv2) '4

            Pub_Com.GetElementBy(Ie, "fraMitBody", "select", "name", "strBukkenKbn").setAttribute("value", "01")
            Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "発　注").click()
            Pub_Com.SleepAndWaitComplete(Ie)
            Pub_Com.SleepAndWaitComplete(Ie)
            Pub_Com.SleepAndWaitComplete(Ie)
            AddProBar(lv2) '5



            Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "発注結果照会へ").click()
            Pub_Com.SleepAndWaitComplete(Ie)
            AddProBar(lv2) '6

            'If PDFMessageBackgroundWorker Is Nothing Then
            '    PDFMessageBackgroundWorker = New BackgroundWorker
            '    Me.PDFMessageBackgroundWorker.RunWorkerAsync(csvFileName)
            'End If
            'PDFBackgroundWorker = New BackgroundWorker
            'Me.PDFBackgroundWorker.RunWorkerAsync(csvFileName)
            'Sai:






            'Dim cie As SHDocVw.InternetExplorerMedium = GetPopupWindowPDF()

            ' WaitWindowPDF()

            'Pub_Com.SleepAndWaitComplete(Ie)
            'AddProBar(lv2) '7

            'Com.Sleep5(3000)





            'AddProBar(lv2) '9

            'Dim ShellWindows As New SHDocVw.ShellWindows
            'For Each childIe As SHDocVw.InternetExplorerMedium In ShellWindows
            '    Dim filename As String = System.IO.Path.GetFileNameWithoutExtension(Ie.FullName).ToLower()
            '    If filename = "iexplore" Then
            '        If CType(childIe, SHDocVw.InternetExplorerMedium).LocationURL.Contains("servlet") Then
            '            'childIe.Document
            '            'Dim sdstr As System.Windows.Forms.SendKeys
            '            CType(childIe, SHDocVw.InternetExplorerMedium).ExecWB(SHDocVw.OLECMDID.OLECMDID_SAVEAS, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_PROMPTUSER)
            '            'Dim frames As mshtml.FramesCollection = CType(CType(childIe, SHDocVw.InternetExplorerMedium).Document, mshtml.HTMLDocument).frames

            '            Pub_Com.SleepAndWaitComplete(Ie, 100)
            '            SavePdf(csvFileName)
            '            'CType(frames.item(1), mshtml.HTMLWindow2).focus()

            '            'SendKeys.Send("(^+{s})")


            '            '    If CType(childIe.Document, mshtml.HTMLDocument).title = titleKey Then
            '            '        If CType(childIe.Document, mshtml.HTMLDocument).url.Contains(fileNameKey) Then
            '            '            Pub_Com.WaitComplete(childIe)
            '            '            Com.Sleep5(500)
            '            '            Return childIe
            '            '        End If
            '            '    End If
            '        End If
            '    End If
            'Next
            'DownLoadPdf(csvFileName)
            If insatu Then
                Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "結果印刷").click()

                'Exit Sub

                'Pub_Com.SleepAndWaitComplete(Ie)
                'Pub_Com.SleepAndWaitComplete(Ie)
                'System.Threading.Thread.Sleep(5000)
                Com.Sleep5(10000)
                'DownLoadPdf(csvFileName)
                'CheckForIllegalCrossThreadCalls = True
                'Gsz_Thread_Test = New Threading.Thread(AddressOf DownLoadPdf)
                'Gsz_Thread_Test.Start(csvFileName)
                'Com.Sleep5(2000)
                'While Gsz_Thread_Test.IsAlive
                '    Com.Sleep5(1000)
                'End While

            End If


            AddProBar(lv2) '10
            Pub_Com.AddMsg("移動CSV：" & csvFileName & "→" & Pub_Com.folder_Nouki_kanryou)
            If System.IO.File.Exists(Pub_Com.folder_Nouki_kanryou & csvFileName) Then
                FileSystem.Rename(Pub_Com.folder_Nouki_kanryou & csvFileName, Pub_Com.folder_Nouki_kanryou & csvFileName & ".bk." & Now.ToString("yyyyMMddHHmmss"))
            End If
            Sleep(2000)
            System.IO.File.Move(Pub_Com.folder_Nouki & csvFileName, Pub_Com.folder_Nouki_kanryou & csvFileName)
            AddProBar(lv2) '11



            Pub_Com.GetElementBy(Ie, "fraMitMenu", "a", "innertext", "[見積一覧を再表示]").click()
            Pub_Com.SleepAndWaitComplete(Ie)




            ' WaitWindowPDF()
            'PDFBackgroundWorker.Dispose()

        Next





        ' PDFMessageBackgroundWorker.Dispose()

        Ie.Quit()

        ProBar = 100

        'MsgBox("完了")


    End Sub

    Public Delegate Sub ToThread(ByVal csvFileName As String)
    Dim Gsz_Thread_Test As Threading.Thread


    'Public Sub DownLoadPdfThread(ByVal csvFileName As Object)
    '    Dim ivo As New ToThread(AddressOf DownLoadPdf)
    '    Invoke(ivo, csvFileName.ToString)
    'End Sub

    '    Public Sub DownLoadPdf(ByVal csvFileName As Object)


    '        Dim cie As SHDocVw.InternetExplorerMedium = GetPopupWindowPDF()
    '        'Pub_Com.SleepAndWaitComplete(cie)
    '        'Pub_Com.SleepAndWaitComplete(cie)
    '        Com.Sleep5(1000)
    '        Dim waitT As Integer = 2000
    'reD:
    '        If Not ClosePdrErr() Then
    '            Try
    '                ClosePdrErr()
    '                Com.Sleep5(500)
    '                cie.Refresh()
    '                Com.Sleep5(1000)
    '                Dim hWnd2 As IntPtr
    '                hWnd2 = FindWindow("#32770", "Windows Internet Explorer")

    '                Dim hButton As IntPtr = FindWindowEx(hWnd2, IntPtr.Zero, "Button", "再試行(&R)")
    '                SendMessage(hButton, BM.CLICK, 0, Nothing)
    '                Com.Sleep5(waitT)
    '                waitT = waitT + 1000
    '            Catch ex As Exception

    '            End Try


    '        End If

    '        If Not ClosePdrErr() Then
    '            ClosePdrErr()
    '            Com.Sleep5(500)
    '            cie.Refresh()
    '            Com.Sleep5(1000)
    '            Dim hWnd2 As IntPtr
    '            hWnd2 = FindWindow("#32770", "Windows Internet Explorer")

    '            Dim hButton As IntPtr = FindWindowEx(hWnd2, IntPtr.Zero, "Button", "再試行(&R)")
    '            SendMessage(hButton, BM.CLICK, 0, Nothing)
    '            Com.Sleep5(waitT)
    '            waitT = waitT + 1000

    '        End If


    '        Pub_Com.SleepAndWaitComplete(cie)
    '        Pub_Com.SleepAndWaitComplete(cie)
    '        Pub_Com.SleepAndWaitComplete(cie)


    '        CType(cie, SHDocVw.InternetExplorerMedium).ExecWB(SHDocVw.OLECMDID.OLECMDID_SAVEAS, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_PROMPTUSER)
    '        Com.Sleep5(500)
    '        Dim hWnd As IntPtr
    '        hWnd = FindWindow("#32770", "コピーを保存")
    '        If hWnd = IntPtr.Zero Then
    '            ClosePdrErr()
    '            Com.Sleep5(500)
    '            cie.Refresh()
    '            Com.Sleep5(1000)
    '            Dim hWnd2 As IntPtr
    '            hWnd2 = FindWindow("#32770", "Windows Internet Explorer")

    '            Dim hButton As IntPtr = FindWindowEx(hWnd2, IntPtr.Zero, "Button", "再試行(&R)")
    '            SendMessage(hButton, BM.CLICK, 0, Nothing)
    '            Com.Sleep5(waitT)
    '            waitT = waitT + 1000
    '            GoTo reD
    '            Pub_Com.SleepAndWaitComplete(cie)
    '            Pub_Com.SleepAndWaitComplete(cie)
    '            Pub_Com.SleepAndWaitComplete(cie)

    '        End If

    '        Com.Sleep5(500)
    '        'SavePdf(csvFileName.ToString)
    '        cie.Quit()

    '        'Gsz_Thread_Test.Abort()

    '    End Sub

    'Step 1 LOGIN IN
    Public Sub DoStep1_Login()


        Pub_Com.AddMsg("OnSiteパスワード入力")
        '＊＊＊ OnSiteパスワード入力
        Pub_Com.GetElementBy(Ie, "", "input", "name", "strPassWord").innerText = ConfigurationManager.AppSettings("OnSitePassword").ToString()
        Pub_Com.GetElementBy(Ie, "", "input", "value", "ログオン").click()

        Pub_Com.AddMsg("業務別総合メニュー")
        ''＊＊＊ 業務別総合メニュー
        Pub_Com.GetElementBy(Ie, "SubHeader", "a", "innertext", "[見積]").click()


        Pub_Com.AddMsg("物販明細")
        ''＊＊＊ 物販明細
        Pub_Com.GetElementBy(Ie, "Main", "input", "value", "物販明細").click()
        Pub_Com.SleepAndWaitComplete(Ie, 100)
        Pub_Com.SleepAndWaitComplete(Ie, 100)



    End Sub

    'Step 1 新規見積もり
    Public Sub DoStep1_PoupuSentaku(ByVal 事業所 As String, ByVal 得意先 As String, ByVal 下店 As String, ByVal 現場名 As String, ByVal 備考 As String, ByVal 日付連番 As String, ByVal fl As String)
        Dim ShellWindows As New SHDocVw.ShellWindows
        Try

            Pub_Com.AddMsg("見積検索 POPUP")
            Dim cIe As SHDocVw.InternetExplorerMedium = GetPopupWindow("OnSite", "mitSearch.asp")

            While cIe Is Nothing
                Com.Sleep5(100)
                ' Com.Sleep5(5100)
                cIe = GetPopupWindow("OnSite", "mitSearch.asp")
            End While

            Pub_Com.SleepAndWaitComplete(cIe)
            Pub_Com.GetElementBy(cIe, "", "input", "name", "strGenbaMei").innerText = 現場名
            Pub_Com.GetElementBy(cIe, "", "input", "name", "strUriJgy").click()
            Pub_Com.SleepAndWaitComplete(cIe)
            AddProBar(lv2) '12


            Pub_Com.AddMsg("事業所検索 POPUP")
            Dim cIe2 As SHDocVw.InternetExplorerMedium = GetPopupWindow("OnSite", "jgyKensaku.asp")
            While cIe2 Is Nothing
                Com.Sleep5(100)
                cIe2 = GetPopupWindow("OnSite", "jgyKensaku.asp")
            End While

            Pub_Com.SleepAndWaitComplete(cIe2)
            Pub_Com.GetElementBy(cIe2, "", "input", "name", "strJgyCd").innerText = 事業所
            Pub_Com.GetElementBy(cIe2, "", "input", "value", "検　索").click()
            AddProBar(lv2) '13
            Pub_Com.SleepAndWaitComplete(cIe)
            Pub_Com.GetElementBy(cIe, "", "input", "value", "検　索").click()

            Pub_Com.SleepAndWaitComplete(Ie)

            Pub_Com.SleepAndWaitComplete(Ie)
            AddProBar(lv2) '14

            '50件以上の場合
            Try
                Dim Doc As mshtml.HTMLDocument
                Dim eles As mshtml.IHTMLElementCollection
                Doc = CType(Ie.Document, mshtml.HTMLDocument)
                Dim fra As mshtml.HTMLWindow2 = Pub_Com.GetFrameWait(Ie, "fraHyou")
                Doc = CType(fra.document, mshtml.HTMLDocument)
                eles = Doc.getElementsByTagName("input")

                For Each ele As mshtml.IHTMLElement In eles
                    Try
                        If ele.getAttribute("value").ToString = "継 続" Then
                            ele.click()
                            Pub_Com.SleepAndWaitComplete(Ie)
                            Exit For
                        End If

                    Catch ex As Exception

                    End Try

                Next
            Catch ex As Exception

            End Try
            AddProBar(lv2) '15

            Exit Sub
        Catch ex As Exception
            'DoStep1_SinkiMitumori(事業所, 得意先, 下店, 現場名, 備考, 日付連番, fl)
            Exit Sub
        End Try
    End Sub

    'Step 2 （WHILE）
    Public Function DoStep2_Set(Optional ByVal startIdx As Integer = 0) As Boolean

        Dim Doc As mshtml.HTMLDocument
        Dim eles As mshtml.IHTMLElementCollection
        Doc = CType(Ie.Document, mshtml.HTMLDocument)
        Dim fra As mshtml.HTMLWindow2 = Pub_Com.GetFrameWait(Ie, "fraHyou")

        Doc = CType(fra.document, mshtml.HTMLDocument)
        eles = Doc.getElementsByTagName("input")

        For i As Integer = startIdx To eles.length - 1
            Dim ele As mshtml.IHTMLElement = CType(eles.item(i), mshtml.IHTMLElement)
            Try
                If ele.getAttribute("name").ToString = "strMitKbnHen" Then
                    Dim tr As mshtml.IHTMLTableRow = CType(ele.parentElement.parentElement, mshtml.IHTMLTableRow)
                    Dim td As mshtml.HTMLTableCell = CType(tr.cells.item(4), mshtml.HTMLTableCell)
                    'Return ele
                    If td.innerText = "作成中" Then
                        ele.click()
                        Pub_Com.GetElementBy(Ie, "fraHead", "input", "value", "発注納期非表示").click()
                        Pub_Com.SleepAndWaitComplete(Ie)


                        Dim Doc1 As mshtml.HTMLDocument = CType(Ie.Document, mshtml.HTMLDocument)
                        Dim fra1 As mshtml.HTMLWindow2 = Pub_Com.GetFrameWait(Ie, "fraMitBody")
                        Doc1 = CType(fra1.document, mshtml.HTMLDocument)

                        If Doc1.body.innerText.IndexOf("発注可能な見積ではありません") > 0 Then
                            Ie.GoBack()
                            Pub_Com.SleepAndWaitComplete(Ie)
                            Return DoStep2_Set(i + 1)
                        Else
                            Return True
                        End If



                    End If
                End If
            Catch ex As Exception

            End Try


        Next

        Return False

    End Function

    'POPUP Window 取得
    Public Function GetPopupWindow(ByVal titleKey As String, ByVal fileNameKey As String) As SHDocVw.InternetExplorerMedium
        Dim ShellWindows As New SHDocVw.ShellWindows
        For Each childIe As SHDocVw.InternetExplorerMedium In ShellWindows
            System.Windows.Forms.Application.DoEvents()
            Dim filename As String = System.IO.Path.GetFileNameWithoutExtension(childIe.FullName).ToLower()
            If filename = "iexplore" Then
                If CType(childIe, SHDocVw.InternetExplorerMedium).LocationURL.Contains(fileNameKey) Then
                    If CType(childIe.Document, mshtml.HTMLDocument).title = titleKey Then
                        If CType(childIe.Document, mshtml.HTMLDocument).url.Contains(fileNameKey) Then
                            Pub_Com.SleepAndWaitComplete(childIe)
                            Com.Sleep5(500)
                            Return childIe
                        End If
                    End If
                End If
            End If
        Next
        Return Nothing
    End Function

    'Public Function GetPopupWindowPDF() As SHDocVw.InternetExplorerMedium
    '    Dim ShellWindows As SHDocVw.ShellWindows = New SHDocVw.ShellWindows
    '    For Each childIe As SHDocVw.InternetExplorerMedium In (ShellWindows)
    '        System.Windows.Forms.Application.DoEvents()
    '        Dim filename As String = System.IO.Path.GetFileNameWithoutExtension(childIe.FullName).ToLower()
    '        If filename = "iexplore" Then
    '            If CType(childIe, SHDocVw.InternetExplorerMedium).LocationURL.Contains("servlet") Then
    '                Return childIe
    '            End If
    '        End If
    '    Next
    '    Return GetPopupWindowPDF()
    'End Function


    Public Function HaveWindowPDF() As Boolean
        Dim ShellWindows As New SHDocVw.ShellWindows
        For Each childIe As SHDocVw.InternetExplorerMedium In ShellWindows
            System.Windows.Forms.Application.DoEvents()
            Dim filename As String = System.IO.Path.GetFileNameWithoutExtension(childIe.FullName).ToLower()
            If filename = "iexplore" Then
                If CType(childIe, SHDocVw.InternetExplorerMedium).LocationURL.Contains("servlet") Then
                    Return True
                End If
            End If
        Next
        Return False
    End Function


    Public Sub WaitWindowPDF()


        Try
            Dim ShellWindows As New SHDocVw.ShellWindows
            For Each childIe As SHDocVw.InternetExplorerMedium In ShellWindows
                System.Windows.Forms.Application.DoEvents()
                Dim filename As String = System.IO.Path.GetFileNameWithoutExtension(Ie.FullName).ToLower()
                If filename = "iexplore" Then
                    While CType(childIe, SHDocVw.InternetExplorerMedium).LocationURL.Contains("servlet")
                        Com.Sleep5(1)
                        System.Windows.Forms.Application.DoEvents()
                    End While
                End If
            Next

        Catch ex As Exception

        End Try


    End Sub

    ''ファイル選択
    'Private Sub BackgroundWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker.DoWork
    '    Dim FileName As String = DirectCast(e.Argument, String)
    '    Dim hWnd As IntPtr
    '    Do While hWnd = IntPtr.Zero
    '        hWnd = FindWindow("#32770", "アップロードするファイルの選択")
    '        Sleep(1)
    '    Loop
    '    Dim hComboBoxEx As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "ComboBoxEx32", String.Empty)
    '    Dim hComboBox As IntPtr = FindWindowEx(hComboBoxEx, IntPtr.Zero, "ComboBox", String.Empty)
    '    Dim hEdit As IntPtr = FindWindowEx(hComboBox, IntPtr.Zero, "Edit", String.Empty)
    '    Do Until IsWindowVisible(hEdit)
    '        Sleep(1)
    '    Loop
    '    SendMessage(hEdit, WM.SETTEXT, 0, New StringBuilder(FileName))
    '    Dim hButton As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "Button", "開く(&O)")
    '    SendMessage(hButton, BM.CLICK, 0, Nothing)
    'End Sub

    'ファイル選択
    'Private Sub PDFMessageBackgroundWorkerDoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles PDFMessageBackgroundWorker.DoWork
    '    Dim FileName As String = DirectCast(e.Argument, String)
    '    PDFMessageBW(FileName)

    'End Sub

    Private Function ClosePdrErr() As Boolean
        Dim hWnd As IntPtr
        hWnd = FindWindow("#32770", "Adobe Reader")

        If hWnd = IntPtr.Zero Then
            Return True
        Else
            Dim hComboBoxEx As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "GroupBox", String.Empty)
            Dim hButton As IntPtr = FindWindowEx(hComboBoxEx, IntPtr.Zero, "Button", "OK")
            SendMessage(hButton, BM.CLICK, 0, Nothing)
            ClosePdrErr = False

        End If

        hWnd = FindWindow("#32770", "Adobe Reader")
        If hWnd <> IntPtr.Zero Then
            SendMessage(hWnd, 10, 0, Nothing)
            ClosePdrErr = False

        End If



    End Function

    'Private Sub PDFMessageBW(ByVal csvFileName As String)

    '    Dim hWnd As IntPtr
    '    Do While hWnd = IntPtr.Zero
    '        hWnd = FindWindow("#32770", "Adobe Reader")
    '        Sleep(1)
    '    Loop
    '    Dim hComboBoxEx As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "GroupBox", String.Empty)
    '    'Dim hComboBox As IntPtr = FindWindowEx(hComboBoxEx, IntPtr.Zero, "ComboBox", String.Empty)
    '    'Dim hEdit As IntPtr = FindWindowEx(hComboBoxEx, IntPtr.Zero, "Edit", String.Empty)
    '    'Do Until IsWindowVisible(hEdit)
    '    '    Sleep(1)
    '    'Loop
    '    Com.Sleep5(100)
    '    Dim hButton As IntPtr = FindWindowEx(hComboBoxEx, IntPtr.Zero, "Button", "OK")
    '    SendMessage(hButton, BM.CLICK, 0, Nothing)
    '    Com.Sleep5(100)



    '    While FindWindow("#32770", "Adobe Reader") <> IntPtr.Zero
    '        Try
    '            hWnd = FindWindow("#32770", "Adobe Reader")
    '            hComboBoxEx = FindWindowEx(hWnd, IntPtr.Zero, "GroupBox", String.Empty)
    '            hButton = FindWindowEx(hComboBoxEx, IntPtr.Zero, "Button", "OK")
    '            SendMessage(hButton, BM.CLICK, 0, Nothing)
    '            Com.Sleep5(1000)
    '            SendMessage(hWnd, 10, 0, Nothing)
    '        Catch ex As Exception

    '        End Try
    '    End While

    '    Dim childIe As SHDocVw.InternetExplorerMedium = GetPopupWindowPDF()
    '    CType(childIe, SHDocVw.InternetExplorerMedium).GoBack()


    '    'Pub_Com.SleepAndWaitComplete(childIe, 100)
    '    'Pub_Com.SleepAndWaitComplete(childIe, 100)

    '    Com.Sleep5(6000)
    '    If HaveWindowPDF() Then
    '        System.Windows.Forms.Application.DoEvents()
    '        Com.Sleep5(1000)
    '        SavePdf(csvFileName)
    '    End If




    '    'Dim ShellWindows As New SHDocVw.ShellWindows
    '    'For Each childIe As SHDocVw.InternetExplorerMedium In ShellWindows
    '    '    System.Windows.Forms.Application.DoEvents()
    '    '    Dim filename As String = System.IO.Path.GetFileNameWithoutExtension(Ie.FullName).ToLower()
    '    '    If filename = "iexplore" Then
    '    '        If CType(childIe, SHDocVw.InternetExplorerMedium).LocationURL.Contains("servlet") Then
    '    '            CType(childIe, SHDocVw.InternetExplorerMedium).GoBack()

    '    '            Com.Sleep5(200)
    '    '            Do Until childIe.ReadyState = WebBrowserReadyState.Complete AndAlso Not childIe.Busy
    '    '                System.Windows.Forms.Application.DoEvents()
    '    '                Com.Sleep5(1)
    '    '            Loop
    '    '            Com.Sleep5(1000)
    '    '            System.Windows.Forms.Application.DoEvents()
    '    '        End If
    '    '    End If
    '    'Next

    '    PDFMessageBW(csvFileName)
    'End Sub

    'Private Sub PDFBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles PDFBackgroundWorker.DoWork
    '    'Dim FileName As String = DirectCast(e.Argument, String)

    '    'Save PDF
    '    Dim cIe2 As SHDocVw.InternetExplorerMedium = GetPopupWindowPDF()
    '    Pub_Com.SleepAndWaitComplete(cIe2, 100)
    '    Pub_Com.SleepAndWaitComplete(cIe2, 100)
    '    Pub_Com.SleepAndWaitComplete(cIe2, 100)
    '    Pub_Com.SleepAndWaitComplete(cIe2, 100)

    '    Com.Sleep5(500)


    '    CType(cIe2, SHDocVw.InternetExplorerMedium).ExecWB(SHDocVw.OLECMDID.OLECMDID_SAVEAS, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_PROMPTUSER)
    '    Com.Sleep5(500)
    '    'SavePdf(FileName)
    '    Com.Sleep5(500)
    '    cIe2.Quit()
    'End Sub


    'Private Sub SavePdf(ByVal csvFileName As String)

    '    Try
    '        Dim pdfPath As String = ConfigurationManager.AppSettings("Pdf_Path").ToString()

    '        If System.IO.File.Exists(pdfPath & csvFileName & ".pdf") Then
    '            FileSystem.Rename(pdfPath & csvFileName & ".pdf", pdfPath & csvFileName & ".pdf" & ".bk." & Now.ToString("yyyyMMddHHmmss"))
    '        End If

    '        Dim hWnd As IntPtr
    '        Do While hWnd = IntPtr.Zero
    '            hWnd = FindWindow("#32770", "コピーを保存")
    '            Sleep(1)
    '        Loop

    '        Dim filePathComboBoxEx32 As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "ComboBoxEx32", String.Empty)
    '        Dim filePathComboBox As IntPtr = FindWindowEx(filePathComboBoxEx32, IntPtr.Zero, "ComboBox", String.Empty)
    '        Dim hEditfilePath As IntPtr = FindWindowEx(filePathComboBox, IntPtr.Zero, "Edit", String.Empty)
    '        SendMessage(hEditfilePath, WM.SETTEXT, 0, New StringBuilder(pdfPath & csvFileName & ".pdf"))

    '        Dim saveButton As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "Button", "保存")
    '        SendMessage(saveButton, BM.CLICK, 0, Nothing)
    '        Com.Sleep5(1000)
    '    Catch ex As Exception

    '    End Try

    '    If FindWindow("#32770", "コピーを保存") <> IntPtr.Zero Then
    '        SavePdf(csvFileName)
    '    End If

    'End Sub




#End Region


    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
        Timer1.Stop()
    End Sub


    Public Sub StartTimer()
        Dim tcb As New Threading.TimerCallback(AddressOf Me.TimerMethod)
        Dim objTimer As System.Threading.Timer
        objTimer = New System.Threading.Timer(tcb, Nothing, TimeSpan.FromSeconds(5), TimeSpan.FromSeconds(10))
    End Sub
    Public Sub TimerMethod(ByVal state As Object)
        'MsgBox("The Timer invoked this method.")

        Do While Timer1.Enabled
            Application.DoEvents()

        Loop

    End Sub

    Private Sub AutoImportNouhinsyo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Com.Sleep5(3100)
    End Sub
End Class