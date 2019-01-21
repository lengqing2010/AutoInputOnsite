Option Explicit On
Option Strict On

Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading.Thread
Imports System.Configuration

Public Class AutoImportNouhinsyo

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

#Region "����"
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

#Region "�F�؏��"



    Private Sub NinnsyouBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles NinnsyouBackgroundWorker.DoWork
        Ninnsyou()
    End Sub

    ''' <summary>
    ''' �F�؏��
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Ninnsyou()
        Dim user As String = ConfigurationManager.AppSettings("User").ToString()
        Dim password As String = ConfigurationManager.AppSettings("Password").ToString()
        Dim hWnd As IntPtr
        Do While hWnd = IntPtr.Zero
            hWnd = FindWindow("#32770", "windows �Z�L�����e�B")
            Sleep(1)
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

        Ninnsyou()
    End Sub

#End Region

#Region "�������s"

    Private WithEvents BackgroundWorker As BackgroundWorker
    Private WithEvents NinnsyouBackgroundWorker As BackgroundWorker
    Dim Ie As New SHDocVw.InternetExplorerMedium

    Private Sub AutoImportCsv_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
        DoAll()
    End Sub

    Public logFileName As String

    'MESSAGE �ǉ�
    Public Sub AddMsg(ByVal msg As String)
        Me.RichTextBox1.Text = Now.ToString("yy/MM/dd HH:mm:ss") & "��:" & msg & vbNewLine & Me.RichTextBox1.Text

        WriteLog(Now.ToString("yy/MM/dd HH:mm:ss") & "��:" & msg, logFileName)
    End Sub
    '���s MAIN
    Public Sub DoAll()

        logFileName = Now.ToString("yyyyMMddHHmmss")

        '�t�H���_�ǂ�
        Dim folder_Hattyuu As String = ConfigurationManager.AppSettings("Folder_Hattyuu").ToString()
        Dim folder_Hattyuu_kanryou As String = ConfigurationManager.AppSettings("Folder_Hattyuu_kanryou").ToString()
        'Dim folder_Nouki As String = ConfigurationManager.AppSettings("Folder_Nouki").ToString()
        If Not System.IO.Directory.Exists(folder_Hattyuu) Then
            My.Computer.FileSystem.CreateDirectory(folder_Hattyuu)
        End If

        If Not System.IO.Directory.Exists(folder_Hattyuu_kanryou) Then
            My.Computer.FileSystem.CreateDirectory(folder_Hattyuu_kanryou)
        End If

        'If Not System.IO.Directory.Exists(folder_Nouki) Then
        '    My.Computer.FileSystem.CreateDirectory(folder_Nouki)
        'End If


        NinnsyouBackgroundWorker = New BackgroundWorker
        BackgroundWorker = New BackgroundWorker

        Me.RichTextBox1.Text = ""

        Dim user As String = ConfigurationManager.AppSettings("User").ToString()
        Dim password As String = ConfigurationManager.AppSettings("Password").ToString()

        If user = "" Then
            user = "china\shil2"
        End If
        If password = "" Then
            password = "asdf@123"
        End If


        Dim authHeader As Object = "Authorization: Basic " + _
        Convert.ToBase64String(System.Text.UnicodeEncoding.UTF8.GetBytes(String.Format("{0}:{1}", user, password))) + "\r\n"


        '������ OnSite�p�X���[�h���͉��
        Dim url As String = ConfigurationManager.AppSettings("PasswordNyuuryoku").ToString()
        Ie.Navigate(url, , , , authHeader)
        Ie.Silent = True
        Ie.Visible = True

        Me.NinnsyouBackgroundWorker.RunWorkerAsync()

        'Sigle
        DoStep1_Login()


        Dim files As List(Of String) = GetAllFiles(folder_Hattyuu, "*.csv")

        For Each fl As String In files

            AddMsg("�捞�F" & fl)

            Dim ���Ə�, ���Ӑ�, ���X, ���ꖼ, ���l, ���t�A�� As String

            ���Ə� = fl.Split("-"c)(0)
            ���Ӑ� = fl.Split("-"c)(1)
            ���X = fl.Split("-"c)(2)
            ���ꖼ = fl.Split("-"c)(3)
            ���l = fl.Split("-"c)(4)
            ���t�A�� = fl.Split("-"c)(5)

            'While
            DoStep2_Sinki(���Ə�, ���Ӑ�, ���X, ���ꖼ, ���l, ���t�A��, fl)

            AddMsg("�ړ�CSV�F" & fl & "��" & folder_Hattyuu_kanryou)
            If System.IO.File.Exists(folder_Hattyuu_kanryou & fl) Then

                FileSystem.Rename(folder_Hattyuu_kanryou & fl, folder_Hattyuu_kanryou & fl & ".bk." & Now.ToString("yyyyMMddHHmmss"))
            End If

            System.IO.File.Move(folder_Hattyuu & fl, folder_Hattyuu_kanryou & fl)

        Next


        NinnsyouBackgroundWorker.Dispose()
        BackgroundWorker.Dispose()

        Ie.Quit()

        MsgBox("����")


    End Sub
    'Step 1 LOGIN IN
    Public Sub DoStep1_Login()

        AddMsg("OnSite�p�X���[�h����")
        '������ OnSite�p�X���[�h����
        GetElementBy(Ie, "", "input", "name", "strPassWord").innerText = ConfigurationManager.AppSettings("OnSitePassword").ToString()
        GetElementBy(Ie, "", "input", "value", "���O�I��").click()

        AddMsg("�Ɩ��ʑ������j���[")
        ''������ �Ɩ��ʑ������j���[
        GetElementBy(Ie, "SubHeader", "a", "innertext", "[����]").click()
        ' WaitComplete(Ie)

        AddMsg("���̖���")
        ''������ ���̖���
        GetElementBy(Ie, "Main", "input", "value", "���̖���").click()
        'WaitComplete(Ie)
        ' System.Threading.Thread.Sleep(1500)


        '�V�K���ς���
        AddMsg("�V�K����")
        DoStep1_SinkiMitumori()

    End Sub
    'Step 1 �V�K���ς���
    Public Sub DoStep1_SinkiMitumori()
        Dim ShellWindows As New SHDocVw.ShellWindows
        Try
            Dim cIe As SHDocVw.InternetExplorerMedium = GetPopupWindow("OnSite", "mitSearch.asp")
            GetElementBy(cIe, "", "input", "value", "�V�K����").click()
            System.Threading.Thread.Sleep(100)
            WaitComplete(Ie)
            Exit Sub
        Catch ex As Exception
            DoStep1_SinkiMitumori()
            Exit Sub
        End Try
    End Sub
    'Step 2 �iWHILE�j
    Public Sub DoStep2_Sinki(ByVal ���Ə� As String, ByVal ���Ӑ� As String, ByVal ���X As String, ByVal ���ꖼ As String, ByVal ���l As String, ByVal ���t�A�� As String, ByVal fl As String)

        Dim folder_Hattyuu As String = ConfigurationManager.AppSettings("Folder_Hattyuu").ToString()

        System.Threading.Thread.Sleep(500)
        WaitComplete(Ie)

        AddMsg("    ���Ə��F" & ���Ə�)
        AddMsg("    ���Ӑ�F" & ���Ӑ�)
        AddMsg("    ���X�F" & ���X)
        AddMsg("    ���l�F" & ���l)
        AddMsg("    ���ꖼ�F" & ���ꖼ)

        '���ό��o����
        GetElementBy(Ie, "fraMitBody", "input", "name", "strJgyCdText").innerText = ���Ə�
        GetElementBy(Ie, "fraMitBody", "input", "name", "strTokMeiText").innerText = ���Ӑ�
        GetElementBy(Ie, "fraMitBody", "input", "name", "strOtdMeiText").innerText = ���X

        GetElementBy(Ie, "fraMitBody", "input", "name", "strBikouMei").innerText = ���l
        GetElementBy(Ie, "fraMitBody", "input", "name", "strGenbaMei").innerText = ���ꖼ

        GetElementBy(Ie, "fraMitBody", "select", "name", "aryKijyunSyouhinBunrui").setAttribute("value", "A0001,�T�b�V,L90000")
        AddMsg("    �[�ޓX�Ȃ��œ�����͂� CLICK")
        GetElementBy(Ie, "fraMitBody", "input", "name", "btnUtiwake").click()
        System.Threading.Thread.Sleep(500)
        WaitComplete(Ie)


        '���ϓ������
        Try
            Dim ele As mshtml.IHTMLElement = GetElementBy(Ie, "fraMitBody", "DIV", "classname", "ttl")
            If ele.innerText <> "���ϓ������" Then
                DoStep2_Sinki(���Ə�, ���Ӑ�, ���X, ���ꖼ, ���l, ���t�A��, fl)
                Exit Sub
            End If
            AddMsg("    ���ϓ������ CSV�捞 CLICK")
            GetElementBy(Ie, "fraMitBody", "input", "value", "CSV�捞").click()
        Catch ex As Exception
            DoStep2_Sinki(���Ə�, ���Ӑ�, ���X, ���ꖼ, ���l, ���t�A��, fl)
            Exit Sub
        End Try


        WaitComplete(Ie)


        Try

            Me.BackgroundWorker.RunWorkerAsync(folder_Hattyuu & fl)
            System.Threading.Thread.Sleep(500)

            'Dim csvPopup As SHDocVw.InternetExplorerMedium = GetPopupWindow("OnSite", "fileYomikomiSiji.asp")

            AddMsg("    ���ϓ������ CSV�捞 �Q�@�� CLICK")
            GetElementBy(GetPopupWindow("OnSite", "fileYomikomiSiji.asp"), "", "input", "value", "�Q�@��").click()
            System.Threading.Thread.Sleep(1000)
            AddMsg("    ���ϓ������ CSV�捞 ��@�� CLICK")
            GetElementBy(GetPopupWindow("OnSite", "fileYomikomiSiji.asp"), "", "input", "value", "��@��").click()
            System.Threading.Thread.Sleep(100)

            WaitComplete(Ie)
            AddMsg("    ���i�R�[�h�������� ���@�� CLICK")
            GetElementBy(Ie, "fraMitBody", "input", "value", "���@��").click()
            WaitComplete(Ie)
            AddMsg("    ���i�R�[�h�������� ���@�� CLICK")
            GetElementBy(Ie, "fraMitBody", "input", "value", "���@��").click()
            WaitComplete(Ie)
            AddMsg("    ���i�R�[�h�������� ���@�� CLICK")
            GetElementBy(Ie, "fraMitBody", "input", "value", "���@��").click()

            WaitComplete(Ie)
            WaitComplete(Ie)
            WaitComplete(Ie)

            '���@����
            WaitComplete(Ie)
            AddMsg("    ���@���� ���@�� CLICK")
            GetElementBy(Ie, "fraMitBody", "input", "value", "���@��").click()
            WaitComplete(Ie)
            WaitComplete(Ie)
            AddMsg("    ���@���� ���@�� CLICK")
            GetElementBy(Ie, "fraMitBody", "input", "value", "���@��").click()
            WaitComplete(Ie)
            WaitComplete(Ie)
            AddMsg("    ���@���� ���@�� CLICK")
            GetElementBy(Ie, "fraMitBody", "input", "value", "���@��").click()
            WaitComplete(Ie)
            WaitComplete(Ie)
            AddMsg("    �P������ ���ϓ�����͂� CLICK")
            GetElementBy(Ie, "fraMitBody", "input", "value", "���ϓ�����͂�").click()
            WaitComplete(Ie)
            Dim ele As mshtml.IHTMLElement = GetElementBy(Ie, "fraMitBody", "DIV", "classname", "ttl")
            Dim kekka As String = ele.innerText
            AddMsg("    �V�K���� CLICK")
            GetElementBy(Ie, "fraMitMenu", "a", "innertext", "[�V�K����]").click()
            WaitComplete(Ie)
            WaitComplete(Ie)

        Catch ex As Exception
            DoStep2_Sinki(���Ə�, ���Ӑ�, ���X, ���ꖼ, ���l, ���t�A��, fl)
            Exit Sub
        End Try

    End Sub
    'POPUP Window �擾
    Public Function GetPopupWindow(ByVal titleKey As String, ByVal fileNameKey As String) As SHDocVw.InternetExplorerMedium
        Dim ShellWindows As New SHDocVw.ShellWindows
        For Each childIe As SHDocVw.InternetExplorerMedium In ShellWindows
            Dim filename As String = System.IO.Path.GetFileNameWithoutExtension(Ie.FullName).ToLower()
            If filename = "iexplore" Then
                If CType(childIe, SHDocVw.InternetExplorerMedium).LocationURL.Contains(fileNameKey) Then
                    If CType(childIe.Document, mshtml.HTMLDocument).title = titleKey Then
                        If CType(childIe.Document, mshtml.HTMLDocument).url.Contains(fileNameKey) Then
                            WaitComplete(childIe)
                            System.Threading.Thread.Sleep(500)
                            Return childIe
                        End If
                    End If
                End If
            End If
        Next
        Return Nothing
    End Function

    '�t�@�C���I��
    Private Sub BackgroundWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker.DoWork
        Dim FileName As String = DirectCast(e.Argument, String)
        Dim hWnd As IntPtr
        Do While hWnd = IntPtr.Zero
            hWnd = FindWindow("#32770", "�A�b�v���[�h����t�@�C���̑I��")
            Sleep(1)
        Loop
        Dim hComboBoxEx As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "ComboBoxEx32", String.Empty)
        Dim hComboBox As IntPtr = FindWindowEx(hComboBoxEx, IntPtr.Zero, "ComboBox", String.Empty)
        Dim hEdit As IntPtr = FindWindowEx(hComboBox, IntPtr.Zero, "Edit", String.Empty)
        Do Until IsWindowVisible(hEdit)
            Sleep(1)
        Loop
        SendMessage(hEdit, WM.SETTEXT, 0, New StringBuilder(FileName))
        Dim hButton As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "Button", "�J��(&O)")
        SendMessage(hButton, BM.CLICK, 0, Nothing)
    End Sub


    ''' <summary>
    ''' Wait For Complete
    ''' </summary>
    ''' <param name="webApp"></param>
    ''' <remarks></remarks>
    Sub WaitComplete(ByRef webApp As SHDocVw.InternetExplorerMedium)
        System.Threading.Thread.Sleep(100)
        For i As Integer = 0 To 10
            Do Until webApp.ReadyState = WebBrowserReadyState.Complete AndAlso Not webApp.Busy
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(10)
            Loop
            System.Threading.Thread.Sleep(10)
        Next

    End Sub

    'Get Element
    Public Function GetElementBy(ByRef webApp As SHDocVw.InternetExplorerMedium, ByVal fraName As String, ByVal tagName As String, ByVal keyName As String, ByVal keyTxt As String) As mshtml.IHTMLElement
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

    'Get element and do soming
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

    'Get Frame
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

    '�t�@�C���擾
    Private Function GetAllFiles(ByVal strDirect As String, ByVal exName As String) As List(Of String)

        Dim flList As New List(Of String)

        If Not (strDirect Is Nothing) Then
            Dim mFileInfo As System.IO.FileInfo
            Dim mDir As System.IO.DirectoryInfo
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

    '���O�o��
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

#End Region

End Class