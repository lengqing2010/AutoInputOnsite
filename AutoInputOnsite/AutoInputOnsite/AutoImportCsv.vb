Option Explicit On
Option Strict On

Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading.Thread
Imports System.Configuration

Public Class AutoImportCsv

    Public _ProBar As Double = 0
    Public lv1 As Double
    Public lv2 As Double
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

    Public Sub AddProBar(ByVal x As Double)
        _ProBar += x
    End Sub
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


    'Private WithEvents NinnsyouBackgroundWorker As BackgroundWorker
    'Private Sub NinnsyouBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles NinnsyouBackgroundWorker.DoWork
    '    Ninnsyou()
    'End Sub

    ''' <summary>
    ''' �F�؏��
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Ninnsyou()

        Dim user As String = ConfigurationManager.AppSettings("User").ToString()
        Dim password As String = ConfigurationManager.AppSettings("Password").ToString()

        If user = "" Then user = "china\shil2"
        If password = "" Then password = "asdf@123"

        Dim idx As Integer = 0

        Dim hWnd As IntPtr
        Do While hWnd = IntPtr.Zero
            hWnd = FindWindow("#32770", "windows �Z�L�����e�B")
            Sleep(10)
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

    Public Ie As SHDocVw.InternetExplorerMedium
    Public Pub_Com As Com




    'Public Sub Com.Sleep5(ByVal Interval As Double)
    '    'Dim __time As DateTime = DateTime.Now
    '    'Dim __Span As Int64 = Interval * 10000
    '    'While (DateTime.Now.Ticks - __time.Ticks < __Span)
    '    '    Application.DoEvents()
    '    'End While
    '    'Timer1.Interval = Interval
    '    'Timer1.Enabled = True
    '    'Timer1.Start()

    '    Dim Start As Long
    '    Start = timeGetTime
    '    Do While (timeGetTime < Start + CLng(Interval))
    '        Windows.Forms.Application.DoEvents()
    '        Sleep(1)
    '    Loop

    'End Sub
    '���s
    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
        DoAll()
    End Sub

    '���s MAIN
    Public Sub DoAll()

        Pub_Com = New Com("���o���׍쐬" & Now.ToString("yyyyMMddHHmmss"))
        If Pub_Com.file_list_hattyuu.Count = 0 Then
            ProBar = 100
            Exit Sub
        End If

        lv1 = 90 / Pub_Com.file_list_hattyuu.Count
        Ie = New SHDocVw.InternetExplorerMedium

        'NinnsyouBackgroundWorker = New BackgroundWorker
        BackgroundWorker = New BackgroundWorker

        Dim authHeader As Object = "Authorization: Basic " + _
        Convert.ToBase64String(System.Text.UnicodeEncoding.UTF8.GetBytes(String.Format("{0}:{1}", Pub_Com.user, Pub_Com.password))) + "\r\n"


        Dim thr As New Threading.Thread(AddressOf Ninnsyou)
        thr.Start()

        'Pub_Com.StartWK(NinnsyouBackgroundWorker)
        '������ OnSite�p�X���[�h���͉��
        Ie.Navigate(Pub_Com.url, , , , authHeader)
        Ie.Silent = True
        Ie.Visible = True



        'Pub_Com.StopWK(NinnsyouBackgroundWorker)

        'Me.NinnsyouBackgroundWorker.RunWorkerAsync()


        ProBar = 5
        DoStep1_Login() '���������O�C��
        ProBar = 10


        'Dim idx As Integer = 1

        For Each fl As String In Pub_Com.file_list_hattyuu

            lv2 = lv1 / 10
            Pub_Com.AddMsg("�捞�F" & fl)

            Dim ���Ə�, ���Ӑ�, ���X, ���ꖼ, ���l, ���t�A�� As String

            ���Ə� = fl.Split("-"c)(0)
            ���Ӑ� = fl.Split("-"c)(1)
            ���X = fl.Split("-"c)(2)
            ���ꖼ = fl.Split("-"c)(3)
            ���l = fl.Split("-"c)(4)
            ���t�A�� = fl.Split("-"c)(5)


            'While
            DoStep2_Sinki(���Ə�, ���Ӑ�, ���X, ���ꖼ, ���l, ���t�A��, fl)


            'ProBar = CInt(20 + CInt(80 / Pub_Com.file_list_hattyuu.Count) * idx * 0.75 - 1)

            Pub_Com.AddMsg("�ړ�CSV�F" & fl & "��" & Pub_Com.folder_Hattyuu_kanryou)
            If System.IO.File.Exists(Pub_Com.folder_Hattyuu_kanryou & fl) Then
                FileSystem.Rename(Pub_Com.folder_Hattyuu_kanryou & fl, Pub_Com.folder_Hattyuu_kanryou & fl & ".bk." & Now.ToString("yyyyMMddHHmmss"))
            End If
            System.IO.File.Move(Pub_Com.folder_Hattyuu & fl, Pub_Com.folder_Hattyuu_kanryou & fl)
            AddProBar(lv2) '10

            'idx += 1

        Next

        thr.Abort()

        ProBar = 100

        ' NinnsyouBackgroundWorker.Dispose()
        BackgroundWorker.Dispose()
        ' NinnsyouBackgroundWorker = Nothing
        BackgroundWorker = Nothing

        Ie.Quit()

    End Sub

    'Step 1 LOGIN IN
    Public Sub DoStep1_Login()
        ' Pub_Com.StartWK(NinnsyouBackgroundWorker)

        Pub_Com.AddMsg("OnSite�p�X���[�h����")
        '������ OnSite�p�X���[�h����
        Pub_Com.GetElementBy(Ie, "", "input", "name", "strPassWord").innerText = ConfigurationManager.AppSettings("OnSitePassword").ToString()
        Pub_Com.GetElementBy(Ie, "", "input", "value", "���O�I��").click()
        ' Pub_Com.StopWK(NinnsyouBackgroundWorker)

        Pub_Com.SleepAndWaitComplete(Ie)
        ' Pub_Com.StartWK(NinnsyouBackgroundWorker)
        '
        Pub_Com.AddMsg("�Ɩ��ʑ������j���[")
        ''������ �Ɩ��ʑ������j���[
        Pub_Com.GetElementBy(Ie, "SubHeader", "a", "innertext", "[����]").click()
        Pub_Com.SleepAndWaitComplete(Ie)
        'Pub_Com.StopWK(NinnsyouBackgroundWorker)

        'Pub_Com.StartWK(NinnsyouBackgroundWorker)

        Pub_Com.AddMsg("���̖���")


        ''������ ���̖���
        Pub_Com.GetElementBy(Ie, "Main", "input", "value", "���̖���").click()
        Pub_Com.SleepAndWaitComplete(Ie)
        ' Com.Sleep5(1500)
        'Pub_Com.StopWK(NinnsyouBackgroundWorker)

        '�V�K���ς���
        Pub_Com.AddMsg("�V�K����")
        DoStep1_SinkiMitumori()

    End Sub

    'Step 1 �V�K���ς���
    Public Sub DoStep1_SinkiMitumori()
        Dim ShellWindows As New SHDocVw.ShellWindows
        'Try

        'Pub_Com.StartWK(NinnsyouBackgroundWorker)

        Dim cIe As SHDocVw.InternetExplorerMedium = GetPopupWindow("OnSite", "mitSearch.asp")
        While cIe Is Nothing
            Com.Sleep5(1000)
            cIe = GetPopupWindow("OnSite", "mitSearch.asp")
        End While

        Pub_Com.GetElementBy(cIe, "", "input", "value", "�V�K����").click()
        Com.Sleep5(100)



        Pub_Com.SleepAndWaitComplete(Ie)
        'Pub_Com.StopWK(NinnsyouBackgroundWorker)

        Exit Sub
        'Catch ex As Exception
        '    DoStep1_SinkiMitumori()
        '    Exit Sub
        'End Try
    End Sub

    'Step 2 �iWHILE�j
    Public Sub DoStep2_Sinki(ByVal ���Ə� As String, ByVal ���Ӑ� As String, ByVal ���X As String, ByVal ���ꖼ As String, ByVal ���l As String, ByVal ���t�A�� As String, ByVal fl As String)

        Dim folder_Hattyuu As String = ConfigurationManager.AppSettings("Folder_Hattyuu").ToString()

        Com.Sleep5(500)
        Pub_Com.SleepAndWaitComplete(Ie)

        Pub_Com.AddMsg("    ���Ə��F" & ���Ə�)
        Pub_Com.AddMsg("    ���Ӑ�F" & ���Ӑ�)
        Pub_Com.AddMsg("    ���X�F" & ���X)
        Pub_Com.AddMsg("    ���l�F" & ���l)
        Pub_Com.AddMsg("    ���ꖼ�F" & ���ꖼ)

        Pub_Com.SleepAndWaitComplete(Ie)
        Dim fra As mshtml.HTMLWindow2 = Pub_Com.GetFrameWait(Ie, "fraMitBody")

        Pub_Com.GetElement(fra, "input", "name", "strJgyCdText").innerText = ���Ə�
        Pub_Com.GetElement(fra, "input", "name", "strTokMeiText").innerText = ���Ӑ�
        Pub_Com.GetElement(fra, "input", "name", "strOtdMeiText").innerText = ���X
        Pub_Com.GetElement(fra, "input", "name", "strBikouMei").innerText = ���l
        Pub_Com.GetElement(fra, "input", "name", "strGenbaMei").innerText = ���ꖼ
        Pub_Com.GetElement(fra, "select", "name", "aryKijyunSyouhinBunrui").setAttribute("value", "A0001,�T�b�V,L90000")

        Pub_Com.AddMsg("    �[�ޓX�Ȃ��œ�����͂� CLICK")

        Pub_Com.GetElement(fra, "input", "name", "btnUtiwake").click()

        AddProBar(lv2) '1



        ''���ό��o����
        'Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "name", "strJgyCdText").innerText = ���Ə�
        'Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "name", "strTokMeiText").innerText = ���Ӑ�
        'Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "name", "strOtdMeiText").innerText = ���X

        'Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "name", "strBikouMei").innerText = ���l
        'Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "name", "strGenbaMei").innerText = ���ꖼ

        'Pub_Com.GetElementBy(Ie, "fraMitBody", "select", "name", "aryKijyunSyouhinBunrui").setAttribute("value", "A0001,�T�b�V,L90000")
        'Pub_Com.AddMsg("    �[�ޓX�Ȃ��œ�����͂� CLICK")
        'Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "name", "btnUtiwake").click()


        Com.Sleep5(500)
        Pub_Com.SleepAndWaitComplete(Ie)


        '���ϓ������
        Try
            Dim ele1 As mshtml.IHTMLElement = Pub_Com.GetElementBy(Ie, "fraMitBody", "DIV", "classname", "ttl")
            If ele1.innerText <> "���ϓ������" Then
                DoStep2_Sinki(���Ə�, ���Ӑ�, ���X, ���ꖼ, ���l, ���t�A��, fl)
                Exit Sub
            End If
            Pub_Com.AddMsg("    ���ϓ������ CSV�捞 CLICK")
            Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "CSV�捞").click()
        Catch ex As Exception
            DoStep2_Sinki(���Ə�, ���Ӑ�, ���X, ���ꖼ, ���l, ���t�A��, fl)
            Exit Sub
        End Try
        AddProBar(lv2) '2

        Pub_Com.SleepAndWaitComplete(Ie)




        Me.BackgroundWorker.RunWorkerAsync(folder_Hattyuu & fl)
        Com.Sleep5(500)
        'Dim csvPopup As SHDocVw.InternetExplorerMedium = GetPopupWindow("OnSite", "fileYomikomiSiji.asp")



        Pub_Com.AddMsg("    ���ϓ������ CSV�捞 �Q�@�� CLICK")

        Dim fra1 As SHDocVw.InternetExplorerMedium = GetPopupWindow("OnSite", "fileYomikomiSiji.asp")
        Pub_Com.GetElementBy(fra1, "", "input", "value", "�Q�@��").click()

        Com.Sleep5(1500)
        Try
            While Pub_Com.GetElementBy(fra1, "", "input", "name", "strFilename").getAttribute("value").ToString = ""
                Com.Sleep5(1)
            End While
        Catch ex As Exception
        End Try

        Com.Sleep5(500)
        AddProBar(lv2) '3

        Pub_Com.AddMsg("    ���ϓ������ CSV�捞 ��@�� CLICK")
        Pub_Com.GetElementBy(GetPopupWindow("OnSite", "fileYomikomiSiji.asp"), "", "input", "value", "��@��").click()
        Com.Sleep5(1000)
        AddProBar(lv2) '4

        Pub_Com.SleepAndWaitComplete(Ie)
        Pub_Com.AddMsg("    ���i�R�[�h�������� ���@�� CLICK")
        Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "���@��").click()
        Pub_Com.SleepAndWaitComplete(Ie)
        Pub_Com.AddMsg("    ���i�R�[�h�������� ���@�� CLICK")
        Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "���@��").click()
        Pub_Com.SleepAndWaitComplete(Ie)
        Pub_Com.AddMsg("    ���i�R�[�h�������� ���@�� CLICK")
        Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "���@��").click()
        AddProBar(lv2) '5
        Pub_Com.SleepAndWaitComplete(Ie)
        Pub_Com.SleepAndWaitComplete(Ie)
        Pub_Com.SleepAndWaitComplete(Ie)

        '���@����
        Pub_Com.SleepAndWaitComplete(Ie)
        Pub_Com.AddMsg("    ���@���� ���@�� CLICK")
        Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "���@��").click()
        Pub_Com.SleepAndWaitComplete(Ie)
        Pub_Com.SleepAndWaitComplete(Ie)
        Pub_Com.AddMsg("    ���@���� ���@�� CLICK")
        Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "���@��").click()
        Pub_Com.SleepAndWaitComplete(Ie)
        Pub_Com.SleepAndWaitComplete(Ie)
        Pub_Com.AddMsg("    ���@���� ���@�� CLICK")
        Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "���@��").click()
        Pub_Com.SleepAndWaitComplete(Ie)
        Pub_Com.SleepAndWaitComplete(Ie)
        AddProBar(lv2) '6
        Pub_Com.AddMsg("    �P������ ���ϓ�����͂� CLICK")
        Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "���ϓ�����͂�").click()
        Pub_Com.SleepAndWaitComplete(Ie)
        AddProBar(lv2) '7
        Dim ele As mshtml.IHTMLElement = Pub_Com.GetElementBy(Ie, "fraMitBody", "DIV", "classname", "ttl")
        Dim kekka As String = ele.innerText
        AddProBar(lv2) '8
        Pub_Com.AddMsg("    �V�K���� CLICK")
        Pub_Com.GetElementBy(Ie, "fraMitMenu", "a", "innertext", "[�V�K����]").click()
        Pub_Com.SleepAndWaitComplete(Ie)
        Pub_Com.SleepAndWaitComplete(Ie)
        AddProBar(lv2) '9


    End Sub

    'POPUP Window �擾
    Public Function GetPopupWindow(ByVal titleKey As String, ByVal fileNameKey As String) As SHDocVw.InternetExplorerMedium
        Dim ShellWindows As New SHDocVw.ShellWindows
        For Each childIe As SHDocVw.InternetExplorerMedium In ShellWindows
            Dim filename As String = System.IO.Path.GetFileNameWithoutExtension(Ie.FullName).ToLower()
            If filename = "iexplore" Then
                If CType(childIe, SHDocVw.InternetExplorerMedium).LocationURL.Contains(fileNameKey) Then

                    If CType(childIe.Document, mshtml.HTMLDocument).title.Contains("���i��񂪖���") Then
                        MsgBox("���i��񂪖���")


                        Dim thrs() As Process = Process.GetProcessesByName("AutoInputOnsite")
                        For Each tr As Process In thrs
                            tr.Kill()

                        Next
                        Application.Exit()
                    End If
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

#End Region

End Class