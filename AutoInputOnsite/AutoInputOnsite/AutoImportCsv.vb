Option Explicit On
Option Strict On

Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading.Thread
Imports System.Configuration

Public Class AutoImportCsv

    Public ProBar As Integer = 1

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

        If user = "" Then
            user = "china\shil2"
        End If
        If password = "" Then
            password = "asdf@123"
        End If

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
    Public Ie As New SHDocVw.InternetExplorerMedium
    Public Pub_Com As Com

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



        NinnsyouBackgroundWorker = New BackgroundWorker
        BackgroundWorker = New BackgroundWorker

        Dim authHeader As Object = "Authorization: Basic " + _
        Convert.ToBase64String(System.Text.UnicodeEncoding.UTF8.GetBytes(String.Format("{0}:{1}", Pub_Com.user, Pub_Com.password))) + "\r\n"

        '������ OnSite�p�X���[�h���͉��
        Ie.Navigate(Pub_Com.url, , , , authHeader)
        Ie.Silent = True
        Ie.Visible = True

        Me.NinnsyouBackgroundWorker.RunWorkerAsync()

        '���������O�C��
        DoStep1_Login()


        Dim idx As Integer = 1

        For Each fl As String In Pub_Com.file_list_hattyuu

            Pub_Com.AddMsg("�捞�F" & fl)

            Dim ���Ə�, ���Ӑ�, ���X, ���ꖼ, ���l, ���t�A�� As String

            ���Ə� = fl.Split("-"c)(0)
            ���Ӑ� = fl.Split("-"c)(1)
            ���X = fl.Split("-"c)(2)
            ���ꖼ = fl.Split("-"c)(3)
            ���l = fl.Split("-"c)(4)
            ���t�A�� = fl.Split("-"c)(5)

            ProBar = CInt(20 + CInt(80 / Pub_Com.file_list_hattyuu.Count) * idx * 0.25 - 1)

            'While
            DoStep2_Sinki(���Ə�, ���Ӑ�, ���X, ���ꖼ, ���l, ���t�A��, fl)


            ProBar = CInt(20 + CInt(80 / Pub_Com.file_list_hattyuu.Count) * idx * 0.75 - 1)

            Pub_Com.AddMsg("�ړ�CSV�F" & fl & "��" & Pub_Com.folder_Hattyuu_kanryou)
            If System.IO.File.Exists(Pub_Com.folder_Hattyuu_kanryou & fl) Then

                FileSystem.Rename(Pub_Com.folder_Hattyuu_kanryou & fl, Pub_Com.folder_Hattyuu_kanryou & fl & ".bk." & Now.ToString("yyyyMMddHHmmss"))
            End If
            System.IO.File.Move(Pub_Com.folder_Hattyuu & fl, Pub_Com.folder_Hattyuu_kanryou & fl)

            ProBar = CInt(20 + CInt(80 / Pub_Com.file_list_hattyuu.Count) * idx * 1 - 1)

            idx += 1

        Next

        ProBar = 100

        NinnsyouBackgroundWorker.Dispose()
        BackgroundWorker.Dispose()

        Ie.Quit()

    End Sub

    'Step 1 LOGIN IN
    Public Sub DoStep1_Login()

        Pub_Com.AddMsg("OnSite�p�X���[�h����")
        '������ OnSite�p�X���[�h����
        Pub_Com.GetElementBy(Ie, "", "input", "name", "strPassWord").innerText = ConfigurationManager.AppSettings("OnSitePassword").ToString()
        Pub_Com.GetElementBy(Ie, "", "input", "value", "���O�I��").click()

        ProBar = 5

        Pub_Com.AddMsg("�Ɩ��ʑ������j���[")
        ''������ �Ɩ��ʑ������j���[
        Pub_Com.GetElementBy(Ie, "SubHeader", "a", "innertext", "[����]").click()
        ' Pub_Com.WaitComplete(Ie)

        ProBar = 10
        Pub_Com.AddMsg("���̖���")
        ''������ ���̖���
        Pub_Com.GetElementBy(Ie, "Main", "input", "value", "���̖���").click()
        'Pub_Com.WaitComplete(Ie)
        ' System.Threading.Thread.Sleep(1500)

        ProBar = 15

        '�V�K���ς���
        Pub_Com.AddMsg("�V�K����")
        DoStep1_SinkiMitumori()

        ProBar = 20

    End Sub

    'Step 1 �V�K���ς���
    Public Sub DoStep1_SinkiMitumori()
        Dim ShellWindows As New SHDocVw.ShellWindows
        Try
            Dim cIe As SHDocVw.InternetExplorerMedium = GetPopupWindow("OnSite", "mitSearch.asp")
            Pub_Com.GetElementBy(cIe, "", "input", "value", "�V�K����").click()
            System.Threading.Thread.Sleep(100)
            Pub_Com.WaitComplete(Ie)
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
        Pub_Com.WaitComplete(Ie)

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





        ''���ό��o����
        'Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "name", "strJgyCdText").innerText = ���Ə�
        'Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "name", "strTokMeiText").innerText = ���Ӑ�
        'Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "name", "strOtdMeiText").innerText = ���X

        'Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "name", "strBikouMei").innerText = ���l
        'Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "name", "strGenbaMei").innerText = ���ꖼ

        'Pub_Com.GetElementBy(Ie, "fraMitBody", "select", "name", "aryKijyunSyouhinBunrui").setAttribute("value", "A0001,�T�b�V,L90000")
        'Pub_Com.AddMsg("    �[�ޓX�Ȃ��œ�����͂� CLICK")
        'Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "name", "btnUtiwake").click()


        System.Threading.Thread.Sleep(500)
        Pub_Com.WaitComplete(Ie)


        '���ϓ������
        Try
            Dim ele As mshtml.IHTMLElement = Pub_Com.GetElementBy(Ie, "fraMitBody", "DIV", "classname", "ttl")
            If ele.innerText <> "���ϓ������" Then
                DoStep2_Sinki(���Ə�, ���Ӑ�, ���X, ���ꖼ, ���l, ���t�A��, fl)
                Exit Sub
            End If
            Pub_Com.AddMsg("    ���ϓ������ CSV�捞 CLICK")
            Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "CSV�捞").click()
        Catch ex As Exception
            DoStep2_Sinki(���Ə�, ���Ӑ�, ���X, ���ꖼ, ���l, ���t�A��, fl)
            Exit Sub
        End Try


        Pub_Com.WaitComplete(Ie)


        Try

            Me.BackgroundWorker.RunWorkerAsync(folder_Hattyuu & fl)
            System.Threading.Thread.Sleep(500)
            'Dim csvPopup As SHDocVw.InternetExplorerMedium = GetPopupWindow("OnSite", "fileYomikomiSiji.asp")

            Pub_Com.AddMsg("    ���ϓ������ CSV�捞 �Q�@�� CLICK")
            Pub_Com.GetElementBy(GetPopupWindow("OnSite", "fileYomikomiSiji.asp"), "", "input", "value", "�Q�@��").click()

            While Pub_Com.GetElementBy(GetPopupWindow("OnSite", "fileYomikomiSiji.asp"), "", "input", "name", "strFilename").getAttribute("value").ToString = ""
                System.Threading.Thread.Sleep(1)
            End While
            System.Threading.Thread.Sleep(100)


            Pub_Com.AddMsg("    ���ϓ������ CSV�捞 ��@�� CLICK")
            Pub_Com.GetElementBy(GetPopupWindow("OnSite", "fileYomikomiSiji.asp"), "", "input", "value", "��@��").click()
            System.Threading.Thread.Sleep(1000)


            Pub_Com.WaitComplete(Ie)
            Pub_Com.AddMsg("    ���i�R�[�h�������� ���@�� CLICK")
            Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "���@��").click()
            Pub_Com.WaitComplete(Ie)
            Pub_Com.AddMsg("    ���i�R�[�h�������� ���@�� CLICK")
            Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "���@��").click()
            Pub_Com.WaitComplete(Ie)
            Pub_Com.AddMsg("    ���i�R�[�h�������� ���@�� CLICK")
            Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "���@��").click()

            Pub_Com.WaitComplete(Ie)
            Pub_Com.WaitComplete(Ie)
            Pub_Com.WaitComplete(Ie)

            '���@����
            Pub_Com.WaitComplete(Ie)
            Pub_Com.AddMsg("    ���@���� ���@�� CLICK")
            Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "���@��").click()
            Pub_Com.WaitComplete(Ie)
            Pub_Com.WaitComplete(Ie)
            Pub_Com.AddMsg("    ���@���� ���@�� CLICK")
            Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "���@��").click()
            Pub_Com.WaitComplete(Ie)
            Pub_Com.WaitComplete(Ie)
            Pub_Com.AddMsg("    ���@���� ���@�� CLICK")
            Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "���@��").click()
            Pub_Com.WaitComplete(Ie)
            Pub_Com.WaitComplete(Ie)
            Pub_Com.AddMsg("    �P������ ���ϓ�����͂� CLICK")
            Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "���ϓ�����͂�").click()
            Pub_Com.WaitComplete(Ie)
            Dim ele As mshtml.IHTMLElement = Pub_Com.GetElementBy(Ie, "fraMitBody", "DIV", "classname", "ttl")
            Dim kekka As String = ele.innerText
            Pub_Com.AddMsg("    �V�K���� CLICK")
            Pub_Com.GetElementBy(Ie, "fraMitMenu", "a", "innertext", "[�V�K����]").click()
            Pub_Com.WaitComplete(Ie)
            Pub_Com.WaitComplete(Ie)

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
                            Pub_Com.WaitComplete(childIe)
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

#End Region

End Class