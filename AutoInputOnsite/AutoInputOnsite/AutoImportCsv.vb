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
    'Public Declare Function timeGetTime Lib "winmm.dll" () As Long
    'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


#Region "�������s"



    Public Ie As SHDocVw.InternetExplorerMedium
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

        lv1 = 90 / Pub_Com.file_list_hattyuu.Count


        Dim authHeader As Object = "Authorization: Basic " + _
        Convert.ToBase64String(System.Text.UnicodeEncoding.UTF8.GetBytes(String.Format("{0}:{1}", Pub_Com.user, Pub_Com.password))) + "\r\n"


        '������ OnSite�p�X���[�h���͉��
        Ie.Navigate(Pub_Com.url, , , , authHeader)
        Ie.Silent = True
        Ie.Visible = True


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


            Pub_Com.AddMsg("�ړ�CSV�F" & fl & "��" & Pub_Com.folder_Hattyuu_kanryou)
            If System.IO.File.Exists(Pub_Com.folder_Hattyuu_kanryou & fl) Then
                FileSystem.Rename(Pub_Com.folder_Hattyuu_kanryou & fl, Pub_Com.folder_Hattyuu_kanryou & fl & ".bk." & Now.ToString("yyyyMMddHHmmss"))
            End If
            System.IO.File.Move(Pub_Com.folder_Hattyuu & fl, Pub_Com.folder_Hattyuu_kanryou & fl)
            AddProBar(lv2) '10

            'idx += 1

        Next


        ProBar = 100


    End Sub

    'Step 1 LOGIN IN
    Public Sub DoStep1_Login()


        Pub_Com.AddMsg("OnSite�p�X���[�h����")
        '������ OnSite�p�X���[�h����
        Pub_Com.GetElementBy(Ie, "", "input", "name", "strPassWord").innerText = ConfigurationManager.AppSettings("OnSitePassword").ToString()
        Pub_Com.GetElementBy(Ie, "", "input", "value", "���O�I��").click()

        Pub_Com.SleepAndWaitComplete(Ie)

        '
        Pub_Com.AddMsg("�Ɩ��ʑ������j���[")
        ''������ �Ɩ��ʑ������j���[
        Pub_Com.GetElementBy(Ie, "SubHeader", "a", "innertext", "[����]").click()
        Pub_Com.SleepAndWaitComplete(Ie)


        Pub_Com.AddMsg("���̖���")


        ''������ ���̖���
        Pub_Com.GetElementBy(Ie, "Main", "input", "value", "���̖���").click()
        Pub_Com.SleepAndWaitComplete(Ie)

        '�V�K���ς���
        Pub_Com.AddMsg("�V�K����")
        DoStep1_SinkiMitumori()

    End Sub

    'Step 1 �V�K���ς���
    Public Sub DoStep1_SinkiMitumori()
        Dim ShellWindows As New SHDocVw.ShellWindows

        Dim cIe As SHDocVw.InternetExplorerMedium = GetPopupWindow("OnSite", "mitSearch.asp")
        While cIe Is Nothing
            Com.Sleep5(1000)
            cIe = GetPopupWindow("OnSite", "mitSearch.asp")
        End While

        Pub_Com.GetElementBy(cIe, "", "input", "value", "�V�K����").click()
        Com.Sleep5(100)



        Pub_Com.SleepAndWaitComplete(Ie)


        Exit Sub

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
        Com.Sleep5(500)

        Pub_Com.AddMsg("    ���ϓ������ CSV�捞 �Q�@�� CLICK")

        Dim fra1 As SHDocVw.InternetExplorerMedium = GetPopupWindow("OnSite", "fileYomikomiSiji.asp")
        While fra1 Is Nothing
            fra1 = GetPopupWindow("OnSite", "fileYomikomiSiji.asp")
        End While

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

#End Region

End Class