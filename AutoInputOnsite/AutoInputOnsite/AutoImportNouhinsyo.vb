Option Explicit On
Option Strict On

Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading.Thread
Imports System.Configuration

Public Class AutoImportNouhinsyo

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


    Private Sub SavePdf(ByVal csvFileName As String)
        Dim pdfPath As String = ConfigurationManager.AppSettings("Pdf_Path").ToString()

        If System.IO.File.Exists(pdfPath & csvFileName & ".pdf") Then
            FileSystem.Rename(pdfPath & csvFileName & ".pdf", pdfPath & csvFileName & ".pdf" & ".bk." & Now.ToString("yyyyMMddHHmmss"))
        End If

        Dim hWnd As IntPtr
        Do While hWnd = IntPtr.Zero
            hWnd = FindWindow("#32770", "�R�s�[��ۑ�")
            Sleep(1)
        Loop
        Dim filePathComboBoxEx32 As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "ComboBoxEx32", String.Empty)
        Dim filePathComboBox As IntPtr = FindWindowEx(filePathComboBoxEx32, IntPtr.Zero, "ComboBox", String.Empty)
        Dim hEditfilePath As IntPtr = FindWindowEx(filePathComboBox, IntPtr.Zero, "Edit", String.Empty)
        SendMessage(hEditfilePath, WM.SETTEXT, 0, New StringBuilder(pdfPath & csvFileName & ".pdf"))

        Dim saveButton As IntPtr = FindWindowEx(hWnd, IntPtr.Zero, "Button", "�ۑ�")
        Dim hEdit2 As IntPtr = FindWindowEx(saveButton, IntPtr.Zero, "Button", "�ۑ�")
        SendMessage(saveButton, BM.CLICK, 0, Nothing)

        'Dim CtrlNotifySink1 As IntPtr = FindWindowEx(DirectUIHWND, IntPtr.Zero, "CtrlNotifySink", String.Empty)
        'Dim CtrlNotifySink2 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink1, "CtrlNotifySink", String.Empty)
        'Dim CtrlNotifySink3 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink2, "CtrlNotifySink", String.Empty)
        'Dim CtrlNotifySink4 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink3, "CtrlNotifySink", String.Empty)
        'Dim CtrlNotifySink5 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink4, "CtrlNotifySink", String.Empty)
        'Dim CtrlNotifySink6 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink5, "CtrlNotifySink", String.Empty)
        'Dim CtrlNotifySink7 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink6, "CtrlNotifySink", String.Empty)
        'Dim CtrlNotifySink8 As IntPtr = FindWindowEx(DirectUIHWND, CtrlNotifySink7, "CtrlNotifySink", String.Empty)

        'Dim hEdit As IntPtr = FindWindowEx(CtrlNotifySink7, IntPtr.Zero, "Edit", String.Empty)
        'SendMessage(hEdit, WM.SETTEXT, 0, New StringBuilder(user))
        'Dim hEdit1 As IntPtr = FindWindowEx(CtrlNotifySink8, IntPtr.Zero, "Edit", String.Empty)
        'SendMessage(hEdit1, WM.SETTEXT, 0, New StringBuilder(password))
        'Dim hEdit2 As IntPtr = FindWindowEx(CtrlNotifySink3, IntPtr.Zero, "Button", "OK")
        'SendMessage(hEdit2, BM.CLICK, 0, Nothing)

        'SavePdf()
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

        'SavePdf("aaaa.csv")
        Pub_Com = New Com("�[���w�蔭��" & Now.ToString("yyyyMMddHHmmss"))
        If Pub_Com.file_list_Nouki.Count = 0 Then
            ProBar = 100
            Exit Sub
        End If

        '����
        Dim firsOpenKbn As Boolean = True

        '������New Object
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

        'CSV �t�@�C���� �捞
        For zzz As Integer = 0 To Pub_Com.file_list_Nouki.Count - 1

            Dim csvFileName As String = Pub_Com.file_list_Nouki(zzz).ToString.Trim
            Dim csvFileNames() As String = Pub_Com.file_list_Nouki(zzz).ToString.Trim.Split("-"c)
            Dim ���Ə�, ���Ӑ�, ���X, ���ꖼ, ���l, ���t�A�� As String
            ���Ə� = csvFileNames(0)
            ���Ӑ� = csvFileNames(1)
            ���X = csvFileNames(2)
            ���ꖼ = csvFileNames(3)
            ���l = csvFileNames(4)
            ���t�A�� = csvFileNames(5)


            If firsOpenKbn = False Then
                Pub_Com.GetElementBy(Ie, "fraHead", "input", "value", "�i������").click()
                Pub_Com.SleepAndWaitComplete(Ie)
            End If
            firsOpenKbn = False

            Pub_Com.AddMsg("�捞�F" & Pub_Com.file_list_Nouki(zzz).ToString.Trim)

            ProBar = CInt(20 + CInt(80 / Pub_Com.file_list_Nouki.Count) * idx * 0.25 - 1)
            '���ό���
            Pub_Com.AddMsg("���ό���")
            DoStep1_PoupuSentaku(���Ə�, ���Ӑ�, ���X, ���ꖼ, ���l, ���t�A��, csvFileName)
            Pub_Com.SleepAndWaitComplete(Ie)


            ProBar = CInt(20 + CInt(80 / Pub_Com.file_list_Nouki.Count) * idx * 0.35 - 1)
            '�[�����ݒ�
            If Not DoStep2_Set() Then
                Continue For
            End If
            Pub_Com.SleepAndWaitComplete(Ie)
            Pub_Com.SleepAndWaitComplete(Ie)
            Pub_Com.SleepAndWaitComplete(Ie)

            '�Y���f�[�^������܂��� NEXT
            Dim fraTmp As mshtml.HTMLWindow2 = Pub_Com.GetFrameByName(Ie, "fraHyou")
            If fraTmp IsNot Nothing Then
                If fraTmp.document.body.innerText.IndexOf("�Y���f�[�^������܂���") >= 0 Then
                    Continue For
                End If
            End If


            'CSV�t�@�C�����e�捞
            Dim strData As String() = System.IO.File.ReadAllLines(Pub_Com.folder_Nouki & csvFileName)
            Dim code As String = ""
            Dim nouki As String = ""

            ProBar = CInt(20 + CInt(80 / Pub_Com.file_list_Nouki.Count) * idx * 0.4 - 1)

            'CSV LINES
            For jjj As Integer = 0 To strData.Length - 1

                If strData(jjj).Trim <> "" Then

                    '�R�[�h �[��
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
                                MsgBox("�[�i��]���F[" & nouki & "]������܂���")
                                Exit Sub
                            End If
                        End If

                    Next

                End If

            Next

            ProBar = CInt(20 + CInt(80 / Pub_Com.file_list_Nouki.Count) * idx * 0.8 - 1)

            Pub_Com.GetElementBy(Ie, "fraMitBody", "select", "name", "strBukkenKbn").setAttribute("value", "01")
            Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "���@��").click()
            Pub_Com.SleepAndWaitComplete(Ie)
            Pub_Com.SleepAndWaitComplete(Ie)
            Pub_Com.SleepAndWaitComplete(Ie)

            Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "�������ʏƉ��").click()
            Pub_Com.SleepAndWaitComplete(Ie)

            Pub_Com.GetElementBy(Ie, "fraMitBody", "input", "value", "���ʈ��").click()
            Pub_Com.SleepAndWaitComplete(Ie)
            'Save PDF
            Dim cIe2 As SHDocVw.InternetExplorerMedium = GetPopupWindowPDF()
            Pub_Com.SleepAndWaitComplete(cIe2, 100)
            Pub_Com.SleepAndWaitComplete(cIe2, 100)
            Pub_Com.SleepAndWaitComplete(cIe2, 100)
            Pub_Com.SleepAndWaitComplete(cIe2, 100)

            System.Threading.Thread.Sleep(500)

            CType(cIe2, SHDocVw.InternetExplorerMedium).ExecWB(SHDocVw.OLECMDID.OLECMDID_SAVEAS, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_PROMPTUSER)
            System.Threading.Thread.Sleep(500)
            SavePdf(csvFileName)
            System.Threading.Thread.Sleep(500)

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
            '            '            System.Threading.Thread.Sleep(500)
            '            '            Return childIe
            '            '        End If
            '            '    End If
            '        End If
            '    End If
            'Next
            cIe2.Quit()
            Pub_Com.GetElementBy(Ie, "fraMitMenu", "a", "innertext", "[���ψꗗ���ĕ\��]").click()
            Pub_Com.SleepAndWaitComplete(Ie)




            Pub_Com.AddMsg("�ړ�CSV�F" & csvFileName & "��" & Pub_Com.folder_Nouki_kanryou)
            If System.IO.File.Exists(Pub_Com.folder_Nouki_kanryou & csvFileName) Then
                FileSystem.Rename(Pub_Com.folder_Nouki_kanryou & csvFileName, Pub_Com.folder_Nouki_kanryou & csvFileName & ".bk." & Now.ToString("yyyyMMddHHmmss"))
            End If

            System.IO.File.Move(Pub_Com.folder_Nouki & csvFileName, Pub_Com.folder_Nouki_kanryou & csvFileName)
            ProBar = CInt(20 + CInt(80 / Pub_Com.file_list_Nouki.Count) * idx * 1 - 1)
            idx += 1

        Next


        NinnsyouBackgroundWorker.Dispose()
        BackgroundWorker.Dispose()

        Ie.Quit()

        ProBar = 100

        'MsgBox("����")


    End Sub

    'Step 1 LOGIN IN
    Public Sub DoStep1_Login()

        ProBar = 5

        Pub_Com.AddMsg("OnSite�p�X���[�h����")
        '������ OnSite�p�X���[�h����
        Pub_Com.GetElementBy(Ie, "", "input", "name", "strPassWord").innerText = ConfigurationManager.AppSettings("OnSitePassword").ToString()
        Pub_Com.GetElementBy(Ie, "", "input", "value", "���O�I��").click()

        ProBar = 10
        Pub_Com.AddMsg("�Ɩ��ʑ������j���[")
        ''������ �Ɩ��ʑ������j���[
        Pub_Com.GetElementBy(Ie, "SubHeader", "a", "innertext", "[����]").click()

        ProBar = 15
        Pub_Com.AddMsg("���̖���")
        ''������ ���̖���
        Pub_Com.GetElementBy(Ie, "Main", "input", "value", "���̖���").click()
        Pub_Com.SleepAndWaitComplete(Ie, 200)



    End Sub

    'Step 1 �V�K���ς���
    Public Sub DoStep1_PoupuSentaku(ByVal ���Ə� As String, ByVal ���Ӑ� As String, ByVal ���X As String, ByVal ���ꖼ As String, ByVal ���l As String, ByVal ���t�A�� As String, ByVal fl As String)
        Dim ShellWindows As New SHDocVw.ShellWindows
        Try

            Pub_Com.AddMsg("���ό��� POPUP")
            Dim cIe As SHDocVw.InternetExplorerMedium = GetPopupWindow("OnSite", "mitSearch.asp")

            While cIe Is Nothing
                System.Threading.Thread.Sleep(100)
                cIe = GetPopupWindow("OnSite", "mitSearch.asp")
            End While

            Pub_Com.SleepAndWaitComplete(cIe)
            Pub_Com.GetElementBy(cIe, "", "input", "name", "strGenbaMei").innerText = ���ꖼ
            Pub_Com.GetElementBy(cIe, "", "input", "name", "strUriJgy").click()
            Pub_Com.SleepAndWaitComplete(cIe)



            Pub_Com.AddMsg("���Ə����� POPUP")
            Dim cIe2 As SHDocVw.InternetExplorerMedium = GetPopupWindow("OnSite", "jgyKensaku.asp")
            While cIe2 Is Nothing
                System.Threading.Thread.Sleep(100)
                cIe2 = GetPopupWindow("OnSite", "jgyKensaku.asp")
            End While

            Pub_Com.SleepAndWaitComplete(cIe2)
            Pub_Com.GetElementBy(cIe2, "", "input", "name", "strJgyCd").innerText = ���Ə�
            Pub_Com.GetElementBy(cIe2, "", "input", "value", "���@��").click()

            Pub_Com.SleepAndWaitComplete(cIe2)
            Pub_Com.SleepAndWaitComplete(cIe)
            Pub_Com.GetElementBy(cIe, "", "input", "value", "���@��").click()

            Pub_Com.SleepAndWaitComplete(Ie)

            Pub_Com.SleepAndWaitComplete(Ie)

            '50���ȏ�̏ꍇ
            Try
                Dim Doc As mshtml.HTMLDocument
                Dim eles As mshtml.IHTMLElementCollection
                Doc = CType(Ie.Document, mshtml.HTMLDocument)
                Dim fra As mshtml.HTMLWindow2 = Pub_Com.GetFrameWait(Ie, "fraHyou")
                Doc = CType(fra.document, mshtml.HTMLDocument)
                eles = Doc.getElementsByTagName("input")

                For Each ele As mshtml.IHTMLElement In eles
                    Try
                        If ele.getAttribute("value").ToString = "�p ��" Then
                            ele.click()
                            Pub_Com.SleepAndWaitComplete(Ie)
                            Exit For
                        End If

                    Catch ex As Exception

                    End Try

                Next
            Catch ex As Exception

            End Try


            Exit Sub
        Catch ex As Exception
            'DoStep1_SinkiMitumori(���Ə�, ���Ӑ�, ���X, ���ꖼ, ���l, ���t�A��, fl)
            Exit Sub
        End Try
    End Sub

    'Step 2 �iWHILE�j
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
                    If td.innerText = "�쐬��" Then
                        ele.click()
                        Pub_Com.GetElementBy(Ie, "fraHead", "input", "value", "�����[����\��").click()
                        Pub_Com.SleepAndWaitComplete(Ie)


                        Dim Doc1 As mshtml.HTMLDocument = CType(Ie.Document, mshtml.HTMLDocument)
                        Dim fra1 As mshtml.HTMLWindow2 = Pub_Com.GetFrameWait(Ie, "fraMitBody")
                        Doc1 = CType(fra1.document, mshtml.HTMLDocument)

                        If Doc1.body.innerText.IndexOf("�����\�Ȍ��ςł͂���܂���") > 0 Then
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

    'POPUP Window �擾
    Public Function GetPopupWindow(ByVal titleKey As String, ByVal fileNameKey As String) As SHDocVw.InternetExplorerMedium
        Dim ShellWindows As New SHDocVw.ShellWindows
        For Each childIe As SHDocVw.InternetExplorerMedium In ShellWindows
            System.Windows.Forms.Application.DoEvents()
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

    Public Function GetPopupWindowPDF() As SHDocVw.InternetExplorerMedium
        Dim ShellWindows As New SHDocVw.ShellWindows
        For Each childIe As SHDocVw.InternetExplorerMedium In ShellWindows
            System.Windows.Forms.Application.DoEvents()
            Dim filename As String = System.IO.Path.GetFileNameWithoutExtension(Ie.FullName).ToLower()
            If filename = "iexplore" Then
                If CType(childIe, SHDocVw.InternetExplorerMedium).LocationURL.Contains("servlet") Then
                    Return childIe
                End If
            End If
        Next
        'System.Threading.Thread.Sleep(10)
        Return GetPopupWindowPDF()
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

    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox1.TextChanged

    End Sub
End Class