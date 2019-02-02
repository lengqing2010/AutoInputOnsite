Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading.Thread

Public Class RawBrowserPopup
    Inherits System.Windows.Forms.Form

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
Public Shared Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function FindWindowEx( _
        ByVal parentHandle As IntPtr, _
        ByVal childAfter As IntPtr, _
        ByVal lclassName As String, _
        ByVal windowTitle As String) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function SendMessage( _
        ByVal hWnd As IntPtr, _
        ByVal Msg As Integer, _
        ByVal wParam As Integer, _
        ByVal lParam As StringBuilder) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function IsWindowVisible( _
        ByVal hWnd As IntPtr) As Boolean
    End Function

    Private Enum WM As Integer
        SETTEXT = &HC
    End Enum

    Private Enum BM As Integer
        CLICK = &HF5
    End Enum

    'Private WithEvents WebBrowser As New WebBrowser
    Private WithEvents BackgroundWorker As New BackgroundWorker

    Dim WithEvents rawBrowser As RawWebBrowser = New RawWebBrowser()

    Sub New()
        MyBase.New()
        rawBrowser.Dock = DockStyle.Fill
        Me.Controls.Add(rawBrowser)
        rawBrowser.Visible = True
    End Sub

    Public ReadOnly Property WebBrowser() As Object
        Get
            Return rawBrowser
        End Get
    End Property



    Private WithEvents activeX As SHDocVw.InternetExplorer

    Private Sub rawBrowser_Initialized(ByVal sender As Object, ByVal e As EventArgs) Handles rawBrowser.Initialized
        activeX = rawBrowser.ActiveXInstance
        AddHandler activeX.NewWindow2, AddressOf activeX_NewWindow2
        AddHandler activeX.WindowClosing, AddressOf activeX_WindowClosing

        'AddHandler activeX.DocumentComplete, AddressOf a
        Me.BackgroundWorker.RunWorkerAsync(Application.ExecutablePath)

    End Sub

    Private Sub activeX_NewWindow2(ByRef ppDisp As Object, ByRef Cancel As Boolean)
        Dim popup As RawBrowserPopup = New RawBrowserPopup()
        popup.Visible = True
        ppDisp = popup.WebBrowser.ActiveXInstance
    End Sub

    Private Sub activeX_WindowClosing(ByVal IsChildWindow As Boolean, ByRef Cancel As Boolean)
        Cancel = True
        Me.Close()
    End Sub

    Private Sub IE_DocumentComplete(ByVal pDisp As Object, ByRef URL As Object) Handles activeX.DocumentComplete
        Dim Document As Object = activeX.Document
        If Document IsNot Nothing Then
            'Dim doc As HtmlDocument = CType(Document, HtmlDocument)

            'Dim InputElement As HtmlElement = doc.GetElementById("upload")
            'InputElement.InvokeMember("click")
            activeX.Document.GetElementById("fileUpload").Click()
        End If
    End Sub

    'Private Sub a(ByRef pdisp As Object, ByRef url As Object)
    '    Dim Document As HtmlDocument = DirectCast(pdisp, WebBrowser).Document
    '    Dim InputElement As HtmlElement = Document.GetElementById("upload")
    '    InputElement.InvokeMember("click")
    'End Sub

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


    Private Sub InitializeComponent()
        Me.SuspendLayout()
        '
        'RawBrowserPopup
        '
        Me.ClientSize = New System.Drawing.Size(284, 262)
        Me.Name = "RawBrowserPopup"
        Me.ResumeLayout(False)

    End Sub
End Class