Public Class Form1

    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwd As Long, ByVal buf As String, ByVal bln As Long) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwd As Long, ByVal Msg As Long, ByVal wpara As Long, ByVal lpara As Long) As Long
    Private Declare Function GetWindow Lib "user32" (ByVal hwd As Long, ByVal cmd As Long) As Long
    Private Sub Form_Load()
        'Timer1.Interval = 1000
        Timer1_Timer()
    End Sub

    Private Sub Timer1_Timer()
        'Dim hWnd As Long
        'Dim hChild As Long
        'Dim sWndText As String
        'Dim lRet As Long

        ''�e�E�C���h�E�̖��O"WEB"�̃n���h�������o
        'hWnd = FindWindow(vbNullString, "windows �Z�L�����e�B")



        ''���̐e�n���h������q�E�C���h�E�̃n���h�����o
        'hChild = GetWindow(hWnd, 6)
        ''���̎q�n���h������{�^���̃n���h�������o�H
        'hChild = GetWindow(hChild, 5)

        'sWndText = Space(256)
        ''���̃{�^���̃n���h���̕\�����𒊏o�H
        'lRet = GetWindowText(hChild, sWndText, Len(sWndText))
        'lRet = InStr(sWndText, vbNullChar)
        ''sWndText = Trim(Left(sWndText, lRet - 1))
        'sWndText = sWndText.Substring(0, lRet - 1)
        ''�{�^������"OK"�Ȃ�N���b�N
        'If sWndText = "OK" Then
        '    '�q�E�C���h�E�̃{�^�����A�N�e�B�u�ɂ���H
        '    SendMessage(hChild, &H6, 2, IntPtr.Zero)
        '    '�q�E�C���h�E�̃{�^�����N���b�N�ɂ���
        '    SendMessage(hChild, &HF5, 1, IntPtr.Zero)
        'End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1_Timer()
        Timer1.Interval = 1000
    End Sub
End Class