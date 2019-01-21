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

        ''親ウインドウの名前"WEB"のハンドルを検出
        'hWnd = FindWindow(vbNullString, "windows セキュリティ")



        ''その親ハンドルから子ウインドウのハンドル検出
        'hChild = GetWindow(hWnd, 6)
        ''その子ハンドルからボタンのハンドルを検出？
        'hChild = GetWindow(hChild, 5)

        'sWndText = Space(256)
        ''そのボタンのハンドルの表示名を抽出？
        'lRet = GetWindowText(hChild, sWndText, Len(sWndText))
        'lRet = InStr(sWndText, vbNullChar)
        ''sWndText = Trim(Left(sWndText, lRet - 1))
        'sWndText = sWndText.Substring(0, lRet - 1)
        ''ボタン名が"OK"ならクリック
        'If sWndText = "OK" Then
        '    '子ウインドウのボタンをアクティブにする？
        '    SendMessage(hChild, &H6, 2, IntPtr.Zero)
        '    '子ウインドウのボタンをクリックにする
        '    SendMessage(hChild, &HF5, 1, IntPtr.Zero)
        'End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1_Timer()
        Timer1.Interval = 1000
    End Sub
End Class