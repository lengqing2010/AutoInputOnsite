Public Class Form1

    Dim myWebView

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try


            myWebView = CreateObject("InternetExplorer.Application")
            'Dim uri As System.Uri = New Uri("http://bpn-ap-01d/ons/mit/Scripts/midasiSinkiNyuuryoku.asp?strMitKbn=&strServerName=&strBpnFlg=1&datNoCacheDummy=2018%2F12%2F26+16%3A07%3A46")

            Dim myUri As UriBuilder = New UriBuilder("http://bpn-ap-01d/ons/mit/Scripts/midasiSinkiNyuuryoku.asp?strMitKbn=&strServerName=&strBpnFlg=1&datNoCacheDummy=2018%2F12%2F26+16%3A07%3A46")
            myUri.UserName = "china\shil2"
            myUri.Password = "asdf\@123"
            myUri.Host = "http://bpn-ap-01d"
            myWebView.Navigate(myUri.Uri)
            myWebView.Visible = True

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class