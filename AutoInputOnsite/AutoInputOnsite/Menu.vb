Public Class Menu

    Public fra As AutoImportCsv
    Public fra2 As AutoImportNouhinsyo

    ''' <summary>
    ''' 見出＆明細作成
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnMs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMs.Click
        Timer1.Stop()
        Timer1.Start()

        pb1.Value = 0
        pb2.Value = 0


        fra = New AutoImportCsv
        If cbHyouji.Checked Then
            fra.Show()
        End If
        fra.DoAll()
        fra.Close()
        pb1.Value = 100

        Me.TopMost = False
        MsgBox("完了")
        Me.TopMost = True

        fra.Dispose()
        fra = Nothing
    End Sub

    ''' <summary>
    ''' 納期指定発注
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnNouhin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNouhin.Click
        Timer1.Stop()
        Timer1.Start()
        pb1.Value = 0
        pb2.Value = 0
        fra2 = New AutoImportNouhinsyo
        If cbHyouji.Checked Then
            fra2.Show()
        End If
        fra2.DoAll()
        fra2.Close()
        pb2.Value = 100
        Me.TopMost = False
        MsgBox("完了")
        Me.TopMost = True
        fra2.Dispose()
        fra2 = Nothing
    End Sub

    ''' <summary>
    ''' 見出＆明細作成  AND  納期指定発注
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAll.Click
        Timer1.Stop()
        Timer1.Start()
        pb1.Value = 0
        pb2.Value = 0

        fra = New AutoImportCsv
        If cbHyouji.Checked Then
            fra.Show()
        End If

        fra.DoAll()
        fra.Close()

        fra.Dispose()
        fra = Nothing

        pb1.Value = 100

        System.Threading.Thread.Sleep(500)

        fra2 = New AutoImportNouhinsyo
        If cbHyouji.Checked Then
            fra2.Show()
        End If
        fra2.DoAll()
        fra2.Close()
        fra2.Dispose()
        fra2 = Nothing

        pb1.Value = 100
        pb2.Value = 100
        Me.TopMost = False
        MsgBox("完了")
        Me.TopMost = True
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If fra IsNot Nothing Then
            pb1.Value = fra.ProBar
        End If
        If fra2 IsNot Nothing Then
            pb2.Value = fra2.ProBar
        End If
    End Sub
End Class