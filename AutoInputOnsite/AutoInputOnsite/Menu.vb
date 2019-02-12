Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading.Thread
Imports System.Configuration

Public Class Menu

    Public fra As AutoImportCsv
    Public fra2 As AutoImportNouhinsyo
    Public Ie As SHDocVw.InternetExplorerMedium
    Private WithEvents BackgroundWorker As BackgroundWorker

    ''' <summary>
    ''' 見出＆明細作成
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnMs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMs.Click
        StartAllJunbi()
        DoAutoImportCsv()
        kanryou()
    End Sub
    Private Sub DoAutoImportCsv()
        fra = New AutoImportCsv
        fra.Ie = Ie
        If cbHyouji.Checked Then fra.Show()
        fra.DoAll()
        fra.Close()
        pb1.Value = 100
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

        StartAllJunbi()
        AutoImportNouhinsyo()
      kanryou()
    End Sub
    Private Sub AutoImportNouhinsyo()
        fra2 = New AutoImportNouhinsyo
        fra2.Ie = Ie
        If cbHyouji.Checked Then
            fra2.Show()
        End If
        fra2.insatu = Me.cbInsatu.Checked
        fra2.DoAll()
        fra2.Close()
        fra2.Dispose()
        fra2 = Nothing
        pb2.Value = 100
    End Sub

    ''' <summary>
    ''' 見出＆明細作成  AND  納期指定発注
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAll.Click
        StartAllJunbi()
        DoAutoImportCsv()
        System.Threading.Thread.Sleep(500)
        AutoImportNouhinsyo()
        kanryou()
    End Sub

    ''' <summary>
    ''' 完了
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub kanryou()

        Try
            Ie.Quit()
        Catch ex As Exception

        End Try
        MsgBox("完了")
    End Sub

    ''' <summary>
    ''' Start 準備
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub StartAllJunbi()

re:
        Try
            Ie = Nothing
            Ie = New SHDocVw.InternetExplorerMedium
        Catch ex As Exception
            System.Threading.Thread.Sleep(2000)
            GoTo re
        End Try

        If BackgroundWorker IsNot Nothing Then BackgroundWorker.Dispose()
        BackgroundWorker = New BackgroundWorker
        BackgroundWorker.RunWorkerAsync()

        Timer1.Stop()
        Timer1.Start()
        pb1.Value = 0
        pb2.Value = 0

    End Sub

    ''' <summary>
    ''' Pro bar 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If fra IsNot Nothing Then
            pb1.Value = fra.ProBar
        End If
        If fra2 IsNot Nothing Then
            pb2.Value = fra2.ProBar
        End If
    End Sub

    ''' <summary>
    ''' Menu_Disposed
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Menu_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        Try
            Ie.Quit()
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Menu_Load
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Menu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Ie.Quit()
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Windows Object Something
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BackgroundWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker.DoWork
        Dim com As New Com("")
        com.NewWindowsCom()
    End Sub

End Class