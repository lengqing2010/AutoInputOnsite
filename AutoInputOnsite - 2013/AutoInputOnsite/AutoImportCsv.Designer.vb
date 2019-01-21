<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AutoImportCsv
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使用して変更できます。  
    'コード エディタを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.rtbxMsg = New System.Windows.Forms.RichTextBox()
        Me.SuspendLayout()
        '
        'btnRun
        '
        Me.btnRun.Location = New System.Drawing.Point(367, 0)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(75, 23)
        Me.btnRun.TabIndex = 0
        Me.btnRun.Text = "実行"
        Me.btnRun.UseVisualStyleBackColor = True
        '
        'rtbxMsg
        '
        Me.rtbxMsg.Location = New System.Drawing.Point(0, 29)
        Me.rtbxMsg.Name = "rtbxMsg"
        Me.rtbxMsg.Size = New System.Drawing.Size(678, 209)
        Me.rtbxMsg.TabIndex = 1
        Me.rtbxMsg.Text = ""
        '
        'AutoImportCsv
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(690, 249)
        Me.Controls.Add(Me.rtbxMsg)
        Me.Controls.Add(Me.btnRun)
        Me.Name = "AutoImportCsv"
        Me.Text = "AutoImportCsv"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnRun As System.Windows.Forms.Button
    Friend WithEvents rtbxMsg As System.Windows.Forms.RichTextBox
End Class
