<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Menu
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Menu))
        Me.btnMs = New System.Windows.Forms.Button
        Me.btnNouhin = New System.Windows.Forms.Button
        Me.btnAll = New System.Windows.Forms.Button
        Me.cbHyouji = New System.Windows.Forms.CheckBox
        Me.pb1 = New System.Windows.Forms.ProgressBar
        Me.pb2 = New System.Windows.Forms.ProgressBar
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.cbInsatu = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'btnMs
        '
        Me.btnMs.Font = New System.Drawing.Font("HGP創英ﾌﾟﾚｾﾞﾝｽEB", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnMs.ForeColor = System.Drawing.Color.Blue
        Me.btnMs.Location = New System.Drawing.Point(12, 113)
        Me.btnMs.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.btnMs.Name = "btnMs"
        Me.btnMs.Size = New System.Drawing.Size(122, 34)
        Me.btnMs.TabIndex = 0
        Me.btnMs.Text = "見出＆明細作成"
        Me.btnMs.UseVisualStyleBackColor = True
        '
        'btnNouhin
        '
        Me.btnNouhin.Font = New System.Drawing.Font("HGP創英ﾌﾟﾚｾﾞﾝｽEB", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnNouhin.ForeColor = System.Drawing.Color.Blue
        Me.btnNouhin.Location = New System.Drawing.Point(135, 113)
        Me.btnNouhin.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.btnNouhin.Name = "btnNouhin"
        Me.btnNouhin.Size = New System.Drawing.Size(118, 34)
        Me.btnNouhin.TabIndex = 1
        Me.btnNouhin.Text = "納期指定発注"
        Me.btnNouhin.UseVisualStyleBackColor = True
        '
        'btnAll
        '
        Me.btnAll.Font = New System.Drawing.Font("HGP創英ﾌﾟﾚｾﾞﾝｽEB", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnAll.ForeColor = System.Drawing.Color.Blue
        Me.btnAll.Location = New System.Drawing.Point(253, 113)
        Me.btnAll.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.btnAll.Name = "btnAll"
        Me.btnAll.Size = New System.Drawing.Size(251, 34)
        Me.btnAll.TabIndex = 2
        Me.btnAll.Text = "見出＆明細作成 And 納期指定発注"
        Me.btnAll.UseVisualStyleBackColor = True
        '
        'cbHyouji
        '
        Me.cbHyouji.AutoSize = True
        Me.cbHyouji.Font = New System.Drawing.Font("HGP創英ﾌﾟﾚｾﾞﾝｽEB", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cbHyouji.ForeColor = System.Drawing.Color.GhostWhite
        Me.cbHyouji.Location = New System.Drawing.Point(23, 176)
        Me.cbHyouji.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbHyouji.Name = "cbHyouji"
        Me.cbHyouji.Size = New System.Drawing.Size(86, 19)
        Me.cbHyouji.TabIndex = 3
        Me.cbHyouji.Text = "明細表示"
        Me.cbHyouji.UseVisualStyleBackColor = True
        '
        'pb1
        '
        Me.pb1.Location = New System.Drawing.Point(12, 29)
        Me.pb1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pb1.Name = "pb1"
        Me.pb1.Size = New System.Drawing.Size(584, 26)
        Me.pb1.TabIndex = 4
        '
        'pb2
        '
        Me.pb2.Location = New System.Drawing.Point(12, 75)
        Me.pb2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pb2.Name = "pb2"
        Me.pb2.Size = New System.Drawing.Size(584, 26)
        Me.pb2.TabIndex = 4
        '
        'Timer1
        '
        Me.Timer1.Interval = 10
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("HGP創英ﾌﾟﾚｾﾞﾝｽEB", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.GhostWhite
        Me.Label1.Location = New System.Drawing.Point(12, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 15)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "見出＆明細作成"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("HGP創英ﾌﾟﾚｾﾞﾝｽEB", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.GhostWhite
        Me.Label2.Location = New System.Drawing.Point(12, 59)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(97, 15)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "納期指定発注"
        '
        'cbInsatu
        '
        Me.cbInsatu.AutoSize = True
        Me.cbInsatu.Checked = True
        Me.cbInsatu.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbInsatu.Font = New System.Drawing.Font("HGP創英ﾌﾟﾚｾﾞﾝｽEB", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cbInsatu.ForeColor = System.Drawing.Color.GhostWhite
        Me.cbInsatu.Location = New System.Drawing.Point(135, 176)
        Me.cbInsatu.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbInsatu.Name = "cbInsatu"
        Me.cbInsatu.Size = New System.Drawing.Size(56, 19)
        Me.cbInsatu.TabIndex = 7
        Me.cbInsatu.Text = "印刷"
        Me.cbInsatu.UseVisualStyleBackColor = True
        '
        'Menu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ClientSize = New System.Drawing.Size(609, 208)
        Me.Controls.Add(Me.cbInsatu)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.pb2)
        Me.Controls.Add(Me.pb1)
        Me.Controls.Add(Me.cbHyouji)
        Me.Controls.Add(Me.btnAll)
        Me.Controls.Add(Me.btnNouhin)
        Me.Controls.Add(Me.btnMs)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "Menu"
        Me.Opacity = 0.9
        Me.Text = "「見出＆明細作成」・「納期指定発注」"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnMs As System.Windows.Forms.Button
    Friend WithEvents btnNouhin As System.Windows.Forms.Button
    Friend WithEvents btnAll As System.Windows.Forms.Button
    Friend WithEvents cbHyouji As System.Windows.Forms.CheckBox
    Friend WithEvents pb1 As System.Windows.Forms.ProgressBar
    Friend WithEvents pb2 As System.Windows.Forms.ProgressBar
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbInsatu As System.Windows.Forms.CheckBox
End Class
