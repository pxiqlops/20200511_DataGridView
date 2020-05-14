<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button10 = New System.Windows.Forms.Button()
        Me.Button11 = New System.Windows.Forms.Button()
        Me.Button12 = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToOrderColumns = True
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(25, 68)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 51
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(728, 428)
        Me.DataGridView1.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(778, 27)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(146, 35)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "表示初期化"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(778, 68)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(146, 35)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "csvファイル読込"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(778, 109)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(146, 35)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "xlsファイル読込"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(778, 150)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(146, 35)
        Me.Button4.TabIndex = 4
        Me.Button4.Text = "xlsxファイル読込"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(778, 191)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(146, 35)
        Me.Button5.TabIndex = 6
        Me.Button5.Text = "mdbファイル読込"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(778, 232)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(146, 35)
        Me.Button6.TabIndex = 5
        Me.Button6.Text = "accdbファイル読込"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(778, 273)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(146, 35)
        Me.Button7.TabIndex = 9
        Me.Button7.Text = "csvファイル出力"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Button9
        '
        Me.Button9.Location = New System.Drawing.Point(778, 355)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(146, 35)
        Me.Button9.TabIndex = 10
        Me.Button9.Text = "Excelファイル出力"
        Me.Button9.UseVisualStyleBackColor = True
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(778, 314)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(146, 35)
        Me.Button8.TabIndex = 11
        Me.Button8.Text = "txtファイル出力"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(25, 27)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(728, 22)
        Me.TextBox1.TabIndex = 12
        '
        'Button10
        '
        Me.Button10.Location = New System.Drawing.Point(930, 68)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(146, 35)
        Me.Button10.TabIndex = 13
        Me.Button10.Text = "csv上書き保存"
        Me.Button10.UseVisualStyleBackColor = True
        '
        'Button11
        '
        Me.Button11.Enabled = False
        Me.Button11.Font = New System.Drawing.Font("MS UI Gothic", 7.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Button11.Location = New System.Drawing.Point(930, 109)
        Me.Button11.Name = "Button11"
        Me.Button11.Size = New System.Drawing.Size(146, 35)
        Me.Button11.TabIndex = 14
        Me.Button11.Text = "xls上書き保存" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "※使用不可"
        Me.Button11.UseVisualStyleBackColor = True
        Me.Button11.UseWaitCursor = True
        '
        'Button12
        '
        Me.Button12.Location = New System.Drawing.Point(930, 150)
        Me.Button12.Name = "Button12"
        Me.Button12.Size = New System.Drawing.Size(146, 35)
        Me.Button12.TabIndex = 15
        Me.Button12.Text = "xlsx上書き保存"
        Me.Button12.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1106, 531)
        Me.Controls.Add(Me.Button12)
        Me.Controls.Add(Me.Button11)
        Me.Controls.Add(Me.Button10)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.Button9)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "Form1"
        Me.Text = "ファイル操作検証"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents Button7 As Button
    Friend WithEvents Button9 As Button
    Friend WithEvents Button8 As Button
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Button10 As Button
    Friend WithEvents Button11 As Button
    Friend WithEvents Button12 As Button
End Class
