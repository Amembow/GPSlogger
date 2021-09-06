<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class GPSLogger
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(GPSLogger))
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.OpenFileDialog2 = New System.Windows.Forms.OpenFileDialog()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Button2.Location = New System.Drawing.Point(12, 12)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(112, 23)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Start logging"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Button3.Location = New System.Drawing.Point(12, 38)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(112, 23)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "Stop"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Button1.Location = New System.Drawing.Point(12, 67)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(112, 30)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "WKT読み込み"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'OpenFileDialog2
        '
        Me.OpenFileDialog2.FileName = "OpenFileDialog2"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.SystemColors.Control
        Me.Label1.Location = New System.Drawing.Point(682, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 12)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = ":未読み込み"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(552, 114)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(91, 19)
        Me.TextBox1.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label2.Location = New System.Drawing.Point(499, 117)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(47, 12)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "座標系："
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(552, 146)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(91, 19)
        Me.TextBox2.TabIndex = 8
        Me.TextBox2.Text = "0"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label3.Location = New System.Drawing.Point(487, 149)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(59, 12)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "更新間隔："
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label4.Location = New System.Drawing.Point(14, 134)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(35, 12)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "左右："
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label5.Location = New System.Drawing.Point(55, 134)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(25, 12)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Null"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label6.Location = New System.Drawing.Point(14, 154)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(35, 12)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "誤差："
        '
        'Label7
        '
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label7.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label7.Location = New System.Drawing.Point(8, 101)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(635, 3)
        Me.Label7.TabIndex = 13
        Me.Label7.Text = "Label7"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label8.Location = New System.Drawing.Point(55, 154)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(25, 12)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "Null"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label9.Location = New System.Drawing.Point(14, 114)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(35, 12)
        Me.Label9.TabIndex = 15
        Me.Label9.Text = "方位："
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label10.Location = New System.Drawing.Point(55, 114)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(25, 12)
        Me.Label10.TabIndex = 16
        Me.Label10.Text = "Null"
        '
        'Label11
        '
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.ForeColor = System.Drawing.SystemColors.Control
        Me.Label11.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label11.Location = New System.Drawing.Point(144, 45)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(251, 16)
        Me.Label11.TabIndex = 17
        Me.Label11.Text = "■■■■■■■■■■■■■■■■■■■■"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label12
        '
        Me.Label12.ForeColor = System.Drawing.SystemColors.Control
        Me.Label12.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label12.Location = New System.Drawing.Point(400, 42)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(251, 23)
        Me.Label12.TabIndex = 18
        Me.Label12.Text = "■■■■■■■■■■■■■■■■■■■■"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label13
        '
        Me.Label13.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label13.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label13.Location = New System.Drawing.Point(660, 16)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(3, 150)
        Me.Label13.TabIndex = 19
        Me.Label13.Text = "Label13"
        '
        'Label14
        '
        Me.Label14.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label14.Location = New System.Drawing.Point(143, 42)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(505, 1)
        Me.Label14.TabIndex = 20
        Me.Label14.Text = "Label14"
        '
        'Label15
        '
        Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label15.Location = New System.Drawing.Point(143, 32)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(1, 10)
        Me.Label15.TabIndex = 21
        Me.Label15.Text = "Label15"
        '
        'Label16
        '
        Me.Label16.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label16.Location = New System.Drawing.Point(396, 32)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(1, 10)
        Me.Label16.TabIndex = 22
        Me.Label16.Text = "Label16"
        '
        'Label17
        '
        Me.Label17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label17.Location = New System.Drawing.Point(647, 32)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(1, 10)
        Me.Label17.TabIndex = 23
        Me.Label17.Text = "Label17"
        '
        'Label18
        '
        Me.Label18.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label18.Location = New System.Drawing.Point(271, 32)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(1, 10)
        Me.Label18.TabIndex = 24
        Me.Label18.Text = "Label18"
        '
        'Label19
        '
        Me.Label19.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label19.Location = New System.Drawing.Point(521, 32)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(1, 10)
        Me.Label19.TabIndex = 25
        Me.Label19.Text = "Label19"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.GPSlogger.My.Resources.Resources.A2
        Me.PictureBox1.InitialImage = Global.GPSlogger.My.Resources.Resources.A2
        Me.PictureBox1.Location = New System.Drawing.Point(684, 33)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(100, 100)
        Me.PictureBox1.TabIndex = 26
        Me.PictureBox1.TabStop = False
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(739, 147)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(45, 19)
        Me.TextBox3.TabIndex = 27
        Me.TextBox3.Text = "1"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label20.Location = New System.Drawing.Point(674, 150)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(59, 12)
        Me.Label20.TabIndex = 28
        Me.Label20.Text = "角度倍率："
        '
        'GPSLogger
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ClientSize = New System.Drawing.Size(806, 177)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "GPSLogger"
        Me.Text = "GPSLogger"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents OpenFileDialog2 As OpenFileDialog
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents Label12 As Label
    Friend WithEvents Label13 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents Label16 As Label
    Friend WithEvents Label17 As Label
    Friend WithEvents Label18 As Label
    Friend WithEvents Label19 As Label
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents Label20 As Label
End Class
