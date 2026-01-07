<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FmCaculate_Rpt
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FmCaculate_Rpt))
        Me.FG = New System.Windows.Forms.DataGridView
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.FG2 = New System.Windows.Forms.DataGridView
        Me.BtnSave = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.Button7 = New System.Windows.Forms.Button
        Me.FontStype = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.FontColor = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.CLock = New System.Windows.Forms.CheckBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.Button8 = New System.Windows.Forms.Button
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.Button9 = New System.Windows.Forms.Button
        Me.Button10 = New System.Windows.Forms.Button
        Me.Button11 = New System.Windows.Forms.Button
        Me.Button12 = New System.Windows.Forms.Button
        Me.Button13 = New System.Windows.Forms.Button
        CType(Me.FG, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FG2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'FG
        '
        Me.FG.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.FG.BackgroundColor = System.Drawing.Color.White
        Me.FG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.FG.Location = New System.Drawing.Point(667, -19)
        Me.FG.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.FG.Name = "FG"
        Me.FG.Size = New System.Drawing.Size(10, 60)
        Me.FG.TabIndex = 0
        Me.FG.Visible = False
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(189, 137)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(60, 30)
        Me.TextBox1.TabIndex = 2
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"BLS", "INC", "CAF", "PRO1", "PRO2", "PRO3", "PRO4"})
        Me.ComboBox1.Location = New System.Drawing.Point(54, 137)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(78, 29)
        Me.ComboBox1.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 142)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 21)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Type:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(138, 142)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(50, 21)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Rpt_ID"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(255, 142)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 21)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Name:"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(309, 137)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.ReadOnly = True
        Me.TextBox2.Size = New System.Drawing.Size(429, 30)
        Me.TextBox2.TabIndex = 6
        '
        'FG2
        '
        Me.FG2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.FG2.BackgroundColor = System.Drawing.Color.White
        Me.FG2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.FG2.Location = New System.Drawing.Point(12, 247)
        Me.FG2.Name = "FG2"
        Me.FG2.Size = New System.Drawing.Size(725, 436)
        Me.FG2.TabIndex = 8
        '
        'BtnSave
        '
        Me.BtnSave.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnSave.Location = New System.Drawing.Point(50, 6)
        Me.BtnSave.Name = "BtnSave"
        Me.BtnSave.Size = New System.Drawing.Size(84, 35)
        Me.BtnSave.TabIndex = 92
        Me.BtnSave.Tag = "3006"
        Me.BtnSave.Text = "Save"
        Me.BtnSave.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Red
        Me.Button2.Location = New System.Drawing.Point(15, 6)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(35, 35)
        Me.Button2.TabIndex = 91
        Me.Button2.Tag = "9999"
        Me.Button2.Text = "Exit"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Blue
        Me.Button1.Location = New System.Drawing.Point(220, 170)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(41, 35)
        Me.Button1.TabIndex = 93
        Me.Button1.Text = "/"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.Blue
        Me.Button3.Location = New System.Drawing.Point(179, 170)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(41, 35)
        Me.Button3.TabIndex = 94
        Me.Button3.Text = "*"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.Blue
        Me.Button4.Location = New System.Drawing.Point(97, 170)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(41, 35)
        Me.Button4.TabIndex = 96
        Me.Button4.Text = "+"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.ForeColor = System.Drawing.Color.Blue
        Me.Button5.Location = New System.Drawing.Point(138, 170)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(41, 35)
        Me.Button5.TabIndex = 95
        Me.Button5.Text = "-"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button6.ForeColor = System.Drawing.Color.Blue
        Me.Button6.Location = New System.Drawing.Point(15, 170)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(41, 35)
        Me.Button6.TabIndex = 98
        Me.Button6.Text = "("
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button7.ForeColor = System.Drawing.Color.Blue
        Me.Button7.Location = New System.Drawing.Point(56, 170)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(41, 35)
        Me.Button7.TabIndex = 97
        Me.Button7.Text = ")"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'FontStype
        '
        Me.FontStype.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.FontStype.FormattingEnabled = True
        Me.FontStype.Items.AddRange(New Object() {"ຕົວຈາງ", "ຕົວເຂັ້ມ"})
        Me.FontStype.Location = New System.Drawing.Point(501, 209)
        Me.FontStype.Name = "FontStype"
        Me.FontStype.Size = New System.Drawing.Size(124, 29)
        Me.FontStype.TabIndex = 99
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(438, 212)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(44, 21)
        Me.Label4.TabIndex = 100
        Me.Label4.Text = "Stype"
        '
        'FontColor
        '
        Me.FontColor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.FontColor.FormattingEnabled = True
        Me.FontColor.Items.AddRange(New Object() {"ສີ ດໍາ", "ສີ ແດງ", "ສີ ຂຽວ"})
        Me.FontColor.Location = New System.Drawing.Point(501, 174)
        Me.FontColor.Name = "FontColor"
        Me.FontColor.Size = New System.Drawing.Size(124, 29)
        Me.FontColor.TabIndex = 101
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(438, 178)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 21)
        Me.Label5.TabIndex = 102
        Me.Label5.Text = "ForeColor"
        '
        'CLock
        '
        Me.CLock.AutoSize = True
        Me.CLock.Location = New System.Drawing.Point(631, 179)
        Me.CLock.Name = "CLock"
        Me.CLock.Size = New System.Drawing.Size(113, 25)
        Me.CLock.TabIndex = 103
        Me.CLock.Text = "Lock for save"
        Me.CLock.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(631, 209)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(84, 25)
        Me.CheckBox1.TabIndex = 104
        Me.CheckBox1.Text = "Underline"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Button8
        '
        Me.Button8.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button8.ForeColor = System.Drawing.Color.Blue
        Me.Button8.Location = New System.Drawing.Point(179, 206)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(82, 35)
        Me.Button8.TabIndex = 105
        Me.Button8.Text = "DEL"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'TextBox3
        '
        Me.TextBox3.BackColor = System.Drawing.Color.Black
        Me.TextBox3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox3.ForeColor = System.Drawing.Color.White
        Me.TextBox3.Location = New System.Drawing.Point(15, 43)
        Me.TextBox3.Multiline = True
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.ReadOnly = True
        Me.TextBox3.Size = New System.Drawing.Size(725, 90)
        Me.TextBox3.TabIndex = 106
        '
        'Button9
        '
        Me.Button9.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button9.ForeColor = System.Drawing.Color.Blue
        Me.Button9.Location = New System.Drawing.Point(263, 206)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(82, 35)
        Me.Button9.TabIndex = 108
        Me.Button9.Text = "Refresh"
        Me.Button9.UseVisualStyleBackColor = True
        '
        'Button10
        '
        Me.Button10.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button10.ForeColor = System.Drawing.Color.Blue
        Me.Button10.Location = New System.Drawing.Point(346, 206)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(86, 35)
        Me.Button10.TabIndex = 109
        Me.Button10.Text = "Test"
        Me.Button10.UseVisualStyleBackColor = True
        '
        'Button11
        '
        Me.Button11.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button11.ForeColor = System.Drawing.Color.Blue
        Me.Button11.Location = New System.Drawing.Point(263, 169)
        Me.Button11.Name = "Button11"
        Me.Button11.Size = New System.Drawing.Size(169, 35)
        Me.Button11.TabIndex = 110
        Me.Button11.Text = "<= BSP"
        Me.Button11.UseVisualStyleBackColor = True
        '
        'Button12
        '
        Me.Button12.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button12.ForeColor = System.Drawing.Color.Blue
        Me.Button12.Location = New System.Drawing.Point(16, 206)
        Me.Button12.Name = "Button12"
        Me.Button12.Size = New System.Drawing.Size(81, 35)
        Me.Button12.TabIndex = 112
        Me.Button12.Text = "Cast(("
        Me.Button12.UseVisualStyleBackColor = True
        '
        'Button13
        '
        Me.Button13.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button13.ForeColor = System.Drawing.Color.Blue
        Me.Button13.Location = New System.Drawing.Point(97, 206)
        Me.Button13.Name = "Button13"
        Me.Button13.Size = New System.Drawing.Size(82, 35)
        Me.Button13.TabIndex = 111
        Me.Button13.Text = ")As Float)"
        Me.Button13.UseVisualStyleBackColor = True
        '
        'FmCaculate_Rpt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 21.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Silver
        Me.ClientSize = New System.Drawing.Size(750, 660)
        Me.Controls.Add(Me.Button12)
        Me.Controls.Add(Me.Button13)
        Me.Controls.Add(Me.Button11)
        Me.Controls.Add(Me.Button10)
        Me.Controls.Add(Me.Button9)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.CLock)
        Me.Controls.Add(Me.FontColor)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.FontStype)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.BtnSave)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.FG2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.FG)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "FmCaculate_Rpt"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FmCaculate_Rpt"
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FG2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FG As System.Windows.Forms.DataGridView
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents FG2 As System.Windows.Forms.DataGridView
    Friend WithEvents BtnSave As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents FontStype As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents FontColor As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents CLock As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Button9 As System.Windows.Forms.Button
    Friend WithEvents Button10 As System.Windows.Forms.Button
    Friend WithEvents Button11 As System.Windows.Forms.Button
    Friend WithEvents Button12 As System.Windows.Forms.Button
    Friend WithEvents Button13 As System.Windows.Forms.Button
End Class
