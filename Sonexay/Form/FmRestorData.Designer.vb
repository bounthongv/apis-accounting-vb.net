<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FmRestorData
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FmRestorData))
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.cmdrestore = New System.Windows.Forms.Button
        Me.cmdsearch = New System.Windows.Forms.Button
        Me.txtSaveIn = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.FG = New System.Windows.Forms.DataGridView
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.dts = New System.Windows.Forms.DateTimePicker
        Me.dtt = New System.Windows.Forms.DateTimePicker
        Me.BtnExit = New System.Windows.Forms.Button
        CType(Me.FG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Saysettha OT", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(76, 4)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(204, 28)
        Me.TextBox1.TabIndex = 47
        '
        'cmdrestore
        '
        Me.cmdrestore.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdrestore.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdrestore.ForeColor = System.Drawing.Color.Navy
        Me.cmdrestore.Location = New System.Drawing.Point(6, 391)
        Me.cmdrestore.Name = "cmdrestore"
        Me.cmdrestore.Size = New System.Drawing.Size(155, 29)
        Me.cmdrestore.TabIndex = 46
        Me.cmdrestore.Text = "ເອີ້ນຂໍ້ມູນກັບຄືນມາ"
        Me.cmdrestore.UseVisualStyleBackColor = False
        '
        'cmdsearch
        '
        Me.cmdsearch.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdsearch.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdsearch.ForeColor = System.Drawing.Color.Navy
        Me.cmdsearch.Location = New System.Drawing.Point(361, 35)
        Me.cmdsearch.Name = "cmdsearch"
        Me.cmdsearch.Size = New System.Drawing.Size(139, 33)
        Me.cmdsearch.TabIndex = 45
        Me.cmdsearch.Text = "ຊອກຫາຕາມແພັມ"
        Me.cmdsearch.UseVisualStyleBackColor = False
        '
        'txtSaveIn
        '
        Me.txtSaveIn.BackColor = System.Drawing.Color.White
        Me.txtSaveIn.Font = New System.Drawing.Font("Saysettha OT", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSaveIn.ForeColor = System.Drawing.Color.Blue
        Me.txtSaveIn.Location = New System.Drawing.Point(76, 36)
        Me.txtSaveIn.Name = "txtSaveIn"
        Me.txtSaveIn.ReadOnly = True
        Me.txtSaveIn.Size = New System.Drawing.Size(285, 28)
        Me.txtSaveIn.TabIndex = 44
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(-2, 40)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(80, 21)
        Me.Label5.TabIndex = 43
        Me.Label5.Text = "ບ່ອນຈັດເກັບ"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(51, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(19, 21)
        Me.Label1.TabIndex = 210
        Me.Label1.Text = "ຊື່"
        '
        'FG
        '
        Me.FG.DataSource = Nothing
        Me.FG.Location = New System.Drawing.Point(5, 104)
        Me.FG.Name = "FG"

        Me.FG.Size = New System.Drawing.Size(497, 285)
        Me.FG.TabIndex = 211
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton1.Location = New System.Drawing.Point(286, 6)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(75, 25)
        Me.RadioButton1.TabIndex = 213
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "ຈາກແຟັມ"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton2.Location = New System.Drawing.Point(361, 7)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(106, 25)
        Me.RadioButton2.TabIndex = 214
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "ເລືອກຕາມວັນທີ"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.Button1)
        Me.Panel3.Controls.Add(Me.Label3)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Controls.Add(Me.dts)
        Me.Panel3.Controls.Add(Me.dtt)
        Me.Panel3.Location = New System.Drawing.Point(5, 67)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(499, 31)
        Me.Panel3.TabIndex = 210
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Button1.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Navy
        Me.Button1.Location = New System.Drawing.Point(355, 1)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(139, 27)
        Me.Button1.TabIndex = 215
        Me.Button1.Text = "ຊອກຫາຕາມວັນທີ"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(185, 4)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 21)
        Me.Label3.TabIndex = 216
        Me.Label3.Text = "ຫາວັນທີ"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(7, 4)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 21)
        Me.Label2.TabIndex = 215
        Me.Label2.Text = "ແຕ່ວັນທີ"
        '
        'dts
        '
        Me.dts.CustomFormat = "dd/MM/yyyy"
        Me.dts.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dts.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dts.Location = New System.Drawing.Point(70, 1)
        Me.dts.Name = "dts"
        Me.dts.Size = New System.Drawing.Size(101, 26)
        Me.dts.TabIndex = 168
        '
        'dtt
        '
        Me.dtt.CustomFormat = "dd/MM/yyyy"
        Me.dtt.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtt.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtt.Location = New System.Drawing.Point(247, -1)
        Me.dtt.Name = "dtt"
        Me.dtt.Size = New System.Drawing.Size(106, 26)
        Me.dtt.TabIndex = 169
        '
        'BtnExit
        '
        Me.BtnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnExit.Location = New System.Drawing.Point(2, 3)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(35, 35)
        Me.BtnExit.TabIndex = 292
        Me.BtnExit.Tag = "9999"
        Me.BtnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'FmRestorData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(508, 422)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.RadioButton2)
        Me.Controls.Add(Me.RadioButton1)
        Me.Controls.Add(Me.FG)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.cmdrestore)
        Me.Controls.Add(Me.cmdsearch)
        Me.Controls.Add(Me.txtSaveIn)
        Me.Controls.Add(Me.Label5)
        Me.Name = "FmRestorData"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FmRestorData"
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents cmdrestore As System.Windows.Forms.Button
    Friend WithEvents cmdsearch As System.Windows.Forms.Button
    Friend WithEvents txtSaveIn As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents FG As System.Windows.Forms.DataGridView
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents dts As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtt As System.Windows.Forms.DateTimePicker
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents BtnExit As System.Windows.Forms.Button
End Class
