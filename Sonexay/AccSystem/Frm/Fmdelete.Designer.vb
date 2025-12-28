<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Fmdelete
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
        Me.BtnDelete = New System.Windows.Forms.Button
        Me.Cmdelete = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Dt = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.RY = New System.Windows.Forms.RadioButton
        Me.RD = New System.Windows.Forms.RadioButton
        Me.RM = New System.Windows.Forms.RadioButton
        Me.yy = New System.Windows.Forms.DateTimePicker
        Me.Ds = New System.Windows.Forms.DateTimePicker
        Me.Myy = New System.Windows.Forms.DateTimePicker
        Me.DMonth = New System.Windows.Forms.ComboBox
        Me.L5 = New System.Windows.Forms.TextBox
        Me.Lb = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'BtnDelete
        '
        Me.BtnDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnDelete.Location = New System.Drawing.Point(37, 3)
        Me.BtnDelete.Name = "BtnDelete"
        Me.BtnDelete.Size = New System.Drawing.Size(100, 35)
        Me.BtnDelete.TabIndex = 21
        Me.BtnDelete.Text = "ລຶບ"
        Me.BtnDelete.UseVisualStyleBackColor = True
        '
        'Cmdelete
        '
        Me.Cmdelete.FormattingEnabled = True
        Me.Cmdelete.Items.AddRange(New Object() {"ລຶບລາຍການເຄື່ອນໄຫວບັນຊີປະຈຳວັນ", "ລຶບລາຍການຍອດຍົກປະຈຳປີ"})
        Me.Cmdelete.Location = New System.Drawing.Point(101, 237)
        Me.Cmdelete.Name = "Cmdelete"
        Me.Cmdelete.Size = New System.Drawing.Size(256, 29)
        Me.Cmdelete.TabIndex = 174
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(198, 145)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(60, 21)
        Me.Label4.TabIndex = 45541
        Me.Label4.Text = "ຮອດວັນທີ"
        '
        'Dt
        '
        Me.Dt.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.Dt.Location = New System.Drawing.Point(262, 139)
        Me.Dt.Name = "Dt"
        Me.Dt.Size = New System.Drawing.Size(95, 30)
        Me.Dt.TabIndex = 45533
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(240, 176)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(18, 21)
        Me.Label2.TabIndex = 45539
        Me.Label2.Text = "ປີ"
        '
        'RY
        '
        Me.RY.AutoSize = True
        Me.RY.Location = New System.Drawing.Point(9, 207)
        Me.RY.Name = "RY"
        Me.RY.Size = New System.Drawing.Size(66, 25)
        Me.RY.TabIndex = 45538
        Me.RY.TabStop = True
        Me.RY.Text = "ປະຈຳປີ"
        Me.RY.UseVisualStyleBackColor = True
        '
        'RD
        '
        Me.RD.AutoSize = True
        Me.RD.Location = New System.Drawing.Point(9, 142)
        Me.RD.Name = "RD"
        Me.RD.Size = New System.Drawing.Size(83, 25)
        Me.RD.TabIndex = 45537
        Me.RD.TabStop = True
        Me.RD.Text = "ປະຈຳວັນທີ"
        Me.RD.UseVisualStyleBackColor = True
        '
        'RM
        '
        Me.RM.AutoSize = True
        Me.RM.Checked = True
        Me.RM.Location = New System.Drawing.Point(9, 174)
        Me.RM.Name = "RM"
        Me.RM.Size = New System.Drawing.Size(88, 25)
        Me.RM.TabIndex = 45535
        Me.RM.TabStop = True
        Me.RM.Text = "ປະຈຳເດືອນ"
        Me.RM.UseVisualStyleBackColor = True
        '
        'yy
        '
        Me.yy.CustomFormat = "yyyy"
        Me.yy.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.yy.Location = New System.Drawing.Point(101, 204)
        Me.yy.Name = "yy"
        Me.yy.Size = New System.Drawing.Size(94, 30)
        Me.yy.TabIndex = 45534
        '
        'Ds
        '
        Me.Ds.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.Ds.Location = New System.Drawing.Point(101, 139)
        Me.Ds.Name = "Ds"
        Me.Ds.Size = New System.Drawing.Size(94, 30)
        Me.Ds.TabIndex = 45532
        '
        'Myy
        '
        Me.Myy.CustomFormat = "yyyy"
        Me.Myy.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.Myy.Location = New System.Drawing.Point(262, 172)
        Me.Myy.Name = "Myy"
        Me.Myy.Size = New System.Drawing.Size(95, 30)
        Me.Myy.TabIndex = 45530
        '
        'DMonth
        '
        Me.DMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.DMonth.FormattingEnabled = True
        Me.DMonth.Items.AddRange(New Object() {"ມັງກອນ", "ກຸມພາ", "ມີນາ", "ເມສາ", "ພຶດສະພາ", "ມີຖຸນາ", "ກໍລະກົດ", "ສິງຫາ", "ກັນຍາ", "ຕຸລາ", "ພະຈິກ", "ທັນວາ"})
        Me.DMonth.Location = New System.Drawing.Point(101, 172)
        Me.DMonth.Name = "DMonth"
        Me.DMonth.Size = New System.Drawing.Size(94, 29)
        Me.DMonth.TabIndex = 45528
        '
        'L5
        '
        Me.L5.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(221, Byte), Integer), CType(CType(254, Byte), Integer))
        Me.L5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.L5.Location = New System.Drawing.Point(102, 270)
        Me.L5.Name = "L5"
        Me.L5.Size = New System.Drawing.Size(256, 30)
        Me.L5.TabIndex = 45543
        Me.L5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Lb
        '
        Me.Lb.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(221, Byte), Integer), CType(CType(254, Byte), Integer))
        Me.Lb.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Lb.Location = New System.Drawing.Point(102, -31)
        Me.Lb.Name = "Lb"
        Me.Lb.Size = New System.Drawing.Size(308, 30)
        Me.Lb.TabIndex = 45542
        Me.Lb.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 241)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 21)
        Me.Label1.TabIndex = 45544
        Me.Label1.Text = "ລາຍການທີຈະລືບ"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Saysettha OT", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(18, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(327, 34)
        Me.Label3.TabIndex = 45545
        Me.Label3.Text = "ກະລຸນນາເລືອກລາຍການທີຈະລືບກອນ"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Location = New System.Drawing.Point(5, 64)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(352, 71)
        Me.GroupBox1.TabIndex = 45546
        Me.GroupBox1.TabStop = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.ForeColor = System.Drawing.Color.Red
        Me.Label7.Location = New System.Drawing.Point(66, 44)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(245, 21)
        Me.Label7.TabIndex = 45544
        Me.Label7.Text = "ແລະຂໍ້ມູນນັ້ນຈະຖືກລຶບໂດຍບໍ່ສາມາດເອື້ນຄືນໄດ້"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.Color.Red
        Me.Label6.Location = New System.Drawing.Point(55, 18)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(294, 21)
        Me.Label6.TabIndex = 45543
        Me.Label6.Text = ": ການລືບຂໍ້ມູນໃນທີ່ນີ້ແມ່ນຈະລຶບຂໍ້ມູນທັງໝົດທີທ່ານເລືອກ"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Red
        Me.Label5.Location = New System.Drawing.Point(4, 18)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(55, 21)
        Me.Label5.TabIndex = 45542
        Me.Label5.Text = "ຄຳເຕືອນ"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.ForeColor = System.Drawing.Color.Red
        Me.Label8.Location = New System.Drawing.Point(-1, 272)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(97, 21)
        Me.Label8.TabIndex = 45545
        Me.Label8.Text = "ລຶບຢູ່ໃນລະຫວ່າງ"
        '
        'Button1
        '
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(3, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(35, 35)
        Me.Button1.TabIndex = 45547
        Me.Button1.Tag = "9999"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Fmdelete
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 21.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(221, Byte), Integer), CType(CType(254, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(366, 311)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.L5)
        Me.Controls.Add(Me.Lb)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Dt)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.RY)
        Me.Controls.Add(Me.RD)
        Me.Controls.Add(Me.RM)
        Me.Controls.Add(Me.yy)
        Me.Controls.Add(Me.Ds)
        Me.Controls.Add(Me.Myy)
        Me.Controls.Add(Me.DMonth)
        Me.Controls.Add(Me.Cmdelete)
        Me.Controls.Add(Me.BtnDelete)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Fmdelete"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Fmdelete"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnDelete As System.Windows.Forms.Button
    Friend WithEvents Cmdelete As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Dt As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents RY As System.Windows.Forms.RadioButton
    Friend WithEvents RD As System.Windows.Forms.RadioButton
    Friend WithEvents RM As System.Windows.Forms.RadioButton
    Friend WithEvents yy As System.Windows.Forms.DateTimePicker
    Friend WithEvents Ds As System.Windows.Forms.DateTimePicker
    Friend WithEvents Myy As System.Windows.Forms.DateTimePicker
    Friend WithEvents DMonth As System.Windows.Forms.ComboBox
    Friend WithEvents L5 As System.Windows.Forms.TextBox
    Friend WithEvents Lb As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
