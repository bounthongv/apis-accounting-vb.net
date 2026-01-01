<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FmTrialBalanceReport
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FmTrialBalanceReport))
        Me.RaParent = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.RY = New System.Windows.Forms.RadioButton
        Me.RD = New System.Windows.Forms.RadioButton
        Me.RP = New System.Windows.Forms.RadioButton
        Me.RM = New System.Windows.Forms.RadioButton
        Me.yy = New System.Windows.Forms.DateTimePicker
        Me.Dt = New System.Windows.Forms.DateTimePicker
        Me.Ds = New System.Windows.Forms.DateTimePicker
        Me.Pyy = New System.Windows.Forms.DateTimePicker
        Me.Myy = New System.Windows.Forms.DateTimePicker
        Me.Period = New System.Windows.Forms.ComboBox
        Me.DMonth = New System.Windows.Forms.ComboBox
        Me.Lb = New System.Windows.Forms.TextBox
        Me.L5 = New System.Windows.Forms.TextBox
        Me.FG = New System.Windows.Forms.DataGridView
        Me.BtnRefresh = New System.Windows.Forms.Button
        Me.BtnPreview = New System.Windows.Forms.Button
        Me.BtnExit = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Balance = New System.Windows.Forms.TextBox
        Me.Ac_Type = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.Ac_Nme = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.txtSumAmountDr = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtSumAmountCr = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        CType(Me.FG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'RaParent
        '
        Me.RaParent.AutoSize = True
        Me.RaParent.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.RaParent.Location = New System.Drawing.Point(687, 168)
        Me.RaParent.Name = "RaParent"
        Me.RaParent.Size = New System.Drawing.Size(128, 28)
        Me.RaParent.TabIndex = 5
        Me.RaParent.Text = "ສະເພາະບັຍຊີແມ່"
        Me.RaParent.UseVisualStyleBackColor = True
        Me.RaParent.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Saysettha OT", 15.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Blue
        Me.Label1.Location = New System.Drawing.Point(652, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(174, 34)
        Me.Label1.TabIndex = 45508
        Me.Label1.Text = "ໃບດູນດ່ຽງສຳຮອງ"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(261, 49)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(68, 24)
        Me.Label4.TabIndex = 45527
        Me.Label4.Text = "ຮອດວັນທີ"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(304, 124)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(20, 24)
        Me.Label3.TabIndex = 45526
        Me.Label3.Text = "ປີ"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(305, 85)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 24)
        Me.Label2.TabIndex = 45525
        Me.Label2.Text = "ປີ"
        '
        'RY
        '
        Me.RY.AutoSize = True
        Me.RY.Location = New System.Drawing.Point(3, 158)
        Me.RY.Name = "RY"
        Me.RY.Size = New System.Drawing.Size(74, 28)
        Me.RY.TabIndex = 45524
        Me.RY.TabStop = True
        Me.RY.Text = "ປະຈຳປີ"
        Me.RY.UseVisualStyleBackColor = True
        '
        'RD
        '
        Me.RD.AutoSize = True
        Me.RD.Location = New System.Drawing.Point(3, 46)
        Me.RD.Name = "RD"
        Me.RD.Size = New System.Drawing.Size(93, 28)
        Me.RD.TabIndex = 45523
        Me.RD.TabStop = True
        Me.RD.Text = "ປະຈຳວັນທີ"
        Me.RD.UseVisualStyleBackColor = True
        '
        'RP
        '
        Me.RP.AutoSize = True
        Me.RP.Location = New System.Drawing.Point(3, 121)
        Me.RP.Name = "RP"
        Me.RP.Size = New System.Drawing.Size(91, 28)
        Me.RP.TabIndex = 45522
        Me.RP.TabStop = True
        Me.RP.Text = "ປະຈຳງວດ"
        Me.RP.UseVisualStyleBackColor = True
        '
        'RM
        '
        Me.RM.AutoSize = True
        Me.RM.Checked = True
        Me.RM.Location = New System.Drawing.Point(4, 83)
        Me.RM.Name = "RM"
        Me.RM.Size = New System.Drawing.Size(100, 28)
        Me.RM.TabIndex = 45521
        Me.RM.TabStop = True
        Me.RM.Text = "ປະຈຳເດືອນ"
        Me.RM.UseVisualStyleBackColor = True
        '
        'yy
        '
        Me.yy.CustomFormat = "yyyy"
        Me.yy.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.yy.Location = New System.Drawing.Point(136, 157)
        Me.yy.Name = "yy"
        Me.yy.Size = New System.Drawing.Size(119, 34)
        Me.yy.TabIndex = 45520
        '
        'Dt
        '
        Me.Dt.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.Dt.Location = New System.Drawing.Point(331, 41)
        Me.Dt.Name = "Dt"
        Me.Dt.Size = New System.Drawing.Size(120, 34)
        Me.Dt.TabIndex = 45519
        '
        'Ds
        '
        Me.Ds.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.Ds.Location = New System.Drawing.Point(136, 42)
        Me.Ds.Name = "Ds"
        Me.Ds.Size = New System.Drawing.Size(118, 34)
        Me.Ds.TabIndex = 45518
        '
        'Pyy
        '
        Me.Pyy.CustomFormat = "yyyy"
        Me.Pyy.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.Pyy.Location = New System.Drawing.Point(331, 118)
        Me.Pyy.Name = "Pyy"
        Me.Pyy.Size = New System.Drawing.Size(120, 34)
        Me.Pyy.TabIndex = 45517
        '
        'Myy
        '
        Me.Myy.CustomFormat = "yyyy"
        Me.Myy.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.Myy.Location = New System.Drawing.Point(332, 79)
        Me.Myy.Name = "Myy"
        Me.Myy.Size = New System.Drawing.Size(120, 34)
        Me.Myy.TabIndex = 45516
        '
        'Period
        '
        Me.Period.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Period.FormattingEnabled = True
        Me.Period.Items.AddRange(New Object() {"ງວດທີ 1", "ງວດທີ 2", "ງວດທີ 3", "ງວດທີ 4"})
        Me.Period.Location = New System.Drawing.Point(136, 119)
        Me.Period.Name = "Period"
        Me.Period.Size = New System.Drawing.Size(118, 32)
        Me.Period.TabIndex = 45515
        '
        'DMonth
        '
        Me.DMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.DMonth.FormattingEnabled = True
        Me.DMonth.Items.AddRange(New Object() {"ມັງກອນ", "ກຸມພາ", "ມີນາ", "ເມສາ", "ພຶດສະພາ", "ມີຖຸນາ", "ກໍລະກົດ", "ສິງຫາ", "ກັນຍາ", "ຕຸລາ", "ພະຈິກ", "ທັນວາ"})
        Me.DMonth.Location = New System.Drawing.Point(136, 81)
        Me.DMonth.Name = "DMonth"
        Me.DMonth.Size = New System.Drawing.Size(118, 32)
        Me.DMonth.TabIndex = 45514
        '
        'Lb
        '
        Me.Lb.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(221, Byte), Integer), CType(CType(254, Byte), Integer))
        Me.Lb.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Lb.Location = New System.Drawing.Point(460, 77)
        Me.Lb.Name = "Lb"
        Me.Lb.Size = New System.Drawing.Size(552, 34)
        Me.Lb.TabIndex = 45530
        Me.Lb.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'L5
        '
        Me.L5.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(221, Byte), Integer), CType(CType(254, Byte), Integer))
        Me.L5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.L5.Location = New System.Drawing.Point(460, 115)
        Me.L5.Name = "L5"
        Me.L5.Size = New System.Drawing.Size(552, 34)
        Me.L5.TabIndex = 45531
        Me.L5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'FG
        '
        Me.FG.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FG.DataSource = Nothing
        Me.FG.Location = New System.Drawing.Point(6, 210)
        Me.FG.Name = "FG"
        Me.FG.Size = New System.Drawing.Size(1006, 514)
        Me.FG.TabIndex = 45532
        '
        'BtnRefresh
        '
        Me.BtnRefresh.Image = CType(resources.GetObject("BtnRefresh.Image"), System.Drawing.Image)
        Me.BtnRefresh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnRefresh.Location = New System.Drawing.Point(136, 3)
        Me.BtnRefresh.Name = "BtnRefresh"
        Me.BtnRefresh.Size = New System.Drawing.Size(100, 35)
        Me.BtnRefresh.TabIndex = 45534
        Me.BtnRefresh.Text = "ເອີ້ນຂໍ້ມູນ"
        Me.BtnRefresh.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnRefresh.UseVisualStyleBackColor = True
        '
        'BtnPreview
        '
        Me.BtnPreview.Image = CType(resources.GetObject("BtnPreview.Image"), System.Drawing.Image)
        Me.BtnPreview.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnPreview.Location = New System.Drawing.Point(36, 3)
        Me.BtnPreview.Name = "BtnPreview"
        Me.BtnPreview.Size = New System.Drawing.Size(100, 35)
        Me.BtnPreview.TabIndex = 45512
        Me.BtnPreview.Text = "ວິວ/ເບິ່ງ"
        Me.BtnPreview.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnPreview.UseVisualStyleBackColor = True
        '
        'BtnExit
        '
        Me.BtnExit.Image = Global.ApPBank10.My.Resources.Resources._Exit
        Me.BtnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnExit.Location = New System.Drawing.Point(3, 4)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(35, 35)
        Me.BtnExit.TabIndex = 45535
        Me.BtnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Balance)
        Me.Panel1.Controls.Add(Me.Ac_Type)
        Me.Panel1.Controls.Add(Me.Label18)
        Me.Panel1.Controls.Add(Me.Ac_Nme)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label16)
        Me.Panel1.Controls.Add(Me.txtSumAmountDr)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.txtSumAmountCr)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Location = New System.Drawing.Point(12, 621)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1007, 103)
        Me.Panel1.TabIndex = 45551
        Me.Panel1.Visible = False
        '
        'Balance
        '
        Me.Balance.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Balance.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Balance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Balance.Location = New System.Drawing.Point(830, 66)
        Me.Balance.Name = "Balance"
        Me.Balance.ReadOnly = True
        Me.Balance.Size = New System.Drawing.Size(160, 30)
        Me.Balance.TabIndex = 45548
        Me.Balance.Text = "0.00"
        Me.Balance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Ac_Type
        '
        Me.Ac_Type.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Ac_Type.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Ac_Type.ForeColor = System.Drawing.Color.Black
        Me.Ac_Type.Location = New System.Drawing.Point(94, 36)
        Me.Ac_Type.Name = "Ac_Type"
        Me.Ac_Type.ReadOnly = True
        Me.Ac_Type.Size = New System.Drawing.Size(610, 30)
        Me.Ac_Type.TabIndex = 45543
        Me.Ac_Type.Text = "0.00"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label18.Location = New System.Drawing.Point(8, 40)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(97, 24)
        Me.Label18.TabIndex = 45548
        Me.Label18.Text = "ປະເພດບັນຊື :" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Ac_Nme
        '
        Me.Ac_Nme.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Ac_Nme.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Ac_Nme.ForeColor = System.Drawing.Color.Black
        Me.Ac_Nme.Location = New System.Drawing.Point(94, 4)
        Me.Ac_Nme.Name = "Ac_Nme"
        Me.Ac_Nme.ReadOnly = True
        Me.Ac_Nme.Size = New System.Drawing.Size(610, 30)
        Me.Ac_Nme.TabIndex = 45545
        Me.Ac_Nme.Text = "0.00"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label5.ForeColor = System.Drawing.Color.Red
        Me.Label5.Location = New System.Drawing.Point(752, 72)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(84, 24)
        Me.Label5.TabIndex = 45549
        Me.Label5.Text = "ຄ່າຜິດດ່ຽງ :"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label16.Location = New System.Drawing.Point(38, 8)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(62, 24)
        Me.Label16.TabIndex = 45549
        Me.Label16.Text = "ຊື່ບັນຊື :" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'txtSumAmountDr
        '
        Me.txtSumAmountDr.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtSumAmountDr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSumAmountDr.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtSumAmountDr.Location = New System.Drawing.Point(830, 3)
        Me.txtSumAmountDr.Name = "txtSumAmountDr"
        Me.txtSumAmountDr.ReadOnly = True
        Me.txtSumAmountDr.Size = New System.Drawing.Size(160, 30)
        Me.txtSumAmountDr.TabIndex = 45544
        Me.txtSumAmountDr.Text = "0.00"
        Me.txtSumAmountDr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label11.Location = New System.Drawing.Point(713, 7)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(131, 24)
        Me.Label11.TabIndex = 45547
        Me.Label11.Text = "ຈຳນວນເງິນຈົດຫນື້ :"
        '
        'txtSumAmountCr
        '
        Me.txtSumAmountCr.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtSumAmountCr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSumAmountCr.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtSumAmountCr.Location = New System.Drawing.Point(830, 35)
        Me.txtSumAmountCr.Name = "txtSumAmountCr"
        Me.txtSumAmountCr.ReadOnly = True
        Me.txtSumAmountCr.Size = New System.Drawing.Size(160, 30)
        Me.txtSumAmountCr.TabIndex = 45542
        Me.txtSumAmountCr.Text = "0.00"
        Me.txtSumAmountCr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label12.Location = New System.Drawing.Point(723, 38)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(120, 24)
        Me.Label12.TabIndex = 45546
        Me.Label12.Text = "ຈຳນວນເງິນຈົດມີ :"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(331, 168)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(228, 28)
        Me.CheckBox1.TabIndex = 45552
        Me.CheckBox1.Text = "ພາຍຫຼັງສ້າງໃບລາຍງານຜົນໄດ້ຮັບ"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'FmTrialBalanceReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(221, Byte), Integer), CType(CType(254, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1016, 736)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.RaParent)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.BtnRefresh)
        Me.Controls.Add(Me.FG)
        Me.Controls.Add(Me.L5)
        Me.Controls.Add(Me.Lb)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Dt)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.RY)
        Me.Controls.Add(Me.RD)
        Me.Controls.Add(Me.RP)
        Me.Controls.Add(Me.RM)
        Me.Controls.Add(Me.yy)
        Me.Controls.Add(Me.Ds)
        Me.Controls.Add(Me.Pyy)
        Me.Controls.Add(Me.Myy)
        Me.Controls.Add(Me.Period)
        Me.Controls.Add(Me.DMonth)
        Me.Controls.Add(Me.BtnPreview)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 6, 4, 6)
        Me.Name = "FmTrialBalanceReport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FmTrialBalanceReport"
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents BtnPreview As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents RY As System.Windows.Forms.RadioButton
    Friend WithEvents RD As System.Windows.Forms.RadioButton
    Friend WithEvents RP As System.Windows.Forms.RadioButton
    Friend WithEvents RM As System.Windows.Forms.RadioButton
    Friend WithEvents yy As System.Windows.Forms.DateTimePicker
    Friend WithEvents Dt As System.Windows.Forms.DateTimePicker
    Friend WithEvents Ds As System.Windows.Forms.DateTimePicker
    Friend WithEvents Pyy As System.Windows.Forms.DateTimePicker
    Friend WithEvents Myy As System.Windows.Forms.DateTimePicker
    Friend WithEvents Period As System.Windows.Forms.ComboBox
    Friend WithEvents DMonth As System.Windows.Forms.ComboBox
    Friend WithEvents Lb As System.Windows.Forms.TextBox
    Friend WithEvents L5 As System.Windows.Forms.TextBox
    Friend WithEvents FG As System.Windows.Forms.DataGridView
    Friend WithEvents BtnRefresh As System.Windows.Forms.Button
    Friend WithEvents RaParent As System.Windows.Forms.CheckBox
    Friend WithEvents BtnExit As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Balance As System.Windows.Forms.TextBox
    Friend WithEvents Ac_Type As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Ac_Nme As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtSumAmountDr As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtSumAmountCr As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
End Class
