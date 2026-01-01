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
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.ReCr = New System.Windows.Forms.TextBox
        Me.OpDr = New System.Windows.Forms.TextBox
        Me.BOpDr = New System.Windows.Forms.TextBox
        Me.OpCr = New System.Windows.Forms.TextBox
        Me.AmtDr = New System.Windows.Forms.TextBox
        Me.BAmtDr = New System.Windows.Forms.TextBox
        Me.AmtCr = New System.Windows.Forms.TextBox
        Me.ReDr = New System.Windows.Forms.TextBox
        Me.BReDr = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.BalanceType = New System.Windows.Forms.ComboBox
        Me.BtnExit = New System.Windows.Forms.Button
        Me.BtnRefresh = New System.Windows.Forms.Button
        Me.Cx = New System.Windows.Forms.ComboBox
        Me.Off_Usr = New System.Windows.Forms.ComboBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.CheckBnk = New System.Windows.Forms.CheckBox
        Me.RGL = New System.Windows.Forms.RadioButton
        Me.RDtail = New System.Windows.Forms.RadioButton
        Me.RGroup = New System.Windows.Forms.RadioButton
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.RT = New System.Windows.Forms.RadioButton
        Me.yyt = New System.Windows.Forms.DateTimePicker
        Me.Ct = New System.Windows.Forms.ComboBox
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Button4 = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.CMB_Curr = New System.Windows.Forms.ComboBox
        Me.txtRate = New System.Windows.Forms.TextBox
        Me.txtRate2 = New System.Windows.Forms.TextBox
        Me.txtcurr_name2 = New System.Windows.Forms.TextBox
        Me.CheckBox4 = New System.Windows.Forms.CheckBox
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.TxtHeader = New System.Windows.Forms.TextBox
        Me.TxtS1 = New System.Windows.Forms.TextBox
        Me.TxtS2 = New System.Windows.Forms.TextBox
        Me.TxtS3 = New System.Windows.Forms.TextBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.TxtS4 = New System.Windows.Forms.TextBox
        Me.TxtPP = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.RD = New System.Windows.Forms.RadioButton
        Me.RM = New System.Windows.Forms.RadioButton
        Me.RY = New System.Windows.Forms.RadioButton
        Me.RP = New System.Windows.Forms.RadioButton
        Me.Label27 = New System.Windows.Forms.Label
        Me.CheckBox3 = New System.Windows.Forms.CheckBox
        Me.nn = New System.Windows.Forms.Label
        Me.BtnPreview = New System.Windows.Forms.Button
        CType(Me.FG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'RaParent
        '
        Me.RaParent.AutoSize = True
        Me.RaParent.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.RaParent.Location = New System.Drawing.Point(349, -23)
        Me.RaParent.Name = "RaParent"
        Me.RaParent.Size = New System.Drawing.Size(130, 31)
        Me.RaParent.TabIndex = 5
        Me.RaParent.Text = "ສະເພາະບັຍຊີແມ່"
        Me.RaParent.UseVisualStyleBackColor = True
        Me.RaParent.Visible = False
        '
        'yy
        '
        Me.yy.CustomFormat = "yyyy"
        Me.yy.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.yy.Location = New System.Drawing.Point(135, 193)
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
        Me.Period.FormattingEnabled = True
        Me.Period.Items.AddRange(New Object() {"ງວດທີ 1", "ງວດທີ 2", "ງວດທີ 3", "ງວດທີ 4"})
        Me.Period.Location = New System.Drawing.Point(136, 119)
        Me.Period.Name = "Period"
        Me.Period.Size = New System.Drawing.Size(118, 35)
        Me.Period.TabIndex = 45515
        '
        'DMonth
        '
        Me.DMonth.FormattingEnabled = True
        Me.DMonth.Items.AddRange(New Object() {"ມັງກອນ", "ກຸມພາ", "ມີນາ", "ເມສາ", "ພຶດສະພາ", "ມີຖຸນາ", "ກໍລະກົດ", "ສິງຫາ", "ກັນຍາ", "ຕຸລາ", "ພະຈິກ", "ທັນວາ"})
        Me.DMonth.Location = New System.Drawing.Point(136, 81)
        Me.DMonth.Name = "DMonth"
        Me.DMonth.Size = New System.Drawing.Size(118, 35)
        Me.DMonth.TabIndex = 45514
        '
        'Lb
        '
        Me.Lb.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(221, Byte), Integer), CType(CType(254, Byte), Integer))
        Me.Lb.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Lb.Location = New System.Drawing.Point(460, 43)
        Me.Lb.Name = "Lb"
        Me.Lb.Size = New System.Drawing.Size(552, 34)
        Me.Lb.TabIndex = 45530
        Me.Lb.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'L5
        '
        Me.L5.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(221, Byte), Integer), CType(CType(254, Byte), Integer))
        Me.L5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.L5.Location = New System.Drawing.Point(460, 81)
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
        Me.FG.Location = New System.Drawing.Point(6, 230)
        Me.FG.Name = "FG"

        Me.FG.Size = New System.Drawing.Size(1301, 386)
        Me.FG.TabIndex = 45532
        Me.FG.Tag = "8006"
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.Label17)
        Me.Panel1.Controls.Add(Me.ReCr)
        Me.Panel1.Controls.Add(Me.OpDr)
        Me.Panel1.Controls.Add(Me.BOpDr)
        Me.Panel1.Controls.Add(Me.OpCr)
        Me.Panel1.Controls.Add(Me.AmtDr)
        Me.Panel1.Controls.Add(Me.BAmtDr)
        Me.Panel1.Controls.Add(Me.AmtCr)
        Me.Panel1.Controls.Add(Me.ReDr)
        Me.Panel1.Controls.Add(Me.BReDr)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label16)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Location = New System.Drawing.Point(6, 630)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1309, 92)
        Me.Panel1.TabIndex = 45551
        Me.Panel1.Tag = "1"
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label11.Location = New System.Drawing.Point(652, 59)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(157, 21)
        Me.Label11.TabIndex = 45568
        Me.Label11.Tag = "2064"
        Me.Label11.Text = "ຄ່າຜິດດ່ຽງຍອດເຫລືອ:"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.Location = New System.Drawing.Point(853, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(81, 24)
        Me.Label17.TabIndex = 45566
        Me.Label17.Tag = "2061"
        Me.Label17.Text = "ຍອດເຫລືອ(ມີ)"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ReCr
        '
        Me.ReCr.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ReCr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ReCr.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.ReCr.Location = New System.Drawing.Point(809, 23)
        Me.ReCr.Name = "ReCr"
        Me.ReCr.ReadOnly = True
        Me.ReCr.Size = New System.Drawing.Size(160, 30)
        Me.ReCr.TabIndex = 45554
        Me.ReCr.Text = "0.00"
        Me.ReCr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'OpDr
        '
        Me.OpDr.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.OpDr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OpDr.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.OpDr.Location = New System.Drawing.Point(3, 23)
        Me.OpDr.Name = "OpDr"
        Me.OpDr.ReadOnly = True
        Me.OpDr.Size = New System.Drawing.Size(160, 30)
        Me.OpDr.TabIndex = 45552
        Me.OpDr.Text = "0.00"
        Me.OpDr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'BOpDr
        '
        Me.BOpDr.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.BOpDr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BOpDr.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.BOpDr.Location = New System.Drawing.Point(164, 56)
        Me.BOpDr.Name = "BOpDr"
        Me.BOpDr.ReadOnly = True
        Me.BOpDr.Size = New System.Drawing.Size(160, 30)
        Me.BOpDr.TabIndex = 45551
        Me.BOpDr.Text = "0.00"
        Me.BOpDr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'OpCr
        '
        Me.OpCr.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.OpCr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OpCr.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.OpCr.Location = New System.Drawing.Point(164, 23)
        Me.OpCr.Name = "OpCr"
        Me.OpCr.ReadOnly = True
        Me.OpCr.Size = New System.Drawing.Size(160, 30)
        Me.OpCr.TabIndex = 45550
        Me.OpCr.Text = "0.00"
        Me.OpCr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'AmtDr
        '
        Me.AmtDr.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.AmtDr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AmtDr.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.AmtDr.Location = New System.Drawing.Point(325, 23)
        Me.AmtDr.Name = "AmtDr"
        Me.AmtDr.ReadOnly = True
        Me.AmtDr.Size = New System.Drawing.Size(160, 30)
        Me.AmtDr.TabIndex = 45548
        Me.AmtDr.Text = "0.00"
        Me.AmtDr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'BAmtDr
        '
        Me.BAmtDr.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.BAmtDr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BAmtDr.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.BAmtDr.Location = New System.Drawing.Point(486, 56)
        Me.BAmtDr.Name = "BAmtDr"
        Me.BAmtDr.ReadOnly = True
        Me.BAmtDr.Size = New System.Drawing.Size(160, 30)
        Me.BAmtDr.TabIndex = 45547
        Me.BAmtDr.Text = "0.00"
        Me.BAmtDr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'AmtCr
        '
        Me.AmtCr.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.AmtCr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AmtCr.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.AmtCr.Location = New System.Drawing.Point(486, 23)
        Me.AmtCr.Name = "AmtCr"
        Me.AmtCr.ReadOnly = True
        Me.AmtCr.Size = New System.Drawing.Size(160, 30)
        Me.AmtCr.TabIndex = 45546
        Me.AmtCr.Text = "0.00"
        Me.AmtCr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'ReDr
        '
        Me.ReDr.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ReDr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ReDr.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.ReDr.Location = New System.Drawing.Point(647, 23)
        Me.ReDr.Name = "ReDr"
        Me.ReDr.ReadOnly = True
        Me.ReDr.Size = New System.Drawing.Size(160, 30)
        Me.ReDr.TabIndex = 45544
        Me.ReDr.Text = "0.00"
        Me.ReDr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'BReDr
        '
        Me.BReDr.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.BReDr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BReDr.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.BReDr.Location = New System.Drawing.Point(809, 56)
        Me.BReDr.Name = "BReDr"
        Me.BReDr.ReadOnly = True
        Me.BReDr.Size = New System.Drawing.Size(160, 30)
        Me.BReDr.TabIndex = 45542
        Me.BReDr.Text = "0.00"
        Me.BReDr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(686, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(91, 24)
        Me.Label9.TabIndex = 45558
        Me.Label9.Tag = "2060"
        Me.Label9.Text = "ຍອດເຫລືອ(ຫນີ້)"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(501, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(123, 24)
        Me.Label8.TabIndex = 45557
        Me.Label8.Tag = "2059"
        Me.Label8.Text = "ເຄື່ອນໄຫວໃນເດືອນ(ມີ)"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(337, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(133, 24)
        Me.Label7.TabIndex = 45556
        Me.Label7.Tag = "2058"
        Me.Label7.Text = "ເຄື່ອນໄຫວໃນເດືອນ(ຫນີ)"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(179, -2)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(111, 24)
        Me.Label6.TabIndex = 45555
        Me.Label6.Tag = "2057"
        Me.Label6.Text = "ຍອດຍົກເບື້ອງຕົ້ນ(ມີ)"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(21, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(121, 24)
        Me.Label5.TabIndex = 45553
        Me.Label5.Tag = "2056"
        Me.Label5.Text = "ຍອດຍົກເບື້ອງຕົ້ນ(ຫນີ)"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label16.Location = New System.Drawing.Point(9, 60)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(154, 21)
        Me.Label16.TabIndex = 45560
        Me.Label16.Tag = "2062"
        Me.Label16.Text = "ຄ່າຜິດດ່ຽງຍອດຍົກເບື້ອງຕົ້ນ:"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label10.Location = New System.Drawing.Point(321, 59)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(166, 21)
        Me.Label10.TabIndex = 45567
        Me.Label10.Tag = "2063"
        Me.Label10.Text = "ຄ່າຜິດດ່ຽງເຄື່ອນໄຫວໃນເດືອນ:"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(462, 159)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(57, 27)
        Me.Label12.TabIndex = 45557
        Me.Label12.Tag = "2055"
        Me.Label12.Text = "ປະເພດ"
        '
        'BalanceType
        '
        Me.BalanceType.FormattingEnabled = True
        Me.BalanceType.Items.AddRange(New Object() {"ໃບດູນດ່ຽງທົ່ວໄປ", "ໃບດູນດ່ຽງລາຍຮັບ-ລາຍຈ່າຍ"})
        Me.BalanceType.Location = New System.Drawing.Point(523, 157)
        Me.BalanceType.Name = "BalanceType"
        Me.BalanceType.Size = New System.Drawing.Size(168, 35)
        Me.BalanceType.TabIndex = 45556
        '
        'BtnExit
        '
        Me.BtnExit.Image = Global.ApPBank10.My.Resources.Resources.Exit1
        Me.BtnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnExit.Location = New System.Drawing.Point(3, 4)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(35, 35)
        Me.BtnExit.TabIndex = 45535
        Me.BtnExit.Tag = "9999"
        Me.BtnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'BtnRefresh
        '
        Me.BtnRefresh.Image = Global.ApPBank10.My.Resources.Resources.Refresh
        Me.BtnRefresh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnRefresh.Location = New System.Drawing.Point(150, 3)
        Me.BtnRefresh.Name = "BtnRefresh"
        Me.BtnRefresh.Size = New System.Drawing.Size(109, 35)
        Me.BtnRefresh.TabIndex = 45534
        Me.BtnRefresh.Tag = "1018"
        Me.BtnRefresh.Text = "ເອີ້ນຂໍ້ມູນ"
        Me.BtnRefresh.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnRefresh.UseVisualStyleBackColor = True
        '
        'Cx
        '
        Me.Cx.FormattingEnabled = True
        Me.Cx.Items.AddRange(New Object() {" ທັງໝົດ", " 3 ຕົວ", " 4 ຕົວ", " 5 ຕົວ", " 6 ຕົວ", " 7 ຕົວ"})
        Me.Cx.Location = New System.Drawing.Point(330, 222)
        Me.Cx.Name = "Cx"
        Me.Cx.Size = New System.Drawing.Size(120, 35)
        Me.Cx.TabIndex = 45559
        Me.Cx.Visible = False
        '
        'Off_Usr
        '
        Me.Off_Usr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Off_Usr.FormattingEnabled = True
        Me.Off_Usr.Location = New System.Drawing.Point(523, 124)
        Me.Off_Usr.Name = "Off_Usr"
        Me.Off_Usr.Size = New System.Drawing.Size(168, 32)
        Me.Off_Usr.TabIndex = 45561
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(932, 124)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(201, 31)
        Me.CheckBox2.TabIndex = 45563
        Me.CheckBox2.Tag = ""
        Me.CheckBox2.Text = "ສະເພາະລາຍການເຄືອນໄຫວ"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(698, 126)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(230, 31)
        Me.CheckBox1.TabIndex = 45564
        Me.CheckBox1.Tag = "4006"
        Me.CheckBox1.Text = "ພາຍຫຼັງສ້າງໃບລາຍງານຜົນໄດ້ຮັບ"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(269, 226)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(58, 27)
        Me.Label14.TabIndex = 45568
        Me.Label14.Text = "ສະແດງ"
        Me.Label14.Visible = False
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(259, 46)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(70, 27)
        Me.Label15.TabIndex = 45567
        Me.Label15.Tag = "2054"
        Me.Label15.Text = "ຮອດວັນທີ"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(271, 121)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(22, 27)
        Me.Label18.TabIndex = 45566
        Me.Label18.Tag = "2049"
        Me.Label18.Text = "ປີ"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(271, 81)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(22, 27)
        Me.Label19.TabIndex = 45565
        Me.Label19.Tag = "2049"
        Me.Label19.Text = "ປີ"
        '
        'CheckBnk
        '
        Me.CheckBnk.AutoSize = True
        Me.CheckBnk.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.CheckBnk.Location = New System.Drawing.Point(270, 15)
        Me.CheckBnk.Name = "CheckBnk"
        Me.CheckBnk.Size = New System.Drawing.Size(15, 14)
        Me.CheckBnk.TabIndex = 45562
        Me.CheckBnk.Tag = ""
        Me.CheckBnk.UseVisualStyleBackColor = True
        Me.CheckBnk.Visible = False
        '
        'RGL
        '
        Me.RGL.AutoSize = True
        Me.RGL.Checked = True
        Me.RGL.Location = New System.Drawing.Point(13, 3)
        Me.RGL.Name = "RGL"
        Me.RGL.Size = New System.Drawing.Size(96, 31)
        Me.RGL.TabIndex = 45569
        Me.RGL.TabStop = True
        Me.RGL.Text = "ແບບທົ່ວໄປ"
        Me.RGL.UseVisualStyleBackColor = True
        '
        'RDtail
        '
        Me.RDtail.AutoSize = True
        Me.RDtail.Location = New System.Drawing.Point(113, 4)
        Me.RDtail.Name = "RDtail"
        Me.RDtail.Size = New System.Drawing.Size(125, 31)
        Me.RDtail.TabIndex = 45570
        Me.RDtail.Text = "ຕາມຫມວດບັນຊີ"
        Me.RDtail.UseVisualStyleBackColor = True
        '
        'RGroup
        '
        Me.RGroup.AutoSize = True
        Me.RGroup.Location = New System.Drawing.Point(243, 4)
        Me.RGroup.Name = "RGroup"
        Me.RGroup.Size = New System.Drawing.Size(130, 31)
        Me.RGroup.TabIndex = 45571
        Me.RGroup.Text = "ສະເພາະບັນຊີແມ່"
        Me.RGroup.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(333, 5)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(118, 34)
        Me.TextBox1.TabIndex = 45572
        '
        'RT
        '
        Me.RT.AutoSize = True
        Me.RT.Location = New System.Drawing.Point(6, 160)
        Me.RT.Name = "RT"
        Me.RT.Size = New System.Drawing.Size(102, 31)
        Me.RT.TabIndex = 45641
        Me.RT.Tag = "7078"
        Me.RT.Text = "ປະຈຳເດືອນ"
        Me.RT.UseVisualStyleBackColor = True
        '
        'yyt
        '
        Me.yyt.CustomFormat = "yyyy"
        Me.yyt.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.yyt.Location = New System.Drawing.Point(332, 156)
        Me.yyt.Name = "yyt"
        Me.yyt.Size = New System.Drawing.Size(118, 34)
        Me.yyt.TabIndex = 45640
        '
        'Ct
        '
        Me.Ct.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Ct.FormattingEnabled = True
        Me.Ct.Items.AddRange(New Object() {"6 ເດືອນຕົ້ນປີ", "6 ເດືອນທ້າຍປີ"})
        Me.Ct.Location = New System.Drawing.Point(135, 156)
        Me.Ct.Name = "Ct"
        Me.Ct.Size = New System.Drawing.Size(119, 35)
        Me.Ct.TabIndex = 45639
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.RGL)
        Me.Panel2.Controls.Add(Me.RDtail)
        Me.Panel2.Controls.Add(Me.RGroup)
        Me.Panel2.Location = New System.Drawing.Point(724, 190)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(389, 36)
        Me.Panel2.TabIndex = 45642
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button4.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.Image = CType(resources.GetObject("Button4.Image"), System.Drawing.Image)
        Me.Button4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button4.Location = New System.Drawing.Point(460, 5)
        Me.Button4.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(99, 34)
        Me.Button4.TabIndex = 45981
        Me.Button4.Tag = "3033"
        Me.Button4.Text = "Export"
        Me.Button4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button4.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(431, 195)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(86, 24)
        Me.Label2.TabIndex = 45982
        Me.Label2.Tag = "2020"
        Me.Label2.Text = "ສະກຸນເງິນ"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CMB_Curr
        '
        Me.CMB_Curr.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CMB_Curr.FormattingEnabled = True
        Me.CMB_Curr.Location = New System.Drawing.Point(523, 191)
        Me.CMB_Curr.Name = "CMB_Curr"
        Me.CMB_Curr.Size = New System.Drawing.Size(77, 30)
        Me.CMB_Curr.TabIndex = 45983
        Me.CMB_Curr.Text = "EQVL"
        '
        'txtRate
        '
        Me.txtRate.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtRate.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRate.Location = New System.Drawing.Point(606, 192)
        Me.txtRate.Name = "txtRate"
        Me.txtRate.Size = New System.Drawing.Size(85, 29)
        Me.txtRate.TabIndex = 45984
        Me.txtRate.Text = "1"
        Me.txtRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtRate2
        '
        Me.txtRate2.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtRate2.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRate2.Location = New System.Drawing.Point(697, 192)
        Me.txtRate2.Name = "txtRate2"
        Me.txtRate2.Size = New System.Drawing.Size(22, 29)
        Me.txtRate2.TabIndex = 45985
        Me.txtRate2.Text = "1"
        Me.txtRate2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtcurr_name2
        '
        Me.txtcurr_name2.BackColor = System.Drawing.Color.White
        Me.txtcurr_name2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcurr_name2.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcurr_name2.ForeColor = System.Drawing.Color.Blue
        Me.txtcurr_name2.Location = New System.Drawing.Point(1041, 77)
        Me.txtcurr_name2.Name = "txtcurr_name2"
        Me.txtcurr_name2.Size = New System.Drawing.Size(100, 30)
        Me.txtcurr_name2.TabIndex = 46031
        Me.txtcurr_name2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtcurr_name2.Visible = False
        '
        'CheckBox4
        '
        Me.CheckBox4.AutoSize = True
        Me.CheckBox4.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox4.Location = New System.Drawing.Point(309, 193)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(114, 31)
        Me.CheckBox4.TabIndex = 46034
        Me.CheckBox4.Text = "ທຽບເທົ່າເງິນ"
        Me.CheckBox4.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.Controls.Add(Me.TxtHeader)
        Me.Panel3.Controls.Add(Me.TxtS1)
        Me.Panel3.Controls.Add(Me.TxtS2)
        Me.Panel3.Controls.Add(Me.TxtS3)
        Me.Panel3.Controls.Add(Me.Label20)
        Me.Panel3.Controls.Add(Me.TxtS4)
        Me.Panel3.Controls.Add(Me.TxtPP)
        Me.Panel3.Controls.Add(Me.Label22)
        Me.Panel3.Controls.Add(Me.Label3)
        Me.Panel3.Controls.Add(Me.Label4)
        Me.Panel3.Controls.Add(Me.Label13)
        Me.Panel3.Controls.Add(Me.Label21)
        Me.Panel3.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel3.Location = New System.Drawing.Point(1137, 5)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(177, 226)
        Me.Panel3.TabIndex = 46062
        '
        'TxtHeader
        '
        Me.TxtHeader.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtHeader.Location = New System.Drawing.Point(84, 17)
        Me.TxtHeader.Name = "TxtHeader"
        Me.TxtHeader.Size = New System.Drawing.Size(86, 30)
        Me.TxtHeader.TabIndex = 46042
        Me.TxtHeader.Visible = False
        '
        'TxtS1
        '
        Me.TxtS1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtS1.Location = New System.Drawing.Point(102, 50)
        Me.TxtS1.Name = "TxtS1"
        Me.TxtS1.Size = New System.Drawing.Size(68, 30)
        Me.TxtS1.TabIndex = 46043
        '
        'TxtS2
        '
        Me.TxtS2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtS2.Location = New System.Drawing.Point(102, 82)
        Me.TxtS2.Name = "TxtS2"
        Me.TxtS2.Size = New System.Drawing.Size(68, 30)
        Me.TxtS2.TabIndex = 46044
        '
        'TxtS3
        '
        Me.TxtS3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtS3.Location = New System.Drawing.Point(102, 114)
        Me.TxtS3.Name = "TxtS3"
        Me.TxtS3.Size = New System.Drawing.Size(68, 30)
        Me.TxtS3.TabIndex = 46045
        '
        'Label20
        '
        Me.Label20.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.Location = New System.Drawing.Point(3, 183)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(93, 24)
        Me.Label20.TabIndex = 46050
        Me.Label20.Text = "ທີ່"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TxtS4
        '
        Me.TxtS4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtS4.Location = New System.Drawing.Point(102, 146)
        Me.TxtS4.Name = "TxtS4"
        Me.TxtS4.Size = New System.Drawing.Size(68, 30)
        Me.TxtS4.TabIndex = 46046
        '
        'TxtPP
        '
        Me.TxtPP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPP.Location = New System.Drawing.Point(102, 179)
        Me.TxtPP.Name = "TxtPP"
        Me.TxtPP.Size = New System.Drawing.Size(68, 30)
        Me.TxtPP.TabIndex = 46047
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.Location = New System.Drawing.Point(36, 18)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(45, 29)
        Me.Label22.TabIndex = 46048
        Me.Label22.Text = "ຫົວຂໍ້"
        Me.Label22.Visible = False
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(3, 83)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(97, 24)
        Me.Label3.TabIndex = 46053
        Me.Label3.Text = "ລາຍເຊັນ2"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(3, 121)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(97, 24)
        Me.Label4.TabIndex = 46052
        Me.Label4.Text = "ລາຍເຊັນ3"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(3, 153)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(97, 24)
        Me.Label13.TabIndex = 46051
        Me.Label13.Text = "ລາຍເຊັນ4"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label21
        '
        Me.Label21.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.Location = New System.Drawing.Point(3, 51)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(97, 24)
        Me.Label21.TabIndex = 46049
        Me.Label21.Text = "ລາຍເຊັນ1"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'RD
        '
        Me.RD.AutoSize = True
        Me.RD.Location = New System.Drawing.Point(6, 47)
        Me.RD.Name = "RD"
        Me.RD.Size = New System.Drawing.Size(95, 31)
        Me.RD.TabIndex = 46063
        Me.RD.Tag = "5032"
        Me.RD.Text = "ປະຈຳວັນທີ"
        Me.RD.UseVisualStyleBackColor = True
        '
        'RM
        '
        Me.RM.AutoSize = True
        Me.RM.Checked = True
        Me.RM.Location = New System.Drawing.Point(6, 81)
        Me.RM.Name = "RM"
        Me.RM.Size = New System.Drawing.Size(102, 31)
        Me.RM.TabIndex = 46064
        Me.RM.TabStop = True
        Me.RM.Tag = "5033"
        Me.RM.Text = "ປະຈຳເດືອນ"
        Me.RM.UseVisualStyleBackColor = True
        '
        'RY
        '
        Me.RY.AutoSize = True
        Me.RY.Location = New System.Drawing.Point(6, 196)
        Me.RY.Name = "RY"
        Me.RY.Size = New System.Drawing.Size(76, 31)
        Me.RY.TabIndex = 46065
        Me.RY.TabStop = True
        Me.RY.Tag = "5035"
        Me.RY.Text = "ປະຈຳປີ"
        Me.RY.UseVisualStyleBackColor = True
        '
        'RP
        '
        Me.RP.AutoSize = True
        Me.RP.Location = New System.Drawing.Point(6, 122)
        Me.RP.Name = "RP"
        Me.RP.Size = New System.Drawing.Size(93, 31)
        Me.RP.TabIndex = 46066
        Me.RP.TabStop = True
        Me.RP.Tag = "5034"
        Me.RP.Text = "ປະຈຳງວດ"
        Me.RP.UseVisualStyleBackColor = True
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(454, 127)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(69, 27)
        Me.Label27.TabIndex = 46067
        Me.Label27.Tag = "2011"
        Me.Label27.Text = "ໜ່ວຍງານ"
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Location = New System.Drawing.Point(698, 155)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(241, 31)
        Me.CheckBox3.TabIndex = 46068
        Me.CheckBox3.Tag = "4006"
        Me.CheckBox3.Text = "ໃບດູນດ່ຽງປະຈຳປີ ຫລັງການປັບປຸງ"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'nn
        '
        Me.nn.AutoSize = True
        Me.nn.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.nn.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.nn.Location = New System.Drawing.Point(688, 11)
        Me.nn.Name = "nn"
        Me.nn.Size = New System.Drawing.Size(133, 29)
        Me.nn.TabIndex = 46069
        Me.nn.Tag = "2065"
        Me.nn.Text = "ໃບດູນດ່ຽງສຳຮອງ"
        '
        'BtnPreview
        '
        Me.BtnPreview.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnPreview.Location = New System.Drawing.Point(44, 3)
        Me.BtnPreview.Name = "BtnPreview"
        Me.BtnPreview.Size = New System.Drawing.Size(105, 35)
        Me.BtnPreview.TabIndex = 46070
        Me.BtnPreview.Tag = "3006"
        Me.BtnPreview.Text = "ວິວ/ເບິ່ງ"
        Me.BtnPreview.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnPreview.UseVisualStyleBackColor = True
        '
        'FmTrialBalanceReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 27.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(221, Byte), Integer), CType(CType(254, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1318, 736)
        Me.Controls.Add(Me.BtnPreview)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.nn)
        Me.Controls.Add(Me.CheckBox3)
        Me.Controls.Add(Me.Label27)
        Me.Controls.Add(Me.RP)
        Me.Controls.Add(Me.RY)
        Me.Controls.Add(Me.RM)
        Me.Controls.Add(Me.RD)
        Me.Controls.Add(Me.CheckBox4)
        Me.Controls.Add(Me.txtcurr_name2)
        Me.Controls.Add(Me.txtRate2)
        Me.Controls.Add(Me.txtRate)
        Me.Controls.Add(Me.CMB_Curr)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.RT)
        Me.Controls.Add(Me.yyt)
        Me.Controls.Add(Me.Ct)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.CheckBnk)
        Me.Controls.Add(Me.Off_Usr)
        Me.Controls.Add(Me.Cx)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.BalanceType)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.RaParent)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.BtnRefresh)
        Me.Controls.Add(Me.FG)
        Me.Controls.Add(Me.L5)
        Me.Controls.Add(Me.Lb)
        Me.Controls.Add(Me.Dt)
        Me.Controls.Add(Me.yy)
        Me.Controls.Add(Me.Ds)
        Me.Controls.Add(Me.Pyy)
        Me.Controls.Add(Me.Myy)
        Me.Controls.Add(Me.Period)
        Me.Controls.Add(Me.DMonth)
        Me.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 6, 4, 6)
        Me.Name = "FmTrialBalanceReport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FmTrialBalanceReport"
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
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
    Friend WithEvents ReDr As System.Windows.Forms.TextBox
    Friend WithEvents OpDr As System.Windows.Forms.TextBox
    Friend WithEvents BOpDr As System.Windows.Forms.TextBox
    Friend WithEvents OpCr As System.Windows.Forms.TextBox
    Friend WithEvents AmtDr As System.Windows.Forms.TextBox
    Friend WithEvents BAmtDr As System.Windows.Forms.TextBox
    Friend WithEvents AmtCr As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ReCr As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents BReDr As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents BalanceType As System.Windows.Forms.ComboBox
    Friend WithEvents Cx As System.Windows.Forms.ComboBox
    Friend WithEvents Off_Usr As System.Windows.Forms.ComboBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents CheckBnk As System.Windows.Forms.CheckBox
    Friend WithEvents RGL As System.Windows.Forms.RadioButton
    Friend WithEvents RDtail As System.Windows.Forms.RadioButton
    Friend WithEvents RGroup As System.Windows.Forms.RadioButton
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents RT As System.Windows.Forms.RadioButton
    Friend WithEvents yyt As System.Windows.Forms.DateTimePicker
    Friend WithEvents Ct As System.Windows.Forms.ComboBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CMB_Curr As System.Windows.Forms.ComboBox
    Friend WithEvents txtRate As System.Windows.Forms.TextBox
    Friend WithEvents txtRate2 As System.Windows.Forms.TextBox
    Friend WithEvents txtcurr_name2 As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox4 As System.Windows.Forms.CheckBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents TxtHeader As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtS1 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtS2 As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents TxtS3 As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents TxtS4 As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents TxtPP As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents RD As System.Windows.Forms.RadioButton
    Friend WithEvents RM As System.Windows.Forms.RadioButton
    Friend WithEvents RY As System.Windows.Forms.RadioButton
    Friend WithEvents RP As System.Windows.Forms.RadioButton
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents nn As System.Windows.Forms.Label
    Friend WithEvents BtnPreview As System.Windows.Forms.Button
End Class
