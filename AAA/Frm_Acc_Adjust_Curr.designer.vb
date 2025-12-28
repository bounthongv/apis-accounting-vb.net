<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_Acc_Adjust_Curr
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_Acc_Adjust_Curr))
        Me.FG = New AxVSFlex8U.AxVSFlexGrid
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtSumTotalAmountCr = New System.Windows.Forms.TextBox
        Me.DDR = New System.Windows.Forms.TextBox
        Me.CCR = New System.Windows.Forms.TextBox
        Me.Dr = New System.Windows.Forms.TextBox
        Me.txtSumAmountDr = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.txtSumAmountCr = New System.Windows.Forms.TextBox
        Me.txtSumTotalAmountDr = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Cr = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.txtRate = New System.Windows.Forms.TextBox
        Me.txtAmount = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtBill_no = New System.Windows.Forms.TextBox
        Me.txt_dt = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.CMBBK_ID = New System.Windows.Forms.ComboBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtBook_nm = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.CMB_Curr = New System.Windows.Forms.ComboBox
        Me.txtAmount_Later = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtAmount_Lak = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.txt_descrip = New System.Windows.Forms.TextBox
        Me.txt_descripE = New System.Windows.Forms.TextBox
        Me.Txt_Referno = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.txtUSD_Rate = New System.Windows.Forms.TextBox
        Me.Button6 = New System.Windows.Forms.Button
        Me.txtAmount_USD = New System.Windows.Forms.TextBox
        Me.txtAC_code_nm = New System.Windows.Forms.TextBox
        Me.txtAC_code = New System.Windows.Forms.TextBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.txtRete_AVG = New System.Windows.Forms.TextBox
        Me.Label36 = New System.Windows.Forms.Label
        Me.txtamt = New System.Windows.Forms.TextBox
        Me.Label37 = New System.Windows.Forms.Label
        Me.txtAC_type = New System.Windows.Forms.TextBox
        Me.Label38 = New System.Windows.Forms.Label
        Me.Label39 = New System.Windows.Forms.Label
        Me.txtDiff = New System.Windows.Forms.TextBox
        Me.txt_Curr = New System.Windows.Forms.TextBox
        Me.Button5 = New System.Windows.Forms.Button
        Me.BtnSearch = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.BtnSave = New System.Windows.Forms.Button
        Me.BtnAddNew = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.txtRateUSD = New System.Windows.Forms.TextBox
        CType(Me.FG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'FG
        '
        Me.FG.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FG.DataSource = Nothing
        Me.FG.Location = New System.Drawing.Point(4, 270)
        Me.FG.Name = "FG"
        Me.FG.OcxState = CType(resources.GetObject("FG.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FG.Size = New System.Drawing.Size(1284, 184)
        Me.FG.TabIndex = 363
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.Label12)
        Me.Panel2.Controls.Add(Me.txtSumTotalAmountCr)
        Me.Panel2.Controls.Add(Me.DDR)
        Me.Panel2.Controls.Add(Me.CCR)
        Me.Panel2.Controls.Add(Me.Dr)
        Me.Panel2.Controls.Add(Me.txtSumAmountDr)
        Me.Panel2.Controls.Add(Me.Label15)
        Me.Panel2.Controls.Add(Me.txtSumAmountCr)
        Me.Panel2.Controls.Add(Me.txtSumTotalAmountDr)
        Me.Panel2.Controls.Add(Me.Label17)
        Me.Panel2.Controls.Add(Me.Label18)
        Me.Panel2.Controls.Add(Me.Label11)
        Me.Panel2.Controls.Add(Me.Label16)
        Me.Panel2.Controls.Add(Me.Cr)
        Me.Panel2.Controls.Add(Me.Label14)
        Me.Panel2.Controls.Add(Me.Label13)
        Me.Panel2.Location = New System.Drawing.Point(4, 461)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1284, 80)
        Me.Panel2.TabIndex = 363
        Me.Panel2.Tag = "1"
        Me.Panel2.Visible = False
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.BackColor = System.Drawing.SystemColors.HighlightText
        Me.Label12.Location = New System.Drawing.Point(10, 41)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(69, 21)
        Me.Label12.TabIndex = 194
        Me.Label12.Tag = "2028"
        Me.Label12.Text = "ລວມຈົດມີ :"
        '
        'txtSumTotalAmountCr
        '
        Me.txtSumTotalAmountCr.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtSumTotalAmountCr.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSumTotalAmountCr.ForeColor = System.Drawing.Color.Black
        Me.txtSumTotalAmountCr.Location = New System.Drawing.Point(401, 37)
        Me.txtSumTotalAmountCr.Name = "txtSumTotalAmountCr"
        Me.txtSumTotalAmountCr.ReadOnly = True
        Me.txtSumTotalAmountCr.Size = New System.Drawing.Size(220, 29)
        Me.txtSumTotalAmountCr.TabIndex = 186
        Me.txtSumTotalAmountCr.Text = "0.00"
        Me.txtSumTotalAmountCr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'DDR
        '
        Me.DDR.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DDR.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.DDR.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DDR.ForeColor = System.Drawing.Color.Black
        Me.DDR.Location = New System.Drawing.Point(1063, 5)
        Me.DDR.Name = "DDR"
        Me.DDR.ReadOnly = True
        Me.DDR.Size = New System.Drawing.Size(212, 29)
        Me.DDR.TabIndex = 190
        Me.DDR.Text = "0.00"
        Me.DDR.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'CCR
        '
        Me.CCR.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CCR.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.CCR.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CCR.ForeColor = System.Drawing.Color.Black
        Me.CCR.Location = New System.Drawing.Point(1063, 39)
        Me.CCR.Name = "CCR"
        Me.CCR.ReadOnly = True
        Me.CCR.Size = New System.Drawing.Size(212, 29)
        Me.CCR.TabIndex = 187
        Me.CCR.Text = "0.00"
        Me.CCR.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Dr
        '
        Me.Dr.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Dr.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Dr.ForeColor = System.Drawing.Color.Black
        Me.Dr.Location = New System.Drawing.Point(716, 6)
        Me.Dr.Name = "Dr"
        Me.Dr.ReadOnly = True
        Me.Dr.Size = New System.Drawing.Size(219, 29)
        Me.Dr.TabIndex = 188
        Me.Dr.Text = "0.00"
        Me.Dr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtSumAmountDr
        '
        Me.txtSumAmountDr.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtSumAmountDr.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSumAmountDr.ForeColor = System.Drawing.Color.Black
        Me.txtSumAmountDr.Location = New System.Drawing.Point(83, 4)
        Me.txtSumAmountDr.Name = "txtSumAmountDr"
        Me.txtSumAmountDr.ReadOnly = True
        Me.txtSumAmountDr.Size = New System.Drawing.Size(208, 29)
        Me.txtSumAmountDr.TabIndex = 191
        Me.txtSumAmountDr.Text = "0.00"
        Me.txtSumAmountDr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.BackColor = System.Drawing.SystemColors.Control
        Me.Label15.Location = New System.Drawing.Point(955, 45)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(106, 21)
        Me.Label15.TabIndex = 197
        Me.Label15.Tag = "2033"
        Me.Label15.Text = "ຄ່າຜິດດ່ຽງມີ(ກີບ) :"
        '
        'txtSumAmountCr
        '
        Me.txtSumAmountCr.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtSumAmountCr.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSumAmountCr.ForeColor = System.Drawing.Color.Black
        Me.txtSumAmountCr.Location = New System.Drawing.Point(83, 37)
        Me.txtSumAmountCr.Name = "txtSumAmountCr"
        Me.txtSumAmountCr.ReadOnly = True
        Me.txtSumAmountCr.Size = New System.Drawing.Size(208, 29)
        Me.txtSumAmountCr.TabIndex = 185
        Me.txtSumAmountCr.Text = "0.00"
        Me.txtSumAmountCr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtSumTotalAmountDr
        '
        Me.txtSumTotalAmountDr.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtSumTotalAmountDr.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSumTotalAmountDr.ForeColor = System.Drawing.Color.Black
        Me.txtSumTotalAmountDr.Location = New System.Drawing.Point(401, 4)
        Me.txtSumTotalAmountDr.Name = "txtSumTotalAmountDr"
        Me.txtSumTotalAmountDr.ReadOnly = True
        Me.txtSumTotalAmountDr.Size = New System.Drawing.Size(220, 29)
        Me.txtSumTotalAmountDr.TabIndex = 192
        Me.txtSumTotalAmountDr.Text = "0.00"
        Me.txtSumTotalAmountDr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.BackColor = System.Drawing.SystemColors.Control
        Me.Label17.Location = New System.Drawing.Point(950, 9)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(111, 21)
        Me.Label17.TabIndex = 200
        Me.Label17.Tag = "2031"
        Me.Label17.Text = "ຄ່າຜິດດ່ຽງໜີ້(ກີບ) :"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.BackColor = System.Drawing.SystemColors.HighlightText
        Me.Label18.Location = New System.Drawing.Point(303, 42)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(100, 21)
        Me.Label18.TabIndex = 198
        Me.Label18.Tag = "2032"
        Me.Label18.Text = "ລວມຈົດມີ (ກີບ) :" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.BackColor = System.Drawing.SystemColors.HighlightText
        Me.Label11.Location = New System.Drawing.Point(3, 8)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(74, 21)
        Me.Label11.TabIndex = 195
        Me.Label11.Tag = "2026"
        Me.Label11.Text = "ລວມຈົດໜີ້ :"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.BackColor = System.Drawing.SystemColors.HighlightText
        Me.Label16.Location = New System.Drawing.Point(297, 8)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(105, 21)
        Me.Label16.TabIndex = 199
        Me.Label16.Tag = "2030"
        Me.Label16.Text = "ລວມຈົດໜີ (ກີບ) :" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Cr
        '
        Me.Cr.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Cr.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cr.ForeColor = System.Drawing.Color.Black
        Me.Cr.Location = New System.Drawing.Point(716, 39)
        Me.Cr.Name = "Cr"
        Me.Cr.ReadOnly = True
        Me.Cr.Size = New System.Drawing.Size(219, 29)
        Me.Cr.TabIndex = 189
        Me.Cr.Text = "0.00"
        Me.Cr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.BackColor = System.Drawing.SystemColors.HighlightText
        Me.Label14.Location = New System.Drawing.Point(632, 43)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(82, 21)
        Me.Label14.TabIndex = 193
        Me.Label14.Tag = "2029"
        Me.Label14.Text = "ຄ່າຜິດດ່ຽງມີ :"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.BackColor = System.Drawing.SystemColors.HighlightText
        Me.Label13.Location = New System.Drawing.Point(627, 9)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(87, 21)
        Me.Label13.TabIndex = 196
        Me.Label13.Tag = "2027"
        Me.Label13.Text = "ຄ່າຜິດດ່ຽງໜີ້ :"
        '
        'txtRate
        '
        Me.txtRate.BackColor = System.Drawing.Color.PaleGreen
        Me.txtRate.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRate.ForeColor = System.Drawing.Color.Blue
        Me.txtRate.Location = New System.Drawing.Point(466, 171)
        Me.txtRate.Name = "txtRate"
        Me.txtRate.Size = New System.Drawing.Size(73, 30)
        Me.txtRate.TabIndex = 370
        Me.txtRate.Text = "0.00"
        Me.txtRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtAmount
        '
        Me.txtAmount.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAmount.ForeColor = System.Drawing.Color.Black
        Me.txtAmount.Location = New System.Drawing.Point(150, 138)
        Me.txtAmount.MaxLength = 12
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(151, 32)
        Me.txtAmount.TabIndex = 373
        Me.txtAmount.Text = "0.00"
        Me.txtAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(20, 142)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(134, 24)
        Me.Label3.TabIndex = 374
        Me.Label3.Tag = "2017"
        Me.Label3.Text = "ມູນຄ່າເດີມໃນບັນຊີ :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(961, 49)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 21)
        Me.Label1.TabIndex = 375
        Me.Label1.Text = "Label1"
        Me.Label1.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(1030, 52)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 21)
        Me.Label2.TabIndex = 376
        Me.Label2.Text = "Label2"
        Me.Label2.Visible = False
        '
        'txtBill_no
        '
        Me.txtBill_no.BackColor = System.Drawing.Color.White
        Me.txtBill_no.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBill_no.Location = New System.Drawing.Point(615, 42)
        Me.txtBill_no.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtBill_no.Name = "txtBill_no"
        Me.txtBill_no.Size = New System.Drawing.Size(123, 30)
        Me.txtBill_no.TabIndex = 377
        Me.txtBill_no.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txt_dt
        '
        Me.txt_dt.CustomFormat = "dd/MM/yyyy"
        Me.txt_dt.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_dt.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txt_dt.Location = New System.Drawing.Point(386, 41)
        Me.txt_dt.Name = "txt_dt"
        Me.txt_dt.ShowUpDown = True
        Me.txt_dt.Size = New System.Drawing.Size(111, 34)
        Me.txt_dt.TabIndex = 379
        Me.txt_dt.Value = New Date(2009, 12, 28, 0, 0, 0, 0)
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(529, 43)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(79, 28)
        Me.Label4.TabIndex = 378
        Me.Label4.Text = "ໃບຢັ້ງຢືນ :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(318, 43)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(62, 28)
        Me.Label6.TabIndex = 382
        Me.Label6.Text = "ວັນທີ:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CMBBK_ID
        '
        Me.CMBBK_ID.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CMBBK_ID.FormattingEnabled = True
        Me.CMBBK_ID.Location = New System.Drawing.Point(149, 40)
        Me.CMBBK_ID.Name = "CMBBK_ID"
        Me.CMBBK_ID.Size = New System.Drawing.Size(152, 32)
        Me.CMBBK_ID.TabIndex = 384
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(19, 44)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(129, 24)
        Me.Label7.TabIndex = 385
        Me.Label7.Tag = "2017"
        Me.Label7.Text = "ປື້ມບັນຊີປະຈໍາວັນ :"
        '
        'txtBook_nm
        '
        Me.txtBook_nm.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBook_nm.Location = New System.Drawing.Point(1475, 49)
        Me.txtBook_nm.Name = "txtBook_nm"
        Me.txtBook_nm.ReadOnly = True
        Me.txtBook_nm.Size = New System.Drawing.Size(315, 34)
        Me.txtBook_nm.TabIndex = 386
        Me.txtBook_nm.Visible = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(69, 175)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(85, 24)
        Me.Label8.TabIndex = 387
        Me.Label8.Tag = "2017"
        Me.Label8.Text = "ສະກຸນເງິນ :"
        '
        'CMB_Curr
        '
        Me.CMB_Curr.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CMB_Curr.FormattingEnabled = True
        Me.CMB_Curr.Items.AddRange(New Object() {"LAK", "THB", "USD"})
        Me.CMB_Curr.Location = New System.Drawing.Point(150, 172)
        Me.CMB_Curr.Name = "CMB_Curr"
        Me.CMB_Curr.Size = New System.Drawing.Size(76, 29)
        Me.CMB_Curr.TabIndex = 388
        '
        'txtAmount_Later
        '
        Me.txtAmount_Later.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtAmount_Later.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAmount_Later.Location = New System.Drawing.Point(149, 234)
        Me.txtAmount_Later.Name = "txtAmount_Later"
        Me.txtAmount_Later.ReadOnly = True
        Me.txtAmount_Later.Size = New System.Drawing.Size(792, 34)
        Me.txtAmount_Later.TabIndex = 467
        Me.txtAmount_Later.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(17, 240)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(127, 24)
        Me.Label9.TabIndex = 468
        Me.Label9.Tag = "2017"
        Me.Label9.Text = "ຂຽນເປັນຕົວໜັງສື :"
        '
        'txtAmount_Lak
        '
        Me.txtAmount_Lak.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAmount_Lak.ForeColor = System.Drawing.Color.Black
        Me.txtAmount_Lak.Location = New System.Drawing.Point(732, 169)
        Me.txtAmount_Lak.MaxLength = 12
        Me.txtAmount_Lak.Name = "txtAmount_Lak"
        Me.txtAmount_Lak.ReadOnly = True
        Me.txtAmount_Lak.Size = New System.Drawing.Size(209, 32)
        Me.txtAmount_Lak.TabIndex = 469
        Me.txtAmount_Lak.Text = "0.00"
        Me.txtAmount_Lak.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(240, 172)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(228, 28)
        Me.Label10.TabIndex = 470
        Me.Label10.Text = "ອັດຕາແລກປ່ຽນມື້ປິດບັນຊີ :"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label19
        '
        Me.Label19.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.Location = New System.Drawing.Point(16, 77)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(130, 28)
        Me.Label19.TabIndex = 472
        Me.Label19.Text = "ເນື້ອໃນລາຍການ :"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txt_descrip
        '
        Me.txt_descrip.BackColor = System.Drawing.Color.White
        Me.txt_descrip.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_descrip.Location = New System.Drawing.Point(149, 73)
        Me.txt_descrip.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txt_descrip.Name = "txt_descrip"
        Me.txt_descrip.Size = New System.Drawing.Size(792, 30)
        Me.txt_descrip.TabIndex = 473
        '
        'txt_descripE
        '
        Me.txt_descripE.BackColor = System.Drawing.Color.White
        Me.txt_descripE.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_descripE.Location = New System.Drawing.Point(965, 175)
        Me.txt_descripE.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txt_descripE.Name = "txt_descripE"
        Me.txt_descripE.Size = New System.Drawing.Size(185, 30)
        Me.txt_descripE.TabIndex = 475
        Me.txt_descripE.Visible = False
        '
        'Txt_Referno
        '
        Me.Txt_Referno.BackColor = System.Drawing.Color.White
        Me.Txt_Referno.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Txt_Referno.Location = New System.Drawing.Point(826, 39)
        Me.Txt_Referno.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Txt_Referno.Name = "Txt_Referno"
        Me.Txt_Referno.Size = New System.Drawing.Size(115, 30)
        Me.Txt_Referno.TabIndex = 478
        '
        'Label22
        '
        Me.Label22.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label22.Font = New System.Drawing.Font("Saysettha OT", 14.0!, System.Drawing.FontStyle.Bold)
        Me.Label22.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label22.Location = New System.Drawing.Point(545, -5)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(293, 56)
        Me.Label22.TabIndex = 480
        Me.Label22.Text = "ປັບປຸງເງິນຕາ"
        Me.Label22.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtUSD_Rate
        '
        Me.txtUSD_Rate.BackColor = System.Drawing.Color.PaleGreen
        Me.txtUSD_Rate.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUSD_Rate.ForeColor = System.Drawing.Color.Blue
        Me.txtUSD_Rate.Location = New System.Drawing.Point(754, 39)
        Me.txtUSD_Rate.Multiline = True
        Me.txtUSD_Rate.Name = "txtUSD_Rate"
        Me.txtUSD_Rate.ReadOnly = True
        Me.txtUSD_Rate.Size = New System.Drawing.Size(66, 32)
        Me.txtUSD_Rate.TabIndex = 431
        Me.txtUSD_Rate.Text = "0.00"
        Me.txtUSD_Rate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Button6.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button6.Image = CType(resources.GetObject("Button6.Image"), System.Drawing.Image)
        Me.Button6.Location = New System.Drawing.Point(4, 8)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(37, 30)
        Me.Button6.TabIndex = 358
        Me.Button6.UseVisualStyleBackColor = False
        '
        'txtAmount_USD
        '
        Me.txtAmount_USD.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAmount_USD.ForeColor = System.Drawing.Color.Black
        Me.txtAmount_USD.Location = New System.Drawing.Point(965, 133)
        Me.txtAmount_USD.MaxLength = 12
        Me.txtAmount_USD.Name = "txtAmount_USD"
        Me.txtAmount_USD.ReadOnly = True
        Me.txtAmount_USD.Size = New System.Drawing.Size(170, 32)
        Me.txtAmount_USD.TabIndex = 45981
        Me.txtAmount_USD.Text = "0.00"
        Me.txtAmount_USD.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtAmount_USD.Visible = False
        '
        'txtAC_code_nm
        '
        Me.txtAC_code_nm.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtAC_code_nm.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAC_code_nm.ForeColor = System.Drawing.Color.Black
        Me.txtAC_code_nm.Location = New System.Drawing.Point(303, 104)
        Me.txtAC_code_nm.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtAC_code_nm.Name = "txtAC_code_nm"
        Me.txtAC_code_nm.Size = New System.Drawing.Size(423, 30)
        Me.txtAC_code_nm.TabIndex = 45992
        '
        'txtAC_code
        '
        Me.txtAC_code.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAC_code.ForeColor = System.Drawing.Color.Black
        Me.txtAC_code.Location = New System.Drawing.Point(150, 105)
        Me.txtAC_code.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtAC_code.Name = "txtAC_code"
        Me.txtAC_code.Size = New System.Drawing.Size(116, 30)
        Me.txtAC_code.TabIndex = 45990
        '
        'Label26
        '
        Me.Label26.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.ForeColor = System.Drawing.Color.Blue
        Me.Label26.Location = New System.Drawing.Point(4, 108)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(141, 28)
        Me.Label26.TabIndex = 45989
        Me.Label26.Text = "ເລກບັນຊີ:"
        Me.Label26.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtRete_AVG
        '
        Me.txtRete_AVG.BackColor = System.Drawing.Color.PaleGreen
        Me.txtRete_AVG.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRete_AVG.ForeColor = System.Drawing.Color.Blue
        Me.txtRete_AVG.Location = New System.Drawing.Point(466, 138)
        Me.txtRete_AVG.Name = "txtRete_AVG"
        Me.txtRete_AVG.Size = New System.Drawing.Size(73, 30)
        Me.txtRete_AVG.TabIndex = 45993
        Me.txtRete_AVG.Text = "0.00"
        Me.txtRete_AVG.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label36
        '
        Me.Label36.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label36.Location = New System.Drawing.Point(301, 140)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(167, 28)
        Me.Label36.TabIndex = 45994
        Me.Label36.Text = "ອັດຕາແລກປ່ຽນສະເລ່ຍ :"
        Me.Label36.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtamt
        '
        Me.txtamt.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtamt.ForeColor = System.Drawing.Color.Black
        Me.txtamt.Location = New System.Drawing.Point(732, 136)
        Me.txtamt.MaxLength = 12
        Me.txtamt.Name = "txtamt"
        Me.txtamt.Size = New System.Drawing.Size(209, 32)
        Me.txtamt.TabIndex = 45995
        Me.txtamt.Text = "0.00"
        Me.txtamt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label37
        '
        Me.Label37.AutoSize = True
        Me.Label37.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label37.Location = New System.Drawing.Point(603, 141)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(128, 24)
        Me.Label37.TabIndex = 45996
        Me.Label37.Tag = "2017"
        Me.Label37.Text = "ມູນຄ່າກີບໃນບັນຊີ :"
        '
        'txtAC_type
        '
        Me.txtAC_type.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtAC_type.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAC_type.ForeColor = System.Drawing.Color.Black
        Me.txtAC_type.Location = New System.Drawing.Point(732, 104)
        Me.txtAC_type.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtAC_type.Name = "txtAC_type"
        Me.txtAC_type.Size = New System.Drawing.Size(209, 30)
        Me.txtAC_type.TabIndex = 45997
        '
        'Label38
        '
        Me.Label38.AutoSize = True
        Me.Label38.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label38.Location = New System.Drawing.Point(543, 173)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(188, 24)
        Me.Label38.TabIndex = 45998
        Me.Label38.Tag = "2017"
        Me.Label38.Text = "ມູນຄ່າກີບໃນມື້ປິດບັນຊີບັນຊີ :"
        '
        'Label39
        '
        Me.Label39.AutoSize = True
        Me.Label39.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label39.ForeColor = System.Drawing.Color.Red
        Me.Label39.Location = New System.Drawing.Point(427, 207)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(302, 24)
        Me.Label39.TabIndex = 46000
        Me.Label39.Tag = "2017"
        Me.Label39.Text = "ຈຳນວນເງິນຜິດດ່ຽງ (+) ກຳໄລ / (-) ຂາດທຶນ :"
        '
        'txtDiff
        '
        Me.txtDiff.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDiff.ForeColor = System.Drawing.Color.Black
        Me.txtDiff.Location = New System.Drawing.Point(732, 202)
        Me.txtDiff.MaxLength = 12
        Me.txtDiff.Name = "txtDiff"
        Me.txtDiff.ReadOnly = True
        Me.txtDiff.Size = New System.Drawing.Size(209, 32)
        Me.txtDiff.TabIndex = 45999
        Me.txtDiff.Text = "0.00"
        Me.txtDiff.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txt_Curr
        '
        Me.txt_Curr.BackColor = System.Drawing.Color.White
        Me.txt_Curr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_Curr.Location = New System.Drawing.Point(965, 98)
        Me.txt_Curr.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txt_Curr.Name = "txt_Curr"
        Me.txt_Curr.Size = New System.Drawing.Size(60, 30)
        Me.txt_Curr.TabIndex = 46001
        Me.txt_Curr.Text = "LAK"
        Me.txt_Curr.Visible = False
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Button5.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button5.Image = CType(resources.GetObject("Button5.Image"), System.Drawing.Image)
        Me.Button5.Location = New System.Drawing.Point(266, 105)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(35, 30)
        Me.Button5.TabIndex = 45991
        Me.Button5.Text = "..."
        Me.Button5.UseVisualStyleBackColor = False
        '
        'BtnSearch
        '
        Me.BtnSearch.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.BtnSearch.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.BtnSearch.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSearch.Image = CType(resources.GetObject("BtnSearch.Image"), System.Drawing.Image)
        Me.BtnSearch.Location = New System.Drawing.Point(160, 296)
        Me.BtnSearch.Name = "BtnSearch"
        Me.BtnSearch.Size = New System.Drawing.Size(34, 26)
        Me.BtnSearch.TabIndex = 364
        Me.BtnSearch.Tag = "3012"
        Me.BtnSearch.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnSearch.UseVisualStyleBackColor = False
        Me.BtnSearch.Visible = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(384, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(113, 30)
        Me.Button1.TabIndex = 361
        Me.Button1.Tag = "3007"
        Me.Button1.Text = "ພີມ"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.UseVisualStyleBackColor = False
        Me.Button1.Visible = False
        '
        'BtnSave
        '
        Me.BtnSave.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.BtnSave.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.BtnSave.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSave.Image = CType(resources.GetObject("BtnSave.Image"), System.Drawing.Image)
        Me.BtnSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnSave.Location = New System.Drawing.Point(267, 8)
        Me.BtnSave.Name = "BtnSave"
        Me.BtnSave.Size = New System.Drawing.Size(113, 30)
        Me.BtnSave.TabIndex = 360
        Me.BtnSave.Tag = "3006"
        Me.BtnSave.Text = "ບັນທຶກ"
        Me.BtnSave.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnSave.UseVisualStyleBackColor = False
        '
        'BtnAddNew
        '
        Me.BtnAddNew.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.BtnAddNew.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.BtnAddNew.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAddNew.Image = CType(resources.GetObject("BtnAddNew.Image"), System.Drawing.Image)
        Me.BtnAddNew.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAddNew.Location = New System.Drawing.Point(150, 8)
        Me.BtnAddNew.Name = "BtnAddNew"
        Me.BtnAddNew.Size = New System.Drawing.Size(115, 30)
        Me.BtnAddNew.TabIndex = 359
        Me.BtnAddNew.Tag = "3003"
        Me.BtnAddNew.Text = "ເພີ່ມໃໝ່"
        Me.BtnAddNew.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnAddNew.UseVisualStyleBackColor = False
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(160, 328)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(34, 26)
        Me.Button2.TabIndex = 46002
        Me.Button2.Tag = "3012"
        Me.Button2.Text = "....."
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'txtRateUSD
        '
        Me.txtRateUSD.BackColor = System.Drawing.Color.PaleGreen
        Me.txtRateUSD.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRateUSD.ForeColor = System.Drawing.Color.Blue
        Me.txtRateUSD.Location = New System.Drawing.Point(947, 206)
        Me.txtRateUSD.Name = "txtRateUSD"
        Me.txtRateUSD.Size = New System.Drawing.Size(73, 30)
        Me.txtRateUSD.TabIndex = 46003
        Me.txtRateUSD.Text = "0.00"
        Me.txtRateUSD.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Frm_Acc_Adjust_Curr
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 21.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.HighlightText
        Me.ClientSize = New System.Drawing.Size(1289, 539)
        Me.Controls.Add(Me.txtRateUSD)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.txt_Curr)
        Me.Controls.Add(Me.Label39)
        Me.Controls.Add(Me.txtDiff)
        Me.Controls.Add(Me.txtAmount_USD)
        Me.Controls.Add(Me.Txt_Referno)
        Me.Controls.Add(Me.Label38)
        Me.Controls.Add(Me.txtAC_type)
        Me.Controls.Add(Me.txtamt)
        Me.Controls.Add(Me.Label37)
        Me.Controls.Add(Me.txtRete_AVG)
        Me.Controls.Add(Me.Label36)
        Me.Controls.Add(Me.txtAC_code_nm)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.txtAC_code)
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.txt_descripE)
        Me.Controls.Add(Me.txtUSD_Rate)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txt_descrip)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.txtAmount_Lak)
        Me.Controls.Add(Me.txtRate)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.txtAmount_Later)
        Me.Controls.Add(Me.CMB_Curr)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txtBook_nm)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.CMBBK_ID)
        Me.Controls.Add(Me.txtBill_no)
        Me.Controls.Add(Me.txt_dt)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.BtnSearch)
        Me.Controls.Add(Me.txtAmount)
        Me.Controls.Add(Me.FG)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.BtnSave)
        Me.Controls.Add(Me.BtnAddNew)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label22)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Blue
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "Frm_Acc_Adjust_Curr"
        Me.Text = "Currency Adjustment"
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents BtnSave As System.Windows.Forms.Button
    Friend WithEvents BtnAddNew As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents FG As AxVSFlex8U.AxVSFlexGrid
    Friend WithEvents BtnSearch As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtSumTotalAmountCr As System.Windows.Forms.TextBox
    Friend WithEvents Dr As System.Windows.Forms.TextBox
    Friend WithEvents txtSumAmountDr As System.Windows.Forms.TextBox
    Friend WithEvents DDR As System.Windows.Forms.TextBox
    Friend WithEvents Cr As System.Windows.Forms.TextBox
    Friend WithEvents CCR As System.Windows.Forms.TextBox
    Friend WithEvents txtSumAmountCr As System.Windows.Forms.TextBox
    Friend WithEvents txtSumTotalAmountDr As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtRate As System.Windows.Forms.TextBox
    Friend WithEvents txtAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtBill_no As System.Windows.Forms.TextBox
    Friend WithEvents txt_dt As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents CMBBK_ID As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtBook_nm As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents CMB_Curr As System.Windows.Forms.ComboBox
    Friend WithEvents txtAmount_Later As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtAmount_Lak As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents txt_descrip As System.Windows.Forms.TextBox
    Friend WithEvents txt_descripE As System.Windows.Forms.TextBox
    Friend WithEvents Txt_Referno As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents txtUSD_Rate As System.Windows.Forms.TextBox
    Friend WithEvents txtAmount_USD As System.Windows.Forms.TextBox
    Friend WithEvents txtAC_code_nm As System.Windows.Forms.TextBox
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents txtAC_code As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents txtRete_AVG As System.Windows.Forms.TextBox
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents txtamt As System.Windows.Forms.TextBox
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents txtAC_type As System.Windows.Forms.TextBox
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents txtDiff As System.Windows.Forms.TextBox
    Friend WithEvents txt_Curr As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents txtRateUSD As System.Windows.Forms.TextBox
End Class
