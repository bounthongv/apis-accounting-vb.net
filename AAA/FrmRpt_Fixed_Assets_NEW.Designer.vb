<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmRpt_Fixed_Assets_NEW
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmRpt_Fixed_Assets_NEW))
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtCode = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtName = New System.Windows.Forms.TextBox
        Me.TxtNameE = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtValue = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtRemain = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.FG = New System.Windows.Forms.DataGridView
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtGrp = New System.Windows.Forms.TextBox
        Me.txtGrpNm = New System.Windows.Forms.ComboBox
        Me.LEng = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.TxtDesription = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.DateIn = New System.Windows.Forms.DateTimePicker
        Me.TxtPeriod = New System.Windows.Forms.TextBox
        Me.TxtDr = New System.Windows.Forms.TextBox
        Me.TxtCr = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.BtnSearch = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.TxtDrNm = New System.Windows.Forms.TextBox
        Me.TxtCrNm = New System.Windows.Forms.TextBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.CmbBook = New System.Windows.Forms.ComboBox
        Me.Cmb = New System.Windows.Forms.ComboBox
        Me.Button6 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.txtBookName = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.txtBill_Amt = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.Button7 = New System.Windows.Forms.Button
        Me.Button8 = New System.Windows.Forms.Button
        Me.Button9 = New System.Windows.Forms.Button
        Me.Button10 = New System.Windows.Forms.Button
        Me.dpFromDate = New System.Windows.Forms.DateTimePicker
        Me.DTUse = New System.Windows.Forms.DateTimePicker
        Me.DatePreV = New System.Windows.Forms.DateTimePicker
        Me.TextBox1 = New System.Windows.Forms.TextBox
        CType(Me.FG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(1088, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 27)
        Me.Label1.TabIndex = 283
        Me.Label1.Text = "Code:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label1.Visible = False
        '
        'TxtCode
        '
        Me.TxtCode.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtCode.Location = New System.Drawing.Point(1178, 38)
        Me.TxtCode.Name = "TxtCode"
        Me.TxtCode.Size = New System.Drawing.Size(100, 30)
        Me.TxtCode.TabIndex = 284
        Me.TxtCode.Visible = False
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(1173, 15)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(84, 27)
        Me.Label2.TabIndex = 285
        Me.Label2.Text = "Name (LA):"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label2.Visible = False
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(999, 86)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(94, 27)
        Me.Label3.TabIndex = 286
        Me.Label3.Text = "Value:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label3.Visible = False
        '
        'TxtName
        '
        Me.TxtName.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtName.Location = New System.Drawing.Point(1021, 13)
        Me.TxtName.Name = "TxtName"
        Me.TxtName.Size = New System.Drawing.Size(263, 30)
        Me.TxtName.TabIndex = 287
        Me.TxtName.Visible = False
        '
        'TxtNameE
        '
        Me.TxtNameE.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtNameE.Location = New System.Drawing.Point(1021, 44)
        Me.TxtNameE.Name = "TxtNameE"
        Me.TxtNameE.Size = New System.Drawing.Size(263, 30)
        Me.TxtNameE.TabIndex = 288
        Me.TxtNameE.Visible = False
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(1030, 83)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(104, 27)
        Me.Label4.TabIndex = 289
        Me.Label4.Text = "Desription:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label4.Visible = False
        '
        'TxtValue
        '
        Me.TxtValue.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValue.Location = New System.Drawing.Point(1021, 76)
        Me.TxtValue.Name = "TxtValue"
        Me.TxtValue.Size = New System.Drawing.Size(263, 30)
        Me.TxtValue.TabIndex = 290
        Me.TxtValue.Text = "0"
        Me.TxtValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtValue.Visible = False
        '
        'Label5
        '
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(945, 113)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(116, 27)
        Me.Label5.TabIndex = 292
        Me.Label5.Text = "Adjust Period:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label5.Visible = False
        '
        'Label6
        '
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(918, 112)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(94, 27)
        Me.Label6.TabIndex = 293
        Me.Label6.Text = "Remain:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label6.Visible = False
        '
        'TxtRemain
        '
        Me.TxtRemain.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtRemain.Location = New System.Drawing.Point(1021, 108)
        Me.TxtRemain.Name = "TxtRemain"
        Me.TxtRemain.Size = New System.Drawing.Size(263, 30)
        Me.TxtRemain.TabIndex = 294
        Me.TxtRemain.Text = "0"
        Me.TxtRemain.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtRemain.Visible = False
        '
        'Label9
        '
        Me.Label9.ForeColor = System.Drawing.Color.Black
        Me.Label9.Location = New System.Drawing.Point(710, 146)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(94, 27)
        Me.Label9.TabIndex = 300
        Me.Label9.Text = "Dr:"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label9.Visible = False
        '
        'FG
        '
        Me.FG.AllowUserToAddRows = False
        Me.FG.AllowUserToDeleteRows = False
        Me.FG.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.FG.Location = New System.Drawing.Point(16, 148)
        Me.FG.Name = "FG"
        Me.FG.ReadOnly = True
        Me.FG.Size = New System.Drawing.Size(1245, 270)
        Me.FG.TabIndex = 303
        Me.FG.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.FG.MultiSelect = False
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Saysettha OT", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Blue
        Me.Label11.Location = New System.Drawing.Point(652, 1)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(350, 36)
        Me.Label11.TabIndex = 304
        Me.Label11.Text = "Fixed Assets Register"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtGrp
        '
        Me.txtGrp.Enabled = False
        Me.txtGrp.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGrp.Location = New System.Drawing.Point(293, 44)
        Me.txtGrp.Name = "txtGrp"
        Me.txtGrp.Size = New System.Drawing.Size(78, 30)
        Me.txtGrp.TabIndex = 307
        Me.txtGrp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtGrpNm
        '
        Me.txtGrpNm.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGrpNm.FormattingEnabled = True
        Me.txtGrpNm.Items.AddRange(New Object() {"LAK", "THB", "USD"})
        Me.txtGrpNm.Location = New System.Drawing.Point(167, 78)
        Me.txtGrpNm.Name = "txtGrpNm"
        Me.txtGrpNm.Size = New System.Drawing.Size(204, 35)
        Me.txtGrpNm.TabIndex = 306
        '
        'LEng
        '
        Me.LEng.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LEng.Location = New System.Drawing.Point(7, 81)
        Me.LEng.Name = "LEng"
        Me.LEng.Size = New System.Drawing.Size(151, 24)
        Me.LEng.TabIndex = 305
        Me.LEng.Tag = "2010"
        Me.LEng.Text = "Adjustment Type:"
        Me.LEng.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label12
        '
        Me.Label12.ForeColor = System.Drawing.Color.Black
        Me.Label12.Location = New System.Drawing.Point(1170, 47)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(84, 27)
        Me.Label12.TabIndex = 309
        Me.Label12.Text = "Name (EN):"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label12.Visible = False
        '
        'TxtDesription
        '
        Me.TxtDesription.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtDesription.Location = New System.Drawing.Point(1008, 80)
        Me.TxtDesription.Name = "TxtDesription"
        Me.TxtDesription.Size = New System.Drawing.Size(341, 30)
        Me.TxtDesription.TabIndex = 310
        Me.TxtDesription.Visible = False
        '
        'Label13
        '
        Me.Label13.ForeColor = System.Drawing.Color.Black
        Me.Label13.Location = New System.Drawing.Point(50, 45)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(114, 27)
        Me.Label13.TabIndex = 311
        Me.Label13.Text = "Date Adjust:"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DateIn
        '
        Me.DateIn.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateIn.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateIn.Location = New System.Drawing.Point(167, 44)
        Me.DateIn.Name = "DateIn"
        Me.DateIn.Size = New System.Drawing.Size(120, 30)
        Me.DateIn.TabIndex = 312
        '
        'TxtPeriod
        '
        Me.TxtPeriod.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtPeriod.Location = New System.Drawing.Point(1063, 112)
        Me.TxtPeriod.Name = "TxtPeriod"
        Me.TxtPeriod.Size = New System.Drawing.Size(83, 30)
        Me.TxtPeriod.TabIndex = 313
        Me.TxtPeriod.Text = "0"
        Me.TxtPeriod.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtPeriod.Visible = False
        '
        'TxtDr
        '
        Me.TxtDr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtDr.Location = New System.Drawing.Point(790, 41)
        Me.TxtDr.Name = "TxtDr"
        Me.TxtDr.Size = New System.Drawing.Size(131, 30)
        Me.TxtDr.TabIndex = 314
        '
        'TxtCr
        '
        Me.TxtCr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtCr.Location = New System.Drawing.Point(789, 73)
        Me.TxtCr.Name = "TxtCr"
        Me.TxtCr.Size = New System.Drawing.Size(131, 30)
        Me.TxtCr.TabIndex = 315
        '
        'Label7
        '
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(729, 78)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(55, 27)
        Me.Label7.TabIndex = 316
        Me.Label7.Text = "Cr:"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'BtnSearch
        '
        Me.BtnSearch.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSearch.Location = New System.Drawing.Point(926, 40)
        Me.BtnSearch.Name = "BtnSearch"
        Me.BtnSearch.Size = New System.Drawing.Size(34, 30)
        Me.BtnSearch.TabIndex = 46037
        Me.BtnSearch.Tag = "3012"
        Me.BtnSearch.Text = "....."
        Me.BtnSearch.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnSearch.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.Location = New System.Drawing.Point(926, 73)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(34, 30)
        Me.Button4.TabIndex = 46038
        Me.Button4.Tag = "3012"
        Me.Button4.Text = "....."
        Me.Button4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button4.UseVisualStyleBackColor = True
        '
        'TxtDrNm
        '
        Me.TxtDrNm.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtDrNm.Location = New System.Drawing.Point(962, 40)
        Me.TxtDrNm.Name = "TxtDrNm"
        Me.TxtDrNm.Size = New System.Drawing.Size(392, 30)
        Me.TxtDrNm.TabIndex = 46039
        '
        'TxtCrNm
        '
        Me.TxtCrNm.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtCrNm.Location = New System.Drawing.Point(962, 72)
        Me.TxtCrNm.Name = "TxtCrNm"
        Me.TxtCrNm.Size = New System.Drawing.Size(392, 30)
        Me.TxtCrNm.TabIndex = 46040
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox2.Location = New System.Drawing.Point(729, 4)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(143, 31)
        Me.CheckBox2.TabIndex = 46041
        Me.CheckBox2.Text = "Full Depreciation"
        Me.CheckBox2.UseVisualStyleBackColor = True
        Me.CheckBox2.Visible = False
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(758, 8)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(128, 31)
        Me.CheckBox1.TabIndex = 46042
        Me.CheckBox1.Text = "Days By date"
        Me.CheckBox1.UseVisualStyleBackColor = True
        Me.CheckBox1.Visible = False
        '
        'Label8
        '
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(374, 43)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(94, 27)
        Me.Label8.TabIndex = 46043
        Me.Label8.Text = "A/C Book:"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label10
        '
        Me.Label10.ForeColor = System.Drawing.Color.Black
        Me.Label10.Location = New System.Drawing.Point(374, 82)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(94, 27)
        Me.Label10.TabIndex = 46044
        Me.Label10.Text = "Currency:"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CmbBook
        '
        Me.CmbBook.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbBook.FormattingEnabled = True
        Me.CmbBook.Location = New System.Drawing.Point(474, 45)
        Me.CmbBook.Name = "CmbBook"
        Me.CmbBook.Size = New System.Drawing.Size(86, 30)
        Me.CmbBook.TabIndex = 46046
        '
        'Cmb
        '
        Me.Cmb.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cmb.FormattingEnabled = True
        Me.Cmb.Location = New System.Drawing.Point(474, 78)
        Me.Cmb.Name = "Cmb"
        Me.Cmb.Size = New System.Drawing.Size(86, 30)
        Me.Cmb.TabIndex = 46047
        '
        'Button6
        '
        Me.Button6.Font = New System.Drawing.Font("Saysettha OT", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button6.Location = New System.Drawing.Point(804, 1)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(156, 36)
        Me.Button6.TabIndex = 46048
        Me.Button6.Text = "approve Item"
        Me.Button6.UseVisualStyleBackColor = True
        Me.Button6.Visible = False
        '
        'Button5
        '
        Me.Button5.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.Location = New System.Drawing.Point(81, 6)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(83, 36)
        Me.Button5.TabIndex = 46049
        Me.Button5.Text = "Find"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'txtBookName
        '
        Me.txtBookName.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtBookName.Location = New System.Drawing.Point(566, 43)
        Me.txtBookName.Name = "txtBookName"
        Me.txtBookName.ReadOnly = True
        Me.txtBookName.Size = New System.Drawing.Size(186, 30)
        Me.txtBookName.TabIndex = 46050
        '
        'Label14
        '
        Me.Label14.ForeColor = System.Drawing.Color.Black
        Me.Label14.Location = New System.Drawing.Point(729, 44)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(55, 27)
        Me.Label14.TabIndex = 46051
        Me.Label14.Text = "Dr:"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtBill_Amt
        '
        Me.txtBill_Amt.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBill_Amt.Location = New System.Drawing.Point(474, 111)
        Me.txtBill_Amt.Name = "txtBill_Amt"
        Me.txtBill_Amt.Size = New System.Drawing.Size(278, 30)
        Me.txtBill_Amt.TabIndex = 46052
        Me.txtBill_Amt.Text = "0"
        Me.txtBill_Amt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label15
        '
        Me.Label15.ForeColor = System.Drawing.Color.Black
        Me.Label15.Location = New System.Drawing.Point(250, 113)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(218, 27)
        Me.Label15.TabIndex = 46053
        Me.Label15.Text = "Monthly Depreciation:"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button7
        '
        Me.Button7.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button7.Location = New System.Drawing.Point(167, 6)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(144, 36)
        Me.Button7.TabIndex = 46054
        Me.Button7.Text = "approve Group"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Button8
        '
        Me.Button8.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button8.Location = New System.Drawing.Point(313, 6)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(93, 36)
        Me.Button8.TabIndex = 46055
        Me.Button8.Text = "Preview"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'Button9
        '
        Me.Button9.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button9.Location = New System.Drawing.Point(408, 6)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(105, 36)
        Me.Button9.TabIndex = 46056
        Me.Button9.Text = "Preview 2"
        Me.Button9.UseVisualStyleBackColor = True
        '
        'Button10
        '
        Me.Button10.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button10.Location = New System.Drawing.Point(566, 4)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(83, 36)
        Me.Button10.TabIndex = 46057
        Me.Button10.Text = "Find"
        Me.Button10.UseVisualStyleBackColor = True
        '
        'dpFromDate
        '
        Me.dpFromDate.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpFromDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dpFromDate.Location = New System.Drawing.Point(124, 116)
        Me.dpFromDate.Name = "dpFromDate"
        Me.dpFromDate.Size = New System.Drawing.Size(120, 30)
        Me.dpFromDate.TabIndex = 46058
        Me.dpFromDate.Visible = False
        '
        'DTUse
        '
        Me.DTUse.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DTUse.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTUse.Location = New System.Drawing.Point(764, 107)
        Me.DTUse.Name = "DTUse"
        Me.DTUse.Size = New System.Drawing.Size(120, 30)
        Me.DTUse.TabIndex = 46058
        Me.DTUse.Visible = False
        '
        'DatePreV
        '
        Me.DatePreV.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DatePreV.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DatePreV.Location = New System.Drawing.Point(16, 116)
        Me.DatePreV.Name = "DatePreV"
        Me.DatePreV.Size = New System.Drawing.Size(102, 30)
        Me.DatePreV.TabIndex = 46059
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(877, 158)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(83, 30)
        Me.TextBox1.TabIndex = 46060
        Me.TextBox1.Text = "0"
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TextBox1.Visible = False
        '
        'FrmRpt_Fixed_Assets_NEW
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1273, 423)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.DatePreV)
        Me.Controls.Add(Me.DTUse)
        Me.Controls.Add(Me.dpFromDate)
        Me.Controls.Add(Me.Button10)
        Me.Controls.Add(Me.Button9)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.txtBill_Amt)
        Me.Controls.Add(Me.txtBookName)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.CmbBook)
        Me.Controls.Add(Me.Cmb)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.TxtCrNm)
        Me.Controls.Add(Me.TxtDrNm)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.BtnSearch)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.TxtCr)
        Me.Controls.Add(Me.TxtDr)
        Me.Controls.Add(Me.TxtPeriod)
        Me.Controls.Add(Me.DateIn)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.TxtDesription)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.txtGrp)
        Me.Controls.Add(Me.txtGrpNm)
        Me.Controls.Add(Me.LEng)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.FG)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "FrmRpt_Fixed_Assets_NEW"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FrmRpt_Fixed_Assets_NEW"
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtCode As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtName As System.Windows.Forms.TextBox
    Friend WithEvents TxtNameE As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtValue As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtRemain As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents FG As System.Windows.Forms.DataGridView
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtGrp As System.Windows.Forms.TextBox
    Friend WithEvents txtGrpNm As System.Windows.Forms.ComboBox
    Friend WithEvents LEng As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TxtDesription As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents DateIn As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtPeriod As System.Windows.Forms.TextBox
    Friend WithEvents TxtDr As System.Windows.Forms.TextBox
    Friend WithEvents TxtCr As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents BtnSearch As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents TxtDrNm As System.Windows.Forms.TextBox
    Friend WithEvents TxtCrNm As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents CmbBook As System.Windows.Forms.ComboBox
    Friend WithEvents Cmb As System.Windows.Forms.ComboBox
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents txtBookName As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtBill_Amt As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents Button9 As System.Windows.Forms.Button
    Friend WithEvents Button10 As System.Windows.Forms.Button
    Friend WithEvents dpFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents DTUse As System.Windows.Forms.DateTimePicker
    Friend WithEvents DatePreV As System.Windows.Forms.DateTimePicker
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
End Class
