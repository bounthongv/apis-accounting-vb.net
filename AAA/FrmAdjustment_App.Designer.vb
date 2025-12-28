<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmAdjustment_App
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmAdjustment_App))
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
        Me.FG = New AxVSFlex8U.AxVSFlexGrid
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtGrp = New System.Windows.Forms.TextBox
        Me.txtGrpNm = New System.Windows.Forms.ComboBox
        Me.LEng = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.BtnAddNew2 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
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
        Me.txtRate = New System.Windows.Forms.TextBox
        Me.txtcurr_name2 = New System.Windows.Forms.TextBox
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
        Me.FG.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FG.DataSource = Nothing
        Me.FG.Location = New System.Drawing.Point(16, 126)
        Me.FG.Name = "FG"
        Me.FG.OcxState = CType(resources.GetObject("FG.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FG.Size = New System.Drawing.Size(1118, 292)
        Me.FG.TabIndex = 303
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Saysettha OT", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Blue
        Me.Label11.Location = New System.Drawing.Point(402, 3)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(350, 36)
        Me.Label11.TabIndex = 304
        Me.Label11.Text = "ຂໍ້ມູນການດັດປັບ"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtGrp
        '
        Me.txtGrp.Enabled = False
        Me.txtGrp.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGrp.Location = New System.Drawing.Point(300, 42)
        Me.txtGrp.Name = "txtGrp"
        Me.txtGrp.Size = New System.Drawing.Size(123, 30)
        Me.txtGrp.TabIndex = 307
        Me.txtGrp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtGrpNm
        '
        Me.txtGrpNm.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGrpNm.FormattingEnabled = True
        Me.txtGrpNm.Items.AddRange(New Object() {"LAK", "THB", "USD"})
        Me.txtGrpNm.Location = New System.Drawing.Point(174, 78)
        Me.txtGrpNm.Name = "txtGrpNm"
        Me.txtGrpNm.Size = New System.Drawing.Size(249, 32)
        Me.txtGrpNm.TabIndex = 306
        '
        'LEng
        '
        Me.LEng.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LEng.Location = New System.Drawing.Point(14, 81)
        Me.LEng.Name = "LEng"
        Me.LEng.Size = New System.Drawing.Size(151, 24)
        Me.LEng.TabIndex = 305
        Me.LEng.Tag = "2010"
        Me.LEng.Text = "Adjustment Type:"
        Me.LEng.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Image = CType(resources.GetObject("Button3.Image"), System.Drawing.Image)
        Me.Button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button3.Location = New System.Drawing.Point(993, 34)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(100, 35)
        Me.Button3.TabIndex = 282
        Me.Button3.Tag = "3004"
        Me.Button3.Text = "ລຶບ"
        Me.Button3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button3.UseVisualStyleBackColor = True
        Me.Button3.Visible = False
        '
        'BtnAddNew2
        '
        Me.BtnAddNew2.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAddNew2.Image = CType(resources.GetObject("BtnAddNew2.Image"), System.Drawing.Image)
        Me.BtnAddNew2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAddNew2.Location = New System.Drawing.Point(1008, 34)
        Me.BtnAddNew2.Name = "BtnAddNew2"
        Me.BtnAddNew2.Size = New System.Drawing.Size(114, 35)
        Me.BtnAddNew2.TabIndex = 281
        Me.BtnAddNew2.Tag = "3001"
        Me.BtnAddNew2.Text = "ເພີ່ມໃຫມ່"
        Me.BtnAddNew2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnAddNew2.UseVisualStyleBackColor = True
        Me.BtnAddNew2.Visible = False
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(1122, 34)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(114, 35)
        Me.Button2.TabIndex = 280
        Me.Button2.Tag = "3002"
        Me.Button2.Text = "ບັນທຶກ"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'Button1
        '
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(7, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(35, 35)
        Me.Button1.TabIndex = 279
        Me.Button1.Tag = "9999"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.UseVisualStyleBackColor = True
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
        Me.Label13.Location = New System.Drawing.Point(57, 45)
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
        Me.DateIn.Location = New System.Drawing.Point(174, 44)
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
        Me.TxtDr.Location = New System.Drawing.Point(805, 146)
        Me.TxtDr.Name = "TxtDr"
        Me.TxtDr.Size = New System.Drawing.Size(163, 30)
        Me.TxtDr.TabIndex = 314
        Me.TxtDr.Visible = False
        '
        'TxtCr
        '
        Me.TxtCr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtCr.Location = New System.Drawing.Point(805, 178)
        Me.TxtCr.Name = "TxtCr"
        Me.TxtCr.Size = New System.Drawing.Size(163, 30)
        Me.TxtCr.TabIndex = 315
        Me.TxtCr.Visible = False
        '
        'Label7
        '
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(709, 181)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(94, 27)
        Me.Label7.TabIndex = 316
        Me.Label7.Text = "Cr:"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label7.Visible = False
        '
        'BtnSearch
        '
        Me.BtnSearch.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSearch.Location = New System.Drawing.Point(968, 145)
        Me.BtnSearch.Name = "BtnSearch"
        Me.BtnSearch.Size = New System.Drawing.Size(34, 30)
        Me.BtnSearch.TabIndex = 46037
        Me.BtnSearch.Tag = "3012"
        Me.BtnSearch.Text = "....."
        Me.BtnSearch.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnSearch.UseVisualStyleBackColor = True
        Me.BtnSearch.Visible = False
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.Location = New System.Drawing.Point(968, 178)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(34, 30)
        Me.Button4.TabIndex = 46038
        Me.Button4.Tag = "3012"
        Me.Button4.Text = "....."
        Me.Button4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button4.UseVisualStyleBackColor = True
        Me.Button4.Visible = False
        '
        'TxtDrNm
        '
        Me.TxtDrNm.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtDrNm.Location = New System.Drawing.Point(1008, 145)
        Me.TxtDrNm.Name = "TxtDrNm"
        Me.TxtDrNm.Size = New System.Drawing.Size(392, 30)
        Me.TxtDrNm.TabIndex = 46039
        Me.TxtDrNm.Visible = False
        '
        'TxtCrNm
        '
        Me.TxtCrNm.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtCrNm.Location = New System.Drawing.Point(1008, 177)
        Me.TxtCrNm.Name = "TxtCrNm"
        Me.TxtCrNm.Size = New System.Drawing.Size(392, 30)
        Me.TxtCrNm.TabIndex = 46040
        Me.TxtCrNm.Visible = False
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox2.Location = New System.Drawing.Point(438, 46)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(141, 28)
        Me.CheckBox2.TabIndex = 46041
        Me.CheckBox2.Text = "Full Depreciation"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(437, 86)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(126, 28)
        Me.CheckBox1.TabIndex = 46042
        Me.CheckBox1.Text = "Days By date"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(578, 44)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(94, 27)
        Me.Label8.TabIndex = 46043
        Me.Label8.Text = "A/C Book:"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label10
        '
        Me.Label10.ForeColor = System.Drawing.Color.Black
        Me.Label10.Location = New System.Drawing.Point(578, 81)
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
        Me.CmbBook.Location = New System.Drawing.Point(678, 46)
        Me.CmbBook.Name = "CmbBook"
        Me.CmbBook.Size = New System.Drawing.Size(125, 30)
        Me.CmbBook.TabIndex = 46046
        '
        'Cmb
        '
        Me.Cmb.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cmb.FormattingEnabled = True
        Me.Cmb.Location = New System.Drawing.Point(678, 77)
        Me.Cmb.Name = "Cmb"
        Me.Cmb.Size = New System.Drawing.Size(86, 30)
        Me.Cmb.TabIndex = 46047
        '
        'Button6
        '
        Me.Button6.Font = New System.Drawing.Font("Saysettha OT", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button6.Location = New System.Drawing.Point(949, 9)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(156, 36)
        Me.Button6.TabIndex = 46048
        Me.Button6.Text = "approve"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Font = New System.Drawing.Font("Saysettha OT", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.Location = New System.Drawing.Point(174, 5)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(120, 36)
        Me.Button5.TabIndex = 46049
        Me.Button5.Text = "Find"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'txtBookName
        '
        Me.txtBookName.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtBookName.Location = New System.Drawing.Point(809, 46)
        Me.txtBookName.Name = "txtBookName"
        Me.txtBookName.ReadOnly = True
        Me.txtBookName.Size = New System.Drawing.Size(334, 30)
        Me.txtBookName.TabIndex = 46050
        '
        'txtRate
        '
        Me.txtRate.BackColor = System.Drawing.Color.White
        Me.txtRate.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRate.Location = New System.Drawing.Point(770, 78)
        Me.txtRate.Name = "txtRate"
        Me.txtRate.Size = New System.Drawing.Size(85, 29)
        Me.txtRate.TabIndex = 46051
        Me.txtRate.Text = "1"
        Me.txtRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtcurr_name2
        '
        Me.txtcurr_name2.BackColor = System.Drawing.Color.White
        Me.txtcurr_name2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcurr_name2.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcurr_name2.ForeColor = System.Drawing.Color.Blue
        Me.txtcurr_name2.Location = New System.Drawing.Point(758, 3)
        Me.txtcurr_name2.Name = "txtcurr_name2"
        Me.txtcurr_name2.Size = New System.Drawing.Size(100, 30)
        Me.txtcurr_name2.TabIndex = 46052
        Me.txtcurr_name2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtcurr_name2.Visible = False
        '
        'FrmAdjustment_App
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 21.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1146, 423)
        Me.Controls.Add(Me.txtcurr_name2)
        Me.Controls.Add(Me.txtRate)
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
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.TxtRemain)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TxtValue)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TxtNameE)
        Me.Controls.Add(Me.TxtName)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TxtCode)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.BtnAddNew2)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "FrmAdjustment_App"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FrmAdjustment_App"
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents BtnAddNew2 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
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
    Friend WithEvents FG As AxVSFlex8U.AxVSFlexGrid
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
    Friend WithEvents txtRate As System.Windows.Forms.TextBox
    Friend WithEvents txtcurr_name2 As System.Windows.Forms.TextBox
End Class
