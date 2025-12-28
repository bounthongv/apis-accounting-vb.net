<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FmReceipt
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FmReceipt))
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.BtnExit = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label13 = New System.Windows.Forms.Label
        Me.ComboBox3 = New System.Windows.Forms.ComboBox
        Me.FG2 = New AxVSFlex8U.AxVSFlexGrid
        Me.FG1 = New AxVSFlex8U.AxVSFlexGrid
        Me.BtnPreview = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.BackPage = New System.Windows.Forms.Button
        Me.FirstPage = New System.Windows.Forms.Button
        Me.LasthPage = New System.Windows.Forms.Button
        Me.NextPage = New System.Windows.Forms.Button
        Me.lblpage_total = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Cmb = New System.Windows.Forms.ComboBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.Label28 = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.TextBox26 = New System.Windows.Forms.TextBox
        Me.BtnAddNew = New System.Windows.Forms.Button
        Me.BtnDelete = New System.Windows.Forms.Button
        Me.TextBox25 = New System.Windows.Forms.TextBox
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.Button4 = New System.Windows.Forms.Button
        Me.BtnRefresh = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TextBox11 = New System.Windows.Forms.TextBox
        Me.Remark = New System.Windows.Forms.TextBox
        Me.Label122 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.FGPaper = New AxVSFlex8U.AxVSFlexGrid
        Me.Label8 = New System.Windows.Forms.Label
        Me.Unit = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.Rate = New System.Windows.Forms.TextBox
        Me.Curr = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtAmt_letter = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TotalLAK = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.FGRate = New AxVSFlex8U.AxVSFlexGrid
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.Cashier = New System.Windows.Forms.TextBox
        Me.Payment = New System.Windows.Forms.TextBox
        Me.Label118 = New System.Windows.Forms.Label
        Me.Receipt_No = New System.Windows.Forms.TextBox
        Me.Label117 = New System.Windows.Forms.Label
        Me.Indate = New System.Windows.Forms.DateTimePicker
        Me.Bnk_Ac_Name = New System.Windows.Forms.TextBox
        Me.Bnk_Ac_Code = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.TotelAmt = New System.Windows.Forms.TextBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.Label121 = New System.Windows.Forms.Label
        Me.Label123 = New System.Windows.Forms.Label
        Me.Button15 = New System.Windows.Forms.Button
        Me.Button17 = New System.Windows.Forms.Button
        Me.Button18 = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.FG2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FG1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.FGPaper, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.FGRate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Alignment = System.Windows.Forms.TabAlignment.Bottom
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(0, -5)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1050, 592)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.TabPage1.Controls.Add(Me.BtnExit)
        Me.TabPage1.Controls.Add(Me.Button2)
        Me.TabPage1.Controls.Add(Me.Label13)
        Me.TabPage1.Controls.Add(Me.ComboBox3)
        Me.TabPage1.Controls.Add(Me.FG2)
        Me.TabPage1.Controls.Add(Me.FG1)
        Me.TabPage1.Controls.Add(Me.BtnPreview)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.TextBox2)
        Me.TabPage1.Controls.Add(Me.BackPage)
        Me.TabPage1.Controls.Add(Me.FirstPage)
        Me.TabPage1.Controls.Add(Me.LasthPage)
        Me.TabPage1.Controls.Add(Me.NextPage)
        Me.TabPage1.Controls.Add(Me.lblpage_total)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.Cmb)
        Me.TabPage1.Controls.Add(Me.Label24)
        Me.TabPage1.Controls.Add(Me.Label28)
        Me.TabPage1.Controls.Add(Me.Label23)
        Me.TabPage1.Controls.Add(Me.Label21)
        Me.TabPage1.Controls.Add(Me.TextBox26)
        Me.TabPage1.Controls.Add(Me.BtnAddNew)
        Me.TabPage1.Controls.Add(Me.BtnDelete)
        Me.TabPage1.Controls.Add(Me.TextBox25)
        Me.TabPage1.Controls.Add(Me.DateTimePicker2)
        Me.TabPage1.Controls.Add(Me.DateTimePicker1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 4)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1042, 555)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "   ລາຍການ    "
        '
        'BtnExit
        '
        Me.BtnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnExit.Location = New System.Drawing.Point(3, 5)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(35, 35)
        Me.BtnExit.TabIndex = 292
        Me.BtnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(359, 6)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(100, 35)
        Me.Button2.TabIndex = 281
        Me.Button2.Text = "ເອີ້ນຂໍ້ມູນ"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(436, 82)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(94, 24)
        Me.Label13.TabIndex = 280
        Me.Label13.Text = "ສະຖານະພາບ"
        '
        'ComboBox3
        '
        Me.ComboBox3.FormattingEnabled = True
        Me.ComboBox3.Items.AddRange(New Object() {"===ທັງຫມົດ===", "ຍັງບໍ່ທັນເຄື່ອນໄຫວ", "ເຄື່ອນໄຫວແລ້ວ"})
        Me.ComboBox3.Location = New System.Drawing.Point(536, 78)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(252, 32)
        Me.ComboBox3.TabIndex = 279
        '
        'FG2
        '
        Me.FG2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.FG2.DataSource = Nothing
        Me.FG2.Location = New System.Drawing.Point(7, 112)
        Me.FG2.Name = "FG2"
        Me.FG2.OcxState = CType(resources.GetObject("FG2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FG2.Size = New System.Drawing.Size(238, 400)
        Me.FG2.TabIndex = 276
        '
        'FG1
        '
        Me.FG1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FG1.DataSource = Nothing
        Me.FG1.Location = New System.Drawing.Point(251, 113)
        Me.FG1.Name = "FG1"
        Me.FG1.OcxState = CType(resources.GetObject("FG1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FG1.Size = New System.Drawing.Size(784, 403)
        Me.FG1.TabIndex = 275
        '
        'BtnPreview
        '
        Me.BtnPreview.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnPreview.Location = New System.Drawing.Point(253, 6)
        Me.BtnPreview.Name = "BtnPreview"
        Me.BtnPreview.Size = New System.Drawing.Size(107, 35)
        Me.BtnPreview.TabIndex = 274
        Me.BtnPreview.Text = "ວິວ/ເບິ່ງ"
        Me.BtnPreview.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(26, 520)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(47, 24)
        Me.Label3.TabIndex = 182
        Me.Label3.Text = "ມູນຄ່າ"
        '
        'TextBox2
        '
        Me.TextBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TextBox2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox2.Location = New System.Drawing.Point(79, 518)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(166, 26)
        Me.TextBox2.TabIndex = 212
        '
        'BackPage
        '
        Me.BackPage.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BackPage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.BackPage.ForeColor = System.Drawing.Color.Blue
        Me.BackPage.Location = New System.Drawing.Point(289, 521)
        Me.BackPage.Name = "BackPage"
        Me.BackPage.Size = New System.Drawing.Size(37, 23)
        Me.BackPage.TabIndex = 210
        Me.BackPage.Text = "<<"
        Me.BackPage.UseVisualStyleBackColor = True
        '
        'FirstPage
        '
        Me.FirstPage.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.FirstPage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.FirstPage.ForeColor = System.Drawing.Color.Blue
        Me.FirstPage.Location = New System.Drawing.Point(251, 521)
        Me.FirstPage.Name = "FirstPage"
        Me.FirstPage.Size = New System.Drawing.Size(39, 23)
        Me.FirstPage.TabIndex = 209
        Me.FirstPage.Text = "|<<"
        Me.FirstPage.UseVisualStyleBackColor = True
        '
        'LasthPage
        '
        Me.LasthPage.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LasthPage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.LasthPage.ForeColor = System.Drawing.Color.Blue
        Me.LasthPage.Location = New System.Drawing.Point(443, 522)
        Me.LasthPage.Name = "LasthPage"
        Me.LasthPage.Size = New System.Drawing.Size(38, 23)
        Me.LasthPage.TabIndex = 208
        Me.LasthPage.Text = ">>|"
        Me.LasthPage.UseVisualStyleBackColor = True
        '
        'NextPage
        '
        Me.NextPage.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.NextPage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NextPage.ForeColor = System.Drawing.Color.Blue
        Me.NextPage.Location = New System.Drawing.Point(406, 522)
        Me.NextPage.Name = "NextPage"
        Me.NextPage.Size = New System.Drawing.Size(37, 23)
        Me.NextPage.TabIndex = 207
        Me.NextPage.Text = ">>"
        Me.NextPage.UseVisualStyleBackColor = True
        '
        'lblpage_total
        '
        Me.lblpage_total.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblpage_total.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblpage_total.ForeColor = System.Drawing.Color.Blue
        Me.lblpage_total.Location = New System.Drawing.Point(326, 522)
        Me.lblpage_total.Name = "lblpage_total"
        Me.lblpage_total.ReadOnly = True
        Me.lblpage_total.Size = New System.Drawing.Size(79, 20)
        Me.lblpage_total.TabIndex = 206
        Me.lblpage_total.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(642, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 24)
        Me.Label2.TabIndex = 39
        Me.Label2.Text = "ສະກຸນເງິນ"
        '
        'Cmb
        '
        Me.Cmb.FormattingEnabled = True
        Me.Cmb.Location = New System.Drawing.Point(721, 38)
        Me.Cmb.Name = "Cmb"
        Me.Cmb.Size = New System.Drawing.Size(67, 32)
        Me.Cmb.TabIndex = 38
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(471, 42)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(28, 24)
        Me.Label24.TabIndex = 25
        Me.Label24.Text = "ຫາ"
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(44, 75)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(20, 24)
        Me.Label28.TabIndex = 24
        Me.Label28.Text = "ຊື່"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(232, 41)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(58, 24)
        Me.Label23.TabIndex = 23
        Me.Label23.Text = "ແຕ່ວັນທີ"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(9, 40)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(65, 24)
        Me.Label21.TabIndex = 22
        Me.Label21.Text = "ຮັບເລກທີ"
        '
        'TextBox26
        '
        Me.TextBox26.Location = New System.Drawing.Point(79, 72)
        Me.TextBox26.Name = "TextBox26"
        Me.TextBox26.Size = New System.Drawing.Size(344, 34)
        Me.TextBox26.TabIndex = 21
        '
        'BtnAddNew
        '
        Me.BtnAddNew.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAddNew.Location = New System.Drawing.Point(38, 6)
        Me.BtnAddNew.Name = "BtnAddNew"
        Me.BtnAddNew.Size = New System.Drawing.Size(114, 35)
        Me.BtnAddNew.TabIndex = 18
        Me.BtnAddNew.Text = "ເພີ່ມໃຫມ່"
        Me.BtnAddNew.UseVisualStyleBackColor = True
        '
        'BtnDelete
        '
        Me.BtnDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnDelete.Location = New System.Drawing.Point(152, 6)
        Me.BtnDelete.Name = "BtnDelete"
        Me.BtnDelete.Size = New System.Drawing.Size(100, 35)
        Me.BtnDelete.TabIndex = 14
        Me.BtnDelete.Text = "ລຶບ"
        Me.BtnDelete.UseVisualStyleBackColor = True
        '
        'TextBox25
        '
        Me.TextBox25.Location = New System.Drawing.Point(79, 36)
        Me.TextBox25.Name = "TextBox25"
        Me.TextBox25.Size = New System.Drawing.Size(144, 34)
        Me.TextBox25.TabIndex = 4
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker2.Location = New System.Drawing.Point(505, 37)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(131, 34)
        Me.DateTimePicker2.TabIndex = 3
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(293, 36)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(130, 34)
        Me.DateTimePicker1.TabIndex = 2
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.TabPage2.Controls.Add(Me.Button4)
        Me.TabPage2.Controls.Add(Me.BtnRefresh)
        Me.TabPage2.Controls.Add(Me.Button1)
        Me.TabPage2.Controls.Add(Me.Label6)
        Me.TabPage2.Controls.Add(Me.Panel2)
        Me.TabPage2.Controls.Add(Me.Panel1)
        Me.TabPage2.Controls.Add(Me.ComboBox1)
        Me.TabPage2.Controls.Add(Me.Label5)
        Me.TabPage2.Controls.Add(Me.TextBox11)
        Me.TabPage2.Controls.Add(Me.Remark)
        Me.TabPage2.Controls.Add(Me.Label122)
        Me.TabPage2.Controls.Add(Me.GroupBox2)
        Me.TabPage2.Controls.Add(Me.TotalLAK)
        Me.TabPage2.Controls.Add(Me.GroupBox1)
        Me.TabPage2.Controls.Add(Me.Label9)
        Me.TabPage2.Controls.Add(Me.Label10)
        Me.TabPage2.Controls.Add(Me.Button3)
        Me.TabPage2.Controls.Add(Me.Cashier)
        Me.TabPage2.Controls.Add(Me.Payment)
        Me.TabPage2.Controls.Add(Me.Label118)
        Me.TabPage2.Controls.Add(Me.Receipt_No)
        Me.TabPage2.Controls.Add(Me.Label117)
        Me.TabPage2.Controls.Add(Me.Indate)
        Me.TabPage2.Controls.Add(Me.Bnk_Ac_Name)
        Me.TabPage2.Controls.Add(Me.Bnk_Ac_Code)
        Me.TabPage2.Controls.Add(Me.Label25)
        Me.TabPage2.Controls.Add(Me.TotelAmt)
        Me.TabPage2.Controls.Add(Me.Label26)
        Me.TabPage2.Controls.Add(Me.Label121)
        Me.TabPage2.Controls.Add(Me.Label123)
        Me.TabPage2.Controls.Add(Me.Button15)
        Me.TabPage2.Controls.Add(Me.Button17)
        Me.TabPage2.Controls.Add(Me.Button18)
        Me.TabPage2.Controls.Add(Me.Label4)
        Me.TabPage2.Location = New System.Drawing.Point(4, 4)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(1042, 555)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = " ເພີ່ມລາຍການ "
        '
        'Button4
        '
        Me.Button4.Image = CType(resources.GetObject("Button4.Image"), System.Drawing.Image)
        Me.Button4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button4.Location = New System.Drawing.Point(338, 2)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(107, 34)
        Me.Button4.TabIndex = 281
        Me.Button4.Text = "ວິວ/ເບິ່ງ"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'BtnRefresh
        '
        Me.BtnRefresh.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnRefresh.Image = CType(resources.GetObject("BtnRefresh.Image"), System.Drawing.Image)
        Me.BtnRefresh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnRefresh.Location = New System.Drawing.Point(239, 2)
        Me.BtnRefresh.Name = "BtnRefresh"
        Me.BtnRefresh.Size = New System.Drawing.Size(100, 34)
        Me.BtnRefresh.TabIndex = 280
        Me.BtnRefresh.Text = "ເອີ້ນຂໍ້ມູນ"
        Me.BtnRefresh.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnRefresh.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(380, -26)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 279
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(257, 153)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(86, 21)
        Me.Label6.TabIndex = 271
        Me.Label6.Text = "ເງິນທີ່ຕ້ອງຈ່າຍ"
        '
        'Panel2
        '
        Me.Panel2.Location = New System.Drawing.Point(251, 117)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(28, 64)
        Me.Panel2.TabIndex = 278
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.RadioButton1)
        Me.Panel1.Controls.Add(Me.RadioButton2)
        Me.Panel1.Location = New System.Drawing.Point(116, 117)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(138, 64)
        Me.Panel1.TabIndex = 277
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.RadioButton1.Location = New System.Drawing.Point(7, -1)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(76, 28)
        Me.RadioButton1.TabIndex = 275
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "ແບບພີມ"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.RadioButton2.Location = New System.Drawing.Point(6, 26)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(132, 28)
        Me.RadioButton2.TabIndex = 276
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "ແບບອັດຕາໂນມັດ"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'ComboBox1
        '
        Me.ComboBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(349, 117)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(408, 32)
        Me.ComboBox1.TabIndex = 273
        Me.ComboBox1.Tag = "dfgdfg"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(288, 120)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(55, 24)
        Me.Label5.TabIndex = 274
        Me.Label5.Text = "ປະເພດ"
        '
        'TextBox11
        '
        Me.TextBox11.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox11.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox11.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.TextBox11.Location = New System.Drawing.Point(349, 151)
        Me.TextBox11.Name = "TextBox11"
        Me.TextBox11.Size = New System.Drawing.Size(408, 29)
        Me.TextBox11.TabIndex = 268
        Me.TextBox11.Text = "0.00"
        Me.TextBox11.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Remark
        '
        Me.Remark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Remark.Location = New System.Drawing.Point(114, 183)
        Me.Remark.Name = "Remark"
        Me.Remark.Size = New System.Drawing.Size(903, 34)
        Me.Remark.TabIndex = 270
        '
        'Label122
        '
        Me.Label122.AutoSize = True
        Me.Label122.Location = New System.Drawing.Point(1, 186)
        Me.Label122.Name = "Label122"
        Me.Label122.Size = New System.Drawing.Size(109, 24)
        Me.Label122.TabIndex = 271
        Me.Label122.Text = "ເນື້ອໃນລາຍການ"
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.TextBox1)
        Me.GroupBox2.Controls.Add(Me.Label14)
        Me.GroupBox2.Controls.Add(Me.FGPaper)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.Unit)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.Rate)
        Me.GroupBox2.Controls.Add(Me.Curr)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Controls.Add(Me.txtAmt_letter)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Location = New System.Drawing.Point(3, 213)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(830, 331)
        Me.GroupBox2.TabIndex = 269
        Me.GroupBox2.TabStop = False
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox1.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(656, 256)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(168, 29)
        Me.TextBox1.TabIndex = 269
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label14
        '
        Me.Label14.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(577, 257)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(75, 24)
        Me.Label14.TabIndex = 268
        Me.Label14.Text = "ຍອດເຫລືອ"
        '
        'FGPaper
        '
        Me.FGPaper.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FGPaper.DataSource = Nothing
        Me.FGPaper.Location = New System.Drawing.Point(3, 15)
        Me.FGPaper.Name = "FGPaper"
        Me.FGPaper.OcxState = CType(resources.GetObject("FGPaper.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FGPaper.Size = New System.Drawing.Size(784, 342)
        Me.FGPaper.TabIndex = 267
        '
        'Label8
        '
        Me.Label8.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label8.AutoSize = True
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(33, 257)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(74, 24)
        Me.Label8.TabIndex = 252
        Me.Label8.Text = "ຈຳນວນໃບ"
        '
        'Unit
        '
        Me.Unit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Unit.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.Unit.Location = New System.Drawing.Point(120, 256)
        Me.Unit.Name = "Unit"
        Me.Unit.ReadOnly = True
        Me.Unit.Size = New System.Drawing.Size(71, 29)
        Me.Unit.TabIndex = 251
        Me.Unit.Text = "0.00"
        Me.Unit.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label12
        '
        Me.Label12.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(373, 257)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(103, 24)
        Me.Label12.TabIndex = 265
        Me.Label12.Text = "ອັດຕາແລກປ່ຽນ"
        '
        'Rate
        '
        Me.Rate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Rate.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.Rate.Location = New System.Drawing.Point(482, 256)
        Me.Rate.Name = "Rate"
        Me.Rate.Size = New System.Drawing.Size(89, 29)
        Me.Rate.TabIndex = 263
        Me.Rate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Curr
        '
        Me.Curr.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Curr.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.Curr.Location = New System.Drawing.Point(283, 258)
        Me.Curr.Name = "Curr"
        Me.Curr.Size = New System.Drawing.Size(89, 26)
        Me.Curr.TabIndex = 266
        Me.Curr.Text = "LAK"
        Me.Curr.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label11
        '
        Me.Label11.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(207, 256)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(73, 24)
        Me.Label11.TabIndex = 264
        Me.Label11.Text = "ສະກຸນເງິນ"
        '
        'txtAmt_letter
        '
        Me.txtAmt_letter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAmt_letter.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtAmt_letter.ForeColor = System.Drawing.Color.Blue
        Me.txtAmt_letter.Location = New System.Drawing.Point(120, 287)
        Me.txtAmt_letter.Name = "txtAmt_letter"
        Me.txtAmt_letter.ReadOnly = True
        Me.txtAmt_letter.Size = New System.Drawing.Size(704, 34)
        Me.txtAmt_letter.TabIndex = 266
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(0, 289)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(117, 24)
        Me.Label1.TabIndex = 267
        Me.Label1.Text = "ເງິນເປັນຕົວຫນັງສື"
        '
        'TotalLAK
        '
        Me.TotalLAK.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TotalLAK.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TotalLAK.Location = New System.Drawing.Point(833, 147)
        Me.TotalLAK.Name = "TotalLAK"
        Me.TotalLAK.ReadOnly = True
        Me.TotalLAK.Size = New System.Drawing.Size(184, 29)
        Me.TotalLAK.TabIndex = 268
        Me.TotalLAK.Text = "0.00"
        Me.TotalLAK.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.FGRate)
        Me.GroupBox1.Location = New System.Drawing.Point(836, 213)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(181, 331)
        Me.GroupBox1.TabIndex = 265
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "ຕາຕະລາງອັດຕາແລກປ່ຽນ"
        '
        'FGRate
        '
        Me.FGRate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom), System.Windows.Forms.AnchorStyles)
        Me.FGRate.DataSource = Nothing
        Me.FGRate.Location = New System.Drawing.Point(4, 16)
        Me.FGRate.Name = "FGRate"
        Me.FGRate.OcxState = CType(resources.GetObject("FGRate.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FGRate.Size = New System.Drawing.Size(190, 411)
        Me.FGRate.TabIndex = 0
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(52, 86)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(57, 24)
        Me.Label9.TabIndex = 261
        Me.Label9.Text = "ລົງວັນທີ"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(10, 52)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(99, 24)
        Me.Label10.TabIndex = 260
        Me.Label10.Text = "ເລກບິນຮັບເງິນ"
        '
        'Button3
        '
        Me.Button3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button3.Location = New System.Drawing.Point(719, 45)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(37, 35)
        Me.Button3.TabIndex = 230
        Me.Button3.Text = "...."
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Cashier
        '
        Me.Cashier.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Cashier.Location = New System.Drawing.Point(833, 72)
        Me.Cashier.Name = "Cashier"
        Me.Cashier.Size = New System.Drawing.Size(184, 34)
        Me.Cashier.TabIndex = 229
        '
        'Payment
        '
        Me.Payment.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Payment.Location = New System.Drawing.Point(833, 35)
        Me.Payment.Name = "Payment"
        Me.Payment.Size = New System.Drawing.Size(184, 34)
        Me.Payment.TabIndex = 228
        '
        'Label118
        '
        Me.Label118.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label118.AutoSize = True
        Me.Label118.Location = New System.Drawing.Point(784, 75)
        Me.Label118.Name = "Label118"
        Me.Label118.Size = New System.Drawing.Size(39, 24)
        Me.Label118.TabIndex = 189
        Me.Label118.Text = "ຜູ້ຮັບ"
        '
        'Receipt_No
        '
        Me.Receipt_No.Location = New System.Drawing.Point(115, 45)
        Me.Receipt_No.Name = "Receipt_No"
        Me.Receipt_No.ReadOnly = True
        Me.Receipt_No.Size = New System.Drawing.Size(139, 34)
        Me.Receipt_No.TabIndex = 227
        '
        'Label117
        '
        Me.Label117.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label117.AutoSize = True
        Me.Label117.Location = New System.Drawing.Point(777, 38)
        Me.Label117.Name = "Label117"
        Me.Label117.Size = New System.Drawing.Size(46, 24)
        Me.Label117.TabIndex = 188
        Me.Label117.Text = "ຜູ້ຈ່າຍ"
        '
        'Indate
        '
        Me.Indate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.Indate.Location = New System.Drawing.Point(115, 81)
        Me.Indate.Name = "Indate"
        Me.Indate.Size = New System.Drawing.Size(139, 34)
        Me.Indate.TabIndex = 226
        '
        'Bnk_Ac_Name
        '
        Me.Bnk_Ac_Name.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Bnk_Ac_Name.Location = New System.Drawing.Point(349, 81)
        Me.Bnk_Ac_Name.Name = "Bnk_Ac_Name"
        Me.Bnk_Ac_Name.Size = New System.Drawing.Size(407, 34)
        Me.Bnk_Ac_Name.TabIndex = 187
        '
        'Bnk_Ac_Code
        '
        Me.Bnk_Ac_Code.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Bnk_Ac_Code.Location = New System.Drawing.Point(349, 45)
        Me.Bnk_Ac_Code.Name = "Bnk_Ac_Code"
        Me.Bnk_Ac_Code.ReadOnly = True
        Me.Bnk_Ac_Code.Size = New System.Drawing.Size(369, 34)
        Me.Bnk_Ac_Code.TabIndex = 186
        Me.Bnk_Ac_Code.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label25
        '
        Me.Label25.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label25.AutoSize = True
        Me.Label25.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label25.Location = New System.Drawing.Point(762, 150)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(66, 24)
        Me.Label25.TabIndex = 225
        Me.Label25.Text = "ທຽບ(ກີບ)"
        '
        'TotelAmt
        '
        Me.TotelAmt.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TotelAmt.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TotelAmt.Location = New System.Drawing.Point(833, 110)
        Me.TotelAmt.Name = "TotelAmt"
        Me.TotelAmt.ReadOnly = True
        Me.TotelAmt.Size = New System.Drawing.Size(184, 29)
        Me.TotelAmt.TabIndex = 182
        Me.TotelAmt.Text = "0.00"
        Me.TotelAmt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label26
        '
        Me.Label26.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label26.AutoSize = True
        Me.Label26.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label26.Location = New System.Drawing.Point(777, 113)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(47, 24)
        Me.Label26.TabIndex = 181
        Me.Label26.Text = "ມູນຄ່າ"
        '
        'Label121
        '
        Me.Label121.AutoSize = True
        Me.Label121.Location = New System.Drawing.Point(277, 55)
        Me.Label121.Name = "Label121"
        Me.Label121.Size = New System.Drawing.Size(66, 24)
        Me.Label121.TabIndex = 210
        Me.Label121.Text = "ເລກບັນຊີ"
        '
        'Label123
        '
        Me.Label123.AutoSize = True
        Me.Label123.Location = New System.Drawing.Point(260, 86)
        Me.Label123.Name = "Label123"
        Me.Label123.Size = New System.Drawing.Size(83, 24)
        Me.Label123.TabIndex = 207
        Me.Label123.Text = "ຊື່ຜູ້ມອບເງິນ"
        '
        'Button15
        '
        Me.Button15.Image = CType(resources.GetObject("Button15.Image"), System.Drawing.Image)
        Me.Button15.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button15.Location = New System.Drawing.Point(444, 2)
        Me.Button15.Name = "Button15"
        Me.Button15.Size = New System.Drawing.Size(114, 35)
        Me.Button15.TabIndex = 202
        Me.Button15.Text = "ອອກ"
        Me.Button15.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button15.UseVisualStyleBackColor = True
        '
        'Button17
        '
        Me.Button17.Image = CType(resources.GetObject("Button17.Image"), System.Drawing.Image)
        Me.Button17.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button17.Location = New System.Drawing.Point(125, 2)
        Me.Button17.Name = "Button17"
        Me.Button17.Size = New System.Drawing.Size(114, 35)
        Me.Button17.TabIndex = 200
        Me.Button17.Text = "ບັນທຶກ"
        Me.Button17.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button17.UseVisualStyleBackColor = True
        '
        'Button18
        '
        Me.Button18.Image = CType(resources.GetObject("Button18.Image"), System.Drawing.Image)
        Me.Button18.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button18.Location = New System.Drawing.Point(11, 2)
        Me.Button18.Name = "Button18"
        Me.Button18.Size = New System.Drawing.Size(114, 35)
        Me.Button18.TabIndex = 199
        Me.Button18.Text = "ເພີ່ມໃຫມ່"
        Me.Button18.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button18.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Saysettha OT", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Blue
        Me.Label4.Location = New System.Drawing.Point(571, 3)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(202, 34)
        Me.Label4.TabIndex = 272
        Me.Label4.Text = "ການມອບ-ຮັບເງິນສົດ"
        '
        'FmReceipt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1028, 590)
        Me.Controls.Add(Me.TabControl1)
        Me.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4, 6, 4, 6)
        Me.Name = "FmReceipt"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FmReceipt"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.FG2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FG1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.FGPaper, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.FGRate, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TextBox25 As System.Windows.Forms.TextBox
    Friend WithEvents DateTimePicker2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents TextBox26 As System.Windows.Forms.TextBox
    Friend WithEvents BtnAddNew As System.Windows.Forms.Button
    Friend WithEvents BtnDelete As System.Windows.Forms.Button
    Friend WithEvents Label121 As System.Windows.Forms.Label
    Friend WithEvents Label123 As System.Windows.Forms.Label
    Friend WithEvents Button15 As System.Windows.Forms.Button
    Friend WithEvents Button17 As System.Windows.Forms.Button
    Friend WithEvents Button18 As System.Windows.Forms.Button
    Friend WithEvents Receipt_No As System.Windows.Forms.TextBox
    Friend WithEvents Indate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Cashier As System.Windows.Forms.TextBox
    Friend WithEvents Payment As System.Windows.Forms.TextBox
    Friend WithEvents Label118 As System.Windows.Forms.Label
    Friend WithEvents Label117 As System.Windows.Forms.Label
    Friend WithEvents Bnk_Ac_Name As System.Windows.Forms.TextBox
    Friend WithEvents Bnk_Ac_Code As System.Windows.Forms.TextBox
    Friend WithEvents TotelAmt As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Unit As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Cmb As System.Windows.Forms.ComboBox
    Friend WithEvents Rate As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Curr As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents FGRate As AxVSFlex8U.AxVSFlexGrid
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtAmt_letter As System.Windows.Forms.TextBox
    Friend WithEvents TotalLAK As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Remark As System.Windows.Forms.TextBox
    Friend WithEvents Label122 As System.Windows.Forms.Label
    'Friend WithEvents CachedCryloanpayment1 As ApPBank10.CachedCryloanpayment
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents BackPage As System.Windows.Forms.Button
    Friend WithEvents FirstPage As System.Windows.Forms.Button
    Friend WithEvents LasthPage As System.Windows.Forms.Button
    Friend WithEvents NextPage As System.Windows.Forms.Button
    Friend WithEvents lblpage_total As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents BtnPreview As System.Windows.Forms.Button
    Friend WithEvents FG1 As AxVSFlex8U.AxVSFlexGrid
    Friend WithEvents FG2 As AxVSFlex8U.AxVSFlexGrid
    Friend WithEvents FGPaper As AxVSFlex8U.AxVSFlexGrid
    Friend WithEvents TextBox11 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents ComboBox3 As System.Windows.Forms.ComboBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents BtnRefresh As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents BtnExit As System.Windows.Forms.Button
End Class
