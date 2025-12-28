<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Rate_setting
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Rate_setting))
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.o = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.p = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtTHB_LAK = New System.Windows.Forms.TextBox
        Me.txtUSD_LAK = New System.Windows.Forms.TextBox
        Me.txtEUR_THB = New System.Windows.Forms.TextBox
        Me.txtEUR_LAK = New System.Windows.Forms.TextBox
        Me.txtEUR_USD = New System.Windows.Forms.TextBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.txtUSD_THB = New System.Windows.Forms.TextBox
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.Label9 = New System.Windows.Forms.Label
        Me.BtnAddNew = New System.Windows.Forms.Button
        Me.BtnDel = New System.Windows.Forms.Button
        Me.txtCerrent = New System.Windows.Forms.TextBox
        Me.DTrate = New System.Windows.Forms.DateTimePicker
        Me.BtnExit = New System.Windows.Forms.Button
        Me.FG_Rate = New AxVSFlex8U.AxVSFlexGrid
        Me.Label6 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Label10 = New System.Windows.Forms.Label
        Me.Cmb_Component = New System.Windows.Forms.ComboBox
        Me.txt_Component_id = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.FG_Curr = New AxVSFlex8U.AxVSFlexGrid
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.txtRate = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Panel6 = New System.Windows.Forms.Panel
        Me.txtRate2 = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.txtcurr_name2 = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.BtnSave = New System.Windows.Forms.Button
        Me.CMB_Curr = New System.Windows.Forms.ComboBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.Panel7 = New System.Windows.Forms.Panel
        Me.txtcurr_name = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.txtCurr = New System.Windows.Forms.TextBox
        Me.BtnSave2 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.dpToDate = New System.Windows.Forms.DateTimePicker
        Me.dpFromDate = New System.Windows.Forms.DateTimePicker
        Me.Label15 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.CMB_Curr_SSS = New System.Windows.Forms.ComboBox
        Me.Panel1.SuspendLayout()
        Me.Panel4.SuspendLayout()
        CType(Me.FG_Rate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        CType(Me.FG_Curr, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel5.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.Panel7.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(4, 5)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(109, 25)
        Me.Label1.TabIndex = 49
        Me.Label1.Text = "ປະຈໍາວັນທີ່:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.FromArgb(CType(CType(187, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.Label7.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Blue
        Me.Label7.Location = New System.Drawing.Point(31, 84)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(94, 29)
        Me.Label7.TabIndex = 52
        Me.Label7.Text = "ຢູໂຣ-ກີບ"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label7.Visible = False
        '
        'o
        '
        Me.o.BackColor = System.Drawing.Color.FromArgb(CType(CType(187, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.o.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.o.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.o.ForeColor = System.Drawing.Color.Black
        Me.o.Location = New System.Drawing.Point(296, 158)
        Me.o.Name = "o"
        Me.o.Size = New System.Drawing.Size(93, 26)
        Me.o.TabIndex = 52
        Me.o.Text = "ຢູໂຣ-ໂດລາ"
        Me.o.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.FromArgb(CType(CType(187, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.Label3.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Blue
        Me.Label3.Location = New System.Drawing.Point(18, 99)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(218, 30)
        Me.Label3.TabIndex = 52
        Me.Label3.Text = "SDR-ໂດລາ/ SDR - DOLLAR :"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label3.Visible = False
        '
        'p
        '
        Me.p.BackColor = System.Drawing.Color.FromArgb(CType(CType(187, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.p.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.p.ForeColor = System.Drawing.Color.Blue
        Me.p.Location = New System.Drawing.Point(18, 68)
        Me.p.Name = "p"
        Me.p.Size = New System.Drawing.Size(218, 29)
        Me.p.TabIndex = 52
        Me.p.Text = "ກີບ-ໂດລາ / KIP - DOLLAR :"
        Me.p.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.FromArgb(CType(CType(187, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(84, 167)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 27)
        Me.Label4.TabIndex = 52
        Me.Label4.Text = "ຢູໂຣ-ບາດ"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.FromArgb(CType(CType(187, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.Label2.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Blue
        Me.Label2.Location = New System.Drawing.Point(18, 35)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(218, 30)
        Me.Label2.TabIndex = 51
        Me.Label2.Text = "ກີບ-THB / KIP - THB :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.FromArgb(CType(CType(187, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.Label8.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(265, 59)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(120, 25)
        Me.Label8.TabIndex = 111
        Me.Label8.Text = "ອັດຕາແລກປ່ຽນ"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.FromArgb(CType(CType(187, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.Label5.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(269, 82)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(116, 34)
        Me.Label5.TabIndex = 112
        Me.Label5.Text = "(Rate ExChange)"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtTHB_LAK
        '
        Me.txtTHB_LAK.BackColor = System.Drawing.Color.White
        Me.txtTHB_LAK.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTHB_LAK.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTHB_LAK.ForeColor = System.Drawing.Color.Blue
        Me.txtTHB_LAK.Location = New System.Drawing.Point(238, 35)
        Me.txtTHB_LAK.Name = "txtTHB_LAK"
        Me.txtTHB_LAK.Size = New System.Drawing.Size(162, 30)
        Me.txtTHB_LAK.TabIndex = 124
        Me.txtTHB_LAK.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtUSD_LAK
        '
        Me.txtUSD_LAK.BackColor = System.Drawing.Color.White
        Me.txtUSD_LAK.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUSD_LAK.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUSD_LAK.ForeColor = System.Drawing.Color.Blue
        Me.txtUSD_LAK.Location = New System.Drawing.Point(238, 67)
        Me.txtUSD_LAK.Name = "txtUSD_LAK"
        Me.txtUSD_LAK.Size = New System.Drawing.Size(162, 30)
        Me.txtUSD_LAK.TabIndex = 127
        Me.txtUSD_LAK.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtEUR_THB
        '
        Me.txtEUR_THB.BackColor = System.Drawing.Color.White
        Me.txtEUR_THB.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtEUR_THB.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEUR_THB.ForeColor = System.Drawing.Color.Blue
        Me.txtEUR_THB.Location = New System.Drawing.Point(84, 197)
        Me.txtEUR_THB.Name = "txtEUR_THB"
        Me.txtEUR_THB.Size = New System.Drawing.Size(100, 30)
        Me.txtEUR_THB.TabIndex = 126
        Me.txtEUR_THB.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtEUR_LAK
        '
        Me.txtEUR_LAK.BackColor = System.Drawing.Color.White
        Me.txtEUR_LAK.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtEUR_LAK.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEUR_LAK.ForeColor = System.Drawing.Color.Blue
        Me.txtEUR_LAK.Location = New System.Drawing.Point(128, 86)
        Me.txtEUR_LAK.Name = "txtEUR_LAK"
        Me.txtEUR_LAK.Size = New System.Drawing.Size(162, 30)
        Me.txtEUR_LAK.TabIndex = 129
        Me.txtEUR_LAK.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtEUR_LAK.Visible = False
        '
        'txtEUR_USD
        '
        Me.txtEUR_USD.BackColor = System.Drawing.Color.White
        Me.txtEUR_USD.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtEUR_USD.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEUR_USD.ForeColor = System.Drawing.Color.Blue
        Me.txtEUR_USD.Location = New System.Drawing.Point(296, 183)
        Me.txtEUR_USD.Name = "txtEUR_USD"
        Me.txtEUR_USD.Size = New System.Drawing.Size(93, 30)
        Me.txtEUR_USD.TabIndex = 128
        Me.txtEUR_USD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(187, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.txtUSD_THB)
        Me.Panel1.Controls.Add(Me.Panel4)
        Me.Panel1.Controls.Add(Me.txtTHB_LAK)
        Me.Panel1.Controls.Add(Me.txtUSD_LAK)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.p)
        Me.Panel1.Location = New System.Drawing.Point(1086, 46)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(235, 152)
        Me.Panel1.TabIndex = 130
        Me.Panel1.Visible = False
        '
        'txtUSD_THB
        '
        Me.txtUSD_THB.BackColor = System.Drawing.Color.White
        Me.txtUSD_THB.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUSD_THB.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUSD_THB.ForeColor = System.Drawing.Color.Blue
        Me.txtUSD_THB.Location = New System.Drawing.Point(239, 99)
        Me.txtUSD_THB.Name = "txtUSD_THB"
        Me.txtUSD_THB.Size = New System.Drawing.Size(161, 30)
        Me.txtUSD_THB.TabIndex = 132
        Me.txtUSD_THB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtUSD_THB.Visible = False
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.Label9)
        Me.Panel4.Controls.Add(Me.o)
        Me.Panel4.Controls.Add(Me.BtnAddNew)
        Me.Panel4.Controls.Add(Me.Label4)
        Me.Panel4.Controls.Add(Me.BtnDel)
        Me.Panel4.Controls.Add(Me.Label7)
        Me.Panel4.Controls.Add(Me.txtEUR_THB)
        Me.Panel4.Controls.Add(Me.txtEUR_USD)
        Me.Panel4.Controls.Add(Me.txtCerrent)
        Me.Panel4.Controls.Add(Me.Label5)
        Me.Panel4.Controls.Add(Me.Label8)
        Me.Panel4.Controls.Add(Me.txtEUR_LAK)
        Me.Panel4.Location = New System.Drawing.Point(444, 35)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(116, 94)
        Me.Panel4.TabIndex = 46017
        Me.Panel4.Visible = False
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.FromArgb(CType(CType(187, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.Label9.Font = New System.Drawing.Font("Saysettha OT", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Red
        Me.Label9.Location = New System.Drawing.Point(19, 37)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(271, 34)
        Me.Label9.TabIndex = 134
        Me.Label9.Text = "(Rate ExChange)"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BtnAddNew
        '
        Me.BtnAddNew.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAddNew.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAddNew.Location = New System.Drawing.Point(190, 154)
        Me.BtnAddNew.Name = "BtnAddNew"
        Me.BtnAddNew.Size = New System.Drawing.Size(100, 30)
        Me.BtnAddNew.TabIndex = 40
        Me.BtnAddNew.Text = "ເພີ່ມໃໝ່"
        Me.BtnAddNew.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnAddNew.UseVisualStyleBackColor = True
        '
        'BtnDel
        '
        Me.BtnDel.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnDel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnDel.Location = New System.Drawing.Point(210, 118)
        Me.BtnDel.Name = "BtnDel"
        Me.BtnDel.Size = New System.Drawing.Size(100, 30)
        Me.BtnDel.TabIndex = 42
        Me.BtnDel.Text = "ລຶບ"
        Me.BtnDel.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BtnDel.UseVisualStyleBackColor = True
        '
        'txtCerrent
        '
        Me.txtCerrent.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCerrent.Location = New System.Drawing.Point(84, 134)
        Me.txtCerrent.Name = "txtCerrent"
        Me.txtCerrent.Size = New System.Drawing.Size(100, 30)
        Me.txtCerrent.TabIndex = 130
        '
        'DTrate
        '
        Me.DTrate.CalendarFont = New System.Drawing.Font("Saysettha OT", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DTrate.CustomFormat = "dd/MM/yyyy"
        Me.DTrate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTrate.Location = New System.Drawing.Point(116, 3)
        Me.DTrate.Name = "DTrate"
        Me.DTrate.ShowUpDown = True
        Me.DTrate.Size = New System.Drawing.Size(110, 30)
        Me.DTrate.TabIndex = 131
        Me.DTrate.Value = New Date(2009, 12, 31, 0, 0, 0, 0)
        '
        'BtnExit
        '
        Me.BtnExit.Image = CType(resources.GetObject("BtnExit.Image"), System.Drawing.Image)
        Me.BtnExit.Location = New System.Drawing.Point(8, 7)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(32, 30)
        Me.BtnExit.TabIndex = 108
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'FG_Rate
        '
        Me.FG_Rate.DataSource = Nothing
        Me.FG_Rate.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FG_Rate.Location = New System.Drawing.Point(0, 0)
        Me.FG_Rate.Name = "FG_Rate"
        Me.FG_Rate.OcxState = CType(resources.GetObject("FG_Rate.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FG_Rate.Size = New System.Drawing.Size(588, 331)
        Me.FG_Rate.TabIndex = 52
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Font = New System.Drawing.Font("Saysettha OT", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Red
        Me.Label6.Location = New System.Drawing.Point(88, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(185, 40)
        Me.Label6.TabIndex = 133
        Me.Label6.Text = "ເພີ່ມສະກຸນເງິນ"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.RoyalBlue
        Me.Panel2.Controls.Add(Me.BtnExit)
        Me.Panel2.Controls.Add(Me.Label10)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1081, 40)
        Me.Panel2.TabIndex = 135
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.Blue
        Me.Label10.Font = New System.Drawing.Font("Times New Roman", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.White
        Me.Label10.Location = New System.Drawing.Point(6, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(891, 40)
        Me.Label10.TabIndex = 136
        Me.Label10.Text = "Exchange Rate"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Cmb_Component
        '
        Me.Cmb_Component.Font = New System.Drawing.Font("Saysettha OT", 9.75!)
        Me.Cmb_Component.FormattingEnabled = True
        Me.Cmb_Component.Location = New System.Drawing.Point(172, 46)
        Me.Cmb_Component.Name = "Cmb_Component"
        Me.Cmb_Component.Size = New System.Drawing.Size(424, 32)
        Me.Cmb_Component.TabIndex = 46014
        Me.Cmb_Component.Visible = False
        '
        'txt_Component_id
        '
        Me.txt_Component_id.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_Component_id.ForeColor = System.Drawing.Color.Black
        Me.txt_Component_id.Location = New System.Drawing.Point(90, 45)
        Me.txt_Component_id.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_Component_id.Name = "txt_Component_id"
        Me.txt_Component_id.ReadOnly = True
        Me.txt_Component_id.Size = New System.Drawing.Size(80, 30)
        Me.txt_Component_id.TabIndex = 46013
        Me.txt_Component_id.Visible = False
        '
        'Label25
        '
        Me.Label25.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.ForeColor = System.Drawing.Color.Black
        Me.Label25.Location = New System.Drawing.Point(4, 45)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(80, 28)
        Me.Label25.TabIndex = 46012
        Me.Label25.Tag = "2007"
        Me.Label25.Text = "ໜ່ວຍງານ:"
        Me.Label25.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label25.Visible = False
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Panel3.Controls.Add(Me.FG_Rate)
        Me.Panel3.Location = New System.Drawing.Point(12, 228)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(588, 331)
        Me.Panel3.TabIndex = 46015
        '
        'FG_Curr
        '
        Me.FG_Curr.DataSource = Nothing
        Me.FG_Curr.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FG_Curr.Location = New System.Drawing.Point(0, 0)
        Me.FG_Curr.Name = "FG_Curr"
        Me.FG_Curr.OcxState = CType(resources.GetObject("FG_Curr.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FG_Curr.Size = New System.Drawing.Size(442, 362)
        Me.FG_Curr.TabIndex = 46016
        '
        'Panel5
        '
        Me.Panel5.Controls.Add(Me.FG_Curr)
        Me.Panel5.Location = New System.Drawing.Point(14, 111)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(442, 362)
        Me.Panel5.TabIndex = 46017
        '
        'txtRate
        '
        Me.txtRate.BackColor = System.Drawing.Color.White
        Me.txtRate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRate.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRate.ForeColor = System.Drawing.Color.Blue
        Me.txtRate.Location = New System.Drawing.Point(116, 65)
        Me.txtRate.Name = "txtRate"
        Me.txtRate.Size = New System.Drawing.Size(110, 30)
        Me.txtRate.TabIndex = 46019
        Me.txtRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Blue
        Me.Label11.Location = New System.Drawing.Point(-5, 65)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(118, 30)
        Me.Label11.TabIndex = 46018
        Me.Label11.Text = "ອັດຕາ / Rate :"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel6
        '
        Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel6.Controls.Add(Me.txtRate2)
        Me.Panel6.Controls.Add(Me.Label18)
        Me.Panel6.Controls.Add(Me.txtcurr_name2)
        Me.Panel6.Controls.Add(Me.Label17)
        Me.Panel6.Controls.Add(Me.Button3)
        Me.Panel6.Controls.Add(Me.BtnSave)
        Me.Panel6.Controls.Add(Me.CMB_Curr)
        Me.Panel6.Controls.Add(Me.Label12)
        Me.Panel6.Controls.Add(Me.DTrate)
        Me.Panel6.Controls.Add(Me.Label1)
        Me.Panel6.Controls.Add(Me.txtRate)
        Me.Panel6.Controls.Add(Me.Label11)
        Me.Panel6.Location = New System.Drawing.Point(12, 81)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(588, 107)
        Me.Panel6.TabIndex = 46020
        '
        'txtRate2
        '
        Me.txtRate2.BackColor = System.Drawing.Color.White
        Me.txtRate2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRate2.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRate2.ForeColor = System.Drawing.Color.Blue
        Me.txtRate2.Location = New System.Drawing.Point(351, 65)
        Me.txtRate2.Name = "txtRate2"
        Me.txtRate2.Size = New System.Drawing.Size(100, 30)
        Me.txtRate2.TabIndex = 46031
        Me.txtRate2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label18
        '
        Me.Label18.BackColor = System.Drawing.SystemColors.Control
        Me.Label18.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.Blue
        Me.Label18.Location = New System.Drawing.Point(228, 65)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(118, 30)
        Me.Label18.TabIndex = 46030
        Me.Label18.Text = "ອັດຕາ / Rate :"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtcurr_name2
        '
        Me.txtcurr_name2.BackColor = System.Drawing.Color.White
        Me.txtcurr_name2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcurr_name2.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcurr_name2.ForeColor = System.Drawing.Color.Blue
        Me.txtcurr_name2.Location = New System.Drawing.Point(351, 34)
        Me.txtcurr_name2.Name = "txtcurr_name2"
        Me.txtcurr_name2.Size = New System.Drawing.Size(100, 30)
        Me.txtcurr_name2.TabIndex = 46028
        Me.txtcurr_name2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label17.ForeColor = System.Drawing.Color.Blue
        Me.Label17.Location = New System.Drawing.Point(251, 37)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(100, 24)
        Me.Label17.TabIndex = 46029
        Me.Label17.Tag = "2023"
        Me.Label17.Text = "ເປັນພາສາລາວ :"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button3.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.Black
        Me.Button3.Image = CType(resources.GetObject("Button3.Image"), System.Drawing.Image)
        Me.Button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button3.Location = New System.Drawing.Point(351, 3)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(100, 30)
        Me.Button3.TabIndex = 46026
        Me.Button3.Tag = "3006"
        Me.Button3.Text = "ລືບ"
        Me.Button3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button3.UseVisualStyleBackColor = False
        '
        'BtnSave
        '
        Me.BtnSave.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.BtnSave.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.BtnSave.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSave.Image = CType(resources.GetObject("BtnSave.Image"), System.Drawing.Image)
        Me.BtnSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnSave.Location = New System.Drawing.Point(232, 3)
        Me.BtnSave.Name = "BtnSave"
        Me.BtnSave.Size = New System.Drawing.Size(113, 30)
        Me.BtnSave.TabIndex = 46024
        Me.BtnSave.Tag = "3004"
        Me.BtnSave.Text = "ບັນທຶກ"
        Me.BtnSave.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnSave.UseVisualStyleBackColor = False
        '
        'CMB_Curr
        '
        Me.CMB_Curr.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CMB_Curr.FormattingEnabled = True
        Me.CMB_Curr.Location = New System.Drawing.Point(116, 34)
        Me.CMB_Curr.Name = "CMB_Curr"
        Me.CMB_Curr.Size = New System.Drawing.Size(110, 29)
        Me.CMB_Curr.TabIndex = 46021
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label12.ForeColor = System.Drawing.Color.Blue
        Me.Label12.Location = New System.Drawing.Point(28, 36)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(85, 24)
        Me.Label12.TabIndex = 46020
        Me.Label12.Tag = "2023"
        Me.Label12.Text = "ສະກຸນເງິນ :"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel7
        '
        Me.Panel7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel7.Controls.Add(Me.txtcurr_name)
        Me.Panel7.Controls.Add(Me.Label16)
        Me.Panel7.Controls.Add(Me.Button2)
        Me.Panel7.Controls.Add(Me.txtCurr)
        Me.Panel7.Controls.Add(Me.BtnSave2)
        Me.Panel7.Controls.Add(Me.Button5)
        Me.Panel7.Controls.Add(Me.Label13)
        Me.Panel7.Controls.Add(Me.Label6)
        Me.Panel7.Controls.Add(Me.Panel5)
        Me.Panel7.Location = New System.Drawing.Point(606, 81)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(463, 480)
        Me.Panel7.TabIndex = 46021
        '
        'txtcurr_name
        '
        Me.txtcurr_name.BackColor = System.Drawing.Color.White
        Me.txtcurr_name.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcurr_name.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcurr_name.ForeColor = System.Drawing.Color.Blue
        Me.txtcurr_name.Location = New System.Drawing.Point(348, 75)
        Me.txtcurr_name.Name = "txtcurr_name"
        Me.txtcurr_name.Size = New System.Drawing.Size(108, 30)
        Me.txtcurr_name.TabIndex = 46025
        Me.txtcurr_name.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label16.ForeColor = System.Drawing.Color.Blue
        Me.Label16.Location = New System.Drawing.Point(248, 78)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(100, 24)
        Me.Label16.TabIndex = 46027
        Me.Label16.Tag = "2023"
        Me.Label16.Text = "ເປັນພາສາລາວ :"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button2.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(250, 43)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(87, 30)
        Me.Button2.TabIndex = 46025
        Me.Button2.Tag = "3006"
        Me.Button2.Text = "ລືບ"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button2.UseVisualStyleBackColor = False
        '
        'txtCurr
        '
        Me.txtCurr.BackColor = System.Drawing.Color.White
        Me.txtCurr.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCurr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCurr.ForeColor = System.Drawing.Color.Blue
        Me.txtCurr.Location = New System.Drawing.Point(131, 75)
        Me.txtCurr.Name = "txtCurr"
        Me.txtCurr.Size = New System.Drawing.Size(113, 30)
        Me.txtCurr.TabIndex = 46024
        Me.txtCurr.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'BtnSave2
        '
        Me.BtnSave2.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.BtnSave2.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.BtnSave2.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSave2.Image = CType(resources.GetObject("BtnSave2.Image"), System.Drawing.Image)
        Me.BtnSave2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnSave2.Location = New System.Drawing.Point(131, 43)
        Me.BtnSave2.Name = "BtnSave2"
        Me.BtnSave2.Size = New System.Drawing.Size(113, 30)
        Me.BtnSave2.TabIndex = 46023
        Me.BtnSave2.Tag = "3004"
        Me.BtnSave2.Text = "ບັນທຶກ"
        Me.BtnSave2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnSave2.UseVisualStyleBackColor = False
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Button5.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button5.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.Image = CType(resources.GetObject("Button5.Image"), System.Drawing.Image)
        Me.Button5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button5.Location = New System.Drawing.Point(14, 43)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(115, 30)
        Me.Button5.TabIndex = 46022
        Me.Button5.Tag = "3003"
        Me.Button5.Text = "ເພີ່ມໃໝ່"
        Me.Button5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button5.UseVisualStyleBackColor = False
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label13.ForeColor = System.Drawing.Color.Blue
        Me.Label13.Location = New System.Drawing.Point(44, 78)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(85, 24)
        Me.Label13.TabIndex = 46021
        Me.Label13.Tag = "2023"
        Me.Label13.Text = "ສະກຸນເງິນ :"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Saysettha OT", 10.0!)
        Me.Label14.Location = New System.Drawing.Point(320, 201)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(33, 25)
        Me.Label14.TabIndex = 46024
        Me.Label14.Tag = "2013"
        Me.Label14.Text = "ຫາ:"
        '
        'dpToDate
        '
        Me.dpToDate.CustomFormat = "dd/MM/yyyy"
        Me.dpToDate.Font = New System.Drawing.Font("Saysettha OT", 10.0!)
        Me.dpToDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpToDate.Location = New System.Drawing.Point(356, 195)
        Me.dpToDate.Name = "dpToDate"
        Me.dpToDate.ShowUpDown = True
        Me.dpToDate.Size = New System.Drawing.Size(109, 31)
        Me.dpToDate.TabIndex = 46023
        '
        'dpFromDate
        '
        Me.dpFromDate.CustomFormat = "dd/MM/yyyy"
        Me.dpFromDate.Font = New System.Drawing.Font("Saysettha OT", 10.0!)
        Me.dpFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpFromDate.Location = New System.Drawing.Point(208, 194)
        Me.dpFromDate.Name = "dpFromDate"
        Me.dpFromDate.ShowUpDown = True
        Me.dpFromDate.Size = New System.Drawing.Size(110, 31)
        Me.dpFromDate.TabIndex = 46022
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Saysettha OT", 10.0!)
        Me.Label15.Location = New System.Drawing.Point(145, 199)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(62, 25)
        Me.Label15.TabIndex = 46025
        Me.Label15.Tag = "2013"
        Me.Label15.Text = "ແຕ່ວັນທີ:"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(471, 196)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(96, 30)
        Me.Button1.TabIndex = 46022
        Me.Button1.Tag = "3009"
        Me.Button1.Text = "ຄົ້ນຫາ"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.UseVisualStyleBackColor = False
        '
        'CMB_Curr_SSS
        '
        Me.CMB_Curr_SSS.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CMB_Curr_SSS.FormattingEnabled = True
        Me.CMB_Curr_SSS.Location = New System.Drawing.Point(12, 194)
        Me.CMB_Curr_SSS.Name = "CMB_Curr_SSS"
        Me.CMB_Curr_SSS.Size = New System.Drawing.Size(127, 29)
        Me.CMB_Curr_SSS.TabIndex = 46026
        '
        'Rate_setting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(1081, 560)
        Me.ControlBox = False
        Me.Controls.Add(Me.CMB_Curr_SSS)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.dpToDate)
        Me.Controls.Add(Me.dpFromDate)
        Me.Controls.Add(Me.Panel7)
        Me.Controls.Add(Me.Panel6)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Cmb_Component)
        Me.Controls.Add(Me.txt_Component_id)
        Me.Controls.Add(Me.Label25)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Blue
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "Rate_setting"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Rate setting"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        CType(Me.FG_Rate, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        CType(Me.FG_Curr, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel5.ResumeLayout(False)
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout()
        Me.Panel7.ResumeLayout(False)
        Me.Panel7.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnAddNew As System.Windows.Forms.Button
    Friend WithEvents BtnDel As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents o As System.Windows.Forms.Label
    Friend WithEvents p As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents BtnExit As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtTHB_LAK As System.Windows.Forms.TextBox
    Friend WithEvents txtUSD_LAK As System.Windows.Forms.TextBox
    Friend WithEvents txtEUR_THB As System.Windows.Forms.TextBox
    Friend WithEvents txtEUR_LAK As System.Windows.Forms.TextBox
    Friend WithEvents txtEUR_USD As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtCerrent As System.Windows.Forms.TextBox
    Friend WithEvents DTrate As System.Windows.Forms.DateTimePicker
    Friend WithEvents FG_Rate As AxVSFlex8U.AxVSFlexGrid
    Friend WithEvents txtUSD_THB As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Cmb_Component As System.Windows.Forms.ComboBox
    Friend WithEvents txt_Component_id As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents FG_Curr As AxVSFlex8U.AxVSFlexGrid
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents txtRate As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents Panel7 As System.Windows.Forms.Panel
    Friend WithEvents CMB_Curr As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents BtnSave2 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents txtCurr As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents dpToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dpFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents BtnSave As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents CMB_Curr_SSS As System.Windows.Forms.ComboBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtcurr_name As System.Windows.Forms.TextBox
    Friend WithEvents txtcurr_name2 As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents txtRate2 As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
End Class
