<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmRpt_Group
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmRpt_Group))
        Me.txtGrp = New System.Windows.Forms.TextBox
        Me.cmbGrp = New System.Windows.Forms.ComboBox
        Me.btnPrint = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.optMon = New System.Windows.Forms.RadioButton
        Me.optYear = New System.Windows.Forms.RadioButton
        Me.DTMon = New System.Windows.Forms.DateTimePicker
        Me.DTYear = New System.Windows.Forms.DateTimePicker
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Button6 = New System.Windows.Forms.Button
        Me.CheckBox5 = New System.Windows.Forms.CheckBox
        Me.Button4 = New System.Windows.Forms.Button
        Me.chkBranch = New System.Windows.Forms.CheckBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtDep = New System.Windows.Forms.TextBox
        Me.chkSum = New System.Windows.Forms.CheckBox
        Me.txtSec = New System.Windows.Forms.TextBox
        Me.txtCode = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.cmbDeprt = New System.Windows.Forms.ComboBox
        Me.cmbSec = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.txtAcc = New System.Windows.Forms.TextBox
        Me.btnShow = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dtTerm = New System.Windows.Forms.DateTimePicker
        Me.optTerm = New System.Windows.Forms.RadioButton
        Me.cmbTerm = New System.Windows.Forms.ComboBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.FG = New AxVSFlex8U.AxVSFlexGrid
        Me.CmbShow = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtCertify = New System.Windows.Forms.TextBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.Label24 = New System.Windows.Forms.Label
        Me.CmbCompany = New System.Windows.Forms.ComboBox
        Me.txtCompany = New System.Windows.Forms.TextBox
        Me.TxtLH = New System.Windows.Forms.TextBox
        Me.GHead = New System.Windows.Forms.GroupBox
        Me.Signal5 = New System.Windows.Forms.RichTextBox
        Me.Signal4 = New System.Windows.Forms.RichTextBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Place = New System.Windows.Forms.RichTextBox
        Me.Signal3 = New System.Windows.Forms.RichTextBox
        Me.Signal2 = New System.Windows.Forms.RichTextBox
        Me.Signal1 = New System.Windows.Forms.RichTextBox
        Me.Head_Nm = New System.Windows.Forms.RichTextBox
        Me.S5 = New System.Windows.Forms.Label
        Me.S4 = New System.Windows.Forms.Label
        Me.P = New System.Windows.Forms.Label
        Me.S3 = New System.Windows.Forms.Label
        Me.S2 = New System.Windows.Forms.Label
        Me.S1 = New System.Windows.Forms.Label
        Me.H = New System.Windows.Forms.Label
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.FGIT = New AxVSFlex8U.AxVSFlexGrid
        Me.CheckBox3 = New System.Windows.Forms.CheckBox
        Me.CheckBox4 = New System.Windows.Forms.CheckBox
        Me.Button5 = New System.Windows.Forms.Button
        Me.Exchange = New System.Windows.Forms.TextBox
        Me.cmbCurr = New System.Windows.Forms.ComboBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtDrNm = New System.Windows.Forms.TextBox
        Me.TxtCrNm = New System.Windows.Forms.TextBox
        Me.BtnSearch = New System.Windows.Forms.Button
        Me.Button7 = New System.Windows.Forms.Button
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.FG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GHead.SuspendLayout()
        CType(Me.FGIT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtGrp
        '
        Me.txtGrp.Enabled = False
        Me.txtGrp.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGrp.Location = New System.Drawing.Point(189, 95)
        Me.txtGrp.Name = "txtGrp"
        Me.txtGrp.Size = New System.Drawing.Size(112, 35)
        Me.txtGrp.TabIndex = 98
        '
        'cmbGrp
        '
        Me.cmbGrp.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbGrp.FormattingEnabled = True
        Me.cmbGrp.Items.AddRange(New Object() {"LAK", "THB", "USD"})
        Me.cmbGrp.Location = New System.Drawing.Point(304, 98)
        Me.cmbGrp.Name = "cmbGrp"
        Me.cmbGrp.Size = New System.Drawing.Size(609, 32)
        Me.cmbGrp.TabIndex = 97
        '
        'btnPrint
        '
        Me.btnPrint.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnPrint.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnPrint.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.ForeColor = System.Drawing.Color.Black
        Me.btnPrint.Image = CType(resources.GetObject("btnPrint.Image"), System.Drawing.Image)
        Me.btnPrint.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPrint.Location = New System.Drawing.Point(168, 12)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(83, 35)
        Me.btnPrint.TabIndex = 100
        Me.btnPrint.Text = "A4"
        Me.btnPrint.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnPrint.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.Location = New System.Drawing.Point(7, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(57, 35)
        Me.Button1.TabIndex = 99
        Me.Button1.UseVisualStyleBackColor = False
        '
        'optMon
        '
        Me.optMon.AutoSize = True
        Me.optMon.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optMon.ForeColor = System.Drawing.Color.Black
        Me.optMon.Location = New System.Drawing.Point(82, 31)
        Me.optMon.Name = "optMon"
        Me.optMon.Size = New System.Drawing.Size(101, 28)
        Me.optMon.TabIndex = 101
        Me.optMon.Text = "ປະຈຳເດືອນ"
        Me.optMon.UseVisualStyleBackColor = True
        '
        'optYear
        '
        Me.optYear.AutoSize = True
        Me.optYear.Checked = True
        Me.optYear.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optYear.ForeColor = System.Drawing.Color.Black
        Me.optYear.Location = New System.Drawing.Point(755, 31)
        Me.optYear.Name = "optYear"
        Me.optYear.Size = New System.Drawing.Size(75, 28)
        Me.optYear.TabIndex = 102
        Me.optYear.TabStop = True
        Me.optYear.Text = "ປະຈຳປີ"
        Me.optYear.UseVisualStyleBackColor = True
        '
        'DTMon
        '
        Me.DTMon.CustomFormat = "MM/yyyy"
        Me.DTMon.Enabled = False
        Me.DTMon.Font = New System.Drawing.Font("Saysettha OT", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DTMon.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTMon.Location = New System.Drawing.Point(189, 26)
        Me.DTMon.Name = "DTMon"
        Me.DTMon.ShowUpDown = True
        Me.DTMon.Size = New System.Drawing.Size(112, 41)
        Me.DTMon.TabIndex = 103
        '
        'DTYear
        '
        Me.DTYear.CustomFormat = "yyyy"
        Me.DTYear.Font = New System.Drawing.Font("Saysettha OT", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DTYear.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTYear.Location = New System.Drawing.Point(836, 25)
        Me.DTYear.Name = "DTYear"
        Me.DTYear.ShowUpDown = True
        Me.DTYear.Size = New System.Drawing.Size(77, 41)
        Me.DTYear.TabIndex = 105
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Button6)
        Me.GroupBox2.Controls.Add(Me.CheckBox5)
        Me.GroupBox2.Controls.Add(Me.Button4)
        Me.GroupBox2.Controls.Add(Me.chkBranch)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.txtDep)
        Me.GroupBox2.Controls.Add(Me.chkSum)
        Me.GroupBox2.Controls.Add(Me.txtSec)
        Me.GroupBox2.Controls.Add(Me.txtCode)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.cmbDeprt)
        Me.GroupBox2.Controls.Add(Me.cmbSec)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.cmbGrp)
        Me.GroupBox2.Controls.Add(Me.txtGrp)
        Me.GroupBox2.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.GroupBox2.Location = New System.Drawing.Point(12, 141)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(929, 183)
        Me.GroupBox2.TabIndex = 109
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "ຂອບເຂດການລາຍງານ"
        '
        'Button6
        '
        Me.Button6.Font = New System.Drawing.Font("Saysettha OT", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button6.Location = New System.Drawing.Point(757, 135)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(156, 36)
        Me.Button6.TabIndex = 144
        Me.Button6.Text = "ໂອນເຂົ້າບັນຊີ"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'CheckBox5
        '
        Me.CheckBox5.AutoSize = True
        Me.CheckBox5.Location = New System.Drawing.Point(6, 68)
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.Size = New System.Drawing.Size(111, 28)
        Me.CheckBox5.TabIndex = 109
        Me.CheckBox5.Text = "CheckBox5"
        Me.CheckBox5.UseVisualStyleBackColor = True
        Me.CheckBox5.Visible = False
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Button4.Enabled = False
        Me.Button4.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.Black
        Me.Button4.Image = CType(resources.GetObject("Button4.Image"), System.Drawing.Image)
        Me.Button4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button4.Location = New System.Drawing.Point(649, 135)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(108, 36)
        Me.Button4.TabIndex = 143
        Me.Button4.Text = "ສົມທຽບ"
        Me.Button4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button4.UseVisualStyleBackColor = False
        '
        'chkBranch
        '
        Me.chkBranch.AutoSize = True
        Me.chkBranch.Location = New System.Drawing.Point(16, 136)
        Me.chkBranch.Name = "chkBranch"
        Me.chkBranch.Size = New System.Drawing.Size(105, 28)
        Me.chkBranch.TabIndex = 134
        Me.chkBranch.Text = "ບໍ່ແຍກສາຂາ"
        Me.chkBranch.UseVisualStyleBackColor = True
        Me.chkBranch.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(324, 138)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(143, 24)
        Me.Label4.TabIndex = 133
        Me.Label4.Text = "ສະເພາະລະຫັດຊັບສິນ"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(6, 101)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(177, 24)
        Me.Label2.TabIndex = 132
        Me.Label2.Text = "ໝວດຊັບສິນ"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtDep
        '
        Me.txtDep.Enabled = False
        Me.txtDep.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDep.Location = New System.Drawing.Point(189, 57)
        Me.txtDep.Name = "txtDep"
        Me.txtDep.Size = New System.Drawing.Size(112, 35)
        Me.txtDep.TabIndex = 131
        '
        'chkSum
        '
        Me.chkSum.AutoSize = True
        Me.chkSum.ForeColor = System.Drawing.Color.Blue
        Me.chkSum.Location = New System.Drawing.Point(161, 137)
        Me.chkSum.Name = "chkSum"
        Me.chkSum.Size = New System.Drawing.Size(140, 28)
        Me.chkSum.TabIndex = 112
        Me.chkSum.Text = "ສັງລວມຕາມໝວດ"
        Me.chkSum.UseVisualStyleBackColor = True
        '
        'txtSec
        '
        Me.txtSec.Enabled = False
        Me.txtSec.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSec.Location = New System.Drawing.Point(189, 19)
        Me.txtSec.Name = "txtSec"
        Me.txtSec.Size = New System.Drawing.Size(112, 35)
        Me.txtSec.TabIndex = 130
        '
        'txtCode
        '
        Me.txtCode.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCode.Location = New System.Drawing.Point(473, 135)
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(174, 35)
        Me.txtCode.TabIndex = 111
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(17, 23)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(166, 24)
        Me.Label3.TabIndex = 129
        Me.Label3.Text = "ສຳນັກງານ"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbDeprt
        '
        Me.cmbDeprt.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbDeprt.FormattingEnabled = True
        Me.cmbDeprt.Items.AddRange(New Object() {"ລະຫັດບັນຊີ / Account Code", "ຊື່ບັນຊີ (ລາວ) / Ac Name (L)", "ຊື່ບັນຊີ (ອັງກິດ) / Ac Name (E)"})
        Me.cmbDeprt.Location = New System.Drawing.Point(304, 58)
        Me.cmbDeprt.Name = "cmbDeprt"
        Me.cmbDeprt.Size = New System.Drawing.Size(609, 32)
        Me.cmbDeprt.TabIndex = 128
        '
        'cmbSec
        '
        Me.cmbSec.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbSec.FormattingEnabled = True
        Me.cmbSec.Items.AddRange(New Object() {"ລະຫັດບັນຊີ / Account Code", "ຊື່ບັນຊີ (ລາວ) / Ac Name (L)", "ຊື່ບັນຊີ (ອັງກິດ) / Ac Name (E)"})
        Me.cmbSec.Location = New System.Drawing.Point(304, 20)
        Me.cmbSec.Name = "cmbSec"
        Me.cmbSec.Size = New System.Drawing.Size(609, 32)
        Me.cmbSec.TabIndex = 126
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(10, 60)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(173, 24)
        Me.Label6.TabIndex = 127
        Me.Label6.Text = "ພະແນກ"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("Saysettha OT", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Location = New System.Drawing.Point(903, 351)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(156, 36)
        Me.Button3.TabIndex = 136
        Me.Button3.Text = "ໂອນເຂົ້າບັນຊີ"
        Me.Button3.UseVisualStyleBackColor = True
        Me.Button3.Visible = False
        '
        'txtAcc
        '
        Me.txtAcc.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAcc.Location = New System.Drawing.Point(1029, 179)
        Me.txtAcc.Name = "txtAcc"
        Me.txtAcc.Size = New System.Drawing.Size(117, 30)
        Me.txtAcc.TabIndex = 108
        '
        'btnShow
        '
        Me.btnShow.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnShow.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnShow.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnShow.Image = CType(resources.GetObject("btnShow.Image"), System.Drawing.Image)
        Me.btnShow.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnShow.Location = New System.Drawing.Point(68, 12)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(98, 35)
        Me.btnShow.TabIndex = 116
        Me.btnShow.Tag = "3015"
        Me.btnShow.Text = "Show"
        Me.btnShow.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnShow.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Saysettha OT", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(478, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(208, 34)
        Me.Label1.TabIndex = 117
        Me.Label1.Text = "ລາຍງານແບບສັງລວມ"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dtTerm)
        Me.GroupBox1.Controls.Add(Me.optTerm)
        Me.GroupBox1.Controls.Add(Me.cmbTerm)
        Me.GroupBox1.Controls.Add(Me.DTMon)
        Me.GroupBox1.Controls.Add(Me.optMon)
        Me.GroupBox1.Controls.Add(Me.optYear)
        Me.GroupBox1.Controls.Add(Me.DTYear)
        Me.GroupBox1.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.Black
        Me.GroupBox1.Location = New System.Drawing.Point(12, 66)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(929, 73)
        Me.GroupBox1.TabIndex = 108
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "ຊ່ວງເວລາລາຍງານ"
        '
        'dtTerm
        '
        Me.dtTerm.CustomFormat = "yyyy"
        Me.dtTerm.Enabled = False
        Me.dtTerm.Font = New System.Drawing.Font("Saysettha OT", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtTerm.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtTerm.Location = New System.Drawing.Point(543, 25)
        Me.dtTerm.Name = "dtTerm"
        Me.dtTerm.ShowUpDown = True
        Me.dtTerm.Size = New System.Drawing.Size(77, 41)
        Me.dtTerm.TabIndex = 108
        '
        'optTerm
        '
        Me.optTerm.AutoSize = True
        Me.optTerm.ForeColor = System.Drawing.Color.Black
        Me.optTerm.Location = New System.Drawing.Point(358, 31)
        Me.optTerm.Name = "optTerm"
        Me.optTerm.Size = New System.Drawing.Size(93, 28)
        Me.optTerm.TabIndex = 107
        Me.optTerm.TabStop = True
        Me.optTerm.Text = "ປະຈຳງວດ"
        Me.optTerm.UseVisualStyleBackColor = True
        '
        'cmbTerm
        '
        Me.cmbTerm.Enabled = False
        Me.cmbTerm.Font = New System.Drawing.Font("Saysettha OT", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbTerm.FormattingEnabled = True
        Me.cmbTerm.Items.AddRange(New Object() {"  I", "  II", "  III", "  IV"})
        Me.cmbTerm.Location = New System.Drawing.Point(457, 25)
        Me.cmbTerm.Name = "cmbTerm"
        Me.cmbTerm.Size = New System.Drawing.Size(80, 39)
        Me.cmbTerm.TabIndex = 106
        Me.cmbTerm.Tag = ""
        Me.cmbTerm.Text = "  I"
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button2.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(253, 12)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(101, 35)
        Me.Button2.TabIndex = 118
        Me.Button2.Text = "Legal"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button2.UseVisualStyleBackColor = False
        '
        'FG
        '
        Me.FG.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FG.DataSource = Nothing
        Me.FG.Location = New System.Drawing.Point(12, 330)
        Me.FG.Name = "FG"
        Me.FG.OcxState = CType(resources.GetObject("FG.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FG.Size = New System.Drawing.Size(1259, 217)
        Me.FG.TabIndex = 111
        '
        'CmbShow
        '
        Me.CmbShow.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbShow.FormattingEnabled = True
        Me.CmbShow.Items.AddRange(New Object() {"ສະແດງທັງໝົດ", "ສະເພາະລາຍການພວມນຳໃຊ້", "ສະເພາະລາຍການສະສາງແລ້ວ"})
        Me.CmbShow.Location = New System.Drawing.Point(737, 45)
        Me.CmbShow.Name = "CmbShow"
        Me.CmbShow.Size = New System.Drawing.Size(188, 32)
        Me.CmbShow.TabIndex = 135
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(606, 50)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(125, 24)
        Me.Label5.TabIndex = 133
        Me.Label5.Text = "ສະແດງລາຍການ"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TxtCertify
        '
        Me.TxtCertify.Enabled = False
        Me.TxtCertify.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtCertify.Location = New System.Drawing.Point(758, 7)
        Me.TxtCertify.Name = "TxtCertify"
        Me.TxtCertify.Size = New System.Drawing.Size(167, 35)
        Me.TxtCertify.TabIndex = 137
        '
        'TextBox2
        '
        Me.TextBox2.Enabled = False
        Me.TextBox2.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox2.Location = New System.Drawing.Point(1203, 190)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(113, 35)
        Me.TextBox2.TabIndex = 138
        Me.TextBox2.Visible = False
        '
        'Label26
        '
        Me.Label26.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.ForeColor = System.Drawing.Color.Black
        Me.Label26.Location = New System.Drawing.Point(1122, 88)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(89, 24)
        Me.Label26.TabIndex = 142
        Me.Label26.Text = "Company"
        Me.Label26.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label26.Visible = False
        '
        'Label24
        '
        Me.Label24.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.ForeColor = System.Drawing.Color.Black
        Me.Label24.Location = New System.Drawing.Point(687, 13)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(68, 24)
        Me.Label24.TabIndex = 141
        Me.Label24.Text = "ເລກທີ"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CmbCompany
        '
        Me.CmbCompany.Font = New System.Drawing.Font("Saysettha Lao", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbCompany.ForeColor = System.Drawing.Color.Black
        Me.CmbCompany.FormattingEnabled = True
        Me.CmbCompany.Items.AddRange(New Object() {"LAK", "THB", "USD"})
        Me.CmbCompany.Location = New System.Drawing.Point(1217, 84)
        Me.CmbCompany.Name = "CmbCompany"
        Me.CmbCompany.Size = New System.Drawing.Size(63, 37)
        Me.CmbCompany.TabIndex = 140
        Me.CmbCompany.Visible = False
        '
        'txtCompany
        '
        Me.txtCompany.Enabled = False
        Me.txtCompany.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCompany.Location = New System.Drawing.Point(1203, 139)
        Me.txtCompany.Name = "txtCompany"
        Me.txtCompany.Size = New System.Drawing.Size(117, 35)
        Me.txtCompany.TabIndex = 139
        Me.txtCompany.Visible = False
        '
        'TxtLH
        '
        Me.TxtLH.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtLH.Location = New System.Drawing.Point(1029, 243)
        Me.TxtLH.Name = "TxtLH"
        Me.TxtLH.Size = New System.Drawing.Size(117, 30)
        Me.TxtLH.TabIndex = 137
        '
        'GHead
        '
        Me.GHead.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GHead.Controls.Add(Me.Signal5)
        Me.GHead.Controls.Add(Me.Signal4)
        Me.GHead.Controls.Add(Me.TextBox1)
        Me.GHead.Controls.Add(Me.Place)
        Me.GHead.Controls.Add(Me.Signal3)
        Me.GHead.Controls.Add(Me.Signal2)
        Me.GHead.Controls.Add(Me.Signal1)
        Me.GHead.Controls.Add(Me.Head_Nm)
        Me.GHead.Controls.Add(Me.S5)
        Me.GHead.Controls.Add(Me.S4)
        Me.GHead.Controls.Add(Me.P)
        Me.GHead.Controls.Add(Me.S3)
        Me.GHead.Controls.Add(Me.S2)
        Me.GHead.Controls.Add(Me.S1)
        Me.GHead.Controls.Add(Me.H)
        Me.GHead.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GHead.ForeColor = System.Drawing.Color.Black
        Me.GHead.Location = New System.Drawing.Point(1233, 45)
        Me.GHead.Name = "GHead"
        Me.GHead.Size = New System.Drawing.Size(68, 279)
        Me.GHead.TabIndex = 45765
        Me.GHead.TabStop = False
        Me.GHead.Text = "Header and Footer"
        Me.GHead.Visible = False
        '
        'Signal5
        '
        Me.Signal5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Signal5.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Signal5.ForeColor = System.Drawing.Color.Black
        Me.Signal5.Location = New System.Drawing.Point(107, 197)
        Me.Signal5.Multiline = False
        Me.Signal5.Name = "Signal5"
        Me.Signal5.Size = New System.Drawing.Size(0, 31)
        Me.Signal5.TabIndex = 45766
        Me.Signal5.Text = ""
        '
        'Signal4
        '
        Me.Signal4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Signal4.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Signal4.ForeColor = System.Drawing.Color.Black
        Me.Signal4.Location = New System.Drawing.Point(107, 162)
        Me.Signal4.Multiline = False
        Me.Signal4.Name = "Signal4"
        Me.Signal4.Size = New System.Drawing.Size(0, 31)
        Me.Signal4.TabIndex = 45765
        Me.Signal4.Text = ""
        '
        'TextBox1
        '
        Me.TextBox1.Enabled = False
        Me.TextBox1.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(107, 270)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(34, 30)
        Me.TextBox1.TabIndex = 45763
        Me.TextBox1.Text = "001"
        '
        'Place
        '
        Me.Place.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Place.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Place.ForeColor = System.Drawing.Color.Black
        Me.Place.Location = New System.Drawing.Point(107, 232)
        Me.Place.Multiline = False
        Me.Place.Name = "Place"
        Me.Place.Size = New System.Drawing.Size(0, 31)
        Me.Place.TabIndex = 161
        Me.Place.Text = ""
        '
        'Signal3
        '
        Me.Signal3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Signal3.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Signal3.ForeColor = System.Drawing.Color.Black
        Me.Signal3.Location = New System.Drawing.Point(107, 127)
        Me.Signal3.Multiline = False
        Me.Signal3.Name = "Signal3"
        Me.Signal3.Size = New System.Drawing.Size(0, 31)
        Me.Signal3.TabIndex = 160
        Me.Signal3.Text = ""
        '
        'Signal2
        '
        Me.Signal2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Signal2.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Signal2.ForeColor = System.Drawing.Color.Black
        Me.Signal2.Location = New System.Drawing.Point(107, 92)
        Me.Signal2.Multiline = False
        Me.Signal2.Name = "Signal2"
        Me.Signal2.Size = New System.Drawing.Size(0, 31)
        Me.Signal2.TabIndex = 159
        Me.Signal2.Text = ""
        '
        'Signal1
        '
        Me.Signal1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Signal1.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Signal1.ForeColor = System.Drawing.Color.Black
        Me.Signal1.Location = New System.Drawing.Point(107, 57)
        Me.Signal1.Multiline = False
        Me.Signal1.Name = "Signal1"
        Me.Signal1.Size = New System.Drawing.Size(0, 31)
        Me.Signal1.TabIndex = 158
        Me.Signal1.Text = ""
        '
        'Head_Nm
        '
        Me.Head_Nm.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Head_Nm.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Head_Nm.ForeColor = System.Drawing.Color.Black
        Me.Head_Nm.Location = New System.Drawing.Point(107, 22)
        Me.Head_Nm.Multiline = False
        Me.Head_Nm.Name = "Head_Nm"
        Me.Head_Nm.Size = New System.Drawing.Size(0, 31)
        Me.Head_Nm.TabIndex = 157
        Me.Head_Nm.Text = ""
        '
        'S5
        '
        Me.S5.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.S5.Location = New System.Drawing.Point(9, 201)
        Me.S5.Name = "S5"
        Me.S5.Size = New System.Drawing.Size(98, 26)
        Me.S5.TabIndex = 45767
        Me.S5.Tag = "2106"
        Me.S5.Text = "Signal 5:"
        Me.S5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'S4
        '
        Me.S4.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.S4.Location = New System.Drawing.Point(9, 167)
        Me.S4.Name = "S4"
        Me.S4.Size = New System.Drawing.Size(98, 26)
        Me.S4.TabIndex = 45768
        Me.S4.Tag = "2106"
        Me.S4.Text = "Signal 4:"
        Me.S4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'P
        '
        Me.P.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.P.Location = New System.Drawing.Point(5, 237)
        Me.P.Name = "P"
        Me.P.Size = New System.Drawing.Size(103, 26)
        Me.P.TabIndex = 45764
        Me.P.Tag = "2106"
        Me.P.Text = "Place Name:"
        Me.P.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'S3
        '
        Me.S3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.S3.Location = New System.Drawing.Point(8, 131)
        Me.S3.Name = "S3"
        Me.S3.Size = New System.Drawing.Size(99, 26)
        Me.S3.TabIndex = 45764
        Me.S3.Tag = "2106"
        Me.S3.Text = "Signal 3:"
        Me.S3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'S2
        '
        Me.S2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.S2.Location = New System.Drawing.Point(9, 97)
        Me.S2.Name = "S2"
        Me.S2.Size = New System.Drawing.Size(98, 26)
        Me.S2.TabIndex = 45764
        Me.S2.Tag = "2106"
        Me.S2.Text = "Signal 2:"
        Me.S2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'S1
        '
        Me.S1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.S1.Location = New System.Drawing.Point(12, 59)
        Me.S1.Name = "S1"
        Me.S1.Size = New System.Drawing.Size(95, 26)
        Me.S1.TabIndex = 45764
        Me.S1.Tag = "2106"
        Me.S1.Text = "Signal 1:"
        Me.S1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'H
        '
        Me.H.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.H.Location = New System.Drawing.Point(3, 25)
        Me.H.Name = "H"
        Me.H.Size = New System.Drawing.Size(104, 26)
        Me.H.TabIndex = 45763
        Me.H.Tag = "2106"
        Me.H.Text = "Heading Name:"
        Me.H.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(370, 5)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(101, 28)
        Me.CheckBox1.TabIndex = 144
        Me.CheckBox1.Text = "ລາຍລະອຽດ"
        Me.CheckBox1.UseVisualStyleBackColor = True
        Me.CheckBox1.Visible = False
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox2.Location = New System.Drawing.Point(370, 32)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(239, 28)
        Me.CheckBox2.TabIndex = 45779
        Me.CheckBox2.Text = "ສະຫລຸບ ລວມຍອດມູນຄ່າຊັບສົມບັດ"
        Me.CheckBox2.UseVisualStyleBackColor = True
        Me.CheckBox2.Visible = False
        '
        'FGIT
        '
        Me.FGIT.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FGIT.DataSource = Nothing
        Me.FGIT.Location = New System.Drawing.Point(1126, 53)
        Me.FGIT.Name = "FGIT"
        Me.FGIT.OcxState = CType(resources.GetObject("FGIT.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FGIT.Size = New System.Drawing.Size(145, 250)
        Me.FGIT.TabIndex = 45780
        Me.FGIT.Visible = False
        '
        'CheckBox3
        '
        Me.CheckBox3.Appearance = System.Windows.Forms.Appearance.Button
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Location = New System.Drawing.Point(1068, 13)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(110, 23)
        Me.CheckBox3.TabIndex = 45781
        Me.CheckBox3.Text = "ລາຍການສົ່ງເອກະສານ"
        Me.CheckBox3.UseVisualStyleBackColor = True
        Me.CheckBox3.Visible = False
        '
        'CheckBox4
        '
        Me.CheckBox4.AutoSize = True
        Me.CheckBox4.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox4.Location = New System.Drawing.Point(947, 18)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(101, 28)
        Me.CheckBox4.TabIndex = 45782
        Me.CheckBox4.Text = "ລາຍລະອຽດ"
        Me.CheckBox4.UseVisualStyleBackColor = True
        Me.CheckBox4.Visible = False
        '
        'Button5
        '
        Me.Button5.Font = New System.Drawing.Font("Saysettha OT", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.Location = New System.Drawing.Point(947, 309)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(156, 36)
        Me.Button5.TabIndex = 45783
        Me.Button5.Text = "ໂອນເຂົ້າບັນຊີ DATA Asset"
        Me.Button5.UseVisualStyleBackColor = True
        Me.Button5.Visible = False
        '
        'Exchange
        '
        Me.Exchange.Enabled = False
        Me.Exchange.Font = New System.Drawing.Font("Saysettha OT", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Exchange.Location = New System.Drawing.Point(813, 318)
        Me.Exchange.Name = "Exchange"
        Me.Exchange.Size = New System.Drawing.Size(77, 28)
        Me.Exchange.TabIndex = 169
        Me.Exchange.Text = "1"
        Me.Exchange.Visible = False
        '
        'cmbCurr
        '
        Me.cmbCurr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbCurr.FormattingEnabled = True
        Me.cmbCurr.Items.AddRange(New Object() {"LAK", "THB", "USD"})
        Me.cmbCurr.Location = New System.Drawing.Point(698, 320)
        Me.cmbCurr.Name = "cmbCurr"
        Me.cmbCurr.Size = New System.Drawing.Size(71, 29)
        Me.cmbCurr.TabIndex = 170
        Me.cmbCurr.Text = "LAK"
        Me.cmbCurr.Visible = False
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(947, 183)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(76, 24)
        Me.Label7.TabIndex = 45784
        Me.Label7.Text = "Debit:"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(947, 248)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(76, 24)
        Me.Label8.TabIndex = 45785
        Me.Label8.Text = "Credit:"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TxtDrNm
        '
        Me.TxtDrNm.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtDrNm.Location = New System.Drawing.Point(951, 211)
        Me.TxtDrNm.Name = "TxtDrNm"
        Me.TxtDrNm.Size = New System.Drawing.Size(320, 30)
        Me.TxtDrNm.TabIndex = 45786
        '
        'TxtCrNm
        '
        Me.TxtCrNm.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtCrNm.Location = New System.Drawing.Point(951, 276)
        Me.TxtCrNm.Name = "TxtCrNm"
        Me.TxtCrNm.Size = New System.Drawing.Size(320, 30)
        Me.TxtCrNm.TabIndex = 45787
        '
        'BtnSearch
        '
        Me.BtnSearch.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSearch.Location = New System.Drawing.Point(1152, 179)
        Me.BtnSearch.Name = "BtnSearch"
        Me.BtnSearch.Size = New System.Drawing.Size(34, 30)
        Me.BtnSearch.TabIndex = 46036
        Me.BtnSearch.Tag = "3012"
        Me.BtnSearch.Text = "....."
        Me.BtnSearch.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnSearch.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button7.Location = New System.Drawing.Point(1152, 244)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(34, 30)
        Me.Button7.TabIndex = 46037
        Me.Button7.Tag = "3012"
        Me.Button7.Text = "....."
        Me.Button7.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button7.UseVisualStyleBackColor = True
        '
        'FrmRpt_Group
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.AutoSize = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(187, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1276, 550)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.BtnSearch)
        Me.Controls.Add(Me.TxtCrNm)
        Me.Controls.Add(Me.TxtDrNm)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.cmbCurr)
        Me.Controls.Add(Me.Exchange)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.CheckBox4)
        Me.Controls.Add(Me.CheckBox3)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.GHead)
        Me.Controls.Add(Me.TxtLH)
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.Label24)
        Me.Controls.Add(Me.CmbCompany)
        Me.Controls.Add(Me.txtCompany)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.TxtCertify)
        Me.Controls.Add(Me.CmbShow)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtAcc)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnShow)
        Me.Controls.Add(Me.FG)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.FGIT)
        Me.ForeColor = System.Drawing.Color.Black
        Me.Name = "FrmRpt_Group"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Report by Group"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GHead.ResumeLayout(False)
        Me.GHead.PerformLayout()
        CType(Me.FGIT, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtGrp As System.Windows.Forms.TextBox
    Friend WithEvents cmbGrp As System.Windows.Forms.ComboBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents optMon As System.Windows.Forms.RadioButton
    Friend WithEvents optYear As System.Windows.Forms.RadioButton
    Friend WithEvents DTMon As System.Windows.Forms.DateTimePicker
    Friend WithEvents DTYear As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtAcc As System.Windows.Forms.TextBox
    Friend WithEvents FG As AxVSFlex8U.AxVSFlexGrid
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtCode As System.Windows.Forms.TextBox
    Friend WithEvents chkSum As System.Windows.Forms.CheckBox
    Friend WithEvents txtDep As System.Windows.Forms.TextBox
    Friend WithEvents txtSec As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmbDeprt As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSec As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents chkBranch As System.Windows.Forms.CheckBox
    Friend WithEvents optTerm As System.Windows.Forms.RadioButton
    Friend WithEvents cmbTerm As System.Windows.Forms.ComboBox
    Friend WithEvents dtTerm As System.Windows.Forms.DateTimePicker
    Friend WithEvents CmbShow As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents TxtCertify As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents CmbCompany As System.Windows.Forms.ComboBox
    Friend WithEvents txtCompany As System.Windows.Forms.TextBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents TxtLH As System.Windows.Forms.TextBox
    Friend WithEvents GHead As System.Windows.Forms.GroupBox
    Friend WithEvents Signal5 As System.Windows.Forms.RichTextBox
    Friend WithEvents Signal4 As System.Windows.Forms.RichTextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Place As System.Windows.Forms.RichTextBox
    Friend WithEvents Signal3 As System.Windows.Forms.RichTextBox
    Friend WithEvents Signal2 As System.Windows.Forms.RichTextBox
    Friend WithEvents Signal1 As System.Windows.Forms.RichTextBox
    Friend WithEvents Head_Nm As System.Windows.Forms.RichTextBox
    Friend WithEvents S5 As System.Windows.Forms.Label
    Friend WithEvents S4 As System.Windows.Forms.Label
    Friend WithEvents P As System.Windows.Forms.Label
    Friend WithEvents S3 As System.Windows.Forms.Label
    Friend WithEvents S2 As System.Windows.Forms.Label
    Friend WithEvents S1 As System.Windows.Forms.Label
    Friend WithEvents H As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents FGIT As AxVSFlex8U.AxVSFlexGrid
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox4 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox5 As System.Windows.Forms.CheckBox
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Exchange As System.Windows.Forms.TextBox
    Friend WithEvents cmbCurr As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtDrNm As System.Windows.Forms.TextBox
    Friend WithEvents TxtCrNm As System.Windows.Forms.TextBox
    Friend WithEvents BtnSearch As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
End Class
