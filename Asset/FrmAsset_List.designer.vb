<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmAsset_List
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmAsset_List))
        Me.CndClose = New System.Windows.Forms.Button
        Me.cmbSort = New System.Windows.Forms.ComboBox
        Me.lSort = New System.Windows.Forms.Label
        Me.cmbGrp = New System.Windows.Forms.ComboBox
        Me.txtGrp = New System.Windows.Forms.TextBox
        Me.txtCode = New System.Windows.Forms.TextBox
        Me.txtNm = New System.Windows.Forms.TextBox
        Me.DTUSE = New System.Windows.Forms.DateTimePicker
        Me.chkDT = New System.Windows.Forms.CheckBox
        Me.cmbSec = New System.Windows.Forms.ComboBox
        Me.txtNo = New System.Windows.Forms.TextBox
        Me.txtSec = New System.Windows.Forms.TextBox
        Me.cmbDeprt = New System.Windows.Forms.ComboBox
        Me.txtDep = New System.Windows.Forms.TextBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.CmbShow = New System.Windows.Forms.ComboBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.DG = New System.Windows.Forms.DataGridView
        Me.Label8 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnDel = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnAdd = New System.Windows.Forms.Button
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.Panel1 = New System.Windows.Forms.Panel
        CType(Me.DG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'CndClose
        '
        Me.CndClose.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.CndClose.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.CndClose.Image = CType(resources.GetObject("CndClose.Image"), System.Drawing.Image)
        Me.CndClose.Location = New System.Drawing.Point(8, 19)
        Me.CndClose.Name = "CndClose"
        Me.CndClose.Size = New System.Drawing.Size(60, 36)
        Me.CndClose.TabIndex = 24
        Me.CndClose.UseVisualStyleBackColor = False
        '
        'cmbSort
        '
        Me.cmbSort.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbSort.FormattingEnabled = True
        Me.cmbSort.Items.AddRange(New Object() {"ລະຫັດ / Code", "ຊື່ / Name", "ວັນທີ / Date"})
        Me.cmbSort.Location = New System.Drawing.Point(1205, 20)
        Me.cmbSort.Name = "cmbSort"
        Me.cmbSort.Size = New System.Drawing.Size(136, 29)
        Me.cmbSort.TabIndex = 48
        '
        'lSort
        '
        Me.lSort.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lSort.ForeColor = System.Drawing.Color.Black
        Me.lSort.Location = New System.Drawing.Point(1091, 20)
        Me.lSort.Name = "lSort"
        Me.lSort.Size = New System.Drawing.Size(112, 30)
        Me.lSort.TabIndex = 50
        Me.lSort.Tag = "2007"
        Me.lSort.Text = "ລຽງຕາມ"
        Me.lSort.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbGrp
        '
        Me.cmbGrp.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbGrp.FormattingEnabled = True
        Me.cmbGrp.Items.AddRange(New Object() {"ລະຫັດບັນຊີ / Account Code", "ຊື່ບັນຊີ (ລາວ) / Ac Name (L)", "ຊື່ບັນຊີ (ອັງກິດ) / Ac Name (E)"})
        Me.cmbGrp.Location = New System.Drawing.Point(96, 56)
        Me.cmbGrp.Name = "cmbGrp"
        Me.cmbGrp.Size = New System.Drawing.Size(178, 29)
        Me.cmbGrp.TabIndex = 51
        '
        'txtGrp
        '
        Me.txtGrp.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGrp.Location = New System.Drawing.Point(487, 37)
        Me.txtGrp.Name = "txtGrp"
        Me.txtGrp.Size = New System.Drawing.Size(57, 22)
        Me.txtGrp.TabIndex = 53
        Me.txtGrp.Visible = False
        '
        'txtCode
        '
        Me.txtCode.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCode.Location = New System.Drawing.Point(885, 59)
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(78, 30)
        Me.txtCode.TabIndex = 55
        '
        'txtNm
        '
        Me.txtNm.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNm.Location = New System.Drawing.Point(1206, 57)
        Me.txtNm.Name = "txtNm"
        Me.txtNm.Size = New System.Drawing.Size(137, 30)
        Me.txtNm.TabIndex = 57
        '
        'DTUSE
        '
        Me.DTUSE.CustomFormat = "dd/MM/yyyy"
        Me.DTUSE.Enabled = False
        Me.DTUSE.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DTUSE.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTUSE.Location = New System.Drawing.Point(715, 62)
        Me.DTUSE.Name = "DTUSE"
        Me.DTUSE.ShowUpDown = True
        Me.DTUSE.Size = New System.Drawing.Size(102, 30)
        Me.DTUSE.TabIndex = 60
        '
        'chkDT
        '
        Me.chkDT.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDT.ForeColor = System.Drawing.Color.Black
        Me.chkDT.Location = New System.Drawing.Point(611, 63)
        Me.chkDT.Name = "chkDT"
        Me.chkDT.Size = New System.Drawing.Size(100, 28)
        Me.chkDT.TabIndex = 62
        Me.chkDT.Tag = "5001"
        Me.chkDT.Text = "ປ.ດ ນຳໃຊ້"
        Me.chkDT.UseVisualStyleBackColor = True
        '
        'cmbSec
        '
        Me.cmbSec.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbSec.FormattingEnabled = True
        Me.cmbSec.Items.AddRange(New Object() {"ລະຫັດບັນຊີ / Account Code", "ຊື່ບັນຊີ (ລາວ) / Ac Name (L)", "ຊື່ບັນຊີ (ອັງກິດ) / Ac Name (E)"})
        Me.cmbSec.Location = New System.Drawing.Point(58, 16)
        Me.cmbSec.Name = "cmbSec"
        Me.cmbSec.Size = New System.Drawing.Size(189, 29)
        Me.cmbSec.TabIndex = 63
        '
        'txtNo
        '
        Me.txtNo.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNo.Location = New System.Drawing.Point(1023, 59)
        Me.txtNo.Name = "txtNo"
        Me.txtNo.Size = New System.Drawing.Size(101, 30)
        Me.txtNo.TabIndex = 65
        '
        'txtSec
        '
        Me.txtSec.Location = New System.Drawing.Point(25, 44)
        Me.txtSec.Name = "txtSec"
        Me.txtSec.Size = New System.Drawing.Size(41, 34)
        Me.txtSec.TabIndex = 67
        Me.txtSec.Visible = False
        '
        'cmbDeprt
        '
        Me.cmbDeprt.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbDeprt.FormattingEnabled = True
        Me.cmbDeprt.Items.AddRange(New Object() {"ລະຫັດບັນຊີ / Account Code", "ຊື່ບັນຊີ (ລາວ) / Ac Name (L)", "ຊື່ບັນຊີ (ອັງກິດ) / Ac Name (E)"})
        Me.cmbDeprt.Location = New System.Drawing.Point(117, 51)
        Me.cmbDeprt.Name = "cmbDeprt"
        Me.cmbDeprt.Size = New System.Drawing.Size(193, 29)
        Me.cmbDeprt.TabIndex = 68
        '
        'txtDep
        '
        Me.txtDep.Enabled = False
        Me.txtDep.Location = New System.Drawing.Point(1126, 97)
        Me.txtDep.Name = "txtDep"
        Me.txtDep.Size = New System.Drawing.Size(77, 34)
        Me.txtDep.TabIndex = 126
        Me.txtDep.Visible = False
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(592, -14)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(85, 34)
        Me.Button2.TabIndex = 127
        Me.Button2.Tag = "3022"
        Me.Button2.Text = "ໂອນຍ້າຍ"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'CmbShow
        '
        Me.CmbShow.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbShow.FormattingEnabled = True
        Me.CmbShow.Items.AddRange(New Object() {"ສະແດງທັງໝົດ", "ລາຍການພວມນຳໃຊ້", "ລາຍການສະສາງແລ້ວ"})
        Me.CmbShow.Location = New System.Drawing.Point(357, 57)
        Me.CmbShow.Name = "CmbShow"
        Me.CmbShow.Size = New System.Drawing.Size(190, 29)
        Me.CmbShow.TabIndex = 136
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(290, 60)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(67, 24)
        Me.Label7.TabIndex = 137
        Me.Label7.Tag = ""
        Me.Label7.Text = "ລາຍການ"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DG
        '
        Me.DG.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DG.Location = New System.Drawing.Point(12, 98)
        Me.DG.Name = "DG"
        Me.DG.Size = New System.Drawing.Size(1253, 290)
        Me.DG.TabIndex = 138
        '
        'Label8
        '
        Me.Label8.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(12, 390)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(1253, 24)
        Me.Label8.TabIndex = 139
        Me.Label8.Tag = ""
        Me.Label8.Text = "0"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(450, 19)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(97, 36)
        Me.Button1.TabIndex = 158
        Me.Button1.Tag = "3008"
        Me.Button1.Text = "Find"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.UseVisualStyleBackColor = False
        '
        'btnPrint
        '
        Me.btnPrint.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnPrint.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnPrint.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.Image = CType(resources.GetObject("btnPrint.Image"), System.Drawing.Image)
        Me.btnPrint.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPrint.Location = New System.Drawing.Point(357, 19)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(92, 36)
        Me.btnPrint.TabIndex = 157
        Me.btnPrint.Tag = "3009"
        Me.btnPrint.Text = "ພິມ"
        Me.btnPrint.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnPrint.UseVisualStyleBackColor = False
        '
        'btnDel
        '
        Me.btnDel.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnDel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnDel.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDel.Image = CType(resources.GetObject("btnDel.Image"), System.Drawing.Image)
        Me.btnDel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDel.Location = New System.Drawing.Point(275, 19)
        Me.btnDel.Name = "btnDel"
        Me.btnDel.Size = New System.Drawing.Size(81, 36)
        Me.btnDel.TabIndex = 156
        Me.btnDel.Tag = "3005"
        Me.btnDel.Text = "ລຶບ"
        Me.btnDel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnDel.UseVisualStyleBackColor = False
        '
        'btnEdit
        '
        Me.btnEdit.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnEdit.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnEdit.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnEdit.Image = CType(resources.GetObject("btnEdit.Image"), System.Drawing.Image)
        Me.btnEdit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnEdit.Location = New System.Drawing.Point(183, 19)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(91, 36)
        Me.btnEdit.TabIndex = 155
        Me.btnEdit.Tag = "3004"
        Me.btnEdit.Text = "ແປງ"
        Me.btnEdit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnEdit.UseVisualStyleBackColor = False
        '
        'btnAdd
        '
        Me.btnAdd.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnAdd.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnAdd.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAdd.Image = CType(resources.GetObject("btnAdd.Image"), System.Drawing.Image)
        Me.btnAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAdd.Location = New System.Drawing.Point(71, 19)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(111, 36)
        Me.btnAdd.TabIndex = 154
        Me.btnAdd.Tag = "3002"
        Me.btnAdd.Text = "ເພີ່ມໃໝ່"
        Me.btnAdd.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnAdd.UseVisualStyleBackColor = False
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Black
        Me.Label9.Location = New System.Drawing.Point(-19, 54)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(135, 24)
        Me.Label9.TabIndex = 162
        Me.Label9.Tag = "2002"
        Me.Label9.Text = "ພະແນກ"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Black
        Me.Label11.Location = New System.Drawing.Point(-42, 19)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(97, 24)
        Me.Label11.TabIndex = 161
        Me.Label11.Tag = "2001"
        Me.Label11.Text = "ສຳນັກງານ"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(1068, 61)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(135, 24)
        Me.Label3.TabIndex = 168
        Me.Label3.Tag = "2006"
        Me.Label3.Text = "ຊື່ຊັບສິນ"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(958, 62)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 24)
        Me.Label2.TabIndex = 170
        Me.Label2.Tag = "2004"
        Me.Label2.Text = "ລະຫັດ"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(811, 61)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 24)
        Me.Label5.TabIndex = 171
        Me.Label5.Tag = "2004"
        Me.Label5.Text = "ລະຫັດ"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(4, 57)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(90, 25)
        Me.Label1.TabIndex = 172
        Me.Label1.Tag = "2003"
        Me.Label1.Text = "ໝວດ"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(552, 24)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(111, 28)
        Me.CheckBox1.TabIndex = 173
        Me.CheckBox1.Text = "ສະເພາະທີ່ດິນ"
        Me.CheckBox1.UseVisualStyleBackColor = True
        Me.CheckBox1.Visible = False
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmbSec)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.cmbDeprt)
        Me.Panel1.Controls.Add(Me.txtSec)
        Me.Panel1.Location = New System.Drawing.Point(567, 152)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(127, 85)
        Me.Panel1.TabIndex = 174
        Me.Panel1.Visible = False
        '
        'FrmAsset_List
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(187, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1277, 414)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnDel)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.DG)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.CmbShow)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.txtDep)
        Me.Controls.Add(Me.txtNo)
        Me.Controls.Add(Me.chkDT)
        Me.Controls.Add(Me.DTUSE)
        Me.Controls.Add(Me.txtNm)
        Me.Controls.Add(Me.txtCode)
        Me.Controls.Add(Me.txtGrp)
        Me.Controls.Add(Me.cmbGrp)
        Me.Controls.Add(Me.lSort)
        Me.Controls.Add(Me.cmbSort)
        Me.Controls.Add(Me.CndClose)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmAsset_List"
        Me.ShowIcon = False
        Me.Text = "Assets List"
        CType(Me.DG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CndClose As System.Windows.Forms.Button
    Friend WithEvents cmbSort As System.Windows.Forms.ComboBox
    Friend WithEvents lSort As System.Windows.Forms.Label
    Friend WithEvents cmbGrp As System.Windows.Forms.ComboBox
    Friend WithEvents txtGrp As System.Windows.Forms.TextBox
    Friend WithEvents txtCode As System.Windows.Forms.TextBox
    Friend WithEvents txtNm As System.Windows.Forms.TextBox
    Friend WithEvents DTUSE As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkDT As System.Windows.Forms.CheckBox
    Friend WithEvents cmbSec As System.Windows.Forms.ComboBox
    Friend WithEvents txtNo As System.Windows.Forms.TextBox
    Friend WithEvents txtSec As System.Windows.Forms.TextBox
    Friend WithEvents cmbDeprt As System.Windows.Forms.ComboBox
    Friend WithEvents txtDep As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents CmbShow As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents DG As System.Windows.Forms.DataGridView
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnDel As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
End Class
