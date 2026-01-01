<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmUser
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmUser))
        Me.txtUsr_id = New System.Windows.Forms.TextBox
        Me.txtConfrim = New System.Windows.Forms.TextBox
        Me.txtPWD = New System.Windows.Forms.TextBox
        Me.txtUsr_nm = New System.Windows.Forms.TextBox
        Me.cmbpermision = New System.Windows.Forms.ComboBox
        Me.cmbUsrPermit = New System.Windows.Forms.ComboBox
        Me.lblID = New System.Windows.Forms.Label
        Me.lblNm = New System.Windows.Forms.Label
        Me.lblPass = New System.Windows.Forms.Label
        Me.lblConfirm = New System.Windows.Forms.Label
        Me.lblPermission = New System.Windows.Forms.Label
        Me.lblSec = New System.Windows.Forms.Label
        Me.lblPermissions = New System.Windows.Forms.Label
        Me.txtDep_ID = New System.Windows.Forms.TextBox
        Me.txtDep_Nm = New System.Windows.Forms.TextBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.Fg = New System.Windows.Forms.DataGridView
        Me.Button1 = New System.Windows.Forms.Button
        Me.BtnDel = New System.Windows.Forms.Button
        Me.BtnSave = New System.Windows.Forms.Button
        Me.BtnAddNew = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.FgSec = New System.Windows.Forms.DataGridView
        Me.Label8 = New System.Windows.Forms.Label
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtConfrimPass = New System.Windows.Forms.TextBox
        Me.txtNewPass = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtOldPass = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.FgItem = New System.Windows.Forms.DataGridView
        Me.Label4 = New System.Windows.Forms.Label
        Me.cmbCompany = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Sub_Company = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.CheckBox3 = New System.Windows.Forms.CheckBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Button6 = New System.Windows.Forms.Button
        CType(Me.Fg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.FgSec, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        CType(Me.FgItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtUsr_id
        '
        Me.txtUsr_id.Location = New System.Drawing.Point(82, 35)
        Me.txtUsr_id.Name = "txtUsr_id"
        Me.txtUsr_id.Size = New System.Drawing.Size(137, 30)
        Me.txtUsr_id.TabIndex = 0
        '
        'txtConfrim
        '
        Me.txtConfrim.Location = New System.Drawing.Point(616, 3)
        Me.txtConfrim.Name = "txtConfrim"
        Me.txtConfrim.PasswordChar = Global.Microsoft.VisualBasic.ChrW(64)
        Me.txtConfrim.Size = New System.Drawing.Size(132, 30)
        Me.txtConfrim.TabIndex = 69
        '
        'txtPWD
        '
        Me.txtPWD.Location = New System.Drawing.Point(309, 35)
        Me.txtPWD.Name = "txtPWD"
        Me.txtPWD.PasswordChar = Global.Microsoft.VisualBasic.ChrW(64)
        Me.txtPWD.Size = New System.Drawing.Size(113, 30)
        Me.txtPWD.TabIndex = 70
        '
        'txtUsr_nm
        '
        Me.txtUsr_nm.Location = New System.Drawing.Point(82, 67)
        Me.txtUsr_nm.Name = "txtUsr_nm"
        Me.txtUsr_nm.Size = New System.Drawing.Size(340, 30)
        Me.txtUsr_nm.TabIndex = 71
        '
        'cmbpermision
        '
        Me.cmbpermision.FormattingEnabled = True
        Me.cmbpermision.Items.AddRange(New Object() {"Admin", "Sub-Admin", "Border-Admin", "User"})
        Me.cmbpermision.Location = New System.Drawing.Point(845, 4)
        Me.cmbpermision.Name = "cmbpermision"
        Me.cmbpermision.Size = New System.Drawing.Size(123, 29)
        Me.cmbpermision.TabIndex = 75
        Me.cmbpermision.Text = "Admin"
        '
        'cmbUsrPermit
        '
        Me.cmbUsrPermit.FormattingEnabled = True
        Me.cmbUsrPermit.Items.AddRange(New Object() {"Administrator", "User"})
        Me.cmbUsrPermit.Location = New System.Drawing.Point(403, -33)
        Me.cmbUsrPermit.Name = "cmbUsrPermit"
        Me.cmbUsrPermit.Size = New System.Drawing.Size(133, 29)
        Me.cmbUsrPermit.TabIndex = 77
        Me.cmbUsrPermit.Text = "Administrator"
        Me.cmbUsrPermit.Visible = False
        '
        'lblID
        '
        Me.lblID.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblID.Location = New System.Drawing.Point(-2, 38)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(84, 26)
        Me.lblID.TabIndex = 78
        Me.lblID.Text = "ລະຫັດຜູ້ໃຊ້:"
        Me.lblID.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblNm
        '
        Me.lblNm.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNm.Location = New System.Drawing.Point(4, 70)
        Me.lblNm.Name = "lblNm"
        Me.lblNm.Size = New System.Drawing.Size(73, 27)
        Me.lblNm.TabIndex = 79
        Me.lblNm.Text = "ຊື່ຜູ້ໃຊ້:"
        Me.lblNm.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPass
        '
        Me.lblPass.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPass.Location = New System.Drawing.Point(220, 38)
        Me.lblPass.Name = "lblPass"
        Me.lblPass.Size = New System.Drawing.Size(85, 21)
        Me.lblPass.TabIndex = 80
        Me.lblPass.Text = "ລະຫັດຜ່ານ:"
        Me.lblPass.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblConfirm
        '
        Me.lblConfirm.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblConfirm.Location = New System.Drawing.Point(496, 7)
        Me.lblConfirm.Name = "lblConfirm"
        Me.lblConfirm.Size = New System.Drawing.Size(123, 21)
        Me.lblConfirm.TabIndex = 81
        Me.lblConfirm.Text = "ຢໍ້າຄືນລະຫັດຜ່ານ:"
        Me.lblConfirm.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPermission
        '
        Me.lblPermission.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPermission.Location = New System.Drawing.Point(295, -30)
        Me.lblPermission.Name = "lblPermission"
        Me.lblPermission.Size = New System.Drawing.Size(101, 21)
        Me.lblPermission.TabIndex = 82
        Me.lblPermission.Text = "ສິດໃຊ້ໂປຣແກຣມ:"
        Me.lblPermission.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblPermission.Visible = False
        '
        'lblSec
        '
        Me.lblSec.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSec.Location = New System.Drawing.Point(522, 39)
        Me.lblSec.Name = "lblSec"
        Me.lblSec.Size = New System.Drawing.Size(97, 21)
        Me.lblSec.TabIndex = 83
        Me.lblSec.Text = "ພະແນກ:"
        Me.lblSec.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPermissions
        '
        Me.lblPermissions.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPermissions.Location = New System.Drawing.Point(767, 7)
        Me.lblPermissions.Name = "lblPermissions"
        Me.lblPermissions.Size = New System.Drawing.Size(72, 20)
        Me.lblPermissions.TabIndex = 85
        Me.lblPermissions.Text = "ສິດທິນຳໃຊ້"
        Me.lblPermissions.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtDep_ID
        '
        Me.txtDep_ID.Location = New System.Drawing.Point(616, 35)
        Me.txtDep_ID.Name = "txtDep_ID"
        Me.txtDep_ID.Size = New System.Drawing.Size(98, 30)
        Me.txtDep_ID.TabIndex = 93
        '
        'txtDep_Nm
        '
        Me.txtDep_Nm.BackColor = System.Drawing.Color.White
        Me.txtDep_Nm.Location = New System.Drawing.Point(748, 35)
        Me.txtDep_Nm.Name = "txtDep_Nm"
        Me.txtDep_Nm.ReadOnly = True
        Me.txtDep_Nm.Size = New System.Drawing.Size(220, 30)
        Me.txtDep_Nm.TabIndex = 95
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button3.Location = New System.Drawing.Point(309, 2)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(113, 35)
        Me.Button3.TabIndex = 96
        Me.Button3.Text = "ປ່ຽນລະຫັດຜ່ານ"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Fg
        '
        Me.Fg.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Fg.DataSource = Nothing
        Me.Fg.Location = New System.Drawing.Point(4, 99)
        Me.Fg.Name = "Fg"

        Me.Fg.Size = New System.Drawing.Size(964, 420)
        Me.Fg.TabIndex = 86
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(715, 35)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(33, 29)
        Me.Button1.TabIndex = 94
        Me.Button1.UseVisualStyleBackColor = True
        '
        'BtnDel
        '
        Me.BtnDel.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnDel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnDel.Location = New System.Drawing.Point(220, 2)
        Me.BtnDel.Name = "BtnDel"
        Me.BtnDel.Size = New System.Drawing.Size(89, 35)
        Me.BtnDel.TabIndex = 91
        Me.BtnDel.Text = "ລຶບ"
        Me.BtnDel.UseVisualStyleBackColor = True
        '
        'BtnSave
        '
        Me.BtnSave.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnSave.Location = New System.Drawing.Point(135, 2)
        Me.BtnSave.Name = "BtnSave"
        Me.BtnSave.Size = New System.Drawing.Size(84, 35)
        Me.BtnSave.TabIndex = 90
        Me.BtnSave.Text = "ບັນທຶກ"
        Me.BtnSave.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnSave.UseVisualStyleBackColor = True
        '
        'BtnAddNew
        '
        Me.BtnAddNew.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAddNew.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAddNew.Location = New System.Drawing.Point(48, 2)
        Me.BtnAddNew.Name = "BtnAddNew"
        Me.BtnAddNew.Size = New System.Drawing.Size(86, 35)
        Me.BtnAddNew.TabIndex = 89
        Me.BtnAddNew.Text = "ເພີ່ມໃໝ່"
        Me.BtnAddNew.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnAddNew.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(4, 2)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(35, 35)
        Me.Button2.TabIndex = 88
        Me.Button2.Tag = "999"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.FgSec)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Location = New System.Drawing.Point(616, 153)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(352, 153)
        Me.Panel1.TabIndex = 97
        Me.Panel1.Visible = False
        '
        'FgSec
        '
        Me.FgSec.DataSource = Nothing
        Me.FgSec.Location = New System.Drawing.Point(2, 16)
        Me.FgSec.Name = "FgSec"
        Me.FgSec.Size = New System.Drawing.Size(410, 254)
        Me.FgSec.TabIndex = 45490
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Blue
        Me.Label8.Location = New System.Drawing.Point(-6, -6)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(399, 29)
        Me.Label8.TabIndex = 45488
        Me.Label8.Text = "ລາຍການພາກສ່ວນ"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.Button5)
        Me.Panel3.Controls.Add(Me.Button4)
        Me.Panel3.Controls.Add(Me.Label3)
        Me.Panel3.Controls.Add(Me.txtConfrimPass)
        Me.Panel3.Controls.Add(Me.txtNewPass)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Controls.Add(Me.txtOldPass)
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Location = New System.Drawing.Point(160, 99)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(262, 135)
        Me.Panel3.TabIndex = 98
        Me.Panel3.Visible = False
        '
        'Button5
        '
        Me.Button5.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button5.Location = New System.Drawing.Point(2, 99)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(124, 31)
        Me.Button5.TabIndex = 104
        Me.Button5.Text = "ບັນທຶກ"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button4.Location = New System.Drawing.Point(126, 99)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(131, 31)
        Me.Button4.TabIndex = 99
        Me.Button4.Text = "ອອກ"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(-2, 70)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(121, 21)
        Me.Label3.TabIndex = 100
        Me.Label3.Text = "ຢໍ້າຄືນລະຫັດຜ່ານ:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtConfrimPass
        '
        Me.txtConfrimPass.Location = New System.Drawing.Point(125, 67)
        Me.txtConfrimPass.Name = "txtConfrimPass"
        Me.txtConfrimPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(64)
        Me.txtConfrimPass.Size = New System.Drawing.Size(133, 30)
        Me.txtConfrimPass.TabIndex = 99
        '
        'txtNewPass
        '
        Me.txtNewPass.Location = New System.Drawing.Point(125, 35)
        Me.txtNewPass.Name = "txtNewPass"
        Me.txtNewPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(64)
        Me.txtNewPass.Size = New System.Drawing.Size(133, 30)
        Me.txtNewPass.TabIndex = 103
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(3, 38)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(118, 21)
        Me.Label2.TabIndex = 102
        Me.Label2.Text = "ລະຫັດຜ່ານໃໝ່:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtOldPass
        '
        Me.txtOldPass.Location = New System.Drawing.Point(125, 3)
        Me.txtOldPass.Name = "txtOldPass"
        Me.txtOldPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(64)
        Me.txtOldPass.Size = New System.Drawing.Size(133, 30)
        Me.txtOldPass.TabIndex = 101
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(118, 21)
        Me.Label1.TabIndex = 100
        Me.Label1.Text = "ລະຫັດຜ່ານເກົ່າ:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel4
        '
        Me.Panel4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel4.Controls.Add(Me.FgItem)
        Me.Panel4.Controls.Add(Me.Label4)
        Me.Panel4.Location = New System.Drawing.Point(616, 308)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(352, 211)
        Me.Panel4.TabIndex = 99
        Me.Panel4.Visible = False
        '
        'FgItem
        '
        Me.FgItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FgItem.DataSource = Nothing
        Me.FgItem.Location = New System.Drawing.Point(3, 19)
        Me.FgItem.Name = "FgItem"

        Me.FgItem.Size = New System.Drawing.Size(392, 187)
        Me.FgItem.TabIndex = 66
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Blue
        Me.Label4.Location = New System.Drawing.Point(-1, -7)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(399, 29)
        Me.Label4.TabIndex = 45489
        Me.Label4.Text = "ລາຍການເມນູ"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmbCompany
        '
        Me.cmbCompany.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbCompany.FormattingEnabled = True
        Me.cmbCompany.Location = New System.Drawing.Point(616, 67)
        Me.cmbCompany.Name = "cmbCompany"
        Me.cmbCompany.Size = New System.Drawing.Size(132, 29)
        Me.cmbCompany.TabIndex = 45517
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(513, 70)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(97, 21)
        Me.Label5.TabIndex = 45518
        Me.Label5.Text = "ສາຂາ:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Sub_Company
        '
        Me.Sub_Company.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Sub_Company.FormattingEnabled = True
        Me.Sub_Company.Location = New System.Drawing.Point(845, 68)
        Me.Sub_Company.Name = "Sub_Company"
        Me.Sub_Company.Size = New System.Drawing.Size(123, 29)
        Me.Sub_Company.TabIndex = 45522
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(767, 71)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(72, 20)
        Me.Label6.TabIndex = 45523
        Me.Label6.Text = "ເຂດ"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.CheckBox3)
        Me.Panel2.Controls.Add(Me.CheckBox2)
        Me.Panel2.Controls.Add(Me.CheckBox1)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Location = New System.Drawing.Point(616, 99)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(352, 52)
        Me.Panel2.TabIndex = 45524
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Location = New System.Drawing.Point(254, 21)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(69, 25)
        Me.CheckBox3.TabIndex = 45494
        Me.CheckBox3.Text = "ລຶບຂໍ້ມີນ"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(139, 19)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(83, 25)
        Me.CheckBox2.TabIndex = 45493
        Me.CheckBox2.Text = "ແກ້ໄຂຂໍ້ມູນ"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(34, 19)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(76, 25)
        Me.CheckBox1.TabIndex = 45492
        Me.CheckBox1.Text = "ປ້ອນຂໍ້ມູນ"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Blue
        Me.Label7.Location = New System.Drawing.Point(-16, -6)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(399, 29)
        Me.Label7.TabIndex = 45491
        Me.Label7.Text = "ການອະນຸມັດ"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button6
        '
        Me.Button6.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button6.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button6.Location = New System.Drawing.Point(422, 2)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(98, 35)
        Me.Button6.TabIndex = 45525
        Me.Button6.Text = "ຮັບຜິດຊອບ"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'FrmUser
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 21.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(187, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(972, 523)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Sub_Company)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.cmbCompany)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.txtDep_Nm)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.txtDep_ID)
        Me.Controls.Add(Me.BtnDel)
        Me.Controls.Add(Me.BtnSave)
        Me.Controls.Add(Me.BtnAddNew)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Fg)
        Me.Controls.Add(Me.lblPermissions)
        Me.Controls.Add(Me.lblSec)
        Me.Controls.Add(Me.lblPermission)
        Me.Controls.Add(Me.lblPass)
        Me.Controls.Add(Me.lblNm)
        Me.Controls.Add(Me.lblID)
        Me.Controls.Add(Me.cmbUsrPermit)
        Me.Controls.Add(Me.cmbpermision)
        Me.Controls.Add(Me.txtUsr_nm)
        Me.Controls.Add(Me.txtPWD)
        Me.Controls.Add(Me.txtConfrim)
        Me.Controls.Add(Me.txtUsr_id)
        Me.Controls.Add(Me.lblConfirm)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "FrmUser"
        Me.Text = "List of user"
        CType(Me.Fg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        CType(Me.FgSec, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        CType(Me.FgItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtUsr_id As System.Windows.Forms.TextBox
    Friend WithEvents txtConfrim As System.Windows.Forms.TextBox
    Friend WithEvents txtPWD As System.Windows.Forms.TextBox
    Friend WithEvents txtUsr_nm As System.Windows.Forms.TextBox
    Friend WithEvents cmbpermision As System.Windows.Forms.ComboBox
    Friend WithEvents cmbUsrPermit As System.Windows.Forms.ComboBox
    Friend WithEvents lblID As System.Windows.Forms.Label
    Friend WithEvents lblNm As System.Windows.Forms.Label
    Friend WithEvents lblPass As System.Windows.Forms.Label
    Friend WithEvents lblConfirm As System.Windows.Forms.Label
    Friend WithEvents lblPermission As System.Windows.Forms.Label
    Friend WithEvents lblSec As System.Windows.Forms.Label
    Friend WithEvents lblPermissions As System.Windows.Forms.Label
    Friend WithEvents BtnDel As System.Windows.Forms.Button
    Friend WithEvents BtnSave As System.Windows.Forms.Button
    Friend WithEvents BtnAddNew As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents txtDep_ID As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents txtDep_Nm As System.Windows.Forms.TextBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Fg As System.Windows.Forms.DataGridView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtConfrimPass As System.Windows.Forms.TextBox
    Friend WithEvents txtNewPass As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtOldPass As System.Windows.Forms.TextBox
    Friend WithEvents FgSec As System.Windows.Forms.DataGridView
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents FgItem As System.Windows.Forms.DataGridView
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmbCompany As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Sub_Company As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Button6 As System.Windows.Forms.Button
End Class
