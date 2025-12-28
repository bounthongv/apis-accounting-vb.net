<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmList_ASST
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmList_ASST))
        Me.CndClose = New System.Windows.Forms.Button
        Me.btnAdd = New System.Windows.Forms.Button
        Me.cmbSort = New System.Windows.Forms.ComboBox
        Me.lSort = New System.Windows.Forms.Label
        Me.cmbGrp = New System.Windows.Forms.ComboBox
        Me.txtGrp = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.txtCode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtNm = New System.Windows.Forms.TextBox
        Me.DTUSE = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.chkDT = New System.Windows.Forms.CheckBox
        Me.cmbSec = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtNo = New System.Windows.Forms.TextBox
        Me.txtSec = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cmbDeprt = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.DG = New System.Windows.Forms.DataGridView
        Me.txtDep = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        CType(Me.DG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CndClose
        '
        Me.CndClose.Image = CType(resources.GetObject("CndClose.Image"), System.Drawing.Image)
        Me.CndClose.Location = New System.Drawing.Point(5, 2)
        Me.CndClose.Name = "CndClose"
        Me.CndClose.Size = New System.Drawing.Size(60, 50)
        Me.CndClose.TabIndex = 24
        Me.CndClose.UseVisualStyleBackColor = True
        '
        'btnAdd
        '
        Me.btnAdd.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAdd.ForeColor = System.Drawing.Color.Black
        Me.btnAdd.Image = CType(resources.GetObject("btnAdd.Image"), System.Drawing.Image)
        Me.btnAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAdd.Location = New System.Drawing.Point(69, 0)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(101, 52)
        Me.btnAdd.TabIndex = 25
        Me.btnAdd.Text = "ເລືອກ"
        Me.btnAdd.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'cmbSort
        '
        Me.cmbSort.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbSort.FormattingEnabled = True
        Me.cmbSort.Items.AddRange(New Object() {"ລະຫັດ / Code", "ຊື່ / Name", "ວັນທີ / Date"})
        Me.cmbSort.Location = New System.Drawing.Point(1114, 2)
        Me.cmbSort.Name = "cmbSort"
        Me.cmbSort.Size = New System.Drawing.Size(136, 32)
        Me.cmbSort.TabIndex = 48
        '
        'lSort
        '
        Me.lSort.AutoSize = True
        Me.lSort.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lSort.ForeColor = System.Drawing.Color.Black
        Me.lSort.Location = New System.Drawing.Point(1046, 2)
        Me.lSort.Name = "lSort"
        Me.lSort.Size = New System.Drawing.Size(62, 24)
        Me.lSort.TabIndex = 50
        Me.lSort.Text = "ລຽງຕາມ"
        '
        'cmbGrp
        '
        Me.cmbGrp.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbGrp.FormattingEnabled = True
        Me.cmbGrp.Items.AddRange(New Object() {"ລະຫັດບັນຊີ / Account Code", "ຊື່ບັນຊີ (ລາວ) / Ac Name (L)", "ຊື່ບັນຊີ (ອັງກິດ) / Ac Name (E)"})
        Me.cmbGrp.Location = New System.Drawing.Point(69, 54)
        Me.cmbGrp.Name = "cmbGrp"
        Me.cmbGrp.Size = New System.Drawing.Size(328, 32)
        Me.cmbGrp.TabIndex = 51
        '
        'txtGrp
        '
        Me.txtGrp.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGrp.Location = New System.Drawing.Point(439, 18)
        Me.txtGrp.Name = "txtGrp"
        Me.txtGrp.Size = New System.Drawing.Size(57, 22)
        Me.txtGrp.TabIndex = 53
        Me.txtGrp.Visible = False
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(176, 2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 52)
        Me.Button1.TabIndex = 54
        Me.Button1.Text = "Find"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.UseVisualStyleBackColor = True
        '
        'txtCode
        '
        Me.txtCode.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCode.Location = New System.Drawing.Point(933, 4)
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(107, 30)
        Me.txtCode.TabIndex = 55
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(876, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 24)
        Me.Label1.TabIndex = 56
        Me.Label1.Text = "ລະຫັດ"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(1048, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 24)
        Me.Label2.TabIndex = 58
        Me.Label2.Text = "ຊື່ຊັບສິນ"
        '
        'txtNm
        '
        Me.txtNm.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNm.Location = New System.Drawing.Point(1113, 36)
        Me.txtNm.Name = "txtNm"
        Me.txtNm.Size = New System.Drawing.Size(137, 30)
        Me.txtNm.TabIndex = 57
        '
        'DTUSE
        '
        Me.DTUSE.CustomFormat = "MM/yyyy"
        Me.DTUSE.Enabled = False
        Me.DTUSE.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DTUSE.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTUSE.Location = New System.Drawing.Point(781, 54)
        Me.DTUSE.Name = "DTUSE"
        Me.DTUSE.ShowUpDown = True
        Me.DTUSE.Size = New System.Drawing.Size(89, 26)
        Me.DTUSE.TabIndex = 60
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(18, 57)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(45, 24)
        Me.Label4.TabIndex = 61
        Me.Label4.Text = "ໝວດ"
        '
        'chkDT
        '
        Me.chkDT.AutoSize = True
        Me.chkDT.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDT.Location = New System.Drawing.Point(678, 53)
        Me.chkDT.Name = "chkDT"
        Me.chkDT.Size = New System.Drawing.Size(97, 28)
        Me.chkDT.TabIndex = 62
        Me.chkDT.Text = "ປ.ດ ນຳໃຊ້"
        Me.chkDT.UseVisualStyleBackColor = True
        '
        'cmbSec
        '
        Me.cmbSec.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbSec.FormattingEnabled = True
        Me.cmbSec.Items.AddRange(New Object() {"ລະຫັດບັນຊີ / Account Code", "ຊື່ບັນຊີ (ລາວ) / Ac Name (L)", "ຊື່ບັນຊີ (ອັງກິດ) / Ac Name (E)"})
        Me.cmbSec.Location = New System.Drawing.Point(731, -6)
        Me.cmbSec.Name = "cmbSec"
        Me.cmbSec.Size = New System.Drawing.Size(241, 32)
        Me.cmbSec.TabIndex = 63
        Me.cmbSec.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(876, 43)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(51, 24)
        Me.Label5.TabIndex = 66
        Me.Label5.Text = "ລະຫັດ"
        '
        'txtNo
        '
        Me.txtNo.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNo.Location = New System.Drawing.Point(933, 40)
        Me.txtNo.Name = "txtNo"
        Me.txtNo.Size = New System.Drawing.Size(107, 30)
        Me.txtNo.TabIndex = 65
        '
        'txtSec
        '
        Me.txtSec.Location = New System.Drawing.Point(552, 40)
        Me.txtSec.Name = "txtSec"
        Me.txtSec.Size = New System.Drawing.Size(41, 34)
        Me.txtSec.TabIndex = 67
        Me.txtSec.Visible = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(357, -1)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(59, 24)
        Me.Label6.TabIndex = 69
        Me.Label6.Text = "ພະແນກ"
        Me.Label6.Visible = False
        '
        'cmbDeprt
        '
        Me.cmbDeprt.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbDeprt.FormattingEnabled = True
        Me.cmbDeprt.Items.AddRange(New Object() {"ລະຫັດບັນຊີ / Account Code", "ຊື່ບັນຊີ (ລາວ) / Ac Name (L)", "ຊື່ບັນຊີ (ອັງກິດ) / Ac Name (E)"})
        Me.cmbDeprt.Location = New System.Drawing.Point(422, -3)
        Me.cmbDeprt.Name = "cmbDeprt"
        Me.cmbDeprt.Size = New System.Drawing.Size(264, 32)
        Me.cmbDeprt.TabIndex = 68
        Me.cmbDeprt.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(657, -3)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(74, 24)
        Me.Label3.TabIndex = 71
        Me.Label3.Text = "ສຳນັກງານ"
        Me.Label3.Visible = False
        '
        'DG
        '
        Me.DG.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DG.Location = New System.Drawing.Point(5, 92)
        Me.DG.Name = "DG"
        Me.DG.Size = New System.Drawing.Size(1254, 399)
        Me.DG.TabIndex = 72
        '
        'txtDep
        '
        Me.txtDep.Location = New System.Drawing.Point(439, 47)
        Me.txtDep.Name = "txtDep"
        Me.txtDep.Size = New System.Drawing.Size(41, 34)
        Me.txtDep.TabIndex = 67
        Me.txtDep.Visible = False
        '
        'Label8
        '
        Me.Label8.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(5, 493)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(1254, 23)
        Me.Label8.TabIndex = 140
        Me.Label8.Tag = ""
        Me.Label8.Text = "0"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'FrmList_ASST
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(187, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1262, 516)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.DG)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.cmbDeprt)
        Me.Controls.Add(Me.txtDep)
        Me.Controls.Add(Me.txtSec)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtNo)
        Me.Controls.Add(Me.cmbSec)
        Me.Controls.Add(Me.chkDT)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.DTUSE)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtNm)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtCode)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.txtGrp)
        Me.Controls.Add(Me.cmbGrp)
        Me.Controls.Add(Me.lSort)
        Me.Controls.Add(Me.cmbSort)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.CndClose)
        Me.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmList_ASST"
        Me.ShowIcon = False
        Me.Text = "Assets List"
        CType(Me.DG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CndClose As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents cmbSort As System.Windows.Forms.ComboBox
    Friend WithEvents lSort As System.Windows.Forms.Label
    Friend WithEvents cmbGrp As System.Windows.Forms.ComboBox
    Friend WithEvents txtGrp As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents txtCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtNm As System.Windows.Forms.TextBox
    Friend WithEvents DTUSE As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents chkDT As System.Windows.Forms.CheckBox
    Friend WithEvents cmbSec As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtNo As System.Windows.Forms.TextBox
    Friend WithEvents txtSec As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cmbDeprt As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DG As System.Windows.Forms.DataGridView
    Friend WithEvents txtDep As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
End Class
