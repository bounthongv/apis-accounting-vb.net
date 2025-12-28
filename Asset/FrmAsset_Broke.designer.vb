<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmAsset_Broke
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmAsset_Broke))
        Me.cmbSort = New System.Windows.Forms.ComboBox
        Me.txtGrp = New System.Windows.Forms.TextBox
        Me.cmbGrp = New System.Windows.Forms.ComboBox
        Me.txtCode = New System.Windows.Forms.TextBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.cmbDeprt = New System.Windows.Forms.ComboBox
        Me.cmbSec = New System.Windows.Forms.ComboBox
        Me.DG = New System.Windows.Forms.DataGridView
        Me.txtSec = New System.Windows.Forms.TextBox
        Me.txtDep = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnDel = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnAdd = New System.Windows.Forms.Button
        Me.CndClose = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Dt = New System.Windows.Forms.DateTimePicker
        Me.Dst = New System.Windows.Forms.DateTimePicker
        CType(Me.DG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmbSort
        '
        Me.cmbSort.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbSort.FormattingEnabled = True
        Me.cmbSort.Items.AddRange(New Object() {"ເລກທີສະສາງ / BrokenID", "ລະຫັດ / AssetID", "ຊື່ລາຍການ / Asset Name"})
        Me.cmbSort.Location = New System.Drawing.Point(756, 70)
        Me.cmbSort.Name = "cmbSort"
        Me.cmbSort.Size = New System.Drawing.Size(231, 32)
        Me.cmbSort.TabIndex = 58
        '
        'txtGrp
        '
        Me.txtGrp.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGrp.Location = New System.Drawing.Point(1143, 42)
        Me.txtGrp.Name = "txtGrp"
        Me.txtGrp.Size = New System.Drawing.Size(57, 22)
        Me.txtGrp.TabIndex = 63
        Me.txtGrp.Visible = False
        '
        'cmbGrp
        '
        Me.cmbGrp.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbGrp.FormattingEnabled = True
        Me.cmbGrp.Items.AddRange(New Object() {"ລະຫັດບັນຊີ / Account Code", "ຊື່ບັນຊີ (ລາວ) / Ac Name (L)", "ຊື່ບັນຊີ (ອັງກິດ) / Ac Name (E)"})
        Me.cmbGrp.Location = New System.Drawing.Point(382, 70)
        Me.cmbGrp.Name = "cmbGrp"
        Me.cmbGrp.Size = New System.Drawing.Size(245, 32)
        Me.cmbGrp.TabIndex = 62
        '
        'txtCode
        '
        Me.txtCode.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCode.Location = New System.Drawing.Point(632, 30)
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(121, 30)
        Me.txtCode.TabIndex = 65
        Me.txtCode.Visible = False
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(1152, 75)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 67
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'cmbDeprt
        '
        Me.cmbDeprt.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbDeprt.FormattingEnabled = True
        Me.cmbDeprt.Items.AddRange(New Object() {"ລະຫັດບັນຊີ / Account Code", "ຊື່ບັນຊີ (ລາວ) / Ac Name (L)", "ຊື່ບັນຊີ (ອັງກິດ) / Ac Name (E)"})
        Me.cmbDeprt.Location = New System.Drawing.Point(1013, 23)
        Me.cmbDeprt.Name = "cmbDeprt"
        Me.cmbDeprt.Size = New System.Drawing.Size(283, 32)
        Me.cmbDeprt.TabIndex = 73
        Me.cmbDeprt.Visible = False
        '
        'cmbSec
        '
        Me.cmbSec.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbSec.FormattingEnabled = True
        Me.cmbSec.Items.AddRange(New Object() {"ລະຫັດບັນຊີ / Account Code", "ຊື່ບັນຊີ (ລາວ) / Ac Name (L)", "ຊື່ບັນຊີ (ອັງກິດ) / Ac Name (E)"})
        Me.cmbSec.Location = New System.Drawing.Point(991, -1)
        Me.cmbSec.Name = "cmbSec"
        Me.cmbSec.Size = New System.Drawing.Size(223, 32)
        Me.cmbSec.TabIndex = 72
        Me.cmbSec.Visible = False
        '
        'DG
        '
        Me.DG.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DG.Location = New System.Drawing.Point(4, 108)
        Me.DG.Name = "DG"
        Me.DG.Size = New System.Drawing.Size(1262, 294)
        Me.DG.TabIndex = 78
        '
        'txtSec
        '
        Me.txtSec.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSec.Location = New System.Drawing.Point(322, 104)
        Me.txtSec.Name = "txtSec"
        Me.txtSec.Size = New System.Drawing.Size(57, 22)
        Me.txtSec.TabIndex = 79
        Me.txtSec.Visible = False
        '
        'txtDep
        '
        Me.txtDep.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDep.Location = New System.Drawing.Point(385, 108)
        Me.txtDep.Name = "txtDep"
        Me.txtDep.Size = New System.Drawing.Size(57, 22)
        Me.txtDep.TabIndex = 80
        Me.txtDep.Visible = False
        '
        'Label8
        '
        Me.Label8.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(4, 405)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(1262, 23)
        Me.Label8.TabIndex = 142
        Me.Label8.Tag = ""
        Me.Label8.Text = "0"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Black
        Me.Label11.Location = New System.Drawing.Point(890, 3)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(97, 24)
        Me.Label11.TabIndex = 159
        Me.Label11.Tag = "2001"
        Me.Label11.Text = "ສຳນັກງານ"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label11.Visible = False
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(873, 26)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(134, 24)
        Me.Label3.TabIndex = 160
        Me.Label3.Tag = "2002"
        Me.Label3.Text = "ພະແນກ"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label3.Visible = False
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(559, 33)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(68, 24)
        Me.Label5.TabIndex = 161
        Me.Label5.Tag = "2004"
        Me.Label5.Text = "ລະຫັດ"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label5.Visible = False
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(284, 73)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(92, 24)
        Me.Label2.TabIndex = 162
        Me.Label2.Tag = "2003"
        Me.Label2.Text = "ໝວດ"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(634, 74)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(116, 24)
        Me.Label4.TabIndex = 163
        Me.Label4.Tag = "2007"
        Me.Label4.Text = "ລຽງຕາມ"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Saysettha OT", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(722, -8)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(186, 38)
        Me.Label6.TabIndex = 164
        Me.Label6.Tag = "2050"
        Me.Label6.Text = "ລາຍການສະສາງ"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(470, 13)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(82, 52)
        Me.Button1.TabIndex = 158
        Me.Button1.Tag = "3008"
        Me.Button1.Text = "Find"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.Image = CType(resources.GetObject("btnPrint.Image"), System.Drawing.Image)
        Me.btnPrint.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPrint.Location = New System.Drawing.Point(382, 13)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(86, 52)
        Me.btnPrint.TabIndex = 157
        Me.btnPrint.Tag = "3009"
        Me.btnPrint.Text = "ພິມ"
        Me.btnPrint.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnDel
        '
        Me.btnDel.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDel.Image = CType(resources.GetObject("btnDel.Image"), System.Drawing.Image)
        Me.btnDel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDel.Location = New System.Drawing.Point(292, 12)
        Me.btnDel.Name = "btnDel"
        Me.btnDel.Size = New System.Drawing.Size(87, 52)
        Me.btnDel.TabIndex = 156
        Me.btnDel.Tag = "3005"
        Me.btnDel.Text = "ລຶບ"
        Me.btnDel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnDel.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnEdit.Image = CType(resources.GetObject("btnEdit.Image"), System.Drawing.Image)
        Me.btnEdit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnEdit.Location = New System.Drawing.Point(182, 12)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(109, 52)
        Me.btnEdit.TabIndex = 155
        Me.btnEdit.Tag = "3004"
        Me.btnEdit.Text = "ແປງ"
        Me.btnEdit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnAdd
        '
        Me.btnAdd.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAdd.Image = CType(resources.GetObject("btnAdd.Image"), System.Drawing.Image)
        Me.btnAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAdd.Location = New System.Drawing.Point(62, 12)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(120, 52)
        Me.btnAdd.TabIndex = 154
        Me.btnAdd.Tag = "3002"
        Me.btnAdd.Text = "ເພີ່ມໃໝ່"
        Me.btnAdd.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'CndClose
        '
        Me.CndClose.Image = CType(resources.GetObject("CndClose.Image"), System.Drawing.Image)
        Me.CndClose.Location = New System.Drawing.Point(4, 12)
        Me.CndClose.Name = "CndClose"
        Me.CndClose.Size = New System.Drawing.Size(55, 52)
        Me.CndClose.TabIndex = 24
        Me.CndClose.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(184, 74)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(24, 21)
        Me.Label1.TabIndex = 45553
        Me.Label1.Text = "To"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(0, 75)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(73, 21)
        Me.Label7.TabIndex = 45552
        Me.Label7.Text = "Start Date"
        '
        'Dt
        '
        Me.Dt.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.Dt.Location = New System.Drawing.Point(212, 69)
        Me.Dt.Name = "Dt"
        Me.Dt.Size = New System.Drawing.Size(103, 30)
        Me.Dt.TabIndex = 45551
        '
        'Dst
        '
        Me.Dst.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.Dst.Location = New System.Drawing.Point(79, 70)
        Me.Dst.Name = "Dst"
        Me.Dst.Size = New System.Drawing.Size(103, 30)
        Me.Dst.TabIndex = 45550
        '
        'FrmAsset_Broke
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(187, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1268, 428)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Dt)
        Me.Controls.Add(Me.Dst)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnDel)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txtDep)
        Me.Controls.Add(Me.txtSec)
        Me.Controls.Add(Me.DG)
        Me.Controls.Add(Me.cmbDeprt)
        Me.Controls.Add(Me.cmbSec)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.txtCode)
        Me.Controls.Add(Me.txtGrp)
        Me.Controls.Add(Me.cmbGrp)
        Me.Controls.Add(Me.cmbSort)
        Me.Controls.Add(Me.CndClose)
        Me.Controls.Add(Me.Label4)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmAsset_Broke"
        Me.ShowIcon = False
        Me.Text = "Assets Broken"
        CType(Me.DG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CndClose As System.Windows.Forms.Button
    Friend WithEvents cmbSort As System.Windows.Forms.ComboBox
    Friend WithEvents txtGrp As System.Windows.Forms.TextBox
    Friend WithEvents cmbGrp As System.Windows.Forms.ComboBox
    Friend WithEvents txtCode As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents cmbDeprt As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSec As System.Windows.Forms.ComboBox
    Friend WithEvents DG As System.Windows.Forms.DataGridView
    Friend WithEvents txtSec As System.Windows.Forms.TextBox
    Friend WithEvents txtDep As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnDel As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Dt As System.Windows.Forms.DateTimePicker
    Friend WithEvents Dst As System.Windows.Forms.DateTimePicker
End Class
