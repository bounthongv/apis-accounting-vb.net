<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmDepartmentNew
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmDepartmentNew))
        Me.LDetail = New System.Windows.Forms.Label
        Me.LEng = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtNm = New System.Windows.Forms.TextBox
        Me.txtID = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtRemark = New System.Windows.Forms.TextBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.CmbCompany = New System.Windows.Forms.ComboBox
        Me.txtCompany = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtNmE = New System.Windows.Forms.TextBox
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdNew = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'LDetail
        '
        Me.LDetail.AutoSize = True
        Me.LDetail.Font = New System.Drawing.Font("Saysettha OT", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LDetail.Location = New System.Drawing.Point(349, 6)
        Me.LDetail.Name = "LDetail"
        Me.LDetail.Size = New System.Drawing.Size(210, 43)
        Me.LDetail.TabIndex = 58
        Me.LDetail.Tag = "2046"
        Me.LDetail.Text = "ລາຍລະອຽດພະແນກ"
        Me.LDetail.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LEng
        '
        Me.LEng.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LEng.Location = New System.Drawing.Point(-27, 103)
        Me.LEng.Name = "LEng"
        Me.LEng.Size = New System.Drawing.Size(223, 24)
        Me.LEng.TabIndex = 71
        Me.LEng.Tag = "2048"
        Me.LEng.Text = "ຊື່ພະແນກ(ລາວ)"
        Me.LEng.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(17, 62)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(179, 24)
        Me.Label1.TabIndex = 63
        Me.Label1.Tag = "2047"
        Me.Label1.Text = "ລະຫັດພະແນກ"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtNm
        '
        Me.txtNm.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNm.Location = New System.Drawing.Point(198, 98)
        Me.txtNm.Name = "txtNm"
        Me.txtNm.Size = New System.Drawing.Size(723, 35)
        Me.txtNm.TabIndex = 59
        '
        'txtID
        '
        Me.txtID.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtID.Location = New System.Drawing.Point(198, 57)
        Me.txtID.Name = "txtID"
        Me.txtID.Size = New System.Drawing.Size(107, 35)
        Me.txtID.TabIndex = 96
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(26, 176)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(170, 24)
        Me.Label2.TabIndex = 103
        Me.Label2.Tag = "2033"
        Me.Label2.Text = "ໝາຍເຫດ"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtRemark
        '
        Me.txtRemark.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRemark.Location = New System.Drawing.Point(198, 176)
        Me.txtRemark.Multiline = True
        Me.txtRemark.Name = "txtRemark"
        Me.txtRemark.Size = New System.Drawing.Size(723, 116)
        Me.txtRemark.TabIndex = 102
        '
        'Label26
        '
        Me.Label26.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.ForeColor = System.Drawing.Color.Black
        Me.Label26.Location = New System.Drawing.Point(330, 62)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(116, 24)
        Me.Label26.TabIndex = 116
        Me.Label26.Tag = "2049"
        Me.Label26.Text = "ສໍານັກງານ"
        Me.Label26.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CmbCompany
        '
        Me.CmbCompany.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbCompany.ForeColor = System.Drawing.Color.Black
        Me.CmbCompany.FormattingEnabled = True
        Me.CmbCompany.Items.AddRange(New Object() {"LAK", "THB", "USD"})
        Me.CmbCompany.Location = New System.Drawing.Point(452, 57)
        Me.CmbCompany.Name = "CmbCompany"
        Me.CmbCompany.Size = New System.Drawing.Size(308, 37)
        Me.CmbCompany.TabIndex = 114
        '
        'txtCompany
        '
        Me.txtCompany.Enabled = False
        Me.txtCompany.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCompany.Location = New System.Drawing.Point(766, 57)
        Me.txtCompany.Name = "txtCompany"
        Me.txtCompany.Size = New System.Drawing.Size(155, 35)
        Me.txtCompany.TabIndex = 113
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(-19, 140)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(215, 24)
        Me.Label3.TabIndex = 118
        Me.Label3.Tag = "2109"
        Me.Label3.Text = "ຊື່ພະແນກ(ອັງກິດ)"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtNmE
        '
        Me.txtNmE.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNmE.Location = New System.Drawing.Point(198, 136)
        Me.txtNmE.Name = "txtNmE"
        Me.txtNmE.Size = New System.Drawing.Size(723, 35)
        Me.txtNmE.TabIndex = 117
        '
        'cmdSave
        '
        Me.cmdSave.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.cmdSave.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdSave.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.Image = CType(resources.GetObject("cmdSave.Image"), System.Drawing.Image)
        Me.cmdSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSave.Location = New System.Drawing.Point(198, 15)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(107, 33)
        Me.cmdSave.TabIndex = 55
        Me.cmdSave.Tag = "3004"
        Me.cmdSave.Text = "ບັນທຶກ"
        Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'cmdNew
        '
        Me.cmdNew.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.cmdNew.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdNew.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdNew.Image = CType(resources.GetObject("cmdNew.Image"), System.Drawing.Image)
        Me.cmdNew.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdNew.Location = New System.Drawing.Point(85, 15)
        Me.cmdNew.Name = "cmdNew"
        Me.cmdNew.Size = New System.Drawing.Size(111, 33)
        Me.cmdNew.TabIndex = 56
        Me.cmdNew.Tag = "3003"
        Me.cmdNew.Text = "ເພີ່ມໃໝ່"
        Me.cmdNew.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdNew.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.Location = New System.Drawing.Point(21, 15)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(57, 33)
        Me.Button1.TabIndex = 57
        Me.Button1.UseVisualStyleBackColor = False
        '
        'FrmDepartmentNew
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(187, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1001, 356)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtNmE)
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.CmbCompany)
        Me.Controls.Add(Me.txtCompany)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtRemark)
        Me.Controls.Add(Me.txtID)
        Me.Controls.Add(Me.LEng)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtNm)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdNew)
        Me.Controls.Add(Me.LDetail)
        Me.Controls.Add(Me.Button1)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "FrmDepartmentNew"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FrmGrpNew"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LDetail As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents cmdNew As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents LEng As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtNm As System.Windows.Forms.TextBox
    Friend WithEvents txtID As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtRemark As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents CmbCompany As System.Windows.Forms.ComboBox
    Friend WithEvents txtCompany As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtNmE As System.Windows.Forms.TextBox
End Class
