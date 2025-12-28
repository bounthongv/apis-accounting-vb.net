<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrNewAcc
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrNewAcc))
        Me.txtAc_code = New System.Windows.Forms.TextBox
        Me.txtAccName = New System.Windows.Forms.TextBox
        Me.txtAccName_E = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.DtmDate = New System.Windows.Forms.DateTimePicker
        Me.BtnSave = New System.Windows.Forms.Button
        Me.BtnExit = New System.Windows.Forms.Button
        Me.BntNew = New System.Windows.Forms.Button
        Me.CFi = New System.Windows.Forms.RadioButton
        Me.CAs = New System.Windows.Forms.RadioButton
        Me.CEx = New System.Windows.Forms.RadioButton
        Me.CIn = New System.Windows.Forms.RadioButton
        Me.TxtWISE_Orginal = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'txtAc_code
        '
        Me.txtAc_code.Location = New System.Drawing.Point(82, 41)
        Me.txtAc_code.MaxLength = 20
        Me.txtAc_code.Name = "txtAc_code"
        Me.txtAc_code.Size = New System.Drawing.Size(143, 30)
        Me.txtAc_code.TabIndex = 0
        Me.txtAc_code.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtAccName
        '
        Me.txtAccName.Location = New System.Drawing.Point(82, 77)
        Me.txtAccName.Name = "txtAccName"
        Me.txtAccName.Size = New System.Drawing.Size(510, 30)
        Me.txtAccName.TabIndex = 1
        '
        'txtAccName_E
        '
        Me.txtAccName_E.Location = New System.Drawing.Point(82, 113)
        Me.txtAccName_E.Name = "txtAccName_E"
        Me.txtAccName_E.Size = New System.Drawing.Size(510, 30)
        Me.txtAccName_E.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 46)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 21)
        Me.Label1.TabIndex = 119
        Me.Label1.Tag = "2002"
        Me.Label1.Text = "ລະຫັດບັນຊີ :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(0, 81)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(85, 21)
        Me.Label2.TabIndex = 120
        Me.Label2.Tag = "2047"
        Me.Label2.Text = "ຊື່ບັນຊີ(ລາວ) :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(-8, 117)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(93, 21)
        Me.Label3.TabIndex = 121
        Me.Label3.Tag = "2048"
        Me.Label3.Text = "ຊື່ບັນຊີ(ອັງກິດ) :"
        '
        'DtmDate
        '
        Me.DtmDate.CustomFormat = "dd/MM/yyyy"
        Me.DtmDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtmDate.Location = New System.Drawing.Point(228, 41)
        Me.DtmDate.Name = "DtmDate"
        Me.DtmDate.Size = New System.Drawing.Size(95, 30)
        Me.DtmDate.TabIndex = 122
        '
        'BtnSave
        '
        Me.BtnSave.Image = CType(resources.GetObject("BtnSave.Image"), System.Drawing.Image)
        Me.BtnSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnSave.Location = New System.Drawing.Point(124, 3)
        Me.BtnSave.Name = "BtnSave"
        Me.BtnSave.Size = New System.Drawing.Size(101, 35)
        Me.BtnSave.TabIndex = 116
        Me.BtnSave.Tag = "3002"
        Me.BtnSave.Text = "ບັນທຶກ"
        Me.BtnSave.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnSave.UseVisualStyleBackColor = True
        '
        'BtnExit
        '
        Me.BtnExit.Image = CType(resources.GetObject("BtnExit.Image"), System.Drawing.Image)
        Me.BtnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnExit.Location = New System.Drawing.Point(4, 3)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(37, 35)
        Me.BtnExit.TabIndex = 118
        Me.BtnExit.Tag = "9999"
        Me.BtnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'BntNew
        '
        Me.BntNew.Image = CType(resources.GetObject("BntNew.Image"), System.Drawing.Image)
        Me.BntNew.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BntNew.Location = New System.Drawing.Point(40, 3)
        Me.BntNew.Name = "BntNew"
        Me.BntNew.Size = New System.Drawing.Size(84, 35)
        Me.BntNew.TabIndex = 117
        Me.BntNew.Tag = "3001"
        Me.BntNew.Text = " ເພີ່ມໃໝ່"
        Me.BntNew.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BntNew.UseVisualStyleBackColor = True
        '
        'CFi
        '
        Me.CFi.AutoSize = True
        Me.CFi.Location = New System.Drawing.Point(536, 44)
        Me.CFi.Name = "CFi"
        Me.CFi.Size = New System.Drawing.Size(57, 25)
        Me.CFi.TabIndex = 123
        Me.CFi.Tag = "5031"
        Me.CFi.Text = "ໜີ້ສິນ"
        Me.CFi.UseVisualStyleBackColor = True
        '
        'CAs
        '
        Me.CAs.AutoSize = True
        Me.CAs.Checked = True
        Me.CAs.Location = New System.Drawing.Point(478, 44)
        Me.CAs.Name = "CAs"
        Me.CAs.Size = New System.Drawing.Size(61, 25)
        Me.CAs.TabIndex = 124
        Me.CAs.TabStop = True
        Me.CAs.Tag = "5030"
        Me.CAs.Text = "ຊັບສິນ"
        Me.CAs.UseVisualStyleBackColor = True
        '
        'CEx
        '
        Me.CEx.AutoSize = True
        Me.CEx.Location = New System.Drawing.Point(401, 44)
        Me.CEx.Name = "CEx"
        Me.CEx.Size = New System.Drawing.Size(75, 25)
        Me.CEx.TabIndex = 125
        Me.CEx.Tag = "5029"
        Me.CEx.Text = "ລາຍຈ່າຍ"
        Me.CEx.UseVisualStyleBackColor = True
        '
        'CIn
        '
        Me.CIn.AutoSize = True
        Me.CIn.Location = New System.Drawing.Point(329, 44)
        Me.CIn.Name = "CIn"
        Me.CIn.Size = New System.Drawing.Size(68, 25)
        Me.CIn.TabIndex = 126
        Me.CIn.Tag = "5028"
        Me.CIn.Text = "ລາຍຮັບ"
        Me.CIn.UseVisualStyleBackColor = True
        '
        'TxtWISE_Orginal
        '
        Me.TxtWISE_Orginal.Location = New System.Drawing.Point(382, 3)
        Me.TxtWISE_Orginal.MaxLength = 20
        Me.TxtWISE_Orginal.Name = "TxtWISE_Orginal"
        Me.TxtWISE_Orginal.Size = New System.Drawing.Size(143, 30)
        Me.TxtWISE_Orginal.TabIndex = 127
        Me.TxtWISE_Orginal.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(279, 6)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(97, 21)
        Me.Label4.TabIndex = 128
        Me.Label4.Tag = "2002"
        Me.Label4.Text = "WISE Orginal:"
        '
        'FrNewAcc
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 21.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(606, 152)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TxtWISE_Orginal)
        Me.Controls.Add(Me.CIn)
        Me.Controls.Add(Me.CEx)
        Me.Controls.Add(Me.CFi)
        Me.Controls.Add(Me.DtmDate)
        Me.Controls.Add(Me.BtnSave)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.BntNew)
        Me.Controls.Add(Me.txtAccName_E)
        Me.Controls.Add(Me.txtAccName)
        Me.Controls.Add(Me.txtAc_code)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.CAs)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "FrNewAcc"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FrNewAccount"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtAc_code As System.Windows.Forms.TextBox
    Friend WithEvents txtAccName As System.Windows.Forms.TextBox
    Friend WithEvents txtAccName_E As System.Windows.Forms.TextBox
    Friend WithEvents BtnSave As System.Windows.Forms.Button
    Friend WithEvents BtnExit As System.Windows.Forms.Button
    Friend WithEvents BntNew As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DtmDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents CFi As System.Windows.Forms.RadioButton
    Friend WithEvents CAs As System.Windows.Forms.RadioButton
    Friend WithEvents CEx As System.Windows.Forms.RadioButton
    Friend WithEvents CIn As System.Windows.Forms.RadioButton
    Friend WithEvents TxtWISE_Orginal As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
End Class
