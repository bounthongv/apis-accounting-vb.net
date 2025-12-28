<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FmReceipt_List
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FmReceipt_List))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BtnExit = New System.Windows.Forms.Button
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Ac_Bnk_Coode = New System.Windows.Forms.RadioButton
        Me.TAc_Bnk_Coode = New System.Windows.Forms.TextBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.RRight = New System.Windows.Forms.RadioButton
        Me.RPercent = New System.Windows.Forms.RadioButton
        Me.RLeft = New System.Windows.Forms.RadioButton
        Me.Rfull = New System.Windows.Forms.RadioButton
        Me.Button3 = New System.Windows.Forms.Button
        Me.SearchId = New System.Windows.Forms.TextBox
        Me.RdName = New System.Windows.Forms.RadioButton
        Me.RdDate = New System.Windows.Forms.RadioButton
        Me.RdId = New System.Windows.Forms.RadioButton
        Me.DtmStartDate = New System.Windows.Forms.DateTimePicker
        Me.DtmToDate = New System.Windows.Forms.DateTimePicker
        Me.Button1 = New System.Windows.Forms.Button
        Me.SearchName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.FG = New AxVSFlex8U.AxVSFlexGrid
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.FG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.BtnExit)
        Me.GroupBox1.Controls.Add(Me.ComboBox1)
        Me.GroupBox1.Controls.Add(Me.Ac_Bnk_Coode)
        Me.GroupBox1.Controls.Add(Me.TAc_Bnk_Coode)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.Button3)
        Me.GroupBox1.Controls.Add(Me.SearchId)
        Me.GroupBox1.Controls.Add(Me.RdName)
        Me.GroupBox1.Controls.Add(Me.RdDate)
        Me.GroupBox1.Controls.Add(Me.RdId)
        Me.GroupBox1.Controls.Add(Me.DtmStartDate)
        Me.GroupBox1.Controls.Add(Me.DtmToDate)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.SearchName)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Location = New System.Drawing.Point(7, -7)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(450, 219)
        Me.GroupBox1.TabIndex = 167
        Me.GroupBox1.TabStop = False
        '
        'BtnExit
        '

        Me.BtnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnExit.Location = New System.Drawing.Point(3, 14)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(35, 35)
        Me.BtnExit.TabIndex = 292
        Me.BtnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'ComboBox1
        '
        Me.ComboBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"ການຮັບເງິນຈາກການເປີດບັນຊີເງິນຝາກໃຫມ່", "ການຮັບເງິນຈາກການມອບເງິນເຂົ້າບັນຊີ", "ການຈ່າຍຈາກການຖອນເງິນ", "ການຮັບເງິນຈາກການຝາກເງິນດ່ອນ", "ການຈ່າຍເງິນໃຫ້ລູກຄ້າຈາກການໂອນ", "ການຈ່າຍເງິນຈາກການປ່ອຍກູ້", "ການຮັບເງິນຈາກການຊຳລະຕົ້ນທືນເງິນປ່ອຍກກູ້", "ການຮັບເງິນຈາກການຊຳລະດອກເບ້ຍເງິນປ່ອຍກກູ້", "ການຮັບເງິນຈາກການປັບໄຫມ", "ການຈ່າຍເງິນໃຫ້ຄັງຍ່ອຍ", "ການຮັບເງິນຈາກຄັງຍ່ອຍ", "ການຮັບເງິນຈາກການແລກປ່ຽນເງິນຕາ", "ການຈ່າຍເງິນຈາກການແລກປ່ຽນເງິນຕາ"})
        Me.ComboBox1.Location = New System.Drawing.Point(33, 181)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(408, 29)
        Me.ComboBox1.TabIndex = 274
        Me.ComboBox1.Tag = "dfgdfg"
        '
        'Ac_Bnk_Coode
        '
        Me.Ac_Bnk_Coode.AutoSize = True
        Me.Ac_Bnk_Coode.Location = New System.Drawing.Point(19, 119)
        Me.Ac_Bnk_Coode.Name = "Ac_Bnk_Coode"
        Me.Ac_Bnk_Coode.Size = New System.Drawing.Size(125, 25)
        Me.Ac_Bnk_Coode.TabIndex = 176
        Me.Ac_Bnk_Coode.Text = "ຕາມເລກບັນຊີເງິນກູ້"
        Me.Ac_Bnk_Coode.UseVisualStyleBackColor = True
        '
        'TAc_Bnk_Coode
        '
        Me.TAc_Bnk_Coode.Location = New System.Drawing.Point(148, 117)
        Me.TAc_Bnk_Coode.Name = "TAc_Bnk_Coode"
        Me.TAc_Bnk_Coode.Size = New System.Drawing.Size(220, 30)
        Me.TAc_Bnk_Coode.TabIndex = 175
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RRight)
        Me.GroupBox2.Controls.Add(Me.RPercent)
        Me.GroupBox2.Controls.Add(Me.RLeft)
        Me.GroupBox2.Controls.Add(Me.Rfull)
        Me.GroupBox2.Location = New System.Drawing.Point(369, 16)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(75, 159)
        Me.GroupBox2.TabIndex = 174
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "ໃສ່ຂໍ້ມູນ"
        '
        'RRight
        '
        Me.RRight.AutoSize = True
        Me.RRight.Location = New System.Drawing.Point(6, 91)
        Me.RRight.Name = "RRight"
        Me.RRight.Size = New System.Drawing.Size(68, 25)
        Me.RRight.TabIndex = 3
        Me.RRight.TabStop = True
        Me.RRight.Text = "ລົງທ້າຍ"
        Me.RRight.UseVisualStyleBackColor = True
        '
        'RPercent
        '
        Me.RPercent.AutoSize = True
        Me.RPercent.Location = New System.Drawing.Point(6, 125)
        Me.RPercent.Name = "RPercent"
        Me.RPercent.Size = New System.Drawing.Size(67, 25)
        Me.RPercent.TabIndex = 2
        Me.RPercent.TabStop = True
        Me.RPercent.Text = "ຄຳບັນຈຸ"
        Me.RPercent.UseVisualStyleBackColor = True
        '
        'RLeft
        '
        Me.RLeft.AutoSize = True
        Me.RLeft.Location = New System.Drawing.Point(6, 59)
        Me.RLeft.Name = "RLeft"
        Me.RLeft.Size = New System.Drawing.Size(61, 25)
        Me.RLeft.TabIndex = 1
        Me.RLeft.TabStop = True
        Me.RLeft.Text = "ຂື້ນຕົ້ນ"
        Me.RLeft.UseVisualStyleBackColor = True
        '
        'Rfull
        '
        Me.Rfull.AutoSize = True
        Me.Rfull.Location = New System.Drawing.Point(7, 26)
        Me.Rfull.Name = "Rfull"
        Me.Rfull.Size = New System.Drawing.Size(49, 25)
        Me.Rfull.TabIndex = 0
        Me.Rfull.TabStop = True
        Me.Rfull.Text = "ເຕັມ"
        Me.Rfull.UseVisualStyleBackColor = True
        '
        'Button3
        '

        Me.Button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button3.Location = New System.Drawing.Point(37, 14)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(95, 35)
        Me.Button3.TabIndex = 171
        Me.Button3.Text = "ເພີ່ມໃຫມ່"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'SearchId
        '
        Me.SearchId.Location = New System.Drawing.Point(148, 52)
        Me.SearchId.Name = "SearchId"
        Me.SearchId.Size = New System.Drawing.Size(95, 30)
        Me.SearchId.TabIndex = 168
        '
        'RdName
        '
        Me.RdName.AutoSize = True
        Me.RdName.Location = New System.Drawing.Point(20, 86)
        Me.RdName.Name = "RdName"
        Me.RdName.Size = New System.Drawing.Size(83, 25)
        Me.RdName.TabIndex = 169
        Me.RdName.Text = "ຕາມລາຍຊື່"
        Me.RdName.UseVisualStyleBackColor = True
        '
        'RdDate
        '
        Me.RdDate.AutoSize = True
        Me.RdDate.Location = New System.Drawing.Point(19, 153)
        Me.RdDate.Name = "RdDate"
        Me.RdDate.Size = New System.Drawing.Size(80, 25)
        Me.RdDate.TabIndex = 169
        Me.RdDate.Text = "ຕາມເວລາ"
        Me.RdDate.UseVisualStyleBackColor = True
        '
        'RdId
        '
        Me.RdId.AutoSize = True
        Me.RdId.Checked = True
        Me.RdId.Location = New System.Drawing.Point(20, 53)
        Me.RdId.Name = "RdId"
        Me.RdId.Size = New System.Drawing.Size(85, 25)
        Me.RdId.TabIndex = 169
        Me.RdId.TabStop = True
        Me.RdId.Text = "ຕາມລະຫັດ"
        Me.RdId.UseVisualStyleBackColor = True
        '
        'DtmStartDate
        '
        Me.DtmStartDate.CustomFormat = "dd/MM/yyyy"
        Me.DtmStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtmStartDate.Location = New System.Drawing.Point(148, 151)
        Me.DtmStartDate.Name = "DtmStartDate"
        Me.DtmStartDate.Size = New System.Drawing.Size(95, 30)
        Me.DtmStartDate.TabIndex = 165
        '
        'DtmToDate
        '
        Me.DtmToDate.CustomFormat = "dd/MM/yyyy"
        Me.DtmToDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtmToDate.Location = New System.Drawing.Point(271, 151)
        Me.DtmToDate.Name = "DtmToDate"
        Me.DtmToDate.Size = New System.Drawing.Size(97, 30)
        Me.DtmToDate.TabIndex = 164
        '
        'Button1
        '

        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(133, 14)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(117, 35)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = " ຄົ້ນຫາ"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'SearchName
        '
        Me.SearchName.Location = New System.Drawing.Point(148, 84)
        Me.SearchName.Name = "SearchName"
        Me.SearchName.Size = New System.Drawing.Size(220, 30)
        Me.SearchName.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(247, 157)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(41, 21)
        Me.Label3.TabIndex = 163
        Me.Label3.Text = "ເຖິງ :"
        '
        'FG
        '
        Me.FG.DataSource = Nothing
        Me.FG.Location = New System.Drawing.Point(7, 218)
        Me.FG.Name = "FG"
        Me.FG.OcxState = CType(resources.GetObject("FG.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FG.Size = New System.Drawing.Size(450, 502)
        Me.FG.TabIndex = 168
        '
        'FmReceipt_List
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 21.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(461, 636)
        Me.Controls.Add(Me.FG)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FmReceipt_List"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FmReceipt_List"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents SearchId As System.Windows.Forms.TextBox
    Friend WithEvents RdName As System.Windows.Forms.RadioButton
    Friend WithEvents RdDate As System.Windows.Forms.RadioButton
    Friend WithEvents RdId As System.Windows.Forms.RadioButton
    Friend WithEvents DtmStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents DtmToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents SearchName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents FG As AxVSFlex8U.AxVSFlexGrid
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents RRight As System.Windows.Forms.RadioButton
    Friend WithEvents RPercent As System.Windows.Forms.RadioButton
    Friend WithEvents RLeft As System.Windows.Forms.RadioButton
    Friend WithEvents Rfull As System.Windows.Forms.RadioButton
    Friend WithEvents Ac_Bnk_Coode As System.Windows.Forms.RadioButton
    Friend WithEvents TAc_Bnk_Coode As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents BtnExit As System.Windows.Forms.Button
End Class
