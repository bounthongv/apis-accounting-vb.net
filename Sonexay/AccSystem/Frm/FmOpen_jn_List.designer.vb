<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FmOpen_jn_List
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FmOpen_jn_List))
        Me.Button1 = New System.Windows.Forms.Button
        Me.BtnDelete = New System.Windows.Forms.Button
        Me.BtnEdit = New System.Windows.Forms.Button
        Me.BtnExit = New System.Windows.Forms.Button
        Me.BntNew = New System.Windows.Forms.Button
        Me.FG = New System.Windows.Forms.DataGridView
        Me.BtnRefresh = New System.Windows.Forms.Button
        Me.BtnPreview = New System.Windows.Forms.Button
        Me.txtSumAmountCr = New System.Windows.Forms.TextBox
        Me.txtSumAmountDr = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label1 = New System.Windows.Forms.Label
        Me.BalanceDr = New System.Windows.Forms.TextBox
        Me.ChAllSty = New System.Windows.Forms.CheckBox
        Me.txtUserId = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.yy = New System.Windows.Forms.DateTimePicker
        Me.Ds = New System.Windows.Forms.DateTimePicker
        Me.Off_Usr = New System.Windows.Forms.ComboBox
        Me.Label27 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        CType(Me.FG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(298, 9)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(84, 34)
        Me.Button1.TabIndex = 168
        Me.Button1.Tag = "00"
        Me.Button1.Text = "ລ໋ອກປະຈຳປີ"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'BtnDelete
        '
        Me.BtnDelete.Image = CType(resources.GetObject("BtnDelete.Image"), System.Drawing.Image)
        Me.BtnDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnDelete.Location = New System.Drawing.Point(214, 8)
        Me.BtnDelete.Name = "BtnDelete"
        Me.BtnDelete.Size = New System.Drawing.Size(84, 35)
        Me.BtnDelete.TabIndex = 167
        Me.BtnDelete.Tag = "3004"
        Me.BtnDelete.Text = "    ລືບ"
        Me.BtnDelete.UseVisualStyleBackColor = True
        '
        'BtnEdit
        '
        Me.BtnEdit.Image = CType(resources.GetObject("BtnEdit.Image"), System.Drawing.Image)
        Me.BtnEdit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnEdit.Location = New System.Drawing.Point(130, 8)
        Me.BtnEdit.Name = "BtnEdit"
        Me.BtnEdit.Size = New System.Drawing.Size(84, 35)
        Me.BtnEdit.TabIndex = 166
        Me.BtnEdit.Tag = "3003"
        Me.BtnEdit.Text = "ແກ້ໄຂ"
        Me.BtnEdit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnEdit.UseVisualStyleBackColor = True
        '
        'BtnExit
        '
        Me.BtnExit.Image = CType(resources.GetObject("BtnExit.Image"), System.Drawing.Image)
        Me.BtnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnExit.Location = New System.Drawing.Point(11, 7)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(35, 35)
        Me.BtnExit.TabIndex = 164
        Me.BtnExit.Tag = "9999"
        Me.BtnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'BntNew
        '
        Me.BntNew.Image = CType(resources.GetObject("BntNew.Image"), System.Drawing.Image)
        Me.BntNew.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BntNew.Location = New System.Drawing.Point(45, 8)
        Me.BntNew.Name = "BntNew"
        Me.BntNew.Size = New System.Drawing.Size(84, 35)
        Me.BntNew.TabIndex = 162
        Me.BntNew.Tag = "3001"
        Me.BntNew.Text = " ເພີ່ມໃໝ່"
        Me.BntNew.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BntNew.UseVisualStyleBackColor = True
        '
        'FG
        '
        Me.FG.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FG.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.FG.BackgroundColor = System.Drawing.Color.White
        Me.FG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.FG.Location = New System.Drawing.Point(12, 47)
        Me.FG.Name = "FG"
        Me.FG.Size = New System.Drawing.Size(1075, 527)
        Me.FG.TabIndex = 169
        Me.FG.Tag = "8004"
        '
        'BtnRefresh
        '
        Me.BtnRefresh.Image = CType(resources.GetObject("BtnRefresh.Image"), System.Drawing.Image)
        Me.BtnRefresh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnRefresh.Location = New System.Drawing.Point(466, 9)
        Me.BtnRefresh.Name = "BtnRefresh"
        Me.BtnRefresh.Size = New System.Drawing.Size(100, 35)
        Me.BtnRefresh.TabIndex = 45539
        Me.BtnRefresh.Tag = "3005"
        Me.BtnRefresh.Text = "ເອີ້ນຂໍ້ມູນ"
        Me.BtnRefresh.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnRefresh.UseVisualStyleBackColor = True
        '
        'BtnPreview
        '
        Me.BtnPreview.Image = CType(resources.GetObject("BtnPreview.Image"), System.Drawing.Image)
        Me.BtnPreview.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnPreview.Location = New System.Drawing.Point(566, 9)
        Me.BtnPreview.Name = "BtnPreview"
        Me.BtnPreview.Size = New System.Drawing.Size(100, 35)
        Me.BtnPreview.TabIndex = 45541
        Me.BtnPreview.Tag = "3006"
        Me.BtnPreview.Text = "ວິວ/ເບິ່ງ"
        Me.BtnPreview.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnPreview.UseVisualStyleBackColor = True
        '
        'txtSumAmountCr
        '
        Me.txtSumAmountCr.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtSumAmountCr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSumAmountCr.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtSumAmountCr.Location = New System.Drawing.Point(576, 9)
        Me.txtSumAmountCr.Name = "txtSumAmountCr"
        Me.txtSumAmountCr.ReadOnly = True
        Me.txtSumAmountCr.Size = New System.Drawing.Size(160, 30)
        Me.txtSumAmountCr.TabIndex = 45542
        Me.txtSumAmountCr.Text = "0.00"
        Me.txtSumAmountCr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtSumAmountDr
        '
        Me.txtSumAmountDr.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtSumAmountDr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSumAmountDr.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtSumAmountDr.Location = New System.Drawing.Point(305, 8)
        Me.txtSumAmountDr.Name = "txtSumAmountDr"
        Me.txtSumAmountDr.ReadOnly = True
        Me.txtSumAmountDr.Size = New System.Drawing.Size(160, 30)
        Me.txtSumAmountDr.TabIndex = 45544
        Me.txtSumAmountDr.Text = "0.00"
        Me.txtSumAmountDr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label11.Location = New System.Drawing.Point(223, 13)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(79, 21)
        Me.Label11.TabIndex = 45547
        Me.Label11.Tag = "2008"
        Me.Label11.Text = "ລວມຈົດຫນື້ :"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label12.Location = New System.Drawing.Point(504, 12)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(69, 21)
        Me.Label12.TabIndex = 45546
        Me.Label12.Tag = "2009"
        Me.Label12.Text = "ລວມຈົດມີ :"
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.BalanceDr)
        Me.Panel1.Controls.Add(Me.txtSumAmountDr)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.txtSumAmountCr)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Location = New System.Drawing.Point(11, 591)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1007, 44)
        Me.Panel1.TabIndex = 45550
        Me.Panel1.Tag = "1"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(762, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 21)
        Me.Label1.TabIndex = 45551
        Me.Label1.Tag = "2010"
        Me.Label1.Text = "ຄ່າຜິດດ່ຽງ :"
        '
        'BalanceDr
        '
        Me.BalanceDr.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.BalanceDr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BalanceDr.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.BalanceDr.Location = New System.Drawing.Point(840, 9)
        Me.BalanceDr.Name = "BalanceDr"
        Me.BalanceDr.ReadOnly = True
        Me.BalanceDr.Size = New System.Drawing.Size(160, 30)
        Me.BalanceDr.TabIndex = 45548
        Me.BalanceDr.Text = "0.00"
        Me.BalanceDr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'ChAllSty
        '
        Me.ChAllSty.AutoSize = True
        Me.ChAllSty.Location = New System.Drawing.Point(789, 12)
        Me.ChAllSty.Name = "ChAllSty"
        Me.ChAllSty.Size = New System.Drawing.Size(93, 25)
        Me.ChAllSty.TabIndex = 45551
        Me.ChAllSty.Tag = "4010"
        Me.ChAllSty.Text = "ແບບສັງລວມ"
        Me.ChAllSty.UseVisualStyleBackColor = True
        '
        'txtUserId
        '
        Me.txtUserId.Font = New System.Drawing.Font("Saysettha OT", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUserId.Location = New System.Drawing.Point(957, -25)
        Me.txtUserId.Name = "txtUserId"
        Me.txtUserId.Size = New System.Drawing.Size(61, 28)
        Me.txtUserId.TabIndex = 45552
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(1094, 15)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 21)
        Me.Label2.TabIndex = 45556
        Me.Label2.Tag = "2049"
        Me.Label2.Text = "ປະຈຳປີ"
        '
        'yy
        '
        Me.yy.CustomFormat = "yyyy"
        Me.yy.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.yy.Location = New System.Drawing.Point(1148, 9)
        Me.yy.Name = "yy"
        Me.yy.Size = New System.Drawing.Size(85, 30)
        Me.yy.TabIndex = 45557
        '
        'Ds
        '
        Me.Ds.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.Ds.Location = New System.Drawing.Point(628, -32)
        Me.Ds.Name = "Ds"
        Me.Ds.Size = New System.Drawing.Size(94, 30)
        Me.Ds.TabIndex = 45558
        '
        'Off_Usr
        '
        Me.Off_Usr.Font = New System.Drawing.Font("Saysettha OT", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Off_Usr.FormattingEnabled = True
        Me.Off_Usr.Location = New System.Drawing.Point(944, 12)
        Me.Off_Usr.Name = "Off_Usr"
        Me.Off_Usr.Size = New System.Drawing.Size(143, 26)
        Me.Off_Usr.TabIndex = 45613
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(878, 15)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(60, 21)
        Me.Label27.TabIndex = 45612
        Me.Label27.Tag = "2011"
        Me.Label27.Text = "ໜ່ວຍງານ"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(695, -38)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(100, 30)
        Me.TextBox1.TabIndex = 45614
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(382, 9)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(84, 35)
        Me.Button2.TabIndex = 45615
        Me.Button2.Tag = "00"
        Me.Button2.Text = "ປົດລ໋ອກປະຈຳປີ"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(672, 13)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(61, 25)
        Me.CheckBox1.TabIndex = 45616
        Me.CheckBox1.Tag = "4010"
        Me.CheckBox1.Text = "ພາສາ"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'FmOpen_jn_List
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 21.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1157, 648)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Off_Usr)
        Me.Controls.Add(Me.Label27)
        Me.Controls.Add(Me.Ds)
        Me.Controls.Add(Me.yy)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtUserId)
        Me.Controls.Add(Me.ChAllSty)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.BtnPreview)
        Me.Controls.Add(Me.BtnRefresh)
        Me.Controls.Add(Me.FG)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.BtnDelete)
        Me.Controls.Add(Me.BtnEdit)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.BntNew)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "FmOpen_jn_List"
        Me.Padding = New System.Windows.Forms.Padding(12, 15, 12, 15)
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "FmOpen_jn_List"
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents BtnDelete As System.Windows.Forms.Button
    Friend WithEvents BtnEdit As System.Windows.Forms.Button
    Friend WithEvents BtnExit As System.Windows.Forms.Button
    Friend WithEvents BntNew As System.Windows.Forms.Button
    Friend WithEvents FG As System.Windows.Forms.DataGridView
    Friend WithEvents BtnRefresh As System.Windows.Forms.Button
    Friend WithEvents BtnPreview As System.Windows.Forms.Button
    Friend WithEvents txtSumAmountCr As System.Windows.Forms.TextBox
    Friend WithEvents txtSumAmountDr As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents BalanceDr As System.Windows.Forms.TextBox
    Friend WithEvents ChAllSty As System.Windows.Forms.CheckBox
    Friend WithEvents txtUserId As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents yy As System.Windows.Forms.DateTimePicker
    Friend WithEvents Ds As System.Windows.Forms.DateTimePicker
    Friend WithEvents Off_Usr As System.Windows.Forms.ComboBox
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox

End Class
