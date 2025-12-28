<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class fmShartOfAccDetail
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(fmShartOfAccDetail))
        Me.AxVSFlexGrid1 = New AxVSFlex8U.AxVSFlexGrid
        Me.FG = New AxVSFlex8U.AxVSFlexGrid
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.ChAll = New System.Windows.Forms.CheckBox
        Me.Button5 = New System.Windows.Forms.Button
        Me.ChbLang = New System.Windows.Forms.CheckBox
        Me.RdName = New System.Windows.Forms.RadioButton
        Me.RdId = New System.Windows.Forms.RadioButton
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.similar = New System.Windows.Forms.RadioButton
        Me.Rdlasth = New System.Windows.Forms.RadioButton
        Me.txtSearchName = New System.Windows.Forms.TextBox
        Me.txtSearchId = New System.Windows.Forms.TextBox
        Me.Button4 = New System.Windows.Forms.Button
        Me.txtSty = New System.Windows.Forms.TextBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.BntNew = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.BtnExit = New System.Windows.Forms.Button
        Me.Panel6 = New System.Windows.Forms.Panel
        Me.P15 = New System.Windows.Forms.RadioButton
        Me.txtSC15 = New System.Windows.Forms.TextBox
        Me.LbPage = New System.Windows.Forms.Label
        Me.p25 = New System.Windows.Forms.RadioButton
        Me.Button12 = New System.Windows.Forms.Button
        Me.NextPage = New System.Windows.Forms.Button
        Me.p1000 = New System.Windows.Forms.RadioButton
        Me.LasthPage = New System.Windows.Forms.Button
        Me.CmbPage = New System.Windows.Forms.ComboBox
        Me.lblpage_total = New System.Windows.Forms.TextBox
        Me.p500 = New System.Windows.Forms.RadioButton
        Me.FirstPage = New System.Windows.Forms.Button
        Me.p250 = New System.Windows.Forms.RadioButton
        Me.BackPage = New System.Windows.Forms.Button
        Me.p100 = New System.Windows.Forms.RadioButton
        Me.p50 = New System.Windows.Forms.RadioButton
        CType(Me.AxVSFlexGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.SuspendLayout()
        '
        'AxVSFlexGrid1
        '
        Me.AxVSFlexGrid1.DataSource = Nothing
        Me.AxVSFlexGrid1.Location = New System.Drawing.Point(12, 54)
        Me.AxVSFlexGrid1.Name = "AxVSFlexGrid1"
        Me.AxVSFlexGrid1.OcxState = CType(resources.GetObject("AxVSFlexGrid1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxVSFlexGrid1.Size = New System.Drawing.Size(192, 192)
        Me.AxVSFlexGrid1.TabIndex = 0
        '
        'FG
        '
        Me.FG.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FG.DataSource = Nothing
        Me.FG.Location = New System.Drawing.Point(10, 41)
        Me.FG.Name = "FG"
        Me.FG.OcxState = CType(resources.GetObject("FG.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FG.Size = New System.Drawing.Size(851, 412)
        Me.FG.TabIndex = 1
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.ChAll)
        Me.Panel1.Controls.Add(Me.Button5)
        Me.Panel1.Controls.Add(Me.ChbLang)
        Me.Panel1.Controls.Add(Me.RdName)
        Me.Panel1.Controls.Add(Me.RdId)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.txtSearchName)
        Me.Panel1.Controls.Add(Me.txtSearchId)
        Me.Panel1.Controls.Add(Me.Button4)
        Me.Panel1.Location = New System.Drawing.Point(217, 130)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(303, 206)
        Me.Panel1.TabIndex = 130
        Me.Panel1.Visible = False
        '
        'ChAll
        '
        Me.ChAll.AutoSize = True
        Me.ChAll.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.ChAll.Location = New System.Drawing.Point(160, 15)
        Me.ChAll.Name = "ChAll"
        Me.ChAll.Size = New System.Drawing.Size(82, 25)
        Me.ChAll.TabIndex = 8
        Me.ChAll.Text = "Show All"
        Me.ChAll.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(179, 161)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(75, 33)
        Me.Button5.TabIndex = 7
        Me.Button5.Text = "ຍົກເລີກ"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'ChbLang
        '
        Me.ChbLang.AutoSize = True
        Me.ChbLang.Enabled = False
        Me.ChbLang.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.ChbLang.Location = New System.Drawing.Point(216, 80)
        Me.ChbLang.Name = "ChbLang"
        Me.ChbLang.Size = New System.Drawing.Size(76, 25)
        Me.ChbLang.TabIndex = 6
        Me.ChbLang.Text = "Englisth"
        Me.ChbLang.UseVisualStyleBackColor = True
        '
        'RdName
        '
        Me.RdName.AutoSize = True
        Me.RdName.Location = New System.Drawing.Point(113, 78)
        Me.RdName.Name = "RdName"
        Me.RdName.Size = New System.Drawing.Size(83, 25)
        Me.RdName.TabIndex = 5
        Me.RdName.Text = "ຕາມຊື່ບັນຊີ"
        Me.RdName.UseVisualStyleBackColor = True
        '
        'RdId
        '
        Me.RdId.AutoSize = True
        Me.RdId.Checked = True
        Me.RdId.Location = New System.Drawing.Point(16, 78)
        Me.RdId.Name = "RdId"
        Me.RdId.Size = New System.Drawing.Size(85, 25)
        Me.RdId.TabIndex = 4
        Me.RdId.TabStop = True
        Me.RdId.Text = "ຕາມລະຫັດ"
        Me.RdId.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.similar)
        Me.GroupBox1.Controls.Add(Me.Rdlasth)
        Me.GroupBox1.Location = New System.Drawing.Point(10, 101)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(282, 54)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "ຮູບແບບການຄົ້ນຫາ"
        '
        'similar
        '
        Me.similar.AutoSize = True
        Me.similar.Location = New System.Drawing.Point(168, 21)
        Me.similar.Name = "similar"
        Me.similar.Size = New System.Drawing.Size(84, 25)
        Me.similar.TabIndex = 1
        Me.similar.Text = "ຄຳທີບັນຈຸຢູ່"
        Me.similar.UseVisualStyleBackColor = True
        '
        'Rdlasth
        '
        Me.Rdlasth.AutoSize = True
        Me.Rdlasth.Checked = True
        Me.Rdlasth.Location = New System.Drawing.Point(6, 21)
        Me.Rdlasth.Name = "Rdlasth"
        Me.Rdlasth.Size = New System.Drawing.Size(85, 25)
        Me.Rdlasth.TabIndex = 0
        Me.Rdlasth.TabStop = True
        Me.Rdlasth.Text = "ຄຳທີ່ຂື້ນຕົນ"
        Me.Rdlasth.UseVisualStyleBackColor = True
        '
        'txtSearchName
        '
        Me.txtSearchName.Enabled = False
        Me.txtSearchName.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSearchName.Location = New System.Drawing.Point(10, 44)
        Me.txtSearchName.Name = "txtSearchName"
        Me.txtSearchName.Size = New System.Drawing.Size(282, 30)
        Me.txtSearchName.TabIndex = 2
        '
        'txtSearchId
        '
        Me.txtSearchId.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSearchId.Location = New System.Drawing.Point(10, 10)
        Me.txtSearchId.Name = "txtSearchId"
        Me.txtSearchId.Size = New System.Drawing.Size(100, 30)
        Me.txtSearchId.TabIndex = 1
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(9, 161)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(147, 33)
        Me.Button4.TabIndex = 0
        Me.Button4.Text = "ຊອກຫາ"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'txtSty
        '
        Me.txtSty.Location = New System.Drawing.Point(454, 8)
        Me.txtSty.Name = "txtSty"
        Me.txtSty.Size = New System.Drawing.Size(274, 30)
        Me.txtSty.TabIndex = 132
        '
        'Button2
        '
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(317, 6)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(122, 35)
        Me.Button2.TabIndex = 133
        Me.Button2.Text = "ຊອກຫາທັງໝົດ"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Image = CType(resources.GetObject("Button6.Image"), System.Drawing.Image)
        Me.Button6.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button6.Location = New System.Drawing.Point(151, 6)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(84, 35)
        Me.Button6.TabIndex = 128
        Me.Button6.Text = "ຕົກລົງ"
        Me.Button6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button6.UseVisualStyleBackColor = True
        '
        'BntNew
        '
        Me.BntNew.Image = CType(resources.GetObject("BntNew.Image"), System.Drawing.Image)
        Me.BntNew.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BntNew.Location = New System.Drawing.Point(42, 6)
        Me.BntNew.Name = "BntNew"
        Me.BntNew.Size = New System.Drawing.Size(110, 35)
        Me.BntNew.TabIndex = 129
        Me.BntNew.Text = "ເປີດບັນຊີໃໝ່"
        Me.BntNew.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BntNew.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(234, 6)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(84, 35)
        Me.Button1.TabIndex = 127
        Me.Button1.Text = "ຊອກຫາ"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Timer1
        '
        '
        'BtnExit
        '
        Me.BtnExit.Image = Global.ApPBank10.My.Resources.Resources.Exit1
        Me.BtnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnExit.Location = New System.Drawing.Point(10, 5)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(35, 35)
        Me.BtnExit.TabIndex = 131
        Me.BtnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'Panel6
        '
        Me.Panel6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel6.Controls.Add(Me.P15)
        Me.Panel6.Controls.Add(Me.txtSC15)
        Me.Panel6.Controls.Add(Me.LbPage)
        Me.Panel6.Controls.Add(Me.p25)
        Me.Panel6.Controls.Add(Me.Button12)
        Me.Panel6.Controls.Add(Me.NextPage)
        Me.Panel6.Controls.Add(Me.p1000)
        Me.Panel6.Controls.Add(Me.LasthPage)
        Me.Panel6.Controls.Add(Me.CmbPage)
        Me.Panel6.Controls.Add(Me.lblpage_total)
        Me.Panel6.Controls.Add(Me.p500)
        Me.Panel6.Controls.Add(Me.FirstPage)
        Me.Panel6.Controls.Add(Me.p250)
        Me.Panel6.Controls.Add(Me.BackPage)
        Me.Panel6.Controls.Add(Me.p100)
        Me.Panel6.Controls.Add(Me.p50)
        Me.Panel6.Location = New System.Drawing.Point(10, 466)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(1138, 50)
        Me.Panel6.TabIndex = 45561
        '
        'P15
        '
        Me.P15.AutoSize = True
        Me.P15.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.P15.Location = New System.Drawing.Point(3, 26)
        Me.P15.Name = "P15"
        Me.P15.Size = New System.Drawing.Size(80, 19)
        Me.P15.TabIndex = 219
        Me.P15.Text = "Customize"
        Me.P15.UseVisualStyleBackColor = True
        '
        'txtSC15
        '
        Me.txtSC15.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSC15.ForeColor = System.Drawing.Color.Blue
        Me.txtSC15.Location = New System.Drawing.Point(365, 2)
        Me.txtSC15.MaxLength = 4
        Me.txtSC15.Name = "txtSC15"
        Me.txtSC15.Size = New System.Drawing.Size(35, 20)
        Me.txtSC15.TabIndex = 220
        Me.txtSC15.Text = "16"
        Me.txtSC15.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'LbPage
        '
        Me.LbPage.AutoSize = True
        Me.LbPage.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LbPage.ForeColor = System.Drawing.Color.Blue
        Me.LbPage.Location = New System.Drawing.Point(388, 28)
        Me.LbPage.Name = "LbPage"
        Me.LbPage.Size = New System.Drawing.Size(73, 15)
        Me.LbPage.TabIndex = 178
        Me.LbPage.Text = "RecordTotal"
        '
        'p25
        '
        Me.p25.AutoSize = True
        Me.p25.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.p25.Location = New System.Drawing.Point(91, 26)
        Me.p25.Name = "p25"
        Me.p25.Size = New System.Drawing.Size(37, 19)
        Me.p25.TabIndex = 216
        Me.p25.Text = "25"
        Me.p25.UseVisualStyleBackColor = True
        '
        'Button12
        '
        Me.Button12.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button12.ForeColor = System.Drawing.Color.Blue
        Me.Button12.Location = New System.Drawing.Point(1, 0)
        Me.Button12.Name = "Button12"
        Me.Button12.Size = New System.Drawing.Size(78, 23)
        Me.Button12.TabIndex = 218
        Me.Button12.Text = "Start New"
        Me.Button12.UseVisualStyleBackColor = True
        '
        'NextPage
        '
        Me.NextPage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NextPage.ForeColor = System.Drawing.Color.Blue
        Me.NextPage.Location = New System.Drawing.Point(290, 1)
        Me.NextPage.Name = "NextPage"
        Me.NextPage.Size = New System.Drawing.Size(37, 23)
        Me.NextPage.TabIndex = 207
        Me.NextPage.Text = ">>"
        Me.NextPage.UseVisualStyleBackColor = True
        '
        'p1000
        '
        Me.p1000.AutoSize = True
        Me.p1000.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.p1000.Location = New System.Drawing.Point(329, 26)
        Me.p1000.Name = "p1000"
        Me.p1000.Size = New System.Drawing.Size(52, 19)
        Me.p1000.TabIndex = 215
        Me.p1000.Text = "1,000"
        Me.p1000.UseVisualStyleBackColor = True
        '
        'LasthPage
        '
        Me.LasthPage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.LasthPage.ForeColor = System.Drawing.Color.Blue
        Me.LasthPage.Location = New System.Drawing.Point(326, 1)
        Me.LasthPage.Name = "LasthPage"
        Me.LasthPage.Size = New System.Drawing.Size(38, 23)
        Me.LasthPage.TabIndex = 208
        Me.LasthPage.Text = ">>|"
        Me.LasthPage.UseVisualStyleBackColor = True
        '
        'CmbPage
        '
        Me.CmbPage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbPage.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbPage.FormattingEnabled = True
        Me.CmbPage.Location = New System.Drawing.Point(79, 1)
        Me.CmbPage.Name = "CmbPage"
        Me.CmbPage.Size = New System.Drawing.Size(60, 22)
        Me.CmbPage.TabIndex = 217
        '
        'lblpage_total
        '
        Me.lblpage_total.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblpage_total.ForeColor = System.Drawing.Color.Blue
        Me.lblpage_total.Location = New System.Drawing.Point(215, 2)
        Me.lblpage_total.Name = "lblpage_total"
        Me.lblpage_total.ReadOnly = True
        Me.lblpage_total.Size = New System.Drawing.Size(75, 20)
        Me.lblpage_total.TabIndex = 206
        Me.lblpage_total.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'p500
        '
        Me.p500.AutoSize = True
        Me.p500.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.p500.Location = New System.Drawing.Point(279, 26)
        Me.p500.Name = "p500"
        Me.p500.Size = New System.Drawing.Size(43, 19)
        Me.p500.TabIndex = 214
        Me.p500.Text = "500"
        Me.p500.UseVisualStyleBackColor = True
        '
        'FirstPage
        '
        Me.FirstPage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.FirstPage.ForeColor = System.Drawing.Color.Blue
        Me.FirstPage.Location = New System.Drawing.Point(139, 1)
        Me.FirstPage.Name = "FirstPage"
        Me.FirstPage.Size = New System.Drawing.Size(39, 23)
        Me.FirstPage.TabIndex = 209
        Me.FirstPage.Text = "|<<"
        Me.FirstPage.UseVisualStyleBackColor = True
        '
        'p250
        '
        Me.p250.AutoSize = True
        Me.p250.Checked = True
        Me.p250.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.p250.Location = New System.Drawing.Point(229, 26)
        Me.p250.Name = "p250"
        Me.p250.Size = New System.Drawing.Size(43, 19)
        Me.p250.TabIndex = 213
        Me.p250.TabStop = True
        Me.p250.Text = "250"
        Me.p250.UseVisualStyleBackColor = True
        '
        'BackPage
        '
        Me.BackPage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.BackPage.ForeColor = System.Drawing.Color.Blue
        Me.BackPage.Location = New System.Drawing.Point(177, 1)
        Me.BackPage.Name = "BackPage"
        Me.BackPage.Size = New System.Drawing.Size(37, 23)
        Me.BackPage.TabIndex = 210
        Me.BackPage.Text = "<<"
        Me.BackPage.UseVisualStyleBackColor = True
        '
        'p100
        '
        Me.p100.AutoSize = True
        Me.p100.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.p100.Location = New System.Drawing.Point(179, 26)
        Me.p100.Name = "p100"
        Me.p100.Size = New System.Drawing.Size(43, 19)
        Me.p100.TabIndex = 212
        Me.p100.Text = "100"
        Me.p100.UseVisualStyleBackColor = True
        '
        'p50
        '
        Me.p50.AutoSize = True
        Me.p50.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.p50.Location = New System.Drawing.Point(135, 26)
        Me.p50.Name = "p50"
        Me.p50.Size = New System.Drawing.Size(37, 19)
        Me.p50.TabIndex = 211
        Me.p50.Text = "50"
        Me.p50.UseVisualStyleBackColor = True
        '
        'fmShartOfAccDetail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 21.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(873, 526)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel6)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.txtSty)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.BntNew)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.FG)
        Me.Controls.Add(Me.AxVSFlexGrid1)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "fmShartOfAccDetail"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "fmShartOfAccDetail"
        CType(Me.AxVSFlexGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents AxVSFlexGrid1 As AxVSFlex8U.AxVSFlexGrid
    Friend WithEvents FG As AxVSFlex8U.AxVSFlexGrid
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents BntNew As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents ChbLang As System.Windows.Forms.CheckBox
    Friend WithEvents RdName As System.Windows.Forms.RadioButton
    Friend WithEvents RdId As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents similar As System.Windows.Forms.RadioButton
    Friend WithEvents Rdlasth As System.Windows.Forms.RadioButton
    Friend WithEvents txtSearchName As System.Windows.Forms.TextBox
    Friend WithEvents txtSearchId As System.Windows.Forms.TextBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents txtSty As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents BtnExit As System.Windows.Forms.Button
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents P15 As System.Windows.Forms.RadioButton
    Friend WithEvents txtSC15 As System.Windows.Forms.TextBox
    Friend WithEvents LbPage As System.Windows.Forms.Label
    Friend WithEvents p25 As System.Windows.Forms.RadioButton
    Friend WithEvents Button12 As System.Windows.Forms.Button
    Friend WithEvents NextPage As System.Windows.Forms.Button
    Friend WithEvents p1000 As System.Windows.Forms.RadioButton
    Friend WithEvents LasthPage As System.Windows.Forms.Button
    Friend WithEvents CmbPage As System.Windows.Forms.ComboBox
    Friend WithEvents lblpage_total As System.Windows.Forms.TextBox
    Friend WithEvents p500 As System.Windows.Forms.RadioButton
    Friend WithEvents FirstPage As System.Windows.Forms.Button
    Friend WithEvents p250 As System.Windows.Forms.RadioButton
    Friend WithEvents BackPage As System.Windows.Forms.Button
    Friend WithEvents p100 As System.Windows.Forms.RadioButton
    Friend WithEvents p50 As System.Windows.Forms.RadioButton
    Friend WithEvents ChAll As System.Windows.Forms.CheckBox
End Class
