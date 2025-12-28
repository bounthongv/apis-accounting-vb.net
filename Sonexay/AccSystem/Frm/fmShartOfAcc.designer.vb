<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class fmShartOfAcc
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(fmShartOfAcc))
        Me.BtnPreview = New System.Windows.Forms.Button
        Me.CmbPrinSelete = New System.Windows.Forms.ComboBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.similar = New System.Windows.Forms.RadioButton
        Me.Button5 = New System.Windows.Forms.Button
        Me.Rdlasth = New System.Windows.Forms.RadioButton
        Me.ChbLang = New System.Windows.Forms.CheckBox
        Me.Button4 = New System.Windows.Forms.Button
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.RdName = New System.Windows.Forms.RadioButton
        Me.RdId = New System.Windows.Forms.RadioButton
        Me.txtSearchName = New System.Windows.Forms.TextBox
        Me.txtSearchId = New System.Windows.Forms.TextBox
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtNewId = New System.Windows.Forms.TextBox
        Me.txtOldId = New System.Windows.Forms.TextBox
        Me.Button7 = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.BtnDelete = New System.Windows.Forms.Button
        Me.BtnEdit = New System.Windows.Forms.Button
        Me.BtnExit = New System.Windows.Forms.Button
        Me.BntNew = New System.Windows.Forms.Button
        Me.FG = New AxVSFlex8U.AxVSFlexGrid
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.BackPage = New System.Windows.Forms.Button
        Me.lblpage_total = New System.Windows.Forms.TextBox
        Me.NextPage = New System.Windows.Forms.Button
        Me.FirstPage = New System.Windows.Forms.Button
        Me.LasthPage = New System.Windows.Forms.Button
        Me.GrPage = New System.Windows.Forms.GroupBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.p25 = New System.Windows.Forms.RadioButton
        Me.Button8 = New System.Windows.Forms.Button
        Me.p1000 = New System.Windows.Forms.RadioButton
        Me.CmbPage = New System.Windows.Forms.ComboBox
        Me.p500 = New System.Windows.Forms.RadioButton
        Me.p250 = New System.Windows.Forms.RadioButton
        Me.p100 = New System.Windows.Forms.RadioButton
        Me.p50 = New System.Windows.Forms.RadioButton
        Me.CheckBox3 = New System.Windows.Forms.CheckBox
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.txtLng = New System.Windows.Forms.TextBox
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.FG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GrPage.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'BtnPreview
        '
        Me.BtnPreview.Image = CType(resources.GetObject("BtnPreview.Image"), System.Drawing.Image)
        Me.BtnPreview.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnPreview.Location = New System.Drawing.Point(468, 1)
        Me.BtnPreview.Name = "BtnPreview"
        Me.BtnPreview.Size = New System.Drawing.Size(84, 35)
        Me.BtnPreview.TabIndex = 120
        Me.BtnPreview.Tag = "3006"
        Me.BtnPreview.Text = "ວິວ/ເບິ່ງ"
        Me.BtnPreview.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnPreview.UseVisualStyleBackColor = True
        '
        'CmbPrinSelete
        '
        Me.CmbPrinSelete.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbPrinSelete.FormattingEnabled = True
        Me.CmbPrinSelete.Items.AddRange(New Object() {"ພີມທຸຫລາຍການ (Prin All)", "ເລືອກລາຍການພິມ (Prin All Items)"})
        Me.CmbPrinSelete.Location = New System.Drawing.Point(715, 4)
        Me.CmbPrinSelete.Name = "CmbPrinSelete"
        Me.CmbPrinSelete.Size = New System.Drawing.Size(211, 32)
        Me.CmbPrinSelete.TabIndex = 125
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(372, 1)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(96, 35)
        Me.Button2.TabIndex = 127
        Me.Button2.Tag = "3010"
        Me.Button2.Text = "ປ່ຽນລະຫັດ"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.similar)
        Me.Panel1.Controls.Add(Me.Button5)
        Me.Panel1.Controls.Add(Me.Rdlasth)
        Me.Panel1.Controls.Add(Me.ChbLang)
        Me.Panel1.Controls.Add(Me.Button4)
        Me.Panel1.Location = New System.Drawing.Point(389, 313)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(303, 70)
        Me.Panel1.TabIndex = 128
        Me.Panel1.Tag = "1"
        Me.Panel1.Visible = False
        '
        'similar
        '
        Me.similar.AutoSize = True
        Me.similar.Location = New System.Drawing.Point(116, 3)
        Me.similar.Name = "similar"
        Me.similar.Size = New System.Drawing.Size(84, 28)
        Me.similar.TabIndex = 1
        Me.similar.Tag = "5027"
        Me.similar.Text = "ຄຳທີບັນຈຸຢູ່"
        Me.similar.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(180, 29)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(75, 33)
        Me.Button5.TabIndex = 7
        Me.Button5.Tag = "3018"
        Me.Button5.Text = "ຍົກເລີກ"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Rdlasth
        '
        Me.Rdlasth.AutoSize = True
        Me.Rdlasth.Checked = True
        Me.Rdlasth.Location = New System.Drawing.Point(17, 3)
        Me.Rdlasth.Name = "Rdlasth"
        Me.Rdlasth.Size = New System.Drawing.Size(85, 28)
        Me.Rdlasth.TabIndex = 0
        Me.Rdlasth.TabStop = True
        Me.Rdlasth.Tag = "5026"
        Me.Rdlasth.Text = "ຄຳທີ່ຂື້ນຕົນ"
        Me.Rdlasth.UseVisualStyleBackColor = True
        '
        'ChbLang
        '
        Me.ChbLang.AutoSize = True
        Me.ChbLang.Enabled = False
        Me.ChbLang.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.ChbLang.Location = New System.Drawing.Point(309, 80)
        Me.ChbLang.Name = "ChbLang"
        Me.ChbLang.Size = New System.Drawing.Size(76, 28)
        Me.ChbLang.TabIndex = 6
        Me.ChbLang.Text = "Englisth"
        Me.ChbLang.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(10, 29)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(147, 33)
        Me.Button4.TabIndex = 0
        Me.Button4.Tag = "3026"
        Me.Button4.Text = "ຊອກຫາ"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Checked = True
        Me.CheckBox2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox2.Location = New System.Drawing.Point(116, 6)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(110, 28)
        Me.CheckBox2.TabIndex = 8
        Me.CheckBox2.Tag = "4004"
        Me.CheckBox2.Text = "ຊອກຫາທັງຫມົດ"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'RdName
        '
        Me.RdName.AutoSize = True
        Me.RdName.Location = New System.Drawing.Point(117, 71)
        Me.RdName.Name = "RdName"
        Me.RdName.Size = New System.Drawing.Size(83, 28)
        Me.RdName.TabIndex = 5
        Me.RdName.Tag = "5025"
        Me.RdName.Text = "ຕາມຊື່ບັນຊີ"
        Me.RdName.UseVisualStyleBackColor = True
        '
        'RdId
        '
        Me.RdId.AutoSize = True
        Me.RdId.Checked = True
        Me.RdId.Location = New System.Drawing.Point(15, 71)
        Me.RdId.Name = "RdId"
        Me.RdId.Size = New System.Drawing.Size(85, 28)
        Me.RdId.TabIndex = 4
        Me.RdId.TabStop = True
        Me.RdId.Tag = "5024"
        Me.RdId.Text = "ຕາມລະຫັດ"
        Me.RdId.UseVisualStyleBackColor = True
        '
        'txtSearchName
        '
        Me.txtSearchName.Enabled = False
        Me.txtSearchName.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSearchName.Location = New System.Drawing.Point(10, 37)
        Me.txtSearchName.Name = "txtSearchName"
        Me.txtSearchName.Size = New System.Drawing.Size(282, 30)
        Me.txtSearchName.TabIndex = 2
        '
        'txtSearchId
        '
        Me.txtSearchId.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSearchId.Location = New System.Drawing.Point(10, 3)
        Me.txtSearchId.Name = "txtSearchId"
        Me.txtSearchId.Size = New System.Drawing.Size(100, 30)
        Me.txtSearchId.TabIndex = 1
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.txtNewId)
        Me.Panel2.Controls.Add(Me.txtOldId)
        Me.Panel2.Controls.Add(Me.Button7)
        Me.Panel2.Controls.Add(Me.Button6)
        Me.Panel2.Location = New System.Drawing.Point(370, 38)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(354, 122)
        Me.Panel2.TabIndex = 129
        Me.Panel2.Tag = "1"
        Me.Panel2.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(22, 49)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(74, 24)
        Me.Label2.TabIndex = 5
        Me.Label2.Tag = "2036"
        Me.Label2.Text = "ລະຫັດໃໝ່ :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(21, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 24)
        Me.Label1.TabIndex = 5
        Me.Label1.Tag = "2035"
        Me.Label1.Text = "ລະຫັດເກົ່າ :"
        '
        'txtNewId
        '
        Me.txtNewId.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNewId.ForeColor = System.Drawing.Color.Blue
        Me.txtNewId.Location = New System.Drawing.Point(134, 46)
        Me.txtNewId.Name = "txtNewId"
        Me.txtNewId.Size = New System.Drawing.Size(199, 30)
        Me.txtNewId.TabIndex = 3
        '
        'txtOldId
        '
        Me.txtOldId.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtOldId.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOldId.ForeColor = System.Drawing.Color.Blue
        Me.txtOldId.Location = New System.Drawing.Point(134, 12)
        Me.txtOldId.Name = "txtOldId"
        Me.txtOldId.ReadOnly = True
        Me.txtOldId.Size = New System.Drawing.Size(199, 30)
        Me.txtOldId.TabIndex = 2
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(240, 84)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(93, 32)
        Me.Button7.TabIndex = 1
        Me.Button7.Tag = "3018"
        Me.Button7.Text = "ຍົກເລິກ"
        Me.Button7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(134, 84)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(100, 32)
        Me.Button6.TabIndex = 0
        Me.Button6.Tag = "3017"
        Me.Button6.Text = "ຕົກລົງ"
        Me.Button6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(289, 1)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(84, 35)
        Me.Button1.TabIndex = 126
        Me.Button1.Tag = "3026"
        Me.Button1.Text = "ຊອກຫາ"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.UseVisualStyleBackColor = True
        '
        'BtnDelete
        '
        Me.BtnDelete.Image = CType(resources.GetObject("BtnDelete.Image"), System.Drawing.Image)
        Me.BtnDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnDelete.Location = New System.Drawing.Point(206, 1)
        Me.BtnDelete.Name = "BtnDelete"
        Me.BtnDelete.Size = New System.Drawing.Size(84, 35)
        Me.BtnDelete.TabIndex = 124
        Me.BtnDelete.Tag = "3004"
        Me.BtnDelete.Text = "    ລືບ"
        Me.BtnDelete.UseVisualStyleBackColor = True
        '
        'BtnEdit
        '
        Me.BtnEdit.Image = CType(resources.GetObject("BtnEdit.Image"), System.Drawing.Image)
        Me.BtnEdit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnEdit.Location = New System.Drawing.Point(123, 1)
        Me.BtnEdit.Name = "BtnEdit"
        Me.BtnEdit.Size = New System.Drawing.Size(84, 35)
        Me.BtnEdit.TabIndex = 123
        Me.BtnEdit.Tag = "3003"
        Me.BtnEdit.Text = "ແກ້ໄຂ"
        Me.BtnEdit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnEdit.UseVisualStyleBackColor = True
        '
        'BtnExit
        '
        Me.BtnExit.Image = CType(resources.GetObject("BtnExit.Image"), System.Drawing.Image)
        Me.BtnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnExit.Location = New System.Drawing.Point(5, 1)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(35, 35)
        Me.BtnExit.TabIndex = 121
        Me.BtnExit.Tag = "9999"
        Me.BtnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'BntNew
        '
        Me.BntNew.Image = CType(resources.GetObject("BntNew.Image"), System.Drawing.Image)
        Me.BntNew.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BntNew.Location = New System.Drawing.Point(40, 1)
        Me.BntNew.Name = "BntNew"
        Me.BntNew.Size = New System.Drawing.Size(84, 35)
        Me.BntNew.TabIndex = 119
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
        Me.FG.DataSource = Nothing
        Me.FG.Location = New System.Drawing.Point(5, 42)
        Me.FG.Name = "FG"
        Me.FG.OcxState = CType(resources.GetObject("FG.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FG.Size = New System.Drawing.Size(1010, 667)
        Me.FG.TabIndex = 1
        Me.FG.Tag = "8003"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.CheckBox1.Location = New System.Drawing.Point(563, 7)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(141, 28)
        Me.CheckBox1.TabIndex = 131
        Me.CheckBox1.Tag = "4005"
        Me.CheckBox1.Text = "ພີມສະເພາະບັນຊີແມ່"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(289, 42)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 23)
        Me.Button3.TabIndex = 132
        Me.Button3.Text = "Button3"
        Me.Button3.UseVisualStyleBackColor = True
        Me.Button3.Visible = False
        '
        'BackPage
        '
        Me.BackPage.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BackPage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.BackPage.ForeColor = System.Drawing.Color.Blue
        Me.BackPage.Location = New System.Drawing.Point(163, 12)
        Me.BackPage.Name = "BackPage"
        Me.BackPage.Size = New System.Drawing.Size(37, 23)
        Me.BackPage.TabIndex = 210
        Me.BackPage.Text = "<<"
        Me.BackPage.UseVisualStyleBackColor = True
        '
        'lblpage_total
        '
        Me.lblpage_total.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblpage_total.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblpage_total.ForeColor = System.Drawing.Color.Blue
        Me.lblpage_total.Location = New System.Drawing.Point(200, 13)
        Me.lblpage_total.Name = "lblpage_total"
        Me.lblpage_total.ReadOnly = True
        Me.lblpage_total.Size = New System.Drawing.Size(61, 20)
        Me.lblpage_total.TabIndex = 206
        Me.lblpage_total.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'NextPage
        '
        Me.NextPage.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.NextPage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NextPage.ForeColor = System.Drawing.Color.Blue
        Me.NextPage.Location = New System.Drawing.Point(260, 12)
        Me.NextPage.Name = "NextPage"
        Me.NextPage.Size = New System.Drawing.Size(37, 23)
        Me.NextPage.TabIndex = 207
        Me.NextPage.Text = ">>"
        Me.NextPage.UseVisualStyleBackColor = True
        '
        'FirstPage
        '
        Me.FirstPage.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.FirstPage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.FirstPage.ForeColor = System.Drawing.Color.Blue
        Me.FirstPage.Location = New System.Drawing.Point(125, 12)
        Me.FirstPage.Name = "FirstPage"
        Me.FirstPage.Size = New System.Drawing.Size(39, 23)
        Me.FirstPage.TabIndex = 209
        Me.FirstPage.Text = "|<<"
        Me.FirstPage.UseVisualStyleBackColor = True
        '
        'LasthPage
        '
        Me.LasthPage.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LasthPage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.LasthPage.ForeColor = System.Drawing.Color.Blue
        Me.LasthPage.Location = New System.Drawing.Point(296, 12)
        Me.LasthPage.Name = "LasthPage"
        Me.LasthPage.Size = New System.Drawing.Size(38, 23)
        Me.LasthPage.TabIndex = 208
        Me.LasthPage.Text = ">>|"
        Me.LasthPage.UseVisualStyleBackColor = True
        '
        'GrPage
        '
        Me.GrPage.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GrPage.Controls.Add(Me.Label3)
        Me.GrPage.Controls.Add(Me.p25)
        Me.GrPage.Controls.Add(Me.Button8)
        Me.GrPage.Controls.Add(Me.p1000)
        Me.GrPage.Controls.Add(Me.CmbPage)
        Me.GrPage.Controls.Add(Me.p500)
        Me.GrPage.Controls.Add(Me.p250)
        Me.GrPage.Controls.Add(Me.p100)
        Me.GrPage.Controls.Add(Me.p50)
        Me.GrPage.Controls.Add(Me.BackPage)
        Me.GrPage.Controls.Add(Me.FirstPage)
        Me.GrPage.Controls.Add(Me.lblpage_total)
        Me.GrPage.Controls.Add(Me.LasthPage)
        Me.GrPage.Controls.Add(Me.NextPage)
        Me.GrPage.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GrPage.Location = New System.Drawing.Point(6, 711)
        Me.GrPage.Name = "GrPage"
        Me.GrPage.Size = New System.Drawing.Size(638, 37)
        Me.GrPage.TabIndex = 211
        Me.GrPage.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Blue
        Me.Label3.Location = New System.Drawing.Point(6, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(78, 15)
        Me.Label3.TabIndex = 219
        Me.Label3.Text = "Select Page :"
        '
        'p25
        '
        Me.p25.AutoSize = True
        Me.p25.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.p25.Location = New System.Drawing.Point(344, 15)
        Me.p25.Name = "p25"
        Me.p25.Size = New System.Drawing.Size(37, 19)
        Me.p25.TabIndex = 216
        Me.p25.Text = "25"
        Me.p25.UseVisualStyleBackColor = True
        '
        'Button8
        '
        Me.Button8.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button8.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button8.ForeColor = System.Drawing.Color.Blue
        Me.Button8.Location = New System.Drawing.Point(2, 11)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(78, 23)
        Me.Button8.TabIndex = 218
        Me.Button8.Text = "Start"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'p1000
        '
        Me.p1000.AutoSize = True
        Me.p1000.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.p1000.Location = New System.Drawing.Point(582, 14)
        Me.p1000.Name = "p1000"
        Me.p1000.Size = New System.Drawing.Size(52, 19)
        Me.p1000.TabIndex = 215
        Me.p1000.Text = "1,000"
        Me.p1000.UseVisualStyleBackColor = True
        '
        'CmbPage
        '
        Me.CmbPage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbPage.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbPage.FormattingEnabled = True
        Me.CmbPage.Location = New System.Drawing.Point(80, 12)
        Me.CmbPage.Name = "CmbPage"
        Me.CmbPage.Size = New System.Drawing.Size(45, 22)
        Me.CmbPage.TabIndex = 217
        '
        'p500
        '
        Me.p500.AutoSize = True
        Me.p500.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.p500.Location = New System.Drawing.Point(532, 13)
        Me.p500.Name = "p500"
        Me.p500.Size = New System.Drawing.Size(43, 19)
        Me.p500.TabIndex = 214
        Me.p500.Text = "500"
        Me.p500.UseVisualStyleBackColor = True
        '
        'p250
        '
        Me.p250.AutoSize = True
        Me.p250.Checked = True
        Me.p250.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.p250.Location = New System.Drawing.Point(482, 14)
        Me.p250.Name = "p250"
        Me.p250.Size = New System.Drawing.Size(43, 19)
        Me.p250.TabIndex = 213
        Me.p250.TabStop = True
        Me.p250.Text = "250"
        Me.p250.UseVisualStyleBackColor = True
        '
        'p100
        '
        Me.p100.AutoSize = True
        Me.p100.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.p100.Location = New System.Drawing.Point(432, 14)
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
        Me.p50.Location = New System.Drawing.Point(388, 14)
        Me.p50.Name = "p50"
        Me.p50.Size = New System.Drawing.Size(37, 19)
        Me.p50.TabIndex = 211
        Me.p50.Text = "50"
        Me.p50.UseVisualStyleBackColor = True
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.CheckBox3.Location = New System.Drawing.Point(932, 6)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(83, 28)
        Me.CheckBox3.TabIndex = 212
        Me.CheckBox3.Text = "ແບບສອງ"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.CheckBox2)
        Me.Panel3.Controls.Add(Me.RdId)
        Me.Panel3.Controls.Add(Me.RdName)
        Me.Panel3.Controls.Add(Me.txtSearchId)
        Me.Panel3.Controls.Add(Me.txtSearchName)
        Me.Panel3.Location = New System.Drawing.Point(389, 215)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(303, 99)
        Me.Panel3.TabIndex = 213
        Me.Panel3.Tag = "1"
        Me.Panel3.Visible = False
        '
        'txtLng
        '
        Me.txtLng.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtLng.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtLng.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLng.Location = New System.Drawing.Point(588, 4)
        Me.txtLng.Name = "txtLng"
        Me.txtLng.ReadOnly = True
        Me.txtLng.Size = New System.Drawing.Size(34, 30)
        Me.txtLng.TabIndex = 214
        Me.txtLng.Text = "LNG"
        Me.txtLng.Visible = False
        '
        'fmShartOfAcc
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1020, 749)
        Me.Controls.Add(Me.txtLng)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.CheckBox3)
        Me.Controls.Add(Me.GrPage)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.CmbPrinSelete)
        Me.Controls.Add(Me.BtnDelete)
        Me.Controls.Add(Me.BtnEdit)
        Me.Controls.Add(Me.BtnPreview)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.BntNew)
        Me.Controls.Add(Me.FG)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "fmShartOfAcc"
        Me.Text = "frmShartOfAcc"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GrPage.ResumeLayout(False)
        Me.GrPage.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FG As AxVSFlex8U.AxVSFlexGrid
    Friend WithEvents BtnDelete As System.Windows.Forms.Button
    Friend WithEvents BtnEdit As System.Windows.Forms.Button
    Friend WithEvents BtnPreview As System.Windows.Forms.Button
    Friend WithEvents BtnExit As System.Windows.Forms.Button
    Friend WithEvents BntNew As System.Windows.Forms.Button
    Friend WithEvents CmbPrinSelete As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtSearchName As System.Windows.Forms.TextBox
    Friend WithEvents txtSearchId As System.Windows.Forms.TextBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents RdName As System.Windows.Forms.RadioButton
    Friend WithEvents RdId As System.Windows.Forms.RadioButton
    Friend WithEvents similar As System.Windows.Forms.RadioButton
    Friend WithEvents Rdlasth As System.Windows.Forms.RadioButton
    Friend WithEvents ChbLang As System.Windows.Forms.CheckBox
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtNewId As System.Windows.Forms.TextBox
    Friend WithEvents txtOldId As System.Windows.Forms.TextBox
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents BackPage As System.Windows.Forms.Button
    Friend WithEvents lblpage_total As System.Windows.Forms.TextBox
    Friend WithEvents NextPage As System.Windows.Forms.Button
    Friend WithEvents FirstPage As System.Windows.Forms.Button
    Friend WithEvents LasthPage As System.Windows.Forms.Button
    Friend WithEvents GrPage As System.Windows.Forms.GroupBox
    Friend WithEvents p50 As System.Windows.Forms.RadioButton
    Friend WithEvents p1000 As System.Windows.Forms.RadioButton
    Friend WithEvents p500 As System.Windows.Forms.RadioButton
    Friend WithEvents p250 As System.Windows.Forms.RadioButton
    Friend WithEvents p100 As System.Windows.Forms.RadioButton
    Friend WithEvents p25 As System.Windows.Forms.RadioButton
    Friend WithEvents CmbPage As System.Windows.Forms.ComboBox
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents txtLng As System.Windows.Forms.TextBox
End Class
