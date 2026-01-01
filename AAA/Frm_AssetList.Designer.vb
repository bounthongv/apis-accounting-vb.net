<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_AssetList
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_AssetList))
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtCode = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtName = New System.Windows.Forms.TextBox
        Me.TxtNameE = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtValue = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtRemain = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.FG = New System.Windows.Forms.DataGridView
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtGrp = New System.Windows.Forms.TextBox
        Me.txtGrpNm = New System.Windows.Forms.ComboBox
        Me.LEng = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.TxtDesription = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.DateIn = New System.Windows.Forms.DateTimePicker
        Me.TxtPeriod = New System.Windows.Forms.TextBox
        Me.TxtDr = New System.Windows.Forms.TextBox
        Me.TxtCr = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.BtnSearch = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.TxtDrNm = New System.Windows.Forms.TextBox
        Me.TxtCrNm = New System.Windows.Forms.TextBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.BtnAddNew2 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.BtnRefresh = New System.Windows.Forms.Button
        Me.BtnEdit = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        CType(Me.FG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(66, 51)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 27)
        Me.Label1.TabIndex = 283
        Me.Label1.Text = "Code:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TxtCode
        '
        Me.TxtCode.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtCode.Location = New System.Drawing.Point(156, 49)
        Me.TxtCode.Name = "TxtCode"
        Me.TxtCode.Size = New System.Drawing.Size(100, 30)
        Me.TxtCode.TabIndex = 284
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(66, 117)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(84, 27)
        Me.Label2.TabIndex = 285
        Me.Label2.Text = "Name (LA):"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(53, 180)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(94, 27)
        Me.Label3.TabIndex = 286
        Me.Label3.Text = "Value:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TxtName
        '
        Me.TxtName.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtName.Location = New System.Drawing.Point(156, 116)
        Me.TxtName.Name = "TxtName"
        Me.TxtName.Size = New System.Drawing.Size(263, 30)
        Me.TxtName.TabIndex = 287
        '
        'TxtNameE
        '
        Me.TxtNameE.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtNameE.Location = New System.Drawing.Point(156, 147)
        Me.TxtNameE.Name = "TxtNameE"
        Me.TxtNameE.Size = New System.Drawing.Size(263, 30)
        Me.TxtNameE.TabIndex = 288
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(434, 115)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(104, 27)
        Me.Label4.TabIndex = 289
        Me.Label4.Text = "Desription:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TxtValue
        '
        Me.TxtValue.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtValue.Location = New System.Drawing.Point(156, 179)
        Me.TxtValue.Name = "TxtValue"
        Me.TxtValue.Size = New System.Drawing.Size(263, 30)
        Me.TxtValue.TabIndex = 290
        Me.TxtValue.Text = "0"
        Me.TxtValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(679, 147)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(116, 27)
        Me.Label5.TabIndex = 292
        Me.Label5.Text = "Adjust Period:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label6
        '
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(53, 215)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(94, 27)
        Me.Label6.TabIndex = 293
        Me.Label6.Text = "Remain:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TxtRemain
        '
        Me.TxtRemain.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtRemain.Location = New System.Drawing.Point(156, 211)
        Me.TxtRemain.Name = "TxtRemain"
        Me.TxtRemain.Size = New System.Drawing.Size(263, 30)
        Me.TxtRemain.TabIndex = 294
        Me.TxtRemain.Text = "0"
        Me.TxtRemain.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label9
        '
        Me.Label9.ForeColor = System.Drawing.Color.Black
        Me.Label9.Location = New System.Drawing.Point(444, 180)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(94, 27)
        Me.Label9.TabIndex = 300
        Me.Label9.Text = "Dr:"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'FG
        '
        Me.FG.AllowUserToAddRows = False
        Me.FG.AllowUserToDeleteRows = False
        Me.FG.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.FG.Location = New System.Drawing.Point(16, 245)
        Me.FG.Name = "FG"
        Me.FG.ReadOnly = True
        Me.FG.Size = New System.Drawing.Size(1118, 173)
        Me.FG.TabIndex = 303
        Me.FG.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.FG.MultiSelect = False
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Saysettha OT", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Blue
        Me.Label11.Location = New System.Drawing.Point(752, 9)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(369, 36)
        Me.Label11.TabIndex = 304
        Me.Label11.Text = "Fixed Assets Register List"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtGrp
        '
        Me.txtGrp.Enabled = False
        Me.txtGrp.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGrp.Location = New System.Drawing.Point(319, 49)
        Me.txtGrp.Name = "txtGrp"
        Me.txtGrp.Size = New System.Drawing.Size(100, 30)
        Me.txtGrp.TabIndex = 307
        Me.txtGrp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtGrpNm
        '
        Me.txtGrpNm.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGrpNm.FormattingEnabled = True
        Me.txtGrpNm.Items.AddRange(New Object() {"LAK", "THB", "USD"})
        Me.txtGrpNm.Location = New System.Drawing.Point(156, 82)
        Me.txtGrpNm.Name = "txtGrpNm"
        Me.txtGrpNm.Size = New System.Drawing.Size(263, 32)
        Me.txtGrpNm.TabIndex = 306
        '
        'LEng
        '
        Me.LEng.Font = New System.Drawing.Font("Saysettha OT", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LEng.Location = New System.Drawing.Point(-4, 85)
        Me.LEng.Name = "LEng"
        Me.LEng.Size = New System.Drawing.Size(151, 24)
        Me.LEng.TabIndex = 305
        Me.LEng.Tag = "2010"
        Me.LEng.Text = "Adjustment Type:"
        Me.LEng.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label12
        '
        Me.Label12.ForeColor = System.Drawing.Color.Black
        Me.Label12.Location = New System.Drawing.Point(63, 149)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(84, 27)
        Me.Label12.TabIndex = 309
        Me.Label12.Text = "Name (EN):"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TxtDesription
        '
        Me.TxtDesription.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtDesription.Location = New System.Drawing.Point(539, 114)
        Me.TxtDesription.Name = "TxtDesription"
        Me.TxtDesription.Size = New System.Drawing.Size(341, 30)
        Me.TxtDesription.TabIndex = 310
        '
        'Label13
        '
        Me.Label13.ForeColor = System.Drawing.Color.Black
        Me.Label13.Location = New System.Drawing.Point(432, 148)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(104, 27)
        Me.Label13.TabIndex = 311
        Me.Label13.Text = "Date In:"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DateIn
        '
        Me.DateIn.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateIn.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateIn.Location = New System.Drawing.Point(539, 147)
        Me.DateIn.Name = "DateIn"
        Me.DateIn.Size = New System.Drawing.Size(120, 30)
        Me.DateIn.TabIndex = 312
        '
        'TxtPeriod
        '
        Me.TxtPeriod.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtPeriod.Location = New System.Drawing.Point(797, 146)
        Me.TxtPeriod.Name = "TxtPeriod"
        Me.TxtPeriod.Size = New System.Drawing.Size(83, 30)
        Me.TxtPeriod.TabIndex = 313
        Me.TxtPeriod.Text = "0"
        Me.TxtPeriod.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtDr
        '
        Me.TxtDr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtDr.Location = New System.Drawing.Point(539, 180)
        Me.TxtDr.Name = "TxtDr"
        Me.TxtDr.Size = New System.Drawing.Size(163, 30)
        Me.TxtDr.TabIndex = 314
        '
        'TxtCr
        '
        Me.TxtCr.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtCr.Location = New System.Drawing.Point(539, 212)
        Me.TxtCr.Name = "TxtCr"
        Me.TxtCr.Size = New System.Drawing.Size(163, 30)
        Me.TxtCr.TabIndex = 315
        '
        'Label7
        '
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(443, 215)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(94, 27)
        Me.Label7.TabIndex = 316
        Me.Label7.Text = "Cr:"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'BtnSearch
        '
        Me.BtnSearch.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSearch.Location = New System.Drawing.Point(702, 179)
        Me.BtnSearch.Name = "BtnSearch"
        Me.BtnSearch.Size = New System.Drawing.Size(34, 30)
        Me.BtnSearch.TabIndex = 46037
        Me.BtnSearch.Tag = "3012"
        Me.BtnSearch.Text = "....."
        Me.BtnSearch.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnSearch.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.Location = New System.Drawing.Point(702, 212)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(34, 30)
        Me.Button4.TabIndex = 46038
        Me.Button4.Tag = "3012"
        Me.Button4.Text = "....."
        Me.Button4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button4.UseVisualStyleBackColor = True
        '
        'TxtDrNm
        '
        Me.TxtDrNm.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtDrNm.Location = New System.Drawing.Point(742, 179)
        Me.TxtDrNm.Name = "TxtDrNm"
        Me.TxtDrNm.Size = New System.Drawing.Size(392, 30)
        Me.TxtDrNm.TabIndex = 46039
        '
        'TxtCrNm
        '
        Me.TxtCrNm.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtCrNm.Location = New System.Drawing.Point(742, 211)
        Me.TxtCrNm.Name = "TxtCrNm"
        Me.TxtCrNm.Size = New System.Drawing.Size(392, 30)
        Me.TxtCrNm.TabIndex = 46040
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Image = CType(resources.GetObject("Button3.Image"), System.Drawing.Image)
        Me.Button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button3.Location = New System.Drawing.Point(256, 8)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(100, 35)
        Me.Button3.TabIndex = 282
        Me.Button3.Tag = "3004"
        Me.Button3.Text = "ລຶບ"
        Me.Button3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button3.UseVisualStyleBackColor = True
        '
        'BtnAddNew2
        '
        Me.BtnAddNew2.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAddNew2.Image = CType(resources.GetObject("BtnAddNew2.Image"), System.Drawing.Image)
        Me.BtnAddNew2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAddNew2.Location = New System.Drawing.Point(42, 8)
        Me.BtnAddNew2.Name = "BtnAddNew2"
        Me.BtnAddNew2.Size = New System.Drawing.Size(114, 35)
        Me.BtnAddNew2.TabIndex = 281
        Me.BtnAddNew2.Tag = "3001"
        Me.BtnAddNew2.Text = "ເພີ່ມໃຫມ່"
        Me.BtnAddNew2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnAddNew2.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(622, 13)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(114, 35)
        Me.Button2.TabIndex = 280
        Me.Button2.Tag = "3002"
        Me.Button2.Text = "ບັນທຶກ"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(7, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(35, 35)
        Me.Button1.TabIndex = 279
        Me.Button1.Tag = "9999"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.UseVisualStyleBackColor = True
        '
        'BtnRefresh
        '
        Me.BtnRefresh.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.BtnRefresh.Image = CType(resources.GetObject("BtnRefresh.Image"), System.Drawing.Image)
        Me.BtnRefresh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnRefresh.Location = New System.Drawing.Point(356, 8)
        Me.BtnRefresh.Name = "BtnRefresh"
        Me.BtnRefresh.Size = New System.Drawing.Size(100, 34)
        Me.BtnRefresh.TabIndex = 46042
        Me.BtnRefresh.Tag = "3005"
        Me.BtnRefresh.Text = "ເອີ້ນຂໍ້ມູນ"
        Me.BtnRefresh.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnRefresh.UseVisualStyleBackColor = False
        '
        'BtnEdit
        '
        Me.BtnEdit.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.BtnEdit.Image = CType(resources.GetObject("BtnEdit.Image"), System.Drawing.Image)
        Me.BtnEdit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnEdit.Location = New System.Drawing.Point(156, 8)
        Me.BtnEdit.Name = "BtnEdit"
        Me.BtnEdit.Size = New System.Drawing.Size(100, 34)
        Me.BtnEdit.TabIndex = 46041
        Me.BtnEdit.Tag = "3003"
        Me.BtnEdit.Text = "ແກ້ໄຂ"
        Me.BtnEdit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnEdit.UseVisualStyleBackColor = False
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Button5.Image = CType(resources.GetObject("Button5.Image"), System.Drawing.Image)
        Me.Button5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button5.Location = New System.Drawing.Point(456, 8)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(103, 35)
        Me.Button5.TabIndex = 46043
        Me.Button5.Tag = "3006"
        Me.Button5.Text = "ເບິ່ງ/ພີມ"
        Me.Button5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button5.UseVisualStyleBackColor = False
        '
        'Frm_AssetList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 21.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1146, 423)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.BtnRefresh)
        Me.Controls.Add(Me.BtnEdit)
        Me.Controls.Add(Me.TxtCrNm)
        Me.Controls.Add(Me.TxtDrNm)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.BtnSearch)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.TxtCr)
        Me.Controls.Add(Me.TxtDr)
        Me.Controls.Add(Me.TxtPeriod)
        Me.Controls.Add(Me.DateIn)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.TxtDesription)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.txtGrp)
        Me.Controls.Add(Me.txtGrpNm)
        Me.Controls.Add(Me.LEng)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.FG)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.TxtRemain)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TxtValue)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TxtNameE)
        Me.Controls.Add(Me.TxtName)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TxtCode)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.BtnAddNew2)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "Frm_AssetList"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Frm_AssetList"
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents BtnAddNew2 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtCode As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtName As System.Windows.Forms.TextBox
    Friend WithEvents TxtNameE As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtValue As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtRemain As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents FG As System.Windows.Forms.DataGridView
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtGrp As System.Windows.Forms.TextBox
    Friend WithEvents txtGrpNm As System.Windows.Forms.ComboBox
    Friend WithEvents LEng As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TxtDesription As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents DateIn As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtPeriod As System.Windows.Forms.TextBox
    Friend WithEvents TxtDr As System.Windows.Forms.TextBox
    Friend WithEvents TxtCr As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents BtnSearch As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents TxtDrNm As System.Windows.Forms.TextBox
    Friend WithEvents TxtCrNm As System.Windows.Forms.TextBox
    Friend WithEvents BtnRefresh As System.Windows.Forms.Button
    Friend WithEvents BtnEdit As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
End Class
