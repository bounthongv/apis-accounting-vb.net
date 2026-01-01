<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_StatementOld
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_StatementOld))
        Me.FG = New System.Windows.Forms.DataGridView
        Me.Label11 = New System.Windows.Forms.Label
        Me.BtnExit = New System.Windows.Forms.Button
        Me.BtnPreview = New System.Windows.Forms.Button
        Me.Dt = New System.Windows.Forms.DateTimePicker
        Me.Ds = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtAccCode = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtAccName = New System.Windows.Forms.TextBox
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.TxtCustID = New System.Windows.Forms.TextBox
        Me.TxtSuppID = New System.Windows.Forms.TextBox
        Me.CmbSupp = New System.Windows.Forms.ComboBox
        Me.CmbCust = New System.Windows.Forms.ComboBox
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.CheckBox4 = New System.Windows.Forms.CheckBox
        Me.BtnSearch = New System.Windows.Forms.Button
        Me.TxtOpen = New System.Windows.Forms.TextBox
        Me.TxtDebit = New System.Windows.Forms.TextBox
        Me.TxtCredit = New System.Windows.Forms.TextBox
        Me.TxtEnd = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.CMB_Curr = New System.Windows.Forms.ComboBox
        CType(Me.FG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'FG
        '
        Me.FG.AllowUserToAddRows = False
        Me.FG.AllowUserToDeleteRows = False
        Me.FG.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.FG.Location = New System.Drawing.Point(16, 168)
        Me.FG.Name = "FG"
        Me.FG.ReadOnly = True
        Me.FG.Size = New System.Drawing.Size(1230, 258)
        Me.FG.TabIndex = 303
        Me.FG.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.FG.MultiSelect = False
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Saysettha OT", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Blue
        Me.Label11.Location = New System.Drawing.Point(513, 4)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(270, 47)
        Me.Label11.TabIndex = 304
        Me.Label11.Text = "Account Statement"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BtnExit
        '
        Me.BtnExit.Image = Global.ApPBank10.My.Resources.Resources.Exit1
        Me.BtnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnExit.Location = New System.Drawing.Point(9, 8)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(35, 35)
        Me.BtnExit.TabIndex = 45545
        Me.BtnExit.Tag = "9999"
        Me.BtnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'BtnPreview
        '
        Me.BtnPreview.Image = Global.ApPBank10.My.Resources.Resources.Preview
        Me.BtnPreview.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnPreview.Location = New System.Drawing.Point(44, 8)
        Me.BtnPreview.Name = "BtnPreview"
        Me.BtnPreview.Size = New System.Drawing.Size(100, 35)
        Me.BtnPreview.TabIndex = 45544
        Me.BtnPreview.Tag = "3006"
        Me.BtnPreview.Text = "ວິວ/ເບິ່ງ"
        Me.BtnPreview.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnPreview.UseVisualStyleBackColor = True
        '
        'Dt
        '
        Me.Dt.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.Dt.Location = New System.Drawing.Point(355, 15)
        Me.Dt.Name = "Dt"
        Me.Dt.Size = New System.Drawing.Size(103, 30)
        Me.Dt.TabIndex = 45547
        '
        'Ds
        '
        Me.Ds.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.Ds.Location = New System.Drawing.Point(196, 12)
        Me.Ds.Name = "Ds"
        Me.Ds.Size = New System.Drawing.Size(110, 30)
        Me.Ds.TabIndex = 45546
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(151, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 21)
        Me.Label1.TabIndex = 45548
        Me.Label1.Text = "Start"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(312, 15)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(24, 21)
        Me.Label2.TabIndex = 45549
        Me.Label2.Text = "To"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(40, 71)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(114, 21)
        Me.Label3.TabIndex = 45550
        Me.Label3.Text = "Account Number:"
        '
        'TxtAccCode
        '
        Me.TxtAccCode.Location = New System.Drawing.Point(150, 68)
        Me.TxtAccCode.Name = "TxtAccCode"
        Me.TxtAccCode.Size = New System.Drawing.Size(156, 30)
        Me.TxtAccCode.TabIndex = 45551
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(40, 105)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(68, 21)
        Me.Label4.TabIndex = 45552
        Me.Label4.Text = "Currency:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(42, 135)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(102, 21)
        Me.Label5.TabIndex = 45553
        Me.Label5.Text = "Account Name:"
        '
        'TxtAccName
        '
        Me.TxtAccName.Location = New System.Drawing.Point(150, 132)
        Me.TxtAccName.Name = "TxtAccName"
        Me.TxtAccName.Size = New System.Drawing.Size(308, 30)
        Me.TxtAccName.TabIndex = 45554
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.TxtCustID)
        Me.Panel4.Controls.Add(Me.TxtSuppID)
        Me.Panel4.Controls.Add(Me.CmbSupp)
        Me.Panel4.Controls.Add(Me.CmbCust)
        Me.Panel4.Controls.Add(Me.RadioButton2)
        Me.Panel4.Controls.Add(Me.RadioButton1)
        Me.Panel4.Location = New System.Drawing.Point(532, 73)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(350, 89)
        Me.Panel4.TabIndex = 46034
        '
        'TxtCustID
        '
        Me.TxtCustID.Location = New System.Drawing.Point(318, 3)
        Me.TxtCustID.Name = "TxtCustID"
        Me.TxtCustID.Size = New System.Drawing.Size(63, 30)
        Me.TxtCustID.TabIndex = 128
        Me.TxtCustID.Visible = False
        '
        'TxtSuppID
        '
        Me.TxtSuppID.Location = New System.Drawing.Point(318, 38)
        Me.TxtSuppID.Name = "TxtSuppID"
        Me.TxtSuppID.Size = New System.Drawing.Size(63, 30)
        Me.TxtSuppID.TabIndex = 127
        Me.TxtSuppID.Visible = False
        '
        'CmbSupp
        '
        Me.CmbSupp.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbSupp.FormattingEnabled = True
        Me.CmbSupp.Location = New System.Drawing.Point(101, 34)
        Me.CmbSupp.Name = "CmbSupp"
        Me.CmbSupp.Size = New System.Drawing.Size(211, 29)
        Me.CmbSupp.TabIndex = 126
        '
        'CmbCust
        '
        Me.CmbCust.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbCust.FormattingEnabled = True
        Me.CmbCust.Location = New System.Drawing.Point(101, 3)
        Me.CmbCust.Name = "CmbCust"
        Me.CmbCust.Size = New System.Drawing.Size(211, 29)
        Me.CmbCust.TabIndex = 125
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(11, 34)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(76, 25)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.Text = "Supplier"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(11, 3)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(84, 25)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Customer"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'CheckBox4
        '
        Me.CheckBox4.AutoSize = True
        Me.CheckBox4.Location = New System.Drawing.Point(355, 73)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(171, 25)
        Me.CheckBox4.TabIndex = 46033
        Me.CheckBox4.Text = "Only Customer/Supplier"
        Me.CheckBox4.UseVisualStyleBackColor = True
        '
        'BtnSearch
        '
        Me.BtnSearch.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSearch.Location = New System.Drawing.Point(312, 68)
        Me.BtnSearch.Name = "BtnSearch"
        Me.BtnSearch.Size = New System.Drawing.Size(34, 30)
        Me.BtnSearch.TabIndex = 46035
        Me.BtnSearch.Tag = "3012"
        Me.BtnSearch.Text = "....."
        Me.BtnSearch.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnSearch.UseVisualStyleBackColor = True
        '
        'TxtOpen
        '
        Me.TxtOpen.Location = New System.Drawing.Point(1075, 29)
        Me.TxtOpen.Name = "TxtOpen"
        Me.TxtOpen.Size = New System.Drawing.Size(171, 30)
        Me.TxtOpen.TabIndex = 46036
        '
        'TxtDebit
        '
        Me.TxtDebit.Location = New System.Drawing.Point(1075, 64)
        Me.TxtDebit.Name = "TxtDebit"
        Me.TxtDebit.Size = New System.Drawing.Size(171, 30)
        Me.TxtDebit.TabIndex = 46037
        '
        'TxtCredit
        '
        Me.TxtCredit.Location = New System.Drawing.Point(1075, 97)
        Me.TxtCredit.Name = "TxtCredit"
        Me.TxtCredit.Size = New System.Drawing.Size(171, 30)
        Me.TxtCredit.TabIndex = 46038
        '
        'TxtEnd
        '
        Me.TxtEnd.Location = New System.Drawing.Point(1075, 129)
        Me.TxtEnd.Name = "TxtEnd"
        Me.TxtEnd.Size = New System.Drawing.Size(171, 30)
        Me.TxtEnd.TabIndex = 46039
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(953, 35)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(116, 21)
        Me.Label6.TabIndex = 46040
        Me.Label6.Text = "Opening Balance:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(925, 69)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(144, 21)
        Me.Label7.TabIndex = 46041
        Me.Label7.Text = "Total Movement-Debit:"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(920, 103)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(149, 21)
        Me.Label8.TabIndex = 46042
        Me.Label8.Text = "Total Movement-Credit:"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(961, 135)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(108, 21)
        Me.Label9.TabIndex = 46043
        Me.Label9.Text = "Ending Balance:"
        '
        'CMB_Curr
        '
        Me.CMB_Curr.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CMB_Curr.FormattingEnabled = True
        Me.CMB_Curr.Items.AddRange(New Object() {"LAK", "THB", "USD"})
        Me.CMB_Curr.Location = New System.Drawing.Point(150, 100)
        Me.CMB_Curr.Name = "CMB_Curr"
        Me.CMB_Curr.Size = New System.Drawing.Size(77, 30)
        Me.CMB_Curr.TabIndex = 46044
        Me.CMB_Curr.Text = "LAK"
        '
        'Frm_Statement
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 21.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1258, 439)
        Me.Controls.Add(Me.CMB_Curr)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.TxtEnd)
        Me.Controls.Add(Me.TxtCredit)
        Me.Controls.Add(Me.TxtDebit)
        Me.Controls.Add(Me.TxtOpen)
        Me.Controls.Add(Me.BtnSearch)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.CheckBox4)
        Me.Controls.Add(Me.TxtAccName)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TxtAccCode)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Dt)
        Me.Controls.Add(Me.Ds)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.BtnPreview)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.FG)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "Frm_Statement"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Frm_Statement"
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FG As System.Windows.Forms.DataGridView
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents BtnExit As System.Windows.Forms.Button
    Friend WithEvents BtnPreview As System.Windows.Forms.Button
    Friend WithEvents Dt As System.Windows.Forms.DateTimePicker
    Friend WithEvents Ds As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtAccCode As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtAccName As System.Windows.Forms.TextBox
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents TxtCustID As System.Windows.Forms.TextBox
    Friend WithEvents TxtSuppID As System.Windows.Forms.TextBox
    Friend WithEvents CmbSupp As System.Windows.Forms.ComboBox
    Friend WithEvents CmbCust As System.Windows.Forms.ComboBox
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents CheckBox4 As System.Windows.Forms.CheckBox
    Friend WithEvents BtnSearch As System.Windows.Forms.Button
    Friend WithEvents TxtOpen As System.Windows.Forms.TextBox
    Friend WithEvents TxtDebit As System.Windows.Forms.TextBox
    Friend WithEvents TxtCredit As System.Windows.Forms.TextBox
    Friend WithEvents TxtEnd As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents CMB_Curr As System.Windows.Forms.ComboBox
End Class
