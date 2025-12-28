<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FmRptProItem
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FmRptProItem))
        Me.DG = New System.Windows.Forms.DataGridView
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.AC_Code = New System.Windows.Forms.TextBox
        Me.Rpt_Type = New System.Windows.Forms.ComboBox
        Me.RPT_ID = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.BtnMove = New System.Windows.Forms.Button
        Me.BtnSearch = New System.Windows.Forms.Button
        Me.DgItems = New System.Windows.Forms.DataGridView
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.ComCurr = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.CAmt = New System.Windows.Forms.CheckBox
        Me.COP = New System.Windows.Forms.CheckBox
        Me.Button6 = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.BtnExit = New System.Windows.Forms.Button
        Me.BtnEdit = New System.Windows.Forms.Button
        CType(Me.DG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DgItems, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'DG
        '
        Me.DG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DG.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DG.Location = New System.Drawing.Point(3, 30)
        Me.DG.Name = "DG"
        Me.DG.Size = New System.Drawing.Size(439, 366)
        Me.DG.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(637, 27)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(55, 24)
        Me.Label3.TabIndex = 182
        Me.Label3.Text = "ປະເພດ"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(438, 26)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(66, 24)
        Me.Label2.TabIndex = 181
        Me.Label2.Text = "ເລກບັນຊີ"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(321, 28)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 24)
        Me.Label1.TabIndex = 180
        Me.Label1.Text = "ລະຫັດ"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(760, 22)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(56, 36)
        Me.Button3.TabIndex = 179
        Me.Button3.Text = "DEL"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'AC_Code
        '
        Me.AC_Code.Location = New System.Drawing.Point(504, 22)
        Me.AC_Code.Name = "AC_Code"
        Me.AC_Code.Size = New System.Drawing.Size(101, 34)
        Me.AC_Code.TabIndex = 178
        '
        'Rpt_Type
        '
        Me.Rpt_Type.FormattingEnabled = True
        Me.Rpt_Type.Items.AddRange(New Object() {"Dr-Cr", "Cr-Dr", "Dr", "Cr"})
        Me.Rpt_Type.Location = New System.Drawing.Point(698, 24)
        Me.Rpt_Type.Name = "Rpt_Type"
        Me.Rpt_Type.Size = New System.Drawing.Size(56, 32)
        Me.Rpt_Type.TabIndex = 177
        Me.Rpt_Type.Text = "Dr-Cr"
        '
        'RPT_ID
        '
        Me.RPT_ID.Location = New System.Drawing.Point(375, 23)
        Me.RPT_ID.Name = "RPT_ID"
        Me.RPT_ID.Size = New System.Drawing.Size(64, 34)
        Me.RPT_ID.TabIndex = 176
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.DG)
        Me.GroupBox1.Location = New System.Drawing.Point(10, 135)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(445, 399)
        Me.GroupBox1.TabIndex = 184
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "ເນື້ອນໃນລາຍການ"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.BtnMove)
        Me.GroupBox2.Controls.Add(Me.BtnSearch)
        Me.GroupBox2.Controls.Add(Me.DgItems)
        Me.GroupBox2.Location = New System.Drawing.Point(461, 135)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(587, 398)
        Me.GroupBox2.TabIndex = 185
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "ລາຍການບັນຊີ"
        '
        'BtnMove
        '
        Me.BtnMove.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnMove.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.BtnMove.Location = New System.Drawing.Point(4, 30)
        Me.BtnMove.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.BtnMove.Name = "BtnMove"
        Me.BtnMove.Size = New System.Drawing.Size(23, 25)
        Me.BtnMove.TabIndex = 307
        Me.BtnMove.Tag = "3011"
        Me.BtnMove.Text = "X"
        Me.BtnMove.UseVisualStyleBackColor = True
        Me.BtnMove.Visible = False
        '
        'BtnSearch
        '
        Me.BtnSearch.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSearch.Location = New System.Drawing.Point(110, 34)
        Me.BtnSearch.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.BtnSearch.Name = "BtnSearch"
        Me.BtnSearch.Size = New System.Drawing.Size(35, 25)
        Me.BtnSearch.TabIndex = 306
        Me.BtnSearch.Tag = "3012"
        Me.BtnSearch.Text = "....."
        Me.BtnSearch.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnSearch.UseVisualStyleBackColor = True
        Me.BtnSearch.Visible = False
        '
        'DgItems
        '
        Me.DgItems.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgItems.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgItems.Location = New System.Drawing.Point(3, 30)
        Me.DgItems.Name = "DgItems"
        Me.DgItems.Size = New System.Drawing.Size(581, 365)
        Me.DgItems.TabIndex = 0
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Button1)
        Me.GroupBox4.Controls.Add(Me.ComCurr)
        Me.GroupBox4.Controls.Add(Me.Label4)
        Me.GroupBox4.Controls.Add(Me.CAmt)
        Me.GroupBox4.Controls.Add(Me.COP)
        Me.GroupBox4.Controls.Add(Me.AC_Code)
        Me.GroupBox4.Controls.Add(Me.RPT_ID)
        Me.GroupBox4.Controls.Add(Me.Label1)
        Me.GroupBox4.Controls.Add(Me.Label3)
        Me.GroupBox4.Controls.Add(Me.Label2)
        Me.GroupBox4.Controls.Add(Me.Rpt_Type)
        Me.GroupBox4.Controls.Add(Me.Button3)
        Me.GroupBox4.Location = New System.Drawing.Point(9, 67)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(824, 65)
        Me.GroupBox4.TabIndex = 186
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "ເລືອກລາຍການ"
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(606, 22)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(35, 34)
        Me.Button1.TabIndex = 308
        Me.Button1.Tag = "3012"
        Me.Button1.Text = "....."
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ComCurr
        '
        Me.ComCurr.FormattingEnabled = True
        Me.ComCurr.Items.AddRange(New Object() {"ທຽບກີບ", "ເງິນເດີມ"})
        Me.ComCurr.Location = New System.Drawing.Point(247, 23)
        Me.ComCurr.Name = "ComCurr"
        Me.ComCurr.Size = New System.Drawing.Size(73, 32)
        Me.ComCurr.TabIndex = 186
        Me.ComCurr.Text = "All"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(169, 26)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(79, 24)
        Me.Label4.TabIndex = 185
        Me.Label4.Text = "ປະເພດເງິນ"
        '
        'CAmt
        '
        Me.CAmt.AutoSize = True
        Me.CAmt.Checked = True
        Me.CAmt.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CAmt.Location = New System.Drawing.Point(81, 26)
        Me.CAmt.Name = "CAmt"
        Me.CAmt.Size = New System.Drawing.Size(92, 28)
        Me.CAmt.TabIndex = 183
        Me.CAmt.Text = "ເຄື່ອນໄຫວ"
        Me.CAmt.UseVisualStyleBackColor = True
        '
        'COP
        '
        Me.COP.AutoSize = True
        Me.COP.Checked = True
        Me.COP.CheckState = System.Windows.Forms.CheckState.Checked
        Me.COP.Location = New System.Drawing.Point(8, 25)
        Me.COP.Name = "COP"
        Me.COP.Size = New System.Drawing.Size(77, 28)
        Me.COP.TabIndex = 184
        Me.COP.Text = "ຍອດຍົກ"
        Me.COP.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button6.Location = New System.Drawing.Point(196, 20)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(123, 49)
        Me.Button6.TabIndex = 183
        Me.Button6.Text = "ປັບປູງບັນຊີຍ່ອຍ"
        Me.Button6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button6.UseVisualStyleBackColor = True
        Me.Button6.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Saysettha OT", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(321, 28)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(392, 34)
        Me.Label5.TabIndex = 184
        Me.Label5.Text = "ລາຍລະອຽດການຄິດໄລ່ອົງປະກອບຂອງຊັບສິນ"
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(839, 18)
        Me.Button2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(35, 34)
        Me.Button2.TabIndex = 309
        Me.Button2.Tag = "3012"
        Me.Button2.Text = "....."
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Controls.Add(Me.BtnExit)
        Me.GroupBox3.Controls.Add(Me.BtnEdit)
        Me.GroupBox3.Controls.Add(Me.Button6)
        Me.GroupBox3.Location = New System.Drawing.Point(9, -10)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(824, 75)
        Me.GroupBox3.TabIndex = 1
        Me.GroupBox3.TabStop = False
        '
        'BtnExit
        '
        Me.BtnExit.BackgroundImage = Global.ApPBank10.My.Resources.Resources.Exit1
        Me.BtnExit.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnExit.Location = New System.Drawing.Point(6, 20)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(49, 49)
        Me.BtnExit.TabIndex = 174
        Me.BtnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'BtnEdit
        '
        Me.BtnEdit.Image = CType(resources.GetObject("BtnEdit.Image"), System.Drawing.Image)
        Me.BtnEdit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnEdit.Location = New System.Drawing.Point(55, 20)
        Me.BtnEdit.Name = "BtnEdit"
        Me.BtnEdit.Size = New System.Drawing.Size(141, 49)
        Me.BtnEdit.TabIndex = 175
        Me.BtnEdit.Text = "ແກ້ໄຂຂໍ້ມູນຫລັກ"
        Me.BtnEdit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnEdit.UseVisualStyleBackColor = True
        Me.BtnEdit.Visible = False
        '
        'FmRptProItem
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(221, Byte), Integer), CType(CType(254, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1050, 541)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox4)
        Me.Font = New System.Drawing.Font("Saysettha OT", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Margin = New System.Windows.Forms.Padding(4, 6, 4, 6)
        Me.Name = "FmRptProItem"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FmRptProItem"
        CType(Me.DG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.DgItems, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DG As System.Windows.Forms.DataGridView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents AC_Code As System.Windows.Forms.TextBox
    Friend WithEvents Rpt_Type As System.Windows.Forms.ComboBox
    Friend WithEvents RPT_ID As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents DgItems As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents ComCurr As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents CAmt As System.Windows.Forms.CheckBox
    Friend WithEvents COP As System.Windows.Forms.CheckBox
    Friend WithEvents BtnMove As System.Windows.Forms.Button
    Friend WithEvents BtnSearch As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents BtnEdit As System.Windows.Forms.Button
    Friend WithEvents BtnExit As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
End Class
