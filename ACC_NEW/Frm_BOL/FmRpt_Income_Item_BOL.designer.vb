<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FmRpt_Income_Item_BOL
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FmRpt_Income_Item_BOL))
        Me.FG2 = New System.Windows.Forms.DataGridView()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.BtnMove = New System.Windows.Forms.Button()
        Me.BtnSearch = New System.Windows.Forms.Button()
        Me.FG = New System.Windows.Forms.DataGridView()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.BtnExit = New System.Windows.Forms.Button()
        Me.BtnEdit = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.AC_Code = New System.Windows.Forms.TextBox()
        Me.Rpt_Type = New System.Windows.Forms.ComboBox()
        Me.RPT_ID = New System.Windows.Forms.TextBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button6 = New System.Windows.Forms.Button()
        CType(Me.FG2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'FG2
        '
        Me.FG2.AllowUserToAddRows = False
        Me.FG2.AllowUserToDeleteRows = False
        Me.FG2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.FG2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FG2.Location = New System.Drawing.Point(435, 40)
        Me.FG2.Name = "FG2"
        Me.FG2.Size = New System.Drawing.Size(498, 378)
        Me.FG2.TabIndex = 12
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(277, -32)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(100, 30)
        Me.TextBox1.TabIndex = 13
        '
        'BtnMove
        '
        Me.BtnMove.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnMove.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.BtnMove.Location = New System.Drawing.Point(435, 65)
        Me.BtnMove.Name = "BtnMove"
        Me.BtnMove.Size = New System.Drawing.Size(37, 26)
        Me.BtnMove.TabIndex = 149
        Me.BtnMove.Text = "X"
        Me.BtnMove.UseVisualStyleBackColor = True
        '
        'BtnSearch
        '
        Me.BtnSearch.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSearch.Location = New System.Drawing.Point(530, 65)
        Me.BtnSearch.Name = "BtnSearch"
        Me.BtnSearch.Size = New System.Drawing.Size(34, 26)
        Me.BtnSearch.TabIndex = 148
        Me.BtnSearch.Text = "....."
        Me.BtnSearch.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnSearch.UseVisualStyleBackColor = True
        Me.BtnSearch.Visible = False
        '
        'FG
        '
        Me.FG.AllowUserToAddRows = False
        Me.FG.AllowUserToDeleteRows = False
        Me.FG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.FG.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.FG.Location = New System.Drawing.Point(5, 40)
        Me.FG.Name = "FG"
        Me.FG.Size = New System.Drawing.Size(424, 378)
        Me.FG.TabIndex = 150
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Button1.Location = New System.Drawing.Point(6, 64)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(37, 26)
        Me.Button1.TabIndex = 151
        Me.Button1.Text = "X"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'BtnExit
        '
        Me.BtnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnExit.Location = New System.Drawing.Point(5, 4)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(35, 35)
        Me.BtnExit.TabIndex = 152
        Me.BtnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'BtnEdit
        '
        Me.BtnEdit.Image = CType(resources.GetObject("BtnEdit.Image"), System.Drawing.Image)
        Me.BtnEdit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnEdit.Location = New System.Drawing.Point(40, 4)
        Me.BtnEdit.Name = "BtnEdit"
        Me.BtnEdit.Size = New System.Drawing.Size(134, 35)
        Me.BtnEdit.TabIndex = 153
        Me.BtnEdit.Text = "ແກ້ໄຂຂໍ້ມູນຫລັກ"
        Me.BtnEdit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnEdit.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(435, -42)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(187, 35)
        Me.Button2.TabIndex = 154
        Me.Button2.Text = "ແກ້ໄຂລະຫັດບັນຊີຜົນໄດ້ຮົບ"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button2.UseVisualStyleBackColor = True
        '
        'AC_Code
        '
        Me.AC_Code.Location = New System.Drawing.Point(645, 4)
        Me.AC_Code.Name = "AC_Code"
        Me.AC_Code.Size = New System.Drawing.Size(125, 30)
        Me.AC_Code.TabIndex = 163
        '
        'Rpt_Type
        '
        Me.Rpt_Type.FormattingEnabled = True
        Me.Rpt_Type.Items.AddRange(New Object() {"In", "Out"})
        Me.Rpt_Type.Location = New System.Drawing.Point(823, 5)
        Me.Rpt_Type.Name = "Rpt_Type"
        Me.Rpt_Type.Size = New System.Drawing.Size(56, 29)
        Me.Rpt_Type.TabIndex = 162
        Me.Rpt_Type.Text = "In"
        '
        'RPT_ID
        '
        Me.RPT_ID.Location = New System.Drawing.Point(486, 5)
        Me.RPT_ID.Name = "RPT_ID"
        Me.RPT_ID.Size = New System.Drawing.Size(91, 30)
        Me.RPT_ID.TabIndex = 161
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(886, 4)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(47, 30)
        Me.Button3.TabIndex = 164
        Me.Button3.Text = "DEL"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(776, 7)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 21)
        Me.Label3.TabIndex = 172
        Me.Label3.Text = "ປະເພດ"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(583, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 21)
        Me.Label2.TabIndex = 171
        Me.Label2.Text = "ເລກບັນຊີ"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(436, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(44, 21)
        Me.Label1.TabIndex = 170
        Me.Label1.Text = "ລະຫັດ"
        '
        'Button6
        '
        Me.Button6.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button6.Location = New System.Drawing.Point(294, 4)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(135, 35)
        Me.Button6.TabIndex = 173
        Me.Button6.Text = "ປັບປູງອັດຕະໂນມັດ"
        Me.Button6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button6.UseVisualStyleBackColor = True
        '
        'FmInCome
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 21.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(221, Byte), Integer), CType(CType(254, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(937, 430)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.AC_Code)
        Me.Controls.Add(Me.Rpt_Type)
        Me.Controls.Add(Me.RPT_ID)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.BtnEdit)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.FG)
        Me.Controls.Add(Me.BtnMove)
        Me.Controls.Add(Me.BtnSearch)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.FG2)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "FmInCome"
        Me.Text = "FmInCome"
        CType(Me.FG2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FG2 As System.Windows.Forms.DataGridView
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents BtnMove As System.Windows.Forms.Button
    Friend WithEvents BtnSearch As System.Windows.Forms.Button
    Friend WithEvents FG As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents BtnExit As System.Windows.Forms.Button
    Friend WithEvents BtnEdit As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents AC_Code As System.Windows.Forms.TextBox
    Friend WithEvents Rpt_Type As System.Windows.Forms.ComboBox
    Friend WithEvents RPT_ID As System.Windows.Forms.TextBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button6 As System.Windows.Forms.Button
End Class
