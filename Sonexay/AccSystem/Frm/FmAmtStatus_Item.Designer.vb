<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FmAmtStatus_Item
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FmAmtStatus_Item))
        Me.FG = New System.Windows.Forms.DataGridView
        Me.FG2 = New System.Windows.Forms.DataGridView
        Me.CRem = New System.Windows.Forms.CheckBox
        Me.CAmt = New System.Windows.Forms.CheckBox
        Me.COP = New System.Windows.Forms.CheckBox
        Me.Button5 = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.AC_Code = New System.Windows.Forms.TextBox
        Me.Rpt_Type = New System.Windows.Forms.ComboBox
        Me.RPT_ID = New System.Windows.Forms.TextBox
        Me.BtnEdit = New System.Windows.Forms.Button
        Me.BtnExit = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.BtnMove = New System.Windows.Forms.Button
        Me.BtnSearch = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        CType(Me.FG, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FG2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'FG
        '
        Me.FG.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.FG.DataSource = Nothing
        Me.FG.Location = New System.Drawing.Point(7, 46)
        Me.FG.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.FG.Name = "FG"
        Me.FG.OcxState = CType(resources.GetObject("FG.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FG.Size = New System.Drawing.Size(612, 440)
        Me.FG.TabIndex = 0
        '
        'FG2
        '
        Me.FG2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.FG2.DataSource = Nothing
        Me.FG2.Location = New System.Drawing.Point(626, 46)
        Me.FG2.Name = "FG2"
        Me.FG2.OcxState = CType(resources.GetObject("FG2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FG2.Size = New System.Drawing.Size(645, 440)
        Me.FG2.TabIndex = 1
        '
        'CRem
        '
        Me.CRem.AutoSize = True
        Me.CRem.Location = New System.Drawing.Point(894, 8)
        Me.CRem.Name = "CRem"
        Me.CRem.Size = New System.Drawing.Size(85, 25)
        Me.CRem.TabIndex = 189
        Me.CRem.Text = "ຍອດເຫລືອ"
        Me.CRem.UseVisualStyleBackColor = True
        '
        'CAmt
        '
        Me.CAmt.AutoSize = True
        Me.CAmt.Location = New System.Drawing.Point(818, 8)
        Me.CAmt.Name = "CAmt"
        Me.CAmt.Size = New System.Drawing.Size(82, 25)
        Me.CAmt.TabIndex = 188
        Me.CAmt.Text = "ເຄື່ອນໄຫວ"
        Me.CAmt.UseVisualStyleBackColor = True
        '
        'COP
        '
        Me.COP.AutoSize = True
        Me.COP.Location = New System.Drawing.Point(717, 8)
        Me.COP.Name = "COP"
        Me.COP.Size = New System.Drawing.Size(95, 25)
        Me.COP.TabIndex = 190
        Me.COP.Text = "ຍອດຍົກຕົ້ນປີ"
        Me.COP.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Enabled = False
        Me.Button5.Location = New System.Drawing.Point(155, 3)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(77, 35)
        Me.Button5.TabIndex = 187
        Me.Button5.Text = "ສ້າງສູດ"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(508, 11)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 21)
        Me.Label3.TabIndex = 186
        Me.Label3.Text = "ປະເພດ"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(346, 11)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 21)
        Me.Label2.TabIndex = 185
        Me.Label2.Text = "ເລກບັນຊີ"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(232, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(44, 21)
        Me.Label1.TabIndex = 184
        Me.Label1.Text = "ລະຫັດ"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(982, 8)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(47, 30)
        Me.Button3.TabIndex = 183
        Me.Button3.Text = "DEL"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'AC_Code
        '
        Me.AC_Code.Location = New System.Drawing.Point(411, 6)
        Me.AC_Code.Name = "AC_Code"
        Me.AC_Code.Size = New System.Drawing.Size(89, 30)
        Me.AC_Code.TabIndex = 182
        '
        'Rpt_Type
        '
        Me.Rpt_Type.FormattingEnabled = True
        Me.Rpt_Type.Items.AddRange(New Object() {"Dr-Cr", "Cr-Dr"})
        Me.Rpt_Type.Location = New System.Drawing.Point(562, 6)
        Me.Rpt_Type.Name = "Rpt_Type"
        Me.Rpt_Type.Size = New System.Drawing.Size(56, 29)
        Me.Rpt_Type.TabIndex = 181
        Me.Rpt_Type.Text = "Dr-Cr"
        '
        'RPT_ID
        '
        Me.RPT_ID.Location = New System.Drawing.Point(283, 6)
        Me.RPT_ID.Name = "RPT_ID"
        Me.RPT_ID.Size = New System.Drawing.Size(62, 30)
        Me.RPT_ID.TabIndex = 180
        '
        'BtnEdit
        '
        Me.BtnEdit.Enabled = False
        Me.BtnEdit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnEdit.Location = New System.Drawing.Point(40, 3)
        Me.BtnEdit.Name = "BtnEdit"
        Me.BtnEdit.Size = New System.Drawing.Size(116, 35)
        Me.BtnEdit.TabIndex = 179
        Me.BtnEdit.Text = "ແກ້ໄຂຂໍ້ມູນຫລັກ"
        Me.BtnEdit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnEdit.UseVisualStyleBackColor = True
        '
        'BtnExit
        '
        Me.BtnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnExit.Location = New System.Drawing.Point(5, 3)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(35, 35)
        Me.BtnExit.TabIndex = 178
        Me.BtnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Button1.Location = New System.Drawing.Point(7, 46)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(37, 26)
        Me.Button1.TabIndex = 177
        Me.Button1.Text = "X"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'BtnMove
        '
        Me.BtnMove.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnMove.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.BtnMove.Location = New System.Drawing.Point(626, 46)
        Me.BtnMove.Name = "BtnMove"
        Me.BtnMove.Size = New System.Drawing.Size(37, 26)
        Me.BtnMove.TabIndex = 176
        Me.BtnMove.Text = "X"
        Me.BtnMove.UseVisualStyleBackColor = True
        '
        'BtnSearch
        '
        Me.BtnSearch.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSearch.Location = New System.Drawing.Point(787, 46)
        Me.BtnSearch.Name = "BtnSearch"
        Me.BtnSearch.Size = New System.Drawing.Size(34, 26)
        Me.BtnSearch.TabIndex = 175
        Me.BtnSearch.Text = "....."
        Me.BtnSearch.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnSearch.UseVisualStyleBackColor = True
        Me.BtnSearch.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(622, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(18, 21)
        Me.Label4.TabIndex = 192
        Me.Label4.Text = "ປີ"
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"N-0", "N-1", "N-2"})
        Me.ComboBox1.Location = New System.Drawing.Point(655, 8)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(56, 29)
        Me.ComboBox1.TabIndex = 191
        Me.ComboBox1.Text = "N-0"
        '
        'FmAmtStatus_Item
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 21.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(221, Byte), Integer), CType(CType(254, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1280, 493)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.CRem)
        Me.Controls.Add(Me.CAmt)
        Me.Controls.Add(Me.COP)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.AC_Code)
        Me.Controls.Add(Me.Rpt_Type)
        Me.Controls.Add(Me.RPT_ID)
        Me.Controls.Add(Me.BtnEdit)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.BtnMove)
        Me.Controls.Add(Me.BtnSearch)
        Me.Controls.Add(Me.FG2)
        Me.Controls.Add(Me.FG)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "FmAmtStatus_Item"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FmAmtStatus_Item"
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FG2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FG As System.Windows.Forms.DataGridView
    Friend WithEvents FG2 As System.Windows.Forms.DataGridView
    Friend WithEvents CRem As System.Windows.Forms.CheckBox
    Friend WithEvents CAmt As System.Windows.Forms.CheckBox
    Friend WithEvents COP As System.Windows.Forms.CheckBox
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents AC_Code As System.Windows.Forms.TextBox
    Friend WithEvents Rpt_Type As System.Windows.Forms.ComboBox
    Friend WithEvents RPT_ID As System.Windows.Forms.TextBox
    Friend WithEvents BtnEdit As System.Windows.Forms.Button
    Friend WithEvents BtnExit As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents BtnMove As System.Windows.Forms.Button
    Friend WithEvents BtnSearch As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
End Class
