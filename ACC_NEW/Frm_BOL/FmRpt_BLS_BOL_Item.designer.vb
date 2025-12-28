<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FmRpt_BLS_BOL_Item
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FmRpt_BLS_BOL_Item))
        Me.FG2 = New AxVSFlex8U.AxVSFlexGrid
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.BtnMove = New System.Windows.Forms.Button
        Me.BtnSearch = New System.Windows.Forms.Button
        Me.FG = New AxVSFlex8U.AxVSFlexGrid
        Me.Button1 = New System.Windows.Forms.Button
        Me.AC_Code = New System.Windows.Forms.TextBox
        Me.Rpt_Type = New System.Windows.Forms.ComboBox
        Me.RPT_ID = New System.Windows.Forms.TextBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.BtnEdit = New System.Windows.Forms.Button
        Me.BtnExit = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel2 = New System.Windows.Forms.Panel
        CType(Me.FG2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'FG2
        '
        Me.FG2.DataSource = Nothing
        Me.FG2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FG2.Location = New System.Drawing.Point(0, 0)
        Me.FG2.Name = "FG2"
        Me.FG2.OcxState = CType(resources.GetObject("FG2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FG2.Size = New System.Drawing.Size(578, 403)
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
        Me.BtnMove.Location = New System.Drawing.Point(55, 27)
        Me.BtnMove.Name = "BtnMove"
        Me.BtnMove.Size = New System.Drawing.Size(37, 26)
        Me.BtnMove.TabIndex = 149
        Me.BtnMove.Text = "X"
        Me.BtnMove.UseVisualStyleBackColor = True
        '
        'BtnSearch
        '
        Me.BtnSearch.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSearch.Location = New System.Drawing.Point(150, 27)
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
        Me.FG.DataSource = Nothing
        Me.FG.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FG.Location = New System.Drawing.Point(0, 0)
        Me.FG.Name = "FG"
        Me.FG.OcxState = CType(resources.GetObject("FG.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FG.Size = New System.Drawing.Size(629, 403)
        Me.FG.TabIndex = 150
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Button1.Location = New System.Drawing.Point(7, 24)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(37, 26)
        Me.Button1.TabIndex = 151
        Me.Button1.Text = "X"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'AC_Code
        '
        Me.AC_Code.Location = New System.Drawing.Point(799, 8)
        Me.AC_Code.Name = "AC_Code"
        Me.AC_Code.Size = New System.Drawing.Size(127, 30)
        Me.AC_Code.TabIndex = 160
        '
        'Rpt_Type
        '
        Me.Rpt_Type.FormattingEnabled = True
        Me.Rpt_Type.Items.AddRange(New Object() {"In", "Out"})
        Me.Rpt_Type.Location = New System.Drawing.Point(986, 7)
        Me.Rpt_Type.Name = "Rpt_Type"
        Me.Rpt_Type.Size = New System.Drawing.Size(56, 29)
        Me.Rpt_Type.TabIndex = 159
        Me.Rpt_Type.Text = "In"
        '
        'RPT_ID
        '
        Me.RPT_ID.Location = New System.Drawing.Point(640, 8)
        Me.RPT_ID.Name = "RPT_ID"
        Me.RPT_ID.Size = New System.Drawing.Size(91, 30)
        Me.RPT_ID.TabIndex = 158
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(1048, 6)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(47, 30)
        Me.Button3.TabIndex = 165
        Me.Button3.Text = "DEL"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(294, -29)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(75, 23)
        Me.Button4.TabIndex = 166
        Me.Button4.Text = "Button4"
        Me.Button4.UseVisualStyleBackColor = True
        Me.Button4.Visible = False
        '
        'Button2
        '
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(435, -39)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(213, 35)
        Me.Button2.TabIndex = 154
        Me.Button2.Text = "ແກ້ໄຂລະຫັດບັນຊີຂອງສັບສົມບັດ"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button2.UseVisualStyleBackColor = True
        '
        'BtnEdit
        '
        Me.BtnEdit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnEdit.Location = New System.Drawing.Point(40, 4)
        Me.BtnEdit.Name = "BtnEdit"
        Me.BtnEdit.Size = New System.Drawing.Size(135, 35)
        Me.BtnEdit.TabIndex = 153
        Me.BtnEdit.Text = "ແກ້ໄຂຂໍ້ມູນຫລັກ"
        Me.BtnEdit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnEdit.UseVisualStyleBackColor = True
        '
        'BtnExit
        '
        Me.BtnExit.Image = CType(resources.GetObject("BtnExit.Image"), System.Drawing.Image)
        Me.BtnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnExit.Location = New System.Drawing.Point(5, 4)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(35, 35)
        Me.BtnExit.TabIndex = 152
        Me.BtnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(590, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(44, 21)
        Me.Label1.TabIndex = 167
        Me.Label1.Text = "ລະຫັດ"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(737, 14)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 21)
        Me.Label2.TabIndex = 168
        Me.Label2.Text = "ເລກບັນຊີ"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(932, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 21)
        Me.Label3.TabIndex = 169
        Me.Label3.Text = "ປະເພດ"
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(174, 4)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(77, 35)
        Me.Button5.TabIndex = 170
        Me.Button5.Text = "ສ້າງສູດ"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button6.Location = New System.Drawing.Point(294, 4)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(135, 35)
        Me.Button6.TabIndex = 171
        Me.Button6.Text = "ປັບປູງອັດຕະໂນມັດ"
        Me.Button6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.FG)
        Me.Panel1.Location = New System.Drawing.Point(5, 45)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(629, 403)
        Me.Panel1.TabIndex = 172
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.Controls.Add(Me.BtnMove)
        Me.Panel2.Controls.Add(Me.BtnSearch)
        Me.Panel2.Controls.Add(Me.FG2)
        Me.Panel2.Location = New System.Drawing.Point(640, 45)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(578, 403)
        Me.Panel2.TabIndex = 173
        '
        'FmRpt_BLS_BOL_Item
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 21.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(221, Byte), Integer), CType(CType(254, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1230, 450)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.AC_Code)
        Me.Controls.Add(Me.Rpt_Type)
        Me.Controls.Add(Me.RPT_ID)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.BtnEdit)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.TextBox1)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "FmRpt_BLS_BOL_Item"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FmBLS"
        CType(Me.FG2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FG2 As AxVSFlex8U.AxVSFlexGrid
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents BtnMove As System.Windows.Forms.Button
    Friend WithEvents BtnSearch As System.Windows.Forms.Button
    Friend WithEvents FG As AxVSFlex8U.AxVSFlexGrid
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents BtnExit As System.Windows.Forms.Button
    Friend WithEvents BtnEdit As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents AC_Code As System.Windows.Forms.TextBox
    Friend WithEvents Rpt_Type As System.Windows.Forms.ComboBox
    Friend WithEvents RPT_ID As System.Windows.Forms.TextBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
End Class
