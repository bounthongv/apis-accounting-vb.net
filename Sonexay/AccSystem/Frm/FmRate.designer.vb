<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FmRate
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FmRate))
        Me.BtnEdit2 = New System.Windows.Forms.Button
        Me.BtnSave = New System.Windows.Forms.Button
        Me.BtnAddNew2 = New System.Windows.Forms.Button
        Me.Curr = New System.Windows.Forms.TextBox
        Me.FG = New System.Windows.Forms.DataGridView
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Rate = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.CurrName = New System.Windows.Forms.TextBox
        Me.BtnDelete = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.FG2 = New System.Windows.Forms.DataGridView
        Me.Label43 = New System.Windows.Forms.Label
        Me.Curr_Last = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Rate2 = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.BtnExit = New System.Windows.Forms.Button
        CType(Me.FG, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FG2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BtnEdit2
        '
        Me.BtnEdit2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnEdit2.Location = New System.Drawing.Point(457, 3)
        Me.BtnEdit2.Name = "BtnEdit2"
        Me.BtnEdit2.Size = New System.Drawing.Size(134, 35)
        Me.BtnEdit2.TabIndex = 10
        Me.BtnEdit2.Text = "ແກ້ໄຂ"
        Me.BtnEdit2.UseVisualStyleBackColor = True
        '
        'BtnSave
        '
        Me.BtnSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnSave.Location = New System.Drawing.Point(220, 2)
        Me.BtnSave.Name = "BtnSave"
        Me.BtnSave.Size = New System.Drawing.Size(121, 35)
        Me.BtnSave.TabIndex = 9
        Me.BtnSave.Text = "ບັນທຶກ"
        Me.BtnSave.UseVisualStyleBackColor = True
        '
        'BtnAddNew2
        '
        Me.BtnAddNew2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAddNew2.Location = New System.Drawing.Point(100, 2)
        Me.BtnAddNew2.Name = "BtnAddNew2"
        Me.BtnAddNew2.Size = New System.Drawing.Size(121, 35)
        Me.BtnAddNew2.TabIndex = 8
        Me.BtnAddNew2.Text = "ເພີ່ມໃຫມ່"
        Me.BtnAddNew2.UseVisualStyleBackColor = True
        '
        'Curr
        '
        Me.Curr.Location = New System.Drawing.Point(100, 40)
        Me.Curr.Name = "Curr"
        Me.Curr.Size = New System.Drawing.Size(97, 30)
        Me.Curr.TabIndex = 12
        '
        'FG
        '
        Me.FG.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.FG.Location = New System.Drawing.Point(4, 105)
        Me.FG.Name = "FG"
        Me.FG.Size = New System.Drawing.Size(745, 423)
        Me.FG.TabIndex = 13
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 44)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(95, 24)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "ສະກຸນເງິນ(ຊື່ຫຍໍ້)"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(195, 46)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(99, 24)
        Me.Label2.TabIndex = 16
        Me.Label2.Text = "ສະກຸນເງິນ(ຊື່ເຕັມ)"
        '
        'Rate
        '
        Me.Rate.Location = New System.Drawing.Point(100, 72)
        Me.Rate.Name = "Rate"
        Me.Rate.Size = New System.Drawing.Size(97, 30)
        Me.Rate.TabIndex = 15
        Me.Rate.Text = "0.00"
        Me.Rate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(48, 75)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 24)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "ອັດຕາຊື້"
        '
        'CurrName
        '
        Me.CurrName.Location = New System.Drawing.Point(295, 41)
        Me.CurrName.Name = "CurrName"
        Me.CurrName.Size = New System.Drawing.Size(296, 30)
        Me.CurrName.TabIndex = 17
        '
        'BtnDelete
        '
        Me.BtnDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnDelete.Location = New System.Drawing.Point(340, 3)
        Me.BtnDelete.Name = "BtnDelete"
        Me.BtnDelete.Size = New System.Drawing.Size(118, 34)
        Me.BtnDelete.TabIndex = 19
        Me.BtnDelete.Text = "ລຶບ"
        Me.BtnDelete.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(384, 78)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(64, 24)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "ຈຳນວນໃບ"
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(446, 73)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(42, 30)
        Me.TextBox1.TabIndex = 22
        Me.TextBox1.Text = "0"
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'FG2
        '
        Me.FG2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FG2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.FG2.Location = New System.Drawing.Point(753, 105)
        Me.FG2.Name = "FG2"
        Me.FG2.Size = New System.Drawing.Size(180, 524)
        Me.FG2.TabIndex = 23
        '
        'Label43
        '
        Me.Label43.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label43.Font = New System.Drawing.Font("Saysettha OT", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label43.ForeColor = System.Drawing.Color.Blue
        Me.Label43.Location = New System.Drawing.Point(628, 21)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(288, 61)
        Me.Label43.TabIndex = 298
        Me.Label43.Text = "ລາຍການອັດຕາແລກປ່ຽນ"
        Me.Label43.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Curr_Last
        '
        Me.Curr_Last.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Curr_Last.Location = New System.Drawing.Point(544, 73)
        Me.Curr_Last.Name = "Curr_Last"
        Me.Curr_Last.Size = New System.Drawing.Size(47, 30)
        Me.Curr_Last.TabIndex = 300
        Me.Curr_Last.Text = "Kip"
        Me.Curr_Last.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(228, 77)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(66, 24)
        Me.Label6.TabIndex = 302
        Me.Label6.Text = "ອັດຕາຂາຍ"
        '
        'Rate2
        '
        Me.Rate2.Location = New System.Drawing.Point(295, 73)
        Me.Rate2.Name = "Rate2"
        Me.Rate2.Size = New System.Drawing.Size(91, 30)
        Me.Rate2.TabIndex = 301
        Me.Rate2.Text = "0.00"
        Me.Rate2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(488, 76)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(58, 24)
        Me.Label5.TabIndex = 299
        Me.Label5.Text = "ຊື່ທັງທ້າຍ"
        '
        'BtnExit
        '
        Me.BtnExit.Image = CType(resources.GetObject("BtnExit.Image"), System.Drawing.Image)
        Me.BtnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnExit.Location = New System.Drawing.Point(4, 2)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(35, 35)
        Me.BtnExit.TabIndex = 303
        Me.BtnExit.Tag = "9999"
        Me.BtnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'FmRate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(936, 533)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Rate2)
        Me.Controls.Add(Me.Curr_Last)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label43)
        Me.Controls.Add(Me.FG2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.BtnDelete)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.CurrName)
        Me.Controls.Add(Me.Rate)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.FG)
        Me.Controls.Add(Me.Curr)
        Me.Controls.Add(Me.BtnEdit2)
        Me.Controls.Add(Me.BtnSave)
        Me.Controls.Add(Me.BtnAddNew2)
        Me.Controls.Add(Me.Label2)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "FmRate"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FmRate"
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FG2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnEdit2 As System.Windows.Forms.Button
    Friend WithEvents BtnSave As System.Windows.Forms.Button
    Friend WithEvents BtnAddNew2 As System.Windows.Forms.Button
    Friend WithEvents Curr As System.Windows.Forms.TextBox
    Friend WithEvents FG As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Rate As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CurrName As System.Windows.Forms.TextBox
    Friend WithEvents BtnDelete As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents FG2 As System.Windows.Forms.DataGridView
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents Curr_Last As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Rate2 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents BtnExit As System.Windows.Forms.Button
End Class
