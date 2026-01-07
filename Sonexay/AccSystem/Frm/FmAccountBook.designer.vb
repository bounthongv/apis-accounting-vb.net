<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FmAccountBook
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FmAccountBook))
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.FG = New System.Windows.Forms.DataGridView
        Me.txtid = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtSty = New System.Windows.Forms.TextBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Cty = New System.Windows.Forms.ComboBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.BtnAddNew2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        CType(Me.FG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'FG
        '
        Me.FG.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.FG.BackgroundColor = System.Drawing.Color.White
        Me.FG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.FG.Location = New System.Drawing.Point(8, 138)
        Me.FG.Name = "FG"
        Me.FG.Size = New System.Drawing.Size(485, 253)
        Me.FG.TabIndex = 0
        Me.FG.Tag = "8005"
        '
        'txtid
        '
        Me.txtid.Location = New System.Drawing.Point(110, 39)
        Me.txtid.Name = "txtid"
        Me.txtid.Size = New System.Drawing.Size(84, 30)
        Me.txtid.TabIndex = 123
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(4, 42)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(107, 24)
        Me.Label1.TabIndex = 124
        Me.Label1.Tag = "2014"
        Me.Label1.Text = "ລະຫັດປື້ມ:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(110, 72)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(383, 30)
        Me.txtName.TabIndex = 125
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(4, 75)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(107, 24)
        Me.Label2.TabIndex = 126
        Me.Label2.Tag = "2052"
        Me.Label2.Text = "ຊື່ປື້ມ(ລາວ) :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtSty
        '
        Me.txtSty.Location = New System.Drawing.Point(248, -32)
        Me.txtSty.Name = "txtSty"
        Me.txtSty.Size = New System.Drawing.Size(100, 30)
        Me.txtSty.TabIndex = 127
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(110, 105)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(383, 30)
        Me.TextBox1.TabIndex = 125
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(27, 108)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(84, 24)
        Me.Label3.TabIndex = 126
        Me.Label3.Tag = "5053"
        Me.Label3.Text = "ຊື່ປື້ມ(ອັງກິດ) :"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(200, 44)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(67, 24)
        Me.Label4.TabIndex = 129
        Me.Label4.Tag = "2019"
        Me.Label4.Text = "ປະເພດປື້ມ:"
        '
        'Cty
        '
        Me.Cty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Cty.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cty.FormattingEnabled = True
        Me.Cty.Items.AddRange(New Object() {"ACC", "RCTY"})
        Me.Cty.Location = New System.Drawing.Point(266, 38)
        Me.Cty.Name = "Cty"
        Me.Cty.Size = New System.Drawing.Size(228, 32)
        Me.Cty.TabIndex = 274
        Me.Cty.Tag = "dfgdfg"
        '
        'Button1
        '
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(4, 1)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(35, 35)
        Me.Button1.TabIndex = 275
        Me.Button1.Tag = "9999"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(153, 1)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(114, 35)
        Me.Button2.TabIndex = 276
        Me.Button2.Tag = "3002"
        Me.Button2.Text = "ບັນທຶກ"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button2.UseVisualStyleBackColor = True
        '
        'BtnAddNew2
        '
        Me.BtnAddNew2.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAddNew2.Image = CType(resources.GetObject("BtnAddNew2.Image"), System.Drawing.Image)
        Me.BtnAddNew2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAddNew2.Location = New System.Drawing.Point(39, 1)
        Me.BtnAddNew2.Name = "BtnAddNew2"
        Me.BtnAddNew2.Size = New System.Drawing.Size(114, 35)
        Me.BtnAddNew2.TabIndex = 277
        Me.BtnAddNew2.Tag = "3001"
        Me.BtnAddNew2.Text = "ເພີ່ມໃຫມ່"
        Me.BtnAddNew2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnAddNew2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Image = CType(resources.GetObject("Button3.Image"), System.Drawing.Image)
        Me.Button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button3.Location = New System.Drawing.Point(266, 2)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(100, 35)
        Me.Button3.TabIndex = 278
        Me.Button3.Tag = "3004"
        Me.Button3.Text = "ລຶບ"
        Me.Button3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button3.UseVisualStyleBackColor = True
        '
        'FmAccountBook
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(500, 397)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.BtnAddNew2)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Cty)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtSty)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.txtid)
        Me.Controls.Add(Me.FG)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "FmAccountBook"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FmAccountBook"
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents FG As System.Windows.Forms.DataGridView
    Friend WithEvents txtid As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtSty As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Cty As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents BtnAddNew2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
End Class
