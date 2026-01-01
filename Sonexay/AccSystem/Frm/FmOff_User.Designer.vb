<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FmOff_User
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FmOff_User))
        Me.FG = New System.Windows.Forms.DataGridView
        Me.BtnExit = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.FG2 = New System.Windows.Forms.DataGridView
        Me.Label14 = New System.Windows.Forms.Label
        Me.CmbUsr = New System.Windows.Forms.ComboBox
        CType(Me.FG, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FG2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'FG
        '
        Me.FG.DataSource = Nothing
        Me.FG.Location = New System.Drawing.Point(10, 40)
        Me.FG.Name = "FG"
        Me.FG.Size = New System.Drawing.Size(350, 417)
        Me.FG.TabIndex = 0
        '
        'BtnExit
        '
        Me.BtnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnExit.Location = New System.Drawing.Point(7, 4)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(35, 35)
        Me.BtnExit.TabIndex = 292
        Me.BtnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(42, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(110, 34)
        Me.Button1.TabIndex = 291
        Me.Button1.Text = "ບັນທຶກ  "
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FG2
        '
        Me.FG2.DataSource = Nothing
        Me.FG2.Location = New System.Drawing.Point(372, 40)
        Me.FG2.Name = "FG2"
        Me.FG2.Size = New System.Drawing.Size(350, 417)
        Me.FG2.TabIndex = 293
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(286, 8)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(74, 21)
        Me.Label14.TabIndex = 295
        Me.Label14.Text = "ລະຫັດໃໝ່ :"
        '
        'CmbUsr
        '
        Me.CmbUsr.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbUsr.FormattingEnabled = True
        Me.CmbUsr.Location = New System.Drawing.Point(372, 6)
        Me.CmbUsr.Name = "CmbUsr"
        Me.CmbUsr.Size = New System.Drawing.Size(97, 30)
        Me.CmbUsr.TabIndex = 299
        '
        'FmOff_User
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 21.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(734, 467)
        Me.Controls.Add(Me.CmbUsr)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.FG2)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.FG)
        Me.Font = New System.Drawing.Font("Saysettha OT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximizeBox = False
        Me.Name = "FmOff_User"
        Me.Text = "FmOff_User"
        CType(Me.FG, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FG2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FG As System.Windows.Forms.DataGridView
    Friend WithEvents BtnExit As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents FG2 As System.Windows.Forms.DataGridView
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents CmbUsr As System.Windows.Forms.ComboBox
End Class
