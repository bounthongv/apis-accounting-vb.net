<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Fm_Image
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
        Me.Img_ID = New System.Windows.Forms.TextBox
        Me.ImgType = New System.Windows.Forms.TextBox
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.a123456789 = New System.Windows.Forms.PictureBox
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.a123456789, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Img_ID
        '
        Me.Img_ID.Location = New System.Drawing.Point(12, 215)
        Me.Img_ID.Name = "Img_ID"
        Me.Img_ID.ReadOnly = True
        Me.Img_ID.Size = New System.Drawing.Size(81, 20)
        Me.Img_ID.TabIndex = 267
        '
        'ImgType
        '
        Me.ImgType.Location = New System.Drawing.Point(116, 215)
        Me.ImgType.Name = "ImgType"
        Me.ImgType.ReadOnly = True
        Me.ImgType.Size = New System.Drawing.Size(81, 20)
        Me.ImgType.TabIndex = 268
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Location = New System.Drawing.Point(12, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(185, 206)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 266
        Me.PictureBox1.TabStop = False
        '
        'a123456789
        '
        Me.a123456789.Location = New System.Drawing.Point(12, 3)
        Me.a123456789.Name = "a123456789"
        Me.a123456789.Size = New System.Drawing.Size(185, 206)
        Me.a123456789.TabIndex = 270
        Me.a123456789.TabStop = False
        '
        'Fm_Image
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(116, 0)
        Me.Controls.Add(Me.a123456789)
        Me.Controls.Add(Me.ImgType)
        Me.Controls.Add(Me.Img_ID)
        Me.Controls.Add(Me.PictureBox1)
        Me.Name = "Fm_Image"
        Me.Text = "Fm_Image"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.a123456789, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Img_ID As System.Windows.Forms.TextBox
    Friend WithEvents ImgType As System.Windows.Forms.TextBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents a123456789 As System.Windows.Forms.PictureBox
End Class
