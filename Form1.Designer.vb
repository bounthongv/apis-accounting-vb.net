<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.Record = New System.Windows.Forms.TextBox
        Me.FirstRecord = New System.Windows.Forms.Button
        Me.BackRecord = New System.Windows.Forms.Button
        Me.LastRecord = New System.Windows.Forms.Button
        Me.NextRecord = New System.Windows.Forms.Button
        Me.Cust_ID = New System.Windows.Forms.TextBox
        Me.Nme = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Record
        '
        Me.Record.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Record.Location = New System.Drawing.Point(88, 39)
        Me.Record.Name = "Record"
        Me.Record.Size = New System.Drawing.Size(72, 20)
        Me.Record.TabIndex = 292
        Me.Record.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'FirstRecord
        '
        Me.FirstRecord.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.FirstRecord.Location = New System.Drawing.Point(17, 38)
        Me.FirstRecord.Name = "FirstRecord"
        Me.FirstRecord.Size = New System.Drawing.Size(36, 23)
        Me.FirstRecord.TabIndex = 291
        Me.FirstRecord.Text = "|<<"
        Me.FirstRecord.UseVisualStyleBackColor = True
        '
        'BackRecord
        '
        Me.BackRecord.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.BackRecord.Location = New System.Drawing.Point(52, 38)
        Me.BackRecord.Name = "BackRecord"
        Me.BackRecord.Size = New System.Drawing.Size(36, 23)
        Me.BackRecord.TabIndex = 290
        Me.BackRecord.Text = "<<"
        Me.BackRecord.UseVisualStyleBackColor = True
        '
        'LastRecord
        '
        Me.LastRecord.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.LastRecord.Location = New System.Drawing.Point(195, 38)
        Me.LastRecord.Name = "LastRecord"
        Me.LastRecord.Size = New System.Drawing.Size(36, 23)
        Me.LastRecord.TabIndex = 289
        Me.LastRecord.Text = ">>|"
        Me.LastRecord.UseVisualStyleBackColor = True
        '
        'NextRecord
        '
        Me.NextRecord.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.NextRecord.Location = New System.Drawing.Point(160, 38)
        Me.NextRecord.Name = "NextRecord"
        Me.NextRecord.Size = New System.Drawing.Size(36, 23)
        Me.NextRecord.TabIndex = 288
        Me.NextRecord.Text = ">>"
        Me.NextRecord.UseVisualStyleBackColor = True
        '
        'Cust_ID
        '
        Me.Cust_ID.Location = New System.Drawing.Point(17, 67)
        Me.Cust_ID.Name = "Cust_ID"
        Me.Cust_ID.Size = New System.Drawing.Size(214, 20)
        Me.Cust_ID.TabIndex = 293
        '
        'Nme
        '
        Me.Nme.Location = New System.Drawing.Point(17, 93)
        Me.Nme.Name = "Nme"
        Me.Nme.Size = New System.Drawing.Size(214, 20)
        Me.Nme.TabIndex = 294
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(238, 125)
        Me.Controls.Add(Me.Nme)
        Me.Controls.Add(Me.Cust_ID)
        Me.Controls.Add(Me.Record)
        Me.Controls.Add(Me.FirstRecord)
        Me.Controls.Add(Me.BackRecord)
        Me.Controls.Add(Me.LastRecord)
        Me.Controls.Add(Me.NextRecord)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Record As System.Windows.Forms.TextBox
    Friend WithEvents FirstRecord As System.Windows.Forms.Button
    Friend WithEvents BackRecord As System.Windows.Forms.Button
    Friend WithEvents LastRecord As System.Windows.Forms.Button
    Friend WithEvents NextRecord As System.Windows.Forms.Button
    Friend WithEvents Cust_ID As System.Windows.Forms.TextBox
    Friend WithEvents Nme As System.Windows.Forms.TextBox
End Class
