<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Conection_To_Servee
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
        Me.jhkh = New System.Windows.Forms.PictureBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtServerName = New System.Windows.Forms.TextBox
        Me.txtDatabaseName = New System.Windows.Forms.TextBox
        Me.txtServerUser = New System.Windows.Forms.TextBox
        Me.txtServerPassword = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btnConect = New System.Windows.Forms.Button
        Me.BtnCancel = New System.Windows.Forms.Button
        Me.cmdsearch = New System.Windows.Forms.Button
        Me.txtSaveIn = New System.Windows.Forms.TextBox
        Me.cmdrestore = New System.Windows.Forms.Button
        Me.RdBackUp = New System.Windows.Forms.RadioButton
        Me.RdRestor = New System.Windows.Forms.RadioButton
        Me.Label5 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.dtActi = New System.Windows.Forms.DateTimePicker
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        CType(Me.jhkh, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'jhkh
        '
        Me.jhkh.Location = New System.Drawing.Point(1, 1)
        Me.jhkh.Name = "jhkh"
        Me.jhkh.Size = New System.Drawing.Size(262, 28)
        Me.jhkh.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.jhkh.TabIndex = 34
        Me.jhkh.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtServerName)
        Me.GroupBox1.Controls.Add(Me.txtDatabaseName)
        Me.GroupBox1.Controls.Add(Me.txtServerUser)
        Me.GroupBox1.Controls.Add(Me.txtServerPassword)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Location = New System.Drawing.Point(20, 23)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(225, 126)
        Me.GroupBox1.TabIndex = 33
        Me.GroupBox1.TabStop = False
        '
        'txtServerName
        '
        Me.txtServerName.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtServerName.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtServerName.Location = New System.Drawing.Point(6, 24)
        Me.txtServerName.MaxLength = 20
        Me.txtServerName.Name = "txtServerName"
        Me.txtServerName.Size = New System.Drawing.Size(213, 22)
        Me.txtServerName.TabIndex = 11
        Me.txtServerName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtDatabaseName
        '
        Me.txtDatabaseName.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDatabaseName.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtDatabaseName.Location = New System.Drawing.Point(6, 61)
        Me.txtDatabaseName.MaxLength = 20
        Me.txtDatabaseName.Name = "txtDatabaseName"
        Me.txtDatabaseName.Size = New System.Drawing.Size(213, 22)
        Me.txtDatabaseName.TabIndex = 12
        Me.txtDatabaseName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtServerUser
        '
        Me.txtServerUser.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtServerUser.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtServerUser.Location = New System.Drawing.Point(5, 99)
        Me.txtServerUser.MaxLength = 15
        Me.txtServerUser.Name = "txtServerUser"
        Me.txtServerUser.Size = New System.Drawing.Size(105, 22)
        Me.txtServerUser.TabIndex = 13
        Me.txtServerUser.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtServerPassword
        '
        Me.txtServerPassword.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtServerPassword.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtServerPassword.Location = New System.Drawing.Point(114, 99)
        Me.txtServerPassword.MaxLength = 12
        Me.txtServerPassword.Name = "txtServerPassword"
        Me.txtServerPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtServerPassword.Size = New System.Drawing.Size(105, 22)
        Me.txtServerPassword.TabIndex = 14
        Me.txtServerPassword.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Silver
        Me.Label1.Location = New System.Drawing.Point(80, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 15)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Server Name"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Silver
        Me.Label4.Location = New System.Drawing.Point(138, 83)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(59, 15)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Password"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Silver
        Me.Label3.Location = New System.Drawing.Point(48, 84)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(33, 15)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "User"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Silver
        Me.Label2.Location = New System.Drawing.Point(73, 46)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(90, 15)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Database Name"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnConect)
        Me.GroupBox2.Controls.Add(Me.BtnCancel)
        Me.GroupBox2.Location = New System.Drawing.Point(20, 143)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(225, 32)
        Me.GroupBox2.TabIndex = 32
        Me.GroupBox2.TabStop = False
        '
        'btnConect
        '
        Me.btnConect.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.btnConect.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnConect.ForeColor = System.Drawing.Color.Navy
        Me.btnConect.Location = New System.Drawing.Point(3, 7)
        Me.btnConect.Name = "btnConect"
        Me.btnConect.Size = New System.Drawing.Size(107, 23)
        Me.btnConect.TabIndex = 4
        Me.btnConect.Text = "Conected"
        Me.btnConect.UseVisualStyleBackColor = False
        '
        'BtnCancel
        '
        Me.BtnCancel.BackColor = System.Drawing.Color.Silver
        Me.BtnCancel.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnCancel.ForeColor = System.Drawing.Color.Navy
        Me.BtnCancel.Location = New System.Drawing.Point(112, 7)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(107, 23)
        Me.BtnCancel.TabIndex = 5
        Me.BtnCancel.Text = "Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = False
        '
        'cmdsearch
        '
        Me.cmdsearch.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdsearch.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdsearch.ForeColor = System.Drawing.Color.Navy
        Me.cmdsearch.Location = New System.Drawing.Point(25, 269)
        Me.cmdsearch.Name = "cmdsearch"
        Me.cmdsearch.Size = New System.Drawing.Size(105, 23)
        Me.cmdsearch.TabIndex = 37
        Me.cmdsearch.Text = "Browser"
        Me.cmdsearch.UseVisualStyleBackColor = False
        '
        'txtSaveIn
        '
        Me.txtSaveIn.BackColor = System.Drawing.Color.White
        Me.txtSaveIn.ForeColor = System.Drawing.Color.Blue
        Me.txtSaveIn.Location = New System.Drawing.Point(25, 246)
        Me.txtSaveIn.Name = "txtSaveIn"
        Me.txtSaveIn.ReadOnly = True
        Me.txtSaveIn.Size = New System.Drawing.Size(214, 22)
        Me.txtSaveIn.TabIndex = 36
        '
        'cmdrestore
        '
        Me.cmdrestore.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdrestore.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdrestore.ForeColor = System.Drawing.Color.Navy
        Me.cmdrestore.Location = New System.Drawing.Point(134, 269)
        Me.cmdrestore.Name = "cmdrestore"
        Me.cmdrestore.Size = New System.Drawing.Size(105, 23)
        Me.cmdrestore.TabIndex = 39
        Me.cmdrestore.Text = "Ok"
        Me.cmdrestore.UseVisualStyleBackColor = False
        '
        'RdBackUp
        '
        Me.RdBackUp.AutoSize = True
        Me.RdBackUp.ForeColor = System.Drawing.Color.Silver
        Me.RdBackUp.Location = New System.Drawing.Point(147, 197)
        Me.RdBackUp.Name = "RdBackUp"
        Me.RdBackUp.Size = New System.Drawing.Size(95, 19)
        Me.RdBackUp.TabIndex = 6
        Me.RdBackUp.TabStop = True
        Me.RdBackUp.Text = "BackUp Data"
        Me.RdBackUp.UseVisualStyleBackColor = True
        '
        'RdRestor
        '
        Me.RdRestor.AutoSize = True
        Me.RdRestor.ForeColor = System.Drawing.Color.Silver
        Me.RdRestor.Location = New System.Drawing.Point(147, 219)
        Me.RdRestor.Name = "RdRestor"
        Me.RdRestor.Size = New System.Drawing.Size(88, 19)
        Me.RdRestor.TabIndex = 7
        Me.RdRestor.TabStop = True
        Me.RdRestor.Text = "Restor Data"
        Me.RdRestor.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Silver
        Me.Label5.Location = New System.Drawing.Point(26, 231)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(76, 15)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "File Address"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(25, 207)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(105, 22)
        Me.TextBox1.TabIndex = 40
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Silver
        Me.Label6.Location = New System.Drawing.Point(26, 193)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(62, 15)
        Me.Label6.TabIndex = 41
        Me.Label6.Text = "File Name"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(479, 6)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 43
        Me.Button1.Text = "Search"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'dtActi
        '
        Me.dtActi.CustomFormat = " dd / MM / yyyy"
        Me.dtActi.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtActi.Location = New System.Drawing.Point(269, 7)
        Me.dtActi.Name = "dtActi"
        Me.dtActi.Size = New System.Drawing.Size(99, 22)
        Me.dtActi.TabIndex = 205
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = " dd / MM / yyyy"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(374, 7)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(99, 22)
        Me.DateTimePicker1.TabIndex = 206
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox1.ForeColor = System.Drawing.Color.Red
        Me.CheckBox1.Location = New System.Drawing.Point(23, 177)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(189, 19)
        Me.CheckBox1.TabIndex = 208
        Me.CheckBox1.Text = "BackUp And Retor Data bases"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Conection_To_Servee
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(266, 314)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.dtActi)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.cmdrestore)
        Me.Controls.Add(Me.RdRestor)
        Me.Controls.Add(Me.cmdsearch)
        Me.Controls.Add(Me.txtSaveIn)
        Me.Controls.Add(Me.RdBackUp)
        Me.Controls.Add(Me.jhkh)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label6)
        Me.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "Conection_To_Servee"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        CType(Me.jhkh, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents jhkh As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtServerName As System.Windows.Forms.TextBox
    Friend WithEvents txtDatabaseName As System.Windows.Forms.TextBox
    Friend WithEvents txtServerUser As System.Windows.Forms.TextBox
    Friend WithEvents txtServerPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnConect As System.Windows.Forms.Button
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
    Friend WithEvents cmdrestore As System.Windows.Forms.Button
    Friend WithEvents cmdsearch As System.Windows.Forms.Button
    Friend WithEvents txtSaveIn As System.Windows.Forms.TextBox
    Friend WithEvents RdBackUp As System.Windows.Forms.RadioButton
    Friend WithEvents RdRestor As System.Windows.Forms.RadioButton
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents dtActi As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox

End Class
