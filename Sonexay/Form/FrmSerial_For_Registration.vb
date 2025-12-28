Option Explicit On
Option Strict On
Public Class FrmSerial_For_Registration
    Dim rsProj As New ADODB.Recordset
    Public editProj As Boolean
    Dim Registrater As String
    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        If MDForMain <> "Bee" Then
            FmLogin.Close()
        End If
        If MDForMain = "Bee" Then
            Me.Close()
        End If
    End Sub
    Private Sub FrmSerial_For_Registration_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Conn As New ADODB.Connection
        ConnectAccess()
        txtKey1.Clear()
        txtkey2.Clear()
        txtkey3.Clear()
        txtkey4.Clear()
        Registrater = txtKey1.Text & "-" & txtkey2.Text & "-" & txtkey3.Text & "-" & txtkey4.Text
        txtForrever.Visible = False
    End Sub
    Private Sub SaveSerialUpdat()
        Dim rsProj As New ADODB.Recordset
        editProj = True
        Call LoadAcData("select SerialID FROM SerialUpdat Where SerialID='" & "001" & "'", rsProj)
        conn.Execute("UPDATE SerialUpdat SET Serial = '" & Registrater & "' , " & _
                       "StartData = '" & txtStartDate.Text.Trim & "' , " & _
                             "RightsByLaw= '" & RightsByLaws & "' , " & _
                                                       "SerialAge = '" & txtUsingDay.Text.Trim & "' WHERE SerialID = '" & "001" & "' ")
        MsgBox("Register Finish")

        If MDForMain <> "Bee" Then
            FmLogin.Visible = True
            Me.Close()
        End If
        If MDForMain = "Bee" Then
            Me.Close()
        End If
    End Sub
    Private Sub BtnRegistration_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRegistration.Click
        Call UpdateUsingDate()
        Call SaveSerialIn()
    End Sub
   
    Private Sub UpdateUsingDate()
        Dim rs As New ADODB.Recordset
        Call LoadAcData("select * from UsingDate where date_ID='" & "001" & "'", rs)
        conn.Execute("Update UsingDate Set Using_Date ='" & "0" & "' " & _
                                      " WHERE date_ID='" & "001" & "' ")
    End Sub
    Private Sub CheckSerialIn()
        Dim rs As New ADODB.Recordset
        Call LoadAcData("SELECT * FROM CheckSerialIn WHERE Serial = '" & Registrater & "'", rs)
        If rs.RecordCount > 0 Then
            PictureBox1.Visible = False
            BtnRegistration.Enabled = False
            txtStartDate.Text = ""
            txtUsingDay.Text = ""
        End If
        If rs.RecordCount = 0 Then
            PictureBox1.Visible = True
            BtnRegistration.Enabled = True
            txtStartDate.Text = DtmStartDate.Text
            txtUsingDay.Text = MDAgeRegistrtion
        End If
    End Sub
    Private Sub ChecSerialOut()
        If Registrater = "" Then
            PictureBox1.Visible = False
            BtnRegistration.Enabled = False
            txtStartDate.Text = ""
            txtUsingDay.Text = ""
            txtForrever.Visible = False
        End If
        Call LoadAcData("Select * from CheckSerialOut WHERE Serial='" & Registrater & "' ", rsProj)
        With rsProj
            If .RecordCount <> 0 Then
                MDAgeRegistrtion = (.Fields("Age").Value.ToString)
                RightsByLaws = (.Fields("RightsByLaw").Value.ToString)
                Call CheckSerialIn()
            End If
            If .RecordCount = 0 Then
                MDAgeRegistrtion = ""
                PictureBox1.Visible = False
                BtnRegistration.Enabled = False
                txtStartDate.Text = ""
                txtUsingDay.Text = ""
            End If
        End With

        'If RightsByLaws = "Forever" Then

        'End If
        txtForrever.Visible = False
        If RightsByLaws = "Forever" Then
            txtForrever.Visible = True
        End If
    End Sub
    Private Sub SaveSerialIn()
        Dim rs As New ADODB.Recordset
        Call LoadAcData("select * from CheckSerialIn where Serial='" & "001" & "'", rs)
        If rs.RecordCount <> 0 Then
        Else
            conn.Execute("Insert into CheckSerialIn(Serial,SerialID) " & _
                         " Values('" & Registrater & "', '" & "001" & "' )")
        End If
        Call SaveSerialUpdat()
    End Sub
    Private Sub txtRegistrater_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call ChecSerialOut()
    End Sub

    Private Sub txtKey1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtKey1.KeyPress
        If e.KeyChar = Chr(13) Then
            txtkey2.Focus()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKey1.TextChanged
        If Len(txtkey2.Text) = 4 Then txtkey2.Focus() : Exit Sub
        'Registrater = txtKey1.Text & "-" & txtkey2.Text & "-" & txtkey3.Text & "-" & txtkey4.Text
        'ChecSerialOut()

    End Sub

    Private Sub txtkey2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtkey2.KeyPress
        If e.KeyChar = Chr(13) Then
            txtkey3.Focus()
        End If
    End Sub

    Private Sub txtkey2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtkey2.TextChanged
        If Len(txtkey2.Text) = 4 Then txtkey3.Focus() : Exit Sub
        'Registrater = txtKey1.Text & "-" & txtkey2.Text & "-" & txtkey3.Text & "-" & txtkey4.Text
        'ChecSerialOut()
    End Sub

    Private Sub txtkey3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtkey3.KeyPress
        If e.KeyChar = Chr(13) Then
            txtkey4.Focus()
        End If
    End Sub

    Private Sub txtkey3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtkey3.TextChanged
        If Len(txtkey3.Text) = 4 Then txtkey4.Focus() : Exit Sub
        'Registrater = txtKey1.Text & "-" & txtkey2.Text & "-" & txtkey3.Text & "-" & txtkey4.Text
        'ChecSerialOut()
    End Sub

    Private Sub txtkey4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtkey4.KeyPress
        If e.KeyChar = Chr(13) Then
            txtKey1.Focus()
        End If
    End Sub

    Private Sub txtkey4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtkey4.TextChanged
        If Len(txtkey4.Text) = 4 Then Exit Sub
        'Registrater = txtKey1.Text & "-" & txtkey2.Text & "-" & txtkey3.Text & "-" & txtkey4.Text
        'ChecSerialOut
    End Sub

    Private Sub txtStartDate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStartDate.TextChanged

    End Sub

    Private Sub txtForrever_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtForrever.TextChanged

    End Sub
End Class