Public Class FmLocKey

    Private Sub txtPassword_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPassword.KeyPress
        If e.KeyChar = Chr(13) Then
            If MPws = txtPassword.Text Then

                txtPassword.BackColor = Color.White
                Close()
            Else
                txtPassword.Clear()
                txtPassword.BackColor = Color.Red
            End If
        End If
    End Sub

    Private Sub txtPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPassword.TextChanged
        txtPassword.BackColor = Color.White
    End Sub

    Private Sub FmLocKey_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MULockey = False

    End Sub

    Private Sub FmLocKey_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MULockey = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If MessageBox.Show("ທ່ານຕ້ອງການອອກຈາກໂປຼແກຣມບໍ?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            FmLogin.Close()
        End If

    End Sub
End Class