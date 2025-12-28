Public Class FmLock
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If MPermit = "Admin" Then
            Call UpdateRemoveExS()
            LoadScExS()
        Else
            FmLogin.Close()
        End If
     
        'Close()
    End Sub

    Private Sub FmLock_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
     

    End Sub

    Private Sub FmLock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If MPermit = "Admin" Then
            Button1.Text = "ປົດລ໋ອກ"
            Button3.Enabled = True
        Else
            Button1.Text = "ອອກຈາກລະບົບ"
            Button3.Enabled = False
        End If

        If MULockey = True Then

            Button2.Enabled = False
        ElseIf MULockey = False Then

            Button2.Enabled = True
        End If
        Timer1.Enabled = True
        'ConnectSQL2()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        LoadScExS()
      
        If CloseAll = 0 Then
            Call UpdateRemoveExS()
            LoadScExS()
            FmMain.Timer1.Enabled = True
            Label1.Text = "ລະບົບນີ້ຖືກລ໋ອກແລ້ວກະລຸນນາຕິດຕໍ່ພວກເຮົາ"
            Me.Close()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        FmLocKey.ShowDialog()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Close()
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        FmMain.Button2.Visible = True
        Me.Hide()
        FmMain.Focus()
    End Sub
End Class