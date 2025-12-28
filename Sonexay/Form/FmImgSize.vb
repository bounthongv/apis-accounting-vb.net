Public Class FmImgSize

    Private Sub FmImgSixt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadSqlData("SELECT * FROM Ap_SizeImg ", RSC)
        With RSC
            Do Until .EOF = True
                TextBox1.Text = Trim(.Fields("b_x").Value)
                TextBox2.Text = Trim(.Fields("b_y").Value)
                TextBox3.Text = Trim(.Fields("g_x").Value)
                TextBox4.Text = Trim(.Fields("g_y").Value)
                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        CNN.Execute("update Ap_SizeImg set b_x = '" & TextBox1.Text & "' , b_y = '" & TextBox2.Text & "', g_x = '" & TextBox3.Text & "' ,g_y = '" & TextBox4.Text & "'")
        Close()
    End Sub

    Private Sub TextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.LostFocus

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If IsNumeric(TextBox1.Text) = False Then TextBox1.Text = 1
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        If IsNumeric(TextBox2.Text) = False Then TextBox2.Text = 2
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        If IsNumeric(TextBox4.Text) = False Then TextBox4.Text = 1
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        If IsNumeric(TextBox3.Text) = False Then TextBox3.Text = 1
    End Sub

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Close()
    End Sub
End Class