Public Class FmAccountBook
    Private Sub SetupGrid()
        FG.AllowUserToAddRows = False
        FG.AllowUserToDeleteRows = False
        FG.ReadOnly = True
        FG.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG.MultiSelect = False
        FG.RowHeadersVisible = False
        
        FG.Columns.Clear()
        FG.Columns.Add("Col0", "ລ/ດ")
        FG.Columns.Add("Col1", "ລະຫັດປື້ມ´")
        FG.Columns.Add("Col2", "ຊື່ປື້ບັນຊີ")
        
        For Each col As DataGridViewColumn In FG.Columns
            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next
    End Sub

    Private Sub FmAccountBook_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FG.AllowUserToResizeColumns = True
        FG.AllowUserToResizeRows = True

        txtid.Enabled = True
        Cty.Text = "Acc"
        LoadListFG()
        SetupGrid()
        SetControlText(Me)
    End Sub
    Public Sub LoadListFG()
        FG.Rows.Clear()
        With RSC
            Call LoadSqlData("SELECT * FROM  Books where Type='" & Cty.Text & "' ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    Dim row As Integer = FG.Rows.Add()
                    FG.Rows(row).Cells(0).Value = .AbsolutePosition
                    FG.Rows(row).Cells(1).Value = Trim(CStr(.Fields("bookid").Value))
                    FG.Rows(row).Cells(2).Value = (.Fields("bookname").Value)
                    .MoveNext()
                End While
            Else
                ' DataGridView doesn't need to pre-allocate empty rows
            End If
        End With
    End Sub

   

    Private Sub FG_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FG.CellClick
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            txtid.Enabled = False
            txtid.Text = FG.Rows(e.RowIndex).Cells(1).Value?.ToString()
            txtName.Text = FG.Rows(e.RowIndex).Cells(2).Value?.ToString()
        End If
    End Sub
 

    Private Sub BntNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtid.Enabled = True
        txtid.Clear()
        txtName.Clear()
        txtid.Focus()
    End Sub

    Private Sub txtid_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtid.KeyPress
        If e.KeyChar = Chr(13) Then
            txtName.Focus()
        End If
    End Sub
    Private Sub txtName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtName.GotFocus
        'SetLao(1)
    End Sub

    Private Sub txtName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtName.KeyPress, TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call LoadSqlData("SELECT bookid FROM Books WHERE bookid = '" & Trim(txtid.Text) & "'", RSC)
            If RSC.RecordCount = 0 Then
                CNN.Execute("INSERT INTO Books( bookid,  bookname ) " & _
                    "Values('" & txtid.Text.Trim & "', N'" & txtName.Text.Trim & "')")
            Else
                CNN.Execute("UPDATE Books SET bookname='" & txtName.Text & "' WHERE bookid = '" & txtid.Text.Trim & "'")
            End If
            If RSC.State = ConnectionState.Open Then RSC.Close()
            MsgBox("ການບັນທຶກສຳເລັດ!", MsgBoxStyle.OkOnly)
            LoadListFG()
            txtid.Focus()
        End If
    End Sub
    Private Sub Cty_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cty.SelectedIndexChanged
        LoadListFG()
    End Sub
    Private Sub txtName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtName.LostFocus
        'SetLao(0)
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If txtid.Text = "" Then MsgBox("ກະລຸນາເລືອກກ່ອນ!", MsgBoxStyle.OkOnly) : Exit Sub
        If MessageBox.Show("ທ່ານຕ້ອງລຶບລະຫັດ " & FG.get_TextMatrix(FG.Row, 1) & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            CNN.Execute("DELETE FROM Books WHERE bookid='" & FG.get_TextMatrix(FG.Row, 1) & "'")
            txtid.Clear()
            txtName.Clear()
            LoadListFG()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call LoadSqlData("SELECT bookid FROM Books WHERE bookid = '" & Trim(txtid.Text) & "'", RSC)
        If RSC.RecordCount = 0 Then
            CNN.Execute("INSERT INTO Books( bookid, Type,  bookname , booknameE ) " & _
                "Values('" & txtid.Text.Trim & "', '" & Cty.Text & "' , N'" & txtName.Text.Trim & "' ,'" & TextBox1.Text & "')")
        Else
            CNN.Execute("UPDATE Books SET  Type=N'" & Cty.Text & "' ,  bookname=N'" & txtName.Text & "' ,  booknameE=N'" & TextBox1.Text & "'  WHERE bookid = '" & txtid.Text.Trim & "'")
        End If
        If RSC.State = ConnectionState.Open Then RSC.Close()
        MsgBox("ການບັນທຶກສຳເລັດ!", MsgBoxStyle.OkOnly)
        txtid.Focus()
        LoadListFG()
    End Sub

    Private Sub BtnAddNew2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew2.Click
        txtid.Enabled = True
        txtName.Text = ""
        TextBox1.Text = ""

    End Sub
End Class