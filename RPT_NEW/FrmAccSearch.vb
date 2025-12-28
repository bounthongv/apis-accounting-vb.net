Public Class FrmAccSearch
    Private Sub LoadDG()
        Dim ds As New DataSet
        Try
            ConnectCL()
            SqlClient = "Select Ac_Code, Name_L  from Acc_Code where  Ac_Code like N'" & txtSearch.Text & "%'  Or Name_L like N'%" & txtSearch.Text & "%'  order by AC_CODE "
            LoadCN()
            da.Fill(ds, "Ac_Code")
            DG.DataSource = ds.Tables("Ac_Code")
            cn.Close()
            DG.Columns(0).HeaderText = "ລະຫັດບັນຊີ" : DG.Columns(0).Width = "100"
            DG.Columns(1).HeaderText = "ຊື່ບັນຊີ" : DG.Columns(1).Width = "500"
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub txtAmt_letter_E_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSearch.KeyPress
        If e.KeyChar = Chr(13) Then
            LoadDG()
        End If
    End Sub

    Private Sub DG_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DG.DoubleClick
        Close()
    End Sub

    Private Sub DG_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DG.KeyPress
        If e.KeyChar = Chr(13) Then
            'MsgBox("tyh")
        End If
    End Sub

    Private Sub FrmAccSearch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadDG()
    End Sub

    Private Sub DG_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DG.CellContentClick

    End Sub

    Private Sub DG_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DG.MouseClick
        MuAcCode = (DG.Item(0, DG.CurrentRow.Index).Value().ToString)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        LoadDG()
    End Sub

    Private Sub txtSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.TextChanged

    End Sub
End Class