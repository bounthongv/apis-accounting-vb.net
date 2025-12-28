Public Class FmCaculate_Pro
    Dim adstr As String
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        LoadDG()
    End Sub
    Private Sub LoadDG()
        DG2.RowCount = 1
        TextBox3.Clear()
        Dim ds As New DataSet
        Try
            ConnectCL()
            SqlClient = "Select   RptID, Des , FML, Fnb,Und,Cor from So_Rpt_Pro  Where RptType = '" & ComboBox1.Text & "'  Order by RptId"
            LoadCN()
            da.Fill(ds, " So_Rpt_Pro")
            DG.DataSource = ds.Tables(" So_Rpt_Pro")
            cn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try
        DG.Columns(0).HeaderText = "ລະຫັດ " : DG.Columns(0).Width = "55"
        DG.Columns(1).HeaderText = "ເນື້ອໃນພາສາລາວ " : DG.Columns(1).Width = "380"
        DG.Columns(2).HeaderText = "ສູດຄິໄລ່ " : DG.Columns(2).Width = "154"
        DG.Columns(3).Visible = False
        DG.Columns(4).Visible = False
        DG.Columns(5).Visible = False
        For i = 0 To DG.RowCount - 1
            If DG.Item(3, i).Value() = 1 Then
                If DG.Item(4, i).Value() = 1 Then
                    DG.Item(0, i).Style.Font = New Font(DG.DefaultCellStyle.Font, FontStyle.Bold Or FontStyle.Underline)
                    DG.Item(1, i).Style.Font = New Font(DG.DefaultCellStyle.Font, FontStyle.Bold Or FontStyle.Underline)
                    DG.Item(2, i).Style.Font = New Font(DG.DefaultCellStyle.Font, FontStyle.Bold Or FontStyle.Underline)
                Else
                    DG.Item(0, i).Style.Font = New Font(DG.DefaultCellStyle.Font, FontStyle.Bold)
                    DG.Item(1, i).Style.Font = New Font(DG.DefaultCellStyle.Font, FontStyle.Bold)
                    DG.Item(2, i).Style.Font = New Font(DG.DefaultCellStyle.Font, FontStyle.Bold)
                End If
            Else
                If DG.Item(4, i).Value() = 1 Then
                    DG.Item(0, i).Style.Font = New Font(DG.DefaultCellStyle.Font, FontStyle.Underline)
                    DG.Item(1, i).Style.Font = New Font(DG.DefaultCellStyle.Font, FontStyle.Underline)
                    DG.Item(2, i).Style.Font = New Font(DG.DefaultCellStyle.Font, FontStyle.Underline)
                End If
            End If
            If DG.Item(5, i).Value() = 1 Then
                DG.Item(0, i).Style.ForeColor = Color.Red
                DG.Item(1, i).Style.ForeColor = Color.Red
                DG.Item(2, i).Style.ForeColor = Color.Red
            ElseIf DG.Item(5, i).Value() = 2 Then
                DG.Item(0, i).Style.ForeColor = Color.Blue
                DG.Item(1, i).Style.ForeColor = Color.Blue
                DG.Item(2, i).Style.ForeColor = Color.Blue
            End If
        Next

    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Close()
    End Sub
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        If DG.RowCount = 0 Then Exit Sub
        Dim n As Integer = DG2.Rows.Add()
        DG2.Rows.Item(n).Cells(0).Value = Button12.Text
        Showtable()
    End Sub
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        DG2.Rows.RemoveAt(DG2.RowCount - 2)
        Showtable()
    End Sub
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        If DG.RowCount = 0 Then Exit Sub
        Dim n As Integer = DG2.Rows.Add()
        DG2.Rows.Item(n).Cells(0).Value = Button13.Text
        Showtable()
    End Sub
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If DG.RowCount = 0 Then Exit Sub
        Dim n As Integer = DG2.Rows.Add()
        DG2.Rows.Item(n).Cells(0).Value = Button6.Text
        Showtable()
    End Sub
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        If DG.RowCount = 0 Then Exit Sub
        Dim n As Integer = DG2.Rows.Add()
        DG2.Rows.Item(n).Cells(0).Value = Button7.Text
        Showtable()
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If DG.RowCount = 0 Then Exit Sub
        Dim n As Integer = DG2.Rows.Add()
        DG2.Rows.Item(n).Cells(0).Value = Button4.Text
        Showtable()
    End Sub
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If DG.RowCount = 0 Then Exit Sub
        Dim n As Integer = DG2.Rows.Add()
        DG2.Rows.Item(n).Cells(0).Value = Button5.Text
        Showtable()
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If DG.RowCount = 0 Then Exit Sub
        Dim n As Integer = DG2.Rows.Add()
        DG2.Rows.Item(n).Cells(0).Value = Button3.Text
        Showtable()
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If DG.RowCount = 0 Then Exit Sub
        Dim n As Integer = DG2.Rows.Add()
        DG2.Rows.Item(n).Cells(0).Value = Button1.Text
        Showtable()
    End Sub
    Private Sub DG_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DG.MouseClick

        If CLock.Checked = False Then
            TextBox1.Text = (DG.Item(0, DG.CurrentRow.Index).Value().ToString)
            TextBox2.Text = (DG.Item(1, DG.CurrentRow.Index).Value().ToString)
            Exit Sub
        End If
        If DG.RowCount = 0 Then Exit Sub


        Dim ds As New DataSet
        Try
            ConnectCL()
            SqlClient = "select CLT_Str  from Caculate_Rpt where Rpt_Id = '" & (DG.Item(0, DG.CurrentRow.Index).Value().ToString) & "' And Rpt_Type = '" & ComboBox1.Text & "'  Order by  Rpt_id ,cnt  "
            LoadCN()
            da.Fill(ds, "Caculate_Rpt")
            If ds.Tables(0).Rows.Count > 0 Then
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    Dim n1 As Integer = DG2.Rows.Add()
                    DG2.Rows.Item(n1).Cells(0).Value = ds.Tables("Caculate_Rpt").Rows(i).Item("CLT_Str").ToString
                Next i
            Else
                Dim n2 As Integer = DG2.Rows.Add()
                DG2.Rows.Item(n2).Cells(0).Value = (DG.Item(0, DG.CurrentRow.Index).Value().ToString)
            End If
      
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Showtable()
    End Sub
    Private Sub Showtable()
        If CLock.Checked = True Then
            Dim c As Integer = DG2.RowCount - 2
            Dim s As String = ""
            For i = 0 To c
                s = s & (DG2.Item(0, i).Value().ToString)
            Next i
            TextBox3.Text = s
        Else
            DG2.RowCount = 1
            'DG2.DataSource = 0 : Exit Sub
        End If
  
    End Sub

    Private Sub DG_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DG.CellContentClick

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        DG2.RowCount = 1
        TextBox3.Clear()
    End Sub

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        For i = 0 To DG2.RowCount - 2
            SqlClient = "Insert Into Caculate_Rpt (Rpt_Id , CLT_Str , Rpt_Type ) Values ('" & TextBox1.Text & "' , '" & (DG2.Item(0, i).Value().ToString) & "' , '" & ComboBox1.Text & "') "
            CnnEdit()
        Next i
        Dim Und As Integer = 0
        If CheckBox1.Checked = True Then Und = 1
        SqlClient = "Update So_Rpt_Pro set StrCal = '" & TextBox3.Text & "' ,  Cor = '" & FontColor.SelectedIndex & "' ,  Fnb = '" & FontStype.SelectedIndex & "'   , Und = '" & Und & "' where RptID = '" & TextBox1.Text & "' and RptType = '" & ComboBox1.Text & "'  "
        CnnEdit()
        LoadDG()
    End Sub

    Private Sub FmCaculate_Pro_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
        FontColor.SelectedIndex = 0
        FontStype.SelectedIndex = 0

    End Sub
End Class