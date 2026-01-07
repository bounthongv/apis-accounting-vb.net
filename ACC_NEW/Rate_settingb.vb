Public Class Rate_settingb
    ' Migrated from ADODB to ADO.NET - using DbHelper
    Public EditActive As Boolean
    Dim itemfgacc As Boolean
    Dim Sql As String
    Dim Dtsql As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub
    Private Sub Rate_setting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Initialize DataGridView columns
        SetupDataGridView()
        
        Call LoadText()
        Call LoadLIST()
    End Sub
    Public Sub Langs()
        BtnAddNew.Text = "AddNew"
        BtnSave.Text = "Save"
        BtnDel.Text = "Delete"
        Label1.Text = "Date:"
        Label2.Text = "THB-LAK"
        Label3.Text = "USD-THB"
        Label4.Text = "EUR-THB"
        p.Text = "USD-LAK"
        o.Text = "EUR-USD"
        Label7.Text = "EUR-LAK"
    End Sub
    
    Private Sub SetupDataGridView()
        ' Setup DataGridView columns equivalent to FormatString
        FG_Rate.Columns.Clear()
        FG_Rate.Columns.Add("Col0", "ລ/ດ")
        FG_Rate.Columns.Add("Col1", "ສະກຸນເງິນ")
        FG_Rate.Columns.Add("Col2", "ວັນທີ່(Date)")
        FG_Rate.Columns.Add("Col3", "THB-LAK")
        FG_Rate.Columns.Add("Col4", "USD-THB")
        FG_Rate.Columns.Add("Col5", "EUR-THB")
        FG_Rate.Columns.Add("Col6", "USD-LAK")
        FG_Rate.Columns.Add("Col7", "EUR-USD")
        FG_Rate.Columns.Add("Col8", "EUR-LAK")
        FG_Rate.Columns.Add("Col9", "ຜູ້ໃຊ້(User)")
        FG_Rate.Columns.Add("Col10", "Last update")
        
        ' Set column widths and alignment
        For i As Integer = 3 To 8
            If FG_Rate.Columns.Count > i Then
                FG_Rate.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                FG_Rate.Columns(i).DefaultCellStyle.Format = "#,##0.00"
            End If
        Next
        
        FG_Rate.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        FG_Rate.AllowUserToAddRows = False
        FG_Rate.AllowUserToDeleteRows = False
        FG_Rate.ReadOnly = True
    End Sub
    Public Sub LangLao()
        BtnAddNew.Text = "ເພີ່ມໃໝ່"
        BtnSave.Text = "ບັນທຶກ"
        BtnDel.Text = "ລຶບ"
        Label1.Text = "ວັນທີ່:"
        Label2.Text = "ບາດ-ກີບ"
        Label3.Text = "ໂດລາ-ບາດ"
        Label4.Text = "ຢູໂຣ-ບາດ"
        p.Text = "ໂດລາ-ກີບ"
        o.Text = "ຢູໂຣ-ໂດລາ"
        Label7.Text = "ຢູໂຣ-ກີບ"
    End Sub
    Private Sub LoadLIST()
        ' Clear existing rows
        FG_Rate.Rows.Clear()
        
        ' Modern database access using DbHelper
        Dim dt As DataTable = DbHelper.GetDataTable("select * from AP_Rate_history WHERE curr<>'' " & Sql & " ORDER BY rate_dt DESC")
        
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim rowIndex As Integer = FG_Rate.Rows.Add()
                FG_Rate.Rows(rowIndex).Cells("Col0").Value = rowIndex + 1
                FG_Rate.Rows(rowIndex).Cells("Col1").Value = GetStr(row("curr"))
                FG_Rate.Rows(rowIndex).Cells("Col2").Value = Format(CDate(row("rate_dt")), "dd/MM/yyyy")
                FG_Rate.Rows(rowIndex).Cells("Col3").Value = GetSafeDouble(row("THB_LAK"))
                FG_Rate.Rows(rowIndex).Cells("Col4").Value = GetSafeDouble(row("USD_THB"))
                FG_Rate.Rows(rowIndex).Cells("Col5").Value = GetSafeDouble(row("EUR_THB"))
                FG_Rate.Rows(rowIndex).Cells("Col6").Value = GetSafeDouble(row("USD_LAK"))
                FG_Rate.Rows(rowIndex).Cells("Col7").Value = GetSafeDouble(row("EUR_USD"))
                FG_Rate.Rows(rowIndex).Cells("Col8").Value = GetSafeDouble(row("EUR_LAK"))
                FG_Rate.Rows(rowIndex).Cells("Col9").Value = GetStr(row("user_updt"))
                FG_Rate.Rows(rowIndex).Cells("Col10").Value = GetStr(row("lst_updt"))
            Next
        End If
    End Sub
    Private Sub LoadText()
        txtCerrent.Text = ""
        txtTHB_LAK.Text = "1.00"
        txtUSD_THB.Text = "1.00"
        txtEUR_THB.Text = "1.00"
        txtUSD_LAK.Text = "1.00"
        txtEUR_USD.Text = "1.00"
        txtEUR_LAK.Text = "1.00"
        DTrate.Text = Date.Today

    End Sub
    
    ' Helper function to safely convert to Double
    Private Function GetSafeDouble(ByVal value As Object) As Double
        If value Is Nothing OrElse IsDBNull(value) Then
            Return 0.0
        End If
        Dim result As Double
        If Double.TryParse(value.ToString, result) Then
            Return result
        Else
            Return 0.0
        End If
    End Function
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew.Click
        Call LoadText()
        txtTHB_LAK.Focus()
        txtCerrent.Enabled = True
    End Sub
    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        'CloseFrom()
        'Apimage = True
        Me.Close()
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        Try
            ' Start transaction for atomic operations
            DbHelper.ExecuteNonQuery("BEGIN TRANSACTION")

            DbHelper.ExecuteNonQuery("update AP_Rate set THB_LAK=" & CDbl(txtTHB_LAK.Text) & ", USD_THB=" & CDbl(txtUSD_THB.Text) & ", EUR_THB=" & CDbl(txtEUR_THB.Text) & ", USD_LAK=" & CDbl(txtUSD_LAK.Text) & ", EUR_USD=" & CDbl(txtEUR_USD.Text) & ", EUR_LAK=" & CDbl(txtEUR_LAK.Text) & ", " & _
            "User_updt=N'" & Trim(MUserName) & "' ,PC_nm='" & Trim(MDServerName) & "', LAK=" & CDbl(txtEUR_LAK.Text) & ", THB=" & CDbl(txtEUR_THB.Text) & ", USD=" & CDbl(txtEUR_USD.Text) & ", EUR=1, lst_updt=getdate(), rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "' WHERE (status=1) AND (curr='EUR')")

            DbHelper.ExecuteNonQuery("update AP_Rate set THB_LAK=" & CDbl(txtTHB_LAK.Text) & ", USD_THB=" & CDbl(txtUSD_THB.Text) & ", EUR_THB=" & CDbl(txtEUR_THB.Text) & ", USD_LAK=" & CDbl(txtUSD_LAK.Text) & ", EUR_USD=" & CDbl(txtEUR_USD.Text) & ", EUR_LAK=" & CDbl(txtEUR_LAK.Text) & ", " & _
            "User_updt=N'" & Trim(MUserName) & "' ,PC_nm='" & Trim(MDServerName) & "', LAK=" & CDbl(txtUSD_LAK.Text) & ", THB=" & CDbl(txtUSD_THB.Text) & ", USD=1, EUR=" & 1 / CDbl(txtEUR_USD.Text) & ", lst_updt=getdate(), rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "' WHERE (status=1) AND (curr='USD')")

            DbHelper.ExecuteNonQuery("update AP_Rate set THB_LAK=" & CDbl(txtTHB_LAK.Text) & ", USD_THB=" & CDbl(txtUSD_THB.Text) & ", EUR_THB=" & CDbl(txtEUR_THB.Text) & ", USD_LAK=" & CDbl(txtUSD_LAK.Text) & ", EUR_USD=" & CDbl(txtEUR_USD.Text) & ", EUR_LAK=" & CDbl(txtEUR_LAK.Text) & ", " & _
            "User_updt=N'" & Trim(MUserName) & "' ,PC_nm='" & Trim(MDServerName) & "', LAK=" & CDbl(txtTHB_LAK.Text) & ", THB=1, USD=" & 1 / CDbl(txtUSD_THB.Text) & ", EUR=" & 1 / CDbl(txtEUR_THB.Text) & ", lst_updt=getdate(), rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "' WHERE (status=1) AND (curr='THB')")

            DbHelper.ExecuteNonQuery("update AP_Rate set THB_LAK=" & CDbl(txtTHB_LAK.Text) & ", USD_THB=" & CDbl(txtUSD_THB.Text) & ", EUR_THB=" & CDbl(txtEUR_THB.Text) & ", USD_LAK=" & CDbl(txtUSD_LAK.Text) & ", EUR_USD=" & CDbl(txtEUR_USD.Text) & ", EUR_LAK=" & CDbl(txtEUR_LAK.Text) & ", " & _
            "User_updt=N'" & Trim(MUserName) & "' ,PC_nm='" & Trim(MDServerName) & "', LAK=1, THB=" & 1 / CDbl(txtTHB_LAK.Text) & ", USD=" & 1 / CDbl(txtUSD_LAK.Text) & ", EUR=" & 1 / CDbl(txtEUR_LAK.Text) & ", lst_updt=getdate(), rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "' WHERE (status=1) AND (curr='LAK')")

            ' Modern database access using DbHelper
            Dim dtCheck As DataTable = DbHelper.GetDataTable("select curr from AP_Rate_history WHERE rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "'")

            If dtCheck.Rows.Count = 0 Then
                DbHelper.ExecuteNonQuery("insert into AP_Rate_history(Curr, Rate_dt, LAK, THB, USD, EUR, THB_LAK, USD_THB, EUR_THB, USD_LAK, EUR_USD, EUR_LAK, user_updt, Lst_updt, Pc_nm)" & _
                    " select Curr,'" & Format(DTrate.Value, " yyyy-MM-dd") & "', LAK, THB, USD, EUR, " & CDbl(txtTHB_LAK.Text) & ", " & CDbl(txtUSD_THB.Text) & ", " & CDbl(txtEUR_THB.Text) & ", " & CDbl(txtUSD_LAK.Text) & ", " & CDbl(txtEUR_USD.Text) & ", " & CDbl(txtEUR_LAK.Text) & ", N'" & Trim(MUserName) & "', getdate(), '" & Trim(MDServerName) & "' from AP_Rate WHERE status=1")
            Else
                DbHelper.ExecuteNonQuery("update A set A.curr=B.curr,A. rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "', A.LAK=B.LAK, A.THB=B.THB, A.USD=B.USD, A.EUR=B.EUR, " & _
                    " A.THB_LAK=" & CDbl(txtTHB_LAK.Text) & ", A.USD_THB=" & CDbl(txtUSD_THB.Text) & ", A.EUR_THB=" & CDbl(txtEUR_THB.Text) & ", A.USD_LAK=" & CDbl(txtUSD_LAK.Text) & ", " & _
                    " A.EUR_USD=" & CDbl(txtEUR_USD.Text) & ", A.EUR_LAK=" & CDbl(txtEUR_LAK.Text) & ", A.lst_updt=getdate(), User_updt=N'" & Trim(MUserName) & "', pc_nm='" & Trim(MDServerName) & "' from AP_Rate_history A, AP_Rate B WHERE (A.rate_dt=B.rate_dt) AND (B.status=1)")
            End If

            ' Commit transaction
            DbHelper.ExecuteNonQuery("COMMIT TRANSACTION")

            MsgBox("Save Complete", MsgBoxStyle.OkOnly)
            Call LoadLIST()

        Catch ex As Exception
            ' Rollback on error
            Try
                DbHelper.ExecuteNonQuery("ROLLBACK TRANSACTION")
            Catch rollbackEx As Exception
                ' Ignore rollback errors
            End Try
            MsgBox("Error saving data: " & ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub FG_Rate_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles FG_Rate.CellDoubleClick
        If e.RowIndex >= 0 Then
            txtTHB_LAK.Text = Format(GetSafeDouble(FG_Rate.Rows(e.RowIndex).Cells("Col3").Value), "#,##0.00")
            txtUSD_THB.Text = Format(GetSafeDouble(FG_Rate.Rows(e.RowIndex).Cells("Col4").Value), "#,##0.00")
            txtEUR_THB.Text = Format(GetSafeDouble(FG_Rate.Rows(e.RowIndex).Cells("Col5").Value), "#,##0.00")
            txtUSD_LAK.Text = Format(GetSafeDouble(FG_Rate.Rows(e.RowIndex).Cells("Col6").Value), "#,##0.00")
            txtEUR_USD.Text = Format(GetSafeDouble(FG_Rate.Rows(e.RowIndex).Cells("Col7").Value), "#,##0.00")
            txtEUR_LAK.Text = Format(GetSafeDouble(FG_Rate.Rows(e.RowIndex).Cells("Col8").Value), "#,##0.00")
            DTrate.Value = CDate(FG_Rate.Rows(e.RowIndex).Cells("Col2").Value)
        End If
    End Sub

    Private Sub FG_Rate_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles FG_Rate.CellClick
        If e.RowIndex >= 0 Then
            txtTHB_LAK.Text = Format(GetSafeDouble(FG_Rate.Rows(e.RowIndex).Cells("Col3").Value), "#,##0.00")
            txtUSD_THB.Text = Format(GetSafeDouble(FG_Rate.Rows(e.RowIndex).Cells("Col4").Value), "#,##0.00")
            txtEUR_THB.Text = Format(GetSafeDouble(FG_Rate.Rows(e.RowIndex).Cells("Col5").Value), "#,##0.00")
            txtUSD_LAK.Text = Format(GetSafeDouble(FG_Rate.Rows(e.RowIndex).Cells("Col6").Value), "#,##0.00")
            txtEUR_USD.Text = Format(GetSafeDouble(FG_Rate.Rows(e.RowIndex).Cells("Col7").Value), "#,##0.00")
            txtEUR_LAK.Text = Format(GetSafeDouble(FG_Rate.Rows(e.RowIndex).Cells("Col8").Value), "#,##0.00")
            Dtsql = Format(CDate(FG_Rate.Rows(e.RowIndex).Cells("Col2").Value), "dd/MM/yyyy")
        End If
    End Sub
    Private Sub FG_Rate_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles FG_Rate.SelectionChanged
        If FG_Rate.CurrentRow IsNot Nothing Then
            Dim rowIndex As Integer = FG_Rate.CurrentRow.Index
            txtTHB_LAK.Text = Format(GetSafeDouble(FG_Rate.Rows(rowIndex).Cells("Col3").Value), "#,##0.00")
            txtUSD_THB.Text = Format(GetSafeDouble(FG_Rate.Rows(rowIndex).Cells("Col4").Value), "#,##0.00")
            txtEUR_THB.Text = Format(GetSafeDouble(FG_Rate.Rows(rowIndex).Cells("Col5").Value), "#,##0.00")
            txtUSD_LAK.Text = Format(GetSafeDouble(FG_Rate.Rows(rowIndex).Cells("Col6").Value), "#,##0.00")
            txtEUR_USD.Text = Format(GetSafeDouble(FG_Rate.Rows(rowIndex).Cells("Col7").Value), "#,##0.00")
            txtEUR_LAK.Text = Format(GetSafeDouble(FG_Rate.Rows(rowIndex).Cells("Col8").Value), "#,##0.00")
            Dtsql = Format(CDate(FG_Rate.Rows(rowIndex).Cells("Col2").Value), "dd/MM/yyyy")
        End If
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDel.Click
        If FG_Rate.CurrentRow Is Nothing Then Exit Sub
        
        Dim rowIndex As Integer = FG_Rate.CurrentRow.Index
        Dim AccCD As String = GetStr(FG_Rate.Rows(rowIndex).Cells("Col2").Value)
        
        If FG_Rate.Rows.Count = 1 Then 
            MsgBox("You do not delete , because is list last  ", MsgBoxStyle.OkOnly) 
            Exit Sub
        End If
        
        If MessageBox.Show("Do you want to delete " & AccCD & " yes or no ?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            DbHelper.ExecuteNonQuery("DELETE FROM AP_Rate_history WHERE (rate_dt)= '" & Format(CDate(FG_Rate.Rows(rowIndex).Cells("Col2").Value), "yyyy-MM-dd") & "' ")
            Call LoadLIST()
        End If
    End Sub

    Private Sub txtTHB_LAK_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTHB_LAK.KeyPress
        'Select Case Asc(e.KeyChar)
        '    Case 48 To 57, 8
        '    Case Else
        '        e.Handled = True
        'End Select

        If e.KeyChar = Chr(13) Then
            txtUSD_THB.Focus()
        End If
    End Sub
    Private Sub txtTHB_LAK_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTHB_LAK.LostFocus
        txtTHB_LAK.Text = Format(CDbl(txtTHB_LAK.Text), "#,##0.00")
        If Trim(txtTHB_LAK.Text) = "" Or Trim(txtTHB_LAK.Text) = 0 Then txtTHB_LAK.Text = "1.00"
    End Sub
    Private Sub txtTHB_LAK_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTHB_LAK.TextChanged
        'txtTHB_LAK.Text = Format(CDbl(txtTHB_LAK.Text), "#,##0.00")
        If Trim(txtTHB_LAK.Text) = "" Or Trim(txtTHB_LAK.Text) = 0 Then txtTHB_LAK.Text = "1.00"
    End Sub

    Private Sub txtEUR_THB_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEUR_THB.KeyPress
        'Select Case Asc(e.KeyChar)
        '    Case 48 To 57, 8
        '    Case Else
        '        e.Handled = True
        'End Select
        If e.KeyChar = Chr(13) Then
            txtEUR_USD.Focus()
        End If
    End Sub
    Private Sub txtEUR_THB_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEUR_THB.LostFocus
        txtEUR_THB.Text = Format(CDbl(txtEUR_THB.Text), "#,##0.00")
        txtEUR_LAK.Text = Format(CDbl(txtTHB_LAK.Text) * CDbl(txtEUR_THB.Text), "#,##0.00")
        txtEUR_USD.Text = Format(CDbl(txtEUR_THB.Text) / CDbl(txtUSD_THB.Text), "#,##0.00")
    End Sub
    Private Sub txtEUR_THB_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEUR_THB.TextChanged
        If Trim(txtEUR_THB.Text) = "" Or Trim(txtEUR_THB.Text) = 0 Then txtEUR_THB.Text = "1.00"
    End Sub

    Private Sub txtUSD_LAK_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtUSD_LAK.KeyPress
        'Select Case Asc(e.KeyChar)
        '    Case 48 To 57, 8
        '    Case Else
        '        e.Handled = True
        'End Select
        If e.KeyChar = Chr(13) Then
            txtEUR_THB.Focus()
        End If
    End Sub
    Private Sub txtUSD_LAK_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUSD_LAK.LostFocus
        txtUSD_LAK.Text = Format(CDbl(txtUSD_LAK.Text), "#,##0.00")
    End Sub
    Private Sub txtUSD_LAK_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUSD_LAK.TextChanged
        If Trim(txtUSD_LAK.Text) = "" Or Trim(txtUSD_LAK.Text) = 0 Then txtUSD_LAK.Text = "1.00"
    End Sub

    Private Sub txtEUR_USD_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEUR_USD.KeyPress
        'Select Case Asc(e.KeyChar)
        '    Case 48 To 57, 8
        '    Case Else
        '        e.Handled = True
        'End Select
        If e.KeyChar = Chr(13) Then
            txtEUR_LAK.Focus()
        End If
    End Sub
    Private Sub txtEUR_USD_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEUR_USD.LostFocus
        txtEUR_USD.Text = Format(CDbl(txtEUR_USD.Text), "#,##0.00")
    End Sub
    Private Sub txtEUR_USD_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEUR_USD.TextChanged
        If Trim(txtEUR_USD.Text) = "" Or Trim(txtEUR_USD.Text) = 0 Then txtEUR_USD.Text = "1.00"
    End Sub

    Private Sub txtEUR_LAK_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEUR_LAK.KeyPress
        'Select Case Asc(e.KeyChar)
        '    Case 48 To 57, 8
        '    Case Else
        '        e.Handled = True
        'End Select
        If e.KeyChar = Chr(13) Then
            Call Button2_Click(sender, e)
        End If
    End Sub
    Private Sub txtEUR_LAK_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEUR_LAK.LostFocus
        txtEUR_LAK.Text = Format(CDbl(txtEUR_LAK.Text), "#,##0.00")
    End Sub
    Private Sub txtEUR_LAK_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEUR_LAK.TextChanged
        If Trim(txtEUR_LAK.Text) = "" Or Trim(txtEUR_LAK.Text) = 0 Then txtEUR_LAK.Text = "1.00"
    End Sub

    Private Sub txtUSD_THB_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtUSD_THB.KeyPress
        'Select Case Asc(e.KeyChar)
        '    Case 48 To 57, 8
        '    Case Else
        '        e.Handled = True
        'End Select
        If e.KeyChar = Chr(13) Then
            txtUSD_LAK.Focus()
        End If
    End Sub
    Private Sub txtUSD_THB_LostFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUSD_THB.LostFocus
        txtUSD_THB.Text = Format(CDbl(txtUSD_THB.Text), "#,##0.00")
    End Sub
    Private Sub txtUSD_THB_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUSD_THB.TextChanged
        If Trim(txtUSD_THB.Text) = "" Or Trim(txtUSD_THB.Text) = 0 Then txtUSD_THB.Text = "1.00"
    End Sub


End Class