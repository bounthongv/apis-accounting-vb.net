Public Class Rate_settingb
    Public RSC As New ADODB.Recordset
    Public EditActive As Boolean
    Dim itemfgacc As Boolean
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim Dtsql As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub
    Private Sub Rate_setting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'If MDLanguage = 0 Then
        'Call LangLao()
        'Else
        '    Call Langs()
        'End If
        Call LoadText()
        Call LoadLIST()
        'If MWorkSetting = "" Then
        '    DTrate.Value = Date.Today
        'Else
        '    DTrate.Value = MWorkSetting
        'End If
        Call LoadSqlData("select * from AP_Rate_history", rs)
        'If MDWrite = 0 Then
        '    'BtnAddNew.Enabled = False
        'Else
        '    'BtnAddNew.Enabled = True
        'End If
        'If MDEdit = 0 Then
        '    'BtnSave.Enabled = False
        'Else
        '    'BtnSave.Enabled = True
        'End If
        'If MDDelete = 0 Then
        '    'BtnDel.Enabled = False
        'Else
        '    'BtnDel.Enabled = True
        'End If
        'If Mpermiss = "Admin" Then
        '    BtnDel.Enabled = True
        'Else
        '    BtnDel.Enabled = False

        'End If
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
        FG_Rate.FormatString = "NO|< Cerrent|<  ວັນທີ່(Date)   |<  THB-LAK  |< USD-THB  |< EUR-THB |< USD-LAK |< EUR-USD |< EUR-LAK  |< ຜູ້ໃຊ້(User)  |< Last update"

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
        FG_Rate.FormatString = "ລຝດ|<ສະກຸນເງິນ|<  ວັນທີ່(Date)   |<  ບາດ-ກີບ   |< ໂດລາ-ບາດ |< ຢູໂຣ-ບາດ |< ໂດລາ-ກີບ  |< ຢູໂຣ-ໂດລາ |< ຢູໂຣ-ກີບ  |< ຜູ້ໃຊ້(User)  |< Last update"

    End Sub
    Private Sub LoadLIST()
        FG_Rate.Rows = 1
        With rs
            Call LoadSqlData("select *  from AP_Rate_history WHERE curr<>'' " & Sql & " ORDER BY rate_dt DESC ", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    FG_Rate.AddItem(.AbsolutePosition & _
                    Chr(9) & (.Fields("curr").Value.ToString) & _
                    Chr(9) & Format(.Fields("rate_dt").Value, "dd/MM/yyyy") & _
                    Chr(9) & Format(CDbl(.Fields("THB_LAK").Value), "#,##0.00") & _
                    Chr(9) & Format(CDbl(.Fields("USD_THB").Value), "#,##0.00") & _
                    Chr(9) & Format(CDbl(.Fields("EUR_THB").Value), "#,##0.00") & _
                    Chr(9) & Format(CDbl(.Fields("USD_LAK").Value), "#,##0.00") & _
                    Chr(9) & Format(CDbl(.Fields("EUR_USD").Value), "#,##0.00") & _
                    Chr(9) & Format(CDbl(.Fields("EUR_LAK").Value), "#,##0.00") & _
                    Chr(9) & (.Fields("user_updt").Value.ToString) & _
                    Chr(9) & (.Fields("lst_updt").Value.ToString))
                    .MoveNext()
                End While
            End If
        End With
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
        On Error GoTo ll
        CNN.BeginTrans()

        CNN.Execute("update AP_Rate set THB_LAK=" & CDbl(txtTHB_LAK.Text) & ", USD_THB=" & CDbl(txtUSD_THB.Text) & ", EUR_THB=" & CDbl(txtEUR_THB.Text) & ", USD_LAK=" & CDbl(txtUSD_LAK.Text) & ", EUR_USD=" & CDbl(txtEUR_USD.Text) & ", EUR_LAK=" & CDbl(txtEUR_LAK.Text) & ", " & _
        "User_updt=N'" & Trim(MUserName) & "' ,PC_nm='" & Trim(MDServerName) & "', LAK=" & CDbl(txtEUR_LAK.Text) & ", THB=" & CDbl(txtEUR_THB.Text) & ", USD=" & CDbl(txtEUR_USD.Text) & ", EUR=1, lst_updt=getdate(), rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "' WHERE (status=1) AND (curr='EUR')")

        CNN.Execute("update AP_Rate set THB_LAK=" & CDbl(txtTHB_LAK.Text) & ", USD_THB=" & CDbl(txtUSD_THB.Text) & ", EUR_THB=" & CDbl(txtEUR_THB.Text) & ", USD_LAK=" & CDbl(txtUSD_LAK.Text) & ", EUR_USD=" & CDbl(txtEUR_USD.Text) & ", EUR_LAK=" & CDbl(txtEUR_LAK.Text) & ", " & _
        "User_updt=N'" & Trim(MUserName) & "' ,PC_nm='" & Trim(MDServerName) & "', LAK=" & CDbl(txtUSD_LAK.Text) & ", THB=" & CDbl(txtUSD_THB.Text) & ", USD=1, EUR=" & 1 / CDbl(txtEUR_USD.Text) & ", lst_updt=getdate(), rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "' WHERE (status=1) AND (curr='USD')")

        CNN.Execute("update AP_Rate set THB_LAK=" & CDbl(txtTHB_LAK.Text) & ", USD_THB=" & CDbl(txtUSD_THB.Text) & ", EUR_THB=" & CDbl(txtEUR_THB.Text) & ", USD_LAK=" & CDbl(txtUSD_LAK.Text) & ", EUR_USD=" & CDbl(txtEUR_USD.Text) & ", EUR_LAK=" & CDbl(txtEUR_LAK.Text) & ", " & _
        "User_updt=N'" & Trim(MUserName) & "' ,PC_nm='" & Trim(MDServerName) & "', LAK=" & CDbl(txtTHB_LAK.Text) & ", THB=1, USD=" & 1 / CDbl(txtUSD_THB.Text) & ", EUR=" & 1 / CDbl(txtEUR_THB.Text) & ", lst_updt=getdate(), rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "' WHERE (status=1) AND (curr='THB')")

        CNN.Execute("update AP_Rate set THB_LAK=" & CDbl(txtTHB_LAK.Text) & ", USD_THB=" & CDbl(txtUSD_THB.Text) & ", EUR_THB=" & CDbl(txtEUR_THB.Text) & ", USD_LAK=" & CDbl(txtUSD_LAK.Text) & ", EUR_USD=" & CDbl(txtEUR_USD.Text) & ", EUR_LAK=" & CDbl(txtEUR_LAK.Text) & ", " & _
        "User_updt=N'" & Trim(MUserName) & "' ,PC_nm='" & Trim(MDServerName) & "', LAK=1, THB=" & 1 / CDbl(txtTHB_LAK.Text) & ", USD=" & 1 / CDbl(txtUSD_LAK.Text) & ", EUR=" & 1 / CDbl(txtEUR_LAK.Text) & ", lst_updt=getdate(), rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "' WHERE (status=1) AND (curr='LAK')")

        Dim RsRate As New ADODB.Recordset
        With RsRate
            Call LoadSqlData("select curr from AP_Rate_history WHERE rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "'", RsRate)
            If .RecordCount = 0 Then
                CNN.Execute("insert into AP_Rate_history(Curr, Rate_dt, LAK, THB, USD, EUR, THB_LAK, USD_THB, EUR_THB, USD_LAK, EUR_USD, EUR_LAK, user_updt, Lst_updt, Pc_nm)" & _
                " select Curr,'" & Format(DTrate.Value, " yyyy-MM-dd") & "', LAK, THB, USD, EUR, " & CDbl(txtTHB_LAK.Text) & ", " & CDbl(txtUSD_THB.Text) & ", " & CDbl(txtEUR_THB.Text) & ", " & CDbl(txtUSD_LAK.Text) & ", " & CDbl(txtEUR_USD.Text) & ", " & CDbl(txtEUR_LAK.Text) & ", N'" & Trim(MUserName) & "', getdate(), '" & Trim(MDServerName) & "' from AP_Rate WHERE status=1")
            Else
                CNN.Execute("update A set A.curr=B.curr,A. rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "', A.LAK=B.LAK, A.THB=B.THB, A.USD=B.USD, A.EUR=B.EUR, " & _
                " A.THB_LAK=" & CDbl(txtTHB_LAK.Text) & ", A.USD_THB=" & CDbl(txtUSD_THB.Text) & ", A.EUR_THB=" & CDbl(txtEUR_THB.Text) & ", A.USD_LAK=" & CDbl(txtUSD_LAK.Text) & ", " & _
                " A.EUR_USD=" & CDbl(txtEUR_USD.Text) & ", A.EUR_LAK=" & CDbl(txtEUR_LAK.Text) & ", A.lst_updt=getdate(), User_updt=N'" & Trim(MUserName) & "', pc_nm='" & Trim(MDServerName) & "' from AP_Rate_history A, AP_Rate B WHERE (A.rate_dt=B.rate_dt) AND (B.status=1)")
            End If
        End With

        RsRate = Nothing
ll:
        If Err.Number = 0 Then
            CNN.CommitTrans()
            MsgBox("Save Complete", MsgBoxStyle.OkOnly)
        Else
            CNN.RollbackTrans()
            MsgBox(Err.Description)
        End If
        Call LoadLIST()
    End Sub

    Private Sub FG_Rate_DblClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles FG_Rate.DblClick

        txtTHB_LAK.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 3)), "#,##0.00")
        txtUSD_THB.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 4)), "#,##0.00")
        txtEUR_THB.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 5)), "#,##0.00")
        txtUSD_LAK.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 6)), "#,##0.00")
        txtEUR_USD.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 7)), "#,##0.00")
        txtEUR_LAK.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 8)), "#,##0.00")
        DTrate.Value = (FG_Rate.get_TextMatrix(FG_Rate.Row, 2))
    End Sub

    Private Sub FG_Rate_MouseUpEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseUpEvent) Handles FG_Rate.MouseUpEvent

        txtTHB_LAK.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 3)), "#,##0.00")
        txtUSD_THB.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 4)), "#,##0.00")
        txtEUR_THB.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 5)), "#,##0.00")
        txtUSD_LAK.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 6)), "#,##0.00")
        txtEUR_USD.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 7)), "#,##0.00")
        txtEUR_LAK.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 8)), "#,##0.00")
        ' DTrate.Value = Format(CDate(FG_Rate.get_TextMatrix(FG_Rate.Row, 2)), "dd/MM/yyyy")
        Dtsql = Format(CDate(FG_Rate.get_TextMatrix(FG_Rate.Row, 2)), "dd/MM/yyyy")
        DTrate.Value = Dtsql
    End Sub
    Private Sub FG_Rate_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG_Rate.SelChange
        txtTHB_LAK.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 3)), "#,##0.00")
        txtUSD_THB.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 4)), "#,##0.00")
        txtEUR_THB.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 5)), "#,##0.00")
        txtUSD_LAK.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 6)), "#,##0.00")
        txtEUR_USD.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 7)), "#,##0.00")
        txtEUR_LAK.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 8)), "#,##0.00")
        ' DTrate.Value = Format(CDate(FG_Rate.get_TextMatrix(FG_Rate.Row, 2)), "dd/MM/yyyy")
        Dtsql = Format(CDate(FG_Rate.get_TextMatrix(FG_Rate.Row, 2)), "dd/MM/yyyy")
        DTrate.Value = Dtsql
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDel.Click
        AccCD = FG_Rate.get_TextMatrix(FG_Rate.Row, 2)
        If FG_Rate.Rows = 2 Then MsgBox("You do not delete , because is list last  ", MsgBoxStyle.OkOnly) : Exit Sub
        'Call LoadData(" SELECT user_updt FROM AP_Rate_History WHERE (user_updt)= " & MUserName & " ", rs)
        'If rs.RecordCount <> 0 Then MsgBox("You do not delete , because is list last  ", MsgBoxStyle.OkOnly) : Exit Sub
        If MessageBox.Show("Do you want to delete " & AccCD & " yes or no ?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            CNN.Execute("DELETE FROM AP_Rate_History WHERE (rate_dt)= '" & Format(DTrate.Value, "yyyy-MM-dd") & "' ")
            If FG_Rate.Rows = 2 Then
                FG_Rate.Rows = 1
                FG_Rate.Rows = 2
            Else
                FG_Rate.RemoveItem(FG_Rate.Row)
            End If
        End If
        Call LoadLIST()
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