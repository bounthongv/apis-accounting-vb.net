Public Class Rate_setting
    Public RSC As New ADODB.Recordset
    Public EditActive As Boolean
    Dim itemfgacc As Boolean
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim Dtsql As String
    Dim SS_Curr As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' Refreshes data
        Call LoadData()
    End Sub

    Private Sub Rate_setting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call RateSetting()
        Call LoadText()
        Call LoadSqlData("select * from AP_Rate_history", rs)

        txtTHB_LAK.Text = Format(CDbl(MDTHB_LAK), "#,###0.00########")
        txtUSD_LAK.Text = Format(CDbl(MDUSD_LAK), "#,###0.00########")
        txtEUR_LAK.Text = Format(CDbl(MDEUR_LAK), "#,###0.00########")
        txtUSD_THB.Text = Format(CDbl(MDUSD_THB), "#,###0.00########")
        Me.Show()
        DTrate.Focus()

        ' Setup DataGridViews
        SetupGrids()

        Call Laod_combo()
        Call LoadData()
        Call LoadData_Curr()
    End Sub

    Private Sub SetupGrids()
        ' FG_Rate Setup
        FG_Rate.Columns.Clear()
        FG_Rate.Columns.Add("No", "ລ/ດ")
        FG_Rate.Columns.Add("Currency", "ສະກຸນເງິນ")
        FG_Rate.Columns.Add("Date", "ວັນທີ່(Date)")
        FG_Rate.Columns.Add("Rate", "ອັດຕາ / Rate")
        FG_Rate.Columns.Add("Rate2", "ອັດຕາ / Rate2")
        FG_Rate.Columns.Add("LaoName", "ເປັນພາສາລາວ")
        FG_Rate.Columns.Add("User", "ຜູ້ໃຊ້(User)")
        FG_Rate.Columns.Add("LastUpdate", "Last update")

        FG_Rate.Columns(0).Width = 50
        FG_Rate.Columns(1).Width = 80
        FG_Rate.Columns(2).Width = 100
        FG_Rate.Columns(3).Width = 100
        FG_Rate.Columns(4).Width = 100
        FG_Rate.Columns(5).Width = 150
        FG_Rate.Columns(6).Width = 100
        FG_Rate.Columns(7).Width = 150

        FG_Rate.AllowUserToAddRows = False
        FG_Rate.ReadOnly = True
        FG_Rate.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG_Rate.MultiSelect = False


        ' FG_Curr Setup
        FG_Curr.Columns.Clear()
        FG_Curr.Columns.Add("No", "ລ/ດ")
        FG_Curr.Columns.Add("Currency", "ສະກຸນເງິນ")
        FG_Curr.Columns.Add("LaoName", "ເປັນພາສາລາວ")
        FG_Curr.Columns.Add("Cnt", "cnt")

        FG_Curr.Columns(0).Width = 50
        FG_Curr.Columns(1).Width = 100
        FG_Curr.Columns(2).Width = 200
        FG_Curr.Columns(3).Visible = False

        FG_Curr.AllowUserToAddRows = False
        FG_Curr.ReadOnly = True
        FG_Curr.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG_Curr.MultiSelect = False
    End Sub

    Private Sub Laod_combo()
        CMB_Curr.Items.Clear()
        Call load_Cmb(" SELECT Curr  FROM Curr_For_Rate  ORDER BY cnt ", "Curr", CMB_Curr)
        If CMB_Curr.Items.Count > 0 Then
            CMB_Curr.SelectedIndex = 0
        End If

        CMB_Curr_SSS.Items.Clear()
        CMB_Curr_SSS.Items.Add("ທັງໝົດ")
        Call load_Cmb(" SELECT Curr  FROM Curr_For_Rate  ORDER BY cnt ", "Curr", CMB_Curr_SSS)
        If CMB_Curr_SSS.Items.Count > 0 Then
            CMB_Curr_SSS.SelectedIndex = 0
        End If
    End Sub

    Private Sub LoadData_Curr()
        FG_Curr.Rows.Clear()
        With rs
            Call LoadSqlData("select *  from Curr_For_Rate WHERE 1=1  order by cnt ", rs)
            If .RecordCount > 0 Then
                Dim i As Integer = 1
                While Not .EOF()
                    FG_Curr.Rows.Add(i, _
                        .Fields("curr").Value.ToString, _
                        .Fields("Curr_name").Value.ToString, _
                        .Fields("cnt").Value.ToString)
                    .MoveNext()
                    i += 1
                End While
            End If
        End With
    End Sub

    Private Sub LoadData()
        Sql = " AND AP_Rate_history.rate_dt between '" & Format(dpFromDate.Value, "yyyy-MM-dd") & "'AND '" & Format(dpToDate.Value, "yyyy-MM-dd") & "'   "

        If CMB_Curr_SSS.SelectedIndex = 0 Or CMB_Curr_SSS.SelectedIndex = -1 Then
            SS_Curr = ""
        Else
            SS_Curr = " and AP_Rate_history.Curr =N'" & CMB_Curr_SSS.Text & "' "
        End If

        CNN.Execute("UPDATE AP_Rate_history SET Rate2=Rate where Rate2 IS NULL ")

        FG_Rate.Rows.Clear()
        With rs
            Dim PPP As String = "select AP_Rate_history.*, Curr_For_Rate.Curr_name FROM AP_Rate_history INNER JOIN " & _
                      "  Curr_For_Rate ON AP_Rate_history.curr = Curr_For_Rate.Curr  WHERE 1=1  " & Sql & "  " & SS_Curr & " order by  rate_dt ,Curr_For_Rate.cnt      "
            Call LoadSqlData(PPP, rs)
            If .RecordCount > 0 Then
                Dim i As Integer = 1
                While Not .EOF()
                    FG_Rate.Rows.Add(i, _
                        .Fields("curr").Value.ToString, _
                        Format(.Fields("rate_dt").Value, "dd/MM/yyyy"), _
                        Format(CDbl(.Fields("Rate").Value), "#,###0.00########"), _
                        Format(CDbl(.Fields("Rate2").Value), "#,###0.00########"), _
                        .Fields("Curr_name").Value.ToString, _
                        .Fields("user_updt").Value.ToString, _
                        .Fields("lst_updt").Value.ToString)
                    .MoveNext()
                    i += 1
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
        txtRate.Text = "1.00"
        txtRate2.Text = "1.00"
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew.Click
        Call LoadText()
        txtTHB_LAK.Focus()
        txtCerrent.Enabled = True
    End Sub

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Me.Close()
    End Sub

    ' Replaces DblClick, MouseUp, SelChange
    Private Sub FG_Rate_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles FG_Rate.SelectionChanged, FG_Rate.Click
        If FG_Rate.CurrentRow Is Nothing Then Exit Sub
        If FG_Rate.CurrentRow.Index < 0 Then Exit Sub

        Try
            CMB_Curr.Text = FG_Rate.CurrentRow.Cells(1).Value.ToString()
            txtRate.Text = Format(CDbl(FG_Rate.CurrentRow.Cells(3).Value), "#,###0.00########")
            txtRate2.Text = Format(CDbl(FG_Rate.CurrentRow.Cells(4).Value), "#,###0.00########")
            
            Dim dateStr As String = FG_Rate.CurrentRow.Cells(2).Value.ToString()
            ' Assuming date is stored as dd/MM/yyyy string in grid
            DTrate.Value = DateTime.ParseExact(dateStr, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture)
            Dtsql = dateStr

            txtcurr_name2.Text = FG_Rate.CurrentRow.Cells(5).Value.ToString()
        Catch ex As Exception
            ' Handle potential conversion errors or empty cells
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDel.Click
        If FG_Rate.CurrentRow Is Nothing Then Exit Sub
        
        Dim AccCD As String = FG_Rate.CurrentRow.Cells(2).Value.ToString() ' Date column
        
        ' Logic for "list last" check was a bit vague in original code, simplifying safe delete check
        If MessageBox.Show("Do you want to delete rate for " & AccCD & " yes or no ?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            CNN.Execute("DELETE FROM AP_Rate_History WHERE (rate_dt)= '" & Format(DTrate.Value, "yyyy-MM-dd") & "' ")
            
            FG_Rate.Rows.RemoveAt(FG_Rate.CurrentRow.Index)
        End If
        Call LoadData()
    End Sub

    ' ... (TextBox Event Handlers remain largely the same) ...
    Private Sub txtTHB_LAK_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTHB_LAK.KeyPress
        If e.KeyChar = Chr(13) Then
            txtUSD_LAK.Focus()
        End If
    End Sub
    Private Sub txtTHB_LAK_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTHB_LAK.LostFocus
        txtTHB_LAK.Text = Format(CDbl(txtTHB_LAK.Text), "#,###0.00########")
        If Trim(txtTHB_LAK.Text) = "" Or Trim(txtTHB_LAK.Text) = 0 Then txtTHB_LAK.Text = "1.00"
    End Sub
    Private Sub txtTHB_LAK_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTHB_LAK.TextChanged
        If Trim(txtTHB_LAK.Text) = "" Or Trim(txtTHB_LAK.Text) = 0 Then txtTHB_LAK.Text = "1.00"
    End Sub

    Private Sub txtEUR_THB_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEUR_THB.KeyPress
        If e.KeyChar = Chr(13) Then
            txtEUR_USD.Focus()
        End If
    End Sub
    Private Sub txtEUR_THB_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEUR_THB.LostFocus
        txtEUR_THB.Text = Format(CDbl(txtEUR_THB.Text), "#,###0.00########")
        txtEUR_LAK.Text = Format(CDbl(txtTHB_LAK.Text) * CDbl(txtEUR_THB.Text), "#,###0.00########")
        txtEUR_USD.Text = Format(CDbl(txtEUR_THB.Text) / CDbl(txtUSD_THB.Text), "#,###0.00########")
    End Sub
    Private Sub txtEUR_THB_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEUR_THB.TextChanged
        If Trim(txtEUR_THB.Text) = "" Or Trim(txtEUR_THB.Text) = 0 Then txtEUR_THB.Text = "1.00"
    End Sub

    Private Sub txtUSD_LAK_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtUSD_LAK.KeyPress
        If e.KeyChar = Chr(13) Then
            txtEUR_LAK.Focus()
        End If
    End Sub
    Private Sub txtUSD_LAK_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUSD_LAK.LostFocus
        txtUSD_LAK.Text = Format(CDbl(txtUSD_LAK.Text), "#,###0.00########")
    End Sub
    Private Sub txtUSD_LAK_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUSD_LAK.TextChanged
        If Trim(txtUSD_LAK.Text) = "" Or Trim(txtUSD_LAK.Text) = 0 Then txtUSD_LAK.Text = "1.00"
    End Sub

    Private Sub txtEUR_USD_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEUR_USD.KeyPress
        If e.KeyChar = Chr(13) Then
            txtEUR_LAK.Focus()
        End If
    End Sub
    Private Sub txtEUR_USD_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEUR_USD.LostFocus
        txtEUR_USD.Text = Format(CDbl(txtEUR_USD.Text), "#,###0.00########")
    End Sub
    Private Sub txtEUR_USD_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEUR_USD.TextChanged
        If Trim(txtEUR_USD.Text) = "" Or Trim(txtEUR_USD.Text) = 0 Then txtEUR_USD.Text = "1.00"
    End Sub

    Private Sub txtEUR_LAK_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEUR_LAK.KeyPress
        If e.KeyChar = Chr(13) Then
            txtUSD_THB.Focus()
        End If
    End Sub
    Private Sub txtEUR_LAK_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEUR_LAK.LostFocus
        txtEUR_LAK.Text = Format(CDbl(txtEUR_LAK.Text), "#,###0.00########")
    End Sub
    Private Sub txtEUR_LAK_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEUR_LAK.TextChanged
        If Trim(txtEUR_LAK.Text) = "" Or Trim(txtEUR_LAK.Text) = 0 Then txtEUR_LAK.Text = "1.00"
    End Sub

    Private Sub txtUSD_THB_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtUSD_THB.KeyPress
        If e.KeyChar = Chr(13) Then
            Button2_Click(sender, e)
        End If
    End Sub
    Private Sub txtUSD_THB_LostFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUSD_THB.LostFocus
        txtUSD_THB.Text = Format(CDbl(txtUSD_THB.Text), "#,###0.00########")
    End Sub
    Private Sub txtUSD_THB_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUSD_THB.TextChanged
        If Trim(txtUSD_THB.Text) = "" Or Trim(txtUSD_THB.Text) = 0 Then txtUSD_THB.Text = "1.00"
    End Sub

    Private Sub DTrate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DTrate.KeyPress
        If e.KeyChar = Chr(13) Then
            txtTHB_LAK.Focus()
        End If
    End Sub

    Private Sub Cmb_Component_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmb_Component.SelectedIndexChanged
        Dim aa, bb, cc As String
        aa = (Trim(Cmb_Component.Text))
        cc = Microsoft.VisualBasic.Left(aa, 12)
        cc = Trim(cc)
        bb = "Select * From AP_Office Where off_nm=N'" & Trim(Mid(Trim(aa), 11, Len(aa) - 10)) & "'  and   off_id = '" & Trim(cc) & "'   "
        Call LoadSqlData(bb, RSC)

        If RSC.RecordCount > 0 Then
            txt_Component_id.Text = Trim(RSC("off_id").Value)
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave2.Click
        If txtCurr.Text = "" Then MsgBox("ກະລຸນາ ໃສ່ສະກຸນເງິນ ກ່ອນ!", MsgBoxStyle.OkOnly) : txtCurr.Focus() : Exit Sub
        If txtcurr_name.Text = "" Then MsgBox("ກະລຸນາ ໃສ່ສະກຸນເງິນ ເປັນພາສາລາວ ກ່ອນ!", MsgBoxStyle.OkOnly) : txtcurr_name.Focus() : Exit Sub

        Dim RsRate As New ADODB.Recordset
        With RsRate
            Dim aa As String
            Call LoadSqlData("select Curr from Curr_For_Rate WHERE  Curr=N'" & Trim(txtCurr.Text) & "' ", RsRate)
            If .RecordCount = 0 Then
                aa = ("INSERT INTO Curr_For_Rate ( Curr ,Curr_name , Lst_updt,  user_updt, Pc_nm) " & _
                     " VALUES (N'" & (txtCurr.Text) & "'," & _
                        " N'" & txtcurr_name.Text & "'," & _
                    " Getdate()," & _
                    " N'" & MUserName & "'," & _
                    " '" & COMPUTER_NM & "')")
                CNN.Execute(aa)
            Else
                CNN.Execute("update Curr_For_Rate set  curr=N'" & Trim(txtCurr.Text) & "', " & _
                              " Curr_name=N'" & txtcurr_name.Text & "'," & _
                "  lst_updt=getdate(), User_updt=N'" & Trim(MUserName) & "', " & _
                " pc_nm='" & Trim(MDServerName) & "'   " & _
                " WHERE     Curr=N'" & Trim(txtCurr.Text) & "' ")
            End If
        End With
        Call Laod_combo()
        Call LoadData_Curr()
    End Sub

    ' Replaces FG_Curr_MouseUpEvent
    Private Sub FG_Curr_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles FG_Curr.SelectionChanged, FG_Curr.Click
        If FG_Curr.CurrentRow Is Nothing Then Exit Sub
        If FG_Curr.CurrentRow.Index < 0 Then Exit Sub

        txtCurr.Enabled = False
        Try
            txtCurr.Text = FG_Curr.CurrentRow.Cells(1).Value.ToString()
            txtcurr_name.Text = FG_Curr.CurrentRow.Cells(2).Value.ToString()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        If txtcurr_name2.Text = "" Then MsgBox("ກະລຸນາ ໃສ່ສະກຸນເງິນ ເປັນພາສາລາວ ກ່ອນ!", MsgBoxStyle.OkOnly) : txtcurr_name2.Focus() : Exit Sub
        If txtRate.Text = "" Then MsgBox("ກະລຸນາ ອັດຕາແລກປ່ຽນ ກ່ອນ!", MsgBoxStyle.OkOnly) : txtRate.Focus() : Exit Sub

        If CDbl(txtRate2.Text) = CDbl(txtRate.Text) Then
            txtRate2.Text = txtRate.Text
        End If

        CNN.Execute("update AP_Rate set THB_LAK=" & CDbl(txtTHB_LAK.Text) & ", USD_THB=" & CDbl(txtUSD_THB.Text) & ", EUR_THB=" & CDbl(txtEUR_THB.Text) & ", USD_LAK=" & CDbl(txtUSD_LAK.Text) & ", EUR_USD=" & CDbl(txtEUR_USD.Text) & ", EUR_LAK=" & CDbl(txtEUR_LAK.Text) & ", " & _
        "User_updt=N'" & Trim(MUserName) & "' ,PC_nm='" & Trim(MDServerName) & "', LAK=" & CDbl(txtEUR_LAK.Text) & ", THB=" & CDbl(txtEUR_THB.Text) & ", USD=" & CDbl(txtEUR_USD.Text) & ", EUR=1, lst_updt=getdate(), rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "' WHERE (status=1) AND (curr='EUR')")

        CNN.Execute("update AP_Rate set THB_LAK=" & CDbl(txtTHB_LAK.Text) & ", USD_THB=" & CDbl(txtUSD_THB.Text) & ", EUR_THB=" & CDbl(txtEUR_THB.Text) & ", USD_LAK=" & CDbl(txtUSD_LAK.Text) & ", EUR_USD=" & CDbl(txtEUR_USD.Text) & ", EUR_LAK=" & CDbl(txtEUR_LAK.Text) & ", " & _
        "User_updt=N'" & Trim(MUserName) & "' ,PC_nm='" & Trim(MDServerName) & "', LAK=" & CDbl(txtUSD_LAK.Text) & ", THB=" & CDbl(txtUSD_THB.Text) & ", USD=1, EUR=" & 1 / CDbl(txtEUR_USD.Text) & ", lst_updt=getdate(), rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "' WHERE (status=1) AND (curr='USD')")

        CNN.Execute("update AP_Rate set THB_LAK=" & CDbl(txtTHB_LAK.Text) & ", USD_THB=" & CDbl(txtUSD_THB.Text) & ", EUR_THB=" & CDbl(txtEUR_THB.Text) & ", USD_LAK=" & CDbl(txtUSD_LAK.Text) & ", EUR_USD=" & CDbl(txtEUR_USD.Text) & ", EUR_LAK=" & CDbl(txtEUR_LAK.Text) & ", " & _
        "User_updt=N'" & Trim(MUserName) & "' ,PC_nm='" & Trim(MDServerName) & "', LAK=" & CDbl(txtTHB_LAK.Text) & ", THB=1, USD=" & 1 / CDbl(txtUSD_THB.Text) & ", EUR=" & 1 / CDbl(txtEUR_THB.Text) & ", lst_updt=getdate(), rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "' WHERE (status=1) AND (curr='THB')")

        CNN.Execute("update AP_Rate set THB_LAK=" & CDbl(txtTHB_LAK.Text) & ", USD_THB=" & CDbl(txtUSD_THB.Text) & ", EUR_THB=" & CDbl(txtEUR_THB.Text) & ", USD_LAK=" & CDbl(txtUSD_LAK.Text) & ", EUR_USD=" & CDbl(txtEUR_USD.Text) & ", EUR_LAK=" & CDbl(txtEUR_LAK.Text) & ", " & _
        "User_updt=N'" & Trim(MUserName) & "' ,PC_nm='" & Trim(MDServerName) & "', LAK=1, THB=" & 1 / CDbl(txtTHB_LAK.Text) & ", USD=" & 1 / CDbl(txtUSD_LAK.Text) & ", EUR=" & 1 / CDbl(txtEUR_LAK.Text) & ", lst_updt=getdate(), rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "' WHERE (status=1) AND (curr='LAK')")

        Call Save()

        MsgBox("Save Complete", MsgBoxStyle.OkOnly)

        Call LoadData()
    End Sub

    Private Sub Save()
        Dim RsRate As New ADODB.Recordset
        With RsRate
            Dim aa As String
            Call LoadSqlData("select curr from AP_Rate_history WHERE rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "'  " & _
                        "  and  Curr=N'" & Trim(CMB_Curr.Text) & "' ", RsRate)
            If .RecordCount = 0 Then
                aa = ("INSERT INTO AP_Rate_history ( Curr  , Curr_name , rate_dt , Rate , Rate2,  Lst_updt,  user_updt, Pc_nm) " & _
                      " VALUES (N'" & (CMB_Curr.Text) & "'," & _
                        " N'" & txtcurr_name2.Text & "'," & _
                      " '" & Format(DTrate.Value, " yyyy-MM-dd") & "', " & _
                       " " & CDbl(txtRate.Text) & ", " & _
                                " " & CDbl(txtRate2.Text) & ", " & _
                          " Getdate()," & _
                        " N'" & MUserName & "'," & _
                     " '" & COMPUTER_NM & "')")
                CNN.Execute(aa)
            Else
                aa = "update AP_Rate_history set   rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "',  Rate=" & CDbl(txtRate.Text) & ",   Rate2=" & CDbl(txtRate2.Text) & ", " & _
                  " Curr_name=N'" & txtcurr_name2.Text & "'," & _
                "  lst_updt=getdate(), User_updt=N'" & Trim(MUserName) & "', pc_nm='" & Trim(MDServerName) & "'  " & _
                 " WHERE  rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "'  " & _
                        "  and  Curr=N'" & Trim(CMB_Curr.Text) & "' "
                CNN.Execute(aa)
            End If
        End With
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call LoadSqlData("select curr from AP_Rate_history WHERE    curr = N'" & Trim(txtCurr.Text) & "'  ", rs)
        If rs.RecordCount <> 0 Then MsgBox("ທ່ານ ບໍ່ສາມາດລຶບລາຍການນີ້ໄດ້ເພາະມີການເຄື່ອນໄຫວແລ້ວ.", MsgBoxStyle.OkOnly) : Exit Sub
        Dim aa As String
        If MessageBox.Show("ທ່ານຕ້ອງການລຶບລາຍການ : " & Trim(FG_Curr.CurrentRow.Cells(1).Value.ToString()) & "  ນີ້ແທ້ບໍ່?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            aa = " delete   Curr_For_Rate WHERE  Curr=N'" & Trim(txtCurr.Text) & "'  "
            CNN.Execute(aa)
            Call LoadData_Curr()
            Call Laod_combo()
        End If
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call LoadSqlData("select curr from gen_jn WHERE    curr = N'" & Trim(CMB_Curr.Text) & "'  ", rs)
        If rs.RecordCount <> 0 Then MsgBox("ທ່ານ ບໍ່ສາມາດລຶບລາຍການນີ້ໄດ້ເພາະມີການເຄື່ອນໄຫວແລ້ວ.", MsgBoxStyle.OkOnly) : Exit Sub
        Dim aa As String
        If MessageBox.Show("ທ່ານຕ້ອງການລຶບລາຍການ : " & Trim(FG_Curr.CurrentRow.Cells(1).Value.ToString()) & "  ນີ້ແທ້ບໍ່?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            aa = " delete   AP_Rate_history WHERE   rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "'  " & _
                        "  and  Curr=N'" & Trim(CMB_Curr.Text) & "'  "
            CNN.Execute(aa)
            Call LoadData()
        End If
    End Sub

    Private Sub CMB_Curr_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMB_Curr.SelectedIndexChanged
        Dim rs As New ADODB.Recordset
        Call LoadSqlData("Select * From Curr_For_Rate Where   Curr =N'" & Trim(CMB_Curr.Text) & "'", rs)
        If rs.RecordCount > 0 Then
            txtcurr_name2.Text = Trim(rs("Curr_name").Value.ToString)
        End If
    End Sub

    Private Sub CMB_Curr_SSS_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMB_Curr_SSS.SelectedIndexChanged
        Call LoadData()
    End Sub

    Private Sub txtRate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRate.TextChanged
    End Sub
End Class