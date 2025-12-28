Public Class Rate_setting
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

        Call RateSetting()
        '===============================
        If Off_Id = "00" Then
            'Office_code = ""
            'Office_code = "  and  len(AP_Office.off_id)=6 or  len(AP_Office.off_id)=8  or len(AP_Office.off_id)=10 and  AP_Office.off_id <>'" & Off_Id & "'  "
        Else
            'If Len(Off_Id) = 6 Or Len(Off_Id) = 8 Or Len(Off_Id) = 10 Then
            '    Office_code = "  and   AP_Office.off_id='" & (Off_Id) & "'        "
            'Else
            '    Office_code = "  and  left(AP_Office.off_id,'" & Len(Off_Id) & "')>='" & Off_Id & "' and  left(AP_Office.off_id,'" & Len(Off_Id) & "')<='" & Off_Id & "'    and  len(AP_Office.off_id)>4       "
            'End If

            'Office_code = "  and  left(AP_Office.off_id,'" & Len(Off_Id) & "')>='" & Off_Id & "' and  left(AP_Office.off_id,'" & Len(Off_Id) & "')<='" & Off_Id & "'    "
            'Office_code = "  and   AP_Office.off_id='" & (Off_Id) & "'        "

        End If


        Dim rsc As New ADODB.Recordset
        Dim aa As String
        'If Off_Id = "00" Then
        '    aa = "select off_id,off_nm from AP_Office  where 1=1     order by off_id "
        '    Call LoadSqlData(aa, rsc)
        '    If rsc.RecordCount > 0 Then
        '        Cmb_Component.Items.Clear()
        '        While Not rsc.EOF()
        '            Cmb_Component.Items.Add(Trim(rsc.Fields("off_id").Value.ToString) + Space(11 - Len(Trim(rsc.Fields("off_id").Value.ToString))) & " " & (Trim(rsc.Fields("off_nm").Value.ToString)))
        '            rsc.MoveNext()
        '        End While
        '        Cmb_Component.SelectedIndex = 0
        '    End If

        'Else
        '    S_code = "  and  left(AP_Office.off_id,'" & Len(Off_Id) & "')>='" & Off_Id & "' and  left(AP_Office.off_id,'" & Len(Off_Id) & "')<='" & Off_Id & "'    "

        '    Call LoadSqlData("select off_id,off_nm from AP_Office  where 1=1    " & S_code & "        order by off_id ", rsc)
        '    If rsc.RecordCount > 0 Then
        '        Cmb_Component.Items.Clear()
        '        While Not rsc.EOF()
        '            Cmb_Component.Items.Add(Trim(rsc.Fields("off_id").Value.ToString) + Space(11 - Len(Trim(rsc.Fields("off_id").Value.ToString))) & " " & (Trim(rsc.Fields("off_nm").Value.ToString)))
        '            rsc.MoveNext()
        '        End While
        '        Cmb_Component.SelectedIndex = 0
        '    End If


        'End If




        'If MDLanguage = 0 Then
        '    Call LangLao()
        'Else
        '    Call Langs()
        'End If
        Call LoadText()

        'If MWorkSetting = "" Then
        '    DTrate.Value = Date.Today
        'Else
        '    DTrate.Value = MWorkSetting
        'End If
        Call LoadSqlData("select * from AP_Rate_history", rs)
        'If MDWrite = 0 Then
        '    BtnAddNew.Enabled = False
        'Else
        '    BtnAddNew.Enabled = True
        'End If
        'If MDEDIT = 0 Then
        '    BtnSave.Enabled = False
        '    BtnSave2.Enabled = False
        'Else
        '    BtnSave.Enabled = True
        '    BtnSave2.Enabled = True
        'End If
        'If MDDelete = 0 Then
        '    BtnDel.Enabled = False
        'Else
        '    BtnDel.Enabled = True
        'End If
        txtTHB_LAK.Text = Format(CDbl(MDTHB_LAK), "#,###0.00########")
        txtUSD_LAK.Text = Format(CDbl(MDUSD_LAK), "#,###0.00########")
        txtEUR_LAK.Text = Format(CDbl(MDEUR_LAK), "#,###0.00########")
        txtUSD_THB.Text = Format(CDbl(MDUSD_THB), "#,###0.00########")
        Me.Show()
        DTrate.Focus()

        FG_Rate.FormatString = "ລ/ດ|<ສະກຸນເງິນ|<  ວັນທີ່(Date)   |< ອັດຕາ / Rate  |< ອັດຕາ / Rate2  |< ເປັນພາສາລາວ|< ຜູ້ໃຊ້(User)  |< Last update"
        FG_Rate.set_ColHidden(1, False)
        FG_Curr.FormatString = "ລ/ດ|< ສະກຸນເງິນ  |< ເປັນພາສາລາວ   |<cnt "
        Call Laod_combo()
        Call LoadData()
        Call LoadData_Curr()
    End Sub
    Private Sub Laod_combo()
        CMB_Curr.Items.Clear()
        'CMBBK_ID.Items.Add("ທັງໝົດ")
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
        FG_Curr.Rows = 1
        With rs
            Call LoadSqlData("select *  from Curr_For_Rate WHERE 1=1  order by cnt ", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    FG_Curr.AddItem(.AbsolutePosition & _
                  Chr(9) & (.Fields("curr").Value.ToString) & _
                     Chr(9) & (.Fields("Curr_name").Value.ToString) & _
                     Chr(9) & (.Fields("cnt").Value.ToString))
                    'FG_Curr.AddItem(.AbsolutePosition & _
                    'Chr(9) & (.Fields("curr").Value.ToString) & _
                    'Chr(9) & (.Fields("user_updt").Value.ToString) & _
                    'Chr(9) & (.Fields("lst_updt").Value.ToString))
                    .MoveNext()
                End While
            End If
        End With
    End Sub

    Public Sub Langs()
        BtnAddNew.Text = "AddNew"
        BtnSave.Text = "Save"
        BtnDel.Text = "Delete"
        Label1.Text = "Date:"
        Label2.Text = "LAK-SDR"
        Label3.Text = "SDR-USD"
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
        Label1.Text = "ອັດຕາແລກປ່ຽນປະຈໍາວັນທີ່ :"
        Label2.Text = "ກີບ-SDR / KIP - SDR :"
        Label3.Text = "SDR-ໂດລາ/ SDR - DOLLAR :"
        'Label4.Text = "ຢູໂຣ-ບາດ"
        p.Text = "ກີບ-ໂດລາ / KIP - DOLLAR :"
        'o.Text = "ຢູໂຣ-ໂດລາ"
        'Label7.Text = "ກີບ-ຢູໂຣ / KIP - EURO :"
        FG_Rate.FormatString = "ລຝດ|<ສະກຸນເງິນ|<  ວັນທີ່(Date)   |<  ບາດ-ກີບ   |< ໂດລາ-ບາດ |< ຢູໂຣ-ບາດ |< ໂດລາ-ກີບ  |< ຢູໂຣ-ໂດລາ |< ຢູໂຣ-ກີບ  |< ຜູ້ໃຊ້(User)  |< Last update"

    End Sub
    Private Sub LoadData()
        Sql = " AND AP_Rate_history.rate_dt between '" & Format(dpFromDate.Value, "yyyy-MM-dd") & "'AND '" & Format(dpToDate.Value, "yyyy-MM-dd") & "'   "

        If CMB_Curr_SSS.SelectedIndex = 0 Then
            SS_Curr = ""
        Else
            SS_Curr = " and AP_Rate_history.Curr =N'" & CMB_Curr_SSS.Text & "' "
        End If

        CNN.Execute("UPDATE AP_Rate_history SET Rate2=Rate where Rate2 IS NULL ")

        FG_Rate.Rows = 1
        With rs
            Dim PPP As String = "select AP_Rate_history.* FROM         AP_Rate_history INNER JOIN " & _
                      "  Curr_For_Rate ON AP_Rate_history.curr = Curr_For_Rate.Curr  WHERE 1=1  " & Sql & "  " & SS_Curr & " order by  rate_dt ,Curr_For_Rate.cnt      "
            Call LoadSqlData(PPP, rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    FG_Rate.AddItem(.AbsolutePosition & _
                    Chr(9) & (.Fields("curr").Value.ToString) & _
                    Chr(9) & Format(.Fields("rate_dt").Value, "dd/MM/yyyy") & _
                     Chr(9) & Format(CDbl(.Fields("Rate").Value), "#,###0.00########") & _
                                 Chr(9) & Format(CDbl(.Fields("Rate2").Value), "#,###0.00########") & _
                        Chr(9) & (.Fields("Curr_name").Value.ToString) & _
                    Chr(9) & (.Fields("user_updt").Value.ToString) & _
                    Chr(9) & (.Fields("lst_updt").Value.ToString))
                    .MoveNext()
                End While
            End If
        End With
        'FG_Rate.Rows = 1
        'With rs
        '    Call LoadSqlData("select *  from AP_Rate_history WHERE curr<>'' " & Sql & " order by curr", rs)
        '    If .RecordCount > 0 Then
        '        While Not .EOF()
        '            FG_Rate.AddItem(.AbsolutePosition & _
        '            Chr(9) & (.Fields("curr").Value.ToString) & _
        '            Chr(9) & Format(.Fields("rate_dt").Value, "dd/MM/yyyy") & _
        '            Chr(9) & Format(CDbl(.Fields("THB_LAK").Value), "#,###0.00########") & _
        '            Chr(9) & Format(CDbl(.Fields("USD_THB").Value), "#,###0.00########") & _
        '            Chr(9) & Format(CDbl(.Fields("EUR_THB").Value), "#,###0.00########") & _
        '            Chr(9) & Format(CDbl(.Fields("USD_LAK").Value), "#,###0.00########") & _
        '            Chr(9) & Format(CDbl(.Fields("EUR_USD").Value), "#,###0.00########") & _
        '            Chr(9) & Format(CDbl(.Fields("EUR_LAK").Value), "#,###0.00########") & _
        '            Chr(9) & (.Fields("user_updt").Value.ToString) & _
        '            Chr(9) & (.Fields("lst_updt").Value.ToString))
        '            .MoveNext()
        '        End While
        '    End If
        'End With
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
        'Apimage = True
        Me.Close()
    End Sub


    Private Sub FG_Rate_DblClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles FG_Rate.DblClick
        CMB_Curr.Text = (FG_Rate.get_TextMatrix(FG_Rate.Row, 1))
        txtRate.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 3)), "#,###0.00########")
        txtRate2.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 4)), "#,###0.00########")
        DTrate.Value = Format(CDate(FG_Rate.get_TextMatrix(FG_Rate.Row, 2)), "dd/MM/yyyy")
        Dtsql = Format(CDate(FG_Rate.get_TextMatrix(FG_Rate.Row, 2)), "dd/MM/yyyy")


        txtcurr_name2.Text = (FG_Rate.get_TextMatrix(FG_Rate.Row, 5))

    End Sub

    Private Sub FG_Rate_MouseUpEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseUpEvent) Handles FG_Rate.MouseUpEvent
        CMB_Curr.Text = (FG_Rate.get_TextMatrix(FG_Rate.Row, 1))

        txtRate.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 3)), "#,###0.00########")
        txtRate2.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 4)), "#,###0.00########")
        DTrate.Value = Format(CDate(FG_Rate.get_TextMatrix(FG_Rate.Row, 2)), "dd/MM/yyyy")
        Dtsql = Format(CDate(FG_Rate.get_TextMatrix(FG_Rate.Row, 2)), "dd/MM/yyyy")

        txtcurr_name2.Text = (FG_Rate.get_TextMatrix(FG_Rate.Row, 5))

    End Sub
    Private Sub FG_Rate_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG_Rate.SelChange

        'CMB_Curr.Text = (FG_Rate.get_TextMatrix(FG_Rate.Row, 1))
        txtRate.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 3)), "#,###0.00########")
        txtRate2.Text = Format(CDbl(FG_Rate.get_TextMatrix(FG_Rate.Row, 4)), "#,###0.00########")
        DTrate.Value = Format(CDate(FG_Rate.get_TextMatrix(FG_Rate.Row, 2)), "dd/MM/yyyy")
        Dtsql = Format(CDate(FG_Rate.get_TextMatrix(FG_Rate.Row, 2)), "dd/MM/yyyy")

        DTrate.Value = Dtsql
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDel.Click
        AccCD = FG_Rate.get_TextMatrix(FG_Rate.Row, 2)
        If FG_Rate.Rows = 2 Then MsgBox("You do not delete , because is list last  ", MsgBoxStyle.OkOnly) : Exit Sub
        'Call LoadSqlData(" SELECT user_updt FROM AP_Rate_History WHERE (user_updt)= " & MUserName & " ", rs)
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
        Call LoadData()
    End Sub

    Private Sub txtTHB_LAK_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTHB_LAK.KeyPress
        'Select Case Asc(e.KeyChar)
        '    Case 48 To 57, 8
        '    Case Else
        '        e.Handled = True
        'End Select

        If e.KeyChar = Chr(13) Then
            txtUSD_LAK.Focus()
        End If
    End Sub
    Private Sub txtTHB_LAK_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTHB_LAK.LostFocus
        txtTHB_LAK.Text = Format(CDbl(txtTHB_LAK.Text), "#,###0.00########")
        If Trim(txtTHB_LAK.Text) = "" Or Trim(txtTHB_LAK.Text) = 0 Then txtTHB_LAK.Text = "1.00"
    End Sub
    Private Sub txtTHB_LAK_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTHB_LAK.TextChanged
        'txtTHB_LAK.Text = Format(CDbl(txtTHB_LAK.Text), "#,###0.00########")
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
        txtEUR_THB.Text = Format(CDbl(txtEUR_THB.Text), "#,###0.00########")
        txtEUR_LAK.Text = Format(CDbl(txtTHB_LAK.Text) * CDbl(txtEUR_THB.Text), "#,###0.00########")
        txtEUR_USD.Text = Format(CDbl(txtEUR_THB.Text) / CDbl(txtUSD_THB.Text), "#,###0.00########")
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
        txtEUR_USD.Text = Format(CDbl(txtEUR_USD.Text), "#,###0.00########")
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
        'Select Case Asc(e.KeyChar)
        '    Case 48 To 57, 8
        '    Case Else
        '        e.Handled = True
        'End Select
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


    Private Sub DTrate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTrate.ValueChanged

    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click

    End Sub

    Private Sub Cmb_Component_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmb_Component.SelectedIndexChanged
        Dim aa, bb, cc As String
        aa = (Trim(Cmb_Component.Text))
        cc = Microsoft.VisualBasic.Left(aa, 12)
        cc = Trim(cc)
        'MsgBox(aa)
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

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        txtCurr.Text = ""
        txtCurr.Enabled = True


    End Sub

    Private Sub FG_Curr_MouseUpEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseUpEvent) Handles FG_Curr.MouseUpEvent
        txtCurr.Enabled = False
        txtCurr.Text = FG_Curr.get_TextMatrix(FG_Curr.Row, 1)
        txtcurr_name.Text = FG_Curr.get_TextMatrix(FG_Curr.Row, 2)
    End Sub

    Private Sub FG_Curr_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG_Curr.SelChange

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        If txtcurr_name2.Text = "" Then MsgBox("ກະລຸນາ ໃສ່ສະກຸນເງິນ ເປັນພາສາລາວ ກ່ອນ!", MsgBoxStyle.OkOnly) : txtcurr_name2.Focus() : Exit Sub
        If txtRate.Text = "" Then MsgBox("ກະລຸນາ ອັດຕາແລກປ່ຽນ ກ່ອນ!", MsgBoxStyle.OkOnly) : txtRate.Focus() : Exit Sub

        If CDbl(txtRate2.Text) = CDbl(txtRate.Text) Then
            txtRate2.Text = txtRate.Text
        End If

        'On Error GoTo ll
        'CNN.BeginTrans()
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
            'aa = " delete from AP_Rate_history WHERE rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "' and  Com_id=N'" & Trim(txt_Component_id.Text) & "'  "
            'CNN.Execute(aa)

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
        If MessageBox.Show("ທ່ານຕ້ອງການລຶບລາຍການ : " & Trim(FG_Curr.get_TextMatrix(FG_Curr.Row, 1)) & "  ນີ້ແທ້ບໍ່?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
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
        If MessageBox.Show("ທ່ານຕ້ອງການລຶບລາຍການ : " & Trim(FG_Curr.get_TextMatrix(FG_Curr.Row, 1)) & "  ນີ້ແທ້ບໍ່?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            aa = " delete   AP_Rate_history WHERE   rate_dt='" & Format(DTrate.Value, "yyyy-MM-dd") & "'  " & _
                        "  and  Curr=N'" & Trim(CMB_Curr.Text) & "'  "
            CNN.Execute(aa)
            Call LoadData()
        End If
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Call LoadData()
    End Sub

    Private Sub CMB_Curr_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMB_Curr.SelectedIndexChanged
        Dim rs As New ADODB.Recordset
        Call LoadSqlData("Select * From Curr_For_Rate Where   Curr =N'" & Trim(CMB_Curr.Text) & "'", rs)
        If rs.RecordCount > 0 Then
            txtcurr_name2.Text = Trim(rs("Curr_name").Value.ToString)

        End If
    End Sub

    Private Sub CMB_Curr_SSS_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMB_Curr_SSS.SelectedIndexChanged

    End Sub

    Private Sub txtRate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRate.TextChanged
        'If txtRate2.Text = "" Then
        '    txtRate2.Text = txtRate.Text
        'End If
    End Sub
End Class