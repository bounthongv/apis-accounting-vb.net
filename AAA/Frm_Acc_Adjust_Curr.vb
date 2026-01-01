Public Class Frm_Acc_Adjust_Curr
    Dim Amount As Double
    Dim Amt As Double
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

    Private Sub Frm_Acc_Adjust_Curr_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Setup columns for DataGridView
        FG.Columns.Clear()
        FG.Columns.Add("col0", "ລ/ດ") ' 0
        FG.Columns.Add("col1", "ເລກບັນຊີໜີ້") ' 1
        FG.Columns.Add("col2", "ເລກບັນຊີມີ") ' 2
        FG.Columns.Add("col3", "ຊື່ບັນຊີພາສາລາວ") ' 3
        FG.Columns.Add("col4", "ຊື່ບັນຊີພາສາອັງກິດ") ' 4
        FG.Columns.Add("col5", "ຈໍານວນເງິນຈົດໜີ້") ' 5
        FG.Columns.Add("col6", "ຈໍານວນເງິນຈົດມີ") ' 6
        FG.Columns.Add("col7", "ສະກຸນເງິນ") ' 7
        FG.Columns.Add("col8", "ອັດຕາແລກປ່ຽນ") ' 8
        FG.Columns.Add("col9", "ມູນຄ່າເປັນກີບ ໜີ້") ' 9
        FG.Columns.Add("col10", "ມູນຄ່າເປັນກີບ ມີ") ' 10
        FG.Columns.Add("col11", "ມູນຄ່າເປັນ(USD) ໜີ້") ' 11
        FG.Columns.Add("col12", "ມູນຄ່າເປັນ(USD) ມີ") ' 12

        FG.Columns(5).Visible = False
        FG.Columns(6).Visible = False
        FG.Columns(7).Visible = False
        FG.Columns(8).Visible = False
        CMB_Curr.Items.Clear()
        Call load_Cmb(" SELECT Curr  FROM Curr_For_Rate  ORDER BY cnt ", "Curr", CMB_Curr)
        If CMB_Curr.Items.Count > 0 Then
            CMB_Curr.SelectedIndex = 0
        End If
 
        'Call RateSetting()
        If EditActive = True Then
            Call Loaddata()
        Else
            Call Addnew()
            Edit_Pro = 0
        End If
        CMBBK_ID.Items.Clear()
        Call load_Cmb(" SELECT bookid,bookname  FROM books  ORDER BY bookid ", "bookid", CMBBK_ID)
        If CMBBK_ID.Items.Count > 0 Then
            CMBBK_ID.SelectedIndex = 0
        End If
    End Sub

    Private Sub Loaddata()
        Dim rs As New ADODB.Recordset
        Dim sa As String = ""
        txtBill_no.Enabled = False
        FG.Rows.Clear()
        With rs
            sa = " SELECT     AP_ACC_Adjust.*,  " & _
               " AP_Office.off_id, AP_Office.off_nm,  " & _
         "     books.bookname " & _
           "  FROM   AP_ACC_Adjust INNER JOIN " & _
                   "    books ON AP_ACC_Adjust.Book = books.bookid INNER JOIN " & _
                   "    AP_Office ON AP_ACC_Adjust.Com_id = AP_Office.off_id  WHERE 1=1 AND  AP_ACC_Adjust.certify='" & SaleID & "' "
            Call LoadSqlData(sa, rs)


            If .RecordCount <> 0 Then
                MD_Curr = (.Fields("Curr").Value.ToString)
                CMB_Curr.Text = (.Fields("Curr").Value.ToString)
                txtBill_no.Text = (.Fields("certify").Value.ToString)
                Txt_Referno.Text = (.Fields("Referno").Value.ToString)
                txt_dt.Value = Format((.Fields("date_work").Value.ToString))
                CMBBK_ID.Text = Format((.Fields("Book").Value.ToString))
                txtBook_nm.Text = Format((.Fields("bookname").Value.ToString))
                txt_descrip.Text = Format((.Fields("descrip").Value.ToString))
                txt_descripE.Text = Format((.Fields("descripE").Value.ToString))
                txtAmount.Text = Format(CDbl(.Fields("amount").Value), "#,##0.00")
                txt_Curr.Text = Format((.Fields("Curr").Value.ToString))

                txtAmount_Lak.Text = Format(CDbl(.Fields("net_amt").Value), "#,##0.00")
                txtAmount_USD.Text = Format(CDbl(.Fields("net_usd").Value), "#,##0.00")
                txtAmount_Later.Text = Format((.Fields("Amount_Later").Value.ToString))
                txtSumAmountDr.Text = Format(CDbl(.Fields("AmountDr").Value), "#,##0.00")
                txtSumAmountCr.Text = Format(CDbl(.Fields("AmountCr").Value), "#,##0.00")
                txtSumTotalAmountDr.Text = Format(CDbl(.Fields("TotalAmountDr").Value), "#,##0.00")
                txtSumTotalAmountCr.Text = Format(CDbl(.Fields("TotalAmountCr").Value), "#,##0.00")

                'txtSupp_Nm.Text = Format((.Fields("K_Nm").Value.ToString))
                txtRate.Text = Format(CDbl(.Fields("Rate").Value), "#,##0.00")

            End If
        End With


        If rs.State = ConnectionState.Open Then rs.Close()
        FG.Rows.Clear()
        With rs
            Call LoadSqlData("  SELECT * from AP_ACC_adjust_Item WHERE AP_ACC_adjust_Item.certify='" & SaleID & "' ", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    Dim rowData As String() = { _
                        .AbsolutePosition.ToString(), _
                        .Fields("code_dr").Value.ToString(), _
                        Format(.Fields("code_cr").Value.ToString()), _
                        Format(.Fields("ac_name").Value.ToString()), _
                        "", _
                        Format(CDbl(.Fields("amount_dr").Value), "#,##0.00"), _
                        Format(CDbl(.Fields("amount_Cr").Value), "#,##0.00"), _
                        Format(.Fields("Curr_i").Value.ToString()), _
                        Format(CDbl(.Fields("Rate_i").Value), "#,##0.00"), _
                        Format(CDbl(.Fields("amt_dr").Value), "#,##0.00"), _
                        Format(CDbl(.Fields("amt_cr").Value), "#,##0.00") _
                    }
                    FG.Rows.Add(rowData)
                    .MoveNext()
                End While
            End If
        End With

    End Sub

    Private Sub FG_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles FG.CellEndEdit
        AccId = ""
        AccName = ""
        If FG.Rows(FG.Rows.Count - 1).Cells(3).Value.ToString() <> "" Then
            Exit Sub
        End If

        If e.ColumnIndex = 1 Then
            R = e.RowIndex
            L = e.ColumnIndex
            If FG.Rows(R).Cells(1).Value.ToString() <> "" Then

                AccId = FG.Rows(R).Cells(1).Value.ToString()
                LoadText()
                FG.Rows(R).Cells(3).Value = AccName
                Call load_Ac()
                If FG.Rows(R).Cells(3).Value.ToString() = "" Then
                    FG.Rows(R).Cells(1).Value = ""
                    MD_FROM = Me.Name
                    If AccId <> "" Then
                        MD_KH = AccId
                    End If
                    fmShartOfAccDetail.ShowDialog()
                End If
            End If
        End If

        If e.ColumnIndex = 2 Then
            R = e.RowIndex
            L = e.ColumnIndex
            If FG.Rows(R).Cells(2).Value.ToString() <> "" Then

                AccId = FG.Rows(R).Cells(2).Value.ToString()
                LoadText()
                FG.Rows(R).Cells(3).Value = AccName
                FG.Rows(R).Cells(4).Value = AccNamee
                If L = 1 Then
                    Call load_Ac()
                Else
                    Call load_Ac2()
                End If

                If FG.Rows(R).Cells(3).Value.ToString() = "" Then
                    FG.Rows(R).Cells(1).Value = ""
                    MD_FROM = Me.Name
                    If AccId <> "" Then
                        MD_KH = AccId
                    End If
                    fmShartOfAccDetail.ShowDialog()
                End If
            End If
        End If
        Call Load_Calculate()
    End Sub

    Private Sub load_Ac()
        If L = 1 Then
            MD_AMTDR = Format(CDbl(txtDiff.Text), "#,##0.00")
            MD_AMTCR = 0
            MD_AMTDR_LAK = Format(CDbl(txtDiff.Text), "#,##0.00")
            MD_AMTCR_LAK = 0
        Else
            MD_AMTCR = Format(CDbl(txtDiff.Text), "#,##0.00")
            MD_AMTDR = 0
            MD_AMTCR_LAK = Format(CDbl(txtDiff.Text), "#,##0.00")
            MD_AMTDR_LAK = 0
        End If

        If MD_AMTDR < 0 Then
            MD_AMTDR = CDbl(MD_AMTDR * -1)
        End If
        If MD_AMTCR < 0 Then
            MD_AMTCR = CDbl(MD_AMTCR * -1)
        End If

        FG.Rows(R).Cells(L).Value = FG.Rows(R).Cells(1).Value.ToString()
        FG.Rows(R).Cells(5).Value = Format(CDbl(MD_AMTDR), "#,##0.00")
        FG.Rows(R).Cells(6).Value = Format(CDbl(MD_AMTCR), "#,##0.00")
        FG.Rows(R).Cells(7).Value = "LAK"
        FG.Rows(R).Cells(8).Value = 1
        FG.Rows(R).Cells(9).Value = Format(CDbl(MD_AMTDR), "#,##0.00")
        FG.Rows(R).Cells(10).Value = Format(CDbl(MD_AMTCR), "#,##0.00")
    End Sub
    Private Sub load_Ac2()
        If L = 1 Then
            MD_AMTDR = Format(CDbl(txtDiff.Text), "#,##0.00")
            MD_AMTCR = 0
            MD_AMTDR_LAK = Format(CDbl(txtDiff.Text), "#,##0.00")
            MD_AMTCR_LAK = 0
        Else
            MD_AMTCR = Format(CDbl(txtDiff.Text), "#,##0.00")
            MD_AMTDR = 0
            MD_AMTCR_LAK = Format(CDbl(txtDiff.Text), "#,##0.00")
            MD_AMTDR_LAK = 0
        End If

        If MD_AMTDR < 0 Then
            MD_AMTDR = CDbl(MD_AMTDR * -1)
        End If
        If MD_AMTCR < 0 Then
            MD_AMTCR = CDbl(MD_AMTCR * -1)
        End If

        FG.Rows(R).Cells(5).Value = Format(CDbl(MD_AMTDR), "#,##0.00")
        FG.Rows(R).Cells(6).Value = Format(CDbl(MD_AMTCR), "#,##0.00")
        FG.Rows(R).Cells(7).Value = "LAK"
        FG.Rows(R).Cells(8).Value = 1
        FG.Rows(R).Cells(9).Value = Format(CDbl(MD_AMTDR), "#,##0.00")
        FG.Rows(R).Cells(10).Value = Format(CDbl(MD_AMTCR), "#,##0.00")
    End Sub
    Public Sub LoadText()
        AccName = ""
        AccNamee = ""
        Call LoadSqlData("SELECT * FROM Acc_Code WHERE AC_CODE = N'" & AccId & "'", RSC)
        With RSC
            Do Until .EOF = True
                AccId = Trim(.Fields("AC_CODE").Value)
                AccName = Trim(.Fields("Name_L").Value)
                AccNamee = Trim(.Fields("Name_e").Value)
                .MoveNext()
            Loop
        End With
    End Sub
    Private Sub FG_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles FG.CellDoubleClick
        If e.RowIndex < 0 Or e.RowIndex = FG.Rows.Count - 1 Then Exit Sub
        AccCD = FG.Rows(e.RowIndex).Cells(1).Value.ToString() & FG.Rows(e.RowIndex).Cells(2).Value.ToString()
        If MessageBox.Show("ທ່ານຕ້ອງການລືບລາຍການ'" & AccCD & "' ນີ້ ແທ້ ຫຼື ບໍ່ ?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            FG.Rows.RemoveAt(e.RowIndex)
            Call Calculate()
        End If
    End Sub

    Private Sub FG_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles FG.KeyUp
        If e.KeyCode = Keys.Enter Then
            If FG.CurrentCell.ColumnIndex = 4 Then
                FG.CurrentRow.Cells(9).Value = Format(CDbl(FG.CurrentRow.Cells(5).Value), "#,##0.00")
            End If
            If FG.CurrentCell.ColumnIndex = 5 Then
                FG.CurrentRow.Cells(10).Value = Format(CDbl(FG.CurrentRow.Cells(6).Value), "#,##0.00")
            End If

            FG.CurrentRow.Cells(5).Value = Format(CDbl(FG.CurrentRow.Cells(5).Value), "#,##0.00")
            FG.CurrentRow.Cells(6).Value = Format(CDbl(FG.CurrentRow.Cells(6).Value), "#,##0.00")

            If CMB_Curr.Text = "USD" Then
                FG.CurrentRow.Cells(7).Value = "USD"
                FG.CurrentRow.Cells(9).Value = Format(CDbl(FG.CurrentRow.Cells(5).Value), "#,##0.00")
                FG.CurrentRow.Cells(10).Value = Format(CDbl(FG.CurrentRow.Cells(6).Value), "#,##0.00")
            ElseIf CMB_Curr.Text = "THB" Then
                FG.CurrentRow.Cells(7).Value = "THB"
                If CDbl(FG.CurrentRow.Cells(5).Value) > 0 Then
                    MD_AMTDR = FG.CurrentRow.Cells(5).Value
                Else
                    MD_AMTCR = FG.CurrentRow.Cells(6).Value
                End If
            Else
                FG.CurrentRow.Cells(7).Value = "LAK"
                FG.CurrentRow.Cells(9).Value = Format(CDbl(FG.CurrentRow.Cells(5).Value), "#,##0.00")
                FG.CurrentRow.Cells(10).Value = Format(CDbl(FG.CurrentRow.Cells(6).Value), "#,##0.00")
            End If

            If txt_descrip.Text = "" Then
                txt_descrip.Text = FG.Rows(0).Cells(3).Value.ToString()
            End If
        End If
        Call Calculate()
    End Sub
    Public Sub LoadColor()
        ' FG cell styling not applied
    End Sub
    Private Sub Load_Calculate()
        For i = 0 To FG.Rows.Count - 1
            Call Calculate()
        Next i
    End Sub
    Private Sub Calculate()
        Dim i As Integer
        Dim amt1, amt2, amt3, amt4 As Double
        amt1 = 0 : amt2 = 0 : amt3 = 0 : amt4 = 0

        For i = 0 To FG.Rows.Count - 1
            amt1 = amt1 + CDbl(FG.Rows(i).Cells(5).Value)
            amt2 = amt2 + CDbl(FG.Rows(i).Cells(6).Value)
            amt3 = amt3 + CDbl(FG.Rows(i).Cells(9).Value)
            amt4 = amt4 + CDbl(FG.Rows(i).Cells(10).Value)
        Next i
        txtSumAmountDr.Text = Format(amt1, "#,##0.00")
        txtSumAmountCr.Text = Format(amt2, "#,##0.00")
        txtSumTotalAmountDr.Text = Format(amt3, "#,##0.00")
        txtSumTotalAmountCr.Text = Format(amt4, "#,##0.00")

        If CDbl(txtSumAmountDr.Text) > CDbl(txtSumAmountCr.Text) Then
            Dr.Text = Format(CDbl(txtSumAmountCr.Text) - CDbl(txtSumAmountDr.Text), "#,##0.00")
            Cr.Text = "0.00"
            DDR.Text = Format(CDbl(txtSumTotalAmountCr.Text) - CDbl(txtSumTotalAmountDr.Text), "#,##0.00")
            CCR.Text = "0.00"
        End If

        If CDbl(txtSumAmountCr.Text) > CDbl(txtSumAmountDr.Text) Then
            Cr.Text = Format(CDbl(txtSumAmountDr.Text) - CDbl(txtSumAmountCr.Text), "#,##0.00")
            Dr.Text = "0.00"
            CCR.Text = Format(CDbl(txtSumTotalAmountDr.Text) - CDbl(txtSumAmountCr.Text), "#,##0.00")
            DDR.Text = "0.00"
        End If

        If CDbl(txtSumAmountDr.Text) = CDbl(txtSumAmountCr.Text) Then
            Cr.Text = "0.00" : CCR.Text = "0.00" : Dr.Text = "0.00" : DDR.Text = "0.00"
        End If
    End Sub
    Private Sub FG_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles FG.MouseDown
        BtnSearch.Visible = False
    End Sub

    Private Sub FG_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelectionChanged
        If FG.CurrentRow Is Nothing Then Exit Sub
        MD_Curr = CMB_Curr.Text
        MD_Rate = txtRate.Text
        R = FG.CurrentRow.Index
        L = FG.CurrentCell.ColumnIndex
        Label1.Text = R
        Label2.Text = L
        Load_fg()
    End Sub
    Private Sub Amount_Later()
        If CMB_Curr.Text = "LAK" Then
            txtAmount_Lak.Text = Format(CDbl(txtAmount.Text), "##,##0.00")
        Else
            txtAmount_Lak.Text = Format(CDbl(txtAmount.Text) * CDbl(txtRate.Text), "##,##0.00")
        End If
        txtAmount_Later.Text = Letter_amt(txtAmount) & " ກີບ"
    End Sub
    Public Function Letter_amt(ByVal Txt As TextBox, Optional ByVal CurrKIP As Boolean = False) As String
        If Val(txtAmount.Text) <> 0 Then
            Letter_amt = CMoney(Format(CDbl(Txt.Text), "##0.00"))
        Else
            Letter_amt = ""
        End If
    End Function

    Private Sub BtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch.Click
        fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
        FG.Focus()
        If Label2.Text = "1" Then
        End If
        If txt_descrip.Text = "" Then
            txt_descrip.Text = FG.Rows(0).Cells(3).Value.ToString()
        End If
        If FG.Rows(R).Cells(12).Value.ToString() = "" Then
            FG.Rows(R).Cells(12).Value = Txt_Referno.Text
        End If
    End Sub

    Public Sub Load_fg()
        For i = 0 To FG.Rows.Count - 2
            FG.Rows(i).Cells(0).Value = i + 1
        Next i
    End Sub

    Private Sub BtnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew.Click
        Call Addnew()
        Load_fg()
    End Sub
    Private Sub Addnew()
        FG.Rows.Clear()
        FG.Rows.Add()
        txt_dt.Value = Date.Today
        Txt_Referno.Text = ""
        txtBill_no.Text = ""
        txt_descrip.Text = ""
        txt_descripE.Text = ""
        txtAmount.Text = "0.00"
        txtAmount_Later.Text = "ກີບ"
        txtSumAmountDr.Text = "0.00"
        txtSumAmountCr.Text = "0.00"
        txtSumTotalAmountDr.Text = "0.00"
        txtSumTotalAmountCr.Text = "0.00"
        txtamt.Text = "0.00"
        txtAmount_Lak.Text = "0.00"
        txtDiff.Text = "0.00"
        Txt_Referno.Focus()
        txtAmount_USD.Text = "0.00"
        CMB_Curr.SelectedIndex = 0
        Call RateSetting()
        txtRate.Text = Format(MD_Rate, "#,##0.00")
        txtRateUSD.Text = Format(MDUSD_LAK, "#,##0.00")
        Call Amount_Later()
    End Sub

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        If CDbl(txtSumTotalAmountDr.Text) <> CDbl(txtSumTotalAmountCr.Text) Then MsgBox("ບັນຊີນີ້ບໍ່ທັນບໍ່ທັນດູນດ່ຽງ ບໍ່ສາມາດບັນທຶກໄດ້!", MsgBoxStyle.OkOnly) : Exit Sub
        If FG.Rows.Count = 0 Then MsgBox("ທ່ານຍັງບໍ່ທັນໄດ້ລົງລາຍການບັນຊີ", MsgBoxStyle.OkOnly) : Exit Sub
        If txtBill_no.Text = "" Then Call AutoNumber()

        If txt_descrip.Text = "" Then txt_descrip.Text = txtAC_code_nm.Text

        Call LoadSqlData("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book =N'" & CMBBK_ID.Text & "' And  certify = N'" & txtBill_no.Text & "' ", RSC)
        If RSC.RecordCount > 0 Then
            If txtBill_no.Enabled = True Then
                MsgBox("ເລກລະຫັດ : " & Trim(txtBill_no.Text) & " ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
                Exit Sub
            End If
        End If

        Call Load_Delete()
        Call SaveItems()
        MsgBox("ບັນທຶກສຳເລັດ")
    End Sub
    Private Sub Load_Delete()
        CNN.Execute("DELETE FROM gen_jn WHERE certify=N'" & txtBill_no.Text & "' ")
    End Sub

    Private Sub AutoNumber()
        Dim VIOT As New ADODB.Recordset
        Call LoadSqlData("SELECT top 1 right(certify,5) As Bill_no from gen_jn WHERE 1=1 AND Book=N'" & CMBBK_ID.Text & "' Order by certify DESC", VIOT)
        If VIOT.RecordCount <> 0 Then
            txtBill_no.Text = "AJ" & Format(CDbl(VIOT.Fields("Bill_no").Value) + 1, "00000")
        Else
            txtBill_no.Text = "AJ00001"
        End If
    End Sub
    
    Private Sub SaveItems()
        Dim J As Integer
        For J = 0 To FG.Rows.Count - 1
            If CMB_Curr.Text = "USD" Then
                FG.Rows(J).Cells(11).Value = Format(CDbl(FG.Rows(J).Cells(9).Value), "#,##0.00")
                FG.Rows(J).Cells(12).Value = Format(CDbl(FG.Rows(J).Cells(10).Value), "#,##0.00")
            Else
                FG.Rows(J).Cells(11).Value = Format(CDbl(FG.Rows(J).Cells(9).Value) / CDbl(txtRateUSD.Text), "#,##0.00")
                FG.Rows(J).Cells(12).Value = Format(CDbl(FG.Rows(J).Cells(10).Value) / CDbl(txtRateUSD.Text), "#,##0.00")
            End If
        Next J

        For i = 0 To FG.Rows.Count - 1
            Dim sa As String = "INSERT INTO gen_jn ( date_work, book, certify, descrip, descripe, amount, curr, rate, Rate_USD, net_amt, code_dr, code_cr, ac_code, ac_name, ac_namee, " & _
            " amount_dr, amount_cr, Curr_i, Rate_i, amt_dr, amt_cr, amt_USD_Dr, amt_USD_Cr, company, office_id, com_id, last_update, last_user, pc_nm,AG,Frm) " & _
            " VALUES ( '" & Format(txt_dt.Value, "yyyy-MM-dd") & "'," & _
            "  N'" & Apostrophe(CMBBK_ID.Text) & "'," & _
             "  N'" & Apostrophe(txtBill_no.Text) & "'," & _
               "  N'" & Apostrophe(txt_descrip.Text) & "'," & _
                 "  N'" & Apostrophe(txt_descripE.Text) & "'," & _
              " " & CDbl(txtAmount.Text) & ", " & _
                 "  N'LAK', 1, " & CDbl(txtRateUSD.Text) & ", " & CDbl(txtDiff.Text) & ", " & _
            " N'" & Apostrophe(FG.Rows(i).Cells(1).Value.ToString()) & "'," & _
            " N'" & Apostrophe(FG.Rows(i).Cells(2).Value.ToString()) & "'," & _
            " N'" & Apostrophe(FG.Rows(i).Cells(1).Value.ToString()) + Apostrophe(FG.Rows(i).Cells(2).Value.ToString()) & "'," & _
            " N'" & Apostrophe(FG.Rows(i).Cells(3).Value.ToString()) & "'," & _
            " N'" & Apostrophe(FG.Rows(i).Cells(4).Value.ToString()) & "', " & _
            " " & CDbl(FG.Rows(i).Cells(5).Value) & ", " & _
             " " & CDbl(FG.Rows(i).Cells(6).Value) & "," & _
            " N'" & Apostrophe(FG.Rows(i).Cells(7).Value.ToString()) & "'," & _
            " " & CDbl(FG.Rows(i).Cells(8).Value) & "," & _
            " " & CDbl(FG.Rows(i).Cells(9).Value) & ", " & _
            " " & CDbl(FG.Rows(i).Cells(10).Value) & ", " & _
            " " & CDbl(FG.Rows(i).Cells(11).Value) & ", " & _
            " " & CDbl(FG.Rows(i).Cells(12).Value) & ", " & _
            "  N'" & Apostrophe(MuSubOff2) & "'," & _
            "  N'" & Apostrophe(Off_Id) & "'," & _
            "  N'" & Apostrophe(Off_Id) & "'," & _
            " Getdate()," & _
            " N'" & Apostrophe(MUserName) & "'," & _
            " N'" & Apostrophe(COMPUTER_NM) & "',0,'1' )"
            CNN.Execute(sa)
        Next i
    End Sub

    Private Sub Amount_USD()
        If CMB_Curr.Text = "USD" Then
            txtAmount_USD.Text = Format(CDbl(txtAmount.Text), "#,##0.00")
            txtAmount_Lak.Text = Format(CDbl(txtRate.Text) * CDbl(txtAmount.Text), "#,##0.00")
        ElseIf CMB_Curr.Text = "THB" Then
            txtAmount_Lak.Text = Format(CDbl(txtRate.Text) * CDbl(txtAmount.Text), "#,##0.00")
            txtAmount_USD.Text = Format(CDbl(txtAmount_Lak.Text) / CDbl(MDUSD_LAK), "#,##0.00")
        ElseIf CMB_Curr.Text = "LAK" Then
            txtAmount_Lak.Text = Format(CDbl(txtAmount.Text), "#,##0.00")
            txtAmount_USD.Text = Format(CDbl(txtAmount.Text) / CDbl(MDUSD_LAK), "#,##0.00")
        End If
    End Sub

    Private Sub CMBBK_ID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBBK_ID.SelectedIndexChanged
        Dim rs As New ADODB.Recordset
        Call LoadSqlData("Select * From books Where bookid =N'" & Trim(CMBBK_ID.Text) & "'", rs)
        If rs.RecordCount > 0 Then
            txtBook_nm.Text = Trim(rs("bookid").Value).ToString
        End If
    End Sub

    Private Sub txtAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call Amount_Later()
            txtAmount.Text = Format(CDbl(txtAmount.Text), "##,##0.00")
            Amount_USD()
        End If
    End Sub
    
    Private Sub txtAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAmount.LostFocus
        txtAmount.Text = Format(CDbl(txtAmount.Text), "##,##0.00")
    End Sub

    Private Sub txtAmount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmount.TextChanged
        Call Amount_Later()
    End Sub

    Private Sub txt_descrip_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_descrip.KeyDown
        If e.KeyCode = Keys.Enter Then txt_descripE.Focus()
    End Sub

    Private Sub txt_descripE_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_descripE.KeyDown
        If e.KeyCode = Keys.Enter Then
            txtAmount.Focus()
            txtAmount.SelectAll()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' Printing logic...
    End Sub

    Private Sub txtRate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRate.KeyDown
        If e.KeyCode = Keys.Enter Then
            If CDbl(txtRate.Text) < 1 Then txtRate.Text = "1.00"
            txtRate.Text = Format(CDbl(txtRate.Text), "##,##0.00")
            Amount_USD()
            Calculate()
        End If
    End Sub

    Private Sub txtAC_code_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAC_code.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim rs As New ADODB.Recordset
            Call LoadSqlData("Select * from ACC_CODE Where AC_CODE =N'" & Trim(txtAC_code.Text) & "'", rs)
            If rs.RecordCount > 0 Then
                txtAC_code_nm.Text = Trim(rs("Name_L").Value)
                txtAC_type.Text = Trim(rs("Acc_Type").Value.ToString)
            End If
            sumNew()
        End If
    End Sub

    Private Sub sumNew()
        ' sumNew implementation...
    End Sub

    Private Sub CMB_Curr_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMB_Curr.KeyDown
        If e.KeyCode = Keys.Enter Then
            FG.Rows.Clear()
            FG.Rows.Add()
            FG.Rows.Add()

            If CDbl(txtDiff.Text) > 0 Then
                FG.Rows(0).Cells(1).Value = txtAC_code.Text
                FG.Rows(0).Cells(3).Value = txtAC_code_nm.Text
                FG.Rows(0).Cells(5).Value = "0.00"
                FG.Rows(0).Cells(6).Value = "0.00"
                FG.Rows(0).Cells(7).Value = "LAK"
                FG.Rows(0).Cells(8).Value = "1.00"
                FG.Rows(0).Cells(9).Value = Format(CDbl(txtDiff.Text), "#,##0.00")
                FG.Rows(0).Cells(10).Value = "0.00"

                FG.Rows(1).Cells(2).Value = "5106100.00.000"
                FG.Rows(1).Cells(3).Value = "ກຳໄລຈາກການແລກປ່ຽນເງິນຕາຕ່າງປະເທດ"
                FG.Rows(1).Cells(9).Value = "0.00"
                FG.Rows(1).Cells(10).Value = Format(CDbl(txtDiff.Text), "#,##0.00")
            ElseIf CDbl(txtDiff.Text) < 0 Then
                FG.Rows(0).Cells(1).Value = "4106100.00.000"
                FG.Rows(0).Cells(3).Value = "ຂາດທຶນຈາກການແລກປ່ຽນເງິນຕາ ຕ່າງປະເທດ"
                FG.Rows(0).Cells(9).Value = Format(CDbl(txtDiff.Text) * -1, "#,##0.00")
                FG.Rows(0).Cells(10).Value = "0.00"

                FG.Rows(1).Cells(2).Value = txtAC_code.Text
                FG.Rows(1).Cells(3).Value = txtAC_code_nm.Text
                FG.Rows(1).Cells(9).Value = "0.00"
                FG.Rows(1).Cells(10).Value = Format(CDbl(txtDiff.Text) * -1, "#,##0.00")
            End If
            Calculate()
        End If
    End Sub

    Private Sub CMB_Curr_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMB_Curr.SelectedIndexChanged
        Call RateSetting()
        txtRate.Text = Format(MD_Rate, "#,##0.00")
        txtRateUSD.Text = Format(MDUSD_LAK, "#,##0.00")
        txtAmount_Lak.Text = Format(CDbl(txtAmount.Text) * CDbl(txtRate.Text), "##,##0.00")
        txtDiff.Text = Format(CDbl(txtAmount_Lak.Text) - CDbl(txtamt.Text), "#,##0.00")
        Call Amount_Later()
        Amount_USD()
    End Sub
End Class