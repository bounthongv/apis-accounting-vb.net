Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Data
Imports System.Drawing
Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.Linq
Imports System.Xml.Linq
Imports ApPBank10.Module

Public Class FmNsewJeneralJournal
    Dim sql As String
    Dim Rate1 As String
    Dim MdCertifyId, MdCertifyId2, Sdate As String
    Dim RateType As String
    Dim IVN As String
    Dim Book As String
    Dim Amount_In_Word As String
    Dim MDCust_Supp As String = ""
    Dim MCS As String
    Dim R, L As Integer
    Dim rowIdx As Integer

    ' DataGridView Helper Methods
    Private Function GetGridValue(ByVal grid As DataGridView, ByVal row As Integer, ByVal col As Integer) As String
        If row >= 0 AndAlso row < grid.RowCount AndAlso col >= 0 AndAlso col < grid.ColumnCount Then
            If grid.Rows(row).Cells(col).Value IsNot Nothing Then
                Return grid.Rows(row).Cells(col).Value.ToString()
            End If
        End If
        Return ""
    End Function

    Private Sub SetGridValue(ByVal grid As DataGridView, ByVal row As Integer, ByVal col As Integer, ByVal value As String)
        If row >= 0 AndAlso row < grid.RowCount AndAlso col >= 0 AndAlso col < grid.ColumnCount Then
            grid.Rows(row).Cells(col).Value = value
        End If
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        fmShartOfAccDetail.ShowDialog()
    End Sub

    Private Sub Savedata()
        Dim J As Integer
        For J = 0 To FG.Rows.Count - 1
            If GetGridValue(FG, J, 10) = "USD" Then
                SetGridValue(FG, J, 14, Format(CDbl(Val(GetGridValue(FG, J, 5))), "#,##0.00"))
                SetGridValue(FG, J, 15, Format(CDbl(Val(GetGridValue(FG, J, 6))), "#,##0.00"))
            Else
                SetGridValue(FG, J, 14, Format(CDbl(Val(GetGridValue(FG, J, 12))) / CDbl(Val(txtRate.Text)), "#,##0.00"))
                SetGridValue(FG, J, 15, Format(CDbl(Val(GetGridValue(FG, J, 13))) / CDbl(Val(txtRate.Text)), "#,##0.00"))
            End If
        Next J

        If MuLng = "L" Then
            Amount_In_Word = txtAmt_letter.Text
        Else
            Amount_In_Word = txtAmt_letter_E.Text
        End If
        MuSubOff = Mid(Off_Usr.Text, 1, 5)
        
        For i = 0 To FG.Rows.Count - 1
            If GetGridValue(FG, i, 1) = "" And GetGridValue(FG, i, 2) = "" Then
                FG.Rows.Clear()
                FG.Rows.Add()
                AutoNumber()
                Call NewText()
                Exit Sub
            End If

            Dim sqlInsert As String
            If CheckBox3.Checked Then
                If CheckBox4.Checked Then
                    sqlInsert = "INSERT INTO gen_jn( date_work, ac_Name, book, certify,Referno, cheque_no ,descrip ,descripe ,amount , curr ,rate, Rate_USD, net_amt ,code_dr ,code_cr ,ac_code    ,amt_dr , amt_cr , amt_USD_Dr, amt_USD_Cr , amount_dr ,amount_cr ,  certis, lock ,rec_lock , last_update , last_user , descflag , Cat_ID , certifyID  ,company  ,Office_ID, Cust_Supp, CustID,SuppID , del , AG,Frm) " & _
                              "Values('" & Format(dtActi.Value, "yyyy/MM/dd") & "',N'" & Trim(GetGridValue(FG, i, 3)) & "',N'" & CmbBook.Text & "',N'" & txtInvoice.Text & "',N'" & TxtReferno.Text & "','" & Apostrophe(txtChecq.Text) & "',N'" & Trim(Apostrophe(txtDesc.Text)) & "',N'" & Trim(Apostrophe(txtDescE.Text)) & "'," & CDbl(Val(txtAmount.Text)) & ",'" & GetGridValue(FG, i, 10) & "'," & CDbl(Val(GetGridValue(FG, i, 11))) & "," & CDbl(Val(txtRateUSD.Text)) & ",'" & "0" & "','" & GetGridValue(FG, i, 1) & "','" & GetGridValue(FG, i, 2) & "','" & GetGridValue(FG, i, 1) & GetGridValue(FG, i, 2) & "'," & CDbl(Val(GetGridValue(FG, i, 5))) & "," & CDbl(Val(GetGridValue(FG, i, 6))) & "," & CDbl(Val(GetGridValue(FG, i, 14))) & "," & CDbl(Val(GetGridValue(FG, i, 15))) & "," & CDbl(0) & "," & CDbl(0) & ",'" & "3" & "','" & "4" & "','" & "5" & "','" & dtActi.Text & "','" & MUserID & "','" & "8" & "','" & GetGridValue(FG, i, 7) & "','" & MdCertifyId & "','" & MuSubOff & "' ,'" & MuSubOff & "','" & MDCust_Supp & "',N'" & TxtCustID.Text & "',N'" & TxtSuppID.Text & "' , 0,1,0)"
                Else
                    sqlInsert = "INSERT INTO gen_jn( date_work, ac_Name, book, certify, Referno, cheque_no ,descrip ,descripe ,amount , curr ,rate, Rate_USD, net_amt ,code_dr ,code_cr ,ac_code    ,amt_dr , amt_cr , amt_USD_Dr, amt_USD_Cr , amount_dr ,amount_cr  ,  certis, lock ,rec_lock , last_update , last_user , descflag , Cat_ID , certifyID  ,company  ,Office_ID, Cust_Supp, CustID,SuppID , del , AG,Frm) " & _
                               "Values('" & Format(dtActi.Value, "yyyy/MM/dd") & "',N'" & GetGridValue(FG, i, 3) & "',N'" & CmbBook.Text & "',N'" & txtInvoice.Text & "',N'" & TxtReferno.Text & "','" & Apostrophe(txtChecq.Text) & "',N'" & Trim(Apostrophe(txtDesc.Text)) & "',N'" & Trim(Apostrophe(txtDescE.Text)) & "'," & CDbl(Val(txtAmount.Text)) & ",'" & GetGridValue(FG, i, 10) & "'," & CDbl(Val(GetGridValue(FG, i, 11))) & "," & CDbl(Val(txtRateUSD.Text)) & ",'" & "0" & "','" & GetGridValue(FG, i, 1) & "','" & GetGridValue(FG, i, 2) & "','" & GetGridValue(FG, i, 1) & GetGridValue(FG, i, 2) & "'," & CDbl(Val(GetGridValue(FG, i, 5))) & "," & CDbl(Val(GetGridValue(FG, i, 6))) & "," & CDbl(Val(GetGridValue(FG, i, 14))) & "," & CDbl(Val(GetGridValue(FG, i, 15))) & "," & CDbl(0) & "," & CDbl(0) & ",'" & "3" & "','" & "4" & "','" & "5" & "','" & dtActi.Text & "','" & MUserID & "','" & "8" & "','" & GetGridValue(FG, i, 7) & "','" & MdCertifyId & "','" & MuSubOff & "' ,'" & MuSubOff & "','" & MDCust_Supp & "',N'',N'' , 0,1,0)"
                End If
            Else
                If CheckBox4.Checked Then
                    sqlInsert = "INSERT INTO gen_jn( date_work, ac_Name, book, certify, Referno,cheque_no ,descrip ,descripe ,amount , curr ,rate,  Rate_USD,net_amt ,code_dr ,code_cr ,ac_code  , amount_dr ,amount_cr ,amt_dr , amt_cr  ,amt_USD_Dr, amt_USD_Cr ,certis, lock ,rec_lock , last_update , last_user , descflag , Cat_ID , certifyID  ,company   ,Office_ID, Cust_Supp, CustID,SuppID , del, AG,Frm) " & _
                                      "Values('" & Format(dtActi.Value, "yyyy/MM/dd") & "',N'" & GetGridValue(FG, i, 3) & "',N'" & CmbBook.Text & "',N'" & txtInvoice.Text & "',N'" & TxtReferno.Text & "','" & Apostrophe(txtChecq.Text) & "',N'" & Trim(Apostrophe(txtDesc.Text)) & "',N'" & Trim(Apostrophe(txtDescE.Text)) & "'," & CDbl(Val(txtAmount.Text)) & ",'" & GetGridValue(FG, i, 10) & "'," & CDbl(Val(GetGridValue(FG, i, 11))) & "," & CDbl(Val(txtRateUSD.Text)) & ",'" & "0" & "','" & GetGridValue(FG, i, 1) & "','" & GetGridValue(FG, i, 2) & "','" & GetGridValue(FG, i, 1) & GetGridValue(FG, i, 2) & "'," & CDbl(Val(GetGridValue(FG, i, 5))) & "," & CDbl(Val(GetGridValue(FG, i, 6))) & "," & CDbl(Val(GetGridValue(FG, i, 12))) & "," & CDbl(Val(GetGridValue(FG, i, 13))) & "," & CDbl(Val(GetGridValue(FG, i, 14))) & "," & CDbl(Val(GetGridValue(FG, i, 15))) & ",'" & "3" & "','" & "4" & "','" & "5" & "','" & dtActi.Text & "','" & MUserID & "','" & "8" & "','" & GetGridValue(FG, i, 7) & "','" & MdCertifyId & "','" & MuSubOff & "' , '" & MuSubOff & "','" & MDCust_Supp & "',N'" & TxtCustID.Text & "',N'" & TxtSuppID.Text & "' , 0,0,0)"
                Else
                    sqlInsert = "INSERT INTO gen_jn( date_work, ac_Name, book, certify, Referno,cheque_no ,descrip ,descripe ,amount , curr ,rate,  Rate_USD,net_amt ,code_dr ,code_cr ,ac_code  , amount_dr ,amount_cr ,amt_dr , amt_cr  ,amt_USD_Dr, amt_USD_Cr , certis, lock ,rec_lock , last_update , last_user , descflag , Cat_ID , certifyID  ,company   ,Office_ID, Cust_Supp, CustID, SuppID , del, AG,Frm) " & _
                          "Values('" & Format(dtActi.Value, "yyyy/MM/dd") & "',N'" & GetGridValue(FG, i, 3) & "',N'" & CmbBook.Text & "',N'" & txtInvoice.Text & "',N'" & TxtReferno.Text & "','" & Apostrophe(txtChecq.Text) & "',N'" & Trim(Apostrophe(txtDesc.Text)) & "',N'" & Trim(Apostrophe(txtDescE.Text)) & "'," & CDbl(Val(txtAmount.Text)) & ",'" & GetGridValue(FG, i, 10) & "'," & CDbl(Val(GetGridValue(FG, i, 11))) & "," & CDbl(Val(txtRateUSD.Text)) & ",'" & "0" & "','" & GetGridValue(FG, i, 1) & "','" & GetGridValue(FG, i, 2) & "','" & GetGridValue(FG, i, 1) & GetGridValue(FG, i, 2) & "'," & CDbl(Val(GetGridValue(FG, i, 5))) & "," & CDbl(Val(GetGridValue(FG, i, 6))) & "," & CDbl(Val(GetGridValue(FG, i, 12))) & "," & CDbl(Val(GetGridValue(FG, i, 13))) & "," & CDbl(Val(GetGridValue(FG, i, 14))) & "," & CDbl(Val(GetGridValue(FG, i, 15))) & ",'" & "3" & "','" & "4" & "','" & "5" & "','" & dtActi.Text & "','" & MUserID & "','" & "8" & "','" & GetGridValue(FG, i, 7) & "','" & MdCertifyId & "','" & MuSubOff & "' , '" & MuSubOff & "','" & MDCust_Supp & "',N'',N'' , 0,0,0)"
                End If
            End If
            DbHelper.ExecuteNonQuery(sqlInsert)
        Next i
        MuSubOff = MuSubOff2
        LngId = "6001" : MsgRpt()
    End Sub

    Private Sub FG_CellEndEdit(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FG.CellEndEdit
        If e.ColumnIndex = 1 OrElse e.ColumnIndex = 2 Then
            If GetGridValue(FG, e.RowIndex, 1) & GetGridValue(FG, e.RowIndex, 2) <> "" Then
                'If Len(GetGridValue(FG, e.RowIndex, 1) & GetGridValue(FG, e.RowIndex, 2)) <> 7 Then
                '    LngId = "6005" : MsgRpt()
                '    Exit Sub
                'End If
            End If
        End If
        If txtInvoice.BackColor = Color.Red Then
            txtInvoice.Focus()
            Exit Sub
        End If
        If txtInvoice.Text = "" Then
            MsgBox("ກະລຸນນາໃສ່ເລກໃບຢັງຢືນກອ່ນ!", MsgBoxStyle.OkOnly)
            txtInvoice.BackColor = Color.Red
            txtInvoice.Focus()
            Exit Sub
        End If
        If TxtReferno.Text = "" Then
            MsgBox("ກະລຸນນາໃສ່ເລກອ້າງອີງກ່ອນ!", MsgBoxStyle.OkOnly)
            'TxtReferno.BackColor = Color.Red
            TxtReferno.Focus()
            Exit Sub
        End If

        '=========jjjjjjj
        txtTotal_Amt_LAK.Text = CDbl(txtRate.Text) * CDbl(txtAmount.Text)
        If ChCat_ID.Checked = False Then
            AfterEdit2()
            loadColor()
            'MsgBox("jj")
            txtAmount.Focus()
        Else
            'AfterEdit()
            'loadColor()
            'txtAmount.Focus()
        End If

        If txtDesc.Text = "" Then
            txtDesc.Text = GetGridValue(FG, e.RowIndex, 3)
            txtDescE.Text = GetGridValue(FG, e.RowIndex, 4)
        End If
        If CDbl(txtAmount.Text) = 0 Then
            If FG.CurrentRow IsNot Nothing AndAlso FG.CurrentRow.Index = 1 Then
                txtAmount.Text = Format(CDbl((CDbl(GetGridValue(FG, e.RowIndex, 5)) + CDbl(GetGridValue(FG, e.RowIndex, 6)))), "##,##0.00")
            End If
        End If
        If CDbl(txtSumAmountDr.Text) >= CDbl(txtSumAmountCr.Text) Then
            txtAmount.Text = txtSumAmountDr.Text
        End If
        If CDbl(txtSumAmountDr.Text) <= CDbl(txtSumAmountCr.Text) Then
            txtAmount.Text = txtSumAmountCr.Text
        End If
    End Sub
    Private Sub AfterEdit()
        BtnMove.Visible = False
        '*************************Col-1-*********************
        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 1 Then
            R = FG.CurrentCell.RowIndex
            L = FG.CurrentCell.ColumnIndex
If GetGridValue(FG, R, 1) = "" Then
                MDSearchAcccode = GetGridValue(FG, R, 1)
                fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                fmShartOfAccDetail.ShowDialog()
            Else
                SumAmountDr()
                txtSumAmountDr.Text = CDbl(txtSumAmountDr.Text) - CDbl(GetGridValue(FG, R, 5))
                SetGridValue(FG, R, 5, Format(CDbl(CDbl(txtAmount.Text) - CDbl(txtSumAmountDr.Text)), "#,##0.00"))
                SetGridValue(FG, R, 6, "0.00")
                SetGridValue(FG, R, 7, 0)
                SetGridValue(FG, R, 10, Cmb.Text)
                SetGridValue(FG, R, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
                SetGridValue(FG, R, 12, Format(CDbl(CDbl(txtRate.Text) * GetGridValue(FG, R, 5)), "#,##0.00"))
                SetGridValue(FG, R, 13, "0.00")
                SetGridValue(FG, R, 14, Format(CDbl(CDbl(GetGridValue(FG, R, 12))) / CDbl(MDUSD_LAK), "#,##0.00"))
                SetGridValue(FG, R, 15, "0.00")
                SumAmountDr()

                If ChDe.Checked = True Then
                    If MuLng = "L" Then
                        FG.CurrentCell = FG.Rows(R).Cells(3)
                    End If
                    If MuLng = "E" Then
                        FG.CurrentCell = FG.Rows(R).Cells(4)
                    End If
                    If R < FG.Rows.Count - 1 AndAlso GetGridValue(FG, FG.Rows.Count - 2, 3) <> "" Then
                        FG.Rows.Add()
                    End If
                    Exit Sub
                End If

                FG.CurrentCell = FG.Rows(R).Cells(5)
                If R < FG.Rows.Count - 1 AndAlso GetGridValue(FG, FG.Rows.Count - 2, 3) <> "" Then
                    FG.Rows.Add()
                End If
Exit Sub
        End If
        If GetGridValue(FG, R, 1) <> "" Then
                AccId = GetGridValue(FG, R, 1)
                MDSearchAcccode = GetGridValue(FG, R, 1)
                LoadText()
                SetGridValue(FG, R, 3, AccName)
                SetGridValue(FG, R, 4, AccNamee)
                If GetGridValue(FG, R, 3) = "" Then
                    SetGridValue(FG, R, 1, "")
                    fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                    fmShartOfAccDetail.ShowDialog()
                End If
                SumAmountDr()
                txtSumAmountDr.Text = CDbl(txtSumAmountDr.Text) - CDbl(GetGridValue(FG, R, 5))
                SetGridValue(FG, R, 5, Format(CDbl(CDbl(txtAmount.Text) - CDbl(txtSumAmountDr.Text)), "#,##0.00"))
                SetGridValue(FG, R, 6, "0.00")
                SetGridValue(FG, R, 7, 0)
                SetGridValue(FG, R, 10, Cmb.Text)
                SetGridValue(FG, R, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
                SetGridValue(FG, R, 12, Format(CDbl(CDbl(txtRate.Text) * CDbl(GetGridValue(FG, R, 5))), "#,##0.00"))
                SetGridValue(FG, R, 13, "0.00")
                SetGridValue(FG, R, 14, Format(CDbl(CDbl(GetGridValue(FG, R, 12))) / CDbl(MDUSD_LAK), "#,##0.00"))
                SetGridValue(FG, R, 15, "0.00")
                SumAmountDr()

                If ChDe.Checked = True Then
                    If MuLng = "L" Then
                        FG.CurrentCell = FG.Rows(R).Cells(3)
                    End If
                    If MuLng = "E" Then
                        FG.CurrentCell = FG.Rows(R).Cells(4)
                    End If
                    If R < FG.Rows.Count - 1 AndAlso GetGridValue(FG, FG.Rows.Count - 1, 3) <> "" Then
                        FG.Rows.Add()
                    End If
                    Exit Sub
                End If

                FG.CurrentCell = FG.Rows(R).Cells(5)
                If R < FG.Rows.Count - 1 AndAlso GetGridValue(FG, FG.Rows.Count - 1, 3) <> "" Then
                    FG.Rows.Add()
                End If
                Exit Sub
            End If
        End If
'*************************Col-2-*********************
        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 2 Then
            R = FG.CurrentCell.RowIndex
            L = FG.CurrentCell.ColumnIndex
            If GetGridValue(FG, R, 2) = "" Then
                MDSearchAcccode = GetGridValue(FG, R, 2)
                fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                fmShartOfAccDetail.ShowDialog()
            End If
            If GetGridValue(FG, R, 2) <> "" Then
                AccId = GetGridValue(FG, R, 2)
                LoadText()
                SetGridValue(FG, R, 3, AccName)
                SetGridValue(FG, R, 4, AccNamee)
                If GetGridValue(FG, R, 3) = "" Then
                    MDSearchAcccode = GetGridValue(FG, R, 2)
                    fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                    fmShartOfAccDetail.ShowDialog()
                End If
                SumAmountDr()
                txtSumAmountCr.Text = CDbl(txtSumAmountCr.Text) - CDbl(GetGridValue(FG, R, 6))
                SetGridValue(FG, R, 6, Format(CDbl(CDbl(txtAmount.Text) - CDbl(txtSumAmountCr.Text)), "#,##0.00"))
                SetGridValue(FG, R, 5, "0.00")
                SetGridValue(FG, R, 7, 0)
                SetGridValue(FG, R, 10, Cmb.Text)
                SetGridValue(FG, R, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
                SetGridValue(FG, R, 13, Format(CDbl(CDbl(txtRate.Text) * GetGridValue(FG, R, 6)), "#,##0.00"))
                SetGridValue(FG, R, 12, "0.00")
                SetGridValue(FG, R, 15, Format(CDbl(CDbl(GetGridValue(FG, R, 13))) / CDbl(MDUSD_LAK), "#,##0.00"))
                SetGridValue(FG, R, 14, "0.00")
                SumAmountDr()
                If ChDe.Checked = True Then
                    If MuLng = "L" Then
                        FG.CurrentCell = FG.Rows(R).Cells(3)
                    End If
                    If MuLng = "E" Then
                        FG.CurrentCell = FG.Rows(R).Cells(4)
                    End If
                    If GetGridValue(FG, FG.RowCount - 2, 3) <> "" Then
                        FG.Rows.Add()
                    End If
                    Exit Sub
                End If
                FG.CurrentCell = FG.Rows(R).Cells(6)
                If GetGridValue(FG, FG.RowCount - 2, 3) <> "" Then
                    FG.Rows.Add()
                    FG.CurrentCell = FG.Rows(R).Cells(6)
                End If
                Exit Sub
            End If
            Exit Sub
End If


        '====*************************

        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 3 Then
            If Button3.Text = "ໂຊຂໍ້ມູນແບບທົ່ວໄປ" Then
                FG.CurrentCell = FG.Rows(FG.CurrentCell.RowIndex).Cells(4)
                Exit Sub
            Else
                If GetGridValue(FG, FG.CurrentCell.RowIndex, 1) <> "" Then
                    FG.CurrentCell = FG.Rows(FG.CurrentCell.RowIndex).Cells(5)
                    Exit Sub
                End If
                If GetGridValue(FG, FG.CurrentCell.RowIndex, 2) <> "" Then
                    FG.CurrentCell = FG.Rows(FG.CurrentCell.RowIndex).Cells(6)
                    Exit Sub
                End If
                'MsgBox("ໂຊຂໍ້ມູນແບບລະອຽດ")
            End If
End If
        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 4 Then
            If GetGridValue(FG, FG.CurrentCell.RowIndex, 1) <> "" Then
                FG.CurrentCell = FG.Rows(FG.CurrentCell.RowIndex).Cells(5)
                Exit Sub
            End If
            If GetGridValue(FG, FG.CurrentCell.RowIndex, 2) <> "" Then
                FG.CurrentCell = FG.Rows(FG.CurrentCell.RowIndex).Cells(6)
                Exit Sub
            End If
            'MsgBox("ໂຊຂໍ້ມູນແບບລະອຽດ")
        End If
        '====*************************




'*************************Col-6-*********************
        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 6 Then
            Dim rowIdx As Integer = FG.CurrentCell.RowIndex
            If IsNumeric(GetGridValue(FG, rowIdx, 6)) = False Then
                MessageBox.Show("ກະລຸນນາໃສ່ໂຕເລກ")
                SetGridValue(FG, rowIdx, 6, "0.00")
                Exit Sub
            End If
            If CDbl(GetGridValue(FG, rowIdx, 6)) = 0 Then
                MessageBox.Show("ກະລຸນນາມູນຄ່າກ່ອນ")
                SetGridValue(FG, rowIdx, 6, "0.00")
                Exit Sub
            End If
            SetGridValue(FG, rowIdx, 6, Format(CDbl(GetGridValue(FG, rowIdx, 6)), "#,##0.00"))
            SetGridValue(FG, rowIdx, 5, "0.00")
            SetGridValue(FG, rowIdx, 7, 0)
            SetGridValue(FG, rowIdx, 10, Cmb.Text)
            SetGridValue(FG, rowIdx, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
            SetGridValue(FG, rowIdx, 13, Format(CDbl(CDbl(txtRate.Text) * GetGridValue(FG, rowIdx, 6)), "#,##0.00"))
            SetGridValue(FG, rowIdx, 12, "0.00")
            SetGridValue(FG, rowIdx, 15, Format(CDbl(CDbl(GetGridValue(FG, rowIdx, 13))) / CDbl(MDUSD_LAK), "#,##0.00"))
            SetGridValue(FG, rowIdx, 14, "0.00")
            SumAmountDr()
            If CDbl(txtTotal_Amt_LAK.Text) > 0 Then
                If Chk_Preview.Checked = False Then


                    If CDbl(txtSumTotalAmountDr.Text) = CDbl(txtTotal_Amt_LAK.Text) Then
                        If CDbl(txtSumTotalAmountCr.Text) = CDbl(txtTotal_Amt_LAK.Text) Then
                            If CDbl(txtSumTotalAmountDr.Text) = CDbl(txtSumTotalAmountCr.Text) Then
                                'kkkkkkkkkkk
                                If MessageBox.Show("ບັນຊີນີ້ດູນດ່ຽງແລ້ວ ທ່ານຕອ້ງການບັນທຶຫລືບໍ່!", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                                    If txtInvoice.Enabled = True Then
                                        AutoNumber()
                                        If txtInvoice.Text = "" Then
                                            MsgBox("ກະລຸນນາໃສ່ເລກໃບຢັງຢືນກອ່ນ!", MsgBoxStyle.OkOnly)
                                            txtInvoice.BackColor = Color.Red
                                            txtInvoice.Focus()
                                            Exit Sub
                                        End If


                                        'xxxxxxxxxxxxxx

                                        Dim mNum As Integer = 0
                                        If IsNumeric(Microsoft.VisualBasic.Right(txtInvoice.Text, 3)) = False Then MsgBox("3 ໂຕທາງທ້າຍຕ້ອງເປັນໂຕເລກເທົານັ້ນ") : txtInvoice.BackColor = Color.Red : txtInvoice.Focus() : Exit Sub
                                        ' Using DbHelper instead of ADODB
                                        Dim dtCertify As DataTable = DbHelper.GetDataTable("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book =N'" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "'And  ReferNO = N'" & TxtReferno.Text & "' And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And  month(date_work)=" & Format(CDate(dtActi.Value), "MM") & " And  day(date_work)=" & Format(CDate(dtActi.Value), "dd") & "  Order by  Right(certify,3) DESC ")

                                        If dtCertify.Rows.Count = 0 Then
                                            mNum = 0
                                        Else
                                            mNum = Val(DbHelper.GetStr(dtCertify.Rows(0)("certify")))
                                        End If
                                        mNum = mNum + 1

                                        'If Int(Microsoft.VisualBasic.Right(txtInvoice.Text, 3)) > mNum Then
                                        '    If Len(CStr(mNum)) = 1 Then
                                        '        MsgBox("ເລກທ້າຍ 3 ໂຕບໍ່ໃຫ້ເກີນເລກ " & "00" & mNum)
                                        '    ElseIf Len(CStr(mNum)) = 2 Then
                                        '        MsgBox("ເລກທ້າຍ 3 ໂຕບໍ່ໃຫ້ເກີນເລກ " & "0" & mNum)
                                        '    ElseIf Len(CStr(mNum)) = 3 Then
                                        '        MsgBox("ເລກທ້າຍ 3 ໂຕບໍ່ໃຫ້ເກີນເລກ " & mNum)
                                        '    End If
                                        '    txtInvoice.BackColor = Color.Red
                                        '    txtInvoice.Focus()
                                        '    Exit Sub
                                        'End If
                                        'xxxxxxxxxxxxxxxxxxxxxxxxx

                                        'Call LoadSqlData("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book ='" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "' And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And month(date_work)=" & Format(CDate(dtActi.Value), "MM") & "  Order by  Right(certify,7) DESC ", RSC)
Dim dtCheck As DataTable = DbHelper.GetDataTable("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book =N'" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "'And  ReferNO = N'" & TxtReferno.Text & "' And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And  month(date_work)=" & Format(CDate(dtActi.Value), "MM") & " And  day(date_work)=" & Format(CDate(dtActi.Value), "dd") & "  Order by  Right(certify,3) DESC ")

                                        If dtCheck.Rows.Count > 0 Then
                                            MsgBox("ເລກລະຫັດ : " & Trim(txtInvoice.Text) & " ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
                                            txtInvoice.BackColor = Color.Red
                                            txtInvoice.Focus()
                                            Exit Sub
                                        End If

                                        Savedata()
                                    Else
                                        'DbHelper.ExecuteNonQuery("DELETE FROM gen_jn WHERE book ='" & GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 16) & "' And certify  = '" & txtInvoice.Text & "'   And  year(date_work)='" & Format(CDate(GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 1)), "yyyy") & "' ")
                                        DbHelper.ExecuteNonQuery("DELETE FROM gen_jn WHERE book =N'" & GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 16) & "' And certify  =N'" & txtInvoice.Text & "'  And  ReferNO = N'" & TxtReferno.Text & "'  And   date_work='" & Format(CDate(GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 1)), "yyyy-MM-dd") & "' ")

                                        Savedata()
                                    End If
                                    If CheckBox1.Checked Then
                                        Call LoadReport()
                                    End If
                                    txtInvoice.Enabled = True
                                    CmbBook.Enabled = True
                                    Panel1.Visible = False
                                    BtnMove.Visible = False
                                    BtnSearch.Visible = False
        FG.Rows.Clear()
        FG.Rows.Add()
        FG.Rows.Add()
                                    If FG.CurrentRow IsNot Nothing Then FG.CurrentRow.Index = 1
If FG.CurrentRow IsNot Nothing Then FG.CurrentCell = FG.Rows(1).Cells(1)

                                    If CheckBox2.Checked Then
                                        Close()
                                    End If
                                    Exit Sub
                                End If
                                'kkkkkkkkkkkkk
                            End If
                        End If
                    End If
                End If
                'aaaaaaaaaaaa

                If FG.CurrentRow IsNot Nothing AndAlso GetGridValue(FG, FG.CurrentRow.Index, 1) <> "" Then
                    If FG.CurrentRow.Index < FG.Rows.Count - 1 AndAlso GetGridValue(FG, FG.CurrentRow.Index + 1, 1) <> "" Then
                        FG.CurrentCell = FG.Rows(FG.CurrentRow.Index + 1).Cells(1)
                        Exit Sub
                    End If
                    If FG.CurrentRow.Index < FG.Rows.Count - 1 Then
                        FG.CurrentCell = FG.Rows(FG.CurrentRow.Index + 1).Cells(2)
                    End If
                    SumAmountDr()
                    '*****************
                Else
                    If FG.CurrentRow.Index < FG.Rows.Count - 1 AndAlso GetGridValue(FG, FG.CurrentRow.Index + 1, 2) <> "" Then
                        FG.CurrentCell = FG.Rows(FG.CurrentRow.Index + 1).Cells(2)
                        Exit Sub
                    End If
                    If FG.CurrentRow.Index < FG.Rows.Count - 1 Then
                        FG.CurrentCell = FG.Rows(FG.CurrentRow.Index + 1).Cells(1)
                    End If
                    SumAmountDr()
                End If

                Exit Sub
                    End If
                    ' FG.CurrentRow.Index is read-only, cannot assign
                    FG.CurrentCell = FG.Rows(1).Cells(2)
                    SumAmountDr()
                    '*****************
                Else
                    If GetGridValue(FG,FG.CurrentRow.Index + 1, 2) <> "" Then
                        ' FG.CurrentRow.Index is read-only, cannot assign
                        FG.CurrentCell = FG.Rows(1).Cells(2)
                        Exit Sub
                    End If
                    ' FG.CurrentRow.Index is read-only, cannot assign
                    FG.CurrentCell = FG.Rows(1).Cells(1)
                    SumAmountDr()
                End If

                Exit Sub
        CmbBook.Focus()
    End Sub
Private Sub AfterEdit2()
        BtnMove.Visible = False
        '*************************Col-1-*********************
        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 1 Then
            R = FG.CurrentCell.RowIndex
            L = FG.CurrentCell.ColumnIndex
            If GetGridValue(FG, R, 1) = "" Then
                MDSearchAcccode = GetGridValue(FG, R, 1)
                AccId = GetGridValue(FG, R, 1)

                fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                fmShartOfAccDetail.ShowDialog()

            End If
            If GetGridValue(FG, R, 1) <> "" Then
                AccId = GetGridValue(FG, R, 1)
                MDSearchAcccode = GetGridValue(FG, R, 1)

                LoadText()
                SetGridValue(FG, R, 3, AccName)
                SetGridValue(FG, R, 4, AccNamee)
                If GetGridValue(FG, R, 3) = "" Then
                    SetGridValue(FG, R, 1, "")
                    fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                    fmShartOfAccDetail.ShowDialog()
                    LoadText()
                End If
                SumAmountDr()
                txtSumAmountDr.Text = CDbl(txtSumAmountDr.Text) - CDbl(GetGridValue(FG, R, 5))
                SetGridValue(FG, R, 5, Format(CDbl(CDbl(txtAmount.Text) - CDbl(txtSumAmountDr.Text)), "#,##0.00"))
                SetGridValue(FG, R, 6, "0.00")
                SetGridValue(FG, R, 7, 0)
                SetGridValue(FG, R, 10, Cmb.Text)
                SetGridValue(FG, R, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
                SetGridValue(FG, R, 12, Format(CDbl(CDbl(txtRate.Text) * CDbl(GetGridValue(FG, R, 5))), "#,##0.00"))
                SetGridValue(FG, R, 13, "0.00")
                SetGridValue(FG, R, 14, Format(CDbl(CDbl(GetGridValue(FG, R, 12))) / CDbl(MDUSD_LAK), "#,##0.00"))
                SetGridValue(FG, R, 15, "0.00")
                SumAmountDr()

                If ChDe.Checked = True Then
                    If MuLng = "L" Then
                        FG.CurrentCell = FG.Rows(R).Cells(3)
                    End If
                    If MuLng = "E" Then
                        FG.CurrentCell = FG.Rows(R).Cells(4)
                    End If

                    If GetGridValue(FG, FG.Rows.Count - 1, 3) <> "" Then
                        FG.Rows.Add()
                    End If
                    Exit Sub
                End If

                FG.CurrentCell = FG.Rows(R).Cells(5)
                If GetGridValue(FG, FG.Rows.Count - 1, 3) <> "" Then
                    FG.Rows.Add()
                    Exit Sub
                End If
                Exit Sub
            End If
            Exit Sub
        End If
'*************************Col-2-*********************
        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 2 Then
            R = FG.CurrentCell.RowIndex
            L = FG.CurrentCell.ColumnIndex
            If GetGridValue(FG, R, 2) = "" Then
                MDSearchAcccode = GetGridValue(FG, R, 2)
                fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                fmShartOfAccDetail.ShowDialog()
            End If
            If GetGridValue(FG, R, 2) <> "" Then
                AccId = GetGridValue(FG, R, 2)
                LoadText()
                SetGridValue(FG, R, 3, AccName)
                SetGridValue(FG, R, 4, AccNamee)
                If GetGridValue(FG, R, 3) = "" Then
                    MDSearchAcccode = GetGridValue(FG, R, 2)
                    fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                    fmShartOfAccDetail.ShowDialog()
                End If
                SumAmountDr()
                txtSumAmountCr.Text = CDbl(txtSumAmountCr.Text) - CDbl(GetGridValue(FG, R, 6))
                SetGridValue(FG, R, 6, Format(CDbl(CDbl(txtAmount.Text) - CDbl(txtSumAmountCr.Text)), "#,##0.00"))
                SetGridValue(FG, R, 5, "0.00")
                SetGridValue(FG, R, 7, 0)
                SetGridValue(FG, R, 10, Cmb.Text)
                SetGridValue(FG, R, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
                SetGridValue(FG, R, 13, Format(CDbl(CDbl(txtRate.Text) * CDbl(GetGridValue(FG, R, 6))), "#,##0.00"))
                SetGridValue(FG, R, 12, "0.00")
                SetGridValue(FG, R, 15, Format(CDbl(CDbl(GetGridValue(FG, R, 13))) / CDbl(MDUSD_LAK), "#,##0.00"))
                SetGridValue(FG, R, 14, "0.00")
                SumAmountDr()

                If ChDe.Checked = True Then
                    If MuLng = "L" Then
                        FG.CurrentCell = FG.Rows(R).Cells(3)
                    End If
                    If MuLng = "E" Then
                        FG.CurrentCell = FG.Rows(R).Cells(4)
                    End If
                    If GetGridValue(FG, FG.Rows.Count - 1, 3) <> "" Then
                        FG.Rows.Add()
                    End If
                    Exit Sub
                End If
                FG.CurrentCell = FG.Rows(R).Cells(6)
                If GetGridValue(FG, FG.Rows.Count - 1, 3) <> "" Then
                    FG.Rows.Add()
                    FG.CurrentCell = FG.Rows(R).Cells(6)
                End If
                Exit Sub
            End If
            Exit Sub
        End If

        '====*************************

If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 3 Then
            If Button3.Text = "ໂຊຂໍ້ມູນແບບທົ່ວໄປ" Then
                FG.CurrentCell = FG.Rows(FG.CurrentCell.RowIndex).Cells(4)
                Exit Sub
            Else
                If GetGridValue(FG, FG.CurrentCell.RowIndex, 1) <> "" Then
                    FG.CurrentCell = FG.Rows(FG.CurrentCell.RowIndex).Cells(5)
                    Exit Sub
                End If
                If GetGridValue(FG, FG.CurrentCell.RowIndex, 2) <> "" Then
                    FG.CurrentCell = FG.Rows(FG.CurrentCell.RowIndex).Cells(6)
                    Exit Sub
                End If
                'MsgBox("ໂຊຂໍ້ມູນແບບລະອຽດ")
            End If
        End If
        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 4 Then
            If GetGridValue(FG, FG.CurrentCell.RowIndex, 1) <> "" Then
                FG.CurrentCell = FG.Rows(FG.CurrentCell.RowIndex).Cells(5)
                Exit Sub
            End If
            If GetGridValue(FG, FG.CurrentCell.RowIndex, 2) <> "" Then
                FG.CurrentCell = FG.Rows(FG.CurrentCell.RowIndex).Cells(6)
                Exit Sub
            End If
            'MsgBox("ໂຊຂໍ້ມູນແບບລະອຽດ")
        End If

'*************************Col-6-*********************
        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 6 Then
            rowIdx = FG.CurrentCell.RowIndex
            If IsNumeric(GetGridValue(FG, rowIdx, 6)) = False Then
                MessageBox.Show("ກະລຸນນາໃສ່ໂຕເລກ")
                SetGridValue(FG, rowIdx, 6, "0.00")
                Exit Sub
            End If
            If CDbl(GetGridValue(FG, rowIdx, 6)) = 0 Then
                MessageBox.Show("ກະລຸນນາມູນຄ່າກ່ອນ")
                SetGridValue(FG, rowIdx, 6, "0.00")
                Exit Sub
            End If
            ACCode = GetGridValue(FG, rowIdx, 2)
            SetGridValue(FG, rowIdx, 6, Format(CDbl(GetGridValue(FG, rowIdx, 6)), "#,##0.00"))
            SetGridValue(FG, rowIdx, 5, "0.00")
            SetGridValue(FG, rowIdx, 7, 0)
            SetGridValue(FG, rowIdx, 10, Cmb.Text)
            SetGridValue(FG, rowIdx, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
            SetGridValue(FG, rowIdx, 13, Format(CDbl(CDbl(txtRate.Text) * CDbl(GetGridValue(FG, rowIdx, 6))), "#,##0.00"))
            SetGridValue(FG, rowIdx, 12, "0.00")
            SetGridValue(FG, rowIdx, 15, Format(CDbl(CDbl(GetGridValue(FG, rowIdx, 13))) / CDbl(MDUSD_LAK), "#,##0.00"))
            SetGridValue(FG, rowIdx, 14, "0.00")
            SumAmountDr()
            'Remain = CDbl(GetGridValue(FG, rowIdx, 13))
            If CDbl(txtTotal_Amt_LAK.Text) > 0 Then
                If CDbl(txtTotal_Amt_LAK.Text) > 0 Then
                    If Chk_Preview.Checked = False Then
                        If CDbl(txtSumTotalAmountDr.Text) = CDbl(txtTotal_Amt_LAK.Text) Then
                            If CDbl(txtSumTotalAmountCr.Text) = CDbl(txtTotal_Amt_LAK.Text) Then
                                If CDbl(txtSumTotalAmountDr.Text) = CDbl(txtSumTotalAmountCr.Text) Then
                                    'kkkkkkkkkkk


                                    'If Microsoft.VisualBasic.Left(ACCode, 3) = "110" Or Microsoft.VisualBasic.Left(ACCode, 3) = "112" Or Microsoft.VisualBasic.Left(ACCode, 3) = "113" Then
                                    '    Call Load_Gen_Jn()
                                    '    Remain = CDbl(CDbl(Open_jn) + CDbl(SumDr)) - CDbl(SumCr)

                                    '    If Remain >= 0 Then
                                    '        Remain = Format(CDbl(Remain), "##,##0.00")
                                    '    Else
                                    '        Remain = "(" & Format(CDbl(Remain * (-1)), "##,##0.00") & ")"
                                    '    End If
                                    '    SumDr = CDbl(GetGridValue(FG, rowIdx, 13))
                                    '    If CDbl(SumDr) > CDbl(Remain) Then
                                    '        SetGridValue(FG, rowIdx, 13, Format(CDbl(Remain), "#,##0.00"))
                                    '        SetGridValue(FG, rowIdx, 6, Format(CDbl(Remain) / CDbl(txtRate.Text), "#,##0.00"))
                                    '    Else
                                    '    End If

                                    'End If




                                    If MessageBox.Show("ບັນຊີນີ້ດູນດ່ຽງແລ້ວ ທ່ານຕອ້ງການບັນທຶຫລືບໍ່!", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                                        If txtInvoice.Enabled = True Then
                                            AutoNumber()

                                            If txtInvoice.Text = "" Then
                                                MsgBox("ກະລຸນນາໃສ່ເລກໃບຢັງຢືນກອ່ນ!", MsgBoxStyle.OkOnly)
                                                txtInvoice.BackColor = Color.Red
                                                txtInvoice.Focus()
                                                Exit Sub
                                            End If

                                            'xxxxxxxxxxxxxx

Dim mNum As Integer = 0
                                            If IsNumeric(Microsoft.VisualBasic.Right(txtInvoice.Text, 3)) = False Then MsgBox("3 ໂຕທາງທ້າຍຕ້ອງເປັນໂຕເລກເທົານັ້ນ") : txtInvoice.BackColor = Color.Red : txtInvoice.Focus() : Exit Sub
                                            Dim dtNum As DataTable = DbHelper.GetDataTable("SELECT  top 1 Right(certify,3) As  certify  FROM Gen_jn WHERE   book ='" & CmbBook.Text & "' And    year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And month(date_work)=" & Format(CDate(dtActi.Value), "MM") & "  Order by  Right(certify,7) DESC ")
                                            If dtNum.Rows.Count = 0 Then
                                                mNum = 0
                                            Else
                                                mNum = Val(dtNum.Rows(0)("certify").ToString())
                                            End If
                                            mNum = mNum + 1

                                            'If Int(Microsoft.VisualBasic.Right(txtInvoice.Text, 3)) > mNum Then
                                            '    If Len(CStr(mNum)) = 1 Then
                                            '        MsgBox("ເລກທ້າຍ 3 ໂຕບໍ່ໃຫ້ເກີນເລກ " & "00" & mNum)
                                            '    ElseIf Len(CStr(mNum)) = 2 Then
                                            '        MsgBox("ເລກທ້າຍ 3 ໂຕບໍ່ໃຫ້ເກີນເລກ " & "0" & mNum)
                                            '    ElseIf Len(CStr(mNum)) = 3 Then
                                            '        MsgBox("ເລກທ້າຍ 3 ໂຕບໍ່ໃຫ້ເກີນເລກ " & mNum)
                                            '    End If
                                            '    txtInvoice.BackColor = Color.Red
                                            '    txtInvoice.Focus()
                                            '    Exit Sub
                                            'End If
                                            'xxxxxxxxxxxxxxxxxxxxxxxxx


                                            Dim dtCheck As DataTable = DbHelper.GetDataTable("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book ='" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "' And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And  month(date_work)=" & Format(CDate(dtActi.Value), "MM") & " Order by  Right(certify,7) DESC ")
                                            If dtCheck.Rows.Count > 0 Then
                                                MsgBox("ເລກລະຫັດ : " & Trim(txtInvoice.Text) & " ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
                                                txtInvoice.BackColor = Color.Red
                                                txtInvoice.Focus()
                                                Exit Sub
                                            End If

                                            Savedata()
                                        Else
                                            DbHelper.ExecuteNonQuery("DELETE FROM gen_jn WHERE book ='" & GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 16) & "' And certify  = '" & txtInvoice.Text & "'   And  year(date_work)='" & Format(CDate(GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 1)), "yyyy") & "' ")
                                            Savedata()
                                        End If

                                        If CheckBox1.Checked Then

                                            Call LoadReport()
                                            'MsgBox(11)
                                        End If

                                        txtInvoice.Enabled = True
                                        CmbBook.Enabled = True
                                        Panel1.Visible = False
                                        BtnMove.Visible = False
                                        BtnSearch.Visible = False
FG.Rows.Clear()
        FG.Rows.Add()
        FG.Rows.Add()
                                        If FG.CurrentRow IsNot Nothing Then FG.CurrentRow.Index = 1
If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 1 Then
                                        CmbBook.Focus()


                                        If CheckBox2.Checked Then

                                            Close()
                                        End If
                                        Exit Sub
                                    End If
                                    'kkkkkkkkkkkkk
                                End If
                            End If
                        End If
                    End If

                End If

                If FG.CurrentRow IsNot Nothing AndAlso GetGridValue(FG, FG.CurrentRow.Index, 1) <> "" Then
                    If GetGridValue(FG,FG.CurrentRow.Index + 1, 1) <> "" Then
                        ' FG.CurrentRow.Index is read-only, cannot assign
                        FG.CurrentCell = FG.Rows(1).Cells(1)
                        Exit Sub
                    End If
                    ' FG.CurrentRow.Index is read-only, cannot assign
                    FG.CurrentCell = FG.Rows(1).Cells(2)
                    SumAmountDr()
                    '*****************
                Else
                    If GetGridValue(FG,FG.CurrentRow.Index + 1, 2) <> "" Then
                        ' FG.CurrentRow.Index is read-only, cannot assign
                        FG.CurrentCell = FG.Rows(1).Cells(2)
                        Exit Sub
                    End If
                    ' FG.CurrentRow.Index is read-only, cannot assign
                    FG.CurrentCell = FG.Rows(1).Cells(1)
                    SumAmountDr()
                End If
                Exit Sub
            End If
        End If
        '*************************Col-4*********************
If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 5 Then
            rowIdx = FG.CurrentCell.RowIndex
            If IsNumeric(GetGridValue(FG, rowIdx, 5)) = False Then
                MessageBox.Show("ກະລຸນນາໃສ່ໂຕເລກ")
                SetGridValue(FG, rowIdx, 5, "0.00")
                Exit Sub
            End If

            If CDbl(GetGridValue(FG, rowIdx, 5)) = 0 Then
                MessageBox.Show("ກະລຸນນາມູນຄ່າກ່ອນ")
                SetGridValue(FG, rowIdx, 5, "0.00")
                Exit Sub
            End If
            ACCode = GetGridValue(FG, rowIdx, 1)
            SetGridValue(FG, rowIdx, 5, Format(CDbl(GetGridValue(FG, rowIdx, 5)), "#,##0.00"))
            SetGridValue(FG, rowIdx, 6, "0.00")
            SetGridValue(FG, rowIdx, 7, 0)
            SetGridValue(FG, rowIdx, 10, Cmb.Text)
            SetGridValue(FG, rowIdx, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
            SetGridValue(FG, rowIdx, 12, Format(CDbl(CDbl(txtRate.Text) * CDbl(GetGridValue(FG, rowIdx, 5))), "#,##0.00"))
            SetGridValue(FG, rowIdx, 13, "0.00")
            SetGridValue(FG, rowIdx, 14, Format(CDbl(CDbl(GetGridValue(FG, rowIdx, 12))) / CDbl(MDUSD_LAK), "#,##0.00"))
            SetGridValue(FG, rowIdx, 15, "0.00")
            SumAmountDr()
            If CDbl(txtTotal_Amt_LAK.Text) > 0 Then
                If Chk_Preview.Checked = False Then
                    If CDbl(txtSumTotalAmountDr.Text) = CDbl(txtTotal_Amt_LAK.Text) Then
                        If CDbl(txtSumTotalAmountCr.Text) = CDbl(txtTotal_Amt_LAK.Text) Then
                            If CDbl(txtSumTotalAmountDr.Text) = CDbl(txtSumTotalAmountCr.Text) Then
                                If MessageBox.Show("ບັນຊີນີ້ດູນດ່ຽງແລ້ວ ທ່ານຕອ້ງການບັນທຶຫລືບໍ່!", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                                    If txtInvoice.Enabled = True Then
                                        AutoNumber()

                                        If txtInvoice.Text = "" Then
                                            MsgBox("ກະລຸນນາໃສ່ເລກໃບຢັງຢືນກອ່ນ!", MsgBoxStyle.OkOnly)
                                            txtInvoice.BackColor = Color.Red
                                            txtInvoice.Focus()
                                            Exit Sub
                                        End If

                                        'xxxxxxxxxxxxxx

Dim mNum As Integer = 0
                                        If IsNumeric(Microsoft.VisualBasic.Right(txtInvoice.Text, 3)) = False Then MsgBox("3 ໂຕທາງທ້າຍຕ້ອງເປັນໂຕເລກເທົານັ້ນ") : txtInvoice.BackColor = Color.Red : txtInvoice.Focus() : Exit Sub
                                        Dim dtNum As DataTable = DbHelper.GetDataTable("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book ='" & CmbBook.Text & "' And year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And Month(date_work)=" & Format(CDate(dtActi.Value), "MM") & "  Order by  Right(certify,7) DESC ")
                                        If dtNum.Rows.Count = 0 Then
                                            mNum = 0
                                        Else
                                            mNum = Val(dtNum.Rows(0)("certify").ToString())
                                        End If
                                        mNum = mNum + 1

                                        'If Int(Microsoft.VisualBasic.Right(txtInvoice.Text, 3)) > mNum Then
                                        '    If Len(CStr(mNum)) = 1 Then
                                        '        MsgBox("ເລກທ້າຍ 3 ໂຕບໍ່ໃຫ້ເກີນເລກ " & "00" & mNum)
                                        '    ElseIf Len(CStr(mNum)) = 2 Then
                                        '        MsgBox("ເລກທ້າຍ 3 ໂຕບໍ່ໃຫ້ເກີນເລກ " & "0" & mNum)
                                        '    ElseIf Len(CStr(mNum)) = 3 Then
                                        '        MsgBox("ເລກທ້າຍ 3 ໂຕບໍ່ໃຫ້ເກີນເລກ " & mNum)
                                        '    End If
                                        '    txtInvoice.BackColor = Color.Red
                                        '    txtInvoice.Focus()
                                        '    Exit Sub
                                        'End If
                                        'xxxxxxxxxxxxxxxxxxxxxxxxx

                                        If MdCertifyAuto = 1 Then
                                            If txtInvoice.Enabled = True Then
                                                AutoNumber()
                                            End If
                                        End If

                                        Dim dtCheck2 As DataTable = DbHelper.GetDataTable("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book ='" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "' And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And Month(date_work)=" & Format(CDate(dtActi.Value), "MM") & " Order by  Right(certify,7) DESC ")
                                        If dtCheck2.Rows.Count > 0 Then
                                            MsgBox("ເລກລະຫັດ : " & Trim(txtInvoice.Text) & " ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
                                            txtInvoice.BackColor = Color.Red
                                            txtInvoice.Focus()
                                            Exit Sub
                                        End If


                                        Savedata()
                                    Else

                                        DbHelper.ExecuteNonQuery("DELETE FROM gen_jn WHERE book ='" & GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 16) & "' And certify  = '" & txtInvoice.Text & "'   And  year(date_work)='" & Format(CDate(GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 1)), "yyyy") & "' ")
                                        Savedata()
                                    End If
                                    If CheckBox1.Checked Then
                                        Call LoadReport()
                                    End If
                                    txtInvoice.Enabled = True
                                    CmbBook.Enabled = True
                                    Panel1.Visible = False
                                    BtnMove.Visible = False
                                    BtnSearch.Visible = False
                                    FG.Rows.Clear()
                                    FG.Rows.Add()
                                    FG.Rows.Add()
                                    If FG.CurrentRow IsNot Nothing Then FG.CurrentCell = FG.Rows(1).Cells(1)
                                    Exit Sub

                                    If CheckBox2.Checked Then
                                        Close()
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                If FG.CurrentRow IsNot Nothing AndAlso GetGridValue(FG, FG.CurrentRow.Index, 1) <> "" Then
                    If GetGridValue(FG,FG.CurrentRow.Index + 1, 1) <> "" Then
                        ' FG.CurrentRow.Index is read-only, cannot assign
                        FG.CurrentCell = FG.Rows(1).Cells(1)
                        Exit Sub
                    End If
                    ' FG.CurrentRow.Index is read-only, cannot assign
                    FG.CurrentCell = FG.Rows(1).Cells(2)
                    SumAmountDr()
                    '*****************
                Else
                    If GetGridValue(FG,FG.CurrentRow.Index + 1, 2) <> "" Then
                        ' FG.CurrentRow.Index is read-only, cannot assign
                        FG.CurrentCell = FG.Rows(1).Cells(2)
                        Exit Sub
                    End If
                    ' FG.CurrentRow.Index is read-only, cannot assign
                    FG.CurrentCell = FG.Rows(1).Cells(1)
                    SumAmountDr()
                End If
                Exit Sub
            End If
        End If
        '*************************Col-7-*********************
        End If
        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 7 Then
            Exit Sub
        End If

    End Sub
    Public Sub AutoNumber()
        Dim dtCheck As DataTable
        Dim mNum As Integer

        ' Using ADO.NET instead of ADODB
        Dim ss As String = ""
        ss = "SELECT top 1 Right(certify,3) As  certify   FROM  gen_jn where Frm=0 and book =N'" & CmbBook.Text & "' And  year(date_work)='" & Format(dtActi.Value, "yyyy") & "'  " & _
        " And  month(date_work)='" & Format(dtActi.Value, "MM") & "'  and LEFT(company,2)='" & Off_Id & "' Order by  Right(certify,3) DESC"
        dtCheck = DbHelper.GetDataTable(ss)
        If dtCheck.Rows.Count = 0 Then
            MdCertifyId = "001"
        Else
            mNum = Val(dtCheck.Rows(0)("certify").ToString)
            mNum = CDbl(mNum) + 1
            If Len(CStr(mNum).Trim) = 1 Then
                MdCertifyId = "00" & CStr(mNum)
            ElseIf Len(CStr(mNum).Trim) = 2 Then
                MdCertifyId = "0" & CStr(mNum)
            ElseIf Len(CStr(mNum).Trim) = 3 Then
                MdCertifyId = CStr(mNum)
            End If
        End If
        If MdCertifyAuto = 1 Then
            txtInvoice.Text = Trim(CmbBook.Text) & Format(dtActi.Value, "yyMM") & MdCertifyId
        End If


    End Sub




    Public Sub Autox()

        ' ''Dim mNum As Integer
        'If txtInvoice.Enabled = True Then
        '    If MdCertifyAuto = 0 Then
        '        If txtAuto.Text <> "" Then


        '            txtAuto.Text = CDbl(txtAuto.Text) + 1



        '            If Len(CStr(txtAuto.Text).Trim) = 1 Then
        '                txtInvoice.Text = CmbBook.Text & "000000" & CStr(txtAuto.Text)
        '            ElseIf Len(CStr(txtAuto.Text).Trim) = 2 Then
        '                txtInvoice.Text = CmbBook.Text & "00000" & CStr(txtAuto.Text)
        '            ElseIf Len(CStr(txtAuto.Text).Trim) = 3 Then
        '                txtInvoice.Text = CmbBook.Text & "0000" & CStr(txtAuto.Text)
        '            ElseIf Len(CStr(txtAuto.Text).Trim) = 4 Then
        '                txtInvoice.Text = CmbBook.Text & "000" & CStr(txtAuto.Text)
        '            ElseIf Len(CStr(txtAuto.Text).Trim) = 5 Then
        '                txtInvoice.Text = CmbBook.Text & "00" & CStr(txtAuto.Text)
        '            ElseIf Len(CStr(txtAuto.Text).Trim) = 6 Then
        '                txtInvoice.Text = CmbBook.Text & CStr(txtAuto.Text)
        '            End If
        '        End If
        '    End If
        'End If



    End Sub



Public Sub Foucus()
        FG.Focus()
        If FG.CurrentRow IsNot Nothing Then FG.CurrentCell = FG.Rows(R).Cells(L + 4)
    End Sub
    Private Sub NewText()
        dtActi.Text = MWorkSetting
        txtInvoice.Text = ""
        txtDesc.Text = ""
        txtDescE.Text = ""
        TxtReferno.Text = ""
        'CheckBox4.Checked = False
        txtAmount.Text = 0
        SumAmountDr()
    End Sub
    Public Sub LoadCurr()
        Dim dtCurr As DataTable = DbHelper.GetDataTable("SELECT Curr FROM Ap_RateSeting WHERE Curr <> '' order by Curr")
        Cmb.Items.Clear()
        If dtCurr.Rows.Count <> 0 Then
            For Each row As DataRow In dtCurr.Rows
                Cmb.Items.Add(Trim(row("Curr").ToString()))
            Next
        End If
        If CmbBook.Enabled = True Then
            Cmb.SelectedIndex = 0
        End If

    End Sub

    Private Sub FmNsewJeneralJournal_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Call FmJeneralJournal_List.LoadMonthSQL()
    End Sub
    Private Sub loadOffice_User()
        Off_Usr.Items.Clear()
        Dim dtOffice As DataTable
        If MPermit = "User" Then
            dtOffice = DbHelper.GetDataTable("select sub_id , off_add2  from  Ap_office Where left(sub_id,2) = '" & Off_Id & "' And Substring(sub_id,4,2) <> '00'  Order by sub_id")
        Else
            dtOffice = DbHelper.GetDataTable("select sub_id , off_add2  from  Ap_office Where sub_id <> '00-00' And Substring(sub_id,4,2) <> '00'  Order by sub_id")
        End If
        
        For Each row As DataRow In dtOffice.Rows
            Off_Usr.Items.Add((row("sub_id").ToString()) & " " & row("off_add2").ToString())
        Next
        Off_Usr.SelectedIndex = 0
        'Off_Usr.SelectedIndex = 0
    End Sub
    Private Sub FmNsewJeneralJournal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If MuLng = "E" Then
            txtAmt_letter.Visible = False
            txtAmt_letter_E.Visible = True
FG.CurrentCell.ColumnIndexumns(3).Visible = False
            FG.CurrentCell.ColumnIndexumns(4).Visible = True
        Else
FG.CurrentCell.ColumnIndexumns(3).Visible = True
            FG.CurrentCell.ColumnIndexumns(4).Visible = False
            txtAmt_letter.Visible = True
            txtAmt_letter_E.Visible = False
        End If
        Call loadOffice_User()
        txtInvoice.BackColor = Color.White
        ' FG.BackColorFixed not available in DataGridView

        'LoadCurr()
        If MdCertifyAuto = 1 Then
            txtInvoice.ReadOnly = True
        Else
            txtInvoice.Text = ""
            txtInvoice.ReadOnly = False
        End If

        BtnSearch.Visible = False
        BtnMove.Visible = False

        ' FG.FormatString not available in DataGridView = "^ລ/ດ |< ເລກບັນຊີໜີ           |< ເລກບັນຊີມີ           |< ຊື່ບັນຊີ (ລາວ)                                            |< ຊື່ບັນຊີ (ອັງກິດ)                     |> ຈຳນວນເງິນຈົດໜີ້        |> ຈຳນວນເງິນຈົດມີ     |^ລະຫັດ|< ຕົ້ນທຶນພາສາ (ລາວ)   |< ຕົ້ນທຶນພາສາ (ອັງກິດ)             |< ສະກຸນເງິນເງິນ |> ອັດຕາແລກປ່ຽນ |> ມູນຄ່າໜີ້          |> ມູນຄ່າມີ            |> 1111    |> 22       "
        'Fg2.FormatString = "^|>       |<    |<                                   "
        'Fg2.FormatString = "^|<                        "
        'FG.Size = New System.Drawing.Point(962, 287)
        If CmbBook.Enabled = True Then
            dtActi.Enabled = True
            curr_Last.Text = "Kip"
            txtRate.Text = "1.00"
            Cmb.Text = "LAK"
        Else
            'dtActi.Enabled = False
        End If

        LoadBook()
If FG.CurrentRow IsNot Nothing Then FG.CurrentCell = FG.Rows(1).Cells(1)
        FG.BackColorSel = Color.White

        LoadTableId()
        Panel1.Visible = False
        loadRate()
        Call RateSetting()

        Cmb.Items.Clear()
        Call load_Cmb(" SELECT Curr  FROM Curr_For_Rate  ORDER BY cnt ", "Curr", Cmb)
        If Cmb.Items.Count > 0 Then
            Cmb.SelectedIndex = 0
        End If

        CmbCust.Items.Clear()
        Call load_Cmb(" SELECT Name  FROM Customer  ORDER BY cnt ", "Name", CmbCust)
        If CmbCust.Items.Count > 0 Then
            CmbCust.SelectedIndex = 0
        End If

        CmbSupp.Items.Clear()
        Call load_Cmb(" SELECT Name  FROM Supplier  ORDER BY cnt ", "Name", CmbSupp)
        If CmbSupp.Items.Count > 0 Then
            CmbSupp.SelectedIndex = 0
        End If

        Call RateSetting()

        If txtInvoice.Enabled = False Then
            txtInvoice.Text = MDInvoiceNo
            LoadSQL()
            LoadListFG()
            FG.Rows.Count = FG.Rows.Count + 1
        Else
FG.Rows.Clear()
        FG.Rows.Add()
        FG.Rows.Add()
            AutoNumber()
            Call NewText()
            AutoNumber()
        End If

        Call FormatText()


        Button3.Text = "Show GLN"
        Call ShowList()
        Button3.Text = "Show All"
        Button7.Text = "Connect SerVer"

ChCat_ID.Checked = False
        For J = 0 To FG.Rows.Count - 1
            If Trim(GetGridValue(FG, J, 3)) <> "" Then
                If ChCat_ID.Checked = True Then
                    FG.Rows(J).Cells(7).Style.BackColor = Color.LightCyan
                Else
                    FG.Rows(J).Cells(7).Style.BackColor = Color.White
                    FG.Rows(J).Cells(8).Style.ForeColor = Color.Gray
                    FG.Rows(J).Cells(9).Style.ForeColor = Color.Gray

                    SetGridValue(FG, J, 7, "0")
                    SetGridValue(FG, J, 8, "ບໍ່ເລືອກ")
                    SetGridValue(FG, J, 9, "No Selete")
                End If
            End If
        Next J
        Call loadColor()
        'FG.AllowUserResizing = VSFlex8U.AllowUserResizeSettings.flexResizeBoth
        ' FG.ExtendLastCol not available in DataGridView
        SetControlText(Me)
        CheckBox4.Text = "Only Customer/Supplier"
        If MuLng = "E" Then
            Label24.Text = "Ref No. :"
        Else
            Label24.Text = "ເອກະສານ :"
        End If
        Button7.Text = "Connect SerVer"
        Chk_Preview.Text = "To be continued"

        '' FG.FormatString not available in DataGridView = "^ລ/ດ |< ເລກບັນຊີໜີ           |< ເລກບັນຊີມີ           |< ຊື່ບັນຊີ (ລາວ)                                            |< ຊື່ບັນຊີ (ອັງກິດ)                     |> ຈຳນວນເງິນຈົດໜີ້        |> ຈຳນວນເງິນຈົດມີ     |^ລະຫັດ|< ຕົ້ນທຶນພາສາ (ລາວ)   |< ຕົ້ນທຶນພາສາ (ອັງກິດ)             |< ສະກຸນເງິນເງິນ |> ອັດຕາແລກປ່ຽນ |> ມູນຄ່າໜີ້          |> ມູນຄ່າມີ            "
    End Sub

    Private Sub LoadBook()
        Dim dtBooks As DataTable = DbHelper.GetDataTable("SELECT * FROM books WHERE bookid <> '' ")
        CmbBook.Items.Clear()
        If dtBooks.Rows.Count <> 0 Then
            For Each row As DataRow In dtBooks.Rows
                CmbBook.Items.Add(Trim(row("bookid").ToString()))
            Next
        End If
        CmbBook.Text = "GL"
        Dim dtBook As DataTable = DbHelper.GetDataTable("SELECT * FROM books WHERE bookid = N'" & CmbBook.Text & "'")
        If dtBook.Rows.Count > 0 Then
            txtBookName.Text = Trim(dtBook.Rows(0)("bookname").ToString())
        End If
 

    End Sub

Public Sub AddAcc2()

        BtnMove.Visible = False
        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 1 Then
            rowIdx = FG.CurrentCell.RowIndex
            SetGridValue(FG, rowIdx, 6, "0.00")
            SetGridValue(FG, rowIdx, 7, 0)
            SetGridValue(FG, rowIdx, 10, Cmb.Text)
            SetGridValue(FG, rowIdx, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
            SetGridValue(FG, rowIdx, 12, Format(CDbl(CDbl(txtRate.Text) * CDbl(GetGridValue(FG, rowIdx, 4))), "#,##0.00"))
            SetGridValue(FG, rowIdx, 13, "0.00")
            SumAmountDr()
            If ChDe.Checked = True Then
                If MuLng = "L" Then
                    FG.CurrentCell = FG.Rows(rowIdx).Cells(3)
                End If
                If MuLng = "E" Then
                    FG.CurrentCell = FG.Rows(rowIdx).Cells(4)
                End If
                If GetGridValue(FG, FG.Rows.Count - 1, 3) <> "" Then
                    FG.Rows.Add()
                End If
                Exit Sub
            End If
            FG.CurrentCell = FG.Rows(rowIdx).Cells(5)
            'Call loadColor()
        End If
        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 2 Then
            rowIdx = FG.CurrentCell.RowIndex
            AccId = GetGridValue(FG, rowIdx, 2)
            LoadText()
            SetGridValue(FG, rowIdx, 3, AccName)
            MDSearchAcccode = GetGridValue(FG, rowIdx, 2)
            SetGridValue(FG, rowIdx, 5, "0.00")
            SetGridValue(FG, rowIdx, 7, 0)
            SetGridValue(FG, rowIdx, 10, Cmb.Text)
            SetGridValue(FG, rowIdx, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
            SetGridValue(FG, rowIdx, 13, Format(CDbl(CDbl(txtRate.Text) * CDbl(GetGridValue(FG, rowIdx, 6))), "#,##0.00"))
            SetGridValue(FG, rowIdx, 12, "0.00")
            SumAmountDr()
            If ChDe.Checked = True Then
                If MuLng = "L" Then
                    FG.CurrentCell = FG.Rows(rowIdx).Cells(3)
                End If
                If MuLng = "E" Then
                    FG.CurrentCell = FG.Rows(rowIdx).Cells(4)
                End If
                If GetGridValue(FG, FG.Rows.Count - 1, 3) <> "" Then
                    FG.Rows.Add()
                End If
                Exit Sub
            End If
            FG.CurrentCell = FG.Rows(rowIdx).Cells(5)
            'Call loadColor()
        End If

        SumAmountDr()

    End Sub
Public Sub AddAcc()

        BtnMove.Visible = False
        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 1 Then
            rowIdx = FG.CurrentCell.RowIndex

            txtSumAmountDr.Text = CDbl(txtSumAmountDr.Text) - CDbl(GetGridValue(FG, rowIdx, 5))
            SetGridValue(FG, rowIdx, 5, Format(CDbl(CDbl(txtAmount.Text) - CDbl(txtSumAmountDr.Text)), "#,##0.00"))
            SetGridValue(FG, rowIdx, 6, "0.00")
            SetGridValue(FG, rowIdx, 7, 0)
            SetGridValue(FG, rowIdx, 10, Cmb.Text)
            SetGridValue(FG, rowIdx, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
            SetGridValue(FG, rowIdx, 12, Format(CDbl(CDbl(txtRate.Text) * CDbl(GetGridValue(FG, rowIdx, 5))), "#,##0.00"))
            SetGridValue(FG, rowIdx, 13, "0.00")
            SumAmountDr()
            If ChDe.Checked = True Then
                If MuLng = "L" Then
                    FG.CurrentCell = FG.Rows(rowIdx).Cells(3)
                End If
                If MuLng = "E" Then
                    FG.CurrentCell = FG.Rows(rowIdx).Cells(4)
                End If
                If GetGridValue(FG, FG.Rows.Count - 1, 3) <> "" Then
                    FG.Rows.Add()
                End If
                Exit Sub
            End If

            FG.CurrentCell = FG.Rows(rowIdx).Cells(5)

            If GetGridValue(FG, FG.Rows.Count - 1, 3) <> "" Then
                FG.Rows.Add()
                'Timer1.Enabled = True

            End If
            'Call loadColor()
        End If
        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 2 Then
            rowIdx = FG.CurrentCell.RowIndex

            AccId = GetGridValue(FG, rowIdx, 2)
            LoadText()
            SetGridValue(FG, rowIdx, 3, AccName)
            MDSearchAcccode = GetGridValue(FG, rowIdx, 2)
            txtSumAmountCr.Text = CDbl(txtSumAmountCr.Text) - CDbl(GetGridValue(FG, rowIdx, 6))
            SetGridValue(FG, rowIdx, 6, Format(CDbl(CDbl(txtAmount.Text) - CDbl(txtSumAmountCr.Text)), "#,##0.00"))
            SetGridValue(FG, rowIdx, 5, "0.00")
            SetGridValue(FG, rowIdx, 7, 0)
            SetGridValue(FG, rowIdx, 10, Cmb.Text)
            SetGridValue(FG, rowIdx, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
            SetGridValue(FG, rowIdx, 13, Format(CDbl(CDbl(txtRate.Text) * CDbl(GetGridValue(FG, rowIdx, 6))), "#,##0.00"))
            SetGridValue(FG, rowIdx, 12, "0.00")
            SumAmountDr()
            If ChDe.Checked = True Then
                If MuLng = "L" Then
                    FG.CurrentCell = FG.Rows(rowIdx).Cells(3)
                End If
                If MuLng = "E" Then
                    FG.CurrentCell = FG.Rows(rowIdx).Cells(4)
                End If
                If GetGridValue(FG, FG.Rows.Count - 1, 3) <> "" Then
                    FG.Rows.Add()
                End If
                Exit Sub
            End If

            FG.CurrentCell = FG.Rows(rowIdx).Cells(6)

            If GetGridValue(FG, FG.Rows.Count - 1, 3) <> "" Then
                FG.Rows.Add()
                'Timer1.Enabled = True
            End If
            'Call loadColor()
        End If
        SumAmountDr()
        FG.Focus()

    End Sub
    Public Sub SumAmountDr()
        Dim i As Integer
        Dim AmountDr, AmountCr, TotalAmountDr, TotalAmountCr As Double
        AmountDr = 0
        AmountCr = 0
        
        ' Remove empty rows first
        For i = FG.Rows.Count - 1 To 0 Step -1
            If GetGridValue(FG, i, 3) = "" Then
                FG.Rows.RemoveAt(i)
            End If
        Next
        
        ' Recalculate amounts
        For i = 0 To FG.Rows.Count - 1
            If GetGridValue(FG, i, 1) <> "" Or GetGridValue(FG, i, 2) <> "" Then
                AmountDr = AmountDr + CDbl(GetGridValue(FG, i, 5))
                AmountCr = AmountCr + CDbl(GetGridValue(FG, i, 6))
                TotalAmountDr = TotalAmountDr + CDbl(GetGridValue(FG, i, 12))
                TotalAmountCr = TotalAmountCr + CDbl(GetGridValue(FG, i, 13))
            End If
        Next
        
        txtSumAmountDr.Text = Format(AmountDr, "#,##0.00")
        txtSumAmountCr.Text = Format(AmountCr, "#,##0.00")
        txtSumTotalAmountDr.Text = Format(TotalAmountDr, "#,##0.00")
        txtSumTotalAmountCr.Text = Format(TotalAmountCr, "#,##0.00")


        Dr.Text = CDbl(txtSumAmountDr.Text) - CDbl(txtSumAmountCr.Text)
        Cr.Text = CDbl(txtSumAmountCr.Text) - CDbl(txtSumAmountDr.Text)
        DDR.Text = CDbl(txtSumTotalAmountDr.Text) - CDbl(txtSumTotalAmountCr.Text)
        CCR.Text = CDbl(txtSumTotalAmountCr.Text) - CDbl(txtSumTotalAmountDr.Text)
        Dr.Text = Format(CDbl(Dr.Text), "#,##0.00")
        Cr.Text = Format(CDbl(Cr.Text), "#,##0.00")
        DDR.Text = Format(CDbl(DDR.Text), "#,##0.00")
        CCR.Text = Format(CDbl(CCR.Text), "#,##0.00")

        If CDbl(Dr.Text) < 0 Then Dr.Text = "0.00"
        If CDbl(Cr.Text) < 0 Then Cr.Text = "0.00"
        If CDbl(DDR.Text) < 0 Then DDR.Text = "0.00"
        If CDbl(CCR.Text) < 0 Then CCR.Text = "0.00"

    End Sub
    Private Sub BtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch.Click
        fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
        FG.Focus()
        Call loadColor()
    End Sub

    Public Sub LoadText()
        AccName = ""
        Dim dtAcc As DataTable = DbHelper.GetDataTable("SELECT * FROM Acc_Code WHERE AC_CODE = N'" & AccId & "'")
        If dtAcc.Rows.Count > 0 Then
            AccId = Trim(dtAcc.Rows(0)("AC_CODE").ToString())
            AccName = Trim(dtAcc.Rows(0)("Name_L").ToString())
            AccNamee = Trim(dtAcc.Rows(0)("Name_E").ToString())
        End If
    End Sub
    Public Sub LoadDesc()
        Dim dtAcc As DataTable = DbHelper.GetDataTable("SELECT * FROM Acc_Code WHERE AC_CODE = N'" & AccId & "'")
        If dtAcc.Rows.Count > 0 Then
            AccId = Trim(dtAcc.Rows(0)("AC_CODE").ToString())
            AccName = Trim(dtAcc.Rows(0)("Name_L").ToString())
            AccNamee = Trim(dtAcc.Rows(0)("Name_E").ToString())
        End If
    End Sub
    Private Sub FG_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles FG.SelectionChanged

        BtnSearch.Visible = False
        BtnMove.Visible = False
    End Sub
    Dim ACCode As String
    Private Sub FG_KeyUp(ByVal sender As Object, ByVal e As KeyEventArgs) Handles FG.KeyUp
If e.KeyCode = Keys.Enter Then
            rowIdx = If(FG.CurrentRow IsNot Nothing, FG.CurrentCell.RowIndex, 0)
            If Cmb.Text = "USD" Then
                SetGridValue(FG, rowIdx, 14, Format(CDbl(GetGridValue(FG, rowIdx, 5)), "#,##0.00"))
                SetGridValue(FG, rowIdx, 15, Format(CDbl(GetGridValue(FG, rowIdx, 6)), "#,##0.00"))
            Else
                SetGridValue(FG, rowIdx, 14, Format(CDbl(GetGridValue(FG, rowIdx, 12)) / CDbl(txtRateUSD.Text), "#,##0.00"))
                SetGridValue(FG, rowIdx, 15, Format(CDbl(GetGridValue(FG, rowIdx, 13)) / CDbl(txtRateUSD.Text), "#,##0.00"))
            End If
        End If
      
        If Microsoft.VisualBasic.Left(ACCode, 3) = "110" Or Microsoft.VisualBasic.Left(ACCode, 3) = "112" Or Microsoft.VisualBasic.Left(ACCode, 3) = "113" Then

         

            Call Load_Gen_Jn()
            Remain = CDbl(CDbl(Open_jn) + CDbl(SumDr)) - CDbl(SumCr)

            If Remain >= 0 Then
                Remain = Format(CDbl(Remain), "##,##0.00")
            Else
                Remain = "(" & Format(CDbl(Remain * (-1)), "##,##0.00") & ")"
            End If
rowIdx = If(FG.CurrentRow IsNot Nothing, FG.CurrentCell.RowIndex, 0)
            If CDbl(GetGridValue(FG, rowIdx, 14)) > CDbl(Remain) Then
            Else
                SetGridValue(FG, rowIdx, 14, Format(CDbl(Remain), "#,##0.00"))
            End If
            If CDbl(GetGridValue(FG, rowIdx, 15)) > CDbl(Remain) Then
            Else
                SetGridValue(FG, rowIdx, 15, Format(CDbl(Remain), "#,##0.00"))
            End If

        End If

    End Sub
    Dim Remain As Double
    Dim SumDr As Double
    Dim SumCr As Double
    Dim Op As Double
    Dim Open_jn As Double
    Private Sub Load_Gen_Jn()
        Dim OfUsr1 As String = Microsoft.VisualBasic.Left(Off_Usr.Text, 5)
        Dim OfUsr2 As String = Mid(Off_Usr.Text, 4, 2)
        Dim OfUsr3 As String = Microsoft.VisualBasic.Left(Off_Usr.Text, 2)
        If OfUsr1 = "00-00" Then
            MULook2 = ""
        Else
            If OfUsr2 = "00" Then
                MULook2 = "  And  Left(company,2)= '" & OfUsr3 & "' "
            Else
                MULook2 = "  And company= '" & OfUsr1 & "' "
            End If
        End If
    
        
Dim s As String = "SELECT sum(Amt_dr) as Amt_dr   , sum(Amt_cr) as Amt_cr    FROM gen_jn WHERE ac_code = '" & ACCode & "' and gen_jn.date_work BETWEEN '" & Format(dtActi.Value, "yyyy-MM-dd") & "' AND '" & Format(dtActi.Value, "yyyy-MM-dd") & "'   " & MULook2 & "  "
        Dim dt1 As DataTable = DbHelper.GetDataTable(s)
        If dt1.Rows.Count <> 0 Then
            SumDr = Format(CDbl(Trim(dt1.Rows(0)("Amt_dr").ToString())), "#,##0.00")
            SumCr = Format(CDbl(Trim(dt1.Rows(0)("Amt_cr").ToString())), "#,##0.00")
        End If

        Dim dt2 As DataTable = DbHelper.GetDataTable("select  amount_dr , amount_cr from Open_jn where ac_code='" & ACCode & "'   and  year(Date_work)= '" & Format(CDate(dtActi.Value), "yyyy") & "'  " & MULook2 & "   ")
        Op = 0
        If dt2.Rows.Count <> 0 Then
            Op = CDbl(Trim(dt2.Rows(0)("amount_dr").ToString())) - CDbl(Trim(dt2.Rows(0)("amount_cr").ToString()))
        End If
        Dim dss As Date
        dss = DateAdd(DateInterval.Day, -1, dtActi.Value)
        Dim dt3 As DataTable = DbHelper.GetDataTable("select SUM(amount_dr) AS amount_dr ,SUM(amount_cr) AS amount_cr from Gen_jn where ac_code=N'" & ACCode & "'  And gen_jn.date_work   BETWEEN '" & "1-1-" & Format(dtActi.Value, "yyyy") & "' AND '" & Format(dss, "yyyy-MM-dd") & "' " & MULook2 & " group by ac_code ")
        If dt3.Rows.Count <> 0 Then
            Op = Op + CDbl(CDbl(Trim(dt3.Rows(0)("amount_dr").ToString())) - CDbl(Trim(dt3.Rows(0)("amount_Cr").ToString())))
        End If

        If Op >= 0 Then
            Open_jn = Format(CDbl(Op), "##,##0.00")
        Else
            Open_jn = "(" & Format(CDbl(Op * (-1)), "##,##0.00") & ")"
        End If

    End Sub
    Private Sub FG_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles FG.MouseDown
        If txtInvoice.BackColor = Color.Red Then
            txtInvoice.Focus()
            Exit Sub
        End If
        MouseDownEvent()
    End Sub

    Public Sub MouseDownEvent()
        Select Case MouseButtons
            Case Windows.Forms.MouseButtons.Right
                FG.BeginEdit(True)
            Case Windows.Forms.MouseButtons.Left
                If FG.Rows.Count >= 3 Then
                    BtnMove.Visible = True

                    BtnMove.Top = CInt((FG.GetCellDisplayRectangle(FG.CurrentCell.ColumnIndex, FG.CurrentCell.RowIndex, False).Top / 15) + FG.Top)

                    If FG.CurrentRow IsNot Nothing AndAlso GetGridValue(FG, FG.CurrentRow.Index, FG.CurrentCell.ColumnIndex) <> "" Then
                        If FG.CurrentCell = FG.Rows(1).Cells(7) Then
                            If ChCat_ID.Checked = True Then
                                Panel1.Visible = True
                                Bee.Row = 1
                                Bee.Col = 3
                                Bee.Focus()
                            End If

                        Else
                            Panel1.Visible = False
                        End If
                    End If

                    Panel1.Top = CInt((FG.GetCellDisplayRectangle(FG.CurrentCell.ColumnIndex, FG.CurrentCell.RowIndex, False).Top / 15) + FG.Top)
                    Panel1.Left = CInt(FG.Left + (FG.GetCellDisplayRectangle(FG.CurrentCell.ColumnIndex, FG.CurrentCell.RowIndex, False).Left / 15) + (FG.GetCellDisplayRectangle(FG.CurrentCell.ColumnIndex, FG.CurrentCell.RowIndex, False).Width / 250))
If FG.CurrentRow IsNot Nothing AndAlso FG.CurrentRow.Index = FG.Rows.Count - 1 Then
                        BtnMove.Visible = False
                    End If
                Else
                    BtnMove.Visible = False
                End If
                If FG.CurrentCell = FG.Rows(1).Cells(1) Then
                    If FG.CurrentRow IsNot Nothing AndAlso GetGridValue(FG, FG.CurrentRow.Index, 2) <> "" Then
                        BtnSearch.Visible = False
                    Else
                        BtnSearch.Visible = True
                    End If
                End If
                If FG.CurrentCell = FG.Rows(1).Cells(2) Then
                If FG.CurrentRow IsNot Nothing AndAlso GetGridValue(FG, FG.CurrentRow.Index, 1) <> "" Then
                        BtnSearch.Visible = False
                    Else
                        BtnSearch.Visible = True
                    End If
                End If
                If FG.CurrentCell = FG.Rows(1).Cells(1) Then
                    BtnSearch.Left = CInt(FG.Left + (FG.GetCellDisplayRectangle(FG.CurrentCell.ColumnIndex, FG.CurrentCell.RowIndex, False).Left / 15) + (FG.GetCellDisplayRectangle(FG.CurrentCell.ColumnIndex, FG.CurrentCell.RowIndex, False).Width / 21.8))
                    BtnSearch.Top = CInt((FG.GetCellDisplayRectangle(FG.CurrentCell.ColumnIndex, FG.CurrentCell.RowIndex, False).Top / 15) + FG.Top)


                End If
                If FG.CurrentCell = FG.Rows(1).Cells(2) Then
                    BtnSearch.Size = New System.Drawing.Point(34, 26)
                    BtnSearch.Left = CInt(FG.Left + (FG.GetCellDisplayRectangle(FG.CurrentCell.ColumnIndex, FG.CurrentCell.RowIndex, False).Left / 15) + (FG.GetCellDisplayRectangle(FG.CurrentCell.ColumnIndex, FG.CurrentCell.RowIndex, False).Width / 22.2))
                    BtnSearch.Top = CInt((FG.GetCellDisplayRectangle(FG.CurrentCell.ColumnIndex, FG.CurrentCell.RowIndex, False).Top / 15) + FG.Top)
                End If

                If FG.CurrentRow IsNot Nothing AndAlso FG.CurrentRow.Index = FG.Rows.Count - 1 Then
                    BtnMove.Visible = False
                End If

        End Select
    End Sub

    Private Sub btnmove_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnMove.Click
        If FG.CurrentRow IsNot Nothing Then FG.Rows.RemoveAt(FG.CurrentRow.Index)
        SumAmountDr()
        Panel1.Visible = False
        BtnMove.Visible = False
        BtnSearch.Visible = False

    End Sub


    'Sum
    Private Sub FG_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelectionChanged
        If txtInvoice.BackColor = Color.Red Then
            txtInvoice.Focus()
            Exit Sub
        End If

        'If txxdes Then
        FG.BackColor = Color.White
        ' FG.BackColorAlternate not available in DataGridView
        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 1 Then
            R = If(FG.CurrentRow IsNot Nothing, FG.CurrentRow.Index, 0)
            L = FG.CurrentCell.ColumnIndex
            If FG.CurrentRow IsNot Nothing AndAlso GetGridValue(FG, FG.CurrentRow.Index, 2) <> "" Then
                FG.BackColorSel = Color.White
                FG.ReadOnly = True
            Else
                FG.BackColorSel = Color.SkyBlue
                FG.ReadOnly = False
                AccId = If(FG.CurrentRow IsNot Nothing, GetGridValue(FG, FG.CurrentRow.Index, FG.CurrentCell.ColumnIndex), "")
            End If
        End If

        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 2 Then
            R = If(FG.CurrentRow IsNot Nothing, FG.CurrentRow.Index, 0)
            L = FG.CurrentCell.ColumnIndex
            If FG.CurrentRow IsNot Nothing AndAlso GetGridValue(FG, FG.CurrentRow.Index, 1) <> "" Then
                FG.BackColorSel = Color.White
                FG.ReadOnly = True
                'FG.BackColorSel = Color.White
            Else
                FG.BackColorSel = Color.SkyBlue
                FG.ReadOnly = False
                AccId = If(FG.CurrentRow IsNot Nothing, GetGridValue(FG, FG.CurrentRow.Index, FG.CurrentCell.ColumnIndex), "")
            End If
        End If



        If FG.CurrentCell = FG.Rows(1).Cells(3) Then
            'FG.BackColorSel = Color.White
            FG.ReadOnly = False
            FG.BackColorSel = Color.SkyBlue
        End If

        If FG.CurrentCell = FG.Rows(1).Cells(4) Then
            'FG.BackColorSel = Color.White
            FG.ReadOnly = False
            FG.BackColorSel = Color.SkyBlue
        End If


        If FG.CurrentCell = FG.Rows(1).Cells(5) Then
            If FG.CurrentRow IsNot Nothing AndAlso GetGridValue(FG, FG.CurrentRow.Index, 2) <> "" Then
                FG.ReadOnly = True
                FG.BackColorSel = Color.White
            Else
                FG.ReadOnly = False
                FG.BackColorSel = Color.SkyBlue
            End If
        End If

        If FG.CurrentCell = FG.Rows(1).Cells(6) Then
            If FG.CurrentRow IsNot Nothing AndAlso GetGridValue(FG, FG.CurrentRow.Index, 1) <> "" Then
                FG.BackColorSel = Color.White
                FG.ReadOnly = True

            Else
                FG.BackColorSel = Color.SkyBlue
                FG.ReadOnly = False
            End If
        End If

        If FG.CurrentCell = FG.Rows(1).Cells(7) Then
            FG.BackColorSel = Color.SkyBlue
            FG.ReadOnly = False
        End If


        If ChCat_ID.Checked = False Then
            If FG.CurrentCell IsNot Nothing AndAlso (FG.CurrentCell.ColumnIndex = 3 OrElse FG.CurrentCell.ColumnIndex = 4 OrElse FG.CurrentCell.ColumnIndex = 7 OrElse FG.CurrentCell.ColumnIndex = 8 OrElse FG.CurrentCell.ColumnIndex = 9 OrElse FG.CurrentCell.ColumnIndex = 10 OrElse FG.CurrentCell.ColumnIndex = 11 OrElse FG.CurrentCell.ColumnIndex = 12 OrElse FG.CurrentCell.ColumnIndex = 13) Then
                FG.BackColorSel = Color.White
                FG.ReadOnly = True
            End If
        Else
            If FG.CurrentCell IsNot Nothing AndAlso (FG.CurrentCell.ColumnIndex = 3 OrElse FG.CurrentCell.ColumnIndex = 4 OrElse FG.CurrentCell.ColumnIndex = 8 OrElse FG.CurrentCell.ColumnIndex = 9 OrElse FG.CurrentCell.ColumnIndex = 10 OrElse FG.CurrentCell.ColumnIndex = 11 OrElse FG.CurrentCell.ColumnIndex = 12 OrElse FG.CurrentCell.ColumnIndex = 13) Then
                FG.BackColorSel = Color.White
                FG.ReadOnly = True
            End If
        End If
        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex > 2 Then
            If FG.CurrentRow IsNot Nothing AndAlso GetGridValue(FG, FG.CurrentRow.Index, 1) = "" Then
                If FG.CurrentRow IsNot Nothing AndAlso GetGridValue(FG, FG.CurrentRow.Index, 2) = "" Then
            FG.ReadOnly = True
        ' FG.BackColorSel not available in DataGridView
                End If
            End If
        End If
        'If CheckBox3.Checked Then
        '    CheckBox3.
        'End If
        If FG.CurrentCell IsNot Nothing AndAlso (FG.CurrentCell.ColumnIndex = 3 OrElse FG.CurrentCell.ColumnIndex = 4 OrElse FG.CurrentCell.ColumnIndex = 5 OrElse FG.CurrentCell.ColumnIndex = 6 OrElse FG.CurrentCell.ColumnIndex = 7 OrElse FG.CurrentCell.ColumnIndex = 8 OrElse FG.CurrentCell.ColumnIndex = 9 OrElse FG.CurrentCell.ColumnIndex = 10 OrElse FG.CurrentCell.ColumnIndex = 11 OrElse FG.CurrentCell.ColumnIndex = 12 OrElse FG.CurrentCell.ColumnIndex = 13) Then
            BtnSearch.Visible = False
        End If

        If FG.CurrentCell IsNot Nothing AndAlso GetGridValue(FG, FG.CurrentCell.RowIndex, 5) = "" Then
            Panel1.Visible = False
        End If
        If GetGridValue(FG, FG.CurrentCell.RowIndex, 7) = "0" Then
            SetGridValue(FG, rowIdx, 8, "ບໍ່ເລືອກ")
            SetGridValue(FG, rowIdx, 9, "No Selete")
        ElseIf GetGridValue(FG, FG.CurrentCell.RowIndex, 7) = "1" Then
            SetGridValue(FG, rowIdx, 8, "ຮັບໃຊ້ການພະລິດ")
            SetGridValue(FG, rowIdx, 9, "Use build")

        ElseIf GetGridValue(FG, FG.CurrentCell.RowIndex, 7) = "2" Then
            SetGridValue(FG, rowIdx, 8, "ຮັບໃຊ້ການຈຳໜ່າຍ")
            SetGridValue(FG, rowIdx, 9, "Use Sell")
        ElseIf GetGridValue(FG, FG.CurrentCell.RowIndex, 7) = "3" Then
            SetGridValue(FG, rowIdx, 8, "ຮັບໃຊ້ບໍລິຫານ")
            SetGridValue(FG, rowIdx, 9, "Use manage ")

        ElseIf GetGridValue(FG, FG.CurrentCell.RowIndex,7) = "4" Then
            SetGridValue(FG, rowIdx, 8, "ຕົ້ນທຶນຂາຍ/ຕົ້ນທຶນບໍລິຫານ")
            SetGridValue(FG, rowIdx, 9, "Sell capital/manage capital ")
        End If

    End Sub

    Private Sub CmbBook_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles CmbBook.KeyPress
        If e.KeyChar = Chr(13) Then
            txtAmount.Focus()
        End If
    End Sub

    Private Sub CmbBook_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbBook.SelectedIndexChanged
        txtInvoice.Text = CmbBook.Text & MdCertifyId
Dim dtBook As DataTable = DbHelper.GetDataTable("SELECT * FROM books WHERE bookid = N'" & CmbBook.Text & "'")
        If dtBook.Rows.Count > 0 Then
            txtBookName.Text = Trim(dtBook.Rows(0)("bookname").ToString())
        End If
        If txtInvoice.Enabled = True Then
            AutoNumber()
        End If
    End Sub

    Private Sub txtAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAmount.KeyPress
        If e.KeyChar = Chr(13) Then
            txtTotal_Amt_LAK.Text = CDbl(txtRate.Text) * CDbl(txtAmount.Text)
            Call FormatText()
If GetGridValue(FG, 0, 1) = "" And GetGridValue(FG, 0, 2) = "" Then
                FG.CurrentCell = FG.Rows(0).Cells(1)
            End If
            If GetGridValue(FG, 0, 1) <> "" Then
                FG.CurrentCell = FG.Rows(0).Cells(1)
            End If
        End If
            '        FG.CurrentCell = FG.Rows(1).Cells(7)
            '        FG.CellForeColor = Color.Gray
            '        If FG.CurrentRow IsNot Nothing Then FG.CurrentCell = FG.Rows.Count(FG.CurrentRow.Index).Cells( 8
            '        FG.CellForeColor = Color.Gray
            '        If FG.CurrentRow IsNot Nothing Then FG.CurrentCell = FG.Rows.Count(FG.CurrentRow.Index).Cells( 9
            '        FG.CellForeColor = Color.Gray
            '    Else
            '        FG.CurrentCell = FG.Rows(1).Cells(7)
            '        FG.CellForeColor = Color.Black
            '        If FG.CurrentRow IsNot Nothing Then FG.CurrentCell = FG.Rows.Count(FG.CurrentRow.Index).Cells( 8
            '        FG.CellForeColor = Color.Black
            '        If FG.CurrentRow IsNot Nothing Then FG.CurrentCell = FG.Rows.Count(FG.CurrentRow.Index).Cells( 9
            '        FG.CellForeColor = Color.Black
            '    End If
            '    If ChCat_ID.Checked = False Then

            '        FG.CurrentCell = FG.Rows(1).Cells(7)

            '        FG.CellBackColor = Color.White
            '        FG.CellForeColor = Color.Gray
            '        If FG.CurrentRow IsNot Nothing Then FG.CurrentCell = FG.Rows.Count(FG.CurrentRow.Index).Cells( 8
            '        FG.CellForeColor = Color.Gray
            '        If FG.CurrentRow IsNot Nothing Then FG.CurrentCell = FG.Rows.Count(FG.CurrentRow.Index).Cells( 9
            '        FG.CellForeColor = Color.Gray

            '    End If



        ' FG.CurrentRow.Index is read-only
        If FG.CurrentRow IsNot Nothing Then FG.CurrentCell = FG.Rows(FgR).Cells(FgC)
        ' FG.Redraw not available in DataGridView
    End Sub

    Private Sub Cmb_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cmb.Click
        'LoadCurr()
    End Sub

    Private Sub Cmb_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmb.SelectedIndexChanged

Dim dtCurr As DataTable = DbHelper.GetDataTable("Select * From Curr_For_Rate Where   Curr =N'" & Trim(Cmb.Text) & "'")
        If dtCurr.Rows.Count > 0 Then
            txtcurr_name2.Text = Trim(dtCurr.Rows(0)("Curr_name").ToString())
        End If

        MDRate_DT = " and rate_dt<='" & Format(dtActi.Value, "yyyy-MM-dd") & "'  "
        SS_Curr = " and AP_Rate_history.Curr =N'" & Cmb.Text & "' "
        Call RateSetting()
        txtRate.Text = Format(MD_Rate, "#,##0.00")
        txtRateUSD.Text = Format(MDUSD_LAK, "#,##0.00")
        curr_Last.Text = MDRate_Curr



        txtAmt_letter.Text = Letter_amt(txtAmount)
        txtAmt_letter_E.Text = LetterEng_amt(txtAmount)
        Call FormatText()

        'For i = 0 To FG.Rows.Count - 1
        '    'If If FG.CurrentRow IsNot Nothing Then FG.CurrentRow.Index > 1 Then
        '    SetGridValue(FG,i, 10, Cmb.Text)
        '    SetGridValue(FG,i, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
        '    SetGridValue(FG,i, 12, Format(CDbl(CDbl(txtRate.Text) * GetGridValue(FG,i, 5)), "#,##0.00"))
        '    SetGridValue(FG,i, 13, Format(CDbl(CDbl(txtRate.Text) * GetGridValue(FG,i, 6)), "#,##0.00"))
        '    'End If

        'Next
    End Sub
    Public Sub LoadSetRate()
Dim dtRateSet As DataTable = DbHelper.GetDataTable("select * from Ap_RateSeting where Curr='" & Cmb.Text & "'")
        If dtRateSet.Rows.Count > 0 Then
            txtRate.Text = Trim(dtRateSet.Rows(0)("Rate").ToString())
            curr_Last.Text = Trim(dtRateSet.Rows(0)("curr_Last").ToString())
        End If
        'loadRate()
        txtAmt_letter.Text = Letter_amt(txtAmount)
        txtAmt_letter_E.Text = LetterEng_amt(txtAmount)
        Call FormatText()
    End Sub



    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        IVN = txtInvoice.Text
        SumAmountDr()
        If MdCertifyAuto = 1 Then
            If txtInvoice.Enabled = True Then
                AutoNumber()
            End If
        End If

Dim i As Integer
        For i = 0 To FG.Rows.Count - 1
            If GetGridValue(FG,i, 1) & GetGridValue(FG,i, 2) <> "" Then
                'MsgBox(Len(GetGridValue(FG,i, 1) & GetGridValue(FG,i, 2)))
                If Len(GetGridValue(FG,i, 1) & GetGridValue(FG,i, 2)) <> 7 Then
                    'LngId = "6005" : MsgRpt()
                    'Exit Sub
                End If
            End If
        Next i

Dim k As Integer
        For k = 0 To FG.Rows.Count - 1
            If GetGridValue(FG,k, 1) <> "" Then
                If GetGridValue(FG,k, 5) = 0 Then
                    MsgBox("ກະລຸນນາໃສ່ມູນຄ່າກ່ອນ")
                    Exit Sub
                End If
            End If
            If GetGridValue(FG,k, 2) <> "" Then
                If GetGridValue(FG,k, 6) = 0 Then
                    MsgBox("ກະລຸນນາໃສ່ມູນຄ່າກ່ອນ")
                    Exit Sub
                End If
            End If
        Next k

        If txtInvoice.Text = "" Then
            MsgBox("ກະລຸນນາໃສ່ເລກໃບຢັງຢືນກອ່ນ!", MsgBoxStyle.OkOnly)
            txtInvoice.BackColor = Color.Red
            txtInvoice.Focus()
            Exit Sub
        End If
        If TxtReferno.Text = "" Then
            MsgBox("ກະລຸນນາໃສ່ເລກອ້າງອີງກ່ອນ!", MsgBoxStyle.OkOnly)
            'TxtReferno.BackColor = Color.Red
            TxtReferno.Focus()
            Exit Sub
        End If



        'Call LoadSqlData("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book ='" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "' And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And  Month(date_work)=" & Format(CDate(dtActi.Value), "MM") & " " & _
        '                 "  and LEFT(company,2)='" & Off_Id & "' Order by  Right(certify,3) DESC ", RSC)
        'Call LoadSqlData("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book =N'" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "' And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And  Month(date_work)=" & Format(CDate(dtActi.Value), "MM") & "  And  day(date_work)=" & Format(CDate(dtActi.Value), "dd") & " " & _
        '             "  and LEFT(company,2)=N'" & Off_Id & "' Order by  Right(certify,3) DESC ", RSC)
        Dim dtCheck3 As DataTable = DbHelper.GetDataTable("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book =N'" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "'  And  ReferNO = N'" & TxtReferno.Text & "'And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And  Month(date_work)=" & Format(CDate(dtActi.Value), "MM") & "  And  day(date_work)=" & Format(CDate(dtActi.Value), "dd") & " " & _
             "  and LEFT(company,2)=N'" & Off_Id & "' Order by  Right(certify,3) DESC ")

        If dtCheck3.Rows.Count > 0 Then
            If txtInvoice.Enabled = True Then
                MsgBox("ເລກລະຫັດ : " & Trim(txtInvoice.Text) & " ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
                txtInvoice.BackColor = Color.Red
                txtInvoice.Focus()
                Exit Sub
            End If

        End If
        'Dim J As Integer
        'For J = 1 To FG.Rows.Count
        '    '===============
        '    SetGridValue(FG,J, 10, Cmb.Text)
        '    SetGridValue(FG,J, 11, Format(CDbl(txtRate.Text), "#,##0.00"))


        '    If Cmb.Text = "USD" Then
        '        SetGridValue(FG,J, 14, Format(CDbl(GetGridValue(FG,J, 5)), "#,##0.00"))
        '        SetGridValue(FG,J, 15, Format(CDbl(GetGridValue(FG,J, 6)), "#,##0.00"))
        '    Else
        '        SetGridValue(FG,J, 14, Format(CDbl(GetGridValue(FG,J, 12)) / CDbl(txtRateUSD.Text), "#,##0.00"))
        '        SetGridValue(FG,J, 15, Format(CDbl(GetGridValue(FG,J, 13)) / CDbl(txtRateUSD.Text), "#,##0.00"))
        '        'SetGridValue(FG, rowIdx, 14, Format(CDbl(If FG.CurrentRow IsNot Nothing Then GetGridValue(FG, 12)) / CDbl(txtRateUSD.Text), "#,##0.00"))
        '        'SetGridValue(FG, rowIdx, 15, Format(CDbl(GetGridValue(FG, rowIdx, 13)) / CDbl(txtRateUSD.Text), "#,##0.00"))
        '    End If

        'Next J

        If GetGridValue(FG,1, 6) = "" Then MsgBox("ກະລຸນນາລົງບັນຊີເງິນກອ່ນ!", MsgBoxStyle.OkOnly) : Exit Sub
        If CDbl(txtSumAmountDr.Text) = 0 Then MsgBox("ການລົງບັນຊີເງິນບໍ່ຖຶກຕ້ອງ ບໍ່ສາມາດບັນທຶກໄດ້!", MsgBoxStyle.OkOnly) : Exit Sub
        If CDbl(txtSumAmountCr.Text) = 0 Then MsgBox("ການລົງບັນຊີເງິນບໍ່ຖຶກຕ້ອງ ບໍ່ສາມາດບັນທຶກໄດ້!", MsgBoxStyle.OkOnly) : Exit Sub
        If CDbl(CCR.Text) <> 0 Then MsgBox("ບັນຊີນີ້ບໍ່ທັນບໍ່ທັນດູນດ່ຽງ ບໍ່ສາມາດບັນທຶກໄດ້!", MsgBoxStyle.OkOnly) : Exit Sub
        If CDbl(DDR.Text) <> 0 Then MsgBox("ບັນຊີນີ້ບໍ່ທັນບໍ່ທັນດູນດ່ຽງ ບໍ່ສາມາດບັນທຶກໄດ້!", MsgBoxStyle.OkOnly) : Exit Sub

        'If MessageBox.Show("ບັນຊີນີ້ດູນດ່ຽງແລ້ວ ທ່ານຕອ້ງການບັນທຶຫລືບໍ່!", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
        If txtInvoice.Enabled = True Then
            AutoNumber()
            MdCertifyId2 = txtInvoice.Text
            Sdate = dtActi.Text

            If txtInvoice.Text = "" Then
                MsgBox("ກະລຸນນາໃສ່ເລກໃບຢັງຢືນກອ່ນ!", MsgBoxStyle.OkOnly)
                txtInvoice.BackColor = Color.Red
                txtInvoice.Focus()
                Exit Sub
            End If

            'Call LoadSqlData("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book ='" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "' And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And  month(date_work)=" & Format(CDate(dtActi.Value), "MM") & "  Order by  Right(certify,3) DESC ", RSC)
            Dim dtCheck4 As DataTable = DbHelper.GetDataTable("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book =N'" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "'And  ReferNO = N'" & TxtReferno.Text & "' And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And  month(date_work)=" & Format(CDate(dtActi.Value), "MM") & " And  day(date_work)=" & Format(CDate(dtActi.Value), "dd") & "  Order by  Right(certify,3) DESC ")

            If dtCheck4.Rows.Count > 0 Then
                MsgBox("ເລກລະຫັດ : " & Trim(txtInvoice.Text) & " ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
                txtInvoice.BackColor = Color.Red
                txtInvoice.Focus()
                Exit Sub
            End If

            'MuSubOff = Mid(Off_Usr.Text, 1, 5)

            Savedata()
            'MuSubOff = MuSubOff2
        Else
            'CNN.Execute("DELETE FROM gen_jn WHERE book =N'" & GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 16) & "' And certify  = '" & txtInvoice.Text & "'   And  year(date_work)='" & Format(CDate(GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 1)), "yyyy") & "' ")
            DbHelper.ExecuteNonQuery("DELETE FROM gen_jn WHERE book =N'" & GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 16) & "' And certify  =N'" & txtInvoice.Text & "'  And  ReferNO = N'" & TxtReferno.Text & "'  And   date_work='" & Format(CDate(GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 1)), "yyyy-MM-dd") & "' ")

            MdCertifyId2 = txtInvoice.Text
            Sdate = dtActi.Text
            MuSubOff = Mid(Off_Usr.Text, 1, 5)
            Savedata()
            MuSubOff = MuSubOff2
        End If
        MsgBox("ບັນທຶກສໍາເລັດ")
        txtInvoice.Enabled = True
        CmbBook.Enabled = True
        Panel1.Visible = False
        BtnMove.Visible = False
        BtnSearch.Visible = False
        FG.Rows.Count = 1
        FG.Rows.Count = 2
        If CheckBox1.Checked Then
            Call LoadReport()
        End If
        If CheckBox2.Checked Then
            Close()
        End If
        'End If


    End Sub

    Private Sub BtnAddNew2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew2.Click
        FG.Rows.Count = 1
        FG.Rows.Count = 2
        AutoNumber()
        txtInvoice.Enabled = True
        CmbBook.Enabled = True
        Panel1.Visible = False
        BtnMove.Visible = False
        BtnSearch.Visible = False
    End Sub

    Private Sub Bee_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Bee.SelectionChanged

    End Sub

    Private Sub FG_CellBeginEdit(ByVal sender As Object, ByVal e As DataGridViewCellCancelEventArgs) Handles FG.CellBeginEdit
        If txtInvoice.BackColor = Color.Red Then
            txtInvoice.Focus()
            e.Cancel = True  ' Cancel the edit operation
            Exit Sub
        End If
        ' FG.CellBackColor/ForeColor not available in DataGridView
        BtnMove.Visible = False
        BtnSearch.Visible = False



    End Sub

Private Sub Cat()
        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.ColumnIndex = 7 Then
            rowIdx = FG.CurrentCell.RowIndex
            If GetGridValue(FG, rowIdx, 7) = "0" Then
                SetGridValue(FG, rowIdx, 8, "ບໍ່ເລືອກ")
                SetGridValue(FG, rowIdx, 9, "No Selete")
            ElseIf GetGridValue(FG, rowIdx, 7) = "1" Then
                SetGridValue(FG, rowIdx, 8, "ຮັບໃຊ້ການພະລິດ")
                SetGridValue(FG, rowIdx, 9, "Use build")

            ElseIf GetGridValue(FG, rowIdx, 7) = "2" Then
                SetGridValue(FG, rowIdx, 8, "ຮັບໃຊ້ການຈຳໜ່າຍ")
                SetGridValue(FG, rowIdx, 9, "Use Sell")
            ElseIf GetGridValue(FG, rowIdx, 7) = "3" Then
                SetGridValue(FG, rowIdx, 8, "ຮັບໃຊ້ບໍລິຫານ")
                SetGridValue(FG, rowIdx, 9, "Use manage ")

            ElseIf GetGridValue(FG, rowIdx, 7) = "4" Then
                SetGridValue(FG, rowIdx, 8, "ຕົ້ນທຶນຂາຍ/ຕົ້ນທຶນບໍລິຫານ")
                SetGridValue(FG, rowIdx, 9, "Sell capital/manage capital ")
            ElseIf Val(GetGridValue(FG, rowIdx, 7)) > 4 Then
                MessageBox.Show("ລະຫັດບໍ່ຖືກຕ້ອງ ລະຫັດມີພຽງເລກ 0 ຫາເລກ 4 ເທົ່ານັ້ນ")
                SetGridValue(FG, rowIdx, 8, "ບໍ່ລເລືອກ")
                SetGridValue(FG, rowIdx, 9, "No Selete")
                SetGridValue(FG, rowIdx, 7, "0")
                Exit Sub
            ElseIf Val(GetGridValue(FG, rowIdx, 7)) < 0 Then
                MessageBox.Show("ລະຫັດບໍ່ຖືກຕ້ອງ ລະຫັດມີພຽງເລກ 0 ຫາເລກ 4 ເທົ່ານັ້ນ")
                SetGridValue(FG, rowIdx, 8, "ບໍ່ລເລືອກ")
                SetGridValue(FG, rowIdx, 9, "No Selete")
                SetGridValue(FG, rowIdx, 7, "0")
                Exit Sub
            End If
        End If


    End Sub
    Private Sub LoadSQL()
        sql = ""
        'And  year(date_work)='" & Format(CDate(If FG.CurrentRow IsNot Nothing Then GetGridValue(FG, 2) Else ""), "yyyy") & "'
        'sql = " AND certify  = " & txtInvoice.Text & "  And  year(date_work)='" & Format(CDate(GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 1)), "yyyy") & "' "
        sql = " AND certify  =N" & txtInvoice.Text

    End Sub
    Public Sub LoadListFG()
        DbHelper.ExecuteNonQuery("UPDATE gen_jn set Cust_Supp=0 where Cust_Supp is null")
FG.Rows.Clear()
        
        Dim PK As String = "select * from gen_jn WHERE book =N'" & GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 16) & "' And certify =N'" & txtInvoice.Text & "' And ReferNO =N'" & GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 3) & "' And year(date_work)='" & Format(CDate(GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 1)), "yyyy") & "' order by cnt"
        Dim dt As DataTable = DbHelper.GetDataTable(PK)
        
        If dt.Rows.Count > 0 Then
For Each row As DataRow In dt.Rows
                Dim newRowIdx As Integer = FG.Rows.Add()
                If CDbl(Val(row("AG"))) = 1 Then
                    SetGridValue(FG, newRowIdx, 1, row("code_dr").ToString())
                    SetGridValue(FG, newRowIdx, 2, row("code_cr").ToString())
                    SetGridValue(FG, newRowIdx, 3, row("descrip").ToString())
                    SetGridValue(FG, newRowIdx, 4, row("descripe").ToString())
                    SetGridValue(FG, newRowIdx, 5, Format(CDbl(Val(row("amt_dr"))), "##,##0.00"))
                    SetGridValue(FG, newRowIdx, 6, Format(CDbl(Val(row("amt_cr"))), "##,##0.00"))
                    SetGridValue(FG, newRowIdx, 7, row("Cat_ID").ToString())
                Else
                    SetGridValue(FG, newRowIdx, 1, row("code_dr").ToString())
                    SetGridValue(FG, newRowIdx, 2, row("code_cr").ToString())
                    SetGridValue(FG, newRowIdx, 3, row("descrip").ToString())
                    SetGridValue(FG, newRowIdx, 4, row("descripe").ToString())
                    SetGridValue(FG, newRowIdx, 5, Format(CDbl(Val(row("amount_dr"))), "##,##0.00"))
                    SetGridValue(FG, newRowIdx, 6, Format(CDbl(Val(row("amount_cr"))), "##,##0.00"))
                    SetGridValue(FG, newRowIdx, 7, row("Cat_ID").ToString())
                    SetGridValue(FG, newRowIdx, 10, row("Curr").ToString())
                    SetGridValue(FG, newRowIdx, 11, Format(CDbl(Val(row("Rate"))), "##,##0.00"))
                    SetGridValue(FG, newRowIdx, 12, Format(CDbl(Val(row("amt_Dr"))), "##,##0.00"))
                    SetGridValue(FG, newRowIdx, 13, Format(CDbl(Val(row("amt_cr"))), "##,##0.00"))
                    SetGridValue(FG, newRowIdx, 14, Format(CDbl(Val(row("amt_USD_Dr"))), "##,##0.00"))
                    SetGridValue(FG, newRowIdx, 15, Format(CDbl(Val(row("amt_USD_Cr"))), "##,##0.00"))
                End If
            Next
        Else
For k = 0 To 15
                FG.Rows.Add()
            Next
        End If
        Dim PP As String = "select top 1 * from gen_jn WHERE book =N'" & GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 16) & "' And certify =N'" & txtInvoice.Text & "' And ReferNO =N'" & GetGridValue(FmJeneralJournal_List.FG, 3) & "' And date_work='" & Format(CDate(GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 1)), "yyyy-MM-dd") & "' order by cnt"
        Dim dtPP As DataTable = DbHelper.GetDataTable(PP)
        
        If dtPP.Rows.Count > 0 Then
            Dim row As DataRow = dtPP.Rows(0)
            Book = row("book").ToString().Trim()
            dtActi.Text = Format(CDate(row("date_work")), "dd/MM/yyyy")
            txtAmount.Text = Format(CDbl(Val(row("amount"))), "#,##0.00")
            txtDesc.Text = row("descrip").ToString().Trim()
            txtDescE.Text = row("descripe").ToString().Trim()
            Rate1 = Format(CDbl(Val(row("rate"))), "#,##0.00")
            Cmb.Text = row("curr").ToString().Trim()
            CmbBook.Text = Book
            txtInvoice.Text = MDInvoiceNo
            TxtCustID.Text = row("CustID").ToString().Trim()
            TxtSuppID.Text = row("SuppID").ToString().Trim()
            txtRateUSD.Text = Format(CDbl(Val(row("rate_USD"))), "#,##0.00")
            TxtReferno.Text = row("Referno").ToString().Trim()
            MCS = row("Cust_Supp").ToString().Trim()

            If MCS = "1" Then
                CheckBox4.Checked = True
                If TxtCustID.Text <> "" Then
                    Dim dtCust As DataTable = DbHelper.GetDataTable("select * from Customer WHERE 1=1 and Code=N'" & TxtCustID.Text & "'")
                    If dtCust.Rows.Count > 0 Then
                        CmbCust.Text = dtCust.Rows(0)("Name").ToString().Trim()
                    End If
                    RadioButton1.Checked = True
                End If
                If TxtSuppID.Text <> "" Then
                    Dim dtSupp As DataTable = DbHelper.GetDataTable("select * from Supplier WHERE 1=1 and Code=N'" & TxtSuppID.Text & "'")
                    If dtSupp.Rows.Count > 0 Then
                        CmbSupp.Text = dtSupp.Rows(0)("Name").ToString().Trim()
                    End If
                    RadioButton2.Checked = True
                End If
            Else
                CheckBox4.Checked = False
            End If
            CheckBox3.Checked = (CDbl(Val(row("AG"))) = 1)
        End If

        txtRate.Text = Rate1

For i = 0 To FG.Rows.Count - 1
            SetGridValue(FG, i, 12, Format(CDbl(Val(GetGridValue(FG, i, 11))) * CDbl(Val(GetGridValue(FG, i, 5))), "#,##0.00"))
            SetGridValue(FG, i, 13, Format(CDbl(Val(GetGridValue(FG, i, 11))) * CDbl(Val(GetGridValue(FG, i, 6))), "#,##0.00"))
            SetGridValue(FG, i, 14, Format(CDbl(Val(GetGridValue(FG, i, 12))) / CDbl(Val(txtRate.Text)), "#,##0.00"))
            SetGridValue(FG, i, 15, Format(CDbl(Val(GetGridValue(FG, i, 13))) / CDbl(Val(txtRate.Text)), "#,##0.00"))
            Cat(i)
        Next i

        For i = 0 To FG.Rows.Count - 1
            Dim acCode As String = GetGridValue(FG, i, 1) & GetGridValue(FG, i, 2)
            If acCode <> "" Then
                Dim dtAcc As DataTable = DbHelper.GetDataTable("select * from gen_jn WHERE Ac_Code=N'" & acCode & "' and book =N'" & GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 16) & "' And certify  =N'" & txtInvoice.Text & "'  And ReferNO  =N'" & GetGridValue(FmJeneralJournal_List.FG, 3) & "'    And  year(date_work)='" & Format(CDate(GetGridValue(FmJeneralJournal_List.FG, FmJeneralJournal_List.CurrentRow.Index, 1)), "yyyy") & "' order by cnt")
                If dtAcc.Rows.Count > 0 Then
                    SetGridValue(FG, i, 3, dtAcc.Rows(0)("ac_Name").ToString().Trim())
                    SetGridValue(FG, i, 4, dtAcc.Rows(0)("ac_Namee").ToString().Trim())
                End If
            End If
        Next i

        SumAmountDr()
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ShowList()
    End Sub
    Private Sub ShowList()
        FG.CurrentCell.ColumnIndexumns(7, True)
        FG.CurrentCell.ColumnIndexumns(8, True)
        If Button3.Text = "Show All" Then


            FG.CurrentCell.ColumnIndexumns(4, False)
            FG.CurrentCell.ColumnIndexumns(9, False)
            'FG.CurrentCell.ColumnIndexumns(10, False)
            'FG.CurrentCell.ColumnIndexumns(11, False)
            'FG.CurrentCell.ColumnIndexumns(12, False)
            'FG.CurrentCell.ColumnIndexumns(13, False)
            'FG.CurrentCell.ColumnIndexumns(14, True)
            'FG.CurrentCell.ColumnIndexumns(15, True)
            Button3.Text = "Show GLN"
            Exit Sub
        End If
        If Button3.Text = "Show GLN" Then
            If MuLng = "E" Then
                FG.CurrentCell.ColumnIndexumns(3, True)
                FG.CurrentCell.ColumnIndexumns(4, False)
            Else
                FG.CurrentCell.ColumnIndexumns(3, False)
                FG.CurrentCell.ColumnIndexumns(4, True)
            End If
            'FG.CurrentCell.ColumnIndexumns(4, True)
            FG.CurrentCell.ColumnIndexumns(9, True)
            'FG.CurrentCell.ColumnIndexumns(10, True)
            'FG.CurrentCell.ColumnIndexumns(11, True)
            'FG.CurrentCell.ColumnIndexumns(12, True)
            'FG.CurrentCell.ColumnIndexumns(13, True)
            'FG.CurrentCell.ColumnIndexumns(14, True)
            'FG.CurrentCell.ColumnIndexumns(15, True)
            Button3.Text = "Show All"
            Exit Sub
        End If
    End Sub


    Private Sub ChCat_ID_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChCat_ID.CheckedChanged

For J = 0 To FG.Rows.Count - 1
            If Trim(GetGridValue(FG, J, 3)) <> "" Then
                If ChCat_ID.Checked = True Then
                    FG.Rows(J).Cells(7).Style.BackColor = Color.LightCyan
                Else
                    FG.Rows(J).Cells(7).Style.BackColor = Color.White
                    FG.Rows(J).Cells(8).Style.ForeColor = Color.Gray
                    FG.Rows(J).Cells(9).Style.ForeColor = Color.Gray

                    SetGridValue(FG, J, 7, "0")
                    SetGridValue(FG, J, 8, "ບໍ່ເລືອກ")
                    SetGridValue(FG, J, 9, "No Selete")
                End If
            End If
        Next J


    End Sub

    Private Sub txtRate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRate.TextChanged

    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        FmCalcu.Show()
        FmCalcu.BringToFront()
    End Sub

    Private Sub txtDesc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDesc.GotFocus
        'SetLao(1)
    End Sub

    Private Sub txtDesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDesc.KeyPress
        If e.KeyChar = Chr(13) Then
            txtDescE.Focus()
        End If
    End Sub
    Private Sub txtDescE_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescE.GotFocus
        'SetLao(0)
    End Sub

    Private Sub txtDescE_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDescE.KeyPress
        If e.KeyChar = Chr(13) Then
            txtAmount.Focus()
        End If
    End Sub

    Private Sub Label23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        FmCalcu.Show()
        FmCalcu.BringToFront()
    End Sub

    Private Sub Label24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        FmCalcu.Show()
        FmCalcu.BringToFront()
    End Sub

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Call Close()
    End Sub
    Private Sub LoadReport()
        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & " 'As Crl_Lng  ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
        LngId = "7002" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDP ,"
        LngId = "7003" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
        LngId = "7004" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Descrip ,"
        LngId = "7005" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_No ,"
        LngId = "7006" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Ac_Code ,"
        LngId = "7007" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Certify ,"
        LngId = "7008" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Dr ,"
        LngId = "7009" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Cr ,"
        LngId = "7010" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Amt_LAK ,"
        LngId = "7011" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Sign ,"
        LngId = "7012" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_SignNme ,"
        LngId = "7026" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InDate ,"
        LngId = "7028" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_AmtCurr ,"
        LngId = "7029" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Curr ,"
        LngId = "7030" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Rate ,"
        LngId = "7031" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_User ,"
        LngId = "7032" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Amount ,"
        LngId = "7033" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Number ,"

        LngId = "7039" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Amount_Total	 ,"
        LngId = "7069" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_In_Word	 ," 
        'LngId = "7038" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As  Crl_Ac_Name	 ,"
        LngId = "5019" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As  Crl_Ac_Name	 ,"
        LngId = "7122" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As  Refno ,"
        LngId = "7123" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As  Amount ,"
        LngId = "7124" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As  Fore  ,"
        LngId = "7125" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As  Exchange ,"
        LngId = "7126" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As  Amount_LAK ,"



        If MuLng = "L" Then
            MuLngRpt = MuLngRpt & "N'" & Amount_In_Word & "' As Crl_Amt_In_Word	 ,"
        Else
            MuLngRpt = MuLngRpt & "N'" & Amount_In_Word & "' As Crl_Amt_In_Word	 ,"
        End If


        'SLF = MuLngRpt & " gen_jn.company ,gen_jn.Date_Work , gen_jn.certify, gen_jn.referno,  gen_jn.ac_code, gen_jn.descrip , gen_jn.descripe , gen_jn.amt_dr, gen_jn.amt_cr, gen_jn.ac_name  AS Name_L , gen_jn.ac_namee AS Name_E  "
        SLF = MuLngRpt & "   gen_jn.Rate_USD, gen_jn.company ,gen_jn.Date_Work , gen_jn.Curr, gen_jn.certify, gen_jn.referno, gen_jn.ac_code, gen_jn.descrip , gen_jn.descripe , gen_jn.amount_dr, gen_jn.amount_cr, gen_jn.amt_dr, gen_jn.amt_cr, gen_jn.ac_name  AS Name_L , gen_jn.ac_namee AS Name_E  "

        'SLF = MuLngRpt & " gen_jn.company ,gen_jn.Date_Work , gen_jn.certify, gen_jn.ac_code, gen_jn.descrip , gen_jn.descripe , gen_jn.amt_dr, gen_jn.amt_cr, Acc_Code.Name_L AS Name_L , Acc_Code.Name_E AS Name_E  "
        'SLF = RptSjOff & "    gen_jn.certify, gen_jn.ac_code, gen_jn.descrip, gen_jn.amt_dr, gen_jn.amt_cr, Acc_Code.Name_L AS Name_L , "
        Call LoadLoGO()
        Dim dt As DataTable
        Try
            dt = DbHelper.GetDataTable("SELECT   " & SLF & "  FROM         gen_jn INNER JOIN    Acc_Code ON gen_jn.ac_code = Acc_Code.Ac_Code WHERE gen_jn.certify = N'" & MdCertifyId2 & "' And  year(date_work)=" & Format(dtActi.Value, "yyyy") & "  order by gen_jn.cnt")
            If dt.Rows.Count = 0 Then 
                MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            End If
        Catch ex As Exception
            VSysError = True
            MessageBox.Show("Database Error: " & ex.Message, "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim Rpt As New CryNewsJerneralJournal
        If MdShowLOGO = 1 Then
            Rpt.Subreports(0).SetDataSource(RsLOGO)
        End If
        Rpt.SetDataSource(dt)
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.DisplayGroupTree = False
        'FrmPreview.MdiParent = FmMain
        FrmPreview.WindowState = FormWindowState.Maximized
        FrmPreview.ShowDialog()
        FrmPreview.Focus()
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        MdSearchDataList = "FmNsewJeneralJournal"
        'FmRate.ShowDialog()
        Rate_setting.ShowDialog()
    End Sub

    Private Sub ChDe_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChDe.CheckedChanged

    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub


    Private Sub CheckBox2_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged

    End Sub


    Private Sub txtAmt_letter_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmt_letter.TextChanged

    End Sub

    Private Sub txtChecq_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtChecq.TextChanged

    End Sub

    Private Sub txtInvoice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtInvoice.KeyPress
        If e.KeyChar = Chr(13) Then

            'If Len(CStr(txtInvoice.Text).Trim) = 1 Then
            '    txtInvoice.Text = CmbBook.Text & "000000" & CStr(txtInvoice.Text)
            'ElseIf Len(CStr(txtInvoice.Text).Trim) = 2 Then
            '    txtInvoice.Text = CmbBook.Text & "00000" & CStr(txtInvoice.Text)
            'ElseIf Len(CStr(txtInvoice.Text).Trim) = 3 Then
            '    txtInvoice.Text = CmbBook.Text & "0000" & CStr(txtInvoice.Text)
            'ElseIf Len(CStr(txtInvoice.Text).Trim) = 4 Then
            '    txtInvoice.Text = CmbBook.Text & "000" & CStr(txtInvoice.Text)
            'ElseIf Len(CStr(txtInvoice.Text).Trim) = 5 Then
            '    txtInvoice.Text = CmbBook.Text & "00" & CStr(txtInvoice.Text)
            'ElseIf Len(CStr(txtInvoice.Text).Trim) >= 6 Then
            '    txtInvoice.Text = CmbBook.Text & Microsoft.VisualBasic.Right(txtInvoice.Text, 7)
            'End If

            txtAmount.Focus()
        End If
    End Sub

    Private Sub txtInvoice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoice.LostFocus

        'If Len(CStr(txtInvoice.Text).Trim) = 1 Then
        '    txtInvoice.Text = CmbBook.Text & "000000" & CStr(txtInvoice.Text)
        'ElseIf Len(CStr(txtInvoice.Text).Trim) = 2 Then
        '    txtInvoice.Text = CmbBook.Text & "00000" & CStr(txtInvoice.Text)
        'ElseIf Len(CStr(txtInvoice.Text).Trim) = 3 Then
        '    txtInvoice.Text = CmbBook.Text & "0000" & CStr(txtInvoice.Text)
        'ElseIf Len(CStr(txtInvoice.Text).Trim) = 4 Then
        '    txtInvoice.Text = CmbBook.Text & "000" & CStr(txtInvoice.Text)
        'ElseIf Len(CStr(txtInvoice.Text).Trim) = 5 Then
        '    txtInvoice.Text = CmbBook.Text & "00" & CStr(txtInvoice.Text)
        'ElseIf Len(CStr(txtInvoice.Text).Trim) >= 6 Then
        '    txtInvoice.Text = CmbBook.Text & Microsoft.VisualBasic.Right(txtInvoice.Text, 7)
        '    'txtInvoice.Text = CmbBook.Text & CStr(txtInvoice.Text)
        'End If
        'If IsNumeric(Microsoft.VisualBasic.Right(txtInvoice.Text, 7)) = False Then txtInvoice.BackColor = Color.Red : txtInvoice.Focus() : Exit Sub
    End Sub

    Private Sub txtInvoice_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtInvoice.TextChanged

        txtInvoice.BackColor = Color.White

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'If IsNumeric(txtAuto.Text) = False Then txtAuto.Text = "" : Exit Sub
        'If txtAuto.Text = "" Then txtAuto.Text = "0" : Exit Sub
    End Sub

    Private Sub Label25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Autox()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'Call LoadSetRate()
For i = 0 To FG.Rows.Count - 1
            SetGridValue(FG, i, 10, Cmb.Text)
            SetGridValue(FG, i, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
            SetGridValue(FG, i, 12, Format(CDbl(CDbl(txtRate.Text) * CDbl(GetGridValue(FG, i, 5))), "#,##0.00"))
            SetGridValue(FG, i, 13, Format(CDbl(CDbl(txtRate.Text) * CDbl(GetGridValue(FG, i, 6))), "#,##0.00"))
        Next
        '===============
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Set_date_working.ShowDialog()
    End Sub

    Private Sub dtActi_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtActi.ValueChanged
        If txtInvoice.Enabled = True Then
            AutoNumber()
        End If
    End Sub

    Private Sub Off_Usr_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Off_Usr.SelectedIndexChanged

    End Sub

    Private Sub CmbCust_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCust.SelectedIndexChanged
If RadioButton1.Checked Then 
            Dim dtCust As DataTable = DbHelper.GetDataTable("select    * from Customer where 1=1  and Name=N'" & CmbCust.Text & "'   ")
            If dtCust.Rows.Count > 0 Then
                TxtCustID.Text = Trim(dtCust.Rows(0)("Code").ToString())
            Else
                TxtCustID.Text = ""
            End If

        End If
    End Sub

Private Sub CmbSupp_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbSupp.SelectedIndexChanged
        If RadioButton2.Checked Then
            Dim dtSupp As DataTable = DbHelper.GetDataTable("select    * from Supplier where 1=1  and Name=N'" & CmbSupp.Text & "'   ")
            If dtSupp.Rows.Count > 0 Then
                TxtSuppID.Text = Trim(dtSupp.Rows(0)("Code").ToString())
            Else
                TxtSuppID.Text = ""
            End If
   
        End If
    End Sub

Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked Then
            Dim dtSupp As DataTable = DbHelper.GetDataTable("select    * from Supplier where 1=1  and Name=N'" & CmbSupp.Text & "'   ")
            If dtSupp.Rows.Count > 0 Then
                TxtSuppID.Text = Trim(dtSupp.Rows(0)("Code").ToString())
            Else
                TxtSuppID.Text = ""
            End If
          
        Else
            TxtSuppID.Text = ""
        End If
    
        CmbSupp.Enabled = True
        CmbCust.Enabled = False
    End Sub

Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            Dim dtCust As DataTable = DbHelper.GetDataTable("select    * from Customer where 1=1  and Name=N'" & CmbCust.Text & "'   ")
            If dtCust.Rows.Count > 0 Then
                TxtCustID.Text = Trim(dtCust.Rows(0)("Code").ToString())
            Else
                TxtCustID.Text = ""
            End If
       
        Else
            TxtCustID.Text = ""
        End If

        CmbSupp.Enabled = False
        CmbCust.Enabled = True
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked Then
            Panel4.Enabled = True

        Else
            Panel4.Enabled = False
        End If
        RadioButton1.Checked = True
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

        Call ConnectAccess()


        Try
            Dim dt As DataTable = DbHelper.GetDataTable("Select * from Conect")
            If dt.Rows.Count <> 0 Then
                Dim row As DataRow = dt.Rows(0)
                MDServerName = row("ServerName").ToString()
                MDDatabaName = row("DatabaseName").ToString()
                MDServerUser = row("UserName").ToString()
                MDServerPassword = row("UserPassword").ToString()
                MDSeriaAccess = row("PartitionSeria").ToString()
                'SPW = CStr(row("SavePassword").ToString())
                'SUSID = CStr(row("SaveUserID").ToString())
            End If
        Catch ex As Exception
            VSysError = True
            MessageBox.Show("Database Error: " & ex.Message, "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Call ConnectSQL()
        If VSysError = True Then

            'FrmData_server.Show()
            'Me.Hide()
            'Exit Sub
        Else
            MsgBox("Complete")
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        'IsMdiContainer = True
        'Panel4.Visible = False
        'FrmCustomer.MdiParent = Me
        'FrmCustomer.WindowState = FormWindowState.Maximized

        FrmCustomer.ShowDialog()
        FrmCustomer.Focus()

        CmbCust.Items.Clear()
        Call load_Cmb(" SELECT Name  FROM Customer  ORDER BY cnt ", "Name", CmbCust)
        If CmbCust.Items.Count > 0 Then
            CmbCust.SelectedIndex = 0
        End If

        'CmbSupp.Items.Clear()
        'Call load_Cmb(" SELECT Name  FROM Supplier  ORDER BY cnt ", "Name", CmbSupp)
        'If CmbSupp.Items.Count > 0 Then
        '    CmbSupp.SelectedIndex = 0
        'End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        'IsMdiContainer = True
        'Panel4.Visible = False
        'FrmSupplier.MdiParent = Me
        'FrmSupplier.WindowState = FormWindowState.Maximized
        FrmSupplier.ShowDialog()
        FrmSupplier.Focus()

        'CmbCust.Items.Clear()
        'Call load_Cmb(" SELECT Name  FROM Customer  ORDER BY cnt ", "Name", CmbCust)
        'If CmbCust.Items.Count > 0 Then
        '    CmbCust.SelectedIndex = 0
        'End If

        CmbSupp.Items.Clear()
        Call load_Cmb(" SELECT Name  FROM Supplier  ORDER BY cnt ", "Name", CmbSupp)
        If CmbSupp.Items.Count > 0 Then
            CmbSupp.SelectedIndex = 0
        End If
    End Sub

    Private Sub SetupGrid()
        FG.CurrentCell.ColumnIndexumns.Clear()
        FG.CurrentCell.ColumnIndexumns.Add("Col0", "ລ/ດ")
        FG.CurrentCell.ColumnIndexumns.Add("Col1", "ເລກບັນຊີໜີ")
        FG.CurrentCell.ColumnIndexumns.Add("Col2", "ເລກບັນຊີມີ")
        FG.CurrentCell.ColumnIndexumns.Add("Col3", "ຊື່ບັນຊີ (ລາວ)")
        FG.CurrentCell.ColumnIndexumns.Add("Col4", "ຊື່ບັນຊີ (ອັງກິດ)")
        FG.CurrentCell.ColumnIndexumns.Add("Col5", "ຈຳນວນເງິນຈົດໜີ້")
        FG.CurrentCell.ColumnIndexumns.Add("Col6", "ຈຳນວນເງິນຈົດມີ")
        FG.CurrentCell.ColumnIndexumns.Add("Col7", "ລະຫັດ")
        FG.CurrentCell.ColumnIndexumns.Add("Col8", "ຕົ້ນທຶນພາສາ (ລາວ)")
        FG.CurrentCell.ColumnIndexumns.Add("Col9", "ຕົ້ນທຶນພາສາ (ອັງກິດ)")
        FG.CurrentCell.ColumnIndexumns.Add("Col10", "ສະກຸນເງິນເງິນ")
        FG.CurrentCell.ColumnIndexumns.Add("Col11", "ອັດຕາແລກປ່ຽນ")
        FG.CurrentCell.ColumnIndexumns.Add("Col12", "ມູນຄ່າໜີ້")
        FG.CurrentCell.ColumnIndexumns.Add("Col13", "ມູນຄ່າມີ")
        FG.CurrentCell.ColumnIndexumns.Add("Col14", "1111")
        FG.CurrentCell.ColumnIndexumns.Add("Col15", "22")
        
        FG.AllowUserToAddRows = False
    End Sub


End Class