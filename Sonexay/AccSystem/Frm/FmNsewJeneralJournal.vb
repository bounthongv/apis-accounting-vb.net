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
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        fmShartOfAccDetail.ShowDialog()
    End Sub

    Private Sub Savedata()
        Dim J As Integer
        For J = 1 To FG.Rows - 1
            '===============
            'FG.set_TextMatrix(J, 10, Cmb.Text)
            'FG.set_TextMatrix(J, 11, Format(CDbl(txtRate.Text), "#,##0.00"))


            'If Cmb.Text = "USD" Then
            '    FG.set_TextMatrix(J, 14, Format(CDbl(FG.get_TextMatrix(J, 5)), "#,##0.00"))
            '    FG.set_TextMatrix(J, 15, Format(CDbl(FG.get_TextMatrix(J, 6)), "#,##0.00"))
            'Else
            '    FG.set_TextMatrix(J, 14, Format(CDbl(FG.get_TextMatrix(J, 12)) / CDbl(txtRateUSD.Text), "#,##0.00"))
            '    FG.set_TextMatrix(J, 15, Format(CDbl(FG.get_TextMatrix(J, 13)) / CDbl(txtRateUSD.Text), "#,##0.00"))
            '    'FG.set_TextMatrix(FG.Row, 14, Format(CDbl(FG.get_TextMatrix(FG.Row, 12)) / CDbl(txtRateUSD.Text), "#,##0.00"))
            '    'FG.set_TextMatrix(FG.Row, 15, Format(CDbl(FG.get_TextMatrix(FG.Row, 13)) / CDbl(txtRateUSD.Text), "#,##0.00"))
            'End If
            If FG.get_TextMatrix(J, 10) = "USD" Then
                FG.set_TextMatrix(J, 14, Format(CDbl(FG.get_TextMatrix(J, 5)), "#,##0.00"))
                FG.set_TextMatrix(J, 15, Format(CDbl(FG.get_TextMatrix(J, 6)), "#,##0.00"))
            Else
                FG.set_TextMatrix(J, 14, Format(CDbl(FG.get_TextMatrix(J, 12)) / CDbl(txtRate.Text), "#,##0.00"))
                FG.set_TextMatrix(J, 15, Format(CDbl(FG.get_TextMatrix(J, 13)) / CDbl(txtRate.Text), "#,##0.00"))
                'FG.set_TextMatrix(FG.Row, 14, Format(CDbl(FG.get_TextMatrix(FG.Row, 12)) / CDbl(txtRateUSD.Text), "#,##0.00"))
                'FG.set_TextMatrix(FG.Row, 15, Format(CDbl(FG.get_TextMatrix(FG.Row, 13)) / CDbl(txtRateUSD.Text), "#,##0.00"))
            End If

        Next J

        If MuLng = "L" Then
            Amount_In_Word = txtAmt_letter.Text
        Else
            Amount_In_Word = txtAmt_letter_E.Text
        End If
        MuSubOff = Mid(Off_Usr.Text, 1, 5)
        For i = 1 To FG.Rows - 1
            If FG.get_TextMatrix(i, 1) = "" And FG.get_TextMatrix(i, 2) = "" Then
                FG.Rows = 1
                FG.Rows = 2
                AutoNumber()
                Call NewText()
                Exit Sub
            End If
            Autox()
            txtAmount.Text = txtSumAmountDr.Text


            If CheckBox4.Checked = True Then
                MDCust_Supp = 1
            Else
                MDCust_Supp = 0
            End If

            If TxtReferno.Text = "" Then
                TxtReferno.Text = txtInvoice.Text
            End If

            If CheckBox3.Checked = True Then
                If CheckBox4.Checked = True Then
                    Dim KKK As String = "INSERT INTO gen_jn( date_work, ac_Name, book, certify,Referno, cheque_no ,descrip ,descripe ,amount , curr ,rate, Rate_USD, net_amt ,code_dr ,code_cr ,ac_code    ,amt_dr , amt_cr , amt_USD_Dr, amt_USD_Cr , amount_dr ,amount_cr ,  certis, lock ,rec_lock , last_update , last_user , descflag , Cat_ID , certifyID  ,company  ,Office_ID, Cust_Supp, CustID,SuppID , del , AG,Frm) " & _
                                  "Values('" & Format(dtActi.Value, "yyyy/MM/dd") & "',N'" & Trim(FG.get_TextMatrix(i, 3)) & "',N'" & CmbBook.Text & "',N'" & txtInvoice.Text & "',N'" & TxtReferno.Text & "','" & Apostrophe(txtChecq.Text) & "',N'" & Trim(Apostrophe(txtDesc.Text)) & "',N'" & Trim(Apostrophe(txtDescE.Text)) & "'," & CDbl(txtAmount.Text) & ",'" & FG.get_TextMatrix(i, 10) & "'," & CDbl(FG.get_TextMatrix(i, 11)) & "," & CDbl(txtRateUSD.Text) & ",'" & "0" & "','" & FG.get_TextMatrix(i, 1) & "','" & FG.get_TextMatrix(i, 2) & "','" & FG.get_TextMatrix(i, 1) & FG.get_TextMatrix(i, 2) & "'," & CDbl(FG.get_TextMatrix(i, 5)) & "," & CDbl(FG.get_TextMatrix(i, 6)) & "," & CDbl(FG.get_TextMatrix(i, 14)) & "," & CDbl(FG.get_TextMatrix(i, 15)) & "," & CDbl(0) & "," & CDbl(0) & ",'" & "3" & "','" & "4" & "','" & "5" & "','" & dtActi.Text & "','" & MUserID & "','" & "8" & "','" & FG.get_TextMatrix(i, 7) & "','" & MdCertifyId & "','" & MuSubOff & "' ,'" & MuSubOff & "','" & MDCust_Supp & "',N'" & TxtCustID.Text & "',N'" & TxtSuppID.Text & "' , 0,1,0)"
                    CNN.Execute(KKK)
                Else
                    Dim KKK As String = "INSERT INTO gen_jn( date_work, ac_Name, book, certify, Referno, cheque_no ,descrip ,descripe ,amount , curr ,rate, Rate_USD, net_amt ,code_dr ,code_cr ,ac_code    ,amt_dr , amt_cr , amt_USD_Dr, amt_USD_Cr , amount_dr ,amount_cr  ,  certis, lock ,rec_lock , last_update , last_user , descflag , Cat_ID , certifyID  ,company  ,Office_ID, Cust_Supp, CustID,SuppID , del , AG,Frm) " & _
                               "Values('" & Format(dtActi.Value, "yyyy/MM/dd") & "',N'" & FG.get_TextMatrix(i, 3) & "',N'" & CmbBook.Text & "',N'" & txtInvoice.Text & "',N'" & TxtReferno.Text & "','" & Apostrophe(txtChecq.Text) & "',N'" & Trim(Apostrophe(txtDesc.Text)) & "',N'" & Trim(Apostrophe(txtDescE.Text)) & "'," & CDbl(txtAmount.Text) & ",'" & FG.get_TextMatrix(i, 10) & "'," & CDbl(FG.get_TextMatrix(i, 11)) & "," & CDbl(txtRateUSD.Text) & ",'" & "0" & "','" & FG.get_TextMatrix(i, 1) & "','" & FG.get_TextMatrix(i, 2) & "','" & FG.get_TextMatrix(i, 1) & FG.get_TextMatrix(i, 2) & "'," & CDbl(FG.get_TextMatrix(i, 5)) & "," & CDbl(FG.get_TextMatrix(i, 6)) & "," & CDbl(FG.get_TextMatrix(i, 14)) & "," & CDbl(FG.get_TextMatrix(i, 15)) & "," & CDbl(0) & "," & CDbl(0) & ",'" & "3" & "','" & "4" & "','" & "5" & "','" & dtActi.Text & "','" & MUserID & "','" & "8" & "','" & FG.get_TextMatrix(i, 7) & "','" & MdCertifyId & "','" & MuSubOff & "' ,'" & MuSubOff & "','" & MDCust_Supp & "',N'',N'' , 0,1,0)"
                    CNN.Execute(KKK)
                End If

            Else
                If CheckBox4.Checked = True Then
                    Dim KK As String = "INSERT INTO gen_jn( date_work, ac_Name, book, certify, Referno,cheque_no ,descrip ,descripe ,amount , curr ,rate,  Rate_USD,net_amt ,code_dr ,code_cr ,ac_code  , amount_dr ,amount_cr ,amt_dr , amt_cr  ,amt_USD_Dr, amt_USD_Cr ,certis, lock ,rec_lock , last_update , last_user , descflag , Cat_ID , certifyID  ,company   ,Office_ID, Cust_Supp, CustID,SuppID , del, AG,Frm) " & _
                                      "Values('" & Format(dtActi.Value, "yyyy/MM/dd") & "',N'" & FG.get_TextMatrix(i, 3) & "',N'" & CmbBook.Text & "',N'" & txtInvoice.Text & "',N'" & TxtReferno.Text & "','" & Apostrophe(txtChecq.Text) & "',N'" & Trim(Apostrophe(txtDesc.Text)) & "',N'" & Trim(Apostrophe(txtDescE.Text)) & "'," & CDbl(txtAmount.Text) & ",'" & FG.get_TextMatrix(i, 10) & "'," & CDbl(FG.get_TextMatrix(i, 11)) & "," & CDbl(txtRateUSD.Text) & ",'" & "0" & "','" & FG.get_TextMatrix(i, 1) & "','" & FG.get_TextMatrix(i, 2) & "','" & FG.get_TextMatrix(i, 1) & FG.get_TextMatrix(i, 2) & "'," & CDbl(FG.get_TextMatrix(i, 5)) & "," & CDbl(FG.get_TextMatrix(i, 6)) & "," & CDbl(FG.get_TextMatrix(i, 12)) & "," & CDbl(FG.get_TextMatrix(i, 13)) & "," & CDbl(FG.get_TextMatrix(i, 14)) & "," & CDbl(FG.get_TextMatrix(i, 15)) & ",'" & "3" & "','" & "4" & "','" & "5" & "','" & dtActi.Text & "','" & MUserID & "','" & "8" & "','" & FG.get_TextMatrix(i, 7) & "','" & MdCertifyId & "','" & MuSubOff & "' , '" & MuSubOff & "','" & MDCust_Supp & "',N'" & TxtCustID.Text & "',N'" & TxtSuppID.Text & "' , 0,0,0)"
                    CNN.Execute(KK)
                Else
                    Dim KK As String = "INSERT INTO gen_jn( date_work, ac_Name, book, certify, Referno,cheque_no ,descrip ,descripe ,amount , curr ,rate,  Rate_USD,net_amt ,code_dr ,code_cr ,ac_code  , amount_dr ,amount_cr ,amt_dr , amt_cr  ,amt_USD_Dr, amt_USD_Cr , certis, lock ,rec_lock , last_update , last_user , descflag , Cat_ID , certifyID  ,company   ,Office_ID, Cust_Supp, CustID, SuppID , del, AG,Frm) " & _
                          "Values('" & Format(dtActi.Value, "yyyy/MM/dd") & "',N'" & FG.get_TextMatrix(i, 3) & "',N'" & CmbBook.Text & "',N'" & txtInvoice.Text & "',N'" & TxtReferno.Text & "','" & Apostrophe(txtChecq.Text) & "',N'" & Trim(Apostrophe(txtDesc.Text)) & "',N'" & Trim(Apostrophe(txtDescE.Text)) & "'," & CDbl(txtAmount.Text) & ",'" & FG.get_TextMatrix(i, 10) & "'," & CDbl(FG.get_TextMatrix(i, 11)) & "," & CDbl(txtRateUSD.Text) & ",'" & "0" & "','" & FG.get_TextMatrix(i, 1) & "','" & FG.get_TextMatrix(i, 2) & "','" & FG.get_TextMatrix(i, 1) & FG.get_TextMatrix(i, 2) & "'," & CDbl(FG.get_TextMatrix(i, 5)) & "," & CDbl(FG.get_TextMatrix(i, 6)) & "," & CDbl(FG.get_TextMatrix(i, 12)) & "," & CDbl(FG.get_TextMatrix(i, 13)) & "," & CDbl(FG.get_TextMatrix(i, 14)) & "," & CDbl(FG.get_TextMatrix(i, 15)) & ",'" & "3" & "','" & "4" & "','" & "5" & "','" & dtActi.Text & "','" & MUserID & "','" & "8" & "','" & FG.get_TextMatrix(i, 7) & "','" & MdCertifyId & "','" & MuSubOff & "' , '" & MuSubOff & "','" & MDCust_Supp & "',N'',N'' , 0,0,0)"
                    CNN.Execute(KK)

                End If
            End If

        Next i
        MuSubOff = MuSubOff2
        LngId = "6001" : MsgRpt()
    End Sub

    Private Sub FG_AfterEdit(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_AfterEditEvent) Handles FG.AfterEdit
        If FG.Col = 1 Or FG.Col = 2 Then
            If FG.get_TextMatrix(FG.Row, 1) & FG.get_TextMatrix(FG.Row, 2) <> "" Then
                'If Len(FG.get_TextMatrix(FG.Row, 1) & FG.get_TextMatrix(FG.Row, 2)) <> 7 Then
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
            txtDesc.Text = FG.get_TextMatrix(1, 3)
            txtDescE.Text = FG.get_TextMatrix(1, 4)
        End If
        If CDbl(txtAmount.Text) = 0 Then
            If FG.Row = 2 Then
                txtAmount.Text = Format(CDbl((CDbl(FG.get_TextMatrix(1, 5)) + CDbl(FG.get_TextMatrix(1, 6)))), "##,##0.00")
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
        If FG.Col = 1 Then
            R = FG.Row()
            L = FG.Col
            If FG.get_TextMatrix(FG.Row, 1) = "" Then
                MDSearchAcccode = (FG.get_TextMatrix(FG.Row, 1))
                fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                fmShartOfAccDetail.ShowDialog()
            End If
            If FG.get_TextMatrix(FG.Row, 1) <> "" Then
                AccId = FG.get_TextMatrix(FG.Row, 1)
                MDSearchAcccode = (FG.get_TextMatrix(FG.Row, 1))
                LoadText()
                FG.set_TextMatrix(FG.Row, 3, AccName)
                FG.set_TextMatrix(FG.Row, 4, AccNamee)
                If FG.get_TextMatrix(FG.Row, 3) = "" Then
                    FG.set_TextMatrix(FG.Row, 1, "")
                    fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                    fmShartOfAccDetail.ShowDialog()
                End If
                SumAmountDr()
                txtSumAmountDr.Text = CDbl(txtSumAmountDr.Text) - CDbl(FG.get_TextMatrix(FG.Row, 5))
                FG.set_TextMatrix(FG.Row, 5, Format(CDbl(CDbl(txtAmount.Text) - CDbl(txtSumAmountDr.Text)), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 6, "0.00")
                FG.set_TextMatrix(FG.Row, 7, 0)
                FG.set_TextMatrix(FG.Row, 10, Cmb.Text)
                FG.set_TextMatrix(FG.Row, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 12, Format(CDbl(CDbl(txtRate.Text) * FG.get_TextMatrix(FG.Row, 5)), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 13, "0.00")
                FG.set_TextMatrix(FG.Row, 14, Format(CDbl(CDbl(FG.get_TextMatrix(FG.Row, 12))) / CDbl(MDUSD_LAK), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 15, "0.00")
                SumAmountDr()

                If ChDe.Checked = True Then
                    If MuLng = "L" Then
                        FG.Col = 3
                    End If
                    If MuLng = "E" Then
                        FG.Col = 4
                    End If
                    If FG.get_TextMatrix(FG.Rows - 1, 3) <> "" Then
                        FG.Rows = FG.Rows + 1
                    End If
                    Exit Sub
                End If

                FG.Col = 5
                If FG.get_TextMatrix(FG.Rows - 1, 3) <> "" Then
                    FG.Rows = FG.Rows + 1
                    Exit Sub
                End If
                Exit Sub
            End If
            Exit Sub
        End If
        '*************************Col-2-*********************
        If FG.Col = 2 Then
            R = FG.Row()
            L = FG.Col
            If FG.get_TextMatrix(FG.Row, 2) = "" Then
                MDSearchAcccode = (FG.get_TextMatrix(FG.Row, 2))
                fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                fmShartOfAccDetail.ShowDialog()
            End If
            If FG.get_TextMatrix(FG.Row, 2) <> "" Then
                AccId = FG.get_TextMatrix(FG.Row, FG.Col)
                LoadText()
                FG.set_TextMatrix(FG.Row, 3, AccName)
                FG.set_TextMatrix(FG.Row, 4, AccNamee)
                If FG.get_TextMatrix(FG.Row, 3) = "" Then
                    MDSearchAcccode = (FG.get_TextMatrix(FG.Row, 2))
                    fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                    fmShartOfAccDetail.ShowDialog()
                End If
                SumAmountDr()
                txtSumAmountCr.Text = CDbl(txtSumAmountCr.Text) - CDbl(FG.get_TextMatrix(FG.Row, 6))
                FG.set_TextMatrix(FG.Row, 6, Format(CDbl(CDbl(txtAmount.Text) - CDbl(txtSumAmountCr.Text)), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 5, "0.00")
                FG.set_TextMatrix(FG.Row, 7, 0)
                FG.set_TextMatrix(FG.Row, 10, Cmb.Text)
                FG.set_TextMatrix(FG.Row, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 13, Format(CDbl(CDbl(txtRate.Text) * FG.get_TextMatrix(FG.Row, 6)), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 12, "0.00")
                FG.set_TextMatrix(FG.Row, 15, Format(CDbl(CDbl(FG.get_TextMatrix(FG.Row, 13))) / CDbl(MDUSD_LAK), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 14, "0.00")
                SumAmountDr()
                If ChDe.Checked = True Then
                    If MuLng = "L" Then
                        FG.Col = 3
                    End If
                    If MuLng = "E" Then
                        FG.Col = 4
                    End If
                    If FG.get_TextMatrix(FG.Rows - 1, 3) <> "" Then
                        FG.Rows = FG.Rows + 1
                    End If
                    Exit Sub
                End If
                FG.Col = 6
                If FG.get_TextMatrix(FG.Rows - 1, 3) <> "" Then
                    FG.Rows = FG.Rows + 1
                    FG.Col = 6
                End If
                Exit Sub
            End If
            Exit Sub
        End If


        '====*************************

        If FG.Col = 3 Then
            If Button3.Text = "ໂຊຂໍ້ມູນແບບທົ່ວໄປ" Then
                FG.Col = 4
                Exit Sub
            Else
                If FG.get_TextMatrix(FG.Row, 1) <> "" Then
                    FG.Col = 5
                    Exit Sub
                End If
                If FG.get_TextMatrix(FG.Row, 2) <> "" Then
                    FG.Col = 6
                    Exit Sub
                End If
                'MsgBox("ໂຊຂໍ້ມູນແບບລະອຽດ")
            End If
        End If
        If FG.Col = 4 Then
            If FG.get_TextMatrix(FG.Row, 1) <> "" Then
                FG.Col = 5
                Exit Sub
            End If
            If FG.get_TextMatrix(FG.Row, 2) <> "" Then
                FG.Col = 6
                Exit Sub
            End If
            'MsgBox("ໂຊຂໍ້ມູນແບບລະອຽດ")
        End If
        '====*************************




        '*************************Col-5-*********************
        If FG.Col = 6 Then
            If IsNumeric(FG.get_TextMatrix(FG.Row, 6)) = False Then
                MessageBox.Show("ກະລຸນນາໃສ່ໂຕເລກ")
                FG.set_TextMatrix(FG.Row, 6, "0.00")
                Exit Sub
            End If
            If CDbl(FG.get_TextMatrix(FG.Row, 6)) = 0 Then
                MessageBox.Show("ກະລຸນນາມູນຄ່າກ່ອນ")
                FG.set_TextMatrix(FG.Row, 6, "0.00")
                Exit Sub
            End If
            FG.set_TextMatrix(FG.Row, 6, Format(CDbl(FG.get_TextMatrix(FG.Row, FG.Col)), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 5, "0.00")
            FG.set_TextMatrix(FG.Row, 7, 0)
            FG.set_TextMatrix(FG.Row, 10, Cmb.Text)
            FG.set_TextMatrix(FG.Row, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 13, Format(CDbl(CDbl(txtRate.Text) * FG.get_TextMatrix(FG.Row, 6)), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 12, "0.00")
            FG.set_TextMatrix(FG.Row, 15, Format(CDbl(CDbl(FG.get_TextMatrix(FG.Row, 13))) / CDbl(MDUSD_LAK), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 14, "0.00")
            SumAmountDr()
            FG.Col = 7
            Exit Sub
        End If
        '*************************Col-4*********************
        If FG.Col = 5 Then
            If IsNumeric(FG.get_TextMatrix(FG.Row, 5)) = False Then
                MessageBox.Show("ກະລຸນນາໃສ່ໂຕເລກ")
                FG.set_TextMatrix(FG.Row, 5, "0.00")
                Exit Sub
            End If

            If CDbl(FG.get_TextMatrix(FG.Row, 5)) = 0 Then
                MessageBox.Show("ກະລຸນນາມູນຄ່າກ່ອນ")
                FG.set_TextMatrix(FG.Row, 5, "0.00")
                Exit Sub
            End If

            FG.set_TextMatrix(FG.Row, 5, Format(CDbl(FG.get_TextMatrix(FG.Row, FG.Col)), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 6, "0.00")
            FG.set_TextMatrix(FG.Row, 7, 0)
            FG.set_TextMatrix(FG.Row, 10, Cmb.Text)
            FG.set_TextMatrix(FG.Row, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 12, Format(CDbl(CDbl(txtRate.Text) * FG.get_TextMatrix(FG.Row, 5)), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 13, "0.00")
            FG.set_TextMatrix(FG.Row, 14, Format(CDbl(CDbl(FG.get_TextMatrix(FG.Row, 12))) / CDbl(MDUSD_LAK), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 15, "0.00")
            SumAmountDr()
            FG.Col = 7
            Exit Sub
        End If
        '*************************Col-7-*********************
        If FG.Col = 7 Then
            If IsNumeric(FG.get_TextMatrix(FG.Row, 5)) = False Then
                MessageBox.Show("ກະລຸນນາໃສ່ໂຕເລກ")
                FG.set_TextMatrix(FG.Row, 5, "0.00")
                Exit Sub
            End If
            If FG.get_TextMatrix(FG.Row, 7) = "0" Then
                FG.set_TextMatrix(FG.Row, 8, "ບໍ່ເລືອກ")
                FG.set_TextMatrix(FG.Row, 9, "No Selete")
            ElseIf FG.get_TextMatrix(FG.Row, 7) = "1" Then
                FG.set_TextMatrix(FG.Row, 8, "ຮັບໃຊ້ການພະລິດ")
                FG.set_TextMatrix(FG.Row, 9, "Use build")

            ElseIf FG.get_TextMatrix(FG.Row, 7) = "2" Then
                FG.set_TextMatrix(FG.Row, 8, "ຮັບໃຊ້ການຈຳໜ່າຍ")
                FG.set_TextMatrix(FG.Row, 9, "Use Sell")
            ElseIf FG.get_TextMatrix(FG.Row, 7) = "3" Then
                FG.set_TextMatrix(FG.Row, 8, "ຮັບໃຊ້ບໍລິຫານ")
                FG.set_TextMatrix(FG.Row, 9, "Use manage ")

            ElseIf FG.get_TextMatrix(FG.Row, 7) = "4" Then
                FG.set_TextMatrix(FG.Row, 8, "ຕົ້ນທຶນຂາຍ/ຕົ້ນທຶນບໍລິຫານ")
                FG.set_TextMatrix(FG.Row, 9, "Sell capital/manage capital ")
            ElseIf FG.get_TextMatrix(FG.Row, 7) > 4 Then
                MessageBox.Show("ລະຫັດບໍ່ຖືກຕ້ອງ ລະຫັດມີພຽງເລກ 0 ຫາເລກ 4 ເທົ່ານັ້ນ")
                FG.set_TextMatrix(FG.Row, 8, "ບໍ່ລເລືອກ")
                FG.set_TextMatrix(FG.Row, 9, "No Selete")
                FG.set_TextMatrix(FG.Row, 7, "0")
                Exit Sub
            ElseIf FG.get_TextMatrix(FG.Row, 7) < 0 Then
                MessageBox.Show("ລະຫັດບໍ່ຖືກຕ້ອງ ລະຫັດມີພຽງເລກ 0 ຫາເລກ 4 ເທົ່ານັ້ນ")
                FG.set_TextMatrix(FG.Row, 8, "ບໍ່ລເລືອກ")
                FG.set_TextMatrix(FG.Row, 9, "No Selete")
                FG.set_TextMatrix(FG.Row, 7, "0")
                Exit Sub
            End If
            'aaaaaaaaaaaa
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

                                        Dim srNum As New ADODB.Recordset
                                        Dim mNum As Integer = 0
                                        If IsNumeric(Microsoft.VisualBasic.Right(txtInvoice.Text, 3)) = False Then MsgBox("3 ໂຕທາງທ້າຍຕ້ອງເປັນໂຕເລກເທົານັ້ນ") : txtInvoice.BackColor = Color.Red : txtInvoice.Focus() : Exit Sub
                                        'Call LoadSqlData("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book ='" & CmbBook.Text & "' And    year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And month(date_work)=" & Format(CDate(dtActi.Value), "MM") & " Order by  Right(certify,7) DESC ", srNum)
                                        Call LoadSqlData("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book =N'" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "'And  ReferNO = N'" & TxtReferno.Text & "' And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And  month(date_work)=" & Format(CDate(dtActi.Value), "MM") & " And  day(date_work)=" & Format(CDate(dtActi.Value), "dd") & "  Order by  Right(certify,3) DESC ", RSC)

                                        If srNum.RecordCount = 0 Then
                                            mNum = 0
                                        Else
                                            mNum = Val(srNum.Fields("certify").Value.ToString)
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
                                        Call LoadSqlData("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book =N'" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "'And  ReferNO = N'" & TxtReferno.Text & "' And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And  month(date_work)=" & Format(CDate(dtActi.Value), "MM") & " And  day(date_work)=" & Format(CDate(dtActi.Value), "dd") & "  Order by  Right(certify,3) DESC ", RSC)

                                        If RSC.RecordCount > 0 Then
                                            MsgBox("ເລກລະຫັດ : " & Trim(txtInvoice.Text) & " ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
                                            txtInvoice.BackColor = Color.Red
                                            txtInvoice.Focus()
                                            If RSC.State = ConnectionState.Open Then RSC.Close()
                                            Exit Sub
                                        End If

                                        Savedata()
                                    Else
                                        'CNN.Execute("DELETE FROM gen_jn WHERE book ='" & FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 16) & "' And certify  = '" & txtInvoice.Text & "'   And  year(date_work)='" & Format(CDate(FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 1)), "yyyy") & "' ")
                                        CNN.Execute("DELETE FROM gen_jn WHERE book =N'" & FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 16) & "' And certify  =N'" & txtInvoice.Text & "'  And  ReferNO = N'" & TxtReferno.Text & "'  And   date_work='" & Format(CDate(FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 1)), "yyyy-MM-dd") & "' ")

                                        Savedata()
                                    End If
                                    If CheckBox1.Checked = True Then
                                        Call LoadReport()
                                    End If
                                    txtInvoice.Enabled = True
                                    CmbBook.Enabled = True
                                    Panel1.Visible = False
                                    BtnMove.Visible = False
                                    BtnSearch.Visible = False
                                    FG.Rows = 1
                                    FG.Rows = 2
                                    FG.Row = 1
                                    FG.Col = 1

                                    If CheckBox2.Checked = True Then
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


                If FG.get_TextMatrix(FG.Row, 1) <> "" Then
                    If FG.get_TextMatrix(FG.Row + 1, 1) <> "" Then
                        FG.Row = FG.Row + 1
                        FG.Col = 1
                        Exit Sub
                    End If
                    FG.Row = FG.Row + 1
                    FG.Col = 2
                    SumAmountDr()
                    '*****************
                Else
                    If FG.get_TextMatrix(FG.Row + 1, 2) <> "" Then
                        FG.Row = FG.Row + 1
                        FG.Col = 2
                        Exit Sub
                    End If
                    FG.Row = FG.Row + 1
                    FG.Col = 1
                    SumAmountDr()
                End If

                Exit Sub
            End If
        End If
        CmbBook.Focus()
    End Sub
    Private Sub AfterEdit2()
        BtnMove.Visible = False
        '*************************Col-1-*********************
        If FG.Col = 1 Then
            R = FG.Row()
            L = FG.Col
            If FG.get_TextMatrix(FG.Row, 1) = "" Then
                MDSearchAcccode = (FG.get_TextMatrix(FG.Row, 1))
                AccId = FG.get_TextMatrix(FG.Row, 1)

                fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                fmShartOfAccDetail.ShowDialog()

            End If
            If FG.get_TextMatrix(FG.Row, 1) <> "" Then
                AccId = FG.get_TextMatrix(FG.Row, 1)
                MDSearchAcccode = (FG.get_TextMatrix(FG.Row, 1))


                LoadText()
                FG.set_TextMatrix(FG.Row, 3, AccName)
                FG.set_TextMatrix(FG.Row, 4, AccNamee)
                If FG.get_TextMatrix(FG.Row, 3) = "" Then

                    FG.set_TextMatrix(FG.Row, 1, "")
                    fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                    fmShartOfAccDetail.ShowDialog()
                    LoadText()

                End If
                SumAmountDr()
                txtSumAmountDr.Text = CDbl(txtSumAmountDr.Text) - CDbl(FG.get_TextMatrix(FG.Row, 5))
                FG.set_TextMatrix(FG.Row, 5, Format(CDbl(CDbl(txtAmount.Text) - CDbl(txtSumAmountDr.Text)), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 6, "0.00")
                FG.set_TextMatrix(FG.Row, 7, 0)
                FG.set_TextMatrix(FG.Row, 10, Cmb.Text)
                FG.set_TextMatrix(FG.Row, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 12, Format(CDbl(CDbl(txtRate.Text) * FG.get_TextMatrix(FG.Row, 5)), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 13, "0.00")
                FG.set_TextMatrix(FG.Row, 14, Format(CDbl(CDbl(FG.get_TextMatrix(FG.Row, 12))) / CDbl(MDUSD_LAK), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 15, "0.00")
                SumAmountDr()

                If ChDe.Checked = True Then
                    If MuLng = "L" Then
                        FG.Col = 3
                    End If
                    If MuLng = "E" Then
                        FG.Col = 4
                    End If


                    If FG.get_TextMatrix(FG.Rows - 1, 3) <> "" Then
                        FG.Rows = FG.Rows + 1
                    End If
                    Exit Sub
                End If

                FG.Col = 5
                If FG.get_TextMatrix(FG.Rows - 1, 3) <> "" Then
                    FG.Rows = FG.Rows + 1
                    Exit Sub
                End If
                Exit Sub
            End If
            Exit Sub
        End If
        '*************************Col-2-*********************
        If FG.Col = 2 Then
            R = FG.Row()
            L = FG.Col
            If FG.get_TextMatrix(FG.Row, 2) = "" Then
                MDSearchAcccode = (FG.get_TextMatrix(FG.Row, 2))
                fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                fmShartOfAccDetail.ShowDialog()
            End If
            If FG.get_TextMatrix(FG.Row, 2) <> "" Then
                AccId = FG.get_TextMatrix(FG.Row, FG.Col)
                LoadText()
                FG.set_TextMatrix(FG.Row, 3, AccName)
                FG.set_TextMatrix(FG.Row, 4, AccNamee)
                If FG.get_TextMatrix(FG.Row, 3) = "" Then
                    MDSearchAcccode = (FG.get_TextMatrix(FG.Row, 2))
                    fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                    fmShartOfAccDetail.ShowDialog()
                End If
                SumAmountDr()
                txtSumAmountCr.Text = CDbl(txtSumAmountCr.Text) - CDbl(FG.get_TextMatrix(FG.Row, 6))
                FG.set_TextMatrix(FG.Row, 6, Format(CDbl(CDbl(txtAmount.Text) - CDbl(txtSumAmountCr.Text)), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 5, "0.00")
                FG.set_TextMatrix(FG.Row, 7, 0)
                FG.set_TextMatrix(FG.Row, 10, Cmb.Text)
                FG.set_TextMatrix(FG.Row, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 13, Format(CDbl(CDbl(txtRate.Text) * FG.get_TextMatrix(FG.Row, 6)), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 12, "0.00")
                FG.set_TextMatrix(FG.Row, 15, Format(CDbl(CDbl(FG.get_TextMatrix(FG.Row, 13))) / CDbl(MDUSD_LAK), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 14, "0.00")
                SumAmountDr()

                If ChDe.Checked = True Then
                    If MuLng = "L" Then
                        FG.Col = 3
                    End If
                    If MuLng = "E" Then
                        FG.Col = 4
                    End If
                    If FG.get_TextMatrix(FG.Rows - 1, 3) <> "" Then
                        FG.Rows = FG.Rows + 1
                    End If
                    Exit Sub
                End If
                FG.Col = 6
                If FG.get_TextMatrix(FG.Rows - 1, 3) <> "" Then
                    FG.Rows = FG.Rows + 1
                    FG.Col = 6
                End If
                Exit Sub
            End If
            Exit Sub
        End If

        '====*************************

        If FG.Col = 3 Then
            If Button3.Text = "ໂຊຂໍ້ມູນແບບທົ່ວໄປ" Then
                FG.Col = 4
                Exit Sub
            Else
                If FG.get_TextMatrix(FG.Row, 1) <> "" Then
                    FG.Col = 5
                    Exit Sub
                End If
                If FG.get_TextMatrix(FG.Row, 2) <> "" Then
                    FG.Col = 6
                    Exit Sub
                End If
                'MsgBox("ໂຊຂໍ້ມູນແບບລະອຽດ")
            End If
        End If
        If FG.Col = 4 Then
            If FG.get_TextMatrix(FG.Row, 1) <> "" Then
                FG.Col = 5
                Exit Sub
            End If
            If FG.get_TextMatrix(FG.Row, 2) <> "" Then
                FG.Col = 6
                Exit Sub
            End If
            'MsgBox("ໂຊຂໍ້ມູນແບບລະອຽດ")
        End If

        '*************************Col-5-*********************
        If FG.Col = 6 Then
            If IsNumeric(FG.get_TextMatrix(FG.Row, 6)) = False Then
                MessageBox.Show("ກະລຸນນາໃສ່ໂຕເລກ")
                FG.set_TextMatrix(FG.Row, 6, "0.00")
                Exit Sub
            End If
            If CDbl(FG.get_TextMatrix(FG.Row, 6)) = 0 Then
                MessageBox.Show("ກະລຸນນາມູນຄ່າກ່ອນ")
                FG.set_TextMatrix(FG.Row, 6, "0.00")
                Exit Sub
            End If
            ACCode = FG.get_TextMatrix(FG.Row, 2)
            FG.set_TextMatrix(FG.Row, 6, Format(CDbl(FG.get_TextMatrix(FG.Row, FG.Col)), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 5, "0.00")
            FG.set_TextMatrix(FG.Row, 7, 0)
            FG.set_TextMatrix(FG.Row, 10, Cmb.Text)
            FG.set_TextMatrix(FG.Row, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 13, Format(CDbl(CDbl(txtRate.Text) * FG.get_TextMatrix(FG.Row, 6)), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 12, "0.00")
            FG.set_TextMatrix(FG.Row, 15, Format(CDbl(CDbl(FG.get_TextMatrix(FG.Row, 13))) / CDbl(MDUSD_LAK), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 14, "0.00")
            SumAmountDr()
            'Remain = CDbl(FG.get_TextMatrix(FG.Row, 13))
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
                                    '    SumDr = CDbl(FG.get_TextMatrix(FG.Row, 13))
                                    '    If CDbl(SumDr) > CDbl(Remain) Then
                                    '        FG.set_TextMatrix(FG.Row, 13, Format(CDbl(Remain), "#,##0.00"))
                                    '        FG.set_TextMatrix(FG.Row, 6, Format(CDbl(Remain) / CDbl(txtRate.Text), "#,##0.00"))
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

                                            Dim srNum As New ADODB.Recordset
                                            Dim mNum As Integer = 0
                                            If IsNumeric(Microsoft.VisualBasic.Right(txtInvoice.Text, 3)) = False Then MsgBox("3 ໂຕທາງທ້າຍຕ້ອງເປັນໂຕເລກເທົານັ້ນ") : txtInvoice.BackColor = Color.Red : txtInvoice.Focus() : Exit Sub
                                            Call LoadSqlData("SELECT  top 1 Right(certify,3) As  certify  FROM Gen_jn WHERE   book ='" & CmbBook.Text & "' And    year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And month(date_work)=" & Format(CDate(dtActi.Value), "MM") & "  Order by  Right(certify,7) DESC ", srNum)
                                            If srNum.RecordCount = 0 Then
                                                mNum = 0
                                            Else
                                                mNum = Val(srNum.Fields("certify").Value.ToString)
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


                                            Call LoadSqlData("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book ='" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "' And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And  month(date_work)=" & Format(CDate(dtActi.Value), "MM") & " Order by  Right(certify,7) DESC ", RSC)
                                            If RSC.RecordCount > 0 Then
                                                MsgBox("ເລກລະຫັດ : " & Trim(txtInvoice.Text) & " ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
                                                txtInvoice.BackColor = Color.Red
                                                txtInvoice.Focus()
                                                If RSC.State = ConnectionState.Open Then RSC.Close()
                                                Exit Sub
                                            End If

                                            Savedata()
                                        Else
                                            CNN.Execute("DELETE FROM gen_jn WHERE book ='" & FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 16) & "' And certify  = '" & txtInvoice.Text & "'   And  year(date_work)='" & Format(CDate(FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 1)), "yyyy") & "' ")
                                            Savedata()
                                        End If

                                        If CheckBox1.Checked = True Then

                                            Call LoadReport()
                                            'MsgBox(11)
                                        End If

                                        txtInvoice.Enabled = True
                                        CmbBook.Enabled = True
                                        Panel1.Visible = False
                                        BtnMove.Visible = False
                                        BtnSearch.Visible = False
                                        FG.Rows = 1
                                        FG.Rows = 2
                                        FG.Row = 1
                                        FG.Col = 1
                                        CmbBook.Focus()


                                        If CheckBox2.Checked = True Then

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

                If FG.get_TextMatrix(FG.Row, 1) <> "" Then
                    If FG.get_TextMatrix(FG.Row + 1, 1) <> "" Then
                        FG.Row = FG.Row + 1
                        FG.Col = 1
                        Exit Sub
                    End If
                    FG.Row = FG.Row + 1
                    FG.Col = 2
                    SumAmountDr()
                    '*****************
                Else
                    If FG.get_TextMatrix(FG.Row + 1, 2) <> "" Then
                        FG.Row = FG.Row + 1
                        FG.Col = 2
                        Exit Sub
                    End If
                    FG.Row = FG.Row + 1
                    FG.Col = 1
                    SumAmountDr()
                End If
                Exit Sub
            End If
        End If
        '*************************Col-4*********************
        If FG.Col = 5 Then
            If IsNumeric(FG.get_TextMatrix(FG.Row, 5)) = False Then
                MessageBox.Show("ກະລຸນນາໃສ່ໂຕເລກ")
                FG.set_TextMatrix(FG.Row, 5, "0.00")
                Exit Sub
            End If

            If CDbl(FG.get_TextMatrix(FG.Row, 5)) = 0 Then
                MessageBox.Show("ກະລຸນນາມູນຄ່າກ່ອນ")
                FG.set_TextMatrix(FG.Row, 5, "0.00")
                Exit Sub
            End If
            ACCode = FG.get_TextMatrix(FG.Row, 1)
            FG.set_TextMatrix(FG.Row, 5, Format(CDbl(FG.get_TextMatrix(FG.Row, FG.Col)), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 6, "0.00")
            FG.set_TextMatrix(FG.Row, 7, 0)
            FG.set_TextMatrix(FG.Row, 10, Cmb.Text)
            FG.set_TextMatrix(FG.Row, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 12, Format(CDbl(CDbl(txtRate.Text) * FG.get_TextMatrix(FG.Row, 5)), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 13, "0.00")
            FG.set_TextMatrix(FG.Row, 14, Format(CDbl(CDbl(FG.get_TextMatrix(FG.Row, 12))) / CDbl(MDUSD_LAK), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 15, "0.00")
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

                                        Dim srNum As New ADODB.Recordset
                                        Dim mNum As Integer = 0
                                        If IsNumeric(Microsoft.VisualBasic.Right(txtInvoice.Text, 3)) = False Then MsgBox("3 ໂຕທາງທ້າຍຕ້ອງເປັນໂຕເລກເທົານັ້ນ") : txtInvoice.BackColor = Color.Red : txtInvoice.Focus() : Exit Sub
                                        Call LoadSqlData("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book ='" & CmbBook.Text & "' And year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And Month(date_work)=" & Format(CDate(dtActi.Value), "MM") & "  Order by  Right(certify,7) DESC ", srNum)
                                        If srNum.RecordCount = 0 Then
                                            mNum = 0
                                        Else
                                            mNum = Val(srNum.Fields("certify").Value.ToString)
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

                                        Call LoadSqlData("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book ='" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "' And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And Month(date_work)=" & Format(CDate(dtActi.Value), "MM") & " Order by  Right(certify,7) DESC ", RSC)
                                        If RSC.RecordCount > 0 Then
                                            MsgBox("ເລກລະຫັດ : " & Trim(txtInvoice.Text) & " ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
                                            txtInvoice.BackColor = Color.Red
                                            txtInvoice.Focus()
                                            If RSC.State = ConnectionState.Open Then RSC.Close()
                                            Exit Sub
                                        End If


                                        Savedata()
                                    Else

                                        CNN.Execute("DELETE FROM gen_jn WHERE book ='" & FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 16) & "' And certify  = '" & txtInvoice.Text & "'   And  year(date_work)='" & Format(CDate(FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 1)), "yyyy") & "' ")
                                        Savedata()
                                    End If
                                    If CheckBox1.Checked = True Then
                                        Call LoadReport()
                                    End If
                                    txtInvoice.Enabled = True
                                    CmbBook.Enabled = True
                                    Panel1.Visible = False
                                    BtnMove.Visible = False
                                    BtnSearch.Visible = False
                                    FG.Rows = 1
                                    FG.Rows = 2
                                    FG.Row = 1
                                    FG.Col = 1
                                    Exit Sub

                                    If CheckBox2.Checked = True Then
                                        Close()
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                If FG.get_TextMatrix(FG.Row, 1) <> "" Then
                    If FG.get_TextMatrix(FG.Row + 1, 1) <> "" Then
                        FG.Row = FG.Row + 1
                        FG.Col = 1
                        Exit Sub
                    End If
                    FG.Row = FG.Row + 1
                    FG.Col = 2
                    SumAmountDr()
                    '*****************
                Else
                    If FG.get_TextMatrix(FG.Row + 1, 2) <> "" Then
                        FG.Row = FG.Row + 1
                        FG.Col = 2
                        Exit Sub
                    End If
                    FG.Row = FG.Row + 1
                    FG.Col = 1
                    SumAmountDr()
                End If
                Exit Sub
            End If
        End If
        '*************************Col-7-*********************
        If FG.Col = 7 Then
            'If FG.get_TextMatrix(FG.Row, 1) <> "" Then
            '    If FG.get_TextMatrix(FG.Row + 1, 1) <> "" Then
            '        FG.Row = FG.Row + 1
            '        FG.Col = 1
            '        Exit Sub
            '    End If
            '    FG.Row = FG.Row + 1
            '    FG.Col = 2
            '    SumAmountDr()
            '    '*****************
            'Else
            '    If FG.get_TextMatrix(FG.Row + 1, 2) <> "" Then
            '        FG.Row = FG.Row + 1
            '        FG.Col = 2
            '        Exit Sub
            '    End If
            '    FG.Row = FG.Row + 1
            '    FG.Col = 1
            '    SumAmountDr()
            'End If
            Exit Sub
        End If

    End Sub
    Public Sub AutoNumber()

        Dim srNum As New ADODB.Recordset
        Dim mNum As Integer

        'Call LoadSqlData("SELECT top 1 Right(certify,7) As  certify  FROM gen_jn where year(date_work)='" & Format(dtActi.Value, "yyyy") & "'   Order by   Right(certify,7) DESC", srNum)
        Dim ss As String = ""
        ss = "SELECT top 1 Right(certify,3) As  certify   FROM  gen_jn where Frm=0 and book =N'" & CmbBook.Text & "' And  year(date_work)='" & Format(dtActi.Value, "yyyy") & "'  " & _
        " And  month(date_work)='" & Format(dtActi.Value, "MM") & "'  and LEFT(company,2)='" & Off_Id & "' Order by  Right(certify,3) DESC"
        Call LoadSqlData(ss, srNum)
        If srNum.RecordCount = 0 Then
            MdCertifyId = "001"
        Else
            mNum = Val(srNum.Fields("certify").Value.ToString)
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
        FG.Row = R
        FG.Col = L + 4
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
        Dim Comm As ADODB.Command
        Dim rsat As New ADODB.Recordset
        Comm = New ADODB.Command
        Comm.ActiveConnection = CNN
        Cmb.Items.Clear()
        Comm.CommandText = "SELECT Curr FROM Ap_RateSeting WHERE Curr <> '" & "" & " order by Curr'"
        rsat = Comm.Execute
        If rsat.RecordCount <> 0 Then
            While Not rsat.EOF()
                Cmb.Items.Add(Trim(rsat.Fields("Curr").Value))
                rsat.MoveNext()
            End While
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
        If MPermit = "User" Then
            LoadSqlData("select sub_id , off_add2  from  Ap_office Where left(sub_id,2) = '" & Off_Id & "' And Substring(sub_id,4,2) <> '00'  Order by sub_id", RSC)
        Else
            LoadSqlData("select sub_id , off_add2  from  Ap_office Where sub_id <> '00-00' And Substring(sub_id,4,2) <> '00'  Order by sub_id", RSC)
        End If
        With RSC
            Do Until .EOF = True
                Off_Usr.Items.Add((.Fields("sub_id").Value) & " " & (.Fields("off_add2").Value))
                .MoveNext()
            Loop
        End With
        Off_Usr.SelectedIndex = 0
        'Off_Usr.SelectedIndex = 0
    End Sub
    Private Sub FmNsewJeneralJournal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If MuLng = "E" Then
            txtAmt_letter.Visible = False
            txtAmt_letter_E.Visible = True
            FG.set_ColHidden(3, True)
            FG.set_ColHidden(4, False)
        Else
            FG.set_ColHidden(3, False)
            FG.set_ColHidden(4, True)
            txtAmt_letter.Visible = True
            txtAmt_letter_E.Visible = False
        End If
        Call loadOffice_User()
        txtInvoice.BackColor = Color.White
        FG.BackColorFixed = Color.LightGray

        'LoadCurr()
        If MdCertifyAuto = 1 Then
            txtInvoice.ReadOnly = True
        Else
            txtInvoice.Text = ""
            txtInvoice.ReadOnly = False
        End If

        BtnSearch.Visible = False
        BtnMove.Visible = False

        FG.FormatString = "^ລ/ດ |< ເລກບັນຊີໜີ           |< ເລກບັນຊີມີ           |< ຊື່ບັນຊີ (ລາວ)                                            |< ຊື່ບັນຊີ (ອັງກິດ)                     |> ຈຳນວນເງິນຈົດໜີ້        |> ຈຳນວນເງິນຈົດມີ     |^ລະຫັດ|< ຕົ້ນທຶນພາສາ (ລາວ)   |< ຕົ້ນທຶນພາສາ (ອັງກິດ)             |< ສະກຸນເງິນເງິນ |> ອັດຕາແລກປ່ຽນ |> ມູນຄ່າໜີ້          |> ມູນຄ່າມີ            |> 1111    |> 22       "
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
        FG.Row = 1
        FG.Col = 1
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
            FG.Rows = FG.Rows + 1
        Else
            FG.Rows = 1
            FG.Rows = 2
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
        For J = 1 To FG.Rows - 1
            FG.Row = J
            If Trim(FG.get_TextMatrix(J, 3)) <> "" Then
                If ChCat_ID.Checked = True Then
                    FG.Col = 7
                    FG.CellBackColor = Color.LightCyan
                Else
                    FG.Col = 7
                    FG.CellBackColor = Color.White
                    FG.CellBackColor = Color.White
                    FG.CellForeColor = Color.Gray
                    FG.Col = 8
                    FG.CellForeColor = Color.Gray
                    FG.Col = 9
                    FG.CellForeColor = Color.Gray

                    FG.set_TextMatrix(FG.Row, 7, "0")
                    FG.set_TextMatrix(FG.Row, 8, "ບໍ່ເລືອກ")
                    FG.set_TextMatrix(FG.Row, 9, "No Selete")
                End If
            End If
        Next J
        Call loadColor()
        'FG.AllowUserResizing = VSFlex8U.AllowUserResizeSettings.flexResizeBoth
        FG.ExtendLastCol = True
        SetControlText(Me)
        CheckBox4.Text = "Only Customer/Supplier"
        If MuLng = "E" Then
            Label24.Text = "Ref No. :"
        Else
            Label24.Text = "ເອກະສານ :"
        End If
        Button7.Text = "Connect SerVer"
        Chk_Preview.Text = "To be continued"

        'FG.FormatString = "^ລ/ດ |< ເລກບັນຊີໜີ           |< ເລກບັນຊີມີ           |< ຊື່ບັນຊີ (ລາວ)                                            |< ຊື່ບັນຊີ (ອັງກິດ)                     |> ຈຳນວນເງິນຈົດໜີ້        |> ຈຳນວນເງິນຈົດມີ     |^ລະຫັດ|< ຕົ້ນທຶນພາສາ (ລາວ)   |< ຕົ້ນທຶນພາສາ (ອັງກິດ)             |< ສະກຸນເງິນເງິນ |> ອັດຕາແລກປ່ຽນ |> ມູນຄ່າໜີ້          |> ມູນຄ່າມີ            "
    End Sub

    Private Sub LoadBook()
        Dim rst As New ADODB.Recordset
        CmbBook.Items.Clear()
        Comm = New ADODB.Command
        Comm.ActiveConnection = CNN
        Comm.CommandText = "SELECT * FROM books WHERE bookid <> '" & "" & " '"
        rst = Comm.Execute
        If rst.RecordCount <> 0 Then
            While Not rst.EOF()
                CmbBook.Items.Add(Trim(rst.Fields("bookid").Value))
                rst.MoveNext()
            End While
        End If
        CmbBook.Text = "GL"
        LoadSqlData("SELECT * FROM books WHERE bookid = N'" & CmbBook.Text & "'", RSC)
        With RSC
            Do Until .EOF = True
                txtBookName.Text = Trim(.Fields("bookname").Value)
                .MoveNext()
            Loop
        End With
 

    End Sub

    Public Sub AddAcc2()

        BtnMove.Visible = False
        If FG.Col = 1 Then
            FG.set_TextMatrix(FG.Row, 6, "0.00")
            FG.set_TextMatrix(FG.Row, 7, 0)
            FG.set_TextMatrix(FG.Row, 10, Cmb.Text)
            FG.set_TextMatrix(FG.Row, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 12, Format(CDbl(CDbl(txtRate.Text) * FG.get_TextMatrix(FG.Row, 4)), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 13, "0.00")
            SumAmountDr()
            If ChDe.Checked = True Then
                If MuLng = "L" Then
                    FG.Col = 3
                End If
                If MuLng = "E" Then
                    FG.Col = 4
                End If
                If FG.get_TextMatrix(FG.Rows - 1, 3) <> "" Then
                    FG.Rows = FG.Rows + 1

                End If
                Exit Sub
            End If
            FG.Col = 5
            'Call loadColor()
        End If
        If FG.Col = 2 Then
            AccId = FG.get_TextMatrix(FG.Row, FG.Col)
            LoadText()
            FG.set_TextMatrix(FG.Row, 3, AccName)
            MDSearchAcccode = (FG.get_TextMatrix(FG.Row, 2))
            FG.set_TextMatrix(FG.Row, 5, "0.00")
            FG.set_TextMatrix(FG.Row, 7, 0)
            FG.set_TextMatrix(FG.Row, 10, Cmb.Text)
            FG.set_TextMatrix(FG.Row, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 13, Format(CDbl(CDbl(txtRate.Text) * FG.get_TextMatrix(FG.Row, 6)), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 12, "0.00")
            SumAmountDr()
            If ChDe.Checked = True Then
                If MuLng = "L" Then
                    FG.Col = 3
                End If
                If MuLng = "E" Then
                    FG.Col = 4
                End If
                If FG.get_TextMatrix(FG.Rows - 1, 3) <> "" Then
                    FG.Rows = FG.Rows + 1

                End If
                Exit Sub
            End If
            FG.Col = 5
            'Call loadColor()
        End If

        SumAmountDr()

    End Sub
    Public Sub AddAcc()

        BtnMove.Visible = False
        If FG.Col = 1 Then

            txtSumAmountDr.Text = CDbl(txtSumAmountDr.Text) - CDbl(FG.get_TextMatrix(FG.Row, 5))
            FG.set_TextMatrix(FG.Row, 5, Format(CDbl(CDbl(txtAmount.Text) - CDbl(txtSumAmountDr.Text)), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 6, "0.00")
            FG.set_TextMatrix(FG.Row, 7, 0)
            FG.set_TextMatrix(FG.Row, 10, Cmb.Text)
            FG.set_TextMatrix(FG.Row, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 12, Format(CDbl(CDbl(txtRate.Text) * FG.get_TextMatrix(FG.Row, 5)), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 13, "0.00")
            SumAmountDr()
            If ChDe.Checked = True Then
                If MuLng = "L" Then
                    FG.Col = 3
                End If
                If MuLng = "E" Then
                    FG.Col = 4
                End If
                If FG.get_TextMatrix(FG.Rows - 1, 3) <> "" Then
                    FG.Rows = FG.Rows + 1
                End If
                Exit Sub
            End If

            FG.Col = 5

            If FG.get_TextMatrix(FG.Rows - 1, 3) <> "" Then
                FG.Rows = FG.Rows + 1
                'Timer1.Enabled = True

            End If
            'Call loadColor()
        End If
        If FG.Col = 2 Then

            AccId = FG.get_TextMatrix(FG.Row, FG.Col)
            LoadText()
            FG.set_TextMatrix(FG.Row, 3, AccName)
            MDSearchAcccode = (FG.get_TextMatrix(FG.Row, 2))
            txtSumAmountCr.Text = CDbl(txtSumAmountCr.Text) - CDbl(FG.get_TextMatrix(FG.Row, 6))
            FG.set_TextMatrix(FG.Row, 6, Format(CDbl(CDbl(txtAmount.Text) - CDbl(txtSumAmountCr.Text)), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 5, "0.00")
            FG.set_TextMatrix(FG.Row, 7, 0)
            FG.set_TextMatrix(FG.Row, 10, Cmb.Text)
            FG.set_TextMatrix(FG.Row, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 13, Format(CDbl(CDbl(txtRate.Text) * FG.get_TextMatrix(FG.Row, 6)), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 12, "0.00")
            SumAmountDr()
            If ChDe.Checked = True Then
                If MuLng = "L" Then
                    FG.Col = 3
                End If
                If MuLng = "E" Then
                    FG.Col = 4
                End If
                If FG.get_TextMatrix(FG.Rows - 1, 3) <> "" Then
                    FG.Rows = FG.Rows + 1
                End If
                Exit Sub
            End If


            FG.Col = 6

            If FG.get_TextMatrix(FG.Rows - 1, 3) <> "" Then
                FG.Rows = FG.Rows + 1
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
        For i = 1 To FG.Rows - 1
            If i <> FG.Rows - 1 Then
                If FG.get_TextMatrix(i, 3) = "" Then
                    FG.RemoveItem()
                End If
            End If
            AmountDr = AmountDr + CDbl(FG.get_TextMatrix(i, 5))
            AmountCr = AmountCr + CDbl(FG.get_TextMatrix(i, 6))
            TotalAmountDr = TotalAmountDr + CDbl(FG.get_TextMatrix(i, 12))
            TotalAmountCr = TotalAmountCr + CDbl(FG.get_TextMatrix(i, 13))
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
        LoadSqlData("SELECT * FROM Acc_Code WHERE AC_CODE = N'" & AccId & "'", RSC)
        With RSC
            Do Until .EOF = True
                AccId = Trim(.Fields("AC_CODE").Value)
                AccName = Trim(.Fields("Name_L").Value)
                AccNamee = Trim(.Fields("Name_E").Value)
                .MoveNext()
            Loop
        End With


    End Sub
    Public Sub LoadDesc()


        LoadSqlData("SELECT * FROM Acc_Code WHERE AC_CODE = N'" & AccId & "'", RSC)
        With RSC
            Do Until .EOF = True
                AccId = Trim(.Fields("AC_CODE").Value)
                AccName = Trim(.Fields("Name_L").Value)
                AccNamee = Trim(.Fields("Name_E").Value)
                .MoveNext()
            Loop
        End With
    End Sub
    Private Sub FG_AfterScroll(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_AfterScrollEvent) Handles FG.AfterScroll

        BtnSearch.Visible = False
        BtnMove.Visible = False
    End Sub
    Dim ACCode As String
    Private Sub FG_KeyUpEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_KeyUpEvent) Handles FG.KeyUpEvent
        If e.keyCode = 13 Then
            If Cmb.Text = "USD" Then
                FG.set_TextMatrix(FG.Row, 14, Format(CDbl(FG.get_TextMatrix(FG.Row, 5)), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 15, Format(CDbl(FG.get_TextMatrix(FG.Row, 6)), "#,##0.00"))
            Else
                FG.set_TextMatrix(FG.Row, 14, Format(CDbl(FG.get_TextMatrix(FG.Row, 12)) / CDbl(txtRateUSD.Text), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 15, Format(CDbl(FG.get_TextMatrix(FG.Row, 13)) / CDbl(txtRateUSD.Text), "#,##0.00"))
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
            If CDbl(FG.get_TextMatrix(FG.Row, 14)) > CDbl(Remain) Then
            Else
                FG.set_TextMatrix(FG.Row, 14, Format(CDbl(Remain), "#,##0.00"))
            End If
            If CDbl(FG.get_TextMatrix(FG.Row, 15)) > CDbl(Remain) Then
            Else
                FG.set_TextMatrix(FG.Row, 15, Format(CDbl(Remain), "#,##0.00"))
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
    
        Dim RSC1 As New ADODB.Recordset
        Dim s As String = "SELECT sum(Amt_dr) as Amt_dr   , sum(Amt_cr) as Amt_cr    FROM gen_jn WHERE ac_code = '" & ACCode & "' and gen_jn.date_work BETWEEN '" & Format(dtActi.Value, "yyyy-MM-dd") & "' AND '" & Format(dtActi.Value, "yyyy-MM-dd") & "'   " & MULook2 & "  "
        LoadSqlData(s, RSC1)
        If RSC1.RecordCount <> 0 Then
            SumDr = Format(CDbl(Trim(RSC1.Fields("Amt_dr").Value)), "#,##0.00")
            SumCr = Format(CDbl(Trim(RSC1.Fields("Amt_cr").Value)), "#,##0.00")
        End If

        LoadSqlData("select  amount_dr , amount_cr from Open_jn where ac_code='" & ACCode & "'   and  year(Date_work)= '" & Format(CDate(dtActi.Value), "yyyy") & "'  " & MULook2 & "   ", RSC)
        Op = 0
        If RSC.RecordCount <> 0 Then
            Op = CDbl(Trim(RSC.Fields("amount_dr").Value)) - CDbl(Trim(RSC.Fields("amount_cr").Value))
        End If
        Dim dss As Date
        dss = DateAdd(DateInterval.Day, -1, dtActi.Value)
        Dim RSC2 As New ADODB.Recordset
        LoadSqlData("select SUM(amount_dr) AS amount_dr ,SUM(amount_cr) AS amount_cr from Gen_jn where ac_code=N'" & ACCode & "'  And gen_jn.date_work   BETWEEN '" & "1-1-" & Format(dtActi.Value, "yyyy") & "' AND '" & Format(dss, "yyyy-MM-dd") & "' " & MULook2 & " group by ac_code ", RSC2)
        If RSC2.RecordCount <> 0 Then
            Op = Op + CDbl(CDbl(Trim(RSC2.Fields("amount_dr").Value)) - CDbl(Trim(RSC2.Fields("amount_Cr").Value)))
        End If

        If Op >= 0 Then
            Open_jn = Format(CDbl(Op), "##,##0.00")
        Else
            Open_jn = "(" & Format(CDbl(Op * (-1)), "##,##0.00") & ")"
        End If

    End Sub
    Private Sub FG_MouseDownEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseDownEvent) Handles FG.MouseDownEvent
        If txtInvoice.BackColor = Color.Red Then
            txtInvoice.Focus()
            Exit Sub
        End If
        MouseDownEvent()
    End Sub

    Public Sub MouseDownEvent()
        Select Case MouseButtons
            Case Windows.Forms.MouseButtons.Right
                FG.EditCell()
            Case Windows.Forms.MouseButtons.Left
                If FG.Rows >= 3 Then
                    BtnMove.Visible = True

                    BtnMove.Top = CInt((FG.CellTop / 15) + FG.Top)

                    If FG.get_TextMatrix(FG.Row, FG.Col) <> "" Then
                        If FG.Col = 7 Then
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

                    Panel1.Top = CInt((FG.CellTop / 15) + FG.Top)
                    Panel1.Left = CInt(FG.Left + (FG.CellLeft / 15) + (FG.CellWidth / 250))
                    If FG.Row = FG.Rows - 1 Then
                        BtnMove.Visible = False
                    End If
                Else
                    BtnMove.Visible = False
                End If
                If FG.Col = 1 Then
                    If FG.get_TextMatrix(FG.Row, 2) <> "" Then
                        BtnSearch.Visible = False
                    Else
                        BtnSearch.Visible = True
                    End If
                End If
                If FG.Col = 2 Then
                    If FG.get_TextMatrix(FG.Row, 1) <> "" Then
                        BtnSearch.Visible = False
                    Else
                        BtnSearch.Visible = True
                    End If
                End If
                If FG.Col = 1 Then
                    BtnSearch.Left = CInt(FG.Left + (FG.CellLeft / 15) + (FG.CellWidth / 21.8))
                    BtnSearch.Top = CInt((FG.CellTop / 15) + FG.Top)


                End If
                If FG.Col = 2 Then
                    BtnSearch.Size = New System.Drawing.Point(34, 26)
                    BtnSearch.Left = CInt(FG.Left + (FG.CellLeft / 15) + (FG.CellWidth / 22.2))
                    BtnSearch.Top = CInt((FG.CellTop / 15) + FG.Top)
                End If

                If FG.Row = FG.Rows - 1 Then
                    BtnMove.Visible = False
                End If

        End Select
    End Sub

    Private Sub btnmove_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnMove.Click
        FG.RemoveItem()
        SumAmountDr()
        Panel1.Visible = False
        BtnMove.Visible = False
        BtnSearch.Visible = False

    End Sub


    'Sum
    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelChange
        If txtInvoice.BackColor = Color.Red Then
            txtInvoice.Focus()
            Exit Sub
        End If

        'If txxdes Then
        FG.BackColor = Color.White
        FG.BackColorAlternate = Color.White
        If FG.Col = 1 Then
            R = FG.Row()
            L = FG.Col
            If FG.get_TextMatrix(FG.Row, 2) <> "" Then
                FG.BackColorSel = Color.White
                FG.Editable = VSFlex8U.EditableSettings.flexEDNone
            Else
                FG.BackColorSel = Color.SkyBlue
                FG.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
                AccId = FG.get_TextMatrix(FG.Row, FG.Col)
            End If
        End If

        If FG.Col = 2 Then
            R = FG.Row()
            L = FG.Col
            If FG.get_TextMatrix(FG.Row, 1) <> "" Then
                FG.BackColorSel = Color.White
                FG.Editable = VSFlex8U.EditableSettings.flexEDNone
                'FG.BackColorSel = Color.White
            Else
                FG.BackColorSel = Color.SkyBlue
                FG.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
                AccId = FG.get_TextMatrix(FG.Row, FG.Col)
            End If
        End If



        If FG.Col = 3 Then
            'FG.BackColorSel = Color.White
            FG.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
            FG.BackColorSel = Color.SkyBlue
        End If

        If FG.Col = 4 Then
            'FG.BackColorSel = Color.White
            FG.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
            FG.BackColorSel = Color.SkyBlue
        End If


        If FG.Col = 5 Then
            If FG.get_TextMatrix(FG.Row, 2) <> "" Then
                FG.Editable = VSFlex8U.EditableSettings.flexEDNone
                FG.BackColorSel = Color.White
            Else
                FG.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
                FG.BackColorSel = Color.SkyBlue
            End If
        End If

        If FG.Col = 6 Then
            If FG.get_TextMatrix(FG.Row, 1) <> "" Then
                FG.BackColorSel = Color.White
                FG.Editable = VSFlex8U.EditableSettings.flexEDNone

            Else
                FG.BackColorSel = Color.SkyBlue
                FG.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
            End If
        End If

        If FG.Col = 7 Then
            FG.BackColorSel = Color.SkyBlue
            FG.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
        End If


        If ChCat_ID.Checked = False Then
            If FG.Col = 3 Or FG.Col = 4 Or FG.Col = 7 Or FG.Col = 8 Or FG.Col = 9 Or FG.Col = 10 Or FG.Col = 11 Or FG.Col = 12 Or FG.Col = 13 Then
                FG.BackColorSel = Color.White
                FG.Editable = VSFlex8U.EditableSettings.flexEDNone
            End If
        Else
            If FG.Col = 3 Or FG.Col = 4 Or FG.Col = 8 Or FG.Col = 9 Or FG.Col = 10 Or FG.Col = 11 Or FG.Col = 12 Or FG.Col = 13 Then
                FG.BackColorSel = Color.White
                FG.Editable = VSFlex8U.EditableSettings.flexEDNone
            End If
        End If
        If FG.Col > 2 Then
            If FG.get_TextMatrix(FG.Row, 1) = "" Then
                If FG.get_TextMatrix(FG.Row, 2) = "" Then
                    FG.Editable = VSFlex8U.EditableSettings.flexEDNone
                    FG.BackColorSel = Color.White
                End If
            End If
        End If
        'If CheckBox3.Checked = True Then
        '    CheckBox3.
        'End If
        If FG.Col = 3 Or FG.Col = 4 Or FG.Col = 5 Or FG.Col = 6 Or FG.Col = 7 Or FG.Col = 8 Or FG.Col = 9 Or FG.Col = 10 Or FG.Col = 11 Or FG.Col = 12 Or FG.Col = 13 Then
            BtnSearch.Visible = False
        End If

        If FG.get_TextMatrix(FG.Row, 5) = "" Then
            Panel1.Visible = False
        End If
        If FG.get_TextMatrix(FG.Row, 7) = "0" Then
            FG.set_TextMatrix(FG.Row, 8, "ບໍ່ເລືອກ")
            FG.set_TextMatrix(FG.Row, 9, "No Selete")
        ElseIf FG.get_TextMatrix(FG.Row, 7) = "1" Then
            FG.set_TextMatrix(FG.Row, 8, "ຮັບໃຊ້ການພະລິດ")
            FG.set_TextMatrix(FG.Row, 9, "Use build")

        ElseIf FG.get_TextMatrix(FG.Row, 7) = "2" Then
            FG.set_TextMatrix(FG.Row, 8, "ຮັບໃຊ້ການຈຳໜ່າຍ")
            FG.set_TextMatrix(FG.Row, 9, "Use Sell")
        ElseIf FG.get_TextMatrix(FG.Row, 7) = "3" Then
            FG.set_TextMatrix(FG.Row, 8, "ຮັບໃຊ້ບໍລິຫານ")
            FG.set_TextMatrix(FG.Row, 9, "Use manage ")

        ElseIf FG.get_TextMatrix(FG.Row, 7) = "4" Then
            FG.set_TextMatrix(FG.Row, 8, "ຕົ້ນທຶນຂາຍ/ຕົ້ນທຶນບໍລິຫານ")
            FG.set_TextMatrix(FG.Row, 9, "Sell capital/manage capital ")
        End If

    End Sub

    Private Sub CmbBook_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles CmbBook.KeyPress
        If e.KeyChar = Chr(13) Then
            txtAmount.Focus()
        End If
    End Sub

    Private Sub CmbBook_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbBook.SelectedIndexChanged
        txtInvoice.Text = CmbBook.Text & MdCertifyId
        LoadSqlData("SELECT * FROM books WHERE bookid = N'" & CmbBook.Text & "'", RSC)
        With RSC
            Do Until .EOF = True
                txtBookName.Text = Trim(.Fields("bookname").Value)
                .MoveNext()
            Loop
        End With
        If txtInvoice.Enabled = True Then
            AutoNumber()
        End If
    End Sub

    Private Sub txtAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAmount.KeyPress
        If e.KeyChar = Chr(13) Then
            txtTotal_Amt_LAK.Text = CDbl(txtRate.Text) * CDbl(txtAmount.Text)
            Call FormatText()
            If FG.get_TextMatrix(1, 1) = "" And FG.get_TextMatrix(1, 2) = "" Then
                FG.Row = 1
                FG.Col = 1
            End If
            If FG.get_TextMatrix(1, 1) <> "" Then
                FG.Row = 1
                FG.Col = 1
            End If
            If FG.get_TextMatrix(1, 2) <> "" Then
                FG.Row = 1
                FG.Col = 2
            End If
            FG.Focus()
        End If
    End Sub

    Private Sub txtAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAmount.LostFocus
        Call FormatText()
    End Sub

    Private Sub txtAmount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmount.TextChanged
        If IsNumeric(txtAmount.Text) = False Then txtAmount.Text = "0" : Exit Sub
        If txtAmount.Text = "" Then txtAmount.Text = "0" : Exit Sub
        If txtAmount.Text = 0 Then
            txtAmt_letter.Text = ""
        End If
        If IsNumeric(txtAmount.Text) = False Then txtAmount.Clear() : Exit Sub
        txtAmt_letter.Text = Letter_amt(txtAmount)
        txtAmt_letter_E.Text = LetterEng_amt(txtAmount)

    End Sub
    Public Function Letter_amt(ByVal Txt As TextBox, Optional ByVal CurrKIP As Boolean = False) As String

        If Val(txtAmount.Text) <> 0 Then

            RateType = ": " & Cmb.Text

            Letter_amt = CMoney(Format(CDbl(Txt.Text), "##0.00")) & RateType
        Else
            Letter_amt = ""
        End If

    End Function
    Public Function LetterEng_amt(ByVal Txt As TextBox, Optional ByVal CurrKIP As Boolean = False) As String
        If Val(txtAmount.Text) <> 0 Then
            LetterEng_amt = CMoneyEng(Format(CDbl(Txt.Text), "##0.00")) & ": " & Cmb.Text
        Else
            LetterEng_amt = ""
        End If
    End Function
    Private Sub txtsearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        FG.Focus()
        R = FG.Row
        L = FG.Col
    End Sub



    'Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    '    Static cntd As Integer = 0
    '    Dim a As String
    '    If cntd > 0 Then
    '        a = cntd.ToString
    '    Else
    '        Timer1.Enabled = False
    '        Foucus()
    '    End If
    'End Sub

    Private Sub LoadTableId()
        Bee.set_TextMatrix(1, 2, "ບໍ່ເລືອກ")
        Bee.set_TextMatrix(2, 2, "ຮັບໃຊ້ໃນການພະລິດ")
        Bee.set_TextMatrix(3, 2, "ຮັບໃຊ້ໃນການຈຳໜ່າຍ")
        Bee.set_TextMatrix(4, 2, "ຮັບໃຊ້ບໍລິຫານ")
        Bee.set_TextMatrix(5, 2, "ຕົ້ນທື່ນຂາຍ/ຕົ້ນທຶນບໍລິຫານ")

        Bee.set_TextMatrix(1, 3, "No Selete")
        Bee.set_TextMatrix(2, 3, "Use build")
        Bee.set_TextMatrix(3, 3, "Use Sell")
        Bee.set_TextMatrix(4, 3, "Use manage")
        Bee.set_TextMatrix(5, 3, "Sell capital/manage capital ")

        Bee.set_TextMatrix(1, 4, " 0.   ບໍ່ເລືອກ")
        Bee.set_TextMatrix(2, 4, " 1.   ຮັບໃຊ້ໃນການພະລິດ")
        Bee.set_TextMatrix(3, 4, " 2.   ຮັບໃຊ້ໃນການຈຳໜ່າຍ")
        Bee.set_TextMatrix(4, 4, " 3.   ຮັບໃຊ້ບໍລິຫານ")
        Bee.set_TextMatrix(5, 4, " 4.   ຕົ້ນທື່ນຂາຍ/ຕົ້ນທຶນບໍລິຫານ")
    End Sub

    Private Sub Bee_AfterEdit(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_AfterEditEvent) Handles Bee.AfterEdit

    End Sub

    Private Sub Bee_DblClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bee.DblClick
        FG.set_TextMatrix(FG.Row, 7, Bee.get_TextMatrix(Bee.Row, 1))
        FG.set_TextMatrix(FG.Row, 8, Bee.get_TextMatrix(Bee.Row, 2))
        FG.set_TextMatrix(FG.Row, 9, Bee.get_TextMatrix(Bee.Row, 3))
        Panel1.Visible = False
        loadColor()
        'If FG.get_TextMatrix(FG.Row + 1, 6) <> "" Then
        '    If FG.get_TextMatrix(FG.Row + 1, 1) = "" Then
        '        FG.Col = 2
        '        FG.Row = FG.Row + 1
        '    Else
        '        FG.Col = 1
        '        FG.Row = FG.Row + 1
        '    End If
        'End If
        'If FG.get_TextMatrix(FG.Row + 1, 3) = "" Then
        '    If FG.get_TextMatrix(FG.Row, 1) = "" Then
        '        FG.Col = 1
        '        FG.Row = FG.Row + 1
        '    Else
        '        FG.Col = 2
        '        FG.Row = FG.Row + 1
        '    End If
        'End If
        FG.Focus()
    End Sub




    Private Sub Bee_KeyUpEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_KeyUpEvent) Handles Bee.KeyUpEvent
        If e.keyCode = 13 Then
            FG.set_TextMatrix(FG.Row, 7, Bee.get_TextMatrix(Bee.Row, 1))
            FG.set_TextMatrix(FG.Row, 8, Bee.get_TextMatrix(Bee.Row, 2))
            FG.set_TextMatrix(FG.Row, 9, Bee.get_TextMatrix(Bee.Row, 3))
            Panel1.Visible = False
            loadColor()
            'If FG.get_TextMatrix(FG.Row + 1, 6) <> "" Then
            '    If FG.get_TextMatrix(FG.Row + 1, 1) = "" Then
            '        FG.Col = 2
            '        FG.Row = FG.Row + 1
            '    Else
            '        FG.Col = 1
            '        FG.Row = FG.Row + 1
            '    End If
            'End If
            'If FG.get_TextMatrix(FG.Row + 1, 3) = "" Then
            '    If FG.get_TextMatrix(FG.Row, 1) = "" Then
            '        FG.Col = 1
            '        FG.Row = FG.Row + 1
            '    Else
            '        FG.Col = 2
            '        FG.Row = FG.Row + 1
            '    End If
            'End If
            FG.Focus()
        End If
    End Sub


    Public Sub loadRate()
        txtUserId.Text = MDoff_id
        txtTotal_Amt_LAK.Text = CDbl(txtAmount.Text) * CDbl(txtRate.Text)
        'FormatText()
    End Sub
    Public Sub FormatText()
        txtRate.Text = Format(CDbl(txtRate.Text), "#,##0.00")
        txtTotal_Amt_LAK.Text = Format(CDbl(txtTotal_Amt_LAK.Text), "#,##0.00")
        txtAmount.Text = Format(CDbl(txtAmount.Text), "#,##0.00")
        txtAmt_letter.Text = Letter_amt(txtAmount)
    End Sub








    Public Sub loadColor()
        FgR = FG.Row
        FgC = FG.Col
        Dim rmRS As New ADODB.Recordset
        Dim J As Integer
        FG.Redraw = False



        FG.Col = 1
        FG.Row = FG.Rows - 1
        FG.CellBackColor = Color.LightCyan
        FG.Col = 2
        FG.CellBackColor = Color.LightCyan

        For J = 1 To FG.Rows - 1
            FG.Row = J
            If Trim(FG.get_TextMatrix(J, 1)) <> "" Then
                FG.Col = 1
                FG.CellBackColor = Color.LightCyan
                FG.CellFontBold = True
                'FG.Col = 3
                'FG.CellBackColor = Color.White
                'FG.Col = 4
                'FG.CellBackColor = Color.White
                FG.Col = 5
                FG.CellBackColor = Color.LightCyan
                FG.CellFontBold = True
                'FG.Col = 6
                'FG.CellForeColor = Color.Gray
                'FG.Col = 7
                'FG.CellBackColor = Color.LightCyan
                'FG.Col = 12
                'FG.CellBackColor = Color.LightCyan
                FG.set_TextMatrix(J, 0, J)




            Else
                If Trim(FG.get_TextMatrix(J, 1)) = "" Then
                    FG.Col = 1
                    FG.CellBackColor = Color.White

                End If

                If Trim(FG.get_TextMatrix(J, 2)) <> "" Then
                    FG.Col = 2
                    FG.CellBackColor = Color.LightCyan
                    FG.CellFontBold = True
                    'FG.Col = 3
                    'FG.CellBackColor = Color.White
                    'FG.Col = 4
                    'FG.CellBackColor = Color.White
                    FG.Col = 6
                    FG.CellFontBold = True
                    FG.CellBackColor = Color.LightCyan
                    'FG.Col = 5
                    'FG.CellForeColor = Color.Gray
                    'FG.Col = 7
                    'FG.CellBackColor = Color.LightCyan
                    'FG.Col = 13
                    'FG.CellBackColor = Color.LightCyan

                    FG.set_TextMatrix(J, 0, J)

                End If
                If Trim(FG.get_TextMatrix(J, 2)) = "" Then
                    FG.Col = 2
                    FG.CellBackColor = Color.White

                End If
            End If

            '    If Trim(FG.get_TextMatrix(J, 7)) = "0" Then
            '        FG.Col = 7
            '        FG.CellForeColor = Color.Gray
            '        FG.Col = 8
            '        FG.CellForeColor = Color.Gray
            '        FG.Col = 9
            '        FG.CellForeColor = Color.Gray
            '    Else
            '        FG.Col = 7
            '        FG.CellForeColor = Color.Black
            '        FG.Col = 8
            '        FG.CellForeColor = Color.Black
            '        FG.Col = 9
            '        FG.CellForeColor = Color.Black
            '    End If
            '    If ChCat_ID.Checked = False Then

            '        FG.Col = 7

            '        FG.CellBackColor = Color.White
            '        FG.CellForeColor = Color.Gray
            '        FG.Col = 8
            '        FG.CellForeColor = Color.Gray
            '        FG.Col = 9
            '        FG.CellForeColor = Color.Gray

            '    End If
        Next J


        FG.Row = FgR
        FG.Col = FgC
        FG.Redraw = True
    End Sub

    Private Sub Cmb_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cmb.Click
        'LoadCurr()
    End Sub

    Private Sub Cmb_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmb.SelectedIndexChanged

        Dim rs As New ADODB.Recordset
        Call LoadSqlData("Select * From Curr_For_Rate Where   Curr =N'" & Trim(Cmb.Text) & "'", rs)
        If rs.RecordCount > 0 Then
            txtcurr_name2.Text = Trim(rs("Curr_name").Value.ToString)
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

        'For i = 1 To FG.Rows - 2
        '    'If FG.Row > 1 Then
        '    FG.set_TextMatrix(i, 10, Cmb.Text)
        '    FG.set_TextMatrix(i, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
        '    FG.set_TextMatrix(i, 12, Format(CDbl(CDbl(txtRate.Text) * FG.get_TextMatrix(i, 5)), "#,##0.00"))
        '    FG.set_TextMatrix(i, 13, Format(CDbl(CDbl(txtRate.Text) * FG.get_TextMatrix(i, 6)), "#,##0.00"))
        '    'End If

        'Next
    End Sub
    Public Sub LoadSetRate()
        LoadSqlData("select * from Ap_RateSeting where Curr='" & Cmb.Text & "'", RSC)
        With RSC
            Do Until .EOF = True
                txtRate.Text = Trim(.Fields("Rate").Value)
                curr_Last.Text = Trim(.Fields("curr_Last").Value)
                .MoveNext()
            Loop
        End With
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
        For i = 1 To FG.Rows - 1
            If FG.get_TextMatrix(i, 1) & FG.get_TextMatrix(i, 2) <> "" Then
                'MsgBox(Len(FG.get_TextMatrix(i, 1) & FG.get_TextMatrix(i, 2)))
                If Len(FG.get_TextMatrix(i, 1) & FG.get_TextMatrix(i, 2)) <> 7 Then
                    'LngId = "6005" : MsgRpt()
                    'Exit Sub
                End If
            End If
        Next i

        Dim k As Integer
        For k = 1 To FG.Rows - 2
            If FG.get_TextMatrix(k, 1) <> "" Then
                If FG.get_TextMatrix(k, 5) = 0 Then
                    MsgBox("ກະລຸນນາໃສ່ມູນຄ່າກ່ອນ")
                    Exit Sub
                End If
            End If
            If FG.get_TextMatrix(k, 2) <> "" Then
                If FG.get_TextMatrix(k, 6) = 0 Then
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
        Call LoadSqlData("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book =N'" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "'  And  ReferNO = N'" & TxtReferno.Text & "'And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And  Month(date_work)=" & Format(CDate(dtActi.Value), "MM") & "  And  day(date_work)=" & Format(CDate(dtActi.Value), "dd") & " " & _
             "  and LEFT(company,2)=N'" & Off_Id & "' Order by  Right(certify,3) DESC ", RSC)

        If RSC.RecordCount > 0 Then
            If txtInvoice.Enabled = True Then
                MsgBox("ເລກລະຫັດ : " & Trim(txtInvoice.Text) & " ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
                txtInvoice.BackColor = Color.Red
                txtInvoice.Focus()
                If RSC.State = ConnectionState.Open Then RSC.Close()
                Exit Sub
            End If

        End If
        'Dim J As Integer
        'For J = 1 To FG.Rows - 1
        '    '===============
        '    FG.set_TextMatrix(J, 10, Cmb.Text)
        '    FG.set_TextMatrix(J, 11, Format(CDbl(txtRate.Text), "#,##0.00"))


        '    If Cmb.Text = "USD" Then
        '        FG.set_TextMatrix(J, 14, Format(CDbl(FG.get_TextMatrix(J, 5)), "#,##0.00"))
        '        FG.set_TextMatrix(J, 15, Format(CDbl(FG.get_TextMatrix(J, 6)), "#,##0.00"))
        '    Else
        '        FG.set_TextMatrix(J, 14, Format(CDbl(FG.get_TextMatrix(J, 12)) / CDbl(txtRateUSD.Text), "#,##0.00"))
        '        FG.set_TextMatrix(J, 15, Format(CDbl(FG.get_TextMatrix(J, 13)) / CDbl(txtRateUSD.Text), "#,##0.00"))
        '        'FG.set_TextMatrix(FG.Row, 14, Format(CDbl(FG.get_TextMatrix(FG.Row, 12)) / CDbl(txtRateUSD.Text), "#,##0.00"))
        '        'FG.set_TextMatrix(FG.Row, 15, Format(CDbl(FG.get_TextMatrix(FG.Row, 13)) / CDbl(txtRateUSD.Text), "#,##0.00"))
        '    End If

        'Next J

        If FG.get_TextMatrix(1, 6) = "" Then MsgBox("ກະລຸນນາລົງບັນຊີເງິນກອ່ນ!", MsgBoxStyle.OkOnly) : Exit Sub
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
            Call LoadSqlData("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book =N'" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "'And  ReferNO = N'" & TxtReferno.Text & "' And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And  month(date_work)=" & Format(CDate(dtActi.Value), "MM") & " And  day(date_work)=" & Format(CDate(dtActi.Value), "dd") & "  Order by  Right(certify,3) DESC ", RSC)

            If RSC.RecordCount > 0 Then
                MsgBox("ເລກລະຫັດ : " & Trim(txtInvoice.Text) & " ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
                txtInvoice.BackColor = Color.Red
                txtInvoice.Focus()
                If RSC.State = ConnectionState.Open Then RSC.Close()
                Exit Sub
            End If

            'MuSubOff = Mid(Off_Usr.Text, 1, 5)

            Savedata()
            'MuSubOff = MuSubOff2
        Else
            'CNN.Execute("DELETE FROM gen_jn WHERE book =N'" & FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 16) & "' And certify  = '" & txtInvoice.Text & "'   And  year(date_work)='" & Format(CDate(FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 1)), "yyyy") & "' ")
            CNN.Execute("DELETE FROM gen_jn WHERE book =N'" & FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 16) & "' And certify  =N'" & txtInvoice.Text & "'  And  ReferNO = N'" & TxtReferno.Text & "'  And   date_work='" & Format(CDate(FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 1)), "yyyy-MM-dd") & "' ")

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
        FG.Rows = 1
        FG.Rows = 2
        If CheckBox1.Checked = True Then
            Call LoadReport()
        End If
        If CheckBox2.Checked = True Then
            Close()
        End If
        'End If


    End Sub

    Private Sub BtnAddNew2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew2.Click
        FG.Rows = 1
        FG.Rows = 2
        AutoNumber()
        txtInvoice.Enabled = True
        CmbBook.Enabled = True
        Panel1.Visible = False
        BtnMove.Visible = False
        BtnSearch.Visible = False
    End Sub

    Private Sub Bee_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Bee.SelChange

    End Sub

    Private Sub FG_StartEdit(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_StartEditEvent) Handles FG.StartEdit
        If txtInvoice.BackColor = Color.Red Then
            txtInvoice.Focus()
            Exit Sub
        End If
        FG.CellBackColor = Color.DeepSkyBlue
        FG.CellForeColor = Color.Black
        BtnMove.Visible = False
        BtnSearch.Visible = False



    End Sub

    Private Sub Cat()
        If FG.Col = 7 Then
            If FG.get_TextMatrix(FG.Row, 7) = "0" Then
                FG.set_TextMatrix(FG.Row, 8, "ບໍ່ເລືອກ")
                FG.set_TextMatrix(FG.Row, 9, "No Selete")
            ElseIf FG.get_TextMatrix(FG.Row, 7) = "1" Then
                FG.set_TextMatrix(FG.Row, 8, "ຮັບໃຊ້ການພະລິດ")
                FG.set_TextMatrix(FG.Row, 9, "Use build")

            ElseIf FG.get_TextMatrix(FG.Row, 7) = "2" Then
                FG.set_TextMatrix(FG.Row, 8, "ຮັບໃຊ້ການຈຳໜ່າຍ")
                FG.set_TextMatrix(FG.Row, 9, "Use Sell")
            ElseIf FG.get_TextMatrix(FG.Row, 7) = "3" Then
                FG.set_TextMatrix(FG.Row, 8, "ຮັບໃຊ້ບໍລິຫານ")
                FG.set_TextMatrix(FG.Row, 9, "Use manage ")

            ElseIf FG.get_TextMatrix(FG.Row, 7) = "4" Then
                FG.set_TextMatrix(FG.Row, 8, "ຕົ້ນທຶນຂາຍ/ຕົ້ນທຶນບໍລິຫານ")
                FG.set_TextMatrix(FG.Row, 9, "Sell capital/manage capital ")
            ElseIf FG.get_TextMatrix(FG.Row, 7) > 4 Then
                MessageBox.Show("ລະຫັດບໍ່ຖືກຕ້ອງ ລະຫັດມີພຽງເລກ 0 ຫາເລກ 4 ເທົ່ານັ້ນ")
                FG.set_TextMatrix(FG.Row, 8, "ບໍ່ລເລືອກ")
                FG.set_TextMatrix(FG.Row, 9, "No Selete")
                FG.set_TextMatrix(FG.Row, 7, "0")
                Exit Sub
            ElseIf FG.get_TextMatrix(FG.Row, 7) < 0 Then
                MessageBox.Show("ລະຫັດບໍ່ຖືກຕ້ອງ ລະຫັດມີພຽງເລກ 0 ຫາເລກ 4 ເທົ່ານັ້ນ")
                FG.set_TextMatrix(FG.Row, 8, "ບໍ່ລເລືອກ")
                FG.set_TextMatrix(FG.Row, 9, "No Selete")
                FG.set_TextMatrix(FG.Row, 7, "0")
                Exit Sub
            End If
        End If


    End Sub
    Private Sub LoadSQL()
        sql = ""
        'And  year(date_work)='" & Format(CDate(FG.get_TextMatrix(FG.Row, 2)), "yyyy") & "'
        'sql = " AND certify  = " & txtInvoice.Text & "  And  year(date_work)='" & Format(CDate(FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 1)), "yyyy") & "' "
        sql = " AND certify  =N" & txtInvoice.Text

    End Sub
    Public Sub LoadListFG()
        CNN.Execute("UPDATE gen_jn set Cust_Supp=0 where   Cust_Supp is null")
        FG.Rows = 1
        With RSC
            Dim PK As String = "select *  from gen_jn WHERE book =N'" & FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 16) & "' And certify  =N'" & txtInvoice.Text & "'   And ReferNO  =N'" & FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 3) & "'  And  year(date_work)='" & Format(CDate(FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 1)), "yyyy") & "' order by cnt"
            Call LoadSqlData(PK, RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    If CDbl(.Fields("AG").Value) = 1 Then
                        FG.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("code_dr").Value.ToString)) & _
                    "" & vbTab & Trim(CStr(.Fields("code_cr").Value.ToString)) & _
                     "" & vbTab & Trim((.Fields("descrip").Value.ToString)) & _
                      "" & vbTab & Trim((.Fields("descripe").Value.ToString)) & _
                         "" & vbTab & Format(CDbl(.Fields("amt_dr").Value), "##,##0.00") & _
                           "" & vbTab & Format(CDbl(.Fields("amt_cr").Value), "##,##0.00") & _
                            "" & vbTab & Trim(CStr(.Fields("Cat_ID").Value.ToString)))

                    Else
                        FG.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("code_dr").Value.ToString)) & _
                          "" & vbTab & Trim(CStr(.Fields("code_cr").Value.ToString)) & _
                           "" & vbTab & Trim((.Fields("descrip").Value.ToString)) & _
                            "" & vbTab & Trim((.Fields("descripe").Value.ToString)) & _
                               "" & vbTab & Format(CDbl(.Fields("amount_dr").Value), "##,##0.00") & _
                                 "" & vbTab & Format(CDbl(.Fields("amount_cr").Value), "##,##0.00") & _
                                  "" & vbTab & Trim(CStr(.Fields("Cat_ID").Value.ToString)) & _
                                   "" & vbTab & Trim(CStr(.Fields("Cat_ID").Value.ToString)) & _
                                    "" & vbTab & Trim(CStr(.Fields("Cat_ID").Value.ToString)) & _
                                      "" & vbTab & Trim((.Fields("Curr").Value.ToString)) & _
                                                       "" & vbTab & Format(CDbl(.Fields("Rate").Value), "##,##0.00") & _
                         "" & vbTab & Format(CDbl(.Fields("amt_Dr").Value), "##,##0.00") & _
                          "" & vbTab & Format(CDbl(.Fields("amt_cr").Value), "##,##0.00") & _
                                           "" & vbTab & Format(CDbl(.Fields("amt_USD_Dr").Value), "##,##0.00") & _
                                    "" & vbTab & Format(CDbl(.Fields("amt_USD_Cr").Value), "##,##0.00"))
                    End If
                    .MoveNext()

                End While
            Else
                FG.Rows = 16
            End If
            '==
            'Call LoadSqlData("select top 1 *  from gen_jn WHERE book ='" & FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 16) & "' And certify  = '" & txtInvoice.Text & "'   And  year(date_work)='" & Format(CDate(FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 1)), "yyyy") & "' order by cnt", RSC)
            Dim RSK As New ADODB.Recordset
            With RSK 
                Dim PP As String = "select top 1 *  from gen_jn WHERE book =N'" & FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 16) & "' And certify  =N'" & txtInvoice.Text & "'  And ReferNO  =N'" & FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 3) & "'    And   date_work='" & Format(CDate(FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 1)), "yyyy-MM-dd") & "' order by cnt"
                Call LoadSqlData(PP, RSK)
                If .RecordCount > 0 Then
                    While Not .EOF


                        Book = Trim(.Fields("book").Value)
                        dtActi.Text = Format(CDate(Trim(.Fields("date_work").Value)), "dd/MM/yyyy")
                        txtAmount.Text = Format(CDbl(Trim(.Fields("amount").Value)), "#,##0.00")
                        txtDesc.Text = Trim(.Fields("descrip").Value.ToString)
                        txtDescE.Text = Trim(.Fields("descripe").Value.ToString)
                        Rate1 = Format(CDbl(Trim(.Fields("rate").Value)), "#,##0.00")
                        Cmb.Text = Trim(.Fields("curr").Value)
                        CmbBook.Text = Book
                        txtInvoice.Text = MDInvoiceNo
                        TxtCustID.Text = Trim(.Fields("CustID").Value.ToString)
                        TxtSuppID.Text = Trim(.Fields("SuppID").Value.ToString)
                        txtRateUSD.Text = Format(CDbl(Trim(.Fields("rate_USD").Value)), "#,##0.00")
                        TxtReferno.Text = Trim(.Fields("Referno").Value.ToString)


                        MCS = Trim(.Fields("Cust_Supp").Value.ToString)
                        .MoveNext()
                    End While
                    If MCS = 1 Then
                        CheckBox4.Checked = True
                        If TxtCustID.Text <> "" Then
                            Dim RSCUST As New ADODB.Recordset
                            Call LoadSqlData("select *  from Customer WHERE 1=1 and  Code=N'" & TxtCustID.Text & "'  ", RSCUST)
                            If RSCUST.RecordCount > 0 Then
                                CmbCust.Text = Trim(RSCUST.Fields("Name").Value.ToString)
                            Else
                                CmbCust.Text = ""
                            End If
                            If TxtCustID.Text <> "" Then
                                RadioButton1.Checked = True
                            Else
                                RadioButton1.Checked = False
                            End If
                        End If

                        If TxtSuppID.Text <> "" Then
                            Dim RSCSUPPT As New ADODB.Recordset
                            Call LoadSqlData("select *  from Supplier WHERE 1=1 and  Code=N'" & TxtSuppID.Text & "'  ", RSCSUPPT)
                            If RSCSUPPT.RecordCount > 0 Then
                                CmbSupp.Text = Trim(RSCSUPPT.Fields("Name").Value.ToString)
                            Else
                                CmbSupp.Text = ""
                            End If
                            If TxtSuppID.Text <> "" Then
                                RadioButton2.Checked = True
                            Else
                                RadioButton2.Checked = False
                            End If
                        End If

                    Else
                        CheckBox4.Checked = False
                        RadioButton1.Checked = False
                        RadioButton2.Checked = False
                    End If
                    If AG = 1 Then
                        CheckBox3.Checked = True
                    Else
                        CheckBox3.Checked = False
                    End If
                End If
            End With

            txtRate.Text = Rate1

            Dim i As Integer
            For i = 1 To FG.Rows - 1
                '===============
                'FG.set_TextMatrix(i, 10, Cmb.Text)
                'FG.set_TextMatrix(i, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
                'If CDbl(.Fields("amount_dr").Value) = 1 Then

                'End If
                FG.set_TextMatrix(i, 12, Format(CDbl(CDbl(FG.get_TextMatrix(i, 11)) * FG.get_TextMatrix(i, 5)), "#,##0.00"))
                FG.set_TextMatrix(i, 13, Format(CDbl(CDbl(FG.get_TextMatrix(i, 11)) * FG.get_TextMatrix(i, 6)), "#,##0.00"))
                FG.set_TextMatrix(i, 14, Format(CDbl(FG.get_TextMatrix(i, 12)) / CDbl(txtRate.Text), "#,##0.00"))
                FG.set_TextMatrix(i, 15, Format(CDbl(FG.get_TextMatrix(i, 13)) / CDbl(txtRate.Text), "#,##0.00"))
                '12345
                If FG.get_TextMatrix(FG.Row, 7) = "0" Then
                    FG.set_TextMatrix(i, 8, "ບໍ່ເລືອກ")
                    FG.set_TextMatrix(i, 8, "No Selete")

                ElseIf FG.get_TextMatrix(FG.Row, 7) = "1" Then
                    FG.set_TextMatrix(i, 8, "ຮັບໃຊ້ການພະລິດ")
                    FG.set_TextMatrix(i, 8, "Use build")
                ElseIf FG.get_TextMatrix(FG.Row, 7) = "2" Then
                    FG.set_TextMatrix(i, 8, "ຮັບໃຊ້ການຈຳໜ່າຍ")
                    FG.set_TextMatrix(i, 8, "Use Sell")
                ElseIf FG.get_TextMatrix(FG.Row, 7) = "3" Then
                    FG.set_TextMatrix(i, 8, "ຮັບໃຊ້ບໍລິຫານ")
                    FG.set_TextMatrix(i, 8, "Use manage")

                ElseIf FG.get_TextMatrix(FG.Row, 7) = "4" Then
                    FG.set_TextMatrix(i, 8, "ຕົ້ນທຶນຂາຍ/ຕົ້ນທຶນບໍລິຫານ")
                    FG.set_TextMatrix(i, 8, "capital/manage capital")
                End If
            Next i
        End With

        For i = 1 To FG.Rows - 1
            Call LoadSqlData("select *  from gen_jn WHERE Ac_Code=N'" & FG.get_TextMatrix(i, 1) & FG.get_TextMatrix(i, 2) & "' and book =N'" & FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 16) & "' And certify  =N'" & txtInvoice.Text & "'  And ReferNO  =N'" & FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 3) & "'    And  year(date_work)='" & Format(CDate(FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 1)), "yyyy") & "' order by cnt", RSC)
            If RSC.RecordCount > 0 Then
                FG.set_TextMatrix(i, 3, Trim(RSC.Fields("ac_Name").Value.ToString))
                FG.set_TextMatrix(i, 4, Trim(RSC.Fields("ac_Namee").Value.ToString))

            End If
        Next i

        'For i = 1 To FG.Rows - 1
        '    Call LoadSqlData("select *  from Acc_Code WHERE Ac_Code='" & FG.get_TextMatrix(i, 1) & FG.get_TextMatrix(i, 2) & "'order by Ac_Code", RSC)
        '    If RSC.RecordCount > 0 Then
        '        FG.set_TextMatrix(i, 3, Trim(RSC.Fields("Name_L").Value))
        '        FG.set_TextMatrix(i, 4, Trim(RSC.Fields("Name_E").Value))

        '    End If
        'Next i

        SumAmountDr()
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ShowList()
    End Sub
    Private Sub ShowList()
        FG.set_ColHidden(7, True)
        FG.set_ColHidden(8, True)
        If Button3.Text = "Show All" Then


            FG.set_ColHidden(4, False)
            FG.set_ColHidden(9, False)
            'FG.set_ColHidden(10, False)
            'FG.set_ColHidden(11, False)
            'FG.set_ColHidden(12, False)
            'FG.set_ColHidden(13, False)
            'FG.set_ColHidden(14, True)
            'FG.set_ColHidden(15, True)
            Button3.Text = "Show GLN"
            Exit Sub
        End If
        If Button3.Text = "Show GLN" Then
            If MuLng = "E" Then
                FG.set_ColHidden(3, True)
                FG.set_ColHidden(4, False)
            Else
                FG.set_ColHidden(3, False)
                FG.set_ColHidden(4, True)
            End If
            'FG.set_ColHidden(4, True)
            FG.set_ColHidden(9, True)
            'FG.set_ColHidden(10, True)
            'FG.set_ColHidden(11, True)
            'FG.set_ColHidden(12, True)
            'FG.set_ColHidden(13, True)
            'FG.set_ColHidden(14, True)
            'FG.set_ColHidden(15, True)
            Button3.Text = "Show All"
            Exit Sub
        End If
    End Sub


    Private Sub ChCat_ID_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChCat_ID.CheckedChanged

        For J = 1 To FG.Rows - 1
            FG.Row = J
            If Trim(FG.get_TextMatrix(J, 3)) <> "" Then
                If ChCat_ID.Checked = True Then
                    FG.Col = 7
                    FG.CellBackColor = Color.LightCyan
                Else
                    FG.Col = 7
                    FG.CellBackColor = Color.White
                    FG.CellBackColor = Color.White
                    FG.CellForeColor = Color.Gray
                    FG.Col = 8
                    FG.CellForeColor = Color.Gray
                    FG.Col = 9
                    FG.CellForeColor = Color.Gray

                    FG.set_TextMatrix(FG.Row, 7, "0")
                    FG.set_TextMatrix(FG.Row, 8, "ບໍ່ເລືອກ")
                    FG.set_TextMatrix(FG.Row, 9, "No Selete")
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
        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            .Open("SELECT   " & SLF & "  FROM         gen_jn INNER JOIN    Acc_Code ON gen_jn.ac_code = Acc_Code.Ac_Code WHERE gen_jn.certify = N'" & MdCertifyId2 & "' And  year(date_work)=" & Format(dtActi.Value, "yyyy") & "  order by gen_jn.cnt", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim Rpt As New CryNewsJerneralJournal
        If MdShowLOGO = 1 Then
            Rpt.Subreports(0).SetDataSource(RsLOGO)
        End If
        Rpt.SetDataSource(Rs)
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
        For i = 1 To FG.Rows - 2
            'If FG.Row > 1 Then
            FG.set_TextMatrix(i, 10, Cmb.Text)
            FG.set_TextMatrix(i, 11, Format(CDbl(txtRate.Text), "#,##0.00"))
            FG.set_TextMatrix(i, 12, Format(CDbl(CDbl(txtRate.Text) * FG.get_TextMatrix(i, 5)), "#,##0.00"))
            FG.set_TextMatrix(i, 13, Format(CDbl(CDbl(txtRate.Text) * FG.get_TextMatrix(i, 6)), "#,##0.00"))
            'End If
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
        If RadioButton1.Checked = True Then 
            Dim Rs As New ADODB.Recordset
            With Rs
                Call LoadSqlData("select    * from Customer where 1=1  and Name=N'" & CmbCust.Text & "'   ", Rs)
                If .RecordCount > 0 Then
                    TxtCustID.Text = Trim(.Fields("Code").Value.ToString)
                Else
                    TxtCustID.Text = ""
                End If
            End With

        End If
    End Sub

    Private Sub CmbSupp_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbSupp.SelectedIndexChanged
        If RadioButton2.Checked = True Then
            Dim Rs As New ADODB.Recordset
            With Rs
                Call LoadSqlData("select    * from Supplier where 1=1  and Name=N'" & CmbSupp.Text & "'   ", Rs)
                If .RecordCount > 0 Then
                    TxtSuppID.Text = Trim(.Fields("Code").Value.ToString)
                Else
                    TxtSuppID.Text = ""
                End If
            End With
   
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            Dim Rs As New ADODB.Recordset
            With Rs
                Call LoadSqlData("select    * from Supplier where 1=1  and Name=N'" & CmbSupp.Text & "'   ", Rs)
                If .RecordCount > 0 Then
                    TxtSuppID.Text = Trim(.Fields("Code").Value.ToString)
                Else
                    TxtSuppID.Text = ""
                End If
            End With
          
        Else
            TxtSuppID.Text = ""
        End If
    
        CmbSupp.Enabled = True
        CmbCust.Enabled = False
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            Dim Rs As New ADODB.Recordset
            With Rs
                Call LoadSqlData("select    * from Customer where 1=1  and Name=N'" & CmbCust.Text & "'   ", Rs)
                If .RecordCount > 0 Then
                    TxtCustID.Text = Trim(.Fields("Code").Value.ToString)
                Else
                    TxtCustID.Text = ""
                End If
            End With
       
        Else
            TxtCustID.Text = ""
        End If

        CmbSupp.Enabled = False
        CmbCust.Enabled = True
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            Panel4.Enabled = True

        Else
            Panel4.Enabled = False
        End If
        RadioButton1.Checked = True
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

        Call ConnectAccess()


        Dim Conn As New ADODB.Connection
        Dim rsProj As New ADODB.Recordset
        Call LoadAcData("Select * from Conect ", rsProj)
        With rsProj
            If .RecordCount <> 0 Then
                MDServerName = (.Fields("ServerName").Value.ToString)
                MDDatabaName = (.Fields("DatabaseName").Value.ToString)
                MDServerUser = (.Fields("UserName").Value.ToString)
                MDServerPassword = (.Fields("UserPassword").Value.ToString)
                MDSeriaAccess = (.Fields("PartitionSeria").Value.ToString)
                'SPW = CStr((.Fields("SavePassword").Value.ToString))
                'SUSID = CStr((.Fields("SaveUserID").Value.ToString))
            End If
        End With
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
End Class