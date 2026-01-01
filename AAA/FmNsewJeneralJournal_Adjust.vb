Public Class FmNsewJeneralJournal_Adjust
    Dim sql As String
    Dim Rate1 As String
    Dim MdCertifyId, MdCertifyId2, Sdate As String
    Dim RateType As String
    Dim IVN As String
    Dim Book As String
    Dim Amount_In_Word As String
    Dim FgR As Integer
    Dim FgC As Integer

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        fmShartOfAccDetail.ShowDialog()
    End Sub

    Private Sub SetupGrid()
        FG.Columns.Clear()
        FG.Columns.Add("No", "ລ/ດ") ' 0
        FG.Columns.Add("CodeDr", "ເລກບັນຊີໜີ") ' 1
        FG.Columns.Add("CodeCr", "ເລກບັນຊີມີ") ' 2
        FG.Columns.Add("AccL", "ຊື່ບັນຊີ (ລາວ)") ' 3
        FG.Columns.Add("AccE", "ຊື່ບັນຊີ (ອັງກິດ)") ' 4
        FG.Columns.Add("AmtDr", "ຈຳນວນເງິນຈົດໜີ້") ' 5
        FG.Columns.Add("AmtCr", "ຈຳນວນເງິນຈົດມີ") ' 6
        FG.Columns.Add("CatID", "ລະຫັດ") ' 7
        FG.Columns.Add("CatL", "ຕົ້ນທຶນພາສາ (ລາວ)") ' 8
        FG.Columns.Add("CatE", "ຕົ້ນທຶນພາສາ (ອັງກິດ)") ' 9
        FG.Columns.Add("Curr", "ສະກຸນເງິນເງິນ") ' 10
        FG.Columns.Add("Rate", "ອັດຕາແລກປ່ຽນ") ' 11
        FG.Columns.Add("ValDr", "ມູນຄ່າໜີ້") ' 12
        FG.Columns.Add("ValCr", "ມູນຄ່າມີ") ' 13
        FG.Columns.Add("Col14", "1111") ' 14
        FG.Columns.Add("Col15", "22") ' 15

        FG.Columns(0).Width = 50
        FG.Columns(1).Width = 100
        FG.Columns(2).Width = 100
        FG.Columns(3).Width = 200
        FG.Columns(4).Width = 200
        FG.Columns(5).Width = 120
        FG.Columns(6).Width = 120
        FG.Columns(7).Width = 60
        FG.Columns(8).Width = 150
        FG.Columns(9).Width = 150
        FG.Columns(10).Width = 80
        FG.Columns(11).Width = 100
        FG.Columns(12).Width = 120
        FG.Columns(13).Width = 120
        FG.Columns(14).Width = 80
        FG.Columns(15).Width = 80

        FG.AllowUserToAddRows = False
        FG.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub SetupGridBee()
        Bee.Columns.Clear()
        Bee.Columns.Add("No", "ລ/ດ") ' 0
        Bee.Columns.Add("Col1", "") ' 1
        Bee.Columns.Add("Col2", "") ' 2
        Bee.Columns.Add("Col3", "") ' 3
        Bee.AllowUserToAddRows = False
    End Sub

    Private Sub Savedata()
        Dim J As Integer
        For J = 0 To FG.Rows.Count - 1
            '===============
            FG.Rows(J).Cells(10).Value = Cmb.Text
            FG.Rows(J).Cells(11).Value = Format(CDbl(txtRate.Text), "#,##0.00")

            Dim cell5Val As Object = FG.Rows(J).Cells(5).Value
            Dim cell6Val As Object = FG.Rows(J).Cells(6).Value
            Dim cell12Val As Object = FG.Rows(J).Cells(12).Value
            Dim cell13Val As Object = FG.Rows(J).Cells(13).Value

            If Cmb.Text = "USD" Then
                FG.Rows(J).Cells(14).Value = Format(CDbl(If(cell5Val Is Nothing OrElse cell5Val.ToString() = "", 0, cell5Val)), "#,##0.00")
                FG.Rows(J).Cells(15).Value = Format(CDbl(If(cell6Val Is Nothing OrElse cell6Val.ToString() = "", 0, cell6Val)), "#,##0.00")
            Else
                FG.Rows(J).Cells(14).Value = Format(CDbl(If(cell12Val Is Nothing OrElse cell12Val.ToString() = "", 0, cell12Val)) / CDbl(txtRateUSD.Text), "#,##0.00")
                FG.Rows(J).Cells(15).Value = Format(CDbl(If(cell13Val Is Nothing OrElse cell13Val.ToString() = "", 0, cell13Val)) / CDbl(txtRateUSD.Text), "#,##0.00")
            End If
        Next J

        If MuLng = "L" Then
            Amount_In_Word = txtAmt_letter.Text
        Else
            Amount_In_Word = txtAmt_letter_E.Text
        End If
        MuSubOff = Mid(Off_Usr.Text, 1, 5)
        For i = 0 To FG.Rows.Count - 1
            Dim cell1Val As Object = FG.Rows(i).Cells(1).Value
            Dim cell2Val As Object = FG.Rows(i).Cells(2).Value
            Dim cell3Val As Object = FG.Rows(i).Cells(3).Value
            Dim cell5Val As Object = FG.Rows(i).Cells(5).Value
            Dim cell6Val As Object = FG.Rows(i).Cells(6).Value
            Dim cell7Val As Object = FG.Rows(i).Cells(7).Value
            Dim cell10Val As Object = FG.Rows(i).Cells(10).Value
            Dim cell11Val As Object = FG.Rows(i).Cells(11).Value
            Dim cell12Val As Object = FG.Rows(i).Cells(12).Value
            Dim cell13Val As Object = FG.Rows(i).Cells(13).Value
            Dim cell14Val As Object = FG.Rows(i).Cells(14).Value
            Dim cell15Val As Object = FG.Rows(i).Cells(15).Value

            If (cell1Val Is Nothing OrElse cell1Val.ToString() = "") And (cell2Val Is Nothing OrElse cell2Val.ToString() = "") Then
                FG.Rows.Clear()
                FG.Rows.Add()
                AutoNumber()
                'Call NewText()
                Exit Sub
            End If
            Autox()
            txtAmount.Text = txtSumAmountDr.Text
            If CheckBox3.Checked = True Then
                Dim KKK As String = "INSERT INTO AP_ACC_adjust_Item( date_work, ac_Name, book, certify,cheque_no ,descrip ,descripe ,amount , curr ,rate, Rate_USD, net_amt ,code_dr ,code_cr ,ac_code    ,amt_dr , amt_cr , amount_dr ,amount_cr  ,certis, lock ,rec_lock , last_update , last_user , descflag , Cat_ID , certifyID  ,company  ,Office_ID  , del , AG) " & _
                   "Values('" & Format(dtActi.Value, "yyyy/MM/dd") & "',N'" & Apostrophe(If(cell3Val Is Nothing, "", cell3Val.ToString())) & "',N'" & CmbBook.Text & "',N'" & txtInvoice.Text & "','" & Apostrophe(txtChecq.Text) & "',N'" & Apostrophe(txtDesc.Text) & "',N'" & Apostrophe(txtDescE.Text) & "'," & CDbl(txtAmount.Text) & ",'" & If(cell10Val Is Nothing, "", cell10Val.ToString()) & "'," & CDbl(If(cell11Val Is Nothing OrElse cell11Val.ToString() = "", 0, cell11Val)) & "," & CDbl(txtRateUSD.Text) & ",'" & "0" & "','" & If(cell1Val Is Nothing, "", cell1Val.ToString()) & "','" & If(cell2Val Is Nothing, "", cell2Val.ToString()) & "','" & (If(cell1Val Is Nothing, "", cell1Val.ToString()) & If(cell2Val Is Nothing, "", cell2Val.ToString())) & "'," & CDbl(If(cell5Val Is Nothing OrElse cell5Val.ToString() = "", 0, cell5Val)) & "," & CDbl(If(cell6Val Is Nothing OrElse cell6Val.ToString() = "", 0, cell6Val)) & "," & CDbl(0) & "," & CDbl(0) & ",'" & "3" & "','" & "4" & "','" & "5" & "','" & dtActi.Text & "','" & MUserID & "','" & "8" & "','" & If(cell7Val Is Nothing, "", cell7Val.ToString()) & "','" & MdCertifyId & "','" & MuSubOff & "' ,'" & MuSubOff & "' , 0,1)"
                CNN.Execute(KKK)
            Else
                CNN.Execute("INSERT INTO AP_ACC_adjust_Item( date_work, ac_Name, book, certify,cheque_no ,descrip ,descripe ,amount , curr ,rate, Rate_USD, net_amt ,code_dr ,code_cr ,ac_code  , amount_dr ,amount_cr ,amt_dr , amt_cr   ,amt_USD_Dr, amt_USD_Cr  ,certis, lock ,rec_lock , last_update , last_user , descflag , Cat_ID , certifyID  ,company   ,Office_ID , del, AG) " & _
                   "Values('" & Format(dtActi.Value, "yyyy/MM/dd") & "',N'" & Apostrophe(If(cell3Val Is Nothing, "", cell3Val.ToString())) & "',N'" & CmbBook.Text & "',N'" & txtInvoice.Text & "','" & Apostrophe(txtChecq.Text) & "',N'" & Apostrophe(txtDesc.Text) & "',N'" & Apostrophe(txtDescE.Text) & "'," & CDbl(txtAmount.Text) & ",'" & If(cell10Val Is Nothing, "", cell10Val.ToString()) & "'," & CDbl(If(cell11Val Is Nothing OrElse cell11Val.ToString() = "", 0, cell11Val)) & "," & CDbl(txtRateUSD.Text) & ",'" & "0" & "','" & If(cell1Val Is Nothing, "", cell1Val.ToString()) & "','" & If(cell2Val Is Nothing, "", cell2Val.ToString()) & "','" & (If(cell1Val Is Nothing, "", cell1Val.ToString()) & If(cell2Val Is Nothing, "", cell2Val.ToString())) & "'," & CDbl(If(cell5Val Is Nothing OrElse cell5Val.ToString() = "", 0, cell5Val)) & "," & CDbl(If(cell6Val Is Nothing OrElse cell6Val.ToString() = "", 0, cell6Val)) & "," & CDbl(If(cell12Val Is Nothing OrElse cell12Val.ToString() = "", 0, cell12Val)) & "," & CDbl(If(cell13Val Is Nothing OrElse cell13Val.ToString() = "", 0, cell13Val)) & "," & CDbl(If(cell14Val Is Nothing OrElse cell14Val.ToString() = "", 0, cell14Val)) & "," & CDbl(If(cell15Val Is Nothing OrElse cell15Val.ToString() = "", 0, cell15Val)) & ",'" & "3" & "','" & "4" & "','" & "5" & "','" & dtActi.Text & "','" & MUserID & "','" & "8" & "','" & If(cell7Val Is Nothing, "", cell7Val.ToString()) & "','" & MdCertifyId & "','" & MuSubOff & "' , '" & MuSubOff & "' , 0,0)")
            End If
        Next i
        MuSubOff = MuSubOff2
        LngId = "6001" : MsgRpt()
    End Sub

    Private Sub FG_CellEndEdit(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FG.CellEndEdit
        Dim colIndex As Integer = e.ColumnIndex
        Dim rowIndex As Integer = e.RowIndex

        If colIndex = 1 Or colIndex = 2 Then
            ' Empty check logic
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
        '=========jjjjjjj
        txtTotal_Amt_LAK.Text = Format(CDbl(txtRate.Text) * CDbl(txtAmount.Text), "#,##0.00")
        If ChCat_ID.Checked = False Then
            AfterEdit(rowIndex, colIndex)
            loadColor()
            'MsgBox("jj")
            txtAmount.Focus()
        Else
            'AfterEdit()
            'loadColor()
            'txtAmount.Focus()
        End If

        If txtDesc.Text = "" Then
            Dim cell3Val As Object = FG.Rows(0).Cells(3).Value
            Dim cell4Val As Object = FG.Rows(0).Cells(4).Value
            txtDesc.Text = If(cell3Val Is Nothing, "", cell3Val.ToString())
            txtDescE.Text = If(cell4Val Is Nothing, "", cell4Val.ToString())
        End If
        If CDbl(txtAmount.Text) = 0 Then
            If rowIndex = 1 Then ' 1-based index 2 was the second row. In 0-based it's 1.
                Dim cell1_5Val As Object = FG.Rows(0).Cells(5).Value
                Dim cell1_6Val As Object = FG.Rows(0).Cells(6).Value
                txtAmount.Text = Format(CDbl((CDbl(If(cell1_5Val Is Nothing OrElse cell1_5Val.ToString() = "", 0, cell1_5Val)) + CDbl(If(cell1_6Val Is Nothing OrElse cell1_6Val.ToString() = "", 0, cell1_6Val)))), "##,##0.00")
            End If
        End If
        If CDbl(txtSumAmountDr.Text) >= CDbl(txtSumAmountCr.Text) Then
            txtAmount.Text = txtSumAmountDr.Text
        End If
        If CDbl(txtSumAmountDr.Text) <= CDbl(txtSumAmountCr.Text) Then
            txtAmount.Text = txtSumAmountCr.Text
        End If
    End Sub
    Private Function GetValue(ByVal cellVal As Object) As Double
        If cellVal Is Nothing OrElse cellVal.ToString() = "" Then
            Return 0
        Else
            If IsNumeric(cellVal) Then
                Return CDbl(cellVal)
            End If
            Return 0
        End If
    End Function

    Private Function GetString(ByVal cellVal As Object) As String
        If cellVal Is Nothing Then Return ""
        Return cellVal.ToString()
    End Function

    Private Sub AfterEdit(Optional ByVal rowIndex As Integer = -1, Optional ByVal colIndex As Integer = -1)
        If rowIndex = -1 Then rowIndex = FG.CurrentCell.RowIndex
        If colIndex = -1 Then colIndex = FG.CurrentCell.ColumnIndex

        BtnMove.Visible = False
        '*************************Col-1-*********************
        If colIndex = 1 Then
            R = rowIndex
            L = colIndex
            Dim cell1Val As Object = FG.Rows(rowIndex).Cells(1).Value
            If cell1Val Is Nothing OrElse cell1Val.ToString() = "" Then
                MDSearchAcccode = ""
                fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                fmShartOfAccDetail.ShowDialog()
            End If
            cell1Val = FG.Rows(rowIndex).Cells(1).Value
            If cell1Val IsNot Nothing AndAlso cell1Val.ToString() <> "" Then
                AccId = cell1Val.ToString()
                MDSearchAcccode = AccId
                LoadText()
                FG.Rows(rowIndex).Cells(3).Value = AccName
                FG.Rows(rowIndex).Cells(4).Value = AccNamee
                If FG.Rows(rowIndex).Cells(3).Value Is Nothing OrElse FG.Rows(rowIndex).Cells(3).Value.ToString() = "" Then
                    FG.Rows(rowIndex).Cells(1).Value = ""
                    fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                    fmShartOfAccDetail.ShowDialog()
                End If
                SumAmountDr()
                txtSumAmountDr.Text = (CDbl(txtSumAmountDr.Text) - GetValue(FG.Rows(rowIndex).Cells(5).Value)).ToString()
                FG.Rows(rowIndex).Cells(5).Value = Format(CDbl(CDbl(txtAmount.Text) - CDbl(txtSumAmountDr.Text)), "#,##0.00")
                FG.Rows(rowIndex).Cells(6).Value = "0.00"
                FG.Rows(rowIndex).Cells(7).Value = 0
                FG.Rows(rowIndex).Cells(10).Value = Cmb.Text
                FG.Rows(rowIndex).Cells(11).Value = Format(CDbl(txtRate.Text), "#,##0.00")
                FG.Rows(rowIndex).Cells(12).Value = Format(CDbl(CDbl(txtRate.Text) * GetValue(FG.Rows(rowIndex).Cells(5).Value)), "#,##0.00")
                FG.Rows(rowIndex).Cells(13).Value = "0.00"
                SumAmountDr()

                If ChDe.Checked = True Then
                    If FG.Rows(FG.Rows.Count - 1).Cells(3).Value IsNot Nothing AndAlso FG.Rows(FG.Rows.Count - 1).Cells(3).Value.ToString() <> "" Then
                        FG.Rows.Add()
                    End If
                    Exit Sub
                End If

                If FG.Rows(FG.Rows.Count - 1).Cells(3).Value IsNot Nothing AndAlso FG.Rows(FG.Rows.Count - 1).Cells(3).Value.ToString() <> "" Then
                    FG.Rows.Add()
                    Exit Sub
                End If
                Exit Sub
            End If
            Exit Sub
        End If
        '*************************Col-2-*********************
        If colIndex = 2 Then
            R = rowIndex
            L = colIndex
            Dim cell2Val As Object = FG.Rows(rowIndex).Cells(2).Value
            If cell2Val Is Nothing OrElse cell2Val.ToString() = "" Then
                MDSearchAcccode = ""
                fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                fmShartOfAccDetail.ShowDialog()
            End If
            cell2Val = FG.Rows(rowIndex).Cells(2).Value
            If cell2Val IsNot Nothing AndAlso cell2Val.ToString() <> "" Then
                AccId = cell2Val.ToString()
                LoadText()
                FG.Rows(rowIndex).Cells(3).Value = AccName
                FG.Rows(rowIndex).Cells(4).Value = AccNamee
                If FG.Rows(rowIndex).Cells(3).Value Is Nothing OrElse FG.Rows(rowIndex).Cells(3).Value.ToString() = "" Then
                    MDSearchAcccode = cell2Val.ToString()
                    fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_DR"
                    fmShartOfAccDetail.ShowDialog()
                End If
                SumAmountDr()
                txtSumAmountCr.Text = (CDbl(txtSumAmountCr.Text) - GetValue(FG.Rows(rowIndex).Cells(6).Value)).ToString()
                FG.Rows(rowIndex).Cells(6).Value = Format(CDbl(CDbl(txtAmount.Text) - CDbl(txtSumAmountCr.Text)), "#,##0.00")
                FG.Rows(rowIndex).Cells(5).Value = "0.00"
                FG.Rows(rowIndex).Cells(7).Value = 0
                FG.Rows(rowIndex).Cells(10).Value = Cmb.Text
                FG.Rows(rowIndex).Cells(11).Value = Format(CDbl(txtRate.Text), "#,##0.00")
                FG.Rows(rowIndex).Cells(13).Value = Format(CDbl(CDbl(txtRate.Text) * GetValue(FG.Rows(rowIndex).Cells(6).Value)), "#,##0.00")
                FG.Rows(rowIndex).Cells(12).Value = "0.00"
                SumAmountDr()
                If ChDe.Checked = True Then
                    If FG.Rows(FG.Rows.Count - 1).Cells(3).Value IsNot Nothing AndAlso FG.Rows(FG.Rows.Count - 1).Cells(3).Value.ToString() <> "" Then
                        FG.Rows.Add()
                    End If
                    Exit Sub
                End If
                If FG.Rows(FG.Rows.Count - 1).Cells(3).Value IsNot Nothing AndAlso FG.Rows(FG.Rows.Count - 1).Cells(3).Value.ToString() <> "" Then
                    FG.Rows.Add()
                End If
                Exit Sub
            End If
            Exit Sub
        End If

        '*************************Col-5/6 (Amounts)*********************
        If colIndex = 5 Or colIndex = 6 Then
            Dim cellVal As Object = FG.Rows(rowIndex).Cells(colIndex).Value
            If IsNumeric(cellVal) = False Then
                MessageBox.Show("ກະລຸນນາໃສ່ໂຕເລກ")
                FG.Rows(rowIndex).Cells(colIndex).Value = "0.00"
                Exit Sub
            End If

            FG.Rows(rowIndex).Cells(colIndex).Value = Format(CDbl(cellVal), "#,##0.00")
            If colIndex = 5 Then
                FG.Rows(rowIndex).Cells(6).Value = "0.00"
                FG.Rows(rowIndex).Cells(12).Value = Format(CDbl(CDbl(txtRate.Text) * CDbl(FG.Rows(rowIndex).Cells(5).Value)), "#,##0.00")
                FG.Rows(rowIndex).Cells(13).Value = "0.00"
            Else
                FG.Rows(rowIndex).Cells(5).Value = "0.00"
                FG.Rows(rowIndex).Cells(13).Value = Format(CDbl(CDbl(txtRate.Text) * CDbl(FG.Rows(rowIndex).Cells(6).Value)), "#,##0.00")
                FG.Rows(rowIndex).Cells(12).Value = "0.00"
            End If
            
            FG.Rows(rowIndex).Cells(7).Value = 0
            FG.Rows(rowIndex).Cells(10).Value = Cmb.Text
            FG.Rows(rowIndex).Cells(11).Value = Format(CDbl(txtRate.Text), "#,##0.00")
            
            SumAmountDr()
            
            ' Balancing and Save confirmation logic
            If CDbl(txtTotal_Amt_LAK.Text) > 0 Then
                If CDbl(txtSumTotalAmountDr.Text) = CDbl(txtTotal_Amt_LAK.Text) AndAlso CDbl(txtSumTotalAmountCr.Text) = CDbl(txtTotal_Amt_LAK.Text) Then
                    If MessageBox.Show("ບັນຊີນີ້ດູນດ່ຽງແລ້ວ ທ່ານຕອ້ງການບັນທຶຫລືບໍ່!", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                        If txtInvoice.Enabled = True Then
                            AutoNumber()
                            If txtInvoice.Text = "" Then
                                MsgBox("ກະລຸນນາໃສ່ເລກໃບຢັງຢືນກອ່ນ!", MsgBoxStyle.OkOnly)
                                txtInvoice.BackColor = Color.Red
                                txtInvoice.Focus()
                                Exit Sub
                            End If

                             ' Verify if Invoice exists
                            Call LoadSqlData("SELECT top 1 Right(certify,3) As certify FROM AP_ACC_adjust_Item WHERE book ='" & CmbBook.Text & "' And certify = N'" & txtInvoice.Text & "' And year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And Month(date_work)=" & Format(CDate(dtActi.Value), "MM") & " Order by Right(certify,7) DESC ", RSC)
                            If RSC.RecordCount > 0 Then
                                MsgBox("ເລກລະຫັດ : " & Trim(txtInvoice.Text) & " ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
                                txtInvoice.BackColor = Color.Red
                                txtInvoice.Focus()
                                If RSC.State = ConnectionState.Open Then RSC.Close()
                                Exit Sub
                            End If

                            Savedata()
                        Else
                            ' Update existing
                            Dim prevListFG As DataGridView = FmJeneralJournal_Adjust_List.FG
                            Dim prevRowIndex As Integer = prevListFG.CurrentCell.RowIndex
                            CNN.Execute("DELETE FROM AP_ACC_adjust_Item WHERE book ='" & prevListFG.Rows(prevRowIndex).Cells(15).Value.ToString() & "' And certify  = '" & txtInvoice.Text & "'   And  year(date_work)='" & Format(CDate(prevListFG.Rows(prevRowIndex).Cells(0).Value), "yyyy") & "' ")
                            Savedata()
                        End If
                        
                        If CheckBox1.Checked = True Then Call LoadReport()

                        txtInvoice.Enabled = True
                        CmbBook.Enabled = True
                        Panel1.Visible = False
                        BtnMove.Visible = False
                        BtnSearch.Visible = False
                        FG.Rows.Clear()
                        FG.Rows.Add()
                        FG.Rows.Add()
                        FG.CurrentCell = FG.Rows(0).Cells(1)
                        CmbBook.Focus()

                        If CheckBox2.Checked = True Then Close()
                        Exit Sub
                    End If
                End If
            End If

            ' Navigation logic
            Dim cell1Val_nav As Object = FG.Rows(rowIndex).Cells(1).Value
            If cell1Val_nav IsNot Nothing AndAlso cell1Val_nav.ToString() <> "" Then
                If rowIndex + 1 < FG.Rows.Count Then
                    Dim nextCell1Val As Object = FG.Rows(rowIndex + 1).Cells(1).Value
                    If nextCell1Val IsNot Nothing AndAlso nextCell1Val.ToString() <> "" Then
                        FG.CurrentCell = FG.Rows(rowIndex + 1).Cells(1)
                        Exit Sub
                    End If
                    FG.CurrentCell = FG.Rows(rowIndex + 1).Cells(2)
                End If
                SumAmountDr()
            Else
                If rowIndex + 1 < FG.Rows.Count Then
                    Dim nextCell2Val As Object = FG.Rows(rowIndex + 1).Cells(2).Value
                    If nextCell2Val IsNot Nothing AndAlso nextCell2Val.ToString() <> "" Then
                        FG.CurrentCell = FG.Rows(rowIndex + 1).Cells(2)
                        Exit Sub
                    End If
                    FG.CurrentCell = FG.Rows(rowIndex + 1).Cells(1)
                End If
                SumAmountDr()
            End If
            Exit Sub
        End If

        '*************************Col-7-*********************
        If colIndex = 7 Then
            ' Logic for Col 7
            Dim cellVal As Object = FG.Rows(rowIndex).Cells(7).Value
            If IsNumeric(cellVal) Then
                Dim code As Integer = CInt(cellVal)
                If code >= 0 And code <= 4 Then
                    Select Case code
                        Case 0
                            FG.Rows(rowIndex).Cells(8).Value = "ບໍ່ເລືອກ"
                            FG.Rows(rowIndex).Cells(9).Value = "No Selete"
                        Case 1
                            FG.Rows(rowIndex).Cells(8).Value = "ຮັບໃຊ້ການພະລິດ"
                            FG.Rows(rowIndex).Cells(9).Value = "Use build"
                        Case 2
                            FG.Rows(rowIndex).Cells(8).Value = "ຮັບໃຊ້ການຈຳໜ່າຍ"
                            FG.Rows(rowIndex).Cells(9).Value = "Use Sell"
                        Case 3
                            FG.Rows(rowIndex).Cells(8).Value = "ຮັບໃຊ້ບໍລິຫານ"
                            FG.Rows(rowIndex).Cells(9).Value = "Use manage "
                        Case 4
                            FG.Rows(rowIndex).Cells(8).Value = "ຕົ້ນທຶນຂາຍ/ຕົ້ນທຶນບໍລິຫານ"
                            FG.Rows(rowIndex).Cells(9).Value = "Sell capital/manage capital "
                    End Select
                Else
                    MessageBox.Show("ລະຫັດມີພຽງເລກ 0 ຫາເລກ 4 ເທົ່ານັ້ນ")
                    FG.Rows(rowIndex).Cells(7).Value = "0"
                    FG.Rows(rowIndex).Cells(8).Value = "ບໍ່ເລືອກ"
                    FG.Rows(rowIndex).Cells(9).Value = "No Selete"
                End If
            End If
            Exit Sub
        End If
    End Sub
    Public Sub AutoNumber()

        Dim srNum As New ADODB.Recordset
        Dim mNum As Integer

        'Call LoadSqlData("SELECT top 1 Right(certify,7) As  certify  FROM AP_ACC_adjust_Item where year(date_work)='" & Format(dtActi.Value, "yyyy") & "'   Order by   Right(certify,7) DESC", srNum)
        Dim ss As String = ""
        ss = "SELECT top 1 Right(certify,3) As  certify   FROM  AP_ACC_adjust_Item where  book =N'" & CmbBook.Text & "' And  year(date_work)='" & Format(dtActi.Value, "yyyy") & "'  " & _
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
        If R >= 0 And R < FG.Rows.Count And (L + 4) >= 0 And (L + 4) < FG.Columns.Count Then
            FG.CurrentCell = FG.Rows(R).Cells(L + 4)
        End If
    End Sub
    Private Sub NewText()
        dtActi.Text = MWorkSetting
        txtDesc.Text = ""
        txtDescE.Text = ""
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
        Call FmJeneralJournal_Adjust_List.LoadMonthSQL()
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
        SetupGrid()
        SetupGridBee()
        LoadCurr()
        LoadBook()
        loadOffice_User()
        ' The original form loads some data into the grid upon load.
        ' If FmJeneralJournal_Adjust_List is the source of data, then
        ' the following line would simulate that. Needs further investigation
        ' if this form is standalone or always opened from a list.
        ' For now, assume it's opened from the list and data is passed or loaded.
        ' Dim prevRowIndex As Integer = FmJeneralJournal_Adjust_List.FG.CurrentCell.RowIndex
        ' txtInvoice.Text = GetString(FmJeneralJournal_Adjust_List.FG.CurrentRow.Cells(4).Value)
        ' LoadListFG()
        ' SumAmountDr()
        ' loadColor()
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

    Public Sub AddAcc()
        If FG.CurrentCell Is Nothing Then Exit Sub
        Dim rowIndex As Integer = FG.CurrentCell.RowIndex
        Dim colIndex As Integer = FG.CurrentCell.ColumnIndex

        If colIndex = 1 Then
            AccId = If(FG.Rows(rowIndex).Cells(1).Value Is Nothing, "", FG.Rows(rowIndex).Cells(1).Value.ToString())
            LoadText()
            FG.Rows(rowIndex).Cells(3).Value = AccName
            MDSearchAcccode = If(FG.Rows(rowIndex).Cells(1).Value Is Nothing, "", FG.Rows(rowIndex).Cells(1).Value.ToString())
            txtSumAmountDr.Text = (CDbl(txtSumAmountDr.Text) - CDbl(If(FG.Rows(rowIndex).Cells(5).Value Is Nothing OrElse FG.Rows(rowIndex).Cells(5).Value.ToString() = "", 0, FG.Rows(rowIndex).Cells(5).Value))).ToString()
            FG.Rows(rowIndex).Cells(5).Value = Format(CDbl(CDbl(txtAmount.Text) - CDbl(txtSumAmountDr.Text)), "#,##0.00")
            FG.Rows(rowIndex).Cells(6).Value = "0.00"
            FG.Rows(rowIndex).Cells(7).Value = 0
            FG.Rows(rowIndex).Cells(10).Value = Cmb.Text
            FG.Rows(rowIndex).Cells(11).Value = Format(CDbl(txtRate.Text), "#,##0.00")
            FG.Rows(rowIndex).Cells(12).Value = Format(CDbl(CDbl(txtRate.Text) * CDbl(If(FG.Rows(rowIndex).Cells(5).Value Is Nothing OrElse FG.Rows(rowIndex).Cells(5).Value.ToString() = "", 0, FG.Rows(rowIndex).Cells(5).Value))), "#,##0.00")
            FG.Rows(rowIndex).Cells(13).Value = "0.00"
            SumAmountDr()
            
            If ChDe.Checked = True Then
                ' Logic for column focus omitted, but adding Row:
                If FG.Rows(FG.Rows.Count - 1).Cells(3).Value IsNot Nothing AndAlso FG.Rows(FG.Rows.Count - 1).Cells(3).Value.ToString() <> "" Then
                    FG.Rows.Add()
                End If
                Exit Sub
            End If

            'FG.Col = 5 ' Removed as per instruction to replace FG.Col for access, not setting cursor

            If FG.Rows(FG.Rows.Count - 1).Cells(3).Value IsNot Nothing AndAlso FG.Rows(FG.Rows.Count - 1).Cells(3).Value.ToString() <> "" Then
                FG.Rows.Add()
                'Timer1.Enabled = True

            End If
            'Call loadColor()
        End If
        If colIndex = 2 Then
            AccId = If(FG.Rows(rowIndex).Cells(2).Value Is Nothing, "", FG.Rows(rowIndex).Cells(2).Value.ToString())
            LoadText()
            FG.Rows(rowIndex).Cells(3).Value = AccName
            MDSearchAcccode = If(FG.Rows(rowIndex).Cells(2).Value Is Nothing, "", FG.Rows(rowIndex).Cells(2).Value.ToString())
            txtSumAmountCr.Text = (CDbl(txtSumAmountCr.Text) - CDbl(If(FG.Rows(rowIndex).Cells(6).Value Is Nothing OrElse FG.Rows(rowIndex).Cells(6).Value.ToString() = "", 0, FG.Rows(rowIndex).Cells(6).Value))).ToString()
            FG.Rows(rowIndex).Cells(6).Value = Format(CDbl(CDbl(txtAmount.Text) - CDbl(txtSumAmountCr.Text)), "#,##0.00")
            FG.Rows(rowIndex).Cells(5).Value = "0.00"
            FG.Rows(rowIndex).Cells(7).Value = 0
            FG.Rows(rowIndex).Cells(10).Value = Cmb.Text
            FG.Rows(rowIndex).Cells(11).Value = Format(CDbl(txtRate.Text), "#,##0.00")
            FG.Rows(rowIndex).Cells(13).Value = Format(CDbl(CDbl(txtRate.Text) * CDbl(If(FG.Rows(rowIndex).Cells(6).Value Is Nothing OrElse FG.Rows(rowIndex).Cells(6).Value.ToString() = "", 0, FG.Rows(rowIndex).Cells(6).Value))), "#,##0.00")
            FG.Rows(rowIndex).Cells(12).Value = "0.00"
            SumAmountDr()

            If ChDe.Checked = True Then
                If FG.Rows(FG.Rows.Count - 1).Cells(3).Value IsNot Nothing AndAlso FG.Rows(FG.Rows.Count - 1).Cells(3).Value.ToString() <> "" Then
                    FG.Rows.Add()
                End If
                Exit Sub
            End If

            'FG.Col = 6 ' Removed as per instruction to replace FG.Col for access, not setting cursor

            If FG.Rows(FG.Rows.Count - 1).Cells(3).Value IsNot Nothing AndAlso FG.Rows(FG.Rows.Count - 1).Cells(3).Value.ToString() <> "" Then
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
        TotalAmountDr = 0
        TotalAmountCr = 0
        
        Dim rowsToRemove As New List(Of Integer)
        For i = 0 To FG.Rows.Count - 1
            Dim cell3Val As Object = FG.Rows(i).Cells(3).Value
            If i <> FG.Rows.Count - 1 Then
                If cell3Val Is Nothing OrElse cell3Val.ToString() = "" Then
                    rowsToRemove.Add(i)
                End If
            End If
            
            Dim cell5Val As Object = FG.Rows(i).Cells(5).Value
            Dim cell6Val As Object = FG.Rows(i).Cells(6).Value
            Dim cell12Val As Object = FG.Rows(i).Cells(12).Value
            Dim cell13Val As Object = FG.Rows(i).Cells(13).Value
            
            AmountDr = AmountDr + CDbl(If(cell5Val Is Nothing OrElse cell5Val.ToString() = "", 0, cell5Val))
            AmountCr = AmountCr + CDbl(If(cell6Val Is Nothing OrElse cell6Val.ToString() = "", 0, cell6Val))
            TotalAmountDr = TotalAmountDr + CDbl(If(cell12Val Is Nothing OrElse cell12Val.ToString() = "", 0, cell12Val))
            TotalAmountCr = TotalAmountCr + CDbl(If(cell13Val Is Nothing OrElse cell13Val.ToString() = "", 0, cell13Val))
        Next
        
        For j As Integer = rowsToRemove.Count - 1 To 0 Step -1
            FG.Rows.RemoveAt(rowsToRemove(j))
        Next

        txtSumAmountDr.Text = Format(AmountDr, "#,##0.00")
        txtSumAmountCr.Text = Format(AmountCr, "#,##0.00")
        txtSumTotalAmountDr.Text = Format(TotalAmountDr, "#,##0.00")
        txtSumTotalAmountCr.Text = Format(TotalAmountCr, "#,##0.00")


        Dr.Text = (CDbl(txtSumAmountDr.Text) - CDbl(txtSumAmountCr.Text)).ToString()
        Cr.Text = (CDbl(txtSumAmountCr.Text) - CDbl(txtSumAmountDr.Text)).ToString()
        DDR.Text = (CDbl(txtSumTotalAmountDr.Text) - CDbl(txtSumTotalAmountCr.Text)).ToString()
        CCR.Text = (CDbl(txtSumTotalAmountCr.Text) - CDbl(txtSumTotalAmountDr.Text)).ToString()
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
        'fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal"
        fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal_Adjust"
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
    Private Sub FG_Scroll(ByVal sender As Object, ByVal e As ScrollEventArgs) Handles FG.Scroll
        BtnSearch.Visible = False
        BtnMove.Visible = False
    End Sub


    Private Sub FG_CellMouseDown(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) Handles FG.CellMouseDown
        If txtInvoice.BackColor = Color.Red Then
            txtInvoice.Focus()
            Exit Sub
        End If
        ' RowIndex < 0 means header cell
        If e.RowIndex >= 0 Then
            FG.CurrentCell = FG.Rows(e.RowIndex).Cells(e.ColumnIndex)
            MouseDownEvent(e.RowIndex, e.ColumnIndex)
        End If
    End Sub

    Public Sub MouseDownEvent(Optional ByVal rIndex As Integer = -1, Optional ByVal cIndex As Integer = -1)
        If rIndex = -1 Then rIndex = FG.CurrentCell.RowIndex
        If cIndex = -1 Then cIndex = FG.CurrentCell.ColumnIndex
        
        If MouseButtons = Windows.Forms.MouseButtons.Right Then
            FG.BeginEdit(True)
        ElseIf MouseButtons = Windows.Forms.MouseButtons.Left Then
            If FG.Rows.Count >= 3 Then
                BtnMove.Visible = True
                Dim rect As Rectangle = FG.GetCellDisplayRectangle(cIndex, rIndex, False)
                BtnMove.Top = rect.Top + FG.Top
            End If
            
            Dim cellVal As Object = FG.Rows(rIndex).Cells(cIndex).Value
            If cellVal IsNot Nothing AndAlso cellVal.ToString() <> "" Then
                If cIndex = 7 Then
                    If ChCat_ID.Checked = True Then
                        Panel1.Visible = True
                        ' Suggesting some default focus for Bee
                        If Bee.Rows.Count > 0 Then
                            Bee.CurrentCell = Bee.Rows(0).Cells(3)
                            Bee.Focus()
                        End If
                    End If
                Else
                    Panel1.Visible = False
                End If
            End If

            Dim cellDR As Object = FG.Rows(rIndex).Cells(1).Value
            Dim cellCR As Object = FG.Rows(rIndex).Cells(2).Value

            If cIndex = 1 Then
                BtnSearch.Visible = (cellCR Is Nothing OrElse cellCR.ToString() = "")
            ElseIf cIndex = 2 Then
                BtnSearch.Visible = (cellDR Is Nothing OrElse cellDR.ToString() = "")
            Else
                BtnSearch.Visible = False
            End If

            If BtnSearch.Visible Then
                Dim rect As Rectangle = FG.GetCellDisplayRectangle(cIndex, rIndex, False)
                BtnSearch.Top = rect.Top + FG.Top
                BtnSearch.Left = rect.Right + FG.Left - BtnSearch.Width
            End If

            If rIndex = FG.Rows.Count - 1 Then
                BtnMove.Visible = False
            End If
        End If
    End Sub

    Private Sub btnmove_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnMove.Click
        If FG.CurrentCell IsNot Nothing Then
            FG.Rows.RemoveAt(FG.CurrentCell.RowIndex)
            SumAmountDr()
        End If
        Panel1.Visible = False
        BtnMove.Visible = False
        BtnSearch.Visible = False
    End Sub



    ' Helper to simulate VSFlexGrid Redraw property logic if needed
    ' Private _redraw As Boolean = True

    Private Sub FG_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles FG.SelectionChanged
        If txtInvoice.BackColor = Color.Red Then
            txtInvoice.Focus()
            Exit Sub
        End If
        
        If FG.CurrentCell Is Nothing Then Exit Sub
        Dim r As Integer = FG.CurrentCell.RowIndex
        Dim c As Integer = FG.CurrentCell.ColumnIndex

        If c = 1 Then
            Dim cell2Val As Object = FG.Rows(r).Cells(2).Value
            If cell2Val IsNot Nothing AndAlso cell2Val.ToString() <> "" Then
                FG.Rows(r).ReadOnly = True
            Else
                FG.Rows(r).ReadOnly = False
                AccId = If(FG.Rows(r).Cells(c).Value Is Nothing, "", FG.Rows(r).Cells(c).Value.ToString())
            End If
        End If

        If c = 2 Then
            Dim cell1Val As Object = FG.Rows(r).Cells(1).Value
            If cell1Val IsNot Nothing AndAlso cell1Val.ToString() <> "" Then
                FG.Rows(r).ReadOnly = True
            Else
                FG.Rows(r).ReadOnly = False
                AccId = If(FG.Rows(r).Cells(c).Value Is Nothing, "", FG.Rows(r).Cells(c).Value.ToString())
            End If
        End If

        If c >= 3 And c <= 4 Then
            FG.Rows(r).ReadOnly = False
        End If

        If c > 2 Then
            Dim cell1Val As Object = FG.Rows(r).Cells(1).Value
            Dim cell2Val As Object = FG.Rows(r).Cells(2).Value
            If (cell1Val Is Nothing OrElse cell1Val.ToString() = "") AndAlso (cell2Val Is Nothing OrElse cell2Val.ToString() = "") Then
                ' Logic for non-editable if no account
            End If
        End If

        If c >= 3 And c <= 13 Then
            BtnSearch.Visible = False
        End If

        Dim cell5Val As Object = FG.Rows(r).Cells(5).Value
        If cell5Val Is Nothing OrElse cell5Val.ToString() = "" Then
            Panel1.Visible = False
        End If

        Dim cell7Val As Object = FG.Rows(r).Cells(7).Value
        If cell7Val IsNot Nothing Then
            Select Case cell7Val.ToString()
                Case "0"
                    FG.Rows(r).Cells(8).Value = "ບໍ່ເລືອກ"
                    FG.Rows(r).Cells(9).Value = "No Selete"
                Case "1"
                    FG.Rows(r).Cells(8).Value = "ຮັບໃຊ້ການພະລິດ"
                    FG.Rows(r).Cells(9).Value = "Use build"
                Case "2"
                    FG.Rows(r).Cells(8).Value = "ຮັບໃຊ້ການຈຳໜ່າຍ"
                    FG.Rows(r).Cells(9).Value = "Use Sell"
                Case "3"
                    FG.Rows(r).Cells(8).Value = "ຮັບໃຊ້ບໍລິຫານ"
                    FG.Rows(r).Cells(9).Value = "Use manage "
                Case "4"
                    FG.Rows(r).Cells(8).Value = "ຕົ້ນທຶນຂາຍ/ຕົ້ນທຶນບໍລິຫານ"
                    FG.Rows(r).Cells(9).Value = "Sell capital/manage capital "
            End Select
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
            txtTotal_Amt_LAK.Text = Format(CDbl(txtRate.Text) * CDbl(txtAmount.Text), "#,##0.00")
            Call FormatText()
            Dim cell1_1Val As Object = FG.Rows(0).Cells(1).Value
            Dim cell1_2Val As Object = FG.Rows(0).Cells(2).Value
            If (cell1_1Val Is Nothing OrElse cell1_1Val.ToString() = "") And (cell1_2Val Is Nothing OrElse cell1_2Val.ToString() = "") Then
                FG.CurrentCell = FG.Rows(0).Cells(1)
            ElseIf cell1_1Val IsNot Nothing AndAlso cell1_1Val.ToString() <> "" Then
                FG.CurrentCell = FG.Rows(0).Cells(1)
            ElseIf cell1_2Val IsNot Nothing AndAlso cell1_2Val.ToString() <> "" Then
                FG.CurrentCell = FG.Rows(0).Cells(2)
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
        If FG.CurrentCell IsNot Nothing Then
            R = FG.CurrentCell.RowIndex
            L = FG.CurrentCell.ColumnIndex
        End If
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
        Bee.Rows.Clear()
        Bee.Rows.Add("0", "ບໍ່ເລືອກ", "No Selete", " 0.   ບໍ່ເລືອກ")
        Bee.Rows.Add("1", "ຮັບໃຊ້ໃນການພະລິດ", "Use build", " 1.   ຮັບໃຊ້ໃນການພະລິດ")
        Bee.Rows.Add("2", "ຮັບໃຊ້ໃນການຈຳໜ່າຍ", "Use Sell", " 2.   ຮັບໃຊ້ໃນການຈຳໜ່າຍ")
        Bee.Rows.Add("3", "ຮັບໃຊ້ບໍລິຫານ", "Use manage", " 3.   ຮັບໃຊ້ບໍລິຫານ")
        Bee.Rows.Add("4", "ຕົ້ນທື່ນຂາຍ/ຕົ້ນທຶນບໍລິຫານ", "Sell capital/manage capital ", " 4.   ຕົ້ນທື່ນຂາຍ/ຕົ້ນທຶນບໍລິຫານ")
    End Sub

    ' Legacy Bee_AfterEdit event removed - DataGridView uses CellEndEdit

    Private Sub Bee_CellDoubleClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles Bee.CellDoubleClick
        Dim rowIndex As Integer = FG.CurrentCell.RowIndex
        FG.Rows(rowIndex).Cells(7).Value = Bee.CurrentRow.Cells(0).Value
        FG.Rows(rowIndex).Cells(8).Value = Bee.CurrentRow.Cells(1).Value
        FG.Rows(rowIndex).Cells(9).Value = Bee.CurrentRow.Cells(2).Value
        Panel1.Visible = False
        loadColor()
        FG.Focus()
    End Sub

    Private Sub Bee_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles Bee.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim rowIndex As Integer = FG.CurrentCell.RowIndex
            FG.Rows(rowIndex).Cells(7).Value = Bee.CurrentRow.Cells(0).Value
            FG.Rows(rowIndex).Cells(8).Value = Bee.CurrentRow.Cells(1).Value
            FG.Rows(rowIndex).Cells(9).Value = Bee.CurrentRow.Cells(2).Value
            Panel1.Visible = False
            loadColor()
            FG.Focus()
            e.Handled = True
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
        If FG.CurrentCell Is Nothing Then Exit Sub
        Dim FgR As Integer = FG.CurrentCell.RowIndex
        Dim FgC As Integer = FG.CurrentCell.ColumnIndex

        ' FG.Redraw = False ' Not needed for DGV or use SuspendLayout()

        For J As Integer = 0 To FG.Rows.Count - 1
            Dim row As DataGridViewRow = FG.Rows(J)
            
            ' Apply numbering
            row.Cells(0).Value = J + 1

            Dim cell1Val As String = GetString(row.Cells(1).Value)
            If cell1Val.Trim() <> "" Then
                row.Cells(1).Style.BackColor = Color.LightCyan
                row.Cells(1).Style.Font = New Font(FG.Font, FontStyle.Bold)
                
                row.Cells(5).Style.BackColor = Color.LightCyan
                row.Cells(5).Style.Font = New Font(FG.Font, FontStyle.Bold)
            Else
                row.Cells(1).Style.BackColor = Color.White
            End If

            Dim cell2Val As String = GetString(row.Cells(2).Value)
            If cell2Val.Trim() <> "" Then
                row.Cells(2).Style.BackColor = Color.LightCyan
                row.Cells(2).Style.Font = New Font(FG.Font, FontStyle.Bold)
                
                row.Cells(6).Style.BackColor = Color.LightCyan
                row.Cells(6).Style.Font = New Font(FG.Font, FontStyle.Bold)
            Else
                row.Cells(2).Style.BackColor = Color.White
            End If

            If J = 0 Then
                 ' Special case if needed, originally code had J=1 to Rows-1
            End If
        Next J

        If FgR >= 0 And FgR < FG.Rows.Count And FgC >= 0 And FgC < FG.Columns.Count Then
            FG.CurrentCell = FG.Rows(FgR).Cells(FgC)
        End If
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

        For i = 0 To FG.Rows.Count - 1
            'If FG.Row > 1 Then
            FG.Rows(i).Cells(10).Value = Cmb.Text
            FG.Rows(i).Cells(11).Value = Format(CDbl(txtRate.Text), "#,##0.00")
            FG.Rows(i).Cells(12).Value = Format(CDbl(CDbl(txtRate.Text) * GetValue(FG.Rows(i).Cells(5).Value)), "#,##0.00")
            FG.Rows(i).Cells(13).Value = Format(CDbl(CDbl(txtRate.Text) * GetValue(FG.Rows(i).Cells(6).Value)), "#,##0.00")
            'End If

        Next
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

        ' Legacy loop removed

        Dim k As Integer
        For k = 0 To FG.Rows.Count - 1
            Dim cell1v As String = GetString(FG.Rows(k).Cells(1).Value)
            Dim cell2v As String = GetString(FG.Rows(k).Cells(2).Value)
            Dim cell5v As Double = GetValue(FG.Rows(k).Cells(5).Value)
            Dim cell6v As Double = GetValue(FG.Rows(k).Cells(6).Value)

            If cell1v <> "" Then
                If cell5v = 0 Then
                    MsgBox("ກະລຸນນາໃສ່ມູນຄ່າກ່ອນ")
                    Exit Sub
                End If
            End If
            If cell2v <> "" Then
                If cell6v = 0 Then
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


        'Call LoadSqlData("SELECT  top 1 Right(certify,3) As  certify    FROM AP_ACC_adjust_Item WHERE   book ='" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "' And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And  Month(date_work)=" & Format(CDate(dtActi.Value), "MM") & " " & _
        '                 "  and LEFT(company,2)='" & Off_Id & "' Order by  Right(certify,3) DESC ", RSC)
        Call LoadSqlData("SELECT  top 1 Right(certify,3) As  certify    FROM AP_ACC_adjust_Item WHERE   book =N'" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "' And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And  Month(date_work)=" & Format(CDate(dtActi.Value), "MM") & "  And  day(date_work)=" & Format(CDate(dtActi.Value), "dd") & " " & _
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


        If FG.Rows.Count > 0 AndAlso GetString(FG.Rows(0).Cells(6).Value) = "" Then MsgBox("ກະລຸນນາລົງບັນຊີເງິນກອ່ນ!", MsgBoxStyle.OkOnly) : Exit Sub
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

            'Call LoadSqlData("SELECT  top 1 Right(certify,3) As  certify    FROM AP_ACC_adjust_Item WHERE   book ='" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "' And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And  month(date_work)=" & Format(CDate(dtActi.Value), "MM") & "  Order by  Right(certify,3) DESC ", RSC)
            Call LoadSqlData("SELECT  top 1 Right(certify,3) As  certify    FROM AP_ACC_adjust_Item WHERE   book =N'" & CmbBook.Text & "' And  certify = N'" & txtInvoice.Text & "' And  year(date_work)=" & Format(CDate(dtActi.Value), "yyyy") & " And  month(date_work)=" & Format(CDate(dtActi.Value), "MM") & " And  day(date_work)=" & Format(CDate(dtActi.Value), "dd") & "  Order by  Right(certify,3) DESC ", RSC)

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
            'CNN.Execute("DELETE FROM AP_ACC_adjust_Item WHERE book =N'" & FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 16) & "' And certify  = '" & txtInvoice.Text & "'   And  year(date_work)='" & Format(CDate(FmJeneralJournal_List.FG.get_TextMatrix(FmJeneralJournal_List.FG.Row, 1)), "yyyy") & "' ")
            CNN.Execute("DELETE FROM AP_ACC_adjust_Item WHERE book =N'" & GetString(FmJeneralJournal_Adjust_List.FG.CurrentRow.Cells(16).Value) & "' And certify  =N'" & txtInvoice.Text & "'   And   date_work='" & Format(CDate(GetString(FmJeneralJournal_Adjust_List.FG.CurrentRow.Cells(1).Value)), "yyyy-MM-dd") & "' ")

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
            FG.Rows.Clear()
            'FG.Rows = 2 ' Removed legacy row count setting
        If CheckBox1.Checked = True Then
            Call LoadReport()
        End If
        If CheckBox2.Checked = True Then
            Close()
        End If
        'End If


    End Sub

    Private Sub BtnAddNew2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew2.Click
            FG.Rows.Clear()
            'FG.Rows = 2 ' Removed legacy row count setting
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
            e.Cancel = True
            Exit Sub
        End If
        ' Note: CellBackColor and CellForeColor are not directly available in DataGridView
        ' You would need to set the style for the specific cell if needed
        BtnMove.Visible = False
        BtnSearch.Visible = False


    End Sub

    Private Sub Cat()
        If FG.CurrentCell Is Nothing Then Exit Sub
        Dim r As Integer = FG.CurrentCell.RowIndex
        Dim c As Integer = FG.CurrentCell.ColumnIndex
        
        If c = 7 Then
            Dim valAsStr As String = GetString(FG.Rows(r).Cells(7).Value)
            If valAsStr = "0" Then
                FG.Rows(r).Cells(8).Value = "ບໍ່ເລືອກ"
                FG.Rows(r).Cells(9).Value = "No Selete"
            ElseIf valAsStr = "1" Then
                FG.Rows(r).Cells(8).Value = "ຮັບໃຊ້ການພະລິດ"
                FG.Rows(r).Cells(9).Value = "Use build"

            ElseIf valAsStr = "2" Then
                FG.Rows(r).Cells(8).Value = "ຮັບໃຊ້ການຈຳໜ່າຍ"
                FG.Rows(r).Cells(9).Value = "Use Sell"
            ElseIf valAsStr = "3" Then
                FG.Rows(r).Cells(8).Value = "ຮັບໃຊ້ບໍລິຫານ"
                FG.Rows(r).Cells(9).Value = "Use manage "

            ElseIf valAsStr = "4" Then
                FG.Rows(r).Cells(8).Value = "ຕົ້ນທຶນຂາຍ/ຕົ້ນທຶນບໍລິຫານ"
                FG.Rows(r).Cells(9).Value = "Sell capital/manage capital "
            ElseIf IsNumeric(valAsStr) AndAlso CInt(valAsStr) > 4 Then
                MessageBox.Show("ລະຫັດບໍ່ຖືກຕ້ອງ ລະຫັດມີພຽງເລກ 0 ຫາເລກ 4 ເທົ່ານັ້ນ")
                FG.Rows(r).Cells(8).Value = "ບໍ່ລເລືອກ"
                FG.Rows(r).Cells(9).Value = "No Selete"
                FG.Rows(r).Cells(7).Value = "0"
                Exit Sub
            ElseIf IsNumeric(valAsStr) AndAlso CInt(valAsStr) < 0 Then
                MessageBox.Show("ລະຫັດບໍ່ຖືກຕ້ອງ ລະຫັດມີພຽງເລກ 0 ຫາເລກ 4 ເທົ່ານັ້ນ")
                FG.Rows(r).Cells(8).Value = "ບໍ່ລເລືອກ"
                FG.Rows(r).Cells(9).Value = "No Selete"
                FG.Rows(r).Cells(7).Value = "0"
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
        FG.Rows.Clear()
        With RSC
            Call LoadSqlData("select *  from AP_ACC_adjust_Item WHERE book =N'" & GetString(FmJeneralJournal_Adjust_List.FG.CurrentRow.Cells(16).Value) & "' And certify  =N'" & txtInvoice.Text & "'   And  year(date_work)='" & Format(CDate(GetString(FmJeneralJournal_Adjust_List.FG.CurrentRow.Cells(1).Value)), "yyyy") & "' order by cnt", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    If CDbl(.Fields("AG").Value) = 1 Then
                        FG.Rows.Add(.AbsolutePosition, Trim(CStr(.Fields("code_dr").Value)), Trim(CStr(.Fields("code_cr").Value)), Trim(CStr(.Fields("descrip").Value.ToString)), Trim(CStr(.Fields("descripe").Value)), Format(CDbl(.Fields("amt_dr").Value), "##,##0.00"), Format(CDbl(.Fields("amt_cr").Value), "##,##0.00"), Trim(CStr(.Fields("Cat_ID").Value)))
                    Else
                        FG.Rows.Add(.AbsolutePosition, Trim(CStr(.Fields("code_dr").Value)), Trim(CStr(.Fields("code_cr").Value)), Trim(CStr(.Fields("descrip").Value.ToString)), Trim(CStr(.Fields("descripe").Value)), Format(CDbl(.Fields("amount_dr").Value), "##,##0.00"), Format(CDbl(.Fields("amount_cr").Value), "##,##0.00"), Trim(CStr(.Fields("Cat_ID").Value)))
                    End If
                    .MoveNext()
                End While
            Else
                For k As Integer = 1 To 16
                    FG.Rows.Add()
                Next
            End If
            '==
            Dim prevRowIndex As Integer = FmJeneralJournal_Adjust_List.FG.CurrentCell.RowIndex
            Dim PP As String = "select top 1 *  from AP_ACC_adjust_Item WHERE book =N'" & FmJeneralJournal_Adjust_List.FG.Rows(prevRowIndex).Cells(16).Value.ToString() & "' And certify  =N'" & txtInvoice.Text & "'   And   date_work='" & Format(CDate(FmJeneralJournal_Adjust_List.FG.Rows(prevRowIndex).Cells(0).Value), "yyyy-MM-dd") & "' order by cnt"
            Call LoadSqlData(PP, RSC)

            If .RecordCount > 0 Then
                Book = Trim(.Fields("book").Value)
                dtActi.Text = Format(CDate(Trim(.Fields("date_work").Value)), "dd/MM/yyyy")
                txtAmount.Text = Format(CDbl(Trim(.Fields("amount").Value)), "#,##0.00")
                txtDesc.Text = Trim(.Fields("descrip").Value)
                txtDescE.Text = Trim(.Fields("descripe").Value)
                Rate1 = Format(CDbl(Trim(.Fields("rate").Value)), "#,##0.00")
                Cmb.Text = Trim(.Fields("curr").Value)
                CmbBook.Text = Book
                txtInvoice.Text = MDInvoiceNo

                If AG = 1 Then
                    CheckBox3.Checked = True
                Else
                    CheckBox3.Checked = False
                End If

            End If
            txtRate.Text = Rate1

            Dim i As Integer
            For i = 0 To FG.Rows.Count - 1
                '===============
                FG.Rows(i).Cells(10).Value = Cmb.Text
                FG.Rows(i).Cells(11).Value = Format(CDbl(txtRate.Text), "#,##0.00")
                
                Dim cell5Val As Object = FG.Rows(i).Cells(5).Value
                Dim cell6Val As Object = FG.Rows(i).Cells(6).Value
                
                FG.Rows(i).Cells(12).Value = Format(CDbl(CDbl(txtRate.Text) * CDbl(If(cell5Val Is Nothing OrElse cell5Val.ToString() = "", 0, cell5Val))), "#,##0.00")
                FG.Rows(i).Cells(13).Value = Format(CDbl(CDbl(txtRate.Text) * CDbl(If(cell6Val Is Nothing OrElse cell6Val.ToString() = "", 0, cell6Val))), "#,##0.00")

                Dim cell7Val As Object = FG.Rows(i).Cells(7).Value
                If cell7Val IsNot Nothing AndAlso cell7Val.ToString() = "0" Then
                    FG.Rows(i).Cells(8).Value = "ບໍ່ເລືອກ"
                    FG.Rows(i).Cells(9).Value = "No Selete"
                ElseIf cell7Val IsNot Nothing AndAlso cell7Val.ToString() = "1" Then
                    FG.Rows(i).Cells(8).Value = "ຮັບໃຊ້ການພະລິດ"
                    FG.Rows(i).Cells(9).Value = "Use build"
                ElseIf cell7Val IsNot Nothing AndAlso cell7Val.ToString() = "2" Then
                    FG.Rows(i).Cells(8).Value = "ຮັບໃຊ້ການຈຳໜ່າຍ"
                    FG.Rows(i).Cells(9).Value = "Use Sell"
                ElseIf cell7Val IsNot Nothing AndAlso cell7Val.ToString() = "3" Then
                    FG.Rows(i).Cells(8).Value = "ຮັບໃຊ້ບໍລິຫານ"
                    FG.Rows(i).Cells(9).Value = "Use manage"
                ElseIf cell7Val IsNot Nothing AndAlso cell7Val.ToString() = "4" Then
                    FG.Rows(i).Cells(8).Value = "ຕົ້ນທຶນຂາຍ/ຕົ້ນທຶນບໍລິຫານ"
                    FG.Rows(i).Cells(9).Value = "capital/manage capital"
                End If
            Next i
        End With

        For i = 0 To FG.Rows.Count - 1
            Dim cell1Val As Object = FG.Rows(i).Cells(1).Value
            Dim cell2Val As Object = FG.Rows(i).Cells(2).Value
            Dim prevRowIndexFmList As Integer = FmJeneralJournal_Adjust_List.FG.CurrentCell.RowIndex
            Call LoadSqlData("select *  from AP_ACC_adjust_Item WHERE Ac_Code='" & If(cell1Val Is Nothing, "", cell1Val.ToString()) & If(cell2Val Is Nothing, "", cell2Val.ToString()) & "' and book ='" & FmJeneralJournal_Adjust_List.FG.Rows(prevRowIndexFmList).Cells(16).Value.ToString() & "' And certify  = '" & txtInvoice.Text & "'   And  year(date_work)='" & Format(CDate(FmJeneralJournal_Adjust_List.FG.Rows(prevRowIndexFmList).Cells(0).Value), "yyyy") & "' order by cnt", RSC)
            If RSC.RecordCount > 0 Then
                FG.Rows(i).Cells(3).Value = Trim(RSC.Fields("ac_Name").Value.ToString())
                'FG.Rows(i).Cells(4).Value = Trim(RSC.Fields("ac_Namee").Value.ToString())
            End If
        Next i

        SumAmountDr()
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ShowList()
    End Sub
    Private Sub ShowList()
        FG.Columns(7).Visible = False
        FG.Columns(8).Visible = False
        If Button3.Text = "Show All" Then


            FG.Columns(4).Visible = False ' Original was True, but the provided snippet has False for this case. Sticking to the snippet.
            FG.Columns(9).Visible = True
            FG.Columns(10).Visible = True
            FG.Columns(11).Visible = True
            FG.Columns(12).Visible = True
            FG.Columns(13).Visible = True
            FG.Columns(14).Visible = False
            FG.Columns(15).Visible = False
            Button3.Text = "Show GLN"
            Exit Sub
        End If
        If Button3.Text = "Show GLN" Then
            If MuLng = "E" Then
                FG.Columns(3).Visible = True
                FG.Columns(4).Visible = False
            Else
                FG.Columns(3).Visible = False
                FG.Columns(4).Visible = True
            End If
            'FG.set_ColHidden(4, True)
            FG.Columns(9).Visible = True
            FG.Columns(10).Visible = True
            FG.Columns(11).Visible = True
            FG.Columns(12).Visible = True
            FG.Columns(13).Visible = True
            FG.Columns(14).Visible = True
            FG.Columns(15).Visible = True
            Button3.Text = "Show All"
            Exit Sub
        End If
    End Sub


    Private Sub ChCat_ID_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChCat_ID.CheckedChanged
        For J = 0 To FG.Rows.Count - 1
            Dim cell3Val As Object = FG.Rows(J).Cells(3).Value
            If cell3Val IsNot Nothing AndAlso cell3Val.ToString() <> "" Then
                If ChCat_ID.Checked = True Then
                    ' Simulation of highlighting/selecting col 7
                    FG.Rows(J).Cells(7).Style.BackColor = Color.LightCyan
                Else
                    FG.Rows(J).Cells(7).Style.BackColor = Color.White
                    FG.Rows(J).Cells(7).Style.ForeColor = Color.Gray
                    FG.Rows(J).Cells(8).Style.ForeColor = Color.Gray
                    FG.Rows(J).Cells(9).Style.ForeColor = Color.Gray

                    FG.Rows(J).Cells(7).Value = "0"
                    FG.Rows(J).Cells(8).Value = "ບໍ່ເລືອກ"
                    FG.Rows(J).Cells(9).Value = "No Selete"
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
        LngId = "7038" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As  Crl_Ac_Name	 ,"
        If MuLng = "L" Then
            MuLngRpt = MuLngRpt & "N'" & Amount_In_Word & "' As Crl_Amt_In_Word	 ,"
        Else
            MuLngRpt = MuLngRpt & "N'" & Amount_In_Word & "' As Crl_Amt_In_Word	 ,"
        End If


        SLF = MuLngRpt & " AP_ACC_adjust_Item.company ,AP_ACC_adjust_Item.Date_Work , AP_ACC_adjust_Item.certify, AP_ACC_adjust_Item.ac_code, AP_ACC_adjust_Item.descrip , AP_ACC_adjust_Item.descripe , AP_ACC_adjust_Item.amt_dr, AP_ACC_adjust_Item.amt_cr, AP_ACC_adjust_Item.ac_name  AS Name_L , AP_ACC_adjust_Item.ac_namee AS Name_E  "

        'SLF = MuLngRpt & " AP_ACC_adjust_Item.company ,AP_ACC_adjust_Item.Date_Work , AP_ACC_adjust_Item.certify, AP_ACC_adjust_Item.ac_code, AP_ACC_adjust_Item.descrip , AP_ACC_adjust_Item.descripe , AP_ACC_adjust_Item.amt_dr, AP_ACC_adjust_Item.amt_cr, Acc_Code.Name_L AS Name_L , Acc_Code.Name_E AS Name_E  "
        'SLF = RptSjOff & "    AP_ACC_adjust_Item.certify, AP_ACC_adjust_Item.ac_code, AP_ACC_adjust_Item.descrip, AP_ACC_adjust_Item.amt_dr, AP_ACC_adjust_Item.amt_cr, Acc_Code.Name_L AS Name_L , "
        Call LoadLoGO()
        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            .Open("SELECT   " & SLF & "  FROM         AP_ACC_adjust_Item INNER JOIN    Acc_Code ON AP_ACC_adjust_Item.ac_code = Acc_Code.Ac_Code WHERE AP_ACC_adjust_Item.certify = N'" & MdCertifyId2 & "' And  year(date_work)=" & Format(dtActi.Value, "yyyy") & "  order by AP_ACC_adjust_Item.cnt", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
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
        For i = 0 To FG.Rows.Count - 1
            FG.Rows(i).Cells(10).Value = Cmb.Text
            FG.Rows(i).Cells(11).Value = Format(CDbl(txtRate.Text), "#,##0.00")
            FG.Rows(i).Cells(12).Value = Format(CDbl(CDbl(txtRate.Text) * GetValue(FG.Rows(i).Cells(5).Value)), "#,##0.00")
            FG.Rows(i).Cells(13).Value = Format(CDbl(CDbl(txtRate.Text) * GetValue(FG.Rows(i).Cells(6).Value)), "#,##0.00")
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

    
    Public Sub AddAcc2()
        ' TODO: Implement AddAcc2 logic (originally used to add account from chart)
    End Sub
End Class