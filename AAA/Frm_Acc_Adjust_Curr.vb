Public Class Frm_Acc_Adjust_Curr
    Dim Amount As Double
    Dim Amt As Double
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

    Private Sub Frm_Acc_Adjust_Curr_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'FG.FormatString = "^ລ/ດ|<ເລກບັນຊີໜີ້         |<ເລກບັນຊີມີ         |ຊື່ບັນຊີພາສາລາວ                                |>ຈໍານວນເງິນຈົດໜີ້       |>ຈໍານວນເງິນຈົດມີ      |<ກິດຈະກຳ|<ປະເພດລາຍຈ່າຍ|^ສະກຸນເງິນ|^ອັດຕາແລກປ່ຽນ |>ມູນຄ່າເປັນກີບ ໜີ້   |>ມູນຄ່າເປັນກີບ ມີ   |>ມູນຄ່າເປັນ(USD) ໜີ້   |>ມູນຄ່າເປັນ(USD) ມີ  |>ເອກກະສານອ້າງອິງ"
        FG.FormatString = "^ລ/ດ|<ເລກບັນຊີໜີ້         |<ເລກບັນຊີມີ         |<ຊື່ບັນຊີພາສາລາວ                   |<ຊື່ບັນຊີພາສາອັງກິດ    |>ຈໍານວນເງິນຈົດໜີ້       |>ຈໍານວນເງິນຈົດມີ     |^ສະກຸນເງິນ|^ອັດຕາແລກປ່ຽນ |>ມູນຄ່າເປັນກີບ ໜີ້           |>ມູນຄ່າເປັນກີບ ມີ       |> ມູນຄ່າເປັນ(USD) ໜີ້     |>ມູນຄ່າເປັນ(USD) ມີ        "
        FG.set_ColHidden(5, True)
        FG.set_ColHidden(6, True)
        FG.set_ColHidden(7, True)
        FG.set_ColHidden(8, True)
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
        FG.Rows = 1
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
        '''''F"^ລ/ດ|<ເລກບັນຊີໜີ້         |<ເລກບັນຊີມີ         |ຊື່ບັນຊີພາສາລາວ                                |>ຈໍານວນເງິນຈົດໜີ້       |>ຈໍານວນເງິນຈົດມີ     |^ສະກຸນເງິນ|^ອັດຕາແລກປ່ຽນ |>ມູນຄ່າເປັນກີບ ໜີ້   |>ມູນຄ່າເປັນກີບ ມີ        "
        FG.Rows = 1
        With rs
            Call LoadSqlData("  SELECT * from AP_ACC_adjust_Item WHERE AP_ACC_adjust_Item.certify='" & SaleID & "' ", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    FG.AddItem(.AbsolutePosition & _
                    Chr(9) & (.Fields("code_dr").Value.ToString) & _
                    Chr(9) & Format(.Fields("code_cr").Value.ToString) & _
                    Chr(9) & Format(.Fields("ac_name").Value.ToString) & _
                    Chr(9) & Format(CDbl(.Fields("amount_dr").Value), "#,##0.00") & _
                    Chr(9) & Format(CDbl(.Fields("amount_Cr").Value), "#,##0.00") & _
                    Chr(9) & Format(.Fields("Curr_i").Value.ToString) & _
                    Chr(9) & Format(CDbl(.Fields("Rate_i").Value), "#,##0.00") & _
                    Chr(9) & Format(CDbl(.Fields("amt_dr").Value), "#,##0.00") & _
                         Chr(9) & Format(CDbl(.Fields("amt_cr").Value), "#,##0.00"))
                    .MoveNext()
                End While
                'Fg.Rows = Fg.Rows + 1
            End If
        End With

    End Sub

    Private Sub FG_AfterEdit(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_AfterEditEvent) Handles FG.AfterEdit

        'MsgBox(FG.get_TextMatrix(FG.Rows - 1, 3))
        AccId = ""
        AccName = ""
        If FG.get_TextMatrix(FG.Rows - 1, 3) <> "" Then
            FG.Rows = FG.Rows + 1
            Exit Sub
        End If

        If FG.Col = 1 Then
            R = FG.Row()
            L = FG.Col
            If FG.get_TextMatrix(FG.Row, 1) <> "" Then

                AccId = FG.get_TextMatrix(FG.Row, 1)
                LoadText()
                FG.set_TextMatrix(FG.Row, 3, AccName)
                Call load_Ac()
                If FG.get_TextMatrix(FG.Row, 3) = "" Then
                    FG.set_TextMatrix(FG.Row, 1, "")
                    MD_FROM = Me.Name
                    If AccId <> "" Then
                        MD_KH = AccId
                    End If
                    fmShartOfAccDetail.ShowDialog()
                End If
            End If
        End If

        If FG.Col = 2 Then
            R = FG.Row()
            L = FG.Col
            If FG.get_TextMatrix(FG.Row, 2) <> "" Then

                AccId = FG.get_TextMatrix(FG.Row, 2)
                LoadText()
                FG.set_TextMatrix(FG.Row, 3, AccName)
                FG.set_TextMatrix(FG.Row, 4, AccNamee)
                If L = "1" Then
                    Call load_Ac()
                Else
                    Call load_Ac2()
                End If

                If FG.get_TextMatrix(FG.Row, 3) = "" Then
                    FG.set_TextMatrix(FG.Row, 1, "")
                    MD_FROM = Me.Name
                    If AccId <> "" Then
                        MD_KH = AccId
                    End If
                    fmShartOfAccDetail.ShowDialog()
                End If
            End If
        End If
        Call Load_Calculate()

        'If FG.Col = 6 Then
        '    If FG.get_TextMatrix(FG.Row, 6) <> "" Then
        '        Cat_id = ""
        '        Call LoadSqlData("SELECT * FROM AP_Activity WHERE Act_id = N'" & FG.get_TextMatrix(FG.Row, 6) & "'", RSC)
        '        With RSC
        '            If .RecordCount > 0 Then
        '                Do Until .EOF = True
        '                    Cat_id = Trim(.Fields("Act_id").Value)
        '                    .MoveNext()
        '                Loop
        '            Else
        '                'Frm_Activity_code_Item.ShowDialog()
        '            End If
        '        End With
        '        FG.set_TextMatrix(FG.Row, 6, Cat_id)
        '    End If
        'End If

        'If FG.Col = 7 Then
        '    If FG.get_TextMatrix(FG.Row, 7) <> "" Then
        '        Cat_id = ""
        '        Call LoadSqlData("SELECT * FROM AP_List_of_Category WHERE cat_id = N'" & FG.get_TextMatrix(FG.Row, 7) & "'", RSC)
        '        With RSC
        '            If .RecordCount > 0 Then
        '                Do Until .EOF = True
        '                    Cat_id = Trim(.Fields("cat_id").Value)
        '                    .MoveNext()
        '                Loop
        '            Else
        '                'Frm_Of_Catecory_item.ShowDialog()
        '            End If
        '        End With
        '        FG.set_TextMatrix(FG.Row, 7, Cat_id)
        '    End If
        'End If

    End Sub
    Private Sub load_Ac()
        If L = "1" Then
            If R = "1" Then
                MD_AMTDR = Format(CDbl(txtDiff.Text), "#,##0.00")
                MD_AMTCR = 0
                MD_AMTDR_LAK = Format(CDbl(txtDiff.Text), "#,##0.00")
                MD_AMTCR_LAK = 0
            Else
                MD_AMTDR = Format(CDbl(txtDiff.Text), "#,##0.00")
                MD_AMTCR = 0
                MD_AMTDR_LAK = Format(CDbl(txtDiff.Text), "#,##0.00")
                MD_AMTCR_LAK = 0

            End If
        Else
            If R = "1" Then
                MD_AMTCR = Format(CDbl(txtDiff.Text), "#,##0.00")
                MD_AMTDR = 0
                MD_AMTCR_LAK = Format(CDbl(txtDiff.Text), "#,##0.00")
                MD_AMTDR_LAK = 0
            Else
                MD_AMTCR = CDbl(txtSumAmountDr.Text - txtSumAmountCr.Text)
                MD_AMTDR = 0
                MD_AMTCR_LAK = Format(CDbl(txtSumTotalAmountDr.Text - txtSumTotalAmountCr.Text), "#,##0.00")
                MD_AMTDR_LAK = 0
            End If
        End If

        If MD_AMTDR < 0 Then
            MD_AMTDR = CDbl(MD_AMTDR * -1)
        End If
        If MD_AMTCR < 0 Then
            MD_AMTCR = CDbl(MD_AMTCR * -1)
        End If
        '  FG.FormatString = "^ລ/ດ|<ເລກບັນຊີໜີ້         |<ເລກບັນຊີມີ         |ຊື່ບັນຊີພາສາລາວ                                |>ຈໍານວນເງິນຈົດໜີ້       |>ຈໍານວນເງິນຈົດມີ"
        '     |<ກິດຈະກຳ|<ປະເພດລາຍຈ່າຍ|^ສະກຸນເງິນ|^ອັດຕາແລກປ່ຽນ |>ມູນຄ່າເປັນກີບ ໜີ້   |>ມູນຄ່າເປັນກີບ ມີ   |>ມູນຄ່າເປັນ(USD) ໜີ້   |>ມູນຄ່າເປັນ(USD) ມີ        "

        FG.set_TextMatrix(R, L, FG.get_TextMatrix(FG.Row, 1))
        FG.set_TextMatrix(R, 5, Format(CDbl(MD_AMTDR), "#,##0.00"))
        FG.set_TextMatrix(R, 6, Format(CDbl(MD_AMTCR), "#,##0.00"))
        FG.set_TextMatrix(R, 7, "LAK")
        FG.set_TextMatrix(R, 8, 1)
        FG.set_TextMatrix(R, 9, Format(CDbl(MD_AMTDR), "#,##0.00"))
        FG.set_TextMatrix(R, 10, Format(CDbl(MD_AMTCR), "#,##0.00"))
        If FG.get_TextMatrix(FG.Rows - 1, 3) <> "" Then
            FG.Rows = FG.Rows + 1
        End If
    End Sub
    Private Sub load_Ac2()
        If L = "1" Then
            If R = "1" Then
                MD_AMTDR = Format(CDbl(txtDiff.Text), "#,##0.00")
                MD_AMTCR = 0
                MD_AMTDR_LAK = Format(CDbl(txtDiff.Text), "#,##0.00")
                MD_AMTCR_LAK = 0
            Else
                MD_AMTDR = CDbl(txtSumAmountDr.Text - txtSumAmountCr.Text)
                MD_AMTCR = 0
                MD_AMTDR_LAK = Format(CDbl(txtSumTotalAmountDr.Text - txtSumTotalAmountCr.Text), "#,##0.00")
                MD_AMTCR_LAK = 0
            End If
        Else
            If R = "1" Then
                MD_AMTCR = Format(CDbl(txtDiff.Text), "#,##0.00")
                MD_AMTDR = 0
                MD_AMTCR_LAK = Format(CDbl(txtDiff.Text), "#,##0.00")
                MD_AMTDR_LAK = 0
            Else
                'MD_AMTCR = CDbl(txtSumAmountDr.Text - txtSumAmountCr.Text)
                MD_AMTCR = Format(CDbl(txtDiff.Text), "#,##0.00")
                MD_AMTDR = 0
                MD_AMTCR_LAK = Format(CDbl(txtDiff.Text), "#,##0.00")
                MD_AMTDR_LAK = 0
            End If
        End If

        If MD_AMTDR < 0 Then
            MD_AMTDR = CDbl(MD_AMTDR * -1)
        End If
        If MD_AMTCR < 0 Then
            MD_AMTCR = CDbl(MD_AMTCR * -1)
        End If

        FG.set_TextMatrix(R, 5, Format(CDbl(MD_AMTDR), "#,##0.00"))
        FG.set_TextMatrix(R, 6, Format(CDbl(MD_AMTCR), "#,##0.00"))
        FG.set_TextMatrix(R, 7, "LAK")
        FG.set_TextMatrix(R, 8, 1)
        FG.set_TextMatrix(R, 9, Format(CDbl(MD_AMTDR), "#,##0.00"))
        FG.set_TextMatrix(R, 10, Format(CDbl(MD_AMTCR), "#,##0.00"))
        If FG.get_TextMatrix(FG.Rows - 1, 3) <> "" Then
            FG.Rows = FG.Rows + 1
        End If
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
    Private Sub FG_DblClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles FG.DblClick
        If FG.Row = FG.Rows - 1 Then Exit Sub
        AccCD = FG.get_TextMatrix(FG.Row, 1) & FG.get_TextMatrix(FG.Row, 2)
        If MessageBox.Show("ທ່ານຕ້ອງການລືບລາຍການ'" & AccCD & "' ນີ້ ແທ້ ຫຼື ບໍ່ ?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            FG.RemoveItem()
            Call Calculate()
        End If

    End Sub

    Private Sub FG_KeyUpEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_KeyUpEvent) Handles FG.KeyUpEvent
        If e.keyCode = 13 Then
            If FG.Col = 4 Then
                FG.set_TextMatrix(FG.Row, 9, Format(CDbl(FG.get_TextMatrix(FG.Row, 5)), "#,##0.00"))

            End If
            If FG.Col = 5 Then
                FG.set_TextMatrix(FG.Row, 10, Format(CDbl(FG.get_TextMatrix(FG.Row, 6)), "#,##0.00"))
            End If

            FG.set_TextMatrix(FG.Row, 5, Format(CDbl(FG.get_TextMatrix(FG.Row, 5)), "#,##0.00"))
            FG.set_TextMatrix(FG.Row, 6, Format(CDbl(FG.get_TextMatrix(FG.Row, 6)), "#,##0.00"))

            If CMB_Curr.Text = "USD" Then
                FG.set_TextMatrix(FG.Row, 7, "USD")
                FG.set_TextMatrix(FG.Row, 9, Format(CDbl(FG.get_TextMatrix(FG.Row, 5)), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 10, Format(CDbl(FG.get_TextMatrix(FG.Row, 6)), "#,##0.00"))
            ElseIf CMB_Curr.Text = "THB" Then
                FG.set_TextMatrix(FG.Row, 7, "THB")
                If FG.get_TextMatrix(FG.Row, 5) > 0 Then
                    MD_AMTDR = FG.get_TextMatrix(FG.Row, 5)
                Else
                    MD_AMTCR = FG.get_TextMatrix(FG.Row, 6)
                End If
            Else
                FG.set_TextMatrix(FG.Row, 7, "LAK")
                FG.set_TextMatrix(FG.Row, 9, Format(CDbl(FG.get_TextMatrix(FG.Row, 5)), "#,##0.00"))
                FG.set_TextMatrix(FG.Row, 10, Format(CDbl(FG.get_TextMatrix(FG.Row, 6)), "#,##0.00"))
            End If

            If txt_descrip.Text = "" Then
                txt_descrip.Text = FG.get_TextMatrix(1, 3)
            End If
        End If
        Call Calculate()

    End Sub
    Public Sub LoadColor()
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
                FG.Col = 4
                FG.CellBackColor = Color.LightCyan
                FG.CellFontBold = True
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
                    FG.Col = 5
                    FG.CellFontBold = True
                    FG.CellBackColor = Color.LightCyan
                    FG.set_TextMatrix(J, 0, J)
                End If
                If Trim(FG.get_TextMatrix(J, 2)) = "" Then
                    FG.Col = 2
                    FG.CellBackColor = Color.White

                End If
            End If
        Next J


        FG.Row = FgR
        FG.Col = FgC
        FG.Redraw = True
    End Sub
    Private Sub Load_Calculate()
        For i = 1 To FG.Rows - 1
            'FG.set_TextMatrix(FG.Row, 8, Format((FG.get_TextMatrix(i, 4)) * CDbl(FG.get_TextMatrix(i, 7)), "#,##0.00"))
            'FG.set_TextMatrix(FG.Row, 9, Format((FG.get_TextMatrix(i, 5)) * CDbl(FG.get_TextMatrix(i, 7)), "#,##0.00"))
            Call Calculate()
        Next i
    End Sub
    Private Sub Calculate()

        Dim i As Integer
        Dim amt1 As Double
        Dim amt2 As Double
        Dim amt3 As Double
        Dim amt4 As Double
        amt1 = 0
        amt2 = 0
        amt3 = 0
        amt4 = 0
        '  FG.FormatString = "^ລ/ດ|<ເລກບັນຊີໜີ້         |<ເລກບັນຊີມີ         |ຊື່ບັນຊີພາສາລາວ                                |>ຈໍານວນເງິນຈົດໜີ້       |>ຈໍານວນເງິນຈົດມີ"
        '     |<ກິດຈະກຳ|<ປະເພດລາຍຈ່າຍ|^ສະກຸນເງິນ|^ອັດຕາແລກປ່ຽນ |>ມູນຄ່າເປັນກີບ ໜີ້   |>ມູນຄ່າເປັນກີບ ມີ   |>ມູນຄ່າເປັນ(USD) ໜີ້   |>ມູນຄ່າເປັນ(USD) ມີ        "

        For i = 1 To FG.Rows - 1
            amt1 = amt1 + CDbl(FG.get_TextMatrix(i, 5))
            amt2 = amt2 + CDbl(FG.get_TextMatrix(i, 6))
            amt3 = amt3 + CDbl(FG.get_TextMatrix(i, 9))
            amt4 = amt4 + CDbl(FG.get_TextMatrix(i, 10))
        Next i
        txtSumAmountDr.Text = Format(amt1, "#,##0.00")
        txtSumAmountCr.Text = Format(amt2, "#,##0.00")
        txtSumTotalAmountDr.Text = Format(amt3, "#,##0.00")
        txtSumTotalAmountCr.Text = Format(amt4, "#,##0.00")

        If txtSumAmountDr.Text > txtSumAmountCr.Text Then
            Dr.Text = Format(CDbl(txtSumAmountCr.Text - txtSumAmountDr.Text), "#,##0.00")
            Cr.Text = "0.00"

            DDR.Text = Format(CDbl(txtSumTotalAmountCr.Text - txtSumTotalAmountDr.Text), "#,##0.00")
            CCR.Text = "0.00"
        End If

        If txtSumAmountCr.Text > txtSumAmountDr.Text Then
            Cr.Text = Format(CDbl(txtSumAmountDr.Text - txtSumAmountCr.Text), "#,##0.00")
            Dr.Text = "0.00"

            CCR.Text = Format(CDbl(txtSumTotalAmountDr.Text - txtSumAmountCr.Text), "#,##0.00")
            DDR.Text = "0.00"
        End If

        If txtSumAmountDr.Text = txtSumAmountCr.Text Then
            Cr.Text = "0.00"
            CCR.Text = "0.00"
            Dr.Text = "0.00"
            DDR.Text = "0.00"
        End If

        Call LoadColor()
    End Sub
    Private Sub FG_MouseDownEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseDownEvent) Handles FG.MouseDownEvent
        BtnSearch.Visible = False
        Select Case MouseButtons
            Case Windows.Forms.MouseButtons.Right
                FG.EditCell()
            Case Windows.Forms.MouseButtons.Left
                If FG.Col = 1 Then
                    'BtnSearch.Visible = True
                Else
                    'BtnSearch.Visible = False
                End If
                'If FG.Col = 1 Then
                '    BtnSearch.Left = CInt(FG.Left + (FG.CellLeft / 12) + (FG.CellWidth / 21))
                '    BtnSearch.Top = CInt((FG.CellTop / 15) + FG.Top)
                '    'Exit Sub
                'End If
                'If FG.Col = 2 Then
                '    BtnSearch.Visible = True
                '    BtnSearch.Left = CInt(FG.Left + (FG.CellLeft / 15) + (FG.CellWidth / 21))
                '    BtnSearch.Top = CInt((FG.CellTop / 15) + FG.Top)
                '    'Exit Sub 
                'End If

                'If FG.Col = 2 Then
                '    BtnSearch.Visible = True
                'Else
                '    BtnSearch.Visible = False
                'End If
                'If FG.Col = 2 Then
                '    BtnSearch.Left = CInt(FG.Left + (FG.CellLeft / 15) + (FG.CellWidth / 21))
                '    BtnSearch.Top = CInt((FG.CellTop / 15) + FG.Top)
                'End If

        End Select
    End Sub

    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelChange
        MD_Curr = CMB_Curr.Text
        MD_Rate = txtRate.Text
        If FG.Col = 1 Or FG.Col = 2 Or FG.Col = 5 Or FG.Col = 6 Or FG.Col = 9 Or FG.Col = 10 Then
            FG.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
        Else
            FG.Editable = VSFlex8U.EditableSettings.flexEDNone
        End If
        R = FG.Row()
        L = FG.Col
        Label1.Text = R
        Label2.Text = L
        If FG.Col = 1 Then
            If FG.get_TextMatrix(FG.Row, 2) <> "" Then
                FG.Editable = VSFlex8U.EditableSettings.flexEDNone
            Else
                FG.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
            End If
        End If
        If FG.Col = 2 Then
            If FG.get_TextMatrix(FG.Row, 1) <> "" Then
                FG.BackColorSel = Color.White
                FG.Editable = VSFlex8U.EditableSettings.flexEDNone
            Else
                FG.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
            End If
        End If


        Load_fg()
    End Sub
    Private Sub Amount_Later()
        'txtAmount.Text = IIf(IsNumeric(txtAmount.Text) = False, "0.00", Format(CDbl(txtAmount.Text), "##,##0.00"))
        'txtAmount.Text = Format(CDbl(txtAmount.Text), "##,##0.00")
        If CMB_Curr.Text = "LAK" Then
            txtAmount_Lak.Text = Format(CDbl(txtAmount.Text), "##,##0.00")
        Else

            txtAmount_Lak.Text = Format(CDbl(txtAmount.Text * txtRate.Text), "##,##0.00")
        End If
        curr = ""
        If CMB_Curr.SelectedIndex = 0 Then
            txtAmount_Later.Text = Letter_amt(txtAmount) & " ກີບ"
        ElseIf CMB_Curr.SelectedIndex = 1 Then

            txtAmount_Later.Text = Letter_amt(txtAmount) & " ກີບ"
        Else

            txtAmount_Later.Text = Letter_amt(txtAmount) & " ກີບ"
        End If

    End Sub
    Public Function Letter_amt(ByVal Txt As TextBox, Optional ByVal CurrKIP As Boolean = False) As String
        If Val(txtAmount.Text) <> 0 Then
            Letter_amt = CMoney(Format(CDbl(Txt.Text), "##0.00"))
        Else
            Letter_amt = ""
        End If
    End Function


    Private Sub BtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch.Click
        'MD_FROM = Me.Name
        'MD_KH = ""
        'Frm_Acccode_items.ShowDialog()
        fmShartOfAccDetail.txtSty.Text = "NsewJeneralJournal"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
        FG.Focus()
        Button3_Click(sender, e)
        'MsgBox(FG.get_TextMatrix(FG.Row, 1))
        FG.Focus()
        'R = FG.Row
        'L = FG.Col

        'If FG.Col = 1 Then
        '    FG.Col = FG.Col + 4
        'ElseIf FG.Col = 5 Then
        '    FG.Col = FG.Col + 1
        'ElseIf FG.Col = 6 Then
        'End If



        If Label2.Text = 1 Then
            'FG.Focus(
            'FG.set_TextMatrix(R, 4, FG.get_TextMatrix(FG.Row, 3))
        Else

        End If
        If txt_descrip.Text = "" Then
            txt_descrip.Text = FG.get_TextMatrix(1, 3)
        End If
        If FG.get_TextMatrix(FG.Row, 14) = "" Then
            FG.set_TextMatrix(FG.Row, 14, Txt_Referno.Text)
        End If
    End Sub

    Public Sub Load_fg()
        For i = 1 To FG.Rows - 2
            FG.set_TextMatrix(i, 0, i)
        Next i
    End Sub

    Private Sub BtnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew.Click
        Call Addnew()
        Load_fg()
    End Sub
    Private Sub Addnew()

        FG.Rows = 1

        FG.AddItem(FG.Row)
        'txtStff_Id.Text = MUserID
        'txtStff_NmL.Text = MUserName
        'If MWorkSetting = "" Then
        txt_dt.Value = Date.Today
        'Else
        '    txt_dt.Value = MWorkSetting

        'End If
        Txt_Referno.Text = ""
        txtBill_no.Text = ""

        'CMBBK_ID.SelectedIndex = 0
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
        'Call LoadCMB()
        txtAmount_USD.Text = "0.00"
        CMB_Curr.SelectedIndex = 0

        MDRate_DT = " and rate_dt<='" & Format(txt_dt.Value, "yyyy-MM-dd") & "'  "
        SS_Curr = " and AP_Rate_history.Curr =N'" & CMB_Curr.Text & "' "
        Call RateSetting()
        txtRate.Text = Format(MD_Rate, "#,##0.00")
        txtRateUSD.Text = Format(MDUSD_LAK, "#,##0.00")


        Call Amount_Later()

    End Sub

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        'If FG.get_TextMatrix(FG.Row, 1) = "" Or Me.FG.Rows = 1 Then Exit Sub
        If CDbl(txtSumTotalAmountDr.Text) <> CDbl(txtSumTotalAmountCr.Text) Then MsgBox("ບັນຊີນີ້ບໍ່ທັນບໍ່ທັນດູນດ່ຽງ ບໍ່ສາມາດບັນທຶກໄດ້!", MsgBoxStyle.OkOnly) : Exit Sub

        If FG.Rows = 1 Then MsgBox("ທ່ານຍັງບໍ່ທັນໄດ້ລົງລາຍການບັນຊີ", MsgBoxStyle.OkOnly) : Exit Sub
        If txtBill_no.Text = "" Then
            Call AutoNumber()
        End If

        If txt_descrip.Text = "" Then
            txt_descrip.Text = txtAC_code_nm.Text

        End If
        Call LoadSqlData("SELECT  top 1 Right(certify,3) As  certify    FROM Gen_jn WHERE   book =N'" & CMBBK_ID.Text & "' And  certify = N'" & txtBill_no.Text & "' And  year(date_work)=" & Format(CDate(txt_dt.Value), "yyyy") & " And  Month(date_work)=" & Format(CDate(txt_dt.Value), "MM") & "  And  day(date_work)=" & Format(CDate(txt_dt.Value), "dd") & " " & _
                    "  and LEFT(company,2)=N'" & Off_Id & "' Order by  Right(certify,3) DESC ", RSC)
        If RSC.RecordCount > 0 Then
            If txtBill_no.Enabled = True Then
                MsgBox("ເລກລະຫັດ : " & Trim(txtBill_no.Text) & " ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
                txtBill_no.BackColor = Color.Red
                txtBill_no.Focus()
                If RSC.State = ConnectionState.Open Then RSC.Close()
                Exit Sub
            End If

        End If

        'Call save()
        Call Load_Delete()
        'Call SaveItems_E()
        Call SaveItems()

        MsgBox("ບັນທຶກສຳເລັດ")
    End Sub
    Private Sub Load_Delete()
        'CNN.Execute("DELETE FROM AP_ACC_Adjust_Item WHERE certify=N'" & txtBill_no.Text & "' ")
        CNN.Execute("DELETE FROM gen_jn WHERE certify=N'" & txtBill_no.Text & "' ")
    End Sub

    Private Sub AutoNumber()
        Dim VIOT As New ADODB.Recordset
        Call LoadSqlData("SELECT top 1 right(certify,5) As Bill_no from gen_jn WHERE 1=1 AND Book=N'" & CMBBK_ID.Text & "' Order by certify DESC", VIOT)
        If VIOT.RecordCount <> 0 Then
            txtBill_no.Text = Format(CDbl(CDbl(VIOT.Fields("Bill_no").Value) + 1), "00000")
        Else
            txtBill_no.Text = Format(1, "00000")
        End If
        txtBill_no.Text = "AJ" & txtBill_no.Text

    End Sub
    Private Sub save()
        'Dim rs As New ADODB.Recordset
        'With rs
        '    Call LoadSqlData("SELECT certify FROM AP_ACC_Gen WHERE certify = '" & txtBill_no.Text & "'", rs)
        '    If .RecordCount = 0 Then
        '        'Dim sa As String =
        '        CNN.Execute("INSERT INTO AP_ACC_Gen (certify, date_work, Book, descrip, descripE, amount, Curr, Rate, net_amt, Amount_Later, AmountDr, AmountCr, TotalAmountDr, TotalAmountCr,office_id, com_id, last_update,  " & _
        '             " last_user, pc_nm) " & _
        '           " VALUES('" & (txtBill_no.Text) & "'," & _
        '           " '" & Format(txt_dt.Value, "yyyy-MM-dd") & "'," & _
        '              " N'" & Apostrophe(CMBBK_ID.Text) & "'," & _
        '                " N'" & Apostrophe(txt_descrip.Text) & "'," & _
        '                  " N'" & Apostrophe(txt_descripE.Text) & "'," & _
        '                   " " & CDbl(txtAmount.Text) & "," & _
        '                    " N'" & (txt_Curr.Text) & "'," & _
        '                 " " & CDbl(txtRate.Text) & "," & _
        '                 " " & CDbl(txtAmount_Lak.Text) & "," & _
        '                 " N'" & Apostrophe(txtAmount_Later.Text) & "'," & _
        '                 " " & CDbl(txtSumAmountDr.Text) & "," & _
        '                 " " & CDbl(txtSumAmountCr.Text) & "," & _
        '                 " " & CDbl(txtSumTotalAmountDr.Text) & "," & _
        '                 " " & CDbl(txtSumTotalAmountCr.Text) & "," & _
        '                  "  N'" & Apostrophe(Off_Id) & "'," & _
        '                  "  N'" & Apostrophe(Off_Id) & "'," & _
        '           " Getdate()," & _
        '           " N'" & MUserName & "'," & _
        '           " '" & COMPUTER_NM & "')")
        '    Else
        '        CNN.Execute("UPDATE AP_ACC_Gen SET " & _
        '           " date_work='" & Format(txt_dt.Value, "yyyy-MM-dd") & "'," & _
        '              " Book=N'" & Apostrophe(CMBBK_ID.Text) & "'," & _
        '                " descrip=N'" & Apostrophe(txt_descrip.Text) & "'," & _
        '                  " descripE=N'" & Apostrophe(txt_descripE.Text) & "'," & _
        '                   " amount=" & CDbl(txtAmount.Text) & "," & _
        '                    " Curr=N'" & (txt_Curr.Text) & "'," & _
        '                 " Rate=" & CDbl(txtRate.Text) & "," & _
        '                 " net_amt=" & CDbl(txtAmount_Lak.Text) & "," & _
        '                 " net_USD=" & CDbl(txtAmount_USD.Text) & "," & _
        '                 " Amount_Later=N'" & Apostrophe(txtAmount_Later.Text) & "'," & _
        '                 " AmountDr=" & CDbl(txtSumAmountDr.Text) & "," & _
        '                 " AmountCr=" & CDbl(txtSumAmountCr.Text) & "," & _
        '                 " TotalAmountDr=" & CDbl(txtSumTotalAmountDr.Text) & "," & _
        '                 " TotalAmountCr=" & CDbl(txtSumTotalAmountCr.Text) & "," & _
        '                     " office_id=N'" & Apostrophe(Off_Id) & "'," & _
        '         " com_id=N'" & Apostrophe(Off_Id) & "'," & _
        '             " last_update=Getdate()," & _
        '             " last_user=N'" & MUserName & "'," & _
        '             " pc_nm='" & COMPUTER_NM & "'" & _
        '             " WHERE certify= '" & (txtBill_no.Text) & "'")
        '    End If
        'End With
    End Sub
    Private Sub SaveItems()
        Dim J As Integer
        For J = 1 To FG.Rows - 1
            '===============
 

            If CMB_Curr.Text = "USD" Then
                FG.set_TextMatrix(J, 11, Format(CDbl(FG.get_TextMatrix(J, 9)), "#,##0.00"))
                FG.set_TextMatrix(J, 12, Format(CDbl(FG.get_TextMatrix(J, 10)), "#,##0.00"))
            Else
                FG.set_TextMatrix(J, 11, Format(CDbl(FG.get_TextMatrix(J, 9)) / CDbl(txtRateUSD.Text), "#,##0.00"))
                FG.set_TextMatrix(J, 12, Format(CDbl(FG.get_TextMatrix(J, 10)) / CDbl(txtRateUSD.Text), "#,##0.00"))
                'FG.set_TextMatrix(FG.Row, 14, Format(CDbl(FG.get_TextMatrix(FG.Row, 12)) / CDbl(txtRateUSD.Text), "#,##0.00"))
                'FG.set_TextMatrix(FG.Row, 15, Format(CDbl(FG.get_TextMatrix(FG.Row, 13)) / CDbl(txtRateUSD.Text), "#,##0.00"))
            End If

        Next J

        Dim Rschk As New ADODB.Recordset
        Dim i As Integer
        With Rschk
            For i = 1 To FG.Rows - 1
                Dim sk As String = "Select * FROM gen_jn  WHERE certify=N'" & txtBill_no.Text & "' AND ac_code='" & Apostrophe(FG.get_TextMatrix(i, 1)) + Apostrophe(FG.get_TextMatrix(i, 2)) & "' "
                Call LoadSqlData(sk, Rschk)
                If Rschk.RecordCount <> 0 Then
                    sk = "DELETE FROM gen_jn WHERE 1=1 AND certify='" & txtBill_no.Text & "' "
                    CNN.Execute(sk)
                End If
                'FG.FormatString = "^ລ/ດ|<ລະຫັດ  1   |<ຊື່ ວັດຖຸ    2           |^ຫົວໜ່ວຍ 6  |^ຈໍານວນ7   |<ລາຄາ    8   |<ລວມມູນຄ່າ  9     |^ສະກຸນເງິນ 13|ອັດຕາແລກປ່ຽນ 14 |<ລາຄາ  (ກີບ) 15    |<ລວມມູນຄ່າເປັນ  (ກີບ)  16    |<Pro_ID 17     "
                Dim sa As String = "INSERT INTO gen_jn ( date_work, book, certify, descrip, descripe, amount, curr, rate, Rate_USD, net_amt, code_dr, code_cr, ac_code, ac_name, ac_namee, " & _
                " amount_dr, amount_cr, Curr_i, Rate_i, amt_dr, amt_cr, amt_USD_Dr, amt_USD_Cr, my_lock,lock, company, office_id, com_id, last_update, last_user, pc_nm,AG,Frm) " & _
                " VALUES ( '" & Format(txt_dt.Value, "yyyy-MM-dd") & "'," & _
                "  N'" & Apostrophe(CMBBK_ID.Text) & "'," & _
                 "  N'" & Apostrophe(txtBill_no.Text) & "'," & _
                   "  N'" & Apostrophe(txt_descrip.Text) & "'," & _
                     "  N'" & Apostrophe(txt_descripE.Text) & "'," & _
                  " " & CDbl(txtAmount.Text) & ", " & _
                     "  N'" & "LAK" & "'," & _
                  " " & CDbl(1) & ", " & _
                      " " & CDbl(txtRateUSD.Text) & ", " & _
                   " " & CDbl(txtDiff.Text) & ", " & _
                " N'" & Apostrophe(FG.get_TextMatrix(i, 1)) & "'," & _
                " N'" & Apostrophe(FG.get_TextMatrix(i, 2)) & "'," & _
                " N'" & Apostrophe(FG.get_TextMatrix(i, 1)) + Apostrophe(FG.get_TextMatrix(i, 2)) & "'," & _
                 " N'" & Apostrophe(FG.get_TextMatrix(i, 3)) & "'," & _
                " N'" & Apostrophe(FG.get_TextMatrix(i, 4)) & "', " & _
                " " & CDbl(FG.get_TextMatrix(i, 5)) & ", " & _
                 " " & CDbl(FG.get_TextMatrix(i, 6)) & "," & _
                    " N'" & Apostrophe(FG.get_TextMatrix(i, 7)) & "'," & _
                    " " & CDbl(FG.get_TextMatrix(i, 8)) & "," & _
                  " " & CDbl(FG.get_TextMatrix(i, 9)) & ", " & _
                  " " & CDbl(FG.get_TextMatrix(i, 10)) & ", " & _
                        " " & CDbl(FG.get_TextMatrix(i, 11)) & ", " & _
                  " " & CDbl(FG.get_TextMatrix(i, 12)) & ", " & _
                  " " & CDbl(0) & "," & _
                        " " & CDbl(0) & "," & _
                  "  N'" & Apostrophe(MuSubOff2) & "'," & _
                         "  N'" & Apostrophe(MuSubOff2) & "'," & _
                         "  N'" & Apostrophe(MuSubOff2) & "'," & _
                " Getdate()," & _
                " N'" & Apostrophe(MUserName) & "'," & _
                " N'" & Apostrophe(MDServerName) & "',0,'1' )"
                CNN.Execute(sa)
            Next i
        End With
        CNN.Execute("delete gen_jn WHERE ac_code='' ")
    End Sub

    Private Sub CMB_Curr_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMB_Curr.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                FG.Rows = 1

                FG.Rows = 2
                'FG.Rows = FG.Rows + 1
                'FG.set_TextMatrix(FG.Rows - 1, 1, FG.get_TextMatrix(FG.Row, 1))
                'FG.set_TextMatrix(FG.Rows - 1, 2, FG.get_TextMatrix(FG.Row, 2))
                'FG.FormatString = "^ລ/ດ|<ເລກບັນຊີໜີ້ |<ເລກບັນຊີມີ| ຊື່ບັນຊີພາສາລາວ  |>ຈໍານວນເງິນຈົດໜີ້  |>ຈໍານວນເງິນຈົດມີ |<ກິດຈະກຳ|<ປະເພດລາຍຈ່າຍ|^ສະກຸນເງິນ|^ອັດຕາແລກປ່ຽນ |>ມູນຄ່າເປັນກີບ ໜີ້ |>ມູນຄ່າເປັນກີບ ມີ |>ມູນຄ່າເປັນ(USD) ໜີ້   |>ມູນຄ່າເປັນ(USD) ມີ  |>ເອກກະສານອ້າງອິງ"
                If txtDiff.Text > 0 Then
                    FG.set_TextMatrix(1, 1, txtAC_code.Text)
                    FG.set_TextMatrix(1, 3, txtAC_code_nm.Text)
                    FG.set_TextMatrix(1, 5, Format(CDbl(0), "#,##0.00"))
                    FG.set_TextMatrix(1, 6, Format(CDbl(0), "#,##0.00"))
                    FG.set_TextMatrix(1, 7, "LAK")
                    FG.set_TextMatrix(1, 8, Format(CDbl(1), "#,##0.00"))
                    FG.set_TextMatrix(1, 9, Format(CDbl(txtDiff.Text), "#,##0.00"))
                    FG.set_TextMatrix(1, 10, Format(CDbl(0), "#,##0.00"))

                    FG.Rows = FG.Rows + 1
                    FG.set_TextMatrix(2, 2, "5106100.00.000")
                    FG.set_TextMatrix(2, 3, "ກຳໄລຈາກການແລກປ່ຽນເງິນຕາຕ່າງປະເທດ")
                    FG.set_TextMatrix(2, 5, Format(CDbl(0), "#,##0.00"))
                    FG.set_TextMatrix(2, 6, Format(CDbl(0), "#,##0.00"))
                    FG.set_TextMatrix(2, 7, "LAK")
                    FG.set_TextMatrix(2, 8, Format(CDbl(1), "#,##0.00"))
                    FG.set_TextMatrix(2, 9, Format(CDbl(0), "#,##0.00"))
                    FG.set_TextMatrix(2, 10, Format(CDbl(txtDiff.Text), "#,##0.00"))
                Else
                    FG.set_TextMatrix(1, 1, "4106100.00.000")
                    FG.set_TextMatrix(1, 3, "ຂາດທຶນຈາກການແລກປ່ຽນເງິນຕາ ຕ່າງປະເທດ")
                    FG.set_TextMatrix(1, 5, Format(CDbl(0), "#,##0.00"))
                    FG.set_TextMatrix(1, 6, Format(CDbl(0), "#,##0.00"))
                    FG.set_TextMatrix(1, 7, "LAK")
                    FG.set_TextMatrix(1, 8, Format(CDbl(1), "#,##0.00"))
                    'FG.set_TextMatrix(1, 9, Format(CDbl(txtAmount_Lak.Text - txtamt.Text) * -1, "#,##0.00"))
                    FG.set_TextMatrix(1, 9, Format(CDbl(txtDiff.Text) * -1, "#,##0.00"))
                    FG.set_TextMatrix(1, 10, Format(CDbl(0), "#,##0.00"))

                    FG.Rows = FG.Rows + 1
                    FG.set_TextMatrix(2, 2, txtAC_code.Text)
                    FG.set_TextMatrix(2, 3, txtAC_code_nm.Text)
                    FG.set_TextMatrix(2, 5, Format(CDbl(0), "#,##0.00"))
                    FG.set_TextMatrix(2, 6, Format(CDbl(0), "#,##0.00"))
                    FG.set_TextMatrix(2, 7, "LAK")
                    FG.set_TextMatrix(2, 8, Format(CDbl(1), "#,##0.00"))
                    FG.set_TextMatrix(2, 9, Format(CDbl(0), "#,##0.00"))
                    'FG.set_TextMatrix(2, 10, Format(CDbl(txtAmount_Lak.Text - txtamt.Text) * -1, "#,##0.00"))
                    FG.set_TextMatrix(2, 10, Format(CDbl(txtDiff.Text) * -1, "#,##0.00"))
                End If

        End Select
        Dim J As Integer
        For J = 1 To FG.Rows - 1
            '===============


            If CMB_Curr.Text = "USD" Then
                FG.set_TextMatrix(J, 11, Format(CDbl(FG.get_TextMatrix(J, 9)), "#,##0.00"))
                FG.set_TextMatrix(J, 12, Format(CDbl(FG.get_TextMatrix(J, 10)), "#,##0.00"))
            Else
                FG.set_TextMatrix(J, 11, Format(CDbl(FG.get_TextMatrix(J, 9)) / CDbl(txtRateUSD.Text), "#,##0.00"))
                FG.set_TextMatrix(J, 12, Format(CDbl(FG.get_TextMatrix(J, 10)) / CDbl(txtRateUSD.Text), "#,##0.00"))
                'FG.set_TextMatrix(FG.Row, 14, Format(CDbl(FG.get_TextMatrix(FG.Row, 12)) / CDbl(txtRateUSD.Text), "#,##0.00"))
                'FG.set_TextMatrix(FG.Row, 15, Format(CDbl(FG.get_TextMatrix(FG.Row, 13)) / CDbl(txtRateUSD.Text), "#,##0.00"))
            End If

        Next J
    End Sub

    Private Sub CMB_Curr_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMB_Curr.SelectedIndexChanged
        Dim aRS As New ADODB.Recordset
        MDRate_DT = " and rate_dt<='" & Format(txt_dt.Value, "yyyy-MM-dd") & "'  "
        SS_Curr = " and AP_Rate_history.Curr =N'" & CMB_Curr.Text & "' "
        Call RateSetting()
        txtRate.Text = Format(MD_Rate, "#,##0.00")
        txtRateUSD.Text = Format(MDUSD_LAK, "#,##0.00")
        'Call LoadSqlData("select top 1  * from AP_Rate_history where curr='" & CMB_Curr.Text & "' order by rate_dt desc ", aRS)
        'If aRS.RecordCount <> 0 Then
        '    txtRate.Text = Format(CDbl(aRS.Fields("rate").Value), "##,##0.00")
        'Else
        '    txtRate.Text = "1.00"
        'End If
        txtAmount_Lak.Text = Format(CDbl(txtAmount.Text * txtRate.Text), "##,##0.00")
        'txtDiff.Text = Format(CDbl(txtamt.Text - txtAmount_Lak.Text), "#,##0.00")
        txtDiff.Text = Format(CDbl(txtAmount_Lak.Text - txtamt.Text), "#,##0.00")

        Call Amount_Later()
        Amount_USD()
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
        Else
            txtAmount_Lak.Text = Format(CDbl(txtAmount.Text), "#,##0.00")
            txtAmount_Lak.Text = Format(CDbl(txtAmount_Lak.Text), "#,##0.00")
            txtAmount_USD.Text = Format(CDbl(txtAmount_Lak.Text) / CDbl(MDUSD_LAK), "#,##0.00")

        End If


        'If CMB_Curr.Text = "USD" Then
        '    txtAmount_USD.Text = Format(CDbl(txtAmount.Text), "#,##0.00")
        '    txtAmount_Lak.Text = Format(CDbl(txtRate.Text) * CDbl(txtAmount.Text), "#,##0.00")
        'ElseIf CMB_Curr.Text = "SDR" Then

        '    txtAmount_USD.Text = Format(CDbl(MDUSD_THB) * CDbl(txtAmount.Text), "#,##0.00")

        '    txtAmount_Lak.Text = Format(CDbl(txtRate.Text) * CDbl(txtAmount.Text), "#,##0.00")

        'Else

        '    txtAmount_USD.Text = Format(CDbl(txtAmount.Text) / CDbl(txtRate.Text), "#,##0.00")
        '    txtAmount_Lak.Text = Format(CDbl(txtAmount.Text), "#,##0.00")

        'End If

    End Sub

    Private Sub CMBBK_ID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBBK_ID.SelectedIndexChanged
        Dim rs As New ADODB.Recordset
        Call LoadSqlData("Select * From books Where   bookid =N'" & Trim(CMBBK_ID.Text) & "'", rs)
        If rs.RecordCount > 0 Then
            txtBook_nm.Text = Trim(rs("bookid").Value).ToString
        End If
    End Sub

    Private Sub txtAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount.KeyDown

        Select Case e.KeyCode
            Case Keys.Enter
                Call Amount_Later()
                txtAmount.Text = Format(CDbl(txtAmount.Text), "##,##0.00")
                Amount_USD()
        End Select
    End Sub

    Private Sub txtAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAmount.KeyPress

    End Sub

    Private Sub txtAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAmount.LostFocus
        txtAmount.Text = Format(CDbl(txtAmount.Text), "##,##0.00")
    End Sub

    Private Sub txtAmount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmount.TextChanged
        Call Amount_Later()
    End Sub


    Private Sub txt_descrip_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_descrip.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                txt_descripE.Focus()
        End Select
    End Sub

    Private Sub txt_descrip_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_descrip.TextChanged

    End Sub

    Private Sub txt_descripE_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_descripE.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                txtAmount.Focus()
                txtAmount.SelectAll()
        End Select
    End Sub

    Private Sub txt_descripE_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_descripE.TextChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call Load_Calculate()
        'Call LoadColor()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'With RSC
        '    Dim ss As String = " SELECT     dbo.AP_ACC_Adjust.*, AP_Donnor.Don_Sym, AP_Donnor.Don_Nm_L, AP_Donnor.Don_Nm_E,   " & _
        '     "  dbo.books.bookname,dbo.AP_Office.off_nm,  dbo.AP_Office.off_StrtL, dbo.AP_Office.off_NoL,   " & _
        '  "   dbo.AP_Office.off_VillageL, dbo.AP_Office.Off_DistL, dbo.AP_Office.Off_ProvL , dbo.AP_Office.Tel, dbo.AP_Office.Fax, dbo.AP_Office.com_logo, " & _
        '"     dbo.AP_ACC_adjust_Item.code_dr, dbo.AP_ACC_adjust_Item.code_cr, dbo.AP_ACC_adjust_Item.ac_code, dbo.AP_ACC_adjust_Item.ac_name,dbo.AP_ACC_adjust_Item.amount_dr, dbo.AP_ACC_adjust_Item.amount_cr, dbo.AP_ACC_adjust_Item.Curr_i, " & _
        '"    dbo.AP_ACC_adjust_Item.Rate_i, dbo.AP_ACC_adjust_Item.amt_dr, dbo.AP_ACC_adjust_Item.amt_cr , dbo.AP_ACC_adjust_Item.Activity_id,dbo.AP_ACC_adjust_Item.cat_id, AP_ACC_adjust_Item.amt_USD_dr,AP_ACC_adjust_Item.amt_USD_cr " & _
        '     "                    FROM         dbo.AP_ACC_Adjust INNER JOIN   " & _
        '     "                   books ON AP_ACC_Adjust.Book = books.bookid INNER JOIN " & _
        '         "     AP_Donnor ON AP_ACC_Adjust.don_id = AP_Donnor.Don_ID INNER JOIN " & _
        '         "     AP_Office ON AP_ACC_Adjust.Com_id = AP_Office.off_id INNER JOIN " & _
        ' "    dbo.AP_ACC_adjust_Item ON dbo.AP_ACC_Adjust.certify = dbo.AP_ACC_adjust_Item.certify     " & _
        '    " WHERE  ( AP_ACC_Adjust.certify = '" & txtBill_no.Text & "')    "
        '    Call LoadSqlData(ss, RSC)
        '    If .RecordCount = 0 Then MsgBox("Data empty", vbInformation, "Check") : Exit Sub
        '    Dim Frm As New FrmPreview
        '    Dim Rpt As New CrystalReport_ACC_Adjust
        '    Rpt.SetDataSource(RSC)
        '    Rpt.Refresh()
        '    Frm.ReportViewer.ReportSource = Rpt
        '    Frm.ReportViewer.Zoom(100%)
        '    Frm.ReportViewer.DisplayGroupTree = False
        '    Frm.WindowState = FormWindowState.Maximized
        '    Frm.Show()
        'End With
    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click

    End Sub

    Private Sub txtRate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRate.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                If CDbl(txtRate.Text) < 1 Then
                    txtRate.Text = "1.00"
                End If
                txtRate.Text = Format(CDbl(txtRate.Text), "##,##0.00")



                If CMB_Curr.Text = "USD" Then
                    txtAmount_USD.Text = Format(CDbl(txtAmount.Text), "#,##0.00")
                    txtAmount_Lak.Text = Format(CDbl(txtRate.Text) * CDbl(txtAmount.Text), "#,##0.00")
                ElseIf CMB_Curr.Text = "THB" Then


                    txtAmount_Lak.Text = Format(CDbl(txtRate.Text) * CDbl(txtAmount.Text), "#,##0.00")
                    txtAmount_USD.Text = Format(CDbl(txtAmount_Lak.Text) / CDbl(MDUSD_LAK), "#,##0.00")

                Else


                    txtAmount_Lak.Text = Format(CDbl(txtAmount.Text), "#,##0.00")
                    txtAmount_USD.Text = Format(CDbl(txtAmount.Text) / CDbl(MDUSD_LAK), "#,##0.00")

                End If

                Sum()
        End Select
    End Sub

    Private Sub txtAC_code_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAC_code.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                Dim rs As New ADODB.Recordset
                With rs
                    Call LoadSqlData("Select *  from ACC_CODE Where   AC_CODE =N'" & Trim(txtAC_code.Text) & "'", rs)
                    If rs.RecordCount > 0 Then

                        txtAC_code_nm.Text = Trim(rs("Name_L").Value)
                        txtAC_type.Text = Trim(rs("Acc_Type").Value.ToString)
                    End If
                End With

                sumNew()
                'Sum()

        End Select
    End Sub

    Private Sub sumNew()
        Dim rsOp As New ADODB.Recordset
        Dim rssum As New ADODB.Recordset
        Dim rs As New ADODB.Recordset
        Dim OpAmountDr, OpAmountCr, OpAmtDr, OpAmtCr, AmountDr, AmountCr, AmtDr, AmtCr, BalAmount, BalAmt As Double
        Dim ss As String = "select name_l,name_e,acc_type,acc_typee from acc_code where ltrim(ac_code)='" & Trim(txtAC_code.Text) & "'"
        Call LoadSqlData(ss, rs)

        With rs
            If .RecordCount <> 0 Then
                If Lang = 0 Then   'Lao interface
                    txtAC_code_nm.Text = Trim(.Fields("Name_L").Value)
                    txtAC_type.Text = Trim(.Fields("acc_type").Value)
                Else
                    txtAC_code_nm.Text = Trim(.Fields("!Name_E").Value)
                    txtAC_type.Text = Trim(.Fields("acc_typee").Value)
                End If
                ss = "select amount_dr ,amount_cr,amt_dr, amt_cr  from Open_jn where ltrim(ac_code)='" & Trim(txtAC_code.Text) & "' and year(date_work)=" & Year(txt_dt.Value) & " and company=N'" & MuSubOff2 & "' "
                Call LoadSqlData(ss, rsOp)
                With rsOp
                    If .RecordCount = 0 Then
                        OpAmountDr = 0
                        OpAmountCr = 0
                        OpAmtDr = 0
                        OpAmtCr = 0
                    Else
                        OpAmountDr = CDbl(.Fields("amount_dr").Value)
                        OpAmountCr = CDbl(.Fields("amount_cr").Value)
                        OpAmtDr = CDbl(.Fields("amt_dr").Value)
                        OpAmtCr = CDbl(.Fields("amt_cr").Value)
                    End If
                End With

                ss = "select isnull(sum(amount_dr),0)  as amount_dr,  isnull(sum(amount_cr),0)  as amount_cr,   isnull(sum(amt_dr),0)  as amt_dr,  isnull(sum(amt_cr),0)  as amt_cr from gen_jn where office_id=N'" & MuSubOff2 & "' and ac_code='" & Trim(txtAC_code.Text) & "' and year(date_work)=" & Year(txt_dt.Value) & " and date_work<=convert(datetime,'" & Format(txt_dt.Value, "yyyy-MM-dd") & "',102)  "
                Call LoadSqlData(ss, rssum)
                With rssum
                    If .RecordCount = 0 Then
                        AmountDr = 0
                        AmountCr = 0
                        AmtDr = 0
                        AmtCr = 0
                    Else
                        AmountDr = CDbl(.Fields("amount_dr").Value)
                        AmountCr = CDbl(.Fields("amount_cr").Value)
                        AmtDr = CDbl(.Fields("amt_dr").Value)
                        AmtCr = CDbl(.Fields("amt_cr").Value)
                    End If
                End With
                BalAmount = (OpAmountDr - OpAmountCr) + (AmountDr - AmountCr)
                BalAmt = (OpAmtDr - OpAmtCr) + (AmtDr - AmtCr)
                txtAmount.Text = Format(BalAmount, "#,##0.00")
                txtamt.Text = Format(BalAmt, "#,##0.00")

                If Val(txtAmount.Text) <> 0 Then txtRete_AVG.Text = Format(CDbl(txtamt.Text) / CDbl(txtAmount.Text), "#,##0.00")
                CMB_Curr.Focus()
                CMB_Curr.SelectedIndex = 0
            End If
            txtDiff.Text = Format(CDbl(txtamt.Text) - CDbl(txtAmount_Lak.Text), "#,##0.00")
        End With

    End Sub

    Private Sub Sum()
        Dim aa As String

        aa = "    delete RPT_Adjus_Curr "
        CNN.Execute(aa)
        'and Curr='" & CMB_Curr.Text & "'
        'aa = "SELECT ac_code  FROM gen_jn  where    ac_code='" & txtAC_code.Text & "' and Curr='" & CMB_Curr.Text & "' and  year(date_work)='" & Format(CDate(txt_dt.Value), "yyyy") & "' and  date_work<='" & Format(CDate(txt_dt.Value), "yyyy-MM-dd") & "'  " & shr_Department & "  " & Shr_Donnor & "  " & Office_com & " "
        aa = "SELECT ac_code  FROM gen_jn  where    ac_code='" & txtAC_code.Text & "' and  year(date_work)='" & Format(CDate(txt_dt.Value), "yyyy") & "' and  date_work<='" & Format(CDate(txt_dt.Value), "yyyy-MM-dd") & "'   "
        Call LoadSqlData(aa, RSC)

        If RSC.RecordCount > 0 Then
            aa = "     insert into  RPT_Adjus_Curr (ac_code,Amount_dr,Amount_cr,amt_dr,amt_cr) " & _
                " SELECT  ac_code,SUM (Amount_dr),SUM (Amount_cr) ,SUM (amt_dr),SUM (amt_cr)  from  Open_jn  where   ac_code='" & txtAC_code.Text & "'  and  year(date_work)='" & Format(CDate(txt_dt.Value), "yyyy") & "' " & _
                      "   group by  ac_code   "
            CNN.Execute(aa)
            'aa = "     insert into  RPT_Adjus_Curr (ac_code,Amount_dr,Amount_cr,amt_dr,amt_cr) " & _
            '    " SELECT  ac_code,SUM (Amount_dr),SUM (Amount_cr) ,SUM (amt_dr),SUM (amt_cr)  from  gen_jn  where   ac_code='" & txtAC_code.Text & "' and Curr='" & CMB_Curr.Text & "'  and  year(date_work)='" & Format(CDate(txt_dt.Value), "yyyy") & "' and  date_work<='" & Format(CDate(txt_dt.Value), "yyyy-MM-dd") & "'  " & shr_Department & "  " & Shr_Donnor & "  " & Office_com & " " & _
            '          "   group by  ac_code   "
            aa = "     insert into  RPT_Adjus_Curr (ac_code,Amount_dr,Amount_cr,amt_dr,amt_cr) " & _
                " SELECT  ac_code,SUM (Amount_dr),SUM (Amount_cr) ,SUM (amt_dr),SUM (amt_cr)  from  gen_jn  where   ac_code='" & txtAC_code.Text & "'  and  year(date_work)='" & Format(CDate(txt_dt.Value), "yyyy") & "' and  date_work<='" & Format(CDate(txt_dt.Value), "yyyy-MM-dd") & "'  " & _
                      "   group by  ac_code   "
            CNN.Execute(aa)
            aa = "    delete RPT_Adjus_Curr1 "
            CNN.Execute(aa)
            '     aa = "     insert into  RPT_Adjus_Curr1 (ac_code,Amount,amt) " & _
            '" SELECT  ac_code,sum(Amount_dr-Amount_cr),sum (amt_dr-amt_cr)  from  RPT_Adjus_Curr  group by  ac_code   "
            '     CNN.Execute(aa)
            aa = "     insert into  RPT_Adjus_Curr1 (ac_code,Amount,amt) " & _
          " SELECT  ac_code,sum(Amount_dr),sum (amt_dr)  from  RPT_Adjus_Curr  group by  ac_code   "
            CNN.Execute(aa)

            aa = "SELECT *  FROM RPT_Adjus_Curr1   "
            Call LoadSqlData(aa, RSC)
            If RSC.RecordCount > 0 Then
                txtAmount.Text = Format(CDbl(Trim(RSC("Amount").Value)), "#,##0.00")
                Amt = Format(CDbl(Trim(RSC("amt").Value)), "#,##0.00")
                If txtAmount.Text > 0 Then txtRete_AVG.Text = Format(CDbl(Amt / txtAmount.Text), "#,##0.00")

                txtamt.Text = Format(CDbl(Trim(RSC("amt").Value)), "#,##0.00")
                txtAmount_Lak.Text = Format(CDbl(txtAmount.Text * txtRate.Text), "#,##0.00")

                txtDiff.Text = Format(CDbl(txtamt.Text - txtAmount_Lak.Text), "#,##0.00")
            End If

        End If

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        fmShartOfAccDetail.ShowDialog()
        If Acc_Code = "" Then Exit Sub

        Dim rs As New ADODB.Recordset
        With rs
            Call LoadSqlData("Select *  from ACC_CODE Where   AC_CODE =N'" & Trim(Acc_Code) & "'", rs)
            If rs.RecordCount > 0 Then
                txtAC_code.Text = Trim(rs("AC_CODE").Value)
                txtAC_code_nm.Text = Trim(rs("Name_L").Value)
                txtAC_type.Text = Trim(rs("Acc_Type").Value.ToString)
            End If
        End With


        Sum()
        With rs
            Call LoadSqlData("Select open_amt_dr-open_amt_cr  as amt  from RPT_Barande  Where   AC_CODE =N'" & Trim(txtAC_code.Text) & "'", rs)
            If rs.RecordCount > 0 Then
                txtAmount.Text = Format(CDbl(Trim(rs("amt").Value)), "#,##0.00")
            End If
        End With

        With rs
            Call LoadSqlData("Select   rate    from gen_jn  Where   AC_CODE =N'" & Trim(txtAC_code.Text) & "'", rs)
            If rs.RecordCount > 0 Then
                txtRete_AVG.Text = Format(CDbl(Trim(rs("rate").Value)), "#,##0.00")
            End If
        End With

        txtamt.Text = Format(CDbl(txtRete_AVG.Text) * CDbl(txtAmount.Text), "#,##0.00")

    End Sub

    Private Sub txtAC_code_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAC_code.TextChanged

    End Sub
End Class