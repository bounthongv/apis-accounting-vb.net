Public Class FmReceipt
    Dim sql As String
    Dim MDCurrency As String
    Dim MDRate As String
    Dim At As Boolean
    Dim lastCurr As String


    Private Sub FmReceipt_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Call MdiCNum()
        At = 0
        BtnPreview.Enabled = False
        BtnDelete.Enabled = False
        TextBox11.Focus()
        FmReceipt_List.loadSQL()
        FmReceipt_List.LoadListFG()
    End Sub
    Private Sub AtotoInsert()
        Dim rst As Double
        rst = TextBox11.Text
        Dim i As Integer
        If At = 0 Then
            Exit Sub
        End If
        For i = 1 To FGPaper.Rows - 1
            If CDbl(rst) >= CDbl(FGPaper.get_TextMatrix(i, 5)) Then
                FGPaper.set_TextMatrix(i, 2, Int(CDbl(rst) / CDbl(FGPaper.get_TextMatrix(i, 5))))
                rst = CDbl(rst) - CDbl(CDbl(FGPaper.get_TextMatrix(i, 5)) * CDbl(FGPaper.get_TextMatrix(i, 2)))
                FGPaper.set_TextMatrix(i, 3, Format(CDbl(FGPaper.get_TextMatrix(i, 2) * FGPaper.get_TextMatrix(i, 5)), "##,##0.00"))
                FGPaper.set_TextMatrix(i, 4, Format(CDbl(FGPaper.get_TextMatrix(i, 3) * CDbl(Rate.Text)), "##,##0.00"))
            Else
                FGPaper.set_TextMatrix(i, 2, 0)
            End If
        Next i
        For i = 1 To FGPaper.Rows - 1
            FGPaper.Row = i
            'MsgBox(FGPaper.get_TextMatrix(i, 2))
            If FGPaper.get_TextMatrix(i, 2) = 0 Then
                FGPaper.Col = 1
                FGPaper.CellBackColor = Color.White
                FGPaper.Col = 2
                FGPaper.CellBackColor = Color.White
                FGPaper.Col = 4
                FGPaper.CellBackColor = Color.White
                FGPaper.Col = 5
                FGPaper.CellBackColor = Color.White
                FGPaper.Col = 3
                FGPaper.CellBackColor = Color.White
            Else
                FGPaper.Col = 1
                FGPaper.CellBackColor = Color.LightCyan
                FGPaper.Col = 2
                FGPaper.CellBackColor = Color.LightCyan
                FGPaper.Col = 4
                FGPaper.CellBackColor = Color.LightCyan
                FGPaper.Col = 5
                FGPaper.CellBackColor = Color.LightCyan
                FGPaper.Col = 3
                FGPaper.CellBackColor = Color.LightCyan
            End If
        Next i
        Call SumData()

    End Sub
    Private Sub FmReceipt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboBox1.BackColor = Color.White
        FGPaper.set_ColHidden(5, True)
        At = 1
        'ComboBox1.Text = "ການຊື້ເງິນຕາ"

        BtnDelete.Enabled = False
        BtnPreview.Enabled = False
        Cashier.Text = MUserName
        TabControl1.SelectedIndex = 1
        LoadCurr()
        LoadListFGRate()

        BtnDelete.Enabled = False
        FG1.FormatString = "^ລ/ດ |< ເລກບິນ   |<  ມື້ລົງສັນຍາ |< ຊື່ບັນຊີ         |>ມູນຄ່າເງິນ          |<ເງິນ |>ທຽບກີບ          "
        FG2.FormatString = "^ລ/ດ|< ເງິນ     |<ຈ/ນ|< ມູນຄ່າ     "
        FGPaper.FormatString = "^ລ/ດ|< ເງິນ (ໃບ)                  |^ຈ/ນ(ໃບ)  | ມູນຄ່າ              |>ທຽບກີບ               |"
        FGRate.FormatString = "^ລ/ດ|^ ເງິນ |ອັດຕາ      "



        Cmb.SelectedIndex = 0


        Curr.Text = "LAK"
        Rate.Text = "1.00"
        LoadListFGRatePaper()
        TotalLAK.Text = "0.00"
        TotelAmt.Text = "0.00"
        txtAmt_letter.Text = ""
        RadioButton2.Checked = True
        'TextBox11.ReadOnly = False
        Call SumData()
        Me.WindowState = FormWindowState.Maximized
        FGRate.Size = New System.Drawing.Size(172, 427)
        FG1.Size = New System.Drawing.Size(764, 403)
        FGPaper.Size = New System.Drawing.Size(824, 237)
        TextBox11.Focus()
        Call LoadSqlData("select Curr_last  from Ap_RateSeting WHERE curr= '" & Curr.Text & "'", RSC)
        If RSC.RecordCount <> 0 Then
            lastCurr = RSC.Fields("Curr_last").Value
        End If
        'ComboBox1.Text = MDReceiptType
        ComboBox1.Items.Clear()

        Call load_Cmb(" SELECT bookname FROM  books  where Type='RCTY' ", "bookname", ComboBox1)
        ComboBox1.SelectedIndex = 0
    End Sub


    Private Sub LoadCurr()
        Dim Comm As ADODB.Command
        Dim rsat As New ADODB.Recordset
        Comm = New ADODB.Command
        Comm.ActiveConnection = CNN
        Comm.CommandText = "SELECT Curr FROM Ap_RateSeting WHERE Curr <> '" & "" & " order by Curr'"
        rsat = Comm.Execute
        If rsat.RecordCount <> 0 Then
            While Not rsat.EOF()
                Cmb.Items.Add(Trim(rsat.Fields("Curr").Value))
                rsat.MoveNext()
            End While
            Cmb.Items.Add("==ທັງຫມົດ==")
        End If
    End Sub
    Public Sub FormateText()


        'Rate.Text = Format(CDbl(Rate.Text), "##,##0.00")
        TotelAmt.Text = Format(CDbl(TotelAmt.Text), "##,##0.00")
        TotalLAK.Text = Format(CDbl(TotalLAK.Text), "##,##0.00")

    End Sub

    Public Sub SumData()
        Dim TotalAm, TotalL, Un As Double
        For i = 1 To FGPaper.Rows - 1
            Un = Un + CDbl(FGPaper.get_TextMatrix(i, 2))
        Next
        Unit.Text = Format(CDbl(Un), "##,##0.00")
        For i = 1 To FGPaper.Rows - 1
            TotalAm = TotalAm + CDbl(FGPaper.get_TextMatrix(i, 3))
            FGPaper.set_TextMatrix(i, 3, Format(CDbl(CDbl(FGPaper.get_TextMatrix(i, 3))), "##,##0.00"))
            'TotalL = TotalL + CDbl(FGPaper.get_TextMatrix(i, 4))
        Next
        TotelAmt.Text = Format(CDbl(TotalAm), "##,##0.00")
        For i = 1 To FGPaper.Rows - 1
            TotalL = TotalL + CDbl(FGPaper.get_TextMatrix(i, 4))
        Next
        TotalLAK.Text = Format(CDbl(TotalL), "##,##0.00")
        TextBox1.Text = Format(CDbl(TextBox11.Text) - CDbl(TotelAmt.Text), "##,##0.00")
    End Sub
    Public Sub loadColor()
        Dim rmRS As New ADODB.Recordset
        Dim J As Integer
        FGPaper.Redraw = False
        For J = 1 To FGPaper.Rows - 1
            If CDbl(FGPaper.get_TextMatrix(J, 2)) > 0 Then
                FGPaper.Row = J
                FGPaper.Col = 4
                FGPaper.CellBackColor = Color.LightCyan
            End If
        Next J
        FGPaper.Redraw = True
    End Sub
    Private Sub LoadListFG()
        'loadSQL()
        FG1.Rows = 1
        With RSC
            Call LoadSqlData("select *  from Ap_Receipt WHERE Receipt_No<>''" & sql & " order by Receipt_No", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG1.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("Receipt_No").Value)) & _
                                "" & vbTab & Trim(CStr(.Fields("InDate").Value)) & _
                                   "" & vbTab & Trim(CStr(.Fields("Bnk_Ac_Name").Value)) & _
                                 "" & vbTab & Trim(Format(CDbl(.Fields("Amt").Value), "##,##0.00")) & _
                                    "" & vbTab & Trim(CStr(.Fields("Curr").Value)) & _
                                       "" & vbTab & Trim(Format(CDbl(.Fields("Amt_In_LAK").Value), "##,##0.00")))
                    .MoveNext()
                End While
            Else
                FG1.Rows = 1
            End If
        End With
        SumData()
    End Sub
    Private Sub LoadListFGRatePaper()
        FGPaper.Rows = 1
        With RSC
            Call LoadSqlData("select Curr,Paper  from Ap_MoneyPaper Where Curr='" & Curr.Text & "' order by Paper DESC", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FGPaper.AddItem(.AbsolutePosition & vbTab & "ເງິນໃບ (" & Curr.Text & ") =>  " & Trim(Format(CDbl(.Fields("Paper").Value), "##,##0.00")) & _
                                       "" & vbTab & "0" & _
                                        "" & vbTab & "0.00" & _
                                         "" & vbTab & "0.00" & _
                                         "" & vbTab & Trim(CDbl(.Fields("Paper").Value)))
                    .MoveNext()
                End While
            Else
                FGPaper.Rows = 2
            End If
        End With
    End Sub
    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        BtnDelete.Enabled = False
        BtnPreview.Enabled = False
        AutoNumberLAK()
        If ComboBox1.Text = "" Then MessageBox.Show("ກະລຸນນາເລືອກປະເພດກ່ອນ") : ComboBox1.BackColor = Color.Red : Exit Sub
        CNN.Execute("insert into Ap_Receipt (Receipt_No , Receipt_Type , InDate , Bnk_Ac_Code , Bnk_Ac_Name , Amt , Curr , Rate , Amt_In_LAK , Payer , Cashier , Last_User , Send_To , Last_UpDate , Remark , Status , Company ) values('" & Receipt_No.Text & "' , N'" & ComboBox1.Text & "' , '" & Format(Indate.Value, "MM-dd-yyyy") & "' , '" & Bnk_Ac_Code.Text & "' , N'" & Bnk_Ac_Name.Text & "' , '" & CDbl(TotelAmt.Text) & "' , '" & Curr.Text & "' , '" & CDbl(Rate.Text) & "' , '" & CDbl(TotalLAK.Text) & "' , '" & Payment.Text & "' , N'" & Cashier.Text & "' , N'" & MUserName & "' , '" & "No" & "' , '" & Format(MWorkSetting, "MM-dd-yyyy") & "' , '" & Remark.Text & "', '0',N'" & MuSubOff & "')")
        For i = 1 To FGPaper.Rows - 1
            If CDbl(FGPaper.get_TextMatrix(i, 2)) > 0 Then
                CNN.Execute("insert into Ap_ReceipItem (Receip_No , Amt , Unit ) values('" & Receipt_No.Text & "' , '" & CDbl(FGPaper.get_TextMatrix(i, 5)) & "' , '" & CDbl(FGPaper.get_TextMatrix(i, 2)) & "')")
            End If
        Next
        MessageBox.Show("ການບັນທຶກສຳເລັດ")
     
      

    End Sub
    Private Sub LoadListFGRate()
        FGRate.Rows = 1
        With RSC
            Call LoadSqlData("select Curr,Rate  from Ap_RateSeting order by Curr", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FGRate.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("Curr").Value)) & _
                       "" & vbTab & Trim(Format(CDbl(.Fields("Rate").Value), "##,##0.00")))
                    .MoveNext()
                End While
            Else
                FGRate.Rows = 2
            End If
        End With
    End Sub
    Private Sub LoadListFG2()
        FG2.Rows = 1
        With RSC
            Call LoadSqlData("select *  from Ap_ReceipItem WHERE Receip_No='" & FG1.get_TextMatrix(FG1.Row, 1) & "' order by Amt DESC", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG2.AddItem(.AbsolutePosition & vbTab & Trim(Format(CDbl(.Fields("Amt").Value), "##,##0.00")) & _
                                "" & vbTab & Trim(CStr(.Fields("Unit").Value)) & _
                                "" & vbTab & Format(CDbl(CDbl(Trim(CStr(.Fields("Amt").Value))) * CDbl(Trim(CStr(.Fields("Unit").Value)))), "##,##0.00"))
                    .MoveNext()
                End While
            Else
                FG2.Rows = 16
            End If
        End With
    End Sub
    Private Sub AutoNumberLAK()
        Dim srNum As New ADODB.Recordset
        Dim mNum As Integer
        Call LoadSqlData("SELECT top 1  Receipt_No FROM  Ap_Receipt Order by Receipt_No DESC", srNum)
        If srNum.RecordCount = 0 Then
            Receipt_No.Text = "000001"
        Else
            'mNum = Microsoft.VisualBasic.Right(CDbl(Val(srNum.Fields("Receipt_No").Value)), 1) + 1
            mNum = CDbl(Val(srNum.Fields("Receipt_No").Value)) + 1
            If Len(CStr(mNum).Trim) = 1 Then
                Receipt_No.Text = "00000" & CStr(mNum)
            ElseIf Len(CStr(mNum).Trim) = 2 Then
                Receipt_No.Text = "0000" & CStr(mNum)
            ElseIf Len(CStr(mNum).Trim) = 3 Then
                Receipt_No.Text = "000" & CStr(mNum)
            ElseIf Len(CStr(mNum).Trim) = 4 Then
                Receipt_No.Text = "00" & CStr(mNum)
            ElseIf Len(CStr(mNum).Trim) = 5 Then
                Receipt_No.Text = "0" & CStr(mNum)
            ElseIf Len(CStr(mNum).Trim) >= 6 Then
                Receipt_No.Text = CStr(mNum)
            End If
        End If
    End Sub

    Private Sub FG1_DblClick(ByVal sender As Object, ByVal e As System.EventArgs)

        '-----------LAK
        'If MDCurrency = "LAK" Then
        '    For i = 1 To FG2.Rows - 1
        '        If FG2.get_TextMatrix(i, 3) = "50,000.00" Then
        '            LAK1.Text = FG2.get_TextMatrix(i, 2)
        '            MessageBox.Show("ok")
        '        ElseIf FG2.get_TextMatrix(i, 3) = "20,000.00" Then
        '            LAK2.Text = FG2.get_TextMatrix(i, 2)
        '        ElseIf FG2.get_TextMatrix(i, 3) = "10,000.00" Then
        '            LAK3.Text = FG2.get_TextMatrix(i, 2)
        '        ElseIf FG2.get_TextMatrix(i, 3) = "5,000.00" Then
        '            LAK4.Text = FG2.get_TextMatrix(i, 2)
        '        ElseIf FG2.get_TextMatrix(i, 3) = "2,000.00" Then
        '            LAK5.Text = FG2.get_TextMatrix(i, 2)
        '        ElseIf FG2.get_TextMatrix(i, 3) = "1,000.00" Then
        '            LAK6.Text = FG2.get_TextMatrix(i, 2)
        '        ElseIf FG2.get_TextMatrix(i, 3) = "1,000.00" Then
        '            LAK7.Text = FG2.get_TextMatrix(i, 2)
        '        End If
        '    Next

        'End If


    End Sub

    Private Sub FG1_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

  
  
    Private Sub AxVSFlexGrid2_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FGRate.SelChange
        If FGRate.Row And FGRate.Col > 0 Then
            If FGRate.get_TextMatrix(1, 1) <> "" Then
                Curr.Text = FGRate.get_TextMatrix(FGRate.Row, 1)
                Rate.Text = FGRate.get_TextMatrix(FGRate.Row, 2)
                LoadListFGRatePaper()
                TotalLAK.Text = "0.00"
                TotelAmt.Text = "0.00"
                txtAmt_letter.Text = ""


                Call LoadSqlData("select Curr_last  from Ap_RateSeting WHERE curr= '" & Curr.Text & "'", RSC)
                If RSC.RecordCount <> 0 Then
                    lastCurr = RSC.Fields("Curr_last").Value
                End If


            End If
        End If
    End Sub

 

    Private Sub FGPaper_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

  
    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        AutoNumberLAK()
        BtnDelete.Enabled = False
        BtnPreview.Enabled = False
        Receipt_No.Clear()
        Bnk_Ac_Code.Clear()
        Bnk_Ac_Name.Clear()
        Indate.Text = ""
        txtAmt_letter.Clear()
        Remark.Clear()

    End Sub

    Private Sub txtAmt_letter_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmt_letter.TextChanged

    End Sub
    Public Function Letter_amt(ByVal Txt As TextBox, Optional ByVal CurrKIP As Boolean = False) As String

        If Val(TotelAmt.Text) <> 0 Then
            Letter_amt = CMoney(Format(CDbl(Txt.Text), "##0.00"))
        Else
            Letter_amt = ""
        End If

    End Function

    Private Sub TotelAmt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TotelAmt.TextChanged
        txtAmt_letter.Text = Letter_amt(TotelAmt) & ": " & lastCurr
    End Sub

    Private Sub Cmb_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmb.SelectedIndexChanged
        sql = ""
        If Cmb.Text <> "==ທັງຫມົດ==" Then
            sql = " AND Curr = '" & Cmb.Text & "'  AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
        Else
            sql = "   AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
        End If
        'sql = "   AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
        LoadListFG()
    End Sub

  
    Private Sub BtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelete.Click

        If MessageBox.Show("ທ່ານຕ້ອງລຶບລະຫັດ " & Bnk_Ac_Code.Text & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            CNN.Execute("delete from Ap_ReceipItem where Receip_No ='" & FG1.get_TextMatrix(FG1.Row, 1) & "' ")
            CNN.Execute("delete from Ap_Receipt where Receipt_No ='" & FG1.get_TextMatrix(FG1.Row, 1) & "' ")
            loadSQL()
            LoadListFG()
            FG2.Rows = 1
            FG2.Rows = 2
            BtnDelete.Enabled = False
            BtnPreview.Enabled = False
        End If
    End Sub

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Close()

    End Sub

    Private Sub BtnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew.Click
        BtnPreview.Enabled = False
        BtnDelete.Enabled = False
        Receipt_No.Clear()
        Bnk_Ac_Code.Clear()
        Bnk_Ac_Name.Clear()
        Indate.Text = ""
        txtAmt_letter.Clear()
        Remark.Clear()
        TabControl1.SelectedIndex = 1
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        BtnPreview.Enabled = False
        BtnDelete.Enabled = False
        Close()
    End Sub

    Private Sub BtnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        sql = ""
        sql = " AND Receip_No = '" & FG1.get_TextMatrix(FG1.Row, 1) & "' "
      

    End Sub

    Private Sub FG1_SelChange_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG1.SelChange
        If FG1.Row And FG1.Col > 0 Then
            If FG1.get_TextMatrix(1, 1) <> "" Then
                BtnDelete.Enabled = True
                BtnPreview.Enabled = True
                LoadListFG2()
            End If
        End If
    End Sub

    Private Sub FGPaper_AfterEdit(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_AfterEditEvent) Handles FGPaper.AfterEdit
        If FGPaper.Col = 2 Then

            If IsNumeric(FGPaper.get_TextMatrix(FGPaper.Row, 2)) = False Then MsgBox("ກະລຸນນາປ້ອນເປັນໂຕເລກ") : FGPaper.set_TextMatrix(FGPaper.Row, 2, "0") : FGPaper.set_TextMatrix(FGPaper.Row, 3, "0.00") : FGPaper.set_TextMatrix(FGPaper.Row, 4, "0.00")
            If FGPaper.get_TextMatrix(FGPaper.Row, 2) = "" Then FGPaper.set_TextMatrix(FGPaper.Row, 2, "0") : FGPaper.set_TextMatrix(FGPaper.Row, 3, "0.00") : FGPaper.set_TextMatrix(FGPaper.Row, 4, "0.00")
            FGPaper.set_TextMatrix(FGPaper.Row, 3, Format(CDbl(FGPaper.get_TextMatrix(FGPaper.Row, 2) * FGPaper.get_TextMatrix(FGPaper.Row, 5)), "##,##0.00"))
            If FGPaper.get_TextMatrix(FGPaper.Row, FGPaper.Col) > 0 Then
                FGPaper.Col = 1
                FGPaper.CellBackColor = Color.LightCyan
                FGPaper.Col = 3
                FGPaper.CellBackColor = Color.LightCyan
                FGPaper.Col = 4
                FGPaper.CellBackColor = Color.LightCyan
                FGPaper.Col = 5
                FGPaper.CellBackColor = Color.LightCyan
                FGPaper.Col = 2
                FGPaper.CellBackColor = Color.LightCyan
            Else
                FGPaper.Col = 1
                FGPaper.CellBackColor = Color.White
                FGPaper.Col = 3
                FGPaper.CellBackColor = Color.White
                FGPaper.Col = 4
                FGPaper.CellBackColor = Color.White
                FGPaper.Col = 5
                FGPaper.CellBackColor = Color.White
                FGPaper.Col = 2
                FGPaper.CellBackColor = Color.White
            End If
        End If
        If FGPaper.Col = 3 Then
            If IsNumeric(FGPaper.get_TextMatrix(FGPaper.Row, 3)) = False Then MsgBox("ກະລຸນນາປ້ອນເປັນໂຕເລກ") : FGPaper.set_TextMatrix(FGPaper.Row, 2, "0") : FGPaper.set_TextMatrix(FGPaper.Row, 3, "0.00") : FGPaper.set_TextMatrix(FGPaper.Row, 4, "0.00")
            If FGPaper.get_TextMatrix(FGPaper.Row, 3) = "" Then FGPaper.set_TextMatrix(FGPaper.Row, 2, "0") : FGPaper.set_TextMatrix(FGPaper.Row, 3, "0.00") : FGPaper.set_TextMatrix(FGPaper.Row, 4, "0.00")
            FGPaper.set_TextMatrix(FGPaper.Row, 2, CDbl(FGPaper.get_TextMatrix(FGPaper.Row, 3) / FGPaper.get_TextMatrix(FGPaper.Row, 5)))
            If FGPaper.get_TextMatrix(FGPaper.Row, FGPaper.Col) > 0 Then
                FGPaper.Col = 1
                FGPaper.CellBackColor = Color.LightCyan
                FGPaper.Col = 2
                FGPaper.CellBackColor = Color.LightCyan
                FGPaper.Col = 4
                FGPaper.CellBackColor = Color.LightCyan
                FGPaper.Col = 5
                FGPaper.CellBackColor = Color.LightCyan
                FGPaper.Col = 3
                FGPaper.CellBackColor = Color.LightCyan
            Else
                FGPaper.Col = 1
                FGPaper.CellBackColor = Color.White
                FGPaper.Col = 2
                FGPaper.CellBackColor = Color.White
                FGPaper.Col = 4
                FGPaper.CellBackColor = Color.White
                FGPaper.Col = 5
                FGPaper.CellBackColor = Color.White
                FGPaper.Col = 3
                FGPaper.CellBackColor = Color.White
            End If
        End If
        FGPaper.set_TextMatrix(FGPaper.Row, 4, Format(CDbl(FGPaper.get_TextMatrix(FGPaper.Row, 3) * CDbl(Rate.Text)), "##,##0.00"))


        FGPaper.Row = FGPaper.Row + 1
        Call SumData()
    End Sub



    
    Private Sub FGPaper_SelChange_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FGPaper.SelChange

        If FGPaper.Col = 2 Then
            If IsNumeric(FGPaper.get_TextMatrix(FGPaper.Row, 2)) = False Then MsgBox("ກະລຸນນາປ້ອນເປັນໂຕເລກ") : FGPaper.set_TextMatrix(FGPaper.Row, 2, "0") : FGPaper.set_TextMatrix(FGPaper.Row, 3, "0.00") : FGPaper.set_TextMatrix(FGPaper.Row, 4, "0.00")

        End If


        If FGPaper.Col = 2 Or FGPaper.Col = 3 Then
            FGPaper.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
            FGPaper.FocusRect = VSFlex8U.FocusRectSettings.flexFocusNone
        Else
            FGPaper.Editable = VSFlex8U.EditableSettings.flexEDNone
            FGPaper.FocusRect = VSFlex8U.FocusRectSettings.flexFocusLight
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        LoadListFGRatePaper()
        'TextBox11.ReadOnly = True

        LoadListFGRatePaper()
        Call SumData()
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
    


        Call AtotoInsert()
        'TextBox11.ReadOnly = False
        Call SumData()
        TextBox11.Focus()
    End Sub

    Private Sub TextBox11_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox11.KeyPress
        If e.KeyChar = Chr(13) Then
            TextBox11.Text = Format(CDbl(TextBox11.Text), "##,##0.00")
            Call AtotoInsert()
        End If
    End Sub

    Private Sub TextBox11_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox11.LostFocus
        TextBox11.Text = Format(CDbl(TextBox11.Text), "##,##0.00")
        If RadioButton2.Checked = True Then
            Call AtotoInsert()
        End If

    End Sub



    Private Sub ComboBox1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.MouseEnter
        ComboBox1.BackColor = Color.White
    End Sub

    Public Sub loadSQL()
        'sql = ""

        'If TextBox25.Text <> "" Then
        '    sql = " AND Bnk_Ac_Code = '" & TextBox25.Text & "' "
        'ElseIf TextBox26.Text <> "" Then
        '    sql = " AND (Bnk_Ac_Name  Like N'%" & TextBox26.Text.Trim & "%')"
        'End If

        'If ComboBox2.Text <> "===ທັງໝົດ===" Then
        '    sql = " AND Receipt_Type = N'" & TextBox25.Text & "' "
        'End If
        'If Cmb.Text <> "==ທັງຫມົດ==" Then
        '    sql = " AND Curr = '" & Cmb.Text & "' "
        'End If
        'If ComboBox3.Text = "ເຄື່ອນໄຫວແລ້ວ" Then
        sql = " AND Status = '1' "
        'End If
        'If ComboBox3.Text = "ຍັງບໍ່ທັນເຄື່ອນໄຫວ" Then
        '    sql = " AND Status = '0' "
        'End If
        ''sql = " AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "' "
    End Sub

    Private Sub TextBox25_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox25.KeyPress
        If e.KeyChar = Chr(13) Then
            sql = ""
            sql = " AND Receipt_No = '" & TextBox25.Text & "' AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
            'sql = "   AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
            LoadListFG()
        End If


    End Sub

    Private Sub TextBox26_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox26.KeyPress
        If e.KeyChar = Chr(13) Then
            sql = ""
            sql = " AND (Bnk_Ac_Name  Like N'%" & TextBox26.Text.Trim & "%')  AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
            'sql = "   AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
            LoadListFG()
        End If
    End Sub

    Private Sub DateTimePicker1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePicker1.KeyPress
        If e.KeyChar = Chr(13) Then
            sql = ""
            sql = " AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "' "

            LoadListFG()
        End If
    End Sub

  

    Private Sub DateTimePicker2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePicker2.KeyPress
        If e.KeyChar = Chr(13) Then
            sql = ""
            sql = " AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "' "
            LoadListFG()
        End If
    End Sub


    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        sql = ""
        If ComboBox3.Text = "ເຄື່ອນໄຫວແລ້ວ" Then
            sql = " AND Status = '1'  AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
        Else
            sql = "   AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"

        End If
        If ComboBox3.Text = "ຍັງບໍ່ທັນເຄື່ອນໄຫວ" Then
            sql = " AND Status = '0'  AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
        Else
            sql = "   AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
        End If
        'sql = "   AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
        LoadListFG()
    End Sub

    Private Sub TextBox25_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox25.TextChanged

    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged

    End Sub

    Private Sub TextBox26_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox26.TextChanged

    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'TextBox26.Text = Mid(TextBox26.ToString, 1, 1)
        MsgBox(Mid(ComboBox1.Text.ToString, 4, 7))
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        sql = ""
        sql = " AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "' "

        LoadListFG()
    End Sub

    Private Sub BtnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRefresh.Click
        If RadioButton1.Checked = True Then
            LoadListFGRatePaper()
            'TextBox11.ReadOnly = True

            LoadListFGRatePaper()
            Call SumData()
        Else
            Call AtotoInsert()
            'TextBox11.ReadOnly = False
            Call SumData()
            TextBox11.Focus()
        End If
    End Sub

    Private Sub TextBox11_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox11.TextChanged

    End Sub

   

    Private Sub BtnExit_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click

    End Sub
End Class