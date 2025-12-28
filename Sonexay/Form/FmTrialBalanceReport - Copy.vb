Public Class FmTrialBalanceReport
    Dim MonthLetter1 As String
    Dim MdStartDate As Date
    Dim MdToDate As Date
    Dim sql As String
    Dim VCode1, VCode2, VCode3, VCode4, VCode5, VCode6, VCode7, VCode8, VCode9 As String
    Dim RsOpen As New ADODB.Recordset
    Dim RsOpenMonth As New ADODB.Recordset
    Dim RsRpt As New ADODB.Recordset
    Dim AmtOpenDR, AmtOpenCR, AmtOpenMonthDR, AmtOpenMonthCR As Double
    Dim VOpenDate As Date
    Dim RptNme As String
    Dim RSC12 As New ADODB.Recordset
    Dim d, p As String

    Dim RSCP As New ADODB.Recordset

    Private Sub DMonth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DMonth.SelectedIndexChanged
        LoadMonth()
        Call AddData()
        Call LoadListFG()
        Dim i As Integer
        For i = 1 To FG.Rows - 1

            If CDbl(CDbl(FG.get_TextMatrix(i, 3)) + CDbl(FG.get_TextMatrix(i, 5))) - CDbl(CDbl(FG.get_TextMatrix(i, 4)) + CDbl(FG.get_TextMatrix(i, 6))) >= 0 Then
                FG.set_TextMatrix(i, 7, Format(CDbl(CDbl(FG.get_TextMatrix(i, 3)) + CDbl(FG.get_TextMatrix(i, 5))) - CDbl(CDbl(FG.get_TextMatrix(i, 4)) + CDbl(FG.get_TextMatrix(i, 6))), "##,##0.00"))
            Else
                FG.set_TextMatrix(i, 8, Format(CDbl(CDbl(FG.get_TextMatrix(i, 4)) + CDbl(FG.get_TextMatrix(i, 6))) - CDbl(CDbl(FG.get_TextMatrix(i, 3)) + CDbl(FG.get_TextMatrix(i, 5))), "##,##0.00"))
            End If
        Next i
    End Sub

    Private Sub LoadMonth()
        '---------------------------------
        If DMonth.Text = "ມັງກອນ" Then
            MdStartDate = Format(CDate("01/01/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/01/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ມັງກອນ"
        ElseIf DMonth.Text = "ກຸມພາ" Then
            Dim Day As String
            Dim MM As Date
            Dim Fromm As Date
            MdStartDate = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
            Fromm = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
            MM = Format(CDate("01/3/" & Year(MdStartDate)), "dd-MM-yyyy")
            Day = DateDiff(DateInterval.Day, Fromm, MM)
            MdToDate = Format(CDate(Day & "/02" & "/" & Year(MdStartDate)), "dd-MM-yyyy")
            MonthLetter1 = "ກຸມພາ"
            Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
        ElseIf DMonth.Text = "ມີນາ" Then
            MdStartDate = Format(CDate("01/03/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/03/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ມີນາ"
        ElseIf DMonth.Text = "ເມສາ" Then
            MdStartDate = Format(CDate("01/04/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/04/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ເມສາ"
        ElseIf DMonth.Text = "ພຶດສະພາ" Then
            MdStartDate = Format(CDate("01/05/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/05/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ພຶດສະພາ"
        ElseIf DMonth.Text = "ມີຖຸນາ" Then
            MdStartDate = Format(CDate("01/06/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/06/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ມີຖຸນາ"
        ElseIf DMonth.Text = "ກໍລະກົດ" Then
            MdStartDate = Format(CDate("01/07/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/07/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ກໍລະກົດ"
        ElseIf DMonth.Text = "ສິງຫາ" Then
            MdStartDate = Format(CDate("01/08/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/08/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ສິງຫາ"
        ElseIf DMonth.Text = "ກັນຍາ" Then
            MdStartDate = Format(CDate("01/09/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/09/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ກັນຍາ"
        ElseIf DMonth.Text = "ຕຸລາ" Then
            MdStartDate = Format(CDate("01/10/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/10/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ຕຸລາ"
        ElseIf DMonth.Text = "ພະຈິກ" Then
            MdStartDate = Format(CDate("01/11/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/11/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ພະຈິກ"
        ElseIf DMonth.Text = "ທັນວາ" Then
            MdStartDate = Format(CDate("01/12/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/12/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ທັນວາ"
        End If
        '-----------------
        Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub


    Private Sub Period_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Period.SelectedIndexChanged
        LoadPeriod()
    End Sub

    Private Sub LoadPeriod()

        If Period.Text = "ງວດທີ 1" Then
            MdStartDate = Format(CDate("01/01/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/03/" & Year(Pyy.Value)), "dd-MM-yyyy")
            Lb.Text = "ປະຈຳງວດ " & "1" & " ປີ " & Pyy.Text
        ElseIf Period.Text = "ງວດທີ 2" Then
            MdStartDate = Format(CDate("01/04/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/06/" & Year(Pyy.Value)), "dd-MM-yyyy")
            Lb.Text = "ປະຈຳງວດ " & "2" & " ປີ " & Pyy.Text
        ElseIf Period.Text = "ງວດທີ 3" Then
            MdStartDate = Format(CDate("01/07/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/09/" & Year(Pyy.Value)), "dd-MM-yyyy")
            Lb.Text = "ປະຈຳງວດ " & "3" & " ປີ " & Pyy.Text
        ElseIf Period.Text = "ງວດທີ 4" Then
            MdStartDate = Format(CDate("01/10/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/12/" & Year(Pyy.Value)), "dd-MM-yyyy")
            Lb.Text = "ປະຈຳງວດ " & "4" & " ປີ " & Pyy.Text
        End If
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub

    Private Sub Ds_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ds.ValueChanged

        Dt.Text = Ds.Text
        LoadDay()
    End Sub
    Private Sub LoadDay()
        MdStartDate = Format(CDate(Ds.Value), "dd-MM-yyyy")
        MdToDate = Format(CDate(Dt.Value), "dd-MM-yyyy")
        Lb.Text = "ປະຈຳວັນທີ " & MdStartDate & " ຫາວັນທີ " & MdToDate
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub

    Private Sub DateTimePicker5_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yy.ValueChanged
        Call LoadYear()
    End Sub

    Private Sub LoadYear()
        MdStartDate = Format(CDate("01/1/" & Year(yy.Value)), "dd-MM-yyyy")
        MdToDate = Format(CDate("31/12/" & Year(yy.Value)), "dd-MM-yyyy")
        Lb.Text = "ປະຈຳປີ " & yy.Text
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub

    Private Sub Loadsql()

        sql = ""

        'sql = Microsoft.VisualBasic.Left()

        sql = " AND GIN.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' "

        'sql = " AND Cust_ID = '" & Microsoft.VisualBasic.Left(ComboBox1.Text, 2) & "' "



    End Sub

    Private Sub FmTrialBalanceReport_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Num = Num - 1
        Call MdiCNum()
    End Sub
    Private Sub FmTrialBalanceReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FG.AllowUserResizing = VSFlex8U.AllowUserResizeSettings.flexResizeBoth
        Period.Text = "ງວດທີ 1"
        DMonth.Text = "ມັງກອນ"
        LoadMonth()
        FG.FormatString = "^ລ/ດ |<ລະຫັດບັນຊີ ||ຍອດຍົກເບື້ອງ (ຫນີ້) |ຍອດຍົກເບື້ອງ (ມີ) | ການເຄື່ອນໄຫວ (ຫນີ້) | ການເຄື່ອນໄຫວ (ມີ) | ຍອດເຫລືອ (ຫນີ້)     | ຍອດເຫລືອ (ມີ)      "
        FG.ExtendLastCol = True



        RD.Checked = True
        Ds.Value = MWorkSetting
        Myy.Value = MWorkSetting
        yy.Value = MWorkSetting

        Pyy.Value = MWorkSetting
        If Format(MWorkSetting, "MM") = 1 Then
            DMonth.SelectedIndex = 0
            Period.SelectedIndex = 0
        ElseIf Format(MWorkSetting, "MM") = 2 Then
            DMonth.SelectedIndex = 1
            Period.SelectedIndex = 0
        ElseIf Format(MWorkSetting, "MM") = 3 Then
            DMonth.SelectedIndex = 2
            Period.SelectedIndex = 0
        ElseIf Format(MWorkSetting, "MM") = 4 Then
            DMonth.SelectedIndex = 3
            Period.SelectedIndex = 1
        ElseIf Format(MWorkSetting, "MM") = 5 Then
            DMonth.SelectedIndex = 4
            Period.SelectedIndex = 1
        ElseIf Format(MWorkSetting, "MM") = 6 Then
            DMonth.SelectedIndex = 5
            Period.SelectedIndex = 1
        ElseIf Format(MWorkSetting, "MM") = 7 Then
            DMonth.SelectedIndex = 6
            Period.SelectedIndex = 2
        ElseIf Format(MWorkSetting, "MM") = 8 Then
            DMonth.SelectedIndex = 7
            Period.SelectedIndex = 2
        ElseIf Format(MWorkSetting, "MM") = 9 Then
            DMonth.SelectedIndex = 8
            Period.SelectedIndex = 2
        ElseIf Format(MWorkSetting, "MM") = 10 Then
            DMonth.SelectedIndex = 9
            Period.SelectedIndex = 3
        ElseIf Format(MWorkSetting, "MM") = 11 Then
            DMonth.SelectedIndex = 10
            Period.SelectedIndex = 3
        ElseIf Format(MWorkSetting, "MM") = 12 Then
            DMonth.SelectedIndex = 11
            Period.SelectedIndex = 3
        End If

        selectLoad()
        'Call AddData()
        'Call LoadListFG()
        'Call CaRemain()
    End Sub
    Private Sub selectLoad()
        DMonth.Enabled = False
        Myy.Enabled = False
        Period.Enabled = False
        Pyy.Enabled = False
        Ds.Enabled = False
        Dt.Enabled = False
        yy.Enabled = False
        If RM.Checked = True Then
            DMonth.Enabled = True
            Myy.Enabled = True
            LoadMonth()
        ElseIf RP.Checked = True Then
            Period.Enabled = True
            Pyy.Enabled = True
            LoadPeriod()
        ElseIf RD.Checked = True Then
            Ds.Enabled = True
            Dt.Enabled = True
            LoadDay()
        ElseIf RY.Checked = True Then
            yy.Enabled = True
            LoadYear()
        End If
    End Sub

    Private Sub Myy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Myy.ValueChanged
        LoadMonth()
    End Sub

    Private Sub RM_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RM.CheckedChanged
        selectLoad()
    End Sub

    Private Sub Pyy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Pyy.ValueChanged
        LoadPeriod()
    End Sub

    Private Sub Dt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Dt.ValueChanged
        LoadDay()
    End Sub

    Private Sub RP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RP.CheckedChanged
        selectLoad()

    End Sub

    Private Sub RD_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RD.CheckedChanged
        selectLoad()
    End Sub

    Private Sub RY_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RY.CheckedChanged
        selectLoad()
    End Sub

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Close()
        'If DMonth.Text <> DMonth.Items Then

        'End If
    End Sub

    Private Sub BtnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        Call AddData()
        Call LoadReport()
    End Sub
    Private Sub LoadReport()
        '  ,0 as off_name 
        '  ,0 as off_tel 
        '  ,0 as off_Place
        If RM.Checked = True Then
            RptNme = "ໃບດູນດ່ຽງປະຈຳເດືອນ"
        ElseIf RP.Checked = True Then
            RptNme = "ໃບດູນດ່ຽງປະງວດ"
        ElseIf RD.Checked = True Then
            RptNme = "ໃບດູນດ່ຽງປະຈຳວັນ"
        ElseIf RY.Checked = True Then
            RptNme = "ໃບດູນດ່ຽງປະປີ"
        End If
        If CheckBox1.Checked = True Then
            RptNme = RptNme & CheckBox1.Text
        End If
        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()

            If RaParent.Checked = True Then
                .Open("SELECT * ,N'" & Lb.Text & "'  as Report_Date , N'" & RptNme & "'  as Report_name FROM Ap_balance_6_ChangParent Order by ac_code ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

            Else
                .Open("SELECT * , N'" & Lb.Text & "'  as Report_Date , N'" & RptNme & "'  as Report_name FROM Ap_balance_6_Col Order by ac_code ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

            End If

            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With

        Dim FrmPreview As New FmPreview
        Dim Rpt As New CryTrialBalanceReport_6
        Rpt.SetDataSource(Rs)
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.DisplayGroupTree = False
        FrmPreview.MdiParent = FmMain
        FrmPreview.WindowState = FormWindowState.Maximized
        FrmPreview.Show()
        FrmPreview.Focus()
    End Sub
    Private Sub LoadOpen_Jn()
        CNN.Execute("DELETE FROM Ap_balance_6_col WHERE cnt <> '" & "" & "'")
        CNN.Execute("DELETE FROM Ap_balance_6 WHERE cnt <> '" & "" & "'")
        LoadOpen_Jn1()
        LoadOpen_Jn2()
        LoadOpen_Jn3()
        LoadOpen_Jn6()
        LoadOpen_Jn7()
        LoadOpen_Jn11()
        LoadOpen_Jn14_1()
        If CheckBox1.Checked = True Then
            LoadOpen_Jn16()
        End If
        LoadOpen_Jn14_1()


        LoadOpen_Jn15()
        LoadRaParent()

    End Sub
    Private Sub LoadOpen_Jn16()
        Dim RSC16 As New ADODB.Recordset
        LoadSqlData("  select sum(rem_cr - rem_dr) as x  , count(rem_cr - rem_dr) as y  from Ap_balance_6_col  where Ac_Code Like '1%' or Ac_Code Like '2%' or Ac_Code Like '3%' ", RSC16)

        If CDbl(Trim(RSC16.Fields("y").Value)) > 0 Then

            If CDbl(Trim(RSC16.Fields("x").Value)) > 0 Then
                Rem_Cr = CDbl(Trim(RSC16.Fields("x").Value))
                Rem_Dr = 0
            End If
            If CDbl(Trim(RSC16.Fields("x").Value)) < 0 Then
                Rem_Dr = CDbl(Trim(RSC16.Fields("x").Value)) * CDbl(-1)
                Rem_Cr = 0
            End If
            If CDbl(Trim(RSC16.Fields("x").Value)) <> 0 Then
                Call LoadOpen_Jn17()
            End If
        End If


    End Sub
    Private Sub LoadOpen_Jn17()
        Dim RSC17 As New ADODB.Recordset
        LoadSqlData("   select Ac_Code from Ap_balance_6_col  where Ac_Code ='65'", RSC17)
        If RSC17.RecordCount <> 0 Then
            CNN.Execute(" Update Ap_balance_6_col set Amt_Dr = " & CDbl(Rem_Dr) & "  , Amt_Cr =" & CDbl(Rem_Cr) & " ")
        Else
            CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code ,ac_name , ac_namee , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
             "Values('65', N'" & "***" & "', '" & "***" & "', " & _
             " " & CDbl(0) & ", " & CDbl(0) & ", " & CDbl(Rem_Dr) & ", " & CDbl(Rem_Cr) & ",0 )")
        End If

        CNN.Execute(" Delete Ap_balance_6_col  where Ac_Code Like '1%' or Ac_Code Like '2%' or Ac_Code Like '3%'  ")

    End Sub
    Private Sub LoadOpen_Jn1()
        Dim RSC12 As New ADODB.Recordset
        LoadSqlData("   select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & MULook & " group BY ac_code ", RSC12)
        With RSC12
            Do Until .EOF = True
                VCode1 = CStr(Trim(.Fields("ac_Code").Value))
                CNN.Execute("INSERT INTO Ap_balance_6_col( ac_code   , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
                 "Values('" & CStr(Trim(.Fields("ac_Code").Value)) & "', " & _
                 " " & CDbl(0) & ", " & CDbl(0) & ", " & CDbl(Trim(.Fields("amt_dr").Value)) & ", " & CDbl(Trim(.Fields("amt_cr").Value)) & ",0 )")
                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub LoadOpen_Jn2()
        Dim RSC12 As New ADODB.Recordset
        Dim S As Date
        S = MdStartDate
        S = DateAdd("d", CDbl(-1), MdStartDate)
        LoadSqlData("   select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "'  " & MULook & " group BY ac_code ", RSC12)
        With RSC12
            Do Until .EOF = True
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
               "Values('" & CStr(Trim(.Fields("ac_Code").Value)) & "', " & _
               " " & CDbl(0) & ", " & CDbl(0) & ", " & CDbl(Trim(.Fields("amt_dr").Value)) & ", " & CDbl(Trim(.Fields("amt_cr").Value)) & ",0 )")
                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub LoadOpen_Jn3()
        Dim RSC3 As New ADODB.Recordset
        LoadSqlData("select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr from Open_jn WHERE    date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & MULook & " group BY ac_code", RSC3)
        With RSC3
            Do Until .EOF = True
                VCode3 = (.Fields("ac_Code").Value)
                CNN.Execute("Update Ap_balance_6 set  open_amt_dr='" & CDbl((.Fields("amt_dr").Value)) & "' , open_amt_cr='" & CDbl((.Fields("amt_cr").Value)) & "' where ac_code = '" & (.Fields("ac_Code").Value) & "'")
                LoadOpen_Jn4()
                .MoveNext()
            Loop
        End With
    End Sub


    Private Sub LoadOpen_Jn4()
        Dim RSC4 As New ADODB.Recordset
        With RSC
            LoadSqlData("select ac_Code , amt_dr , amt_cr  from Ap_balance_6  WHERE     ac_code='" & VCode3 & "'  ", RSC4)
            If RSC4.RecordCount > 0 Then
                VCode4 = (RSC4.Fields("ac_Code").Value)
                'MsgBox(VCode4)
                'CNN.Execute("Update Ap_balance_6 set  open_amt_dr=" & CDbl((RSC4.Fields("amount_dr").Value)) & " , open_amt_cr=" & CDbl((RSC4.Fields("amount_dr").Value)) & " where ac_code = '" & (.Fields("ac_code").Value) & "'")
            Else
                'VCode4 = (RSC4.Fields("ac_Code").Value)
                'MsgBox(VCode3 & "n")
                LoadOpen_Jn5()
            End If
        End With
    End Sub

    Private Sub LoadOpen_Jn5()
        Dim RSC5 As New ADODB.Recordset
        With RSC
            LoadSqlData("select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr from Open_jn  WHERE    ac_code='" & VCode3 & "' " & MULook & " group BY ac_code", RSC5)
            If RSC5.RecordCount > 0 Then
                'MsgBox(CStr(Trim(RSC5.Fields("amt_cr").Value)))
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
             "Values('" & CStr(Trim(RSC5.Fields("ac_Code").Value)) & "',  " & _
             " " & CDbl(RSC5.Fields("amt_dr").Value) & ", " & CDbl(RSC5.Fields("amt_cr").Value) & ", " & CDbl(0) & ", " & CDbl(0) & ",0 )")
            Else
                'VCode4 = CStr(Trim(.Fields("ac_Code").Value))
                'MsgBox(VCode3 & "n")
            End If
        End With
    End Sub



    Private Sub LoadOpen_Jn6()
        Dim RSC6 As New ADODB.Recordset
        Dim op_dr, op_cr, amt_dr, amt_cr As Double
        op_dr = 0
        op_cr = 0
        amt_dr = 0
        amt_cr = 0
        LoadSqlData("select Ac_Code , open_amt_dr , open_amt_cr , Amt_dr , Amt_cr from Ap_balance_6  ", RSC6)
        With RSC6
            Do Until .EOF = True
                op_dr = CDbl((.Fields("open_amt_dr").Value))
                op_cr = CDbl((.Fields("open_amt_cr").Value))
                amt_dr = CDbl((.Fields("Amt_dr").Value))
                amt_cr = CDbl((.Fields("Amt_cr").Value))

                If CDbl(op_dr + amt_dr) - CDbl(op_cr + amt_cr) >= 0 Then

                    'MsgBox((.Fields("open_amt_dr").Value) & "++++" & (.Fields("ac_code").Value))
                    CNN.Execute("Update Ap_balance_6 set rem_dr='" & CDbl(op_dr + amt_dr) - CDbl(op_cr + amt_cr) & "' , rem_cr='" & CDbl(0) & "' where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                    'MsgBox(.Fields("open_amt_dr").Value)
                End If
                If CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) >= 0 Then
                    CNN.Execute("Update Ap_balance_6 set rem_dr='" & CDbl(0) & "' , rem_cr='" & CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) & "' where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                End If
                'If CDbl(op_cr + amt_cr) = CDbl(op_dr + amt_dr) Then
                '    'CNN.Execute("Update Ap_balance_6 set rem_dr='" & CDbl(0) & "' , rem_cr='" & CDbl(0) & "' where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                '    CNN.Execute("delete Ap_balance_6_col  where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                'End If
                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub LoadOpen_Jn7()
        Dim RSC7 As New ADODB.Recordset
        LoadSqlData("select ac_code , rem_dr  , rem_cr from Ap_balance_6   ", RSC7)
        With RSC7
            Do Until .EOF = True
                VCode7 = (.Fields("ac_Code").Value)

                CNN.Execute("Update Ap_balance_6_col set  open_amt_dr='" & CDbl((.Fields("rem_dr").Value)) & "' , open_amt_cr='" & CDbl((.Fields("rem_cr").Value)) & "' where ac_code = '" & (.Fields("ac_Code").Value) & "'")
                LoadOpen_Jn8()
                .MoveNext()
            Loop
        End With
    End Sub


    Private Sub LoadOpen_Jn8()
        Dim RSC8 As New ADODB.Recordset
        With RSC
            LoadSqlData("select ac_code   from Ap_balance_6_col  WHERE     ac_code='" & VCode7 & "' ", RSC8)
            If RSC8.RecordCount > 0 Then
                VCode8 = (RSC8.Fields("ac_Code").Value)
            Else
                LoadOpen_Jn9()
            End If
        End With
    End Sub


    Private Sub LoadOpen_Jn9()
        Dim RSC9 As New ADODB.Recordset
        With RSC9
            LoadSqlData("select ac_code , Rem_dr , Rem_cr from Ap_balance_6  WHERE    ac_code='" & VCode7 & "' ", RSC9)
            If RSC9.RecordCount > 0 Then
                CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code ,ac_name , ac_namee , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
                "Values('" & CStr(Trim(RSC9.Fields("ac_Code").Value)) & "', N'" & "***" & "', '" & "***" & "', " & _
                " " & CDbl(RSC9.Fields("rem_dr").Value) & ", " & CDbl(RSC9.Fields("rem_cr").Value) & ", " & CDbl(0) & ", " & CDbl(0) & ",0 )")
            Else
            End If
        End With
    End Sub
    Private Sub LoadOpen_Jn14_1()


        Dim RSC14_1 As New ADODB.Recordset
        LoadSqlData("select Ac_Code , open_amt_dr ,open_amt_cr , amt_dr , amt_cr  from Ap_balance_6_col  ", RSC14_1)
        With RSC14_1
            Do Until .EOF = True
                'MsgBox(CDbl(CDbl(CDbl((.Fields("open_amt_dr").Value)) + CDbl((.Fields("amt_dr").Value))) - CDbl(CDbl((.Fields("amt_cr").Value)) + CDbl((.Fields("amt_cr").Value)))) & " where Ac_Code = '" & (.Fields("Ac_Code").Value))

                If CDbl(CDbl((.Fields("open_amt_dr").Value)) + CDbl((.Fields("amt_dr").Value))) >= CDbl(CDbl((.Fields("open_amt_cr").Value)) + CDbl((.Fields("amt_cr").Value))) Then
                    CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr=" & CDbl(CDbl(CDbl((.Fields("open_amt_dr").Value)) + CDbl((.Fields("amt_dr").Value))) - CDbl(CDbl((.Fields("open_amt_cr").Value)) + CDbl((.Fields("amt_cr").Value)))) & " where Ac_Code = '" & (.Fields("Ac_Code").Value) & "'")
                Else
                    CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr=" & CDbl(CDbl(CDbl((.Fields("open_amt_cr").Value)) + CDbl((.Fields("amt_cr").Value))) - CDbl(CDbl((.Fields("open_amt_dr").Value)) + CDbl((.Fields("amt_dr").Value)))) & " where Ac_Code = '" & (.Fields("Ac_Code").Value) & "'")
                End If
                .MoveNext()
            Loop
        End With
    End Sub
    Private Sub LoadOpen_Jn15()
        CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
    End Sub
    Private Sub Pr()
        For i = 1 To Len(d)
            If Mid(d, i, 1) = "." Then
                p = Microsoft.VisualBasic.Left(d, i - 1)
                Exit Sub
            Else
                p = d
            End If
        Next i
    End Sub
    Private Sub LoadRaParent()

        If RaParent.Checked = True Then

            Dim RSCpl12 As New ADODB.Recordset
     
            LoadSqlData("SELECT  Ac_Code  FROM   Ap_balance_6_col  ", RSCpl12)
            With RSCpl12
                Do Until .EOF = True
                    d = (.Fields("Ac_Code").Value)
                    Call Pr()
                    CNN.Execute("Update Ap_balance_6_col set  Acc_Parent = '" & p & "'  where ac_code='" & (.Fields("ac_code").Value) & "'")
                    .MoveNext()
                Loop
            End With

            LoadPl2()
         

        End If

    End Sub

  

    Private Sub LoadPl2()

        CNN.Execute("Delete Ap_balance_6_ChangParent")
        CNN.Execute("insert Into Ap_balance_6_ChangParent (Ac_Code   ,  open_amt_dr ,  open_amt_cr   , amt_dr , amt_cr      ) select Acc_Parent    , sum(open_amt_dr) as open_amt_dr  , sum(open_amt_cr) as open_amt_cr   ,  sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr       from Ap_balance_6_col group by Acc_Parent ")

        LoadSqlData("select  Ac_Code , open_amt_dr , open_amt_cr from Ap_balance_6_ChangParent   ", RSCP)

        With RSCP
            Do Until .EOF = True
                If CDbl((.Fields("open_amt_dr").Value)) >= CDbl((.Fields("open_amt_cr").Value)) Then
                    ''MsgBox(CDbl((.Fields("open_amt_dr").Value)) & "---" & CDbl((.Fields("open_amt_cr").Value)) & "==" & "9999999999" & "==--" & ((.Fields("Ac_Code").Value)))
                    CNN.Execute("Update Ap_balance_6_ChangParent set open_amt_dr = '" & CDbl((.Fields("open_amt_dr").Value)) - CDbl((.Fields("open_amt_cr").Value)) & "' , open_amt_cr=0  where ac_code='" & (.Fields("ac_code").Value) & "'")

                ElseIf CDbl((.Fields("open_amt_dr").Value)) <= CDbl((.Fields("open_amt_cr").Value)) Then
                    ''MsgBox(CDbl((.Fields("open_amt_dr").Value)) & "---" & CDbl((.Fields("open_amt_cr").Value)) & "==" & CDbl((.Fields("open_amt_cr").Value)) & "==--" & ((.Fields("Ac_Code").Value)))
                    CNN.Execute("Update Ap_balance_6_ChangParent set open_amt_dr=0 , open_amt_cr = '" & CDbl((.Fields("open_amt_cr").Value)) - CDbl((.Fields("open_amt_dr").Value)) & "'   where ac_code='" & (.Fields("ac_code").Value) & "'")

                End If
                .MoveNext()
            Loop
        End With

        LoadSqlData("select  Ac_Code , open_amt_dr , open_amt_cr  , amt_dr , amt_cr from Ap_balance_6_ChangParent   ", RSCP)
        With RSCP
            Do Until .EOF = True
         
                If CDbl(CDbl((.Fields("open_amt_dr").Value)) + CDbl((.Fields("amt_dr").Value))) >= CDbl(CDbl((.Fields("open_amt_cr").Value)) + CDbl((.Fields("amt_cr").Value))) Then
                    CNN.Execute("Update  Ap_balance_6_ChangParent set Rem_cr=0 , Rem_dr=" & CDbl(CDbl(CDbl((.Fields("open_amt_dr").Value)) + CDbl((.Fields("amt_dr").Value))) - CDbl(CDbl((.Fields("open_amt_cr").Value)) + CDbl((.Fields("amt_cr").Value)))) & " where Ac_Code = '" & (.Fields("Ac_Code").Value) & "'")
                Else
                    CNN.Execute("Update  Ap_balance_6_ChangParent set Rem_dr=0 , Rem_cr=" & CDbl(CDbl(CDbl((.Fields("open_amt_cr").Value)) + CDbl((.Fields("amt_cr").Value))) - CDbl(CDbl((.Fields("open_amt_dr").Value)) + CDbl((.Fields("amt_dr").Value)))) & " where Ac_Code = '" & (.Fields("Ac_Code").Value) & "'")
                End If
                .MoveNext()
            Loop
        End With


        Dim RSCbb As New ADODB.Recordset

        LoadSqlData("SELECT  Acc_Code.Ac_Code AS Ac_Code, Acc_Code.Name_L AS Name_L FROM   Acc_Code INNER JOIN    Ap_balance_6_ChangParent ON Acc_Code.Ac_Code = Ap_balance_6_ChangParent.ac_code  ", RSCbb)
        With RSCbb
            Do Until .EOF = True
                CNN.Execute("Update Ap_balance_6_ChangParent set ac_name = N'" & (.Fields("Name_L").Value) & "'  where ac_code='" & (.Fields("ac_code").Value) & "'")
                .MoveNext()
            Loop
        End With

        'Call LoadRaParent2()

    End Sub


    Private Sub LoadPr1()
        Dim RSC1 As New ADODB.Recordset
        LoadSqlData("select * from Ap_balance_6_col where Ac_Code <> '" & (RSCP.Fields("Acc_Parent").Value) & "' ", RSC1)
        'If RSC.RecordCount Then
        With RSC1
            Do Until .EOF = True
                'MsgBox((RSC1.Fields("Acc_Parent").Value))
                'CNN.Execute("Update Ap_balance_6_col set ac_name = N'" & (.Fields("Name_L").Value) & "'   where ac_code='" & (.Fields("ac_code").Value) & "'")
                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub LoadOpen_Jn11()
        Dim RSC11 As New ADODB.Recordset
        Dim op_dr11, op_cr11, amt_dr11, amt_cr11 As Double
        op_dr11 = 0
        op_cr11 = 0
        amt_dr11 = 0
        amt_cr11 = 0
        LoadSqlData("select Ac_Code , open_amt_dr , open_amt_cr , Amt_dr , Amt_cr from Ap_balance_6_col  ", RSC11)
        With RSC11
            Do Until .EOF = True
                op_dr11 = CDbl((.Fields("open_amt_dr").Value))
                op_cr11 = CDbl((.Fields("open_amt_cr").Value))
                amt_dr11 = CDbl((.Fields("Amt_dr").Value))
                amt_cr11 = CDbl((.Fields("Amt_cr").Value))
                If CDbl(op_dr11 + op_cr11) = 0 Then
                    If CDbl(amt_dr11 + amt_cr11) = 0 Then
                        CNN.Execute("delete Ap_balance_6_col  where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                    End If

                End If
                .MoveNext()
            Loop
        End With
    End Sub
    Private Sub AddData()

        LoadOpen_Jn()

    End Sub

    Private Sub BtnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRefresh.Click
        Call AddData()
        Call LoadListFG()
        Call CaRemain()
    End Sub
    Private Sub CaRemain()
        Dim i As Integer
        For i = 1 To FG.Rows - 1

            If CDbl(CDbl(FG.get_TextMatrix(i, 3)) + CDbl(FG.get_TextMatrix(i, 5))) - CDbl(CDbl(FG.get_TextMatrix(i, 4)) + CDbl(FG.get_TextMatrix(i, 6))) >= 0 Then
                FG.set_TextMatrix(i, 7, Format(CDbl(CDbl(FG.get_TextMatrix(i, 3)) + CDbl(FG.get_TextMatrix(i, 5))) - CDbl(CDbl(FG.get_TextMatrix(i, 4)) + CDbl(FG.get_TextMatrix(i, 6))), "##,##0.00"))
            Else
                FG.set_TextMatrix(i, 8, Format(CDbl(CDbl(FG.get_TextMatrix(i, 4)) + CDbl(FG.get_TextMatrix(i, 6))) - CDbl(CDbl(FG.get_TextMatrix(i, 3)) + CDbl(FG.get_TextMatrix(i, 5))), "##,##0.00"))
            End If
        Next i
    End Sub

    Private Sub LoadListFG()
        Dim O_dr, O_cr, Amt_dr, Amt_cr, R_dr, R_Cr As Double
        FG.Rows = 1
        With RSC
            If RaParent.Checked = True Then
                Call LoadSqlData("SELECT * FROM  Ap_balance_6_ChangParent Order by ac_code", RSC)
            Else
                Call LoadSqlData("SELECT * FROM  Ap_balance_6_col Order by ac_code", RSC)
            End If

            If .RecordCount > 0 Then
                While Not .EOF

                    O_dr = Trim(CDbl(.Fields("open_amt_dr").Value))
                    O_cr = Trim(CDbl(.Fields("open_amt_cr").Value))
                    Amt_dr = Trim(CDbl(.Fields("amt_dr").Value))
                    'MsgBox(Trim(CDbl(.Fields("amt_cr").Value)))
                    Amt_cr = Trim(CDbl(.Fields("amt_cr").Value))
                    R_dr = "0,00"
                    R_Cr = "0,00"

                    FG.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("ac_code").Value)) & _
                                "" & vbTab & Trim(CStr(.Fields("ac_name").Value)) & _
                                        "" & vbTab & Format(O_dr, "#,##0.00") & _
                                         "" & vbTab & Format(O_cr, "#,##0.00") & _
                                          "" & vbTab & Format(Amt_dr, "#,##0.00") & _
                                             "" & vbTab & Format(Amt_cr, "#,##0.00") & _
                                                 "" & vbTab & R_dr & _
                                             "" & vbTab & R_Cr & _
                                            "" & vbTab & ((.Fields("cnt").Value)))
                    .MoveNext()
                End While
            Else
                FG.Rows = 2
            End If
        End With
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub








    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        LoadOpen_Jn3()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub BtnExit_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Call Close()
    End Sub

    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelChange

    End Sub
End Class
