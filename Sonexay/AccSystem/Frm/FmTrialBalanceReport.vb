Public Class FmTrialBalanceReport
    Dim MonthLetter1 As String


    'Dim MdStartDate As Date
    'Dim MdToDate As Date
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

 
    Private Sub HeaDer()
        LoadSqlData("SELECT * FROM Header WHERE ID=N'BL01' ", RSC)
        If RSC.RecordCount <> 0 Then
            If MuLng = "L" Then
                TxtHeader.Text = Trim(RSC.Fields("Nm").Value.ToString)
                TxtS1.Text = Trim(RSC.Fields("S1").Value.ToString)
                TxtS2.Text = Trim(RSC.Fields("S2").Value.ToString)
                TxtS3.Text = Trim(RSC.Fields("S3").Value.ToString)
                TxtS4.Text = Trim(RSC.Fields("S4").Value.ToString)
                TxtPP.Text = Trim(RSC.Fields("pp").Value.ToString)
            Else
                TxtHeader.Text = Trim(RSC.Fields("Nm").Value.ToString)
                TxtS1.Text = Trim(RSC.Fields("S1e").Value.ToString)
                TxtS2.Text = Trim(RSC.Fields("S2e").Value.ToString)
                TxtS3.Text = Trim(RSC.Fields("S3e").Value.ToString)
                TxtS4.Text = Trim(RSC.Fields("S4e").Value.ToString)
                TxtPP.Text = Trim(RSC.Fields("ppe").Value.ToString)
            End If
        End If
    End Sub

    Private Sub AddHeader()
        If MuLng = "L" Then
            LoadSqlData("SELECT * FROM Header WHERE ID=N'BL01' ", RSC)
            If RSC.RecordCount = 0 Then
                CNN.Execute("INSERT INTO Header(ID,Nm,S1,S2,S3,S4,PP) " & _
                            " values('BL01',N'" & TxtHeader.Text & "',N'" & TxtS1.Text & "',N'" & TxtS2.Text & "',N'" & TxtS3.Text & "',N'" & TxtS4.Text & "',N'" & TxtPP.Text & "') ")
            Else
                CNN.Execute("UPDATE Header set Nm=N'" & TxtHeader.Text & "',S1=N'" & TxtS1.Text & "',S2=N'" & TxtS2.Text & "',S3=N'" & TxtS3.Text & "',S4=N'" & TxtS4.Text & "',PP=N'" & TxtPP.Text & "' " & _
                            " where ID='BL01' ")
            End If
        Else
            LoadSqlData("SELECT * FROM Header WHERE ID=N'BL01' ", RSC)
            If RSC.RecordCount = 0 Then
                CNN.Execute("INSERT INTO Header(ID,Nm,S1e,S2e,S3e,S4e,PPe) " & _
                            " values('BL01',N'" & TxtHeader.Text & "',N'" & TxtS1.Text & "',N'" & TxtS2.Text & "',N'" & TxtS3.Text & "',N'" & TxtS4.Text & "',N'" & TxtPP.Text & "') ")
            Else
                CNN.Execute("UPDATE Header set Nm=N'" & TxtHeader.Text & "',S1=N'" & TxtS1.Text & "',S2=N'" & TxtS2.Text & "',S3=N'" & TxtS3.Text & "',S4=N'" & TxtS4.Text & "',PP=N'" & TxtPP.Text & "' " & _
                            " where ID='BL01' ")
            End If
        End If
       
    End Sub

    Private Sub DMonth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DMonth.SelectedIndexChanged
        LoadMonth()

        CMB_Curr_SelectedIndexChanged(sender, e)
    End Sub

    Private Sub LoadMonth()
        '---------------------------------

        If DMonth.SelectedIndex = 0 Then

        End If

        'If DMonth.SelectedIndex = 0 Then
        '    MdStartDate = Format(CDate("01/01/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("31/01/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    LngId = "7013" : CallLngStr() : MonthLetter1 = LngStr
        '    MonthLetter1 = "ມັງກອນ"
        'ElseIf DMonth.SelectedIndex = 1 Then
        '    Dim Day As String
        '    Dim MM As Date
        '    Dim Fromm As Date
        '    MdStartDate = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
        '    Fromm = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
        '    MM = Format(CDate("01/3/" & Year(MdStartDate)), "dd-MM-yyyy")
        '    Day = DateDiff(DateInterval.Day, Fromm, MM)
        '    MdToDate = Format(CDate(Day & "/02" & "/" & Year(MdStartDate)), "dd-MM-yyyy")
        '    MonthLetter1 = "ກຸມພາ"
        '    LngId = "7014" : CallLngStr() : MonthLetter1 = LngStr
        'ElseIf DMonth.SelectedIndex = 2 Then
        '    MdStartDate = Format(CDate("01/03/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("31/03/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    LngId = "7015" : CallLngStr() : MonthLetter1 = LngStr
        '    MonthLetter1 = "ມີນາ"
        'ElseIf DMonth.SelectedIndex = 3 Then
        '    MdStartDate = Format(CDate("01/04/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("30/04/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    LngId = "7016" : CallLngStr() : MonthLetter1 = LngStr
        '    MonthLetter1 = "ເມສາ"
        'ElseIf DMonth.SelectedIndex = 4 Then
        '    MdStartDate = Format(CDate("01/05/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("31/05/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    LngId = "7017" : CallLngStr() : MonthLetter1 = LngStr
        '    MonthLetter1 = "ພຶດສະພາ"
        'ElseIf DMonth.SelectedIndex = 5 Then
        '    MdStartDate = Format(CDate("01/06/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("30/06/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    LngId = "7018" : CallLngStr() : MonthLetter1 = LngStr
        '    MonthLetter1 = "ມີຖຸນາ"
        'ElseIf DMonth.SelectedIndex = 6 Then
        '    MdStartDate = Format(CDate("01/07/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("31/07/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    LngId = "7019" : CallLngStr() : MonthLetter1 = LngStr
        '    MonthLetter1 = "ກໍລະກົດ"
        'ElseIf DMonth.SelectedIndex = 7 Then
        '    MdStartDate = Format(CDate("01/08/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("31/08/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    LngId = "7020" : CallLngStr() : MonthLetter1 = LngStr
        '    MonthLetter1 = "ສິງຫາ"
        'ElseIf DMonth.SelectedIndex = 8 Then
        '    MdStartDate = Format(CDate("01/09/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("30/09/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    LngId = "7021" : CallLngStr() : MonthLetter1 = LngStr
        '    MonthLetter1 = "ກັນຍາ"
        'ElseIf DMonth.SelectedIndex = 9 Then
        '    MdStartDate = Format(CDate("01/10/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("31/10/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    LngId = "7022" : CallLngStr() : MonthLetter1 = LngStr
        '    MonthLetter1 = "ຕຸລາ"
        'ElseIf DMonth.SelectedIndex = 10 Then
        '    MdStartDate = Format(CDate("01/11/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("30/11/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    LngId = "7023" : CallLngStr() : MonthLetter1 = LngStr
        '    MonthLetter1 = "ພະຈິກ"
        'ElseIf DMonth.SelectedIndex = 11 Then
        '    MdStartDate = Format(CDate("01/12/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("31/12/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    LngId = "7024" : CallLngStr() : MonthLetter1 = LngStr

        '    MonthLetter1 = "ທັນວາ"
        'End If
        '---------------------------------
        If FmMain.MnLaoLang.Checked = True Then
            If DMonth.Text = "ມັງກອນ" Then
                MdStartDate = Format(CDate("01/01/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/01/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ມັງກອນ"
                DMonth.SelectedIndex = 0
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
                DMonth.SelectedIndex = 1
                Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
            ElseIf DMonth.Text = "ມີນາ" Then
                MdStartDate = Format(CDate("01/03/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/03/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ມີນາ"
                DMonth.SelectedIndex = 2
            ElseIf DMonth.Text = "ເມສາ" Then
                MdStartDate = Format(CDate("01/04/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/04/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ເມສາ"
                DMonth.SelectedIndex = 3
            ElseIf DMonth.Text = "ພຶດສະພາ" Then
                MdStartDate = Format(CDate("01/05/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/05/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ພຶດສະພາ"
                DMonth.SelectedIndex = 4
            ElseIf DMonth.Text = "ມິຖຸນາ" Then
                MdStartDate = Format(CDate("01/06/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/06/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ມິຖຸນາ"
                DMonth.SelectedIndex = 5
            ElseIf DMonth.Text = "ກໍລະກົດ" Then
                MdStartDate = Format(CDate("01/07/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/07/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ກໍລະກົດ"
                DMonth.SelectedIndex = 6
            ElseIf DMonth.Text = "ສິງຫາ" Then
                MdStartDate = Format(CDate("01/08/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/08/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ສິງຫາ"
                DMonth.SelectedIndex = 7
            ElseIf DMonth.Text = "ກັນຍາ" Then
                MdStartDate = Format(CDate("01/09/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/09/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ກັນຍາ"
                DMonth.SelectedIndex = 8
            ElseIf DMonth.Text = "ຕຸລາ" Then
                MdStartDate = Format(CDate("01/10/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/10/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ຕຸລາ"
                DMonth.SelectedIndex = 9
            ElseIf DMonth.Text = "ພະຈິກ" Then
                MdStartDate = Format(CDate("01/11/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/11/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ພະຈິກ"
                DMonth.SelectedIndex = 10
            ElseIf DMonth.Text = "ທັນວາ" Then
                MdStartDate = Format(CDate("01/12/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/12/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ທັນວາ"
                DMonth.SelectedIndex = 11
            End If
            Lb.Text = "ສຳລັບວັນທີ " & (MdToDate.Day) & " " & MonthLetter1 & " " & Year(MdToDate)
        Else

            If DMonth.Text = "January" Then
                MdStartDate = Format(CDate("01/01/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/01/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "January"
                DMonth.SelectedIndex = 0
            ElseIf DMonth.Text = "February" Then
                Dim Day As String
                Dim MM As Date
                Dim Fromm As Date
                MdStartDate = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
                Fromm = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
                MM = Format(CDate("01/3/" & Year(MdStartDate)), "dd-MM-yyyy")
                Day = DateDiff(DateInterval.Day, Fromm, MM)
                MdToDate = Format(CDate(Day & "/02" & "/" & Year(MdStartDate)), "dd-MM-yyyy")
                MonthLetter1 = "February"
                DMonth.SelectedIndex = 1
                Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
            ElseIf DMonth.Text = "March" Then
                MdStartDate = Format(CDate("01/03/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/03/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "March"
                DMonth.SelectedIndex = 2
            ElseIf DMonth.Text = "April" Then
                MdStartDate = Format(CDate("01/04/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/04/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "April"
                DMonth.SelectedIndex = 3
            ElseIf DMonth.Text = "May" Then
                MdStartDate = Format(CDate("01/05/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/05/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "May"
                DMonth.SelectedIndex = 4
            ElseIf DMonth.Text = "June" Then
                MdStartDate = Format(CDate("01/06/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/06/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "June"
                DMonth.SelectedIndex = 5
            ElseIf DMonth.Text = "July" Then
                MdStartDate = Format(CDate("01/07/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/07/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "July"
                DMonth.SelectedIndex = 6
            ElseIf DMonth.Text = "August" Then
                MdStartDate = Format(CDate("01/08/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/08/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "August"
                DMonth.SelectedIndex = 7
            ElseIf DMonth.Text = "September" Then
                MdStartDate = Format(CDate("01/09/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/09/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "September"
                DMonth.SelectedIndex = 8
            ElseIf DMonth.Text = "October" Then
                MdStartDate = Format(CDate("01/10/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/10/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "October"
                DMonth.SelectedIndex = 9
            ElseIf DMonth.Text = "November" Then
                MdStartDate = Format(CDate("01/11/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/11/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "November"
                DMonth.SelectedIndex = 10
            ElseIf DMonth.Text = "December" Then
                MdStartDate = Format(CDate("01/12/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/12/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "December"
                DMonth.SelectedIndex = 11
            End If
            Lb.Text = "For the Month Ended " & (MdToDate.Day) & " " & MonthLetter1 & " " & Year(MdToDate)
        End If

        '-----------------
        Dim m, y As String
        LngId = "7049" : CallLngStr() : m = LngStr & " "
        LngId = "7025" : CallLngStr() : y = LngStr & " "

        'Lb.Text = m & MonthLetter1 & "/" & Year(MdToDate)
        L5.Text = MdStartDate & " => " & MdToDate
        'Lb.Text = "ສຳລັບວັນທີ " & (MdToDate.Day) & " " & MonthLetter1 & " " & Year(MdToDate)
    End Sub


    Private Sub Period_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Period.SelectedIndexChanged
        LoadPeriod()
    End Sub

    Private Sub LoadPeriod()
        Dim p, y As String
        LngId = "7050" : CallLngStr() : p = LngStr & ": "
        LngId = "7025" : CallLngStr() : y = " " & LngStr & ": "
        If Period.SelectedIndex = 0 Then
            MdStartDate = Format(CDate("01/01/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/03/" & Year(Pyy.Value)), "dd-MM-yyyy")
            Lb.Text = p & CDbl(Period.SelectedIndex) + 1 & y & Pyy.Text
        ElseIf Period.SelectedIndex = 1 Then
            MdStartDate = Format(CDate("01/04/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/06/" & Year(Pyy.Value)), "dd-MM-yyyy")
            Lb.Text = p & CDbl(Period.SelectedIndex) + 1 & y & Pyy.Text
        ElseIf Period.SelectedIndex = 2 Then
            MdStartDate = Format(CDate("01/07/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/09/" & Year(Pyy.Value)), "dd-MM-yyyy")
            Lb.Text = p & CDbl(Period.SelectedIndex) + 1 & y & Pyy.Text
        ElseIf Period.SelectedIndex = 3 Then
            MdStartDate = Format(CDate("01/10/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/12/" & Year(Pyy.Value)), "dd-MM-yyyy")
            Lb.Text = p & CDbl(Period.SelectedIndex) + 1 & y & Pyy.Text
        End If
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub

    Private Sub Ds_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ds.ValueChanged

        Dt.Text = Ds.Text
        LoadDay()
    End Sub
    Private Sub LoadDay()
        Dim d, t As String
        LngId = "7051" : CallLngStr() : d = LngStr & ": "
        LngId = "7027" : CallLngStr() : t = LngStr & ": "
        MdStartDate = Format(CDate(Ds.Value), "dd-MM-yyyy")
        MdToDate = Format(CDate(Dt.Value), "dd-MM-yyyy")
        Lb.Text = d & MdStartDate & " " & t & MdToDate
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub

    Private Sub DateTimePicker5_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yy.ValueChanged
        Call LoadYear()
    End Sub

    Private Sub LoadYear()
        Dim y As String
        LngId = "7025" : CallLngStr() : y = LngStr & ": "
        MdStartDate = Format(CDate("01/1/" & Year(yy.Value)), "dd-MM-yyyy")
        MdToDate = Format(CDate("31/12/" & Year(yy.Value)), "dd-MM-yyyy")
        Lb.Text = y & yy.Text
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub

    Private Sub Loadsql()

        sql = ""

        'sql = Microsoft.VisualBasic.Left()

        sql = " AND GIN.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' "

        'sql = " AND Cust_ID = '" & Microsoft.VisualBasic.Left(ComboBox1.Text, 2) & "' "



    End Sub

    Private Sub FmTrialBalanceReport_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Call MdiCNum()
    End Sub
    Private Sub loadOffice_User()
        Off_Usr.Items.Clear()
        LoadSqlData("select sub_id , off_add2  from  Ap_office  Order by sub_id", RSC)
        With RSC
            Do Until .EOF = True
                Off_Usr.Items.Add((.Fields("sub_id").Value) & " " & (.Fields("off_add2").Value))
                .MoveNext()
            Loop
        End With
        Off_Usr.Text = FmLogin.Sub_Company.Text
    End Sub
    Private Sub FmTrialBalanceReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        HeaDer()
        Call loadOffice_User()
        Cx.SelectedIndex = 0
        BalanceType.SelectedIndex = 0
        'FG.AllowUserResizing = VSFlex8U.AllowUserResizeSettings.flexResizeBoth
        Period.Text = "ງວດທີ 1"
        DMonth.Text = "ມັງກອນ"
        LoadMonth()
        If MuLng = "L" Then
            FG.FormatString = "^ລ/ດ |<ລະຫັດບັນຊີ        |<  ເນື້ອນໃນ           |ຍອດຍົກເບື້ອງ (ຫນີ້) |ຍອດຍົກເບື້ອງ (ມີ) | ການເຄື່ອນໄຫວ (ຫນີ້) | ການເຄື່ອນໄຫວ (ມີ) | ຍອດເຫລືອ (ຫນີ້)     | ຍອດເຫລືອ (ມີ)      "
            Label5.Text = "ຍອດຍົກເບື້ອງຕົ້ນ(ຫນີ້)"
            Label6.Text = "ຍອດຍົກເບື້ອງຕົ້ນ(ມີ)"

            Label7.Text = "ເຄື່ອນໄຫວໃນເດືອນ(ຫນີ້)"
            Label8.Text = "ເຄື່ອນໄຫວໃນເດືອນ(ມີ)"

            Label9.Text = "ຍອດເຫລືອ(ຫນີ້)"
            Label17.Text = "ຍອດເຫລືອ(ມີ)"
        Else
            Label5.Text = "Open (Debit)"
            Label6.Text = "Open (Credit)"

            Label7.Text = "Move (Debit)"
            Label8.Text = "Move (Credit)"

            Label9.Text = "Balance (Debit)"
            Label17.Text = "Balance  (Credit)"

            FG.FormatString = "^No |<Code         |<  Acc Name              |Open (Debit)    |Open (Credit)    | Move (Debit)     |Move    (Credit) | Balance (Debit)     | Balance  (Credit)        "

        End If

        'FG.ExtendLastCol = True



        RD.Checked = True
        Ds.Text = MWorkSetting
        Myy.Text = MWorkSetting
        yy.Text = MWorkSetting

        Pyy.Text = MWorkSetting
        If Format(MWorkSetting, "MM") = 1 Then
            Ct.SelectedIndex = 0
            DMonth.SelectedIndex = 0
            Period.SelectedIndex = 0
        ElseIf Format(MWorkSetting, "MM") = 2 Then
            Ct.SelectedIndex = 0
            DMonth.SelectedIndex = 1
            Period.SelectedIndex = 0
        ElseIf Format(MWorkSetting, "MM") = 3 Then
            Ct.SelectedIndex = 0
            DMonth.SelectedIndex = 2
            Period.SelectedIndex = 0
        ElseIf Format(MWorkSetting, "MM") = 4 Then
            Ct.SelectedIndex = 0
            DMonth.SelectedIndex = 3
            Period.SelectedIndex = 1
        ElseIf Format(MWorkSetting, "MM") = 5 Then
            Ct.SelectedIndex = 0
            DMonth.SelectedIndex = 4
            Period.SelectedIndex = 1
        ElseIf Format(MWorkSetting, "MM") = 6 Then
            Ct.SelectedIndex = 0
            DMonth.SelectedIndex = 5
            Period.SelectedIndex = 1
        ElseIf Format(MWorkSetting, "MM") = 7 Then
            Ct.SelectedIndex = 1
            DMonth.SelectedIndex = 6
            Period.SelectedIndex = 2
        ElseIf Format(MWorkSetting, "MM") = 8 Then
            Ct.SelectedIndex = 1
            DMonth.SelectedIndex = 7
            Period.SelectedIndex = 2
        ElseIf Format(MWorkSetting, "MM") = 9 Then
            Ct.SelectedIndex = 1
            DMonth.SelectedIndex = 8
            Period.SelectedIndex = 2
        ElseIf Format(MWorkSetting, "MM") = 10 Then
            Ct.SelectedIndex = 1
            DMonth.SelectedIndex = 9
            Period.SelectedIndex = 3
        ElseIf Format(MWorkSetting, "MM") = 11 Then
            Ct.SelectedIndex = 1
            DMonth.SelectedIndex = 10
            Period.SelectedIndex = 3
        ElseIf Format(MWorkSetting, "MM") = 12 Then
            Ct.SelectedIndex = 1
            DMonth.SelectedIndex = 11
            Period.SelectedIndex = 3
        End If
        Call selectMMM()
        selectLoad()
        'Call AddData()
        'Call LoadListFG()
        'Call CaRemain()
        SetControlText(Me)

        ChgChildForm()
        RGL.Text = "ແບບທົ່ວໄປ"
        RDtail.Text = "ຕາມຫມວດບັນຊີ"
        RGroup.Text = "ສະເພາະບັນຊີແມ່"
        Button4.Text = "Export"
        CMB_Curr.Items.Clear()
        CMB_Curr.Items.Add("EQVL")
        Call load_Cmb(" SELECT Curr  FROM Curr_For_Rate WHERE (Curr='LAK' Or Curr='USD')  ORDER BY cnt ", "Curr", CMB_Curr)
        If CMB_Curr.Items.Count > 0 Then
            CMB_Curr.SelectedIndex = 0
        End If
        ChgChildForm()
        'SetControlText(Me)

        SetControlText(Me)
        CheckBox3.Text = "ໃບດູນດ່ຽງປະຈຳປີ ຫລັງການປັບປຸງ"
        If CMB_Curr.Text = "LAK" Then
            CheckBox4.Text = "ທຽບເທົ່າໂດລາ"
        ElseIf CMB_Curr.Text = "USD" Then
            CheckBox4.Text = "ທຽບເທົ່າກີບ"

        Else
            CheckBox4.Text = "ທຽບເທົ່າໂດລາ"

        End If
        Button4.Text = "Export"
        If MuLng = "L" Then
            CheckBox2.Text = "ສະເພາະລາຍການເຄືອນໄຫວ"
            RGL.Text = "ແບບທົ່ວໄປ"
            RDtail.Text = "ຕາມຫມວດບັນຊີ"
            RGroup.Text = "ສະເພາະບັນຊີແມ່"
            CheckBox3.Text = "ໃບດູນດ່ຽງປະຈຳປີ ຫລັງການປັບປຸງ"
            Label21.Text = "ລາຍເຊັນ1"
            Label3.Text = "ລາຍເຊັນ2"
            Label4.Text = "ລາຍເຊັນ3"
            Label13.Text = "ລາຍເຊັນ4"
            Label20.Text = "ທີ່"
            BtnRefresh.Text = "ເອີ້ນຂໍ້ມູນ"
            Label16.Text = "ຄ່າຜິດດ່ຽງຍອດຍົກເບື້ອງຕົ້ນ:"
            Label10.Text = "ຄ່າຜິດດ່ຽງເຄື່ອນໄຫວໃນເດືອນ:"
            Label11.Text = "ຄ່າຜິດດ່ຽງຍອດເຫລືອ:"
            CheckBox4.Text = "ທຽບເທົ່າເງິນ"

        Else
            CheckBox4.Text = "LAK Prev"
            Label16.Text = "Open Balance:"
            Label10.Text = "Balance Amount:"
            Label11.Text = "Balance Remain:"
            BtnRefresh.Text = "Find"
            Label15.Text = "Rate Prev"
            Label21.Text = "Signature1"
            Label3.Text = "Signature2"
            Label4.Text = "Signature3"
            Label13.Text = "Signature4"
            Label20.Text = "Location"
            CheckBox3.Text = "After Balance Sheet"
            RGL.Text = "Genaeral"
            RDtail.Text = "Acc Group"
            RGroup.Text = "Account Only"
            CheckBox2.Text = "Movement"
        End If
    End Sub
    Private Sub selectMMM()
        DMonth.Items.Clear()
        If FmMain.MnLaoLang.Checked = True Then
            DMonth.Items.Add("ມັງກອນ")
            DMonth.Items.Add("ກຸມພາ")
            DMonth.Items.Add("ມີນາ")
            DMonth.Items.Add("ເມສາ")
            DMonth.Items.Add("ພຶດສະພາ")
            DMonth.Items.Add("ມິຖຸນາ")
            DMonth.Items.Add("ກໍລະກົດ")
            DMonth.Items.Add("ສິງຫາ")
            DMonth.Items.Add("ກັນຍາ")
            DMonth.Items.Add("ຕຸລາ")
            DMonth.Items.Add("ພະຈິກ")
            DMonth.Items.Add("ທັນວາ")
        Else
            DMonth.Items.Add("January")
            DMonth.Items.Add("February")
            DMonth.Items.Add("March")
            DMonth.Items.Add("April")
            DMonth.Items.Add("May")
            DMonth.Items.Add("June")
            DMonth.Items.Add("July")
            DMonth.Items.Add("August")
            DMonth.Items.Add("September")
            DMonth.Items.Add("October")
            DMonth.Items.Add("November")
            DMonth.Items.Add("December")
        End If
    End Sub
    Private Sub selectLoad()
        Ct.Enabled = False
        yyt.Enabled = False
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
        ElseIf RT.Checked = True Then
            Ct.Enabled = True
            yyt.Enabled = True
            Call LoadMt()
        ElseIf RY.Checked = True Then
            yy.Enabled = True
            LoadYear()
        End If
    End Sub
    Private Sub LoadMt()
        If Ct.SelectedIndex = 0 Then
            MdStartDate = Format(CDate("1/1/" & Year(yyt.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/6/" & Year(yyt.Value)), "dd-MM-yyyy")
        Else
            MdStartDate = Format(CDate("1/7/" & Year(yyt.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/12/" & Year(yyt.Value)), "dd-MM-yyyy")
        End If

        Lb.Text = Ct.Text & " " & yyt.Text
        'L5.Text = MdStartDate & " => " & MdToDate

        L5.Text = MdStartDate & " => " & MdToDate


        'L5.Text = MdStartDate & " => " & MdToDate
    End Sub
    Private Sub Myy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Myy.ValueChanged
        LoadMonth()
    End Sub

    Private Sub RM_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        selectLoad()
    End Sub

    Private Sub Pyy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Pyy.ValueChanged
        LoadPeriod()
    End Sub

    Private Sub Dt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Dt.ValueChanged
        LoadDay()
    End Sub

    Private Sub RP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        selectLoad()

    End Sub

 
    Private Sub RY_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        selectLoad()
    End Sub

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Close()
        'If DMonth.Text <> DMonth.Items Then

        'End If
    End Sub

   
    Private Sub LoadReport()
        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As    RptSjUd ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
        LngId = "7002" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDP ,"
        LngId = "7005" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_No ,"
        LngId = "7038" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Descrip ,"
        LngId = "7006" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Ac_Code ,"
        LngId = "7008" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Dr ,"
        LngId = "7009" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Cr ,"
        LngId = "7010" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Amt_LAK ,"
        LngId = "7011" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Sign ,"
        LngId = "7012" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_SignNme ,"
        LngId = "7026" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InDate ,"
        LngId = "7028" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_AmtCurr ,"
        LngId = "7031" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_User ,"
        LngId = "7042" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_OpenAmt ,"
        LngId = "7039" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_TotalAmt ,"
        LngId = "7040" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Balance ,"
        LngId = "7043" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RemAmt ,"
        LngId = "7048" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"

        LngId = "7096" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"

        If CMB_Curr.Text = "EQVL" Then
            If CheckBox4.Checked = True Then
                LngId = "7121" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
            Else
                LngId = "7106" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
            End If
        ElseIf CMB_Curr.Text = "LAK" Then
            LngId = "7097" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        ElseIf CMB_Curr.Text = "USD" Then
            LngId = "7108" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        End If



        LngId = "7098" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As B_Dr ,"
        LngId = "7099" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As B_Cr ,"
        LngId = "7100" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As T_Dr ,"
        LngId = "7101" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As T_Cr ,"
        LngId = "7102" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As R_Dr ,"
        LngId = "7103" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As R_Cr ,"

        LngId = "7104" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Ac_ID ,"
        LngId = "7105" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Detail ,"

        If BalanceType.SelectedIndex = 1 Then
            LngId = "7058" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
        Else
            If RM.Checked = True Or RD.Checked = True Then
                LngId = "7045" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            ElseIf RP.Checked = True Then
                LngId = "7046" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            ElseIf RT.Checked = True Then
                LngId = "7080" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
                LngId = "7081" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
                LngId = "7045" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            ElseIf RY.Checked = True Then
                LngId = "7047" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"

            End If
        End If
        Call LoadLoGO()
        If MuLng = "L" Then
            CNN.Execute(" update Ap_balance_6_col set H1=Acc_Code.H1, H1_Nm=Acc_Code.H1_Nm from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.ac_code=Acc_Code.ac_code ")
        Else
            CNN.Execute(" update Ap_balance_6_col set H1=Acc_Code.H1, H1_Nm=Acc_Code.Name_E from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.ac_code=Acc_Code.ac_code ")

        End If

        CNN.Execute(" update Ap_balance_6_col set H1=left(H1,7) ")
        If RGroup.Checked = True Then
            CNN.Execute(" update Ap_balance_6_col set H3=left(ac_code,3) ")
            CNN.Execute(" update Ap_balance_6_col set H4=left(ac_code,4) ")
            CNN.Execute(" update Ap_balance_6_col set H5=left(ac_code,5) ")

        End If
        If MuLng = "L" Then
            CNN.Execute(" update Ap_balance_6_col set  H3_Nm=Acc_Code.Name_L from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.H3=Acc_Code.ac_code ")
            CNN.Execute(" update Ap_balance_6_col set   H4_Nm=Acc_Code.Name_L from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.H4=Acc_Code.ac_code ")
            CNN.Execute(" update Ap_balance_6_col set   H5_Nm=Acc_Code.Name_L from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.H5=Acc_Code.ac_code ")

        Else
            CNN.Execute(" update Ap_balance_6_col set  H3_Nm=Acc_Code.Name_E from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.H3=Acc_Code.ac_code ")
            CNN.Execute(" update Ap_balance_6_col set   H4_Nm=Acc_Code.Name_E from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.H4=Acc_Code.ac_code ")
            CNN.Execute(" update Ap_balance_6_col set   H5_Nm=Acc_Code.Name_E from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.H5=Acc_Code.ac_code ")

        End If



        If CMB_Curr.SelectedIndex = 1 Then
            'CNN.Execute(" update Ap_balance_6_col set H1=left(H1,3) ")
            CNN.Execute(" update Ap_balance_6_col set ac_code='00-'+ ac_code,H1='00-'+ H1 ")

            CNN.Execute(" update Ap_balance_6_col set H3='00-'+left(H3,3) ")
            CNN.Execute(" update Ap_balance_6_col set H4='00-'+left(H4,4) ")
            CNN.Execute(" update Ap_balance_6_col set H5='00-'+left(H5,5) ")

        ElseIf CMB_Curr.SelectedIndex = 2 Then
            CNN.Execute(" update Ap_balance_6_col set ac_code='01-'+ ac_code,H1='01-'+ H1 ")
            CNN.Execute(" update Ap_balance_6_col set H3='01-'+left(H3,3) ")
            CNN.Execute(" update Ap_balance_6_col set H4='01-'+left(H4,4) ")
            CNN.Execute(" update Ap_balance_6_col set H5='01-'+left(H5,5) ")
        End If
        CNN.Execute(" delete Ap_balance_6_col where AC_name is null ")
        SLF = "SELECT  " & mformat & "  as mformat  ,  " & MuLngRpt & "  *   FROM Ap_balance_6_Col Order by ac_code asc "
        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            .Open(SLF, CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim Rpt As New Object
        If RGL.Checked = True Then
            Rpt = New CryTrialBalanceReport_6_Incom_NEW
        ElseIf RDtail.Checked = True Then
            Rpt = New CryTrialBalanceReport_6_Incom_NEW_Acc
        Else
            'MsgBox("NO") : Exit Sub
            Rpt = New CryTrialBalanceReport_6_Incom_NEW_Mae_3_5
        End If
        'Dim Rpt As New CryTrialBalanceReport_6_Incom_NEW

        If MdShowLOGO = 1 Then
            Rpt.Subreports(0).SetDataSource(RsLOGO)
        End If
        Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S1"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS1.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S2"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS2.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S3"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS3.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S4"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS4.Text

        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("pp"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtPP.Text

        Rpt.SetDataSource(Rs)
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.DisplayGroupTree = False
        FrmPreview.MdiParent = FmMain
        FrmPreview.WindowState = FormWindowState.Maximized
        FrmPreview.Show()
        FrmPreview.Focus()
        'FrmPreview.Focus()
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
        LoadSqlData("   select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook & "  group BY ac_code ", RSC12)
        With RSC12
            Do Until .EOF = True
                VCode1 = CStr(Trim(.Fields("ac_Code").Value))
                CNN.Execute("INSERT INTO Ap_balance_6_col( ac_code   , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  )  select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr  , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As Status from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook & "  group BY ac_code ")
                '" select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr  , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As Status from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook & "  group BY ac_code ")
                '"Values('" & CStr(Trim(.Fields("ac_Code").Value)) & "', " & _
                '" " & CDbl(0) & ", " & CDbl(0) & ", " & CDbl(Trim(.Fields("amt_dr").Value)) & ", " & CDbl(Trim(.Fields("amt_cr").Value)) & ",0 )")
                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub LoadOpen_Jn2()
        Dim S As Date = MdStartDate
        S = DateAdd("d", CDbl(-1), MdStartDate)
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
       " select ac_code , 0 As open_amt_dr , 0 open_amt_cr ,sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 Status from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook & "   group BY ac_code")

    End Sub

    Private Sub LoadOpen_Jn3()
        Dim RSC3 As New ADODB.Recordset
        LoadSqlData("select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr from Open_jn WHERE    date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook & "   group BY ac_code", RSC3)
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
            Else
                LoadOpen_Jn5()
            End If
        End With
    End Sub

    Private Sub LoadOpen_Jn5()
        Dim RSC5 As New ADODB.Recordset
        With RSC
            LoadSqlData("select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr from Open_jn  WHERE    ac_code='" & VCode3 & "' " & MULook & " group BY ac_code", RSC5)
            If RSC5.RecordCount > 0 Then
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
             "Values('" & CStr(Trim(RSC5.Fields("ac_Code").Value)) & "',  " & _
             " " & CDbl(RSC5.Fields("amt_dr").Value) & ", " & CDbl(RSC5.Fields("amt_cr").Value) & ", " & CDbl(0) & ", " & CDbl(0) & ",0 )")
            Else
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
                    CNN.Execute("Update Ap_balance_6 set rem_dr='" & CDbl(op_dr + amt_dr) - CDbl(op_cr + amt_cr) & "' , rem_cr='" & CDbl(0) & "' where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                End If
                If CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) >= 0 Then
                    CNN.Execute("Update Ap_balance_6 set rem_dr='" & CDbl(0) & "' , rem_cr='" & CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) & "' where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                End If
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
                If CDbl(CDbl((.Fields("open_amt_dr").Value)) + CDbl((.Fields("amt_dr").Value))) >= CDbl(CDbl((.Fields("open_amt_cr").Value)) + CDbl((.Fields("amt_cr").Value))) Then
                    CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr=" & CDbl(CDbl(CDbl((.Fields("open_amt_dr").Value)) + CDbl((.Fields("amt_dr").Value))) - CDbl(CDbl((.Fields("open_amt_cr").Value)) + CDbl((.Fields("amt_cr").Value)))) & " where Ac_Code = '" & (.Fields("Ac_Code").Value) & "'")
                Else
                    CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr=" & CDbl(CDbl(CDbl((.Fields("open_amt_cr").Value)) + CDbl((.Fields("amt_cr").Value))) - CDbl(CDbl((.Fields("open_amt_dr").Value)) + CDbl((.Fields("amt_dr").Value)))) & " where Ac_Code = '" & (.Fields("Ac_Code").Value) & "'")
                End If
                .MoveNext()
            Loop
        End With
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


    Private Sub BtnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRefresh.Click
        If BalanceType.SelectedIndex = 0 Then
            Ac_Code = ""
        ElseIf BalanceType.SelectedIndex = 1 Then
            Ac_Code = "And (Left(Ac_Code,1) = '4' or Left(Ac_Code,1) = '5' or Left(Ac_Code,4) = '00.4' or Left(Ac_Code,4) = '00.5') "
            'Ac_Code = "And (Left(Ac_Code,4) = '00.4' or Left(Ac_Code,4) = '00.5') "
            ChangInCom = BalanceType.SelectedIndex

            'MsgBox("dd")

        End If
        If CheckBox1.Checked = True Then
            ChangInCom = 1
        Else
            ChangInCom = 0
        End If


        New_Code = "3901000.00.0000"
        Code_Dr = "4"
        Code_Cr = "5"

        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        'If CMB_Curr.SelectedIndex = 0 Then
        'Call ChangBalance()
        'Else
        Call BLNEW()
        'End If


        If CheckBnk.Checked = True Then
            CNN.Execute("delete Ap_balance_6_col from Ap_balance_6_col , Ap_Rpt_BLS_Item where Ap_balance_6_col.ac_code=Ap_Rpt_BLS_Item.Ac_Code")
            CNN.Execute("delete Ap_balance_6_col from Ap_balance_6_col , Ap_Rpt_Income_Item where Ap_balance_6_col.ac_code=Ap_Rpt_Income_Item.Ac_Code")
        End If
        If CheckBox2.Checked = True Then
            CNN.Execute("delete Ap_balance_6_col where Amt_Dr=0 And Amt_Cr=0 ")
        End If
        If TextBox1.Text <> "" Then
            CNN.Execute("delete Ap_balance_6_col where Left(Ac_Code, " & Len(TextBox1.Text) & ") <> '" & TextBox1.Text & "' ")
        End If
        If MuLng = "L" Then
            CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        Else
            CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_E from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        End If
        '        CNN.Execute(" insert into Ap_balance_6_col(ac_code, ac_name, open_amt_dr, open_amt_cr,  amt_dr, amt_cr,   Rem_dr, Rem_cr,Status) " & _
        '" Select '2382120.00.0000',N'ຄູ່ມູນຄ່າຖານະ ແລກປ່ຽນເງິນຕາ ຕ່າງປະເທດ',sum(open_amt_dr),sum(open_amt_cr),sum(amt_dr),sum(amt_cr),sum(Rem_dr),sum(Rem_cr),1   " & _
        '" from Ap_balance_6_col where left(ac_code,7)='2382120' group by  left(ac_code,7) ")
        '        CNN.Execute("delete  Ap_balance_6_col  where left(ac_code,7)='2382120' and Status is null ")

        Call LoadListFG()
        Call SumData()
    End Sub
    Private Sub BLNEW()
        CNN.Execute("   update gen_jn set Rate_USD=rate where curr='USD' and Rate_USD is null ")
        CNN.Execute("  update gen_jn set Rate_USD=0 where   Rate_USD is null ")
        CNN.Execute("  update gen_jn set amt_USD_Dr=amount_dr  where curr='USD' and (amt_USD_Dr=0 or amt_USD_Dr is null) ")
        CNN.Execute("  update gen_jn set amt_USD_cr= amount_Cr   where curr='USD'  and (amt_USD_cr=0 or amt_USD_cr is null) ")
        CNN.Execute("  update gen_jn set amt_USD_Dr= amt_dr/Rate_USD  where curr='LAK' and Rate_USD<>0")
        CNN.Execute("  update gen_jn set amt_USD_cr= amt_cr/Rate_USD    where curr='LAK'  and Rate_USD<>0")
        '==============OPEN=====
        CNN.Execute("   update Open_jn set Rate_USD=rate where curr='USD' and Rate_USD is null ")
        CNN.Execute("  update Open_jn set Rate_USD=0 where   Rate_USD is null ")
        CNN.Execute("  update Open_jn set amt_USD_Dr=amount_dr  where curr='USD' and (amt_USD_Dr=0 or amt_USD_Dr is null) ")
        CNN.Execute("  update Open_jn set amt_USD_cr= amount_Cr   where curr='USD'  and (amt_USD_cr=0 or amt_USD_cr is null) ")

        CNN.Execute("  update Open_jn set amt_USD_Dr= amt_dr/Rate_USD  where curr='LAK' and Rate_USD<>0")
        CNN.Execute("  update Open_jn set amt_USD_cr= amt_cr/Rate_USD    where curr='LAK'  and Rate_USD<>0")
        '==============Adjust=====
        CNN.Execute("   update AP_ACC_adjust_Item set Rate_USD=rate where curr='USD' and Rate_USD is null ")
        CNN.Execute("  update AP_ACC_adjust_Item set Rate_USD=0 where   Rate_USD is null ")
        CNN.Execute("  update AP_ACC_adjust_Item set amt_USD_Dr=amount_dr  where curr='USD' and (amt_USD_Dr=0 or amt_USD_Dr is null) ")
        CNN.Execute("  update AP_ACC_adjust_Item set amt_USD_cr= amount_Cr   where curr='USD'  and (amt_USD_cr=0 or amt_USD_cr is null) ")

        CNN.Execute("  update AP_ACC_adjust_Item set amt_USD_Dr= amt_dr/Rate_USD  where curr='LAK' and Rate_USD<>0")
        CNN.Execute("  update AP_ACC_adjust_Item set amt_USD_cr= amt_cr/Rate_USD    where curr='LAK'  and Rate_USD<>0")

        '=================NEWWWWW==============
        CNN.Execute("DELETE  Ap_balance_6_col ")
        CNN.Execute("DELETE FROM Ap_balance_6 ")
        Dim B_Curr As String = ""
        If CMB_Curr.SelectedIndex = 0 Then
            B_Curr = ""
        Else
            B_Curr = " AND  Curr=N'" & CMB_Curr.Text & "' "
        End If


        'CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        '" select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")
        CNN.Execute("UPDATE gen_jn set amt_USD_Dr=0 where amt_USD_Dr is null ")
        CNN.Execute("UPDATE gen_jn set amt_USD_Cr=0 where amt_USD_Cr is null ")
        CNN.Execute("UPDATE Open_jn set amt_USD_Dr=0 where amt_USD_Dr is null ")
        CNN.Execute("UPDATE Open_jn set amt_USD_Cr=0 where amt_USD_Cr is null ")
        CNN.Execute("UPDATE AP_ACC_adjust_Item set amt_USD_Dr=0 where amt_USD_Dr is null ")
        CNN.Execute("UPDATE AP_ACC_adjust_Item set amt_USD_Cr=0 where amt_USD_Cr is null ")

        If CMB_Curr.SelectedIndex = 0 Then
            If CheckBox4.Checked = True Then
                '=======LAK===
                Dim GGG As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
               " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amt_USD_Dr)as amt_dr , sum(amt_USD_Cr)as amt_cr  from gen_jn  WHERE 1=1 " & B_Curr & " and gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code "
                CNN.Execute(GGG)

                Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)

                'CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '" select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
                '=======LAK===
                Dim PPP As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code , sum(amt_USD_Dr)as amt_dr , sum(amt_USD_Cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE 1=1 " & B_Curr & " and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code"
                CNN.Execute(PPP)

                '=======LAK===
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code  , sum(amt_USD_Dr) as amt_dr , sum(amt_USD_Cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE  1=1 " & B_Curr & " and   date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

                If CheckBox3.Checked = True Then
                    Dim aa As String
                    aa = "  insert into  Ap_balance_6 (ac_code,amt_dr,amt_cr)   SELECT  AP_ACC_adjust_Item.ac_code,SUM (AP_ACC_adjust_Item.amt_USD_Dr),SUM (AP_ACC_adjust_Item.amt_USD_Cr)  from  AP_ACC_adjust_Item " & _
                            "   where  year(AP_ACC_adjust_Item.date_work) = '" & Format(CDate(MdStartDate), "yyyy") & "' " & MULook2 & "  " & _
                            " group by  ac_code,com_id   "
                    CNN.Execute(aa)

                End If
            Else
                '=======LAK===
                Dim GGG As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
               " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(Amount_Dr)as amt_dr , sum(Amount_cr)as amt_cr  from gen_jn  WHERE 1=1  and Curr=N'LAK'   and gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code "
                CNN.Execute(GGG)

                '            Dim GGu As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '" select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE 1=1   and gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code "
                '            CNN.Execute(GGu)

                Dim USD As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
         " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amount_Dr)* " & CDbl(txtRate.Text) & "  as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & " as amt_cr  from gen_jn  WHERE 1=1 and Curr=N'USD'  and gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code "
                CNN.Execute(USD)

                Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)


                'CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '" select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
                '=======LAK===
                Dim PPP As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                        " select ac_code , sum(Amount_Dr)as amt_dr , sum(Amount_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE 1=1  and Curr=N'LAK'   and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code"
                CNN.Execute(PPP)
                Dim PPPUSD As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code , sum(amount_Dr)* " & CDbl(txtRate.Text) & " as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & " as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE 1=1  and Curr=N'USD'  and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code"
                CNN.Execute(PPPUSD)

                '        '=======LAK===
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code  , sum(Amount_Dr) as amt_dr , sum(Amount_cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1   and Curr=N'LAK'  and date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

                'CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '" select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE  1=1   and Curr=N'LAK'  and   date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
           " select ac_code  , sum(amount_Dr)* " & CDbl(txtRate.Text) & "  as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & "  as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE  1=1  and Curr=N'USD'  and   date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

                If CheckBox3.Checked = True Then
                    Dim aa As String
                    aa = "  insert into  Ap_balance_6 (ac_code,amt_dr,amt_cr)   SELECT  AP_ACC_adjust_Item.ac_code,SUM (AP_ACC_adjust_Item.amt_dr),SUM (AP_ACC_adjust_Item.amt_cr)  from  AP_ACC_adjust_Item " & _
                            "   where  year(AP_ACC_adjust_Item.date_work) = '" & Format(CDate(MdStartDate), "yyyy") & "' " & MULook2 & "  " & _
                            " group by  ac_code,com_id   "
                    CNN.Execute(aa)
 

                End If
            End If


        Else
            '=======Curr===
            '     If CheckBox4.Checked = True Then
            '         If CMB_Curr.Text = "LAK" Then
            '             Dim GGG As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            '" select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amt_USD_Dr)as amt_dr , sum(amt_USD_Cr)as amt_cr  from gen_jn  WHERE 1=1 " & B_Curr & " and gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code "
            '             CNN.Execute(GGG)
            '             Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
            '             Dim PPP As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            '     " select ac_code , sum(amt_USD_Dr)as amt_dr , sum(amt_USD_Cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE 1=1 " & B_Curr & " and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code"
            '             CNN.Execute(PPP)
            '             CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            '             " select ac_code  , sum(amt_USD_Dr) as amt_dr , sum(amt_USD_Cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE  1=1 " & B_Curr & " and   date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
            '             If CheckBox3.Checked = True Then
            '                 Dim aa As String
            '                 aa = "  insert into  Ap_balance_6 (ac_code,amt_dr,amt_cr)   SELECT  AP_ACC_adjust_Item.ac_code,SUM (AP_ACC_adjust_Item.amt_USD_Dr),SUM (AP_ACC_adjust_Item.amt_USD_Cr)  from  AP_ACC_adjust_Item " & _
            '                         "   where  year(AP_ACC_adjust_Item.date_work) = '" & Format(CDate(MdStartDate), "yyyy") & "' " & MULook2 & "  " & _
            '                         " group by  ac_code,com_id   "
            '                 CNN.Execute(aa)
            '             End If
            '         ElseIf CMB_Curr.Text = "USD" Then
            '             Dim GGG As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            '" select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amt_Dr)as amt_dr , sum(amt_Cr)as amt_cr  from gen_jn  WHERE 1=1 " & B_Curr & " and gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code "
            '             CNN.Execute(GGG)
            '             Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
            '             Dim PPP As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            '     " select ac_code , sum(amt_Dr)as amt_dr , sum(amt_Cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE 1=1 " & B_Curr & " and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code"
            '             CNN.Execute(PPP)
            '             CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            '             " select ac_code  , sum(amt_Dr) as amt_dr , sum(amt_Cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE  1=1 " & B_Curr & " and   date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
            '             If CheckBox3.Checked = True Then
            '                 Dim aa As String
            '                 aa = "  insert into  Ap_balance_6 (ac_code,amt_dr,amt_cr)   SELECT  AP_ACC_adjust_Item.ac_code,SUM (AP_ACC_adjust_Item.amt_Dr),SUM (AP_ACC_adjust_Item.amt_Cr)  from  AP_ACC_adjust_Item " & _
            '                         "   where  year(AP_ACC_adjust_Item.date_work) = '" & Format(CDate(MdStartDate), "yyyy") & "' " & MULook2 & "  " & _
            '                         " group by  ac_code,com_id   "
            '                 CNN.Execute(aa)
            '             End If
            '         End If

            '     Else

            Dim GGG As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
" select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(Amount_Dr)as amt_dr , sum(Amount_cr)as amt_cr  from gen_jn  WHERE 1=1 " & B_Curr & " and gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code "
            CNN.Execute(GGG)

            Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
            Dim PPP As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
    " select ac_code , sum(Amount_Dr)as amt_dr , sum(Amount_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE 1=1 " & B_Curr & " and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code"
            CNN.Execute(PPP)
            CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            " select ac_code  , sum(Amount_Dr) as amt_dr , sum(Amount_cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE  1=1 " & B_Curr & " and   date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
            If CheckBox3.Checked = True Then
                Dim aa As String
                aa = "  insert into  Ap_balance_6 (ac_code,open_amt_dr,open_amt_cr,amt_dr,amt_cr)   SELECT  AP_ACC_adjust_Item.ac_code,0,0,SUM (AP_ACC_adjust_Item.Amount_Dr),SUM (AP_ACC_adjust_Item.Amount_cr)  from  AP_ACC_adjust_Item " & _
                        "   where  year(AP_ACC_adjust_Item.date_work) = '" & Format(CDate(MdStartDate), "yyyy") & "' " & MULook2 & "  " & _
                        " group by  ac_code,com_id   "
                CNN.Execute(aa)
            End If
        End If
        'End If

        CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr)as open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")
        If CMB_Curr.SelectedIndex = 0 Then
            '        CNN.Execute(" insert into Ap_balance_6_col(ac_code, ac_name, open_amt_dr, open_amt_cr,  amt_dr, amt_cr,   Rem_dr, Rem_cr,Status) " & _
            '" Select '2382120.00.0000',N'ຄູ່ມູນຄ່າຖານະ ແລກປ່ຽນເງິນຕາ ຕ່າງປະເທດ',sum(open_amt_dr),sum(open_amt_cr),sum(amt_dr),sum(amt_cr),sum(Rem_dr),sum(Rem_cr),1   " & _
            '" from Ap_balance_6_col where left(ac_code,7)='2382120' group by  left(ac_code,7) ")
            '        CNN.Execute("delete  Ap_balance_6_col  where left(ac_code,7)='2382120' and Status is null ")
        End If


        '=================NEWWWWW==============
        'CNN.Execute("DELETE  Ap_balance_6_col ")
        'CNN.Execute("DELETE FROM Ap_balance_6 ")
        'Dim KKKa As String = " insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  from Ap_balance_TB where  date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  " & Ac_Code & "  order by Ac_Code asc "
        'CNN.Execute(KKKa)

        Call Left_AcCode()
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        Call Chang_Incom()
        CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + amt_dr) - (open_amt_cr + amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr + amt_cr) ")
        CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr + amt_cr) - (open_amt_dr + amt_dr) where (open_amt_cr + amt_cr) >= (open_amt_dr + amt_dr) ")
        If MuLng = "L" Then
            CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        Else
            CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_E from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        End If
    End Sub

    Private Sub Left_AcCode()
        Dim L As String
        If MuLeftAcCode > 0 Then
            L = CDbl(MuLeftAcCode) + 2
            Insr = "delete Ap_balance_6 " & _
           "Update Ap_balance_6_col set Acc_Parent = left(Ac_Code," & L & ") " & _
            "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  )  " & _
            " select Acc_Parent ,sum(open_amt_dr) As open_amt_dr ,sum(open_amt_cr) As open_amt_cr  ,sum(amt_dr) As amt_dr ,sum(amt_cr) As amt_cr  from Ap_balance_6_col group by  Acc_Parent " & _
           "  delete Ap_balance_6_col " & _
           "insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  )  " & _
         "select Ac_Code ,sum(open_amt_dr) As open_amt_dr ,sum(open_amt_cr) As open_amt_cr  ,sum(amt_dr) As amt_dr ,sum(amt_cr) As amt_cr  from Ap_balance_6  group by  Ac_Code"
            CNN.Execute(Insr)
        End If
    End Sub
    Private Sub Chang_Incom()
        'or  Ac_Code =  '" & New_Code & "' 
        If ChangInCom = 1 Then
            Insr = "delete  Ap_balance_6  " & _
               "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr)   select '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr) from Ap_balance_6_col where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "' " & _
      "update  Ap_balance_6 set open_amt_dr = 0 where open_amt_dr  is null  " & _
      "update  Ap_balance_6 set open_amt_cr = 0 where open_amt_cr  is null  " & _
      "update  Ap_balance_6 set amt_dr = 0 where amt_dr  is null  " & _
      "update  Ap_balance_6 set amt_cr = 0 where amt_cr  is null   " & _
       "Update  Ap_balance_6 set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr   " & _
      "Update  Ap_balance_6 set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr " & _
      "Update  Ap_balance_6 set   amt_dr = amt_dr  - amt_cr , amt_cr=0  where amt_dr  >= amt_cr   " & _
      "Update  Ap_balance_6 set   amt_cr = amt_cr  - amt_dr , amt_dr=0  where amt_cr  >= amt_dr " & _
         "delete  Ap_balance_6_col  where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "'   " & _
           "  insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr ,status )  " & _
 " select  '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr),1 from Ap_balance_6_col where Ac_Code =  '" & New_Code & "' " & _
    "       delete  Ap_balance_6_col where Ac_Code =  '" & New_Code & "' " & _
  "  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , sum(open_amt_dr) , sum(open_amt_cr) , sum(amt_dr) , sum(amt_cr)  from Ap_balance_6 group by Ac_Code "

            CNN.Execute(Insr)

            '            Insr = "delete  Ap_balance_6  " & _
            '   "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  )   select '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr) from Ap_balance_6_col where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "' or  Ac_Code =  '" & New_Code & "'  " & _
            '"update  Ap_balance_6 set open_amt_dr = 0 where open_amt_dr  is null  " & _
            '"update  Ap_balance_6 set open_amt_cr = 0 where open_amt_cr  is null  " & _
            '"update  Ap_balance_6 set amt_dr = 0 where amt_dr  is null  " & _
            '"update  Ap_balance_6 set amt_cr = 0 where amt_cr  is null   " & _
            '"Update  Ap_balance_6 set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr   " & _
            '"Update  Ap_balance_6 set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr " & _
            '"Update  Ap_balance_6 set   amt_dr = amt_dr  - amt_cr , amt_cr=0  where amt_dr  >= amt_cr   " & _
            '"Update  Ap_balance_6 set   amt_cr = amt_cr  - amt_dr , amt_dr=0  where amt_cr  >= amt_dr " & _
            '"delete  Ap_balance_6_col  where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "'  or  Ac_Code =  '" & New_Code & "'  " & _
            '"  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , sum(open_amt_dr) , sum(open_amt_cr) , sum(amt_dr) , sum(amt_cr)  from Ap_balance_6 group by Ac_Code "
            '            CNN.Execute(Insr)
            '"  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  from Ap_balance_6"
        End If
    End Sub
    Private Sub LoadListFG()
        Dim O_dr, O_cr, Amt_dr, Amt_cr, R_dr, R_Cr As Double
        FG.Rows = 1
        If CMB_Curr.SelectedIndex = 1 Then
            CNN.Execute(" update Ap_balance_6_col set ac_code='00-'+ ac_code ")
        ElseIf CMB_Curr.SelectedIndex = 2 Then
            CNN.Execute(" update Ap_balance_6_col set ac_code='01-'+ ac_code ")
        End If
        CNN.Execute("UPDATE Ap_balance_6_col set open_amt_dr=0 where open_amt_dr is null")
        CNN.Execute("UPDATE Ap_balance_6_col set open_amt_cr=0 where open_amt_cr is null")
        CNN.Execute("UPDATE Ap_balance_6_col set Rem_dr=0 where Rem_dr is null")
        CNN.Execute("UPDATE Ap_balance_6_col set Rem_cr=0 where Rem_cr is null")

        With RSC

            Call LoadSqlData("SELECT * FROM  Ap_balance_6_col Order by ac_code", RSC)

            If .RecordCount > 0 Then
                While Not .EOF

                    O_dr = Trim(CDbl(.Fields("open_amt_dr").Value))
                    O_cr = Trim(CDbl(.Fields("open_amt_cr").Value))
                    Amt_dr = Trim(CDbl(.Fields("amt_dr").Value))
                    Amt_cr = Trim(CDbl(.Fields("amt_cr").Value))
                    R_dr = Trim(CDbl(.Fields("Rem_dr").Value))
                    R_Cr = Trim(CDbl(.Fields("Rem_cr").Value))

                    FG.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("ac_code").Value)) & _
                                "" & vbTab & Trim(CStr(.Fields("ac_name").Value.ToString)) & _
                                        "" & vbTab & Format(O_dr, "##,##0.00") & _
                                         "" & vbTab & Format(O_cr, "##,##0.00") & _
                                          "" & vbTab & Format(Amt_dr, "##,##0.00") & _
                                             "" & vbTab & Format(Amt_cr, "##,##0.00") & _
                                                 "" & vbTab & Format(R_dr, "##,##0.00") & _
                                             "" & vbTab & Format(R_Cr, "##,##0.00") & _
                                            "" & vbTab & ((.Fields("cnt").Value)))
                    .MoveNext()
                End While
            Else
                FG.Rows = 1
            End If
        End With

    End Sub
    Private Sub SumData()

        OpDr.Text = 0
        OpCr.Text = 0
        AmtDr.Text = 0
        AmtCr.Text = 0
        ReDr.Text = 0
        ReCr.Text = 0
        BOpDr.Text = 0
        BAmtDr.Text = 0
        BReDr.Text = 0
        For i = 1 To FG.Rows - 1
            OpDr.Text = CDbl(OpDr.Text) + CDbl(FG.get_TextMatrix(i, 3))
            OpCr.Text = CDbl(OpCr.Text) + CDbl(FG.get_TextMatrix(i, 4))
            AmtDr.Text = CDbl(AmtDr.Text) + CDbl(FG.get_TextMatrix(i, 5))
            AmtCr.Text = CDbl(AmtCr.Text) + CDbl(FG.get_TextMatrix(i, 6))
            ReDr.Text = CDbl(ReDr.Text) + CDbl(FG.get_TextMatrix(i, 7))
            ReCr.Text = CDbl(ReCr.Text) + CDbl(FG.get_TextMatrix(i, 8))
        Next i
        BOpDr.Text = CDbl(OpCr.Text) - CDbl(OpDr.Text)
        BAmtDr.Text = CDbl(AmtCr.Text) - CDbl(AmtDr.Text)
        BReDr.Text = CDbl(ReCr.Text) - CDbl(ReDr.Text)
        OpDr.Text = Format(CDbl(OpDr.Text), "##,##0.00")
        OpCr.Text = Format(CDbl(OpCr.Text), "##,##0.00")
        AmtDr.Text = Format(CDbl(AmtDr.Text), "##,##0.00")
        AmtCr.Text = Format(CDbl(AmtCr.Text), "##,##0.00")
        ReDr.Text = Format(CDbl(ReDr.Text), "##,##0.00")
        ReCr.Text = Format(CDbl(ReCr.Text), "##,##0.00")
        BOpDr.Text = Format(CDbl(BOpDr.Text), "##,##0.00")
        BAmtDr.Text = Format(CDbl(BAmtDr.Text), "##,##0.00")
        BReDr.Text = Format(CDbl(BReDr.Text), "##,##0.00")


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

    Private Sub RT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RT.CheckedChanged
        selectLoad()
    End Sub

    Private Sub Ct_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ct.SelectedIndexChanged
        selectLoad()
    End Sub

    Private Sub yyt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yyt.ValueChanged
        selectLoad()
    End Sub

    Private Sub CheckBnk_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBnk.CheckedChanged

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        AddHeader()
        If RD.Checked = True Then
            Dim s1, s2 As String
            LngId = 7027 : CallLngStr() : s2 = LngStr
            LngId = 7072 : CallLngStr() : s1 = LngStr
            Lb.Text = s1 & " " & Ds.Text & " " & s2 & " " & Dt.Text
        ElseIf RM.Checked = True Then
            Dim s1, s2, s3 As String
            LngId = 7013 + DMonth.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7070 : CallLngStr() : s1 = LngStr
            LngId = 7064 : CallLngStr() : s3 = LngStr
            'Lb.Text = s1 & " (" & s2 & ") " & s3 & " " & Year(MdToDate)
        ElseIf RP.Checked = True Then
            Dim s1, s2, s3 As String
            LngId = 7074 + Period.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7062 : CallLngStr() : s1 = LngStr
            LngId = 7064 : CallLngStr() : s3 = LngStr
            Lb.Text = s1 & "  " & s2 & " " & " " & Year(MdToDate)
        ElseIf RY.Checked = True Then
            Dim s1, s2, s3 As String
            LngId = 7074 + Period.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7062 : CallLngStr() : s1 = LngStr
            LngId = 7064 : CallLngStr() : s3 = LngStr
            Lb.Text = s3 & " " & yy.Text
        End If
        If BalanceType.SelectedIndex = 0 Then
            Ac_Code = ""
        ElseIf BalanceType.SelectedIndex = 1 Then
            'Ac_Code = "And (Left(Ac_Code,1) = '4' or Left(Ac_Code,1) = '5') "
            'Ac_Code = "And (Left(Ac_Code,4) = '00.4' or Left(Ac_Code,4) = '00.5') "
            Ac_Code = "And (Left(Ac_Code,1) = '4' or Left(Ac_Code,1) = '5' or Left(Ac_Code,4) = '00.4' or Left(Ac_Code,4) = '00.5') "
        End If
        New_Code = "3901000"
        Code_Dr = "4"
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        'Call ChangBalance()

        BLNEW()
        If CheckBnk.Checked = True Then
            CNN.Execute("delete Ap_balance_6_col from Ap_balance_6_col , Ap_Rpt_BLS_Item where Ap_balance_6_col.ac_code=Ap_Rpt_BLS_Item.Ac_Code")
            CNN.Execute("delete Ap_balance_6_col from Ap_balance_6_col , Ap_Rpt_Income_Item where Ap_balance_6_col.ac_code=Ap_Rpt_Income_Item.Ac_Code")
        End If
        If CheckBox2.Checked = True Then
            CNN.Execute("delete Ap_balance_6_col where Amt_Dr=0 And Amt_Cr=0 ")
        End If
        If MuLng = "L" Then
            CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        Else
            CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_E from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        End If
        Call LoadReport_Export()
    End Sub
    Private Sub LoadReport_Export()
        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As    RptSjUd ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
        LngId = "7002" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDP ,"
        LngId = "7005" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_No ,"
        LngId = "7038" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Descrip ,"
        LngId = "7006" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Ac_Code ,"
        LngId = "7008" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Dr ,"
        LngId = "7009" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Cr ,"
        LngId = "7010" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Amt_LAK ,"
        LngId = "7011" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Sign ,"
        LngId = "7012" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_SignNme ,"
        LngId = "7026" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InDate ,"
        LngId = "7028" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_AmtCurr ,"
        LngId = "7031" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_User ,"
        LngId = "7042" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_OpenAmt ,"
        LngId = "7039" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_TotalAmt ,"
        LngId = "7040" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Balance ,"
        LngId = "7043" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RemAmt ,"
        LngId = "7048" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
        If BalanceType.SelectedIndex = 1 Then
            LngId = "7058" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
        Else
            If RM.Checked = True Or RD.Checked = True Then
                LngId = "7045" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            ElseIf RP.Checked = True Then
                LngId = "7046" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            ElseIf RT.Checked = True Then
                LngId = "7080" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
                LngId = "7081" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
                LngId = "7045" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            ElseIf RY.Checked = True Then
                LngId = "7047" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"

            End If
        End If
        Call LoadLoGO()
        If MuLng = "L" Then
            CNN.Execute(" update Ap_balance_6_col set H1=Acc_Code.H1, H1_Nm=Acc_Code.H1_Nm from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.ac_code=Acc_Code.ac_code ")
        Else
            CNN.Execute(" update Ap_balance_6_col set H1=Acc_Code.H1, H1_Nm=Acc_Code.Name_E from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.ac_code=Acc_Code.ac_code ")

        End If

        'If CMB_Curr.SelectedIndex = 1 Then
        '    CNN.Execute(" update Ap_balance_6_col set ac_code='00-'+ ac_code,H1='00-'+ H1 ")

        'ElseIf CMB_Curr.SelectedIndex = 2 Then
        '    CNN.Execute(" update Ap_balance_6_col set ac_code='01-'+ ac_code,H1='01-'+ H1 ")
        'End If

        If CMB_Curr.SelectedIndex = 1 Then
            'CNN.Execute(" update Ap_balance_6_col set H1=left(H1,3) ")
            CNN.Execute(" update Ap_balance_6_col set ac_code='00-'+ ac_code,H1='00-'+ H1 ")

            CNN.Execute(" update Ap_balance_6_col set H3='00-'+left(H3,3) ")
            CNN.Execute(" update Ap_balance_6_col set H4='00-'+left(H4,4) ")
            CNN.Execute(" update Ap_balance_6_col set H5='00-'+left(H5,5) ")

        ElseIf CMB_Curr.SelectedIndex = 2 Then
            CNN.Execute(" update Ap_balance_6_col set ac_code='01-'+ ac_code,H1='01-'+ H1 ")
            CNN.Execute(" update Ap_balance_6_col set H3='01-'+left(H3,3) ")
            CNN.Execute(" update Ap_balance_6_col set H4='01-'+left(H4,4) ")
            CNN.Execute(" update Ap_balance_6_col set H5='01-'+left(H5,5) ")
        End If



        SLF = "SELECT  " & mformat & "  as mformat  ,  " & MuLngRpt & "  *   FROM Ap_balance_6_Col Order by ac_code asc "


        'If CMB_Curr.SelectedIndex = 1 Then
        '    CNN.Execute(" update Ap_balance_6_col set ac_code='00-'+ ac_code ")
        'ElseIf CMB_Curr.SelectedIndex = 2 Then
        '    CNN.Execute(" update Ap_balance_6_col set ac_code='01-'+ ac_code ")
        'End If

        'SLF = "SELECT  " & MuLngRpt & "  *   FROM Ap_balance_6_Col Order by ac_code asc "
        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            .Open(SLF, CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim Rpt As New Object
        'Dim Rpt As New CryTrialBalanceReport_6_Incom_NEW
        'If RGL.Checked = True Then
        '    Rpt = New CryTrialBalanceReport_6_Incom_NEW
        'ElseIf RDtail.Checked = True Then
        '    Rpt = New CryTrialBalanceReport_6_Incom_NEW_Acc
        'Else
        '    'MsgBox("NO") : Exit Sub
        '    Rpt = New CryTrialBalanceReport_6_Incom_NEW_Mae
        'End If
        If RGL.Checked = True Then
            Rpt = New CryTrialBalanceReport_6_Incom_NEW
        ElseIf RDtail.Checked = True Then
            Rpt = New CryTrialBalanceReport_6_Incom_NEW_Acc
        Else
            'MsgBox("NO") : Exit Sub
            Rpt = New CryTrialBalanceReport_6_Incom_NEW_Mae_3_5
        End If


        Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S1"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS1.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S2"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS2.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S3"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS3.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S4"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS4.Text

        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("pp"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtPP.Text
        If MdShowLOGO = 1 Then
            Rpt.Subreports(0).SetDataSource(RsLOGO)
        End If

        Rpt.SetDataSource(Rs)
        Rpt.Refresh()
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.ExportReport()
        FrmPreview = Nothing

        'Rpt.SetDataSource(Rs)
        'FrmPreview.ReportViewer.ReportSource = Rpt
        'FrmPreview.ReportViewer.DisplayGroupTree = False
        'FrmPreview.MdiParent = FmMain
        'FrmPreview.WindowState = FormWindowState.Maximized
        'FrmPreview.Show()
        'FrmPreview.Focus()
        'FrmPreview.Focus()
    End Sub

    Private Sub CMB_Curr_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMB_Curr.SelectedIndexChanged

        Dim rs As New ADODB.Recordset
        Call LoadSqlData("Select * From Curr_For_Rate Where   Curr =N'" & Trim(CMB_Curr.Text) & "'", rs)
        If rs.RecordCount > 0 Then
            txtcurr_name2.Text = Trim(rs("Curr_name").Value.ToString)
        End If

        MDRate_DT = " and rate_dt<='" & Format(Dt.Value, "yyyy-MM-dd") & "'  "
        MDRate_DT = " and rate_dt<='" & Format(MdToDate, "yyyy-MM-dd") & "'  "

        SS_Curr = " and AP_Rate_history.Curr =N'" & CMB_Curr.Text & "' "
        If CMB_Curr.SelectedIndex = 0 Then
            SS_Curr = " and AP_Rate_history.Curr =N'USD' "
        Else
            SS_Curr = " and AP_Rate_history.Curr =N'" & CMB_Curr.Text & "' "
        End If

        Call RateSetting()
        txtRate.Text = Format(MD_Rate, "#,##0.00")
        If CMB_Curr.Text = "LAK" Then
            CheckBox4.Text = "ທຽບເທົ່າໂດລາ"
        ElseIf CMB_Curr.Text = "USD" Then
            CheckBox4.Text = "ທຽບເທົ່າກີບ"
        Else
            CheckBox4.Text = "ທຽບເທົ່າໂດລາ"
        End If

    End Sub

 
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged

    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtRate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRate.KeyPress
        If e.KeyChar = Chr(13) Then
            txtRate.Text = Format(CDbl(txtRate.Text), "#,##0.00")
        End If

    End Sub

    Private Sub txtRate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRate.LostFocus
        If txtRate.Text = "" Then
            txtRate.Text = 1
            txtRate.Text = Format(CDbl(txtRate.Text), "#,##0.00")
        End If

    End Sub

    Private Sub txtRate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRate.TextChanged

    End Sub

    Private Sub RD_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RD.CheckedChanged
        selectLoad()
    End Sub

    Private Sub RM_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RM.CheckedChanged
        selectLoad()
    End Sub

    Private Sub RY_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RY.CheckedChanged
        selectLoad()
    End Sub

    Private Sub RP_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RP.CheckedChanged
        selectLoad()
    End Sub
   
    Private Sub BtnPreview_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        AddHeader()
        If RD.Checked = True Then
            Dim s1, s2 As String
            LngId = 7027 : CallLngStr() : s2 = LngStr
            LngId = 7072 : CallLngStr() : s1 = LngStr
            Lb.Text = s1 & " " & Ds.Text & " " & s2 & " " & Dt.Text
        ElseIf RM.Checked = True Then
            Dim s1, s2, s3 As String
            LngId = 7013 + DMonth.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7070 : CallLngStr() : s1 = LngStr
            LngId = 7064 : CallLngStr() : s3 = LngStr
            'Lb.Text = s1 & " (" & s2 & ") " & s3 & " " & Year(MdToDate)
        ElseIf RP.Checked = True Then
            Dim s1, s2, s3 As String
            LngId = 7074 + Period.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7062 : CallLngStr() : s1 = LngStr
            LngId = 7064 : CallLngStr() : s3 = LngStr
            Lb.Text = s1 & "  " & s2 & " " & " " & Year(MdToDate)
        ElseIf RY.Checked = True Then
            Dim s1, s2, s3 As String
            LngId = 7074 + Period.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7062 : CallLngStr() : s1 = LngStr
            LngId = 7064 : CallLngStr() : s3 = LngStr
            Lb.Text = s3 & " " & yy.Text
        End If
        If BalanceType.SelectedIndex = 0 Then
            Ac_Code = ""

        ElseIf BalanceType.SelectedIndex = 1 Then
            Ac_Code = "And (Left(Ac_Code,1) = '4' or Left(Ac_Code,1) = '5') "
            Ac_Code = "And (Left(Ac_Code,1) = '4' or Left(Ac_Code,1) = '5' or Left(Ac_Code,4) = '00.4' or Left(Ac_Code,4) = '00.5') "
            'Ac_Code = "And (Left(Ac_Code,4) = '00.4' or Left(Ac_Code,4) = '00.5') "
        End If
        If CheckBox1.Checked = True Then
            ChangInCom = 1
        Else
            ChangInCom = 0
        End If


        New_Code = "3901000.00.0000"
        Code_Dr = "4"
        Code_Dr1 = "00.4"
        'Code_Dr = "00." & Code_Dr
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        'Call ChangBalance()

        BLNEW()
        'If CMB_Curr.SelectedIndex = 0 Then
        '    Call ChangBalance()
        'Else
        '    Call BLNEW()
        'End If
        'If CheckBnk.Checked = True Then
        '    CNN.Execute("delete Ap_balance_6_col from Ap_balance_6_col , Ap_Rpt_BLS_Item where Ap_balance_6_col.ac_code=Ap_Rpt_BLS_Item.Ac_Code")
        '    CNN.Execute("delete Ap_balance_6_col from Ap_balance_6_col , Ap_Rpt_Income_Item where Ap_balance_6_col.ac_code=Ap_Rpt_Income_Item.Ac_Code")
        'End If
        'If CheckBox2.Checked = True Then
        '    CNN.Execute("delete Ap_balance_6_col where Amt_Dr=0 And Amt_Cr=0 ")
        'End If
        'If MuLng = "L" Then
        '    CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        'Else
        '    CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_E from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        'End If

        If CheckBnk.Checked = True Then
            CNN.Execute("delete Ap_balance_6_col from Ap_balance_6_col , Ap_Rpt_BLS_Item where Ap_balance_6_col.ac_code=Ap_Rpt_BLS_Item.Ac_Code")
            CNN.Execute("delete Ap_balance_6_col from Ap_balance_6_col , Ap_Rpt_Income_Item where Ap_balance_6_col.ac_code=Ap_Rpt_Income_Item.Ac_Code")
        End If
        If CheckBox2.Checked = True Then
            CNN.Execute("delete Ap_balance_6_col where Amt_Dr=0 And Amt_Cr=0 ")
        End If
        If TextBox1.Text <> "" Then
            CNN.Execute("delete Ap_balance_6_col where Left(Ac_Code, " & Len(TextBox1.Text) & ") <> '" & TextBox1.Text & "' ")
        End If
        If MuLng = "L" Then
            CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        Else
            CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_E from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        End If

        Call LoadReport()

    End Sub

  
    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged

    End Sub
End Class
