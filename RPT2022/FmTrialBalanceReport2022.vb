Public Class FmTrialBalanceReport2022
    Dim MonthLetter1 As String


    'Dim MdStartDate As Date
    'Dim MdToDate As Date
    Dim sql As String
    Dim VCode1, VCode2, VCode3, VCode4, VCode5, VCode6, VCode7, VCode8, VCode9 As String
    ' Migrated from ADODB to ADO.NET
    ' Dim RsOpen As New ADODB.Recordset
    ' Dim RsOpenMonth As New ADODB.Recordset
    ' Dim RsRpt As New ADODB.Recordset
    Dim AmtOpenDR, AmtOpenCR, AmtOpenMonthDR, AmtOpenMonthCR As Double
    Dim VOpenDate As Date
    Dim RptNme As String
    ' Migrated from ADODB to ADO.NET
    ' Dim RSC12 As New ADODB.Recordset
    Dim d, p As String

    ' Migrated from ADODB to ADO.NET
    ' Dim RSCP As New ADODB.Recordset

    Private Sub HeaDer()
        Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM Header WHERE ID=N'BL01' ")
        If dt.Rows.Count > 0 Then
            TxtHeader.Text = Trim(dt.Rows(0)("Nm").ToString())
            TxtS1.Text = Trim(dt.Rows(0)("S1").ToString())
            TxtS2.Text = Trim(dt.Rows(0)("S2").ToString())
            TxtS3.Text = Trim(dt.Rows(0)("S3").ToString())
            TxtS4.Text = Trim(dt.Rows(0)("S4").ToString())
            TxtPP.Text = Trim(dt.Rows(0)("pp").ToString())
        End If
    End Sub
    Private Sub AddHeader()
        Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM Header WHERE ID=N'BL01' ")
        If dt.Rows.Count = 0 Then
            DbHelper.ExecuteNonQuery("INSERT INTO Header(ID,Nm,S1,S2,S3,S4,PP) " & _
                        " values('BL01',N'" & TxtHeader.Text & "',N'" & TxtS1.Text & "',N'" & TxtS2.Text & "',N'" & TxtS3.Text & "',N'" & TxtS4.Text & "',N'" & TxtPP.Text & "') ")
        Else
            DbHelper.ExecuteNonQuery("UPDATE Header set Nm=N'" & TxtHeader.Text & "',S1=N'" & TxtS1.Text & "',S2=N'" & TxtS2.Text & "',S3=N'" & TxtS3.Text & "',S4=N'" & TxtS4.Text & "',PP=N'" & TxtPP.Text & "' " & _
                        " where ID='BL01' ")
        End If
    End Sub
    Private Sub DMonth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DMonth.SelectedIndexChanged
        LoadMonth()


    End Sub

    Private Sub LoadMonth()
        '---------------------------------

        If DMonth.SelectedIndex = 0 Then

        End If

        If DMonth.SelectedIndex = 0 Then
            MdStartDate = Format(CDate("01/01/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/01/" & Year(Myy.Value)), "dd-MM-yyyy")
            LngId = "7013" : CallLngStr() : MonthLetter1 = LngStr
        ElseIf DMonth.SelectedIndex = 1 Then
            Dim Day As String
            Dim MM As Date
            Dim Fromm As Date
            MdStartDate = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
            Fromm = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
            MM = Format(CDate("01/3/" & Year(MdStartDate)), "dd-MM-yyyy")
            Day = DateDiff(DateInterval.Day, Fromm, MM)
            MdToDate = Format(CDate(Day & "/02" & "/" & Year(MdStartDate)), "dd-MM-yyyy")
            MonthLetter1 = "ກຸມພາ"
            LngId = "7014" : CallLngStr() : MonthLetter1 = LngStr
        ElseIf DMonth.SelectedIndex = 2 Then
            MdStartDate = Format(CDate("01/03/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/03/" & Year(Myy.Value)), "dd-MM-yyyy")
            LngId = "7015" : CallLngStr() : MonthLetter1 = LngStr
        ElseIf DMonth.SelectedIndex = 3 Then
            MdStartDate = Format(CDate("01/04/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/04/" & Year(Myy.Value)), "dd-MM-yyyy")
            LngId = "7016" : CallLngStr() : MonthLetter1 = LngStr
        ElseIf DMonth.SelectedIndex = 4 Then
            MdStartDate = Format(CDate("01/05/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/05/" & Year(Myy.Value)), "dd-MM-yyyy")
            LngId = "7017" : CallLngStr() : MonthLetter1 = LngStr
        ElseIf DMonth.SelectedIndex = 5 Then
            MdStartDate = Format(CDate("01/06/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/06/" & Year(Myy.Value)), "dd-MM-yyyy")
            LngId = "7018" : CallLngStr() : MonthLetter1 = LngStr
        ElseIf DMonth.SelectedIndex = 6 Then
            MdStartDate = Format(CDate("01/07/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/07/" & Year(Myy.Value)), "dd-MM-yyyy")
            LngId = "7019" : CallLngStr() : MonthLetter1 = LngStr
        ElseIf DMonth.SelectedIndex = 7 Then
            MdStartDate = Format(CDate("01/08/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/08/" & Year(Myy.Value)), "dd-MM-yyyy")
            LngId = "7020" : CallLngStr() : MonthLetter1 = LngStr
        ElseIf DMonth.SelectedIndex = 8 Then
            MdStartDate = Format(CDate("01/09/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/09/" & Year(Myy.Value)), "dd-MM-yyyy")
            LngId = "7021" : CallLngStr() : MonthLetter1 = LngStr
        ElseIf DMonth.SelectedIndex = 9 Then
            MdStartDate = Format(CDate("01/10/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/10/" & Year(Myy.Value)), "dd-MM-yyyy")
            LngId = "7022" : CallLngStr() : MonthLetter1 = LngStr
        ElseIf DMonth.SelectedIndex = 10 Then
            MdStartDate = Format(CDate("01/11/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/11/" & Year(Myy.Value)), "dd-MM-yyyy")
            LngId = "7023" : CallLngStr() : MonthLetter1 = LngStr

        ElseIf DMonth.SelectedIndex = 11 Then
            MdStartDate = Format(CDate("01/12/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/12/" & Year(Myy.Value)), "dd-MM-yyyy")
            LngId = "7024" : CallLngStr() : MonthLetter1 = LngStr
        End If
        '-----------------
        Dim m, y As String
        LngId = "7049" : CallLngStr() : m = LngStr & " "
        LngId = "7025" : CallLngStr() : y = LngStr & " "

        Lb.Text = m & MonthLetter1 & "/" & Year(MdToDate)
        L5.Text = MdStartDate & " => " & MdToDate
        Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
        DTDATE = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)

        'If DMonth.SelectedIndex = "01" Then
        '    MonthLetter1 = "ມັງກອນ"
        'ElseIf DMonth.SelectedIndex = "02" Then
        '    MonthLetter1 = "ກຸມພາ"
        'ElseIf DMonth.SelectedIndex = "03" Then
        '    MonthLetter1 = "ມີນາ"
        'ElseIf DMonth.SelectedIndex = "04" Then
        '    MonthLetter1 = "ເມສາ"
        'ElseIf DMonth.SelectedIndex = "05" Then
        '    MonthLetter1 = "ພຶດສະພາ"
        'ElseIf DMonth.SelectedIndex = "06" Then
        '    MonthLetter1 = "ມີຖຸນາ"
        'ElseIf DMonth.SelectedIndex = "07" Then
        '    MonthLetter1 = "ກໍລະກົດ"
        'ElseIf DMonth.SelectedIndex = "08" Then
        '    MonthLetter1 = "ສິງຫາ"
        'ElseIf DMonth.SelectedIndex = "09" Then
        '    MonthLetter1 = "ກັນຍາ"
        'ElseIf DMonth.SelectedIndex = "10" Then
        '    MonthLetter1 = "ຕຸລາ"
        'ElseIf DMonth.SelectedIndex = "11" Then
        '    MonthLetter1 = "ພະຈິກ"
        'ElseIf DMonth.SelectedIndex = "12" Then
        '    MonthLetter1 = "ທັນວາ"
        'End If
        'DTDATE = "ປະຈຳເດືອນ " & DMonth.Text & " ປີ " & Year(MdToDate)
        DTDATE = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳເດືອນ " & DMonth.Text & " ປີ " & Year(MdToDate)
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
        'DTDATE = "ປະຈຳງວດ " & CDbl(Period.SelectedIndex) + 1 & " ປີ " & Pyy.Text
        DTDATE = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳງວດ " & CDbl(Period.SelectedIndex) + 1 & " ປີ " & Pyy.Text
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
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາວັນທີ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")

        L5.Text = MdStartDate & " => " & MdToDate
        DTDATE = "ແຕ່ວັນທີ " & MdStartDate & " ຫາ " & MdToDate
        DTDATE = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        ' DTDATE02 = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາວັນທີ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳວັນທີ"
    End Sub

    Private Sub DateTimePicker5_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yy.ValueChanged
        Call LoadYear()
    End Sub
    Dim DTDATE As String
    Dim DTDATE02 As String
    Private Sub LoadYear()
        Dim y As String
        LngId = "7025" : CallLngStr() : y = LngStr & ": "
        MdStartDate = Format(CDate("01/1/" & Year(yy.Value)), "dd-MM-yyyy")
        MdToDate = Format(CDate("31/12/" & Year(yy.Value)), "dd-MM-yyyy")
        Lb.Text = y & yy.Text
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd/MM/yyyy") & " ສີ້ນສຸດວັນທີ " & Format(MdToDate, "dd/MM/yyyy")
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ສີ້ນສຸດວັນທີ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")

        L5.Text = MdStartDate & " => " & MdToDate
        'DTDATE = "ປະຈຳປີ " & yy.Text
        'DTDATE = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd/MM/yyyy") & " ສີ້ນສຸດວັນທີ " & Format(MdToDate, "dd/MM/yyyy")
        'DTDATE = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ສີ້ນສຸດວັນທີ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳປີ  " & yy.Text
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
        Dim dt As DataTable = DbHelper.GetDataTable("select sub_id , off_add2  from  Ap_office  Order by sub_id")
        For Each row As DataRow In dt.Rows
            Off_Usr.Items.Add((row("sub_id").ToString()) & " " & (row("off_add2").ToString()))
        Next
        Off_Usr.Text = FmLogin.Sub_Company.Text
    End Sub
    
    Private Sub SetupGrid()
        FG.Columns.Clear()
        FG.Columns.Add("Col1", "ລ/ດ")
        FG.Columns.Add("Col2", "ລະຫັດບັນຊີ")
        FG.Columns.Add("Col3", "ຍອດຍົກເບື້ອງ (ຫນີ້)")
        FG.Columns.Add("Col4", "ຍອດຍົກເບື້ອງ (ມີ)")
        FG.Columns.Add("Col5", "ການເຄື່ອນໄຫວ (ຫນີ້)")
        FG.Columns.Add("Col6", "ການເຄື່ອນໄຫວ (ມີ)")
        FG.Columns.Add("Col7", "ຍອດເຫລືອ (ຫນີ້)")
        FG.Columns.Add("Col8", "ຍອດເຫລືອ (ມີ)")
        FG.Columns.Add("Col9", "Count")
        
        FG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        FG.AllowUserToAddRows = False
        FG.AllowUserToDeleteRows = False
        FG.ReadOnly = True
    End Sub
    Private Sub FmTrialBalanceReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        HeaDer()
        Call loadOffice_User()
        Cx.SelectedIndex = 0
        BalanceType.SelectedIndex = 0
        FG.AllowUserToResizeColumns = True
        FG.AllowUserToResizeRows = True
        Period.Text = "ງວດທີ 1"
        DMonth.Text = "ມັງກອນ"
        LoadMonth()
        SetupGrid()



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

        selectLoad()
        'Call AddData()
        'Call LoadListFG()
        'Call CaRemain()
        SetControlText(Me)
        RGL.Text = "ແບບທົ່ວໄປ"
        RDtail.Text = "ຕາມຫມວດບັນຊີ"
        RGroup.Text = "ສະເພາະບັນຊີແມ່"
        CMB_Curr.Items.Clear()
        CMB_Curr.Items.Add("EQVL")
        Call load_Cmb(" SELECT Curr  FROM Curr_For_Rate WHERE (Curr='LAK' Or Curr='THB'  Or Curr='USD')  ORDER BY cnt ", "Curr", CMB_Curr)
        If CMB_Curr.Items.Count > 0 Then
            CMB_Curr.SelectedIndex = 0
        End If
        CheckBox3.Text = "ໃບດູນດ່ຽງປະຈຳປີ ຫລັງການປັບປຸງ"
        If CMB_Curr.Text = "LAK" Then
            CheckBox4.Text = "ທຽບເທົ່າໂດລາ"
        ElseIf CMB_Curr.Text = "USD" Then
            CheckBox4.Text = "ທຽບເທົ່າກີບ"

        Else
            CheckBox4.Text = "ທຽບເທົ່າໂດລາ"

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
        'DTDATE = "ປະຈຳ " & Ct.Text & " " & yyt.Text
        DTDATE = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳ " & Ct.Text & " " & yyt.Text
        'L5.Text = MdStartDate & " => " & MdToDate
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
        Call Office()
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
            'Lb.Text = s2 & "/" & Year(MdToDate)
        ElseIf RP.Checked = True Then
            Dim s1, s2, s3 As String
            LngId = 7074 + Period.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7062 : CallLngStr() : s1 = LngStr
            LngId = 7064 : CallLngStr() : s3 = LngStr
            Lb.Text = s1 & "  " & s2 & " " & " " & Year(MdToDate)
            Lb.Text = s2 & " " & " " & Year(MdToDate)

        ElseIf RY.Checked = True Then
            Dim s1, s2, s3 As String
            LngId = 7074 + Period.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7062 : CallLngStr() : s1 = LngStr
            LngId = 7064 : CallLngStr() : s3 = LngStr
            'Lb.Text = s3 & " " & yy.Text
            'Lb.Text = yy.Text
        End If
        If BalanceType.SelectedIndex = 0 Then
            Ac_Code = ""
        ElseIf BalanceType.SelectedIndex = 1 Then
            'Ac_Code = "And (Left(Ac_Code,1) = '4' or Left(Ac_Code,1) = '5') "
            'Ac_Code = "And (Left(Ac_Code,4) = '00.4' or Left(Ac_Code,4) = '00.5') "
            Ac_Code = "And (Left(Ac_Code,1) = '4' or Left(Ac_Code,1) = '5') "
        End If
        New_Code = "3901000"
        ' New_Code = "3901"
        Code_Dr = "4"
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        'Call ChangBalance()
        DbHelper.ExecuteNonQuery("DELETE  Ap_balance_6_col ")
        DbHelper.ExecuteNonQuery("DELETE FROM Ap_balance_6 ")
        Call BLNEW()
         
        If CheckBnk.Checked = True Then
            DbHelper.ExecuteNonQuery("delete Ap_balance_6_col from Ap_balance_6_col , Ap_Rpt_BLS_Item where Ap_balance_6_col.ac_code=Ap_Rpt_BLS_Item.Ac_Code")
            DbHelper.ExecuteNonQuery("delete Ap_balance_6_col from Ap_balance_6_col , Ap_Rpt_Income_Item where Ap_balance_6_col.ac_code=Ap_Rpt_Income_Item.Ac_Code")
        End If
        If CheckBox2.Checked = True Then
            DbHelper.ExecuteNonQuery("delete Ap_balance_6_col where Amt_Dr=0 And Amt_Cr=0 ")
        End If
        If CheckBox1.Checked = True Then
            Dim aa As String
 
            aa = " delete  RPT_Balance_5_4"
            DbHelper.ExecuteNonQuery(aa)
            aa = "   insert into  RPT_Balance_5_4 (ac_code,open_amt_dr,open_amt_cr,amt_dr,amt_cr,close_amt_dr,close_amt_cr,lck, Curr) " & _
     "   select LEFT (ac_code,1),0,0,SUM(amt_dr),SUM(amt_cr),0,0,0, Curr from Ap_balance_6_col " & _
     "  where  left(ac_code,'1')>='4' and  left(ac_code,'1')<='5'  group by LEFT (ac_code,1), Curr "
            DbHelper.ExecuteNonQuery(aa)

            aa = "  insert into  RPT_Balance_5_4 (ac_code,open_amt_dr,open_amt_cr,amt_dr,amt_cr,close_amt_dr,close_amt_cr,lck, Curr) " & _
 "  select 1,SUM(open_amt_dr),SUM(open_amt_cr ),SUM(amt_dr ),SUM(amt_cr),SUM(close_amt_dr ),SUM(close_amt_cr ),1, Curr from RPT_Balance_5_4 group by  Curr "
            DbHelper.ExecuteNonQuery(aa)
            aa = " delete  RPT_Balance_5_4 where lck=0  "
            DbHelper.ExecuteNonQuery(aa)
            aa = " delete  RPT_Barande2 "
            DbHelper.ExecuteNonQuery(aa)
            aa = "  insert into  RPT_Barande2 (ac_code,open_amt_dr,open_amt_cr,amt_dr,amt_cr,close_amt_dr,close_amt_cr,lck, Curr) " & _
                   "  select '5_4',0,0,  amt_dr -amt_cr ,0,0,0,0, Curr from RPT_Balance_5_4 where  amt_cr- amt_dr<0  "
            DbHelper.ExecuteNonQuery(aa)
            aa = "  insert into  RPT_Barande2 (ac_code,open_amt_dr,open_amt_cr,amt_cr,amt_dr,close_amt_dr,close_amt_cr,lck, Curr) " & _
                "  select '5_4',0,0, amt_cr- amt_dr,0,0,0,0, Curr from RPT_Balance_5_4 where  amt_cr- amt_dr>=0  "
            DbHelper.ExecuteNonQuery(aa)
            aa = " delete  RPT_Balance_5_4   "
            DbHelper.ExecuteNonQuery(aa)
            aa = "  insert into  RPT_Balance_5_4 (ac_code,open_amt_dr,open_amt_cr,amt_dr,amt_cr,close_amt_dr,close_amt_cr,lck, Curr) " & _
               "  select 3901000,open_amt_dr,open_amt_cr,amt_dr,amt_cr,close_amt_dr,close_amt_cr,0, Curr from RPT_Barande2   "
            DbHelper.ExecuteNonQuery(aa)

            aa = "   insert into  Ap_balance_6_col (ac_code,open_amt_dr,open_amt_cr,amt_dr,amt_cr,Rem_dr,Rem_cr, Curr) " & _
      "   select  ac_code,open_amt_dr,open_amt_cr,amt_dr,amt_cr,close_amt_dr,close_amt_cr, Curr  from RPT_Balance_5_4 "
            DbHelper.ExecuteNonQuery(aa)
            aa = " delete  Ap_balance_6_col where  LEFT (ac_code,1)='4'   "
            DbHelper.ExecuteNonQuery(aa)
            aa = " delete  Ap_balance_6_col where  LEFT (ac_code,1)='5'   "
            DbHelper.ExecuteNonQuery(aa)
 
        End If
        If MuLng = "L" Then
            DbHelper.ExecuteNonQuery("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        Else
            DbHelper.ExecuteNonQuery("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_E from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        End If

        Call LoadReport()
    End Sub
    Private Sub BLNEW()
        DbHelper.ExecuteNonQuery("   update gen_jn set Rate_USD=rate where curr='USD' and Rate_USD is null ")
        DbHelper.ExecuteNonQuery("  update gen_jn set Rate_USD=0 where   Rate_USD is null ")
        DbHelper.ExecuteNonQuery("  update gen_jn set amt_USD_Dr=amount_dr  where curr='USD' and (amt_USD_Dr=0 or amt_USD_Dr is null) ")
        DbHelper.ExecuteNonQuery("  update gen_jn set amt_USD_cr= amount_Cr   where curr='USD'  and (amt_USD_cr=0 or amt_USD_cr is null) ")
        DbHelper.ExecuteNonQuery("  update gen_jn set amt_USD_Dr= amt_dr/Rate_USD  where curr='LAK' and Rate_USD<>0")
        DbHelper.ExecuteNonQuery("  update gen_jn set amt_USD_cr= amt_cr/Rate_USD    where curr='LAK'  and Rate_USD<>0")
        '==============OPEN=====
        DbHelper.ExecuteNonQuery("   update Open_jn set Rate_USD=rate where curr='USD' and Rate_USD is null ")
        DbHelper.ExecuteNonQuery("  update Open_jn set Rate_USD=0 where   Rate_USD is null ")
        DbHelper.ExecuteNonQuery("  update Open_jn set amt_USD_Dr=amount_dr  where curr='USD' and (amt_USD_Dr=0 or amt_USD_Dr is null) ")
        DbHelper.ExecuteNonQuery("  update Open_jn set amt_USD_cr= amount_Cr   where curr='USD'  and (amt_USD_cr=0 or amt_USD_cr is null) ")

        DbHelper.ExecuteNonQuery("  update Open_jn set amt_USD_Dr= amt_dr/Rate_USD  where curr='LAK' and Rate_USD<>0")
        DbHelper.ExecuteNonQuery("  update Open_jn set amt_USD_cr= amt_cr/Rate_USD    where curr='LAK'  and Rate_USD<>0")
        '==============Adjust=====
        DbHelper.ExecuteNonQuery("   update AP_ACC_adjust_Item set Rate_USD=rate where curr='USD' and Rate_USD is null ")
        DbHelper.ExecuteNonQuery("  update AP_ACC_adjust_Item set Rate_USD=0 where   Rate_USD is null ")
        DbHelper.ExecuteNonQuery("  update AP_ACC_adjust_Item set amt_USD_Dr=amount_dr  where curr='USD' and (amt_USD_Dr=0 or amt_USD_Dr is null) ")
        DbHelper.ExecuteNonQuery("  update AP_ACC_adjust_Item set amt_USD_cr= amount_Cr   where curr='USD'  and (amt_USD_cr=0 or amt_USD_cr is null) ")

        DbHelper.ExecuteNonQuery("  update AP_ACC_adjust_Item set amt_USD_Dr= amt_dr/Rate_USD  where curr='LAK' and Rate_USD<>0")
        DbHelper.ExecuteNonQuery("  update AP_ACC_adjust_Item set amt_USD_cr= amt_cr/Rate_USD    where curr='LAK'  and Rate_USD<>0")

        '=================NEWWWWW==============
        DbHelper.ExecuteNonQuery("DELETE  Ap_balance_6_col ")
        DbHelper.ExecuteNonQuery("DELETE FROM Ap_balance_6 ")
        Dim B_Curr As String = ""
        If CMB_Curr.SelectedIndex = 0 Then
            B_Curr = ""
        Else
            B_Curr = " AND  Curr=N'" & CMB_Curr.Text & "' "
        End If
        DbHelper.ExecuteNonQuery("UPDATE gen_jn set amt_USD_Dr=0 where amt_USD_Dr is null ")
        DbHelper.ExecuteNonQuery("UPDATE gen_jn set amt_USD_Cr=0 where amt_USD_Cr is null ")
        DbHelper.ExecuteNonQuery("UPDATE Open_jn set amt_USD_Dr=0 where amt_USD_Dr is null ")
        DbHelper.ExecuteNonQuery("UPDATE Open_jn set amt_USD_Cr=0 where amt_USD_Cr is null ")
        DbHelper.ExecuteNonQuery("UPDATE AP_ACC_adjust_Item set amt_USD_Dr=0 where amt_USD_Dr is null ")
        DbHelper.ExecuteNonQuery("UPDATE AP_ACC_adjust_Item set amt_USD_Cr=0 where amt_USD_Cr is null ")
 
        'CheckBox4.Text = "ທຽບເທົ່າໂດລາ"
        'CheckBox5.Text = "ທຽບເທົ່າບາດ"
        'Label23.Text = "ບາດ-ກີບ" 1
        'Label24.Text = "ໂດລາ-ກີບ" 2

        If CMB_Curr.SelectedIndex = 0 Then
            If CheckBox4.Checked = True Then
                '=======LAK===
                Dim LAK As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
               " select ac_code, 'USD', 0, 0 , sum(amt_Dr) / " & CDbl(txtRate2.Text) & " , sum(amt_Cr) / " & CDbl(txtRate2.Text) & "  from gen_jn  " & _
               " WHERE 1=1 and date_work BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code "
                DbHelper.ExecuteNonQuery(LAK)
                ' Dim THB As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '" select ac_code, Curr, 0, 0 , sum(amt_Dr) / " & CDbl(txtRate.Text) & " , sum(amt_Cr) / " & CDbl(txtRate.Text) & "  from gen_jn  " & _
                '" WHERE 1=1 AND  Curr=N'THB' and date_work BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code, Curr "
                ' DbHelper.ExecuteNonQuery(THB)
                ' Dim USD As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '" select ac_code, Curr, 0, 0 , sum(amount_Dr) , sum(amount_Cr)  from gen_jn  " & _
                '" WHERE 1=1 AND  Curr=N'USD' and date_work BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code, Curr "
                ' DbHelper.ExecuteNonQuery(USD)

                Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
                '=======LAK===
                Dim OLAK As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code, 'USD', sum(amt_Dr) / " & CDbl(txtRate2.Text) & " , sum(amt_cr) / " & CDbl(txtRate2.Text) & " , 0 , 0  from gen_jn  " & _
                " WHERE 1=1 and date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code"
                DbHelper.ExecuteNonQuery(OLAK)
                DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code, Curr, open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
             " select ac_code, 'USD', sum(amt_Dr) / " & CDbl(txtRate2.Text) & " , sum(amt_cr) / " & CDbl(txtRate2.Text) & " , 0, 0 from Open_jn " & _
             " WHERE 1=1 and   date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
                '   '=======THB===
                '   Dim OTHB As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '   " select ac_code, Curr, sum(amt_Dr) / " & CDbl(txtRate.Text) & " , sum(amt_Cr) / " & CDbl(txtRate.Text) & " , 0 , 0  from gen_jn  " & _
                '   " WHERE 1=1 AND Curr=N'THB' and date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code, Curr"
                '   DbHelper.ExecuteNonQuery(OTHB)
                '   DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code, Curr, open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '" select ac_code, Curr, sum(amt_Dr) / " & CDbl(txtRate.Text) & " , sum(amt_Cr) / " & CDbl(txtRate.Text) & " , 0, 0 from Open_jn " & _
                '" WHERE 1=1 AND Curr=N'THB' and   date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code, Curr")
                '    '=======USD===
                '    Dim OUSD As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '    " select ac_code, Curr, sum(amount_Dr) , sum(amount_Cr) , 0 , 0  from gen_jn  " & _
                '    " WHERE 1=1 AND Curr=N'USD' and date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code, Curr"
                '    DbHelper.ExecuteNonQuery(OUSD)
                '    DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code, Curr, open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '" select ac_code, Curr, sum(amount_Dr) , sum(amount_Cr) , 0, 0 from Open_jn " & _
                '" WHERE 1=1 AND Curr=N'USD' and   date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code, Curr")
                '=======LAK===
             
                If CheckBox3.Checked = True Then
                    LAK = "  insert into  Ap_balance_6 (ac_code, Curr, amt_dr, amt_cr) " & _
                    " select ac_code, 'USD', sum(amt_Dr) / " & CDbl(txtRate2.Text) & " , sum(amt_Cr) / " & CDbl(txtRate2.Text) & "  from AP_ACC_adjust_Item " & _
                    " where 1=1 and year(date_work) = '" & Format(CDate(MdStartDate), "yyyy") & "' " & MULook2 & " group by ac_code "
                    DbHelper.ExecuteNonQuery(LAK)
                    'THB = "  insert into  Ap_balance_6 (ac_code, Curr, amt_dr, amt_cr) " & _
                    '" select ac_code, Curr, sum(amt_Dr) / " & CDbl(txtRate.Text) & " , sum(amt_Cr) / " & CDbl(txtRate.Text) & "  from AP_ACC_adjust_Item " & _
                    '" where 1=1 AND Curr=N'THB' and year(date_work) = '" & Format(CDate(MdStartDate), "yyyy") & "' " & MULook2 & " group by ac_code, Curr "
                    'DbHelper.ExecuteNonQuery(THB)
                    'USD = "  insert into  Ap_balance_6 (ac_code, Curr, amt_dr, amt_cr) " & _
                    '" select ac_code, Curr, sum(amount_Dr) , sum(amount_Cr) from AP_ACC_adjust_Item " & _
                    '" where 1=1 AND Curr=N'USD' and year(date_work) = '" & Format(CDate(MdStartDate), "yyyy") & "' " & MULook2 & " group by ac_code, Curr "
                    'DbHelper.ExecuteNonQuery(USD)

                End If

            ElseIf CheckBox5.Checked = True Then
                'CheckBox4.Text = "ທຽບເທົ່າໂດລາ"
                'CheckBox5.Text = "ທຽບເທົ່າບາດ"
                'Label23.Text = "ບາດ-ກີບ" 1
                'Label24.Text = "ໂດລາ-ກີບ" 2
                '=======LAK===
                Dim LAK As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
               " select ac_code, 'THB', 0, 0 , sum(amt_Dr) / " & CDbl(txtRate.Text) & " , sum(amt_Cr) / " & CDbl(txtRate.Text) & "  from gen_jn  " & _
               " WHERE 1=1 and date_work BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code "
                DbHelper.ExecuteNonQuery(LAK)
                ' Dim USD As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '" select ac_code, Curr, 0, 0 , sum(amt_Dr) / " & CDbl(txtRate.Text) & " , sum(amt_Cr) / " & CDbl(txtRate.Text) & "  from gen_jn  " & _
                '" WHERE 1=1 AND  Curr=N'USD' and date_work BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code, Curr "
                ' DbHelper.ExecuteNonQuery(USD)
                ' Dim THB As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '" select ac_code, Curr, 0, 0 , sum(amount_Dr) , sum(amount_Cr)  from gen_jn  " & _
                '" WHERE 1=1 AND  Curr=N'THB' and date_work BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code, Curr "
                ' DbHelper.ExecuteNonQuery(THB)

                Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
                '=======LAK===
                Dim OLAK As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code, 'THB', sum(amt_Dr) / " & CDbl(txtRate.Text) & " , sum(amt_cr) / " & CDbl(txtRate.Text) & " , 0 , 0  from gen_jn  " & _
                " WHERE 1=1 and date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code"
                DbHelper.ExecuteNonQuery(OLAK)
                DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code, Curr, open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
             " select ac_code, 'THB', sum(amt_Dr) / " & CDbl(txtRate.Text) & " , sum(amt_cr) / " & CDbl(txtRate.Text) & " , 0, 0 from Open_jn " & _
             " WHERE 1=1 AND date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
                '    '=======USD===
                '    Dim OUSD As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '    " select ac_code, Curr, sum(amt_Dr) / " & CDbl(txtRate.Text) & " , sum(amt_Cr) / " & CDbl(txtRate.Text) & " , 0 , 0  from gen_jn  " & _
                '    " WHERE 1=1 AND Curr=N'USD' and date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code, Curr"
                '    DbHelper.ExecuteNonQuery(OUSD)
                '    DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code, Curr, open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                ' " select ac_code, Curr, sum(amt_Dr) / " & CDbl(txtRate.Text) & " , sum(amt_Cr) / " & CDbl(txtRate.Text) & " , 0, 0 from Open_jn " & _
                ' " WHERE 1=1 AND Curr=N'USD' and   date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code, Curr")
                '    '=======THB===
                '    Dim OTHB As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '    " select ac_code, Curr, sum(amount_Dr) , sum(amount_Cr) , 0 , 0  from gen_jn  " & _
                '    " WHERE 1=1 AND Curr=N'THB' and date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code, Curr"
                '    DbHelper.ExecuteNonQuery(OTHB)
                '    DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code, Curr, open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '" select ac_code, Curr, sum(amount_Dr) , sum(amount_Cr) , 0, 0 from Open_jn " & _
                '" WHERE 1=1 AND Curr=N'THB' and   date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code, Curr")
                '=======LAK===

                If CheckBox3.Checked = True Then
                    LAK = "  insert into  Ap_balance_6 (ac_code, Curr, amt_dr, amt_cr) " & _
                    " select ac_code, 'THB', sum(amt_Dr) / " & CDbl(txtRate.Text) & " , sum(amt_Cr) / " & CDbl(txtRate.Text) & "  from AP_ACC_adjust_Item " & _
                    " where 1=1 AND year(date_work) = '" & Format(CDate(MdStartDate), "yyyy") & "' " & MULook2 & " group by ac_code "
                    DbHelper.ExecuteNonQuery(LAK)
                    'USD = "  insert into  Ap_balance_6 (ac_code, Curr, amt_dr, amt_cr) " & _
                    '" select ac_code, Curr, sum(amt_Dr) / " & CDbl(txtRate.Text) & " , sum(amt_Cr) / " & CDbl(txtRate.Text) & "  from AP_ACC_adjust_Item " & _
                    '" where 1=1 AND Curr=N'USD' and year(date_work) = '" & Format(CDate(MdStartDate), "yyyy") & "' " & MULook2 & " group by ac_code, Curr "
                    'DbHelper.ExecuteNonQuery(USD)
                    'THB = "  insert into  Ap_balance_6 (ac_code, Curr, amt_dr, amt_cr) " & _
                    '" select ac_code, Curr, sum(amount_Dr) , sum(amount_Cr) from AP_ACC_adjust_Item " & _
                    '" where 1=1 AND Curr=N'THB' and year(date_work) = '" & Format(CDate(MdStartDate), "yyyy") & "' " & MULook2 & " group by ac_code, Curr "
                    'DbHelper.ExecuteNonQuery(THB)

                End If
            Else
                '=======LAK===
                Dim GGG As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
               " select ac_code, 'LAK', 0, 0 , sum(Amt_Dr) , sum(Amt_cr) from gen_jn " & _
               " WHERE 1=1 and date_work BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code "
                DbHelper.ExecuteNonQuery(GGG)
  
                Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
                '=======LAK===
                Dim OLAK As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                       " select ac_code, 'LAK' , sum(Amt_Dr) , sum(Amt_cr), 0, 0 from gen_jn " & _
                        " WHERE 1=1  and date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code "
                DbHelper.ExecuteNonQuery(OLAK)
                '        '=======LAK===
                DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code, Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code , 'LAK' , sum(Amt_Dr) , sum(Amt_cr) , 0, 0 from Open_jn " & _
                " WHERE 1=1 and date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code ")
            
                If CheckBox3.Checked = True Then
                    Dim aa As String
                    aa = "  insert into  Ap_balance_6 (ac_code, Curr, amt_dr, amt_cr) " & _
                     " select ac_code, 'LAK', sum(Amt_Dr) , sum(Amt_cr)  from AP_ACC_adjust_Item " & _
                     " where  year(AP_ACC_adjust_Item.date_work) = '" & Format(CDate(MdStartDate), "yyyy") & "' " & MULook2 & "  " & _
                     " group by  ac_code ,com_id   "
                    DbHelper.ExecuteNonQuery(aa)


                End If
            End If


        Else
          
            Dim GGG As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            " select ac_code, Curr, 0, 0, sum(Amount_Dr), sum(Amount_cr) from gen_jn  WHERE 1=1 " & B_Curr & " and date_work BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code, Curr  "
            DbHelper.ExecuteNonQuery(GGG)

            Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
            Dim PPP As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            " select ac_code, Curr, sum(Amount_Dr), sum(Amount_cr), 0, 0 from gen_jn  WHERE 1=1 " & B_Curr & " and date_work BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code, Curr "
            DbHelper.ExecuteNonQuery(PPP)
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            " select ac_code, Curr, sum(Amount_Dr), sum(Amount_cr), 0, 0 from Open_jn WHERE  1=1 " & B_Curr & " and   date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code, Curr ")
            If CheckBox3.Checked = True Then
                Dim aa As String
                aa = "  insert into  Ap_balance_6 (ac_code, Curr,open_amt_dr,open_amt_cr,amt_dr,amt_cr)   SELECT  AP_ACC_adjust_Item.ac_code, AP_ACC_adjust_Item.Curr,0,0,SUM (AP_ACC_adjust_Item.Amount_Dr),SUM (AP_ACC_adjust_Item.Amount_cr)  from  AP_ACC_adjust_Item " & _
                        "   where  year(AP_ACC_adjust_Item.date_work) = '" & Format(CDate(MdStartDate), "yyyy") & "' " & MULook2 & "  " & _
                        " group by  ac_code, Curr ,com_id   "
                DbHelper.ExecuteNonQuery(aa)
            End If
        End If
        'End If

        DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6_col ( ac_code, Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code , Curr, sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr)as open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code, Curr ")
 
        Call Left_AcCode()
        DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        Call Chang_Incom()
        DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + amt_dr) - (open_amt_cr + amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr + amt_cr) ")
        DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr + amt_cr) - (open_amt_dr + amt_dr) where (open_amt_cr + amt_cr) >= (open_amt_dr + amt_dr) ")
        If MuLng = "L" Then
            DbHelper.ExecuteNonQuery("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        Else
            DbHelper.ExecuteNonQuery("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_E from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        End If



    End Sub
    Private Sub BLNEW22()
        DbHelper.ExecuteNonQuery("   update gen_jn set Rate_USD=rate where curr='USD' and Rate_USD is null ")
        DbHelper.ExecuteNonQuery("  update gen_jn set Rate_USD=0 where   Rate_USD is null ")
        DbHelper.ExecuteNonQuery("  update gen_jn set amt_USD_Dr=amount_dr  where curr='USD' and (amt_USD_Dr=0 or amt_USD_Dr is null) ")
        DbHelper.ExecuteNonQuery("  update gen_jn set amt_USD_cr= amount_Cr   where curr='USD'  and (amt_USD_cr=0 or amt_USD_cr is null) ")
        DbHelper.ExecuteNonQuery("  update gen_jn set amt_USD_Dr= amt_dr/Rate_USD  where curr='LAK' and Rate_USD<>0")
        DbHelper.ExecuteNonQuery("  update gen_jn set amt_USD_cr= amt_cr/Rate_USD    where curr='LAK'  and Rate_USD<>0")
        '==============OPEN=====
        DbHelper.ExecuteNonQuery("   update Open_jn set Rate_USD=rate where curr='USD' and Rate_USD is null ")
        DbHelper.ExecuteNonQuery("  update Open_jn set Rate_USD=0 where   Rate_USD is null ")
        DbHelper.ExecuteNonQuery("  update Open_jn set amt_USD_Dr=amount_dr  where curr='USD' and (amt_USD_Dr=0 or amt_USD_Dr is null) ")
        DbHelper.ExecuteNonQuery("  update Open_jn set amt_USD_cr= amount_Cr   where curr='USD'  and (amt_USD_cr=0 or amt_USD_cr is null) ")

        DbHelper.ExecuteNonQuery("  update Open_jn set amt_USD_Dr= amt_dr/Rate_USD  where curr='LAK' and Rate_USD<>0")
        DbHelper.ExecuteNonQuery("  update Open_jn set amt_USD_cr= amt_cr/Rate_USD    where curr='LAK'  and Rate_USD<>0")
        '==============Adjust=====
        DbHelper.ExecuteNonQuery("   update AP_ACC_adjust_Item set Rate_USD=rate where curr='USD' and Rate_USD is null ")
        DbHelper.ExecuteNonQuery("  update AP_ACC_adjust_Item set Rate_USD=0 where   Rate_USD is null ")
        DbHelper.ExecuteNonQuery("  update AP_ACC_adjust_Item set amt_USD_Dr=amount_dr  where curr='USD' and (amt_USD_Dr=0 or amt_USD_Dr is null) ")
        DbHelper.ExecuteNonQuery("  update AP_ACC_adjust_Item set amt_USD_cr= amount_Cr   where curr='USD'  and (amt_USD_cr=0 or amt_USD_cr is null) ")

        DbHelper.ExecuteNonQuery("  update AP_ACC_adjust_Item set amt_USD_Dr= amt_dr/Rate_USD  where curr='LAK' and Rate_USD<>0")
        DbHelper.ExecuteNonQuery("  update AP_ACC_adjust_Item set amt_USD_cr= amt_cr/Rate_USD    where curr='LAK'  and Rate_USD<>0")

        '=================NEWWWWW==============
        DbHelper.ExecuteNonQuery("DELETE  Ap_balance_6_col ")
        DbHelper.ExecuteNonQuery("DELETE FROM Ap_balance_6 ")
        Dim B_Curr As String = ""
        If CMB_Curr.SelectedIndex = 0 Then
            B_Curr = ""
        Else
            B_Curr = " AND  Curr=N'" & CMB_Curr.Text & "' "
        End If


        'DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        '" select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")
        DbHelper.ExecuteNonQuery("UPDATE gen_jn set amt_USD_Dr=0 where amt_USD_Dr is null ")
        DbHelper.ExecuteNonQuery("UPDATE gen_jn set amt_USD_Cr=0 where amt_USD_Cr is null ")
        DbHelper.ExecuteNonQuery("UPDATE Open_jn set amt_USD_Dr=0 where amt_USD_Dr is null ")
        DbHelper.ExecuteNonQuery("UPDATE Open_jn set amt_USD_Cr=0 where amt_USD_Cr is null ")
        DbHelper.ExecuteNonQuery("UPDATE AP_ACC_adjust_Item set amt_USD_Dr=0 where amt_USD_Dr is null ")
        DbHelper.ExecuteNonQuery("UPDATE AP_ACC_adjust_Item set amt_USD_Cr=0 where amt_USD_Cr is null ")

        'CheckBox4.Text = "ທຽບເທົ່າໂດລາ"
        'CheckBox5.Text = "ທຽບເທົ່າບາດ"
        'Label23.Text = "ບາດ-ກີບ" 1
        'Label24.Text = "ໂດລາ-ກີບ" 2

        If CMB_Curr.SelectedIndex = 0 Then
            If CheckBox4.Checked = True Then
                '=======LAK===
                Dim LAK As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
               " select ac_code, Curr, 0, 0 , sum(amt_USD_Dr) , sum(amt_USD_Cr)  from gen_jn  WHERE 1=1 " & B_Curr & " and date_work BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code, Curr "
                DbHelper.ExecuteNonQuery(LAK)
                Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
                '=======LAK===
                Dim PPP As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code , Curr, sum(amt_USD_Dr)as amt_dr , sum(amt_USD_Cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE 1=1 " & B_Curr & " and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code, Curr"
                DbHelper.ExecuteNonQuery(PPP)
                '=======LAK===
                DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code, Curr, open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code , Curr , sum(amt_USD_Dr) as amt_dr , sum(amt_USD_Cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE  1=1 " & B_Curr & " and   date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code, Curr")

                If CheckBox3.Checked = True Then
                    Dim aa As String
                    aa = "  insert into  Ap_balance_6 (ac_code, Curr,amt_dr,amt_cr)   SELECT  AP_ACC_adjust_Item.ac_code,SUM (AP_ACC_adjust_Item.amt_USD_Dr),SUM (AP_ACC_adjust_Item.amt_USD_Cr)  from  AP_ACC_adjust_Item " & _
                            "   where  year(AP_ACC_adjust_Item.date_work) = '" & Format(CDate(MdStartDate), "yyyy") & "' " & MULook2 & "  " & _
                            " group by  ac_code, Curr,com_id   "
                    DbHelper.ExecuteNonQuery(aa)

                End If

            ElseIf CheckBox5.Checked = True Then

            Else
                '=======LAK===
                Dim GGG As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
               " select ac_code , Curr ,  0 As open_amt_dr, 0 As open_amt_cr , sum(Amount_Dr)as amt_dr , sum(Amount_cr)as amt_cr  from gen_jn   " & _
               " WHERE 1=1  and Curr=N'LAK'   and date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code, Curr "
                DbHelper.ExecuteNonQuery(GGG)
                '=======USD===
                Dim USD As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
               " select ac_code , Curr ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amount_Dr)* " & CDbl(txtRate.Text) & "  as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & " as amt_cr  from gen_jn  " & _
               " WHERE 1=1 and Curr=N'USD'  and date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code, Curr  "
                DbHelper.ExecuteNonQuery(USD)
                '=======THB===
                Dim THB As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
               " select ac_code , Curr ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amount_Dr)* " & CDbl(txtRate.Text) & "  as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & " as amt_cr  from gen_jn  " & _
               " WHERE 1=1 and Curr=N'THB'  and date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code, Curr  "
                DbHelper.ExecuteNonQuery(THB)



                Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
                '=======LAK===
                Dim PPP As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                        " select ac_code, Curr , sum(Amount_Dr)as amt_dr , sum(Amount_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn " & _
                        " WHERE 1=1  and Curr=N'LAK' and date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code, Curr "
                DbHelper.ExecuteNonQuery(PPP)
                Dim PPPUSD As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code, Curr , sum(amount_Dr)* " & CDbl(txtRate.Text) & " as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & " as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn " & _
        " WHERE 1=1  and Curr=N'USD'  and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code, Curr "
                DbHelper.ExecuteNonQuery(PPPUSD)
                '        '=======LAK===
                DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code, Curr  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code, Curr  , sum(Amount_Dr) as amt_dr , sum(Amount_cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1   and Curr=N'LAK'  and date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code, Curr ")
                DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
           " select ac_code  , sum(amount_Dr)* " & CDbl(txtRate.Text) & "  as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & "  as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE  1=1  and Curr=N'USD'  and   date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code, Curr ")

                If CheckBox3.Checked = True Then
                    Dim aa As String
                    aa = "  insert into  Ap_balance_6 (ac_code, Curr,amt_dr,amt_cr)   SELECT  AP_ACC_adjust_Item.ac_code, AP_ACC_adjust_Item.Curr,SUM (AP_ACC_adjust_Item.amt_dr),SUM (AP_ACC_adjust_Item.amt_cr)  from  AP_ACC_adjust_Item " & _
                            "   where  year(AP_ACC_adjust_Item.date_work) = '" & Format(CDate(MdStartDate), "yyyy") & "' " & MULook2 & "  " & _
                            " group by  ac_code, Curr ,com_id   "
                    DbHelper.ExecuteNonQuery(aa)


                End If
            End If


        Else

            Dim GGG As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            " select ac_code, Curr, 0, 0, sum(Amount_Dr), sum(Amount_cr) from gen_jn  WHERE 1=1 " & B_Curr & " and date_work BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code, Curr  "
            DbHelper.ExecuteNonQuery(GGG)

            Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
            Dim PPP As String = "INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            " select ac_code, Curr, sum(Amount_Dr), sum(Amount_cr), 0, 0 from gen_jn  WHERE 1=1 " & B_Curr & " and date_work BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code, Curr "
            DbHelper.ExecuteNonQuery(PPP)
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            " select ac_code, Curr, sum(Amount_Dr), sum(Amount_cr), 0, 0 from Open_jn WHERE  1=1 " & B_Curr & " and   date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code, Curr ")
            If CheckBox3.Checked = True Then
                Dim aa As String
                aa = "  insert into  Ap_balance_6 (ac_code, Curr,open_amt_dr,open_amt_cr,amt_dr,amt_cr)   SELECT  AP_ACC_adjust_Item.ac_code, AP_ACC_adjust_Item.Curr,0,0,SUM (AP_ACC_adjust_Item.Amount_Dr),SUM (AP_ACC_adjust_Item.Amount_cr)  from  AP_ACC_adjust_Item " & _
                        "   where  year(AP_ACC_adjust_Item.date_work) = '" & Format(CDate(MdStartDate), "yyyy") & "' " & MULook2 & "  " & _
                        " group by  ac_code, Curr ,com_id   "
                DbHelper.ExecuteNonQuery(aa)
            End If
        End If
        'End If

        DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6_col ( ac_code , Curr , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code , Curr , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr)as open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code, Curr ")

        Call Left_AcCode()
        DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        Call Chang_Incom()
        DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + amt_dr) - (open_amt_cr + amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr + amt_cr) ")
        DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr + amt_cr) - (open_amt_dr + amt_dr) where (open_amt_cr + amt_cr) >= (open_amt_dr + amt_dr) ")
        If MuLng = "L" Then
            DbHelper.ExecuteNonQuery("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        Else
            DbHelper.ExecuteNonQuery("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_E from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        End If



    End Sub
    Private Sub LoadReport()

        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As RptSjUd ,"
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
            If RD.Checked = True Then
                ' LngId = "7045" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            ElseIf RM.Checked = True Then
                'LngId = "7045" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & " " & Lb.Text & "' As Crl_RptName ,"
            ElseIf RP.Checked = True Then
                ' LngId = "7046" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & " " & Lb.Text & "' As Crl_RptName ,"
            ElseIf RT.Checked = True Then
                LngId = "7078" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
                LngId = "7079" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
                ' LngId = "7045" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            ElseIf RY.Checked = True Then
                ' LngId = "7047" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & " " & Lb.Text & "' As Crl_RptName ,"

            End If
            LngId = "7045" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & " " & DTDATE02 & "' As Crl_RptName ,"
        End If

        Call LoadLoGO()
        Dim MdEnd As String
        MdEnd = "For the Month Ended " & Format(CDate(MdToDate), "dd/MM/yyyy")
        MdEnd = Today
        'DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set H1=left(H1,7) ")
        If RDtail.Checked = True Or RGroup.Checked = True Then
            'DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set H3=substring(ac_code,1,6) ")
            'DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set H4=substring(ac_code,1,7) ")
            'DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set H5=substring(ac_code,1,8) ")
            'DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set H6=substring(ac_code,1,9) ")
            ''DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set H7=substring(ac_code,1,10) ")

            DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set H3=Left(ac_code,3) where len(Left(ac_code,3))=3 ")
            DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set H4=Left(ac_code,4) where len(Left(ac_code,4))=4 ")
            DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set H5=Left(ac_code,5) where len(Left(ac_code,5))=5 ")
            DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set H6=Left(ac_code,6) where len(Left(ac_code,6))=6 ")
            DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set H7=Left(ac_code,7) where len(Left(ac_code,7))=7 ")

            DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set H4_Nm='' where H4_Nm is null ")
            DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set H5_Nm='' where H5_Nm is null ")
            DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set H6_Nm='' where H6_Nm is null ")
            DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set H7_Nm='' where H7_Nm is null ")

        End If
        If MuLng = "L" Then
            DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set  H3_Nm=Acc_Code.Name_L from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.H3=Acc_Code.ac_code ")
            DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set   H4_Nm=Acc_Code.Name_L from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.H4=Acc_Code.ac_code ")
            DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set   H5_Nm=Acc_Code.Name_L from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.H5=Acc_Code.ac_code ")
            DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set   H6_Nm=Acc_Code.Name_L from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.H6=Acc_Code.ac_code ")
            DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set   H7_Nm=Acc_Code.Name_L from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.H7=Acc_Code.ac_code ")
        Else
            DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set  H3_Nm=Acc_Code.Name_E from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.H3=Acc_Code.ac_code ")
            DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set   H4_Nm=Acc_Code.Name_E from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.H4=Acc_Code.ac_code ")
            DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set   H5_Nm=Acc_Code.Name_E from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.H5=Acc_Code.ac_code ")
            DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set   H6_Nm=Acc_Code.Name_E from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.H6=Acc_Code.ac_code ")
            DbHelper.ExecuteNonQuery(" update Ap_balance_6_col set   H7_Nm=Acc_Code.Name_E from Acc_Code,Ap_balance_6_col where Ap_balance_6_col.H7=Acc_Code.ac_code ")
        End If
        Dim ss As String

        If CMB_Curr.SelectedIndex = 0 Then
            If RDtail.Checked = True Or RGroup.Checked = True Then
                ss = " update Ap_balance_6_col set  AC_CODE2='' where Curr='LAK' "
                DbHelper.ExecuteNonQuery(ss)
                ss = " update Ap_balance_6_col set  AC_CODE2='' where Curr='USD' "
                DbHelper.ExecuteNonQuery(ss)
                ss = "  update Ap_balance_6_col set  AC_CODE2='' where Curr='THB'  "
                DbHelper.ExecuteNonQuery(ss)

            Else

                ss = " update Ap_balance_6_col set  AC_CODE2=AC_CODE where Curr='LAK' "
                DbHelper.ExecuteNonQuery(ss)
                ss = " update Ap_balance_6_col set  AC_CODE2=AC_CODE where Curr='USD' "
                DbHelper.ExecuteNonQuery(ss)
                ss = "  update Ap_balance_6_col set  AC_CODE2=AC_CODE where Curr='THB'  "
                DbHelper.ExecuteNonQuery(ss)

            End If
        Else
            If RDtail.Checked = True Or RGroup.Checked = True Then
                ss = " update Ap_balance_6_col set  AC_CODE2='00.' where Curr='LAK' "
                DbHelper.ExecuteNonQuery(ss)
                ss = " update Ap_balance_6_col set  AC_CODE2='01.' where Curr='USD' "
                DbHelper.ExecuteNonQuery(ss)
                ss = "  update Ap_balance_6_col set  AC_CODE2='02.' where Curr='THB'  "
                DbHelper.ExecuteNonQuery(ss)

            Else

                ss = " update Ap_balance_6_col set  AC_CODE2='00.' + AC_CODE where Curr='LAK' "
                DbHelper.ExecuteNonQuery(ss)
                ss = " update Ap_balance_6_col set  AC_CODE2='01.' + AC_CODE where Curr='USD' "
                DbHelper.ExecuteNonQuery(ss)
                ss = "  update Ap_balance_6_col set  AC_CODE2='02.' + AC_CODE where Curr='THB'  "
                DbHelper.ExecuteNonQuery(ss)

            End If
        End If
       

        If CMB_Curr.Text = "LAK" Then 
            CURR = "ຫົວໜ່ວຍ : ກີບ"
        ElseIf CMB_Curr.Text = "USD" Then
            CURR = "ຫົວໜ່ວຍ : ໂດລາ"
        Else
            CURR = "ຫົວໜ່ວຍ : ກີບ"
        End If


        SLF = " SELECT " & MuLngRpt & " * FROM Ap_balance_6_Col Order by ac_code asc "
        Dim dt As DataTable = DbHelper.GetDataTable(SLF)
        If dt.Rows.Count = 0 Then
            MsgBox("ບໍ່ມີຂໍ້ມູນ")
            Exit Sub
        End If
        Dim FrmPreview As New FmPreview : FrmClosing()
        'Dim Rpt As New CryTrialBalanceReport_6_Incom2022
        Dim Rpt As New Object
      
        If BalanceType.SelectedIndex = 1 Then
            If RDtail.Checked = True Then
                Rpt = New CryTrialBalanceReport_6_Incom2022_3_5
            ElseIf RGroup.Checked = True Then
                Rpt = New CryTrialBalanceReport_6_Incom2022_3_5_G
            Else
                Rpt = New CryTrialBalanceReport_6_Incom2022
            End If
        Else
            If RDtail.Checked = True Then
                Rpt = New CryTrialBalanceReport_6_Incom2022_3_5_BType
            ElseIf RGroup.Checked = True Then
                Rpt = New CryTrialBalanceReport_6_Incom2022_3_5_G_BType
            Else
                Rpt = New CryTrialBalanceReport_6_Incom2022_BType
            End If
        End If
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
        'myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
        'myText2.Text = MuOffNEW
        'myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("End"), CrystalDecisions.CrystalReports.Engine.TextObject)
        'myText2.Text = MdEnd
        'myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("OfTel"), CrystalDecisions.CrystalReports.Engine.TextObject)
        'myText2.Text = MDRegister
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("Text4"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = DTDATE
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("txtprint_user"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = MUserName
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("Text1"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = CURR
        Rpt.SetDataSource(dt)
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.DisplayGroupTree = False
        FrmPreview.MdiParent = FmMain
        FrmPreview.WindowState = FormWindowState.Maximized
        FrmPreview.Show()
        FrmPreview.Focus()
        'FrmPreview.Focus()

    End Sub


    Private Sub LoadOpen_Jn16()
        Dim dt As DataTable = DbHelper.GetDataTable("  select sum(rem_cr - rem_dr) as x  , count(rem_cr - rem_dr) as y  from Ap_balance_6_col  where Ac_Code Like '1%' or Ac_Code Like '2%' or Ac_Code Like '3%' ")

        If dt.Rows.Count > 0 AndAlso CDbl(Trim(dt.Rows(0)("y").ToString())) > 0 Then

            If CDbl(Trim(dt.Rows(0)("x").ToString())) > 0 Then
                Rem_Cr = CDbl(Trim(dt.Rows(0)("x").ToString()))
                Rem_Dr = 0
            End If
            If CDbl(Trim(dt.Rows(0)("x").ToString())) < 0 Then
                Rem_Dr = CDbl(Trim(dt.Rows(0)("x").ToString())) * CDbl(-1)
                Rem_Cr = 0
            End If
            If CDbl(Trim(dt.Rows(0)("x").ToString())) <> 0 Then
                Call LoadOpen_Jn17()
            End If
        End If
    End Sub
    Private Sub LoadOpen_Jn17()
        Dim dt As DataTable = DbHelper.GetDataTable("   select Ac_Code from Ap_balance_6_col  where Ac_Code ='65'")
        If dt.Rows.Count > 0 Then
            DbHelper.ExecuteNonQuery(" Update Ap_balance_6_col set Amt_Dr = " & CDbl(Rem_Dr) & "  , Amt_Cr =" & CDbl(Rem_Cr) & " ")
        Else
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6_col ( ac_code ,ac_name , ac_namee , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
             "Values('65', N'" & "***" & "', '" & "***" & "', " & _
             " " & CDbl(0) & ", " & CDbl(0) & ", " & CDbl(Rem_Dr) & ", " & CDbl(Rem_Cr) & ",0 )")
        End If

        DbHelper.ExecuteNonQuery(" Delete Ap_balance_6_col  where Ac_Code Like '1%' or Ac_Code Like '2%' or Ac_Code Like '3%'  ")

    End Sub
    Private Sub LoadOpen_Jn1()
        Dim dt As DataTable = DbHelper.GetDataTable("   select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook & "  group BY ac_code ")
        For Each row As DataRow In dt.Rows
            VCode1 = CStr(Trim(row("ac_Code").ToString()))
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6_col( ac_code   , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  )  select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr  , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As Status from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook & "  group BY ac_code ")
        Next
    End Sub

    Private Sub LoadOpen_Jn2()
        Dim S As Date = MdStartDate
        S = DateAdd("d", CDbl(-1), MdStartDate)
        DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
       " select ac_code , 0 As open_amt_dr , 0 open_amt_cr ,sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 Status from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook & "   group BY ac_code")

    End Sub

    Private Sub LoadOpen_Jn3()
        Dim dt As DataTable = DbHelper.GetDataTable("select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr from Open_jn WHERE    date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook & "   group BY ac_code")
        For Each row As DataRow In dt.Rows
            VCode3 = row("ac_Code").ToString()
            DbHelper.ExecuteNonQuery("Update Ap_balance_6 set  open_amt_dr='" & CDbl((row("amt_dr").ToString())) & "' , open_amt_cr='" & CDbl((row("amt_cr").ToString())) & "' where ac_code = '" & row("ac_Code").ToString() & "'")
            LoadOpen_Jn4()
        Next
    End Sub


    Private Sub LoadOpen_Jn4()
        Dim dt As DataTable = DbHelper.GetDataTable("select ac_Code , amt_dr , amt_cr  from Ap_balance_6  WHERE     ac_code='" & VCode3 & "'  ")
        If dt.Rows.Count > 0 Then
            VCode4 = dt.Rows(0)("ac_Code").ToString()
        Else
            LoadOpen_Jn5()
        End If
    End Sub

    Private Sub LoadOpen_Jn5()
        Dim dt As DataTable = DbHelper.GetDataTable("select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr from Open_jn  WHERE    ac_code='" & VCode3 & "' " & MULook & " group BY ac_code")
        If dt.Rows.Count > 0 Then
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
         "Values('" & CStr(Trim(dt.Rows(0)("ac_Code").ToString())) & "',  " & _
         " " & CDbl(dt.Rows(0)("amt_dr").ToString()) & ", " & CDbl(dt.Rows(0)("amt_cr").ToString()) & ", " & CDbl(0) & ", " & CDbl(0) & ",0 )")
        End If
    End Sub



    Private Sub LoadOpen_Jn6()
        Dim dt As DataTable = DbHelper.GetDataTable("select Ac_Code , open_amt_dr , open_amt_cr , Amt_dr , Amt_cr from Ap_balance_6  ")
        For Each row As DataRow In dt.Rows
            Dim op_dr, op_cr, amt_dr, amt_cr As Double
            op_dr = CDbl((row("open_amt_dr").ToString()))
            op_cr = CDbl((row("open_amt_cr").ToString()))
            amt_dr = CDbl((row("Amt_dr").ToString()))
            amt_cr = CDbl((row("Amt_cr").ToString()))
            If CDbl(op_dr + amt_dr) - CDbl(op_cr + amt_cr) >= 0 Then
                DbHelper.ExecuteNonQuery("Update Ap_balance_6 set rem_dr='" & CDbl(op_dr + amt_dr) - CDbl(op_cr + amt_cr) & "' , rem_cr='" & CDbl(0) & "' where Ac_code='" & row("Ac_Code").ToString() & "'")
            End If
            If CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) >= 0 Then
                DbHelper.ExecuteNonQuery("Update Ap_balance_6 set rem_dr='" & CDbl(0) & "' , rem_cr='" & CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) & "' where Ac_code='" & row("Ac_Code").ToString() & "'")
            End If
        Next
    End Sub

    Private Sub LoadOpen_Jn7()
        Dim dt As DataTable = DbHelper.GetDataTable("select ac_code , rem_dr  , rem_cr from Ap_balance_6   ")
        For Each row As DataRow In dt.Rows
            VCode7 = row("ac_Code").ToString()
            DbHelper.ExecuteNonQuery("Update Ap_balance_6_col set  open_amt_dr='" & CDbl((row("rem_dr").ToString())) & "' , open_amt_cr='" & CDbl((row("rem_cr").ToString())) & "' where ac_code = '" & row("ac_code").ToString() & "'")
            LoadOpen_Jn8()
        Next
    End Sub


    Private Sub LoadOpen_Jn8()
        Dim dt As DataTable = DbHelper.GetDataTable("select ac_code   from Ap_balance_6_col  WHERE     ac_code='" & VCode7 & "' ")
        If dt.Rows.Count > 0 Then
            VCode8 = dt.Rows(0)("ac_Code").ToString()
        Else
            LoadOpen_Jn9()
        End If
    End Sub


    Private Sub LoadOpen_Jn9()
        Dim dt As DataTable = DbHelper.GetDataTable("select ac_code , Rem_dr , Rem_cr from Ap_balance_6  WHERE    ac_code='" & VCode7 & "' ")
        If dt.Rows.Count > 0 Then
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6_col ( ac_code ,ac_name , ac_namee , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
            "Values('" & CStr(Trim(dt.Rows(0)("ac_Code").ToString())) & "', N'" & "***" & "', '" & "***" & "', " & _
            " " & CDbl(dt.Rows(0)("rem_dr").ToString()) & ", " & CDbl(dt.Rows(0)("rem_cr").ToString()) & ", " & CDbl(0) & ", " & CDbl(0) & ",0 )")
        End If
    End Sub
    Private Sub LoadOpen_Jn14_1()
        Dim dt As DataTable = DbHelper.GetDataTable("select Ac_Code , open_amt_dr ,open_amt_cr , amt_dr , amt_cr  from Ap_balance_6_col  ")
        For Each row As DataRow In dt.Rows
            If CDbl(CDbl((row("open_amt_dr").ToString())) + CDbl((row("amt_dr").ToString()))) >= CDbl(CDbl((row("open_amt_cr").ToString())) + CDbl((row("amt_cr").ToString()))) Then
                DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr=" & CDbl(CDbl(CDbl((row("open_amt_dr").ToString())) + CDbl((row("amt_dr").ToString()))) - CDbl(CDbl((row("open_amt_cr").ToString())) + CDbl((row("amt_cr").ToString())))) & " where Ac_Code = '" & row("Ac_Code").ToString() & "'")
            Else
                DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr=" & CDbl(CDbl(CDbl((row("open_amt_cr").ToString())) + CDbl((row("amt_cr").ToString()))) - CDbl(CDbl((row("open_amt_dr").ToString())) + CDbl((row("amt_dr").ToString())))) & " where Ac_Code = '" & row("Ac_Code").ToString() & "'")
            End If
        Next
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

            Dim dt As DataTable = DbHelper.GetDataTable("SELECT  Ac_Code  FROM   Ap_balance_6_col  ")
            For Each row As DataRow In dt.Rows
                d = row("Ac_Code").ToString()
                Call Pr()
                DbHelper.ExecuteNonQuery("Update Ap_balance_6_col set  Acc_Parent = '" & p & "'  where ac_code='" & row("ac_code").ToString() & "'")
            Next

            LoadPl2()

        End If

    End Sub



    Private Sub LoadPl2()

        DbHelper.ExecuteNonQuery("Delete Ap_balance_6_ChangParent")
        DbHelper.ExecuteNonQuery("insert Into Ap_balance_6_ChangParent (Ac_Code   ,  open_amt_dr ,  open_amt_cr   , amt_dr , amt_cr      ) select Acc_Parent    , sum(open_amt_dr) as open_amt_dr  , sum(open_amt_cr) as open_amt_cr   ,  sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr       from Ap_balance_6_col group by Acc_Parent ")

        Dim dtParent As DataTable = DbHelper.GetDataTable("select  Ac_Code , open_amt_dr , open_amt_cr from Ap_balance_6_ChangParent   ")

        For Each row As DataRow In dtParent.Rows
            If CDbl((row("open_amt_dr").ToString())) >= CDbl((row("open_amt_cr").ToString())) Then
                ''MsgBox(CDbl((row("open_amt_dr").ToString())) & "---" & CDbl((row("open_amt_cr").ToString())) & "==" & "9999999999" & "==--" & (row("Ac_Code").ToString()))
                DbHelper.ExecuteNonQuery("Update Ap_balance_6_ChangParent set open_amt_dr = '" & CDbl((row("open_amt_dr").ToString())) - CDbl((row("open_amt_cr").ToString())) & "' , open_amt_cr=0  where ac_code='" & row("ac_code").ToString() & "'")

            ElseIf CDbl((row("open_amt_dr").ToString())) <= CDbl((row("open_amt_cr").ToString())) Then
                ''MsgBox(CDbl((row("open_amt_dr").ToString())) & "---" & CDbl((row("open_amt_cr").ToString())) & "==" & CDbl((row("open_amt_cr").ToString())) & "==--" & (row("Ac_Code").ToString()))
                DbHelper.ExecuteNonQuery("Update Ap_balance_6_ChangParent set open_amt_dr=0 , open_amt_cr = '" & CDbl((row("open_amt_cr").ToString())) - CDbl((row("open_amt_dr").ToString())) & "'   where ac_code='" & row("ac_code").ToString() & "'")

            End If
        Next

        Dim dtChange As DataTable = DbHelper.GetDataTable("select  Ac_Code , open_amt_dr , open_amt_cr  , amt_dr , amt_cr from Ap_balance_6_ChangParent   ")
        For Each row As DataRow In dtChange.Rows
            If CDbl(CDbl((row("open_amt_dr").ToString())) + CDbl((row("amt_dr").ToString()))) >= CDbl(CDbl((row("open_amt_cr").ToString())) + CDbl((row("amt_cr").ToString()))) Then
                DbHelper.ExecuteNonQuery("Update  Ap_balance_6_ChangParent set Rem_cr=0 , Rem_dr=" & CDbl(CDbl(CDbl((row("open_amt_dr").ToString())) + CDbl((row("amt_dr").ToString()))) - CDbl(CDbl((row("open_amt_cr").ToString())) + CDbl((row("amt_cr").ToString())))) & " where Ac_Code = '" & row("Ac_Code").ToString() & "'")
            Else
                DbHelper.ExecuteNonQuery("Update  Ap_balance_6_ChangParent set Rem_dr=0 , Rem_cr=" & CDbl(CDbl(CDbl((row("open_amt_cr").ToString())) + CDbl((row("amt_cr").ToString()))) - CDbl(CDbl((row("open_amt_dr").ToString())) + CDbl((row("amt_dr").ToString())))) & " where Ac_Code = '" & row("Ac_Code").ToString() & "'")
            End If
        Next


        Dim dtCodes As DataTable = DbHelper.GetDataTable("SELECT  Acc_Code.Ac_Code AS Ac_Code, Acc_Code.Name_L AS Name_L FROM   Acc_Code INNER JOIN    Ap_balance_6_ChangParent ON Acc_Code.Ac_Code = Ap_balance_6_ChangParent.ac_code  ")
        For Each row As DataRow In dtCodes.Rows
            DbHelper.ExecuteNonQuery("Update Ap_balance_6_ChangParent set ac_name = N'" & row("Name_L").ToString() & "'  where ac_code='" & row("ac_code").ToString() & "'")
        Next

        'Call LoadRaParent2()

    End Sub


    Private Sub LoadPr1()
        'This method appears to be unused and referenced RSCP which is no longer available
        'Dim dt As DataTable = DbHelper.GetDataTable("select * from Ap_balance_6_col where Ac_Code <> '" & (RSCP.Fields("Acc_Parent").Value) & "' ")
        'If dt.Rows.Count > 0 Then
        '    For Each row As DataRow In dt.Rows
        '        'MsgBox((row("Acc_Parent").ToString()))
        '        'DbHelper.ExecuteNonQuery("Update Ap_balance_6_col set ac_name = N'" & (row("Name_L").ToString()) & "'   where ac_code='" & row("ac_code").ToString() & "'")
        '    Next
        'End If
    End Sub

    Private Sub LoadOpen_Jn11()
        Dim dt As DataTable = DbHelper.GetDataTable("select Ac_Code , open_amt_dr , open_amt_cr , Amt_dr , Amt_cr from Ap_balance_6_col  ")
        For Each row As DataRow In dt.Rows
            Dim op_dr11, op_cr11, amt_dr11, amt_cr11 As Double
            op_dr11 = CDbl((row("open_amt_dr").ToString()))
            op_cr11 = CDbl((row("open_amt_cr").ToString()))
            amt_dr11 = CDbl((row("Amt_dr").ToString()))
            amt_cr11 = CDbl((row("Amt_cr").ToString()))
            If CDbl(op_dr11 + op_cr11) = 0 Then
                If CDbl(amt_dr11 + amt_cr11) = 0 Then
                    DbHelper.ExecuteNonQuery("delete Ap_balance_6_col  where Ac_code='" & row("Ac_Code").ToString() & "'")
                End If
            End If
        Next
    End Sub


    Private Sub BtnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRefresh.Click
        If BalanceType.SelectedIndex = 0 Then
            Ac_Code = ""
        ElseIf BalanceType.SelectedIndex = 1 Then
            'Ac_Code = "And (Left(Ac_Code,1) = '4' or Left(Ac_Code,1) = '5') "
            'Ac_Code = "And (Left(Ac_Code,4) = '00.4' or Left(Ac_Code,4) = '00.5') "
            Ac_Code = "And (Left(Ac_Code,1) = '4' or Left(Ac_Code,1) = '5') "
            ChangInCom = BalanceType.SelectedIndex

            'MsgBox("dd")

        End If
        If CheckBox1.Checked = True Then
            ChangInCom = 1
        Else
            ChangInCom = 0
        End If


        New_Code = "3901000"
        ' New_Code = "3901"
        Code_Dr = "4"
        Code_Cr = "5"
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        Call ChangBalance()
        'Call BLNEW()
        If CheckBnk.Checked = True Then
            DbHelper.ExecuteNonQuery("delete Ap_balance_6_col from Ap_balance_6_col , Ap_Rpt_BLS_Item where Ap_balance_6_col.ac_code=Ap_Rpt_BLS_Item.Ac_Code")
            DbHelper.ExecuteNonQuery("delete Ap_balance_6_col from Ap_balance_6_col , Ap_Rpt_Income_Item where Ap_balance_6_col.ac_code=Ap_Rpt_Income_Item.Ac_Code")
        End If
        If CheckBox2.Checked = True Then
            DbHelper.ExecuteNonQuery("delete Ap_balance_6_col where Amt_Dr=0 And Amt_Cr=0 ")
        End If
        If TextBox1.Text <> "" Then
            DbHelper.ExecuteNonQuery("delete Ap_balance_6_col where Left(Ac_Code, " & Len(TextBox1.Text) & ") <> '" & TextBox1.Text & "' ")
        End If
        If MuLng = "L" Then
            DbHelper.ExecuteNonQuery("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        Else
            DbHelper.ExecuteNonQuery("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_E from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        End If

        Call LoadListFG()
        Call SumData()
    End Sub
    'Private Sub BLNEW()
    '    '=================NEWWWWW==============
    '    DbHelper.ExecuteNonQuery("DELETE  Ap_balance_6_col ")
    '    DbHelper.ExecuteNonQuery("DELETE FROM Ap_balance_6 ")
    '    If RM.Checked = True Then
    '        DbHelper.ExecuteNonQuery(" insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  from Ap_balance_TB where  date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  " & Ac_Code & "  order by Ac_Code asc ")
    '    ElseIf RP.Checked = True Then

    '        DbHelper.ExecuteNonQuery(" insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , 0 , 0 , amt_dr , amt_cr  from Ap_balance_TB where  date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  " & Ac_Code & "   order by Ac_Code asc ")
    '        Dim OP As String = " insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , 0 , 0  from Ap_balance_TB where  month(date_work)= '" & Month(MdStartDate) & "' AND   year(date_work)= '" & Year(MdStartDate) & "'  " & Ac_Code & "    order by Ac_Code asc "
    '        DbHelper.ExecuteNonQuery(OP)
    '        Dim HHH As String = "INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
    '   " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr)as open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code"
    '        DbHelper.ExecuteNonQuery(HHH)
    '    End If


    '    'Call Left_AcCode()
    '    'DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
    '    'DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
    '    'Call Chang_Incom()
    '    DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + amt_dr) - (open_amt_cr + amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr + amt_cr) ")
    '    DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr + amt_cr) - (open_amt_dr + amt_dr) where (open_amt_cr + amt_cr) >= (open_amt_dr + amt_dr) ")
    '    If MuLng = "L" Then
    '        DbHelper.ExecuteNonQuery("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
    '    Else
    '        DbHelper.ExecuteNonQuery("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_E from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
    '    End If

    'End Sub

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
            DbHelper.ExecuteNonQuery(Insr)
        End If
    End Sub
    Private Sub Chang_Incom()
        If ChangInCom = 1 Then
            Insr = "delete  Ap_balance_6  " & _
               "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  )   select '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr) from Ap_balance_6_col where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "'  or  Ac_Code =  '" & New_Code & "'  " & _
      "update  Ap_balance_6 set open_amt_dr = 0 where open_amt_dr  is null  " & _
      "update  Ap_balance_6 set open_amt_cr = 0 where open_amt_cr  is null  " & _
      "update  Ap_balance_6 set amt_dr = 0 where amt_dr  is null  " & _
      "update  Ap_balance_6 set amt_cr = 0 where amt_cr  is null   " & _
       "Update  Ap_balance_6 set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr   " & _
      "Update  Ap_balance_6 set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr " & _
      "Update  Ap_balance_6 set   amt_dr = amt_dr  - amt_cr , amt_cr=0  where amt_dr  >= amt_cr   " & _
      "Update  Ap_balance_6 set   amt_cr = amt_cr  - amt_dr , amt_dr=0  where amt_cr  >= amt_dr " & _
         "delete  Ap_balance_6_col  where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "'  or  Ac_Code =  '" & New_Code & "'  " & _
           "  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  from Ap_balance_6"
            DbHelper.ExecuteNonQuery(Insr)
        End If
    End Sub
    Private Sub LoadListFG()
        Dim O_dr, O_cr, Amt_dr, Amt_cr, R_dr, R_Cr As Double
        FG.Rows.Clear()
        
        Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM  Ap_balance_6_col Order by ac_code")

        If dt.Rows.Count > 0 Then
            Dim rowIndex As Integer = 0
            For Each row As DataRow In dt.Rows
                O_dr = Trim(CDbl(row("open_amt_dr").ToString()))
                O_cr = Trim(CDbl(row("open_amt_cr").ToString()))
                Amt_dr = Trim(CDbl(row("amt_dr").ToString()))
                Amt_cr = Trim(CDbl(row("amt_cr").ToString()))
                R_dr = Trim(CDbl(row("Rem_dr").ToString()))
                R_Cr = Trim(CDbl(row("Rem_cr").ToString()))

                FG.Rows.Add()
                FG.Rows(rowIndex).Cells("Col1").Value = rowIndex + 1
                FG.Rows(rowIndex).Cells("Col2").Value = Trim(CStr(row("ac_code").ToString()))
                FG.Rows(rowIndex).Cells("Col3").Value = Trim(CStr(row("ac_name").ToString()))
                FG.Rows(rowIndex).Cells("Col4").Value = Format(O_dr, "##,##0.00")
                FG.Rows(rowIndex).Cells("Col5").Value = Format(O_cr, "##,##0.00")
                FG.Rows(rowIndex).Cells("Col6").Value = Format(Amt_dr, "##,##0.00")
                FG.Rows(rowIndex).Cells("Col7").Value = Format(Amt_cr, "##,##0.00")
                FG.Rows(rowIndex).Cells("Col8").Value = Format(R_dr, "##,##0.00")
                FG.Rows(rowIndex).Cells("Col9").Value = Format(R_Cr, "##,##0.00")
                rowIndex += 1
            Next
        End If
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
        
        For i = 0 To FG.Rows.Count - 1
            If Not IsDBNull(FG.Rows(i).Cells("Col4").Value) Then
                OpDr.Text = CDbl(OpDr.Text) + CDbl(FG.Rows(i).Cells("Col4").Value)
            End If
            If Not IsDBNull(FG.Rows(i).Cells("Col5").Value) Then
                OpCr.Text = CDbl(OpCr.Text) + CDbl(FG.Rows(i).Cells("Col5").Value)
            End If
            If Not IsDBNull(FG.Rows(i).Cells("Col6").Value) Then
                AmtDr.Text = CDbl(AmtDr.Text) + CDbl(FG.Rows(i).Cells("Col6").Value)
            End If
            If Not IsDBNull(FG.Rows(i).Cells("Col7").Value) Then
                AmtCr.Text = CDbl(AmtCr.Text) + CDbl(FG.Rows(i).Cells("Col7").Value)
            End If
            If Not IsDBNull(FG.Rows(i).Cells("Col8").Value) Then
                ReDr.Text = CDbl(ReDr.Text) + CDbl(FG.Rows(i).Cells("Col8").Value)
            End If
            If Not IsDBNull(FG.Rows(i).Cells("Col9").Value) Then
                ReCr.Text = CDbl(ReCr.Text) + CDbl(FG.Rows(i).Cells("Col9").Value)
            End If
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

    Private Sub FG_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelectionChanged

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

    Private Sub BalanceType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BalanceType.SelectedIndexChanged

    End Sub

    Private Sub CMB_Curr_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMB_Curr.SelectedIndexChanged

        Dim dt As DataTable = DbHelper.GetDataTable("Select * From Curr_For_Rate Where   Curr =N'" & Trim(CMB_Curr.Text) & "'")
        If dt.Rows.Count > 0 Then
            txtcurr_name2.Text = Trim(dt.Rows(0)("Curr_name").ToString())
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
        txtRate2.Text = Format(MD_Rate2, "#,##0.00")
        If CMB_Curr.Text = "LAK" Then
            CheckBox4.Text = "ທຽບເທົ່າໂດລາ"
            CheckBox5.Text = "ທຽບເທົ່າບາດ"
            Label23.Text = "ບາດ-ກີບ"
            Label24.Text = "ໂດລາ-ກີບ"
        ElseIf CMB_Curr.Text = "THB" Then
            CheckBox4.Text = "ທຽບເທົ່າກີບ"
            CheckBox5.Text = "ທຽບເທົ່າໂດລາ"
            'Label23.Text = "ກີບ-ບາດ"
            'Label24.Text = "ໂດລາ-ບາດ"
        ElseIf CMB_Curr.Text = "USD" Then
            CheckBox4.Text = "ທຽບເທົ່າກີບ"
            CheckBox5.Text = "ທຽບເທົ່າບາດ"
            'Label23.Text = "ກີບ-ໂດລາ"
            'Label24.Text = "ບາດ-ໂດລາ"
        Else
            CheckBox4.Text = "ທຽບເທົ່າໂດລາ"
            CheckBox5.Text = "ທຽບເທົ່າບາດ"
            Label23.Text = "ບາດ-ກີບ"
            Label24.Text = "ໂດລາ-ກີບ"
        End If

    End Sub
    Dim CURR As String

    Private Sub txtRate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRate.KeyPress
        If e.KeyChar = Chr(13) Then
            txtRate.Text = Format(CDbl(txtRate.Text), "##,##0.00")
        End If

    End Sub

    Private Sub txtRate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRate.TextChanged

    End Sub

    Private Sub txtRate2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRate2.KeyPress
        If e.KeyChar = Chr(13) Then
            txtRate2.Text = Format(CDbl(txtRate2.Text), "##,##0.00")
        End If
    End Sub

    Private Sub txtRate2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRate2.TextChanged

    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged

        CheckBox5.Checked = False
    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged

        CheckBox4.Checked = False
    End Sub
End Class
