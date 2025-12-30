Public Class FmPostedLedgers_From3
    Dim x_i As Integer
    Dim MonthLetter1, Cp As String
    Dim MdStartDate As Date
    'Dim Op_Dr_Ant As Double
    'Dim Op_Ant_Cr As Double
    Dim MdToDate As Date
    Dim s As Double
    Dim k As String
    Dim VCode1, VCode2, VCode3, VCode4, VCode5, VCode6, VCode7, VCode8, VCode9 As String
    Dim RSC21 As New ADODB.Recordset
    Dim ACCNO, AccNm, Curr As String
    Private Sub HeaDer()
        LoadSqlData("SELECT * FROM Header WHERE ID=N'SP01' ", RSC)
        If RSC.RecordCount <> 0 Then
            TxtHeader.Text = Trim(RSC.Fields("Nm").Value.ToString)
            TxtS1.Text = Trim(RSC.Fields("S1").Value.ToString)
            TxtS2.Text = Trim(RSC.Fields("S2").Value.ToString)
            TxtS3.Text = Trim(RSC.Fields("S3").Value.ToString)
            TxtS4.Text = Trim(RSC.Fields("S4").Value.ToString)
            TxtPP.Text = Trim(RSC.Fields("pp").Value.ToString)
        End If
    End Sub
    Private Sub AddHeader()
        LoadSqlData("SELECT * FROM Header WHERE ID=N'SP01' ", RSC)
        If RSC.RecordCount = 0 Then
            CNN.Execute("INSERT INTO Header(ID,Nm,S1,S2,S3,S4,PP) " & _
                        " values('SP01',N'" & TxtHeader.Text & "',N'" & TxtS1.Text & "',N'" & TxtS2.Text & "',N'" & TxtS3.Text & "',N'" & TxtS4.Text & "',N'" & TxtPP.Text & "') ")
        Else
            CNN.Execute("UPDATE Header set Nm=N'" & TxtHeader.Text & "',S1=N'" & TxtS1.Text & "',S2=N'" & TxtS2.Text & "',S3=N'" & TxtS3.Text & "',S4=N'" & TxtS4.Text & "',PP=N'" & TxtPP.Text & "' " & _
                        " where ID='SP01' ")
        End If
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

    Private Sub Ins()

        Dim RSC21 As New ADODB.Recordset
        Dim x As Double
        Call LoadSqlData("SELECT * FROM  Ap_PostedLedgers where ac_Code= '" & k & "'  order by certify asc ", RSC21)
        With RSC21
            Do Until .EOF = True
                x = CDbl(CDbl(s) + CDbl((RSC21.Fields("amt_dr").Value))) - CDbl((RSC21.Fields("amt_cr").Value))
                CNN.Execute("update Ap_PostedLedgers set remain ='" & CDbl(x) & "'   where cnt = '" & (RSC21.Fields("cnt").Value) & "' ")
                s = x
                .MoveNext()
            Loop
        End With

    End Sub



    Private Sub FmPostedLedgers_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Call MdiCNum()
    End Sub
    Private Sub FmPostedLedgers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        HeaDer()
        Call loadOffice_User()
        FG.AllowUserResizing = VSFlex8U.AllowUserResizeSettings.flexResizeBoth
        RD.Checked = True
        Ds.Text = MWorkSetting
        Myy.Text = MWorkSetting
        yy.Text = MWorkSetting

        Pyy.Text = MWorkSetting
        If Month(MWorkSetting) = 1 Then
            DMonth.SelectedIndex = 0
            Period.SelectedIndex = 0
        ElseIf Month(MWorkSetting) = 2 Then
            DMonth.SelectedIndex = 1
            Period.SelectedIndex = 0
        ElseIf Month(MWorkSetting) = 3 Then
            DMonth.SelectedIndex = 2
            Period.SelectedIndex = 0
        ElseIf Month(MWorkSetting) = 4 Then
            DMonth.SelectedIndex = 3
            Period.SelectedIndex = 1
        ElseIf Month(MWorkSetting) = 5 Then
            DMonth.SelectedIndex = 4
            Period.SelectedIndex = 1
        ElseIf Month(MWorkSetting) = 6 Then
            DMonth.SelectedIndex = 5
            Period.SelectedIndex = 1
        ElseIf Month(MWorkSetting) = 7 Then
            DMonth.SelectedIndex = 6
            Period.SelectedIndex = 2
        ElseIf Month(MWorkSetting) = 8 Then
            DMonth.SelectedIndex = 7
            Period.SelectedIndex = 2
        ElseIf Month(MWorkSetting) = 9 Then
            DMonth.SelectedIndex = 8
            Period.SelectedIndex = 2
        ElseIf Month(MWorkSetting) = 10 Then
            DMonth.SelectedIndex = 9
            Period.SelectedIndex = 3
        ElseIf Month(MWorkSetting) = 11 Then
            DMonth.SelectedIndex = 10
            Period.SelectedIndex = 3
        ElseIf Month(MWorkSetting) = 12 Then
            DMonth.SelectedIndex = 11
            Period.SelectedIndex = 3
        End If
        'Period.Text = "ງວດທີ 1"
        'DMonth.Text = "ມັງກອນ"
        'LoadMonth()
        selectLoad()

        FG.FormatString = "^ ລ/ດ |<  ເລກໃບຢັງຢືນ |ລະຫັດບັນຊີ |  ຍອດຍົກມູນຄ່າເດີມ  |ຍອດຍອກມູນຄ່າເປັນກີບ | ຈົດຫນີ້ມູນຄ່າເດີມ | ຈົດມີມູນຄ່າເດີມ  |  ດຫນີ້ມູນຄ່າເປັນກີບ |ຈົດມີມູນຄ່າເປັນກີບ | ຍອດເຫລືອ     |"

        SetControlText(Me)
        Button4.Text = "Export"
        Button1.Text = "ວິວ/ເບິ່ງ ສອງສະກຸນເງິນ"
        CMB_Curr.Items.Clear()
        CMB_Curr.Items.Add("EQVL")
        Call load_Cmb(" SELECT Curr  FROM Curr_For_Rate WHERE (Curr='LAK' Or Curr='THB' Or Curr='USD')  ORDER BY cnt ", "Curr", CMB_Curr)
        If CMB_Curr.Items.Count > 0 Then
            CMB_Curr.SelectedIndex = 0
        End If
        If MuLng = "L" Then
            'Button2.Text = "ທຽບບັນຊີ"
            CheckBox4.Text = "ທຽບເທົ່າເງິນ"
            If CMB_Curr.Text = "LAK" Then
                CheckBox4.Text = "ທຽບເທົ່າໂດລາ"
            ElseIf CMB_Curr.Text = "USD" Then
                CheckBox4.Text = "ທຽບເທົ່າກີບ"
            Else
                CheckBox4.Text = "ທຽບເທົ່າໂດລາ"
            End If
            Button1.Text = "ວິວ/ເບິ່ງ ສອງສະກຸນເງິນ"
            Button2.Text = "Account Statement"
            'CheckBox5.Text = "ສະແດງເອກະສານຊ້ອນທ້າຍ"
        Else
            'Button2.Text = "EQVL ACC"
            CheckBox4.Text = "EQVL Money"
            If CMB_Curr.Text = "LAK" Then
                CheckBox4.Text = "EQVL USD"
            ElseIf CMB_Curr.Text = "USD" Then
                CheckBox4.Text = "EQVL LAK"
            Else
                CheckBox4.Text = "EQVL USD"
            End If
            Button1.Text = "ວິວ/ເບິ່ງ ສອງສະກຸນເງິນ"
            Button2.Text = "Account Statement"
            'CheckBox5.Text = "Show Doc"



            If MuLng = "L" Then

                Label8.Text = "ລາຍເຊັນ1"
                Label14.Text = "ລາຍເຊັນ2"
                Label7.Text = "ລາຍເຊັນ3"
                Label12.Text = "ລາຍເຊັນ4"
                Label11.Text = "ທີ່"
                Button1.Text = "ວິວ/ເບິ່ງ ສອງສະກຸນເງິນ"
                BtnPreview.Text = "ວິວ/ເບິ່ງ"
                Label13.Text = "ສະແດງ"
            Else
                Button1.Text = "Preview 2 Curr"
                Label8.Text = "Signature1"
                Label14.Text = "Signature2"
                Label7.Text = "Signature3"
                Label12.Text = "Signature4"
                Label11.Text = "Location"
                BtnPreview.Text = "View/Pre"
                Label13.Text = "Show"
            End If

        End If
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
        'Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub

    Private Sub LoadDay()
        MdStartDate = Format(CDate(Ds.Value), "dd-MM-yyyy")
        MdToDate = Format(CDate(Dt.Value), "dd-MM-yyyy")
        'Lb.Text = "ແຕ່ວັນທີ " & MdStartDate & " ຫາວັນທີ " & MdToDate
        'Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາວັນທີ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳວັນທີ"
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub
    Private Sub LoadYear()
        MdStartDate = Format(CDate("01/1/" & Year(yy.Value)), "dd-MM-yyyy")
        MdToDate = Format(CDate("31/12/" & Year(yy.Value)), "dd-MM-yyyy")
        Lb.Text = "ປະຈຳປີ " & yy.Text
        'Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd/MM/yyyy") & " ສີ້ນສຸດວັນທີ " & Format(MdToDate, "dd/MM/yyyy")
        ' Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ສີ້ນສຸດວັນທີ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳປີ  " & Year(MdToDate)
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
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳ" & Period.Text & " ປີ " & Year(MdToDate)
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub

    Private Sub RM_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RM.CheckedChanged
        selectLoad()
    End Sub

    Private Sub RP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RP.CheckedChanged
        selectLoad()
    End Sub

    Private Sub RD_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RD.CheckedChanged
        selectLoad()
    End Sub

    Private Sub DMonth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DMonth.SelectedIndexChanged
        selectLoad()

    End Sub


    Private Sub BtnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        AddHeader()
        selectLoad()
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
            'Lb.Text = s1 & " " & s2 & "/" & Year(MdToDate)
        ElseIf RP.Checked = True Then
            Dim s1, s2, s3 As String
            LngId = 7074 + Period.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7062 : CallLngStr() : s1 = LngStr
            LngId = 7064 : CallLngStr() : s3 = LngStr
            Lb.Text = s1 & " (" & s2 & ") " & s3 & " " & Year(MdToDate)
        ElseIf RY.Checked = True Then
            Dim s1, s2, s3 As String
            LngId = 7074 + Period.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7062 : CallLngStr() : s1 = LngStr
            LngId = 7064 : CallLngStr() : s3 = LngStr
            'Lb.Text = s3 & " " & yy.Text
        End If
        CNN.Execute("DELETE  Ap_PostedLedgers ")
        CNN.Execute("DELETE  Ap_Ope_PostedLedgers_Group ")
        CNN.Execute("DELETE FROM Ap_Open_PostedLedgers ")


        Dim B_Curr As String = ""
        If CMB_Curr.SelectedIndex = 0 Then
            B_Curr = ""
        Else
            B_Curr = " AND  Curr=N'" & CMB_Curr.Text & "' "
        End If


        Dim S As Date
        S = MdStartDate
        S = DateAdd("d", CDbl(-1), MdStartDate)
        Off_Find = Off_Usr.Text : MuTable = "gen_jn." : Call Find_Company()
        If CMB_Curr.SelectedIndex = 0 Then
            If CheckBox4.Checked = True Then
                Dim s11 As String = "INSERT INTO Ap_Open_PostedLedgers( ac_code , Amount_Dr , Amount_Cr , Amt_Dr , Amt_Cr , Curr   ) " & _
                          "select ac_code , sum(amount_dr) , sum(amount_cr) ,  sum(amt_dr) , sum(amt_cr) , Curr  from gen_jn WHERE 1=1 " & B_Curr & " and     gen_jn.date_work   BETWEEN '" & "" & Format(MdStartDate, "yyyy") & "-1-1' AND '" & Format(S, "yyyy-MM-dd") & "' " & MULook2 & "  group BY ac_code , Curr "
                CNN.Execute(s11)
                Dim aa As String
                Off_Find = Off_Usr.Text : MuTable = "Open_jn." : Call Find_Company()
                aa = "INSERT INTO Ap_Open_PostedLedgers( ac_code , Amount_Dr , Amount_Cr , Amt_Dr , Amt_Cr  , Curr  ) " & _
                 "select ac_code , amount_dr , amount_cr , amt_dr , amt_cr , Curr from Open_jn  WHERE  1=1 " & B_Curr & " and     date_work='" & "" & Format(MdStartDate, "yyyy") & "-1-1' " & MULook2 & "  "
                CNN.Execute(aa)
                Off_Find = Off_Usr.Text : MuTable = "gen_jn." : Call Find_Company()
                aa = " INSERT INTO Ap_PostedLedgers( ac_code , ac_name , date_work , descrip , Open_amount , open_amt , amount_dr , amount_cr , amt_dr , amt_cr , certify , remain , Status ,  curr ) " & _
                       "select ac_code ,'' ,date_work, descrip ,0 ,0 ,amount_dr ,amount_cr , amt_dr , amt_cr ,certify, 0 , 0 ,  Curr  from gen_jn    WHERE  1=1 " & B_Curr & " and   gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & MULook2 & "   "
                CNN.Execute(aa)
            Else
              
                Dim s11 As String = "INSERT INTO Ap_Open_PostedLedgers( ac_code , Amount_Dr , Amount_Cr , Amt_Dr , Amt_Cr , Curr   ) " & _
                    "select ac_code , sum(amount_dr) , sum(amount_cr) ,  sum(Amount_dr) , sum(Amount_Cr) , Curr  from gen_jn WHERE 1=1 " & B_Curr & " and     gen_jn.date_work   BETWEEN '" & "" & Format(MdStartDate, "yyyy") & "-1-1' AND '" & Format(S, "yyyy-MM-dd") & "' " & MULook2 & "  group BY ac_code , Curr "
                CNN.Execute(s11)

                Dim aa As String
                Off_Find = Off_Usr.Text : MuTable = "Open_jn." : Call Find_Company()
                aa = "INSERT INTO Ap_Open_PostedLedgers( ac_code , Amount_Dr , Amount_Cr , Amt_Dr , Amt_Cr  , Curr  ) " & _
                 "select ac_code , amount_dr , amount_cr , amt_dr , amt_cr , Curr from Open_jn  WHERE  1=1 " & B_Curr & " and     date_work='" & "" & Format(MdStartDate, "yyyy") & "-1-1' " & MULook2 & "  "
                CNN.Execute(aa)
                Off_Find = Off_Usr.Text : MuTable = "gen_jn." : Call Find_Company()
                aa = " INSERT INTO Ap_PostedLedgers( ac_code , ac_name , date_work , descrip , Open_amount , open_amt , amount_dr , amount_cr , amt_dr , amt_cr , certify , remain , Status ,  curr ) " & _
                       "select ac_code ,'' ,date_work, descrip ,0 ,0 ,amount_dr ,amount_cr , amt_dr , amt_cr ,certify, 0 , 0 ,  Curr  from gen_jn    WHERE 1=1 " & B_Curr & " and    gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & MULook2 & "   "
                CNN.Execute(aa)

            End If
        Else
            Dim s11 As String = "INSERT INTO Ap_Open_PostedLedgers( ac_code , Amount_Dr , Amount_Cr , Amt_Dr , Amt_Cr , Curr   ) " & _
                      "select ac_code , sum(amount_dr) , sum(amount_cr) ,  sum(Amount_dr) , sum(Amount_Cr) , Curr  from gen_jn WHERE 1=1 " & B_Curr & " and     gen_jn.date_work   BETWEEN '" & "" & Format(MdStartDate, "yyyy") & "-1-1' AND '" & Format(S, "yyyy-MM-dd") & "' " & MULook2 & "  group BY ac_code , Curr "
            CNN.Execute(s11)

            Dim aa As String
            Off_Find = Off_Usr.Text : MuTable = "Open_jn." : Call Find_Company()
            aa = "INSERT INTO Ap_Open_PostedLedgers( ac_code , Amount_Dr , Amount_Cr , Amt_Dr , Amt_Cr  , Curr  ) " & _
             "select ac_code , amount_dr , amount_cr , amount_dr , amount_cr , Curr from Open_jn  WHERE  1=1 " & B_Curr & " and     date_work='" & "" & Format(MdStartDate, "yyyy") & "-1-1' " & MULook2 & "  "
            CNN.Execute(aa)
            Off_Find = Off_Usr.Text : MuTable = "gen_jn." : Call Find_Company()
            aa = " INSERT INTO Ap_PostedLedgers( ac_code , ac_name , date_work , descrip , Open_amount , open_amt , amount_dr , amount_cr , amt_dr , amt_cr , certify , remain , Status ,  curr ) " & _
                   "select ac_code ,'' ,date_work, descrip ,0 ,0 ,amount_dr ,amount_cr , amount_dr , amount_cr ,certify, 0 , 0 ,  Curr  from gen_jn    WHERE 1=1 " & B_Curr & " and    gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & MULook2 & "   "
            CNN.Execute(aa)

        End If

        CNN.Execute("Delete Ap_Ope_PostedLedgers_Group ")
        CNN.Execute("Insert Into Ap_Ope_PostedLedgers_Group (Ac_Code  ,Amount, Amt  )" & _
                     " Select Ac_Code ,Sum(Amount_Dr - Amount_Cr)  ,Sum(Amt_Dr - Amt_Cr)  from Ap_Open_PostedLedgers group by Ac_Code ")
        CNN.Execute("Update Ap_PostedLedgers set  Ap_PostedLedgers.Open_Amt = Ap_Ope_PostedLedgers_Group.Amt  , Ap_PostedLedgers.Open_Amount = Ap_Ope_PostedLedgers_Group.Amount  From Ap_PostedLedgers , Ap_Ope_PostedLedgers_Group  Where Ap_PostedLedgers.Ac_Code = Ap_Ope_PostedLedgers_Group.Ac_Code   ")

        CNN.Execute("Delete Ap_Ope_PostedLedgers_Group ")
        CNN.Execute("Insert Into Ap_Ope_PostedLedgers_Group (Ac_Code , Amount  , Amt )" & _
                  " Select Ac_Code ,Sum(Amount_Dr - Amount_Cr)   ,Sum(Amt_Dr - Amt_Cr)   from Ap_Open_PostedLedgers group by Ac_Code ")
        CNN.Execute("Update Ap_PostedLedgers set Ap_PostedLedgers.Open_Amount = Ap_Ope_PostedLedgers_Group.Amount , Ap_PostedLedgers.Open_Amt = Ap_Ope_PostedLedgers_Group.Amt   From Ap_PostedLedgers , Ap_Ope_PostedLedgers_Group  Where Ap_PostedLedgers.Ac_Code = Ap_Ope_PostedLedgers_Group.Ac_Code  ")

        CNN.Execute(" delete Ap_PostedLedgers_Rem ")
        CNN.Execute(" insert into Ap_PostedLedgers_Rem  (ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work)  " & _
         " select ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work from Ap_PostedLedgers order by  Ac_Code,Date_Work , certify ,cnt Asc  ")
        CNN.Execute("delete Ap_PostedLedgers   ")
        CNN.Execute("insert into Ap_PostedLedgers  (ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work)  " & _
         " select ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work from Ap_PostedLedgers_Rem order by cnt   ")

        CNN.Execute(" delete Ap_PostedLedgers_Rem ")
        CNN.Execute(" insert into Ap_PostedLedgers_Rem (Ac_Code , cnt_Mat ,remain )  " & _
          " select Ac_Code , cnt,Open_amt+(select SUM(Amt_Dr-Amt_cr) from Ap_PostedLedgers  where Ac_Code = x.ac_code    And cnt <= x.cnt )as Rem from Ap_PostedLedgers as x  order by  cnt Asc  ")


        CNN.Execute("Update Ap_PostedLedgers set Ap_PostedLedgers.remain = Ap_PostedLedgers_Rem.remain from Ap_PostedLedgers , Ap_PostedLedgers_Rem where  Ap_PostedLedgers.cnt = Ap_PostedLedgers_Rem.cnt_Mat  " & _
            "  update	 Ap_PostedLedgers set remain=0 where remain is null ")

        CNN.Execute(" Update Ap_PostedLedgers set Open_dr = 0 where  Open_dr is Null Update Ap_PostedLedgers set Open_cr = 0 where  Open_cr is Null Update Ap_PostedLedgers set Open_amount = 0 where  Open_amount is Null Update Ap_PostedLedgers set open_amt = 0 where  open_amt is Null Update Ap_PostedLedgers set amount_dr = 0 where  amount_dr is Null Update Ap_PostedLedgers set amount_cr = 0 where  amount_cr is Null Update Ap_PostedLedgers set amt_dr = 0 where  amt_dr is Null Update Ap_PostedLedgers set amt_cr = 0 where  amt_cr is Null Update Ap_PostedLedgers set remain = 0 where  remain is Null ")

        If CheckBox1.Checked = True Then
            CNN.Execute(" delete Ap_PostedLedgers_Rem ")
            CNN.Execute(" insert into Ap_PostedLedgers_Rem (Ac_Code , cnt_Mat ,remain )  " & _
              " select Ac_Code , cnt,Open_amount+(select SUM(Amount_Dr-Amount_cr) from Ap_PostedLedgers  where Ac_Code = x.ac_code    And cnt <= x.cnt )as Rem_Curr from Ap_PostedLedgers as x  order by  cnt Asc  ")
            CNN.Execute("Update Ap_PostedLedgers set Ap_PostedLedgers.Rem_Curr = Ap_PostedLedgers_Rem.remain from Ap_PostedLedgers , Ap_PostedLedgers_Rem where  Ap_PostedLedgers.cnt = Ap_PostedLedgers_Rem.cnt_Mat  " & _
                "  update	 Ap_PostedLedgers set remain=0 where remain is null ")
        End If

        Off_Find = Off_Usr.Text : MuTable = "Open_jn." : Call Find_Company()
        CNN.Execute("update Open_jn set Lck=0 where year(date_work)= " & Format(MdStartDate, "yyyy") & "  " & MULook2 & "")
        CNN.Execute("update Open_jn set Lck=1 from Open_jn, Ap_PostedLedgers   where year(Open_jn.date_work)= " & Format(MdStartDate, "yyyy") & " And  Open_jn.ac_code =  Ap_PostedLedgers.ac_code " & MULook2 & "")
        If CMB_Curr.SelectedIndex = 0 Then
            Dim s7 As String = " Insert into Ap_PostedLedgers (ac_code,Open_dr,Open_Cr,Open_amount,open_amt,amount_dr,amount_cr,amt_dr,amt_cr,Curr,remain,Rem_Curr,Status)  " & _
                             "select ac_code,0,0, Sum(amount_dr-amount_cr),Sum(amt_dr-amt_cr),0,0,0,0,Curr,0,0,0 from Open_jn where   1=1 " & B_Curr & " and  Lck=0 And year(date_work)= " & Format(MdStartDate, "yyyy") & "  " & MULook2 & " Group by ac_code , Curr"
            CNN.Execute(s7)
        Else
            Dim s7 As String = " Insert into Ap_PostedLedgers (ac_code,Open_dr,Open_Cr,Open_amount,open_amt,amount_dr,amount_cr,amt_dr,amt_cr,Curr,remain,Rem_Curr,Status)  " & _
                         "select ac_code,0,0, Sum(amount_dr-amount_cr),Sum(amt_dr-amt_cr),0,0,0,0,Curr,0,0,0 from Open_jn where   1=1 " & B_Curr & " and  Lck=0 And year(date_work)= " & Format(MdStartDate, "yyyy") & "  " & MULook2 & " Group by ac_code , Curr"
            CNN.Execute(s7)
        End If


        CNN.Execute("Update Ap_PostedLedgers set open_amt = Ap_Ope_PostedLedgers_Group.Amt from Ap_PostedLedgers ,Ap_Ope_PostedLedgers_Group  where  Ap_PostedLedgers.Ac_Code = Ap_Ope_PostedLedgers_Group.Ac_Code")

        'CNN.Execute("update Ap_PostedLedgers set Ap_PostedLedgers.ac_name=Acc_Code.Name_L from Ap_PostedLedgers , Acc_Code where Ap_PostedLedgers.Ac_Code = Acc_Code.Ac_Code")
        If MuLng = "L" Then
            CNN.Execute("update Ap_PostedLedgers set Ap_PostedLedgers.ac_name=gen_jn.ac_name, Ap_PostedLedgers.last_user=gen_jn.last_user from Ap_PostedLedgers , gen_jn where Ap_PostedLedgers.Ac_Code = gen_jn.Ac_Code and   Ap_PostedLedgers.certify =gen_jn.certify ")
        Else
            CNN.Execute("update Ap_PostedLedgers set Ap_PostedLedgers.ac_name=gen_jn.ac_namee, Ap_PostedLedgers.last_user=gen_jn.last_user from Ap_PostedLedgers , gen_jn where Ap_PostedLedgers.Ac_Code = gen_jn.Ac_Code and   Ap_PostedLedgers.certify =gen_jn.certify ")

        End If

        If ChDs.Checked = False Then
            CNN.Execute("update Ap_PostedLedgers set descrip = ac_name")

        End If
        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As    RptSjUd ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
        LngId = "7002" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDP ,"
        If CheckBox1.Checked = True Then
            LngId = "7044" : CallLngStr() : MuLngRpt = MuLngRpt & "N'ລາຍງານບັນຊີສໍາຮອງ " & DTDATE02 & "' As Crl_RptName ,"
        Else
            LngId = "7067" : CallLngStr() : MuLngRpt = MuLngRpt & "N'ລາຍງານບັນຊີສໍາຮອງ " & DTDATE02 & "' As Crl_RptName ,"
        End If
        LngId = "7004" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Descrip ,"
        LngId = "7005" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_No ,"
        LngId = "7006" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Ac_Code ,"
        LngId = "7007" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Certify ,"
        LngId = "7008" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Dr ,"
        LngId = "7009" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Cr ,"
        LngId = "7068" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Amt_LAK ,"
        LngId = "7011" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Sign ,"
        LngId = "7012" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_SignNme ,"
        LngId = "7026" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InDate ,"
        LngId = "7028" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_AmtCurr ,"
        LngId = "7029" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Curr ,"
        LngId = "7030" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Rate ,"
        LngId = "7031" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_User ,"
        LngId = "7042" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_OpenAmt ,"
        LngId = "7043" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RemAmt,"
        If Cfrom.Text <> "" Then
            SLF = "SELECT  " & MuLngRpt & "  *   FROM Ap_PostedLedgers Where   left(ac_code,'" & Len(Cfrom.Text) & "')>='" & Cfrom.Text & "' and  left(ac_code,'" & Len(Cto.Text) & "')<='" & Cto.Text & "'   order by  cnt    asc "
        Else
            SLF = "SELECT  " & MuLngRpt & "  *   FROM Ap_PostedLedgers       order by  Ac_Code,Date_Work , certify ,cnt     asc "
        End If
        Call LoadLoGO()
        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            .Open(SLF, CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With

        If CMB_Curr.Text = "LAK" Then
            CURR01 = "ຫົວໜ່ວຍ : ກີບ"
        ElseIf CMB_Curr.Text = "USD" Then
            CURR01 = "ຫົວໜ່ວຍ : ໂດລາ"
        Else
            CURR01 = "ຫົວໜ່ວຍ : ກີບ"
        End If

        Call SUM_Ledgers()

        If CheckBox1.Checked = True Then
            Dim Rpt As New CryPostedLedgersCurr3
            'Dim Rpt As New CrystalReport_posted_ledgers_ALL()
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
            myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("txtprint_user"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = MUserName
            myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("pp"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = TxtPP.Text
            myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("Text3"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = CURR01

            Rpt.SetDataSource(Rs)
            FmPreview.ReportViewer.ReportSource = Rpt
        Else
            Dim Rpt As New Object
            If MuLng = "L" Then
                Rpt = New CryPostedLedgers3
            Else
                Rpt = New CryPostedLedgers_Eng
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
            myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("txtprint_user"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = MUserName
            myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("pp"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = TxtPP.Text
            myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("txtCurr"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = Label6.Text & "" & CMB_Curr.Text

            myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("Text3"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = CURR01

            myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("C_01"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = Format(CDbl(C_01), "#,##0")
            myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("C_02"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = Format(CDbl(C_02), "#,##0")
            myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("Open_amt"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = Format(CDbl(Open_amt), "#,##0")
            myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("Amt_dr"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = Format(CDbl(Amt_dr), "#,##0")
            myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("Amt_cr"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = Format(CDbl(Amt_cr), "#,##0")
            myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("Rem_amt"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = Format(CDbl(Rem_amt), "#,##0")

            Rpt.SetDataSource(Rs)
            FmPreview.ReportViewer.ReportSource = Rpt
        End If
        FmPreview.ReportViewer.DisplayGroupTree = False
        'FmPreview.MdiParent = FmMain
        FmPreview.WindowState = FormWindowState.Maximized
        FmPreview.Show()
        FmPreview.Focus()
    End Sub
    Dim C_01, C_02, Open_amt, Amt_dr, Amt_cr, Rem_amt As Double
    Private Sub SUM_Ledgers()
        Dim RSC1 As New ADODB.Recordset
        Dim RSC2 As New ADODB.Recordset
        Dim RSC3 As New ADODB.Recordset
        LoadSqlData("SELECT Count(Ac_Code) as Ac_Code FROM Ap_PostedLedgers where Amt_dr<>0 and   left(ac_code,'" & Len(Cfrom.Text) & "')>='" & Cfrom.Text & "' and  left(ac_code,'" & Len(Cto.Text) & "')<='" & Cto.Text & "'   ", RSC1)
        If RSC1.RecordCount <> 0 Then
            C_01 = Format(CDbl(RSC1.Fields("Ac_Code").Value), "#,##0")
        End If

        LoadSqlData("SELECT Count(Ac_Code) as Ac_Code FROM Ap_PostedLedgers where Amt_Cr<>0 and   left(ac_code,'" & Len(Cfrom.Text) & "')>='" & Cfrom.Text & "' and  left(ac_code,'" & Len(Cto.Text) & "')<='" & Cto.Text & "'    ", RSC2)
        If RSC2.RecordCount <> 0 Then
            C_02 = Format(CDbl(RSC2.Fields("Ac_Code").Value), "#,##0")
        End If

        LoadSqlData("SELECT SUM(Open_amt) as Open_amt, SUM(Amt_dr) as Amt_dr, SUM(Amt_cr) as Amt_cr FROM Ap_PostedLedgers where   left(ac_code,'" & Len(Cfrom.Text) & "')>='" & Cfrom.Text & "' and  left(ac_code,'" & Len(Cto.Text) & "')<='" & Cto.Text & "'     ", RSC3)
        If RSC.RecordCount <> 0 Then
            Open_amt = Format(CDbl(RSC3.Fields("Open_amt").Value), "#,##0")
            Amt_dr = Format(CDbl(RSC3.Fields("Amt_dr").Value), "#,##0")
            Amt_cr = Format(CDbl(RSC3.Fields("Amt_cr").Value), "#,##0")
        End If
        Rem_amt = CDbl(Open_amt) + CDbl(Amt_dr) - CDbl(Amt_cr)
        Rem_amt = Format(CDbl(Rem_amt), "#,##0")
    End Sub
    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Cfrom.KeyPress

    End Sub


    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cfrom.TextChanged
        Cto.Text = Cfrom.Text
    End Sub

    Private Sub BtnExit_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Close()
    End Sub







    Private Sub Myy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Myy.ValueChanged

    End Sub



    Private Sub Pyy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Pyy.ValueChanged

    End Sub



    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cto.TextChanged

    End Sub

    Private Sub Ds_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ds.ValueChanged
        Dt.Value = Ds.Value

    End Sub

    Private Sub Dt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Dt.ValueChanged
        'Call selectLoad()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        AddHeader()
        selectLoad()
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
            Lb.Text = s1 & " " & s2 & "/" & Year(MdToDate)
        ElseIf RP.Checked = True Then
            Dim s1, s2, s3 As String
            LngId = 7074 + Period.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7062 : CallLngStr() : s1 = LngStr
            LngId = 7064 : CallLngStr() : s3 = LngStr
            Lb.Text = s1 & " (" & s2 & ") " & s3 & " " & Year(MdToDate)
        ElseIf RY.Checked = True Then
            Dim s1, s2, s3 As String
            LngId = 7074 + Period.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7062 : CallLngStr() : s1 = LngStr
            LngId = 7064 : CallLngStr() : s3 = LngStr
            Lb.Text = s3 & " " & yy.Text
        End If
        CNN.Execute("DELETE  Ap_PostedLedgers ")
        CNN.Execute("DELETE  Ap_Ope_PostedLedgers_Group ")
        CNN.Execute("DELETE FROM Ap_Open_PostedLedgers ")
        Dim S As Date
        S = MdStartDate
        S = DateAdd("d", CDbl(-1), MdStartDate)
        Off_Find = Off_Usr.Text : MuTable = "gen_jn." : Call Find_Company()
        Dim s11 As String = "INSERT INTO Ap_Open_PostedLedgers( ac_code , Amount_Dr , Amount_Cr , Amt_Dr , Amt_Cr , Curr   ) " & _
         "select ac_code , sum(amount_dr) , sum(amount_cr) ,  sum(amt_dr) , sum(amt_cr) , Curr  from gen_jn WHERE   gen_jn.date_work   BETWEEN '" & "" & Format(MdStartDate, "yyyy") & "-1-1' AND '" & Format(S, "yyyy-MM-dd") & "' " & MULook2 & "  group BY ac_code , Curr "
        CNN.Execute(s11)
        Dim aa As String
        Off_Find = Off_Usr.Text : MuTable = "Open_jn." : Call Find_Company()
        aa = "INSERT INTO Ap_Open_PostedLedgers( ac_code , Amount_Dr , Amount_Cr , Amt_Dr , Amt_Cr  , Curr  ) " & _
         "select ac_code , amount_dr , amount_cr , amt_dr , amt_cr , Curr from Open_jn  WHERE    date_work='" & "" & Format(MdStartDate, "yyyy") & "-1-1' " & MULook2 & "  "
        CNN.Execute(aa)
        Off_Find = Off_Usr.Text : MuTable = "gen_jn." : Call Find_Company()
        aa = " INSERT INTO Ap_PostedLedgers( ac_code , ac_name , date_work , descrip , Open_amount , open_amt , amount_dr , amount_cr , amt_dr , amt_cr , certify , remain , Status ,  curr ) " & _
               "select ac_code ,'' ,date_work, descrip ,0 ,0 ,amount_dr ,amount_cr , amt_dr , amt_cr ,certify, 0 , 0 ,  Curr  from gen_jn    WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & MULook2 & "   "
        CNN.Execute(aa)

        CNN.Execute("Delete Ap_Ope_PostedLedgers_Group ")
        CNN.Execute("Insert Into Ap_Ope_PostedLedgers_Group (Ac_Code  ,Amount, Amt  )" & _
                     " Select Ac_Code ,Sum(Amount_Dr - Amount_Cr)  ,Sum(Amt_Dr - Amt_Cr)  from Ap_Open_PostedLedgers group by Ac_Code ")
        CNN.Execute("Update Ap_PostedLedgers set  Ap_PostedLedgers.Open_Amt = Ap_Ope_PostedLedgers_Group.Amt  , Ap_PostedLedgers.Open_Amount = Ap_Ope_PostedLedgers_Group.Amount  From Ap_PostedLedgers , Ap_Ope_PostedLedgers_Group  Where Ap_PostedLedgers.Ac_Code = Ap_Ope_PostedLedgers_Group.Ac_Code   ")

        CNN.Execute("Delete Ap_Ope_PostedLedgers_Group ")
        CNN.Execute("Insert Into Ap_Ope_PostedLedgers_Group (Ac_Code , Amount  , Amt )" & _
                  " Select Ac_Code ,Sum(Amount_Dr - Amount_Cr)   ,Sum(Amt_Dr - Amt_Cr)   from Ap_Open_PostedLedgers group by Ac_Code ")
        CNN.Execute("Update Ap_PostedLedgers set Ap_PostedLedgers.Open_Amount = Ap_Ope_PostedLedgers_Group.Amount , Ap_PostedLedgers.Open_Amt = Ap_Ope_PostedLedgers_Group.Amt   From Ap_PostedLedgers , Ap_Ope_PostedLedgers_Group  Where Ap_PostedLedgers.Ac_Code = Ap_Ope_PostedLedgers_Group.Ac_Code  ")

        CNN.Execute(" delete Ap_PostedLedgers_Rem ")
        CNN.Execute(" insert into Ap_PostedLedgers_Rem  (ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work)  " & _
         " select ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work from Ap_PostedLedgers order by  Ac_Code,Date_Work , certify ,cnt Asc  ")
        CNN.Execute("delete Ap_PostedLedgers   ")
        CNN.Execute("insert into Ap_PostedLedgers  (ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work)  " & _
         " select ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work from Ap_PostedLedgers_Rem order by cnt   ")

        CNN.Execute(" delete Ap_PostedLedgers_Rem ")
        CNN.Execute(" insert into Ap_PostedLedgers_Rem (Ac_Code , cnt_Mat ,remain )  " & _
          " select Ac_Code , cnt,Open_amt+(select SUM(Amt_Dr-Amt_cr) from Ap_PostedLedgers  where Ac_Code = x.ac_code    And cnt <= x.cnt )as Rem from Ap_PostedLedgers as x  order by  cnt Asc  ")


        CNN.Execute("Update Ap_PostedLedgers set Ap_PostedLedgers.remain = Ap_PostedLedgers_Rem.remain from Ap_PostedLedgers , Ap_PostedLedgers_Rem where  Ap_PostedLedgers.cnt = Ap_PostedLedgers_Rem.cnt_Mat  " & _
            "  update	 Ap_PostedLedgers set remain=0 where remain is null ")

        CNN.Execute(" Update Ap_PostedLedgers set Open_dr = 0 where  Open_dr is Null Update Ap_PostedLedgers set Open_cr = 0 where  Open_cr is Null Update Ap_PostedLedgers set Open_amount = 0 where  Open_amount is Null Update Ap_PostedLedgers set open_amt = 0 where  open_amt is Null Update Ap_PostedLedgers set amount_dr = 0 where  amount_dr is Null Update Ap_PostedLedgers set amount_cr = 0 where  amount_cr is Null Update Ap_PostedLedgers set amt_dr = 0 where  amt_dr is Null Update Ap_PostedLedgers set amt_cr = 0 where  amt_cr is Null Update Ap_PostedLedgers set remain = 0 where  remain is Null ")

        If CheckBox1.Checked = True Then
            CNN.Execute(" delete Ap_PostedLedgers_Rem ")
            CNN.Execute(" insert into Ap_PostedLedgers_Rem (Ac_Code , cnt_Mat ,remain )  " & _
              " select Ac_Code , cnt,Open_amount+(select SUM(Amount_Dr-Amount_cr) from Ap_PostedLedgers  where Ac_Code = x.ac_code    And cnt <= x.cnt )as Rem_Curr from Ap_PostedLedgers as x  order by  cnt Asc  ")
            CNN.Execute("Update Ap_PostedLedgers set Ap_PostedLedgers.Rem_Curr = Ap_PostedLedgers_Rem.remain from Ap_PostedLedgers , Ap_PostedLedgers_Rem where  Ap_PostedLedgers.cnt = Ap_PostedLedgers_Rem.cnt_Mat  " & _
                "  update	 Ap_PostedLedgers set remain=0 where remain is null ")
        End If

        'CNN.Execute("Update Ap_Ope_PostedLedgers_Group set Lck = 1")
        'CNN.Execute(" Update Ap_Ope_PostedLedgers_Group set Lck = 0 from   Ap_Ope_PostedLedgers_Group , Ap_PostedLedgers where  Ap_Ope_PostedLedgers_Group.Ac_Code= Ap_PostedLedgers.Ac_Code ")
        'CNN.Execute("INSERT INTO Ap_PostedLedgers( ac_code , Open_amount , open_amt    ) select  ac_code ,  Amount , Amt   from  Ap_Ope_PostedLedgers_Group where Lck = 1 ")

        Off_Find = Off_Usr.Text : MuTable = "Open_jn." : Call Find_Company()
        CNN.Execute("update Open_jn set Lck=0 where year(date_work)= " & Format(MdStartDate, "yyyy") & "  " & MULook2 & "")
        CNN.Execute("update Open_jn set Lck=1 from Open_jn, Ap_PostedLedgers   where year(Open_jn.date_work)= " & Format(MdStartDate, "yyyy") & " And  Open_jn.ac_code =  Ap_PostedLedgers.ac_code " & MULook2 & "")
        Dim s7 As String = " Insert into Ap_PostedLedgers (ac_code,Open_dr,Open_Cr,Open_amount,open_amt,amount_dr,amount_cr,amt_dr,amt_cr,Curr,remain,Rem_Curr,Status)  " & _
                    "select ac_code,0,0, Sum(amount_dr-amount_cr),Sum(amt_dr-amt_cr),0,0,0,0,Curr,0,0,0 from Open_jn where Lck=0 And year(date_work)= " & Format(MdStartDate, "yyyy") & "  " & MULook2 & " Group by ac_code , Curr"
        CNN.Execute(s7)

        CNN.Execute("Update Ap_PostedLedgers set open_amt = Ap_Ope_PostedLedgers_Group.Amt from Ap_PostedLedgers ,Ap_Ope_PostedLedgers_Group  where  Ap_PostedLedgers.Ac_Code = Ap_Ope_PostedLedgers_Group.Ac_Code")

        'CNN.Execute("update Ap_PostedLedgers set Ap_PostedLedgers.ac_name=Acc_Code.Name_L from Ap_PostedLedgers , Acc_Code where Ap_PostedLedgers.Ac_Code = Acc_Code.Ac_Code")
        CNN.Execute("update Ap_PostedLedgers set Ap_PostedLedgers.ac_name=gen_jn.ac_name from Ap_PostedLedgers , gen_jn where Ap_PostedLedgers.Ac_Code = gen_jn.Ac_Code and   Ap_PostedLedgers.certify =gen_jn.certify ")

        If ChDs.Checked = False Then
            CNN.Execute("update Ap_PostedLedgers set descrip = ac_name")

        End If
        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As    RptSjUd ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
        LngId = "7002" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDP ,"
        If CheckBox1.Checked = True Then
            LngId = "7044" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
        Else
            LngId = "7067" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
        End If
        LngId = "7004" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Descrip ,"
        LngId = "7005" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_No ,"
        LngId = "7006" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Ac_Code ,"
        LngId = "7007" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Certify ,"
        LngId = "7008" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Dr ,"
        LngId = "7009" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Cr ,"
        LngId = "7068" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Amt_LAK ,"
        LngId = "7011" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Sign ,"
        LngId = "7012" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_SignNme ,"
        LngId = "7026" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InDate ,"
        LngId = "7028" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_AmtCurr ,"
        LngId = "7029" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Curr ,"
        LngId = "7030" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Rate ,"
        LngId = "7031" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_User ,"
        LngId = "7042" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_OpenAmt ,"
        LngId = "7043" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RemAmt,"
        If Cfrom.Text <> "" Then
            SLF = "SELECT  " & MuLngRpt & "  *   FROM Ap_PostedLedgers Where   left(ac_code,'" & Len(Cfrom.Text) & "')>='" & Cfrom.Text & "' and  left(ac_code,'" & Len(Cto.Text) & "')<='" & Cto.Text & "'   order by  cnt    asc "
        Else
            SLF = "SELECT  " & MuLngRpt & "  *   FROM Ap_PostedLedgers       order by  Ac_Code,Date_Work , certify ,cnt     asc "
        End If
        Call LoadLoGO()
        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            .Open(SLF, CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        If CheckBox1.Checked = True Then
            Dim Rpt As New CryPostedLedgersCurr
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
            FmPreview.ReportViewer.ReportSource = Rpt
        Else
            Dim Rpt As New CryPostedLedgers
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
            FmPreview.ReportViewer.ReportSource = Rpt
        End If

        FmPreview.ReportViewer.ExportReport()
        FmPreview = Nothing

        'FmPreview.ReportViewer.DisplayGroupTree = False
        ''FmPreview.MdiParent = FmMain
        'FmPreview.WindowState = FormWindowState.Maximized
        'FmPreview.Show()
        'FmPreview.Focus()
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
        Call RateSetting()
        txtRate.Text = Format(MD_Rate, "#,##0.00")

        If MuLng = "L" Then
            'Button2.Text = "ທຽບບັນຊີ"
            CheckBox4.Text = "ທຽບເທົ່າເງິນ"
            If CMB_Curr.Text = "LAK" Then
                CheckBox4.Text = "ທຽບເທົ່າໂດລາ"
            ElseIf CMB_Curr.Text = "USD" Then
                CheckBox4.Text = "ທຽບເທົ່າກີບ"
            Else
                CheckBox4.Text = "ທຽບເທົ່າໂດລາ"
            End If
        Else
            'Button2.Text = "EQVL ACC"
            CheckBox4.Text = "EQVL Money"
            If CMB_Curr.Text = "LAK" Then
                CheckBox4.Text = "EQVL USD"
            ElseIf CMB_Curr.Text = "USD" Then
                CheckBox4.Text = "EQVL LAK"
            Else
                CheckBox4.Text = "EQVL USD"
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        AddHeader()
        selectLoad()
        Load_all()
        Dim aa As String = ""
        CNN.Execute(" update RPT_Bank_Book set  Referno =gen_jn.Referno from gen_jn where  RPT_Bank_Book.certify= gen_jn.certify and RPT_Bank_Book.cat_id= gen_jn.cnt  ")
        aa = "update RPT_Ledgers set  Referno =gen_jn.Referno from gen_jn where  RPT_Ledgers.certify= gen_jn.certify and RPT_Ledgers.cat_id= gen_jn.cnt    " & _
                     " and   RPT_Ledgers. Date_work = gen_jn.Date_work  and   RPT_Ledgers.Ac_code  = gen_jn.Ac_code and  RPT_Ledgers.amt_dr  + RPT_Ledgers.amt_cr   = gen_jn.amount  "
        CNN.Execute(aa)

        'CNN.Execute("update RPT_Bank_Book set  Ac_code_N =ACC_CODE.AC_CODE2 from ACC_CODE where  RPT_Bank_Book.Ac_code= ACC_CODE.AC_CODE ")
        'CNN.Execute("update RPT_Bank_Book set  Ac_code_dr_N =ACC_CODE.AC_CODE2 from ACC_CODE where  RPT_Bank_Book.Ac_code= ACC_CODE.AC_CODE ")
        'CNN.Execute("update RPT_Bank_Book set  Ac_code_cr_N =ACC_CODE.AC_CODE2 from ACC_CODE where  RPT_Bank_Book.Ac_code= ACC_CODE.AC_CODE ")

        Dim rs As New ADODB.Recordset

        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As    RptSjUd ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
        LngId = "7002" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDP ,"
        If CheckBox1.Checked = True Then
            LngId = "7044" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
        Else
            LngId = "7067" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
        End If
        LngId = "7004" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Descrip ,"
        LngId = "7005" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_No ,"
        LngId = "7006" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Ac_Code ,"
        LngId = "7007" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Certify ,"
        LngId = "7008" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Dr ,"
        LngId = "7009" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Cr ,"
        LngId = "7068" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Amt_LAK ,"
        LngId = "7011" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Sign ,"
        LngId = "7012" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_SignNme ,"
        LngId = "7026" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InDate ,"
        LngId = "7028" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_AmtCurr ,"
        LngId = "7029" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Curr ,"
        LngId = "7030" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Rate ,"
        LngId = "7031" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_User ,"
        LngId = "7042" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_OpenAmt ,"
        LngId = "7043" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RemAmt,"
        'If Cfrom.Text <> "" Then
        '    SLF = "SELECT  " & MuLngRpt & "  *   FROM Ap_PostedLedgers Where   left(ac_code,'" & Len(Cfrom.Text) & "')>='" & Cfrom.Text & "' and  left(ac_code,'" & Len(Cto.Text) & "')<='" & Cto.Text & "'   order by  cnt    asc "
        'Else
        '    SLF = "SELECT  " & MuLngRpt & "  *   FROM Ap_PostedLedgers       order by  Ac_Code,Date_Work , certify ,cnt     asc "
        'End If

        If Cfrom.Text <> "" Then
            'Dim ss As String = " SELECT  " & mformat & "  as mformat  ,     RPT_Bank_Book.*   where 1=1       order by   date_work "
            'Call LoadSqlData(ss, rs)
            'left(ac_code,'" & Len(Cfrom.Text) & "')>='" & Cfrom.Text & "' and  left(ac_code,'" & Len(Cto.Text) & "')<='" & Cto.Text & "'   order by  cnt    asc 
            Dim ss As String = " SELECT  " & mformat & "  as mformat  ,  " & MuLngRpt & "  RPT_Bank_Book.* from RPT_Bank_Book   " & _
                  " where 1=1 and  left(ac_code,'" & Len(Cfrom.Text) & "')>='" & Cfrom.Text & "' and  left(ac_code,'" & Len(Cto.Text) & "')<='" & Cto.Text & "'   order by  cnt    asc  "
            Call LoadSqlData(ss, rs)
        Else
            Dim ss As String = " SELECT  " & mformat & "  as mformat  ,  " & MuLngRpt & "   RPT_Bank_Book.*  from RPT_Bank_Book  where 1=1   order by   date_work "
            Call LoadSqlData(ss, rs)



        End If
        Call LoadLoGO()
        If rs.RecordCount = 0 Then MsgBox("Data empty", vbInformation, "Check") : Exit Sub

        'Dim Rpt As New CryPostedLedgersCurr
        Dim Rpt As New CrystalReport_posted_ledgers_ALL()
        'If MdShowLOGO = 1 Then
        '    Rpt.Subreports(0).SetDataSource(RsLOGO)
        'End If


        'Dim Rpt1 = New CrystalReport_posted_ledgers_ALL
        Dim frm1 = New FmPreview
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

        Rpt.SetDataSource(rs)
        Rpt.Refresh()
        frm1.ReportViewer.ReportSource = Rpt
        frm1.ReportViewer.DisplayGroupTree = False
        frm1.WindowState = FormWindowState.Maximized
        frm1.Show()



    End Sub


    Private Sub Load_all()
        Dim aa As String
        'Dim M_Month1 As Integer
        aa = "    delete RPT_Bank_Book "
        CNN.Execute(aa)
        aa = "    delete RPT_Bank_Book_insert "
        CNN.Execute(aa)
        aa = "    delete RPT_Bank_Book_Group "
        CNN.Execute(aa)
        CNN.CommandTimeout = 0
        Dim S As Date
        'MdStartDate = Format(CDate(dpFromDate.Value), "dd-MM-yyyy")
        S = MdStartDate
        S = DateAdd("d", CDbl(-1), MdStartDate)

        aa = "  insert into   RPT_Bank_Book_insert (Ac_code,opening,opening_USD)   SELECT Ac_code,sum(amt_dr-amt_cr) ,sum(amount_dr-amount_cr) from  gen_jn   " & _
               " INNER JOIN AP_Office on gen_jn.company = AP_Office .sub_ID     " & _
             " where 1=1 " & sql & "  and date_work   <  '" & Format(MdStartDate, "yyyy-MM-dd") & "'  and  Year(date_work)='" & Format(MdStartDate, "yyyy") & "'   group by Ac_code "
        CNN.Execute(aa)

        aa = " insert into   RPT_Bank_Book_insert (Ac_code,opening,opening_USD) SELECT  ac_code,sum(amt_dr-amt_cr) ,sum(amount_dr-amount_cr)   from Open_jn      " & _
               " INNER JOIN AP_Office on Open_jn.company = AP_Office .sub_ID     " & _
             " where 1=1 " & sql & "   and   year(date_work)='" & Format(MdStartDate, "yyyy") & "'  group by Ac_code"
        CNN.Execute(aa)

        ' aa = "INSERT INTO Ap_Open_PostedLedgers( ac_code , Amount_Dr , Amount_Cr , Amt_Dr , Amt_Cr  , Curr  ) " & _
        '"select ac_code , amount_dr , amount_cr , amt_dr , amt_cr , Curr from Open_jn  WHERE  1=1 " & B_Curr & " and     date_work='" & "" & Format(MdStartDate, "yyyy") & "-1-1' " & MULook2 & "  "
        ' CNN.Execute(aa)

        aa = "  insert into   RPT_Bank_Book_Group (Ac_code,opening,opening_USD) select Ac_code,sum(opening),sum(opening_USD) from   RPT_Bank_Book_insert group by Ac_code "
        CNN.Execute(aa)
        aa = " insert into  RPT_Bank_Book (Date_work,certify,descrip,Ac_code,Ac_nm,amt_dr,amt_cr,amt_USD_dr,amt_USD_cr,Curr_i,Rate_i,Cat_ID,cheque_no,Com_id)   " & _
              "   SELECT  Date_work,certify,descrip,ac_code,Ac_name,amt_dr,amt_cr,amount_dr,amount_cr,Curr_i,Rate_i,gen_jn.cnt,cheque_no,Com_id   from  gen_jn   " & _
               " INNER JOIN AP_Office on gen_jn.company = AP_Office .sub_ID     " & _
             " where 1=1 " & sql & "   and Date_work BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'   "
        CNN.Execute(aa)
        CNN.CommandTimeout = 0
        aa = " update RPT_Bank_Book set opening = RPT_Bank_Book_Group.opening, opening_USD = RPT_Bank_Book_Group.opening_USD from RPT_Bank_Book_Group where RPT_Bank_Book.Ac_code=RPT_Bank_Book_Group.Ac_code "
        CNN.Execute(aa)
        aa = "  update	 RPT_Bank_Book set opening=0 where opening is null  "
        CNN.Execute(aa)
        aa = "  update	 RPT_Bank_Book set opening_USD=0 where opening_USD is null  "
        CNN.Execute(aa)

        aa = "  update	 RPT_Bank_Book set Rem =0 where Rem  is null  "
        CNN.Execute(aa)
        aa = "  update	 RPT_Bank_Book set Rem_USD=0 where Rem_USD is null  "
        CNN.Execute(aa)

        aa = " update RPT_Bank_Book set opening_USD = RPT_Bank_Book_Group.opening_USD from RPT_Bank_Book_Group where RPT_Bank_Book.Ac_code=RPT_Bank_Book_Group.Ac_code "
        CNN.Execute(aa)
        aa = "  update	 RPT_Bank_Book set opening_USD=0 where opening_USD is null  "
        CNN.Execute(aa)
        CNN.CommandTimeout = 0

        aa = "update RPT_Bank_Book set lck=0"
        CNN.Execute(aa)
        aa = " insert into RPT_Bank_Book (Date_work,certify,descrip,Ac_code,Ac_nm,amt_dr,amt_cr,amt_USD_dr,amt_USD_cr,Curr_i,Rate_i,Activity_id,Cat_ID,cheque_no,Com_id, opening,opening_USD, rem, lck) " & _
  " select Date_work,certify,descrip,Ac_code,Ac_nm,amt_dr,amt_cr,amt_USD_dr,amt_USD_cr,Curr_i,Rate_i,Activity_id,Cat_ID,cheque_no,Com_id, opening,opening_USD, 0, 1 from RPT_Bank_Book order by AC_CODE,Date_work,substring(certify,2,10),cnt  " & _
  "  delete RPT_Bank_Book where lck=0"
        CNN.Execute(aa)

        aa = "  update	 RPT_Bank_Book set opening=0 where opening is null  "
        CNN.Execute(aa)
        aa = "  update	 RPT_Bank_Book set opening_USD=0 where opening_USD is null  "
        CNN.Execute(aa)
        aa = "  update	 RPT_Bank_Book set Rem_USD=0 where Rem_USD is null  "
        CNN.Execute(aa)

        aa = "   delete RPT_Bank_Book_Rem  "
        CNN.Execute(aa)
        aa = "     insert into RPT_Bank_Book_Rem (Ac_Code , cnt_Mat ,rem,rem_USD )  " & _
"   select Ac_Code , cnt,opening+(select SUM(Amt_Dr-Amt_cr) from RPT_Bank_Book  where Ac_Code = x.ac_code    And cnt <= x.cnt )as Rem  ,opening_USD+(select SUM(Amt_USD_Dr-Amt_USD_cr) " & _
"   from RPT_Bank_Book  where Ac_Code = x.ac_code    And cnt <= x.cnt )as Rem_USD  from RPT_Bank_Book as x  order by  cnt Asc    "
        CNN.Execute(aa)
        aa = "   Update RPT_Bank_Book set RPT_Bank_Book.rem = RPT_Bank_Book_Rem.rem, RPT_Bank_Book.rem_USD = RPT_Bank_Book_Rem.rem_USD from RPT_Bank_Book , RPT_Bank_Book_Rem where  RPT_Bank_Book.cnt = RPT_Bank_Book_Rem.cnt_Mat    " & _
  "       update	 RPT_Bank_Book set rem=0 where rem is null  "
        CNN.Execute(aa)
        CNN.CommandTimeout = 0

        aa = " delete  RPT_Bank_Book_Group  from RPT_Bank_Book_Group as a ,RPT_Bank_Book as b where a.ac_code =b.Ac_code "
        CNN.Execute(aa)
        aa = " insert into  RPT_Bank_Book (Ac_code,opening,amt_dr,amt_cr,rem,opening_USD,rem_USD)   " & _
            "   SELECT ac_code,opening,0,0,opening,opening_USD,opening_USD  from  RPT_Bank_Book_Group   "
        CNN.Execute(aa)
        aa = "   update RPT_Bank_Book set ac_nm =  ACC_CODE.Name_L from ACC_CODE where RPT_Bank_Book.Ac_code= ACC_CODE.AC_CODE"
        CNN.Execute(aa)


        '======== ອັບເດດເລກບັນຊີ ================
        aa = "      update  RPT_Bank_Book set RPT_Bank_Book.Ac_code_dr = gen_jn.code_dr from RPT_Bank_Book,gen_jn " & _
        "  where RPT_Bank_Book.certify = gen_jn.certify and gen_jn.code_dr <>'' "
        CNN.Execute(aa)
        CNN.CommandTimeout = 0
        aa = "   update  RPT_Bank_Book set RPT_Bank_Book.Ac_code_cr = gen_jn.code_cr from RPT_Bank_Book,gen_jn " & _
      "  where RPT_Bank_Book.certify = gen_jn.certify and gen_jn.code_cr <>'' "
        CNN.Execute(aa)
        CNN.Execute(" update	 RPT_Bank_Book set amt_USD_dr=0   where amt_USD_dr is null  ")
        CNN.Execute(" update	 RPT_Bank_Book set amt_USD_cr=0  where amt_USD_cr is null ")

    End Sub
    Private Sub LoadFG_AC_CODE()
        'MdStartDate = Format(CDate(Ds.Value), "dd-MM-yyyy")
        'MdToDate = Format(CDate(Dt.Value), "dd-MM-yyyy")

        CNN.Execute("DELETE  Ap_PostedLedgers ")
        CNN.Execute("DELETE  Ap_Ope_PostedLedgers_Group ")
        CNN.Execute("DELETE FROM Ap_Open_PostedLedgers ")
        Dim S As Date
        S = MdStartDate
        S = DateAdd("d", CDbl(-1), MdStartDate)
        MuTable = "gen_jn." : Call Find_Company()
        Dim CUST_Supp As String

        Dim ACCN As String
        If Cfrom.Text = "" Then
            ACCN = ""
        Else
            'ACCN = " and ac_code=N'" & TxtAccCode.Text & "'  "
            ACCN = "and left(ac_code,'" & Len(Cfrom.Text) & "')>='" & Cfrom.Text & "' and  left(ac_code,'" & Len(Cfrom.Text) & "')<='" & Cfrom.Text & "' "

        End If

        Dim Cur As String = ""
        Cur = " AND Curr=N'" & CMB_Curr.Text & "' "

        Dim s11 As String = "INSERT INTO Ap_Open_PostedLedgers( ac_code , Amount_Dr , Amount_Cr , Amt_Dr , Amt_Cr , Curr   ) " & _
         "select Ac_Code , sum(amount_dr) , sum(amount_cr) ,  sum(amt_dr) , sum(amt_cr) , Curr  from gen_jn WHERE  1=1   " & Cur & "  " & ACCN & "  and gen_jn.date_work   BETWEEN '" & "" & Format(MdStartDate, "yyyy") & "-1-1' AND '" & Format(S, "yyyy-MM-dd") & "'   group BY Ac_Code , Curr "
        CNN.Execute(s11)
        Dim aa As String
        MuTable = "Open_jn." : Call Find_Company()
        aa = "INSERT INTO Ap_Open_PostedLedgers( ac_code , Amount_Dr , Amount_Cr , Amt_Dr , Amt_Cr  , Curr  ) " & _
         "select ac_code , amount_dr , amount_cr , amt_dr , amt_cr , Curr from Open_jn  WHERE  1=1  " & Cur & "  " & ACCN & "   and date_work='" & "" & Format(MdStartDate, "yyyy") & "-1-1'    "
        CNN.Execute(aa)
        MuTable = "gen_jn." : Call Find_Company()
        aa = " INSERT INTO Ap_PostedLedgers( ac_code , ac_name , date_work , descrip , Open_amount , open_amt , amount_dr , amount_cr , amt_dr , amt_cr , certify , remain , Status ,  curr ) " & _
               "select ac_code ,'' ,date_work, descrip ,0 ,0 ,amount_dr ,amount_cr , amt_dr , amt_cr ,certify, 0 , 0 ,  Curr  from gen_jn    WHERE   1=1  " & Cur & "  " & ACCN & "   and gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'   "
        CNN.Execute(aa)
        CNN.Execute("UPDATE Ap_PostedLedgers set ac_code='" & Cfrom.Text & "' ")
        CNN.Execute("UPDATE Ap_Open_PostedLedgers set ac_code='" & Cfrom.Text & "' ")

        CNN.Execute("Delete Ap_Ope_PostedLedgers_Group ")
        CNN.Execute("Insert Into Ap_Ope_PostedLedgers_Group (Ac_Code  ,Amount, Amt  )" & _
                     " Select Ac_Code ,Sum(Amount_Dr - Amount_Cr)  ,Sum(Amt_Dr - Amt_Cr)  from Ap_Open_PostedLedgers group by Ac_Code ")
        CNN.Execute("Update Ap_PostedLedgers set  Ap_PostedLedgers.Open_Amt = Ap_Ope_PostedLedgers_Group.Amt  , Ap_PostedLedgers.Open_Amount = Ap_Ope_PostedLedgers_Group.Amount  From Ap_PostedLedgers , Ap_Ope_PostedLedgers_Group  Where Ap_PostedLedgers.Ac_Code = Ap_Ope_PostedLedgers_Group.Ac_Code   ")

        CNN.Execute("Delete Ap_Ope_PostedLedgers_Group ")
        CNN.Execute("Insert Into Ap_Ope_PostedLedgers_Group (Ac_Code , Amount  , Amt )" & _
                  " Select Ac_Code ,Sum(Amount_Dr - Amount_Cr)   ,Sum(Amt_Dr - Amt_Cr)   from Ap_Open_PostedLedgers group by Ac_Code ")
        CNN.Execute("Update Ap_PostedLedgers set Ap_PostedLedgers.Open_Amount = Ap_Ope_PostedLedgers_Group.Amount , Ap_PostedLedgers.Open_Amt = Ap_Ope_PostedLedgers_Group.Amt   From Ap_PostedLedgers , Ap_Ope_PostedLedgers_Group  Where Ap_PostedLedgers.Ac_Code = Ap_Ope_PostedLedgers_Group.Ac_Code  ")

        CNN.Execute(" delete Ap_PostedLedgers_Rem ")
        CNN.Execute(" insert into Ap_PostedLedgers_Rem  (ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work)  " & _
         " select ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work from Ap_PostedLedgers order by  Ac_Code,Date_Work , certify ,cnt Asc  ")
        CNN.Execute("delete Ap_PostedLedgers   ")


        Dim KK As String = "insert into Ap_PostedLedgers  (ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work)  " & _
   " select ac_code, '', '', 'Opening Balance', 0, 0, 0, 0, 0, 0, 0, 0, Open_amount, '', '', '','" & Format(MdStartDate, "yyyy-MM-dd") & "' from Ap_PostedLedgers_Rem group by ac_code,Open_amount  "
        CNN.Execute(KK)
        CNN.Execute("UPDATE Ap_PostedLedgers set ac_code='" & Cfrom.Text & "' ")

        CNN.Execute("insert into Ap_PostedLedgers  (ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work)  " & _
         " select ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work from Ap_PostedLedgers_Rem order by cnt   ")

        'CNN.Execute("UPDATE Ap_PostedLedgers set ac_code='" & TxtAccCode.Text & "' ")

        CNN.Execute(" delete Ap_PostedLedgers_Rem ")
        CNN.Execute(" insert into Ap_PostedLedgers_Rem (Ac_Code , cnt_Mat ,remain )  " & _
          " select Ac_Code , cnt,Open_amt+(select SUM(Amt_Dr-Amt_cr) from Ap_PostedLedgers  where Ac_Code = x.ac_code    And cnt <= x.cnt )as Rem from Ap_PostedLedgers as x  order by  cnt Asc  ")


        CNN.Execute("Update Ap_PostedLedgers set Ap_PostedLedgers.remain = Ap_PostedLedgers_Rem.remain from Ap_PostedLedgers , Ap_PostedLedgers_Rem where  Ap_PostedLedgers.cnt = Ap_PostedLedgers_Rem.cnt_Mat  " & _
            "  update	 Ap_PostedLedgers set remain=0 where remain is null ")

        CNN.Execute(" Update Ap_PostedLedgers set Open_dr = 0 where  Open_dr is Null Update Ap_PostedLedgers set Open_cr = 0 where  Open_cr is Null Update Ap_PostedLedgers set Open_amount = 0 where  Open_amount is Null Update Ap_PostedLedgers set open_amt = 0 where  open_amt is Null Update Ap_PostedLedgers set amount_dr = 0 where  amount_dr is Null Update Ap_PostedLedgers set amount_cr = 0 where  amount_cr is Null Update Ap_PostedLedgers set amt_dr = 0 where  amt_dr is Null Update Ap_PostedLedgers set amt_cr = 0 where  amt_cr is Null Update Ap_PostedLedgers set remain = 0 where  remain is Null ")

        'If CheckBox1.Checked = True Then
        '    CNN.Execute(" delete Ap_PostedLedgers_Rem ")
        '    CNN.Execute(" insert into Ap_PostedLedgers_Rem (Ac_Code , cnt_Mat ,remain )  " & _
        '      " select Ac_Code , cnt,Open_amount+(select SUM(Amount_Dr-Amount_cr) from Ap_PostedLedgers  where Ac_Code = x.ac_code    And cnt <= x.cnt )as Rem_Curr from Ap_PostedLedgers as x  order by  cnt Asc  ")
        '    CNN.Execute("Update Ap_PostedLedgers set Ap_PostedLedgers.Rem_Curr = Ap_PostedLedgers_Rem.remain from Ap_PostedLedgers , Ap_PostedLedgers_Rem where  Ap_PostedLedgers.cnt = Ap_PostedLedgers_Rem.cnt_Mat  " & _
        '        "  update	 Ap_PostedLedgers set remain=0 where remain is null ")
        'End If


        MuTable = "Open_jn." : Call Find_Company()
        CNN.Execute("update Open_jn set Lck=0 where year(date_work)= " & Format(MdStartDate, "yyyy") & "  ")
        CNN.Execute("update Open_jn set Lck=1 from Open_jn, Ap_PostedLedgers   where year(Open_jn.date_work)= " & Format(MdStartDate, "yyyy") & " And  Open_jn.ac_code =  Ap_PostedLedgers.ac_code ")
        Dim s7 As String = " Insert into Ap_PostedLedgers (ac_code,Open_dr,Open_Cr,Open_amount,open_amt,amount_dr,amount_cr,amt_dr,amt_cr,Curr,remain,Rem_Curr,Status)  " & _
                    "select ac_code,0,0, Sum(amount_dr-amount_cr),Sum(amt_dr-amt_cr),0,0,0,0,Curr,0,0,0 from Open_jn where  1=1  " & Cur & "  and ac_code='" & Cfrom.Text & "' and Lck=0 And year(date_work)= " & Format(MdStartDate, "yyyy") & "  Group by ac_code , Curr"
        CNN.Execute(s7)
        CNN.Execute("UPDATE Ap_PostedLedgers set ac_code='" & Cfrom.Text & "' ")
        CNN.Execute("Update Ap_PostedLedgers set open_amt = Ap_Ope_PostedLedgers_Group.Amt from Ap_PostedLedgers ,Ap_Ope_PostedLedgers_Group  where  Ap_PostedLedgers.Ac_Code = Ap_Ope_PostedLedgers_Group.Ac_Code")

        'CNN.Execute("update Ap_PostedLedgers set Ap_PostedLedgers.ac_name=Acc_Code.Name_L from Ap_PostedLedgers , Acc_Code where Ap_PostedLedgers.Ac_Code = Acc_Code.Ac_Code")
        CNN.Execute("update Ap_PostedLedgers set Ap_PostedLedgers.ac_name=gen_jn.ac_name from Ap_PostedLedgers , gen_jn where Ap_PostedLedgers.Ac_Code = gen_jn.Ac_Code and   Ap_PostedLedgers.certify =gen_jn.certify ")

        CNN.Execute("Update Ap_PostedLedgers set remain = open_amt where Descrip=N'Opening Balance' ")

        'LoadListFG_AC_Code()
        'Call loadColor()
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        If CMB_Curr.SelectedIndex = 0 Then
            MsgBox("ກະລຸນາເລືອກສະກຸນເງິນກ່ອນ", MsgBoxStyle.Exclamation) : CMB_Curr.Focus() : Exit Sub
        End If

        LoadFG_AC_CODE()
        Call Office()

        If CMB_Curr.Text = "LAK" Then
            ACCNO = "00-" & Cfrom.Text
        ElseIf CMB_Curr.Text = "USD" Then
            ACCNO = "01-" & Cfrom.Text
        End If
        Curr = CMB_Curr.Text
        Dim rs As New ADODB.Recordset
        Dim sa As String = " SELECT    *  from Acc_Code   WHERE 1=1 AND  AC_CODE=N'" & Cfrom.Text & "' "
        Call LoadSqlData(sa, rs)
        If rs.RecordCount <> 0 Then
            'TxtAccCode.Text = (.Fields("AC_CODE").Value.ToString)
            AccNm = (rs.Fields("Name_L").Value.ToString) & " / " & (rs.Fields("Name_E").Value.ToString)

        End If

        'AccNm = TxtAccName.Text
        Call LoadLoGO()
        Dim AAA As String = " For the Period " & Ds.Text & " - " & Dt.Text
        SLF = "SELECT N'" & TxtS1.Text & "'  as TxtS1,N'" & TxtS2.Text & "'  as TxtS2,N'" & TxtS3.Text & "'  as TxtS3,N'" & TxtS4.Text & "'  as TxtS4,N'" & TxtPP.Text & "'  as TxtPP, N'" & ACCNO & "'  as ACCN  , N'" & Curr & "'  as Curr  , N'" & AccNm & "'  as AccNm  ,  N'" & AAA & "'  as DD  , N'" & MDSgn1 & "' as S1,N'" & MDSgn2 & "' as S2,N'" & MDSgn3 & "' as S3, N'" & MDSgn4 & "' as S4,  N'" & MDSgn5 & "' as S5,   N'" & MDSgn6 & "' as S6,  N'" & RptPro & "' as pp, " & mformat & "  as mformat  ,    *   FROM Ap_PostedLedgers Order by CNT asc "
        'Dim Rs As New ADODB.Recordset
        Dim RSN As New ADODB.Recordset
        With RSN
            If .State = ConnectionState.Open Then .Close()
            .Open(SLF, CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim Rpt As New Object
        Rpt = New Acc_Statement_Ac_Code
        If MdShowLOGO = 1 Then
            Rpt.Subreports(0).SetDataSource(RsLOGO)
        End If

        Rpt.SetDataSource(RSN)
        Rpt.Refresh()
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.DisplayGroupTree = False
        'FrmPreview.MdiParent = FmMain
        FrmPreview.WindowState = FormWindowState.Maximized
        FrmPreview.Show()
        FrmPreview.Focus()
    End Sub
End Class