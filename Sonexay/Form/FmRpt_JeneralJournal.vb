Public Class FmRpt_JeneralJournal
    Dim CLT_Str As String
    Dim bls1 As String
    Dim MonthLetter1 As String
    Dim MdStartDate As Date
    Dim MdToDate As Date
    Dim MdQuarter As Date

    Dim MdStartDate_Last As Date
    Dim MdToDate_Last As Date
    Dim ny, ly, n_L_y As String
    Dim BK As String = ""
    Dim sql As String
    Dim VCode1, VCode2, VCode3, VCode4, VCode5, VCode6, VCode7, VCode8, VCode9 As String
    Dim RsOpen As New ADODB.Recordset
    Dim RsOpenMonth As New ADODB.Recordset
    Dim RsRpt As New ADODB.Recordset
    Dim AmtOpenDR, AmtOpenCR, AmtOpenMonthDR, AmtOpenMonthCR As Double
    Dim VOpenDate As Date
    Dim RptNme As String
    Dim RSC12 As New ADODB.Recordset
    Dim D, P As String
    Dim RSCIn_M As New ADODB.Recordset

    'select sub_id , off_add2  from  Ap_office  Order by cnt

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

    Private Sub HeaDer()
        LoadSqlData("SELECT * FROM Header WHERE ID=N'J01' ", RSC)
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
        LoadSqlData("SELECT * FROM Header WHERE ID=N'J01' ", RSC)
        If RSC.RecordCount = 0 Then
            CNN.Execute("INSERT INTO Header(ID,Nm,S1,S2,S3,S4,PP) " & _
                        " values('J01',N'" & TxtHeader.Text & "',N'" & TxtS1.Text & "',N'" & TxtS2.Text & "',N'" & TxtS3.Text & "',N'" & TxtS4.Text & "',N'" & TxtPP.Text & "') ")
        Else
            CNN.Execute("UPDATE Header set Nm=N'" & TxtHeader.Text & "',S1=N'" & TxtS1.Text & "',S2=N'" & TxtS2.Text & "',S3=N'" & TxtS3.Text & "',S4=N'" & TxtS4.Text & "',PP=N'" & TxtPP.Text & "' " & _
                        " where ID='J01' ")
        End If
    End Sub



    Private Sub FmRpt_JeneralJournal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        HeaDer()
        RM.Checked = True
        Ds.Text = MWorkSetting
        Myy.Text = MWorkSetting
        yy.Text = MWorkSetting
        Toyy.Text = MWorkSetting
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
        Call SelectLoad()
        SetControlText(Me)
        Call loadOffice_User()
        CheckBox2.Text = "ໃບຢັ້ງຢືນ"
        Label15.Text = "ເຖີງ"
        Button4.Text = "Export"
        cmbBook.Items.Clear()
        cmbBook.Items.Add("All")
        Call load_Cmb("select bookid , bookname  from books   order by bookid   ", "bookid", cmbBook)
        If cmbBook.Items.Count > 0 Then
            cmbBook.SelectedIndex = 0
        End If
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


        If MuLng = "L" Then

            Label8.Text = "ລາຍເຊັນ1"
            Label14.Text = "ລາຍເຊັນ2"
            Label7.Text = "ລາຍເຊັນ3"
            Label12.Text = "ລາຍເຊັນ4"
            Label11.Text = "ທີ່"

            RadioButton8.Text = "ໃບຢັ້ງຢືນ"
            CheckBox2.Text = "ໃບຢັ້ງຢືນ"
            Label15.Text = "ເຖີງ"
            RadioButton7.Text = "ປື້ມບັນຊີ"
        Else
            RadioButton8.Text = "Certify"
            RadioButton7.Text = "Books"
            Label8.Text = "Signature1"
            Label14.Text = "Signature2"
            Label7.Text = "Signature3"
            Label12.Text = "Signature4"
            Label11.Text = "Location"
            CheckBox2.Text = "Certify"
            Label15.Text = "To"

        End If
    End Sub
    Private Sub LoadDay()
        MdStartDate = Format(CDate(Ds.Value), "dd-MM-yyyy")
        MdToDate = Format(CDate(Dt.Value), "dd-MM-yyyy")
        'Lb.Text = "ແຕ່ວັນທີ " & MdStartDate & " ຫາວັນທີ " & MdToDate
        ' Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາວັນທີ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳວັນທີ"
        L5.Text = MdStartDate & " => " & MdToDate
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

    Private Sub LoadYear()
        MdStartDate = Format(CDate("01/1/" & Year(yy.Value)), "dd-MM-yyyy")
        MdToDate = Format(CDate("31/12/" & Year(Toyy.Value)), "dd-MM-yyyy")
        Lb.Text = "ປະຈຳປີ " & yy.Text
        'Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd/MM/yyyy") & " ສີ້ນສຸດວັນທີ " & Format(MdToDate, "dd/MM/yyyy")
        'Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ສີ້ນສຸດວັນທີ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳປີ  " & Year(MdToDate)
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub

    Private Sub BtnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        AddHeader()
        If RD.Checked = True Then
            Dim s1, s2 As String
            LngId = 7027 : CallLngStr() : s2 = LngStr
            LngId = 7072 : CallLngStr() : s1 = LngStr
            'Lb.Text = s1 & " " & Ds.Text & " " & s2 & " " & Dt.Text
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
            'Lb.Text = s1 & " (" & s2 & ") " & s3 & " " & Year(MdToDate)
        ElseIf RY.Checked = True Then
            Dim s1, s2, s3 As String
            LngId = 7074 + Period.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7062 : CallLngStr() : s1 = LngStr
            LngId = 7064 : CallLngStr() : s3 = LngStr
            'Lb.Text = s3 & " " & yy.Text
        End If


        Call Office()
        Dim OfUsr1 As String = Microsoft.VisualBasic.Left(Off_Usr.Text, 5)
        Dim OfUsr2 As String = Mid(Off_Usr.Text, 4, 2)
        Dim OfUsr3 As String = Microsoft.VisualBasic.Left(Off_Usr.Text, 2)
        If OfUsr1 = "00-00" Then
            MULook2 = ""
        Else
            If OfUsr2 = "00" Then
                MULook2 = "  And  Left(gen_jn.company,2)= '" & OfUsr3 & "' "
            Else
                MULook2 = "  And gen_jn.company= '" & OfUsr1 & "' "
            End If
        End If

        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As    RptSjUd ,"

        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
        LngId = "7002" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDP ,"
        LngId = "7003" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & " " & DTDATE02 & "' As Crl_RptName ,"
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

        LngId = "7127" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Cheque ,"
        LngId = "7128" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Exchane ,"
        LngId = "7129" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Rate ,"

 

        LngId = "7042" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_OpenAmt ,"
        LngId = "7043" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RemAmt ,"
        'Dim Book As String
        'If CheckBox2.Checked = True Then
        '    'Book = " and gen_jn.book=N'" & Trim(cmbBook.Text) & "' "
        '    Book = "  and  left(gen_jn.certify,'" & Len(txtbill_no1.Text) & "')>='" & txtbill_no1.Text & "' and  left(gen_jn.certify,'" & Len(txtbill_no2.Text) & "')<='" & txtbill_no2.Text & "'   "
        'Else
        '    'Book = "  and  left(gen_jn.certify,'" & Len(txtbill_no1.Text) & "')>='" & txtbill_no1.Text & "' and  left(gen_jn.certify,'" & Len(txtbill_no2.Text) & "')<='" & txtbill_no2.Text & "'   "
        '    Book = ""
        'End If

        Dim Book As String
        If RadioButton7.Checked = True Then
            If cmbBook.SelectedIndex = 0 Then
                Book = ""
            Else
                Book = " and gen_jn.book=N'" & Trim(cmbBook.Text) & "' "
            End If
        Else
            Book = "  and  left(gen_jn.certify,'" & Len(txtbill_no1.Text) & "')>='" & txtbill_no1.Text & "' and  left(gen_jn.certify,'" & Len(txtbill_no2.Text) & "')<='" & txtbill_no2.Text & "'   "
        End If

        Dim mformat As String = 0

        ' CNN.Execute("update gen_jn set gen_jn.ac_name=Acc_Code.name_L,gen_jn.ac_namee=Acc_Code.name_e from Acc_Code,gen_jn where gen_jn.ac_code=gen_jn.ac_code and gen_jn.ac_name is null ")
        SLF = "SELECT  " & mformat & "  as mformat  , " & MuLngRpt & "   *  ,  gen_jn.ac_name  As AcNmeEx_L , gen_jn.ac_name  As AcNmeEx_E FROM  gen_jn      WHERE Book <>'' And gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  "



        Call LoadLoGO()

        'SLF = "SELECT N'" & Lb.Text & "' As    RptSjUd ," & RptSjOff & "   *  ,Acc_Code.Name_L AS ExAcc_Name FROM  gen_jn       INNER JOIN Acc_Code ON gen_jn.ac_code = Acc_Code.Ac_Code WHERE Book <>'' "
        If CheckBox1.Checked = False Then
            Dim Rs As New ADODB.Recordset
            With Rs
                'gen_jn.date_work,gen_jn.referno ASC
                If .State = ConnectionState.Open Then .Close()
                .Open(" " & SLF & "  " & MULook2 & "  " & Book & "  ORDER BY gen_jn.cnt ASC  ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
                If .EOF Then Exit Sub
            End With
            Dim FrmPreview As New FmPreview : FrmClosing()
            'Dim Rpt As New CryGeneralLedgers
            Dim Rpt As New CrystalReport_General_Jurnal_Curr_List_P
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
            myText2.Text = "ຫົວໜ່ວຍ : ກີບ"

            Rpt.SetDataSource(Rs)
            FrmPreview.ReportViewer.ReportSource = Rpt
            FrmPreview.ReportViewer.DisplayGroupTree = False
            'FrmPreview.MdiParent = FmMain
            FrmPreview.WindowState = FormWindowState.Maximized
            FrmPreview.ShowDialog()
            FrmPreview.Focus()
        Else
            Dim Rs As New ADODB.Recordset
            With Rs
                If .State = ConnectionState.Open Then .Close()
                .Open("  " & SLF & "  " & MULook2 & "  " & Book & "  " & "" & " ORDER BY  gen_jn.date_work , dbo.gen_jn.certify , dbo.gen_jn.cnt ASC ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
                If .EOF Then Exit Sub
            End With
            Dim FrmPreview As New FmPreview : FrmClosing()
            'Dim Rpt As New CryGeneralLedgersUser
            Dim Rpt As New CrystalReport_General_Jurnal_Curr_List_P_Curr
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
            myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("txtprint_user"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = MUserName
            myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("Text3"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = "ຫົວໜ່ວຍ : ກີບ"
            Rpt.SetDataSource(Rs)
            FrmPreview.ReportViewer.ReportSource = Rpt
            FrmPreview.ReportViewer.DisplayGroupTree = False
            'FrmPreview.MdiParent = FmMain
            FrmPreview.WindowState = FormWindowState.Maximized
            FrmPreview.ShowDialog()
            FrmPreview.Focus()
        End If
    End Sub

    Private Sub RD_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RD.CheckedChanged
        Call SelectLoad()
    End Sub
    Private Sub SelectLoad()
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

    Private Sub Ds_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ds.ValueChanged
        Call SelectLoad()
    End Sub

    Private Sub Dt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Dt.ValueChanged
        Call SelectLoad()
    End Sub

    Private Sub RM_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RM.CheckedChanged
        Call SelectLoad()
    End Sub

    Private Sub DMonth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DMonth.SelectedIndexChanged
        Call SelectLoad()
    End Sub

    Private Sub Myy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Myy.ValueChanged
        Call SelectLoad()
    End Sub

    Private Sub Period_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Period.SelectedIndexChanged
        Call SelectLoad()
    End Sub

    Private Sub Pyy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Pyy.ValueChanged
        Call SelectLoad()
    End Sub

    Private Sub yy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yy.ValueChanged
        Call SelectLoad()
    End Sub

    Private Sub Toyy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Toyy.ValueChanged
        Call SelectLoad()
    End Sub

    Private Sub RY_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RY.CheckedChanged
        Call SelectLoad()
    End Sub

    Private Sub RP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RP.CheckedChanged
        Call SelectLoad()
    End Sub

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Close()
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            txtbill_no1.Enabled = True
            txtbill_no2.Enabled = True
        Else
            txtbill_no1.Enabled = False
            txtbill_no2.Enabled = False
        End If

    End Sub

    Private Sub txtbill_no1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtbill_no1.TextChanged
        txtbill_no2.Text = txtbill_no1.Text
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
            Lb.Text = s1 & " (" & s2 & ") " & s3 & " " & Year(MdToDate)
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


        Call Office()
        Dim OfUsr1 As String = Microsoft.VisualBasic.Left(Off_Usr.Text, 5)
        Dim OfUsr2 As String = Mid(Off_Usr.Text, 4, 2)
        Dim OfUsr3 As String = Microsoft.VisualBasic.Left(Off_Usr.Text, 2)
        If OfUsr1 = "00-00" Then
            MULook2 = ""
        Else
            If OfUsr2 = "00" Then
                MULook2 = "  And  Left(gen_jn.company,2)= '" & OfUsr3 & "' "
            Else
                MULook2 = "  And gen_jn.company= '" & OfUsr1 & "' "
            End If
        End If







        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & L5.Text & "' As    RptSjUd ,"

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

        LngId = "7042" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_OpenAmt ,"
        LngId = "7043" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RemAmt ,"
        'Dim Book As String
        'If CheckBox2.Checked = True Then
        '    'Book = " and gen_jn.book=N'" & Trim(cmbBook.Text) & "' "
        '    Book = "  and  left(gen_jn.certify,'" & Len(txtbill_no1.Text) & "')>='" & txtbill_no1.Text & "' and  left(gen_jn.certify,'" & Len(txtbill_no2.Text) & "')<='" & txtbill_no2.Text & "'   "
        'Else
        '    'Book = "  and  left(gen_jn.certify,'" & Len(txtbill_no1.Text) & "')>='" & txtbill_no1.Text & "' and  left(gen_jn.certify,'" & Len(txtbill_no2.Text) & "')<='" & txtbill_no2.Text & "'   "
        '    Book = ""
        'End If


        Dim Book As String
        If RadioButton7.Checked = True Then
            If cmbBook.SelectedIndex = 0 Then
                Book = ""
            Else
                Book = " and gen_jn.book=N'" & Trim(cmbBook.Text) & "' "
            End If
        Else
            Book = "  and  left(gen_jn.certify,'" & Len(txtbill_no1.Text) & "')>='" & txtbill_no1.Text & "' and  left(gen_jn.certify,'" & Len(txtbill_no2.Text) & "')<='" & txtbill_no2.Text & "'   "
        End If



        CNN.Execute("update gen_jn set gen_jn.ac_name=Acc_Code.name_L,gen_jn.ac_namee=Acc_Code.name_e from Acc_Code,gen_jn where gen_jn.ac_code=gen_jn.ac_code and gen_jn.ac_name is null ")
        'SLF = "SELECT " & MuLngRpt & "   *  ,  Acc_Code.Name_L  As AcNmeEx_L , Acc_Code.Name_E  As AcNmeEx_E FROM  gen_jn       INNER JOIN Acc_Code ON gen_jn.ac_code = Acc_Code.Ac_Code WHERE Book <>'' And gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  "
        SLF = "SELECT " & MuLngRpt & "   *  ,  gen_jn.ac_name  As AcNmeEx_L , gen_jn.ac_name  As AcNmeEx_E FROM  gen_jn      WHERE Book <>'' And gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  "

        'MsgBox(SLF)

        Call LoadLoGO()

        'SLF = "SELECT N'" & Lb.Text & "' As    RptSjUd ," & RptSjOff & "   *  ,Acc_Code.Name_L AS ExAcc_Name FROM  gen_jn       INNER JOIN Acc_Code ON gen_jn.ac_code = Acc_Code.Ac_Code WHERE Book <>'' "
        If CheckBox1.Checked = False Then
            Dim Rs As New ADODB.Recordset
            With Rs
                If .State = ConnectionState.Open Then .Close()
                .Open(" " & SLF & "  " & MULook2 & "  " & Book & "  order by  gen_jn.cnt ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
                If .EOF Then Exit Sub
            End With
            Dim FrmPreview As New FmPreview : FrmClosing()
            'Dim Rpt As New CryGeneralLedgersUser
            Dim Rpt As New CrystalReport_General_Jurnal_Curr_List_P
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
            FrmPreview.ReportViewer.ExportReport()
            FrmPreview = Nothing
            'FrmPreview.ReportViewer.DisplayGroupTree = False
            ''FrmPreview.MdiParent = FmMain
            'FrmPreview.WindowState = FormWindowState.Maximized
            'FrmPreview.ShowDialog()
            'FrmPreview.Focus()
        Else
            Dim Rs As New ADODB.Recordset
            With Rs
                If .State = ConnectionState.Open Then .Close()
                .Open("  " & SLF & "  " & MULook2 & "  " & Book & "  " & "" & " ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
                If .EOF Then Exit Sub
            End With
            Dim FrmPreview As New FmPreview : FrmClosing()
            'Dim Rpt As New CryGeneralLedgersUser
            Dim Rpt As New CrystalReport_General_Jurnal_Curr_List_P_Curr
            'If MdShowLOGO = 1 Then
            '    Rpt.Subreports(0).SetDataSource(RsLOGO)
            'End If
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
            FrmPreview.ReportViewer.ExportReport()
            FrmPreview = Nothing
            'FrmPreview.ReportViewer.DisplayGroupTree = False
            ''FrmPreview.MdiParent = FmMain
            'FrmPreview.WindowState = FormWindowState.Maximized
            'FrmPreview.ShowDialog()
            'FrmPreview.Focus()
        End If
    End Sub

    Private Sub cmbBook_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbBook.SelectedIndexChanged
        Dim rs As New ADODB.Recordset
        Call LoadSqlData("Select * From books Where   bookid=N'" & Trim(cmbBook.Text) & "'", rs)
        If rs.RecordCount > 0 Then
            txtbook.Text = Trim(rs("bookname").Value.ToString)
        End If

        If cmbBook.SelectedIndex = 0 Then
            txtbook.Text = ""
            BK = ""
        Else
            BK = " and books.bookid=N'" & cmbBook.Text & "'  "
        End If
    End Sub

    Private Sub RadioButton7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton7.CheckedChanged
        If RadioButton7.Checked = True Then
            cmbBook.Enabled = True
            txtbill_no1.Enabled = False
            txtbill_no2.Enabled = False
        Else
            cmbBook.Enabled = False
            txtbill_no1.Enabled = True
            txtbill_no2.Enabled = True
        End If
    End Sub

    Private Sub RadioButton8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton8.CheckedChanged

    End Sub
End Class