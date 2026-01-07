Public Class FrmRpt_F04
    Dim Biz_Type, Loan_Type, sqlNew, Rpt_ID As String

    Dim RelationShip, Relation As String
    Private Sub LoadHeader()
        Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM Header where ID='F04'")
        If dt.Rows.Count <> 0 Then
            txtHeader.Text = DbHelper.GetStr(dt.Rows(0)("Nm"))
            txtSig1.Text = DbHelper.GetStr(dt.Rows(0)("S1"))
            txtSig2.Text = DbHelper.GetStr(dt.Rows(0)("S2"))
            txtSig3.Text = DbHelper.GetStr(dt.Rows(0)("S3"))
            txtSig4.Text = DbHelper.GetStr(dt.Rows(0)("S4"))
            txtSig5.Text = DbHelper.GetStr(dt.Rows(0)("S5"))
            txtSig6.Text = DbHelper.GetStr(dt.Rows(0)("S6"))
            TxtPP.Text = DbHelper.GetStr(dt.Rows(0)("PP"))
        End If

    End Sub
    Private Sub Frm_loan_Due_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MMM()
        LoadHeader()
        MdStartDate = Date.Now
        DtToDate.Value = Date.Now

        Call SetDatNow()

        DtFromDate.Value = MWorkSetting
        DtToDate.Value = MWorkSetting

        DtYearFormonth.Value = MWorkSetting
        DtYearForPeriod.Value = MWorkSetting

        DtYearforFirsthalfyear.Value = MWorkSetting
        DtAllYear.Value = MWorkSetting


        Call SelectDate()

        RMonth.Checked = True
        'loadOffice_User()
        ComboBox1.SelectedIndex = 0
        Button4.Text = "Export"
    End Sub
    Private Sub SetDatNow()
        DtFromDate.Value = MdStartDate
        DtToDate.Value = MdStartDate
        CFromMonth.SelectedIndex = CDbl(Month(MdStartDate)) - 1
        CToMonth.SelectedIndex = CDbl(Month(MdStartDate)) - 1
        DtYearFormonth.Value = MdStartDate
        If CDbl(Month(MdStartDate)) < 4 Then
            CPeriod.SelectedIndex = 0
        ElseIf CDbl(Month(MdStartDate)) > 3 Then
            CPeriod.SelectedIndex = 1
        ElseIf CDbl(Month(MdStartDate)) > 6 Then
            CPeriod.SelectedIndex = 2
        ElseIf CDbl(Month(MdStartDate)) > 9 Then
            CPeriod.SelectedIndex = 3
        End If
        DtYearForPeriod.Value = MdStartDate
        If CDbl(Month(MdStartDate)) < 7 Then
            CFirsthalfyear.SelectedIndex = 0
        Else
            CFirsthalfyear.SelectedIndex = 1
        End If
        DtYearforFirsthalfyear.Value = MdStartDate
        DtAllYear.Value = MdStartDate

    End Sub
    Private Sub SelectDate()

        If RDate.Checked = True Then
            MdStartDate = DtFromDate.Value
            MdToDate = DtToDate.Value
            DtFromDate.Enabled = True
            DtToDate.Enabled = True


            CFromMonth.Enabled = False
            CToMonth.Enabled = False
            DtYearFormonth.Enabled = False
            CPeriod.Enabled = False
            DtYearForPeriod.Enabled = False
            CFirsthalfyear.Enabled = False
            DtYearforFirsthalfyear.Enabled = False
            DtAllYear.Enabled = False

            If DtFromDate.Text = DtToDate.Text Then
                RmFrdate.Text = "ປະຈໍາວັນທີ: " & Format(MdStartDate, "dd/MM/yyyy")
            Else
                RmFrdate.Text = "ແຕ່ວັນທີ: " & Format(MdStartDate, "dd/MM/yyyy") & " ຫາວັນທີ " & Format(MdToDate, "dd/MM/yyyy")
            End If
        ElseIf RMonth.Checked = True Then
            MdStartDate = Format(CDate("01/" & CDbl(CFromMonth.SelectedIndex + 1) & "/" & Year(DtYearFormonth.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("01/" & CDbl(CToMonth.SelectedIndex + 1) & "/" & Year(DtYearFormonth.Value)), "dd-MM-yyyy")
            Dim SM As Date = DateAdd("M", CDbl(1), MdToDate)
            Dim SD As Date = DateAdd("d", CDbl(-1), SM)
            MdToDate = Format(CDate(SD), "dd-MM-yyyy")
            CFromMonth.Enabled = True
            CToMonth.Enabled = True
            DtYearFormonth.Enabled = True
            DMY.Value = Format(MdStartDate, "dd/MM/yyyy")
            DtFromDate.Enabled = False
            DtToDate.Enabled = False
            CPeriod.Enabled = False
            DtYearForPeriod.Enabled = False
            CFirsthalfyear.Enabled = False
            DtYearforFirsthalfyear.Enabled = False
            DtAllYear.Enabled = False
            If CFromMonth.Text = CToMonth.Text Then
                RmFrdate.Text = "ປະຈໍາເດືອນ: " & CFromMonth.Text & " ປີ: " & Format(MdStartDate, "yyyy")
            Else
                RmFrdate.Text = "ແຕ່ເດືອນ: " & CFromMonth.Text & "/" & Format(MdStartDate, "yyyy") & "ຫາເດືອນ: " & CToMonth.Text & "/" & Format(MdStartDate, "yyyy")
            End If
        ElseIf RPeriod.Checked = True Then
            Dim dt As Date = Format(CDate("01/01/" & Year(DtYearForPeriod.Value)), "dd-MM-yyyy")
            MdStartDate = Format(CDate("1/" & CDbl(CDbl(CDbl(CDbl(CDbl(CPeriod.SelectedIndex) + 2) * 3)) - 5) & "/" & Year(DtYearForPeriod.Value)), "dd-MM-yyyy")
            Dim SM1 As Date = DateAdd(DateInterval.Month, CDbl(3), CDate(Format(CDate(MdStartDate), "yyyy-MM-dd")))
            MdToDate = DateAdd(DateInterval.Day, CDbl(-1), SM1)
            CPeriod.Enabled = True
            DtYearForPeriod.Enabled = True

            DtFromDate.Enabled = False
            DtToDate.Enabled = False
            CFromMonth.Enabled = False
            CToMonth.Enabled = False
            DtYearFormonth.Enabled = False

            CFirsthalfyear.Enabled = False
            DtYearforFirsthalfyear.Enabled = False
            DtAllYear.Enabled = False
            RmFrdate.Text = "ປະຈໍາງວດ: " & CPeriod.Text & " ປີ: " & Format(MdStartDate, "yyyy")
        ElseIf RFirsthalfyear.Checked = True Then
            If CFirsthalfyear.SelectedIndex = 0 Then
                MdStartDate = Format(CDate("1/1/" & Year(DtYearforFirsthalfyear.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/6/" & Year(DtYearforFirsthalfyear.Value)), "dd-MM-yyyy")
            Else
                MdStartDate = Format(CDate("1/7/" & Year(DtYearforFirsthalfyear.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/12/" & Year(DtYearforFirsthalfyear.Value)), "dd-MM-yyyy")
            End If
            CFirsthalfyear.Enabled = True
            DtYearforFirsthalfyear.Enabled = True

            DtFromDate.Enabled = False
            DtToDate.Enabled = False
            CFromMonth.Enabled = False
            CToMonth.Enabled = False
            DtYearFormonth.Enabled = False
            CPeriod.Enabled = False
            DtYearForPeriod.Enabled = False


            DtAllYear.Enabled = False
            RmFrdate.Text = CFirsthalfyear.Text & "/" & Format(MdStartDate, "yyyy")
        ElseIf RYearAll.Checked = True Then
            MdStartDate = Format(CDate("1/1/" & Year(DtAllYear.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/12/" & Year(DtAllYear.Value)), "dd-MM-yyyy")
            DtAllYear.Enabled = True

            DtFromDate.Enabled = False
            DtToDate.Enabled = False
            CFromMonth.Enabled = False
            CToMonth.Enabled = False
            DtYearFormonth.Enabled = False
            CPeriod.Enabled = False
            DtYearForPeriod.Enabled = False
            CFirsthalfyear.Enabled = False
            DtYearforFirsthalfyear.Enabled = False

            RmFrdate.Text = "ປະຈໍາປີ: " & Format(MdStartDate, "yyyy")
        End If
        Label12.Text = Format(MdStartDate, "yyyy/MM/dd") & " To " & Format(MdToDate, "yyyy/MM/dd")
    End Sub
    Private Sub CALLC()
        Dim MDWRITEOFF As String = " AND WRITEOFF=N'N' "

        '===========   N'1. ສິນເຊື່ອປົກກະຕິ (A)'===============
        '=========ACCNO=====
        Dim A1 As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan where 1=1 " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and  loan_grade=N'1. ສິນເຊື່ອປົກກະຕິ (A)' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(A1)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim K1 As String = "Update RPT_F04 set ACCNO=" & DbHelper.GetStr(row("AA")) & "  where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and Grp_Nm=N'1. ສິນເຊື່ອປົກກະຕິ (A)' "
                DbHelper.ExecuteNonQuery(K1)
            Next
        End If
        '=========ACC ALl=====
        Dim A1_Cust As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan where 1=1 " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and  loan_grade=N'1. ສິນເຊື່ອປົກກະຕິ (A)' group by BUSINESSTYPEDESC  , Cust_ID order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(A1_Cust)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim K1 As String = "Update RPT_F04 set  ACC=ACC+" & 1 & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and Grp_Nm=N'1. ສິນເຊື່ອປົກກະຕິ (A)' "
                DbHelper.ExecuteNonQuery(K1)
            Next
        End If
        ' ====ACC W===
        Dim W1 As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and  loan_grade=N'1. ສິນເຊື່ອປົກກະຕິ (A)' and GENDER=N'F' group by BUSINESSTYPEDESC, Cust_ID  order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(W1)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim KW1 As String = "Update RPT_F04 set  ACC_W=ACC_W+" & 1 & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and Grp_Nm=N'1. ສິນເຊື່ອປົກກະຕິ (A)' "
                DbHelper.ExecuteNonQuery(KW1)
            Next
        End If

        ' ====AMT KIP  TT===
        Dim AMT1 As String = "select isnull(sum(Principle_LAK),0) as AA, isnull(sum(Provision_Amt),0) as BB, BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'1. ສິນເຊື່ອປົກກະຕິ (A)' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(AMT1)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim KW1 As String = "Update RPT_F04 set  AMT=" & DbHelper.GetStr(row("AA")) & " , Dep_Amt=" & DbHelper.GetStr(row("BB")) & "  where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and Grp_Nm=N'1. ສິນເຊື່ອປົກກະຕິ (A)' "
                DbHelper.ExecuteNonQuery(KW1)
            Next
        End If
        ' ====AMT KIP  W===
        Dim AMT_W As String = "select isnull(sum(Principle_LAK),0) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'1. ສິນເຊື່ອປົກກະຕິ (A)' and GENDER=N'F' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(AMT_W)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim KW1 As String = "Update RPT_F04 set  AMT_W=" & DbHelper.GetStr(row("AA")) & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and Grp_Nm=N'1. ສິນເຊື່ອປົກກະຕິ (A)' "
                DbHelper.ExecuteNonQuery(KW1)
            Next
        End If

        ''=========== 2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)===============
        Dim A2 As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(A2)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim K2 As String = "Update RPT_F04 set ACCNO=" & DbHelper.GetStr(row("AA")) & "  where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and Grp_Nm=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' "
                DbHelper.ExecuteNonQuery(K2)
            Next
        End If
        '============
        Dim A2_AA As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' group by BUSINESSTYPEDESC, Cust_ID order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(A2_AA)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim K2 As String = "Update RPT_F04 set  ACC=ACC+" & 1 & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and Grp_Nm=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' "
                DbHelper.ExecuteNonQuery(K2)
            Next
        End If
        ' ====W===
        Dim W2 As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' and GENDER=N'F' group by BUSINESSTYPEDESC, Cust_ID order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(W2)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim KW2 As String = "Update RPT_F04 set  ACC_W=ACC_W+" & 1 & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and Grp_Nm=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' "
                DbHelper.ExecuteNonQuery(KW2)
            Next
        End If
        ' ====AMT KIP  TT===
        Dim AMT2 As String = "select isnull(sum(Principle_LAK),0) as AA,  isnull(sum(Provision_Amt),0) as BB, BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(AMT2)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim KW1 As String = "Update RPT_F04 set  AMT=" & DbHelper.GetStr(row("AA")) & " , Dep_Amt=" & DbHelper.GetStr(row("BB")) & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and Grp_Nm=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' "
                DbHelper.ExecuteNonQuery(KW1)
            Next
        End If
        ' ====AMT KIP  W===
        Dim AMT_W2 As String = "select isnull(sum(Principle_LAK),0) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' and GENDER=N'F' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(AMT_W2)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim KW1 As String = "Update RPT_F04 set  AMT_W=" & DbHelper.GetStr(row("AA")) & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and Grp_Nm=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' "
                DbHelper.ExecuteNonQuery(KW1)
            Next
        End If


        ''=========== 3. ສິນເຊື່ອຕໍ່າກວ່າມາດຕະຖານ (C)===============
        Dim A3 As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   left(loan_grade,1)=N'3' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(A3)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim K3 As String = "Update RPT_F04 set ACCNO=" & DbHelper.GetStr(row("AA")) & "  where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and left(Grp_Nm,1)=N'3' "
                DbHelper.ExecuteNonQuery(K3)
            Next
        End If
        '=======
        Dim A3_AAA As String = "select count(*) as AA,BUSINESSTYPEDESC,1 from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and    left(loan_grade,1)=N'3' group by BUSINESSTYPEDESC,Cust_ID  order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(A3_AAA)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim K3 As String = "Update RPT_F04 set  ACC=ACC+" & 1 & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and left(Grp_Nm,1)=N'3' "
                DbHelper.ExecuteNonQuery(K3)
            Next
        End If
        ' ====W===
        Dim W3 As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   left(loan_grade,1)=N'3' and GENDER=N'F' group by BUSINESSTYPEDESC,Cust_ID order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(W3)

        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim KW3 As String = "Update RPT_F04 set  ACC_W=ACC_W+" & 1 & "  where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and left(Grp_Nm,1)=N'3' "
                DbHelper.ExecuteNonQuery(KW3)
            Next
        End If
        ' ====AMT KIP  TT===
        Dim AMT3 As String = "select isnull(sum(Principle_LAK),0) as AA,  isnull(sum(Provision_Amt),0) as BB,  BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   left(loan_grade,1)=N'3' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(AMT3)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim KW1 As String = "Update RPT_F04 set  AMT=" & DbHelper.GetStr(row("AA")) & "  , Dep_Amt=" & DbHelper.GetStr(row("BB")) & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and left(Grp_Nm,1)=N'3' "
                DbHelper.ExecuteNonQuery(KW1)
            Next
        End If
        ' ====AMT KIP  W===
        Dim AMT_W3 As String = "select isnull(sum(Principle_LAK),0) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and    left(loan_grade,1)=N'3' and GENDER=N'F' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(AMT_W3)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim KW1 As String = "Update RPT_F04 set  AMT_W=" & DbHelper.GetStr(row("AA")) & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and left(Grp_Nm,1)=N'3'' "
                DbHelper.ExecuteNonQuery(KW1)
            Next
        End If

        ''=========== 4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)==============
        Dim A4 As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(A4)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim K4 As String = "Update RPT_F04 set ACCNO=" & DbHelper.GetStr(row("AA")) & "  where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and Grp_Nm=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)' "
                DbHelper.ExecuteNonQuery(K4)
            Next
        End If
        '=====
        Dim A4_AA As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)' group by BUSINESSTYPEDESC,Cust_ID order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(A4_AA)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim K4 As String = "Update RPT_F04 set  ACC=ACC+" & 1 & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and Grp_Nm=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)' "
                DbHelper.ExecuteNonQuery(K4)
            Next
        End If
        ' ====W===
        Dim W4 As String = "select count(*) as AA,BUSINESSTYPEDESC  from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)'   and GENDER=N'F' group by BUSINESSTYPEDESC,Cust_ID order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(W4)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim KW4 As String = "Update RPT_F04 set  ACC_W=ACC_W+" & 1 & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and Grp_Nm=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)'  "
                DbHelper.ExecuteNonQuery(KW4)
            Next
        End If
        ' ====AMT KIP  TT===
        Dim AMT4 As String = "select isnull(sum(Principle_LAK),0) as AA,  isnull(sum(Provision_Amt),0) as BB,  BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)'  group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(AMT4)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim KW1 As String = "Update RPT_F04 set  AMT=" & DbHelper.GetStr(row("AA")) & "  , Dep_Amt=" & DbHelper.GetStr(row("BB")) & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and Grp_Nm=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)'  "
                DbHelper.ExecuteNonQuery(KW1)
            Next
        End If
        ' ====AMT KIP  W===
        Dim AMT_W4 As String = "select isnull(sum(Principle_LAK),0) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)'  and GENDER=N'F' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(AMT_W4)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim KW1 As String = "Update RPT_F04 set  AMT_W=" & DbHelper.GetStr(row("AA")) & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and Grp_Nm=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)'  "
                DbHelper.ExecuteNonQuery(KW1)
            Next
        End If


        ''=========== 5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E))==============
        Dim A5 As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   (loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(A5)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim K5 As String = "Update RPT_F04 set ACCNO=" & DbHelper.GetStr(row("AA")) & "  where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and (Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') "
                DbHelper.ExecuteNonQuery(K5)
            Next
        End If

        Dim A5_AA As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   (loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') group by BUSINESSTYPEDESC,Cust_ID order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(A5_AA)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim K5 As String = "Update RPT_F04 set  ACC=ACC+" & 1 & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and (Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') "
                DbHelper.ExecuteNonQuery(K5)
            Next
        End If
        ' ====W===
        Dim W5 As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and  (loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)')   and GENDER=N'F' group by BUSINESSTYPEDESC,Cust_ID order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(W5)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim KW5 As String = "Update RPT_F04 set  ACC_W=ACC_W+" & 1 & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and   (Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') "
                DbHelper.ExecuteNonQuery(KW5)
            Next
        End If
        ' ====AMT KIP  TT===
        Dim AMT5 As String = "select isnull(sum(Principle_LAK),0) as AA,  isnull(sum(Provision_Amt),0) as BB,  BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   (loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(AMT5)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim KW1 As String = "Update RPT_F04 set  AMT=" & DbHelper.GetStr(row("AA")) & "  , Dep_Amt=" & DbHelper.GetStr(row("BB")) & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and (Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') "
                DbHelper.ExecuteNonQuery(KW1)
            Next
        End If
        ' ====AMT KIP  W===
        Dim AMT_W5 As String = "select isnull(sum(Principle_LAK),0) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and  (loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') and GENDER=N'F' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Dim dt As DataTable = DbHelper.GetDataTable(AMT_W5)
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim KW1 As String = "Update RPT_F04 set  AMT_W=" & DbHelper.GetStr(row("AA")) & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(DbHelper.GetStr(row("BUSINESSTYPEDESC")), 3) & "' and (Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') "
                DbHelper.ExecuteNonQuery(KW1)
            Next
        End If

    End Sub
    Private Sub INSERTHEADER()
        Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM Header where ID='F04'")
        If dt.Rows.Count = 0 Then
            Dim k As String = "INSERT INTO Header(ID,Nm,S1,S2,S3,S4,S5,S6,PP) " & _
                          " values(N'F04',N'" & txtHeader.Text & "',N'" & txtSig1.Text & "',N'" & txtSig2.Text & "',N'" & txtSig3.Text & "',N'" & txtSig4.Text & "',N'" & txtSig5.Text & "',N'" & txtSig6.Text & "',N'" & (TxtPP.Text) & "') "
            DbHelper.ExecuteNonQuery(k)
        Else
            DbHelper.ExecuteNonQuery("UPDATE Header set Nm=N'" & txtHeader.Text & "',S1=N'" & txtSig1.Text & "',S2=N'" & txtSig2.Text & "',S3=N'" & txtSig3.Text & "',S4=N'" & txtSig4.Text & "',S5=N'" & txtSig5.Text & "',S6=N'" & txtSig6.Text & "',PP=N'" & (TxtPP.Text) & "' where ID='F04' ")
        End If
    End Sub
    Private Sub UPP_INTER()
        'Dim KK As New ADODB.Recordset
        Dim aa As String = " SELECT  * from AP_MM where month(MM)='" & Month(DMY.Value) & "' and  year(MM)='" & Year(DMY.Value) & "'   "
        Dim dt As DataTable = DbHelper.GetDataTable(aa)
        If dt.Rows.Count = 0 Then
            'Dim HH As String = "INSERT INTO AP_MM (MM,A,B,C,D,E,Amt) values('" & Format(DMY.Value, "yyyy/MM/dd") & "'," & CDbl(txtA.Text) & "," & CDbl(txtB.Text) & "," & CDbl(txtC.Text) & "," & CDbl(txtD.Text) & "," & CDbl(txtE.Text) & ",0 ) "
            'DbHelper.ExecuteNonQuery(HH)
            Dim HH As String = "INSERT INTO AP_MM (MM,A,B,C,D,E,Amt) values('" & Format(DMY.Value, "yyyy/MM/dd") & "',N'" & (txtA.Text) & "',N'" & (txtB.Text) & "',N'" & (txtC.Text) & "',N'" & (txtD.Text) & "',N'" & (txtE.Text) & "',0 ) "
            DbHelper.ExecuteNonQuery(HH)
        Else
            'DbHelper.ExecuteNonQuery("UPDATE AP_MM set A=" & CDbl(txtA.Text) & ",b=" & CDbl(txtB.Text) & ",c=" & CDbl(txtC.Text) & ",d=" & CDbl(txtD.Text) & ", e=" & CDbl(txtE.Text) & "  where month(MM)='" & Month(DMY.Value) & "' and  year(MM)='" & Year(DMY.Value) & "'  ")
            DbHelper.ExecuteNonQuery("UPDATE AP_MM set A=N'" & (txtA.Text) & "',b=N'" & (txtB.Text) & "',c=N'" & (txtC.Text) & "',d=N'" & (txtD.Text) & "', e=N'" & (txtE.Text) & "'  where month(MM)='" & Month(DMY.Value) & "' and  year(MM)='" & Year(DMY.Value) & "'  ")

        End If
        'MsgBox("OK")
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'Call LoadSqlData(" select * from RPT_F01 where ItemID='F01' ", RSC)
        'If RSC.RecordCount = 0 Then
        '    DbHelper.ExecuteNonQuery("INSERT INTO RPT_F01(ItemID) values (N'F01') ")
        'End If
        Call UPP_INTER()
        INSERTHEADER()

        DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set AccNo=0,Acc=0,Acc_W=0,Amt=0,Amt_w=0,Dep=0,Dep_Amt=0  ")
        Call CALLC()
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtA.Text) & " where Grp_Nm=N'1. ສິນເຊື່ອປົກກະຕິ (A)' ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtB.Text) & " where Grp_Nm=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtC.Text) & " where Grp_Nm=N'3. ສິນເຊື່ອຕໍ່າກວ່າມາດຕະຖານ (C)' ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtD.Text) & " where Grp_Nm=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)' ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtE.Text) & " where (Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtA.Text) & " where left(Grp_Nm,1)=N'1' ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtB.Text) & " where left(Grp_Nm,1)=N'2' ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtC.Text) & " where left(Grp_Nm,1)=N'3' ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtD.Text) & " where left(Grp_Nm,1)=N'4' ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtE.Text) & " where left(Grp_Nm,1)=N'5' ")

        DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep='" & (txtA.Text) & "' where left(Grp_Nm,1)=N'1' ")
        DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep='" & (txtB.Text) & "' where left(Grp_Nm,1)=N'2' ")
        DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep='" & (txtC.Text) & "' where left(Grp_Nm,1)=N'3' ")
        DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep='" & (txtD.Text) & "' where left(Grp_Nm,1)=N'4' ")
        DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep='" & (txtE.Text) & "' where left(Grp_Nm,1)=N'5' ")

        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep_Amt=Amt*Dep/100  ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=0 where Dep_Amt=0  ")
        'Dim RsKK As New ADODB.Recordset
        Dim PPP As String = " SELECT N'" & RmFrdate.Text & "' as DD,N'" & txtSig1.Text & "' as S1,N'" & txtSig2.Text & "' as S2,N'" & txtSig3.Text & "' as S3, N'" & txtSig4.Text & "' as S4,  N'" & txtSig5.Text & "' as S5,   N'" & txtSig6.Text & "' as S6,  N'" & TxtPP.Text & "' as pp, * from RPT_F04 order by CNT ASC "
        Dim dt As DataTable = DbHelper.GetDataTable(PPP)
        With dt
            If .Rows.Count = 0 Then MsgBox("Data empty", vbInformation, "Check") : Exit Sub
            Dim Frm As New FmPreview

            Dim Rpt As New F04
            'Dim myTextObjectOnReport As CrystalDecisions.CrystalReports.Engine.TextObject
            'myTextObjectOnReport = CType(Rpt.ReportDefinition.ReportObjects.Item("HD"), CrystalDecisions.CrystalReports.Engine.TextObject)
            'myTextObjectOnReport.Text = txtHeader.Text
            'myTextObjectOnReport = CType(Rpt.ReportDefinition.ReportObjects.Item("NO"), CrystalDecisions.CrystalReports.Engine.TextObject)
            'myTextObjectOnReport.Text = RmFrdate.Text

            Rpt.SetDataSource(dt)
            Frm.ReportViewer.ReportSource = Rpt
            Frm.ReportViewer.DisplayGroupTree = False
            Frm.WindowState = FormWindowState.Maximized
            Frm.Show()
            Rpt = Nothing
        End With
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub RDate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RDate.CheckedChanged
        Call SelectDate()
    End Sub

    Private Sub RMonth_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RMonth.CheckedChanged
        Call SelectDate()
    End Sub

    Private Sub RPeriod_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RPeriod.CheckedChanged
        Call SelectDate()
    End Sub

    Private Sub RFirsthalfyear_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RFirsthalfyear.CheckedChanged
        Call SelectDate()
    End Sub

    Private Sub RYearAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RYearAll.CheckedChanged
        Call SelectDate()
    End Sub

    Private Sub CFromMonth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CFromMonth.SelectedIndexChanged
        CToMonth.SelectedIndex = CFromMonth.SelectedIndex
        Call SelectDate()
    End Sub

    Private Sub CToMonth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CToMonth.SelectedIndexChanged
        If CToMonth.SelectedIndex < CFromMonth.SelectedIndex Then
            CToMonth.SelectedIndex = CFromMonth.SelectedIndex
        End If
        Call SelectDate()
    End Sub

    Private Sub CPeriod_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CPeriod.SelectedIndexChanged
        Call SelectDate()
    End Sub

    Private Sub CFirsthalfyear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CFirsthalfyear.SelectedIndexChanged
        Call SelectDate()
    End Sub

    Private Sub DtAllYear_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DtAllYear.ValueChanged
        SelectDate()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Frm_F01Edit.ShowDialog()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DMY.ValueChanged
        Call MMM()
    End Sub
    Private Sub MMM()
        'Dim KK As New ADODB.Recordset
        Dim aa As String = " SELECT  * from AP_MM where month(MM)='" & Month(DMY.Value) & "' and  year(MM)='" & Year(DMY.Value) & "'   "
        Dim dt As DataTable = DbHelper.GetDataTable(aa)
        If dt.Rows.Count <> 0 Then
            txtA.Text = DbHelper.GetStr(dt.Rows(0)("A"))
            txtB.Text = DbHelper.GetStr(dt.Rows(0)("B"))
            txtC.Text = DbHelper.GetStr(dt.Rows(0)("C"))
            txtD.Text = DbHelper.GetStr(dt.Rows(0)("D"))
            txtE.Text = DbHelper.GetStr(dt.Rows(0)("E"))
        Else
            txtA.Text = ""
            txtB.Text = ""
            txtC.Text = ""
            txtD.Text = ""
            txtE.Text = ""
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        'Call LoadSqlData(" select * from RPT_F01 where ItemID='F01' ", RSC)
        'If RSC.RecordCount = 0 Then
        '    DbHelper.ExecuteNonQuery("INSERT INTO RPT_F01(ItemID) values (N'F01') ")
        'End If
        Call UPP_INTER()
        INSERTHEADER()

        DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set AccNo=0,Acc=0,Acc_W=0,Amt=0,Amt_w=0,Dep=0,Dep_Amt=0  ")
        Call CALLC()
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtA.Text) & " where Grp_Nm=N'1. ສິນເຊື່ອປົກກະຕິ (A)' ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtB.Text) & " where Grp_Nm=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtC.Text) & " where Grp_Nm=N'3. ສິນເຊື່ອຕໍ່າກວ່າມາດຕະຖານ (C)' ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtD.Text) & " where Grp_Nm=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)' ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtE.Text) & " where (Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtA.Text) & " where left(Grp_Nm,1)=N'1' ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtB.Text) & " where left(Grp_Nm,1)=N'2' ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtC.Text) & " where left(Grp_Nm,1)=N'3' ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtD.Text) & " where left(Grp_Nm,1)=N'4' ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=" & CDbl(txtE.Text) & " where left(Grp_Nm,1)=N'5' ")

        DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep='" & (txtA.Text) & "' where left(Grp_Nm,1)=N'1' ")
        DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep='" & (txtB.Text) & "' where left(Grp_Nm,1)=N'2' ")
        DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep='" & (txtC.Text) & "' where left(Grp_Nm,1)=N'3' ")
        DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep='" & (txtD.Text) & "' where left(Grp_Nm,1)=N'4' ")
        DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep='" & (txtE.Text) & "' where left(Grp_Nm,1)=N'5' ")

        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep_Amt=Amt*Dep/100  ")
        'DbHelper.ExecuteNonQuery("UPDATE RPT_F04 set Dep=0 where Dep_Amt=0  ")
        'Dim RsKK As New ADODB.Recordset
        Dim PPP As String = " SELECT N'" & RmFrdate.Text & "' as DD,N'" & txtSig1.Text & "' as S1,N'" & txtSig2.Text & "' as S2,N'" & txtSig3.Text & "' as S3, N'" & txtSig4.Text & "' as S4,  N'" & txtSig5.Text & "' as S5,   N'" & txtSig6.Text & "' as S6,  N'" & TxtPP.Text & "' as pp, * from RPT_F04 order by CNT ASC "
        Dim dt As DataTable = DbHelper.GetDataTable(PPP)
        With dt
            If .Rows.Count = 0 Then MsgBox("Data empty", vbInformation, "Check") : Exit Sub
            Dim Frm As New FmPreview

            Dim Rpt As New F04
            'Dim myTextObjectOnReport As CrystalDecisions.CrystalReports.Engine.TextObject
            'myTextObjectOnReport = CType(Rpt.ReportDefinition.ReportObjects.Item("HD"), CrystalDecisions.CrystalReports.Engine.TextObject)
            'myTextObjectOnReport.Text = txtHeader.Text
            'myTextObjectOnReport = CType(Rpt.ReportDefinition.ReportObjects.Item("NO"), CrystalDecisions.CrystalReports.Engine.TextObject)
            'myTextObjectOnReport.Text = RmFrdate.Text

            Rpt.SetDataSource(dt)
            Rpt.Refresh()
            Frm.ReportViewer.ReportSource = Rpt
            Frm.ReportViewer.ExportReport()
            Frm = Nothing

            'Rpt.SetDataSource(dt)
            'Frm.ReportViewer.ReportSource = Rpt
            'Frm.ReportViewer.DisplayGroupTree = False
            'Frm.WindowState = FormWindowState.Maximized
            'Frm.Show()
            'Rpt = Nothing
        End With
    End Sub
End Class