Imports System.Data.SqlClient

Public Class FrmRpt_F011
    Dim Biz_Type, Loan_Type, sqlNew, Rpt_ID As String

    Dim RelationShip, Relation As String
    Private Sub INSERTHEADER()
        'Call LoadSqlData("SELECT * FROM Header where ID='F01'", RSC)
        Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM Header where ID='F01'")
        If dt.Rows.Count = 0 Then
            Dim k As String = "INSERT INTO Header(ID,Nm,S1,S2,S3,PP) " & _
                          " values(N'F01',N'" & txtHeader.Text & "',N'" & txtSig1.Text & "',N'" & txtSig2.Text & "',N'" & txtSig3.Text & "',N'" & (TxtPP.Text) & "') "
            DbHelper.ExecuteNonQuery(k)
        Else
            DbHelper.ExecuteNonQuery("UPDATE Header set Nm=N'" & txtHeader.Text & "',S1=N'" & txtSig1.Text & "',S2=N'" & txtSig2.Text & "',S3=N'" & txtSig3.Text & "',PP=N'" & (TxtPP.Text) & "' where ID='F01' ")
        End If
    End Sub

    Private Sub LoadHeader()
        'Call LoadSqlData("SELECT * FROM Header where ID='F01'", RSC)
        Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM Header where ID='F01'")
        If dt.Rows.Count <> 0 Then
            txtHeader.Text = dt.Rows(0)("Nm").ToString
            txtSig1.Text = dt.Rows(0)("S1").ToString
            txtSig2.Text = dt.Rows(0)("S2").ToString
            txtSig3.Text = dt.Rows(0)("S3").ToString
            TxtPP.Text = dt.Rows(0)("PP").ToString
        End If

    End Sub
    Private Sub Frm_loan_Due_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        INSERTHEADER()
        DbHelper.ExecuteNonQuery(" UPDATE RPT_F01 set No2=N'" & CFromMonth.Text & "/" & DtYearFormonth.Text & "' WHERE ItemID=N'F01' ")

        'Call LoadSqlData(" select * from RPT_F01 where ItemID='F01' ", RSC)
        Dim dt As DataTable = DbHelper.GetDataTable(" select * from RPT_F01 where ItemID='F01' ")
        If dt.Rows.Count = 0 Then
            DbHelper.ExecuteNonQuery("INSERT INTO RPT_F01(ItemID) values (N'F01') ")
        End If
        'Dim RsKK As New ADODB.Recordset
        'With RsKK
            Dim PPP As String = "SELECT N'" & RmFrdate.Text & "' as DD,N'" & txtSig1.Text & "' as S1,N'" & txtSig2.Text & "' as S2,N'" & txtSig3.Text & "' as S3,N'" & TxtPP.Text & "' as pp, * from RPT_F01 where ItemID=N'F01' "
            'Call LoadSqlData(PPP, RsKK)
            Dim dt2 As DataTable = DbHelper.GetDataTable(PPP)
            If dt2.Rows.Count = 0 Then MsgBox("Data empty", vbInformation, "Check") : Exit Sub
            Dim Frm As New FmPreview

            Dim Rpt As New F01
            'Dim myTextObjectOnReport As CrystalDecisions.CrystalReports.Engine.TextObject
            'myTextObjectOnReport = CType(Rpt.ReportDefinition.ReportObjects.Item("HD"), CrystalDecisions.CrystalReports.Engine.TextObject)
            'myTextObjectOnReport.Text = txtHeader.Text
            'myTextObjectOnReport = CType(Rpt.ReportDefinition.ReportObjects.Item("NO"), CrystalDecisions.CrystalReports.Engine.TextObject)
            'myTextObjectOnReport.Text = RmFrdate.Text
            'myTextObjectOnReport = CType(Rpt.ReportDefinition.ReportObjects.Item("S1"), CrystalDecisions.CrystalReports.Engine.TextObject)
            'myTextObjectOnReport.Text = txtSig1.Text
            'myTextObjectOnReport = CType(Rpt.ReportDefinition.ReportObjects.Item("S2"), CrystalDecisions.CrystalReports.Engine.TextObject)
            'myTextObjectOnReport.Text = txtSig2.Text
            'myTextObjectOnReport = CType(Rpt.ReportDefinition.ReportObjects.Item("S3"), CrystalDecisions.CrystalReports.Engine.TextObject)
            'myTextObjectOnReport.Text = txtSig3.Text
            'myTextObjectOnReport = CType(Rpt.ReportDefinition.ReportObjects.Item("PP"), CrystalDecisions.CrystalReports.Engine.TextObject)
            'myTextObjectOnReport.Text = TxtPP.Text
            Rpt.SetDataSource(dt2)
            Frm.ReportViewer.ReportSource = Rpt
            Frm.ReportViewer.DisplayGroupTree = False
            Frm.WindowState = FormWindowState.Maximized
            Frm.Show()
            Rpt = Nothing
        'End With
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
End Class