Public Class FrmRpt_Group
    Dim rpt As Object
    Dim Sec As String
    Dim MM As String
    Dim SDep As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'ShowShortFrm()
        Me.Close()
    End Sub

    Private Sub footder()
        Dim Signal As CrystalDecisions.CrystalReports.Engine.TextObject
        Signal = CType(rpt.ReportDefinition.ReportObjects.Item("Signal5"), CrystalDecisions.CrystalReports.Engine.TextObject)
        Signal.Text = MDSignal5
        Signal = CType(rpt.ReportDefinition.ReportObjects.Item("Signal4"), CrystalDecisions.CrystalReports.Engine.TextObject)
        Signal.Text = MDSignal4
        Signal = CType(rpt.ReportDefinition.ReportObjects.Item("Signal3"), CrystalDecisions.CrystalReports.Engine.TextObject)
        Signal.Text = MDSignal3
        Signal = CType(rpt.ReportDefinition.ReportObjects.Item("Signal2"), CrystalDecisions.CrystalReports.Engine.TextObject)
        Signal.Text = MDSignal2
        Signal = CType(rpt.ReportDefinition.ReportObjects.Item("Signal1"), CrystalDecisions.CrystalReports.Engine.TextObject)
        Signal.Text = MDSignal1
        Signal = CType(rpt.ReportDefinition.ReportObjects.Item("Place"), CrystalDecisions.CrystalReports.Engine.TextObject)
        Signal.Text = MDPlace
    End Sub
    Private Sub Sig_Load()
        Dim RsSig As New ADODB.Recordset
        With RsSig
            Call LoadSqlData("SELECT * FROM Header_Asset WHERE Head_ID='" & "001" & "' ", RsSig)
            If .RecordCount <> 0 Then
                Head_Nm.Text = .Fields("Head_Nm").Value.ToString
                Signal1.Text = .Fields("Signal1").Value.ToString
                Signal2.Text = .Fields("Signal2").Value.ToString
                Signal3.Text = .Fields("Signal3").Value.ToString
                Signal4.Text = .Fields("Signal4").Value.ToString
                Signal5.Text = .Fields("Signal5").Value.ToString
                Place.Text = .Fields("Place").Value.ToString
            End If
        End With
        RsSig = Nothing
    End Sub
    Private Sub LoadHeader_Asset()
        Dim RsSig As New ADODB.Recordset
        With RsSig
            CNN.Execute("UPDATE Header_Asset SET " & _
                           " Head_Nm = N'" & (Head_Nm.Text) & "'," & _
                           " Signal1 = N'" & (Signal1.Text) & "', " & _
                           " Signal2 = N'" & (Signal2.Text) & "'," & _
                           " Signal3 = N'" & (Signal3.Text) & "', " & _
                         " Signal4 = N'" & (Signal4.Text) & "'," & _
                           " Signal5 = N'" & (Signal5.Text) & "', " & _
                           " Place = N'" & (Place.Text) & "' " & _
                           " WHERE (Head_ID = '" & "001" & "')")

            Call LoadSqlData("SELECT * FROM Header_Asset WHERE Head_ID='" & "001" & "' ", RsSig)
            If .RecordCount <> 0 Then
                MDHead = .Fields("Head_Nm").Value
                MDSignal1 = .Fields("Signal1").Value
                MDSignal2 = .Fields("Signal2").Value
                MDSignal3 = .Fields("Signal3").Value
                MDSignal4 = .Fields("Signal4").Value
                MDSignal5 = .Fields("Signal5").Value
                MDPlace = .Fields("Place").Value
            End If
        End With
        RsSig = Nothing
    End Sub
    Private Sub FrmRpt_Group_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Call Loadlang()
        'SetControlText(Me)
        'ChgChildForm()
        Sig_Load()
        DTMon.Value = Format(Date.Today, "MM/yyyy")
        DTYear.Value = Date.Today
        DTMon.Value = "01/" & Month(DTMon.Value) & "/" & Year(DTMon.Value)
        Call LdGrp()
        Call LdSec()
        'Call LdDep()
        'Call LdCompany()
        cmbTerm.SelectedIndex = 0

        '                    0        1          2                                        3         4           5           6             7             8             9             10            11       12             13                14               15
        If Lang = True Then
            GHead.Text = "Header_Asset and Footer"
            H.Text = "Heading Name:"
            S1.Text = "Signal 1:"
            S2.Text = "Signal 2:"
            S3.Text = "Signal 3:"
            S4.Text = "Signal 4:"
            S5.Text = "Signal 5:"
            P.Text = "Place Name:"
        Else
            GHead.Text = "Header_Asset and Footer"
            H.Text = "Heading Name:"
            S1.Text = "Signal 1:"
            S2.Text = "Signal 2:"
            S3.Text = "Signal 3:"
            S4.Text = "Signal 4:"
            S5.Text = "Signal 5:"
            P.Text = "Place Name:"
        End If
        If Lang = True Then
            Label3.Text = "Office Name"
            Label6.Text = "Section Name"
            Label2.Text = "Group Name"
            Label1.Text = "All Assets Report"
            Label4.Text = "Group Code"
            chkSum.Text = "Follow Group"
            chkBranch.Text = "Don't Branch"
            optMon.Text = "Monthly"
            optTerm.Text = "Termly"
            optYear.Text = "Yearly"
            Label24.Text = "Number No"
            Label5.Text = "Show List"
            GroupBox1.Text = "Time Report"
            GroupBox2.Text = "Scope of Report"
            Button3.Text = "Transfer Acc"
            Button4.Text = "Compare"
            FG.FormatString = "^No. |<Asset ID   |<Number No |<Assets Name                             |^Group|^Uesd Date    |^Sesding|>Amounth    |^Dep Date  |>Year Dep      |>Month/Dep    |^Month |^Pre month|>PreDep     |>Month Dep      |>All Dep         |>Remain         "
            FGIT.FormatString = "^ລດ |^ເລືອກ|<ລະຫັດ|<ຊື່ພະແນກ                         "

            CmbShow.Items.Clear()
            CmbShow.Items.Add("All Show")
            CmbShow.Items.Add("Specific items are used")
            CmbShow.Items.Add("Because of the built")
        Else
            CmbShow.Items.Clear()
            CmbShow.Items.Add("ສະແດງທັງໝົດ")
            CmbShow.Items.Add("ສະເພາະລາຍການພວມນຳໃຊ້")
            CmbShow.Items.Add("ສະເພາະລາຍການລໍຖ້າສະສາງ")
            FG.FormatString = "^ລ/ດ |<ລະຫັດຊັບສິນ |<ເລກປະຈຳຕົວ |<ຊື່ຊັບສິນ                                   |^ໝວດ |^ວັນທີນຳໃຊ້     |^ອາຍຸນຳໃຊ້|>ມູນຄ່າຊັບສິນ    |^ວັນທີສະສາງ  |>ຫຼຸ້ຍຫ້ຽນຕໍ່ປີ       |>ຫຼຸ້ຫ້ຽນຕໍ່ເດືອນ   |^ຈນ ເດືອນ|^ຈນດ ຜ່ານມາ|>ຫັກຜ່ານມາ     |>ຫັກໃນເດືອນ      |>ຫັກສະສົມ          |>ຍັງເທືອ         "
            FGIT.FormatString = "^ລດ |^ເລືອກ|<ລະຫັດ|<ຊື່ພະແນກ                         "

            Button3.Text = "ໂອນເຂົ້າບັນຊີ"
            Button4.Text = "ສົມທຽບ"
            GroupBox1.Text = "ຊ່ວງເວລາລາຍງານ"
            GroupBox2.Text = "ຂອບເຂດການລາຍງານ"
            Label24.Text = "ເລກທີ"
            Label5.Text = "ສະແດງລາຍການ"
            optMon.Text = "ປະຈຳເດືອນ"
            optTerm.Text = "ປະຈຳງວດ"
            optYear.Text = "ປະຈໍາປີ"
            chkBranch.Text = "ບໍ່ແຍກສາຂາ"
            chkSum.Text = "ສັງລວມຕາມໝວດ"
            Label4.Text = "ສະເພາະລະຫັດຊັບສິນ"
            Label1.Text = "ລາຍງານແບບສັງລວມ"
            Label3.Text = "ສຳນັກງານ"
            Label6.Text = "ພະແນກ"
            Label2.Text = "ໝວດຊັບສິນ"
        End If
        CmbShow.SelectedIndex = 0

        FGIT.set_ColDataType(1, VSFlex8U.DataTypeSettings.flexDTBoolean)

        'Call LOADFG1()
        FGIT.Visible = False
    End Sub

    Private Sub LdCompany()
        Dim gRS As New ADODB.Recordset
        CmbCompany.Items.Clear()
        Call LoadSqlData("Select * from office Order by off_id", gRS)
        If gRS.RecordCount <> 0 Then
            While Not gRS.EOF
                CmbCompany.Items.Add(gRS.Fields("off_name").Value)
                gRS.MoveNext()
            End While
        End If
        CmbCompany.SelectedIndex = 0
    End Sub
    Private Sub LdGrp()
        Dim gRS As New ADODB.Recordset
        cmbGrp.Items.Clear()
        If Lang = True Then
            cmbGrp.Items.Add(" All Group ")
            Call LoadSqlData("Select * from Groups Order by Group_ID", gRS)
            If gRS.RecordCount <> 0 Then
                While Not gRS.EOF
                    cmbGrp.Items.Add(gRS.Fields("Group_NmE").Value.ToString)
                    gRS.MoveNext()
                End While
            End If
            cmbGrp.SelectedIndex = 0
        Else
            cmbGrp.Items.Add(" ສະແດງທັງໝົດ ")
            Call LoadSqlData("Select * from Groups Order by Group_ID", gRS)
            If gRS.RecordCount <> 0 Then
                While Not gRS.EOF
                    cmbGrp.Items.Add(gRS.Fields("Group_Nm").Value.ToString)
                    gRS.MoveNext()
                End While
            End If
            cmbGrp.SelectedIndex = 0
        End If
    End Sub

    Private Sub LdSec()
        Dim sRS As New ADODB.Recordset
        cmbSec.Items.Clear()
        If Lang = True Then
            cmbSec.Items.Add("** All Sections ***")
            Call LoadSqlData("Select * from AP_Office  where Off_ID<>'00' Order by Off_ID", sRS)
            If sRS.RecordCount <> 0 Then
                While Not sRS.EOF
                    cmbSec.Items.Add(sRS.Fields("Off_NmE").Value.ToString)
                    sRS.MoveNext()
                End While
            End If
            If cmbSec.Items.Count > 0 Then cmbSec.SelectedIndex = 0
        Else
            If Mpermiss = "Admin" Then
                cmbSec.Items.Add("** ສະແດງທັງໝົດ **")
                Call LoadSqlData("Select * from AP_Office  where Off_ID<>'00' Order by Off_ID", sRS)
                If sRS.RecordCount <> 0 Then
                    While Not sRS.EOF
                        cmbSec.Items.Add(sRS.Fields("Off_Name").Value.ToString)
                        sRS.MoveNext()
                    End While
                End If
                If cmbSec.Items.Count > 0 Then cmbSec.SelectedIndex = 0
            Else
                cmbSec.Items.Clear()
                Call load_Cmb("SELECT Off_Name FROM AP_Office where 1=1 and off_id=N'" & Off_Id & "' ORDER BY off_id ASC", "Off_Name", cmbSec)
                If cmbSec.Items.Count > 0 Then
                    cmbSec.SelectedIndex = 0
                End If
            End If
        End If
    End Sub

    Private Sub LdDep()
        Dim sRS As New ADODB.Recordset
        cmbDeprt.Items.Clear()
        If Lang = True Then
            cmbDeprt.Items.Add("** All Department ***")
            Call LoadSqlData("Select * from Department Order by DepartmentID", sRS)
            If sRS.RecordCount <> 0 Then
                While Not sRS.EOF
                    cmbDeprt.Items.Add(sRS.Fields("DepartmentNmE").Value.ToString)
                    sRS.MoveNext()
                End While
            End If
            If cmbDeprt.Items.Count > 0 Then cmbDeprt.SelectedIndex = 0
        Else
            cmbDeprt.Items.Add("** ສະແດງທັງໝົດ ***")
            Call LoadSqlData("Select * from Department Order by DepartmentID", sRS)
            If sRS.RecordCount <> 0 Then
                While Not sRS.EOF
                    cmbDeprt.Items.Add(sRS.Fields("DepartmentNm").Value.ToString)
                    sRS.MoveNext()
                End While
            End If
            If cmbDeprt.Items.Count > 0 Then cmbDeprt.SelectedIndex = 0
        End If
    End Sub

    Private Sub cmbGrp_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbGrp.SelectedIndexChanged
        Dim gRS As New ADODB.Recordset
        If Lang = True Then
            Call LoadSqlData("select * from Groups Where Group_NmE=N'" & Trim(cmbGrp.Text) & "'", gRS)
            If gRS.RecordCount <> 0 Then
                txtGrp.Text = Trim(gRS.Fields("Group_ID").Value.ToString)
                TxtLH.Text = Trim(gRS.Fields("Ac_Code").Value.ToString)
                txtAcc.Text = Trim(gRS.Fields("Dep_Code").Value.ToString)
                TxtCertify.Text = txtGrp.Text & "." & Format((DTMon.Value), "MM.yy") & "/" & Trim(txtCompany.Text)
            Else
                txtGrp.Text = ""
            End If
        Else
            Call LoadSqlData("select * from Groups Where Group_Nm=N'" & Trim(cmbGrp.Text) & "'", gRS)
            If gRS.RecordCount <> 0 Then
                txtGrp.Text = Trim(gRS.Fields("Group_ID").Value.ToString)
                TxtLH.Text = Trim(gRS.Fields("Ac_Code").Value.ToString)
                txtAcc.Text = Trim(gRS.Fields("Dep_Code").Value.ToString)
                txtCompany.Text = txtSec.Text
                TxtCertify.Text = txtGrp.Text & "." & Format((DTMon.Value), "MM.yy") & "/" & Trim(txtCompany.Text)
            Else
                txtGrp.Text = ""
            End If


        End If
        Call LoadSqlData("select * from Acc_Code Where Ac_Code=N'" & Trim(txtAcc.Text) & "'", gRS)
        If gRS.RecordCount <> 0 Then
            TxtDrNm.Text = Trim(gRS.Fields("Name_L").Value.ToString)
        Else
            TxtDrNm.Text = ""
        End If

        Call LoadSqlData("select * from Acc_Code Where Ac_Code=N'" & Trim(TxtLH.Text) & "'", gRS)
        If gRS.RecordCount <> 0 Then
            TxtCrNm.Text = Trim(gRS.Fields("Name_L").Value.ToString)
        Else
            TxtCrNm.Text = ""
        End If
    End Sub

    Private Sub optYear_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optYear.CheckedChanged
        DTYear.Enabled = True
        DTMon.Enabled = False
        cmbTerm.Enabled = False
        dtTerm.Enabled = False
    End Sub

    Private Sub optMon_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMon.CheckedChanged
        DTYear.Enabled = False
        DTMon.Enabled = True
        cmbTerm.Enabled = False
        dtTerm.Enabled = False
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        cmbGrp.Enabled = True
        chkSum.Checked = False
        chkSum.Enabled = False
        Call cmbGrp_SelectedIndexChanged(sender, e)
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        cmbGrp.Enabled = False
        txtGrp.Text = ""
        chkSum.Enabled = True
    End Sub
    Private Sub SUMMARYASSET()
        Dim Rs, Rs1 As New ADODB.Recordset

        Dim rpt As New Object
        Dim FrmPreview As New FmPreview : FrmClosing()
        rpt = New Crystal_SummaryALL_GrpNEW



        Call Office()

        CNN.Execute("update Rpt_Grp set Rpt_Grp.SectionNm=Sections.SecNmL from Sections,Rpt_Grp where Rpt_Grp.Section=Sections.SecID ")
        CNN.Execute("update Rpt_Grp set Rpt_Grp.DepartmenNm=Department.DepartmentNm from Department,Rpt_Grp where Rpt_Grp.DepartmentID=Department.DepartmentID ")
        CNN.Execute("update Rpt_Grp set Rpt_Grp.Asset_No=Assets.Asset_No  from Assets,Rpt_Grp where Rpt_Grp.AssetID=Assets.AssetID ")
        CNN.Execute("update Rpt_Grp set Rpt_Grp.Using_By=Assets.Using_By,Rpt_Grp.Budget=Assets.Budget from Assets,Rpt_Grp where Rpt_Grp.AssetID=Assets.AssetID ")
        CNN.Execute("update Rpt_Grp set Rpt_Grp.Group_Nm=Groups.Group_Nm from Groups,Rpt_Grp where Rpt_Grp.Group_ID=Groups.Group_ID ")
        CNN.Execute("update Rpt_Grp set Rpt_Grp.Grp_No=Groups.Grp_No from Groups,Rpt_Grp where rpt_grp.Group_ID=groups.Group_ID")


        CNN.Execute("UPDATE Rpt_NEW set amt=0,Dep_year=0,Dep_Month=0,Dep_TT=0,Rem=0")
        Call KKK()
        CNN.Execute("update Rpt_NEW set Amt=(select SUM(Amt) from Rpt_NEW where NO='1') where Grp_Acc=N'I' ")
        CNN.Execute("update Rpt_NEW set Amt=(select SUM(Amt) from Rpt_NEW where NO='2') where Grp_Acc=N'II' ")
        CNN.Execute("update Rpt_NEW set Amt=(select SUM(Amt) from Rpt_NEW where NO='3') where Grp_Acc=N'III' ")
        CNN.Execute("update Rpt_NEW set Dep_year=(select SUM(Dep_year) from Rpt_NEW where NO='1') where Grp_Acc=N'I' ")
        CNN.Execute("update Rpt_NEW set Dep_year=(select SUM(Dep_year) from Rpt_NEW where NO='2') where Grp_Acc=N'II' ")
        CNN.Execute("update Rpt_NEW set Dep_year=(select SUM(Dep_year) from Rpt_NEW where NO='3') where Grp_Acc=N'III' ")
        CNN.Execute("update Rpt_NEW set Dep_Month=(select SUM(Dep_Month) from Rpt_NEW where NO='1') where Grp_Acc=N'I' ")
        CNN.Execute("update Rpt_NEW set Dep_Month=(select SUM(Dep_Month) from Rpt_NEW where NO='2') where Grp_Acc=N'II' ")
        CNN.Execute("update Rpt_NEW set Dep_Month=(select SUM(Dep_Month) from Rpt_NEW where NO='3') where Grp_Acc=N'III' ")
        CNN.Execute("update Rpt_NEW set Dep_TT=(select SUM(Dep_TT) from Rpt_NEW where NO='1') where Grp_Acc=N'I' ")
        CNN.Execute("update Rpt_NEW set Dep_TT=(select SUM(Dep_TT) from Rpt_NEW where NO='2') where Grp_Acc=N'II' ")
        CNN.Execute("update Rpt_NEW set Dep_TT=(select SUM(Dep_TT) from Rpt_NEW where NO='3') where Grp_Acc=N'III' ")
        CNN.Execute("update Rpt_NEW set rem=(select SUM(rem) from Rpt_NEW where NO='1') where Grp_Acc=N'I' ")
        CNN.Execute("update Rpt_NEW set rem=(select SUM(rem) from Rpt_NEW where NO='2') where Grp_Acc=N'II' ")
        CNN.Execute("update Rpt_NEW set rem=(select SUM(rem) from Rpt_NEW where NO='3') where Grp_Acc=N'III' ")
        CNN.Execute("UPDATE Rpt_NEW set amt=0 where amt is null ")
        CNN.Execute("UPDATE Rpt_NEW set Dep_year=0 where Dep_year is null ")
        CNN.Execute("UPDATE Rpt_NEW set Dep_Month=0 where Dep_Month is null ")
        CNN.Execute("UPDATE Rpt_NEW set Dep_TT=0 where Dep_TT is null ")
        CNN.Execute("UPDATE Rpt_NEW set Rem=0 where Rem is null ")
        CNN.Execute("UPDATE Rpt_NEW set Dep_year=0,Dep_Month=0,Dep_TT=0 where Grp_no='201' ")
        CNN.Execute("UPDATE Rpt_NEW set Rem=Amt  where Grp_no='201' ")
        Call LoadSqlData("Select * from Rpt_NEW where amt>0 Order by grp_No ASC", Rs)


        Dim myText3 As CrystalDecisions.CrystalReports.Engine.TextObject
        'myText1 = CType(rpt.ReportDefinition.ReportObjects.Item("txtGrp"), CrystalDecisions.CrystalReports.Engine.TextObject)
        'If txtGrp.Text <> "" Then
        '    myText1.Text = Trim(cmbGrp.Text)
        'End If

        '========================
        Dim mPer As String
        If optMon.Checked = True Then

            mPer = "ປະຈຳເດືອນ " & Format(DTMon.Value, "MM/yyyy")

        ElseIf optYear.Checked = True Then

            mPer = "ປະຈຳປີ " & Format(DTYear.Value, "yyyy")
        Else

            mPer = "ປະຈຳງວດ " & Trim(cmbTerm.Text) & " / " & Format(dtTerm.Value, "yyyy")
        End If
        '======================
        myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Tel"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText3.Text = "Tel: " & OffTel & " Fax: " & OffFax
        myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("HD"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText3.Text = mPer

        myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("OffNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText3.Text = OffName
        myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("SG5"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText3.Text = Sign1
        myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("SG4"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText3.Text = Sign2
        myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("SG3"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText3.Text = Sign3
        myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("SG2"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText3.Text = Sign4
        myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("SG1"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText3.Text = Sign5
        myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("PP"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText3.Text = PlaecL
        If cmbSec.SelectedIndex = 0 Then
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Nm"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = ""
        Else
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Nm"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = cmbSec.Text
        End If

        If Rs.RecordCount = 0 Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
        With rpt
            'Dim Rs1 As New ADODB.Recordset
            'Dim RO As CrystalDecisions.CrystalReports.Engine.ReportObject
            'Dim SRO As CrystalDecisions.CrystalReports.Engine.SubreportObject
            'Dim SubDoc As CrystalDecisions.CrystalReports.Engine.ReportDocument
            'SqlPrint = "SELECT  * from Ap_Image  where Img_Id='" & IMageID & "'"
            'Call LoadSqlData(SqlPrint, Rs1)
            'RO = rpt.ReportDefinition.Sections.Item("Section1").ReportObjects.Item("Subreport1")
            'SRO = CType(RO, CrystalDecisions.CrystalReports.Engine.SubreportObject)
            'SubDoc = SRO.OpenSubreport(SRO.SubreportName)
            'If Rs1.RecordCount > 0 Then
            '    SubDoc.SetDataSource(Rs1)
            '    FmPreview.ReportViewer.ReportSource = SubDoc
            'End If
            rpt.SetDataSource(Rs)
            rpt.Refresh()
            FmPreview.ReportViewer.ReportSource = rpt
            FmPreview.ReportViewer.DisplayGroupTree = False
            FmPreview.WindowState = FormWindowState.Maximized
            FmPreview.Show()
            'Call CloseRs(RSC)
            'Call CloseRs(Rs)
        End With
    End Sub
    Private Sub KKK()
        Dim Sec, SDep As String

        If cmbSec.SelectedIndex = 0 Then
            Sec = ""
        Else
            Sec = " AND section='" & txtSec.Text & "' "
        End If
        If cmbDeprt.SelectedIndex = 0 Then
            SDep = ""
        Else
            SDep = " AND DepartmentID='" & txtDep.Text & "' "
        End If
        Dim KK As String
        If CheckBox4.Checked = True Then
            Dim i As Integer
            For i = 1 To FGIT.Rows - 1
                If FGIT.get_ValueMatrix(i, 1) = True Then
                    Dim kq As String = "UPDATE Sections set Choose='1'  where SecID='" & FGIT.get_TextMatrix(i, 2) & "'  "
                    CNN.Execute(kq)
                End If
            Next



            CNN.Execute("UPDATE Rpt_Grp set Rpt_Grp.Choose=Sections.Choose from Rpt_Grp,Sections where Sections.SecID=Rpt_Grp.Section ")
            KK = " AND Choose=1 "
        Else
            KK = ""
        End If
        '===========201=====
        CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt_KIP),0) as Amt_KIP  from Rpt_Grp  where Grp_No='201' " & Sec & " " & SDep & " " & KK & " ) from " & _
              "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='201'")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Year=(select isnull(sum(Dep_Year),0) as Dep_Year  from Rpt_Grp  where Grp_No='201'  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='201'")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Month=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Grp_No='201' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='201'")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(TTDep),0) as TTDep  from Rpt_Grp  where Grp_No='201' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='201'")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Rem=(select isnull(sum(Remain),0) as Remain  from Rpt_Grp  where Grp_No='201' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='201'")
        '===========203=====
        CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt_KIP),0) as Amt_KIP  from Rpt_Grp  where Grp_No='203' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='203'")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Year=(select isnull(sum(Dep_Year),0) as Dep_Year  from Rpt_Grp  where Grp_No='203' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='203'")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Month=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Grp_No='203' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='203'")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(TTDep),0) as TTDep  from Rpt_Grp  where Grp_No='203' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='203'")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Rem=(select isnull(sum(Remain),0) as Remain  from Rpt_Grp  where Grp_No='203' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='203'")
        '===========212=====rem>0
        CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='212' and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='212' and grp_acc='3' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Year=(select isnull(sum(Dep_Year),0) as Dep_Year  from Rpt_Grp  where Grp_No='212'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='212'  and grp_acc='3' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Month=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Grp_No='212'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='212'  and grp_acc='3' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(TTDep),0) as TTDep  from Rpt_Grp  where Grp_No='212'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='212'  and grp_acc='3' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Rem=(select isnull(sum(Remain),0) as Remain  from Rpt_Grp  where Grp_No='212'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='212' and grp_acc='3' ")
        '===========212=====rem=0
        CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='212' and Remain=0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='212'  and grp_acc='19' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(Amt),0) as TTDep  from Rpt_Grp  where Grp_No='212' and Remain=0 " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='212' and grp_acc='19' ")
        '===========213=====rem>0
        CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='213' and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='213' and grp_acc='4' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Year=(select isnull(sum(Dep_Year),0) as Dep_Year  from Rpt_Grp  where Grp_No='213'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='213'  and grp_acc='4' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Month=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Grp_No='213'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='213'  and grp_acc='4' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(TTDep),0) as TTDep  from Rpt_Grp  where Grp_No='213'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='213'  and grp_acc='4' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Rem=(select isnull(sum(Remain),0) as Remain  from Rpt_Grp  where Grp_No='213'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='213' and grp_acc='4' ")
        '===========213=====rem=0
        CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='213' and Remain=0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='213'  and grp_acc='20' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(Amt),0) as TTDep  from Rpt_Grp  where Grp_No='213' and Remain=0 " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='213' and grp_acc='20' ")
        '===========214=====rem>0
        CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='214' and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='214' and grp_acc='5' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Year=(select isnull(sum(Dep_Year),0) as Dep_Year  from Rpt_Grp  where Grp_No='214'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='214'  and grp_acc='5' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Month=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Grp_No='214'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='214'  and grp_acc='5' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(TTDep),0) as TTDep  from Rpt_Grp  where Grp_No='214'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='214'  and grp_acc='5' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Rem=(select isnull(sum(Remain),0) as Remain  from Rpt_Grp  where Grp_No='214'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='214' and grp_acc='5' ")
        '===========214=====rem=0
        CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='214' and Remain=0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='214'  and grp_acc='21' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(Amt),0) as TTDep  from Rpt_Grp  where Grp_No='214' and Remain=0 " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='214' and grp_acc='21' ")
        '===========217=====rem>0
        CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='217' and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='217' and grp_acc='6' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Year=(select isnull(sum(Dep_Year),0) as Dep_Year  from Rpt_Grp  where Grp_No='217'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='217'  and grp_acc='6' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Month=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Grp_No='217'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='217'  and grp_acc='6' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(TTDep),0) as TTDep  from Rpt_Grp  where Grp_No='217'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='217'  and grp_acc='6' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Rem=(select isnull(sum(Remain),0) as Remain  from Rpt_Grp  where Grp_No='217'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='217' and grp_acc='6' ")
        '===========217=====rem=0
        CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='217' and Remain=0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='217'  and grp_acc='22' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(Amt),0) as TTDep  from Rpt_Grp  where Grp_No='217' and Remain=0 " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='217' and grp_acc='22' ")
        '===========2181=====rem>0
        CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='2181' and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2181' and grp_acc='7' ")
        Dim SOM As String = "update Rpt_NEW set Rpt_NEW.Dep_Year=(select isnull(sum(Dep_Year),0) as Dep_Year  from Rpt_Grp  where Grp_No='2181'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2181'  and grp_acc='7' "
        CNN.Execute(SOM)
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Month=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Grp_No='2181'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2181'  and grp_acc='7' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(TTDep),0) as TTDep  from Rpt_Grp  where Grp_No='2181'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2181'  and grp_acc='7' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Rem=(select isnull(sum(Remain),0) as Remain  from Rpt_Grp  where Grp_No='2181'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2181' and grp_acc='7' ")
        '===========2181=====rem=0
        CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='2181' and Remain=0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2181'  and grp_acc='23' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(Amt),0) as TTDep  from Rpt_Grp  where Grp_No='2181' and Remain=0 " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2181' and grp_acc='23' ")
        '===========2182=====rem>0
        CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='2182' and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2182' and grp_acc='8' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Year=(select isnull(sum(Dep_Year),0) as Dep_Year  from Rpt_Grp  where Grp_No='2182'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2182'  and grp_acc='8' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Month=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Grp_No='2182'  and Remain>0  " & Sec & " " & SDep & " " & KK & "  ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2182'  and grp_acc='8' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(TTDep),0) as TTDep  from Rpt_Grp  where Grp_No='2182'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2182'  and grp_acc='8' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Rem=(select isnull(sum(Remain),0) as Remain  from Rpt_Grp  where Grp_No='2182'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2182' and grp_acc='8' ")
        '===========2182=====rem=0
        CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='2182' and Remain=0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2182'  and grp_acc='24' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(Amt),0) as TTDep  from Rpt_Grp  where Grp_No='2182' and Remain=0 " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2182' and grp_acc='24' ")
        '===========2183=====rem>0
        CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='2183' and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2183' and grp_acc='9' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Year=(select isnull(sum(Dep_Year),0) as Dep_Year  from Rpt_Grp  where Grp_No='2183'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2183'  and grp_acc='9' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Month=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Grp_No='2183'  and Remain>0  " & Sec & " " & SDep & " " & KK & "  ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2183'  and grp_acc='9' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(TTDep),0) as TTDep  from Rpt_Grp  where Grp_No='2183'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2183'  and grp_acc='9' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Rem=(select isnull(sum(Remain),0) as Remain  from Rpt_Grp  where Grp_No='2183'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2183' and grp_acc='9' ")
        '===========2183=====rem=0
        CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt  from Rpt_Grp  where Grp_No='2183' and Remain=0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2183'  and grp_acc='25' ")
        CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(Amt),0) as TTDep  from Rpt_Grp  where Grp_No='2183' and Remain=0 " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2183' and grp_acc='25' ")
    End Sub
    Private Sub KKKDDD()
        Dim Sec, SDep As String

        If cmbSec.SelectedIndex = 0 Then
            Sec = ""
        Else
            Sec = " AND company='" & txtSec.Text & "' "
        End If
        If cmbDeprt.SelectedIndex = 0 Then
            SDep = ""
        Else
            SDep = " AND DepartmentID='" & txtDep.Text & "' "
        End If
        Dim KK As String
        If CheckBox4.Checked = True Then
            Dim i As Integer
            For i = 1 To FGIT.Rows - 1
                If FGIT.get_ValueMatrix(i, 1) = True Then
                    Dim kq As String = "UPDATE Sections set Choose='1'  where SecID='" & FGIT.get_TextMatrix(i, 2) & "'  "
                    CNN.Execute(kq)
                End If
            Next



            CNN.Execute("UPDATE Rpt_Grp set Rpt_Grp.Choose=Sections.Choose from Rpt_Grp,Sections where Sections.SecID=Rpt_Grp.Section ")
            KK = " AND Choose=1 "
        Else
            KK = ""
        End If
        CNN.Execute("UPDATE RPT_DETAILL_sum set amt=0,amt1=0,amt2=0,amt3=0,amt4=0,amt5=0,amt6=0,amt7=0 ")
        ''===============280===================
        'Dim BE3 As String = "update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(prevdep),0) as prevdep  from Rpt_Grp  where Group_ID='217' " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '         "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Grp_Acc='I' "
        'CNN.Execute(BE3)
        Dim kka280 As String = "update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Group_ID='280' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Group_ID=Rpt_Grp.Group_ID and RPT_DETAILL_sum.Grp_No='280' and RPT_DETAILL_sum.Grp_Acc='01'"
        CNN.Execute(kka280)
        CNN.Execute("update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(mondep),0) as mondep  from Rpt_Grp  where Group_ID='280' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Group_ID=Rpt_Grp.Group_ID and RPT_DETAILL_sum.Grp_No='280' and RPT_DETAILL_sum.Grp_Acc='09'")

        ''===============212===================
        'Dim BE3 As String = "update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(prevdep),0) as prevdep  from Rpt_Grp  where Group_ID='217' " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '         "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Grp_Acc='I' "
        'CNN.Execute(BE3)
        Dim kka12 As String = "update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Group_ID='212' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Group_ID=Rpt_Grp.Group_ID and RPT_DETAILL_sum.Grp_No='212' and RPT_DETAILL_sum.Grp_Acc='02'"
        CNN.Execute(kka12)
        CNN.Execute("update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(mondep),0) as mondep  from Rpt_Grp  where Group_ID='212' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Group_ID=Rpt_Grp.Group_ID and RPT_DETAILL_sum.Grp_No='212' and RPT_DETAILL_sum.Grp_Acc='10'")

        ''===============213===================
        'Dim BE3 As String = "update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(prevdep),0) as prevdep  from Rpt_Grp  where Group_ID='217' " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '         "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Grp_Acc='I' "
        'CNN.Execute(BE3)
        Dim kka13 As String = "update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Group_ID='213' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Group_ID=Rpt_Grp.Group_ID and RPT_DETAILL_sum.Grp_No='213' and RPT_DETAILL_sum.Grp_Acc='03'"
        CNN.Execute(kka13)
        CNN.Execute("update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(mondep),0) as mondep  from Rpt_Grp  where Group_ID='213' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Group_ID=Rpt_Grp.Group_ID and RPT_DETAILL_sum.Grp_No='213' and RPT_DETAILL_sum.Grp_Acc='11'")

        ''===============214===================
        'Dim BE3 As String = "update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(prevdep),0) as prevdep  from Rpt_Grp  where Group_ID='217' " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '         "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Grp_Acc='I' "
        'CNN.Execute(BE3)
        Dim kka14 As String = "update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Group_ID='214' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Group_ID=Rpt_Grp.Group_ID and RPT_DETAILL_sum.Grp_No='214' and RPT_DETAILL_sum.Grp_Acc='04'"
        CNN.Execute(kka14)
        CNN.Execute("update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(mondep),0) as mondep  from Rpt_Grp  where Group_ID='214' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Group_ID=Rpt_Grp.Group_ID and RPT_DETAILL_sum.Grp_No='214' and RPT_DETAILL_sum.Grp_Acc='12'")

        ''===============2181===================
        'Dim BE3 As String = "update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(prevdep),0) as prevdep  from Rpt_Grp  where Group_ID='217' " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '         "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Grp_Acc='I' "
        'CNN.Execute(BE3)
        Dim kka81 As String = "update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Group_ID='2181' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Group_ID=Rpt_Grp.Group_ID and RPT_DETAILL_sum.Grp_No='2181' and RPT_DETAILL_sum.Grp_Acc='05'"
        CNN.Execute(kka81)
        CNN.Execute("update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(mondep),0) as mondep  from Rpt_Grp  where Group_ID='2181' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Group_ID=Rpt_Grp.Group_ID and RPT_DETAILL_sum.Grp_No='2181' and RPT_DETAILL_sum.Grp_Acc='13'")

        ''===============2182===================
        'Dim BE3 As String = "update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(prevdep),0) as prevdep  from Rpt_Grp  where Group_ID='217' " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '         "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Grp_Acc='I' "
        'CNN.Execute(BE3)
        Dim kka82 As String = "update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Group_ID='2182' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Group_ID=Rpt_Grp.Group_ID and RPT_DETAILL_sum.Grp_No='2182' and RPT_DETAILL_sum.Grp_Acc='06'"
        CNN.Execute(kka82)
        CNN.Execute("update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(mondep),0) as mondep  from Rpt_Grp  where Group_ID='2182' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Group_ID=Rpt_Grp.Group_ID and RPT_DETAILL_sum.Grp_No='2182' and RPT_DETAILL_sum.Grp_Acc='14'")

        ''===============2183===================
        'Dim BE3 As String = "update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(prevdep),0) as prevdep  from Rpt_Grp  where Group_ID='217' " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '         "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Grp_Acc='I' "
        'CNN.Execute(BE3)
        Dim kka83 As String = "update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Group_ID='2183' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Group_ID=Rpt_Grp.Group_ID and RPT_DETAILL_sum.Grp_No='2183' and RPT_DETAILL_sum.Grp_Acc='07'"
        CNN.Execute(kka83)
        CNN.Execute("update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(mondep),0) as mondep  from Rpt_Grp  where Group_ID='2183' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Group_ID=Rpt_Grp.Group_ID and RPT_DETAILL_sum.Grp_No='2183' and RPT_DETAILL_sum.Grp_Acc='15'")

        ''===============217===================
        Dim BEF As String = "update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(prevdep),0) as prevdep  from Rpt_Grp  where Group_ID='217' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                 "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Grp_Acc='I' "
        CNN.Execute(BEF)
        Dim kka As String = "update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Group_ID='217' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Group_ID=Rpt_Grp.Group_ID and RPT_DETAILL_sum.Grp_No='217' and RPT_DETAILL_sum.Grp_Acc='08'"
        CNN.Execute(kka)
        CNN.Execute("update RPT_DETAILL_sum set RPT_DETAILL_sum.amt4=(select isnull(sum(mondep),0) as mondep  from Rpt_Grp  where Group_ID='217' " & Sec & " " & SDep & "  " & KK & " ) from " & _
                  "Rpt_Grp,RPT_DETAILL_sum where RPT_DETAILL_sum.Group_ID=Rpt_Grp.Group_ID and RPT_DETAILL_sum.Grp_No='217' and RPT_DETAILL_sum.Grp_Acc='16'")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Rem=(select isnull(sum(Remain),0) as Remain  from Rpt_Grp  where Grp_No='201' " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='201'")
        '===========203=====
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt_KIP),0) as Amt_KIP  from Rpt_Grp  where Grp_No='203' " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='203'")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Year=(select isnull(sum(Dep_Year),0) as Dep_Year  from Rpt_Grp  where Grp_No='203' " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='203'")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Month=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Grp_No='203' " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='203'")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(TTDep),0) as TTDep  from Rpt_Grp  where Grp_No='203' " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='203'")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Rem=(select isnull(sum(Remain),0) as Remain  from Rpt_Grp  where Grp_No='203' " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='203'")
        ''===========212=====rem>0
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='212' and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='212' and grp_acc='3' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Year=(select isnull(sum(Dep_Year),0) as Dep_Year  from Rpt_Grp  where Grp_No='212'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='212'  and grp_acc='3' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Month=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Grp_No='212'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='212'  and grp_acc='3' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(TTDep),0) as TTDep  from Rpt_Grp  where Grp_No='212'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='212'  and grp_acc='3' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Rem=(select isnull(sum(Remain),0) as Remain  from Rpt_Grp  where Grp_No='212'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='212' and grp_acc='3' ")
        ''===========212=====rem=0
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='212' and Remain=0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='212'  and grp_acc='19' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(Amt),0) as TTDep  from Rpt_Grp  where Grp_No='212' and Remain=0 " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='212' and grp_acc='19' ")
        ''===========213=====rem>0
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='213' and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='213' and grp_acc='4' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Year=(select isnull(sum(Dep_Year),0) as Dep_Year  from Rpt_Grp  where Grp_No='213'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='213'  and grp_acc='4' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Month=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Grp_No='213'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='213'  and grp_acc='4' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(TTDep),0) as TTDep  from Rpt_Grp  where Grp_No='213'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='213'  and grp_acc='4' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Rem=(select isnull(sum(Remain),0) as Remain  from Rpt_Grp  where Grp_No='213'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='213' and grp_acc='4' ")
        ''===========213=====rem=0
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='213' and Remain=0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='213'  and grp_acc='20' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(Amt),0) as TTDep  from Rpt_Grp  where Grp_No='213' and Remain=0 " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='213' and grp_acc='20' ")
        ''===========214=====rem>0
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='214' and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='214' and grp_acc='5' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Year=(select isnull(sum(Dep_Year),0) as Dep_Year  from Rpt_Grp  where Grp_No='214'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='214'  and grp_acc='5' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Month=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Grp_No='214'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='214'  and grp_acc='5' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(TTDep),0) as TTDep  from Rpt_Grp  where Grp_No='214'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='214'  and grp_acc='5' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Rem=(select isnull(sum(Remain),0) as Remain  from Rpt_Grp  where Grp_No='214'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='214' and grp_acc='5' ")
        ''===========214=====rem=0
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='214' and Remain=0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='214'  and grp_acc='21' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(Amt),0) as TTDep  from Rpt_Grp  where Grp_No='214' and Remain=0 " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='214' and grp_acc='21' ")
        ''===========217=====rem>0
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='217' and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='217' and grp_acc='6' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Year=(select isnull(sum(Dep_Year),0) as Dep_Year  from Rpt_Grp  where Grp_No='217'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='217'  and grp_acc='6' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Month=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Grp_No='217'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='217'  and grp_acc='6' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(TTDep),0) as TTDep  from Rpt_Grp  where Grp_No='217'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='217'  and grp_acc='6' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Rem=(select isnull(sum(Remain),0) as Remain  from Rpt_Grp  where Grp_No='217'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='217' and grp_acc='6' ")
        ''===========217=====rem=0
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='217' and Remain=0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='217'  and grp_acc='22' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(Amt),0) as TTDep  from Rpt_Grp  where Grp_No='217' and Remain=0 " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='217' and grp_acc='22' ")
        ''===========2181=====rem>0
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='2181' and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2181' and grp_acc='7' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Year=(select isnull(sum(Dep_Year),0) as Dep_Year  from Rpt_Grp  where Grp_No='2181'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2181'  and grp_acc='7' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Month=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Grp_No='2181'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2181'  and grp_acc='7' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(TTDep),0) as TTDep  from Rpt_Grp  where Grp_No='2181'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2181'  and grp_acc='7' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Rem=(select isnull(sum(Remain),0) as Remain  from Rpt_Grp  where Grp_No='2181'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2181' and grp_acc='7' ")
        ''===========2181=====rem=0
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='2181' and Remain=0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2181'  and grp_acc='23' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(Amt),0) as TTDep  from Rpt_Grp  where Grp_No='2181' and Remain=0 " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2181' and grp_acc='23' ")
        ''===========2182=====rem>0
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='2182' and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2182' and grp_acc='8' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Year=(select isnull(sum(Dep_Year),0) as Dep_Year  from Rpt_Grp  where Grp_No='2182'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2182'  and grp_acc='8' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Month=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Grp_No='2182'  and Remain>0  " & Sec & " " & SDep & " " & KK & "  ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2182'  and grp_acc='8' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(TTDep),0) as TTDep  from Rpt_Grp  where Grp_No='2182'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2182'  and grp_acc='8' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Rem=(select isnull(sum(Remain),0) as Remain  from Rpt_Grp  where Grp_No='2182'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2182' and grp_acc='8' ")
        ''===========2182=====rem=0
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='2182' and Remain=0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2182'  and grp_acc='24' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(Amt),0) as TTDep  from Rpt_Grp  where Grp_No='2182' and Remain=0 " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2182' and grp_acc='24' ")
        ''===========2183=====rem>0
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt_KIP  from Rpt_Grp  where Grp_No='2183' and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2183' and grp_acc='9' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Year=(select isnull(sum(Dep_Year),0) as Dep_Year  from Rpt_Grp  where Grp_No='2183'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2183'  and grp_acc='9' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_Month=(select isnull(sum(Dep_Month),0) as Dep_Month  from Rpt_Grp  where Grp_No='2183'  and Remain>0  " & Sec & " " & SDep & " " & KK & "  ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2183'  and grp_acc='9' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(TTDep),0) as TTDep  from Rpt_Grp  where Grp_No='2183'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2183'  and grp_acc='9' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Rem=(select isnull(sum(Remain),0) as Remain  from Rpt_Grp  where Grp_No='2183'  and Remain>0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2183' and grp_acc='9' ")
        ''===========2183=====rem=0
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.amt=(select isnull(sum(Amt),0) as Amt  from Rpt_Grp  where Grp_No='2183' and Remain=0  " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2183'  and grp_acc='25' ")
        'CNN.Execute("update Rpt_NEW set Rpt_NEW.Dep_TT=(select isnull(sum(Amt),0) as TTDep  from Rpt_Grp  where Grp_No='2183' and Remain=0 " & Sec & " " & SDep & "  " & KK & " ) from " & _
        '          "Rpt_Grp,Rpt_NEW where Rpt_NEW.Grp_No=Rpt_Grp.Grp_No and Rpt_NEW.Grp_No='2183' and grp_acc='25' ")

        CNN.Execute("UPDATE RPT_DETAILL_sum set amt4=(select isnull(sum(amt4),0) as amt4  from RPT_DETAILL_sum where (grp_acc>='01' and grp_acc<='08')) where grp_acc='II' ")
        CNN.Execute("UPDATE RPT_DETAILL_sum set amt4=(select isnull(sum(amt4),0) as amt4  from RPT_DETAILL_sum where (grp_acc>='09' and grp_acc<='16')) where grp_acc='III' ")
        CNN.Execute("UPDATE RPT_DETAILL_sum set amt4=(select isnull(sum(amt4),0) as amt4  from RPT_DETAILL_sum where (grp_acc>='09' and grp_acc<='16')) where grp_acc='III' ")
        CNN.Execute("UPDATE RPT_DETAILL_sum set amt=(amt1+amt2+amt3+amt4+amt5+amt6+amt7)  ")

    End Sub
    Private Sub DDD()
        Call Office()
        Dim Rs As New ADODB.Recordset
        CNN.Execute("update Rpt_Grp set Rpt_Grp.SectionNm=Sections.SecNmL from Sections,Rpt_Grp where Rpt_Grp.Section=Sections.SecID ")
        CNN.Execute("update Rpt_Grp set Rpt_Grp.DepartmenNm=Department.DepartmentNm from Department,Rpt_Grp where Rpt_Grp.DepartmentID=Department.DepartmentID ")
        CNN.Execute("update Rpt_Grp set Rpt_Grp.Asset_No=Assets.Asset_No  from Assets,Rpt_Grp where Rpt_Grp.AssetID=Assets.AssetID ")
        CNN.Execute("update Rpt_Grp set Rpt_Grp.Using_By=Assets.Using_By,Rpt_Grp.Budget=Assets.Budget from Assets,Rpt_Grp where Rpt_Grp.AssetID=Assets.AssetID ")
        CNN.Execute("update Rpt_Grp set Rpt_Grp.Group_Nm=Groups.Group_Nm from Groups,Rpt_Grp where Rpt_Grp.Group_ID=Groups.Group_ID ")

        If cmbSec.SelectedIndex = 0 Then
            Sec = ""
        Else
            Sec = " AND Section='" & txtSec.Text & "' "
        End If
        If cmbDeprt.SelectedIndex = 0 Then
            SDep = ""
        Else
            SDep = " AND DepartmentID='" & txtDep.Text & "' "
        End If
        Dim Post As String
        If CmbShow.SelectedIndex = 0 Then
            Post = ""
        ElseIf CmbShow.SelectedIndex = 1 Then
            Post = " AND  Deposted='0' "
        Else
            Post = " AND  Deposted='1'"
        End If
        Call LoadSqlData("Select * from Rpt_Grp WHERE 1=1 " & Sec & " " & SDep & " " & Post & " Order by Asset_No,AssetID ASC", Rs)
        Dim FrmPreview As New FmPreview : FrmClosing()
        rpt = New Crystal_Asset_VTE_Detaill
        Dim myText1, myText3 As CrystalDecisions.CrystalReports.Engine.TextObject
        myText1 = CType(rpt.ReportDefinition.ReportObjects.Item("txtGrp"), CrystalDecisions.CrystalReports.Engine.TextObject)
        If txtGrp.Text <> "" Then
            myText1.Text = Trim(cmbGrp.Text)
        End If

        '========================
        Dim mPer As String
        If optMon.Checked = True Then

            mPer = "ປະຈຳເດືອນ " & Format(DTMon.Value, "MM/yyyy")

        ElseIf optYear.Checked = True Then

            mPer = "ປະຈຳປີ " & Format(DTYear.Value, "yyyy")
        Else

            mPer = "ປະຈຳງວດ " & Trim(cmbTerm.Text) & " / " & Format(dtTerm.Value, "yyyy")
        End If
        '======================
        myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Tel"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText3.Text = "Tel: " & OffTel & " Fax: " & OffFax
        myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("DD"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText3.Text = mPer
        If Lang = False Then
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = OffName
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal5"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign5
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal4"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign4
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal3"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign3
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal2"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign2
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal1"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign1
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Place"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = PlaecL
        Else
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = OffNameE
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal5"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign5e
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal4"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign4e
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal3"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign3e
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal2"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign2e
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal1"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign1e
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Place"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = PlaecE
        End If
        If Rs.RecordCount = 0 Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
        With rpt
            Dim Rs1 As New ADODB.Recordset
            'Dim RO As CrystalDecisions.CrystalReports.Engine.ReportObject
            'Dim SRO As CrystalDecisions.CrystalReports.Engine.SubreportObject
            'Dim SubDoc As CrystalDecisions.CrystalReports.Engine.ReportDocument
            'SqlPrint = "SELECT  * from Ap_Image  where Img_Id='" & IMageID & "'"
            'Call LoadSqlData(SqlPrint, Rs1)
            'RO = rpt.ReportDefinition.Sections.Item("Section1").ReportObjects.Item("Subreport1")
            'SRO = CType(RO, CrystalDecisions.CrystalReports.Engine.SubreportObject)
            'SubDoc = SRO.OpenSubreport(SRO.SubreportName)
            'If Rs1.RecordCount > 0 Then
            '    SubDoc.SetDataSource(Rs1)
            '    FmPreview.ReportViewer.ReportSource = SubDoc
            'End If
            rpt.SetDataSource(Rs)
            rpt.Refresh()
            FmPreview.ReportViewer.ReportSource = rpt
            FmPreview.ReportViewer.DisplayGroupTree = False
            FmPreview.WindowState = FormWindowState.Maximized
            FmPreview.Show()
            'Call CloseRs(RSC)
            'Call CloseRs(Rs)
        End With

    End Sub
    Private Sub SUMDDD()

        Call Office()
        Dim RS As New ADODB.Recordset
        Dim FrmPreview As New FmPreview : FrmClosing()
        rpt = New Crystal_SummaryDetaill
        Call KKKDDD()
        Dim hh As String = "Select '" & PlaecE & "' as P ,'" & OffTel & "' as Tel, * from RPT_DETAILL_SUM order by cnt asc  "
        Call LoadSqlData(hh, RS)

        If RS.RecordCount = 0 Then
            MsgBox("NO DATA") : Exit Sub
        End If

        With rpt
            Dim myTextObjectOnReport As CrystalDecisions.CrystalReports.Engine.TextObject
            Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("OffNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = OffName
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Tel"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = "Tel: " & OffTel & "Fax: " & OffFax
            myTextObjectOnReport = CType(rpt.ReportDefinition.ReportObjects.Item("SG1"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myTextObjectOnReport.Text = Sign5
            myTextObjectOnReport = CType(rpt.ReportDefinition.ReportObjects.Item("SG2"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myTextObjectOnReport.Text = Sign3
            myTextObjectOnReport = CType(rpt.ReportDefinition.ReportObjects.Item("SG3"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myTextObjectOnReport.Text = Sign1
            myTextObjectOnReport = CType(rpt.ReportDefinition.ReportObjects.Item("PP"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myTextObjectOnReport.Text = PlaecL
            'myTextObjectOnReport = CType(rpt.ReportDefinition.ReportObjects.Item("HD"), CrystalDecisions.CrystalReports.Engine.TextObject)
            'myTextObjectOnReport.Text = HHDD
            'myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("AAA"), CrystalDecisions.CrystalReports.Engine.TextObject)
            'myText2.Text = PlaecE
            'Dim RO As CrystalDecisions.CrystalReports.Engine.ReportObject
            'Dim SRO As CrystalDecisions.CrystalReports.Engine.SubreportObject
            'Dim SubDoc As CrystalDecisions.CrystalReports.Engine.ReportDocument
            'SqlPrint = "SELECT  * from Ap_Image  where Img_Id='" & IMageID & "'"
            'Call LoadSqlData(SqlPrint, Rs1)
            'RO = rpt.ReportDefinition.Sections.Item("Section1").ReportObjects.Item("Subreport3")
            'SRO = CType(RO, CrystalDecisions.CrystalReports.Engine.SubreportObject)
            'SubDoc = SRO.OpenSubreport(SRO.SubreportName)
            'If Rs1.RecordCount > 0 Then
            '    SubDoc.SetDataSource(Rs1)
            '    FmPreview.ReportViewer.ReportSource = SubDoc
            'End If
            rpt.SetDataSource(RS)
            rpt.Refresh()
            FmPreview.ReportViewer.ReportSource = rpt
            FmPreview.ReportViewer.DisplayGroupTree = False
            FmPreview.WindowState = FormWindowState.Maximized
            FmPreview.Show()
            'Call CloseRs(RSC)
            'Call CloseRs(RS)

        End With
    End Sub
    Private Sub OfficeNEW()
        Dim Rs As New ADODB.Recordset
        With Rs
            Call LoadSqlData("SELECT * FROM AP_Office where off_id='" & txtSec.Text & "' ", Rs)
            If .RecordCount = 0 Then Exit Sub
            OffName = Trim(.Fields("off_Name").Value.ToString)
            OffNameE = Trim(.Fields("off_NameE").Value.ToString)
            'Off_strtl = Trim(.Fields("off_strtl").Value.ToString)
            'Off_VillageL = Trim(.Fields("Off_VillageL").Value.ToString)
            'Off_DistL = Trim(.Fields("Off_DistL").Value.ToString)
            'Off_ProVL = Trim(.Fields("Off_ProVL").Value.ToString)
            OffTel = Trim(.Fields("tel").Value.ToString)
            OffFax = Trim(.Fields("fax").Value.ToString)
            'Sign1 = Trim(.Fields("Signatur_1").Value.ToString)
            'Sign2 = Trim(.Fields("Signatur_2").Value.ToString)
            'Sign3 = Trim(.Fields("Signatur_3").Value.ToString)
            'Sign4 = Trim(.Fields("Signatur_4").Value.ToString)
            'Sign5 = Trim(.Fields("Signatur_5").Value.ToString)
            'OffPlace = Trim(.Fields("Locate_Bill").Value.ToString)
            .MoveNext()
        End With
    End Sub
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Call Office()
        OfficeNEW()
        LoadHeader_Asset()
        Dim Rs, Rs1 As New ADODB.Recordset
        Dim mPer As String
        Dim rpt As Object

        Dim FrmPreview As New FmPreview : FrmClosing()
        If chkBranch.Checked = True Then
            MsgBox("No Report") : Exit Sub
        End If
        If optMon.Checked = True Then
            If Lang = True Then
                mPer = "Monthly " & Format(DTMon.Value, "MM/yyyy")
            Else
                mPer = "ປະຈຳເດືອນ " & Format(DTMon.Value, "MM/yyyy")
            End If
            MM = "ຫຼຸ້ຍຫ້ຽນໃນ ເດືອນ"
        ElseIf optTerm.Checked = True Then
            If Lang = True Then
                mPer = "Termly " & cmbTerm.Text & "/ " & Format(DTYear.Value, "yyyy")
            Else
                mPer = "ປະຈຳງວດ " & cmbTerm.Text & "/ " & Format(DTYear.Value, "yyyy")
            End If
            MM = "ຫຼຸ້ຍຫ້ຽນໃນ ງວດ"
        Else
            If Lang = True Then
                mPer = "Yearly " & Format(DTYear.Value, "yyyy")
            Else
                mPer = "ປະຈຳປີ " & Format(DTYear.Value, "yyyy")
            End If
            MM = "ຫຼຸ້ຍຫ້ຽນໃນ ປີ"
        End If

        If optMon.Checked = True Then
            'CalcMon()
            Call Calc()
        ElseIf optYear.Checked = True Then
            Call CalcYEAR()

        Else
            Call CalcTerm()
        End If
        CNN.Execute("Update Rpt_Grp set Rpt_Grp.Company=Assets.Company ,Rpt_Grp.Section=Assets.Section ,Rpt_Grp.DepartmentID=Assets.DepartmentID  from Rpt_Grp,Assets where Assets.AssetID=Rpt_Grp.AssetID ")

        If CheckBox1.Checked = True Then
            Call DDD()
            Exit Sub
        End If
        If CheckBox2.Checked = True Then

            Call SUMMARYASSET()
            Dim kk As String = "UPDATE Sections set Choose='0' "
            CNN.Execute(kk)
            Exit Sub
        End If
        If CheckBox5.Checked = True Then
            Call SUMDDD()
            Exit Sub
        End If
        If cmbSec.SelectedIndex = 0 Then
            Sec = ""
        Else
            Sec = " AND Section='" & txtSec.Text & "' "
        End If
        If cmbDeprt.SelectedIndex = 0 Then
            SDep = ""
        Else
            SDep = " AND DepartmentID='" & txtDep.Text & "' "
        End If
        Dim Post As String
        If CmbShow.SelectedIndex = 0 Then
            Post = ""
        ElseIf CmbShow.SelectedIndex = 1 Then
            Post = " AND  Deposted='0' "
        Else
            Post = " AND  Deposted='1'"
        End If

        If chkSum.Checked = False Then
            Call LoadSqlData("Select AssetID, Asset_No, Asset_Nm,Asset_NmE, Group_ID, Date_Work, Used_Life, Amt_KIP, Broke_Date, Deposted_Date, Dep_Year, Dep_Month, TTMon, PrevMon, PrevDep, MonDep, TTDep, Remain from Rpt_Grp WHERE 1=1 " & Sec & " " & SDep & " " & Post & " Order by Asset_No,AssetID ASC", Rs)
            If Lang = True Then
                rpt = New CryRpt_GrpEng
            Else
                'If CheckBox1.Checked = True Then
                '    rpt = New Crystal_Asset_VTE_Detaill
                'Else
                rpt = New CryRpt_Grp
                'End If
                Dim myText4 As CrystalDecisions.CrystalReports.Engine.TextObject
                myText4 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText4.Text = OffName
            End If


        Else
            If cmbSec.SelectedIndex = 0 Then
                Sec = ""
            Else
                Sec = " AND A.Section='" & txtSec.Text & "' "
            End If
            If cmbDeprt.SelectedIndex = 0 Then
                SDep = ""
            Else
                SDep = " AND A.DepartmentID='" & txtDep.Text & "' "
            End If
            Dim ok As String = "Select A.assetid,B.Ac_Code,A.asset_no, B.Group_Nm,B.Group_NmE, A.Group_ID,A.asset_Nm, sum(A.Amt_KIP) as Amt_KIP, sum(A.Dep_Year) as Dep_Year, sum(A.Dep_Month) as Dep_Month, sum(A.PrevMon) as PrevMon, sum(A.PrevDep) as PrevDep, sum(A.MonDep) as MonDep, sum(A.TTDep) as TTDep, sum(A.Remain) as Remain, A.Used_Life, A.Date_Work " & _
                          " from Rpt_Grp A INNER JOIN Groups B ON A.Group_ID=B.Group_ID  WHERE 1=1 " & Sec & " " & SDep & " " & Post & " GROUP By A.Group_ID, B.Ac_Code, B.Group_Nm,B.Group_NmE, A.Used_Life, A.Date_Work,asset_Nm,assetid,asset_no  ORDER by A.Group_ID ASC "
            Call LoadSqlData(ok, Rs)
            If Lang = True Then
                rpt = New CryRpt_GrpSumEng
            Else
                rpt = New CryRpt_GrpSum
                Dim myText4 As CrystalDecisions.CrystalReports.Engine.TextObject
                myText4 = CType(rpt.ReportDefinition.ReportObjects.Item("MM"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText4.Text = MM
                If cmbSec.SelectedIndex = 0 Then
                    myText4 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText4.Text = OffName
                Else
                    myText4 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText4.Text = cmbSec.Text
                End If
                If cmbDeprt.SelectedIndex = 0 Then
                    myText4 = CType(rpt.ReportDefinition.ReportObjects.Item("Text17"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText4.Text = cmbDeprt.Text
                Else
                    myText4 = CType(rpt.ReportDefinition.ReportObjects.Item("Text17"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText4.Text = cmbDeprt.Text
                End If
            End If
        End If
        Call Office()
        Dim myText1, myText2, myText3 As CrystalDecisions.CrystalReports.Engine.TextObject
        myText1 = CType(rpt.ReportDefinition.ReportObjects.Item("txtGrp"), CrystalDecisions.CrystalReports.Engine.TextObject)
        If txtGrp.Text <> "" Then
            myText1.Text = Trim(cmbGrp.Text)
        End If
        myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("txtPeriod"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = mPer
        myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Tel"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText3.Text = "Tel: " & OffTel & " Fax: " & OffFax
        If Lang = False Then

            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal5"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign5
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal4"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign4
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal3"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign3
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal2"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign2
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal1"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign1
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Place"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = PlaecL
        Else
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = OffNameE
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal5"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign5e
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal4"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign4e
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal3"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign3e
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal2"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign2e
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal1"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = Sign1e
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Place"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = PlaecE
        End If
        If Rs.RecordCount = 0 Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
        With rpt
            'Dim RO As CrystalDecisions.CrystalReports.Engine.ReportObject
            'Dim SRO As CrystalDecisions.CrystalReports.Engine.SubreportObject
            'Dim SubDoc As CrystalDecisions.CrystalReports.Engine.ReportDocument
            'SqlPrint = "SELECT  * from Ap_Image  where Img_Id='" & IMageID & "'"
            'Call LoadSqlData(SqlPrint, Rs1)
            'RO = rpt.ReportDefinition.Sections.Item("Section1").ReportObjects.Item("Subreport3")
            'SRO = CType(RO, CrystalDecisions.CrystalReports.Engine.SubreportObject)
            'SubDoc = SRO.OpenSubreport(SRO.SubreportName)
            'If Rs1.RecordCount > 0 Then
            '    SubDoc.SetDataSource(Rs1)
            '    FmPreview.ReportViewer.ReportSource = SubDoc
            'End If
            rpt.SetDataSource(Rs)
            rpt.Refresh()
            FmPreview.ReportViewer.ReportSource = rpt
            FmPreview.ReportViewer.DisplayGroupTree = False
            FmPreview.WindowState = FormWindowState.Maximized
            FmPreview.Show()
            'Call CloseRs(RSC)
            'Call CloseRs(Rs)
        End With
    End Sub

    Private Sub Calc()
        Dim cRS As New ADODB.Recordset
        Dim Str As String = ""
        Dim ss As String
        Dim strDate, EndDate As Date
        Dim mym, myy, myr As Integer
        mym = DTMon.Value.Month + 1
        myy = DTMon.Value.Year
        Dim yy As Integer = 1
        If optYear.Checked = True Then
            yy = 12
            myr = DTYear.Value.Year
        ElseIf optMon.Checked = True Then
            myr = DTMon.Value.Year
        Else
            myr = dtTerm.Value.Year
        End If
        Dim mm As Integer = 0
        Dim mTm As Integer = 0
        mm = Month(DTMon.Value) - 1
        strDate = CDate("01/01/" & DTMon.Value.Year.ToString)
        'EndDate = CDate("31/12/" & Trim(DTYear.Value.Year.ToString))
        If DTMon.Value.Month < 12 Then
            EndDate = DateAdd("d", -1, CDate("01/" & DTMon.Value.Month + 1 & "/" & DTMon.Value.Year))
        Else
            EndDate = CDate("31/12/" & DTMon.Value.Year)
        End If


        CNN.Execute("Delete From Rpt_Grp")

        ss = "Insert into Rpt_Grp(AmountRemain,AmountClear,Amt_All,AssetID, Asset_No, Asset_Nm,Asset_NmE, Group_ID, Date_Work, Used_Life, Amt, Amt_KIP, Broke_Date,  Deposted, Deposted_Date, Dep_Year, Dep_Month, TTMon, PrevMon, PrevDep, CurrMon, MonDep, TTDep, Remain, strDate, EndDate) " & _
               "Select AmountRemain,AmountClear,Amt_All,AssetID, Asset_No, Asset_Nm,Asset_NmE, Group_ID, Date_Work, Used_Life, Amt, Amt, Broke_Date,  Deposted, Deposted_Date, Dep_Year, Dep_Month, 0, 0, 0, 0, 0, 0, 0, '" & Format(strDate, "yyyy-MM-dd") & "', '" & Format(EndDate, "yyyy-MM-dd") & "' From Assets Where Date_Work <= '" & Format(EndDate, "yyyy-MM-dd") & "' "
        ss = ss & Str & " Order by Asset_No "
        CNN.Execute(ss)



        'CNN.Execute("Update Rpt_Grp Set TTMon=Used_Life * 12, PrevMon=DateDiff(m, Date_Work, EndDate) , CurrMon=DateDiff(m, Date_Work, strDate) ")
        'CNN.Execute("Update Rpt_Grp Set CurrMon=DateDiff(m, strDate, Deposted_Date)+1 Where Deposted_Date < EndDate AND Deposted_Date is not null")
        CNN.Execute("Update Rpt_Grp Set TTMon=Used_Life * 12, PrevMon=DateDiff(m, Date_Work, EndDate)-1   , CurrMon=DateDiff(m, Date_Work, EndDate) ")
        CNN.Execute("Update Rpt_Grp Set CurrMon=DateDiff(m, strDate, Deposted_Date)   Where Deposted_Date < EndDate AND Deposted_Date is not null")
        '=============
        CNN.Execute("Update Rpt_Grp Set PrevMon=PrevMon+1 where year(Date_Work)<'2016'  ")
        '================

        'CNN.Execute("Update Rpt_Grp Set CurrMon=1 Where CurrMon=0")
        CNN.Execute("Update Rpt_Grp Set PrevMon=0 Where PrevMon < 0")
        CNN.Execute("Update Rpt_Grp Set MonDep= Dep_Month Where PrevMon > 0 AND TTMon <> PrevMon")
        CNN.Execute("Update Rpt_Grp Set MonDep= Dep_Month Where DateDiff(m, Date_Work, EndDate) >= 0 ")
        CNN.Execute("Update Rpt_Grp Set MonDep= 0 Where TTMon=PrevMon ")
        CNN.Execute("Update Rpt_Grp Set MonDep = 0 Where CurrMon=0 ")
        '--------------Depost
        CNN.Execute("Update Rpt_Grp Set PrevMon=DateDiff(m, Date_Work, Deposted_Date)  Where Deposted_Date is not null AND PrevMon > DateDiff(m, Date_Work, Deposted_Date) ")
        CNN.Execute("Update Rpt_Grp Set MonDep= 0 Where Deposted_Date is not null AND DateDiff(m, strdate, Deposted_Date) <= 0 ")
        CNN.Execute("Update Rpt_Grp Set Remain = 0 Where Deposted_Date is not null AND DateDiff(m, strdate, Deposted_Date) <= 0 ")
        CNN.Execute("Update Rpt_Grp Set PrevMon=TTMon where PrevMon > TTMon  ")
        CNN.Execute("Update Rpt_Grp Set PrevDep= PrevMon * Dep_Month")
        CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep Where Deposted_Date is not null AND DateDiff(m, strdate, Deposted_Date) <= 0 ")
        '------------ LA Only
        'CNN.Execute("Update Rpt_Grp Set Rpt_Grp.PrevDep=Open_BL.Open_Amt From Rpt_Grp, Open_BL Where Rpt_Grp.AssetID=Open_BL.AssetID AND Open_BL.Open_Amt <>0 and Year(Open_BL.Date_Work)=" & myr & "")
        '------------
        'CNN.Execute("Update Rpt_Grp Set MonDep =  MonDep * " & yy & " where CurrMon > 12")
        ''CNN.Execute("Update Rpt_Grp Set MonDep =  MonDep * CurrMon where CurrMon <= 12")
        'CNN.Execute("Update Rpt_Grp Set MonDep =  MonDep * CurrMon where CurrMon <= 12")
        CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep ")
        CNN.Execute("Update Rpt_Grp Set MonDep = Amt_KIP-PrevDep where TTDep > Amt_KIP ")
        CNN.Execute("Update Rpt_Grp Set MonDep = 0 where MonDep < 0 ")
        CNN.Execute("Update Rpt_Grp Set Remain = Amt_KIP-TTDep ")
        CNN.Execute("Update Rpt_Grp Set Remain = 0 where Remain < 1 ")
        CNN.Execute("Update Rpt_Grp Set  Amt_KIP=0, MonDep=0, TTDep=0,PrevDep=0 Where Deposted_Date <= strdate")
        CNN.Execute("Update Rpt_Grp Set Remain = 0, Amt_KIP=0,MonDep=0,TTDep=0,PrevDep=0 Where Deposted_Date <= EndDate")
        'CNN.Execute("Update Rpt_Grp Set PrevDep = 0 Where DateDiff(m, Date_Work, EndDate) <=1 ")

        CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep ")
        CNN.Execute("Update Rpt_Grp Set Remain = Amt_KIP-TTDep ")

        CNN.Execute("Update Rpt_Grp Set Remain =0 where  Amt_KIP<TTDep ")


    End Sub
    Private Sub CalcYEAR()
        Dim cRS As New ADODB.Recordset
        Dim Str As String = ""
        Dim ss As String
        Dim strDate, EndDate As Date
        Dim myr As Integer
        Dim yy As Integer = 1
        yy = 12
        myr = DTYear.Value.Year
        strDate = CDate("01/01/" & DTYear.Value.Year)
        EndDate = CDate("31/12/" & DTYear.Value.Year)
        Dim EndPreyear As Date = CDate("31/12/" & DTYear.Value.Year - 1)

        If txtGrp.Text <> "" Then
            Str = " AND Group_ID='" & Trim(txtGrp.Text) & "'"
        End If


        CNN.Execute("Delete From Rpt_Grp")
        ss = "Insert into Rpt_Grp(AmountRemain,AmountClear,Amt_All,AssetID, Asset_No, Asset_Nm,Asset_NmE, Group_ID, Date_Work, Used_Life, Amt, Amt_KIP, Broke_Date,  Deposted, Deposted_Date, Dep_Year, Dep_Month, TTMon, PrevMon, PrevDep, CurrMon, MonDep, TTDep, Remain, strDate, EndDate) " & _
                    "Select AmountRemain,AmountClear,Amt_All,AssetID, Asset_No, Asset_Nm,Asset_NmE, Group_ID, Date_Work, Used_Life, Amt, Amt, Broke_Date,  Deposted, Deposted_Date, Dep_Year, Dep_Month, 0, 0, 0, 0, 0, 0, 0, '" & Format(strDate, "yyyy-MM-dd") & "', '" & Format(EndDate, "yyyy-MM-dd") & "' From Assets Where year(Date_Work) <= '" & Format(EndDate, "yyyy") & "' "
        ss = ss & Str & " Order by Asset_No "
        CNN.Execute(ss)
        'CurrMon=12 
        CNN.Execute("Update Rpt_Grp Set Preyear='" & Format(EndPreyear, "yyyy-MM-dd") & "' ")
        CNN.Execute("Update Rpt_Grp Set TTMon=Used_Life * 12, PrevMon=DateDiff(m, Date_Work,Preyear)   , CurrMon=DateDiff(m, Date_Work, EndDate)  ")

        'CNN.Execute("Update Rpt_Grp Set TTMon=Used_Life * 12, PrevMon=DateDiff(m, Date_Work,strDate)   , CurrMon=DateDiff(m, Date_Work, EndDate)  ")
        CNN.Execute("Update Rpt_Grp Set CurrMon=0 where Deposted_Date < strDate   ")
        CNN.Execute("Update Rpt_Grp Set CurrMon=DateDiff(m, strDate, Deposted_Date)+1 Where Deposted_Date < EndDate AND Deposted_Date >= strDate  and Deposted_Date is not null")
        CNN.Execute("Update Rpt_Grp Set CurrMon=DateDiff(m, Date_Work, EndDate)   Where Date_Work > strDate ")
        'CNN.Execute("Update Rpt_Grp Set CurrMon=DateDiff(m, Date_Work, EndDate)+1  Where Date_Work > strDate ")
        ''=============
        'CNN.Execute("Update Rpt_Grp Set PrevMon=DateDiff(m, Date_Work, strDate)-1  where assetid='FA16.01971' ")
        'CNN.Execute("Update Rpt_Grp Set PrevMon=DateDiff(m, Date_Work, strDate)-1  where assetid='FA16.01966' ")
        'CNN.Execute("Update Rpt_Grp Set TTMon=Used_Life * 12, PrevMon=DateDiff(m, Date_Work, strdate) , CurrMon=DateDiff(m, Date_Work, EndDate) ")
        'CNN.Execute("Update Rpt_Grp Set CurrMon=DateDiff(m, strDate, Deposted_Date)+1 Where Deposted_Date < EndDate AND Deposted_Date is not null")
        '=============
        CNN.Execute("Update Rpt_Grp Set PrevMon=PrevMon+1 where year(Date_Work)<'2016' ")
        'CNN.Execute("Update Rpt_Grp Set PrevMon=PrevMon+1 where year(Date_Work)<'2016' and  group_id<>'217' ")
        CNN.Execute("Update Rpt_Grp Set PrevMon=PrevMon-1 where month(Date_Work)='12'  and year(Date_Work)='2015'  and  group_id='217' ")
        CNN.Execute("Update Rpt_Grp Set PrevMon=PrevMon-1 where month(Date_Work)='12'  and year(Date_Work)='2015'  and  group_id='2182' ")
        '================

        'CNN.Execute("Update Rpt_Grp Set CurrMon=1 Where CurrMon=0")

        CNN.Execute("Update Rpt_Grp Set PrevMon=0 Where PrevMon < 0")
        CNN.Execute("Update Rpt_Grp Set MonDep= Dep_Month Where PrevMon > 0 AND TTMon <> PrevMon")
        CNN.Execute("Update Rpt_Grp Set MonDep= Dep_Month Where DateDiff(m, Date_Work, EndDate) >= 0 ")
        CNN.Execute("Update Rpt_Grp Set MonDep= 0 Where TTMon=PrevMon ")
        CNN.Execute("Update Rpt_Grp Set MonDep = 0 Where CurrMon=0 ")
        '--------------Depost
        CNN.Execute("Update Rpt_Grp Set PrevMon=DateDiff(m, Date_Work, Deposted_Date)  Where Deposted_Date is not null AND PrevMon > DateDiff(m, Date_Work, Deposted_Date) ")

        CNN.Execute("Update Rpt_Grp Set MonDep= 0 Where Deposted_Date is not null AND DateDiff(m, strdate, Deposted_Date) <= 0 ")
        CNN.Execute("Update Rpt_Grp Set Remain = 0 Where Deposted_Date is not null AND DateDiff(m, strdate, Deposted_Date) <= 0 ")
        CNN.Execute("Update Rpt_Grp Set PrevMon=TTMon where PrevMon > TTMon  ")
        CNN.Execute("Update Rpt_Grp Set PrevDep= PrevMon * Dep_Month")
        CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep Where Deposted_Date is not null AND DateDiff(m, strdate, Deposted_Date) <= 0 ")
        '------------ LA Only
        'CNN.Execute("Update Rpt_Grp Set Rpt_Grp.PrevDep=Open_BL.Open_Amt From Rpt_Grp, Open_BL Where Rpt_Grp.AssetID=Open_BL.AssetID AND Open_BL.Open_Amt <>0 and Year(Open_BL.Date_Work)=" & myr & "")
        '------------
        CNN.Execute("Update Rpt_Grp Set MonDep =  MonDep * " & yy & " where CurrMon > 12")
        ''CNN.Execute("Update Rpt_Grp Set MonDep =  MonDep * CurrMon where CurrMon <= 12")
        CNN.Execute("Update Rpt_Grp Set MonDep =  MonDep * CurrMon where CurrMon <= 12")
        CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep ")
        CNN.Execute("Update Rpt_Grp Set MonDep = Amt_KIP-PrevDep where TTDep > Amt_KIP ")
        CNN.Execute("Update Rpt_Grp Set MonDep = 0 where MonDep < 0 ")
        CNN.Execute("Update Rpt_Grp Set Remain = Amt_KIP-TTDep ")
        CNN.Execute("Update Rpt_Grp Set Remain = 0 where Remain < 1 ")
        CNN.Execute("Update Rpt_Grp Set  Amt_KIP=0, MonDep=0, TTDep=0,PrevDep=0 Where Deposted_Date <= strdate")
        CNN.Execute("Update Rpt_Grp Set Remain = 0, Amt_KIP=0,MonDep=0,TTDep=0,PrevDep=0 Where Deposted_Date <= EndDate")
        'CNN.Execute("Update Rpt_Grp Set PrevDep = 0 Where DateDiff(m, Date_Work, EndDate) <=1 ")
        If optTerm.Checked = True Then
            CNN.Execute("Update Rpt_Grp Set MonDep = 0 where CurrMon=0 ")
        End If
        CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep ")
        CNN.Execute("Update Rpt_Grp Set Remain = Amt_KIP-TTDep ")

        CNN.Execute("Update Rpt_Grp Set Remain =0 where  Amt_KIP<TTDep ")

    End Sub
    Private Sub CalcMon()
        Dim cRS As New ADODB.Recordset
        Dim Str As String = ""
        Dim ss As String
        Dim strDate, EndDate As Date
        Dim mym, myy, myr As Integer
        mym = DTMon.Value.Month + 1
        myy = DTMon.Value.Year
        Dim yy As Integer = 1
        myr = DTMon.Value.Year
        Dim mm As Integer = 0
        Dim mTm As Integer = 0
        'mm = Month(DTMon.Value)
        mm = Month(DTMon.Value) - 1
        strDate = CDate("01/" & Trim(DTMon.Value.Month.ToString) & "/" & Trim(DTMon.Value.Year.ToString))
        If DTMon.Value.Month < 12 Then
            EndDate = DateAdd("d", -1, CDate("01/" & DTMon.Value.Month + 1 & "/" & DTMon.Value.Year))
        Else
            EndDate = CDate("31/12/" & DTMon.Value.Year)
        End If

 

        'MsgBox(strDate)
        'MsgBox(EndDate)
        If txtGrp.Text <> "" Then
            Str = " AND Group_ID='" & Trim(txtGrp.Text) & "'"
        End If

        CNN.Execute("Delete From Rpt_Grp")
        'ss = "Insert into Rpt_Grp(AmountRemain,AmountClear,Amt_All,AssetID, Asset_No, Asset_Nm, Group_ID, Date_Work, Used_Life, Amt, Amt_KIP, Broked,Broke_Date, Deposted, Deposted_Date, Dep_Year, Dep_Month, TTMon, PrevMon, PrevDep, CurrMon, MonDep, TTDep, Remain, strDate, EndDate) " & _
        '    "Select AmountRemain,AmountClear,Amt_All,AssetID, Asset_No, Asset_Nm, Group_ID, Date_Work, Used_Life, Amt, Amt_KIP,Broked, Broke_Date, Deposted, Deposted_Date, Dep_Year, Dep_Month, 0, 0, 0, 0, 0, 0, 0, '" & Format(strDate, "yyyy-MM-dd") & "', '" & Format(EndDate, "yyyy-MM-dd") & "' From Assets Where Date_Work <= '" & Format(EndDate, "yyyy-MM-dd") & "' "
        'ss = ss & Str & " Order by Asset_No "
        'CNN.Execute(ss)

        ss = "Insert into Rpt_Grp(AmountRemain,AmountClear,Amt_All,AssetID, Asset_No, Asset_Nm, Group_ID, Date_Work, Used_Life, Amt, Amt_KIP, Broked,Broke_Date, Deposted, Deposted_Date, Dep_Year, Dep_Month,Dep_Day, TTMon, PrevMon, PrevDep, CurrMon, MonDep, TTDep, Remain, strDate, EndDate) " & _
      "Select AmountRemain,AmountClear,Amt_All,AssetID, Asset_No, Asset_Nm, Group_ID, Date_Work+1, Used_Life, Amt, Amt_KIP,Broked, Broke_Date, Deposted, Deposted_Date, Dep_Year, Dep_Month,Dep_Day, 0, 0, 0, 0, 0, 0, 0, '" & Format(strDate, "yyyy-MM-dd") & "', '" & Format(EndDate, "yyyy-MM-dd") & "' From Assets Where Date_Work <= '" & Format(EndDate, "yyyy-MM-dd") & "' "
        ss = ss & Str & " Order by Asset_No "
        CNN.Execute(ss)

        'CNN.Execute("Update Rpt_Grp Set TTMon=Used_Life * 12, PrevMon=DateDiff(m, Date_Work, strDate)-1, CurrMon=DateDiff(m, Date_Work, EndDate) ")
        'CNN.Execute("Update Rpt_Grp Set CurrMon=DateDiff(m, strDate, Deposted_Date)+1 Where Deposted_Date < EndDate AND Deposted_Date is not null")
        'CNN.Execute("Update Rpt_Grp Set PrevMon=0 Where PrevMon < 0")
        'CNN.Execute("Update Rpt_Grp Set MonDep= Dep_Month Where PrevMon > 0 AND TTMon <> PrevMon")
        'CNN.Execute("Update Rpt_Grp Set MonDep= Dep_Month Where DateDiff(m, Date_Work, EndDate) >= 0 ")
        'CNN.Execute("Update Rpt_Grp Set MonDep= 0 Where TTMon=PrevMon ")
        'CNN.Execute("Update Rpt_Grp Set MonDep = 0 Where CurrMon=0 ")


        CNN.Execute("Update Rpt_Grp Set TTMon=Used_Life * 12, PrevMon=DateDiff(day, Date_Work, strDate)-1 , CurrMon=DateDiff(day, Date_Work, EndDate) ")
        CNN.Execute("Update Rpt_Grp Set TTMon=TTMon*30 ")
        CNN.Execute("Update Rpt_Grp Set CurrMon=DateDiff(day, strDate, Deposted_Date)+1 Where Deposted_Date < EndDate AND Deposted_Date is not null")
        CNN.Execute("Update Rpt_Grp Set PrevMon=0 Where PrevMon < 0")
        CNN.Execute("Update Rpt_Grp Set MonDep= Dep_Day Where PrevMon > 0 AND TTMon <> PrevMon")
        CNN.Execute("Update Rpt_Grp Set MonDep= Dep_Day Where DateDiff(day, Date_Work, EndDate) >= 0 ")
        CNN.Execute("Update Rpt_Grp Set MonDep= 0 Where TTMon=PrevMon ")
        CNN.Execute("Update Rpt_Grp Set MonDep = 0 Where CurrMon=0 ")
        '--------------Depost
        CNN.Execute("Update Rpt_Grp Set PrevMon=DateDiff(day, Date_Work, Deposted_Date) Where Deposted_Date is not null AND PrevMon > DateDiff(day, Date_Work, Deposted_Date) ")
        CNN.Execute("Update Rpt_Grp Set MonDep= 0 Where Deposted_Date is not null AND DateDiff(day, strdate, Deposted_Date) <= 0 ")
        CNN.Execute("Update Rpt_Grp Set Remain = 0 Where Deposted_Date is not null AND DateDiff(day, strdate, Deposted_Date) <= 0 ")
        ' LA Only
        CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep Where Deposted_Date is not null AND DateDiff(day, strdate, Deposted_Date) <= 0 ")
        '------------
        'CNN.Execute("Update Rpt_Grp Set Rpt_Grp.PrevDep=Open_BL.Open_Amt From Rpt_Grp, Open_BL Where Rpt_Grp.AssetID=Open_BL.AssetID AND Open_BL.Open_Amt <>0 and Year(Open_BL.Date_Work)=" & myr & "")
        'If mm >= 1 Then
        '    CNN.Execute("Update Rpt_Grp Set PrevDep = PrevDep + Dep_Month * " & mm & " Where year(Date_Work) < year(strdate) ")
        '    CNN.Execute("Update Rpt_Grp Set PrevDep = Dep_Month * (" & mm & " - month(Date_Work)) Where year(Date_Work) >= year(strdate) ")
        'End If
        CNN.Execute("Update Rpt_Grp Set PrevDep = Dep_Day*PrevMon where PrevMon > 0 ")
        'CNN.Execute("Update Rpt_Grp Set PrevDep = Amt_KIP where PrevDep > Amt_KIP ")
        CNN.Execute("Update Rpt_Grp Set TTDep = Dep_Day * CurrMon ")

        CNN.Execute("Update Rpt_Grp Set MonDep = TTDep - PrevDep ")
        CNN.Execute("Update Rpt_Grp Set MonDep = 0 where MonDep < 0 ")
        CNN.Execute("Update Rpt_Grp Set PrevDep = 0 Where DateDiff(day, Date_Work, strdate) <=1 ")
        'CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep ")
        CNN.Execute("Update Rpt_Grp Set Remain = Amt_KIP - TTDep ")
        CNN.Execute("Update Rpt_Grp Set Remain = 0 where Remain < 1 ")
        CNN.Execute("Update Rpt_Grp Set Remain = 0, Amt_KIP=0, MonDep=0 Where Deposted_Date <= strdate")
        CNN.Execute("Update Rpt_Grp Set Remain = 0, Amt_KIP=0 Where Deposted_Date <= EndDate")
     
 
        '' ''--------------Depost
        ' ''CNN.Execute("Update Rpt_Grp Set PrevMon=DateDiff(m, Date_Work, Deposted_Date) Where Deposted_Date is not null AND PrevMon > DateDiff(m, Date_Work, Deposted_Date) ")
        ' ''CNN.Execute("Update Rpt_Grp Set MonDep= 0 Where Deposted_Date is not null AND DateDiff(m, strdate, Deposted_Date) <= 0 ")
        ' ''CNN.Execute("Update Rpt_Grp Set Remain = 0 Where Deposted_Date is not null AND DateDiff(m, strdate, Deposted_Date) <= 0 ")
        '' '' LA Only
        ' ''CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep Where Deposted_Date is not null AND DateDiff(m, strdate, Deposted_Date) <= 0 ")
        '' ''------------
        '' ''CNN.Execute("Update Rpt_Grp Set Rpt_Grp.PrevDep=Open_BL.Open_Amt From Rpt_Grp, Open_BL Where Rpt_Grp.AssetID=Open_BL.AssetID AND Open_BL.Open_Amt <>0 and Year(Open_BL.Date_Work)=" & myr & "")
        '' ''If mm >= 1 Then
        '' ''    CNN.Execute("Update Rpt_Grp Set PrevDep = PrevDep + Dep_Month * " & mm & " Where year(Date_Work) < year(strdate) ")
        '' ''    CNN.Execute("Update Rpt_Grp Set PrevDep = Dep_Month * (" & mm & " - month(Date_Work)) Where year(Date_Work) >= year(strdate) ")
        '' ''End If
        ' ''CNN.Execute("Update Rpt_Grp Set PrevDep = Amt_KIP where PrevDep > Amt_KIP ")
        ' ''CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep ")
        ' ''CNN.Execute("Update Rpt_Grp Set MonDep = Amt_KIP - PrevDep where TTDep > Amt_KIP ")
        ' ''CNN.Execute("Update Rpt_Grp Set MonDep = 0 where MonDep < 0 ")
        ' ''CNN.Execute("Update Rpt_Grp Set PrevDep = 0 Where DateDiff(m, Date_Work, strdate) <=1 ")
        ' ''CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep ")
        ' ''CNN.Execute("Update Rpt_Grp Set Remain = Amt_KIP - TTDep ")
        ' ''CNN.Execute("Update Rpt_Grp Set Remain = 0 where Remain < 1 ")
        ' ''CNN.Execute("Update Rpt_Grp Set Remain = 0, Amt_KIP=0, MonDep=0 Where Deposted_Date <= strdate")
        ' ''CNN.Execute("Update Rpt_Grp Set Remain = 0, Amt_KIP=0 Where Deposted_Date <= EndDate")
    End Sub
    Private Sub CalcMon1()
        Dim cRS As New ADODB.Recordset
        Dim Str As String = ""
        Dim ss As String
        Dim strDate, EndDate As Date
        Dim mym, myy, myr As Integer
        mym = DTMon.Value.Month + 1
        myy = DTMon.Value.Year
        Dim yy As Integer = 1
        myr = DTMon.Value.Year
        Dim mm As Integer = 0
        Dim mTm As Integer = 0
        'mm = Month(DTMon.Value)
        mm = Month(DTMon.Value) - 1
        strDate = CDate("01/" & Trim(DTMon.Value.Month.ToString) & "/" & Trim(DTMon.Value.Year.ToString))
        If DTMon.Value.Month < 12 Then
            EndDate = DateAdd("d", -1, CDate("01/" & DTMon.Value.Month + 1 & "/" & DTMon.Value.Year))
        Else
            EndDate = CDate("31/12/" & DTMon.Value.Year)
        End If
        'MsgBox(strDate)
        'MsgBox(EndDate)
        If txtGrp.Text <> "" Then
            Str = " AND Group_ID='" & Trim(txtGrp.Text) & "'"
        End If

        CNN.Execute("Delete From Rpt_Grp")
        ss = "Insert into Rpt_Grp(AmountRemain,AmountClear,Amt_All,AssetID, Asset_No, Asset_Nm, Group_ID, Date_Work, Used_Life, Amt, Amt_KIP, Broked,Broke_Date, Deposted, Deposted_Date, Dep_Year, Dep_Month, TTMon, PrevMon, PrevDep, CurrMon, MonDep, TTDep, Remain, strDate, EndDate) " & _
            "Select AmountRemain,AmountClear,Amt_All,AssetID, Asset_No, Asset_Nm, Group_ID, Date_Work, Used_Life, Amt, Amt_KIP,Broked, Broke_Date, Deposted, Deposted_Date, Dep_Year, Dep_Month, 0, 0, 0, 0, 0, 0, 0, '" & Format(strDate, "yyyy-MM-dd") & "', '" & Format(EndDate, "yyyy-MM-dd") & "' From Assets Where Date_Work <= '" & Format(EndDate, "yyyy-MM-dd") & "' "
        ss = ss & Str & " Order by Asset_No "
        CNN.Execute(ss)
        CNN.Execute("Update Rpt_Grp Set TTMon=Used_Life * 12, PrevMon=DateDiff(m, Date_Work, strDate)-1, CurrMon=DateDiff(m, Date_Work, EndDate) ")
        CNN.Execute("Update Rpt_Grp Set CurrMon=DateDiff(m, strDate, Deposted_Date)+1 Where Deposted_Date < EndDate AND Deposted_Date is not null")
        CNN.Execute("Update Rpt_Grp Set PrevMon=0 Where PrevMon < 0")
        CNN.Execute("Update Rpt_Grp Set MonDep= Dep_Month Where PrevMon > 0 AND TTMon <> PrevMon")
        CNN.Execute("Update Rpt_Grp Set MonDep= Dep_Month Where DateDiff(m, Date_Work, EndDate) >= 0 ")
        CNN.Execute("Update Rpt_Grp Set MonDep= 0 Where TTMon=PrevMon ")
        CNN.Execute("Update Rpt_Grp Set MonDep = 0 Where CurrMon=0 ")
        '--------------Depost
        CNN.Execute("Update Rpt_Grp Set PrevMon=DateDiff(m, Date_Work, Deposted_Date) Where Deposted_Date is not null AND PrevMon > DateDiff(m, Date_Work, Deposted_Date) ")
        CNN.Execute("Update Rpt_Grp Set MonDep= 0 Where Deposted_Date is not null AND DateDiff(m, strdate, Deposted_Date) <= 0 ")
        CNN.Execute("Update Rpt_Grp Set Remain = 0 Where Deposted_Date is not null AND DateDiff(m, strdate, Deposted_Date) <= 0 ")
        ' LA Only
        CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep Where Deposted_Date is not null AND DateDiff(m, strdate, Deposted_Date) <= 0 ")
        '------------
        CNN.Execute("Update Rpt_Grp Set Rpt_Grp.PrevDep=Open_BL.Open_Amt From Rpt_Grp, Open_BL Where Rpt_Grp.AssetID=Open_BL.AssetID AND Open_BL.Open_Amt <>0 and Year(Open_BL.Date_Work)=" & myr & "")
        If mm >= 1 Then
            CNN.Execute("Update Rpt_Grp Set PrevDep = PrevDep + Dep_Month * " & mm & " Where year(Date_Work) < year(strdate) ")
            CNN.Execute("Update Rpt_Grp Set PrevDep = Dep_Month * (" & mm & " - month(Date_Work)) Where year(Date_Work) >= year(strdate) ")
        End If
        'CNN.Execute("Update Rpt_Grp Set PrevDep=")
        CNN.Execute("Update Rpt_Grp Set PrevDep = Amt_KIP where PrevDep > Amt_KIP ")
        CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep ")
        CNN.Execute("Update Rpt_Grp Set MonDep = Amt_KIP - PrevDep where TTDep > Amt_KIP ")
        CNN.Execute("Update Rpt_Grp Set MonDep = 0 where MonDep < 0 ")
        CNN.Execute("Update Rpt_Grp Set Remain = Amt_KIP - TTDep ")
        CNN.Execute("Update Rpt_Grp Set Remain = Amt_KIP where MonDep = 0 ")
        CNN.Execute("Update Rpt_Grp Set Remain = 0 where Remain < 1 ")
        CNN.Execute("Update Rpt_Grp Set Remain = 0, Amt_KIP=0, MonDep=0 Where Deposted_Date <= strdate")
        CNN.Execute("Update Rpt_Grp Set Remain = 0, Amt_KIP=0 Where Deposted_Date <= EndDate")
        CNN.Execute("Update Rpt_Grp Set PrevDep = 0 Where DateDiff(m, Date_Work, strdate) <=1 ")
        CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep ")

    End Sub

    Private Sub CalcTerm()
        Dim cRS As New ADODB.Recordset
        Dim Str As String = ""
        Dim ss As String
        Dim strDate, EndDate As Date
        Dim mym, myy, myr As Integer
        mym = DTMon.Value.Month + 1
        myy = DTMon.Value.Year
        Dim yy As Integer = 1
        myr = dtTerm.Value.Year
        Dim mm As Integer = 0
        Dim mTm As Integer = 0
        If cmbTerm.SelectedIndex = 0 Then
            strDate = CDate("01/01/" & dtTerm.Value.Year.ToString)
            EndDate = CDate("31/03/" & Trim(dtTerm.Value.Year.ToString))
            mTm = 0
        ElseIf cmbTerm.SelectedIndex = 1 Then
            strDate = CDate("01/04/" & dtTerm.Value.Year.ToString)
            EndDate = CDate("30/06/" & Trim(dtTerm.Value.Year.ToString))
            mTm = 3
        ElseIf cmbTerm.SelectedIndex = 2 Then
            strDate = CDate("01/07/" & dtTerm.Value.Year.ToString)
            EndDate = CDate("30/09/" & Trim(dtTerm.Value.Year.ToString))
            mTm = 6
        Else
            strDate = CDate("01/10/" & dtTerm.Value.Year.ToString)
            EndDate = CDate("31/12/" & Trim(dtTerm.Value.Year.ToString))
            mTm = 9
        End If
        If txtGrp.Text <> "" Then
            Str = " AND Group_ID='" & Trim(txtGrp.Text) & "'"
        End If
        'CNN.Execute("Delete From Rpt_Grp")
        'ss = "Insert into Rpt_Grp(AssetID, Asset_No, Asset_Nm, Group_ID, Date_Work, Used_Life, Amt_KIP, Broke_Date, Deposted_Date, Dep_Year, Dep_Month, TTMon, PrevMon, PrevDep, CurrMon, MonDep, TTDep, Remain, strDate, EndDate) " & _
        '    "Select AssetID, Asset_No, Asset_Nm, Group_ID, Date_Work, Used_Life, Amt_KIP, Broke_Date, Deposted_Date, Dep_Year, Dep_Month, 0, 0, 0, 0, 0, 0, 0, '" & Format(strDate, "yyyy-MM-dd") & "', '" & Format(EndDate, "yyyy-MM-dd") & "' From Assets Where Date_Work <= '" & Format(EndDate, "yyyy-MM-dd") & "' "
        'ss = ss & Str & " Order by Asset_No "
        'CNN.Execute(ss)
        CNN.Execute("Delete From Rpt_Grp")
        ss = "Insert into Rpt_Grp(AmountRemain,AmountClear,Amt_All,AssetID, Asset_No, Asset_Nm, Group_ID, Date_Work, Used_Life, Amt, Amt_KIP, Broke_Date,  Deposted, Deposted_Date, Dep_Year, Dep_Month, TTMon, PrevMon, PrevDep, CurrMon, MonDep, TTDep, Remain, strDate, EndDate) " & _
            "Select AmountRemain,AmountClear,Amt_All,AssetID, Asset_No, Asset_Nm, Group_ID, Date_Work, Used_Life, Amt, Amt_KIP, Broke_Date,  Deposted, Deposted_Date, Dep_Year, Dep_Month, 0, 0, 0, 0, 0, 0, 0, '" & Format(strDate, "yyyy-MM-dd") & "', '" & Format(EndDate, "yyyy-MM-dd") & "' From Assets Where Date_Work <= '" & Format(EndDate, "yyyy-MM-dd") & "' "
        ss = ss & Str & " Order by Asset_No "
        CNN.Execute(ss)
        CNN.Execute("Update Rpt_Grp Set TTMon=Used_Life * 12, PrevMon=DateDiff(m, Date_Work, strDate)-1, CurrMon=DateDiff(m, Date_Work, EndDate) ")
        CNN.Execute("Update Rpt_Grp Set CurrMon=DateDiff(m, strDate, Deposted_Date)+1 Where Deposted_Date < EndDate AND Deposted_Date is not null")
        CNN.Execute("Update Rpt_Grp Set PrevMon=0 Where PrevMon < 0")
        CNN.Execute("Update Rpt_Grp Set MonDep= Dep_Month Where PrevMon > 0 AND TTMon <> PrevMon")
        CNN.Execute("Update Rpt_Grp Set MonDep= Dep_Month Where DateDiff(m, Date_Work, EndDate) >= 0 ")
        CNN.Execute("Update Rpt_Grp Set MonDep= 0 Where TTMon=PrevMon ")
        CNN.Execute("Update Rpt_Grp Set MonDep = 0 Where CurrMon=0 ")
        '--------------Depost
        CNN.Execute("Update Rpt_Grp Set PrevMon=DateDiff(m, Date_Work, Deposted_Date) Where Deposted_Date is not null AND PrevMon > DateDiff(m, Date_Work, Deposted_Date) ")
        CNN.Execute("Update Rpt_Grp Set MonDep= 0 Where Deposted_Date is not null AND DateDiff(m, strdate, Deposted_Date) <= 0 ")
        CNN.Execute("Update Rpt_Grp Set Remain = 0 Where Deposted_Date is not null AND DateDiff(m, strdate, Deposted_Date) <= 0 ")
        CNN.Execute("Update Rpt_Grp Set PrevDep= PrevMon * Dep_Month")
        CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep Where Deposted_Date is not null AND DateDiff(m, strdate, Deposted_Date) <= 0 ")
        '------------ LA Only
        CNN.Execute("Update Rpt_Grp Set Rpt_Grp.PrevDep=Open_BL.Open_Amt From Rpt_Grp, Open_BL Where Rpt_Grp.AssetID=Open_BL.AssetID AND Open_BL.Open_Amt <>0 and Year(Open_BL.Date_Work)=" & myr & "")
        '----------------
        'CNN.Execute("Update Rpt_Grp Set PrevDep = PrevDep + MonDep* " & mTm & " Where CurrMon < TTMon")
        CNN.Execute("Update Rpt_Grp Set PrevDep = PrevDep + MonDep* " & mTm & " Where year(Date_Work) < year(strdate) and PrevDep < Amt_KIP")

        CNN.Execute("Update Rpt_Grp Set MonDep = Dep_Month * 3 Where CurrMon < TTMon or (CurrMon >= TTMon AND TTDep<Amt_KIP) ")
        CNN.Execute("Update Rpt_Grp Set MonDep = Dep_Month * DateDiff(m, Date_Work, EndDate) Where Date_Work > strdate and Date_Work < EndDate")
        CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep ")
        CNN.Execute("Update Rpt_Grp Set MonDep = Amt_KIP-PrevDep where TTDep > Amt_KIP ")
        CNN.Execute("Update Rpt_Grp Set MonDep = 0 where MonDep < 0 ")

        CNN.Execute("Update Rpt_Grp Set Remain = 0, Amt_KIP=0, MonDep=0 Where Deposted_Date <= strdate")
        CNN.Execute("Update Rpt_Grp Set Remain = 0, Amt_KIP=0 Where Deposted_Date <= EndDate")
        CNN.Execute("Update Rpt_Grp Set PrevDep = 0 Where DateDiff(m, Date_Work, strdate) <=1 ")
        CNN.Execute("Update Rpt_Grp Set MonDep = 0 where CurrMon=0 ")
        CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep ")
        CNN.Execute("Update Rpt_Grp Set Remain = Amt_KIP-TTDep ")
        CNN.Execute("Update Rpt_Grp Set Remain = 0 where Remain < 1 ")
    End Sub
    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        If optMon.Checked = True Then
            CalcMon()
        ElseIf optYear.Checked = True Then
            Call Calc()
        Else
            Call CalcTerm()
        End If
        Call LdFG()
    End Sub

    Private Sub LdFG()
        CNN.Execute("Update Rpt_Grp set Rpt_Grp.Company=Assets.Company ,Rpt_Grp.Section=Assets.Section ,Rpt_Grp.DepartmentID=Assets.DepartmentID  from Rpt_Grp,Assets where Assets.AssetID=Rpt_Grp.AssetID ")
        Dim BB As String
        If CmbShow.SelectedIndex = 0 Then
            BB = ""
        ElseIf CmbShow.SelectedIndex = 1 Then
            BB = " AND  Deposted='0' "
        Else
            BB = " AND  Deposted='1'"
        End If


        If cmbSec.SelectedIndex = 0 Then
            Sec = ""
        Else
            Sec = " AND Section='" & txtSec.Text & "' "
        End If
        If cmbDeprt.SelectedIndex = 0 Then
            SDep = ""
        Else
            SDep = " AND DepartmentID='" & txtDep.Text & "' "
        End If
        FG.Rows = 1
        Dim cRS As New ADODB.Recordset
        Call LoadSqlData("SELECT * FROM Rpt_Grp where 1=1 " & BB & " " & Sec & " " & SDep & " Order by Assetid ", cRS)
        If cRS.RecordCount <> 0 Then
            FG.Redraw = False

            While Not cRS.EOF
                FG.AddItem(cRS.AbsolutePosition & Chr(9) & Trim(cRS.Fields("AssetID").Value.ToString) & Chr(9) & Trim(cRS.Fields("Asset_No").Value.ToString) & Chr(9) & Trim(cRS.Fields("Asset_Nm").Value.ToString) & Chr(9) & Trim(cRS.Fields("Group_ID").Value.ToString) & Chr(9) & Format(cRS.Fields("Date_Work").Value, "dd/MM/yyyy") & Chr(9) & cRS.Fields("Used_life").Value.ToString & Chr(9) & Format(cRS.Fields("Amt_KIP").Value, "#,##0.00") & Chr(9) & cRS.Fields("Deposted_Date").Value & _
                           Chr(9) & Format(cRS.Fields("Dep_Year").Value, "#,##0.00") & Chr(9) & Format(cRS.Fields("Dep_Month").Value, "#,##0.00") & Chr(9) & cRS.Fields("TTMon").Value.ToString & Chr(9) & cRS.Fields("PrevMon").Value.ToString & Chr(9) & Format(cRS.Fields("PrevDep").Value, "#,##0.00") & Chr(9) & Format(cRS.Fields("MonDep").Value, "#,##0.00") & Chr(9) & Format(cRS.Fields("TTDep").Value, "#,##0.00") & Chr(9) & Format(cRS.Fields("Remain").Value, "#,##0.00"))
                cRS.MoveNext()
            End While
        Else
            FG.Rows = 1
            FG.Rows = 2
        End If
        FG.Redraw = True
    End Sub

    Private Sub optCode_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        chkSum.Checked = False
        chkSum.Enabled = False
        cmbGrp.Enabled = False
        txtGrp.Text = ""
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call Office()
        OfficeNEW()
        LoadHeader_Asset()
        Dim Rs, Rs1 As New ADODB.Recordset
        Dim mPer, mstDep As String
        Dim myStr As String = ""
        If chkSum.Checked = True Then
            If txtGrp.Text <> "" Then
                myStr = myStr & " AND A.Group_ID='" & Trim(txtGrp.Text) & "'"
            End If

        Else
            If CmbShow.SelectedIndex = 0 Then
                myStr = myStr
            ElseIf CmbShow.SelectedIndex = 1 Then
                myStr = myStr & " AND A.Deposted='0' "
            Else
                myStr = myStr & " AND A.Deposted='1'"
            End If
            If txtGrp.Text <> "" Then
                myStr = myStr & " AND A.Group_ID='" & Trim(txtGrp.Text) & "'"
            End If
            If cmbSec.SelectedIndex > 0 Then
                myStr = myStr & " AND A.Sect_ID='" & txtSec.Text & "'"
            End If
            If cmbDeprt.SelectedIndex > 0 Then
                myStr = myStr & " AND A.DepartmentID='" & txtDep.Text & "'"
            End If
            If txtCode.Text <> "" Then
                myStr = myStr & " AND (Left(A.Asset_No, " & Len(txtCode.Text) & ")='" & Trim(txtCode.Text) & "' )"
            End If
        End If
        Dim FrmPreview As New FmPreview : FrmClosing()
        If optMon.Checked = True Then
            If DTMon.Value.Month < 12 Then
                myStr = myStr & " AND A.Date_Work <= '" & Format(DateAdd("d", -1, CDate("01/" & DTMon.Value.Month + 1 & "/" & DTMon.Value.Year)), "yyyy-MM-dd") & "' "
            Else
                myStr = myStr & " AND A.Date_Work <= '" & Format(CDate("31/12/" & DTMon.Value.Year), "yyyy-MM-dd") & "' "
            End If
            If Lang = True Then
                mPer = "Monthly " & Format(DTMon.Value, "MM/yyyy")
            Else
                mPer = "ປະຈຳເດືອນ " & Format(DTMon.Value, "MM/yyyy")
            End If
            mstDep = "ຫຼຸ້ຍຫ້ຽນ ໃນເດືອນ"
        ElseIf optYear.Checked = True Then
            If Lang = True Then
                mPer = "Yearly " & Format(DTYear.Value, "yyyy")
            Else
                mPer = "ປະຈຳປີ " & Format(DTYear.Value, "yyyy")
            End If
            myStr = myStr & " AND YEAR(A.Date_Work)<=" & Year(DTYear.Value) & " "
            mstDep = "ຫຼຸ້ຍຫ້ຽນ  ໃນປີ"
        Else
            If Lang = True Then
                mPer = "Termly " & Trim(cmbTerm.Text) & " / " & Format(dtTerm.Value, "yyyy")
            Else
                mPer = "ປະຈຳງວດ " & Trim(cmbTerm.Text) & " / " & Format(dtTerm.Value, "yyyy")
            End If
            myStr = myStr & " AND YEAR(A.Date_Work)<=" & Year(DTYear.Value) & " "
            mstDep = "ຫຼຸ້ຍຫ້ຽນ ໃນງວດ"
        End If
        Call Calc()
        'If optMon.Checked = True Then
        '    CalcMon()
        'ElseIf optYear.Checked = True Then
        '    Call Calc()
        'Else
        '    Call CalcTerm()
        'End If
        Call Office()
        CNN.Execute("Update Rpt_Grp set Rpt_Grp.Company=Assets.Company ,Rpt_Grp.Section=Assets.Section ,Rpt_Grp.DepartmentID=Assets.DepartmentID  from Rpt_Grp,Assets where Assets.AssetID=Rpt_Grp.AssetID ")
        If cmbSec.SelectedIndex = 0 Then
            Sec = ""
        Else
            Sec = " AND B.Section='" & txtSec.Text & "' "
        End If
        If cmbDeprt.SelectedIndex = 0 Then
            SDep = ""
        Else
            SDep = " AND B.DepartmentID='" & txtDep.Text & "' "
        End If


        If chkSum.Checked = False Then
            If CmbShow.SelectedIndex = 2 Then
                Dim sss As String = "Select A.Qty, G.Group_Nm,G.Group_Nme, D.DepartmentNm,D.DepartmentNme, A.DepartmentID, A.Model, A.Engin_No, A.Frame_No, A.Serial, A.Used_Life, A.Date_Work, A.AssetID, A.Asset_No, A.Asset_Nm,A.Asset_NmE, A.Group_ID, A.Date_Work, A.Used_Life, A.Amount, A.Curr, A.amt, B.Amt_KIP, A.Broke_Date, A.Deposted_Date, A.Dep_Year, A.Dep_Month, B.TTMon, B.PrevMon, B.PrevDep, B.MonDep , B.TTDep, PST.Descriptions AS Remain, A.regist_no,B.AmountRemain,B.AmountClear, B.Amt_All " & _
                              " from Assets A LEFT OUTER JOIN  Rpt_Grp B ON A.AssetID=B.AssetID LEFT OUTER JOIN Department D ON D.DepartmentID=A.DepartmentID LEFT OUTER JOIN Groups G ON G.Group_ID=A.Group_ID LEFT OUTER JOIN Brokens PST ON A.AssetID=PST.AssetID Where 1=1 " & myStr & " " & Sec & " " & SDep & " Order by A.AssetID "
                Call LoadSqlData(sss, Rs)
                Dim myText5 As CrystalDecisions.CrystalReports.Engine.TextObject
                If Lang = True Then
                    rpt = New CryRpt_GrpLongPostEng
                Else
                    rpt = New CryRpt_GrpLongPost
                    Dim myText4 As CrystalDecisions.CrystalReports.Engine.TextObject
                    myText5 = CType(rpt.ReportDefinition.ReportObjects.Item("Text1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText5.Text = "ສະຫຸບລາຍງານຊັບສົມບັດຄົງທີ ລໍຖ້າສະສາງ"
                    If cmbSec.SelectedIndex = 0 Then
                        myText4 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                        myText4.Text = OffName
                    Else
                        myText4 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                        myText4.Text = cmbSec.Text
                    End If
                    If cmbDeprt.SelectedIndex = 0 Then
                        myText4 = CType(rpt.ReportDefinition.ReportObjects.Item("Text3"), CrystalDecisions.CrystalReports.Engine.TextObject)
                        myText4.Text = cmbDeprt.Text
                    Else
                        myText4 = CType(rpt.ReportDefinition.ReportObjects.Item("Text3"), CrystalDecisions.CrystalReports.Engine.TextObject)
                        myText4.Text = cmbDeprt.Text
                    End If
                End If

            Else
                Dim sss As String = "Select A.Qty, G.Group_Nm,G.Group_Nme, D.DepartmentNm,D.DepartmentNme, A.DepartmentID, A.Model, A.Engin_No, A.Frame_No, A.Serial, A.Used_Life, A.Date_Work, A.AssetID, A.Asset_No, A.Asset_Nm, A.Asset_NmE, A.Group_ID, A.Date_Work, A.Used_Life, A.Amount, A.Curr, A.amt, B.Amt_KIP, A.Broke_Date, A.Deposted_Date, A.Dep_Year, A.Dep_Month, B.TTMon, B.PrevMon, B.PrevDep, B.MonDep , B.TTDep, B.Remain, A.regist_no,B.AmountRemain,B.AmountClear,B.Amt_All,dbo.Sections.SecNmL, dbo.Sections.SecNmE, A.Sect_ID, D.DepartmentNmE, G.Group_NmE, A.Asset_NmE " & _
                              " FROM dbo.Assets AS A INNER JOIN  dbo.Sections ON A.Sect_ID = dbo.Sections.SecID LEFT OUTER JOIN dbo.Rpt_Grp AS B ON A.AssetID = B.AssetID LEFT OUTER JOIN " & _
                              " dbo.Department AS D ON D.DepartmentID = A.DepartmentID LEFT OUTER JOIN dbo.Groups AS G ON G.Group_ID = A.Group_ID Where 1=1 " & myStr & "  " & Sec & " " & SDep & " Order by A.AssetID "
                Call LoadSqlData(sss, Rs)
                If chkBranch.Checked = True Then
                    If Lang = True Then
                        rpt = New CryRpt_GrpLongNoBrEng
                    Else
                        rpt = New CryRpt_GrpLongNoBr
                    End If

                    Call footder()
                Else
                    Dim myText4 As CrystalDecisions.CrystalReports.Engine.TextObject

                    If Lang = False Then
                        rpt = New CryRpt_GrpLong

                        'myText4 = CType(rpt.ReportDefinition.ReportObjects.Item("MM"), CrystalDecisions.CrystalReports.Engine.TextObject)
                        'myText4.Text = MM
                        If cmbSec.SelectedIndex = 0 Then
                            myText4 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                            myText4.Text = OffName
                        Else
                            myText4 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                            myText4.Text = cmbSec.Text
                        End If
                        If cmbDeprt.SelectedIndex = 0 Then
                            myText4 = CType(rpt.ReportDefinition.ReportObjects.Item("Text1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                            myText4.Text = cmbDeprt.Text
                        Else
                            myText4 = CType(rpt.ReportDefinition.ReportObjects.Item("Text1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                            myText4.Text = cmbDeprt.Text
                        End If
                    Else
                        rpt = New CryRpt_GrpLongEng
                    End If
                End If
            End If
            Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("txtPeriod"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = mPer
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Show"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = CmbShow.Text
            If Lang = False Then
                myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Head"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = "ສະຫຸບລາຍງານຊັບສົມບັດຄົງທີ ທີ່ບໍ່ຫມົດ​ຄ່າຫຼຸ້ຍ​ຫ້ຽນ"
            Else
                myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Head"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = "Assets List " & cmbGrp.Text
            End If
            'If optYear.Checked = True Then

            'End If
            Dim myText1 As CrystalDecisions.CrystalReports.Engine.TextObject
            myText1 = CType(rpt.ReportDefinition.ReportObjects.Item("txtDep"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText1.Text = mstDep
            Dim myText3 As CrystalDecisions.CrystalReports.Engine.TextObject
            'myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("txtGrp"), CrystalDecisions.CrystalReports.Engine.TextObject)
            'myText3.Text = Trim(cmbGrp.Text)
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Tel"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = "Tel: " & OffTel & " Fax: " & OffFax
            If Lang = False Then
                'myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                'myText3.Text = OffName
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal5"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign5
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal4"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign4
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal3"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign3
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal2"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign2
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign1
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Place"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = PlaecL
            Else
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = OffNameE
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal5"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign5e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal4"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign4e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal3"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign3e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal2"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign2e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign1e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Place"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = PlaecE
            End If

        Else
            If cmbSec.SelectedIndex = 0 Then
                Sec = ""
            Else
                Sec = " AND C.Section='" & txtSec.Text & "' "
            End If
            If cmbDeprt.SelectedIndex = 0 Then
                SDep = ""
            Else
                SDep = " AND C.DepartmentID='" & txtDep.Text & "' "
            End If

            Dim ss As String = "Select  A.Group_ID, G.Group_Nm, G.Group_NmE, sum(A.Amt) as Amt, sum(A.Amt_KIP) as Amt_KIP, sum(A.Dep_Year) as Dep_Year, sum(A.Dep_Month) as Dep_Month, sum(A.PrevMon) as PrevMon, sum(A.PrevDep) as PrevDep, sum(A.MonDep) as MonDep, sum(A.TTDep) as TTDep, sum(A.Remain) as Remain , " & _
            " sum(A.AmountRemain) as AmountRemain, sum(A.AmountClear) as AmountClear, sum(A.Amt_All) as Amt_All " & _
               " from Assets C LEFT OUTER JOIN Rpt_Grp A ON C.Group_ID=A.Group_ID LEFT OUTER JOIN Groups G ON A.Group_ID=G.Group_ID " & _
               " Where 1=1" & myStr & _
                 " GROUP By A.Group_ID, G.Group_Nm, G.Group_NmE " & _
                " ORDER by A.Group_ID "
            Call LoadSqlData(ss, Rs)
            If Lang = True Then
                rpt = New CryRpt_GrpSumLongEng
            Else
                rpt = New CryRpt_GrpSumLong
                Dim myText4 As CrystalDecisions.CrystalReports.Engine.TextObject
                If cmbSec.SelectedIndex = 0 Then

                    myText4 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText4.Text = OffName
                Else
                    myText4 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText4.Text = cmbSec.Text
                End If
                If cmbDeprt.SelectedIndex = 0 Then
                    myText4 = CType(rpt.ReportDefinition.ReportObjects.Item("D1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText4.Text = cmbDeprt.Text
                Else
                    myText4 = CType(rpt.ReportDefinition.ReportObjects.Item("D1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText4.Text = cmbDeprt.Text
                End If
            End If

            Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
            Dim myText1, myText3 As CrystalDecisions.CrystalReports.Engine.TextObject
            myText1 = CType(rpt.ReportDefinition.ReportObjects.Item("txtSec"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText1.Text = CmbShow.Text
            If Lang = True Then
                myText1 = CType(rpt.ReportDefinition.ReportObjects.Item("Head"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText1.Text = "Assets List " & cmbGrp.Text
            Else
                myText1 = CType(rpt.ReportDefinition.ReportObjects.Item("Head"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText1.Text = "ສະຫຸບລາຍງານຊັບສົມບັດຄົງທີ ທີ່ບໍ່ຫມົດ​ຄ່າຫຼຸ້ຍ​ຫ້ຽນ "
            End If
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("txtPeriod"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = mPer
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Tel"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = "Tel: " & OffTel & " Fax: " & OffFax
            If Lang = False Then
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = OffName
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal5"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign5
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal4"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign4
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal3"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign3
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal2"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign2
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign1
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Place"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = PlaecL
            Else
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = OffNameE
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal5"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign5e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal4"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign4e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal3"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign3e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal2"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign2e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign1e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Place"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = PlaecE
            End If
        End If
        If Rs.RecordCount = 0 Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
        With rpt
            'Dim RO As CrystalDecisions.CrystalReports.Engine.ReportObject
            'Dim SRO As CrystalDecisions.CrystalReports.Engine.SubreportObject
            'Dim SubDoc As CrystalDecisions.CrystalReports.Engine.ReportDocument
            'SqlPrint = "SELECT  * from Ap_Image  where Img_Id='" & IMageID & "'"
            'Call LoadSqlData(SqlPrint, Rs1)
            'RO = rpt.ReportDefinition.Sections.Item("Section1").ReportObjects.Item("Subreport3")
            'SRO = CType(RO, CrystalDecisions.CrystalReports.Engine.SubreportObject)
            'SubDoc = SRO.OpenSubreport(SRO.SubreportName)
            'If Rs1.RecordCount > 0 Then
            '    SubDoc.SetDataSource(Rs1)
            '    FmPreview.ReportViewer.ReportSource = SubDoc
            'End If
            rpt.SetDataSource(Rs)
            rpt.Refresh()
            FmPreview.ReportViewer.ReportSource = rpt
            FmPreview.ReportViewer.DisplayGroupTree = False
            FmPreview.WindowState = FormWindowState.Maximized
            FmPreview.Show()
            'Call CloseRs(RSC)
            'Call CloseRs(Rs)
        End With
    End Sub

    Private Sub cmbSec_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSec.SelectedIndexChanged
        Dim sRS As New ADODB.Recordset
        If Lang = True Then
            Call LoadSqlData("select * from AP_Office Where Off_NameE=N'" & Trim(cmbSec.Text) & "'", sRS)
            If sRS.RecordCount <> 0 Then
                txtSec.Text = Trim(sRS.Fields("Off_ID").Value.ToString)
                Dim dRS As New ADODB.Recordset
                cmbDeprt.Items.Clear()
                cmbDeprt.Items.Add("** All Department ***")
                Call LoadSqlData("Select * from Department Where Left(DepartmentID,2) = '" & Trim(txtSec.Text) & "' Order by DepartmentID", dRS)
                If dRS.RecordCount <> 0 Then
                    While Not dRS.EOF
                        cmbDeprt.Items.Add(dRS.Fields("DepartmentNmE").Value.ToString)
                        dRS.MoveNext()
                    End While
                End If
                If cmbDeprt.Items.Count > 0 Then cmbDeprt.SelectedIndex = 0
            Else
                txtSec.Text = ""
            End If

            If cmbSec.SelectedIndex = 0 Then
                txtCompany.Text = ""
            ElseIf cmbSec.SelectedIndex = 1 Then
                'txtCompany.Text = "01-VTE"
                txtCompany.Text = txtSec.Text
                TxtCertify.Text = txtGrp.Text & "." & Format((DTMon.Value), "MM.yy") & "/" & Trim(txtCompany.Text)
            Else
                txtCompany.Text = ""
                'TextBox1.Text = txtGrp.Text & "." & Format((DTMon.Value), "MM.yy") & "/" & Trim(txtCompany.Text)
            End If
        Else
            Call LoadSqlData("select * from AP_Office Where Off_Name=N'" & Trim(cmbSec.Text) & "'", sRS)
            If sRS.RecordCount <> 0 Then
                txtSec.Text = Trim(sRS.Fields("Off_ID").Value.ToString)
                Dim dRS As New ADODB.Recordset
                cmbDeprt.Items.Clear()
                cmbDeprt.Items.Add("ສະແດງທັງໝົດທຸກພະແນກ")
                Call LoadSqlData("Select * from Department Where Left(DepartmentID,2) = '" & Trim(txtSec.Text) & "' Order by DepartmentID", dRS)
                If dRS.RecordCount <> 0 Then
                    While Not dRS.EOF
                        cmbDeprt.Items.Add(dRS.Fields("DepartmentNm").Value.ToString)
                        dRS.MoveNext()
                    End While
                End If
                If cmbDeprt.Items.Count > 0 Then cmbDeprt.SelectedIndex = 0
            Else
                txtSec.Text = ""
            End If

            If cmbSec.SelectedIndex = 0 Then
                txtCompany.Text = ""
            ElseIf cmbSec.SelectedIndex = 1 Then
                txtCompany.Text = txtSec.Text
                TxtCertify.Text = txtGrp.Text & "." & Format((DTMon.Value), "MM.yy") & "/" & Trim(txtCompany.Text)
            Else
                txtCompany.Text = ""
                'TextBox1.Text = txtGrp.Text & "." & Format((DTMon.Value), "MM.yy") & "/" & Trim(txtCompany.Text)
            End If
        End If
    End Sub

    Private Sub cmbDeprt_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbDeprt.SelectedIndexChanged
        Dim dRS As New ADODB.Recordset
        If Lang = True Then
            Call LoadSqlData("Select * from Department Where DepartmentNmE= N'" & Trim(cmbDeprt.Text) & "' ", dRS)
            If dRS.RecordCount <> 0 Then
                txtDep.Text = Trim(dRS.Fields("DepartmentID").Value.ToString)
                txtCompany.Text = Trim(dRS.Fields("Company").Value.ToString)
            Else
                txtDep.Text = ""
            End If
        Else
            Call LoadSqlData("Select * from Department Where DepartmentNm = N'" & Trim(cmbDeprt.Text) & "' ", dRS)
            If dRS.RecordCount <> 0 Then
                txtDep.Text = Trim(dRS.Fields("DepartmentID").Value.ToString)
                txtCompany.Text = Trim(dRS.Fields("Company").Value.ToString)
            Else
                txtDep.Text = ""
            End If
        End If
        TxtCertify.Text = txtGrp.Text & "." & Format((DTMon.Value), "MM.yy") & "/" & Trim(txtCompany.Text)
    End Sub

    Private Sub optTerm_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTerm.CheckedChanged
        cmbTerm.Enabled = True
        dtTerm.Enabled = True
        DTMon.Enabled = False
        DTYear.Enabled = False
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        CNN.Execute("Update Rpt_Grp set Rpt_Grp.Company=Assets.Company ,Rpt_Grp.Section=Assets.Section ,Rpt_Grp.DepartmentID=Assets.DepartmentID  from Rpt_Grp,Assets where Assets.AssetID=Rpt_Grp.AssetID ")

        Dim SLL As New ADODB.Recordset
        Dim MMDate As Date = CDate("01/" & Trim(DTMon.Value.Month.ToString) & "/" & Trim(DTMon.Value.Year.ToString))
        If optMon.Checked = False Then
            MsgBox("ກະລຸນາເລືອກປະຈໍາເດືອນກ່ອນ") : Exit Sub
        End If
        If cmbSec.Text = "" Then MsgBox("ກະລຸນາເລືອກສາຂາກ່ອນຈະໂອນ") : Exit Sub
        If cmbGrp.SelectedIndex = 0 Then
            MsgBox("ກະລຸນາເລືອກໝວດຊັບສິນທີ່ທ່ານຕ້ອງການຈະໂອນກ່ອນ") : cmbGrp.Focus() : Exit Sub
        End If
        If optMon.Checked = True Then
            If MsgBox("ທ່ານຕ້ອງການໂອນໝວດຊັບສິນ " & Trim(cmbGrp.Text) & " ນີ້ບໍ່?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo + MsgBoxStyle.Question) = Windows.Forms.DialogResult.Yes Then
                Dim KK As String = "select * from gen_jn where company='" & Trim(txtCompany.Text) & "' and date_work='" & Format(MMDate, "yyyy-MM-dd") & "' and certify='" & Trim(TxtCertify.Text) & "' "
                Call LoadSqlData(KK, SLL)
                If SLL.RecordCount = 0 Then
                    'CalcMonCOMPANY()
                    Dim kk2 As String = "SElecT sum(MonDep) as MonDep,Group_ID,section from  Rpt_Grp where  Group_ID='" & Trim(txtGrp.Text) & "' and section='" & Trim(txtSec.Text) & "' group by Group_ID,section "
                    Call LoadSqlData(kk2, RSC)
                    If RSC.RecordCount > 0 Then
                        Dim Lng As String = "INSERT INTO gen_jn(certify,Referno,company,Com_id,don_id,office_id, Book, amount, net_amt, date_work, code_dr,code_cr,ac_code,amount_dr,amount_cr,amt_dr,amt_cr,last_update,last_user,my_lock,amt_USD_dr,amt_USD_cr)  " & _
                                           " VALUES('" & Trim(TxtCertify.Text) & "','" & Trim(TxtCertify.Text) & "','" & Trim(txtCompany.Text) & "','" & Trim(txtCompany.Text) & "','01','" & Trim(txtCompany.Text) & "','Fixd Asset'," & RSC.Fields("MonDep").Value & "," & RSC.Fields("MonDep").Value & ",'" & Format(MMDate, "yyyy-MM-dd") & "','" & txtAcc.Text & "','','" & txtAcc.Text & "'," & RSC.Fields("MonDep").Value & ",0," & RSC.Fields("MonDep").Value & ",0,'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "','0','0','0')"
                        CNN.Execute(Lng)
                        Dim Lng1 As String = "INSERT INTO gen_jn(certify,Referno,company, Com_id,don_id,office_id, Book, amount, net_amt, date_work, code_dr,code_cr,ac_code,amount_dr,amount_cr,amt_dr,amt_cr,last_update,last_user,my_lock,amt_USD_dr,amt_USD_cr)  " & _
                                           " VALUES('" & Trim(TxtCertify.Text) & "','" & Trim(TxtCertify.Text) & "','" & Trim(txtCompany.Text) & "','" & Trim(txtCompany.Text) & "','01','" & Trim(txtCompany.Text) & "','Fixd Asset'," & RSC.Fields("MonDep").Value & "," & RSC.Fields("MonDep").Value & ",'" & Format(MMDate, "yyyy-MM-dd") & "','','" & TxtLH.Text & "','" & TxtLH.Text & "',0," & RSC.Fields("MonDep").Value & ",0," & RSC.Fields("MonDep").Value & ",'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "','0','0','0')"
                        CNN.Execute(Lng1)
                    End If
                Else
                    Dim De As String = "Delete from gen_jn where company='" & Trim(txtCompany.Text) & "' and month(date_work)='" & (MMDate.Month) & "' and year(date_work)='" & (MMDate.Year) & "'"
                    CNN.Execute(De)
                    'CalcMonCOMPANY()
                    Call LoadSqlData("SElecT sum(MonDep) as MonDep,Group_ID from  Rpt_Grp where  Group_ID='" & Trim(txtGrp.Text) & "'  and section='" & Trim(txtSec.Text) & "' group by Group_ID,section ", RSC)
                    If RSC.RecordCount > 0 Then
                        Dim Lng As String = "INSERT INTO gen_jn(certify,Referno,company,Com_id,don_id,office_id, Book, amount, net_amt, date_work, code_dr,code_cr,ac_code,amount_dr,amount_cr,amt_dr,amt_cr,last_update,last_user,my_lock,amt_USD_dr,amt_USD_cr)  " & _
                                        " VALUES('" & Trim(TxtCertify.Text) & "','" & Trim(TxtCertify.Text) & "','" & Trim(txtCompany.Text) & "','" & Trim(txtCompany.Text) & "','01','" & Trim(txtCompany.Text) & "','Fixd Asset'," & RSC.Fields("MonDep").Value & "," & RSC.Fields("MonDep").Value & ",'" & Format(MMDate, "yyyy-MM-dd") & "','" & txtAcc.Text & "','','" & txtAcc.Text & "'," & RSC.Fields("MonDep").Value & ",0," & RSC.Fields("MonDep").Value & ",0,'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "','0','0','0')"
                        CNN.Execute(Lng)
                        Dim Lng1 As String = "INSERT INTO gen_jn(certify,Referno,company, Com_id,don_id,office_id, Book, amount, net_amt, date_work, code_dr,code_cr,ac_code,amount_dr,amount_cr,amt_dr,amt_cr,last_update,last_user,my_lock,amt_USD_dr,amt_USD_cr)  " & _
                                           " VALUES('" & Trim(TxtCertify.Text) & "','" & Trim(TxtCertify.Text) & "','" & Trim(txtCompany.Text) & "','" & Trim(txtCompany.Text) & "','01','" & Trim(txtCompany.Text) & "','Fixd Asset'," & RSC.Fields("MonDep").Value & "," & RSC.Fields("MonDep").Value & ",'" & Format(MMDate, "yyyy-MM-dd") & "','','" & TxtLH.Text & "','" & TxtLH.Text & "',0," & RSC.Fields("MonDep").Value & ",0," & RSC.Fields("MonDep").Value & ",'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "','0','0','0')"
                        CNN.Execute(Lng1)
                    End If
                End If
                CNN.Execute("update Gen_jn set Gen_jn.descrip=Acc_Code.Name_L, Gen_jn.ac_name=Acc_Code.Name_L, Gen_jn.curr='LAK', Gen_jn.rate=1,Gen_jn.rate_i=1, Gen_jn.ac_typee=Acc_Code.Acc_TypeE from Acc_Code,Gen_jn where Gen_jn.certify='" & Trim(TxtCertify.Text) & "' and Gen_jn.AC_Code=ACC_Code.AC_Code ")
                MsgBox("ການໂອນສຳເລັດຜົນ")
            End If

        ElseIf optTerm.Checked = True Then
            MsgBox("ບໍ່ສາມາດໂອນເປັນງວດໄດ້")
        ElseIf optYear.Checked = True Then
            MsgBox("ບໍ່ສາມາດໂອນເປັນປີໄດ້")
        End If
    End Sub
    Private Sub CalcMonCOMPANY()
        Dim cRS As New ADODB.Recordset
        Dim Str As String = ""
        Dim ss As String
        Dim strDate, EndDate As Date
        Dim mym, myy, myr As Integer
        mym = DTMon.Value.Month + 1
        myy = DTMon.Value.Year
        Dim yy As Integer = 1
        myr = DTMon.Value.Year
        Dim mm As Integer = 0
        Dim mTm As Integer = 0
        'mm = Month(DTMon.Value)
        mm = Month(DTMon.Value) - 1
        strDate = CDate("01/" & Trim(DTMon.Value.Month.ToString) & "/" & Trim(DTMon.Value.Year.ToString))
        If DTMon.Value.Month < 12 Then
            EndDate = DateAdd("d", -1, CDate("01/" & DTMon.Value.Month + 1 & "/" & DTMon.Value.Year))
        Else
            EndDate = CDate("31/12/" & DTMon.Value.Year)
        End If
        'MsgBox(strDate)
        'MsgBox(EndDate)
        If txtGrp.Text <> "" Then
            Str = " AND Group_ID='" & Trim(txtGrp.Text) & "'"
        End If
        Str = Str & " AND section='" & Trim(txtSec.Text) & "'"
        CNN.Execute("Delete From Rpt_Grp")
        ss = "Insert into Rpt_Grp(Company,AmountRemain,AmountClear,Amt_All,AssetID, Asset_No, Asset_Nm, Group_ID, Date_Work, Used_Life, Amt, Amt_KIP, Broked,Broke_Date, Deposted, Deposted_Date, Dep_Year, Dep_Month, TTMon, PrevMon, PrevDep, CurrMon, MonDep, TTDep, Remain, strDate, EndDate) " & _
            "Select Company,AmountRemain,AmountClear,Amt_All,AssetID, Asset_No, Asset_Nm, Group_ID, Date_Work, Used_Life, Amt, Amt_KIP,Broked, Broke_Date, Deposted, Deposted_Date, Dep_Year, Dep_Month, 0, 0, 0, 0, 0, 0, 0, '" & Format(strDate, "yyyy-MM-dd") & "', '" & Format(EndDate, "yyyy-MM-dd") & "' From Assets Where Date_Work <= '" & Format(EndDate, "yyyy-MM-dd") & "' "
        ss = ss & Str & " Order by Asset_No "
        CNN.Execute(ss)
        CNN.Execute("Update Rpt_Grp Set TTMon=Used_Life * 12, PrevMon=DateDiff(m, Date_Work, strDate)-1, CurrMon=DateDiff(m, Date_Work, EndDate) ")
        CNN.Execute("Update Rpt_Grp Set CurrMon=DateDiff(m, strDate, Deposted_Date)+1 Where Deposted_Date < EndDate AND Deposted_Date is not null")
        CNN.Execute("Update Rpt_Grp Set PrevMon=0 Where PrevMon < 0")
        CNN.Execute("Update Rpt_Grp Set MonDep= Dep_Month Where PrevMon > 0 AND TTMon <> PrevMon")
        CNN.Execute("Update Rpt_Grp Set MonDep= Dep_Month Where DateDiff(m, Date_Work, EndDate) >= 0 ")
        CNN.Execute("Update Rpt_Grp Set MonDep= 0 Where TTMon=PrevMon ")
        CNN.Execute("Update Rpt_Grp Set MonDep = 0 Where CurrMon=0 ")
        '--------------Depost
        CNN.Execute("Update Rpt_Grp Set PrevMon=DateDiff(m, Date_Work, Deposted_Date) Where Deposted_Date is not null AND PrevMon > DateDiff(m, Date_Work, Deposted_Date) ")
        CNN.Execute("Update Rpt_Grp Set MonDep= 0 Where Deposted_Date is not null AND DateDiff(m, strdate, Deposted_Date) <= 0 ")
        CNN.Execute("Update Rpt_Grp Set Remain = 0 Where Deposted_Date is not null AND DateDiff(m, strdate, Deposted_Date) <= 0 ")
        ' LA Only
        CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep Where Deposted_Date is not null AND DateDiff(m, strdate, Deposted_Date) <= 0 ")
        '------------
        CNN.Execute("Update Rpt_Grp Set Rpt_Grp.PrevDep=Open_BL.Open_Amt From Rpt_Grp, Open_BL Where Rpt_Grp.AssetID=Open_BL.AssetID AND Open_BL.Open_Amt <>0 and Year(Open_BL.Date_Work)=" & myr & "")
        If mm >= 1 Then
            CNN.Execute("Update Rpt_Grp Set PrevDep = PrevDep + Dep_Month * " & mm & " Where year(Date_Work) < year(strdate) ")
            CNN.Execute("Update Rpt_Grp Set PrevDep = Dep_Month * (" & mm & " - month(Date_Work)) Where year(Date_Work) >= year(strdate) ")
        End If
        'CNN.Execute("Update Rpt_Grp Set PrevDep=")
        CNN.Execute("Update Rpt_Grp Set PrevDep = Amt_KIP where PrevDep > Amt_KIP ")
        CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep ")
        CNN.Execute("Update Rpt_Grp Set MonDep = Amt_KIP - PrevDep where TTDep > Amt_KIP ")
        CNN.Execute("Update Rpt_Grp Set MonDep = 0 where MonDep < 0 ")
        CNN.Execute("Update Rpt_Grp Set Remain = Amt_KIP - TTDep ")
        CNN.Execute("Update Rpt_Grp Set Remain = 0 where Remain < 1 ")
        CNN.Execute("Update Rpt_Grp Set Remain = 0, Amt_KIP=0, MonDep=0 Where Deposted_Date <= strdate")
        CNN.Execute("Update Rpt_Grp Set Remain = 0, Amt_KIP=0 Where Deposted_Date <= EndDate")
        CNN.Execute("Update Rpt_Grp Set PrevDep = 0 Where DateDiff(m, Date_Work, strdate) <=1 ")
        CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep ")
    End Sub
    Private Sub CmbCompany_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCompany.SelectedIndexChanged
        Dim gRS As New ADODB.Recordset
        Call LoadSqlData("select * from office Where off_name=N'" & Trim(CmbCompany.Text) & "'", gRS)
        If gRS.RecordCount <> 0 Then
            txtCompany.Text = Trim(gRS.Fields("off_id").Value.ToString)
        Else
            txtCompany.Text = ""
        End If
    End Sub

    Private Sub DTMon_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTMon.ValueChanged
        TxtCertify.Text = txtGrp.Text & "." & Format((DTMon.Value), "MM.yy") & "/" & Trim(txtCompany.Text)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        LoadHeader_Asset()
        Dim Rs As New ADODB.Recordset
        Dim mPer, mstDep As String
        Dim myStr As String = ""
        If chkSum.Checked = True Then
            If txtGrp.Text <> "" Then
                myStr = myStr & " AND A.Group_ID='" & Trim(txtGrp.Text) & "'"
            End If

        Else
            If CmbShow.SelectedIndex = 0 Then
                myStr = myStr
            ElseIf CmbShow.SelectedIndex = 1 Then
                myStr = myStr & " AND A.Deposted='0' "
            Else
                myStr = myStr & " AND A.Deposted='1'"
            End If
            If txtGrp.Text <> "" Then
                myStr = myStr & " AND A.Group_ID='" & Trim(txtGrp.Text) & "'"
            End If
            If cmbSec.SelectedIndex > 0 Then
                myStr = myStr & " AND A.Sect_ID='" & txtSec.Text & "'"
            End If
            If cmbDeprt.SelectedIndex > 0 Then
                myStr = myStr & " AND A.DepartmentID='" & txtDep.Text & "'"
            End If
            If txtCode.Text <> "" Then
                myStr = myStr & " AND (Left(A.Asset_No, " & Len(txtCode.Text) & ")='" & Trim(txtCode.Text) & "' )"
            End If
        End If
        Dim FrmPreview As New FmPreview : FrmClosing()
        If optMon.Checked = True Then
            If DTMon.Value.Month < 12 Then
                myStr = myStr & " AND A.Date_Work <= '" & Format(DateAdd("d", -1, CDate("01/" & DTMon.Value.Month + 1 & "/" & DTMon.Value.Year)), "yyyy-MM-dd") & "' "
            Else
                myStr = myStr & " AND A.Date_Work <= '" & Format(CDate("31/12/" & DTMon.Value.Year), "yyyy-MM-dd") & "' "
            End If
            If Lang = True Then
                mPer = "Monthly " & Format(DTMon.Value, "MM/yyyy")
            Else
                mPer = "ປະຈຳເດືອນ " & Format(DTMon.Value, "MM/yyyy")
            End If
            mstDep = "ຫຼຸ້ຍຫ້ຽນ ໃນເດືອນ"
        ElseIf optYear.Checked = True Then
            If Lang = True Then
                mPer = "Yearly " & Format(DTYear.Value, "yyyy")
            Else
                mPer = "ປະຈຳປີ " & Format(DTYear.Value, "yyyy")
            End If
            myStr = myStr & " AND YEAR(A.Date_Work)<=" & Year(DTYear.Value) & " "
            mstDep = "ຫຼຸ້ຍຫ້ຽນ ໃນປີ"
        Else
            If Lang = True Then
                mPer = "Termly " & Trim(cmbTerm.Text) & " / " & Format(dtTerm.Value, "yyyy")
            Else
                mPer = "ປະຈຳງວດ " & Trim(cmbTerm.Text) & " / " & Format(dtTerm.Value, "yyyy")
            End If
            myStr = myStr & " AND YEAR(A.Date_Work)<=" & Year(DTYear.Value) & " "
            mstDep = "ຫຼຸ້ຍຫ້ຽນ ໃນງວດ"
        End If
        If optMon.Checked = True Then
            CalcMon()
        ElseIf optYear.Checked = True Then
            Call Calc()
        Else
            Call CalcTerm()
        End If
        Call Office()
        If chkSum.Checked = False Then
            If CmbShow.SelectedIndex = 2 Then
                Dim sss As String = "Select A.Qty, G.Group_Nm,G.Group_Nme, D.DepartmentNm,D.DepartmentNme, A.DepartmentID, A.Model, A.Engin_No, A.Frame_No, A.Serial, A.Used_Life, A.Date_Work, A.AssetID, A.Asset_No, A.Asset_Nm, A.Group_ID, A.Date_Work, A.Used_Life, A.Amount, A.Curr, A.amt, B.Amt_KIP, A.Broke_Date, A.Deposted_Date, A.Dep_Year, A.Dep_Month, B.TTMon, B.PrevMon, B.PrevDep, B.MonDep , B.TTDep, PST.Descriptions AS Remain, A.regist_no,B.AmountRemain,B.AmountClear, B.Amt_All " & _
                              " from Assets A LEFT OUTER JOIN  Rpt_Grp B ON A.AssetID=B.AssetID LEFT OUTER JOIN Department D ON D.DepartmentID=A.DepartmentID LEFT OUTER JOIN Groups G ON G.Group_ID=A.Group_ID LEFT OUTER JOIN Brokens PST ON A.AssetID=PST.AssetID Where 1=1 " & myStr & " Order by A.AssetID "
                Call LoadSqlData(sss, Rs)

                rpt = New CryRpt_GrpLongPost
            Else
                Dim sss As String = "Select A.Qty, G.Group_Nm,G.Group_Nme, D.DepartmentNm,D.DepartmentNme, A.DepartmentID, A.Model, A.Engin_No, A.Frame_No, A.Serial, A.Used_Life, A.Date_Work, A.AssetID, A.Asset_No, A.Asset_Nm, A.Asset_NmE, A.Group_ID, A.Date_Work, A.Used_Life, A.Amount, A.Curr, A.amt, B.Amt_KIP, A.Broke_Date, A.Deposted_Date, A.Dep_Year, A.Dep_Month, B.TTMon, B.PrevMon, B.PrevDep, B.MonDep , B.TTDep, B.Remain, A.regist_no,B.AmountRemain,B.AmountClear,B.Amt_All,dbo.Sections.SecNmL, dbo.Sections.SecNmE, A.Sect_ID, D.DepartmentNmE, G.Group_NmE, A.Asset_NmE " & _
                              " FROM dbo.Assets AS A INNER JOIN  dbo.Sections ON A.Sect_ID = dbo.Sections.SecID LEFT OUTER JOIN dbo.Rpt_Grp AS B ON A.AssetID = B.AssetID LEFT OUTER JOIN " & _
                              " dbo.Department AS D ON D.DepartmentID = A.DepartmentID LEFT OUTER JOIN dbo.Groups AS G ON G.Group_ID = A.Group_ID Where 1=1 " & myStr & " Order by A.AssetID "
                Call LoadSqlData(sss, Rs)
                If chkBranch.Checked = True Then
                    If Lang = True Then
                        rpt = New CryRpt_GrpLongNoBrEng
                    Else
                        rpt = New CryRpt_GrpLongNoBr
                    End If

                    Call footder()
                Else

                    If Lang = False Then
                        rpt = New CryRpt_GrpLong
                    Else
                        rpt = New CryRpt_GrpLongEng
                    End If
                End If
            End If
            Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("txtPeriod"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = mPer
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Show"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = CmbShow.Text
            If Lang = False Then
                myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Head"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = "ບັນຊີຊັບສິນ " & cmbGrp.Text
            Else
                myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Head"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = "Assets List " & cmbGrp.Text
            End If

            'Dim myText1 As CrystalDecisions.CrystalReports.Engine.TextObject
            'myText1 = CType(rpt.ReportDefinition.ReportObjects.Item("txtDep"), CrystalDecisions.CrystalReports.Engine.TextObject)
            'myText1.Text = mstDep
            Dim myText3 As CrystalDecisions.CrystalReports.Engine.TextObject
            'myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("txtGrp"), CrystalDecisions.CrystalReports.Engine.TextObject)
            'myText3.Text = Trim(cmbGrp.Text)
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Tel"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = "Tel: " & OffTel & " Fax: " & OffFax
            If Lang = False Then
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = OffName
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal5"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign5
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal4"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign4
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal3"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign3
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal2"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign2
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign1
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Place"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = PlaecL
            Else
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = OffNameE
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal5"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign5e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal4"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign4e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal3"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign3e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal2"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign2e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign1e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Place"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = PlaecE
            End If

        Else
            Dim ss As String = "Select  A.Group_ID, G.Group_Nm, G.Group_NmE, sum(A.Amt) as Amt, sum(A.Amt_KIP) as Amt_KIP, sum(A.Dep_Year) as Dep_Year, sum(A.Dep_Month) as Dep_Month, sum(A.PrevMon) as PrevMon, sum(A.PrevDep) as PrevDep, sum(A.MonDep) as MonDep, sum(A.TTDep) as TTDep, sum(A.Remain) as Remain , " & _
            " sum(A.AmountRemain) as AmountRemain, sum(A.AmountClear) as AmountClear, sum(A.Amt_All) as Amt_All " & _
               " from Assets C LEFT OUTER JOIN Rpt_Grp A ON C.Group_ID=A.Group_ID LEFT OUTER JOIN Groups G ON A.Group_ID=G.Group_ID " & _
               " Where 1=1" & myStr & _
                 " GROUP By A.Group_ID, G.Group_Nm, G.Group_NmE " & _
                " ORDER by A.Group_ID "
            Call LoadSqlData(ss, Rs)
            If Lang = True Then
                rpt = New CryRpt_GrpSumLongEng
            Else
                rpt = New CryRpt_GrpSumLong
            End If

            Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
            Dim myText1, myText3 As CrystalDecisions.CrystalReports.Engine.TextObject
            myText1 = CType(rpt.ReportDefinition.ReportObjects.Item("txtSec"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText1.Text = CmbShow.Text
            If Lang = True Then
                myText1 = CType(rpt.ReportDefinition.ReportObjects.Item("Head"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText1.Text = "Assets List " & cmbGrp.Text
            Else
                myText1 = CType(rpt.ReportDefinition.ReportObjects.Item("Head"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText1.Text = "ບັນຊີຊັບສິນ " & cmbGrp.Text
            End If
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("txtPeriod"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = mPer
            myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Tel"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText3.Text = "Tel: " & OffTel & " Fax: " & OffFax
            If Lang = False Then
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = OffName
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal5"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign5
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal4"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign4
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal3"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign3
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal2"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign2
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign1
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Place"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = PlaecL
            Else
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("OfNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = OffNameE
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal5"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign5e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal4"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign4e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal3"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign3e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal2"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign2e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Signal1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = Sign1e
                myText3 = CType(rpt.ReportDefinition.ReportObjects.Item("Place"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText3.Text = PlaecE
            End If
        End If
        If Rs.RecordCount = 0 Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
        With rpt
            .SetDataSource(Rs)
            FmPreview.ReportViewer.ReportSource = rpt
            FmPreview.ReportViewer.DisplayGroupTree = False
            'FmPreview.MdiParent = APMain
            FmPreview.WindowState = FormWindowState.Maximized
            FmPreview.Show()
            FmPreview.Focus()
        End With
    End Sub

    Private Sub txtCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCode.TextChanged

    End Sub

    Private Sub CmbShow_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbShow.SelectedIndexChanged

    End Sub

    Private Sub FG1_MouseUpEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseUpEvent) Handles FGIT.MouseUpEvent
        If FGIT.Col = 1 Then
            FGIT.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
        Else
            FGIT.Editable = VSFlex8U.EditableSettings.flexEDNone
        End If
    End Sub

    Private Sub FG1_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FGIT.SelChange
        For i = 1 To FGIT.Rows - 1
            If FGIT.get_ValueMatrix(i, 1) = True Then
                Dim kk As String = "UPDATE Sections set Choose='1'  where SecID='" & FGIT.get_TextMatrix(i, 2) & "'  "
                CNN.Execute(kk)
            End If

        Next
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged

        FGIT.set_ColDataType(1, VSFlex8U.DataTypeSettings.flexDTBoolean)
        FGIT.FormatString = "^ລດ |^ເລືອກ|<ລະຫັດ|<ຊື່ພະແນກ                         "

        Call LOADFG1()
    End Sub
    Private Sub LOADFG1()
        'FG1.FormatString = "^ລດ |^ເລືອກ|<ລະຫັດ|<ຊື່ພະແນກ                         "

        Dim rs As New ADODB.Recordset
        FGIT.Rows = 1
        FGIT.FormatString = "^ລດ |^ເລືອກ |<ລະຫັດ|<ຊື່ພະແນກ                         "

        With rs
            Dim kk As String = "select *  from Sections  order by SecID ASC"
            Call LoadSqlData(kk, rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    FGIT.AddItem(.AbsolutePosition & _
                                    Chr(9) & "" & _
                    Chr(9) & (.Fields("SecID").Value.ToString) & _
                      Chr(9) & (.Fields("SecNmL").Value.ToString))
                    .MoveNext()

                End While
            Else
                FGIT.Rows = 2
            End If
        End With
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged

        FGIT.set_ColDataType(1, VSFlex8U.DataTypeSettings.flexDTBoolean)
        FGIT.FormatString = "^ລດ |^ເລືອກ|<ລະຫັດ|<ຊື່ພະແນກ                         "

        Call LOADFG1()
    End Sub

    Private Sub TxtCertify_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCertify.TextChanged

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        CNN.Execute("Update Rpt_Grp set Rpt_Grp.Company=Assets.Company ,Rpt_Grp.Section=Assets.Section ,Rpt_Grp.DepartmentID=Assets.DepartmentID  from Rpt_Grp,Assets where Assets.AssetID=Rpt_Grp.AssetID ")

        Dim SLL As New ADODB.Recordset
        Dim MMDate As Date = CDate("01/" & Trim(DTMon.Value.Month.ToString) & "/" & Trim(DTMon.Value.Year.ToString))
        If optMon.Checked = False Then
            MsgBox("ກະລຸນາເລືອກປະຈໍາເດືອນກ່ອນ") : Exit Sub
        End If
        If cmbSec.Text = "" Then MsgBox("ກະລຸນາເລືອກສາຂາກ່ອນຈະໂອນ") : Exit Sub
        If cmbGrp.SelectedIndex = 0 Then
            MsgBox("ກະລຸນາເລືອກໝວດຊັບສິນທີ່ທ່ານຕ້ອງການຈະໂອນກ່ອນ") : cmbGrp.Focus() : Exit Sub
        End If
        If optMon.Checked = True Then
            If MsgBox("ທ່ານຕ້ອງການໂອນໝວດຊັບສິນ " & Trim(cmbGrp.Text) & " ນີ້ບໍ່?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo + MsgBoxStyle.Question) = Windows.Forms.DialogResult.Yes Then
                Dim KK As String = "select * from gen_jn where company='" & Trim(txtCompany.Text) & "' and date_work='" & Format(MMDate, "yyyy-MM-dd") & "' and certify='" & Trim(TxtCertify.Text) & "' "
                Call LoadSqlData(KK, SLL)
                If SLL.RecordCount = 0 Then
                    'CalcMonCOMPANY()
                    Dim kk2 As String = "SElecT sum(MonDep) as MonDep,Group_ID,section from  Rpt_Grp where  Group_ID='" & Trim(txtGrp.Text) & "' and section='" & Trim(txtSec.Text) & "' group by Group_ID,section "
                    Call LoadSqlData(kk2, RSC)
                    If RSC.RecordCount > 0 Then
                        Dim Lng As String = "INSERT INTO gen_jn(certify,Referno,company,Com_id,don_id,office_id, Book, amount, net_amt, date_work, code_dr,code_cr,ac_code,amount_dr,amount_cr,amt_dr,amt_cr,last_update,last_user,my_lock,amt_USD_dr,amt_USD_cr)  " & _
                                           " VALUES('" & Trim(TxtCertify.Text) & "','" & Trim(TxtCertify.Text) & "','" & Trim(txtCompany.Text) & "','" & Trim(txtCompany.Text) & "','01','" & Trim(txtCompany.Text) & "','Fixd Asset'," & RSC.Fields("MonDep").Value & "," & RSC.Fields("MonDep").Value & ",'" & Format(MMDate, "yyyy-MM-dd") & "','" & txtAcc.Text & "','','" & txtAcc.Text & "'," & RSC.Fields("MonDep").Value & ",0," & RSC.Fields("MonDep").Value & ",0,'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "','0','0','0')"
                        CNN.Execute(Lng)
                        Dim Lng1 As String = "INSERT INTO gen_jn(certify,Referno,company, Com_id,don_id,office_id, Book, amount, net_amt, date_work, code_dr,code_cr,ac_code,amount_dr,amount_cr,amt_dr,amt_cr,last_update,last_user,my_lock,amt_USD_dr,amt_USD_cr)  " & _
                                           " VALUES('" & Trim(TxtCertify.Text) & "','" & Trim(TxtCertify.Text) & "','" & Trim(txtCompany.Text) & "','" & Trim(txtCompany.Text) & "','01','" & Trim(txtCompany.Text) & "','Fixd Asset'," & RSC.Fields("MonDep").Value & "," & RSC.Fields("MonDep").Value & ",'" & Format(MMDate, "yyyy-MM-dd") & "','','" & TxtLH.Text & "','" & TxtLH.Text & "',0," & RSC.Fields("MonDep").Value & ",0," & RSC.Fields("MonDep").Value & ",'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "','0','0','0')"
                        CNN.Execute(Lng1)
                    End If
                Else
                    Dim De As String = "Delete from gen_jn where company='" & Trim(txtCompany.Text) & "' and month(date_work)='" & (MMDate.Month) & "' and year(date_work)='" & (MMDate.Year) & "'"
                    CNN.Execute(De)
                    'CalcMonCOMPANY()
                    Call LoadSqlData("SElecT sum(MonDep) as MonDep,Group_ID from  Rpt_Grp where  Group_ID='" & Trim(txtGrp.Text) & "'  and section='" & Trim(txtSec.Text) & "' group by Group_ID,section ", RSC)
                    If RSC.RecordCount > 0 Then
                        Dim Lng As String = "INSERT INTO gen_jn(certify,Referno,company,Com_id,don_id,office_id, Book, amount, net_amt, date_work, code_dr,code_cr,ac_code,amount_dr,amount_cr,amt_dr,amt_cr,last_update,last_user,my_lock,amt_USD_dr,amt_USD_cr)  " & _
                                        " VALUES('" & Trim(TxtCertify.Text) & "','" & Trim(TxtCertify.Text) & "','" & Trim(txtCompany.Text) & "','" & Trim(txtCompany.Text) & "','01','" & Trim(txtCompany.Text) & "','Fixd Asset'," & RSC.Fields("MonDep").Value & "," & RSC.Fields("MonDep").Value & ",'" & Format(MMDate, "yyyy-MM-dd") & "','" & txtAcc.Text & "','','" & txtAcc.Text & "'," & RSC.Fields("MonDep").Value & ",0," & RSC.Fields("MonDep").Value & ",0,'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "','0','0','0')"
                        CNN.Execute(Lng)
                        Dim Lng1 As String = "INSERT INTO gen_jn(certify,Referno,company, Com_id,don_id,office_id, Book, amount, net_amt, date_work, code_dr,code_cr,ac_code,amount_dr,amount_cr,amt_dr,amt_cr,last_update,last_user,my_lock,amt_USD_dr,amt_USD_cr)  " & _
                                           " VALUES('" & Trim(TxtCertify.Text) & "','" & Trim(TxtCertify.Text) & "','" & Trim(txtCompany.Text) & "','" & Trim(txtCompany.Text) & "','01','" & Trim(txtCompany.Text) & "','Fixd Asset'," & RSC.Fields("MonDep").Value & "," & RSC.Fields("MonDep").Value & ",'" & Format(MMDate, "yyyy-MM-dd") & "','','" & TxtLH.Text & "','" & TxtLH.Text & "',0," & RSC.Fields("MonDep").Value & ",0," & RSC.Fields("MonDep").Value & ",'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "','0','0','0')"
                        CNN.Execute(Lng1)
                    End If
                End If
                CNN.Execute("update Gen_jn set Gen_jn.descrip=Acc_Code.Name_L, Gen_jn.ac_name=Acc_Code.Name_L, Gen_jn.curr='LAK', Gen_jn.rate=1,Gen_jn.rate_i=1, Gen_jn.ac_typee=Acc_Code.Acc_TypeE from Acc_Code,Gen_jn where Gen_jn.certify='" & Trim(TxtCertify.Text) & "' and Gen_jn.AC_Code=ACC_Code.AC_Code ")
                MsgBox("ການໂອນສຳເລັດຜົນ")
            End If

        ElseIf optTerm.Checked = True Then
            MsgBox("ບໍ່ສາມາດໂອນເປັນງວດໄດ້")
        ElseIf optYear.Checked = True Then
            MsgBox("ບໍ່ສາມາດໂອນເປັນປີໄດ້")
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If txtSec.Text = "" Then
            MsgBox("ກະລຸນາເລືອກສໍານັກງານກ່ອນ", MsgBoxStyle.Exclamation) : cmbSec.Focus() : Exit Sub
        End If
        Dim GEN, CNDR, CNCr As String
        CNN.Execute("Update Rpt_Grp set Rpt_Grp.Company=Assets.Company ,Rpt_Grp.Section=Assets.Section ,Rpt_Grp.DepartmentID=Assets.DepartmentID  from Rpt_Grp,Assets where Assets.AssetID=Rpt_Grp.AssetID ")

        Dim SLL As New ADODB.Recordset
        Dim MMDate As Date = CDate("01/" & Trim(DTMon.Value.Month.ToString) & "/" & Trim(DTMon.Value.Year.ToString))
        If optMon.Checked = False Then
            MsgBox("ກະລຸນາເລືອກປະຈໍາເດືອນກ່ອນ") : Exit Sub
        End If
        If cmbSec.Text = "" Then MsgBox("ກະລຸນາເລືອກສາຂາກ່ອນຈະໂອນ") : Exit Sub
        If cmbGrp.SelectedIndex = 0 Then
            MsgBox("ກະລຸນາເລືອກໝວດຊັບສິນທີ່ທ່ານຕ້ອງການຈະໂອນກ່ອນ") : cmbGrp.Focus() : Exit Sub
        End If

        If txtAcc.Text = "" Then MsgBox("ກະລຸນາໃສ່ເລກບັນຊີເບື້ອງໜີກ່ອນ", MsgBoxStyle.Exclamation) : txtAcc.Focus() : Exit Sub
        If TxtLH.Text = "" Then MsgBox("ກະລຸນາໃສ່ເລກບັນຊີເບື້ອງມີກ່ອນ", MsgBoxStyle.Exclamation) : TxtLH.Focus() : Exit Sub

        Dim gRS As New ADODB.Recordset
        Call LoadSqlData("select * from Acc_Code Where Ac_Code=N'" & Trim(txtAcc.Text) & "'", gRS)
        If gRS.RecordCount = 0 Then
            MsgBox("ເລກບັນຊີບໍ່ມີໃນສາລະບານ " & Trim(TxtLH.Text), MsgBoxStyle.Exclamation) : txtAcc.Focus() : Exit Sub
        End If
        Call LoadSqlData("select * from Acc_Code Where Ac_Code=N'" & Trim(TxtLH.Text) & "'", gRS)
        If gRS.RecordCount = 0 Then
            MsgBox("ເລກບັນຊີບໍ່ມີໃນສາລະບານ " & Trim(TxtLH.Text), MsgBoxStyle.Exclamation) : TxtLH.Focus() : Exit Sub
        End If


        If optMon.Checked = True Then
            If MsgBox("ທ່ານຕ້ອງການໂອນໝວດຊັບສິນ " & Trim(cmbGrp.Text) & " ນີ້ບໍ່?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo + MsgBoxStyle.Question) = Windows.Forms.DialogResult.Yes Then
                Dim KK As String = "select * from gen_jn where office_id='" & Trim(MuSubOff2) & "' and date_work='" & Format(MMDate, "yyyy-MM-dd") & "' and certify=N'" & Trim(TxtCertify.Text) & "' "
                Call LoadSqlData(KK, SLL)
                If SLL.RecordCount = 0 Then

                    Dim kk2 As String = "SElecT sum(MonDep) as MonDep,Group_ID,section from  Rpt_Grp where  Group_ID='" & Trim(txtGrp.Text) & "' and section='" & Trim(txtSec.Text) & "' group by Group_ID,section "
                    Call LoadSqlData(kk2, RSC)
                    If RSC.RecordCount > 0 Then
                        'Dim Lng As String = "INSERT INTO gen_jn(certify,Referno,company,Com_id,don_id,office_id, Book, amount, net_amt, date_work, code_dr,code_cr,ac_code,amount_dr,amount_cr,amt_dr,amt_cr,last_update,last_user,my_lock,amt_USD_dr,amt_USD_cr)  " & _
                        '                   " VALUES('" & Trim(TxtCertify.Text) & "','" & Trim(TxtCertify.Text) & "','" & Trim(txtCompany.Text) & "','" & Trim(txtCompany.Text) & "','01','" & Trim(txtCompany.Text) & "','Fixd Asset'," & RSC.Fields("MonDep").Value & "," & RSC.Fields("MonDep").Value & ",'" & Format(MMDate, "yyyy-MM-dd") & "','" & txtAcc.Text & "','','" & txtAcc.Text & "'," & RSC.Fields("MonDep").Value & ",0," & RSC.Fields("MonDep").Value & ",0,'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "','0','0','0')"
                        'CNN.Execute(Lng)
                        'Dim Lng1 As String = "INSERT INTO gen_jn(certify,Referno,company, Com_id,don_id,office_id, Book, amount, net_amt, date_work, code_dr,code_cr,ac_code,amount_dr,amount_cr,amt_dr,amt_cr,last_update,last_user,my_lock,amt_USD_dr,amt_USD_cr)  " & _
                        '                   " VALUES('" & Trim(TxtCertify.Text) & "','" & Trim(TxtCertify.Text) & "','" & Trim(txtCompany.Text) & "','" & Trim(txtCompany.Text) & "','01','" & Trim(txtCompany.Text) & "','Fixd Asset'," & RSC.Fields("MonDep").Value & "," & RSC.Fields("MonDep").Value & ",'" & Format(MMDate, "yyyy-MM-dd") & "','','" & TxtLH.Text & "','" & TxtLH.Text & "',0," & RSC.Fields("MonDep").Value & ",0," & RSC.Fields("MonDep").Value & ",'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "','0','0','0')"
                        'CNN.Execute(Lng1)

                        ' =============AP_ACC_Gen=====================
                        GEN = "INSERT INTO AP_ACC_Gen(certify,date_work, book,Referno, cheque_no,descrip,amount,Curr,rate,net_amt,  AmountDr, AmountCr, " & _
                        " TotalAmountDr, TotalAmountCr, my_lock,Amount_Later, last_update, last_user,Com_id,office_id) " & _
                          " VALUES(N'" & Trim(TxtCertify.Text) & "'," & _
                        " '" & Format(CDate(MMDate), "yyyy-MM-dd") & "'," & _
                               "N'Fixd Asset'," & _
                          "N'" & Trim(TxtCertify.Text) & "'," & _
                                         "N''," & _
                                   "N'ຫຼັກຄ່າເຊື່ອມມູນຄ່າ'," & _
                                       "" & RSC.Fields("MonDep").Value & "," & _
                            "N'" & Trim(cmbCurr.Text) & "'," & _
                               "" & CDbl(Exchange.Text) & "," & _
                                  "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                          "" & RSC.Fields("MonDep").Value & "," & _
                          "" & RSC.Fields("MonDep").Value & "," & _
                           "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                            "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                           " 1," & _
                             "N''," & _
                          " Getdate()," & _
                        "N'" & MUserID & "'," & _
                        "'01','01' )"
                        CNN.Execute(GEN)
                        '        '==========dr =====

                        '        CNDR = "INSERT INTO AP_ACC_Gen_Item(certify,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                        '        " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id) " & _
                        '          " VALUES(N'" & Trim(TxtCertify.Text) & "'," & _
                        '        " '" & Format(CDate(MMDate), "yyyy-MM-dd") & "'," & _
                        '           "N'Fixd Asset'," & _
                        '          "N'" & Trim(TxtCertify.Text) & "'," & _
                        '            "N'" & Trim(TxtCertify.Text) & "'," & _
                        '                         "N''," & _
                        '                       "" & RSC.Fields("MonDep").Value & "," & _
                        '            "N'" & Trim(cmbCurr.Text) & "'," & _
                        '               "" & CDbl(Exchange.Text) & "," & _
                        '                 "N'" & Trim(cmbCurr.Text) & "'," & _
                        '               "" & CDbl(Exchange.Text) & "," & _
                        '                  "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                        '          "N'" & Trim(txtAcc.Text) & "'," & _
                        '           "N''," & _
                        '         "N'" & Trim(txtAcc.Text) & "'," & _
                        '         "N''," & _
                        '          "" & RSC.Fields("MonDep").Value & "," & _
                        '          " 0," & _
                        '            "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                        '          " 0," & _
                        '             " 0," & _
                        '                " 0," & _
                        '           " 1," & _
                        '               " 1," & _
                        '          " Getdate()," & _
                        '        "N'" & MUserID & "'," & _
                        '        "'01'  )"
                        '        CNN.Execute(CNDR)
                        '        '=======CR======================
                        '        CNCr = "INSERT INTO AP_ACC_Gen_Item(certify,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                        '        " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id) " & _
                        '          " VALUES(N'" & Trim(TxtCertify.Text) & "'," & _
                        '        " '" & Format(CDate(MMDate), "yyyy-MM-dd") & "'," & _
                        '           "N'Fixd Asset'," & _
                        '          "N'" & Trim(TxtCertify.Text) & "'," & _
                        '            "N'" & Trim(TxtCertify.Text) & "'," & _
                        '                         "N''," & _
                        '                       "" & RSC.Fields("MonDep").Value & "," & _
                        '            "N'" & Trim(cmbCurr.Text) & "'," & _
                        '               "" & CDbl(Exchange.Text) & "," & _
                        '                 "N'" & Trim(cmbCurr.Text) & "'," & _
                        '               "" & CDbl(Exchange.Text) & "," & _
                        '                  "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                        '                                     "N''," & _
                        '          "N'" & Trim(TxtLH.Text) & "'," & _
                        '         "N'" & Trim(TxtLH.Text) & "'," & _
                        '         "N''," & _
                        '" 0," & _
                        '          "" & RSC.Fields("MonDep").Value & "," & _
                        '          " 0," & _
                        '            "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                        '          " 0," & _
                        '             " 0," & _
                        '           " 1," & _
                        '               " 1," & _
                        '          " Getdate()," & _
                        '        "N'" & MUserID & "'," & _
                        '        "'01'  )"
                        '        CNN.Execute(CNCr)
                        '==========dr =====

                        CNDR = "INSERT INTO gen_jn(certify,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                        " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id) " & _
                          " VALUES(N'" & Trim(TxtCertify.Text) & "'," & _
                        " '" & Format(CDate(MMDate), "yyyy-MM-dd") & "'," & _
                           "N'Fixd Asset'," & _
                          "N'" & Trim(TxtCertify.Text) & "'," & _
                            "N'" & Trim(TxtCertify.Text) & "'," & _
                                         "N''," & _
                                       "" & RSC.Fields("MonDep").Value & "," & _
                            "N'" & Trim(cmbCurr.Text) & "'," & _
                               "" & CDbl(Exchange.Text) & "," & _
                                 "N'" & Trim(cmbCurr.Text) & "'," & _
                               "" & CDbl(Exchange.Text) & "," & _
                                  "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                          "N'" & Trim(txtAcc.Text) & "'," & _
                           "N''," & _
                         "N'" & Trim(txtAcc.Text) & "'," & _
                         "N''," & _
                          "" & RSC.Fields("MonDep").Value & "," & _
                          " 0," & _
                            "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                          " 0," & _
                             " 0," & _
                                " 0," & _
                           " 1," & _
                               " 1," & _
                          " Getdate()," & _
                        "N'" & MUserID & "'," & _
                        "N'" & MuSubOff2 & "'   )"
                        CNN.Execute(CNDR)
                        '=======CR======================
                        CNCr = "INSERT INTO gen_jn(certify,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                        " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id) " & _
                          " VALUES(N'" & Trim(TxtCertify.Text) & "'," & _
                        " '" & Format(CDate(MMDate), "yyyy-MM-dd") & "'," & _
                           "N'Fixd Asset'," & _
                          "N'" & Trim(TxtCertify.Text) & "'," & _
                            "N'" & Trim(TxtCertify.Text) & "'," & _
                                         "N''," & _
                                       "" & RSC.Fields("MonDep").Value & "," & _
                            "N'" & Trim(cmbCurr.Text) & "'," & _
                               "" & CDbl(Exchange.Text) & "," & _
                                 "N'" & Trim(cmbCurr.Text) & "'," & _
                               "" & CDbl(Exchange.Text) & "," & _
                                  "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                                                     "N''," & _
                          "N'" & Trim(TxtLH.Text) & "'," & _
                         "N'" & Trim(TxtLH.Text) & "'," & _
                         "N''," & _
                " 0," & _
                          "" & RSC.Fields("MonDep").Value & "," & _
                          " 0," & _
                            "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                          " 0," & _
                             " 0," & _
                           " 1," & _
                               " 1," & _
                          " Getdate()," & _
                        "N'" & MUserID & "'," & _
                        "N'" & MuSubOff2 & "'   )"
                        CNN.Execute(CNCr)
                    End If
                Else
                    Dim DeGen As String = "Delete from AP_ACC_Gen  where certify=N'" & Trim(TxtCertify.Text) & "' and office_id='" & MuSubOff2 & "' and month(date_work)='" & (MMDate.Month) & "' and year(date_work)='" & (MMDate.Year) & "'"
                    CNN.Execute(DeGen)
                    Dim De As String = "Delete from AP_ACC_Gen_Item where certify=N'" & Trim(TxtCertify.Text) & "' and  office_id='" & MuSubOff2 & "' and month(date_work)='" & (MMDate.Month) & "' and year(date_work)='" & (MMDate.Year) & "'"
                    CNN.Execute(De)
                    Dim Dejn As String = "Delete from gen_jn where certify=N'" & Trim(TxtCertify.Text) & "' and  office_id='" & MuSubOff2 & "' and month(date_work)='" & (MMDate.Month) & "' and year(date_work)='" & (MMDate.Year) & "'"
                    CNN.Execute(Dejn)

                    Call LoadSqlData("SElecT sum(MonDep) as MonDep,Group_ID from  Rpt_Grp where  Group_ID='" & Trim(txtGrp.Text) & "'  and section='" & Trim(txtSec.Text) & "' group by Group_ID,section ", RSC)
                    If RSC.RecordCount > 0 Then

                        'Dim Lng As String = "INSERT INTO gen_jn(certify,Referno,company,Com_id,don_id,office_id, Book, amount, net_amt, date_work, code_dr,code_cr,ac_code,amount_dr,amount_cr,amt_dr,amt_cr,last_update,last_user,my_lock,amt_USD_dr,amt_USD_cr)  " & _
                        '                " VALUES('" & Trim(TxtCertify.Text) & "','" & Trim(TxtCertify.Text) & "','" & Trim(txtCompany.Text) & "','" & Trim(txtCompany.Text) & "','01','" & Trim(txtCompany.Text) & "','Fixd Asset'," & RSC.Fields("MonDep").Value & "," & RSC.Fields("MonDep").Value & ",'" & Format(MMDate, "yyyy-MM-dd") & "','" & txtAcc.Text & "','','" & txtAcc.Text & "'," & RSC.Fields("MonDep").Value & ",0," & RSC.Fields("MonDep").Value & ",0,'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "','0','0','0')"
                        'CNN.Execute(Lng)
                        'Dim Lng1 As String = "INSERT INTO gen_jn(certify,Referno,company, Com_id,don_id,office_id, Book, amount, net_amt, date_work, code_dr,code_cr,ac_code,amount_dr,amount_cr,amt_dr,amt_cr,last_update,last_user,my_lock,amt_USD_dr,amt_USD_cr)  " & _
                        '                   " VALUES('" & Trim(TxtCertify.Text) & "','" & Trim(TxtCertify.Text) & "','" & Trim(txtCompany.Text) & "','" & Trim(txtCompany.Text) & "','01','" & Trim(txtCompany.Text) & "','Fixd Asset'," & RSC.Fields("MonDep").Value & "," & RSC.Fields("MonDep").Value & ",'" & Format(MMDate, "yyyy-MM-dd") & "','','" & TxtLH.Text & "','" & TxtLH.Text & "',0," & RSC.Fields("MonDep").Value & ",0," & RSC.Fields("MonDep").Value & ",'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "','0','0','0')"
                        'CNN.Execute(Lng1)

                        ' =============AP_ACC_Gen=====================
                        GEN = "INSERT INTO AP_ACC_Gen(certify,date_work, book,Referno, cheque_no,descrip,amount,Curr,rate,net_amt,  AmountDr, AmountCr, " & _
                        " TotalAmountDr, TotalAmountCr, my_lock,Amount_Later, last_update, last_user,Com_id,office_id) " & _
                          " VALUES(N'" & Trim(TxtCertify.Text) & "'," & _
                        " '" & Format(CDate(MMDate), "yyyy-MM-dd") & "'," & _
                               "N'Fixd Asset'," & _
                          "N'" & Trim(TxtCertify.Text) & "'," & _
                                         "N''," & _
                                   "N'ຫຼັກຄ່າເຊື່ອມມູນຄ່າ'," & _
                                       "" & RSC.Fields("MonDep").Value & "," & _
                            "N'" & Trim(cmbCurr.Text) & "'," & _
                               "" & CDbl(Exchange.Text) & "," & _
                                  "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                          "" & RSC.Fields("MonDep").Value & "," & _
                          "" & RSC.Fields("MonDep").Value & "," & _
                           "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                            "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                           " 1," & _
                             "N''," & _
                          " Getdate()," & _
                        "N'" & MUserID & "'," & _
                        "N'" & MuSubOff2 & "' ,N'" & MuSubOff2 & "' )"
                        CNN.Execute(GEN)
                        '==========dr =====

                        CNDR = "INSERT INTO gen_jn(certify,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                        " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id,Company,Del,Lock) " & _
                          " VALUES(N'" & Trim(TxtCertify.Text) & "'," & _
                        " '" & Format(CDate(MMDate), "yyyy-MM-dd") & "'," & _
                           "N'Fixd Asset'," & _
                          "N'" & Trim(TxtCertify.Text) & "'," & _
                            "N'" & Trim(TxtCertify.Text) & "'," & _
                                         "N''," & _
                                       "" & RSC.Fields("MonDep").Value & "," & _
                            "N'" & Trim(cmbCurr.Text) & "'," & _
                               "" & CDbl(Exchange.Text) & "," & _
                                 "N'" & Trim(cmbCurr.Text) & "'," & _
                               "" & CDbl(Exchange.Text) & "," & _
                                  "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                          "N'" & Trim(txtAcc.Text) & "'," & _
                           "N''," & _
                         "N'" & Trim(txtAcc.Text) & "'," & _
                         "N''," & _
                          "" & RSC.Fields("MonDep").Value & "," & _
                          " 0," & _
                            "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                          " 0," & _
                             " 0," & _
                                " 0," & _
                           " 1," & _
                               " 1," & _
                          " Getdate()," & _
                        "N'" & MUserID & "'," & _
                        "N'" & MuSubOff2 & "' ,N'" & MuSubOff2 & "',1,0  )"
                        CNN.Execute(CNDR)
                        '=======CR======================
                        CNCr = "INSERT INTO gen_jn(certify,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                        " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id,Company,Del,Lock) " & _
                          " VALUES(N'" & Trim(TxtCertify.Text) & "'," & _
                        " '" & Format(CDate(MMDate), "yyyy-MM-dd") & "'," & _
                           "N'Fixd Asset'," & _
                          "N'" & Trim(TxtCertify.Text) & "'," & _
                            "N'" & Trim(TxtCertify.Text) & "'," & _
                                         "N''," & _
                                       "" & RSC.Fields("MonDep").Value & "," & _
                            "N'" & Trim(cmbCurr.Text) & "'," & _
                               "" & CDbl(Exchange.Text) & "," & _
                                 "N'" & Trim(cmbCurr.Text) & "'," & _
                               "" & CDbl(Exchange.Text) & "," & _
                                  "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                                                     "N''," & _
                          "N'" & Trim(TxtLH.Text) & "'," & _
                         "N'" & Trim(TxtLH.Text) & "'," & _
                         "N''," & _
                " 0," & _
                          "" & RSC.Fields("MonDep").Value & "," & _
                          " 0," & _
                            "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                          " 0," & _
                             " 0," & _
                           " 1," & _
                               " 1," & _
                          " Getdate()," & _
                        "N'" & MUserID & "'," & _
                      "N'" & MuSubOff2 & "' ,N'" & MuSubOff2 & "',1,0  )"
                        CNN.Execute(CNCr)
                        '==========dr =====

                        '        CNDR = "INSERT INTO AP_ACC_Gen_Item(certify,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                        '        " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id) " & _
                        '          " VALUES(N'" & Trim(TxtCertify.Text) & "'," & _
                        '        " '" & Format(CDate(MMDate), "yyyy-MM-dd") & "'," & _
                        '           "N'Fixd Asset'," & _
                        '          "N'" & Trim(TxtCertify.Text) & "'," & _
                        '            "N'" & Trim(TxtCertify.Text) & "'," & _
                        '                         "N''," & _
                        '                       "" & RSC.Fields("MonDep").Value & "," & _
                        '            "N'" & Trim(cmbCurr.Text) & "'," & _
                        '               "" & CDbl(Exchange.Text) & "," & _
                        '                 "N'" & Trim(cmbCurr.Text) & "'," & _
                        '               "" & CDbl(Exchange.Text) & "," & _
                        '                  "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                        '          "N'" & Trim(txtAcc.Text) & "'," & _
                        '           "N''," & _
                        '         "N'" & Trim(txtAcc.Text) & "'," & _
                        '         "N''," & _
                        '          "" & RSC.Fields("MonDep").Value & "," & _
                        '          " 0," & _
                        '            "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                        '          " 0," & _
                        '             " 0," & _
                        '                " 0," & _
                        '           " 1," & _
                        '               " 1," & _
                        '          " Getdate()," & _
                        '        "N'" & MUserID & "'," & _
                        '        "'01'  )"
                        '        CNN.Execute(CNDR)
                        '        '=======CR======================
                        '        CNCr = "INSERT INTO AP_ACC_Gen_Item(certify,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                        '        " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id) " & _
                        '          " VALUES(N'" & Trim(TxtCertify.Text) & "'," & _
                        '        " '" & Format(CDate(MMDate), "yyyy-MM-dd") & "'," & _
                        '           "N'Fixd Asset'," & _
                        '          "N'" & Trim(TxtCertify.Text) & "'," & _
                        '            "N'" & Trim(TxtCertify.Text) & "'," & _
                        '                         "N''," & _
                        '                       "" & RSC.Fields("MonDep").Value & "," & _
                        '            "N'" & Trim(cmbCurr.Text) & "'," & _
                        '               "" & CDbl(Exchange.Text) & "," & _
                        '                 "N'" & Trim(cmbCurr.Text) & "'," & _
                        '               "" & CDbl(Exchange.Text) & "," & _
                        '                  "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                        '                                     "N''," & _
                        '          "N'" & Trim(TxtLH.Text) & "'," & _
                        '         "N'" & Trim(TxtLH.Text) & "'," & _
                        '         "N''," & _
                        '" 0," & _
                        '          "" & RSC.Fields("MonDep").Value & "," & _
                        '          " 0," & _
                        '            "" & RSC.Fields("MonDep").Value * CDbl(Exchange.Text) & "," & _
                        '          " 0," & _
                        '             " 0," & _
                        '           " 1," & _
                        '               " 1," & _
                        '          " Getdate()," & _
                        '        "N'" & MUserID & "'," & _
                        '        "'01'  )"
                        '        CNN.Execute(CNCr)

                    End If
                End If
                CNN.Execute("update AP_ACC_Gen_Item set  AP_ACC_Gen_Item.descrip=Acc_Code.Name_L, AP_ACC_Gen_Item.ac_name=Acc_Code.Name_L,  AP_ACC_Gen_Item.ac_typee=Acc_Code.Acc_TypeE from Acc_Code,AP_ACC_Gen_Item where AP_ACC_Gen_Item.certify='" & Trim(TxtCertify.Text) & "' and AP_ACC_Gen_Item.AC_Code=ACC_Code.AC_Code ")

                MsgBox("ການໂອນສຳເລັດຜົນ")
            End If

        ElseIf optTerm.Checked = True Then
            MsgBox("ບໍ່ສາມາດໂອນເປັນງວດໄດ້")
        ElseIf optYear.Checked = True Then
            MsgBox("ບໍ່ສາມາດໂອນເປັນປີໄດ້")
        End If
    End Sub
 
    Private Sub BtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch.Click
        fmShartOfAccDetail.txtSty.Text = "FrmRpt_Group_DR"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
 
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        fmShartOfAccDetail.txtSty.Text = "FrmRpt_Group_CR"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
    End Sub

    Private Sub txtAcc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAcc.TextChanged

    End Sub
End Class