Public Class FrmAssetNew

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim cRS As New ADODB.Recordset
        Dim nRS As New ADODB.Recordset
        'If txtCertify.Text = "" Then MsgBox("ກະລຸນາໃສ່ເລກບິນກ່ອນ", MsgBoxStyle.Question) : txtCertify.Focus() : Exit Sub
        If txtNm.Text = "" Then MsgBox("ກະລຸນາໃສ່ຊື່ລາຍການພາສາລາວກ່ອນ") : txtNm.Focus() : Exit Sub
        'If txtNmE.Text = "" Then MsgBox("ກະລຸນາໃສ່ຊື່ລາຍການພາສາອັງກິດກ່ອນ") : txtNmE.Focus() : Exit Sub
        If txtID.Enabled = True Then
            RunID()
        End If
        Call LoadSqlData("Select AssetID From Assets Where AssetID='" & Trim(txtID.Text) & "' ", cRS)

        If cRS.RecordCount <> 0 Then
            Dim ss As String
            ss = "Update Assets Set Bar_Code=N'" & Trim(Barcode.Text) & "',Vendor=N'" & TxtVendor.Text & "',Area=N'" & TxtArea.Text & "', Budget=N'" & txtBudget.Text & "',Asset_Nm=N'" & Trim(txtNm.Text) & "',Asset_NmE=N'" & Trim(txtNmE.Text) & "', Group_ID='" & Trim(txtGrp.Text) & "', Group_Nm=N'" & Trim(cmbGrp.Text) & "', Date_Buy='" & Format(DTBuy.Value, "yyyy-MM-dd") & "', Date_Work='" & Format(DTUse.Value, "yyyy-MM-dd") & "',  Qty=" & CDbl(txtQty.Text) & ", Unit=N'" & Trim(txtUnit.Text) & "', Section='" & Trim(txtSec.Text) & "', Using_By=N'" & Trim(TxtHelp3.Text) & "', " & _
                        " Price=" & CDbl(txtPrice.Text) & ", Curr='" & Trim(cmbCurr.Text) & "', Amount=" & CDbl(txtAmt.Text) & ", Rate=" & CDbl(txtRate.Text) & ", Amt=" & CDbl(txtKIP.Text) & ",Amt_KIP=" & CDbl(txtKIP.Text) & ", Used_Life=" & CDbl(txtLife.Text) & ", Dep_Year=" & CDbl(txtYear.Text) & ", Dep_Month=" & CDbl(txtMon.Text) & ", Dep_Day=" & CDbl(txtDay.Text) & ", Remark=N'" & Trim(txtRemark.Text) & "', Engin_No=N'" & Trim(txtEngine.Text) & "', Frame_No=N'" & Trim(txtFrame.Text) & "', Regist_No=N'" & Trim(txtReg.Text) & "', Serial=N'" & Trim(txtSerial.Text) & "', " & _
                        " DepCalc=" & IIf(chkDep.Checked = True, 1, 0) & ",Make_Buy=" & IIf(Chk_Make.Checked = True, 1, 0) & ", Model=N'" & Trim(CmdModelType.Text) & "', Certify=N'" & Trim(txtCertify.Text) & "', Last_Update='" & Format(Date.Today, "yyyy-MM-dd") & "',Last_User='" & MUserID & "', DepartmentID='" & Trim(txtDep.Text) & "', Sect_ID= '" & Trim(txtSec.Text) & "', Asset_No=N'" & Trim(txtCode.Text.ToString) & "', DepartmentNm=N'" & Trim(cmbDeprt.Text) & "', txt_disc=" & CDbl(txt_disc.Text) & ", txt_Rem=" & CDbl(txt_Rem.Text) & " " & _
                        " Where AssetID='" & Trim(txtID.Text.ToString) & "' "
            CNN.Execute(ss)

            '   CNN1.Execute("DELETE FROM gen_jn where Book='Fixd Asset' and certify='" & AANO & "' ")
            '   CNN1.Execute("INSERT INTO gen_jn(certify,company, Book,amount, net_amt, date_work, code_dr,code_cr,ac_code, descrip,amount_dr,amount_cr,amt_dr,amt_cr,curr,Rate,last_update,last_user) " & _
            '  " VALUES('" & Trim(txtCertify.Text) & "','" & Trim(txtCompany.Text) & "','Fixd Asset', " & CDbl(txtPrice.Text) & "," & CDbl(txtKIP.Text) & ",'" & Format(DTBuy.Value, "yyyy-MM-dd") & "','" & txtAcc.Text & "','','" & txtAcc.Text & "',N'" & cmbGrp.Text & "'," & CDbl(txtPrice.Text) & ",0," & CDbl(txtKIP.Text) & ",0, " & _
            '  " '" & cmbCurr.Text & "'," & CDbl(txtRate.Text) & ",'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "')")
            '   CNN1.Execute("INSERT INTO gen_jn(certify,company, Book,amount,net_amt,date_work,code_dr,code_cr,ac_code, descrip,amount_dr,amount_cr,amt_dr,amt_cr,curr,Rate,last_update,last_user)  " & _
            '" VALUES('" & Trim(txtCertify.Text.ToString) & "','" & Trim(txtCompany.Text) & "','Fixd Asset', " & CDbl(txtPrice.Text) & "," & CDbl(txtKIP.Text) & ",'" & Format(DTBuy.Value, "yyyy-MM-dd") & "','','" & txtGrp.Text & "','" & txtGrp.Text & "',N'" & cmbGrp.Text & "',0," & CDbl(txtPrice.Text) & ",0," & CDbl(txtKIP.Text) & ", " & _
            '   " '" & cmbCurr.Text & "'," & CDbl(txtRate.Text) & ",'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "')")
        Else
            RunID()
            Dim PPP As String = "INSERT INTO Assets(AssetID,Bar_Code,Vendor,Area,DepartmentNm,Budget,Asset_NmE, Asset_No, Ac_Code, Asset_Nm, Group_ID, Group_Nm, Date_Buy, Date_Work,  Qty, Unit, Section, Using_By, Price, Curr, Amount, Rate, Amt, Amt_KIP, Used_Life, Dep_Year, Dep_Month,Dep_Day, Broked, Deposted, Remark, Engin_No, Frame_No, Regist_No, Serial, DepCalc, Make_Buy, Model, Certify, Last_Update,Last_User, Company, DepartmentID, Sect_ID, txt_disc, txt_Rem) " & _
                            " VALUES('" & Trim(txtID.Text.ToString) & "','" & Trim(Barcode.Text.ToString) & "',N'" & Trim(TxtVendor.Text) & "',N'" & Trim(TxtArea.Text) & "', N'" & Trim(cmbDeprt.Text) & "',N'" & Trim(txtBudget.Text) & "',N'" & Trim(txtNmE.Text) & "',N'" & Trim(txtCode.Text) & "',  '" & Trim(txtAcc.Text.ToString) & "', N'" & Trim(txtNm.Text) & "', '" & Trim(txtGrp.Text) & "', N'" & Trim(cmbGrp.Text) & "', '" & Format(DTBuy.Value, "yyyy-MM-dd") & "', '" & Format(DTUse.Value, "yyyy-MM-dd") & "', '" & CDbl(txtQty.Text) & "', N'" & Trim(txtUnit.Text) & "', '" & Trim(txtSec.Text) & "', N'" & Trim(TxtHelp3.Text) & "', " & CDbl(txtPrice.Text) & ", '" & Trim(cmbCurr.Text) & "', " & _
                            " " & CDbl(txtAmt.Text) & ", " & CDbl(txtRate.Text) & ", " & CDbl(txtKIP.Text) & ", " & CDbl(txtKIP.Text) & ", " & CDbl(txtLife.Text) & ", " & CDbl(txtYear.Text) & ", " & CDbl(txtMon.Text) & ", " & CDbl(txtDay.Text) & ", 0, 0, N'" & Trim(txtRemark.Text) & "', N'" & Trim(txtEngine.Text) & "', N'" & Trim(txtFrame.Text) & "', N'" & Trim(txtReg.Text) & "', N'" & Trim(txtSerial.Text) & "', " & IIf(chkDep.Checked = True, 1, 0) & "," & IIf(Chk_Make.Checked = True, 1, 0) & ",  N'" & Trim(CmdModelType.Text) & "', N'" & Trim(txtCertify.Text) & "', '" & Format(Date.Today, "yyyy-MM-dd") & "', '" & MUserID & "', '" & Trim(txtCompany.Text) & "', '" & Trim(txtDep.Text) & "', N'" & Trim(txtSec.Text) & "', " & CDbl(txt_disc.Text) & ", " & CDbl(txt_Rem.Text) & " )"
            CNN.Execute(PPP)
            '   CNN1.Execute("INSERT INTO gen_jn(certify, company, Book, amount,net_amt, date_work, code_dr,code_cr,ac_code, descrip,amount_dr,amount_cr,amt_dr,amt_cr,curr,Rate,last_update,last_user) " & _
            '  " VALUES('" & Trim(txtCertify.Text) & "','" & Trim(txtCompany.Text) & "','Fixd Asset', " & CDbl(txtPrice.Text) & "," & CDbl(txtKIP.Text) & ",'" & Format(DTBuy.Value, "yyyy-MM-dd") & "','" & txtAcc.Text & "','','" & txtAcc.Text & "',N'" & cmbGrp.Text & "'," & CDbl(txtPrice.Text) & ",0," & CDbl(txtKIP.Text) & ",0, " & _
            '  " '" & cmbCurr.Text & "'," & CDbl(txtRate.Text) & ",'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "')")
            '   CNN1.Execute("INSERT INTO gen_jn(certify,company, Book, amount, net_amt,date_work,code_dr,code_cr,ac_code, descrip,amount_dr,amount_cr,amt_dr,amt_cr,curr,Rate,last_update,last_user)  " & _
            '" VALUES('" & Trim(txtCertify.Text.ToString) & "','" & Trim(txtCompany.Text) & "','Fixd Asset', " & CDbl(txtPrice.Text) & ", " & CDbl(txtKIP.Text) & ",'" & Format(DTBuy.Value, "yyyy-MM-dd") & "','','" & txtGrp.Text & "','" & txtGrp.Text & "',N'" & cmbGrp.Text & "',0," & CDbl(txtPrice.Text) & ",0," & CDbl(txtKIP.Text) & ", " & _
            '   " '" & cmbCurr.Text & "'," & CDbl(txtRate.Text) & ",'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "')")
        End If
        '==================LACATION Name===================

        Dim rsp As New ADODB.Recordset
        Call LoadSqlData("SELECT * FROM AP_Project where Item_Nm=N'" & Trim(TxtHelp3.Text) & "' ", rsp)
        If rsp.RecordCount = 0 Then
            PJID()
            CNN.Execute("Insert into AP_Project(Item_ID,Item_Nm) values (N'" & Trim(TxtPJID.Text) & "',N'" & Trim(TxtHelp3.Text) & "') ")
        End If
        '===============================
        Dim rsk As New ADODB.Recordset
        Call LoadSqlData("SELECT * FROM Type where TypeNm=N'" & Trim(CmdModelType.Text) & "' ", rsk)
        If rsk.RecordCount = 0 Then
            TYID()
            CNN.Execute("Insert into Type(TypeID,TypeNm) values (N'" & Trim(TxtTYID.Text) & "',N'" & Trim(CmdModelType.Text) & "') ")
        End If
        '===========================
        Dim rsB As New ADODB.Recordset
        Call LoadSqlData("SELECT * FROM AP_Budget_Asset where Item_Nm=N'" & Trim(txtBudget.Text) & "' ", rsB)
        If rsB.RecordCount = 0 Then
            Budget()
            CNN.Execute("Insert into AP_Budget_Asset(Item_ID,Item_Nm) values (N'" & Trim(txtBudgetID.Text) & "',N'" & Trim(txtBudget.Text) & "') ")
        End If
        '======================Unit=====
        Dim rsU As New ADODB.Recordset
        Call LoadSqlData("SELECT * FROM AP_Unit where Item_Nm=N'" & Trim(txtUnit.Text) & "' ", rsU)
        If rsU.RecordCount = 0 Then
            Units()
            CNN.Execute("Insert into AP_Unit(Item_ID,Item_Nm) values (N'" & Trim(txtUnitID.Text) & "',N'" & Trim(txtUnit.Text) & "') ")
        End If
        txtID.Enabled = False
        MsgBox("Save Finish")
    End Sub
    Private Sub Units()
        Dim rRS As New ADODB.Recordset
        Dim mNum2 As String
        Call LoadSqlData("Select Top 1 Item_ID From AP_Unit Order by Item_ID DESC ", rRS)
        If rRS.RecordCount <> 0 Then
            mNum2 = Val(rRS.Fields("Item_ID").Value) + 1
            If Len(CStr(mNum2).Trim) = 1 Then
                txtUnitID.Text = "000" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 2 Then
                txtUnitID.Text = "00" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 3 Then
                txtUnitID.Text = "0" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 4 Then
                txtUnitID.Text = CStr(mNum2)
            End If
        Else
            txtUnitID.Text = "0001"
        End If
    End Sub
    Private Sub Budget()
        Dim rRS As New ADODB.Recordset
        Dim mNum2 As String
        Call LoadSqlData("Select Top 1 Item_ID From AP_Budget_Asset Order by Item_ID DESC ", rRS)
        If rRS.RecordCount <> 0 Then
            mNum2 = Val(rRS.Fields("Item_ID").Value) + 1
            If Len(CStr(mNum2).Trim) = 1 Then
                txtBudgetID.Text = "000" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 2 Then
                txtBudgetID.Text = "00" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 3 Then
                txtBudgetID.Text = "0" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 4 Then
                txtBudgetID.Text = CStr(mNum2)
            End If
        Else
            txtBudgetID.Text = "0001"
        End If
    End Sub
    Private Sub TYID()
        Dim rRS As New ADODB.Recordset
        Dim mNum2 As String
        Call LoadSqlData("Select Top 1 TypeID From Type Order by TypeID DESC ", rRS)
        If rRS.RecordCount <> 0 Then
            mNum2 = Val(rRS.Fields("TypeID").Value) + 1
            If Len(CStr(mNum2).Trim) = 1 Then
                TxtTYID.Text = "000" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 2 Then
                TxtTYID.Text = "00" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 3 Then
                TxtTYID.Text = "0" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 4 Then
                TxtTYID.Text = CStr(mNum2)
            End If
        Else
            TxtTYID.Text = "0001"
        End If
    End Sub
    Private Sub PJID()
        Dim rRS As New ADODB.Recordset
        Dim mNum2 As String
        Call LoadSqlData("Select Top 1 Item_ID From AP_Project Order by Item_ID DESC ", rRS)
        If rRS.RecordCount <> 0 Then
            mNum2 = Val(rRS.Fields("Item_ID").Value) + 1
            If Len(CStr(mNum2).Trim) = 1 Then
                TxtPJID.Text = "000" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 2 Then
                TxtPJID.Text = "00" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 3 Then
                TxtPJID.Text = "0" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 4 Then
                TxtPJID.Text = CStr(mNum2)
            End If
        Else
            TxtPJID.Text = "0001"
        End If
    End Sub
    Private Sub SaveACC()

    End Sub
    Private Sub RunID()
        Dim rRS As New ADODB.Recordset
        Dim mNum2 As String
        Call LoadSqlData("Select Top 1 AssetID, substring(AssetID,6,6) asNo From Assets Order by AssetID DESC ", rRS)
        If rRS.RecordCount <> 0 Then
            mNum2 = Val(rRS.Fields("asNo").Value) + 1
            If Len(CStr(mNum2).Trim) = 1 Then
                txtID.Text = "FA" & Format(Date.Today, "yy") & "." & "0000" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 2 Then
                txtID.Text = "FA" & Format(Date.Today, "yy") & "." & "000" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 3 Then
                txtID.Text = "FA" & Format(Date.Today, "yy") & "." & "00" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 4 Then
                txtID.Text = "FA" & Format(Date.Today, "yy") & "." & "0" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 5 Then
                txtID.Text = CStr(mNum2)
            End If
        Else
            txtID.Text = "FA" & Format(Date.Today, "yy") & "." & "00001"
        End If
    End Sub
    Private Sub LdGrp()
        Dim gRS As New ADODB.Recordset
        cmbGrp.Items.Clear()
        If Lang = True Then
            Call LoadSqlData("Select * from Groups_Asset Order by Group_ID", gRS)
            If gRS.RecordCount <> 0 Then
                While Not gRS.EOF
                    cmbGrp.Items.Add(gRS.Fields("Group_NmE").Value.ToString)
                    gRS.MoveNext()
                End While
            End If
            cmbGrp.SelectedIndex = 0
        Else
            Call LoadSqlData("Select * from Groups_Asset Order by Group_ID", gRS)
            If gRS.RecordCount <> 0 Then
                While Not gRS.EOF
                    cmbGrp.Items.Add(gRS.Fields("Group_Nm").Value.ToString)
                    gRS.MoveNext()
                End While
            End If
            cmbGrp.SelectedIndex = 0
        End If

    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click

        Dim Rs, Rs1 As New ADODB.Recordset
        Dim rpt As New CryAssetItem
        Dim FrmPreview As New FmPreview : FrmClosing()
        Call LoadSqlData("Select * from Assets Where AssetID='" & Trim(txtID.Text) & "' and Company='" & txtCompany.Text & "' ", Rs)
        COM = txtCompany.Text
        Call Office()
        With rpt
            Dim myTextObjectOnReport As CrystalDecisions.CrystalReports.Engine.TextObject
            myTextObjectOnReport = CType(rpt.ReportDefinition.ReportObjects.Item("SG1"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myTextObjectOnReport.Text = Sign5
            myTextObjectOnReport = CType(rpt.ReportDefinition.ReportObjects.Item("SG2"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myTextObjectOnReport.Text = Sign3
            myTextObjectOnReport = CType(rpt.ReportDefinition.ReportObjects.Item("SG3"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myTextObjectOnReport.Text = Sign1
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
            '    FmPreview.CrystalReportViewer1.ReportSource = SubDoc
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
    Private Sub LdDep()
        Dim sRS As New ADODB.Recordset
        cmbDeprt.Items.Clear()
        Call LoadSqlData("Select * from Department Order by DepartmentID", sRS)
        If sRS.RecordCount <> 0 Then
            While Not sRS.EOF
                cmbDeprt.Items.Add(sRS.Fields("DepartmentNm").Value.ToString)
                sRS.MoveNext()
            End While
        End If
        If cmbDeprt.Items.Count > 0 Then cmbDeprt.SelectedIndex = 0
    End Sub
    Private Sub LdSec()
        Dim sRS As New ADODB.Recordset
        cmbSec.Items.Clear()
        If Lang = True Then
            Call LoadSqlData("Select * from AP_Office where Off_ID<>'00' Order by Off_ID", sRS)
            If sRS.RecordCount <> 0 Then
                While Not sRS.EOF
                    cmbSec.Items.Add(sRS.Fields("Off_Name").Value.ToString)
                    sRS.MoveNext()
                End While
            End If
            If cmbSec.Items.Count > 0 Then cmbSec.SelectedIndex = 0
        Else
            If Mpermiss = "Admin" Then
                'cmbSec.Items.Add("** ສະແດງທັງໝົດ **")
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

    Private Sub FrmAssetNew_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Call Loadlang()

        'Call LoadLng()
        'Call LoadLng2()

        Call LdGrp()
        'Call LdCompany()
        'txtCompany.Text = mCompany
        txtCompany.Text = Off_Id
        Call LdSec()
        'Call LdDep()
        TxtHelp3.Items.Clear()
        Call LoadSqlData("Select * from AP_Project Order by Item_ID", RSC)
        If RSC.RecordCount <> 0 Then
            While Not RSC.EOF
                TxtHelp3.Items.Add(RSC.Fields("Item_Nm").Value.ToString)
                RSC.MoveNext()
            End While
        End If
        '========================
        CmdModelType.Items.Clear()
        Call LoadSqlData("Select * from Type  Order by TypeID", RSC)
        If RSC.RecordCount <> 0 Then
            While Not RSC.EOF
                CmdModelType.Items.Add(RSC.Fields("TypeNm").Value.ToString)
                RSC.MoveNext()
            End While
        End If
        '===================================
        txtBudget.Items.Clear()
        Call LoadSqlData("Select * from AP_Budget_Asset  Order by Item_ID", RSC)
        If RSC.RecordCount <> 0 Then
            While Not RSC.EOF
                txtBudget.Items.Add(RSC.Fields("Item_Nm").Value.ToString)
                RSC.MoveNext()
            End While
        End If
        '==================
        txtUnit.Items.Clear()
        Call LoadSqlData("Select * from AP_Unit  Order by Item_ID", RSC)
        If RSC.RecordCount <> 0 Then
            While Not RSC.EOF
                txtUnit.Items.Add(RSC.Fields("Item_Nm").Value.ToString)
                RSC.MoveNext()
            End While
        End If
        If mEdit = True Then
            Call loadAST()
            If Barcode.Text = "" Then
                Button3.Enabled = True
                'Barcode.ReadOnly = False
            Else
                Button3.Enabled = False
                'Barcode.ReadOnly = True
            End If
        Else
            Call cmdNew_Click(sender, e)
        End If
        'If mCompStr = "" Then cmdNew.Enabled = False
        Label30.Text = "ເນື້ອທີ"
        Label31.Text = "/ ມ2"
        Label32.Text = "Barcode"
        Button3.Text = "........"
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
    Private Sub loadAST()
        CNN.Execute("UPDATE Assets set Make_Buy=0 where Make_Buy is null ")


        Dim aRS As New ADODB.Recordset 
        Dim hh As String = "SELECT   dbo.Assets.* , dbo.AP_Office.Off_Name  FROM dbo.Assets INNER JOIN dbo.AP_Office ON dbo.Assets.Sect_ID = dbo.AP_Office.off_id   where Assets.assetID='" & myTemp & "' "
        Call LoadSqlData(hh, aRS)

        If aRS.RecordCount <> 0 Then
            txtID.Text = Trim(aRS.Fields("assetID").Value.ToString)
            txtCode.Text = Trim(aRS.Fields("Asset_No").Value.ToString)
            txtNm.Text = Trim(aRS.Fields("Asset_Nm").Value.ToString)
            txtNmE.Text = Trim(aRS.Fields("Asset_NmE").Value.ToString)
            TxtVendor.Text = Trim(aRS.Fields("Vendor").Value.ToString)
            txtBudget.Text = Trim(aRS.Fields("Budget").Value.ToString)
            txtCompany.Text = Trim(aRS.Fields("Company").Value.ToString)
            txtAcc.Text = Trim(aRS.Fields("Fix_Acc").Value.ToString)
            txtCertify.Text = Trim(aRS.Fields("Certify").Value.ToString)
            'cmbGrp.Text = Trim(aRS.Fields("Group_Nm").Value.ToString)
            txtGrp.Text = Trim(aRS.Fields("Group_ID").Value.ToString)
            txtEngine.Text = Trim(aRS.Fields("Engin_No").Value.ToString)
            txtFrame.Text = Trim(aRS.Fields("Frame_No").Value.ToString)
            txtSerial.Text = Trim(aRS.Fields("Serial").Value.ToString)
            txtReg.Text = Trim(aRS.Fields("Regist_No").Value.ToString)
            txtUnit.Text = Trim(aRS.Fields("Unit").Value.ToString)
            CmdModelType.Text = Trim(aRS.Fields("Model").Value.ToString)
            chkDep.Checked = IIf(aRS.Fields("DepCalc").Value = 1, True, False)
            Chk_Make.Checked = IIf(aRS.Fields("Make_Buy").Value = 1, True, False)
            txtLife.Text = Trim(aRS.Fields("Used_Life").Value.ToString)
            txtPrice.Text = Format(aRS.Fields("Price").Value, "#,##0.00")
            txtAmt.Text = Format(aRS.Fields("Amount").Value, "#,##0.00")
            txtRate.Text = Format(aRS.Fields("Rate").Value, "#,##0.00")
            txtKIP.Text = Format(aRS.Fields("Amt_KIP").Value, "#,##0.00")
            'txtYear.Text = Format(aRS.Fields("Dep_Year").Value, "#,##0.00")
            'txtMon.Text = Format(aRS.Fields("Dep_Month").Value, "#,##0.00")
            txtDay.Text = Format(aRS.Fields("Dep_Day").Value, "#,##0.00")
            txtQty.Text = Format(aRS.Fields("Qty").Value, "#,##0")
            txtPost.Text = Trim(aRS.Fields("Deposted_Date").Value.ToString)
            txtRemark.Text = Trim(aRS.Fields("Remark").Value.ToString)
            DTBuy.Value = Format(aRS.Fields("Date_Buy").Value, "dd/MM/yyyy")
            DTUse.Value = Format(aRS.Fields("Date_Work").Value, "dd/MM/yyyy")
            cmbCurr.Text = Trim(aRS.Fields("Curr").Value.ToString)
            txtSec.Text = Trim(aRS.Fields("sect_ID").Value.ToString)
            cmbSec.Text = Trim(aRS.Fields("Off_Name").Value.ToString)
            txtDep.Text = Trim(aRS.Fields("DepartmentID").Value.ToString)
            cmbDeprt.Text = Trim(aRS.Fields("DepartmentNm").Value.ToString)
            TxtHelp3.Text = Trim(aRS.Fields("Using_By").Value.ToString)
            Barcode.Text = Trim(aRS.Fields("Bar_code").Value.ToString)

            txt_disc.Text = Format(CDbl(aRS.Fields("txt_disc").Value), "#,##0.00")
            txt_Rem.Text = Format(CDbl(aRS.Fields("txt_Rem").Value), "#,##0.00")

            If txt_disc.Text <> "0" Then
                CheckBox1.Checked = True
            Else
                CheckBox1.Checked = False
            End If
        End If
        Call CALLC()
        Call LoadSqlData("SELECT * FROM Groups_Asset where Group_ID='" & txtGrp.Text & "' ", RSC)
        If RSC.RecordCount <> 0 Then
            cmbGrp.Text = Trim(RSC.Fields("Group_Nm").Value.ToString)
        End If
        txtID.Enabled = False


    End Sub

    Private Sub CALLC() 
        Dim yearday, days, dayOfmonth As Integer
        Dim date1, date2 As Date

        dayOfmonth = System.DateTime.DaysInMonth(Year(MWorkSetting), Month(MWorkSetting))
 
        Dim DT_date As Date
   
        DT_date = DateAdd(DateInterval.Year, CDbl(txtLife.Text), CDate(DTUse.Value))

        days = DateDiff(DateInterval.Day, CDate(DTUse.Value), CDate(DT_date))

        dayOfmonth = System.DateTime.DaysInMonth(Year(MWorkSetting), Month(MWorkSetting))

        date1 = CDate(Year(MWorkSetting) & "/01/01")
        date2 = CDate(Year(MWorkSetting) + 1 & "/01/01")
        yearday = DateDiff(DateInterval.Day, date1, date2)

        'txtMon.Text = Format(CDbl(dayOfmonth) * CDbl(txtDay.Text), "#,##0.00")
        'txtYear.Text = Format(CDbl(yearday) * CDbl(txtDay.Text), "#,##0.00")

        txtYear.Text = Format(CDbl(txtKIP.Text) / CDbl(txtLife.Text), "#,#0.00")
        txtMon.Text = Format(CDbl(txtKIP.Text) / (CDbl(txtLife.Text) * 12), "#,#0.00")
        'txtDay.Text = Format(CDbl(txtKIP.Text) / 365, "#,#0.00")
        txtDay.Text = Format(CDbl(txtKIP.Text) / CDbl(days), "#,#0.00")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub cmbGrp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbGrp.Click
        Call cmbGrp_SelectedIndexChanged(sender, e)
    End Sub

    Private Sub cmbGrp_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbGrp.SelectedIndexChanged
        Dim gRS As New ADODB.Recordset
        If Lang = True Then
            Call LoadSqlData("select * from Groups_Asset Where Group_NmE=N'" & Trim(cmbGrp.Text) & "'", gRS)
            If gRS.RecordCount <> 0 Then
                txtGrp.Text = Trim(gRS.Fields("Group_ID").Value.ToString)
                txtAcc.Text = Trim(gRS.Fields("AccountCodeAsCR").Value.ToString)
                If txtCode.Text = "" Then
                    txtCode.Text = txtAcc.Text & "."
                End If
            End If
        Else
            Call LoadSqlData("select * from Groups_Asset Where Group_Nm=N'" & Trim(cmbGrp.Text) & "' ", gRS)
            If gRS.RecordCount <> 0 Then
                txtGrp.Text = Trim(gRS.Fields("Group_ID").Value.ToString)
                txtAcc.Text = Trim(gRS.Fields("AccountCodeAsCR").Value.ToString)
                If txtCode.Text = "" Then
                    txtCode.Text = txtAcc.Text & "."
                End If
            End If
        End If
    End Sub

    Private Sub cmdNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNew.Click
        txtBudget.Text = ""
        txtNmE.Text = ""
        TxtVendor.Text = ""
        txtCode.Text = ""
        TxtHelp3.Text = ""
        txtNm.Text = ""
        'txtCompany.Text = mCompany
        txtAcc.Text = ""
        txtCertify.Text = ""
        txtGrp.Text = ""
        cmbGrp.SelectedIndex = 0
        txtEngine.Text = ""
        txtFrame.Text = ""
        txtSerial.Text = ""
        txtReg.Text = ""
        TxtArea.Text = ""
        txtLife.Text = 1
        txtUnit.Text = ""
        chkDep.Checked = True
        Chk_Make.Checked = True
        txtPrice.Text = 0
        txtAmt.Text = 0
        txtRate.Text = 1
        txtKIP.Text = 0
        txtYear.Text = 0
        txtMon.Text = 0
        txtDay.Text = 0
        txt_disc.Text = 0
        txt_Rem.Text = 0
        txtPost.Text = ""
        txtRemark.Text = ""
        DTBuy.Value = Format(mDate, "dd/MM/yyyy")
        DTUse.Value = Format(mDate, "dd/MM/yyyy")
        cmbCurr.SelectedIndex = 0
        txtCode.Enabled = True
        Call RunID()
        Call RunBarcode()
        txtID.Enabled = True
        Call cmbGrp_SelectedIndexChanged(sender, e)
        txtCode.Focus()
    End Sub

    Private Sub cmbSec_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSec.SelectedIndexChanged
        Dim sRS As New ADODB.Recordset
        If Lang = True Then
            Call LoadSqlData("select * from AP_Office Where Off_NameE=N'" & Trim(cmbSec.Text) & "'", sRS)
            If sRS.RecordCount <> 0 Then
                txtSec.Text = Trim(sRS.Fields("Off_ID").Value.ToString)
                txtCompany.Text = txtSec.Text
                Dim dRS As New ADODB.Recordset
                cmbDeprt.Items.Clear()
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
            'If cmbSec.SelectedIndex = 0 Then
            '    txtCompany.Text = "01-VTE"
            'Else
            '    txtCompany.Text = ""
            'End If
        Else
            Call LoadSqlData("select * from AP_Office Where Off_Name=N'" & Trim(cmbSec.Text) & "'", sRS)
            If sRS.RecordCount <> 0 Then
                txtSec.Text = Trim(sRS.Fields("Off_ID").Value.ToString)
                Dim dRS As New ADODB.Recordset
                cmbDeprt.Items.Clear()
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
            'If cmbSec.SelectedIndex = 0 Then
            '    txtCompany.Text = "01-VTE"
            'Else
            '    txtCompany.Text = ""
            'End If
        End If
    End Sub

    Private Sub cmbDeprt_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbDeprt.SelectedIndexChanged
        Dim dRS As New ADODB.Recordset
        If Lang = True Then
            Call LoadSqlData("Select * from Department Where DepartmentNmE= N'" & Trim(cmbDeprt.Text) & "' ", dRS)
            If dRS.RecordCount <> 0 Then
                txtDep.Text = Trim(dRS.Fields("DepartmentID").Value.ToString)
                txtCompany.Text = Trim(dRS.Fields("Company").Value.ToString)
            End If
        Else
            Call LoadSqlData("Select * from Department Where DepartmentNm = N'" & Trim(cmbDeprt.Text) & "' ", dRS)
            If dRS.RecordCount <> 0 Then
                txtDep.Text = Trim(dRS.Fields("DepartmentID").Value.ToString)
                txtCompany.Text = Trim(dRS.Fields("Company").Value.ToString)
            End If
        End If


    End Sub

    Private Sub txtLife_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtLife.KeyPress
        'If txtLife.Text = 0 Then txtLife.Text = 1
        'txtYear.Text = Format(CDbl(txtKIP.Text) / CDbl(txtLife.Text), "#,#0.00")
        'txtMon.Text = Format(CDbl(txtYear.Text) / 12, "#,#0.00")
        ''txtDay.Text = Format(CDbl(txtMon.Text) / 30, "#,#0.00")
        'txtDay.Text = Format(CDbl(txtYear.Text) / 365, "#,#0.00")
        'txtRemark.Focus()
        If e.KeyChar = Chr(13) Then
            If txtLife.Text = 0 Then txtLife.Text = 1
            Dim yearday, days, dayOfmonth As Integer
            Dim DT_date As Date
            Dim date1, date2 As Date

            DT_date = DateAdd(DateInterval.Year, CDbl(txtLife.Text), CDate(DTUse.Value))

            days = DateDiff(DateInterval.Day, CDate(DTUse.Value), CDate(DT_date))

            dayOfmonth = System.DateTime.DaysInMonth(Year(MWorkSetting), Month(MWorkSetting))

            date1 = CDate(Year(MWorkSetting) & "/01/01")
            date2 = CDate(Year(MWorkSetting) + 1 & "/01/01")
            yearday = DateDiff(DateInterval.Day, date1, date2)
             
            txtAmt.Text = Format(CDbl(txtPrice.Text) * CDbl(txtQty.Text), "#,#0.00")
            'txt_disc.Text = Format(CDbl(txtAmt.Text) * 0.01, "#,#0.00")
            If CheckBox1.Checked = True Then
                txt_disc.Text = Format(CDbl(txtAmt.Text) * 0.01, "#,#0.00")
            Else
                txt_disc.Text = Format(CDbl(0), "#,#0.00")
            End If
            txt_Rem.Text = Format(CDbl(txtAmt.Text) - CDbl(txt_disc.Text), "#,#0.00")
            txtKIP.Text = Format(CDbl(txt_Rem.Text) * CDbl(txtRate.Text), "#,#0.00")

            'txtDay.Text = Format(CDbl(txtKIP.Text) / CDbl(days), "#,#0.00")
            'txtMon.Text = Format(CDbl(dayOfmonth) * CDbl(txtDay.Text), "#,##0.00")
            'txtYear.Text = Format(CDbl(yearday) * CDbl(txtDay.Text), "#,##0.00")

            txtYear.Text = Format(CDbl(txtKIP.Text) / CDbl(txtLife.Text), "#,#0.00")
            txtMon.Text = Format(CDbl(txtKIP.Text) / (CDbl(txtLife.Text) * 12), "#,#0.00")
            'txtDay.Text = Format(CDbl(txtKIP.Text) / 365, "#,#0.00")
            txtDay.Text = Format(CDbl(txtKIP.Text) / CDbl(days), "#,#0.00")
            txtRemark.Focus()


        End If
    End Sub

    Private Sub txtPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPrice.KeyPress

        'txtAmt.Text = Format(CDbl(txtPrice.Text) * CDbl(txtQty.Text), "#,#0.00")
        'txtKIP.Text = Format(CDbl(txtAmt.Text) * CDbl(txtRate.Text), "#,#0.00")
        'txtYear.Text = Format(CDbl(txtKIP.Text) / CDbl(txtLife.Text), "#,#0.00")
        'txtMon.Text = Format(CDbl(txtYear.Text) / 12, "#,#0.00")
        'txtDay.Text = Format(CDbl(txtMon.Text) / 30, "#,#0.00")

        If e.KeyChar = Chr(13) Then

            Dim yearday, days, dayOfmonth As Integer
            Dim DT_date As Date
            Dim date1, date2 As Date

            DT_date = DateAdd(DateInterval.Year, CDbl(txtLife.Text), CDate(DTUse.Value))

            days = DateDiff(DateInterval.Day, CDate(DTUse.Value), CDate(DT_date))

            dayOfmonth = System.DateTime.DaysInMonth(Year(MWorkSetting), Month(MWorkSetting))

            date1 = CDate(Year(MWorkSetting) & "/01/01")
            date2 = CDate(Year(MWorkSetting) + 1 & "/01/01") 
            yearday = DateDiff(DateInterval.Day, date1, date2)

             
            If txtLife.Text = 0 Then txtLife.Text = 1
            txtAmt.Text = Format(CDbl(txtPrice.Text) * CDbl(txtQty.Text), "#,#0.00")
            If CheckBox1.Checked = True Then
                txt_disc.Text = Format(CDbl(CDbl(txtAmt.Text) * CDbl(txtRate.Text)) * 0.01, "#,#0.00")
            Else
                txt_disc.Text = Format(CDbl(0), "#,#0.00")
            End If

            txt_Rem.Text = Format(CDbl(CDbl(txtAmt.Text) * CDbl(txtRate.Text)) - CDbl(txt_disc.Text), "#,#0.00")
            txtKIP.Text = Format(CDbl(txt_Rem.Text) * CDbl(txtRate.Text), "#,#0.00")

            'txtDay.Text = Format(CDbl(txtKIP.Text) / CDbl(days), "#,#0.00")
            'txtMon.Text = Format(CDbl(dayOfmonth) * CDbl(txtDay.Text), "#,##0.00")
            'txtYear.Text = Format(CDbl(yearday) * CDbl(txtDay.Text), "#,##0.00")

            txtYear.Text = Format(CDbl(txtKIP.Text) / CDbl(txtLife.Text), "#,#0.00")
            txtMon.Text = Format(CDbl(txtKIP.Text) / (CDbl(txtLife.Text) * 12), "#,#0.00")
            txtDay.Text = Format(CDbl(txtKIP.Text) / CDbl(days), "#,#0.00")

            txtLife.Focus()
        End If

    End Sub

    Private Sub txtQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtQty.KeyPress
        If e.KeyChar = Chr(13) Then
            If Val(txtQty.Text) = 0 Then txtQty.Text = 1
            If txtLife.Text = 0 Then txtLife.Text = 1
            'txtAmt.Text = Format(CDbl(txtPrice.Text) * CDbl(txtQty.Text), "#,#0.00")
            'txtKIP.Text = Format(CDbl(txtAmt.Text) * CDbl(txtRate.Text), "#,#0.00")
            'txtYear.Text = Format(CDbl(txtKIP.Text) / CDbl(txtLife.Text), "#,#0.00")
            'txtMon.Text = Format(CDbl(txtYear.Text) / 12, "#,#0.00")
            'txtDay.Text = Format(CDbl(txtMon.Text) / 30, "#,#0.00")
            Dim yearday, days, dayOfmonth As Integer
            Dim DT_date As Date
            Dim date1, date2 As Date

            DT_date = DateAdd(DateInterval.Year, CDbl(txtLife.Text), CDate(DTUse.Value))

            days = DateDiff(DateInterval.Day, CDate(DTUse.Value), CDate(DT_date))

            dayOfmonth = System.DateTime.DaysInMonth(Year(MWorkSetting), Month(MWorkSetting))

            date1 = CDate(Year(MWorkSetting) & "/01/01")
            date2 = CDate(Year(MWorkSetting) + 1 & "/01/01")
            yearday = DateDiff(DateInterval.Day, date1, date2)

             
            txtAmt.Text = Format(CDbl(txtPrice.Text) * CDbl(txtQty.Text), "#,#0.00")
            'txtKIP.Text = Format(CDbl(txtAmt.Text) * CDbl(txtRate.Text), "#,#0.00")
            ' txt_disc.Text = Format(CDbl(txtAmt.Text) * 0.01, "#,#0.00")
            If CheckBox1.Checked = True Then
                txt_disc.Text = Format(CDbl(txtAmt.Text) * 0.01, "#,#0.00")
            Else
                txt_disc.Text = Format(CDbl(0), "#,#0.00")
            End If
            txt_Rem.Text = Format(CDbl(txtAmt.Text) - CDbl(txt_disc.Text), "#,#0.00")
            txtKIP.Text = Format(CDbl(txt_Rem.Text) * CDbl(txtRate.Text), "#,#0.00")


            'txtDay.Text = Format(CDbl(txtKIP.Text) / CDbl(days), "#,#0.00")
            'txtMon.Text = Format(CDbl(dayOfmonth) * CDbl(txtDay.Text), "#,##0.00")
            'txtYear.Text = Format(CDbl(yearday) * CDbl(txtDay.Text), "#,##0.00")

            txtYear.Text = Format(CDbl(txtKIP.Text) / CDbl(txtLife.Text), "#,#0.00")
            txtMon.Text = Format(CDbl(txtKIP.Text) / (CDbl(txtLife.Text) * 12), "#,#0.00")
            'txtDay.Text = Format(CDbl(txtKIP.Text) / 365, "#,#0.00")
            txtDay.Text = Format(CDbl(txtKIP.Text) / CDbl(days), "#,#0.00")
        End If
    End Sub

    Private Sub txtPrice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrice.LostFocus
        txtPrice.Text = Format(CDbl(txtPrice.Text), "#,#0.00")
    End Sub

    Private Sub txtPrice_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPrice.TextChanged

    End Sub

    Private Sub txtQty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtQty.TextChanged

    End Sub

    Private Sub txtMon_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMon.TextChanged

    End Sub

    Private Sub txtRate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRate.KeyPress
        If e.KeyChar = Chr(13) Then
            'txtKIP.Text = Format(CDbl(txtAmt.Text) * CDbl(txtRate.Text), "#,#0.00")
            'txtYear.Text = Format(CDbl(txtKIP.Text) / CDbl(txtLife.Text), "#,#0.00")
            'txtMon.Text = Format(CDbl(txtYear.Text) / 12, "#,#0.00")
            'txtDay.Text = Format(CDbl(txtMon.Text) / 30, "#,#0.00")
            Dim yearday, days, dayOfmonth As Integer
            Dim DT_date As Date
            Dim date1, date2 As Date

            DT_date = DateAdd(DateInterval.Year, CDbl(txtLife.Text), CDate(DTUse.Value))

            days = DateDiff(DateInterval.Day, CDate(DTUse.Value), CDate(DT_date))

            dayOfmonth = System.DateTime.DaysInMonth(Year(MWorkSetting), Month(MWorkSetting))

            date1 = CDate(Year(MWorkSetting) & "/01/01")
            date2 = CDate(Year(MWorkSetting) + 1 & "/01/01")
            yearday = DateDiff(DateInterval.Day, date1, date2)


            If txtLife.Text = 0 Then txtLife.Text = 1
            txtAmt.Text = Format(CDbl(txtPrice.Text) * CDbl(txtQty.Text), "#,#0.00")
            txtKIP.Text = Format(CDbl(txtAmt.Text) * CDbl(txtRate.Text), "#,#0.00")

            txtDay.Text = Format(CDbl(txtKIP.Text) / CDbl(days), "#,#0.00")
            txtMon.Text = Format(CDbl(dayOfmonth) * CDbl(txtDay.Text), "#,##0.00")
            txtYear.Text = Format(CDbl(yearday) * CDbl(txtDay.Text), "#,##0.00")
        End If
    End Sub

    Private Sub txtRate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRate.TextChanged

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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim SLl As New ADODB.Recordset
        'If txtCertify.Text = "" Then MsgBox("ກະລຸນາໃສ່ເລກບິນກ່ອນ", MsgBoxStyle.Question) : Exit Sub
        'Call LoadSqlData("Select AssetID From Assets Where AssetID='" & Trim(txtID.Text) & "'", cRS)
        If txtAcc.Text = "" Then MsgBox("ກະລຸນາໃສ່ເລກບັນຊີກ່ອນ", MsgBoxStyle.Exclamation) : txtAcc.Focus() : Exit Sub
        If txtGrp.Text = "" Then MsgBox("ກະລຸນາໃສ່ເລກບັນຊີກ່ອນ", MsgBoxStyle.Exclamation) : txtGrp.Focus() : Exit Sub

        Dim gRS As New ADODB.Recordset
        Call LoadSqlData("select * from Acc_Code Where Ac_Code=N'" & Trim(txtAcc.Text) & "'", gRS)
        If gRS.RecordCount = 0 Then
            MsgBox("ເລກບັນຊີບໍ່ມີໃນສາລະບານ " & Trim(txtAcc.Text), MsgBoxStyle.Exclamation) : txtAcc.Focus() : Exit Sub
        End If

        'Call LoadSqlData("select * from Acc_Code Where Ac_Code=N'" & Trim(txtGrp.Text) & "'", gRS)
        'If gRS.RecordCount = 0 Then
        '    MsgBox("ເລກບັນຊີບໍ່ມີໃນສາລະບານ " & Trim(txtGrp.Text), MsgBoxStyle.Exclamation) : txtGrp.Focus() : Exit Sub
        'End If

        CNN.Execute("DELETE FROM gen_jn where Book=N'Fixd Asset' and certify=N'" & Trim(txtID.Text) & "' ")
        CNN.Execute("INSERT INTO gen_jn(certify,company,Office_ID, Book,amount, net_amt, date_work, code_dr,code_cr,ac_code, descrip,amount_dr,amount_cr,amt_dr,amt_cr,curr,Rate,last_update,last_user,Frm,lock) " & _
           " VALUES('" & Trim(txtID.Text) & "','" & Trim(MuSubOff) & "','" & Trim(MuSubOff) & "','Fixd Asset', " & CDbl(txtPrice.Text) & "," & CDbl(txtKIP.Text) & ",'" & Format(DTBuy.Value, "yyyy-MM-dd") & "','" & txtGrp.Text & "','','" & txtGrp.Text & "',N'" & cmbGrp.Text & "'," & CDbl(txtPrice.Text) & ",0," & CDbl(txtKIP.Text) & ",0, " & _
           " '" & cmbCurr.Text & "'," & CDbl(txtRate.Text) & ",'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "',1,4)")
        CNN.Execute("INSERT INTO gen_jn(certify,company, Office_ID,Book,amount,net_amt,date_work,code_dr,code_cr,ac_code, descrip,amount_dr,amount_cr,amt_dr,amt_cr,curr,Rate,last_update,last_user,Frm,lock)  " & _
     " VALUES('" & Trim(txtID.Text.ToString) & "','" & Trim(MuSubOff) & "','" & Trim(MuSubOff) & "','Fixd Asset', " & CDbl(txtPrice.Text) & "," & CDbl(txtKIP.Text) & ",'" & Format(DTBuy.Value, "yyyy-MM-dd") & "','','" & txtAcc.Text & "','" & txtAcc.Text & "',N'" & cmbGrp.Text & "',0," & CDbl(txtPrice.Text) & ",0," & CDbl(txtKIP.Text) & ", " & _
        " '" & cmbCurr.Text & "'," & CDbl(txtRate.Text) & ",'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "',1,4)")
        CNN.Execute("update Gen_jn set Gen_jn.descrip=Acc_Code.Name_L, Gen_jn.ac_name=Acc_Code.Name_L  from Acc_Code,Gen_jn where Gen_jn.certify='" & Trim(txtID.Text) & "' and Gen_jn.AC_Code=ACC_Code.AC_Code ")
        MsgBox("Transfering Finish")
    End Sub

    Private Sub TxtVendor_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtVendor.KeyPress
        If e.KeyChar = Chr(13) Then
            txtNm.Focus()
        End If
    End Sub

    Private Sub TxtVendor_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtVendor.TextChanged

    End Sub

    Private Sub txtNm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNm.KeyPress
        If e.KeyChar = Chr(13) Then
            txtNmE.Focus()
        End If
    End Sub

    Private Sub txtNm_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNm.TextChanged

    End Sub

    Private Sub txtNmE_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNmE.KeyPress
        If e.KeyChar = Chr(13) Then
            CmdModelType.Focus()
        End If
    End Sub

    Private Sub txtNmE_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNmE.TextChanged

    End Sub

    Private Sub txtModel_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = Chr(13) Then
            txtEngine.Focus()
        End If
    End Sub

    Private Sub txtModel_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtEngine_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEngine.KeyPress
        If e.KeyChar = Chr(13) Then
            TxtHelp3.Focus()
        End If
    End Sub

    Private Sub txtEngine_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEngine.TextChanged

    End Sub

    Private Sub txtLocate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = Chr(13) Then
            txtBudget.Focus()
        End If
    End Sub

    Private Sub txtLocate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtBudget_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = Chr(13) Then
            txtSerial.Focus()
        End If
    End Sub

    Private Sub txtBudget_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtSerial_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSerial.KeyPress
        If e.KeyChar = Chr(13) Then
            txtFrame.Focus()
        End If
    End Sub

    Private Sub txtSerial_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSerial.TextChanged

    End Sub

    Private Sub txtFrame_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFrame.KeyPress
        If e.KeyChar = Chr(13) Then
            txtReg.Focus()
        End If
    End Sub

    Private Sub txtFrame_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFrame.TextChanged

    End Sub

    Private Sub txtReg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtReg.KeyPress
        If e.KeyChar = Chr(13) Then
            txtPrice.Focus()
        End If
    End Sub

    Private Sub txtReg_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtReg.TextChanged

    End Sub

    Private Sub txtLife_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLife.TextChanged

    End Sub

    Private Sub txtRemark_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRemark.KeyPress
        If e.KeyChar = Chr(13) Then
            cmdSave_Click(sender, e)
        End If
    End Sub

    Private Sub txtRemark_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRemark.TextChanged

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub

    Private Sub TxtHelp3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtHelp3.SelectedIndexChanged

    End Sub

    Private Sub chkDep_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDep.CheckedChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call RunBarcode()
    End Sub
    Private Sub RunBarcode()
        Dim tempEven As Long
        Dim tempOdd As Long
        Dim tempTotal
        Dim newBarCode As String
        Dim Checksum As Integer
        Dim Pos As Integer
        Randomize()
        newBarCode = Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10)
        For Pos = 2 To 12 Step 2
            tempEven = tempEven + Val(Mid(newBarCode, Pos, 1))
        Next
        For Pos = 1 To 11 Step 2
            tempOdd = tempOdd + Val(Mid(newBarCode, Pos, 1))
        Next
        tempEven = tempEven * 3
        tempTotal = tempOdd + tempEven
        Checksum = tempTotal Mod 10
        If Checksum > 0 Then
            Checksum = 10 - Checksum
        End If
        newBarCode = newBarCode & Checksum
        If Checksum <> Mid(newBarCode, 13, 1) Then
            MsgBox("ì½¹ñ©®¾Â£©êóúêú¾­¦ñú¤¯½ªò®ñ©®Ò¦¾´¾©­¿Ã§ûÄ©û, ¡½ì÷­¾¦ñú¤ÃÏú!", MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        Me.Barcode.Text = newBarCode
        If (newBarCode) = "0" Then Me.Barcode.Text = "" : RunBarcode()
        Dim Rschk As New ADODB.Recordset
        With Rschk
            Call LoadSqlData("Select Top 1  Bar_Code From Assets WHERE Bar_Code='" & newBarCode & "'", Rschk)
            If .RecordCount <> 0 Then Barcode.Text = "" : RunBarcode()
        End With
    End Sub

    Private Sub txtAmt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAmt.KeyPress

    End Sub

    Private Sub txtAmt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmt.TextChanged

    End Sub

    Private Sub cmbCurr_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCurr.SelectedIndexChanged
        MDR_Curr = " and Curr='" & cmbCurr.Text & "' "
        Call RateSetting() 
        If cmbCurr.SelectedIndex = 0 Then
            txtRate.Text = Format(MDLAK, "#,##0.00")
        ElseIf cmbCurr.SelectedIndex = 1 Then
            txtRate.Text = Format(MDTHB_LAK, "#,##0.00")
        Else
            txtRate.Text = Format(MDUSD_LAK, "#,##0.00")
        End If

        If txtAmt.Text = "" Then txtAmt.Text = 0

        'txtKIP.Text = Format(CDbl(txtAmt.Text) * CDbl(txtRate.Text), "#,#0.00")
        'txtYear.Text = Format(CDbl(txtKIP.Text) / CDbl(txtLife.Text), "#,#0.00")
        'txtMon.Text = Format(CDbl(txtYear.Text) / 12, "#,#0.00")
        'txtDay.Text = Format(CDbl(txtMon.Text) / 30, "#,#0.00")

        Dim yearday, days As Integer
        Dim DT_date As Date

        'DT_date = DateAdd(DateInterval.Year, CDbl(txtLife.Text), CDate(DTUse.Value))

        'days = DateDiff(DateInterval.Day, CDate(DTUse.Value), CDate(DT_date))

        'yearday = System.DateTime.DaysInMonth(Year(DTUse.Value), Month(DTUse.Value))

        'txtAmt.Text = Format(CDbl(txtPrice.Text) * CDbl(txtQty.Text), "#,#0.00")
        'txtKIP.Text = Format(CDbl(txtAmt.Text) * CDbl(txtRate.Text), "#,#0.00")
        'txtDay.Text = Format(CDbl(txtKIP.Text) / CDbl(days), "#,#0.00")
        'txtMon.Text = Format(CDbl(txtDay.Text) * CDbl(yearday), "#,##0.00")
        'txtYear.Text = Format(CDbl(txtDay.Text) * CDbl(days), "#,##0.00")

    End Sub

    Private Sub BtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch.Click

        fmShartOfAccDetail.txtSty.Text = "FrmAssetNew"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()

    End Sub

    Private Sub txtYear_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtYear.TextChanged

    End Sub

    Private Sub Label34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label34.Click

    End Sub

    Private Sub txtDay_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDay.TextChanged

    End Sub

    Private Sub txtBudget_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBudget.SelectedIndexChanged

    End Sub

    Private Sub DTUse_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTUse.ValueChanged
        'DATEADD(YEAR, + Used_Life, Date_Work)
        DTEnd.Text = DateAdd(DateInterval.Year, CDbl(txtLife.Text), CDate(DTUse.Value))

        Dim yearday, days As Integer
        Dim DT_date As Date

        DT_date = DateAdd(DateInterval.Year, CDbl(txtLife.Text), CDate(DTUse.Value))
        days = DateDiff(DateInterval.Day, CDate(DTUse.Value), CDate(DT_date))
        yearday = System.DateTime.DaysInMonth(Year(DTUse.Value), Month(DTUse.Value))

        txtDay.Text = Format(CDbl(txtKIP.Text) / CDbl(days), "#,#0.00")
        If (CDbl(txtLife.Text) * 12) <> 0 Then
            txtMon.Text = Format(CDbl(txtKIP.Text) / (CDbl(txtLife.Text) * 12), "#,##0.00")
        Else
            txtMon.Text = 0
        End If

        txtYear.Text = Format(CDbl(txtKIP.Text) / CDbl(txtLife.Text), "#,##0.00")

        'txtYear.Text = Format(CDbl(txtKIP.Text) / CDbl(txtLife.Text), "#,#0.00")
        'txtMon.Text = Format(CDbl(txtKIP.Text) / (CDbl(txtLife.Text) * 12), "#,#0.00")
        ''txtDay.Text = Format(CDbl(txtKIP.Text) / 365, "#,#0.00")
        'txtDay.Text = Format(CDbl(txtKIP.Text) / CDbl(days), "#,#0.00")

    End Sub

    Private Sub DTBuy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTBuy.ValueChanged

    End Sub
End Class