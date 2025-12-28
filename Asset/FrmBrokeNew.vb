Public Class FrmBrokeNew
    Dim Lin As Integer = 0
    Dim DA As Integer

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub MAHA_ACC()

        If chkAcc.Checked = True Then
            CNN.Execute("Delete from gen_jn where Book='AS' and certify=N'" & Trim(txtID.Text) & "' ")
            If CDbl(txtCost.Text) <= 0 Then
                'CNN.Execute("INSERT INTO gen_jn(certify,company,  Com_id,don_id, office_id, Book,amount, net_amt, date_work, code_dr,code_cr,ac_code, descrip,amount_dr,amount_cr,amt_dr,amt_cr,last_update,last_user,my_lock,amt_USD_dr,amt_USD_cr) " & _
                '  " VALUES('" & Trim(txtID.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','Fixd Asset', " & CDbl(txtKIP.Text) & "," & CDbl(txtKIP.Text) & ",'" & Format(dtDate.Value, "yyyy-MM-dd") & "','" & TextBox2.Text & "','','" & TextBox2.Text & "',N'" & TxtGrpName.Text & "'," & CDbl(TextBox1.Text) & ",0," & CDbl(TextBox1.Text) & ",0,'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "','0','0','0')")
                ''========= ໝມີ
                'CNN.Execute("INSERT INTO gen_jn(certify,company,  Com_id,don_id, office_id, Book,amount, net_amt, date_work, code_dr,code_cr,ac_code, descrip,amount_dr,amount_cr,amt_dr,amt_cr,last_update,last_user,my_lock,amt_USD_dr,amt_USD_cr) " & _
                '" VALUES('" & Trim(txtID.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','Fixd Asset', " & CDbl(txtKIP.Text) & "," & CDbl(txtKIP.Text) & ",'" & Format(dtDate.Value, "yyyy-MM-dd") & "','','" & txtGrp.Text & "','" & txtGrp.Text & "',N'" & TxtGrpName.Text & "',0," & CDbl(txtKIP.Text) & ",0," & CDbl(txtKIP.Text) & ",'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "','0','0','0')")

            Else
                '=== ຫຼຸ້ຍຫ້ຽນໃນເດືອນ Debit
                If CDbl(TxtDep_Month.Text) <> 0 Then
                    Dim CNDR As String = "INSERT INTO gen_jn(certify,Descrip,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                 " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id,AG,Frm) " & _
                   " VALUES(N'" & Trim(txtID.Text) & "'," & _
                        "N'" & (txtDesc.Text) & "'," & _
                 " '" & Format(CDate(dtDate.Value), "yyyy-MM-dd") & "'," & _
                    "N'AS'," & _
                   "N'" & Trim(txtID.Text) & "'," & _
                           "N'" & Trim(txtID.Text) & "'," & _
                                  "N''," & _
                                "" & CDbl(TxtDep_Month.Text) & "," & _
                            "N'LAK'," & _
                        "" & CDbl(1) & "," & _
                          "N'LAK'," & _
                        "" & CDbl(1) & "," & _
                           "" & CDbl(TxtDep_Month.Text) * CDbl(1) & "," & _
                   "N'" & TxtDr.Text & "'," & _
                    "N''," & _
                  "N'" & TxtDr.Text & "'," & _
                  "N''," & _
                   "" & CDbl(TxtDep_Month.Text) & "," & _
                   " 0," & _
                        "" & CDbl(TxtDep_Month.Text) * CDbl(1) & "," & _
                   " 0," & _
                      " 0," & _
                         " 0," & _
                    " 1," & _
                        " 1," & _
                   " Getdate()," & _
                 "N'" & MUserID & "'," & _
                 "N'" & MuSubOff2 & "',0,'1' )"
                    CNN.Execute(CNDR)
                    '=== ຫຼຸ້ຍຫ້ຽນໃນເດືອນ Credit

                    Dim CNCr As String = "INSERT INTO gen_jn(certify,Descrip,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                  " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id,AG,Frm) " & _
                    " VALUES(N'" & Trim(txtID.Text) & "'," & _
                    "N'" & (txtDesc.Text) & "'," & _
                  " '" & Format(CDate(dtDate.Value), "yyyy-MM-dd") & "'," & _
                 "N'AS'," & _
                   "N'" & Trim(txtID.Text) & "'," & _
                      "N'" & Trim(txtID.Text) & "'," & _
                                   "N''," & _
                                 "" & CDbl(TxtDep_Month.Text) & "," & _
                        "N'LAK'," & _
                         "" & CDbl(1) & "," & _
                            "N'LAK'," & _
                         "" & CDbl(1) & "," & _
                            "" & CDbl(TxtDep_Month.Text) * CDbl(1) & "," & _
                                               "N''," & _
                    "N'" & TxtCr.Text & "'," & _
                   "N'" & TxtCr.Text & "'," & _
                   "N''," & _
          " 0," & _
                    "" & CDbl(TxtDep_Month.Text) & "," & _
                    " 0," & _
                      "" & CDbl(TxtDep_Month.Text) * CDbl(1) & "," & _
                    " 0," & _
                       " 0," & _
                     " 1," & _
                         " 1," & _
                    " Getdate()," & _
                  "N'" & MUserID & "'," & _
                  "N'" & MuSubOff2 & "',0,'1')"
                    CNN.Execute(CNCr)
                End If

                '================================================
                '=== ສະສົມ ລວມ Debit
                If CDbl(TextBox1.Text) <> 0 Then


                    Dim DR22 As String = "INSERT INTO gen_jn(certify,Descrip,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
             " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id,AG,Frm) " & _
               " VALUES(N'" & Trim(txtID.Text) & "'," & _
                    "N'" & (txtDesc.Text) & "'," & _
             " '" & Format(CDate(dtDate.Value), "yyyy-MM-dd") & "'," & _
                "N'AS'," & _
               "N'" & Trim(txtID.Text) & "'," & _
                       "N'" & Trim(txtID.Text) & "'," & _
                              "N''," & _
                            "" & CDbl(TextBox1.Text) & "," & _
                        "N'LAK'," & _
                    "" & CDbl(1) & "," & _
                      "N'LAK'," & _
                    "" & CDbl(1) & "," & _
                       "" & CDbl(TextBox1.Text) * CDbl(1) & "," & _
               "N'" & TxtDr22.Text & "'," & _
                "N''," & _
              "N'" & TxtDr22.Text & "'," & _
              "N''," & _
               "" & CDbl(TextBox1.Text) & "," & _
               " 0," & _
                    "" & CDbl(TextBox1.Text) * CDbl(1) & "," & _
               " 0," & _
                  " 0," & _
                     " 0," & _
                " 1," & _
                    " 1," & _
               " Getdate()," & _
             "N'" & MUserID & "'," & _
             "N'" & MuSubOff2 & "',0,'1' )"
                    CNN.Execute(DR22)
                End If
                '=== ສະສົມ ລວມ Net Book Value (NBV)
                If CDbl(txtCost.Text) <> 0 Then


                    Dim DR33 As String = "INSERT INTO gen_jn(certify,Descrip,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
             " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id,AG,Frm) " & _
               " VALUES(N'" & Trim(txtID.Text) & "'," & _
                    "N'" & (txtDesc.Text) & "'," & _
             " '" & Format(CDate(dtDate.Value), "yyyy-MM-dd") & "'," & _
                "N'AS'," & _
               "N'" & Trim(txtID.Text) & "'," & _
                       "N'" & Trim(txtID.Text) & "'," & _
                              "N''," & _
                            "" & CDbl(txtCost.Text) & "," & _
                        "N'LAK'," & _
                    "" & CDbl(1) & "," & _
                      "N'LAK'," & _
                    "" & CDbl(1) & "," & _
                       "" & CDbl(txtCost.Text) * CDbl(1) & "," & _
               "N'" & TxtDr33.Text & "'," & _
                "N''," & _
              "N'" & TxtDr33.Text & "'," & _
              "N''," & _
               "" & CDbl(txtCost.Text) & "," & _
               " 0," & _
                    "" & CDbl(txtCost.Text) * CDbl(1) & "," & _
               " 0," & _
                  " 0," & _
                     " 0," & _
                " 1," & _
                    " 1," & _
               " Getdate()," & _
             "N'" & MUserID & "'," & _
             "N'" & MuSubOff2 & "',0,'1' )"
                    CNN.Execute(DR33)
                End If

                '=== Acquisition Cost / Purchase Price Credit
                If CDbl(txtKIP.Text) <> 0 Then

                    Dim Cr22 As String = "INSERT INTO gen_jn(certify,Descrip,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                  " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id,AG,Frm) " & _
                    " VALUES(N'" & Trim(txtID.Text) & "'," & _
                    "N'" & (txtDesc.Text) & "'," & _
                  " '" & Format(CDate(dtDate.Value), "yyyy-MM-dd") & "'," & _
                 "N'AS'," & _
                   "N'" & Trim(txtID.Text) & "'," & _
                      "N'" & Trim(txtID.Text) & "'," & _
                                   "N''," & _
                                 "" & CDbl(txtKIP.Text) & "," & _
                        "N'LAK'," & _
                         "" & CDbl(1) & "," & _
                            "N'LAK'," & _
                         "" & CDbl(1) & "," & _
                            "" & CDbl(txtKIP.Text) * CDbl(1) & "," & _
                                               "N''," & _
                    "N'" & TxtCr22.Text & "'," & _
                   "N'" & TxtCr22.Text & "'," & _
                   "N''," & _
          " 0," & _
                    "" & CDbl(txtKIP.Text) & "," & _
                    " 0," & _
                      "" & CDbl(txtKIP.Text) * CDbl(1) & "," & _
                    " 0," & _
                       " 0," & _
                     " 1," & _
                         " 1," & _
                    " Getdate()," & _
                  "N'" & MUserID & "'," & _
                  "N'" & MuSubOff2 & "',0,'1')"
                    CNN.Execute(Cr22)
                End If
            End If
        End If
    End Sub
    'Private Sub RunID()
    '    Dim rRS As New ADODB.Recordset
    '    Dim mNum2 As String
    '    Call LoadSqlData("Select Top 1 BrokenID From Brokens Order by BrokenID DESC ", rRS)
    '    If rRS.RecordCount <> 0 Then
    '        mNum2 = Val(rRS.Fields("BrokenID").Value) + 1
    '        If Len(CStr(mNum2).Trim) = 1 Then
    '            txtID.Text = "0000" & CStr(mNum2)
    '        ElseIf Len(CStr(mNum2).Trim) = 2 Then
    '            txtID.Text = "000" & CStr(mNum2)
    '        ElseIf Len(CStr(mNum2).Trim) = 3 Then
    '            txtID.Text = "00" & CStr(mNum2)
    '        ElseIf Len(CStr(mNum2).Trim) = 4 Then
    '            txtID.Text = "0" & CStr(mNum2)
    '        Else
    '            txtID.Text = CStr(mNum2)
    '        End If
    '    Else
    '        txtID.Text = "00001"
    '    End If
    'End Sub
    Private Sub RunID()
        Dim rRS As New ADODB.Recordset
        Dim mNum2 As String
        Call LoadSqlData("Select Top 1 BrokenID, substring(BrokenID,6,6) asNo From Brokens Order by BrokenID DESC ", rRS)
        If rRS.RecordCount <> 0 Then
            mNum2 = Val(rRS.Fields("asNo").Value) + 1
            If Len(CStr(mNum2).Trim) = 1 Then
                txtID.Text = "BK" & Format(Date.Today, "yy") & "." & "0000" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 2 Then
                txtID.Text = "BK" & Format(Date.Today, "yy") & "." & "000" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 3 Then
                txtID.Text = "BK" & Format(Date.Today, "yy") & "." & "00" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 4 Then
                txtID.Text = "BK" & Format(Date.Today, "yy") & "." & "0" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 5 Then
                txtID.Text = CStr(mNum2)
            End If
        Else
            txtID.Text = "BK" & Format(Date.Today, "yy") & "." & "00001"
        End If
    End Sub
    Private Sub FrmAssetNew_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Call Loadlang()
        'SetControlText(Me)
        'ChgChildForm()
        dpFromDate.Value = "01/" & Format((dtDate.Value), "MM") & "/" & Format(dtDate.Value, "yyyy")
        AddText()
        If mEdit = True Then
            Call loadAST()
            txtID.Enabled = False
        Else
            'Call btnAdd_Click(sender, e)
            AddText()
            AddText()
        End If

    End Sub

    Private Sub loadAST()
        Dim aRS As New ADODB.Recordset
        'Call LoadSqlData("Select * from Brokens where BrokenID='" & myTemp & "' ", aRS)
        LoadSqlData("SELECT dbo.Brokens.*, dbo.Groups.Group_Nm, dbo.Groups.Dep_Code, dbo.Groups.Ac_Code, dbo.Groups.AccountCodeAsCR, dbo.Groups.AccountCodeAsDR,  " & _
                      "  dbo.Groups.AccountCodeBrokenDR FROM dbo.Brokens INNER JOIN " & _
                      "   dbo.Groups ON dbo.Brokens.group_id = dbo.Groups.Group_ID where Brokens.BrokenID=N'" & Trim(myTemp) & "' ", aRS)
        With aRS
            If aRS.RecordCount <> 0 Then
                txtID.Text = Trim(aRS.Fields("BrokenID").Value.ToString)
                txtDesc.Text = Trim(aRS.Fields("Descriptions").Value.ToString)
        
                txtAS.Text = Trim(aRS.Fields("AssetID").Value.ToString)
                txtNm.Text = Trim(aRS.Fields("AssetNm").Value.ToString)
                txtRemark.Text = Trim(aRS.Fields("Remark").Value.ToString)
                dtDate.Value = Format(CDate(aRS.Fields("BrokenDate").Value), "dd/MM/yyyy")
                txtGrp.Text = Trim(aRS.Fields("Group_ID").Value.ToString)
                TextBox2.Text = Trim(aRS.Fields("Group_ID").Value.ToString)
                txtKIP.Text = Format(aRS.Fields("Amt_KIP").Value, "#,##0.00")
                txtLife.Text = Format(aRS.Fields("Used_Life").Value, "#,##0")
                txtYear.Text = Format(aRS.Fields("Dep_Year").Value, "#,##0.00")
                txtMon.Text = Format(aRS.Fields("Dep_Month").Value, "#,##0.00")
                TextBox1.Text = Format(aRS.Fields("Amt_all").Value, "#,##0.00")

                TxtDep_Month.Text = Format(aRS.Fields("Dep_Monthly").Value, "#,##0.00")
                TxtDay.Text = Format(aRS.Fields("Dep_Day").Value, "#,##0.00")

                txtCost.Text = Format(aRS.Fields("Cost").Value, "#,##0.00")
                txtAmt.Text = Format(aRS.Fields("Amount").Value, "#,##0.00")

                DTBuy.Value = Format(CDate(aRS.Fields("Date_Buy").Value), "dd/MM/yyyy")
                DTUse.Value = Format(CDate(aRS.Fields("Date_Work").Value), "dd/MM/yyyy")
                txtDepart.Text = Trim(aRS.Fields("DepartmentID").Value.ToString)
                TxtGrpName.Text = Trim(aRS.Fields("Group_Nm").Value.ToString)
                txtBrokenDR.Text = Trim(aRS.Fields("AccountCodeBrokenDR").Value.ToString)
                txtSec_ID.Text = Trim(aRS.Fields("Sect_ID").Value.ToString)

                TxtDr.Text = Trim(aRS.Fields("Dr").Value.ToString)
                Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtDr.Text & "' ", RSC)
                If RSC.RecordCount <> 0 Then
                    TxtDrNm.Text = Trim(RSC.Fields("Name_L").Value.ToString)
                End If

                TxtCr.Text = Trim(aRS.Fields("Cr").Value.ToString)
                Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtCr.Text & "' ", RSC)
                If RSC.RecordCount <> 0 Then
                    TxtCrNm.Text = Trim(RSC.Fields("Name_L").Value.ToString)
                End If

                TxtDr22.Text = Trim(aRS.Fields("Dr22").Value.ToString)
                Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtDr22.Text & "' ", RSC)
                If RSC.RecordCount <> 0 Then
                    TxtDrNm22.Text = Trim(RSC.Fields("Name_L").Value.ToString)
                End If

                TxtDr33.Text = Trim(aRS.Fields("Dr33").Value.ToString)

                Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtDr33.Text & "' ", RSC)
                If RSC.RecordCount <> 0 Then
                    TxtDrNm33.Text = Trim(RSC.Fields("Name_L").Value.ToString)
                End If

                TxtCr22.Text = Trim(aRS.Fields("Cr22").Value.ToString)
                Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtCr22.Text & "' ", RSC)
                If RSC.RecordCount <> 0 Then
                    TxtCrNm22.Text = Trim(RSC.Fields("Name_L").Value.ToString)
                End If

            End If
        End With
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub cmdNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtAS_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAS.KeyPress
        If e.KeyChar = Chr(13) Then
            Dim aRS As New ADODB.Recordset
            Call LoadSqlData("SELECT     dbo.Assets.*, dbo.Groups.Group_Nm AS Expr1, dbo.Groups.Ac_Code AS Expr2, dbo.Groups.AccountCodeAsCR, dbo.Groups.AccountCodeAsDR, " & _
                    " dbo.Groups.AgingFixAssets, dbo.Groups.AccountCodeBrokenDR FROM dbo.Assets INNER JOIN  " & _
                     " dbo.Groups ON dbo.Assets.Group_ID = dbo.Groups.Group_ID where Assets.AssetID='" & Trim(txtAS.Text) & "'", aRS)
            If aRS.RecordCount <> 0 Then
                TextBox2.Text = ""
                txtBrokenDR.Text = Trim(aRS.Fields("AccountCodeBrokenDR").Value.ToString)
                txtNm.Text = Trim(aRS.Fields("Asset_Nm").Value.ToString)
                TxtCompany.Text = Trim(aRS.Fields("Company").Value.ToString)
                txtGrp.Text = Trim(aRS.Fields("Group_ID").Value.ToString)
                TxtGrpName.Text = Trim(aRS.Fields("Group_Nm").Value.ToString)
                TextBox2.Text = Trim(aRS.Fields("Expr2").Value.ToString)
                txtKIP.Text = Format((aRS.Fields("Amt_KIP").Value), "#,##0.00")
                txtYear.Text = Format((aRS.Fields("Dep_Year").Value), "#,##0.00")
                txtMon.Text = Format((aRS.Fields("Dep_Month").Value), "#,##0.00")
                txtLife.Text = Format((aRS.Fields("Used_Life").Value), "#,##0.00")
                txtDepart.Text = Trim(aRS.Fields("DepartmentID").Value.ToString)
                DTBuy.Value = Format(aRS.Fields("Date_Buy").Value, "dd/MM/yyyy")
                DTUse.Value = Format(aRS.Fields("Date_Work").Value, "dd/MM/yyyy")
                txtSec_ID.Text = Trim(aRS.Fields("Sect_ID").Value.ToString)
                'Lin = txtLife.Text * 12
                Lin = DateDiff(DateInterval.Month, DTUse.Value, dtDate.Value)
                TextBox1.Text = Format((Lin * txtMon.Text), "#,##0.00")
                txtCost.Text = Format((txtKIP.Text - TextBox1.Text), "#,##0.00")
                DA = txtCost.Text
                txtAmt.Text = Format((DA), "#,##0.00")
                If CDbl(TextBox1.Text) > CDbl(txtKIP.Text) Then
                    TextBox1.Text = txtKIP.Text
                    txtCost.Text = 0
                    txtAmt.Text = 0
                End If
                txtAmt.Focus()
            Else
                txtNm.Text = ""
                txtGrp.Text = ""
            End If
            Dim D As Integer = DateDiff(DateInterval.Day, (dpFromDate.Value), dtDate.Value)
            Dim MD As Integer = DateDiff(DateInterval.Day, DTUse.Value, dtDate.Value)
            Dim MMamt As Double = 0
            TextBox1.Text = CDbl(txtKIP.Text) / CDbl(txtLife.Text) / 360 * CDbl(MD)
            TextBox1.Text = Format(CDbl(TextBox1.Text), "#,##0.00")
            D = D + 1
            TxtDep_Month.Text = CDbl(txtKIP.Text) / CDbl(txtLife.Text) / 360 * CDbl(D)
            TxtDep_Month.Text = Format(CDbl(TxtDep_Month.Text), "#,##0.00")

            txtCost.Text = Format((txtKIP.Text - TextBox1.Text), "#,##0.00")
            txtAmt.Text = Format(CDbl(txtCost.Text), "#,##0.00")
            If CDbl(TextBox1.Text) > CDbl(txtKIP.Text) Then
                TextBox1.Text = txtKIP.Text
                txtCost.Text = 0
                txtAmt.Text = 0
            End If



        End If
    End Sub

    Private Sub txtAS_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAS.TextChanged

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        myTemp = ""
        FrmList_ASST.ShowDialog()
        If myTemp <> "" Then
            txtAS.Text = myTemp
            txtNm.Text = myTemp1
            Dim aRS As New ADODB.Recordset
            Call LoadSqlData("SELECT     dbo.Assets.*, dbo.Groups_Asset.Group_Nm AS Expr1, dbo.Groups_Asset.Ac_Code AS Expr2, dbo.Groups_Asset.AccountCodeAsCR, dbo.Groups_Asset.AccountCodeAsDR, " & _
                  " dbo.Groups_Asset.AgingFixAssets, dbo.Groups_Asset.AccountCodeBrokenDR FROM dbo.Assets INNER JOIN  " & _
                   " dbo.Groups_Asset ON dbo.Assets.Group_ID = dbo.Groups_Asset.Group_ID where Assets.AssetID='" & Trim(txtAS.Text) & "'", aRS)
            If aRS.RecordCount <> 0 Then
                txtBrokenDR.Text = Trim(aRS.Fields("AccountCodeBrokenDR").Value.ToString)
                txtNm.Text = Trim(aRS.Fields("Asset_Nm").Value.ToString)
                TxtCompany.Text = Trim(aRS.Fields("Company").Value.ToString)
                txtGrp.Text = Trim(aRS.Fields("Group_ID").Value.ToString)
                TxtGrpName.Text = Trim(aRS.Fields("Group_Nm").Value.ToString)
                TextBox2.Text = Trim(aRS.Fields("Ac_Code").Value.ToString)
                txtKIP.Text = Format((aRS.Fields("Amt_KIP").Value), "#,##0.00")
                txtYear.Text = Format((aRS.Fields("Dep_Year").Value), "#,##0.00")
                txtMon.Text = Format((aRS.Fields("Dep_Month").Value), "#,##0.00")
                TxtDay.Text = Format((aRS.Fields("Dep_Day").Value), "#,##0.00")

                txtLife.Text = Format((aRS.Fields("Used_Life").Value), "#,##0")
                txtDepart.Text = Trim(aRS.Fields("DepartmentID").Value.ToString)
                DTBuy.Value = Format(aRS.Fields("Date_Buy").Value, "dd/MM/yyyy")
                DTUse.Value = Format(aRS.Fields("Date_Work").Value, "dd/MM/yyyy")
                txtSec_ID.Text = Trim(aRS.Fields("Sect_ID").Value.ToString)
                TxtCompany.Text = Trim(aRS.Fields("Sect_ID").Value.ToString)
                'Lin = txtLife.Text * 12
                Lin = DateDiff(DateInterval.Month, DTUse.Value, dtDate.Value)
                TextBox1.Text = Format((Lin * txtMon.Text), "#,##0.00")
                txtCost.Text = Format((txtKIP.Text - TextBox1.Text), "#,##0.00")
                'DA = txtCost.Text
                txtAmt.Text = Format(CDbl(txtCost.Text), "#,##0.00")
                If CDbl(TextBox1.Text) > CDbl(txtKIP.Text) Then
                    TextBox1.Text = txtKIP.Text
                    txtCost.Text = 0
                    txtAmt.Text = 0
                End If
                Dim D As Integer = DateDiff(DateInterval.Day, (dpFromDate.Value), dtDate.Value)
                Dim MD As Integer = DateDiff(DateInterval.Day, DTUse.Value, dtDate.Value)
                Dim MMamt As Double = 0
                TextBox1.Text = CDbl(txtKIP.Text) / CDbl(txtLife.Text) / 360 * CDbl(MD)
                TextBox1.Text = Format(CDbl(TextBox1.Text), "#,##0.00")
                D = D + 1
                TxtDep_Month.Text = CDbl(txtKIP.Text) / CDbl(txtLife.Text) / 360 * CDbl(D)
                TxtDep_Month.Text = Format(CDbl(TxtDep_Month.Text), "#,##0.00")

                txtCost.Text = Format((txtKIP.Text - TextBox1.Text), "#,##0.00")
                txtAmt.Text = Format(CDbl(txtCost.Text), "#,##0.00")
                If CDbl(TextBox1.Text) > CDbl(txtKIP.Text) Then
                    TextBox1.Text = txtKIP.Text
                    txtCost.Text = 0
                    txtAmt.Text = 0
                End If


                txtAmt.Focus()
            Else
                txtNm.Text = ""
                txtGrp.Text = ""
            End If
        End If

    End Sub

    Private Sub dtDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtDate.ValueChanged
        dpFromDate.Value = "01/" & Format((dtDate.Value), "MM") & "/" & Format(dtDate.Value, "yyyy")

        Lin = DateDiff(DateInterval.Month, DTUse.Value, dtDate.Value)
        TextBox1.Text = Format((Lin * txtMon.Text), "#,##0")
        txtCost.Text = Format((txtKIP.Text - TextBox1.Text), "#,##0")
        DA = txtCost.Text
        txtAmt.Text = Format((DA), "#,##0")
        If CDbl(TextBox1.Text) > CDbl(txtKIP.Text) Then
            TextBox1.Text = txtKIP.Text
            txtCost.Text = 0
            txtAmt.Text = 0
        End If
        Dim D As Integer = DateDiff(DateInterval.Day, (dpFromDate.Value), dtDate.Value)
        Dim MD As Integer = DateDiff(DateInterval.Day, DTUse.Value, dtDate.Value)
        Dim MMamt As Double = 0
        TextBox1.Text = CDbl(txtKIP.Text) / CDbl(txtLife.Text) / 360 * CDbl(MD)
        TextBox1.Text = Format(CDbl(TextBox1.Text), "#,##0.00")
        D = D + 1
        TxtDep_Month.Text = CDbl(txtKIP.Text) / CDbl(txtLife.Text) / 360 * CDbl(D)
        TxtDep_Month.Text = Format(CDbl(TxtDep_Month.Text), "#,##0.00")

        txtCost.Text = Format((txtKIP.Text - TextBox1.Text), "#,##0.00")
        txtAmt.Text = Format(CDbl(txtCost.Text), "#,##0.00")
        If CDbl(TextBox1.Text) > CDbl(txtKIP.Text) Then
            TextBox1.Text = txtKIP.Text
            txtCost.Text = 0
            txtAmt.Text = 0
        End If



    End Sub

    Private Sub txtLife_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLife.TextChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        CNN.Execute("Delete from gen_jn where Book='AS' and certify='" & Trim(txtID.Text) & "' ")
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        AddText()
        AddText()
    End Sub
    Private Sub AddText()
        Call RunID()
        txtDesc.Text = ""
        txtDepart.Text = ""
        txtCost.Text = 0
        txtKIP.Text = 0
        txtYear.Text = 0
        txtMon.Text = 0
        txtLife.Text = 0
        TextBox1.Text = 0
        txtAmt.Text = 0
        TxtDay.Text = 0
        TxtDep_Month.Text = 0
        txtAS.Text = ""
        txtNm.Text = ""
        txtRemark.Text = ""
        'DTBuy.Value = Format(mDate, "dd/MM/yyyy")
        dtDate.Value = Format(mDate, "dd/MM/yyyy")
        DTUse.Value = Today
        DTBuy.Text = Today

        txtGrp.Text = ""
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim cRS As New ADODB.Recordset
        Dim ss As String
        Call LoadSqlData("select BrokenID from Brokens where AssetID = '" & Trim(txtAS.Text) & "' AND BrokenID<>'" & Trim(txtID.Text.ToString) & "' ", cRS)
        If cRS.RecordCount <> 0 Then
            MsgBox("ລາຍການນີ້ໄດ້ສະສາງແລ້ວ")
            Exit Sub
        End If
        Call LoadSqlData("select BrokenID from Brokens where BrokenID='" & Trim(txtID.Text.ToString) & "' ", cRS)
        If cRS.RecordCount = 0 Then
            RunID()
            CNN.Execute("INSERT INTO Brokens(BrokenID, BrokenDate, AssetID, AssetNm, Descriptions, Cost, Remark, Group_ID, Amount, Amt_KIP, Used_Life, Dep_Year, Dep_Month, Sect_ID, Date_Buy, Date_Work,DepartmentID,Amt_All,Dep_Monthly,Dep_Day,Dr,Cr,Dr22,Dr33,Cr22) " & _
                        " VALUES('" & Trim(txtID.Text.ToString) & "', '" & Format(dtDate.Value, "yyyy-MM-dd") & "', '" & Trim(txtAS.Text.ToArray) & "', N'" & Trim(txtNm.Text.ToString) & "', N'" & Trim(txtDesc.Text.ToString) & "',  " & CDbl(txtCost.Text) & ", N'" & Trim(txtRemark.Text.ToString) & "', '" & Trim(txtGrp.Text.ToString) & "', " & CDbl(txtAmt.Text) & ", " & _
                        " " & CDbl(txtKIP.Text) & ", " & CDbl(txtLife.Text) & ", " & CDbl(txtYear.Text) & ", " & CDbl(txtMon.Text) & ", N'" & Trim(txtSec_ID.Text.ToString) & "', '" & Format(DTBuy.Value, "yyyy-MM-dd") & "', '" & Format(DTUse.Value, "yyyy-MM-dd") & "',N'" & Trim(txtDepart.Text.ToString) & "'," & CDbl(TextBox1.Text) & "," & CDbl(TxtDep_Month.Text) & "," & CDbl(TxtDay.Text) & "," & _
                        " N'" & Trim(TxtDr.Text) & "',N'" & Trim(TxtCr.Text) & "',N'" & Trim(TxtDr22.Text) & "',N'" & Trim(TxtDr33.Text) & "',N'" & Trim(TxtCr22.Text) & "')")
            '==================================================================================
            '==========ACC=======================================================================
 
        Else
            ss = "Update Brokens Set BrokenDate='" & Format(dtDate.Value, "yyyy-MM-dd") & "',AssetID='" & Trim(txtAS.Text.ToArray) & "',AssetNm=N'" & Trim(txtNm.Text.ToString) & "'," & _
            "Descriptions=N'" & Trim(txtDesc.Text) & "',Cost=" & CDbl(txtCost.Text) & ",Remark=N'" & Trim(txtRemark.Text.ToString) & "',Group_ID='" & Trim(txtGrp.Text.ToString) & "'," & _
                "Amount=" & CDbl(txtAmt.Text) & ",Amt_KIP=" & CDbl(txtKIP.Text) & ",Used_Life=" & CDbl(txtLife.Text) & ",Dep_Year=" & CDbl(txtYear.Text) & ",Dep_Month=" & CDbl(txtMon.Text) & "," & _
             " Sect_ID=N'" & Trim(txtSec_ID.Text.ToString) & "',Date_Buy='" & Format(DTBuy.Value, "yyyy-MM-dd") & "',Date_Work='" & Format(DTUse.Value, "yyyy-MM-dd") & "',DepartmentID=N'" & Trim(txtDepart.Text.ToString) & "',Amt_All=" & CDbl(TextBox1.Text) & ",Dep_Monthly=" & CDbl(TxtDep_Month.Text) & ",Dep_Day=" & CDbl(TxtDay.Text) & ", " & _
             " Dr=N'" & Trim(TxtDr.Text) & "',Cr=N'" & Trim(TxtCr.Text) & "',Dr22=N'" & Trim(TxtDr22.Text) & "',Dr33=N'" & Trim(TxtDr33.Text) & "',Cr22=N'" & Trim(TxtCr22.Text) & "' " & _
                " Where BrokenID='" & Trim(txtID.Text.ToString) & "' "
            CNN.Execute(ss)
            '==========ACC=======================================================================
           
        End If
        ss = "Update Assets Set Deposted=1,  Deposted_Date='" & Format(dtDate.Value, "yyyy-MM-dd") & "',AmountRemain=" & CDbl(txtCost.Text) & ",AmountClear=" & CDbl(txtAmt.Text) & ",Amt_All=" & CDbl(TextBox1.Text) & " " & _
        " Where AssetID='" & Trim(txtAS.Text.ToString) & "' "
        CNN.Execute(ss)
   
        '=ACC====
        Call MAHA_ACC()
        If chkAcc.Checked = True Then
            CNN.Execute("update Gen_jn set Gen_jn.descrip=Acc_Code.Name_L, Gen_jn.ac_name=Acc_Code.Name_L, Gen_jn.curr='LAK', Gen_jn.rate=1  from Acc_Code,Gen_jn where Gen_jn.certify='" & Trim(txtID.Text) & "' and Gen_jn.AC_Code=ACC_Code.AC_Code ")
        End If
        MsgBox("Finish")
        'If chkAcc.Checked = True Then
        '    CNN.Execute("Delete from gen_jn where Book='Fixd Asset' and certify='" & Trim(txtID.Text) & "' ")
        '    If CDbl(txtCost.Text) <= 0 Then
        '        CNN.Execute("INSERT INTO gen_jn(certify, company,  Com_id,don_id, office_id,Book,amount, net_amt, date_work, code_dr,code_cr,ac_code, descrip,amount_dr,amount_cr,amt_dr,amt_cr,last_update,last_user,my_lock,amt_USD_dr,amt_USD_cr) " & _
        '          " VALUES('" & Trim(txtID.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','Fixd Asset', " & CDbl(txtKIP.Text) & "," & CDbl(txtKIP.Text) & ",'" & Format(dtDate.Value, "yyyy-MM-dd") & "','" & TextBox2.Text & "','','" & TextBox2.Text & "',N'" & TxtGrpName.Text & "'," & CDbl(txtAmt.Text) & ",0," & CDbl(txtAmt.Text) & ",0,'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "','0','0','0')")
        '        '========= ໝມີ
        '        CNN.Execute("INSERT INTO gen_jn(certify,company,  Com_id,don_id, office_id,Book,amount, net_amt, date_work, code_dr,code_cr,ac_code, descrip,amount_dr,amount_cr,amt_dr,amt_cr,last_update,last_user,my_lock,amt_USD_dr,amt_USD_cr) " & _
        '        " VALUES('" & Trim(txtID.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','Fixd Asset', " & CDbl(txtKIP.Text) & "," & CDbl(txtKIP.Text) & ",'" & Format(dtDate.Value, "yyyy-MM-dd") & "','','" & txtGrp.Text & "','" & txtGrp.Text & "',N'" & TxtGrpName.Text & "',0," & CDbl(txtAmt.Text) & ",0," & CDbl(txtAmt.Text) & ",'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "','0','0','0')")

        '    Else
        '        '=== ສະສົມ ລວມ
        '        ' CNN.Execute("INSERT INTO gen_jn(certify,company,  Com_id,don_id, office_id,Book,amount, net_amt, date_work, code_dr,code_cr,ac_code, descrip,amount_dr,amount_cr,amt_dr,amt_cr,last_update,last_user,my_lock,amt_USD_dr,amt_USD_cr) " & _
        '        '" VALUES('" & Trim(txtID.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','Fixd Asset', " & CDbl(txtKIP.Text) & "," & CDbl(txtKIP.Text) & ",'" & Format(dtDate.Value, "yyyy-MM-dd") & "','" & TextBox2.Text & "','','" & TextBox2.Text & "',N'" & TxtGrpName.Text & "'," & CDbl(TextBox1.Text) & ",0," & CDbl(TextBox1.Text) & ",0,'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "','0','0','0')")
        '        '=== ຍັງເຫຼືອ  ໜີ້
        '        CNN.Execute("INSERT INTO gen_jn(certify,company,  Com_id,don_id, office_id,Book,amount, net_amt, date_work, code_dr,code_cr,ac_code, descrip,amount_dr,amount_cr,amt_dr,amt_cr,last_update,last_user,my_lock,amt_USD_dr,amt_USD_cr) " & _
        '        " VALUES('" & Trim(txtID.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','Fixd Asset', " & CDbl(txtAmt.Text) & "," & CDbl(txtAmt.Text) & ",'" & Format(dtDate.Value, "yyyy-MM-dd") & "','" & txtBrokenDR.Text & "','','" & txtBrokenDR.Text & "',N'" & TxtGrpName.Text & "'," & CDbl(txtAmt.Text) & ",0," & CDbl(txtAmt.Text) & ",0,'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "','0','0','0')")
        '        '=== ມີ
        '        CNN.Execute("INSERT INTO gen_jn(certify,company, Com_id,don_id, office_id, Book,amount, net_amt, date_work, code_dr,code_cr,ac_code, descrip,amount_dr,amount_cr,amt_dr,amt_cr,last_update,last_user,my_lock,amt_USD_dr,amt_USD_cr) " & _
        '      " VALUES('" & Trim(txtID.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','" & Trim(TxtCompany.Text) & "','Fixd Asset', " & CDbl(txtAmt.Text) & "," & CDbl(txtAmt.Text) & ",'" & Format(dtDate.Value, "yyyy-MM-dd") & "','','" & txtGrp.Text & "','" & txtGrp.Text & "',N'" & TxtGrpName.Text & "',0," & CDbl(txtAmt.Text) & ",0," & CDbl(txtAmt.Text) & ",'" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "','0','0','0')")
        '    End If
        'End If
    End Sub

    Private Sub TxtDr_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtDr.KeyPress
        If e.KeyChar = Chr(13) Then
            Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtDr.Text & "' ", RSC)
            If RSC.RecordCount <> 0 Then
                TxtDrNm.Text = Trim(RSC.Fields("Name_L").Value.ToString)
            Else
                MsgBox("ເລກບັນຊີບໍ່ມີໃນລາລະບານ", MsgBoxStyle.Exclamation) : TxtDr.Focus() : Exit Sub
            End If

            TxtCr.Focus()
        End If
    End Sub

    Private Sub TxtDr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtDr.TextChanged

    End Sub

    Private Sub BtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch.Click
        fmShartOfAccDetail.txtSty.Text = "FrmBrokeNew_DR"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        fmShartOfAccDetail.txtSty.Text = "FrmBrokeNew_CR"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
    End Sub

    Private Sub TxtCr_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCr.KeyPress
        If e.KeyChar = Chr(13) Then
            Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtCr.Text & "' ", RSC)
            If RSC.RecordCount <> 0 Then
                TxtCrNm.Text = Trim(RSC.Fields("Name_L").Value.ToString)
            Else
                MsgBox("ເລກບັນຊີບໍ່ມີໃນລາລະບານ", MsgBoxStyle.Exclamation) : TxtCr.Focus() : Exit Sub
            End If

        End If
    End Sub

    Private Sub TxtCr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCr.TextChanged

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        fmShartOfAccDetail.txtSty.Text = "FrmBrokeNew_DR22"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        fmShartOfAccDetail.txtSty.Text = "FrmBrokeNew_DR33"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        fmShartOfAccDetail.txtSty.Text = "FrmBrokeNew_CR22"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
    End Sub

    Private Sub TxtDr22_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtDr22.KeyPress
        If e.KeyChar = Chr(13) Then
            Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtDr22.Text & "' ", RSC)
            If RSC.RecordCount <> 0 Then
                TxtDrNm22.Text = Trim(RSC.Fields("Name_L").Value.ToString)
            Else
                MsgBox("ເລກບັນຊີບໍ່ມີໃນລາລະບານ", MsgBoxStyle.Exclamation) : TxtDr22.Focus() : Exit Sub
            End If

        End If
    End Sub

    Private Sub TxtDr22_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtDr22.TextChanged

    End Sub

    Private Sub TxtDr33_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtDr33.KeyPress
        If e.KeyChar = Chr(13) Then
            Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtDr33.Text & "' ", RSC)
            If RSC.RecordCount <> 0 Then
                TxtDrNm33.Text = Trim(RSC.Fields("Name_L").Value.ToString)
            Else
                MsgBox("ເລກບັນຊີບໍ່ມີໃນລາລະບານ", MsgBoxStyle.Exclamation) : TxtDr33.Focus() : Exit Sub
            End If

        End If
    End Sub

    Private Sub TxtDr33_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtDr33.TextChanged

    End Sub

    Private Sub TxtCr22_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCr22.KeyPress
        If e.KeyChar = Chr(13) Then
            Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtCr22.Text & "' ", RSC)
            If RSC.RecordCount <> 0 Then
                TxtCrNm22.Text = Trim(RSC.Fields("Name_L").Value.ToString)
            Else
                MsgBox("ເລກບັນຊີບໍ່ມີໃນລາລະບານ", MsgBoxStyle.Exclamation) : TxtCr22.Focus() : Exit Sub
            End If

        End If
    End Sub

    Private Sub TxtCr22_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCr22.TextChanged

    End Sub
End Class