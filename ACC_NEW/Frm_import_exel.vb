Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class Frm_import_exel
    Dim amt1 As Double
    Dim amt2 As Double
    Private Sub Frm_import_exel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        FG.FormatString = "^ລ/ດ |<ວັນທີ        |<ເລກທີ່ບິນ    |<ເລກທີໂອໂຕ  |<ແຊັກ             |<ປື້ມ     |<ເນື້ອໃນພາສາລາວ     |<ເນື້ອໃນພາສາ ອັງກິດ  |<ໜີ້     |<ມີ       |<ຈຳນວນເງີນໜີ້     |<ຈຳນວນເງີນມີ    |<ສະກຸນເງິນ    |<ອັດຕາແລກປ່ຽນ    |<ຈຳນວນເງີນກີບ   |<ຈຳນວນເງີນໂດລາ    |<ເລກບັນຊີທະນາຄານ|<ແຫຼງທຶນ   |<ອົງປະກອບ|<ລະຫັດກິດຈະກຳ |<ປະເພດລາຍຈ່າຍ|<ລະຫັດ |<ສຳນັກງານ         "


    End Sub

    Private Sub load_exel()
        Dim cn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\EX_test.xlsx ;Extended Properties=Excel 12.0 Xml;")

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
      
        OpenFileDialog1.Filter = "Excel *.xlsx|*.xlsx|All file|*.*"
        OpenFileDialog1.ShowDialog()
        txtFNm.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim i As Integer
        Dim no As Integer = 0

        FG.Rows = 1
        For i = 0 To DataGridView1.RowCount - 2
            If DataGridView1.Item(1, i).Value.ToString <> "" Then

                'FG.AddItem(DataGridView1.Item(0, i).Value.ToString & Chr(9) & DataGridView1.Item(1, i).Value.ToString)

                FG.AddItem(DataGridView1.Item(0, i).Value.ToString & Chr(9) & DataGridView1.Item(1, i).Value.ToString & Chr(9) & DataGridView1.Item(2, i).Value.ToString & _
             Chr(9) & Trim(DataGridView1.Item(3, i).Value.ToString) & Chr(9) & Trim(DataGridView1.Item(4, i).Value.ToString) & Chr(9) & Trim(DataGridView1.Item(5, i).Value.ToString) & _
              Chr(9) & Trim(DataGridView1.Item(6, i).Value.ToString) & _
                Chr(9) & Trim(DataGridView1.Item(7, i).Value.ToString) & _
                  Chr(9) & Trim(DataGridView1.Item(8, i).Value.ToString) & _
                 Chr(9) & Trim(DataGridView1.Item(9, i).Value.ToString) & _
                  Chr(9) & (DataGridView1.Item(10, i).Value.ToString) & _
                  Chr(9) & (DataGridView1.Item(11, i).Value.ToString) & _
                    Chr(9) & (DataGridView1.Item(12, i).Value.ToString) & _
                      Chr(9) & (DataGridView1.Item(13, i).Value.ToString) & _
                        Chr(9) & (DataGridView1.Item(14, i).Value.ToString) & _
                          Chr(9) & (DataGridView1.Item(15, i).Value.ToString) & _
                            Chr(9) & (DataGridView1.Item(16, i).Value.ToString) & _
                              Chr(9) & (DataGridView1.Item(17, i).Value.ToString) & _
                                Chr(9) & (DataGridView1.Item(18, i).Value.ToString) & _
                                  Chr(9) & (DataGridView1.Item(19, i).Value.ToString) & _
                                    Chr(9) & (DataGridView1.Item(20, i).Value.ToString) & _
                                      Chr(9) & (DataGridView1.Item(21, i).Value.ToString) & _
                Chr(9) & DataGridView1.Item(22, i).Value.ToString)
                'FG.set_TextMatrix(FG.Row, no, no)
            End If

        Next i

        For i = 1 To FG.Rows - 1
            no = no + 1
            FG.set_TextMatrix(i, 0, no)

            If FG.get_TextMatrix(i, 3) = "" Then
                FG.set_TextMatrix(i, 3, FG.get_TextMatrix(i, 2) & "/" & FG.get_TextMatrix(i, 5))
            End If

            FG.set_TextMatrix(i, 10, Format(CDbl(FG.get_TextMatrix(i, 10)), "#,##0.00"))
            FG.set_TextMatrix(i, 11, Format(CDbl(FG.get_TextMatrix(i, 11)), "#,##0.00"))
            FG.set_TextMatrix(i, 13, Format(CDbl(FG.get_TextMatrix(i, 13)), "#,##0.00"))
            FG.set_TextMatrix(i, 14, Format(CDbl(FG.get_TextMatrix(i, 14)), "#,##0.00"))
            FG.set_TextMatrix(i, 15, Format(CDbl(FG.get_TextMatrix(i, 15)), "#,##0.00"))

            'FG.set_TextMatrix(i, no)

        Next
        'Call Calc()
        GroupBox1.Visible = False
        Call amt()
    End Sub
    Private Sub amt()
        amt1 = 0
        amt2 = 0
        Dim i As Integer
        For i = 1 To FG.Rows - 1
            amt1 = amt1 + CDbl(FG.get_TextMatrix(i, 10))
            amt2 = amt2 + CDbl(FG.get_TextMatrix(i, 11))
        Next i

        txtSumAmountDr.Text = Format(CDbl(amt1), "##,##0.00")
        txtSumAmountCr.Text = Format(CDbl(amt1), "##,##0.00")
    End Sub
    
    Private Sub btnExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        GroupBox1.Visible = True
        DataGridView1.DataSource = Nothing
        DataGridView1.RefreshEdit()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim mXLS As String
        Dim MyConnection As System.Data.OleDb.OleDbConnection
        Dim DtSet As System.Data.DataSet
        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
        mXLS = Trim(txtFNm.Text)
        If Len(mXLS) = 0 Then Exit Sub
        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='D:\Software\Exim\PL.xls';Extended Properties=Excel 8.0;")
        MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & mXLS & "';Extended Properties=Excel 8.0;")
        '========= Exell 2010 ລົງມາ ========================
        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & mXLS & "';Extended Properties=Excel 12.0 Xml;")
        '=========  ========================

        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.16.0;Data Source='" & mXLS & "';Extended Properties=Excel 16.0 xml;")

        MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [" & txtSheet.Text & "$]", MyConnection)
        MyCommand.TableMappings.Add("Table", "Net-informations.com")
        DtSet = New System.Data.DataSet
        MyCommand.Fill(DtSet)
        DataGridView1.DataSource = DtSet.Tables(0)
        MyConnection.Close()

        'With cn1
        '    .Provider = "Microsoft.ACE.OLEDB.16.0"
        '    .ConnectionString = "Data Source=" & strfile & ";" & _
        '    "Extended Properties=""Excel 16.0 xml;HDR=No;IMEX=1;Readonly=True"""
        'End With

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        DataGridView1.DataSource = Nothing
        DataGridView1.RefreshEdit()
        GroupBox1.Visible = False
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        DataGridView1.DataSource = Nothing
        DataGridView1.RefreshEdit()
    End Sub

    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelChange
        FG.Editable = VSFlex8U.EditableSettings.flexEDKbd
    End Sub

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        Call amt()
        If CDbl(txtSumAmountDr.Text) <> CDbl(txtSumAmountCr.Text) Then MsgBox("ບັນຊີນີ້ບໍ່ທັນບໍ່ທັນດູນດ່ຽງ ບໍ່ສາມາດບັນທຶກໄດ້!", MsgBoxStyle.OkOnly) : Exit Sub
        Call SaveItems()
        '  Call Insert_Gen_jn()
        Insert_Gen_jn22()
        MsgBox("ບັນທຶກສຳເລັດ")
    End Sub
    Private Sub SaveItems()
        Dim aa As String
        aa = " delete Tmp_Import  "
        CNN.Execute(aa)
        Dim Rschk As New ADODB.Recordset
        Dim i As Integer
        With Rschk
            Dim sk As String = "Select * FROM Tmp_Import  "
            Call LoadSqlData(sk, Rschk)

            For i = 1 To FG.Rows - 1

                If Rschk.RecordCount = 0 Then
                    Dim sa As String = " INSERT INTO Tmp_Import (date_work,Referno,certify, cheque_no,book, descrip, descripe, code_dr, code_cr,ac_code,  " & _
                       " amount_dr,amount_cr,Curr,Rate, amount_LAK, amount_USD,bank_no, doner, Com_id, Activity_id,Cat_ID, office_id,office_nm,last_update, last_user,pc_nm) " & _
                    " VALUES (  N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'," & _
                        " N'" & Apostrophe(FG.get_TextMatrix(i, 2)) & "'," & _
                        " N'" & Apostrophe(FG.get_TextMatrix(i, 3)) & "'," & _
                         " N'" & Apostrophe(FG.get_TextMatrix(i, 4)) & "'," & _
                          " N'" & Trim(Apostrophe(FG.get_TextMatrix(i, 5))) & "'," & _
                           " N'" & Apostrophe(FG.get_TextMatrix(i, 6)) & "'," & _
                            " N'" & Apostrophe(FG.get_TextMatrix(i, 7)) & "'," & _
                             " N'" & Apostrophe(FG.get_TextMatrix(i, 8)) & "'," & _
                              " N'" & Apostrophe(FG.get_TextMatrix(i, 9)) & "'," & _
                      " N'" & Apostrophe(FG.get_TextMatrix(i, 8)) + Apostrophe(FG.get_TextMatrix(i, 9)) & "'," & _
                      " " & CDbl(FG.get_TextMatrix(i, 10)) & ", " & _
                    " " & CDbl(FG.get_TextMatrix(i, 11)) & ", " & _
                     " N'" & Apostrophe(FG.get_TextMatrix(i, 12)) & "'," & _
                     " " & CDbl(FG.get_TextMatrix(i, 13)) & ", " & _
                     " " & CDbl(FG.get_TextMatrix(i, 14)) & ", " & _
                     " " & CDbl(FG.get_TextMatrix(i, 15)) & ", " & _
                       " N'" & Apostrophe(FG.get_TextMatrix(i, 16)) & "'," & _
                      " N'" & Apostrophe(FG.get_TextMatrix(i, 17)) & "'," & _
                       " N'" & Apostrophe(FG.get_TextMatrix(i, 18)) & "'," & _
                        " N'" & Apostrophe(FG.get_TextMatrix(i, 19)) & "'," & _
                         " N'" & Apostrophe(FG.get_TextMatrix(i, 20)) & "'," & _
                          " N'" & Apostrophe(FG.get_TextMatrix(i, 21)) & "'," & _
                           " N'" & Apostrophe(FG.get_TextMatrix(i, 22)) & "'," & _
                               " Getdate()," & _
                    " N'" & Apostrophe(MUserName) & "'," & _
                     " N'" & Apostrophe(MDServerName) & "') "
                    CNN.Execute(sa)
                Else

                    'Dim sa As String = "DELETE FROM gen_jn WHERE 1=1 AND certify='" & txtBill_no.Text & "' "
                    'CNN.Execute(sa)

                End If
            Next i
        End With
        ''CNN.Execute("delete gen_jn WHERE ac_code='' ")
        'CNN.Execute("update Tmp_Import set Tmp_Import.ac_code=Acc_Code.ac_code from Acc_Code,Tmp_Import where Acc_Code.ac_original=Tmp_Import.ac_code ")
        'CNN.Execute("update Tmp_Import set Tmp_Import.code_dr=Acc_Code.ac_code from Acc_Code,Tmp_Import where Acc_Code.ac_original=Tmp_Import.code_dr and Tmp_Import.code_dr<>''  ")
        'CNN.Execute("update Tmp_Import set Tmp_Import.code_cr=Acc_Code.ac_code from Acc_Code,Tmp_Import where Acc_Code.ac_original=Tmp_Import.code_cr and Tmp_Import.code_Cr<>'' ")
    End Sub

    Private Sub Insert_Gen_jn()
        Frm_import_progress.Show()
        Dim aa As String
        Dim RSC As New ADODB.Recordset
        Dim Rschk As New ADODB.Recordset
        Dim i As Integer
        With Rschk
            For i = 1 To FG.Rows - 1
                Dim sk As String = "Select certify FROM gen_jn where  certify=N'" & FG.get_TextMatrix(i, 3) & "'  "
                Call LoadSqlData(sk, Rschk)
                If Rschk.RecordCount = 0 Then
                    aa = "   insert into gen_jn (date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, " & _
          "  code_cr,ac_code,ac_name, amount_dr, amount_cr, curr, rate,   bank_no, " & _
          "    don_id, Com_id, Activity_id, Cat_ID,  office_id,lock, my_lock, del ,AG  ,last_update, last_user,pc_nm) " & _
         "   select  date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr," & _
         "  code_cr,ac_code,descrip, amount_dr, amount_cr, curr, rate,   bank_no, " & _
         " doner, Com_id, Activity_id, Cat_id, office_id  ,0, 0, 1,0, Getdate()," & _
              " N'" & Apostrophe(MUserName) & "'," & _
              " N'" & Apostrophe(MDServerName) & "'" & _
       "    from   Tmp_Import  where  certify=N'" & FG.get_TextMatrix(i, 3) & "'    order by certify  "
                    CNN.Execute(aa)

                    'aa = " update gen_jn set don_id= AP_Donnor.Don_ID  from AP_Donnor where  gen_jn.don_id= AP_Donnor.Don_Sym   "
                    'CNN.Execute(aa)
                    'aa = "  update gen_jn set office_id= AP_Office.off_id  from AP_Office  where  gen_jn.office_id= AP_Office.office_id   "
                    'CNN.Execute(aa)

                    aa = "  update gen_jn set Curr_i = curr where  certify=N'" & FG.get_TextMatrix(i, 3) & "' " & _
                     " update gen_jn set Rate_i  = Rate where  certify=N'" & FG.get_TextMatrix(i, 3) & "' "
                    CNN.Execute(aa)

                    aa = "   update gen_jn set amount =amount_dr + amount_cr where  certify=N'" & FG.get_TextMatrix(i, 3) & "'  " & _
                    "  update gen_jn set amt_dr  =amount_dr where  certify=N'" & FG.get_TextMatrix(i, 3) & "'  " & _
                    " update gen_jn set amt_cr  =amount_cr  where  certify=N'" & FG.get_TextMatrix(i, 3) & "'  "
                    CNN.Execute(aa)

                    aa = "  update gen_jn set  rate =1 where  rate =0 and  certify=N'" & FG.get_TextMatrix(i, 3) & "' "
                    CNN.Execute(aa)

                    aa = "   update gen_jn set amt_USD_dr  =amount_dr / rate  where  certify=N'" & FG.get_TextMatrix(i, 3) & "'   " & _
                   "  update gen_jn set amt_USD_cr  =amount_cr  /rate   where  certify=N'" & FG.get_TextMatrix(i, 3) & "'   "
                    CNN.Execute(aa)

                    aa = "       update gen_jn set amount_dr=0 where amount_dr is null and  certify=N'" & FG.get_TextMatrix(i, 3) & "'  " & _
                    "   update gen_jn set amount_cr=0 where amount_cr is null and  certify=N'" & FG.get_TextMatrix(i, 3) & "'   " & _
                     "  update gen_jn set amt_dr=0 where amt_dr is null and  certify=N'" & FG.get_TextMatrix(i, 3) & "'  " & _
                    "   update  gen_jn set amt_cr=0 where amt_cr is null and  certify=N'" & FG.get_TextMatrix(i, 3) & "'  " & _
                    "  update gen_jn set amt_USD_dr=0 where amt_USD_dr is null and  certify=N'" & FG.get_TextMatrix(i, 3) & "'  " & _
                    "   update  gen_jn set amt_USD_cr=0 where amt_USD_cr is null and  certify=N'" & FG.get_TextMatrix(i, 3) & "'   "
                    CNN.Execute(aa)
                    aa = ""
                    aa = "       insert into AP_ACC_Gen ( date_work, Referno, certify,Book ,   descrip, descripe,amount,net_amt , " & _
                  "  AmountDr, AmountCr,TotalAmountDr,TotalAmountCr,curr,rate,Rate_USD,bank_no,don_id,Com_id,office_id)    " & _
                 "  select     date_work, Referno, certify,Book,    '', '', " & _
                   "  sum(amount_dr) as amount , sum(amount_dr) as net_amt , sum(amount_dr) as amount_dr , " & _
                   "   sum(amount_cr) as amount_cr , sum(amount_dr) as TotalAmountDr , sum(amount_cr) as TotalAmountCr , " & _
                  "   curr,rate,rate,bank_no,don_id,Com_id,  office_id from gen_jn where  1=1    and   certify=N'" & FG.get_TextMatrix(i, 3) & "'  " & _
                   "    group by    date_work, Referno, certify,Book,curr,rate,bank_no, " & _
                   "    don_id, Com_id,  office_id    "
                    CNN.Execute(aa)
                Else

                    aa = "delete gen_jn WHERE certify=N'" & FG.get_TextMatrix(i, 3) & "'  "
                    CNN.Execute(aa)
                    aa = "delete AP_ACC_Gen WHERE certify=N'" & FG.get_TextMatrix(i, 3) & "'  "
                    CNN.Execute(aa)

                    aa = "   insert into gen_jn (date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, " & _
         "  code_cr,ac_code,ac_name, amount_dr, amount_cr, curr, rate,   bank_no, " & _
        "    don_id, Com_id, Activity_id, Cat_ID,  office_id,company, lock, my_lock ,del ,AG , last_update, last_user,pc_nm) " & _
         "   select  date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr," & _
        "  code_cr,ac_code,descrip, amount_dr, amount_cr, curr, rate,   bank_no, " & _
         " doner, Com_id, Activity_id, Cat_id, office_id  , office_id , 0, 0, 1,0, Getdate()," & _
         " N'" & Apostrophe(MUserName) & "'," & _
         " N'" & Apostrophe(MDServerName) & "'" & _
         "    from   Tmp_Import  where  certify=N'" & FG.get_TextMatrix(i, 3) & "'    order by certify  "
                    CNN.Execute(aa)

                    'aa = " update gen_jn set don_id= AP_Donnor.Don_ID  from AP_Donnor where  gen_jn.don_id= AP_Donnor.Don_Sym   "
                    'CNN.Execute(aa)
                    'aa = "  update gen_jn set office_id= AP_Office.off_id  from AP_Office  where  gen_jn.office_id= AP_Office.office_id   "
                    'CNN.Execute(aa)

                    aa = "  update gen_jn set Curr_i = curr where  certify=N'" & FG.get_TextMatrix(i, 3) & "' " & _
                     " update gen_jn set Rate_i  = Rate where  certify=N'" & FG.get_TextMatrix(i, 3) & "' "
                    CNN.Execute(aa)


                    aa = "   update gen_jn set amount =amount_dr + amount_cr where  certify=N'" & FG.get_TextMatrix(i, 3) & "' " & _
                    "  update gen_jn set amt_dr  =amount_dr where  certify=N'" & FG.get_TextMatrix(i, 3) & "'  " & _
                    " update gen_jn set amt_cr  =amount_cr  where  certify=N'" & FG.get_TextMatrix(i, 3) & "'  "
                    CNN.Execute(aa)

                    aa = "  update gen_jn set  rate =1 where  rate =0 and  certify=N'" & FG.get_TextMatrix(i, 3) & "' "
                    CNN.Execute(aa)

                    aa = "   update gen_jn set amt_USD_dr  =amount_dr / rate where  certify=N'" & FG.get_TextMatrix(i, 3) & "'    " & _
                   "  update gen_jn set amt_USD_cr  =amount_cr  /rate   where  certify=N'" & FG.get_TextMatrix(i, 3) & "'   "
                    CNN.Execute(aa)

                    aa = "       update gen_jn set amount_dr=0 where amount_dr is null and  certify=N'" & FG.get_TextMatrix(i, 3) & "'  " & _
                    "   update gen_jn set amount_cr=0 where amount_cr is null and  certify=N'" & FG.get_TextMatrix(i, 3) & "'   " & _
                     "  update gen_jn set amt_dr=0 where amt_dr is null and  certify=N'" & FG.get_TextMatrix(i, 3) & "'  " & _
                    "   update  gen_jn set amt_cr=0 where amt_cr is null and  certify=N'" & FG.get_TextMatrix(i, 3) & "'  " & _
                    "  update gen_jn set amt_USD_dr=0 where amt_USD_dr is null and  certify=N'" & FG.get_TextMatrix(i, 3) & "'  " & _
                    "   update  gen_jn set amt_USD_cr=0 where amt_USD_cr is null  and  certify=N'" & FG.get_TextMatrix(i, 3) & "'  "
                    CNN.Execute(aa)

                    aa = "       insert into AP_ACC_Gen ( date_work, Referno, certify,Book ,   descrip, descripe,amount,net_amt , " & _
                  "  AmountDr, AmountCr,TotalAmountDr,TotalAmountCr,curr,rate,Rate_USD,bank_no,don_id,Com_id,office_id)    " & _
                 "  select     date_work, Referno, certify,Book,    '', '', " & _
                   "  sum(amount_dr) as amount , sum(amount_dr) as net_amt , sum(amount_dr) as amount_dr , " & _
                   "   sum(amount_cr) as amount_cr , sum(amount_dr) as TotalAmountDr , sum(amount_cr) as TotalAmountCr , " & _
                  "   curr,rate,rate,bank_no,don_id,Com_id,  office_id from gen_jn where  1=1    and   certify=N'" & FG.get_TextMatrix(i, 3) & "'  " & _
                   "    group by    date_work, Referno, certify,Book,curr,rate,bank_no, " & _
                   "    don_id, Com_id,  office_id    "
                    CNN.Execute(aa)

                End If
                Frm_import_progress.Refresh()

                Frm_import_progress.Label2.Text = FG.Rows

                Frm_import_progress.Label4.Text = FG.get_TextMatrix(i, 0)

                Frm_import_progress.Label1.Text = FG.get_TextMatrix(i, 2)

            Next
            Frm_import_progress.Close()
        End With
        'Fg1.get_TextMatrix(Fg1.Row, 1)
        'With RSC
        '    aa = "Select * FROM Tmp_Import  "
        '    Call LoadRs(aa, RSC)
        '    If .RecordCount <> 0 Then
        '        While Not .EOF()

        '        End While
        '        .MoveNext()
        '    End If

        'End With
        'aa = "  delete  Tmp_Import  "
        'CNN.Execute(aa)
        aa = "   update AP_ACC_Gen set net_usd  =net_amt / rate   where curr='LAK'  "
        CNN.Execute(aa)
        aa = "   update AP_ACC_Gen set net_usd  =net_amt   where curr='USD'  "
        CNN.Execute(aa)



        aa = "   update AP_ACC_Gen set cheque_no  = gen_jn.cheque_no from gen_jn where  AP_ACC_Gen.certify=gen_jn.certify "
        CNN.Execute(aa)



        aa = "   update AP_ACC_Gen set Com_id =office_id " & _
          "    update gen_jn set Com_id =office_id  " & _
            " update gen_jn set don_id  ='01'  " & _
                "  update AP_ACC_Gen set don_id  ='01' "
        CNN.Execute(aa)


        aa = "   update gen_jn set amt_dr  =  amt_dr * Rate where  amt_dr >0 "
        CNN.Execute(aa)
        aa = "   update gen_jn set amt_cr  =  amt_cr * Rate where  amt_cr >0   "
        CNN.Execute(aa)


    End Sub
    Private Sub Insert_Gen_jn22()
        'Frm_import_progress.Show()
        Dim aa As String
        Dim RSC As New ADODB.Recordset
        Dim Rschk As New ADODB.Recordset
        Dim i As Integer

        aa = "   insert into gen_jn (date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, " & _
"  code_cr,ac_code,ac_name, amount_dr, amount_cr, curr, rate,   bank_no, " & _
"    don_id, Com_id, Activity_id, Cat_ID,  office_id,lock, my_lock, del ,AG  ,last_update, last_user,pc_nm) " & _
"   select  date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr," & _
"  code_cr,ac_code,descrip, amount_dr, amount_cr, curr, rate,   bank_no, " & _
" doner, Com_id, Activity_id, Cat_id, office_id  ,0, 0, 1,0, Getdate()," & _
  " N'" & Apostrophe(MUserName) & "'," & _
  " N'" & Apostrophe(MDServerName) & "'" & _
"    from   Tmp_Import    order by certify  "
            CNN.Execute(aa)

        aa = "  update gen_jn set Curr_i = curr  " & _
         " update gen_jn set Rate_i  = Rate "
            CNN.Execute(aa)

        aa = "   update gen_jn set amount =amount_dr + amount_cr  " & _
        "  update gen_jn set amt_dr  =amount_dr  " & _
        " update gen_jn set amt_cr  =amount_cr   "
            CNN.Execute(aa)

        aa = "  update gen_jn set  rate =1 where  rate =0  "
            CNN.Execute(aa)

        aa = "   update gen_jn set amt_USD_dr  =amount_dr / rate  " & _
       "  update gen_jn set amt_USD_cr  =amount_cr  /rate      "
            CNN.Execute(aa)

        aa = "       update gen_jn set amount_dr=0 where amount_dr is null  " & _
        "   update gen_jn set amount_cr=0 where amount_cr is null " & _
         "  update gen_jn set amt_dr=0 where amt_dr is null  " & _
        "   update  gen_jn set amt_cr=0 where amt_cr is null   " & _
        "  update gen_jn set amt_USD_dr=0 where amt_USD_dr is null  " & _
        "   update  gen_jn set amt_USD_cr=0 where amt_USD_cr is null   "
            CNN.Execute(aa)
            aa = ""
        aa = "       insert into AP_ACC_Gen ( date_work, Referno, certify,Book ,   descrip, descripe,amount,net_amt , " & _
      "  AmountDr, AmountCr,TotalAmountDr,TotalAmountCr,curr,rate,Rate_USD,bank_no,don_id,Com_id,office_id)    " & _
     "  select     date_work, Referno, certify,Book,    '', '', " & _
       "  sum(amount_dr) as amount , sum(amount_dr) as net_amt , sum(amount_dr) as amount_dr , " & _
       "   sum(amount_cr) as amount_cr , sum(amount_dr) as TotalAmountDr , sum(amount_cr) as TotalAmountCr , " & _
      "   curr,rate,rate,bank_no,don_id,Com_id,  office_id from gen_jn  " & _
       "    group by    date_work, Referno, certify,Book,curr,rate,bank_no, " & _
       "    don_id, Com_id,  office_id    "
            CNN.Execute(aa)

            aa = "   update AP_ACC_Gen set net_usd  =net_amt / rate   where curr='LAK'  "
            CNN.Execute(aa)
            aa = "   update AP_ACC_Gen set net_usd  =net_amt   where curr='USD'  "
            CNN.Execute(aa)

            aa = "   update AP_ACC_Gen set cheque_no  = gen_jn.cheque_no from gen_jn where  AP_ACC_Gen.certify=gen_jn.certify "
            CNN.Execute(aa)

            aa = "   update AP_ACC_Gen set Com_id =office_id " & _
              "    update gen_jn set Com_id =office_id  " & _
                " update gen_jn set don_id  ='01'  " & _
                    "  update AP_ACC_Gen set don_id  ='01' "
            CNN.Execute(aa)
        CNN.Execute(" update gen_jn set company =office_id  ")

            aa = "   update gen_jn set amt_dr  =  amt_dr * Rate where  amt_dr >0 "
            CNN.Execute(aa)
            aa = "   update gen_jn set amt_cr  =  amt_cr * Rate where  amt_cr >0   "
            CNN.Execute(aa)


    End Sub
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()
    End Sub
End Class