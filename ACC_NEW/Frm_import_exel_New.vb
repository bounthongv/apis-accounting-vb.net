Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports ApPBank10.Module

Public Class Frm_import_exel_New
    Dim amt1 As Double
    Dim amt2 As Double
Private Sub Frm_import_exel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Setup DataGridView columns
        FG.Columns.Clear()
        FG.Columns.Add("Col0", "ລ/ດ")
        FG.Columns.Add("Col1", "ວັນທີ")
        FG.Columns.Add("Col2", "ເລກທີ່ບິນ")
        FG.Columns.Add("Col3", "ເລກທີໂອໂຕ")
        FG.Columns.Add("Col4", "ແຊັກ")
        FG.Columns.Add("Col5", "ປື້ມ")
        FG.Columns.Add("Col6", "ເນື້ອໃນພາສາລາວ")
        FG.Columns.Add("Col7", "ເນື້ອໃນພາສາອັງກິດ")
        FG.Columns.Add("Col8", "ໜີ້")
        FG.Columns.Add("Col9", "ມີ")
        FG.Columns.Add("Col10", "ຈຳນວນເງີນໜີ້")
        FG.Columns.Add("Col11", "ຈຳນວນເງີນມີ")
        FG.Columns.Add("Col12", "ສະກຸນເງິນ")
        FG.Columns.Add("Col13", "ອັດຕາແລກປ່ຽນ")
        FG.Columns.Add("Col14", "ຈຳນວນເງີນກີບ")
        FG.Columns.Add("Col15", "ຈຳນວນເງີນໂດລາ")
        FG.Columns.Add("Col16", "ເລກບັນຊີທະນາຄານ")
        FG.Columns.Add("Col17", "ແຫຼງທຶນ")
        FG.Columns.Add("Col18", "ອົງປະກອບ")
        FG.Columns.Add("Col19", "ລະຫັດກິດຈະກຳ")
        FG.Columns.Add("Col20", "ປະເພດລາຍຈ່າຍ")
        FG.Columns.Add("Col21", "ລະຫັດ")
        FG.Columns.Add("Col22", "ສຳນັກງານ")

        ' Set column widths and alignment
        For i As Integer = 0 To FG.Columns.Count - 1
            FG.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        Next

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

        FG.Rows.Clear()
        For i = 0 To DataGridView1.RowCount - 2
            If DataGridView1.Item(1, i).Value IsNot Nothing AndAlso DataGridView1.Item(1, i).Value.ToString <> "" Then
                ' Add new row to DataGridView
                Dim rowIndex As Integer = FG.Rows.Add()
                
                ' Populate row with data from DataGridView1
                FG.Rows(rowIndex).Cells(0).Value = "" ' Will be filled with sequence number later
                FG.Rows(rowIndex).Cells(1).Value = If(DataGridView1.Item(0, i).Value IsNot Nothing, DataGridView1.Item(0, i).Value.ToString, "")
                FG.Rows(rowIndex).Cells(2).Value = If(DataGridView1.Item(1, i).Value IsNot Nothing, DataGridView1.Item(1, i).Value.ToString, "")
                FG.Rows(rowIndex).Cells(3).Value = If(DataGridView1.Item(2, i).Value IsNot Nothing, DataGridView1.Item(2, i).Value.ToString, "")
                FG.Rows(rowIndex).Cells(4).Value = If(DataGridView1.Item(3, i).Value IsNot Nothing, DataGridView1.Item(3, i).Value.ToString, "")
                FG.Rows(rowIndex).Cells(5).Value = If(DataGridView1.Item(4, i).Value IsNot Nothing, DataGridView1.Item(4, i).Value.ToString, "")
                FG.Rows(rowIndex).Cells(6).Value = If(DataGridView1.Item(5, i).Value IsNot Nothing, DataGridView1.Item(5, i).Value.ToString, "")
                FG.Rows(rowIndex).Cells(7).Value = If(DataGridView1.Item(6, i).Value IsNot Nothing, DataGridView1.Item(6, i).Value.ToString, "")
                FG.Rows(rowIndex).Cells(8).Value = If(DataGridView1.Item(7, i).Value IsNot Nothing, DataGridView1.Item(7, i).Value.ToString, "")
                FG.Rows(rowIndex).Cells(9).Value = If(DataGridView1.Item(8, i).Value IsNot Nothing, DataGridView1.Item(8, i).Value.ToString, "")
                FG.Rows(rowIndex).Cells(10).Value = If(DataGridView1.Item(9, i).Value IsNot Nothing, Val(DataGridView1.Item(9, i).Value.ToString), 0)
                FG.Rows(rowIndex).Cells(11).Value = If(DataGridView1.Item(10, i).Value IsNot Nothing, Val(DataGridView1.Item(10, i).Value.ToString), 0)
                FG.Rows(rowIndex).Cells(12).Value = If(DataGridView1.Item(11, i).Value IsNot Nothing, DataGridView1.Item(11, i).Value.ToString, "")
                FG.Rows(rowIndex).Cells(13).Value = If(DataGridView1.Item(12, i).Value IsNot Nothing, Val(DataGridView1.Item(12, i).Value.ToString), 0)
                FG.Rows(rowIndex).Cells(14).Value = If(DataGridView1.Item(13, i).Value IsNot Nothing, Val(DataGridView1.Item(13, i).Value.ToString), 0)
                FG.Rows(rowIndex).Cells(15).Value = If(DataGridView1.Item(14, i).Value IsNot Nothing, Val(DataGridView1.Item(14, i).Value.ToString), 0)
                FG.Rows(rowIndex).Cells(16).Value = If(DataGridView1.Item(15, i).Value IsNot Nothing, DataGridView1.Item(15, i).Value.ToString, "")
                FG.Rows(rowIndex).Cells(17).Value = If(DataGridView1.Item(16, i).Value IsNot Nothing, DataGridView1.Item(16, i).Value.ToString, "")
                FG.Rows(rowIndex).Cells(18).Value = If(DataGridView1.Item(17, i).Value IsNot Nothing, DataGridView1.Item(17, i).Value.ToString, "")
                FG.Rows(rowIndex).Cells(19).Value = If(DataGridView1.Item(18, i).Value IsNot Nothing, DataGridView1.Item(18, i).Value.ToString, "")
                FG.Rows(rowIndex).Cells(20).Value = If(DataGridView1.Item(19, i).Value IsNot Nothing, DataGridView1.Item(19, i).Value.ToString, "")
                FG.Rows(rowIndex).Cells(21).Value = If(DataGridView1.Item(20, i).Value IsNot Nothing, DataGridView1.Item(20, i).Value.ToString, "")
                FG.Rows(rowIndex).Cells(22).Value = If(DataGridView1.Item(21, i).Value IsNot Nothing, DataGridView1.Item(21, i).Value.ToString, "")
                FG.Rows(rowIndex).Cells(22).Value = If(DataGridView1.Item(22, i).Value IsNot Nothing, DataGridView1.Item(22, i).Value.ToString, "")
            End If
        Next i

        ' Process the data
        For i = 0 To FG.Rows.Count - 1
            no = no + 1
            FG.Rows(i).Cells(0).Value = no

            If FG.Rows(i).Cells(3).Value Is Nothing OrElse FG.Rows(i).Cells(3).Value.ToString = "" Then
                Dim col2Value As String = If(FG.Rows(i).Cells(2).Value IsNot Nothing, FG.Rows(i).Cells(2).Value.ToString, "")
                Dim col5Value As String = If(FG.Rows(i).Cells(5).Value IsNot Nothing, FG.Rows(i).Cells(5).Value.ToString, "")
                FG.Rows(i).Cells(3).Value = col2Value & "/" & col5Value
            End If

            ' Format numeric columns
            If FG.Rows(i).Cells(10).Value IsNot Nothing AndAlso FG.Rows(i).Cells(10).Value.ToString <> "" Then
                FG.Rows(i).Cells(10).Value = Format(CDbl(FG.Rows(i).Cells(10).Value), "#,##0.00")
            End If
            If FG.Rows(i).Cells(11).Value IsNot Nothing AndAlso FG.Rows(i).Cells(11).Value.ToString <> "" Then
                FG.Rows(i).Cells(11).Value = Format(CDbl(FG.Rows(i).Cells(11).Value), "#,##0.00")
            End If
            If FG.Rows(i).Cells(13).Value IsNot Nothing AndAlso FG.Rows(i).Cells(13).Value.ToString <> "" Then
                FG.Rows(i).Cells(13).Value = Format(CDbl(FG.Rows(i).Cells(13).Value), "#,##0.00")
            End If
            If FG.Rows(i).Cells(14).Value IsNot Nothing AndAlso FG.Rows(i).Cells(14).Value.ToString <> "" Then
                FG.Rows(i).Cells(14).Value = Format(CDbl(FG.Rows(i).Cells(14).Value), "#,##0.00")
            End If
            If FG.Rows(i).Cells(15).Value IsNot Nothing AndAlso FG.Rows(i).Cells(15).Value.ToString <> "" Then
                FG.Rows(i).Cells(15).Value = Format(CDbl(FG.Rows(i).Cells(15).Value), "#,##0.00")
            End If
        Next
        GroupBox1.Visible = False
        Call amt()
    End Sub
Private Sub amt()
        amt1 = 0
        amt2 = 0
        Dim i As Integer
        For i = 0 To FG.Rows.Count - 1
            If FG.Rows(i).Cells(10).Value IsNot Nothing AndAlso FG.Rows(i).Cells(10).Value.ToString <> "" Then
                amt1 = amt1 + CDbl(FG.Rows(i).Cells(10).Value)
            End If
            If FG.Rows(i).Cells(11).Value IsNot Nothing AndAlso FG.Rows(i).Cells(11).Value.ToString <> "" Then
                amt2 = amt2 + CDbl(FG.Rows(i).Cells(11).Value)
            End If
        Next i

        txtSumAmountDr.Text = Format(CDbl(amt1), "##,##0.00")
        txtSumAmountCr.Text = Format(CDbl(amt2), "##,##0.00")
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

Private Sub FG_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelectionChanged
        ' DataGridView is editable by default, no need to set Editable property
    End Sub

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        Call amt()
        If CDbl(txtSumAmountDr.Text) <> CDbl(txtSumAmountCr.Text) Then MsgBox("ບັນຊີນີ້ບໍ່ທັນບໍ່ທັນດູນດ່ຽງ ບໍ່ສາມາດບັນທຶກໄດ້!", MsgBoxStyle.OkOnly) : Exit Sub
        Call SaveItems()
        Call Insert_Gen_jn()
        MsgBox("ບັນທຶກສຳເລັດ")
    End Sub
Private Sub SaveItems()
        Dim aa As String
        aa = " delete Tmp_Import  "
        ExecuteNonQuery(aa)
        
        Dim dtCheck As DataTable = GetDataTable("Select * FROM Tmp_Import")
        Dim i As Integer

        For i = 0 To FG.Rows.Count - 1
            If FG.Rows(i).Cells(1).Value IsNot Nothing AndAlso FG.Rows(i).Cells(1).Value.ToString <> "" Then
                If dtCheck.Rows.Count = 0 Then
                    Dim sa As String = " INSERT INTO Tmp_Import (date_work,Referno,certify, cheque_no,book, descrip, descripe, code_dr, code_cr,ac_code,  " & _
                       " amount_dr,amount_cr,Curr,Rate, amount_LAK, amount_USD,bank_no, doner, Com_id, Activity_id,Cat_ID, office_id,office_nm,last_update, last_user,pc_nm) " & _
                    " VALUES (  N'" & Format(CDate(If(FG.Rows(i).Cells(1).Value IsNot Nothing, FG.Rows(i).Cells(1).Value.ToString, Now())), "yyyy-MM-dd") & "'," & _
                        " N'" & Apostrophe(If(FG.Rows(i).Cells(2).Value IsNot Nothing, FG.Rows(i).Cells(2).Value.ToString, "")) & "'," & _
                        " N'" & Apostrophe(If(FG.Rows(i).Cells(3).Value IsNot Nothing, FG.Rows(i).Cells(3).Value.ToString, "")) & "'," & _
                         " N'" & Apostrophe(If(FG.Rows(i).Cells(4).Value IsNot Nothing, FG.Rows(i).Cells(4).Value.ToString, "")) & "'," & _
                          " N'" & Trim(Apostrophe(If(FG.Rows(i).Cells(5).Value IsNot Nothing, FG.Rows(i).Cells(5).Value.ToString, ""))) & "'," & _
                           " N'" & Apostrophe(If(FG.Rows(i).Cells(6).Value IsNot Nothing, FG.Rows(i).Cells(6).Value.ToString, "")) & "'," & _
                            " N'" & Apostrophe(If(FG.Rows(i).Cells(7).Value IsNot Nothing, FG.Rows(i).Cells(7).Value.ToString, "")) & "'," & _
                             " N'" & Apostrophe(If(FG.Rows(i).Cells(8).Value IsNot Nothing, FG.Rows(i).Cells(8).Value.ToString, "")) & "'," & _
                              " N'" & Apostrophe(If(FG.Rows(i).Cells(9).Value IsNot Nothing, FG.Rows(i).Cells(9).Value.ToString, "")) & "'," & _
                      " N'" & Apostrophe(If(FG.Rows(i).Cells(8).Value IsNot Nothing, FG.Rows(i).Cells(8).Value.ToString, "") + If(FG.Rows(i).Cells(9).Value IsNot Nothing, FG.Rows(i).Cells(9).Value.ToString, "")) & "'," & _
                      " " & CDbl(If(FG.Rows(i).Cells(10).Value IsNot Nothing AndAlso FG.Rows(i).Cells(10).Value.ToString <> "", FG.Rows(i).Cells(10).Value.ToString, "0")) & ", " & _
                    " " & CDbl(If(FG.Rows(i).Cells(11).Value IsNot Nothing AndAlso FG.Rows(i).Cells(11).Value.ToString <> "", FG.Rows(i).Cells(11).Value.ToString, "0")) & ", " & _
                     " N'" & Apostrophe(If(FG.Rows(i).Cells(12).Value IsNot Nothing, FG.Rows(i).Cells(12).Value.ToString, "")) & "'," & _
                     " " & CDbl(If(FG.Rows(i).Cells(13).Value IsNot Nothing AndAlso FG.Rows(i).Cells(13).Value.ToString <> "", FG.Rows(i).Cells(13).Value.ToString, "0")) & ", " & _
                     " " & CDbl(If(FG.Rows(i).Cells(14).Value IsNot Nothing AndAlso FG.Rows(i).Cells(14).Value.ToString <> "", FG.Rows(i).Cells(14).Value.ToString, "0")) & ", " & _
                     " " & CDbl(If(FG.Rows(i).Cells(15).Value IsNot Nothing AndAlso FG.Rows(i).Cells(15).Value.ToString <> "", FG.Rows(i).Cells(15).Value.ToString, "0")) & ", " & _
                       " N'" & Apostrophe(If(FG.Rows(i).Cells(16).Value IsNot Nothing, FG.Rows(i).Cells(16).Value.ToString, "")) & "'," & _
                      " N'" & Apostrophe(If(FG.Rows(i).Cells(17).Value IsNot Nothing, FG.Rows(i).Cells(17).Value.ToString, "")) & "'," & _
                       " N'" & Apostrophe(If(FG.Rows(i).Cells(18).Value IsNot Nothing, FG.Rows(i).Cells(18).Value.ToString, "")) & "'," & _
                        " N'" & Apostrophe(If(FG.Rows(i).Cells(19).Value IsNot Nothing, FG.Rows(i).Cells(19).Value.ToString, "")) & "'," & _
                         " N'" & Apostrophe(If(FG.Rows(i).Cells(20).Value IsNot Nothing, FG.Rows(i).Cells(20).Value.ToString, "")) & "'," & _
                          " N'" & Apostrophe(If(FG.Rows(i).Cells(21).Value IsNot Nothing, FG.Rows(i).Cells(21).Value.ToString, "")) & "'," & _
                           " N'" & Apostrophe(If(FG.Rows(i).Cells(22).Value IsNot Nothing, FG.Rows(i).Cells(22).Value.ToString, "")) & "'," & _
                               " Getdate()," & _
                    " N'" & Apostrophe(MUserName) & "'," & _
                     " N'" & Apostrophe(MDServerName) & "') "
                    ExecuteNonQuery(sa)
                End If
            End If
        Next i
        
        'Update account codes
        ExecuteNonQuery("update Tmp_Import set Tmp_Import.ac_code=Acc_Code.ac_code from Acc_Code,Tmp_Import where Acc_Code.ac_original=Tmp_Import.ac_code ")
        ExecuteNonQuery("update Tmp_Import set Tmp_Import.code_dr=Acc_Code.ac_code from Acc_Code,Tmp_Import where Acc_Code.ac_original=Tmp_Import.code_dr and Tmp_Import.code_dr<>''  ")
        ExecuteNonQuery("update Tmp_Import set Tmp_Import.code_cr=Acc_Code.ac_code from Acc_Code,Tmp_Import where Acc_Code.ac_original=Tmp_Import.code_cr and Tmp_Import.code_Cr<>'' ")
    End Sub

Private Sub Insert_Gen_jn()
        Frm_import_progress.Show()
        Dim aa As String
        Dim i As Integer
        
        For i = 0 To FG.Rows.Count - 1
            If FG.Rows(i).Cells(3).Value IsNot Nothing AndAlso FG.Rows(i).Cells(3).Value.ToString <> "" Then
                Dim sk As String = "Select * FROM gen_jn where  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'  "
                Dim dtCheck As DataTable = GetDataTable(sk)
                If dtCheck.Rows.Count = 0 Then
                    aa = "   insert into gen_jn (date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, " & _
          "  code_cr,ac_code,ac_name, amount_dr, amount_cr, curr, rate,   bank_no, " & _
          "    don_id, Com_id, Activity_id, Cat_ID,  office_id,lock, my_lock, del ,AG  ,last_update, last_user,pc_nm) " & _
         "   select  date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr," & _
         "  code_cr,ac_code,descrip, amount_dr, amount_cr, curr, rate,   bank_no, " & _
         " doner, Com_id, Activity_id, Cat_id, office_id  ,0, 0, 1,0, Getdate()," & _
              " N'" & Apostrophe(MUserName) & "'," & _
              " N'" & Apostrophe(MDServerName) & "'" & _
"    from   Tmp_Import  where  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'    order by certify  "
                    ExecuteNonQuery(aa)

                    'aa = " update gen_jn set don_id= AP_Donnor.Don_ID  from AP_Donnor where  gen_jn.don_id= AP_Donnor.Don_Sym   "
                    'CNN.Execute(aa)
                    'aa = "  update gen_jn set office_id= AP_Office.off_id  from AP_Office  where  gen_jn.office_id= AP_Office.office_id   "
                    'CNN.Execute(aa)

aa = "   update gen_jn set Curr_i = curr where  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "' " & _
                     " update gen_jn set Rate_i  = Rate where  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "' "
                    ExecuteNonQuery(aa)

                    aa = "   update gen_jn set amount =amount_dr + amount_cr where  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'  " & _
                    "  update gen_jn set amt_dr  =amount_dr where  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'  " & _
                    " update gen_jn set amt_cr  =amount_cr  where  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'  "
                    ExecuteNonQuery(aa)

                    aa = "  update gen_jn set  rate =1 where  rate =0 and  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "' "
                    ExecuteNonQuery(aa)

                    aa = "   update gen_jn set amt_USD_dr  =amount_dr / rate  where  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'   " & _
                   "  update gen_jn set amt_USD_cr  =amount_cr  /rate   where  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'   "
                    ExecuteNonQuery(aa)

                    aa = "       update gen_jn set amount_dr=0 where amount_dr is null and  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'  " & _
                    "   update gen_jn set amount_cr=0 where amount_cr is null and  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'   " & _
                     "  update gen_jn set amt_dr=0 where amt_dr is null and  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'  " & _
                    "   update  gen_jn set amt_cr=0 where amt_cr is null and  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'  " & _
                    "  update gen_jn set amt_USD_dr=0 where amt_USD_dr is null and  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'  " & _
                    "   update  gen_jn set amt_USD_cr=0 where amt_USD_cr is null and  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'   "
ExecuteNonQuery(aa)
                    aa = ""
                    aa = "       insert into AP_ACC_Gen ( date_work, Referno, certify,Book ,   descrip, descripe,amount,net_amt , " & _
                  "  AmountDr, AmountCr,TotalAmountDr,TotalAmountCr,curr,rate,Rate_USD,bank_no,don_id,Com_id,office_id)    " & _
                 "  select     date_work, Referno, certify,Book,    '', '', " & _
                   "  sum(amount_dr) as amount , sum(amount_dr) as net_amt , sum(amount_dr) as amount_dr , " & _
                   "   sum(amount_cr) as amount_cr , sum(amount_dr) as TotalAmountDr , sum(amount_cr) as TotalAmountCr , " & _
                  "   curr,rate,rate,bank_no,don_id,Com_id,  office_id from gen_jn where  1=1    and   certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'  " & _
                   "    group by    date_work, Referno, certify,Book,curr,rate,bank_no, " & _
                   "    don_id, Com_id,  office_id    "
                    ExecuteNonQuery(aa)
                Else

                    aa = "delete gen_jn WHERE certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'  "
                    ExecuteNonQuery(aa)
                    aa = "delete AP_ACC_Gen WHERE certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'  "
                    ExecuteNonQuery(aa)

                    aa = "   insert into gen_jn (date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, " & _
         "  code_cr,ac_code,ac_name, amount_dr, amount_cr, curr, rate,   bank_no, " & _
        "    don_id, Com_id, Activity_id, Cat_ID,  office_id,company, lock, my_lock ,del ,AG , last_update, last_user,pc_nm) " & _
         "   select  date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr," & _
        "  code_cr,ac_code,descrip, amount_dr, amount_cr, curr, rate,   bank_no, " & _
         " doner, Com_id, Activity_id, Cat_id, office_id  , office_id , 0, 0, 1,0, Getdate()," & _
         " N'" & Apostrophe(MUserName) & "'," & _
         " N'" & Apostrophe(MDServerName) & "'" & _
"    from   Tmp_Import  where  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'    order by certify  "
                    ExecuteNonQuery(aa)

                    'aa = " update gen_jn set don_id= AP_Donnor.Don_ID  from AP_Donnor where  gen_jn.don_id= AP_Donnor.Don_Sym   "
                    'CNN.Execute(aa)
                    'aa = "  update gen_jn set office_id= AP_Office.off_id  from AP_Office  where  gen_jn.office_id= AP_Office.office_id   "
                    'CNN.Execute(aa)

aa = "  update gen_jn set Curr_i = curr where  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "' " & _
                     " update gen_jn set Rate_i  = Rate where  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "' "
                    ExecuteNonQuery(aa)


                    aa = "   update gen_jn set amount =amount_dr + amount_cr where  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "' " & _
                    "  update gen_jn set amt_dr  =amount_dr where  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'  " & _
                    " update gen_jn set amt_cr  =amount_cr  where  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'  "
                    ExecuteNonQuery(aa)

aa = "  update gen_jn set  rate =1 where  rate =0 and  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "' "
                    ExecuteNonQuery(aa)

aa = "   update gen_jn set amt_USD_dr  =amount_dr / rate where  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'    " & _
                   "  update gen_jn set amt_USD_cr  =amount_cr  /rate   where  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'   "
                    ExecuteNonQuery(aa)

                    aa = "       update gen_jn set amount_dr=0 where amount_dr is null and  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'  " & _
                    "   update gen_jn set amount_cr=0 where amount_cr is null and  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'   " & _
                     "  update gen_jn set amt_dr=0 where amt_dr is null and  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'  " & _
                    "   update  gen_jn set amt_cr=0 where amt_cr is null and  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'  " & _
                    "  update gen_jn set amt_USD_dr=0 where amt_USD_dr is null and  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'  " & _
                    "   update  gen_jn set amt_USD_cr=0 where amt_USD_cr is null  and  certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'  "
                    ExecuteNonQuery(aa)

aa = "       insert into AP_ACC_Gen ( date_work, Referno, certify,Book ,   descrip, descripe,amount,net_amt , " & _
                  "  AmountDr, AmountCr,TotalAmountDr,TotalAmountCr,curr,rate,Rate_USD,bank_no,don_id,Com_id,office_id)    " & _
                 "  select     date_work, Referno, certify,Book,    '', '', " & _
                   "  sum(amount_dr) as amount , sum(amount_dr) as net_amt , sum(amount_dr) as amount_dr , " & _
                   "   sum(amount_cr) as amount_cr , sum(amount_dr) as TotalAmountDr , sum(amount_cr) as TotalAmountCr , " & _
                  "   curr,rate,rate,bank_no,don_id,Com_id,  office_id from gen_jn where  1=1    and   certify=N'" & FG.Rows(i).Cells(3).Value.ToString & "'  " & _
                   "    group by    date_work, Referno, certify,Book,curr,rate,bank_no, " & _
                   "    don_id, Com_id,  office_id    "
                    ExecuteNonQuery(aa)

End If
                Frm_import_progress.Refresh()

                Frm_import_progress.Label2.Text = FG.Rows.Count

                Frm_import_progress.Label4.Text = If(FG.Rows(i).Cells(0).Value IsNot Nothing, FG.Rows(i).Cells(0).Value.ToString, "")

                Frm_import_progress.Label1.Text = If(FG.Rows(i).Cells(2).Value IsNot Nothing, FG.Rows(i).Cells(2).Value.ToString, "")

            End If
        Next
            Frm_import_progress.Close()
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
        ExecuteNonQuery(aa)
        aa = "   update AP_ACC_Gen set net_usd  =net_amt   where curr='USD'  "
        ExecuteNonQuery(aa)



        aa = "   update AP_ACC_Gen set cheque_no  = gen_jn.cheque_no from gen_jn where  AP_ACC_Gen.certify=gen_jn.certify "
        ExecuteNonQuery(aa)



        aa = "   update AP_ACC_Gen set Com_id =office_id " & _
          "    update gen_jn set Com_id =office_id  " & _
            " update gen_jn set don_id  ='01'  " & _
                "  update AP_ACC_Gen set don_id  ='01' "
        ExecuteNonQuery(aa)


        aa = "   update gen_jn set amt_dr  =  amt_dr * Rate where  amt_dr >0 "
        ExecuteNonQuery(aa)
        aa = "   update gen_jn set amt_cr  =  amt_cr * Rate where  amt_cr >0   "
        ExecuteNonQuery(aa)


    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()
    End Sub
End Class