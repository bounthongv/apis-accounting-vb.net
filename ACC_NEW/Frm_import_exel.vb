Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports ApPBank10.Module

Public Class Frm_import_exel
    Dim amt1 As Double
    Dim amt2 As Double
    Private Sub Frm_import_exel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Initialize DataGridView for migrated FG functionality
        InitializeDataGridView()
    End Sub
    
Private Sub InitializeDataGridView()
        ' Configure DataGridView (already created in designer) 
        With DataGridView
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .ReadOnly = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .MultiSelect = False
            .RowHeadersWidth = 50
            .ColumnHeadersHeight = 25
        End With
        
        ' Define columns matching original FG FormatString
        ' 0: ລ/ດ, 1: ວັນທີ, 2: ເລກທີ່ບິນ, 3: ເລກທີໂອໂຕ, 4: ແຊັກ, 5: ປື້ມ, 6: ເນື້ອໃນພາສາລາວ, 7: ເນື້ອໃນພາສາອັງກິດ, 8: ໜີ້, 9: ມີ, 10: ຈຳນວນເງີນໜີ້, 11: ຈຳນວນເງີນມີ, 12: ສະກຸນເງິນ, 13: ອັດຕາແລກປ່ຽນ, 14: ຈຳນວນເງີນກີບ, 15: ຈຳນວນເງີນໂດລາ, 16: ເລກບັນຊີທະນາຄານ, 17: ແຫຼງທຶນ, 18: ອົງປະກອບ, 19: ລະຫັດກິດຈະກຳ, 20: ປະເພດລາຍຈ່າຍ, 21: ລະຫັດ, 22: ສຳນັກງານ
        
        Dim columns As String() = {"ລ/ດ", "ວັນທີ", "ເລກທີ່ບິນ", "ເລກທີໂອໂຕ", "ແຊັກ", "ປື້ມ", "ເນື້ອໃນພາສາລາວ", "ເນື້ອໃນພາສາອັງກິດ", "ໜີ້", "ມີ", "ຈຳນວນເງີນໜີ້", "ຈຳນວນເງີນມີ", "ສະກຸນເງິນ", "ອັດຕາແລກປ່ຽນ", "ຈຳນວນເງີນກີບ", "ຈຳນວນເງີນໂດລາ", "ເລກບັນຊີທະນາຄານ", "ແຫຼງທຶນ", "ອົງປະກອບ", "ລະຫັດກິດຈະກຳ", "ປະເພດລາຍຈ່າຍ", "ລະຫັດ", "ສຳນັກງານ"}
        
        ' Clear existing columns and add new ones
        DataGridView.Columns.Clear()
        For i As Integer = 0 To columns.Length - 1
            Dim column As New DataGridViewTextBoxColumn()
            column.Name = "Col" & i
            column.HeaderText = columns(i)
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            DataGridView.Columns.Add(column)
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

        ' Clear DataGridView (replaces FG.Rows = 1)
        DataGridView.Rows.Clear()
        
        For i = 0 To DataGridView1.RowCount - 2
            If DataGridView1.Item(1, i).Value IsNot Nothing AndAlso DataGridView1.Item(1, i).Value.ToString <> "" Then
                ' Add row to DataGridView (replaces FG.AddItem)
                Dim rowIndex As Integer = DataGridView.Rows.Add()
                
                ' Map Excel columns to DataGridView columns
                DataGridView.Rows(rowIndex).Cells(0).Value = "" ' ລ/ດ - will be filled later
                DataGridView.Rows(rowIndex).Cells(1).Value = If(DataGridView1.Item(0, i).Value IsNot Nothing, DataGridView1.Item(0, i).Value.ToString, "")
                DataGridView.Rows(rowIndex).Cells(2).Value = If(DataGridView1.Item(1, i).Value IsNot Nothing, DataGridView1.Item(1, i).Value.ToString, "")
                DataGridView.Rows(rowIndex).Cells(3).Value = If(DataGridView1.Item(2, i).Value IsNot Nothing, DataGridView1.Item(2, i).Value.ToString, "")
                DataGridView.Rows(rowIndex).Cells(4).Value = If(DataGridView1.Item(3, i).Value IsNot Nothing, Trim(DataGridView1.Item(3, i).Value.ToString), "")
                DataGridView.Rows(rowIndex).Cells(5).Value = If(DataGridView1.Item(4, i).Value IsNot Nothing, Trim(DataGridView1.Item(4, i).Value.ToString), "")
                DataGridView.Rows(rowIndex).Cells(6).Value = If(DataGridView1.Item(5, i).Value IsNot Nothing, Trim(DataGridView1.Item(5, i).Value.ToString), "")
                DataGridView.Rows(rowIndex).Cells(7).Value = If(DataGridView1.Item(6, i).Value IsNot Nothing, Trim(DataGridView1.Item(6, i).Value.ToString), "")
                DataGridView.Rows(rowIndex).Cells(8).Value = If(DataGridView1.Item(7, i).Value IsNot Nothing, Trim(DataGridView1.Item(7, i).Value.ToString), "")
                DataGridView.Rows(rowIndex).Cells(9).Value = If(DataGridView1.Item(8, i).Value IsNot Nothing, Trim(DataGridView1.Item(8, i).Value.ToString), "")
                DataGridView.Rows(rowIndex).Cells(10).Value = If(DataGridView1.Item(9, i).Value IsNot Nothing, Trim(DataGridView1.Item(9, i).Value.ToString), "")
                DataGridView.Rows(rowIndex).Cells(11).Value = If(DataGridView1.Item(10, i).Value IsNot Nothing, DataGridView1.Item(10, i).Value.ToString, "0")
                DataGridView.Rows(rowIndex).Cells(12).Value = If(DataGridView1.Item(11, i).Value IsNot Nothing, DataGridView1.Item(11, i).Value.ToString, "0")
                DataGridView.Rows(rowIndex).Cells(13).Value = If(DataGridView1.Item(12, i).Value IsNot Nothing, DataGridView1.Item(12, i).Value.ToString, "")
                DataGridView.Rows(rowIndex).Cells(14).Value = If(DataGridView1.Item(13, i).Value IsNot Nothing, DataGridView1.Item(13, i).Value.ToString, "1")
                DataGridView.Rows(rowIndex).Cells(15).Value = If(DataGridView1.Item(14, i).Value IsNot Nothing, DataGridView1.Item(14, i).Value.ToString, "0")
                DataGridView.Rows(rowIndex).Cells(16).Value = If(DataGridView1.Item(15, i).Value IsNot Nothing, DataGridView1.Item(15, i).Value.ToString, "0")
                DataGridView.Rows(rowIndex).Cells(17).Value = If(DataGridView1.Item(16, i).Value IsNot Nothing, DataGridView1.Item(16, i).Value.ToString, "")
                DataGridView.Rows(rowIndex).Cells(18).Value = If(DataGridView1.Item(17, i).Value IsNot Nothing, DataGridView1.Item(17, i).Value.ToString, "")
                DataGridView.Rows(rowIndex).Cells(19).Value = If(DataGridView1.Item(18, i).Value IsNot Nothing, DataGridView1.Item(18, i).Value.ToString, "")
                DataGridView.Rows(rowIndex).Cells(20).Value = If(DataGridView1.Item(19, i).Value IsNot Nothing, DataGridView1.Item(19, i).Value.ToString, "")
                DataGridView.Rows(rowIndex).Cells(21).Value = If(DataGridView1.Item(20, i).Value IsNot Nothing, DataGridView1.Item(20, i).Value.ToString, "")
                DataGridView.Rows(rowIndex).Cells(22).Value = If(DataGridView1.Item(21, i).Value IsNot Nothing, DataGridView1.Item(21, i).Value.ToString, "")
                DataGridView.Rows(rowIndex).Cells(23).Value = If(DataGridView1.Item(22, i).Value IsNot Nothing, DataGridView1.Item(22, i).Value.ToString, "")
            End If
        Next i

        ' Process rows (replaces FG.get_TextMatrix loops)
        For i = 0 To DataGridView.Rows.Count - 1
            no = no + 1
            DataGridView.Rows(i).Cells(0).Value = no.ToString() ' Set ລ/ດ

            ' Set certify if empty
            If DataGridView.Rows(i).Cells(3).Value Is Nothing OrElse DataGridView.Rows(i).Cells(3).Value.ToString = "" Then
                Dim certifyValue As String = ""
                If DataGridView.Rows(i).Cells(2).Value IsNot Nothing Then certifyValue &= DataGridView.Rows(i).Cells(2).Value.ToString
                certifyValue &= "/"
                If DataGridView.Rows(i).Cells(5).Value IsNot Nothing Then certifyValue &= DataGridView.Rows(i).Cells(5).Value.ToString
                DataGridView.Rows(i).Cells(3).Value = certifyValue
            End If

            ' Format numeric columns
            If DataGridView.Rows(i).Cells(11).Value IsNot Nothing AndAlso DataGridView.Rows(i).Cells(11).Value.ToString <> "" Then
                DataGridView.Rows(i).Cells(11).Value = Format(CDbl(DataGridView.Rows(i).Cells(11).Value), "#,##0.00")
            End If
            If DataGridView.Rows(i).Cells(12).Value IsNot Nothing AndAlso DataGridView.Rows(i).Cells(12).Value.ToString <> "" Then
                DataGridView.Rows(i).Cells(12).Value = Format(CDbl(DataGridView.Rows(i).Cells(12).Value), "#,##0.00")
            End If
            If DataGridView.Rows(i).Cells(14).Value IsNot Nothing AndAlso DataGridView.Rows(i).Cells(14).Value.ToString <> "" Then
                DataGridView.Rows(i).Cells(14).Value = Format(CDbl(DataGridView.Rows(i).Cells(14).Value), "#,##0.00")
            End If
            If DataGridView.Rows(i).Cells(15).Value IsNot Nothing AndAlso DataGridView.Rows(i).Cells(15).Value.ToString <> "" Then
                DataGridView.Rows(i).Cells(15).Value = Format(CDbl(DataGridView.Rows(i).Cells(15).Value), "#,##0.00")
            End If
            If DataGridView.Rows(i).Cells(16).Value IsNot Nothing AndAlso DataGridView.Rows(i).Cells(16).Value.ToString <> "" Then
                DataGridView.Rows(i).Cells(16).Value = Format(CDbl(DataGridView.Rows(i).Cells(16).Value), "#,##0.00")
            End If
        Next
        
        GroupBox1.Visible = False
        Call amt()
    End Sub
Private Sub amt()
        amt1 = 0
        amt2 = 0
        Dim i As Integer
        For i = 0 To DataGridView.Rows.Count - 1
            Dim drValue As Double = 0
            Dim crValue As Double = 0
            
            If DataGridView.Rows(i).Cells(11).Value IsNot Nothing AndAlso DataGridView.Rows(i).Cells(11).Value.ToString <> "" Then
                Double.TryParse(DataGridView.Rows(i).Cells(11).Value.ToString, drValue)
            End If
            If DataGridView.Rows(i).Cells(12).Value IsNot Nothing AndAlso DataGridView.Rows(i).Cells(12).Value.ToString <> "" Then
                Double.TryParse(DataGridView.Rows(i).Cells(12).Value.ToString, crValue)
            End If
            
            amt1 = amt1 + drValue
            amt2 = amt2 + crValue
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

Private Sub DataGridView_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridView.SelectionChanged
        ' DataGridView is editable by default, no equivalent to FG.Editable needed
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
        DbHelper.ExecuteNonQuery(aa)
        
        Dim i As Integer
        For i = 0 To DataGridView.Rows.Count - 1
            Dim sa As String = " INSERT INTO Tmp_Import (date_work,Referno,certify, cheque_no,book, descrip, descripe, code_dr, code_cr,ac_code,  " & _
               " amount_dr,amount_cr,Curr,Rate, amount_LAK, amount_USD,bank_no, doner, Com_id, Activity_id,Cat_ID, office_id,office_nm,last_update, last_user,pc_nm) " & _
            " VALUES (  N'" & GetCellDateTime(i, 1) & "'," & _
                " N'" & Apostrophe(GetCellText(i, 2)) & "'," & _
                " N'" & Apostrophe(GetCellText(i, 3)) & "'," & _
                 " N'" & Apostrophe(GetCellText(i, 4)) & "'," & _
                  " N'" & Trim(Apostrophe(GetCellText(i, 5))) & "'," & _
                   " N'" & Apostrophe(GetCellText(i, 6)) & "'," & _
                    " N'" & Apostrophe(GetCellText(i, 7)) & "'," & _
                     " N'" & Apostrophe(GetCellText(i, 8)) & "'," & _
                      " N'" & Apostrophe(GetCellText(i, 9)) & "'," & _
              " N'" & Apostrophe(GetCellText(i, 8)) + Apostrophe(GetCellText(i, 9)) & "'," & _
              " " & GetCellDouble(i, 11) & ", " & _
            " " & GetCellDouble(i, 12) & ", " & _
             " N'" & Apostrophe(GetCellText(i, 13)) & "'," & _
             " " & GetCellDouble(i, 14) & ", " & _
             " " & GetCellDouble(i, 15) & ", " & _
             " " & GetCellDouble(i, 16) & ", " & _
               " N'" & Apostrophe(GetCellText(i, 17)) & "'," & _
              " N'" & Apostrophe(GetCellText(i, 18)) & "'," & _
               " N'" & Apostrophe(GetCellText(i, 19)) & "'," & _
                " N'" & Apostrophe(GetCellText(i, 20)) & "'," & _
                 " N'" & Apostrophe(GetCellText(i, 21)) & "'," & _
                  " N'" & Apostrophe(GetCellText(i, 22)) & "'," & _
                  " N'" & Apostrophe(GetCellText(i, 23)) & "'," & _
                       " Getdate()," & _
            " N'" & Apostrophe(MUserName) & "'," & _
             " N'" & Apostrophe(MDServerName) & "') "
            DbHelper.ExecuteNonQuery(sa)
        Next i
    End Sub
    
    ' Helper functions for DataGridView cell access
    Private Function GetCellText(ByVal rowIndex As Integer, ByVal columnIndex As Integer) As String
        If rowIndex >= 0 AndAlso rowIndex < DataGridView.Rows.Count AndAlso columnIndex >= 0 AndAlso columnIndex < DataGridView.ColumnCount Then
            If DataGridView.Rows(rowIndex).Cells(columnIndex).Value IsNot Nothing Then
                Return DataGridView.Rows(rowIndex).Cells(columnIndex).Value.ToString()
            End If
        End If
        Return ""
    End Function
    
    Private Function GetCellDouble(ByVal rowIndex As Integer, ByVal columnIndex As Integer) As Double
        If rowIndex >= 0 AndAlso rowIndex < DataGridView.Rows.Count AndAlso columnIndex >= 0 AndAlso columnIndex < DataGridView.ColumnCount Then
            If DataGridView.Rows(rowIndex).Cells(columnIndex).Value IsNot Nothing Then
                Dim value As Double = 0
                If Double.TryParse(DataGridView.Rows(rowIndex).Cells(columnIndex).Value.ToString, value) Then
                    Return value
                End If
            End If
        End If
        Return 0
    End Function
    
    Private Function GetCellDateTime(ByVal rowIndex As Integer, ByVal columnIndex As Integer) As String
        If rowIndex >= 0 AndAlso rowIndex < DataGridView.Rows.Count AndAlso columnIndex >= 0 AndAlso columnIndex < DataGridView.ColumnCount Then
            If DataGridView.Rows(rowIndex).Cells(columnIndex).Value IsNot Nothing Then
                Dim dateValue As DateTime
                If DateTime.TryParse(DataGridView.Rows(rowIndex).Cells(columnIndex).Value.ToString, dateValue) Then
                    Return Format(dateValue, "yyyy-MM-dd")
                End If
            End If
        End If
        Return DateTime.Now.ToString("yyyy-MM-dd")
    End Function

Private Sub Insert_Gen_jn()
        Frm_import_progress.Show()
        Dim aa As String
        Dim i As Integer
        
        For i = 0 To DataGridView.Rows.Count - 1
            Dim certify As String = GetCellText(i, 3)
            Dim dt As DataTable = DbHelper.GetDataTable("Select certify FROM gen_jn where  certify=N'" & certify & "'")
            
            If dt.Rows.Count = 0 Then
                aa = "   insert into gen_jn (date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, " & _
      "  code_cr,ac_code,ac_name, amount_dr, amount_cr, curr, rate,   bank_no, " & _
      "    don_id, Com_id, Activity_id, Cat_ID,  office_id,lock, my_lock, del ,AG  ,last_update, last_user,pc_nm) " & _
     "   select  date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr," & _
     "  code_cr,ac_code,descrip, amount_dr, amount_cr, curr, rate,   bank_no, " & _
     " doner, Com_id, Activity_id, Cat_id, office_id  ,0, 0, 1,0, Getdate()," & _
          " N'" & Apostrophe(MUserName) & "'," & _
          " N'" & Apostrophe(MDServerName) & "'" & _
   "    from   Tmp_Import  where  certify=N'" & certify & "'    order by certify  "
                DbHelper.ExecuteNonQuery(aa)

                aa = "  update gen_jn set Curr_i = curr where  certify=N'" & certify & "' " & _
                 " update gen_jn set Rate_i  = Rate where  certify=N'" & certify & "' "
                DbHelper.ExecuteNonQuery(aa)

                aa = "   update gen_jn set amount =amount_dr + amount_cr where  certify=N'" & certify & "'  " & _
                "  update gen_jn set amt_dr  =amount_dr where  certify=N'" & certify & "'  " & _
                " update gen_jn set amt_cr  =amount_cr  where  certify=N'" & certify & "'  "
                DbHelper.ExecuteNonQuery(aa)

                aa = "  update gen_jn set  rate =1 where  rate =0 and  certify=N'" & certify & "' "
                DbHelper.ExecuteNonQuery(aa)

                aa = "   update gen_jn set amt_USD_dr  =amount_dr / rate  where  certify=N'" & certify & "'   " & _
               "  update gen_jn set amt_USD_cr  =amount_cr  /rate   where  certify=N'" & certify & "'   "
                DbHelper.ExecuteNonQuery(aa)

                aa = "       update gen_jn set amount_dr=0 where amount_dr is null and  certify=N'" & certify & "'  " & _
                "   update gen_jn set amount_cr=0 where amount_cr is null and  certify=N'" & certify & "'   " & _
                 "  update gen_jn set amt_dr=0 where amt_dr is null and  certify=N'" & certify & "'  " & _
                "   update  gen_jn set amt_cr=0 where amt_cr is null and  certify=N'" & certify & "'  " & _
                "  update gen_jn set amt_USD_dr=0 where amt_USD_dr is null and  certify=N'" & certify & "'  " & _
                "   update  gen_jn set amt_USD_cr=0 where amt_USD_cr is null and  certify=N'" & certify & "'   "
                DbHelper.ExecuteNonQuery(aa)
                
                aa = "       insert into AP_ACC_Gen ( date_work, Referno, certify,Book ,   descrip, descripe,amount,net_amt , " & _
              "  AmountDr, AmountCr,TotalAmountDr,TotalAmountCr,curr,rate,Rate_USD,bank_no,don_id,Com_id,office_id)    " & _
             "  select     date_work, Referno, certify,Book,    '', '', " & _
               "  sum(amount_dr) as amount , sum(amount_dr) as net_amt , sum(amount_dr) as amount_dr , " & _
               "   sum(amount_cr) as amount_cr , sum(amount_dr) as TotalAmountDr , sum(amount_cr) as TotalAmountCr , " & _
              "   curr,rate,rate,bank_no,don_id,Com_id,  office_id from gen_jn where  1=1    and   certify=N'" & certify & "'  " & _
               "    group by    date_work, Referno, certify,Book,curr,rate,bank_no, " & _
               "    don_id, Com_id,  office_id    "
                DbHelper.ExecuteNonQuery(aa)
            Else
                aa = "delete gen_jn WHERE certify=N'" & certify & "'  "
                DbHelper.ExecuteNonQuery(aa)
                aa = "delete AP_ACC_Gen WHERE certify=N'" & certify & "'  "
                DbHelper.ExecuteNonQuery(aa)

                aa = "   insert into gen_jn (date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, " & _
     "  code_cr,ac_code,ac_name, amount_dr, amount_cr, curr, rate,   bank_no, " & _
    "    don_id, Com_id, Activity_id, Cat_ID,  office_id,company, lock, my_lock ,del ,AG , last_update, last_user,pc_nm) " & _
     "   select  date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr," & _
    "  code_cr,ac_code,descrip, amount_dr, amount_cr, curr, rate,   bank_no, " & _
     " doner, Com_id, Activity_id, Cat_id, office_id  , office_id , 0, 0, 1,0, Getdate()," & _
     " N'" & Apostrophe(MUserName) & "'," & _
     " N'" & Apostrophe(MDServerName) & "'" & _
     "    from   Tmp_Import  where  certify=N'" & certify & "'    order by certify  "
                DbHelper.ExecuteNonQuery(aa)

                aa = "  update gen_jn set Curr_i = curr where  certify=N'" & certify & "' " & _
                 " update gen_jn set Rate_i  = Rate where  certify=N'" & certify & "' "
                DbHelper.ExecuteNonQuery(aa)

                aa = "   update gen_jn set amount =amount_dr + amount_cr where  certify=N'" & certify & "' " & _
                "  update gen_jn set amt_dr  =amount_dr where  certify=N'" & certify & "'  " & _
                " update gen_jn set amt_cr  =amount_cr  where  certify=N'" & certify & "'  "
                DbHelper.ExecuteNonQuery(aa)

                aa = "  update gen_jn set  rate =1 where  rate =0 and  certify=N'" & certify & "' "
                DbHelper.ExecuteNonQuery(aa)

                aa = "   update gen_jn set amt_USD_dr  =amount_dr / rate where  certify=N'" & certify & "'    " & _
               "  update gen_jn set amt_USD_cr  =amount_cr  /rate   where  certify=N'" & certify & "'   "
                DbHelper.ExecuteNonQuery(aa)

                aa = "       update gen_jn set amount_dr=0 where amount_dr is null and  certify=N'" & certify & "'  " & _
                "   update gen_jn set amount_cr=0 where amount_cr is null and  certify=N'" & certify & "'   " & _
                 "  update gen_jn set amt_dr=0 where amt_dr is null and  certify=N'" & certify & "'  " & _
                "   update  gen_jn set amt_cr=0 where amt_cr is null and  certify=N'" & certify & "'  " & _
                "  update gen_jn set amt_USD_dr=0 where amt_USD_dr is null and  certify=N'" & certify & "'  " & _
                "   update  gen_jn set amt_USD_cr=0 where amt_USD_cr is null  and  certify=N'" & certify & "'  "
                DbHelper.ExecuteNonQuery(aa)

                aa = "       insert into AP_ACC_Gen ( date_work, Referno, certify,Book ,   descrip, descripe,amount,net_amt , " & _
              "  AmountDr, AmountCr,TotalAmountDr,TotalAmountCr,curr,rate,Rate_USD,bank_no,don_id,Com_id,office_id)    " & _
             "  select     date_work, Referno, certify,Book,    '', '', " & _
               "  sum(amount_dr) as amount , sum(amount_dr) as net_amt , sum(amount_dr) as amount_dr , " & _
               "   sum(amount_cr) as amount_cr , sum(amount_dr) as TotalAmountDr , sum(amount_cr) as TotalAmountCr , " & _
              "   curr,rate,rate,bank_no,don_id,Com_id,  office_id from gen_jn where  1=1    and   certify=N'" & certify & "'  " & _
               "    group by    date_work, Referno, certify,Book,curr,rate,bank_no, " & _
               "    don_id, Com_id,  office_id    "
                DbHelper.ExecuteNonQuery(aa)
            End If
            
            Frm_import_progress.Refresh()
            Frm_import_progress.Label2.Text = DataGridView.Rows.Count.ToString()
            Frm_import_progress.Label4.Text = GetCellText(i, 0)
            Frm_import_progress.Label1.Text = GetCellText(i, 2)
        Next
        
        Frm_import_progress.Close()
        
        aa = "   update AP_ACC_Gen set net_usd  =net_amt / rate   where curr='LAK'  "
        DbHelper.ExecuteNonQuery(aa)
        aa = "   update AP_ACC_Gen set net_usd  =net_amt   where curr='USD'  "
        DbHelper.ExecuteNonQuery(aa)

        aa = "   update AP_ACC_Gen set cheque_no  = gen_jn.cheque_no from gen_jn where  AP_ACC_Gen.certify=gen_jn.certify "
        DbHelper.ExecuteNonQuery(aa)

        aa = "   update AP_ACC_Gen set Com_id =office_id " & _
          "    update gen_jn set Com_id =office_id  " & _
            " update gen_jn set don_id  ='01'  " & _
                "  update AP_ACC_Gen set don_id  ='01' "
        DbHelper.ExecuteNonQuery(aa)

        aa = "   update gen_jn set amt_dr  =  amt_dr * Rate where  amt_dr >0 "
        DbHelper.ExecuteNonQuery(aa)
        aa = "   update gen_jn set amt_cr  =  amt_cr * Rate where  amt_cr >0   "
        DbHelper.ExecuteNonQuery(aa)
    End Sub
Private Sub Insert_Gen_jn22()
        'Frm_import_progress.Show()
        Dim aa As String
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
            DbHelper.ExecuteNonQuery(aa)

        aa = "  update gen_jn set Curr_i = curr  " & _
         " update gen_jn set Rate_i  = Rate "
            DbHelper.ExecuteNonQuery(aa)

        aa = "   update gen_jn set amount =amount_dr + amount_cr  " & _
        "  update gen_jn set amt_dr  =amount_dr  " & _
        " update gen_jn set amt_cr  =amount_cr   "
            DbHelper.ExecuteNonQuery(aa)

        aa = "  update gen_jn set  rate =1 where  rate =0  "
            DbHelper.ExecuteNonQuery(aa)

        aa = "   update gen_jn set amt_USD_dr  =amount_dr / rate  " & _
       "  update gen_jn set amt_USD_cr  =amount_cr  /rate      "
            DbHelper.ExecuteNonQuery(aa)

        aa = "       update gen_jn set amount_dr=0 where amount_dr is null  " & _
        "   update gen_jn set amount_cr=0 where amount_cr is null " & _
         "  update gen_jn set amt_dr=0 where amt_dr is null  " & _
        "   update  gen_jn set amt_cr=0 where amt_cr is null   " & _
        "  update gen_jn set amt_USD_dr=0 where amt_USD_dr is null  " & _
        "   update  gen_jn set amt_USD_cr=0 where amt_USD_cr is null   "
            DbHelper.ExecuteNonQuery(aa)
            
        aa = "       insert into AP_ACC_Gen ( date_work, Referno, certify,Book ,   descrip, descripe,amount,net_amt , " & _
      "  AmountDr, AmountCr,TotalAmountDr,TotalAmountCr,curr,rate,Rate_USD,bank_no,don_id,Com_id,office_id)    " & _
     "  select     date_work, Referno, certify,Book,    '', '', " & _
       "  sum(amount_dr) as amount , sum(amount_dr) as net_amt , sum(amount_dr) as amount_dr , " & _
       "   sum(amount_cr) as amount_cr , sum(amount_dr) as TotalAmountDr , sum(amount_cr) as TotalAmountCr , " & _
      "   curr,rate,rate,bank_no,don_id,Com_id,  office_id from gen_jn  " & _
        "    group by    date_work, Referno, certify,Book,curr,rate,bank_no, " & _
        "    don_id, Com_id,  office_id    "
            DbHelper.ExecuteNonQuery(aa)

            aa = "   update AP_ACC_Gen set net_usd  =net_amt / rate   where curr='LAK'  "
            DbHelper.ExecuteNonQuery(aa)
            aa = "   update AP_ACC_Gen set net_usd  =net_amt   where curr='USD'  "
            DbHelper.ExecuteNonQuery(aa)

            aa = "   update AP_ACC_Gen set cheque_no  = gen_jn.cheque_no from gen_jn where  AP_ACC_Gen.certify=gen_jn.certify "
            DbHelper.ExecuteNonQuery(aa)

            aa = "   update AP_ACC_Gen set Com_id =office_id " & _
              "    update gen_jn set Com_id =office_id  " & _
                " update gen_jn set don_id  ='01'  " & _
                    "  update AP_ACC_Gen set don_id  ='01' "
            DbHelper.ExecuteNonQuery(aa)
        DbHelper.ExecuteNonQuery(" update gen_jn set company =office_id  ")

            aa = "   update gen_jn set amt_dr  =  amt_dr * Rate where  amt_dr >0 "
            DbHelper.ExecuteNonQuery(aa)
            aa = "   update gen_jn set amt_cr  =  amt_cr * Rate where  amt_cr >0   "
            DbHelper.ExecuteNonQuery(aa)
    End Sub
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()
    End Sub
End Class