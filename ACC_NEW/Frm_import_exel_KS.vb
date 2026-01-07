Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports ApPBank10.Module

Public Class Frm_import_exel_KS
    Dim amt1 As Double
    Dim amt2 As Double
    Dim MLAK, MUSD, MTHB, MEUR As Double
    Dim DataGridViewMain As New DataGridView
Private Sub Frm_import_exel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InitializeDataGridView()
        dff()
    End Sub

    Private Sub InitializeDataGridView()
        ' Replace FG with DataGridViewMain
        DataGridViewMain.Dock = DockStyle.Fill
        DataGridViewMain.AllowUserToAddRows = False
        DataGridViewMain.AllowUserToDeleteRows = False
        DataGridViewMain.ReadOnly = False
        DataGridViewMain.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridViewMain.MultiSelect = False
        
        ' Add columns based on the original FG.FormatString
        DataGridViewMain.Columns.Clear()
        DataGridViewMain.Columns.Add("Col0", "ລ/ດ")
        DataGridViewMain.Columns.Add("Col1", "ວັນທີ")
        DataGridViewMain.Columns.Add("Col2", "ເລກທີ່ບິນ")
        DataGridViewMain.Columns.Add("Col3", "ເລກທີໂອໂຕ")
        DataGridViewMain.Columns.Add("Col4", "ແຊັກ")
        DataGridViewMain.Columns.Add("Col5", "ປື້ມ")
        DataGridViewMain.Columns.Add("Col6", "ເນື້ອໃນພາສາລາວ")
        DataGridViewMain.Columns.Add("Col7", "ເນື້ອໃນພາສາອັງກິດ")
        DataGridViewMain.Columns.Add("Col8", "ໜີ້")
        DataGridViewMain.Columns.Add("Col9", "ມີ")
        DataGridViewMain.Columns.Add("Col10", "ຈຳນວນເງີນໜີ້")
        DataGridViewMain.Columns.Add("Col11", "ຈຳນວນເງີນມີ")
        DataGridViewMain.Columns.Add("Col12", "ສະກຸນເງິນ")
        DataGridViewMain.Columns.Add("Col13", "ອັດຕາແລກປ່ຽນ")
        DataGridViewMain.Columns.Add("Col14", "ຈຳນວນເງີນກີບ")
        DataGridViewMain.Columns.Add("Col15", "ຈຳນວນເງີນໂດລາ")
        DataGridViewMain.Columns.Add("Col16", "ເລກບັນຊີທະນາຄານ")
        DataGridViewMain.Columns.Add("Col17", "ແຫຼງທຶນ")
        DataGridViewMain.Columns.Add("Col18", "ອົງປະກອບ")
        DataGridViewMain.Columns.Add("Col19", "ລະຫັດກິດຈະກຳ")
        DataGridViewMain.Columns.Add("Col20", "ປະເພດລາຍຈ່າຍ")
        DataGridViewMain.Columns.Add("Col21", "ລະຫັດ")
        DataGridViewMain.Columns.Add("Col22", "ສຳນັກງານ")
        
        ' Set column widths
        For i As Integer = 0 To DataGridViewMain.Columns.Count - 1
            DataGridViewMain.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next
        
        Panel1.Controls.Add(DataGridViewMain)
    End Sub
Private Sub dff()
        ' Migrate LoadSqlData calls to DbHelper
        Dim dt As DataTable
        
        dt = DbHelper.GetDataTable("select * from Ap_RateSeting where Curr='LAK'")
        If dt.Rows.Count > 0 Then
            MLAK = Trim(GetStr(dt.Rows(0)("Rate")))
        End If
        
        dt = DbHelper.GetDataTable("select * from Ap_RateSeting where Curr='USD'")
        If dt.Rows.Count > 0 Then
            MUSD = Trim(GetStr(dt.Rows(0)("Rate")))
        End If
        
        dt = DbHelper.GetDataTable("select * from Ap_RateSeting where Curr='THB'")
        If dt.Rows.Count > 0 Then
            MTHB = Trim(GetStr(dt.Rows(0)("Rate")))
        End If
        
        dt = DbHelper.GetDataTable("select * from Ap_RateSeting where Curr='EUR'")
        If dt.Rows.Count > 0 Then
            MEUR = Trim(GetStr(dt.Rows(0)("Rate")))
        End If
        'With RSC
        '    Do Until .EOF = True
        '        txtRate.Text = Trim(.Fields("Rate").Value)
        '        Curr_Last.Text = Trim(.Fields("curr_Last").Value)
        '        .MoveNext()
        '    Loop
        'End With
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

        DataGridViewMain.Rows.Clear()
        Dim CU As String
        For i = 0 To DataGridView1.RowCount - 2
            If DataGridView1.Item(1, i).Value IsNot Nothing AndAlso DataGridView1.Item(1, i).Value.ToString <> "" Then
                Dim rowIndex As Integer = DataGridViewMain.Rows.Add()
                
                DataGridViewMain.Rows(rowIndex).Cells(1).Value = DataGridView1.Item(0, i).Value  ' ວັນທີ
                DataGridViewMain.Rows(rowIndex).Cells(2).Value = Microsoft.VisualBasic.Left(DataGridView1.Item(2, i).Value.ToString, 10)  ' ເລກທີ່ບິນ
                DataGridViewMain.Rows(rowIndex).Cells(3).Value = DataGridView1.Item(0, i).Value  ' ເລກທີໂອໂຕ
                DataGridViewMain.Rows(rowIndex).Cells(4).Value = ""  ' ແຊັກ
                DataGridViewMain.Rows(rowIndex).Cells(5).Value = "GL"  ' ປື້ມ
                DataGridViewMain.Rows(rowIndex).Cells(6).Value = DataGridView1.Item(5, i).Value  ' ເນື້ອໃນພາສາລາວ
                DataGridViewMain.Rows(rowIndex).Cells(7).Value = ""  ' ເນື້ອໃນພາສາອັງກິດ
                DataGridViewMain.Rows(rowIndex).Cells(8).Value = DataGridView1.Item(9, i).Value  ' ໜີ້
                DataGridViewMain.Rows(rowIndex).Cells(9).Value = DataGridView1.Item(10, i).Value  ' ມີ
                DataGridViewMain.Rows(rowIndex).Cells(10).Value = DataGridView1.Item(11, i).Value  ' ຈຳນວນເງີນໜີ້
                DataGridViewMain.Rows(rowIndex).Cells(11).Value = DataGridView1.Item(12, i).Value  ' ຈຳນວນເງີນມີ
                DataGridViewMain.Rows(rowIndex).Cells(12).Value = DataGridView1.Item(12, i).Value  ' ສະກຸນເງິນ
                DataGridViewMain.Rows(rowIndex).Cells(13).Value = Microsoft.VisualBasic.Right(If(DataGridView1.Item(8, i).Value, ""), 4)  ' ອັດຕາແລກປ່ຽນ
                DataGridViewMain.Rows(rowIndex).Cells(14).Value = "1"  ' ຈຳນວນເງີນກີບ
                DataGridViewMain.Rows(rowIndex).Cells(15).Value = 0  ' ຈຳນວນເງີນໂດລາ
                DataGridViewMain.Rows(rowIndex).Cells(16).Value = ""  ' ເລກບັນຊີທະນາຄານ
                DataGridViewMain.Rows(rowIndex).Cells(17).Value = ""  ' ແຫຼງທຶນ
                DataGridViewMain.Rows(rowIndex).Cells(18).Value = "01"  ' ອົງປະກອບ
                DataGridViewMain.Rows(rowIndex).Cells(19).Value = "01"  ' ລະຫັດກິດຈະກຳ
                DataGridViewMain.Rows(rowIndex).Cells(20).Value = ""  ' ປະເພດລາຍຈ່າຍ
                DataGridViewMain.Rows(rowIndex).Cells(21).Value = ""  ' ລະຫັດ
                DataGridViewMain.Rows(rowIndex).Cells(22).Value = "01-02"  ' ສຳນັກງານ
            End If
            CU = Microsoft.VisualBasic.Right(If(DataGridView1.Item(8, i).Value, ""), 4)
        Next i

        For i = 0 To DataGridViewMain.Rows.Count - 1
            no = no + 1
            DataGridViewMain.Rows(i).Cells(0).Value = no

            DataGridViewMain.Rows(i).Cells(12).Value = Microsoft.VisualBasic.Left(GetStr(DataGridViewMain.Rows(i).Cells(12).Value), 3)
            
            Dim currCode As String = Trim(GetStr(DataGridViewMain.Rows(i).Cells(12).Value))
            Select Case currCode
                Case "LAK"
                    DataGridViewMain.Rows(i).Cells(13).Value = 1
                Case "USD"
                    DataGridViewMain.Rows(i).Cells(13).Value = MUSD
                Case "THB"
                    DataGridViewMain.Rows(i).Cells(13).Value = MTHB
                Case "EUR"
                    DataGridViewMain.Rows(i).Cells(13).Value = MEUR
            End Select

            Dim rateValue As Double = 0
            If Double.TryParse(GetStr(DataGridViewMain.Rows(i).Cells(13).Value), rateValue) Then
                DataGridViewMain.Rows(i).Cells(13).Value = Format(rateValue, "#,##0.00")
            End If

            Dim amountDr As Double = 0
            Dim amountCr As Double = 0
            Double.TryParse(GetStr(DataGridViewMain.Rows(i).Cells(10).Value), amountDr)
            Double.TryParse(GetStr(DataGridViewMain.Rows(i).Cells(11).Value), amountCr)
            
            DataGridViewMain.Rows(i).Cells(14).Value = Format(amountDr + amountCr, "#,##0.00")
        Next
        
        GroupBox1.Visible = False
    End Sub
Private Sub amt()
        amt1 = 0
        amt2 = 0
        Dim i As Integer
        For i = 0 To DataGridViewMain.Rows.Count - 1
            Dim amountDr As Double = 0
            Dim amountCr As Double = 0
            Double.TryParse(GetStr(DataGridViewMain.Rows(i).Cells(10).Value), amountDr)
            Double.TryParse(GetStr(DataGridViewMain.Rows(i).Cells(11).Value), amountCr)
            amt1 = amt1 + amountDr
            amt2 = amt2 + amountCr
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

Private Sub DataGridView1_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        ' DataGridView is always editable by default, no equivalent to VSFlexGrid.Editable needed
    End Sub

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        Call amt()
        If CDbl(txtSumAmountDr.Text) <> CDbl(txtSumAmountCr.Text) Then MsgBox("ບັນຊີນີ້ບໍ່ທັນບໍ່ທັນດູນດ່ຽງ ບໍ່ສາມາດບັນທຶກໄດ້!", MsgBoxStyle.OkOnly) : Exit Sub
        Call SaveItems()
        Call Insert_Gen_jn()
        MsgBox("ບັນທຶກສຳເລັດ")
    End Sub
Private Sub SaveItems()
        Dim sql As String
        sql = " delete Tmp_Import  "
        DbHelper.ExecuteNonQuery(sql)
        
        Dim i As Integer
        For i = 0 To DataGridViewMain.Rows.Count - 1
            Dim sa As String = " INSERT INTO Tmp_Import (date_work,Referno,certify, cheque_no,book, descrip, descripe, code_dr, code_cr,ac_code,  " & _
               " amount_dr,amount_cr,Curr,Rate, amount_LAK, amount_USD,bank_no, doner, Com_id, Activity_id,Cat_ID, office_id,office_nm,last_update, last_user,pc_nm) " & _
            " VALUES (  '" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'," & _
                " '" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(2).Value)) & "'," & _
                " '" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'," & _
                 " '" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(4).Value)) & "'," & _
                  " '" & Trim(Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(5).Value))) & "'," & _
                   " '" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(6).Value)) & "'," & _
                    " '" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(7).Value)) & "'," & _
                     " '" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(8).Value)) & "'," & _
                      " '" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(9).Value)) & "'," & _
              " '" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(8).Value)) + Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(9).Value)) & "'," & _
              " " & CDbl(GetStr(DataGridViewMain.Rows(i).Cells(10).Value)) & ", " & _
            " " & CDbl(GetStr(DataGridViewMain.Rows(i).Cells(11).Value)) & ", " & _
             " '" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(12).Value)) & "'," & _
             " " & CDbl(GetStr(DataGridViewMain.Rows(i).Cells(13).Value)) & ", " & _
             " " & CDbl(GetStr(DataGridViewMain.Rows(i).Cells(14).Value)) & ", " & _
             " " & CDbl(GetStr(DataGridViewMain.Rows(i).Cells(15).Value)) & ", " & _
               " '" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(16).Value)) & "'," & _
              " '" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(17).Value)) & "'," & _
               " '" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(18).Value)) & "'," & _
                " '" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(19).Value)) & "'," & _
                 " '" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(20).Value)) & "'," & _
                  " '" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(21).Value)) & "'," & _
                   " '" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(22).Value)) & "'," & _
                       " Getdate()," & _
            " '" & Apostrophe(MUserName) & "'," & _
             " '" & Apostrophe(MDServerName) & "') "
            DbHelper.ExecuteNonQuery(sa)
        Next i
        
        DbHelper.ExecuteNonQuery("update Tmp_Import set Tmp_Import.ac_code=Acc_Code.ac_code from Acc_Code,Tmp_Import where Acc_Code.ac_original=Tmp_Import.ac_code ")
        DbHelper.ExecuteNonQuery("update Tmp_Import set Tmp_Import.code_dr=Acc_Code.ac_code from Acc_Code,Tmp_Import where Acc_Code.ac_original=Tmp_Import.code_dr and Tmp_Import.code_dr<>''  ")
        DbHelper.ExecuteNonQuery("update Tmp_Import set Tmp_Import.code_cr=Acc_Code.ac_code from Acc_Code,Tmp_Import where Acc_Code.ac_original=Tmp_Import.code_cr and Tmp_Import.code_Cr<>'' ")
    End Sub

Private Sub Insert_Gen_jn()
        Frm_import_progress.Show()
        Dim sql As String
        Dim i As Integer
        For i = 0 To DataGridViewMain.Rows.Count - 1
            Dim dtCheck As DataTable = DbHelper.GetDataTable("Select * FROM gen_jn where  certify='" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work='" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'  ")
            If dtCheck.Rows.Count = 0 Then
sql = "   insert into gen_jn (date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, " & _
             "  code_cr,ac_code,ac_name, amount_dr, amount_cr, curr, rate,   bank_no, " & _
             "    don_id, Com_id, Activity_id, Cat_ID,  office_id,lock, my_lock, del ,AG  ,last_update, last_user,pc_nm) " & _
            "   select  date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr," & _
            "  code_cr,ac_code,descrip, amount_dr, amount_cr, curr, rate,   bank_no, " & _
            " doner, Com_id, Activity_id, Cat_id, office_id  ,0, 0, 1,0, Getdate()," & _
                 " '" & Apostrophe(MUserName) & "'," & _
                 " '" & Apostrophe(MDServerName) & "'" & _
          "    from   Tmp_Import    where  certify='" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work='" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    order by certify  "
                    DbHelper.ExecuteNonQuery(sql)

                    'aa = " update gen_jn set don_id= AP_Donnor.Don_ID  from AP_Donnor where  gen_jn.don_id= AP_Donnor.Don_Sym   "
                    'CNN.Execute(aa)
                    'aa = "  update gen_jn set office_id= AP_Office.off_id  from AP_Office  where  gen_jn.office_id= AP_Office.office_id   "
                    'CNN.Execute(aa)

sql = "  update gen_jn set Curr_i = curr where   certify='" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work='" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    " & _
                      " update gen_jn set Rate_i  = Rate where   certify='" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work='" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    "
                    DbHelper.ExecuteNonQuery(sql)

                    sql = "   update gen_jn set amount =amount_dr + amount_cr where   certify='" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work='" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'   " & _
                    "  update gen_jn set amt_dr  =amount_dr where   certify='" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work='" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'   " & _
                    " update gen_jn set amt_cr  =amount_cr  where   certify='" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work='" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'     "
                    DbHelper.ExecuteNonQuery(sql)

                    sql = "  update gen_jn set  rate =1 where  rate =0 and   certify='" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work='" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    "
                    DbHelper.ExecuteNonQuery(sql)

                    sql = "   update gen_jn set amt_USD_dr  =amount_dr / rate  where   certify='" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work='" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    " & _
                   "  update gen_jn set amt_USD_cr  =amount_cr  /rate   where   certify='" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work='" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'      "
                    DbHelper.ExecuteNonQuery(sql)

                    sql = "       update gen_jn set amount_dr=0 where amount_dr is null and   certify='" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work='" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    " & _
                    "   update gen_jn set amount_cr=0 where amount_cr is null and   certify='" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work='" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'     " & _
                     "  update gen_jn set amt_dr=0 where amt_dr is null and   certify='" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work='" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    " & _
                    "   update  gen_jn set amt_cr=0 where amt_cr is null and   certify='" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work='" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    " & _
                    "  update gen_jn set amt_USD_dr=0 where amt_USD_dr is null and   certify='" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work='" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    " & _
                    "   update  gen_jn set amt_USD_cr=0 where amt_USD_cr is null and   certify='" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work='" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'     "
                    DbHelper.ExecuteNonQuery(sql)
                    
                    sql = "       insert into AP_ACC_Gen ( date_work, Referno, certify,Book ,   descrip, descripe,amount,net_amt , " & _
                  "  AmountDr, AmountCr,TotalAmountDr,TotalAmountCr,curr,rate,Rate_USD,bank_no,don_id,Com_id,office_id)    " & _
                 "  select     date_work, Referno, certify,Book,    '', '', " & _
                   "  sum(amount_dr) as amount , sum(amount_dr) as net_amt , sum(amount_dr) as amount_dr , " & _
                   "   sum(amount_cr) as amount_cr , sum(amount_dr) as TotalAmountDr , sum(amount_cr) as TotalAmountCr , " & _
                  "   curr,rate,rate,bank_no,don_id,Com_id,  office_id from gen_jn where  1=1    and    certify='" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work='" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    " & _
                   "    group by    date_work, Referno, certify,Book,curr,rate,bank_no, " & _
                   "    don_id, Com_id,  office_id    "
                    DbHelper.ExecuteNonQuery(sql)
Else
                    'Dim sk As String = "Select * FROM gen_jn where  certify=N'" & GetStr(DataGridViewMain.Rows(i).Cells(3).Value) & "'  and date_work=N'" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'  "
                    'Call LoadSqlData(sk, Rschk)
                    aa = "delete gen_jn WHERE certify=N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work='" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'  "
                    DbHelper.ExecuteNonQuery(aa)

                    aa = "   insert into gen_jn (date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, " & _
         "  code_cr,ac_code,ac_name, amount_dr, amount_cr, curr, rate,   bank_no, " & _
        "    don_id, Com_id, Activity_id, Cat_ID,  office_id,company, lock, my_lock ,del ,AG , last_update, last_user,pc_nm) " & _
         "   select  date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr," & _
        "  code_cr,ac_code,descrip, amount_dr, amount_cr, curr, rate,   bank_no, " & _
         " doner, Com_id, Activity_id, Cat_id, office_id  , office_id , 0, 0, 1,0, Getdate()," & _
         " N'" & Apostrophe(MUserName) & "'," & _
         " N'" & Apostrophe(MDServerName) & "'" & _
         "    from   Tmp_Import  where   certify=N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work=N'" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'     order by certify  "
                    DbHelper.ExecuteNonQuery(aa)

                    'aa = " update gen_jn set don_id= AP_Donnor.Don_ID  from AP_Donnor where  gen_jn.don_id= AP_Donnor.Don_Sym   "
                    'DbHelper.ExecuteNonQuery(aa)
                    'aa = "  update gen_jn set office_id= AP_Office.off_id  from AP_Office  where  gen_jn.office_id= AP_Office.office_id   "
                    'DbHelper.ExecuteNonQuery(aa)

                    aa = "  update gen_jn set Curr_i = curr where   certify=N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work=N'" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    " & _
                     " update gen_jn set Rate_i  = Rate where   certify=N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work=N'" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    "
                    DbHelper.ExecuteNonQuery(aa)


                    aa = "   update gen_jn set amount =amount_dr + amount_cr where  certify=N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work=N'" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'   " & _
                    "  update gen_jn set amt_dr  =amount_dr where  certify=N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work=N'" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    " & _
                    " update gen_jn set amt_cr  =amount_cr where  certify=N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work=N'" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    "
                    DbHelper.ExecuteNonQuery(aa)

                    aa = "  update gen_jn set  rate =1 where  rate =0 and  certify=N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work=N'" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    "
                    DbHelper.ExecuteNonQuery(aa)

                    aa = "   update gen_jn set amt_USD_dr  =amount_dr / rate where  certify=N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'   and date_work=N'" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'     " & _
                   "  update gen_jn set amt_USD_cr  =amount_cr  /rate   where  certify=N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work=N'" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'     "
                    DbHelper.ExecuteNonQuery(aa)

                    aa = "       update gen_jn set amount_dr=0 where amount_dr is null and  certify=N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work=N'" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    " & _
                   "   update gen_jn set amount_cr=0 where amount_cr is null and  certify=N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work=N'" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'     " & _
                    "  update gen_jn set amt_dr=0 where amt_dr is null and  certify=N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work=N'" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    " & _
                   "   update  gen_jn set amt_cr=0 where amt_cr is null and  certify=N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work=N'" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    " & _
                   "  update gen_jn set amt_USD_dr=0 where amt_USD_dr is null and  certify=N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work=N'" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    " & _
                   "   update  gen_jn set amt_USD_cr=0 where amt_USD_cr is null  and  certify=N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work=N'" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    "
                    DbHelper.ExecuteNonQuery(aa)

                    aa = "       insert into AP_ACC_Gen ( date_work, Referno, certify,Book ,   descrip, descripe,amount,net_amt , " & _
                  "  AmountDr, AmountCr,TotalAmountDr,TotalAmountCr,curr,rate,Rate_USD,bank_no,don_id,Com_id,office_id)    " & _
                 "  select     date_work, Referno, certify,Book,    '', '', " & _
                   "  sum(amount_dr) as amount , sum(amount_dr) as net_amt , sum(amount_dr) as amount_dr , " & _
                   "   sum(amount_cr) as amount_cr , sum(amount_dr) as TotalAmountDr , sum(amount_cr) as TotalAmountCr , " & _
                  "   curr,rate,rate,bank_no,don_id,Com_id,  office_id from gen_jn where  1=1    and   certify=N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'  and date_work=N'" & Format(CDate(GetStr(DataGridViewMain.Rows(i).Cells(1).Value)), "yyyy-MM-dd") & "'    " & _
                   "    group by    date_work, Referno, certify,Book,curr,rate,bank_no, " & _
                   "    don_id, Com_id,  office_id    "
DbHelper.ExecuteNonQuery(aa)

                End If
                Frm_import_progress.Refresh()

                Frm_import_progress.Label2.Text = DataGridViewMain.Rows.Count

Frm_import_progress.Label4.Text = GetStr(DataGridViewMain.Rows(i).Cells(0).Value)

                Frm_import_progress.Label1.Text = GetStr(DataGridViewMain.Rows(i).Cells(2).Value)

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
        'DbHelper.ExecuteNonQuery(aa)
        Dim aa As String
        aa = "   update AP_ACC_Gen set net_usd  =net_amt / rate   where curr='LAK'  "
        DbHelper.ExecuteNonQuery(aa)
        aa = "   update AP_ACC_Gen set net_usd  =net_amt   where curr='USD'  "
        DbHelper.ExecuteNonQuery(aa)



        aa = "   update AP_ACC_Gen set cheque_no  = gen_jn.cheque_no from gen_jn where  AP_ACC_Gen.certify=gen_jn.certify  and AP_ACC_Gen.cheque_no is null "
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

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()
    End Sub
End Class