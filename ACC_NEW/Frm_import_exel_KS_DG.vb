Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports ApPBank10.Module

Public Class Frm_import_exel_KS_DG
    Dim amt1 As Double
    Dim amt2 As Double
    Dim MLAK, MUSD, MTHB, MEUR As Double

    Dim cn As New OleDbConnection
    Dim cm As New OleDbCommand
    Dim da As OleDbDataAdapter
    Dim dt As New DataTable
    Dim A1 As String = ""
    Dim A2 As String = ""
    Dim A3 As String = ""
    Dim A4 As String = ""
    Dim A5 As String = ""
    Dim A6 As String = ""
    Dim A7 As String = ""
    Dim A8 As String = ""
    Dim A9 As String = ""
    Dim A10 As String = ""
    Dim A11 As String = ""
    Dim A12 As String = ""
    Dim A13 As String = ""
    Dim A14 As String = ""
    Dim A15 As String = ""
    Dim A16 As String = ""
    Dim DtSet As System.Data.DataSet
    Dim DataGridViewMain As New DataGridView

Private Sub Frm_import_exel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InitializeDataGridView()
        dff()
        CheckBox2.Text = "V.S"
    End Sub

    Private Sub InitializeDataGridView()
        ' Replace FG with DataGridViewMain
        DataGridViewMain.Dock = DockStyle.Fill
        DataGridViewMain.AllowUserToAddRows = False
        DataGridViewMain.AllowUserToDeleteRows = False
        DataGridViewMain.ReadOnly = False
        DataGridViewMain.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridViewMain.MultiSelect = False
        
        ' Add columns based on DataGridView configuration
        DataGridViewMain.Columns.Clear()
        DataGridViewMain.Columns.Add("Col0", "ລ/ດ")
        DataGridViewMain.Columns.Add("Col1", "ວັນທີ")
        DataGridViewMain.Columns.Add("Col2", "ເລກທີ່ບິນ")
        DataGridViewMain.Columns.Add("Col3", "ເລກທີໂອໂຕ")
        DataGridViewMain.Columns.Add("Col4", "ແຊັກ")
        DataGridViewMain.Columns.Add("Col5", "ປື້ມ")
        DataGridViewMain.Columns.Add("Col6", "ເນື້ອໃນພາສາລາວ")
        DataGridViewMain.Columns.Add("Col7", "ເນື້ອໃນພາສາອັງກິດ")
        DataGridViewMain.Columns.Add("Col8", "ຜີ້")
        DataGridViewMain.Columns.Add("Col9", "ມີ")
        DataGridViewMain.Columns.Add("Col10", "ຈຳນວນເງີນຜີ້")
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

        ' Clear DataGridViewMain
        DataGridViewMain.Rows.Clear()
        Dim CU As String
        Try
            For i = 0 To DataGridView1.RowCount - 2
                If DataGridView1.Item(2, i).Value IsNot Nothing AndAlso DataGridView1.Item(2, i).Value.ToString <> "" Then
                    ' Add row to DataGridViewMain
                    Dim rowIndex As Integer = DataGridViewMain.Rows.Add()
                    
                    ' Populate columns with data from DataGridView1
                    DataGridViewMain.Rows(rowIndex).Cells(0).Value = rowIndex + 1  ' ລ/ດ
                    DataGridViewMain.Rows(rowIndex).Cells(1).Value = DataGridView1.Item(0, i).Value  ' ວັນທີ
                    DataGridViewMain.Rows(rowIndex).Cells(2).Value = Microsoft.VisualBasic.Left(DataGridView1.Item(2, i).Value.ToString, 10)  ' ເລກທີ່ບິນ
                    DataGridViewMain.Rows(rowIndex).Cells(3).Value = DataGridView1.Item(0, i).Value  ' ເລກທີໂອໂຕ
                    DataGridViewMain.Rows(rowIndex).Cells(4).Value = ""  ' ແຊັກ
                    DataGridViewMain.Rows(rowIndex).Cells(5).Value = "GL"  ' ປື້ມ
                    DataGridViewMain.Rows(rowIndex).Cells(6).Value = DataGridView1.Item(5, i).Value  ' ເນື້ອໃນພາສາລາວ
                    DataGridViewMain.Rows(rowIndex).Cells(7).Value = ""  ' ເນື້ອໃນພາສາອັງກິດ
                    DataGridViewMain.Rows(rowIndex).Cells(8).Value = DataGridView1.Item(9, i).Value  ' ຜີ້
                    DataGridViewMain.Rows(rowIndex).Cells(9).Value = DataGridView1.Item(10, i).Value  ' ມີ
                    DataGridViewMain.Rows(rowIndex).Cells(10).Value = DataGridView1.Item(11, i).Value  ' ຈຳນວນເງີນຜີ້
                    DataGridViewMain.Rows(rowIndex).Cells(11).Value = DataGridView1.Item(12, i).Value  ' ຈຳນວນເງີນມີ
                    DataGridViewMain.Rows(rowIndex).Cells(12).Value = DataGridView1.Item(12, i).Value  ' ສະກຸນເງິນ
                    DataGridViewMain.Rows(rowIndex).Cells(13).Value = Microsoft.VisualBasic.Right(If(DataGridView1.Item(8, i).Value, ""), 4)  ' ອັດຕາແລກປ່ຽນ
                    DataGridViewMain.Rows(rowIndex).Cells(14).Value = "1"  ' ຈຳນວນເງີນກີບ
                    DataGridViewMain.Rows(rowIndex).Cells(15).Value = 0  ' ຈຳນວນເງີນໂດລາ
                    DataGridViewMain.Rows(rowIndex).Cells(16).Value = ""  ' ເລກບັນຊີທະນາຄານ
                    DataGridViewMain.Rows(rowIndex).Cells(17).Value = "01"  ' ແຫຼງທຶນ
                    DataGridViewMain.Rows(rowIndex).Cells(18).Value = "01"  ' ອົງປະກອບ
                    DataGridViewMain.Rows(rowIndex).Cells(19).Value = ""  ' ລະຫັດກິດຈະກຳ
                    DataGridViewMain.Rows(rowIndex).Cells(20).Value = ""  ' ປະເພດລາຍຈ່າຍ
                    DataGridViewMain.Rows(rowIndex).Cells(21).Value = "01-02"  ' ລະຫັດ
                    DataGridViewMain.Rows(rowIndex).Cells(22).Value = ""  ' ສຳນັກງານ
                End If
            Next i
            
            ' Process the imported data
            For i = 0 To DataGridViewMain.Rows.Count - 1
                If DataGridViewMain.Rows(i).Cells(12).Value IsNot Nothing Then
                    Dim currency As String = Microsoft.VisualBasic.Left(GetStr(DataGridViewMain.Rows(i).Cells(12).Value), 3)
                    DataGridViewMain.Rows(i).Cells(12).Value = currency
                    
                    Select Case Trim(currency)
                        Case "LAK"
                            DataGridViewMain.Rows(i).Cells(13).Value = 1
                        Case "USD"
                            DataGridViewMain.Rows(i).Cells(13).Value = MUSD
                        Case "THB"
                            DataGridViewMain.Rows(i).Cells(13).Value = MTHB
                        Case "EUR"
                            DataGridViewMain.Rows(i).Cells(13).Value = MEUR
                    End Select
                    
                    ' Format rate and calculate amounts
                    If DataGridViewMain.Rows(i).Cells(13).Value IsNot Nothing Then
                        DataGridViewMain.Rows(i).Cells(13).Value = Format(CDbl(DataGridViewMain.Rows(i).Cells(13).Value), "#,##0.00")
                    End If
                    
                    Dim amountDr As Double = 0
                    Dim amountCr As Double = 0
                    If DataGridViewMain.Rows(i).Cells(10).Value IsNot Nothing Then
                        Double.TryParse(GetStr(DataGridViewMain.Rows(i).Cells(10).Value), amountDr)
                    End If
                    If DataGridViewMain.Rows(i).Cells(11).Value IsNot Nothing Then
                        Double.TryParse(GetStr(DataGridViewMain.Rows(i).Cells(11).Value), amountCr)
                    End If
                    
                    DataGridViewMain.Rows(i).Cells(14).Value = Format(amountDr + amountCr, "#,##0.00")
                End If
            Next
            
        Catch ex As Exception
            MessageBox.Show("Error importing data: " & ex.Message)
        End Try

        GroupBox1.Visible = False
        amt()
    End Sub
Private Sub amt()
        amt1 = 0
        amt2 = 0
        Dim i As Integer
        For i = 0 To DataGridViewMain.Rows.Count - 1
            Dim amountDr As Double = 0
            Dim amountCr As Double = 0
            
            If DataGridViewMain.Rows(i).Cells(10).Value IsNot Nothing Then
                Double.TryParse(GetStr(DataGridViewMain.Rows(i).Cells(10).Value), amountDr)
            End If
            If DataGridViewMain.Rows(i).Cells(11).Value IsNot Nothing Then
                Double.TryParse(GetStr(DataGridViewMain.Rows(i).Cells(11).Value), amountCr)
            End If
            
            amt1 += amountDr
            amt2 += amountCr
        Next i

        txtSumAmountDr.Text = Format(amt1, "##,##0.00")
        txtSumAmountCr.Text = Format(amt2, "##,##0.00")
    End Sub

    Private Sub btnExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        GroupBox1.Visible = True
        DataGridView1.DataSource = Nothing
        DataGridView1.RefreshEdit()
    End Sub
    Private Sub FillDataGridView(ByVal Query As String)

        da = New OleDbDataAdapter(Query, cn)
        dt.Clear()
        da.Fill(dt)
        ' ITEM, DESCRIPT, QTY, Actual, Diff, UNIT, GW_Kg, Total_GW, CBM_M3, Total_CBM, Load_GW, Load_CBM
        With FG
            '.DataSource = dt

        End With
    End Sub
    Private Sub EXXX()
        Try

 
            Dim mXLS As String
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim DtSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            mXLS = Trim(txtFNm.Text)
            If Len(mXLS) = 0 Then Exit Sub
            'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='D:\Software\Exim\PL.xls';Extended Properties=Excel 8.0;")
            'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & mXLS & "';Extended Properties=Excel 8.0;")
            '========= Exell 2010 ລົງມາ ========================
            'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & mXLS & "';Extended Properties=Excel 12.0 Xml;")
            '=========  ========================
            If CheckBox2.Checked = True Then
                MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source='" & mXLS & " '; " & "Extended Properties=Excel 8.0;")
            Else
                MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & mXLS & "';Extended Properties=Excel 8.0;")
            End If

            'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.16.0;Data Source='" & mXLS & "';Extended Properties=Excel 16.0 xml;")

            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [" & txtSheet.Text & "$]", MyConnection)
            MyCommand.TableMappings.Add("Table", "Net-informations.com")

            DtSet = New System.Data.DataSet
            MyCommand.Fill(DtSet)
            DataGridView1.DataSource = DtSet.Tables(0)
            MyConnection.Close()
            'DataGridView1.Columns(0).HeaderText = "date" : DataGridView1.Columns(0).Width = "120"
            'DataGridView1.Columns(1).HeaderText = "ACCno     " : DataGridView1.Columns(1).Width = "120"
            'With cn1
            '    .Provider = "Microsoft.ACE.OLEDB.16.0"
            '    .ConnectionString = "Data Source=" & strfile & ";" & _
            '    "Extended Properties=""Excel 16.0 xml;HDR=No;IMEX=1;Readonly=True"""
            'End With

        Catch ex As Exception

            MessageBox.Show(ex.Message)
        End Try
        'Dim row As Integer = 0
        'For row = 0 To DataGridView1.RowCount - 2
        '    DataGridView1.Rows(row).Cells(0).Value = row + 1
        'Next
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
 

        Call EXXX()
        Exit Sub
        ''Try


        ''    Dim MyConnection As System.Data.OleDb.OleDbConnection
        ''    Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
        ''    Dim rvl As Boolean = False
        ''    Dim fBrowse As New OpenFileDialog
        ''    With fBrowse
        ''        .Filter = "Excel files(*.xlsx)|*.xlsx|All files (*.*)|*.*"
        ''        .FilterIndex = 1
        ''        .Title = "Import data from Excel file"
        ''    End With


        ''    Dim mXLS As String

        ''    mXLS = Trim(txtFNm.Text)
        ''    If Len(mXLS) = 0 Then Exit Sub
        ''    'MyConnection = New System.Data.OleDb.OleDbConnection("Dsn=Excel Files;dbq=D:\LHSETEST\BCELAccStatement_093110000869828001_04-01-2019_04-01-2019.xlsx;defaultdir=D:\LHSETEST;driverid=1046;maxbuffersize=2048;pagetimeout=5")
        ''    MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source='" & mXLS & " '; " & "Extended Properties=Excel 8.0;")
        ''    'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & mXLS & "';Extended Properties=Excel 8.0;")
        ''    'MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection)
        ''    '  MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from test", MyConnection)
        ''    MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [" & txtSheet.Text & "$]", MyConnection)
        ''    'MyCommand.TableMappings.Add("Table", "Test")
        ''    MyCommand.TableMappings.Add("Table", "Net-informations.com")
        ''    DtSet = New System.Data.DataSet
        ''    'If rvl = True Then DtSet.Tables(0).Rows.Clear()
        ''    MyCommand.Fill(DtSet)
        ''    'If DtSet.Tables(0).Rows.Count > 0 Then
        ''    '    rvl = True
        ''    'Else
        ''    '    rvl = False
        ''    'End If
        ''    'MyConnection.Close()
        ''    With DataGridView1
        ''        .DataSource = DtSet.Tables(0)
        ''        .Refresh()
        ''    End With

        ''Catch ex As Exception

        ''End Try

        '' '' ''Dim row As Integer = 0
        '' '' ''For row = 0 To DataGridView1.RowCount - 2
        '' '' ''    DataGridView1.Rows(row).Cells(0).Value = row + 1
        '' '' ''Next
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
        ' DataGridView equivalent of FG_SelChange
        If DataGridView1.CurrentCell IsNot Nothing Then
            DataGridView1.BeginEdit(False)
        End If
    End Sub
Private Sub KKK()
        Dim KK As String = "  insert into gen_jn ( date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, code_cr, ac_code, amount, amount_dr, amount_cr, curr, rate, net_amt, bank_no,  Com_id,  " & _
                     " Activity_id, Cat_ID, office_id,company, lock, my_lock, del ,AG ,  last_update, last_user, pc_nm,amt_USD_dr,amt_USD_cr) " & _
                      " select date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, code_cr, ac_code,  (amount_dr + amount_cr), amount_dr, amount_cr, curr, rate, amount_LAK,  bank_no,   Com_id,  " & _
                     " Activity_id, Cat_ID, office_id,office_id,  0, 0, 1,0,   Getdate()," & _
                    " N'" & Apostrophe(MUserName) & "'," & _
                     " N'" & Apostrophe(MDServerName) & "',amount_dr / rate ,amount_Cr / rate    from Tmp_Import  "
        DbHelper.ExecuteNonQuery(KK)

        Dim aa As String
        aa = "       update gen_jn set amount_dr=0 where amount_dr is null     " & _
  "   update gen_jn set amount_cr=0 where amount_cr is null     " & _
   "  update gen_jn set amt_dr=0 where amt_dr is null      " & _
  "   update  gen_jn set amt_cr=0 where amt_cr is null   " & _
  "  update gen_jn set amt_USD_dr=0 where amt_USD_dr is null   " & _
  "   update  gen_jn set amt_USD_cr=0 where amt_USD_cr is null     "
        DbHelper.ExecuteNonQuery(aa)

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
        aa = "   update gen_jn set amt_dr  =  amount_dr * Rate where  amount_dr >0 "
        DbHelper.ExecuteNonQuery(aa)
        aa = "   update gen_jn set amt_cr  =  amount_Cr * Rate where  amount_Cr >0   "
        DbHelper.ExecuteNonQuery(aa)

        DbHelper.ExecuteNonQuery("update gen_jn set gen_jn.ac_name=Acc_Code.name_L,gen_jn.ac_namee=Acc_Code.name_E from Acc_Code,gen_jn where gen_jn.Ac_Code=Acc_Code.Ac_Code and  gen_jn.ac_name is null ")
    End Sub
    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        Call amt()
        If CDbl(txtSumAmountDr.Text) <> CDbl(txtSumAmountCr.Text) Then MsgBox("ບັນຊີນີ້ບໍ່ທັນບໍ່ທັນດູນດ່ຽງ ບໍ່ສາມາດບັນທຶກໄດ້!", MsgBoxStyle.OkOnly) : Exit Sub
        Call SaveItems()
        'If CheckBox1.Checked = True Then
        Call KKK()
        'Else
        '    Call Insert_Gen_jn()
        'End If

        MsgBox("ບັນທຶກສຳເລັດ")
    End Sub
Private Sub SaveItems()
        Dim aa As String
        aa = " delete Tmp_Import  "
        DbHelper.ExecuteNonQuery(aa)
        
        Dim dt As DataTable
        dt = DbHelper.GetDataTable("Select * FROM Tmp_Import")
        
        For i = 0 To DataGridViewMain.Rows.Count - 1
            If dt.Rows.Count = 0 Then
                Dim dateValue As Date
                If Date.TryParse(GetStr(DataGridViewMain.Rows(i).Cells(1).Value), dateValue) Then
                    Dim sa As String = " INSERT INTO Tmp_Import (date_work,Referno,certify, cheque_no,book, descrip, descripe, code_dr, code_cr,ac_code,  " & _
                       " amount_dr,amount_cr,Curr,Rate, amount_LAK, amount_USD,bank_no, doner, Com_id, Activity_id,Cat_ID, office_id,office_nm,last_update, last_user,pc_nm) " & _
                    " VALUES (  '" & Format(dateValue, "yyyy-MM-dd") & "'," & _
                        " N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(2).Value)) & "'," & _
                        " N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(3).Value)) & "'," & _
                         " N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(4).Value)) & "'," & _
                          " N'" & Trim(Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(5).Value))) & "'," & _
                           " N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(6).Value)) & "'," & _
                            " N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(7).Value)) & "'," & _
                             " N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(8).Value)) & "'," & _
                              " N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(9).Value)) & "'," & _
                      " N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(8).Value) + GetStr(DataGridViewMain.Rows(i).Cells(9).Value)) & "'," & _
                      " " & CDbl(GetStr(DataGridViewMain.Rows(i).Cells(10).Value)) & ", " & _
                    " " & CDbl(GetStr(DataGridViewMain.Rows(i).Cells(11).Value)) & ", " & _
                     " N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(12).Value)) & "'," & _
                     " " & CDbl(GetStr(DataGridViewMain.Rows(i).Cells(13).Value)) & ", " & _
                     " " & CDbl(GetStr(DataGridViewMain.Rows(i).Cells(14).Value)) & ", " & _
                     " " & CDbl(GetStr(DataGridViewMain.Rows(i).Cells(15).Value)) & ", " & _
                       " N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(16).Value)) & "'," & _
                      " N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(17).Value)) & "'," & _
                       " N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(18).Value)) & "'," & _
                        " N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(19).Value)) & "'," & _
                         " N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(20).Value)) & "'," & _
                          " N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(21).Value)) & "'," & _
                           " N'" & Apostrophe(GetStr(DataGridViewMain.Rows(i).Cells(22).Value)) & "'," & _
                                " Getdate()," & _
                    " N'" & Apostrophe(MUserName) & "'," & _
                     " N'" & Apostrophe(MDServerName) & "') "
                    DbHelper.ExecuteNonQuery(sa)
                End If
            End If
        Next i
        
        ' Update account codes
        DbHelper.ExecuteNonQuery("update Tmp_Import set Tmp_Import.ac_code=Acc_Code.ac_code from Acc_Code,Tmp_Import where Acc_Code.ac_original=Tmp_Import.ac_code ")
        DbHelper.ExecuteNonQuery("update Tmp_Import set Tmp_Import.code_dr=Acc_Code.ac_code from Acc_Code,Tmp_Import where Acc_Code.ac_original=Tmp_Import.code_dr and Tmp_Import.code_dr<>''  ")
        DbHelper.ExecuteNonQuery("update Tmp_Import set Tmp_Import.code_cr=Acc_Code.ac_code from Acc_Code,Tmp_Import where Acc_Code.ac_original=Tmp_Import.code_cr and Tmp_Import.code_Cr<>'' ")
    End Sub

Private Sub Insert_Gen_jn()
        Frm_import_progress.Show()
        Dim aa As String
        Dim dt As DataTable
        
        For i = 0 To DataGridViewMain.Rows.Count - 1
            Dim dateValue As Date
            If Date.TryParse(GetStr(DataGridViewMain.Rows(i).Cells(1).Value), dateValue) Then
                Dim sk As String = "Select * FROM gen_jn where  certify=N'" & GetStr(DataGridViewMain.Rows(i).Cells(3).Value) & "'  and date_work='" & Format(dateValue, "yyyy-MM-dd") & "'  "
                dt = DbHelper.GetDataTable(sk)
                
                If dt.Rows.Count = 0 Then
                    aa = "   insert into gen_jn (date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, " & _
          "  code_cr,ac_code,ac_name, amount_dr, amount_cr, curr,Curr_i, rate, Rate_i,  bank_no, " & _
          "    don_id, Com_id, Activity_id, Cat_ID,  company, office_id,lock, my_lock, del ,AG  ,last_update, last_user,pc_nm,amount,amt_dr,amt_Cr,amt_USD_dr,amt_USD_cr) " & _
         "   select  date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr," & _
         "  code_cr,ac_code,descrip, amount_dr, amount_cr, curr, curr, rate, rate,  bank_no, " & _
         " doner, Com_id, Activity_id, Cat_id, office_id  ,office_id  ,0, 0, 1,0, Getdate()," & _
              " N'" & Apostrophe(MUserName) & "'," & _
              " N'" & Apostrophe(MDServerName) & "', (amount_dr + amount_cr),amount_dr,amount_Cr,amount_dr / rate ,amount_Cr / rate  " & _
       "    from   Tmp_Import    where  certify=N'" & GetStr(DataGridViewMain.Rows(i).Cells(3).Value) & "'  and date_work='" & Format(dateValue, "yyyy-MM-dd") & "'    order by certify  "
                    DbHelper.ExecuteNonQuery(aa)
                Else
                    aa = "delete gen_jn WHERE certify=N'" & GetStr(DataGridViewMain.Rows(i).Cells(3).Value) & "'  and date_work='" & Format(dateValue, "yyyy-MM-dd") & "'  "
                    DbHelper.ExecuteNonQuery(aa)
                    aa = "delete AP_ACC_Gen WHERE certify=N'" & GetStr(DataGridViewMain.Rows(i).Cells(3).Value) & "'  and date_work='" & Format(dateValue, "yyyy-MM-dd") & "'  "
                    DbHelper.ExecuteNonQuery(aa)
                    aa = "   insert into gen_jn (date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, " & _
                            "  code_cr,ac_code,ac_name, amount_dr, amount_cr, curr,Curr_i, rate, Rate_i,  bank_no, " & _
                            "    don_id, Com_id, Activity_id, Cat_ID,  company,office_id,lock, my_lock, del ,AG  ,last_update, last_user,pc_nm,amount,amt_dr,amt_Cr,amt_USD_dr,amt_USD_cr) " & _
                           "   select  date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr," & _
                           "  code_cr,ac_code,descrip, amount_dr, amount_cr, curr, curr, rate, rate,  bank_no, " & _
                           " doner, Com_id, Activity_id, Cat_id, office_id, office_id  ,0, 0, 1,0, Getdate()," & _
                                " N'" & Apostrophe(MUserName) & "'," & _
                                " N'" & Apostrophe(MDServerName) & "', (amount_dr + amount_cr),amount_dr,amount_Cr,amount_dr / rate ,amount_Cr / rate  " & _
                         "    from   Tmp_Import    where  certify=N'" & GetStr(DataGridViewMain.Rows(i).Cells(3).Value) & "'  and date_work='" & Format(dateValue, "yyyy-MM-dd") & "'    order by certify  "
                    DbHelper.ExecuteNonQuery(aa)
                End If
                Frm_import_progress.Refresh()

                Frm_import_progress.Label2.Text = DataGridViewMain.Rows.Count
                Frm_import_progress.Label4.Text = GetStr(DataGridViewMain.Rows(i).Cells(0).Value)
                Frm_import_progress.Label1.Text = GetStr(DataGridViewMain.Rows(i).Cells(2).Value)
            End If
        Next
        Frm_import_progress.Close()
        
        aa = "       update gen_jn set amount_dr=0 where amount_dr is null     " & _
              "   update gen_jn set amount_cr=0 where amount_cr is null     " & _
               "  update gen_jn set amt_dr=0 where amt_dr is null      " & _
              "   update  gen_jn set amt_cr=0 where amt_cr is null   " & _
              "  update gen_jn set amt_USD_dr=0 where amt_USD_dr is null   " & _
              "   update  gen_jn set amt_USD_cr=0 where amt_USD_cr is null     "
        DbHelper.ExecuteNonQuery(aa)

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