Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text

Public Class Frm_import_exel_KS_BL
    Dim amt1 As Double
    Dim amt2 As Double
    Dim MLAK, MUSD, MTHB, MEUR As Double
    Dim MDr As Double
    Dim MCr As Double
    Dim EDr As Double
    Dim ECr As Double
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
    'Dim percentOfGamesWon As Double = (gamesWon + gamesLost) * gamesWon / 100%
    Dim DtSet As System.Data.DataSet

    Private Sub Frm_import_exel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        FG.FormatString = "^ລ/ດ |<ເລກບັນຊີ        |<ຊື່ບັນຊີ                                    |> Be Dr           |>Be Cr              |>Mov Dr            |>Mov Cr          |>End Dr          |>End Cr            "

        dff()
        CheckBox2.Text = "V.S"
        'FG.set_ColHidden(3, True)
        'FG.set_ColHidden(4, True)
        'FG.set_ColHidden(5, True)
        'FG.set_ColHidden(6, True)
        'FG.set_ColHidden(7, True)
        'FG.set_ColHidden(8, True)
        'FG.set_ColHidden(9, True)
    End Sub
    Private Sub dff()
        LoadSqlData("select * from Ap_RateSeting where Curr='LAK' ", RSC)
        If RSC.RecordCount <> 0 Then
            MLAK = Trim(RSC.Fields("Rate").Value)
        End If
        LoadSqlData("select * from Ap_RateSeting where Curr='USD' ", RSC)
        If RSC.RecordCount <> 0 Then
            MUSD = Trim(RSC.Fields("Rate").Value)
        End If
        LoadSqlData("select * from Ap_RateSeting where Curr='THB' ", RSC)
        If RSC.RecordCount <> 0 Then
            MTHB = Trim(RSC.Fields("Rate").Value)
        End If
        LoadSqlData("select * from Ap_RateSeting where Curr='EUR' ", RSC)
        If RSC.RecordCount <> 0 Then
            MEUR = Trim(RSC.Fields("Rate").Value)
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

        FG.Rows = 1
        Dim CU As String
        Try


            For i = 0 To DataGridView1.RowCount - 2
                If DataGridView1.Item(0, i).Value.ToString <> "" Then

                    'FG.AddItem(DataGridView1.Item(0, i).Value.ToString & Chr(9) & DataGridView1.Item(1, i).Value.ToString)
                    CNN.CommandTimeout = 0
                    FG.AddItem(DataGridView1.Item(0, i).Value.ToString & _
                                 Chr(9) & DataGridView1.Item(0, i).Value.ToString & _
                               Chr(9) & DataGridView1.Item(1, i).Value.ToString & _
                           Chr(9) & Format(CDbl(DataGridView1.Item(2, i).Value), "#,##0.00") & _
                         Chr(9) & Format(CDbl(DataGridView1.Item(3, i).Value), "#,##0.00") & _
                             Chr(9) & Format(CDbl(DataGridView1.Item(4, i).Value), "#,##0.00") & _
                                   Chr(9) & Format(CDbl(DataGridView1.Item(5, i).Value), "#,##0.00") & _
                       Chr(9) & Format(CDbl(DataGridView1.Item(6, i).Value), "#,##0.00") & _
                           Chr(9) & Format(CDbl(DataGridView1.Item(7, i).Value), "#,##0.00"))
                    'FG.set_TextMatrix(FG.Row, no, no)
                End If
                'CU = Chr(9) & Microsoft.VisualBasic.Right(DataGridView1.Item(8, i).Value.ToString, 4)

            Next i
        Catch ex As Exception
            MsgBox(ex)
        End Try

        For i = 1 To FG.Rows - 1
            no = no + 1
            FG.set_TextMatrix(i, 0, no)

 
            'FG.set_TextMatrix(i, 12, Microsoft.VisualBasic.Left(FG.get_TextMatrix(i, 12), 3))
            'If Trim(FG.get_TextMatrix(i, 12)) = "LAK" Then
            '    FG.set_TextMatrix(i, 13, 1)

            'ElseIf Trim(FG.get_TextMatrix(i, 12)) = "USD" Then
            '    FG.set_TextMatrix(i, 13, MUSD)
            'ElseIf Trim(FG.get_TextMatrix(i, 12)) = "THB" Then
            '    FG.set_TextMatrix(i, 13, MTHB)
            'ElseIf Trim(FG.get_TextMatrix(i, 12)) = "EUR" Then
            '    FG.set_TextMatrix(i, 13, MEUR)
            'End If

            'FG.set_TextMatrix(i, 13, Format(CDbl(FG.get_TextMatrix(i, 13)), "#,##0.00"))
            'FG.set_TextMatrix(i, 14, Format(CDbl(FG.get_TextMatrix(i, 10)) + CDbl(FG.get_TextMatrix(i, 11)), "#,##0.00"))
            'FG.set_TextMatrix(i, 15, Format(CDbl(FG.get_TextMatrix(i, 10)) + CDbl(FG.get_TextMatrix(i, 11)), "#,##0.00"))
            'FG.set_TextMatrix(i, 11, Format(CDbl(FG.get_TextMatrix(i, 11)), "#,##0.00"))
            'FG.set_TextMatrix(i, 13, Format(CDbl(FG.get_TextMatrix(i, 13)), "#,##0.00"))
            'FG.set_TextMatrix(i, 14, Format(CDbl(FG.get_TextMatrix(i, 14)), "#,##0.00"))
            'FG.set_TextMatrix(i, 15, Format(CDbl(FG.get_TextMatrix(i, 15)), "#,##0.00"))

            'FG.set_TextMatrix(i, no)

        Next
        'Call Calc()
        GroupBox1.Visible = False
        Call amt()
    End Sub
    Private Sub amt()
        amt1 = 0
        amt2 = 0
        MDr = 0
        MCr = 0
        EDr = 0
        ECr = 0
        Dim i As Integer
        For i = 1 To FG.Rows - 1
            amt1 = amt1 + CDbl(FG.get_TextMatrix(i, 3))
            amt2 = amt2 + CDbl(FG.get_TextMatrix(i, 4))
            MDr = MDr + CDbl(FG.get_TextMatrix(i, 5))
            MCr = MCr + CDbl(FG.get_TextMatrix(i, 6))
            EDr = EDr + CDbl(FG.get_TextMatrix(i, 7))
            ECr = ECr + CDbl(FG.get_TextMatrix(i, 8))
        Next i

        txtSumAmountDr.Text = Format(CDbl(amt1), "##,##0.00")
        txtSumAmountCr.Text = Format(CDbl(amt1), "##,##0.00")
        '======
        TxtMdr.Text = Format(CDbl(MDr), "##,##0.00")
        TxtMcr.Text = Format(CDbl(MCr), "##,##0.00")
        '=========
        TxtEdr.Text = Format(CDbl(EDr), "##,##0.00")
        TxtEcr.Text = Format(CDbl(ECr), "##,##0.00")
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

    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelChange
        FG.Editable = VSFlex8U.EditableSettings.flexEDKbd
        'MsgBox(FG.get_TextMatrix(FG.Row, 10))
        'MsgBox(FG.get_TextMatrix(FG.Row, 11))
        'MsgBox(FG.get_TextMatrix(FG.Row, 12))
        'MsgBox(FG.get_TextMatrix(FG.Row, 13))
        'MsgBox(FG.get_TextMatrix(FG.Row, 14))
        'MsgBox(FG.get_TextMatrix(FG.Row, 15))
    End Sub
    Private Sub KKK()
        Dim KK As String = "  insert into gen_jn ( date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, code_cr, ac_code, amount, amount_dr, amount_cr, curr, rate, net_amt, bank_no,  Com_id,  " & _
                     " Activity_id, Cat_ID, office_id,company, lock, my_lock, del ,AG ,  last_update, last_user, pc_nm,amt_USD_dr,amt_USD_cr) " & _
                      " select date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, code_cr, ac_code,  (amount_dr + amount_cr), amount_dr, amount_cr, curr, rate, amount_LAK,  bank_no,   Com_id,  " & _
                     " Activity_id, Cat_ID, office_id,office_id,  0, 0, 1,0,   Getdate()," & _
                    " N'" & Apostrophe(MUserName) & "'," & _
                     " N'" & Apostrophe(MDServerName) & "',amount_dr / rate ,amount_Cr / rate    from Tmp_Import  "
        CNN.Execute(KK)


        Dim aa As String
        aa = "       update gen_jn set amount_dr=0 where amount_dr is null     " & _
  "   update gen_jn set amount_cr=0 where amount_cr is null     " & _
   "  update gen_jn set amt_dr=0 where amt_dr is null      " & _
  "   update  gen_jn set amt_cr=0 where amt_cr is null   " & _
  "  update gen_jn set amt_USD_dr=0 where amt_USD_dr is null   " & _
  "   update  gen_jn set amt_USD_cr=0 where amt_USD_cr is null     "
        CNN.Execute(aa)

        aa = "   update AP_ACC_Gen set net_usd  =net_amt / rate   where curr='LAK'  "
        CNN.Execute(aa)
        aa = "   update AP_ACC_Gen set net_usd  =net_amt   where curr='USD'  "
        CNN.Execute(aa)

        aa = "   update AP_ACC_Gen set cheque_no  = gen_jn.cheque_no from gen_jn where  AP_ACC_Gen.certify=gen_jn.certify  and AP_ACC_Gen.cheque_no is null "
        CNN.Execute(aa)

        aa = "   update AP_ACC_Gen set Com_id =office_id " & _
          "    update gen_jn set Com_id =office_id  " & _
            " update gen_jn set don_id  ='01'  " & _
                "  update AP_ACC_Gen set don_id  ='01' "
        CNN.Execute(aa)
        aa = "   update gen_jn set amt_dr  =  amount_dr * Rate where  amount_dr >0 "
        CNN.Execute(aa)
        aa = "   update gen_jn set amt_cr  =  amount_Cr * Rate where  amount_Cr >0   "
        CNN.Execute(aa)

        CNN.Execute("update gen_jn set gen_jn.ac_name=Acc_Code.name_L,gen_jn.ac_namee=Acc_Code.name_E from Acc_Code,gen_jn where gen_jn.Ac_Code=Acc_Code.Ac_Code and  gen_jn.ac_name is null ")

    End Sub
    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        Call amt()
        If CDbl(txtSumAmountDr.Text) <> CDbl(txtSumAmountCr.Text) Then MsgBox("ບັນຊີນີ້ບໍ່ທັນບໍ່ທັນດູນດ່ຽງ ບໍ່ສາມາດບັນທຶກໄດ້!", MsgBoxStyle.OkOnly) : Exit Sub
        Dim sk As String = "Select * FROM Ap_balance_TB   where   month(date_work)='" & Month(MM.Value) & "' and year(date_work)='" & Year(MM.Value) & "'  "
        Call LoadSqlData(sk, RSC)
        If RSC.RecordCount <> 0 Then
            MsgBox("ຂໍ້ມູນເດືອນ " & Format(CDate(MM.Value), "MM/yyyy") & " ມີແລ້ວ!", MsgBoxStyle.Exclamation) : Exit Sub
        Else
            Call SaveItems()
        End If
        CNN.Execute("delete Ap_balance_TB WHERE ac_code='' ")
        'If CheckBox1.Checked = True Then
        'Call KKK()
        'Else
        '    Call Insert_Gen_jn()
        'End If

        MsgBox("ບັນທຶກສຳເລັດ")
    End Sub
    Private Sub SaveItems()
        Dim aa As String
        aa = " delete Tmp_Import  "
        CNN.Execute(aa)
        Dim Rschk As New ADODB.Recordset
        Dim i As Integer
        ', office_id,office_nm,last_update, last_user,pc_nm
        With Rschk
            Dim sk As String = "Select * FROM Ap_balance_TB   where   month(date_work)='" & Month(MM.Value) & "' and year(date_work)='" & Year(MM.Value) & "'  "
            Call LoadSqlData(sk, Rschk)
            For i = 1 To FG.Rows - 1
                If Rschk.RecordCount = 0 Then
                    Dim sa As String = " INSERT INTO Ap_balance_TB (date_work, ac_code, ac_Name,   open_amt_dr, open_amt_cr, amt_dr, amt_cr, Rem_dr, Rem_cr,office_id,office_nm,last_update, last_user,pc_nm) " & _
                    " VALUES ('" & Format(CDate(MM.Value), "yyyy-MM-dd") & "'," & _
                        " N'" & Apostrophe(FG.get_TextMatrix(i, 1)) & "'," & _
                        " N'" & Apostrophe(FG.get_TextMatrix(i, 2)) & "'," & _
                         " " & CDbl(FG.get_TextMatrix(i, 3)) & ", " & _
                          " " & CDbl(FG.get_TextMatrix(i, 4)) & ", " & _
                            " " & CDbl(FG.get_TextMatrix(i, 5)) & ", " & _
                      " " & CDbl(FG.get_TextMatrix(i, 6)) & ", " & _
                    " " & CDbl(FG.get_TextMatrix(i, 7)) & ", " & _
                       " " & CDbl(FG.get_TextMatrix(i, 8)) & ", " & _
                       " N'01-02',''," & _
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
        'CNN.Execute("delete gen_jn WHERE ac_code='' ")
    End Sub

    Private Sub Insert_Gen_jn()
        Frm_import_progress.Show()
        Dim aa As String
        Dim RSC As New ADODB.Recordset
        Dim Rschk As New ADODB.Recordset
        Dim i As Integer
        With Rschk
            For i = 1 To FG.Rows - 1
                Dim sk As String = "Select * FROM gen_jn where  certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'  "
                Call LoadSqlData(sk, Rschk)
                If Rschk.RecordCount = 0 Then
                    aa = "   insert into gen_jn (date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, " & _
          "  code_cr,ac_code,ac_name, amount_dr, amount_cr, curr,Curr_i, rate, Rate_i,  bank_no, " & _
          "    don_id, Com_id, Activity_id, Cat_ID,  company, office_id,lock, my_lock, del ,AG  ,last_update, last_user,pc_nm,amount,amt_dr,amt_Cr,amt_USD_dr,amt_USD_cr) " & _
         "   select  date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr," & _
         "  code_cr,ac_code,descrip, amount_dr, amount_cr, curr, curr, rate, rate,  bank_no, " & _
         " doner, Com_id, Activity_id, Cat_id, office_id  ,office_id  ,0, 0, 1,0, Getdate()," & _
              " N'" & Apostrophe(MUserName) & "'," & _
              " N'" & Apostrophe(MDServerName) & "', (amount_dr + amount_cr),amount_dr,amount_Cr,amount_dr / rate ,amount_Cr / rate  " & _
       "    from   Tmp_Import    where  certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    order by certify  "
                    CNN.Execute(aa)


                    ' ''aa = "  update gen_jn set Curr_i = curr where   certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    " & _
                    ' '' " update gen_jn set Rate_i  = Rate where   certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    "
                    ' ''CNN.Execute(aa)

                    ''aa = "   update gen_jn set amount =amount_dr + amount_cr where   certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'   " & _
                    ''"  update gen_jn set amt_dr  =amount_dr where   certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'   " & _
                    ''" update gen_jn set amt_cr  =amount_cr  where   certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'     "
                    ''CNN.Execute(aa)

                    ' ''aa = "  update gen_jn set  rate =1 where  rate =0 and   certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    "
                    ' ''CNN.Execute(aa)

                    '' aa = "   update gen_jn set amt_USD_dr  =amount_dr / rate  where   certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    " & _
                    ''"  update gen_jn set amt_USD_cr  =amount_cr  /rate   where   certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'      "
                    '' CNN.Execute(aa)

                    'aa = "       update gen_jn set amount_dr=0 where amount_dr is null and   certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    " & _
                    '"   update gen_jn set amount_cr=0 where amount_cr is null and   certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'     " & _
                    ' "  update gen_jn set amt_dr=0 where amt_dr is null and   certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    " & _
                    '"   update  gen_jn set amt_cr=0 where amt_cr is null and   certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    " & _
                    '"  update gen_jn set amt_USD_dr=0 where amt_USD_dr is null and   certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    " & _
                    '"   update  gen_jn set amt_USD_cr=0 where amt_USD_cr is null and   certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'     "
                    'CNN.Execute(aa)
                    '   aa = ""
                    '   aa = "       insert into AP_ACC_Gen ( date_work, Referno, certify,Book ,   descrip, descripe,amount,net_amt , " & _
                    ' "  AmountDr, AmountCr,TotalAmountDr,TotalAmountCr,curr,rate,Rate_USD,bank_no,don_id,Com_id,office_id)    " & _
                    '"  select     date_work, Referno, certify,Book,    '', '', " & _
                    '  "  sum(amount_dr) as amount , sum(amount_dr) as net_amt , sum(amount_dr) as amount_dr , " & _
                    '  "   sum(amount_cr) as amount_cr , sum(amount_dr) as TotalAmountDr , sum(amount_cr) as TotalAmountCr , " & _
                    ' "   curr,rate,rate,bank_no,don_id,Com_id,  office_id from gen_jn where  1=1    and    certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    " & _
                    '  "    group by    date_work, Referno, certify,Book,curr,rate,bank_no, " & _
                    '  "    don_id, Com_id,  office_id    "
                    '   CNN.Execute(aa)
                Else
                    'Dim sk As String = "Select * FROM gen_jn where  certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'  "
                    'Call LoadSqlData(sk, Rschk)
                    aa = "delete gen_jn WHERE certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work='" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'  "
                    CNN.Execute(aa)
                    aa = "delete AP_ACC_Gen WHERE certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work='" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'  "
                    CNN.Execute(aa)
                    aa = "   insert into gen_jn (date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, " & _
                            "  code_cr,ac_code,ac_name, amount_dr, amount_cr, curr,Curr_i, rate, Rate_i,  bank_no, " & _
                            "    don_id, Com_id, Activity_id, Cat_ID,  company,office_id,lock, my_lock, del ,AG  ,last_update, last_user,pc_nm,amount,amt_dr,amt_Cr,amt_USD_dr,amt_USD_cr) " & _
                           "   select  date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr," & _
                           "  code_cr,ac_code,descrip, amount_dr, amount_cr, curr, curr, rate, rate,  bank_no, " & _
                           " doner, Com_id, Activity_id, Cat_id, office_id, office_id  ,0, 0, 1,0, Getdate()," & _
                                " N'" & Apostrophe(MUserName) & "'," & _
                                " N'" & Apostrophe(MDServerName) & "', (amount_dr + amount_cr),amount_dr,amount_Cr,amount_dr / rate ,amount_Cr / rate  " & _
                         "    from   Tmp_Import    where  certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    order by certify  "
                    CNN.Execute(aa)

                    'aa = " update gen_jn set don_id= AP_Donnor.Don_ID  from AP_Donnor where  gen_jn.don_id= AP_Donnor.Don_Sym   "
                    'CNN.Execute(aa)
                    'aa = "  update gen_jn set office_id= AP_Office.off_id  from AP_Office  where  gen_jn.office_id= AP_Office.office_id   "
                    'CNN.Execute(aa)

                    '' aa = "  update gen_jn set Curr_i = curr where   certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    " & _
                    ''  " update gen_jn set Rate_i  = Rate where   certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    "
                    '' CNN.Execute(aa)


                    '' aa = "   update gen_jn set amount =amount_dr + amount_cr where  certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'   " & _
                    '' "  update gen_jn set amt_dr  =amount_dr where  certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    " & _
                    '' " update gen_jn set amt_cr  =amount_cr  where  certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    "
                    '' CNN.Execute(aa)

                    '' aa = "  update gen_jn set  rate =1 where  rate =0 and  certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    "
                    '' CNN.Execute(aa)

                    '' aa = "   update gen_jn set amt_USD_dr  =amount_dr / rate where  certify=N'" & FG.get_TextMatrix(i, 3) & "'   and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'     " & _
                    ''"  update gen_jn set amt_USD_cr  =amount_cr  /rate   where  certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'     "
                    '' CNN.Execute(aa)

                    'aa = "       update gen_jn set amount_dr=0 where amount_dr is null and  certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    " & _
                    '"   update gen_jn set amount_cr=0 where amount_cr is null and  certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'     " & _
                    ' "  update gen_jn set amt_dr=0 where amt_dr is null and  certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    " & _
                    '"   update  gen_jn set amt_cr=0 where amt_cr is null and  certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    " & _
                    '"  update gen_jn set amt_USD_dr=0 where amt_USD_dr is null and  certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    " & _
                    '"   update  gen_jn set amt_USD_cr=0 where amt_USD_cr is null  and  certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    "
                    'CNN.Execute(aa)

                    '   aa = "       insert into AP_ACC_Gen ( date_work, Referno, certify,Book ,   descrip, descripe,amount,net_amt , " & _
                    ' "  AmountDr, AmountCr,TotalAmountDr,TotalAmountCr,curr,rate,Rate_USD,bank_no,don_id,Com_id,office_id)    " & _
                    '"  select     date_work, Referno, certify,Book,    '', '', " & _
                    '  "  sum(amount_dr) as amount , sum(amount_dr) as net_amt , sum(amount_dr) as amount_dr , " & _
                    '  "   sum(amount_cr) as amount_cr , sum(amount_dr) as TotalAmountDr , sum(amount_cr) as TotalAmountCr , " & _
                    ' "   curr,rate,rate,bank_no,don_id,Com_id,  office_id from gen_jn where  1=1    and   certify=N'" & FG.get_TextMatrix(i, 3) & "'  and date_work=N'" & Format(CDate(FG.get_TextMatrix(i, 1)), "yyyy-MM-dd") & "'    " & _
                    '  "    group by    date_work, Referno, certify,Book,curr,rate,bank_no, " & _
                    '  "    don_id, Com_id,  office_id    "
                    '   CNN.Execute(aa)

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
        aa = "       update gen_jn set amount_dr=0 where amount_dr is null     " & _
              "   update gen_jn set amount_cr=0 where amount_cr is null     " & _
               "  update gen_jn set amt_dr=0 where amt_dr is null      " & _
              "   update  gen_jn set amt_cr=0 where amt_cr is null   " & _
              "  update gen_jn set amt_USD_dr=0 where amt_USD_dr is null   " & _
              "   update  gen_jn set amt_USD_cr=0 where amt_USD_cr is null     "
        CNN.Execute(aa)

        aa = "   update AP_ACC_Gen set net_usd  =net_amt / rate   where curr='LAK'  "
        CNN.Execute(aa)
        aa = "   update AP_ACC_Gen set net_usd  =net_amt   where curr='USD'  "
        CNN.Execute(aa)

        aa = "   update AP_ACC_Gen set cheque_no  = gen_jn.cheque_no from gen_jn where  AP_ACC_Gen.certify=gen_jn.certify  and AP_ACC_Gen.cheque_no is null "
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

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call amt()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim sk As String = "Select * FROM Ap_balance_TB   where   month(date_work)='" & Month(MM.Value) & "' and year(date_work)='" & Year(MM.Value) & "'  "
        Call LoadSqlData(sk, RSC)
        If RSC.RecordCount = 0 Then
            MsgBox("ຂໍ້ມູນເດືອນ " & Format(CDate(MM.Value), "MM/yyyy") & " ບໍ່ທັນມີ!", MsgBoxStyle.Exclamation) : Exit Sub
        End If

        If MessageBox.Show("ທ່ານຕ້ອງລຶບຂໍ້ມູນ  " & Format(CDate(MM.Value), "MM/yyyy") & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            CNN.Execute("DELETE FROM Ap_balance_TB   where   month(date_work)='" & Month(MM.Value) & "' and year(date_work)='" & Year(MM.Value) & "' ")
            MsgBox("Finish")
        End If



    End Sub
End Class