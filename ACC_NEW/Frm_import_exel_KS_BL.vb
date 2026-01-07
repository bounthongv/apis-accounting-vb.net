Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports ApPBank10.Module

Public Class Frm_import_exel_KS_BL
    Dim amt1 As Double
    Dim amt2 As Double
    Dim MLAK, MUSD, MTHB, MEUR As Double
    Dim MDr As Double
    Dim MCr As Double
    Dim EDr As Double
    Dim ECr As Double
    ' Remove legacy OleDb variables - using DbHelper instead
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
    'DtSet moved to local scope in EXXX method

    Private Sub Frm_import_exel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Initialize DataGridView columns
        InitializeDataGridViewColumns()
        
        dff()
        CheckBox2.Text = "V.S"
    End Sub
    
    Private Sub InitializeDataGridViewColumns()
        ' Setup DataGridView columns equivalent to FG.FormatString
        FG.Columns.Clear()
        FG.Columns.Add("Col0", "ລ/ດ")
        FG.Columns.Add("Col1", "ເລກບັນຊີ")
        FG.Columns.Add("Col2", "ຊື່ບັນຊີ")
        FG.Columns.Add("Col3", "Be Dr")
        FG.Columns.Add("Col4", "Be Cr")
        FG.Columns.Add("Col5", "Mov Dr")
        FG.Columns.Add("Col6", "Mov Cr")
        FG.Columns.Add("Col7", "End Dr")
        FG.Columns.Add("Col8", "End Cr")
        
        ' Set column widths and styles
        FG.Columns("Col0").Width = 50
        FG.Columns("Col1").Width = 100
        FG.Columns("Col2").Width = 200
        FG.Columns("Col3").Width = 100
        FG.Columns("Col4").Width = 100
        FG.Columns("Col5").Width = 100
        FG.Columns("Col6").Width = 100
        FG.Columns("Col7").Width = 100
        FG.Columns("Col8").Width = 100
        
        ' Set alignment for numeric columns
        For i As Integer = 3 To 8
            FG.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            FG.Columns(i).DefaultCellStyle.Format = "#,##0.00"
        Next
        
        ' Add row number column
        AddRowNumbers()
    End Sub
    
    Private Sub AddRowNumbers()
        For i As Integer = 0 To FG.Rows.Count - 1
            If i < FG.Rows.Count Then
                FG.Rows(i).Cells("Col0").Value = i + 1
            End If
        Next
    End Sub
    Private Sub dff()
        ' Modern database access using DbHelper
        Dim dt As DataTable = DbHelper.GetDataTable("select * from Ap_RateSeting where Curr='LAK'")
        If dt.Rows.Count > 0 Then
            MLAK = Convert.ToDouble(GetStr(dt.Rows(0)("Rate")))
        End If
        
        dt = DbHelper.GetDataTable("select * from Ap_RateSeting where Curr='USD'")
        If dt.Rows.Count > 0 Then
            MUSD = Convert.ToDouble(GetStr(dt.Rows(0)("Rate")))
        End If
        
        dt = DbHelper.GetDataTable("select * from Ap_RateSeting where Curr='THB'")
        If dt.Rows.Count > 0 Then
            MTHB = Convert.ToDouble(GetStr(dt.Rows(0)("Rate")))
        End If
        
        dt = DbHelper.GetDataTable("select * from Ap_RateSeting where Curr='EUR'")
        If dt.Rows.Count > 0 Then
            MEUR = Convert.ToDouble(GetStr(dt.Rows(0)("Rate")))
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
        Try
            ' Clear existing rows (keep header row)
            FG.Rows.Clear()
            
            ' Import data from DataGridView1 to FG (DataGridView)
            For i As Integer = 0 To DataGridView1.RowCount - 2
                If DataGridView1.Item(0, i).Value IsNot Nothing AndAlso DataGridView1.Item(0, i).Value.ToString <> "" Then
                    Dim rowValues As New List(Of Object)
                    rowValues.Add(FG.Rows.Count + 1) ' Row number
                    rowValues.Add(GetStr(DataGridView1.Item(0, i).Value)) ' Account code
                    rowValues.Add(GetStr(DataGridView1.Item(1, i).Value)) ' Account name
                    
                    ' Add numeric values safely
                    For col As Integer = 2 To 7
                        If DataGridView1.Item(col, i).Value IsNot Nothing Then
                            Dim val As Double
                            If Double.TryParse(DataGridView1.Item(col, i).Value.ToString, val) Then
                                rowValues.Add(val)
                            Else
                                rowValues.Add(0.0)
                            End If
                        Else
                            rowValues.Add(0.0)
                        End If
                    Next
                    
                    FG.Rows.Add(rowValues.ToArray())
                End If
            Next i
            
        Catch ex As Exception
            MsgBox("Error importing data: " & ex.Message)
        End Try
        
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
        
        ' Calculate totals from DataGridView
        For i As Integer = 0 To FG.Rows.Count - 1
            If FG.Rows(i).Cells("Col3").Value IsNot Nothing Then
                amt1 += Convert.ToDouble(FG.Rows(i).Cells("Col3").Value)
            End If
            If FG.Rows(i).Cells("Col4").Value IsNot Nothing Then
                amt2 += Convert.ToDouble(FG.Rows(i).Cells("Col4").Value)
            End If
            If FG.Rows(i).Cells("Col5").Value IsNot Nothing Then
                MDr += Convert.ToDouble(FG.Rows(i).Cells("Col5").Value)
            End If
            If FG.Rows(i).Cells("Col6").Value IsNot Nothing Then
                MCr += Convert.ToDouble(FG.Rows(i).Cells("Col6").Value)
            End If
            If FG.Rows(i).Cells("Col7").Value IsNot Nothing Then
                EDr += Convert.ToDouble(FG.Rows(i).Cells("Col7").Value)
            End If
            If FG.Rows(i).Cells("Col8").Value IsNot Nothing Then
                ECr += Convert.ToDouble(FG.Rows(i).Cells("Col8").Value)
            End If
        Next i

        txtSumAmountDr.Text = Format(amt1, "##,##0.00")
        txtSumAmountCr.Text = Format(amt2, "##,##0.00")
        TxtMdr.Text = Format(MDr, "##,##0.00")
        TxtMcr.Text = Format(MCr, "##,##0.00")
        TxtEdr.Text = Format(EDr, "##,##0.00")
        TxtEcr.Text = Format(ECr, "##,##0.00")
    End Sub

    Private Sub btnExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        GroupBox1.Visible = True
        DataGridView1.DataSource = Nothing
        DataGridView1.RefreshEdit()
    End Sub
    Private Sub FillDataGridView(ByVal Query As String)
        ' Modern approach using DbHelper
        Try
            Dim dt As DataTable = DbHelper.GetDataTable(Query)
            ' For legacy compatibility, you could still use OleDb for Excel files
            ' but for SQL Server data, use DbHelper
        Catch ex As Exception
            MsgBox("Error filling DataGridView: " & ex.Message)
        End Try
    End Sub
    Private Sub EXXX()
        Try
            Dim mXLS As String = Trim(txtFNm.Text)
            If String.IsNullOrEmpty(mXLS) Then Exit Sub
            
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim DtSet As System.Data.DataSet
            
            ' Modern Excel connection with better error handling
            If CheckBox2.Checked = True Then
                MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source='" & mXLS & " '; " & "Extended Properties=Excel 8.0;HDR=YES;")
            Else
                MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & mXLS & "';Extended Properties=Excel 8.0;HDR=YES;")
            End If

            ' Validate sheet name
            If String.IsNullOrEmpty(txtSheet.Text) Then
                MsgBox("Please enter sheet name", MsgBoxStyle.Exclamation)
                Exit Sub
            End If

            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [" & txtSheet.Text & "$]", MyConnection)
            MyCommand.TableMappings.Add("Table", "Net-informations.com")

            DtSet = New System.Data.DataSet
            MyCommand.Fill(DtSet)
            DataGridView1.DataSource = DtSet.Tables(0)
            MyConnection.Close()
            
            ' Auto-size columns for better visibility
            For Each column As DataGridViewColumn In DataGridView1.Columns
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            Next

        Catch ex As Exception
            MessageBox.Show("Error loading Excel file: " & ex.Message, "Excel Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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

    Private Sub FG_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelectionChanged
        ' DataGridView equivalent of SelChange
        FG.ReadOnly = False
        ' You can access selected cell values with:
        ' If FG.CurrentRow IsNot Nothing Then
        '     MsgBox(FG.CurrentRow.Cells("Col10").Value)
        ' End If
    End Sub
    Private Sub KKK()
        ' Modern database execution using DbHelper
        Dim KK As String = "  insert into gen_jn ( date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, code_cr, ac_code, amount, amount_dr, amount_cr, curr, rate, net_amt, bank_no,  Com_id,  " & _
                     " Activity_id, Cat_ID, office_id,company, lock, my_lock, del ,AG ,  last_update, last_user, pc_nm,amt_USD_dr,amt_USD_cr) " & _
                      " select date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, code_cr, ac_code,  (amount_dr + amount_cr), amount_dr, amount_cr, curr, rate, amount_LAK,  bank_no,   Com_id,  " & _
                     " Activity_id, Cat_ID, office_id,office_id,  0, 0, 1,0,   Getdate()," & _
                    " N'" & Apostrophe(MUserName) & "'," & _
                     " N'" & Apostrophe(MDServerName) & "',amount_dr / rate ,amount_Cr / rate    from Tmp_Import  "
        DbHelper.ExecuteNonQuery(KK)

        Dim aa As String = "       update gen_jn set amount_dr=0 where amount_dr is null     " & _
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
        If CDbl(txtSumAmountDr.Text) <> CDbl(txtSumAmountCr.Text) Then 
            MsgBox("ບັນຊີນີ້ບໍ່ທັນບໍ່ທັນດູນດ່ຽງ ບໍ່ສາມາດບັນທຶກໄດ້!", MsgBoxStyle.OkOnly) 
            Exit Sub
        End If
        
        ' Modern database check using DbHelper
        Dim sk As String = "Select * FROM Ap_balance_TB   where   month(date_work)='" & Month(MM.Value) & "' and year(date_work)='" & Year(MM.Value) & "'  "
        Dim dtCheck As DataTable = DbHelper.GetDataTable(sk)
        If dtCheck.Rows.Count > 0 Then
            MsgBox("ຂໍ້ມູນເດືອນ " & Format(CDate(MM.Value), "MM/yyyy") & " ມີແລ້ວ!", MsgBoxStyle.Exclamation) 
            Exit Sub
        Else
            Call SaveItems()
        End If
        
        DbHelper.ExecuteNonQuery("delete Ap_balance_TB WHERE ac_code='' ")
        MsgBox("ບັນທຶກສຳເລັດ")
    End Sub
    Private Sub SaveItems()
        Dim aa As String = " delete Tmp_Import  "
        DbHelper.ExecuteNonQuery(aa)
        
        ' Modern database access using DbHelper
        Dim sk As String = "Select * FROM Ap_balance_TB   where   month(date_work)='" & Month(MM.Value) & "' and year(date_work)='" & Year(MM.Value) & "'  "
        Dim dtCheck As DataTable = DbHelper.GetDataTable(sk)
        
        For i As Integer = 0 To FG.Rows.Count - 1
            If dtCheck.Rows.Count = 0 Then
                ' Safely get values from DataGridView cells
                Dim acCode As String = GetStr(FG.Rows(i).Cells("Col1").Value)
                Dim acName As String = GetStr(FG.Rows(i).Cells("Col2").Value)
                Dim openAmtDr As Double = GetSafeDouble(FG.Rows(i).Cells("Col3").Value)
                Dim openAmtCr As Double = GetSafeDouble(FG.Rows(i).Cells("Col4").Value)
                Dim amtDr As Double = GetSafeDouble(FG.Rows(i).Cells("Col5").Value)
                Dim amtCr As Double = GetSafeDouble(FG.Rows(i).Cells("Col6").Value)
                Dim remDr As Double = GetSafeDouble(FG.Rows(i).Cells("Col7").Value)
                Dim remCr As Double = GetSafeDouble(FG.Rows(i).Cells("Col8").Value)
                
                Dim sa As String = " INSERT INTO Ap_balance_TB (date_work, ac_code, ac_Name,   open_amt_dr, open_amt_cr, amt_dr, amt_cr, Rem_dr, Rem_cr,office_id,office_nm,last_update, last_user,pc_nm) " & _
                " VALUES ('" & Format(CDate(MM.Value), "yyyy-MM-dd") & "'," & _
                    " N'" & Apostrophe(acCode) & "'," & _
                    " N'" & Apostrophe(acName) & "'," & _
                     " " & openAmtDr & ", " & _
                      " " & openAmtCr & ", " & _
                        " " & amtDr & ", " & _
                  " " & amtCr & ", " & _
                " " & remDr & ", " & _
                   " " & remCr & ", " & _
                   " N'01-02',''," & _
                           " Getdate()," & _
                " N'" & Apostrophe(MUserName) & "'," & _
                 " N'" & Apostrophe(MDServerName) & "') "
                DbHelper.ExecuteNonQuery(sa)
            End If
        Next i
    End Sub
    
    ' Helper function to safely convert to Double
    Private Function GetSafeDouble(ByVal value As Object) As Double
        If value Is Nothing OrElse IsDBNull(value) Then
            Return 0.0
        End If
        Dim result As Double
        If Double.TryParse(value.ToString, result) Then
            Return result
        Else
            Return 0.0
        End If
    End Function

    Private Sub Insert_Gen_jn()
        Frm_import_progress.Show()
        Dim aa As String
        Dim i As Integer
        
        For i = 0 To FG.Rows.Count - 1
            ' Safely get values from DataGridView
            Dim certify As String = GetStr(FG.Rows(i).Cells("Col3").Value)
            Dim dateWork As String
            Try
                dateWork = Format(CDate(GetStr(FG.Rows(i).Cells("Col1").Value)), "yyyy-MM-dd")
            Catch
                dateWork = Format(CDate(MM.Value), "yyyy-MM-dd")
            End Try
            
            ' Modern database check using DbHelper
            Dim sk As String = "Select * FROM gen_jn where  certify=N'" & certify & "'  and date_work=N'" & dateWork & "'  "
            Dim dtCheck As DataTable = DbHelper.GetDataTable(sk)
            
            If dtCheck.Rows.Count = 0 Then
                aa = "   insert into gen_jn (date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, " & _
          "  code_cr,ac_code,ac_name, amount_dr, amount_cr, curr,Curr_i, rate, Rate_i,  bank_no, " & _
          "    don_id, Com_id, Activity_id, Cat_ID,  company, office_id,lock, my_lock, del ,AG  ,last_update, last_user,pc_nm,amount,amt_dr,amt_Cr,amt_USD_dr,amt_USD_cr) " & _
         "   select  date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr," & _
         "  code_cr,ac_code,descrip, amount_dr, amount_cr, curr, curr, rate, rate,  bank_no, " & _
         " doner, Com_id, Activity_id, Cat_id, office_id  ,office_id  ,0, 0, 1,0, Getdate()," & _
              " N'" & Apostrophe(MUserName) & "'," & _
              " N'" & Apostrophe(MDServerName) & "', (amount_dr + amount_cr),amount_dr,amount_Cr,amount_dr / rate ,amount_Cr / rate  " & _
       "    from   Tmp_Import    where  certify=N'" & certify & "'  and date_work=N'" & dateWork & "'    order by certify  "
                DbHelper.ExecuteNonQuery(aa)
            Else
                ' Delete existing records and re-insert
                aa = "delete gen_jn WHERE certify=N'" & certify & "'  and date_work=N'" & dateWork & "'  "
                DbHelper.ExecuteNonQuery(aa)
                aa = "delete AP_ACC_Gen WHERE certify=N'" & certify & "'  and date_work=N'" & dateWork & "'  "
                DbHelper.ExecuteNonQuery(aa)
                aa = "   insert into gen_jn (date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr, " & _
                        "  code_cr,ac_code,ac_name, amount_dr, amount_cr, curr,Curr_i, rate, Rate_i,  bank_no, " & _
                        "    don_id, Com_id, Activity_id, Cat_ID,  company,office_id,lock, my_lock, del ,AG  ,last_update, last_user,pc_nm,amount,amt_dr,amt_Cr,amt_USD_dr,amt_USD_cr) " & _
                       "   select  date_work, Referno, certify, cheque_no, book, descrip, descripe, code_dr," & _
                       "  code_cr,ac_code,descrip, amount_dr, amount_cr, curr, curr, rate, rate,  bank_no, " & _
                       " doner, Com_id, Activity_id, Cat_id, office_id, office_id  ,0, 0, 1,0, Getdate()," & _
                            " N'" & Apostrophe(MUserName) & "'," & _
                            " N'" & Apostrophe(MDServerName) & "', (amount_dr + amount_cr),amount_dr,amount_Cr,amount_dr / rate ,amount_Cr / rate  " & _
                     "    from   Tmp_Import    where  certify=N'" & certify & "'  and date_work=N'" & dateWork & "'    order by certify  "
                DbHelper.ExecuteNonQuery(aa)
            End If
            
            ' Update progress
            Frm_import_progress.Refresh()
            Frm_import_progress.Label2.Text = FG.Rows.Count.ToString()
            Frm_import_progress.Label4.Text = GetStr(FG.Rows(i).Cells("Col0").Value)
            Frm_import_progress.Label1.Text = GetStr(FG.Rows(i).Cells("Col2").Value)

        Next
        Frm_import_progress.Close()
    End Sub
    
    ' Additional cleanup method after data insertion
    Private Sub CleanupAfterInsertion()
        ' Using DbHelper instead of CNN.Execute
        Dim aa As String = "       update gen_jn set amount_dr=0 where amount_dr is null     " & _
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Call amt()
        Catch ex As Exception
            MsgBox("Error calculating amounts: " & ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ' Modern database access using DbHelper
        Dim sk As String = "Select * FROM Ap_balance_TB   where   month(date_work)='" & Month(MM.Value) & "' and year(date_work)='" & Year(MM.Value) & "'  "
        Dim dtCheck As DataTable = DbHelper.GetDataTable(sk)
        If dtCheck.Rows.Count = 0 Then
            MsgBox("ຂໍ້ມູນເດືອນ " & Format(CDate(MM.Value), "MM/yyyy") & " ບໍ່ທັນມີ!", MsgBoxStyle.Exclamation) 
            Exit Sub
        End If

        If MessageBox.Show("ທ່ານຕ້ອງລຶບຂໍ້ມູນ  " & Format(CDate(MM.Value), "MM/yyyy") & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            DbHelper.ExecuteNonQuery("DELETE FROM Ap_balance_TB   where   month(date_work)='" & Month(MM.Value) & "' and year(date_work)='" & Year(MM.Value) & "' ")
            MsgBox("Finish")
        End If
    End Sub
    
    ' Helper methods for DataGridView data access
    Private Function GetCellValue(ByVal row As Integer, ByVal colName As String) As Object
        If row >= 0 AndAlso row < FG.Rows.Count AndAlso FG.Columns.Contains(colName) Then
            Return FG.Rows(row).Cells(colName).Value
        End If
        Return Nothing
    End Function
    
    Private Function GetCellText(ByVal row As Integer, ByVal colName As String) As String
        Dim value As Object = GetCellValue(row, colName)
        Return GetStr(value)
    End Function
    
    Private Function GetCellDouble(ByVal row As Integer, ByVal colName As String) As Double
        Dim value As Object = GetCellValue(row, colName)
        Return GetSafeDouble(value)
    End Function
    
    ' DataGridView event for editing
    Private Sub FG_CellEndEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles FG.CellEndEdit
        Try
            ' Recalculate totals after editing
            Call amt()
        Catch ex As Exception
            ' Handle error silently or log it
        End Try
    End Sub
    
    ' DataGridView event for data validation
    Private Sub FG_CellValidating(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles FG.CellValidating
        ' Validate numeric columns
        If e.ColumnIndex >= 3 AndAlso e.ColumnIndex <= 8 Then
            Dim newValue As String = e.FormattedValue.ToString()
            If Not String.IsNullOrEmpty(newValue) Then
                Dim result As Double
                If Not Double.TryParse(newValue, result) Then
                    e.Cancel = True
                    MsgBox("Please enter a valid number", MsgBoxStyle.Exclamation)
                End If
            End If
        End If
    End Sub
End Class