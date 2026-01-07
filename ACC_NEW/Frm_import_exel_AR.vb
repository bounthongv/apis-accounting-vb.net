Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Data

Public Class Frm_import_exel_AR
    Dim amt1 As Double
    Dim amt2 As Double
    Dim MLAK, MUSD, MTHB, MEUR As Double
    Private Sub Frm_import_exel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        FG.Columns.Clear()
        FG.Columns.Add("No", "ລ/ດ")
        FG.Columns.Add("LoanNO", "LOAN ACCOUNT")
        FG.Columns.Add("LoanName", "CONTRACT NAME")
        FG.Columns.Add("LoanDate", "LOAN OPEN DATE")
        FG.Columns.Add("Curr", "CURRENCY")
        FG.Columns.Add("FIX_RATE", "FIX RATE")
        FG.Columns.Add("PRINCIPLE", "AR OS TG PRINCIPLE")
        FG.Columns.Add("INTEREST", "AR OS TG ARCURED INTEREST")
        FG.Columns.Add("GENDER", "GENDER")
        FG.Columns.Add("BUSINESSTYPEDESC", "BUSINESSTYPEDESC")
        FG.Columns.Add("LOAN_GRADE", "LOAN GRADE")
        FG.Columns.Add("Provision", "Provision")
        FG.Columns.Add("Provision_Amt", "Provision Amt")
        FG.Columns.Add("WRITEOFF", "FOR AR WRITEOFF")
        FG.Columns.Add("Int_Call", "AR INT")
        FG.Columns.Add("Cust_ID", "Customer ID")
        
        For Each col As DataGridViewColumn In FG.Columns
            col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next
        dff()

    End Sub
    Private Sub dff()
        Dim dt As DataTable = DbHelper.GetDataTable("select * from Ap_RateSeting where Curr='LAK'")
        If dt.Rows.Count > 0 Then
            MLAK = Trim(dt.Rows(0)("Rate").ToString())
        End If
        dt = DbHelper.GetDataTable("select * from Ap_RateSeting where Curr='USD'")
        If dt.Rows.Count > 0 Then
            MUSD = Trim(dt.Rows(0)("Rate").ToString())
        End If
        dt = DbHelper.GetDataTable("select * from Ap_RateSeting where Curr='THB'")
        If dt.Rows.Count > 0 Then
            MTHB = Trim(dt.Rows(0)("Rate").ToString())
        End If
        dt = DbHelper.GetDataTable("select * from Ap_RateSeting where Curr='EUR'")
        If dt.Rows.Count > 0 Then
            MEUR = Trim(dt.Rows(0)("Rate").ToString())
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
        MM.Value = Microsoft.VisualBasic.Right(DataGridView1.Item(0, 0).Value.ToString, 10)
        FG.Rows.Clear()
        
        For i = 0 To DataGridView1.RowCount - 2
            If DataGridView1.Item(1, i).Value IsNot Nothing AndAlso DataGridView1.Item(1, i).Value.ToString <> "" Then
                FG.Rows.Add()
                Dim rowIndex As Integer = FG.Rows.Count - 1
                
                FG.Rows(rowIndex).Cells("No").Value = no + 1
                FG.Rows(rowIndex).Cells("LoanNO").Value = DataGridView1.Item(0, i).Value.ToString
                FG.Rows(rowIndex).Cells("LoanName").Value = DataGridView1.Item(1, i).Value.ToString
                FG.Rows(rowIndex).Cells("LoanDate").Value = DataGridView1.Item(7, i).Value?.ToString
                FG.Rows(rowIndex).Cells("Curr").Value = DataGridView1.Item(10, i).Value?.ToString
                FG.Rows(rowIndex).Cells("FIX_RATE").Value = DataGridView1.Item(14, i).Value?.ToString
                FG.Rows(rowIndex).Cells("PRINCIPLE").Value = DataGridView1.Item(22, i).Value?.ToString
                FG.Rows(rowIndex).Cells("INTEREST").Value = DataGridView1.Item(24, i).Value?.ToString
                FG.Rows(rowIndex).Cells("GENDER").Value = DataGridView1.Item(31, i).Value?.ToString
                FG.Rows(rowIndex).Cells("BUSINESSTYPEDESC").Value = DataGridView1.Item(32, i).Value?.ToString
                FG.Rows(rowIndex).Cells("LOAN_GRADE").Value = DataGridView1.Item(41, i).Value?.ToString
                FG.Rows(rowIndex).Cells("Provision").Value = DataGridView1.Item(42, i).Value?.ToString
                FG.Rows(rowIndex).Cells("Provision_Amt").Value = DataGridView1.Item(43, i).Value?.ToString
                FG.Rows(rowIndex).Cells("WRITEOFF").Value = DataGridView1.Item(45, i).Value?.ToString
                FG.Rows(rowIndex).Cells("Int_Call").Value = DataGridView1.Item(26, i).Value?.ToString
                FG.Rows(rowIndex).Cells("Cust_ID").Value = DataGridView1.Item(2, i).Value?.ToString
                
                If DataGridView1.Item(22, i).Value?.ToString <> "" Then
                    FG.Rows(rowIndex).Cells("WRITEOFF").Value = 0
                End If
            End If
        Next i
        'Call Calc()
        GroupBox1.Visible = False
        'Call amt()
    End Sub
    Private Sub amt()
        amt1 = 0
        amt2 = 0
        Dim i As Integer
        For i = 0 To FG.Rows.Count - 1
            If FG.Rows(i).Cells("Provision").Value IsNot Nothing AndAlso FG.Rows(i).Cells("Provision").Value.ToString <> "" Then
                amt1 = amt1 + CDbl(FG.Rows(i).Cells("Provision").Value.ToString)
            End If
            If FG.Rows(i).Cells("Provision_Amt").Value IsNot Nothing AndAlso FG.Rows(i).Cells("Provision_Amt").Value.ToString <> "" Then
                amt2 = amt2 + CDbl(FG.Rows(i).Cells("Provision_Amt").Value.ToString)
            End If
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
        ' DataGridView is read-only by default
    End Sub

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        'Call amt()
        If CDbl(txtSumAmountDr.Text) <> CDbl(txtSumAmountCr.Text) Then MsgBox("ບັນຊີນີ້ບໍ່ທັນບໍ່ທັນດູນດ່ຽງ ບໍ່ສາມາດບັນທຶກໄດ້!", MsgBoxStyle.OkOnly) : Exit Sub
        Call SaveItems()
        Call Insert_Gen_jn()
        MsgBox("ບັນທຶກສຳເລັດ")
    End Sub
    Private Sub SaveItems()
        DbHelper.ExecuteNonQuery(" delete TEM_AR  ")
        Dim i As Integer

        For i = 0 To FG.Rows.Count - 1
            If FG.Rows(i).Cells("LoanNO").Value IsNot Nothing AndAlso FG.Rows(i).Cells("LoanNO").Value.ToString <> "" Then
                Dim sa As String = " INSERT INTO TEM_AR (LoanNO, LoanName, LaonDate, Curr, FIX_RATE, PRINCIPLE, INTEREST, GENDER, BUSINESSTYPEDESC, LOAN_GRADE, Provision, Provision_Amt, WRITEOFF,Int_Call,Cust_ID, last_update, last_user,pc_nm) " & _
                " VALUES ( N'" & Apostrophe(FG.Rows(i).Cells("LoanNO").Value.ToString) & "'," & _
                    " N'" & Apostrophe(FG.Rows(i).Cells("LoanName").Value.ToString) & "'," & _
                     " N'" & Format(CDate(MM.Value), "yyyy-MM-dd") & "'," & _
                      " N'" & Trim(Apostrophe(FG.Rows(i).Cells("Curr").Value.ToString)) & "'," & _
                         " " & CDbl(If(FG.Rows(i).Cells("FIX_RATE").Value Is Nothing OrElse FG.Rows(i).Cells("FIX_RATE").Value.ToString = "", "0", FG.Rows(i).Cells("FIX_RATE").Value.ToString)) & ", " & _
                            " " & CDbl(If(FG.Rows(i).Cells("PRINCIPLE").Value Is Nothing OrElse FG.Rows(i).Cells("PRINCIPLE").Value.ToString = "", "0", FG.Rows(i).Cells("PRINCIPLE").Value.ToString)) & ", " & _
                                   " " & CDbl(If(FG.Rows(i).Cells("INTEREST").Value Is Nothing OrElse FG.Rows(i).Cells("INTEREST").Value.ToString = "", "0", FG.Rows(i).Cells("INTEREST").Value.ToString)) & ", " & _
                         " N'" & Apostrophe(FG.Rows(i).Cells("GENDER").Value.ToString) & "'," & _
                          " N'" & Apostrophe(FG.Rows(i).Cells("BUSINESSTYPEDESC").Value.ToString) & "'," & _
                              " N'" & Apostrophe(FG.Rows(i).Cells("LOAN_GRADE").Value.ToString) & "'," & _
                       " " & CDbl(If(FG.Rows(i).Cells("Provision").Value Is Nothing OrElse FG.Rows(i).Cells("Provision").Value.ToString = "", "0", FG.Rows(i).Cells("Provision").Value.ToString)) & ", " & _
                              " " & CDbl(If(FG.Rows(i).Cells("Provision_Amt").Value Is Nothing OrElse FG.Rows(i).Cells("Provision_Amt").Value.ToString = "", "0", FG.Rows(i).Cells("Provision_Amt").Value.ToString)) & ", " & _
                  " N'" & Apostrophe(FG.Rows(i).Cells("WRITEOFF").Value.ToString) & "'," & _
                      " " & CDbl(If(FG.Rows(i).Cells("Int_Call").Value Is Nothing OrElse FG.Rows(i).Cells("Int_Call").Value.ToString = "", "0", FG.Rows(i).Cells("Int_Call").Value.ToString)) & ", " & _
                        " N'" & Apostrophe(FG.Rows(i).Cells("Cust_ID").Value.ToString) & "'," & _
                           " Getdate()," & _
                " N'" & Apostrophe(MUserName) & "'," & _
                 " N'" & Apostrophe(MDServerName) & "') "
                DbHelper.ExecuteNonQuery(sa)
            End If
        Next i
        
        DbHelper.ExecuteNonQuery("update TEM_AR set rate=1 where curr='LAK' ")
        DbHelper.ExecuteNonQuery("update TEM_AR set rate=" & MUSD & " where curr='USD' ")
        DbHelper.ExecuteNonQuery("update TEM_AR set rate =" & MTHB & "  where curr='THB' ")

        DbHelper.ExecuteNonQuery("update TEM_AR set Int_Call=0 where Int_Call is null")
        DbHelper.ExecuteNonQuery("update AP_Loan set Int_Call=0 where Int_Call is null")
        DbHelper.ExecuteNonQuery("update TEM_AR set Provision=0 where Provision is null")
        DbHelper.ExecuteNonQuery("update AP_Loan set Provision_Amt=0 where Provision_Amt is null")

        DbHelper.ExecuteNonQuery(" update TEM_AR set principle_LAK=principle*rate  ")
        DbHelper.ExecuteNonQuery(" update TEM_AR set  Int_LAK=interest*rate  ")
        If CheckBox1.Checked = True Then
            DbHelper.ExecuteNonQuery("update TEM_AR set BUSINESSTYPEDESC=N'1.3 ປະກອບວັດຖຸເຕັກນິກ'  ")
        End If
    End Sub

    Private Sub Insert_Gen_jn()
        Frm_import_progress.Show()
        Dim aa As String
        Dim i As Integer
        
        For i = 0 To FG.Rows.Count - 1
            If FG.Rows(i).Cells("LoanNO").Value IsNot Nothing AndAlso FG.Rows(i).Cells("LoanNO").Value.ToString <> "" Then
                Dim sk As String = "Select * FROM AP_Loan where  LoanNO=N'" & FG.Rows(i).Cells("LoanNO").Value.ToString & "' and month(LaonDate)='" & Month(MM.Value) & "' and year(LaonDate)='" & Year(MM.Value) & "' "
                Dim dt As DataTable = DbHelper.GetDataTable(sk)
                If dt.Rows.Count = 0 Then
                    aa = "   insert into AP_Loan (LoanNO, LoanName, LaonDate, Curr, FIX_RATE, PRINCIPLE, INTEREST, GENDER, BUSINESSTYPEDESC, LOAN_GRADE,  Provision, Provision_Amt, WRITEOFF,rate,PRINCIPLE_lak, Int_LAK,Int_Call, Cust_ID,  last_update, last_user,pc_nm) " & _
     "   select  LoanNO, LoanName, LaonDate, Curr, FIX_RATE, PRINCIPLE, INTEREST, GENDER, BUSINESSTYPEDESC, LOAN_GRADE,  Provision, Provision_Amt, WRITEOFF,rate, PRINCIPLE_lak, Int_LAK,Int_Call, Cust_ID,  Getdate(),N'" & Apostrophe(MUserName) & "'," & _
          " N'" & Apostrophe(MDServerName) & "'" & _
   "    from   TEM_AR  where  LoanNO=N'" & FG.Rows(i).Cells("LoanNO").Value.ToString & "' order by LoanNO  "
                    DbHelper.ExecuteNonQuery(aa)
                Else
                    aa = "delete AP_Loan WHERE   LoanNO=N'" & FG.Rows(i).Cells("LoanNO").Value.ToString & "' and month(LaonDate)='" & Month(MM.Value) & "' and year(LaonDate)='" & Year(MM.Value) & "' "
                    DbHelper.ExecuteNonQuery(aa)
                    aa = "   insert into AP_Loan (LoanNO, LoanName, LaonDate, Curr, FIX_RATE, PRINCIPLE, INTEREST, GENDER, BUSINESSTYPEDESC, LOAN_GRADE,  Provision, Provision_Amt, WRITEOFF,rate,PRINCIPLE_lak, Int_LAK,Int_Call, Cust_ID,  last_update, last_user,pc_nm) " & _
"   select  LoanNO, LoanName, LaonDate, Curr, FIX_RATE, PRINCIPLE, INTEREST, GENDER, BUSINESSTYPEDESC, LOAN_GRADE,  Provision, Provision_Amt, WRITEOFF,rate, PRINCIPLE_lak, Int_LAK,Int_Call, Cust_ID,  Getdate(),N'" & Apostrophe(MUserName) & "'," & _
 " N'" & Apostrophe(MDServerName) & "'" & _
"    from   TEM_AR  where  LoanNO=N'" & FG.Rows(i).Cells("LoanNO").Value.ToString & "' order by LoanNO  "
                    DbHelper.ExecuteNonQuery(aa)
                End If
                Frm_import_progress.Refresh()
                Frm_import_progress.Label2.Text = FG.Rows.Count
                Frm_import_progress.Label4.Text = FG.Rows(i).Cells("No").Value?.ToString
                Frm_import_progress.Label1.Text = FG.Rows(i).Cells("LoanName").Value?.ToString
            End If
        Next
        Frm_import_progress.Close()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim sk As String = "Select * FROM AP_Loan   where   month(LaonDate)='" & Month(MM.Value) & "' and year(LaonDate)='" & Year(MM.Value) & "'  "
        Dim dt As DataTable = DbHelper.GetDataTable(sk)
        If dt.Rows.Count = 0 Then
            MsgBox("ຂໍ້ມູນເດືອນ " & Format(CDate(MM.Value), "MM/yyyy") & " ບໍ່ທັນມີ!", MsgBoxStyle.Exclamation) : Exit Sub
        End If

        If MessageBox.Show("ທ່ານຕ້ອງລຶບຂໍ້ມູນ  " & Format(CDate(MM.Value), "MM/yyyy") & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            DbHelper.ExecuteNonQuery("DELETE FROM AP_Loan   where   month(LaonDate)='" & Month(MM.Value) & "' and year(LaonDate)='" & Year(MM.Value) & "' ")
            MsgBox("Finish")
        End If
    End Sub
End Class