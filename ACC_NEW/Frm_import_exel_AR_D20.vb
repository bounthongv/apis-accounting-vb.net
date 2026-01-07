Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports ApPBank10.Module

Public Class Frm_import_exel_AR_D20
    Dim amt1 As Double
    Dim amt2 As Double
    Dim MLAK, MUSD, MTHB, MEUR As Double
    Private Sub Frm_import_exel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' Initialize DataGridView columns
        InitializeDataGridView()
        dff()

    End Sub

    Private Sub InitializeDataGridView()
        With DataGridViewFG
            .Columns.Clear()
            .Columns.Add("No", "ລ/ດ")
            .Columns.Add("LoanAccount", "LOAN ACCOUNT")
            .Columns.Add("ContractName", "CONTRACT NAME")
            .Columns.Add("LoanOpenDate", "LOAN OPEN DATE")
            .Columns.Add("Currency", "CURRENCY")
            .Columns.Add("FixRate", "FIX RATE")
            .Columns.Add("AROSTGPrinciple", "AR OS TG PRINCIPLE")
            .Columns.Add("AROSTGAccuredInterest", "AR OS TG ARCURED INTEREST")
            .Columns.Add("Gender", "GENDER")
            .Columns.Add("BusinessTypeDesc", "BUSINESSTYPEDESC")
            .Columns.Add("LoanGrade", "LOAN GRADE")
            .Columns.Add("ForARWriteoff", "FOR AR WRITEOFF")
            .Columns.Add("ARInt", "AR INT")
            .Columns.Add("CustomerID", "Customer ID")

            ' Set column widths
            For i As Integer = 0 To .Columns.Count - 1
                .Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            Next

            .ReadOnly = True
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
        End With
    End Sub
    Private Sub dff()
        Try
            Dim dt As DataTable = DbHelper.GetDataTable("select * from Ap_RateSeting where Curr='LAK'")
            If dt.Rows.Count > 0 Then
                MLAK = Trim(DbHelper.GetStr(dt.Rows(0)("Rate")))
            End If

            dt = DbHelper.GetDataTable("select * from Ap_RateSeting where Curr='USD'")
            If dt.Rows.Count > 0 Then
                MUSD = Trim(DbHelper.GetStr(dt.Rows(0)("Rate")))
            End If

            dt = DbHelper.GetDataTable("select * from Ap_RateSeting where Curr='THB'")
            If dt.Rows.Count > 0 Then
                MTHB = Trim(DbHelper.GetStr(dt.Rows(0)("Rate")))
            End If

            dt = DbHelper.GetDataTable("select * from Ap_RateSeting where Curr='EUR'")
            If dt.Rows.Count > 0 Then
                MEUR = Trim(DbHelper.GetStr(dt.Rows(0)("Rate")))
            End If
        Catch ex As Exception
            MessageBox.Show("Error loading exchange rates: " & ex.Message)
        End Try
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
            Dim i As Integer
            Dim no As Integer = 0
            
            If DataGridView1.Rows.Count > 0 Then
                MM.Value = Microsoft.VisualBasic.Right(DataGridView1.Item(0, 0).Value.ToString, 10)
            End If

            DataGridViewFG.Rows.Clear()

            For i = 0 To DataGridView1.RowCount - 2
                If DataGridView1.Item(1, i).Value IsNot Nothing AndAlso DataGridView1.Item(1, i).Value.ToString <> "" Then
                    Dim row As New DataGridViewRow()
                    row.CreateCells(DataGridViewFG)

                    ' Map Excel columns to DataGridView columns
                    ' Column 0: No (will be assigned later)
                    row.Cells(1).Value = DbHelper.GetStr(DataGridView1.Item(0, i).Value) ' LOAN ACCOUNT
                    row.Cells(2).Value = DbHelper.GetStr(DataGridView1.Item(1, i).Value) ' CONTRACT NAME
                    row.Cells(3).Value = DbHelper.GetStr(DataGridView1.Item(7, i).Value) ' LOAN OPEN DATE
                    row.Cells(4).Value = DbHelper.GetStr(DataGridView1.Item(10, i).Value) ' CURRENCY
                    row.Cells(5).Value = DbHelper.GetStr(DataGridView1.Item(14, i).Value) ' FIX RATE
                    row.Cells(6).Value = DbHelper.GetStr(DataGridView1.Item(22, i).Value) ' AR OS TG PRINCIPLE
                    row.Cells(7).Value = DbHelper.GetStr(DataGridView1.Item(24, i).Value) ' AR OS TG ARCURED INTEREST
                    row.Cells(8).Value = DbHelper.GetStr(DataGridView1.Item(31, i).Value) ' GENDER
                    row.Cells(9).Value = DbHelper.GetStr(DataGridView1.Item(32, i).Value) ' BUSINESSTYPEDESC
                    row.Cells(10).Value = DbHelper.GetStr(DataGridView1.Item(41, i).Value) ' LOAN GRADE
                    row.Cells(11).Value = DbHelper.GetStr(DataGridView1.Item(43, i).Value) ' FOR AR WRITEOFF
                    row.Cells(12).Value = DbHelper.GetStr(DataGridView1.Item(24, i).Value) ' AR INT (reusing column 24)
                    row.Cells(13).Value = DbHelper.GetStr(DataGridView1.Item(2, i).Value) ' Customer ID

                    DataGridViewFG.Rows.Add(row)
                End If
            Next i

            ' Assign row numbers
            For i = 0 To DataGridViewFG.Rows.Count - 1
                no = no + 1
                DataGridViewFG.Rows(i).Cells(0).Value = no.ToString()
                ' Clear some values if needed (mimicking original logic)
                If DataGridView1.Item(22, i).Value IsNot Nothing AndAlso DataGridView1.Item(22, i).Value.ToString <> "" Then
                    DataGridViewFG.Rows(i).Cells(11).Value = "0" ' FOR AR WRITEOFF
                End If
            Next

            GroupBox1.Visible = False
            amt()
        Catch ex As Exception
            MessageBox.Show("Error importing data: " & ex.Message)
        End Try
    End Sub
    Private Sub amt()
        amt1 = 0
        amt2 = 0
        Dim i As Integer
        For i = 0 To DataGridViewFG.Rows.Count - 1
            If DataGridViewFG.Rows(i).Cells(6).Value IsNot Nothing AndAlso IsNumeric(DataGridViewFG.Rows(i).Cells(6).Value) Then
                amt1 = amt1 + CDbl(DataGridViewFG.Rows(i).Cells(6).Value) ' AR OS TG PRINCIPLE
            End If
            If DataGridViewFG.Rows(i).Cells(7).Value IsNot Nothing AndAlso IsNumeric(DataGridViewFG.Rows(i).Cells(7).Value) Then
                amt2 = amt2 + CDbl(DataGridViewFG.Rows(i).Cells(7).Value) ' AR OS TG ARCURED INTEREST
            End If
        Next i

        txtSumAmountDr.Text = Format(CDbl(amt1), "##,##0.00")
        txtSumAmountCr.Text = Format(CDbl(amt2), "##,##0.00") ' Fixed to use amt2 instead of amt1
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

    ' FG_SelChange removed - VSFlexGrid replaced with DataGridView

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        'Call amt()
        If CDbl(txtSumAmountDr.Text) <> CDbl(txtSumAmountCr.Text) Then MsgBox("ບັນຊີນີ້ບໍ່ທັນບໍ່ທັນດູນດ່ຽງ ບໍ່ສາມາດບັນທຶກໄດ້!", MsgBoxStyle.OkOnly) : Exit Sub
        Call SaveItems()
        Call Insert_Gen_jn()
        MsgBox("ບັນທຶກສຳເລັດ")
    End Sub
    Private Sub SaveItems()
        Try
            Dim aa As String
            aa = " delete TEM_AR  "
            DbHelper.ExecuteNonQuery(aa)

            Dim i As Integer
            For i = 0 To DataGridViewFG.Rows.Count - 1
                If DataGridViewFG.Rows(i).Cells(1).Value IsNot Nothing AndAlso DataGridViewFG.Rows(i).Cells(1).Value.ToString <> "" Then
                    Dim sa As String = " INSERT INTO TEM_AR (LoanNO, LoanName, LaonDate, Curr, FIX_RATE, PRINCIPLE, INTEREST, GENDER, BUSINESSTYPEDESC, LOAN_GRADE, WRITEOFF,Int_Call,Cust_ID, last_update, last_user,pc_nm) " & _
                    " VALUES ( N'" & Apostrophe(DbHelper.GetStr(DataGridViewFG.Rows(i).Cells(1).Value)) & "'," & _
                        " N'" & Apostrophe(DbHelper.GetStr(DataGridViewFG.Rows(i).Cells(2).Value)) & "'," & _
                         " N'" & Format(CDate(MM.Value), "yyyy-MM-dd") & "'," & _
                          " N'" & Trim(Apostrophe(DbHelper.GetStr(DataGridViewFG.Rows(i).Cells(4).Value))) & "'," & _
                             " " & If(IsNumeric(DataGridViewFG.Rows(i).Cells(5).Value), CDbl(DataGridViewFG.Rows(i).Cells(5).Value), 0) & ", " & _
                                " " & If(IsNumeric(DataGridViewFG.Rows(i).Cells(6).Value), CDbl(DataGridViewFG.Rows(i).Cells(6).Value), 0) & ", " & _
                                       " " & If(IsNumeric(DataGridViewFG.Rows(i).Cells(7).Value), CDbl(DataGridViewFG.Rows(i).Cells(7).Value), 0) & ", " & _
                             " N'" & Apostrophe(DbHelper.GetStr(DataGridViewFG.Rows(i).Cells(8).Value)) & "'," & _
                              " N'" & Apostrophe(DbHelper.GetStr(DataGridViewFG.Rows(i).Cells(9).Value)) & "'," & _
                                  " N'" & Apostrophe(DbHelper.GetStr(DataGridViewFG.Rows(i).Cells(10).Value)) & "'," & _
                      " N'" & Apostrophe(DbHelper.GetStr(DataGridViewFG.Rows(i).Cells(11).Value)) & "'," & _
                          " " & If(IsNumeric(DataGridViewFG.Rows(i).Cells(12).Value), CDbl(DataGridViewFG.Rows(i).Cells(12).Value), 0) & ", " & _
                            " N'" & Apostrophe(DbHelper.GetStr(DataGridViewFG.Rows(i).Cells(13).Value)) & "'," & _
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

            DbHelper.ExecuteNonQuery(" update TEM_AR set principle_LAK=principle*rate  ")
            DbHelper.ExecuteNonQuery(" update TEM_AR set  Int_LAK=interest*rate  ")

            If CheckBox1.Checked = True Then
                DbHelper.ExecuteNonQuery("update TEM_AR set BUSINESSTYPEDESC=N'1.3 ປະກອບວັດຖຸເຕັກນິກ'  ")
            End If

        Catch ex As Exception
            MessageBox.Show("Error saving items: " & ex.Message)
        End Try
    End Sub

    Private Sub Insert_Gen_jn()
        Try
            Frm_import_progress.Show()
            Dim aa As String
            Dim i As Integer

            For i = 0 To DataGridViewFG.Rows.Count - 1
                If DataGridViewFG.Rows(i).Cells(1).Value IsNot Nothing AndAlso DataGridViewFG.Rows(i).Cells(1).Value.ToString <> "" Then
                    Dim sk As String = "Select * FROM AP_Loan where  LoanNO=N'" & Apostrophe(DbHelper.GetStr(DataGridViewFG.Rows(i).Cells(1).Value)) & "' and month(LaonDate)='" & Month(MM.Value) & "' and year(LaonDate)='" & Year(MM.Value) & "' "
                    Dim dt As DataTable = DbHelper.GetDataTable(sk)
                    If dt.Rows.Count = 0 Then
                        aa = "   insert into AP_Loan (LoanNO, LoanName, LaonDate, Curr, FIX_RATE, PRINCIPLE, INTEREST, GENDER, BUSINESSTYPEDESC, LOAN_GRADE, WRITEOFF,rate,PRINCIPLE_lak, Int_LAK,Int_Call, Cust_ID, last_update, last_user,pc_nm) " & _
             "   select  LoanNO, LoanName, LaonDate, Curr, FIX_RATE, PRINCIPLE, INTEREST, GENDER, BUSINESSTYPEDESC, LOAN_GRADE, WRITEOFF,rate, PRINCIPLE_lak, Int_LAK,Int_Call, Cust_ID,  Getdate(),N'" & Apostrophe(MUserName) & "'," & _
                   " N'" & Apostrophe(MDServerName) & "'" & _
            "    from   TEM_AR  where  LoanNO=N'" & Apostrophe(DbHelper.GetStr(DataGridViewFG.Rows(i).Cells(1).Value)) & "' order by LoanNO  "
                        DbHelper.ExecuteNonQuery(aa)
                    Else
                        aa = "delete AP_Loan WHERE   LoanNO=N'" & Apostrophe(DbHelper.GetStr(DataGridViewFG.Rows(i).Cells(1).Value)) & "' and month(LaonDate)='" & Month(MM.Value) & "' and year(LaonDate)='" & Year(MM.Value) & "' "
                        DbHelper.ExecuteNonQuery(aa)
                        aa = "   insert into AP_Loan (LoanNO, LoanName, LaonDate, Curr, FIX_RATE, PRINCIPLE, INTEREST, GENDER, BUSINESSTYPEDESC, LOAN_GRADE, WRITEOFF,rate,PRINCIPLE_lak, Int_LAK,Int_Call, Cust_ID,  last_update, last_user,pc_nm) " & _
        "   select  LoanNO, LoanName, LaonDate, Curr, FIX_RATE, PRINCIPLE, INTEREST, GENDER, BUSINESSTYPEDESC, LOAN_GRADE, WRITEOFF,rate, PRINCIPLE_lak, Int_LAK,Int_Call, Cust_ID,  Getdate(),N'" & Apostrophe(MUserName) & "'," & _
          " N'" & Apostrophe(MDServerName) & "'" & _
        "    from   TEM_AR  where  LoanNO=N'" & Apostrophe(DbHelper.GetStr(DataGridViewFG.Rows(i).Cells(1).Value)) & "' order by LoanNO  "
                        DbHelper.ExecuteNonQuery(aa)
                    End If

                    Frm_import_progress.Refresh()
                    Frm_import_progress.Label2.Text = DataGridViewFG.Rows.Count.ToString()
                    Frm_import_progress.Label4.Text = DbHelper.GetStr(DataGridViewFG.Rows(i).Cells(0).Value)
                    Frm_import_progress.Label1.Text = DbHelper.GetStr(DataGridViewFG.Rows(i).Cells(2).Value)
                End If
            Next
            Frm_import_progress.Close()
        Catch ex As Exception
            If Frm_import_progress IsNot Nothing AndAlso Not Frm_import_progress.IsDisposed Then
                Frm_import_progress.Close()
            End If
            MessageBox.Show("Error inserting data: " & ex.Message)
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Dim sk As String = "Select * FROM AP_Loan   where   month(LaonDate)='" & Month(MM.Value) & "' and year(LaonDate)='" & Year(MM.Value) & "'  "
            Dim dt As DataTable = DbHelper.GetDataTable(sk)
            If dt.Rows.Count = 0 Then
                MsgBox("ຂໍ້ມູນເດືອນ " & Format(CDate(MM.Value), "MM/yyyy") & " ບໍ່ທັນມີ!", MsgBoxStyle.Exclamation) : Exit Sub
            End If

            If MessageBox.Show("ທ່ານຕ້ອງລຶບຂໍ້ມູນ  " & Format(CDate(MM.Value), "MM/yyyy") & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                DbHelper.ExecuteNonQuery("DELETE FROM AP_Loan   where   month(LaonDate)='" & Month(MM.Value) & "' and year(LaonDate)='" & Year(MM.Value) & "' ")
                MsgBox("Finish")
            End If
        Catch ex As Exception
            MessageBox.Show("Error clearing data: " & ex.Message)
        End Try
    End Sub
End Class