Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class Frm_import_exel_AR_D20
    Dim amt1 As Double
    Dim amt2 As Double
    Dim MLAK, MUSD, MTHB, MEUR As Double
    Private Sub Frm_import_exel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        FG.FormatString = "^ລ/ດ |<LOAN ACCOUNT   |<CONTRACT NAME  |<LOAN OPEN DATE   |<CURRENCY    |<FIX RATE  |<AR OS TG PRINCIPLE |< AR OS TG ARCURED INTEREST |<GENDER   |<BUSINESSTYPEDESC  |<LOAN GRADE |<FOR AR WRITEOFF |< AR INT|<Customer ID  "
        dff()

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
        MM.Value = Microsoft.VisualBasic.Right(DataGridView1.Item(0, 0).Value.ToString, 10)
        FG.Rows = 1
        For i = 0 To DataGridView1.RowCount - 2
            If DataGridView1.Item(1, i).Value.ToString <> "" Then

                'FG.AddItem(DataGridView1.Item(0, i).Value.ToString & Chr(9) & DataGridView1.Item(1, i).Value.ToString)

                FG.AddItem(DataGridView1.Item(0, i).Value.ToString & _
                  Chr(9) & DataGridView1.Item(0, i).Value.ToString & _
                  Chr(9) & DataGridView1.Item(1, i).Value.ToString & _
                  Chr(9) & DataGridView1.Item(7, i).Value.ToString & _
                 Chr(9) & DataGridView1.Item(10, i).Value.ToString & _
                Chr(9) & DataGridView1.Item(14, i).Value.ToString & _
               Chr(9) & DataGridView1.Item(22, i).Value.ToString & _
                Chr(9) & DataGridView1.Item(24, i).Value.ToString & _
                Chr(9) & DataGridView1.Item(31, i).Value.ToString & _
                  Chr(9) & DataGridView1.Item(32, i).Value.ToString & _
                 Chr(9) & DataGridView1.Item(41, i).Value.ToString & _
                    Chr(9) & DataGridView1.Item(43, i).Value.ToString & _
                          Chr(9) & DataGridView1.Item(24, i).Value.ToString & _
                       Chr(9) & DataGridView1.Item(2, i).Value.ToString)

                'FG.set_TextMatrix(FG.Row, no, no)
            End If

        Next i

        For i = 1 To FG.Rows - 1
            no = no + 1
            FG.set_TextMatrix(i, 0, no)
            If DataGridView1.Item(22, i).Value.ToString <> "" Then
                FG.set_TextMatrix(i, 22, 0)
            End If
            'If FG.get_TextMatrix(i, 3) = "" Then
            '    FG.set_TextMatrix(i, 3, FG.get_TextMatrix(i, 2) & "/" & FG.get_TextMatrix(i, 5))
            'End If

            'FG.set_TextMatrix(i, 10, Format(CDbl(FG.get_TextMatrix(i, 10)), "#,##0.00"))
            'FG.set_TextMatrix(i, 11, Format(CDbl(FG.get_TextMatrix(i, 11)), "#,##0.00"))
            'FG.set_TextMatrix(i, 13, Format(CDbl(FG.get_TextMatrix(i, 13)), "#,##0.00"))
            'FG.set_TextMatrix(i, 14, Format(CDbl(FG.get_TextMatrix(i, 14)), "#,##0.00"))
            'FG.set_TextMatrix(i, 15, Format(CDbl(FG.get_TextMatrix(i, 15)), "#,##0.00"))

            'FG.set_TextMatrix(i, no)

        Next
        'Call Calc()
        GroupBox1.Visible = False
        'Call amt()
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

    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelChange
        FG.Editable = VSFlex8U.EditableSettings.flexEDKbd
    End Sub

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        'Call amt()
        If CDbl(txtSumAmountDr.Text) <> CDbl(txtSumAmountCr.Text) Then MsgBox("ບັນຊີນີ້ບໍ່ທັນບໍ່ທັນດູນດ່ຽງ ບໍ່ສາມາດບັນທຶກໄດ້!", MsgBoxStyle.OkOnly) : Exit Sub
        Call SaveItems()
        Call Insert_Gen_jn()
        MsgBox("ບັນທຶກສຳເລັດ")
    End Sub
    Private Sub SaveItems()
        Dim aa As String
        aa = " delete TEM_AR  "
        CNN.Execute(aa)
        Dim Rschk As New ADODB.Recordset
        Dim i As Integer
        With Rschk
            Dim sk As String = "Select * FROM TEM_AR  "
            Call LoadSqlData(sk, Rschk)

            For i = 3 To FG.Rows - 1

                If Rschk.RecordCount = 0 Then

                    Dim sa As String = " INSERT INTO TEM_AR (LoanNO, LoanName, LaonDate, Curr, FIX_RATE, PRINCIPLE, INTEREST, GENDER, BUSINESSTYPEDESC, LOAN_GRADE, WRITEOFF,Int_Call,Cust_ID, last_update, last_user,pc_nm) " & _
                    " VALUES ( N'" & Apostrophe(FG.get_TextMatrix(i, 1)) & "'," & _
                        " N'" & Apostrophe(FG.get_TextMatrix(i, 2)) & "'," & _
                         " N'" & Format(CDate(MM.Value), "yyyy-MM-dd") & "'," & _
                          " N'" & Trim(Apostrophe(FG.get_TextMatrix(i, 4))) & "'," & _
                             " " & CDbl(FG.get_TextMatrix(i, 5)) & ", " & _
                                " " & CDbl(FG.get_TextMatrix(i, 6)) & ", " & _
                                       " " & CDbl(FG.get_TextMatrix(i, 7)) & ", " & _
                             " N'" & Apostrophe(FG.get_TextMatrix(i, 8)) & "'," & _
                              " N'" & Apostrophe(FG.get_TextMatrix(i, 9)) & "'," & _
                                  " N'" & Apostrophe(FG.get_TextMatrix(i, 10)) & "'," & _
                      " N'" & Apostrophe(FG.get_TextMatrix(i, 11)) & "'," & _
                          " " & CDbl(FG.get_TextMatrix(i, 12)) & ", " & _
                            " N'" & Apostrophe(FG.get_TextMatrix(i, 13)) & "'," & _
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
        CNN.Execute("update TEM_AR set rate=1 where curr='LAK' ")
        CNN.Execute("update TEM_AR set rate=" & MUSD & " where curr='USD' ")
        CNN.Execute("update TEM_AR set rate =" & MTHB & "  where curr='THB' ")

        CNN.Execute("update TEM_AR set Int_Call=0 where Int_Call is null")
        CNN.Execute("update AP_Loan set Int_Call=0 where Int_Call is null")

        CNN.Execute(" update TEM_AR set principle_LAK=principle*rate  ")
        CNN.Execute(" update TEM_AR set  Int_LAK=interest*rate  ")
        'CNN.Execute(" update TEM_AR set  Int_LAK=interest*rate where Int_Call<>0 ")
        If CheckBox1.Checked = True Then
            CNN.Execute("update TEM_AR set BUSINESSTYPEDESC=N'1.3 ປະກອບວັດຖຸເຕັກນິກ'  ")
        End If


        'CNN.Execute("delete gen_jn WHERE ac_code='' ")
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
            For i = 2 To FG.Rows - 1
                Dim sk As String = "Select * FROM AP_Loan where  LoanNO=N'" & FG.get_TextMatrix(i, 1) & "' and month(LaonDate)='" & Month(MM.Value) & "' and year(LaonDate)='" & Year(MM.Value) & "' "
                Call LoadSqlData(sk, Rschk)
                If Rschk.RecordCount = 0 Then
                    aa = "   insert into AP_Loan (LoanNO, LoanName, LaonDate, Curr, FIX_RATE, PRINCIPLE, INTEREST, GENDER, BUSINESSTYPEDESC, LOAN_GRADE, WRITEOFF,rate,PRINCIPLE_lak, Int_LAK,Int_Call, Cust_ID, last_update, last_user,pc_nm) " & _
         "   select  LoanNO, LoanName, LaonDate, Curr, FIX_RATE, PRINCIPLE, INTEREST, GENDER, BUSINESSTYPEDESC, LOAN_GRADE, WRITEOFF,rate, PRINCIPLE_lak, Int_LAK,Int_Call, Cust_ID,  Getdate(),N'" & Apostrophe(MUserName) & "'," & _
              " N'" & Apostrophe(MDServerName) & "'" & _
       "    from   TEM_AR  where  LoanNO=N'" & FG.get_TextMatrix(i, 1) & "' order by LoanNO  "
                    CNN.Execute(aa)

                Else

                    aa = "delete AP_Loan WHERE   LoanNO=N'" & FG.get_TextMatrix(i, 1) & "' and month(LaonDate)='" & Month(MM.Value) & "' and year(LaonDate)='" & Year(MM.Value) & "' "
                    CNN.Execute(aa)
                    aa = "   insert into AP_Loan (LoanNO, LoanName, LaonDate, Curr, FIX_RATE, PRINCIPLE, INTEREST, GENDER, BUSINESSTYPEDESC, LOAN_GRADE, WRITEOFF,rate,PRINCIPLE_lak, Int_LAK,Int_Call, Cust_ID,  last_update, last_user,pc_nm) " & _
"   select  LoanNO, LoanName, LaonDate, Curr, FIX_RATE, PRINCIPLE, INTEREST, GENDER, BUSINESSTYPEDESC, LOAN_GRADE, WRITEOFF,rate, PRINCIPLE_lak, Int_LAK,Int_Call, Cust_ID,  Getdate(),N'" & Apostrophe(MUserName) & "'," & _
     " N'" & Apostrophe(MDServerName) & "'" & _
"    from   TEM_AR  where  LoanNO=N'" & FG.get_TextMatrix(i, 1) & "' order by LoanNO  "
                    CNN.Execute(aa)


                End If
                Frm_import_progress.Refresh()

                Frm_import_progress.Label2.Text = FG.Rows

                Frm_import_progress.Label4.Text = FG.get_TextMatrix(i, 0)

                Frm_import_progress.Label1.Text = FG.get_TextMatrix(i, 2)

            Next
            Frm_import_progress.Close()
        End With


    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim sk As String = "Select * FROM AP_Loan   where   month(LaonDate)='" & Month(MM.Value) & "' and year(LaonDate)='" & Year(MM.Value) & "'  "
        Call LoadSqlData(sk, RSC)
        If RSC.RecordCount = 0 Then
            MsgBox("ຂໍ້ມູນເດືອນ " & Format(CDate(MM.Value), "MM/yyyy") & " ບໍ່ທັນມີ!", MsgBoxStyle.Exclamation) : Exit Sub
        End If

        If MessageBox.Show("ທ່ານຕ້ອງລຶບຂໍ້ມູນ  " & Format(CDate(MM.Value), "MM/yyyy") & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            CNN.Execute("DELETE FROM AP_Loan   where   month(LaonDate)='" & Month(MM.Value) & "' and year(LaonDate)='" & Year(MM.Value) & "' ")
            MsgBox("Finish")
        End If

    End Sub
End Class