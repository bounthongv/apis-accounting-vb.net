Imports MySql.Data.MySqlClient
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Net

Public Class FmJeneralJournal_List
    Dim SQL As String
    Dim Ch As Double
    Dim LockData As String
    Dim CntNB As String = "cnt"
    Dim RptName As String
    Dim BookId As String
    Dim Sto As String
    Dim SR As String
    Dim OrD As String = "   ASC"
    Dim Op As Double = 0
    Dim x As Integer = 0
    Dim px As Integer = 0
    Dim y As Integer = 0
    Dim x0 As Integer = 0
    Dim y0 As Integer = 0
    Dim s, s0 As String
    Dim Rs1 As Integer = 0
    Dim Rs2 As Integer = 0
    Dim Rt1 As Integer = 0
    Dim D As Integer
    Dim Rt2 As Integer = 0
    Dim Amount_In_Word As String
    Dim MDUPASET As String
    Dim MDUPASETAMT As Double = 0

    ' DataGridView Helper Methods
    Private Function GetGridValue(ByVal grid As DataGridView, ByVal row As Integer, ByVal col As Integer) As String
        If row >= 0 AndAlso row < grid.RowCount AndAlso col >= 0 AndAlso col < grid.ColumnCount Then
            If grid.Rows(row).Cells(col).Value IsNot Nothing Then
                Return grid.Rows(row).Cells(col).Value.ToString()
            End If
        End If
        Return ""
    End Function

    Private Sub SetGridValue(ByVal grid As DataGridView, ByVal row As Integer, ByVal col As Integer, ByVal value As String)
        If row >= 0 AndAlso row < grid.RowCount AndAlso col >= 0 AndAlso col < grid.ColumnCount Then
            grid.Rows(row).Cells(col).Value = value
        End If
    End Sub

    Private Sub SetupGrid()
        ' Clear existing columns
        FG.Columns.Clear()

        ' Add columns based on FormatString from original code
        FG.Columns.Add("Col0", "ລ/ດ")
        FG.Columns.Add("Col1", "ວດ")
        FG.Columns.Add("Col2", "ວັນ")
        FG.Columns.Add("Col3", "ປະຈຳວັນ")
        FG.Columns.Add("Col4", "ວັນ")
        FG.Columns.Add("Col5", "ລະຫັດ")
        FG.Columns.Add("Col6", "ວັນ")
        FG.Columns.Add("Col7", "ວັນ")
        FG.Columns.Add("Col8", "ວັນ")
        FG.Columns.Add("Col9", "ວັນ")
        FG.Columns.Add("Col10", "ວັນ")
        FG.Columns.Add("Col11", "ວັນ")
        FG.Columns.Add("Col12", "ວັນ")
        FG.Columns.Add("Col13", "ວັນ")

        ' Set column properties
        For i As Integer = 0 To FG.Columns.Count - 1
            FG.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        Next
    End Sub
    Private Sub BntNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BntNew.Click
        Panel4.Visible = False
        FmNsewJeneralJournal.txtInvoice.Enabled = True
        FmNsewJeneralJournal.CmbBook.Enabled = True
        'FmNsewJeneralJournal.MdiParent = Me
        'FmNsewJeneralJournal.WindowState = FormWindowState.Maximized


        FmNsewJeneralJournal.Show()
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        LoadMonthSQL()
    End Sub
    Private Sub ColorRadioButton()
        If RadioButton14.Checked = True Then
            dts.Enabled = True
            dtt.Enabled = True
        Else
            dts.Enabled = False
            dtt.Enabled = False
        End If
        If RadioButton1.Checked = True Then
            RadioButton1.BackColor = Color.Aquamarine
        Else
            RadioButton1.BackColor = Color.Gainsboro
        End If
        If RadioButton2.Checked = True Then
            RadioButton2.BackColor = Color.Aquamarine
        Else
            RadioButton2.BackColor = Color.Gainsboro
        End If
        If RadioButton3.Checked = True Then
            RadioButton3.BackColor = Color.Aquamarine
        Else
            RadioButton3.BackColor = Color.Gainsboro
        End If
        If RadioButton4.Checked = True Then
            RadioButton4.BackColor = Color.Aquamarine
        Else
            RadioButton4.BackColor = Color.Gainsboro
        End If
        If RadioButton5.Checked = True Then
            RadioButton5.BackColor = Color.Aquamarine
        Else
            RadioButton5.BackColor = Color.Gainsboro
        End If
        If RadioButton6.Checked = True Then
            RadioButton6.BackColor = Color.Aquamarine
        Else
            RadioButton6.BackColor = Color.Gainsboro
        End If
        If RadioButton7.Checked = True Then
            RadioButton7.BackColor = Color.Aquamarine
        Else
            RadioButton7.BackColor = Color.Gainsboro
        End If
        If RadioButton8.Checked = True Then
            RadioButton8.BackColor = Color.Aquamarine
        Else
            RadioButton8.BackColor = Color.Gainsboro
        End If
        If RadioButton9.Checked = True Then
            RadioButton9.BackColor = Color.Aquamarine
        Else
            RadioButton9.BackColor = Color.Gainsboro
        End If
        If RadioButton10.Checked = True Then
            RadioButton10.BackColor = Color.Aquamarine
        Else
            RadioButton10.BackColor = Color.Gainsboro
        End If
        If RadioButton11.Checked = True Then
            RadioButton11.BackColor = Color.Aquamarine
        Else
            RadioButton11.BackColor = Color.Gainsboro
        End If

        If RadioButton12.Checked = True Then
            RadioButton12.BackColor = Color.Aquamarine
        Else
            RadioButton12.BackColor = Color.Gainsboro
        End If


        If RadioButton13.Checked = True Then
            RadioButton13.BackColor = Color.Aquamarine
        Else
            RadioButton13.BackColor = Color.Gainsboro
        End If
        If RadioButton14.Checked = True Then
            RadioButton14.BackColor = Color.Aquamarine
        Else
            RadioButton14.BackColor = Color.Gainsboro
        End If

    End Sub

    Private Sub ClickMouseRadio2()


        Dim D, D2, Y As String

        If RadioButton1.Checked = True Then
            LngId = "7013" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (01), " & Y & " " & Format(dts.Value, "yyyy")
            RptName = "ປະຈຳເດືອນ ມັງກອນ ປີ " & Format(dts.Value, "yyyy")
        End If
        If RadioButton2.Checked = True Then
            LngId = "7014" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (02), " & Y & " " & Format(dts.Value, "yyyy")
            RptName = "ປະຈຳເດືອນ ກຸມພາ ປີ " & Format(dts.Value, "yyyy")
        End If
        If RadioButton3.Checked = True Then
            LngId = "7015" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (03), " & Y & " " & Format(dts.Value, "yyyy")
            RptName = "ປະຈຳເດືອນ ມີນາ ປີ " & Format(dts.Value, "yyyy")
        End If
        If RadioButton4.Checked = True Then
            LngId = "7016" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (04), " & Y & " " & Format(dts.Value, "yyyy")
            RptName = "ປະຈຳເດືອນ ເມສາ ປີ " & Format(dts.Value, "yyyy")
        End If
        If RadioButton5.Checked = True Then
            LngId = "7017" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (05), " & Y & " " & Format(dts.Value, "yyyy")
            RptName = "ປະຈຳເດືອນ ພຶດສະພາ ປີ " & Format(dts.Value, "yyyy")
        End If
        If RadioButton6.Checked = True Then
            LngId = "7018" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (06), " & Y & " " & Format(dts.Value, "yyyy")
            RptName = "ປະຈຳເດືອນ ມີຖຸນາ ປີ " & Format(dts.Value, "yyyy")
        End If
        If RadioButton7.Checked = True Then
            LngId = "7019" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (07), " & Y & " " & Format(dts.Value, "yyyy")
            RptName = "ປະຈຳເດືອນ ກໍລະກົດ ປີ " & Format(dts.Value, "yyyy")
        End If
        If RadioButton8.Checked = True Then
            LngId = "7020" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (08), " & Y & " " & Format(dts.Value, "yyyy")
            RptName = "ປະຈຳເດືອນ ສິງຫາ ປີ " & Format(dts.Value, "yyyy")
        End If
        If RadioButton9.Checked = True Then
            LngId = "7021" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (09), " & Y & " " & Format(dts.Value, "yyyy")
            RptName = "ປະຈຳເດືອນ ກັນຍາ ປີ " & Format(dts.Value, "yyyy")
        End If
        If RadioButton10.Checked = True Then
            LngId = "7022" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (10), " & Y & " " & Format(dts.Value, "yyyy")
            RptName = "ປະຈຳເດືອນ ຕຸລາ ປີ " & Format(dts.Value, "yyyy")
        End If
        If RadioButton11.Checked = True Then
            LngId = "7023" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (11), " & Y & " " & Format(dts.Value, "yyyy")
            RptName = "ປະຈຳເດືອນ ພະຈິກ ປີ " & Format(dts.Value, "yyyy")
        End If
        If RadioButton12.Checked = True Then
            LngId = "7024" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (12), " & Y & " " & Format(dts.Value, "yyyy")
            RptName = "ປະຈຳເດືອນ ທັນວາ ປີ " & Format(dts.Value, "yyyy")
        End If
        If RadioButton13.Checked = True Then
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = Y & Format(dts.Value, "yyyy")
            dts.Value = "1-1-" & Year(MWorkSetting)
            dtt.Value = "31-12-" & Year(MWorkSetting)
        End If
        If RadioButton14.Checked = True Then
            LngId = "7026" : CallLngStr() : D = LngStr
            LngId = "7027" : CallLngStr() : D2 = LngStr
            RptName = D & Format(dts.Value, "dd/MM/yyyy") & D2 & Format(dtt.Value, "dd/MM/yyyy")
        End If
    End Sub

    Private Sub ClickMouseRadio()
        Dim D, D2, Y As String
        If RadioButton14.Checked = True Then
            dts.Enabled = True
            dtt.Enabled = True
        Else
            dts.Enabled = False
            dtt.Enabled = False
        End If
        If RadioButton1.Checked = True Then
            LngId = "7013" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (01), " & Y & " " & Format(dts.Value, "yyyy")

            dts.Text = "01/01/" & Year(MWorkSetting)
            Dim x As Date
            x = DateAdd(DateInterval.Month, 1, dts.Value)
            dtt.Value = DateAdd(DateInterval.Day, -1, x)
        End If
        If RadioButton2.Checked = True Then
            LngId = "7014" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (02), " & Y & " " & Format(dts.Value, "yyyy")
            dts.Value = "01/02/" & Year(MWorkSetting)
            Dim x As Date
            x = DateAdd(DateInterval.Month, 1, dts.Value)
            dtt.Value = DateAdd(DateInterval.Day, -1, x)
        End If
        If RadioButton3.Checked = True Then
            LngId = "7015" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (03), " & Y & " " & Format(dts.Value, "yyyy")
            'MsgBox(RptName)
            dts.Value = "1-3-" & Year(MWorkSetting)
            Dim x As Date
            x = DateAdd(DateInterval.Month, 1, dts.Value)
            dtt.Value = DateAdd(DateInterval.Day, -1, x)
        End If
        If RadioButton4.Checked = True Then
            LngId = "7016" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (04), " & Y & " " & Format(dts.Value, "yyyy")
            dts.Value = "1-4-" & Year(MWorkSetting)
            Dim x As Date
            x = DateAdd(DateInterval.Month, 1, dts.Value)
            dtt.Value = DateAdd(DateInterval.Day, -1, x)
        End If
        If RadioButton5.Checked = True Then
            LngId = "7017" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (05), " & Y & " " & Format(dts.Value, "yyyy")
            dts.Value = "1-5-" & Year(MWorkSetting)
            Dim x As Date
            x = DateAdd(DateInterval.Month, 1, dts.Value)
            dtt.Value = DateAdd(DateInterval.Day, -1, x)
        End If
        If RadioButton6.Checked = True Then
            LngId = "7018" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (06), " & Y & " " & Format(dts.Value, "yyyy")
            dts.Value = "1-6-" & Year(MWorkSetting)
            Dim x As Date
            x = DateAdd(DateInterval.Month, 1, dts.Value)
            dtt.Value = DateAdd(DateInterval.Day, -1, x)
        End If
        If RadioButton7.Checked = True Then
            LngId = "7019" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (07), " & Y & " " & Format(dts.Value, "yyyy")
            dts.Value = "1-7-" & Year(MWorkSetting)
            Dim x As Date
            x = DateAdd(DateInterval.Month, 1, dts.Value)
            dtt.Value = DateAdd(DateInterval.Day, -1, x)
        End If
        If RadioButton8.Checked = True Then
            LngId = "7020" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (08), " & Y & " " & Format(dts.Value, "yyyy")
            dts.Value = "1-8-" & Year(MWorkSetting)
            Dim x As Date
            x = DateAdd(DateInterval.Month, 1, dts.Value)
            dtt.Value = DateAdd(DateInterval.Day, -1, x)
        End If
        If RadioButton9.Checked = True Then
            LngId = "7021" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (09), " & Y & " " & Format(dts.Value, "yyyy")
            dts.Value = "1-9-" & Year(MWorkSetting)
            Dim x As Date
            x = DateAdd(DateInterval.Month, 1, dts.Value)
            dtt.Value = DateAdd(DateInterval.Day, -1, x)
        End If
        If RadioButton10.Checked = True Then
            LngId = "7022" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (10), " & Y & " " & Format(dts.Value, "yyyy")
            dts.Value = "1-10-" & Year(MWorkSetting)
            Dim x As Date
            x = DateAdd(DateInterval.Month, 1, dts.Value)
            dtt.Value = DateAdd(DateInterval.Day, -1, x)
        End If
        If RadioButton11.Checked = True Then
            LngId = "7023" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (11), " & Y & " " & Format(dts.Value, "yyyy")
            dts.Value = "1-11-" & Year(MWorkSetting)
            Dim x As Date
            x = DateAdd(DateInterval.Month, 1, dts.Value)
            dtt.Value = DateAdd(DateInterval.Day, -1, x)
        End If

        If RadioButton12.Checked = True Then
            LngId = "7024" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (12), " & Y & " " & Format(dts.Value, "yyyy")
            dts.Value = "1-12-" & Year(MWorkSetting)
            Dim x As Date
            x = DateAdd(DateInterval.Month, 1, dts.Value)
            dtt.Value = DateAdd(DateInterval.Day, -1, x)
        End If



        If RadioButton13.Checked = True Then
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = Y & Format(dts.Value, "yyyy")
            dts.Value = "1-1-" & Year(MWorkSetting)
            dtt.Value = "31-12-" & Year(MWorkSetting)
        End If
        If RadioButton14.Checked = True Then
            LngId = "7026" : CallLngStr() : D = LngStr
            LngId = "7027" : CallLngStr() : D2 = LngStr
            RptName = D & Format(dts.Value, "dd/MM/yyyy") & D2 & Format(dtt.Value, "dd/MM/yyyy")

        End If

    End Sub
    Public Sub LoadMonthSQL()

        SQL = ""
        Panel4.Visible = False

        StartLoadDataList()



        MDInvoiceNo = ""
    End Sub
    Private Sub LoadSQLCheckbox()
        SR = ""
        If Rinvioce.Checked = True Then
            'SR = " AND gen_jn.certify = '" & Nme.Text & "' "
            SR = " AND (gen_jn.certify = '" & Nme.Text & "' or gen_jn.referno  Like N'%" & Nme.Text.Trim & "%') "
        End If
        If RCex.Checked = True Then
            'SR = " AND gen_jn.cheque_no = '" & Nme.Text & "' "
            SR = " AND gen_jn.referno = '" & Nme.Text & "' "
        End If

        If RAc_code.Checked = True Then
            SR = " AND (gen_jn.ac_code  Like N'" & Nme.Text.Trim & "%')"
        End If
        If RAcNme.Checked = True Then
            SR = " AND (gen_jn.ac_name  Like N'%" & Nme.Text.Trim & "%')"
        End If
        If RDesc.Checked = True Then
            SR = " AND (gen_jn.descrip  Like N'%" & Nme.Text.Trim & "%')"
        End If
        If RBook.Checked = True Then
            SR = " AND book = '" & Nme.Text & "' "
            If Nme.Text = "<All>  ທັງໝົດ (All books)" Then
                SR = ""
            End If
        End If
        If RAcType.Checked = True Then
            SR = " AND ac_type = '" & Nme.Text & "' "
            If Nme.Text = "<All>  ທັງໝົດ" Then
                SR = ""
            End If
        End If
        If RCurr.Checked = True Then
            SR = " AND gen_jn.curr = '" & Nme.Text & "' "
            If Nme.Text = "==ທັງຫມົດ==" Then
                SR = ""
            End If
        End If
        'ເບິ່ງແບບສັງລວມ
    End Sub




    Private Sub Load_M()
        'RD.Checked = True
        'Ds.Value = MWorkSetting
        'Myy.Value = MWorkSetting
        'yy.Value = MWorkSetting
        'Toyy.Value = MWorkSetting
        'Pyy.Value = MWorkSetting
        If Month(MWorkSetting) = 1 Then
            RadioButton1.Checked = True
        ElseIf Month(MWorkSetting) = 2 Then
            RadioButton2.Checked = True
        ElseIf Month(MWorkSetting) = 3 Then
            RadioButton3.Checked = True
        ElseIf Month(MWorkSetting) = 4 Then
            RadioButton4.Checked = True
        ElseIf Month(MWorkSetting) = 5 Then
            RadioButton5.Checked = True
        ElseIf Month(MWorkSetting) = 6 Then
            RadioButton6.Checked = True
        ElseIf Month(MWorkSetting) = 7 Then
            RadioButton7.Checked = True
        ElseIf Month(MWorkSetting) = 8 Then
            RadioButton8.Checked = True
        ElseIf Month(MWorkSetting) = 9 Then
            RadioButton9.Checked = True
        ElseIf Month(MWorkSetting) = 10 Then
            RadioButton10.Checked = True
        ElseIf Month(MWorkSetting) = 11 Then
            RadioButton11.Checked = True
        ElseIf Month(MWorkSetting) = 12 Then
            RadioButton12.Checked = True
        End If
        LoadMonthSQL()
    End Sub
    Private Sub loadOffice_User()
        Off_Usr.Items.Clear()
        Dim dt As DataTable = DbHelper.GetDataTable("select sub_id , off_add2  from  Ap_office  Order by sub_id")
        For Each row As DataRow In dt.Rows
            Off_Usr.Items.Add((DbHelper.GetStr(row("sub_id"))) & " " & DbHelper.GetStr(row("off_add2")))
        Next
        Off_Usr.Text = FmLogin.Sub_Company.Text
    End Sub

    Private Sub FmJeneralJournal_List_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MuSubOff = MuSubOff2
    End Sub
    Private Sub loadCompany()

        cmbCompany.Items.Clear()
        Dim dt2 As DataTable = DbHelper.GetDataTable("select off_add1 , off_id  from  Ap_office group BY off_id , off_add1")
        For Each row As DataRow In dt2.Rows
            cmbCompany.Items.Add((DbHelper.GetStr(row("off_id"))) & " " & DbHelper.GetStr(row("off_add1")))
        Next
        CmbCompany.SelectedIndex = FmLogin.cmbCompany.SelectedIndex
        If MPermit = "User" Then
            CmbCompany.Enabled = False
        End If
        SUPD = 0
    End Sub
    Private Sub Load_DES()
        If certify.Checked = False Then
            CheckBox3.Checked = False
        End If
        If MASC.Checked Then
            If date_work.Checked = True Then
                'CntNB = "date_work ASC, cnt ASC"
                CntNB = "date_work ASC "
            ElseIf certify.Checked = True Then
                CntNB = "certify ASC, cnt ASC"
            ElseIf cheque_no.Checked = True Then
                CntNB = "cheque_no ASC, cnt ASC"
            ElseIf ac_code.Checked = True Then
                CntNB = "ac_code ASC, cnt ASC"
            ElseIf descrip.Checked = True Then
                CntNB = "descrip ASC, cnt ASC"
            ElseIf Book.Checked = True Then
                CntNB = "Book ASC, cnt ASC"
            ElseIf Curr.Checked = True Then
                CntNB = "Curr ASC, cnt ASC"
            End If

        Else
            If date_work.Checked = True Then
                CntNB = "date_work DESC, cnt DESC"
            ElseIf certify.Checked = True Then
                CntNB = "certify DESC, cnt DESC"
            ElseIf cheque_no.Checked = True Then
                CntNB = "cheque_no DESC, cnt DESC"
            ElseIf ac_code.Checked = True Then
                CntNB = "ac_code DESC, cnt DESC"
            ElseIf descrip.Checked = True Then
                CntNB = "descrip DESC, cnt DESC"
            ElseIf Book.Checked = True Then
                CntNB = "Book DESC, cnt DESC"
            ElseIf Curr.Checked = True Then
                CntNB = "Curr DESC, cnt DESC"
            End If
        End If

    End Sub


    Private Sub FmJeneralJournal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetupGrid()

        DbHelper.ExecuteNonQuery("   update gen_jn set Rate_USD=rate where curr='USD' and Rate_USD is null ")
        DbHelper.ExecuteNonQuery("  update gen_jn set Rate_USD=0 where   Rate_USD is null ")
        DbHelper.ExecuteNonQuery("  update gen_jn set amt_USD_Dr=amount_dr  where curr='USD' and (amt_USD_Dr=0 or amt_USD_Dr is null) ")
        DbHelper.ExecuteNonQuery("  update gen_jn set amt_USD_cr= amount_Cr   where curr='USD'  and (amt_USD_cr=0 or amt_USD_cr is null) ")
        DbHelper.ExecuteNonQuery("  update gen_jn set amt_USD_Dr= amt_dr/Rate_USD  where curr='LAK' and Rate_USD<>0")
        DbHelper.ExecuteNonQuery("  update gen_jn set amt_USD_cr= amt_cr/Rate_USD    where curr='LAK'  and Rate_USD<>0")
        '==============OPEN=====
        DbHelper.ExecuteNonQuery("   update Open_jn set Rate_USD=rate where curr='USD' and Rate_USD is null ")
        DbHelper.ExecuteNonQuery("  update Open_jn set Rate_USD=0 where   Rate_USD is null ")
        DbHelper.ExecuteNonQuery("  update Open_jn set amt_USD_Dr=amount_dr  where curr='USD' and (amt_USD_Dr=0 or amt_USD_Dr is null) ")
        DbHelper.ExecuteNonQuery("  update Open_jn set amt_USD_cr= amount_Cr   where curr='USD'  and (amt_USD_cr=0 or amt_USD_cr is null) ")
        DbHelper.ExecuteNonQuery("  update Open_jn set amt_USD_Dr= amt_dr/Rate_USD  where curr='LAK' and Rate_USD<>0")
        DbHelper.ExecuteNonQuery("  update Open_jn set amt_USD_cr= amt_cr/Rate_USD    where curr='LAK'  and Rate_USD<>0")
        '==============Adjust=====
        DbHelper.ExecuteNonQuery("   update AP_ACC_adjust_Item set Rate_USD=rate where curr='USD' and Rate_USD is null ")
        DbHelper.ExecuteNonQuery("  update AP_ACC_adjust_Item set Rate_USD=0 where   Rate_USD is null ")
        DbHelper.ExecuteNonQuery("  update AP_ACC_adjust_Item set amt_USD_Dr=amount_dr  where curr='USD' and (amt_USD_Dr=0 or amt_USD_Dr is null) ")
        DbHelper.ExecuteNonQuery("  update AP_ACC_adjust_Item set amt_USD_cr= amount_Cr   where curr='USD'  and (amt_USD_cr=0 or amt_USD_cr is null) ")

        DbHelper.ExecuteNonQuery("  update AP_ACC_adjust_Item set amt_USD_Dr= amt_dr/Rate_USD  where curr='LAK' and Rate_USD<>0")
        DbHelper.ExecuteNonQuery("  update AP_ACC_adjust_Item set amt_USD_cr= amt_cr/Rate_USD    where curr='LAK'  and Rate_USD<>0")


        DbHelper.ExecuteNonQuery("update gen_jn set gen_jn.ac_name=Acc_Code.name_L,gen_jn.ac_namee=Acc_Code.name_E from Acc_Code,gen_jn where gen_jn.Ac_Code=Acc_Code.Ac_Code and  gen_jn.ac_name is null ")
        certify.Checked = True
        MDESC.Checked = False
        FG.BackgroundColor = Color.White
        SetControlText(Me)
        Call loadCompany()
        LoadSubCompany()
        Off_Usr.Text = FmLogin.Sub_Company.Text
        ' FG.GridColor - property not available in DataGridView = Color.RoyalBlue
        ' FG.ForeColorFixed - property not available in DataGridView
        ' FG.BackColorFixed - property not available in DataGridView
        'Me.FG.Size = New System.Drawing.Size(20, 32)
        'FG.Anchor = AnchorStyles.Left
        'FG.Anchor = AnchorStyles.Top
        'FG.Anchor = AnchorStyles.Bottom
        'FG.Anchor = AnchorStyles.Right
        RAll.Checked = True
        ComboBox1.Items.Clear()
        Nme.Enabled = False
        Nme.Visible = True
        'RadioButton14.Checked = True
        'LoadMonthSQL()
        FG.CurrentRow.Indexs = 1

        'LoadBooks()
        FG.CurrentRow.Indexs = 2
        'CntNB = "certify , cnt"
        'SQL = " AND month(gen_jn.date_work   ) BETWEEN '" & "11" & "' AND '" & "11" & "' AND year(gen_jn.date_work )='" & 2010 & "' AND gen_jn.Company='" & MuSubOff & "' " & SR & " "
        'Load_M()
        'FG.CurrentRow.Indexs = 13
        'Panel1.Anchor = AnchorStyles.Bottom And AnchorStyles.Left
' FG.AllowUserResizing - property not available in DataGridView
        ' FG.ExtendLastCol - property not available in DataGridView
        MDInvoiceNo = ""
        'FG.AllowUserResizing = VSFlex8U.AllowUserResizeSettings.flexResizeBoth
        'FormOpening2()
        'ChgChildForm()
        CMS7.Text = "ຮູບແບບການຈັດລຽງ"
        date_work.Text = "ວັນທີ"
    End Sub
    Private Sub LoadBooks()
        ComboBox1.Items.Clear()
        Try
            Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM books WHERE bookid <> N'" & "" & "'")
            If dt.Rows.Count <> 0 Then
                ComboBox1.Items.Add("<All>  ທັງໝົດ (All books)")
                For Each row As DataRow In dt.Rows
                    ComboBox1.Items.Add(Trim(row("bookid").ToString()))
                Next
                ComboBox1.Text = "<All>  ທັງໝົດ (All books)"
            End If
        Catch ex As Exception
            VSysError = True
            MessageBox.Show("Database Error in LoadBooks: " & ex.Message, "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub LoadsLurr()
        ComboBox1.Items.Clear()
        Try
            Dim dt As DataTable = DbHelper.GetDataTable("SELECT curr FROM Ap_RateSeting WHERE curr <> N'" & "" & "'")
            If dt.Rows.Count <> 0 Then
                ComboBox1.Items.Add("<All>  ທັງໝົດ")
                For Each row As DataRow In dt.Rows
                    ComboBox1.Items.Add(Trim(row("curr").ToString()))
                Next
                ComboBox1.Text = "<All>  ທັງໝົດ"
            End If
        Catch ex As Exception
            VSysError = True
            MessageBox.Show("Database Error in LoadsLurr: " & ex.Message, "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtBook.Text = ""
        Dim dtBook As DataTable = DbHelper.GetDataTable("SELECT * FROM books WHERE bookname  = N'" & ComboBox1.Text & "'")
        If dtBook.Rows.Count > 0 Then
            txtBook.Text = Trim(dtBook.Rows(0)("bookid").ToString())
        End If
        If ComboBox1.Text = "" Then
            txtBook.Text = "All"
        End If
    End Sub

    Private Sub Load_Gen_Jn()
        Dim OfUsr1 As String = Microsoft.VisualBasic.Left(Off_Usr.Text, 5)
        Dim OfUsr2 As String = Mid(Off_Usr.Text, 4, 2)
        Dim OfUsr3 As String = Microsoft.VisualBasic.Left(Off_Usr.Text, 2)
        If OfUsr1 = "00-00" Then
            MULook2 = ""
        Else
            If OfUsr2 = "00" Then
                MULook2 = "  And  Left(company,2)= '" & OfUsr3 & "' "
            Else
                MULook2 = "  And company= '" & OfUsr1 & "' "
            End If
        End If
        txtDescrip.Text = ""
        txtAc_Code.Text = ""
        txtAc_Code.Text = GetGridValue(FG, FG.CurrentRow.Index, 4)
        txtCurr.Text = GetGridValue(FG, FG.CurrentRow.Index, 9)
        Dim str As String = ", Amount = " & GetGridValue(FG, FG.CurrentRow.Index, 4) & ": " & GetGridValue(FG, FG.CurrentRow.Index, 9)
        If MuLng = "L" Then

            Dim dt3 As DataTable = DbHelper.GetDataTable("SELECT AG  , descrip     FROM gen_jn WHERE cnt = '" & GetGridValue(FG, FG.CurrentRow.Index, 13) & "' " & MULook2 & "  order by cnt")
            If dt3.Rows.Count <> 0 Then
                AG = Trim(DbHelper.GetStr(dt3.Rows(0)("AG")))
                txtDescrip.Text = Trim(DbHelper.GetStr(dt3.Rows(0)("descrip"))) & ", ມູນຄ່າ: " & GetGridValue(FG, FG.CurrentRow.Index, 5) & ": " & GetGridValue(FG, FG.CurrentRow.Index, 9)
            End If
            Dim dt4 As DataTable = DbHelper.GetDataTable("SELECT   Name_L  FROM Acc_Code WHERE Ac_Code = '" & txtAc_Code.Text & "'")
            If dt4.Rows.Count <> 0 Then
                Ac_Name.Text = Trim(DbHelper.GetStr(dt4.Rows(0)("Name_L")))
                'x = Ac_Name.Text
                'MsgBox(x)
            End If

            '========================
        Else
            Dim dt5 As DataTable = DbHelper.GetDataTable("SELECT AG  , descripe    FROM gen_jn WHERE cnt = '" & GetGridValue(FG, FG.CurrentRow.Index, 13) & "' " & MULook2 & "  order by cnt")
            If dt5.Rows.Count <> 0 Then
                AG = Trim(DbHelper.GetStr(dt5.Rows(0)("AG")))
                txtDescrip.Text = Trim(DbHelper.GetStr(dt5.Rows(0)("descripe"))) & ", Amount: " & GetGridValue(FG, FG.CurrentRow.Index, 5) & ": " & GetGridValue(FG, FG.CurrentRow.Index, 9)
            End If
            Dim dt6 As DataTable = DbHelper.GetDataTable("SELECT   Name_E FROM Acc_Code WHERE Ac_Code = '" & txtAc_Code.Text & "'")
            If dt6.Rows.Count <> 0 Then
                Ac_Name.Text = Trim(DbHelper.GetStr(dt6.Rows(0)("Name_E")))
                'x = Ac_Name.Text
            End If
        End If

        ' Dim RSC1 As New ADODB.Recordset ' REMOVED - ADODB migration
        Try
            Dim s As String = "SELECT sum(Amt_dr) as Amt_dr   , sum(Amt_cr) as Amt_cr    FROM gen_jn WHERE ac_code = '" & txtAc_Code.Text & "' and gen_jn.date_work BETWEEN '" & Format(dts.Value, "yyyy-MM-dd") & "' AND '" & Format(dtt.Value, "yyyy-MM-dd") & "'   " & MULook2 & "  "
            Dim dt1 As DataTable = DbHelper.GetDataTable(s)
            If dt1.Rows.Count <> 0 Then
                SumDr.Text = Format(CDbl(dt1.Rows(0)("Amt_dr").ToString()), "#,##0.00")
                SumCr.Text = Format(CDbl(dt1.Rows(0)("Amt_cr").ToString()), "#,##0.00")
            End If

            Dim dtOpen As DataTable = DbHelper.GetDataTable("select  amount_dr , amount_cr from Open_jn where ac_code='" & txtAc_Code.Text & "'   and  year(Date_work)= '" & Format(CDate(dts.Value), "yyyy") & "'  " & MULook2 & "   ")
            Op = 0
            If dtOpen.Rows.Count <> 0 Then
                Op = CDbl(dtOpen.Rows(0)("amount_dr").ToString()) - CDbl(dtOpen.Rows(0)("amount_cr").ToString())
            End If
            Dim dss As Date
            dss = DateAdd(DateInterval.Day, -1, dts.Value)
            Dim dt2 As DataTable = DbHelper.GetDataTable("select SUM(amount_dr) AS amount_dr ,SUM(amount_cr) AS amount_cr from Gen_jn where ac_code=N'" & txtAc_Code.Text & "'  And gen_jn.date_work   BETWEEN '" & "1-1-" & Format(dts.Value, "yyyy") & "' AND '" & Format(dss, "yyyy-MM-dd") & "' " & MULook2 & " group by ac_code ")
            If dt2.Rows.Count <> 0 Then
                Op = Op + CDbl(CDbl(dt2.Rows(0)("amount_dr").ToString()) - CDbl(dt2.Rows(0)("amount_Cr").ToString()))
            End If
        Catch ex As Exception
            VSysError = True
            MessageBox.Show("Database Error: " & ex.Message, "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        If Op >= 0 Then
            Open_jn.ForeColor = Color.Black
            Open_jn.Text = Format(CDbl(Op), "##,##0.00")
        Else
            Open_jn.ForeColor = Color.Red
            Open_jn.Text = "(" & Format(CDbl(Op * (-1)), "##,##0.00") & ")"
        End If

    End Sub
    Public Sub SumAmount()

        TotalDr.Text = "0.00"
        TotalCr.Text = "0.00"
        Balance.Text = "0.00"
        Dim s As String = "SELECT sum(Amt_dr) as Amt_dr , sum(Amt_cr) as Amt_cr   FROM gen_jn WHERE gen_jn.date_work BETWEEN '" & Format(dts.Value, "yyyy-MM-dd") & "' AND '" & Format(dtt.Value, "yyyy-MM-dd") & "'   " & MULook2 & "  "
        Dim dtTemp As DataTable = DbHelper.GetDataTable(s)

        If dtTemp.Rows.Count <> 0 Then
            TotalDr.Text = DbHelper.GetStr(dtTemp.Rows(0)("Amt_dr"))
            TotalCr.Text = DbHelper.GetStr(dtTemp.Rows(0)("Amt_cr"))
        End If
        If TotalDr.Text = "" Then
            TotalDr.Text = "0.00"
        End If
        If TotalCr.Text = "" Then
            TotalCr.Text = "0.00"
        End If

        Balance.Text = CDbl(TotalDr.Text) - CDbl(TotalCr.Text)
        TotalDr.Text = Format(CDbl(TotalDr.Text), "#,##0.00")
        TotalCr.Text = Format(CDbl(TotalCr.Text), "#,##0.00")
        Balance.Text = Format(CDbl(Balance.Text), "#,##0.00")

    End Sub
    Public Sub LoadSQL()


        'SQL = ""
        'Dim Yr As Integer
        'Yr = Year(Today)
        'If RadioButton1.Checked = True Then
        '    SQL = " AND month(gen_jn.date_work   ) BETWEEN '" & Month(dts.Value) & "' AND '" & Month(dtt.Value) & "' AND year(gen_jn.date_work )='" & Yr & "' "
        'End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        LoadMonthSQL()
        'Call StartLoadDataList()
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        LoadMonthSQL()

    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        LoadMonthSQL()
    End Sub

    Private Sub RadioButton5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton5.CheckedChanged

        LoadMonthSQL()
    End Sub

    Private Sub RadioButton6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton6.CheckedChanged
        LoadMonthSQL()

    End Sub

    Private Sub RadioButton8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton8.CheckedChanged
        LoadMonthSQL()
    End Sub

    Private Sub RadioButton9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton9.CheckedChanged
        LoadMonthSQL()
    End Sub

    Private Sub RadioButton10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton10.CheckedChanged
        LoadMonthSQL()
    End Sub

    Private Sub RadioButton7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton7.CheckedChanged
        LoadMonthSQL()
    End Sub

    Private Sub RadioButton11_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton11.CheckedChanged
        LoadMonthSQL()
    End Sub

    Private Sub RadioButton12_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton12.CheckedChanged
        LoadMonthSQL()
    End Sub

    Private Sub RadioButton13_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton13.CheckedChanged
        LoadMonthSQL()
    End Sub

    Private Sub RadioButton14_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton14.CheckedChanged
        LoadMonthSQL()

    End Sub

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click

        Call MdiCNum()
        'FmMain.PictureBox1.Visible = True
        MDInvoiceNo = ""
        Close()

    End Sub

    Private Sub FG_ClickEvent(ByVal sender As Object, ByVal e As System.EventArgs)
        LockData = GetGridValue(FG, FG.CurrentRow.Index, 13)
        If GetGridValue(FG, FG.CurrentRow.Index, 14) = 1 Then
            Button1.Text = "ປົດລອກ"
        End If
        If GetGridValue(FG, FG.CurrentRow.Index, 14) = 0 Then
            Button1.Text = "ລອກຂໍ້ມູນ"
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Panel7.Visible = True



    End Sub






    Private Sub FG_Resize(ByVal sender As Object, ByVal e As System.EventArgs)
        '4444
    End Sub

    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub FmJeneralJournal_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        'FG.Width = Me.Width - 18
        'FG.Height = Me.Height - 340
        'Panel1.Location = New System.Drawing.Point(5, CDbl(FG.Height) - CDbl(-115))
    End Sub


    Private Sub BtnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnEdit.Click
        If LockData = 1 Then
            MsgBox("ບັນຊີນີ້ຖືກລ໋ອກບໍ່ສາມາດແກ້ໄຂໄດ້")
            Exit Sub
        End If
        If LockData = 2 Then
            MsgBox("ບັນຊີນີ້ໄດ້ປິດບັນຊີໄປແລ້ວກບໍ່ສາມາດແກ້ໄຂໄດ້")
            Exit Sub
        End If

        Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT * FROM gen_jn where certify=N'" & MDInvoiceNo & "' and My_Lock=1 ")
        If dtTemp.Rows.Count <> 0 Then
            MsgBox("ລາຍການນີ້ບໍ່ສາມາດແກ້ໄຂໄດ້", MsgBoxStyle.Exclamation)
            Exit Sub
        End If


        Panel4.Visible = False
        If MDInvoiceNo <> "" Then
            FmNsewJeneralJournal.txtInvoice.Enabled = False
            FmNsewJeneralJournal.CmbBook.Enabled = False
            'FmNsewJeneralJournal.ShowDialog()
            FmNsewJeneralJournal.Show()
        Else
            MessageBox.Show("ກະລຸນນາເລືອກກ່ອນ")
            Exit Sub
        End If
    End Sub

    Private Sub BtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelete.Click
        Panel4.Visible = False
        If MDInvoiceNo = "" Then
            MessageBox.Show("ກະລຸນນາເລືອກກ່ອນ", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If LockData = 1 Then
            MessageBox.Show("ບັນຊີນີ້ຖືກລ໋ອກບໍ່ສາມາດລືບໄດ້", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If LockData = 2 Then
            MessageBox.Show("ບັນຊີນີ້ໄດ້ປິດບັນຊີໄປແລ້ວກບໍ່ສາມາດລືບໄດ້", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If MessageBox.Show("ທ່ານຕ້ອງການລຶບ  " & MDInvoiceNo & " ຫລືບໍ່?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

            'DbHelper.ExecuteNonQuery("delete gen_jn where certify =N'" & MDInvoiceNo & "' And   date_work='" & Format(CDate(GetGridValue(FG, FG.CurrentRow.Index, 1)), "yyyy-MM-dd") & "' ")
            DbHelper.ExecuteNonQuery("delete gen_jn where certify =N'" & MDInvoiceNo & "' and  ReferNO =N'" & MDInvoice_RefNo & "'   And   date_work='" & Format(CDate(GetGridValue(FG, FG.CurrentRow.Index, 1)), "yyyy-MM-dd") & "' ")

            DbHelper.ExecuteNonQuery("update Adjustment_List set  Remain=Remain+ " & CDbl(MDUPASETAMT) & "  where Code=N'" & Trim(MDUPASET) & "' ")

        End If
        MDInvoiceNo = ""
        LoadMonthSQL()
    End Sub

    Private Sub FG_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles FG.SelectionChanged
        If FG.CurrentRow.Indexs > 7 Then
            LngId = "8008" : CallLngStr()
            FG.FormatString = LngStr
        Else
            LngId = "8001" : CallLngStr()
            FG.FormatString = LngStr
        End If
    End Sub

    Private Sub FG_Scroll(ByVal sender As Object, ByVal e As ScrollEventArgs) Handles FG.Scroll
        BtnSearch.Visible = False
    End Sub

    Private Sub FG_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FG.CellClick
        MDInvoiceDT = GetGridValue(FG, FG.CurrentRow.Index, 1)
        MDInvoiceNo = GetGridValue(FG, FG.CurrentRow.Index, 2)
        'Call Load_Gen_Jn()

        'MessageBox.Show(GetGridValue(FG, FG.CurrentRow.Index, 2))
    End Sub

    Private Sub FG_CellDoubleClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FG.CellDoubleClick
        If MDInvoiceNo <> "" Then
            FmNsewJeneralJournal.txtInvoice.Enabled = False
            FmNsewJeneralJournal.CmbBook.Enabled = False
            FmNsewJeneralJournal.Show()
        End If



    End Sub

    Private Sub AssetUP()
        Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT * FROM gen_jn where certify=N'" & MDInvoiceNo & "' and My_Lock=1 ")
        If dtTemp.Rows.Count <> 0 Then
            MDUPASET = DbHelper.GetStr(dtTemp.Rows(0)("Referno_Item"))
            MDUPASETAMT = Format(CDbl(DbHelper.GetStr(dtTemp.Rows(0)("amount"))), "#,##0.00")
        Else
            MDUPASET = ""
            MDUPASETAMT = 0
        End If
    End Sub
    Private Sub FG_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs) Handles FG.MouseUp

        ' FG.FocusRect - property not available in DataGridView
        If x > 0 Then
            x0 = x
            s0 = TextBox1.Text
        End If
        If y > 0 Then
            y0 = y
        End If
        x = FG.CurrentCell.ColumnIndex
        y = FG.CurrentRow.Index
        MDInvoiceDT = GetGridValue(FG, FG.CurrentRow.Index, 1)
        MDInvoiceNo = GetGridValue(FG, FG.CurrentRow.Index, 2)
        MDInvoice_RefNo = GetGridValue(FG, FG.CurrentRow.Index, 3)
        TextBox1.Text = GetGridValue(FG, FG.CurrentRow.Index, 2)
        Call AssetUP()
        If MuLng = "L" Then
            txtDescrip.Text = GetGridValue(FG, FG.CurrentRow.Index, 6) & ", ມູນຄ່າ: " & GetGridValue(FG, FG.CurrentRow.Index, 5) & ": " & GetGridValue(FG, FG.CurrentRow.Index, 9)
        Else
            txtDescrip.Text = GetGridValue(FG, FG.CurrentRow.Index, 6) & ", Amout: " & GetGridValue(FG, FG.CurrentRow.Index, 5) & ": " & GetGridValue(FG, FG.CurrentRow.Index, 9)
        End If



        LockData = ""
        LockData = GetGridValue(FG, FG.CurrentRow.Index, 14)
        If LockData = 1 Then
            RadioButton17.Checked = False
            RadioButton18.Checked = True
        Else
            RadioButton17.Checked = True
            RadioButton18.Checked = False
        End If
        If LockData = 1 Or LockData = 2 Then
            LngId = "3027" : CallLngStr() : Button1.Text = LngStr
        Else
            LngId = "3008" : CallLngStr() : Button1.Text = LngStr
        End If
        'Remain.Text = "0.00"
        If FG.CurrentRow.Index > 0 Then
            If FG.CurrentCell.ColumnIndex = 4 Then
                Call Load_Gen_Jn()
                Remain.Text = CDbl(CDbl(Open_jn.Text) + CDbl(SumDr.Text)) - CDbl(SumCr.Text)

                If Remain.Text >= 0 Then
                    Remain.ForeColor = Color.Black
                    Remain.Text = Format(CDbl(Remain.Text), "##,##0.00")
                Else
                    Remain.ForeColor = Color.Red
                    Remain.Text = "(" & Format(CDbl(Remain.Text * (-1)), "##,##0.00") & ")"
                End If
            End If
        End If










        If CheckBox3.Checked = True Then
            ForIs1()
            ForIt1()
            ForIs2()
            ForIt2()
            Dim ne As Integer
            For ne = Rs2 To Rt2
                If Trim(FG.get_TextMatrix(ne, 2)) = s0 Then
                    FG.CurrentRow.Index = ne
                    FG.CurrentCell.ColumnIndex = 2
                    FG.CurrentRow.DefaultCellStyle.BackColor = Color.White
                    FG.CurrentCell.ColumnIndex = 4
                    FG.CurrentRow.DefaultCellStyle.BackColor = Color.White
                    'MsgBox(Trim(FG.get_TextMatrix(J, 2)))
                    If Trim(FG.get_TextMatrix(ne, 10)) <> 0 Then
                        FG.CurrentCell.ColumnIndex = 7
                        FG.CurrentRow.DefaultCellStyle.BackColor = Color.White
                        FG.CurrentCell.ColumnIndex = 10
                        FG.CurrentRow.DefaultCellStyle.BackColor = Color.White
                    Else
                        FG.CurrentCell.ColumnIndex = 8
                        FG.CurrentRow.DefaultCellStyle.BackColor = Color.White
                        FG.CurrentCell.ColumnIndex = 11
                        FG.CurrentRow.DefaultCellStyle.BackColor = Color.White
                    End If
                End If
            Next ne









            For J = Rs1 To Rt1
                'MsgBox(J & " = " & FG.get_TextMatrix(J, 4))
                If Trim(FG.get_TextMatrix(J, 2)) = TextBox1.Text Then
                    FG.CurrentRow.Index = J
                    FG.CurrentCell.ColumnIndex = 2
                    FG.CurrentRow.DefaultCellStyle.BackColor = Color.SkyBlue
                    FG.CurrentCell.ColumnIndex = 4
                    FG.CurrentRow.DefaultCellStyle.BackColor = Color.SkyBlue
                    'MsgBox(Trim(FG.get_TextMatrix(J, 2)))
                    If Trim(FG.get_TextMatrix(J, 10)) <> 0 Then
                        FG.CurrentCell.ColumnIndex = 7
                        FG.CurrentRow.DefaultCellStyle.BackColor = Color.SkyBlue
                        FG.CurrentCell.ColumnIndex = 10
                        FG.CurrentRow.DefaultCellStyle.BackColor = Color.SkyBlue
                    Else
                        FG.CurrentCell.ColumnIndex = 8
                        FG.CurrentRow.DefaultCellStyle.BackColor = Color.SkyBlue
                        FG.CurrentCell.ColumnIndex = 11
                        FG.CurrentRow.DefaultCellStyle.BackColor = Color.SkyBlue
                    End If
                End If
            Next J


            If Trim(FG.get_TextMatrix(FG.CurrentRow.Indexs - 1, 2)) = TextBox1.Text <> 0 Then
                FG.CurrentRow.Index = FG.CurrentRow.Indexs - 1
                FG.CurrentCell.ColumnIndex = 2
                FG.CurrentRow.DefaultCellStyle.BackColor = Color.SkyBlue
                FG.CurrentCell.ColumnIndex = 4
                FG.CurrentRow.DefaultCellStyle.BackColor = Color.SkyBlue
                'MsgBox(Trim(FG.get_TextMatrix(J, 2)))
                If Trim(FG.get_TextMatrix(FG.CurrentRow.Indexs - 1, 10)) <> 0 Then
                    FG.CurrentCell.ColumnIndex = 7
                    FG.CurrentRow.DefaultCellStyle.BackColor = Color.SkyBlue
                    FG.CurrentCell.ColumnIndex = 10
                    FG.CurrentRow.DefaultCellStyle.BackColor = Color.SkyBlue
                Else
                    FG.CurrentCell.ColumnIndex = 8
                    FG.CurrentRow.DefaultCellStyle.BackColor = Color.SkyBlue
                    FG.CurrentCell.ColumnIndex = 11
                    FG.CurrentRow.DefaultCellStyle.BackColor = Color.SkyBlue
                End If
            Else
                FG.CurrentRow.Index = FG.CurrentRow.Indexs - 1
                FG.CurrentCell.ColumnIndex = 2
                FG.CurrentRow.DefaultCellStyle.BackColor = Color.White
                FG.CurrentCell.ColumnIndex = 4
                FG.CurrentRow.DefaultCellStyle.BackColor = Color.White
                'MsgBox(Trim(FG.get_TextMatrix(J, 2)))
                If Trim(FG.get_TextMatrix(FG.CurrentRow.Indexs - 1, 10)) <> 0 Then
                    FG.CurrentCell.ColumnIndex = 7
                    FG.CurrentRow.DefaultCellStyle.BackColor = Color.White
                    FG.CurrentCell.ColumnIndex = 10
                    FG.CurrentRow.DefaultCellStyle.BackColor = Color.White
                Else
                    FG.CurrentCell.ColumnIndex = 8
                    FG.CurrentRow.DefaultCellStyle.BackColor = Color.White
                    FG.CurrentCell.ColumnIndex = 11
                    FG.CurrentRow.DefaultCellStyle.BackColor = Color.White
                End If
            End If
        End If
        FG.CurrentCell.ColumnIndex = x
        FG.CurrentRow.Index = y
        'If MDInvoiceNo <> "" Then

        CMS2.Enabled = True
        CMS3.Enabled = True
        'End If
    End Sub

    Private Sub ForIs2()

        If y0 = FG.CurrentRow.Indexs - 1 Then
            If FG.get_TextMatrix(FG.CurrentRow.Indexs - 1, 2) <> FG.get_TextMatrix(FG.CurrentRow.Indexs - 2, 2) Then
                Rs2 = FG.CurrentRow.Indexs - 1
            End If
        End If
        Dim x0k As Integer = y0
        For i = y0 To y0 * 2
            Rs2 = x0k

            x0k = x0k - 1
            s = FG.get_TextMatrix(x0k, 2)
            'R1 = x + 1
            If s <> s0 Then
                Exit Sub
            End If
        Next
    End Sub

    Private Sub ForIt2()
        If y0 = FG.CurrentRow.Indexs - 1 Then
            Rt2 = FG.CurrentRow.Indexs - 1
            Exit Sub
        End If
        For i = y0 To FG.CurrentRow.Indexs - 1
            Rt2 = i - 1
            s = FG.get_TextMatrix(i, 2)
            If s <> s0 Then
                Exit Sub
            End If
        Next
    End Sub
    Private Sub ForIs1()

        If y = FG.CurrentRow.Indexs - 1 Then
            If FG.get_TextMatrix(FG.CurrentRow.Indexs - 1, 2) <> FG.get_TextMatrix(FG.CurrentRow.Indexs - 2, 2) Then
                Rs1 = FG.CurrentRow.Indexs - 1
            End If
        End If
        Dim x As Integer = y
        For i = y To y * 2
            Rs1 = x
            x = x - 1
            s = FG.get_TextMatrix(x, 2)
            'R1 = x + 1
            If s <> TextBox1.Text Then
                Exit Sub
            End If
        Next
    End Sub

    Private Sub ForIt1()




        For i = y To FG.CurrentRow.Indexs - 1
            Rt1 = i - 1
            s = FG.get_TextMatrix(i, 2)
            If s <> TextBox1.Text Then
                Exit Sub
            End If
        Next
    End Sub
    Private Sub FG_SelectionChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelectionChanged

    End Sub
    Private Sub ReportGen_jn()
        Dim Gen_jn, Open_jn, LopGen_jn As String
        Dim Yr, Yrl As Integer

        Yr = Year(Today)
        Yrl = Year(Today) - 1

        Gen_jn = "" : Open_jn = "" : LopGen_jn = ""
        If txtBook.Text <> "All" Then
            Gen_jn = Gen_jn & " AND Gen_jn.book, '" & Len(txtBook.Text.Trim)
            'Open_jn = Open_jn & " AND Open_jn.book, '" & Len(txtBook.Text.Trim)
        End If
        If RadioButton1.Checked = True Then
            Gen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "1" & "' AND '" & "1" & "' AND year(gen_jn.date_work )='" & Yr & "' AND gen_jn.Company=N'" & MuSubOff & "' "
            LopGen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "12" & "' AND '" & "12" & "' AND year(gen_jn.date_work )='" & Yrl & "' AND gen_jn.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "1" & "' AND '" & "1" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            DbHelper.ExecuteNonQuery(" DELETE FROM Ap_Sum_Gen_jn ")
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_Sum_Gen_jn ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                    " FROM gen_jn  WHERE 1=1 " & Gen_jn & " ")
            '==============================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'   ")
                Next
            Else
                Exit Sub
            End If
            '===========================
            Dim dtTemp2 As DataTable = DbHelper.GetDataTable("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM gen_jn  WHERE 1=1 " & LopGen_jn & " GROUP BY Company,Certify ")
            If dtTemp2.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp2.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE Certify=N'" & DbHelper.GetStr(row("Certify")) & "' AND Company=N'" & DbHelper.GetStr(row("Company")) & "'  ")
                Next
            Else
                Exit Sub
            End If
            '=======================================
            Dim dtTemp3 As DataTable = DbHelper.GetDataTable("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_Gen_jn")
            If dtTemp3.Rows.Count > 0 Then
                For Each row As DataRow In dtTemp3.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Rem_Amt= +( " & CDbl(DbHelper.GetStr(row("Open_Amt"))) & "+" & CDbl(DbHelper.GetStr(row("amount_Dr"))) & " -" & CDbl(DbHelper.GetStr(row("amount_Cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'  ")

                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Rem_Amt  ")
                Next
            Else
                Exit Sub
            End If
        End If
        If RadioButton2.Checked = True Then
            Gen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "2" & "' AND '" & "2" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            LopGen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "1" & "' AND '" & "1" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "2" & "' AND '" & "2" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            DbHelper.ExecuteNonQuery(" DELETE FROM Ap_Sum_Gen_jn ")
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_Sum_Gen_jn ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                    " FROM gen_jn  WHERE 1=1 " & Gen_jn & " ")
            '==============================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'   ")
                Next
            Else
                Exit Sub
            End If
            '===========================
            Dim dtTemp2 As DataTable = DbHelper.GetDataTable("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM gen_jn  WHERE 1=1 " & LopGen_jn & " GROUP BY Company,Certify ")
            If dtTemp2.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp2.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE Certify=N'" & DbHelper.GetStr(row("Certify")) & "' AND Company=N'" & DbHelper.GetStr(row("Company")) & "'  ")
                Next
            Else
                Exit Sub
            End If
            '=======================================
            Dim dtTemp3 As DataTable = DbHelper.GetDataTable("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_Gen_jn")
            If dtTemp3.Rows.Count > 0 Then
                For Each row As DataRow In dtTemp3.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Rem_Amt= +( " & CDbl(DbHelper.GetStr(row("Open_Amt"))) & "+" & CDbl(DbHelper.GetStr(row("amount_Dr"))) & " -" & CDbl(DbHelper.GetStr(row("amount_Cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'  ")

                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Rem_Amt  ")
                Next
            Else
                Exit Sub
            End If
        End If
        If RadioButton3.Checked = True Then
            Gen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "3" & "' AND '" & "3" & "' AND year(gen_jn.date_work )='" & Yr & "AND gen_jn.Company=N'" & MuSubOff & "'' "
            LopGen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "2" & "' AND '" & "2" & "' AND year(gen_jn.date_work )='" & Yr & "AND gen_jn.Company=N'" & MuSubOff & "'' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "3" & "' AND '" & "3" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            DbHelper.ExecuteNonQuery(" DELETE FROM Ap_Sum_Gen_jn ")
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_Sum_Gen_jn ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                    " FROM gen_jn  WHERE 1=1 " & Gen_jn & " ")
            '==============================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'   ")
                Next
            Else
                Exit Sub
            End If
            '===========================
            Dim dtTemp2 As DataTable = DbHelper.GetDataTable("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM gen_jn  WHERE 1=1 " & LopGen_jn & " GROUP BY Company,Certify ")
            If dtTemp2.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp2.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE Certify=N'" & DbHelper.GetStr(row("Certify")) & "' AND Company=N'" & DbHelper.GetStr(row("Company")) & "'  ")
                Next
            Else
                Exit Sub
            End If
            '=======================================
            Dim dtTemp3 As DataTable = DbHelper.GetDataTable("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_Gen_jn")
            If dtTemp3.Rows.Count > 0 Then
                For Each row As DataRow In dtTemp3.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Rem_Amt= +( " & CDbl(DbHelper.GetStr(row("Open_Amt"))) & "+" & CDbl(DbHelper.GetStr(row("amount_Dr"))) & " -" & CDbl(DbHelper.GetStr(row("amount_Cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'  ")

                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Rem_Amt  ")
                Next
            Else
                Exit Sub
            End If
        End If
        If RadioButton4.Checked = True Then
            Gen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "4" & "' AND '" & "4" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            LopGen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "3" & "' AND '" & "3" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "4" & "' AND '" & "4" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            DbHelper.ExecuteNonQuery(" DELETE FROM Ap_Sum_Gen_jn ")
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_Sum_Gen_jn ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM gen_jn  WHERE 1=1 " & Gen_jn & " ")
            '==============================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'   ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '===========================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM gen_jn  WHERE 1=1 " & LopGen_jn & " GROUP BY Company,Certify ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE Certify=N'" & DbHelper.GetStr(row("Certify")) & "' AND Company=N'" & DbHelper.GetStr(row("Company")) & "'  ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_Gen_jn")
            If dtTemp.Rows.Count > 1 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Rem_Amt= +( " & CDbl(DbHelper.GetStr(row("Open_Amt"))) & "+" & CDbl(DbHelper.GetStr(row("amount_dr"))) & " -" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'  ")

                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Rem_Amt  ")
                    Next
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton5.Checked = True Then
            Gen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "5" & "' AND '" & "5" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            LopGen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "4" & "' AND '" & "4" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "5" & "' AND '" & "5" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            DbHelper.ExecuteNonQuery(" DELETE FROM Ap_Sum_Gen_jn ")
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_Sum_Gen_jn ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM gen_jn  WHERE 1=1 " & Gen_jn & " ")
            '==============================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'   ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '===========================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM gen_jn  WHERE 1=1 " & LopGen_jn & " GROUP BY Company,Certify ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE Certify=N'" & DbHelper.GetStr(row("Certify")) & "' AND Company=N'" & DbHelper.GetStr(row("Company")) & "'  ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_Gen_jn")
            If dtTemp.Rows.Count > 1 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Rem_Amt= +( " & CDbl(DbHelper.GetStr(row("Open_Amt"))) & "+" & CDbl(DbHelper.GetStr(row("amount_dr"))) & " -" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'  ")

                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Rem_Amt  ")
                    Next
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton6.Checked = True Then
            Gen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "6" & "' AND '" & "6" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            LopGen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "5" & "' AND '" & "5" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "6" & "' AND '" & "6" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            DbHelper.ExecuteNonQuery(" DELETE FROM Ap_Sum_Gen_jn ")
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_Sum_Gen_jn ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM gen_jn  WHERE 1=1 " & Gen_jn & " ")
            '==============================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'   ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '===========================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM gen_jn  WHERE 1=1 " & LopGen_jn & " GROUP BY Company,Certify ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE Certify=N'" & DbHelper.GetStr(row("Certify")) & "' AND Company=N'" & DbHelper.GetStr(row("Company")) & "'  ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_Gen_jn")
            If dtTemp.Rows.Count > 1 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Rem_Amt= +( " & CDbl(DbHelper.GetStr(row("Open_Amt"))) & "+" & CDbl(DbHelper.GetStr(row("amount_dr"))) & " -" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'  ")

                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Rem_Amt  ")
                    Next
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton7.Checked = True Then
            Gen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "7" & "' AND '" & "7" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            LopGen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "6" & "' AND '" & "6" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "7" & "' AND '" & "7" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            DbHelper.ExecuteNonQuery(" DELETE FROM Ap_Sum_Gen_jn ")
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_Sum_Gen_jn ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM gen_jn  WHERE 1=1 " & Gen_jn & " ")
            '==============================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'   ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '===========================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM gen_jn  WHERE 1=1 " & LopGen_jn & " GROUP BY Company,Certify ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE Certify=N'" & DbHelper.GetStr(row("Certify")) & "' AND Company=N'" & DbHelper.GetStr(row("Company")) & "'  ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_Gen_jn")
            If dtTemp.Rows.Count > 1 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Rem_Amt= +( " & CDbl(DbHelper.GetStr(row("Open_Amt"))) & "+" & CDbl(DbHelper.GetStr(row("amount_dr"))) & " -" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'  ")

                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Rem_Amt  ")
                    Next
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton8.Checked = True Then
            Gen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "8" & "' AND '" & "8" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            LopGen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "7" & "' AND '" & "7" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "8" & "' AND '" & "8" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            DbHelper.ExecuteNonQuery(" DELETE FROM Ap_Sum_Gen_jn ")
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_Sum_Gen_jn ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM gen_jn  WHERE 1=1 " & Gen_jn & " ")
            '==============================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'   ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '===========================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM gen_jn  WHERE 1=1 " & LopGen_jn & " GROUP BY Company,Certify ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE Certify=N'" & DbHelper.GetStr(row("Certify")) & "' AND Company=N'" & DbHelper.GetStr(row("Company")) & "'  ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_Gen_jn")
            If dtTemp.Rows.Count > 1 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Rem_Amt= +( " & CDbl(DbHelper.GetStr(row("Open_Amt"))) & "+" & CDbl(DbHelper.GetStr(row("amount_dr"))) & " -" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'  ")

                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Rem_Amt  ")
                    Next
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton9.Checked = True Then
            Gen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "9" & "' AND '" & "9" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            LopGen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "8" & "' AND '" & "8" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "9" & "' AND '" & "9" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            DbHelper.ExecuteNonQuery(" DELETE FROM Ap_Sum_Gen_jn ")
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_Sum_Gen_jn ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM gen_jn  WHERE 1=1 " & Gen_jn & " ")
            '==============================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'   ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '===========================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM gen_jn  WHERE 1=1 " & LopGen_jn & " GROUP BY Company,Certify ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE Certify=N'" & DbHelper.GetStr(row("Certify")) & "' AND Company=N'" & DbHelper.GetStr(row("Company")) & "'  ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_Gen_jn")
            If dtTemp.Rows.Count > 1 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Rem_Amt= +( " & CDbl(DbHelper.GetStr(row("Open_Amt"))) & "+" & CDbl(DbHelper.GetStr(row("amount_dr"))) & " -" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'  ")

                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Rem_Amt  ")
                    Next
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton10.Checked = True Then
            Gen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "10" & "' AND '" & "10" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            LopGen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "9" & "' AND '" & "9" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "10" & "' AND '" & "10" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            DbHelper.ExecuteNonQuery(" DELETE FROM Ap_Sum_Gen_jn ")
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_Sum_Gen_jn ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM gen_jn  WHERE 1=1 " & Gen_jn & " ")
            '==============================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'   ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '===========================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM gen_jn  WHERE 1=1 " & LopGen_jn & " GROUP BY Company,Certify ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE Certify=N'" & DbHelper.GetStr(row("Certify")) & "' AND Company=N'" & DbHelper.GetStr(row("Company")) & "'  ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_Gen_jn")
            If dtTemp.Rows.Count > 1 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Rem_Amt= +( " & CDbl(DbHelper.GetStr(row("Open_Amt"))) & "+" & CDbl(DbHelper.GetStr(row("amount_dr"))) & " -" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'  ")

                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Rem_Amt  ")
                    Next
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton11.Checked = True Then
            Gen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "11" & "' AND '" & "11" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            LopGen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "10" & "' AND '" & "10" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "11" & "' AND '" & "11" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            DbHelper.ExecuteNonQuery(" DELETE FROM Ap_Sum_Gen_jn ")
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_Sum_Gen_jn ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM gen_jn  WHERE 1=1 " & Gen_jn & " ")
            '==============================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'   ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '===========================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM gen_jn  WHERE 1=1 " & LopGen_jn & " GROUP BY Company,Certify ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE Certify=N'" & DbHelper.GetStr(row("Certify")) & "' AND Company=N'" & DbHelper.GetStr(row("Company")) & "'  ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_Gen_jn")
            If dtTemp.Rows.Count > 1 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Rem_Amt= +( " & CDbl(DbHelper.GetStr(row("Open_Amt"))) & "+" & CDbl(DbHelper.GetStr(row("amount_dr"))) & " -" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'  ")

                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Rem_Amt  ")
                    Next
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton12.Checked = True Then
            Gen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "12" & "' AND '" & "12" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            LopGen_jn = " AND month(gen_jn.date_work   ) BETWEEN '" & "11" & "' AND '" & "11" & "' AND year(gen_jn.date_work )='" & Yr & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "12" & "' AND '" & "12" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            DbHelper.ExecuteNonQuery(" DELETE FROM Ap_Sum_Gen_jn ")
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_Sum_Gen_jn ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM gen_jn  WHERE 1=1 " & Gen_jn & " ")
            '==============================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'   ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '===========================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM gen_jn  WHERE 1=1 " & LopGen_jn & " GROUP BY Company,Certify ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE Certify=N'" & DbHelper.GetStr(row("Certify")) & "' AND Company=N'" & DbHelper.GetStr(row("Company")) & "'  ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_Gen_jn")
            If dtTemp.Rows.Count > 1 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Rem_Amt= +( " & CDbl(DbHelper.GetStr(row("Open_Amt"))) & "+" & CDbl(DbHelper.GetStr(row("amount_dr"))) & " -" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'  ")

                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Rem_Amt  ")
                    Next
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton13.Checked = True Then
            Gen_jn = " AND year(gen_jn.date_work ) BETWEEN '" & Year(dts.Value) & "' AND '" & Year(dtt.Value) & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            LopGen_jn = " AND year(gen_jn.date_work ) BETWEEN '" & Year(dts.Value) - 1 & "' AND '" & Year(dtt.Value) - 1 & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            Open_jn = " AND year(Open_jn.date_work ) BETWEEN '" & Year(dts.Value) & "' AND '" & Year(dtt.Value) - 1 & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            DbHelper.ExecuteNonQuery(" DELETE FROM Ap_Sum_Gen_jn ")
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_Sum_Gen_jn ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM gen_jn  WHERE 1=1 " & Gen_jn & " ")
            '==============================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'   ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '===========================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM gen_jn  WHERE 1=1 " & LopGen_jn & " GROUP BY Company,Certify ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE Certify=N'" & DbHelper.GetStr(row("Certify")) & "' AND Company=N'" & DbHelper.GetStr(row("Company")) & "'  ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_Gen_jn")
            If dtTemp.Rows.Count > 1 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Rem_Amt= +( " & CDbl(DbHelper.GetStr(row("Open_Amt"))) & "+" & CDbl(DbHelper.GetStr(row("amount_dr"))) & " -" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'  ")

                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Rem_Amt  ")
                    Next
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton14.Checked = True Then
            Gen_jn = " AND gen_jn.date_work   BETWEEN '" & Format(dts.Value, "yyyy-MM-dd") & "' AND '" & Format(dtt.Value, "yyyy-MM-dd") & "'AND gen_jn.Company=N'" & MuSubOff & "' "
            LopGen_jn = " AND gen_jn.date_work  < '" & Format(dts.Value, "yyyy-MM-dd") & "' AND gen_jn.Company=N'" & MuSubOff & "' "
            Open_jn = " AND Open_jn.date_work   BETWEEN '" & Format(dts.Value, "yyyy-MM-dd") & "' AND '" & Format(dtt.Value, "yyyy-MM-dd") & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            DbHelper.ExecuteNonQuery(" DELETE FROM Ap_Sum_Gen_jn ")
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_Sum_Gen_jn ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM gen_jn  WHERE 1=1 " & Gen_jn & " ")
            '==============================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'   ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '===========================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM gen_jn  WHERE 1=1 " & LopGen_jn & " GROUP BY Company,Certify ")
            If dtTemp.Rows.Count <> 0 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Open_Amt + (" & CDbl(DbHelper.GetStr(row("amount_dr"))) & "-" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE Certify=N'" & DbHelper.GetStr(row("Certify")) & "' AND Company=N'" & DbHelper.GetStr(row("Company")) & "'  ")
                    Next
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_Gen_jn")
            If dtTemp.Rows.Count > 1 Then
                For Each row As DataRow In dtTemp.Rows
                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Rem_Amt= +( " & CDbl(DbHelper.GetStr(row("Open_Amt"))) & "+" & CDbl(DbHelper.GetStr(row("amount_dr"))) & " -" & CDbl(DbHelper.GetStr(row("amount_cr"))) & " ) WHERE ac_code='" & DbHelper.GetStr(row("ac_code")) & "'  ")

                    DbHelper.ExecuteNonQuery("UPDATE Ap_Sum_Gen_jn SET Open_Amt=Rem_Amt  ")
                    Next
                End While
            Else
                Exit Sub
            End If
        End If

        '================================ReportGen_jn==================================

        'DbHelper.ExecuteNonQuery(" DELETE FROM Ap_Sum_Gen_jn ")
        'DbHelper.ExecuteNonQuery("INSERT INTO Ap_Sum_Gen_jn ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
        '          " SELECT date_work,code_dr,code_cr,ac_code,0,0,Sum(amount_dr-amount_cr),0,0,0,Last_User,Last_Update,Company " & _
        '          " FROM Open_jn   WHERE 1=1 " & Open_jn & " GROUP BY date_work,code_dr,code_cr,ac_code,Last_User,Last_Update,Company  " & _
        '          " UNION ALL " & _
        '          "SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
        '          " FROM gen_jn  WHERE 1=1 " & Gen_jn & " ")
    End Sub




    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Label5.Visible = True
        ClickMouseRadio2()
        Call Office()
        MuLngRpt = RptSjOff

        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & RptName & "' As    RptSjUd ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
        LngId = "7002" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDP ,"
        LngId = "7003" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
        LngId = "7004" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Descrip ,"
        LngId = "7005" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_No ,"
        LngId = "7006" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Ac_Code ,"
        LngId = "7007" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Certify ,"
        LngId = "7008" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Dr ,"
        LngId = "7009" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Cr ,"
        LngId = "7010" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Amt_LAK ,"
        LngId = "7011" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Sign ,"
        LngId = "7012" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_SignNme ,"
        LngId = "7026" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InDate ,"
        LngId = "7028" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_AmtCurr ,"
        LngId = "7029" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Curr ,"
        LngId = "7030" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Rate ,"
        LngId = "7031" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_User ,"

        Panel4.Visible = False
        If FG.get_TextMatrix(1, 1) = "" Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
        'SLF = "SELECT " & MuLngRpt & "   *  ,  Acc_Code.Name_L  As AcNmeEx_L , Acc_Code.Name_E  As AcNmeEx_E FROM  gen_jn       INNER JOIN Acc_Code ON gen_jn.ac_code = Acc_Code.Ac_Code WHERE Book <>'' "
        SLF = "SELECT  " & mformat & "  as mformat  , " & MuLngRpt & "   *  ,    gen_jn.last_user,  gen_jn.ac_name  As AcNmeEx_L , gen_jn.ac_namee  As AcNmeEx_E FROM  gen_jn  WHERE Book <>'' "


        Call LoadLoGO()

        If CheckBox1.Checked = False Then
            Dim dt As DataTable
            Try
                dt = DbHelper.GetDataTable("  " & SLF & "  " & SQL & "order by  gen_jn.cnt ")
                If dt.Rows.Count = 0 Then 
                    MsgBox("ບໍ່ມີຂໍ້ມູນ") : Label5.Visible = False : Exit Sub
                End If
            Catch ex As Exception
                VSysError = True
                MessageBox.Show("Database Error: " & ex.Message, "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Label5.Visible = False
                Exit Sub
            End Try
            Dim FrmPreview As New FmPreview : FrmClosing()
            'Dim Rpt As New CryGeneralLedgers
            Dim Rpt As New CrystalReport_General_Jurnal_Curr_List_P
            If MdShowLOGO = 1 Then
                Rpt.Subreports(0).SetDataSource(RsLOGO)
            End If

            Rpt.SetDataSource(dt)
            FrmPreview.ReportViewer.ReportSource = Rpt
            FrmPreview.ReportViewer.DisplayGroupTree = False
            FrmPreview.MdiParent = FmMain
            FrmPreview.WindowState = FormWindowState.Maximized
            Label5.Visible = False
            FrmPreview.Show()
            FrmPreview.Focus()
        Else
            '========
            Dim dt As DataTable
            Try
                dt = DbHelper.GetDataTable("  " & SLF & "  " & SQL & "    " & "" & " ")
                If dt.Rows.Count = 0 Then 
                    MsgBox("ບໍ່ມີຂໍ້ມູນ") : Label5.Visible = False : Exit Sub
                End If
            Catch ex As Exception
                VSysError = True
                MessageBox.Show("Database Error: " & ex.Message, "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Label5.Visible = False
                Exit Sub
            End Try
            Dim FrmPreview As New FmPreview : FrmClosing()
            'Dim Rpt As New CryGeneralLedgersUser
            Dim Rpt As New CrystalReport_General_Jurnal_Curr_List_P_Curr
            If MdShowLOGO = 1 Then
                Rpt.Subreports(0).SetDataSource(RsLOGO)
            End If

        Rpt.SetDataSource(dtReport)
            FrmPreview.ReportViewer.ReportSource = Rpt
            FrmPreview.ReportViewer.DisplayGroupTree = False
            FrmPreview.MdiParent = FmMain
            FrmPreview.WindowState = FormWindowState.Maximized
            Label5.Visible = False
            FrmPreview.Show()
            FrmPreview.Focus()
        End If
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub RadioButton1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton1.MouseEnter
        RadioButton1.Size = New System.Drawing.Size(65, 30)
        Panel3.Size = New System.Drawing.Size(1261, 35)
        RadioButton1.BackColor = Color.Aquamarine
        'Panel3.BackColor = Color.Aquamarine
    End Sub

    Private Sub RadioButton1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton1.MouseLeave
        RadioButton1.Size = New System.Drawing.Size(60, 25)
        Panel3.Size = New System.Drawing.Size(1261, 29)
        RadioButton1.BackColor = Color.Gainsboro

    End Sub


    Private Sub RadioButton2_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton2.MouseEnter
        RadioButton2.Size = New System.Drawing.Size(65, 30)
        Panel3.Size = New System.Drawing.Size(1261, 35)
        RadioButton2.BackColor = Color.Aquamarine
    End Sub

    Private Sub RadioButton2_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton2.MouseLeave
        RadioButton2.Size = New System.Drawing.Size(60, 25)
        Panel3.Size = New System.Drawing.Size(1261, 29)
        RadioButton2.BackColor = Color.Gainsboro

    End Sub

    Private Sub RadioButton3_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton3.MouseEnter
        RadioButton3.Size = New System.Drawing.Size(65, 30)
        Panel3.Size = New System.Drawing.Size(1261, 35)
        RadioButton3.BackColor = Color.Aquamarine
    End Sub

    Private Sub RadioButton3_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton3.MouseLeave
        RadioButton3.Size = New System.Drawing.Size(60, 25)
        Panel3.Size = New System.Drawing.Size(1261, 29)
        RadioButton3.BackColor = Color.Gainsboro
        ColorRadioButton()
    End Sub

    Private Sub RadioButton4_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton4.MouseEnter
        RadioButton4.Size = New System.Drawing.Size(65, 30)
        Panel3.Size = New System.Drawing.Size(1261, 35)
        RadioButton4.BackColor = Color.Aquamarine
    End Sub

    Private Sub RadioButton4_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton4.MouseLeave
        RadioButton4.Size = New System.Drawing.Size(60, 25)
        Panel3.Size = New System.Drawing.Size(1261, 29)
        RadioButton4.BackColor = Color.Gainsboro
        ColorRadioButton()
    End Sub

    Private Sub RadioButton5_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton5.MouseEnter
        RadioButton5.Size = New System.Drawing.Size(65, 30)
        Panel3.Size = New System.Drawing.Size(1261, 35)
        RadioButton5.BackColor = Color.Aquamarine

    End Sub

    Private Sub RadioButton5_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton5.MouseLeave
        RadioButton5.Size = New System.Drawing.Size(60, 25)
        Panel3.Size = New System.Drawing.Size(1261, 29)
        RadioButton5.BackColor = Color.Gainsboro
        ColorRadioButton()
    End Sub

    Private Sub RadioButton6_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton6.MouseEnter
        RadioButton6.Size = New System.Drawing.Size(65, 30)
        Panel3.Size = New System.Drawing.Size(1261, 35)
        RadioButton6.BackColor = Color.Aquamarine

    End Sub

    Private Sub RadioButton6_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton6.MouseLeave
        RadioButton6.Size = New System.Drawing.Size(60, 25)
        Panel3.Size = New System.Drawing.Size(1261, 29)
        RadioButton6.BackColor = Color.Gainsboro
        ColorRadioButton()
    End Sub
    '7
    Private Sub RadioButton7_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton7.MouseEnter
        RadioButton7.Size = New System.Drawing.Size(65, 30)
        Panel3.Size = New System.Drawing.Size(1261, 35)
        RadioButton7.BackColor = Color.Aquamarine
        'Dim s As String

        's = Panel5.Location(x)
    End Sub

    Private Sub RadioButton7_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton7.MouseLeave
        RadioButton7.Size = New System.Drawing.Size(60, 25)
        Panel3.Size = New System.Drawing.Size(1261, 29)
        RadioButton7.BackColor = Color.Gainsboro
        ColorRadioButton()
    End Sub

    '1261
    '8
    Private Sub RadioButton8_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton8.MouseEnter
        RadioButton8.Size = New System.Drawing.Size(65, 30)
        Panel3.Size = New System.Drawing.Size(1261, 35)
        RadioButton8.BackColor = Color.Aquamarine

    End Sub

    Private Sub RadioButton8_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton8.MouseLeave
        RadioButton8.Size = New System.Drawing.Size(60, 25)
        Panel3.Size = New System.Drawing.Size(1261, 29)
        RadioButton8.BackColor = Color.Gainsboro
        ColorRadioButton()
    End Sub
    '9
    Private Sub RadioButton9_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton9.MouseEnter
        RadioButton9.Size = New System.Drawing.Size(65, 30)
        Panel3.Size = New System.Drawing.Size(1261, 35)
        RadioButton9.BackColor = Color.Aquamarine

    End Sub

    Private Sub RadioButton9_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton9.MouseLeave
        RadioButton9.Size = New System.Drawing.Size(60, 25)
        Panel3.Size = New System.Drawing.Size(1261, 29)
        RadioButton9.BackColor = Color.Gainsboro
        ColorRadioButton()
    End Sub
    '10
    Private Sub RadioButton10_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton10.MouseEnter
        RadioButton10.Size = New System.Drawing.Size(65, 30)
        Panel3.Size = New System.Drawing.Size(1261, 35)
        RadioButton10.BackColor = Color.Aquamarine

    End Sub

    Private Sub RadioButton10_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton10.MouseLeave
        RadioButton10.Size = New System.Drawing.Size(60, 25)
        Panel3.Size = New System.Drawing.Size(1261, 29)
        RadioButton10.BackColor = Color.Gainsboro
        ColorRadioButton()
    End Sub


    '11
    Private Sub RadioButton11_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton11.MouseEnter
        RadioButton11.Size = New System.Drawing.Size(65, 30)
        Panel3.Size = New System.Drawing.Size(1261, 35)
        RadioButton11.BackColor = Color.Aquamarine

    End Sub

    Private Sub RadioButton11_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton11.MouseLeave
        RadioButton11.Size = New System.Drawing.Size(60, 25)
        Panel3.Size = New System.Drawing.Size(1261, 29)
        RadioButton11.BackColor = Color.Gainsboro
        ColorRadioButton()
    End Sub
    '12
    Private Sub RadioButton12_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton12.MouseEnter
        RadioButton12.Size = New System.Drawing.Size(65, 30)
        Panel3.Size = New System.Drawing.Size(1261, 35)
        RadioButton12.BackColor = Color.Aquamarine

    End Sub

    Private Sub RadioButton12_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton12.MouseLeave
        RadioButton12.Size = New System.Drawing.Size(60, 25)
        Panel3.Size = New System.Drawing.Size(1261, 29)
        RadioButton12.BackColor = Color.Gainsboro
        ColorRadioButton()
    End Sub
    Private Sub RadioButton13_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton13.MouseEnter
        RadioButton13.Size = New System.Drawing.Size(65, 30)
        Panel3.Size = New System.Drawing.Size(1261, 35)
        RadioButton13.BackColor = Color.Aquamarine

    End Sub

    Private Sub RadioButton13_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton13.MouseLeave
        RadioButton13.Size = New System.Drawing.Size(60, 25)
        Panel3.Size = New System.Drawing.Size(1261, 29)
        RadioButton13.BackColor = Color.Gainsboro
        ColorRadioButton()
    End Sub

    Private Sub RadioButton14_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton14.MouseEnter
        RadioButton14.Size = New System.Drawing.Size(65, 30)
        Panel3.Size = New System.Drawing.Size(1261, 35)
        RadioButton14.BackColor = Color.Aquamarine

    End Sub

    Private Sub RadioButton14_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton14.MouseLeave
        RadioButton14.Size = New System.Drawing.Size(60, 25)
        Panel3.Size = New System.Drawing.Size(1261, 29)
        RadioButton14.BackColor = Color.Gainsboro
        ColorRadioButton()
        LoadMonthSQL()
    End Sub

    Private Sub BtnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRefresh.Click
        Panel4.Visible = False
        LoadMonthSQL()
    End Sub

    Private Sub RCurr_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RCurr.CheckedChanged
        ComboBox1.Visible = True
        Nme.Enabled = False
        Nme.Visible = False
        LoadsLurr()
        ComboBox1.Focus()
    End Sub

    Private Sub RBook_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBook.CheckedChanged
        Nme.Enabled = False
        Nme.Visible = False
        ComboBox1.Visible = True

        LoadBooks()
        ComboBox1.Focus()
    End Sub

    Private Sub RAcType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RAcType.CheckedChanged
        Nme.Enabled = False
        Nme.Visible = False
        ComboBox1.Visible = True
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("==ທັງຫມົດ==")
        ComboBox1.Items.Add("ບັນຊີຍ່ອຍ (D)")
        ComboBox1.Items.Add("ບັນຊີແມ່ (P)")
        ComboBox1.SelectedIndex = 0
        ComboBox1.Focus()
    End Sub

    Private Sub RAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RAll.CheckedChanged

        ComboBox1.Items.Clear()
        ComboBox1.Visible = True
        Nme.Enabled = False
        Nme.Visible = True
    End Sub

    Private Sub RDesc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RDesc.CheckedChanged
        Nme.Enabled = True
        Nme.Visible = True
        ComboBox1.Visible = False

        Nme.Focus()
    End Sub

    Private Sub RAcNme_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RAcNme.CheckedChanged
        Nme.Enabled = True
        Nme.Visible = True
        ComboBox1.Visible = False
        Nme.Focus()
    End Sub

    Private Sub RAc_code_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RAc_code.CheckedChanged
        Nme.Enabled = True
        Nme.Visible = True
        ComboBox1.Visible = False
        Nme.Focus()
    End Sub

    Private Sub RCex_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RCex.CheckedChanged
        Nme.Enabled = True
        Nme.Visible = True
        ComboBox1.Visible = False
        Nme.Focus()
    End Sub

    Private Sub Rinvioce_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Nme.Enabled = True
        Nme.Visible = True
        ComboBox1.Visible = False
        Nme.Focus()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Nme.Text = ComboBox1.Text
    End Sub

    Private Sub Nme_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Nme.KeyPress
        If e.KeyChar = Chr(13) Then
            LoadMonthSQL()
        End If
    End Sub

    Private Sub TextBox20_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox20.TextChanged

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDescrip.TextChanged

    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TotalDr.TextChanged

    End Sub

    Private Sub TextBox13_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TotalCr.TextChanged

    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Balance.TextChanged

    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub Label11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label11.Click

    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ac_Name.TextChanged

    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Panel4.Visible = False
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click

        ' ''If IsNumeric(Microsoft.VisualBasic.Right(txtNewId.Text, 7)) = False Then MsgBox("7 ໂຕທາງທ້າຍຕ້ອງເປັນໂຕເລກເທົານັ້ນ") : txtNewId.BackColor = Color.Red : txtNewId.Focus() : Exit Sub
        ''Dim srNum As New ADODB.Recordset
        ''Dim mNum As Integer = 0
        ' ''If IsNumeric(Microsoft.VisualBasic.Right(txtNewId.Text, 7)) = False Then MsgBox("7 ໂຕທາງທ້າຍຕ້ອງເປັນໂຕເລກເທົານັ້ນ") : txtNewId.BackColor = Color.Red : txtNewId.Focus() : Exit Sub
        ''Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT top 1 Right(certify,7) As  certify   FROM  gen_jn where  book ='" & CmbBook2.Text & "' And  year(date_work)='" & Format(CDate(GetGridValue(FG, FG.CurrentRow.Index, 1)), "yyyy") & "'   Order by  Right(certify,7) DESC", srNum)
        ''If srNum.RecordCount = 0 Then
        ''    mNum = 0
        ''Else
        ''    mNum = Val(srNum.Fields("certify").Value.ToString)
        ''End If
        ''mNum = mNum + 1

        ''If Int(Microsoft.VisualBasic.Right(txtNewId.Text, 7)) > mNum Then

        ''    If Len(CStr(mNum)) = 1 Then
        ''        MsgBox("ເລກທ້າຍ 7 ໂຕບໍ່ໃຫ້ເກີນເລກ " & "000000" & mNum)
        ''    ElseIf Len(CStr(mNum)) = 2 Then
        ''        MsgBox("ເລກທ້າຍ 7 ໂຕບໍ່ໃຫ້ເກີນເລກ " & "00000" & mNum)
        ''    ElseIf Len(CStr(mNum)) = 3 Then
        ''        MsgBox("ເລກທ້າຍ 7 ໂຕບໍ່ໃຫ້ເກີນເລກ " & "0000" & mNum)
        ''    ElseIf Len(CStr(mNum)) = 4 Then

        ''        MsgBox("ເລກທ້າຍ 7 ໂຕບໍ່ໃຫ້ເກີນເລກ " & "000" & mNum)
        ''    ElseIf Len(CStr(mNum)) = 5 Then
        ''        MsgBox("ເລກທ້າຍ 7 ໂຕບໍ່ໃຫ້ເກີນເລກ " & "00" & mNum)
        ''    ElseIf Len(CStr(mNum)) >= 6 Then
        ''        MsgBox("ເລກທ້າຍ 7 ໂຕບໍ່ໃຫ້ເກີນເລກ " & mNum)
        ''    End If

        ''    txtNewId.BackColor = Color.Red
        ''    txtNewId.Focus()
        ''    Exit Sub

        ''End If














        Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT AC_CODE FROM Gen_jn WHERE   book ='" & CmbBook2.Text & "' And  certify = N'" & txtNewId.Text & "' And  year(date_work)=" & Format(CDate(GetGridValue(FG, FG.CurrentRow.Index, 1)), "yyyy") & " ")
        If dtTemp.Rows.Count > 0 Then
            MsgBox("ເລກລະຫັດ : " & Trim(txtNewId.Text) & " ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
            txtNewId.Focus()
            If RSC.State = ConnectionState.Open Then RSC.Close()
            Exit Sub

        End If







        'Dim dtTemp As DataTable = DbHelper.GetDataTable("select *  from Gen_jn WHERE cnt<>''  " & SQL & "order by certify")
        ''Call LoadData("SELECT * FROM  Gen_jn")
        'If dtTemp.Rows.Count > 0 Then

        'End If


        If MessageBox.Show("ທ່ານຕ້ອງການປ່ຽນລະຫັດ " & txtOldId.Text & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            DbHelper.ExecuteNonQuery("UPDATE gen_jn SET book='" & CmbBook2.Text & "',certify= N'" & txtNewId.Text & "' WHERE   book =N'" & FG.get_TextMatrix(FG.CurrentRow.Index, 16) & "' And  certify  =N'" & MDInvoiceNo & "'   And  year(date_work)='" & Format(CDate(GetGridValue(FG, FG.CurrentRow.Index, 1)), "yyyy") & "' ")

        End If

        Panel4.Visible = False
        LoadMonthSQL()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Panel4.Visible = False
        If MDInvoiceNo = "" Then
            MessageBox.Show("ກະລຸນນາເລືອກກ່ອນ")
            Exit Sub
        End If

        Panel4.Visible = True





        CmbBook2.Items.Clear()
        Dim dtBooks As DataTable = DbHelper.GetDataTable("SELECT * FROM books WHERE bookid <> ''")
        If dtBooks.Rows.Count <> 0 Then
            For Each row As DataRow In dtBooks.Rows
                CmbBook2.Items.Add(Trim(DbHelper.GetStr(row("bookid"))))
            Next
        End If

        CmbBook2.Text = "GL"
        Dim dtTemp As DataTable = DbHelper.GetDataTable("select book , certify from gen_jn WHERE   book ='" & FG.get_TextMatrix(FG.CurrentRow.Index, 16) & "' And  certify  = '" & MDInvoiceNo & "'   And  year(date_work)='" & Format(CDate(GetGridValue(FG, FG.CurrentRow.Index, 1)), "yyyy") & "' order by cnt")

        If dtTemp.Rows.Count > 0 Then
            txtOldId.Text = DbHelper.GetStr(dtTemp.Rows(0)("certify"))
            Books.Text = DbHelper.GetStr(dtTemp.Rows(0)("book"))
        End If
        CmbBook2.Text = Books.Text
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click




        Panel4.Visible = False
        If MDInvoiceNo = "" Then
            MessageBox.Show("ກະລຸນນາເລືອກກ່ອນ")
            Exit Sub
        End If

        ClickMouseRadio2()
        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & " 'As Crl_Lng  ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
        LngId = "7002" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDP ,"
        LngId = "7003" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
        LngId = "7004" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Descrip ,"
        LngId = "7005" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_No ,"
        LngId = "7006" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Ac_Code ,"
        LngId = "7007" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Certify ,"
        LngId = "7008" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Dr ,"
        LngId = "7009" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Cr ,"
        LngId = "7010" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Amt_LAK ,"
        LngId = "7011" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Sign ,"
        LngId = "7012" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_SignNme ,"
        LngId = "7026" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InDate ,"
        LngId = "7028" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_AmtCurr ,"
        LngId = "7029" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Curr ,"
        LngId = "7030" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Rate ,"
        LngId = "7031" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_User ,"
        LngId = "7032" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Amount ,"
        LngId = "7033" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Number ,"

        LngId = "7039" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Amount_Total	 ,"
        LngId = "7069" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_In_Word	 ,"
        'LngId = "7038" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As  Crl_Ac_Name	 ,"
        LngId = "5019" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As  Crl_Ac_Name	 ,"
        LngId = "7122" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As  Refno ,"
        LngId = "7123" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As  Amount ,"
        LngId = "7124" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As  Fore  ,"
        LngId = "7125" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As  Exchange ,"
        LngId = "7126" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As  Amount_LAK ,"

        If MuLng = "L" Then
            MuLngRpt = MuLngRpt & "N'" & Amount_In_Word & "' As Crl_Amt_In_Word	 ,"
        Else
            MuLngRpt = MuLngRpt & "N'" & Amount_In_Word & "' As Crl_Amt_In_Word	 ,"
        End If


        'SLF = MuLngRpt & " gen_jn.company ,gen_jn.Date_Work , gen_jn.certify, gen_jn.ac_code, gen_jn.descrip , gen_jn.descripe , gen_jn.amt_dr, gen_jn.amt_cr, Acc_Code.Name_L AS Name_L , Acc_Code.Name_E AS Name_E  "
        SLF = MuLngRpt & "   gen_jn.last_user,  gen_jn.Rate_USD, gen_jn.company ,gen_jn.Date_Work , gen_jn.Curr, gen_jn.certify, gen_jn.referno, gen_jn.ac_code, gen_jn.descrip , gen_jn.descripe , gen_jn.amount_dr, gen_jn.amount_cr, gen_jn.amt_dr, gen_jn.amt_cr, gen_jn.ac_name  AS Name_L , gen_jn.ac_namee AS Name_E  "
        'SLF = MuLngRpt & " gen_jn.company ,gen_jn.Date_Work , gen_jn.certify, gen_jn.ac_code, gen_jn.descrip , gen_jn.descripe , gen_jn.amt_dr, gen_jn.amt_cr, Acc_Code.Name_L AS Name_L , Acc_Code.Name_E AS Name_E  "
        Call LoadLoGO()
        Dim dtReport As DataTable = DbHelper.GetDataTable("SELECT   " & SLF & "   FROM gen_jn INNER JOIN Acc_Code ON gen_jn.ac_code = Acc_Code.Ac_Code WHERE gen_jn.certify = N'" & MDInvoiceNo & "' And  year(date_work)=" & Format(CDate(GetGridValue(FG, FG.CurrentRow.Index, 1)), "yyyy") & "  order by gen_jn.cnt ASC")
        If dtReport.Rows.Count = 0 Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim Rpt As New CryNewsJerneralJournal
        If MdShowLOGO = 1 Then
            Rpt.Subreports(0).SetDataSource(RsLOGO)
        End If

        Rpt.SetDataSource(Rs)
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.DisplayGroupTree = False
        'FrmPreview.MdiParent = FmMain
        FrmPreview.WindowState = FormWindowState.Maximized
        FrmPreview.ShowDialog()
        FrmPreview.Focus()
    End Sub

    Private Sub txtNewId_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNewId.KeyPress
        'If e.KeyChar = Chr(13) Then
        '    'If IsNumeric(txtNewId.Text) = True Then
        '    If Len(CStr(txtNewId.Text).Trim) = 1 Then
        '        txtNewId.Text = CmbBook2.Text & "000000" & CStr(txtNewId.Text)
        '    ElseIf Len(CStr(txtNewId.Text).Trim) = 2 Then
        '        txtNewId.Text = CmbBook2.Text & "00000" & CStr(txtNewId.Text)
        '    ElseIf Len(CStr(txtNewId.Text).Trim) = 3 Then
        '        txtNewId.Text = CmbBook2.Text & "0000" & CStr(txtNewId.Text)
        '    ElseIf Len(CStr(txtNewId.Text).Trim) = 4 Then
        '        txtNewId.Text = CmbBook2.Text & "000" & CStr(txtNewId.Text)
        '    ElseIf Len(CStr(txtNewId.Text).Trim) = 5 Then
        '        txtNewId.Text = CmbBook2.Text & "00" & CStr(txtNewId.Text)
        '    ElseIf Len(CStr(txtNewId.Text).Trim) >= 6 Then
        '        txtNewId.Text = CmbBook2.Text & Microsoft.VisualBasic.Right(txtNewId.Text, 7)
        '    End If
        '    txtNewId.SelectAll()
        'End If
    End Sub

    Private Sub txtNewId_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNewId.LostFocus
        ''If IsNumeric(txtNewId.Text) = True Then
        'If Len(CStr(txtNewId.Text).Trim) = 1 Then
        '    txtNewId.Text = CmbBook2.Text & "000000" & CStr(txtNewId.Text)
        'ElseIf Len(CStr(txtNewId.Text).Trim) = 2 Then
        '    txtNewId.Text = CmbBook2.Text & "00000" & CStr(txtNewId.Text)
        'ElseIf Len(CStr(txtNewId.Text).Trim) = 3 Then
        '    txtNewId.Text = CmbBook2.Text & "0000" & CStr(txtNewId.Text)
        'ElseIf Len(CStr(txtNewId.Text).Trim) = 4 Then
        '    txtNewId.Text = CmbBook2.Text & "000" & CStr(txtNewId.Text)
        'ElseIf Len(CStr(txtNewId.Text).Trim) = 5 Then
        '    txtNewId.Text = CmbBook2.Text & "00" & CStr(txtNewId.Text)
        'ElseIf Len(CStr(txtNewId.Text).Trim) >= 6 Then
        '    txtNewId.Text = CmbBook2.Text & Microsoft.VisualBasic.Right(txtNewId.Text, 7)

        'End If
    End Sub

    Private Sub txtNewId_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNewId.TextChanged
        txtNewId.BackColor = Color.White
    End Sub

    Private Sub CmbBook2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbBook2.SelectedIndexChanged
        If Len(CStr(txtNewId.Text).Trim) = 1 Then
            txtNewId.Text = CmbBook2.Text & "000000" & CStr(txtNewId.Text)
        ElseIf Len(CStr(txtNewId.Text).Trim) = 2 Then
            txtNewId.Text = CmbBook2.Text & "00000" & CStr(txtNewId.Text)
        ElseIf Len(CStr(txtNewId.Text).Trim) = 3 Then
            txtNewId.Text = CmbBook2.Text & "0000" & CStr(txtNewId.Text)
        ElseIf Len(CStr(txtNewId.Text).Trim) = 4 Then
            txtNewId.Text = CmbBook2.Text & "000" & CStr(txtNewId.Text)
        ElseIf Len(CStr(txtNewId.Text).Trim) = 5 Then
            txtNewId.Text = CmbBook2.Text & "00" & CStr(txtNewId.Text)
        ElseIf Len(CStr(txtNewId.Text).Trim) >= 6 Then
            txtNewId.Text = CmbBook2.Text & Microsoft.VisualBasic.Right(txtNewId.Text, 7)

        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        CMS2.Enabled = True
        CMS3.Enabled = True
        'MsgBox(MDInvoiceNo)
        'Panel3.Size = New System.Drawing.Size(1261, 29)
        If Panel5.Visible = False Then
            Panel5.Visible = True
        Else
            Panel5.Visible = False
        End If

        'Call StartLoadDataList()
    End Sub
    Public Sub loadColor()
        If CheckBox2.Checked = True Then
            Dim J As Integer
            FG.Redraw = False

            For J = 1 To FG.CurrentRow.Indexs - 1
                FG.CurrentRow.Index = J
                If Trim(FG.get_TextMatrix(J, 4)) <> "" Then
                    If Trim(FG.get_TextMatrix(J, 10)) <> 0 Then
                        FG.CurrentCell.ColumnIndex = 7
                        FG.CellFontBold = True
                        FG.CurrentCell.ColumnIndex = 10
                        FG.CellFontBold = True
                    Else
                        FG.CurrentCell.ColumnIndex = 8
                        FG.CellFontBold = True
                        FG.CurrentCell.ColumnIndex = 11
                        FG.CellFontBold = True
                    End If
                    FG.CurrentCell.ColumnIndex = 4
                    FG.CellFontBold = True
                End If
                'MsgBox(Trim(FG.get_TextMatrix(J, 14)))
                Dim C1 As String = "255, 192, 128"
                Dim C2 As Color = Color.Red
                If FG.get_TextMatrix(J, 14) = 1 Then
                    FG.CurrentCell.ColumnIndex = 1
                    FG.CellForeColor = C2
                    'FG.CellForeColor = Color.FromArgb(C1)
                    FG.CurrentCell.ColumnIndex = 2
                    FG.CellForeColor = C2
                    FG.CurrentCell.ColumnIndex = 3
                    FG.CellForeColor = C2
                    FG.CurrentCell.ColumnIndex = 4
                    FG.CellForeColor = C2
                    FG.CurrentCell.ColumnIndex = 6
                    FG.CellForeColor = C2
                    FG.CurrentCell.ColumnIndex = 7
                    FG.CellForeColor = C2
                    FG.CurrentCell.ColumnIndex = 8
                    FG.CellForeColor = C2
                    FG.CurrentCell.ColumnIndex = 9
                    FG.CellForeColor = C2
                    FG.CurrentCell.ColumnIndex = 10

                    FG.CellForeColor = C2
                    FG.CurrentCell.ColumnIndex = 11
                    FG.CellForeColor = C2
                    FG.CurrentCell.ColumnIndex = 12
                    FG.CellForeColor = C2
                End If

                If FG.get_TextMatrix(J, 14) = 2 Then
                    C2 = Color.Gray
                    FG.CurrentCell.ColumnIndex = 1
                    FG.CellForeColor = C2
                    'FG.CellForeColor = Color.FromArgb(C1)
                    FG.CurrentCell.ColumnIndex = 2
                    FG.CellForeColor = C2
                    FG.CurrentCell.ColumnIndex = 3
                    FG.CellForeColor = C2
                    FG.CurrentCell.ColumnIndex = 4
                    FG.CellForeColor = C2
                    FG.CurrentCell.ColumnIndex = 6
                    FG.CellForeColor = C2
                    FG.CurrentCell.ColumnIndex = 7
                    FG.CellForeColor = C2
                    FG.CurrentCell.ColumnIndex = 8
                    FG.CellForeColor = C2
                    FG.CurrentCell.ColumnIndex = 9
                    FG.CellForeColor = C2
                    FG.CurrentCell.ColumnIndex = 10

                    FG.CellForeColor = C2
                    FG.CurrentCell.ColumnIndex = 11
                    FG.CellForeColor = C2
                    FG.CurrentCell.ColumnIndex = 12
                    FG.CellForeColor = C2
                End If



            Next J

            FG.Redraw = True
        End If
    End Sub

    Public Sub StartLoadDataList()
        'MsgBox("00")
        LoadSQLCheckbox()
        ClickMouseRadio()
        SQL = ""
        Panel4.Visible = False
        Call LoadDividePage()
        P = 1
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page
        If FG.get_TextMatrix(1, 1) = "" Then lblpage_total.Text = "0/0"
        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        If CmbPage.Items.Count > 0 Then
            CmbPage.SelectedIndex = 0
        End If
        If FG.get_TextMatrix(1, 1) = "" Then CmbPage.Text = ""
        LbPage.Focus()
    End Sub
    Public Sub PageCnt(ByVal StrSQL As String, ByVal ConStr As String, ByVal PageNum As Long, ByVal RowPerPage As Integer)
        Load_DES()
        Label5.Visible = True
        Label5.BringToFront()
        x0 = 0
        y0 = 0
        'Dim RsLoad As New ADODB.Recordset
        'Dim rssum As New ADODB.Recordset
        Dim i As Integer
        FG.CurrentRow.Index = 1
        Dim x As String
        PageNum = PageNum - 1
        Dim MS As String = "And Company = '" & MuSubOff & "'"
        Dim OfUsr1 As String = Microsoft.VisualBasic.Left(Off_Usr.Text, 5)
        Dim OfUsr2 As String = Mid(Off_Usr.Text, 4, 2)
        Dim OfUsr3 As String = Microsoft.VisualBasic.Left(Off_Usr.Text, 2)
        If OfUsr1 = "00-00" Then
            MULook2 = ""
        Else
            If OfUsr2 = "00" Then
                MULook2 = "  And  Left(company,2)= '" & OfUsr3 & "' "
            Else
                MULook2 = "  And company= '" & OfUsr1 & "'   "
            End If
        End If

        Dim API As String
        If CheckBox4.Checked = True Then
            API = " and API='API' "
        Else
            API = ""
        End If

        DbHelper.ExecuteNonQuery("UPDATE gen_jn set company=office_ID where company is null ")
        DbHelper.ExecuteNonQuery("UPDATE gen_jn set lock=4 where lock is null ")

        DbHelper.ExecuteNonQuery("update gen_jn set Rate_USD=rate where curr='USD' and Rate_USD is null ")
        DbHelper.ExecuteNonQuery("update gen_jn set Rate_USD=0 where   Rate_USD is null ")

        DbHelper.ExecuteNonQuery("update Gen_jn set Gen_jn.Ac_namee=Acc_Code.Name_E from Gen_jn,Acc_Code where Acc_Code.ac_code=Gen_jn.ac_code and Gen_jn.Ac_namee is null")

        SQL = " AND gen_jn.date_work   BETWEEN '" & Format(dts.Value, "yyyy-MM-dd") & "' AND '" & Format(dtt.Value, "yyyy-MM-dd") & "'  " & SR & " " & MULook2 & " "

        x = " date_work , certify , referno , cheque_no , Ac_code,Ac_name, Ac_namee, code_dr , code_cr , descrip , descripe , amount_dr , amount_cr , amount ,  amt_dr , amt_cr  ,  curr , cnt , lock  , book, Company "
        Dim L As String = "select " & x & "  from Gen_jn WHERE certify<>''  " & SQL & " " & API & " order by " & CntNB & ""
        BtnSearch.Visible = False
        Dim dtTemp As DataTable = DbHelper.GetDataTable(L)

        If dtTemp.Rows.Count <> 0 Then
            Dim startIndex As Integer = RowPerPage * PageNum
            LbPage.Text = dtTemp.Rows.Count
            If dtTemp.Rows.Count Mod RowPerPage = 0 Then
                Last_page = dtTemp.Rows.Count / DividePage
            Else
                Last_page = dtTemp.Rows.Count / DividePage + 1
                If P = Last_page Then RowPerPage = dtTemp.Rows.Count Mod RowPerPage
            End If
            FG.Redraw = False
            FG.CurrentRow.Index = 1
            px = 0
            CMS2.Enabled = False
            CMS3.Enabled = False
            For i = 0 To RowPerPage - 1
                Dim rowIndex = startIndex + i
                If rowIndex >= dtTemp.Rows.Count Then Exit For
                Dim row = dtTemp.Rows(rowIndex)
                Dim s As String
                If MuLng = "L" Then s = DbHelper.GetStr(row("descrip")) Else s = DbHelper.GetStr(row("descripe"))
                FG.AddItem((rowIndex + 1) & vbTab & Format(CDate(DbHelper.GetStr(row("date_work"))), "dd/MM/yyyy") & _
                            "" & vbTab & DbHelper.GetStr(row("certify")) & _
                             "" & vbTab & DbHelper.GetStr(row("referno")) & _
                            "" & vbTab & DbHelper.GetStr(row("Ac_code")) & _
                           "" & vbTab & Format(CDbl(DbHelper.GetStr(row("amount"))), "##,##0.00") & _
                             "" & vbTab & s & _
                             "" & vbTab & Format(CDbl(DbHelper.GetStr(row("amount_dr"))), "##,##0.00") & _
                              "" & vbTab & Format(CDbl(DbHelper.GetStr(row("amount_cr"))), "##,##0.00") & _
                                 "" & vbTab & DbHelper.GetStr(row("curr")) & _
                              "" & vbTab & Format(CDbl(DbHelper.GetStr(row("amt_dr"))), "##,##0.00") & _
                              "" & vbTab & Format(CDbl(DbHelper.GetStr(row("amt_cr"))), "##,##0.00") & _
                             "" & vbTab & DbHelper.GetStr(row("company")) & _
                                "" & vbTab & DbHelper.GetStr(row("cnt")) & _
                               "" & vbTab & DbHelper.GetStr(row("lock")) & _
                                "" & vbTab & "" & _
                                  "" & vbTab & DbHelper.GetStr(row("book")) & _
                            "" & vbTab & DbHelper.GetStr(row("referno")))
            Next i
            FG.CurrentRow.Index = FG.CurrentRow.Index - 1
            FG.Redraw = True
            lblpage_total.Text = P & "/" & Last_page
        Else
            FG.CurrentRow.Index = 1
            FG.CurrentRow.Index = 2
        End If



        If FG.get_TextMatrix(1, 1) <> "" Then
            FirstPage.Enabled = True
            BackPage.Enabled = True
            NextPage.Enabled = True
            LasthPage.Enabled = True
            EnterPage.Enabled = True
            LbPage.Text = FG.get_TextMatrix(1, 0) & " To " & FG.get_TextMatrix(FG.CurrentRow.Indexs - 1, 0) & ", Of " & LbPage.Text
            If P = 1 Then
                FirstPage.Enabled = False
                BackPage.Enabled = False
                NextPage.Enabled = True
                LasthPage.Enabled = True
            ElseIf P = Last_page Then
                FirstPage.Enabled = True
                BackPage.Enabled = True
                NextPage.Enabled = False
                LasthPage.Enabled = False
            End If
        Else
            FirstPage.Enabled = False
            BackPage.Enabled = False
            NextPage.Enabled = False
            LasthPage.Enabled = False
            EnterPage.Enabled = False
            Last_page = 0

            LbPage.Text = "0 To 0, Of 0"

        End If

        'MsgBox("kk")
        If NextPage.Enabled = False Then EnterPage.Text = "Back "
        If BackPage.Enabled = False Then EnterPage.Text = "Next  "
        'If P15.Checked = True Then
        Call loadColor()
        'End If

        Ch = 0
        Call SumAmount()
        Label5.Visible = False
    End Sub
    Private Sub LoadDividePage()
        LoadSQLCheckbox()
        ClickMouseRadio()
        SQL = ""
        Panel4.Visible = False
        'MULook2 = "" : If MuSubOff <> "00-00" Then MULook2 = "And gen_jn." & Mid(MULook, 5, CDbl(Len(MULook)) - 4) Else 
        SQL = " AND gen_jn.date_work   BETWEEN '" & Format(dts.Value, "yyyy-MM-dd") & "' AND '" & Format(dtt.Value, "yyyy-MM-dd") & "'  " & SR & " " & MULook2 & " "
        txtSC15.Enabled = False
        P15.ForeColor = Color.Black
        p25.ForeColor = Color.Black
        p50.ForeColor = Color.Black
        p250.ForeColor = Color.Black
        p500.ForeColor = Color.Black
        p1000.ForeColor = Color.Black
        p100.ForeColor = Color.Black
        If P15.Checked = True Then
            DividePage = txtSC15.Text
            txtSC15.Enabled = True
            P15.ForeColor = Color.Red
            'FG.FormatString = "^ລ/ດ |< ວັນທີ     |< ໃບຍັງຢືນ   |<ແຊັກເລກທີ|< ເລກບັນຊີໜີ  |< ເລກບັນຊີມີ  |<ເນື້ອໃນລາຍການ                        | ຈຳນວນເງິນຈົດໜີ້    | ຈຳນວນເງິນຈົດມີ  |<ສະກຸນເງິນ|ຈຳນວນເງິນຈົດໜີ້(ກີບ) |ຈຳນວນເງິນຈົດໜີ້(ກີບ) |< ຕົ້ນທຶນ ||||"
        ElseIf p25.Checked = True Then
            DividePage = 25
            p25.ForeColor = Color.Red
        ElseIf p50.Checked = True Then
            DividePage = 50
            p50.ForeColor = Color.Red
        ElseIf p100.Checked = True Then
            DividePage = 100
            p100.ForeColor = Color.Red
        ElseIf p250.Checked = True Then
            DividePage = 250
            p250.ForeColor = Color.Red
        ElseIf p500.Checked = True Then
            DividePage = 500
            p500.ForeColor = Color.Red
        ElseIf p1000.Checked = True Then
            DividePage = 1000
            p1000.ForeColor = Color.Red
        End If
    End Sub


    Private Sub FirstPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FirstPage.Click
        Call LoadDividePage()
        P = 1
        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page
        If FG.get_TextMatrix(1, 1) = "" Then lblpage_total.Text = "0/0"
        If CmbPage.SelectedIndex >= 0 Then
            CmbPage.SelectedIndex = P - 1
        End If
        If FG.get_TextMatrix(1, 1) = "" Then CmbPage.Text = ""

        LbPage.Focus()
    End Sub

    Private Sub BackPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackPage.Click
        Call LoadDividePage()
        If P = 1 Then Exit Sub
        P = P - 1

        Call LoadSQL()

        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = P & "/" & Last_page
        If CmbPage.SelectedIndex >= 0 Then
            CmbPage.SelectedIndex = P - 1
        End If
        LbPage.Focus()
    End Sub

    Private Sub NextPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NextPage.Click
        Call LoadDividePage()
        If P >= Last_page Then Exit Sub
        P = P + 1
        Call LoadSQL()

        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = P & "/" & Last_page
        If CmbPage.SelectedIndex >= 0 Then
            CmbPage.SelectedIndex = P - 1
        End If
        LbPage.Focus()
    End Sub

    Private Sub LasthPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LasthPage.Click
        Call LoadDividePage()
        P = Last_page
        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = Last_page & "/" & Last_page
        If CmbPage.SelectedIndex >= 0 Then
            CmbPage.SelectedIndex = P - 1
        End If
        LbPage.Focus()
    End Sub

    Private Sub CmbPage_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles CmbPage.KeyPress
        If e.KeyChar = Chr(13) Then
            If CmbPage.Text <> "" Then
                If IsNumeric(CmbPage.Text) = False Then CmbPage.Text = "1" : Exit Sub
                If CDbl(CmbPage.Text) > CDbl(Last_page) Then CmbPage.Text = CDbl(Last_page)
                P = CDbl(CmbPage.Text)
                Call LoadDividePage()
                'If P >= Last_page Then Exit Sub

                Call LoadSQL()
                Call PageCnt(StrSQL, ConString, P, DividePage)
                Me.lblpage_total.Text = P & "/" & Last_page
            End If
        End If
    End Sub

    Private Sub CmbPage_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles CmbPage.MouseDoubleClick
        'MsgBox("MouseDoubleClick")
    End Sub

    Private Sub CmbPage_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles CmbPage.MouseDown
        'MsgBox("MouseDown")
        Ch = 1
    End Sub

    Private Sub CmbPage_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbPage.MouseEnter
        'MsgBox("MouseEnter")
        'Ch = 0
        'Ch = 1
    End Sub

    Private Sub CmbPage_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbPage.MouseHover
        'MsgBox("MouseHover")
    End Sub

    Private Sub CmbPage_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbPage.MouseLeave
        'Ch = 0
        'MsgBox(Ch)
        'MsgBox("MouseLeave")
    End Sub

    Private Sub CmbPage_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles CmbPage.MouseMove
        'MsgBox("MouseMove")
    End Sub

    Private Sub CmbPage_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles CmbPage.MouseUp
        'MsgBox("MouseUp")
        'Ch = 1
    End Sub

    Private Sub CmbPage_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles CmbPage.MouseWheel
        'MsgBox("MouseWheel")
    End Sub



    Private Sub CmbPage_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbPage.SelectedIndexChanged
        If IsNumeric(CmbPage.Text) = False Then CmbPage.Text = "0" : Exit Sub


        If Ch = 1 Then


            If CmbPage.Text <> "" Then
                If IsNumeric(CmbPage.Text) = False Then CmbPage.Text = "1" : Exit Sub
                If CDbl(CmbPage.Text) > CDbl(Last_page) Then CmbPage.Text = CDbl(Last_page)
                P = CDbl(CmbPage.Text)
                Call LoadDividePage()
                'If P >= Last_page Then Exit Sub

                Call LoadSQL()
                Call PageCnt(StrSQL, ConString, P, DividePage)
                Me.lblpage_total.Text = P & "/" & Last_page
            End If



        End If



        ''If txtAmount.Text = "" Then txtAmount.Text = "0" : Exit Sub


        'If CmbPage.Text <> "" Then
        '    P = CDbl(CmbPage.Text)
        '    Call LoadDividePage()
        '    'If P >= Last_page Then Exit Sub

        '    Call LoadSQL()
        '    Call PageCnt(StrSQL, ConString, P, DividePage)
        '    Me.lblpage_total.Text = P & "/" & Last_page
        '    LbPage.Focus()
        'End If
    End Sub

    Private Sub p25_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles p25.CheckedChanged

    End Sub

    Private Sub p25_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p25.MouseClick
        txtSC15.Enabled = False
        P15.ForeColor = Color.Black
        p25.ForeColor = Color.Black
        p50.ForeColor = Color.Black
        p250.ForeColor = Color.Black
        p500.ForeColor = Color.Black
        p1000.ForeColor = Color.Black
        If P15.Checked = True Then
            txtSC15.Enabled = True
            P15.ForeColor = Color.Red : txtSC15.Focus() : txtSC15.SelectAll()
        ElseIf p25.Checked = True Then
            p25.ForeColor = Color.Red
        ElseIf p50.Checked = True Then
            p50.ForeColor = Color.Red
        ElseIf p100.Checked = True Then
            p100.ForeColor = Color.Red
        ElseIf p250.Checked = True Then
            p250.ForeColor = Color.Red
        ElseIf p500.Checked = True Then
            p500.ForeColor = Color.Red
        ElseIf p1000.Checked = True Then
            p1000.ForeColor = Color.Red
        End If
        Call LoadDividePage()
        P = 1

        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page
        If FG.get_TextMatrix(1, 1) = "" Then lblpage_total.Text = "0/0"
        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        If CmbPage.Items.Count > 0 Then
            CmbPage.SelectedIndex = 0
        End If
        If FG.get_TextMatrix(1, 1) = "" Then CmbPage.Text = ""
        'CmbPage.SelectedIndex = 0
    End Sub

    Private Sub p50_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles p50.CheckedChanged

    End Sub

    Private Sub p50_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p50.MouseClick
        txtSC15.Enabled = False
        P15.ForeColor = Color.Black
        p25.ForeColor = Color.Black
        p50.ForeColor = Color.Black
        p250.ForeColor = Color.Black
        p500.ForeColor = Color.Black
        p1000.ForeColor = Color.Black
        If P15.Checked = True Then
            txtSC15.Enabled = True
            P15.ForeColor = Color.Red : txtSC15.Focus() : txtSC15.SelectAll()
        ElseIf p25.Checked = True Then
            p25.ForeColor = Color.Red
        ElseIf p50.Checked = True Then
            p50.ForeColor = Color.Red
        ElseIf p100.Checked = True Then
            p100.ForeColor = Color.Red
        ElseIf p250.Checked = True Then
            p250.ForeColor = Color.Red
        ElseIf p500.Checked = True Then
            p500.ForeColor = Color.Red
        ElseIf p1000.Checked = True Then
            p1000.ForeColor = Color.Red
        End If
        Call LoadDividePage()
        P = 1

        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page
        If FG.get_TextMatrix(1, 1) = "" Then lblpage_total.Text = "0/0"
        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        If CmbPage.Items.Count > 0 Then
            CmbPage.SelectedIndex = 0
        End If
        If FG.get_TextMatrix(1, 1) = "" Then CmbPage.Text = ""
        'CmbPage.SelectedIndex = 0
        'CmbPage.SelectedIndex = 0
    End Sub

    Private Sub p100_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles p100.CheckedChanged

    End Sub

    Private Sub p100_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p100.MouseClick
        txtSC15.Enabled = False
        P15.ForeColor = Color.Black
        p25.ForeColor = Color.Black
        p50.ForeColor = Color.Black
        p250.ForeColor = Color.Black
        p500.ForeColor = Color.Black
        p1000.ForeColor = Color.Black
        If P15.Checked = True Then
            txtSC15.Enabled = True
            P15.ForeColor = Color.Red : txtSC15.Focus() : txtSC15.SelectAll()
        ElseIf p25.Checked = True Then
            p25.ForeColor = Color.Red
        ElseIf p50.Checked = True Then
            p50.ForeColor = Color.Red
        ElseIf p100.Checked = True Then
            p100.ForeColor = Color.Red
        ElseIf p250.Checked = True Then
            p250.ForeColor = Color.Red

        ElseIf p500.Checked = True Then
            p500.ForeColor = Color.Red
        ElseIf p1000.Checked = True Then
            p1000.ForeColor = Color.Red
        End If
        Call LoadDividePage()
        P = 1

        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page
        If FG.get_TextMatrix(1, 1) = "" Then lblpage_total.Text = "0/0"
        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        If CmbPage.Items.Count > 0 Then
            CmbPage.SelectedIndex = 0
        End If
        If FG.get_TextMatrix(1, 1) = "" Then CmbPage.Text = ""
        'CmbPage.SelectedIndex = 0
    End Sub

    Private Sub p250_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles p250.CheckedChanged

    End Sub

    Private Sub p250_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p250.MouseClick
        txtSC15.Enabled = False
        P15.ForeColor = Color.Black
        p25.ForeColor = Color.Black
        p50.ForeColor = Color.Black
        p250.ForeColor = Color.Black
        p500.ForeColor = Color.Black
        p1000.ForeColor = Color.Black
        If P15.Checked = True Then
            txtSC15.Enabled = True
            P15.ForeColor = Color.Red : txtSC15.Focus() : txtSC15.SelectAll()
        ElseIf p25.Checked = True Then
            p25.ForeColor = Color.Red
        ElseIf p50.Checked = True Then
            p50.ForeColor = Color.Red
        ElseIf p100.Checked = True Then
            p100.ForeColor = Color.Red
        ElseIf p250.Checked = True Then
            p250.ForeColor = Color.Red
        ElseIf p500.Checked = True Then
            p500.ForeColor = Color.Red
        ElseIf p1000.Checked = True Then
            p1000.ForeColor = Color.Red
        End If
        Call LoadDividePage()
        P = 1

        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page
        If FG.get_TextMatrix(1, 1) = "" Then lblpage_total.Text = "0/0"
        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        If CmbPage.Items.Count > 0 Then
            CmbPage.SelectedIndex = 0
        End If
        If FG.get_TextMatrix(1, 1) = "" Then CmbPage.Text = ""
        'CmbPage.SelectedIndex = 0
    End Sub

    Private Sub p500_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles p500.CheckedChanged

    End Sub

    Private Sub p500_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p500.MouseClick
        txtSC15.Enabled = False
        P15.ForeColor = Color.Black
        p25.ForeColor = Color.Black
        p50.ForeColor = Color.Black
        p250.ForeColor = Color.Black
        p500.ForeColor = Color.Black
        p1000.ForeColor = Color.Black
        If P15.Checked = True Then
            txtSC15.Enabled = True
            P15.ForeColor = Color.Red : txtSC15.Focus() : txtSC15.SelectAll()
        ElseIf p25.Checked = True Then
            p25.ForeColor = Color.Red
        ElseIf p50.Checked = True Then
            p50.ForeColor = Color.Red
        ElseIf p100.Checked = True Then
            p100.ForeColor = Color.Red
        ElseIf p250.Checked = True Then
            p250.ForeColor = Color.Red
        ElseIf p500.Checked = True Then
            p500.ForeColor = Color.Red
        ElseIf p1000.Checked = True Then
            p1000.ForeColor = Color.Red
        End If
        Call LoadDividePage()
        P = 1

        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page
        If FG.get_TextMatrix(1, 1) = "" Then lblpage_total.Text = "0/0"
        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        If CmbPage.Items.Count > 0 Then
            CmbPage.SelectedIndex = 0
        End If
        If FG.get_TextMatrix(1, 1) = "" Then CmbPage.Text = ""
        'CmbPage.SelectedIndex = 0
    End Sub

    Private Sub p1000_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles p1000.CheckedChanged

    End Sub

    Private Sub p1000_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p1000.MouseClick
        txtSC15.Enabled = False
        P15.ForeColor = Color.Black
        p25.ForeColor = Color.Black
        p50.ForeColor = Color.Black
        p250.ForeColor = Color.Black
        p500.ForeColor = Color.Black
        p1000.ForeColor = Color.Black
        If P15.Checked = True Then
            txtSC15.Enabled = True
            P15.ForeColor = Color.Red : txtSC15.Focus() : txtSC15.SelectAll()
        ElseIf p25.Checked = True Then
            p25.ForeColor = Color.Red
        ElseIf p50.Checked = True Then
            p50.ForeColor = Color.Red
        ElseIf p100.Checked = True Then
            p100.ForeColor = Color.Red
        ElseIf p250.Checked = True Then
            p250.ForeColor = Color.Red
        ElseIf p500.Checked = True Then
            p500.ForeColor = Color.Red
        ElseIf p1000.Checked = True Then
            p1000.ForeColor = Color.Red
        End If
        Call LoadDividePage()
        P = 1

        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page
        If FG.get_TextMatrix(1, 1) = "" Then lblpage_total.Text = "0/0"
        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        If CmbPage.Items.Count > 0 Then
            CmbPage.SelectedIndex = 0
        End If
        If FG.get_TextMatrix(1, 1) = "" Then CmbPage.Text = ""
        'CmbPage.SelectedIndex = 0
    End Sub

    Private Sub P15_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles P15.CheckedChanged

    End Sub

    Private Sub P15_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles P15.MouseClick
        txtSC15.Enabled = False
        P15.ForeColor = Color.Black
        p25.ForeColor = Color.Black
        p50.ForeColor = Color.Black
        p250.ForeColor = Color.Black
        p500.ForeColor = Color.Black
        p1000.ForeColor = Color.Black
        If P15.Checked = True Then
            txtSC15.Enabled = True

        ElseIf p25.Checked = True Then
            p25.ForeColor = Color.Red
        ElseIf p50.Checked = True Then
            p50.ForeColor = Color.Red
        ElseIf p100.Checked = True Then
            p100.ForeColor = Color.Red
        ElseIf p250.Checked = True Then
            p250.ForeColor = Color.Red
        ElseIf p500.Checked = True Then
            p500.ForeColor = Color.Red
        ElseIf p1000.Checked = True Then
            p1000.ForeColor = Color.Red
        End If
        Call LoadDividePage()
        P = 1

        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page
        If FG.get_TextMatrix(1, 1) = "" Then lblpage_total.Text = "0/0"

        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i

        If CmbPage.Items.Count > 0 Then
            CmbPage.SelectedIndex = 0
        End If
        If FG.get_TextMatrix(1, 1) = "" Then CmbPage.Text = ""
        P15.ForeColor = Color.Red : txtSC15.Focus() : txtSC15.SelectAll()
        'CmbPage.SelectedIndex = 0
        'CmbPage.SelectedIndex = 0
    End Sub

    Private Sub txtSC15_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSC15.KeyPress
        If e.KeyChar = Chr(13) Then
            Call StartLoadDataList()
        End If
    End Sub

    Private Sub txtSC_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSC15.TextChanged
        If IsNumeric(txtSC15.Text) = False Then txtSC15.Text = "1" : Exit Sub
        If txtSC15.Text = "0" Then txtSC15.Text = "1" : Exit Sub
        If txtSC15.Text = "" Then txtSC15.Text = "1" : Exit Sub
    End Sub


    Private Sub ComboBox2_QueryAccessibilityHelp(ByVal sender As Object, ByVal e As System.Windows.Forms.QueryAccessibilityHelpEventArgs)

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub dts_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dts.KeyPress
        If e.KeyChar = Chr(13) Then
            LoadMonthSQL()
            dtt.Focus()
        End If
    End Sub

    Private Sub dts_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dts.ValueChanged
        dtt.Text = dts.Text
    End Sub

    Private Sub dtt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dtt.KeyPress
        If e.KeyChar = Chr(13) Then
            LoadMonthSQL()
        End If
    End Sub

    Private Sub dtt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtt.ValueChanged

    End Sub

    Private Sub Panel6_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel6.Paint

    End Sub

    Private Sub LbPage_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles LbPage.KeyPress
        If e.KeyChar = Chr(13) Then

            'If EnterPage.Text = "Next " Then

            EnterPage_Click(sender, e)
            '    'If NextPage.Enabled = False Then EnterPage.Text = "Back" : Exit Sub
            '    Exit Sub
            'End If
            'If EnterPage.Text = "Back" Then
            '    'If P = 2 Then EnterPage.Text = "Next "
            '    BackPage_Click(sender, e)
            '    Exit Sub
            '    'If BackPage.Enabled = False Then EnterPage.Text = "Next " : Exit Sub
            'End If

        End If
    End Sub


    Private Sub EnterPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EnterPage.Click


        If EnterPage.Text = "Next  " Then
            NextPage_Click(sender, e)
            If CmbPage.SelectedIndex = CmbPage.Items.Count - 1 Then
                EnterPage.Text = "Back "
            End If
        Else
            BackPage_Click(sender, e)
            If CmbPage.SelectedIndex = 0 Then
                EnterPage.Text = "Next  "
            End If
        End If





        'If LasthPage.Enabled = False Then EnterPage.Text = "Back" : FirstPage_Click(sender, e) : LbPage.Focus() : Exit Sub
        'If FirstPage.Enabled = False Then EnterPage.Text = "Next " : LasthPage_Click(sender, e) : LbPage.Focus() : Exit Sub

        'If LasthPage.Enabled = False Then EnterPage.Text = "Back" : LbPage.Focus() : Exit Sub
        'If FirstPage.Enabled = False Then EnterPage.Text = "Next " : LbPage.Focus() : Exit Sub

        'If EnterPage.Text = "Back " Then EnterPage.Text = "Next " : LbPage.Focus() : Exit Sub
        'If EnterPage.Text = "Next " Then EnterPage.Text = "Back " : LbPage.Focus() : Exit Sub

    End Sub




    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Panel5.Visible = False




    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CntNB = "certify , cnt"
        Call StartLoadDataList()
        Panel5.Visible = False
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CntNB = "date_work, cnt"
        CheckBox3.Checked = False
        Call StartLoadDataList()
        Panel5.Visible = False
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CntNB = "cheque_no , cnt"
        Call StartLoadDataList()
        CheckBox3.Checked = False
        Panel5.Visible = False
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CntNB = "ac_code, cnt"
        Call StartLoadDataList()
        Panel5.Visible = False
        CheckBox3.Checked = False
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        CntNB = "descrip, cnt"
        Call StartLoadDataList()
        Panel5.Visible = False
        CheckBox3.Checked = False
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CntNB = "book, cnt"
        Call StartLoadDataList()
        CheckBox3.Checked = False
        Panel5.Visible = False
    End Sub



    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CntNB = "curr , cnt"
        Call StartLoadDataList()
        CheckBox3.Checked = False
        Panel5.Visible = False
    End Sub

    Private Sub Label28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Off_Usr_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Off_Usr.SelectedIndexChanged
        MuSubOff = Mid(Off_Usr.Text, 1, 5)
        'Loadfind()
        Panel4.Visible = False

        LoadMonthSQL()
    End Sub

    Private Sub Panel3_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel3.Paint

    End Sub

    Private Sub LbPage_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LbPage.TextChanged

    End Sub

    Private Sub Rinvioce_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Rinvioce.CheckedChanged
        Nme.Enabled = True
        Nme.Visible = True
        ComboBox1.Visible = False
        Nme.Focus()
    End Sub

    Private Sub CmbCompany_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCompany.SelectedIndexChanged
        LoadSubCompany()
    End Sub
    Private Sub LoadSubCompany()
        Off_Usr.Items.Clear()
        Dim dtTemp As DataTable = DbHelper.GetDataTable("select sub_id , off_id , off_add2  from  Ap_office where off_id ='" & Mid(cmbCompany.Text, 1, 2) & "' group BY  sub_id  ,off_id , off_add2")
        For Each row As DataRow In dtTemp.Rows
            Off_Usr.Items.Add(DbHelper.GetStr(row("sub_id")) & " " & DbHelper.GetStr(row("off_add2")))
        Next

        Off_Usr.SelectedIndex = FmLogin.Sub_Company.SelectedIndex
        Off_Id = Mid(cmbCompany.Text, 1, 2)
        SUPD = 0
    End Sub

    Private Sub ContextMenuStrip1_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening
        'If Lng = "L" Then
        LngId = "3001" : CallLngStr() : CMS1.Text = LngStr
        LngId = "3003" : CallLngStr() : CMS2.Text = LngStr
        LngId = "3004" : CallLngStr() : CMS3.Text = LngStr
        LngId = "3005" : CallLngStr() : CMS4.Text = LngStr
        LngId = "3007" : CallLngStr() : CMS5.Text = LngStr
        LngId = "3006" : CallLngStr() : CMS6.Text = LngStr
        LngId = "5043" : CallLngStr() : date_work.Text = LngStr
        LngId = "5015" : CallLngStr() : certify.Text = LngStr
        LngId = "5016" : CallLngStr() : cheque_no.Text = LngStr
        LngId = "5017" : CallLngStr() : ac_code.Text = LngStr
        LngId = "5019" : CallLngStr() : descrip.Text = LngStr
        LngId = "5020" : CallLngStr() : Book.Text = LngStr
        LngId = "5022" : CallLngStr() : Curr.Text = LngStr
        LngId = "6006" : CallLngStr() : CMS7.Text = LngStr
        LngId = "6007" : CallLngStr() : MASC.Text = LngStr
        LngId = "6008" : CallLngStr() : MDESC.Text = LngStr




    End Sub


    Private Sub NextPage_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NextPage.MouseClick
        If NextPage.Enabled = True Then
            EnterPage.Text = "Next  "
        Else
            EnterPage.Text = "Back "
        End If

    End Sub

    Private Sub BackPage_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BackPage.MouseClick
        If BackPage.Enabled = True Then
            EnterPage.Text = "Back "
        Else
            EnterPage.Text = "Next  "

        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged

    End Sub

    Private Sub GdgToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMS1.Click
        Panel4.Visible = False
        FmNsewJeneralJournal.txtInvoice.Enabled = True
        FmNsewJeneralJournal.CmbBook.Enabled = True

        FmNsewJeneralJournal.ShowDialog()
    End Sub



    Private Sub TextBox1_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub DgfdgToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMS2.Click
        Panel4.Visible = False
        If MDInvoiceNo <> "" Then
            If LockData = 1 Then
                MsgBox("ບັນຊີນີ້ຖືກລ໋ອກບໍ່ສາມາດແກ້ໄຂໄດ້")
                Exit Sub
            End If
            If LockData = 2 Then
                MsgBox("ບັນຊີນີ້ໄດ້ປິດບັນຊີໄປແລ້ວກບໍ່ສາມາດແກ້ໄຂໄດ້")
                Exit Sub
            End If
            FmNsewJeneralJournal.txtInvoice.Enabled = False
            FmNsewJeneralJournal.CmbBook.Enabled = False
            FmNsewJeneralJournal.ShowDialog()
        Else
            MessageBox.Show("ກະລຸນນາເລືອກກ່ອນ")
            Exit Sub
        End If
    End Sub

    Private Sub DgdfToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMS3.Click
        Panel4.Visible = False
        If MDInvoiceNo = "" Then
            'MessageBox.Show("ກະລຸນນາເລືອກກ່ອນ")
            MessageBox.Show("ກະລຸນນາເລືອກກ່ອນ", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If LockData = 1 Then
            MessageBox.Show("ບັນຊີນີ້ຖືກລ໋ອກບໍ່ສາມາດລືບໄດ້", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If LockData = 2 Then

            MessageBox.Show("ບັນຊີນີ້ໄດ້ປິດບັນຊີໄປແລ້ວກບໍ່ສາມາດລືບໄດ້", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If



        If MessageBox.Show("ທ່ານຕ້ອງການລຶບ  " & MDInvoiceNo & " ຫລືບໍ່?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

            DbHelper.ExecuteNonQuery("delete gen_jn where certify ='" & MDInvoiceNo & "' And  year(date_work)='" & Format(CDate(GetGridValue(FG, FG.CurrentRow.Index, 1)), "yyyy") & "' ")

        End If
        MDInvoiceNo = ""
        LoadMonthSQL()
    End Sub

    Private Sub ເອນຂມນToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMS4.Click
        Panel4.Visible = False
        LoadMonthSQL()
    End Sub

    Private Sub ຈດລຽງແຕໃຫຍຫານອຍToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MASC.Click
        MDESC.Checked = False
        MASC.Checked = True
        Load_DES()
        Call StartLoadDataList()
    End Sub

    Private Sub ຈດລຽງແຕນອຍຫາໃຫຍToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MDESC.Click

        MDESC.Checked = True
        MASC.Checked = False
        Load_DES()
        Call StartLoadDataList()
    End Sub

    Private Sub ວນທToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles date_work.Click
        date_work.Checked = True
        certify.Checked = False
        cheque_no.Checked = False
        ac_code.Checked = False
        descrip.Checked = False
        Book.Checked = False
        Curr.Checked = False
        Call Load_DES()
        Call StartLoadDataList()
    End Sub

    Private Sub ໃບຢງຢນToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles certify.Click
        date_work.Checked = False
        certify.Checked = True
        cheque_no.Checked = False
        ac_code.Checked = False
        descrip.Checked = False
        Book.Checked = False
        Curr.Checked = False
        Call Load_DES()
        Call StartLoadDataList()
    End Sub

    Private Sub ເລກແຊກToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cheque_no.Click
        date_work.Checked = False
        certify.Checked = False
        cheque_no.Checked = True
        ac_code.Checked = False
        descrip.Checked = False
        Book.Checked = False
        Curr.Checked = False
        Call Load_DES()
        Call StartLoadDataList()
    End Sub

    Private Sub ລະຫດບນຊToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ac_code.Click
        date_work.Checked = False
        certify.Checked = False
        cheque_no.Checked = False
        ac_code.Checked = True
        descrip.Checked = False
        Book.Checked = False
        Curr.Checked = False
        Call Load_DES()
        Call StartLoadDataList()
    End Sub

    Private Sub ລະຫດບນຊToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles descrip.Click
        date_work.Checked = False
        certify.Checked = False
        cheque_no.Checked = False
        ac_code.Checked = False
        descrip.Checked = True
        Book.Checked = False
        Curr.Checked = False
        Call Load_DES()
        Call StartLoadDataList()
    End Sub

    Private Sub ເນອໃນລາຍການToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Book.Click
        date_work.Checked = False
        certify.Checked = False
        cheque_no.Checked = False
        ac_code.Checked = False
        descrip.Checked = False
        Book.Checked = True
        Curr.Checked = False
        Call Load_DES()
        Call StartLoadDataList()
    End Sub

    Private Sub ສະກນເງນToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Curr.Click
        date_work.Checked = False
        certify.Checked = False
        cheque_no.Checked = False
        ac_code.Checked = False
        descrip.Checked = False
        Book.Checked = False
        Curr.Checked = True
        Call Load_DES()
        Call StartLoadDataList()
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMS5.Click
        Button11_Click(sender, e)
    End Sub

    Private Sub ToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMS6.Click
        Button3_Click(sender, e)
    End Sub

    Private Sub TT()
        Dim dtTemp As DataTable = DbHelper.GetDataTable("select top 1 count(cnt) as cnt from gen_jn")
        If dtTemp.Rows.Count > 0 Then
            TextBox2.Text = CDbl(DbHelper.GetStr(dtTemp.Rows(0)("cnt")))
        End If
    End Sub
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TT()

        TextBox4.Text = CDbl(TextBox2.Text) / (txtSC15.Text)
        FG.CurrentRow.Indexs = 1
        With RSC
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT top " & txtSC15.Text & " * FROM  Gen_Jn Order by cnt")
            If .RecordCount > 0 Then
                While Not .EOF
                    'kkkkkk
                    If MuLng = "L" Then s = Trim(CStr(.Fields("descrip").Value)) Else s = Trim(CStr(.Fields("descripe").Value.ToString))

                    'End If

                    FG.AddItem(.AbsolutePosition & vbTab & Format(CDate(Trim(.Fields("date_work").Value)), "dd/MM/yyyy") & _
                                "" & vbTab & Trim(CStr(.Fields("certify").Value)) & _
                                 "" & vbTab & Trim(CStr(.Fields("cheque_no").Value)) & _
                                "" & vbTab & Trim(CStr(.Fields("Ac_code").Value)) & _
                               "" & vbTab & Format(CDbl(Trim(.Fields("amount").Value)), "##,##0.00") & _
                                 "" & vbTab & s & _
                                 "" & vbTab & Format(CDbl(Trim(.Fields("amount_dr").Value)), "##,##0.00") & _
                                  "" & vbTab & Format(CDbl(Trim(.Fields("amount_cr").Value)), "##,##0.00") & _
                                     "" & vbTab & Trim(CStr(.Fields("curr").Value)) & _
                                  "" & vbTab & Format(CDbl(Trim(.Fields("amt_dr").Value)), "##,##0.00") & _
                                  "" & vbTab & Format(CDbl(Trim(.Fields("amt_cr").Value)), "##,##0.00") & _
                                 "" & vbTab & Trim(CStr(.Fields("company").Value)) & _
                                    "" & vbTab & Trim(CStr(.Fields("cnt").Value)) & _
                                   "" & vbTab & Trim(CStr(.Fields("lock").Value)) & _
                                    "" & vbTab & "" & _
                                "" & vbTab & Trim(CStr(.Fields("book").Value)))

                    'kkkkkkkkkkk

                    .MoveNext()
                End While
            Else
                FG.CurrentRow.Indexs = 16
            End If
        End With
    End Sub

    Private Sub Nme_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Nme.TextChanged

    End Sub

    Private Sub Button19_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        TT()
        D = FG.CurrentRow.Indexs - 1
        If CDbl(CDbl(TextBox2.Text) / CDbl(txtSC15.Text)) > Int(CDbl(TextBox2.Text) / CDbl(txtSC15.Text)) Then
            TextBox4.Text = Int(CDbl(TextBox2.Text) / CDbl(txtSC15.Text)) + 1
        Else
            TextBox4.Text = Int(CDbl(TextBox2.Text) / CDbl(txtSC15.Text))
        End If
        TextBox6.Text = 1
        TextBox5.Text = TextBox6.Text & "/" & TextBox4.Text
        FG.CurrentRow.Indexs = 1
        With RSC
            Dim dtTemp As DataTable = DbHelper.GetDataTable("SELECT top " & txtSC15.Text & " * FROM  Gen_Jn  Order by certify , cnt")
            If .RecordCount > 0 Then
                While Not .EOF
                    'kkkkkk
                    If MuLng = "L" Then s = Trim(CStr(.Fields("descrip").Value)) Else s = Trim(CStr(.Fields("descripe").Value.ToString))

                    'End If

                    FG.AddItem(.AbsolutePosition & vbTab & Format(CDate(Trim(.Fields("date_work").Value)), "dd/MM/yyyy") & _
                                "" & vbTab & Trim(CStr(.Fields("certify").Value)) & _
                                 "" & vbTab & Trim(CStr(.Fields("cheque_no").Value)) & _
                                "" & vbTab & Trim(CStr(.Fields("Ac_code").Value)) & _
                               "" & vbTab & Format(CDbl(Trim(.Fields("amount").Value)), "##,##0.00") & _
                                 "" & vbTab & s & _
                                 "" & vbTab & Format(CDbl(Trim(.Fields("amount_dr").Value)), "##,##0.00") & _
                                  "" & vbTab & Format(CDbl(Trim(.Fields("amount_cr").Value)), "##,##0.00") & _
                                     "" & vbTab & Trim(CStr(.Fields("curr").Value)) & _
                                  "" & vbTab & Format(CDbl(Trim(.Fields("amt_dr").Value)), "##,##0.00") & _
                                  "" & vbTab & Format(CDbl(Trim(.Fields("amt_cr").Value)), "##,##0.00") & _
                                 "" & vbTab & Trim(CStr(.Fields("company").Value)) & _
                                    "" & vbTab & Trim(CStr(.Fields("lock").Value)) & _
                                   "" & vbTab & Trim(CStr(.Fields("lock").Value)) & _
                                    "" & vbTab & "" & _
                                "" & vbTab & Trim(CStr(.Fields("book").Value)))

                    'kkkkkkkkkkk

                    .MoveNext()
                End While
            Else
                FG.CurrentRow.Indexs = 16
            End If
        End With
    End Sub

    Private Sub Button18_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click

    End Sub

    Private Sub Button17_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        'TT()

        Dim x As String = " date_work , certify ,descrip,descripe, cheque_no , Ac_code,code_dr , code_cr  , amount_dr , amount_cr , amount ,  amt_dr , amt_cr  ,  curr , cnt , lock  , book, Company "
        Dim s As String = "SELECT top " & txtSC15.Text & " " & x & "  FROM   Gen_Jn where certify > '" & FG.get_TextMatrix(FG.CurrentRow.Indexs - 1, 2) & "' Order by certify , cnt"
        TextBox6.Text = Int(TextBox6.Text) + 1
        TextBox5.Text = TextBox6.Text & "/" & TextBox4.Text
        FG.CurrentRow.Indexs = 1
        Dim R As Integer
        With RSC
            Dim dtTemp As DataTable = DbHelper.GetDataTable(s)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.CurrentRow.Indexs = FG.CurrentRow.Indexs + 1
                    If MuLng = "L" Then s = Trim(CStr(.Fields("descrip").Value)) Else s = Trim(CStr(.Fields("descripe").Value.ToString))
                    R = FG.CurrentRow.Indexs - 1
                    D = D + 1
                    FG.set_TextMatrix(R, 0, D)
                    FG.set_TextMatrix(R, 1, Format(CDate(Trim(.Fields("date_work").Value)), "dd/MM/yyyy"))
                    FG.set_TextMatrix(R, 2, Trim(CStr(.Fields("certify").Value)))
                    FG.set_TextMatrix(R, 3, Trim(CStr(.Fields("cheque_no").Value)))
                    FG.set_TextMatrix(R, 4, Trim(CStr(.Fields("Ac_code").Value)))
                    FG.set_TextMatrix(R, 5, Format(CDbl(Trim(.Fields("amount").Value)), "##,##0.00"))
                    FG.set_TextMatrix(R, 6, s)
                    FG.set_TextMatrix(R, 7, Format(CDbl(Trim(.Fields("amount_dr").Value)), "##,##0.00"))
                    FG.set_TextMatrix(R, 8, Format(CDbl(Trim(.Fields("amount_Cr").Value)), "##,##0.00"))
                    FG.set_TextMatrix(R, 9, Trim(CStr(.Fields("Curr").Value)))
                    FG.set_TextMatrix(R, 10, Format(CDbl(Trim(.Fields("amt_dr").Value)), "##,##0.00"))
                    FG.set_TextMatrix(R, 11, Format(CDbl(Trim(.Fields("amt_dr").Value)), "##,##0.00"))
                    FG.set_TextMatrix(R, 12, Trim(CStr(.Fields("company").Value)))
                    FG.set_TextMatrix(R, 13, Trim(CStr(.Fields("cnt").Value)))
                    FG.set_TextMatrix(R, 14, Trim(CStr(.Fields("lock").Value)))
                    FG.set_TextMatrix(R, 15, "")
                    FG.set_TextMatrix(R, 16, Trim(CStr(.Fields("book").Value)))


                    'kkkkkkkkkkk

                    .MoveNext()
                End While
            Else
                FG.CurrentRow.Indexs = 16
            End If
        End With
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click

    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        If RadioButton15.Checked = True Then
            If LockData = "" Then MsgBox("ກະລຸນາເລືອກກ່ອນ!", MsgBoxStyle.OkOnly) : Exit Sub
            Dim dtTemp As DataTable = DbHelper.GetDataTable("Select lock from gen_jn where certify ='" & MDInvoiceNo & "' And  year(date_work)='" & Format(CDate(GetGridValue(FG, FG.CurrentRow.Index, 1)), "yyyy") & "'")
            If dtTemp.Rows.Count <> 0 Then
                If DbHelper.GetStr(dtTemp.Rows(0)("lock")) = "2" Then
                    MsgBox("ລາຍການນີ້ ໄດ້ປິດບັນຊີໄປແລ້ວບໍ່ສາມາດ " & Button1.Text & "  ໄດ້ອີກ!", MsgBoxStyle.OkOnly)
                    Exit Sub
                End If
            End If

            If MessageBox.Show("ທ່ານຕ້ອງການ " & Button1.Text & " ລະຫັດ " & MDInvoiceNo & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                If LockData = "1" Then
                    DbHelper.ExecuteNonQuery("UPDATE gen_jn SET lock='0' where certify ='" & MDInvoiceNo & "' And  year(date_work)='" & Format(CDate(GetGridValue(FG, FG.CurrentRow.Index, 1)), "yyyy") & "'")
                Else
                    DbHelper.ExecuteNonQuery("UPDATE gen_jn SET lock='1' where certify ='" & MDInvoiceNo & "' And  year(date_work)='" & Format(CDate(GetGridValue(FG, FG.CurrentRow.Index, 1)), "yyyy") & "'")

                End If
                LockData = ""
                LoadMonthSQL()
            End If
        Else
            Dim MS As String = "And Company = '" & MuSubOff & "'"
            Dim OfUsr1 As String = Microsoft.VisualBasic.Left(Off_Usr.Text, 5)
            Dim OfUsr2 As String = Mid(Off_Usr.Text, 4, 2)
            Dim OfUsr3 As String = Microsoft.VisualBasic.Left(Off_Usr.Text, 2)
            If OfUsr1 = "00-00" Then
                MULook2 = ""
            Else
                If OfUsr2 = "00" Then
                    MULook2 = "  And  Left(gen_jn.company,2)= '" & OfUsr3 & "' "
                Else
                    MULook2 = "  And gen_jn.company= '" & OfUsr1 & "' "
                End If
            End If

            Dim dtTemp As DataTable = DbHelper.GetDataTable("Select lock from gen_jn where  gen_jn.date_work   BETWEEN '" & Format(dts.Value, "yyyy-MM-dd") & "' AND '" & Format(dtt.Value, "yyyy-MM-dd") & "'   " & MULook2 & " ")
            If dtTemp.Rows.Count <> 0 Then
                If DbHelper.GetStr(dtTemp.Rows(0)("lock")) = "2" Then
                    MsgBox("ລາຍການພວກນີ້ ໄດ້ປິດບັນຊີໄປແລ້ວບໍ່ສາມາດ " & Button1.Text & "  ໄດ້ອີກ!", MsgBoxStyle.OkOnly)
                    Exit Sub
                End If
            End If

            If MessageBox.Show("ທ່ານຕ້ອງການປົດລ໋ອກຂໍ່ມູນແຕວັນທີ " & dts.Text & " ຫາ " & dtt.Text & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                If RadioButton17.Checked = True Then
                    DbHelper.ExecuteNonQuery("UPDATE gen_jn SET lock='1' where  gen_jn.date_work   BETWEEN '" & Format(dts.Value, "yyyy-MM-dd") & "' AND '" & Format(dtt.Value, "yyyy-MM-dd") & "'   " & MULook2 & " ")
                Else
                    DbHelper.ExecuteNonQuery("UPDATE gen_jn SET lock='0'  where  gen_jn.date_work   BETWEEN '" & Format(dts.Value, "yyyy-MM-dd") & "' AND '" & Format(dtt.Value, "yyyy-MM-dd") & "'   " & MULook2 & " ")
                End If
                LockData = ""
                LoadMonthSQL()
            End If
        End If

    End Sub

    Private Sub Button13_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Panel7.Visible = False
    End Sub

    Private Sub Button21_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        LoadData4()
        LoadDataSQL()
    End Sub
    Dim trn_id As String
    Dim trn_desc As String
    Dim Currency As String
    Dim acc_book As String
    Dim ex_rate As Double
    Dim SUMAMT As Double
    Dim dr_ac, dr_desc As String
    Dim dr_amt As Double
    Dim bis_date As Date
    Dim MdCertifyId As String
    Public Sub LoadData3()

        Dim connStr As String = "Server=" & MDServerName2 & ";Database=" & MDDatabaName2 & ";Uid=" & MDServerUser2 & ";Pwd=" & MDServerPassword2 & ";SslMode=None;"
        'Dim connStr As String = "Server=apis.com.la;Database=apb_msp;Uid=Admin;Pwd=Sql_admin@#2024;Port=3306;SslMode=None;"
        '"Server=apis.com.la;Database=apb_msp;Uid=Admin;Pwd=Sql_admin@#2024;Port=3306;SslMode=None;"
        '"Server=10.151.146.91;Database=apb_msp;Uid=admin;Pwd=ApbAdmin@2025;Port=3306;SslMode=None;"
        Dim conn As New MySqlConnection(connStr)
        conn.Open()

        ' Dim connStr As String = "server=localhost;user id=Admin;password=ApbAdmin@2025;database=your_db"
        'Dim conn As New MySql.Data.MySqlClient.MySqlConnection(connStr)
        'conn.Open()

        Dim tx = conn.BeginTransaction()
        Try

            ' Load pending msp rows
            Dim dtMain As New DataTable()
            Using da As New MySqlDataAdapter("SELECT trn_id, trn_desc, acc_book, Currency, ex_rate, bis_date FROM msp WHERE status = 'Wait' ORDER BY trn_id", conn)
                da.Fill(dtMain)
            End Using


            For Each rMain As DataRow In dtMain.Rows
                Dim trn_id As String = rMain("trn_id").ToString()
                trn_desc = If(IsDBNull(rMain("trn_desc")), "", rMain("trn_desc").ToString())
                acc_book = If(IsDBNull(rMain("acc_book")), "", rMain("acc_book").ToString())
                Currency = If(IsDBNull(rMain("Currency")), "", rMain("Currency").ToString())
                ex_rate = If(IsDBNull(rMain("ex_rate")), 1, Convert.ToDecimal(rMain("ex_rate")))
                If ex_rate = 0 Then
                    ex_rate = 1
                End If
                bis_date = If(IsDBNull(rMain("bis_date")), Date.Today, Convert.ToDateTime(rMain("bis_date")))
                Call AutoNumber()
                ' Get SUMAMT (safe scalar)
                Dim SUMAMT As Decimal = 0
                Dim SUMAMT_LAK As Decimal = 0
                Using cmdSum As New MySqlCommand("SELECT SUM(dr_amt), SUM(dr_amt_LAK) FROM tbl_dr WHERE trn_id = @trn_id", conn, tx)
                    cmdSum.Parameters.AddWithValue("@trn_id", trn_id)
                    Dim res = cmdSum.ExecuteScalar()
                    If res IsNot Nothing AndAlso Not IsDBNull(res) Then SUMAMT = Convert.ToDecimal(res)
                    ''=======================
                    'Using rdr As MySqlDataReader = cmdSum.ExecuteReader(CommandBehavior.SingleRow)
                    '    If rdr.Read() Then
                    '        ' COALESCE ensures non-NULL, but keep checks for safety
                    '        If Not rdr.IsDBNull(0) Then SUMAMT = Convert.ToDecimal(rdr.GetValue(0))
                    '        If Not rdr.IsDBNull(1) Then SUMAMT_LAK = Convert.ToDecimal(rdr.GetValue(1))
                    '    End If
                    'End Using
                End Using

                ' DR rows
                Dim dtDr As New DataTable()
                Using daDr As New MySqlDataAdapter("SELECT dr_ac, dr_amt, dr_amt_lak, dr_desc FROM tbl_dr WHERE trn_id = @trn_id ORDER BY dr_ac", conn)
                    daDr.SelectCommand.Transaction = tx
                    daDr.SelectCommand.Parameters.AddWithValue("@trn_id", trn_id)
                    daDr.Fill(dtDr)
                End Using

                For Each rDr As DataRow In dtDr.Rows
                    Dim dr_ac As String = rDr("dr_ac").ToString()
                    Dim dr_desc As String = If(IsDBNull(rDr("dr_desc")), "", rDr("dr_desc").ToString())
                    Dim dr_amt As Decimal = If(IsDBNull(rDr("dr_amt")), 0, Convert.ToDecimal(rDr("dr_amt")))
                    Dim dr_amt_lak As Decimal = If(IsDBNull(rDr("dr_amt_lak")), 1, Convert.ToDecimal(rDr("dr_amt_lak")))
                    Dim KKK As String = "INSERT INTO gen_jn( date_work, ac_Name, book, certify, Referno ,descrip ,descripe , " & _
                                          " amount , curr ,rate, Rate_USD, net_amt ,code_dr ,code_cr ,ac_code ,amt_dr , amt_cr , amt_USD_Dr, amt_USD_Cr , amount_dr ,amount_cr ,  " & _
                                          " certis, lock ,rec_lock , last_update , last_user  ,company  ,Office_ID, del , AG, Frm, API ) " & _
                                          " Values('" & Format(CDate(bis_date), "yyyy/MM/dd") & "', " & _
                                          " N'" & trn_desc & "', " & _
                                          " N'" & acc_book & "', " & _
                                          " N'" & MdCertifyId & "', " & _
                                          " N'" & trn_id & "', " & _
                                          " N'" & trn_desc & "', " & _
                                          " N'', " & _
                                          " " & CDbl(SUMAMT) & ", " & _
                                          " N'" & Currency & "', " & _
                                          " " & CDbl(ex_rate) & ", " & _
                                          " " & CDbl(ex_rate) & ", " & _
                                          " 0, " & _
                                          " '" & dr_ac & "', " & _
                                          " '', " & _
                                          " '" & dr_ac & "', " & _
                                          " " & CDbl(dr_amt_lak) & ", " & _
                                          " 0, " & _
                                          " " & CDbl(dr_amt) & ", " & _
                                          " 0, " & _
                                          " " & CDbl(dr_amt) / CDbl(ex_rate) & ", " & _
                                          " " & CDbl(0) & ", " & _
                                          " 3, 4, 5, " & _
                                          " '" & Format(CDate(Date.Today), "yyyy/MM/dd") & "', " & _
                                          " '" & MUserID & "', " & _
                                          " '" & MuSubOff & "' , " & _
                                          " '" & MuSubOff & "', 0,1,0, 'API' )"
                    DbHelper.ExecuteNonQuery(KKK)
                Next



                ' CR rows
                Dim dtCr As New DataTable()
                Using daCr As New MySqlDataAdapter("SELECT cr_ac, cr_amt, cr_desc, cr_amt_lak FROM tbl_cr WHERE trn_id = @trn_id ORDER BY cr_ac", conn)
                    daCr.SelectCommand.Transaction = tx
                    daCr.SelectCommand.Parameters.AddWithValue("@trn_id", trn_id)
                    daCr.Fill(dtCr)
                End Using

                For Each rCr As DataRow In dtCr.Rows
                    Dim cr_ac As String = rCr("cr_ac").ToString()
                    Dim cr_desc As String = If(IsDBNull(rCr("cr_desc")), "", rCr("cr_desc").ToString())
                    Dim cr_amt As Decimal = If(IsDBNull(rCr("cr_amt")), 0, Convert.ToDecimal(rCr("cr_amt")))
                    Dim cr_amt_lak As Decimal = If(IsDBNull(rCr("cr_amt_lak")), 0, Convert.ToDecimal(rCr("cr_amt_lak")))
                    Dim KKK As String = "INSERT INTO gen_jn( date_work, ac_Name, book, certify, Referno ,descrip ,descripe , " & _
                                          " amount , curr ,rate, Rate_USD, net_amt ,code_dr ,code_cr ,ac_code ,amt_dr , amt_cr , amt_USD_Dr, amt_USD_Cr , amount_dr ,amount_cr ,  " & _
                                          " certis, lock ,rec_lock , last_update , last_user ,company  ,Office_ID, del , AG, Frm, API ) " & _
                                          "Values('" & Format(CDate(bis_date), "yyyy/MM/dd") & "', " & _
                                          " N'" & trn_desc & "', " & _
                                          " N'" & acc_book & "', " & _
                                          " N'" & MdCertifyId & "', " & _
                                          " N'" & trn_id & "', " & _
                                          " N'" & trn_desc & "', " & _
                                          " N'', " & _
                                          " " & CDbl(SUMAMT) & ", " & _
                                           " N'" & Currency & "', " & _
                                          " " & CDbl(ex_rate) & ", " & _
                                          " " & CDbl(ex_rate) & ", " & _
                                          " '0', " & _
                                          " '', " & _
                                          " '" & cr_ac & "', " & _
                                          " '" & cr_ac & "', " & _
                                          " '', " & _
                                          " " & CDbl(cr_amt_lak) & ", " & _
                                          " '', " & _
                                          " " & CDbl(cr_amt) & ", " & _
                                          " " & CDbl(0) & ", " & _
                                          " " & CDbl(cr_amt) / CDbl(ex_rate) & ", " & _
                                          " '3', '4', '5', " & _
                                          " '" & Format(CDate(Date.Today), "yyyy/MM/dd") & "', " & _
                                          " '" & MUserID & "', " & _
                                          " '" & MuSubOff & "' , " & _
                                          " '" & MuSubOff & "', 0,1,0, 'API' )"
                    DbHelper.ExecuteNonQuery(KKK)

                Next

                Using cmdUp As New MySqlCommand("UPDATE msp SET status=@status WHERE trn_id=@trn_id", conn, tx)
                    cmdUp.Parameters.AddWithValue("@status", "success")
                    cmdUp.Parameters.AddWithValue("@trn_id", trn_id)
                    cmdUp.ExecuteNonQuery()
                End Using

            Next

            MsgBox("ບັນທຶກສໍາເລັດ")
            tx.Commit()
        Catch ex As Exception
            Try
                tx.Rollback()
            Catch
            End Try
            MessageBox.Show("Unexpected error: " & ex.Message)
        Finally
            conn.Close()
        End Try

    End Sub
   
    Public Sub AutoNumber()
        'Dim rs As New ADODB.Recordset
        Dim mNum As Integer
        Dim ss As String

        ss = "SELECT top 1 Right(certify,7) As certify FROM gen_jn WHERE book = N'" & acc_book & "' " & _
             "And year(date_work) = '" & Format(CDate(bis_date), "yyyy") & "' And month(date_work) = '" & Format(CDate(bis_date), "MM") & "' " & _
             "And LEFT(company,2)=N'" & Off_Id & "' Order by Right(certify,7) DESC"

        Dim dtTemp As DataTable = DbHelper.GetDataTable(ss)

        If dtTemp.Rows.Count = 0 Then
            MdCertifyId = Format(1, "0000000")
        Else
            Dim certifyValue As String = DbHelper.GetStr(dtTemp.Rows(0)("certify"))
            If certifyValue.Trim = "" Then
                mNum = 0
            Else
                mNum = Val(certifyValue)
            End If
            mNum = mNum + 1
            MdCertifyId = Format(mNum, "0000000")
        End If

        ' Full certify value: <book><yymm><7-digits>
        MdCertifyId = acc_book & Format(CDate(bis_date), "yyMM") & MdCertifyId

    End Sub
    Public Sub LoadData4()
        Try
            'Dim url As String = "http://10.151.146.91:8000/ping"
            'Dim url As String = "http://" & MDServerName2 & ":8000/ping"
            Dim url As String = "http://10.151.146.96:8000/ping"
            ' Using WebClient (simplest)
            Using client As New WebClient()
                Dim response As String = client.DownloadString(url)
                ' MessageBox.Show("Response from API: " & response, "API Response", MessageBoxButtons.OK, MessageBoxIcon.Information)

                'Dim connStr As String = "Server=" & MDServerName2 & ";Database=" & MDDatabaName2 & ";Uid=" & MDServerUser2 & ";Pwd=" & MDServerPassword2 & ";SslMode=None;"
                ''Dim connStr As String = "Server=apis.com.la;Database=apb_msp;Uid=Admin;Pwd=Sql_admin@#2024;Port=3306;SslMode=None;"

                Dim ConnectionString As String = "Data Source=" & MDServerName & ";Initial Catalog=" & MDDatabaName & ";User ID=" & MDServerUser & ";Password=" & MDServerPassword & ""

                ' Fetch Header Data from API
                Dim dtMsp As DataTable = ApiClient.GetMspData("wait")

                If dtMsp Is Nothing OrElse dtMsp.Rows.Count = 0 Then
                    Console.WriteLine("No pending data found.")
                    Exit Sub
                End If

                Using conn As New SqlConnection(ConnectionString)
                    conn.Open()

                    ' 1. Clear Old Staging Data
                    Dim cmdClear As New SqlCommand("DELETE FROM STAGE_MSP; DELETE FROM STAGE_TBL_DR; DELETE FROM STAGE_TBL_CR;", conn)
                    cmdClear.ExecuteNonQuery()

                    ' 2. Loop through each transaction
                    For Each row As DataRow In dtMsp.Rows
                        Dim trnId As String = row("trn_id").ToString() 
                        Dim isLocalInsertSuccess As Boolean = False

                        ' Insert Header to STAGE_MSP
                        Dim queryMsp As String = "INSERT INTO STAGE_MSP (trn_id, trn_desc, currency, acc_book, status, bis_date, create_date, ex_rate) " & _
                                                 "VALUES (@id, @desc, @curr, @book, @stat, @bis, @create, @rate)"
                        Using cmd As New SqlCommand(queryMsp, conn)
                            cmd.Parameters.AddWithValue("@id", trnId)
                            cmd.Parameters.AddWithValue("@desc", row("trn_desc"))
                            cmd.Parameters.AddWithValue("@curr", row("currency"))
                            cmd.Parameters.AddWithValue("@book", row("acc_book"))
                            cmd.Parameters.AddWithValue("@stat", row("status"))
                            'cmd.Parameters.AddWithValue("@bis", row("bis_date"))
                            'cmd.Parameters.AddWithValue("@create", row("create_date"))
                            'cmd.Parameters.AddWithValue("@rate", row("ex_rate")) 
                            cmd.Parameters.Add("@bis", SqlDbType.DateTime).Value = Convert.ToDateTime(row("bis_date"))
                            cmd.Parameters.Add("@create", SqlDbType.DateTime).Value = Convert.ToDateTime(row("create_date"))
                            cmd.Parameters.AddWithValue("@rate", row("ex_rate"))
                            cmd.ExecuteNonQuery()
                        End Using

                        ' Fetch & Insert Debits to STAGE_TBL_DR
                        Dim dtDr As DataTable = ApiClient.GetDetails("retrieve_dr_trn_id", trnId)
                        If dtDr IsNot Nothing Then
                            For Each drRow As DataRow In dtDr.Rows
                                Dim queryDr As String = "INSERT INTO STAGE_TBL_DR (trn_id, dr_ac, dr_amt, dr_amt_lak, dr_desc) VALUES (@id, @ac, @amt, @lak, @desc)"
                                Using cmdDr As New SqlCommand(queryDr, conn)
                                    cmdDr.Parameters.AddWithValue("@id", trnId)
                                    cmdDr.Parameters.AddWithValue("@ac", drRow("dr_ac"))
                                    cmdDr.Parameters.AddWithValue("@amt", drRow("dr_amt"))
                                    cmdDr.Parameters.AddWithValue("@lak", drRow("dr_amt_lak"))
                                    cmdDr.Parameters.AddWithValue("@desc", drRow("dr_desc"))
                                    cmdDr.ExecuteNonQuery()
                                End Using
                            Next
                        End If

                        ' Fetch & Insert Credits to STAGE_TBL_CR
                        Dim dtCr As DataTable = ApiClient.GetDetails("retrieve_cr_trn_id", trnId)
                        If dtCr IsNot Nothing Then
                            For Each crRow As DataRow In dtCr.Rows
                                Dim queryCr As String = "INSERT INTO STAGE_TBL_CR (trn_id, cr_ac, cr_amt, cr_amt_lak, cr_desc) VALUES (@id, @ac, @amt, @lak, @desc)"
                                Using cmdCr As New SqlCommand(queryCr, conn)
                                    cmdCr.Parameters.AddWithValue("@id", trnId)
                                    cmdCr.Parameters.AddWithValue("@ac", crRow("cr_ac"))
                                    cmdCr.Parameters.AddWithValue("@amt", crRow("cr_amt"))
                                    cmdCr.Parameters.AddWithValue("@lak", crRow("cr_amt_lak"))
                                    cmdCr.Parameters.AddWithValue("@desc", crRow("cr_desc"))
                                    cmdCr.ExecuteNonQuery()
                                End Using
                            Next
                        End If


                        ' 3. Update Status to 'success' via API
                        'If isLocalInsertSuccess Then
                        '    Dim apiUpdated As Boolean = ApiClient.UpdateStatus(trnId, "success")
                        'End If

                        Dim apiUpdated As Boolean = ApiClient.UpdateStatus(trnId, "success")

                    Next
                End Using

                Console.WriteLine("Staging Load Complete.")




            End Using

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "API Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub LoadDataSQL()
        'Dim rsAS As New ADODB.Recordset
        'Dim rsAS2 As New ADODB.Recordset
        'Dim rsAS3 As New ADODB.Recordset
        'Dim rsAS4 As New ADODB.Recordset
        Dim ss As String = " SELECT trn_id, trn_desc, acc_book, Currency, ex_rate, bis_date FROM STAGE_MSP WHERE status = 'Wait' ORDER BY trn_id  "
        Dim dtTemp As DataTable = DbHelper.GetDataTable(ss)
        For Each row As DataRow In dtTemp.Rows
            trn_id = DbHelper.GetStr(row("trn_id"))
            trn_desc = DbHelper.GetStr(row("trn_desc"))
            acc_book = DbHelper.GetStr(row("acc_book"))
            Currency = DbHelper.GetStr(row("Currency"))
            ex_rate = CDbl(DbHelper.GetStr(row("ex_rate")))
            If ex_rate = 0 Then ex_rate = 1
            bis_date = Format(CDate(DbHelper.GetStr(row("bis_date"))), "dd/MM/yyyy")
            Call AutoNumber()
            Dim ss2 As String = " SELECT SUM(dr_amt) as dr_amt  FROM STAGE_TBL_DR WHERE trn_id = N'" & trn_id & "'  "
            Dim dtTemp2 As DataTable = DbHelper.GetDataTable(ss2)
            SUMAMT = 0
            If dtTemp2.Rows.Count > 0 Then
                SUMAMT = CDbl(DbHelper.GetStr(dtTemp2.Rows(0)("dr_amt")))
            End If

            Dim ss3 As String = " SELECT *  FROM STAGE_TBL_DR WHERE trn_id = N'" & trn_id & "'  "
            Dim dtTemp3 As DataTable = DbHelper.GetDataTable(ss3)
            For Each row3 As DataRow In dtTemp3.Rows
                Dim dr_ac As String = DbHelper.GetStr(row3("dr_ac"))
                Dim dr_desc As String = DbHelper.GetStr(row3("dr_desc"))
                Dim dr_amt As Double = CDbl(DbHelper.GetStr(row3("dr_amt")))
                Dim dr_amt_lak As Double = CDbl(DbHelper.GetStr(row3("dr_amt_lak")))
                Dim KKK As String = "INSERT INTO gen_jn( date_work, ac_Name, book, certify, Referno ,descrip ,descripe , " & _
                                       " amount , curr ,rate, Rate_USD, net_amt ,code_dr ,code_cr ,ac_code ,amt_dr , amt_cr , amt_USD_Dr, amt_USD_Cr , amount_dr ,amount_cr ,  " & _
                                       " certis, lock ,rec_lock , last_update , last_user  ,company  ,Office_ID, del , AG, Frm, API ) " & _
                                       " Values('" & Format(CDate(bis_date), "yyyy/MM/dd") & "', " & _
                                       " N'" & dr_desc & "', " & _
                                       " N'" & acc_book & "', " & _
                                       " N'" & MdCertifyId & "', " & _
                                       " N'" & trn_id & "', " & _
                                       " N'" & dr_desc & "', " & _
                                       " N'', " & _
                                       " " & CDbl(SUMAMT) & ", " & _
                                       " N'" & Currency & "', " & _
                                       " " & CDbl(ex_rate) & ", " & _
                                       " " & CDbl(ex_rate) & ", " & _
                                       " 0, " & _
                                       " '" & dr_ac & "', " & _
                                       " '', " & _
                                       " '" & dr_ac & "', " & _
                                       " " & CDbl(dr_amt_lak) & ", " & _
                                       " 0, " & _
                                       " " & CDbl(dr_amt) & ", " & _
                                       " 0, " & _
                                       " " & CDbl(dr_amt) / CDbl(ex_rate) & ", " & _
                                       " " & CDbl(0) & ", " & _
                                       " 3, 4, 5, " & _
                                       " '" & Format(CDate(Date.Today), "yyyy/MM/dd") & "', " & _
                                       " '" & MUserID & "', " & _
                                       " '" & MuSubOff & "' , " & _
                                       " '" & MuSubOff & "', 0,1,0, 'API' )"
                DbHelper.ExecuteNonQuery(KKK)
            Next

            Dim ss4 As String = " SELECT * FROM STAGE_TBL_CR WHERE trn_id = N'" & trn_id & "'  "
            Dim dtTemp4 As DataTable = DbHelper.GetDataTable(ss4)
            For Each row4 As DataRow In dtTemp4.Rows
                Dim cr_ac As String = DbHelper.GetStr(row4("cr_ac"))
                Dim cr_desc As String = DbHelper.GetStr(row4("cr_desc"))
                Dim cr_amt As Double = CDbl(DbHelper.GetStr(row4("cr_amt")))
                Dim cr_amt_lak As Double = CDbl(DbHelper.GetStr(row4("cr_amt_lak")))
                Dim KKK3 As String = "INSERT INTO gen_jn( date_work, ac_Name, book, certify, Referno ,descrip ,descripe , " & _
                                       " amount , curr ,rate, Rate_USD, net_amt ,code_dr ,code_cr ,ac_code ,amt_dr , amt_cr , amt_USD_Dr, amt_USD_Cr , amount_dr ,amount_cr ,  " & _
                                       " certis, lock ,rec_lock , last_update , last_user ,company  ,Office_ID, del , AG, Frm, API ) " & _
                                       "Values('" & Format(CDate(bis_date), "yyyy/MM/dd") & "', " & _
                                       " N'" & cr_desc & "', " & _
                                       " N'" & acc_book & "', " & _
                                       " N'" & MdCertifyId & "', " & _
                                       " N'" & trn_id & "', " & _
                                       " N'" & cr_desc & "', " & _
                                       " N'', " & _
                                       " " & CDbl(SUMAMT) & ", " & _
                                       " N'" & Currency & "', " & _
                                       " " & CDbl(ex_rate) & ", " & _
                                       " " & CDbl(ex_rate) & ", " & _
                                       " '0', " & _
                                       " '', " & _
                                       " '" & cr_ac & "', " & _
                                       " '" & cr_ac & "', " & _
                                       " '', " & _
                                       " " & CDbl(cr_amt_lak) & ", " & _
                                       " '', " & _
                                       " " & CDbl(cr_amt) & ", " & _
                                       " " & CDbl(0) & ", " & _
                                       " " & CDbl(cr_amt) / CDbl(ex_rate) & ", " & _
                                       " '3', '4', '5', " & _
                                       " '" & Format(CDate(Date.Today), "yyyy/MM/dd") & "', " & _
                                       " '" & MUserID & "', " & _
                                       " '" & MuSubOff & "' , " & _
                                       " '" & MuSubOff & "', 0,1,0, 'API' )"
                DbHelper.ExecuteNonQuery(KKK3)
            Next
        Next
                    bis_date = Format(rsAS.Fields("bis_date").Value, "dd/MM/yyyy")
                    Call AutoNumber()
                    Dim ss2 As String = " SELECT SUM(dr_amt) as dr_amt  FROM STAGE_TBL_DR WHERE trn_id = N'" & trn_id & "'  "
                    Dim dtTemp As DataTable = DbHelper.GetDataTable(ss2)
                    If rsAS2.RecordCount > 0 Then
                        SUMAMT = Format(rsAS2.Fields("dr_amt").Value, "#,##0.00")
                    End If

                    Dim ss3 As String = " SELECT *  FROM STAGE_TBL_DR WHERE trn_id = N'" & trn_id & "'  "
                    Dim dtTemp As DataTable = DbHelper.GetDataTable(ss3)
                    If rsAS3.RecordCount > 0 Then
                        While Not rsAS3.EOF()

                            Dim dr_ac As String = Trim(rsAS3.Fields("dr_ac").Value.ToString)
                            Dim dr_desc As String = Trim(rsAS3.Fields("dr_desc").Value.ToString)
                            Dim dr_amt As Double = Format(rsAS3.Fields("dr_amt").Value, "#,##0.00")
                            Dim dr_amt_lak As Double = Format(rsAS3.Fields("dr_amt_lak").Value, "#,##0.00")
                            'If CDbl(dr_amt) = "" Then
                            '    dr_amt = 0
                            'End If
                            'If dr_amt_lak = "" Then
                            '    dr_amt_lak = 0
                            'End If
                            Dim KKK As String = "INSERT INTO gen_jn( date_work, ac_Name, book, certify, Referno ,descrip ,descripe , " & _
                                                  " amount , curr ,rate, Rate_USD, net_amt ,code_dr ,code_cr ,ac_code ,amt_dr , amt_cr , amt_USD_Dr, amt_USD_Cr , amount_dr ,amount_cr ,  " & _
                                                  " certis, lock ,rec_lock , last_update , last_user  ,company  ,Office_ID, del , AG, Frm, API ) " & _
                                                  " Values('" & Format(CDate(bis_date), "yyyy/MM/dd") & "', " & _
                                                  " N'" & dr_desc & "', " & _
                                                  " N'" & acc_book & "', " & _
                                                  " N'" & MdCertifyId & "', " & _
                                                  " N'" & trn_id & "', " & _
                                                  " N'" & dr_desc & "', " & _
                                                  " N'', " & _
                                                  " " & CDbl(SUMAMT) & ", " & _
                                                  " N'" & Currency & "', " & _
                                                  " " & CDbl(ex_rate) & ", " & _
                                                  " " & CDbl(ex_rate) & ", " & _
                                                  " 0, " & _
                                                  " '" & dr_ac & "', " & _
                                                  " '', " & _
                                                  " '" & dr_ac & "', " & _
                                                  " " & CDbl(dr_amt_lak) & ", " & _
                                                  " 0, " & _
                                                  " " & CDbl(dr_amt) & ", " & _
                                                  " 0, " & _
                                                  " " & CDbl(dr_amt) / CDbl(ex_rate) & ", " & _
                                                  " " & CDbl(0) & ", " & _
                                                  " 3, 4, 5, " & _
                                                  " '" & Format(CDate(Date.Today), "yyyy/MM/dd") & "', " & _
                                                  " '" & MUserID & "', " & _
                                                  " '" & MuSubOff & "' , " & _
                                                  " '" & MuSubOff & "', 0,1,0, 'API' )"
                            DbHelper.ExecuteNonQuery(KKK)
                            rsAS3.MoveNext()
                        End While

                    End If



                    Dim ss4 As String = " SELECT * FROM STAGE_TBL_cR WHERE trn_id = N'" & trn_id & "'  "
                    Dim dtTemp As DataTable = DbHelper.GetDataTable(ss4)
                    If rsAS4.RecordCount > 0 Then
                        While Not rsAS4.EOF()
                            Dim cr_ac As String = Trim(rsAS4.Fields("cr_ac").Value.ToString)
                            Dim cr_desc As String = Trim(rsAS4.Fields("cr_desc").Value.ToString)
                            Dim cr_amt As Double = Format(rsAS4.Fields("cr_amt").Value, "#,##0.00")
                            Dim cr_amt_lak As Double = Format(rsAS4.Fields("cr_amt_lak").Value, "#,##0.00")

                            Dim KKK3 As String = "INSERT INTO gen_jn( date_work, ac_Name, book, certify, Referno ,descrip ,descripe , " & _
                                                  " amount , curr ,rate, Rate_USD, net_amt ,code_dr ,code_cr ,ac_code ,amt_dr , amt_cr , amt_USD_Dr, amt_USD_Cr , amount_dr ,amount_cr ,  " & _
                                                  " certis, lock ,rec_lock , last_update , last_user ,company  ,Office_ID, del , AG, Frm, API ) " & _
                                                  "Values('" & Format(CDate(bis_date), "yyyy/MM/dd") & "', " & _
                                                  " N'" & cr_desc & "', " & _
                                                  " N'" & acc_book & "', " & _
                                                  " N'" & MdCertifyId & "', " & _
                                                  " N'" & trn_id & "', " & _
                                                  " N'" & cr_desc & "', " & _
                                                  " N'', " & _
                                                  " " & CDbl(SUMAMT) & ", " & _
                                                   " N'" & Currency & "', " & _
                                                  " " & CDbl(ex_rate) & ", " & _
                                                  " " & CDbl(ex_rate) & ", " & _
                                                  " '0', " & _
                                                  " '', " & _
                                                  " '" & cr_ac & "', " & _
                                                  " '" & cr_ac & "', " & _
                                                  " '', " & _
                                                  " " & CDbl(cr_amt_lak) & ", " & _
                                                  " '', " & _
                                                  " " & CDbl(cr_amt) & ", " & _
                                                  " " & CDbl(0) & ", " & _
                                                  " " & CDbl(cr_amt) / CDbl(ex_rate) & ", " & _
                                                  " '3', '4', '5', " & _
                                                  " '" & Format(CDate(Date.Today), "yyyy/MM/dd") & "', " & _
                                                  " '" & MUserID & "', " & _
                                                  " '" & MuSubOff & "' , " & _
                                                  " '" & MuSubOff & "', 0,1,0, 'API' )"
                            DbHelper.ExecuteNonQuery(KKK3)
                            rsAS4.MoveNext()
                        End While
                    End If


             

                    .MoveNext()
                End While
            End With



        End If

         

        MsgBox("MSP ລົງບັນຊີສໍາເລັດ")
       
    End Sub
    Public Sub LoadDataCancel()
        Try
            ' Dim url As String = "http://10.151.146.91:8000/ping"
            'Dim url As String = "http://" & MDServerName2 & ":8000/ping"
            Dim url As String = "http://10.151.146.96:8000/ping"
            ' Using WebClient (simplest)
            Using client As New WebClient()
                Dim response As String = client.DownloadString(url)
      
                Dim ConnectionString As String = "Data Source=" & MDServerName & ";Initial Catalog=" & MDDatabaName & ";User ID=" & MDServerUser & ";Password=" & MDServerPassword & ""

                ' Fetch Header Data from API
                Dim dtMsp As DataTable = ApiClient.GetMspData("cancel")

                If dtMsp Is Nothing OrElse dtMsp.Rows.Count = 0 Then
                    Console.WriteLine("No pending data found.")
                    Exit Sub
                End If

                Using conn As New SqlConnection(ConnectionString)
                    conn.Open()

                    ' 1. Clear Old Staging Data
                    Dim cmdClear As New SqlCommand("DELETE FROM STAGE_MSP; DELETE FROM STAGE_TBL_DR; DELETE FROM STAGE_TBL_CR;", conn)
                    cmdClear.ExecuteNonQuery()

                    ' 2. Loop through each transaction
                    For Each row As DataRow In dtMsp.Rows
                        Dim trnId As String = row("trn_id").ToString()
                        Dim isLocalInsertSuccess As Boolean = False

                        ' Insert Header to STAGE_MSP
                        Dim queryMsp As String = "INSERT INTO STAGE_MSP (trn_id, trn_desc, currency, acc_book, status, bis_date, create_date, ex_rate) " & _
                                                 "VALUES (@id, @desc, @curr, @book, @stat, @bis, @create, @rate)"
                        Using cmd As New SqlCommand(queryMsp, conn)
                            cmd.Parameters.AddWithValue("@id", trnId)
                            cmd.Parameters.AddWithValue("@desc", row("trn_desc"))
                            cmd.Parameters.AddWithValue("@curr", row("currency"))
                            cmd.Parameters.AddWithValue("@book", row("acc_book"))
                            cmd.Parameters.AddWithValue("@stat", row("status"))
                            'cmd.Parameters.AddWithValue("@bis", row("bis_date"))
                            'cmd.Parameters.AddWithValue("@create", row("create_date"))
                            'cmd.Parameters.AddWithValue("@rate", row("ex_rate")) 
                            cmd.Parameters.Add("@bis", SqlDbType.DateTime).Value = Convert.ToDateTime(row("bis_date"))
                            cmd.Parameters.Add("@create", SqlDbType.DateTime).Value = Convert.ToDateTime(row("create_date"))
                            cmd.Parameters.AddWithValue("@rate", row("ex_rate"))
                            cmd.ExecuteNonQuery()
                        End Using
 

                        Dim apiUpdated As Boolean = ApiClient.UpdateStatus(trnId, "canceled")

                    Next
                End Using

                Console.WriteLine("Staging Load Complete.")




            End Using

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "API Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub LoadDataSQLDelete()
        'Dim rsAS As New ADODB.Recordset
        'Dim rsAS2 As New ADODB.Recordset
        'Dim rsAS3 As New ADODB.Recordset
        'Dim rsAS4 As New ADODB.Recordset
        Dim ss As String = " SELECT trn_id, trn_desc, acc_book, Currency, ex_rate, bis_date FROM STAGE_MSP WHERE status = 'cancel' ORDER BY trn_id  "
        Dim dtTemp As DataTable = DbHelper.GetDataTable(ss)
        For Each row As DataRow In dtTemp.Rows
            trn_id = DbHelper.GetStr(row("trn_id"))
            DbHelper.ExecuteNonQuery(" Delete gen_jn where API='API'  and Referno=N'" & trn_id & "'  ")
        Next



        MsgBox("MSP ຍົກເລີກສໍາເລັດ")

    End Sub
    Private Sub Button22_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        LoadDataCancel()
        LoadDataSQLDelete()
    End Sub
End Class