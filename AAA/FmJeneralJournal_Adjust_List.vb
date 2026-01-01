Public Class FmJeneralJournal_Adjust_List
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

    Private Function GetString(ByVal cellVal As Object) As String
        If cellVal Is Nothing Then Return ""
        Return cellVal.ToString()
    End Function

    Private Sub BntNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BntNew.Click
        Panel4.Visible = False
        FmNsewJeneralJournal_Adjust.txtInvoice.Enabled = True
        FmNsewJeneralJournal_Adjust.CmbBook.Enabled = True
        'FmNsewJeneralJournal.MdiParent = Me
        'FmNsewJeneralJournal.WindowState = FormWindowState.Maximized

        FmNsewJeneralJournal_Adjust.ShowDialog()
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
        End If
        If RadioButton2.Checked = True Then
            LngId = "7014" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (02), " & Y & " " & Format(dts.Value, "yyyy")
        End If
        If RadioButton3.Checked = True Then
            LngId = "7015" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (03), " & Y & " " & Format(dts.Value, "yyyy")
        End If
        If RadioButton4.Checked = True Then
            LngId = "7016" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (04), " & Y & " " & Format(dts.Value, "yyyy")
        End If
        If RadioButton5.Checked = True Then
            LngId = "7017" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (05), " & Y & " " & Format(dts.Value, "yyyy")
        End If
        If RadioButton6.Checked = True Then
            LngId = "7018" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (06), " & Y & " " & Format(dts.Value, "yyyy")
        End If
        If RadioButton7.Checked = True Then
            LngId = "7019" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (07), " & Y & " " & Format(dts.Value, "yyyy")
        End If
        If RadioButton8.Checked = True Then
            LngId = "7020" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (08), " & Y & " " & Format(dts.Value, "yyyy")
        End If
        If RadioButton9.Checked = True Then
            LngId = "7021" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (09), " & Y & " " & Format(dts.Value, "yyyy")
        End If
        If RadioButton10.Checked = True Then
            LngId = "7022" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (10), " & Y & " " & Format(dts.Value, "yyyy")
        End If
        If RadioButton11.Checked = True Then
            LngId = "7023" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (11), " & Y & " " & Format(dts.Value, "yyyy")
        End If
        If RadioButton12.Checked = True Then
            LngId = "7024" : CallLngStr() : D = LngStr
            LngId = "7025" : CallLngStr() : Y = LngStr
            RptName = D & " (12), " & Y & " " & Format(dts.Value, "yyyy")
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
            SR = " AND AP_ACC_adjust_Item.certify = '" & Nme.Text & "' "

        End If
        If RCex.Checked = True Then
            SR = " AND AP_ACC_adjust_Item.cheque_no = '" & Nme.Text & "' "
        End If

        If RAc_code.Checked = True Then
            SR = " AND (AP_ACC_adjust_Item.ac_code  Like N'" & Nme.Text.Trim & "%')"
        End If
        If RAcNme.Checked = True Then
            SR = " AND (AP_ACC_adjust_Item.ac_name  Like N'%" & Nme.Text.Trim & "%')"
        End If
        If RDesc.Checked = True Then
            SR = " AND (AP_ACC_adjust_Item.descrip  Like N'%" & Nme.Text.Trim & "%')"
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
            SR = " AND AP_ACC_adjust_Item.curr = '" & Nme.Text & "' "
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
        LoadSqlData("select sub_id , off_add2  from  Ap_office  Order by sub_id", RSC)
        With RSC
            Do Until .EOF = True
                Off_Usr.Items.Add((.Fields("sub_id").Value) & " " & (.Fields("off_add2").Value))
                .MoveNext()
            Loop
        End With
        Off_Usr.Text = FmLogin.Sub_Company.Text
    End Sub

    Private Sub FmJeneralJournal_List_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MuSubOff = MuSubOff2
    End Sub
    Private Sub loadCompany()

        CmbCompany.Items.Clear()
        LoadSqlData("select off_add1 , off_id  from  Ap_office group BY off_id , off_add1", RSC)
        With RSC
            Do Until .EOF = True
                CmbCompany.Items.Add((.Fields("off_id").Value) & " " & (.Fields("off_add1").Value))
                .MoveNext()
            Loop
        End With
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
                CntNB = "date_work ASC, cnt ASC"
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


    Public Sub SetupGrid()
        FG.Columns.Clear()
        FG.AllowUserToAddRows = False
        FG.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG.ReadOnly = True
        
        FG.Columns.Add("AbsolutePosition", "ລ/ດ") ' 0
        FG.Columns.Add("date_work", "ວັນທີ")   ' 1
        FG.Columns.Add("certify", "ໃບຍັງຢືນ")  ' 2
        FG.Columns.Add("Referno", "ແຊັກເລກທີ") ' 3
        FG.Columns.Add("Ac_code", "ເລກບັນຊີ") ' 4
        FG.Columns.Add("amount", "ຈຳນວນເງິນ") ' 5
        FG.Columns.Add("Ac_name", "ເນື້ອໃນລາຍການ") ' 6
        FG.Columns.Add("amount_dr", "ຈຳນວນເງິນຈົດໜີ້") ' 7
        FG.Columns.Add("amount_cr", "ຈຳນວນເງິນຈົດມີ") ' 8
        FG.Columns.Add("curr", "ສະກຸນເງິນ") ' 9
        FG.Columns.Add("amt_dr", "ຈຳນວນເງິນຈົດໜີ້(ກີບ)") ' 10
        FG.Columns.Add("amt_cr", "ຈຳນວນເງິນຈົດມີ(ກີບ)") ' 11
        FG.Columns.Add("company", "ຕົ້ນທຶນ") ' 12
        FG.Columns.Add("cnt", "cnt") ' 13
        FG.Columns.Add("lock", "lock") ' 14
        FG.Columns.Add("blank", "") ' 15
        FG.Columns.Add("book", "book") ' 16

        FG.Columns("cnt").Visible = False
        FG.Columns("lock").Visible = False
        FG.Columns("blank").Visible = False
        FG.Columns("book").Visible = False
        
        FG.ColumnHeadersDefaultCellStyle.Font = New Font("Lao_Classic3", 10, FontStyle.Bold)
    End Sub

    Private Sub FmJeneralJournal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CNN.Execute("update AP_ACC_adjust_Item set AP_ACC_adjust_Item.ac_name=Acc_Code.name_L,AP_ACC_adjust_Item.ac_namee=Acc_Code.name_E from Acc_Code,AP_ACC_adjust_Item where AP_ACC_adjust_Item.Ac_Code=Acc_Code.Ac_Code and  AP_ACC_adjust_Item.ac_name is null ")
        certify.Checked = True
        MDESC.Checked = False
        FG.BackgroundColor = Color.White
        SetControlText(Me)
        Call loadCompany()
        LoadSubCompany()
        Off_Usr.Text = FmLogin.Sub_Company.Text
        FG.GridColor = Color.RoyalBlue
        
        SetupGrid()
        
        RAll.Checked = True
        ComboBox1.Items.Clear()
        Nme.Enabled = False
        Nme.Visible = True
        
        MDInvoiceNo = ""
    End Sub
    Private Sub LoadBooks()
        Dim rst As New ADODB.Recordset
        ComboBox1.Items.Clear()
        Comm = New ADODB.Command
        Comm.ActiveConnection = CNN
        Comm.CommandText = "SELECT * FROM books WHERE bookid <> N'" & "" & "'"
        rst = Comm.Execute

        If rst.RecordCount <> 0 Then
            ComboBox1.Items.Add("<All>  ທັງໝົດ (All books)")
            While Not rst.EOF()
                ComboBox1.Items.Add(Trim(rst.Fields("bookid").Value))
                rst.MoveNext()
            End While
            ComboBox1.Text = "<All>  ທັງໝົດ (All books)"
        End If
    End Sub
    Private Sub LoadsLurr()
        Dim rst As New ADODB.Recordset
        ComboBox1.Items.Clear()
        Comm = New ADODB.Command
        Comm.ActiveConnection = CNN
        Comm.CommandText = "SELECT curr FROM Ap_RateSeting WHERE curr <> N'" & "" & "'"
        rst = Comm.Execute

        If rst.RecordCount <> 0 Then
            ComboBox1.Items.Add("<All>  ທັງໝົດ")
            While Not rst.EOF()
                ComboBox1.Items.Add(Trim(rst.Fields("curr").Value))
                rst.MoveNext()
            End While
            ComboBox1.Text = "<All>  ທັງໝົດ"
        End If
    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtBook.Text = ""
        LoadSqlData("SELECT * FROM books WHERE bookname  = N'" & ComboBox1.Text & "'", RSC)
        With RSC
            Do Until .EOF = True
                txtBook.Text = Trim(.Fields("bookid").Value)
                .MoveNext()
            Loop
        End With
        If ComboBox1.Text = "" Then
            txtBook.Text = "All"
        End If
    End Sub

    Private Sub Load_AP_ACC_adjust_Item()
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
        If FG.CurrentRow IsNot Nothing Then
            txtAc_Code.Text = If(FG.CurrentRow.Cells(4).Value Is Nothing, "", FG.CurrentRow.Cells(4).Value.ToString())
            txtCurr.Text = If(FG.CurrentRow.Cells(9).Value Is Nothing, "", FG.CurrentRow.Cells(9).Value.ToString())
            Dim str As String = ", Amount = " & If(FG.CurrentRow.Cells(4).Value Is Nothing, "", FG.CurrentRow.Cells(4).Value.ToString()) & ": " & If(FG.CurrentRow.Cells(9).Value Is Nothing, "", FG.CurrentRow.Cells(9).Value.ToString())
            If MuLng = "L" Then
                LoadSqlData("SELECT AG  , descrip     FROM AP_ACC_adjust_Item WHERE cnt = '" & If(FG.CurrentRow.Cells(13).Value Is Nothing, "", FG.CurrentRow.Cells(13).Value.ToString()) & "' " & MULook2 & "  order by cnt", RSC)
                If RSC.RecordCount <> 0 Then
                    AG = Trim(RSC.Fields("AG").Value)
                    txtDescrip.Text = Trim(RSC.Fields("descrip").Value) & ", ມູນຄ່າ: " & If(FG.CurrentRow.Cells(5).Value Is Nothing, "", FG.CurrentRow.Cells(5).Value.ToString()) & ": " & If(FG.CurrentRow.Cells(9).Value Is Nothing, "", FG.CurrentRow.Cells(9).Value.ToString())
                End If
            LoadSqlData("SELECT   Name_L  FROM Acc_Code WHERE Ac_Code = '" & txtAc_Code.Text & "'", RSC)
            If RSC.RecordCount <> 0 Then

                Ac_Name.Text = Trim(RSC.Fields("Name_L").Value)
                'x = Ac_Name.Text
                'MsgBox(x)
            End If

            '========================
            End If ' Close If MuLng = "L" Then
        Else
            LoadSqlData("SELECT AG  , descripe    FROM AP_ACC_adjust_Item WHERE cnt = '" & If(FG.CurrentRow.Cells(13).Value Is Nothing, "", FG.CurrentRow.Cells(13).Value.ToString()) & "' " & MULook2 & "  order by cnt", RSC)
            If RSC.RecordCount <> 0 Then
                AG = Trim(RSC.Fields("AG").Value)
                txtDescrip.Text = Trim(RSC.Fields("descripe").Value) & ", Amount: " & If(FG.CurrentRow.Cells(5).Value Is Nothing, "", FG.CurrentRow.Cells(5).Value.ToString()) & ": " & If(FG.CurrentRow.Cells(9).Value Is Nothing, "", FG.CurrentRow.Cells(9).Value.ToString())
            End If
            LoadSqlData("SELECT   Name_E FROM Acc_Code WHERE Ac_Code = '" & txtAc_Code.Text & "'", RSC)
            If RSC.RecordCount <> 0 Then
                Ac_Name.Text = Trim(RSC.Fields("Name_E").Value)
            End If
        End If

        Dim RSC1 As New ADODB.Recordset
        Dim s As String = "SELECT sum(Amt_dr) as Amt_dr   , sum(Amt_cr) as Amt_cr    FROM AP_ACC_adjust_Item WHERE ac_code = '" & txtAc_Code.Text & "' and AP_ACC_adjust_Item.date_work BETWEEN '" & Format(dts.Value, "yyyy-MM-dd") & "' AND '" & Format(dtt.Value, "yyyy-MM-dd") & "'   " & MULook2 & "  "
        LoadSqlData(s, RSC1)
        If RSC1.RecordCount <> 0 Then
            SumDr.Text = Format(CDbl(Trim(RSC1.Fields("Amt_dr").Value)), "#,##0.00")
            SumCr.Text = Format(CDbl(Trim(RSC1.Fields("Amt_cr").Value)), "#,##0.00")
        End If

        LoadSqlData("select  amount_dr , amount_cr from Open_jn where ac_code='" & txtAc_Code.Text & "'   and  year(Date_work)= '" & Format(CDate(dts.Value), "yyyy") & "'  " & MULook2 & "   ", RSC)
        Op = 0
        If RSC.RecordCount <> 0 Then
            Op = CDbl(Trim(RSC.Fields("amount_dr").Value)) - CDbl(Trim(RSC.Fields("amount_cr").Value))
        End If
        Dim dss As Date
        dss = DateAdd(DateInterval.Day, -1, dts.Value)
        Dim RSC2 As New ADODB.Recordset
        LoadSqlData("select SUM(amount_dr) AS amount_dr ,SUM(amount_cr) AS amount_cr from AP_ACC_adjust_Item where ac_code=N'" & txtAc_Code.Text & "'  And AP_ACC_adjust_Item.date_work   BETWEEN '" & "1-1-" & Format(dts.Value, "yyyy") & "' AND '" & Format(dss, "yyyy-MM-dd") & "' " & MULook2 & " group by ac_code ", RSC2)
        If RSC2.RecordCount <> 0 Then
            Op = Op + CDbl(CDbl(Trim(RSC2.Fields("amount_dr").Value)) - CDbl(Trim(RSC2.Fields("amount_Cr").Value)))
        End If

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
        Dim s As String = "SELECT sum(Amt_dr) as Amt_dr , sum(Amt_cr) as Amt_cr   FROM AP_ACC_adjust_Item WHERE AP_ACC_adjust_Item.date_work BETWEEN '" & Format(dts.Value, "yyyy-MM-dd") & "' AND '" & Format(dtt.Value, "yyyy-MM-dd") & "'   " & MULook2 & "  "
        LoadSqlData(s, RSC)

        If RSC.RecordCount <> 0 Then
            TotalDr.Text = (Trim(RSC.Fields("Amt_cr").Value.ToString))
            TotalCr.Text = (Trim(RSC.Fields("Amt_cr").Value.ToString))
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
        '    SQL = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & Month(dts.Value) & "' AND '" & Month(dtt.Value) & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "' "
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
        LockData = GetString(FG.CurrentRow.Cells(13).Value)
        If GetString(FG.CurrentRow.Cells(14).Value) = 1 Then
            Button1.Text = "ປົດລອກ"
        End If
        If GetString(FG.CurrentRow.Cells(14).Value) = 0 Then
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
        Panel4.Visible = False
        If MDInvoiceNo <> "" Then
            FmNsewJeneralJournal_Adjust.txtInvoice.Enabled = False
            FmNsewJeneralJournal_Adjust.CmbBook.Enabled = False
            FmNsewJeneralJournal_Adjust.ShowDialog()
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
            Dim dateVal As String = If(FG.CurrentRow.Cells(1).Value Is Nothing, "1900-01-01", FG.CurrentRow.Cells(1).Value.ToString())
            CNN.Execute("delete AP_ACC_adjust_Item where certify =N'" & MDInvoiceNo & "' And   date_work='" & Format(CDate(dateVal), "yyyy-MM-dd") & "' ")
        End If
        MDInvoiceNo = ""
        LoadMonthSQL()
    End Sub

    Private Sub FG_AfterRowColChange(ByVal sender As Object, ByVal e As EventArgs) Handles FG.SelectionChanged
        If FG.Rows.Count > 7 Then
            LngId = "8008" : CallLngStr()
            'FG.FormatString = LngStr ' TODO: Convert FormatString logic to DGV header text
        Else
            LngId = "8001" : CallLngStr()
            'FG.FormatString = LngStr ' TODO: Convert FormatString logic to DGV header text
        End If
    End Sub

    Private Sub FG_Scroll(ByVal sender As Object, ByVal e As ScrollEventArgs) Handles FG.Scroll
        BtnSearch.Visible = False
    End Sub

    Private Sub FG_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FG.CellClick
        If e.RowIndex < 0 Then Exit Sub
        
        Dim row As DataGridViewRow = FG.Rows(e.RowIndex)
        MDInvoiceDT = If(row.Cells(1).Value Is Nothing, "", row.Cells(1).Value.ToString())
        MDInvoiceNo = If(row.Cells(2).Value Is Nothing, "", row.Cells(2).Value.ToString())
        
        If x > 0 Then
            x0 = x
            s0 = TextBox1.Text
        End If
        If y > 0 Then
            y0 = y
        End If
        x = e.ColumnIndex
        y = e.RowIndex
        
        TextBox1.Text = MDInvoiceNo
        If MuLng = "L" Then
            txtDescrip.Text = If(row.Cells(6).Value Is Nothing, "", row.Cells(6).Value.ToString()) & ", ມູນຄ່າ: " & If(row.Cells(5).Value Is Nothing, "", row.Cells(5).Value.ToString()) & ": " & If(row.Cells(9).Value Is Nothing, "", row.Cells(9).Value.ToString())
        Else
            txtDescrip.Text = If(row.Cells(6).Value Is Nothing, "", row.Cells(6).Value.ToString()) & ", Amout: " & If(row.Cells(5).Value Is Nothing, "", row.Cells(5).Value.ToString()) & ": " & If(row.Cells(9).Value Is Nothing, "", row.Cells(9).Value.ToString())
        End If

        LockData = ""
        LockData = If(row.Cells(14).Value Is Nothing, "", row.Cells(14).Value.ToString())
        If LockData = "1" Then
            RadioButton17.Checked = False
            RadioButton18.Checked = True
        Else
            RadioButton17.Checked = True
            RadioButton18.Checked = False
        End If
        
        If LockData = "1" Or LockData = "2" Then
            LngId = "3027" : CallLngStr() : Button1.Text = LngStr
        Else
            LngId = "3008" : CallLngStr() : Button1.Text = LngStr
        End If

        If e.ColumnIndex = 4 Then
            Call Load_AP_ACC_adjust_Item()
            Dim openJnVal As Double = 0
            Dim sumDrVal As Double = 0
            Dim sumCrVal As Double = 0
            Double.TryParse(Open_jn.Text.Replace("(", "-").Replace(")", "").Replace(",", ""), openJnVal)
            Double.TryParse(SumDr.Text.Replace(",", ""), sumDrVal)
            Double.TryParse(SumCr.Text.Replace(",", ""), sumCrVal)
            
            Dim remVal As Double = (openJnVal + sumDrVal) - sumCrVal
            If remVal >= 0 Then
                Remain.ForeColor = Color.Black
                Remain.Text = Format(remVal, "##,##0.00")
            Else
                Remain.ForeColor = Color.Red
                Remain.Text = "(" & Format(Math.Abs(remVal), "##,##0.00") & ")"
            End If
        End If

        If CheckBox3.Checked = True Then
            ' Simplified highlighting - in DataGridView we can just loop and set Style
            For Each r As DataGridViewRow In FG.Rows
                Dim rCert As String = If(r.Cells(2).Value Is Nothing, "", r.Cells(2).Value.ToString())
                If rCert = s0 Then
                    ' Reset old selection
                    For Each cell As DataGridViewCell In r.Cells
                        cell.Style.BackColor = Color.Empty
                    Next
                End If
                If rCert = TextBox1.Text Then
                    ' Highlight new selection
                    r.Cells(2).Style.BackColor = Color.SkyBlue
                    r.Cells(4).Style.BackColor = Color.SkyBlue
                    If CDbl(If(r.Cells(10).Value Is Nothing, 0, r.Cells(10).Value)) <> 0 Then
                        r.Cells(7).Style.BackColor = Color.SkyBlue
                        r.Cells(10).Style.BackColor = Color.SkyBlue
                    Else
                        r.Cells(8).Style.BackColor = Color.SkyBlue
                        r.Cells(11).Style.BackColor = Color.SkyBlue
                    End If
                End If
            Next
        End If

        CMS2.Enabled = True
        CMS3.Enabled = True
    End Sub

    Private Sub FG_CellDoubleClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FG.CellDoubleClick
        If e.RowIndex < 0 Then Exit Sub
        If MDInvoiceNo <> "" Then
            FmNsewJeneralJournal_Adjust.txtInvoice.Enabled = False
            FmNsewJeneralJournal_Adjust.CmbBook.Enabled = False
            FmNsewJeneralJournal_Adjust.ShowDialog()
        End If
    End Sub


    ' MouseUp logic merged into CellClick

    Private Sub ForIs2()

        If y0 = FG.Rows.Count - 1 Then
            If GetString(FG.Rows(FG.Rows.Count - 1).Cells(2).Value) <> GetString(FG.Rows(FG.Rows.Count - 2).Cells(2).Value) Then
                Rs2 = FG.Rows.Count - 1
            End If
        End If
        Dim x0k As Integer = y0
        For i = y0 To y0 * 2
            Rs2 = x0k

            x0k = x0k - 1
            s = GetString(FG.Rows(x0k).Cells(2).Value)
            'R1 = x + 1
            If s <> s0 Then
                Exit Sub
            End If
        Next
    End Sub

    Private Sub ForIt2()
        If y0 = FG.Rows.Count - 1 Then
            Rt2 = FG.Rows.Count - 1
            Exit Sub
        End If
        For i = y0 To FG.Rows.Count - 1
            Rt2 = i - 1
            s = GetString(FG.Rows(i).Cells(2).Value)
            If s <> s0 Then
                Exit Sub
            End If
        Next
    End Sub
    Private Sub ForIs1()
        If y = FG.Rows.Count - 1 Then
            Dim valCurrent As String = GetString(FG.Rows(FG.Rows.Count - 1).Cells(2).Value)
            Dim valPrev As String = GetString(FG.Rows(FG.Rows.Count - 2).Cells(2).Value)
            If valCurrent <> valPrev Then
                Rs1 = FG.Rows.Count - 1
            End If
        End If
        Dim x As Integer = y
        For i = y To y * 2
            Rs1 = x
            x = x - 1
            If x < 0 Then Exit For
            s = GetString(FG.Rows(x).Cells(2).Value)
            'R1 = x + 1
            If s <> TextBox1.Text Then
                Exit Sub
            End If
        Next
    End Sub

    Private Sub ForIt1()




        For i = y To FG.Rows.Count - 1
            Rt1 = i - 1
            s = GetString(FG.Rows(i).Cells(2).Value)
            If s <> TextBox1.Text Then
                Exit Sub
            End If
        Next
    End Sub
    ' ForIt1 replacement logic assumed similar legacy loop fix
    Private Sub FG_SelectionChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelectionChanged

    End Sub
    Private Sub ReportAP_ACC_adjust_Item()
        Dim AP_ACC_adjust_Item, Open_jn, LopAP_ACC_adjust_Item As String
        Dim Yr, Yrl As Integer

        Yr = Year(Today)
        Yrl = Year(Today) - 1

        AP_ACC_adjust_Item = "" : Open_jn = "" : LopAP_ACC_adjust_Item = ""
        If txtBook.Text <> "All" Then
            AP_ACC_adjust_Item = AP_ACC_adjust_Item & " AND AP_ACC_adjust_Item.book, '" & Len(txtBook.Text.Trim)
            'Open_jn = Open_jn & " AND Open_jn.book, '" & Len(txtBook.Text.Trim)
        End If
        If RadioButton1.Checked = True Then
            AP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "1" & "' AND '" & "1" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "' AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            LopAP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "12" & "' AND '" & "12" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yrl & "' AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "1" & "' AND '" & "1" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            CNN.Execute(" DELETE FROM Ap_Sum_AP_ACC_adjust_Item ")
            CNN.Execute("INSERT INTO Ap_Sum_AP_ACC_adjust_Item ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM AP_ACC_adjust_Item  WHERE 1=1 " & AP_ACC_adjust_Item & " ")
            '==============================
            Call LoadSqlData("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'   ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '===========================
            Call LoadSqlData("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM AP_ACC_adjust_Item  WHERE 1=1 " & LopAP_ACC_adjust_Item & " GROUP BY Company,Certify ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE Certify=N'" & RSC.Fields("Certify").Value & "' AND Company=N'" & RSC.Fields("Company").Value & "'  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Call LoadSqlData("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_AP_ACC_adjust_Item", RSC)
            If RSC.RecordCount > 1 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Rem_Amt= +( " & CDbl(RSC.Fields("Open_Amt").Value) & "+" & CDbl(RSC.Fields("amount_Dr").Value) & " -" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'  ")

                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Rem_Amt  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton2.Checked = True Then
            AP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "2" & "' AND '" & "2" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            LopAP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "1" & "' AND '" & "1" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "2" & "' AND '" & "2" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            CNN.Execute(" DELETE FROM Ap_Sum_AP_ACC_adjust_Item ")
            CNN.Execute("INSERT INTO Ap_Sum_AP_ACC_adjust_Item ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM AP_ACC_adjust_Item  WHERE 1=1 " & AP_ACC_adjust_Item & " ")
            '==============================
            Call LoadSqlData("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'   ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '===========================
            Call LoadSqlData("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM AP_ACC_adjust_Item  WHERE 1=1 " & LopAP_ACC_adjust_Item & " GROUP BY Company,Certify ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE Certify=N'" & RSC.Fields("Certify").Value & "' AND Company=N'" & RSC.Fields("Company").Value & "'  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Call LoadSqlData("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_AP_ACC_adjust_Item", RSC)
            If RSC.RecordCount > 1 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Rem_Amt= +( " & CDbl(RSC.Fields("Open_Amt").Value) & "+" & CDbl(RSC.Fields("amount_Dr").Value) & " -" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'  ")

                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Rem_Amt  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton3.Checked = True Then
            AP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "3" & "' AND '" & "3" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "'' "
            LopAP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "2" & "' AND '" & "2" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "'' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "3" & "' AND '" & "3" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            CNN.Execute(" DELETE FROM Ap_Sum_AP_ACC_adjust_Item ")
            CNN.Execute("INSERT INTO Ap_Sum_AP_ACC_adjust_Item ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM AP_ACC_adjust_Item  WHERE 1=1 " & AP_ACC_adjust_Item & " ")
            '==============================
            Call LoadSqlData("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'   ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '===========================
            Call LoadSqlData("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM AP_ACC_adjust_Item  WHERE 1=1 " & LopAP_ACC_adjust_Item & " GROUP BY Company,Certify ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE Certify=N'" & RSC.Fields("Certify").Value & "' AND Company=N'" & RSC.Fields("Company").Value & "'  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Call LoadSqlData("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_AP_ACC_adjust_Item", RSC)
            If RSC.RecordCount > 1 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Rem_Amt= +( " & CDbl(RSC.Fields("Open_Amt").Value) & "+" & CDbl(RSC.Fields("amount_Dr").Value) & " -" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'  ")

                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Rem_Amt  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton4.Checked = True Then
            AP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "4" & "' AND '" & "4" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            LopAP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "3" & "' AND '" & "3" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "4" & "' AND '" & "4" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            CNN.Execute(" DELETE FROM Ap_Sum_AP_ACC_adjust_Item ")
            CNN.Execute("INSERT INTO Ap_Sum_AP_ACC_adjust_Item ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM AP_ACC_adjust_Item  WHERE 1=1 " & AP_ACC_adjust_Item & " ")
            '==============================
            Call LoadSqlData("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'   ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '===========================
            Call LoadSqlData("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM AP_ACC_adjust_Item  WHERE 1=1 " & LopAP_ACC_adjust_Item & " GROUP BY Company,Certify ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE Certify=N'" & RSC.Fields("Certify").Value & "' AND Company=N'" & RSC.Fields("Company").Value & "'  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Call LoadSqlData("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_AP_ACC_adjust_Item", RSC)
            If RSC.RecordCount > 1 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Rem_Amt= +( " & CDbl(RSC.Fields("Open_Amt").Value) & "+" & CDbl(RSC.Fields("amount_Dr").Value) & " -" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'  ")

                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Rem_Amt  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton5.Checked = True Then
            AP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "5" & "' AND '" & "5" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            LopAP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "4" & "' AND '" & "4" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "5" & "' AND '" & "5" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            CNN.Execute(" DELETE FROM Ap_Sum_AP_ACC_adjust_Item ")
            CNN.Execute("INSERT INTO Ap_Sum_AP_ACC_adjust_Item ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM AP_ACC_adjust_Item  WHERE 1=1 " & AP_ACC_adjust_Item & " ")
            '==============================
            Call LoadSqlData("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'   ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '===========================
            Call LoadSqlData("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM AP_ACC_adjust_Item  WHERE 1=1 " & LopAP_ACC_adjust_Item & " GROUP BY Company,Certify ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE Certify=N'" & RSC.Fields("Certify").Value & "' AND Company=N'" & RSC.Fields("Company").Value & "'  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Call LoadSqlData("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_AP_ACC_adjust_Item", RSC)
            If RSC.RecordCount > 1 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Rem_Amt= +( " & CDbl(RSC.Fields("Open_Amt").Value) & "+" & CDbl(RSC.Fields("amount_Dr").Value) & " -" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'  ")

                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Rem_Amt  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton6.Checked = True Then
            AP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "6" & "' AND '" & "6" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            LopAP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "5" & "' AND '" & "5" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "6" & "' AND '" & "6" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            CNN.Execute(" DELETE FROM Ap_Sum_AP_ACC_adjust_Item ")
            CNN.Execute("INSERT INTO Ap_Sum_AP_ACC_adjust_Item ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM AP_ACC_adjust_Item  WHERE 1=1 " & AP_ACC_adjust_Item & " ")
            '==============================
            Call LoadSqlData("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'   ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '===========================
            Call LoadSqlData("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM AP_ACC_adjust_Item  WHERE 1=1 " & LopAP_ACC_adjust_Item & " GROUP BY Company,Certify ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE Certify=N'" & RSC.Fields("Certify").Value & "' AND Company=N'" & RSC.Fields("Company").Value & "'  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Call LoadSqlData("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_AP_ACC_adjust_Item", RSC)
            If RSC.RecordCount > 1 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Rem_Amt= +( " & CDbl(RSC.Fields("Open_Amt").Value) & "+" & CDbl(RSC.Fields("amount_Dr").Value) & " -" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'  ")

                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Rem_Amt  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton7.Checked = True Then
            AP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "7" & "' AND '" & "7" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            LopAP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "6" & "' AND '" & "6" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "7" & "' AND '" & "7" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            CNN.Execute(" DELETE FROM Ap_Sum_AP_ACC_adjust_Item ")
            CNN.Execute("INSERT INTO Ap_Sum_AP_ACC_adjust_Item ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM AP_ACC_adjust_Item  WHERE 1=1 " & AP_ACC_adjust_Item & " ")
            '==============================
            Call LoadSqlData("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'   ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '===========================
            Call LoadSqlData("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM AP_ACC_adjust_Item  WHERE 1=1 " & LopAP_ACC_adjust_Item & " GROUP BY Company,Certify ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE Certify=N'" & RSC.Fields("Certify").Value & "' AND Company=N'" & RSC.Fields("Company").Value & "'  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Call LoadSqlData("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_AP_ACC_adjust_Item", RSC)
            If RSC.RecordCount > 1 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Rem_Amt= +( " & CDbl(RSC.Fields("Open_Amt").Value) & "+" & CDbl(RSC.Fields("amount_Dr").Value) & " -" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'  ")

                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Rem_Amt  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton8.Checked = True Then
            AP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "8" & "' AND '" & "8" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            LopAP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "7" & "' AND '" & "7" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "8" & "' AND '" & "8" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            CNN.Execute(" DELETE FROM Ap_Sum_AP_ACC_adjust_Item ")
            CNN.Execute("INSERT INTO Ap_Sum_AP_ACC_adjust_Item ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM AP_ACC_adjust_Item  WHERE 1=1 " & AP_ACC_adjust_Item & " ")
            '==============================
            Call LoadSqlData("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'   ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '===========================
            Call LoadSqlData("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM AP_ACC_adjust_Item  WHERE 1=1 " & LopAP_ACC_adjust_Item & " GROUP BY Company,Certify ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE Certify=N'" & RSC.Fields("Certify").Value & "' AND Company=N'" & RSC.Fields("Company").Value & "'  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Call LoadSqlData("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_AP_ACC_adjust_Item", RSC)
            If RSC.RecordCount > 1 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Rem_Amt= +( " & CDbl(RSC.Fields("Open_Amt").Value) & "+" & CDbl(RSC.Fields("amount_Dr").Value) & " -" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'  ")

                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Rem_Amt  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton9.Checked = True Then
            AP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "9" & "' AND '" & "9" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            LopAP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "8" & "' AND '" & "8" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "9" & "' AND '" & "9" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            CNN.Execute(" DELETE FROM Ap_Sum_AP_ACC_adjust_Item ")
            CNN.Execute("INSERT INTO Ap_Sum_AP_ACC_adjust_Item ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM AP_ACC_adjust_Item  WHERE 1=1 " & AP_ACC_adjust_Item & " ")
            '==============================
            Call LoadSqlData("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'   ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '===========================
            Call LoadSqlData("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM AP_ACC_adjust_Item  WHERE 1=1 " & LopAP_ACC_adjust_Item & " GROUP BY Company,Certify ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE Certify=N'" & RSC.Fields("Certify").Value & "' AND Company=N'" & RSC.Fields("Company").Value & "'  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Call LoadSqlData("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_AP_ACC_adjust_Item", RSC)
            If RSC.RecordCount > 1 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Rem_Amt= +( " & CDbl(RSC.Fields("Open_Amt").Value) & "+" & CDbl(RSC.Fields("amount_Dr").Value) & " -" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'  ")

                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Rem_Amt  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton10.Checked = True Then
            AP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "10" & "' AND '" & "10" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            LopAP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "9" & "' AND '" & "9" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "10" & "' AND '" & "10" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            CNN.Execute(" DELETE FROM Ap_Sum_AP_ACC_adjust_Item ")
            CNN.Execute("INSERT INTO Ap_Sum_AP_ACC_adjust_Item ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM AP_ACC_adjust_Item  WHERE 1=1 " & AP_ACC_adjust_Item & " ")
            '==============================
            Call LoadSqlData("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'   ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '===========================
            Call LoadSqlData("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM AP_ACC_adjust_Item  WHERE 1=1 " & LopAP_ACC_adjust_Item & " GROUP BY Company,Certify ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE Certify=N'" & RSC.Fields("Certify").Value & "' AND Company=N'" & RSC.Fields("Company").Value & "'  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Call LoadSqlData("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_AP_ACC_adjust_Item", RSC)
            If RSC.RecordCount > 1 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Rem_Amt= +( " & CDbl(RSC.Fields("Open_Amt").Value) & "+" & CDbl(RSC.Fields("amount_Dr").Value) & " -" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'  ")

                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Rem_Amt  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton11.Checked = True Then
            AP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "11" & "' AND '" & "11" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            LopAP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "10" & "' AND '" & "10" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "11" & "' AND '" & "11" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            CNN.Execute(" DELETE FROM Ap_Sum_AP_ACC_adjust_Item ")
            CNN.Execute("INSERT INTO Ap_Sum_AP_ACC_adjust_Item ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM AP_ACC_adjust_Item  WHERE 1=1 " & AP_ACC_adjust_Item & " ")
            '==============================
            Call LoadSqlData("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'   ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '===========================
            Call LoadSqlData("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM AP_ACC_adjust_Item  WHERE 1=1 " & LopAP_ACC_adjust_Item & " GROUP BY Company,Certify ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE Certify=N'" & RSC.Fields("Certify").Value & "' AND Company=N'" & RSC.Fields("Company").Value & "'  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Call LoadSqlData("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_AP_ACC_adjust_Item", RSC)
            If RSC.RecordCount > 1 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Rem_Amt= +( " & CDbl(RSC.Fields("Open_Amt").Value) & "+" & CDbl(RSC.Fields("amount_Dr").Value) & " -" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'  ")

                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Rem_Amt  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton12.Checked = True Then
            AP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "12" & "' AND '" & "12" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            LopAP_ACC_adjust_Item = " AND month(AP_ACC_adjust_Item.date_work   ) BETWEEN '" & "11" & "' AND '" & "11" & "' AND year(AP_ACC_adjust_Item.date_work )='" & Yr & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            Open_jn = " AND month(Open_jn.date_work   ) BETWEEN '" & "12" & "' AND '" & "12" & "' AND year(Open_jn.date_work )='" & Yrl & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            CNN.Execute(" DELETE FROM Ap_Sum_AP_ACC_adjust_Item ")
            CNN.Execute("INSERT INTO Ap_Sum_AP_ACC_adjust_Item ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM AP_ACC_adjust_Item  WHERE 1=1 " & AP_ACC_adjust_Item & " ")
            '==============================
            Call LoadSqlData("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'   ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '===========================
            Call LoadSqlData("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM AP_ACC_adjust_Item  WHERE 1=1 " & LopAP_ACC_adjust_Item & " GROUP BY Company,Certify ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE Certify=N'" & RSC.Fields("Certify").Value & "' AND Company=N'" & RSC.Fields("Company").Value & "'  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Call LoadSqlData("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_AP_ACC_adjust_Item", RSC)
            If RSC.RecordCount > 1 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Rem_Amt= +( " & CDbl(RSC.Fields("Open_Amt").Value) & "+" & CDbl(RSC.Fields("amount_Dr").Value) & " -" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'  ")

                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Rem_Amt  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton13.Checked = True Then
            AP_ACC_adjust_Item = " AND year(AP_ACC_adjust_Item.date_work ) BETWEEN '" & Year(dts.Value) & "' AND '" & Year(dtt.Value) & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            LopAP_ACC_adjust_Item = " AND year(AP_ACC_adjust_Item.date_work ) BETWEEN '" & Year(dts.Value) - 1 & "' AND '" & Year(dtt.Value) - 1 & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            Open_jn = " AND year(Open_jn.date_work ) BETWEEN '" & Year(dts.Value) & "' AND '" & Year(dtt.Value) - 1 & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            CNN.Execute(" DELETE FROM Ap_Sum_AP_ACC_adjust_Item ")
            CNN.Execute("INSERT INTO Ap_Sum_AP_ACC_adjust_Item ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM AP_ACC_adjust_Item  WHERE 1=1 " & AP_ACC_adjust_Item & " ")
            '==============================
            Call LoadSqlData("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'   ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '===========================
            Call LoadSqlData("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM AP_ACC_adjust_Item  WHERE 1=1 " & LopAP_ACC_adjust_Item & " GROUP BY Company,Certify ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE Certify=N'" & RSC.Fields("Certify").Value & "' AND Company=N'" & RSC.Fields("Company").Value & "'  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Call LoadSqlData("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_AP_ACC_adjust_Item", RSC)
            If RSC.RecordCount > 1 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Rem_Amt= +( " & CDbl(RSC.Fields("Open_Amt").Value) & "+" & CDbl(RSC.Fields("amount_Dr").Value) & " -" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'  ")

                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Rem_Amt  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
        End If
        If RadioButton14.Checked = True Then
            AP_ACC_adjust_Item = " AND AP_ACC_adjust_Item.date_work   BETWEEN '" & Format(dts.Value, "yyyy-MM-dd") & "' AND '" & Format(dtt.Value, "yyyy-MM-dd") & "'AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            LopAP_ACC_adjust_Item = " AND AP_ACC_adjust_Item.date_work  < '" & Format(dts.Value, "yyyy-MM-dd") & "' AND AP_ACC_adjust_Item.Company=N'" & MuSubOff & "' "
            Open_jn = " AND Open_jn.date_work   BETWEEN '" & Format(dts.Value, "yyyy-MM-dd") & "' AND '" & Format(dtt.Value, "yyyy-MM-dd") & "'AND Open_jn.Company=N'" & MuSubOff & "' "
            'StartLoadDataList()
            '=====================
            CNN.Execute(" DELETE FROM Ap_Sum_AP_ACC_adjust_Item ")
            CNN.Execute("INSERT INTO Ap_Sum_AP_ACC_adjust_Item ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
                   " SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
                   " FROM AP_ACC_adjust_Item  WHERE 1=1 " & AP_ACC_adjust_Item & " ")
            '==============================
            Call LoadSqlData("SELECT Open_jn.ac_code, SUM(ISNULL(Open_jn.amount_dr,0)) AS amount_dr ,SUM(ISNULL(Open_jn.amount_cr,0)) AS amount_cr  " & _
                  " FROM Open_jn   GROUP BY Open_jn.ac_code ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'   ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '===========================
            Call LoadSqlData("SELECT SUM(amount_dr) AS amount_dr,SUM(amount_cr) as amount_cr,Company,Certify  " & _
               " FROM AP_ACC_adjust_Item  WHERE 1=1 " & LopAP_ACC_adjust_Item & " GROUP BY Company,Certify ", RSC)
            If RSC.RecordCount <> 0 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Open_Amt + (" & CDbl(RSC.Fields("amount_Dr").Value) & "-" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE Certify=N'" & RSC.Fields("Certify").Value & "' AND Company=N'" & RSC.Fields("Company").Value & "'  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
            '=======================================
            Call LoadSqlData("SELECT ac_code, Open_Amt, amount_Dr, amount_Cr FROM Ap_Sum_AP_ACC_adjust_Item", RSC)
            If RSC.RecordCount > 1 Then
                While Not RSC.EOF
                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Rem_Amt= +( " & CDbl(RSC.Fields("Open_Amt").Value) & "+" & CDbl(RSC.Fields("amount_Dr").Value) & " -" & CDbl(RSC.Fields("amount_Cr").Value) & " ) WHERE ac_code='" & RSC.Fields("ac_code").Value.ToString & "'  ")

                    CNN.Execute("UPDATE Ap_Sum_AP_ACC_adjust_Item SET Open_Amt=Rem_Amt  ")
                    RSC.MoveNext()
                End While
            Else
                Exit Sub
            End If
        End If

        '================================ReportAP_ACC_adjust_Item==================================

        'CNN.Execute(" DELETE FROM Ap_Sum_AP_ACC_adjust_Item ")
        'CNN.Execute("INSERT INTO Ap_Sum_AP_ACC_adjust_Item ( date_work,code_dr, code_cr, ac_code, Certify,descrip,Open_Amt, amount_Dr, amount_Cr, Rem_Amt,Last_User ,Last_Update,Company)" & _
        '          " SELECT date_work,code_dr,code_cr,ac_code,0,0,Sum(amount_dr-amount_cr),0,0,0,Last_User,Last_Update,Company " & _
        '          " FROM Open_jn   WHERE 1=1 " & Open_jn & " GROUP BY date_work,code_dr,code_cr,ac_code,Last_User,Last_Update,Company  " & _
        '          " UNION ALL " & _
        '          "SELECT date_work,code_dr,code_cr,ac_code,certify,descrip,0,amount_dr,amount_cr,0,Last_User,Last_Update,Company " & _
        '          " FROM AP_ACC_adjust_Item  WHERE 1=1 " & AP_ACC_adjust_Item & " ")
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
        If GetString(FG.Rows(0).Cells(1).Value) = "" Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
        'SLF = "SELECT " & MuLngRpt & "   *  ,  Acc_Code.Name_L  As AcNmeEx_L , Acc_Code.Name_E  As AcNmeEx_E FROM  AP_ACC_adjust_Item       INNER JOIN Acc_Code ON AP_ACC_adjust_Item.ac_code = Acc_Code.Ac_Code WHERE Book <>'' "
        SLF = "SELECT " & MuLngRpt & "   *  ,  AP_ACC_adjust_Item.ac_name  As AcNmeEx_L , AP_ACC_adjust_Item.ac_namee  As AcNmeEx_E FROM  AP_ACC_adjust_Item  WHERE Book <>'' "


        Call LoadLoGO()

        If CheckBox1.Checked = False Then
            Dim Rs As New ADODB.Recordset
            With Rs
                If .State = ConnectionState.Open Then .Close()
                .Open("  " & SLF & "  " & SQL & "order by  AP_ACC_adjust_Item.cnt ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Label5.Visible = False : Exit Sub
                If .EOF Then Exit Sub
            End With
            Dim FrmPreview As New FmPreview : FrmClosing()
            Dim Rpt As New CryGeneralLedgers

            If MdShowLOGO = 1 Then
                Rpt.Subreports(0).SetDataSource(RsLOGO)
            End If

            Rpt.SetDataSource(Rs)
            FrmPreview.ReportViewer.ReportSource = Rpt
            FrmPreview.ReportViewer.DisplayGroupTree = False
            FrmPreview.MdiParent = FmMain
            FrmPreview.WindowState = FormWindowState.Maximized
            Label5.Visible = False
            FrmPreview.Show()
            FrmPreview.Focus()
        Else
            '========
            Dim Rs As New ADODB.Recordset
            With Rs
                If .State = ConnectionState.Open Then .Close()
                .Open("  " & SLF & "  " & SQL & "    " & "" & " ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Label5.Visible = False : Exit Sub
                If .EOF Then Exit Sub
            End With
            Dim FrmPreview As New FmPreview : FrmClosing()
            Dim Rpt As New CryGeneralLedgersUser

            'If MdShowLOGO = 1 Then
            '    Rpt.Subreports(0).SetDataSource(RsLOGO)
            'End If

            Rpt.SetDataSource(Rs)
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
        ''Call LoadSqlData("SELECT top 1 Right(certify,7) As  certify   FROM  AP_ACC_adjust_Item where  book ='" & CmbBook2.Text & "' And  year(date_work)='" & Format(CDate(GetString(FG.CurrentRow.Cells(1).Value)), "yyyy") & "'   Order by  Right(certify,7) DESC", srNum)
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














        Call LoadSqlData("SELECT AC_CODE FROM AP_ACC_adjust_Item WHERE   book ='" & CmbBook2.Text & "' And  certify = N'" & txtNewId.Text & "' And  year(date_work)=" & Format(CDate(GetString(FG.CurrentRow.Cells(1).Value)), "yyyy") & " ", RSC)
        If RSC.RecordCount > 0 Then
            MsgBox("ເລກລະຫັດ : " & Trim(txtNewId.Text) & " ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
            txtNewId.Focus()
            If RSC.State = ConnectionState.Open Then RSC.Close()
            Exit Sub

        End If







        'Call LoadSqlData("select *  from AP_ACC_adjust_Item WHERE cnt<>''  " & SQL & "order by certify", RSC)
        ''Call LoadData("SELECT * FROM  AP_ACC_adjust_Item", RSC)
        'If RSC.RecordCount > 0 Then

        'End If


        If MessageBox.Show("ທ່ານຕ້ອງການປ່ຽນລະຫັດ " & txtOldId.Text & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            CNN.Execute("UPDATE AP_ACC_adjust_Item SET book='" & CmbBook2.Text & "',certify= N'" & txtNewId.Text & "' WHERE   book =N'" & GetString(FG.CurrentRow.Cells(16).Value) & "' And  certify  =N'" & MDInvoiceNo & "'   And  year(date_work)='" & Format(CDate(GetString(FG.CurrentRow.Cells(1).Value)), "yyyy") & "' ")

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





        Dim rst As New ADODB.Recordset
        CmbBook2.Items.Clear()
        Comm = New ADODB.Command
        Comm.ActiveConnection = CNN
        Comm.CommandText = "SELECT * FROM books WHERE bookid <> '" & "" & "'"
        rst = Comm.Execute
        If rst.RecordCount <> 0 Then
            While Not rst.EOF()
                CmbBook2.Items.Add(Trim(rst.Fields("bookid").Value))
                rst.MoveNext()
            End While
        End If

        CmbBook2.Text = "GL"
        If FG.CurrentRow IsNot Nothing Then
            Dim dateStr As String = If(FG.CurrentRow.Cells(1).Value Is Nothing, "1900-01-01", FG.CurrentRow.Cells(1).Value.ToString())
            Dim yearStr As String = Format(CDate(dateStr), "yyyy")
            Dim rowBook As String = If(FG.CurrentRow.Cells(16).Value Is Nothing, "", FG.CurrentRow.Cells(16).Value.ToString())
            
            Call LoadSqlData("select book , certify from AP_ACC_adjust_Item WHERE   book ='" & rowBook & "' And  certify  = '" & MDInvoiceNo & "'   And  year(date_work)='" & yearStr & "' order by cnt", RSC)

            If RSC.RecordCount > 0 Then
                txtOldId.Text = Trim(RSC.Fields("certify").Value)
                Books.Text = Trim(RSC.Fields("book").Value)
            End If
            CmbBook2.Text = Books.Text
        End If
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
        LngId = "7038" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As  Crl_Ac_Name	 ,"
        If MuLng = "L" Then
            MuLngRpt = MuLngRpt & "N'" & Amount_In_Word & "' As Crl_Amt_In_Word	 ,"
        Else
            MuLngRpt = MuLngRpt & "N'" & Amount_In_Word & "' As Crl_Amt_In_Word	 ,"
        End If


        'SLF = MuLngRpt & " AP_ACC_adjust_Item.company ,AP_ACC_adjust_Item.Date_Work , AP_ACC_adjust_Item.certify, AP_ACC_adjust_Item.ac_code, AP_ACC_adjust_Item.descrip , AP_ACC_adjust_Item.descripe , AP_ACC_adjust_Item.amt_dr, AP_ACC_adjust_Item.amt_cr, Acc_Code.Name_L AS Name_L , Acc_Code.Name_E AS Name_E  "
        SLF = MuLngRpt & " AP_ACC_adjust_Item.company ,AP_ACC_adjust_Item.Date_Work , AP_ACC_adjust_Item.certify, AP_ACC_adjust_Item.ac_code, AP_ACC_adjust_Item.descrip , AP_ACC_adjust_Item.descripe , AP_ACC_adjust_Item.amt_dr, AP_ACC_adjust_Item.amt_cr, AP_ACC_adjust_Item.ac_name  AS Name_L , AP_ACC_adjust_Item.ac_namee AS Name_E  "
        'SLF = MuLngRpt & " AP_ACC_adjust_Item.company ,AP_ACC_adjust_Item.Date_Work , AP_ACC_adjust_Item.certify, AP_ACC_adjust_Item.ac_code, AP_ACC_adjust_Item.descrip , AP_ACC_adjust_Item.descripe , AP_ACC_adjust_Item.amt_dr, AP_ACC_adjust_Item.amt_cr, Acc_Code.Name_L AS Name_L , Acc_Code.Name_E AS Name_E  "

        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            Dim dateStr As String = If(FG.CurrentRow.Cells(1).Value Is Nothing, "1900-01-01", FG.CurrentRow.Cells(1).Value.ToString())
            Dim yearStr As String = Format(CDate(dateStr), "yyyy")
            .Open("SELECT   " & SLF & "   FROM AP_ACC_adjust_Item INNER JOIN Acc_Code ON AP_ACC_adjust_Item.ac_code = Acc_Code.Ac_Code WHERE AP_ACC_adjust_Item.certify = N'" & MDInvoiceNo & "' And  year(date_work)=" & yearStr & "  order by AP_ACC_adjust_Item.cnt", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim Rpt As New CryNewsJerneralJournal
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
            For J = 0 To FG.Rows.Count - 1
                Dim row As DataGridViewRow = FG.Rows(J)
                Dim acCodeVal As String = If(row.Cells(4).Value Is Nothing, "", row.Cells(4).Value.ToString())
                If acCodeVal <> "" Then
                    If CDbl(If(row.Cells(10).Value Is Nothing, 0, row.Cells(10).Value)) <> 0 Then
                        row.Cells(7).Style.Font = New Font(FG.Font, FontStyle.Bold)
                        row.Cells(10).Style.Font = New Font(FG.Font, FontStyle.Bold)
                    Else
                        row.Cells(8).Style.Font = New Font(FG.Font, FontStyle.Bold)
                        row.Cells(11).Style.Font = New Font(FG.Font, FontStyle.Bold)
                    End If
                    row.Cells(4).Style.Font = New Font(FG.Font, FontStyle.Bold)
                End If
                
                Dim lockVal As String = If(row.Cells(14).Value Is Nothing, "", row.Cells(14).Value.ToString())
                Dim C2 As Color = Color.Red
                If lockVal = "1" Then
                    For Each cell As DataGridViewCell In row.Cells
                        cell.Style.ForeColor = Color.Red
                    Next
                ElseIf lockVal = "2" Then
                    For Each cell As DataGridViewCell In row.Cells
                        cell.Style.ForeColor = Color.Gray
                    Next
                End If
            Next J
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
        If FG.Rows.Count = 0 Then lblpage_total.Text = "0/0"
        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        If CmbPage.Items.Count > 0 Then
            CmbPage.SelectedIndex = 0
        End If
        If FG.Rows.Count = 0 Then CmbPage.Text = ""
        LbPage.Focus()
    End Sub
    Public Sub PageCnt(ByVal StrSQL As String, ByVal ConStr As String, ByVal PageNum As Long, ByVal RowPerPage As Integer)
        Load_DES()
        Label5.Visible = True
        Label5.BringToFront()
        x0 = 0
        y0 = 0
        Dim RsLoad As New ADODB.Recordset
        Dim rssum As New ADODB.Recordset
        Dim i As Integer
        FG.Rows.Clear()
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
                MULook2 = "  And  Left(AP_ACC_adjust_Item.company,2)= '" & OfUsr3 & "' "
            Else
                MULook2 = "  And AP_ACC_adjust_Item.company= '" & OfUsr1 & "' "
            End If
        End If
        SQL = " AND AP_ACC_adjust_Item.date_work   BETWEEN '" & Format(dts.Value, "yyyy-MM-dd") & "' AND '" & Format(dtt.Value, "yyyy-MM-dd") & "'  " & SR & " " & MULook2 & " "

        x = " date_work , certify , cheque_no , Referno,  Ac_code,Ac_name, Ac_namee, code_dr , code_cr , descrip , descripe , amount_dr , amount_cr , amount ,  amt_dr , amt_cr  ,  curr , cnt , lock  , book, Company "
        Dim L As String = "select " & x & "  from AP_ACC_adjust_Item WHERE certify<>''  " & SQL & "order by " & CntNB & ""
        BtnSearch.Visible = False
        LoadSqlData(L, RSC)

        With RSC
            If .RecordCount <> 0 Then
                .MoveFirst()
                .Move(RowPerPage * PageNum)
                LbPage.Text = Int(.RecordCount)
                If Int(.RecordCount Mod RowPerPage) = 0 Then
                    Last_page = Int(.RecordCount / DividePage)
                Else
                    Last_page = Int(.RecordCount / DividePage) + 1
                    If P = Last_page Then RowPerPage = (.RecordCount Mod RowPerPage)
                End If
                FG.Rows.Clear()
                
                px = 0
                CMS2.Enabled = False
                CMS3.Enabled = False
                
                For i = 0 To RowPerPage - 1
                    Dim s As String
                    If MuLng = "L" Then s = Trim(CStr(.Fields("Ac_name").Value.ToString)) Else s = Trim(CStr(.Fields("Ac_namee").Value.ToString))
                    
                    FG.Rows.Add(.AbsolutePosition, _
                                Format(CDate(Trim(.Fields("date_work").Value)), "dd/MM/yyyy"), _
                                Trim(CStr(.Fields("certify").Value.ToString)), _
                                Trim(CStr(.Fields("Referno").Value.ToString)), _
                                Trim(CStr(.Fields("Ac_code").Value)), _
                                Format(CDbl(Trim(.Fields("amount").Value)), "##,##0.00"), _
                                s, _
                                Format(CDbl(Trim(.Fields("amount_dr").Value)), "##,##0.00"), _
                                Format(CDbl(Trim(.Fields("amount_cr").Value)), "##,##0.00"), _
                                Trim(CStr(.Fields("curr").Value)), _
                                Format(CDbl(Trim(.Fields("amt_dr").Value)), "##,##0.00"), _
                                Format(CDbl(Trim(.Fields("amt_cr").Value)), "##,##0.00"), _
                                Trim(CStr(.Fields("company").Value)), _
                                Trim(CStr(.Fields("cnt").Value)), _
                                Trim(CStr(.Fields("lock").Value)), _
                                "", _
                                Trim(CStr(.Fields("book").Value)))
                    .MoveNext()
                Next i
                lblpage_total.Text = P & "/" & Int(Last_page)
            Else
                FG.Rows.Clear()
            End If
        End With

        If FG.Rows.Count > 0 Then
            FirstPage.Enabled = True
            BackPage.Enabled = True
            NextPage.Enabled = True
            LasthPage.Enabled = True
            EnterPage.Enabled = True
            
            Dim firstRowIdx As String = FG.Rows(0).Cells(0).Value.ToString()
            Dim lastRowIdx As String = FG.Rows(FG.Rows.Count - 1).Cells(0).Value.ToString()
            LbPage.Text = firstRowIdx & " To " & lastRowIdx & ", Of " & LbPage.Text
            
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

        If NextPage.Enabled = False Then EnterPage.Text = "Back "
        If BackPage.Enabled = False Then EnterPage.Text = "Next  "
        Call loadColor()

        Ch = 0
        Call SumAmount()
        Label5.Visible = False
    End Sub
    Private Sub LoadDividePage()
        LoadSQLCheckbox()
        ClickMouseRadio()
        SQL = ""
        Panel4.Visible = False
        'MULook2 = "" : If MuSubOff <> "00-00" Then MULook2 = "And AP_ACC_adjust_Item." & Mid(MULook, 5, CDbl(Len(MULook)) - 4) Else 
        SQL = " AND AP_ACC_adjust_Item.date_work   BETWEEN '" & Format(dts.Value, "yyyy-MM-dd") & "' AND '" & Format(dtt.Value, "yyyy-MM-dd") & "'  " & SR & " " & MULook2 & " "
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
        If GetString(FG.Rows(0).Cells(1).Value) = "" Then lblpage_total.Text = "0/0"
        If CmbPage.SelectedIndex >= 0 Then
            CmbPage.SelectedIndex = P - 1
        End If
        If GetString(FG.Rows(0).Cells(1).Value) = "" Then CmbPage.Text = ""

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
        If GetString(FG.Rows(0).Cells(1).Value) = "" Then lblpage_total.Text = "0/0"
        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        If CmbPage.Items.Count > 0 Then
            CmbPage.SelectedIndex = 0
        End If
        If GetString(FG.Rows(0).Cells(1).Value) = "" Then CmbPage.Text = ""
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
        If GetString(FG.Rows(0).Cells(1).Value) = "" Then lblpage_total.Text = "0/0"
        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        If CmbPage.Items.Count > 0 Then
            CmbPage.SelectedIndex = 0
        End If
        If GetString(FG.Rows(0).Cells(1).Value) = "" Then CmbPage.Text = ""
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
        If GetString(FG.Rows(0).Cells(1).Value) = "" Then lblpage_total.Text = "0/0"
        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        If CmbPage.Items.Count > 0 Then
            CmbPage.SelectedIndex = 0
        End If
        If GetString(FG.Rows(0).Cells(1).Value) = "" Then CmbPage.Text = ""
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
        If GetString(FG.Rows(0).Cells(1).Value) = "" Then lblpage_total.Text = "0/0"
        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        If CmbPage.Items.Count > 0 Then
            CmbPage.SelectedIndex = 0
        End If
        If GetString(FG.Rows(0).Cells(1).Value) = "" Then CmbPage.Text = ""
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
        If GetString(FG.Rows(0).Cells(1).Value) = "" Then lblpage_total.Text = "0/0"
        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        If CmbPage.Items.Count > 0 Then
            CmbPage.SelectedIndex = 0
        End If
        If GetString(FG.Rows(0).Cells(1).Value) = "" Then CmbPage.Text = ""
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
        If FG.Rows.Count = 0 Then lblpage_total.Text = "0/0"
        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        If CmbPage.Items.Count > 0 Then
            CmbPage.SelectedIndex = 0
        End If
        If FG.Rows.Count = 0 Then CmbPage.Text = ""
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
        If FG.Rows.Count = 0 Then lblpage_total.Text = "0/0"

        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i

        If CmbPage.Items.Count > 0 Then
            CmbPage.SelectedIndex = 0
        End If
        If FG.Rows.Count = 0 Then CmbPage.Text = ""
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
        LoadSqlData("select sub_id , off_id , off_add2  from  Ap_office where off_id ='" & Mid(CmbCompany.Text, 1, 2) & "' group BY  sub_id  ,off_id , off_add2", RSC)
        With RSC
            Do Until .EOF = True
                Off_Usr.Items.Add((.Fields("sub_id").Value) & " " & (.Fields("off_add2").Value))
                .MoveNext()
            Loop
        End With

        Off_Usr.SelectedIndex = FmLogin.Sub_Company.SelectedIndex
        Off_Id = Mid(CmbCompany.Text, 1, 2)
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
        FmNsewJeneralJournal_Adjust.txtInvoice.Enabled = True
        FmNsewJeneralJournal_Adjust.CmbBook.Enabled = True

        FmNsewJeneralJournal_Adjust.ShowDialog()
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
            FmNsewJeneralJournal_Adjust.txtInvoice.Enabled = False
            FmNsewJeneralJournal_Adjust.CmbBook.Enabled = False
            FmNsewJeneralJournal_Adjust.ShowDialog()
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
            Dim dateVal As String = If(FG.CurrentRow.Cells(1).Value Is Nothing, "1900-01-01", FG.CurrentRow.Cells(1).Value.ToString())
            CNN.Execute("delete AP_ACC_adjust_Item where certify ='" & MDInvoiceNo & "' And  year(date_work)='" & Format(CDate(dateVal), "yyyy") & "' ")
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
        Call LoadSqlData("select top 1 count(cnt) as cnt from AP_ACC_adjust_Item", RSC)
        If RSC.RecordCount > 0 Then
            TextBox2.Text = CDbl(Trim(RSC.Fields("cnt").Value))
        End If
    End Sub
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TT()

        TextBox4.Text = CDbl(TextBox2.Text) / (txtSC15.Text)
        FG.Rows.Clear()
        With RSC
            Call LoadSqlData("SELECT top " & txtSC15.Text & " * FROM  AP_ACC_adjust_Item Order by cnt", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    Dim sVal As String
                    If MuLng = "L" Then sVal = Trim(CStr(.Fields("descrip").Value)) Else sVal = Trim(CStr(.Fields("descripe").Value.ToString))

                    FG.Rows.Add(.AbsolutePosition, Format(CDate(Trim(.Fields("date_work").Value)), "dd/MM/yyyy"), _
                                Trim(CStr(.Fields("certify").Value)), Trim(CStr(.Fields("cheque_no").Value)), _
                                Trim(CStr(.Fields("Ac_code").Value)), Format(CDbl(Trim(.Fields("amount").Value)), "##,##0.00"), _
                                sVal, Format(CDbl(Trim(.Fields("amount_dr").Value)), "##,##0.00"), _
                                Format(CDbl(Trim(.Fields("amount_cr").Value)), "##,##0.00"), Trim(CStr(.Fields("curr").Value)), _
                                Format(CDbl(Trim(.Fields("amt_dr").Value)), "##,##0.00"), Format(CDbl(Trim(.Fields("amt_cr").Value)), "##,##0.00"), _
                                Trim(CStr(.Fields("company").Value)), Trim(CStr(.Fields("cnt").Value)), _
                                Trim(CStr(.Fields("lock").Value)), "", Trim(CStr(.Fields("book").Value)))

                    .MoveNext()
                End While
            Else
                'FG.Rows = 16
            End If
        End With
    End Sub

    Private Sub Nme_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Nme.TextChanged

    End Sub

    Private Sub Button19_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        TT()
        D = FG.Rows.Count
        If CDbl(CDbl(TextBox2.Text) / CDbl(txtSC15.Text)) > Int(CDbl(TextBox2.Text) / CDbl(txtSC15.Text)) Then
            TextBox4.Text = Int(CDbl(TextBox2.Text) / CDbl(txtSC15.Text)) + 1
        Else
            TextBox4.Text = Int(CDbl(TextBox2.Text) / CDbl(txtSC15.Text))
        End If
        TextBox6.Text = 1
        TextBox5.Text = TextBox6.Text & "/" & TextBox4.Text
        FG.Rows.Clear()
        With RSC
            Call LoadSqlData("SELECT top " & txtSC15.Text & " * FROM  AP_ACC_adjust_Item  Order by certify , cnt", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    Dim sVal As String
                    If MuLng = "L" Then sVal = Trim(CStr(.Fields("descrip").Value)) Else sVal = Trim(CStr(.Fields("descripe").Value.ToString))

                    FG.Rows.Add(.AbsolutePosition, Format(CDate(Trim(.Fields("date_work").Value)), "dd/MM/yyyy"), _
                                Trim(CStr(.Fields("certify").Value)), Trim(CStr(.Fields("cheque_no").Value)), _
                                Trim(CStr(.Fields("Ac_code").Value)), Format(CDbl(Trim(.Fields("amount").Value)), "##,##0.00"), _
                                sVal, Format(CDbl(Trim(.Fields("amount_dr").Value)), "##,##0.00"), _
                                Format(CDbl(Trim(.Fields("amount_cr").Value)), "##,##0.00"), Trim(CStr(.Fields("curr").Value)), _
                                Format(CDbl(Trim(.Fields("amt_dr").Value)), "##,##0.00"), Format(CDbl(Trim(.Fields("amt_cr").Value)), "##,##0.00"), _
                                Trim(CStr(.Fields("company").Value)), Trim(CStr(.Fields("lock").Value)), _
                                Trim(CStr(.Fields("lock").Value)), "", Trim(CStr(.Fields("book").Value)))

                    .MoveNext()
                End While
            Else
                'FG.Rows = 16
            End If
        End With
    End Sub

    Private Sub Button18_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click

    End Sub

    Private Sub Button17_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        'TT()

        Dim xVal As String = " date_work , certify ,descrip,descripe, cheque_no , Ac_code,code_dr , code_cr  , amount_dr , amount_cr , amount ,  amt_dr , amt_cr  ,  curr , cnt , lock  , book, Company "
        Dim lastCertify As String = ""
        If FG.Rows.Count > 0 Then
            lastCertify = If(FG.Rows(FG.Rows.Count - 1).Cells(2).Value Is Nothing, "", FG.Rows(FG.Rows.Count - 1).Cells(2).Value.ToString())
        End If
        
        Dim sSQL As String = "SELECT top " & txtSC15.Text & " " & xVal & "  FROM   AP_ACC_adjust_Item where certify > '" & lastCertify & "' Order by certify , cnt"
        TextBox6.Text = Int(TextBox6.Text) + 1
        TextBox5.Text = TextBox6.Text & "/" & TextBox4.Text
        FG.Rows.Clear()
        With RSC
            Call LoadSqlData(sSQL, RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    Dim sVal As String
                    If MuLng = "L" Then sVal = Trim(CStr(.Fields("descrip").Value)) Else sVal = Trim(CStr(.Fields("descripe").Value.ToString))
                    D = D + 1
                    
                    FG.Rows.Add(D, Format(CDate(Trim(.Fields("date_work").Value)), "dd/MM/yyyy"), _
                                Trim(CStr(.Fields("certify").Value)), Trim(CStr(.Fields("cheque_no").Value)), _
                                Trim(CStr(.Fields("Ac_code").Value)), Format(CDbl(Trim(.Fields("amount").Value)), "##,##0.00"), _
                                sVal, Format(CDbl(Trim(.Fields("amount_dr").Value)), "##,##0.00"), _
                                Format(CDbl(Trim(.Fields("amount_Cr").Value)), "##,##0.00"), Trim(CStr(.Fields("Curr").Value)), _
                                Format(CDbl(Trim(.Fields("amt_dr").Value)), "##,##0.00"), Format(CDbl(Trim(.Fields("amt_dr").Value)), "##,##0.00"), _
                                Trim(CStr(.Fields("company").Value)), Trim(CStr(.Fields("cnt").Value)), _
                                Trim(CStr(.Fields("lock").Value)), "", Trim(CStr(.Fields("book").Value)))

                    .MoveNext()
                End While
            Else
                'FG.Rows = 16
            End If
        End With
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click

    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        If RadioButton15.Checked = True Then
            If LockData = "" Then MsgBox("ກະລຸນາເລືອກກ່ອນ!", MsgBoxStyle.OkOnly) : Exit Sub
            Dim dateStr As String = If(FG.CurrentRow.Cells(1).Value Is Nothing, "1900-01-01", FG.CurrentRow.Cells(1).Value.ToString())
            Dim yearStr As String = Format(CDate(dateStr), "yyyy")
            LoadSqlData("Select lock from AP_ACC_adjust_Item where certify ='" & MDInvoiceNo & "' And  year(date_work)='" & yearStr & "'", RSC)
            If RSC.RecordCount <> 0 Then
                If (RSC.Fields("lock").Value) = "2" Then
                    MsgBox("ລາຍການນີ້ ໄດ້ປິດບັນຊີໄປແລ້ວບໍ່ສາມາດ " & Button1.Text & "  ໄດ້ອີກ!", MsgBoxStyle.OkOnly)
                    Exit Sub
                End If
            End If

            If MessageBox.Show("ທ່ານຕ້ອງການ " & Button1.Text & " ລະຫັດ " & MDInvoiceNo & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                If LockData = "1" Then
                    CNN.Execute("UPDATE AP_ACC_adjust_Item SET lock='0' where certify ='" & MDInvoiceNo & "' And  year(date_work)='" & yearStr & "'")
                Else
                    CNN.Execute("UPDATE AP_ACC_adjust_Item SET lock='1' where certify ='" & MDInvoiceNo & "' And  year(date_work)='" & yearStr & "'")
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
                    MULook2 = "  And  Left(AP_ACC_adjust_Item.company,2)= '" & OfUsr3 & "' "
                Else
                    MULook2 = "  And AP_ACC_adjust_Item.company= '" & OfUsr1 & "' "
                End If
            End If

            LoadSqlData("Select lock from AP_ACC_adjust_Item where  AP_ACC_adjust_Item.date_work   BETWEEN '" & Format(dts.Value, "yyyy-MM-dd") & "' AND '" & Format(dtt.Value, "yyyy-MM-dd") & "'   " & MULook2 & " ", RSC)
            If RSC.RecordCount <> 0 Then
                If (RSC.Fields("lock").Value) = "2" Then
                    MsgBox("ລາຍການພວກນີ້ ໄດ້ປິດບັນຊີໄປແລ້ວບໍ່ສາມາດ " & Button1.Text & "  ໄດ້ອີກ!", MsgBoxStyle.OkOnly)
                    Exit Sub
                End If
            End If

            If MessageBox.Show("ທ່ານຕ້ອງການປົດລ໋ອກຂໍ່ມູນແຕວັນທີ " & dts.Text & " ຫາ " & dtt.Text & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                If RadioButton17.Checked = True Then
                    CNN.Execute("UPDATE AP_ACC_adjust_Item SET lock='1' where  AP_ACC_adjust_Item.date_work   BETWEEN '" & Format(dts.Value, "yyyy-MM-dd") & "' AND '" & Format(dtt.Value, "yyyy-MM-dd") & "'   " & MULook2 & " ")
                Else
                    CNN.Execute("UPDATE AP_ACC_adjust_Item SET lock='0'  where  AP_ACC_adjust_Item.date_work   BETWEEN '" & Format(dts.Value, "yyyy-MM-dd") & "' AND '" & Format(dtt.Value, "yyyy-MM-dd") & "'   " & MULook2 & " ")
                End If
                LockData = ""
                LoadMonthSQL()
            End If
        End If

    End Sub

    Private Sub Button13_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Panel7.Visible = False
    End Sub
End Class