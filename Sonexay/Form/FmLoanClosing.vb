Public Class FmLoanClosing
    Dim MDTab As Integer
    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        CNN.Execute("insert into Ap_LoanClosing(Bnk_Ac_Code , Date_Work ,  Open_Amt , Paid_Amt , Paid_Inte , Rem_Amt , Last_Action , Action_Type , Las_Udate , Statuss  ) " & _
"Values('" & Bnk_Ac_Code.Text & "' , '" & Date_Work.Text & "' , '" & Open_Amt.Text & "' , '" & Paid_Amt.Text & "' , '" & Inte_Amt.Text & "' , '" & Rem_Amt.Text & "' , '" & Last_Action.Text & "' , '" & Action_Type.Text & "' , '" & Las_Udate.Text & "' , '" & Statuss.Text & "' )")
        MessageBox.Show("ok")
    End Sub

    Private Sub BtnEdit2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnEdit2.Click
        CNN.Execute(" update Ap_LoanClosing set " & _
      "Date_Work='" & Date_Work.Text & "' " & _
     ", Open_Amt='" & Open_Amt.Text & "' " & _
    ", Paid_Amt='" & Paid_Amt.Text & "' " & _
    ", Paid_Inte='" & Inte_Amt.Text & "' " & _
    ", Rem_Amt='" & Rem_Amt.Text & "' " & _
    ", Last_Action='" & Last_Action.Text & "' " & _
    ", Action_Type='" & Action_Type.Text & "' " & _
    ", Las_Udate='" & Las_Udate.Text & "' " & _
     ", Statuss='" & Statuss.Text & "' " & _
      " where Bnk_Ac_Code='" & Bnk_Ac_Code.Text & "' ")
        MessageBox.Show("ok")
    End Sub

    Private Sub BtnAddNew2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew2.Click
        Date_Work.Text = ""
        Open_Amt.Text = ""
        Paid_Amt.Text = ""
        Inte_Amt.Text = ""
        Rem_Amt.Text = ""
        Last_Action.Text = ""
        Action_Type.Text = ""
        Las_Udate.Text = ""
        Statuss.Text = ""
        Bnk_Ac_Code.Text = ""
    End Sub


    Private Sub BtnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew.Click
        ' Logic for adding a new record, possibly clearing input fields
    End Sub

    Private Sub FmLoanClosing_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim colRemovedTabs As New Collection()
        'Dim TabPage1 As TabPage
        'TabPage1 = FmOPenForm.TabControl1.TabPages(MDTab)
        'FmOPenForm.TabControl1.Controls.Remove(TabPage1)
        'MDTabIndex = MDTabIndex - 1
        'FmMain.ToolStripMenuItem55.Enabled = True
        'If MDTabIndex = 0 Then
        '    FmOPenForm.Close()
        'End If
    End Sub

    Private Sub FmLoanClosing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        BtnEdit2.Enabled = False
        BtnEdit.Enabled = False
        BtnDelete.Enabled = False
        SetupGrid() ' Call SetupGrid here
    End Sub

    Private Sub BtnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRefresh.Click
        LoadDataIntoGrid()
    End Sub

    Private Sub LoadDataIntoGrid()
        FG.Rows.Clear()
        Dim R As New ADODB.Recordset
        ' Assuming LoadAcData is a global function to load data
        Call LoadAcData("SELECT Bnk_Ac_Code, Date_Work, Open_Amt, Paid_Amt, Paid_Inte, Rem_Amt, Last_Action, Action_Type, Las_Udate, Statuss FROM Ap_LoanClosing", R)
        If R.RecordCount <> 0 Then
            While Not R.EOF
                FG.Rows.Add(R.Fields("Bnk_Ac_Code").Value, _
                            Format(CDate(R.Fields("Date_Work").Value.ToString), "dd/MM/yyyy"), _
                            R.Fields("Open_Amt").Value, _
                            R.Fields("Paid_Amt").Value, _
                            R.Fields("Paid_Inte").Value, _
                            R.Fields("Rem_Amt").Value, _
                            R.Fields("Last_Action").Value, _
                            R.Fields("Action_Type").Value, _
                            Format(CDate(R.Fields("Las_Udate").Value.ToString), "dd/MM/yyyy"), _
                            R.Fields("Statuss").Value)
                R.MoveNext()
            End While
        End If
    End Sub

    Private Sub BtnExit2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit2.Click
        Close()
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub SetupGrid()
        FG.Columns.Clear()
        FG.Rows.Clear()
        FG.AllowUserToAddRows = False
        FG.AllowUserToDeleteRows = False
        FG.ReadOnly = True
        FG.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG.MultiSelect = False
        FG.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        FG.RowHeadersVisible = False

        Dim c As DataGridViewColumn

        c = New DataGridViewTextBoxColumn() : c.Name = "Bnk_Ac_Code" : c.HeaderText = "ເລກບັນຊີເງິນກູ້" : FG.Columns.Add(c)
        c = New DataGridViewTextBoxColumn() : c.Name = "Date_Work" : c.HeaderText = "ວັນທີເຄືອນໄຫວ" : FG.Columns.Add(c)
        c = New DataGridViewTextBoxColumn() : c.Name = "Open_Amt" : c.HeaderText = "ຍອດຍົກມາ" : FG.Columns.Add(c)
        c = New DataGridViewTextBoxColumn() : c.Name = "Paid_Amt" : c.HeaderText = "ຈໍານວນເງິນຈ່າຍຄືນ" : c.DefaultCellStyle.Format = "N2" : c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight : FG.Columns.Add(c)
        c = New DataGridViewTextBoxColumn() : c.Name = "Paid_Inte" : c.HeaderText = "ຈ່າຍດອກເບ້ຍ" : c.DefaultCellStyle.Format = "N2" : c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight : FG.Columns.Add(c)
        c = New DataGridViewTextBoxColumn() : c.Name = "Rem_Amt" : c.HeaderText = "ຄ້າງຈ່າຍ" : c.DefaultCellStyle.Format = "N2" : c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight : FG.Columns.Add(c)
        c = New DataGridViewTextBoxColumn() : c.Name = "Last_Action" : c.HeaderText = "ຜູ້ປັບປຸງຫລ້າສຸດ" : FG.Columns.Add(c)
        c = New DataGridViewTextBoxColumn() : c.Name = "Action_Type" : c.HeaderText = "ວັນທີປັບປຸງ" : FG.Columns.Add(c)
        c = New DataGridViewTextBoxColumn() : c.Name = "Las_Udate" : c.HeaderText = "ວັນທີເຄື່ອນໄຫວຫລ້າສຸດ" : FG.Columns.Add(c)
        c = New DataGridViewTextBoxColumn() : c.Name = "Statuss" : c.HeaderText = "ສະຖານະ" : c.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill : FG.Columns.Add(c)

        UpdateColumnHeaders()
    End Sub

    Private Sub UpdateColumnHeaders()
        If FG.Columns.Count < 10 Then Exit Sub
        If MuLng = "L" Then
            FG.Columns("Bnk_Ac_Code").HeaderText = "ເລກບັນຊີເງິນກູ້"
            FG.Columns("Date_Work").HeaderText = "ວັນທີເຄືອນໄຫວ"
            FG.Columns("Open_Amt").HeaderText = "ຍອດຍົກມາ"
            FG.Columns("Paid_Amt").HeaderText = "ຈໍານວນເງິນຈ່າຍຄືນ"
            FG.Columns("Paid_Inte").HeaderText = "ຈ່າຍດອກເບ້ຍ"
            FG.Columns("Rem_Amt").HeaderText = "ຄ້າງຈ່າຍ"
            FG.Columns("Last_Action").HeaderText = "ຜູ້ປັບປຸງຫລ້າສຸດ"
            FG.Columns("Action_Type").HeaderText = "ວັນທີປັບປຸງ"
            FG.Columns("Las_Udate").HeaderText = "ວັນທີເຄື່ອນໄຫວຫລ້າສຸດ"
            FG.Columns("Statuss").HeaderText = "ສະຖານະ"
        Else
            FG.Columns("Bnk_Ac_Code").HeaderText = "Bank Account Code"
            FG.Columns("Date_Work").HeaderText = "Work Date"
            FG.Columns("Open_Amt").HeaderText = "Open Amount"
            FG.Columns("Paid_Amt").HeaderText = "Paid Amount"
            FG.Columns("Paid_Inte").HeaderText = "Paid Interest"
            FG.Columns("Rem_Amt").HeaderText = "Remaining Amount"
            FG.Columns("Last_Action").HeaderText = "Last Action"
            FG.Columns("Action_Type").HeaderText = "Action Type"
            FG.Columns("Las_Udate").HeaderText = "Last Update Date"
            FG.Columns("Statuss").HeaderText = "Status"
        End If
    End Sub
End Class