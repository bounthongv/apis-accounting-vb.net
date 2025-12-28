Public Class FrmList_ASST
    Dim cn As New Odbc.OdbcConnection
    Dim mSql, DDD As String

    Private Sub LdData(ByVal str As String)
        'Dim cRS As New ADODB.Recordset
        'Dim ss As String
        'ss = "SELECT * FROM ASSETS Where 1=1 " & str
        'Call LoadSqlData(ss, cRS)
        'If cRS.RecordCount <> 0 Then
        '    FG.Rows = 1
        '    FG.Redraw = False
        '    While Not cRS.EOF
        '        FG.AddItem(cRS.AbsolutePosition & Chr(9) & Trim(cRS.Fields("AssetID").Value.ToString) & Chr(9) & Trim(cRS.Fields("Asset_No").Value.ToString) & Chr(9) & Trim(cRS.Fields("Sect_ID").Value.ToString) & Chr(9) & Trim(cRS.Fields("Asset_Nm").Value.ToString) & Chr(9) & Trim(cRS.Fields("Group_ID").Value.ToString) & Chr(9) & Format(cRS.Fields("Date_Work").Value, "dd/MM/yyyy") & Chr(9) & cRS.Fields("Qty").Value.ToString & Chr(9) & Format(cRS.Fields("Amount").Value, "#,##0.00") & Chr(9) & Trim(cRS.Fields("Curr").Value.ToString) & Chr(9) & Format(cRS.Fields("Rate").Value, "#,##0.00") & Chr(9) & Format(cRS.Fields("Amt_KIP").Value, "#,##0.00"))
        '        cRS.MoveNext()
        '    End While
        'Else
        '    FG.Rows = 1
        '    FG.Rows = 2
        'End If
        'FG.Redraw = True
    End Sub

    Private Sub FrmAcc_Code_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        mSql = " Order by Asset_No"
        Call LdGrp()
        'Call LdSec()
        'Call LdDep()
        cmbSort.SelectedIndex = 0
        Call Button1_Click(sender, e)
        If mCompStr = "" Then btnAdd.Enabled = False
    End Sub
    Private Sub LdDep()
        Dim sRS As New ADODB.Recordset
        cmbDeprt.Items.Clear()
        cmbDeprt.Items.Add("** All Department ***")
        Call LoadSqlData("Select * from Department Order by DepartmentID", sRS)
        If sRS.RecordCount <> 0 Then
            While Not sRS.EOF
                cmbDeprt.Items.Add(sRS.Fields("DepartmentNm").Value.ToString)
                sRS.MoveNext()
            End While
        End If
        'If cmbDeprt.Items.Count > 0 Then cmbDeprt.SelectedIndex = 1
    End Sub
    Private Sub LdSec()
        Dim sRS As New ADODB.Recordset
        cmbSec.Items.Clear()
        cmbSec.Items.Add("** All Sections ***")
        Call LoadSqlData("Select * from Sections Order by SecID", sRS)
        If sRS.RecordCount <> 0 Then
            While Not sRS.EOF
                cmbSec.Items.Add(sRS.Fields("SecNmL").Value.ToString)
                sRS.MoveNext()
            End While
        End If
        If cmbSec.Items.Count > 0 Then cmbSec.SelectedIndex = 1
    End Sub

    Private Sub LdGrp()
        Dim gRS As New ADODB.Recordset
        cmbGrp.Items.Clear()
        cmbGrp.Items.Add("** All Groups ***")
        Call LoadSqlData("Select * from Groups Order by Group_ID", gRS)
        If gRS.RecordCount <> 0 Then
            While Not gRS.EOF
                cmbGrp.Items.Add(gRS.Fields("Group_Nm").Value.ToString)
                gRS.MoveNext()
            End While
        End If
        cmbGrp.SelectedIndex = 0
    End Sub
    Private Sub CndClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CndClose.Click
        myTemp = ""
        myTemp1 = ""
        Me.Close()
    End Sub


    Private Sub cmbSort_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSort.SelectedIndexChanged
        If cmbSort.SelectedIndex = 0 Then
            mSql = " Order by Section, AssetID"
        ElseIf cmbSort.SelectedIndex = 1 Then
            mSql = " Order by Asset_Nm"
        Else
            mSql = " Order by Date_Work"
        End If
        Call Button1_Click(sender, e)
    End Sub

    Private Sub cmbGrp_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbGrp.SelectedIndexChanged
        Dim gRS As New ADODB.Recordset
        Call LoadSqlData("select * from Groups Where Group_Nm=N'" & Trim(cmbGrp.Text) & "'", gRS)
        If gRS.RecordCount <> 0 Then
            txtGrp.Text = Trim(gRS.Fields("Group_ID").Value.ToString)
        Else
            txtGrp.Text = ""
        End If
        Call Button1_Click(sender, e)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        LoadDG()
    End Sub
    Private Sub LoadDG()
        Dim cRS As New ADODB.Recordset
        DDD = ""
        If cmbGrp.SelectedIndex > 0 Then
            DDD = DDD & " AND Group_ID='" & txtGrp.Text & "'"
        End If
        If cmbSec.SelectedIndex > 0 Then
            DDD = DDD & " AND Sect_ID='" & txtSec.Text & "'"
        End If
        If cmbDeprt.SelectedIndex > 0 Then
            DDD = DDD & " AND DepartmentID='" & txtDep.Text & "'"
        End If
        If txtCode.Text <> "" Then
            DDD = DDD & " AND (Left(Asset_No, " & Len(txtCode.Text) & ")='" & Trim(txtCode.Text) & "' OR Left(AssetID, " & Len(txtCode.Text) & ")='" & Trim(txtCode.Text) & "')"
        End If
        If txtNm.Text <> "" Then
            DDD = DDD & " AND Asset_Nm Like '%" & Trim(txtNm.Text) & "%'"
        End If
        If chkDT.Checked = True Then
            DDD = DDD & " AND Month(Date_Work)= " & DTUSE.Value.Month & " AND YEAR(Date_Work) = " & DTUSE.Value.Year & ""
        End If

        DG.DataSource = 0
        Dim ds As New DataSet
        Try
            ConnectCL()
            sql = "Select 0,AssetID,Asset_No,Asset_Nm,Group_ID,Date_Work,Used_Life,Amount,Curr,Rate,Amt_KIP,Deposted_Date from ASSETS where 1=1 " & DDD & " and   Deposted='0' "
            'SQL = "Select SecID,SecNmL,SecNmE,Remark from Sections  " & mSql & " "
            LoadCN_DG()
            da.Fill(ds, " ASSETS")
            DG.DataSource = ds.Tables(" ASSETS")
            cn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try
        'FG.FormatString = "^No.   |<ລະຫັດຊັບສິນ   |<ເລກປະຈຳຕົວ     |<ຊື່ຊັບສິນ                          |^ວັນທີນຳໃຊ້     |^ອາຍຸນຳໃຊ້|>ມູນຄ່າເດີມ     |^ສະກຸນເງິນ|>ອັດຕາແລກປ່ຽນ|>ມູນຄ່າຊັບສິນ    |^ວັນທີສະສາງ    "

        DG.Columns(0).HeaderText = "ລ/ດ" : DG.Columns(0).Width = "40"
        DG.Columns(1).HeaderText = "ລະຫັດຊັບສິນ" : DG.Columns(1).Width = "100"
        DG.Columns(2).HeaderText = "ເລກປະຈຳຕົວ" : DG.Columns(2).Width = "150"
        DG.Columns(3).HeaderText = "ຊື່ຊັບສິນ      " : DG.Columns(3).Width = "250"
        DG.Columns(4).HeaderText = "ໝວດ" : DG.Columns(4).Width = "50"
        DG.Columns(5).HeaderText = "ວັນທີນຳໃຊ້" : DG.Columns(5).Width = "90"
        DG.Columns(6).HeaderText = "ອາຍຸນຳໃຊ້" : DG.Columns(6).Width = "80"
        DG.Columns(7).HeaderText = "ມູນຄ່າເດີມ " : DG.Columns(7).Width = "125"
        DG.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DG.Columns(7).DefaultCellStyle.Format = "##,##0.00"
        DG.Columns(8).HeaderText = "ສະກຸນເງິນ" : DG.Columns(8).Width = "70"
        DG.Columns(9).HeaderText = "ອັດຕາ" : DG.Columns(9).Width = "80"
        DG.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DG.Columns(9).DefaultCellStyle.Format = "##,##0.00"
        DG.Columns(10).HeaderText = "ມູນຄ່າຊັບສິນ" : DG.Columns(10).Width = "125"
        DG.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DG.Columns(10).DefaultCellStyle.Format = "##,##0"
        DG.Columns(11).HeaderText = "ວັນທີສະສາງ" : DG.Columns(11).Width = "90"
        'DG.Columns(12).HeaderText = "ຈຳນວນ" : DG.Columns(12).Width = "50"
        Dim counter As Integer = 0
        For Each dr As DataGridViewRow In DG.Rows
            counter += 1
        Next
        Label8.Text = "ລາຍການທັງໝົດ " & counter.ToString - 1 & " ລາຍການ"
        For i As Integer = 0 To DG.Rows.Count - 1
            DG.Rows(i).HeaderCell.Value = i.ToString()
        Next i
        Dim row As Integer = 0
        For row = 0 To DG.RowCount - 2
            DG.Rows(row).Cells(0).Value = row + 1
        Next
        DG.ReadOnly = True
    End Sub
    Private Sub txtCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCode.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Button1_Click(sender, e)
        End If
    End Sub

    Private Sub txtCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCode.TextChanged

    End Sub

    Private Sub txtNm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNm.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Button1_Click(sender, e)
        End If
    End Sub

    Private Sub DTUSE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTUSE.ValueChanged
        Call Button1_Click(sender, e)
    End Sub

    Private Sub chkDT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDT.CheckedChanged
        If chkDT.Checked = True Then
            DTUSE.Enabled = True
        Else
            DTUSE.Enabled = False
        End If
    End Sub

    Private Sub lSort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lSort.Click

    End Sub

    Private Sub cmbSec_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSec.SelectedIndexChanged
        Dim sRS As New ADODB.Recordset
        Call LoadSqlData("select * from Sections Where SecNmL=N'" & Trim(cmbSec.Text) & "'", sRS)
        If sRS.RecordCount <> 0 Then
            txtSec.Text = Trim(sRS.Fields("SecID").Value.ToString)
            Dim dRS As New ADODB.Recordset
            cmbDeprt.Items.Clear()
            cmbDeprt.Items.Add("** All Department ***")
            Call LoadSqlData("Select * from Department Where Left(DepartmentID,2) = '" & Trim(txtSec.Text) & "' Order by DepartmentID", dRS)
            If dRS.RecordCount <> 0 Then
                While Not dRS.EOF
                    cmbDeprt.Items.Add(dRS.Fields("DepartmentNm").Value.ToString)
                    dRS.MoveNext()
                End While
            End If
            If cmbDeprt.Items.Count > 0 Then cmbDeprt.SelectedIndex = 1
        Else
            txtSec.Text = ""
        End If
        Call Button1_Click(sender, e)
    End Sub

    Private Sub cmbCeprt_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbDeprt.SelectedIndexChanged
        Dim dRS As New ADODB.Recordset
        If Lang = True Then
            Call LoadSqlData("Select * from Department Where DepartmentNmE= N'" & Trim(cmbDeprt.Text) & "' ", dRS)
            If dRS.RecordCount <> 0 Then
                txtDep.Text = Trim(dRS.Fields("DepartmentID").Value.ToString)
            End If
        Else
            Call LoadSqlData("Select * from Department Where DepartmentNm = N'" & Trim(cmbDeprt.Text) & "' ", dRS)
            If dRS.RecordCount <> 0 Then
                txtDep.Text = Trim(dRS.Fields("DepartmentID").Value.ToString)
            End If
        End If
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        myTemp = Trim(DG.Item(1, DG.CurrentRow.Index).Value.ToString())
        myTemp1 = Trim(DG.Item(3, DG.CurrentRow.Index).Value.ToString())

        Me.Close()
    End Sub

    Private Sub FG_DblClick(ByVal sender As Object, ByVal e As System.EventArgs)
        Call btnAdd_Click(sender, e)
    End Sub

    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub DG_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DG.CellContentClick

    End Sub

    Private Sub DG_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DG.DoubleClick
        myTemp = Trim(DG.Item(1, DG.CurrentRow.Index).Value.ToString())
        myTemp1 = Trim(DG.Item(3, DG.CurrentRow.Index).Value.ToString())

        Me.Close()
    End Sub
End Class