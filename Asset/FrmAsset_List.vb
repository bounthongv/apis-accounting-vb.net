Public Class FrmAsset_List
    Dim cn As New Odbc.OdbcConnection
    Dim mSql As String
    Dim str, DDD As String
    Private Sub FrmAcc_Code_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Call Loadlang()
        'SetControlText(Me)
        'ChgChildForm()
        'LoadLng2()
        'SetControlText2(Me)
        'ChgChildForm2()
        'ສະແດງທັງໝົດ()
        'ລາຍການພວມນຳໃຊ້()
        'ລາຍການສະສາງແລ້ວ()
        Call LoadLuangG()
        'mSql = " Order by Asset_No"
        Call LdGrp()
        Call LdSec()
        'Call LdDep()
        cmbGrp.SelectedIndex = 0
        cmbSec.SelectedIndex = 0
        cmbSort.SelectedIndex = 1
        CmbShow.SelectedIndex = 0
        LoadDG()
        'If mCompStr = "" Then btnAdd.Enabled = False
        'Call loadUSR()
    End Sub
    Private Sub loadUSR()
        '===============Add/Save==========
        If MDWrite = 1 Then
            btnAdd.Enabled = True
        Else
            btnAdd.Enabled = False
        End If
        '===============Edit/Save==========
        If MDEdit = 1 Then
            btnEdit.Enabled = True
        Else
            btnEdit.Enabled = False
        End If
        '===============Delete==========
        If MDDelete = 1 Then
            btnDel.Enabled = True
        Else
            btnDel.Enabled = False
        End If
    End Sub

    Private Sub LoadLuangG()
        If Lang = True Then
            'FG.FormatString = "^No. |<Asset ID   |<Number No.    |^Section |<Asset Name                   |^Group |^Used Date     |^By Use|>Value Price     |^Currency|>Rate  |>Assets Amount   |^Deposal Date |^ Qty.     "
            CmbShow.Items.Clear()
            CmbShow.Items.Add("Show All")
            CmbShow.Items.Add("Uesd List")
            CmbShow.Items.Add("Deposal List")
            Label7.Text = "List"
            Label1.Text = "Group"
            Label2.Text = "Code"
            Label5.Text = "Code"
            Label3.Text = "Asset Name"
            lSort.Text = "Order"
            chkDT.Text = "Use Date"
            CheckBox1.Text = "Land"
        Else
            CheckBox1.Text = "ສະເພາະທີ່ດິນ"
            Label1.Text = "ໝວດ"
            Label2.Text = "ລະຫັດ"
            Label5.Text = "ລະຫັດ"
            Label3.Text = "ຊື່ຊັບສິນ"
            lSort.Text = "ລຽງຕາມ"
            Label7.Text = "ລາຍການ"
            chkDT.Text = "ດ.ປ ນໍາໃຊ້"
            'FG.FormatString = "^ລ/ດ |<ລະຫັດຊັບສິນ   |<ເລກປະຈຳຕົວ     |^ພາກສ່ວນ |<ຊື່ຊັບສິນ                           |^ໝວດ |^ວັນທີນຳໃຊ້     |^ອາຍຸນຳໃຊ້|>ມູນຄ່າເດີມ     |^ສະກຸນເງິນ|>ອັດຕາແລກປ່ຽນ|>ມູນຄ່າຊັບສິນ    |^ວັນທີສະສາງ   |^ຈຳນວນ   "
            CmbShow.Items.Clear()
            CmbShow.Items.Add("ສະແດງທັງໝົດ")
            CmbShow.Items.Add("ລາຍການພວມນຳໃຊ້")
            CmbShow.Items.Add("ລາຍການສະສາງແລ້ວ")
        End If
    End Sub
    Private Sub LdDep()
        Dim sRS As New ADODB.Recordset
        cmbDeprt.Items.Clear()
        If Lang = True Then
            cmbDeprt.Items.Add("** All Department ***")
            Call LoadSqlData("Select * from Department Order by DepartmentID", sRS)
            If sRS.RecordCount <> 0 Then
                While Not sRS.EOF
                    cmbDeprt.Items.Add(sRS.Fields("DepartmentNmE").Value.ToString)
                    sRS.MoveNext()
                End While
            End If
            If cmbDeprt.Items.Count > 0 Then cmbDeprt.SelectedIndex = 0
        Else
            cmbDeprt.Items.Add("** ສະແດງທັງໝົດ ***")
            Call LoadSqlData("Select * from Department Order by DepartmentID", sRS)
            If sRS.RecordCount <> 0 Then
                While Not sRS.EOF
                    cmbDeprt.Items.Add(sRS.Fields("DepartmentNm").Value.ToString)
                    sRS.MoveNext()
                End While
            End If
            If cmbDeprt.Items.Count > 0 Then cmbDeprt.SelectedIndex = 0
        End If

    End Sub
    Private Sub LdSec()
        Dim sRS As New ADODB.Recordset
        cmbSec.Items.Clear()
        If Lang = True Then
            cmbSec.Items.Add("**ທັງໝົດ**")
            Call LoadSqlData("Select * from AP_Office  where Off_ID<>'00' Order by Off_ID", sRS)
            If sRS.RecordCount <> 0 Then
                While Not sRS.EOF
                    cmbSec.Items.Add(sRS.Fields("Off_Name").Value.ToString)
                    sRS.MoveNext()
                End While
            End If
            If cmbSec.Items.Count > 0 Then cmbSec.SelectedIndex = 0
        Else
            If Mpermiss = "Admin" Then
                cmbSec.Items.Add("** ສະແດງທັງໝົດ **")
                Call LoadSqlData("Select * from AP_Office  where Off_ID<>'00' Order by Off_ID", sRS)
                If sRS.RecordCount <> 0 Then
                    While Not sRS.EOF
                        cmbSec.Items.Add(sRS.Fields("Off_Name").Value.ToString)
                        sRS.MoveNext()
                    End While
                End If
                If cmbSec.Items.Count > 0 Then cmbSec.SelectedIndex = 0
            Else
                cmbSec.Items.Clear()
                Call load_Cmb("SELECT Off_Name FROM AP_Office where 1=1 and off_id=N'" & Off_Id & "' ORDER BY off_id ASC", "Off_Name", cmbSec)
                If cmbSec.Items.Count > 0 Then
                    cmbSec.SelectedIndex = 0
                End If
            End If

        End If
    End Sub

    Private Sub LdGrp()
        Dim gRS As New ADODB.Recordset
        cmbGrp.Items.Clear()
        If Lang = True Then
            cmbGrp.Items.Add("** All Groups_Asset ***")
            Call LoadSqlData("Select * from Groups_Asset Order by Group_ID", gRS)
            If gRS.RecordCount <> 0 Then
                While Not gRS.EOF
                    cmbGrp.Items.Add(gRS.Fields("Group_NmE").Value.ToString)
                    gRS.MoveNext()
                End While
            End If
            cmbGrp.SelectedIndex = 0
        Else
            cmbGrp.Items.Add("** ສະແດງທັງໝົດ ***")
            Call LoadSqlData("Select * from Groups_Asset Order by Group_ID", gRS)
            If gRS.RecordCount <> 0 Then
                While Not gRS.EOF
                    cmbGrp.Items.Add(gRS.Fields("Group_Nm").Value.ToString)
                    gRS.MoveNext()
                End While
            End If
            cmbGrp.SelectedIndex = 0
        End If

    End Sub
    Private Sub CndClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CndClose.Click
        'ShowShortFrm()
        Me.Close()
    End Sub

    Private Sub FrmCategory_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        'FG.Width = Me.Width - 50
        'FG.Height = Me.Height - 120
    End Sub

    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Call btnEdit_Click(sender, e)
    End Sub

    Private Sub cmbSort_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSort.SelectedIndexChanged
        If cmbSort.SelectedIndex = 0 Then
            mSql = " Order by Sect_ID, AssetID"
        ElseIf cmbSort.SelectedIndex = 1 Then
            mSql = " Order by Asset_Nm"
        Else
            mSql = " Order by Date_Work"
        End If
        'Call Button1_Click(sender, e)
    End Sub

    Private Sub cmdDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub MDOFF()
        If Lang = False Then
        Else

        End If
    End Sub
    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub cmbGrp_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbGrp.SelectedIndexChanged
        Dim gRS As New ADODB.Recordset
        Call LoadSqlData("select * from Groups_Asset Where Group_Nm=N'" & Trim(cmbGrp.Text) & "'", gRS)
        If gRS.RecordCount <> 0 Then
            txtGrp.Text = Trim(gRS.Fields("Group_ID").Value.ToString)
        Else
            txtGrp.Text = ""
        End If
        'Call Button1_Click(sender, e)
    End Sub
 
    Private Sub LoadDGEng()
        Dim cRS As New ADODB.Recordset
        DDD = ""
        If CmbShow.SelectedIndex = 0 Then
            DDD = DDD
        ElseIf CmbShow.SelectedIndex = 1 Then
            DDD = DDD & " AND  Deposted='0' "
        Else
            DDD = DDD & " AND  Deposted='1'"
        End If
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
            sql = "Select 0,AssetID,Asset_No,Asset_Nm,Group_ID,Date_Work,Used_Life,Amount,Curr,Rate,txt_disc,txt_Rem,Amt_KIP,Deposted_Date,Qty from ASSETS where 1=1 " & DDD & " "
            'SQL = "Select SecID,SecNmL,SecNmE,Remark from Sections  " & mSql & " "
            LoadCN_DG()
            da.Fill(ds, " ASSETS")
            DG.DataSource = ds.Tables(" ASSETS")
            cn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try

        DG.Columns(0).HeaderText = "No " : DG.Columns(0).Width = "40"
        DG.Columns(1).HeaderText = "Asset ID" : DG.Columns(1).Width = "100"
        DG.Columns(2).HeaderText = "Number No" : DG.Columns(2).Width = "150"
        DG.Columns(3).HeaderText = "Asset Name      " : DG.Columns(3).Width = "250"
        DG.Columns(4).HeaderText = "Group" : DG.Columns(4).Width = "50"
        DG.Columns(5).HeaderText = "Used Date" : DG.Columns(5).Width = "90"
        DG.Columns(6).HeaderText = "Life " : DG.Columns(6).Width = "80"
        DG.Columns(7).HeaderText = "Original " : DG.Columns(7).Width = "125"
        DG.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DG.Columns(7).DefaultCellStyle.Format = "##,##0.00"
        DG.Columns(8).HeaderText = "Currency" : DG.Columns(8).Width = "70"
        DG.Columns(9).HeaderText = "Exchange" : DG.Columns(9).Width = "80"
        DG.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DG.Columns(9).DefaultCellStyle.Format = "##,##0.00"
        DG.Columns(10).HeaderText = "ຄ່າຊາກ 0.01%" : DG.Columns(10).Width = "125"
        DG.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DG.Columns(10).DefaultCellStyle.Format = "##,##0.00"
        DG.Columns(11).HeaderText = "ມູນຄ່າຮັບຮູ້ເບື້ອງຕົ້ນ" : DG.Columns(11).Width = "125"
        DG.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DG.Columns(11).DefaultCellStyle.Format = "##,##0.00"
        DG.Columns(12).HeaderText = "Amount" : DG.Columns(12).Width = "125"
        DG.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DG.Columns(12).DefaultCellStyle.Format = "##,##0.00"
        DG.Columns(13).HeaderText = "Depost Date" : DG.Columns(13).Width = "90"
        DG.Columns(14).HeaderText = "Quatity" : DG.Columns(12).Width = "50"
        Dim counter As Integer = 0
        For Each dr As DataGridViewRow In DG.Rows
            counter += 1
        Next
        Label8.Text = "List All " & counter.ToString - 1 & " Items"

        For i As Integer = 0 To DG.Rows.Count - 1
            DG.Rows(i).HeaderCell.Value = i.ToString()
        Next i
        Dim row As Integer = 0
        For row = 0 To DG.RowCount - 2
            DG.Rows(row).Cells(0).Value = row + 1
        Next
        DG.ReadOnly = True
    End Sub
    Private Sub LoadDG()
        Dim cRS As New ADODB.Recordset
        DDD = ""
        If CmbShow.SelectedIndex = 0 Then
            DDD = DDD
        ElseIf CmbShow.SelectedIndex = 1 Then
            DDD = DDD & " AND  Deposted='0' "
        ElseIf CmbShow.SelectedIndex = 2 Then
            DDD = DDD & " AND  Deposted='1'"
        End If
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
        '==============Post
        'Dim PP As String
        'If ComboBox1.SelectedIndex = 0 Then
        '    DDD = DDD
        'ElseIf ComboBox1.SelectedIndex = 1 Then
        '    DDD = DDD & " AND  Deposted='0' "
        'ElseIf ComboBox1.SelectedIndex = 2 Then
        '    DDD = DDD & " AND  Deposted='1'"
        'End If
        DG.DataSource = 0
        Dim ds As New DataSet
        Try
            ConnectCL()
            sql = "Select 0,AssetID,Asset_No,Asset_Nm,Group_ID,Date_Work,Used_Life,Amount,Curr,Rate,txt_disc,txt_Rem,Amt_KIP,Deposted_Date,Qty from ASSETS where 1=1 " & DDD & " "
            'SQL = "Select SecID,SecNmL,SecNmE,Remark from Sections  " & mSql & " "
            LoadCN_DG()
            da.Fill(ds, " ASSETS")
            DG.DataSource = ds.Tables(" ASSETS")
            cn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try
        If Lang = True Then
            DG.Columns(0).HeaderText = "No " : DG.Columns(0).Width = "40"
            DG.Columns(1).HeaderText = "Asset ID" : DG.Columns(1).Width = "100"
            DG.Columns(2).HeaderText = "Number No" : DG.Columns(2).Width = "150"
            DG.Columns(3).HeaderText = "Asset Name      " : DG.Columns(3).Width = "250"
            DG.Columns(4).HeaderText = "Group" : DG.Columns(4).Width = "50"
            DG.Columns(5).HeaderText = "Used Date" : DG.Columns(5).Width = "90"
            DG.Columns(6).HeaderText = "Life " : DG.Columns(6).Width = "80"
            DG.Columns(7).HeaderText = "Original " : DG.Columns(7).Width = "125"
            DG.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DG.Columns(7).DefaultCellStyle.Format = "##,##0.00"
            DG.Columns(8).HeaderText = "Currency" : DG.Columns(8).Width = "70"
            DG.Columns(9).HeaderText = "Exchange" : DG.Columns(9).Width = "80"
            DG.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DG.Columns(9).DefaultCellStyle.Format = "##,##0.00"
            DG.Columns(10).HeaderText = "ຄ່າຊາກ 0.01%" : DG.Columns(10).Width = "125"
            DG.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DG.Columns(10).DefaultCellStyle.Format = "##,##0.00"
            DG.Columns(11).HeaderText = "ມູນຄ່າຮັບຮູ້ເບື້ອງຕົ້ນ" : DG.Columns(11).Width = "125"
            DG.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DG.Columns(11).DefaultCellStyle.Format = "##,##0.00"
            DG.Columns(12).HeaderText = "Amount" : DG.Columns(12).Width = "125"
            DG.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DG.Columns(12).DefaultCellStyle.Format = "##,##0.00"
            DG.Columns(13).HeaderText = "Depost Date" : DG.Columns(13).Width = "90"
            DG.Columns(14).HeaderText = "Quatity" : DG.Columns(14).Width = "50"
            Dim counter As Integer = 0
            For Each dr As DataGridViewRow In DG.Rows
                counter += 1
            Next
            Label8.Text = "List All " & counter.ToString - 1 & " Items"

        Else
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
            'DG.Columns(10).HeaderText = "ຄ່າຊາກ 0.01%" : DG.Columns(10).Width = "125"
            'DG.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'DG.Columns(10).DefaultCellStyle.Format = "##,##0.00"
            DG.Columns(10).HeaderText = "ຄ່າຊາກ 0.01%" : DG.Columns(10).Width = "125"
            DG.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DG.Columns(10).DefaultCellStyle.Format = "##,##0.00"
            DG.Columns(11).HeaderText = "ມູນຄ່າຮັບຮູ້ເບື້ອງຕົ້ນ" : DG.Columns(11).Width = "125"
            DG.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DG.Columns(11).DefaultCellStyle.Format = "##,##0.00"
            DG.Columns(12).HeaderText = "ມູນຄ່າເງິນກີບ" : DG.Columns(12).Width = "125"
            DG.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DG.Columns(12).DefaultCellStyle.Format = "##,##0.00"
            DG.Columns(13).HeaderText = "ວັນທີສະສາງ" : DG.Columns(13).Width = "90"
            DG.Columns(14).HeaderText = "ຈຳນວນ" : DG.Columns(12).Width = "50"
            Dim counter As Integer = 0
            For Each dr As DataGridViewRow In DG.Rows
                counter += 1
            Next
            Label8.Text = "ລາຍການທັງໝົດ " & counter.ToString - 1 & " ລາຍການ"

        End If

        For i As Integer = 0 To DG.Rows.Count - 1
            DG.Rows(i).HeaderCell.Value = i.ToString()
        Next i
        Dim row As Integer = 0
        For row = 0 To DG.RowCount - 2
            DG.Rows(row).Cells(0).Value = row + 1
        Next
        DG.ReadOnly = True
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        mEdit = False
        FrmAssetNew.ShowDialog()
    End Sub

    Private Sub txtCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCode.KeyPress
        If e.KeyChar = Chr(13) Then
            'Call Button1_Click(sender, e)
        End If
    End Sub

    Private Sub txtNm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNm.KeyPress
        If e.KeyChar = Chr(13) Then
            'Call Button1_Click(sender, e)
        End If
    End Sub

    Private Sub DTUSE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTUSE.ValueChanged
        'Call Button1_Click(sender, e)
    End Sub

    Private Sub chkDT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDT.CheckedChanged
        If chkDT.Checked = True Then
            DTUSE.Enabled = True
        Else
            DTUSE.Enabled = False
        End If
    End Sub

    Private Sub cmbSec_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSec.SelectedIndexChanged
        Dim sRS As New ADODB.Recordset
        If Lang = True Then
            Call LoadSqlData("select * from AP_Office Where Off_NameE=N'" & Trim(cmbSec.Text) & "'", sRS)
            If sRS.RecordCount <> 0 Then
                txtSec.Text = Trim(sRS.Fields("Off_ID").Value.ToString)
                Dim dRS As New ADODB.Recordset
                cmbDeprt.Items.Clear()
                cmbDeprt.Items.Add("** All Department ***")
                Call LoadSqlData("Select * from Department Where Left(DepartmentID,2) = '" & Trim(txtSec.Text) & "' Order by DepartmentID", dRS)
                If dRS.RecordCount <> 0 Then
                    While Not dRS.EOF
                        cmbDeprt.Items.Add(dRS.Fields("DepartmentNmE").Value.ToString)
                        dRS.MoveNext()
                    End While
                End If
                If cmbDeprt.Items.Count > 0 Then cmbDeprt.SelectedIndex = 1
            Else
                txtSec.Text = ""
            End If
        Else
            Call LoadSqlData("select * from AP_Office Where Off_Name=N'" & Trim(cmbSec.Text) & "'", sRS)
            If sRS.RecordCount <> 0 Then
                txtSec.Text = Trim(sRS.Fields("Off_ID").Value.ToString)
                Dim dRS As New ADODB.Recordset
                cmbDeprt.Items.Clear()
                cmbDeprt.Items.Add("** ສະແດງທັງໝົດ ***")
                Call LoadSqlData("Select * from Department Where Left(DepartmentID,2) = '" & Trim(txtSec.Text) & "' Order by DepartmentID", dRS)
                If dRS.RecordCount <> 0 Then
                    While Not dRS.EOF
                        cmbDeprt.Items.Add(dRS.Fields("DepartmentNm").Value.ToString)
                        dRS.MoveNext()
                    End While
                End If
                'If cmbDeprt.Items.Count > 0 Then cmbDeprt.SelectedIndex = 1
            Else
                txtSec.Text = ""
            End If
        End If

    End Sub

    Private Sub cmbCeprt_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbDeprt.SelectedIndexChanged
        Dim dRS As New ADODB.Recordset
        If Lang = True Then
            Call LoadSqlData("Select * from Department Where DepartmentNmE= N'" & Trim(cmbDeprt.Text) & "' ", dRS)
            If dRS.RecordCount <> 0 Then
                txtDep.Text = Trim(dRS.Fields("DepartmentID").Value.ToString)
            End If
        Else
            Call LoadSqlData("Select * from Department Where DepartmentNm = N'" & Trim(cmbDeprt.Text) & "' and company='" & txtSec.Text & "' ", dRS)
            If dRS.RecordCount <> 0 Then
                txtDep.Text = Trim(dRS.Fields("DepartmentID").Value.ToString)
            End If
        End If

    End Sub

    Private Sub txtCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCode.TextChanged

    End Sub

    Private Sub txtNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNo.TextChanged

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        EditActive = True
        mEdit = True
        If Lang = True Then
            If myTemp = "" Then MsgBox("Plase Select Assets List") : Exit Sub
        Else
            If myTemp = "" Then MsgBox("ກະລຸນາເລືອກລາຍການທີ່ທ່ານຕ້ອງການຈະໂອນກ່ອນ") : Exit Sub
        End If
        'Frm_List_History.ShowDialog()
    End Sub


    Private Sub DG_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DG.DoubleClick
        If myTemp = "" Then
            Exit Sub
        End If
        myTemp = Trim(DG.Item(1, DG.CurrentRow.Index).Value.ToString())
        Call btnEdit_Click(sender, e)
    End Sub

    Private Sub DG_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DG.MouseClick
        'If myTemp = "" Then
        '    Exit Sub
        'End If
        Try


     
            myTemp = Trim(DG.Item(1, DG.CurrentRow.Index).Value.ToString())
            AssetID = myTemp
            ASName = Trim(DG.Item(4, DG.CurrentRow.Index).Value.ToString())
            'Call LoadSqlData("select * from assets where AssetID='" & myTemp & "'", RSC)
            'If RSC.RecordCount > 0 Then
            '    AANO = RSC.Fields("certify").Value
            'End If
            'Dim row As Integer = 0
            'For row = 0 To DG.RowCount - 2
            '    DG.Rows(row).Cells(0).Value = row + 1
            'Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CmbShow_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbShow.SelectedIndexChanged

    End Sub

    Private Sub DG_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DG.CellContentClick
        DG.ReadOnly = True
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        mEdit = False
        FrmAssetNew.ShowDialog()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        'If myTemp = "" Then
        '    Exit Sub
        'End If
        mEdit = True
        myTemp = Trim(DG.Item(1, DG.CurrentRow.Index).Value.ToString())
        FrmAssetNew.ShowDialog()
        FrmAssetNew.Focus()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel.Click
        If myTemp = "" Then
            Exit Sub
        End If
        Call LoadSqlData("SELECT * FROM Brokens where AssetID='" & myTemp & "' ", RSC)
        If RSC.RecordCount > 0 Then
            MsgBox("ທ່ານບໍ່ສາມາດລຶບລາຍການນີ້ໄດ້ເພາະມີສະສາງແລ້ວ", MsgBoxStyle.Exclamation) : Exit Sub
        End If

        If MsgBox("ທ່ານຕ້ອງການລຶບລາຍການຊັບສິນເລກລະຫັດ " & myTemp & " ນີ້ບໍ່?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2) = Windows.Forms.DialogResult.Yes Then
            CNN.Execute("DELETE FROM Assets WHERE AssetID='" & myTemp & "' ")
            CNN.Execute("DELETE FROM gen_jn where Book='Fixd Asset' and certify=N'" & Trim(myTemp) & "' ")

            'Call Button1_Click(sender, e)
        End If
    End Sub
    Private Sub OfficeNEW()
        Dim Rs As New ADODB.Recordset
        With Rs
            Call LoadSqlData("SELECT * FROM AP_Office where off_id='" & txtSec.Text & "' ", Rs)
            If .RecordCount = 0 Then Exit Sub
            OffName = Trim(.Fields("off_Nm").Value.ToString)
            OffNameE = Trim(.Fields("off_NmE").Value.ToString)
            Off_strtl = Trim(.Fields("off_strtl").Value.ToString)
            Off_VillageL = Trim(.Fields("Off_VillageL").Value.ToString)
            Off_DistL = Trim(.Fields("Off_DistL").Value.ToString)
            Off_ProVL = Trim(.Fields("Off_ProVL").Value.ToString)
            OffTel = Trim(.Fields("tel").Value.ToString)
            OffFax = Trim(.Fields("fax").Value.ToString)
            Sign1 = Trim(.Fields("Signatur_1").Value.ToString)
            Sign2 = Trim(.Fields("Signatur_2").Value.ToString)
            Sign3 = Trim(.Fields("Signatur_3").Value.ToString)
            Sign4 = Trim(.Fields("Signatur_4").Value.ToString)
            Sign5 = Trim(.Fields("Signatur_5").Value.ToString)
            OffPlace = Trim(.Fields("Locate_Bill").Value.ToString)
            .MoveNext()
        End With
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Call Office()
        OfficeNEW()
        Dim Rs, Rs1 As New ADODB.Recordset
        Dim rpt As Object
        If Lang = False Then
            If CheckBox1.Checked = True Then
                rpt = New CryAss_ListLand
            Else
                rpt = New CryAss_List
            End If

        Else
            rpt = New CryAss_ListEng
        End If
        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim cRS As New ADODB.Recordset
        Dim str As String = ""
        Dim ss As String
        DDD = ""
        If CmbShow.SelectedIndex = 0 Then
            DDD = DDD
        ElseIf CmbShow.SelectedIndex = 1 Then
            DDD = DDD & " AND  Deposted='0' "
        ElseIf CmbShow.SelectedIndex = 2 Then
            DDD = DDD & " AND  Deposted='1'"
        End If
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
        'ss = "Select * from Assets Where 1=1" & str & " Order by AssetID"
        'Call LoadSqlData(ss, Rs)
        Dim kkk As String
        If CmbShow.Text = "ສະແດງທັງໝົດ" Then
            kkk = ""
        ElseIf CmbShow.Text = "ລາຍການພວມນຳໃຊ້" Then
            kkk = " AND  Deposted='0' "
        Else
            kkk = " AND  Deposted='1'"
        End If

        If CheckBox1.Checked = True Then
            ss = "Select * from ASSETS where 1=1 " & DDD & " and Group_ID='201' Order by AssetID"
            Call LoadSqlData(ss, Rs)
        Else
            ss = "Select * from ASSETS where 1=1 " & DDD & " " & kkk & " Order by AssetID"
            Call LoadSqlData(ss, Rs)
        End If


        If Rs.RecordCount = 0 Then
            MsgBox("Data Emtry") : Exit Sub
        End If
        With rpt
            '.SetDataSource(Rs)


            Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
            If CheckBox1.Checked = True Then
                myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("OffNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = OffName
                myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Tel"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = "Tel: " & OffTel & "Fax: " & OffFax
                myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = Sign1
                myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg2"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = Sign2
                myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg3"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = Sign3
                myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg4"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = Sign4
                myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg5"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = Sign5
                myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("place"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = PlaecL

                Dim RO As CrystalDecisions.CrystalReports.Engine.ReportObject
                Dim SRO As CrystalDecisions.CrystalReports.Engine.SubreportObject
                Dim SubDoc As CrystalDecisions.CrystalReports.Engine.ReportDocument
                SqlPrint = "SELECT  * from Ap_Image  where Img_Id='" & IMageID & "'"
                Call LoadSqlData(SqlPrint, Rs1)
                RO = rpt.ReportDefinition.Sections.Item("Section1").ReportObjects.Item("Subreport1")
                SRO = CType(RO, CrystalDecisions.CrystalReports.Engine.SubreportObject)
                SubDoc = SRO.OpenSubreport(SRO.SubreportName)
                If Rs1.RecordCount > 0 Then
                    SubDoc.SetDataSource(Rs1)
                    FmPreview.ReportViewer.ReportSource = SubDoc
                End If
            Else
                Dim myText1 As CrystalDecisions.CrystalReports.Engine.TextObject
                myText1 = CType(rpt.ReportDefinition.ReportObjects.Item("txtSec"), CrystalDecisions.CrystalReports.Engine.TextObject)
                If cmbSec.SelectedIndex = 0 Then
                    myText1.Text = Trim(cmbSec.Text)
                Else
                    myText1.Text = "ພາກສ່ວນ " & Trim(cmbSec.Text)
                End If
                'Dim RO As CrystalDecisions.CrystalReports.Engine.ReportObject
                'Dim SRO As CrystalDecisions.CrystalReports.Engine.SubreportObject
                'Dim SubDoc As CrystalDecisions.CrystalReports.Engine.ReportDocument
                'SqlPrint = "SELECT  * from Ap_Image  where Img_Id='" & IMageID & "'"
                'Call LoadSqlData(SqlPrint, Rs1)
                'RO = rpt.ReportDefinition.Sections.Item("Section1").ReportObjects.Item("Subreport3")
                'SRO = CType(RO, CrystalDecisions.CrystalReports.Engine.SubreportObject)
                'SubDoc = SRO.OpenSubreport(SRO.SubreportName)
                'If Rs1.RecordCount > 0 Then
                '    SubDoc.SetDataSource(Rs1)
                '    FmPreview.CrystalReportViewer1.ReportSource = SubDoc
                'End If

                If Lang = False Then
                    myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("OffNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText2.Text = OffName
                    myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Tel"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText2.Text = "Tel: " & OffTel & "Fax: " & OffFax
                    myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText2.Text = Sign1
                    myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg2"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText2.Text = Sign2
                    myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg3"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText2.Text = Sign3
                    myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg4"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText2.Text = Sign4
                    myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg5"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText2.Text = Sign5
                    myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("place"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText2.Text = PlaecL
                Else
                    myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("OffNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText2.Text = OffNameE
                    myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Tel"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText2.Text = "Tel: " & OffTel & "Fax: " & OffFax
                    myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText2.Text = Sign1e
                    myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg2"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText2.Text = Sign2e
                    myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg3"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText2.Text = Sign3e
                    myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg4"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText2.Text = Sign3e
                    myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg5"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText2.Text = Sign5e
                    myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("place"), CrystalDecisions.CrystalReports.Engine.TextObject)
                    myText2.Text = PlaecE
                End If
            End If
            rpt.SetDataSource(Rs)
            rpt.Refresh()
            FmPreview.ReportViewer.ReportSource = rpt
            FmPreview.ReportViewer.DisplayGroupTree = False
            FmPreview.WindowState = FormWindowState.Maximized
            FmPreview.Show()
            'Call CloseRs(RSC)
            'Call CloseRs(Rs)
        End With
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Lang = False Then
            LoadDG()
        Else
            LoadDGEng()
        End If
        LoadLuangG()
    End Sub

    Private Sub txtNm_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class