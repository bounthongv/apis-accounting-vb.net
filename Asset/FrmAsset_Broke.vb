Public Class FrmAsset_Broke
    Dim cn As New Odbc.OdbcConnection
    Dim mSql As String
 
    Private Sub FrmAcc_Code_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Call Loadlang()
        'SetControlText(Me)
        'ChgChildForm()
 
        'If mCompStr = "" Then btnAdd.Enabled = False
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
        'LdSec()

        cmbGrp.SelectedIndex = 0
        cmbSort.SelectedIndex = 0
        LoadDG()
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
    Private Sub LoadDG()
        'Dim DT As String = " AND Brokens.BrokenDate    BETWEEN '" & Format(ds.Value, "yyyy-MM-dd") & "' AND '" & Format(DT.Value, "yyyy-MM-dd") & "'"
        DG.DataSource = 0
        Dim ds As New DataSet
        Try
            ConnectCL()

            sql = "Select 0,BrokenID,BrokenDate,AssetID,group_id,AssetNm,Descriptions,Amt_KIP,Amount from Brokens where 1=1   AND Brokens.BrokenDate    BETWEEN '" & Format(Dst.Value, "yyyy-MM-dd") & "' AND '" & Format(Dt.Value, "yyyy-MM-dd") & "' " & mSql & " "
            'SQL = "Select SecID,SecNmL,SecNmE,Remark from Sections  " & mSql & " "
            LoadCN_DG()
            da.Fill(ds, " Brokens")
            DG.DataSource = ds.Tables(" Brokens")
            cn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try
        'DG.Columns(-1).Visible = False
        If Lang = True Then
            DG.Columns(0).HeaderText = "No " : DG.Columns(0).Width = "50"
            DG.Columns(1).HeaderText = "Depost ID" : DG.Columns(1).Width = "150"
            DG.Columns(2).HeaderText = "Date   " : DG.Columns(2).Width = "100"
            DG.Columns(3).HeaderText = "Asset ID" : DG.Columns(3).Width = "100"
            DG.Columns(4).HeaderText = "Group" : DG.Columns(4).Width = "100"
            DG.Columns(5).HeaderText = "Asset Name         " : DG.Columns(5).Width = "250"
            DG.Columns(6).HeaderText = "Description          " : DG.Columns(6).Width = "250"
            DG.Columns(7).HeaderText = "Remain     " : DG.Columns(7).Width = "150"
            DG.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DG.Columns(7).DefaultCellStyle.Format = "##,##0.00"
            DG.Columns(8).HeaderText = "Depost Values     " : DG.Columns(8).Width = "150"
            DG.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DG.Columns(8).DefaultCellStyle.Format = "##,##0.00"
            Dim counter As Integer = 0
            For Each dr As DataGridViewRow In DG.Rows
                counter += 1
            Next
            Label8.Text = "List All " & counter.ToString - 1 & " Items"
        Else
            DG.Columns(0).HeaderText = "ລ/ດ" : DG.Columns(0).Width = "50"
            DG.Columns(1).HeaderText = "ລະຫັດ" : DG.Columns(1).Width = "100"
            DG.Columns(2).HeaderText = "ວັນທີ   " : DG.Columns(2).Width = "100"
            DG.Columns(3).HeaderText = "ລະຫັດຊັບສິນ" : DG.Columns(3).Width = "100"
            DG.Columns(4).HeaderText = "ໝວດ" : DG.Columns(4).Width = "80"
            DG.Columns(5).HeaderText = "ຊື່ຊັບສິນ         " : DG.Columns(5).Width = "300"
            DG.Columns(6).HeaderText = "ເນື້ອໃນ          " : DG.Columns(6).Width = "250"
            DG.Columns(7).HeaderText = "ມູນຄ່າຍັງເຫຼືອ     " : DG.Columns(7).Width = "150"
            DG.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DG.Columns(7).DefaultCellStyle.Format = "##,##0.00"
            DG.Columns(8).HeaderText = "ມູນຄ່າສະສາງ     " : DG.Columns(8).Width = "150"
            DG.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DG.Columns(8).DefaultCellStyle.Format = "##,##0.00"
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
            cmbSec.Items.Add("** All Sections ***")
            Call LoadSqlData("Select * from Sections Order by SecID", sRS)
            If sRS.RecordCount <> 0 Then
                While Not sRS.EOF
                    cmbSec.Items.Add(sRS.Fields("SecNmE").Value.ToString)
                    sRS.MoveNext()
                End While
            End If
            If cmbSec.Items.Count > 0 Then cmbSec.SelectedIndex = 0
        Else
            cmbSec.Items.Add("** ສະແດງທັງໝົດ ***")
            Call LoadSqlData("Select * from Sections Order by SecID", sRS)
            If sRS.RecordCount <> 0 Then
                While Not sRS.EOF
                    cmbSec.Items.Add(sRS.Fields("SecNmL").Value.ToString)
                    sRS.MoveNext()
                End While
            End If
            If cmbSec.Items.Count > 0 Then cmbSec.SelectedIndex = 0
        End If
    End Sub
    Private Sub CndClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CndClose.Click
        Me.Close()
    End Sub

    Private Sub cmdDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub cmbSort_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSort.SelectedIndexChanged
        If cmbSec.SelectedIndex > 0 Then
            mSql = mSql & " AND Sect_ID='" & txtSec.Text & "'"
        End If
        If cmbDeprt.SelectedIndex > 0 Then
            mSql = mSql & " AND DepartmentID='" & txtDep.Text & "'"
        End If
        If cmbGrp.SelectedIndex <> 0 Then
            If cmbSort.SelectedIndex = 0 Then
                mSql = mSql & " AND group_id = '" & Trim(txtGrp.Text) & "'  "
            ElseIf cmbSort.SelectedIndex = 1 Then
                mSql = mSql & " AND group_id = '" & Trim(txtGrp.Text) & "' "
            Else
                mSql = mSql & " AND group_id = '" & Trim(txtGrp.Text) & "'  "
            End If
        Else
            'If cmbSort.SelectedIndex = 0 Then
            '    mSql = " Order by BrokenID"
            'ElseIf cmbSort.SelectedIndex = 1 Then
            '    mSql = " Order by AssetID"
            'Else
            '    mSql = " Order by AssetNm"
            'End If
        End If
        'Call LdData(mSql)
        LoadDG()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub cmbGrp_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbGrp.SelectedIndexChanged
        Dim gRS As New ADODB.Recordset
        Call LoadSqlData("select * from Groups Where Group_Nm=N'" & Trim(cmbGrp.Text) & "'", gRS)
        If gRS.RecordCount <> 0 Then
            txtGrp.Text = Trim(gRS.Fields("Group_ID").Value.ToString)
            mSql = " AND group_id = '" & Trim(txtGrp.Text) & "'  "
        Else
            txtGrp.Text = ""
            'mSql = " Order by BrokenID "
        End If
        'Call LdData(mSql)
        'LoadDG()
    End Sub

    Private Sub txtCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCode.KeyPress
        If e.KeyChar = Chr(13) Then
            mSql = " AND left(AssetID," & Len(txtCode.Text) & ")= '" & Trim(txtCode.Text) & "'"
            'Call LdData(mSql)
            LoadDG()
        End If
    End Sub

    Private Sub cmbSec_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSec.SelectedIndexChanged
        'Dim sRS As New ADODB.Recordset
        'If Lang = True Then
        '    Call LoadSqlData("select * from Sections Where SecNmE=N'" & Trim(cmbSec.Text) & "'", sRS)
        '    If sRS.RecordCount <> 0 Then
        '        txtSec.Text = Trim(sRS.Fields("SecID").Value.ToString)
        '        Dim dRS As New ADODB.Recordset
        '        cmbDeprt.Items.Clear()
        '        cmbDeprt.Items.Add("** All Department ***")
        '        Call LoadSqlData("Select * from Department Where Left(DepartmentID,2) = '" & Trim(txtSec.Text) & "' Order by DepartmentID", dRS)
        '        If dRS.RecordCount <> 0 Then
        '            While Not dRS.EOF
        '                cmbDeprt.Items.Add(dRS.Fields("DepartmentNmE").Value.ToString)
        '                dRS.MoveNext()
        '            End While
        '        End If
        '        If cmbDeprt.Items.Count > 0 Then cmbDeprt.SelectedIndex = 1
        '    Else
        '        txtSec.Text = ""
        '    End If
        'Else
        '    Call LoadSqlData("select * from Sections Where SecNmL=N'" & Trim(cmbSec.Text) & "'", sRS)
        '    If sRS.RecordCount <> 0 Then
        '        txtSec.Text = Trim(sRS.Fields("SecID").Value.ToString)
        '        Dim dRS As New ADODB.Recordset
        '        cmbDeprt.Items.Clear()
        '        cmbDeprt.Items.Add("** ສະແດງທັງໝົດ ***")
        '        Call LoadSqlData("Select * from Department Where Left(DepartmentID,2) = '" & Trim(txtSec.Text) & "' Order by DepartmentID", dRS)
        '        If dRS.RecordCount <> 0 Then
        '            While Not dRS.EOF
        '                cmbDeprt.Items.Add(dRS.Fields("DepartmentNm").Value.ToString)
        '                dRS.MoveNext()
        '            End While
        '        End If
        '        If cmbDeprt.Items.Count > 0 Then cmbDeprt.SelectedIndex = 1
        '    Else
        '        txtSec.Text = ""
        '    End If
        'End If

    End Sub

    Private Sub cmbDeprt_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbDeprt.SelectedIndexChanged
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

    Private Sub DG_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DG.CellContentClick

    End Sub

    Private Sub DG_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DG.MouseClick
        myTemp = Trim(DG.Item(1, DG.CurrentRow.Index).Value.ToString())
        AssetID = Trim(DG.Item(3, DG.CurrentRow.Index).Value.ToString())
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        mEdit = False
        FrmBrokeNew.ShowDialog()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        mEdit = True
        myTemp = Trim(DG.Item(1, DG.CurrentRow.Index).Value.ToString())
        FrmBrokeNew.ShowDialog()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel.Click
        If MsgBox("ທ່ານຕ້ອງການລຶບລາຍການຊັບສິນເລກລະຫັດ " & myTemp & " ນີ້ບໍ່?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2) = Windows.Forms.DialogResult.Yes Then
            CNN.Execute("Update Assets Set Deposted=0, AmountRemain=0, AmountClear=0, Amt_All=0, Deposted_Date=Null Where AssetID='" & AssetID & "' ")
            CNN.Execute("DELETE FROM Brokens WHERE BrokenID=N'" & myTemp & "' ")
            CNN.Execute("DELETE FROM Gen_jn WHERE certify=N'" & myTemp & "' ")
        End If
        Button1_Click(sender, e)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim Rs, Rs1 As New ADODB.Recordset
        Dim rpt As Object
        If Lang = False Then
            rpt = New CryAss_Broke
        Else
            rpt = New CryAss_BrokeEng
        End If

        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim str As String = ""
        If cmbGrp.SelectedIndex <> 0 Then
            If cmbSort.SelectedIndex = 0 Then
                str = " AND group_id = '" & Trim(txtGrp.Text) & "' Order by BrokenID"
            ElseIf cmbSort.SelectedIndex = 1 Then
                str = " AND group_id = '" & Trim(txtGrp.Text) & "' Order by AssetID"
            Else
                str = " AND group_id = '" & Trim(txtGrp.Text) & "' Order by AssetNm"
            End If
        Else
            If cmbSort.SelectedIndex = 0 Then
                str = " Order by BrokenID"
            ElseIf cmbSort.SelectedIndex = 1 Then
                str = " Order by AssetID"
            Else
                str = " Order by AssetNm"
            End If
        End If
        Call LoadSqlData("Select * from Brokens Where 1=1 " & mSql & " ", Rs)
        CNN.Execute("UPDATE Brokens set Brokens.Asset_NmE=Assets.Asset_NmE from Assets,Brokens where Brokens.AssetID=Assets.AssetID ")
        If Rs.RecordCount = 0 Then
            MsgBox("No data")
            Exit Sub
        End If

        With rpt
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
            '    FmPreview.ReportViewer.ReportSource = SubDoc
            'End If
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
        mSql = ""

        If txtCode.Text = "" Then
            If cmbSort.SelectedIndex = 1 Then
                mSql = mSql & " AND group_id = '" & Trim(txtGrp.Text) & "'  "
            ElseIf cmbSort.SelectedIndex = 2 Then
                mSql = mSql & " AND group_id = '" & Trim(txtGrp.Text) & "'  "
            ElseIf cmbSort.SelectedIndex = 3 Then
                mSql = mSql & " AND group_id = '" & Trim(txtGrp.Text) & "'   "
            End If
        Else
            mSql = mSql & " AND AssetID= '" & Trim(txtCode.Text) & "'"
        End If
        If cmbSec.SelectedIndex > 0 Then
            mSql = mSql & " AND Sect_ID='" & txtSec.Text & "'"
        End If
        If cmbDeprt.SelectedIndex > 0 Then
            mSql = mSql & " AND DepartmentID='" & txtDep.Text & "'"
        End If
        If cmbGrp.SelectedIndex > 0 Then
            mSql = mSql & " AND group_id='" & txtGrp.Text & "'"
        End If
        'Call LdData(mSql)
        LoadDG()
    End Sub

    Private Sub Dt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Dt.ValueChanged

    End Sub
End Class