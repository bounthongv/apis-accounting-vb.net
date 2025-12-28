Public Class FrmAsset_List_LSNEW
    Dim cn As New Odbc.OdbcConnection
    Dim mSql, GrpID As String

    Private Sub FrmAcc_Code_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Call Loadlang()

        'If Lang = False Then
        '    FG.FormatString = "^ລ/ດ |<ລະຫັດໝວດ |<ຊື່ໝວດຊັບສິນ  (ພາສາລາວ)                       |<ຊື່ໝວດຊັບສິນ (ພາສາອັງກິດ)                      |<ເລກບັນຊີຊັບສິນ|< ບັນຊີຊື້   |<ບັນຊີຄິດໄລ່ຫຼຸຍຫ້ຽນ|<ບັນຊີລາຍຈ່າຍຄ່າຫຼຸຍຫ້ຽນ|< ບັນຊີສະສາງ    "
        'Else
        '    FG.FormatString = "^No. |<Group Code|<Group Name (Lao)                             |<Group Name (English)                        |<ເລກບັນຊີຊັບສິນ|< ບັນຊີຊື້   |<ບັນຊີຄິດໄລ່ຫຼຸຍຫ້ຽນ|<ບັນຊີລາຍຈ່າຍຄ່າຫຼຸຍຫ້ຽນ|< ບັນຊີສະສາງ    "

        'End If
        mSql = " Order by Group_ID"
        Call LoadDG()
        'If mCompStr = "" Then btnAdd.Enabled = False
        'Call loadUSR()
    End Sub
    Private Sub loadUSR()
        '===============Add/Save==========
        'If MDWrite = 1 Then
        '    btnAdd.Enabled = True
        'Else
        '    btnAdd.Enabled = False
        'End If
        '===============Edit/Save==========
        If MDEdit = 1 Then
            btnEdit.Enabled = True
        Else
            btnEdit.Enabled = False
        End If
        '===============Delete==========
        'If MDDelete = 1 Then
        '    btnDel.Enabled = True
        'Else
        '    btnDel.Enabled = False
        'End If
    End Sub
    Private Sub CndClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CndClose.Click
        'ShowShortFrm()
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
        If cmbSort.SelectedIndex = 0 Then
            mSql = " Order by Group_ID"
        ElseIf cmbSort.SelectedIndex = 1 Then
            mSql = " Order by Group_Nm"
        Else
            mSql = " Order by Ac_Code"
        End If
        Call LoadDG()
    End Sub
    Private Sub cmbSort_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub LoadDG()
        DG.DataSource = 0
        Dim ds As New DataSet
        Try
            ConnectCL()
            sql = "Select 0,Group_ID,Group_Nm,Group_NmE,AccountCodeAsDR,AccountCodeAsCR,Ac_Code,Dep_Code,AccountCodeBrokenDR,Grp_no from Groups_Asset  " & mSql & " "
            LoadCN_DG()
            da.Fill(ds, "Groups_Asset")
            DG.DataSource = ds.Tables("Groups_Asset")
            cn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try

        'DG.Columns(-1).Visible = False
        If Lang = True Then
            DG.Columns(0).HeaderText = "No " : DG.Columns(0).Width = "50"
            DG.Columns(1).HeaderText = "Code " : DG.Columns(1).Width = "90"
            DG.Columns(2).HeaderText = "Group Name  (Lao)      " : DG.Columns(2).Width = "300"
            DG.Columns(3).HeaderText = "Group Name  (Eng)      " : DG.Columns(3).Width = "300"
            DG.Columns(4).HeaderText = "Asset Account " : DG.Columns(4).Width = "100"
            DG.Columns(5).HeaderText = "Buy Account  " : DG.Columns(5).Width = "100"
            DG.Columns(6).HeaderText = "Dep Calculation Account" : DG.Columns(6).Width = "135"
            DG.Columns(7).HeaderText = "Dep Expense Account" : DG.Columns(7).Width = "165"
            DG.Columns(8).HeaderText = "Post Account " : DG.Columns(8).Width = "100"
            DG.Columns(9).HeaderText = "Head Code " : DG.Columns(9).Width = "100"
            Dim counter As Integer = 0
            For Each dr As DataGridViewRow In DG.Rows
                counter += 1
            Next
            Label8.Text = "List All " & counter.ToString - 1 & " Items"
        Else
            DG.Columns(0).HeaderText = "ລ/ດ" : DG.Columns(0).Width = "50"
            DG.Columns(1).HeaderText = "ລະຫັດ" : DG.Columns(1).Width = "90"
            DG.Columns(2).HeaderText = "ຊື່ໝວດຊັບສິນ  (ພາສາລາວ)     " : DG.Columns(2).Width = "300"
            DG.Columns(3).HeaderText = "ຊື່ໝວດຊັບສິນ (ພາສາອັງກິດ)     " : DG.Columns(3).Width = "300"
            DG.Columns(4).HeaderText = "ບັນຊີຊັບສິນ " : DG.Columns(4).Width = "100"
            DG.Columns(5).HeaderText = "ບັນຊີຊື້  " : DG.Columns(5).Width = "100"
            DG.Columns(6).HeaderText = "ບັນຊີຄິດໄລ່ຫຼຸຍຫ້ຽນ" : DG.Columns(6).Width = "135"
            DG.Columns(7).HeaderText = "ບັນຊີລາຍຈ່າຍຄ່າຫຼຸຍຫ້ຽນ" : DG.Columns(7).Width = "165"
            DG.Columns(8).HeaderText = "ບັນຊີສະສາງ " : DG.Columns(8).Width = "100"
            DG.Columns(9).HeaderText = "ລະຫັດຫຼັກ " : DG.Columns(9).Width = "100"
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

    Private Sub DG_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DG.CellContentClick

    End Sub

    Private Sub DG_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DG.DoubleClick

        GrpIDOrder = DG.Item(1, DG.CurrentRow.Index).Value.ToString()
        GrpNm = DG.Item(2, DG.CurrentRow.Index).Value.ToString()
        Me.Close()
    End Sub

    Private Sub DG_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DG.MouseClick
        myTemp = DG.Item(1, DG.CurrentRow.Index).Value.ToString()
    End Sub

    Private Sub cmdNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        mEdit = False
        FrmGrpNew_LS.ShowDialog()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        If myTemp = "" Then Exit Sub
        mEdit = True
        myTemp = DG.Item(1, DG.CurrentRow.Index).Value.ToString()
        FrmGrpNew_LS.ShowDialog()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel.Click
        If myTemp = "" Then Exit Sub
        Dim SSL As New ADODB.Recordset
        Call LoadSqlData("SELECT Group_ID from Assets where Group_ID='" & myTemp & "' ", SSL)
        If SSL.RecordCount > 0 Then
            MsgBox("ລາຍການຊັບສິນໄດ້ຖຶກນຳໃຊ້ແລ້ວ ທ່ານບໍ່ສາມາດລຶບໄດ້", MsgBoxStyle.Exclamation) : Exit Sub
        End If
        If MsgBox("ທ່ານຕ້ອງການລຶບລາຍການຊັບສິນເລກລະຫັດ " & myTemp & " ນີ້ບໍ່?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2) = Windows.Forms.DialogResult.Yes Then
            CNN.Execute("DELETE FROM Groups_Asset WHERE Group_ID='" & myTemp & "' ")
        End If
        Call LoadDG()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim Rs, Rs1 As New ADODB.Recordset
        'Dim rpt As New Cry_Asset_Group
        Dim rpt As Object
        If Lang = False Then
            rpt = New Cry_Asset_Group
        Else
            rpt = New Cry_Asset_GroupE
        End If

        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim str As String = ""
        Call LoadSqlData("Select * from Groups_Asset Where 1=1" & mSql & "  ", Rs)

        With rpt
            'Dim RO As CrystalDecisions.CrystalReports.Engine.ReportObject
            'Dim SRO As CrystalDecisions.CrystalReports.Engine.SubreportObject
            'Dim SubDoc As CrystalDecisions.CrystalReports.Engine.ReportDocument
            'SqlPrint = "SELECT  * from Ap_Image  where Img_Id='" & IMageID & "'"
            'Call LoadData(SqlPrint, Rs1)
            'RO = rpt.ReportDefinition.Sections.Item("Section1").ReportObjects.Item("Subreport3")
            'SRO = CType(RO, CrystalDecisions.CrystalReports.Engine.SubreportObject)
            'SubDoc = SRO.OpenSubreport(SRO.SubreportName)
            'If Rs1.RecordCount > 0 Then
            '    SubDoc.SetDataSource(Rs1)
            '    FmPreview.CrystalReportViewer1.ReportSource = SubDoc
            'End If
            rpt.SetDataSource(Rs)
            FrmPreview.ReportViewer.ReportSource = rpt
            FrmPreview.ReportViewer.DisplayGroupTree = False
            FrmPreview.WindowState = FormWindowState.Maximized
            FrmPreview.Show()
            FrmPreview.Focus()
        End With
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        LoadDG()
    End Sub
End Class