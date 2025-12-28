Public Class FrmDepartment_List
    Dim cn As New Odbc.OdbcConnection
    Dim mSql, Section, SecPView As String

    Private Sub FrmAcc_Code_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Call Loadlang()
        'SetControlText(Me)
        'ChgChildForm()
        LoadDG()
        'If mCompStr = "" Then btnEdit.Enabled = False
        LdSec()
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
        'If MDEdit = 1 Then
        '    btnEdit.Enabled = True
        'Else
        '    btnEdit.Enabled = False
        'End If
        '===============Delete==========
        'If MDDelete = 1 Then
        '    btnDel.Enabled = True
        'Else
        '    btnDel.Enabled = False
        'End If
    End Sub
    Private Sub LdSec()
        Dim sRS As New ADODB.Recordset
        cmbSec.Items.Clear()
        cmbSec.Items.Add("**ທັງໝົດ**")
        Call LoadSqlData("Select * from AP_Office where Off_ID<>'00' Order by Off_ID", sRS)
        If sRS.RecordCount <> 0 Then
            While Not sRS.EOF
                cmbSec.Items.Add(sRS.Fields("off_name").Value.ToString)
                sRS.MoveNext()
            End While
        End If
        'If cmbSec.Items.Count > 0 Then cmbSec.SelectedIndex = 1
    End Sub
    Private Sub CndClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CndClose.Click

        Me.Close()
    End Sub

    Private Sub FrmCategory_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        'FG.Width = Me.Width - 50
        'FG.Height = Me.Height - 100
    End Sub

    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Call btnEdit_Click(sender, e)
    End Sub

    Private Sub cmdDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call LoadSqlData("SELECT * FROM Assets where DepartmentID='" & myTemp & "' ", RSC)
        If RSC.RecordCount > 0 Then
            MsgBox("ທ່ານບໍ່ສາມາດລຶບລາຍການພະແນກນີ້ໄດ້ເພາະມີການເຄື່ອນໄຫວແລ້ວ") : Exit Sub
        End If
        If MsgBox("ທ່ານຕ້ອງການລຶບລາຍ " & myTemp & " ນີ້ບໍ່?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2) = Windows.Forms.DialogResult.Yes Then
            CNN.Execute("DELETE FROM Department WHERE DepartmentID='" & myTemp & "' ")
            Call LoadDG()
        End If
    End Sub

    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        mEdit = True
        myTemp = Trim(DG.Item(1, DG.CurrentRow.Index).Value.ToString())
        FrmDepartmentNew.ShowDialog()
        'mSql = " Order by DepartmentID"
        'Call LdData(mSql)
        LoadDG()
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        mEdit = False

        FrmDepartmentNew.ShowDialog()

        Call LoadDG()
    End Sub

    Private Sub cmbSort_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSec.SelectedIndexChanged
        Dim SSL As New ADODB.Recordset
        Call LoadSqlData("Select * from AP_Office Where Off_Name=N'" & Trim(cmbSec.Text) & "' ", SSL)
        If SSL.RecordCount = 0 Then
            Section = ""
            SecPView = ""
            txtSec.Text = ""
        Else
            txtSec.Text = Trim(SSL.Fields("Off_ID").Value.ToString)
            Section = ""
            SecPView = ""
            'Section = " And  DepartmentID = '" & txtSec.Text & "'  "
            'SecPView = " And  Department.DepartmentID  = '" & txtSec.Text & "'  "
            Section = " And Left(DepartmentID,2) = '" & txtSec.Text & "'  "
            SecPView = " And Left(Department.DepartmentID,2) = '" & txtSec.Text & "'  "
        End If
        Call LoadDG()
    End Sub
    Private Sub LoadDG()
        DG.DataSource = 0
        Dim ds As New DataSet
        Try
            ConnectCL()
            sql = "Select 0,DepartmentID,DepartmentNm,DepartmentNmE,DepartmentRemark,Company  from Department where 1=1 " & Section & " Order by DepartmentID ASC "
            'SQL = "Select SecID,SecNmL,SecNmE,Remark from Sections  " & mSql & " "
            LoadCN_DG()
            da.Fill(ds, " Department")
            DG.DataSource = ds.Tables(" Department")
            cn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try
        'FG.FormatString = "^ລ/ດ |<ລະຫັດພະແນກ |<ຊື່ພະແນກ (ພາສາລາວ)                         |<ຊື່ພະແນກ (ພາສາອັງກິດ)                  |<ໝາຍເຫດ                  | ລະຫັດສາຂາ      "
        If Lang = True Then
            DG.Columns(0).HeaderText = "No" : DG.Columns(0).Width = "50"
            DG.Columns(1).HeaderText = "Code" : DG.Columns(1).Width = "100"
            DG.Columns(2).HeaderText = "Department Name (Lao)  " : DG.Columns(2).Width = "300"
            DG.Columns(3).HeaderText = "Department Name (Eng) " : DG.Columns(3).Width = "300"
            DG.Columns(4).HeaderText = "Remark" : DG.Columns(4).Width = "300"
            DG.Columns(5).HeaderText = "Branch Code" : DG.Columns(5).Width = "150"
            Dim counter As Integer = 0
            For Each dr As DataGridViewRow In DG.Rows
                counter += 1
            Next
            Label8.Text = "List All " & counter.ToString - 1 & " Items "
        Else
            DG.Columns(0).HeaderText = "ລ/ດ" : DG.Columns(0).Width = "50"
            DG.Columns(1).HeaderText = "ລະຫັດ" : DG.Columns(1).Width = "100"
            DG.Columns(2).HeaderText = "ຊື່ພະແນກ (ພາສາລາວ)  " : DG.Columns(2).Width = "300"
            DG.Columns(3).HeaderText = "ຊື່ພະແນກ (ພາສາອັງກິດ) " : DG.Columns(3).Width = "300"
            DG.Columns(4).HeaderText = "ໝາຍເຫດ" : DG.Columns(4).Width = "300"
            DG.Columns(5).HeaderText = "ລະຫັດສາຂາ" : DG.Columns(5).Width = "100"
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        LoadDG()
    End Sub

    Private Sub DG_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DG.DoubleClick
        'Call btnEdit_Click(sender, e)
    End Sub

    Private Sub DG_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DG.MouseClick
        myTemp = Trim(DG.Item(1, DG.CurrentRow.Index).Value.ToString())
    End Sub

    Private Sub DG_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DG.CellContentClick

    End Sub

    Private Sub cmdNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        mEdit = False

        FrmDepartmentNew.ShowDialog()

        Call LoadDG()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        mEdit = True
        myTemp = Trim(DG.Item(1, DG.CurrentRow.Index).Value.ToString())
        FrmDepartmentNew.ShowDialog()
        'mSql = " Order by DepartmentID"
        'Call LdData(mSql)
        LoadDG()
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel.Click
        Call LoadSqlData("SELECT * FROM Assets where DepartmentID='" & myTemp & "' ", RSC)
        If RSC.RecordCount > 0 Then
            MsgBox("ທ່ານບໍ່ສາມາດລຶບລາຍການພະແນກນີ້ໄດ້ເພາະມີການເຄື່ອນໄຫວແລ້ວ", MsgBoxStyle.Exclamation) : Exit Sub
        End If
        If MsgBox("ທ່ານຕ້ອງການລຶບລາຍ " & myTemp & " ນີ້ບໍ່?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2) = Windows.Forms.DialogResult.Yes Then
            CNN.Execute("DELETE FROM Department WHERE DepartmentID='" & myTemp & "' ")
            Call LoadDG()
        End If
    End Sub

    Private Sub Button2_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Office()
        Dim Rs, Rs1 As New ADODB.Recordset
        Dim rpt As Object
        If Lang = False Then
            rpt = New Cry_Asset_Department
        Else
            rpt = New Cry_Asset_DepartmentE
        End If
        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim str As String = ""
        'Call LoadData("Select * from Department Where 1=1" & str & " Order by DepartmentID", Rs)
        Call LoadSqlData("SELECT dbo.AP_Office.Off_ID, dbo.AP_Office.Off_Nm, dbo.AP_Office.Off_NmE, dbo.Department.DepartmentID, dbo.Department.DepartmentNm, dbo.Department.DepartmentNmE " & _
        "FROM dbo.AP_Office INNER JOIN dbo.Department ON  (dbo.AP_Office.Off_ID) =  left(dbo.Department.DepartmentID,2) where 1=1 " & SecPView & " ", Rs)
        If Rs.RecordCount = 0 Then
            MsgBox("Date Entry", vbInformation, "Information")
        Else
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
                '    FmPreview.ReportViewer.ReportSource = SubDoc
                'End If
                rpt.SetDataSource(Rs)
                FrmPreview.ReportViewer.ReportSource = rpt
                FrmPreview.ReportViewer.DisplayGroupTree = False
                FrmPreview.WindowState = FormWindowState.Maximized
                FrmPreview.Show()
                FrmPreview.Focus()
            End With
        End If
    End Sub

    Private Sub Button2_Click_3(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        LoadDG()
    End Sub
End Class