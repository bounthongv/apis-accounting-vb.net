Imports System.Data
Imports System.Windows.Forms

Public Class FmBLS_Item
    
    ' Helper functions for FG2 grid
    Private Function GetGridValue2(ByVal grid As DataGridView, ByVal row As Integer, ByVal col As Integer) As String
        Try
            If grid.RowCount <= row OrElse row < 0 Then Return ""
            If grid.ColumnCount <= col OrElse col < 0 Then Return ""
            If grid.Rows(row).Cells(col).Value Is Nothing Then Return ""
            Return grid.Rows(row).Cells(col).Value.ToString()
        Catch ex As Exception
            Return ""
        End Try
    End Function
    
    Private Sub SetGridValue2(ByVal grid As DataGridView, ByVal row As Integer, ByVal col As Integer, ByVal value As Object)
        Try
            If row < 0 OrElse col < 0 Then Exit Sub
            While grid.RowCount <= row
                grid.Rows.Add()
            End While
            While grid.ColumnCount <= col
                grid.Columns.Add("Col" & grid.ColumnCount)
            End While
            grid.Rows(row).Cells(col).Value = value
        Catch ex As Exception
            ' Ignore
        End Try
    End Sub
Public Sub LoadListFG()
        ' Clear existing rows
        FG.Rows.Clear()
        FG.Columns.Clear()
        
        ' Setup columns
        FG.Columns.Add("Col0", "ລ/ດ")
        FG.Columns.Add("Col1", "ເນື້ອໃນ (ພາສາລາວ)")
        FG.Columns.Add("Col2", "ເນື້ອໃນ (ພາສາອັງກິດ)")
        FG.Columns.Add("Col3", "Chart of Accounts Codes")
        FG.Columns.Add("Col4", "Group")
        FG.Columns.Add("Col5", "Group Name")
        
        ' Set column widths
        For Each col As DataGridViewColumn In FG.Columns
            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next
        
        ' Load data using DataTable
        Dim dtBLS As DataTable = DbHelper.GetDataTable("SELECT * FROM Ap_Rpt_BLS ORDER BY CNT ASC")
        
        If dtBLS.Rows.Count > 0 Then
            For Each row As DataRow In dtBLS.Rows
                Dim newRow As DataGridViewRow = FG.Rows(FG.Rows.Add())
                newRow.Cells("Col0").Value = row("Rpt_ID").ToString()
                newRow.Cells("Col1").Value = row("Description").ToString()
                newRow.Cells("Col2").Value = row("Descriptione").ToString()
                newRow.Cells("Col3").Value = row("Chart_of_Accounts_Codes").ToString()
                newRow.Cells("Col4").Value = row("Grp").ToString()
                newRow.Cells("Col5").Value = row("Grp_Nme").ToString()
            Next
        End If
    End Sub





Public Sub MouseDownEvent()
        If FG2.CurrentRow IsNot Nothing Then
            AC_Code.Text = GetGridValue2(FG2, FG2.CurrentRow.Index, 2)
            Rpt_Type.Text = GetGridValue2(FG2, FG2.CurrentRow.Index, 5)
            TXTCNT.Text = GetGridValue2(FG2, FG2.CurrentRow.Index, 6)
        Else
            Exit Sub
        End If
        
        If FG2.CurrentCell IsNot Nothing Then
            If FG2.CurrentCell.ColumnIndex = 5 Then
                FG2.ReadOnly = False
                'MsgBox(FG2.CurrentCell.ColumnIndex)
            Else
                FG2.ReadOnly = True
            End If
            BtnSearch.Visible = True
            Select Case MouseButtons
                Case Windows.Forms.MouseButtons.Right
                    FG2.BeginEdit(True)
                Case Windows.Forms.MouseButtons.Left
                    'MsgBox(FG2.CurrentCell.ColumnIndex)
                    If FG2.CurrentCell.ColumnIndex = 2 Then
                        BtnSearch.Visible = True
                    Else
                        BtnSearch.Visible = False
                    End If
                    If FG2.CurrentRow IsNot Nothing AndAlso FG2.CurrentRow.Index = FG2.Rows.Count - 1 Then
                        BtnMove.Visible = False
                    End If
            End Select
        End If
    End Sub





Private Sub loadBankItem()
        ' Clear existing rows and setup columns
        FG2.Rows.Clear()
        FG2.Columns.Clear()
        
        ' Setup columns
        FG2.Columns.Add("Col0", "ລ/ດ")
        FG2.Columns.Add("Col1", "ລະຫັດ")
        FG2.Columns.Add("Col2", "ລະຫັດບັນຊີ")
        FG2.Columns.Add("Col3", "ຊື່ບັນຊີ(ພາສາລາວ)")
        FG2.Columns.Add("Col4", "ຊື່ບັນຊີ(ພາສາອັງກິດ)")
        FG2.Columns.Add("Col5", "ສະຖານະພາບ")
        FG2.Columns.Add("Col6", "CNT")
        
        ' Set column widths
        For Each col As DataGridViewColumn In FG2.Columns
            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next
        
        ' Load data using DataTable
        Dim dtBLSItem As DataTable = DbHelper.GetDataTable("SELECT * FROM Ap_Rpt_BLS_Item where Rpt_ID=N'" & TextBox1.Text & "' Order by Ac_Code ")
        
        If dtBLSItem.Rows.Count > 0 Then
            For i As Integer = 0 To dtBLSItem.Rows.Count - 1
                Dim row As DataRow = dtBLSItem.Rows(i)
                Dim newRow As DataGridViewRow = FG2.Rows(FG2.Rows.Add())
                newRow.Cells("Col0").Value = (i + 1).ToString()
                newRow.Cells("Col1").Value = row("Rpt_ID").ToString()
                newRow.Cells("Col2").Value = row("Ac_Code").ToString()
                newRow.Cells("Col3").Value = row("Ac_Name").ToString()
                newRow.Cells("Col4").Value = row("Ac_NameE").ToString()
                newRow.Cells("Col5").Value = row("Rpt_Type").ToString()
                newRow.Cells("Col6").Value = row("CNT").ToString()
            Next
        End If
    End Sub

Private Sub FG2_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles FG2.CellEndEdit
        Button2.Enabled = True
    End Sub

    Private Sub FG2_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles FG2.Scroll
        BtnSearch.Visible = False
        BtnMove.Visible = False
    End Sub

    Private Sub FG2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles FG2.MouseDown
        MouseDownEvent()
    End Sub

    Private Sub BtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch.Click
        fmShartOfAccDetail.txtSty.Text = "FmBLS"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
        FG.Focus()
    End Sub

    Private Sub FmBankReportId_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button2.Enabled = False
        BtnMove.Visible = False
        BtnSearch.Visible = False
FG.Size = New System.Drawing.Size(519, 378)
        LoadListFG()
        'FG.AllowUserToResizeColumns = True
        'FG.AllowUserToResizeRows = True
    End Sub

    Private Sub BtnMove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnMove.Click
        'FG2.RemoveItem()

        CNN.Execute("delete Ap_Rpt_BLS_Item where Ac_Code like '" & AC_Code.Text & "%' And Rpt_ID =N'" & RPT_ID.Text & "' And Rpt_Type =N'" & Rpt_Type.Text & "'  And CNT =N'" & TXTCNT.Text & "' ")
        BtnMove.Visible = False
        Call loadBankItem()
    End Sub

Private Sub FG_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles FG.MouseDown
        Select Case MouseButtons
            Case Windows.Forms.MouseButtons.Right
            Case Windows.Forms.MouseButtons.Left
                If FG.CurrentRow IsNot Nothing AndAlso FG.CurrentRow.Index = FG.Rows.Count - 1 Then
                    Button1.Visible = False
                Else
                    Button1.Visible = True
                End If
                If FG.CurrentRow IsNot Nothing Then
                    Button1.Top = FG.CurrentRow.GetCellDisplayRectangle(0, False).Top + FG.Top
                End If
        End Select
    End Sub

    Private Sub FG_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelectionChanged
        If FG.CurrentRow IsNot Nothing Then
            TextBox1.Text = GetGridValue2(FG, FG.CurrentRow.Index, 1)
            RPT_ID.Text = GetGridValue2(FG, FG.CurrentRow.Index, 1)
            Call loadBankItem()
        End If

        Button2.Enabled = False
        BtnMove.Visible = False
        BtnSearch.Visible = False
    End Sub

Private Sub BtnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnEdit.Click
        Button2.Enabled = False
        BtnMove.Visible = False
        BtnSearch.Visible = False
        MsgBox("ການບັນຶກສຳເລັດຜົນ")
        
        For i As Integer = 0 To FG.Rows.Count - 1
            If GetGridValue2(FG, i, 1) = "" And GetGridValue2(FG, i, 2) = "" Then
                Continue For
            End If

            CNN.Execute("Update Ap_Rpt_BLS Set  Description = N'" & Apostrophe(GetGridValue2(FG, i, 2)) & "' ,  Descriptione = N'" & Apostrophe(GetGridValue2(FG, i, 3)) & "' , Chart_of_Accounts_Codes = N'" & GetGridValue2(FG, i, 4) & "'  Where Rpt_ID = '" & GetGridValue2(FG, i, 1) & "'")
        Next i
    End Sub

Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        MsgBox("ການບັນຶກສຳເລັດຜົນ")
        CNN.Execute("delete Ap_Rpt_BLS_Item where Rpt_ID = '" & TextBox1.Text & "' ")
        
        For i As Integer = 0 To FG2.Rows.Count - 1
            If GetGridValue2(FG2, i, 1) = "" And GetGridValue2(FG2, i, 2) = "" Then
                Continue For
            End If
            CNN.Execute("INSERT INTO Ap_Rpt_BLS_Item( Rpt_ID,  Ac_Code , Ac_Name , Ac_NameE ,Amt_Dr , Amt_Cr , BLS  , Rpt_Type) " & _
                 "Values('" & GetGridValue2(FG2, i, 1) & "', N'" & GetGridValue2(FG2, i, 2) & "', N'" & GetGridValue2(FG2, i, 3) & "','" & GetGridValue2(FG2, i, 4) & "','" & CDbl(0) & "','" & CDbl(0) & "','" & "ALL" & "' ,'" & GetGridValue2(FG2, i, 5) & "')")
        Next i
    End Sub

Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If FG.CurrentRow IsNot Nothing Then
            FG.Rows.RemoveAt(FG.CurrentRow.Index)
        End If
        Button1.Visible = False
    End Sub
    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click

        Call Close()

    End Sub

Private Sub AC_Code_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles AC_Code.KeyPress
        If e.KeyChar = Chr(13) Then
            Dim dtCheck As DataTable = DbHelper.GetDataTable("Select top 1 Rpt_ID , Ac_Code from Ap_Rpt_BLS_Item where  Ac_Code like '" & AC_Code.Text & "%'  And Rpt_ID <> '" & RPT_ID.Text & "' ")
            If dtCheck.Rows.Count > 0 Then
                MsgBox("ເລກບັນຊີ " & dtCheck.Rows(0)("Ac_Code").ToString() & " ມີຢູ່ " & dtCheck.Rows(0)("Rpt_ID").ToString() & " ແລ້ວ")
                Exit Sub
            End If

            CNN.Execute("delete Ap_Rpt_BLS_Item where Ac_Code like '" & AC_Code.Text & "%' And Rpt_ID = '" & RPT_ID.Text & "' And Rpt_Type = '" & Rpt_Type.Text & "'  insert into Ap_Rpt_BLS_Item (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type) select N'" & RPT_ID.Text & "' ,  Ac_Code , Name_L , '" & Rpt_Type.Text & "' from Acc_Code where Ac_Code like '" & AC_Code.Text & "%'  ")
            If FG.CurrentRow IsNot Nothing Then
                TextBox1.Text = GetGridValue2(FG, FG.CurrentRow.Index, 1)
            End If
            Call loadBankItem()
        End If
    End Sub

    Private Sub AC_Code_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AC_Code.TextChanged

    End Sub

Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        CNN.Execute("delete Ap_Rpt_BLS_Item where  Rpt_ID =N'" & RPT_ID.Text & "'   ")
        If FG.CurrentRow IsNot Nothing Then
            TextBox1.Text = GetGridValue2(FG, FG.CurrentRow.Index, 1)
        End If
        Call loadBankItem()
    End Sub

Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If FG.Rows.Count > 18 Then
            FG.CurrentCell = FG.Rows(18).Cells(2)
            FG.BeginEdit(True)
        End If
    End Sub

Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        CNN.Execute("delete Ap_Rpt_Item")
        
        Dim dtBLSItem As DataTable = DbHelper.GetDataTable("SELECT * FROM Ap_Rpt_BLS_Item Order by Ac_Code ")
        If dtBLSItem.Rows.Count > 0 Then
            For Each row As DataRow In dtBLSItem.Rows
                Dim acCode As String = Trim(row("Ac_Code").ToString())
                Dim rptId As String = Trim(row("Rpt_ID").ToString())
                Dim rptType As String = Trim(row("Rpt_Type").ToString())
                
                CNN.Execute("delete Ap_Rpt_Item where Ac_Code like '" & acCode & "%' And Rpt_ID = '" & rptId & "' And Rpt_Type = '" & rptType & "' " & _
                            " insert into Ap_Rpt_Item (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type) select '" & rptId & "' ,  Ac_Code , Name_L , '" & rptType & "' from Acc_Code where Ac_Code like '" & acCode & "%'  ")
            Next
        End If
        
        CNN.Execute("delete Ap_Rpt_BLS_Item")
        CNN.Execute(" insert into Ap_Rpt_BLS_Item  (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type) select Rpt_ID , Ac_Code , Ac_Name, Rpt_Type from Ap_Rpt_Item")
        MsgBox("Ok")
    End Sub

End Class