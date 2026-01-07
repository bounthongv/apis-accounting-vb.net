Imports System.Data
Imports System.Drawing

Public Class FmBLS_Item_Old
    
    ' Helper function to get column index by name
    Private Function GetColumnIndex(dgv As DataGridView, columnName As String) As Integer
        For i As Integer = 0 To dgv.Columns.Count - 1
            If dgv.Columns(i).Name = columnName Then
                Return i
            End If
        Next
        Return -1
    End Function
    Public Sub LoadListFG()
        FG.Rows.Clear()
        FG.Columns.Clear()
        
        ' Setup columns
        FG.Columns.Add("Col1", "ລ/ດ")
        FG.Columns.Add("Rpt_ID", "ລະຫັດ")
        FG.Columns.Add("Description", "ເນື້ອໃນ (ພາສາລາວ)")
        FG.Columns.Add("Descriptione", "ເນື້ອໃນ (ພາສາອັງກິດ)")
        FG.Columns.Add("Chart_of_Accounts_Codes", "")
        FG.Columns.Add("Grp", "")
        FG.Columns.Add("Grp_Nme", "")
        
        ' Set column widths
        FG.Columns("Col1").Width = 40
        FG.Columns("Rpt_ID").Width = 80
        FG.Columns("Description").Width = 200
        FG.Columns("Descriptione").Width = 200
        FG.Columns("Chart_of_Accounts_Codes").Width = 100
        FG.Columns("Grp").Width = 50
        FG.Columns("Grp_Nme").Width = 100
        
        ' Load data using DbHelper
        Try
            Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM Ap_Rpt_BLS_Old WHERE 1=1 " & RPT_GRP & " ORDER BY GRP, Rpt_ID ASC")
            
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim row As String() = {
                    (i + 1).ToString(),
                    DbHelper.GetStr(dt.Rows(i)("Rpt_ID")),
                    DbHelper.GetStr(dt.Rows(i)("Description")),
                    DbHelper.GetStr(dt.Rows(i)("Descriptione")),
                    DbHelper.GetStr(dt.Rows(i)("Chart_of_Accounts_Codes")),
                    DbHelper.GetStr(dt.Rows(i)("Grp")),
                    DbHelper.GetStr(dt.Rows(i)("Grp_Nme"))
                }
                FG.Rows.Add(row)
            Next
        Catch ex As Exception
            MsgBox("Error loading data: " & ex.Message)
        End Try
    End Sub





    Public Sub MouseDownEvent()
        If FG2.CurrentRow >= 0 AndAlso FG2.CurrentRow < FG2.Rows.Count Then
            AC_Code.Text = If(FG2.CurrentRow >= 0, CStr(FG2.Rows(FG2.CurrentRow).Cells("Ac_Code").Value), "")
            Rpt_Type.Text = If(FG2.CurrentRow >= 0, CStr(FG2.Rows(FG2.CurrentRow).Cells("Rpt_Type").Value), "")
            TXTCNT.Text = If(FG2.CurrentRow >= 0, CStr(FG2.Rows(FG2.CurrentRow).Cells("CNT").Value), "")
        End If
        
        If FG2.CurrentCell.ColumnIndex = GetColumnIndex(FG2, "Rpt_Type") Then
            FG2.ReadOnly = False
        Else
            FG2.ReadOnly = True
        End If
        
        BtnSearch.Visible = True
        Select Case MouseButtons
            Case Windows.Forms.MouseButtons.Right
                If FG2.CurrentCell IsNot Nothing Then
                    FG2.BeginEdit(True)
                End If
            Case Windows.Forms.MouseButtons.Left
                If FG2.CurrentCell IsNot Nothing Then
                    If FG2.CurrentCell.ColumnIndex = GetColumnIndex(FG2, "Ac_Code") Then
                        BtnSearch.Visible = True
                    Else
                        BtnSearch.Visible = False
                    End If
                    
                    If FG2.CurrentRow = FG2.Rows.Count - 1 Then
                        BtnMove.Visible = False
                    Else
                        BtnMove.Visible = True
                    End If
                    
                    ' Position buttons near current cell
                    Dim cellRect As Rectangle = FG2.GetCellDisplayRectangle(FG2.CurrentCell.ColumnIndex, FG2.CurrentRow, False)
                    BtnSearch.Left = FG2.Left + cellRect.Left + cellRect.Width - BtnSearch.Width
                    BtnSearch.Top = FG2.Top + cellRect.Top
                    BtnMove.Top = FG2.Top + cellRect.Top
                End If
        End Select
    End Sub





    Private Sub loadBankItem()
        FG2.Rows.Clear()
        FG2.Columns.Clear()
        
        ' Setup columns
        FG2.Columns.Add("Col1", "ລ/ດ")
        FG2.Columns.Add("Rpt_ID", "ລະຫັດ")
        FG2.Columns.Add("Ac_Code", "ລະຫັດບັນຊີ")
        FG2.Columns.Add("Ac_Name", "ຊື່ບັນຊີ(ພາສາລາວ)")
        FG2.Columns.Add("Ac_NameE", "ຊື່ບັນຊີ(ພາສາອັງກິດ)")
        FG2.Columns.Add("Rpt_Type", "ສະຖານະພາບ")
        FG2.Columns.Add("CNT", "CNT")
        
        ' Set column widths
        FG2.Columns("Col1").Width = 40
        FG2.Columns("Rpt_ID").Width = 80
        FG2.Columns("Ac_Code").Width = 100
        FG2.Columns("Ac_Name").Width = 250
        FG2.Columns("Ac_NameE").Width = 200
        FG2.Columns("Rpt_Type").Width = 80
        FG2.Columns("CNT").Width = 50
        
        ' Load data using DbHelper
        Try
            Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM Ap_Rpt_BLS_Item_Old WHERE Rpt_ID=N'" & TextBox1.Text & "' " & RPT_GRP & " ORDER BY Ac_Code")
            
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim row As String() = {
                    (i + 1).ToString(),
                    DbHelper.GetStr(dt.Rows(i)("Rpt_ID")),
                    DbHelper.GetStr(dt.Rows(i)("Ac_Code")),
                    DbHelper.GetStr(dt.Rows(i)("Ac_Name")),
                    DbHelper.GetStr(dt.Rows(i)("Ac_NameE")),
                    DbHelper.GetStr(dt.Rows(i)("Rpt_Type")),
                    DbHelper.GetStr(dt.Rows(i)("CNT"))
                }
                FG2.Rows.Add(row)
            Next
        Catch ex As Exception
            MsgBox("Error loading bank items: " & ex.Message)
        End Try
    End Sub

    Private Sub FG2_CellEndEdit(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FG2.CellEndEdit
        Button2.Enabled = True
    End Sub

    Private Sub FG2_Scroll(ByVal sender As Object, ByVal e As ScrollEventArgs) Handles FG2.Scroll
        BtnSearch.Visible = False
        BtnMove.Visible = False
    End Sub

    Private Sub FG2_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles FG2.MouseDown
        MouseDownEvent()
    End Sub

    Private Sub FG2_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles FG2.SelectionChanged

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
    End Sub

    Private Sub BtnMove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnMove.Click
        'FG2.RemoveItem()

        CNN.Execute("delete Ap_Rpt_BLS_Item_Old where Ac_Code like '" & AC_Code.Text & "%' And Rpt_ID =N'" & RPT_ID.Text & "' And Rpt_Type =N'" & Rpt_Type.Text & "'  And CNT =N'" & TXTCNT.Text & "' " & RPT_GRP & " ")
        BtnMove.Visible = False
        Call loadBankItem()
    End Sub

    Private Sub FG_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles FG.MouseDown
        Select Case e.Button
            Case MouseButtons.Right
            Case MouseButtons.Left
                If FG.CurrentRow = FG.Rows.Count - 1 Then
                    Button1.Visible = False
                Else
                    Button1.Visible = True
                End If
                If FG.CurrentRow >= 0 Then
                    Dim cellRect As Rectangle = FG.GetCellDisplayRectangle(0, FG.CurrentRow, False)
                    Button1.Top = FG.Top + cellRect.Top
                End If
        End Select
    End Sub

    Private Sub FG_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles FG.SelectionChanged
        If FG.CurrentRow >= 0 Then
            TextBox1.Text = CStr(FG.Rows(FG.CurrentRow).Cells("Rpt_ID").Value)
            RPT_ID.Text = CStr(FG.Rows(FG.CurrentRow).Cells("Rpt_ID").Value)
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
            If CStr(FG.Rows(i).Cells("Rpt_ID").Value) = "" And CStr(FG.Rows(i).Cells("Description").Value) = "" Then
                Exit For
            End If

            Dim sql As String = "UPDATE Ap_Rpt_BLS_Old SET Description = N'" & Apostrophe(CStr(FG.Rows(i).Cells("Description").Value)) & _
                               "', Descriptione = N'" & Apostrophe(CStr(FG.Rows(i).Cells("Descriptione").Value)) & _
                               "', Chart_of_Accounts_Codes = N'" & CStr(FG.Rows(i).Cells("Chart_of_Accounts_Codes").Value) & _
                               "' WHERE Rpt_ID = '" & CStr(FG.Rows(i).Cells("Rpt_ID").Value) & "'"
            
            DbHelper.ExecuteNonQuery(sql)
        Next i
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        MsgBox("ການບັນຶກສຳເລັດຜົນ")
        DbHelper.ExecuteNonQuery("DELETE FROM Ap_Rpt_BLS_Item_Old WHERE Rpt_ID = '" & TextBox1.Text & "'")
        
        For i As Integer = 0 To FG2.Rows.Count - 1
            If CStr(FG2.Rows(i).Cells("Rpt_ID").Value) = "" And CStr(FG2.Rows(i).Cells("Ac_Code").Value) = "" Then
                Exit For
            End If
            
            Dim sql As String = "INSERT INTO Ap_Rpt_BLS_Item_Old(Rpt_ID, Ac_Code, Ac_Name, Ac_NameE, Amt_Dr, Amt_Cr, BLS, Rpt_Type) " & _
                               "VALUES('" & CStr(FG2.Rows(i).Cells("Rpt_ID").Value) & "', N'" & CStr(FG2.Rows(i).Cells("Ac_Code").Value) & _
                               "', N'" & CStr(FG2.Rows(i).Cells("Ac_Name").Value) & "', '" & CStr(FG2.Rows(i).Cells("Ac_NameE").Value) & _
                               "', 0, 0, 'ALL', '" & CStr(FG2.Rows(i).Cells("Rpt_Type").Value) & "')"
            
            DbHelper.ExecuteNonQuery(sql)
        Next i
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If FG.CurrentRow >= 0 AndAlso FG.CurrentRow < FG.Rows.Count - 1 Then
            FG.Rows.RemoveAt(FG.CurrentRow)
            Button1.Visible = False
        End If
    End Sub
    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click

        Call Close()

    End Sub

    Private Sub AC_Code_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles AC_Code.KeyPress
        If e.KeyChar = Chr(13) Then
            Try
                Dim dt As DataTable = DbHelper.GetDataTable("SELECT TOP 1 Rpt_ID, Ac_Code FROM Ap_Rpt_BLS_Item_Old WHERE Ac_Code LIKE '" & AC_Code.Text & "%' AND Rpt_ID <> '" & RPT_ID.Text & "' " & RPT_GRP)
                
                If dt.Rows.Count > 0 Then
                    MsgBox("ເລກບັນຊີ " & DbHelper.GetStr(dt.Rows(0)("Ac_Code")) & " ມີຢູ່ " & DbHelper.GetStr(dt.Rows(0)("Rpt_ID")) & " ແລ້ວ")
                    Exit Sub
                End If

                DbHelper.ExecuteNonQuery("DELETE FROM Ap_Rpt_BLS_Item_Old WHERE Ac_Code LIKE '" & AC_Code.Text & "%' AND Rpt_ID = '" & RPT_ID.Text & "' AND Rpt_Type = '" & Rpt_Type.Text & "' " & RPT_GRP)
                
                Dim SS As String = "INSERT INTO Ap_Rpt_BLS_Item_Old (Rpt_ID, Ac_Code, Ac_Name, Rpt_Type, GRP) SELECT N'" & RPT_ID.Text & "', Ac_Code, Name_L, '" & Rpt_Type.Text & "', " & RPT_GRPID & " FROM Acc_Code WHERE Ac_Code LIKE '" & AC_Code.Text & "%'"
                DbHelper.ExecuteNonQuery(SS)
                
                If FG.CurrentRow >= 0 Then
                    TextBox1.Text = CStr(FG.Rows(FG.CurrentRow).Cells("Rpt_ID").Value)
                End If
                Call loadBankItem()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub AC_Code_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AC_Code.TextChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        DbHelper.ExecuteNonQuery("DELETE FROM Ap_Rpt_BLS_Item_Old WHERE Rpt_ID =N'" & RPT_ID.Text & "' " & RPT_GRP)
        
        If FG.CurrentRow >= 0 Then
            TextBox1.Text = CStr(FG.Rows(FG.CurrentRow).Cells("Rpt_ID").Value)
        End If
        Call loadBankItem()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If FG.Rows.Count > 18 Then
            FG.CurrentCell = FG.Rows(18).Cells("Description")
            FG.Focus()
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Try
            DbHelper.ExecuteNonQuery("DELETE FROM Ap_Rpt_Item")
            
            Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM Ap_Rpt_BLS_Item_Old ORDER BY Ac_Code")
            
            For Each row As DataRow In dt.Rows
                Dim acCode As String = DbHelper.GetStr(row("Ac_Code"))
                Dim rptId As String = DbHelper.GetStr(row("Rpt_ID"))
                Dim rptType As String = DbHelper.GetStr(row("Rpt_Type"))
                
                DbHelper.ExecuteNonQuery("DELETE FROM Ap_Rpt_Item WHERE Ac_Code LIKE '" & acCode & "%' AND Rpt_ID = '" & rptId & "' AND Rpt_Type = '" & rptType & "'")
                DbHelper.ExecuteNonQuery("INSERT INTO Ap_Rpt_BLS_Item_Old (Rpt_ID, Ac_Code, Ac_Name, Rpt_Type) SELECT '" & rptId & "', Ac_Code, Name_L, '" & rptType & "' FROM Acc_Code WHERE Ac_Code LIKE '" & acCode & "%'")
            Next
            
            DbHelper.ExecuteNonQuery("DELETE FROM Ap_Rpt_BLS_Item_Old")
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_Rpt_BLS_Item_Old (Rpt_ID, Ac_Code, Ac_Name, Rpt_Type) SELECT Rpt_ID, Ac_Code, Ac_Name, Rpt_Type FROM Ap_Rpt_Item")
            MsgBox("Ok")
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub RPT_ID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RPT_ID.TextChanged

    End Sub

    Private Sub Rpt_Type_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Rpt_Type.SelectedIndexChanged

    End Sub

    Private Sub TXTCNT_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTCNT.TextChanged

    End Sub
End Class