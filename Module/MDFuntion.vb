
Option Explicit On
Option Strict On

Imports System.Data.SqlClient
'MDSeriaCom = CDbl(Mid(MDPartitionSeria.ShowDriveInfo(Application.StartupPath), 2, Len(Trim(MDPartitionSeria.ShowDriveInfo(Application.StartupPath)) - 1)))
Module MDFuntion
    Public Function Apostrophe(ByVal sFieldString As String) As String
        If CBool(InStr(sFieldString, "'")) Then
            Dim iLen As Integer
            Dim i As Integer
            Dim apostr As String
            iLen = Len(sFieldString)
            i = 1
            Do While i <= iLen
                If Mid(sFieldString, i, 1) = "'" Then
                    apostr = CStr(i)
                    sFieldString = Left(sFieldString, CInt(apostr)) & "'" & Right(sFieldString, iLen - CInt(apostr))
                    iLen = Len(sFieldString)
                    i = i + 1
                End If
                i = i + 1
            Loop
        End If
        Apostrophe = sFieldString
    End Function
    
    Public Sub Check_Pair(ByVal mFG As DataGridView, ByVal mCol As Integer)
        Dim Pair, Pair1, i As Integer
        Dim Plus As Boolean
        Plus = False
        Pair = 1
        For i = 0 To mFG.Rows.Count - 2 ' DataGridView is 0-based
            If GetGridValue(mFG, i, 0) <> "" Then
                Pair1 = Pair
                Plus = False
            Else
                If Plus = False Then
                    Pair = Pair + 1
                    Plus = True
                End If
            End If
            SetGridValue(mFG, i, mCol, CStr(Pair1))
        Next i
    End Sub
    Public Sub GridEdit(ByVal g As DataGridView, ByVal C As TextBox, ByVal KeyAscii As Integer)
        '--------------------------------------------------
        ' prepare the control
        ' On Error Resume Next

        '--------------------------------------------------
        ' show it at the right place
        If g.EditType = DataGridViewEditType.OnKeystroke Then
            If g.CurrentCell Is Nothing OrElse g.CurrentCell.RowIndex <> g.CurrentRow.Index OrElse g.CurrentCell.ColumnIndex <> g.CurrentCell.ColumnIndex Then
                g.CurrentCell = g.CurrentCell
            End If
        End If

        ' and let it rip
        On Error Resume Next
        On Error GoTo 0

    End Sub
    Public Sub EditKeyCode(ByVal g As DataGridView, ByVal C As TextBox, ByVal KeyCode%, ByVal Shift%)
        ' standard edit control processing
        Select Case KeyCode
            Case 27 ' esc
                C.Visible = False
                g.Focus()
            Case 13 ' enter
                g.Focus()
                Application.DoEvents()
            Case 40 ' down
                g.Focus()
                Application.DoEvents()
                If g.CurrentRow IsNot Nothing AndAlso g.CurrentRow.Index < g.Rows.Count - 1 Then
                    g.CurrentCell = g.Rows(g.CurrentRow.Index + 1).Cells(g.CurrentCell.ColumnIndex)
                End If
            Case 38 ' up
                g.Focus()
                Application.DoEvents()
                If g.CurrentRow IsNot Nothing AndAlso g.CurrentRow.Index > 0 Then
                    g.CurrentCell = g.Rows(g.CurrentRow.Index - 1).Cells(g.CurrentCell.ColumnIndex)
                End If
        End Select
    End Sub
    Public Sub load_Cmb(ByVal sql As String, ByVal Str As String, ByVal CmbFg As ComboBox)
        Dim dt As DataTable = GetDataTable(sql)
        
        CmbFg.BeginUpdate()
        For Each row As DataRow In dt.Rows
            CmbFg.Items.Add(IIf(IsDBNull(row(Str)) = True, "", Trim(CStr(row(Str)))))
        Next
        CmbFg.EndUpdate()
    End Sub
    ' Helper functions for DataGridView
    Private Function GetGridValue(ByVal grid As DataGridView, ByVal row As Integer, ByVal col As Integer) As String
        Try
            If grid.Rows.Count <= row OrElse row < 0 Then Return ""
            If grid.ColumnCount <= col OrElse col < 0 Then Return ""
            If grid.Rows(row).Cells(col).Value Is Nothing Then Return ""
            Return grid.Rows(row).Cells(col).Value.ToString()
        Catch ex As Exception
            Return ""
        End Try
    End Function
    
    Private Sub SetGridValue(ByVal grid As DataGridView, ByVal row As Integer, ByVal col As Integer, ByVal value As Object)
        Try
            If row < 0 Then Exit Sub
            While grid.RowCount <= row
                grid.Rows.Add()
            End While
            If col < grid.ColumnCount Then
                grid.Rows(row).Cells(col).Value = value
            End If
        Catch ex As Exception
            ' Ignore
        End Try
    End Sub

End Module
