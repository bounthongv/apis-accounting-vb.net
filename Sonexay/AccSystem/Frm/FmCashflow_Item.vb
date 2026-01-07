Public Class FmCashflow_Item1

#Region "DataGridView Helper Methods"

    ''' <summary>
    ''' Gets cell value safely from DataGridView
    ''' </summary>
    Private Function GetGridValue(grid As DataGridView, row As Integer, col As Integer) As String
        If row >= 0 AndAlso row < grid.RowCount AndAlso col >= 0 AndAlso col < grid.ColumnCount Then
            If grid.Rows(row).Cells(col).Value IsNot Nothing Then
                Return grid.Rows(row).Cells(col).Value.ToString()
            End If
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Sets cell value safely in DataGridView
    ''' </summary>
    Private Sub SetGridValue(grid As DataGridView, row As Integer, col As Integer, value As Object)
        If row >= 0 AndAlso row < grid.RowCount AndAlso col >= 0 AndAlso col < grid.ColumnCount Then
            grid.Rows(row).Cells(col).Value = value
        End If
    End Sub

    ''' <summary>
    ''' Sets up DataGridView with common properties
    ''' </summary>
    Private Sub SetupGrid(grid As DataGridView, ParamArray columns() As String)
        grid.AllowUserToAddRows = False
        grid.AllowUserToDeleteRows = False
        grid.ReadOnly = False
        grid.RowHeadersVisible = False
        
        grid.Columns.Clear()
        For Each col As String In columns
            grid.Columns.Add(col, col)
        Next
        
        grid.AutoResizeColumns()
    End Sub

    ''' <summary>
    ''' Sets up single column width
    ''' </summary>
    Private Sub SetupGridColumn(grid As DataGridView, columnIndex As Integer, width As Integer)
        If columnIndex >= 0 AndAlso columnIndex < grid.ColumnCount Then
            grid.Columns(columnIndex).Width = width
        End If
    End Sub

#End Region
    Public Sub LoadListFG()
        SetupGrid(FG, "No", "Rpt_ID", "Description", "Descriptione", "Chart_of_Accounts_Codes", "Grp", "Grp_Nme")
        FG.Rows.Clear()
        
        With RSC
            Call LoadSqlData("SELECT * FROM  Ap_Rpt_Cashflow  order by Rpt_ID ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.Rows.Add(.AbsolutePosition, _
                                Trim(CStr(.Fields("Rpt_ID").Value.ToString)), _
                                CStr(.Fields("Description").Value.ToString), _
                                CStr(.Fields("Descriptione").Value.ToString), _
                                CStr(.Fields("Chart_of_Accounts_Codes").Value.ToString), _
                                CStr(.Fields("Grp").Value.ToString), _
                                CStr(.Fields("Grp_Nme").Value.ToString))
                    .MoveNext()
                End While
            End If
        End With
    End Sub

    Public Sub MouseDownEvent()
        If FG2.CurrentCell IsNot Nothing Then
            Dim rowIndex As Integer = FG2.CurrentCell.RowIndex
            AC_Code.Text = GetGridValue(FG2, rowIndex, 2)
            Rpt_Type.Text = GetGridValue(FG2, rowIndex, 5)

            If GetGridValue(FG2, rowIndex, 6) = 1 Then
                COP.Checked = True
            Else
                COP.Checked = False
            End If
            If GetGridValue(FG2, rowIndex, 7) = 1 Then
                CLa.Checked = True
            Else
                CLa.Checked = False
            End If
            If GetGridValue(FG2, rowIndex, 8) = 1 Then
                CAmt.Checked = True
            Else
                CAmt.Checked = False
            End If
            If GetGridValue(FG2, rowIndex, 9) = 1 Then
                CRem.Checked = True
            Else
                CRem.Checked = False
            End If
        End If

        If FG2.Col = 5 Then
            FG2.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
        Else
            FG2.Editable = VSFlex8U.EditableSettings.flexEDNone
        End If
        BtnSearch.Visible = True
        Select Case MouseButtons
            Case Windows.Forms.MouseButtons.Right
                FG2.EditCell()
            Case Windows.Forms.MouseButtons.Left
                If FG2.Col = 2 Then
                    BtnSearch.Visible = True
                Else
                    BtnSearch.Visible = False
                End If
                If FG2.Row = FG2.Rows - 1 Then
                    BtnMove.Visible = False
                Else
                    BtnMove.Visible = True
                End If
                BtnSearch.Left = CInt(FG2.Left + (FG2.CellLeft / 15) + (FG2.CellWidth / 22.8))
                BtnSearch.Top = CInt((FG2.CellTop / 15) + FG2.Top)
                BtnMove.Top = CInt((FG2.CellTop / 15) + FG2.Top)
        End Select
    End Sub

    Private Sub loadBankItem()
        SetupGrid(FG2, "No", "Rpt_ID", "Ac_Code", "Ac_Name", "Ac_NameE", "Rpt_Type", "Select_Open_Amt", "Select_Last_Amt", "Select_Amt", "Select_Rem_Amt")
        FG2.Rows.Clear()
        
        With RSC
            Call LoadSqlData("SELECT * FROM  Ap_Rpt_Cashflow_Item where Rpt_ID=   '" & TextBox1.Text & "' Order by Ac_Code ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG2.Rows.Add(.AbsolutePosition, _
                                 Trim(CStr(.Fields("Rpt_ID").Value.ToString)), _
                                 Trim(CStr(.Fields("Ac_Code").Value.ToString)), _
                                 Trim(CStr(.Fields("Ac_Name").Value.ToString)), _
                                 Trim(CStr(.Fields("Ac_NameE").Value.ToString)), _
                                 Trim(CStr(.Fields("Rpt_Type").Value.ToString)), _
                                 Trim(CStr(.Fields("Select_Open_Amt").Value.ToString)), _
                                 Trim(CStr(.Fields("Select_Last_Amt").Value.ToString)), _
                                 Trim(CStr(.Fields("Select_Amt").Value.ToString)), _
                                 Trim(CStr(.Fields("Select_Rem_Amt").Value.ToString)))
                    .MoveNext()
                End While
            End If
        End With
        ' DataGridView - automatic row management
    End Sub

    Private Sub FG2_CellEndEdit(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FG2.CellEndEdit
        Button2.Enabled = True
    End Sub

    Private Sub FG2_Scroll(ByVal sender As Object, ByVal e As ScrollEventArgs) Handles FG2.Scroll
        BtnSearch.Visible = False
        BtnMove.Visible = False
    End Sub

    Private Sub FG2_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FG2.CellClick
        MouseDownEvent()
        
        If FG2.CurrentCell IsNot Nothing Then
            Dim colIndex As Integer = FG2.CurrentCell.ColumnIndex
            Dim rowIndex As Integer = FG2.CurrentCell.RowIndex
            
            ' Handle column-specific logic
            If colIndex = 5 Then
                ' DataGridView doesn't need Editable property setting for column 5
            End If
            
            ' Show/hide search button based on column
            If colIndex = 2 Then
                BtnSearch.Visible = True
            Else
                BtnSearch.Visible = False
            End If
            
            ' Show/hide move button based on row
            If rowIndex = FG2.RowCount - 1 Then
                BtnMove.Visible = False
            Else
                BtnMove.Visible = True
            End If
            
            ' Position buttons
            If FG2.CurrentCell IsNot Nothing Then
                BtnSearch.Left = FG2.Left + FG2.CurrentCell.OwningColumn.Left
                BtnSearch.Top = FG2.CurrentCell.OwningRow.Top + FG2.Top
                BtnMove.Top = FG2.CurrentCell.OwningRow.Top + FG2.Top
            End If
        End If
    End Sub

    Private Sub BtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch.Click
        fmShartOfAccDetail.txtSty.Text = "FmBLS"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
        FG.Focus()
    End Sub

    Private Sub FmCashflow_Item_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button2.Enabled = False
        BtnMove.Visible = False
        BtnSearch.Visible = False
        FG.Size = New System.Drawing.Size(409, 378)
        
        ' Configure DataGridView properties
        ConfigureDataGridViewProperties()
        
        LoadListFG()
    End Sub

    ''' <summary>
    ''' Configures DataGridView properties for all grids in form
    ''' </summary>
    Private Sub ConfigureDataGridViewProperties()
        ' Configure FG
        FG.AllowUserToAddRows = False
        FG.AllowUserToDeleteRows = False
        FG.AllowUserToResizeColumns = True
        FG.AllowUserToResizeRows = False
        FG.ReadOnly = True
        FG.RowHeadersVisible = False
        FG.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG.MultiSelect = False
        
        ' Configure FG2
        FG2.AllowUserToAddRows = False
        FG2.AllowUserToDeleteRows = False
        FG2.AllowUserToResizeColumns = True
        FG2.AllowUserToResizeRows = False
        FG2.ReadOnly = False
        FG2.RowHeadersVisible = False
        FG2.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
        FG2.MultiSelect = False
    End Sub

    Private Sub BtnMove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnMove.Click
        CNN.Execute("delete Ap_Rpt_Cashflow_Item where Ac_Code like '" & AC_Code.Text & "%' And Rpt_ID = '" & RPT_ID.Text & "' And Rpt_Type = '" & Rpt_Type.Text & "' ")
        BtnMove.Visible = False
        Call loadBankItem()
    End Sub

    Private Sub FG_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FG.CellClick
        Select Case MouseButtons
            Case Windows.Forms.MouseButtons.Right
            Case Windows.Forms.MouseButtons.Left
                If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.RowIndex = FG.RowCount - 1 Then
                    Button1.Visible = False
                Else
                    Button1.Visible = True
                End If
                If FG.CurrentCell IsNot Nothing Then
                    Button1.Top = FG.CurrentCell.OwningRow.Top + FG.Top
                End If
        End Select
    End Sub

    Private Sub FG_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelectionChanged
        If FG.CurrentCell IsNot Nothing Then
            Dim rowIndex As Integer = FG.CurrentCell.RowIndex
            TextBox1.Text = GetGridValue(FG, rowIndex, 1)
            RPT_ID.Text = GetGridValue(FG, rowIndex, 1)
        End If

        'MsgBox(RPT_ID.Text)
        Call loadBankItem()

        Button2.Enabled = False
        BtnMove.Visible = False
        BtnSearch.Visible = False
    End Sub

    Private Sub BtnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnEdit.Click
        Button2.Enabled = False
        BtnMove.Visible = False
        BtnSearch.Visible = False
        MsgBox("ການບັນຶກສຳເລັດຜົນ")
        CNN.Execute("delete Ap_Rpt_Cashflow")
        Dim i As Integer
        For i = 0 To FG.RowCount - 1
            If GetGridValue(FG, i, 1) = "" And GetGridValue(FG, i, 2) = "" Then
                Exit Sub
            End If
            CNN.Execute("INSERT INTO Ap_Rpt_Cashflow( Rpt_ID,  Description , Descriptione  ,Chart_of_Accounts_Codes , Grp , Grp_Nme ) " & _
                                    "Values('" & GetGridValue(FG, i, 1) & "', N'" & Apostrophe(GetGridValue(FG, i, 2)) & "','" & Apostrophe(GetGridValue(FG, i, 3)) & "',N'" & Apostrophe(GetGridValue(FG, i, 4)) & "' ,N'" & Apostrophe(GetGridValue(FG, i, 5)) & "' ,N'" & Apostrophe(GetGridValue(FG, i, 6)) & "')")
        Next i
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        MsgBox("ການບັນຶກສຳເລັດຜົນ")
        CNN.Execute("delete Ap_Rpt_Cashflow_Item where Rpt_ID = '" & TextBox1.Text & "' ")
        Dim i As Integer
        For i = 0 To FG2.RowCount - 1

            If GetGridValue(FG2, i, 1) = "" And GetGridValue(FG2, i, 2) = "" Then
                Exit Sub
            End If
            CNN.Execute("INSERT INTO Ap_Rpt_Cashflow_Item( Rpt_ID,  Ac_Code , Ac_Name , Ac_NameE ,Amt_Dr , Amt_Cr , BLS  , Rpt_Type) " & _
                 "Values('" & GetGridValue(FG2, i, 1) & "', N'" & GetGridValue(FG2, i, 2) & "', N'" & GetGridValue(FG2, i, 3) & "','" & GetGridValue(FG2, i, 4) & "','" & CDbl(0) & "','" & CDbl(0) & "','" & "ALL" & "' ,'" & GetGridValue(FG2, i, 5) & "')")
        Next i
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.RowIndex >= 0 Then
            FG.Rows.RemoveAt(FG.CurrentCell.RowIndex)
        End If
        Button1.Visible = False
    End Sub

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Call Close()
    End Sub

    Private Sub AC_Code_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles AC_Code.KeyPress
        Dim OP_Amt, Amt, Rem_Amt, Last_Amt As String
        OP_Amt = 0
        Amt = 0
        Rem_Amt = 0
        Last_Amt = 0
        If COP.Checked = True Then
            OP_Amt = 1
        End If
        If CAmt.Checked = True Then
            Amt = 1
        End If
        If CRem.Checked = True Then
            Rem_Amt = 1
        End If
        If CLa.Checked = True Then
            Last_Amt = 1
        End If
        If e.KeyChar = Chr(13) Then
            CNN.Execute("delete Ap_Rpt_Cashflow_Item where Ac_Code like '" & AC_Code.Text & "%' And Rpt_ID = '" & RPT_ID.Text & "' And Rpt_Type = '" & Rpt_Type.Text & "'  insert into Ap_Rpt_Cashflow_Item (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type , Select_Open_Amt , Select_Amt , Select_Rem_Amt , Select_Last_Amt) select '" & RPT_ID.Text & "' ,  Ac_Code , Name_L , '" & Rpt_Type.Text & "' , " & OP_Amt & " , " & Amt & " , " & Rem_Amt & " , " & Last_Amt & " from Acc_Code where Ac_Code like '" & AC_Code.Text & "%'  and acc_type=N'ບັນຊີແມ່ (P)'  ")
            If FG.CurrentCell IsNot Nothing Then
                TextBox1.Text = GetGridValue(FG, FG.CurrentCell.RowIndex, 1)
            End If
            Call loadBankItem()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        CNN.Execute("delete Ap_Rpt_Cashflow_Item where  Rpt_ID = '" & RPT_ID.Text & "'   ")
        If FG.CurrentCell IsNot Nothing Then
            TextBox1.Text = GetGridValue(FG, FG.CurrentCell.RowIndex, 1)
        End If
        Call loadBankItem()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        ' DataGridView navigation equivalent
        If FG.RowCount > 18 AndAlso FG.ColumnCount > 2 Then
            FG.CurrentCell = FG(18, 2)
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

    End Sub

    Private Sub AC_Code_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AC_Code.TextChanged

    End Sub
 

    Private Sub FG2_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG2.SelectionChanged

    End Sub
End Class