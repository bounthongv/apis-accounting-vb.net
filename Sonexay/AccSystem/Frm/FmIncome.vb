Public Class FmIncome

    Private Function GetGridValue(ByVal grid As DataGridView, ByVal row As Integer, ByVal col As Integer) As String
        If row >= 0 AndAlso row < grid.RowCount AndAlso col >= 0 AndAlso col < grid.ColumnCount Then
            If grid.Rows(row).Cells(col).Value IsNot Nothing Then
                Return grid.Rows(row).Cells(col).Value.ToString()
            End If
        End If
        Return ""
    End Function

    Private Sub SetGridValue(ByVal grid As DataGridView, ByVal row As Integer, ByVal col As Integer, ByVal value As String)
        If row >= 0 AndAlso row < grid.RowCount AndAlso col >= 0 AndAlso col < grid.ColumnCount Then
            grid.Rows(row).Cells(col).Value = value
        End If
    End Sub

    Private Sub SetupGrid(ByVal grid As DataGridView, ByVal ParamArray columnHeaders() As String)
        grid.Columns.Clear()
        For i As Integer = 0 To columnHeaders.Length - 1
            grid.Columns.Add("Col" & i, columnHeaders(i))
        Next
        For i As Integer = 0 To grid.Columns.Count - 1
            grid.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        Next
    End Sub


Public Sub LoadListFG()
        FG.Rows.Clear()
        SetupGrid(FG, "ລ/ດ", "ລະຫັດ", "ເນື້ອໃນ (ພາສາລາວ)", "ເນື້ອໃນ (ພາສາອັງກິດ", "", "", "")
        
        With RSC
            Call LoadSqlData("SELECT * FROM  Ap_Rpt_Income order by  CNT ASC  ", RSC)

            If .RecordCount > 0 Then
                While Not .EOF
                    FG.Rows.Add(.AbsolutePosition.ToString(), 
                                Trim(CStr(.Fields("Rpt_ID").Value.ToString)),
                                CStr(.Fields("Description").Value.ToString),
                                Trim(CStr(.Fields("Descriptione").Value.ToString)),
                                Trim(CStr(.Fields("Chart_of_Accounts_Codes").Value.ToString)),
                                Trim(CStr(.Fields("Grp").Value.ToString)),
                                Trim(CStr(.Fields("Grp_Nme").Value.ToString)))
                    .MoveNext()
                End While
            End If
        End With
    End Sub





Public Sub MouseDownEvent()
        If FG2.CurrentRow IsNot Nothing Then
            AC_Code.Text = GetGridValue(FG2, FG2.CurrentRow.Index, 2)
            Rpt_Type.Text = GetGridValue(FG2, FG2.CurrentRow.Index, 5)
            TXTCNT.Text = GetGridValue(FG2, FG2.CurrentRow.Index, 6)
            
            If FG2.CurrentCell IsNot Nothing AndAlso FG2.CurrentCell.ColumnIndex = 5 Then
                FG2.ReadOnly = False
            Else
                FG2.ReadOnly = True
            End If
        End If
        
        BtnSearch.Visible = True
        Select Case MouseButtons
            Case Windows.Forms.MouseButtons.Right
                If FG2.CurrentCell IsNot Nothing Then
                    FG2.BeginEdit(True)
                End If
            Case Windows.Forms.MouseButtons.Left
                If FG2.CurrentCell IsNot Nothing Then
                    If FG2.CurrentCell.ColumnIndex = 2 Then
                        BtnSearch.Visible = True
                    Else
                        BtnSearch.Visible = False
                    End If
                    
                    If FG2.CurrentCell.RowIndex = FG2.RowCount - 1 Then
                        BtnMove.Visible = False
                    Else
                        BtnMove.Visible = True
                    End If
                    
                    BtnSearch.Left = CInt(FG2.Left + 100)
                    BtnSearch.Top = CInt(FG2.Top + FG2.CurrentCell.RowIndex * 22)
                    BtnMove.Top = CInt(FG2.Top + FG2.CurrentCell.RowIndex * 22)
                End If
        End Select
    End Sub





Private Sub loadBankItem()
        FG2.Rows.Clear()
        SetupGrid(FG2, "ລ/ດ", "ລະຫັດ", "ລະຫັດບັນຊີ", "ຊື່ບັນຊີ(ພາສາລາວ)", "ຊື່ບັນຊີ(ພາສາອັງກິດ)", "ສະຖານະພາບ", "CNT")
        
        With RSC
            Call LoadSqlData("SELECT * FROM  Ap_Rpt_Income_Item where Rpt_ID=N'" & TextBox1.Text & "' Order by Ac_Code ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG2.Rows.Add(.AbsolutePosition.ToString(),
                                Trim(CStr(.Fields("Rpt_ID").Value.ToString)),
                                Trim(CStr(.Fields("Ac_Code").Value.ToString)),
                                Trim(CStr(.Fields("Ac_Name").Value.ToString)),
                                Trim(CStr(.Fields("Ac_NameE").Value.ToString)),
                                Trim(CStr(.Fields("Rpt_Type").Value.ToString)),
                                Trim(CStr(.Fields("CNT").Value.ToString)))
                    .MoveNext()
                End While
            End If
        End With
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
    End Sub

Private Sub BtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch.Click
        fmShartOfAccDetail.txtSty.Text = "FmInCome"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
        FG.Focus()
        CNN.Execute("delete Ap_Rpt_Income_Item where Rpt_ID =N'" & TextBox1.Text & "' ")
        Dim i As Integer
        For i = 0 To FG2.RowCount - 1
            If GetGridValue(FG2, i, 1) = "" And GetGridValue(FG2, i, 2) = "" Then
                Exit Sub
            End If
            CNN.Execute("INSERT INTO Ap_Rpt_Income_Item( Rpt_ID,  Ac_Code , Ac_Name , Ac_NameE   , BLS , Rpt_Type ) " & _
                 "Values('" & GetGridValue(FG2, i, 1) & "', N'" & GetGridValue(FG2, i, 2) & "', N'" & GetGridValue(FG2, i, 3) & "','" & GetGridValue(FG2, i, 4) & "','" & "ALL" & "' , '" & GetGridValue(FG2, i, 5) & "')")
        Next i
    End Sub

Private Sub FmBankReportId_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button2.Enabled = False
        BtnMove.Visible = False
        BtnSearch.Visible = False
        FG.Size = New System.Drawing.Size(519, 378)
        FG.AllowUserToAddRows = False
        FG.AllowUserToDeleteRows = False
        FG.ReadOnly = True
        FG.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG.MultiSelect = False
        
        FG2.AllowUserToAddRows = False
        FG2.AllowUserToDeleteRows = False
        FG2.ReadOnly = True
        FG2.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG2.MultiSelect = False
        
        LoadListFG()
    End Sub

    Private Sub BtnMove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnMove.Click
        CNN.Execute("delete Ap_Rpt_Income_Item where Ac_Code like '" & AC_Code.Text & "%' And Rpt_ID =N'" & RPT_ID.Text & "' And Rpt_Type =N'" & Rpt_Type.Text & "' and CNT=N'" & TXTCNT.Text & "' ")
        BtnMove.Visible = False
        Call loadBankItem()
    End Sub

Private Sub FG_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FG.CellClick
        Select Case MouseButtons
            Case Windows.Forms.MouseButtons.Right
            Case Windows.Forms.MouseButtons.Left
                If e.RowIndex = FG.RowCount - 1 Then
                    Button1.Visible = False
                Else
                    Button1.Visible = True
                End If
                Button1.Top = CInt(FG.Top + e.RowIndex * 20)
        End Select
    End Sub

    Private Sub FG_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles FG.SelectionChanged
        If FG.CurrentRow IsNot Nothing Then
            TextBox1.Text = GetGridValue(FG, FG.CurrentRow.Index, 1)
            RPT_ID.Text = GetGridValue(FG, FG.CurrentRow.Index, 1)
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
        Dim i As Integer
        For i = 0 To FG.RowCount - 1
            If GetGridValue(FG, i, 1) = "" And GetGridValue(FG, i, 2) = "" Then
                Exit Sub
            End If
            CNN.Execute("Update Ap_Rpt_Income Set  Description = N'" & Apostrophe(GetGridValue(FG, i, 2)) & "' ,  Descriptione = N'" & Apostrophe(GetGridValue(FG, i, 3)) & "' , Chart_of_Accounts_Codes = N'" & GetGridValue(FG, i, 4) & "'  Where Rpt_ID = '" & GetGridValue(FG, i, 1) & "'")
        Next i
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'MsgBox("ການບັນຶກສຳເລັດຜົນ")

        'CNN.Execute("delete Ap_Rpt_Income_Item where Rpt_ID = '" & TextBox1.Text & "' ")
        'Dim i As Integer
        'For i = 1 To FG2.Rows - 1

        '    If FG2.get_TextMatrix(i, 1) = "" And FG2.get_TextMatrix(i, 2) = "" Then
        '        Exit Sub
        '    End If
        '    CNN.Execute("INSERT INTO Ap_Rpt_Income_Item( Rpt_ID,  Ac_Code , Ac_Name , Ac_NameE   , BLS , Rpt_Type ) " & _
        '         "Values('" & FG2.get_TextMatrix(i, 1) & "', N'" & FG2.get_TextMatrix(i, 2) & "', N'" & FG2.get_TextMatrix(i, 3) & "','" & FG2.get_TextMatrix(i, 4) & "','" & "ALL" & "' , '" & FG2.get_TextMatrix(i, 5) & "')")
        'Next i
    End Sub

Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If FG.CurrentRow IsNot Nothing AndAlso FG.CurrentRow.Index >= 0 Then
            FG.Rows.RemoveAt(FG.CurrentRow.Index)
        End If
        Button1.Visible = False
    End Sub
    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click

        Call Close()

    End Sub

    Private Sub RPT_ID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RPT_ID.TextChanged

    End Sub

    Private Sub Rpt_Type_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Rpt_Type.SelectedIndexChanged

    End Sub

Private Sub AC_Code_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles AC_Code.KeyPress
        If e.KeyChar = Chr(13) Then
            LoadSqlData("Select top 1 Rpt_ID , Ac_Code from Ap_Rpt_Income_Item where  Ac_Code like '" & AC_Code.Text & "%'  And Rpt_ID <> '" & RPT_ID.Text & "'  ", RSC)
            If RSC.RecordCount <> 0 Then
                MsgBox("ເລກບັນຊີ " & Trim(CStr(RSC.Fields("Ac_Code").Value.ToString)) & " ມີຢູ່ " & Trim(CStr(RSC.Fields("Rpt_ID").Value.ToString)) & " ແລ້ວ")
                Exit Sub
            End If

            CNN.Execute("delete Ap_Rpt_Income_Item where Ac_Code like '" & AC_Code.Text & "%' And Rpt_ID = '" & RPT_ID.Text & "' And Rpt_Type = '" & Rpt_Type.Text & "'  insert into Ap_Rpt_Income_Item (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type) select '" & RPT_ID.Text & "' ,  Ac_Code , Name_L , '" & Rpt_Type.Text & "' from Acc_Code where Ac_Code like '" & AC_Code.Text & "%'  ")
            
            If FG.CurrentRow IsNot Nothing Then
                TextBox1.Text = GetGridValue(FG, FG.CurrentRow.Index, 1)
            End If
            Call loadBankItem()
        End If
    End Sub

    Private Sub AC_Code_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AC_Code.TextChanged

    End Sub

Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        CNN.Execute("delete Ap_Rpt_Income_Item where  Rpt_ID=N'" & RPT_ID.Text & "'   ")
        If FG.CurrentRow IsNot Nothing Then
            TextBox1.Text = GetGridValue(FG, FG.CurrentRow.Index, 1)
        End If
        Call loadBankItem()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        CNN.Execute("delete Ap_Rpt_Item")
        With RSC
            Call LoadSqlData("SELECT * FROM  Ap_Rpt_Income_Item  Order by Ac_Code  ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    CNN.Execute("delete Ap_Rpt_Item where Ac_Code like '" & Trim(CStr(.Fields("Ac_Code").Value.ToString)) & "%' And Rpt_ID = '" & Trim(CStr(.Fields("Rpt_ID").Value.ToString)) & "' And Rpt_Type = '" & Trim(CStr(.Fields("Rpt_Type").Value.ToString)) & "' " & _
                                " insert into Ap_Rpt_Item (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type) select '" & Trim(CStr(.Fields("Rpt_ID").Value.ToString)) & "' ,  Ac_Code , Name_L , '" & Trim(CStr(.Fields("Rpt_Type").Value.ToString)) & "' from Acc_Code where Ac_Code like '" & Trim(CStr(.Fields("Ac_Code").Value.ToString)) & "%'  ")
                    .MoveNext()
                End While
            Else
            End If
        End With
        CNN.Execute("delete Ap_Rpt_Income_Item")
        CNN.Execute(" insert into Ap_Rpt_Income_Item  (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type) select Rpt_ID , Ac_Code , Ac_Name, Rpt_Type from Ap_Rpt_Item")
        MsgBox("Ok")
    End Sub
End Class