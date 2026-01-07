Imports System.Windows.Forms
Imports System.Drawing

Public Class FmIncome_Old


    Public Sub LoadListFG()
        FG.Rows.Clear()
        FG.Columns.Clear()
        
        ' Setup columns
        FG.Columns.Add("Col0", "ລ/ດ")
        FG.Columns.Add("Col1", "ລະຫັດ")
        FG.Columns.Add("Col2", "ເນື້ອໃນ (ພາສາລາວ)")
        FG.Columns.Add("Col3", "ເນື້ອໃນ (ພາສາອັງກິດ)")
        FG.Columns.Add("Col4", "")
        FG.Columns.Add("Col5", "")
        FG.Columns.Add("Col6", "")

        Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM Ap_Rpt_Income_Old order by CNT ASC")

        If dt.Rows.Count > 0 Then
            For i As Integer = 0 To dt.Rows.Count - 1
                FG.Rows.Add((i + 1).ToString(),
                            GetStr(dt.Rows(i)("Rpt_ID")),
                            GetStr(dt.Rows(i)("Description")),
                            GetStr(dt.Rows(i)("Descriptione")),
                            GetStr(dt.Rows(i)("Chart_of_Accounts_Codes")),
                            GetStr(dt.Rows(i)("Grp")),
                            GetStr(dt.Rows(i)("Grp_Nme")))
            Next
        End If
        
        ' Add empty row
        FG.Rows.Add()
    End Sub





Public Sub MouseDownEvent()
        If FG2.CurrentCell IsNot Nothing AndAlso FG2.CurrentCell.RowIndex >= 0 AndAlso FG2.CurrentCell.RowIndex < FG2.Rows.Count - 1 Then
            AC_Code.Text = GetStr(FG2.Rows(FG2.CurrentCell.RowIndex).Cells(2).Value)
            Rpt_Type.Text = GetStr(FG2.Rows(FG2.CurrentCell.RowIndex).Cells(5).Value)
            TXTCNT.Text = GetStr(FG2.Rows(FG2.CurrentCell.RowIndex).Cells(6).Value)
        End If
        
        If FG2.CurrentCell IsNot Nothing AndAlso FG2.CurrentCell.ColumnIndex = 5 Then
            FG2.ReadOnly = False
        Else
            FG2.ReadOnly = True
        End If
        
        BtnSearch.Visible = True
        Select Case MouseButtons
            Case Windows.Forms.MouseButtons.Right
                If FG2.CurrentCell IsNot Nothing AndAlso FG2.CurrentCell.RowIndex >= 0 Then
                    FG2.BeginEdit(False)
                End If
            Case Windows.Forms.MouseButtons.Left
                If FG2.CurrentCell IsNot Nothing AndAlso FG2.CurrentCell.ColumnIndex = 2 Then
                    BtnSearch.Visible = True
                Else
                    BtnSearch.Visible = False
                End If
                If FG2.CurrentCell IsNot Nothing AndAlso FG2.CurrentCell.RowIndex = FG2.Rows.Count - 1 Then
                    BtnMove.Visible = False
                Else
                    BtnMove.Visible = True
                End If
                If FG2.CurrentCell IsNot Nothing AndAlso FG2.CurrentCell.RowIndex >= 0 Then
                    Dim cellRect As Rectangle = FG2.GetCellDisplayRectangle(FG2.CurrentCell.ColumnIndex, FG2.CurrentCell.RowIndex, False)
                    BtnSearch.Left = FG2.Left + cellRect.Left
                    BtnSearch.Top = FG2.Top + cellRect.Top
                    BtnMove.Top = FG2.Top + cellRect.Top
                End If
        End Select
    End Sub





Private Sub loadBankItem()
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

        Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM Ap_Rpt_Income_Item_Old where Rpt_ID=N'" & TextBox1.Text & "' Order by Ac_Code")
        
        If dt.Rows.Count > 0 Then
            For i As Integer = 0 To dt.Rows.Count - 1
                FG2.Rows.Add((i + 1).ToString(),
                            GetStr(dt.Rows(i)("Rpt_ID")),
                            GetStr(dt.Rows(i)("Ac_Code")),
                            GetStr(dt.Rows(i)("Ac_Name")),
                            GetStr(dt.Rows(i)("Ac_NameE")),
                            GetStr(dt.Rows(i)("Rpt_Type")),
                            GetStr(dt.Rows(i)("CNT")))
            Next
        End If
        
        ' Add empty row
        FG2.Rows.Add()
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
        fmShartOfAccDetail.txtSty.Text = "FmInCome"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
        FG.Focus()
        DbHelper.ExecuteNonQuery("delete Ap_Rpt_Income_Item_Old where Rpt_ID =N'" & TextBox1.Text & "' ")
        
        For i As Integer = 0 To FG2.Rows.Count - 2 ' Skip last empty row
            If GetStr(FG2.Rows(i).Cells(1).Value) = "" AndAlso GetStr(FG2.Rows(i).Cells(2).Value) = "" Then
                Exit For
            End If
            DbHelper.ExecuteNonQuery("INSERT INTO Ap_Rpt_Income_Item_Old( Rpt_ID,  Ac_Code , Ac_Name , Ac_NameE   , BLS , Rpt_Type ) " & _
                 "Values('" & GetStr(FG2.Rows(i).Cells(1).Value) & "', N'" & GetStr(FG2.Rows(i).Cells(2).Value) & "', N'" & GetStr(FG2.Rows(i).Cells(3).Value) & "','" & GetStr(FG2.Rows(i).Cells(4).Value) & "','" & "ALL" & "' , '" & GetStr(FG2.Rows(i).Cells(5).Value) & "')")
        Next
    End Sub

Private Sub FmBankReportId_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button2.Enabled = False
        BtnMove.Visible = False
        BtnSearch.Visible = False
        FG.Size = New System.Drawing.Size(519, 378)
        LoadListFG()
        FG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    End Sub

Private Sub BtnMove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnMove.Click
        DbHelper.ExecuteNonQuery("delete Ap_Rpt_Income_Item_Old where Ac_Code like '" & AC_Code.Text & "%' And Rpt_ID =N'" & RPT_ID.Text & "' And Rpt_Type =N'" & Rpt_Type.Text & "' and CNT=N'" & TXTCNT.Text & "' ")
        BtnMove.Visible = False
        loadBankItem()
    End Sub

Private Sub FG_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles FG.MouseDown
        Select Case MouseButtons
            Case Windows.Forms.MouseButtons.Right
            Case Windows.Forms.MouseButtons.Left
                If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.RowIndex = FG.Rows.Count - 1 Then
                    Button1.Visible = False
                Else
                    Button1.Visible = True
                End If
                If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.RowIndex >= 0 Then
                    Dim cellRect As Rectangle = FG.GetCellDisplayRectangle(0, FG.CurrentCell.RowIndex, False)
                    Button1.Top = FG.Top + cellRect.Top
                End If
        End Select
    End Sub

Private Sub FG_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles FG.SelectionChanged
        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.RowIndex >= 0 AndAlso FG.CurrentCell.RowIndex < FG.Rows.Count - 1 Then
            TextBox1.Text = GetStr(FG.Rows(FG.CurrentCell.RowIndex).Cells(1).Value)
            RPT_ID.Text = GetStr(FG.Rows(FG.CurrentCell.RowIndex).Cells(1).Value)
            loadBankItem()
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
        
        For i As Integer = 0 To FG.Rows.Count - 2 ' Skip last empty row
            If GetStr(FG.Rows(i).Cells(1).Value) = "" AndAlso GetStr(FG.Rows(i).Cells(2).Value) = "" Then
                Exit For
            End If
            DbHelper.ExecuteNonQuery("Update Ap_Rpt_Income_Old Set  Description = N'" & Apostrophe(GetStr(FG.Rows(i).Cells(2).Value)) & "' ,  Descriptione = N'" & Apostrophe(GetStr(FG.Rows(i).Cells(3).Value)) & "' , Chart_of_Accounts_Codes = N'" & GetStr(FG.Rows(i).Cells(4).Value) & "'  Where Rpt_ID = '" & GetStr(FG.Rows(i).Cells(1).Value) & "'")
        Next
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
        If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.RowIndex >= 0 AndAlso FG.CurrentCell.RowIndex < FG.Rows.Count - 1 Then
            FG.Rows.RemoveAt(FG.CurrentCell.RowIndex)
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
            Dim dt As DataTable = DbHelper.GetDataTable("Select top 1 Rpt_ID , Ac_Code from Ap_Rpt_Income_Item_Old where  Ac_Code like '" & AC_Code.Text & "%'  And Rpt_ID <> '" & RPT_ID.Text & "'")
            If dt.Rows.Count > 0 Then
                MsgBox("ເລກບັນຊີ " & GetStr(dt.Rows(0)("Ac_Code")) & " ມີຢູ່ " & GetStr(dt.Rows(0)("Rpt_ID")) & " ແລ້ວ")
                Exit Sub
            End If

            DbHelper.ExecuteNonQuery("delete Ap_Rpt_Income_Item_Old where Ac_Code like '" & AC_Code.Text & "%' And Rpt_ID = '" & RPT_ID.Text & "' And Rpt_Type = '" & Rpt_Type.Text & "'  insert into Ap_Rpt_Income_Item_Old (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type) select '" & RPT_ID.Text & "' ,  Ac_Code , Name_L , '" & Rpt_Type.Text & "' from Acc_Code where Ac_Code like '" & AC_Code.Text & "%' and acc_type=N'ບັນຊີແມ່ (P)' ")
            
            If FG.CurrentCell IsNot Nothing AndAlso FG.CurrentCell.RowIndex >= 0 AndAlso FG.CurrentCell.RowIndex < FG.Rows.Count - 1 Then
                TextBox1.Text = GetStr(FG.Rows(FG.CurrentCell.RowIndex).Cells(1).Value)
            End If
            loadBankItem()
        End If
    End Sub

    Private Sub AC_Code_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AC_Code.TextChanged

    End Sub

Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        DbHelper.ExecuteNonQuery("delete Ap_Rpt_Income_Item_Old where  Rpt_ID=N'" & RPT_ID.Text & "'   ")
        If FG.CurrentRowIndex >= 0 AndAlso FG.CurrentRowIndex < FG.Rows.Count - 1 Then
            TextBox1.Text = GetStr(FG.Rows(FG.CurrentRowIndex).Cells(1).Value)
        End If
        loadBankItem()
    End Sub

Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        DbHelper.ExecuteNonQuery("delete Ap_Rpt_Item")
        
        Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM Ap_Rpt_Income_Item_Old Order by Ac_Code")
        If dt.Rows.Count > 0 Then
            For i As Integer = 0 To dt.Rows.Count - 1
                DbHelper.ExecuteNonQuery("delete Ap_Rpt_Item where Ac_Code like '" & GetStr(dt.Rows(i)("Ac_Code")) & "%' And Rpt_ID = '" & GetStr(dt.Rows(i)("Rpt_ID")) & "' And Rpt_Type = '" & GetStr(dt.Rows(i)("Rpt_Type")) & "' " & _
                            " insert into Ap_Rpt_Item (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type) select '" & GetStr(dt.Rows(i)("Rpt_ID")) & "' ,  Ac_Code , Name_L , '" & GetStr(dt.Rows(i)("Rpt_Type")) & "' from Acc_Code where Ac_Code like '" & GetStr(dt.Rows(i)("Ac_Code")) & "%'  ")
            Next
        End If
        
        DbHelper.ExecuteNonQuery("delete Ap_Rpt_Income_Item_Old")
        DbHelper.ExecuteNonQuery(" insert into Ap_Rpt_Income_Item_Old  (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type) select Rpt_ID , Ac_Code , Ac_Name, Rpt_Type from Ap_Rpt_Item")
        MsgBox("Ok")
    End Sub
End Class