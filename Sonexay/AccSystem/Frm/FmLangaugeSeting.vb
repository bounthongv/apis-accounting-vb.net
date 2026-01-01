Public Class FmLangaugeSeting
    Dim sql As String
    Private Sub FmLangaugeSeting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Setup DataGridView Columns
        FG.Columns.Clear()
        FG.Columns.Add("No", "No")
        FG.Columns.Add("LangID", "Langauge ID")
        FG.Columns.Add("LangL", "Langauge Name(Laos)")
        FG.Columns.Add("LangE", "Langauge Name (English)")
        FG.Columns.Add("Toolbox", "ToolBox")

        ' Set Column Widths (Approximation based on original format string spaces)
        FG.Columns(0).Width = 50
        FG.Columns(1).Width = 100
        FG.Columns(2).Width = 300
        FG.Columns(3).Width = 300
        FG.Columns(4).Width = 150

        FG.AllowUserToAddRows = False
        FG.ReadOnly = True
        FG.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG.MultiSelect = False

        Call loadToolbox()
        'Call LaodList()
    End Sub

    Private Sub loadToolbox()
        Toolbox.Items.Clear()

        Toolbox.Items.Add("ToolBox All")
        LoadAcData("select Toolbox  from  Langauge group by Toolbox  Order by Toolbox", RSC)
        With RSC
            Do Until .EOF = True
                Toolbox.Items.Add((.Fields("Toolbox").Value))
                .MoveNext()
            Loop
        End With
        Toolbox.SelectedIndex = 0
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call LaodList()
        LoadLng()

    End Sub

    Public Sub LaodList()
        sql = "where LngID <>'' "
        If txtSearch.Text <> "" Then
            sql = sql & " and ( LngL Like  '%" & txtSearch.Text & "%' Or  LngE Like  '%" & txtSearch.Text & "%'  ) "
        End If
        If Toolbox.SelectedIndex > 0 Then
            sql = sql & " And ( Toolbox =  '" & Toolbox.Text & "' ) "
        End If
        
        FG.Rows.Clear()
        
        Call LoadAcData("Select * from Langauge  " & sql & "  order by   LngID ", RSC)
        With RSC
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.Rows.Add(.AbsolutePosition, _
                                Trim(CStr(.Fields("LngID").Value)), _
                                Trim(CStr(.Fields("LngL").Value.ToString)), _
                                Trim(CStr(.Fields("LngE").Value.ToString)), _
                                Trim(CStr(.Fields("Toolbox").Value.ToString)))
                    .MoveNext()
                End While
            End If
        End With
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        For i As Integer = 0 To FG.Rows.Count - 1
            ' Note: DataGridView is 0-based. 
            ' Col 1 (LangID), Col 2 (LangL), Col 3 (LangE) match the indices defined in Load event.
            conn.Execute("Update Langauge set LngL = '" & FG.Rows(i).Cells(2).Value.ToString() & "'  , LngE = '" & FG.Rows(i).Cells(3).Value.ToString() & "' Where LngID =  '" & FG.Rows(i).Cells(1).Value.ToString() & "'  ")
        Next i
        Call LaodList()
        LoadLng()

    End Sub

    Private Sub FG_DblClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles FG.DoubleClick, FG.CellDoubleClick
        If FG.CurrentRow IsNot Nothing AndAlso FG.CurrentRow.Index >= 0 Then
            Panel1.Visible = True
            LngL.Focus()
        End If
    End Sub

    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelectionChanged, FG.Click
        If FG.CurrentRow IsNot Nothing AndAlso FG.CurrentRow.Index >= 0 Then
            LngId.Text = FG.CurrentRow.Cells(1).Value.ToString()
            LngL.Text = FG.CurrentRow.Cells(2).Value.ToString()
            LngE.Text = FG.CurrentRow.Cells(3).Value.ToString()
            Toolbox2.Text = FG.CurrentRow.Cells(4).Value.ToString()
        End If
    End Sub

    Private Sub txtSearch_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSearch.KeyPress
        If e.KeyChar = Chr(13) Then
            LaodList()
        End If
    End Sub

    Private Sub txtSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.TextChanged

    End Sub

    Private Sub Toolbox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Toolbox.SelectedIndexChanged
        'Call LaodList()
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Panel1.Visible = False
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If FG.CurrentRow Is Nothing Then Exit Sub
        
        conn.Execute("Update Langauge set LngL = '" & LngL.Text & "'  , LngE = '" & LngE.Text & "' Where LngID =  '" & LngId.Text & "'  ")
        
        FG.CurrentRow.Cells(1).Value = LngId.Text
        FG.CurrentRow.Cells(2).Value = LngL.Text
        FG.CurrentRow.Cells(3).Value = LngE.Text
        FG.CurrentRow.Cells(4).Value = Toolbox2.Text
        
        Panel1.Visible = False
    End Sub

    Private Sub LngL_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles LngL.KeyPress
        If e.KeyChar = Chr(13) Then
            LngE.Focus()
        End If
    End Sub

    Private Sub LngL_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LngL.TextChanged

    End Sub

    Private Sub LngE_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles LngE.KeyPress
        If e.KeyChar = Chr(13) Then
            If CheckBox1.Checked = True Then
                If FG.CurrentRow IsNot Nothing Then
                    conn.Execute("Update Langauge set LngL = '" & LngL.Text & "'  , LngE = '" & LngE.Text & "' Where LngID =  '" & LngId.Text & "'  ")
                    
                    FG.CurrentRow.Cells(1).Value = LngId.Text
                    FG.CurrentRow.Cells(2).Value = LngL.Text
                    FG.CurrentRow.Cells(3).Value = LngE.Text
                    FG.CurrentRow.Cells(4).Value = Toolbox2.Text
                    
                    Panel1.Visible = False
                End If
            End If
        End If
    End Sub

    Private Sub LngE_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LngE.TextChanged
    
    End Sub
End Class