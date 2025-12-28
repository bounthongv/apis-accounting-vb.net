Public Class FmLangaugeSeting
    Dim sql As String
    Private Sub FmLangaugeSeting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        FG.Cols = 5

        FG.FormatString = "^No    |<Langauge ID |< Langauge Name(Laos)                                         |< Langauge Name (English)                                        |< ToolBok      "
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
        FG.Rows = 1
        Call LoadAcData("Select * from Langauge  " & sql & "  order by   LngID ", RSC)
        With RSC
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("LngID").Value)) & _
                                                        "" & vbTab & Trim(CStr(.Fields("LngL").Value.ToString)) & _
                                                         "" & vbTab & Trim(CStr(.Fields("LngE").Value.ToString)) & _
                                                             "" & vbTab & Trim(CStr(.Fields("Toolbox").Value.ToString)))
                    .MoveNext()
                End While
            Else
                FG.Rows = 2
            End If
        End With
        FG.AllowUserResizing = VSFlex8U.AllowUserResizeSettings.flexResizeColumns
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        For i = 1 To FG.Rows - 1
            conn.Execute("Update Langauge set LngL = '" & FG.get_TextMatrix(i, 2) & "'  , LngE = '" & FG.get_TextMatrix(i, 3) & "' Where LngID =  '" & FG.get_TextMatrix(i, 1) & "'  ")
        Next i
        Call LaodList()
        LoadLng()

    End Sub

    Private Sub FG_DblClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles FG.DblClick
        If FG.Row > 0 Then
            Panel1.Visible = True
            LngL.Focus()
        End If

    End Sub

    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelChange
        If FG.Row > 0 Then

            LngId.Text = FG.get_TextMatrix(FG.Row, 1)
            LngL.Text = FG.get_TextMatrix(FG.Row, 2)
            LngE.Text = FG.get_TextMatrix(FG.Row, 3)
            Toolbox2.Text = FG.get_TextMatrix(FG.Row, 4)
        End If
  

        'If FG.Col = 2 Or FG.Col = 3 Then
        '    FG.FocusRect = VSFlex8U.FocusRectSettings.flexFocusInset
        '    'FG.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
        'Else

        '    'FG.Editable = VSFlex8U.EditableSettings.flexEDNone
        '    FG.FocusRect = VSFlex8U.FocusRectSettings.flexFocusLight
        'End If
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
        conn.Execute("Update Langauge set LngL = '" & LngL.Text & "'  , LngE = '" & LngE.Text & "' Where LngID =  '" & LngId.Text & "'  ")
        FG.set_TextMatrix(FG.Row, 1, LngId.Text)
        FG.set_TextMatrix(FG.Row, 2, LngL.Text)
        FG.set_TextMatrix(FG.Row, 3, LngE.Text)
        FG.set_TextMatrix(FG.Row, 4, Toolbox2.Text)
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
                conn.Execute("Update Langauge set LngL = '" & LngL.Text & "'  , LngE = '" & LngE.Text & "' Where LngID =  '" & LngId.Text & "'  ")
                FG.set_TextMatrix(FG.Row, 1, LngId.Text)
                FG.set_TextMatrix(FG.Row, 2, LngL.Text)
                FG.set_TextMatrix(FG.Row, 3, LngE.Text)
                FG.set_TextMatrix(FG.Row, 4, Toolbox2.Text)
                Panel1.Visible = False
            End If
        End If
    End Sub

    Private Sub LngE_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LngE.TextChanged
    
    End Sub
End Class