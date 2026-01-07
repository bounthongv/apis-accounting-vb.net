Public Class FmAmtStatus_Item

    Private Sub FmAmtStatus_Item_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FG2.FormatString = "^ລ/ດ |< ລະຫັດ |<ລະຫັດບັນຊີ  |<ຊື່ບັນຊີ(ພາສາລາວ)                   |<|<Status      |<Open|<Amt|<Rem"
        FG.FormatString = "^ລ/ດ|< ລະຫັດ|<ເນື້ອໃນ (ພາສາລາວ)           |<       |<Amt1|<Amt2|<Amt3|<Amt4|<Amt5|<Amt6|<"
        Call LoadListFG()
        BtnEdit.Enabled = True
    End Sub

    Public Sub LoadListFG()
        FG.Rows = 1
        With RSC
            'Call LoadData("select *  from Open_jn WHERE ac_code<>''  " & Sql & "order by ac_code", RSC)

            Call LoadSqlData(" SELECT * FROM  AP_Rpt_Amt_Status  order by Rpt_ID ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("Rpt_ID").Value.ToString)) & _
                                    "" & vbTab & (CStr(.Fields("Description").Value.ToString)) & _
                                       "" & vbTab & (CStr(.Fields("Descriptione").Value.ToString)) & _
                                           "" & vbTab & "X" & _
                                           "" & vbTab & "X" & _
                                                "" & vbTab & "X" & _
                                         "" & vbTab & "X" & _
                                          "" & vbTab & "X" & _
                                                "" & vbTab & "X" & _
                                        "" & vbTab & (CStr(.Fields("cnt").Value.ToString)))
                    .MoveNext()
                End While
            Else
            End If
        End With
        FG.Rows = CDbl(FG.Rows) + 1
    End Sub

    Private Sub FG_CellMouseDown(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) Handles FG.CellMouseDown
        If e.ColumnIndex > 3 And e.ColumnIndex < 10 Then
            RPT_ID.Text = FG.Rows(e.RowIndex).Cells(1).Value.ToString() & "/0" & e.ColumnIndex - 3
            Call loadBankItem()
        Else
            'RPT_ID.Text = ""
            FG2.Rows = 1
            FG2.Rows = 2

        End If
    End Sub
    Private Sub loadBankItem()
        FG2.Rows = 1
        With RSC
            'Call LoadData("select *  from Open_jn WHERE ac_code<>''  " & Sql & "order by ac_code", RSC)

            Call LoadSqlData("SELECT * FROM  AP_Rpt_AmtStatus_Item where Rpt_ID=   '" & RPT_ID.Text & "' Order by Ac_Code ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG2.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("Rpt_ID").Value.ToString)) & _
                    "" & vbTab & Trim(CStr(.Fields("Ac_Code").Value.ToString)) & _
                       "" & vbTab & Trim(CStr(.Fields("Ac_Name").Value.ToString)) & _
                            "" & vbTab & Trim(CStr(.Fields("Ac_NameE").Value.ToString)) & _
                                "" & vbTab & Trim(CStr(.Fields("Rpt_Type").Value.ToString)) & _
                                    "" & vbTab & Trim(CStr(.Fields("Select_Open_Amt").Value.ToString)) & _
                                            "" & vbTab & Trim(CStr(.Fields("Select_Amt").Value.ToString)) & _
                    "" & vbTab & Trim(CStr(.Fields("Select_Rem_Amt").Value.ToString)))
                    .MoveNext()
                End While
            Else
            End If
        End With
        FG2.Rows = CDbl(FG2.Rows) + 1
    End Sub


    Private Sub FG_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelectionChanged
        If FG.CurrentRow IsNot Nothing AndAlso FG.CurrentCell IsNot Nothing Then
            RPT_ID.Text = FG.CurrentRow.Cells(1).Value.ToString()
            If FG.CurrentCell.ColumnIndex = 2 Or FG.CurrentCell.ColumnIndex = 3 Then
                ' DataGridView is always editable by default, no equivalent to VSFlexGrid.Editable needed
                'MsgBox(FG.CurrentCell.ColumnIndex)
            Else
                ' DataGridView is always editable by default, no equivalent to VSFlexGrid.Editable needed
            End If
        End If
        'MsgBox(RPT_ID.Text)
    End Sub

    Private Sub AC_Code_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles AC_Code.KeyPress
        Dim OP_Amt, Amt, Rem_Amt As String
        OP_Amt = 0
        Amt = 0
        Rem_Amt = 0

        If COP.Checked = True Then
            OP_Amt = 1
        End If
        If CAmt.Checked = True Then

            Amt = 1
        End If
        If CRem.Checked = True Then
            Rem_Amt = 1
        End If
        If e.KeyChar = Chr(13) Then
            Dim MUTY As String = Rpt_Type.Text & "&" & ComboBox1.Text
            Dim s As String = "delete AP_Rpt_AmtStatus_Item where Ac_Code like '" & AC_Code.Text & "%' And Rpt_ID = '" & RPT_ID.Text & "' And Rpt_Type = '" & Apostrophe(MUTY) & "' " & _
                        " insert into AP_Rpt_AmtStatus_Item (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type , Select_Open_Amt , Select_Amt , Select_Rem_Amt ) " & _
                        " select '" & RPT_ID.Text & "' ,  Ac_Code , Name_L , '" & Apostrophe(MUTY) & "' , " & OP_Amt & " , " & Amt & " , " & Rem_Amt & "  " & _
                        " from Acc_Code where Ac_Code like '" & AC_Code.Text & "%'  "
            CNN.Execute(s)
            'TextBox1.Text = FG.get_TextMatrix(FG.Row, 1)
            Call loadBankItem()
        End If
    End Sub

    Private Sub AC_Code_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AC_Code.TextChanged

    End Sub

    Private Sub BtnMove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnMove.Click
        CNN.Execute("delete AP_Rpt_AmtStatus_Item where Ac_Code like '" & AC_Code.Text & "%' And Rpt_ID = '" & RPT_ID.Text & "' And Rpt_Type = '" & Rpt_Type.Text & "' ")
        BtnMove.Visible = False
        Call loadBankItem()
    End Sub

    Private Sub FG2_CellMouseDown(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) Handles FG2.CellMouseDown
        MouseDownEvent()
    End Sub

    Private Sub FG2_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG2.SelectionChanged

    End Sub
    Public Sub MouseDownEvent()
        AC_Code.Text = FG2.get_TextMatrix(FG2.Row, 2)
        Rpt_Type.Text = FG2.get_TextMatrix(FG2.Row, 5)

        If FG2.get_TextMatrix(FG2.Row, 6) = 1 Then
            COP.Checked = True
        Else
            COP.Checked = False
        End If
      
        If FG2.get_TextMatrix(FG2.Row, 8) = 1 Then
            CAmt.Checked = True
        Else
            CAmt.Checked = False
        End If
        If FG2.get_TextMatrix(FG2.Row, 9) = 1 Then
            CRem.Checked = True
        Else
            CRem.Checked = False
        End If

        If FG2.Col = 5 Then
            FG2.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
            'MsgBox(FG.Col)
        Else
            FG2.Editable = VSFlex8U.EditableSettings.flexEDNone
        End If
        BtnSearch.Visible = True
        Select Case MouseButtons
            Case Windows.Forms.MouseButtons.Right
                FG2.EditCell()
            Case Windows.Forms.MouseButtons.Left
                'MsgBox(FG.Col)
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
                'BtnMove.Left = CInt(FG.Left + (FG.CellLeft / 15) + (FG.CellWidth / 22.8))
                BtnMove.Top = CInt((FG2.CellTop / 15) + FG2.Top)
        End Select
    End Sub

    Private Sub BtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch.Click
        fmShartOfAccDetail.txtSty.Text = "FmAmtStatus_Item"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
        FG.Focus()
    End Sub

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'CNN.Execute("delete Ap_Rpt_BLS_Item_Old where  Rpt_ID =N'" & RPT_ID.Text & "'   ")
        CNN.Execute("delete AP_Rpt_AmtStatus_Item where   Rpt_ID =N'" & RPT_ID.Text & "' And Rpt_Type =N'" & Rpt_Type.Text & "' ")

        'TextBox1.Text = FG.get_TextMatrix(FG.Row, 1)
        Call loadBankItem()
    End Sub

    Private Sub BtnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnEdit.Click

        BtnMove.Visible = False
        BtnSearch.Visible = False
        MsgBox("ການບັນທຶກສຳເລັດຜົນ")
        'CNN.Execute("delete Ap_Rpt_BLS_Old")
        Dim i As Integer
        For i = 1 To FG.Rows - 1
            If FG.get_TextMatrix(i, 1) = "" And FG.get_TextMatrix(i, 2) = "" Then
                Exit Sub
            End If

            CNN.Execute("Update AP_Rpt_Amt_Status Set  Description = N'" & Apostrophe(FG.get_TextMatrix(i, 2)) & "' ,  Descriptione = N'" & Apostrophe(FG.get_TextMatrix(i, 3)) & "'  Where Rpt_ID = '" & FG.get_TextMatrix(i, 1) & "'")


            'CNN.Execute("INSERT INTO Ap_Rpt_BLS_Old( Rpt_ID,  Description , Descriptione  ,Chart_of_Accounts_Codes , Grp , Grp_Nme ) " & _
            '     "Values('" & FG.get_TextMatrix(i, 1) & "', N'" & Apostrophe(FG.get_TextMatrix(i, 2)) & "','" & Apostrophe(FG.get_TextMatrix(i, 3)) & "',N'" & Apostrophe(FG.get_TextMatrix(i, 4)) & "' ,N'" & Apostrophe(FG.get_TextMatrix(i, 5)) & "' ,N'" & Apostrophe(FG.get_TextMatrix(i, 6)) & "')")
        Next i

    End Sub
End Class