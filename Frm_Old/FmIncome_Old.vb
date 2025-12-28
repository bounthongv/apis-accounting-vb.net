Public Class FmIncome_Old


    Public Sub LoadListFG()
        FG.Rows = 1
        With RSC
            'Call LoadData("select *  from Open_jn WHERE ac_code<>''  " & Sql & "order by ac_code", RSC)
            FG.FormatString = "^ລ/ດ |< ລະຫັດ    |<ເນື້ອໃນ (ພາສາລາວ)                      |<ເນື້ອໃນ (ພາສາອັງກິດ)     |<         |<|< "

            Call LoadSqlData("SELECT * FROM  Ap_Rpt_Income_Old order by  CNT ASC  ", RSC)

            If .RecordCount > 0 Then
                While Not .EOF
                    FG.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("Rpt_ID").Value.ToString)) & _
                                    "" & vbTab & (CStr(.Fields("Description").Value.ToString)) & _
                                       "" & vbTab & Trim(CStr(.Fields("Descriptione").Value.ToString)) & _
                                          "" & vbTab & Trim(CStr(.Fields("Chart_of_Accounts_Codes").Value.ToString)) & _
                                             "" & vbTab & Trim(CStr(.Fields("Grp").Value.ToString)) & _
                                        "" & vbTab & Trim(CStr(.Fields("Grp_Nme").Value.ToString)))
                    .MoveNext()
                End While
            Else
            End If
        End With
        FG.Rows = CDbl(FG.Rows) + 1
    End Sub





    Public Sub MouseDownEvent()
        AC_Code.Text = FG2.get_TextMatrix(FG2.Row, 2)
        Rpt_Type.Text = FG2.get_TextMatrix(FG2.Row, 5)
        TXTCNT.Text = FG2.get_TextMatrix(FG2.Row, 6)
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





    Private Sub loadBankItem()
        FG2.Rows = 1
        With RSC
            'Call LoadData("select *  from Open_jn WHERE ac_code<>''  " & Sql & "order by ac_code", RSC)
            FG2.FormatString = "^ລ/ດ |< ລະຫັດ   |<ລະຫັດບັນຊີ    |<ຊື່ບັນຊີ(ພາສາລາວ)                               |<ຊື່ບັນຊີ(ພາສາອັງກິດ)                  |<ສະຖານະພາບ|<CNT   "
            Call LoadSqlData("SELECT * FROM  Ap_Rpt_Income_Item_Old where Rpt_ID=N'" & TextBox1.Text & "' Order by Ac_Code ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG2.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("Rpt_ID").Value.ToString)) & _
                    "" & vbTab & Trim(CStr(.Fields("Ac_Code").Value.ToString)) & _
                       "" & vbTab & Trim(CStr(.Fields("Ac_Name").Value.ToString)) & _
                                 "" & vbTab & Trim(CStr(.Fields("Ac_NameE").Value.ToString)) & _
                                          "" & vbTab & Trim(CStr(.Fields("Rpt_Type").Value.ToString)) & _
                    "" & vbTab & Trim(CStr(.Fields("CNT").Value.ToString)))
                    .MoveNext()
                End While
            Else
            End If
        End With
        FG2.Rows = CDbl(FG2.Rows) + 1
    End Sub

    Private Sub FG2_AfterCollapse(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_AfterCollapseEvent) Handles FG2.AfterCollapse

    End Sub

    Private Sub FG2_AfterEdit(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_AfterEditEvent) Handles FG2.AfterEdit
        Button2.Enabled = True
    End Sub

    Private Sub FG2_AfterMoveColumn(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_AfterMoveColumnEvent) Handles FG2.AfterMoveColumn

    End Sub

    Private Sub FG2_AfterMoveRow(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_AfterMoveRowEvent) Handles FG2.AfterMoveRow

    End Sub

    Private Sub FG2_AfterRowColChange(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_AfterRowColChangeEvent) Handles FG2.AfterRowColChange

    End Sub

    Private Sub FG2_AfterScroll(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_AfterScrollEvent) Handles FG2.AfterScroll
        BtnSearch.Visible = False
        BtnMove.Visible = False
    End Sub

    Private Sub FG2_AfterSelChange(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_AfterSelChangeEvent) Handles FG2.AfterSelChange

    End Sub

    Private Sub FG2_AfterSort(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_AfterSortEvent) Handles FG2.AfterSort
        'BtnSearch.Visible = False
        'BtnMove.Visible = False
    End Sub

    Private Sub FG2_AfterUserFreeze(ByVal sender As Object, ByVal e As System.EventArgs) Handles FG2.AfterUserFreeze

    End Sub

    Private Sub FG2_AfterUserResize(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_AfterUserResizeEvent) Handles FG2.AfterUserResize

    End Sub

    Private Sub FG2_MouseDownEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseDownEvent) Handles FG2.MouseDownEvent
        MouseDownEvent()
    End Sub

    Private Sub FG2_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG2.SelChange

    End Sub

    Private Sub BtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch.Click
        fmShartOfAccDetail.txtSty.Text = "FmInCome"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
        FG.Focus()
        CNN.Execute("delete Ap_Rpt_Income_Item_Old where Rpt_ID =N'" & TextBox1.Text & "' ")
        Dim i As Integer
        For i = 1 To FG2.Rows - 1

            If FG2.get_TextMatrix(i, 1) = "" And FG2.get_TextMatrix(i, 2) = "" Then
                Exit Sub
            End If
            CNN.Execute("INSERT INTO Ap_Rpt_Income_Item_Old( Rpt_ID,  Ac_Code , Ac_Name , Ac_NameE   , BLS , Rpt_Type ) " & _
                 "Values('" & FG2.get_TextMatrix(i, 1) & "', N'" & FG2.get_TextMatrix(i, 2) & "', N'" & FG2.get_TextMatrix(i, 3) & "','" & FG2.get_TextMatrix(i, 4) & "','" & "ALL" & "' , '" & FG2.get_TextMatrix(i, 5) & "')")
        Next i
    End Sub

    Private Sub FmBankReportId_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button2.Enabled = False
        BtnMove.Visible = False
        BtnSearch.Visible = False
        FG.Size = New System.Drawing.Size(519, 378)
        LoadListFG()
        FG.AutoResize = True
    End Sub

    Private Sub BtnMove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnMove.Click
        CNN.Execute("delete Ap_Rpt_Income_Item_Old where Ac_Code like '" & AC_Code.Text & "%' And Rpt_ID =N'" & RPT_ID.Text & "' And Rpt_Type =N'" & Rpt_Type.Text & "' and CNT=N'" & TXTCNT.Text & "' ")
        BtnMove.Visible = False
        Call loadBankItem()
    End Sub

    Private Sub FG_MouseDownEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseDownEvent) Handles FG.MouseDownEvent
        Select Case MouseButtons
            Case Windows.Forms.MouseButtons.Right
            Case Windows.Forms.MouseButtons.Left
                If FG.Row = FG.Rows - 1 Then
                    Button1.Visible = False
                Else
                    Button1.Visible = True
                End If
                Button1.Top = CInt((FG.CellTop / 15) + FG.Top)
        End Select
    End Sub

    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelChange
        TextBox1.Text = FG.get_TextMatrix(FG.Row, 1)
        RPT_ID.Text = FG.get_TextMatrix(FG.Row, 1)
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
        'CNN.Execute("delete Ap_Rpt_Income")
        Dim i As Integer
        For i = 1 To FG.Rows - 1
            If FG.get_TextMatrix(i, 1) = "" And FG.get_TextMatrix(i, 2) = "" Then
                Exit Sub
            End If
            CNN.Execute("Update Ap_Rpt_Income_Old Set  Description = N'" & Apostrophe(FG.get_TextMatrix(i, 2)) & "' ,  Descriptione = N'" & Apostrophe(FG.get_TextMatrix(i, 3)) & "' , Chart_of_Accounts_Codes = N'" & FG.get_TextMatrix(i, 4) & "'  Where Rpt_ID = '" & FG.get_TextMatrix(i, 1) & "'")
            'CNN.Execute("INSERT INTO Ap_Rpt_Income( Rpt_ID,  Description , Descriptione    , Chart_of_Accounts_Codes , Grp , Grp_Nme ) " & _
            '"Values('" & FG.get_TextMatrix(i, 1) & "', N'" & Apostrophe(FG.get_TextMatrix(i, 2)) & "','" & Apostrophe(FG.get_TextMatrix(i, 3)) & "',N'" & Apostrophe(FG.get_TextMatrix(i, 4)) & "' ,N'" & Apostrophe(FG.get_TextMatrix(i, 5)) & "' ,N'" & Apostrophe(FG.get_TextMatrix(i, 6)) & "')")
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
        FG.RemoveItem()
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


            LoadSqlData("Select top 1 Rpt_ID , Ac_Code from Ap_Rpt_Income_Item_Old where  Ac_Code like '" & AC_Code.Text & "%'  And Rpt_ID <> '" & RPT_ID.Text & "'  ", RSC)
            If RSC.RecordCount <> 0 Then
                MsgBox("ເລກບັນຊີ " & Trim(CStr(RSC.Fields("Ac_Code").Value.ToString)) & " ມີຢູ່ " & Trim(CStr(RSC.Fields("Rpt_ID").Value.ToString)) & " ແລ້ວ")
                Exit Sub
            End If






            CNN.Execute("delete Ap_Rpt_Income_Item_Old where Ac_Code like '" & AC_Code.Text & "%' And Rpt_ID = '" & RPT_ID.Text & "' And Rpt_Type = '" & Rpt_Type.Text & "'  insert into Ap_Rpt_Income_Item_Old (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type) select '" & RPT_ID.Text & "' ,  Ac_Code , Name_L , '" & Rpt_Type.Text & "' from Acc_Code where Ac_Code like '" & AC_Code.Text & "%' and acc_type=N'ບັນຊີແມ່ (P)' ")
            TextBox1.Text = FG.get_TextMatrix(FG.Row, 1)
            Call loadBankItem()
        End If

        'If e.KeyChar = Chr(13) Then
        '    CNN.Execute("delete Ap_Rpt_Income_Item where Ac_Code like '" & AC_Code.Text & "%' And Rpt_ID = '" & RPT_ID.Text & "' An Rpt_Type = '" & Rpt_Type.Text & "'  insert into Ap_Rpt_Income_Item (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type) select '" & RPT_ID.Text & "' ,  Ac_Code , Name_L from Acc_Code where Ac_Code like '" & AC_Code.Text & "%' , '" & Rpt_Type.Text & "' ", )
        '    TextBox1.Text = FG.get_TextMatrix(FG.Row, 1)
        '    Call loadBankItem()
        'End If
    End Sub

    Private Sub AC_Code_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AC_Code.TextChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        CNN.Execute("delete Ap_Rpt_Income_Item_Old where  Rpt_ID=N'" & RPT_ID.Text & "'   ")
        TextBox1.Text = FG.get_TextMatrix(FG.Row, 1)
        Call loadBankItem()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        CNN.Execute("delete Ap_Rpt_Item")
        With RSC
            Call LoadSqlData("SELECT * FROM  Ap_Rpt_Income_Item_Old  Order by Ac_Code  ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    CNN.Execute("delete Ap_Rpt_Item where Ac_Code like '" & Trim(CStr(.Fields("Ac_Code").Value.ToString)) & "%' And Rpt_ID = '" & Trim(CStr(.Fields("Rpt_ID").Value.ToString)) & "' And Rpt_Type = '" & Trim(CStr(.Fields("Rpt_Type").Value.ToString)) & "' " & _
                                " insert into Ap_Rpt_Item (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type) select '" & Trim(CStr(.Fields("Rpt_ID").Value.ToString)) & "' ,  Ac_Code , Name_L , '" & Trim(CStr(.Fields("Rpt_Type").Value.ToString)) & "' from Acc_Code where Ac_Code like '" & Trim(CStr(.Fields("Ac_Code").Value.ToString)) & "%'  ")
                    .MoveNext()
                End While
            Else
            End If
        End With
        CNN.Execute("delete Ap_Rpt_Income_Item_Old")
        CNN.Execute(" insert into Ap_Rpt_Income_Item_Old  (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type) select Rpt_ID , Ac_Code , Ac_Name, Rpt_Type from Ap_Rpt_Item")
        MsgBox("Ok")
    End Sub
End Class