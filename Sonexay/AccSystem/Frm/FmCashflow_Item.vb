Public Class FmCashflow_Item1
    Public Sub LoadListFG()
        FG.Rows = 1
        With RSC
            FG.FormatString = "^ລ/ດ |< ລະຫັດ    |<ເນື້ອໃນ (ພາສາລາວ)                       |<ເນື້ອໃນ (ພາສາອັງກິດ)                       |<                                       |<|< "
            Call LoadSqlData("SELECT * FROM  Ap_Rpt_Cashflow  order by Rpt_ID ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("Rpt_ID").Value.ToString)) & _
                                    "" & vbTab & (CStr(.Fields("Description").Value.ToString)) & _
                                       "" & vbTab & (CStr(.Fields("Descriptione").Value.ToString)) & _
                                           "" & vbTab & (CStr(.Fields("Chart_of_Accounts_Codes").Value.ToString)) & _
                                               "" & vbTab & (CStr(.Fields("Grp").Value.ToString)) & _
                                        "" & vbTab & (CStr(.Fields("Grp_Nme").Value.ToString)))
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

        If FG2.get_TextMatrix(FG2.Row, 6) = 1 Then
            COP.Checked = True
        Else
            COP.Checked = False
        End If
        If FG2.get_TextMatrix(FG2.Row, 7) = 1 Then
            CLa.Checked = True
        Else
            CLa.Checked = False
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
        FG2.Rows = 1
        With RSC
            FG2.FormatString = "^ລ/ດ |< ລະຫັດ |<ລະຫັດບັນຊີ|<ຊື່ບັນຊີ(ພາສາລາວ)                   |<|<Status|<Open|<Last|<Amt|<Rem"
            Call LoadSqlData("SELECT * FROM  Ap_Rpt_Cashflow_Item where Rpt_ID=   '" & TextBox1.Text & "' Order by Ac_Code ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG2.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("Rpt_ID").Value.ToString)) & _
                    "" & vbTab & Trim(CStr(.Fields("Ac_Code").Value.ToString)) & _
                       "" & vbTab & Trim(CStr(.Fields("Ac_Name").Value.ToString)) & _
                            "" & vbTab & Trim(CStr(.Fields("Ac_NameE").Value.ToString)) & _
                                "" & vbTab & Trim(CStr(.Fields("Rpt_Type").Value.ToString)) & _
                                    "" & vbTab & Trim(CStr(.Fields("Select_Open_Amt").Value.ToString)) & _
                                        "" & vbTab & Trim(CStr(.Fields("Select_Last_Amt").Value.ToString)) & _
                                            "" & vbTab & Trim(CStr(.Fields("Select_Amt").Value.ToString)) & _
                    "" & vbTab & Trim(CStr(.Fields("Select_Rem_Amt").Value.ToString)))
                    .MoveNext()
                End While
            Else
            End If
        End With
        FG2.Rows = CDbl(FG2.Rows) + 1
    End Sub

    Private Sub FG2_AfterEdit(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_AfterEditEvent) Handles FG2.AfterEdit
        Button2.Enabled = True
    End Sub

    Private Sub FG2_AfterScroll(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_AfterScrollEvent) Handles FG2.AfterScroll
        BtnSearch.Visible = False
        BtnMove.Visible = False
    End Sub

    Private Sub FG2_MouseDownEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseDownEvent) Handles FG2.MouseDownEvent
        MouseDownEvent()
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
        LoadListFG()
        FG.AllowUserResizing = VSFlex8U.AllowUserResizeSettings.flexResizeBoth
    End Sub

    Private Sub BtnMove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnMove.Click
        CNN.Execute("delete Ap_Rpt_Cashflow_Item where Ac_Code like '" & AC_Code.Text & "%' And Rpt_ID = '" & RPT_ID.Text & "' And Rpt_Type = '" & Rpt_Type.Text & "' ")
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
        For i = 1 To FG.Rows - 1
            If FG.get_TextMatrix(i, 1) = "" And FG.get_TextMatrix(i, 2) = "" Then
                Exit Sub
            End If
            CNN.Execute("INSERT INTO Ap_Rpt_Cashflow( Rpt_ID,  Description , Descriptione  ,Chart_of_Accounts_Codes , Grp , Grp_Nme ) " & _
                                    "Values('" & FG.get_TextMatrix(i, 1) & "', N'" & Apostrophe(FG.get_TextMatrix(i, 2)) & "','" & Apostrophe(FG.get_TextMatrix(i, 3)) & "',N'" & Apostrophe(FG.get_TextMatrix(i, 4)) & "' ,N'" & Apostrophe(FG.get_TextMatrix(i, 5)) & "' ,N'" & Apostrophe(FG.get_TextMatrix(i, 6)) & "')")
        Next i
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        MsgBox("ການບັນຶກສຳເລັດຜົນ")
        CNN.Execute("delete Ap_Rpt_Cashflow_Item where Rpt_ID = '" & TextBox1.Text & "' ")
        Dim i As Integer
        For i = 1 To FG2.Rows - 1

            If FG2.get_TextMatrix(i, 1) = "" And FG2.get_TextMatrix(i, 2) = "" Then
                Exit Sub
            End If
            CNN.Execute("INSERT INTO Ap_Rpt_Cashflow_Item( Rpt_ID,  Ac_Code , Ac_Name , Ac_NameE ,Amt_Dr , Amt_Cr , BLS  , Rpt_Type) " & _
                 "Values('" & FG2.get_TextMatrix(i, 1) & "', N'" & FG2.get_TextMatrix(i, 2) & "', N'" & FG2.get_TextMatrix(i, 3) & "','" & FG2.get_TextMatrix(i, 4) & "','" & CDbl(0) & "','" & CDbl(0) & "','" & "ALL" & "' ,'" & FG2.get_TextMatrix(i, 5) & "')")
        Next i
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        FG.RemoveItem()
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
            TextBox1.Text = FG.get_TextMatrix(FG.Row, 1)
            Call loadBankItem()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        CNN.Execute("delete Ap_Rpt_Cashflow_Item where  Rpt_ID = '" & RPT_ID.Text & "'   ")
        TextBox1.Text = FG.get_TextMatrix(FG.Row, 1)
        Call loadBankItem()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        FG.Row = 18
        FG.Col = 2
        FG.FocusRect = VSFlex8U.FocusRectSettings.flexFocusInset
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

    End Sub

    Private Sub AC_Code_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AC_Code.TextChanged

    End Sub
 

    Private Sub FG2_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG2.SelChange

    End Sub
End Class