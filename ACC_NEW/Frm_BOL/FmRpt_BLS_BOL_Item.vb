Public Class FmRpt_BLS_BOL_Item

    Private Sub FmBankReportId_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button2.Enabled = False
        BtnMove.Visible = False
        BtnSearch.Visible = False
        
        SetupGrid()
        SetupGrid2()
        
        LoadListFG()
    End Sub

    Private Sub SetupGrid()
        FG.Columns.Clear()
        FG.Columns.Add("No", "ລ/ດ")
        FG.Columns.Add("Rpt_ID", "ລະຫັດ")
        FG.Columns.Add("Description", "ເນື້ອໃນ (ພາສາລາວ)")
        FG.Columns.Add("Descriptione", "ເນື້ອໃນ (ພາສາອັງກິດ)")
        FG.Columns.Add("Chart_of_Accounts_Codes", "")
        FG.Columns.Add("Grp", "")
        FG.Columns.Add("Grp_Nme", "")

        FG.Columns(0).Width = 50
        FG.Columns(1).Width = 80
        FG.Columns(2).Width = 200
        FG.Columns(3).Width = 200
        FG.Columns(4).Width = 100
        FG.Columns(5).Width = 50
        FG.Columns(6).Width = 100

        FG.AllowUserToAddRows = False
        FG.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub SetupGrid2()
        FG2.Columns.Clear()
        FG2.Columns.Add("No", "ລ/ດ")
        FG2.Columns.Add("Rpt_ID", "ລະຫັດ")
        FG2.Columns.Add("Ac_Code", "ລະຫັດບັນຊີ")
        FG2.Columns.Add("Ac_Name", "ຊື່ບັນຊີ(ພາສາລາວ)")
        FG2.Columns.Add("Ac_NameE", "ຊື່ບັນຊີ(ພາສາອັງກິດ)")
        FG2.Columns.Add("Rpt_Type", "ສະຖານນະພາບ")

        FG2.Columns(0).Width = 50
        FG2.Columns(1).Width = 80
        FG2.Columns(2).Width = 100
        FG2.Columns(3).Width = 200
        FG2.Columns(4).Width = 200
        FG2.Columns(5).Width = 100

        FG2.AllowUserToAddRows = False
        FG2.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Public Sub LoadListFG()
        FG.Rows.Clear()
        With RSC
            Call LoadSqlData("SELECT * FROM  Ap_Rpt_BLS_BOL  order by cnt ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.Rows.Add(.AbsolutePosition, _
                                Trim(CStr(.Fields("Rpt_ID").Value.ToString)), _
                                (CStr(.Fields("Description").Value.ToString)), _
                                (CStr(.Fields("Descriptione").Value.ToString)), _
                                (CStr(.Fields("Chart_of_Accounts_Codes").Value.ToString)), _
                                (CStr(.Fields("Grp").Value.ToString)), _
                                (CStr(.Fields("Grp_Nme").Value.ToString)))
                    .MoveNext()
                End While
            End If
        End With
    End Sub

    Private Sub loadBankItem()
        FG2.Rows.Clear()
        With RSC
            Call LoadSqlData("SELECT * FROM  Ap_Rpt_BLS_BOL_Item where Rpt_ID=   '" & TextBox1.Text & "' Order by Ac_Code ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG2.Rows.Add(.AbsolutePosition, _
                                Trim(CStr(.Fields("Rpt_ID").Value.ToString)), _
                                Trim(CStr(.Fields("Ac_Code").Value.ToString)), _
                                Trim(CStr(.Fields("Ac_Name").Value.ToString)), _
                                Trim(CStr(.Fields("Ac_NameE").Value.ToString)), _
                                Trim(CStr(.Fields("Rpt_Type").Value.ToString)))
                    .MoveNext()
                End While
            End If
        End With
    End Sub

    Private Sub FG2_CellEndEdit(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FG2.CellEndEdit
        Button2.Enabled = True
    End Sub

    Private Sub FG2_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles FG2.MouseDown
        If FG2.CurrentRow Is Nothing Then Exit Sub
        
        AC_Code.Text = If(FG2.CurrentRow.Cells(2).Value Is Nothing, "", FG2.CurrentRow.Cells(2).Value.ToString())
        Rpt_Type.Text = If(FG2.CurrentRow.Cells(5).Value Is Nothing, "", FG2.CurrentRow.Cells(5).Value.ToString())

        BtnSearch.Visible = True
        
        If e.Button = Windows.Forms.MouseButtons.Left Then
            If FG2.CurrentCell IsNot Nothing AndAlso FG2.CurrentCell.ColumnIndex = 2 Then
                BtnSearch.Visible = True
            Else
                BtnSearch.Visible = False
            End If
            
            If FG2.CurrentRow.Index = FG2.Rows.Count - 1 Then
                BtnMove.Visible = False
            Else
                BtnMove.Visible = True
            End If

            Dim rect As Rectangle = FG2.GetCellDisplayRectangle(If(FG2.CurrentCell IsNot Nothing, FG2.CurrentCell.ColumnIndex, 0), FG2.CurrentRow.Index, False)
            BtnSearch.Left = FG2.Left + rect.Left + (rect.Width / 2)
            BtnSearch.Top = FG2.Top + rect.Top
            BtnMove.Top = FG2.Top + rect.Top
        End If
    End Sub

    Private Sub FG2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles FG2.SelectionChanged
        ' Handle selection change logic if needed
    End Sub

    Private Sub FG_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles FG.MouseDown
        If FG.CurrentRow Is Nothing Then Exit Sub
        
        If e.Button = Windows.Forms.MouseButtons.Left Then
            If FG.CurrentRow.Index = FG.Rows.Count - 1 Then
                Button1.Visible = False
            Else
                Button1.Visible = True
            End If
            
            Dim colIdx As Integer = If(FG.CurrentCell IsNot Nothing, FG.CurrentCell.ColumnIndex, 0)
            Dim rect As Rectangle = FG.GetCellDisplayRectangle(colIdx, FG.CurrentRow.Index, False)
            Button1.Top = FG.Top + rect.Top
        End If
    End Sub

    Private Sub FG_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles FG.SelectionChanged
        If FG.CurrentRow Is Nothing Then Exit Sub
        
        TextBox1.Text = If(FG.CurrentRow.Cells(1).Value Is Nothing, "", FG.CurrentRow.Cells(1).Value.ToString())
        RPT_ID.Text = If(FG.CurrentRow.Cells(1).Value Is Nothing, "", FG.CurrentRow.Cells(1).Value.ToString())
        Call loadBankItem()

        Button2.Enabled = False
        BtnMove.Visible = False
        BtnSearch.Visible = False
    End Sub

    Private Sub BtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch.Click
        fmShartOfAccDetail.txtSty.Text = "FmBLS"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
        FG.Focus()
    End Sub

    Private Sub BtnMove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnMove.Click
        CNN.Execute("delete Ap_Rpt_BLS_BOL_Item where Ac_Code like '" & AC_Code.Text & "%' And Rpt_ID = '" & RPT_ID.Text & "' And Rpt_Type = '" & Rpt_Type.Text & "' ")
        BtnMove.Visible = False
        Call loadBankItem()
    End Sub

    Private Sub BtnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnEdit.Click
        Button2.Enabled = False
        BtnMove.Visible = False
        BtnSearch.Visible = False
        
        Dim i As Integer
        For i = 0 To FG.Rows.Count - 1
            Dim v1 As String = If(FG.Rows(i).Cells(1).Value Is Nothing, "", FG.Rows(i).Cells(1).Value.ToString())
            Dim v2 As String = If(FG.Rows(i).Cells(2).Value Is Nothing, "", FG.Rows(i).Cells(2).Value.ToString())
            Dim v3 As String = If(FG.Rows(i).Cells(3).Value Is Nothing, "", FG.Rows(i).Cells(3).Value.ToString())
            Dim v4 As String = If(FG.Rows(i).Cells(4).Value Is Nothing, "", FG.Rows(i).Cells(4).Value.ToString())

            If v1 = "" And v2 = "" Then Continue For
            
            CNN.Execute("Update Ap_Rpt_BLS_BOL Set  Description = N'" & Apostrophe(v2) & "' ,  Descriptione = N'" & Apostrophe(v3) & "' , Chart_of_Accounts_Codes = N'" & v4 & "'  Where Rpt_ID = '" & v1 & "'")
        Next i
        MsgBox("ການບັນທຶກສຳເລັດຜົນ")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        MsgBox("ການບັນຶກສຳເລັດຜົນ")
        CNN.Execute("delete Ap_Rpt_BLS_BOL_Item where Rpt_ID = '" & TextBox1.Text & "' ")
        Dim i As Integer
        For i = 0 To FG2.Rows.Count - 1
            Dim v1 As String = If(FG2.Rows(i).Cells(1).Value Is Nothing, "", FG2.Rows(i).Cells(1).Value.ToString())
            Dim v2 As String = If(FG2.Rows(i).Cells(2).Value Is Nothing, "", FG2.Rows(i).Cells(2).Value.ToString())
            Dim v3 As String = If(FG2.Rows(i).Cells(3).Value Is Nothing, "", FG2.Rows(i).Cells(3).Value.ToString())
            Dim v4 As String = If(FG2.Rows(i).Cells(4).Value Is Nothing, "", FG2.Rows(i).Cells(4).Value.ToString())
            Dim v5 As String = If(FG2.Rows(i).Cells(5).Value Is Nothing, "", FG2.Rows(i).Cells(5).Value.ToString())

            If v1 = "" And v2 = "" Then Continue For
            
            CNN.Execute("INSERT INTO Ap_Rpt_BLS_BOL_Item( Rpt_ID,  Ac_Code , Ac_Name , Ac_NameE ,Amt_Dr , Amt_Cr , BLS  , Rpt_Type) " & _
                 "Values('" & v1 & "', N'" & v2 & "', N'" & v3 & "','" & v4 & "','" & CDbl(0) & "','" & CDbl(0) & "','" & "ALL" & "' ,'" & v5 & "')")
        Next i
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If FG.CurrentRow IsNot Nothing Then
            FG.Rows.RemoveAt(FG.CurrentRow.Index)
            Button1.Visible = False
        End If
    End Sub

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Me.Close()
    End Sub

    Private Sub AC_Code_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles AC_Code.KeyPress
        If e.KeyChar = Chr(13) Then
            LoadSqlData("Select top 1 Rpt_ID , Ac_Code from Ap_Rpt_BLS_BOL_Item where  Ac_Code like '" & AC_Code.Text & "%'  And Rpt_ID <> '" & RPT_ID.Text & "'  ", RSC)
            If RSC.RecordCount <> 0 Then
                MsgBox("ເລກບັນຊີ " & Trim(CStr(RSC.Fields("Ac_Code").Value.ToString)) & " ມີຢູ່ " & Trim(CStr(RSC.Fields("Rpt_ID").Value.ToString)) & " ແລ້ວ")
                Exit Sub
            End If

            CNN.Execute("delete Ap_Rpt_BLS_BOL_Item where Ac_Code like '" & AC_Code.Text & "%' And Rpt_ID = '" & RPT_ID.Text & "' And Rpt_Type = '" & Rpt_Type.Text & "'  insert into Ap_Rpt_BLS_BOL_Item (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type) select N'" & RPT_ID.Text & "' ,  Ac_Code , Name_L , '" & Rpt_Type.Text & "' from Acc_Code where Ac_Code like '" & AC_Code.Text & "%'  ")
            
            If FG.CurrentRow IsNot Nothing Then
                TextBox1.Text = If(FG.CurrentRow.Cells(1).Value Is Nothing, "", FG.CurrentRow.Cells(1).Value.ToString())
                Call loadBankItem()
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        CNN.Execute("delete Ap_Rpt_BLS_BOL_Item where  Rpt_ID = '" & RPT_ID.Text & "'   ")
        If FG.CurrentRow IsNot Nothing Then
            TextBox1.Text = If(FG.CurrentRow.Cells(1).Value Is Nothing, "", FG.CurrentRow.Cells(1).Value.ToString())
            Call loadBankItem()
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        CNN.Execute("delete Ap_Rpt_Item")
        With RSC
            Call LoadSqlData("SELECT * FROM  Ap_Rpt_BLS_BOL_Item  Order by Ac_Code  ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    Dim rptId As String = Trim(CStr(.Fields("Rpt_ID").Value.ToString))
                    Dim acCode As String = Trim(CStr(.Fields("Ac_Code").Value.ToString))
                    Dim rptType As String = Trim(CStr(.Fields("Rpt_Type").Value.ToString))
                    
                    CNN.Execute("delete Ap_Rpt_Item where Ac_Code like '" & acCode & "%' And Rpt_ID = '" & rptId & "' And Rpt_Type = '" & rptType & "' " & _
                                " insert into Ap_Rpt_Item (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type) select '" & rptId & "' ,  Ac_Code , Name_L , '" & rptType & "' from Acc_Code where ac_code like '" & acCode & "%'  ")
                    .MoveNext()
                End While
            End If
        End With
        CNN.Execute("delete Ap_Rpt_BLS_BOL_Item")
        CNN.Execute(" insert into Ap_Rpt_BLS_BOL_Item  (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type) select Rpt_ID , Ac_Code , Ac_Name, Rpt_Type from Ap_Rpt_Item")
        MsgBox("Ok")
    End Sub

End Class