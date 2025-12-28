Public Class FmRptProItem

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Close()
    End Sub

    Private Sub FmRptProItem_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = "FmRptProItem(" & MUTY & ")"
        ComCurr.SelectedIndex = 0
        LoadDG()
    End Sub
    Private Sub LoadDG()
        Dim ds As New DataSet
        Try
            ConnectCL()
            SqlClient = "Select   RptID, Grp, Des, DesE, FML from So_Rpt_Pro  Where RptType = '" & MUTY & "'  Order by RptId"
            LoadCN()
            da.Fill(ds, " So_Rpt_Pro")
            DG.DataSource = ds.Tables(" So_Rpt_Pro")
            cn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try
        DG.Columns(0).HeaderText = "ລະຫັດ " : DG.Columns(0).Width = "55"
        DG.Columns(1).HeaderText = "ໜວດ " : DG.Columns(1).Width = "50"
        DG.Columns(2).HeaderText = "ເນື້ອໃນພາສາລາວ " : DG.Columns(2).Width = "350"
        DG.Columns(3).HeaderText = "ເນື້ອໃນພາສາອັງກິ " : DG.Columns(3).Width = "350"
        DG.Columns(4).HeaderText = "ເລກບັນຊີ " : DG.Columns(4).Width = "100"
    End Sub

    Private Sub DgItems_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DgItems.MouseClick
        Label4.Text = DgItems.CurrentRow.Index
        If (DgItems.Item(0, DgItems.CurrentRow.Index).Value().ToString) <> "" Then
            RPT_ID.Text = (DgItems.Item(0, DgItems.CurrentRow.Index).Value().ToString)
            AC_Code.Text = (DgItems.Item(1, DgItems.CurrentRow.Index).Value().ToString)
            Rpt_Type.Text = (DgItems.Item(3, DgItems.CurrentRow.Index).Value().ToString)
            ComCurr.SelectedIndex = (DgItems.Item(4, DgItems.CurrentRow.Index).Value().ToString)
            COP.Checked = (DgItems.Item(5, DgItems.CurrentRow.Index).Value().ToString)
            CAmt.Checked = (DgItems.Item(6, DgItems.CurrentRow.Index).Value().ToString)
        Else
            ComCurr.SelectedIndex = 0
        End If

      
        'Mouse_Click()
    End Sub
    Private Sub LoadDGItems()
        DgItems.DataSource = ""
        Dim ds As New DataSet
        Try
            ConnectCL()
            SqlClient = "Select  RptID, AcCode, Des, RptStatus,  CurrType, SelOpen, SelAmt from So_Rpt_Proitems where RptID = '" & RPT_ID.Text & "'  And  RptType = '" & MUTY & "' Order by AcCode "
            LoadCN()
            da.Fill(ds, " So_Rpt_Proitems")
            DgItems.DataSource = ds.Tables(" So_Rpt_Proitems")
            cn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try
        DgItems.Columns(0).HeaderText = "ລະຫັດ " : DgItems.Columns(0).Width = "60"
        DgItems.Columns(1).HeaderText = "ເລກບັນຊີ " : DgItems.Columns(1).Width = "75"
        DgItems.Columns(2).HeaderText = "ເນື້ອໃນພາສາລາວ " : DgItems.Columns(2).Width = "300"
        DgItems.Columns(3).HeaderText = "ສະຖານນະ " : DgItems.Columns(3).Width = "75"
        DgItems.Columns(4).HeaderText = "ເງິນ " : DgItems.Columns(4).Width = "60"
        DgItems.Columns(5).HeaderText = "ຍອດຍົກ " : DgItems.Columns(5).Width = "75"
        DgItems.Columns(6).HeaderText = "ເຄື່ອນໄຫວ" : DgItems.Columns(6).Width = "75"
    End Sub
    Private Sub DG_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DG.CellContentClick

    End Sub
    Private Sub Mouse_Click()
        Invalidate()
        BtnSearch.Visible = False
        BtnMove.Visible = False
        'If DgItems.CurrentCell.ColumnIndex = 1 Then
        '    'If DgItems.CurrentRow.Index() <> DgItems.RowCount - 1 Then
        '    '    If DgItems.Item(0, DgItems.CurrentRow.Index).Value() <> "" Then
        '    '        BtnSearch.Visible = False
        '    '    Else
        '    BtnSearch.Visible = True
        '    'End If
        'Else
        '    BtnSearch.Visible = True
        '    'End If
        'End If
        'If DG.CurrentCell.ColumnIndex = 0 Then
        '    If DG.CurrentRow.Index() <> DG.RowCount - 1 Then
        '        If DG.Item(1, DG.CurrentRow.Index).Value() <> "" Then
        '            BtnSearch.Visible = False
        '        Else
        '            BtnSearch.Visible = True
        '        End If
        '    Else
        '        BtnSearch.Visible = True
        '    End If
        'End If
        If DgItems.Item(1, DgItems.CurrentRow.Index).Value().ToString <> "" Then
            BtnMove.Visible = True
        End If
    End Sub
    Private Sub Form1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        Dim rect As Rectangle = DG.GetCellDisplayRectangle(DgItems.CurrentCellAddress.X(), DgItems.CurrentCellAddress.Y(), False)
        Dim x1 As Integer = DgItems.Left
        Dim y1 As Integer = DgItems.Top
        BtnSearch.Location = New Point(GroupBox2.Location.X + rect.Right - 33, GroupBox2.Location.Y + rect.Bottom - 22)
        BtnMove.Location = New Point(5, GroupBox2.Location.Y + rect.Bottom - 22)
    End Sub
    Private Sub DG_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DG.MouseClick
        RPT_ID.Text = (DG.Item(0, DG.CurrentRow.Index).Value().ToString)
        LoadDGItems()
    End Sub

    Private Sub AC_Code_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles AC_Code.KeyPress
        If e.KeyChar = Chr(13) Then
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
            SqlClient = "delete So_Rpt_Proitems where AcCode like '" & AC_Code.Text & "%' And RptID = '" & RPT_ID.Text & "' And RptType = '" & Apostrophe(MUTY) & "' " & _
                      " insert into So_Rpt_Proitems (RptID , MainAcCode, AcCode , Des, RptType , RptStatus , SelOpen , SelAmt , CurrType ) " & _
                      " select '" & RPT_ID.Text & "' , '" & AC_Code.Text & "' ,  Ac_Code , Name_L , '" & Apostrophe(MUTY) & "' , '" & Apostrophe(Rpt_Type.Text) & "' ,  " & OP_Amt & " , " & Amt & " , '" & ComCurr.SelectedIndex & "'  " & _
                      " from Acc_Code where Ac_Code like '" & AC_Code.Text & "%'  "
            CnnEdit()
            LoadDGItems()
        End If
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        SqlClient = "delete So_Rpt_Proitems where  RptID = '" & RPT_ID.Text & "' And RptType = '" & Apostrophe(MUTY) & "' And AcCode like '" & AC_Code.Text & "%'  "
        CnnEdit()
        LoadDGItems()
    End Sub

    Private Sub DgItems_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgItems.CellContentClick
        'Label4.Text = DG.InvalidateRow
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        SqlClient = " Update  So_Rpt_Proitems lck = 0  "
        LoadCN()

        'Dim ds As New DataSet
        Try
            ConnectCL()
            SqlClient = "select *  from SELECT * FROM  So_Rpt_Proitems  Order by Ac_Code  where RptType = '" & MUTY & "' Order by  Rptid ,cnt  "
            LoadCN()
            da.Fill(ds, "gen_jn")
            For i = 0 To ds.Tables(0).Rows.Count - 1
                Dim ss As String = ds.Tables("Caculate_Start").Rows(i).Item("cnt").ToString

            Next i
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try








        CNN.Execute("delete Ap_Rpt_Item")
        With RSC
            Call LoadSqlData("SELECT * FROM  Ap_Rpt_BLS_Item  Order by Ac_Code  ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    CNN.Execute("delete Ap_Rpt_Item where Ac_Code like '" & Trim(CStr(.Fields("Ac_Code").Value.ToString)) & "%' And Rpt_ID = '" & Trim(CStr(.Fields("Rpt_ID").Value.ToString)) & "' And Rpt_Type = '" & Trim(CStr(.Fields("Rpt_Type").Value.ToString)) & "' " & _
                                " insert into Ap_Rpt_Item (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type) select '" & Trim(CStr(.Fields("Rpt_ID").Value.ToString)) & "' ,  Ac_Code , Name_L , '" & Trim(CStr(.Fields("Rpt_Type").Value.ToString)) & "' from Acc_Code where Ac_Code like '" & Trim(CStr(.Fields("Ac_Code").Value.ToString)) & "%'  ")
                    .MoveNext()
                End While
            Else
            End If
        End With
        CNN.Execute("delete Ap_Rpt_BLS_Item")
        CNN.Execute(" insert into Ap_Rpt_BLS_Item  (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type) select Rpt_ID , Ac_Code , Ac_Name, Rpt_Type from Ap_Rpt_Item")
        MsgBox("Ok")

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        FrmAccSearch.ShowDialog()
        If MuAcCode <> "" Then
            AC_Code.Text = MuAcCode
        End If
    End Sub

    Private Sub BtnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnEdit.Click
        SqlClient = "Update So_Rpt_Pro set Lck = 0 where  RptType = '" & Apostrophe(MUTY) & "'  "
        CnnEdit()
        For i = 0 To DG.RowCount - 2
            Dim ds As New DataSet
            Try
                ConnectCL()
                SqlClient = "select top 1 RptId  from So_Rpt_Pro where RptId = '" & (DG.Item(0, i).Value().ToString) & "' And RptType = '" & MUTY & "'   "
                LoadCN()
                da.Fill(ds, "So_Rpt_Pro")
                If ds.Tables(0).Rows.Count > 0 Then
                    If (DG.Item(0, i).Value().ToString) <> "" Then
                        SqlClient = " Update So_Rpt_Pro set  Lck = 1 , Des = N'" & (DG.Item(2, i).Value().ToString) & "' , DesE = '" & (DG.Item(3, i).Value().ToString) & "'  , FML = '" & (DG.Item(4, i).Value().ToString) & "'  where RptId = '" & (DG.Item(0, i).Value().ToString) & "' And RptType = '" & MUTY & "' "

                        CnnEdit()
                    End If
                Else
                    If (DG.Item(0, i).Value().ToString) <> "" Then
                        SqlClient = "Insert Into So_Rpt_Pro (RptID, Grp, RptType, Des, DesE, OpenAmt, Amt, AmtA, FML, StrCal, Fnb, Und, Cor, Lck ) values " & _
                        "'" & (DG.Item(0, i).Value().ToString) & "', '" & (DG.Item(1, i).Value().ToString) & "', '" & MUTY & "',N '" & (DG.Item(3, i).Value().ToString) & "', '" & (DG.Item(4, i).Value().ToString) & "', 0, 0, 0, '" & (DG.Item(5, i).Value().ToString) & "', '', 0, 0, 0, 1 "
                        CnnEdit()
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        Next i
        SqlClient = "Delete So_Rpt_ProDetail from So_Rpt_ProDetail As a ,So_Rpt_Pro As b  where a.RptID = b.RptID and b.lck=0 and b.RptType = '" & Apostrophe(MUTY) & "'  "
        CnnEdit()
        SqlClient = "Delete So_Rpt_Pro where Lck = 0 And  RptType = '" & Apostrophe(MUTY) & "'  "
        CnnEdit()
    End Sub

    Private Sub AC_Code_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AC_Code.TextChanged

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'FmAutoAdd.ShowDialog()
        Dim RSC1 As New ADODB.Recordset
        With RSC1
            Call LoadSqlData("select *  from Table_1  Order by  Ac ", RSC1)
            If .RecordCount > 0 Then
                While Not .EOF()
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
                    AC_Code.Text = (RSC1.Fields("Ac").Value.ToString)
                    SqlClient = "delete So_Rpt_Proitems where AcCode like '" & AC_Code.Text & "%' And RptID = '" & RPT_ID.Text & "' And RptType = '" & Apostrophe(MUTY) & "'  insert into So_Rpt_Proitems (RptID , MainAcCode, AcCode , Des, RptType , RptStatus , SelOpen , SelAmt , CurrType ) select '" & RPT_ID.Text & "' , '" & AC_Code.Text & "' ,  Ac_Code , Name_L , '" & Apostrophe(MUTY) & "' , '" & Apostrophe(Rpt_Type.Text) & "' ,  " & OP_Amt & " , " & Amt & " , '" & ComCurr.SelectedIndex & "'  from Acc_Code where Ac_Code like '" & AC_Code.Text & "%'  "
                    CnnEdit()
                    LoadDGItems()
                    .MoveNext()
                End While
            End If
        End With
        'LoadCombo()
    End Sub
    Private Sub LoadCombo()
        Dim array() As String = New System.IO.StreamReader(My.Application.Info.DirectoryPath & "\Des.txt").ReadToEnd().Split(vbCrLf)
        If array IsNot Nothing Then
            'ComboBox1.Items.Clear()
            For Each element As String In array
                'ComboBox1.Items.Add(element)
            Next
        End If

        'For i = 0 To ComboBox1.Items.Count - 1
        '    ComboBox1.SelectedIndex = i
        '    AC_Code.Text = ComboBox1.Text


        '    Dim OP_Amt, Amt, Rem_Amt, Last_Amt As String
        '    OP_Amt = 0
        '    Amt = 0
        '    Rem_Amt = 0
        '    Last_Amt = 0
        '    If COP.Checked = True Then
        '        OP_Amt = 1
        '    End If
        '    If CAmt.Checked = True Then
        '        Amt = 1
        '    End If
        '    SqlClient = "delete So_Rpt_Proitems where AcCode like '" & AC_Code.Text & "%' And RptID = '" & RPT_ID.Text & "' And RptType = '" & Apostrophe(MUTY) & "'  insert into So_Rpt_Proitems (RptID , MainAcCode, AcCode , Des, RptType , RptStatus , SelOpen , SelAmt , CurrType ) select '" & RPT_ID.Text & "' , '" & AC_Code.Text & "' ,  Ac_Code , Name_L , '" & Apostrophe(MUTY) & "' , '" & Apostrophe(Rpt_Type.Text) & "' ,  " & OP_Amt & " , " & Amt & " , '" & ComCurr.SelectedIndex & "'  from Acc_Code where Ac_Code like '" & AC_Code.Text & "%'  "
        '    'MsgBox(AC_Code.Text)
        '    CnnEdit()
        '    LoadDGItems()
        'Next i
    End Sub

  
    Private Sub BtnMove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnMove.Click
    End Sub
End Class