Public Class FmRate

    Private Sub FmRate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LockFormSiz()
        SetupGrids()
        
        BtnSave.Enabled = True
        BtnDelete.Enabled = False
        BtnEdit2.Enabled = False
        Curr.Enabled = True
        LoadListFG()
        'FG.Size = New System.Drawing.Size(745, 423) ' Size is set in designer
        'FG2.Size = New System.Drawing.Size(180, 423)
    End Sub

    Private Sub SetupGrids()
        ' FG Setup
        ' FG.FormatString = "^ດ/ລ |ສະກຸນເງິນ (ຊື່ຫຍໍ້)       |ສະກຸນເງິນ (ຊື່ເຕັມ)       |ອັດຕາຊື້         |ອັດຕາຂາຍ      |ຊື່ລົງທ້າຍ   |"
        FG.Columns.Clear()
        FG.Columns.Add("No", "ດ/ລ")
        FG.Columns.Add("Currency", "ສະກຸນເງິນ (ຊື່ຫຍໍ້)")
        FG.Columns.Add("CurrencyFull", "ສະກຸນເງິນ (ຊື່ເຕັມ)")
        FG.Columns.Add("BuyRate", "ອັດຕາຊື້")
        FG.Columns.Add("SellRate", "ອັດຕາຂາຍ")
        FG.Columns.Add("Suffix", "ຊື່ລົງທ້າຍ")
        FG.Columns.Add("Cnt", "cnt") ' Hidden column for ID if needed, based on LoadListFG it loads cnt at index 6

        FG.Columns(0).Width = 50
        FG.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        FG.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        FG.Columns(1).Width = 120
        FG.Columns(2).Width = 200
        FG.Columns(3).Width = 100
        FG.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        FG.Columns(4).Width = 100
        FG.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        FG.Columns(5).Width = 100
        FG.Columns(6).Visible = False

        FG.AllowUserToAddRows = False
        FG.ReadOnly = True
        FG.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG.MultiSelect = False


        ' FG2 Setup
        ' FG2.FormatString = "^ດ/ລ |^ປະເພດເງິນ       "
        FG2.Columns.Clear()
        FG2.Columns.Add("No", "ດ/ລ")
        FG2.Columns.Add("MoneyType", "ປະເພດເງິນ")

        FG2.Columns(0).Width = 50
        FG2.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        FG2.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        FG2.Columns(1).Width = 100
        FG2.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        FG2.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        FG2.AllowUserToAddRows = False
        FG2.ReadOnly = False ' Allow editing for MoneyType? Original code had AfterEdit event.
        ' However, original code Logic in BtnEdit2 iterates FG2 and reads values. 
        ' Also "FG2_AfterEdit" handles formatting.
        ' Let's assume it needs to be editable.
        FG2.Columns(0).ReadOnly = True
    End Sub

    Private Sub LockFormSiz()
        MaximizeBox = False
        ControlBox = False
        FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
    End Sub
    Public Sub Lng()
        LoadLng()
        'SetControlText(Me)
    End Sub

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        If Curr.Text = "" Then MessageBox.Show("ກະລຸນນາໃສ່ສະກຸນເງິນກ່ອນ") : Exit Sub : Curr.Focus() : End
        If CDbl(Rate.Text) = 0 Then MessageBox.Show("ກະລຸນນາໃສ່ອັດຕາແລກປ່ຽນກ່ອນເງິນກ່ອນ") : Rate.Focus() : Exit Sub : End

        Call LoadSqlData("SELECT  Curr FROM Ap_RateSeting WHERE Curr = '" & Trim(Curr.Text) & "'", RSC)
        If RSC.RecordCount > 0 Then
            MsgBox("  ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
            Curr.Focus()
            If RSC.State = ConnectionState.Open Then RSC.Close()
            Exit Sub
        End If
        CNN.Execute("insert into Ap_RateSeting(Curr,Curr_Name,Rate ,Rate2,Curr_Last)values('" & Curr.Text & "',N'" & CurrName.Text & "','" & CDbl(Rate.Text) & "','" & CDbl(Rate2.Text) & "',N'" & Curr_Last.Text & "')")

        SaveDateStatus()

        BtnSave.Enabled = True
        BtnDelete.Enabled = False
        BtnEdit2.Enabled = False
        Curr.Enabled = True

        CNN.Execute("delete FROM Ap_MoneyPaper where Curr='" & Curr.Text & "'  ")
        For i = 0 To FG2.Rows.Count - 1
             'MessageBox.Show(CDbl(FG2.get_TextMatrix(i, 1)))
             If FG2.Rows(i).Cells(1).Value IsNot Nothing AndAlso IsNumeric(FG2.Rows(i).Cells(1).Value) Then
                CNN.Execute("insert into Ap_MoneyPaper(Curr,Paper)values('" & Curr.Text & "','" & CDbl(FG2.Rows(i).Cells(1).Value) & "')")
             End If
        Next
        LoadListFG()
    End Sub
    Private Sub SaveDateStatus()
        CNN.Execute("insert into Ap_RateStatus(in_date , Curr , Curr_Name , Rate ,  Last_User)values('" & Format(MWorkSetting, "MM-dd-yyyy") & "','" & Curr.Text & "',N'" & CurrName.Text & "','" & CDbl(Rate.Text) & "',N'" & MUserName & "')")
        Dim srNum As New ADODB.Recordset
        Dim mNum As Integer
        Call LoadSqlData("SELECT top 1 cnt FROM  Ap_RateStatus Order by cnt DESC", srNum)
        If srNum.RecordCount = 0 Then
        Else
            mNum = CDbl(Val(srNum.Fields("cnt").Value)) - 30000
            CNN.Execute("delete FROM Ap_RateStatus WHERE cnt <= '" & mNum & "' ")
        End If
    End Sub

    Private Sub BtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelete.Click
        CNN.Execute("delete FROM Ap_RateSeting WHERE Curr = '" & Trim(Curr.Text) & "' ")
        BtnSave.Enabled = True
        BtnDelete.Enabled = False
        BtnEdit2.Enabled = False
        Curr.Enabled = True
        Call LoadListFG()
    End Sub

    Private Sub FG_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles FG.SelectionChanged, FG.Click
        If FG.CurrentRow Is Nothing Then Exit Sub
        If FG.CurrentRow.Index < 0 Then Exit Sub

        Try
            If FG.CurrentRow.Cells(1).Value IsNot Nothing AndAlso FG.CurrentRow.Cells(1).Value.ToString() <> "" Then
                BtnSave.Enabled = False
                BtnDelete.Enabled = True
                BtnEdit2.Enabled = True
                Curr.Enabled = False

                Curr.Text = FG.CurrentRow.Cells(1).Value.ToString()
                CurrName.Text = FG.CurrentRow.Cells(2).Value.ToString()
                Rate.Text = Format(CDbl(FG.CurrentRow.Cells(3).Value), "0.00")
                Rate2.Text = Format(CDbl(FG.CurrentRow.Cells(4).Value), "0.00")
                Curr_Last.Text = FG.CurrentRow.Cells(5).Value.ToString()
                LoadFG2()
                TextBox1.Text = FG2.Rows.Count.ToString()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' Handled by SelectionChanged
    End Sub

    Private Sub LoadFG2()
        FG2.Rows.Clear()
        With RSC
            Call LoadSqlData("SELECT * FROM  Ap_MoneyPaper where curr='" & Curr.Text & "' ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG2.Rows.Add(.AbsolutePosition, _
                                 Trim(Format(CDbl(.Fields("Paper").Value), "##,##0.00")))
                    .MoveNext()
                End While
            Else
                'FG2.Rows = 2 ' Original logic added empty rows? 
                ' Let's just keep it clear or add one if needed? 
                ' The TextBox1 controls row count.
            End If
        End With
    End Sub

    Private Sub FmRate_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Call MdiCNum()
        BtnSave.Enabled = True
        BtnDelete.Enabled = False
        BtnEdit2.Enabled = False
        If MdSearchDataList = "FmNsewJeneralJournal" Then
            FmNsewJeneralJournal_Adjust.LoadSetRate()
            FmNsewJeneralJournal_Adjust.LoadCurr()
        End If
    End Sub



    Private Sub BtnAddNew2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew2.Click
        BtnSave.Enabled = True
        BtnDelete.Enabled = False
        BtnEdit2.Enabled = False
        Curr.Enabled = True
        Curr.Text = ""
        CurrName.Text = ""
        Rate.Text = ""
    End Sub
    Private Sub LoadListFG()
        FG.Rows.Clear()
        With RSC
            Call LoadSqlData("SELECT * FROM  Ap_RateSeting ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.Rows.Add(.AbsolutePosition, _
                                Trim(CStr(.Fields("Curr").Value)), _
                                Trim(CStr(.Fields("Curr_Name").Value)), _
                                Trim(Format(CDbl(.Fields("Rate").Value), "##,##0.00")), _
                                Trim(Format(CDbl(.Fields("Rate2").Value), "##,##0.00")), _
                                Trim(CStr(.Fields("Curr_Last").Value)), _
                                .Fields("cnt").Value)
                    .MoveNext()
                End While
            Else
                ' FG.Rows = 16 ' Don't need to force rows in DataGridView
            End If
        End With
    End Sub
    Private Sub BtnEdit2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnEdit2.Click
        CNN.Execute("delete FROM Ap_RateSeting where Curr='" & Curr.Text & "'  ")
        CNN.Execute("insert into Ap_RateSeting(Curr,Curr_Name,Rate ,Rate2,Curr_Last)values('" & Curr.Text & "',N'" & CurrName.Text & "','" & CDbl(Rate.Text) & "','" & CDbl(Rate2.Text) & "',N'" & Curr_Last.Text & "')")
        Call LoadListFG()
        
        ' This block re-inserts everything? Seems dangerous or redundant logic in original code.
        ' Original logic:
        ' CNN.Execute("delete FROM Ap_RateSeting  ")
        ' For i = 1 To FG.Rows - 1 ...
        ' It seems to be trying to update ALL records based on grid content? 
        ' Or maybe just updating status history?
        ' The original code loops through FG (which was just reloaded via LoadListFG)
        ' Let's try to preserve the logic but use DGV.
        
        CNN.Execute("delete FROM Ap_RateSeting")
        For i = 0 To FG.Rows.Count - 1
             Dim r As DataGridViewRow = FG.Rows(i)
             ' Columns: 1=Curr, 2=CurrName, 3=Rate, 4=Rate2, 5=Curr_Last
             CNN.Execute("insert into Ap_RateSeting(Curr,Curr_Name,Rate ,Rate2,Curr_Last)values('" & r.Cells(1).Value.ToString() & "',N'" & r.Cells(2).Value.ToString() & "'," & CDbl(r.Cells(3).Value) & "," & CDbl(r.Cells(4).Value) & ",N'" & r.Cells(5).Value.ToString() & "')")
             CNN.Execute("insert into Ap_RateStatus(in_date , Curr , Curr_Name , Rate ,  Last_User)values('" & Format(MWorkSetting, "MM-dd-yyyy") & "','" & r.Cells(1).Value.ToString() & "',N'" & r.Cells(2).Value.ToString() & "','" & CDbl(r.Cells(3).Value) & "','" & MUserName & "')")
        Next

        BtnSave.Enabled = True
        BtnDelete.Enabled = False
        BtnEdit2.Enabled = False
        Curr.Enabled = True

        For i = 0 To FG2.Rows.Count - 1
            ' FG2.Row = i ' Not needed
            ' FG2.Col = 1 ' Not needed
            Dim val As Object = FG2.Rows(i).Cells(1).Value
            If val Is Nothing OrElse val.ToString() = "" Then 
                ' Skip empty? Original said messagebox and exit sub.
                ' But original loop started at 1?
                ' MessageBox.Show("ໃສ່ໃບເງິນບໍ່ຄບຖ້ວນກະລຸນນາໃສ່ໃຫມ່") : Exit Sub
            ElseIf IsNumeric(val) AndAlso CDbl(val) <= 0 Then
                 MessageBox.Show("ໃສ່ໃບເງິນບໍ່ຄບຖ້ວນກະລຸນນາໃສ່ໃຫມ່") : Exit Sub
            End If
        Next
        
        ' Original code had commented out block for Ap_MoneyPaper update here.
        
        SaveDateStatus()
        Call LoadListFG()
    End Sub

    Private Sub Rate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Rate.KeyPress
        If e.KeyChar = Chr(13) Then
            Rate2.Focus()
        End If
    End Sub
    Private Sub Rate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Rate.TextChanged
        If IsNumeric(Rate.Text) = False Then Rate.Text = "0.00" : Exit Sub
        If Rate.Text = "" Then Rate.Text = "0.00" : Exit Sub
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            If Not IsNumeric(TextBox1.Text) Then Exit Sub
            Dim count As Integer = CInt(TextBox1.Text)
            If count < 2 Then Exit Sub
            If count > 20 Then 
                count = 20
                TextBox1.Text = "20"
            End If
            
            FG2.RowCount = count
            ' Re-number rows?
            For i As Integer = 0 To FG2.RowCount - 1
                FG2.Rows(i).Cells(0).Value = i + 1
            Next
            
            Curr_Last.Focus()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If IsNumeric(TextBox1.Text) = False Then TextBox1.Text = "0" : Exit Sub
        If TextBox1.Text = "" Then TextBox1.Text = "0" : Exit Sub
    End Sub

    ' Replaces FG2_AfterEdit
    Private Sub FG2_CellEndEdit(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FG2.CellEndEdit
        If e.ColumnIndex = 1 Then
             Dim val As Object = FG2.Rows(e.RowIndex).Cells(1).Value
             If val IsNot Nothing AndAlso IsNumeric(val) Then
                 FG2.Rows(e.RowIndex).Cells(1).Value = Format(CDbl(val), "##,##0.00")
             End If
        End If
    End Sub

  
    Private Sub FG2_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub BtnExit2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click

    End Sub

    Private Sub Curr_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Curr.KeyPress
        If e.KeyChar = Chr(13) Then
            CurrName.Focus()
        End If
    End Sub

    Private Sub Curr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Curr.TextChanged

    End Sub

    Private Sub CurrName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles CurrName.KeyPress
        If e.KeyChar = Chr(13) Then
            Rate.Focus()
        End If
    End Sub

    Private Sub CurrName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CurrName.TextChanged

    End Sub

    Private Sub Rate2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Rate2.KeyPress
        If e.KeyChar = Chr(13) Then
            TextBox1.Focus()
        End If
    End Sub

    Private Sub Rate2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Rate2.TextChanged

    End Sub

    Private Sub Curr_Last_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Curr_Last.KeyPress

    End Sub

    Private Sub Curr_Last_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Curr_Last.TextChanged

    End Sub

    Private Sub L1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MuLng = "L"
        LoadLng()
        SetControlText(Me)
    End Sub

    Private Sub L2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MuLng = "E"
        LoadLng()
        SetControlText(Me)
    End Sub


    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Close()
    End Sub
End Class