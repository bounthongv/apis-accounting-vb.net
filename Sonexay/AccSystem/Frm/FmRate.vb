Public Class FmRate

    Private Sub FmRate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LockFormSiz()
        FG.FormatString = "^ດ/ລ |ສະກຸນເງິນ (ຊື່ຫຍໍ້)       |ສະກຸນເງິນ (ຊື່ເຕັມ)       |ອັດຕາຊື້         |ອັດຕາຂາຍ      |ຊື່ລົງທ້າຍ   |"
        FG2.FormatString = "^ດ/ລ |^ປະເພດເງິນ       "
        BtnSave.Enabled = True
        BtnDelete.Enabled = False
        BtnEdit2.Enabled = False
        Curr.Enabled = True
        LoadListFG()
        FG.Size = New System.Drawing.Size(745, 423)
        FG2.Size = New System.Drawing.Size(180, 423)
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
        For i = 1 To FG2.Rows - 1
            'MessageBox.Show(CDbl(FG2.get_TextMatrix(i, 1)))
            CNN.Execute("insert into Ap_MoneyPaper(Curr,Paper)values('" & Curr.Text & "','" & CDbl(FG2.get_TextMatrix(i, 1)) & "')")
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

    Private Sub FG_MouseUpEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseUpEvent) Handles FG.MouseUpEvent
        If FG.Row And FG.Col > 0 Then
            If FG.get_TextMatrix(1, 1) <> "" Then
                BtnSave.Enabled = False
                BtnDelete.Enabled = True
                BtnEdit2.Enabled = True
                Curr.Enabled = False

                Curr.Text = FG.get_TextMatrix(FG.Row, 1)
                CurrName.Text = FG.get_TextMatrix(FG.Row, 2)
                Rate.Text = FG.get_TextMatrix(FG.Row, 3)
                Rate2.Text = FG.get_TextMatrix(FG.Row, 4)
                Curr_Last.Text = FG.get_TextMatrix(FG.Row, 5)
                LoadFG2()
                TextBox1.Text = FG2.Rows - 1
            End If
        End If
    End Sub

    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelChange


    End Sub
    Private Sub LoadFG2()
        FG2.Rows = 1
        With RSC
            Call LoadSqlData("SELECT * FROM  Ap_MoneyPaper where curr='" & Curr.Text & "' ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG2.AddItem(.AbsolutePosition & vbTab & Trim(Format(CDbl(.Fields("Paper").Value), "##,##0.00")))
                    .MoveNext()
                End While
            Else
                FG2.Rows = 2
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
        FG.Rows = 1
        With RSC
            Call LoadSqlData("SELECT * FROM  Ap_RateSeting ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("Curr").Value)) & _
                                 "" & vbTab & Trim(CStr(.Fields("Curr_Name").Value)) & _
                                 "" & vbTab & Trim(Format(CDbl(.Fields("Rate").Value), "##,##0.00")) & _
                                    "" & vbTab & Trim(Format(CDbl(.Fields("Rate2").Value), "##,##0.00")) & _
                                  "" & vbTab & Trim(CStr(.Fields("Curr_Last").Value)) & _
                                 "" & vbTab & ((.Fields("cnt").Value)))
                    .MoveNext()
                End While
            Else
                FG.Rows = 16
            End If
        End With
    End Sub
    Private Sub BtnEdit2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnEdit2.Click
        CNN.Execute("delete FROM Ap_RateSeting where Curr='" & Curr.Text & "'  ")
        CNN.Execute("insert into Ap_RateSeting(Curr,Curr_Name,Rate ,Rate2,Curr_Last)values('" & Curr.Text & "',N'" & CurrName.Text & "','" & CDbl(Rate.Text) & "','" & CDbl(Rate2.Text) & "',N'" & Curr_Last.Text & "')")
        Call LoadListFG()
        CNN.Execute("delete FROM Ap_RateSeting  ")
        For i = 1 To FG.Rows - 1
            CNN.Execute("insert into Ap_RateSeting(Curr,Curr_Name,Rate ,Rate2,Curr_Last)values('" & FG.get_TextMatrix(i, 1) & "',N'" & FG.get_TextMatrix(i, 2) & "'," & CDbl(FG.get_TextMatrix(i, 3)) & "," & CDbl(FG.get_TextMatrix(i, 4)) & ",N'" & FG.get_TextMatrix(i, 5) & "')")
            CNN.Execute("insert into Ap_RateStatus(in_date , Curr , Curr_Name , Rate ,  Last_User)values('" & Format(MWorkSetting, "MM-dd-yyyy") & "','" & FG.get_TextMatrix(i, 1) & "',N'" & FG.get_TextMatrix(i, 2) & "','" & CDbl(FG.get_TextMatrix(i, 3)) & "','" & MUserName & "')")
        Next
        BtnSave.Enabled = True
        BtnDelete.Enabled = False
        BtnEdit2.Enabled = False
        Curr.Enabled = True

        For i = 1 To FG2.Rows - 1
            FG2.Row = i
            FG2.Col = 1
            If FG2.get_TextMatrix(i, 1) = "" Then MessageBox.Show("ໃສ່ໃບເງິນບໍ່ຄບຖ້ວນກະລຸນນາໃສ່ໃຫມ່") : Exit Sub
            If FG2.get_TextMatrix(i, 1) <= 0 Then MessageBox.Show("ໃສ່ໃບເງິນບໍ່ຄບຖ້ວນກະລຸນນາໃສ່ໃຫມ່") : Exit Sub
        Next
        'CNN.Execute("delete FROM Ap_MoneyPaper where Curr='" & Curr.Text & "'  ")
        'For i = 1 To FG2.Rows - 1
        '    'MessageBox.Show(CDbl(FG2.get_TextMatrix(i, 1)))
        '    CNN.Execute("insert into Ap_MoneyPaper(Curr,Paper)values('" & Curr.Text & "','" & CDbl(FG2.get_TextMatrix(i, 1)) & "')")
        'Next
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
            If CDbl(TextBox1.Text) < 2 Then Exit Sub
            If CDbl(TextBox1.Text) > 20 Then TextBox1.Text = 20 : FG2.Rows = 20 : Exit Sub
            FG2.Rows = CDbl(TextBox1.Text) + 1
            Curr_Last.Focus()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If IsNumeric(TextBox1.Text) = False Then TextBox1.Text = "0" : Exit Sub
        If TextBox1.Text = "" Then TextBox1.Text = "0" : Exit Sub
    End Sub

    Private Sub FG2_AfterEdit(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_AfterEditEvent) Handles FG2.AfterEdit
        FG2.set_TextMatrix(FG2.Row, 1, Format(CDbl(FG2.get_TextMatrix(FG2.Row, 1)), "##,##0.00"))
        FG2.Row = FG2.Row + 1
    End Sub

  
    Private Sub FG2_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG2.SelChange

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