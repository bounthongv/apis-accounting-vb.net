Public Class FmRateStatus
    Dim sql As String
    Dim DT As New Date
    Dim mNum As String
    Private Sub BtnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRefresh.Click
        LoadListFG()
    End Sub

    Private Sub LoadListFG()
        sql = ""
      
        sql = " AND Ap_RateStatus.In_Date   BETWEEN '" & Format(DateTimePicker4.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker3.Value, "yyyy-MM-dd") & "' "
        FG.Rows = 1
        With RSC
            Call LoadSqlData("SELECT * FROM  Ap_RateStatus where Curr <>''" & sql & "  order by in_date asc", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("In_Date").Value)) & _
                                 "" & vbTab & Trim(CStr(.Fields("Curr").Value)) & _
                                 "" & vbTab & ((.Fields("Curr_Name").Value)) & _
                                 "" & vbTab & Trim(Format(CDbl(.Fields("Rate").Value), "##,##0.00")))
                    .MoveNext()
                End While
                Label1.Text = "ອັດຕາແລກປ່ຽນໃນວັນທີ : " & Format(CDate(FG.get_TextMatrix(1, 1)), "dd/MM/yyyy") & " ຫາວັນທີ " & Format(CDate(FG.get_TextMatrix(FG.Rows - 1, 1)), "dd/MM/yyyy")
            Else
                '===========

                LoadSqlData("select top 1  In_Date from Ap_RateStatus where In_date<'" & Format(DateTimePicker4.Value, "MM-dd-yyyy") & "' Order by In_Date DESC", RSC)
                With RSC
                    Do Until .EOF = True
                        DT = Format(CDate(.Fields("In_Date").Value), "dd-MM-yyyy")
                        .MoveNext()
                    Loop
                End With
                'DateTimePicker4.Text = DT

                sql = ""
                sql = " AND Ap_RateStatus.In_Date   BETWEEN '" & Format(DT, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker3.Value, "yyyy-MM-dd") & "' "
                If Cmb.Text <> "==ທັງຫມົດ==" Then
                    sql = " AND Curr = '" & Cmb.Text & "' "
                End If
                Call LoadSqlData("SELECT * FROM  Ap_RateStatus where Curr <>''" & sql & "   order by in_date asc", RSC)
                If .RecordCount > 0 Then
                    While Not .EOF
                        FG.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("In_Date").Value)) & _
                                     "" & vbTab & Trim(CStr(.Fields("Curr").Value)) & _
                                     "" & vbTab & ((.Fields("Curr_Name").Value)) & _
                                     "" & vbTab & Trim(Format(CDbl(.Fields("Rate").Value), "##,##0.00")))
                        .MoveNext()
                    End While
                    Label1.Text = "ອັດຕາແລກປ່ຽນໃນວັນທີ : " & Format(DT, "dd/MM/yyyy") & " ຫາວັນທີ " & Format(CDate(FG.get_TextMatrix(FG.Rows - 1, 1)), "dd/MM/yyyy")
                End If


                '=============
            End If
        End With
    End Sub

    Private Sub BtnExit2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Close()
    End Sub

    Private Sub DateTimePicker4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePicker4.KeyPress
        If e.KeyChar = Chr(13) Then
            LoadListFG()
        End If
    End Sub

    Private Sub DateTimePicker4_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker4.ValueChanged
        DateTimePicker3.Text = DateTimePicker4.Text
    End Sub

    Private Sub DateTimePicker3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePicker3.KeyPress
        If e.KeyChar = Chr(13) Then
            LoadListFG()
        End If
    End Sub

    Private Sub DateTimePicker3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker3.ValueChanged

    End Sub

    Private Sub FmRateStatus_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Call MdiCNum()
    End Sub

    Private Sub FmRateStatus_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadCurr()
        FG.FormatString = "^ລ/ດ  |ວັນທີ               |^ ສະກຸນເງິນ        | ສະກຸນເງິນ (ຊື່ເຕັມ)                                     <| ອັດຕາແລກປ່ຽນ                          "
    End Sub
    Private Sub LoadCurr()
        Dim Comm As ADODB.Command
        Dim rsat As New ADODB.Recordset
        Comm = New ADODB.Command
        Comm.ActiveConnection = CNN
        Comm.CommandText = "SELECT Curr FROM Ap_RateSeting WHERE Curr <> '" & "" & " order by Curr'"
        rsat = Comm.Execute
        If rsat.RecordCount <> 0 Then
            While Not rsat.EOF()
                Cmb.Items.Add(Trim(rsat.Fields("Curr").Value))
                rsat.MoveNext()
            End While
            Cmb.Items.Add("==ທັງຫມົດ==")
        End If
    End Sub

 

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Close()
    End Sub
End Class