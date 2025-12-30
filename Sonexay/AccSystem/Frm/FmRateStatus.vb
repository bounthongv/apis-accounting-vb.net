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
        FG.Rows.Clear()
        With RSC
            Call LoadSqlData("SELECT * FROM  Ap_RateStatus where Curr <>''" & sql & "  order by in_date asc", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.Rows.Add(.AbsolutePosition, _
                                Trim(CStr(.Fields("In_Date").Value)), _
                                Trim(CStr(.Fields("Curr").Value)), _
                                .Fields("Curr_Name").Value, _
                                Trim(Format(CDbl(.Fields("Rate").Value), "##,##0.00")))
                    .MoveNext()
                End While
                If FG.Rows.Count > 0 Then
                    Label1.Text = "ອັດຕາແລກປ່ຽນໃນວັນທີ : " & Format(CDate(FG.Rows(0).Cells(1).Value), "dd/MM/yyyy") & " ຫາວັນທີ " & Format(CDate(FG.Rows(FG.Rows.Count - 1).Cells(1).Value), "dd/MM/yyyy")
                End If
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
                        FG.Rows.Add(.AbsolutePosition, _
                                    Trim(CStr(.Fields("In_Date").Value)), _
                                    Trim(CStr(.Fields("Curr").Value)), _
                                    .Fields("Curr_Name").Value, _
                                    Trim(Format(CDbl(.Fields("Rate").Value), "##,##0.00")))
                        .MoveNext()
                    End While
                    If FG.Rows.Count > 0 Then
                        Label1.Text = "ອັດຕາແລກປ່ຽນໃນວັນທີ : " & Format(DT, "dd/MM/yyyy") & " ຫາວັນທີ " & Format(CDate(FG.Rows(FG.Rows.Count - 1).Cells(1).Value), "dd/MM/yyyy")
                    End If
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
        SetupGrid()
    End Sub

    Private Sub SetupGrid()
        FG.Columns.Clear()
        FG.Columns.Add("No", "ລ/ດ")
        FG.Columns.Add("Date", "ວັນທີ")
        FG.Columns.Add("Currency", "ສະກຸນເງິນ")
        FG.Columns.Add("CurrencyName", "ສະກຸນເງິນ (ຊື່ເຕັມ)")
        FG.Columns.Add("Rate", "ອັດຕາແລກປ່ຽນ")

        ' Formatting based on FormatString = "^ລ/ດ  |ວັນທີ               |^ ສະກຸນເງິນ        | ສະກຸນເງິນ (ຊື່ເຕັມ)                                     <| ອັດຕາແລກປ່ຽນ                          "
        
        FG.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        FG.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        FG.Columns(0).Width = 50

        FG.Columns(1).Width = 100

        FG.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        FG.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        FG.Columns(2).Width = 80

        FG.Columns(3).Width = 250

        FG.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        FG.Columns(4).Width = 150

        FG.AllowUserToAddRows = False
        FG.ReadOnly = True
        FG.SelectionMode = DataGridViewSelectionMode.FullRowSelect
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