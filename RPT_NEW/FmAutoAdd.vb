Public Class FmAutoAdd

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ''Dim array() As String = New System.IO.StreamReader(My.Application.Info.DirectoryPath & "\Des.txt").ReadToEnd().Split(vbCrLf)
        'Dim array() As String = New System.IO.StreamReader(TextBox1.Text)
        'If array IsNot Nothing Then
        '    ComboBox1.Items.Clear()
        '    For Each element As String In array
        '        ComboBox1.Items.Add(element)
        '    Next
        'End If

        'For i = 0 To ComboBox1.Items.Count - 1
        '    ComboBox1.SelectedIndex = ComboBox1.SelectedIndex + 1
        '    FmRptProItem.AC_Code.Text = ComboBox1.Text


        '    Dim OP_Amt, Amt, Rem_Amt, Last_Amt As String
        '    OP_Amt = 0
        '    Amt = 0
        '    Rem_Amt = 0
        '    Last_Amt = 0
        '    If FmRptProItem.COP.Checked = True Then
        '        OP_Amt = 1
        '    End If
        '    If FmRptProItem.CAmt.Checked = True Then
        '        Amt = 1
        '    End If
        '    SqlClient = "delete So_Rpt_Proitems where AcCode like '" & FmRptProItem.AC_Code.Text & "%' And RptID = '" & FmRptProItem.RPT_ID.Text & "' And RptType = '" & Apostrophe(MUTY) & "' " & _
        '              " insert into So_Rpt_Proitems (RptID , MainAcCode, AcCode , Des, RptType , RptStatus , SelOpen , SelAmt , CurrType ) " & _
        '              " select '" & FmRptProItem.RPT_ID.Text & "' , '" & FmRptProItem.AC_Code.Text & "' ,  Ac_Code , Name_L , '" & Apostrophe(MUTY) & "' , '" & Apostrophe(FmRptProItem.Rpt_Type.Text) & "' ,  " & OP_Amt & " , " & Amt & " , '" & FmRptProItem.ComCurr.SelectedIndex & "'  " & _
        '              " from Acc_Code where Ac_Code like '" & FmRptProItem.AC_Code.Text & "%'  "
        '    CnnEdit()
        '    'FmRptPro.LoadDGItems()
        'Next i
    End Sub
    Private Sub LoadCombo()

 
    End Sub

End Class