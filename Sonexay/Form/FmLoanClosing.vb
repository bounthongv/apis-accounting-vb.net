Public Class FmLoanClosing
    Dim MDTab As Integer
    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        CNN.Execute("insert into Ap_LoanClosing(Bnk_Ac_Code , Date_Work ,  Open_Amt , Paid_Amt , Paid_Inte , Rem_Amt , Last_Action , Action_Type , Las_Udate , Statuss  ) " & _
"Values('" & Bnk_Ac_Code.Text & "' , '" & Date_Work.Text & "' , '" & Open_Amt.Text & "' , '" & Paid_Amt.Text & "' , '" & Inte_Amt.Text & "' , '" & Rem_Amt.Text & "' , '" & Last_Action.Text & "' , '" & Action_Type.Text & "' , '" & Las_Udate.Text & "' , '" & Statuss.Text & "' )")
        MessageBox.Show("ok")
    End Sub

    Private Sub BtnEdit2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnEdit2.Click
        CNN.Execute(" update Ap_LoanClosing set " & _
      "Date_Work='" & Date_Work.Text & "' " & _
     ", Open_Amt='" & Open_Amt.Text & "' " & _
    ", Paid_Amt='" & Paid_Amt.Text & "' " & _
    ", Paid_Inte='" & Inte_Amt.Text & "' " & _
    ", Rem_Amt='" & Rem_Amt.Text & "' " & _
    ", Last_Action='" & Last_Action.Text & "' " & _
    ", Action_Type='" & Action_Type.Text & "' " & _
    ", Las_Udate='" & Las_Udate.Text & "' " & _
     ", Statuss='" & Statuss.Text & "' " & _
      " where Bnk_Ac_Code='" & Bnk_Ac_Code.Text & "' ")
        MessageBox.Show("ok")
    End Sub

    Private Sub BtnAddNew2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew2.Click
        Date_Work.Text = ""
        Open_Amt.Text = ""
        Paid_Amt.Text = ""
        Inte_Amt.Text = ""
        Rem_Amt.Text = ""
        Last_Action.Text = ""
        Action_Type.Text = ""
        Las_Udate.Text = ""
        Statuss.Text = ""
        Bnk_Ac_Code.Text = ""
    End Sub


    Private Sub BtnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew.Click

    End Sub

    Private Sub FmLoanClosing_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim colRemovedTabs As New Collection()
        'Dim TabPage1 As TabPage
        'TabPage1 = FmOPenForm.TabControl1.TabPages(MDTab)
        'FmOPenForm.TabControl1.Controls.Remove(TabPage1)
        'MDTabIndex = MDTabIndex - 1
        'FmMain.ToolStripMenuItem55.Enabled = True
        'If MDTabIndex = 0 Then
        '    FmOPenForm.Close()
        'End If
    End Sub

    Private Sub FmLoanClosing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        BtnEdit2.Enabled = False
        BtnEdit.Enabled = False
        BtnDelete.Enabled = False
    End Sub


    Private Sub BtnExit2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit2.Click
        Close()
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Close()
    End Sub
End Class