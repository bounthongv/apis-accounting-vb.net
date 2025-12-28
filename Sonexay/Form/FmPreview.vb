Public Class FmPreview

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        'Call MdiCNum()
   
        Me.Close()

    End Sub

    Private Sub ReportViewer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReportViewer.Load

    End Sub

    Private Sub FmPreview_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If FmNme = "FmNsewJeneralJournal" Then
            FmNsewJeneralJournal_Adjust.ShowDialog()
        End If
        If FmNme = "FmNsewJeneralJournal" Then
            FmNsewJeneralJournal_Adjust.ShowDialog()
        End If
    End Sub

    Private Sub FmPreview_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'SetControlText(Me)
        'Me.ReportViewer.RefreshReport()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        FmCalcu.Show()
    End Sub
End Class