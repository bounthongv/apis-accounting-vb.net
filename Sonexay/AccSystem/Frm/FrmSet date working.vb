Public Class Set_date_working
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        MWorkSetting = MWorkSetting
        Me.Close()
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        MWorkSetting = (MCWorkSetting.SelectionRange.Start)
        Me.Close()
    End Sub
    Private Sub Set_date_working_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MWorkSetting = MWorkSetting
    End Sub
    Private Sub MCWorkSetting_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles MCWorkSetting.DateChanged
        MWorkSetting = (MCWorkSetting.SelectionRange.Start)
    End Sub
End Class