Public Class FmShow
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Static cntd As Integer = 1
        cntd -= 1
        If cntd > 0 Then
        Else
            ' ===============
            FmMain.PictureBox3.Image = FmLogin.PictureBox3.Image
            Fm_Image.Img_ID.Text = FmLogin.txtUserId.Text
            Fm_Image.ImgType.Text = "Back"
            Call LoadPhoto()
            FmMain.BackgroundImage = Fm_Image.PictureBox1.Image
            '=======
            'FmMain.PictureBox3.Image = FmLogin.PictureBox3.Image
            'Fm_Image.Img_ID.Text = "a"
            'Fm_Image.ImgType.Text = "LOGO"
            'Call LoadPhoto()
            'FmMain.PictureBox4.Image = Fm_Image.PictureBox1.Image
            FmMain.VS.Text = FmLogin.VS.Text
            '=================
            cntd = 10
            Timer1.Enabled = False
            FmMain.WindowState = FormWindowState.Maximized
            Me.Visible = False
            Dim Rs As New ADODB.Recordset
            With Rs
                If .State = ConnectionState.Open Then .Close()
                .Open("SELECT Img_Id , ImgType FROM Ap_Image WHERE Img_Id='a' and ImgType='User' ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                If .RecordCount = 0 Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            End With
            Dim FrmPreview As New FmPreview : FrmClosing()
            Dim Rpt As New CryNewsJerneralJournal
            Rpt.SetDataSource(Rs)
            FrmPreview.ReportViewer.ReportSource = Rpt
            FrmPreview.ReportViewer.DisplayGroupTree = False
            FrmPreview.WindowState = FormWindowState.Maximized
            FrmPreview.Focus()
            FmMain.Show()
        End If
    End Sub

    Private Sub FmShow_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Enabled = True
    End Sub
End Class