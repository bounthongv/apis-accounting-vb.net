
Public Class FmMain
    Dim sql As String
    Dim StrFIlePath As String
    Dim StrFilename As String
    Dim rs As New ADODB.Recordset
    Dim ImageSlno As Integer
    Dim ImageSlno2 As Integer
    Dim s As Integer
    Dim SPW As String = ""
    Dim SUSID As String = ""
    Dim con As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source = Connection.mdb;Persist Security Info=False;Jet OLEDB:Database Password=2459428") ' connection srting 
    Private Sub FmMain_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Application.Exit()
    End Sub
    Private Sub FmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing


        'If MessageBox.Show("ທ່ານຕ້ອງການອອກຈາກໂປຼແກຣມບໍ?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
        '    Exit Sub
        '    Me.Close()
        'Else
        '    e.Cancel = True
        'End If
        
    End Sub

    Public Sub LoadHideMenu()
        Call LoadSqlData("  SELECT Ints,cnt  FROM Ap_Section_Item where Ints = '0' order by cnt ", RSC)
        With RSC
            Do Until .EOF = True
                If (.Fields("cnt").Value) = 1 Then
                    MnOff.Visible = False
                ElseIf (.Fields("cnt").Value) = 2 Then
                    MnOffSub.Visible = False
                ElseIf (.Fields("cnt").Value) = 3 Then
                    MnDateSeting.Visible = False
                End If
                .MoveNext()
            Loop
        End With

    End Sub

    Private Sub FmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblpage_total.Text = "0/0"
        'MnSystem.BackgroundImage = Image.FromFile(My.Application.Info.DirectoryPath & "\Menu.png")
        'MnAt.BackgroundImage = Image.FromFile(My.Application.Info.DirectoryPath & "\Menu.png")
        'MnAcReport.BackgroundImage = Image.FromFile(My.Application.Info.DirectoryPath & "\Menu.png")
        'MnFinance.BackgroundImage = Image.FromFile(My.Application.Info.DirectoryPath & "\Menu.png")
        'ToolStripMenuItem2.BackgroundImage = Image.FromFile(My.Application.Info.DirectoryPath & "\Menu.png")
        'ToolStripMenuItem34.BackgroundImage = Image.FromFile(My.Application.Info.DirectoryPath & "\Menu.png")
        'MnHelp.BackgroundImage = Image.FromFile(My.Application.Info.DirectoryPath & "\Menu.png")
        LoadLng()
        MuLngL.Checked = True
        MuLngE.Checked = False
        MuLng = "L"
        SetControlText(Me)
        ChgChildForm()
        MnLaoLang.Checked = True
        MnEngLang.Checked = False

        PictureBox7.BackgroundImage = Image.FromFile(My.Application.Info.DirectoryPath & "\BackGroundLOGO.png")
        MuLng = "L"
        'FG.Cols = 30
        'FG.Size = New System.Drawing.Size(1028, 43)
        'MuLng = "L"
        'LoadLng()
        'SetControlText(Me)
        '=======
        Call Office()
        'MsgBox(RptSjOff)
        'Dim a As String = Environment.MachineName.ToString
        ''Timer1.Start()
        'If a <> MDServerName Then
        '    ToolStripMenuItem57.Enabled = False
        'Else
        '    ToolStripMenuItem57.Enabled = True
        'End If
        FormMain = "ApBank"

        Panel3.Visible = False
        Call Login()

        Timer1.Enabled = True
        '==============
        MWorkSetting = Today
        LoadHideMenu()
        Load_Curr()
        HideMenu()
        LoadLngnnn()
        LoadlangMM()
    End Sub

    Public Sub Load_Curr()
        Dim rs As New ADODB.Recordset
        With rs
            Call LoadSqlData("SELECT * FROM Curr", rs)
            If .RecordCount <> 0 Then
                MDCurr = (.Fields("Curr").Value)
            End If
        End With

    End Sub
    Public Sub Login()
        Dim rsProj As New ADODB.Recordset
        With rsProj
            Call LoadSqlData("select Sec_ID from AP_Users where Usr_id='" & MUserID & "'  ", rsProj)
            If rsProj.RecordCount <> 0 Then
                MSection = (.Fields("Sec_ID").Value)
            End If
        End With

        If MSection = 1 Then
            'MnDailyEntry.Visible = True
            'ToolStripMenuItem16.Visible = False
            'ToolStripMenuItem17.Visible = False
            MnAt.Visible = False

            '====================================================================
            Call LoadSqlData(" SELECT Ints FROM Ap_Section_Item WHERE Sec_ID ='" & MSection & "' AND Sec_Nm=N'" & "ລະຫັດບັນຊີ" & "'", RSC)
            If RSC.RecordCount <> 0 Then
                MD1 = RSC.Fields("Ints").Value
            End If
            If MD1 = 1 Then
                ToolStripMenuItem18.Visible = True
            Else
                ToolStripMenuItem18.Visible = False
            End If

            '=====================================================================
            Call LoadSqlData(" SELECT Ints FROM Ap_Section_Item WHERE Sec_ID ='" & MSection & "' AND Sec_Nm=N'" & "ການແລກປ່ຽນເງິນຕາ" & "'", RSC)
            If RSC.RecordCount <> 0 Then
                MD1 = RSC.Fields("Ints").Value
            End If


            '=====================================================================


        End If
        LoadInfo()
    End Sub

    Private Sub LoadImage()
        Call LoadAcData("select * from TblImage where BackID='" & "00001" & "'", rs)
        If rs.RecordCount = 0 Then
            Exit Sub
        Else
            On Error GoTo hang
hang:
            If Err.Number = 0 Then
                Me.BackgroundImage = Image.FromFile(rs.Fields("FileAddress").Value.ToString)
                VSysError = False
            Else
                VSysError = True
                MessageBox.Show("No Background " & (rs.Fields("FileAddress").Value.ToString) & "  in data base please Select Image To Background")
                OpenFileDialog1.ShowDialog()
                StrFilename = OpenFileDialog1.SafeFileName
                StrFIlePath = OpenFileDialog1.FileName
                If StrFIlePath = "" Or StrFIlePath = "OpenFileDialog1" Then Exit Sub
                conn.Execute("UPDATE TblImage SET FileAddress ='" & StrFIlePath & "' ," & _
                                       " FileNmae='" & StrFIlePath & "' " & _
                                      " WHERE BackID='" & "00001" & "' ")
                Call LoadAcData("select * from TblImage where BackID='" & "00001" & "'", rs)
                If rs.RecordCount <> 0 Then
                    While Not rs.EOF
                        Me.BackgroundImage = Image.FromFile(rs.Fields("FileAddress").Value.ToString)
                        Me.BackgroundImageLayout = ImageLayout.Stretch
                        rs.MoveNext()
                    End While
                End If

            End If
            Me.BackgroundImageLayout = ImageLayout.Stretch
        End If
    End Sub


    Private Sub ToolStripMenuItem43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub




    Private Sub ToolStripMenuItem45_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'FmTrialBalanceReport___10_ColuMne_.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem18.Click

        fmShartOfAcc.GrPage.SendToBack()

        fmShartOfAcc.FG.Size = New System.Drawing.Size(1019, 665)
        IsMdiContainer = True
        Panel4.Visible = False
        fmShartOfAcc.MdiParent = Me
        fmShartOfAcc.WindowState = FormWindowState.Maximized
        fmShartOfAcc.Show()
        fmShartOfAcc.Focus()

    End Sub

    'Private Sub ToolStripMenuItem20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    IsMdiContainer = True
    '    Panel4.Visible = False
    '    FmOpen_jn_List.MdiParent = Me
    '    FmOpen_jn_List.WindowState = FormWindowState.Maximized
    '    FmOpen_jn_List.Show()
    '    FmOpen_jn_List.Focus()
    '    FmOpen_jn_List.FG.Size = New System.Drawing.Size(1007, 565)
    '    FmOpen_jn_List..Text = "1/1/" & Format(CDate(MWorkSetting), "yyyy")
    '    FmOpen_jn_List.dts.Text = FmOpen_jn_List.dtt.Text
    '    FmOpen_jn_List.LoadListFG()
    '    FmOpen_jn_List.FG.Size = New System.Drawing.Size(1007, 500)
    'End Sub

    Private Sub ToolStripMenuItem19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub















    Private Sub ToolStripMenuItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub



    Private Sub ລາຍງານປະຈຳວນເງນກToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'FmReportRepayfinace.ShowDialog()


        'FrmReport.WindowState = FormWindowState.Maximized

    End Sub

    Private Sub ລາຍການສຳນກງານToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnOff.Click

        IsMdiContainer = True
        Panel4.Visible = False
        Office_AP.MdiParent = Me
        Office_AP.WindowState = FormWindowState.Maximized
        Office_AP.Show()
    End Sub

    Private Sub ຂມນສາຂາToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnOffSub.Click

        IsMdiContainer = True
        Panel4.Visible = False
        FmOffice_AP.MdiParent = Me
        FmOffice_AP.WindowState = FormWindowState.Maximized
        FmOffice_AP.Show()

    End Sub





    Private Sub ToolStripMenuItem31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub



    Private Sub ToolStripMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnDateSeting.Click
        Set_date_working.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem37.Click
        Shell("calc.exe", vbNormalFocus)
    End Sub

    Private Sub ToolStripMenuItem38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem38.Click
        Set_date_working.ShowDialog()
    End Sub



    Private Sub ToolStripMenuItem35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem35.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TileHorizontalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TileVerticalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub mnBackGround_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnBackGround.Click
        LoadImgSize()
        Try

            Fm_Image.OpenFileDialog1.Filter = "Bmp Files(*.bmp)|*.bmp|Gif Files(*.gif)|*.gif|Jpg Files(*.jpg)|*.jpg"
            Fm_Image.OpenFileDialog1.ShowDialog()
            Dim s As String
            s = Fm_Image.OpenFileDialog1.FileName
            Dim objImage As System.Drawing.Image = System.Drawing.Image.FromFile(s)
            If objImage.Width > b_x Or objImage.Height > b_y Then MsgBox("ຂະຫນາດຮູບ (" & objImage.Width & " x " & objImage.Height & ") ໃຫ່ຍເກີນຂະຫນາດ (" & b_x & " x " & b_y & ") ") : Fm_Image.PictureBox1.Image = Fm_Image.a123456789.Image : Exit Sub
            'If objImage.Width > 1024 Or objImage.Height > 768 Then MsgBox("ຂະຫນາດຮູບ (" & objImage.Width & " x " & objImage.Height & ") ໃຫ່ຍເກີນຂະຫນາດ (1024 x 768) ") : Exit Sub
            Me.BackgroundImage = Image.FromFile(Fm_Image.OpenFileDialog1.FileName)

            Fm_Image.Img_ID.Text = FmLogin.txtUserId.Text
            Fm_Image.ImgType.Text = "Back"
            deleteImage()
            Insert_Image()



            'TextBox4.Text = Fm_Image.OpenFileDialog1.SafeFileName
            'ImageFrom.Text = Fm_Image.OpenFileDialog1.FileName
        Catch

        End Try

    End Sub





    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        txtOldPass.Text = ""
        txtNewPass.Text = ""
        txtConfrimPass.Text = ""
        Panel3.Visible = False
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'If txtConfrimPass.Text = "" Then MsgBox("ກະລຸນາຢືນຢັນລະຫັດຜ່ານຂອງທ່ານ", MsgBoxStyle.OkOnly) : txtConfrimPass.Focus() : Exit Sub
        If txtConfrimPass.Text <> txtNewPass.Text Then MsgBox("ລະຫັດຜ່ານຂອງທ່ານບໍ່ຖືກຕ້ອງມກະລຸນາປ່ຽນ", MsgBoxStyle.OkOnly) : txtConfrimPass.Text = "" : Exit Sub
        CNN.Execute("UPDATE AP_Users SET " & _
                " PWD=N'" & txtNewPass.Text & "'," & _
                " lst_updt=Getdate()," & _
                " lst_usr=N'" & MDServerUser & "'," & _
                " pc_nm=N'" & MDServerName & "' " & _
                " WHERE Usr_id=N'" & (MUserID) & "' ")
        MsgBox("ບັນທຶກສໍາເລັດຜົນ!", MsgBoxStyle.OkOnly)
        Call Button4_Click(sender, e)
    End Sub

    Private Sub ToolStripMenuItem29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'TTime.Text = "ເວລາ :" & Format(TimeOfDay) & " |"
        Label22.Text = Format(Now, "hh:mm:ss tt")
        If Microsoft.VisualBasic.Right(Label22.Text, 2) = "AM" Then
            Time.Text = "ເວລາ: " & Format(Now, "hh:mm:ss") & " ຕອນເຊົ້າ"
        Else
            Time.Text = "ເວລາ: " & Format(Now, "hh:mm:ss") & " ຕອນແລງ"
        End If

        LoadScExS()
        'Label31.Text = " ທຸກໆລະບົບຈະຖືກປິດການໃຊ້ງານພາຍໃນ  (" & CDbl(Format(MWorkSetting, "ss")) - 60 & ":) ;ວິນາທີ"

        If CloseAll = 1 Then
            'If MPermit = "Admin" Then
            Panel2.Visible = True
            'Else
            'Me.Hide()

            'FmLock.Label1.Text = "ລະບົບນີ້ຖືກລ໋ອກແລ້ວກະລຸນນາຕິດຕໍ່ພວກເຮົາ"
            Timer1.Enabled = False
     
            'ConnectSQL2()
            Exit Sub

            'FmLock.Close()
            'FmLogin.Close()
        Else
            Panel2.Visible = False
            'ConnectSQL()
        End If

    End Sub





    Private Sub ລາຍງານການແລກປຽນເງນຕາປະຈາວນToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'FmRateServiceReport.MdiParent = Me
        'FmRateServiceReport.WindowState = FormWindowState.Maximized
        'FmRateServiceReport.Show()
    End Sub




    Private Sub ToolStripMenuItem49_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem49.Click
        FmCalcu.Show()
        FormOpening2()
        FmCalcu.BringToFront()
    End Sub



    Private Sub ລາງຂມນToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If MessageBox.Show("ທ່ານຕ້ອງການລ້າງຂໍ້ມູນນີ້ ແມ່ນຫຼືບໍ່ ?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            ''CNN.Execute("delete Ap_Calculate")
            ''CNN.Execute("delete Ap_Calculate_Item")
            ''CNN.Execute("delete Ap_Customer")
            ''CNN.Execute("delete Ap_Image where ImgType ='Cust'")
            ''CNN.Execute("delete Ap_Load_Interes")
            'CNN.Execute("delete Ap_Loan")
            ''CNN.Execute("delete Ap_LoanClosing")
            ''CNN.Execute("delete Ap_loanpayment")
            ''CNN.Execute("delete Ap_ReceipItem")
            ''CNN.Execute("delete Ap_Receipt")
            ''CNN.Execute("delete Ap_Save_Money")
            ''CNN.Execute("delete Ap_Save_Money_Return")
            ''CNN.Execute("delete Ap_SaveDeposit")
            ''CNN.Execute("delete Ap_SaveTransfer")
            ''CNN.Execute("delete Ap_SaveWithdr")
            ''CNN.Execute("delete Ap_Saving")
            ''CNN.Execute("delete Ap_Service_History")
            'CNN.Execute("delete gen_jn")
            'CNN.Execute("delete Open_jn")


            'CNN.Execute("delete Ap_ProgramNo ")
            'CNN.Execute("delete Ap_LoanPaymentTeble")
            'CNN.Execute("delete Ap_LoanPaymentTebleAll")
            'CNN.Execute("delete Ap_RementPayLoan")
            'CNN.Execute("delete Ap_Repayfinace")
            'CNN.Execute("delete Ap_ProgramRepay")
            'MsgBox("ການລ້າງຂໍ້ມູນສຳເລັດ")
        Else
            Exit Sub
        End If

    End Sub

    Private Sub ToolStripMenuItem32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MuLng = "L"
        'MuLng = "L"
        'SetControlText(Me)
        'ChgChildForm()
        'FmRate.Lng()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'CNN.Close()

        'VSysError = True
        'RSC.Close()
        'Comm.Cancel()
        MsgBox("ok")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'CNN.Open()
        ConnectSQL()
        'VSysError = False
        'RSC.Close()
        'Comm.Cancel()
        MsgBox("ok")
    End Sub

    Private Sub ToolStripMenuItem50_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripMenuItem17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

 
    Private Sub ToolStripMenuItem52_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem52.Click
        LoadImgSize()
        Try
            'Fm_Image.OpenFileDialog1.Filter = "Bmp Files(*.bmp)|*.bmp|Gif Files(*.gif)|*.gif|Jpg Files(*.jpg)|*.jpg"
            Fm_Image.OpenFileDialog1.ShowDialog()
            Dim s As String
            s = Fm_Image.OpenFileDialog1.FileName
            Dim objImage As System.Drawing.Image = System.Drawing.Image.FromFile(s)
            If objImage.Width > g_x Or objImage.Height > g_y Then MsgBox("ຂະຫນາດຮູບ (" & objImage.Width & " x " & objImage.Height & ") ໃຫ່ຍເກີນຂະຫນາດ (" & g_x & " x " & g_y & ") ") : Fm_Image.PictureBox1.Image = Fm_Image.a123456789.Image : Exit Sub
            'If objImage.Width > 240 Or objImage.Height > 320 Then MsgBox("ຂະຫນາດຮູບ (" & objImage.Width & " x " & objImage.Height & ") ໃຫ່ຍເກີນຂະຫນາດ (240 x 320) ") : Exit Sub

            PictureBox3.Image = Image.FromFile(Fm_Image.OpenFileDialog1.FileName)
            Fm_Image.Img_ID.Text = FmLogin.txtUserId.Text
            Fm_Image.ImgType.Text = "User"
            Update_Image()



            Dim svl As String
            svl = ""
            Dim srNum As New ADODB.Recordset
            Call LoadSqlData("SELECT  ImgType , Img_Id , cnt FROM Ap_Image where ImgType='User' and Img_Id= '" & FmLogin.txtUserId.Text & "' ", srNum)
            svl = Val(srNum.Fields("cnt").Value.ToString)
            If svl = "" Then
                Fm_Image.Img_ID.Text = FmLogin.txtUserId.Text
                Fm_Image.ImgType.Text = "User"
                Insert_Image()
            End If


            'TextBox4.Text = Fm_Image.OpenFileDialog1.SafeFileName
            'ImageFrom.Text = Fm_Image.OpenFileDialog1.FileName

        Catch

        End Try
    End Sub

    Private Sub ToolStripMenuItem53_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem53.Click
        FmAccountBook.ShowDialog()
    End Sub


    Private Sub ToolStripMenuItem54_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub



    Private Sub ToolStripMenuItem31_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnUser.Click

        IsMdiContainer = True
        Panel4.Visible = False
        FrmUser.MdiParent = Me
        FrmUser.WindowState = FormWindowState.Maximized
        FrmUser.Show()
        'IsMdiContainer = True
        'FrmUser_DDC.Visible = False
        'FrmUser_DDC.MdiParent = Me
        ''FrmUser_DDC.WindowState = FormWindowState.Maximized
        'FrmUser_DDC.Show()

    End Sub

    Private Sub ToolStripMenuItem29_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnChangUser.Click
        FmLogOut.ShowDialog()
        Call Login()
    End Sub

    Private Sub ToolStripMenuItem30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnChangPsw.Click
        Call LoadSqlData("select * from AP_Users where Usr_id='" & MUserID & "' ", rs)
        With rs
            If .RecordCount <> 0 Then
                TextBox1.Text = MUserID & " / " & MuSubOff
                txtOldPass.Text = rs.Fields("PWD").Value
            End If
        End With
        Panel3.Visible = True
    End Sub

    Private Sub MenuStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub


    Private Sub Label29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label29.Click
        'Form1.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnImgSize.Click
        FmImgSize.ShowDialog()
    End Sub

  

    Private Sub ToolStripMenuItem57_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem57.Click
        FrmBackup_Data.Show()
    End Sub

    Private Sub ToolStripMenuItem55_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Fmdelete.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem56_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripMenuItem51_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripMenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ສງປດທກລະບບToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ສງປດທກລະບບToolStripMenuItem.Click
        'If CloseAll = 1 Then
        '    FmLock.ShowDialog()
        'End If


        UpdateExS()
    End Sub
  


    Private Sub ToolStripMenuItem33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MuLng = "E"
    End Sub



    Private Sub ໃບສະຫບຊບສມບດToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ໃບສະຫບຊບສມບດToolStripMenuItem.Click
        'FmRpt_BLS.nn.Text = ໃບສະຫບຊບສມບດToolStripMenuItem.Text





        FmRpt_BLS.ShowDialog()
        'FmRpt_BLS_BOL.nn.Text = ໃບສະຫບຊບສມບດToolStripMenuItem.Text

        'FmRpt_BLS_BOL.ShowDialog()

    End Sub


    Private Sub ໃບລາຍງານຜນໄດຮບToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ໃບລາຍງານຜນໄດຮບToolStripMenuItem.Click
        'Axt = 1
        'FmRpt_Income.nn.Text = ໃບລາຍງານຜນໄດຮບToolStripMenuItem.Text

        FmRpt_Income.ShowDialog()
        'FmRpt_Income_BOL.Show()
    End Sub
    Private Sub ToolStripMenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem8.Click
        LngId = "6002" : MsgQtin() : If MsgSL = 1 Then FmLogin.Close()
    End Sub


    Private Sub ໃບດນດຽງບນຊສຳຮອງToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ໃບດນດຽງບນຊສຳຮອງToolStripMenuItem.Click
        ''FmTrialBalanceReport.FG.Size = New System.Drawing.Size(1006, 407)
        'IsMdiContainer = True
        'Panel4.Visible = False
        'FmTrialBalanceReport.MdiParent = Me
        'FmTrialBalanceReport.WindowState = FormWindowState.Maximized
        'FmTrialBalanceReport.Show()
        'FmTrialBalanceReport.Focus()

        FmTrialBalanceReport2022.FG.Size = New System.Drawing.Size(1006, 407)
        IsMdiContainer = True
        Panel4.Visible = False
        FmTrialBalanceReport2022.MdiParent = Me
        FmTrialBalanceReport2022.WindowState = FormWindowState.Maximized
        FmTrialBalanceReport2022.Show()
        FmTrialBalanceReport2022.Focus()
    End Sub

    Private Sub ບນຊແຍກປະເພດToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ບນຊແຍກປະເພດToolStripMenuItem.Click

        FmPostedLedgers.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem7_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem7.Click
        FmOpen_jn_List.FG.Size = New System.Drawing.Size(1260, 542)
        'FmJeneralJournal_List.FG.Size = New System.Drawing.Size(1260, 388)
        IsMdiContainer = True
        Panel4.Visible = False
        FmOpen_jn_List.MdiParent = Me
        FmOpen_jn_List.WindowState = FormWindowState.Maximized
        FmOpen_jn_List.Show()
        FmOpen_jn_List.Focus()
    End Sub
    Private Sub ToolStripMenuItem9_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem9.Click
        IsMdiContainer = True
        Panel4.Visible = False
        FmJeneralJournal_List.MdiParent = Me
        FmJeneralJournal_List.WindowState = FormWindowState.Maximized
        FmJeneralJournal_List.Show()
        FmJeneralJournal_List.Focus()
        'FmJeneralJournal_List.FG.Size = New System.Drawing.Size(1260, 388)
        FmJeneralJournal_List.FG.set_ColHidden(11, False)
        FmJeneralJournal_List.FG.set_ColHidden(5, True)
        FmJeneralJournal_List.FG.set_ColHidden(12, False)
        FmJeneralJournal_List.FG.set_ColHidden(13, True)
        FmJeneralJournal_List.FG.set_ColHidden(14, True)
        FmJeneralJournal_List.FG.set_ColHidden(15, True)
        FmJeneralJournal_List.FG.set_ColHidden(16, True)
        'FmJeneralJournal_List.FG.set_ColHidden(17, True)
    End Sub
    Private Sub ToolStripMenuItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem11.Click
        Fmdelete.ShowDialog()
    End Sub
    Private Sub ToolStripMenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem10.Click
        FmClosing.ShowDialog()
    End Sub
    Private Sub ສດຄດໄລToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ສດຄດໄລToolStripMenuItem.Click
        FmCaculate_Rpt.ShowDialog()
    End Sub
    Private Sub KpkoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KpkoToolStripMenuItem.Click
        FmRpt_JeneralJournal.nn.Text = KpkoToolStripMenuItem.Text
        FmRpt_JeneralJournal.ShowDialog()
    End Sub
    Private Sub ລາວToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MuLngL.Click
        'LoadLng()
        'MuLngL.Checked = True
        'MuLngE.Checked = False
        'MuLng = "L"
        'SetControlText(Me)
        'ChgChildForm()

        'MnLaoLang.Checked = True
        'MnEngLang.Checked = False


        'PictureBox7.BackgroundImage = Image.FromFile(My.Application.Info.DirectoryPath & "\BackGroundLOGO.png")
        'MuLng = "L"
        LoadLng()
        MuLngL.Checked = True
        MuLngE.Checked = False
        MuLng = "L"
        Lang = False
        SetControlText(Me)
        ChgChildForm()
        MnLaoLang.Checked = True
        MnEngLang.Checked = False
        LoadLngnnn()
        LoadlangMM()
        PictureBox7.BackgroundImage = Image.FromFile(My.Application.Info.DirectoryPath & "\BackGroundLOGO.png")
        MuLng = "L"
        'FG.Cols = 30
        'FG.Size = New System.Drawing.Size(1028, 43)
        'MuLng = "L"
        'LoadLng()
        'SetControlText(Me)
        '=======
        Call Office()
        'MsgBox(RptSjOff)
        'Dim a As String = Environment.MachineName.ToString
        ''Timer1.Start()
        'If a <> MDServerName Then
        '    ToolStripMenuItem57.Enabled = False
        'Else
        '    ToolStripMenuItem57.Enabled = True
        'End If
        FormMain = "ApBank"

        Panel3.Visible = False
        Call Login()
        Me.MenuStrip1.Refresh()
        Timer1.Enabled = True
        '==============
        MWorkSetting = Today
        LoadHideMenu()
        Load_Curr()
    End Sub
    Private Sub ອງກດToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MuLngE.Click
        LoadLng()
        MuLngL.Checked = False
        MuLngE.Checked = True
        MuLng = "E"
        Lang = True
        SetControlText(Me)
        ChgChildForm()
        MnLaoLang.Checked = False
        MnEngLang.Checked = True
        LoadLngnnn()
        LoadlangMM()
    End Sub
    Private Sub ໃບລາຍງານກະແສຄງເງນລວມToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ໃບລາຍງານກະແສຄງເງນລວມToolStripMenuItem.Click
        FmRpt_BLS8.ShowDialog()
    End Sub
    Private Sub LangToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LangToolStripMenuItem.Click
        IsMdiContainer = True
        Panel4.Visible = False
        FmLangaugeSeting.MdiParent = Me
        FmLangaugeSeting.WindowState = FormWindowState.Maximized
        FmLangaugeSeting.Show()
    End Sub
    Private Sub ຕງຄາອດຕາແລກປຽນToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ຕງຄາອດຕາແລກປຽນToolStripMenuItem.Click
        'FmRate.ShowDialog()
        'Rate_setting.MdiParent = Me
        'Rate_setting.WindowState = FormWindowState.Maximized
        Rate_setting.ShowDialog()
    End Sub
    Private Sub ການປຽນແປງຂອງເງນຕາToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ການປຽນແປງຂອງເງນຕາToolStripMenuItem.Click
        FmRateStatus.ShowDialog()
    End Sub
    Private Sub ສວນປຽນແປງທນToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ສວນປຽນແປງທນToolStripMenuItem.Click
        FmAmtStatus123.ShowDialog()
    End Sub
    Private Sub ປຽນຖານຂມນToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ປຽນຖານຂມນToolStripMenuItem.Click
        System.Diagnostics.Process.Start(My.Application.Info.DirectoryPath & "\Conection_To_Server.exe")
    End Sub
    Private Sub CmbForm_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles CmbForm.MouseDown
        Call FormOpening()
    End Sub
    Private Sub CmbForm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbForm.SelectedIndexChanged
        For i As Integer = 0 To (My.Application.OpenForms.Count - 1)
            If My.Application.OpenForms.Item(i).Name() = CmbForm.Text Then
                My.Application.OpenForms.Item(i).BringToFront()
            End If
        Next i
        lblpage_total.Text = "0/0"
        If CmbForm.SelectedIndex >= 0 Then
            lblpage_total.Text = CmbForm.SelectedIndex + 1 & "/" & CmbForm.Items.Count
        End If
        If EnterPageM.Text = "Next" Then
            If CmbForm.SelectedIndex = CmbForm.Items.Count - 1 Then
                EnterPageM.Text = "Back"
            End If
        Else

            If CmbForm.SelectedIndex = 0 Then
                EnterPageM.Text = "Next"
            End If
        End If
    End Sub
    Private Sub FirstPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FirstPage.Click
        Call FormOpening()
        If CmbForm.Items.Count > 0 Then
            CmbForm.SelectedIndex = 0
        End If
    End Sub
    Private Sub LasthPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LasthPage.Click
        Call FormOpening()
        If CmbForm.Items.Count > 0 Then
            CmbForm.SelectedIndex = CmbForm.Items.Count - 1
        End If
    End Sub
    Private Sub NextPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NextPage.Click
        If CmbForm.SelectedIndex < CmbForm.Items.Count - 1 Then
            CmbForm.SelectedIndex = CmbForm.SelectedIndex + 1
        End If
        lblpage_total.Text = "0/0"
        If CmbForm.SelectedIndex >= 0 Then
            lblpage_total.Text = CmbForm.SelectedIndex + 1 & "/" & CmbForm.Items.Count
        End If
    End Sub
    Private Sub BackPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackPage.Click
        If CmbForm.SelectedIndex > 0 Then
            CmbForm.SelectedIndex = CmbForm.SelectedIndex - 1
        End If
        lblpage_total.Text = "0/0"
        If CmbForm.SelectedIndex >= 0 Then
            lblpage_total.Text = CmbForm.SelectedIndex + 1 & "/" & CmbForm.Items.Count
        End If
    End Sub
    Private Sub MnSystem_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles MnSystem.MouseEnter
        'MnSystem.BackgroundImage = Image.FromFile(My.Application.Info.DirectoryPath & "\Menu.png")
    End Sub
    Private Sub MnSystem_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles MnSystem.MouseLeave
        'MnSystem.BackgroundImage = NonSerializedAttribute
    End Sub
    Private Sub EnterPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EnterPageM.Click
        If EnterPageM.Text = "Next" Then
            NextPage_Click(sender, e)
            If CmbForm.SelectedIndex = CmbForm.Items.Count - 1 Then
                EnterPageM.Text = "Back"
            End If
        Else
            BackPage_Click(sender, e)
            If CmbForm.SelectedIndex = 0 Then
                EnterPageM.Text = "Next"
            End If
        End If
    End Sub
    Private Sub NextPage_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NextPage.MouseClick
        'Next  
        EnterPageM.Text = "Next"
        If CmbForm.SelectedIndex = CmbForm.Items.Count - 1 Then
            EnterPageM.Text = "Back"
        End If
    End Sub
    Private Sub BackPage_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BackPage.MouseClick
        EnterPageM.Text = "Back"
        If CmbForm.SelectedIndex = 0 Then
            EnterPageM.Text = "Next"
        End If
    End Sub
    Private Sub lblpage_total_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblpage_total.TextChanged
        NextPage.Enabled = True
        LasthPage.Enabled = True
        BackPage.Enabled = True
        FirstPage.Enabled = True
        Button3.Enabled = True
        EnterPageM.Enabled = True
        CmbForm.Enabled = True
        lblpage_total.Enabled = True
        If CmbForm.Items.Count > 1 Then
            If CmbForm.SelectedIndex = CmbForm.Items.Count - 1 Then
                NextPage.Enabled = False
                LasthPage.Enabled = False
            End If
            If CmbForm.SelectedIndex = 0 Then
                BackPage.Enabled = False
                FirstPage.Enabled = False
            End If
        End If
        If lblpage_total.Text = "0/0" Then
            lblpage_total.Enabled = False
            NextPage.Enabled = False
            LasthPage.Enabled = False
            BackPage.Enabled = False
            FirstPage.Enabled = False
            Button3.Enabled = False
            EnterPageM.Enabled = False
            CmbForm.Enabled = False
        End If
        If lblpage_total.Text = "1/1" Then
            lblpage_total.Enabled = False
            NextPage.Enabled = False
            LasthPage.Enabled = False
            BackPage.Enabled = False
            FirstPage.Enabled = False
            EnterPageM.Enabled = False
            CmbForm.Enabled = False
        End If
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        For i As Integer = 0 To (My.Application.OpenForms.Count - 1)
            If My.Application.OpenForms.Item(i).Name() = CmbForm.Text Then
                My.Application.OpenForms.Item(i).BringToFront()
            End If
        Next i
    End Sub
    Private Sub ລາຍງານການຄດໄລອງປະກອບຂອງຊບສນToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ລາຍງານການຄດໄລອງປະກອບຂອງຊບສນToolStripMenuItem.Click
        FmRptPro_New.LoadDateNow()
        FmRptPro_New.txtRptNme.Text = ລາຍງານການຄດໄລອງປະກອບຂອງຊບສນToolStripMenuItem.Text
        MUTY = "PRO2"
        FmRptPro_New.ShowDialog()
    End Sub
    Private Sub ຕາຕະລາງຄດໄລອງປະກອບຂອງໜສນToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ຕາຕະລາງຄດໄລອງປະກອບຂອງໜສນToolStripMenuItem.Click
        FmRptPro_New.txtRptNme.Text = ຕາຕະລາງຄດໄລອງປະກອບຂອງໜສນToolStripMenuItem.Text
        MUTY = "PRO3"
        FmRptPro_New.ShowDialog()
    End Sub
    Private Sub ການຄດໄລຄວາມສຽງຂອງຊບສນໃນໃບສະຫບຊບສມບດToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ການຄດໄລຄວາມສຽງຂອງຊບສນໃນໃບສະຫບຊບສມບດToolStripMenuItem.Click
        FmRptPro_New.txtRptNme.Text = ການຄດໄລຄວາມສຽງຂອງຊບສນໃນໃບສະຫບຊບສມບດToolStripMenuItem.Text
        MUTY = "PRO4"
        FmRptPro_New.ShowDialog()
    End Sub
    Private Sub ToolStripMenuItem12_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem12.Click
        FmRptPro_New.txtRptNme.Text = ToolStripMenuItem12.Text
        MUTY = "PRO1"
        FmRptPro_New.ShowDialog()
    End Sub
    Private Sub Label22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label22.Click
        CNN.Execute("delete Caculate_Rpt_Update  ")
        CNN.Execute("insert into Caculate_Rpt_Update (Rpt_Id,CLT_Amt)  " & _
"select Rpt_Id , STUFF((  select ' '+b.CLT_Amt from Caculate_Rpt b   where b.Rpt_Id = a.Rpt_Id    " & _
"order by b.cnt for xml path('a'), type).value('.','nvarchar(max)'),1,1,'') As  CLT_Amt      " & _
" from Caculate_Rpt a where CLT_Amt <>''group by Rpt_Id ")
    End Sub

    Private Sub CaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CaToolStripMenuItem.Click
        'FmCaculate_Pro.ShowDialog()
        Frm_Group_accode.Show()
    End Sub





    Private Sub ToolStripMenuItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem13.Click
        'frmBackup_Restore.Show()
        Form1111.Show()
    End Sub

    Private Sub ToolStripMenuItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem14.Click
        IsMdiContainer = True
        Panel4.Visible = False
        Frm_import_exel.MdiParent = Me
        Frm_import_exel.WindowState = FormWindowState.Maximized
        Frm_import_exel.ShowIcon = False
        Frm_import_exel.Show()

        'IsMdiContainer = True
        'Panel4.Visible = False
        'Frm_import_exel_New.MdiParent = Me
        'Frm_import_exel_New.WindowState = FormWindowState.Maximized
        'Frm_import_exel_New.ShowIcon = False
        'Frm_import_exel_New.Show()

    End Sub

    Private Sub ToolStripMenuItem15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem15.Click

        FrmRpt_F01.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem16.Click


        FrmRpt_F04.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem17_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem17.Click

        FrmRpt_F05.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem19_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem19.Click
        FrmRpt_F06.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem20.Click
        FrmRpt_F07.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem21.Click
        FrmRpt_F08.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem22_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem22.Click
        IsMdiContainer = True
        Panel4.Visible = False


        Frm_import_exel_AR.MdiParent = Me
        Frm_import_exel_AR.WindowState = FormWindowState.Maximized
        Frm_import_exel_AR.ShowIcon = False
        Frm_import_exel_AR.Show()
    End Sub

    Private Sub ToolStripMenuItem24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem24.Click
        'IsMdiContainer = True
        'Panel4.Visible = False
        'Frm_import_exel_KS.MdiParent = Me
        'Frm_import_exel_KS.WindowState = FormWindowState.Maximized
        'Frm_import_exel_KS.ShowIcon = False
        'Frm_import_exel_KS.Show()
        '=======
        IsMdiContainer = True
        Panel4.Visible = False
        Frm_import_exel_KS_DG.MdiParent = Me
        Frm_import_exel_KS_DG.WindowState = FormWindowState.Maximized
        Frm_import_exel_KS_DG.ShowIcon = False
        Frm_import_exel_KS_DG.Show()

        'IsMdiContainer = True
        'Panel4.Visible = False
        'frmImportExcel.MdiParent = Me
        'frmImportExcel.WindowState = FormWindowState.Maximized
        'frmImportExcel.ShowIcon = False
        'frmImportExcel.Show()
    End Sub

    Private Sub ToolStripMenuItem39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem39.Click
        'FrmData_server.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem25.Click
        FrmImport_Rate.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem26.Click
        IsMdiContainer = True
        Panel4.Visible = False
        Frm_import_exel_KS_BL.MdiParent = Me
        Frm_import_exel_KS_BL.WindowState = FormWindowState.Maximized
        Frm_import_exel_KS_BL.ShowIcon = False
        Frm_import_exel_KS_BL.Show()

    End Sub

    Private Sub ToolStripMenuItem28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem28.Click
        FmRptPro_New.txtRptNme.Text = ToolStripMenuItem28.Text
        MUTY = "PRO6"
        FmRptPro_New.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem30_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem30.Click
        MDBSL = "1"
        FmRpt_BLS_BCEL.nn.Text = ToolStripMenuItem30.Text
 
        FmRpt_BLS_BCEL.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem31_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem31.Click
        FmRpt_Income_Old2.nn.Text = ToolStripMenuItem31.Text

        FmRpt_Income_Old2.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem32_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem32.Click
        'FmRpt_BLS8.nn.Text = ToolStripMenuItem32.Text

        'FmRpt_BLS8.ShowDialog()
        'FmRpt_CashNew.nn.Text = ToolStripMenuItem32.Text 
        'FmRpt_CashNew.ShowDialog()
        FmRpt_BLS_NEW.nn.Text = ToolStripMenuItem32.Text


        FmRpt_BLS_NEW.ShowDialog()

    End Sub

    Private Sub ToolStripMenuItem33_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem33.Click
        FmAmtStatusNEW.nn.Text = ToolStripMenuItem33.Text
         
        FmAmtStatusNEW.ShowDialog()
         
    End Sub

    Private Sub ToolStripMenuItem45_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem45.Click
        FmRptPro.txtRptNme.Text = ToolStripMenuItem45.Text
        MUTY = "PRO4"
        FmRptPro.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem44_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem44.Click
        FmRptPro_New.txtRptNme.Text = ToolStripMenuItem44.Text
        MUTY = "PRO1"
        FmRptPro.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem42.Click
        MDBSL = "2"
        FmRpt_BLS_BCEL.nn.Text = ToolStripMenuItem42.Text
        FmRpt_BLS_BCEL.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem43_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem43.Click
        MDBSL = "3"
        FmRpt_BLS_BCEL.nn.Text = ToolStripMenuItem43.Text
        FmRpt_BLS_BCEL.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem48_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem48.Click
        FmRptPro_01.txtRptNme.Text = ToolStripMenuItem48.Text
        FmRptProItem.Label5.Text = ToolStripMenuItem48.Text

        MUTY = "PRO1"


        FmRptPro_01.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem50_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem50.Click
        FmRptPro.LoadDateNow()
        FmRptPro.txtRptNme.Text = ToolStripMenuItem50.Text
        MUTY = "PRO2"
        FmRptPro.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem51_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem51.Click
        FmRptPro.txtRptNme.Text = ToolStripMenuItem51.Text
        MUTY = "PRO3"
        FmRptPro.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem54_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem54.Click
        FmRptPro.txtRptNme.Text = ToolStripMenuItem54.Text
        MUTY = "PRO4"
        FmRptPro.ShowDialog()
    End Sub

    Private Sub ຕງຄາສຳປະສດToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ຕງຄາສຳປະສດToolStripMenuItem.Click
        Frm_mformat.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem47_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem47.Click
        EditActive = False
        'Frm_Acc_Adjust_Curr.MdiParent = Me
        'Frm_Acc_Adjust_Curr.WindowState = FormWindowState.Maximized
        'Frm_Acc_Adjust_Curr.Show()
        'Frm_Acc_Adjust_Curr.ShowDialog()
        IsMdiContainer = True
        Panel4.Visible = False
        Frm_Acc_Adjust_Curr.MdiParent = Me
        Frm_Acc_Adjust_Curr.WindowState = FormWindowState.Maximized
        Frm_Acc_Adjust_Curr.Show()
        Frm_Acc_Adjust_Curr.Focus()
    End Sub

    Private Sub ToolStripMenuItem46_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem46.Click
        IsMdiContainer = True
        Panel4.Visible = False
        FmJeneralJournal_Adjust_List.MdiParent = Me
        FmJeneralJournal_Adjust_List.WindowState = FormWindowState.Maximized
        FmJeneralJournal_Adjust_List.Show()
        FmJeneralJournal_Adjust_List.Focus()
        'FmJeneralJournal_List.FG.Size = New System.Drawing.Size(1260, 388)
        FmJeneralJournal_Adjust_List.FG.set_ColHidden(11, False)
        FmJeneralJournal_Adjust_List.FG.set_ColHidden(5, True)
        FmJeneralJournal_Adjust_List.FG.set_ColHidden(12, False)
        FmJeneralJournal_Adjust_List.FG.set_ColHidden(13, True)
        FmJeneralJournal_Adjust_List.FG.set_ColHidden(14, True)
        FmJeneralJournal_Adjust_List.FG.set_ColHidden(15, True)
        FmJeneralJournal_Adjust_List.FG.set_ColHidden(16, True)
    End Sub

    Private Sub ToolStripMenuItem55_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem55.Click
        FmRptPro_New.txtRptNme.Text = ToolStripMenuItem55.Text
        MUTY = "PRO6"
        FmRptPro_New.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem56_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem56.Click
        IsMdiContainer = True
        Panel4.Visible = False
        FrmCustomer.MdiParent = Me
        FrmCustomer.WindowState = FormWindowState.Maximized
        FrmCustomer.Show()
        FrmCustomer.Focus()
        'FrmCustomer.MdiParent = Me
        'FrmCustomer.WindowState = FormWindowState.Maximized
        'FrmCustomer.Show()
    End Sub

    Private Sub ToolStripMenuItem58_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem58.Click
        IsMdiContainer = True
        Panel4.Visible = False
        FrmSupplier.MdiParent = Me
        FrmSupplier.WindowState = FormWindowState.Maximized
        FrmSupplier.Show()
        FrmSupplier.Focus()
    End Sub

    Private Sub ToolStripMenuItem59_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem59.Click
        IsMdiContainer = True
        Panel4.Visible = False
        Frm_Statement.MdiParent = Me
        Frm_Statement.WindowState = FormWindowState.Maximized
        Frm_Statement.Show()
        Frm_Statement.Focus()
    End Sub

    Private Sub ToolStripMenuItem61_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem61.Click
 
        IsMdiContainer = True
        Panel4.Visible = False
        FrmAsset_List_LSNEW.MdiParent = Me
        FrmAsset_List_LSNEW.WindowState = FormWindowState.Maximized
        FrmAsset_List_LSNEW.Show()
        FrmAsset_List_LSNEW.Focus()
    End Sub

    Private Sub ToolStripMenuItem62_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem62.Click
        IsMdiContainer = True
        Panel4.Visible = False
        FrmDepartment_List.MdiParent = Me
        FrmDepartment_List.WindowState = FormWindowState.Maximized
        FrmDepartment_List.Show()
        FrmDepartment_List.Focus()
    End Sub

    Private Sub ToolStripMenuItem63_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem63.Click
        IsMdiContainer = True
        Panel4.Visible = False
        FrmAsset_List.MdiParent = Me
        FrmAsset_List.WindowState = FormWindowState.Maximized
        FrmAsset_List.Show()
        FrmAsset_List.Focus()
    End Sub

    Private Sub ToolStripMenuItem64_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem64.Click
        IsMdiContainer = True
        Panel4.Visible = False
        FrmRpt_Group.MdiParent = Me
        FrmRpt_Group.WindowState = FormWindowState.Maximized
        FrmRpt_Group.Show()
        FrmRpt_Group.Focus()
    End Sub

    Private Sub ToolStripMenuItem66_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem66.Click
        IsMdiContainer = True
        Panel4.Visible = False
        FrmAsset_Grp.Label3.Text = ToolStripMenuItem66.Text
        FrmGrpNew_LS.LDetail.Text = ToolStripMenuItem66.Text

        FrmAsset_Grp.MdiParent = Me
        FrmAsset_Grp.WindowState = FormWindowState.Maximized
        FrmAsset_Grp.Show()
        FrmAsset_Grp.Focus()
    End Sub

    Private Sub ToolStripMenuItem68_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem68.Click
        IsMdiContainer = True
        Panel4.Visible = False
        FrmAdjustment_List.Label11.Text = ToolStripMenuItem68.Text

        FrmAdjustment_List.MdiParent = Me
        FrmAdjustment_List.WindowState = FormWindowState.Maximized
        FrmAdjustment_List.Show()
        FrmAdjustment_List.Focus()
    End Sub

    Private Sub ToolStripMenuItem69_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem69.Click
        IsMdiContainer = True
        Panel4.Visible = False
        FrmAdjustment_App.Label11.Text = ToolStripMenuItem69.Text

        FrmAdjustment_App.MdiParent = Me
        FrmAdjustment_App.WindowState = FormWindowState.Maximized
        FrmAdjustment_App.Show()
        FrmAdjustment_App.Focus()
    End Sub

    Private Sub ToolStripMenuItem5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem5.Click
        'Call txtUserID.Focus()

        Call ConnectAccess()
 

        Dim Conn As New ADODB.Connection
        Dim rsProj As New ADODB.Recordset
        Call LoadAcData("Select * from Conect ", rsProj)
        With rsProj
            If .RecordCount <> 0 Then
                MDServerName = (.Fields("ServerName").Value.ToString)
                MDDatabaName = (.Fields("DatabaseName").Value.ToString)
                MDServerUser = (.Fields("UserName").Value.ToString)
                MDServerPassword = (.Fields("UserPassword").Value.ToString)
                MDSeriaAccess = (.Fields("PartitionSeria").Value.ToString)
                'SPW = CStr((.Fields("SavePassword").Value.ToString))
                'SUSID = CStr((.Fields("SaveUserID").Value.ToString))
            End If
        End With
        Call ConnectSQL()
        If VSysError = True Then

            'FrmData_server.Show()
            'Me.Hide()
            'Exit Sub
        Else
            MsgBox("Complete")
        End If
    End Sub

    Private Sub ToolStripMenuItem67_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripMenuItem70_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripMenuItem71_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem71.Click
        IsMdiContainer = True
        Panel4.Visible = False
        FrmAsset_List.MdiParent = Me
        FrmAsset_List.WindowState = FormWindowState.Maximized
        FrmAsset_List.Show()
        FrmAsset_List.Focus()
    End Sub

    Private Sub ToolStripMenuItem72_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem72.Click
        'IsMdiContainer = True
        'Panel4.Visible = False
        'FrmRpt_Fixed_Assets.MdiParent = Me
        'FrmRpt_Fixed_Assets.WindowState = FormWindowState.Maximized
        'FrmRpt_Fixed_Assets.Show()
        'FrmRpt_Fixed_Assets.Focus()
        IsMdiContainer = True
        Panel4.Visible = False
        FrmRpt_Fixed_Assets_NEW.MdiParent = Me
        FrmRpt_Fixed_Assets_NEW.WindowState = FormWindowState.Maximized
        FrmRpt_Fixed_Assets_NEW.Show()
        FrmRpt_Fixed_Assets_NEW.Focus()
    End Sub

    Private Sub ToolStripMenuItem67_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem67.Click
        IsMdiContainer = True
        Panel4.Visible = False
        FrmAsset_Broke.MdiParent = Me
        FrmAsset_Broke.WindowState = FormWindowState.Maximized
        FrmAsset_Broke.Show()
        FrmAsset_Broke.Focus()
    End Sub

    Private Sub Label28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label28.Click

    End Sub

    Private Sub ລາຍງToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ລາຍງToolStripMenuItem.Click
        'EditActive = False
        'FrmRpt_BLS_liability.MdiParent = Me
        'FrmRpt_BLS_liability.WindowState = FormWindowState.Maximized
        FmRpt_Depreciation.ShowDialog()
    End Sub

    Private Sub ບນຊແຍກປະເພດແບບທ2ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ບນຊແຍກປະເພດແບບທ2ToolStripMenuItem.Click

        FmPostedLedgers_From2.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem36.Click
        System.Diagnostics.Process.Start(".\\APTAX_NEW.pdf")
    End Sub

    Private Sub ບນຊສຳຮອງToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ບນຊສຳຮອງToolStripMenuItem.Click
        FmPostedLedgers_From3.ShowDialog()
    End Sub
End Class