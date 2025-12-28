'Imports System.IO
Public Class FmLogOut
    Dim SPW As String = ""
    Dim SUSID As String = ""
    Dim Ck, Status1 As Integer
    Dim rs As New ADODB.Recordset
    Dim RSCC4 As New ADODB.Recordset
    Dim MUserName2, MUserID2, MPermit2, MSection2, MPws2 As String








    'Dim ImageSlno As Integer
    'Public con As New OleDb.OleDbConnection("Provider=SQLOLEDB;User id=" & MDServerUser & ";database=" & MDDatabaName & ";password=" & MDServerPassword & ";data source=" & MDServerName & "")
    Private Sub FmLogOut_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub SavePsswordAndUserID()
        Call LoadSqlData("select * from AP_Users where Usr_id='" & Trim(Apostrophe(txtUserId.Text)) & "' ", rs)
        With rs
            If .RecordCount = 0 Then
                'MessageBox.Show("User ID ບໍ່ຖືກຕ້ອງ", "ບໍ່ສາມາດເຂົ້າລະບົບໄດ້", MessageBoxButtons.OK, MessageBoxIcon.Error) : txtUserId.Focus() : Exit Sub
            Else
                MUserID2 = Trim(.Fields("Usr_id").Value.ToString)
                MUserName = Trim(.Fields("Usr_nm").Value.ToString)
                cmbCompany.Text = Trim(.Fields("Company").Value.ToString)
                Sub_Company.Text = Trim(.Fields("Sub_Company").Value.ToString)
                MUserID = Trim(.Fields("Usr_id").Value)
                MPws = Trim(.Fields("PWD").Value)
                MPermit = Trim(.Fields("permision").Value)




                Label11.Text = Trim(.Fields("Usr_nm").Value.ToString)
                Label12.Text = MPermit
                MSection = (.Fields("Sec_ID").Value)

                txtPassword.Focus()
            End If
        End With
    End Sub


    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
    
        Me.Close()

    End Sub

    Private Sub txtUserId_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUserId.GotFocus
        BtnOk.Enabled = False

    End Sub

    Private Sub txtUserId_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtUserId.KeyPress
        If e.KeyChar = Chr(13) Then
            Call loadCompany()
            Call LoadUser()
            If Label11.Text = "" Then
                MessageBox.Show("User ID ບໍ່ຖືກຕ້ອງ", "ບໍ່ສາມາດເຂົ້າລະບົບໄດ້", MessageBoxButtons.OK, MessageBoxIcon.Error) : txtUserId.Focus() : Exit Sub
            End If
            Fm_Image.Img_ID.Text = txtUserId.Text
            Fm_Image.ImgType.Text = "User"
            Call LoadPhoto()
            PictureBox3.BorderStyle = BorderStyle.Fixed3D
            PictureBox3.Image = Fm_Image.PictureBox1.Image

        End If
    End Sub


    Private Sub LoadUser()
        MPws = ""
        Label11.Text = ""
        Label12.Text = ""
        Call LoadSqlData("select * from AP_Users where Usr_id='" & Trim(Apostrophe(txtUserId.Text)) & "' ", rs)
        With rs
            If .RecordCount = 0 Then
                'MessageBox.Show("User ID ບໍ່ຖືກຕ້ອງ", "ບໍ່ສາມາດເຂົ້າລະບົບໄດ້", MessageBoxButtons.OK, MessageBoxIcon.Error) : txtUserId.Focus() : Exit Sub
            Else
              
                MUserName2 = Trim(.Fields("Usr_nm").Value.ToString)
                cmbCompany.Text = Trim(.Fields("Company").Value.ToString)
                Sub_Company.Text = Trim(.Fields("Sub_Company").Value.ToString)
                MUserID2 = Trim(.Fields("Usr_id").Value)
                MPws2 = Trim(.Fields("PWD").Value)
                MPermit2 = Trim(.Fields("permision").Value)
                Label11.Text = Trim(.Fields("Usr_nm").Value.ToString)
                Label12.Text = MPermit
                MSection2 = (.Fields("Sec_ID").Value)
                txtPassword.Focus()
            End If
        End With
    End Sub

    Private Sub txtUserId_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUserId.LostFocus


        Call loadCompany()
        'Call LoadUser()
        'If Label11.Text = "" Then
        '    MessageBox.Show("User ID ບໍ່ຖືກຕ້ອງ", "ບໍ່ສາມາດເຂົ້າລະບົບໄດ້", MessageBoxButtons.OK, MessageBoxIcon.Error) : txtUserId.Focus() : Exit Sub
        'End If
        'Fm_Image.Img_ID.Text = txtUserId.Text
        'Fm_Image.ImgType.Text = "User"
        'Call LoadPhoto()
        'PictureBox3.Image = Fm_Image.PictureBox1.Image









        Call LoadUser()

        'If MPws <> "" Then
        '    CheckBox1.Checked = True
        'Else
        '    CheckBox1.Checked = False
        'End If
        Fm_Image.Img_ID.Text = txtUserId.Text
        Fm_Image.ImgType.Text = "User"
        Call LoadPhoto()
        PictureBox3.BorderStyle = BorderStyle.Fixed3D
        PictureBox3.Image = Fm_Image.PictureBox1.Image

        cmbCompany.Enabled = False
        Sub_Company.Enabled = False
        If Label12.Text = "Sub-Admin" Then
            cmbCompany.Enabled = False
            Sub_Company.Enabled = True

        End If
        If Label12.Text = "Admin" Then

            cmbCompany.Enabled = True
            Sub_Company.Enabled = True
        End If
        If Label12.Text = "User" Then
            cmbCompany.Enabled = False
            Sub_Company.Enabled = False

        End If

    End Sub
    'Private Sub LoadPhoto(ByVal slno As Decimal)

    '    Try
    '        MsgBox("lk")
    '        Dim str As String = "SELECT * FROM Ap_Image WHERE Img_Id = '000010'"
    '        con.Open()
    '        Dim cmd As New OleDb.OleDbCommand(str, con)
    '        Dim b() As Byte
    '        b = cmd.ExecuteScalar()
    '        con.Close()
    '        If (b.Length > 0) Then
    '            Dim stream As New MemoryStream(b, True)
    '            stream.Write(b, 0, b.Length)
    '            DrawToScale(New Bitmap(stream))
    '            stream.Close()
    '        End If
    '    Catch ex As Exception
    '        MsgBox("No Image Is There")
    '    End Try

    'End Sub

    'Private Sub DrawToScale(ByVal bmp As Image)
    '    PictureBox4.Image = New Bitmap(bmp)
    'End Sub
    'Sub CheckUserNm()
    '    Dim rsProj As New ADODB.Recordset
    '    With rsProj
    '        Call LoadSqlData("select * from AP_Users where Usr_id='" & Trim(Apostrophe(txtUserId.Text)) & "' AND Company ='" & cmbCompany.Text & "' ", rsProj)
    '        Do Until .EOF = True
    '            MUserID = Trim(.Fields("Usr_id").Value)
    '            MUserName = Trim(.Fields("Usr_nm").Value)
    '            MPws = Trim(.Fields("PWD").Value)
    '            MPermit = Trim(.Fields("permision").Value)
    '            'ForStaff = (.Fields("ForStaff").Value)
    '            'MDWrite = (.Fields("Write_bit").Value)
    '            'MDEdit = (.Fields("Edit_bit").Value)
    '            'MDDelete = (.Fields("Delete_bit").Value)
    '            MSection = (.Fields("Sec_ID").Value)
    '            .MoveNext()
    '        Loop
    '    End With
    'End Sub

    Private Sub txtUserId_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUserId.TextChanged
        BtnOk.Enabled = False
    End Sub
   

    Private Sub LoadPartition()
        ConnectPartition()
        Dim Conn As New ADODB.Connection
        Dim rsProj As New ADODB.Recordset
        Call LoadAcData("Select * from PTSSeria ", rsProj)
        With rsProj
            If .RecordCount <> 0 Then
                MDSeriaAccess = CDbl((.Fields("PartitionSeria").Value.ToString))
            End If
        End With
    End Sub
    Private Sub loadCheckComputerCode()

        Dim RSC As New ADODB.Recordset
        With RSC
            LoadSqlData("select *  from  Ap_ComputerMember Where ComCode = '" & MDSeriaCom & "' ", RSC)
            If RSC.RecordCount = 0 Then
                Ck = 0
            End If
        End With

        'Dim RSC As New ADODB.Recordset
        With RSC
            LoadSqlData("select *  from  Ap_ComputerMember Where ComCode = '" & MDSeriaCom & "' And Status = 1 ", RSC)
            If RSC.RecordCount = 0 Then
                'MsgBox("jj")
                Status1 = 0
            End If
        End With
    End Sub
    Private Sub loadCompany()
        cmbCompany.Items.Clear()
        LoadSqlData("select off_add1 , off_id  from  Ap_office group BY off_id , off_add1", RSC)
        With RSC
            Do Until .EOF = True
                cmbCompany.Items.Add((.Fields("off_id").Value) & " " & (.Fields("off_add1").Value))
                .MoveNext()
            Loop
        End With

        SUPD = 0
    End Sub

    Private Sub LoadSubCompany()
        Sub_Company.Items.Clear()
        LoadSqlData("select sub_id , off_id , off_add2  from  Ap_office where off_id ='" & Mid(cmbCompany.Text, 1, 2) & "' group BY  sub_id  ,off_id , off_add2", RSC)
        With RSC
            Do Until .EOF = True
                Sub_Company.Items.Add((.Fields("sub_id").Value) & " " & (.Fields("off_add2").Value))
                .MoveNext()
            Loop
        End With

        Sub_Company.SelectedIndex = 0
        Off_Id = Mid(cmbCompany.Text, 1, 2)
        SUPD = 0
    End Sub
    Private Sub BtnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOk.Click


        If ChSavPW.Checked = True Then
            conn.Execute(" Update  conect set SavePassword = '" & txtPassword.Text & "' ")
        Else
            conn.Execute(" Update  conect set SavePassword = '' ")
        End If
        If ChSavUserID.Checked = True Then
            conn.Execute(" Update  conect set SaveUserID = '" & txtUserId.Text & "' ")
        Else
            conn.Execute(" Update  conect set SaveUserID = '' ")
        End If

        Ck = 1
        Status1 = 1

       

        'Call CheckUserNm()

        LoadChecOwner()
        If ScOwner = 0 Then
            LoadCheLock()
            If ScLock = 1 Then

                Me.Close()
                Exit Sub
            End If
            LoadCheckRecor()
            LoadPermitRecord()
            If CDbl(ScRecordUsing) >= CDbl(ScPermitRecord) Then
                LockProgrome()
            End If
            LoadChecPermitSave()
            LoadChecScSaving()
            If CDbl(ScSaving) >= CDbl(ScPermitSave) Then
                LockProgrome()
            End If
        End If






        'MsgBox(MPws)
        'MsgBox(txtPassword.Text)
        If txtUserId.Text <> MUserID2 Then MsgBox(" No have UserID in data base ", MsgBoxStyle.OkOnly) : txtUserId.Focus() : Exit Sub
        If txtPassword.Text <> MPws2 Then MsgBox("No have Password in data base ", MsgBoxStyle.OkOnly) : txtPassword.Focus() : Exit Sub







        MuSubOff = Mid(Sub_Company.Text, 1, 5)

        Call LoadfindLogOut()

        '==========7

        MuLng = "L"
        '=======
        Call Office()
        'MsgBox(RptSjOff)
        Dim a As String = Environment.MachineName.ToString
        'Timer1.Start()
        If a <> MDServerName Then
            FmMain.ToolStripMenuItem57.Enabled = False
        Else
            FmMain.ToolStripMenuItem57.Enabled = True
        End If
        FormMain = "ApBank"
        FmMain.Panel3.Visible = False
        Call FmMain.Login()





        FmMain.PictureBox3.Image = PictureBox3.Image
        Fm_Image.Img_ID.Text = txtUserId.Text
        Fm_Image.ImgType.Text = "Back"
        Call LoadPhoto()
        FmMain.BackgroundImage = Fm_Image.PictureBox1.Image




        'Timer1.Enabled = True




        MUserName = MUserName2

        MUserID = MUserID2
        MPermit = MPermit2
        MSection = MSection2
        MPws = MPws2

        Me.Close()
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub cmbCompany_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCompany.SelectedIndexChanged

        LoadSubCompany()
    End Sub

    Private Sub txtPassword_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPassword.GotFocus
        If txtUserId.Text <> "" Then
            If txtPassword.Text = MPws Then
                BtnOk.Enabled = True
            Else
                BtnOk.Enabled = False
            End If
        Else
            BtnOk.Enabled = False
        End If
    End Sub

    Private Sub txtPassword_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPassword.KeyPress
        If e.KeyChar = Chr(13) Then
            If txtUserId.Text = "" Then
                MessageBox.Show("ກະລຸນນາໃສ່  User ID ກ່ອນ", "ບໍ່ສາມາດເຂົ້າລະບົບໄດ້", MessageBoxButtons.OK, MessageBoxIcon.Error) : txtUserId.Focus() : Exit Sub
            End If
            BtnOk_Click(sender, e)
        End If
    End Sub

    Private Sub txtPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPassword.TextChanged
        If txtUserId.Text <> "" Then
            If txtPassword.Text = MPws Then
                BtnOk.Enabled = True
            Else
                BtnOk.Enabled = False
            End If
        Else
            BtnOk.Enabled = False
        End If

    End Sub

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'MsgBox(Mid(cmbCompany.Text, 1, 2))
    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub


    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click
        'With RSCC4
        '    Call LoadSqlData("select Bnk_Ac_Code3 , count(Bnk_Ac_Code) from Ap_Loan group by Bnk_Ac_Code3 having count(Bnk_Ac_Code3)>1", RSCC4)
        '    If RSCC4.RecordCount <> 0 Then
        '        While Not .EOF()

        '            x()

        '            .MoveNext()
        '        End While
        '    End If
        'End With
        'MsgBox("ok")
    End Sub
    Private Sub x()
        Dim X1 As Double
        X1 = 0
        With RSC
            Call LoadSqlData("select *  from Ap_Loan where Bnk_Ac_Code3='" & (RSCC4.Fields("Bnk_Ac_Code3").Value.ToString) & "'", RSC)
            If .RecordCount <> 0 Then
                While Not .EOF()
                    X1 = X1 + 1
                    CNN.Execute(" Update Ap_Loan set Bnk_Ac_Code=Bnk_Ac_Code3 + '-' + '" & X1 & "' where cnt = '" & (.Fields("cnt").Value.ToString) & "' ")
                    '(.Fields("Withdr_No").Value.ToString) & _

                    .MoveNext()
                End While
            End If
        End With
        'MsgBox("ok")
    End Sub
    Private Sub Label30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label30.Click

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChSavUserID.CheckedChanged
        'If MPws <> "" Then
        '    CheckBox1.Checked = True
        'Else
        '    CheckBox1.Checked = False
        'End If
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Process.Start("")
    End Sub


    Private Sub Sub_Company_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Sub_Company.SelectedIndexChanged

    End Sub
End Class