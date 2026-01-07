'Imports System.IO
Public Class FmLogOut
    Dim SPW As String = ""
    Dim SUSID As String = ""
    Dim Ck, Status1 As Integer
    ' Dim rs As New ADODB.Recordset ' REMOVED - ADODB migration
    ' Dim RSCC4 As New ADODB.Recordset ' REMOVED - ADODB migration
    Dim MUserName2, MUserID2, MPermit2, MSection2, MPws2 As String








    'Dim ImageSlno As Integer
    'Public con As New OleDb.OleDbConnection("Provider=SQLOLEDB;User id=" & MDServerUser & ";database=" & MDDatabaName & ";password=" & MDServerPassword & ";data source=" & MDServerName & "")
    Private Sub FmLogOut_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub SavePsswordAndUserID()
        Dim dt As DataTable = DbHelper.GetDataTable("select * from AP_Users where Usr_id='" & Trim(Apostrophe(txtUserId.Text)) & "' ")
        If dt.Rows.Count = 0 Then
            'MessageBox.Show("User ID ບໍ່ຖືກຕ້ອງ", "ບໍ່ສາມາດເຂົ້າລະບົບໄດ້", MessageBoxButtons.OK, MessageBoxIcon.Error) : txtUserId.Focus() : Exit Sub
        Else
            Dim row As DataRow = dt.Rows(0)
            MUserID2 = Trim(DbHelper.GetStr(row("Usr_id")))
            MUserName = Trim(DbHelper.GetStr(row("Usr_nm")))
            cmbCompany.Text = Trim(DbHelper.GetStr(row("Company")))
            Sub_Company.Text = Trim(DbHelper.GetStr(row("Sub_Company")))
            MUserID = Trim(DbHelper.GetStr(row("Usr_id")))
            MPws = Trim(DbHelper.GetStr(row("PWD")))
            MPermit = Trim(DbHelper.GetStr(row("permision")))

            Label11.Text = Trim(DbHelper.GetStr(row("Usr_nm")))
            Label12.Text = MPermit
            MSection = DbHelper.GetStr(row("Sec_ID"))

            txtPassword.Focus()
        End If
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
        Dim dt As DataTable = DbHelper.GetDataTable("select * from AP_Users where Usr_id='" & Trim(Apostrophe(txtUserId.Text)) & "' ")
        If dt.Rows.Count = 0 Then
            'MessageBox.Show("User ID ບໍ່ຖືກຕ້ອງ", "ບໍ່ສາມາດເຂົ້າລະບົບໄດ້", MessageBoxButtons.OK, MessageBoxIcon.Error) : txtUserId.Focus() : Exit Sub
        Else
            Dim row As DataRow = dt.Rows(0)
            MUserName2 = Trim(DbHelper.GetStr(row("Usr_nm")))
            cmbCompany.Text = Trim(DbHelper.GetStr(row("Company")))
            Sub_Company.Text = Trim(DbHelper.GetStr(row("Sub_Company")))
            MUserID2 = Trim(DbHelper.GetStr(row("Usr_id")))
            MPws2 = Trim(DbHelper.GetStr(row("PWD")))
            MPermit2 = Trim(DbHelper.GetStr(row("permision")))
            Label11.Text = Trim(DbHelper.GetStr(row("Usr_nm")))
            Label12.Text = MPermit
            MSection2 = DbHelper.GetStr(row("Sec_ID"))
            txtPassword.Focus()
        End If
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
        Dim dt As DataTable = DbHelper.GetDataTable("Select * from PTSSeria ")
        If dt.Rows.Count <> 0 Then
            MDSeriaAccess = CDbl(DbHelper.GetStr(dt.Rows(0)("PartitionSeria")))
        End If
    End Sub
    Private Sub loadCheckComputerCode()
        Dim dt1 As DataTable = DbHelper.GetDataTable("select *  from  Ap_ComputerMember Where ComCode = '" & MDSeriaCom & "' ")
        If dt1.Rows.Count = 0 Then
            Ck = 0
        End If

        Dim dt2 As DataTable = DbHelper.GetDataTable("select *  from  Ap_ComputerMember Where ComCode = '" & MDSeriaCom & "' And Status = 1 ")
        If dt2.Rows.Count = 0 Then
            'MsgBox("jj")
            Status1 = 0
        End If
    End Sub
    Private Sub loadCompany()
        cmbCompany.Items.Clear()
        Dim dt As DataTable = DbHelper.GetDataTable("select off_add1 , off_id  from  Ap_office group BY off_id , off_add1")
        For Each row As DataRow In dt.Rows
            cmbCompany.Items.Add(DbHelper.GetStr(row("off_id")) & " " & DbHelper.GetStr(row("off_add1")))
        Next
        SUPD = 0
    End Sub

    Private Sub LoadSubCompany()
        Sub_Company.Items.Clear()
        Dim dt As DataTable = DbHelper.GetDataTable("select sub_id , off_id , off_add2  from  Ap_office where off_id ='" & Mid(cmbCompany.Text, 1, 2) & "' group BY  sub_id  ,off_id , off_add2")
        For Each row As DataRow In dt.Rows
            Sub_Company.Items.Add(DbHelper.GetStr(row("sub_id")) & " " & DbHelper.GetStr(row("off_add2")))
        Next

        Sub_Company.SelectedIndex = 0
        Off_Id = Mid(cmbCompany.Text, 1, 2)
        SUPD = 0
    End Sub
    Private Sub BtnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOk.Click


        If ChSavPW.Checked = True Then
            DbHelper.ExecuteNonQuery(" Update  conect set SavePassword = '" & txtPassword.Text & "' ")
        Else
            DbHelper.ExecuteNonQuery(" Update  conect set SavePassword = '' ")
        End If
        If ChSavUserID.Checked = True Then
            DbHelper.ExecuteNonQuery(" Update  conect set SaveUserID = '" & txtUserId.Text & "' ")
        Else
            DbHelper.ExecuteNonQuery(" Update  conect set SaveUserID = '' ")
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
        ' Retrieve duplicate Bnk_Ac_Code3 values
        Dim dt As DataTable = DbHelper.GetDataTable("select Bnk_Ac_Code3, count(Bnk_Ac_Code) from Ap_Loan group by Bnk_Ac_Code3 having count(Bnk_Ac_Code3)>1")
        If dt.Rows.Count <> 0 Then
            For Each row As DataRow In dt.Rows
                Dim bnkAcCode3 As String = DbHelper.GetStr(row("Bnk_Ac_Code3"))
                x(bnkAcCode3) ' Pass the Bnk_Ac_Code3 value as a parameter
            Next
        End If
        'MsgBox("ok")
    End Sub
    Private Sub x(bnkAcCode3 As String)
        Dim X1 As Double
        X1 = 0
        Dim dt As DataTable = DbHelper.GetDataTable("select * from Ap_Loan where Bnk_Ac_Code3='" & bnkAcCode3 & "'")
        If dt.Rows.Count <> 0 Then
            For Each row As DataRow In dt.Rows
                X1 = X1 + 1
                DbHelper.ExecuteNonQuery(" Update Ap_Loan set Bnk_Ac_Code=Bnk_Ac_Code3 + '-' + '" & X1 & "' where cnt = '" & DbHelper.GetStr(row("cnt")) & "' ")
            Next
        End If
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