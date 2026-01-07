'Imports System.IO
Imports System.Data.SqlClient

Public Class FmLogin
    Dim SPW  As String =""
    Dim SUSID As String = ""
    Dim Ck, Status1 As Integer
    ' Dim rs As New ADODB.Recordset ' REMOVED - ADODB migration
    ' Dim RSCC4 As New ADODB.Recordset ' REMOVED - ADODB migration
    'Dim ImageSlno As Integer
    'Public con As New OleDb.OleDbConnection("Provider=SQLOLEDB;User id=" & MDServerUser & ";database=" & MDDatabaName & ";password=" & MDServerPassword & ";data source=" & MDServerName & "")
    Private Sub FmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label3.Text = "Account " & Format(Now, "yyyy")
        'Call LoadImageLocation()
        'PictureBox7.BackgroundImage = Image.FromFile(My.Application.Info.DirectoryPath & "\BackGroundLOGO.png")
        Call ConnectAccess()
        Call LoadServer()
        ' ==============
        Call ConnectSQL()

        'Call SavePsswordAndUserID()



        'loadCheckComputerCode()

        'If Ck = 0 Then
        '    Me.Hide()
        '    FmPartition.TextBox1.Text = MDSeriaCom
        '    FmPartition.ShowDialog()
        '    'FrmShow.Close()
        '    Exit Sub
        'End If

        'If Status1 = 0 Then
        '    Me.Hide()
        '    MsgBox("ບໍ່ທັນໄດ້ອະນຸມັດໃຫ້ເຂົ້າໃຊ້")
        '    Me.Close()
        'End If


        'Call load_Cmb("  select off_add1 , off_id  from  office group BY off_id , off_add1 ", "off_add1 and off_id ", cmbCompany)
        'cmbCompany.SelectedIndex = 0


        'Call UpdateRemoveExS()
        '===============

    End Sub

    Private Sub SavePsswordAndUserID()
        Dim dt As DataTable = DbHelper.GetDataTable("select * from AP_Users where Usr_id='" & Trim(Apostrophe(txtUserId.Text)) & "' ")
        If dt.Rows.Count = 0 Then
            'MessageBox.Show("User ID ບໍ່ຖືກຕ້ອງ", "ບໍ່ສາມາດເຂົ້າລະບົບໄດ້", MessageBoxButtons.OK, MessageBoxIcon.Error) : txtUserId.Focus() : Exit Sub
        Else

            MUserName = Trim(DbHelper.GetStr(dt.Rows(0)("Usr_nm")))
            cmbCompany.Text = Trim(DbHelper.GetStr(dt.Rows(0)("Company")))
            Sub_Company.Text = Trim(DbHelper.GetStr(dt.Rows(0)("Sub_Company")))
            MUserID = Trim(DbHelper.GetStr(dt.Rows(0)("Usr_id")))
            MPws = Trim(DbHelper.GetStr(dt.Rows(0)("PWD")))
            MPermit = Trim(DbHelper.GetStr(dt.Rows(0)("permision")))




            Label11.Text = Trim(DbHelper.GetStr(dt.Rows(0)("Usr_nm")))
            Label12.Text = MPermit
            MSection = DbHelper.GetStr(dt.Rows(0)("Sec_ID"))

            txtPassword.Focus()
        End If
    End Sub


    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        If FormMain = "ApBank" Then
            Me.Hide()
        Else
            Me.Close()
        End If
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
            'Fm_Image.Img_ID.Text = txtUserId.Text
            'Fm_Image.ImgType.Text = "User"
            'Call LoadPhoto()
            'cmbCompany.SelectedIndex = 0
            'Sub_Company.SelectedIndex = 0
            'PictureBox3.BorderStyle = BorderStyle.Fixed3D
            'PictureBox3.Image = Fm_Image.PictureBox1.Image

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

            MUserName = Trim(DbHelper.GetStr(dt.Rows(0)("Usr_nm")))
            cmbCompany.Text = Trim(DbHelper.GetStr(dt.Rows(0)("Company")))
            Sub_Company.Text = Trim(DbHelper.GetStr(dt.Rows(0)("Sub_Company")))
            'MsgBox(Trim(DbHelper.GetStr(dt.Rows(0)("Sub_Company"))))
            MUserID = Trim(DbHelper.GetStr(dt.Rows(0)("Usr_id")))
            MPws = Trim(DbHelper.GetStr(dt.Rows(0)("PWD")))
            MPermit = Trim(DbHelper.GetStr(dt.Rows(0)("permision")))
            Mpermiss = Trim(DbHelper.GetStr(dt.Rows(0)("permision")))
            Label11.Text = Trim(DbHelper.GetStr(dt.Rows(0)("Usr_nm")))
            Label12.Text = MPermit
            'MuLng = Trim(DbHelper.GetStr(dt.Rows(0)("Lng")))
            'ForStaff = DbHelper.GetStr(dt.Rows(0)("ForStaff"))
            'MDWrite = DbHelper.GetStr(dt.Rows(0)("Write_bit"))
            'MDEdit = DbHelper.GetStr(dt.Rows(0)("Edit_bit"))
            'MDDelete = DbHelper.GetStr(dt.Rows(0)("Delete_bit"))
            MSection = DbHelper.GetStr(dt.Rows(0)("Sec_ID"))
            'If MPermit = "Admi" Then
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
        'Fm_Image.Img_ID.Text = txtUserId.Text
        'Fm_Image.ImgType.Text = "User"
        'Call LoadPhoto()
        'PictureBox3.BorderStyle = BorderStyle.Fixed3D
        'PictureBox3.Image = Fm_Image.PictureBox1.Image

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
    Public Sub LoadServer()
        Dim dtProj As DataTable = DbHelper.GetDataTable("Select * from Conect ")
        If dtProj.Rows.Count <> 0 Then
            MDServerName = DbHelper.GetStr(dtProj.Rows(0)("ServerName"))
            MDDatabaName = DbHelper.GetStr(dtProj.Rows(0)("DatabaseName"))
            MDServerUser = DbHelper.GetStr(dtProj.Rows(0)("UserName"))
            MDServerPassword = DbHelper.GetStr(dtProj.Rows(0)("UserPassword"))
            MDSeriaAccess = DbHelper.GetStr(dtProj.Rows(0)("PartitionSeria"))
            SPW = CStr((DbHelper.GetStr(dtProj.Rows(0)("SavePassword"))))
            SUSID = CStr((DbHelper.GetStr(dtProj.Rows(0)("SaveUserID"))))
        End If
        'MsgBox(MDDatabaName)
        'Dim rsProj2 As New ADODB.Recordset
        'Call LoadAcData("Select * from Conect where SvID='002' ", rsProj2)
        'With rsProj2
        '    If .RecordCount <> 0 Then
        '        MDServerName2 = (.Fields("ServerName").Value.ToString)
        '        MDDatabaName2 = (.Fields("DatabaseName").Value.ToString)
        '        MDServerUser2 = (.Fields("UserName").Value.ToString)
        '        MDServerPassword2 = (.Fields("UserPassword").Value.ToString)
        '    End If
        'End With
        If CStr(SPW) <> "" Then
            txtPassword.Text = SPW
            ChSavPW.Checked = True
        Else
            ChSavPW.Checked = False
        End If



        If CStr(SUSID) <> "" Then
            txtUserId.Text = SUSID
            ChSavUserID.Checked = True
        Else
            ChSavUserID.Checked = False
        End If
    End Sub

    Private Sub LoadPartition()
        ConnectPartition()
        Dim dtProj As DataTable = DbHelper.GetDataTable("Select * from PTSSeria ")
        If dtProj.Rows.Count <> 0 Then
            MDSeriaAccess = CDbl((DbHelper.GetStr(dtProj.Rows(0)("PartitionSeria"))))
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
            cmbCompany.Items.Add((DbHelper.GetStr(row("off_id"))) & " " & DbHelper.GetStr(row("off_add1")))
        Next

        SUPD = 0
    End Sub

    Private Sub LoadSubCompany()
        Sub_Company.Items.Clear()
        Dim dt As DataTable = DbHelper.GetDataTable("select sub_id , off_id , off_add2  from  Ap_office where off_id ='" & Mid(cmbCompany.Text, 1, 2) & "' group BY  sub_id  ,off_id , off_add2")
        For Each row As DataRow In dt.Rows
            Sub_Company.Items.Add((DbHelper.GetStr(row("sub_id"))) & " " & DbHelper.GetStr(row("off_add2")))
        Next

        Sub_Company.SelectedIndex = 0
        Off_Id = Mid(cmbCompany.Text, 1, 2)
        SUPD = 0
    End Sub
    Private Sub BtnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOk.Click
        MWorkSetting = Date.Now
        If ChSavPW.Checked = True Then
            DbHelper.ExecuteNonQuery(" Update  conect set SavePassword = '" & txtPassword.Text & "' ")
        Else
            DbHelper.ExecuteNonQuery(" Update  conect set SavePassword ='' ")
        End If
        If ChSavUserID.Checked = True Then
            DbHelper.ExecuteNonQuery(" Update  conect set SaveUserID = '" & txtUserId.Text & "' ")
        Else
            DbHelper.ExecuteNonQuery(" Update  conect set SaveUserID = '' ")
        End If
        Ck = 1
        Status1 = 1
        MDSeriaCom = Environment.MachineName.ToString & "." & CDbl(Mid(MDPartitionSeria.ShowDriveInfo(Application.StartupPath), 2, Len(Trim(MDPartitionSeria.ShowDriveInfo(Application.StartupPath)) - 1)))
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
        If txtUserId.Text <> MUserID Then MsgBox(" No have UserID in data base ", MsgBoxStyle.OkOnly) : txtUserId.Focus() : Exit Sub
        If txtPassword.Text <> MPws Then MsgBox("No have Password in data base ", MsgBoxStyle.OkOnly) : txtPassword.Focus() : Exit Sub
        If FormMain = "ApBank" Then
            Me.Hide()
            If Application.OpenForms.Count > 1 Then
                Conection_To_Servee.Close()
            End If
        Else
            Me.Hide()
            Dim frmMain As New FmMain
            frmMain.WindowState = FormWindowState.Maximized
            MuSubOff = Mid(Sub_Company.Text, 1, 5)
            MuSubOff2 = MuSubOff
            Call Loadfind()
            FmShow.Show()
            'FmShow.Focus()
        End If
        DbHelper.ExecuteNonQuery(" Update  MbDtUse set UsrId = '" & txtUserId.Text & "' , Psw ='" & txtPassword.Text & "'  , OffId = '" & MuSubOff & "' ")
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

        Dim cn As SqlConnection
        Dim cmd As SqlCommand
      
        DbHelper.ExecuteNonQuery("Delete Ap_balance_6_col")

        Try
      
            cn = New SqlConnection("initial catalog=Ap_AccountBCEL4;integrated security=true;data source=SONEXAY7;Password =sql;User ID =sa")
            cn.Open()
            cmd = New SqlCommand
            cmd.Connection = cn
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = Apostrophe("Update_List")
            'With cmd.Parameters
            '    .AddWithValue("@empid", Me.txtEmpNo)
            'End With

            Dim dr As SqlDataReader = cmd.ExecuteReader
            If dr.Read Then
                'txtEmpNo.Text = dr.Item(0)
                'TxtEmpName.Text = dr.Item(1)
                MessageBox.Show("Data Searched")
            Else
                MessageBox.Show("Data Not Available")

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    Private Sub x()
        Dim X1 As Double
        X1 = 0
        ' This sub appears to be using RSCC4 which was removed, commenting out for safety
        ' With RSC
        '     Call LoadSqlData("select *  from Ap_Loan where Bnk_Ac_Code3='" & (RSCC4.Fields("Bnk_Ac_Code3").Value.ToString) & "'", RSC)
        '     If .RecordCount <> 0 Then
        '         While Not .EOF()
        '             X1 = X1 + 1
        '             DbHelper.ExecuteNonQuery(" Update Ap_Loan set Bnk_Ac_Code=Bnk_Ac_Code3 + '-' + '" & X1 & "' where cnt = '" & (.Fields("cnt").Value.ToString) & "' ")
        '             '(.Fields("Withdr_No").Value.ToString) & _

        '             .MoveNext()
        '         End While
        '     End If
        ' End With
        'MsgBox("ok")
    End Sub
    Private Sub Label30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

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

    Private Sub VS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VS.Click

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class