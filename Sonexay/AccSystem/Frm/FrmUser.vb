Public Class FrmUser
    Public RSC As New ADODB.Recordset
    Public EditActive As Boolean
    Dim itemfgacc As Boolean
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim MDSection As Integer = 0
    Dim MDCheckWrite, MDCheckEdit, MDCheckDelete, MDCheckForstaff As Integer
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub FrmUser_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Call MdiCNum()
        SUPD = 0
    End Sub
    Private Sub FrmUser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SUPD = 0
        Call SetupGrid()
        Call LoadData()
        Call LoadSection()
        txtUsr_id.Enabled = True
        txtUsr_id.Focus()
        FgSec.Size = New System.Drawing.Size(392, 173)
     
        Call loadCompany()
        Call LoadSubCompany()
    End Sub

    Private Sub SetupGrid()
        With Fg
            .Columns.Clear()
            .Columns.Add("No", "ລ/ດ")
            .Columns.Add("UserID", "ລະຫັດຜູ້ໃຊ້")
            .Columns.Add("UserName", "ລາຍການຜູ້ໃຊ້")
            .Columns.Add("Permission", "ສິດໃຊ້ໂປຣແກຣມ")
            .Columns.Add("Section", "ພາກສ່ວນ")
            .Columns.Add("Branch", "ສາຂາ")

            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .ReadOnly = True
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .MultiSelect = False
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        End With

        With FgSec
            .Columns.Clear()
            .Columns.Add("No", "ລ/ດ")
            .Columns.Add("SecID", "ລະຫັດພາກສ່ວນ")
            .Columns.Add("SecName", "ລາຍການພາກສ່ວນ")

            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .ReadOnly = True
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .MultiSelect = False
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        End With

        With FgItem
            .Columns.Clear()
            .Columns.Add("No", "ລ/ດ")
            Dim chkCol As New DataGridViewCheckBoxColumn()
            chkCol.Name = "Ints"
            chkCol.HeaderText = "Check"
            .Columns.Add(chkCol)
            .Columns.Add("SecID", "Sec_ID")
            .Columns.Add("SecName", "Sec_Nm")

            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .MultiSelect = False
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            
            .ReadOnly = False
            .Columns(0).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = True
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
        'cmbCompany.Text = "08 ໄຊຍະບູລີ"
        SUPD = 0
    End Sub
    Private Sub LoadData()
        If MPermit = "Admin" Then
            Sql = ""
        Else
            Sql = "AND Company='" & MuSubOff & "'"
        End If
        Fg.Rows.Clear()
        With rs
            Call LoadSqlData("select AP_Users.Usr_id,AP_Users.Usr_nm,AP_Users.permision,Ap_Section.Sec_Nm,AP_Users.Company from AP_Users" & _
                        " INNER JOIN Ap_Section ON AP_Users.Sec_ID=Ap_Section.Sec_ID " & _
                        " WHERE 1=1 " & Sql & "  order by Usr_id", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    Fg.Rows.Add(.AbsolutePosition, _
                    .Fields("Usr_id").Value.ToString, _
                    .Fields("Usr_nm").Value.ToString, _
                    .Fields("permision").Value.ToString, _
                    .Fields("Sec_Nm").Value.ToString, _
                    .Fields("Company").Value.ToString)
                    .MoveNext()
                End While
            End If
        End With
    End Sub
    Private Sub LoadSection()
        FgSec.Rows.Clear()
        With rs
            Call LoadSqlData("select * from Ap_Section" & _
                        " WHERE Sec_ID > 0  order by Sec_ID", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    FgSec.Rows.Add(.AbsolutePosition, _
                    .Fields("Sec_ID").Value.ToString, _
                    .Fields("Sec_Nm").Value.ToString)
                    .MoveNext()
                End While
            End If
        End With
    End Sub
    Private Sub ClearText()
        txtUsr_id.Text = ""
        txtUsr_nm.Text = ""
        txtPWD.Text = ""
        txtConfrim.Text = ""
        txtDep_ID.Text = "0"
        txtDep_Nm.Text = "Administrator"
        cmbpermision.Text = "Admin"
        cmbUsrPermit.Text = "Administrator"
        'cmbCompany.SelectedIndex = 0
        txtUsr_id.Focus()
        SUPD = 0
    End Sub

    Private Sub cmbpermision_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbpermision.KeyPress
        If e.KeyChar = Chr(13) Then
            txtDep_ID.Focus()
        End If
    End Sub
    Private Sub cmbpermision_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbpermision.SelectedIndexChanged
        If cmbpermision.SelectedIndex = 0 Then
            Panel1.Visible = False
            Panel2.Visible = False
            Panel4.Visible = False
            Button1.Enabled = False
            txtDep_ID.Enabled = False
            txtDep_Nm.Enabled = False
            txtDep_ID.Text = "0"
            txtDep_Nm.Text = "Administrator"
            'ElseIf cmbpermision.SelectedIndex = 1 Then
            '    Button1.Enabled = False
            '    txtDep_ID.Enabled = False
            '    txtDep_Nm.Enabled = False
            '    txtDep_ID.Text = "0"
            '    txtDep_Nm.Text = "Administrator"
        Else
            Panel1.Visible = True
            Panel2.Visible = True
            Panel4.Visible = True
            Button1.Enabled = True
            txtDep_ID.Enabled = True
            txtDep_Nm.Enabled = True
            txtDep_ID.Text = ""
            txtDep_Nm.Text = ""
        End If
    End Sub

    Private Sub cmbUsrPermit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbUsrPermit.KeyPress
        If e.KeyChar = Chr(13) Then
            cmbpermision.Focus()
        End If
    End Sub
    Private Sub cmbUsrPermit_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbUsrPermit.SelectedIndexChanged
        'If cmbpermision.Text = "Administrator" Then
        '    CheckWrite_bit.Enabled = False
        '    CheckEdit_bit.Enabled = False
        '    CheckDelete_bit.Enabled = False
        '    CheckWrite_bit.Checked = False
        '    CheckEdit_bit.Checked = False
        '    CheckDelete_bit.Checked = False
        'End If
        'If cmbpermision.Text = "User" Then
        '    CheckWrite_bit.Enabled = True
        '    CheckEdit_bit.Enabled = True
        '    CheckDelete_bit.Enabled = True
        '    CheckWrite_bit.Checked = True
        '    CheckEdit_bit.Checked = True
        '    CheckDelete_bit.Checked = True
        'End If
    End Sub
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew.Click
        Call ClearText()
        txtUsr_id.Enabled = True
    End Sub
    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        If txtUsr_id.Text = "" Then MsgBox("ກະລຸນາປ້ອນລະຫັດ , ກ່ອນ!", MsgBoxStyle.OkOnly) : txtUsr_id.Focus() : Exit Sub
        If txtDep_ID.Text = "" Then MsgBox("ກະລຸນາປ້ອນລາຍການພາກສ່ວນ , ກ່ອນ!", MsgBoxStyle.OkOnly) : txtDep_ID.Focus() : Exit Sub
        If txtUsr_id.Enabled = True Then
            Call LoadSqlData("SELECT Usr_id FROM AP_Users WHERE Usr_id = '" & Trim(txtUsr_id.Text) & "' AND Company='" & MuSubOff & "' ", RSC)
            If RSC.RecordCount > 0 Then
                MsgBox("ລະຫັດຜູ້ໃຊ້ : " & Trim(txtUsr_id.Text) & " ມີແລ້ວ, ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
                txtUsr_id.Focus()
                If RSC.State = ConnectionState.Open Then RSC.Close()
                Exit Sub
            End If
            If RSC.State = ConnectionState.Open Then RSC.Close()
        End If
        Call Save()
        MsgBox("ບັນທຶກສໍາເລັດຜົນ!", MsgBoxStyle.OkOnly)
        Call LoadData()
        SUPD = 0
    End Sub
    Private Sub Save()
        Call LoadSqlData("SELECT * FROM AP_Users WHERE Usr_id = '" & txtUsr_id.Text & "'  ", rs)
        With rs
            If .RecordCount = 0 Then
                CNN.Execute("INSERT INTO AP_Users (Usr_id,Sec_ID,Usr_nm,permision,UsrPermit, " & _
                "PWD,Company ,Sub_Company,lst_updt,lst_usr,pc_nm) " & _
                   " VALUES('" & (txtUsr_id.Text) & "'," & _
                   " N'" & (txtDep_ID.Text) & "'," & _
                   " N'" & (txtUsr_nm.Text) & "'," & _
                   " N'" & (cmbpermision.Text) & "'," & _
                   " N'" & (cmbUsrPermit.Text) & "'," & _
                   " N'" & (txtPWD.Text) & "'," & _
                   " N'" & cmbCompany.Text & "'," & _
                      " N'" & Sub_Company.Text & "'," & _
                              " Getdate()," & _
                   " N'" & MUserName & "'," & _
                   " N'" & MDServerName & "')")

                '=======insertImage

                CNN.Execute("insert into Ap_Image (Img_Id,ImgType ,Img)  select  N'" & txtUsr_id.Text & "',ImgType ,Img  from Ap_Image2 where Img_Id = 'a' And ImgType = 'User'")
                CNN.Execute("insert into Ap_Image (Img_Id,ImgType ,Img)  select  N'" & txtUsr_id.Text & "',ImgType ,Img  from Ap_Image2 where Img_Id = 'a' And ImgType = 'Back'")
                'If SUPD = 1 Then
                '    Fm_Image.Img_ID.Text = txtUsr_id.Text
                '    Fm_Image.ImgType.Text = "User"
                '    Insert_Image()
                '    Fm_Image.Img_ID.Text = txtUsr_id.Text
                '    Fm_Image.ImgType.Text = "Back"
                '    Insert_Image()
                'End If

                '================

            Else
                'Conn.Execute("INSERT INTO AP_Users_ED ( Usr_id, Usr_nm, permision, Sec_id, UsrPermit, Write_bit, Edit_bit, Delete_bit, PWD, lst_usr, lst_updt, pc_nm) " & _
                '" Select Usr_id, Usr_nm, permision, Sec_id, UsrPermit, Write_bit, Edit_bit, Delete_bit, PWD, lst_usr, lst_updt, pc_nm From AP_Users WHERE Usr_id='" & Trim(Me.txtUsr_id.Text) & "'")
                CNN.Execute("UPDATE AP_Users SET " & _
                   " Usr_nm=N'" & txtUsr_nm.Text & "'," & _
                   " Sec_ID=N'" & (txtDep_ID.Text) & "'," & _
                   " permision=N'" & (cmbpermision.Text) & "'," & _
                   " UsrPermit=N'" & (cmbUsrPermit.Text) & "'," & _
                   " PWD=N'" & txtPWD.Text & "'," & _
                   " lst_updt=Getdate()," & _
                   " lst_usr=N'" & MUserName & "'," & _
                   " Company=N'" & cmbCompany.Text & "'," & _
                     " Sub_Company=N'" & Sub_Company.Text & "'," & _
                   " pc_nm=N'" & MDServerName & "' " & _
                   "WHERE Usr_id=N'" & (txtUsr_id.Text) & "' ")

                '=======updateImage
                'If SUPD = 1 Then
                '    Fm_Image.Img_ID.Text = txtUsr_id.Text
                '    Fm_Image.ImgType.Text = "User"
                '    deleteImage()
                '    Insert_Image()
                'End If

                '=======
            End If
        End With

   



    End Sub
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDel.Click
        If Fg.Rows.Count = 0 Then MsgBox("ທ່ານບໍ່ສາມາດລຶບລາຍການນີ້ໄດ້, ເພາະແມ່ນລາຍການສຸດທ້າຍແລ້ວ ", MsgBoxStyle.OkOnly) : Exit Sub
        Dim Rsch As New ADODB.Recordset
        If MessageBox.Show("ທ່ານຕ້ອງການລຶບລາຍການ '" & (txtUsr_id.Text) & "' ແມ່ນ ຫຼື ບໍ່ ?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            CNN.BeginTrans()
            CNN.Execute("delete from AP_Users where Usr_id='" & Trim(txtUsr_id.Text) & "'  ")
            '========deleteImage======
            'Fm_Image.Img_ID.Text = txtUsr_id.Text
            'Fm_Image.ImgType.Text = "User"
            'deleteImage()
            'Fm_Image.Img_ID.Text = txtUsr_id.Text
            'Fm_Image.ImgType.Text = "Back"
            'deleteImage()
            '===============
            Call ClearText()
            Call LoadData()
            CNN.CommitTrans()
        End If
        SUPD = 0
    End Sub


    Private Sub Fg_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Fg.SelectionChanged
        If Fg.CurrentRow Is Nothing Then Exit Sub
        If Fg.CurrentRow.Index < 0 Then Exit Sub
        
        Dim usrId As String = ""
        If Fg.CurrentRow.Cells(1).Value IsNot Nothing Then
             usrId = Fg.CurrentRow.Cells(1).Value.ToString()
        End If

        If usrId = "" Then Exit Sub

        txtUsr_id.Text = usrId
        If Fg.CurrentRow.Cells(5).Value IsNot Nothing Then
             cmbCompany.Text = Fg.CurrentRow.Cells(5).Value.ToString()
        End If

        txtUsr_id.Enabled = False
        Dim FGSel As New ADODB.Recordset
        With FGSel
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .Open("Select AP_Users.*,Ap_Section.Sec_Nm From AP_Users" & _
                  " INNER JOIN Ap_Section ON AP_Users.Sec_id=Ap_Section.Sec_id Where(AP_Users.Usr_id='" & _
             usrId & "' AND Company='" & MuSubOff & "' )", CNN, _
             ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .RecordCount <> 0 Then
                cmbCompany.Text = (.Fields("Company").Value.ToString)
                Sub_Company.Text = (.Fields("Sub_Company").Value.ToString)
                txtOldPass.Text = (.Fields("PWD").Value.ToString)
                txtNewPass.Text = (.Fields("PWD").Value.ToString)
                cmbpermision.Text = (.Fields("permision").Value.ToString)
                txtUsr_id.Text = (.Fields("Usr_id").Value.ToString)
                txtUsr_nm.Text = (.Fields("Usr_nm").Value.ToString)
                txtDep_ID.Text = (.Fields("Sec_id").Value)
                txtDep_Nm.Text = (.Fields("Sec_Nm").Value.ToString)
                cmbUsrPermit.Text = (.Fields("UsrPermit").Value.ToString)
                txtPWD.Text = (.Fields("PWD").Value.ToString)
                txtConfrim.Text = (.Fields("PWD").Value.ToString)
                MDCheckForstaff = (.Fields("ForStaff").Value)
                MDCheckWrite = (.Fields("Write_bit").Value)
                MDCheckEdit = (.Fields("Edit_bit").Value)
                MDCheckDelete = (.Fields("Delete_bit").Value)
            End If
        End With

        If cmbpermision.Text = "Admin" Then
            Panel1.Visible = False
            Panel2.Visible = False
            Panel4.Visible = False
            Button1.Enabled = False
            txtDep_ID.Enabled = False
            txtDep_Nm.Enabled = False
            txtDep_ID.Text = "0"
            txtDep_Nm.Text = "Administrator"
        Else
            Panel1.Visible = True
            Panel2.Visible = True
            Panel4.Visible = True
            Button1.Enabled = True
            txtDep_ID.Enabled = True
            txtDep_Nm.Enabled = True
            txtDep_ID.Text = ""
            txtDep_Nm.Text = ""
        End If
        SUPD = 0
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Panel1.Visible = True
    End Sub

    Private Sub txtDep_ID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDep_ID.KeyPress
        If e.KeyChar = Chr(13) Then
            Call LoadSqlData("select *  from AP_Section WHERE Sec_ID='" & txtDep_ID.Text & "' ", rs)
            If rs.RecordCount = 0 Then
                Button1_Click(sender, e)
            Else
                txtDep_Nm.Text = rs.Fields("Sec_Nm").Value.ToString
            End If
        End If
    End Sub
    Private Sub txtDep_ID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDep_ID.TextChanged
        'Call loadrs("select *  from AP_Sections WHERE Sec_id='" & txtDep_ID.Text & "' ", rs)
        'If rs.RecordCount > 0 Then
        '    txtDep_Nm.Text = rs.Fields("Sec_nmL").Value.ToString
        'End If
    End Sub

    Private Sub txtUsr_id_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtUsr_id.KeyPress
        If e.KeyChar = Chr(13) Then
            txtUsr_nm.Focus()
        End If
    End Sub

    Private Sub txtUsr_id_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUsr_id.TextChanged

    End Sub

    Private Sub txtUsr_nm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtUsr_nm.KeyPress
        If e.KeyChar = Chr(13) Then
            txtPWD.Focus()
        End If
    End Sub

    Private Sub txtUsr_nm_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUsr_nm.TextChanged

    End Sub

    Private Sub txtPWD_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPWD.KeyPress
        If e.KeyChar = Chr(13) Then
            txtConfrim.Focus()
        End If
    End Sub

    Private Sub txtPWD_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPWD.TextChanged

    End Sub

    Private Sub txtConfrim_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtConfrim.KeyPress
        If e.KeyChar = Chr(13) Then
            cmbUsrPermit.Focus()
        End If
    End Sub

    Private Sub txtOldCost_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Panel1.Visible = False
        Exit Sub
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If txtUsr_id.Text = "" Then MsgBox("ກະລຸນາເລືອກລາຍການກ່ອນ", MsgBoxStyle.OkOnly) : Exit Sub
        Panel3.Visible = True
        SUPD = 0
    End Sub

    Private Sub FgSec_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FgSec.SelectionChanged
        If FgSec.CurrentRow Is Nothing Then Exit Sub
        If FgSec.CurrentRow.Index < 0 Then Exit Sub
        
        Dim secId As String = ""
        If FgSec.CurrentRow.Cells(1).Value IsNot Nothing Then
             secId = FgSec.CurrentRow.Cells(1).Value.ToString()
        End If

        If secId = "" Then Exit Sub
        
        txtDep_ID.Text = secId
        If FgSec.CurrentRow.Cells(2).Value IsNot Nothing Then
            txtDep_Nm.Text = FgSec.CurrentRow.Cells(2).Value.ToString()
        End If

        FgItem.Rows.Clear()
        With rs
            Call LoadSqlData("select * from Ap_Section_Item" & _
                        " WHERE Sec_ID='" & secId & "'  order by Sec_ID", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    Dim checkVal As Boolean = False
                    Try
                        If Not IsDBNull(.Fields("Ints").Value) Then
                             checkVal = CBool(.Fields("Ints").Value)
                        End If
                    Catch ex As Exception
                        checkVal = False
                    End Try

                    FgItem.Rows.Add(.AbsolutePosition, _
                    checkVal, _
                    .Fields("Sec_ID").Value.ToString, _
                    .Fields("Sec_Nm").Value.ToString)
                    .MoveNext()
                End While
            End If
        End With
        Panel4.Visible = True
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Panel3.Visible = False
        SUPD = 0
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If txtConfrimPass.Text = "" Then MsgBox("ກະລຸນາຢືນຢັນລະຫັດຜ່ານຂອງທ່ານ", MsgBoxStyle.OkOnly) : txtConfrimPass.Focus() : Exit Sub
        If txtConfrimPass.Text <> txtNewPass.Text Then MsgBox("ລະຫັດຜ່ານຂອງທ່ານບໍ່ຖືກຕ້ອງມກະລຸນາປ່ຽນ", MsgBoxStyle.OkOnly) : txtConfrimPass.Text = "" : Exit Sub
        CNN.Execute("UPDATE AP_Users SET " & _
                " PWD=N'" & txtNewPass.Text & "'," & _
                " lst_updt=Getdate()," & _
                " lst_usr=N'" & MDServerUser & "'," & _
                " pc_nm=N'" & MDServerName & "' " & _
                " WHERE Usr_id=N'" & (txtUsr_id.Text) & "' AND Company='" & MuSubOff & "'")
        MsgBox("ບັນທຶກສໍາເລັດຜົນ!", MsgBoxStyle.OkOnly)
        Call LoadData()
        SUPD = 0
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Panel4.Visible = False
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button9_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim i As Integer
        For i = 0 To FgItem.Rows.Count - 1
            Dim checkVal As Boolean = False
            If FgItem.Rows(i).Cells(1).Value IsNot Nothing Then
                checkVal = CBool(FgItem.Rows(i).Cells(1).Value)
            End If
            
            Dim secID As String = FgItem.Rows(i).Cells(2).Value.ToString()
            Dim secName As String = FgItem.Rows(i).Cells(3).Value.ToString()
            
            Dim ints As Integer = 0
            If checkVal Then ints = 1
            
            CNN.Execute("UPDATE Ap_Section_Item SET Ints=" & ints & " WHERE Sec_ID='" & secID & "' AND Sec_Nm=N'" & secName & "' ")
        Next i
        MsgBox("ສໍາເລັດຜົນ!", MsgBoxStyle.OkOnly)
    End Sub

    Private Sub FgItem_CellContentClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FgItem.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex = 1 Then ' Checkbox column
             FgItem.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub FgItem_CellValueChanged(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FgItem.CellValueChanged
         If e.RowIndex < 0 Then Exit Sub
         If e.ColumnIndex = 1 Then
            Dim checkVal As Boolean = False
            If FgItem.Rows(e.RowIndex).Cells(1).Value IsNot Nothing Then
                checkVal = CBool(FgItem.Rows(e.RowIndex).Cells(1).Value)
            End If
            
            Dim secID As String = FgItem.Rows(e.RowIndex).Cells(2).Value.ToString()
            Dim secName As String = FgItem.Rows(e.RowIndex).Cells(3).Value.ToString()
            
            Dim ints As Integer = 0
            If checkVal Then ints = 1
             
             CNN.Execute("UPDATE Ap_Section_Item SET Ints=" & ints & " WHERE Sec_ID='" & secID & "' AND Sec_Nm=N'" & secName & "' ")
         End If
    End Sub

  

    Private Sub Button9_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Fm_Image.Img_ID.Text = txtUsr_id.Text
        Fm_Image.ImgType.Text = "User"
        Update_Image()
    End Sub

    Private Sub cmbCompany_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCompany.SelectedIndexChanged
        LoadSubCompany()
    End Sub
    Private Sub LoadSubCompany()
        Sub_Company.Items.Clear()
        LoadSqlData("select sub_id , off_id , off_add2  from  Ap_office where off_id ='08' group BY  sub_id  ,off_id , off_add2", RSC)

        With RSC
            Do Until .EOF = True
                Sub_Company.Items.Add((.Fields("sub_id").Value) & " " & (.Fields("off_add2").Value))
                .MoveNext()
            Loop
        End With

        SUPD = 0
    End Sub

    Private Sub Sub_Company_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Sub_Company.SelectedIndexChanged

    End Sub

    Private Sub Button6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        FmOff_User.ShowDialog()
    End Sub
End Class