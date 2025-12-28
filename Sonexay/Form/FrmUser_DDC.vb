Public Class FrmUser_DDC

    Public RSC As New ADODB.Recordset
    Public EditActive As Boolean
    Dim itemfgacc As Boolean
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim MDSection As Integer = 0
    Dim MDCheckWrite, MDCheckEdit, MDCheckDelete, MDCheckForstaff, DP_ID As Integer
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub FrmUser_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        'Call MdiCNum()
        SUPD = 0
    End Sub
    Private Sub FrmUser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Fg.FormatString = "^ລ/ດ|<ລະຫັດຜູ້ໃຊ້   |<ຊື່ ຜູ້ໃຊ້               |<ພາກ​ສ່ວນ/ສຳ​ນັກ​ງານ              |<ຮັບຜິດຊອບ  |<ໃນນາມ"

        'If MDWrite = 0 Then
        '    BtnAddNew.Enabled = False
        '    BtnSave.Enabled = False
        'Else
        '    BtnAddNew.Enabled = True
        '    BtnSave.Enabled = True
        'End If
        'If MDDelete = 0 Then
        '    BtnDel.Enabled = False
        'Else
        '    BtnDel.Enabled = True
        'End If

        FgSec.Visible = True
        FgItem.Visible = True
        SUPD = 0

        Call LoadData()
        Call LoadSection()
        txtUsr_id.Enabled = True
        txtUsr_id.Focus()


        'SetControlText(Me)
        Fg.Cols = 6
        FgSec.Cols = 3
        FgItem.Cols = 4
        FG2.Cols = 3
        FG2.set_ColDataType(1, VSFlex8U.DataTypeSettings.flexDTBoolean)
        FgItem.set_ColDataType(1, VSFlex8U.DataTypeSettings.flexDTBoolean)
        If FG2.Col = 1 Then
            FG2.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
        End If
        If FgItem.Col = 1 Then
            FgItem.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
        End If

        FgItem.FormatString = "^ລ/ດ|    |< |<                     "
        FgSec.FormatString = "^ລ/ດ||<ລະຫັດພາກສ່ວນ  |<ລາຍການພາກສ່ວນ               "
        FG2.FormatString = "^ລ/ດ|    |^User   |<            "
        FG_DP.FormatString = "^ລ/ດ|<ເລືອກ|<ລະຫັດ   |<ກົມກອງ               "

        FgSec.set_ColHidden(1, True)
        FgSec.set_ColHidden(3, True)
        FgSec.BackColorSel = Color.SkyBlue
        FgItem.set_ColHidden(2, True)
        FG2.Rows = 1
        FG2.Rows = 2
        Fg.Size = New System.Drawing.Size(557, 521)
        FgSec.Size = New System.Drawing.Size(411, 200)
        FgItem.Size = New System.Drawing.Size(411, 271)
        FG2.Size = New System.Drawing.Size(265, 480)
        'Call LoadSubCompany()
        EditActive = False
        'Fg.set_ColHidden(5, False)
    End Sub
    Public Sub LoadListFG2()
        Dim aa As String
        aa = "update Ap_Office set Lck=0 update Ap_Office set Lck=1 from  Ap_Office_User , Ap_office where Ap_Office_User.Sub_Id = Ap_office.Sub_Id And Usr_Id =  '" & txtUsr_id.Text & "' "
        CNN.Execute(aa)
        FG2.Rows = 1
        With RSC
            Call LoadSqlData("select Sub_Id , Off_Add2,Lck  from Ap_office order by Sub_Id", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG2.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("Lck").Value)) & _
                                "" & vbTab & Trim(CStr(.Fields("Sub_Id").Value)) & _
                    "" & vbTab & Trim(CStr(.Fields("Off_Add2").Value)))
                    .MoveNext()
                End While
            Else
                FG2.Rows = 2
            End If
        End With
        'If txtUsr_id.Text = "a" Then
        '    Exit Sub
        'End If
        For i = 1 To FG2.Rows - 1
            If FG2.get_TextMatrix(i, 2) = "00-00" Then
                FG2.Col = 2
                FG2.Row = i
                FG2.CellForeColor = Color.Red
                FG2.CellFontBold = True
                FG2.Col = 3
                FG2.CellForeColor = Color.Red
                FG2.CellFontBold = True
            End If
            If Mid(FG2.get_TextMatrix(i, 2), 1, 2) <> "00" Then
                If Mid(FG2.get_TextMatrix(i, 2), 4, 2) = "00" Then
                    FG2.Col = 2
                    FG2.Row = i
                    FG2.CellForeColor = Color.Blue
                    FG2.CellFontBold = True
                    FG2.Col = 3
                    FG2.CellForeColor = Color.Blue
                    FG2.CellFontBold = True
                End If
            End If
        Next i
    End Sub
    Private Sub LoadData()
        'Fg.FormatString = "^ລ/ດ|<ລະຫັດຜູ້ໃຊ້   |<ຊື່ ຜູ້ໃຊ້              |<ພາກ​ສ່ວນ/ສຳ​ນັກ​ງານ         |<ຮັບຜິດຊອບ         |<ໃນນາມ"
        Fg.Rows = 1
        With rs
            Call LoadSqlData("select Usr_id,Usr_nm,permision   from Ap_Users Order by Sec_id ", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    Fg.AddItem(.AbsolutePosition & _
                    Chr(9) & (.Fields("Usr_id").Value.ToString) & _
                    Chr(9) & (.Fields("Usr_nm").Value.ToString) & _
                       Chr(9) & "" & _
                     Chr(9) & "" & _
                    Chr(9) & (.Fields("permision").Value.ToString))
                    .MoveNext()
                End While
            End If
        End With
    End Sub
    Private Sub LoadSection()
        FgSec.Rows = 1
        With rs
            Call LoadSqlData("select * from Ap_Section Order by Sec_ID ", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    FgSec.AddItem(.AbsolutePosition & _
                    Chr(9) & (.Fields("Sec_ID").Value.ToString) & _
                    Chr(9) & (.Fields("Sec_Nm").Value.ToString))
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

        txt_Component_id.Text = ""
        txt_Component_nm.Text = ""

        Panel3.Visible = False
        cmbpermision.Text = "Admin"
        cmbUsrPermit.Text = "Administrator"
        'cmbCompany.SelectedIndex = 0
        txtUsr_id.Focus()
        SUPD = 0
    End Sub




    Private Sub cmbUsrPermit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbUsrPermit.KeyPress
        If e.KeyChar = Chr(13) Then
            cmbpermision.Focus()
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew.Click
        Call ClearText()
        txtUsr_id.Enabled = True
        EditActive = False
    End Sub
    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        If txtUsr_id.Text = "" Then MsgBox("ກະລຸນາປ້ອນລະຫັດ , ກ່ອນ!", MsgBoxStyle.OkOnly) : txtUsr_id.Focus() : Exit Sub
        'If txt_Component_id.Text = "" Then MsgBox("ກະລຸນາເລືອກ ບໍລິສັດ , ກ່ອນ!", MsgBoxStyle.OkOnly) : txt_Component_id.Focus() : Exit Sub

        If txtUsr_id.Enabled = True Then
            Call LoadSqlData("SELECT Usr_id FROM Ap_Users WHERE Usr_id = N'" & Trim(txtUsr_id.Text) & "' ", RSC)
            If RSC.RecordCount > 0 Then
                MsgBox("ລະຫັດຜູ້ໃຊ້ : " & Trim(txtUsr_id.Text) & " ມີແລ້ວ, ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
                txtUsr_id.Focus()
                If RSC.State = ConnectionState.Open Then RSC.Close()
                Exit Sub
            End If
            If RSC.State = ConnectionState.Open Then RSC.Close()
        End If
        Call Save()

        If txtUsr_id.Enabled = True Then
            Call InsertMenuStrip_Usr()
        End If

        Call LoadData()
        SUPD = 0
        CNN.Execute("delete from  AP_Users_List WHERE  Usr_id = N'" & txtUsr_id.Text & "'")
        Call Save_Item()

        MsgBox("ບັນທຶກສໍາເລັດຜົນ!", MsgBoxStyle.OkOnly)
    End Sub
    Public Sub InsertMenuStrip_Usr()
        Usr = txtUsr_id.Text
        CNN.Execute(" Delete Ap_Section   Delete Ap_Section_AdNew  ")

        Dim N As Integer = 0
        For i = 0 To FmMain.MenuStrip1.Items.Count - 1
            N = N + 1
            If FmMain.MenuStrip1.Items.Item(i).Text <> "" Then
                CNN.Execute("Insert Into Ap_Section (Sec_ID , Sec_Nm , Menu_Nm ) VAlues (" & N & " ,N'" & FmMain.MenuStrip1.Items.Item(i).Text & "'   ,N'" & FmMain.MenuStrip1.Items.Item(i).Name & "' )")
            End If
        Next i
        'LoadUsr()
        'MsgBox(Usr)
        CNN.Execute("Update Ap_Section_AdNew set Ints=Ap_Section_Item.Ints from Ap_Section_AdNew , Ap_Section_Item where Ap_Section_AdNew.Menu_Nm = Ap_Section_Item.Menu_Nm And Ap_Section_AdNew.Usr_Id = Ap_Section_Item.Usr_Id ")
        CNN.Execute("delete Ap_Section_Item where Usr_Id='" & Usr & "'")
        CNN.Execute("insert  into Ap_Section_Item (Sec_ID, Sec_Nm, Ints, Menu_Nm , Usr_Id) select Sec_ID, Sec_Nm, Ints, Menu_Nm , Usr_Id from Ap_Section_AdNew Order by cnt ")
        CNN.Execute("Update Ap_Section_Item set Ints = 1 where Usr_Id='" & Usr & "'")
    End Sub
     
    Private Sub Save_Item()
        Dim Rschk As New ADODB.Recordset
        Dim i As Integer
        With Rschk
            If cmbpermision.SelectedIndex = 1 Then
                Call LoadSqlData("SELECT * FROM AP_Users_List WHERE Usr_id = N'" & txtUsr_id.Text & "'", Rschk)
                For i = 1 To FG_DP.Rows - 1
                    If .RecordCount = 0 Then
                        If FG_DP.get_ValueMatrix(i, 1) = True Then
                            CNN.Execute("INSERT INTO  AP_Users_List (Usr_id,Stff_Id,Usr_nm,permision,permision_id,status,txt_DP_ID,Cmb_DP,Get_date, lst_usr, pc_nm) " & _
                                " VALUES(N'" & (txtUsr_id.Text) & "'," & _
                                      " N'" & (TextBox1.Text) & "'," & _
                           " N'" & (txtUsr_nm.Text) & "'," & _
                               " N'" & (cmbpermision.Text) & "'," & _
                           " N'" & (txt_permision_id.Text) & "'," & _
                             " N'" & 1 & "'," & _
                                " N'" & Apostrophe(FG_DP.get_TextMatrix(i, 2)) & "'," & _
                                  " N'" & Apostrophe(FG_DP.get_TextMatrix(i, 3)) & "'," & _
                                   " Getdate()," & _
                              " N'" & MUserName & "'," & _
                              " '" & COMPUTER_NM & "')")
                        End If
                    End If
                Next i

            ElseIf cmbpermision.SelectedIndex = 2 Then

                Call LoadSqlData("SELECT * FROM AP_Users_List WHERE Usr_id = N'" & txtUsr_id.Text & "'", Rschk)
                For i = 1 To FG_DP.Rows - 1
                    If .RecordCount = 0 Then
                        If FG_DP.get_ValueMatrix(i, 1) = True Then
                            CNN.Execute("INSERT INTO  AP_Users_List (Usr_id,Stff_Id,Usr_nm,permision,permision_id,txt_DP_ID,Cmb_DP,txt_DP_sub_ID,Cmb_DP_sub,Get_date, lst_usr, pc_nm) " & _
                                " VALUES(N'" & (txtUsr_id.Text) & "'," & _
                                      " N'" & (TextBox1.Text) & "'," & _
                           " N'" & (txtUsr_nm.Text) & "'," & _
                               " N'" & (cmbpermision.Text) & "'," & _
                           " N'" & (txt_permision_id.Text) & "'," & _
                               " N'" & (txt_DP_ID.Text) & "'," & _
                                   " N'" & (Cmb_DP.Text) & "'," & _
                                " N'" & Apostrophe(FG_DP.get_TextMatrix(i, 2)) & "'," & _
                                  " N'" & Apostrophe(FG_DP.get_TextMatrix(i, 3)) & "'," & _
                                   " Getdate()," & _
                              " N'" & MUserName & "'," & _
                              " '" & COMPUTER_NM & "')")
                        End If
                    End If
                Next i
            Else
                Call LoadSqlData("SELECT * FROM AP_Users_List WHERE Usr_id = N'" & txtUsr_id.Text & "'", Rschk)
                For i = 1 To FG_DP.Rows - 1
                    If .RecordCount = 0 Then
                        If FG_DP.get_ValueMatrix(i, 1) = True Then
                            CNN.Execute("INSERT INTO  AP_Users_List (Usr_id,Stff_Id,Usr_nm,permision,permision_id,txt_DP_ID,Cmb_DP,Get_date, lst_usr, pc_nm) " & _
                                " VALUES(N'" & (txtUsr_id.Text) & "'," & _
                                      " N'" & (TextBox1.Text) & "'," & _
                           " N'" & (txtUsr_nm.Text) & "'," & _
                               " N'" & (cmbpermision.Text) & "'," & _
                           " N'" & (txt_permision_id.Text) & "'," & _
                                " N'" & Apostrophe(FG_DP.get_TextMatrix(i, 2)) & "'," & _
                                  " N'" & Apostrophe(FG_DP.get_TextMatrix(i, 3)) & "'," & _
                                   " Getdate()," & _
                              " N'" & MUserName & "'," & _
                              " '" & COMPUTER_NM & "')")
                        End If
                    End If
                Next i

            End If
        End With


    End Sub
    Private Sub Save()
        If CheckBox6.Checked = False Then
            MDCheckWrite = 0
        Else
            MDCheckWrite = 1
        End If
        If CheckBox5.Checked = False Then
            MDCheckEdit = 0
        Else
            MDCheckEdit = 1
        End If
        If CheckBox4.Checked = False Then
            MDCheckDelete = 0
        Else
            MDCheckDelete = 1
        End If


        Dim rs As New ADODB.Recordset
        Call LoadSqlData("SELECT Usr_id  FROM Ap_Users WHERE Usr_id =N'" & Trim(txtUsr_id.Text) & "' ", rs)
        With rs
            If .RecordCount = 0 Then
                CNN.Execute("INSERT INTO Ap_Users (Usr_id,Usr_nm,permision,Sec_id,off_id,off_nm,Department, " & _
                "PWD ,lst_updt , Write_bit,Edit_bit,Delete_bit ) " & _
                   " VALUES( N'" & Trim(txtUsr_id.Text) & "'," & _
                   " N'" & Trim(txtUsr_nm.Text) & "'," & _
                   " N'" & (cmbpermision.Text) & "'," & _
                      " N'" & Trim(txt_permision_id.Text) & "'," & _
                     " N'" & Trim(txt_Component_id.Text) & "'," & _
                     " N'" & Trim(txt_Component_nm.Text) & "'," & _
                   " N'" & (TextBox1.Text) & "'," & _
                   " N'" & (txtPWD.Text) & "'," & _
                              " Getdate()," & _
                       " N'" & MDCheckWrite & "'," & _
                   " N'" & MDCheckEdit & "'," & _
                   " N'" & MDCheckDelete & "')")

                'CNN.Execute("insert into Ap_Image (Img_Id,ImgType ,Img)  select  N'" & txtUsr_id.Text & "',ImgType ,Img  from Ap_Image2 where Img_Id = 'a' And ImgType = 'User'")
                'CNN.Execute("insert into Ap_Image (Img_Id,ImgType ,Img)  select  N'" & txtUsr_id.Text & "',ImgType ,Img  from Ap_Image2 where Img_Id = 'a' And ImgType = 'Back'")
            Else
                CNN.Execute("UPDATE Ap_Users SET " & _
                   " Usr_nm=N'" & txtUsr_nm.Text & "'," & _
                       " off_id=N'" & Trim(txt_Component_id.Text) & "'," & _
                     " off_nm=N'" & Trim(txt_Component_nm.Text) & "'," & _
                   " permision=N'" & Trim(cmbpermision.Text) & "'," & _
                    " Sec_id=N'" & Trim(txt_permision_id.Text) & "'," & _
                   " Department=N'" & Trim(TextBox1.Text) & "'," & _
                   " PWD=N'" & txtPWD.Text & "'," & _
                       " Write_bit='" & MDCheckWrite & "'," & _
                   " Edit_bit='" & MDCheckEdit & "'," & _
                   " Delete_bit='" & MDCheckDelete & "'," & _
                   " lst_updt=Getdate()" & _
                   "WHERE Usr_id=N'" & Trim(txtUsr_id.Text) & "' ")

            End If
        End With

        Call LoadSqlData("SELECT Stff_Id FROM AP_Staffs WHERE Stff_Id =N'" & Trim(txtUsr_id.Text) & "'", rs)
        With rs
            If .RecordCount = 0 Then
                CNN.Execute("INSERT INTO AP_Staffs (Stff_Id,Stff_nmL, Lst_updt,Lst_usr,Pc_nm) " & _
                 " VALUES(N'" & Apostrophe(txtUsr_id.Text.Trim) & "'," & _
                 " N'" & Apostrophe(txtUsr_nm.Text.Trim) & "'," & _
                               " Getdate()," & _
                 " N'" & MUserName & "'," & _
                 " N'" & MDServerName & "')")
            Else
                CNN.Execute("  update  AP_Staffs set    " & _
               " Stff_nmL=N'" & Apostrophe(txtUsr_nm.Text.Trim) & "'," & _
                             " Lst_updt=Getdate()," & _
               " Lst_usr=N'" & MUserName & "', " & _
               " Pc_nm= N'" & MDServerName & "'  where  Stff_Id=N'" & Trim(txtUsr_id.Text) & "'  ")
            End If
        End With

        'LoadSqlData("Select Usr_id From Ap_Users Where Usr_id = '" & txtUsr_id.Text & "' ", RSC)
        If txtUsr_id.Text = "a" Then
            CNN.Execute("Delete Ap_Office_User Where Usr_Id =N'" & Trim(txtUsr_id.Text) & "'")
            For i = 1 To FG2.Rows - 1
                CNN.Execute("Insert into Ap_Office_User (Usr_Id,Sub_Id) VAlues ( N'" & Trim(txtUsr_id.Text) & "',N'" & Trim(FG2.get_TextMatrix(i, 2)) & "') ")
            Next i
            'CNN.Execute(" update Ap_Office_User set off_Id=Ap_Office.off_Id,off_add1=Ap_office.off_add1 ,off_add2=Ap_office.off_add2 from  Ap_Office_User , Ap_office where Ap_Office_User.Sub_Id = Ap_office.Sub_Id")
            CNN.Execute("Update Ap_Section_Item set Ints =1 Where Usr_Id = N'" & Trim(txtUsr_id.Text) & "'")
        End If

    End Sub
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDel.Click
        If txtUsr_id.Text = "a" Then MsgBox("a ບໍ່ສາມາມລືບໄດ້ຍ້ອນມັນເປັນລະຫັດຫລັກ , !", MsgBoxStyle.OkOnly) : txtUsr_id.Focus() : Exit Sub
        If Fg.Rows = 2 Then MsgBox("ທ່ານບໍ່ສາມາດລຶບລາຍການນີ້ໄດ້, ເພາະແມ່ນລາຍການສຸດທ້າຍແລ້ວ ", MsgBoxStyle.OkOnly) : Exit Sub
        AccCD = txtUsr_id.Text
        'Call LoadSqlData("Select Bill_no From AP_SaleForStock WHERE Stff_Id='" & AccCD & "'", RSC)
        'If RSC.RecordCount <> 0 Then MsgBox("You do not delete '" & AccCD & "' because activity", MsgBoxStyle.OkOnly) : RSC = Nothing : Exit Sub

        Dim Rsch As New ADODB.Recordset
        If MessageBox.Show("ທ່ານຕ້ອງການລຶບລາຍການ '" & (txtUsr_id.Text) & "' ແມ່ນ ຫຼື ບໍ່ ?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            CNN.Execute("delete from Ap_Users where Usr_id=N'" & Trim(txtUsr_id.Text) & "'  ")
            CNN.Execute("delete from AP_Users_List where Usr_id=N'" & Trim(txtUsr_id.Text) & "'  ")


            '========deleteImage======
            'Fm_Image.Img_ID.Text = txtUsr_id.Text
            'Fm_Image.ImgType.Text = "User"
            'deleteImage()
            'Fm_Image.Img_ID.Text = txtUsr_id.Text
            'Fm_Image.ImgType.Text = "Back"
            'deleteImage()
            '===============
            Call ClearText()
            Call InsertMenuStrip_Usr()
            Call LoadData()
        End If
        'SUPD = 0
    End Sub

    Private Sub Fg_ClickEvent(ByVal sender As Object, ByVal e As System.EventArgs) Handles Fg.ClickEvent

     

    End Sub

    Private Sub Fg_DblClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Fg.DblClick
        If Fg.get_TextMatrix(Fg.Row, 1) = "" Then Exit Sub
        txtUsr_id.Text = Fg.get_TextMatrix(Fg.Row, 1)
        txtOldPass.Text = Fg.get_TextMatrix(Fg.Row, 1)
        txtUsr_nm.Text = Fg.get_TextMatrix(Fg.Row, 2)
        LoadSqlData("Select Usr_id,Usr_nm,permision,Department , off_id,off_nm , PWD , Write_bit ,Edit_bit , Delete_bit  From  Ap_Users where  USr_Id = N'" & Trim(txtUsr_id.Text) & "' ", RSC)
        txtUsr_id.Enabled = False
        If RSC.RecordCount <> 0 Then
            cmbpermision.Text = (RSC.Fields("permision").Value.ToString)
            txtUsr_nm.Text = (RSC.Fields("Usr_nm").Value.ToString)

            txt_Component_id.Text = (RSC.Fields("off_id").Value.ToString)
            txt_Component_nm.Text = (RSC.Fields("off_nm").Value.ToString)

            TextBox1.Text = (RSC.Fields("Department").Value.ToString)
            txtPWD.Text = (RSC.Fields("PWD").Value.ToString)
            'MsgBox(RSC.Fields("Write_bit").Value)
            MDCheckWrite = (RSC.Fields("Write_bit").Value)
            MDCheckEdit = (RSC.Fields("Edit_bit").Value)
            MDCheckDelete = (RSC.Fields("Delete_bit").Value)


            CheckBox6.Checked = (RSC.Fields("Write_bit").Value)
            CheckBox5.Checked = (RSC.Fields("Edit_bit").Value)
            CheckBox4.Checked = (RSC.Fields("Delete_bit").Value)


        End If





        If txtUsr_id.Text = "a" Then
            'FG2.set_ColHidden(1, True)
            'FgItem.set_ColHidden(1, True)
            FG2.Editable = VSFlex8U.EditableSettings.flexEDNone
            FgItem.Editable = VSFlex8U.EditableSettings.flexEDNone
            'FG2.ForeColor = Color.Silver
            'FgItem.ForeColor = Color.Silver
            'FG2.Enabled = True
        Else
            'FG2.set_ColHidden(1, False)
            'FgItem.set_ColHidden(1, False)
            FG2.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
            FgItem.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
            'FG2.ForeColor = Color.Black
            'FgItem.ForeColor = Color.Black
        End If
        FgItem.Rows = 1
        FgItem.Rows = 2
        EditActive = True

        If cmbpermision.SelectedIndex = 0 Then
            Cmb_DP.Items.Clear()
            Cmb_DP.Text = ""
            Cmb_DP.Enabled = False
            txt_permision_id.Text = "0"
            txt_DP_ID.Text = "00"
            Load_DP_00()
            'For i = 1 To FG_DP.Rows - 1
            '    FG_DP.set_TextMatrix(i, 1, True)
            'Next

        ElseIf cmbpermision.SelectedIndex = 1 Then
            FG_DP.FormatString = "^ລ/ດ|<ເລືອກ|<ລະຫັດ   |<ສຳນັກງານ    "
            txt_permision_id.Text = "1"
            Cmb_DP.Items.Clear()
            Cmb_DP.Text = ""
            Cmb_DP.Enabled = False
            Load_DP()
            Load_DP_Edit()
            If FG_DP.Col = 1 Then
                FG_DP.Editable = VSFlex8U.EditableSettings.flexEDNone
            End If
            FG_DP.set_ColDataType(1, VSFlex8U.DataTypeSettings.flexDTBoolean)

        Else
            FG_DP.FormatString = "^ລ/ດ|<ເລືອກ|<ລະຫັດ   |<ສຳນັກງານ     "
            txt_permision_id.Text = "2"
            Cmb_DP.Enabled = True
            Cmb_DP.Items.Clear()
            Call load_Cmb("select Depart_nm from Depart where Depart_id<>'00'  ", "Depart_nm", Cmb_DP)
            Cmb_DP.SelectedIndex = 0
            Load_DP_Sub_Edit()
            Load_DP_Sub()
            Load_DP_Sub_Edit()
        End If

        Call LoadListFG2()


    End Sub

    Private Sub Fg_MouseUpEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseUpEvent) Handles Fg.MouseUpEvent

        'If Fg.get_TextMatrix(Fg.Row, 1) = "" Then Exit Sub
        'txtUsr_id.Text = Fg.get_TextMatrix(Fg.Row, 1)
        'txtOldPass.Text = Fg.get_TextMatrix(Fg.Row, 1)
        'LoadSqlData("Select Usr_id,Usr_nm,permision,Department  , PWD , Write_bit ,Edit_bit , Delete_bit  From  Ap_Users where  USr_Id = '" & txtUsr_id.Text & "' ", RSC)
        'txtUsr_id.Enabled = False
        'If RSC.RecordCount <> 0 Then
        '    cmbpermision.Text = (RSC.Fields("permision").Value.ToString)
        '    txtUsr_nm.Text = (RSC.Fields("Usr_nm").Value.ToString)
        '    TextBox1.Text = (RSC.Fields("Department").Value.ToString)
        '    txtPWD.Text = (RSC.Fields("PWD").Value.ToString)
        '    'MsgBox(RSC.Fields("Write_bit").Value)
        '    MDCheckWrite = (RSC.Fields("Write_bit").Value)
        '    MDCheckEdit = (RSC.Fields("Edit_bit").Value)
        '    MDCheckDelete = (RSC.Fields("Delete_bit").Value)


        '    CheckBox6.Checked = (RSC.Fields("Write_bit").Value)
        '    CheckBox5.Checked = (RSC.Fields("Edit_bit").Value)
        '    CheckBox4.Checked = (RSC.Fields("Delete_bit").Value)


        'End If





        'If txtUsr_id.Text = "a" Then
        '    'FG2.set_ColHidden(1, True)
        '    'FgItem.set_ColHidden(1, True)
        '    FG2.Editable = VSFlex8U.EditableSettings.flexEDNone
        '    FgItem.Editable = VSFlex8U.EditableSettings.flexEDNone
        '    'FG2.ForeColor = Color.Silver
        '    'FgItem.ForeColor = Color.Silver
        '    'FG2.Enabled = True
        'Else
        '    'FG2.set_ColHidden(1, False)
        '    'FgItem.set_ColHidden(1, False)
        '    FG2.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
        '    FgItem.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
        '    'FG2.ForeColor = Color.Black
        '    'FgItem.ForeColor = Color.Black
        'End If
        'FgItem.Rows = 1
        'FgItem.Rows = 2
        'Call LoadListFG2()
    End Sub







    Private Sub Fg_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Fg.SelChange


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



    Private Sub FgSec_DblClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles FgSec.DblClick
        'Panel1.Visible = False
        'txtDep_ID.Text = FgSec.get_TextMatrix(FgSec.Row, 1)
        'txtDep_Nm.Text = FgSec.get_TextMatrix(FgSec.Row, 2)
    End Sub

    Private Sub FgSec_MouseUpEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseUpEvent) Handles FgSec.MouseUpEvent
        FgItem.Cols = 4

        If FgSec.Row = 0 Then Exit Sub
        If FgSec.get_TextMatrix(FgSec.Row, 1) = "" Then Exit Sub
        FgItem.Rows = 1
        With rs
            Dim s As String = "select * from Ap_Section_Item" & _
                        " WHERE Sec_ID='" & FgSec.get_TextMatrix(FgSec.Row, 1) & "'  And Usr_Id = '" & txtUsr_id.Text & "' ORDER BY  cnt "
            Call LoadSqlData(s, rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    FgItem.AddItem(.AbsolutePosition & _
                    Chr(9) & (.Fields("Ints").Value) & _
                    Chr(9) & (.Fields("cnt").Value.ToString) & _
                    Chr(9) & (.Fields("Sec_Nm").Value.ToString))
                    .MoveNext()
                End While
            End If
        End With

        'If txtUsr_id.Text = "a" Then
        '    Exit Sub
        'End If


        For i = 1 To FgItem.Rows - 1
            If Microsoft.VisualBasic.Left(FgItem.get_TextMatrix(i, 3), 3) = "*  " Then
                FgItem.Row = i
                FgItem.Col = 3
                FgItem.CellFontBold = True
                FgItem.CellForeColor = Color.Blue
            End If

        Next i
        FgItem.Row = 1
        FgItem.Col = 3
        FgItem.CellFontBold = True
        FgItem.CellForeColor = Color.Red
        FgItem.Col = 2
    End Sub

    Private Sub FgSec_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FgSec.SelChange


    End Sub








    Private Sub Button9_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim i As Integer
        For i = 1 To FgItem.Rows - 1
            If FgItem.get_TextMatrix(i, 1) = True Then
                CNN.Execute("UPDATE Ap_Section_Item SET Ints=1 WHERE Sec_ID='" & FgItem.get_TextMatrix(i, 2) & "'AND Sec_Nm=N'" & FgItem.get_TextMatrix(i, 3) & "' ")
            ElseIf FgItem.get_TextMatrix(FgItem.Row, 1) = True Then
                CNN.Execute("UPDATE Ap_Section_Item SET Ints=1 WHERE Sec_ID='" & FgItem.get_TextMatrix(FgItem.Row, 2) & "' AND Sec_Nm=N'" & FgItem.get_TextMatrix(FgItem.Row, 3) & "' ")
            ElseIf FgItem.get_TextMatrix(FgItem.Row, 1) = False Then
                CNN.Execute("UPDATE Ap_Section_Item SET Ints=0 WHERE Sec_ID='" & FgItem.get_TextMatrix(FgItem.Row, 2) & "' AND Sec_Nm=N'" & FgItem.get_TextMatrix(FgItem.Row, 3) & "' ")
            Else
                CNN.Execute("UPDATE Ap_Section_Item SET Ints=0 WHERE Sec_ID='" & FgItem.get_TextMatrix(i, 2) & "' AND Sec_Nm=N'" & FgItem.get_TextMatrix(i, 3) & "' ")
            End If
        Next i
        MsgBox("ສໍາເລັດຜົນ!", MsgBoxStyle.OkOnly)
    End Sub

    Private Sub FgItem_MouseUpEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseUpEvent)
        If FgItem.Row > 0 Then
            CNN.Execute("Update Ap_Section_Item set Ints = " & FgItem.get_TextMatrix(FgItem.Row, 1) & " where cnt = " & FgItem.get_TextMatrix(FgItem.Row, 2) & "  And Usr_Id = '" & txtUsr_id.Text & "' ")
        End If
    End Sub





    Private Sub Button9_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Fm_Image.Img_ID.Text = txtUsr_id.Text
        'Fm_Image.ImgType.Text = "User"
        'Update_Image()
    End Sub




    Private Sub Sub_Company_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub



    Private Sub FgItem_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub FgItem_MouseUpEvent1(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseUpEvent) Handles FgItem.MouseUpEvent
        If FgItem.Row = 1 Then
            If FgItem.get_TextMatrix(1, 2) <> "" Then
                'MsgBox(FgItem.get_TextMatrix(1, 2))
                For i = 1 To FgItem.Rows - 1
                    FgItem.set_TextMatrix(i, 1, FgItem.get_TextMatrix(1, 1))
                    CNN.Execute("Update Ap_Section_Item set Ints = " & FgItem.get_TextMatrix(1, 1) & " where cnt = " & FgItem.get_TextMatrix(i, 2) & "  And Usr_Id = '" & txtUsr_id.Text & "' ")
                Next
            End If
        End If
        If FgItem.Row > 1 Then
            CNN.Execute("Update Ap_Section_Item set Ints = " & FgItem.get_TextMatrix(FgItem.Row, 1) & " where cnt = " & FgItem.get_TextMatrix(FgItem.Row, 2) & "  And Usr_Id = '" & txtUsr_id.Text & "' ")
        End If
    End Sub

    Private Sub Fg4_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FgItem.SelChange

    End Sub

    Private Sub FG2_MouseDownEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseDownEvent) Handles FG2.MouseDownEvent

        Dim s As String = FG2.get_TextMatrix(FG2.Row, 2)
        Dim d As String = FG2.get_TextMatrix(FG2.Row, 1)
        If s = "00-00" Then
            If d = True Then
                For i = 1 To FG2.Rows - 1
                    FG2.set_TextMatrix(i, 1, True)
                Next i
            End If
            If d = False Then
                For i = 1 To FG2.Rows - 1
                    FG2.set_TextMatrix(i, 1, False)
                Next i
            End If
        Else
            If Mid(s, 4, 2) = "00" Then
                For i = 1 To FG2.Rows - 1
                    If Mid(s, 1, 2) = Mid(FG2.get_TextMatrix(i, 2), 1, 2) Then
                        If d = True Then
                            FG2.set_TextMatrix(i, 1, True)
                        End If
                        If d = False Then
                            FG2.set_TextMatrix(i, 1, False)
                        End If
                    End If
                Next i
            End If
        End If

        LoadSqlData("Select Usr_id From Ap_Users Where Usr_id = '" & txtUsr_id.Text & "' ", RSC)
        If RSC.RecordCount <> 0 Then
            CNN.Execute("Delete Ap_Office_User Where Usr_Id = '" & txtUsr_id.Text & "'")
            For i = 1 To FG2.Rows - 1
                If FG2.get_TextMatrix(i, 1) = True Then
                    CNN.Execute("Insert into Ap_Office_User (Usr_Id,Sub_Id) VAlues ( '" & txtUsr_id.Text & "','" & FG2.get_TextMatrix(i, 2) & "') ")
                End If
            Next i
            CNN.Execute(" update Ap_Office_User set off_Id=Ap_Office.off_Id,off_add1=Ap_office.off_add1 ,off_add2=Ap_office.off_add2 from  Ap_Office_User , Ap_office where Ap_Office_User.Sub_Id = Ap_office.Sub_Id")
        End If

    End Sub

    Private Sub FG2_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG2.SelChange

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Panel3.Visible = True
        txtNewPass.Focus()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If txtUsr_id.Text = "a" Then MsgBox("a ບໍ່ສາມາມປ່ຽນໄດ້ຍ້ອນມັນເປັນລະຫັດຫລັກ , !", MsgBoxStyle.OkOnly) : txtUsr_id.Focus() : Exit Sub
        If txtNewPass.Text = "" Then MsgBox("ກະລຸນາຢືນຢັນລະຫັດຜ່ານຂອງທ່ານ", MsgBoxStyle.OkOnly) : txtNewPass.Focus() : Exit Sub
        Call LoadSqlData("select Usr_id from Ap_Users where Usr_Id = N'" & txtNewPass.Text & "' ", rs) : If rs.RecordCount <> 0 Then MsgBox(txtNewPass.Text & " ມີໃນຖານຂໍ້ມູນແລ້ວບໍ່າສມາດປ່ຽນໄດ້ , !", MsgBoxStyle.OkOnly) : Exit Sub
        CNN.Execute("UPDATE Ap_Users SET  Usr_Id = N'" & txtNewPass.Text & "'  Where   Usr_Id = N'" & txtOldPass.Text & "' ")
        CNN.Execute("UPDATE Ap_Office_User SET  Usr_Id = N'" & txtNewPass.Text & "'  Where   Usr_Id = N'" & txtOldPass.Text & "' ")
        CNN.Execute("UPDATE Ap_gen_jn SET  Usr_Id = N'" & txtNewPass.Text & "'  Where   Usr_Id = N'" & txtOldPass.Text & "' ")
        CNN.Execute("UPDATE Ap_Section_Item SET  Usr_Id = N'" & txtNewPass.Text & "'  Where   Usr_Id = N'" & txtOldPass.Text & "' ")
        MsgBox("ບັນທຶກສໍາເລັດຜົນ!", MsgBoxStyle.OkOnly)
        Call LoadData()

        SUPD = 0
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Panel3.Visible = False
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        InsertMenuStrip1()
        Call LoadSection()
        MsgBox("OK")
    End Sub

 

    Private Sub cmbpermision_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbpermision.SelectedIndexChanged
        If cmbpermision.SelectedIndex = 0 Then
            FG_DP.FormatString = "^ລ/ດ|<ເລືອກ|<ລະຫັດ   |<ສຳນັກງານ      "
            Cmb_DP.Items.Clear()
            Cmb_DP.Text = ""
            Cmb_DP.Enabled = False
            txt_permision_id.Text = "0"
            txt_DP_ID.Text = "00"
            Load_DP_00()
            'For i = 1 To FG_DP.Rows - 1
            '    FG_DP.set_TextMatrix(i, 1, True)
            'Next

        Else
            FG_DP.FormatString = "^ລ/ດ|<ເລືອກ|<ລະຫັດ   |<ສຳນັກງານ      "
            txt_permision_id.Text = "1"
            Cmb_DP.Items.Clear()
            Cmb_DP.Text = ""
            Cmb_DP.Enabled = False
            Load_DP()
            Load_DP_Edit()
            If FG_DP.Col = 1 Then
                FG_DP.Editable = VSFlex8U.EditableSettings.flexEDNone
            End If
            FG_DP.set_ColDataType(1, VSFlex8U.DataTypeSettings.flexDTBoolean)

            'Else
            '    FG_DP.FormatString = "^ລ/ດ|<ເລືອກ|<ລະຫັດ   |<ສຳນັກງານ         "
            '    txt_permision_id.Text = "2"
            '    Cmb_DP.Enabled = True
            '    Cmb_DP.Items.Clear()
            '    Call load_Cmb("select Depart_nm from Depart where Depart_id<>'00'  ", "Depart_nm", Cmb_DP)
            '    Cmb_DP.SelectedIndex = 0
            '    'Load_DP_Sub()
            '    'Load_DP_Sub_Edit()
        End If
    End Sub
    Public Sub Load_DP_Sub()
        FG_DP.Rows = 1
        With rs
            Call LoadSqlData("SELECT   * from Depart_sub  where 1=1  and Depart_id=N'" & txt_DP_ID.Text & "'  order by  Depart_sub_id", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    FG_DP.AddItem(.AbsolutePosition & _
                                         Chr(9) & 0 & _
                                Chr(9) & (.Fields("Depart_sub_id").Value.ToString) & _
                    Chr(9) & (.Fields("Depart_sub_nm").Value.ToString))
                    .MoveNext()
                End While
            Else
                FG_DP.Rows = 1
            End If
        End With
    End Sub
    Public Sub Load_DP()
        FG_DP.Rows = 1
        With rs
            Call LoadSqlData("SELECT   * from AP_Office where 1=1  order by  off_id", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    FG_DP.AddItem(.AbsolutePosition & _
                                      Chr(9) & 0 & _
                          Chr(9) & (.Fields("off_id").Value.ToString) & _
                    Chr(9) & (.Fields("off_name").Value.ToString))
                    .MoveNext()
                End While
            Else
                FG_DP.Rows = 2
            End If
        End With
    End Sub
    Public Sub Load_DP_00()
        FG_DP.Rows = 1
        With rs
            Call LoadSqlData("SELECT   * from Depart where Depart_id='00'  order by  Depart_id", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    FG_DP.AddItem(.AbsolutePosition & _
                                       Chr(9) & 1 & _
                          Chr(9) & (.Fields("Depart_id").Value.ToString) & _
                    Chr(9) & (.Fields("Depart_nm").Value.ToString))
                    .MoveNext()
                End While
            Else
                FG_DP.Rows = 2
            End If
        End With
    End Sub

    Public Sub Load_DP_Edit()
        With rs
            Call LoadSqlData("SELECT     AP_Users_List.Usr_id, AP_Office.off_id, AP_Office.off_name, AP_Users_List.status " & _
                   "   FROM         AP_Users_List INNER JOIN  " & _
                  "    AP_Office ON AP_Users_List.txt_DP_ID = AP_Office.off_id  where  Usr_id = N'" & txtUsr_id.Text & "'    order by  off_id", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    For i = 1 To FG_DP.Rows - 1
                        If FG_DP.get_TextMatrix(i, 2) = Trim((.Fields("off_id").Value).ToString) Then
                            FG_DP.set_TextMatrix(i, 1, True)
                        End If
                    Next
                    .MoveNext()
                End While
            End If
        End With
    End Sub
    Public Sub Load_DP_Sub_Edit()
        With rs
            Call LoadSqlData("SELECT   * from AP_Users_List  where  Usr_id = '" & txtUsr_id.Text & "'    order by  txt_DP_sub_ID", rs)
            If .RecordCount > 0 Then
                Cmb_DP.Text = Trim((.Fields("Cmb_DP").Value).ToString)
                txt_DP_ID.Text = Trim((.Fields("txt_DP_ID").Value).ToString)
                While Not .EOF()
                    For i = 1 To FG_DP.Rows - 1
                        'MsgBox(FG_DP.get_TextMatrix(i, 2))
                        'MsgBox(Trim((.Fields("txt_DP_sub_ID").Value).ToString))
                        If FG_DP.get_TextMatrix(i, 2) = Trim((.Fields("txt_DP_sub_ID").Value).ToString) Then
                            FG_DP.set_TextMatrix(i, 1, True)
                        End If
                    Next
                    .MoveNext()
                End While
            End If
        End With
    End Sub
    Private Sub Cmb_DP_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmb_DP.SelectedIndexChanged
        Dim RSC As New ADODB.Recordset
        Call LoadSqlData("Select * From Depart Where Depart_nm=N'" & Trim(Cmb_DP.Text) & "'   ", RSC)
        If RSC.RecordCount > 0 Then
            txt_DP_ID.Text = Trim(RSC("Depart_id").Value)
        End If
        Load_DP_Sub()
        'Load_DP_Sub_Edit()

    End Sub

    Private Sub FG_DP_ClickEvent(ByVal sender As Object, ByVal e As System.EventArgs) Handles FG_DP.ClickEvent
        If FG_DP.Col = 1 Then
            FG_DP.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
        End If
    End Sub

    Private Sub FG_DP_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG_DP.SelChange
        If FG_DP.Col = 1 Then
            FG_DP.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
        End If
    End Sub

    Private Sub txtOldPass_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOldPass.TextChanged

    End Sub

    Private Sub Button7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        'FrmStaff_item.ShowDialog()
        'If StaffID = "" Then
        '    txtUsr_id.Focus() : Exit Sub
        'Else
        '    txtUsr_id.Text = StaffID
        '    txtUsr_nm.Text = StaffNm
        'End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        'Frm_Of_component_for_user_Item.ShowDialog()
        'If Cat_id = "" Then
        '    txt_Component_id.Focus() : Exit Sub
        'Else
        '    txt_Component_id.Text = Cat_id
        '    txt_Component_nm.Text = Cat_nm

        'End If
    End Sub
End Class
