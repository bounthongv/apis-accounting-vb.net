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
        SetupGrid()
        SetupGridSec()
        SetupGridItem()
        SetupGrid2()
        SetupGridDP()

        FgSec.Visible = True
        FgItem.Visible = True
        SUPD = 0

        Call LoadData()
        Call LoadSection()
        txtUsr_id.Enabled = True
        txtUsr_id.Focus()

        EditActive = False
    End Sub

    Private Sub SetupGrid()
        Fg.Columns.Clear()
        Fg.Columns.Add("No", "ລ/ດ")
        Fg.Columns.Add("Usr_id", "ລະຫັດຜູ້ໃຊ້")
        Fg.Columns.Add("Usr_nm", "ຊື່ ຜູ້ໃຊ້")
        Fg.Columns.Add("Section", "ພາກ​ສ່ວນ/ສຳ​ນັກ​ງານ")
        Fg.Columns.Add("Responsibility", "ຮັບຜິດຊອບ")
        Fg.Columns.Add("InName", "ໃນນາມ")

        Fg.Columns(0).Width = 50
        Fg.Columns(1).Width = 100
        Fg.Columns(2).Width = 150
        Fg.Columns(3).Width = 150
        Fg.Columns(4).Width = 100
        Fg.Columns(5).Width = 100
        
        Fg.AllowUserToAddRows = False
        Fg.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Fg.ReadOnly = True
    End Sub

    Private Sub SetupGridSec()
        FgSec.Columns.Clear()
        FgSec.Columns.Add("No", "ລ/ດ")
        FgSec.Columns.Add("ID_Hidden", "")
        FgSec.Columns.Add("Sec_ID", "ລະຫັດພາກສ່ວນ")
        FgSec.Columns.Add("Sec_Nm", "ລາຍການພາກສ່ວນ")
        
        FgSec.Columns(1).Visible = False
        
        FgSec.AllowUserToAddRows = False
        FgSec.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FgSec.ReadOnly = True
    End Sub

    Private Sub SetupGridItem()
        FgItem.Columns.Clear()
        FgItem.Columns.Add("No", "ລ/ດ")
        Dim chkCol As New DataGridViewCheckBoxColumn()
        chkCol.HeaderText = ""
        chkCol.Name = "Check"
        FgItem.Columns.Add(chkCol)
        FgItem.Columns.Add("ID", "")
        FgItem.Columns.Add("Name", "")
        
        FgItem.Columns(2).Visible = False
        
        FgItem.AllowUserToAddRows = False
        FgItem.SelectionMode = DataGridViewSelectionMode.CellSelect
    End Sub

    Private Sub SetupGrid2()
        FG2.Columns.Clear()
        FG2.Columns.Add("No", "ລ/ດ")
        Dim chkCol As New DataGridViewCheckBoxColumn()
        chkCol.HeaderText = ""
        chkCol.Name = "Check"
        FG2.Columns.Add(chkCol)
        FG2.Columns.Add("User", "User")
        FG2.Columns.Add("Name", "")
        
        FG2.AllowUserToAddRows = False
        FG2.SelectionMode = DataGridViewSelectionMode.CellSelect
    End Sub

    Private Sub SetupGridDP()
        FG_DP.Columns.Clear()
        FG_DP.Columns.Add("No", "ລ/ດ")
        Dim chkCol As New DataGridViewCheckBoxColumn()
        chkCol.HeaderText = "ເລືອກ"
        chkCol.Name = "Check"
        FG_DP.Columns.Add(chkCol)
        FG_DP.Columns.Add("ID", "ລະຫັດ")
        FG_DP.Columns.Add("Name", "ກົມກອງ")
        
        FG_DP.AllowUserToAddRows = False
        FG_DP.SelectionMode = DataGridViewSelectionMode.CellSelect
    End Sub
    Public Sub LoadListFG2()
        Dim aa As String
        aa = "update Ap_Office set Lck=0 update Ap_Office set Lck=1 from  Ap_Office_User , Ap_office where Ap_Office_User.Sub_Id = Ap_office.Sub_Id And Usr_Id =  '" & txtUsr_id.Text & "' "
        CNN.Execute(aa)
        FG2.Rows.Clear()
        With RSC
            Call LoadSqlData("select Sub_Id , Off_Add2,Lck  from Ap_office order by Sub_Id", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG2.Rows.Add(.AbsolutePosition, _
                                (.Fields("Lck").Value), _
                                Trim(CStr(.Fields("Sub_Id").Value)), _
                                Trim(CStr(.Fields("Off_Add2").Value)))
                    .MoveNext()
                End While
            End If
        End With

        For i = 0 To FG2.Rows.Count - 1
            Dim subId As String = If(FG2.Rows(i).Cells(2).Value Is Nothing, "", FG2.Rows(i).Cells(2).Value.ToString())
            If subId = "00-00" Then
                FG2.Rows(i).Cells(2).Style.ForeColor = Color.Red
                FG2.Rows(i).Cells(2).Style.Font = New Font(FG2.Font, FontStyle.Bold)
                FG2.Rows(i).Cells(3).Style.ForeColor = Color.Red
                FG2.Rows(i).Cells(3).Style.Font = New Font(FG2.Font, FontStyle.Bold)
            End If
            If Mid(subId, 1, 2) <> "00" Then
                If Mid(subId, 4, 2) = "00" Then
                    FG2.Rows(i).Cells(2).Style.ForeColor = Color.Blue
                    FG2.Rows(i).Cells(2).Style.Font = New Font(FG2.Font, FontStyle.Bold)
                    FG2.Rows(i).Cells(3).Style.ForeColor = Color.Blue
                    FG2.Rows(i).Cells(3).Style.Font = New Font(FG2.Font, FontStyle.Bold)
                End If
            End If
        Next i
    End Sub

    Private Sub LoadData()
        Fg.Rows.Clear()
        With rs
            Call LoadSqlData("select Usr_id,Usr_nm,permision   from Ap_Users Order by Sec_id ", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    Fg.Rows.Add(.AbsolutePosition, _
                                (.Fields("Usr_id").Value.ToString), _
                                (.Fields("Usr_nm").Value.ToString), _
                                "", "", _
                                (.Fields("permision").Value.ToString))
                    .MoveNext()
                End While
            End If
        End With
    End Sub

    Private Sub LoadSection()
        FgSec.Rows.Clear()
        With rs
            Call LoadSqlData("select * from Ap_Section Order by Sec_ID ", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    FgSec.Rows.Add(.AbsolutePosition, _
                                   "", _
                                   (.Fields("Sec_ID").Value.ToString), _
                                   (.Fields("Sec_Nm").Value.ToString))
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
                For i = 0 To FG_DP.Rows.Count - 1
                    If .RecordCount = 0 Then
                        If CBool(If(FG_DP.Rows(i).Cells(1).Value Is Nothing, False, FG_DP.Rows(i).Cells(1).Value)) = True Then
                            CNN.Execute("INSERT INTO  AP_Users_List (Usr_id,Stff_Id,Usr_nm,permision,permision_id,status,txt_DP_ID,Cmb_DP,Get_date, lst_usr, pc_nm) " & _
                                " VALUES(N'" & (txtUsr_id.Text) & "'," & _
                                      " N'" & (TextBox1.Text) & "'," & _
                           " N'" & (txtUsr_nm.Text) & "'," & _
                                " N'" & (cmbpermision.Text) & "'," & _
                           " N'" & (txt_permision_id.Text) & "'," & _
                             " N'" & 1 & "'," & _
                                " N'" & Apostrophe(If(FG_DP.Rows(i).Cells(2).Value Is Nothing, "", FG_DP.Rows(i).Cells(2).Value.ToString())) & "'," & _
                                  " N'" & Apostrophe(If(FG_DP.Rows(i).Cells(3).Value Is Nothing, "", FG_DP.Rows(i).Cells(3).Value.ToString())) & "'," & _
                                   " Getdate()," & _
                              " N'" & MUserName & "'," & _
                              " '" & COMPUTER_NM & "')")
                        End If
                    End If
                Next i

            ElseIf cmbpermision.SelectedIndex = 2 Then

                Call LoadSqlData("SELECT * FROM AP_Users_List WHERE Usr_id = N'" & txtUsr_id.Text & "'", Rschk)
                For i = 0 To FG_DP.Rows.Count - 1
                    If .RecordCount = 0 Then
                        If CBool(If(FG_DP.Rows(i).Cells(1).Value Is Nothing, False, FG_DP.Rows(i).Cells(1).Value)) = True Then
                            CNN.Execute("INSERT INTO  AP_Users_List (Usr_id,Stff_Id,Usr_nm,permision,permision_id,txt_DP_ID,Cmb_DP,txt_DP_sub_ID,Cmb_DP_sub,Get_date, lst_usr, pc_nm) " & _
                                " VALUES(N'" & (txtUsr_id.Text) & "'," & _
                                      " N'" & (TextBox1.Text) & "'," & _
                           " N'" & (txtUsr_nm.Text) & "'," & _
                                " N'" & (cmbpermision.Text) & "'," & _
                           " N'" & (txt_permision_id.Text) & "'," & _
                                " N'" & (txt_DP_ID.Text) & "'," & _
                                    " N'" & (Cmb_DP.Text) & "'," & _
                                 " N'" & Apostrophe(If(FG_DP.Rows(i).Cells(2).Value Is Nothing, "", FG_DP.Rows(i).Cells(2).Value.ToString())) & "'," & _
                                   " N'" & Apostrophe(If(FG_DP.Rows(i).Cells(3).Value Is Nothing, "", FG_DP.Rows(i).Cells(3).Value.ToString())) & "'," & _
                                    " Getdate()," & _
                              " N'" & MUserName & "'," & _
                              " '" & COMPUTER_NM & "')")
                        End If
                    End If
                Next i
            Else
                Call LoadSqlData("SELECT * FROM AP_Users_List WHERE Usr_id = N'" & txtUsr_id.Text & "'", Rschk)
                For i = 0 To FG_DP.Rows.Count - 1
                    If .RecordCount = 0 Then
                        If CBool(If(FG_DP.Rows(i).Cells(1).Value Is Nothing, False, FG_DP.Rows(i).Cells(1).Value)) = True Then
                            CNN.Execute("INSERT INTO  AP_Users_List (Usr_id,Stff_Id,Usr_nm,permision,permision_id,txt_DP_ID,Cmb_DP,Get_date, lst_usr, pc_nm) " & _
                                " VALUES(N'" & (txtUsr_id.Text) & "'," & _
                                      " N'" & (TextBox1.Text) & "'," & _
                           " N'" & (txtUsr_nm.Text) & "'," & _
                                " N'" & (cmbpermision.Text) & "'," & _
                           " N'" & (txt_permision_id.Text) & "'," & _
                                " N'" & Apostrophe(If(FG_DP.Rows(i).Cells(2).Value Is Nothing, "", FG_DP.Rows(i).Cells(2).Value.ToString())) & "'," & _
                                  " N'" & Apostrophe(If(FG_DP.Rows(i).Cells(3).Value Is Nothing, "", FG_DP.Rows(i).Cells(3).Value.ToString())) & "'," & _
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
            For i = 0 To FG2.Rows.Count - 1
                Dim subId As String = If(FG2.Rows(i).Cells(2).Value Is Nothing, "", FG2.Rows(i).Cells(2).Value.ToString())
                CNN.Execute("Insert into Ap_Office_User (Usr_Id,Sub_Id) VAlues ( N'" & Trim(txtUsr_id.Text) & "',N'" & Trim(subId) & "') ")
            Next i
            'CNN.Execute(" update Ap_Office_User set off_Id=Ap_Office.off_Id,off_add1=Ap_office.off_add1 ,off_add2=Ap_office.off_add2 from  Ap_Office_User , Ap_office where Ap_Office_User.Sub_Id = Ap_office.Sub_Id")
            CNN.Execute("Update Ap_Section_Item set Ints =1 Where Usr_Id = N'" & Trim(txtUsr_id.Text) & "'")
        End If

    End Sub
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDel.Click
        If txtUsr_id.Text = "a" Then MsgBox("a ບໍ່ສາມາມລືບໄດ້ຍ້ອນມັນເປັນລະຫັດຫລັກ , !", MsgBoxStyle.OkOnly) : txtUsr_id.Focus() : Exit Sub
        If Fg.Rows.Count <= 1 Then MsgBox("ທ່ານບໍ່ສາມາດລຶບລາຍການນີ້ໄດ້, ເພາະແມ່ນລາຍການສຸດທ້າຍແລ້ວ ", MsgBoxStyle.OkOnly) : Exit Sub
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

    Private Sub Fg_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles Fg.CellClick
    End Sub

    Private Sub Fg_CellDoubleClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles Fg.CellDoubleClick
        If e.RowIndex < 0 Then Exit Sub
        If Fg.Rows(e.RowIndex).Cells(1).Value Is Nothing Then Exit Sub
        
        txtUsr_id.Text = Fg.Rows(e.RowIndex).Cells(1).Value.ToString()
        txtOldPass.Text = Fg.Rows(e.RowIndex).Cells(1).Value.ToString()
        txtUsr_nm.Text = Fg.Rows(e.RowIndex).Cells(2).Value.ToString()
        
        LoadSqlData("Select Usr_id,Usr_nm,permision,Department , off_id,off_nm , PWD , Write_bit ,Edit_bit , Delete_bit  From  Ap_Users where  USr_Id = N'" & Trim(txtUsr_id.Text) & "' ", RSC)
        txtUsr_id.Enabled = False
        If RSC.RecordCount <> 0 Then
            cmbpermision.Text = (RSC.Fields("permision").Value.ToString)
            txtUsr_nm.Text = (RSC.Fields("Usr_nm").Value.ToString)

            txt_Component_id.Text = (RSC.Fields("off_id").Value.ToString)
            txt_Component_nm.Text = (RSC.Fields("off_nm").Value.ToString)

            TextBox1.Text = (RSC.Fields("Department").Value.ToString)
            txtPWD.Text = (RSC.Fields("PWD").Value.ToString)
            
            MDCheckWrite = (RSC.Fields("Write_bit").Value)
            MDCheckEdit = (RSC.Fields("Edit_bit").Value)
            MDCheckDelete = (RSC.Fields("Delete_bit").Value)

            CheckBox6.Checked = (RSC.Fields("Write_bit").Value)
            CheckBox5.Checked = (RSC.Fields("Edit_bit").Value)
            CheckBox4.Checked = (RSC.Fields("Delete_bit").Value)
        End If

        If txtUsr_id.Text = "a" Then
            FG2.ReadOnly = True
            FgItem.ReadOnly = True
        Else
            FG2.ReadOnly = False
            FgItem.ReadOnly = False
        End If
        
        FgItem.Rows.Clear()
        EditActive = True

        If cmbpermision.SelectedIndex = 0 Then
            Cmb_DP.Items.Clear()
            Cmb_DP.Text = ""
            Cmb_DP.Enabled = False
            txt_permision_id.Text = "0"
            txt_DP_ID.Text = "00"
            Load_DP_00()
        ElseIf cmbpermision.SelectedIndex = 1 Then
            FG_DP.Columns(1).HeaderText = "ເລືອກ"
            FG_DP.Columns(2).HeaderText = "ລະຫັດ"
            FG_DP.Columns(3).HeaderText = "ສຳນັກງານ"
            
            txt_permision_id.Text = "1"
            Cmb_DP.Items.Clear()
            Cmb_DP.Text = ""
            Cmb_DP.Enabled = False
            Load_DP()
            Load_DP_Edit()
        Else
            FG_DP.Columns(1).HeaderText = "ເລືອກ"
            FG_DP.Columns(2).HeaderText = "ລະຫັດ"
            FG_DP.Columns(3).HeaderText = "ສຳນັກງານ"
            
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

    Private Sub Fg_MouseUpEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseUpEvent)

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







    Private Sub Fg_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Fg.SelectionChanged


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



    Private Sub FgSec_CellDoubleClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FgSec.CellDoubleClick
    End Sub

    Private Sub FgSec_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FgSec.CellClick
        If e.RowIndex < 0 Then Exit Sub
        Dim secId As String = If(FgSec.Rows(e.RowIndex).Cells(2).Value Is Nothing, "", FgSec.Rows(e.RowIndex).Cells(2).Value.ToString())
        If secId = "" Then Exit Sub
        
        FgItem.Rows.Clear()
        With rs
            Dim s As String = "select * from Ap_Section_Item" & _
                        " WHERE Sec_ID='" & secId & "'  And Usr_Id = '" & txtUsr_id.Text & "' ORDER BY  cnt "
            Call LoadSqlData(s, rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    FgItem.Rows.Add(.AbsolutePosition, _
                                   (.Fields("Ints").Value), _
                                   (.Fields("cnt").Value.ToString), _
                                   (.Fields("Sec_Nm").Value.ToString))
                    .MoveNext()
                End While
            End If
        End With

        For i = 0 To FgItem.Rows.Count - 1
            Dim secNm As String = If(FgItem.Rows(i).Cells(3).Value Is Nothing, "", FgItem.Rows(i).Cells(3).Value.ToString())
            If Microsoft.VisualBasic.Left(secNm, 3) = "*  " Then
                FgItem.Rows(i).Cells(3).Style.Font = New Font(FgItem.Font, FontStyle.Bold)
                FgItem.Rows(i).Cells(3).Style.ForeColor = Color.Blue
            End If
        Next i
        
        If FgItem.Rows.Count > 0 Then
            FgItem.Rows(0).Cells(3).Style.Font = New Font(FgItem.Font, FontStyle.Bold)
            FgItem.Rows(0).Cells(3).Style.ForeColor = Color.Red
        End If
    End Sub

    Private Sub FgSec_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FgSec.SelectionChanged


    End Sub








    Private Sub Button9_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim i As Integer
        For i = 0 To FgItem.Rows.Count - 1
            Dim isChecked As Boolean = CBool(If(FgItem.Rows(i).Cells(1).Value Is Nothing, False, FgItem.Rows(i).Cells(1).Value))
            Dim secId As String = If(FgItem.Rows(i).Cells(2).Value Is Nothing, "", FgItem.Rows(i).Cells(2).Value.ToString())
            Dim secNm As String = If(FgItem.Rows(i).Cells(3).Value Is Nothing, "", FgItem.Rows(i).Cells(3).Value.ToString())
            
            If isChecked = True Then
                CNN.Execute("UPDATE Ap_Section_Item SET Ints=1 WHERE Sec_ID='" & secId & "'AND Sec_Nm=N'" & secNm & "' ")
            Else
                CNN.Execute("UPDATE Ap_Section_Item SET Ints=0 WHERE Sec_ID='" & secId & "' AND Sec_Nm=N'" & secNm & "' ")
            End If
        Next i
        MsgBox("ສໍາເລັດຜົນ!", MsgBoxStyle.OkOnly)
    End Sub

    Private Sub FgItem_MouseUpEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseUpEvent)
        If FgItem.CurrentRow IsNot Nothing AndAlso FgItem.CurrentRow.Index >= 0 Then
            CNN.Execute("Update Ap_Section_Item set Ints = " & Convert.ToString(FgItem.CurrentRow.Cells(1).Value) & " where cnt = " & Convert.ToString(FgItem.CurrentRow.Cells(2).Value) & "  And Usr_Id = '" & txtUsr_id.Text & "' ")
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

    Private Sub FgItem_CellContentClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FgItem.CellContentClick
        If e.ColumnIndex = 1 Then
            FgItem.CommitEdit(DataGridViewDataErrorContexts.Commit)
            Dim isChecked As Boolean = CBool(If(FgItem.Rows(e.RowIndex).Cells(1).Value Is Nothing, False, FgItem.Rows(e.RowIndex).Cells(1).Value))
            
            If e.RowIndex = 0 Then
                Dim cnt_Val As String = If(FgItem.Rows(0).Cells(2).Value Is Nothing, "", FgItem.Rows(0).Cells(2).Value.ToString())
                If cnt_Val <> "" Then
                    For i = 0 To FgItem.Rows.Count - 1
                        FgItem.Rows(i).Cells(1).Value = isChecked
                        Dim rowCnt As String = If(FgItem.Rows(i).Cells(2).Value Is Nothing, "", FgItem.Rows(i).Cells(2).Value.ToString())
                        CNN.Execute("Update Ap_Section_Item set Ints = " & If(isChecked, 1, 0) & " where cnt = " & rowCnt & "  And Usr_Id = '" & txtUsr_id.Text & "' ")
                    Next
                End If
            Else
                Dim rowCnt As String = If(FgItem.Rows(e.RowIndex).Cells(2).Value Is Nothing, "", FgItem.Rows(e.RowIndex).Cells(2).Value.ToString())
                CNN.Execute("Update Ap_Section_Item set Ints = " & If(isChecked, 1, 0) & " where cnt = " & rowCnt & "  And Usr_Id = '" & txtUsr_id.Text & "' ")
            End If
        End If
    End Sub

    Private Sub FgItem_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FgItem.SelectionChanged

    End Sub

    Private Sub FG2_CellContentClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FG2.CellContentClick
        If e.RowIndex < 0 Or e.ColumnIndex <> 1 Then Exit Sub
        
        FG2.CommitEdit(DataGridViewDataErrorContexts.Commit)
        Dim subId_Val As String = If(FG2.Rows(e.RowIndex).Cells(2).Value Is Nothing, "", FG2.Rows(e.RowIndex).Cells(2).Value.ToString())
        Dim isChecked As Boolean = CBool(If(FG2.Rows(e.RowIndex).Cells(1).Value Is Nothing, False, FG2.Rows(e.RowIndex).Cells(1).Value))
        
        If subId_Val = "00-00" Then
            For i = 0 To FG2.Rows.Count - 1
                FG2.Rows(i).Cells(1).Value = isChecked
            Next i
        Else
            If Mid(subId_Val, 4, 2) = "00" Then
                For i = 0 To FG2.Rows.Count - 1
                    Dim rowSubId As String = If(FG2.Rows(i).Cells(2).Value Is Nothing, "", FG2.Rows(i).Cells(2).Value.ToString())
                    If Mid(subId_Val, 1, 2) = Mid(rowSubId, 1, 2) Then
                        FG2.Rows(i).Cells(1).Value = isChecked
                    End If
                Next i
            End If
        End If

        LoadSqlData("Select Usr_id From Ap_Users Where Usr_id = '" & txtUsr_id.Text & "' ", RSC)
        If RSC.RecordCount <> 0 Then
            CNN.Execute("Delete Ap_Office_User Where Usr_Id = '" & txtUsr_id.Text & "'")
            For i = 0 To FG2.Rows.Count - 1
                Dim rowChecked As Boolean = CBool(If(FG2.Rows(i).Cells(1).Value Is Nothing, False, FG2.Rows(i).Cells(1).Value))
                If rowChecked = True Then
                    Dim rowSubId As String = If(FG2.Rows(i).Cells(2).Value Is Nothing, "", FG2.Rows(i).Cells(2).Value.ToString())
                    CNN.Execute("Insert into Ap_Office_User (Usr_Id,Sub_Id) VAlues ( '" & txtUsr_id.Text & "','" & rowSubId & "') ")
                End If
            Next i
            CNN.Execute(" update Ap_Office_User set off_Id=Ap_Office.off_Id,off_add1=Ap_office.off_add1 ,off_add2=Ap_office.off_add2 from  Ap_Office_User , Ap_office where Ap_Office_User.Sub_Id = Ap_office.Sub_Id")
        End If
    End Sub

    Private Sub FG2_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG2.SelectionChanged

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
            'FG_DP.FormatString = "^ລ/ດ|<ເລືອກ|<ລະຫັດ   |<ສຳນັກງານ      "
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
            'FG_DP.FormatString = "^ລ/ດ|<ເລືອກ|<ລະຫັດ   |<ສຳນັກງານ      "
            txt_permision_id.Text = "1"
            Cmb_DP.Items.Clear()
            Cmb_DP.Text = ""
            Cmb_DP.Enabled = False
            Load_DP()
            Load_DP_Edit()
            ' Legacy VSFlexGrid configuration removed

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
        FG_DP.Rows.Clear()
        With rs
            Call LoadSqlData("SELECT   * from Depart_sub  where 1=1  and Depart_id=N'" & txt_DP_ID.Text & "'  order by  Depart_sub_id", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    FG_DP.Rows.Add(.AbsolutePosition, _
                                   False, _
                                   (.Fields("Depart_sub_id").Value.ToString), _
                                   (.Fields("Depart_sub_nm").Value.ToString))
                    .MoveNext()
                End While
            End If
        End With
    End Sub
    Public Sub Load_DP()
        FG_DP.Rows.Clear()
        With rs
            Call LoadSqlData("SELECT   * from AP_Office where 1=1  order by  off_id", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    FG_DP.Rows.Add(.AbsolutePosition, _
                                   False, _
                                   (.Fields("off_id").Value.ToString), _
                                   (.Fields("off_name").Value.ToString))
                    .MoveNext()
                End While
            End If
        End With
    End Sub
    Public Sub Load_DP_00()
        FG_DP.Rows.Clear()
        With rs
            Call LoadSqlData("SELECT   * from Depart where Depart_id='00'  order by  Depart_id", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    FG_DP.Rows.Add(.AbsolutePosition, _
                                   True, _
                                   (.Fields("Depart_id").Value.ToString), _
                                   (.Fields("Depart_nm").Value.ToString))
                    .MoveNext()
                End While
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
                    For i = 0 To FG_DP.Rows.Count - 1
                        Dim offId As String = If(FG_DP.Rows(i).Cells(2).Value Is Nothing, "", FG_DP.Rows(i).Cells(2).Value.ToString())
                        If offId = Trim((.Fields("off_id").Value).ToString) Then
                            FG_DP.Rows(i).Cells(1).Value = True
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
                    For i = 0 To FG_DP.Rows.Count - 1
                        Dim subId As String = If(FG_DP.Rows(i).Cells(2).Value Is Nothing, "", FG_DP.Rows(i).Cells(2).Value.ToString())
                        If subId = Trim((.Fields("txt_DP_sub_ID").Value).ToString) Then
                            FG_DP.Rows(i).Cells(1).Value = True
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
    End Sub

    Private Sub FG_DP_CellContentClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FG_DP.CellContentClick
        ' Handling checkbox interaction if needed beyond ReadOnly toggle
    End Sub

    Private Sub FG_DP_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG_DP.SelectionChanged
    End Sub

    Private Sub txtOldPass_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOldPass.TextChanged
    End Sub

    Private Sub Button7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
    End Sub
End Class
