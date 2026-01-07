Public Class fmShartOfAcc
    Dim Status As Boolean
    Dim AcTypeLao, AcTypeEng As String
    Dim SQL As String
    Dim a, Cs As String
    
    Private Sub SetupGrid()
        FG.Columns.Clear()
        FG.Columns.Add("Col0", "ລ/ດ")
        FG.Columns.Add("Col1", "ເລກລະຫັດ")
        FG.Columns.Add("Col2", "WISE Orginal")
        FG.Columns.Add("Col3", "ຊື່ພາສາລາວ")
        FG.Columns.Add("Col4", "ຊື່ພາສາອັງກິດ")
        FG.Columns.Add("Col5", "ປະເພດບັນຊີ ພາສາລາວ")
        FG.Columns.Add("Col6", "ປະເພດບັນຊີ ພາສາອັງກິດ")
        FG.Columns.Add("Col7", "Print_Status")
        
        ' Set column widths
        FG.Columns("Col0").Width = 50
        FG.Columns("Col1").Width = 150
        FG.Columns("Col2").Width = 120
        FG.Columns("Col3").Width = 250
        FG.Columns("Col4").Width = 250
        FG.Columns("Col5").Width = 120
        FG.Columns("Col6").Width = 120
        FG.Columns("Col7").Width = 80
        
        ' Hide WISE Orginal column (column index 2)
        FG.Columns("Col2").Visible = False
        FG.Columns("Col7").Visible = False ' Hide print status column initially
    End Sub

    Private Sub fmShartOfAcc_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Call MdiCNum()
    End Sub
    Private Sub LoadDividePage()
        If p25.Checked = True Then
            DividePage = 25
        ElseIf p50.Checked = True Then
            DividePage = 50
        ElseIf p100.Checked = True Then
            DividePage = 100
        ElseIf p250.Checked = True Then
            DividePage = 250
        ElseIf p500.Checked = True Then
            DividePage = 500
        ElseIf p1000.Checked = True Then
            DividePage = 1000
        End If
    End Sub
Private Sub frmShartOfAcc_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        p250.Checked = True
        p250.ForeColor = Color.Red
        SetupGrid()
        CmbPrinSelete.Text = "ພີມທຸຫລາຍການ (Prin All)"
        Rdlasth.Enabled = False
        similar.Enabled = False
        StartLoadDataList()

        SetControlText(Me)
        If MuLng = "L" Then
            ' Update column headers for Lao language
            FG.Columns("Col0").HeaderText = "ລ/ດ"
            FG.Columns("Col1").HeaderText = "ເລກລະຫັດ"
            FG.Columns("Col2").HeaderText = "WISE Orginal"
            FG.Columns("Col3").HeaderText = "ຊື່ພາສາລາວ"
            FG.Columns("Col4").HeaderText = "ຊື່ພາສາອັງກິດ"
            FG.Columns("Col5").HeaderText = "ປະເພດບັນຊີ ພາສາລາວ"
            FG.Columns("Col6").HeaderText = "ປະເພດບັນຊີ ພາສາອັງກິດ"
        Else
            ' Update column headers for English language
            FG.Columns("Col0").HeaderText = "No"
            FG.Columns("Col1").HeaderText = "Ac Code"
            FG.Columns("Col2").HeaderText = "WISE Orginal"
            FG.Columns("Col3").HeaderText = "Name Lao"
            FG.Columns("Col4").HeaderText = "Name Eng"
            FG.Columns("Col5").HeaderText = "Type Lao"
            FG.Columns("Col6").HeaderText = "Type Eng"
        End If

    End Sub



    Private Sub BntNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BntNew.Click
        FrNewAcc.txtAc_code.Enabled = True
        FrNewAcc.ShowDialog()
        'MDActivated = False
    End Sub

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        'FmMain.PictureBox1.Visible = True
        Me.Close()
    End Sub

Private Sub FG_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles FG.CellClick
        If e.RowIndex < 0 Then Exit Sub
        
        If CmbPrinSelete.Text = "ເລືອກລາຍການພິມ (Prin All Items)" Then
            If Convert.ToBoolean(FG.Rows(e.RowIndex).Cells("Col7").Value) = True Then
                FG.Rows(e.RowIndex).Cells("Col7").Value = False
                DbHelper.ExecuteNonQuery("UPDATE Acc_Code SET Print_status='" & "0" & "' WHERE AC_CODE = '" & FG.Rows(e.RowIndex).Cells("Col1").Value.ToString() & "'")

                If Convert.ToBoolean(FG.Rows(e.RowIndex).Cells("Col7").Value) = False Then
                    For I = 0 To FG.Columns.Count - 1
                        FG.Rows(e.RowIndex).Cells(I).Style.BackColor = Color.SkyBlue
                    Next I
                End If



            Else
                FG.Rows(e.RowIndex).Cells("Col7").Value = True
                DbHelper.ExecuteNonQuery("UPDATE Acc_Code SET Print_status='" & "1" & "' WHERE AC_CODE = '" & FG.Rows(e.RowIndex).Cells("Col1").Value.ToString() & "'")
                If Convert.ToBoolean(FG.Rows(e.RowIndex).Cells("Col7").Value) = True Then
                    For I = 0 To FG.Columns.Count - 1
                        FG.Rows(e.RowIndex).Cells(I).Style.BackColor = Color.White
                    Next
                End If

            End If

        End If
    End Sub



Private Sub FG_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles FG.CellDoubleClick
        If e.RowIndex < 0 Then Exit Sub
        FrNewAcc.txtAc_code.Text = FG.Rows(e.RowIndex).Cells("Col1").Value.ToString()
        FrNewAcc.txtAc_code.Enabled = False
        FrNewAcc.ShowDialog()
        'MDActivated = True
    End Sub



Private Sub FG_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles FG.SelectionChanged
        If FG.CurrentRow IsNot Nothing Then
            FrNewAcc.txtAc_code.Text = FG.CurrentRow.Cells("Col1").Value.ToString()
            txtOldId.Text = FG.CurrentRow.Cells("Col1").Value.ToString()
            FrNewAcc.txtAc_code.Enabled = False
        End If
    End Sub

    Private Sub BtnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnEdit.Click
        If FrNewAcc.txtAc_code.Enabled = True Then MsgBox("ກະລຸນາເລືອກກ່ອນ!", MsgBoxStyle.OkOnly) : Exit Sub
        FrNewAcc.ShowDialog()
    End Sub

Private Sub BtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelete.Click
        If FrNewAcc.txtAc_code.Enabled = True Then MsgBox("ກະລຸນາເລືອກກ່ອນ!", MsgBoxStyle.OkOnly) : Exit Sub

        If FG.CurrentRow Is Nothing Then Exit Sub
        Dim acCode As String = FG.CurrentRow.Cells("Col1").Value.ToString()

        Dim dtGen As DataTable = DbHelper.GetDataTable("SELECT AC_CODE FROM Gen_jn WHERE AC_CODE = '" & acCode & "'")
        If dtGen.Rows.Count >0 Then
            MsgBox("ເລກລະຫັດ : " & acCode & " ມີການເຄີ່ຶນໄຫວຢູ່ບັນຊີປະຈຳວັນແລ້ວບໍ່ສາມາດລຶບໄດ້!", MsgBoxStyle.OkOnly)
            txtNewId.Focus()
            Exit Sub

        End If


        Dim dtOpen As DataTable = DbHelper.GetDataTable("SELECT AC_CODE FROM Open_jn WHERE AC_CODE = '" & acCode & "'")
        If dtOpen.Rows.Count >0 Then
            MsgBox("ເລກລະຫັດ : " & acCode & " ມີການເຄີ່ຶນໄຫວຢູ່ຍອດໍແລ້ວບໍ່ສາມາດລຶບໄດ້!", MsgBoxStyle.OkOnly)
            txtNewId.Focus()
            Exit Sub

        End If




        If MessageBox.Show("ທ່ານຕ້ອງລຶບລະຫັດ " & acCode & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            DbHelper.ExecuteNonQuery("DELETE FROM Acc_Code WHERE AC_CODE='" & acCode & "'")
            FrNewAcc.txtAc_code.Enabled = True
        End If
        LoadSQL()
        StartLoadDataList()
    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Panel3.Visible = True
        Panel1.Visible = True
    End Sub



    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Panel2.Visible = False
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'CNN.Execute("UPDATE Acc_Code SET Print_status='0'")
        'For k = 1 To FG.Rows - 1
        '    FG.set_TextMatrix(k, 6, False)
        'Next k
        'Dim rmRS As New ADODB.Recordset
        'Dim I, J As Integer
        'FG.Redraw = False
        'For J = 1 To FG.Rows - 1
        '    If Trim(FG.get_TextMatrix(J, 6)) = True Then
        '        FG.Row = J
        '        For I = 1 To FG.Cols - 1
        '            FG.Col = I
        '            FG.CellBackColor = Color.SkyBlue
        '        Next I
        '    Else

        '        FG.Row = J
        '        For I = 1 To FG.Cols - 1
        '            FG.Col = I
        '            FG.CellBackColor = Color.White
        '        Next

        '    End If

        'Next J
        'FG.Redraw = True
        'If CmbPrinSelete.Text = "ເລືອກລາຍການພິມ (Prin All Items)" Then
        '    CheckBox1.Checked = False
        '    CheckBox1.Enabled = False
        'Else
        '    CheckBox1.Enabled = True
        'End If




        Panel2.Visible = True
        Panel1.Visible = False
        'txtOldId.Clear()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Cs = 0
        For i = 1 To Len(Trim(txtNewId.Text))
            If Mid(Trim(txtNewId.Text), i, 1) = "." Then
                AcTypeLao = "ບັນຊີຍ່ອຍ (D)"
                AcTypeEng = "Detail Account"
                Exit For
            Else
                AcTypeLao = "ບັນຊີແມ່ (P)"
                AcTypeEng = "Parent Account"
            End If
        Next i

        Dim dtCheck As DataTable = DbHelper.GetDataTable("SELECT AC_CODE FROM Acc_Code WHERE AC_CODE = '" & Trim(txtNewId.Text) & "'")
        If dtCheck.Rows.Count > 0 Then
            MsgBox("ເລກລະຫັດ : " & Trim(txtNewId.Text) & " ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
            txtNewId.Focus()
            Exit Sub

        End If

        If txtOldId.Text = "" Then MsgBox("ກະລຸນາ ເລືອກລະຫັດກ່ອນກ່ອນ!", MsgBoxStyle.OkOnly) : Exit Sub
        If txtNewId.Text = "" Then MsgBox("ກະລຸນາ ໃສ່ລະຫັດໃໝ່ກ່ອນ!", MsgBoxStyle.OkOnly) : Exit Sub





        If FG.CurrentRow Is Nothing Then Exit Sub
        Dim dtGen As DataTable = DbHelper.GetDataTable("SELECT AC_CODE FROM Gen_jn WHERE AC_CODE = '" & FG.CurrentRow.Cells("Col1").Value.ToString() & "'")






        If dtGen.Rows.Count > 0 Then

            Cs = 1

        End If






        If FG.CurrentRow Is Nothing Then Exit Sub
        Dim dtOpen As DataTable = DbHelper.GetDataTable("SELECT AC_CODE FROM Open_jn WHERE AC_CODE = '" & FG.CurrentRow.Cells("Col1").Value.ToString() & "'", RSC)
        If dtOpen.Rows.Count > 0 Then

            Cs = 1
        End If


        If Cs = 1 Then
            If MessageBox.Show("ລະຫັດ " & txtOldId.Text & "  ມີການເຄີ່ຶນໄຫວແລ້ວທ່ານຕ້ອງການປ່ຽນລະຫັດນີ້ ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

                DbHelper.ExecuteNonQuery("UPDATE gen_jn SET AC_CODE='" & txtNewId.Text & "',Code_dr= N'" & txtNewId.Text & "' WHERE AC_CODE = '" & txtOldId.Text & "' And Code_dr =  '" & txtOldId.Text & "'")

                DbHelper.ExecuteNonQuery("UPDATE gen_jn SET AC_CODE='" & txtNewId.Text & "',Code_Cr= N'" & txtNewId.Text & "' WHERE AC_CODE = '" & txtOldId.Text & "' And Code_Cr =  '" & txtOldId.Text & "'")

                DbHelper.ExecuteNonQuery("UPDATE Open_jn SET AC_CODE='" & txtNewId.Text & "',Code_dr= N'" & txtNewId.Text & "' WHERE AC_CODE = '" & txtOldId.Text & "' And Code_dr =  '" & txtOldId.Text & "'")
                DbHelper.ExecuteNonQuery("UPDATE Open_jn SET AC_CODE='" & txtNewId.Text & "',Code_Cr= N'" & txtNewId.Text & "' WHERE AC_CODE = '" & txtOldId.Text & "' And Code_Cr =  '" & txtOldId.Text & "'")

                DbHelper.ExecuteNonQuery("UPDATE Acc_Code SET AC_CODE='" & txtNewId.Text & "',acc_type= N'" & AcTypeLao & "',acc_typee='" & AcTypeEng & "' WHERE AC_CODE = '" & txtOldId.Text & "'")

                StartLoadDataList()
                Panel2.Visible = False

            End If

        Else

            If MessageBox.Show("ທ່ານຕ້ອງການປ່ຽນລະຫັດ " & txtOldId.Text & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

                DbHelper.ExecuteNonQuery("UPDATE gen_jn SET AC_CODE='" & txtNewId.Text & "',Code_dr= N'" & txtNewId.Text & "' WHERE AC_CODE = '" & txtOldId.Text & "' And Code_dr =  '" & txtOldId.Text & "'")
                DbHelper.ExecuteNonQuery("UPDATE gen_jn SET AC_CODE='" & txtNewId.Text & "',Code_Cr= N'" & txtNewId.Text & "' WHERE AC_CODE = '" & txtOldId.Text & "' And Code_Cr =  '" & txtOldId.Text & "'")

                DbHelper.ExecuteNonQuery("UPDATE Open_jn SET AC_CODE='" & txtNewId.Text & "',Code_dr= N'" & txtNewId.Text & "' WHERE AC_CODE = '" & txtOldId.Text & "' And Code_dr =  '" & txtOldId.Text & "'")
                DbHelper.ExecuteNonQuery("UPDATE Open_jn SET AC_CODE='" & txtNewId.Text & "',Code_Cr= N'" & txtNewId.Text & "' WHERE AC_CODE = '" & txtOldId.Text & "' And Code_Cr =  '" & txtOldId.Text & "'")

                DbHelper.ExecuteNonQuery("UPDATE Acc_Code SET AC_CODE='" & txtNewId.Text & "',acc_type= N'" & AcTypeLao & "',acc_typee='" & AcTypeEng & "' WHERE AC_CODE = '" & txtOldId.Text & "'")

                StartLoadDataList()
                Panel2.Visible = False

            End If

        End If







    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Panel1.Visible = False
        StartLoadDataList()
    End Sub
Public Sub PageCnt(ByVal StrSQL As String, ByVal ConStr As String, ByVal PageNum As Long, ByVal RowPerPage As Integer)
        'Me.Enabled = False
        ' Dim RsLoad As New ADODB.Recordset ' REMOVED - ADODB migration
        ' Dim rssum As New ADODB.Recordset ' REMOVED - ADODB migration
        Dim i As Integer
        FG.Rows.Clear()


    
        Panel3.Visible = False

        PageNum = PageNum - 1
        Dim dt As DataTable = DbHelper.GetDataTable("select *  from Acc_Code WHERE AC_CODE<>''  " & SQL & "order by AC_CODE")
        If dt.Rows.Count <> 0 Then
            Dim startRow As Integer = RowPerPage * PageNum
            If Int(dt.Rows.Count Mod RowPerPage) = 0 Then
                Last_page = Int(dt.Rows.Count / DividePage)
            Else
                Last_page = Int(dt.Rows.Count / DividePage) + 1
                If P = Last_page Then RowPerPage = (dt.Rows.Count Mod RowPerPage)
            End If
            
            For i = startRow To Math.Min(startRow + RowPerPage - 1, dt.Rows.Count - 1)
                Dim row As DataRow = dt.Rows(i)
                FG.Rows.Add(i, Trim(CStr(row("AC_CODE"))), _
                               Trim(CStr(row("Ac_Original"))), _
                           Trim(CStr(row("Name_L"))), _
                             Trim(CStr(row("Name_E"))), _
                               Trim(CStr(row("Acc_Type"))), _
                                 Trim(CStr(row("Acc_TypeE"))), _
                                   DbHelper.GetStr(row("Print_status")))
            Next
        End If

        FirstPage.Enabled = True
        BackPage.Enabled = True
        NextPage.Enabled = True
        LasthPage.Enabled = True
        If P = 1 Then
            FirstPage.Enabled = False
            BackPage.Enabled = False
            NextPage.Enabled = True
            LasthPage.Enabled = True
        ElseIf P = Last_page Then
            FirstPage.Enabled = True
            BackPage.Enabled = True
            NextPage.Enabled = False
            LasthPage.Enabled = False
        End If

    End Sub

    Public Sub LoadSQL()

        SQL = ""
        If RdName.Checked = True Then
            If ChbLang.Checked = True Then
                If Rdlasth.Checked = True Then
                    SQL = SQL & " AND left(Name_E, N'" & Len(txtSearchName.Text.Trim) & "')= N'" & txtSearchName.Text.Trim & "' "
                End If

                If similar.Checked = True Then
                    SQL = SQL & " AND (Name_E  Like N'%" & txtSearchName.Text.Trim & "%')"
                End If
            Else
                '********************
                If Rdlasth.Checked = True Then
                    SQL = SQL & " AND left(Name_L, N'" & Len(txtSearchName.Text.Trim) & "')= N'" & txtSearchName.Text.Trim & "' "
                End If

                If similar.Checked = True Then
                    SQL = SQL & " AND (Name_L  Like N'%" & txtSearchName.Text.Trim & "%')"
                End If
            End If
        End If


        If RdId.Checked = True Then

            SQL = SQL & " AND left(AC_CODE, '" & Len(txtSearchId.Text.Trim) & "')= '" & txtSearchId.Text.Trim & "' "
            'SQL = SQL & " AND AC_CODE = '" & txtSearchId.Text & "' "
        End If
        If CheckBox2.Checked = True Then
            SQL = ""
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Panel1.Visible = False
        Panel3.Visible = False
    End Sub


    Private Sub ChbLang_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChbLang.CheckedChanged
        If ChbLang.Checked = True Then
            ChbLang.Text = "Lao"
        Else
            ChbLang.Text = "Englisth"

        End If
    End Sub

    Private Sub RdId_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RdId.CheckedChanged
        If RdId.Checked = True Then
            txtSearchId.Enabled = True
            txtSearchName.Enabled = False
            ChbLang.Enabled = False
            txtSearchId.Focus()
            Rdlasth.Enabled = False
            similar.Enabled = False
        Else
            Rdlasth.Enabled = True
            similar.Enabled = True
            txtSearchId.Enabled = False
            txtSearchName.Enabled = True
            ChbLang.Enabled = True
            txtSearchName.Focus()
        End If
    End Sub

    Private Sub RdName_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RdName.CheckedChanged
        If RdId.Checked = True Then
            txtSearchId.Enabled = True
            txtSearchName.Enabled = False
            ChbLang.Enabled = False
            txtSearchId.Focus()
            Rdlasth.Enabled = False
            similar.Enabled = False
        Else
            Rdlasth.Enabled = True
            similar.Enabled = True
            txtSearchId.Enabled = False
            txtSearchName.Enabled = True
            ChbLang.Enabled = True
            txtSearchName.Focus()
        End If
    End Sub

    Private Sub txtSearchId_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSearchId.KeyPress
        If e.KeyChar = Chr(13) Then
            StartLoadDataList()
            Panel1.Visible = False
        End If
    End Sub

    Private Sub txtSearchId_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchId.TextChanged

    End Sub

    Private Sub txtSearchName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSearchName.KeyPress
        If e.KeyChar = Chr(13) Then
            StartLoadDataList()
            Panel1.Visible = False
        End If
    End Sub


    Private Sub BtnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & " 'As Crl_Lng  ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
        LngId = "7002" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDP ,"
        LngId = "7035" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
        LngId = "7004" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Descrip ,"
        LngId = "7005" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_No ,"
        LngId = "7006" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Ac_Code ,"
        LngId = "7011" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Sign ,"
        LngId = "7012" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_SignNme ,"
        LngId = "7026" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InDate ,"
        LngId = "7036" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Len ,"
        LngId = "7037" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_AcType ,"
        LngId = "7038" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_AcNme ,"
        Call LoadLoGO()
        ' Dim Rs As New ADODB.Recordset ' COMMENTED OUT - ADODB migration
        Dim dt As DataTable = DbHelper.GetDataTable("SELECT " & MuLngRpt & " * FROM Acc_Code WHERE 1=1 " & SQL & "  And Ac_Code<>'' Order by Left(Ac_Code,3) ,Left(Ac_Code,4) , Left(Ac_Code,5), Left(Ac_Code,6) , Left(Ac_Code,7)")
        If dt.Rows.Count = 0 Then MsgBox("Data empty", vbInformation, "Check") : Exit Sub
            If CheckBox3.Checked = True Then
                Dim Frm As New FmPreview
                Dim Rpt As New CryShartOfAcc2
                Rpt.SetDataSource(dt)
                Frm.ReportViewer.ReportSource = Rpt
                Frm.ReportViewer.DisplayGroupTree = False
                Frm.WindowState = FormWindowState.Maximized
                Frm.Show()
                Rpt = Nothing
            Else
                Dim FrmPreview As New FmPreview
                Dim Rpt As New CryShartOfAcc
                If MdShowLOGO = 1 Then
                    Rpt.Subreports(0).SetDataSource(RsLOGO)
                End If
                Rpt.SetDataSource(dt)
                FrmPreview.ReportViewer.ReportSource = Rpt
                FrmPreview.ReportViewer.DisplayGroupTree = False
                FrmPreview.MdiParent = FmMain
                FrmPreview.WindowState = FormWindowState.Maximized
                FrmPreview.Show()
                FrmPreview.Focus()
            End If
    End Sub

Private Sub CmbPrinSelete_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbPrinSelete.SelectedIndexChanged
        'CNN.Execute("UPDATE Acc_Code SET Print_status='0'")
        For k = 0 To FG.Rows.Count - 1
            FG.Rows(k).Cells("Col7").Value = False
        Next k
        
        For J = 0 To FG.Rows.Count - 1
            If Convert.ToBoolean(FG.Rows(J).Cells("Col7").Value) = True Then
                For I = 0 To FG.Columns.Count - 1
                    FG.Rows(J).Cells(I).Style.BackColor = Color.SkyBlue
                Next I
            Else
                For I = 0 To FG.Columns.Count - 1
                    FG.Rows(J).Cells(I).Style.BackColor = Color.White
                Next
            End If
        Next J
        
        If CmbPrinSelete.Text = "ເລືອກລາຍການພິມ (Prin All Items)" Then
            CheckBox1.Checked = False
            CheckBox1.Enabled = False
        Else
            CheckBox1.Enabled = True
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        a = 0
        Dim dtPrint As DataTable = DbHelper.GetDataTable("select * from Acc_Code where Print_status = '" & "1" & "'")
        If dtPrint.Rows.Count > 0 Then a = 1

        SQL = ""
        If CmbPrinSelete.Text = "ເລືອກລາຍການພິມ (Prin All Items)" Then
            SQL = " AND Print_status = '" & "1" & "' "

            If a = 0 Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
        End If


        ' Dim Rs As New ADODB.Recordset ' COMMENTED OUT - ADODB migration
        Dim dtReport As DataTable = DbHelper.GetDataTable("SELECT * FROM Acc_Code WHERE 1=1 " & SQL & " ")
        If dtReport.Rows.Count = 0 Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim Rpt As New CryShartOfAcc
        Rpt.SetDataSource(dtReport)
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ShowDialog()
        FrmPreview.Focus()
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

If FG.CurrentRow Is Nothing Then Exit Sub
        If IsNumeric(FG.CurrentRow.Cells("Col1").Value.ToString()) = False Then
            MsgBox("S")
        Else
            MsgBox("0")
        End If
    End Sub

    Private Sub LasthPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LasthPage.Click
        Call LoadDividePage()
        P = Last_page
        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = Last_page & "/" & Last_page
        If CmbPage.SelectedIndex >= 0 Then
            CmbPage.SelectedIndex = P - 1
        End If
    End Sub

    Private Sub NextPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NextPage.Click
        Call LoadDividePage()
        If P >= Last_page Then Exit Sub
        P = P + 1
        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = P & "/" & Last_page
        If CmbPage.SelectedIndex >= 0 Then
            CmbPage.SelectedIndex = P - 1
        End If
    End Sub

    Private Sub BackPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackPage.Click
        Call LoadDividePage()
        If P = 1 Then Exit Sub
        P = P - 1

        Call LoadSQL()

        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = P & "/" & Last_page
        If CmbPage.SelectedIndex >= 0 Then
            CmbPage.SelectedIndex = P - 1
        End If
    End Sub

    Private Sub FirstPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FirstPage.Click

        'If No_record = True Then Exit Sub
        Call LoadDividePage()
        P = 1
        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page

        If CmbPage.SelectedIndex >= 0 Then
            CmbPage.SelectedIndex = P - 1
        End If

    End Sub

    Private Sub p25_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles p25.CheckedChanged
        'FirstPage.Enabled = False
        'BackPage.Enabled = False
        'LasthPage.Enabled = False
        'NextPage.Enabled = False




    End Sub

    Private Sub p50_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles p50.CheckedChanged
        'FirstPage.Enabled = False
        'BackPage.Enabled = False
        'LasthPage.Enabled = False
        'NextPage.Enabled = False
        'Call LoadDividePage()
        'P = 1

        'Call LoadSQL()
        'Call PageCnt(StrSQL, ConString, P, DividePage)
        'Me.lblpage_total.Text = "1/" & Last_page

    End Sub

    Private Sub p10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles p100.CheckedChanged


    End Sub

    Private Sub p250_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles p250.CheckedChanged





        'FirstPage.Enabled = False
        'BackPage.Enabled = False
        'LasthPage.Enabled = False
        'NextPage.Enabled = False

    End Sub

    Private Sub p500_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles p500.CheckedChanged



    End Sub

    Private Sub p100_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles p1000.CheckedChanged

    End Sub

    Private Sub GrPage_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GrPage.Enter

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub p250_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p250.MouseClick

        p25.ForeColor = Color.Black
        p50.ForeColor = Color.Black
        p250.ForeColor = Color.Black
        p500.ForeColor = Color.Black
        p1000.ForeColor = Color.Black
        If p25.Checked = True Then
            p25.ForeColor = Color.Red
        ElseIf p50.Checked = True Then
            p50.ForeColor = Color.Red
        ElseIf p250.Checked = True Then
            p250.ForeColor = Color.Red
        ElseIf p500.Checked = True Then
            p500.ForeColor = Color.Red
        ElseIf p1000.Checked = True Then
            p1000.ForeColor = Color.Red
        End If
        Call LoadDividePage()
        P = 1

        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page







        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i

        'Call LoadDividePage()
        'P = 1
        'CmbPage.SelectedIndex = 0
        'Call LoadSQL()
        'Call PageCnt(StrSQL, ConString, P, DividePage)
        'Me.lblpage_total.Text = "1/" & Last_page
    End Sub

    Private Sub p25_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p25.MouseClick
        p25.ForeColor = Color.Black
        p50.ForeColor = Color.Black
        p250.ForeColor = Color.Black
        p500.ForeColor = Color.Black
        p1000.ForeColor = Color.Black
        If p25.Checked = True Then
            p25.ForeColor = Color.Red
        ElseIf p50.Checked = True Then
            p50.ForeColor = Color.Red
        ElseIf p250.Checked = True Then
            p250.ForeColor = Color.Red
        ElseIf p500.Checked = True Then
            p500.ForeColor = Color.Red
        ElseIf p1000.Checked = True Then
            p1000.ForeColor = Color.Red
        End If
        Call LoadDividePage()
        P = 1

        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page
        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        'CmbPage.SelectedIndex = 0
    End Sub

    Private Sub p50_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p50.MouseClick
        p25.ForeColor = Color.Black
        p50.ForeColor = Color.Black
        p250.ForeColor = Color.Black
        p500.ForeColor = Color.Black
        p1000.ForeColor = Color.Black
        If p25.Checked = True Then
            p25.ForeColor = Color.Red
        ElseIf p50.Checked = True Then
            p50.ForeColor = Color.Red
        ElseIf p250.Checked = True Then
            p250.ForeColor = Color.Red
        ElseIf p500.Checked = True Then
            p500.ForeColor = Color.Red
        ElseIf p1000.Checked = True Then
            p1000.ForeColor = Color.Red
        End If
        Call LoadDividePage()
        P = 1

        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page

        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        'CmbPage.SelectedIndex = 0
    End Sub

    Private Sub p1000_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p1000.MouseClick
        p25.ForeColor = Color.Black
        p50.ForeColor = Color.Black
        p250.ForeColor = Color.Black
        p500.ForeColor = Color.Black
        p1000.ForeColor = Color.Black
        If p25.Checked = True Then
            p25.ForeColor = Color.Red
        ElseIf p50.Checked = True Then
            p50.ForeColor = Color.Red
        ElseIf p250.Checked = True Then
            p250.ForeColor = Color.Red
        ElseIf p500.Checked = True Then
            p500.ForeColor = Color.Red
        ElseIf p1000.Checked = True Then
            p1000.ForeColor = Color.Red
        End If
        Call LoadDividePage()
        P = 1

        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page
        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        'CmbPage.SelectedIndex = 0
    End Sub

    Private Sub p500_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p500.MouseClick
        p25.ForeColor = Color.Black
        p50.ForeColor = Color.Black
        p250.ForeColor = Color.Black
        p500.ForeColor = Color.Black
        p1000.ForeColor = Color.Black
        If p25.Checked = True Then
            p25.ForeColor = Color.Red
        ElseIf p50.Checked = True Then
            p50.ForeColor = Color.Red
        ElseIf p250.Checked = True Then
            p250.ForeColor = Color.Red
        ElseIf p500.Checked = True Then
            p500.ForeColor = Color.Red
        ElseIf p1000.Checked = True Then
            p1000.ForeColor = Color.Red
        End If
        Call LoadDividePage()
        P = 1

        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page


        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        'CmbPage.SelectedIndex = 0
    End Sub

    Private Sub p100_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p100.MouseClick
        p25.ForeColor = Color.Black
        p50.ForeColor = Color.Black
        p250.ForeColor = Color.Black
        p500.ForeColor = Color.Black
        p1000.ForeColor = Color.Black
        If p25.Checked = True Then
            p25.ForeColor = Color.Red
        ElseIf p50.Checked = True Then
            p50.ForeColor = Color.Red
        ElseIf p250.Checked = True Then
            p250.ForeColor = Color.Red
        ElseIf p500.Checked = True Then
            p500.ForeColor = Color.Red
        ElseIf p1000.Checked = True Then
            p1000.ForeColor = Color.Red
        End If
        Call LoadDividePage()
        P = 1

        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page

        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        'CmbPage.SelectedIndex = 0
    End Sub

    Private Sub CmbPage_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles CmbPage.MouseDown


    End Sub

    Private Sub CmbPage_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbPage.SelectedIndexChanged
        If CmbPage.Text <> "" Then
            P = CDbl(CmbPage.Text)
            Call LoadDividePage()
            'If P >= Last_page Then Exit Sub

            Call LoadSQL()
            Call PageCnt(StrSQL, ConString, P, DividePage)
            Me.lblpage_total.Text = P & "/" & Last_page
        End If

    End Sub

    Private Sub Button8_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        StartLoadDataList()
    End Sub
    Public Sub StartLoadDataList()
        p25.ForeColor = Color.Black
        p50.ForeColor = Color.Black
        p250.ForeColor = Color.Black
        p500.ForeColor = Color.Black
        p1000.ForeColor = Color.Black
        If p25.Checked = True Then
            p25.ForeColor = Color.Red
        ElseIf p50.Checked = True Then
            p50.ForeColor = Color.Red
        ElseIf p100.Checked = True Then
            p100.ForeColor = Color.Red
        ElseIf p250.Checked = True Then
            p250.ForeColor = Color.Red
        ElseIf p500.Checked = True Then
            p500.ForeColor = Color.Red
        ElseIf p1000.Checked = True Then
            p1000.ForeColor = Color.Red
        End If
        Call LoadDividePage()
        P = 1

        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page
        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        If CmbPage.SelectedIndex >= 0 Then
            'CmbPage.SelectedIndex = 0
        End If
    End Sub

    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtLng_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLng.TextChanged
        'If MuLng = "L" Then
        '    FG.set_ColHidden(3, True)
        '    FG.set_ColHidden(5, True)
        '    FG.set_ColHidden(2, False)
        '    FG.set_ColHidden(4, False)
        'ElseIf MuLng = "E" Then
        '    FG.set_ColHidden(3, False)
        '    FG.set_ColHidden(5, False)
        '    FG.set_ColHidden(2, True)
        '    FG.set_ColHidden(4, True)
        'End If
    End Sub
End Class