Public NotInheritable Class FmOpen_jn_List
    Dim RptNme As String
    Dim Ac_Code As String
    Dim sql As String
    Dim mylock As Integer = 0

    ' DataGridView Helper Methods
    Private Function GetGridValue(ByVal grid As DataGridView, ByVal row As Integer, ByVal col As Integer) As String
        If row >= 0 AndAlso row < grid.RowCount AndAlso col >= 0 AndAlso col < grid.ColumnCount Then
            If grid.Rows(row).Cells(col).Value IsNot Nothing Then
                Return grid.Rows(row).Cells(col).Value.ToString()
            End If
        End If
        Return ""
    End Function

    Private Sub SetGridValue(ByVal grid As DataGridView, ByVal row As Integer, ByVal col As Integer, ByVal value As String)
        If row >= 0 AndAlso row < grid.RowCount AndAlso col >= 0 AndAlso col < grid.ColumnCount Then
            grid.Rows(row).Cells(col).Value = value
        End If
    End Sub
    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Close()
    End Sub
    Private Sub BntNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BntNew.Click
        FmNewOpen_jn.txtCode_dr.Enabled = True
        FmNewOpen_jn.txtCode_cr.Enabled = True
        FmNewOpen_jn.BtnSearch_dr.Enabled = True
        FmNewOpen_jn.BtnSearch_cr.Enabled = True
        FmNewOpen_jn.ShowDialog()
    End Sub

    Private Sub SetupGrid()
        FG.AllowUserToAddRows = False
        FG.AllowUserToDeleteRows = False
        FG.ReadOnly = True
        FG.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG.MultiSelect = False
        FG.RowHeadersVisible = False
        
        FG.Columns.Clear()
        FG.Columns.Add("Col0", "ລ/ດ")
        FG.Columns.Add("Col1", "ວັນທີ")
        FG.Columns.Add("Col2", "ເລກບັນຊີເບືອງໜີ")
        FG.Columns.Add("Col3", "ເລກບັນຊີເບືອງມີ")
        FG.Columns.Add("Col4", "ເລກບັນຊີ")
        FG.Columns.Add("Col5", "ຊືບັນຊີ(ລາວ)")
        FG.Columns.Add("Col6", "ຈຳນວນເງິນຈົດໜີ້")
        FG.Columns.Add("Col7", "ຈຳນວນເງິນຈົດມີ")
        FG.Columns.Add("Col8", "ເງິນ")
        FG.Columns.Add("Col9", "ອັດຕາ")
        FG.Columns.Add("Col10", "ມູນຄ່າໜີ້")
        FG.Columns.Add("Col11", "ມູນຄ່າມີ")
        FG.Columns.Add("Col12", "ສາຂາ")
        FG.Columns.Add("Col13", "cnt")
        FG.Columns.Add("Col14", "Ac_original")
        
        For Each col As DataGridViewColumn In FG.Columns
            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next
    End Sub

    Private Sub FmOpen_jn_List_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Call MdiCNum()
    End Sub
  
    Private Sub FmOpen_jn_List_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'MsgBox(Year(MWorkSetting))

        SetupGrid()


        'ChAllSty.Visible = False


        'MsgBox(FmLogin.Sub_Company.Text)

        FG.AllowUserToResizeColumns = True
        FG.AllowUserToResizeRows = True
        sql = " AND code_cr, '" & ""
        'LoadListFG()

        'loadCompany()
        If ChAllSty.Checked = True Then
            BntNew.Enabled = False
            BtnEdit.Enabled = False
            BtnDelete.Enabled = False
            Button1.Enabled = False
            FmNewOpen_jn.BntNew.Enabled = False
            FmNewOpen_jn.BtnSave.Enabled = False




        Else



            BntNew.Enabled = True
            BtnEdit.Enabled = True
            BtnDelete.Enabled = True
            Button1.Enabled = True
            FmNewOpen_jn.BntNew.Enabled = True
            FmNewOpen_jn.BtnSave.Enabled = True


            If Mid(FmLogin.Sub_Company.Text, 4, 2) = "00" Then
                BntNew.Enabled = False
                BtnEdit.Enabled = False
                BtnDelete.Enabled = False
                Button1.Enabled = False
                FmNewOpen_jn.BntNew.Enabled = False
                FmNewOpen_jn.BtnSave.Enabled = False
            End If
        End If

        Ds.Text = MWorkSetting
        yy.Value = Ds.Value
        'Year(MWorkSetting)
        'MsgBox("77")
        SetControlText(Me)

        Call loadOffice_User()
        Button1.Text = "ລ໋ອກ"
        Button2.Text = "ປົດລ໋ອກ"
        CheckBox1.Text = "ພາສາ"
        SetupGrid()

    End Sub

    Public Sub LoadLook()
        If MPermit = "Admin" Then
            MULook2 = ""
            If Mid(FmLogin.Sub_Company.Text, 1, 5) = "00-00" Then
                MULook2 = ""
            Else
                If Mid(FmLogin.Sub_Company.Text, 4, 2) = "00" Then
                    MULook2 = " AND (Open_jn.company  Like N'" & Mid(FmLogin.Sub_Company.Text, 1, 2) & "%')"
                Else
                    MULook2 = " AND Open_jn.company = '" & Mid(FmLogin.Sub_Company.Text, 1, 5) & "' "
                End If
            End If
        End If
        If MPermit = "Sub-Admin" Then
            If Mid(FmLogin.Sub_Company.Text, 4, 2) = "00" Then
                MULook2 = " AND (Open_jn.company  Like N'" & Mid(FmLogin.Sub_Company.Text, 1, 2) & "%')"
            Else
                MULook2 = " AND Open_jn.company = '" & Mid(FmLogin.Sub_Company.Text, 1, 5) & "' "
            End If
        End If
        If MPermit = "Border-Admin" Then
            MULook2 = " AND Open_jn.company = '" & Mid(FmLogin.Sub_Company.Text, 1, 5) & "' "
        End If

        If MPermit = "User" Then
            MULook2 = " AND Open_jn.Last_User = N'" & MUserName & "' "
        End If
        '===============
    End Sub

    Private Sub loadOffice_User()
        Off_Usr.Items.Clear()
        Dim dtOffice As DataTable = DbHelper.GetDataTable("select sub_id , off_add2  from  Ap_office  Order by sub_id")
        If dtOffice.Rows.Count > 0 Then
            For Each row As DataRow In dtOffice.Rows
                Off_Usr.Items.Add((row("sub_id").ToString()) & " " & (row("off_add2").ToString()))
            Next
        End If
        Off_Usr.Text = FmLogin.Sub_Company.Text
    End Sub

    Public Sub LoadListFG()
        Dim OfUsr1 As String = Microsoft.VisualBasic.Left(Off_Usr.Text, 5)
        Dim OfUsr2 As String = Mid(Off_Usr.Text, 4, 2)
        Dim OfUsr3 As String = Microsoft.VisualBasic.Left(Off_Usr.Text, 2)
        If OfUsr1 = "00-00" Then
            MULook2 = ""
        Else
            If OfUsr2 = "00" Then
                MULook2 = "  And  Left(Open_jn.company,2)= '" & OfUsr3 & "' "
            Else
                MULook2 = "  And Open_jn.company= '" & OfUsr1 & "' "
            End If
        End If
        If ChAllSty.Checked = False Then
            FG.Rows = 1
            FG.Cols = 14
            FG.set_ColHidden(4, True)
            FG.set_ColHidden(1, False)
        SetupGrid()

                Dim s As String = "SELECT * FROM  Open_jn   where  Open_jn.date_work   BETWEEN '" & Format(yy.Value, "yyyy") & "-1-1" & "' AND '" & Format(yy.Value, "yyyy") & "-1-1" & "' " & MULook2 & "  And  Open_jn.ac_code Like '%" & TextBox1.Text & "%' Order by ac_code "
                Dim dtOpenJn As DataTable = DbHelper.GetDataTable(s)
                If dtOpenJn.Rows.Count > 0 Then
            For Each dr As DataRow In dtOpenJn.Rows
                Dim row As Integer = FG.Rows.Add()
                FG.Rows(row).Cells(0).Value = dr("date_work")
                FG.Rows(row).Cells(1).Value = Format(CDate(dr("date_work").ToString()), "dd/MM/yyyy")
                FG.Rows(row).Cells(2).Value = Trim(CStr(dr("code_dr").ToString()))
                FG.Rows(row).Cells(3).Value = Trim(CStr(dr("code_cr").ToString()))
                FG.Rows(row).Cells(4).Value = Trim(CStr(dr("ac_code").ToString()))
                FG.Rows(row).Cells(5).Value = Trim(CStr(dr("ac_name").ToString()))
                FG.Rows(row).Cells(6).Value = Trim(CStr(dr("ac_namee").ToString()))
                ' Continue with remaining fields if they exist
                If dr.Table.Columns.Contains("amount_dr") Then
                    FG.Rows(row).Cells(7).Value = Format(CDbl(Trim(dr("amount_dr").ToString())), "##,##0.00")
                End If
                If dr.Table.Columns.Contains("amount_cr") Then
                    FG.Rows(row).Cells(8).Value = Format(CDbl(Trim(dr("amount_cr").ToString())), "##,##0.00")
                End If
                If dr.Table.Columns.Contains("curr") Then
                    FG.Rows(row).Cells(9).Value = Trim(CStr(dr("curr").ToString()))
                End If
                If dr.Table.Columns.Contains("rate") Then
                    FG.Rows(row).Cells(10).Value = Format(CDbl(Trim(dr("rate").ToString())), "##,##0.00")
                End If
                If dr.Table.Columns.Contains("amt_dr") Then
                    FG.Rows(row).Cells(11).Value = Format(CDbl(Trim(dr("amt_dr").ToString())), "##,##0.00")
                End If
                If dr.Table.Columns.Contains("amt_cr") Then
                    FG.Rows(row).Cells(12).Value = Format(CDbl(Trim(dr("amt_cr").ToString())), "##,##0.00")
                End If
                If dr.Table.Columns.Contains("Company") Then
                    FG.Rows(row).Cells(13).Value = Trim(CStr(dr("Company").ToString()))
                End If
                If dr.Table.Columns.Contains("CNT") Then
                    FG.Rows(row).Cells(14).Value = Trim(CStr(dr("CNT").ToString()))
                End If
                If dr.Table.Columns.Contains("Ac_original") Then
                    FG.Rows(row).Cells(15).Value = Trim(CStr(dr("Ac_original").ToString()))
                End If
            Next
        Else
            ' Handle the case where no data was found
        End If
            '=====
        Else
            SetupGrid()
            FG.Rows.Clear()
            ' DataGridView handles columns automatically based on SetupGrid()
            FG.Columns("Col4").Visible = True
            FG.Columns("Col1").Visible = False

            'Call LoadData("select *  from Open_jn WHERE ac_code<>''  " & Sql & "order by ac_code", RSC)
            'MsgBox(MULookSelct)
            Dim s As String = " SELECT   sum(Open_jn.amt_dr) as amt_dr , sum(Open_jn.amt_cr) as amt_cr ,   Open_jn.ac_code AS ac_code , Acc_Code.Name_L FROM         Open_jn INNER JOIN   Acc_Code ON Open_jn.ac_code = Acc_Code.Ac_Code   where  Open_jn.date_work   BETWEEN '" & Year(MWorkSetting) & "-1-1" & "' AND '" & Year(MWorkSetting) & "-1-1" & "' " & MULook2 & " group by Open_jn.ac_code, Acc_Code.Name_L   having   SUM((Open_jn.amt_cr)-(Open_jn.amt_dr)) <> 0  And  Open_jn.ac_code Like '%" & TextBox1.Text & "%'  Order by Open_jn.ac_code "
            Dim dtSummary As DataTable = DbHelper.GetDataTable(s)
            If dtSummary.Rows.Count > 0 Then
                For Each dr As DataRow In dtSummary.Rows
                    Dim Ac_Code, Code_Dr, Code_Cr, Amt_Dr, Amt_Cr As String
                    Ac_Code = Trim(CStr(dr("ac_code").ToString()))
                    Amt_Dr = Format(CDbl(Trim(dr("amt_dr").ToString())), "##,##0.00")
                    Amt_Cr = Format(CDbl(Trim(dr("amt_cr").ToString())), "##,##0.00")
                    If CDbl(Trim(dr("amt_dr").ToString())) > CDbl(Trim(dr("amt_cr").ToString())) Then
                        Code_Dr = Ac_Code
                        Code_Cr = ""
                        Amt_Dr = Format(CDbl(CDbl(Trim(dr("amt_dr").ToString())) - CDbl(Trim(dr("amt_Cr").ToString()))), "##,##0.00")
                        Amt_Cr = "0.00"
                    Else
                        Code_Cr = Ac_Code
                        Code_Dr = ""
                        Amt_Cr = Format(CDbl(CDbl(Trim(dr("amt_Cr").ToString())) - CDbl(Trim(dr("amt_dr").ToString()))), "##,##0.00")
                        Amt_Dr = "0.00"
                    End If

                    If Amt_Dr <> Amt_Cr Then
                        Dim row As Integer = FG.Rows.Add()
                        FG.Rows(row).Cells(0).Value = FG.Rows.Count ' Equivalent to .AbsolutePosition
                        FG.Rows(row).Cells(1).Value = Ac_Code
                        FG.Rows(row).Cells(2).Value = Code_Dr
                        FG.Rows(row).Cells(3).Value = Code_Cr
                        FG.Rows(row).Cells(4).Value = Trim(CStr(dr("Name_L").ToString()))
                        FG.Rows(row).Cells(5).Value = Amt_Dr
                        FG.Rows(row).Cells(6).Value = Amt_Cr
                    End If
                Next
            Else
                ' Handle the case where no data was found
            End If
        End If

        Call SumAmountDr()
        Call loadColor()

        'Ac_Nme.Text = "SELECT * FROM  Open_jn   where  Open_jn.date_work   BETWEEN '" & Format(dts.Value, "yyyy-MM-dd") & "' AND '" & Format(dtt.Value, "yyyy-MM-dd") & "' " & MULook & ""

    End Sub

    Private Sub Load_For_Edit()
        'MsgBox("asd")
        FmNewOpen_jn.txtCode_dr.Text = GetGridValue(FG, FG.CurrentRow.Index, 2)
        FmNewOpen_jn.txtCode_cr.Text = GetGridValue(FG, FG.CurrentRow.Index, 3)
        'FmNewOpen_jn.Label7.Text = GetGridValue(FG, FG.CurrentRow.Index, 13)
        FmNewOpen_jn.txtCode_dr.Enabled = False
        FmNewOpen_jn.txtCode_cr.Enabled = False
        FmNewOpen_jn.BtnSearch_dr.Enabled = False
        FmNewOpen_jn.BtnSearch_cr.Enabled = False
    End Sub


    Public Sub SumAmountDr()
        Dim i As Integer
        Dim AmountDr, AmountCr As Double
        AmountDr = 0
        AmountCr = 0
        For i = 1 To FG.Rows - 1
            If ChAllSty.Checked = False Then
                AmountDr = AmountDr + CDbl(GetGridValue(FG, i, 11))
                AmountCr = AmountCr + CDbl(GetGridValue(FG, i, 12))
            Else
                AmountDr = AmountDr + CDbl(GetGridValue(FG, i, 5))
                AmountCr = AmountCr + CDbl(GetGridValue(FG, i, 6))
            End If
        Next i
        'MsgBox(TotalAmountCr)

        txtSumAmountDr.Text = Format(AmountDr, "#,##0.00")
        txtSumAmountCr.Text = Format(AmountCr, "#,##0.00")

        BalanceDr.Text = CDbl(txtSumAmountCr.Text) - CDbl(txtSumAmountDr.Text)
        BalanceDr.Text = Format(CDbl(BalanceDr.Text), "#,##0.00")
        'Dr.Text = CDbl(txtSumAmountDr.Text) - CDbl(txtSumAmountCr.Text)
        'Cr.Text = CDbl(txtSumAmountCr.Text) - CDbl(txtSumAmountDr.Text)
        'DDR.Text = CDbl(txtSumTotalAmountDr.Text) - CDbl(txtSumTotalAmountCr.Text)
        'CCR.Text = CDbl(txtSumTotalAmountCr.Text) - CDbl(txtSumTotalAmountDr.Text)
        'Dr.Text = Format(CDbl(Dr.Text), "#,##0.00")
        'Cr.Text = Format(CDbl(Cr.Text), "#,##0.00")
        'DDR.Text = Format(CDbl(DDR.Text), "#,##0.00")
        'CCR.Text = Format(CDbl(CCR.Text), "#,##0.00")

    End Sub

    Public Sub loadColor()
        'Dim rmRS As New ADODB.Recordset
        'Dim J As Integer
        'FG.Redraw = False
        'For J = 1 To FG.Rows - 1
        '    If Trim(FG.get_TextMatrix(J, 13)) = True Then
        '        FG.Row = J
        '        FG.Col = 13
        '        FG.CellBackColor = Color.Blue
        '    End If
        'Next J
        FG.Redraw = True
    End Sub
    Private Sub BtnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnEdit.Click
        If Ac_Code = "" Then Exit Sub
        mylock = 0 : LockAcc() : If mylock = 1 Then Exit Sub
        If MdAtv = False Or GetGridValue(FG, FG.CurrentRow.Index, 1) = "" Then MessageBox.Show("ກະລຸນນາເລືອກກ່ອນ") : Exit Sub
        'If FG.get_TextMatrix(FG.Row, 13L) = True Then MessageBox.Show("ຂໍ້ມູນນີ້ໄດຖືກລ໋ອດແລ້ວບໍ່ສາມາດແກ້ໄຂໄດ້") : Exit Sub
        FmNewOpen_jn.ShowDialog()
    End Sub
    Private Sub LockAcc()
        Dim dtLockCheck As DataTable = DbHelper.GetDataTable("SELECT Date_work FROM  Open_jn where  My_Lock =1 And ac_code='" & GetGridValue(FG, FG.CurrentRow.Index, 4) & "' And Company='" & GetGridValue(FG, FG.CurrentRow.Index, 13) & "' And year(date_work)= '" & Format(CDate(GetGridValue(FG, FG.CurrentRow.Index, 1)), "yyyy") & "'   ")
            If dtLockCheck.Rows.Count <> 0 Then
                MsgBox("ບັນຊີປີ '" & CDbl(yy.Text) & "' ໄດ້ລ໋ອດໄວ້ແລ້ວທ່ານຕ້ອງປົດລ໋ອດກ່ອນ")
                mylock = 1
            End If

    End Sub
    Private Sub BtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelete.Click
        mylock = 0 : LockAcc() : If mylock = 1 Then Exit Sub
        If MdAtv = False Or GetGridValue(FG, FG.CurrentRow.Index, 1) = "" Then MessageBox.Show("ກະລຸນນາເລືອກກ່ອນ") : Exit Sub
        'If GetGridValue(FG, FG.CurrentRow.Index, 13) = True Then MessageBox.Show("ຂໍ້ມູນນີ້ໄດຖືກລ໋ອດແລ້ວບໍ່ສາມາດແກ້ໄຂໄດ້") : Exit Sub
        If MdAtv = False Then MessageBox.Show("ກະລຸນນາເລືອກກ່ອນ") : Exit Sub
        If MessageBox.Show("ທ່ານຕ້ອງລຶບລະຫັດ " & GetGridValue(FG, FG.CurrentRow.Index, 4) & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            CNN.Execute("DELETE FROM Open_jn WHERE ac_code='" & GetGridValue(FG, FG.CurrentRow.Index, 4) & "' And Company='" & GetGridValue(FG, FG.CurrentRow.Index, 13) & "' And cnt='" & GetGridValue(FG, FG.CurrentRow.Index, 14) & "'  And year(date_work)= '" & Format(CDate(GetGridValue(FG, FG.CurrentRow.Index, 1)), "yyyy") & "' ")
        End If
        MdAtv = False
        LoadListFG()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If MdAtv = False Then MessageBox.Show("ກະລຸນນາເລືອກກ່ອນ") : Exit Sub
        If MessageBox.Show("ທ່ານຕ້ອງການລ໋ອກປີ " & yy.Text & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            CNN.Execute("UPDATE Open_jn SET  my_lock = '" & "1" & "' WHERE Year(Date_Work) = '" & yy.Text & "'")
        End If
        MdAtv = False
        LoadListFG()

    End Sub
    Private Sub FG_CellDoubleClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FG.CellDoubleClick

        'MessageBox.Show(fg.get_TextMatrix(fg.Row,FG.Col))




        'If GetGridValue(FG, FG.CurrentRow.Index, 13) = True Then MessageBox.Show("ຂໍ້ມູນນີ້ໄດຖືກລ໋ອດແລ້ວບໍ່ສາມາດແກ້ໄຂໄດ້") : Exit Sub


        Load_For_Edit()
        mylock = 0 : LockAcc() : If mylock = 1 Then Exit Sub
        'MsgBox("sdf")
        If ChAllSty.Checked = False Then
            FmNewOpen_jn.ShowDialog()
        End If

    End Sub

    Private Sub FG_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelectionChanged
        'If GetGridValue(FG, FG.CurrentRow.Index, 13) <> MuSubOff Then
        '    BtnEdit.Enabled
        'End If
        MdAtv = True
        'MsgBox(Format(CDate(GetGridValue(FG, FG.CurrentRow.Index, 1)), "yyyy"))
        Ac_Code = GetGridValue(FG, FG.CurrentRow.Index, 4)
        'FmNewOpen_jn.DtmYearDate.Value = Format(CDate(GetGridValue(FG, FG.CurrentRow.Index, 1)), "yyyy")
        'MsgBox(FmNewOpen_jn.DtmYearDate.Text)
        FmNewOpen_jn.txtCode_dr.Text = GetGridValue(FG, FG.CurrentRow.Index, 2)
        FmNewOpen_jn.txtCode_cr.Text = GetGridValue(FG, FG.CurrentRow.Index, 3)

        FmNewOpen_jn.txtCode_dr.Text = GetGridValue(FG, FG.CurrentRow.Index, 2)
        FmNewOpen_jn.txtCode_cr.Text = GetGridValue(FG, FG.CurrentRow.Index, 3)
        FmNewOpen_jn.LAbel9.Text = GetGridValue(FG, FG.CurrentRow.Index, 13)
        'MuSubOffOp()

        BtnEdit.Enabled = True
        BtnDelete.Enabled = True
        Button1.Enabled = True
        FmNewOpen_jn.txtCode_dr.Enabled = False
        FmNewOpen_jn.txtCode_cr.Enabled = False
        FmNewOpen_jn.BtnSearch_dr.Enabled = False
        FmNewOpen_jn.BtnSearch_cr.Enabled = False
    End Sub

    Private Sub BtnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRefresh.Click
        'LoadLook()
        BntNew.Enabled = True
        CNN.Execute("  Update Open_jn set amount_cr =0 where amount_cr is null Update Open_jn set amount_Dr =0 where amount_Dr is null")
        LoadListFG()

    End Sub

    Private Sub BtnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        'If MdAtv = False Then MessageBox.Show("ກະລຸນນາເລືອກກ່ອນ") : Exit Sub
        Call LoadReport()

    End Sub
    Public Sub LoadReport()
        If CheckBox1.Checked = True Then
            MuLng = "E"
        Else
            MuLng = "L"
        End If


        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng ,"
        LngId = "7025" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & " " & yy.Text & " ' As RptSjUd ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
        LngId = "7002" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDP ,"
        LngId = "7085" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
        LngId = "7004" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Descrip ,"
        LngId = "7005" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_No ,"
        LngId = "7006" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Ac_Code ,"
        LngId = "7007" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Certify ,"
        LngId = "7008" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Dr ,"
        LngId = "7009" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Cr ,"
        LngId = "7010" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Amt_LAK ,"
        LngId = "7011" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Sign ,"
        LngId = "7012" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_SignNme ,"
        LngId = "7026" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InDate ,"
        LngId = "7028" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_AmtCurr ,"
        LngId = "7029" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Curr ,"
        LngId = "7030" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Rate ,"
        LngId = "7031" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_User ,"
        LngId = "7039" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_TotalAmt ,"
        LngId = "7040" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Balace ,"
        LngId = "7041" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Company ,"

        SLF = "SELECT " & MuLngRpt & "    "


        Dim OfUsr1 As String = Microsoft.VisualBasic.Left(Off_Usr.Text, 5)
        Dim OfUsr2 As String = Mid(Off_Usr.Text, 4, 2)
        Dim OfUsr3 As String = Microsoft.VisualBasic.Left(Off_Usr.Text, 2)
        If OfUsr1 = "00-00" Then
            MULook2 = ""
        Else
            If OfUsr2 = "00" Then
                MULook2 = "  And  Left(Open_jn.company,2)= '" & OfUsr3 & "' "
            Else
                MULook2 = "  And Open_jn.company= '" & OfUsr1 & "' "
            End If
        End If
        Call LoadLoGO()
        CNN.Execute("UPDATE Open_jn set Open_jn.Ac_original=Acc_Code.Ac_original from Acc_Code,Open_jn where Open_jn.ac_code=Acc_Code.ac_code and (Open_jn.Ac_original is null or Open_jn.Ac_original='') ")
        Dim Rs As DataTable
        If ChAllSty.Checked = True Then
            'Dim s As String = "    " & SLF & "   sum(Open_jn.amt_dr) as amt_dr , sum(Open_jn.amt_cr) as amt_cr ,   Open_jn.ac_code AS ac_code , Acc_Code.Name_L FROM         Open_jn INNER JOIN   Acc_Code ON Open_jn.ac_code = Acc_Code.Ac_Code   where  Open_jn.date_work   BETWEEN '" & Format(MWorkSetting, "yyyy") & "-1-1" & "' AND '" & Format(MWorkSetting, "yyyy") & "-1-1" & "' " & MULook2 & " group by Open_jn.ac_code, Acc_Code.Name_L  having   SUM((Open_jn.amt_cr)-(Open_jn.amt_dr)) <> 0 Order by  ac_code"

            Rs = DbHelper.GetDataTable(s)
        Else
            Dim s2 As String = "    " & SLF & "   sum(Open_jn.amt_dr) as amt_dr , sum(Open_jn.amt_cr) as amt_cr ,   Open_jn.ac_code AS ac_code , Open_jn.ac_name  FROM   Open_jn    where  Open_jn.date_work   BETWEEN '" & Format(MWorkSetting, "yyyy") & "-1-1" & "' AND '" & Format(MWorkSetting, "yyyy") & "-1-1" & "' " & MULook2 & " group by Open_jn.ac_code, Open_jn.ac_name     having   SUM((Open_jn.amt_cr)-(Open_jn.amt_dr)) <> 0 Order by  ac_code"

            Rs = DbHelper.GetDataTable(s2)
        End If

        If Rs.Rows.Count = 0 Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
        If ChAllSty.Checked = False Then
            Dim FrmPreview As New FmPreview : FrmClosing()
            Dim Rpt As New Crybandon_year

            If MdShowLOGO = 1 Then
                Rpt.Subreports(0).SetDataSource(RsLOGO)
            End If

            Rpt.SetDataSource(Rs.Tables(0))
            FrmPreview.ReportViewer.ReportSource = Rpt
            FrmPreview.ReportViewer.DisplayGroupTree = False
            FrmPreview.MdiParent = FmMain
            FrmPreview.WindowState = FormWindowState.Maximized
            FrmPreview.Show()
            FrmPreview.Focus()
        Else

            Dim FrmPreview As New FmPreview : FrmClosing()
            Dim Rpt As New CryOpen_yearAllSty

            If MdShowLOGO = 1 Then
                Rpt.Subreports(0).SetDataSource(RsLOGO)
            End If

            Rpt.SetDataSource(Rs.Tables(0))
            FrmPreview.ReportViewer.ReportSource = Rpt
            FrmPreview.ReportViewer.DisplayGroupTree = False
            FrmPreview.MdiParent = FmMain
            FrmPreview.WindowState = FormWindowState.Maximized
            FrmPreview.Show()
            FrmPreview.Focus()
        End If

    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

  


    Private Sub Sub_Company_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ChAllSty_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChAllSty.CheckedChanged
        If ChAllSty.Checked = True Then
            BntNew.Enabled = False
            BtnEdit.Enabled = False
            BtnDelete.Enabled = False
            Button1.Enabled = False
            FmNewOpen_jn.BntNew.Enabled = False
            FmNewOpen_jn.BtnSave.Enabled = False


        Else

            BntNew.Enabled = True
            BtnEdit.Enabled = True
            BtnDelete.Enabled = True
            Button1.Enabled = True
            FmNewOpen_jn.BntNew.Enabled = True
            FmNewOpen_jn.BtnSave.Enabled = True
            If Mid(FmLogin.Sub_Company.Text, 4, 2) = "00" Then
                BntNew.Enabled = False
                BtnEdit.Enabled = False
                BtnDelete.Enabled = False
                Button1.Enabled = False
                FmNewOpen_jn.BntNew.Enabled = False
                FmNewOpen_jn.BtnSave.Enabled = False
            End If
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub yy_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles yy.KeyPress
        If e.KeyChar = Chr(13) Then
            BntNew.Enabled = True
            LoadListFG()
        End If
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            BntNew.Enabled = True
            LoadListFG()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If MdAtv = False Then MessageBox.Show("ກະລຸນນາເລືອກກ່ອນ") : Exit Sub
        If MessageBox.Show("ທ່ານຕ້ອງການປົດລ໋ອກປີ " & yy.Text & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            CNN.Execute("UPDATE Open_jn SET  my_lock = '" & "0" & "' WHERE Year(Date_Work) = '" & yy.Text & "'")
        End If
        MdAtv = False
        LoadListFG()
    End Sub
End Class
