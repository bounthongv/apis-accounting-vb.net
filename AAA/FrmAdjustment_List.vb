Public Class FrmAdjustment_List

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()

    End Sub

    Private Sub FrmAdjustment_List_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetupGrid()
        Call load_Cmb(" SELECT Curr  FROM Curr_For_Rate  ORDER BY cnt ", "Curr", Cmb)
        If Cmb.Items.Count > 0 Then
            Cmb.SelectedIndex = 0
        End If

        LdGrp()
        LoadListFG()
        Call AddNew()
    End Sub

    Private Sub SetupGrid()
        ' Clear and setup DataGridView columns
        FG.Columns.Clear()
        FG.Columns.Add("No", "No.")
        FG.Columns.Add("Code", "Code")
        FG.Columns.Add("AdjustmentLA", "Adjustment (LA)")
        FG.Columns.Add("AdjustmentEN", "Adjustment (EN)")
        FG.Columns.Add("Value", "Value")
        FG.Columns.Add("RemainingValue", "Remaining Value")
        FG.Columns.Add("DateIN", "Date IN")
        FG.Columns.Add("AdjustPeriod", "Adjust Period")
        FG.Columns.Add("Dr", "Dr")
        FG.Columns.Add("Cr", "Cr")
        FG.Columns.Add("Currency", "Currency")
        FG.Columns.Add("ForeignDebit", "Foreign Debit")
        FG.Columns.Add("ForeignCredit", "Foreign Credit")

        ' Set column widths
        FG.Columns(0).Width = 50
        FG.Columns(1).Width = 100
        FG.Columns(2).Width = 150
        FG.Columns(3).Width = 150
        FG.Columns(4).Width = 100
        FG.Columns(5).Width = 120
        FG.Columns(6).Width = 100
        FG.Columns(7).Width = 100
        FG.Columns(8).Width = 100
        FG.Columns(9).Width = 100
        FG.Columns(10).Width = 80
        FG.Columns(11).Width = 120
        FG.Columns(12).Width = 120

        ' Configure DataGridView properties
        FG.AllowUserToAddRows = False
        FG.ReadOnly = True
        FG.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG.MultiSelect = False
    End Sub

    Private Sub LdGrp()
        Dim gRS As New ADODB.Recordset
        txtGrpNm.Items.Clear()
        If Lang = True Then
            txtGrpNm.Items.Add("All Group")
            Call LoadSqlData("Select * from Groups Order by Group_ID", gRS)
            If gRS.RecordCount <> 0 Then
                While Not gRS.EOF
                    txtGrpNm.Items.Add(gRS.Fields("Group_NmE").Value.ToString)
                    gRS.MoveNext()
                End While
            End If
            txtGrpNm.SelectedIndex = 0
        Else
            txtGrpNm.Items.Add("ທັງໝົດ ")
            Call LoadSqlData("Select * from Groups Order by Group_ID", gRS)
            If gRS.RecordCount <> 0 Then
                While Not gRS.EOF
                    txtGrpNm.Items.Add(gRS.Fields("Group_Nm").Value.ToString)
                    gRS.MoveNext()
                End While
            End If
            txtGrpNm.SelectedIndex = 0
        End If
    End Sub

    Private Sub BtnAddNew2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew2.Click
        Call AddNew()
    End Sub
    Private Sub AddNew()
        TxtCode.Text = ""
        TxtName.Text = ""
        TxtNameE.Text = ""
        TxtValue.Text = "0"
        TxtRemain.Text = "0"
        TxtDesription.Text = ""
        TxtDr.Text = ""
        TxtCr.Text = ""
        TxtDrNm.Text = ""
        TxtCrNm.Text = ""
        TxtPeriod.Text = "0"
        TxtCode.Enabled = True
        TxtCode.Focus()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TxtCode.Text = "" Then MsgBox("ກະລຸນາເລືອກກ່ອນ!", MsgBoxStyle.OkOnly) : Exit Sub
        If FG.CurrentRow Is Nothing Then Exit Sub
        If MessageBox.Show("ທ່ານຕ້ອງລຶບລະຫັດ " & FG.CurrentRow.Cells(1).Value.ToString() & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            CNN.Execute("DELETE FROM Adjustment_List WHERE Code=N'" & FG.CurrentRow.Cells(1).Value.ToString() & "'")

            LoadListFG()
            Call AddNew()
        End If
    End Sub
    Public Sub LoadListFG()
        Dim GrpNM As String
        If txtGrpNm.SelectedIndex = 0 Then
            GrpNM = ""
        Else
            GrpNM = " AND GrpID=N'" & Trim(txtGrp.Text) & "' "
        End If

        ' Clear existing rows
        FG.Rows.Clear()

        With RSC
            Call LoadSqlData("SELECT * FROM  Adjustment_List where 1=1 " & GrpNM & " order by Code ASC  ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.Rows.Add(.AbsolutePosition, _
                      Trim(CStr(.Fields("Code").Value)), _
                      Trim(CStr(.Fields("Name").Value.ToString)), _
                      Trim(CStr(.Fields("NameE").Value.ToString)), _
                      Format(CDbl(Trim(.Fields("Value").Value)), "##,##0.00"), _
                      Format(CDbl(Trim(.Fields("Remain").Value)), "##,##0.00"), _
                      Format(CDate(Trim(.Fields("DateIn").Value)), "dd/MM/yyyy"), _
                      Trim(CStr(.Fields("Period").Value.ToString)), _
                      Trim(CStr(.Fields("Dr").Value.ToString)), _
                      Trim(CStr(.Fields("Cr").Value.ToString)), _
                      Trim(CStr(.Fields("Curr").Value.ToString)), _
                      Trim(CStr(.Fields("Dr_Curr").Value.ToString)), _
                      Trim(CStr(.Fields("Cr_Curr").Value.ToString)))
                    .MoveNext()
                End While
            End If
        End With

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If txtGrpNm.SelectedIndex = 0 Then
            MsgBox("ກະລຸນາເລືອກໝວດຊັບສິນກ່ອນ!", MsgBoxStyle.Exclamation) : txtGrpNm.Focus() : Exit Sub
        End If
        If TxtPeriod.Text = 0 Then
            MsgBox("Before Adjust Period!", MsgBoxStyle.Exclamation) : TxtPeriod.Focus() : Exit Sub
        End If

        If TxtCode.Text = "" Then MsgBox("Input Code", MsgBoxStyle.Exclamation) : TxtCode.Focus() : Exit Sub

        If TxtCode.Enabled = True Then
            Call LoadSqlData("SELECT * FROM Adjustment_List WHERE Code =N'" & Trim(TxtCode.Text) & "'", RSC)
            If RSC.RecordCount <> 0 Then
                MsgBox("ລະຫັດມີແລ້ວ!", MsgBoxStyle.Exclamation) : TxtCode.Focus() : Exit Sub
            End If
        End If


        Call LoadSqlData("SELECT * FROM Adjustment_List WHERE Code =N'" & Trim(TxtCode.Text) & "'", RSC)
        If RSC.RecordCount = 0 Then
            CNN.Execute("INSERT INTO Adjustment_List(Code, GrpID, GrpIDNm, Name, NameE, Desription, DateIn, Period, Value, Remain, Dr, DrNm, Cr, CrNm,Curr,Rate,Dr_Curr,Dr_CurrNm,Cr_Curr,Cr_CurrNm) " & _
                "Values(N'" & Trim(TxtCode.Text) & "',N'" & Trim(txtGrp.Text) & "',N'" & Trim(txtGrpNm.Text) & "' ,N'" & Trim(TxtName.Text) & "' , N'" & Trim(TxtNameE.Text) & "' ,N'" & Trim(TxtDesription.Text) & "','" & Format(DateIn.Value, "yyyy-MM-dd") & "'," & CDbl(TxtPeriod.Text) & "," & CDbl(TxtValue.Text) & "," & CDbl(TxtRemain.Text) & ",N'" & Trim(TxtDr.Text) & "',N'" & Trim(TxtDrNm.Text) & "',N'" & Trim(TxtCr.Text) & "',N'" & Trim(TxtCrNm.Text) & "',N'" & Trim(Cmb.Text) & "'," & CDbl(txtRate.Text) & ",N'" & Trim(TxtDr_Curr.Text) & "',N'" & Trim(TxtDrNm_Curr.Text) & "',N'" & Trim(TxtCr_Curr.Text) & "',N'" & Trim(TxtCrNm_Curr.Text) & "')")
        Else
            CNN.Execute("DELETE Adjustment_List WHERE Code =N'" & Trim(TxtCode.Text) & "' ")
            CNN.Execute("INSERT INTO Adjustment_List(Code, GrpID, GrpIDNm, Name, NameE, Desription, DateIn, Period, Value, Remain, Dr, DrNm, Cr, CrNm,Curr,Rate,Dr_Curr,Dr_CurrNm,Cr_Curr,Cr_CurrNm) " & _
                "Values(N'" & Trim(TxtCode.Text) & "',N'" & Trim(txtGrp.Text) & "',N'" & Trim(txtGrpNm.Text) & "' ,N'" & Trim(TxtName.Text) & "' , N'" & Trim(TxtNameE.Text) & "' ,N'" & Trim(TxtDesription.Text) & "','" & Format(DateIn.Value, "yyyy-MM-dd") & "'," & CDbl(TxtPeriod.Text) & "," & CDbl(TxtValue.Text) & "," & CDbl(TxtRemain.Text) & ",N'" & Trim(TxtDr.Text) & "',N'" & Trim(TxtDrNm.Text) & "',N'" & Trim(TxtCr.Text) & "',N'" & Trim(TxtCrNm.Text) & "',N'" & Trim(Cmb.Text) & "'," & CDbl(txtRate.Text) & ",N'" & Trim(TxtDr_Curr.Text) & "',N'" & Trim(TxtDrNm_Curr.Text) & "',N'" & Trim(TxtCr_Curr.Text) & "',N'" & Trim(TxtCrNm_Curr.Text) & "')")

        End If
        CNN.Execute("UPDATE Adjustment_List set day=Value/Period  WHERE Code =N'" & Trim(TxtCode.Text) & "'  ")

        If RSC.State = ConnectionState.Open Then RSC.Close()
        MsgBox("ການບັນທຶກສຳເລັດ!", MsgBoxStyle.OkOnly)
        TxtCode.Focus()
        LoadListFG()
    End Sub

    Private Sub FG_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles FG.SelectionChanged, FG.Click
        If FG.CurrentRow Is Nothing Then Exit Sub
        If FG.CurrentRow.Index < 0 Then Exit Sub

        Try
            TxtCode.Text = FG.CurrentRow.Cells(1).Value.ToString()
            TxtName.Text = FG.CurrentRow.Cells(2).Value.ToString()
            Call LoadText()
            TxtCode.Enabled = False
        Catch ex As Exception
            ' Handle potential conversion errors or empty cells
        End Try
    End Sub
    Private Sub LoadText()

        CNN.Execute("    update Adjustment_List set Rate=1  where curr='LAK' ")
        CNN.Execute("   update Adjustment_List set Rate=1  where Rate is null ")

        Call LoadSqlData("SELECT * FROM Adjustment_List WHERE Code =N'" & Trim(TxtCode.Text) & "'", RSC)
        If RSC.RecordCount = 0 Then
            AddNew()
        Else
            TxtCode.Text = Trim(RSC.Fields("Code").Value.ToString)
            TxtName.Text = Trim(RSC.Fields("Name").Value.ToString)
            TxtNameE.Text = Trim(RSC.Fields("NameE").Value.ToString)

            TxtValue.Text = Format(RSC.Fields("Value").Value, "#,##0.00")
            TxtRemain.Text = Format(RSC.Fields("Remain").Value, "#,##0.00")
            TxtPeriod.Text = Format(RSC.Fields("Period").Value, "#,##0.00")

            DateIn.Value = Format(RSC.Fields("DateIn").Value, "dd/MM/yyyy")
            TxtDr.Text = Trim(RSC.Fields("Dr").Value.ToString)
            TxtDrNm.Text = Trim(RSC.Fields("DrNm").Value.ToString)
            TxtCr.Text = Trim(RSC.Fields("Cr").Value.ToString)
            TxtCrNm.Text = Trim(RSC.Fields("CrNm").Value.ToString)
            TxtDesription.Text = Trim(RSC.Fields("Desription").Value.ToString)
            txtGrp.Text = Trim(RSC.Fields("GrpID").Value.ToString)
            txtGrpNm.Text = Trim(RSC.Fields("GrpIDNm").Value.ToString)
            Cmb.Text = Trim(RSC.Fields("Curr").Value.ToString)
            txtRate.Text = Format(RSC.Fields("Rate").Value, "#,##0.00")

            TxtDr_Curr.Text = Trim(RSC.Fields("Dr_Curr").Value.ToString)
            TxtDrNm_Curr.Text = Trim(RSC.Fields("Dr_CurrNm").Value.ToString)
            TxtCr_Curr.Text = Trim(RSC.Fields("Cr_Curr").Value.ToString)
            TxtCrNm_Curr.Text = Trim(RSC.Fields("Cr_CurrNm").Value.ToString)

        End If
    End Sub

    Private Sub txtGrpNm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtGrpNm.SelectedIndexChanged
        Dim gRS As New ADODB.Recordset
        If Lang = True Then
            Call LoadSqlData("select * from Groups Where Group_NmE=N'" & Trim(txtGrpNm.Text) & "'", gRS)
            If gRS.RecordCount <> 0 Then
                txtGrp.Text = Trim(gRS.Fields("Group_ID").Value.ToString)
            Else
                txtGrp.Text = ""
            End If
        Else
            Call LoadSqlData("select * from Groups Where Group_Nm=N'" & Trim(txtGrpNm.Text) & "' ", gRS)
            If gRS.RecordCount <> 0 Then
                txtGrp.Text = Trim(gRS.Fields("Group_ID").Value.ToString)
            Else
                txtGrp.Text = ""
            End If
        End If
        TxtName.Focus()
        LoadListFG()
    End Sub

    Private Sub BtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch.Click
        fmShartOfAccDetail.txtSty.Text = "FrmAdjustment_List_Dr"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        fmShartOfAccDetail.txtSty.Text = "FrmAdjustment_List_Cr"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
    End Sub

    Private Sub TxtDr_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtDr.KeyPress
        If e.KeyChar = Chr(13) Then
            Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtDr.Text & "' ", RSC)
            If RSC.RecordCount <> 0 Then
                TxtDrNm.Text = Trim(RSC.Fields("Name_L").Value.ToString)
            Else
                MsgBox("ເລກບັນຊີບໍ່ມີໃນລາລະບານ", MsgBoxStyle.Exclamation) : TxtDr.Focus() : Exit Sub
            End If

            TxtCr.Focus()
        End If
    End Sub

    Private Sub TxtDr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtDr.TextChanged

    End Sub

    Private Sub TxtCr_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCr.KeyPress
        If e.KeyChar = Chr(13) Then
            Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtCr.Text & "' ", RSC)
            If RSC.RecordCount <> 0 Then
                TxtCrNm.Text = Trim(RSC.Fields("Name_L").Value.ToString)
            Else
                MsgBox("ເລກບັນຊີບໍ່ມີໃນລາລະບານ", MsgBoxStyle.Exclamation) : TxtCr.Focus() : Exit Sub
            End If
            Button2.Focus()
        End If


    End Sub

    Private Sub TxtCr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCr.TextChanged

    End Sub

    Private Sub TxtValue_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtValue.KeyPress
        If e.KeyChar = Chr(13) Then
            TxtValue.Text = Format(CDbl(TxtValue.Text), "#,#0.00")
            If TxtRemain.Text = 0 Then
                TxtRemain.Text = TxtValue.Text
                TxtRemain.Text = Format(CDbl(TxtRemain.Text), "#,#0.00")
            End If
            TxtRemain.Focus()
        End If
  
    End Sub

    Private Sub TxtValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtValue.TextChanged

    End Sub

    Private Sub TxtRemain_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtRemain.KeyPress
        If e.KeyChar = Chr(13) Then
            TxtRemain.Text = Format(CDbl(TxtRemain.Text), "#,#0.00")

            TxtDesription.Focus()

        End If
    End Sub

    Private Sub TxtRemain_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtRemain.TextChanged

    End Sub

    Private Sub TxtPeriod_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtPeriod.KeyPress
        If e.KeyChar = Chr(13) Then
            TxtPeriod.Text = Format(CDbl(TxtPeriod.Text), "#,#0.00")

            TxtDr.Focus()

        End If
    End Sub

    Private Sub TxtPeriod_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtPeriod.TextChanged

    End Sub

    Private Sub TxtName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtName.KeyPress
        If e.KeyChar = Chr(13) Then
            TxtNameE.Focus()
        End If
    End Sub

    Private Sub TxtName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtName.TextChanged

    End Sub

    Private Sub TxtNameE_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtNameE.KeyPress
        If e.KeyChar = Chr(13) Then
            TxtValue.Focus()
        End If
    End Sub

    Private Sub TxtNameE_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtNameE.TextChanged

    End Sub

    Private Sub TxtDesription_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtDesription.KeyPress
        If e.KeyChar = Chr(13) Then
            DateIn.Focus()
        End If
    End Sub

    Private Sub TxtDesription_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtDesription.TextChanged

    End Sub

    Private Sub DateIn_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateIn.KeyPress
        If e.KeyChar = Chr(13) Then
            TxtPeriod.Focus()
        End If
    End Sub

    Private Sub DateIn_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateIn.ValueChanged

    End Sub

    Private Sub Cmb_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmb.SelectedIndexChanged

        Dim rs As New ADODB.Recordset
        Call LoadSqlData("Select * From Curr_For_Rate Where   Curr =N'" & Trim(Cmb.Text) & "'", rs)
        If rs.RecordCount > 0 Then
            txtcurr_name2.Text = Trim(rs("Curr_name").Value.ToString)
        End If

        MDRate_DT = " and rate_dt<='" & Format(DateIn.Value, "yyyy-MM-dd") & "'  "
        'MDRate_DT = " and rate_dt<='" & Format(MdToDate, "yyyy-MM-dd") & "'  "

        SS_Curr = " and AP_Rate_history.Curr =N'" & Cmb.Text & "' "
        Call RateSetting()
        txtRate.Text = Format(MD_Rate, "#,##0.00")

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        fmShartOfAccDetail.txtSty.Text = "FrmAdjustment_List_Dr_Curr"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        fmShartOfAccDetail.txtSty.Text = "FrmAdjustment_List_Cr_Curr"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
    End Sub

    Private Sub TxtDr_Curr_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtDr_Curr.KeyPress
        If e.KeyChar = Chr(13) Then
            Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtDr_Curr.Text & "' ", RSC)
            If RSC.RecordCount <> 0 Then
                TxtDrNm_Curr.Text = Trim(RSC.Fields("Name_L").Value.ToString)
            Else
                MsgBox("ເລກບັນຊີບໍ່ມີໃນລາລະບານ", MsgBoxStyle.Exclamation) : TxtDr_Curr.Focus() : Exit Sub
            End If

            TxtCr_Curr.Focus()
        End If
    End Sub

    Private Sub TxtDr_Curr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtDr_Curr.TextChanged

    End Sub

    Private Sub TxtCr_Curr_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCr_Curr.KeyPress
        If e.KeyChar = Chr(13) Then
            Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtCr_Curr.Text & "' ", RSC)
            If RSC.RecordCount <> 0 Then
                TxtCrNm_Curr.Text = Trim(RSC.Fields("Name_L").Value.ToString)
            Else
                MsgBox("ເລກບັນຊີບໍ່ມີໃນລາລະບານ", MsgBoxStyle.Exclamation) : TxtCr_Curr.Focus() : Exit Sub
            End If

            TxtCr_Curr.Focus()
        End If
    End Sub

    Private Sub TxtCr_Curr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCr_Curr.TextChanged

    End Sub
End Class