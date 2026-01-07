Public Class Frm_AssetAdd

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()

    End Sub

    Private Sub FrmAdjustment_List_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetupGrid()
        LdGrp()
        LoadListFG()
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
        FG.Columns.Add("AdjustingPeriod", "Adjusting Period")
        FG.Columns.Add("Dr", "Dr")
        FG.Columns.Add("Cr", "Cr")

        ' Set column widths
        FG.Columns(0).Width = 50
        FG.Columns(1).Width = 100
        FG.Columns(2).Width = 150
        FG.Columns(3).Width = 150
        FG.Columns(4).Width = 100
        FG.Columns(5).Width = 120
        FG.Columns(6).Width = 100
        FG.Columns(7).Width = 120
        FG.Columns(8).Width = 100
        FG.Columns(9).Width = 100

        ' Configure DataGridView properties
        FG.AllowUserToAddRows = False
        FG.ReadOnly = True
        FG.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG.MultiSelect = False
    End Sub

    Private Sub LdGrp()
        Dim dt As DataTable = DbHelper.GetDataTable("Select * from Groups Order by Group_ID")
        txtGrpNm.Items.Clear()
        If Lang = True Then
            txtGrpNm.Items.Add("All Group")
            If dt.Rows.Count > 0 Then
                For Each row As DataRow In dt.Rows
                    txtGrpNm.Items.Add(DbHelper.GetStr(row("Group_NmE")))
                Next
            End If
            txtGrpNm.SelectedIndex = 0
        Else
            txtGrpNm.Items.Add("ທັງໝົດ ")
            If dt.Rows.Count > 0 Then
                For Each row As DataRow In dt.Rows
                    txtGrpNm.Items.Add(DbHelper.GetStr(row("Group_Nm")))
                Next
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

'    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
'        If TxtCode.Text = "" Then MsgBox("ກະລຸນາເລືອກກ່ອນ!", MsgBoxStyle.OkOnly) : Exit Sub
'        If FG.CurrentRow Is Nothing Then Exit Sub
'        If MessageBox.Show("ທ່ານຕ້ອງລຶບລະຫັດ " & FG.CurrentRow.Cells(1).Value.ToString() & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
'            CNN.Execute("DELETE FROM Adjustment_List WHERE Code=N'" & FG.CurrentRow.Cells(1).Value.ToString() & "'")
'
'            LoadListFG()
'            Call AddNew()
'        End If
'    End Sub
    Public Sub LoadListFG()
        Dim GrpNM As String
        If txtGrpNm.SelectedIndex = 0 Then
            GrpNM = ""
        Else
            GrpNM = " AND GrpID=N'" & Trim(txtGrp.Text) & "' "
        End If

        ' Clear existing rows
        FG.Rows.Clear()

        Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM  Adjustment_List where 1=1 " & GrpNM & " order by Code ASC")
        If dt.Rows.Count > 0 Then
            Dim rowNum As Integer = 1
            For Each row As DataRow In dt.Rows
                FG.Rows.Add(rowNum, _
                    DbHelper.GetStr(row("Code")), _
                    DbHelper.GetStr(row("Name")), _
                    DbHelper.GetStr(row("NameE")), _
                    Format(CDbl(DbHelper.GetStr(row("Value"))), "##,##0.00"), _
                    Format(CDbl(DbHelper.GetStr(row("Remain"))), "##,##0.00"), _
                    Format(CDate(DbHelper.GetStr(row("DateIn"))), "dd/MM/yyyy"), _
                    DbHelper.GetStr(row("Period")), _
                    DbHelper.GetStr(row("Dr")), _
                    DbHelper.GetStr(row("Cr")))
                rowNum += 1
            Next
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If txtGrpNm.SelectedIndex = 0 Then
            MsgBox("ກະລຸນາເລືອກໝວດຊັບສິນກ່ອນ!", MsgBoxStyle.Exclamation) : txtGrpNm.Focus() : Exit Sub
        End If
        If TxtPeriod.Text = 0 Then
            MsgBox("Before Adjust Period!", MsgBoxStyle.Exclamation) : txtGrpNm.Focus() : Exit Sub
        End If

        If TxtCode.Text = "" Then MsgBox("", MsgBoxStyle.Exclamation) : TxtCode.Focus() : Exit Sub

        If TxtCode.Enabled = True Then
            Dim dtCheck As DataTable = DbHelper.GetDataTable("SELECT * FROM Adjustment_List WHERE Code =N'" & Trim(TxtCode.Text) & "'")
            If dtCheck.Rows.Count <> 0 Then
                MsgBox("ລະຫັດມີແລ້ວ!", MsgBoxStyle.Exclamation) : TxtCode.Focus() : Exit Sub
            End If
        End If


        Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM Adjustment_List WHERE Code =N'" & Trim(TxtCode.Text) & "'")
        If dt.Rows.Count = 0 Then
            DbHelper.ExecuteNonQuery("INSERT INTO Adjustment_List(Code, GrpID, GrpIDNm, Name, NameE, Desription, DateIn, Period, Value, Remain, Dr, DrNm, Cr, CrNm) " & _
                "Values(N'" & Trim(TxtCode.Text) & "',N'" & Trim(txtGrp.Text) & "',N'" & Trim(txtGrpNm.Text) & "' ,N'" & Trim(TxtName.Text) & "' , N'" & Trim(TxtNameE.Text) & "' ,N'" & Trim(TxtDesription.Text) & "','" & Format(DateIn.Value, "yyyy-MM-dd") & "'," & CDbl(TxtPeriod.Text) & "," & CDbl(TxtValue.Text) & "," & CDbl(TxtRemain.Text) & ",N'" & Trim(TxtDr.Text) & "',N'" & Trim(TxtDrNm.Text) & "',N'" & Trim(TxtCr.Text) & "',N'" & Trim(TxtCrNm.Text) & "')")
        Else
            DbHelper.ExecuteNonQuery("DELETE Adjustment_List WHERE Code =N'" & Trim(TxtCode.Text) & "' ")
            DbHelper.ExecuteNonQuery("INSERT INTO Adjustment_List(Code, GrpID, GrpIDNm, Name, NameE, Desription, DateIn, Period, Value, Remain, Dr, DrNm, Cr, CrNm) " & _
             "Values(N'" & Trim(TxtCode.Text) & "',N'" & Trim(txtGrp.Text) & "',N'" & Trim(txtGrpNm.Text) & "' ,N'" & Trim(TxtName.Text) & "' , N'" & Trim(TxtNameE.Text) & "' ,N'" & Trim(TxtDesription.Text) & "','" & Format(DateIn.Value, "yyyy-MM-dd") & "'," & CDbl(TxtPeriod.Text) & "," & CDbl(TxtValue.Text) & "," & CDbl(TxtRemain.Text) & ",N'" & Trim(TxtDr.Text) & "',N'" & Trim(TxtDrNm.Text) & "',N'" & Trim(TxtCr.Text) & "',N'" & Trim(TxtCrNm.Text) & "')")

        End If
        DbHelper.ExecuteNonQuery("UPDATE Adjustment_List set day=Value/Period  WHERE Code =N'" & Trim(TxtCode.Text) & "'  ")
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
        Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM Adjustment_List WHERE Code =N'" & Trim(TxtCode.Text) & "'")
        If dt.Rows.Count = 0 Then
            AddNew()
        Else
            TxtCode.Text = DbHelper.GetStr(dt.Rows(0)("Code"))
            TxtName.Text = DbHelper.GetStr(dt.Rows(0)("Name"))
            TxtNameE.Text = DbHelper.GetStr(dt.Rows(0)("NameE"))

            TxtValue.Text = Format(CDbl(DbHelper.GetStr(dt.Rows(0)("Value"))), "#,##0.00")
            TxtRemain.Text = Format(CDbl(DbHelper.GetStr(dt.Rows(0)("Remain"))), "#,##0.00")
            TxtPeriod.Text = Format(CDbl(DbHelper.GetStr(dt.Rows(0)("Period"))), "#,##0.00")

            DateIn.Value = CDate(DbHelper.GetStr(dt.Rows(0)("DateIn")))
            TxtDr.Text = DbHelper.GetStr(dt.Rows(0)("Dr"))
            TxtDrNm.Text = DbHelper.GetStr(dt.Rows(0)("DrNm"))
            TxtCr.Text = DbHelper.GetStr(dt.Rows(0)("Cr"))
            TxtCrNm.Text = DbHelper.GetStr(dt.Rows(0)("CrNm"))
            TxtDesription.Text = DbHelper.GetStr(dt.Rows(0)("Desription"))
            txtGrp.Text = DbHelper.GetStr(dt.Rows(0)("GrpID"))
            txtGrpNm.Text = DbHelper.GetStr(dt.Rows(0)("GrpIDNm"))
        End If
    End Sub

    Private Sub txtGrpNm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtGrpNm.SelectedIndexChanged
        If Lang = True Then
            Dim dt As DataTable = DbHelper.GetDataTable("select * from Groups Where Group_NmE=N'" & Trim(txtGrpNm.Text) & "'")
            If dt.Rows.Count > 0 Then
                txtGrp.Text = DbHelper.GetStr(dt.Rows(0)("Group_ID"))
            Else
                txtGrp.Text = ""
            End If
        Else
            Dim dt As DataTable = DbHelper.GetDataTable("select * from Groups Where Group_Nm=N'" & Trim(txtGrpNm.Text) & "' ")
            If dt.Rows.Count > 0 Then
                txtGrp.Text = DbHelper.GetStr(dt.Rows(0)("Group_ID"))
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
            Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtDr.Text & "' ")
            If dt.Rows.Count > 0 Then
                TxtDrNm.Text = DbHelper.GetStr(dt.Rows(0)("Name_L"))
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
            Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtCr.Text & "' ")
            If dt.Rows.Count > 0 Then
                TxtCrNm.Text = DbHelper.GetStr(dt.Rows(0)("Name_L"))
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
End Class