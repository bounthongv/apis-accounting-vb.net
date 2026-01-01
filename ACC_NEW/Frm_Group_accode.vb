Public Class Frm_Group_accode
    Dim rs As New ADODB.Recordset
    Dim ac_A, ac_B As String

    Private Sub Frm_Department_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetupGrid()
        LoadData()
    End Sub

    Private Sub SetupGrid()
        fg.Columns.Clear()
        fg.Columns.Add("No", "ລດ")
        fg.Columns.Add("AC_CODE", "ລະ​ຫັດ ບັນ​ຊີ")
        fg.Columns.Add("Name_L", "ຊື່ ບັນ​ຊີ")
        fg.Columns.Add("H1", "ລະ​ຫັດ ບັນ​ຊີ​ແມ່")
        fg.Columns.Add("H1_nm", "ຊື່ ບັນ​ຊີ​ແມ່")

        fg.Columns(0).Width = 50
        fg.Columns(1).Width = 120
        fg.Columns(2).Width = 300
        fg.Columns(3).Width = 120
        fg.Columns(4).Width = 300

        fg.AllowUserToAddRows = False
        fg.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        fg.ReadOnly = True
    End Sub

    Private Sub AutoNumber()
        Dim VIOT As New ADODB.Recordset
        Dim VIOTNEW As String
        Call LoadSqlData("SELECT top 1 Don_ID from AP_Donnor   Order by Don_ID DESC", VIOT)
        If VIOT.RecordCount <> 0 Then
            VIOTNEW = Format(Val(Mid(VIOT.Fields("Don_ID").Value, 1, 2)) + 1, "00")
        Else
            VIOTNEW = "01"
        End If
        Txt_ID.Text = Trim(CStr(VIOTNEW.ToString))
    End Sub

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        Call save()
        MsgBox("ຄຳນວນ ສຳເລັດຜົນ!", MsgBoxStyle.OkOnly)
        Call LoadData()
        Txt_name_L.Text = ""
        Txt_name_L.Focus()
    End Sub

    Public Sub save()
        Dim aa As String
        Dim AcTypeLao, AcTypeEng As String
        For j = 0 To fg.Rows.Count - 1
            ac_A = ""
            ac_B = ""
            Dim acCode As String = If(fg.Rows(j).Cells(1).Value Is Nothing, "", fg.Rows(j).Cells(1).Value.ToString())
            
            For i = 1 To Len(Trim(acCode))
                If Mid(Trim(acCode), i, 1) = "." Then
                    AcTypeLao = "ບັນຊີຍ່ອຍ (D)"
                    AcTypeEng = "Detail Account"
                    Exit For
                Else
                    AcTypeLao = "ບັນຊີແມ່ (P)"
                    AcTypeEng = "Parent Account"
                End If

                ac_B = (Mid(Trim(acCode), i, 1))
                ac_A = ac_A + ac_B
            Next i

            aa = " update ACC_CODE set H1 =N'" & ac_A & "'  where ac_code='" & Trim(acCode) & "'  "
            CNN.Execute(aa)
        Next j

        aa = " delete ACC_CODE2   "
        CNN.Execute(aa)
        aa = " INSERT INTO ACC_CODE2  (AC_CODE, Name_L,Name_E )   " & _
             "  select AC_CODE, Name_L, Name_E  from  ACC_CODE order by ac_code  "
        CNN.Execute(aa)

        aa = " update ACC_CODE set H1_nm =ACC_CODE2.Name_L from ACC_CODE2  where ACC_CODE2.ac_code=ACC_CODE.H1  "
        CNN.Execute(aa)
    End Sub

    Public Sub LoadData()
        fg.Rows.Clear()
        With rs
            Call LoadSqlData("SELECT   * from ACC_CODE  order by  AC_CODE", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    fg.Rows.Add(.AbsolutePosition, _
                                (.Fields("AC_CODE").Value.ToString), _
                                (.Fields("Name_L").Value.ToString), _
                                (.Fields("H1").Value.ToString), _
                                (.Fields("H1_nm").Value.ToString))
                    .MoveNext()
                End While
            End If
        End With
    End Sub

    Private Sub fg_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles fg.CellClick
        If e.RowIndex < 0 Then Exit Sub
        
        Txt_ID.Enabled = False
        Txt_ID.Text = If(fg.Rows(e.RowIndex).Cells(1).Value Is Nothing, "", fg.Rows(e.RowIndex).Cells(1).Value.ToString())
        txtsym.Text = If(fg.Rows(e.RowIndex).Cells(2).Value Is Nothing, "", fg.Rows(e.RowIndex).Cells(2).Value.ToString())
        Txt_name_L.Text = If(fg.Rows(e.RowIndex).Cells(3).Value Is Nothing, "", fg.Rows(e.RowIndex).Cells(3).Value.ToString())
        Txt_name_E.Text = If(fg.Rows(e.RowIndex).Cells(4).Value Is Nothing, "", fg.Rows(e.RowIndex).Cells(4).Value.ToString())
    End Sub

    Private Sub fg_CellDoubleClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles fg.CellDoubleClick
        If e.RowIndex < 0 Then Exit Sub
        Txt_ID.Enabled = False
        Txt_ID.Text = If(fg.Rows(e.RowIndex).Cells(1).Value Is Nothing, "", fg.Rows(e.RowIndex).Cells(1).Value.ToString())
        txtsym.Text = If(fg.Rows(e.RowIndex).Cells(2).Value Is Nothing, "", fg.Rows(e.RowIndex).Cells(2).Value.ToString())
        Txt_name_L.Text = If(fg.Rows(e.RowIndex).Cells(3).Value Is Nothing, "", fg.Rows(e.RowIndex).Cells(3).Value.ToString())
        Txt_name_E.Text = If(fg.Rows(e.RowIndex).Cells(4).Value Is Nothing, "", fg.Rows(e.RowIndex).Cells(4).Value.ToString())
    End Sub

    Private Sub fg_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fg.SelectionChanged
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

    Private Sub BtnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew.Click
        Txt_ID.Text = ""
        Txt_name_L.Text = ""
        Txt_name_E.Text = ""
        txtsym.Text = ""
        txtpersen.Text = "0"
        Txt_ID.Visible = True
        Txt_ID.Enabled = True
        txtsym.Text = ""
        AutoNumber()
    End Sub

    Private Sub TxtPV_NM_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmb_Sections.SelectedIndexChanged
        Dim RSC As New ADODB.Recordset
        Call LoadSqlData("Select * From AP_Sections Where Sec_nmL=N'" & Trim(Cmb_Sections.Text) & "'   ", RSC)
        If RSC.RecordCount > 0 Then
            txtSection_ID.Text = Trim(RSC("Sec_id").Value)
        End If
        AutoNumber()
        Txt_name_L.Text = ""
        Call LoadData()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        If fg.CurrentRow Is Nothing Then Exit Sub
        Dim acCode As String = If(fg.CurrentRow.Cells(1).Value Is Nothing, "", fg.CurrentRow.Cells(1).Value.ToString())
        If MessageBox.Show("ທ່ານຕ້ອງການລຶບລາຍການນີ້: " & Trim(acCode) & "  ນີ້ແທ້ບໍ່?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            CNN.Execute("Delete From AP_Donnor Where  don_id=N'" & Trim(acCode) & "' ")
            Call LoadData()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        fg.Rows.Clear()
        fg.Columns.Clear()
        fg.Columns.Add("No", "ລດ")
        fg.Columns.Add("office_id", "ລະ​ຫັດ")
        fg.Columns.Add("AC_CODE", "ເລກບັນຊີໃນສາລະບານ")
        fg.Columns.Add("gen", "ເລກບັນຊີໃນການໂອນ ບໍ່ມີໃນສາລະບານ")

        With rs
            Call LoadSqlData("SELECT    dbo.gen_jn.office_id, dbo.ACC_CODE.AC_CODE, dbo.gen_jn.ac_code AS gen  " & _
                             "FROM         dbo.ACC_CODE RIGHT OUTER JOIN " & _
                             "   dbo.gen_jn ON dbo.ACC_CODE.AC_CODE = dbo.gen_jn.ac_code " & _
                             " WHERE     (dbo.ACC_CODE.AC_CODE IS NULL  OR dbo.ACC_CODE.AC_CODE = '') " & _
                             " GROUP BY dbo.gen_jn.office_id, dbo.ACC_CODE.AC_CODE, dbo.gen_jn.ac_code  ORDER BY gen  ", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    fg.Rows.Add(.AbsolutePosition, _
                                (.Fields("office_id").Value.ToString), _
                                (.Fields("AC_CODE").Value.ToString), _
                                (.Fields("gen").Value.ToString))
                    .MoveNext()
                End While
            End If
        End With
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        fg.Rows.Clear()
        fg.Columns.Clear()
        fg.Columns.Add("No", "ລດ")
        fg.Columns.Add("office_id", "ລະ​ຫັດ")
        fg.Columns.Add("AC_CODE", "ເລກບັນຊີໃນສາລະບານ")
        fg.Columns.Add("gen", "ເລກບັນຊີໃນການໂອນ ບໍ່ມີໃນສາລະບານ")
        fg.Columns.Add("AC_CodeTY", "ເລກບັນຊີໃນການແມັດກັບສາລະບານ")

        With rs
            Call LoadSqlData("SELECT    dbo.gen_jn.office_id, dbo.ACC_CODE.AC_CODE, dbo.gen_jn.ac_code AS gen, dbo.gen_jn.AC_CodeTY, dbo.ACC_CODE.chk_id " & _
                             "FROM         dbo.ACC_CODE RIGHT OUTER JOIN " & _
                             "   dbo.gen_jn ON dbo.ACC_CODE.AC_CODE = dbo.gen_jn.ac_code " & _
                             " WHERE     (dbo.gen_jn.open_amt_dr + dbo.gen_jn.open_amt_cr + dbo.gen_jn.amt_dr + dbo.gen_jn.amt_cr + dbo.gen_jn.Rem_dr + dbo.gen_jn.Rem_cr) > 0 " & _
                             "  AND (dbo.ACC_CODE.AC_CODE IS NULL OR   dbo.ACC_CODE.AC_CODE = '') " & _
                             " GROUP BY dbo.gen_jn.office_id, dbo.ACC_CODE.AC_CODE, dbo.gen_jn.ac_code, dbo.gen_jn.AC_CodeTY, dbo.ACC_CODE.chk_id ORDER BY gen  ", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    fg.Rows.Add(.AbsolutePosition, _
                                (.Fields("office_id").Value.ToString), _
                                (.Fields("AC_CODE").Value.ToString), _
                                (.Fields("gen").Value.ToString), _
                                (.Fields("AC_CodeTY").Value.ToString))
                    .MoveNext()
                End While
            End If
        End With
    End Sub
End Class