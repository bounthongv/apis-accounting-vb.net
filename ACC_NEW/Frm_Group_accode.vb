Public Class Frm_Group_accode
    Dim rs As New ADODB.Recordset
    Dim ac_A, ac_B As String
    Private Sub Frm_Department_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Call Loadlang()
        'SetControlText(Me)
        'ChgChildForm()
        'BtnSave.Text = "ຄຳນວນ"
        'If Lang = False Then
        '    Label2.Text = "ແຫລ່ງລາຍຮັບ"
        'Else
        '    Label2.Text = "Donnor "
        'End If
        fg.FormatString = "ລດ |<ລະ​ຫັດ ບັນ​ຊີ |<ຊື່ ບັນ​ຊີ                                   |<ລະ​ຫັດ ບັນ​ຊີ​ແມ່ |<ຊື່ ບັນ​ຊີ​ແມ່  "
        'fg.set_ColHidden(1, True)
        ''fg.set_ColHidden(2, True)
        'Cmb_Sections.Items.Clear()
        'Call load_Cmb("select Sec_nmL from AP_Sections  ", "Sec_nmL", Cmb_Sections)
        'Cmb_Sections.SelectedIndex = 0

        LoadData()
        'BtnAddNew_Click(sender, e)
    End Sub
    Private Sub AutoNumber()
        Dim VIOT As New ADODB.Recordset
        Dim VIOTNEW As String
        'Call LoadSqlData("SELECT top 1 Depart_id from Department where  Sec_id='" & txtSection_ID.Text & "'  Order by DP_ID DESC", VIOT)
        Call LoadSqlData("SELECT top 1 Don_ID from AP_Donnor   Order by Don_ID DESC", VIOT)
        If VIOT.RecordCount <> 0 Then
            VIOTNEW = Format(Val(Mid(VIOT.Fields("Don_ID").Value, 1, 2)) + 1, "00")
        Else
            VIOTNEW = "01"

        End If
        Txt_ID.Text = Trim(CStr(VIOTNEW.ToString))
    End Sub

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        'If Txt_ID.Text = "" Then MsgBox("Type ID !", MsgBoxStyle.OkOnly) : Txt_ID.Focus() : Exit Sub

        Call save()
        MsgBox("ຄຳນວນ ສຳເລັດຜົນ!", MsgBoxStyle.OkOnly)
        Call LoadData()
        'AutoNumber()
        Txt_name_L.Text = ""
        Txt_name_L.Focus()
    End Sub
    Public Sub save()

        Dim aa As String
        Dim AcTypeLao, AcTypeEng As String
        For j = 1 To fg.Rows - 1
            ac_A = ""
            ac_B = ""
            For i = 1 To Len(Trim(fg.get_TextMatrix(j, 1)))
                'MsgBox(Mid(Trim(txtAc_code.Text), i, 1))
                If Mid(Trim(fg.get_TextMatrix(j, 1)), i, 1) = "." Then
                    AcTypeLao = "ບັນຊີຍ່ອຍ (D)"
                    AcTypeEng = "Detail Account"

                    Exit For
                Else
                    AcTypeLao = "ບັນຊີແມ່ (P)"
                    AcTypeEng = "Parent Account"
                End If

                ac_B = (Mid(Trim(fg.get_TextMatrix(j, 1)), i, 1))
                ac_A = ac_A + ac_B
            Next i

            aa = " update ACC_CODE set H1 =N'" & ac_A & "'  where ac_code='" & Trim(fg.get_TextMatrix(j, 1)) & "'  "
            CNN.Execute(aa)
        Next j

        aa = " delete ACC_CODE2   "
        CNN.Execute(aa)
        'aa = " INSERT INTO ACC_CODE2  (AC_CODE, Name_L, Name_E, Get_date, lst_usr, pc_nm)   " & _
        '"  select AC_CODE, Name_L, Name_E, Get_date, lst_usr, pc_nm from  ACC_CODE order by ac_code  "
        'CNN.Execute(aa)
        aa = " INSERT INTO ACC_CODE2  (AC_CODE, Name_L,Name_E )   " & _
 "  select AC_CODE, Name_L, Name_E  from  ACC_CODE order by ac_code  "
        CNN.Execute(aa)

        aa = " update ACC_CODE set H1_nm =ACC_CODE2.Name_L from ACC_CODE2  where ACC_CODE2.ac_code=ACC_CODE.H1  "
        CNN.Execute(aa)
        'aa = " update ACC_CODE set H2_nm =ACC_CODE2.Name_E from ACC_CODE2  where ACC_CODE2.ac_code=ACC_CODE.H1  "
        'CNN.Execute(aa)


        'aa = " update ACC_CODE set H1 =Name_L   where ac_code=ac_code  "
        'Conn.Execute(aa)


        'Call LoadSqlData("SELECT * FROM AP_Donnor WHERE Don_ID = '" & Txt_ID.Text & "'  ", rs)
        'With rs
        '    If .RecordCount = 0 Then
        '        Conn.Execute("INSERT INTO AP_Donnor (Don_ID,Don_Sym,Don_Nm_L,Don_Nm_E,Percen) " & _
        '           " VALUES('" & (Txt_ID.Text) & "'," & _
        '                   " N'" & (txtsym.Text) & "'," & _
        '                        " N'" & (Txt_name_L.Text) & "'," & _
        '                              " N'" & (Txt_name_E.Text) & "'," & _
        '                       " " & CDbl(txtpersen.Text) & ")")
        '    Else
        '        Conn.Execute("UPDATE AP_Donnor SET " & _
        '                                 " Don_Sym=N'" & txtsym.Text & "', " & _
        '                                   " Don_Nm_L=N'" & Txt_name_L.Text & "', " & _
        '                                     " Don_Nm_E=N'" & Txt_name_E.Text & "' ," & _
        '                                       " Percen=" & CDbl(txtpersen.Text) & " " & _
        '           " WHERE Don_ID= '" & (Txt_ID.Text) & "'  ")
        '    End If
        'End With

    End Sub
    Public Sub LoadData()
        fg.Rows = 1
        With rs
            Call LoadSqlData("SELECT   * from ACC_CODE  order by  AC_CODE", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    fg.AddItem(.AbsolutePosition & _
                          Chr(9) & (.Fields("AC_CODE").Value.ToString) & _
                             Chr(9) & (.Fields("Name_L").Value.ToString) & _
                                 Chr(9) & (.Fields("H1").Value.ToString) & _
                               Chr(9) & (.Fields("H1_nm").Value.ToString))
                    .MoveNext()
                End While
            Else
                fg.Rows = 2
            End If
        End With
    End Sub

    Private Sub fg_ClickEvent(ByVal sender As Object, ByVal e As System.EventArgs) Handles fg.ClickEvent

        'Dim aa As String
        'Dim rs As New ADODB.Recordset
        'aa = "SELECT * from AP_Sections where Sec_id='" & fg.get_TextMatrix(fg.Row, 1) & "'  "
        'Call LoadSqlData(aa, rs)
        'With rs
        '    If .RecordCount > 0 Then
        '        Cmb_Sections.Text = Trim(.Fields("Sec_nmL").Value.ToString)
        '    End If
        'End With

        Txt_ID.Enabled = False
        Txt_ID.Text = fg.get_TextMatrix(fg.Row, 1)
        txtsym.Text = fg.get_TextMatrix(fg.Row, 2)
        Txt_name_L.Text = fg.get_TextMatrix(fg.Row, 3)
        Txt_name_E.Text = fg.get_TextMatrix(fg.Row, 4)
        txtpersen.Text = fg.get_TextMatrix(fg.Row, 5)

    End Sub

    Private Sub fg_DblClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles fg.DblClick
        'If fg.Row = 0 Or fg.get_TextMatrix(fg.Row, 1) = "" Then Exit Sub
        Txt_ID.Enabled = False
        Txt_ID.Text = fg.get_TextMatrix(fg.Row, 1)
        txtsym.Text = fg.get_TextMatrix(fg.Row, 2)
        Txt_name_L.Text = fg.get_TextMatrix(fg.Row, 3)
        Txt_name_E.Text = fg.get_TextMatrix(fg.Row, 4)
        txtpersen.Text = fg.get_TextMatrix(fg.Row, 5)
    End Sub

    Private Sub fg_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fg.SelChange

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


    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

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
        If MessageBox.Show("ທ່ານຕ້ອງການລຶບລາຍການນີ້: " & Trim(fg.get_TextMatrix(fg.Row, 1)) & "  ນີ້ແທ້ບໍ່?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            conn.Execute("Delete From AP_Donnor Where  don_id=N'" & Trim(fg.get_TextMatrix(fg.Row, 1)) & "' ")
            Call LoadData()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'Dim ss As String = " SELECT    * from AP_Donnor    where 1=1 order by  don_id  "
        'Call LoadSqlData(ss, rs)
        'If rs.RecordCount = 0 Then MsgBox("Data empty", vbInformation, "Check") : Exit Sub
        'Dim Rpt1 = New CrystalReport_Donnor
        'Dim frm1 = New FrmPreview
        ''Dim myTextObjectOnReport As CrystalDecisions.CrystalReports.Engine.TextObject
        ''myTextObjectOnReport = CType(Rpt1.ReportDefinition.ReportObjects.Item("Text14"), CrystalDecisions.CrystalReports.Engine.TextObject)
        ''myTextObjectOnReport.Text = txtNo_id.Text
        ''myTextObjectOnReport = CType(Rpt1.ReportDefinition.ReportObjects.Item("Text8"), CrystalDecisions.CrystalReports.Engine.TextObject)
        ''myTextObjectOnReport.Text = Cmb_Depart_sub.Text

        ''If OptYear.Checked = True Then

        ''    myTextObjectOnReport = CType(Rpt1.ReportDefinition.ReportObjects.Item("Text20"), CrystalDecisions.CrystalReports.Engine.TextObject)
        ''    myTextObjectOnReport.Text = "ສົກປີ " & Format((DT_year1.Value), "yyyy") & " - " & Format((DT_year2.Value), "yyyy")

        ''End If

        ''myTextObjectOnReport = CType(Rpt1.ReportDefinition.ReportObjects.Item("Text21"), CrystalDecisions.CrystalReports.Engine.TextObject)
        ''myTextObjectOnReport.Text = "ປີ, " & Format((DT_year1.Value), "yyyy") & " - " & Format((DT_year2.Value), "yyyy")

        ''myTextObjectOnReport = CType(Rpt1.ReportDefinition.ReportObjects.Item("Text22"), CrystalDecisions.CrystalReports.Engine.TextObject)
        ''myTextObjectOnReport.Text = "ປີ, " & Format((DT_year1.Value), "yyyy") & " - " & Format((DT_year2.Value), "yyyy")

        'Rpt1.SetDataSource(rs)
        'Rpt1.Refresh()
        'frm1.ReportViewer.ReportSource = Rpt1
        'frm1.ReportViewer.DisplayGroupTree = False
        'frm1.WindowState = FormWindowState.Maximized
        'frm1.Show()
        'Rpt1 = Nothing

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        fg.Rows = 1
        fg.FormatString = "ລດ |<ລະ​ຫັດ |<ເລກບັນຊີໃນສາລະບານ    |<ເລກບັນຊີໃນການໂອນ ບໍ່ມີໃນສາລະບານ    "


        With rs
            '            Call LoadRs("SELECT   dbo.gen_jn.office_id,   dbo.ACC_CODE.AC_CODE, dbo.gen_jn.ac_code AS gen " & _
            '" FROM  dbo.ACC_CODE  RIGHT OUTER JOIN   " & _
            '                 "     dbo.gen_jn ON dbo.ACC_CODE.AC_CODE = dbo.gen_jn.ac_code   " & _
            '          "  where Acc_Code.ac_code Is null     group by  dbo.gen_jn.office_id, ACC_CODE.ac_code, dbo.gen_jn.AC_CODE    order by ac_code asc  ", rs)
            Call LoadSqlData("SELECT    dbo.gen_jn.office_id, dbo.ACC_CODE.AC_CODE, dbo.gen_jn.ac_code AS gen  " & _
  "FROM         dbo.ACC_CODE RIGHT OUTER JOIN " & _
                     "   dbo.gen_jn ON dbo.ACC_CODE.AC_CODE = dbo.gen_jn.ac_code " & _
  " WHERE     (dbo.ACC_CODE.AC_CODE IS NULL  OR dbo.ACC_CODE.AC_CODE = '') " & _
  " GROUP BY dbo.gen_jn.office_id, dbo.ACC_CODE.AC_CODE, dbo.gen_jn.ac_code  ORDER BY gen  ", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    fg.AddItem(.AbsolutePosition & _
                          Chr(9) & (.Fields("office_id").Value.ToString) & _
                             Chr(9) & (.Fields("AC_CODE").Value.ToString) & _
                                 Chr(9) & (.Fields("gen").Value.ToString))
                    .MoveNext()
                End While
            Else
                fg.Rows = 2
            End If
        End With
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        fg.Rows = 1
        fg.FormatString = "ລດ |<ລະ​ຫັດ |<ເລກບັນຊີໃນສາລະບານ    |<ເລກບັນຊີໃນການໂອນ ບໍ່ມີໃນສາລະບານ   |<ເລກບັນຊີໃນການແມັດກັບສາລະບານ   "


        With rs
            Call LoadSqlData("SELECT    dbo.gen_jn.office_id, dbo.ACC_CODE.AC_CODE, dbo.gen_jn.ac_code AS gen, dbo.gen_jn.AC_CodeTY, dbo.ACC_CODE.chk_id " & _
"FROM         dbo.ACC_CODE RIGHT OUTER JOIN " & _
                   "   dbo.gen_jn ON dbo.ACC_CODE.AC_CODE = dbo.gen_jn.ac_code " & _
" WHERE     (dbo.gen_jn.open_amt_dr + dbo.gen_jn.open_amt_cr + dbo.gen_jn.amt_dr + dbo.gen_jn.amt_cr + dbo.gen_jn.Rem_dr + dbo.gen_jn.Rem_cr) > 0 " & _
                    "  AND (dbo.Ap_balance_TB.ac_code IS NULL OR   dbo.Ap_balance_TB.ac_code = '') " & _
" GROUP BY dbo.gen_jn.office_id, dbo.ACC_CODE.AC_CODE, dbo.gen_jn.ac_code, dbo.gen_jn.AC_CodeTY, dbo.ACC_CODE.chk_id ORDER BY gen  ", rs)
            If .RecordCount > 0 Then
                While Not .EOF()
                    fg.AddItem(.AbsolutePosition & _
                          Chr(9) & (.Fields("office_id").Value.ToString) & _
                             Chr(9) & (.Fields("AC_CODE").Value.ToString) & _
                                      Chr(9) & (.Fields("gen").Value.ToString) & _
                                 Chr(9) & (.Fields("AC_CodeTY").Value.ToString))
                    .MoveNext()
                End While
            Else
                fg.Rows = 2
            End If
        End With
    End Sub
End Class