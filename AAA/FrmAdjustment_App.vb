Public Class FrmAdjustment_App

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()

    End Sub

    Private Sub FrmAdjustment_List_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateIn.Value = DateAdd("d", -1, DateAdd("m", DateDiff("m", DateIn.Value, DateIn.Value) + 1, CDate(Month(DateIn.Value) & "/" & Year(DateIn.Value))))
        'SetControlText(Me)

        SetupGrid()
        LdGrp()
        LoadListFG()
        LoadBook()
        Cmb.Items.Clear()
        Call load_Cmb(" SELECT Curr  FROM Curr_For_Rate  ORDER BY cnt ", "Curr", Cmb)
        If Cmb.Items.Count > 0 Then
            Cmb.SelectedIndex = 0
        End If

        ' Handle column visibility based on language setting
        If FmMain.MuLngL.Checked = True Then
            FG.Columns(3).Visible = False  ' Adjustment (EN) column
        Else
            FG.Columns(2).Visible = False  ' Adjustment (LA) column
        End If
        'FG.Columns(17).Visible = False  ' Hidden column
    End Sub

    Private Sub SetupGrid()
        ' Clear and setup DataGridView columns
        FG.Columns.Clear()
        FG.Columns.Add("No", "No.")
        FG.Columns.Add("Code", "Code")
        FG.Columns.Add("AdjustmentLA", "Adjustment (LA)")
        FG.Columns.Add("AdjustmentEN", "Adjustment (EN)")
        FG.Columns.Add("DateIN", "Date IN")
        FG.Columns.Add("Period", "Period")
        FG.Columns.Add("AdjustValue", "Adjust Value")
        FG.Columns.Add("RemainValue", "Remain Value")
        FG.Columns.Add("Description", "Description")
        FG.Columns.Add("Dr", "Dr")
        FG.Columns.Add("Cr", "Cr")
        FG.Columns.Add("LastAdjustDate", "Last Adjust Date")
        FG.Columns.Add("ExpectDay", "Expect Day")
        FG.Columns.Add("ExpectValueToAdjust", "Expect Value to Adjust")
        FG.Columns.Add("ExpectRemainValue", "Expect Remain Value")
        FG.Columns.Add("Select", "Select")  ' Boolean column
        FG.Columns.Add("Currency", "Currency")
        FG.Columns.Add("ExchangeRate", "Exchange Rate")
        FG.Columns.Add("ForeignDr", "Foreign Dr")
        FG.Columns.Add("ForeignCr", "Foreign Cr")

        ' Set column widths
        FG.Columns(0).Width = 50
        FG.Columns(1).Width = 80
        FG.Columns(2).Width = 150
        FG.Columns(3).Width = 150
        FG.Columns(4).Width = 100
        FG.Columns(5).Width = 80
        FG.Columns(6).Width = 100
        FG.Columns(7).Width = 100
        FG.Columns(8).Width = 120
        FG.Columns(9).Width = 100
        FG.Columns(10).Width = 100
        FG.Columns(11).Width = 100
        FG.Columns(12).Width = 80
        FG.Columns(13).Width = 150
        FG.Columns(14).Width = 120
        FG.Columns(15).Width = 50
        FG.Columns(16).Width = 80
        FG.Columns(17).Width = 100
        FG.Columns(18).Width = 100
        FG.Columns(19).Width = 100

        ' Configure DataGridView properties
        FG.AllowUserToAddRows = False
        FG.ReadOnly = True
        FG.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG.MultiSelect = False

        ' Set the Select column as boolean (checkbox)
        FG.Columns(15).DefaultCellStyle.NullValue = False
        FG.Columns(15).ValueType = GetType(Boolean)
    End Sub
    Private Sub LoadBook()
        Dim rst As New ADODB.Recordset
        CmbBook.Items.Clear()
        Comm = New ADODB.Command
        Comm.ActiveConnection = CNN
        Comm.CommandText = "SELECT * FROM books WHERE bookid <> '" & "" & " '"
        rst = Comm.Execute
        If rst.RecordCount <> 0 Then
            While Not rst.EOF()
                CmbBook.Items.Add(Trim(rst.Fields("bookid").Value))
                rst.MoveNext()
            End While
        End If
        CmbBook.Text = "GL"
        LoadSqlData("SELECT * FROM books WHERE bookid = N'" & CmbBook.Text & "'", RSC)
        With RSC
            Do Until .EOF = True
                txtBookName.Text = Trim(.Fields("bookname").Value)
                .MoveNext()
            Loop
        End With


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
        Dim CurNM As String
        If txtGrpNm.SelectedIndex = 0 Then
            GrpNM = ""
        Else
            GrpNM = " AND GrpID=N'" & Trim(txtGrp.Text) & "' "
        End If
        CurNM = " AND Curr=N'" & Trim(Cmb.Text) & "' "

        ' Clear existing rows
        FG.Rows.Clear()

        CNN.Execute("UPDATE Adjustment_List set day=Value/Period where day is nuLL ")
        CNN.Execute("UPDATE Adjustment_List set day=Value/Period  ")
        With RSC
            'select DateDiff(d, DateIn , '2021-11-29')  from Adjustment_List
            StrDate = CDate("01/" & Trim(DateIn.Value.Month.ToString) & "/" & Trim(DateIn.Value.Year.ToString))
            Call LoadSqlData("SELECT *, (DateDiff(d, '" & Format(CDate(StrDate), "yyyy/MM/dd") & "' , '" & Format(CDate(DateIn.Value), "yyyy/MM/dd") & "')+1) as ExpectDay FROM  Adjustment_List where 1=1 " & GrpNM & "  " & CurNM & "and Remain>0 order by Code ASC  ", RSC)

            If .RecordCount > 0 Then
                While Not .EOF
                    Dim Rema As Double = 0
                    ' Add row to DataGridView
                    FG.Rows.Add(.AbsolutePosition, _
                      Trim(CStr(.Fields("Code").Value)), _
                      Trim(CStr(.Fields("Name").Value.ToString)), _
                      Trim(CStr(.Fields("NameE").Value.ToString)), _
                      Format(CDate(Trim(.Fields("DateIn").Value)), "dd/MM/yyyy"), _
                      Trim(CStr(.Fields("Period").Value.ToString)), _
                      Format(CDbl(Trim(.Fields("Value").Value)), "##,##0.00"), _
                      Format(CDbl(Trim(.Fields("Remain").Value)), "##,##0.00"), _
                      Trim(CStr(.Fields("Desription").Value.ToString)), _
                      Trim(CStr(.Fields("Dr").Value.ToString)), _
                      Trim(CStr(.Fields("Cr").Value.ToString)), _
                      Format(CDate(DateIn.Value), "dd/MM/yyyy"), _
                      Trim(CStr(.Fields("ExpectDay").Value.ToString)), _
                      0, _  ' Expect Value to Adjust
                      Format(CDbl(Rema), "##,##0.00"), _  ' Expect Remain Value
                      False, _  ' Default value for the boolean select column
                      Trim(CStr(.Fields("Curr").Value.ToString)), _
                      Format(CDbl(Trim(.Fields("RAte").Value)), "##,##0.00"), _
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

        If TxtCode.Text = "" Then MsgBox("", MsgBoxStyle.Exclamation) : TxtCode.Focus() : Exit Sub

        If TxtCode.Enabled = True Then
            Call LoadSqlData("SELECT * FROM Adjustment_List WHERE Code =N'" & Trim(TxtCode.Text) & "'", RSC)
            If RSC.RecordCount <> 0 Then
                MsgBox("ລະຫັດມີແລ້ວ!", MsgBoxStyle.Exclamation) : TxtCode.Focus() : Exit Sub
            End If
        End If


        Call LoadSqlData("SELECT * FROM Adjustment_List WHERE Code =N'" & Trim(TxtCode.Text) & "'", RSC)
        If RSC.RecordCount = 0 Then
            CNN.Execute("INSERT INTO Adjustment_List(Code, GrpID, GrpIDNm, Name, NameE, Desription, DateIn, Period, Value, Remain, Dr, DrNm, Cr, CrNm) " & _
                "Values(N'" & Trim(TxtCode.Text) & "',N'" & Trim(txtGrp.Text) & "',N'" & Trim(txtGrpNm.Text) & "' ,N'" & Trim(TxtName.Text) & "' , N'" & Trim(TxtNameE.Text) & "' ,N'" & Trim(TxtDesription.Text) & "','" & Format(DateIn.Value, "yyyy-MM-dd") & "'," & CDbl(TxtPeriod.Text) & "," & CDbl(TxtValue.Text) & "," & CDbl(TxtRemain.Text) & ",N'" & Trim(TxtDr.Text) & "',N'" & Trim(TxtDrNm.Text) & "',N'" & Trim(TxtCr.Text) & "',N'" & Trim(TxtCrNm.Text) & "')")
        Else
            CNN.Execute("DELETE Adjustment_List WHERE Code =N'" & Trim(TxtCode.Text) & "' ")
            CNN.Execute("INSERT INTO Adjustment_List(Code, GrpID, GrpIDNm, Name, NameE, Desription, DateIn, Period, Value, Remain, Dr, DrNm, Cr, CrNm) " & _
             "Values(N'" & Trim(TxtCode.Text) & "',N'" & Trim(txtGrp.Text) & "',N'" & Trim(txtGrpNm.Text) & "' ,N'" & Trim(TxtName.Text) & "' , N'" & Trim(TxtNameE.Text) & "' ,N'" & Trim(TxtDesription.Text) & "','" & Format(DateIn.Value, "yyyy-MM-dd") & "'," & CDbl(TxtPeriod.Text) & "," & CDbl(TxtValue.Text) & "," & CDbl(TxtRemain.Text) & ",N'" & Trim(TxtDr.Text) & "',N'" & Trim(TxtDrNm.Text) & "',N'" & Trim(TxtCr.Text) & "',N'" & Trim(TxtCrNm.Text) & "')")

        End If
        If RSC.State = ConnectionState.Open Then RSC.Close()
        MsgBox("ການບັນທຶກສຳເລັດ!", MsgBoxStyle.OkOnly)
        TxtCode.Focus()
        LoadListFG()
    End Sub

    Private Sub FG_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FG.CellClick
        ' Handle checkbox column (column 15) click
        If e.ColumnIndex = 15 AndAlso e.RowIndex >= 0 Then
            ' Toggle the checkbox value
            Dim currentValue As Boolean = False
            If FG.Rows(e.RowIndex).Cells(15).Value IsNot Nothing Then
                Boolean.TryParse(FG.Rows(e.RowIndex).Cells(15).Value.ToString(), currentValue)
            End If
            FG.Rows(e.RowIndex).Cells(15).Value = Not currentValue
        End If
    End Sub

    Private Sub FG_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles FG.SelectionChanged, FG.Click
        If FG.CurrentRow Is Nothing Then Exit Sub
        If FG.CurrentRow.Index < 0 Then Exit Sub

        Try
            TxtCode.Text = FG.CurrentRow.Cells(1).Value.ToString()
            TxtName.Text = FG.CurrentRow.Cells(2).Value.ToString()
            'Call LoadText()
            TxtCode.Enabled = False
        Catch ex As Exception
            ' Handle potential conversion errors or empty cells
        End Try
    End Sub
    Private Sub LoadText()
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
        End If
    End Sub

    Private Sub txtGrpNm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtGrpNm.SelectedIndexChanged
        Dim gRS As New ADODB.Recordset
        If Lang = True Then
            Call LoadSqlData("select * from Groups Where Group_NmE=N'" & Trim(txtGrpNm.Text) & "'", gRS)
            If gRS.RecordCount <> 0 Then
                txtGrp.Text = Trim(gRS.Fields("Group_ID").Value.ToString)
            End If
        Else
            Call LoadSqlData("select * from Groups Where Group_Nm=N'" & Trim(txtGrpNm.Text) & "' ", gRS)
            If gRS.RecordCount <> 0 Then
                txtGrp.Text = Trim(gRS.Fields("Group_ID").Value.ToString)
            End If
        End If
        TxtName.Focus()
        LoadListFG()
        FGCal()
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
        LoadListFG()
        FGCal()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If txtGrpNm.SelectedIndex = 0 Then
            MsgBox("ກະລຸນາເລືອກໝວດຊັບສິນກ່ອນ!", MsgBoxStyle.Exclamation) : txtGrpNm.Focus() : Exit Sub
        End If
        If MessageBox.Show("ທ່ານຕ້ອງການໂອນໄປບັນຊີແທ້ ຫຼື ບໍ່ ! ", "ຢັ້ງຢືນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            For i = 0 To FG.Rows.Count - 1
                ' FG.Row = i (Not needed in DataGridView for manual loop)
                If CBool(FG.Rows(i).Cells(15).Value) = True Then
                    Dim MDcertify As String
                    MDcertify = CmbBook.Text & "." & Trim(FG.Rows(i).Cells(1).Value.ToString()) & "." & Format(CDate(DateIn.Value), "dd/MM/yyyy")
                    '====== Dr =========
                    Dim DeGen As String = "Delete from AP_ACC_Gen  where certify=N'" & Trim(MDcertify) & "' and office_id='" & MuSubOff2 & "' and  date_work='" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "'  "
                    CNN.Execute(DeGen)
                    Dim De As String = "Delete from AP_ACC_Gen_Item where certify=N'" & Trim(MDcertify) & "' and  office_id='" & MuSubOff2 & "'  and  date_work='" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "' "
                    CNN.Execute(De)
                    Dim Dejn As String = "Delete from gen_jn where certify=N'" & Trim(MDcertify) & "' and  office_id='" & MuSubOff2 & "' and  date_work='" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "' "
                    CNN.Execute(Dejn)

                    If CDbl(FG.Rows(i).Cells(13).Value) <> 0 Then
                        'CNN.Execute("INSERT INTO gen_jn(certify,Referno, Book,date_work, code_dr,code_cr,ac_code,ac_name,descrip,amount, amount_dr,amount_cr,amt_dr,amt_Cr, curr,rate,curr_i,rate_i, net_amt,my_lock,don_id,Com_id,Office_ID, last_update,last_user) " & _
                        '                    " VALUES('" & MDcertify & "','" & MDcertify & "','" & CmbBook.Text & "','" & Format(DateIn.Value, "yyyy-MM-dd") & "','" & (FG.get_TextMatrix(FG.Row, 9)) & "','','" & (FG.get_TextMatrix(FG.Row, 9)) & "','',''," & CDbl(FG.get_TextMatrix(FG.Row, 13)) & "," & CDbl(FG.get_TextMatrix(FG.Row, 13)) & ",0," & CDbl(FG.get_TextMatrix(FG.Row, 13)) & ",0,'LAK','1','LAK','1'," & CDbl(FG.get_TextMatrix(FG.Row, 7)) & ",'1','01','" & Trim(KK) & "','" & Trim(KK) & "','" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "')")
                        '          Dim CNDR As String = "INSERT INTO gen_jn(certify,Descrip,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                        '        " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id,AG,Frm) " & _
                        '          " VALUES(N'" & Trim(MDcertify) & "'," & _
                        '               "N'" & (FG.Rows(i).Cells(8).Value.ToString()) & "'," & _
                        '        " '" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "'," & _
                        '           "N'" & CmbBook.Text & "'," & _
                        '          "N'" & Trim(MDcertify) & "'," & _
                        '            "N'" & (FG.Rows(i).Cells(1).Value.ToString()) & "'," & _
                        '                         "N''," & _
                        '                       "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) & "," & _
                        '              "N'" & (FG.Rows(i).Cells(16).Value.ToString()) & "'," & _
                        '               "" & CDbl(FG.Rows(i).Cells(17).Value.ToString()) & "," & _
                        '                   "N'" & (FG.Rows(i).Cells(16).Value.ToString()) & "'," & _
                        '               "" & CDbl(FG.Rows(i).Cells(17).Value.ToString()) & "," & _
                        '                  "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(FG.Rows(i).Cells(17).Value.ToString()) & "," & _
                        '          "N'" & (FG.Rows(i).Cells(9).Value.ToString()) & "'," & _
                        '           "N''," & _
                        '         "N'" & (FG.Rows(i).Cells(9).Value.ToString()) & "'," & _
                        '         "N''," & _
                        '          "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) & "," & _
                        '          " 0," & _
                        '               "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(FG.Rows(i).Cells(17).Value.ToString()) & "," & _
                        '          " 0," & _
                        '             " 0," & _
                        '                " 0," & _
                        '           " 1," & _
                        '               " 1," & _
                        '          " Getdate()," & _
                        '        "N'" & MUserID & "'," & _
                        '        "N'" & MuSubOff2 & "',0,'1' )"
                        '          CNN.Execute(CNDR)
                        '          '====== Cr =========
                        '          Dim CNCr As String = "INSERT INTO gen_jn(certify,Descrip,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                        '        " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id,AG,Frm) " & _
                        '          " VALUES(N'" & Trim(MDcertify) & "'," & _
                        '              "N'" & (FG.Rows(i).Cells(8).Value.ToString()) & "'," & _
                        '        " '" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "'," & _
                        '       "N'" & CmbBook.Text & "'," & _
                        '          "N'" & Trim(MDcertify) & "'," & _
                        '           "N'" & (FG.Rows(i).Cells(1).Value.ToString()) & "'," & _
                        '                         "N''," & _
                        '                       "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) & "," & _
                        '               "N'" & (FG.Rows(i).Cells(16).Value.ToString()) & "'," & _
                        '               "" & CDbl(FG.Rows(i).Cells(17).Value.ToString()) & "," & _
                        '                   "N'" & (FG.Rows(i).Cells(16).Value.ToString()) & "'," & _
                        '               "" & CDbl(FG.Rows(i).Cells(17).Value.ToString()) & "," & _
                        '                  "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(FG.Rows(i).Cells(17).Value.ToString()) & "," & _
                        '                                     "N''," & _
                        '          "N'" & (FG.Rows(i).Cells(10).Value.ToString()) & "'," & _
                        '         "N'" & (FG.Rows(i).Cells(10).Value.ToString()) & "'," & _
                        '         "N''," & _
                        '" 0," & _
                        '          "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) & "," & _
                        '          " 0," & _
                        '            "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(FG.Rows(i).Cells(17).Value.ToString()) & "," & _
                        '          " 0," & _
                        '             " 0," & _
                        '           " 1," & _
                        '               " 1," & _
                        '          " Getdate()," & _
                        '        "N'" & MUserID & "'," & _
                        '        "N'" & MuSubOff2 & "',0,'1')"
                        '          CNN.Execute(CNCr)
                        '============CNDR_Curr===============

                        If (FG.Rows(i).Cells(16).Value.ToString()) <> "LAK" Then
                            Dim CNDR_Curr As String = "INSERT INTO gen_jn(certify,Descrip,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                         " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id,AG,Frm) " & _
                           " VALUES(N'" & Trim(MDcertify) & "'," & _
                                "N'" & (FG.Rows(i).Cells(8).Value.ToString()) & "'," & _
                         " '" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "'," & _
                            "N'" & CmbBook.Text & "'," & _
                           "N'" & Trim(MDcertify) & "'," & _
                             "N'" & (FG.Rows(i).Cells(1).Value.ToString()) & "'," & _
                                          "N''," & _
                                        "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) & "," & _
                               "N'" & (FG.Rows(i).Cells(16).Value.ToString()) & "'," & _
                                "" & CDbl(txtRate.Text) & "," & _
                                    "N'" & (FG.Rows(i).Cells(16).Value.ToString()) & "'," & _
                                "" & CDbl(txtRate.Text) & "," & _
                                   "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                           "N'" & (FG.Rows(i).Cells(18).Value.ToString()) & "'," & _
                            "N''," & _
                          "N'" & (FG.Rows(i).Cells(18).Value.ToString()) & "'," & _
                          "N''," & _
                           "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) & "," & _
                           " 0," & _
                                "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                           " 0," & _
                              " 0," & _
                                 " 0," & _
                            " 1," & _
                                " 1," & _
                           " Getdate()," & _
                         "N'" & MUserID & "'," & _
                         "N'" & MuSubOff2 & "',0,'1' )"
                            CNN.Execute(CNDR_Curr)
                            '====== Cr =========
                            Dim CNCr_Curr As String = "INSERT INTO gen_jn(certify,Descrip,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                          " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id,AG,Frm) " & _
                            " VALUES(N'" & Trim(MDcertify) & "'," & _
                                "N'" & (FG.Rows(i).Cells(8).Value.ToString()) & "'," & _
                          " '" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "'," & _
                         "N'" & CmbBook.Text & "'," & _
                            "N'" & Trim(MDcertify) & "'," & _
                             "N'" & (FG.Rows(i).Cells(1).Value.ToString()) & "'," & _
                                           "N''," & _
                                         "" & CDbl(FG.Rows(i).Cells(13).Value) & "," & _
                                 "N'" & (FG.Rows(i).Cells(16).Value.ToString()) & "'," & _
                                 "" & CDbl(txtRate.Text) & "," & _
                                     "N'" & (FG.Rows(i).Cells(16).Value.ToString()) & "'," & _
                                 "" & CDbl(txtRate.Text) & "," & _
                                    "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                                                       "N''," & _
                            "N'" & (FG.Rows(i).Cells(19).Value.ToString()) & "'," & _
                           "N'" & (FG.Rows(i).Cells(19).Value.ToString()) & "'," & _
                           "N''," & _
                  " 0," & _
                            "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) & "," & _
                            " 0," & _
                              "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                            " 0," & _
                               " 0," & _
                             " 1," & _
                                 " 1," & _
                            " Getdate()," & _
                          "N'" & MUserID & "'," & _
                          "N'" & MuSubOff2 & "',0,'1')"
                            CNN.Execute(CNCr_Curr)


                            Dim CNDR As String = "INSERT INTO gen_jn(certify,Descrip,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                                   " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id,AG,Frm) " & _
                                     " VALUES(N'" & Trim(MDcertify) & "'," & _
                                          "N'" & (FG.Rows(i).Cells(8).Value.ToString()) & "'," & _
                                   " '" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "'," & _
                                      "N'" & CmbBook.Text & "'," & _
                                     "N'" & Trim(MDcertify) & "'," & _
                                       "N'" & (FG.Rows(i).Cells(1).Value.ToString()) & "'," & _
                                                    "N''," & _
                                                "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                                                 "N'LAK'," & _
                                          "" & 1 & "," & _
                                              "N'LAK'," & _
                                          "" & 1 & "," & _
                                             "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                                     "N'" & (FG.Rows(i).Cells(9).Value.ToString()) & "'," & _
                                      "N''," & _
                                    "N'" & (FG.Rows(i).Cells(9).Value.ToString()) & "'," & _
                                    "N''," & _
                                   "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                                     " 0," & _
                                          "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                                     " 0," & _
                                        " 0," & _
                                           " 0," & _
                                      " 1," & _
                                          " 1," & _
                                     " Getdate()," & _
                                   "N'" & MUserID & "'," & _
                                   "N'" & MuSubOff2 & "',0,'1' )"
                            CNN.Execute(CNDR)
                            '====== Cr =========
                            Dim CNCr As String = "INSERT INTO gen_jn(certify,Descrip,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                          " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id,AG,Frm) " & _
                            " VALUES(N'" & Trim(MDcertify) & "'," & _
                                "N'" & (FG.Rows(i).Cells(8).Value.ToString()) & "'," & _
                          " '" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "'," & _
                         "N'" & CmbBook.Text & "'," & _
                            "N'" & Trim(MDcertify) & "'," & _
                             "N'" & (FG.Rows(i).Cells(1).Value.ToString()) & "'," & _
                                           "N''," & _
                                          "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                                       "N'LAK'," & _
                                 "" & 1 & "," & _
                                            "N'LAK'," & _
                                 "" & 1 & "," & _
                                    "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                                                       "N''," & _
                            "N'" & (FG.Rows(i).Cells(10).Value.ToString()) & "'," & _
                           "N'" & (FG.Rows(i).Cells(10).Value.ToString()) & "'," & _
                           "N''," & _
                  " 0," & _
                               "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                            " 0," & _
                              "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                            " 0," & _
                               " 0," & _
                             " 1," & _
                                 " 1," & _
                            " Getdate()," & _
                          "N'" & MUserID & "'," & _
                          "N'" & MuSubOff2 & "',0,'1')"
                            CNN.Execute(CNCr)
                        Else
                            '==LAK===
                            Dim CNDR As String = "INSERT INTO gen_jn(certify,Descrip,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                                 " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id,AG,Frm) " & _
                                   " VALUES(N'" & Trim(MDcertify) & "'," & _
                                        "N'" & (FG.Rows(i).Cells(8).Value.ToString()) & "'," & _
                                 " '" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "'," & _
                                    "N'" & CmbBook.Text & "'," & _
                                   "N'" & Trim(MDcertify) & "'," & _
                                     "N'" & (FG.Rows(i).Cells(1).Value.ToString()) & "'," & _
                                                  "N''," & _
                                              "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                                               "N'LAK'," & _
                                        "" & 1 & "," & _
                                            "N'LAK'," & _
                                        "" & 1 & "," & _
                                           "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                                   "N'" & (FG.Rows(i).Cells(9).Value.ToString()) & "'," & _
                                    "N''," & _
                                  "N'" & (FG.Rows(i).Cells(9).Value.ToString()) & "'," & _
                                  "N''," & _
                                 "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                                   " 0," & _
                                        "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                                   " 0," & _
                                      " 0," & _
                                         " 0," & _
                                    " 1," & _
                                        " 1," & _
                                   " Getdate()," & _
                                 "N'" & MUserID & "'," & _
                                 "N'" & MuSubOff2 & "',0,'1' )"
                            CNN.Execute(CNDR)
                            '====== Cr =========
                            Dim CNCr As String = "INSERT INTO gen_jn(certify,Descrip,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                          " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id,AG,Frm) " & _
                            " VALUES(N'" & Trim(MDcertify) & "'," & _
                                "N'" & (FG.Rows(i).Cells(8).Value.ToString()) & "'," & _
                          " '" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "'," & _
                         "N'" & CmbBook.Text & "'," & _
                            "N'" & Trim(MDcertify) & "'," & _
                             "N'" & (FG.Rows(i).Cells(1).Value.ToString()) & "'," & _
                                           "N''," & _
                                          "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                                       "N'LAK'," & _
                                 "" & 1 & "," & _
                                            "N'LAK'," & _
                                 "" & 1 & "," & _
                                    "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                                                       "N''," & _
                            "N'" & (FG.Rows(i).Cells(10).Value.ToString()) & "'," & _
                           "N'" & (FG.Rows(i).Cells(10).Value.ToString()) & "'," & _
                           "N''," & _
                  " 0," & _
                               "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                            " 0," & _
                              "" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) * CDbl(txtRate.Text) & "," & _
                            " 0," & _
                               " 0," & _
                             " 1," & _
                                 " 1," & _
                            " Getdate()," & _
                          "N'" & MUserID & "'," & _
                          "N'" & MuSubOff2 & "',0,'1')"
                            CNN.Execute(CNCr)
                        End If

                    End If
                    CNN.Execute("update AP_ACC_Gen_Item set  AP_ACC_Gen_Item.descrip=Acc_Code.Name_L, AP_ACC_Gen_Item.ac_name=Acc_Code.Name_L,  AP_ACC_Gen_Item.ac_typee=Acc_Code.Acc_TypeE from Acc_Code,AP_ACC_Gen_Item where AP_ACC_Gen_Item.certify='" & Trim(MDcertify) & "' and AP_ACC_Gen_Item.AC_Code=ACC_Code.AC_Code ")

                    CNN.Execute("update gen_jn set  gen_jn.ac_name=Acc_Code.Name_L, gen_jn.ac_namee=Acc_Code.Name_E from Acc_Code,gen_jn where gen_jn.certify=N'" & Trim(MDcertify) & "' and gen_jn.AC_Code=ACC_Code.AC_Code ")
                    'CNN.Execute("update Adjustment_List set  Remain= " & CDbl(FG.Rows(i).Cells(7).Value.ToString()) & "-" & CDbl(FG.Rows(i).Cells(13).Value.ToString()) & " where Code=N'" & (FG.Rows(i).Cells(1).Value.ToString()) & "' ")
                    CNN.Execute("update Adjustment_List set  Remain= " & CDbl(FG.Rows(i).Cells(14).Value.ToString()) & "  where Code=N'" & (FG.Rows(i).Cells(1).Value.ToString()) & "' ")

                End If
            Next
            MsgBox("Finish")
        End If

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        LoadListFG()
        UPP()
        FGCal()

    End Sub
    Private Sub UPP()

        For i = 0 To FG.Rows.Count - 1
            Dim sk As String = "SELECT top 1 rate_dt,rate FROM  AP_Rate_history  where  curr=N'" & (FG.Rows(i).Cells(16).Value.ToString()) & "'     and rate_dt<='" & Format(DateIn.Value, "yyyy-MM-dd") & "'   order by rate_dt desc  "
            Call LoadSqlData(sk, RSC)
            If RsC.RecordCount <> 0 Then
                'conn.Execute("Update Curr_For_Rate set rate=" & (RSC.Fields("rate").Value) & " where  curr=N'" & (FG.Rows(i).Cells(1).Value.ToString()) & "' ")
                'FG.Rows(i).Cells(17).Value = (RSC.Fields("rate").Value)
                FG.Rows(i).Cells(17).Value = Format(CDbl(RSC.Fields("rate").Value), "#,##0.00")
            Else
                FG.Rows(i).Cells(17).Value = 1
                'conn.Execute("Update Curr_For_Rate set rate=1  where  curr=N'" & (FG.Rows(i).Cells(1).Value.ToString()) & "' ")

            End If

        Next
    End Sub

    Private Sub FGCal()
        For i = 0 To FG.Rows.Count - 1
            If FG.Rows(i).Cells(1).Value IsNot Nothing AndAlso FG.Rows(i).Cells(1).Value.ToString() <> "" Then
                StrDate = CDate("01/" & Trim(DateIn.Value.Month.ToString) & "/" & Trim(DateIn.Value.Year.ToString))
                StrMM = Format(CDate(DateIn.Value), "dd/MM/yyyy")
                Call LoadSqlData("SELECT *, (DateDiff(d, '" & Format(CDate(StrDate), "yyyy/MM/dd") & "' , '" & Format(CDate(DateIn.Value), "yyyy/MM/dd") & "')+1) as ExpectDay FROM  Adjustment_List where 1=1 " & GrpNm & " order by Code ASC  ", RSC)

                '========================
                If Format(CDate(DateIn.Value), "MM/yyyy") = Format(CDate(FG.Rows(i).Cells(4).Value), "MM/yyyy") Then
                    FG.Rows(i).Cells(11).Value = Format(CDate(StrMM), "dd/MM/yyyy")
                    FG.Rows(i).Cells(12).Value = DateDiff(DateInterval.Day, CDate(FG.Rows(i).Cells(4).Value), CDate(FG.Rows(i).Cells(11).Value)) + 1

                    'Label10.Text = DateDiff(DateInterval.Day, DISDATE.Value, RECDATE.Value)
                End If

                Dim D As Double = Format(CDbl(FG.Rows(i).Cells(6).Value) / CDbl(FG.Rows(i).Cells(5).Value), "#,##0.00")
                FG.Rows(i).Cells(13).Value = Format(CDbl(D) * CDbl(FG.Rows(i).Cells(12).Value), "#,##0.00")
                FG.Rows(i).Cells(13).Value = Format(CDbl(D) * CDbl(FG.Rows(i).Cells(12).Value), "#,##0.00")
                Dim AMT As Double = Math.Round(Val(FG.Rows(i).Cells(6).Value / CDbl(FG.Rows(i).Cells(5).Value) * CDbl(FG.Rows(i).Cells(12).Value)), 2)
                FG.Rows(i).Cells(13).Value = Math.Round(Val(FG.Rows(i).Cells(6).Value / CDbl(FG.Rows(i).Cells(5).Value) * CDbl(FG.Rows(i).Cells(12).Value)), 2)
                'Dim AMT As Double = Math.Round(Val(FG.Rows(i).Cells(12).Value.ToString()), 2)

                FG.Rows(i).Cells(13).Value = Format(CDbl(FG.Rows(i).Cells(13).Value), "#,##0.00")

                'FG.Rows(i).Cells(14).Value = Format(CDbl(FG.Rows(i).Cells(7).Value.ToString() - (FG.Rows(i).Cells(13).Value.ToString())), "#,##0.00")
                If CheckBox2.Checked = True Then
                    FG.Rows(i).Cells(13).Value = Format(CDbl(FG.Rows(i).Cells(7).Value), "#,##0.00")
                End If
                FG.Rows(i).Cells(14).Value = Format(CDbl(FG.Rows(i).Cells(7).Value) - CDbl(FG.Rows(i).Cells(13).Value), "#,##0.00")

            End If
        Next i
    End Sub

    Private Sub CmbBook_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbBook.SelectedIndexChanged
        'txtInvoice.Text = CmbBook.Text & MdCertifyId
        LoadSqlData("SELECT * FROM books WHERE bookid = N'" & CmbBook.Text & "'", RSC)
        With RSC
            Do Until .EOF = True
                txtBookName.Text = Trim(.Fields("bookname").Value)
                .MoveNext()
            Loop
        End With
   
    End Sub

    Private Sub Cmb_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmb.SelectedIndexChanged

        'Dim rs As New ADODB.Recordset
        'Call LoadSqlData("Select * From Curr_For_Rate Where   Curr =N'" & Trim(Cmb.Text) & "'", rs)
        'If rs.RecordCount > 0 Then
        '    txtcurr_name2.Text = Trim(rs("Curr_name").Value.ToString)
        'End If

        'MDRate_DT = " and rate_dt<='" & Format(dtActi.Value, "yyyy-MM-dd") & "'  "
        'SS_Curr = " and AP_Rate_history.Curr =N'" & Cmb.Text & "' "
        'Call RateSetting()
        'txtRate.Text = Format(MD_Rate, "#,##0.00")

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
        LoadListFG()
        FGCal()
    End Sub
End Class
