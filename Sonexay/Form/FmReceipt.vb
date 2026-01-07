Public Class FmReceipt
    Dim sql As String
    Dim MDCurrency As String
    Dim MDRate As String
    Dim At As Boolean
    Dim lastCurr As String

#Region "DataGridView Helper Methods"

    ''' <summary>
    ''' Gets cell value safely from DataGridView
    ''' </summary>
    Private Function GetGridValue(grid As DataGridView, row As Integer, col As Integer) As String
        If row >= 0 AndAlso row < grid.RowCount AndAlso col >= 0 AndAlso col < grid.ColumnCount Then
            If grid.Rows(row).Cells(col).Value IsNot Nothing Then
                Return grid.Rows(row).Cells(col).Value.ToString()
            End If
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Sets cell value safely in DataGridView
    ''' </summary>
    Private Sub SetGridValue(grid As DataGridView, row As Integer, col As Integer, value As Object)
        If row >= 0 AndAlso row < grid.RowCount AndAlso col >= 0 AndAlso col < grid.ColumnCount Then
            grid.Rows(row).Cells(col).Value = value
        End If
    End Sub

    ''' <summary>
    ''' Sets up DataGridView with common properties
    ''' </summary>
    Private Sub SetupGrid(grid As DataGridView, ParamArray columns() As String)
        grid.AllowUserToAddRows = False
        grid.AllowUserToDeleteRows = False
        grid.ReadOnly = False
        grid.RowHeadersVisible = False
        
        grid.Columns.Clear()
        For Each col As String In columns
            grid.Columns.Add(col, col)
        Next
        
        grid.AutoResizeColumns()
    End Sub

    ''' <summary>
    ''' Sets up single column width
    ''' </summary>
    Private Sub SetupGridColumn(grid As DataGridView, columnIndex As Integer, width As Integer)
        If columnIndex >= 0 AndAlso columnIndex < grid.ColumnCount Then
            grid.Columns(columnIndex).Width = width
        End If
    End Sub

#End Region

    ''' <summary>
    ''' Configures DataGridView properties for all grids in the form
    ''' </summary>
    Private Sub ConfigureDataGridViewProperties()
        ' Configure FG1
        FG1.AllowUserToAddRows = False
        FG1.AllowUserToDeleteRows = False
        FG1.ReadOnly = True
        FG1.RowHeadersVisible = False
        FG1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG1.MultiSelect = False
        
        ' Configure FG2
        FG2.AllowUserToAddRows = False
        FG2.AllowUserToDeleteRows = False
        FG2.ReadOnly = True
        FG2.RowHeadersVisible = False
        FG2.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG2.MultiSelect = False
        
        ' Configure FGPaper
        FGPaper.AllowUserToAddRows = False
        FGPaper.AllowUserToDeleteRows = False
        FGPaper.ReadOnly = False
        FGPaper.RowHeadersVisible = False
        FGPaper.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
        FGPaper.MultiSelect = False
        
        ' Configure FGRate
        FGRate.AllowUserToAddRows = False
        FGRate.AllowUserToDeleteRows = False
        FGRate.ReadOnly = True
        FGRate.RowHeadersVisible = False
        FGRate.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FGRate.MultiSelect = False
    End Sub

    Private Sub FmReceipt_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Call MdiCNum()
        At = 0
        BtnPreview.Enabled = False
        BtnDelete.Enabled = False
        TextBox11.Focus()
        FmReceipt_List.loadSQL()
        FmReceipt_List.LoadListFG()
    End Sub
    Private Sub AtotoInsert()
        Dim rst As Double
        rst = TextBox11.Text
        Dim i As Integer
        If At = 0 Then
            Exit Sub
        End If
        For i = 0 To FGPaper.RowCount - 1
            If CDbl(rst) >= CDbl(GetGridValue(FGPaper, i, 5)) Then
                SetGridValue(FGPaper, i, 2, Int(CDbl(rst) / CDbl(GetGridValue(FGPaper, i, 5))))
                rst = CDbl(rst) - CDbl(CDbl(GetGridValue(FGPaper, i, 5)) * CDbl(GetGridValue(FGPaper, i, 2)))
                SetGridValue(FGPaper, i, 3, Format(CDbl(GetGridValue(FGPaper, i, 2) * GetGridValue(FGPaper, i, 5)), "##,##0.00"))
                SetGridValue(FGPaper, i, 4, Format(CDbl(GetGridValue(FGPaper, i, 3) * CDbl(Rate.Text)), "##,##0.00"))
            Else
                SetGridValue(FGPaper, i, 2, 0)
            End If
        Next i
        For i = 0 To FGPaper.RowCount - 1
            'MsgBox(GetGridValue(FGPaper, i, 2))
            If GetGridValue(FGPaper, i, 2) = 0 Then
                ' Set white background for columns 1,2,3,4,5
                For col As Integer = 1 To 5
                    If col < FGPaper.ColumnCount Then
                        FGPaper.Rows(i).Cells(col).Style.BackColor = Color.White
                    End If
                Next
            Else
                ' Set LightCyan background for columns 1,2,3,4,5
                For col As Integer = 1 To 5
                    If col < FGPaper.ColumnCount Then
                        FGPaper.Rows(i).Cells(col).Style.BackColor = Color.LightCyan
                    End If
                Next
            End If
        Next i
        Call SumData()

    End Sub
    Private Sub FmReceipt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboBox1.BackColor = Color.White
        If FGPaper.ColumnCount > 5 Then FGPaper.Columns(5).Visible = False
        At = 1
        'ComboBox1.Text = "ການຊື້ເງິນຕາ"

        BtnDelete.Enabled = False
        BtnPreview.Enabled = False
        Cashier.Text = MUserName
        TabControl1.SelectedIndex = 1
        LoadCurr()
        LoadListFGRate()

        BtnDelete.Enabled = False
        ' FormatString replaced by SetupGrid calls in LoadListFG methods



        Cmb.SelectedIndex = 0


        Curr.Text = "LAK"
        Rate.Text = "1.00"
        LoadListFGRatePaper()
        TotalLAK.Text = "0.00"
        TotelAmt.Text = "0.00"
        txtAmt_letter.Text = ""
        RadioButton2.Checked = True
        'TextBox11.ReadOnly = False
        Call SumData()
        
        ' Configure DataGridView properties
        ConfigureDataGridViewProperties()
        
        Me.WindowState = FormWindowState.Maximized
        FGRate.Size = New System.Drawing.Size(172, 427)
        FG1.Size = New System.Drawing.Size(764, 403)
        FGPaper.Size = New System.Drawing.Size(824, 237)
        TextBox11.Focus()
        Dim dt As DataTable = DbHelper.GetDataTable("select Curr_last  from Ap_RateSeting WHERE curr= '" & Curr.Text & "'")
        If dt.Rows.Count <> 0 Then
            lastCurr = DbHelper.GetStr(dt.Rows(0)("Curr_last"))
        End If
        'ComboBox1.Text = MDReceiptType
        ComboBox1.Items.Clear()

        Call load_Cmb(" SELECT bookname FROM  books  where Type='RCTY' ", "bookname", ComboBox1)
        ComboBox1.SelectedIndex = 0
    End Sub


    Private Sub LoadCurr()
        Dim Comm As ADODB.Command
        ' Dim rsat As New ADODB.Recordset ' REMOVED - ADODB migration
        Comm = New ADODB.Command
        Comm.ActiveConnection = CNN
        Comm.CommandText = "SELECT Curr FROM Ap_RateSeting WHERE Curr <> '" & "" & " order by Curr'"
        rsat = Comm.Execute
        If dt.Rows.Count <> 0 Then
            While Not rsat.EOF()
                Cmb.Items.Add(Trim(DbHelper.GetStr(dt.Rows(0)("Curr"))))
                rsat.MoveNext()
            End While
            Cmb.Items.Add("==ທັງຫມົດ==")
        End If
    End Sub
    Public Sub FormateText()


        'Rate.Text = Format(CDbl(Rate.Text), "##,##0.00")
        TotelAmt.Text = Format(CDbl(TotelAmt.Text), "##,##0.00")
        TotalLAK.Text = Format(CDbl(TotalLAK.Text), "##,##0.00")

    End Sub

    Public Sub SumData()
        Dim TotalAm, TotalL, Un As Double
        For i = 0 To FGPaper.RowCount - 1
            Un = Un + CDbl(GetGridValue(FGPaper, i, 2))
        Next
        Unit.Text = Format(CDbl(Un), "##,##0.00")
        For i = 0 To FGPaper.RowCount - 1
            TotalAm = TotalAm + CDbl(GetGridValue(FGPaper, i, 3))
            SetGridValue(FGPaper, i, 3, Format(CDbl(CDbl(GetGridValue(FGPaper, i, 3))), "##,##0.00"))
            'TotalL = TotalL + CDbl(GetGridValue(FGPaper, i, 4))
        Next
        TotelAmt.Text = Format(CDbl(TotalAm), "##,##0.00")
        For i = 0 To FGPaper.RowCount - 1
            TotalL = TotalL + CDbl(GetGridValue(FGPaper, i, 4))
        Next
        TotalLAK.Text = Format(CDbl(TotalL), "##,##0.00")
        TextBox1.Text = Format(CDbl(TextBox11.Text) - CDbl(TotelAmt.Text), "##,##0.00")
    End Sub
    Public Sub loadColor()
        ' Dim rmRS As New ADODB.Recordset ' REMOVED - ADODB migration
        Dim J As Integer
        ' FGPaper.Redraw = False ' DataGridView doesn't need Redraw
        For J = 0 To FGPaper.RowCount - 1
            If CDbl(GetGridValue(FGPaper, J, 2)) > 0 Then
                If 4 < FGPaper.ColumnCount Then
                    FGPaper.Rows(J).Cells(4).Style.BackColor = Color.LightCyan
                End If
            End If
        Next J
        ' FGPaper.Redraw = True ' DataGridView doesn't need Redraw
    End Sub
    Private Sub LoadListFG()
        'loadSQL()
        SetupGrid(FG1, "No", "Receipt_No", "InDate", "Bnk_Ac_Name", "Amt", "Curr", "Amt_In_LAK")
        FG1.Rows.Clear()
        
        With RSC
            Dim dt As DataTable = DbHelper.GetDataTable("select *  from Ap_Receipt WHERE Receipt_No<>''" & sql & " order by Receipt_No")
            If dt.Rows.Count > 0 Then
                While Not .EOF
                    For Each row As DataRow In dt.Rows
                        FG1.Rows.Add(Trim(CStr(row("Receipt_No"))), _
                                     Trim(CStr(row("InDate"))), _
                                     Trim(CStr(row("Bnk_Ac_Name"))), _
                                     Trim(Format(CDbl(row("Amt")), "##,##0.00")), _
                                     Trim(CStr(row("Curr"))), _
                                     Trim(Format(CDbl(row("Amt_In_LAK")), "##,##0.00")))
                    Next
                    .MoveNext()
                End While
            End If
        End With
        SumData()
    End Sub
    Private Sub LoadListFGRatePaper()
        SetupGrid(FGPaper, "No", "Description", "Unit", "Total_Curr", "Total_LAK", "Paper_Value")
        FGPaper.Rows.Clear()
        
        With RSC
            Dim dt2 As DataTable = DbHelper.GetDataTable("select Curr,Paper  from Ap_MoneyPaper Where Curr='" & Curr.Text & "' order by Paper DESC")
            If dt2.Rows.Count > 0 Then
                    FGPaper.Rows.Add(0, "ເງິນໃບ (" & Curr.Text & ") =>  " & Trim(Format(CDbl(dt2.Rows(0)("Paper")), "##,##0.00")), _
                                     "0", "0.00", "0.00", Trim(CDbl(dt2.Rows(0)("Paper"))))
                    .MoveNext()
                End While
            End If
        End With
    End Sub
    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        BtnDelete.Enabled = False
        BtnPreview.Enabled = False
        AutoNumberLAK()
        If ComboBox1.Text = "" Then MessageBox.Show("ກະລຸນນາເລືອກປະເພດກ່ອນ") : ComboBox1.BackColor = Color.Red : Exit Sub
        CNN.Execute("insert into Ap_Receipt (Receipt_No , Receipt_Type , InDate , Bnk_Ac_Code , Bnk_Ac_Name , Amt , Curr , Rate , Amt_In_LAK , Payer , Cashier , Last_User , Send_To , Last_UpDate , Remark , Status , Company ) values('" & Receipt_No.Text & "' , N'" & ComboBox1.Text & "' , '" & Format(Indate.Value, "MM-dd-yyyy") & "' , '" & Bnk_Ac_Code.Text & "' , N'" & Bnk_Ac_Name.Text & "' , '" & CDbl(TotelAmt.Text) & "' , '" & Curr.Text & "' , '" & CDbl(Rate.Text) & "' , '" & CDbl(TotalLAK.Text) & "' , '" & Payment.Text & "' , N'" & Cashier.Text & "' , N'" & MUserName & "' , '" & "No" & "' , '" & Format(MWorkSetting, "MM-dd-yyyy") & "' , '" & Remark.Text & "', '0',N'" & MuSubOff & "')")
        For i = 0 To FGPaper.RowCount - 1
            If CDbl(GetGridValue(FGPaper, i, 2)) > 0 Then
                CNN.Execute("insert into Ap_ReceipItem (Receip_No , Amt , Unit ) values('" & Receipt_No.Text & "' , '" & CDbl(GetGridValue(FGPaper, i, 5)) & "' , '" & CDbl(GetGridValue(FGPaper, i, 2)) & "')")
            End If
        Next
        MessageBox.Show("ການບັນທຶກສຳເລັດ")
     
      

    End Sub
    Private Sub LoadListFGRate()
        SetupGrid(FGRate, "No", "Curr", "Rate")
        FGRate.Rows.Clear()
        
        With RSC
            Dim dt3 As DataTable = DbHelper.GetDataTable("select Curr,Rate  from Ap_RateSeting order by Curr")
            If dt3.Rows.Count > 0 Then
                    For Each row As DataRow In dt3.Rows
                        FGRate.Rows.Add(Trim(CStr(row("Curr"))), _
                                        Trim(Format(CDbl(row("Rate")), "##,##0.00")))
                    Next
                    .MoveNext()
                End While
            End If
        End With
    End Sub
    Private Sub LoadListFG2()
        SetupGrid(FG2, "No", "Amt", "Unit", "Total")
        FG2.Rows.Clear()
        
        With RSC
            Dim dt4 As DataTable = DbHelper.GetDataTable("select *  from Ap_ReceipItem WHERE Receip_No='" & GetGridValue(FG1, FG1.CurrentCell.RowIndex, 1) & "' order by Amt DESC")
            If dt4.Rows.Count > 0 Then
                    For Each row As DataRow In dt4.Rows
                        FG2.Rows.Add(Trim(Format(CDbl(row("Amt")), "##,##0.00")), _
                                     Trim(CStr(row("Unit"))), _
                                     Format(CDbl(CDbl(Trim(CStr(row("Amt")))) * CDbl(Trim(CStr(row("Unit"))))), "##,##0.00"))
                    Next
                    .MoveNext()
                End While
            End If
        End With
    End Sub
    Private Sub AutoNumberLAK()
        Dim dtNum As DataTable = DbHelper.GetDataTable("SELECT top 1  Receipt_No FROM  Ap_Receipt Order by Receipt_No DESC")
        If dtNum.Rows.Count = 0 Then
            Receipt_No.Text = "000001"
        Else
            'mNum = Microsoft.VisualBasic.Right(CDbl(Val(dtNum.Rows(0)("Receipt_No"))), 1) + 1
            mNum = CDbl(Val(dtNum.Rows(0)("Receipt_No"))) + 1
            If Len(CStr(mNum).Trim) = 1 Then
                Receipt_No.Text = "00000" & CStr(mNum)
            ElseIf Len(CStr(mNum).Trim) = 2 Then
                Receipt_No.Text = "0000" & CStr(mNum)
            ElseIf Len(CStr(mNum).Trim) = 3 Then
                Receipt_No.Text = "000" & CStr(mNum)
            ElseIf Len(CStr(mNum).Trim) = 4 Then
                Receipt_No.Text = "00" & CStr(mNum)
            ElseIf Len(CStr(mNum).Trim) = 5 Then
                Receipt_No.Text = "0" & CStr(mNum)
            ElseIf Len(CStr(mNum).Trim) >= 6 Then
                Receipt_No.Text = CStr(mNum)
            End If
        End If
    End Sub

    Private Sub FG1_DblClick(ByVal sender As Object, ByVal e As System.EventArgs)

        '-----------LAK
        'If MDCurrency = "LAK" Then
        '    For i = 1 To FG2.Rows - 1
        '        If FG2.get_TextMatrix(i, 3) = "50,000.00" Then
        '            LAK1.Text = FG2.get_TextMatrix(i, 2)
        '            MessageBox.Show("ok")
        '        ElseIf FG2.get_TextMatrix(i, 3) = "20,000.00" Then
        '            LAK2.Text = FG2.get_TextMatrix(i, 2)
        '        ElseIf FG2.get_TextMatrix(i, 3) = "10,000.00" Then
        '            LAK3.Text = FG2.get_TextMatrix(i, 2)
        '        ElseIf FG2.get_TextMatrix(i, 3) = "5,000.00" Then
        '            LAK4.Text = FG2.get_TextMatrix(i, 2)
        '        ElseIf FG2.get_TextMatrix(i, 3) = "2,000.00" Then
        '            LAK5.Text = FG2.get_TextMatrix(i, 2)
        '        ElseIf FG2.get_TextMatrix(i, 3) = "1,000.00" Then
        '            LAK6.Text = FG2.get_TextMatrix(i, 2)
        '        ElseIf FG2.get_TextMatrix(i, 3) = "1,000.00" Then
        '            LAK7.Text = FG2.get_TextMatrix(i, 2)
        '        End If
        '    Next

        'End If


    End Sub

    Private Sub FG1_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

  
  
    Private Sub FGRate_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FGRate.SelectionChanged
        If FGRate.CurrentCell IsNot Nothing AndAlso FGRate.CurrentCell.RowIndex >= 0 AndAlso FGRate.CurrentCell.ColumnIndex >= 0 Then
            If FGRate.RowCount > 0 AndAlso GetGridValue(FGRate, 0, 1) <> "" Then
                Curr.Text = GetGridValue(FGRate, FGRate.CurrentCell.RowIndex, 1)
                Rate.Text = GetGridValue(FGRate, FGRate.CurrentCell.RowIndex, 2)
                LoadListFGRatePaper()
                TotalLAK.Text = "0.00"
                TotelAmt.Text = "0.00"
                txtAmt_letter.Text = ""


                Dim dtLast As DataTable = DbHelper.GetDataTable("select Curr_last  from Ap_RateSeting WHERE curr= '" & Curr.Text & "'")
                If dtLast.Rows.Count <> 0 Then
                    lastCurr = DbHelper.GetStr(dtLast.Rows(0)("Curr_last"))
                End If


            End If
        End If
    End Sub

 

    Private Sub FGPaper_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

  
    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        AutoNumberLAK()
        BtnDelete.Enabled = False
        BtnPreview.Enabled = False
        Receipt_No.Clear()
        Bnk_Ac_Code.Clear()
        Bnk_Ac_Name.Clear()
        Indate.Text = ""
        txtAmt_letter.Clear()
        Remark.Clear()

    End Sub

    Private Sub txtAmt_letter_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmt_letter.TextChanged

    End Sub
    Public Function Letter_amt(ByVal Txt As TextBox, Optional ByVal CurrKIP As Boolean = False) As String

        If Val(TotelAmt.Text) <> 0 Then
            Letter_amt = CMoney(Format(CDbl(Txt.Text), "##0.00"))
        Else
            Letter_amt = ""
        End If

    End Function

    Private Sub TotelAmt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TotelAmt.TextChanged
        txtAmt_letter.Text = Letter_amt(TotelAmt) & ": " & lastCurr
    End Sub

    Private Sub Cmb_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmb.SelectedIndexChanged
        sql = ""
        If Cmb.Text <> "==ທັງຫມົດ==" Then
            sql = " AND Curr = '" & Cmb.Text & "'  AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
        Else
            sql = "   AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
        End If
        'sql = "   AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
        LoadListFG()
    End Sub

  
    Private Sub BtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelete.Click

        If MessageBox.Show("ທ່ານຕ້ອງລຶບລະຫັດ " & Bnk_Ac_Code.Text & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            CNN.Execute("delete from Ap_ReceipItem where Receip_No ='" & GetGridValue(FG1, FG1.CurrentCell.RowIndex, 1) & "' ")
            CNN.Execute("delete from Ap_Receipt where Receipt_No ='" & GetGridValue(FG1, FG1.CurrentCell.RowIndex, 1) & "' ")
            loadSQL()
            LoadListFG()
            FG2.Rows.Clear()
            ' Add empty row if needed
            BtnDelete.Enabled = False
            BtnPreview.Enabled = False
        End If
    End Sub

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Close()

    End Sub

    Private Sub BtnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew.Click
        BtnPreview.Enabled = False
        BtnDelete.Enabled = False
        Receipt_No.Clear()
        Bnk_Ac_Code.Clear()
        Bnk_Ac_Name.Clear()
        Indate.Text = ""
        txtAmt_letter.Clear()
        Remark.Clear()
        TabControl1.SelectedIndex = 1
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        BtnPreview.Enabled = False
        BtnDelete.Enabled = False
        Close()
    End Sub

    Private Sub BtnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        sql = ""
        If FG1.CurrentCell IsNot Nothing Then
            sql = " AND Receip_No = '" & GetGridValue(FG1, FG1.CurrentCell.RowIndex, 1) & "' "
        End If
      

    End Sub

    Private Sub FG1_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG1.SelectionChanged
        If FG1.CurrentCell IsNot Nothing AndAlso FG1.CurrentCell.RowIndex >= 0 AndAlso FG1.CurrentCell.ColumnIndex >= 0 Then
            If FG1.RowCount > 0 AndAlso GetGridValue(FG1, 0, 1) <> "" Then
                BtnDelete.Enabled = True
                BtnPreview.Enabled = True
                LoadListFG2()
            End If
        End If
    End Sub

    Private Sub FGPaper_CellEndEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles FGPaper.CellEndEdit
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim rowIndex As Integer = e.RowIndex
            Dim colIndex As Integer = e.ColumnIndex
            
            If colIndex = 2 Then
                If IsNumeric(GetGridValue(FGPaper, rowIndex, 2)) = False Then MsgBox("ກະລຸນນາປ້ອນເປັນໂຕເລກ") : SetGridValue(FGPaper, rowIndex, 2, "0") : SetGridValue(FGPaper, rowIndex, 3, "0.00") : SetGridValue(FGPaper, rowIndex, 4, "0.00")
                If GetGridValue(FGPaper, rowIndex, 2) = "" Then SetGridValue(FGPaper, rowIndex, 2, "0") : SetGridValue(FGPaper, rowIndex, 3, "0.00") : SetGridValue(FGPaper, rowIndex, 4, "0.00")
                SetGridValue(FGPaper, rowIndex, 3, Format(CDbl(GetGridValue(FGPaper, rowIndex, 2) * GetGridValue(FGPaper, rowIndex, 5)), "##,##0.00"))
                
                If CDbl(GetGridValue(FGPaper, rowIndex, colIndex)) > 0 Then
                    ' Set LightCyan background for columns 1,2,3,4,5
                    For col As Integer = 1 To 5
                        If col < FGPaper.ColumnCount Then
                            FGPaper.Rows(rowIndex).Cells(col).Style.BackColor = Color.LightCyan
                        End If
                    Next
                Else
                    ' Set White background for columns 1,2,3,4,5
                    For col As Integer = 1 To 5
                        If col < FGPaper.ColumnCount Then
                            FGPaper.Rows(rowIndex).Cells(col).Style.BackColor = Color.White
                        End If
                    Next
                End If
            End If
            
            If colIndex = 3 Then
                If IsNumeric(GetGridValue(FGPaper, rowIndex, 3)) = False Then MsgBox("ກະລຸນນາປ້ອນເປັນໂຕເລກ") : SetGridValue(FGPaper, rowIndex, 2, "0") : SetGridValue(FGPaper, rowIndex, 3, "0.00") : SetGridValue(FGPaper, rowIndex, 4, "0.00")
                If GetGridValue(FGPaper, rowIndex, 3) = "" Then SetGridValue(FGPaper, rowIndex, 2, "0") : SetGridValue(FGPaper, rowIndex, 3, "0.00") : SetGridValue(FGPaper, rowIndex, 4, "0.00")
                SetGridValue(FGPaper, rowIndex, 2, CDbl(GetGridValue(FGPaper, rowIndex, 3) / GetGridValue(FGPaper, rowIndex, 5)))
                
                If CDbl(GetGridValue(FGPaper, rowIndex, colIndex)) > 0 Then
                    ' Set LightCyan background for columns 1,2,3,4,5
                    For col As Integer = 1 To 5
                        If col < FGPaper.ColumnCount Then
                            FGPaper.Rows(rowIndex).Cells(col).Style.BackColor = Color.LightCyan
                        End If
                    Next
                Else
                    ' Set White background for columns 1,2,3,4,5
                    For col As Integer = 1 To 5
                        If col < FGPaper.ColumnCount Then
                            FGPaper.Rows(rowIndex).Cells(col).Style.BackColor = Color.White
                        End If
                    Next
                End If
            End If
            
            SetGridValue(FGPaper, rowIndex, 4, Format(CDbl(GetGridValue(FGPaper, rowIndex, 3) * CDbl(Rate.Text)), "##,##0.00"))
            
            ' Move to next row if possible
            If rowIndex + 1 < FGPaper.RowCount Then
                FGPaper.CurrentCell = FGPaper(rowIndex + 1, colIndex)
            End If
            Call SumData()
        End If
    End Sub



    
    Private Sub FGPaper_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FGPaper.SelectionChanged
        If FGPaper.CurrentCell IsNot Nothing AndAlso FGPaper.CurrentCell.RowIndex >= 0 AndAlso FGPaper.CurrentCell.ColumnIndex >= 0 Then
            Dim rowIndex As Integer = FGPaper.CurrentCell.RowIndex
            Dim colIndex As Integer = FGPaper.CurrentCell.ColumnIndex
            
            If colIndex = 2 Then
                If IsNumeric(GetGridValue(FGPaper, rowIndex, 2)) = False Then MsgBox("ກະລຸນນາປ້ອນເປັນໂຕເລກ") : SetGridValue(FGPaper, rowIndex, 2, "0") : SetGridValue(FGPaper, rowIndex, 3, "0.00") : SetGridValue(FGPaper, rowIndex, 4, "0.00")
            End If

            If colIndex = 2 Or colIndex = 3 Then
                ' Additional validation logic if needed
            End If
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        LoadListFGRatePaper()
        'TextBox11.ReadOnly = True

        LoadListFGRatePaper()
        Call SumData()
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
    


        Call AtotoInsert()
        'TextBox11.ReadOnly = False
        Call SumData()
        TextBox11.Focus()
    End Sub

    Private Sub TextBox11_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox11.KeyPress
        If e.KeyChar = Chr(13) Then
            TextBox11.Text = Format(CDbl(TextBox11.Text), "##,##0.00")
            Call AtotoInsert()
        End If
    End Sub

    Private Sub TextBox11_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox11.LostFocus
        TextBox11.Text = Format(CDbl(TextBox11.Text), "##,##0.00")
        If RadioButton2.Checked = True Then
            Call AtotoInsert()
        End If

    End Sub



    Private Sub ComboBox1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.MouseEnter
        ComboBox1.BackColor = Color.White
    End Sub

    Public Sub loadSQL()
        'sql = ""

        'If TextBox25.Text <> "" Then
        '    sql = " AND Bnk_Ac_Code = '" & TextBox25.Text & "' "
        'ElseIf TextBox26.Text <> "" Then
        '    sql = " AND (Bnk_Ac_Name  Like N'%" & TextBox26.Text.Trim & "%')"
        'End If

        'If ComboBox2.Text <> "===ທັງໝົດ===" Then
        '    sql = " AND Receipt_Type = N'" & TextBox25.Text & "' "
        'End If
        'If Cmb.Text <> "==ທັງຫມົດ==" Then
        '    sql = " AND Curr = '" & Cmb.Text & "' "
        'End If
        'If ComboBox3.Text = "ເຄື່ອນໄຫວແລ້ວ" Then
        sql = " AND Status = '1' "
        'End If
        'If ComboBox3.Text = "ຍັງບໍ່ທັນເຄື່ອນໄຫວ" Then
        '    sql = " AND Status = '0' "
        'End If
        ''sql = " AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "' "
    End Sub

    Private Sub TextBox25_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox25.KeyPress
        If e.KeyChar = Chr(13) Then
            sql = ""
            sql = " AND Receipt_No = '" & TextBox25.Text & "' AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
            'sql = "   AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
            LoadListFG()
        End If


    End Sub

    Private Sub TextBox26_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox26.KeyPress
        If e.KeyChar = Chr(13) Then
            sql = ""
            sql = " AND (Bnk_Ac_Name  Like N'%" & TextBox26.Text.Trim & "%')  AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
            'sql = "   AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
            LoadListFG()
        End If
    End Sub

    Private Sub DateTimePicker1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePicker1.KeyPress
        If e.KeyChar = Chr(13) Then
            sql = ""
            sql = " AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "' "

            LoadListFG()
        End If
    End Sub

  

    Private Sub DateTimePicker2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTimePicker2.KeyPress
        If e.KeyChar = Chr(13) Then
            sql = ""
            sql = " AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "' "
            LoadListFG()
        End If
    End Sub


    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        sql = ""
        If ComboBox3.Text = "ເຄື່ອນໄຫວແລ້ວ" Then
            sql = " AND Status = '1'  AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
        Else
            sql = "   AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"

        End If
        If ComboBox3.Text = "ຍັງບໍ່ທັນເຄື່ອນໄຫວ" Then
            sql = " AND Status = '0'  AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
        Else
            sql = "   AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
        End If
        'sql = "   AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'"
        LoadListFG()
    End Sub

    Private Sub TextBox25_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox25.TextChanged

    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged

    End Sub

    Private Sub TextBox26_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox26.TextChanged

    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'TextBox26.Text = Mid(TextBox26.ToString, 1, 1)
        MsgBox(Mid(ComboBox1.Text.ToString, 4, 7))
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        sql = ""
        sql = " AND Ap_Receipt.InDate    BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "' "

        LoadListFG()
    End Sub

    Private Sub BtnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRefresh.Click
        If RadioButton1.Checked = True Then
            LoadListFGRatePaper()
            'TextBox11.ReadOnly = True

            LoadListFGRatePaper()
            Call SumData()
        Else
            Call AtotoInsert()
            'TextBox11.ReadOnly = False
            Call SumData()
            TextBox11.Focus()
        End If
    End Sub

    Private Sub TextBox11_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox11.TextChanged

    End Sub

   

    Private Sub BtnExit_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click

    End Sub
End Class