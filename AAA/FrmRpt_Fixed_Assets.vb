Public Class FrmRpt_Fixed_Assets

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()

    End Sub

    Private Sub FrmAdjustment_List_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateIn.Value = DateAdd("d", -1, DateAdd("m", DateDiff("m", DateIn.Value, DateIn.Value) + 1, CDate(Month(DateIn.Value) & "/" & Year(DateIn.Value))))
        'SetControlText(Me)
        'FG.FormatString = "^No. |<Code  |< Adjustment (LA)    |< Adjustment (EN) |^ Date IN  |<Period |> Adjust Value    |>Remain Value   |< Desription |< Dr            |< Cr         |^Last Adjust Date|^Expect Day|>Expect Value to Adjust|>Expect Remain Value|^Select"
        FG.FormatString = "^No. |<Code        |< Description    |<Location  |< Asset Code  |<Model  |<Serial Number|<Invoice|^  Date     |<Currency|<Useful Life|>Purchase Cost|>per Year |>Monthly    |>Prev Month  |>Accumulated|>Net Book Value (NBV)|<Broken Date|<Disposal Date|>Gain or loss o |^Select "

        LdGrp()
        LoadListFG()
        LoadBook()
        Cmb.Items.Clear()
        Call load_Cmb(" SELECT Curr  FROM Curr_For_Rate  ORDER BY cnt ", "Curr", Cmb)
        If Cmb.Items.Count > 0 Then
            Cmb.SelectedIndex = 0
        End If

        FG.set_ColDataType(20, VSFlex8U.DataTypeSettings.flexDTBoolean)
        If FmMain.MuLngL.Checked = True Then
            FG.set_ColHidden(3, True)
        Else
            FG.set_ColHidden(2, True)
        End If
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
        CmbBook.Text = "AS"
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
            Call LoadSqlData("Select * from Groups_Asset Order by Group_ID", gRS)
            If gRS.RecordCount <> 0 Then
                While Not gRS.EOF
                    txtGrpNm.Items.Add(gRS.Fields("Group_NmE").Value.ToString)
                    gRS.MoveNext()
                End While
            End If
            txtGrpNm.SelectedIndex = 0
        Else
            txtGrpNm.Items.Add("ທັງໝົດ ")
            Call LoadSqlData("Select * from Groups_Asset Order by Group_ID", gRS)
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
        If MessageBox.Show("ທ່ານຕ້ອງລຶບລະຫັດ " & FG.get_TextMatrix(FG.Row, 1) & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            CNN.Execute("DELETE FROM Adjustment_List WHERE Code=N'" & FG.get_TextMatrix(FG.Row, 1) & "'")

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
        FG.Rows = 1
        Dim cRS As New ADODB.Recordset
        Dim Str As String = ""
        Dim ss As String
        Dim strDate, EndDate As Date
        Dim mym, myy, myr As Integer
        mym = DateIn.Value.Month + 1
        myy = DateIn.Value.Year
        Dim yy As Integer = 1
        myr = DateIn.Value.Year
        Dim mm As Integer = 0
        Dim mTm As Integer = 0
        'mm = Month(DTMon.Value)
        mm = Month(DateIn.Value) - 1
        strDate = CDate("01/" & Trim(DateIn.Value.Month.ToString) & "/" & Trim(DateIn.Value.Year.ToString))
        If DateIn.Value.Month < 12 Then
            EndDate = DateAdd("d", -1, CDate("01/" & DateIn.Value.Month + 1 & "/" & DateIn.Value.Year))
        Else
            'EndDate = CDate("31/12/" & DateIn.Value.Year)
            EndDate = Format(CDate(DateIn.Value), "dd/MM/yyyy")
        End If


        If txtGrp.Text <> "" Then
            Str = " AND Group_ID='" & Trim(txtGrp.Text) & "'"
        End If

        CNN.Execute("Delete From Rpt_Grp")
        'ss = "Insert into Rpt_Grp(AmountRemain,AmountClear,Amt_All,AssetID, Asset_No, Asset_Nm, Group_ID, Date_Work, Used_Life, Amt, Amt_KIP, Broked,Broke_Date, Deposted, Deposted_Date, Dep_Year, Dep_Month, TTMon, PrevMon, PrevDep, CurrMon, MonDep, TTDep, Remain, strDate, EndDate) " & _
        '    "Select AmountRemain,AmountClear,Amt_All,AssetID, Asset_No, Asset_Nm, Group_ID, Date_Work, Used_Life, Amt, Amt_KIP,Broked, Broke_Date, Deposted, Deposted_Date, Dep_Year, Dep_Month, 0, 0, 0, 0, 0, 0, 0, '" & Format(strDate, "yyyy-MM-dd") & "', '" & Format(EndDate, "yyyy-MM-dd") & "' From Assets Where Date_Work <= '" & Format(EndDate, "yyyy-MM-dd") & "' "
        'ss = ss & Str & " Order by Asset_No "
        'CNN.Execute(ss)

        ss = "Insert into Rpt_Grp(disposal, Curr,Vendor,Using_By,Model,Serial,AmountRemain,AmountClear,Amt_All,AssetID, Asset_No, Asset_Nm, Group_ID, Date_Work, Used_Life, Amt, Amt_KIP, Broked,Broke_Date, Deposted, Deposted_Date, Dep_Year, Dep_Month,Dep_Day, TTMon, PrevMon, PrevDep, CurrMon, MonDep, TTDep, Remain, strDate, EndDate) " & _
      "Select 0, Curr,Vendor,Using_By,Model,Serial,AmountRemain,AmountClear,Amt_All,AssetID, Asset_No, Asset_Nm, Group_ID, Date_Work , Used_Life, Amt, Amt_KIP,Broked, Broke_Date, Deposted, Deposted_Date, Dep_Year, Dep_Month,Dep_Day, 0, 0, 0, 0, 0, 0, 0, '" & Format(strDate, "yyyy-MM-dd") & "', '" & Format(EndDate, "yyyy-MM-dd") & "' " & _
      " From Assets Where Date_Work <= '" & Format(EndDate, "yyyy-MM-dd") & "' "
        ss = ss & Str & " Order by Asset_No "
        CNN.Execute(ss)

        'CNN.Execute("Update Rpt_Grp Set TTMon=Used_Life * 12, PrevMon=DateDiff(m, Date_Work, strDate)-1, CurrMon=DateDiff(m, Date_Work, EndDate) ")
        'CNN.Execute("Update Rpt_Grp Set CurrMon=DateDiff(m, strDate, Deposted_Date)+1 Where Deposted_Date < EndDate AND Deposted_Date is not null")
        'CNN.Execute("Update Rpt_Grp Set PrevMon=0 Where PrevMon < 0")
        'CNN.Execute("Update Rpt_Grp Set MonDep= Dep_Month Where PrevMon > 0 AND TTMon <> PrevMon")
        'CNN.Execute("Update Rpt_Grp Set MonDep= Dep_Month Where DateDiff(m, Date_Work, EndDate) >= 0 ")
        'CNN.Execute("Update Rpt_Grp Set MonDep= 0 Where TTMon=PrevMon ")
        'CNN.Execute("Update Rpt_Grp Set MonDep = 0 Where CurrMon=0 ")


        CNN.Execute("Update Rpt_Grp Set TTMon=Used_Life * 12, PrevMon=DateDiff(day, Date_Work, strDate)   , CurrMon=DateDiff(day, Date_Work, EndDate)+1    ")
        CNN.Execute("Update Rpt_Grp Set TTMon=TTMon*30 ")
        CNN.Execute("Update Rpt_Grp Set CurrMon=DateDiff(day, strDate, Deposted_Date)+1 Where Deposted_Date < EndDate AND Deposted_Date is not null")
        CNN.Execute("Update Rpt_Grp Set PrevMon=0 Where PrevMon < 0")
        CNN.Execute("Update Rpt_Grp Set MonDep= Dep_Day Where PrevMon > 0 AND TTMon <> PrevMon")
        CNN.Execute("Update Rpt_Grp Set MonDep= Dep_Day Where DateDiff(day, Date_Work, EndDate) >= 0 ")
        CNN.Execute("Update Rpt_Grp Set MonDep= 0 Where TTMon=PrevMon ")
        CNN.Execute("Update Rpt_Grp Set MonDep = 0 Where CurrMon=0 ")
        '--------------Depost
        CNN.Execute("Update Rpt_Grp Set PrevMon=DateDiff(day, Date_Work, Deposted_Date) Where Deposted_Date is not null AND PrevMon > DateDiff(day, Date_Work, Deposted_Date) ")
        CNN.Execute("Update Rpt_Grp Set MonDep= 0 Where Deposted_Date is not null AND DateDiff(day, strdate, Deposted_Date) <= 0 ")
        CNN.Execute("Update Rpt_Grp Set Remain = 0 Where Deposted_Date is not null AND DateDiff(day, strdate, Deposted_Date) <= 0 ")
        ' LA Only
        CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep Where Deposted_Date is not null AND DateDiff(day, strdate, Deposted_Date) <= 0 ")
        '------------
        'CNN.Execute("Update Rpt_Grp Set Rpt_Grp.PrevDep=Open_BL.Open_Amt From Rpt_Grp, Open_BL Where Rpt_Grp.AssetID=Open_BL.AssetID AND Open_BL.Open_Amt <>0 and Year(Open_BL.Date_Work)=" & myr & "")
        'If mm >= 1 Then
        '    CNN.Execute("Update Rpt_Grp Set PrevDep = PrevDep + Dep_Month * " & mm & " Where year(Date_Work) < year(strdate) ")
        '    CNN.Execute("Update Rpt_Grp Set PrevDep = Dep_Month * (" & mm & " - month(Date_Work)) Where year(Date_Work) >= year(strdate) ")
        'End If
        CNN.Execute("Update Rpt_Grp Set PrevDep = Dep_Day*PrevMon where PrevMon > 0 ")
        'CNN.Execute("Update Rpt_Grp Set PrevDep = Amt_KIP where PrevDep > Amt_KIP ")
        CNN.Execute("Update Rpt_Grp Set TTDep = Dep_Day * CurrMon ")

        CNN.Execute("Update Rpt_Grp Set MonDep = TTDep - PrevDep ")
        CNN.Execute("Update Rpt_Grp Set MonDep = 0 where MonDep < 0 ")
        CNN.Execute("Update Rpt_Grp Set PrevDep = 0 Where DateDiff(day, Date_Work, strdate) <=1 ")
        'CNN.Execute("Update Rpt_Grp Set TTDep = PrevDep + MonDep ")

        CNN.Execute(" update Rpt_Grp set MonDep=0 where   TTDep>Amt   ")
        CNN.Execute("update Rpt_Grp set PrevDep=Amt,TTDep=Amt where   TTDep>Amt    ")

        CNN.Execute("Update Rpt_Grp Set Remain = Amt_KIP - TTDep ")
        CNN.Execute("Update Rpt_Grp Set Remain = 0 where Remain < 1 ")
        CNN.Execute("Update Rpt_Grp Set Remain = 0, Amt_KIP=0, MonDep=0 Where Deposted_Date <= strdate")
        CNN.Execute("Update Rpt_Grp Set Remain = 0, Amt_KIP=0 Where Deposted_Date <= EndDate")


        With cRS
            'strDate = CDate("01/" & Trim(DateIn.Value.Month.ToString) & "/" & Trim(DateIn.Value.Year.ToString))
            'Call LoadSqlData("SELECT *, (DateDiff(d, '" & Format(CDate(strDate), "yyyy/MM/dd") & "' , '" & Format(CDate(DateIn.Value), "yyyy/MM/dd") & "')+1) as ExpectDay FROM  Adjustment_List where 1=1 " & GrpNM & " and Remain>0 order by Code ASC  ", RSC)
            Call LoadSqlData("SELECT * FROM Rpt_Grp where 1=1  Order by Assetid ", cRS)
            If .RecordCount > 0 Then
                While Not .EOF
                    Dim Rema As Double = 0
                    'FG.FormatString = "^No. |<Code        |< Description    |<Location  |< Asset Code  |<Model  |<Serial Number|<Invoice|^  Date     |<Currency|<Useful Life|>Purchase Cost|>per Year |>Monthly    |>Prev Month  |>Accumulated|>Net Book Value (NBV)|^Select"

                    FG.AddItem(cRS.AbsolutePosition & Chr(9) & Trim(cRS.Fields("AssetID").Value.ToString) & _
                               Chr(9) & Trim(cRS.Fields("Asset_Nm").Value.ToString) & _
                               Chr(9) & Trim(cRS.Fields("Using_By").Value.ToString) & _
                               Chr(9) & Trim(cRS.Fields("Asset_No").Value.ToString) & _
                           Chr(9) & Trim(cRS.Fields("Model").Value.ToString) & _
                              Chr(9) & Trim(cRS.Fields("Serial").Value.ToString) & _
                                   Chr(9) & Trim(cRS.Fields("Vendor").Value.ToString) & _
                                   Chr(9) & Format(cRS.Fields("Date_Work").Value, "dd/MM/yyyy") & _
                                             Chr(9) & Trim(cRS.Fields("Curr").Value.ToString) & _
                                            Chr(9) & cRS.Fields("Used_life").Value.ToString & _
                               Chr(9) & Format(cRS.Fields("Amt_KIP").Value, "#,##0.00") & _
                                    Chr(9) & 20 & _
                       Chr(9) & Format(cRS.Fields("MonDep").Value, "#,##0.00") & _
                       Chr(9) & Format(cRS.Fields("PrevDep").Value, "#,##0.00") & _
                       Chr(9) & Format(cRS.Fields("TTDep").Value, "#,##0.00") & _
                       Chr(9) & Format(cRS.Fields("Remain").Value, "#,##0.00") & _
                            Chr(9) & Trim(cRS.Fields("Broke_Date").Value.ToString) & _
                                 Chr(9) & Trim(cRS.Fields("Deposted_Date").Value.ToString) & _
                                  Chr(9) & Format(cRS.Fields("disposal").Value, "#,##0.00") & _
                      "" & vbTab & 0)
                    ' Chr(9) & cRS.Fields("Deposted_Date").Value & _
                    'Chr(9) & Format(cRS.Fields("Dep_Year").Value, "#,##0.00") & _
                    'Chr(9) & Format(cRS.Fields("Dep_Month").Value, "#,##0.00") & _
                    'Chr(9) & cRS.Fields("TTMon").Value.ToString & _
                    'Chr(9) & cRS.Fields("PrevMon").Value.ToString & _
                    .MoveNext()
                End While
            Else
                FG.Rows = 2
            End If
        End With

    End Sub
    Private Sub Sum()
        Dim i As Integer
        Dim total As Double = 0
        For i = 1 To FG.Rows - 1
            total = total + (CDbl(FG.get_TextMatrix(i, 13)))
        Next i
  
        txtBill_Amt.Text = Format(CDbl(total), "#,##0.00")
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

    Private Sub FG_MouseUpEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseUpEvent) Handles FG.MouseUpEvent
        If FG.Col = 17 Then
            FG.Editable = VSFlex8U.EditableSettings.flexEDKbdMouse
        Else

        End If
    End Sub

    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelChange
        TxtCode.Text = FG.get_TextMatrix(FG.Row, 1)
        TxtName.Text = FG.get_TextMatrix(FG.Row, 2)
        'Call LoadText()
        TxtCode.Enabled = False
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
            Call LoadSqlData("select * from Groups_Asset Where Group_NmE=N'" & Trim(txtGrpNm.Text) & "'", gRS)
            If gRS.RecordCount <> 0 Then
                txtGrp.Text = Trim(gRS.Fields("Group_ID").Value.ToString)
            Else
                txtGrp.Text = ""
            End If
        Else
            Call LoadSqlData("select * from Groups_Asset Where Group_Nm=N'" & Trim(txtGrpNm.Text) & "' ", gRS)
            If gRS.RecordCount <> 0 Then
                txtGrp.Text = Trim(gRS.Fields("Group_ID").Value.ToString)
                TxtDr.Text = Trim(gRS.Fields("AccountCodeAsDR").Value.ToString)
                TxtCr.Text = Trim(gRS.Fields("AccountCodeAsCR").Value.ToString)
            Else
                txtGrp.Text = ""
                TxtDr.Text = ""
                TxtCr.Text = ""
            End If
        End If
        Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtDr.Text & "' ", RSC)
        If RSC.RecordCount <> 0 Then
            TxtDrNm.Text = Trim(RSC.Fields("Name_L").Value.ToString)
        End If

        Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtCr.Text & "' ", RSC)
        If RSC.RecordCount <> 0 Then
            TxtCrNm.Text = Trim(RSC.Fields("Name_L").Value.ToString)
        End If


        TxtName.Focus()
        LoadListFG()
        'FGCal()
        Sum()
    End Sub

    Private Sub BtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch.Click
        fmShartOfAccDetail.txtSty.Text = "FrmRpt_Fixed_Assets_DR"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
        TxtDr.Text = DDDR
        TxtDrNm.Text = DDDRN
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        fmShartOfAccDetail.txtSty.Text = "FrmRpt_Fixed_Assets_CR"
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
        'FGCal()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If txtGrpNm.SelectedIndex = 0 Then
            MsgBox("ກະລຸນາເລືອກໝວດຊັບສິນກ່ອນ!", MsgBoxStyle.Exclamation) : txtGrpNm.Focus() : Exit Sub
        End If
        If MessageBox.Show("ທ່ານຕ້ອງການໂອນໄປບັນຊີແທ້ ຫຼື ບໍ່ ! ", "ຢັ້ງຢືນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            For i = 1 To FG.Rows - 1
                FG.Row = i
                If FG.get_ValueMatrix(i, 17) = True Then
                    Dim MDcertify As String
                    MDcertify = CmbBook.Text & "." & Trim(FG.get_TextMatrix(i, 1)) & "." & Format(CDate(DateIn.Value), "dd/MM/yyyy")
                    '====== Dr =========
                    Dim DeGen As String = "Delete from AP_ACC_Gen  where certify=N'" & Trim(MDcertify) & "' and office_id='" & MuSubOff2 & "' and  date_work='" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "'  "
                    CNN.Execute(DeGen)
                    Dim De As String = "Delete from AP_ACC_Gen_Item where certify=N'" & Trim(MDcertify) & "' and  office_id='" & MuSubOff2 & "'  and  date_work='" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "' "
                    CNN.Execute(De)
                    Dim Dejn As String = "Delete from gen_jn where certify=N'" & Trim(MDcertify) & "' and  office_id='" & MuSubOff2 & "' and  date_work='" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "' "
                    CNN.Execute(Dejn)

                    If CDbl(FG.get_TextMatrix(i, 13)) <> 0 Then
                        'CNN.Execute("INSERT INTO gen_jn(certify,Referno, Book,date_work, code_dr,code_cr,ac_code,ac_name,descrip,amount, amount_dr,amount_cr,amt_dr,amt_Cr, curr,rate,curr_i,rate_i, net_amt,my_lock,don_id,Com_id,Office_ID, last_update,last_user) " & _
                        '                    " VALUES('" & MDcertify & "','" & MDcertify & "','" & CmbBook.Text & "','" & Format(DateIn.Value, "yyyy-MM-dd") & "','" & (FG.get_TextMatrix(FG.Row, 9)) & "','','" & (FG.get_TextMatrix(FG.Row, 9)) & "','',''," & CDbl(FG.get_TextMatrix(FG.Row, 13)) & "," & CDbl(FG.get_TextMatrix(FG.Row, 13)) & ",0," & CDbl(FG.get_TextMatrix(FG.Row, 13)) & ",0,'LAK','1','LAK','1'," & CDbl(FG.get_TextMatrix(FG.Row, 7)) & ",'1','01','" & Trim(KK) & "','" & Trim(KK) & "','" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "')")
                        Dim CNDR As String = "INSERT INTO gen_jn(certify,Descrip,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
                      " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id,AG,Frm) " & _
                        " VALUES(N'" & Trim(MDcertify) & "'," & _
                             "N'" & (FG.get_TextMatrix(i, 2)) & "'," & _
                      " '" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "'," & _
                         "N'" & CmbBook.Text & "'," & _
                        "N'" & Trim(MDcertify) & "'," & _
                          "N'" & (FG.get_TextMatrix(i, 1)) & "'," & _
                                       "N''," & _
                                     "" & CDbl(FG.get_TextMatrix(i, 13)) & "," & _
                          "N'" & Trim(Cmb.Text) & "'," & _
                             "" & CDbl(1) & "," & _
                               "N'" & Trim(Cmb.Text) & "'," & _
                             "" & CDbl(1) & "," & _
                                "" & CDbl(FG.get_TextMatrix(i, 13)) * CDbl(1) & "," & _
                        "N'" & TxtDr.Text & "'," & _
                         "N''," & _
                       "N'" & TxtDr.Text & "'," & _
                       "N''," & _
                        "" & CDbl(FG.get_TextMatrix(i, 13)) & "," & _
                        " 0," & _
                             "" & CDbl(FG.get_TextMatrix(i, 13)) * CDbl(1) & "," & _
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
                            "N'" & (FG.get_TextMatrix(i, 2)) & "'," & _
                      " '" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "'," & _
                     "N'" & CmbBook.Text & "'," & _
                        "N'" & Trim(MDcertify) & "'," & _
                         "N'" & (FG.get_TextMatrix(i, 1)) & "'," & _
                                       "N''," & _
                                     "" & CDbl(FG.get_TextMatrix(i, 13)) & "," & _
                          "N'" & Trim(Cmb.Text) & "'," & _
                             "" & CDbl(1) & "," & _
                               "N'" & Trim(Cmb.Text) & "'," & _
                             "" & CDbl(1) & "," & _
                                "" & CDbl(FG.get_TextMatrix(i, 13)) * CDbl(1) & "," & _
                                                   "N''," & _
                        "N'" & TxtCr.Text & "'," & _
                       "N'" & TxtCr.Text & "'," & _
                       "N''," & _
              " 0," & _
                        "" & CDbl(FG.get_TextMatrix(i, 13)) & "," & _
                        " 0," & _
                          "" & CDbl(FG.get_TextMatrix(i, 13)) * CDbl(1) & "," & _
                        " 0," & _
                           " 0," & _
                         " 1," & _
                             " 1," & _
                        " Getdate()," & _
                      "N'" & MUserID & "'," & _
                      "N'" & MuSubOff2 & "',0,'1')"
                        CNN.Execute(CNCr)
                    End If
                    CNN.Execute("update AP_ACC_Gen_Item set  AP_ACC_Gen_Item.descrip=Acc_Code.Name_L, AP_ACC_Gen_Item.ac_name=Acc_Code.Name_L,  AP_ACC_Gen_Item.ac_typee=Acc_Code.Acc_TypeE from Acc_Code,AP_ACC_Gen_Item where AP_ACC_Gen_Item.certify='" & Trim(MDcertify) & "' and AP_ACC_Gen_Item.AC_Code=ACC_Code.AC_Code ")

                    CNN.Execute("update gen_jn set  gen_jn.ac_name=Acc_Code.Name_L, gen_jn.ac_namee=Acc_Code.Name_E from Acc_Code,gen_jn where gen_jn.certify=N'" & Trim(MDcertify) & "' and gen_jn.AC_Code=ACC_Code.AC_Code ")
                    'CNN.Execute("update Adjustment_List set  Remain= " & CDbl(FG.get_TextMatrix(i, 7)) & "-" & CDbl(FG.get_TextMatrix(i, 13)) & " where Code=N'" & (FG.get_TextMatrix(i, 1)) & "' ")
                    'CNN.Execute("update Adjustment_List set  Remain= " & CDbl(FG.get_TextMatrix(i, 14)) & "  where Code=N'" & (FG.get_TextMatrix(i, 1)) & "' ")

                End If
            Next
            MsgBox("Finish")
        End If

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        LoadListFG()
        'FGCal()
        Sum()
    End Sub
    Private Sub FGCal()
        For i = 1 To FG.Rows - 1
            If FG.get_TextMatrix(i, 1) <> "" Then
                StrDate = CDate("01/" & Trim(DateIn.Value.Month.ToString) & "/" & Trim(DateIn.Value.Year.ToString))
                StrMM = Format(CDate(DateIn.Value), "dd/MM/yyyy")
                Call LoadSqlData("SELECT *, (DateDiff(d, '" & Format(CDate(StrDate), "yyyy/MM/dd") & "' , '" & Format(CDate(DateIn.Value), "yyyy/MM/dd") & "')+1) as ExpectDay FROM  Adjustment_List where 1=1 " & GrpNm & " order by Code ASC  ", RSC)

                '========================
                If Format(CDate(DateIn.Value), "MM/yyyy") = Format(CDate(FG.get_TextMatrix(i, 4)), "MM/yyyy") Then
                    FG.set_TextMatrix(i, 11, Format(CDate(StrMM), "dd/MM/yyyy"))
                    FG.set_TextMatrix(i, 12, DateDiff(DateInterval.Day, CDate(FG.get_TextMatrix(i, 4)), CDate(FG.get_TextMatrix(i, 11))) + 1)

                    'Label10.Text = DateDiff(DateInterval.Day, DISDATE.Value, RECDATE.Value)
                End If

                Dim D As Double = Format(CDbl(FG.get_TextMatrix(i, 6)) / CDbl(FG.get_TextMatrix(i, 5)), "#,##0.00")
                FG.set_TextMatrix(i, 13, Format(CDbl(D) * CDbl(FG.get_TextMatrix(i, 12)), "#,##0.00"))
                FG.set_TextMatrix(i, 13, Format(CDbl(D) * CDbl(FG.get_TextMatrix(i, 12)), "#,##0.00"))
                Dim AMT As Double = Math.Round(Val(FG.get_TextMatrix(i, 6) / CDbl(FG.get_TextMatrix(i, 5)) * CDbl(FG.get_TextMatrix(i, 12))), 2)
                FG.set_TextMatrix(i, 13, Math.Round(Val(FG.get_TextMatrix(i, 6) / CDbl(FG.get_TextMatrix(i, 5)) * CDbl(FG.get_TextMatrix(i, 12))), 2))
                'Dim AMT As Double = Math.Round(Val(FG.get_TextMatrix(i, 12)), 2)

                FG.set_TextMatrix(i, 13, Format(CDbl(FG.get_TextMatrix(i, 13)), "#,##0.00"))

                'FG.set_TextMatrix(i, 14, Format(CDbl(FG.get_TextMatrix(i, 7) - (FG.get_TextMatrix(i, 13))), "#,##0.00"))
                If CheckBox2.Checked = True Then
                    FG.set_TextMatrix(i, 13, Format(CDbl(FG.get_TextMatrix(i, 7)), "#,##0.00"))
                End If
                FG.set_TextMatrix(i, 14, Format(CDbl(FG.get_TextMatrix(i, 7) - (FG.get_TextMatrix(i, 13))), "#,##0.00"))

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

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Sum()
        If txtGrpNm.SelectedIndex = 0 Then
            MsgBox("ກະລຸນາເລືອກໝວດຊັບສິນກ່ອນ!", MsgBoxStyle.Exclamation) : txtGrpNm.Focus() : Exit Sub
        End If
        If TxtDr.Text = "" Then
            MsgBox("ກະລຸນາເລືອກເລກບັນຊີກ່ອນ!", MsgBoxStyle.Exclamation) : TxtDr.Focus() : Exit Sub
        End If
        If TxtCr.Text = "" Then
            MsgBox("ກະລຸນາເລືອກເລກບັນຊີກ່ອນ!", MsgBoxStyle.Exclamation) : TxtCr.Focus() : Exit Sub
        End If

        If MessageBox.Show("ທ່ານຕ້ອງການໂອນໄປບັນຊີແທ້ ຫຼື ບໍ່ ! ", "ຢັ້ງຢືນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then


            Dim MDcertify As String
            MDcertify = CmbBook.Text & "." & Trim(txtGrp.Text) & "." & Format(CDate(DateIn.Value), "dd/MM/yyyy")
            '====== Dr =========
            Dim DeGen As String = "Delete from AP_ACC_Gen  where certify=N'" & Trim(MDcertify) & "' and office_id='" & MuSubOff2 & "' and  date_work='" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "'  "
            CNN.Execute(DeGen)
            Dim De As String = "Delete from AP_ACC_Gen_Item where certify=N'" & Trim(MDcertify) & "' and  office_id='" & MuSubOff2 & "'  and  date_work='" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "' "
            CNN.Execute(De)
            Dim Dejn As String = "Delete from gen_jn where certify=N'" & Trim(MDcertify) & "' and  office_id='" & MuSubOff2 & "' and  date_work='" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "' "
            CNN.Execute(Dejn)

            If CDbl(txtBill_Amt.Text) <> 0 Then
                'CNN.Execute("INSERT INTO gen_jn(certify,Referno, Book,date_work, code_dr,code_cr,ac_code,ac_name,descrip,amount, amount_dr,amount_cr,amt_dr,amt_Cr, curr,rate,curr_i,rate_i, net_amt,my_lock,don_id,Com_id,Office_ID, last_update,last_user) " & _
                '                    " VALUES('" & MDcertify & "','" & MDcertify & "','" & CmbBook.Text & "','" & Format(DateIn.Value, "yyyy-MM-dd") & "','" & (FG.get_TextMatrix(FG.Row, 9)) & "','','" & (FG.get_TextMatrix(FG.Row, 9)) & "','',''," & CDbl(FG.get_TextMatrix(FG.Row, 13)) & "," & CDbl(FG.get_TextMatrix(FG.Row, 13)) & ",0," & CDbl(FG.get_TextMatrix(FG.Row, 13)) & ",0,'LAK','1','LAK','1'," & CDbl(FG.get_TextMatrix(FG.Row, 7)) & ",'1','01','" & Trim(KK) & "','" & Trim(KK) & "','" & Format(Date.Today, "yyyy-MM-dd") & "','" & MUserID & "')")
                Dim CNDR As String = "INSERT INTO gen_jn(certify,Descrip,date_work, book,Referno,Referno_Item, cheque_no,amount,Curr,rate,Curr_i,rate_i,net_amt, code_dr, code_cr, ac_code, ac_name, amount_dr, amount_cr, " & _
              " amt_dr, amt_cr,amt_USD_dr,amt_USD_Cr, my_lock,rec_lock, last_update, last_user, office_id,AG,Frm) " & _
                " VALUES(N'" & Trim(MDcertify) & "'," & _
                     "N'" & (txtGrpNm.Text) & "'," & _
              " '" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "'," & _
                 "N'" & CmbBook.Text & "'," & _
                "N'" & Trim(MDcertify) & "'," & _
                        "N'" & Trim(MDcertify) & "'," & _
                               "N''," & _
                             "" & CDbl(txtBill_Amt.Text) & "," & _
                  "N'" & Trim(Cmb.Text) & "'," & _
                     "" & CDbl(1) & "," & _
                       "N'" & Trim(Cmb.Text) & "'," & _
                     "" & CDbl(1) & "," & _
                        "" & CDbl(txtBill_Amt.Text) * CDbl(1) & "," & _
                "N'" & TxtDr.Text & "'," & _
                 "N''," & _
               "N'" & TxtDr.Text & "'," & _
               "N''," & _
                "" & CDbl(txtBill_Amt.Text) & "," & _
                " 0," & _
                     "" & CDbl(txtBill_Amt.Text) * CDbl(1) & "," & _
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
                "N'" & (txtGrpNm.Text) & "'," & _
              " '" & Format(CDate(DateIn.Value), "yyyy-MM-dd") & "'," & _
             "N'" & CmbBook.Text & "'," & _
                "N'" & Trim(MDcertify) & "'," & _
                 "N'" & Trim(MDcertify) & "'," & _
                               "N''," & _
                             "" & CDbl(txtBill_Amt.Text) & "," & _
                  "N'" & Trim(Cmb.Text) & "'," & _
                     "" & CDbl(1) & "," & _
                       "N'" & Trim(Cmb.Text) & "'," & _
                     "" & CDbl(1) & "," & _
                        "" & CDbl(txtBill_Amt.Text) * CDbl(1) & "," & _
                                           "N''," & _
                "N'" & TxtCr.Text & "'," & _
               "N'" & TxtCr.Text & "'," & _
               "N''," & _
      " 0," & _
                "" & CDbl(txtBill_Amt.Text) & "," & _
                " 0," & _
                  "" & CDbl(txtBill_Amt.Text) * CDbl(1) & "," & _
                " 0," & _
                   " 0," & _
                 " 1," & _
                     " 1," & _
                " Getdate()," & _
              "N'" & MUserID & "'," & _
              "N'" & MuSubOff2 & "',0,'1')"
                CNN.Execute(CNCr)
            End If
            CNN.Execute("update AP_ACC_Gen_Item set  AP_ACC_Gen_Item.descrip=Acc_Code.Name_L, AP_ACC_Gen_Item.ac_name=Acc_Code.Name_L,  AP_ACC_Gen_Item.ac_typee=Acc_Code.Acc_TypeE from Acc_Code,AP_ACC_Gen_Item where AP_ACC_Gen_Item.certify='" & Trim(MDcertify) & "' and AP_ACC_Gen_Item.AC_Code=ACC_Code.AC_Code ")

            CNN.Execute("update gen_jn set  gen_jn.ac_name=Acc_Code.Name_L, gen_jn.ac_namee=Acc_Code.Name_E from Acc_Code,gen_jn where gen_jn.certify=N'" & Trim(MDcertify) & "' and gen_jn.AC_Code=ACC_Code.AC_Code ")
            'CNN.Execute("update Adjustment_List set  Remain= " & CDbl(FG.get_TextMatrix(i, 7)) & "-" & CDbl(FG.get_TextMatrix(i, 13)) & " where Code=N'" & (FG.get_TextMatrix(i, 1)) & "' ")
            'CNN.Execute("update Adjustment_List set  Remain= " & CDbl(FG.get_TextMatrix(i, 14)) & "  where Code=N'" & (FG.get_TextMatrix(i, 1)) & "' ")
            MsgBox("Finish")
        End If




    End Sub
    Private Sub OfficeNEW()
        Dim Rs As New ADODB.Recordset
        With Rs
            Call LoadSqlData("SELECT * FROM AP_Office where off_id='" & Off_Id & "' ", Rs)
            If .RecordCount = 0 Then Exit Sub
            OffName = Trim(.Fields("off_name").Value.ToString)
            OffNameE = Trim(.Fields("off_namee").Value.ToString)
            'Off_strtl = Trim(.Fields("off_strtl").Value.ToString)
            'Off_VillageL = Trim(.Fields("Off_VillageL").Value.ToString)
            'Off_DistL = Trim(.Fields("Off_DistL").Value.ToString)
            'Off_ProVL = Trim(.Fields("Off_ProVL").Value.ToString)
            OffTel = Trim(.Fields("tel").Value.ToString)
            OffFax = Trim(.Fields("fax").Value.ToString)
            Sign1 = Trim(.Fields("Sign1").Value.ToString)
            Sign2 = Trim(.Fields("Sign2").Value.ToString)
            Sign3 = Trim(.Fields("Sign3").Value.ToString)
            Sign4 = Trim(.Fields("Sign4").Value.ToString)
            Sign5 = Trim(.Fields("Sign5").Value.ToString)
            OffPlace = Trim(.Fields("Place").Value.ToString)
            .MoveNext()
        End With
    End Sub
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        LoadListFG()
        'FGCal()
        Sum()
        PPPP()

    End Sub
    Private Sub PPPP()
        Call Office()
        OfficeNEW()
        Dim Rs, Rs1 As New ADODB.Recordset
        Dim rpt As Object
  
        rpt = New Cry_Fixed_Assets_Register

        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim cRS As New ADODB.Recordset
        Dim str As String = ""
        Dim ss As String

        CNN.Execute("update Rpt_Grp set Rpt_Grp.Group_Nm=Groups_Asset.Group_Nm from Rpt_Grp,Groups_Asset where Groups_Asset.Group_ID=Rpt_Grp.Group_ID")

        ss = "Select * from Rpt_Grp  Order by AssetID"
        Call LoadSqlData(ss, Rs)
 


        If Rs.RecordCount = 0 Then
            MsgBox("Data Emtry") : Exit Sub
        End If
        With rpt
            Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("OffNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = OffName
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Tel"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = "Tel: " & OffTel & "Fax: " & OffFax
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg1"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = Sign1
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg2"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = Sign2
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg3"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = Sign3
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg4"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = Sign4
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg5"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = Sign5
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("place"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = PlaecL
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("HD"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = "Fixed Assets Register As Of  " & DateIn.Text
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("End"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = "Month ended " & DateIn.Text

            'Dim RO As CrystalDecisions.CrystalReports.Engine.ReportObject
            'Dim SRO As CrystalDecisions.CrystalReports.Engine.SubreportObject
            'Dim SubDoc As CrystalDecisions.CrystalReports.Engine.ReportDocument
            'SqlPrint = "SELECT  * from Ap_Image  where Img_Id='" & IMageID & "'"
            'Call LoadSqlData(SqlPrint, Rs1)
            'RO = rpt.ReportDefinition.Sections.Item("Section1").ReportObjects.Item("Subreport1")
            'SRO = CType(RO, CrystalDecisions.CrystalReports.Engine.SubreportObject)
            'SubDoc = SRO.OpenSubreport(SRO.SubreportName)
            'If Rs1.RecordCount > 0 Then
            '    SubDoc.SetDataSource(Rs1)
            '    FmPreview.ReportViewer.ReportSource = SubDoc
            'End If

            rpt.SetDataSource(Rs)
            rpt.Refresh()
            FmPreview.ReportViewer.ReportSource = rpt
            FmPreview.ReportViewer.DisplayGroupTree = False
            FmPreview.WindowState = FormWindowState.Maximized
            FmPreview.Show()
            'Call CloseRs(RSC)
            'Call CloseRs(Rs)
        End With
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        LoadListFG()
        'FGCal()
        Sum()
        PPPP_ALL()
    End Sub
    Private Sub PPPP_ALL()
        Call Office()
        OfficeNEW()
        Dim Rs, Rs1 As New ADODB.Recordset
        Dim rpt As Object

        rpt = New Cry_Fixed_Assets_Register_All

        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim cRS As New ADODB.Recordset
        Dim str As String = ""
        Dim ss As String

        CNN.Execute("update Rpt_Grp set Rpt_Grp.Group_Nm=Groups_Asset.Group_Nm from Rpt_Grp,Groups_Asset where Groups_Asset.Group_ID=Rpt_Grp.Group_ID")

        ss = "Select * from Rpt_Grp  Order by AssetID"
        Call LoadSqlData(ss, Rs)



        If Rs.RecordCount = 0 Then
            MsgBox("Data Emtry") : Exit Sub
        End If
        With rpt
            Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("OffNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = OffName
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Tel"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = "Tel: " & OffTel & "Fax: " & OffFax
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg1"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = Sign1
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg2"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = Sign2
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg3"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = Sign3
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg4"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = Sign4
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("Sg5"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = Sign5
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("place"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = PlaecL
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("HD"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = "Fixed Assets Register As Of  " & DateIn.Text
            myText2 = CType(rpt.ReportDefinition.ReportObjects.Item("End"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = "Month ended  " & DateIn.Text

            Dim RO As CrystalDecisions.CrystalReports.Engine.ReportObject
            Dim SRO As CrystalDecisions.CrystalReports.Engine.SubreportObject
            Dim SubDoc As CrystalDecisions.CrystalReports.Engine.ReportDocument
            SqlPrint = " Select * from Rpt_Grp  Order by AssetID "
            Call LoadSqlData(SqlPrint, Rs1)
            RO = rpt.ReportDefinition.Sections.Item("Section2").ReportObjects.Item("Subreport1")
            SRO = CType(RO, CrystalDecisions.CrystalReports.Engine.SubreportObject)
            SubDoc = SRO.OpenSubreport(SRO.SubreportName)
            If Rs1.RecordCount > 0 Then
                SubDoc.SetDataSource(Rs1)
                FmPreview.ReportViewer.ReportSource = SubDoc
            End If

            rpt.SetDataSource(Rs)
            rpt.Refresh()
            FmPreview.ReportViewer.ReportSource = rpt
            FmPreview.ReportViewer.DisplayGroupTree = False
            FmPreview.WindowState = FormWindowState.Maximized
            FmPreview.Show()
            'Call CloseRs(RSC)
            'Call CloseRs(Rs)
        End With
    End Sub

End Class