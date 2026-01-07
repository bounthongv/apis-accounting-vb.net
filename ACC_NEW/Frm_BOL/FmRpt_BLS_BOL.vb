Public Class FmRpt_BLS_BOL
    Dim r As String
    Dim CLT_Str, CLT_Last_Str As String
    Dim bls1 As String
    Dim MonthLetter1 As String
    Dim MdStartDate As Date
    Dim MdToDate As Date

    Dim MdStartDate_Last As Date
    Dim MdToDate_Last As Date
    Dim ny, ly, n_L_y As String

    Dim sql As String
    Dim AmtOpenDR, AmtOpenCR, AmtOpenMonthDR, AmtOpenMonthCR As Double
    Dim VCode1, VCode2, VCode3, VCode4, VCode5, VCode6, VCode7, VCode8, VCode9 As String
    Dim MdQuarter As Date
     'Dim RsOpen As New ADODB.Recordset
     'Dim RsOpenMonth As New ADODB.Recordset
     'Dim RsRpt As New ADODB.Recordset
     Dim VOpenDate As Date
     Dim RptNme As String
     'Dim RSC12 As New ADODB.Recordset
     'Dim RSCIn_M As New ADODB.Recordset


    Private Sub ChangBalance()
        New_Code = "3901"
        Code_Dr = "4"
        Code_Cr = "5"
        Ac_Code = ""
        'MsgBox(MdStartDate & "==" & MdToDate)

        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()


        DbHelper.ExecuteNonQuery("DELETE  Ap_balance_6_col ")
        DbHelper.ExecuteNonQuery("DELETE FROM Ap_balance_6 ")
        DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")
        Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
        DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE    date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr)as open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")
        DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        Call Chang_Incom()
        DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + amt_dr) - (open_amt_cr + amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr + amt_cr) ")
        DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr + amt_cr) - (open_amt_dr + amt_dr) where (open_amt_cr + amt_cr) >= (open_amt_dr + amt_dr) ")
        DbHelper.ExecuteNonQuery("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")

    End Sub

    Private Sub Chang_Incom()
        Insr = "delete  Ap_balance_6  " & _
           "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  )   select '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr) from Ap_balance_6_col where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "'  or  Ac_Code =  '" & New_Code & "'  " & _
  "update  Ap_balance_6 set open_amt_dr = 0 where open_amt_dr  is null  " & _
  "update  Ap_balance_6 set open_amt_cr = 0 where open_amt_cr  is null  " & _
  "update  Ap_balance_6 set amt_dr = 0 where amt_dr  is null  " & _
  "update  Ap_balance_6 set amt_cr = 0 where amt_cr  is null   " & _
   "Update  Ap_balance_6 set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr   " & _
  "Update  Ap_balance_6 set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr " & _
  "Update  Ap_balance_6 set   amt_dr = amt_dr  - amt_cr , amt_cr=0  where amt_dr  >= amt_cr   " & _
  "Update  Ap_balance_6 set   amt_cr = amt_cr  - amt_dr , amt_dr=0  where amt_cr  >= amt_dr " & _
     "delete  Ap_balance_6_col  where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "'  or  Ac_Code =  '" & New_Code & "'  " & _
       "  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  from Ap_balance_6"
        DbHelper.ExecuteNonQuery(Insr)
    End Sub


    Private Sub BtnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()

        Off_Find = Off_Usr.Text : MuTable = ""
        Off_Find2 = Off_Usr.Text : MuTable = ""
        'MsgBox(Off_Find)
        Call Find_Company()
        Call ChangBalance()

        Call ChangBalance()
        DbHelper.ExecuteNonQuery("update  Ap_Rpt_BLS_BOL_Item set Last_amt_dr  =  " & CDbl(0) & " , Last_amt_cr  = " & CDbl(0) & " ,amt_dr  =  " & CDbl(0) & " , amt_cr  = " & CDbl(0) & "  ")
        DbHelper.ExecuteNonQuery("update Ap_Rpt_BLS_BOL set Last_Amt = 0 , Amt  =  " & CDbl(0) & "  ")
        DbHelper.ExecuteNonQuery("DELETE FROM Ap_Rpt_BLS_BOL_Detail ")
        SelcectIn()
        UpdateIIn()
        SelectOut()
        UpdateOut()
        Update_Sum()
        If RD.Checked = True Then
            Dim s1, s2 As String
            LngId = 7027 : CallLngStr() : s2 = LngStr
            LngId = 7072 : CallLngStr() : s1 = LngStr
            'Lb.Text = s1 & " " & Ds.Text & " " & s2 & " 
        ElseIf RM.Checked = True Then
            Dim s1, s2, s3 As String
            LngId = 7013 + DMonth.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7070 : CallLngStr() : s1 = LngStr
            LngId = 7064 : CallLngStr() : s3 = LngStr

            Lb.Text = s1 & " " & s2 & "/" & Year(MdToDate)
        ElseIf RP.Checked = True Then
            Dim s1, s2, s3 As String
            LngId = 7074 + Period.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7062 : CallLngStr() : s1 = LngStr
            LngId = 7064 : CallLngStr() : s3 = LngStr
            Lb.Text = s1 & " " & s2 & " " & Year(MdToDate)
            'ElseIf Rhalf.Checked = True Then
            '    Dim s1, s2 As String
            '    'LngId = 7078 + M1.SelectedIndex : CallLngStr() : s2 = LngStr
            '    LngId = 7082 : CallLngStr() : s1 = LngStr
            '    LngId = 7027 : CallLngStr() : s2 = LngStr
            '    Lb.Text = s1 & " " & M1.SelectedIndex + 1 & " " & s2 & " " & M2.SelectedIndex + 1 & " (" & Year(MdToDate) & " )"



        ElseIf RY.Checked = True Then
            Dim s1, s2, s3 As String
            LngId = 7074 + Period.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7062 : CallLngStr() : s1 = LngStr
            LngId = 7064 : CallLngStr() : s3 = LngStr
            Lb.Text = s3 & " " & yy.Text
        End If
        If CheckBox1.Checked = False Then
            Call LoadReport()
        Else
            Call LoadReportItem()
        End If
        'MdStartDate = d1
        'MdToDate = d2
    End Sub
     Private Sub LoadOpen_Jn16()
         'Dim RSC16 As New ADODB.Recordset
         Dim dt As DataTable = DbHelper.GetDataTable("  select sum(rem_cr - rem_dr) as x  , count(rem_cr - rem_dr) as y  from Ap_balance_6_col  where Ac_Code Like '1%' or Ac_Code Like '2%' or Ac_Code Like '3%' ")

         If dt.Rows.Count > 0 Then
             Dim row As DataRow = dt.Rows(0)
             If CDbl(Trim(DbHelper.GetStr(row("y")))) > 0 Then

                 If CDbl(Trim(DbHelper.GetStr(row("x")))) > 0 Then
                     Rem_Cr = CDbl(Trim(DbHelper.GetStr(row("x"))))
                     Rem_Dr = 0
                 End If
                 If CDbl(Trim(DbHelper.GetStr(row("x")))) < 0 Then
                     Rem_Dr = CDbl(Trim(DbHelper.GetStr(row("x")))) * CDbl(-1)
                     Rem_Cr = 0
                 End If
                 If CDbl(Trim(DbHelper.GetStr(row("x")))) <> 0 Then
                     Call LoadOpen_Jn17()
                 End If
             End If
         End If


     End Sub
     Private Sub LoadOpen_Jn17()
         'Dim RSC17 As New ADODB.Recordset
         Dim dt As DataTable = DbHelper.GetDataTable("   select Ac_Code from Ap_balance_6_col  where Ac_Code ='65'")
         If dt.Rows.Count <> 0 Then
             DbHelper.ExecuteNonQuery(" Update Ap_balance_6_col set Amt_Dr = " & CDbl(Rem_Dr) & "  , Amt_Cr =" & CDbl(Rem_Cr) & " ")
         Else
             DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6_col ( ac_code ,ac_name , ac_namee , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
              "Values('65', N'" & "***" & "', '" & "***" & "', " & _
              " " & CDbl(0) & ", " & CDbl(0) & ", " & CDbl(Rem_Dr) & ", " & CDbl(Rem_Cr) & ",0 )")
         End If

         DbHelper.ExecuteNonQuery(" Delete Ap_balance_6_col  where Ac_Code Like '1%' or Ac_Code Like '2%' or Ac_Code Like '3%'  ")

     End Sub
    Private Sub SelcectIn()

        DbHelper.ExecuteNonQuery("Update Ap_Rpt_BLS_BOL_Item set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_BOL_Item , Ap_balance_6_col " & _
                "where Ap_Rpt_BLS_BOL_Item.ac_code = Ap_balance_6_col.ac_code And  Rpt_Type = 'In'")

        DbHelper.ExecuteNonQuery("Insert into Ap_Rpt_BLS_BOL_Detail (  Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type )" & _
         " select   Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type from Ap_Rpt_BLS_BOL_Item where Rpt_Type = 'In' And ( Amt_Dr <>0 or Amt_Cr <>0  or Last_Amt_Dr <>0 or Last_Amt_Cr <>0 )")

    End Sub


     Private Sub SelcectInLast()
         Dim dt As DataTable = DbHelper.GetDataTable("select * from Ap_Rpt_BLS_BOL_Item where  Rpt_Type = 'In'")
         For Each row As DataRow In dt.Rows
             UpdateIIn_ItemLast(row)
         Next
     End Sub

     Private Sub UpdateIIn_Item(ByVal row As DataRow)
         'Dim RSCkk As New ADODB.Recordset
         Dim dt As DataTable = DbHelper.GetDataTable(" select * from Ap_balance_6_col   where ac_code =  '" & DbHelper.GetStr(row("Ac_Code")) & "' ")
         For Each innerRow As DataRow In dt.Rows
             DbHelper.ExecuteNonQuery("Insert into Ap_Rpt_BLS_BOL_Detail (  Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type ) values ( '" & DbHelper.GetStr(row("Rpt_Id")) & "' , '" & DbHelper.GetStr(innerRow("Ac_Code")) & "' , N'" & DbHelper.GetStr(innerRow("Ac_Name")) & "'   , " & CDbl(DbHelper.GetStr(innerRow("open_amt_dr"))) & " , " & CDbl(DbHelper.GetStr(innerRow("open_amt_Cr"))) & "   , " & CDbl(DbHelper.GetStr(innerRow("Rem_dr"))) & " , " & CDbl(DbHelper.GetStr(innerRow("Rem_cr"))) & " , 'In')")
             DbHelper.ExecuteNonQuery("update  Ap_Rpt_BLS_BOL_Item set  Last_amt_dr  =  Last_amt_dr+" & CDbl(DbHelper.GetStr(innerRow("open_amt_dr"))) & " , Last_amt_cr  = Last_amt_cr+" & CDbl(DbHelper.GetStr(innerRow("open_amt_Cr"))) & " , Amt_Dr  =  Amt_Dr+" & CDbl(DbHelper.GetStr(innerRow("Rem_dr"))) & " , Amt_Cr  = Amt_Cr+" & CDbl(DbHelper.GetStr(innerRow("Rem_cr"))) & "   where ac_code = '" & DbHelper.GetStr(row("ac_code")) & "' And  Rpt_Type = 'In' ")
         Next
     End Sub


     Private Sub UpdateIIn_ItemLast(ByVal row As DataRow)
         'Dim RSCkk As New ADODB.Recordset
         Dim dt As DataTable = DbHelper.GetDataTable(" select * from Ap_balance_6_col   where ac_code =  '" & DbHelper.GetStr(row("Ac_Code")) & "' ")
         For Each innerRow As DataRow In dt.Rows
             DbHelper.ExecuteNonQuery("Insert into Ap_Rpt_BLS_BOL_Detail (Ac_Code_Parent , Rpt_Id , Ac_Code , Ac_Name  , Last_Amt_Dr , Last_Amt_Cr , Rpt_Type) values (  '" & DbHelper.GetStr(row("Ac_Code")) & "' , '" & DbHelper.GetStr(row("Rpt_Id")) & "' , '" & DbHelper.GetStr(innerRow("Ac_Code")) & "' , N'" & DbHelper.GetStr(innerRow("Ac_Name")) & "'  , " & CDbl(DbHelper.GetStr(innerRow("Rem_dr"))) & " , " & CDbl(DbHelper.GetStr(innerRow("Rem_cr"))) & ", 'In')")
             DbHelper.ExecuteNonQuery("update  Ap_Rpt_BLS_BOL_Item set Amt_Dr  =  Amt_Dr+" & CDbl(DbHelper.GetStr(innerRow("Rem_dr"))) & " , Amt_Cr  = Amt_Cr+" & CDbl(DbHelper.GetStr(innerRow("Rem_cr"))) & "   where ac_code = '" & DbHelper.GetStr(row("ac_code")) & "' And  Rpt_Type = 'In' ")
         Next
     End Sub

    Private Sub UpdateIIn()
        DbHelper.ExecuteNonQuery("delete Ap_Rpt_BLS_BOL_Stock ")
        DbHelper.ExecuteNonQuery(" insert into Ap_Rpt_BLS_BOL_Stock ( Rpt_ID , Last_Amt_Dr , Last_Amt_Cr, Amt_Dr , Amt_Cr)" & _
                     "  select Rpt_ID , sum(Last_Amt_Dr) As Last_Amt_Dr , sum(Last_Amt_Cr) As Last_Amt_Cr ,sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr   from Ap_Rpt_BLS_BOL_Item  where  Rpt_Type = 'In' group by Rpt_ID")
        DbHelper.ExecuteNonQuery("Update Ap_Rpt_BLS_BOL set Amt = Ap_Rpt_BLS_BOL_Stock.Amt_Dr-Ap_Rpt_BLS_BOL_Stock.Amt_cr ,Last_Amt =Ap_Rpt_BLS_BOL_Stock.Last_Amt_dr-Ap_Rpt_BLS_BOL_Stock.Last_Amt_Cr  from Ap_Rpt_BLS_BOL ,Ap_Rpt_BLS_BOL_Stock where  Ap_Rpt_BLS_BOL.Rpt_ID=Ap_Rpt_BLS_BOL_Stock.Rpt_ID")
    End Sub
     Private Sub UpdateIInLast()
         'Dim RSC As New ADODB.Recordset
         Dim dt As DataTable = DbHelper.GetDataTable("select Rpt_ID, sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr  from Ap_Rpt_BLS_BOL_Item  where  Rpt_Type = 'In' group by Rpt_ID ")
         For Each row As DataRow In dt.Rows
             DbHelper.ExecuteNonQuery("Update Ap_Rpt_BLS_BOL set " & _
                         " Last_Amt ='" & CDbl(CDbl(DbHelper.GetStr(row("Amt_dr"))) - CDbl(DbHelper.GetStr(row("Amt_cr")))) & "' " & _
                            " where Rpt_ID = '" & DbHelper.GetStr(row("Rpt_ID")) & "' ")
         Next
     End Sub


    Private Sub SelectOut()
        DbHelper.ExecuteNonQuery("Update Ap_Rpt_BLS_BOL_Item set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_BOL_Item , Ap_balance_6_col " & _
          "where Ap_Rpt_BLS_BOL_Item.ac_code = Ap_balance_6_col.ac_code And  Rpt_Type = 'Out'")

        DbHelper.ExecuteNonQuery("Insert into Ap_Rpt_BLS_BOL_Detail (  Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type )" & _
         " select   Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type from Ap_Rpt_BLS_BOL_Item where Rpt_Type = 'Out' And ( Amt_Dr <>0 or Amt_Cr <>0  or Last_Amt_Dr <>0 or Last_Amt_Cr <>0 )")

        'LoadSqlData("select * from Ap_Rpt_BLS_BOL_Item where  Rpt_Type = 'Out' ", RSCIn_M)
        'With RSCIn_M
        '    Do Until .EOF = True
        '        Call UpdateOut_Item()
        '        .MoveNext()
        '    Loop
        'End With
        'If RSCIn_M.State = ConnectionState.Open Then RSCIn_M.Close()
    End Sub
     Private Sub SelectOutLast()

         Dim dt As DataTable = DbHelper.GetDataTable("select * from Ap_Rpt_BLS_BOL_Item where  Rpt_Type = 'Out' ")
         For Each row As DataRow In dt.Rows
             UpdateOut_ItemLast(row)
         Next
         'If RSCIn_M.State = ConnectionState.Open Then RSCIn_M.Close()
     End Sub

     Private Sub UpdateOut_Item(ByVal row As DataRow)
         'Dim RSCkk As New ADODB.Recordset
         Dim dt As DataTable = DbHelper.GetDataTable(" select * from Ap_balance_6_col   where ac_code =  '" & DbHelper.GetStr(row("Ac_Code")) & "' ")
         For Each innerRow As DataRow In dt.Rows
             DbHelper.ExecuteNonQuery("Insert into Ap_Rpt_BLS_BOL_Detail (Ac_Code_Parent , Rpt_Id , Ac_Code , Ac_Name  ,  Last_Amt_Dr , Last_Amt_Cr, Amt_Dr , Amt_Cr , Rpt_Type ) values (  '" & DbHelper.GetStr(row("Ac_Code")) & "' , '" & DbHelper.GetStr(row("Rpt_Id")) & "' , '" & DbHelper.GetStr(innerRow("Ac_Code")) & "' , N'" & DbHelper.GetStr(innerRow("Ac_Name")) & "'  ,   " & CDbl(DbHelper.GetStr(innerRow("Open_Amt_dr"))) & " , " & CDbl(DbHelper.GetStr(innerRow("Open_Amt_cr"))) & " , " & CDbl(DbHelper.GetStr(innerRow("Rem_dr"))) & " , " & CDbl(DbHelper.GetStr(innerRow("Rem_cr"))) & " , 'Out' )")
             DbHelper.ExecuteNonQuery("update  Ap_Rpt_BLS_BOL_Item set Last_Amt_Dr  =  Last_Amt_Dr+" & CDbl(DbHelper.GetStr(innerRow("Open_Amt_dr"))) & " , Last_Amt_Cr  = Last_Amt_Cr+" & CDbl(DbHelper.GetStr(innerRow("Open_Amt_cr"))) & "  , Amt_Dr  =  Amt_Dr+" & CDbl(DbHelper.GetStr(innerRow("Rem_dr"))) & " , Amt_Cr  = Amt_Cr+" & CDbl(DbHelper.GetStr(innerRow("Rem_cr"))) & "   where ac_code = '" & DbHelper.GetStr(row("ac_code")) & "' And  Rpt_Type = 'Out' ")
         Next

     End Sub

     Private Sub UpdateOut_ItemLast(ByVal row As DataRow)
         'Dim RSCkk As New ADODB.Recordset
         Dim dt As DataTable = DbHelper.GetDataTable(" select * from Ap_balance_6_col   where ac_code =  '" & DbHelper.GetStr(row("Ac_Code")) & "' ")
         For Each innerRow As DataRow In dt.Rows
             DbHelper.ExecuteNonQuery("Insert into Ap_Rpt_BLS_BOL_Detail (Ac_Code_Parent , Rpt_Id , Ac_Code , Ac_Name  , Last_Amt_Dr , Last_Amt_Cr , Rpt_Type) values (  '" & DbHelper.GetStr(row("Ac_Code")) & "' , '" & DbHelper.GetStr(row("Rpt_Id")) & "' , '" & DbHelper.GetStr(innerRow("Ac_Code")) & "' , N'" & DbHelper.GetStr(innerRow("Ac_Name")) & "'  , " & CDbl(DbHelper.GetStr(innerRow("Rem_dr"))) & " , " & CDbl(DbHelper.GetStr(innerRow("Rem_cr"))) & ", 'Out')")
             DbHelper.ExecuteNonQuery("update  Ap_Rpt_BLS_BOL_Item set Amt_Dr  =  Amt_Dr+" & CDbl(DbHelper.GetStr(innerRow("Rem_dr"))) & " , Amt_Cr  = Amt_Cr+" & CDbl(DbHelper.GetStr(innerRow("Rem_cr"))) & "   where ac_code = '" & DbHelper.GetStr(row("ac_code")) & "' And  Rpt_Type = 'Out' ")
         Next

     End Sub



    'Private Sub UpdateOutLast()
    '    Dim RSC As New ADODB.Recordset
    '    LoadSqlData("select Rpt_ID, sum(Last_Amt_Dr) As Last_Amt_Dr , sum(Last_Amt_Cr) As Last_Amt_Cr  from Ap_Rpt_BLS_BOL_Item  where  Rpt_Type = 'Out' group by Rpt_ID ", RSC)
    '    With RSC
    '        Do Until .EOF = True
    '            DbHelper.ExecuteNonQuery("Update Ap_Rpt_BLS_BOL set " & _
    '                     " Last_Amt ='" & CDbl(CDbl((.Fields("Last_Amt_cr").Value)) - CDbl((.Fields("Last_Amt_dr").Value))) & "' " & _
    '                        " where Rpt_ID = '" & (.Fields("Rpt_ID").Value) & "' ")
    '            .MoveNext()
    '        Loop
    '    End With
    'End Sub

    Private Sub UpdateOut()
        DbHelper.ExecuteNonQuery("delete Ap_Rpt_BLS_BOL_Stock ")
        DbHelper.ExecuteNonQuery(" insert into Ap_Rpt_BLS_BOL_Stock ( Rpt_ID , Last_Amt_Dr , Last_Amt_Cr, Amt_Dr , Amt_Cr)" & _
                     "  select Rpt_ID , sum(Last_Amt_Dr) As Last_Amt_Dr , sum(Last_Amt_Cr) As Last_Amt_Cr ,sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr   from Ap_Rpt_BLS_BOL_Item  where  Rpt_Type = 'Out' group by Rpt_ID")
        DbHelper.ExecuteNonQuery("Update Ap_Rpt_BLS_BOL set Amt = Ap_Rpt_BLS_BOL_Stock.Amt_Dr-Ap_Rpt_BLS_BOL_Stock.Amt_Dr ,Last_Amt =Ap_Rpt_BLS_BOL_Stock.Last_Amt_dr-Ap_Rpt_BLS_BOL_Stock.Last_Amt_Cr  from Ap_Rpt_BLS_BOL ,Ap_Rpt_BLS_BOL_Stock where  Ap_Rpt_BLS_BOL.Rpt_ID=Ap_Rpt_BLS_BOL_Stock.Rpt_ID")
        DbHelper.ExecuteNonQuery("Update Ap_Rpt_BLS_BOL set Amt = Ap_Rpt_BLS_BOL_Stock.Amt_Cr-Ap_Rpt_BLS_BOL_Stock.Amt_Dr ,Last_Amt =Ap_Rpt_BLS_BOL_Stock.Last_Amt_Cr-Ap_Rpt_BLS_BOL_Stock.Last_Amt_Dr  from Ap_Rpt_BLS_BOL ,Ap_Rpt_BLS_BOL_Stock where  Ap_Rpt_BLS_BOL.Rpt_ID=Ap_Rpt_BLS_BOL_Stock.Rpt_ID")
        Dim RSC As New ADODB.Recordset
        'LoadSqlData("select Rpt_ID, sum(Last_Amt_dr) As Last_Amt_Dr , sum(Last_Amt_cr) As Last_Amt_cr , sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr  from Ap_Rpt_BLS_BOL_Item  where  Rpt_Type = 'Out' group by Rpt_ID ", RSC)
        'With RSC
        '    Do Until .EOF = True
        '        DbHelper.ExecuteNonQuery("Update Ap_Rpt_BLS_BOL set " & _
        '                 " Amt ='" & CDbl(CDbl((.Fields("Amt_cr").Value)) - CDbl((.Fields("Amt_dr").Value))) & "' " & _
        '                   " ,Last_Amt ='" & CDbl(CDbl((.Fields("Last_Amt_cr").Value)) - CDbl((.Fields("Last_Amt_dr").Value))) & "' " & _
        '                    " where Rpt_ID = '" & (.Fields("Rpt_ID").Value) & "' ")
        '        .MoveNext()
        '    Loop
        'End With
    End Sub


     Private Sub UpdateOutLast()
         'Dim RSC As New ADODB.Recordset
         Dim dt As DataTable = DbHelper.GetDataTable("select Rpt_ID, sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr  from Ap_Rpt_BLS_BOL_Item  where  Rpt_Type = 'Out' group by Rpt_ID ")
         For Each row As DataRow In dt.Rows
             DbHelper.ExecuteNonQuery("Update Ap_Rpt_BLS_BOL set " & _
                      " Last_Amt ='" & CDbl(CDbl(DbHelper.GetStr(row("Amt_cr"))) - CDbl(DbHelper.GetStr(row("Amt_dr")))) & "' " & _
                         " where Rpt_ID = '" & DbHelper.GetStr(row("Rpt_ID")) & "' ")
         Next
     End Sub
    Private Sub Update_Sum()
        DbHelper.ExecuteNonQuery("update Ap_Rpt_BLS_BOL_Detail set  Rpt_Name=Ap_Rpt_BLS_BOL.Description from   Ap_Rpt_BLS_BOL_Detail , Ap_Rpt_BLS_BOL  where Ap_Rpt_BLS_BOL_Detail.Rpt_Id = Ap_Rpt_BLS_BOL.Rpt_Id")
        DbHelper.ExecuteNonQuery(" Update Caculate_Rpt set  CLT_Amt  = CLT_Str ,  CLT_Last_Amt  = CLT_Str where CLT_Str = '+' or CLT_Str = '-' or CLT_Str = '*' or CLT_Str = '+' or CLT_Str = '/' or CLT_Str = '(' or CLT_Str=')' ")
        DbHelper.ExecuteNonQuery("delete Caculate_Lock")
        DbHelper.ExecuteNonQuery("delete Caculate_Start")
        DbHelper.ExecuteNonQuery(" Insert Into Caculate_Start (Rpt_Id , Rpt_Type , clt_Str , clt_Amt , Clt_Last_Amt ) select Rpt_Id , Rpt_Type , clt_Str , clt_Amt , Clt_Last_Amt from Caculate_Rpt where Rpt_Type = 'BLS_BOL'  Order by  Rpt_id ,cnt asc  ")
        DbHelper.ExecuteNonQuery("update Caculate_Start set lck =0")
        DbHelper.ExecuteNonQuery("Insert into Caculate_Lock (cnt_Mt)  SELECT  (SELECT     TOP 1 cnt FROM Caculate_Start AS B WHERE(Rpt_Id = A.Rpt_Id   ) ORDER BY cnt desc) AS cnt FROM Caculate_Start  AS A  GROUP BY Rpt_Id ORDER BY Rpt_Id")
        DbHelper.ExecuteNonQuery("update  Caculate_Start set lck=1 from Caculate_Start ,Caculate_Lock  where Caculate_Start.cnt=Caculate_Lock.cnt_MT")
        DbHelper.ExecuteNonQuery("  Update Caculate_Start set Caculate_Start.Amt = Ap_Rpt_BLS_BOL.Amt , Caculate_Start.Last_Amt = Ap_Rpt_BLS_BOL.Last_Amt   from Caculate_Start , Ap_Rpt_BLS_BOL  where  Caculate_Start.CLT_Str  = Ap_Rpt_BLS_BOL.Rpt_Id  ")
        DbHelper.ExecuteNonQuery("Update Caculate_Start set lck_Amt=0")
        DbHelper.ExecuteNonQuery("Update Caculate_Start set lck_Amt=1 where CLT_Str <> '+' And CLT_Str <> '-' And CLT_Str <> '*' And CLT_Str <> '+' And CLT_Str <> '/' And CLT_Str <> '(' And CLT_Str<>')'")
         'Dim RSC1 As New ADODB.Recordset
         CLT_Str = ""
         CLT_Last_Str = ""
         Dim dt As DataTable = DbHelper.GetDataTable("select *  from Caculate_Start where Rpt_Type = 'BLS_BOL'  Order by  Rpt_id ,cnt asc")
         If dt.Rows.Count > 0 Then
             For Each row As DataRow In dt.Rows
                 If DbHelper.GetStr(row("lck_Amt")) = "1" Then
                     CLT_Str = CLT_Str & CDbl(DbHelper.GetStr(row("Amt")))
                     CLT_Last_Str = CLT_Last_Str & CDbl(DbHelper.GetStr(row("Last_Amt")))
                 Else
                     CLT_Str = CLT_Str & Trim(DbHelper.GetStr(row("CLT_Amt")))
                     CLT_Last_Str = CLT_Last_Str & Trim(DbHelper.GetStr(row("CLT_Last_Amt")))
                 End If
                 If DbHelper.GetStr(row("lck")) = "1" Then
                     'Dim s As String = " Update  Ap_Rpt_BLS_BOL set Amt = " & (CLT_Str) & " , Last_Amt = " & (CLT_Last_Str) & " where  Rpt_ID =   '" & DbHelper.GetStr(row("Rpt_ID")) & "'"
                     'DbHelper.ExecuteNonQuery(s)
                     On Error GoTo hang
 hang:
                     If Err.Number = 0 Then
                         Dim s As String = " Update  Ap_Rpt_BLS_BOL set Amt = " & CLT_Str & " , Last_Amt = " & CLT_Last_Str & " where  Rpt_ID =   '" & DbHelper.GetStr(row("Rpt_ID")) & "'"
                         DbHelper.ExecuteNonQuery(s)
                     Else
                         MessageBox.Show("ສູດຄິດໄລ່ຂອງ " & DbHelper.GetStr(row("Rpt_ID")) & " = " & CLT_Last_Str & " ບໍ່ຖຶກຕ້ອງກະລຸນນາກວດສອບຄືນໃຫມ່")
                         Exit Sub
                     End If
                     CLT_Str = ""
                     CLT_Last_Str = ""
                 End If
             Next
         End If


        '        DbHelper.ExecuteNonQuery("update Ap_Rpt_BLS_BOL_Detail set  Rpt_Name=Ap_Rpt_BLS_BOL.Description from   Ap_Rpt_BLS_BOL_Detail , Ap_Rpt_BLS_BOL  where Ap_Rpt_BLS_BOL_Detail.Rpt_Id = Ap_Rpt_BLS_BOL.Rpt_Id")
        '        DbHelper.ExecuteNonQuery(" Update Caculate_Rpt set  CLT_Amt  = CLT_Str ,  CLT_Last_Amt  = CLT_Str where CLT_Str = '+' or CLT_Str = '-' or CLT_Str = '*' or CLT_Str = '+' or CLT_Str = '/' or CLT_Str = '(' or CLT_Str=')' Or CLT_Str<>'Cast(('   Or CLT_Str<>')As Float)'")
        '        DbHelper.ExecuteNonQuery("delete Caculate_Lock")
        '        DbHelper.ExecuteNonQuery("delete Caculate_Start")
        '        DbHelper.ExecuteNonQuery(" Insert Into Caculate_Start (Rpt_Id , Rpt_Type , clt_Str , clt_Amt , Clt_Last_Amt ) " & _
        '                    " select Rpt_Id , Rpt_Type , clt_Str , clt_Amt , Clt_Last_Amt from Caculate_Rpt where Rpt_Type = 'BLS_BOL'  Order by  Rpt_id ,cnt asc  ")
        '        DbHelper.ExecuteNonQuery("update Caculate_Start set lck =0")
        '        DbHelper.ExecuteNonQuery("Insert into Caculate_Lock (cnt_Mt)  SELECT  (SELECT     TOP 1 cnt FROM Caculate_Start AS B WHERE(Rpt_Id = A.Rpt_Id   ) ORDER BY cnt desc) AS cnt FROM Caculate_Start  AS A  GROUP BY Rpt_Id ORDER BY Rpt_Id")
        '        DbHelper.ExecuteNonQuery("update  Caculate_Start set lck=1 from Caculate_Start ,Caculate_Lock  where Caculate_Start.cnt=Caculate_Lock.cnt_MT")
        '        DbHelper.ExecuteNonQuery("  Update Caculate_Start set Caculate_Start.Amt = Ap_Rpt_BLS_BOL.Amt , Caculate_Start.Last_Amt = Ap_Rpt_BLS_BOL.Last_Amt   from Caculate_Start , Ap_Rpt_BLS_BOL  where  Caculate_Start.CLT_Str  = Ap_Rpt_BLS_BOL.Rpt_Id  ")
        '        DbHelper.ExecuteNonQuery("Update Caculate_Start set lck_Amt=0")
        '        DbHelper.ExecuteNonQuery("Update Caculate_Start set lck_Amt=1 " & _
        '                    " where CLT_Str <> '+' And CLT_Str <> '-' And CLT_Str <> '*' And CLT_Str <> '+' And CLT_Str <> '/' And " & _
        '                    " CLT_Str <> '(' And CLT_Str<>')' And CLT_Str<>'Cast(('   And CLT_Str<>')As Float)' ")
        '        Dim RSC1 As New ADODB.Recordset
        '        CLT_Str = ""
        '        CLT_Last_Str = ""
        '        With RSC1
        '            Call LoadSqlData("select *  from Caculate_Start where Rpt_Type = 'BLS_BOL'  Order by  Rpt_id ,cnt asc", RSC1)
        '            If .RecordCount > 0 Then
        '                While Not .EOF()
        '                    Dim Rpt_id As String
        '                    Rpt_id = (RSC1.Fields("Rpt_id").Value.ToString)
        '                    If (RSC1.Fields("lck_Amt").Value.ToString) = "1" Then
        '                        CLT_Str = CLT_Str & (RSC1.Fields("Amt").Value.ToString)
        '                        CLT_Last_Str = CLT_Last_Str & (RSC1.Fields("Last_Amt").Value.ToString)
        '                    Else
        '                        CLT_Str = CLT_Str & (RSC1.Fields("CLT_Amt").Value.ToString)
        '                        CLT_Last_Str = CLT_Last_Str & (RSC1.Fields("CLT_Last_Amt").Value.ToString)
        '                    End If

        '                    If (RSC1.Fields("lck").Value.ToString) = "1" Then
        '                        On Error GoTo hang
        'hang:
        '                        If Err.Number = 0 Then
        '                            Dim s As String = " Update  Ap_Rpt_BLS_BOL set Amt = " & CLT_Str & " , Last_Amt = " & CLT_Last_Str & " where  Rpt_ID =   '" & (RSC1.Fields("Rpt_ID").Value.ToString) & "'"
        '                            DbHelper.ExecuteNonQuery(s)
        '                        Else
        '                            Dim cnt As String
        '                            Dim lck As String
        '                            cnt = (RSC1.Fields("cnt").Value.ToString)
        '                            lck = (RSC1.Fields("lck").Value.ToString)
        '                            MessageBox.Show("ສູດຄິດໄລ່ຂອງ " & (RSC1.Fields("Rpt_ID").Value.ToString) & " = " & CLT_Last_Str & " ບໍ່ຖຶກຕ້ອງກະລຸນນາກວດສອບຄືນໃຫມ່")
        '                            Exit Sub
        '                        End If
        '                        CLT_Str = ""
        '                        CLT_Last_Str = ""
        '                    End If
        '                    .MoveNext()
        '                End While
        '            End If
        '        End With

    End Sub



     Private Sub LoadOpen_Jn1()
         'Dim RSC12 As New ADODB.Recordset


         Dim add As String = "select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & MULook & " group BY ac_code "

         'MsgBox(add)

         Dim dt As DataTable = DbHelper.GetDataTable(add)
         For Each row As DataRow In dt.Rows
             VCode1 = CStr(Trim(DbHelper.GetStr(row("ac_Code"))))
             DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6_col( ac_code   , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
              "Values('" & CStr(Trim(DbHelper.GetStr(row("ac_Code")))) & "', " & _
              " " & CDbl(0) & ", " & CDbl(0) & ", " & CDbl(Trim(DbHelper.GetStr(row("amt_dr")))) & ", " & CDbl(Trim(DbHelper.GetStr(row("amt_cr")))) & ",0 )")
         Next
     End Sub









    Private Sub LoadOpen_Jn12()
        DbHelper.ExecuteNonQuery("Update Ap_balance_6_col set Quarter_dr=0,Quarter_cr=0")

        Dim DS, DT As Date
        If Format(MdStartDate, "dd/MM") = "01/01" Then
            Exit Sub
        End If
        If Format(MdStartDate, "MM") = "01" Or Format(MdStartDate, "MM") = "02" Or Format(MdStartDate, "MM") = "03" Then
            DS = "01/01/" & Format(MdStartDate, "yyyy")
            DT = DateAdd(DateInterval.Day, -1, MdStartDate)
        End If

        If Format(MdStartDate, "MM") = "04" Or Format(MdStartDate, "MM") = "05" Or Format(MdStartDate, "MM") = "06" Then
            DS = "01/04/" & Format(MdStartDate, "yyyy")
            DT = DateAdd(DateInterval.Day, -1, MdStartDate)
        End If

        If Format(MdStartDate, "MM") = "07" Or Format(MdStartDate, "MM") = "08" Or Format(MdStartDate, "MM") = "09" Then
            DS = "01/07/" & Format(MdStartDate, "yyyy")
            DT = DateAdd(DateInterval.Day, -1, MdStartDate)
        End If

        If Format(MdStartDate, "MM") = "10" Or Format(MdStartDate, "MM") = "11" Or Format(MdStartDate, "MM") = "12" Then
            DS = "01/10/" & Format(MdStartDate, "yyyy")
            DT = DateAdd(DateInterval.Day, -1, MdStartDate)
        End If

         Dim dt As DataTable = DbHelper.GetDataTable("   select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(DS, "yyyy-MM-dd") & "' AND '" & Format(DT, "yyyy-MM-dd") & "' " & MULook & " group BY ac_code ")


         For Each row As DataRow In dt.Rows
             LoadOpen_Jn14(row)
         Next

    End Sub

     Private Sub LoadOpen_Jn14(ByVal row As DataRow)
         'Dim RSC14 As New ADODB.Recordset
         Dim dt As DataTable = DbHelper.GetDataTable("   select * from Ap_balance_6  where ac_code = '" & CStr(Trim(DbHelper.GetStr(row("ac_code")))) & "'  ")

         If dt.Rows.Count <> 0 Then
             DbHelper.ExecuteNonQuery("Update Ap_balance_6_col set Quarter_dr= " & CDbl(Trim(DbHelper.GetStr(row("amt_dr")))) & "  , Quarter_cr= " & CDbl(Trim(DbHelper.GetStr(row("amt_cr")))) & "  where ac_code = '" & CStr(Trim(DbHelper.GetStr(row("ac_code")))) & "'  ")
         Else
             MsgBox("ggg")
             DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6_col( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status, Quarter_dr , Quarter_cr  ) " & _
               "Values('" & CStr(Trim(DbHelper.GetStr(row("ac_Code")))) & "', " & _
               " 0 , 0 , 0 , 0,0, " & CDbl(Trim(DbHelper.GetStr(row("amt_dr")))) & " , " & CDbl(Trim(DbHelper.GetStr(row("amt_cr")))) & " )")

         End If
     End Sub
     Private Sub LoadOpen_Jn14_1()
         'Dim RSC14_1 As New ADODB.Recordset
         Dim dt As DataTable = DbHelper.GetDataTable("select * from Ap_balance_6_col  ")
         For Each row As DataRow In dt.Rows
             If CDbl(CDbl(DbHelper.GetStr(row("open_amt_dr"))) + CDbl(DbHelper.GetStr(row("amt_dr")))) >= CDbl(CDbl(DbHelper.GetStr(row("open_amt_cr"))) + CDbl(DbHelper.GetStr(row("amt_cr")))) Then
                 DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr=" & CDbl(CDbl(CDbl(DbHelper.GetStr(row("open_amt_dr"))) + CDbl(DbHelper.GetStr(row("amt_dr")))) - CDbl(CDbl(DbHelper.GetStr(row("open_amt_cr"))) + CDbl(DbHelper.GetStr(row("amt_cr"))))) & " where Ac_Code = '" & DbHelper.GetStr(row("Ac_Code")) & "'")
             Else
                 DbHelper.ExecuteNonQuery("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr=" & CDbl(CDbl(CDbl(DbHelper.GetStr(row("open_amt_cr"))) + CDbl(DbHelper.GetStr(row("amt_cr")))) - CDbl(CDbl(DbHelper.GetStr(row("open_amt_dr"))) + CDbl(DbHelper.GetStr(row("amt_dr"))))) & " where Ac_Code = '" & DbHelper.GetStr(row("Ac_Code")) & "'")
             End If
         Next
     End Sub
    Private Sub LoadOpen_Jn2()
        Dim RSC12 As New ADODB.Recordset

        Dim S As Date
        S = MdStartDate
        S = DateAdd("d", CDbl(-1), MdStartDate)
        'LoadSqlData("SELECT GIN.ac_code, ACC.Name_L, ACC.Name_E, SUM(GIN.amount_dr)AS amount_dr , SUM(GIN.amount_cr)AS amount_cr , SUM(GIN.amt_dr) AS amt_dr, SUM(GIN.amt_cr) AS amt_cr FROM Acc_Code ACC INNER JOIN gen_jn GIN ON ACC.AC_CODE = GIN.ac_code WHERE  GIN.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' GROUP BY GIN.ac_code, ACC.Name_L, ACC.Name_E  Order by GIN.AC_Code DESC  ", RSC12)

         Dim dt As DataTable = DbHelper.GetDataTable("   select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "'  " & MULook & " group BY ac_code ")


         For Each row As DataRow In dt.Rows
             DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
            "Values('" & CStr(Trim(DbHelper.GetStr(row("ac_Code")))) & "', " & _
            " " & CDbl(0) & ", " & CDbl(0) & ", " & CDbl(Trim(DbHelper.GetStr(row("amt_dr")))) & ", " & CDbl(Trim(DbHelper.GetStr(row("amt_cr")))) & ",0 )")
         Next
    End Sub

     Private Sub LoadOpen_Jn3()
         'Dim RSC3 As New ADODB.Recordset

         Dim dt As DataTable = DbHelper.GetDataTable("select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr from Open_jn WHERE    date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & MULook & " group BY ac_code")


         For Each row As DataRow In dt.Rows
             VCode3 = DbHelper.GetStr(row("ac_Code"))
             DbHelper.ExecuteNonQuery("Update Ap_balance_6 set  open_amt_dr='" & CDbl(DbHelper.GetStr(row("amt_dr"))) & "' , open_amt_cr='" & CDbl(DbHelper.GetStr(row("amt_cr"))) & "' where ac_code = '" & DbHelper.GetStr(row("ac_Code")) & "'")
             LoadOpen_Jn4()
         Next
     End Sub


     Private Sub LoadOpen_Jn4()
         'Dim RSC4 As New ADODB.Recordset
         Dim dt As DataTable = DbHelper.GetDataTable("select * from Ap_balance_6  WHERE     ac_code='" & VCode3 & "'  ")
         If dt.Rows.Count > 0 Then
             Dim row As DataRow = dt.Rows(0)
             VCode4 = DbHelper.GetStr(row("ac_Code"))
         Else
             LoadOpen_Jn5()
         End If
     End Sub

     Private Sub LoadOpen_Jn5()
         'Dim RSC5 As New ADODB.Recordset
         Dim dt As DataTable = DbHelper.GetDataTable("select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr from Open_jn  WHERE    ac_code='" & VCode3 & "' " & MULook & " group BY ac_code")
         If dt.Rows.Count > 0 Then
             Dim row As DataRow = dt.Rows(0)
             DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
          "Values('" & CStr(Trim(DbHelper.GetStr(row("ac_Code")))) & "',  " & _
          " " & CDbl(DbHelper.GetStr(row("amt_dr"))) & ", " & CDbl(DbHelper.GetStr(row("amt_cr"))) & ", " & CDbl(0) & ", " & CDbl(0) & ",0 )")
         End If
     End Sub



     Private Sub LoadOpen_Jn6()
         'Dim RSC6 As New ADODB.Recordset
         Dim op_dr, op_cr, amt_dr, amt_cr As Double
         op_dr = 0
         op_cr = 0
         amt_dr = 0
         amt_cr = 0
         Dim dt As DataTable = DbHelper.GetDataTable("select Ac_Code , open_amt_dr , open_amt_cr , Amt_dr , Amt_cr from Ap_balance_6  ")
         For Each row As DataRow In dt.Rows
             op_dr = CDbl(DbHelper.GetStr(row("open_amt_dr")))
             op_cr = CDbl(DbHelper.GetStr(row("open_amt_cr")))
             amt_dr = CDbl(DbHelper.GetStr(row("Amt_dr")))
             amt_cr = CDbl(DbHelper.GetStr(row("Amt_cr")))

             If CDbl(op_dr + amt_dr) - CDbl(op_cr + amt_cr) >= 0 Then

                 DbHelper.ExecuteNonQuery("Update Ap_balance_6 set rem_dr='" & CDbl(op_dr + amt_dr) - CDbl(op_cr + amt_cr) & "' , rem_cr='" & CDbl(0) & "' where Ac_code='" & DbHelper.GetStr(row("Ac_Code")) & "'")
             End If
             If CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) >= 0 Then
                 DbHelper.ExecuteNonQuery("Update Ap_balance_6 set rem_dr='" & CDbl(0) & "' , rem_cr='" & CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) & "' where Ac_code='" & DbHelper.GetStr(row("Ac_Code")) & "'")
             End If
         Next
     End Sub

     Private Sub LoadOpen_Jn7()
         'Dim RSC7 As New ADODB.Recordset
         Dim dt As DataTable = DbHelper.GetDataTable("select ac_code , rem_dr  , rem_cr from Ap_balance_6   ")
         For Each row As DataRow In dt.Rows
             VCode7 = DbHelper.GetStr(row("ac_Code"))

             DbHelper.ExecuteNonQuery("Update Ap_balance_6_col set  open_amt_dr='" & CDbl(DbHelper.GetStr(row("rem_dr"))) & "' , open_amt_cr='" & CDbl(DbHelper.GetStr(row("rem_cr"))) & "' where ac_code = '" & DbHelper.GetStr(row("ac_Code")) & "'")
             LoadOpen_Jn8()
         Next
     End Sub


     Private Sub LoadOpen_Jn8()
         'Dim RSC8 As New ADODB.Recordset
         Dim dt As DataTable = DbHelper.GetDataTable("select * from Ap_balance_6_col  WHERE     ac_code='" & VCode7 & "' ")
         If dt.Rows.Count > 0 Then
             Dim row As DataRow = dt.Rows(0)
             VCode8 = DbHelper.GetStr(row("ac_Code"))
         Else
             LoadOpen_Jn9()
         End If
     End Sub


     Private Sub LoadOpen_Jn9()
         'Dim RSC9 As New ADODB.Recordset
         Dim dt As DataTable = DbHelper.GetDataTable("select ac_code , Rem_dr , Rem_cr from Ap_balance_6  WHERE    ac_code='" & VCode7 & "' ")
         If dt.Rows.Count > 0 Then
             Dim row As DataRow = dt.Rows(0)
             DbHelper.ExecuteNonQuery("INSERT INTO Ap_balance_6_col ( ac_code ,ac_name , ac_namee , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
             "Values('" & CStr(Trim(DbHelper.GetStr(row("ac_Code")))) & "', N'" & "***" & "', '" & "***" & "', " & _
             " " & CDbl(DbHelper.GetStr(row("rem_dr"))) & ", " & CDbl(DbHelper.GetStr(row("rem_cr"))) & ", " & CDbl(0) & ", " & CDbl(0) & ",0 )")
         End If
     End Sub
    Private Sub LoadOpen_Jn15()
        DbHelper.ExecuteNonQuery("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")

        'Dim RSC10 As New ADODB.Recordset
        'AmtOpenDR = 0
        'AmtOpenCR = 0
        'LoadSqlData("select Ac_Code , Name_L , Name_E , Acc_TypeE from Acc_Code  ", RSC10)
        'With RSC10
        '    Do Until .EOF = True
        '        DbHelper.ExecuteNonQuery("Update Ap_balance_6_col set ac_name = N'" & (.Fields("Name_L").Value) & "'   where ac_code='" & (.Fields("ac_code").Value) & "'")
        '        .MoveNext()
        '    Loop
        'End With
    End Sub

     Private Sub LoadOpen_Jn11()
         'Dim RSC11 As New ADODB.Recordset
         Dim op_dr11, op_cr11, amt_dr11, amt_cr11 As Double
         op_dr11 = 0
         op_cr11 = 0
         amt_dr11 = 0
         amt_cr11 = 0
         Dim dt As DataTable = DbHelper.GetDataTable("select Ac_Code , open_amt_dr , open_amt_cr , Amt_dr , Amt_cr from Ap_balance_6_col  ")
         For Each row As DataRow In dt.Rows
             op_dr11 = CDbl(DbHelper.GetStr(row("open_amt_dr")))
             op_cr11 = CDbl(DbHelper.GetStr(row("open_amt_cr")))
             amt_dr11 = CDbl(DbHelper.GetStr(row("Amt_dr")))
             amt_cr11 = CDbl(DbHelper.GetStr(row("Amt_cr")))
             If CDbl(op_dr11 + op_cr11) = 0 Then
                 If CDbl(amt_dr11 + amt_cr11) = 0 Then
                     DbHelper.ExecuteNonQuery("delete Ap_balance_6_col  where Ac_code='" & DbHelper.GetStr(row("Ac_Code")) & "'")
                 End If

             End If
         Next
     End Sub

    Private Sub selectLoad()
        DMonth.Enabled = False
        Myy.Enabled = False
        Period.Enabled = False
        Pyy.Enabled = False
        Ds.Enabled = False
        Dt.Enabled = False
        yy.Enabled = False
        M2.Enabled = False
        M1.Enabled = False
        Hyy.Enabled = False
        Ct.Enabled = False
        yyt.Enabled = False
        If RD.Checked = True Then
            Ds.Enabled = True
            Dt.Enabled = True
            LoadDay()
        ElseIf RM.Checked = True Then
            DMonth.Enabled = True
            Myy.Enabled = True
            LoadMonth()
        ElseIf RP.Checked = True Then
            Period.Enabled = True
            Pyy.Enabled = True
            LoadPeriod()
        ElseIf RD.Checked = True Then
            Ds.Enabled = True
            Dt.Enabled = True
            LoadDay()
        ElseIf Rhalf.Checked = True Then
            M1.Enabled = True
            M2.Enabled = True
            Hyy.Enabled = True
            LoadHalfYear()
        ElseIf RT.Checked = True Then
            Ct.Enabled = True
            yyt.Enabled = True
            Call LoadMt()
        ElseIf RY.Checked = True Then
            yy.Enabled = True
            LoadYear()
        End If
    End Sub
    Private Sub LoadMt()
        If Ct.SelectedIndex = 0 Then
            MdStartDate = Format(CDate("1/1/" & Year(yyt.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/6/" & Year(yyt.Value)), "dd-MM-yyyy")
        Else
            MdStartDate = Format(CDate("1/7/" & Year(yyt.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/12/" & Year(yyt.Value)), "dd-MM-yyyy")
        End If

        Lb.Text = Ct.Text & " " & yyt.Text
        'L5.Text = MdStartDate & " => " & MdToDate

        L5.Text = MdStartDate & " => " & MdToDate


        'L5.Text = MdStartDate & " => " & MdToDate
    End Sub
    Private Sub LoadDay()

        Ds.Enabled = True
        Dt.Enabled = True
        MdStartDate = Ds.Value
        MdToDate = Dt.Value
        Dim s1, s2 As String
        Dim s As Integer = 1
        If Ds.Value = Dt.Value Then
            LngId = 2069 : CallLngStr()
            Lb.Text = LngStr & " " & Ds.Text
        Else
            LngId = 2070 : CallLngStr() : s1 = LngStr
            LngId = 2054 : CallLngStr() : s2 = LngStr
            Lb.Text = s1 & " " & Format(MdStartDate, "dd/MM/yyyy") & " " & s2 & " " & Format(MdToDate, "dd/MM/yyyy")
        End If
        'MdStartDate = Format(CDate(Ds.Value), "dd-MM-yyyy")
        'MdToDate = Format(CDate(Dt.Value), "dd-MM-yyyy")
        'Lb.Text = "ແຕ່ວັນທີ " & MdStartDate & " ຫາວັນທີ " & MdToDate
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub
    Private Sub LoadMonth()
        '---------------------------------
        If DMonth.Text = "01" Then
            MdStartDate = Format(CDate("01/01/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/01/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ມັງກອນ"
        ElseIf DMonth.Text = "02" Then
            Dim Day As String
            Dim MM As Date
            Dim Fromm As Date
            MdStartDate = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
            Fromm = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
            MM = Format(CDate("01/3/" & Year(MdStartDate)), "dd-MM-yyyy")
            Day = DateDiff(DateInterval.Day, Fromm, MM)
            MdToDate = Format(CDate(Day & "/02" & "/" & Year(MdStartDate)), "dd-MM-yyyy")
            MonthLetter1 = "ກຸມພາ"
            Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
        ElseIf DMonth.Text = "03" Then
            MdStartDate = Format(CDate("01/03/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/03/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ມີນາ"
        ElseIf DMonth.Text = "04" Then
            MdStartDate = Format(CDate("01/04/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/04/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ເມສາ"
        ElseIf DMonth.Text = "05" Then
            MdStartDate = Format(CDate("01/05/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/05/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ພຶດສະພາ"
        ElseIf DMonth.Text = "06" Then
            MdStartDate = Format(CDate("01/06/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/06/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ມີຖຸນາ"
        ElseIf DMonth.Text = "07" Then
            MdStartDate = Format(CDate("01/07/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/07/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ກໍລະກົດ"
        ElseIf DMonth.Text = "08" Then
            MdStartDate = Format(CDate("01/08/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/08/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ສິງຫາ"
        ElseIf DMonth.Text = "09" Then
            MdStartDate = Format(CDate("01/09/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/09/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ກັນຍາ"
        ElseIf DMonth.Text = "10" Then
            MdStartDate = Format(CDate("01/10/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/10/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ຕຸລາ"
        ElseIf DMonth.Text = "11" Then
            MdStartDate = Format(CDate("01/11/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/11/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ພະຈິກ"
        ElseIf DMonth.Text = "12" Then
            MdStartDate = Format(CDate("01/12/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/12/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ທັນວາ"
        End If
        '-----------------
        Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub

    Private Sub LoadPeriod()
        If Period.Text = "ງວດທີ 1" Then
            MdStartDate = Format(CDate("01/01/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/03/" & Year(Pyy.Value)), "dd-MM-yyyy")
            Lb.Text = "ປະຈຳງວດ " & "1" & " ປີ " & Pyy.Text
        ElseIf Period.Text = "ງວດທີ 2" Then
            MdStartDate = Format(CDate("01/04/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/06/" & Year(Pyy.Value)), "dd-MM-yyyy")
            Lb.Text = "ປະຈຳງວດ " & "2" & " ປີ " & Pyy.Text
        ElseIf Period.Text = "ງວດທີ 3" Then
            MdStartDate = Format(CDate("01/07/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/09/" & Year(Pyy.Value)), "dd-MM-yyyy")
            Lb.Text = "ປະຈຳງວດ " & "3" & " ປີ " & Pyy.Text
        ElseIf Period.Text = "ງວດທີ 4" Then
            MdStartDate = Format(CDate("01/10/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/12/" & Year(Pyy.Value)), "dd-MM-yyyy")
            Lb.Text = "ປະຈຳງວດ " & "4" & " ປີ " & Pyy.Text
        End If
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub

    Private Sub LoadYear()
        MdStartDate = Format(CDate("01/1/" & Year(yy.Value)), "dd-MM-yyyy")
        MdToDate = Format(CDate("31/12/" & Year(yy.Value)), "dd-MM-yyyy")
        Lb.Text = "ປະຈຳປີ " & yy.Text
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub

     Private Sub UpdateData2()
         'Dim RSC2 As New ADODB.Recordset
         Dim op_dr, op_cr, amt_dr, amt_cr As Double
         op_dr = 0
         op_cr = 0
         amt_dr = 0
         amt_cr = 0
         Dim dt As DataTable = DbHelper.GetDataTable("select ac_code , sum(open_amt_dr) as open_amt_dr , sum(open_amt_cr) as open_amt_cr  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  from Ap_balance_6_col where ac_code ='" & bls1 & "'   group by  ac_code ")
         For Each row As DataRow In dt.Rows
             op_dr = CDbl(DbHelper.GetStr(row("open_amt_dr")))
             op_cr = CDbl(DbHelper.GetStr(row("open_amt_cr")))
             amt_dr = CDbl(DbHelper.GetStr(row("Amt_dr")))
             amt_cr = CDbl(DbHelper.GetStr(row("Amt_cr")))
             If CDbl(op_dr + amt_dr) - CDbl(op_cr + amt_cr) >= 0 Then
                 DbHelper.ExecuteNonQuery("Update Ap_Rpt_BLS_BOL_Item set amt_dr='" & CDbl(op_dr + amt_dr) - CDbl(op_cr + amt_cr) & "' , amt_cr='" & CDbl(0) & "' where Ac_code='" & DbHelper.GetStr(row("Ac_Code")) & "'")
             End If
             If CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) >= 0 Then
                 DbHelper.ExecuteNonQuery("Update Ap_Rpt_BLS_BOL_Item set amt_dr='" & CDbl(0) & "' , amt_cr='" & CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) & "' where Ac_code='" & DbHelper.GetStr(row("Ac_Code")) & "'")
             End If
         Next
     End Sub


    Private Sub LoadReport()
        Dim RPT_ID As String


        RPT_ID = " "
        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As    RptSjUd ,"
        LngId = "7059" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_PP ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
        LngId = "7002" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDP ,"
        LngId = "7004" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Descrip ,"
        LngId = "7005" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_No ,"
        LngId = "7006" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Ac_Code ,"
        LngId = "7008" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Dr ,"
        LngId = "7009" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Cr ,"
        LngId = "7010" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Amt_LAK ,"
        LngId = "7011" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Sign ,"
        LngId = "7012" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_SignNme ,"
        LngId = "7026" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InDate ,"
        LngId = "7088" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_AmtCurr ,"
        LngId = "7031" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_User ,"
        LngId = "7042" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_OpenAmt ,"
        LngId = "7039" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_TotalAmt ,"
        LngId = "7040" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Balance ,"
        LngId = "7043" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RemAmt ,"


        If RD.Checked = True Then
            LngId = "7090" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7091" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        ElseIf RM.Checked = True Then
            LngId = "7060" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7061" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        ElseIf RP.Checked = True Then
            LngId = "7083" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7063" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        ElseIf Rhalf.Checked = True Then

            LngId = "7060" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7061" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"

        ElseIf RT.Checked = True Then
            LngId = "7080" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7081" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"

        ElseIf RY.Checked = True Then
            LngId = "7084" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7065" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"

        End If





        If r1.Checked = True Then
            LngId = "7052" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            r = "And GRP>0"
        ElseIf r2.Checked = True Then
            LngId = "7053" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            r = "And GRP<5"
        ElseIf r3.Checked = True Then
            LngId = "7054" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            r = "And GRP>5"
        End If
        SLF = "SELECT  " & MuLngRpt & "  *   FROM Ap_Rpt_BLS_BOL  "

        Call LoadLoGO()

         Dim dt As DataTable
         If r2.Checked = True Then
             dt = DbHelper.GetDataTable(" " & SLF & " where  BL_Type='1'  " & RPT_ID & "  order by Rpt_ID asc  ")
         ElseIf r3.Checked = True Then
             dt = DbHelper.GetDataTable(" " & SLF & " where BL_Type='2' " & RPT_ID & "  order by Rpt_ID asc  ")
         Else
             dt = DbHelper.GetDataTable(" " & SLF & " where grp<>'' " & RPT_ID & " " & r & "order by Rpt_ID asc  ")
         End If

         If dt.Rows.Count = 0 Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
        Dim FrmPreview As New FmPreview : FrmClosing()

         Dim Rpt As New CryRpt_BLS_BOL
         If MdShowLOGO = 1 Then
             Rpt.Subreports(0).SetDataSource(RsLOGO)
         End If
         Rpt.SetDataSource(dt)
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.DisplayGroupTree = False
        FrmPreview.WindowState = FormWindowState.Maximized
        FrmPreview.Show()
        FrmPreview.Focus()
    End Sub
     Private Sub loadOffice_User()
         Off_Usr.Items.Clear()
         Dim dt As DataTable = DbHelper.GetDataTable("select sub_id , off_add2  from  Ap_office  Order by sub_id")
         For Each row As DataRow In dt.Rows
             Off_Usr.Items.Add(DbHelper.GetStr(row("sub_id")) & " " & DbHelper.GetStr(row("off_add2")))
         Next
         Off_Usr.Text = FmLogin.Sub_Company.Text
     End Sub


    Private Sub LoadReportItem()


        Dim RPT_ID As String
        RPT_ID = " "
        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As    RptSjUd ,"
        LngId = "7059" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_PP ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
        LngId = "7002" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDP ,"
        LngId = "7004" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Descrip ,"
        LngId = "7005" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_No ,"
        LngId = "7006" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Ac_Code ,"
        LngId = "7008" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Dr ,"
        LngId = "7009" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Cr ,"
        LngId = "7010" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Amt_LAK ,"
        LngId = "7011" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Sign ,"
        LngId = "7012" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_SignNme ,"
        LngId = "7026" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InDate ,"
        LngId = "7028" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_AmtCurr ,"
        LngId = "7031" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_User ,"
        LngId = "7042" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_OpenAmt ,"
        LngId = "7039" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_TotalAmt ,"
        LngId = "7040" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Balance ,"
        LngId = "7043" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RemAmt ,"
        If RM.Checked = True Then
            LngId = "7060" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7061" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        ElseIf RP.Checked = True Then
            LngId = "7083" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7063" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        ElseIf Rhalf.Checked = True Then
            LngId = "7080" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7081" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        ElseIf RY.Checked = True Then
            LngId = "7084" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7065" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        End If

        If r1.Checked = True Then
            LngId = "7052" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            r = "And GRP>0"
        ElseIf r2.Checked = True Then
            LngId = "7053" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            r = "And GRP<5"
        ElseIf r3.Checked = True Then
            LngId = "7054" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            r = "And GRP>5"
        End If
        SLF = "SELECT  " & MuLngRpt & "  *   FROM Ap_Rpt_BLS_BOL_Detail  "
        Call LoadLoGO()
         Dim dt As DataTable = DbHelper.GetDataTable(SLF)

         If dt.Rows.Count = 0 Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
        Dim FrmPreview As New FmPreview : FrmClosing()
         Dim Rpt As New CryRpt_BLS_BOL_Item
         If MdShowLOGO = 1 Then
             Rpt.Subreports(0).SetDataSource(RsLOGO)
         End If
         Rpt.SetDataSource(dt)
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.DisplayGroupTree = False

        FrmPreview.WindowState = FormWindowState.Maximized
        FrmPreview.ShowDialog()
        FrmPreview.Focus()
    End Sub


    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Call Close()
    End Sub

    Private Sub RM_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RM.CheckedChanged
        selectLoad()
    End Sub

    Private Sub RP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RP.CheckedChanged
        selectLoad()
    End Sub

    Private Sub RD_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RD.CheckedChanged
        selectLoad()
    End Sub

    Private Sub RY_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RY.CheckedChanged
        selectLoad()
    End Sub

    Private Sub Period_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Period.SelectedIndexChanged
        LoadPeriod()
    End Sub

    Private Sub Pyy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Pyy.ValueChanged
        LoadPeriod()
    End Sub

    Private Sub DMonth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DMonth.SelectedIndexChanged
        LoadMonth()
    End Sub

    Private Sub Myy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Myy.ValueChanged
        LoadMonth()
    End Sub

    Private Sub Ds_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ds.ValueChanged
        Dt.Value = Ds.Value
        LoadDay()
    End Sub

    Private Sub Dt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Dt.ValueChanged
        LoadDay()
    End Sub

    Private Sub yy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yy.ValueChanged
        Toyy.Value = yy.Value
        Call LoadYear()
    End Sub

    Private Sub Toyy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Toyy.ValueChanged
        Call LoadYear()
    End Sub

    Private Sub RaParent_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RaParent.CheckedChanged

    End Sub

    Private Sub FmRpt_BLS_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Call MdiCNum()
    End Sub

    Private Sub FmRpt_BLS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        Ds.Text = MWorkSetting
        Myy.Text = MWorkSetting
        yy.Text = MWorkSetting
        Toyy.Text = MWorkSetting
        'MsgBox(MWorkSetting)
        Pyy.Text = MWorkSetting

        RM.Checked = True
        Ct.SelectedIndex = 0
        If Month(MWorkSetting) = 1 Then
            DMonth.SelectedIndex = 0
            Period.SelectedIndex = 0
            M1.SelectedIndex = 0
            M2.SelectedIndex = 0

        ElseIf Month(MWorkSetting) = 2 Then
            DMonth.SelectedIndex = 1
            Period.SelectedIndex = 0
            M1.SelectedIndex = 1
            M2.SelectedIndex = 1

        ElseIf Month(MWorkSetting) = 3 Then
            DMonth.SelectedIndex = 2
            Period.SelectedIndex = 0
            M1.SelectedIndex = 2
            M2.SelectedIndex = 2

        ElseIf Month(MWorkSetting) = 4 Then
            DMonth.SelectedIndex = 3
            Period.SelectedIndex = 1
            M1.SelectedIndex = 3
            M2.SelectedIndex = 3
        ElseIf Month(MWorkSetting) = 5 Then
            DMonth.SelectedIndex = 4
            Period.SelectedIndex = 1
            M1.SelectedIndex = 4
            M2.SelectedIndex = 4

        ElseIf Month(MWorkSetting) = 6 Then
            DMonth.SelectedIndex = 5
            Period.SelectedIndex = 1
            M1.SelectedIndex = 5
            M2.SelectedIndex = 5

        ElseIf Month(MWorkSetting) = 7 Then
            DMonth.SelectedIndex = 6
            Period.SelectedIndex = 2
            M1.SelectedIndex = 6
            M2.SelectedIndex = 6

        ElseIf Month(MWorkSetting) = 8 Then
            DMonth.SelectedIndex = 7
            Period.SelectedIndex = 2
            M1.SelectedIndex = 7
            M2.SelectedIndex = 7
            M1.SelectedIndex = 7
            M2.SelectedIndex = 7

        ElseIf Month(MWorkSetting) = 9 Then
            DMonth.SelectedIndex = 8
            Period.SelectedIndex = 2
            M1.SelectedIndex = 8
            M2.SelectedIndex = 8

        ElseIf Month(MWorkSetting) = 10 Then
            DMonth.SelectedIndex = 9
            Period.SelectedIndex = 3
            M1.SelectedIndex = 9
            M2.SelectedIndex = 9

        ElseIf Month(MWorkSetting) = 11 Then
            DMonth.SelectedIndex = 10
            Period.SelectedIndex = 3
            M1.SelectedIndex = 10
            M2.SelectedIndex = 10

        ElseIf Month(MWorkSetting) = 12 Then
            DMonth.SelectedIndex = 11
            Period.SelectedIndex = 3
            M1.SelectedIndex = 11
            M2.SelectedIndex = 11

        End If

        Call selectLoad()
        Call Click_Last()
        'SetControlText(Me)
        Call loadOffice_User()
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        FmRpt_BLS_BOL_Item.ShowDialog()
        FmRpt_BLS_BOL_Item.Focus()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub Click_Last()
        If RD.Checked = True Then
            ny = Format(CDate(MdStartDate), "MM/yyyy")
            ly = Format(CDate(MdStartDate_Last), "MM/yyyy")
            MdStartDate_Last = DateAdd(DateInterval.Day, -1, MdStartDate)
            MdToDate_Last = DateAdd(DateInterval.Day, -1, MdToDate)
        ElseIf RM.Checked = True Then
            MdStartDate_Last = DateAdd(DateInterval.Month, -1, MdStartDate)
            MdToDate_Last = DateAdd(DateInterval.Day, -1, MdStartDate)
            LL6.Text = ""
            If DMonth.SelectedIndex > 0 Then
                LL6.Text = Format(MdStartDate_Last, "MM/yyyy")
            End If
            LL5.Text = Format(MdStartDate, "MM/yyyy")
        ElseIf RP.Checked = True Then
            MdStartDate_Last = DateAdd(DateInterval.Month, -3, MdStartDate)
            MdToDate_Last = DateAdd(DateInterval.Day, -1, MdStartDate)
            LL6.Text = ""
            If Period.SelectedIndex > 0 Then
                LL6.Text = "ງວດທີ " & Period.SelectedIndex
            End If
            LL5.Text = Period.Text
        ElseIf RY.Checked = True Then
            MdStartDate_Last = DateAdd(DateInterval.Year, -1, MdStartDate)
            MdToDate_Last = DateAdd(DateInterval.Day, -1, MdStartDate)
            LL5.Text = Format(MdStartDate, "yyyy")
            LL6.Text = Format(MdStartDate_Last, "yyyy")
        End If
        LL2.Text = MdStartDate_Last
        LL4.Text = MdToDate_Last
        LL1.Text = MdStartDate
        LL3.Text = MdToDate
    End Sub
    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub Button2_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Click_Last()
    End Sub

    Private Sub half_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles half.SelectedIndexChanged
        LoadHalfYear()
    End Sub
    Private Sub LoadHalfYear()

        If M1.SelectedIndex < M2.SelectedIndex Then
            Dim s As Double = M1.SelectedIndex + 2
            Dim x As Double = M2.SelectedIndex + 1
            MdStartDate = Format(CDate("01/" & s & "/" & Year(Hyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("01/" & x & "/" & Year(Hyy.Value)), "dd-MM-yyyy")
            Dim SM As Date = DateAdd("M", CDbl(1), MdToDate)
            Dim SD As Date = DateAdd("d", CDbl(-1), SM)
            MdToDate = Format(CDate(SD), "dd-MM-yyyy")
            Lb.Text = "ປະຈຳເດືອນ " & M1.Text & " ຫາ " & M2.Text & "/" & yy.Text
        Else
            Dim s As Double = M1.SelectedIndex + 1
            Dim x As Double = M2.SelectedIndex + 1
            MdStartDate = Format(CDate("01/" & s & "/" & Year(Hyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("01/" & x & "/" & Year(Hyy.Value)), "dd-MM-yyyy")
            Dim SM As Date = DateAdd("M", CDbl(1), MdToDate)
            Dim SD As Date = DateAdd("d", CDbl(-1), SM)
            MdToDate = Format(CDate(SD), "dd-MM-yyyy")
            Lb.Text = "ປະຈຳເດືອນ " & M1.Text & "/" & yy.Text
        End If

        L5.Text = MdStartDate & " => " & MdToDate
        'ElseIf half.SelectedIndex = 1 Then
        '    MdStartDate = Format(CDate("01/7/" & Year(Hyy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("31/12/" & Year(Hyy.Value)), "dd-MM-yyyy")
        '    Lb.Text = "6 ເດືອນທ້າຍປີ " & yy.Text
        '    L5.Text = MdStartDate & " => " & MdToDate

    End Sub
    Private Sub Rhalf_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        selectLoad()
    End Sub

    Private Sub Label27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label27.Click

    End Sub

    Private Sub Off_Usr_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Off_Usr.SelectedIndexChanged

    End Sub

    Private Sub Rhalf_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Rhalf.CheckedChanged
        selectLoad()
    End Sub

    Private Sub RT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RT.CheckedChanged
        selectLoad()
    End Sub

    Private Sub M1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles M1.SelectedIndexChanged
        If M2.SelectedIndex < M1.SelectedIndex Then
            M2.SelectedIndex = M1.SelectedIndex
        End If

        LoadHalfYear()
    End Sub

    Private Sub M2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles M2.SelectedIndexChanged
        If M2.SelectedIndex < M1.SelectedIndex Then
            M1.SelectedIndex = M2.SelectedIndex
        End If


        LoadHalfYear()
    End Sub

    Private Sub yyt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yyt.ValueChanged
        selectLoad()
    End Sub

    Private Sub Ct_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ct.SelectedIndexChanged
        selectLoad()
    End Sub
End Class