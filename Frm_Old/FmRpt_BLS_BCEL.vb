Imports CrystalDecisions.Shared

Public Class FmRpt_BLS_BCEL
    Dim r As String
    Dim CLT_Str, CLT_Last_Str As String
    Dim bls1 As String
    Dim MonthLetter1 As String
    Dim MonthLetter_Last As String
    Dim Month_IN As String

    Dim Month_Last As String

    Dim MdStartDate As Date
    Dim MdToDate As Date
    Dim MdStartDate_MM As Date
    Dim MdToDate_MM As Date
    Dim Month_IN_MM As String
    Dim MdStartDate_Last As Date
    Dim MdToDate_Last As Date
    Dim ny, ly, n_L_y As String

    Dim sql As String
    Dim AmtOpenDR, AmtOpenCR, AmtOpenMonthDR, AmtOpenMonthCR As Double
    Dim VCode1, VCode2, VCode3, VCode4, VCode5, VCode6, VCode7, VCode8, VCode9 As String
    Dim MdQuarter As Date
    Dim RsOpen As New ADODB.Recordset
    Dim RsOpenMonth As New ADODB.Recordset
    Dim RsRpt As New ADODB.Recordset
    Dim VOpenDate As Date
    Dim RptNme As String
    Dim RSC12 As New ADODB.Recordset
    Dim RSCIn_M As New ADODB.Recordset
    Public Rpt As New Object
    Private Sub HeaDer()
        LoadSqlData("SELECT * FROM Header WHERE ID=N'B01' ", RSC)
        If RSC.RecordCount <> 0 Then
            TxtHeader.Text = Trim(RSC.Fields("Nm").Value.ToString)
            If MuLng = "L" Then
                TxtS1.Text = Trim(RSC.Fields("S1").Value.ToString)
                TxtS2.Text = Trim(RSC.Fields("S2").Value.ToString)
                TxtS3.Text = Trim(RSC.Fields("S3").Value.ToString)
                TxtS4.Text = Trim(RSC.Fields("S4").Value.ToString)
                TxtPP.Text = Trim(RSC.Fields("pp").Value.ToString)
            Else
                TxtS1.Text = Trim(RSC.Fields("S1e").Value.ToString)
                TxtS2.Text = Trim(RSC.Fields("S2e").Value.ToString)
                TxtS3.Text = Trim(RSC.Fields("S3e").Value.ToString)
                TxtS4.Text = Trim(RSC.Fields("S4e").Value.ToString)
                TxtPP.Text = Trim(RSC.Fields("ppe").Value.ToString)
            End If



        End If
    End Sub
    Private Sub AddHeader()
        If MuLng = "L" Then
            LoadSqlData("SELECT * FROM Header WHERE ID=N'B01' ", RSC)
            If RSC.RecordCount = 0 Then
                CNN.Execute("INSERT INTO Header(ID,Nm,S1,S2,S3,S4,PP) " & _
                            " values('B01',N'" & TxtHeader.Text & "',N'" & TxtS1.Text & "',N'" & TxtS2.Text & "',N'" & TxtS3.Text & "',N'" & TxtS4.Text & "',N'" & TxtPP.Text & "') ")
            Else
                CNN.Execute("UPDATE Header set Nm=N'" & TxtHeader.Text & "',S1=N'" & TxtS1.Text & "',S2=N'" & TxtS2.Text & "',S3=N'" & TxtS3.Text & "',S4=N'" & TxtS4.Text & "',PP=N'" & TxtPP.Text & "' " & _
                            " where ID='B01' ")
            End If
        Else
            LoadSqlData("SELECT * FROM Header WHERE ID=N'B01' ", RSC)
            If RSC.RecordCount = 0 Then
                CNN.Execute("INSERT INTO Header(ID,Nm,S1e,S2e,S3e,S4e,PPe) " & _
                            " values('B01',N'" & TxtHeader.Text & "',N'" & TxtS1.Text & "',N'" & TxtS2.Text & "',N'" & TxtS3.Text & "',N'" & TxtS4.Text & "',N'" & TxtPP.Text & "') ")
            Else
                CNN.Execute("UPDATE Header set Nm=N'" & TxtHeader.Text & "',S1e=N'" & TxtS1.Text & "',S2e=N'" & TxtS2.Text & "',S3e=N'" & TxtS3.Text & "',S4e=N'" & TxtS4.Text & "',PPe=N'" & TxtPP.Text & "' " & _
                            " where ID='B01' ")
            End If
        End If

    End Sub
    Private Sub ChangBalance()
        New_Code = "3901000"
        New_Code4 = "00.3901000"
        Code_Dr = "4"
        Code_Dr1 = "00.4"
        Code_Cr = "5"
        Code_Cr1 = "00.5"
        Ac_Code = ""
        'MsgBox(MdStartDate & "==" & MdToDate)

        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()


        CNN.Execute("DELETE  Ap_balance_6_col ")
        CNN.Execute("DELETE FROM Ap_balance_6 ")
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")
        Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE    date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr)as open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        Call Chang_Incom()
        CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + amt_dr) - (open_amt_cr + amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr + amt_cr) ")
        CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr + amt_cr) - (open_amt_dr + amt_dr) where (open_amt_cr + amt_cr) >= (open_amt_dr + amt_dr) ")
        CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")

    End Sub

    Private Sub Chang_Incom()
        'MDACC00 = 0
        If MDACC00 = "0" Then
            New_Code = New_Code
            Insr = "delete  Ap_balance_6  " & _
   "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  )   select '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr) from Ap_balance_6_col where  left(Ac_Code,1) ='" & Code_Dr & "'   Or left(Ac_Code,1)='" & Code_Cr & "'   or  Ac_Code =  '" & New_Code & "'    " & _
"update  Ap_balance_6 set open_amt_dr = 0 where open_amt_dr  is null  " & _
"update  Ap_balance_6 set open_amt_cr = 0 where open_amt_cr  is null  " & _
"update  Ap_balance_6 set amt_dr = 0 where amt_dr  is null  " & _
"update  Ap_balance_6 set amt_cr = 0 where amt_cr  is null   " & _
"Update  Ap_balance_6 set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr   " & _
"Update  Ap_balance_6 set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr " & _
"Update  Ap_balance_6 set   amt_dr = amt_dr  - amt_cr , amt_cr=0  where amt_dr  >= amt_cr   " & _
"Update  Ap_balance_6 set   amt_cr = amt_cr  - amt_dr , amt_dr=0  where amt_cr  >= amt_dr " & _
"delete  Ap_balance_6_col  where   left(Ac_Code,1) ='" & Code_Dr & "'  Or  left(Ac_Code,1)='" & Code_Cr & "' or   Ac_Code =  '" & New_Code & "'    " & _
"  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  from Ap_balance_6"
            CNN.Execute(Insr)
        Else

            New_Code = New_Code4
            Insr = "delete  Ap_balance_6  " & _
   "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  )   select '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr) from Ap_balance_6_col where   left(Ac_Code,4) ='" & Code_Dr1 & "'  or left(Ac_Code,4)='" & Code_Cr1 & "'    or  Ac_Code =  '" & New_Code & "'  " & _
"update  Ap_balance_6 set open_amt_dr = 0 where open_amt_dr  is null  " & _
"update  Ap_balance_6 set open_amt_cr = 0 where open_amt_cr  is null  " & _
"update  Ap_balance_6 set amt_dr = 0 where amt_dr  is null  " & _
"update  Ap_balance_6 set amt_cr = 0 where amt_cr  is null   " & _
"Update  Ap_balance_6 set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr   " & _
"Update  Ap_balance_6 set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr " & _
"Update  Ap_balance_6 set   amt_dr = amt_dr  - amt_cr , amt_cr=0  where amt_dr  >= amt_cr   " & _
"Update  Ap_balance_6 set   amt_cr = amt_cr  - amt_dr , amt_dr=0  where amt_cr  >= amt_dr " & _
"delete  Ap_balance_6_col  where    left(Ac_Code,4) ='" & Code_Dr1 & "'  or left(Ac_Code,4)='" & Code_Cr1 & "'  or  Ac_Code =  '" & New_Code & "'   " & _
"  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  from Ap_balance_6"
            CNN.Execute(Insr)
        End If

        '      Insr = "delete  Ap_balance_6  " & _
        '         "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  )   select '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr) from Ap_balance_6_col where  left(Ac_Code,1) ='" & Code_Dr & "' or left(Ac_Code,4) ='" & Code_Dr1 & "'  Or left(Ac_Code,1)='" & Code_Cr & "' or left(Ac_Code,4)='" & Code_Cr1 & "'  or  Ac_Code =  '" & New_Code & "'  or  Ac_Code =  '" & New_Code4 & "'  " & _
        '"update  Ap_balance_6 set open_amt_dr = 0 where open_amt_dr  is null  " & _
        '"update  Ap_balance_6 set open_amt_cr = 0 where open_amt_cr  is null  " & _
        '"update  Ap_balance_6 set amt_dr = 0 where amt_dr  is null  " & _
        '"update  Ap_balance_6 set amt_cr = 0 where amt_cr  is null   " & _
        ' "Update  Ap_balance_6 set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr   " & _
        '"Update  Ap_balance_6 set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr " & _
        '"Update  Ap_balance_6 set   amt_dr = amt_dr  - amt_cr , amt_cr=0  where amt_dr  >= amt_cr   " & _
        '"Update  Ap_balance_6 set   amt_cr = amt_cr  - amt_dr , amt_dr=0  where amt_cr  >= amt_dr " & _
        '   "delete  Ap_balance_6_col  where   left(Ac_Code,1) ='" & Code_Dr & "' or left(Ac_Code,4) ='" & Code_Dr1 & "'  Or  left(Ac_Code,1)='" & Code_Cr & "' or left(Ac_Code,4)='" & Code_Cr1 & "'  or  Ac_Code =  '" & New_Code & "'  or  Ac_Code =  '" & New_Code4 & "'  " & _
        '     "  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  from Ap_balance_6"
        '      CNN.Execute(Insr)
    End Sub

    Private Sub BLNEW()


        New_Code = "3901000"
        New_Code4 = "00.3901000"

        Code_Dr = "4"
        Code_Dr1 = "00.4"
        Code_Cr = "5"
        Code_Cr1 = "00.5"

        Ac_Code = ""
        'MsgBox(MdStartDate & "==" & MdToDate)

        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        Dim B_Curr As String = ""
        If CMB_Curr.SelectedIndex = 0 Then
            B_Curr = ""
        Else
            B_Curr = " AND  Curr=N'" & CMB_Curr.Text & "' "
        End If


        CNN.Execute("DELETE  Ap_balance_6_col ")
        CNN.Execute("DELETE FROM Ap_balance_6 ")
        If CMB_Curr.SelectedIndex = 0 Then
            If CheckBox4.Checked = True Then
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code, open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code, 0, 0, sum(amt_Dr) / " & CDbl(txtRate2.Text) & " , sum(amt_Cr) / " & CDbl(txtRate2.Text) & " from gen_jn  " & _
                " WHERE  1=1 and date_work BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")
                Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code , sum(amt_Dr) / " & CDbl(txtRate2.Text) & " , sum(amt_Cr) / " & CDbl(txtRate2.Text) & " , 0 , 0 from gen_jn  " & _
                " WHERE  1=1 and date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code  , sum(amt_Dr) / " & CDbl(txtRate2.Text) & " , sum(amt_Cr) / " & CDbl(txtRate2.Text) & " , 0 , 0 from Open_jn " & _
                " WHERE 1=1 and date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
            ElseIf CheckBox6.Checked = True Then
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code, open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code, 0, 0, sum(amt_Dr) / " & CDbl(txtRate.Text) & " , sum(amt_Cr) / " & CDbl(txtRate.Text) & " from gen_jn  " & _
                " WHERE  1=1 and date_work BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")
                Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code , sum(amt_Dr) / " & CDbl(txtRate.Text) & " , sum(amt_Cr) / " & CDbl(txtRate.Text) & " , 0 , 0 from gen_jn  " & _
                " WHERE  1=1 and date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code  , sum(amt_Dr) / " & CDbl(txtRate.Text) & " , sum(amt_Cr) / " & CDbl(txtRate.Text) & " , 0 , 0 from Open_jn " & _
                " WHERE 1=1 and date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

            Else

                '=======LAK===
                Dim GGG As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
               " select ac_code , 0 , 0 , sum(Amt_Dr) , sum(Amt_cr) from gen_jn " & _
               " WHERE 1=1  and date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code "
                CNN.Execute(GGG)
                Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
                '=======LAK===
                Dim PPP As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                        " select ac_code , sum(Amt_Dr) , sum(Amt_cr) , 0 , 0  from gen_jn  WHERE 1=1 and date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code"
                CNN.Execute(PPP)
                '        '=======LAK===
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code , sum(Amt_Dr) , sum(Amt_cr) , 0 , 0  from Open_jn WHERE 1=1 and date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

            End If
        Else
            CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
  " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(Amount_Dr)as amt_dr , sum(Amount_Cr)as amt_cr  from gen_jn  WHERE  1=1 " & B_Curr & " and  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")
            Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
            CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            " select ac_code , sum(Amount_Dr)as amt_dr , sum(Amount_Cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE 1=1 " & B_Curr & " and  gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
            CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            " select ac_code  , sum(Amount_Dr) as amt_dr , sum(Amount_Cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE  1=1 " & B_Curr & " and    date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

        End If
        CNN.Execute("UPDATE Ap_balance_6 set Ac_Code = left(Ac_Code,7) ")


        CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr)as open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")
        If CMB_Curr.SelectedIndex = 0 Then
            '        CNN.Execute(" insert into Ap_balance_6_col(ac_code, ac_name, open_amt_dr, open_amt_cr,  amt_dr, amt_cr,   Rem_dr, Rem_cr,Status) " & _
            '" Select '2382120.00.0000',N'ຄູ່ມູນຄ່າຖານະ ແລກປ່ຽນເງິນຕາ ຕ່າງປະເທດ',sum(open_amt_dr),sum(open_amt_cr),sum(amt_dr),sum(amt_cr),sum(Rem_dr),sum(Rem_cr),1   " & _
            '" from Ap_balance_6_col where left(ac_code,7)='2382120' group by  left(ac_code,7) ")
            '        CNN.Execute("delete  Ap_balance_6_col  where left(ac_code,7)='2382120' and Status is null ")
        ElseIf CMB_Curr.SelectedIndex = 1 Then
            CNN.Execute(" insert into Ap_balance_6_col(ac_code, ac_name, open_amt_dr, open_amt_cr,  amt_dr, amt_cr,   Rem_dr, Rem_cr,Status) " & _
" Select '2382120.00.0000',N'ຄູ່ມູນຄ່າຖານະ ແລກປ່ຽນເງິນຕາ ຕ່າງປະເທດ',sum(open_amt_dr),sum(open_amt_cr),sum(amt_dr),sum(amt_cr),sum(Rem_dr),sum(Rem_cr),1   " & _
" from Ap_balance_6_col where left(ac_code,7)='2382120' group by  left(ac_code,7) ")
            CNN.Execute("delete  Ap_balance_6_col  where left(ac_code,7)='2382120' and Status is null ")
        Else

            'CNN.Execute("delete  Ap_balance_6_col  where left(ac_code,7)='2382120'")

        End If

        CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        Call Chang_Incom()
        CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + amt_dr) - (open_amt_cr + amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr + amt_cr) ")
        CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr + amt_cr) - (open_amt_dr + amt_dr) where (open_amt_cr + amt_cr) >= (open_amt_dr + amt_dr) ")
        CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        'CNN.Execute("UPDATE Ap_balance_6_col set Ac_Code = left(Ac_Code,7) ")
    End Sub

    Private Sub BtnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        Call AddHeader()
        Call selectLoad()
        Call Office()
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        'Call ChangBalance()
        'BLNEW()
        'CNN.Execute("update  Ap_Rpt_BLS_Item_Old set Last_amt_dr  =  " & CDbl(0) & " , Last_amt_cr  = " & CDbl(0) & " ,amt_dr  =  " & CDbl(0) & " , amt_cr  = " & CDbl(0) & "  ")
        'CNN.Execute("update Ap_Rpt_BLS_Old set Last_Amt = 0 , Amt  =  " & CDbl(0) & "  ")
        'CNN.Execute("DELETE FROM Ap_Rpt_BLS_Detail ")
        'SelcectIn()
        'UpdateIIn()
        'SelectOut()
        'UpdateOut()
        'Update_Sum()
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

            'Lb.Text = s1 & " " & s2 & "/" & Year(MdToDate)
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
            ' Lb.Text = s3 & " " & yy.Text
        End If

        If r1.Checked = True Then
            CRLD = " "
        ElseIf r2.Checked = True Then
            CRLD = "And GRP=1"
        ElseIf r3.Checked = True Then
            CRLD = "And GRP=2"
        End If
        Call BLNEW()
        CNN.Execute("update  Ap_Rpt_BLS_Item_Old set Last_amt_dr  =  " & CDbl(0) & " , Last_amt_cr  = " & CDbl(0) & " ,amt_dr  =  " & CDbl(0) & " , amt_cr  = " & CDbl(0) & "  ")
        CNN.Execute("update Ap_Rpt_BLS_Old set Last_Amt = 0 , Amt  =  " & CDbl(0) & "  ")
       
        If r1.Checked = True Then
            CNN.Execute("DELETE FROM Ap_Rpt_BLS_Detail ")
            CRLD = "And GRP=1"
            CNN.Execute("update  Ap_Rpt_BLS_Item_Old set Last_amt_dr  =  " & CDbl(0) & " , Last_amt_cr  = " & CDbl(0) & " ,amt_dr  =  " & CDbl(0) & " , amt_cr  = " & CDbl(0) & " where GRP=1 ")
            CNN.Execute("update Ap_Rpt_BLS_Old set Last_Amt = 0 , Amt  =  " & CDbl(0) & "  where GRP=1  ")
            SelcectIn()
            UpdateIIn()
            SelectOut()
            UpdateOut()
            Update_Sum()
            '====================GRP=2====================================
            CNN.Execute("DELETE FROM Ap_Rpt_BLS_Detail ")
            CRLD = "And GRP=2"
            CNN.Execute("update  Ap_Rpt_BLS_Item_Old set Last_amt_dr  =  " & CDbl(0) & " , Last_amt_cr  = " & CDbl(0) & " ,amt_dr  =  " & CDbl(0) & " , amt_cr  = " & CDbl(0) & " where GRP=2 ")
            CNN.Execute("update Ap_Rpt_BLS_Old set Last_Amt = 0 , Amt  =  " & CDbl(0) & "  where GRP=2  ")
            SelcectIn()
            UpdateIIn()
            SelectOut()
            UpdateOut()
            Update_Sum()
            '====================06.10====================================
            CNN.Execute(" update Ap_Rpt_BLS_Old set Amt= ( select sum(Amt_Cr-Amt_Dr) from Ap_Rpt_BLS_Item_Old where Rpt_ID='06.10'  )  from Ap_Rpt_BLS_Old where Rpt_ID='06.10' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt= ( select sum(Last_Amt_Cr-Last_Amt_Dr) from Ap_Rpt_BLS_Item_Old where Rpt_ID='06.10' )  from Ap_Rpt_BLS_Old where Rpt_ID='06.10' And GRP=2 ")
        ElseIf r2.Checked = True Then
            CNN.Execute("DELETE FROM Ap_Rpt_BLS_Detail ")
            SelcectIn()
            UpdateIIn()
            SelectOut()
            UpdateOut()
            Update_Sum()
        ElseIf r3.Checked = True Then

            CNN.Execute("DELETE FROM Ap_Rpt_BLS_Detail ")
            SelcectIn()
            UpdateIIn()
            SelectOut()
            UpdateOut()
            Update_Sum()
            '====================06.10====================================
            CNN.Execute(" update Ap_Rpt_BLS_Old set Amt= ( select sum(Amt_Cr-Amt_Dr) from Ap_Rpt_BLS_Item_Old where Rpt_ID='06.10' )  from Ap_Rpt_BLS_Old where Rpt_ID='06.10'  ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt= ( select sum(Last_Amt_Cr-Last_Amt_Dr) from Ap_Rpt_BLS_Item_Old where Rpt_ID='06.10' )  from Ap_Rpt_BLS_Old where Rpt_ID='06.10'  ")
        End If

        If CMB_Curr.Text = "LAK" Then
            CURR01 = "ຫົວໜ່ວຍ : ກີບ"
        ElseIf CMB_Curr.Text = "USD" Then
            CURR01 = "ຫົວໜ່ວຍ : ໂດລາ"
        Else
            CURR01 = "ຫົວໜ່ວຍ : ກີບ"
        End If
        If CheckBox1.Checked = False Then
            Call LoadReport()
        Else
            Call LoadReportItem()
        End If
        'MdStartDate = d1
        'MdToDate = d2
    End Sub
    Private Sub M_M()
        Dim RPT_ID As String

        Dim AA As String = Lb.Text
        RPT_ID = " "
        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As RptSjUd ,"
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

        If CMB_Curr.Text = "EQVL" Then
            If CheckBox4.Checked = True Then
                LngId = "7121" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"

            Else
                LngId = "7106" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"

            End If
        ElseIf CMB_Curr.Text = "LAK" Then
            LngId = "7097" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        ElseIf CMB_Curr.Text = "USD" Then
            LngId = "7108" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        End If


        'If CMB_Curr.SelectedIndex = 0 Then
        '    LngId = "7106" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        'ElseIf CMB_Curr.SelectedIndex = 1 Then
        '    LngId = "7107" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        'ElseIf CMB_Curr.SelectedIndex = 2 Then
        '    LngId = "7108" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        'End If

        If RD.Checked = True Then
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"
            LngId = "7090" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7091" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        ElseIf RM.Checked = True Then
            LngId = "7060" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7061" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"

            LngId = "7096" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
        ElseIf RP.Checked = True Then
            LngId = "7083" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7063" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7112" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"
        ElseIf Rhalf.Checked = True Then
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"
            LngId = "7060" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7061" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"

        ElseIf RT.Checked = True Then
            LngId = "7080" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7081" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"
        ElseIf RY.Checked = True Then
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"
            LngId = "7084" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7065" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7113" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
        End If





        If r1.Checked = True Then
            LngId = "7092" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            r = "And GRP>0"
            'r = "And GRP<4"
        ElseIf r2.Checked = True Then
            LngId = "7093" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            'r = "And GRP<5"
            r = "And GRP<2"
        ElseIf r3.Checked = True Then
            LngId = "7054" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            'r = "And GRP>5"
            'r = "And GRP>1 and GRP<=3"
            r = "And GRP>1"
        End If


        Dim Doc As Integer = 0
        If CheckBox5.Checked = True Then
            Doc = 1
        Else
            Doc = 0
        End If
        SLF = "SELECT    " & Doc & "  as Doc  ,  N'" & MuOffDep & "'  as RptSjoff_Dep  ,  " & mformat & "  as mformat  ,  " & MuLngRpt & "  *, N'" & Month_IN & "' as In_Mn, N'" & Month_IN_MM & "' as Last_Mn   FROM Ap_Rpt_BLS_Old_M_M  "



        Call LoadLoGO()

        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            .Open(" " & SLF & " where grp<>'' " & RPT_ID & " " & r & "order by Rpt_Id asc  ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        Dim FrmPreview As New FmPreview : FrmClosing()

        'Dim Rpt As New CryRpt_BLSM_M
        Dim Rpt As New Object
        If RM.Checked = True Then
            Rpt = New CryRpt_BLS_M
        Else
            Rpt = New CryRpt_BLS
        End If
        'Dim Rpt As New CryRpt_BLS_NNN
        If MdShowLOGO = 1 Then
            Rpt.Subreports(0).SetDataSource(RsLOGO)
        End If
        Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S1"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS1.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S2"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS2.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S3"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS3.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S4"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS4.Text

        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("pp"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtPP.Text


        Rpt.SetDataSource(Rs)
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.DisplayGroupTree = False
        FrmPreview.WindowState = FormWindowState.Maximized
        FrmPreview.Show()
        FrmPreview.Focus()
    End Sub
    Private Sub LoadOpen_Jn16()
        Dim RSC16 As New ADODB.Recordset
        LoadSqlData("  select sum(rem_cr - rem_dr) as x  , count(rem_cr - rem_dr) as y  from Ap_balance_6_col  where Ac_Code Like '1%' or Ac_Code Like '2%' or Ac_Code Like '3%' ", RSC16)

        If CDbl(Trim(RSC16.Fields("y").Value)) > 0 Then

            If CDbl(Trim(RSC16.Fields("x").Value)) > 0 Then
                Rem_Cr = CDbl(Trim(RSC16.Fields("x").Value))
                Rem_Dr = 0
            End If
            If CDbl(Trim(RSC16.Fields("x").Value)) < 0 Then
                Rem_Dr = CDbl(Trim(RSC16.Fields("x").Value)) * CDbl(-1)
                Rem_Cr = 0
            End If
            If CDbl(Trim(RSC16.Fields("x").Value)) <> 0 Then
                Call LoadOpen_Jn17()
            End If
        End If


    End Sub
    Private Sub LoadOpen_Jn17()
        Dim RSC17 As New ADODB.Recordset
        LoadSqlData("   select Ac_Code from Ap_balance_6_col  where Ac_Code ='65'", RSC17)
        If RSC17.RecordCount <> 0 Then
            CNN.Execute(" Update Ap_balance_6_col set Amt_Dr = " & CDbl(Rem_Dr) & "  , Amt_Cr =" & CDbl(Rem_Cr) & " ")
        Else
            CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code ,ac_name , ac_namee , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
             "Values('65', N'" & "***" & "', '" & "***" & "', " & _
             " " & CDbl(0) & ", " & CDbl(0) & ", " & CDbl(Rem_Dr) & ", " & CDbl(Rem_Cr) & ",0 )")
        End If

        CNN.Execute(" Delete Ap_balance_6_col  where Ac_Code Like '1%' or Ac_Code Like '2%' or Ac_Code Like '3%'  ")

    End Sub
    Private Sub SelcectIn()
        If CMB_Curr.SelectedIndex = 0 Then
            CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
        "where Ap_Rpt_BLS_Item_Old.ac_code = Ap_balance_6_col.ac_code And  Rpt_Type = 'In'  and left(Ap_balance_6_col.ac_code,7)<>'2382120' " & CRLD & " ")

        ElseIf CMB_Curr.SelectedIndex = 1 Then
            CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
              "where Ap_Rpt_BLS_Item_Old.ac_code = Ap_balance_6_col.ac_code And  Rpt_Type = 'In' " & CRLD & " ")
        Else
            CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
     "where Ap_Rpt_BLS_Item_Old.ac_code = Ap_balance_6_col.ac_code And  Rpt_Type = 'In'  and left(Ap_balance_6_col.ac_code,7)<>'2382120' " & CRLD & " ")
        End If


        CNN.Execute("Insert into Ap_Rpt_BLS_Detail (  Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type )" & _
         " select   Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type from Ap_Rpt_BLS_Item_Old where Rpt_Type = 'In' And ( Amt_Dr <>0 or Amt_Cr <>0  or Last_Amt_Dr <>0 or Last_Amt_Cr <>0 )  " & CRLD & "  ")

    End Sub


    Private Sub SelcectInLast()
        LoadSqlData("select * from Ap_Rpt_BLS_Item_Old where  Rpt_Type = 'In'", RSCIn_M)
        With RSCIn_M
            Do Until .EOF = True
                UpdateIIn_ItemLast()
                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub UpdateIIn_Item()
        Dim RSCkk As New ADODB.Recordset
        LoadSqlData(" select * from Ap_balance_6_col   where ac_code =  '" & (RSCIn_M.Fields("Ac_Code").Value) & "' ", RSCkk)
        With RSCkk
            Do Until .EOF = True
                CNN.Execute("Insert into Ap_Rpt_BLS_Detail (  Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type ) values ( '" & CStr((RSCIn_M.Fields("Rpt_Id").Value)) & "' , '" & CStr((.Fields("Ac_Code").Value)) & "' , N'" & CStr((.Fields("Ac_Name").Value)) & "'   , " & CDbl((.Fields("open_amt_dr").Value)) & " , " & CDbl((.Fields("open_amt_Cr").Value)) & "   , " & CDbl((.Fields("Rem_dr").Value)) & " , " & CDbl((.Fields("Rem_cr").Value)) & " , 'In')")
                CNN.Execute("update  Ap_Rpt_BLS_Item_Old set  Last_amt_dr  =  Last_amt_dr+" & CDbl((.Fields("open_amt_dr").Value)) & " , Last_amt_cr  = Last_amt_cr+" & CDbl((.Fields("open_amt_Cr").Value)) & " , Amt_Dr  =  Amt_Dr+" & CDbl((.Fields("Rem_dr").Value)) & " , Amt_Cr  = Amt_Cr+" & CDbl((.Fields("Rem_cr").Value)) & "   where ac_code = '" & (RSCIn_M.Fields("ac_code").Value) & "' And  Rpt_Type = 'In' ")

                .MoveNext()
            Loop
        End With
    End Sub


    Private Sub UpdateIIn_ItemLast()
        Dim RSCkk As New ADODB.Recordset
        LoadSqlData(" select * from Ap_balance_6_col   where ac_code =  '" & (RSCIn_M.Fields("Ac_Code").Value) & "' ", RSCkk)
        With RSCkk
            Do Until .EOF = True
                CNN.Execute("Insert into Ap_Rpt_BLS_Detail (Ac_Code_Parent , Rpt_Id , Ac_Code , Ac_Name  , Last_Amt_Dr , Last_Amt_Cr , Rpt_Type) values (  '" & CStr((RSCIn_M.Fields("Ac_Code").Value)) & "' , '" & CStr((RSCIn_M.Fields("Rpt_Id").Value)) & "' , '" & CStr((.Fields("Ac_Code").Value)) & "' , N'" & CStr((.Fields("Ac_Name").Value)) & "'  , " & CDbl((.Fields("Rem_dr").Value)) & " , " & CDbl((.Fields("Rem_cr").Value)) & " , 'In' )")
                CNN.Execute("update  Ap_Rpt_BLS_Item_Old set Amt_Dr  =  Amt_Dr+" & CDbl((.Fields("Rem_dr").Value)) & " , Amt_Cr  = Amt_Cr+" & CDbl((.Fields("Rem_cr").Value)) & "   where ac_code = '" & (RSCIn_M.Fields("ac_code").Value) & "' And  Rpt_Type = 'In' ")
                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub UpdateIIn()
        CNN.Execute("delete Ap_Rpt_BLS_Stock ")
        CNN.Execute(" insert into Ap_Rpt_BLS_Stock ( Rpt_ID , Last_Amt_Dr , Last_Amt_Cr, Amt_Dr , Amt_Cr)" & _
                     "  select Rpt_ID , sum(Last_Amt_Dr) As Last_Amt_Dr , sum(Last_Amt_Cr) As Last_Amt_Cr ,sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr   from Ap_Rpt_BLS_Item_Old  where  Rpt_Type = 'In' " & CRLD & " group by Rpt_ID")
        CNN.Execute("Update Ap_Rpt_BLS_Old set Amt = Ap_Rpt_BLS_Stock.Amt_Dr-Ap_Rpt_BLS_Stock.Amt_cr ,Last_Amt =Ap_Rpt_BLS_Stock.Last_Amt_dr-Ap_Rpt_BLS_Stock.Last_Amt_Cr  from Ap_Rpt_BLS_Old ,Ap_Rpt_BLS_Stock where  Ap_Rpt_BLS_Old.Rpt_ID=Ap_Rpt_BLS_Stock.Rpt_ID")
    End Sub
    Private Sub UpdateIInLast()
        Dim RSC As New ADODB.Recordset
        LoadSqlData("select Rpt_ID, sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr  from Ap_Rpt_BLS_Item_Old  where  Rpt_Type = 'In' group by Rpt_ID ", RSC)
        With RSC
            Do Until .EOF = True
                CNN.Execute("Update Ap_Rpt_BLS_Old set " & _
                            " Last_Amt ='" & CDbl(CDbl((.Fields("Amt_dr").Value)) - CDbl((.Fields("Amt_cr").Value))) & "' " & _
                               " where Rpt_ID = '" & (.Fields("Rpt_ID").Value) & "' ")
                .MoveNext()
            Loop
        End With
    End Sub


    Private Sub SelectOut()
        '     If CMB_Curr.SelectedIndex = 2 Then
        '         CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
        '"where left(Ap_Rpt_BLS_Item_Old.ac_code,7) =left(Ap_balance_6_col.ac_code,7) And   Rpt_Type = 'Out'   ")
        '         '            CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
        '         '"where left(Ap_Rpt_BLS_Item_Old.ac_code,7) =left(Ap_balance_6_col.ac_code,7) And  Rpt_Type = 'In'  and left(Ap_balance_6_col.ac_code,7)='2382120'  ")
        '         '            CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
        '         '    "where left(Ap_Rpt_BLS_Item_Old.ac_code,7) = left(Ap_balance_6_col.ac_code,7) And  Rpt_Type = 'In'  ")
        '     ElseIf CMB_Curr.SelectedIndex = 0 Then
        '         'CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
        '         '   "where left(Ap_Rpt_BLS_Item_Old.ac_code,7) =left(Ap_balance_6_col.ac_code,7) And  Rpt_Type = 'Out' and left(Ap_balance_6_col.ac_code,7)<>'2382120' ")
        '         CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
        ' "where left(Ap_Rpt_BLS_Item_Old.ac_code,7) =left(Ap_balance_6_col.ac_code,7) And  Rpt_Type = 'Out'  ")
        '     Else
        '         CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
        '              "where left(Ap_Rpt_BLS_Item_Old.ac_code,7) =left(Ap_balance_6_col.ac_code,7) And  Rpt_Type = 'Out'   and left(Ap_balance_6_col.ac_code,7)<>'2382120' ")

        '     End If

        If CMB_Curr.SelectedIndex = 0 Then
            CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
   "where Ap_Rpt_BLS_Item_Old.ac_code = Ap_balance_6_col.ac_code And   Rpt_Type = 'Out'  " & CRLD & "    ")

        ElseIf CMB_Curr.SelectedIndex = 1 Then
            'CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
            '   "where left(Ap_Rpt_BLS_Item_Old.ac_code,7) =left(Ap_balance_6_col.ac_code,7) And  Rpt_Type = 'Out' and left(Ap_balance_6_col.ac_code,7)<>'2382120' ")
            CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
    "where Ap_Rpt_BLS_Item_Old.ac_code = Ap_balance_6_col.ac_code And  Rpt_Type = 'Out' and left(Ap_balance_6_col.ac_code,7)<>'2382120' " & CRLD & " ")
        Else
            CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
                 "where Ap_Rpt_BLS_Item_Old.ac_code = Ap_balance_6_col.ac_code And  Rpt_Type = 'Out'  " & CRLD & "   ")

        End If



        CNN.Execute("Insert into Ap_Rpt_BLS_Detail (  Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type )" & _
         " select   Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type from Ap_Rpt_BLS_Item_Old where Rpt_Type = 'Out' And ( Amt_Dr <>0 or Amt_Cr <>0  or Last_Amt_Dr <>0 or Last_Amt_Cr <>0 ) " & CRLD & "")

        'LoadSqlData("select * from Ap_Rpt_BLS_Item_Old where  Rpt_Type = 'Out' ", RSCIn_M)
        'With RSCIn_M
        '    Do Until .EOF = True
        '        Call UpdateOut_Item()
        '        .MoveNext()
        '    Loop
        'End With
        'If RSCIn_M.State = ConnectionState.Open Then RSCIn_M.Close()
    End Sub
    Private Sub SelectOutLast()

        LoadSqlData("select * from Ap_Rpt_BLS_Item_Old where  Rpt_Type = 'Out' ", RSCIn_M)
        With RSCIn_M
            Do Until .EOF = True
                Call UpdateOut_ItemLast()
                .MoveNext()
            Loop
        End With
        'If RSCIn_M.State = ConnectionState.Open Then RSCIn_M.Close()
    End Sub

    Private Sub UpdateOut_Item()
        Dim RSCkk As New ADODB.Recordset
        LoadSqlData(" select * from Ap_balance_6_col   where ac_code =  '" & (RSCIn_M.Fields("Ac_Code").Value) & "' ", RSCkk)
        With RSCkk
            Do Until .EOF = True
                CNN.Execute("Insert into Ap_Rpt_BLS_Detail (Ac_Code_Parent , Rpt_Id , Ac_Code , Ac_Name  ,  Last_Amt_Dr , Last_Amt_Cr, Amt_Dr , Amt_Cr , Rpt_Type ) values (  '" & CStr((RSCIn_M.Fields("Ac_Code").Value)) & "' , '" & CStr((RSCIn_M.Fields("Rpt_Id").Value)) & "' , '" & CStr((.Fields("Ac_Code").Value)) & "' , N'" & CStr((.Fields("Ac_Name").Value)) & "'  ,   " & CDbl((.Fields("Open_Amt_dr").Value)) & " , " & CDbl((.Fields("Open_Amt_cr").Value)) & " , " & CDbl((.Fields("Rem_dr").Value)) & " , " & CDbl((.Fields("Rem_cr").Value)) & " , 'Out' )")
                CNN.Execute("update  Ap_Rpt_BLS_Item_Old set Last_Amt_Dr  =  Last_Amt_Dr+" & CDbl((.Fields("Open_Amt_dr").Value)) & " , Last_Amt_Cr  = Last_Amt_Cr+" & CDbl((.Fields("Open_Amt_cr").Value)) & "  , Amt_Dr  =  Amt_Dr+" & CDbl((.Fields("Rem_dr").Value)) & " , Amt_Cr  = Amt_Cr+" & CDbl((.Fields("Rem_cr").Value)) & "   where ac_code = '" & (RSCIn_M.Fields("ac_code").Value) & "' And  Rpt_Type = 'Out' ")

                .MoveNext()
            Loop
        End With

    End Sub

    Private Sub UpdateOut_ItemLast()
        Dim RSCkk As New ADODB.Recordset
        LoadSqlData(" select * from Ap_balance_6_col   where ac_code =  '" & (RSCIn_M.Fields("Ac_Code").Value) & "' ", RSCkk)
        With RSCkk
            Do Until .EOF = True
                CNN.Execute("Insert into Ap_Rpt_BLS_Detail (Ac_Code_Parent , Rpt_Id , Ac_Code , Ac_Name  , Last_Amt_Dr , Last_Amt_Cr , Rpt_Type) values (  '" & CStr((RSCIn_M.Fields("Ac_Code").Value)) & "' , '" & CStr((RSCIn_M.Fields("Rpt_Id").Value)) & "' , '" & CStr((.Fields("Ac_Code").Value)) & "' , N'" & CStr((.Fields("Ac_Name").Value)) & "'  , " & CDbl((.Fields("Rem_dr").Value)) & " , " & CDbl((.Fields("Rem_cr").Value)) & ", 'Out')")
                CNN.Execute("update  Ap_Rpt_BLS_Item_Old set Amt_Dr  =  Amt_Dr+" & CDbl((.Fields("Rem_dr").Value)) & " , Amt_Cr  = Amt_Cr+" & CDbl((.Fields("Rem_cr").Value)) & "   where ac_code = '" & (RSCIn_M.Fields("ac_code").Value) & "' And  Rpt_Type = 'Out' ")
                .MoveNext()
            Loop
        End With

    End Sub



    'Private Sub UpdateOutLast()
    '    Dim RSC As New ADODB.Recordset
    '    LoadSqlData("select Rpt_ID, sum(Last_Amt_Dr) As Last_Amt_Dr , sum(Last_Amt_Cr) As Last_Amt_Cr  from Ap_Rpt_BLS_Item_Old  where  Rpt_Type = 'Out' group by Rpt_ID ", RSC)
    '    With RSC
    '        Do Until .EOF = True
    '            CNN.Execute("Update Ap_Rpt_BLS set " & _
    '                     " Last_Amt ='" & CDbl(CDbl((.Fields("Last_Amt_cr").Value)) - CDbl((.Fields("Last_Amt_dr").Value))) & "' " & _
    '                        " where Rpt_ID = '" & (.Fields("Rpt_ID").Value) & "' ")
    '            .MoveNext()
    '        Loop
    '    End With
    'End Sub

    Private Sub UpdateOut()
        CNN.Execute("delete Ap_Rpt_BLS_Stock ")
        CNN.Execute(" insert into Ap_Rpt_BLS_Stock ( Rpt_ID , Last_Amt_Dr , Last_Amt_Cr, Amt_Dr , Amt_Cr)" & _
                     "  select Rpt_ID , sum(Last_Amt_Dr) As Last_Amt_Dr , sum(Last_Amt_Cr) As Last_Amt_Cr ,sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr   from Ap_Rpt_BLS_Item_Old  where  Rpt_Type = 'Out' " & CRLD & " group by Rpt_ID  having (sum(Last_Amt_Dr)+sum(Last_Amt_Cr)+sum(Amt_Dr)+sum(Amt_Cr)) <>0 ")
        CNN.Execute("Update Ap_Rpt_BLS_Old set Amt = Ap_Rpt_BLS_Stock.Amt_Dr-Ap_Rpt_BLS_Stock.Amt_Dr ,Last_Amt =Ap_Rpt_BLS_Stock.Last_Amt_dr-Ap_Rpt_BLS_Stock.Last_Amt_Cr  from Ap_Rpt_BLS_Old ,Ap_Rpt_BLS_Stock where  Ap_Rpt_BLS_Old.Rpt_ID=Ap_Rpt_BLS_Stock.Rpt_ID " & CRLD & "")
        CNN.Execute("Update Ap_Rpt_BLS_Old set Amt = Ap_Rpt_BLS_Stock.Amt_Cr-Ap_Rpt_BLS_Stock.Amt_Dr ,Last_Amt =Ap_Rpt_BLS_Stock.Last_Amt_Cr-Ap_Rpt_BLS_Stock.Last_Amt_Dr  from Ap_Rpt_BLS_Old ,Ap_Rpt_BLS_Stock where  Ap_Rpt_BLS_Old.Rpt_ID=Ap_Rpt_BLS_Stock.Rpt_ID " & CRLD & "")
        Dim RSC As New ADODB.Recordset
        'LoadSqlData("select Rpt_ID, sum(Last_Amt_dr) As Last_Amt_Dr , sum(Last_Amt_cr) As Last_Amt_cr , sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr  from Ap_Rpt_BLS_Item_Old  where  Rpt_Type = 'Out' group by Rpt_ID ", RSC)
        'With RSC
        '    Do Until .EOF = True
        '        CNN.Execute("Update Ap_Rpt_BLS set " & _
        '                 " Amt ='" & CDbl(CDbl((.Fields("Amt_cr").Value)) - CDbl((.Fields("Amt_dr").Value))) & "' " & _
        '                   " ,Last_Amt ='" & CDbl(CDbl((.Fields("Last_Amt_cr").Value)) - CDbl((.Fields("Last_Amt_dr").Value))) & "' " & _
        '                    " where Rpt_ID = '" & (.Fields("Rpt_ID").Value) & "' ")
        '        .MoveNext()
        '    Loop
        'End With
    End Sub


    Private Sub UpdateOutLast()
        Dim RSC As New ADODB.Recordset
        LoadSqlData("select Rpt_ID, sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr  from Ap_Rpt_BLS_Item_Old  where  Rpt_Type = 'Out' group by Rpt_ID ", RSC)
        With RSC
            Do Until .EOF = True
                CNN.Execute("Update Ap_Rpt_BLS_Old set " & _
                         " Last_Amt ='" & CDbl(CDbl((.Fields("Amt_cr").Value)) - CDbl((.Fields("Amt_dr").Value))) & "' " & _
                            " where Rpt_ID = '" & (.Fields("Rpt_ID").Value) & "' ")
                .MoveNext()
            Loop
        End With
    End Sub
    Private Sub Update_Sum()
        CNN.Execute("update Ap_Rpt_BLS_Detail set  Rpt_Name=Ap_Rpt_BLS_Old.Description from   Ap_Rpt_BLS_Detail , Ap_Rpt_BLS_Old  where Ap_Rpt_BLS_Detail.Rpt_Id = Ap_Rpt_BLS_Old.Rpt_Id")
        CNN.Execute(" Update Caculate_Rpt set  CLT_Amt  = CLT_Str ,  CLT_Last_Amt  = CLT_Str where CLT_Str = '+' or CLT_Str = '-' or CLT_Str = '*' or CLT_Str = '+' or CLT_Str = '/' or CLT_Str = '(' or CLT_Str=')' Or CLT_Str<>'Cast(('   Or CLT_Str<>')As Float)'")
        CNN.Execute("delete Caculate_Lock")
        CNN.Execute("delete Caculate_Start")
        CNN.Execute(" Insert Into Caculate_Start (Rpt_Id , Rpt_Type , clt_Str , clt_Amt , Clt_Last_Amt ) select Rpt_Id , Rpt_Type , clt_Str , clt_Amt , Clt_Last_Amt from Caculate_Rpt where Rpt_Type = 'BLS'  Order by  Rpt_id ,cnt asc  ")
        CNN.Execute("update Caculate_Start set lck =0")
        CNN.Execute("Insert into Caculate_Lock (cnt_Mt)  SELECT  (SELECT     TOP 1 cnt FROM Caculate_Start AS B WHERE(Rpt_Id = A.Rpt_Id   ) ORDER BY cnt desc) AS cnt FROM Caculate_Start  AS A  GROUP BY Rpt_Id ORDER BY Rpt_Id")
        CNN.Execute("update  Caculate_Start set lck=1 from Caculate_Start ,Caculate_Lock  where Caculate_Start.cnt=Caculate_Lock.cnt_MT")
        CNN.Execute("  Update Caculate_Start set Caculate_Start.Amt = Ap_Rpt_BLS_Old.Amt , Caculate_Start.Last_Amt = Ap_Rpt_BLS_Old.Last_Amt   from Caculate_Start , Ap_Rpt_BLS_Old  where  Caculate_Start.CLT_Str  = Ap_Rpt_BLS_Old.Rpt_Id  ")
        CNN.Execute("Update Caculate_Start set lck_Amt=0")
        CNN.Execute("Update Caculate_Start set lck_Amt=1 where CLT_Str <> '+' And CLT_Str <> '-' And CLT_Str <> '*' And CLT_Str <> '+' And CLT_Str <> '/' And CLT_Str <> '(' And CLT_Str<>')' And CLT_Str<>'Cast(('   And CLT_Str<>')As Float)' ")
        Dim RSC1 As New ADODB.Recordset
        CLT_Str = ""
        CLT_Last_Str = ""
        With RSC1
            Call LoadSqlData("select *  from Caculate_Start where Rpt_Type = 'BLS'  Order by  Rpt_id ,cnt asc", RSC1)
            If .RecordCount > 0 Then
                While Not .EOF()
                    If (RSC1.Fields("lck_Amt").Value.ToString) = "1" Then
                        CLT_Str = CLT_Str & (RSC1.Fields("Amt").Value.ToString)
                        CLT_Last_Str = CLT_Last_Str & (RSC1.Fields("Last_Amt").Value.ToString)
                    Else
                        CLT_Str = CLT_Str & (RSC1.Fields("CLT_Amt").Value.ToString)
                        CLT_Last_Str = CLT_Last_Str & (RSC1.Fields("CLT_Last_Amt").Value.ToString)
                    End If
                    If (RSC1.Fields("lck").Value.ToString) = "1" Then
                        On Error GoTo hang
hang:
                        If Err.Number = 0 Then
                            Dim s As String = " Update  Ap_Rpt_BLS_Old set Amt = " & CLT_Str & " , Last_Amt = " & CLT_Last_Str & " where  Rpt_ID =   '" & (RSC1.Fields("Rpt_ID").Value.ToString) & "'"
                            CNN.Execute(s)
                        Else
                            'MessageBox.Show("ສູດຄິດໄລ່ຂອງ " & (RSC1.Fields("Rpt_ID").Value.ToString) & " = " & CLT_Last_Str & " ບໍ່ຖຶກຕ້ອງກະລຸນນາກວດສອບຄືນໃຫມ່")
                            Exit Sub
                        End If
                        CLT_Str = ""
                        CLT_Last_Str = ""
                    End If
                    .MoveNext()
                End While
            End If
        End With

    End Sub



    Private Sub LoadOpen_Jn1()
        Dim RSC12 As New ADODB.Recordset


        Dim add As String = "select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & MULook & " group BY ac_code "

        'MsgBox(add)

        LoadSqlData(add, RSC12)
        With RSC12
            Do Until .EOF = True
                VCode1 = CStr(Trim(.Fields("ac_Code").Value))
                CNN.Execute("INSERT INTO Ap_balance_6_col( ac_code   , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
                 "Values('" & CStr(Trim(.Fields("ac_Code").Value)) & "', " & _
                 " " & CDbl(0) & ", " & CDbl(0) & ", " & CDbl(Trim(.Fields("amt_dr").Value)) & ", " & CDbl(Trim(.Fields("amt_cr").Value)) & ",0 )")

                .MoveNext()
            Loop
        End With
    End Sub









    Private Sub LoadOpen_Jn12()
        CNN.Execute("Update Ap_balance_6_col set Quarter_dr=0,Quarter_cr=0")

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

        LoadSqlData("   select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(DS, "yyyy-MM-dd") & "' AND '" & Format(DT, "yyyy-MM-dd") & "' " & MULook & " group BY ac_code ", RSC12)


        With RSC12
            Do Until .EOF = True
                Call LoadOpen_Jn14()
                .MoveNext()
            Loop
        End With

    End Sub

    Private Sub LoadOpen_Jn14()
        Dim RSC14 As New ADODB.Recordset
        LoadSqlData("   select * from Ap_balance_6  where ac_code = '" & CStr(Trim(RSC12.Fields("ac_code").Value)) & "'  ", RSC14)

        If RSC14.RecordCount <> 0 Then
            CNN.Execute("Update Ap_balance_6_col set Quarter_dr= " & CDbl(Trim(RSC12.Fields("amt_dr").Value)) & "  , Quarter_cr= " & CDbl(Trim(RSC12.Fields("amt_cr").Value)) & "  where ac_code = '" & CStr(Trim(RSC12.Fields("ac_code").Value)) & "'  ")
        Else
            MsgBox("ggg")
            CNN.Execute("INSERT INTO Ap_balance_6_col( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status, Quarter_dr , Quarter_cr  ) " & _
              "Values('" & CStr(Trim(RSC12.Fields("ac_Code").Value)) & "', " & _
              " 0 , 0 , 0 , 0,0, " & CDbl(Trim(RSC12.Fields("amt_dr").Value)) & " , " & CDbl(Trim(RSC12.Fields("amt_cr").Value)) & " )")

        End If
    End Sub
    Private Sub LoadOpen_Jn14_1()
        Dim RSC14_1 As New ADODB.Recordset
        LoadSqlData("select * from Ap_balance_6_col  ", RSC14_1)
        With RSC14_1
            Do Until .EOF = True
                If CDbl(CDbl((.Fields("open_amt_dr").Value)) + CDbl((.Fields("amt_dr").Value))) >= CDbl(CDbl((.Fields("open_amt_cr").Value)) + CDbl((.Fields("amt_cr").Value))) Then
                    CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr=" & CDbl(CDbl(CDbl((.Fields("open_amt_dr").Value)) + CDbl((.Fields("amt_dr").Value))) - CDbl(CDbl((.Fields("open_amt_cr").Value)) + CDbl((.Fields("amt_cr").Value)))) & " where Ac_Code = '" & (.Fields("Ac_Code").Value) & "'")
                Else
                    CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr=" & CDbl(CDbl(CDbl((.Fields("open_amt_cr").Value)) + CDbl((.Fields("amt_cr").Value))) - CDbl(CDbl((.Fields("open_amt_dr").Value)) + CDbl((.Fields("amt_dr").Value)))) & " where Ac_Code = '" & (.Fields("Ac_Code").Value) & "'")
                End If
                .MoveNext()
            Loop
        End With
    End Sub
    Private Sub LoadOpen_Jn2()
        Dim RSC12 As New ADODB.Recordset

        Dim S As Date
        S = MdStartDate
        S = DateAdd("d", CDbl(-1), MdStartDate)
        'LoadSqlData("SELECT GIN.ac_code, ACC.Name_L, ACC.Name_E, SUM(GIN.amount_dr)AS amount_dr , SUM(GIN.amount_cr)AS amount_cr , SUM(GIN.amt_dr) AS amt_dr, SUM(GIN.amt_cr) AS amt_cr FROM Acc_Code ACC INNER JOIN gen_jn GIN ON ACC.AC_CODE = GIN.ac_code WHERE  GIN.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' GROUP BY GIN.ac_code, ACC.Name_L, ACC.Name_E  Order by GIN.AC_Code DESC  ", RSC12)

        LoadSqlData("   select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "'  " & MULook & " group BY ac_code ", RSC12)


        With RSC12
            Do Until .EOF = True
                'VCode2 = (.Fields("ac_Code").Value)
                'MsgBox(VCode)
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
               "Values('" & CStr(Trim(.Fields("ac_Code").Value)) & "', " & _
               " " & CDbl(0) & ", " & CDbl(0) & ", " & CDbl(Trim(.Fields("amt_dr").Value)) & ", " & CDbl(Trim(.Fields("amt_cr").Value)) & ",0 )")

                'LoadOpen_Jn2()
                'LoadOpen_Jn3()
                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub LoadOpen_Jn3()
        Dim RSC3 As New ADODB.Recordset

        LoadSqlData("select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr from Open_jn WHERE    date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & MULook & " group BY ac_code", RSC3)


        With RSC3
            Do Until .EOF = True
                VCode3 = (.Fields("ac_Code").Value)
                'MsgBox((.Fields("amt_dr").Value) & "__" & (.Fields("amt_cr").Value))
                CNN.Execute("Update Ap_balance_6 set  open_amt_dr='" & CDbl((.Fields("amt_dr").Value)) & "' , open_amt_cr='" & CDbl((.Fields("amt_cr").Value)) & "' where ac_code = '" & (.Fields("ac_Code").Value) & "'")
                LoadOpen_Jn4()
                .MoveNext()
            Loop
        End With
    End Sub


    Private Sub LoadOpen_Jn4()
        Dim RSC4 As New ADODB.Recordset
        With RSC
            LoadSqlData("select * from Ap_balance_6  WHERE     ac_code='" & VCode3 & "'  ", RSC4)
            If RSC4.RecordCount > 0 Then
                VCode4 = (RSC4.Fields("ac_Code").Value)
                'MsgBox(VCode4)
                'CNN.Execute("Update Ap_balance_6 set  open_amt_dr=" & CDbl((RSC4.Fields("amount_dr").Value)) & " , open_amt_cr=" & CDbl((RSC4.Fields("amount_dr").Value)) & " where ac_code = '" & (.Fields("ac_code").Value) & "'")
            Else
                'VCode4 = (RSC4.Fields("ac_Code").Value)
                'MsgBox(VCode3 & "n")
                LoadOpen_Jn5()
            End If
        End With
    End Sub

    Private Sub LoadOpen_Jn5()
        Dim RSC5 As New ADODB.Recordset
        With RSC
            LoadSqlData("select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr from Open_jn  WHERE    ac_code='" & VCode3 & "' " & MULook & " group BY ac_code", RSC5)
            If RSC5.RecordCount > 0 Then
                'MsgBox(CStr(Trim(RSC5.Fields("amt_cr").Value)))
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
             "Values('" & CStr(Trim(RSC5.Fields("ac_Code").Value)) & "',  " & _
             " " & CDbl(RSC5.Fields("amt_dr").Value) & ", " & CDbl(RSC5.Fields("amt_cr").Value) & ", " & CDbl(0) & ", " & CDbl(0) & ",0 )")
            Else
                'VCode4 = CStr(Trim(.Fields("ac_Code").Value))
                'MsgBox(VCode3 & "n")
            End If
        End With
    End Sub



    Private Sub LoadOpen_Jn6()
        Dim RSC6 As New ADODB.Recordset
        Dim op_dr, op_cr, amt_dr, amt_cr As Double
        op_dr = 0
        op_cr = 0
        amt_dr = 0
        amt_cr = 0
        LoadSqlData("select Ac_Code , open_amt_dr , open_amt_cr , Amt_dr , Amt_cr from Ap_balance_6  ", RSC6)
        With RSC6
            Do Until .EOF = True
                op_dr = CDbl((.Fields("open_amt_dr").Value))
                op_cr = CDbl((.Fields("open_amt_cr").Value))
                amt_dr = CDbl((.Fields("Amt_dr").Value))
                amt_cr = CDbl((.Fields("Amt_cr").Value))

                If CDbl(op_dr + amt_dr) - CDbl(op_cr + amt_cr) >= 0 Then

                    'MsgBox((.Fields("open_amt_dr").Value) & "++++" & (.Fields("ac_code").Value))
                    CNN.Execute("Update Ap_balance_6 set rem_dr='" & CDbl(op_dr + amt_dr) - CDbl(op_cr + amt_cr) & "' , rem_cr='" & CDbl(0) & "' where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                    'MsgBox(.Fields("open_amt_dr").Value)
                End If
                If CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) >= 0 Then
                    CNN.Execute("Update Ap_balance_6 set rem_dr='" & CDbl(0) & "' , rem_cr='" & CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) & "' where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                End If
                'If CDbl(op_cr + amt_cr) = CDbl(op_dr + amt_dr) Then
                '    'CNN.Execute("Update Ap_balance_6 set rem_dr='" & CDbl(0) & "' , rem_cr='" & CDbl(0) & "' where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                '    CNN.Execute("delete Ap_balance_6_col  where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                'End If
                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub LoadOpen_Jn7()
        Dim RSC7 As New ADODB.Recordset
        LoadSqlData("select ac_code , rem_dr  , rem_cr from Ap_balance_6   ", RSC7)
        With RSC7
            Do Until .EOF = True
                VCode7 = (.Fields("ac_Code").Value)

                CNN.Execute("Update Ap_balance_6_col set  open_amt_dr='" & CDbl((.Fields("rem_dr").Value)) & "' , open_amt_cr='" & CDbl((.Fields("rem_cr").Value)) & "' where ac_code = '" & (.Fields("ac_Code").Value) & "'")
                LoadOpen_Jn8()
                .MoveNext()
            Loop
        End With
    End Sub


    Private Sub LoadOpen_Jn8()
        Dim RSC8 As New ADODB.Recordset
        With RSC
            LoadSqlData("select * from Ap_balance_6_col  WHERE     ac_code='" & VCode7 & "' ", RSC8)
            If RSC8.RecordCount > 0 Then
                VCode8 = (RSC8.Fields("ac_Code").Value)
            Else
                LoadOpen_Jn9()
            End If
        End With
    End Sub


    Private Sub LoadOpen_Jn9()
        Dim RSC9 As New ADODB.Recordset
        With RSC9
            LoadSqlData("select ac_code , Rem_dr , Rem_cr from Ap_balance_6  WHERE    ac_code='" & VCode7 & "' ", RSC9)
            If RSC9.RecordCount > 0 Then
                CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code ,ac_name , ac_namee , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
                "Values('" & CStr(Trim(RSC9.Fields("ac_Code").Value)) & "', N'" & "***" & "', '" & "***" & "', " & _
                " " & CDbl(RSC9.Fields("rem_dr").Value) & ", " & CDbl(RSC9.Fields("rem_cr").Value) & ", " & CDbl(0) & ", " & CDbl(0) & ",0 )")
            Else
            End If
        End With
    End Sub
    Private Sub LoadOpen_Jn15()
        CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")

        'Dim RSC10 As New ADODB.Recordset
        'AmtOpenDR = 0
        'AmtOpenCR = 0
        'LoadSqlData("select Ac_Code , Name_L , Name_E , Acc_TypeE from Acc_Code  ", RSC10)
        'With RSC10
        '    Do Until .EOF = True
        '        CNN.Execute("Update Ap_balance_6_col set ac_name = N'" & (.Fields("Name_L").Value) & "'   where ac_code='" & (.Fields("ac_code").Value) & "'")
        '        .MoveNext()
        '    Loop
        'End With
    End Sub

    Private Sub LoadOpen_Jn11()
        Dim RSC11 As New ADODB.Recordset
        Dim op_dr11, op_cr11, amt_dr11, amt_cr11 As Double
        op_dr11 = 0
        op_cr11 = 0
        amt_dr11 = 0
        amt_cr11 = 0
        LoadSqlData("select Ac_Code , open_amt_dr , open_amt_cr , Amt_dr , Amt_cr from Ap_balance_6_col  ", RSC11)
        With RSC11
            Do Until .EOF = True
                op_dr11 = CDbl((.Fields("open_amt_dr").Value))
                op_cr11 = CDbl((.Fields("open_amt_cr").Value))
                amt_dr11 = CDbl((.Fields("Amt_dr").Value))
                amt_cr11 = CDbl((.Fields("Amt_cr").Value))
                If CDbl(op_dr11 + op_cr11) = 0 Then
                    If CDbl(amt_dr11 + amt_cr11) = 0 Then
                        CNN.Execute("delete Ap_balance_6_col  where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                    End If

                End If
                .MoveNext()
            Loop
        End With
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
            MdStartDate_MM = Format(CDate("1/1/" & Year(yyt.Value) - 1), "dd-MM-yyyy")
            MdToDate_MM = Format(CDate("30/6/" & Year(yyt.Value) - 1), "dd-MM-yyyy")
        Else
            MdStartDate = Format(CDate("1/7/" & Year(yyt.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/12/" & Year(yyt.Value)), "dd-MM-yyyy")
            MdStartDate_MM = Format(CDate("1/7/" & Year(yyt.Value) - 1), "dd-MM-yyyy")
            MdToDate_MM = Format(CDate("31/12/" & Year(yyt.Value) - 1), "dd-MM-yyyy")
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
            Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາວັນທີ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        End If
        'MdStartDate = Format(CDate(Ds.Value), "dd-MM-yyyy")
        'MdToDate = Format(CDate(Dt.Value), "dd-MM-yyyy")
        'Lb.Text = "ແຕ່ວັນທີ " & MdStartDate & " ຫາວັນທີ " & MdToDate
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳວັນທີ"
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub
    Dim INNM, LastNM As String
    Private Sub LoadMonth()
        '---------------------------------
        If FmMain.MnLaoLang.Checked = True Then
            If DMonth.Text = "ມັງກອນ" Then
                MdStartDate = Format(CDate("01/01/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/01/" & Year(Myy.Value)), "dd-MM-yyyy")

                MdStartDate_MM = Format(CDate("01/01/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/01/" & Year(Myy.Value) - 1), "dd-MM-yyyy")

                MonthLetter1 = "ມັງກອນ"
                MonthLetter_Last = "ທັນວາ"
                DMonth.SelectedIndex = 0
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)

                INNM = "ທ້າຍເດືອນ " & Format(CDate(MdStartDate), "MM/yyyy")
                LastNM = "ທ້າຍປີ " & Format(CDate(MdStartDate_MM), "yyyy")
            ElseIf DMonth.Text = "ກຸມພາ" Then
                Dim Day As String
                Dim MM As Date
                Dim Fromm As Date
                MdStartDate = Format(CDate("01/02/" & Year(Myy.Value)), "dd-MM-yyyy")
                Fromm = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
                MM = Format(CDate("01/3/" & Year(MdStartDate)), "dd-MM-yyyy")
                Day = DateDiff(DateInterval.Day, Fromm, MM)
                MdToDate = Format(CDate(Day & "/02" & "/" & Year(Myy.Value)), "dd-MM-yyyy")

                MdStartDate_MM = Format(CDate("01/02/" & Year(MdStartDate) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate(Day & "/02" & "/" & Year(MdStartDate) - 1), "dd-MM-yyyy")

                MonthLetter1 = "ກຸມພາ"
                MonthLetter_Last = "ມັງກອນ"
                DMonth.SelectedIndex = 1
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
                INNM = "ທ້າຍເດືອນ " & Format(CDate(MdStartDate), "MM/yyyy")
                LastNM = "ທ້າຍເດືອນ " & Format(CDate(dpMonthPrev.Value), "MM/yyyy")

            ElseIf DMonth.Text = "ມີນາ" Then
                MdStartDate = Format(CDate("01/03/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/03/" & Year(Myy.Value)), "dd-MM-yyyy")

                MdStartDate_MM = Format(CDate("01/03/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/03/" & Year(Myy.Value) - 1), "dd-MM-yyyy")

                MonthLetter1 = "ມີນາ"
                MonthLetter_Last = "ກຸມພາ"
                DMonth.SelectedIndex = 2
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                INNM = "ທ້າຍເດືອນ " & Format(CDate(MdStartDate), "MM/yyyy")
                LastNM = "ທ້າຍເດືອນ " & Format(CDate(dpMonthPrev.Value), "MM/yyyy")
            ElseIf DMonth.Text = "ເມສາ" Then
                MdStartDate = Format(CDate("01/04/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/04/" & Year(Myy.Value)), "dd-MM-yyyy")

                MdStartDate_MM = Format(CDate("01/04/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("30/04/" & Year(Myy.Value) - 1), "dd-MM-yyyy")

                MonthLetter1 = "ເມສາ"
                MonthLetter_Last = "ມີນາ"
                DMonth.SelectedIndex = 3
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                INNM = "ທ້າຍເດືອນ " & Format(CDate(MdStartDate), "MM/yyyy")
                LastNM = "ທ້າຍເດືອນ " & Format(CDate(dpMonthPrev.Value), "MM/yyyy")
            ElseIf DMonth.Text = "ພຶດສະພາ" Then
                MdStartDate = Format(CDate("01/05/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/05/" & Year(Myy.Value)), "dd-MM-yyyy")

                MdStartDate_MM = Format(CDate("01/05/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/05/" & Year(Myy.Value) - 1), "dd-MM-yyyy")

                MonthLetter1 = "ພຶດສະພາ"
                MonthLetter_Last = "ເມສາ"
                DMonth.SelectedIndex = 4
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                INNM = "ທ້າຍເດືອນ " & Format(CDate(MdStartDate), "MM/yyyy")
                LastNM = "ທ້າຍເດືອນ " & Format(CDate(dpMonthPrev.Value), "MM/yyyy")
            ElseIf DMonth.Text = "ມິຖຸນາ" Then
                MdStartDate = Format(CDate("01/06/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/06/" & Year(Myy.Value)), "dd-MM-yyyy")


                MdStartDate_MM = Format(CDate("01/06/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("30/06/" & Year(Myy.Value) - 1), "dd-MM-yyyy")

                MonthLetter1 = "ມິຖຸນາ"
                MonthLetter_Last = "ພຶດສະພາ"
                DMonth.SelectedIndex = 5
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                INNM = "ທ້າຍເດືອນ " & Format(CDate(MdStartDate), "MM/yyyy")
                LastNM = "ທ້າຍເດືອນ " & Format(CDate(dpMonthPrev.Value), "MM/yyyy")
            ElseIf DMonth.Text = "ກໍລະກົດ" Then
                MdStartDate = Format(CDate("01/07/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/07/" & Year(Myy.Value)), "dd-MM-yyyy")

                MdStartDate_MM = Format(CDate("01/07/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/07/" & Year(Myy.Value) - 1), "dd-MM-yyyy")

                MonthLetter1 = "ກໍລະກົດ"
                MonthLetter_Last = "ມີຖຸນາ"
                DMonth.SelectedIndex = 6
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                INNM = "ທ້າຍເດືອນ " & Format(CDate(MdStartDate), "MM/yyyy")
                LastNM = "ທ້າຍເດືອນ " & Format(CDate(dpMonthPrev.Value), "MM/yyyy")
            ElseIf DMonth.Text = "ສິງຫາ" Then
                MdStartDate = Format(CDate("01/08/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/08/" & Year(Myy.Value)), "dd-MM-yyyy")

                MdStartDate_MM = Format(CDate("01/08/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/08/" & Year(Myy.Value) - 1), "dd-MM-yyyy")

                MonthLetter1 = "ສິງຫາ"
                MonthLetter_Last = "ກໍລະກົດ"
                DMonth.SelectedIndex = 7
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                INNM = "ທ້າຍເດືອນ " & Format(CDate(MdStartDate), "MM/yyyy")
                LastNM = "ທ້າຍເດືອນ " & Format(CDate(dpMonthPrev.Value), "MM/yyyy")
            ElseIf DMonth.Text = "ກັນຍາ" Then
                MdStartDate = Format(CDate("01/09/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/09/" & Year(Myy.Value)), "dd-MM-yyyy")

                MdStartDate_MM = Format(CDate("01/09/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("30/09/" & Year(Myy.Value) - 1), "dd-MM-yyyy")

                MonthLetter1 = "ກັນຍາ"
                MonthLetter_Last = "ສິງຫາ"
                DMonth.SelectedIndex = 8
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                INNM = "ທ້າຍເດືອນ " & Format(CDate(MdStartDate), "MM/yyyy")
                LastNM = "ທ້າຍເດືອນ " & Format(CDate(dpMonthPrev.Value), "MM/yyyy")
            ElseIf DMonth.Text = "ຕຸລາ" Then
                MdStartDate = Format(CDate("01/10/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/10/" & Year(Myy.Value)), "dd-MM-yyyy")

                MdStartDate_MM = Format(CDate("01/10/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/10/" & Year(Myy.Value) - 1), "dd-MM-yyyy")

                MonthLetter1 = "ຕຸລາ"
                MonthLetter_Last = "ກັນຍາ"
                DMonth.SelectedIndex = 9
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                INNM = "ທ້າຍເດືອນ " & Format(CDate(MdStartDate), "MM/yyyy")
                LastNM = "ທ້າຍເດືອນ " & Format(CDate(dpMonthPrev.Value), "MM/yyyy")
            ElseIf DMonth.Text = "ພະຈິກ" Then
                MdStartDate = Format(CDate("01/11/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/11/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdStartDate_MM = Format(CDate("01/11/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("30/11/" & Year(Myy.Value) - 1), "dd-MM-yyyy")

                MonthLetter1 = "ພະຈິກ"
                MonthLetter_Last = "ຕຸລາ"
                DMonth.SelectedIndex = 10
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                INNM = "ທ້າຍເດືອນ " & Format(CDate(MdStartDate), "MM/yyyy")
                LastNM = "ທ້າຍເດືອນ " & Format(CDate(dpMonthPrev.Value), "MM/yyyy")
            ElseIf DMonth.Text = "ທັນວາ" Then
                MdStartDate = Format(CDate("01/12/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/12/" & Year(Myy.Value)), "dd-MM-yyyy")

                MdStartDate_MM = Format(CDate("01/12/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/12/" & Year(Myy.Value) - 1), "dd-MM-yyyy")

                MonthLetter1 = "ທັນວາ"
                MonthLetter_Last = "ພະຈິກ"
                DMonth.SelectedIndex = 11
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                INNM = "ທ້າຍເດືອນ " & Format(CDate(MdStartDate), "MM/yyyy")
                LastNM = "ທ້າຍເດືອນ " & Format(CDate(dpMonthPrev.Value), "MM/yyyy")
            End If

            'Month_Last = Format(dpMonthPrev.Value, "MM/yyyy")
            Month_Last = "[" & MonthLetter_Last & " " & Format(dpMonthPrev.Value, "yyyy") & "]"

            Month_IN = "[" & MonthLetter1 & " " & Format(MdToDate, "yyyy") & "]"
            Month_IN_MM = "[" & MonthLetter1 & " " & Format(MdToDate, "yyyy") - 1 & "]"

            ' Lb.Text = "ສຳລັບວັນທີ " & (MdToDate.Day) & " " & MonthLetter1 & " " & Year(MdToDate)
            Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
        Else

            If DMonth.Text = "January" Then
                MdStartDate = Format(CDate("01/01/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/01/" & Year(Myy.Value)), "dd-MM-yyyy")

                MdStartDate_MM = Format(CDate("01/01/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/01/" & Year(Myy.Value) - 1), "dd-MM-yyyy")

                MonthLetter1 = "January"
                MonthLetter_Last = "December"
                DMonth.SelectedIndex = 0
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "February" Then
                Dim Day As String
                Dim MM As Date
                Dim Fromm As Date
                MdStartDate = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
                MdStartDate_MM = Format(CDate("01/02/" & Year(MdStartDate) - 1), "dd-MM-yyyy")

                Fromm = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
                MM = Format(CDate("01/3/" & Year(MdStartDate)), "dd-MM-yyyy")
                Day = DateDiff(DateInterval.Day, Fromm, MM)
                MdToDate = Format(CDate(Day & "/02" & "/" & Year(MdStartDate)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate(Day & "/02" & "/" & Year(MdStartDate) - 1), "dd-MM-yyyy")

                MonthLetter1 = "February"
                MonthLetter_Last = "January"
                DMonth.SelectedIndex = 1
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
            ElseIf DMonth.Text = "March" Then
                MdStartDate = Format(CDate("01/03/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/03/" & Year(Myy.Value)), "dd-MM-yyyy")

                MdStartDate_MM = Format(CDate("01/03/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/03/" & Year(Myy.Value) - 1), "dd-MM-yyyy")

                MonthLetter1 = "March"
                MonthLetter_Last = "February"
                DMonth.SelectedIndex = 2
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "April" Then
                MdStartDate = Format(CDate("01/04/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/04/" & Year(Myy.Value)), "dd-MM-yyyy")

                MdStartDate_MM = Format(CDate("01/04/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("30/04/" & Year(Myy.Value) - 1), "dd-MM-yyyy")

                MonthLetter1 = "April"
                MonthLetter_Last = "March"
                MonthLetter_Last = DMonth.SelectedIndex = 3
                DMonth.SelectedIndex = 3
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "May" Then
                MdStartDate = Format(CDate("01/05/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/05/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdStartDate_MM = Format(CDate("01/05/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/05/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MonthLetter1 = "May"
                MonthLetter_Last = "April"
                DMonth.SelectedIndex = 4
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "June" Then
                MdStartDate = Format(CDate("01/06/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/06/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdStartDate_MM = Format(CDate("01/06/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("30/06/" & Year(Myy.Value) - 1), "dd-MM-yyyy")

                MonthLetter1 = "June"
                MonthLetter_Last = "May"
                DMonth.SelectedIndex = 5
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "July" Then
                MdStartDate = Format(CDate("01/07/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/07/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdStartDate_MM = Format(CDate("01/07/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/07/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MonthLetter1 = "July"
                MonthLetter_Last = "June"
                DMonth.SelectedIndex = 6
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "August" Then
                MdStartDate = Format(CDate("01/08/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/08/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdStartDate_MM = Format(CDate("01/08/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/08/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MonthLetter1 = "August"
                MonthLetter_Last = "July"
                DMonth.SelectedIndex = 7
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "September" Then
                MdStartDate = Format(CDate("01/09/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/09/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdStartDate_MM = Format(CDate("01/09/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("30/09/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MonthLetter1 = "September"
                MonthLetter_Last = "August"
                DMonth.SelectedIndex = 8
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "October" Then
                MdStartDate = Format(CDate("01/10/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/10/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdStartDate_MM = Format(CDate("01/10/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/10/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MonthLetter1 = "October"
                MonthLetter_Last = "September"
                DMonth.SelectedIndex = 9
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "November" Then
                MdStartDate = Format(CDate("01/11/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/11/" & Year(Myy.Value)), "dd-MM-yyyy")

                MdStartDate_MM = Format(CDate("01/11/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("30/11/" & Year(Myy.Value) - 1), "dd-MM-yyyy")

                MonthLetter1 = "November"
                MonthLetter_Last = "October"
                DMonth.SelectedIndex = 10
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "December" Then
                MdStartDate = Format(CDate("01/12/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/12/" & Year(Myy.Value)), "dd-MM-yyyy")

                MdStartDate_MM = Format(CDate("01/12/" & Year(Myy.Value) - 1), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/12/" & Year(Myy.Value) - 1), "dd-MM-yyyy")

                MonthLetter1 = "December"
                MonthLetter_Last = "November"
                DMonth.SelectedIndex = 11
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            End If

            'dpMonthPrev.Value = DateAdd("m", -1, MdToDate) 
            Month_Last = "[" & MonthLetter_Last & " " & Format(dpMonthPrev.Value, "yyyy") & "]"
            Month_IN = "[" & MonthLetter1 & " " & Format(MdToDate, "yyyy") & "]"
            Month_IN_MM = "[" & MonthLetter1 & " " & Format(MdToDate, "yyyy") - 1 & "]"
            Lb.Text = "For the Month Ended " & (MdToDate.Day) & " " & MonthLetter1 & " " & Year(MdToDate)
        End If

        '-----------------
        'Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)

        'dpMonthPrev.Value = DateAdd("m", -1, MdToDate) 
        'Month_Last = Format(dpMonthPrev.Value, "MM/yyyy") 
        'Month_IN = Format(MdToDate, "MM/yyyy")
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub

    Private Sub LoadPeriod()
        If Period.Text = "ງວດທີ 1" Then
            MdStartDate = Format(CDate("01/01/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/03/" & Year(Pyy.Value)), "dd-MM-yyyy")

            MdStartDate_MM = Format(CDate("01/01/" & Year(Pyy.Value) - 1), "dd-MM-yyyy")
            MdToDate_MM = Format(CDate("31/03/" & Year(Pyy.Value) - 1), "dd-MM-yyyy")

            Lb.Text = "ປະຈຳງວດ " & "1" & " ປີ " & Pyy.Text
            INNM = "ທ້າຍງວດ 1 " & Format(CDate(MdStartDate), "MM/yyyy")
            LastNM = "ທ້າຍງວດ 4 " & Format(CDate(MdStartDate_MM), "MM/yyyy")
        ElseIf Period.Text = "ງວດທີ 2" Then
            MdStartDate = Format(CDate("01/04/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/06/" & Year(Pyy.Value)), "dd-MM-yyyy")

            MdStartDate_MM = Format(CDate("01/04/" & Year(Pyy.Value) - 1), "dd-MM-yyyy")
            MdToDate_MM = Format(CDate("30/06/" & Year(Pyy.Value) - 1), "dd-MM-yyyy")

            Lb.Text = "ປະຈຳງວດ " & "2" & " ປີ " & Pyy.Text
            INNM = "ທ້າຍງວດ 2 " & Format(CDate(MdStartDate), "MM/yyyy")
            LastNM = "ທ້າຍງວດ 1 " & Format(CDate(MdStartDate), "MM/yyyy")
        ElseIf Period.Text = "ງວດທີ 3" Then
            MdStartDate = Format(CDate("01/07/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/09/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdStartDate_MM = Format(CDate("01/07/" & Year(Pyy.Value) - 1), "dd-MM-yyyy")
            MdToDate_MM = Format(CDate("30/09/" & Year(Pyy.Value) - 1), "dd-MM-yyyy")

            Lb.Text = "ປະຈຳງວດ " & "3" & " ປີ " & Pyy.Text
            INNM = "ທ້າຍງວດ 3 " & Format(CDate(MdStartDate), "MM/yyyy")
            LastNM = "ທ້າຍງວດ 2 " & Format(CDate(MdStartDate), "MM/yyyy")
        ElseIf Period.Text = "ງວດທີ 4" Then
            MdStartDate = Format(CDate("01/10/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/12/" & Year(Pyy.Value)), "dd-MM-yyyy")

            MdStartDate_MM = Format(CDate("01/10/" & Year(Pyy.Value) - 1), "dd-MM-yyyy")
            MdToDate_MM = Format(CDate("31/12/" & Year(Pyy.Value) - 1), "dd-MM-yyyy")

            Lb.Text = "ປະຈຳງວດ " & "4" & " ປີ " & Pyy.Text
            INNM = "ທ້າຍງວດ 4 " & Format(CDate(MdStartDate), "MM/yyyy")
            LastNM = "ທ້າຍງວດ 3 " & Format(CDate(MdStartDate), "MM/yyyy")
        End If
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳ" & Period.Text & " ປີ " & Year(MdToDate)
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub

    Private Sub LoadYear()
        MdStartDate = Format(CDate("01/1/" & Year(yy.Value)), "dd-MM-yyyy")
        MdToDate = Format(CDate("31/12/" & Year(yy.Value)), "dd-MM-yyyy")

        MdStartDate_MM = Format(CDate("01/1/" & Year(yy.Value) - 1), "dd-MM-yyyy")
        MdToDate_MM = Format(CDate("31/12/" & Year(yy.Value) - 1), "dd-MM-yyyy")

        Month_IN = "ປະຈຳປີ " & Format(MdToDate, "yyyy")
        Month_Last = "ປະຈຳປີ " & Year(yy.Value) - 1

        Lb.Text = "ປະຈຳປີ " & yy.Text
        'Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd/MM/yyyy") & " ສີ້ນສຸດວັນທີ " & Format(MdToDate, "dd/MM/yyyy")
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ສີ້ນສຸດວັນທີ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        L5.Text = MdStartDate & " => " & MdToDate
        INNM = "ທ້າຍປີ " & Format(CDate(MdStartDate), "yyyy")
        LastNM = "ທ້າຍປີ " & Format(CDate(MdStartDate_MM), "yyyy")
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳປີ  " & Year(MdToDate)
    End Sub

    Private Sub UpdateData2()
        Dim RSC2 As New ADODB.Recordset
        Dim op_dr, op_cr, amt_dr, amt_cr As Double
        op_dr = 0
        op_cr = 0
        amt_dr = 0
        amt_cr = 0
        LoadSqlData("select ac_code , sum(open_amt_dr) as open_amt_dr , sum(open_amt_cr) as open_amt_cr  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  from Ap_balance_6_col where ac_code ='" & bls1 & "'   group by  ac_code ", RSC2)
        With RSC2
            Do Until .EOF = True
                op_dr = CDbl((.Fields("open_amt_dr").Value))
                op_cr = CDbl((.Fields("open_amt_cr").Value))
                amt_dr = CDbl((.Fields("Amt_dr").Value))
                amt_cr = CDbl((.Fields("Amt_cr").Value))
                'MsgBox(amt_cr)
                If CDbl(op_dr + amt_dr) - CDbl(op_cr + amt_cr) >= 0 Then
                    CNN.Execute("Update Ap_Rpt_BLS_Item_Old set amt_dr='" & CDbl(op_dr + amt_dr) - CDbl(op_cr + amt_cr) & "' , amt_cr='" & CDbl(0) & "' where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                End If
                If CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) >= 0 Then
                    CNN.Execute("Update Ap_Rpt_BLS_Item_Old set amt_dr='" & CDbl(0) & "' , amt_cr='" & CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) & "' where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                End If
                'CNN.Execute("update Ap_Rpt_BLS_Item_Old set amt_dr  =  " & CDbl((.Fields("amt_dr").Value)) & " , amt_cr  = " & CDbl((.Fields("amt_cr").Value)) & "   where Ac_code=  '" & (.Fields("Ac_code").Value) & "' ")
                .MoveNext()
            Loop
        End With
    End Sub


    Private Sub LoadReport()
        Dim RPT_ID As String

        Dim AA As String = Lb.Text
        RPT_ID = " "
        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As RptSjUd ,"
        LngId = "7059" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_PP ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
        LngId = "7002" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDP ,"
        'LngId = "7004" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Descrip ,"
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

        If CMB_Curr.Text = "EQVL" Then
            If CheckBox4.Checked = True Then
                LngId = "7121" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"

            Else
                LngId = "7106" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"

            End If
        ElseIf CMB_Curr.Text = "LAK" Then
            LngId = "7097" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        ElseIf CMB_Curr.Text = "USD" Then
            LngId = "7108" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        End If


        'If CMB_Curr.SelectedIndex = 0 Then
        '    LngId = "7106" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        'ElseIf CMB_Curr.SelectedIndex = 1 Then
        '    LngId = "7107" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        'ElseIf CMB_Curr.SelectedIndex = 2 Then
        '    LngId = "7108" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        'End If

        If RD.Checked = True Then
            'LngId = "7090" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            'LngId = "7091" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"
            Month_IN = ""
            Month_Last = ""
            LngId = "7113" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
        ElseIf RM.Checked = True Then
            'LngId = "7060" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            'LngId = "7061" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"

            LngId = "7096" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
        ElseIf RP.Checked = True Then
            Month_IN = ""
            Month_Last = ""
            'LngId = "7083" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            'LngId = "7063" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7112" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
        ElseIf Rhalf.Checked = True Then
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"
            'LngId = "7060" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            'LngId = "7061" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7113" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
            Month_IN = ""
            Month_Last = ""
        ElseIf RT.Checked = True Then
            Month_IN = ""
            Month_Last = ""
            'LngId = "7080" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            'LngId = "7081" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"
            LngId = "7113" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
        ElseIf RY.Checked = True Then
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"
            'LngId = "7084" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            'LngId = "7065" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7113" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
        End If

        Dim CRLD As String
        If r1.Checked = True Then
            CRLD = "ລາຍການ"
            LngId = "7092" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & " " & DTDATE02 & "' As Crl_RptName ,"
            r = " "
            'r = "And GRP<4"
        ElseIf r2.Checked = True Then
            CRLD = "ລາຍການຊັບສິນ"
            LngId = "7093" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & " " & DTDATE02 & "' As Crl_RptName ,"
            'r = "And GRP<5"
            r = "And GRP=1"
        ElseIf r3.Checked = True Then
            CRLD = "ລາຍການໜີ້ສິນ"
            LngId = "7054" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & " " & DTDATE02 & "' As Crl_RptName ,"
            'r = "And GRP>5"
            'r = "And GRP>1 and GRP<=3"
            r = "And GRP=2"
        End If
        Dim Doc As Integer = 0
        If CheckBox5.Checked = True Then
            Doc = 1
        Else
            Doc = 0
        End If

        If r1.Checked = True Then
            '======================GRP=1=========================
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '01%' And GRP=1 ) where Rpt_ID='01' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '01%' And GRP=1 ) where Rpt_ID='01' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '02%' And GRP=1 ) where Rpt_ID='02' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '02%' And GRP=1 ) where Rpt_ID='02' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '04%' And GRP=1 ) where Rpt_ID='04' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '04%' And GRP=1 ) where Rpt_ID='04' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '08%' And GRP=1 ) where Rpt_ID='08' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '08%' And GRP=1 ) where Rpt_ID='08' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '10%' And GRP=1 ) where Rpt_ID='10' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '10%' And GRP=1 ) where Rpt_ID='10' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where ( Rpt_ID='01' or Rpt_ID='02' or Rpt_ID='03' or Rpt_ID='04' or Rpt_ID='05' or Rpt_ID='06' or Rpt_ID='07' or Rpt_ID='08' or Rpt_ID='09' or Rpt_ID='10' ) And GRP=1 ) where Rpt_ID='11' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where ( Rpt_ID='01' or Rpt_ID='02' or Rpt_ID='03' or Rpt_ID='04' or Rpt_ID='05' or Rpt_ID='06' or Rpt_ID='07' or Rpt_ID='08' or Rpt_ID='09' or Rpt_ID='10' ) And GRP=1 ) where Rpt_ID='11' And GRP=1 ")
            '======================GRP=2=========================
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '02%' And GRP=2 ) where Rpt_ID='02' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '02%' And GRP=2 ) where Rpt_ID='02' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '01%' And GRP=2 ) where Rpt_ID='01' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '01%' And GRP=2 ) where Rpt_ID='01' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where ( Rpt_ID='05.1' or Rpt_ID='05.2' or Rpt_ID='05.3' or Rpt_ID='05.4') And GRP=2 ) where Rpt_ID='05' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where ( Rpt_ID='05.1' or Rpt_ID='05.2' or Rpt_ID='05.3' or Rpt_ID='05.4') And GRP=2 ) where Rpt_ID='05' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where ( Rpt_ID='01' or Rpt_ID='02' or Rpt_ID='03' or Rpt_ID='04' or Rpt_ID='05' ) And GRP=2 ) where Rpt_ID='05.5' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where ( Rpt_ID='01' or Rpt_ID='02' or Rpt_ID='03' or Rpt_ID='04' or Rpt_ID='05' ) And GRP=2 ) where Rpt_ID='05.5' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '06%' And GRP=2 ) where Rpt_ID='06' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '06%' And GRP=2 ) where Rpt_ID='06' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where ( Rpt_ID='05.5' or Rpt_ID='06' ) And GRP=2 ) where Rpt_ID='07' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where ( Rpt_ID='05.5' or Rpt_ID='06' ) And GRP=2 ) where Rpt_ID='07' And GRP=2 ")

        ElseIf r2.Checked = True Then
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '01%' And GRP=1 ) where Rpt_ID='01' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '01%' And GRP=1 ) where Rpt_ID='01' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '02%' And GRP=1 ) where Rpt_ID='02' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '02%' And GRP=1 ) where Rpt_ID='02' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '04%' And GRP=1 ) where Rpt_ID='04' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '04%' And GRP=1 ) where Rpt_ID='04' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '08%' And GRP=1 ) where Rpt_ID='08' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '08%' And GRP=1 ) where Rpt_ID='08' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '10%' And GRP=1 ) where Rpt_ID='10' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '10%' And GRP=1 ) where Rpt_ID='10' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where ( Rpt_ID='01' or Rpt_ID='02' or Rpt_ID='03' or Rpt_ID='04' or Rpt_ID='05' or Rpt_ID='06' or Rpt_ID='07' or Rpt_ID='08' or Rpt_ID='09' or Rpt_ID='10' ) And GRP=1 ) where Rpt_ID='11' And GRP=1 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where ( Rpt_ID='01' or Rpt_ID='02' or Rpt_ID='03' or Rpt_ID='04' or Rpt_ID='05' or Rpt_ID='06' or Rpt_ID='07' or Rpt_ID='08' or Rpt_ID='09' or Rpt_ID='10' ) And GRP=1 ) where Rpt_ID='11' And GRP=1 ")

        ElseIf r3.Checked = True Then

            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '02%' And GRP=2 ) where Rpt_ID='02' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '02%' And GRP=2 ) where Rpt_ID='02' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '01%' And GRP=2 ) where Rpt_ID='01' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '01%' And GRP=2 ) where Rpt_ID='01' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where ( Rpt_ID='05.1' or Rpt_ID='05.2' or Rpt_ID='05.3' or Rpt_ID='05.4') And GRP=2 ) where Rpt_ID='05' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where ( Rpt_ID='05.1' or Rpt_ID='05.2' or Rpt_ID='05.3' or Rpt_ID='05.4') And GRP=2 ) where Rpt_ID='05' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where ( Rpt_ID='01' or Rpt_ID='02' or Rpt_ID='03' or Rpt_ID='04' or Rpt_ID='05' ) And GRP=2 ) where Rpt_ID='05.5' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where ( Rpt_ID='01' or Rpt_ID='02' or Rpt_ID='03' or Rpt_ID='04' or Rpt_ID='05' ) And GRP=2 ) where Rpt_ID='05.5' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '06%' And GRP=2 ) where Rpt_ID='06' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where Rpt_ID like '06%' And GRP=2 ) where Rpt_ID='06' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set amt=(select sum(Amt) from Ap_Rpt_BLS_Old where ( Rpt_ID='05.5' or Rpt_ID='06' ) And GRP=2 ) where Rpt_ID='07' And GRP=2 ")
            CNN.Execute(" update Ap_Rpt_BLS_Old set Last_Amt=(select sum(Last_Amt) from Ap_Rpt_BLS_Old where ( Rpt_ID='05.5' or Rpt_ID='06' ) And GRP=2 ) where Rpt_ID='07' And GRP=2 ") 
        End If
        
        'INNM = "ທ້າຍປີ " & Format(CDate(MdStartDate), "MM/yyyy")
        'LastNM = "ທ້າຍປີ " & Format(CDate(MdStartDate_MM), "MM-yyyy")

        'SLF = "SELECT  " & Doc & "  as Doc  , N'" & MuOffDep & "'  as RptSjoff_Dep  , " & mformat & "  as mformat  ,  " & MuLngRpt & "  *, N'" & Month_IN & "' as In_Mn, N'" & Month_Last & "' as Last_Mn   FROM Ap_Rpt_BLS_Old  "
        SLF = "SELECT  " & Doc & "  as Doc  , N'" & MuOffDep & "'  as RptSjoff_Dep  , " & mformat & "  as mformat  ,  " & MuLngRpt & "  * , N'" & INNM & "' As Crl_InMounth ,N'" & LastNM & "' As Crl_LastMounth , N'" & INNM & "' as In_Mn, N'" & LastNM & "' as Last_Mn, N'" & CRLD & "'  As Crl_Descrip   FROM Ap_Rpt_BLS_Old  "

        Call LoadLoGO()

        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            .Open(" " & SLF & " where grp<>'' " & RPT_ID & " " & r & " order by GRP, Rpt_Id   ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        Dim FrmPreview As New FmPreview : FrmClosing()
        'Dim Rpt As New CryRpt_BLS
        Dim Rpt As New Object
        If RM.Checked = True Then
            Rpt = New CryRpt_BLS_M
        Else
            Rpt = New CryRpt_BLS
        End If

        'Dim Rpt As New CryRpt_BLS_NNN
        If MdShowLOGO = 1 Then
            Rpt.Subreports(0).SetDataSource(RsLOGO)
        End If
        Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S1"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS1.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S2"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS2.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S3"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS3.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S4"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS4.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("txtprint_user"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = MUserName
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("pp"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtPP.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("Text2"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = Lb.Text
        'myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("Text6"), CrystalDecisions.CrystalReports.Engine.TextObject)
        'myText2.Text = MDRegister
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("Text6"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = CURR01

        Rpt.SetDataSource(Rs)
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.DisplayGroupTree = False
        FrmPreview.WindowState = FormWindowState.Maximized
        FrmPreview.Show()
        FrmPreview.Focus()
    End Sub
    Private Sub loadOffice_User()
        Off_Usr.Items.Clear()
        LoadSqlData("select sub_id , off_add2  from  Ap_office  Order by sub_id", RSC)
        With RSC
            Do Until .EOF = True
                Off_Usr.Items.Add((.Fields("sub_id").Value) & " " & (.Fields("off_add2").Value))
                .MoveNext()
            Loop
        End With
        Off_Usr.Text = FmLogin.Sub_Company.Text
    End Sub

    Dim CRLD As String
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
            CRLD = " "
        ElseIf r2.Checked = True Then
            CRLD = "And GRP=1"
        ElseIf r3.Checked = True Then
            CRLD = "And GRP=2"
        End If
        BLNEW()
        CNN.Execute("update  Ap_Rpt_BLS_Item_Old set Last_amt_dr  =  " & CDbl(0) & " , Last_amt_cr  = " & CDbl(0) & " ,amt_dr  =  " & CDbl(0) & " , amt_cr  = " & CDbl(0) & "  ")
        CNN.Execute("update Ap_Rpt_BLS_Old set Last_Amt = 0 , Amt  =  " & CDbl(0) & "  ")
        CNN.Execute("DELETE FROM Ap_Rpt_BLS_Detail ")
        SelcectIn()
        UpdateIIn()
        SelectOut()
        UpdateOut()
        Update_Sum()

        If r1.Checked = True Then
            LngId = "7052" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & " " & DTDATE02 & "' As Crl_RptName ,"
            r = "And GRP>0"
        ElseIf r2.Checked = True Then
            LngId = "7053" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & " " & DTDATE02 & "' As Crl_RptName ,"
            r = "And GRP<5"
        ElseIf r3.Checked = True Then
            LngId = "7054" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & " " & DTDATE02 & "' As Crl_RptName ,"
            r = "And GRP>5"
        End If
        SLF = "SELECT     N'" & MuOffDep & "'  as RptSjoff_Dep  , " & mformat & "  as mformat  ,   " & MuLngRpt & "  *   FROM Ap_Rpt_BLS_Detail  "
        Call LoadLoGO()
        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            .Open(SLF, CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

            '.Open("SELECT *  ,N'" & txtReport_name & "' as txtReport_name  FROM Ap_Rpt_BLS where Amount_in_million_Kip <>0  ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim Rpt As New CryRpt_BLSItem
        If MdShowLOGO = 1 Then
            Rpt.Subreports(0).SetDataSource(RsLOGO)
        End If
        Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S1"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS1.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S2"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS2.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S3"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS3.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S4"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS4.Text

        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("pp"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtPP.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("Text2"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = Lb.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("Text6"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = CURR01

        Rpt.SetDataSource(Rs)
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
        CMB_Curr_SelectedIndexChanged(sender, e)
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
        'Toyy.Value = yy.Value
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

        Call HeaDer()

        Ds.Text = MWorkSetting
        Myy.Text = MWorkSetting
        yy.Text = MWorkSetting
        Toyy.Text = MWorkSetting
        'MsgBox(MWorkSetting)
        Pyy.Text = MWorkSetting
        MdToDate = MWorkSetting
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
        Call selectMMM()
        Call selectLoad()
        Call Click_Last()
        SetControlText(Me)
        Call loadOffice_User()
        If MDBSL = "1" Then
            r1.Checked = True
        ElseIf MDBSL = "2" Then
            r2.Checked = True
        ElseIf MDBSL = "3" Then
            r3.Checked = True
        End If
        CMB_Curr.Items.Clear()
        CMB_Curr.Items.Add("EQVL")
        Call load_Cmb(" SELECT Curr  FROM Curr_For_Rate WHERE (Curr='LAK' Or Curr='THB'  Or Curr='USD')  ORDER BY cnt ", "Curr", CMB_Curr)
        If CMB_Curr.Items.Count > 0 Then
            CMB_Curr.SelectedIndex = 0
        End If
        'If MuLng = "L" Then
        '    Button2.Text = "ທຽບບັນຊີ"
        '    CheckBox4.Text = "ທຽບເທົ່າເງິນ"

        'Else
        '    Button2.Text = "EQVL ACC"
        '    CheckBox4.Text = "EQVL Money"
        'End If
        'If CMB_Curr.Text = "LAK" Then
        '    CheckBox4.Text = "ທຽບເທົ່າໂດລາ"
        'ElseIf CMB_Curr.Text = "USD" Then
        '    CheckBox4.Text = "ທຽບເທົ່າກີບ"
        'Else
        '    CheckBox4.Text = "ທຽບເທົ່າໂດລາ"
        'End If
        If MuLng = "L" Then
            Button2.Text = "ທຽບບັນຊີ"
            CheckBox4.Text = "ທຽບເທົ່າເງິນ"
            If CMB_Curr.Text = "LAK" Then
                CheckBox4.Text = "ທຽບເທົ່າໂດລາ"
            ElseIf CMB_Curr.Text = "USD" Then
                CheckBox4.Text = "ທຽບເທົ່າກີບ"
            Else
                CheckBox4.Text = "ທຽບເທົ່າໂດລາ"
            End If
            CheckBox5.Text = "ສະແດງເອກະສານຊ້ອນທ້າຍ"
            Label10.Text = "ລາຍເຊັນ1"
            Label14.Text = "ລາຍເຊັນ2"
            Label13.Text = "ລາຍເຊັນ3"
            Label12.Text = "ລາຍເຊັນ4"
            Label11.Text = "ທີ່"
            Label15.Text = "ອັດຕາຜ່ານມາ"
            Button3.Text = "Export"
            Button4.Text = "Export (ທຽບບັນຊີ)"
        Else
            Button3.Text = "Export"
            Button4.Text = "Export (EQVL)"

            Label15.Text = "Rate Prev"
            Label10.Text = "Signature1"
            Label14.Text = "Signature2"
            Label13.Text = "Signature3"
            Label12.Text = "Signature4"
            Label11.Text = "Location"
            Button2.Text = "EQVL ACC"
            CheckBox4.Text = "EQVL Money"
            If CMB_Curr.Text = "LAK" Then
                CheckBox4.Text = "EQVL USD"
            ElseIf CMB_Curr.Text = "USD" Then
                CheckBox4.Text = "EQVL LAK"
            Else
                CheckBox4.Text = "EQVL USD"
            End If
            CheckBox5.Text = "Show Doc"
        End If

    End Sub
    Private Sub selectMMM()
        DMonth.Items.Clear()
        If FmMain.MnLaoLang.Checked = True Then
            DMonth.Items.Add("ມັງກອນ")
            DMonth.Items.Add("ກຸມພາ")
            DMonth.Items.Add("ມີນາ")
            DMonth.Items.Add("ເມສາ")
            DMonth.Items.Add("ພຶດສະພາ")
            DMonth.Items.Add("ມິຖຸນາ")
            DMonth.Items.Add("ກໍລະກົດ")
            DMonth.Items.Add("ສິງຫາ")
            DMonth.Items.Add("ກັນຍາ")
            DMonth.Items.Add("ຕຸລາ")
            DMonth.Items.Add("ພະຈິກ")
            DMonth.Items.Add("ທັນວາ")
        Else
            DMonth.Items.Add("January")
            DMonth.Items.Add("February")
            DMonth.Items.Add("March")
            DMonth.Items.Add("April")
            DMonth.Items.Add("May")
            DMonth.Items.Add("June")
            DMonth.Items.Add("July")
            DMonth.Items.Add("August")
            DMonth.Items.Add("September")
            DMonth.Items.Add("October")
            DMonth.Items.Add("November")
            DMonth.Items.Add("December")
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If r1.Checked = True Then
            RPT_GRP = " "
            RPT_GRPID = 0
        ElseIf r2.Checked = True Then
            RPT_GRP = "And GRP=1"
            RPT_GRPID = 1
        ElseIf r3.Checked = True Then
            RPT_GRP = "And GRP=2"
            RPT_GRPID = 2
        End If
        FmBLS_Item_Old.ShowDialog()
        FmBLS_Item_Old.Focus()
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
            MdStartDate_MM = Format(CDate("01/" & s & "/" & Year(Hyy.Value) - 1), "dd-MM-yyyy")
            MdToDate_MM = Format(CDate("01/" & x & "/" & Year(Hyy.Value) - 1), "dd-MM-yyyy")


            Dim SM As Date = DateAdd("M", CDbl(1), MdToDate)
            Dim SD As Date = DateAdd("d", CDbl(-1), SM)
            MdToDate = Format(CDate(SD), "dd-MM-yyyy")
            Lb.Text = "ປະຈຳເດືອນ " & M1.Text & " ຫາ " & M2.Text & "/" & yy.Text
        Else
            Dim s As Double = M1.SelectedIndex + 1
            Dim x As Double = M2.SelectedIndex + 1
            MdStartDate = Format(CDate("01/" & s & "/" & Year(Hyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("01/" & x & "/" & Year(Hyy.Value)), "dd-MM-yyyy")
            MdStartDate_MM = Format(CDate("01/" & s & "/" & Year(Hyy.Value) - 1), "dd-MM-yyyy")
            MdToDate_MM = Format(CDate("01/" & x & "/" & Year(Hyy.Value) - 1), "dd-MM-yyyy")

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

    Private Sub CMB_Curr_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMB_Curr.SelectedIndexChanged

        Dim rs As New ADODB.Recordset
        Call LoadSqlData("Select * From Curr_For_Rate Where   Curr =N'" & Trim(CMB_Curr.Text) & "'", rs)
        If rs.RecordCount > 0 Then
            txtcurr_name2.Text = Trim(rs("Curr_name").Value.ToString)
        End If

        MDRate_DT = " and rate_dt<='" & Format(Dt.Value, "yyyy-MM-dd") & "'  "
        MDRate_DT = " and rate_dt<='" & Format(MdToDate, "yyyy-MM-dd") & "'  "
        If CMB_Curr.SelectedIndex = 0 Then
            SS_Curr = " and AP_Rate_history.Curr =N'USD' "
        Else
            SS_Curr = " and AP_Rate_history.Curr =N'" & CMB_Curr.Text & "' "
        End If


        Call RateSetting()
        txtRate.Text = Format(MD_Rate, "#,##0.00")
        'MDRate_DT = " and rate_dt<='" & Format(MdToDate_MM, "yyyy-MM-dd") & "'  "
        'Call RateSetting()
        txtRate_Last.Text = Format(MD_Rate, "#,##0.00")

        If MuLng = "L" Then
            Button2.Text = "ທຽບບັນຊີ"
            CheckBox4.Text = "ທຽບເທົ່າເງິນ"
            CheckBox6.Text = "ທຽບເທົ່າບາດ"
            Label23.Text = "ບາດ-ກີບ"
            Label24.Text = "ໂດລາ-ກີບ"
            If CMB_Curr.Text = "LAK" Then
                CheckBox4.Text = "ທຽບເທົ່າໂດລາ"
                CheckBox6.Text = "ທຽບເທົ່າບາດ"
                Label23.Text = "ບາດ-ກີບ"
                Label24.Text = "ໂດລາ-ກີບ"
            ElseIf CMB_Curr.Text = "USD" Then
                CheckBox4.Text = "ທຽບເທົ່າກີບ"
            ElseIf CMB_Curr.Text = "THB" Then
                CheckBox4.Text = "ທຽບເທົ່າກີບ"
            Else
                CheckBox4.Text = "ທຽບເທົ່າໂດລາ"
                CheckBox6.Text = "ທຽບເທົ່າບາດ"
                Label23.Text = "ບາດ-ກີບ"
                Label24.Text = "ໂດລາ-ກີບ"
            End If
        Else
            Button2.Text = "EQVL ACC"
            CheckBox4.Text = "EQVL Money"
            If CMB_Curr.Text = "LAK" Then
                CheckBox4.Text = "EQVL USD"
            ElseIf CMB_Curr.Text = "USD" Then
                CheckBox4.Text = "EQVL LAK"
            Else
                CheckBox4.Text = "EQVL USD"
            End If
        End If

    End Sub
    Private Sub AAA()
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        'Call ChangBalance()
        Call LoadSqlData("select * from Ap_Rpt_BLS_Old_M_M ", RSC)
        If RSC.RecordCount = 0 Then
            CNN.Execute("insert into Ap_Rpt_BLS_Old_M_M(Rpt_ID, Description, Descriptione, Amt, Chart_of_Accounts_Codes, Grp, Rpt_Type, Grp_Nme, Fnt, Clor, Last_Amt, Udln, Lck, clt_Str, clt_Amt, NL, x) " & _
         " select Rpt_ID, Description, Descriptione, Amt, Chart_of_Accounts_Codes, Grp, Rpt_Type, Grp_Nme, Fnt, Clor, Last_Amt, Udln, Lck, clt_Str, clt_Amt, NL, x from Ap_Rpt_BLS_Old ")
        End If

        BLNEW()
        CNN.Execute("update  Ap_Rpt_BLS_Item_Old set Last_amt_dr  =  " & CDbl(0) & " , Last_amt_cr  = " & CDbl(0) & " ,amt_dr  =  " & CDbl(0) & " , amt_cr  = " & CDbl(0) & "  ")
        CNN.Execute("update Ap_Rpt_BLS_Old set Last_Amt = 0 , Amt  =  " & CDbl(0) & "  ")
        CNN.Execute("update Ap_Rpt_BLS_Old_M_M set Last_Amt = 0 , Amt  =  " & CDbl(0) & "  ")

        CNN.Execute("DELETE FROM Ap_Rpt_BLS_Detail ")
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

            'Lb.Text = s1 & " " & s2 & "/" & Year(MdToDate)
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

        Dim MI As String = "delete AP_Sum "
        MI = MI & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,amt,0 from Ap_Rpt_BLS_Old     where rpt_ID='4.1'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.1'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.2'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.3'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.4'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.5'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.6'  "
        MI = MI & "  update AP_Sum set amt=(select sum(amt2)+sum(amt3) from AP_Sum where (rpt_ID='4.1.1')  ) where rpt_ID='4.1'   "
        MI = MI & "  delete  AP_Sum  where rpt_ID='4.1.1'   "
        MI = MI & " Update Ap_Rpt_BLS_Old set Ap_Rpt_BLS_Old.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_BLS_Old    where Ap_Rpt_BLS_Old.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(MI)
        Dim Last As String = "delete AP_Sum "
        Last = Last & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,amt,0 from Ap_Rpt_BLS_Old     where rpt_ID='4.1'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.1'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.2'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.3'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.4'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.5'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.6'  "
        Last = Last & "  update AP_Sum set amt=(select sum(amt2)+sum(amt3) from AP_Sum where (rpt_ID='4.1.1')  ) where rpt_ID='4.1'   "
        Last = Last & "  delete  AP_Sum  where rpt_ID='4.1.1'   "
        Last = Last & " Update Ap_Rpt_BLS_Old set Ap_Rpt_BLS_Old.Last_amt=AP_Sum.amt from AP_Sum,Ap_Rpt_BLS_Old    where Ap_Rpt_BLS_Old.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(Last)


    End Sub

    Private Sub BLNEW_MM()


        New_Code = "3901000"
        New_Code4 = "00.3901000"

        Code_Dr = "4"
        Code_Dr1 = "00.4"
        Code_Cr = "5"
        Code_Cr1 = "00.5"

        Ac_Code = ""
        'MsgBox(MdStartDate & "==" & MdToDate)

        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        Dim B_Curr As String = ""
        If CMB_Curr.SelectedIndex = 0 Then
            B_Curr = ""
        Else
            B_Curr = " AND  Curr=N'" & CMB_Curr.Text & "' "
        End If


        CNN.Execute("DELETE  Ap_balance_6_col ")
        CNN.Execute("DELETE FROM Ap_balance_6 ")
        If CMB_Curr.SelectedIndex = 0 Then
            If CheckBox4.Checked = True Then
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
 " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amt_USD_Dr)as amt_dr , sum(amt_USD_Cr)as amt_cr  from gen_jn  WHERE  1=1 " & B_Curr & " and   gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")
                Dim S As Date = MdStartDate_MM : S = DateAdd("d", CDbl(-1), MdStartDate_MM)
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code , sum(amt_USD_Dr)as amt_dr , sum(amt_USD_Cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE  1=1 " & B_Curr & " and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code  , sum(amt_USD_Dr) as amt_dr , sum(amt_USD_Cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1 " & B_Curr & " and     date_work='" & "1-1-" & Format(MdStartDate_MM, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
            Else

                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
               " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  1=1  and Curr=N'LAK'   and   gen_jn.date_work   BETWEEN '" & Format(MdStartDate_MM, "yyyy-MM-dd") & "' AND '" & Format(MdToDate_MM, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")

                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
           " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amount_Dr)* " & CDbl(txtRate_Last.Text) & " as amt_dr , sum(amount_cr)* " & CDbl(txtRate_Last.Text) & " as amt_cr  from gen_jn  WHERE  1=1  and Curr=N'USD'   and   gen_jn.date_work   BETWEEN '" & Format(MdStartDate_MM, "yyyy-MM-dd") & "' AND '" & Format(MdToDate_MM, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")

                Dim S As Date = MdStartDate_MM : S = DateAdd("d", CDbl(-1), MdStartDate_MM)
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE  1=1   and Curr=N'LAK'   and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate_MM, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
           " select ac_code , sum(amount_Dr)* " & CDbl(txtRate_Last.Text) & " as amt_dr , sum(amount_cr)* " & CDbl(txtRate_Last.Text) & " as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE  1=1   and Curr=N'USD'   and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate_MM, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")


                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1  and Curr=N'LAK' and     date_work='" & "1-1-" & Format(MdStartDate_MM, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
         " select ac_code  , sum(amt_Dr)  as amt_dr , sum(amt_cr)   as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1  and Curr=N'USD'  and date_work='" & "1-1-" & Format(MdStartDate_MM, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")


                '============
                '                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '      " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  1=1  and Curr=N'LAK'   and   gen_jn.date_work   BETWEEN '" & Format(MdStartDate_MM, "yyyy-MM-dd") & "' AND '" & Format(MdToDate_MM, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")
                '                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '" select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr ,sum(amount_Dr)* " & CDbl(txtRate_Last.Text) & "  , sum(amount_Cr)* " & CDbl(txtRate_Last.Text) & "  from gen_jn  WHERE  1=1  and Curr=N'USD'   and   gen_jn.date_work   BETWEEN '" & Format(MdStartDate_MM, "yyyy-MM-dd") & "' AND '" & Format(MdToDate_MM, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")

                '                Dim S As Date = MdStartDate_MM : S = DateAdd("d", CDbl(-1), MdStartDate_MM)
                '                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '                " select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE  1=1  and Curr=N'LAK'   and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate_MM, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

                '                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '                " select ac_code  , sum(amt_dr)* " & CDbl(txtRate_Last.Text) & "    , sum(amt_Cr)* " & CDbl(txtRate_Last.Text) & "     , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1  and Curr=N'USD'   and     date_work='" & "1-1-" & Format(MdStartDate_MM, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
                '                '===============
                '                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '           " select ac_code , sum(amount_Dr)as amt_dr , sum(amount_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE  1=1  and Curr=N'LAK'   and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate_MM, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

                '                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                '                " select ac_code  , sum(amount_Dr)* " & CDbl(txtRate_Last.Text) & "  , sum(amount_Dr)* " & CDbl(txtRate_Last.Text) & "  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1 and Curr=N'USD'   and     date_work='" & "1-1-" & Format(MdStartDate_MM, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")



            End If
        Else
            CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
  " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(Amount_Dr)as amt_dr , sum(Amount_Cr)as amt_cr  from gen_jn  WHERE  1=1 " & B_Curr & " and  gen_jn.date_work   BETWEEN '" & Format(MdStartDate_MM, "yyyy-MM-dd") & "' AND '" & Format(MdToDate_MM, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")
            Dim S As Date = MdStartDate_MM : S = DateAdd("d", CDbl(-1), MdStartDate_MM)
            CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            " select ac_code , sum(Amount_Dr)as amt_dr , sum(Amount_Cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE 1=1 " & B_Curr & " and  gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate_MM, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
            CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            " select ac_code  , sum(Amount_Dr) as amt_dr , sum(Amount_Cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE  1=1 " & B_Curr & " and    date_work='" & "1-1-" & Format(MdStartDate_MM, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

        End If

        CNN.Execute("UPDATE Ap_balance_6 set Ac_Code = left(Ac_Code,7) ")

        CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr)as open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")
        If CMB_Curr.SelectedIndex = 0 Then
            '        CNN.Execute(" insert into Ap_balance_6_col(ac_code, ac_name, open_amt_dr, open_amt_cr,  amt_dr, amt_cr,   Rem_dr, Rem_cr,Status) " & _
            '" Select '2382120.00.0000',N'ຄູ່ມູນຄ່າຖານະ ແລກປ່ຽນເງິນຕາ ຕ່າງປະເທດ',sum(open_amt_dr),sum(open_amt_cr),sum(amt_dr),sum(amt_cr),sum(Rem_dr),sum(Rem_cr),1   " & _
            '" from Ap_balance_6_col where left(ac_code,7)='2382120' group by  left(ac_code,7) ")
            '        CNN.Execute("delete  Ap_balance_6_col  where left(ac_code,7)='2382120' and Status is null ")
        ElseIf CMB_Curr.SelectedIndex = 1 Then
            CNN.Execute(" insert into Ap_balance_6_col(ac_code, ac_name, open_amt_dr, open_amt_cr,  amt_dr, amt_cr,   Rem_dr, Rem_cr,Status) " & _
" Select '2382120.00.0000',N'ຄູ່ມູນຄ່າຖານະ ແລກປ່ຽນເງິນຕາ ຕ່າງປະເທດ',sum(open_amt_dr),sum(open_amt_cr),sum(amt_dr),sum(amt_cr),sum(Rem_dr),sum(Rem_cr),1   " & _
" from Ap_balance_6_col where left(ac_code,7)='2382120' group by  left(ac_code,7) ")
            CNN.Execute("delete  Ap_balance_6_col  where left(ac_code,7)='2382120' and Status is null ")
        Else

            'CNN.Execute("delete  Ap_balance_6_col  where left(ac_code,7)='2382120'")

        End If

        CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        Call Chang_Incom()
        CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + amt_dr) - (open_amt_cr + amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr + amt_cr) ")
        CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr + amt_cr) - (open_amt_dr + amt_dr) where (open_amt_cr + amt_cr) >= (open_amt_dr + amt_dr) ")
        CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")

    End Sub
    Private Sub Button2_Click_3(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        AddHeader()
        Call selectLoad()
        Call Office()
        Call AAA()
        CNN.Execute(" UPDATE Ap_Rpt_BLS_Old_M_M set Ap_Rpt_BLS_Old_M_M.amt=Ap_Rpt_BLS_Old.amt from Ap_Rpt_BLS_Old_M_M,Ap_Rpt_BLS_Old where Ap_Rpt_BLS_Old.rpt_ID=Ap_Rpt_BLS_Old_M_M.rpt_ID ")

        BLNEW_MM()
        CNN.Execute("update  Ap_Rpt_BLS_Item_Old set Last_amt_dr  =  " & CDbl(0) & " , Last_amt_cr  = " & CDbl(0) & " ,amt_dr  =  " & CDbl(0) & " , amt_cr  = " & CDbl(0) & "  ")
        CNN.Execute("update Ap_Rpt_BLS_Old set Last_Amt = 0 , Amt  =  " & CDbl(0) & "  ")
        'CNN.Execute("update Ap_Rpt_BLS_Old_M_M set Last_Amt = 0 , Amt  =  " & CDbl(0) & "  ")

        CNN.Execute("DELETE FROM Ap_Rpt_BLS_Detail ")
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

            'Lb.Text = s1 & " " & s2 & "/" & Year(MdToDate)
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

        Dim MI As String = "delete AP_Sum "
        MI = MI & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,amt,0 from Ap_Rpt_BLS_Old     where rpt_ID='4.1'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.1'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.2'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.3'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.4'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.5'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.6'  "
        MI = MI & "  update AP_Sum set amt=(select sum(amt2)+sum(amt3) from AP_Sum where (rpt_ID='4.1.1')  ) where rpt_ID='4.1'   "
        MI = MI & "  delete  AP_Sum  where rpt_ID='4.1.1'   "
        MI = MI & " Update Ap_Rpt_BLS_Old set Ap_Rpt_BLS_Old.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_BLS_Old    where Ap_Rpt_BLS_Old.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(MI)
        Dim Last As String = "delete AP_Sum "
        Last = Last & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,amt,0 from Ap_Rpt_BLS_Old     where rpt_ID='4.1'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.1'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.2'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.3'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.4'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.5'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.6'  "
        Last = Last & "  update AP_Sum set amt=(select sum(amt2)+sum(amt3) from AP_Sum where (rpt_ID='4.1.1')  ) where rpt_ID='4.1'   "
        Last = Last & "  delete  AP_Sum  where rpt_ID='4.1.1'   "
        Last = Last & " Update Ap_Rpt_BLS_Old set Ap_Rpt_BLS_Old.Last_amt=AP_Sum.amt from AP_Sum,Ap_Rpt_BLS_Old    where Ap_Rpt_BLS_Old.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(Last)

        If CheckBox1.Checked = False Then
            If CheckBox2.Checked = True Then
                CNN.Execute(" UPDATE Ap_Rpt_BLS_Old_M_M set Ap_Rpt_BLS_Old_M_M.Last_amt=Ap_Rpt_BLS_Old.amt from Ap_Rpt_BLS_Old_M_M,Ap_Rpt_BLS_Old where Ap_Rpt_BLS_Old.rpt_ID=Ap_Rpt_BLS_Old_M_M.rpt_ID ")

                Call M_M()
                Exit Sub
            End If
            'Call LoadReport()
        Else
            Call LoadReportItem()
        End If
    End Sub

    Private Sub txtRate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRate.KeyPress
        If e.KeyChar = Chr(13) Then
            txtRate.Text = Format(CDbl(txtRate.Text), "#,##0.00")
        End If
    End Sub

    Private Sub txtRate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRate.LostFocus
        If txtRate.Text = "" Then
            txtRate.Text = 1
            txtRate.Text = Format(CDbl(txtRate.Text), "#,##0.00")
        End If
    End Sub

    Private Sub txtRate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRate.TextChanged

    End Sub

    Private Sub txtRate_Last_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRate_Last.KeyPress
        If e.KeyChar = Chr(13) Then
            txtRate_Last.Text = Format(CDbl(txtRate_Last.Text), "#,##0.00")
        End If
    End Sub

    Private Sub txtRate_Last_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRate_Last.LostFocus
        If txtRate_Last.Text = "" Then
            txtRate_Last.Text = 1
            txtRate_Last.Text = Format(CDbl(txtRate_Last.Text), "#,##0.00")
        End If
    End Sub

    Private Sub txtRate_Last_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRate_Last.TextChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call AddHeader()
        Call selectLoad()
        Call Office()
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        'Call ChangBalance()
        BLNEW()
        CNN.Execute("update  Ap_Rpt_BLS_Item_Old set Last_amt_dr  =  " & CDbl(0) & " , Last_amt_cr  = " & CDbl(0) & " ,amt_dr  =  " & CDbl(0) & " , amt_cr  = " & CDbl(0) & "  ")
        CNN.Execute("update Ap_Rpt_BLS_Old set Last_Amt = 0 , Amt  =  " & CDbl(0) & "  ")
        CNN.Execute("DELETE FROM Ap_Rpt_BLS_Detail ")
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

            'Lb.Text = s1 & " " & s2 & "/" & Year(MdToDate)
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

        Dim MI As String = "delete AP_Sum "
        MI = MI & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,amt,0 from Ap_Rpt_BLS_Old     where rpt_ID='4.1'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.1'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.2'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.3'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.4'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.5'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.6'  "
        MI = MI & "  update AP_Sum set amt=(select sum(amt2)+sum(amt3) from AP_Sum where (rpt_ID='4.1.1')  ) where rpt_ID='4.1'   "
        MI = MI & "  delete  AP_Sum  where rpt_ID='4.1.1'   "
        MI = MI & " Update Ap_Rpt_BLS_Old set Ap_Rpt_BLS_Old.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_BLS_Old    where Ap_Rpt_BLS_Old.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(MI)
        Dim Last As String = "delete AP_Sum "
        Last = Last & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,amt,0 from Ap_Rpt_BLS_Old     where rpt_ID='4.1'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.1'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.2'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.3'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.4'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.5'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.6'  "
        Last = Last & "  update AP_Sum set amt=(select sum(amt2)+sum(amt3) from AP_Sum where (rpt_ID='4.1.1')  ) where rpt_ID='4.1'   "
        Last = Last & "  delete  AP_Sum  where rpt_ID='4.1.1'   "
        Last = Last & " Update Ap_Rpt_BLS_Old set Ap_Rpt_BLS_Old.Last_amt=AP_Sum.amt from AP_Sum,Ap_Rpt_BLS_Old    where Ap_Rpt_BLS_Old.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(Last)


        If CheckBox1.Checked = False Then
            'If CheckBox2.Checked = True Then
            '    Call M_M()
            '    Exit Sub
            'End If
            Call LoadReport_Exp()
        Else
            Call LoadReportItem()
        End If
        'MdStartDate = d1
        'MdToDate = d2
    End Sub
    Private Sub LoadReport_Exp()
        Dim RPT_ID As String

        Dim AA As String = Lb.Text
        RPT_ID = " "
        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As RptSjUd ,"
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

        If CMB_Curr.Text = "EQVL" Then
            If CheckBox4.Checked = True Then
                LngId = "7121" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"

            Else
                LngId = "7106" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"

            End If
        ElseIf CMB_Curr.Text = "LAK" Then
            LngId = "7097" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        ElseIf CMB_Curr.Text = "USD" Then
            LngId = "7108" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        End If


        'If CMB_Curr.SelectedIndex = 0 Then
        '    LngId = "7106" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        'ElseIf CMB_Curr.SelectedIndex = 1 Then
        '    LngId = "7107" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        'ElseIf CMB_Curr.SelectedIndex = 2 Then
        '    LngId = "7108" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        'End If

        If RD.Checked = True Then
            LngId = "7090" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7091" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"
            Month_IN = ""
            Month_Last = ""
            LngId = "7113" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
        ElseIf RM.Checked = True Then
            LngId = "7060" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7061" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"

            LngId = "7096" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
        ElseIf RP.Checked = True Then
            Month_IN = ""
            Month_Last = ""
            LngId = "7083" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7063" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7112" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
        ElseIf Rhalf.Checked = True Then
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"
            LngId = "7060" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7061" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7113" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
            Month_IN = ""
            Month_Last = ""
        ElseIf RT.Checked = True Then
            Month_IN = ""
            Month_Last = ""
            LngId = "7080" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7081" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"
            LngId = "7113" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
        ElseIf RY.Checked = True Then
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"
            LngId = "7084" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7065" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7113" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
        End If





        If r1.Checked = True Then
            LngId = "7092" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            r = "And GRP>0"
            'r = "And GRP<4"
        ElseIf r2.Checked = True Then
            LngId = "7093" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            'r = "And GRP<5"
            r = "And GRP<2"
        ElseIf r3.Checked = True Then
            LngId = "7054" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            'r = "And GRP>5"
            'r = "And GRP>1 and GRP<=3"
            r = "And GRP>1"
        End If
        Dim Doc As Integer = 0
        If CheckBox5.Checked = True Then
            Doc = 1
        Else
            Doc = 0
        End If


        SLF = "SELECT  " & Doc & "  as Doc  , N'" & MuOffDep & "'  as RptSjoff_Dep  , " & mformat & "  as mformat  ,  " & MuLngRpt & "  *, N'" & Month_IN & "' as In_Mn, N'" & Month_Last & "' as Last_Mn   FROM Ap_Rpt_BLS_Old  "

        Call LoadLoGO()

        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            .Open(" " & SLF & " where grp<>'' " & RPT_ID & " " & r & "order by Rpt_Id asc  ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        Dim FrmPreview As New FmPreview : FrmClosing()
        'Dim Rpt As New CryRpt_BLS
        'Dim Rpt As New Object
        If RM.Checked = True Then
            Rpt = New CryRpt_BLS_M_Exp
        Else
            Rpt = New CryRpt_BLS_Exp
        End If

        'Dim Rpt As New CryRpt_BLS_NNN
        If MdShowLOGO = 1 Then
            Rpt.Subreports(0).SetDataSource(RsLOGO)
        End If
        Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S1"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS1.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S2"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS2.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S3"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS3.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S4"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS4.Text

        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("pp"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtPP.Text

        Rpt.SetDataSource(Rs)
        'Rpt.Refresh()
        'FrmPreview.ReportViewer.ReportSource = Rpt
        'Rpt.ReportViewer.exportreport(Rs)
        'Rpt.ExportReport()
        'Rpt = Nothing
        Rpt.SetDataSource(Rs)
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.DisplayGroupTree = False
        FrmPreview.WindowState = FormWindowState.Maximized
        FrmPreview.Show()
        FrmPreview.Focus()


    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        AddHeader()
        Call selectLoad()
        Call Office()
        Call AAA()
        CNN.Execute(" UPDATE Ap_Rpt_BLS_Old_M_M set Ap_Rpt_BLS_Old_M_M.amt=Ap_Rpt_BLS_Old.amt from Ap_Rpt_BLS_Old_M_M,Ap_Rpt_BLS_Old where Ap_Rpt_BLS_Old.rpt_ID=Ap_Rpt_BLS_Old_M_M.rpt_ID ")

        BLNEW_MM()
        CNN.Execute("update  Ap_Rpt_BLS_Item_Old set Last_amt_dr  =  " & CDbl(0) & " , Last_amt_cr  = " & CDbl(0) & " ,amt_dr  =  " & CDbl(0) & " , amt_cr  = " & CDbl(0) & "  ")
        CNN.Execute("update Ap_Rpt_BLS_Old set Last_Amt = 0 , Amt  =  " & CDbl(0) & "  ")
        'CNN.Execute("update Ap_Rpt_BLS_Old_M_M set Last_Amt = 0 , Amt  =  " & CDbl(0) & "  ")

        CNN.Execute("DELETE FROM Ap_Rpt_BLS_Detail ")
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

            'Lb.Text = s1 & " " & s2 & "/" & Year(MdToDate)
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

        Dim MI As String = "delete AP_Sum "
        MI = MI & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,amt,0 from Ap_Rpt_BLS_Old     where rpt_ID='4.1'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.1'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.2'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.3'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.4'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.5'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.6'  "
        MI = MI & "  update AP_Sum set amt=(select sum(amt2)+sum(amt3) from AP_Sum where (rpt_ID='4.1.1')  ) where rpt_ID='4.1'   "
        MI = MI & "  delete  AP_Sum  where rpt_ID='4.1.1'   "
        MI = MI & " Update Ap_Rpt_BLS_Old set Ap_Rpt_BLS_Old.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_BLS_Old    where Ap_Rpt_BLS_Old.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(MI)
        Dim Last As String = "delete AP_Sum "
        Last = Last & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,amt,0 from Ap_Rpt_BLS_Old     where rpt_ID='4.1'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.1'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.2'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.3'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.4'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.5'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '4.1.1'  ,0 ,Last_amt from Ap_Rpt_BLS_Old     where rpt_ID='4.1.6'  "
        Last = Last & "  update AP_Sum set amt=(select sum(amt2)+sum(amt3) from AP_Sum where (rpt_ID='4.1.1')  ) where rpt_ID='4.1'   "
        Last = Last & "  delete  AP_Sum  where rpt_ID='4.1.1'   "
        Last = Last & " Update Ap_Rpt_BLS_Old set Ap_Rpt_BLS_Old.Last_amt=AP_Sum.amt from AP_Sum,Ap_Rpt_BLS_Old    where Ap_Rpt_BLS_Old.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(Last)

        If CheckBox1.Checked = False Then
            If CheckBox2.Checked = True Then
                CNN.Execute(" UPDATE Ap_Rpt_BLS_Old_M_M set Ap_Rpt_BLS_Old_M_M.Last_amt=Ap_Rpt_BLS_Old.amt from Ap_Rpt_BLS_Old_M_M,Ap_Rpt_BLS_Old where Ap_Rpt_BLS_Old.rpt_ID=Ap_Rpt_BLS_Old_M_M.rpt_ID ")

                Call M_M_Export()
                Exit Sub
            End If
            'Call LoadReport()
        Else
            Call LoadReportItem()
        End If
    End Sub
    Private Sub M_M_Export()
        Dim RPT_ID As String

        Dim AA As String = Lb.Text
        RPT_ID = " "
        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As RptSjUd ,"
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

        If CMB_Curr.Text = "EQVL" Then
            If CheckBox4.Checked = True Then
                LngId = "7121" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"

            Else
                LngId = "7106" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"

            End If
        ElseIf CMB_Curr.Text = "LAK" Then
            LngId = "7097" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        ElseIf CMB_Curr.Text = "USD" Then
            LngId = "7108" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        End If


        'If CMB_Curr.SelectedIndex = 0 Then
        '    LngId = "7106" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        'ElseIf CMB_Curr.SelectedIndex = 1 Then
        '    LngId = "7107" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        'ElseIf CMB_Curr.SelectedIndex = 2 Then
        '    LngId = "7108" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        'End If

        If RD.Checked = True Then
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"
            LngId = "7090" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7091" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        ElseIf RM.Checked = True Then
            LngId = "7060" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7061" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"

            LngId = "7096" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
        ElseIf RP.Checked = True Then
            LngId = "7083" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7063" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7112" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"
        ElseIf Rhalf.Checked = True Then
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"
            LngId = "7060" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7061" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"

        ElseIf RT.Checked = True Then
            LngId = "7080" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7081" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"
        ElseIf RY.Checked = True Then
            LngId = "7109" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_BL ,"
            LngId = "7084" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7065" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            LngId = "7113" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
        End If





        If r1.Checked = True Then
            LngId = "7092" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            r = "And GRP>0"
            'r = "And GRP<4"
        ElseIf r2.Checked = True Then
            LngId = "7093" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            'r = "And GRP<5"
            r = "And GRP<2"
        ElseIf r3.Checked = True Then
            LngId = "7054" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
            'r = "And GRP>5"
            'r = "And GRP>1 and GRP<=3"
            r = "And GRP>1"
        End If


        Dim Doc As Integer = 0
        If CheckBox5.Checked = True Then
            Doc = 1
        Else
            Doc = 0
        End If
        SLF = "SELECT    " & Doc & "  as Doc  ,  N'" & MuOffDep & "'  as RptSjoff_Dep  ,  " & mformat & "  as mformat  ,  " & MuLngRpt & "  *, N'" & Month_IN & "' as In_Mn, N'" & Month_IN_MM & "' as Last_Mn   FROM Ap_Rpt_BLS_Old_M_M  "



        Call LoadLoGO()

        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            .Open(" " & SLF & " where grp<>'' " & RPT_ID & " " & r & "order by Rpt_Id asc  ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        Dim FrmPreview As New FmPreview : FrmClosing()

        'Dim Rpt As New CryRpt_BLSM_M

        If RM.Checked = True Then
            Rpt = New CryRpt_BLS_M_Exp
        Else
            Rpt = New CryRpt_BLS_Exp
        End If
        'Dim Rpt As New CryRpt_BLS_NNN
        If MdShowLOGO = 1 Then
            Rpt.Subreports(0).SetDataSource(RsLOGO)
        End If
        Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S1"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS1.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S2"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS2.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S3"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS3.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("S4"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtS4.Text

        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("pp"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = TxtPP.Text

        'Rpt.SetDataSource(Rs)
        'Rpt.Refresh()
        'FrmPreview.ReportViewer.ReportSource = Rpt
        'Rpt.ReportViewer.ExportReport()
        'Rpt = Nothing
        Rpt.SetDataSource(Rs)
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.DisplayGroupTree = False
        FrmPreview.WindowState = FormWindowState.Maximized
        FrmPreview.Show()
        FrmPreview.Focus()
    End Sub

    Private Sub txtRate2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRate2.KeyPress
        If e.KeyChar = Chr(13) Then
            txtRate2.Text = Format(CDbl(txtRate2.Text), "##,##0.00")
        End If
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRate2.TextChanged

    End Sub
End Class