Public Class FmRpt_Income
    Dim CLT_Str, CLT_Last_Str, sk As String
    Dim bls1 As String
    Dim MonthLetter1 As String
    Dim MdStartDate As Date
    Dim MdToDate As Date
    Dim MdQuarter As Date
    Dim MdStartDate_Last As Date
    Dim MdToDate_Last As Date
    Dim ny, ly, n_L_y As String
    Dim sql As String
    Dim VCode1, VCode2, VCode3, VCode4, VCode5, VCode6, VCode7, VCode8, VCode9 As String
    Dim RsOpen As New ADODB.Recordset
    Dim RsOpenMonth As New ADODB.Recordset
    Dim RsRpt As New ADODB.Recordset
    Dim AmtOpenDR, AmtOpenCR, AmtOpenMonthDR, AmtOpenMonthCR As Double
    Dim VOpenDate As Date
    Dim RptNme As String
    Dim RSC12 As New ADODB.Recordset
    Dim D, P As String
    Dim RSCIn_M As New ADODB.Recordset

    Private Sub ChangBalance()
        New_Code = "3901000"
        Code_Dr = "4"
        Code_Cr = "5"
        Ac_Code = ""
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        CNN.Execute("DELETE  Ap_balance_6_col ")
        CNN.Execute("DELETE FROM Ap_balance_6 ")

        If RY.Checked = True Then

            CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")
            Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
            S = DateAdd("y", CDbl(-1), MdStartDate)
            'Dim T As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
            Dim dd As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            " select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & CDbl(Format(MdStartDate, "yyyy")) - 1 & "-1-1" & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code"
            CNN.Execute(dd)
            CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            " select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE    date_work='" & CDbl(Format(MdStartDate, "yyyy")) - 1 & "-1-1" & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

        Else

            Dim PPP As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
             " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code "
            CNN.Execute(PPP)
            Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
            CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            " select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
            CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            " select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE    date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

        End If



        CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr)as open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        'Call Chang_Incom()
        CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + amt_dr) - (open_amt_cr + amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr + amt_cr) ")
        CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr + amt_cr) - (open_amt_dr + amt_dr) where (open_amt_cr + amt_cr) >= (open_amt_dr + amt_dr) ")
        CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")

    End Sub

    Private Sub Chang_Incom()

        Insr = "delete  Ap_balance_6  " & _
         "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) values ('" & New_Code & "' ,0,0,0,0) " & _
      "update Ap_balance_6 set  open_amt_Dr =  " & _
      "(select top 1  (select  (Sum(open_amt_dr))-( Sum(open_amt_cr)) As open_amt_dr from Ap_balance_6_col where left(Ac_Code,1)='" & Code_Dr & "' )  As Dr " & _
      "from Ap_balance_6_col )  where  Ac_Code ='" & New_Code & "'  " & _
  "update Ap_balance_6 set  open_amt_cr =  " & _
  "(select top 1  (select  (Sum(open_amt_cr))-( Sum(open_amt_dr)) As open_amt_dr from Ap_balance_6_col where left(Ac_Code,1)='" & Code_Cr & "'  )  As Cr " & _
   "from Ap_balance_6_col ) where  Ac_Code ='" & New_Code & "'   " & _
  "update Ap_balance_6 set  amt_Dr = " & _
  "(select top 1  (select  (Sum(amt_dr))-( Sum(amt_cr)) As amt_dr from Ap_balance_6_col where left(Ac_Code,1)='" & Code_Dr & "'  )  As Dr " & _
  "from Ap_balance_6_col )  where  Ac_Code ='" & New_Code & "'  " & _
  "update Ap_balance_6 set  amt_cr =  " & _
  "(select top 1  (select  (Sum(amt_cr))-( Sum(amt_dr)) As amt_dr from Ap_balance_6_col where  left(Ac_Code,1)='" & Code_Cr & "' )  As Cr " & _
  "from Ap_balance_6_col ) where  Ac_Code ='" & New_Code & "'  " & _
  "update  Ap_balance_6_col set open_amt_dr = 0 where open_amt_dr  is null  " & _
  "update  Ap_balance_6_col set open_amt_cr = 0 where open_amt_cr  is null  " & _
  "update  Ap_balance_6_col set amt_dr = 0 where amt_dr  is null  " & _
  "update  Ap_balance_6_col set amt_cr = 0 where amt_cr  is null   " & _
  " Update  Ap_balance_6 set   open_amt_dr = (open_amt_cr  - open_amt_dr ) , open_amt_cr=0  where (open_amt_cr  - open_amt_dr )>= 0 " & _
   "Update  Ap_balance_6 set   open_amt_cr = (open_amt_dr  - open_amt_cr) , open_amt_dr=0  where (open_amt_cr  - open_amt_dr )<= 0 " & _
  "Update  Ap_balance_6 set   amt_dr = (amt_cr  - amt_dr ) , amt_cr=0  where (amt_cr  - amt_dr )>= 0 " & _
  " Update  Ap_balance_6 set   amt_cr = (amt_dr  - amt_cr) , amt_dr=0  where (amt_cr  - amt_dr )<= 0 " & _
   "  update Ap_balance_6 set  Ap_balance_6.open_amt_dr = Ap_balance_6.open_amt_dr + Ap_balance_6_col.open_amt_dr , Ap_balance_6.open_amt_cr = Ap_balance_6.open_amt_cr + Ap_balance_6_col.open_amt_cr   ,  Ap_balance_6.amt_dr = Ap_balance_6.amt_dr + Ap_balance_6_col.amt_dr    ,  Ap_balance_6.amt_cr = Ap_balance_6.amt_cr + Ap_balance_6_col.amt_cr    from Ap_balance_6 , Ap_balance_6_col   where  Ap_balance_6.Ac_Code = Ap_balance_6_col.Ac_Code      " & _
   "Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr " & _
    "Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr " & _
    "delete  Ap_balance_6_col   where left(Ac_Code,1)='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "'  Or Ac_Code = '" & New_Code & "' " & _
     "  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr from Ap_balance_6"
        CNN.Execute(Insr)

    End Sub


    Private Sub BtnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        'Click_Last()
        Call SelectLoad()
        '============
        'CNN.Execute("update  Ap_Rpt_Income_Item set Last_Amt_dr  =  0 , Last_Amt_Cr  =  0 , amt_dr  = 0 , amt_cr  = 0  ")
        Dim saa As String = "update  Ap_Rpt_Income_Item set Last_Amt_dr  =  0 , Last_Amt_Cr  =  0 , amt_dr  = 0 , amt_cr  = 0 "
        CNN.Execute(saa)

        CNN.Execute("update Ap_Rpt_Income set  Last_Amt  = 0 , Amt  = 0    ")
        CNN.Execute("DELETE FROM Ap_Rpt_Incon_Detail ")
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        ChangBalance()
        ' ''    CNN.Execute("DELETE  Ap_balance_6_col ")
        ' ''    CNN.Execute("DELETE FROM Ap_balance_6 ")
        ' ''    CNN.Execute(" insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  from Ap_balance_TB where  date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  order by Ac_Code asc ")
        ' ''    Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
        ' ''    Dim ppp As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ' ''          " select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from Ap_balance_TB  WHERE  date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "'   group BY ac_code"
        ' ''    CNN.Execute(ppp)
        ' ''    'Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
        ' ''    'Dim ppp As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ' ''    '      " select ac_code , sum(open_amt_dr), sum(open_amt_cr)  , sum(amt_dr)   ,sum(amt_cr)    from Ap_balance_TB  WHERE  date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'   group BY ac_code"
        ' ''    'CNN.Execute(ppp)

        ' ''    CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ' ''" select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr)as open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")
        ' ''    CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        ' ''    CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")

        ' ''    CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + amt_dr) - (open_amt_cr + amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr + amt_cr) ")
        ' ''    CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr + amt_cr) - (open_amt_dr + amt_dr) where (open_amt_cr + amt_cr) >= (open_amt_dr + amt_dr) ")
        ' ''    CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")

        SelcectInLast()
        UpdateIInLast()
        SelectOutLast()
        UpdateOut()
        Update_Sum()


        If RD.Checked = True Then
            Dim s1, s2 As String
            LngId = 7027 : CallLngStr() : s2 = LngStr
            LngId = 7072 : CallLngStr() : s1 = LngStr
            Lb.Text = s1 & " " & Ds.Text & " " & s2 & " " & Dt.Text
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
        ElseIf Rhalf.Checked = True Then
            Dim s1, s2 As String
            'LngId = 7078 + M1.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7089 : CallLngStr() : s1 = LngStr

            LngId = 7027 : CallLngStr() : s2 = LngStr
            'Lb.Text = s1 & " " & M1.SelectedIndex + 1 & " " & s2 & " " & M2.SelectedIndex + 1 & " (" & Year(MdToDate) & " )"


        ElseIf RT.Checked = True Then
            Dim s1, s7 As String
            LngId = 7049 : CallLngStr() : s1 = LngStr

            If Ct.SelectedIndex = 0 Then
                s7 = "0" & Ct.SelectedIndex + 1
                LngId = 7078 : CallLngStr()
                Lb.Text = LngStr & " " & yyt.Text
            Else
                s7 = Ct.SelectedIndex + 1
                LngId = 7079 : CallLngStr()
                Lb.Text = LngStr & " " & yyt.Text
            End If


        ElseIf RY.Checked = True Then
            Dim s1, s2, s3 As String
            LngId = 7074 + Period.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7062 : CallLngStr() : s1 = LngStr
            LngId = 7064 : CallLngStr() : s3 = LngStr
            Lb.Text = s3 & " " & yy.Text
        End If


        If CheckBox3.Checked = True Then
            CNN.Execute("Update Ap_Rpt_Income set Amt=Amt+Last_Amt")
        End If
        '========1 =========
        CNN.Execute("update Ap_Rpt_Income set amt=(select isnull(sum(amt),0) as amt  from Ap_Rpt_Income where (rpt_ID='1.1.1' or  rpt_ID='1.1.1.1'  or  rpt_ID='1.1.1.2'  or  rpt_ID='1.1.1.2.1'  or  rpt_ID='1.1.1.2.2' )) where  rpt_ID='1' ")

        '========2 =========
        CNN.Execute("update Ap_Rpt_Income set amt=(select isnull(sum(amt),0) as amt  from Ap_Rpt_Income where (rpt_ID='1.1.1.3' or  rpt_ID='1.1.1.4'  or  rpt_ID='1.1.2'  or  rpt_ID='1.1.2.5'  or  rpt_ID='1.13' )) where  rpt_ID='1.1.1.2.3' ")


        Dim MI As String = "delete AP_Sum "
        MI = MI & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,amt,0 from Ap_Rpt_Income     where rpt_ID='1'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  'I'  ,0 ,amt from Ap_Rpt_Income     where rpt_ID='1.1.1.2.3'  "
        MI = MI & "  update AP_Sum set amt=(select sum(amt2)-sum(amt3) from AP_Sum where (rpt_ID='1' or rpt_ID='I')  ) where rpt_ID='I'   "
        MI = MI & "  delete  AP_Sum  where rpt_ID='1'   "
        MI = MI & " Update Ap_Rpt_Income set Ap_Rpt_Income.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Income    where Ap_Rpt_Income.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(MI)
        '========1,1=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.1.1' and  rpt_ID<='1.1.5') where  rpt_ID='1.1' ")
        '========1,2=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.2.1' and  rpt_ID<='1.2.5') where  rpt_ID='1.2' ")
        '========I=========
        '========Sum I =========
        Dim SI As String = "delete AP_Sum "
        SI = SI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,amt,0 from Ap_Rpt_Income     where rpt_ID='1.1' "
        SI = SI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  'I'  ,0 ,amt from Ap_Rpt_Income     where rpt_ID='1.2'  "
        SI = SI & " update AP_Sum set amt=(select sum(amt2)-sum(amt3) from AP_Sum where (rpt_ID='1.1' or rpt_ID='I')  ) where rpt_ID='I' "
        SI = SI & " delete  AP_Sum  where rpt_ID='1.1' "
        SI = SI & " Update Ap_Rpt_Income set Ap_Rpt_Income.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Income    where Ap_Rpt_Income.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(SI)
        '========3=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.3.1' and  rpt_ID<='1.3.2') where  rpt_ID='1.3' ")
        '========4=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.4.1' and  rpt_ID<='1.4.2') where  rpt_ID='1.4'")
        '========5========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID='1.5.1' ) where  rpt_ID='1.5' ")
        '========6======== 
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.6.1' and  rpt_ID<='1.6.2') where  rpt_ID='1.6' ")
        '========7======== 
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID='1.7.1') where  rpt_ID='1.7' ")
        '========8======== 
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.8.1' and  rpt_ID<='1.8.2') where  rpt_ID='1.8'")
        '========9======== 
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.9.1' and  rpt_ID<='1.9.5') where  rpt_ID='1.9' ")
        '========10======= 
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.10.1' and  rpt_ID<='1.10.5') where  rpt_ID='1.10'")


        '========11=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.11.1' and  rpt_ID<='1.11.2') where  rpt_ID='1.11'")

        '========12=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.12.1' and  rpt_ID<='1.12.2') where  rpt_ID='1.12'")

        '========13=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.13.1' and  rpt_ID<='1.13.2') where  rpt_ID='1.13'")
        '========14=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.14.1' and  rpt_ID<='1.14.2') where  rpt_ID='1.14'")
        '========15=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.15.1' and  rpt_ID<='1.15.5') where  rpt_ID='1.15' ")
        '========16========= 
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.16.1' and  rpt_ID<='1.16.2') where  rpt_ID='1.16'")
        '========17=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.17.1' and  rpt_ID<='1.17.2') where  rpt_ID='1.17'")
        '========18=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.18.1' and  rpt_ID<='1.18.4') where  rpt_ID='1.18' ")
        '========19=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.19.1' and  rpt_ID<='1.19.2') where  rpt_ID='1.19'")
        '========20=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.20.1' and  rpt_ID<='1.20.4') where  rpt_ID='1.20'")
        '========21=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.21.1' and  rpt_ID<='1.21.3') where  rpt_ID='1.21'")
        CNN.Execute("UPDATE Ap_Rpt_Income set amt=0 where amt is null")
        '========Sum III=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where (rpt_ID='1.15' or  rpt_ID='1.16' or  rpt_ID='1.17' or  rpt_ID='1.18' or  rpt_ID='1.19' or  rpt_ID='1.20')) where  rpt_ID='III' ")

        '========Sum II =========
        Dim SII As String = "delete AP_Sum "
        SII = SII & " insert into AP_Sum(Rpt_ID,amt2,amt3) select 'II' ,amt,0 from Ap_Rpt_Income     where rpt_ID='I' "
        SII = SII & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  'II'  ,0 ,amt from Ap_Rpt_Income     where (rpt_ID='3' or rpt_ID='4' or rpt_ID='5' or rpt_ID='6' or rpt_ID='7' or rpt_ID='8' or rpt_ID='9' or rpt_ID='10' or rpt_ID='11' or rpt_ID='12' or rpt_ID='13' or rpt_ID='14') "
        SII = SII & " update AP_Sum set amt=(select sum(amt2)-sum(amt3) from AP_Sum where (rpt_ID='II' or rpt_ID='3' or rpt_ID='4' or rpt_ID='5' or rpt_ID='6' or rpt_ID='7' or rpt_ID='8' or rpt_ID='9' or rpt_ID='10' or rpt_ID='11' or rpt_ID='12' or rpt_ID='13' or rpt_ID='14' )  ) where rpt_ID='II' "
        SII = SII & " Update Ap_Rpt_Income set Ap_Rpt_Income.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Income    where Ap_Rpt_Income.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(SII)
        'SII = SII & " delete  AP_Sum  where rpt_ID='I' "

        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where (rpt_ID='I' or  rpt_ID='1.3' or  rpt_ID='1.4' or  rpt_ID='1.5' or  rpt_ID='1.6' or  rpt_ID='1.7' or  rpt_ID='1.8' or  rpt_ID='1.9' or  rpt_ID='1.10' or  rpt_ID='1.11' or  rpt_ID='1.12' or  rpt_ID='1.13'  or  rpt_ID='1.14' )) where  rpt_ID='II' ")
        ''========Sum III=========
        '========Sum IV=========
        Dim IVV As String = "delete AP_Sum "
        IVV = IVV & " insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,amt,0 from Ap_Rpt_Income     where rpt_ID='II' "
        IVV = IVV & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  'IV'  ,0 ,amt from Ap_Rpt_Income     where rpt_ID='III'  "
        IVV = IVV & " update AP_Sum set amt=(select sum(amt2)-sum(amt3) from AP_Sum where (rpt_ID='II' or rpt_ID='IV')  ) where rpt_ID='IV' "
        IVV = IVV & " delete  AP_Sum  where rpt_ID='II' "
        IVV = IVV & " Update Ap_Rpt_Income set Ap_Rpt_Income.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Income    where Ap_Rpt_Income.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(IVV)

        '========Sum  V=========
        Dim VV As String = "delete AP_Sum "
        VV = VV & " insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,amt,0 from Ap_Rpt_Income     where rpt_ID='IV' "
        VV = VV & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  'V'  ,0 ,amt from Ap_Rpt_Income     where rpt_ID='1.21'  "
        VV = VV & " update AP_Sum set amt=(select sum(amt2)-sum(amt3) from AP_Sum where (rpt_ID='IV' or rpt_ID='V')  ) where rpt_ID='V' "
        VV = VV & " delete  AP_Sum  where rpt_ID='IV' "
        VV = VV & " Update Ap_Rpt_Income set Ap_Rpt_Income.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Income    where Ap_Rpt_Income.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(VV)

        If CheckBox1.Checked = False Then
            Call LoadReport()
        Else
            Call LoadReportItem()
        End If
    End Sub


    Private Sub LoadOpen_Jn1()
        Dim RSC12 As New ADODB.Recordset

        LoadSqlData("   select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & MULook & " group BY ac_code ", RSC12)


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
        If Format(MdStartDate, "MM") = "01" Or Format(MdStartDate, "MM") = "02" Or Format(MdStartDate, "MM") = "03" Then
            DS = "01/01/" & Format(MdStartDate, "yyyy")
            DT = MdToDate
        End If

        If Format(MdStartDate, "MM") = "04" Or Format(MdStartDate, "MM") = "05" Or Format(MdStartDate, "MM") = "06" Then
            DS = "01/04/" & Format(MdStartDate, "yyyy")
            DT = MdToDate
        End If

        If Format(MdStartDate, "MM") = "07" Or Format(MdStartDate, "MM") = "08" Or Format(MdStartDate, "MM") = "09" Then
            DS = "01/07/" & Format(MdStartDate, "yyyy")
            DT = MdToDate
        End If

        If Format(MdStartDate, "MM") = "10" Or Format(MdStartDate, "MM") = "11" Or Format(MdStartDate, "MM") = "12" Then
            DS = "01/10/" & Format(MdStartDate, "yyyy")
            DT = MdToDate
        End If


        LoadSqlData("   select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(DS, "yyyy-MM-dd") & "' AND '" & Format(DT, "yyyy-MM-dd") & "' " & MULook & " group BY ac_code ", RSC12)


        With RSC12
            Do Until .EOF = True

                Call LoadOpen_Jn14()
                .MoveNext()
            Loop
        End With

    End Sub
    Private Sub LoadOpen_Jn12_11()
        CNN.Execute("Update Ap_balance_6_col set Amt_Last_M_dr=0,Amt_Last_M_cr=0")
        Dim DS, DT As Date
        If Format(MdStartDate, "MM") <> "01" Then
            DS = DateAdd(DateInterval.Month, -1, MdStartDate)
            DT = DateAdd(DateInterval.Month, -1, MdToDate)
            LoadSqlData("   select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy") & "-1-1" & "' AND '" & Format(DT, "yyyy-MM-dd") & "' " & MULook & " group BY ac_code ", RSC12)
            With RSC12
                Do Until .EOF = True

                    Call LoadOpen_Jn14_11()
                    .MoveNext()
                Loop
            End With
        End If
    End Sub
    Private Sub LoadOpen_Jn14_11()
        Dim RSC14 As New ADODB.Recordset
        LoadSqlData("   select * from Ap_balance_6_col  where ac_code = '" & CStr(Trim(RSC12.Fields("ac_code").Value)) & "'  ", RSC14)
        If RSC14.RecordCount <> 0 Then
            CNN.Execute("Update Ap_balance_6_col set Amt_Last_M_dr= " & CDbl(Trim(RSC12.Fields("amt_dr").Value)) & "  , Amt_Last_M_cr= " & CDbl(Trim(RSC12.Fields("amt_cr").Value)) & "  where ac_code = '" & CStr(Trim(RSC12.Fields("ac_code").Value)) & "'  ")
        Else
            CNN.Execute("INSERT INTO Ap_balance_6_col( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status, Quarter_dr , Quarter_cr , Amt_Last_M_dr , Amt_Last_M_cr  ) " & _
              "Values('" & CStr(Trim(RSC12.Fields("ac_Code").Value)) & "', " & _
              " 0 , 0 , 0 , 0,0, 0 , 0 , " & CDbl(Trim(RSC12.Fields("amt_dr").Value)) & " , " & CDbl(Trim(RSC12.Fields("amt_cr").Value)) & " )")
        End If
    End Sub
    Private Sub LoadOpen_Jn14()
        Dim RSC14 As New ADODB.Recordset
        LoadSqlData("   select * from Ap_balance_6_col  where ac_code = '" & CStr(Trim(RSC12.Fields("ac_code").Value)) & "'  ", RSC14)
        If RSC14.RecordCount <> 0 Then
            CNN.Execute("Update Ap_balance_6_col set Quarter_dr= " & CDbl(Trim(RSC12.Fields("amt_dr").Value)) & "  , Quarter_cr= " & CDbl(Trim(RSC12.Fields("amt_cr").Value)) & "  where ac_code = '" & CStr(Trim(RSC12.Fields("ac_code").Value)) & "'  ")
        Else
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
        LoadSqlData("   select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "'  " & MULook & " group BY ac_code ", RSC12)
        With RSC12
            Do Until .EOF = True
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
               "Values('" & CStr(Trim(.Fields("ac_Code").Value)) & "', " & _
               " " & CDbl(0) & ", " & CDbl(0) & ", " & CDbl(Trim(.Fields("amt_dr").Value)) & ", " & CDbl(Trim(.Fields("amt_cr").Value)) & ",0 )")
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
            Else
                LoadOpen_Jn5()
            End If
        End With
    End Sub

    Private Sub LoadOpen_Jn5()
        Dim RSC5 As New ADODB.Recordset
        With RSC
            LoadSqlData("select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr from Open_jn  WHERE    ac_code='" & VCode3 & "' " & MULook & " group BY ac_code", RSC5)
            If RSC5.RecordCount > 0 Then
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr , Status  ) " & _
             "Values('" & CStr(Trim(RSC5.Fields("ac_Code").Value)) & "',  " & _
             " " & CDbl(RSC5.Fields("amt_dr").Value) & ", " & CDbl(RSC5.Fields("amt_cr").Value) & ", " & CDbl(0) & ", " & CDbl(0) & ",0 )")
            Else
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
                    CNN.Execute("Update Ap_balance_6 set rem_dr='" & CDbl(op_dr + amt_dr) - CDbl(op_cr + amt_cr) & "' , rem_cr='" & CDbl(0) & "' where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                End If
                If CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) >= 0 Then
                    CNN.Execute("Update Ap_balance_6 set rem_dr='" & CDbl(0) & "' , rem_cr='" & CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) & "' where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                End If
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
    End Sub
    Private Sub KKK()
        For i = 1 To Len(D)
            If Mid(D, i, 1) = "." Then
                P = Microsoft.VisualBasic.Left(D, i - 1)
                Exit Sub
            Else
                P = D
            End If

        Next i
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

    Private Sub SelcectIn()
        LoadSqlData("select * from Ap_Rpt_Income_Item where  Rpt_Type = 'In'", RSCIn_M)
        With RSCIn_M
            Do Until .EOF = True
                UpdateIIn_Item()
                .MoveNext()
            Loop
        End With
    End Sub


    Private Sub SelcectInLast()
        LoadSqlData("select * from Ap_Rpt_Income_Item where  Rpt_Type = 'In'  ", RSCIn_M)
        With RSCIn_M
            Do Until .EOF = True
                UpdateIIn_ItemLast()
                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub UpdateIIn_Item()
        Dim RSCkk As New ADODB.Recordset
        LoadSqlData(" select * from Ap_balance_6_col   where ac_code Like  '" & (RSCIn_M.Fields("Ac_Code").Value) & "%' ", RSCkk)
        With RSCkk
            Do Until .EOF = True
                CNN.Execute("insert into Ap_Rpt_Incon_Detail (Rpt_ID , Ac_Code , Ac_Name  , Amt_Last_M_Dr  , Amt_Last_M_Cr , Amt_M_Dr , Amt_M_Cr , Amt_Q_Dr , Amt_Q_Cr , Amt_Y_Dr , Amt_Y_Cr , Ac_Code_Parent) Values ( '" & CStr((RSCIn_M.Fields("Rpt_ID").Value)) & "' ,  '" & CStr((.Fields("Ac_Code").Value)) & "' ,  N'" & CStr((.Fields("Ac_Name").Value)) & "'  ,  " & CDbl((.Fields("Amt_Last_M_Dr").Value)) & " , " & CDbl((.Fields("Amt_Last_M_Cr").Value)) & " ,  " & CDbl((.Fields("amt_dr").Value)) & " , " & CDbl((.Fields("amt_cr").Value)) & " , " & CDbl((.Fields("Quarter_dr").Value)) & " ,   " & CDbl((.Fields("Quarter_cr").Value)) & " , " & CDbl((.Fields("Rem_dr").Value)) & " , " & CDbl((.Fields("Rem_cr").Value)) & " , " & CStr((RSCIn_M.Fields("Ac_Code").Value)) & ") ")
                CNN.Execute("update  Ap_Rpt_Income_Item set Amt_Last_M_Dr  =  Amt_Last_M_Dr+" & CDbl((.Fields("Amt_Last_M_Dr").Value)) & " , Amt_Last_M_Cr  =  Amt_Last_M_Cr+" & CDbl((.Fields("Amt_Last_M_Cr").Value)) & " , amt_M_dr  =  amt_M_dr+" & CDbl((.Fields("amt_dr").Value)) & " , amt_M_cr  = amt_M_cr+" & CDbl((.Fields("amt_cr").Value)) & "  , amt_Q_dr  = amt_Q_dr+ " & CDbl(CDbl((.Fields("Quarter_dr").Value))) & " , amt_Q_cr  = amt_Q_cr+" & CDbl(CDbl((.Fields("Quarter_cr").Value))) & " , amt_y_dr  =  amt_y_dr+" & CDbl((.Fields("Rem_dr").Value)) & " , amt_y_cr  = amt_y_cr+" & CDbl((.Fields("Rem_cr").Value)) & "    where ac_code Like  '" & (RSCIn_M.Fields("Ac_Code").Value) & "%' And  Rpt_Type = 'In' ")
                .MoveNext()
            Loop
        End With

    End Sub
    Private Sub UpdateIIn_ItemLast()
        Dim RSCkk As New ADODB.Recordset
        Dim PPa As String = " select * from Ap_balance_6_col   where ac_code =N'" & (RSCIn_M.Fields("Ac_Code").Value) & "' "
        LoadSqlData(PPa, RSCkk)
        With RSCkk
            Do Until .EOF = True
                CNN.Execute("Insert into Ap_Rpt_Incon_Detail (  Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type ) values ( '" & CStr((RSCIn_M.Fields("Rpt_Id").Value)) & "' , '" & CStr((.Fields("Ac_Code").Value)) & "' , N'" & CStr((.Fields("Ac_Name").Value)) & "'   , " & CDbl((.Fields("open_amt_dr").Value)) & " , " & CDbl((.Fields("open_amt_Cr").Value)) & "   , " & CDbl((.Fields("Amt_dr").Value)) & " , " & CDbl((.Fields("Amt_cr").Value)) & " , 'In')")
                CNN.Execute("update  Ap_Rpt_Income_Item set  Last_amt_dr  =  Last_amt_dr+" & CDbl((.Fields("open_amt_dr").Value)) & " , Last_amt_cr  = Last_amt_cr+" & CDbl((.Fields("open_amt_Cr").Value)) & " , Amt_Dr  =  Amt_Dr+" & CDbl((.Fields("Amt_dr").Value)) & " , Amt_Cr  = Amt_Cr+" & CDbl((.Fields("Amt_cr").Value)) & "   where ac_code = '" & (RSCIn_M.Fields("ac_code").Value) & "' And  Rpt_Type = 'In' ")
                .MoveNext()
            Loop
        End With
    End Sub
    Private Sub Update_Sum()
        CNN.Execute("update Ap_Rpt_Incon_Detail set  Rpt_Name=Ap_Rpt_Income.Description from   Ap_Rpt_Incon_Detail , Ap_Rpt_Income where Ap_Rpt_Incon_Detail.Rpt_Id = Ap_Rpt_Income.Rpt_Id")
        CNN.Execute(" Update Caculate_Rpt set  CLT_Amt  = CLT_Str ,  CLT_Last_Amt  = CLT_Str where CLT_Str = '+' or CLT_Str = '-' or CLT_Str = '*' or CLT_Str = '+' or CLT_Str = '/' or CLT_Str = '(' or CLT_Str=')' Or CLT_Str<>'Cast(('   Or CLT_Str<>')As Float)'")
        CNN.Execute("delete Caculate_Lock")
        CNN.Execute("delete Caculate_Start")
        CNN.Execute(" Insert Into Caculate_Start (Rpt_Id , Rpt_Type , clt_Str , clt_Amt , Clt_Last_Amt ) select Rpt_Id , Rpt_Type , clt_Str , clt_Amt , Clt_Last_Amt from Caculate_Rpt where Rpt_Type = 'INC'  Order by  Rpt_id ,cnt asc ")
        CNN.Execute("update Caculate_Start set lck =0")
        CNN.Execute("Insert into Caculate_Lock (cnt_Mt)  SELECT  (SELECT     TOP 1 cnt FROM Caculate_Start AS B WHERE(Rpt_Id = A.Rpt_Id   ) ORDER BY cnt desc) AS cnt FROM Caculate_Start  AS A  GROUP BY Rpt_Id ORDER BY Rpt_Id")
        CNN.Execute("update  Caculate_Start set lck=1 from Caculate_Start ,Caculate_Lock  where Caculate_Start.cnt=Caculate_Lock.cnt_MT")
        CNN.Execute("  Update Caculate_Start set Caculate_Start.Amt = Ap_Rpt_Income.Amt , Caculate_Start.Last_Amt = Ap_Rpt_Income.Last_Amt   from Caculate_Start , Ap_Rpt_Income  where  Caculate_Start.CLT_Str  = Ap_Rpt_Income.Rpt_Id  ")
        CNN.Execute("Update Caculate_Start set lck_Amt=0")
        CNN.Execute("Update Caculate_Start set lck_Amt=1 where CLT_Str <> '+' And CLT_Str <> '-' And CLT_Str <> '*' And CLT_Str <> '+' And CLT_Str <> '/' And CLT_Str <> '(' And CLT_Str<>')' And CLT_Str<>')' And CLT_Str<>'Cast(('   And CLT_Str<>')As Float)' ")
        Dim RSC1 As New ADODB.Recordset
        CLT_Str = ""
        CLT_Last_Str = ""




        With RSC1
            Call LoadSqlData("select *  from Caculate_Start where Rpt_Type = 'INC'    Order by  Rpt_id ,cnt asc", RSC1)
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
                            Dim s1 As String = " Update  Ap_Rpt_Income set Amt = " & CLT_Str & " , Last_Amt = " & CLT_Last_Str & " where  Rpt_ID =   '" & (RSC1.Fields("Rpt_ID").Value.ToString) & "' "
                            CNN.Execute(s1)
                        Else
                            Dim s As String = " Update  Ap_Rpt_Income set Amt = " & CLT_Str & " , Last_Amt = " & CLT_Last_Str & " where  Rpt_ID =   '" & (RSC1.Fields("Rpt_ID").Value.ToString) & "' "
                            Dim k As String = s

                            'MessageBox.Show("ສູດຄິດໄລ່ຂອງ " & (RSC1.Fields("Rpt_ID").Value.ToString) & " = " & CLT_Str & " ບໍ່ຖຶກຕ້ອງກະລຸນນາກວດສອບຄືນໃຫມ່")
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
    Private Sub Update_Sumxx()
        Dim InteIn_M As Double = 0
        Dim InteIn_Q As Double = 0
        Dim InteIn_Y As Double = 0
        Dim xm1 As Double = 0
        Dim xm2 As Double = 0
        Dim xq1 As Double = 0
        Dim xq2 As Double = 0
        Dim xy1 As Double = 0
        Dim xy2 As Double = 0


        '==========5
        Dim GOI_M As Double = 0
        Dim GOI_Q As Double = 0
        Dim GOI_Y As Double = 0
        Dim GOIM3 As Double = 0
        Dim GOIM4 As Double = 0
        Dim GOIQ3 As Double = 0
        Dim GOIQ4 As Double = 0
        Dim GOIY3 As Double = 0
        Dim GOIY4 As Double = 0

        '==========7
        Dim NI_M As Double = 0
        Dim NI_Q As Double = 0
        Dim NI_Y As Double = 0
        Dim NIM5 As Double = 0
        Dim NIM6 As Double = 0
        Dim NIQ5 As Double = 0
        Dim NIQ6 As Double = 0
        Dim NIY5 As Double = 0
        Dim NIY6 As Double = 0


        '==========10
        Dim PBT_M As Double = 0
        Dim PBT_Q As Double = 0
        Dim PBT_Y As Double = 0
        Dim PBTM7 As Double = 0
        Dim PBTM8 As Double = 0
        Dim PBTM9 As Double = 0
        Dim PBTQ7 As Double = 0
        Dim PBTQ8 As Double = 0
        Dim PBTQ9 As Double = 0
        Dim PBTY7 As Double = 0
        Dim PBTY8 As Double = 0
        Dim PBTY9 As Double = 0

        '==========12
        Dim NP_M As Double = 0
        Dim NP_Q As Double = 0
        Dim NP_Y As Double = 0
        Dim NPM10 As Double = 0
        Dim NPM11 As Double = 0
        Dim NPQ10 As Double = 0
        Dim NPQ11 As Double = 0
        Dim NPY10 As Double = 0
        Dim NPY11 As Double = 0

        Dim RSC As New ADODB.Recordset
        LoadSqlData("select * from Ap_Rpt_Income  ", RSC)
        With RSC
            Do Until .EOF = True

                If (.Fields("Grp").Value) = "01" Then
                    xm1 = xm1 + CDbl((.Fields("Current_Month").Value))
                    xq1 = xq1 + CDbl((.Fields("Quarter_to_Date").Value))
                    xy1 = xy1 + CDbl((.Fields("Year_to_Date").Value))
                End If
                If (.Fields("Grp").Value) = "02" Then
                    xm2 = xm2 + CDbl((.Fields("Current_Month").Value))
                    xq2 = xq2 + CDbl((.Fields("Quarter_to_Date").Value))
                    xy2 = xy2 + CDbl((.Fields("Year_to_Date").Value))
                End If

                .MoveNext()
            Loop
        End With
        'MsgBox(xy1)
        'MsgBox(xy2)
        'MsgBox(InteIn_Y)
        InteIn_M = xm1 - xm2
        InteIn_Q = xq1 - xq2
        InteIn_Y = xy1 - xy2
        CNN.Execute("Update Ap_Rpt_Income set " & _
                    "Current_Month =" & InteIn_M & " , " & _
                    " Quarter_to_Date =" & InteIn_Q & " , " & _
                    " Year_to_Date =" & InteIn_Y & " " & _
                    "where Rpt_ID='03'")
        '===========5


        LoadSqlData("select * from Ap_Rpt_Income  ", RSC)
        With RSC
            Do Until .EOF = True

                '==========5
                If (.Fields("Grp").Value) = "03" Then
                    GOIM3 = GOIM3 + CDbl((.Fields("Current_Month").Value))
                    GOIQ3 = GOIQ3 + CDbl((.Fields("Quarter_to_Date").Value))
                    GOIY3 = GOIY3 + CDbl((.Fields("Year_to_Date").Value))
                End If
                If (.Fields("Grp").Value) = "04" Then
                    GOIM4 = GOIM4 + CDbl((.Fields("Current_Month").Value))
                    GOIQ4 = GOIQ4 + CDbl((.Fields("Quarter_to_Date").Value))
                    GOIY4 = GOIY4 + CDbl((.Fields("Year_to_Date").Value))
                End If

                .MoveNext()
            Loop
        End With


        GOI_M = GOIM3 + GOIM4
        GOI_Q = GOIQ3 + GOIQ4
        GOI_Y = GOIY3 + GOIY4

        CNN.Execute("Update Ap_Rpt_Income set " & _
                    "Current_Month =" & GOI_M & " , " & _
                    " Quarter_to_Date =" & GOI_Q & " , " & _
                    " Year_to_Date =" & GOI_Y & " " & _
                    "where Rpt_ID='05'")


        '===========7



        LoadSqlData("select sum(Current_Month) As Current_Month , sum(Quarter_to_Date) As Quarter_to_Date  , sum(Year_to_Date) As Year_to_Date  from Ap_Rpt_Income where Rpt_ID Like '06.05%' ", RSC)
        If RSC.RecordCount <> 0 Then
            CNN.Execute("Update Ap_Rpt_Income set Current_Month=" & CDbl((RSC.Fields("Current_Month").Value)) & " , Quarter_to_Date=" & CDbl((RSC.Fields("Quarter_to_Date").Value)) & " , Year_to_Date=" & CDbl((RSC.Fields("Year_to_Date").Value)) & " where Rpt_ID='06.05'")
        End If

        '===========7









        LoadSqlData("select * from Ap_Rpt_Income  ", RSC)
        With RSC
            Do Until .EOF = True

                '==========7
                If (.Fields("Grp").Value) = "05" Then
                    NIM5 = NIM5 + CDbl((.Fields("Current_Month").Value))
                    NIQ5 = NIQ5 + CDbl((.Fields("Quarter_to_Date").Value))
                    NIY5 = NIY5 + CDbl((.Fields("Year_to_Date").Value))
                End If
                If (.Fields("Rpt_Id").Value) = "06.01" Or (.Fields("Rpt_Id").Value) = "06.02" Or (.Fields("Rpt_Id").Value) = "06.03" Or (.Fields("Rpt_Id").Value) = "06.04" Or (.Fields("Rpt_Id").Value) = "06.05" Or (.Fields("Rpt_Id").Value) = "06.06" Or (.Fields("Rpt_Id").Value) = "06.07" Or (.Fields("Rpt_Id").Value) = "06.08" Then

                    'MsgBox(NIM6)

                    NIM6 = NIM6 + CDbl((.Fields("Current_Month").Value))
                    NIQ6 = NIQ6 + CDbl((.Fields("Quarter_to_Date").Value))
                    NIY6 = NIY6 + CDbl((.Fields("Year_to_Date").Value))
                End If


                .MoveNext()
            Loop
        End With
        NI_M = NIM5 - NIM6
        NI_Q = NIQ5 - NIQ6
        NI_Y = NIY5 - NIY6

        CNN.Execute("Update Ap_Rpt_Income set " & _
                    "Current_Month =" & NI_M & " , " & _
                    " Quarter_to_Date =" & NI_Q & " , " & _
                    " Year_to_Date =" & NI_Y & " " & _
                    "where Rpt_ID='07'")

        '===========10
        LoadSqlData("select * from Ap_Rpt_Income  ", RSC)
        With RSC
            Do Until .EOF = True
                '==========10
                If (.Fields("Grp").Value) = "07" Then
                    PBTM7 = PBTM7 + CDbl((.Fields("Current_Month").Value))
                    PBTQ7 = PBTQ7 + CDbl((.Fields("Quarter_to_Date").Value))
                    PBTY7 = PBTY7 + CDbl((.Fields("Year_to_Date").Value))
                End If
                If (.Fields("Grp").Value) = "08" Then
                    PBTM8 = PBTM8 + CDbl((.Fields("Current_Month").Value))
                    PBTQ8 = PBTQ8 + CDbl((.Fields("Quarter_to_Date").Value))
                    PBTY8 = PBTY8 + CDbl((.Fields("Year_to_Date").Value))
                End If
                If (.Fields("Grp").Value) = "09" Then
                    PBTM9 = PBTM9 + CDbl((.Fields("Current_Month").Value))
                    PBTQ9 = PBTQ9 + CDbl((.Fields("Quarter_to_Date").Value))
                    PBTY9 = PBTY9 + CDbl((.Fields("Year_to_Date").Value))
                End If
                .MoveNext()
            Loop
        End With
        PBT_M = CDbl(PBTM7 - PBTM8) + PBTM9
        PBT_Q = CDbl(PBTQ7 - PBTQ8) + PBTQ9
        PBT_Y = CDbl(PBTY7 - PBTY8) + PBTY9
        CNN.Execute("Update Ap_Rpt_Income set " & _
                    "Current_Month =" & PBT_M & " , " & _
                    " Quarter_to_Date =" & PBT_Q & " , " & _
                    " Year_to_Date =" & PBT_Y & " " & _
                    "where Rpt_ID='10'")
        '===========12
        LoadSqlData("select * from Ap_Rpt_Income  ", RSC)
        With RSC
            Do Until .EOF = True
                '==========12
                If (.Fields("Grp").Value) = "10" Then
                    NPM10 = NPM10 + CDbl((.Fields("Current_Month").Value))
                    NPQ10 = NPQ10 + CDbl((.Fields("Quarter_to_Date").Value))
                    NPY10 = NPY10 + CDbl((.Fields("Year_to_Date").Value))
                End If
                If (.Fields("Grp").Value) = "11" Then
                    NPM11 = NPM11 + CDbl((.Fields("Current_Month").Value))
                    NPQ11 = NPQ11 + CDbl((.Fields("Quarter_to_Date").Value))
                    NPY11 = NPY11 + CDbl((.Fields("Year_to_Date").Value))
                End If
                .MoveNext()
            Loop
        End With
        NP_M = NPM10 - NPM11
        NP_Q = NPQ10 - NPQ11
        NP_Y = NPY10 - NPY11
        CNN.Execute("Update Ap_Rpt_Income set " & _
                    "Current_Month =" & NP_M & " , " & _
                    " Quarter_to_Date =" & NP_Q & " , " & _
                    " Year_to_Date =" & NP_Y & " " & _
                    "where Rpt_ID='12'")
    End Sub



    Private Sub UpdateOut()
        Dim RSC As New ADODB.Recordset
        LoadSqlData("select Rpt_ID, sum(Last_Amt_dr) As Last_Amt_Dr , sum(Last_Amt_cr) As Last_Amt_cr , sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr    from Ap_Rpt_Income_Item  where  Rpt_Type = 'Out'   group by Rpt_ID ", RSC)
        With RSC
            Do Until .EOF = True
                CNN.Execute("Update Ap_Rpt_Income set " & _
                    " Amt ='" & CDbl(CDbl((.Fields("Amt_dr").Value)) - CDbl((.Fields("Amt_cr").Value))) & "' " & _
                      " ,Last_Amt ='" & CDbl(CDbl((.Fields("Last_Amt_dr").Value)) - CDbl((.Fields("Last_Amt_cr").Value))) & "' " & _
                       " where Rpt_ID = '" & (.Fields("Rpt_ID").Value) & "' ")
                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub UpdateOutLast()
        Dim RSC As New ADODB.Recordset
        LoadSqlData("select Rpt_ID         , sum(Amt_Last_M_Dr) As Amt_Last_M_Dr , sum(Amt_Last_M_Cr) As Amt_Last_M_Cr                      , sum(Amt_M_Dr) As Amt_M_Dr , sum(Amt_M_Cr) As Amt_M_Cr  , sum(Amt_Q_Dr) As Amt_Q_Dr ,  sum(Amt_Q_Cr) As Amt_Q_Cr , sum(Amt_y_Dr) As Amt_y_Dr , sum(Amt_y_Cr) As Amt_y_Cr  from Ap_Rpt_Income_Item  where  Rpt_Type = 'Out' group by Rpt_ID ", RSC)
        With RSC
            Do Until .EOF = True
                CNN.Execute("Update Ap_Rpt_Income set " & _
                              " Last_Month ='" & CDbl(CDbl((.Fields("Amt_Last_M_Dr").Value)) - CDbl((.Fields("Amt_Last_M_Cr").Value))) & "' ," & _
                            " Current_Month ='" & CDbl(CDbl((.Fields("Amt_M_dr").Value)) - CDbl((.Fields("Amt_M_cr").Value))) & "' ," & _
                              " Quarter_to_Date ='" & CDbl(CDbl((.Fields("Amt_Q_dr").Value)) - CDbl((.Fields("Amt_Q_cr").Value))) & "' ," & _
                                " Year_to_Date ='" & CDbl(CDbl((.Fields("Amt_y_dr").Value)) - CDbl((.Fields("Amt_y_cr").Value))) & "' " & _
                            " where Rpt_ID = '" & (.Fields("Rpt_ID").Value) & "' ")
                .MoveNext()
            Loop
        End With
    End Sub
    Private Sub UpdateIIn()
        Dim RSC As New ADODB.Recordset
        LoadSqlData("select Rpt_ID  , sum(Amt_Last_M_Dr) As Amt_Last_M_Dr , sum(Amt_Last_M_Cr) As Amt_Last_M_Cr , sum(Amt_M_Dr) As Amt_M_Dr , sum(Amt_M_Cr) As Amt_M_Cr  , sum(Amt_Q_Dr) As Amt_Q_Dr ,  sum(Amt_Q_Cr) As Amt_Q_Cr , sum(Amt_y_Dr) As Amt_y_Dr , sum(Amt_y_Cr) As Amt_y_Cr  from Ap_Rpt_Income_Item  where  Rpt_Type = 'In' group by Rpt_ID ", RSC)
        With RSC
            Do Until .EOF = True
                CNN.Execute("Update Ap_Rpt_Income set " & _
                               " Last_Month ='" & CDbl(CDbl((.Fields("Amt_Last_M_Cr").Value)) - CDbl((.Fields("Amt_Last_M_Dr").Value))) & "' ," & _
                            " Current_Month ='" & CDbl(CDbl((.Fields("Amt_M_cr").Value)) - CDbl((.Fields("Amt_M_dr").Value))) & "' ," & _
                              " Quarter_to_Date ='" & CDbl(CDbl((.Fields("Amt_Q_cr").Value)) - CDbl((.Fields("Amt_Q_dr").Value))) & "' ," & _
                                " Year_to_Date ='" & CDbl(CDbl((.Fields("Amt_y_cr").Value)) - CDbl((.Fields("Amt_y_dr").Value))) & "' " & _
                            " where Rpt_ID = '" & (.Fields("Rpt_ID").Value) & "' ")
                .MoveNext()
            Loop
        End With
    End Sub
    Private Sub UpdateIInLast()
        Dim RSC As New ADODB.Recordset
        LoadSqlData("select Rpt_ID , sum(Last_Amt_Dr) As Last_Amt_Dr , sum(Last_Amt_Cr) As Last_Amt_Cr ,sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr   from Ap_Rpt_Income_Item  where  Rpt_Type = 'In'   group by Rpt_ID ", RSC)
        With RSC
            Do Until .EOF = True

                CNN.Execute("Update Ap_Rpt_Income set " & _
                     " Amt = '" & CDbl(CDbl((.Fields("Amt_cr").Value)) - CDbl((.Fields("Amt_dr").Value))) & "' " & _
                       " , Last_Amt ='" & CDbl(CDbl((.Fields("Last_Amt_cr").Value)) - CDbl((.Fields("Last_Amt_dr").Value))) & "' " & _
                        " where Rpt_ID = '" & (.Fields("Rpt_ID").Value) & "' ")
                .MoveNext()
            Loop
        End With
    End Sub


    Private Sub SelectOut()

        LoadSqlData("select * from Ap_Rpt_Income_Item where  Rpt_Type = 'Out'", RSCIn_M)
        With RSCIn_M
            Do Until .EOF = True
                Call UpdateOut_ItemLast()
                .MoveNext()
            Loop
        End With
        'If RSCIn_M.State = ConnectionState.Open Then RSCIn_M.Close()
    End Sub

    Private Sub SelectOutLast()

        LoadSqlData("select * from Ap_Rpt_Income_Item where  Rpt_Type = 'Out'  ", RSCIn_M)
        With RSCIn_M
            Do Until .EOF = True
                Call UpdateOut_Item()
                .MoveNext()
            Loop
        End With
        'If RSCIn_M.State = ConnectionState.Open Then RSCIn_M.Close()
    End Sub

    Private Sub UpdateOut_Item()
        Dim RSCkk As New ADODB.Recordset
        LoadSqlData(" select * from Ap_balance_6_col   where ac_code =N'" & (RSCIn_M.Fields("Ac_Code").Value) & "' ", RSCkk)
        With RSCkk
            Do Until .EOF = True
                'MsgBox((RSCIn_M.Fields("Ac_Code").Value))
                Dim ppp As String = "Insert into Ap_Rpt_Incon_Detail (Ac_Code_Parent , Rpt_Id , Ac_Code , Ac_Name  ,  Last_Amt_Dr , Last_Amt_Cr, Amt_Dr , Amt_Cr , Rpt_Type ) values (  '" & CStr((RSCIn_M.Fields("Ac_Code").Value.ToString)) & "' , '" & CStr((RSCIn_M.Fields("Rpt_Id").Value.ToString)) & "' , '" & CStr((.Fields("Ac_Code").Value.ToString)) & "' , N'" & CStr((.Fields("Ac_Name").Value.ToString)) & "'  ,   " & CDbl((.Fields("Open_Amt_dr").Value)) & " , " & CDbl((.Fields("Open_Amt_cr").Value)) & " , " & CDbl((.Fields("Amt_dr").Value)) & " , " & CDbl((.Fields("Amt_cr").Value)) & " , 'Out' )"
                CNN.Execute(ppp)
                CNN.Execute("update  Ap_Rpt_Income_Item set Last_Amt_Dr  =  Last_Amt_Dr+" & CDbl((.Fields("Open_Amt_dr").Value)) & " , Last_Amt_Cr  = Last_Amt_Cr+" & CDbl((.Fields("Open_Amt_cr").Value)) & "  , Amt_Dr  =  Amt_Dr+" & CDbl((.Fields("Amt_dr").Value)) & " , Amt_Cr  = Amt_Cr+" & CDbl((.Fields("Amt_cr").Value)) & "   where ac_code = '" & (RSCIn_M.Fields("ac_code").Value) & "' And  Rpt_Type = 'Out' ")

                .MoveNext()
            Loop
        End With
    End Sub
    Private Sub UpdateOut_ItemLast()
        Dim RSCkk As New ADODB.Recordset
        LoadSqlData(" select * from Ap_balance_6_col   where ac_code Like  '" & (RSCIn_M.Fields("Ac_Code").Value) & "%' ", RSCkk)
        With RSCkk
            Do Until .EOF = True
                CNN.Execute("insert into Ap_Rpt_Incon_Detail (Rpt_ID , Ac_Code , Ac_Name  , Amt_Last_M_Dr  , Amt_Last_M_Cr , Amt_M_Dr , Amt_M_Cr , Amt_Q_Dr , Amt_Q_Cr , Amt_Y_Dr , Amt_Y_Cr , Ac_Code_Parent) Values ( '" & CStr((RSCIn_M.Fields("Rpt_ID").Value)) & "' ,  '" & CStr((.Fields("Ac_Code").Value)) & "' ,  N'" & CStr((.Fields("Ac_Name").Value)) & "'  ,  " & CDbl((.Fields("Amt_Last_M_Dr").Value)) & " , " & CDbl((.Fields("Amt_Last_M_Cr").Value)) & " ,  " & CDbl((.Fields("amt_dr").Value)) & " , " & CDbl((.Fields("amt_cr").Value)) & " , " & CDbl((.Fields("Quarter_dr").Value)) & " ,   " & CDbl((.Fields("Quarter_cr").Value)) & " , " & CDbl((.Fields("Rem_dr").Value)) & " , " & CDbl((.Fields("Rem_cr").Value)) & " , " & CStr((RSCIn_M.Fields("Ac_Code").Value)) & ") ")
                CNN.Execute("update  Ap_Rpt_Income_Item set Amt_Last_M_Dr  =  Amt_Last_M_Dr+" & CDbl((.Fields("Amt_Last_M_Dr").Value)) & " , Amt_Last_M_Cr  =  Amt_Last_M_Cr+" & CDbl((.Fields("Amt_Last_M_Cr").Value)) & " , amt_M_dr  =  amt_M_dr+" & CDbl((.Fields("amt_dr").Value)) & " , amt_M_cr  = amt_M_cr+" & CDbl((.Fields("amt_cr").Value)) & "  , amt_Q_dr  = amt_Q_dr+ " & CDbl(CDbl((.Fields("Quarter_dr").Value))) & " , amt_Q_cr  = amt_Q_cr+" & CDbl(CDbl((.Fields("Quarter_cr").Value))) & " , amt_y_dr  =  amt_y_dr+" & CDbl((.Fields("Rem_dr").Value)) & " , amt_y_cr  = amt_y_cr+" & CDbl((.Fields("Rem_cr").Value)) & "    where ac_code Like  '" & (RSCIn_M.Fields("Ac_Code").Value) & "%' And  Rpt_Type = 'Out' ")
                .MoveNext()
            Loop
        End With

    End Sub

    Private Sub SelectLoad()
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
    End Sub
    Private Sub LoadHalfYear()
        Dim s As Double = M1.SelectedIndex + 1
        Dim x As Double = M2.SelectedIndex + 1
        MdStartDate = Format(CDate("01/" & s & "/" & Year(Hyy.Value)), "dd-MM-yyyy")
        MdToDate = Format(CDate("01/" & x & "/" & Year(Hyy.Value)), "dd-MM-yyyy")
        Dim SM As Date = DateAdd("M", CDbl(1), MdToDate)
        Dim SD As Date = DateAdd("d", CDbl(-1), SM)
        MdToDate = Format(CDate(SD), "dd-MM-yyyy")
        If M1.SelectedIndex < M2.SelectedIndex Then
            Lb.Text = "ປະຈຳເດືອນ " & M1.Text & " ຫາ " & M2.Text & "/" & yy.Text
        Else
            Lb.Text = "ປະຈຳເດືອນ " & M1.Text & "/" & yy.Text
        End If
        'Dim s As Double = M1.SelectedIndex + 1
        'Dim x As Double = M2.SelectedIndex + 1
        'MdStartDate = Format(CDate("01/" & s & "/" & Year(Hyy.Value)), "dd-MM-yyyy")
        'MdToDate = Format(CDate("01/" & x & "/" & Year(Hyy.Value)), "dd-MM-yyyy")
        'Dim SM As Date = DateAdd("M", CDbl(1), MdToDate)
        'Dim SD As Date = DateAdd("d", CDbl(-1), SM)
        'MdToDate = Format(CDate(SD), "dd-MM-yyyy")

        'Lb.Text = "ປະຈຳເດືອນ " & yy.Text
        L5.Text = MdStartDate & " => " & MdToDate


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
        'MdStartDate = Format(CDate(Ds.Value), "dd-MM-yyyy")
        'MdToDate = Format(CDate(Dt.Value), "dd-MM-yyyy")
        'Lb.Text = "ແຕ່ວັນທີ " & MdStartDate & " ຫາວັນທີ " & MdToDate
        'L5.Text = MdStartDate & " => " & MdToDate
    End Sub
    Private Sub LoadMonth()
        '---------------------------------
        If DMonth.Text = "ມັງກອນ" Then
            MdStartDate = Format(CDate("01/01/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/01/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ມັງກອນ"
            'LngId = "7001" : CallLngStr() : MonthLetter1 = LngStr
        ElseIf DMonth.Text = "ກຸມພາ" Then
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
        ElseIf DMonth.Text = "ມີນາ" Then
            MdStartDate = Format(CDate("01/03/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/03/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ມີນາ"
        ElseIf DMonth.Text = "ເມສາ" Then
            MdStartDate = Format(CDate("01/04/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/04/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ເມສາ"
        ElseIf DMonth.Text = "ພຶດສະພາ" Then
            MdStartDate = Format(CDate("01/05/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/05/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ພຶດສະພາ"
        ElseIf DMonth.Text = "ມີຖຸນາ" Then
            MdStartDate = Format(CDate("01/06/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/06/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ມີຖຸນາ"
        ElseIf DMonth.Text = "ກໍລະກົດ" Then
            MdStartDate = Format(CDate("01/07/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/07/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ກໍລະກົດ"
        ElseIf DMonth.Text = "ສິງຫາ" Then
            MdStartDate = Format(CDate("01/08/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/08/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ສິງຫາ"
        ElseIf DMonth.Text = "ກັນຍາ" Then
            MdStartDate = Format(CDate("01/09/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/09/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ກັນຍາ"
        ElseIf DMonth.Text = "ຕຸລາ" Then
            MdStartDate = Format(CDate("01/10/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/10/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ຕຸລາ"
        ElseIf DMonth.Text = "ພະຈິກ" Then
            MdStartDate = Format(CDate("01/11/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/11/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ພະຈິກ"
        ElseIf DMonth.Text = "ທັນວາ" Then
            MdStartDate = Format(CDate("01/12/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/12/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ທັນວາ"
        End If
        '-----------------





        Dim s1, s2, s3 As String
        LngId = 7013 + DMonth.SelectedIndex : CallLngStr() : s2 = LngStr
        LngId = 7070 : CallLngStr() : s1 = LngStr
        LngId = 7071 : CallLngStr() : s3 = LngStr
        Lb.Text = s1 & " " & s2 & "/ " & Year(MdToDate)
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub

    Private Sub LoadPeriod()

        If Period.SelectedIndex = 0 Then
            MdStartDate = Format(CDate("01/01/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/03/" & Year(Pyy.Value)), "dd-MM-yyyy")
            Lb.Text = "ປະຈຳງວດ " & "1" & " ປີ " & Pyy.Text
        ElseIf Period.SelectedIndex = 1 Then
            MdStartDate = Format(CDate("01/04/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/06/" & Year(Pyy.Value)), "dd-MM-yyyy")
            Lb.Text = "ປະຈຳງວດ " & "2" & " ປີ " & Pyy.Text
        ElseIf Period.SelectedIndex = 2 Then
            MdStartDate = Format(CDate("01/07/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/09/" & Year(Pyy.Value)), "dd-MM-yyyy")
            Lb.Text = "ປະຈຳງວດ " & "3" & " ປີ " & Pyy.Text
        ElseIf Period.SelectedIndex = 3 Then
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
    Private Sub LoadReport()
        Dim RPT_ID As String
        RPT_ID = " "
        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As    RptSjUd ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
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
        LngId = "7055" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"

        Call LoadLoGO()

        'CDate(DateAdd(DateInterval.Month, MdStartDate, 1)
        Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
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

        SLF = "SELECT    " & mformat & "  as mformat  , " & MuLngRpt & "  *   FROM Ap_Rpt_Income  "

        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            Dim ny, ly As String
            ny = CDbl(Year(MdStartDate))
            ly = CDbl(Year(MdStartDate)) - 1
            .Open(" " & SLF & " where grp<>'' order by CNT asc  ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        Dim FrmPreview As New FmPreview : FrmClosing()
        'Dim Rpt As New CryLOGO
        Dim Rpt As New CryRpt_Income
        If MdShowLOGO = 1 Then
            Rpt.Subreports(0).SetDataSource(RsLOGO)
        End If
        Rpt.SetDataSource(Rs)
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.DisplayGroupTree = False
        'FrmPreview.MdiParent = FmMain
        FrmPreview.WindowState = FormWindowState.Maximized
        FrmPreview.Show()
        FrmPreview.Focus()
    End Sub

    Private Sub LoadReportItem()
        Dim RPT_ID As String
        RPT_ID = " "
        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As    RptSjUd ,"
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
        LngId = "7055" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
        If RM.Checked = True Then
            LngId = "7060" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7061" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        ElseIf RP.Checked = True Then
            LngId = "7062" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7063" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        ElseIf Rhalf.Checked = True Then
            LngId = "7080" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7081" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        ElseIf RY.Checked = True Then
            LngId = "7064" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7065" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        End If
        SLF = "SELECT  " & MuLngRpt & "  *   FROM Ap_Rpt_Incon_Detail  "
        Call LoadLoGO()
        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            n_L_y = " N'" & LL5.Text & "' As Now_Year , N'" & LL6.Text & "' As Last_Year ,  "
            .Open(" " & SLF & "   ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim Rpt As New CryRpt_Income_Itemxx
        If MdShowLOGO = 1 Then
            Rpt.Subreports(0).SetDataSource(RsLOGO)
        End If
        Rpt.SetDataSource(Rs)
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.DisplayGroupTree = False
        FrmPreview.Show()
        FrmPreview.Focus()
    End Sub



    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Call Close()
    End Sub

    Private Sub FmRpt_Income_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Call MdiCNum()
    End Sub

    Private Sub FmRpt_BLS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call loadOffice_User()
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
        M2.SelectedIndex = M1.SelectedIndex
        Call SelectLoad()
        'SetControlText(Me)

        RM.Checked = True
        Button2.Text = "Export"
    End Sub
    Private Sub loadOffice_User()
        Off_Usr.Items.Clear()
        LoadSqlData("select sub_id , off_add2  from  Ap_office  Order by sub_id", RSC)
        With RSC
            Do Until .EOF = True
                'MsgBox("ghf")
                Off_Usr.Items.Add((.Fields("sub_id").Value) & " " & (.Fields("off_add2").Value))
                .MoveNext()
            Loop
        End With
        Off_Usr.Text = FmLogin.Sub_Company.Text
    End Sub

    Private Sub RM_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RM.CheckedChanged
        Call SelectLoad()
    End Sub

    Private Sub RP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RP.CheckedChanged
        Call SelectLoad()
    End Sub

    Private Sub RD_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RD.CheckedChanged
        Call SelectLoad()
    End Sub

    Private Sub RY_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RY.CheckedChanged
        Call SelectLoad()
    End Sub

    Private Sub Period_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Period.SelectedIndexChanged
        Call SelectLoad()
    End Sub

    Private Sub yy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yy.ValueChanged
        Call SelectLoad()
    End Sub

    Private Sub Ds_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ds.ValueChanged
        Dt.Value = Ds.Value
        Call SelectLoad()
    End Sub

    Private Sub DMonth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DMonth.SelectedIndexChanged
        Call SelectLoad()
    End Sub

    Private Sub Myy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Myy.ValueChanged
        Call SelectLoad()
    End Sub

    Private Sub Pyy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Pyy.ValueChanged
        Call SelectLoad()
    End Sub

    Private Sub Dt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Dt.ValueChanged
        Call SelectLoad()
    End Sub

    Private Sub Toyy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Toyy.ValueChanged
        Call SelectLoad()
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        FmIncome_Old.ShowDialog()
        FmIncome_Old.Focus()


    End Sub

    Private Sub RdIn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub RdOut_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ButtoClick_Lastn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

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
                LL6.Text = "ສະສົດເດືອນຜ່ານມາ" & vbCrLf & "ຮອດ " & Format(MdStartDate_Last, "MM/yyyy")
            End If
            LL5.Text = Format(MdStartDate, "MM/yyyy")
        ElseIf RP.Checked = True Then
            MdStartDate_Last = DateAdd(DateInterval.Month, -3, MdStartDate)
            MdToDate_Last = DateAdd(DateInterval.Day, -1, MdStartDate)
            LL6.Text = ""
            If Period.SelectedIndex > 0 Then
                LL6.Text = "ງວດຜ່ານມາ" & vbCrLf & "ຮອດ " & Format(MdStartDate_Last, "MM/yyyy")
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

    Private Sub Rhalf_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Rhalf.CheckedChanged

        Call SelectLoad()
    End Sub

    Private Sub half_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles M1.SelectedIndexChanged
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

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RT.CheckedChanged
        Call SelectLoad()
    End Sub

    Private Sub Ct_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ct.SelectedIndexChanged
        Call SelectLoad()
    End Sub

    Private Sub yyt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yyt.ValueChanged
        Call SelectLoad()
    End Sub

    Private Sub Hyy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Hyy.ValueChanged
        Call SelectLoad()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Click_Last()
        Call SelectLoad()
        '============
        'CNN.Execute("update  Ap_Rpt_Income_Item set Last_Amt_dr  =  0 , Last_Amt_Cr  =  0 , amt_dr  = 0 , amt_cr  = 0  ")
        Dim saa As String = "update  Ap_Rpt_Income_Item set Last_Amt_dr  =  0 , Last_Amt_Cr  =  0 , amt_dr  = 0 , amt_cr  = 0 "
        CNN.Execute(saa)

        CNN.Execute("update Ap_Rpt_Income set  Last_Amt  = 0 , Amt  = 0    ")
        CNN.Execute("DELETE FROM Ap_Rpt_Incon_Detail ")
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        ChangBalance()
        ' ''    CNN.Execute("DELETE  Ap_balance_6_col ")
        ' ''    CNN.Execute("DELETE FROM Ap_balance_6 ")
        ' ''    CNN.Execute(" insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  from Ap_balance_TB where  date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  order by Ac_Code asc ")
        ' ''    Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
        ' ''    Dim ppp As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ' ''          " select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from Ap_balance_TB  WHERE  date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "'   group BY ac_code"
        ' ''    CNN.Execute(ppp)
        ' ''    'Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
        ' ''    'Dim ppp As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ' ''    '      " select ac_code , sum(open_amt_dr), sum(open_amt_cr)  , sum(amt_dr)   ,sum(amt_cr)    from Ap_balance_TB  WHERE  date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'   group BY ac_code"
        ' ''    'CNN.Execute(ppp)

        ' ''    CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ' ''" select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr)as open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")
        ' ''    CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        ' ''    CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")

        ' ''    CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + amt_dr) - (open_amt_cr + amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr + amt_cr) ")
        ' ''    CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr + amt_cr) - (open_amt_dr + amt_dr) where (open_amt_cr + amt_cr) >= (open_amt_dr + amt_dr) ")
        ' ''    CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")

        SelcectInLast()
        UpdateIInLast()
        SelectOutLast()
        UpdateOut()
        Update_Sum()


        If RD.Checked = True Then
            Dim s1, s2 As String
            LngId = 7027 : CallLngStr() : s2 = LngStr
            LngId = 7072 : CallLngStr() : s1 = LngStr
            Lb.Text = s1 & " " & Ds.Text & " " & s2 & " " & Dt.Text
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
        ElseIf Rhalf.Checked = True Then
            Dim s1, s2 As String
            'LngId = 7078 + M1.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7089 : CallLngStr() : s1 = LngStr

            LngId = 7027 : CallLngStr() : s2 = LngStr
            'Lb.Text = s1 & " " & M1.SelectedIndex + 1 & " " & s2 & " " & M2.SelectedIndex + 1 & " (" & Year(MdToDate) & " )"


        ElseIf RT.Checked = True Then
            Dim s1, s7 As String
            LngId = 7049 : CallLngStr() : s1 = LngStr

            If Ct.SelectedIndex = 0 Then
                s7 = "0" & Ct.SelectedIndex + 1
                LngId = 7078 : CallLngStr()
                Lb.Text = LngStr & " " & yyt.Text
            Else
                s7 = Ct.SelectedIndex + 1
                LngId = 7079 : CallLngStr()
                Lb.Text = LngStr & " " & yyt.Text
            End If


        ElseIf RY.Checked = True Then
            Dim s1, s2, s3 As String
            LngId = 7074 + Period.SelectedIndex : CallLngStr() : s2 = LngStr
            LngId = 7062 : CallLngStr() : s1 = LngStr
            LngId = 7064 : CallLngStr() : s3 = LngStr
            Lb.Text = s3 & " " & yy.Text
        End If


        If CheckBox3.Checked = True Then
            CNN.Execute("Update Ap_Rpt_Income set Amt=Amt+Last_Amt")
        End If
        '========1 =========
        CNN.Execute("update Ap_Rpt_Income set amt=(select isnull(sum(amt),0) as amt  from Ap_Rpt_Income where (rpt_ID='1.1.1' or  rpt_ID='1.1.1.1'  or  rpt_ID='1.1.1.2'  or  rpt_ID='1.1.1.2.1'  or  rpt_ID='1.1.1.2.2' )) where  rpt_ID='1' ")

        '========2 =========
        CNN.Execute("update Ap_Rpt_Income set amt=(select isnull(sum(amt),0) as amt  from Ap_Rpt_Income where (rpt_ID='1.1.1.3' or  rpt_ID='1.1.1.4'  or  rpt_ID='1.1.2'  or  rpt_ID='1.1.2.5'  or  rpt_ID='1.13' )) where  rpt_ID='1.1.1.2.3' ")


        Dim MI As String = "delete AP_Sum "
        MI = MI & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,amt,0 from Ap_Rpt_Income     where rpt_ID='1'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  'I'  ,0 ,amt from Ap_Rpt_Income     where rpt_ID='1.1.1.2.3'  "
        MI = MI & "  update AP_Sum set amt=(select sum(amt2)-sum(amt3) from AP_Sum where (rpt_ID='1' or rpt_ID='I')  ) where rpt_ID='I'   "
        MI = MI & "  delete  AP_Sum  where rpt_ID='1'   "
        MI = MI & " Update Ap_Rpt_Income set Ap_Rpt_Income.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Income    where Ap_Rpt_Income.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(MI)
        '========1,1=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.1.1' and  rpt_ID<='1.1.5') where  rpt_ID='1.1' ")
        '========1,2=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.2.1' and  rpt_ID<='1.2.5') where  rpt_ID='1.2' ")
        '========I=========
        '========Sum I =========
        Dim SI As String = "delete AP_Sum "
        SI = SI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,amt,0 from Ap_Rpt_Income     where rpt_ID='1.1' "
        SI = SI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  'I'  ,0 ,amt from Ap_Rpt_Income     where rpt_ID='1.2'  "
        SI = SI & " update AP_Sum set amt=(select sum(amt2)-sum(amt3) from AP_Sum where (rpt_ID='1.1' or rpt_ID='I')  ) where rpt_ID='I' "
        SI = SI & " delete  AP_Sum  where rpt_ID='1.1' "
        SI = SI & " Update Ap_Rpt_Income set Ap_Rpt_Income.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Income    where Ap_Rpt_Income.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(SI)
        '========3=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.3.1' and  rpt_ID<='1.3.2') where  rpt_ID='1.3' ")
        '========4=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.4.1' and  rpt_ID<='1.4.2') where  rpt_ID='1.4'")
        '========5========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID='1.5.1' ) where  rpt_ID='1.5' ")
        '========6======== 
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.6.1' and  rpt_ID<='1.6.2') where  rpt_ID='1.6' ")
        '========7======== 
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID='1.7.1') where  rpt_ID='1.7' ")
        '========8======== 
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.8.1' and  rpt_ID<='1.8.2') where  rpt_ID='1.8'")
        '========9======== 
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.9.1' and  rpt_ID<='1.9.5') where  rpt_ID='1.9' ")
        '========10======= 
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.10.1' and  rpt_ID<='1.10.5') where  rpt_ID='1.10'")


        '========11=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.11.1' and  rpt_ID<='1.11.2') where  rpt_ID='1.11'")

        '========12=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.12.1' and  rpt_ID<='1.12.2') where  rpt_ID='1.12'")

        '========13=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.13.1' and  rpt_ID<='1.13.2') where  rpt_ID='1.13'")
        '========14=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.14.1' and  rpt_ID<='1.14.2') where  rpt_ID='1.14'")
        '========15=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.15.1' and  rpt_ID<='1.15.5') where  rpt_ID='1.15' ")
        '========16========= 
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.16.1' and  rpt_ID<='1.16.2') where  rpt_ID='1.16'")
        '========17=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.17.1' and  rpt_ID<='1.17.2') where  rpt_ID='1.17'")
        '========18=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.18.1' and  rpt_ID<='1.18.4') where  rpt_ID='1.18' ")
        '========19=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.19.1' and  rpt_ID<='1.19.2') where  rpt_ID='1.19'")
        '========20=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.20.1' and  rpt_ID<='1.20.4') where  rpt_ID='1.20'")
        '========21=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where rpt_ID>='1.21.1' and  rpt_ID<='1.21.3') where  rpt_ID='1.21'")
        CNN.Execute("UPDATE Ap_Rpt_Income set amt=0 where amt is null")
        '========Sum III=========
        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where (rpt_ID='1.15' or  rpt_ID='1.16' or  rpt_ID='1.17' or  rpt_ID='1.18' or  rpt_ID='1.19' or  rpt_ID='1.20')) where  rpt_ID='III' ")

        '========Sum II =========
        Dim SII As String = "delete AP_Sum "
        SII = SII & " insert into AP_Sum(Rpt_ID,amt2,amt3) select 'II' ,amt,0 from Ap_Rpt_Income     where rpt_ID='I' "
        SII = SII & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  'II'  ,0 ,amt from Ap_Rpt_Income     where (rpt_ID='3' or rpt_ID='4' or rpt_ID='5' or rpt_ID='6' or rpt_ID='7' or rpt_ID='8' or rpt_ID='9' or rpt_ID='10' or rpt_ID='11' or rpt_ID='12' or rpt_ID='13' or rpt_ID='14') "
        SII = SII & " update AP_Sum set amt=(select sum(amt2)-sum(amt3) from AP_Sum where (rpt_ID='II' or rpt_ID='3' or rpt_ID='4' or rpt_ID='5' or rpt_ID='6' or rpt_ID='7' or rpt_ID='8' or rpt_ID='9' or rpt_ID='10' or rpt_ID='11' or rpt_ID='12' or rpt_ID='13' or rpt_ID='14' )  ) where rpt_ID='II' "
        SII = SII & " Update Ap_Rpt_Income set Ap_Rpt_Income.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Income    where Ap_Rpt_Income.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(SII)
        'SII = SII & " delete  AP_Sum  where rpt_ID='I' "

        CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where (rpt_ID='I' or  rpt_ID='1.3' or  rpt_ID='1.4' or  rpt_ID='1.5' or  rpt_ID='1.6' or  rpt_ID='1.7' or  rpt_ID='1.8' or  rpt_ID='1.9' or  rpt_ID='1.10' or  rpt_ID='1.11' or  rpt_ID='1.12' or  rpt_ID='1.13'  or  rpt_ID='1.14' )) where  rpt_ID='II' ")
        ''========Sum III=========
        '========Sum IV=========
        Dim IVV As String = "delete AP_Sum "
        IVV = IVV & " insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,amt,0 from Ap_Rpt_Income     where rpt_ID='II' "
        IVV = IVV & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  'IV'  ,0 ,amt from Ap_Rpt_Income     where rpt_ID='III'  "
        IVV = IVV & " update AP_Sum set amt=(select sum(amt2)-sum(amt3) from AP_Sum where (rpt_ID='II' or rpt_ID='IV')  ) where rpt_ID='IV' "
        IVV = IVV & " delete  AP_Sum  where rpt_ID='II' "
        IVV = IVV & " Update Ap_Rpt_Income set Ap_Rpt_Income.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Income    where Ap_Rpt_Income.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(IVV)

        '========Sum  V=========
        Dim VV As String = "delete AP_Sum "
        VV = VV & " insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,amt,0 from Ap_Rpt_Income     where rpt_ID='IV' "
        VV = VV & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  'V'  ,0 ,amt from Ap_Rpt_Income     where rpt_ID='1.21'  "
        VV = VV & " update AP_Sum set amt=(select sum(amt2)-sum(amt3) from AP_Sum where (rpt_ID='IV' or rpt_ID='V')  ) where rpt_ID='V' "
        VV = VV & " delete  AP_Sum  where rpt_ID='IV' "
        VV = VV & " Update Ap_Rpt_Income set Ap_Rpt_Income.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Income    where Ap_Rpt_Income.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(VV)

        If CheckBox1.Checked = False Then
            Call LoadReport_Export()
        Else
            Call LoadReportItem_Export()
        End If
    End Sub
    Private Sub LoadReport_Export()
        Dim RPT_ID As String
        RPT_ID = " "
        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As    RptSjUd ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
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
        LngId = "7055" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"

        Call LoadLoGO()

        'CDate(DateAdd(DateInterval.Month, MdStartDate, 1)
        Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
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

        SLF = "SELECT  " & MuLngRpt & "  *   FROM Ap_Rpt_Income  "

        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            Dim ny, ly As String
            ny = CDbl(Year(MdStartDate))
            ly = CDbl(Year(MdStartDate)) - 1
            .Open(" " & SLF & " where grp<>'' order by CNT asc  ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        Dim FrmPreview As New FmPreview : FrmClosing()
        'Dim Rpt As New CryLOGO
        Dim Rpt As New CryRpt_Income
        If MdShowLOGO = 1 Then
            Rpt.Subreports(0).SetDataSource(RsLOGO)
        End If
        Rpt.SetDataSource(Rs)
        Rpt.Refresh()
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.ExportReport()
        FrmPreview = Nothing

        'Rpt.SetDataSource(Rs)
        'FrmPreview.ReportViewer.ReportSource = Rpt
        'FrmPreview.ReportViewer.DisplayGroupTree = False
        ''FrmPreview.MdiParent = FmMain
        'FrmPreview.WindowState = FormWindowState.Maximized
        'FrmPreview.Show()
        'FrmPreview.Focus()
    End Sub

    Private Sub LoadReportItem_Export()
        Dim RPT_ID As String
        RPT_ID = " "
        Call Office()
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As    RptSjUd ,"
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
        LngId = "7055" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
        If RM.Checked = True Then
            LngId = "7060" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7061" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        ElseIf RP.Checked = True Then
            LngId = "7062" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7063" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        ElseIf Rhalf.Checked = True Then
            LngId = "7080" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7081" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        ElseIf RY.Checked = True Then
            LngId = "7064" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7065" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        End If
        SLF = "SELECT  " & MuLngRpt & "  *   FROM Ap_Rpt_Incon_Detail  "
        Call LoadLoGO()
        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            n_L_y = " N'" & LL5.Text & "' As Now_Year , N'" & LL6.Text & "' As Last_Year ,  "
            .Open(" " & SLF & "   ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim Rpt As New CryRpt_Income_Itemxx
        If MdShowLOGO = 1 Then
            Rpt.Subreports(0).SetDataSource(RsLOGO)
        End If
        Rpt.SetDataSource(Rs)
        Rpt.Refresh()
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.ExportReport()
        FrmPreview = Nothing

        'Rpt.SetDataSource(Rs)
        'FrmPreview.ReportViewer.ReportSource = Rpt
        'FrmPreview.ReportViewer.DisplayGroupTree = False
        'FrmPreview.Show()
        'FrmPreview.Focus()
    End Sub
End Class