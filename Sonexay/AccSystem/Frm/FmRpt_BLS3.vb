Public Class FmRpt_BLS
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
    Dim RsOpen As New ADODB.Recordset
    Dim RsOpenMonth As New ADODB.Recordset
    Dim RsRpt As New ADODB.Recordset
    Dim VOpenDate As Date
    Dim RptNme As String
    Dim RSC12 As New ADODB.Recordset
    Dim RSCIn_M As New ADODB.Recordset


    Private Sub ChangBalance()
        New_Code = "3901000"
        Code_Dr = "4"
        Code_Cr = "5"
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
        CNN.Execute(Insr)
    End Sub
    Private Sub LoadAPLOAN()
        Dim MDWRITEOFF As String = " AND WRITEOFF=N'N' "
        ' ====AMT KIP  TT===
        Dim AMT1 As String = "select isnull(sum(Principle_LAK),0) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1 " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'1. ສິນເຊື່ອປົກກະຕິ (A)' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Call LoadSqlData(AMT1, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim KW1 As String = "Update TEM_F04 set  AMT=" & RSC.Fields("AA").Value & " where  rpt_ID='1.5.1'  "
                CNN.Execute(KW1)
                RSC.MoveNext()
            End While
        End If
        Dim AMT2 As String = "select isnull(sum(Principle_LAK),0) as AA,BUSINESSTYPEDESC  from AP_Loan  where 1=1  " & MDWRITEOFF & "  and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Call LoadSqlData(AMT2, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim KW1 As String = "Update TEM_F04 set  AMT=" & RSC.Fields("AA").Value & " where  rpt_ID='1.5.2'  "
                CNN.Execute(KW1)
                RSC.MoveNext()
            End While
        End If
        Dim AMT3 As String = "select isnull(sum(Principle_LAK),0) as AA  from AP_Loan  where 1=1  " & MDWRITEOFF & "  and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   (loan_grade=N'3. ສິນເຊື່ອຕໍ່າກວ່າມາດຕະຖານ (C)' or loan_grade=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)' or loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)')   "
        Call LoadSqlData(AMT3, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim KW1 As String = "Update TEM_F04 set  AMT=" & RSC.Fields("AA").Value & " where  rpt_ID='1.5.3'  "
                CNN.Execute(KW1)
                RSC.MoveNext()
            End While
        End If
    End Sub
    Private Sub UPP_INTER()
   
        Dim KK As New ADODB.Recordset
        Dim aa As String = " SELECT  * from AP_MM where month(MM)='" & Month(DMY.Value) & "' and  year(MM)='" & Year(DMY.Value) & "'   "
        Call LoadSqlData(aa, KK)
        If KK.RecordCount = 0 Then
            'Dim HH As String = "INSERT INTO AP_MM (MM,A,B,C,D,E,Amt) values('" & Format(DMY.Value, "yyyy/MM/dd") & "'," & CDbl(txtA.Text) & "," & CDbl(txtB.Text) & "," & CDbl(txtC.Text) & "," & CDbl(txtD.Text) & "," & CDbl(txtE.Text) & ",0 ) "
            'CNN.Execute(HH)
            Dim HH As String = "INSERT INTO AP_MM (MM,A,B,C,D,E,Amt) values('" & Format(DMY.Value, "yyyy/MM/dd") & "',N'" & (txtA.Text) & "',N'" & (txtB.Text) & "',N'" & (txtC.Text) & "',N'" & (txtD.Text) & "',N'" & (txtE.Text) & "',0 ) "
            CNN.Execute(HH)
        Else
            'CNN.Execute("UPDATE AP_MM set A=" & CDbl(txtA.Text) & ",b=" & CDbl(txtB.Text) & ",c=" & CDbl(txtC.Text) & ",d=" & CDbl(txtD.Text) & ", e=" & CDbl(txtE.Text) & "  where month(MM)='" & Month(DMY.Value) & "' and  year(MM)='" & Year(DMY.Value) & "'  ")
            CNN.Execute("UPDATE AP_MM set A=N'" & (txtA.Text) & "',b=N'" & (txtB.Text) & "',c=N'" & (txtC.Text) & "',d=N'" & (txtD.Text) & "', e=N'" & (txtE.Text) & "'  where month(MM)='" & Month(DMY.Value) & "' and  year(MM)='" & Year(DMY.Value) & "'  ")

        End If
        'MsgBox("OK")
    End Sub
    Private Sub Incom()
        'Click_Last()
        Call selectLoad()
        '============
        'CNN.Execute("update  Ap_Rpt_Income_Item set Last_Amt_dr  =  0 , Last_Amt_Cr  =  0 , amt_dr  = 0 , amt_cr  = 0  ")
        Dim sa1 As String = "update  Ap_Rpt_Income_Item set Last_Amt_dr  =  0 , Last_Amt_Cr  =  0 , amt_dr  = 0 , amt_cr  = 0 "
        CNN.Execute(sa1)

        CNN.Execute("update Ap_Rpt_Income set  Last_Amt  = 0 , Amt  = 0    ")
        CNN.Execute("DELETE FROM Ap_Rpt_Incon_Detail ")
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        ' ''    ChangBalance11()
        ' ''    CNN.Execute("DELETE  Ap_balance_6_col ")
        ' ''    CNN.Execute("DELETE FROM Ap_balance_6 ")
        ' ''    'CNN.Execute(" insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  from Ap_balance_TB where  date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  order by Ac_Code asc ")

        ' ''Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
        ' ''    Dim ppp As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ' ''          " select ac_code , sum(open_amt_dr), sum(open_amt_cr)  , sum(amt_dr)   ,sum(amt_cr)    from Ap_balance_TB  WHERE  date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'   group BY ac_code"
        ' ''    CNN.Execute(ppp)
        ' ''    'Dim ppp As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ' ''    '    " select ac_code , (sum(open_amt_dr)+sum(amt_dr)) as amt_dr , (sum(open_amt_cr)+sum(amt_cr))as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from Ap_balance_TB  WHERE  date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "'   group BY ac_code"
        ' ''    'CNN.Execute(ppp)
        ' ''    CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ' ''" select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr)as open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")
        ' ''    CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        ' ''    CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")

        ' ''    CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + amt_dr) - (open_amt_cr + amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr + amt_cr) ")
        ' ''    CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr + amt_cr) - (open_amt_dr + amt_dr) where (open_amt_cr + amt_cr) >= (open_amt_dr + amt_dr) ")
        ' ''    CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")

        SelcectInLast11()
        UpdateIInLast11()
        SelectOutLast11()
        UpdateOut11()
        'Update_Sum11()


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

        'CNN.Execute("update Ap_Rpt_Income set amt=(select sum(amt) from Ap_Rpt_Income where (rpt_ID='I' or  rpt_ID='1.2.1' or  rpt_ID='1.2.1.1.3' or  rpt_ID='1.2.2' or  rpt_ID='1.2.2.2.2' or  rpt_ID='1.2.3' or  rpt_ID='1.2.4' or  rpt_ID='1.2.4.3' or  rpt_ID='1.2.10' or  rpt_ID='1.1.11' or  rpt_ID='1.1.12' or  rpt_ID='1.1.13' or  rpt_ID='1.1.14')) where  rpt_ID='II' ")
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

    End Sub
    Private Sub CALLC()
        Dim MDWRITEOFF As String = " AND WRITEOFF=N'N' "

        '===========   N'1. ສິນເຊື່ອປົກກະຕິ (A)'===============
        '=========ACCNO=====
        Dim A1 As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan where 1=1 " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and  loan_grade=N'1. ສິນເຊື່ອປົກກະຕິ (A)' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Call LoadSqlData(A1, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim K1 As String = "Update RPT_F04 set ACCNO=" & RSC.Fields("AA").Value & "  where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and Grp_Nm=N'1. ສິນເຊື່ອປົກກະຕິ (A)' "
                CNN.Execute(K1)
                RSC.MoveNext()
            End While
        End If
        '=========ACC ALl=====
        Dim A1_Cust As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan where 1=1 " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and  loan_grade=N'1. ສິນເຊື່ອປົກກະຕິ (A)' group by BUSINESSTYPEDESC  , Cust_ID order by BUSINESSTYPEDESC "
        Call LoadSqlData(A1_Cust, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim K1 As String = "Update RPT_F04 set  ACC=ACC+" & 1 & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and Grp_Nm=N'1. ສິນເຊື່ອປົກກະຕິ (A)' "
                CNN.Execute(K1)
                RSC.MoveNext()
            End While
        End If
        ' ====ACC W===
        Dim W1 As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and  loan_grade=N'1. ສິນເຊື່ອປົກກະຕິ (A)' and GENDER=N'F' group by BUSINESSTYPEDESC, Cust_ID  order by BUSINESSTYPEDESC "
        Call LoadSqlData(W1, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim KW1 As String = "Update RPT_F04 set  ACC_W=ACC_W+" & 1 & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and Grp_Nm=N'1. ສິນເຊື່ອປົກກະຕິ (A)' "
                CNN.Execute(KW1)
                RSC.MoveNext()
            End While
        End If

        ' ====AMT KIP  TT===
        Dim AMT1 As String = "select isnull(sum(Principle_LAK),0) as AA, isnull(sum(Provision_Amt),0) as BB, BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'1. ສິນເຊື່ອປົກກະຕິ (A)' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Call LoadSqlData(AMT1, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim KW1 As String = "Update RPT_F04 set  AMT=" & RSC.Fields("AA").Value & " , Dep_Amt=" & RSC.Fields("BB").Value & "  where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and Grp_Nm=N'1. ສິນເຊື່ອປົກກະຕິ (A)' "
                CNN.Execute(KW1)
                RSC.MoveNext()
            End While
        End If
        ' ====AMT KIP  W===
        Dim AMT_W As String = "select isnull(sum(Principle_LAK),0) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'1. ສິນເຊື່ອປົກກະຕິ (A)' and GENDER=N'F' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Call LoadSqlData(AMT_W, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim KW1 As String = "Update RPT_F04 set  AMT_W=" & RSC.Fields("AA").Value & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and Grp_Nm=N'1. ສິນເຊື່ອປົກກະຕິ (A)' "
                CNN.Execute(KW1)
                RSC.MoveNext()
            End While
        End If

        ''=========== 2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)===============
        Dim A2 As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Call LoadSqlData(A2, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim K2 As String = "Update RPT_F04 set ACCNO=" & RSC.Fields("AA").Value & "  where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and Grp_Nm=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' "
                CNN.Execute(K2)
                RSC.MoveNext()
            End While
        End If
        '============
        Dim A2_AA As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' group by BUSINESSTYPEDESC, Cust_ID order by BUSINESSTYPEDESC "
        Call LoadSqlData(A2_AA, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim K2 As String = "Update RPT_F04 set  ACC=ACC+" & 1 & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and Grp_Nm=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' "
                CNN.Execute(K2)
                RSC.MoveNext()
            End While
        End If
        ' ====W===
        Dim W2 As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' and GENDER=N'F' group by BUSINESSTYPEDESC, Cust_ID order by BUSINESSTYPEDESC "
        Call LoadSqlData(W2, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim KW2 As String = "Update RPT_F04 set  ACC_W=ACC_W+" & 1 & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and Grp_Nm=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' "
                CNN.Execute(KW2)
                RSC.MoveNext()
            End While
        End If
        ' ====AMT KIP  TT===
        Dim AMT2 As String = "select isnull(sum(Principle_LAK),0) as AA,  isnull(sum(Provision_Amt),0) as BB, BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Call LoadSqlData(AMT2, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim KW1 As String = "Update RPT_F04 set  AMT=" & RSC.Fields("AA").Value & " , Dep_Amt=" & RSC.Fields("BB").Value & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and Grp_Nm=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' "
                CNN.Execute(KW1)
                RSC.MoveNext()
            End While
        End If
        ' ====AMT KIP  W===
        Dim AMT_W2 As String = "select isnull(sum(Principle_LAK),0) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' and GENDER=N'F' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Call LoadSqlData(AMT_W2, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim KW1 As String = "Update RPT_F04 set  AMT_W=" & RSC.Fields("AA").Value & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and Grp_Nm=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' "
                CNN.Execute(KW1)
                RSC.MoveNext()
            End While
        End If


        ''=========== 3. ສິນເຊື່ອຕໍ່າກວ່າມາດຕະຖານ (C)===============
        Dim A3 As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   left(loan_grade,1)=N'3' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Call LoadSqlData(A3, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim K3 As String = "Update RPT_F04 set ACCNO=" & RSC.Fields("AA").Value & "  where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and left(Grp_Nm,1)=N'3' "
                CNN.Execute(K3)
                RSC.MoveNext()
            End While
        End If
        '=======
        Dim A3_AAA As String = "select count(*) as AA,BUSINESSTYPEDESC,1 from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and    left(loan_grade,1)=N'3' group by BUSINESSTYPEDESC,Cust_ID  order by BUSINESSTYPEDESC "
        Call LoadSqlData(A3_AAA, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim K3 As String = "Update RPT_F04 set  ACC=ACC+" & 1 & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and left(Grp_Nm,1)=N'3' "
                CNN.Execute(K3)
                RSC.MoveNext()
            End While
        End If
        ' ====W===
        Dim W3 As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   left(loan_grade,1)=N'3' and GENDER=N'F' group by BUSINESSTYPEDESC,Cust_ID order by BUSINESSTYPEDESC "
        Call LoadSqlData(W3, RSC)

        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim KW3 As String = "Update RPT_F04 set  ACC_W=ACC_W+" & 1 & "  where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and left(Grp_Nm,1)=N'3' "
                CNN.Execute(KW3)
                RSC.MoveNext()
            End While
        End If
        ' ====AMT KIP  TT===
        Dim AMT3 As String = "select isnull(sum(Principle_LAK),0) as AA,  isnull(sum(Provision_Amt),0) as BB,  BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   left(loan_grade,1)=N'3' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Call LoadSqlData(AMT3, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim KW1 As String = "Update RPT_F04 set  AMT=" & RSC.Fields("AA").Value & "  , Dep_Amt=" & RSC.Fields("BB").Value & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and left(Grp_Nm,1)=N'3' "
                CNN.Execute(KW1)
                RSC.MoveNext()
            End While
        End If
        ' ====AMT KIP  W===
        Dim AMT_W3 As String = "select isnull(sum(Principle_LAK),0) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and    left(loan_grade,1)=N'3' and GENDER=N'F' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Call LoadSqlData(AMT_W3, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim KW1 As String = "Update RPT_F04 set  AMT_W=" & RSC.Fields("AA").Value & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and left(Grp_Nm,1)=N'3'' "
                CNN.Execute(KW1)
                RSC.MoveNext()
            End While
        End If

        ''=========== 4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)==============
        Dim A4 As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Call LoadSqlData(A4, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim K4 As String = "Update RPT_F04 set ACCNO=" & RSC.Fields("AA").Value & "  where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and Grp_Nm=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)' "
                CNN.Execute(K4)
                RSC.MoveNext()
            End While
        End If
        '=====
        Dim A4_AA As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)' group by BUSINESSTYPEDESC,Cust_ID order by BUSINESSTYPEDESC "
        Call LoadSqlData(A4_AA, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim K4 As String = "Update RPT_F04 set  ACC=ACC+" & 1 & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and Grp_Nm=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)' "
                CNN.Execute(K4)
                RSC.MoveNext()
            End While
        End If
        ' ====W===
        Dim W4 As String = "select count(*) as AA,BUSINESSTYPEDESC  from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)'   and GENDER=N'F' group by BUSINESSTYPEDESC,Cust_ID order by BUSINESSTYPEDESC "
        Call LoadSqlData(W4, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim KW4 As String = "Update RPT_F04 set  ACC_W=ACC_W+" & 1 & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and Grp_Nm=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)'  "
                CNN.Execute(KW4)
                RSC.MoveNext()
            End While
        End If
        ' ====AMT KIP  TT===
        Dim AMT4 As String = "select isnull(sum(Principle_LAK),0) as AA,  isnull(sum(Provision_Amt),0) as BB,  BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)'  group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Call LoadSqlData(AMT4, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim KW1 As String = "Update RPT_F04 set  AMT=" & RSC.Fields("AA").Value & "  , Dep_Amt=" & RSC.Fields("BB").Value & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and Grp_Nm=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)'  "
                CNN.Execute(KW1)
                RSC.MoveNext()
            End While
        End If
        ' ====AMT KIP  W===
        Dim AMT_W4 As String = "select isnull(sum(Principle_LAK),0) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   loan_grade=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)'  and GENDER=N'F' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Call LoadSqlData(AMT_W4, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim KW1 As String = "Update RPT_F04 set  AMT_W=" & RSC.Fields("AA").Value & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and Grp_Nm=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)'  "
                CNN.Execute(KW1)
                RSC.MoveNext()
            End While
        End If


        ''=========== 5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E))==============
        Dim A5 As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   (loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Call LoadSqlData(A5, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim K5 As String = "Update RPT_F04 set ACCNO=" & RSC.Fields("AA").Value & "  where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and (Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') "
                CNN.Execute(K5)
                RSC.MoveNext()
            End While
        End If

        Dim A5_AA As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   (loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') group by BUSINESSTYPEDESC,Cust_ID order by BUSINESSTYPEDESC "
        Call LoadSqlData(A5_AA, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim K5 As String = "Update RPT_F04 set  ACC=ACC+" & 1 & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and (Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') "
                CNN.Execute(K5)
                RSC.MoveNext()
            End While
        End If
        ' ====W===
        Dim W5 As String = "select count(*) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and  (loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)')   and GENDER=N'F' group by BUSINESSTYPEDESC,Cust_ID order by BUSINESSTYPEDESC "
        Call LoadSqlData(W5, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim KW5 As String = "Update RPT_F04 set  ACC_W=ACC_W+" & 1 & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and   (Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') "
                CNN.Execute(KW5)
                RSC.MoveNext()
            End While
        End If
        ' ====AMT KIP  TT===
        Dim AMT5 As String = "select isnull(sum(Principle_LAK),0) as AA,  isnull(sum(Provision_Amt),0) as BB,  BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and   (loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Call LoadSqlData(AMT5, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim KW1 As String = "Update RPT_F04 set  AMT=" & RSC.Fields("AA").Value & "  , Dep_Amt=" & RSC.Fields("BB").Value & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and (Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') "
                CNN.Execute(KW1)
                RSC.MoveNext()
            End While
        End If
        ' ====AMT KIP  W===
        Dim AMT_W5 As String = "select isnull(sum(Principle_LAK),0) as AA,BUSINESSTYPEDESC from AP_Loan  where 1=1  " & MDWRITEOFF & " and LaonDate BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  and  (loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or loan_grade=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') and GENDER=N'F' group by BUSINESSTYPEDESC order by BUSINESSTYPEDESC "
        Call LoadSqlData(AMT_W5, RSC)
        If RSC.RecordCount > 0 Then
            While Not RSC.EOF
                Dim KW1 As String = "Update RPT_F04 set  AMT_W=" & RSC.Fields("AA").Value & " where left(rpt_ID,3)=N'" & Microsoft.VisualBasic.Left(RSC.Fields("BUSINESSTYPEDESC").Value, 3) & "' and (Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' or Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ (​E)') "
                CNN.Execute(KW1)
                RSC.MoveNext()
            End While
        End If

    End Sub
    Private Sub Arr()
        CNN.Execute("UPDATE RPT_F04 set AccNo=0,Acc=0,Acc_W=0,Amt=0,Amt_w=0,Dep=0,Dep_Amt=0  ")
        Call CALLC()
        'CNN.Execute("UPDATE RPT_F04 set Dep=" & CDbl(txtA.Text) & " where Grp_Nm=N'1. ສິນເຊື່ອປົກກະຕິ (A)' ")
        'CNN.Execute("UPDATE RPT_F04 set Dep=" & CDbl(txtB.Text) & " where Grp_Nm=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' ")
        'CNN.Execute("UPDATE RPT_F04 set Dep=" & CDbl(txtC.Text) & " where Grp_Nm=N'3. ສິນເຊື່ອຕໍ່າກວ່າມາດຕະຖານ (C)' ")
        'CNN.Execute("UPDATE RPT_F04 set Dep=" & CDbl(txtD.Text) & " where Grp_Nm=N'4. ສິນເຊື່ອທີ່ໜ້າສົງໃສ (D)' ")
        'CNN.Execute("UPDATE RPT_F04 set Dep=" & CDbl(txtE.Text) & " where Grp_Nm=N'5. ສິນເຊື່ອທີ່ເປັນໜີ້ສູນ  (E)' ")
        'CNN.Execute("UPDATE RPT_F04 set Dep_Amt=Amt*Dep/100  ")
        'CNN.Execute("UPDATE RPT_F04 set Dep=0 where Dep_Amt=0  ")
        CNN.Execute("UPDATE RPT_F04 set Dep='" & (txtA.Text) & "' where left(Grp_Nm,1)=N'1' ")
        CNN.Execute("UPDATE RPT_F04 set Dep='" & (txtB.Text) & "' where left(Grp_Nm,1)=N'2' ")
        CNN.Execute("UPDATE RPT_F04 set Dep='" & (txtC.Text) & "' where left(Grp_Nm,1)=N'3' ")
        CNN.Execute("UPDATE RPT_F04 set Dep='" & (txtD.Text) & "' where left(Grp_Nm,1)=N'4' ")
        CNN.Execute("UPDATE RPT_F04 set Dep='" & (txtE.Text) & "' where left(Grp_Nm,1)=N'5' ")

    End Sub
    Private Sub BtnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        Call Incom()
        Call Arr()
        Call UPP_INTER()

        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        Call ChangBalance()
        '=======
        ' ''CNN.Execute("DELETE  Ap_balance_6_col ")
        ' ''CNN.Execute("DELETE FROM Ap_balance_6 ")
        ' ''If RM.Checked = True Then
        ' ''    CNN.Execute(" insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  from Ap_balance_TB where  date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  order by Ac_Code asc ")
        ' ''    'ElseIf RP.Checked = True Then
        ' ''Else
        ' ''    CNN.Execute(" insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , 0 , 0 , amt_dr , amt_cr  from Ap_balance_TB where  date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  order by Ac_Code asc ")
        ' ''    'Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
        ' ''    'CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ' ''    '" select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        ' ''    'CNN.Execute(" insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , 0 , 0  from Ap_balance_TB where  date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(S, "yyyy-MM-dd") & "'  order by Ac_Code asc ")
        ' ''    Dim OP As String = " insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , 0 , 0  from Ap_balance_TB where  month(date_work)= '" & Month(MdStartDate) & "' AND   year(date_work)= '" & Year(MdStartDate) & "'   order by Ac_Code asc "
        ' ''    CNN.Execute(OP)
        ' ''    Dim HHH As String = "INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ' '' " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr)as open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code"
        ' ''    CNN.Execute(HHH)
        ' ''End If


        'CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        'CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        'Call Chang_Incom()
 
        ''CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + amt_dr) - (open_amt_cr + amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr + amt_cr) ")
        ''CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr + amt_cr) - (open_amt_dr + amt_dr) where (open_amt_cr + amt_cr) >= (open_amt_dr + amt_dr) ")
        ''CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")

        '==========
        CNN.Execute("update  Ap_Rpt_BLS_Item set Last_amt_dr  =  " & CDbl(0) & " , Last_amt_cr  = " & CDbl(0) & " ,amt_dr  =  " & CDbl(0) & " , amt_cr  = " & CDbl(0) & "  ")
        CNN.Execute("update Ap_Rpt_BLS set Last_Amt = 0 , Amt  =  " & CDbl(0) & "  ")
        CNN.Execute("DELETE FROM Ap_Rpt_BLS_Detail ")
        CNN.Execute("update   TEM_F04 set amt=0 ")
        Call LoadAPLOAN()

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
        '=========
 
        '=========
        CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=" & CDbl(txtAA.Text) & " where rpt_id='2.16.6' ")
        'CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=" & CDbl(txtAA.Text) & " where rpt_id='2.16.9' ")
       

        '========5.3
        CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='1.5.3' ")

        Call LoadSqlData("select sum(amt) as aaa from RPT_F04 where (left(grp_nm,1)=N'3' or left(grp_nm,1)=N'4'or left(grp_nm,1)=N'5' ) ", RSC)
        'dep_amt
        If RSC.RecordCount <> 0 Then
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=" & CDbl(RSC.Fields("aaa").Value) & " where rpt_id='1.5.3' ")
        Else
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='1.5.3' ")
        End If
        '========5.4
        CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='1.5.4' ")

        Call LoadSqlData("select sum(dep_amt) as aaa from RPT_F04 where (left(grp_nm,1)=N'3' or left(grp_nm,1)=N'4'or left(grp_nm,1)=N'5' ) ", RSC)
        'dep_amt
        If RSC.RecordCount <> 0 Then
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=" & CDbl(RSC.Fields("aaa").Value) & " where rpt_id='1.5.4' ")
        Else
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='1.5.4' ")
        End If


        CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='2.16.7' ")

        Call LoadSqlData("select sum(dep_amt) as aaa from RPT_F04 where (grp_nm=N'1. ສິນເຊື່ອປົກກະຕິ (A)' or grp_nm=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' ) ", RSC)
        If RSC.RecordCount <> 0 Then
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=" & CDbl(RSC.Fields("aaa").Value) & " where rpt_id='2.16.7' ")
        Else
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='2.16.7' ")
        End If
        '====
        Call LoadSqlData("select  amt  from Ap_Rpt_Income where (rpt_id=N'V' ) ", RSC)
        If RSC.RecordCount <> 0 Then
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=" & CDbl(RSC.Fields("amt").Value) & " where rpt_id='2.16.10' ")
        Else
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='2.16.10' ")
        End If

        '======== LOAN ====UPDATE =====
        CNN.Execute("update Ap_Rpt_BLS set Ap_Rpt_BLS.amt=TEM_F04.amt from TEM_F04,Ap_Rpt_BLS where TEM_F04.rpt_id=Ap_Rpt_BLS.rpt_id ")
        '========5.2
        CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='1.5.2' ")

        Call LoadSqlData("select sum(amt) as aaa from RPT_F04 where (left(grp_nm,1)=N'2' or left(grp_nm,1)=N'3'or left(grp_nm,1)=N'4' ) ", RSC)
        'dep_amt
        If RSC.RecordCount <> 0 Then
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=" & CDbl(RSC.Fields("aaa").Value) & " where rpt_id='1.5.2' ")
        Else
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='1.5.2' ")
        End If
        '========5.3
        CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='1.5.3' ")

        Call LoadSqlData("select sum(amt) as aaa from RPT_F04 where (left(grp_nm,1)=N'5' ) ", RSC)
        'dep_amt
        If RSC.RecordCount <> 0 Then
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=" & CDbl(RSC.Fields("aaa").Value) & " where rpt_id='1.5.3' ")
        Else
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='1.5.3' ")
        End If

        '========1 =========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='1.1.1' and  rpt_ID<='1.1.3') where  rpt_ID='1.1'")
        '========2 =========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='1.2.1' and  rpt_ID<='1.2.3') where  rpt_ID='1.2'")
        '========3 =========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='1.3.1' and  rpt_ID<='1.3.2') where  rpt_ID='1.3'")
        '========4=========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='1.4.1' and  rpt_ID<='1.4.3') where  rpt_ID='1.4'")

        '========5========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='1.5.1' and  rpt_ID<='1.5.3') where  rpt_ID='1.5'")
        CNN.Execute("update Ap_Rpt_BLS set amt=amt-(select sum(amt) from Ap_Rpt_BLS where rpt_ID='1.5.4') where  rpt_ID='1.5' ")
        '========6========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='1.6.1' and  rpt_ID<='1.6.4') where  rpt_ID='1.6'")
        '========7========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='1.7.1' and  rpt_ID<='1.7.8') where  rpt_ID='1.7'")
        '========8========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='1.8.1' and  rpt_ID<='1.8.3') where  rpt_ID='1.8'")
        '========9=======
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID='1.9.1') where  rpt_ID='1.9'")
        '========10========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='1.10.1' and  rpt_ID<='1.10.3') where  rpt_ID='1.10'")
        '========11=======
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='2.11.1' and  rpt_ID<='2.11.4') where  rpt_ID='2.11'")
        '========12=======
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='2.12.1' and  rpt_ID<='2.12.3') where  rpt_ID='2.12'")
        '========13=======
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID='2.13.1' ) where  rpt_ID='2.13'")
        '========14=======
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='2.14.1' and  rpt_ID<='2.14.2') where  rpt_ID='2.14'")
        '========15=======
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='2.15.1' and  rpt_ID<='2.15.3') where  rpt_ID='2.15'")
        '========16=======
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='2.16.1' and  rpt_ID<='2.16.13') where  rpt_ID='2.16' ")
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where (rpt_ID='2.16.1' or  rpt_ID='2.16.2' or  rpt_ID='2.16.3' or  rpt_ID='2.16.4' or  rpt_ID='2.16.5' or  rpt_ID='2.16.6' or  rpt_ID='2.16.7' or  rpt_ID='2.16.8' or  rpt_ID='2.16.9' or  rpt_ID='2.16.10' or  rpt_ID='2.16.11' or  rpt_ID='2.16.12' or  rpt_ID='2.16.13' )) where  rpt_ID='2.16' ")


        '========I=========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where (rpt_ID='1.1' or  rpt_ID='1.2' or  rpt_ID='1.2' or  rpt_ID='1.4' or  rpt_ID='1.5' or  rpt_ID='1.6' or  rpt_ID='1.7' or  rpt_ID='1.8' or  rpt_ID='1.9' or  rpt_ID='1.10')) where  rpt_ID='I' ")
        '========II=========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where (rpt_ID='2.11' or  rpt_ID='2.12' or  rpt_ID='2.13' or  rpt_ID='2.14' or  rpt_ID='2.15' or  rpt_ID='2.16')) where  rpt_ID='II' ")
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where (rpt_ID='2.11' or  rpt_ID='2.12' or  rpt_ID='2.13' or  rpt_ID='2.14' or  rpt_ID='2.15' or  rpt_ID='2.16')) where  rpt_ID='II' ")


        If CheckBox1.Checked = False Then
            Call LoadReport()
        Else
            Call LoadReportItem()
        End If
        'MdStartDate = d1
        'MdToDate = d2
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

        CNN.Execute("Update Ap_Rpt_BLS_Item set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item , Ap_balance_6_col " & _
                "where Ap_Rpt_BLS_Item.ac_code = Ap_balance_6_col.ac_code And  Rpt_Type = 'In'")

        CNN.Execute("Insert into Ap_Rpt_BLS_Detail (  Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type )" & _
         " select   Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type from Ap_Rpt_BLS_Item where Rpt_Type = 'In' And ( Amt_Dr <>0 or Amt_Cr <>0  or Last_Amt_Dr <>0 or Last_Amt_Cr <>0 )")

    End Sub


    Private Sub SelcectInLast()
        LoadSqlData("select * from Ap_Rpt_BLS_Item where  Rpt_Type = 'In'", RSCIn_M)
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
                CNN.Execute("update  Ap_Rpt_BLS_Item set  Last_amt_dr  =  Last_amt_dr+" & CDbl((.Fields("open_amt_dr").Value)) & " , Last_amt_cr  = Last_amt_cr+" & CDbl((.Fields("open_amt_Cr").Value)) & " , Amt_Dr  =  Amt_Dr+" & CDbl((.Fields("Rem_dr").Value)) & " , Amt_Cr  = Amt_Cr+" & CDbl((.Fields("Rem_cr").Value)) & "   where ac_code = '" & (RSCIn_M.Fields("ac_code").Value) & "' And  Rpt_Type = 'In' ")

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
                CNN.Execute("update  Ap_Rpt_BLS_Item set Amt_Dr  =  Amt_Dr+" & CDbl((.Fields("Rem_dr").Value)) & " , Amt_Cr  = Amt_Cr+" & CDbl((.Fields("Rem_cr").Value)) & "   where ac_code = '" & (RSCIn_M.Fields("ac_code").Value) & "' And  Rpt_Type = 'In' ")
                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub UpdateIIn()
        CNN.Execute("delete Ap_Rpt_BLS_Stock ")
        CNN.Execute(" insert into Ap_Rpt_BLS_Stock ( Rpt_ID , Last_Amt_Dr , Last_Amt_Cr, Amt_Dr , Amt_Cr)" & _
                     "  select Rpt_ID , sum(Last_Amt_Dr) As Last_Amt_Dr , sum(Last_Amt_Cr) As Last_Amt_Cr ,sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr   from Ap_Rpt_BLS_Item  where  Rpt_Type = 'In' group by Rpt_ID")
        CNN.Execute("Update Ap_Rpt_BLS set Amt = Ap_Rpt_BLS_Stock.Amt_Dr-Ap_Rpt_BLS_Stock.Amt_cr ,Last_Amt =Ap_Rpt_BLS_Stock.Last_Amt_dr-Ap_Rpt_BLS_Stock.Last_Amt_Cr  from Ap_Rpt_BLS ,Ap_Rpt_BLS_Stock where  Ap_Rpt_BLS.Rpt_ID=Ap_Rpt_BLS_Stock.Rpt_ID")
    End Sub
    Private Sub UpdateIInLast()
        Dim RSC As New ADODB.Recordset
        LoadSqlData("select Rpt_ID, sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr  from Ap_Rpt_BLS_Item  where  Rpt_Type = 'In' group by Rpt_ID ", RSC)
        With RSC
            Do Until .EOF = True
                CNN.Execute("Update Ap_Rpt_BLS set " & _
                            " Last_Amt ='" & CDbl(CDbl((.Fields("Amt_dr").Value)) - CDbl((.Fields("Amt_cr").Value))) & "' " & _
                               " where Rpt_ID = '" & (.Fields("Rpt_ID").Value) & "' ")
                .MoveNext()
            Loop
        End With
    End Sub


    Private Sub SelectOut()
        CNN.Execute("Update Ap_Rpt_BLS_Item set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item , Ap_balance_6_col " & _
                  "where Ap_Rpt_BLS_Item.ac_code = Ap_balance_6_col.ac_code And  Rpt_Type = 'Out'")

        CNN.Execute("Insert into Ap_Rpt_BLS_Detail (  Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type )" & _
         " select   Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type from Ap_Rpt_BLS_Item where Rpt_Type = 'Out' And ( Amt_Dr <>0 or Amt_Cr <>0  or Last_Amt_Dr <>0 or Last_Amt_Cr <>0 )")

        'LoadSqlData("select * from Ap_Rpt_BLS_Item where  Rpt_Type = 'Out' ", RSCIn_M)
        'With RSCIn_M
        '    Do Until .EOF = True
        '        Call UpdateOut_Item()
        '        .MoveNext()
        '    Loop
        'End With
        'If RSCIn_M.State = ConnectionState.Open Then RSCIn_M.Close()
    End Sub
    Private Sub SelectOutLast()

        LoadSqlData("select * from Ap_Rpt_BLS_Item where  Rpt_Type = 'Out' ", RSCIn_M)
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
                CNN.Execute("update  Ap_Rpt_BLS_Item set Last_Amt_Dr  =  Last_Amt_Dr+" & CDbl((.Fields("Open_Amt_dr").Value)) & " , Last_Amt_Cr  = Last_Amt_Cr+" & CDbl((.Fields("Open_Amt_cr").Value)) & "  , Amt_Dr  =  Amt_Dr+" & CDbl((.Fields("Rem_dr").Value)) & " , Amt_Cr  = Amt_Cr+" & CDbl((.Fields("Rem_cr").Value)) & "   where ac_code = '" & (RSCIn_M.Fields("ac_code").Value) & "' And  Rpt_Type = 'Out' ")

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
                CNN.Execute("update  Ap_Rpt_BLS_Item set Amt_Dr  =  Amt_Dr+" & CDbl((.Fields("Rem_dr").Value)) & " , Amt_Cr  = Amt_Cr+" & CDbl((.Fields("Rem_cr").Value)) & "   where ac_code = '" & (RSCIn_M.Fields("ac_code").Value) & "' And  Rpt_Type = 'Out' ")
                .MoveNext()
            Loop
        End With

    End Sub



    'Private Sub UpdateOutLast()
    '    Dim RSC As New ADODB.Recordset
    '    LoadSqlData("select Rpt_ID, sum(Last_Amt_Dr) As Last_Amt_Dr , sum(Last_Amt_Cr) As Last_Amt_Cr  from Ap_Rpt_BLS_Item  where  Rpt_Type = 'Out' group by Rpt_ID ", RSC)
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
                     "  select Rpt_ID , sum(Last_Amt_Dr) As Last_Amt_Dr , sum(Last_Amt_Cr) As Last_Amt_Cr ,sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr   from Ap_Rpt_BLS_Item  where  Rpt_Type = 'Out' group by Rpt_ID")
        CNN.Execute("Update Ap_Rpt_BLS set Amt = Ap_Rpt_BLS_Stock.Amt_Dr-Ap_Rpt_BLS_Stock.Amt_Dr ,Last_Amt =Ap_Rpt_BLS_Stock.Last_Amt_dr-Ap_Rpt_BLS_Stock.Last_Amt_Cr  from Ap_Rpt_BLS ,Ap_Rpt_BLS_Stock where  Ap_Rpt_BLS.Rpt_ID=Ap_Rpt_BLS_Stock.Rpt_ID")
        CNN.Execute("Update Ap_Rpt_BLS set Amt = Ap_Rpt_BLS_Stock.Amt_Cr-Ap_Rpt_BLS_Stock.Amt_Dr ,Last_Amt =Ap_Rpt_BLS_Stock.Last_Amt_Cr-Ap_Rpt_BLS_Stock.Last_Amt_Dr  from Ap_Rpt_BLS ,Ap_Rpt_BLS_Stock where  Ap_Rpt_BLS.Rpt_ID=Ap_Rpt_BLS_Stock.Rpt_ID")
        Dim RSC As New ADODB.Recordset
        'LoadSqlData("select Rpt_ID, sum(Last_Amt_dr) As Last_Amt_Dr , sum(Last_Amt_cr) As Last_Amt_cr , sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr  from Ap_Rpt_BLS_Item  where  Rpt_Type = 'Out' group by Rpt_ID ", RSC)
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
        LoadSqlData("select Rpt_ID, sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr  from Ap_Rpt_BLS_Item  where  Rpt_Type = 'Out' group by Rpt_ID ", RSC)
        With RSC
            Do Until .EOF = True
                CNN.Execute("Update Ap_Rpt_BLS set " & _
                         " Last_Amt ='" & CDbl(CDbl((.Fields("Amt_cr").Value)) - CDbl((.Fields("Amt_dr").Value))) & "' " & _
                            " where Rpt_ID = '" & (.Fields("Rpt_ID").Value) & "' ")
                .MoveNext()
            Loop
        End With
    End Sub
    Private Sub Update_Sum()
        CNN.Execute("update Ap_Rpt_BLS_Detail set  Rpt_Name=Ap_Rpt_BLS.Description from   Ap_Rpt_BLS_Detail , Ap_Rpt_BLS  where Ap_Rpt_BLS_Detail.Rpt_Id = Ap_Rpt_BLS.Rpt_Id")
        CNN.Execute(" Update Caculate_Rpt set  CLT_Amt  = CLT_Str ,  CLT_Last_Amt  = CLT_Str where CLT_Str = '+' or CLT_Str = '-' or CLT_Str = '*' or CLT_Str = '+' or CLT_Str = '/' or CLT_Str = '(' or CLT_Str=')' Or CLT_Str<>'Cast(('   Or CLT_Str<>')As Float)'")
        CNN.Execute("delete Caculate_Lock")
        CNN.Execute("delete Caculate_Start")
        CNN.Execute(" Insert Into Caculate_Start (Rpt_Id , Rpt_Type , clt_Str , clt_Amt , Clt_Last_Amt ) select Rpt_Id , Rpt_Type , clt_Str , clt_Amt , Clt_Last_Amt from Caculate_Rpt where Rpt_Type = 'BLS'  Order by  Rpt_id ,cnt asc  ")
        CNN.Execute("update Caculate_Start set lck =0")
        CNN.Execute("Insert into Caculate_Lock (cnt_Mt)  SELECT  (SELECT     TOP 1 cnt FROM Caculate_Start AS B WHERE(Rpt_Id = A.Rpt_Id   ) ORDER BY cnt desc) AS cnt FROM Caculate_Start  AS A  GROUP BY Rpt_Id ORDER BY Rpt_Id")
        CNN.Execute("update  Caculate_Start set lck=1 from Caculate_Start ,Caculate_Lock  where Caculate_Start.cnt=Caculate_Lock.cnt_MT")
        CNN.Execute("  Update Caculate_Start set Caculate_Start.Amt = Ap_Rpt_BLS.Amt , Caculate_Start.Last_Amt = Ap_Rpt_BLS.Last_Amt   from Caculate_Start , Ap_Rpt_BLS  where  Caculate_Start.CLT_Str  = Ap_Rpt_BLS.Rpt_Id  ")
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
                            Dim s As String = " Update  Ap_Rpt_BLS set Amt = " & CLT_Str & " , Last_Amt = " & CLT_Last_Str & " where  Rpt_ID =   '" & (RSC1.Fields("Rpt_ID").Value.ToString) & "'"
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
        If DMonth.Text = "ມັງກອນ" Then
            MdStartDate = Format(CDate("01/01/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/01/" & Year(Myy.Value)), "dd-MM-yyyy")
            DMY.Value = Format(CDate("01/01/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ມັງກອນ"
        ElseIf DMonth.Text = "ກຸມພາ" Then
            Dim Day As String
            Dim MM As Date
            Dim Fromm As Date
            MdStartDate = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
            Fromm = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
            MM = Format(CDate("01/3/" & Year(MdStartDate)), "dd-MM-yyyy")
            Day = DateDiff(DateInterval.Day, Fromm, MM)
            MdToDate = Format(CDate(Day & "/02" & "/" & Year(MdStartDate)), "dd-MM-yyyy")
            DMY.Value = Format(CDate("01/02/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ກຸມພາ"
            Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
        ElseIf DMonth.Text = "ມີນາ" Then
            MdStartDate = Format(CDate("01/03/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/03/" & Year(Myy.Value)), "dd-MM-yyyy")
            DMY.Value = Format(CDate("01/03/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ມີນາ"
        ElseIf DMonth.Text = "ເມສາ" Then
            MdStartDate = Format(CDate("01/04/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/04/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ເມສາ"
            DMY.Value = Format(CDate("01/04/" & Year(Myy.Value)), "dd-MM-yyyy")
        ElseIf DMonth.Text = "ພຶດສະພາ" Then
            MdStartDate = Format(CDate("01/05/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/05/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ພຶດສະພາ"
            DMY.Value = Format(CDate("01/05/" & Year(Myy.Value)), "dd-MM-yyyy")
        ElseIf DMonth.Text = "ມີຖຸນາ" Then
            MdStartDate = Format(CDate("01/06/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/06/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ມີຖຸນາ"
            DMY.Value = Format(CDate("01/06/" & Year(Myy.Value)), "dd-MM-yyyy")
        ElseIf DMonth.Text = "ກໍລະກົດ" Then
            MdStartDate = Format(CDate("01/07/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/07/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ກໍລະກົດ"
            DMY.Value = Format(CDate("01/07/" & Year(Myy.Value)), "dd-MM-yyyy")
        ElseIf DMonth.Text = "ສິງຫາ" Then
            MdStartDate = Format(CDate("01/08/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/08/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ສິງຫາ"
            DMY.Value = Format(CDate("01/08/" & Year(Myy.Value)), "dd-MM-yyyy")
        ElseIf DMonth.Text = "ກັນຍາ" Then
            MdStartDate = Format(CDate("01/09/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/09/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ກັນຍາ"
            DMY.Value = Format(CDate("01/09/" & Year(Myy.Value)), "dd-MM-yyyy")
        ElseIf DMonth.Text = "ຕຸລາ" Then
            MdStartDate = Format(CDate("01/10/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/10/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ຕຸລາ"
            DMY.Value = Format(CDate("01/10/" & Year(Myy.Value)), "dd-MM-yyyy")
        ElseIf DMonth.Text = "ພະຈິກ" Then
            MdStartDate = Format(CDate("01/11/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/11/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ພະຈິກ"
            DMY.Value = Format(CDate("01/11/" & Year(Myy.Value)), "dd-MM-yyyy")
        ElseIf DMonth.Text = "ທັນວາ" Then
            MdStartDate = Format(CDate("01/12/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/12/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "ທັນວາ"
            DMY.Value = Format(CDate("01/12/" & Year(Myy.Value)), "dd-MM-yyyy")
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
                    CNN.Execute("Update Ap_Rpt_BLS_Item set amt_dr='" & CDbl(op_dr + amt_dr) - CDbl(op_cr + amt_cr) & "' , amt_cr='" & CDbl(0) & "' where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                End If
                If CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) >= 0 Then
                    CNN.Execute("Update Ap_Rpt_BLS_Item set amt_dr='" & CDbl(0) & "' , amt_cr='" & CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) & "' where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                End If
                'CNN.Execute("update Ap_Rpt_BLS_Item set amt_dr  =  " & CDbl((.Fields("amt_dr").Value)) & " , amt_cr  = " & CDbl((.Fields("amt_cr").Value)) & "   where Ac_code=  '" & (.Fields("Ac_code").Value) & "' ")
                .MoveNext()
            Loop
        End With
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

        'CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=" & CDbl(txtA.Text) & " where rpt_id='2.16.5' ")

        SLF = "SELECT   " & mformat & "  as mformat  , " & MuLngRpt & "  *   FROM Ap_Rpt_BLS  "
        'CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where (rpt_id='1.1' or rpt_id='1.2' or rpt_id='1.3' or rpt_id='1.4'))   where rpt_id='1.5'")
        Call LoadLoGO()

        Dim Rs As New ADODB.Recordset
        With Rs

            If .State = ConnectionState.Open Then .Close()
            'If CheckBox1.Checked = True Then
            .Open(" " & SLF & " where grp<>'' " & RPT_ID & " " & r & " Order by CNT asc  ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            'Else
            '    .Open(" " & SLF & " where grp<>'' " & RPT_ID & " " & r & " and F=0 Order by CNT asc  ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            'End If

            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        Dim FrmPreview As New FmPreview : FrmClosing()

        Dim Rpt As New CryRpt_BLS
        If MdShowLOGO = 1 Then
            Rpt.Subreports(0).SetDataSource(RsLOGO)
        End If
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
        SLF = "SELECT  " & MuLngRpt & "  *   FROM Ap_Rpt_BLS_Detail  "
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

        MMM()
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
        SetControlText(Me)
        Label5.Text = "ມູນຄ່າ"
        Call loadOffice_User()
        Button2.Text = "Export"
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        FmLBS_Item.ShowDialog()
        FmLBS_Item.Focus()
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

    Private Sub DMY_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DMY.ValueChanged
        Call MMM()
    End Sub
    Private Sub MMM()
        CNN.Execute("UPDATE AP_MM set Amt=0 where Amt is null")
        Dim KK As New ADODB.Recordset
        Dim aa As String = " SELECT  * from AP_MM where month(MM)='" & Month(DMY.Value) & "' and  year(MM)='" & Year(DMY.Value) & "'   "
        Call LoadSqlData(aa, KK)
        If KK.RecordCount <> 0 Then
            txtA.Text = (KK.Fields("A").Value.ToString)
            txtB.Text = (KK.Fields("B").Value.ToString)
            txtC.Text = (KK.Fields("C").Value.ToString)
            txtD.Text = (KK.Fields("D").Value.ToString)
            txtE.Text = (KK.Fields("E").Value.ToString)
        Else
            txtA.Text = ""
            txtB.Text = ""
            txtC.Text = ""
            txtD.Text = ""
            txtE.Text = ""
        End If
        'Dim KK As New ADODB.Recordset
        'Dim aa As String = " SELECT  * from AP_MM where month(MM)='" & Month(DMY.Value) & "' and  year(MM)='" & Year(DMY.Value) & "'   "
        'Call LoadSqlData(aa, KK)
        'If KK.RecordCount <> 0 Then
        '    txtA.Text = Format(CDbl(KK.Fields("A").Value), "#,##0.00")
        '    txtB.Text = Format(CDbl(KK.Fields("B").Value), "#,##0.00")
        '    txtC.Text = Format(CDbl(KK.Fields("C").Value), "#,##0.00")
        '    txtD.Text = Format(CDbl(KK.Fields("D").Value), "#,##0.00")
        '    txtE.Text = Format(CDbl(KK.Fields("E").Value), "#,##0.00")
        '    txtAA.Text = Format(CDbl(KK.Fields("Amt").Value), "#,##0.00")
        'Else
        '    txtAA.Text = 0
        '    txtA.Text = 0
        '    txtB.Text = 0
        '    txtC.Text = 0
        '    txtD.Text = 0
        '    txtE.Text = 0
        'End If
    End Sub

    Private Sub txtA_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAA.KeyPress
        If e.KeyChar = Chr(13) Then
            If txtAA.Text = "" Then txtAA.Text = 0
            txtAA.Text = Format(CDbl(txtAA.Text), "#,##0.00")
        End If

    End Sub
    Private Sub ChangBalance11()
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

    Private Sub SelcectInLast11()
        LoadSqlData("select * from Ap_Rpt_Income_Item where  Rpt_Type = 'In'  ", RSCIn_M)
        With RSCIn_M
            Do Until .EOF = True
                UpdateIIn_ItemLast222()
                .MoveNext()
            Loop
        End With
    End Sub
    Private Sub UpdateIIn_ItemLast222()
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
    Private Sub UpdateIIn_ItemLast111()
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
    Private Sub UpdateIInLast11()
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
    Private Sub SelectOutLast11()

        LoadSqlData("select * from Ap_Rpt_Income_Item where  Rpt_Type = 'Out'  ", RSCIn_M)
        With RSCIn_M
            Do Until .EOF = True
                Call UpdateOut_Item22()
                .MoveNext()
            Loop
        End With
    End Sub
    Private Sub UpdateOut_Item22()
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
    Private Sub UpdateOut11()
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
    Private Sub Update_Sum11()
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
                            'Dim s2 As String = " Update  Ap_Rpt_Income set Amt = " & CLT_Str & " + " & CLT_Last_Str & " where  Rpt_ID =N'" & (RSC1.Fields("Rpt_ID").Value.ToString) & "' "
                            'CNN.Execute(s2)
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
    Private Sub txtA_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAA.TextChanged

    End Sub

    Private Sub Button2_Click_3(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call Incom()
        Call Arr()
        Call UPP_INTER()

        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        Call ChangBalance()
        '=======
        ' ''CNN.Execute("DELETE  Ap_balance_6_col ")
        ' ''CNN.Execute("DELETE FROM Ap_balance_6 ")
        ' ''If RM.Checked = True Then
        ' ''    CNN.Execute(" insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  from Ap_balance_TB where  date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  order by Ac_Code asc ")
        ' ''    'ElseIf RP.Checked = True Then
        ' ''Else
        ' ''    CNN.Execute(" insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , 0 , 0 , amt_dr , amt_cr  from Ap_balance_TB where  date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  order by Ac_Code asc ")
        ' ''    'Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
        ' ''    'CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ' ''    '" select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        ' ''    'CNN.Execute(" insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , 0 , 0  from Ap_balance_TB where  date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(S, "yyyy-MM-dd") & "'  order by Ac_Code asc ")
        ' ''    Dim OP As String = " insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , 0 , 0  from Ap_balance_TB where  month(date_work)= '" & Month(MdStartDate) & "' AND   year(date_work)= '" & Year(MdStartDate) & "'   order by Ac_Code asc "
        ' ''    CNN.Execute(OP)
        ' ''    Dim HHH As String = "INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ' '' " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr)as open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code"
        ' ''    CNN.Execute(HHH)
        ' ''End If


        'CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        'CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        'Call Chang_Incom()

        ''CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + amt_dr) - (open_amt_cr + amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr + amt_cr) ")
        ''CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr + amt_cr) - (open_amt_dr + amt_dr) where (open_amt_cr + amt_cr) >= (open_amt_dr + amt_dr) ")
        ''CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")

        '==========
        CNN.Execute("update  Ap_Rpt_BLS_Item set Last_amt_dr  =  " & CDbl(0) & " , Last_amt_cr  = " & CDbl(0) & " ,amt_dr  =  " & CDbl(0) & " , amt_cr  = " & CDbl(0) & "  ")
        CNN.Execute("update Ap_Rpt_BLS set Last_Amt = 0 , Amt  =  " & CDbl(0) & "  ")
        CNN.Execute("DELETE FROM Ap_Rpt_BLS_Detail ")
        CNN.Execute("update   TEM_F04 set amt=0 ")
        Call LoadAPLOAN()

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
        '=========

        '=========
        CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=" & CDbl(txtAA.Text) & " where rpt_id='2.16.6' ")
        'CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=" & CDbl(txtAA.Text) & " where rpt_id='2.16.9' ")


        '========5.3
        CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='1.5.3' ")

        Call LoadSqlData("select sum(amt) as aaa from RPT_F04 where (left(grp_nm,1)=N'3' or left(grp_nm,1)=N'4'or left(grp_nm,1)=N'5' ) ", RSC)
        'dep_amt
        If RSC.RecordCount <> 0 Then
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=" & CDbl(RSC.Fields("aaa").Value) & " where rpt_id='1.5.3' ")
        Else
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='1.5.3' ")
        End If
        '========5.4
        CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='1.5.4' ")

        Call LoadSqlData("select sum(dep_amt) as aaa from RPT_F04 where (left(grp_nm,1)=N'3' or left(grp_nm,1)=N'4'or left(grp_nm,1)=N'5' ) ", RSC)
        'dep_amt
        If RSC.RecordCount <> 0 Then
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=" & CDbl(RSC.Fields("aaa").Value) & " where rpt_id='1.5.4' ")
        Else
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='1.5.4' ")
        End If


        CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='2.16.7' ")

        Call LoadSqlData("select sum(dep_amt) as aaa from RPT_F04 where (grp_nm=N'1. ສິນເຊື່ອປົກກະຕິ (A)' or grp_nm=N'2. ສິນເຊື່ອຄວນເອົາໃຈໃສ່ (B)' ) ", RSC)
        If RSC.RecordCount <> 0 Then
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=" & CDbl(RSC.Fields("aaa").Value) & " where rpt_id='2.16.7' ")
        Else
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='2.16.7' ")
        End If
        '====
        Call LoadSqlData("select  amt  from Ap_Rpt_Income where (rpt_id=N'V' ) ", RSC)
        If RSC.RecordCount <> 0 Then
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=" & CDbl(RSC.Fields("amt").Value) & " where rpt_id='2.16.10' ")
        Else
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='2.16.10' ")
        End If

        '======== LOAN ====UPDATE =====
        CNN.Execute("update Ap_Rpt_BLS set Ap_Rpt_BLS.amt=TEM_F04.amt from TEM_F04,Ap_Rpt_BLS where TEM_F04.rpt_id=Ap_Rpt_BLS.rpt_id ")
        '========5.2
        CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='1.5.2' ")

        Call LoadSqlData("select sum(amt) as aaa from RPT_F04 where (left(grp_nm,1)=N'2' or left(grp_nm,1)=N'3'or left(grp_nm,1)=N'4' ) ", RSC)
        'dep_amt
        If RSC.RecordCount <> 0 Then
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=" & CDbl(RSC.Fields("aaa").Value) & " where rpt_id='1.5.2' ")
        Else
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='1.5.2' ")
        End If
        '========5.3
        CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='1.5.3' ")

        Call LoadSqlData("select sum(amt) as aaa from RPT_F04 where (left(grp_nm,1)=N'5' ) ", RSC)
        'dep_amt
        If RSC.RecordCount <> 0 Then
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=" & CDbl(RSC.Fields("aaa").Value) & " where rpt_id='1.5.3' ")
        Else
            CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=0 where rpt_id='1.5.3' ")
        End If

        '========1 =========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='1.1.1' and  rpt_ID<='1.1.3') where  rpt_ID='1.1'")
        '========2 =========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='1.2.1' and  rpt_ID<='1.2.3') where  rpt_ID='1.2'")
        '========3 =========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='1.3.1' and  rpt_ID<='1.3.2') where  rpt_ID='1.3'")
        '========4=========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='1.4.1' and  rpt_ID<='1.4.3') where  rpt_ID='1.4'")

        '========5========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='1.5.1' and  rpt_ID<='1.5.3') where  rpt_ID='1.5'")
        CNN.Execute("update Ap_Rpt_BLS set amt=amt-(select sum(amt) from Ap_Rpt_BLS where rpt_ID='1.5.4') where  rpt_ID='1.5' ")
        '========6========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='1.6.1' and  rpt_ID<='1.6.4') where  rpt_ID='1.6'")
        '========7========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='1.7.1' and  rpt_ID<='1.7.8') where  rpt_ID='1.7'")
        '========8========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='1.8.1' and  rpt_ID<='1.8.3') where  rpt_ID='1.8'")
        '========9=======
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID='1.9.1') where  rpt_ID='1.9'")
        '========10========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='1.10.1' and  rpt_ID<='1.10.3') where  rpt_ID='1.10'")
        '========11=======
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='2.11.1' and  rpt_ID<='2.11.4') where  rpt_ID='2.11'")
        '========12=======
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='2.12.1' and  rpt_ID<='2.12.3') where  rpt_ID='2.12'")
        '========13=======
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID='2.13.1' ) where  rpt_ID='2.13'")
        '========14=======
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='2.14.1' and  rpt_ID<='2.14.2') where  rpt_ID='2.14'")
        '========15=======
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='2.15.1' and  rpt_ID<='2.15.3') where  rpt_ID='2.15'")
        '========16=======
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where rpt_ID>='2.16.1' and  rpt_ID<='2.16.13') where  rpt_ID='2.16' ")
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where (rpt_ID='2.16.1' or  rpt_ID='2.16.2' or  rpt_ID='2.16.3' or  rpt_ID='2.16.4' or  rpt_ID='2.16.5' or  rpt_ID='2.16.6' or  rpt_ID='2.16.7' or  rpt_ID='2.16.8' or  rpt_ID='2.16.9' or  rpt_ID='2.16.10' or  rpt_ID='2.16.11' or  rpt_ID='2.16.12' or  rpt_ID='2.16.13' )) where  rpt_ID='2.16' ")


        '========I=========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where (rpt_ID='1.1' or  rpt_ID='1.2' or  rpt_ID='1.2' or  rpt_ID='1.4' or  rpt_ID='1.5' or  rpt_ID='1.6' or  rpt_ID='1.7' or  rpt_ID='1.8' or  rpt_ID='1.9' or  rpt_ID='1.10')) where  rpt_ID='I' ")
        '========II=========
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where (rpt_ID='2.11' or  rpt_ID='2.12' or  rpt_ID='2.13' or  rpt_ID='2.14' or  rpt_ID='2.15' or  rpt_ID='2.16')) where  rpt_ID='II' ")
        CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where (rpt_ID='2.11' or  rpt_ID='2.12' or  rpt_ID='2.13' or  rpt_ID='2.14' or  rpt_ID='2.15' or  rpt_ID='2.16')) where  rpt_ID='II' ")


        If CheckBox1.Checked = False Then
            Call LoadReport_Export()
        Else
            Call LoadReportItem_Export()
        End If
    End Sub

    Private Sub LoadReportItem_Export()


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
        SLF = "SELECT  " & MuLngRpt & "  *   FROM Ap_Rpt_BLS_Detail  "
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

 
        Rpt.SetDataSource(Rs)
        Rpt.Refresh()
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.ExportReport()
        FrmPreview = Nothing
        'FrmPreview.ReportViewer.ReportSource = Rpt
        'FrmPreview.ReportViewer.DisplayGroupTree = False

        'FrmPreview.WindowState = FormWindowState.Maximized
        'FrmPreview.ShowDialog()
        'FrmPreview.Focus()
    End Sub
    Private Sub LoadReport_Export()
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

        'CNN.Execute("UPDATE   Ap_Rpt_BLS set amt=" & CDbl(txtA.Text) & " where rpt_id='2.16.5' ")

        SLF = "SELECT  " & MuLngRpt & "  *   FROM Ap_Rpt_BLS  "
        'CNN.Execute("update Ap_Rpt_BLS set amt=(select sum(amt) from Ap_Rpt_BLS where (rpt_id='1.1' or rpt_id='1.2' or rpt_id='1.3' or rpt_id='1.4'))   where rpt_id='1.5'")
        Call LoadLoGO()

        Dim Rs As New ADODB.Recordset
        With Rs

            If .State = ConnectionState.Open Then .Close()
            'If CheckBox1.Checked = True Then
            .Open(" " & SLF & " where grp<>'' " & RPT_ID & " " & r & " Order by CNT asc  ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            'Else
            '    .Open(" " & SLF & " where grp<>'' " & RPT_ID & " " & r & " and F=0 Order by CNT asc  ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            'End If

            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        Dim FrmPreview As New FmPreview : FrmClosing()

        Dim Rpt As New CryRpt_BLS
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
        'FrmPreview.WindowState = FormWindowState.Maximized
        'FrmPreview.Show()
        'FrmPreview.Focus()
    End Sub
End Class