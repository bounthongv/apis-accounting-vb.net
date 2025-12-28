Public Class FmAmtStatus1234

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
    Private Sub FmAmtStatus_statement_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        RM.Checked = True
        Ds.Text = MWorkSetting
        Myy.Text = MWorkSetting
        yy.Text = MWorkSetting
        Toyy.Text = MWorkSetting
        Pyy.Text = MWorkSetting
        If Month(MWorkSetting) = 1 Then
            DMonth.SelectedIndex = 0
            Period.SelectedIndex = 0
        ElseIf Month(MWorkSetting) = 2 Then
            DMonth.SelectedIndex = 1
            Period.SelectedIndex = 0
        ElseIf Month(MWorkSetting) = 3 Then
            DMonth.SelectedIndex = 2
            Period.SelectedIndex = 0
        ElseIf Month(MWorkSetting) = 4 Then
            DMonth.SelectedIndex = 3
            Period.SelectedIndex = 1
        ElseIf Month(MWorkSetting) = 5 Then
            DMonth.SelectedIndex = 4
            Period.SelectedIndex = 1
        ElseIf Month(MWorkSetting) = 6 Then
            DMonth.SelectedIndex = 5
            Period.SelectedIndex = 1
        ElseIf Month(MWorkSetting) = 7 Then
            DMonth.SelectedIndex = 6
            Period.SelectedIndex = 2
        ElseIf Month(MWorkSetting) = 8 Then
            DMonth.SelectedIndex = 7
            Period.SelectedIndex = 2
        ElseIf Month(MWorkSetting) = 9 Then
            DMonth.SelectedIndex = 8
            Period.SelectedIndex = 2
        ElseIf Month(MWorkSetting) = 10 Then
            DMonth.SelectedIndex = 9
            Period.SelectedIndex = 3
        ElseIf Month(MWorkSetting) = 11 Then
            DMonth.SelectedIndex = 10
            Period.SelectedIndex = 3
        ElseIf Month(MWorkSetting) = 12 Then
            DMonth.SelectedIndex = 11
            Period.SelectedIndex = 3
        End If
        Call selectLoad()
        Call Click_Last()
        'SetControlText(Me)
        Call loadOffice_User()
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

            CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")
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




    End Sub

    Private Sub Call_ALL()
        CNN.Execute("update Ap_Rpt_Income set  Last_Amt  = 0 , Amt  = 0    ")
        CNN.Execute("DELETE FROM Ap_Rpt_Incon_Detail ")
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        ChangBalance()
        SelcectInLast()
        UpdateIInLast()
        SelectOutLast()
        UpdateOut()
        Update_Sum()
    End Sub
    Private Sub cal_neung()
        Dim aa As String

        aa = "  update Ap_Rpt_Income_Item set Last_Amt_Dr =0,Last_Amt_Cr =0,Amt_Dr =0,Amt_Cr =0 "
        CNN.Execute(aa)
        aa = "  update Ap_Rpt_Income_Item set Last_Amt_Dr =Ap_balance_6_col.open_amt_dr , " & _
   "  Last_Amt_Cr =Ap_balance_6_col.open_amt_cr,Amt_Dr =Ap_balance_6_col.amt_dr ,Amt_Cr =Ap_balance_6_col.amt_cr  " & _
      "  from Ap_balance_6_col  where Ap_Rpt_Income_Item.Ac_Code =Ap_balance_6_col.ac_code    "
        CNN.Execute(aa)
        '====================
        aa = "update AP_Rpt_Amt_Status set Amt4=   (select     sum(Last_Amt_cr+ Amt_cr) -sum(Last_Amt_Dr+ Amt_Dr) from Ap_Rpt_Income_Item where  RPT_ID='1.1.1.1')  where rpt_id='18'"
        CNN.Execute(aa)
        aa = "update AP_Rpt_Amt_Status set Amt4=  Amt4+(select     sum(Last_Amt_cr+ Amt_cr) -sum(Last_Amt_Dr+ Amt_Dr) from Ap_Rpt_Income_Item where  RPT_ID='1.1.1.2.1')  where rpt_id='18'"
        CNN.Execute(aa)
        aa = "update AP_Rpt_Amt_Status set Amt4=  Amt4+(select     sum(Last_Amt_cr+ Amt_cr) -sum(Last_Amt_Dr+ Amt_Dr) from Ap_Rpt_Income_Item where  RPT_ID='1.1.1.2.3')  where rpt_id='18'"
        CNN.Execute(aa)
        aa = "update AP_Rpt_Amt_Status set Amt4=  Amt4+(select     sum(Last_Amt_cr+ Amt_cr) -sum(Last_Amt_Dr+ Amt_Dr) from Ap_Rpt_Income_Item where  RPT_ID='1.1.1.3')  where rpt_id='18'"
        CNN.Execute(aa)
        aa = "update AP_Rpt_Amt_Status set Amt4=  Amt4+(select     sum(Last_Amt_cr+ Amt_cr) -sum(Last_Amt_Dr+ Amt_Dr) from Ap_Rpt_Income_Item where  RPT_ID='1.1.1.4')  where rpt_id='18'"
        CNN.Execute(aa)
        aa = "update AP_Rpt_Amt_Status set Amt4=  Amt4+(select     sum(Last_Amt_cr+ Amt_cr) -sum(Last_Amt_Dr+ Amt_Dr) from Ap_Rpt_Income_Item where  RPT_ID='1.1.2.5')  where rpt_id='18'"
        CNN.Execute(aa)
        '====================
        aa = "update AP_Rpt_Amt_Status set Amt4=Amt4-  (select sum(Last_Amt_Dr+ Amt_Dr)- sum(Last_Amt_cr+ Amt_cr)  from Ap_Rpt_Income_Item where  RPT_ID='1.2.1.1.2')  where rpt_id='18'"
        CNN.Execute(aa)
        aa = "update AP_Rpt_Amt_Status set Amt4=Amt4-  (select sum(Last_Amt_Dr+ Amt_Dr)- sum(Last_Amt_cr+ Amt_cr)  from Ap_Rpt_Income_Item where  RPT_ID='1.2.1.1.5')  where rpt_id='18'"
        CNN.Execute(aa)
        aa = "update AP_Rpt_Amt_Status set Amt4=Amt4-  (select sum(Last_Amt_Dr+ Amt_Dr)- sum(Last_Amt_cr+ Amt_cr)  from Ap_Rpt_Income_Item where  RPT_ID='1.2.2.2.1')  where rpt_id='18'"
        CNN.Execute(aa)
        aa = "update AP_Rpt_Amt_Status set Amt4=Amt4-  (select sum(Last_Amt_Dr+ Amt_Dr)- sum(Last_Amt_cr+ Amt_cr)  from Ap_Rpt_Income_Item where  RPT_ID='1.2.2.2.2')  where rpt_id='18'"
        CNN.Execute(aa)
        aa = "update AP_Rpt_Amt_Status set Amt4=Amt4-  (select sum(Last_Amt_Dr+ Amt_Dr)- sum(Last_Amt_cr+ Amt_cr)  from Ap_Rpt_Income_Item where  RPT_ID='1.2.2.2.3')  where rpt_id='18'"
        CNN.Execute(aa)
        aa = "update AP_Rpt_Amt_Status set Amt4=Amt4-  (select sum(Last_Amt_Dr+ Amt_Dr)- sum(Last_Amt_cr+ Amt_cr)  from Ap_Rpt_Income_Item where  RPT_ID='1.2.2.2.4')  where rpt_id='18'"
        CNN.Execute(aa)

        aa = "update AP_Rpt_Amt_Status set Amt4=Amt4-  (select sum(Last_Amt_Dr+ Amt_Dr)- sum(Last_Amt_cr+ Amt_cr)  from Ap_Rpt_Income_Item where  RPT_ID='1.2.4.2')  where rpt_id='18'"
        CNN.Execute(aa)
        'aa = "update AP_Rpt_Amt_Status set Amt4=Amt4- (select   isnull(sum(Last_Amt_Dr+ Amt_Dr)- sum(Last_Amt_cr+ Amt_cr),0)  from Ap_Rpt_Income_Item where  RPT_ID='1.2.4.1')  where rpt_id='18'"
        'CNN.Execute(aa)


    End Sub
    Private Sub BtnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        'Click_Last()
        Dim aa As String = ""
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()

        Off_Find = Off_Usr.Text : MuTable = ""
        Off_Find2 = Off_Usr.Text : MuTable = ""
        'MsgBox(Off_Find)
        Call Find_Company()

        CNN.Execute("DELETE  Ap_balance_6_col ")
        CNN.Execute("DELETE FROM Ap_balance_6 ")
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
        " select ac_code  ,  0 , 0  , 0  , 0  , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")
        Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr      ) " & _
        " select  ac_code  ,  0 , 0   , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  , 0 , 0   from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        CNN.Execute("INSERT INTO Ap_balance_6 (  ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
        " select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr   ,  0 , 0  , 0 , 0  from Open_jn WHERE    date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

        CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr  ) " & _
        " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr) as open_amt_cr , sum(Amt_Last_M_Dr) as Amt_Last_M_Dr , sum(Amt_Last_M_Cr) as Amt_Last_M_Cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")

        CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        New_Code = "3901000"
        Code_Dr = "4"
        Code_Cr = "5"
        Call Chang_Incom()

        CNN.Execute("update  Ap_balance_6_col set Amt_Last_M_dr = 0 where Amt_Last_M_dr  is null")
        CNN.Execute("update  Ap_balance_6_col set Amt_Last_M_cr = 0 where Amt_Last_M_cr  is null")

        CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + Amt_Last_M_Dr + amt_dr) - (open_amt_cr + Amt_Last_M_Cr + amt_cr) where (open_amt_dr + Amt_Last_M_Dr + amt_dr) >= (open_amt_cr + Amt_Last_M_Cr + amt_cr) ")
        CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr +  Amt_Last_M_Cr + amt_cr) - (open_amt_dr + Amt_Last_M_Dr + amt_dr) where (open_amt_cr + Amt_Last_M_Cr + amt_cr) >= (open_amt_dr + Amt_Last_M_Dr + amt_dr) ")
        CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        'Call ChangBalance()

        Call Call_ALL()


        CNN.Execute("update AP_Rpt_Amt_Status set Amt1=0 ,amt2=0 , Amt3=0 , Amt4=0  ,Amt5=0")
        Dim ds As Date
        ds = DateAdd(DateInterval.Year, -1, MdStartDate)

        'MsgBox(DisT)
        CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '3108210%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")

        CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '3202%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")
        'CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '320%'  And Ac_Code<> '3202' And Year(Date_Work)=" & Year(ds) & ")  where rpt_id='01'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where  Ac_Code Like '3810000%'    And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  Amt4+(select sum(Amt_Cr- Amt_Dr) from Open_jn where    Ac_Code Like '3908000%'    And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")

        CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_dr) from Gen_jn where Ac_Code Like '3810000%'  And  Year(Date_Work)='" & Year(ds) & "') * -1 where rpt_id='03'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt_cr-Amt_Dr) from Gen_jn where Ac_Code Like '3202000%'  And  Year(Date_Work)='" & Year(ds) & "')  where rpt_id='04'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_cr-Amt_Dr) from Gen_jn where Ac_Code Like '3202000%' And  Year(Date_Work)='" & Year(ds) & "') * -1 where rpt_id='04'")

        CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '3203%'  And Ac_Code <> '3202%' And  Year(Date_Work)='" & Year(ds) & "')  where rpt_id='05'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum( Amt_cr) from Gen_jn where Ac_Code Like '3203%'  And Ac_Code <> '3202%' And Year(Date_Work)='" & Year(ds) & "') * -1  where rpt_id='05'")

        CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum( Amt_cr) from Gen_jn where Ac_Code Like '3108210%'  And Ac_Code <> '3202%' And Year(Date_Work)='" & Year(ds) & "') * -1  where rpt_id='06'")
        'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Cr) from Open_jn where Ac_Code Like '3108210%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='07'")

        CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Dr) from Gen_jn where Ac_Code Like '3901000%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='07'")

        CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status  where rpt_id='01' Or rpt_id='06' )  where rpt_id='08'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt2) from AP_Rpt_Amt_Status  where rpt_id='01' Or rpt_id='02'  Or rpt_id='03' Or rpt_id='04' Or rpt_id='05' Or rpt_id='06' Or rpt_id='07' )  where rpt_id='08'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status  where rpt_id='01' Or rpt_id='03' Or rpt_id='04' Or rpt_id='05' Or rpt_id='07' )  where rpt_id='08'")

        CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status  where rpt_id='08' )  where rpt_id='10'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt2) from AP_Rpt_Amt_Status  where rpt_id='08' )  where rpt_id='10'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status  where rpt_id='08' )  where rpt_id='10'")

        CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status  where rpt_id='10' Or  rpt_id='11' )  where rpt_id='12'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status  where rpt_id='10'  Or  rpt_id='11' )  where rpt_id='12'")
        '============
        ds = DateAdd(DateInterval.Year, 0, MdStartDate)
        CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '3108210%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='12'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '3202%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='12'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt_Dr) from Open_jn where Ac_Code Like '320%'  And Ac_Code<> '3202' And Year(Date_Work)=" & Year(ds) & ")  where rpt_id='12'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '3810000%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='12'")

        CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '23626%'   And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='14'")

        CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt_cr-Amt_Dr) from Gen_jn where Ac_Code Like '3202000%'   And date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "')   where rpt_id='15'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_cr-Amt_Dr) from Gen_jn where Ac_Code Like '3202000%'  And date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "')  * -1 where rpt_id='15'")

        CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '3203%'  And Ac_Code <> '3202%' And  Year(Date_Work)='" & Year(ds) & "') where rpt_id='15'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum( Amt_cr) from Gen_jn where Ac_Code Like '3203%'  And Ac_Code <> '3202%' And  Year(Date_Work)='" & Year(ds) & "') * -1  where rpt_id='16'")

        CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum( Amt_cr) from Gen_jn where Ac_Code Like '3108210%'  And Ac_Code <> '3202%' And  Year(Date_Work)='" & Year(ds) & "') where rpt_id='17'")


        Call bar_neung()
        'aa = "update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Dr) from Gen_jn where Ac_Code Like '3901000%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='18'"
        'aa = "update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Dr) from Gen_jn where Ac_Code Like '3901000%' And date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "')  where rpt_id='18'"
        'CNN.Execute(aa)
        'date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'
        CNN.Execute("update AP_Rpt_Amt_Status set  Amt4=0  where Amt4 is null ")

        aa = "update AP_Rpt_Amt_Status set Amt4=  (select sum(Rem_dr-Rem_cr) from Ap_balance_6_col where Ac_Code Like '3901000%'  )  where rpt_id='18'   "
        CNN.Execute(aa)
        'aa = "update AP_Rpt_Amt_Status set Amt4=  (select sum(Rem_dr+Rem_cr) from Ap_balance_6_col where Ac_Code Like '3901000%'  )  where rpt_id='18'   "
        'CNN.Execute(aa)

        'Call cal_neung()


        CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status  where rpt_id='12' Or rpt_id='17' )  where rpt_id='19'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt2) from AP_Rpt_Amt_Status  where rpt_id='12' Or rpt_id='15' )  where rpt_id='19'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt3) from AP_Rpt_Amt_Status  where rpt_id='12' Or rpt_id='16'  )  where rpt_id='19'")

        CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status  where   rpt_id='12' Or rpt_id='14'  Or rpt_id='15' Or rpt_id='15' Or rpt_id='16' Or rpt_id='18')  where rpt_id='19'")

        CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status  where  rpt_id='10'   Or rpt_id='11')  where rpt_id='12'")

        CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status  where  rpt_id='12'   Or rpt_id='18')  where rpt_id='19'")

        CNN.Execute("update AP_Rpt_Amt_Status set Amt1=0 where amt1 Is null")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt2=0 where amt2 Is null")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt3=0 where amt3 Is null")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt4=0 where amt4 Is null")

        If CheckBox1.Checked = False Then
            Call LoadReport()
        Else
            'Call LoadReportItem()
        End If
    End Sub
    Private Sub bar_neung()
        Dim aa As String
        CNN.Execute("DELETE  Ap_balance_6_col ")
        CNN.Execute("DELETE FROM Ap_balance_6 ")
        aa = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
        " select ac_code  ,  0 , 0  , 0  , 0  , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code "
        CNN.Execute(aa)

        Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)

        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr      ) " & _
        " select  ac_code  ,  0 , 0   , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  , 0 , 0   from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

        aa = "INSERT INTO Ap_balance_6 (  ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
        " select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr   ,  0 , 0  , 0 , 0  from Open_jn WHERE    date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code"
        CNN.Execute(aa)

        CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr  ) " & _
        " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr) as open_amt_cr , sum(Amt_Last_M_Dr) as Amt_Last_M_Dr , sum(Amt_Last_M_Cr) as Amt_Last_M_Cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")

        CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")

        CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr +  amt_dr) - (open_amt_cr +  amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr +  amt_cr) ")
        CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr +  amt_cr) - (open_amt_dr +  amt_dr) where (open_amt_cr +  amt_cr) >= (open_amt_dr +  amt_dr) ")

        LoadSqlData("select *  from  Ap_balance_6_col   where Ac_Code Like '3901000%'  ", RSC)
        If RSC.RecordCount > 0 Then
            Dim ram_dr As Double = 0
            Dim ram_cr As Double = 0

            ram_dr = RSC.Fields("Rem_dr").Value
            ram_cr = RSC.Fields("Rem_cr").Value
            If CDbl(ram_dr) + CDbl(ram_cr) = 0 Then
                'New_Code = "3901000"
                Code_Dr = "4"
                Code_Cr = "5"
                Dim amt As Double = 0
                Dim amt_4 As Double = 0
                Dim amt_5 As Double = 0

                Dim RSC4 As New ADODB.Recordset
                aa = "select   sum(amt_dr-amt_cr) as amt4  from gen_jn  " & _
                " WHERE LEFT(Ac_Code,1)= '4' and  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'   group BY   LEFT(Ac_Code,1)"
                LoadSqlData(aa, RSC4)
                If RSC.RecordCount > 0 Then
                    amt_4 = RSC4.Fields("amt4").Value
                End If

                Dim RSC5 As New ADODB.Recordset
                aa = "select   sum(amt_cr-amt_dr) as amt5  from gen_jn  " & _
                " WHERE LEFT(Ac_Code,1)= '5' and  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'   group BY   LEFT(Ac_Code,1)"
                LoadSqlData(aa, RSC5)
                If RSC.RecordCount > 0 Then
                    amt_5 = RSC5.Fields("amt5").Value
                End If

                amt = amt_5 - amt_4
                CNN.Execute("Update  Ap_balance_6_col set    amt_dr = " & CDbl(amt) & "   where ac_code='3901000' ")
                'CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= ( amt_dr) - ( amt_cr) where ( amt_dr) >= (  amt_cr) and     ac_code='3901000' ")
                'CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (  amt_cr) - (  amt_dr) where ( amt_cr) >= (  amt_dr) and   ac_code='3901000' ")
                CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr +  amt_dr) - (open_amt_cr +  amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr +  amt_cr) ")
                CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr +  amt_cr) - (open_amt_dr +  amt_dr) where (open_amt_cr +  amt_cr) >= (open_amt_dr +  amt_dr) ")

                'If amt > 0 Then
                '    CNN.Execute("Update  Ap_balance_6_col set    amt_dr = " & CDbl(amt) & "   where'3901000' ")
                'Else
                '    CNN.Execute("Update  Ap_balance_6_col set    amt_cr = " & CDbl(amt) & "   where'3901000' ")
                'End If

            End If

        End If

        CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")


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
        LoadSqlData("select * from Ap_Rpt_Cashflow_Item where  Rpt_Type = 'In'", RSCIn_M)
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
        LoadSqlData(" select * from Ap_balance_6_col   where ac_code =  '" & (RSCIn_M.Fields("Ac_Code").Value) & "' ", RSCkk)
        With RSCkk
            Do Until .EOF = True
                CNN.Execute("Insert into Ap_Rpt_Cashflow_Detail (  Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type ) values ( '" & CStr((RSCIn_M.Fields("Rpt_Id").Value)) & "' , '" & CStr((.Fields("Ac_Code").Value)) & "' , N'" & CStr((.Fields("Ac_Name").Value)) & "'   , " & CDbl((.Fields("open_amt_dr").Value)) & " , " & CDbl((.Fields("open_amt_Cr").Value)) & "   , " & CDbl((.Fields("Rem_dr").Value)) & " , " & CDbl((.Fields("Rem_cr").Value)) & " , 'In')")
                CNN.Execute("update  Ap_Rpt_Cashflow_Item set  Last_amt_dr  =  Last_amt_dr+" & CDbl((.Fields("open_amt_dr").Value)) & " , Last_amt_cr  = Last_amt_cr+" & CDbl((.Fields("open_amt_Cr").Value)) & " , Amt_Dr  =  Amt_Dr+" & CDbl((.Fields("Rem_dr").Value)) & " , Amt_Cr  = Amt_Cr+" & CDbl((.Fields("Rem_cr").Value)) & "   where ac_code = '" & (RSCIn_M.Fields("ac_code").Value) & "' And  Rpt_Type = 'In' ")

                .MoveNext()
            Loop
        End With
    End Sub


    Private Sub UpdateIIn_ItemLast()
        Dim RSCkk As New ADODB.Recordset
        LoadSqlData(" select * from Ap_balance_6_col   where ac_code =  '" & (RSCIn_M.Fields("Ac_Code").Value) & "' ", RSCkk)
        With RSCkk
            Do Until .EOF = True
                CNN.Execute("Insert into Ap_Rpt_Incon_Detail (  Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type ) values ( '" & CStr((RSCIn_M.Fields("Rpt_Id").Value)) & "' , '" & CStr((.Fields("Ac_Code").Value)) & "' , N'" & CStr((.Fields("Ac_Name").Value)) & "'   , " & CDbl((.Fields("open_amt_dr").Value)) & " , " & CDbl((.Fields("open_amt_Cr").Value)) & "   , " & CDbl((.Fields("Amt_dr").Value)) & " , " & CDbl((.Fields("Amt_cr").Value)) & " , 'In')")
                CNN.Execute("update  Ap_Rpt_Income_Item set  Last_amt_dr  =  Last_amt_dr+" & CDbl((.Fields("open_amt_dr").Value)) & " , Last_amt_cr  = Last_amt_cr+" & CDbl((.Fields("open_amt_Cr").Value)) & " , Amt_Dr  =  Amt_Dr+" & CDbl((.Fields("Amt_dr").Value)) & " , Amt_Cr  = Amt_Cr+" & CDbl((.Fields("Amt_cr").Value)) & "   where ac_code = '" & (RSCIn_M.Fields("ac_code").Value) & "' And  Rpt_Type = 'In' ")
                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub UpdateIIn()
        Dim RSC As New ADODB.Recordset
        LoadSqlData("select Rpt_ID , sum(Last_Amt_Dr) As Last_Amt_Dr , sum(Last_Amt_Cr) As Last_Amt_Cr ,sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr   from Ap_Rpt_Cashflow_Item  where  Rpt_Type = 'In' group by Rpt_ID ", RSC)
        With RSC
            Do Until .EOF = True
                CNN.Execute("Update Ap_Rpt_Cashflow set " & _
                            " Amt = '" & CDbl(CDbl((.Fields("Amt_dr").Value)) - CDbl((.Fields("Amt_cr").Value))) & "' " & _
                              " , Last_Amt ='" & CDbl(CDbl((.Fields("Last_Amt_dr").Value)) - CDbl((.Fields("Last_Amt_cr").Value))) & "' " & _
                               " where Rpt_ID = '" & (.Fields("Rpt_ID").Value) & "' ")
                'MsgBox(.Fields("Rpt_ID").Value)
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

        LoadSqlData("select * from Ap_Rpt_Cashflow_Item where  Rpt_Type = 'Out' ", RSCIn_M)
        With RSCIn_M
            Do Until .EOF = True
                Call UpdateOut_Item()
                .MoveNext()
            Loop
        End With
        'If RSCIn_M.State = CNNectionState.Open Then RSCIn_M.Close()
    End Sub
    Private Sub SelectOutLast()

        LoadSqlData("select * from Ap_Rpt_Income_Item where  Rpt_Type = 'Out'  ", RSCIn_M)
        With RSCIn_M
            Do Until .EOF = True
                Call UpdateOut_Item()
                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub UpdateOut_Item()
        Dim RSCkk As New ADODB.Recordset
        LoadSqlData(" select * from Ap_balance_6_col   where ac_code =  '" & (RSCIn_M.Fields("Ac_Code").Value) & "' ", RSCkk)
        With RSCkk
            Do Until .EOF = True
                'MsgBox((RSCIn_M.Fields("Ac_Code").Value))
                CNN.Execute("Insert into Ap_Rpt_Incon_Detail (Ac_Code_Parent , Rpt_Id , Ac_Code , Ac_Name  ,  Last_Amt_Dr , Last_Amt_Cr, Amt_Dr , Amt_Cr , Rpt_Type ) values (  '" & CStr((RSCIn_M.Fields("Ac_Code").Value)) & "' , '" & CStr((RSCIn_M.Fields("Rpt_Id").Value)) & "' , '" & CStr((.Fields("Ac_Code").Value)) & "' , N'" & CStr((.Fields("Ac_Name").Value)) & "'  ,   " & CDbl((.Fields("Open_Amt_dr").Value)) & " , " & CDbl((.Fields("Open_Amt_cr").Value)) & " , " & CDbl((.Fields("Amt_dr").Value)) & " , " & CDbl((.Fields("Amt_cr").Value)) & " , 'Out' )")
                CNN.Execute("update  Ap_Rpt_Income_Item set Last_Amt_Dr  =  Last_Amt_Dr+" & CDbl((.Fields("Open_Amt_dr").Value)) & " , Last_Amt_Cr  = Last_Amt_Cr+" & CDbl((.Fields("Open_Amt_cr").Value)) & "  , Amt_Dr  =  Amt_Dr+" & CDbl((.Fields("Amt_dr").Value)) & " , Amt_Cr  = Amt_Cr+" & CDbl((.Fields("Amt_cr").Value)) & "   where ac_code = '" & (RSCIn_M.Fields("ac_code").Value) & "' And  Rpt_Type = 'Out' ")

                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub UpdateOut_ItemLast()
        Dim RSCkk As New ADODB.Recordset
        LoadSqlData(" select * from Ap_balance_6_col   where ac_code =  '" & (RSCIn_M.Fields("Ac_Code").Value) & "' ", RSCkk)
        With RSCkk
            Do Until .EOF = True
                CNN.Execute("Insert into Ap_Rpt_Cashflow_Detail (Ac_Code_Parent , Rpt_Id , Ac_Code , Ac_Name  , Last_Amt_Dr , Last_Amt_Cr , Rpt_Type) values (  '" & CStr((RSCIn_M.Fields("Ac_Code").Value)) & "' , '" & CStr((RSCIn_M.Fields("Rpt_Id").Value)) & "' , '" & CStr((.Fields("Ac_Code").Value)) & "' , N'" & CStr((.Fields("Ac_Name").Value)) & "'  , " & CDbl((.Fields("Rem_dr").Value)) & " , " & CDbl((.Fields("Rem_cr").Value)) & ", 'Out')")
                CNN.Execute("update  Ap_Rpt_Cashflow_Item set Amt_Dr  =  Amt_Dr+" & CDbl((.Fields("Rem_dr").Value)) & " , Amt_Cr  = Amt_Cr+" & CDbl((.Fields("Rem_cr").Value)) & "   where ac_code = '" & (RSCIn_M.Fields("ac_code").Value) & "' And  Rpt_Type = 'Out' ")
                .MoveNext()
            Loop
        End With

    End Sub



    'Private Sub UpdateOutLast()
    '    Dim RSC As New ADODB.Recordset
    '    LoadSqlData("select Rpt_ID, sum(Last_Amt_Dr) As Last_Amt_Dr , sum(Last_Amt_Cr) As Last_Amt_Cr  from Ap_Rpt_Cashflow_Item  where  Rpt_Type = 'Out' group by Rpt_ID ", RSC)
    '    With RSC
    '        Do Until .EOF = True
    '            CNN.Execute("Update Ap_Rpt_Cashflow set " & _
    '                     " Last_Amt ='" & CDbl(CDbl((.Fields("Last_Amt_cr").Value)) - CDbl((.Fields("Last_Amt_dr").Value))) & "' " & _
    '                        " where Rpt_ID = '" & (.Fields("Rpt_ID").Value) & "' ")
    '            .MoveNext()
    '        Loop
    '    End With
    'End Sub

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
        LoadSqlData("select Rpt_ID, sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr  from Ap_Rpt_Cashflow_Item  where  Rpt_Type = 'Out' group by Rpt_ID ", RSC)
        With RSC
            Do Until .EOF = True
                CNN.Execute("Update Ap_Rpt_Cashflow set " & _
                         " Last_Amt ='" & CDbl(CDbl((.Fields("Amt_cr").Value)) - CDbl((.Fields("Amt_dr").Value))) & "' " & _
                            " where Rpt_ID = '" & (.Fields("Rpt_ID").Value) & "' ")
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
        If RM.Checked = True Then
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
        ElseIf RY.Checked = True Then
            yy.Enabled = True
            LoadYear()
        End If
    End Sub

    Private Sub LoadDay()
        MdStartDate = Format(CDate(Ds.Value), "dd-MM-yyyy")
        MdToDate = Format(CDate(Dt.Value), "dd-MM-yyyy")
        Lb.Text = "ແຕ່ວັນທີ " & MdStartDate & " ຫາວັນທີ " & MdToDate
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub
    Private Sub LoadMonth()
        '---------------------------------
        If DMonth.Text = "01" Then
            MdStartDate = Format(CDate("01/01/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/01/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "01"
        ElseIf DMonth.Text = "02" Then
            Dim Day As String
            Dim MM As Date
            Dim Fromm As Date
            MdStartDate = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
            Fromm = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
            MM = Format(CDate("01/3/" & Year(MdStartDate)), "dd-MM-yyyy")
            Day = DateDiff(DateInterval.Day, Fromm, MM)
            MdToDate = Format(CDate(Day & "/02" & "/" & Year(MdStartDate)), "dd-MM-yyyy")
            MonthLetter1 = "02"
            Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
        ElseIf DMonth.Text = "03" Then
            MdStartDate = Format(CDate("01/03/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/03/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "03"
        ElseIf DMonth.Text = "04" Then
            MdStartDate = Format(CDate("01/04/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/04/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "04"
        ElseIf DMonth.Text = "05" Then
            MdStartDate = Format(CDate("01/05/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/05/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "05"
        ElseIf DMonth.Text = "06" Then
            MdStartDate = Format(CDate("01/06/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/06/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "06"
        ElseIf DMonth.Text = "07" Then
            MdStartDate = Format(CDate("01/07/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/07/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "07"
        ElseIf DMonth.Text = "08" Then
            MdStartDate = Format(CDate("01/08/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/08/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "08"
        ElseIf DMonth.Text = "09" Then
            MdStartDate = Format(CDate("01/09/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/09/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "09"
        ElseIf DMonth.Text = "10" Then
            MdStartDate = Format(CDate("01/10/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/10/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "10"
        ElseIf DMonth.Text = "11" Then
            MdStartDate = Format(CDate("01/11/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/11/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "11"
        ElseIf DMonth.Text = "12" Then
            MdStartDate = Format(CDate("01/12/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/12/" & Year(Myy.Value)), "dd-MM-yyyy")
            MonthLetter1 = "12"
        End If
        '-----------------
        Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & "/" & Year(MdToDate)
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
                    CNN.Execute("Update Ap_Rpt_Cashflow_Item set amt_dr='" & CDbl(op_dr + amt_dr) - CDbl(op_cr + amt_cr) & "' , amt_cr='" & CDbl(0) & "' where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                End If
                If CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) >= 0 Then
                    CNN.Execute("Update Ap_Rpt_Cashflow_Item set amt_dr='" & CDbl(0) & "' , amt_cr='" & CDbl(op_cr + amt_cr) - CDbl(op_dr + amt_dr) & "' where Ac_code='" & (.Fields("Ac_Code").Value) & "'")
                End If
                'CNN.Execute("update Ap_Rpt_Cashflow_Item set amt_dr  =  " & CDbl((.Fields("amt_dr").Value)) & " , amt_cr  = " & CDbl((.Fields("amt_cr").Value)) & "   where Ac_code=  '" & (.Fields("Ac_code").Value) & "' ")
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
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
        LngId = "7066" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
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
        LngId = "7048" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
        'If RM.Checked = True Then
        '    LngId = "7060" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
        '    LngId = "7061" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        'ElseIf RP.Checked = True Then
        '    LngId = "7062" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
        '    LngId = "7063" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        'ElseIf RY.Checked = True Then
        '    LngId = "7064" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
        '    LngId = "7065" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        'End If
        'SLF = "SELECT  " & MuLngRpt & "  *   FROM Ap_Rpt_Amt_Status  "
        Call LoadLoGO()
        Dim aa As String
        aa = " insert into  "
        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            Dim s As String = " SELECT " & MuLngRpt & "  *  FROM Ap_Rpt_Amt_Status Order by   Rpt_Id asc  "
            .Open(s, CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With




        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim Rpt As New CryAmt_Status
        If MdShowLOGO = 1 Then
            Rpt.Subreports(0).SetDataSource(RsLOGO)
        End If
        Rpt.SetDataSource(Rs)
        FrmPreview.ReportViewer.ReportSource = Rpt
        'FrmPreview.ReportViewer.DisplayGroupTree = False
        FrmPreview.WindowState = FormWindowState.Maximized
        FrmPreview.Show()
        FrmPreview.Focus()

    End Sub


    Private Sub LoadReportItem()
        Dim RPT_ID As String
        RPT_ID = " "
        Call Office()
        MuLngRpt = RptSjOff
        'MuLngRpt = ""
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As    RptSjUd ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
        LngId = "7066" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
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
        LngId = "7048" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
        'If RM.Checked = True Then
        '    LngId = "7060" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
        '    LngId = "7061" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        'ElseIf RP.Checked = True Then
        '    LngId = "7062" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
        '    LngId = "7063" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        'ElseIf RY.Checked = True Then
        '    LngId = "7064" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
        '    LngId = "7065" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        'End If
        'SLF = "SELECT  " & MuLngRpt & "  *   FROM Ap_Rpt_Amt_Status  "
        Call LoadLoGO()
        Dim aa As String
        aa = " insert into  "
        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            Dim s As String = " SELECT " & MuLngRpt & "  *  FROM Ap_Rpt_Amt_Status Order by   Rpt_Id asc  "
            .Open(s, CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With




        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim Rpt As New CryAmt_Status
        If MdShowLOGO = 1 Then
            Rpt.Subreports(0).SetDataSource(RsLOGO)
        End If
        Rpt.SetDataSource(Rs)
        FrmPreview.ReportViewer.ReportSource = Rpt
        'FrmPreview.ReportViewer.DisplayGroupTree = False
        FrmPreview.WindowState = FormWindowState.Maximized
        FrmPreview.Show()
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
        LoadDay()
    End Sub

    Private Sub Dt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Dt.ValueChanged
        LoadDay()
    End Sub

    Private Sub yy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yy.ValueChanged
        Call LoadYear()
    End Sub

    Private Sub Toyy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Toyy.ValueChanged
        Call LoadYear()
    End Sub

    Private Sub RaParent_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RaParent.CheckedChanged

    End Sub

    Private Sub FmCashflow_statement_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Call MdiCNum()

    End Sub




    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        FmAmtStatus_Item.ShowDialog()
        FmAmtStatus_Item.Focus()
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
End Class