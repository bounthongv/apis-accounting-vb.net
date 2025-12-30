Public Class FmRpt_BLS_NEW
    Dim MdStartDate2, MdToDate2 As Date
    Dim r As String
    Dim CLT_Str, CLT_Last_Str As String
    Dim bls1 As String
    Dim MonthLetter1 As String
    Dim MdStartDate As Date
    Dim MdToDate As Date

    Dim MdStartDate_PRV As Date
    Dim MdToDate_PRV As Date


    Dim MdStartDate_Last As Date
    Dim MdToDate_Last As Date
    Dim ny, ly, n_L_y As String
    Dim MdStartDate_MM As Date
    Dim MdToDate_MM As Date
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
    Private Sub HeaDer()
        LoadSqlData("SELECT * FROM Header WHERE ID=N'B03' ", RSC)
        If RSC.RecordCount <> 0 Then
            If MuLng = "L" Then
                TxtHeader.Text = Trim(RSC.Fields("Nm").Value.ToString)
                TxtS1.Text = Trim(RSC.Fields("S1").Value.ToString)
                TxtS2.Text = Trim(RSC.Fields("S2").Value.ToString)
                TxtS3.Text = Trim(RSC.Fields("S3").Value.ToString)
                TxtS4.Text = Trim(RSC.Fields("S4").Value.ToString)
                TxtPP.Text = Trim(RSC.Fields("pp").Value.ToString)
            Else
                TxtHeader.Text = Trim(RSC.Fields("Nm").Value.ToString)
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
            LoadSqlData("SELECT * FROM Header WHERE ID=N'B03' ", RSC)
            If RSC.RecordCount = 0 Then
                CNN.Execute("INSERT INTO Header(ID,Nm,S1,S2,S3,S4,PP) " & _
                            " values('B03',N'" & TxtHeader.Text & "',N'" & TxtS1.Text & "',N'" & TxtS2.Text & "',N'" & TxtS3.Text & "',N'" & TxtS4.Text & "',N'" & TxtPP.Text & "') ")
            Else
                CNN.Execute("UPDATE Header set Nm=N'" & TxtHeader.Text & "',S1=N'" & TxtS1.Text & "',S2=N'" & TxtS2.Text & "',S3=N'" & TxtS3.Text & "',S4=N'" & TxtS4.Text & "',PP=N'" & TxtPP.Text & "' " & _
                            " where ID='B03' ")
            End If
        Else
            LoadSqlData("SELECT * FROM Header WHERE ID=N'B03' ", RSC)
            If RSC.RecordCount = 0 Then
                CNN.Execute("INSERT INTO Header(ID,Nm,S1e,S2e,S3e,S4e,PPe) " & _
                            " values('B03',N'" & TxtHeader.Text & "',N'" & TxtS1.Text & "',N'" & TxtS2.Text & "',N'" & TxtS3.Text & "',N'" & TxtS4.Text & "',N'" & TxtPP.Text & "') ")
            Else
                CNN.Execute("UPDATE Header set Nm=N'" & TxtHeader.Text & "',S1e=N'" & TxtS1.Text & "',S2e=N'" & TxtS2.Text & "',S3e=N'" & TxtS3.Text & "',S4e=N'" & TxtS4.Text & "',PPe=N'" & TxtPP.Text & "' " & _
                            " where ID='B03' ")
            End If
        End If

    End Sub
    Private Sub FmCashflow_statement_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        HeaDer()
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
        selectMMM()
        Call selectLoad()
        Call Click_Last()
        SetControlText(Me)
        Call loadOffice_User()
        CMB_Curr.Items.Clear()
        CMB_Curr.Items.Add("EQVL")
        Call load_Cmb(" SELECT Curr  FROM Curr_For_Rate WHERE (Curr='LAK' Or Curr='THB Or Curr='USD')  ORDER BY cnt ", "Curr", CMB_Curr)
        If CMB_Curr.Items.Count > 0 Then
            CMB_Curr.SelectedIndex = 0
        End If

        If MuLng = "L" Then

            Label10.Text = "ລາຍເຊັນ1"
            Label14.Text = "ລາຍເຊັນ2"
            Label13.Text = "ລາຍເຊັນ3"
            Label12.Text = "ລາຍເຊັນ4"
            Label11.Text = "ທີ່"
            Label15.Text = "ອັດຕາຜ່ານມາ"
            Ct.Items.Clear()
            Ct.Items.Add("6 ເດືອນຕົ້ນປີ")
            Ct.Items.Add("6 ເດືອນທ້າຍປີ")
            CheckBox6.Text = "ທຽບເທົ່າໂດລາ"

        Else
            Label15.Text = "Rate Prev"
            Label10.Text = "Signature1"
            Label14.Text = "Signature2"
            Label13.Text = "Signature3"
            Label12.Text = "Signature4"
            Label11.Text = "Location"
            CheckBox6.Text = "EQVL USD"

            Ct.Items.Clear()
            Ct.Items.Add("First half year")
            Ct.Items.Add("Second half year")
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

        CNN.Execute("DELETE  Ap_balance_6_col ")
        CNN.Execute("DELETE FROM Ap_balance_6 ")
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
        " select ac_code  ,  0 , 0  , 0  , 0  , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook & "  group BY ac_code ")
        Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
        Dim AA As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr      ) " & _
            " select  ac_code  ,  0 , 0   , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  , 0 , 0   from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook & "   group BY ac_code"
        CNN.Execute(AA)

        CNN.Execute("INSERT INTO Ap_balance_6 (  ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
        " select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr   ,  0 , 0  , 0 , 0  from Open_jn WHERE    date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook & "   group BY ac_code")
        CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr  ) " & _
        " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr) as open_amt_cr , sum(Amt_Last_M_Dr) as Amt_Last_M_Dr , sum(Amt_Last_M_Cr) as Amt_Last_M_Cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")

        CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        Call Chang_Incom()
        CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + Amt_Last_M_Dr + amt_dr) - (open_amt_cr + Amt_Last_M_Cr + amt_cr) where (open_amt_dr + Amt_Last_M_Dr + amt_dr) >= (open_amt_cr + Amt_Last_M_Cr + amt_cr) ")
        CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr +  Amt_Last_M_Cr + amt_cr) - (open_amt_dr + Amt_Last_M_Dr + amt_dr) where (open_amt_cr + Amt_Last_M_Cr + amt_cr) >= (open_amt_dr + Amt_Last_M_Dr + amt_dr) ")
        CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")

    End Sub

    Private Sub Chang_Incom()
        If MDACC00 = 0 Then
            New_Code = New_Code
            '            Insr = "delete  Ap_balance_6  " & _
            '   "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  )   select '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr) from Ap_balance_6_col where  left(Ac_Code,1) ='" & Code_Dr & "'   Or left(Ac_Code,1)='" & Code_Cr & "'   or  Ac_Code =  '" & New_Code & "'    " & _
            '"update  Ap_balance_6 set open_amt_dr = 0 where open_amt_dr  is null  " & _
            '"update  Ap_balance_6 set open_amt_cr = 0 where open_amt_cr  is null  " & _
            '"update  Ap_balance_6 set amt_dr = 0 where amt_dr  is null  " & _
            '"update  Ap_balance_6 set amt_cr = 0 where amt_cr  is null   " & _
            '"Update  Ap_balance_6 set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr   " & _
            '"Update  Ap_balance_6 set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr " & _
            '"Update  Ap_balance_6 set   amt_dr = amt_dr  - amt_cr , amt_cr=0  where amt_dr  >= amt_cr   " & _
            '"Update  Ap_balance_6 set   amt_cr = amt_cr  - amt_dr , amt_dr=0  where amt_cr  >= amt_dr " & _
            '"delete  Ap_balance_6_col  where   left(Ac_Code,1) ='" & Code_Dr & "'  Or  left(Ac_Code,1)='" & Code_Cr & "' or   Ac_Code =  '" & New_Code & "'    " & _
            '"  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  from Ap_balance_6"
            '            CNN.Execute(Insr)
            If Month(MdStartDate) <> 12 Then
                Insr = "delete  Ap_balance_6  " & _
                 "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr)   select '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr) from Ap_balance_6_col where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "' " & _
        "update  Ap_balance_6 set open_amt_dr = 0 where open_amt_dr  is null  " & _
        "update  Ap_balance_6 set open_amt_cr = 0 where open_amt_cr  is null  " & _
        "update  Ap_balance_6 set amt_dr = 0 where amt_dr  is null  " & _
        "update  Ap_balance_6 set amt_cr = 0 where amt_cr  is null   " & _
         "Update  Ap_balance_6 set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr   " & _
        "Update  Ap_balance_6 set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr " & _
        "Update  Ap_balance_6 set   amt_dr = amt_dr  - amt_cr , amt_cr=0  where amt_dr  >= amt_cr   " & _
        "Update  Ap_balance_6 set   amt_cr = amt_cr  - amt_dr , amt_dr=0  where amt_cr  >= amt_dr " & _
           "delete  Ap_balance_6_col  where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "'   " & _
             "  insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr ,status )  " & _
    " select  '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr),1 from Ap_balance_6_col where Ac_Code =  '" & New_Code & "' " & _
      "       delete  Ap_balance_6_col where Ac_Code =  '" & New_Code & "' " & _
    "  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , sum(open_amt_dr) , sum(open_amt_cr) , sum(amt_dr) , sum(amt_cr)  from Ap_balance_6 group by Ac_Code "

                CNN.Execute(Insr)

            Else
                Insr = "delete  Ap_balance_6  " & _
              "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr)   select '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr) from Ap_balance_6_col where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "' " & _
     "update  Ap_balance_6 set open_amt_dr = 0 where open_amt_dr  is null  " & _
     "update  Ap_balance_6 set open_amt_cr = 0 where open_amt_cr  is null  " & _
     "update  Ap_balance_6 set amt_dr = 0 where amt_dr  is null  " & _
     "update  Ap_balance_6 set amt_cr = 0 where amt_cr  is null   " & _
      "Update  Ap_balance_6 set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr   " & _
     "Update  Ap_balance_6 set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr " & _
     "Update  Ap_balance_6 set   amt_dr = amt_dr  - amt_cr , amt_cr=0  where amt_dr  >= amt_cr   " & _
     "Update  Ap_balance_6 set   amt_cr = amt_cr  - amt_dr , amt_dr=0  where amt_cr  >= amt_dr " & _
     " delete TEST_ABC insert into TEST_ABC (Rpt_Id,amt) select '4.1.6' ,amt_cr-amt_Dr  from Ap_balance_6  where Ac_Code =  '" & New_Code & "' " & _
        "delete  Ap_balance_6_col  where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "'   " & _
          "  insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr ,status )  " & _
    " select  '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr),1 from Ap_balance_6_col where Ac_Code =  '" & New_Code & "' " & _
    "       delete  Ap_balance_6_col where Ac_Code =  '" & New_Code & "' " & _
    "  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , sum(open_amt_dr) , sum(open_amt_cr) , sum(amt_dr) , sum(amt_cr)  from Ap_balance_6 group by Ac_Code "

                CNN.Execute(Insr)

            End If



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

        '   Insr = "delete  Ap_balance_6  " & _
        '       "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) values ('" & New_Code & "' ,0,0,0,0) " & _
        '    "update Ap_balance_6 set  open_amt_Dr =  " & _
        '    "(select top 1  (select  (Sum(open_amt_dr))-( Sum(open_amt_cr)) As open_amt_dr from Ap_balance_6_col where left(Ac_Code,1)='" & Code_Dr & "' )  As Dr " & _
        '    "from Ap_balance_6_col )  where  Ac_Code ='" & New_Code & "'  " & _
        '"update Ap_balance_6 set  open_amt_cr =  " & _
        '"(select top 1  (select  (Sum(open_amt_cr))-( Sum(open_amt_dr)) As open_amt_dr from Ap_balance_6_col where left(Ac_Code,1)='" & Code_Cr & "'  )  As Cr " & _
        ' "from Ap_balance_6_col ) where  Ac_Code ='" & New_Code & "'   " & _
        '"update Ap_balance_6 set  amt_Dr = " & _
        '"(select top 1  (select  (Sum(amt_dr))-( Sum(amt_cr)) As amt_dr from Ap_balance_6_col where left(Ac_Code,1)='" & Code_Dr & "'  )  As Dr " & _
        '"from Ap_balance_6_col )  where  Ac_Code ='" & New_Code & "'  " & _
        '"update Ap_balance_6 set  amt_cr =  " & _
        '"(select top 1  (select  (Sum(amt_cr))-( Sum(amt_dr)) As amt_dr from Ap_balance_6_col where  left(Ac_Code,1)='" & Code_Cr & "' )  As Cr " & _
        '"from Ap_balance_6_col ) where  Ac_Code ='" & New_Code & "'  " & _
        '"update  Ap_balance_6_col set open_amt_dr = 0 where open_amt_dr  is null  " & _
        '"update  Ap_balance_6_col set open_amt_cr = 0 where open_amt_cr  is null  " & _
        '"update  Ap_balance_6_col set amt_dr = 0 where amt_dr  is null  " & _
        '"update  Ap_balance_6_col set amt_cr = 0 where amt_cr  is null   " & _
        '" Update  Ap_balance_6 set   open_amt_dr = (open_amt_cr  - open_amt_dr ) , open_amt_cr=0  where (open_amt_cr  - open_amt_dr )>= 0 " & _
        ' "Update  Ap_balance_6 set   open_amt_cr = (open_amt_dr  - open_amt_cr) , open_amt_dr=0  where (open_amt_cr  - open_amt_dr )<= 0 " & _
        '"Update  Ap_balance_6 set   amt_dr = (amt_cr  - amt_dr ) , amt_cr=0  where (amt_cr  - amt_dr )>= 0 " & _
        '" Update  Ap_balance_6 set   amt_cr = (amt_dr  - amt_cr) , amt_dr=0  where (amt_cr  - amt_dr )<= 0 " & _
        ' "  update Ap_balance_6 set  Ap_balance_6.open_amt_dr = Ap_balance_6.open_amt_dr + Ap_balance_6_col.open_amt_dr , Ap_balance_6.open_amt_cr = Ap_balance_6.open_amt_cr + Ap_balance_6_col.open_amt_cr   ,  Ap_balance_6.amt_dr = Ap_balance_6.amt_dr + Ap_balance_6_col.amt_dr    ,  Ap_balance_6.amt_cr = Ap_balance_6.amt_cr + Ap_balance_6_col.amt_cr    from Ap_balance_6 , Ap_balance_6_col   where  Ap_balance_6.Ac_Code = Ap_balance_6_col.Ac_Code      " & _
        ' "Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr " & _
        '  "Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr " & _
        '  "delete  Ap_balance_6_col   where left(Ac_Code,1)='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "'  Or Ac_Code = '" & New_Code & "' " & _
        '   "  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr ,   amt_cr , amt_dr from Ap_balance_6"
        '   CNN.Execute(Insr)
    End Sub
    Private Sub SelcectIn_BLS()

        'CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
        '         "where left(Ap_Rpt_BLS_Item_Old.ac_code,7) =left(Ap_balance_6_col.ac_code,7) And  Rpt_Type = 'In'")

        If CMB_Curr.SelectedIndex <> 2 Then
            CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
        "where left(Ap_Rpt_BLS_Item_Old.ac_code,7) = left(Ap_balance_6_col.ac_code,7) And  Rpt_Type = 'In'  ")

        Else
            CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
                    "where left(Ap_Rpt_BLS_Item_Old.ac_code,7) = left(Ap_balance_6_col.ac_code,7) And  Rpt_Type = 'In' and left(Ap_balance_6_col.ac_code,7)<>'2382120' ")

        End If

        CNN.Execute("Insert into Ap_Rpt_BLS_Detail (  Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type )" & _
         " select   Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type from Ap_Rpt_BLS_Item_Old where Rpt_Type = 'In' And ( Amt_Dr <>0 or Amt_Cr <>0  or Last_Amt_Dr <>0 or Last_Amt_Cr <>0 )")

    End Sub

    Private Sub UpdateIIn_BLS()
        CNN.Execute("delete Ap_Rpt_BLS_Stock ")
        CNN.Execute(" insert into Ap_Rpt_BLS_Stock ( Rpt_ID , Last_Amt_Dr , Last_Amt_Cr, Amt_Dr , Amt_Cr)" & _
                     "  select Rpt_ID , sum(Last_Amt_Dr) As Last_Amt_Dr , sum(Last_Amt_Cr) As Last_Amt_Cr ,sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr   from Ap_Rpt_BLS_Item_Old  where  Rpt_Type = 'In' group by Rpt_ID")
        CNN.Execute("Update Ap_Rpt_BLS_Old set Amt = Ap_Rpt_BLS_Stock.Amt_Dr-Ap_Rpt_BLS_Stock.Amt_cr ,Last_Amt =Ap_Rpt_BLS_Stock.Last_Amt_dr-Ap_Rpt_BLS_Stock.Last_Amt_Cr  from Ap_Rpt_BLS_Old ,Ap_Rpt_BLS_Stock where  Ap_Rpt_BLS_Old.Rpt_ID=Ap_Rpt_BLS_Stock.Rpt_ID")
    End Sub

    Private Sub SelectOut_BLS()
        'CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
        '          "where left(Ap_Rpt_BLS_Item_Old.ac_code,7) =left(Ap_balance_6_col.ac_code,7) And  Rpt_Type = 'Out'")
        If CMB_Curr.SelectedIndex = 2 Then
            CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
   "where left(Ap_Rpt_BLS_Item_Old.ac_code,7) =left(Ap_balance_6_col.ac_code,7) And   Rpt_Type = 'Out'   ")
            '            CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
            '"where left(Ap_Rpt_BLS_Item_Old.ac_code,7) =left(Ap_balance_6_col.ac_code,7) And  Rpt_Type = 'In'  and left(Ap_balance_6_col.ac_code,7)='2382120'  ")
            '            CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
            '    "where left(Ap_Rpt_BLS_Item_Old.ac_code,7) = left(Ap_balance_6_col.ac_code,7) And  Rpt_Type = 'In'  ")
        ElseIf CMB_Curr.SelectedIndex = 0 Then
            CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
               "where left(Ap_Rpt_BLS_Item_Old.ac_code,7) =left(Ap_balance_6_col.ac_code,7) And  Rpt_Type = 'Out' and left(Ap_balance_6_col.ac_code,7)<>'2382120' ")
        Else
            CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
                 "where left(Ap_Rpt_BLS_Item_Old.ac_code,7) =left(Ap_balance_6_col.ac_code,7) And  Rpt_Type = 'Out'   and left(Ap_balance_6_col.ac_code,7)<>'2382120' ")

        End If


        CNN.Execute("Insert into Ap_Rpt_BLS_Detail (  Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type )" & _
         " select   Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type from Ap_Rpt_BLS_Item_Old where Rpt_Type = 'Out' And ( Amt_Dr <>0 or Amt_Cr <>0  or Last_Amt_Dr <>0 or Last_Amt_Cr <>0 )")

        'LoadSqlData("select * from Ap_Rpt_BLS_Item_Old where  Rpt_Type = 'Out' ", RSCIn_M)
        'With RSCIn_M
        '    Do Until .EOF = True
        '        Call UpdateOut_Item()
        '        .MoveNext()
        '    Loop
        'End With
        'If RSCIn_M.State = ConnectionState.Open Then RSCIn_M.Close()
    End Sub

    Private Sub UpdateOut_BLS()
        CNN.Execute("delete Ap_Rpt_BLS_Stock ")
        CNN.Execute(" insert into Ap_Rpt_BLS_Stock ( Rpt_ID , Last_Amt_Dr , Last_Amt_Cr, Amt_Dr , Amt_Cr)" & _
                     "  select Rpt_ID , sum(Last_Amt_Dr) As Last_Amt_Dr , sum(Last_Amt_Cr) As Last_Amt_Cr ,sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr   from Ap_Rpt_BLS_Item_Old  where  Rpt_Type = 'Out' group by Rpt_ID")
        CNN.Execute("Update Ap_Rpt_BLS_Old set Amt = Ap_Rpt_BLS_Stock.Amt_Dr-Ap_Rpt_BLS_Stock.Amt_Dr ,Last_Amt =Ap_Rpt_BLS_Stock.Last_Amt_dr-Ap_Rpt_BLS_Stock.Last_Amt_Cr  from Ap_Rpt_BLS_Old ,Ap_Rpt_BLS_Stock where  Ap_Rpt_BLS_Old.Rpt_ID=Ap_Rpt_BLS_Stock.Rpt_ID")
        CNN.Execute("Update Ap_Rpt_BLS_Old set Amt = Ap_Rpt_BLS_Stock.Amt_Cr-Ap_Rpt_BLS_Stock.Amt_Dr ,Last_Amt =Ap_Rpt_BLS_Stock.Last_Amt_Cr-Ap_Rpt_BLS_Stock.Last_Amt_Dr  from Ap_Rpt_BLS_Old ,Ap_Rpt_BLS_Stock where  Ap_Rpt_BLS_Old.Rpt_ID=Ap_Rpt_BLS_Stock.Rpt_ID")
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
    Private Sub Update_Sum_BLS()
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

        'CNN.Execute("Update Ap_Rpt_BLS_Old set Amt=0  where Rpt_Id='4.1.6' ")
        'CNN.Execute("Update Ap_Rpt_BLS_Old set Last_Amt=0  where Rpt_Id='4.1.6' ")


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
    Private Sub BLS()
        Call Office()
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        'Call ChangBalance()
        BLNEW()
        CNN.Execute("update  Ap_Rpt_BLS_Item_Old set Last_amt_dr  =  " & CDbl(0) & " , Last_amt_cr  = " & CDbl(0) & " ,amt_dr  =  " & CDbl(0) & " , amt_cr  = " & CDbl(0) & "  ")
        CNN.Execute("update Ap_Rpt_BLS_Old set Last_Amt = 0 , Amt  =  " & CDbl(0) & "  ")
        CNN.Execute("DELETE FROM Ap_Rpt_BLS_Detail ")
        SelcectIn_BLS()
        UpdateIIn_BLS()
        SelectOut_BLS()
        UpdateOut_BLS()
        Update_Sum_BLS()
    End Sub
    Private Sub BLNEW()


        New_Code = "3901000"
        New_Code4 = "00.3901000"
        New_Code = "3901000"
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
        ''        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ''               " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  1=1  and Curr=N'LAK'   and   gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")

        ''        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ''   " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amount_Dr)* " & CDbl(txtRate.Text) & " as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & " as amt_cr  from gen_jn  WHERE  1=1  and Curr=N'USD'   and   gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")

        ''        Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
        ''        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ''        " select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE  1=1   and Curr=N'LAK'   and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        ''        Dim KK As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ''      " select ac_code , sum(amount_Dr)* " & CDbl(txtRate.Text) & " as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & " as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE  1=1   and Curr=N'USD'   and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code"
        ''        CNN.Execute(KK)


        ''        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ''        " select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1  and Curr=N'LAK' and     date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        ''        '       CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ''        '" select ac_code  , sum(amount_Dr)* " & CDbl(txtRate.Text) & "  as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & "  as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1  and Curr=N'USD'  and date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        ''        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        ''" select ac_code  , sum(amt_dr)  as amt_dr , sum(amt_cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1  and Curr=N'USD'  and date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

        Dim GGG As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
              " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(Amount_Dr)as amt_dr , sum(Amount_cr)as amt_cr  from gen_jn  WHERE 1=1  and Curr=N'LAK'   and gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code "
        CNN.Execute(GGG)

        Dim USD As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
 " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amount_Dr)* " & CDbl(txtRate.Text) & "  as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & " as amt_cr  from gen_jn  WHERE 1=1 and Curr=N'USD'  and gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code "
        CNN.Execute(USD)

        Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)


        'CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        '" select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        '=======LAK===
        Dim PPP As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code , sum(Amount_Dr)as amt_dr , sum(Amount_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE 1=1  and Curr=N'LAK'   and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code"
        CNN.Execute(PPP)
        Dim PPPUSD As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
" select ac_code , sum(amount_Dr)* " & CDbl(txtRate.Text) & " as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & " as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE 1=1  and Curr=N'USD'  and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code"
        CNN.Execute(PPPUSD)

        '        '=======LAK===
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code  , sum(Amount_Dr) as amt_dr , sum(Amount_cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1   and Curr=N'LAK'  and date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")


        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
   " select ac_code  , sum(amount_Dr)* " & CDbl(txtRate.Text) & "  as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & "  as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE  1=1  and Curr=N'USD'  and   date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")


        CNN.Execute("UPDATE Ap_balance_6 set Ac_Code = left(Ac_Code,7) ")


        CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr)as open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")
        '    If CMB_Curr.SelectedIndex = 0 Then
        '        CNN.Execute(" insert into Ap_balance_6_col(ac_code, ac_name, open_amt_dr, open_amt_cr,  amt_dr, amt_cr,   Rem_dr, Rem_cr,Status) " & _
        '" Select '2382120.00.0000',N'ຄູ່ມູນຄ່າຖານະ ແລກປ່ຽນເງິນຕາ ຕ່າງປະເທດ',sum(open_amt_dr),sum(open_amt_cr),sum(amt_dr),sum(amt_cr),sum(Rem_dr),sum(Rem_cr),1   " & _
        '" from Ap_balance_6_col where left(ac_code,7)='2382120' group by  left(ac_code,7) ")
        '        CNN.Execute("delete  Ap_balance_6_col  where left(ac_code,7)='2382120' and Status is null ")
        '    End If
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

        'If Month(MdStartDate) = 12 Then
        '    CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (amt_dr) - (amt_cr) where (amt_dr) >= (amt_cr) ")
        '    CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (amt_cr) - (amt_dr) where (amt_cr) >= (amt_dr) ")

        'End If

        CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")




        '        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        '" select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  1=1 " & B_Curr & " and   gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")
        '        Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
        '        Dim KK As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        '         " select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE  1=1 " & B_Curr & " and gen_jn.date_work   BETWEEN '" & Format(S, "yyyy-MM-01") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code"
        '        CNN.Execute(KK)

        '        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        '        " select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1 " & B_Curr & " and     date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

        '        CNN.Execute("UPDATE Ap_balance_6 set Ac_Code = left(Ac_Code,7) ")

        '        CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        '        " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr)as open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")

        '        CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        '        CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        '        Call Chang_Incom()
        '        CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + amt_dr) - (open_amt_cr + amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr + amt_cr) ")
        '        CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr + amt_cr) - (open_amt_dr + amt_dr) where (open_amt_cr + amt_cr) >= (open_amt_dr + amt_dr) ")
        '        CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")

    End Sub

    Private Sub BLNEW22()


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
            If Month(MdStartDate) = 1 Then
                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                      " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  1=1  and Curr=N'LAK'   and   month(gen_jn.date_work)='12' and Year(gen_jn.date_work)='" & Year(MdStartDate) - 1 & "'  " & Ac_Code & " " & MULook2 & "  group BY ac_code ")

                CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
           " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amount_Dr)* " & CDbl(txtRate.Text) & " as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & " as amt_cr  from gen_jn  WHERE  1=1  and Curr=N'USD'   and     month(gen_jn.date_work)='12' and Year(gen_jn.date_work)='" & Year(MdStartDate) - 1 & "'  " & Ac_Code & " " & MULook2 & "  group BY ac_code ")

            End If

            ''            Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
            ''            CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            ''            " select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE  1=1   and Curr=N'LAK'   and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
            ''            CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            ''       " select ac_code , sum(amount_Dr)* " & CDbl(txtRate.Text) & " as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & " as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE  1=1   and Curr=N'USD'   and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")


            ''            CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            ''            " select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1  and Curr=N'LAK' and     date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
            ''            '       CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            ''            '" select ac_code  , sum(amount_Dr)* " & CDbl(txtRate.Text) & "  as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & "  as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1  and Curr=N'USD'  and date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
            ''            CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            ''" select ac_code  , sum(amt_dr)  as amt_dr , sum(amt_cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1  and Curr=N'USD'  and date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")


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
        '=============
        Insr = "delete  Ap_balance_6  " & _
             "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr)   select '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr) from Ap_balance_6_col where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "' " & _
    "update  Ap_balance_6 set open_amt_dr = 0 where open_amt_dr  is null  " & _
    "update  Ap_balance_6 set open_amt_cr = 0 where open_amt_cr  is null  " & _
    "update  Ap_balance_6 set amt_dr = 0 where amt_dr  is null  " & _
    "update  Ap_balance_6 set amt_cr = 0 where amt_cr  is null   " & _
     "Update  Ap_balance_6 set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr   " & _
    "Update  Ap_balance_6 set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr " & _
    "Update  Ap_balance_6 set   amt_dr = amt_dr  - amt_cr , amt_cr=0  where amt_dr  >= amt_cr   " & _
    "Update  Ap_balance_6 set   amt_cr = amt_cr  - amt_dr , amt_dr=0  where amt_cr  >= amt_dr " & _
    " delete TEST_ABC insert into TEST_ABC (Rpt_Id,amt) select '4.1.6' ,amt_cr-amt_Dr  from Ap_balance_6  where Ac_Code =  '" & New_Code & "' " & _
       "delete  Ap_balance_6_col  where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "'   " & _
         "  insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr ,status )  " & _
   " select  '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr),1 from Ap_balance_6_col where Ac_Code =  '" & New_Code & "' " & _
   "       delete  Ap_balance_6_col where Ac_Code =  '" & New_Code & "' " & _
   "  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , sum(open_amt_dr) , sum(open_amt_cr) , sum(amt_dr) , sum(amt_cr)  from Ap_balance_6 group by Ac_Code "

        CNN.Execute(Insr)
        '===============
        CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + amt_dr) - (open_amt_cr + amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr + amt_cr) ")
        CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr + amt_cr) - (open_amt_dr + amt_dr) where (open_amt_cr + amt_cr) >= (open_amt_dr + amt_dr) ")
        CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        'CNN.Execute("UPDATE Ap_balance_6_col set Ac_Code = left(Ac_Code,7) ")
    End Sub

    Private Sub BtnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        AddHeader()
        If CMB_Curr.Text = "LAK" Then
            CURR01 = "ຫົວໜ່ວຍ : ກີບ"
        ElseIf CMB_Curr.Text = "USD" Then
            CURR01 = "ຫົວໜ່ວຍ : ໂດລາ"
        Else
            CURR01 = "ຫົວໜ່ວຍ : ກີບ"
        End If
        If CheckBox1.Checked = False Then

            CaCashflow()

            BLS()
            'BLNEW()
            If Month(MdStartDate) <> 12 Then
                BLNEW22()
            End If

            'Chang_Incom12()

            CNN.Execute("update Ap_Rpt_Cashflow set Ap_Rpt_Cashflow.amt=(select   amt-Last_amt from Ap_Rpt_BLS_Old where Rpt_Id='4.1.6') where Rpt_Id='02' ")
            CNN.Execute("update Ap_Rpt_Cashflow set Ap_Rpt_Cashflow.Last_amt=(select   Last_amt from Ap_Rpt_BLS_Old where Rpt_Id='4.1.6') where Rpt_Id='02' ")
            If RM.Checked = True Then
                If Month(MdStartDate) = 1 Then
                    CNN.Execute("   update Ap_Rpt_Cashflow set Ap_Rpt_Cashflow.Last_amt=(select   amt  from TEST_ABC where Rpt_Id='4.1.6') where Rpt_Id='02'")
                End If
                If Month(MdStartDate) = 12 Then
                    CNN.Execute("   update Ap_Rpt_Cashflow set Ap_Rpt_Cashflow.amt=(select   amt  from TEST_ABC where Rpt_Id='4.1.6') where Rpt_Id='02'")
                End If
                'CaCashflow()
                If DMonth.SelectedIndex >= 1 Then
                    BLS_PreV()

                    'CNN.Execute("update Ap_Rpt_Cashflow set Ap_Rpt_Cashflow.amt=(select   amt-Last_amt from Ap_Rpt_BLS_Old where Rpt_Id='4.1.6') where Rpt_Id='02' ")
                    'CNN.Execute("UPDATE Ap_Rpt_BLS_Old set  ")
                    CNN.Execute("update Ap_Rpt_Cashflow set Ap_Rpt_Cashflow.Last_amt= (select   amt-Last_amt from Ap_Rpt_BLS_Old where Rpt_Id='4.1.6') where Rpt_Id='02' ")

                End If
                Dim sa As String = "update Ap_Rpt_Cashflow set  Amt=(    select sum(amt_cr)-sum(amt_dr) from gen_jn   where left(ac_code,3)='144'  and date_work BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  " & MULook2 & " ) where Rpt_Id='33'"
                CNN.Execute(sa)
                sa = "update Ap_Rpt_Cashflow set  Last_amt=(   select sum(amt_cr)-sum(amt_dr) from gen_jn   where left(ac_code,3)='144'  and date_work BETWEEN '" & Format(MdStartDate_PRV, "yyyy-MM-dd") & "' AND '" & Format(MdToDate_PRV, "yyyy-MM-dd") & "'  " & MULook2 & " ) where Rpt_Id='33'"
                CNN.Execute(sa)

            End If

            If RP.Checked = True Then
                Call BLS_LAST()
                If Period.SelectedIndex = 0 Then

                    Dim LAST As String = "update Ap_Rpt_Cashflow set  Last_amt=(   select sum(amt_cr)-sum(amt_dr) from gen_jn   where (left(ac_code,1)='4' or left(ac_code,1)='5' )  and date_work BETWEEN '" & Format(MdStartDate_PRV, "yyyy-MM-dd") & "' AND '" & Format(MdToDate_PRV, "yyyy-MM-dd") & "'  " & MULook2 & " ) where Rpt_Id='02'"
                    CNN.Execute(LAST)

                ElseIf Period.SelectedIndex > 1 Then
                    Dim PPP3 As String = "update Ap_Rpt_Cashflow set  amt=(   select sum(amt_cr)-sum(amt_dr) from gen_jn   where (left(ac_code,1)='4' or left(ac_code,1)='5' )  and date_work BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  " & MULook2 & " ) where Rpt_Id='02'"
                    CNN.Execute(PPP3)
                    MdStartDate2 = DateAdd(DateInterval.Month, -3, MdStartDate)
                    MdToDate2 = DateAdd(DateInterval.Day, -1, MdStartDate)
                    Dim LAST As String = "update Ap_Rpt_Cashflow set  Last_amt=(   select sum(amt_cr)-sum(amt_dr) from gen_jn   where (left(ac_code,1)='4' or left(ac_code,1)='5' )  and date_work BETWEEN '" & Format(MdStartDate2, "yyyy-MM-dd") & "' AND '" & Format(MdToDate2, "yyyy-MM-dd") & "'  " & MULook2 & " ) where Rpt_Id='02'"
                    CNN.Execute(LAST)
                End If
            End If

            If RT.Checked = True Then
                If Ct.SelectedIndex = 0 Then
                    Dim yearamt As String = "update Ap_Rpt_Cashflow set  amt=(   select sum(amt_cr)-sum(amt_dr) from gen_jn   where (left(ac_code,1)='4' or left(ac_code,1)='5' )   and   date_work BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  " & MULook2 & " ) where Rpt_Id='02'"
                    CNN.Execute(yearamt)
                    Dim LAST_year As String = "update Ap_Rpt_Cashflow set  Last_amt=(   select sum(amt_cr)-sum(amt_dr) from gen_jn   where (left(ac_code,1)='4' or left(ac_code,1)='5' )  and date_work BETWEEN '" & Format(MdStartDate_MM, "yyyy-MM-dd") & "' AND '" & Format(MdToDate_MM, "yyyy-MM-dd") & "'  " & MULook2 & " ) where Rpt_Id='02'"
                    CNN.Execute(LAST_year)
                Else
                    Dim yearamt As String = "update Ap_Rpt_Cashflow set  amt=(   select sum(amt_cr)-sum(amt_dr) from gen_jn   where (left(ac_code,1)='4' or left(ac_code,1)='5' )   and   date_work BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  " & MULook2 & " ) where Rpt_Id='02'"
                    CNN.Execute(yearamt)
                    Dim LAST_year As String = "update Ap_Rpt_Cashflow set  Last_amt=(   select sum(amt_cr)-sum(amt_dr) from gen_jn   where (left(ac_code,1)='4' or left(ac_code,1)='5' )  and   date_work BETWEEN '" & Format(MdStartDate_MM, "yyyy-MM-dd") & "' AND '" & Format(MdToDate_MM, "yyyy-MM-dd") & "'  " & MULook2 & " ) where Rpt_Id='02'"
                    CNN.Execute(LAST_year)
                End If

            End If


            If RY.Checked = True Then
                Dim yearamt As String = "update Ap_Rpt_Cashflow set  amt=(   select sum(amt_cr)-sum(amt_dr) from gen_jn   where (left(ac_code,1)='4' or left(ac_code,1)='5' )   and year(date_work)='" & Year(MdStartDate) & "'  " & MULook2 & " ) where Rpt_Id='02'"
                CNN.Execute(yearamt)
                Dim LAST_year As String = "update Ap_Rpt_Cashflow set  Last_amt=(   select sum(amt_cr)-sum(amt_dr) from gen_jn   where (left(ac_code,1)='4' or left(ac_code,1)='5' )  and year(date_work)='" & Year(MdStartDate) - 1 & "'  " & MULook2 & " ) where Rpt_Id='02'"
                CNN.Execute(LAST_year)
            End If

            Dim MI As String = "delete AP_Sum "
            MI = MI & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select '10' ,amt,0 from Ap_Rpt_Cashflow     where rpt_ID='02'  "
            MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '03'  ,0 ,amt from Ap_Rpt_Cashflow     where rpt_ID='03'  "
            MI = MI & "  update AP_Sum set amt=(select sum(amt2)+sum(amt3) from AP_Sum where  (rpt_ID='10' or rpt_ID='03')  ) where rpt_ID='10'   "
            MI = MI & "  delete  AP_Sum  where rpt_ID='03'   "
            MI = MI & " Update Ap_Rpt_Cashflow set Ap_Rpt_Cashflow.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Cashflow    where Ap_Rpt_Cashflow.rpt_ID=AP_Sum.rpt_ID "
            CNN.Execute(MI)

            Call LoadReport()
        Else

            CaCashflow()
            BLS()
            CNN.Execute("update Ap_Rpt_Cashflow set Ap_Rpt_Cashflow.amt=(select   amt-Last_amt from Ap_Rpt_BLS_Old where Rpt_Id='4.1.5') where Rpt_Id='02' ")
            CNN.Execute("update Ap_Rpt_Cashflow set Ap_Rpt_Cashflow.Last_amt=(select  Last_amt from Ap_Rpt_BLS_Old where Rpt_Id='4.1.5') where Rpt_Id='02' ")
            'CaCashflow()

            Dim MI As String = "delete AP_Sum "
            MI = MI & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select '10' ,amt,0 from Ap_Rpt_Cashflow     where rpt_ID='02'  "
            MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '03'  ,0 ,amt from Ap_Rpt_Cashflow     where rpt_ID='03'  "
            MI = MI & "  update AP_Sum set amt=(select sum(amt2)+sum(amt3) from AP_Sum where  (rpt_ID='10' or rpt_ID='03')  ) where rpt_ID='10'   "
            MI = MI & "  delete  AP_Sum  where rpt_ID='03'   "
            MI = MI & " Update Ap_Rpt_Cashflow set Ap_Rpt_Cashflow.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Cashflow    where Ap_Rpt_Cashflow.rpt_ID=AP_Sum.rpt_ID "
            CNN.Execute(MI)

            Dim Last_amt As String = "delete AP_Sum "
            Last_amt = Last_amt & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select '10' ,Last_amt,0 from Ap_Rpt_Cashflow     where rpt_ID='02'  "
            Last_amt = Last_amt & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '03'  ,0 ,Last_amt from Ap_Rpt_Cashflow     where rpt_ID='03'  "
            Last_amt = Last_amt & "  update AP_Sum set amt=(select sum(amt2)+sum(amt3) from AP_Sum where  (rpt_ID='10' or rpt_ID='03')  ) where rpt_ID='10'   "
            Last_amt = Last_amt & "  delete  AP_Sum  where rpt_ID='03'   "
            Last_amt = Last_amt & " Update Ap_Rpt_Cashflow set Ap_Rpt_Cashflow.Last_amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Cashflow    where Ap_Rpt_Cashflow.rpt_ID=AP_Sum.rpt_ID "
            CNN.Execute(Last_amt)
            '========29	ກະແສເງິນສົດສຸດທິຈາກກິດຈະກຳດຳເນີນງານ 
            Dim MI29 As String = "delete AP_Sum "
            MI29 = MI29 & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select '29' ,0,0 from Ap_Rpt_Cashflow     where rpt_ID='29'  "
            MI29 = MI29 & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '111'  ,0 ,amt from Ap_Rpt_Cashflow     where rpt_ID='10'  "
            MI29 = MI29 & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '111'  ,0 ,amt from Ap_Rpt_Cashflow     where rpt_ID='20'  "
            MI29 = MI29 & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '111'  ,0 ,amt from Ap_Rpt_Cashflow     where rpt_ID='28'  "
            MI29 = MI29 & "  update AP_Sum set amt=(select sum(amt2)+sum(amt3) from AP_Sum where  (rpt_ID='29' or rpt_ID='111')  ) where rpt_ID='29'   "
            MI29 = MI29 & "  delete  AP_Sum  where rpt_ID='111'   "
            MI29 = MI29 & " Update Ap_Rpt_Cashflow set Ap_Rpt_Cashflow.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Cashflow    where Ap_Rpt_Cashflow.rpt_ID=AP_Sum.rpt_ID "
            CNN.Execute(MI29)

            Dim MI29Last_amt As String = "delete AP_Sum "
            MI29Last_amt = MI29Last_amt & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select '29' ,0,0 from Ap_Rpt_Cashflow     where rpt_ID='29'  "
            MI29Last_amt = MI29Last_amt & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '111'  ,0 ,Last_amt from Ap_Rpt_Cashflow     where rpt_ID='10'  "
            MI29Last_amt = MI29Last_amt & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '111'  ,0 ,Last_amt from Ap_Rpt_Cashflow     where rpt_ID='20'  "
            MI29Last_amt = MI29Last_amt & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '111'  ,0 ,Last_amt from Ap_Rpt_Cashflow     where rpt_ID='28'  "
            MI29Last_amt = MI29Last_amt & "  update AP_Sum set amt=(select sum(amt2)+sum(amt3) from AP_Sum where  (rpt_ID='29' or rpt_ID='111')  ) where rpt_ID='29'   "
            MI29Last_amt = MI29Last_amt & "  delete  AP_Sum  where rpt_ID='111'   "
            MI29Last_amt = MI29Last_amt & " Update Ap_Rpt_Cashflow set Ap_Rpt_Cashflow.Last_amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Cashflow    where Ap_Rpt_Cashflow.rpt_ID=AP_Sum.rpt_ID "
            CNN.Execute(MI29Last_amt)
            '=========42	ເງິນສົດ ແລະ ທຽບເທົ່າເງິນສົດເພີ່ມຂຶ້ນ ຫຼື ຫຼຸດລົງສຸດທິໃນປີ====
            Dim MI42 As String = "delete AP_Sum "
            MI42 = MI42 & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select '42' ,0,0 from Ap_Rpt_Cashflow     where rpt_ID='42'  "
            MI42 = MI42 & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '111'  ,0 ,amt from Ap_Rpt_Cashflow     where rpt_ID='29'  "
            MI42 = MI42 & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '111'  ,0 ,amt from Ap_Rpt_Cashflow     where rpt_ID='34'  "
            MI42 = MI42 & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '111'  ,0 ,amt from Ap_Rpt_Cashflow     where rpt_ID='40'  "
            MI42 = MI42 & "  update AP_Sum set amt=(select sum(amt2)+sum(amt3) from AP_Sum where  (rpt_ID='42' or rpt_ID='111')  ) where rpt_ID='42'   "
            MI42 = MI42 & "  delete  AP_Sum  where rpt_ID='111'   "
            MI42 = MI42 & " Update Ap_Rpt_Cashflow set Ap_Rpt_Cashflow.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Cashflow    where Ap_Rpt_Cashflow.rpt_ID=AP_Sum.rpt_ID "
            CNN.Execute(MI42)
            '=========42	ເງິນສົດ ແລະ ທຽບເທົ່າເງິນສົດເພີ່ມຂຶ້ນ ຫຼື ຫຼຸດລົງສຸດທິໃນປີ====
            Dim MI42Last_amt As String = "delete AP_Sum "
            MI42Last_amt = MI42Last_amt & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select '42' ,0,0 from Ap_Rpt_Cashflow     where rpt_ID='42'  "
            MI42Last_amt = MI42Last_amt & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '111'  ,0 ,Last_amt from Ap_Rpt_Cashflow     where rpt_ID='29'  "
            MI42Last_amt = MI42Last_amt & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '111'  ,0 ,Last_amt from Ap_Rpt_Cashflow     where rpt_ID='34'  "
            MI42Last_amt = MI42Last_amt & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '111'  ,0 ,Last_amt from Ap_Rpt_Cashflow     where rpt_ID='40'  "
            MI42Last_amt = MI42Last_amt & "  update AP_Sum set amt=(select sum(amt2)+sum(amt3) from AP_Sum where  (rpt_ID='42' or rpt_ID='111')  ) where rpt_ID='42'   "
            MI42Last_amt = MI42Last_amt & "  delete  AP_Sum  where rpt_ID='111'   "
            MI42Last_amt = MI42Last_amt & " Update Ap_Rpt_Cashflow set Ap_Rpt_Cashflow.Last_amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Cashflow    where Ap_Rpt_Cashflow.rpt_ID=AP_Sum.rpt_ID "
            CNN.Execute(MI42Last_amt)

            Call LoadReportItem()
        End If
        'MdStartDate = d1
        'MdToDate = d2
    End Sub
    Private Sub CaCashflow_Detail()

    End Sub
    Private Sub CaCashflow()
        Click_Last()
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        CNN.Execute("DELETE  Ap_balance_6_col ")
        CNN.Execute("DELETE FROM Ap_balance_6 ")
        Dim S7 As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
        " select ac_code  ,  0 , 0  , 0  , 0  , sum(amount_Dr)as amt_dr , sum(amount_cr)as amt_cr  from gen_jn  WHERE curr='LAK' and  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code"
        CNN.Execute(S7)

        Dim Susd As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
  " select ac_code  ,  0 , 0  , 0  , 0  , sum(amount_Dr)* " & CDbl(txtRate.Text) & "  as amt_dr  , sum(amount_Cr)* " & CDbl(txtRate.Text) & "  as amt_Cr   from gen_jn  WHERE   curr='USD' and gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code"
        CNN.Execute(Susd)

        Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr      ) " & _
        " select  ac_code  ,  0 , 0   , sum(amount_Dr) as amt_dr , sum(amount_Cr) as amt_cr  , 0 , 0   from gen_jn  WHERE  curr='LAK' and gen_jn.date_work    BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr      ) " & _
    " select  ac_code  ,  0 , 0   ,  sum(amount_Dr)* " & CDbl(txtRate.Text) & "  as amt_dr ,  sum(amount_Cr)* " & CDbl(txtRate.Text) & "  as amt_cr , 0 , 0   from gen_jn  WHERE   curr='USD' and  gen_jn.date_work BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        '==============
        CNN.Execute("INSERT INTO Ap_balance_6 (  ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
        " select ac_code  , sum(amount_Dr) as amt_dr , sum(amount_Cr) as amt_cr   ,  0 , 0  , 0 , 0  from Open_jn WHERE   curr='LAK' and     date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        CNN.Execute("INSERT INTO Ap_balance_6 (  ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
    " select ac_code  , sum(amount_Dr)* " & CDbl(txtRate.Text) & "  as amt_dr ,  sum(amount_Cr)* " & CDbl(txtRate.Text) & "  as amt_Cr ,  0 , 0  , 0 , 0  from Open_jn WHERE   curr='USD' and     date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        '============
        CNN.Execute("UPDATE Ap_balance_6 set Ac_Code = left(Ac_Code,7) ")


        CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr  ) " & _
        " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr) as open_amt_cr , sum(Amt_Last_M_Dr) as Amt_Last_M_Dr , sum(Amt_Last_M_Cr) as Amt_Last_M_Cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")



        CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + Amt_Last_M_Dr + amt_dr) - (open_amt_cr + Amt_Last_M_Cr + amt_cr) where (open_amt_dr + Amt_Last_M_Dr + amt_dr) >= (open_amt_cr + Amt_Last_M_Cr + amt_cr) ")
        CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr +  Amt_Last_M_Cr + amt_cr) - (open_amt_dr + Amt_Last_M_Dr + amt_dr) where (open_amt_cr + Amt_Last_M_Cr + amt_cr) >= (open_amt_dr + Amt_Last_M_Dr + amt_dr) ")
        CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")

        CNN.Execute("update  Ap_Rpt_Cashflow_Item set Open_amt_dr  = 0 , OPen_amt_cr  = 0,Last_amt_dr  = 0 , Last_amt_cr  = 0 ,amt_dr  =  0 , amt_cr  = 0,  Rem_amt_dr  =  0 , Rem_amt_cr  = 0  , Amt=0  ")
        CNN.Execute("update Ap_Rpt_Cashflow set Last_Amt = 0 , Amt  =  0    ")
        CNN.Execute("DELETE FROM Ap_Rpt_Cashflow_Detail ")

        CNN.Execute("Update Ap_Rpt_Cashflow_Item set Open_amt_dr= Ap_balance_6_col.open_amt_dr , Open_amt_cr= Ap_balance_6_col.open_amt_cr , Last_amt_dr = Ap_balance_6_col.amt_Last_M_dr , Last_amt_cr = Ap_balance_6_col.amt_Last_M_cr , amt_dr = Ap_balance_6_col.amt_dr , amt_cr = Ap_balance_6_col.amt_cr , Rem_amt_dr = Ap_balance_6_col.Rem_dr , Rem_amt_cr = Ap_balance_6_col.Rem_cr  from Ap_Rpt_Cashflow_Item , Ap_balance_6_col " & _
        " where Ap_Rpt_Cashflow_Item.ac_code = Ap_balance_6_col.ac_code And  Rpt_Type = 'Dr-Cr'")
        CNN.Execute("Update Ap_Rpt_Cashflow_Item set Open_amt_dr= Ap_balance_6_col.open_amt_dr , Open_amt_cr= Ap_balance_6_col.open_amt_cr , Last_amt_dr = Ap_balance_6_col.amt_Last_M_dr , Last_amt_cr = Ap_balance_6_col.amt_Last_M_cr , amt_dr = Ap_balance_6_col.amt_dr , amt_cr = Ap_balance_6_col.amt_cr , Rem_amt_dr = Ap_balance_6_col.Rem_dr , Rem_amt_cr = Ap_balance_6_col.Rem_cr  from Ap_Rpt_Cashflow_Item , Ap_balance_6_col " & _
        "where Ap_Rpt_Cashflow_Item.ac_code = Ap_balance_6_col.ac_code And  Rpt_Type = 'Cr-Dr'")

        '        CNN.Execute("Update Ap_Rpt_Cashflow_Item set Open_amt_dr= Ap_balance_6_col.open_amt_dr , Open_amt_cr= Ap_balance_6_col.open_amt_cr , Last_amt_dr = Ap_balance_6_col.amt_Last_M_dr , Last_amt_cr = Ap_balance_6_col.amt_Last_M_cr , amt_dr = Ap_balance_6_col.amt_dr , amt_cr = Ap_balance_6_col.amt_cr , Rem_amt_dr = Ap_balance_6_col.Rem_dr , Rem_amt_cr = Ap_balance_6_col.Rem_cr  from Ap_Rpt_Cashflow_Item , Ap_balance_6_col " & _
        '" where Ap_Rpt_Cashflow_Item.ac_code = Ap_balance_6_col.ac_code And  Rpt_Type = 'Dr'")
        '        CNN.Execute("Update Ap_Rpt_Cashflow_Item set Open_amt_dr= Ap_balance_6_col.open_amt_dr , Open_amt_cr= Ap_balance_6_col.open_amt_cr , Last_amt_dr = Ap_balance_6_col.amt_Last_M_dr , Last_amt_cr = Ap_balance_6_col.amt_Last_M_cr , amt_dr = Ap_balance_6_col.amt_dr , amt_cr = Ap_balance_6_col.amt_cr , Rem_amt_dr = Ap_balance_6_col.Rem_dr , Rem_amt_cr = Ap_balance_6_col.Rem_cr  from Ap_Rpt_Cashflow_Item , Ap_balance_6_col " & _
        '        "where Ap_Rpt_Cashflow_Item.ac_code = Ap_balance_6_col.ac_code And  Rpt_Type = 'Cr'")
  


        CNN.Execute("Insert into Ap_Rpt_Cashflow_Detail (  Rpt_Id , Ac_Code , Ac_Name , Amt )" & _
        " select   Rpt_Id , Ac_Code , Ac_Name , Amt from Ap_Rpt_Cashflow_Item  ")
        CNN.Execute("update Ap_Rpt_Cashflow_Item set Open_Amt_dr=0 , Open_Amt_Cr=0 where Select_OPen_Amt=0")
        CNN.Execute("update Ap_Rpt_Cashflow_Item set Last_Amt_dr=0 , Last_Amt_Cr=0 where Select_Last_Amt=0")
        CNN.Execute("update Ap_Rpt_Cashflow_Item set Amt_dr=0 , Amt_Cr=0 where Select_Amt=0 ")
        CNN.Execute("update Ap_Rpt_Cashflow_Item set Rem_Amt_dr=0 , Rem_Amt_Cr=0 where Select_Rem_Amt=0")
        'CNN.Execute("update Ap_Rpt_Cashflow_Item set Amt=((Open_Amt_dr+Last_Amt_dr +Amt_dr+Rem_Amt_dr)-(Open_Amt_cr+Last_Amt_cr +Amt_cr+Rem_Amt_cr)) where Rpt_Type = 'Dr-Cr' ")
        'CNN.Execute("update Ap_Rpt_Cashflow_Item set Amt=((Open_Amt_cr+Last_Amt_cr +Amt_cr+Rem_Amt_cr)-(Open_Amt_dr+Last_Amt_dr +Amt_dr+Rem_Amt_dr)) where Rpt_Type = 'Cr-Dr'")

        CNN.Execute("update Ap_Rpt_Cashflow_Item set Amt=((Open_Amt_dr+Last_Amt_dr +Amt_dr+Rem_Amt_dr)-(Open_Amt_cr+Last_Amt_cr +Amt_cr+Rem_Amt_cr)) where Rpt_Type = 'Dr-Cr' ")
        CNN.Execute("update Ap_Rpt_Cashflow_Item set Amt=((Open_Amt_cr+Last_Amt_cr +Amt_cr+Rem_Amt_cr)-(Open_Amt_dr+Last_Amt_dr +Amt_dr+Rem_Amt_dr)) where Rpt_Type = 'Cr-Dr'")


        'CNN.Execute("update Ap_Rpt_Cashflow_Item set Amt=(Amt_cr) where Rpt_Type = 'Cr' ")
        'CNN.Execute("update Ap_Rpt_Cashflow_Item set Amt=(Amt_dr) where Rpt_Type = 'Dr'")


        CNN.Execute("delete Ap_Rpt_BLS_Stock ")
        CNN.Execute(" insert into Ap_Rpt_BLS_Stock ( Rpt_ID , Amt)" & _
        "select Rpt_ID , sum(Amt) As Amt   from Ap_Rpt_Cashflow_Item     group by Rpt_ID")
        CNN.Execute("Update Ap_Rpt_Cashflow set Amt = Ap_Rpt_BLS_Stock.Amt from Ap_Rpt_Cashflow ,Ap_Rpt_BLS_Stock where  Ap_Rpt_Cashflow.Rpt_ID=Ap_Rpt_BLS_Stock.Rpt_ID")
        CNN.Execute("update Ap_Rpt_Cashflow_Detail set  Rpt_Name=Ap_Rpt_Cashflow.Description from   Ap_Rpt_Cashflow_Detail , Ap_Rpt_Cashflow  where Ap_Rpt_Cashflow_Detail.Rpt_Id = Ap_Rpt_Cashflow.Rpt_Id")

        CNN.Execute("Delete Ap_Rpt_Cashflow2 ")
        CNN.Execute("Insert into Ap_Rpt_Cashflow2 (Rpt_Id ) Select Rpt_Id from Ap_Rpt_Cashflow ")
        CNN.Execute("Update Ap_Rpt_Cashflow2 set Amt = Ap_Rpt_Cashflow.Amt from Ap_Rpt_Cashflow2 , Ap_Rpt_Cashflow where Ap_Rpt_Cashflow2.rpt_Id = Ap_Rpt_Cashflow.Rpt_Id  ")
        '2============================
        'If Month(MdStartDate) > 1 Then
        If RM.Checked = True Then
            'MsgBox("1")
            MdStartDate2 = DateAdd(DateInterval.Month, -1, MdStartDate)
            MdToDate2 = DateAdd(DateInterval.Day, -1, MdStartDate)
            Call ChaneLastMonth()
            'MsgBox("2")
        ElseIf RP.Checked = True Then
            MdStartDate2 = DateAdd(DateInterval.Month, -3, MdStartDate)
            MdToDate2 = DateAdd(DateInterval.Day, -1, MdStartDate)
            If Period.SelectedIndex > 0 Then
                MdStartDate2 = DateAdd(DateInterval.Month, -3, MdStartDate)
                MdToDate2 = DateAdd(DateInterval.Day, -1, MdStartDate)
                'MsgBox(MdStartDate2 & "  =  " & MdToDate2)
            End If
            Call ChaneLastMonth()
        ElseIf RT.Checked = True Then
            MdStartDate2 = DateAdd(DateInterval.Month, -6, MdStartDate)
            MdToDate2 = DateAdd(DateInterval.Day, -1, MdStartDate)
            If Period.SelectedIndex > 0 Then
                MdStartDate2 = DateAdd(DateInterval.Month, -6, MdStartDate)
                MdToDate2 = DateAdd(DateInterval.Day, -1, MdStartDate)
                'MsgBox(MdStartDate2 & "  =  " & MdToDate2)
            End If
            Call ChaneLastMonth()

        ElseIf RY.Checked = True Then
            MdStartDate2 = Format(MdStartDate, "yyyy") - 1 & "-1-1"
            MdToDate2 = Format(MdToDate, "yyyy") - 1 & "-12-31"
            Call ChaneLastMonth()
        End If

        'End If


        CNN.Execute("Update Ap_Rpt_Cashflow2 set Last_Amt = Ap_Rpt_Cashflow.Amt from Ap_Rpt_Cashflow2 , Ap_Rpt_Cashflow where Ap_Rpt_Cashflow2.rpt_Id = Ap_Rpt_Cashflow.Rpt_ID ")
        CNN.Execute("Update Ap_Rpt_Cashflow set Amt = Ap_Rpt_Cashflow2.Amt , Last_Amt = Ap_Rpt_Cashflow2.Last_Amt from Ap_Rpt_Cashflow , Ap_Rpt_Cashflow2 where Ap_Rpt_Cashflow.rpt_Id = Ap_Rpt_Cashflow2.Rpt_ID ")





        Update_Sum()

    End Sub
    Private Sub ChaneLastMonth()
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        CNN.Execute("DELETE  Ap_balance_6_col ")
        CNN.Execute("DELETE FROM Ap_balance_6 ")
        Dim S7 As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
        " select ac_code  ,  0 , 0  , 0  , 0  , sum(amount_dr)as amt_dr , sum(amount_Cr)as amt_cr  from gen_jn  WHERE curr='LAK' and  gen_jn.date_work   BETWEEN '" & Format(MdStartDate2, "yyyy-MM-dd") & "' AND '" & Format(MdToDate2, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code"
        CNN.Execute(S7)
        Dim S7p As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
    " select ac_code  ,  0 , 0  , 0  , 0  , sum(amount_Dr)* " & CDbl(txtRate_Last.Text) & "  as amt_dr , sum(amount_Cr)* " & CDbl(txtRate_Last.Text) & "  as amt_Cr  from gen_jn  WHERE  curr='USD' and gen_jn.date_work   BETWEEN '" & Format(MdStartDate2, "yyyy-MM-dd") & "' AND '" & Format(MdToDate2, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code"
        CNN.Execute(S7p)

        Dim S As Date = MdStartDate2 : S = DateAdd("d", CDbl(-1), MdStartDate2)
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr      ) " & _
        " select  ac_code  ,  0 , 0   , sum(amount_dr) as amt_dr , sum(amount_Cr) as amt_cr  , 0 , 0   from gen_jn  WHERE curr='LAK' and   gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate2, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr      ) " & _
" select  ac_code  ,  0 , 0   ,  sum(amount_Dr)* " & CDbl(txtRate_Last.Text) & "  as amt_dr ,  sum(amount_Cr)* " & CDbl(txtRate_Last.Text) & "  as amt_cr  , 0 , 0   from gen_jn  WHERE curr='USD' and   gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate2, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

        CNN.Execute("INSERT INTO Ap_balance_6 (  ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
        " select ac_code  , sum(amount_Dr) as amt_dr , sum(amount_Cr) as amt_cr   ,  0 , 0  , 0 , 0  from Open_jn WHERE  curr='LAK' and   date_work='" & "1-1-" & Format(MdStartDate2, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        CNN.Execute("INSERT INTO Ap_balance_6 (  ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
   " select ac_code  ,  sum(amount_Dr)* " & CDbl(txtRate_Last.Text) & " as amt_dr ,  sum(amount_Cr)* " & CDbl(txtRate_Last.Text) & "  as amt_cr   ,  0 , 0  , 0 , 0  from Open_jn WHERE  curr='USD' and   date_work='" & "1-1-" & Format(MdStartDate2, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

        CNN.Execute("UPDATE Ap_balance_6 set Ac_Code = left(Ac_Code,7) ")

        CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr  ) " & _
        " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr) as open_amt_cr , sum(Amt_Last_M_Dr) as Amt_Last_M_Dr , sum(Amt_Last_M_Cr) as Amt_Last_M_Cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")

        CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + Amt_Last_M_Dr + amt_dr) - (open_amt_cr + Amt_Last_M_Cr + amt_cr) where (open_amt_dr + Amt_Last_M_Dr + amt_dr) >= (open_amt_cr + Amt_Last_M_Cr + amt_cr) ")
        CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr +  Amt_Last_M_Cr + amt_cr) - (open_amt_dr + Amt_Last_M_Dr + amt_dr) where (open_amt_cr + Amt_Last_M_Cr + amt_cr) >= (open_amt_dr + Amt_Last_M_Dr + amt_dr) ")
        CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        CNN.Execute("update  Ap_Rpt_Cashflow_Item set Open_amt_dr  = 0 , OPen_amt_cr  = 0,Last_amt_dr  = 0 , Last_amt_cr  = 0 ,amt_dr  =  0 , amt_cr  = 0,  Rem_amt_dr  =  0 , Rem_amt_cr  = 0  , Amt=0  ")
        CNN.Execute("update Ap_Rpt_Cashflow set Last_Amt = 0 , Amt  =  0  ")
        CNN.Execute("DELETE FROM Ap_Rpt_Cashflow_Detail ")

        CNN.Execute("Update Ap_Rpt_Cashflow_Item set Open_amt_dr= Ap_balance_6_col.open_amt_dr , Open_amt_cr= Ap_balance_6_col.open_amt_cr , Last_amt_dr = Ap_balance_6_col.amt_Last_M_dr , Last_amt_cr = Ap_balance_6_col.amt_Last_M_cr , amt_dr = Ap_balance_6_col.amt_dr , amt_cr = Ap_balance_6_col.amt_cr , Rem_amt_dr = Ap_balance_6_col.Rem_dr , Rem_amt_cr = Ap_balance_6_col.Rem_cr  from Ap_Rpt_Cashflow_Item , Ap_balance_6_col " & _
        " where Ap_Rpt_Cashflow_Item.ac_code = Ap_balance_6_col.ac_code And  Rpt_Type = 'Dr-Cr'")
        CNN.Execute("Update Ap_Rpt_Cashflow_Item set Open_amt_dr= Ap_balance_6_col.open_amt_dr , Open_amt_cr= Ap_balance_6_col.open_amt_cr , Last_amt_dr = Ap_balance_6_col.amt_Last_M_dr , Last_amt_cr = Ap_balance_6_col.amt_Last_M_cr , amt_dr = Ap_balance_6_col.amt_dr , amt_cr = Ap_balance_6_col.amt_cr , Rem_amt_dr = Ap_balance_6_col.Rem_dr , Rem_amt_cr = Ap_balance_6_col.Rem_cr  from Ap_Rpt_Cashflow_Item , Ap_balance_6_col " & _
        "where Ap_Rpt_Cashflow_Item.ac_code = Ap_balance_6_col.ac_code And  Rpt_Type = 'Cr-Dr'")
        CNN.Execute("Insert into Ap_Rpt_Cashflow_Detail (  Rpt_Id , Ac_Code , Ac_Name , Amt )" & _
        " select   Rpt_Id , Ac_Code , Ac_Name , Amt from Ap_Rpt_Cashflow_Item  ")
        CNN.Execute("update Ap_Rpt_Cashflow_Item set Open_Amt_dr=0 , Open_Amt_Cr=0 where Select_OPen_Amt=0")
        CNN.Execute("update Ap_Rpt_Cashflow_Item set Last_Amt_dr=0 , Last_Amt_Cr=0 where Select_Last_Amt=0")
        CNN.Execute("update Ap_Rpt_Cashflow_Item set Amt_dr=0 , Amt_Cr=0 where Select_Amt=0 ")
        CNN.Execute("update Ap_Rpt_Cashflow_Item set Rem_Amt_dr=0 , Rem_Amt_Cr=0 where Select_Rem_Amt=0")
        CNN.Execute("update Ap_Rpt_Cashflow_Item set Amt=((Open_Amt_dr+Last_Amt_dr +Amt_dr+Rem_Amt_dr)-(Open_Amt_cr+Last_Amt_cr +Amt_cr+Rem_Amt_cr)) where Rpt_Type = 'Dr-Cr' ")
        CNN.Execute("update Ap_Rpt_Cashflow_Item set Amt=((Open_Amt_cr+Last_Amt_cr +Amt_cr+Rem_Amt_cr)-(Open_Amt_dr+Last_Amt_dr +Amt_dr+Rem_Amt_dr)) where Rpt_Type = 'Cr-Dr'")

        'CNN.Execute("update Ap_Rpt_Cashflow_Item set Amt=((Open_Amt_cr+Last_Amt_cr +Amt_cr+Rem_Amt_cr)) where Rpt_Type = 'Cr' ")
        'CNN.Execute("update Ap_Rpt_Cashflow_Item set Amt=((Open_Amt_dr+Last_Amt_dr +Amt_dr+Rem_Amt_dr)) where Rpt_Type = 'Dr'")


        CNN.Execute("delete Ap_Rpt_BLS_Stock ")
        CNN.Execute(" insert into Ap_Rpt_BLS_Stock ( Rpt_ID , Amt)" & _
        "select Rpt_ID , sum(Amt) As Amt   from Ap_Rpt_Cashflow_Item     group by Rpt_ID")
        CNN.Execute("Update Ap_Rpt_Cashflow set Amt = Ap_Rpt_BLS_Stock.Amt from Ap_Rpt_Cashflow ,Ap_Rpt_BLS_Stock where  Ap_Rpt_Cashflow.Rpt_ID=Ap_Rpt_BLS_Stock.Rpt_ID")
        CNN.Execute("update Ap_Rpt_Cashflow_Detail set  Rpt_Name=Ap_Rpt_Cashflow.Description from   Ap_Rpt_Cashflow_Detail , Ap_Rpt_Cashflow  where Ap_Rpt_Cashflow_Detail.Rpt_Id = Ap_Rpt_Cashflow.Rpt_Id")


        ' ''Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        ' ''CNN.Execute("DELETE  Ap_balance_6_col ")
        ' ''CNN.Execute("DELETE FROM Ap_balance_6 ")
        ' ''Dim S7 As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
        ' ''" select ac_code  ,  0 , 0  , 0  , 0  , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate2, "yyyy-MM-dd") & "' AND '" & Format(MdToDate2, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code"
        ' ''CNN.Execute(S7)

        ' ''Dim S As Date = MdStartDate2 : S = DateAdd("d", CDbl(-1), MdStartDate2)
        ' ''CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr      ) " & _
        ' ''" select  ac_code  ,  0 , 0   , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  , 0 , 0   from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate2, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        ' ''CNN.Execute("INSERT INTO Ap_balance_6 (  ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
        ' ''" select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr   ,  0 , 0  , 0 , 0  from Open_jn WHERE    date_work='" & "1-1-" & Format(MdStartDate2, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        ' ''CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr  ) " & _
        ' ''" select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr) as open_amt_cr , sum(Amt_Last_M_Dr) as Amt_Last_M_Dr , sum(Amt_Last_M_Cr) as Amt_Last_M_Cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")
        ' ''CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        ' ''CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        ' ''CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + Amt_Last_M_Dr + amt_dr) - (open_amt_cr + Amt_Last_M_Cr + amt_cr) where (open_amt_dr + Amt_Last_M_Dr + amt_dr) >= (open_amt_cr + Amt_Last_M_Cr + amt_cr) ")
        ' ''CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr +  Amt_Last_M_Cr + amt_cr) - (open_amt_dr + Amt_Last_M_Dr + amt_dr) where (open_amt_cr + Amt_Last_M_Cr + amt_cr) >= (open_amt_dr + Amt_Last_M_Dr + amt_dr) ")
        ' ''CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        ' ''CNN.Execute("update  Ap_Rpt_Cashflow_Item set Open_amt_dr  = 0 , OPen_amt_cr  = 0,Last_amt_dr  = 0 , Last_amt_cr  = 0 ,amt_dr  =  0 , amt_cr  = 0,  Rem_amt_dr  =  0 , Rem_amt_cr  = 0  , Amt=0  ")
        ' ''CNN.Execute("update Ap_Rpt_Cashflow set Last_Amt = 0 , Amt  =  0  ")
        ' ''CNN.Execute("DELETE FROM Ap_Rpt_Cashflow_Detail ")

        ' ''CNN.Execute("Update Ap_Rpt_Cashflow_Item set Open_amt_dr= Ap_balance_6_col.open_amt_dr , Open_amt_cr= Ap_balance_6_col.open_amt_cr , Last_amt_dr = Ap_balance_6_col.amt_Last_M_dr , Last_amt_cr = Ap_balance_6_col.amt_Last_M_cr , amt_dr = Ap_balance_6_col.amt_dr , amt_cr = Ap_balance_6_col.amt_cr , Rem_amt_dr = Ap_balance_6_col.Rem_dr , Rem_amt_cr = Ap_balance_6_col.Rem_cr  from Ap_Rpt_Cashflow_Item , Ap_balance_6_col " & _
        ' ''" where Ap_Rpt_Cashflow_Item.ac_code = Ap_balance_6_col.ac_code And  Rpt_Type = 'Dr-Cr'")
        ' ''CNN.Execute("Update Ap_Rpt_Cashflow_Item set Open_amt_dr= Ap_balance_6_col.open_amt_dr , Open_amt_cr= Ap_balance_6_col.open_amt_cr , Last_amt_dr = Ap_balance_6_col.amt_Last_M_dr , Last_amt_cr = Ap_balance_6_col.amt_Last_M_cr , amt_dr = Ap_balance_6_col.amt_dr , amt_cr = Ap_balance_6_col.amt_cr , Rem_amt_dr = Ap_balance_6_col.Rem_dr , Rem_amt_cr = Ap_balance_6_col.Rem_cr  from Ap_Rpt_Cashflow_Item , Ap_balance_6_col " & _
        ' ''"where Ap_Rpt_Cashflow_Item.ac_code = Ap_balance_6_col.ac_code And  Rpt_Type = 'Cr-Dr'")
        ' ''CNN.Execute("Insert into Ap_Rpt_Cashflow_Detail (  Rpt_Id , Ac_Code , Ac_Name , Amt )" & _
        ' ''" select   Rpt_Id , Ac_Code , Ac_Name , Amt from Ap_Rpt_Cashflow_Item  ")
        ' ''CNN.Execute("update Ap_Rpt_Cashflow_Item set Open_Amt_dr=0 , Open_Amt_Cr=0 where Select_OPen_Amt=0")
        ' ''CNN.Execute("update Ap_Rpt_Cashflow_Item set Last_Amt_dr=0 , Last_Amt_Cr=0 where Select_Last_Amt=0")
        ' ''CNN.Execute("update Ap_Rpt_Cashflow_Item set Amt_dr=0 , Amt_Cr=0 where Select_Amt=0 ")
        ' ''CNN.Execute("update Ap_Rpt_Cashflow_Item set Rem_Amt_dr=0 , Rem_Amt_Cr=0 where Select_Rem_Amt=0")
        ' ''CNN.Execute("update Ap_Rpt_Cashflow_Item set Amt=((Open_Amt_dr+Last_Amt_dr +Amt_dr+Rem_Amt_dr)-(Open_Amt_cr+Last_Amt_cr +Amt_cr+Rem_Amt_cr)) where Rpt_Type = 'Dr-Cr' ")
        ' ''CNN.Execute("update Ap_Rpt_Cashflow_Item set Amt=((Open_Amt_cr+Last_Amt_cr +Amt_cr+Rem_Amt_cr)-(Open_Amt_dr+Last_Amt_dr +Amt_dr+Rem_Amt_dr)) where Rpt_Type = 'Cr-Dr'")
        ' ''CNN.Execute("delete Ap_Rpt_BLS_Stock ")
        ' ''CNN.Execute(" insert into Ap_Rpt_BLS_Stock ( Rpt_ID , Amt)" & _
        ' ''"select Rpt_ID , sum(Amt) As Amt   from Ap_Rpt_Cashflow_Item     group by Rpt_ID")
        ' ''CNN.Execute("Update Ap_Rpt_Cashflow set Amt = Ap_Rpt_BLS_Stock.Amt from Ap_Rpt_Cashflow ,Ap_Rpt_BLS_Stock where  Ap_Rpt_Cashflow.Rpt_ID=Ap_Rpt_BLS_Stock.Rpt_ID")
        ' ''CNN.Execute("update Ap_Rpt_Cashflow_Detail set  Rpt_Name=Ap_Rpt_Cashflow.Description from   Ap_Rpt_Cashflow_Detail , Ap_Rpt_Cashflow  where Ap_Rpt_Cashflow_Detail.Rpt_Id = Ap_Rpt_Cashflow.Rpt_Id")

        'CNN.Execute(" Update Caculate_Rpt set  CLT_Amt  = CLT_Str ,  CLT_Last_Amt  = CLT_Str where CLT_Str = '+' or CLT_Str = '-' or CLT_Str = '*' or CLT_Str = '+' or CLT_Str = '/' or CLT_Str = '(' or CLT_Str=')' ")

        'CNN.Execute("delete Caculate_Lock")
        'CNN.Execute("delete Caculate_Start")
        'CNN.Execute(" Insert Into Caculate_Start (Rpt_Id , Rpt_Type , clt_Str , clt_Amt , Clt_Last_Amt ) select Rpt_Id , Rpt_Type , clt_Str , clt_Amt , Clt_Last_Amt from Caculate_Rpt where Rpt_Type = 'CAF'  Order by  Rpt_id ,cnt asc ")
        'CNN.Execute("update Caculate_Start set lck =0")
        'CNN.Execute("Insert into Caculate_Lock (cnt_Mt)  SELECT  (SELECT     TOP 1 cnt FROM Caculate_Start AS B WHERE(Rpt_Id = A.Rpt_Id   ) ORDER BY cnt desc) AS cnt FROM Caculate_Start  AS A  GROUP BY Rpt_Id ORDER BY Rpt_Id")
        'CNN.Execute("update  Caculate_Start set lck=1 from Caculate_Start ,Caculate_Lock  where Caculate_Start.cnt=Caculate_Lock.cnt_MT")
        'CNN.Execute("  Update Caculate_Start set Caculate_Start.Amt = Ap_Rpt_Cashflow.Amt , Caculate_Start.Last_Amt = Ap_Rpt_Cashflow.Last_Amt   from Caculate_Start , Ap_Rpt_Cashflow  where  Caculate_Start.CLT_Str  = Ap_Rpt_Cashflow.Rpt_Id  ")
        'CNN.Execute("Update Caculate_Start set lck_Amt=0")
        'CNN.Execute("Update Caculate_Start set lck_Amt=1 where CLT_Str <> '+' And CLT_Str <> '-' And CLT_Str <> '*' And CLT_Str <> '+' And CLT_Str <> '/' And CLT_Str <> '(' And CLT_Str<>')'")

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
        LoadSqlData("select * from Ap_Rpt_Cashflow_Item where  Rpt_Type = 'In'", RSCIn_M)
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
                CNN.Execute("Insert into Ap_Rpt_Cashflow_Detail (Ac_Code_Parent , Rpt_Id , Ac_Code , Ac_Name  , Last_Amt_Dr , Last_Amt_Cr , Rpt_Type) values (  '" & CStr((RSCIn_M.Fields("Ac_Code").Value)) & "' , '" & CStr((RSCIn_M.Fields("Rpt_Id").Value)) & "' , '" & CStr((.Fields("Ac_Code").Value)) & "' , N'" & CStr((.Fields("Ac_Name").Value)) & "'  , " & CDbl((.Fields("Rem_dr").Value)) & " , " & CDbl((.Fields("Rem_cr").Value)) & " , 'In' )")
                CNN.Execute("update  Ap_Rpt_Cashflow_Item set Amt_Dr  =  Amt_Dr+" & CDbl((.Fields("Rem_dr").Value)) & " , Amt_Cr  = Amt_Cr+" & CDbl((.Fields("Rem_cr").Value)) & "   where ac_code = '" & (RSCIn_M.Fields("ac_code").Value) & "' And  Rpt_Type = 'In' ")
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
        LoadSqlData("select Rpt_ID, sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr  from Ap_Rpt_Cashflow_Item  where  Rpt_Type = 'In' group by Rpt_ID ", RSC)
        With RSC
            Do Until .EOF = True
                CNN.Execute("Update Ap_Rpt_Cashflow set " & _
                            " Last_Amt ='" & CDbl(CDbl((.Fields("Amt_dr").Value)) - CDbl((.Fields("Amt_cr").Value))) & "' " & _
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
        'If RSCIn_M.State = ConnectionState.Open Then RSCIn_M.Close()
    End Sub
    Private Sub SelectOutLast()

        LoadSqlData("select * from Ap_Rpt_Cashflow_Item where  Rpt_Type = 'Out' ", RSCIn_M)
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
                CNN.Execute("Insert into Ap_Rpt_Cashflow_Detail (Ac_Code_Parent , Rpt_Id , Ac_Code , Ac_Name  ,  Last_Amt_Dr , Last_Amt_Cr, Amt_Dr , Amt_Cr , Rpt_Type ) values (  '" & CStr((RSCIn_M.Fields("Ac_Code").Value)) & "' , '" & CStr((RSCIn_M.Fields("Rpt_Id").Value)) & "' , '" & CStr((.Fields("Ac_Code").Value)) & "' , N'" & CStr((.Fields("Ac_Name").Value)) & "'  ,   " & CDbl((.Fields("Open_Amt_dr").Value)) & " , " & CDbl((.Fields("Open_Amt_cr").Value)) & " , " & CDbl((.Fields("Rem_dr").Value)) & " , " & CDbl((.Fields("Rem_cr").Value)) & " , 'Out' )")
                CNN.Execute("update  Ap_Rpt_Cashflow_Item set Last_Amt_Dr  =  Last_Amt_Dr+" & CDbl((.Fields("Open_Amt_dr").Value)) & " , Last_Amt_Cr  = Last_Amt_Cr+" & CDbl((.Fields("Open_Amt_cr").Value)) & "  , Amt_Dr  =  Amt_Dr+" & CDbl((.Fields("Rem_dr").Value)) & " , Amt_Cr  = Amt_Cr+" & CDbl((.Fields("Rem_cr").Value)) & "   where ac_code = '" & (RSCIn_M.Fields("ac_code").Value) & "' And  Rpt_Type = 'Out' ")

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
        LoadSqlData("select Rpt_ID, sum(Last_Amt_dr) As Last_Amt_Dr , sum(Last_Amt_cr) As Last_Amt_cr , sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr  from Ap_Rpt_Cashflow_Item  where  Rpt_Type = 'Out' group by Rpt_ID ", RSC)
        With RSC
            Do Until .EOF = True
                CNN.Execute("Update Ap_Rpt_Cashflow set " & _
                         " Amt ='" & CDbl(CDbl((.Fields("Amt_cr").Value)) - CDbl((.Fields("Amt_dr").Value))) & "' " & _
                           " ,Last_Amt ='" & CDbl(CDbl((.Fields("Last_Amt_cr").Value)) - CDbl((.Fields("Last_Amt_dr").Value))) & "' " & _
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

        CNN.Execute(" Update Caculate_Rpt set  CLT_Amt  = CLT_Str ,  CLT_Last_Amt  = CLT_Str where CLT_Str = '+' or CLT_Str = '-' or CLT_Str = '*' or CLT_Str = '+' or CLT_Str = '/' or CLT_Str = '(' or CLT_Str=')' Or CLT_Str<>'Cast(('   Or CLT_Str<>')As Float)' ")

        CNN.Execute("delete Caculate_Lock")
        CNN.Execute("delete Caculate_Start")
        CNN.Execute(" Insert Into Caculate_Start (Rpt_Id , Rpt_Type , clt_Str , clt_Amt , Clt_Last_Amt ) select Rpt_Id , Rpt_Type , clt_Str , clt_Amt , Clt_Last_Amt from Caculate_Rpt where Rpt_Type = 'CAF'  Order by  Rpt_id ,cnt asc ")
        CNN.Execute("update Caculate_Start set lck =0")
        CNN.Execute("Insert into Caculate_Lock (cnt_Mt)  SELECT  (SELECT     TOP 1 cnt FROM Caculate_Start AS B WHERE(Rpt_Id = A.Rpt_Id   ) ORDER BY cnt desc) AS cnt FROM Caculate_Start  AS A  GROUP BY Rpt_Id ORDER BY Rpt_Id")
        CNN.Execute("update  Caculate_Start set lck=1 from Caculate_Start ,Caculate_Lock  where Caculate_Start.cnt=Caculate_Lock.cnt_MT")
        CNN.Execute("  Update Caculate_Start set Caculate_Start.Amt = Ap_Rpt_Cashflow.Amt , Caculate_Start.Last_Amt = Ap_Rpt_Cashflow.Last_Amt   from Caculate_Start , Ap_Rpt_Cashflow  where  Caculate_Start.CLT_Str  = Ap_Rpt_Cashflow.Rpt_Id  ")
        CNN.Execute("Update Caculate_Start set lck_Amt=0")
        CNN.Execute("Update Caculate_Start set lck_Amt=1 where CLT_Str <> '+' And CLT_Str <> '-' And CLT_Str <> '*' And CLT_Str <> '+' And CLT_Str <> '/' And CLT_Str <> '(' And CLT_Str<>')' And CLT_Str<>'Cast(('   And CLT_Str<>')As Float)' ")
        'MsgBox("44")




        Dim RSC1 As New ADODB.Recordset
        CLT_Str = ""
        CLT_Last_Str = ""
        With RSC1
            Call LoadSqlData("select *  from Caculate_Start where Rpt_Type = 'CAF'  Order by  Rpt_id ,cnt asc", RSC1)
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

                            CNN.Execute(" Update  Ap_Rpt_Cashflow set Amt = " & CLT_Str & " , Last_Amt = " & CLT_Last_Str & " where  Rpt_ID =   '" & (RSC1.Fields("Rpt_ID").Value.ToString) & "' ")
                            'If (RSC1.Fields("Rpt_ID").Value.ToString) = 20 Then
                            '    MsgBox("hh")
                            'End If
                        Else
                            Dim s As String = " Update  Ap_Rpt_Cashflow set Amt = " & CLT_Str & " , Last_Amt = " & CLT_Last_Str & " where  Rpt_ID =   '" & (RSC1.Fields("Rpt_ID").Value.ToString) & "' "
                            s = 0
                            MessageBox.Show("ສູດຄິດໄລ່ຂອງ " & (RSC1.Fields("Rpt_ID").Value.ToString) & " = " & CLT_Str & " ບໍ່ຖຶກຕ້ອງກະລຸນນາກວດສອບຄືນໃຫມ່")
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
            MdStartDate = Format(CDate("01/01/" & Year(yyt.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/06/" & Year(yyt.Value)), "dd-MM-yyyy")
            MdStartDate_MM = Format(CDate("01/07/" & Year(yyt.Value) - 1), "dd-MM-yyyy")
            MdToDate_MM = Format(CDate("31/12/" & Year(yyt.Value) - 1), "dd-MM-yyyy")
        Else
            MdStartDate = Format(CDate("01/07/" & Year(yyt.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/12/" & Year(yyt.Value)), "dd-MM-yyyy")
            MdStartDate_MM = Format(CDate("01/01/" & Year(yyt.Value)), "dd-MM-yyyy")
            MdToDate_MM = Format(CDate("30/06/" & Year(yyt.Value)), "dd-MM-yyyy")
        End If

        Lb.Text = Ct.Text & " " & yyt.Text
        'L5.Text = MdStartDate & " => " & MdToDate

        L5.Text = MdStartDate & " => " & MdToDate


        'L5.Text = MdStartDate & " => " & MdToDate
    End Sub
    Private Sub LoadDay()
        MdStartDate = Format(CDate(Ds.Value), "dd-MM-yyyy")
        MdToDate = Format(CDate(Dt.Value), "dd-MM-yyyy")
        Lb.Text = "ແຕ່ວັນທີ " & MdStartDate & " ຫາວັນທີ " & MdToDate
        L5.Text = MdStartDate & " => " & MdToDate
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳວັນທີ"
    End Sub
    Dim INNM, LastNM As String
    Private Sub LoadMonth()
         
        '---------------------------------
        If FmMain.MnLaoLang.Checked = True Then
            If DMonth.Text = "ມັງກອນ" Then
                MdStartDate = Format(CDate("01/01/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/01/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ມັງກອນ"
                DMonth.SelectedIndex = 0
                dpMonthPrev.Value = DateAdd("m", 1, MdStartDate)
                INNM = "ທ້າຍເດືອນ " & Format(CDate(MdToDate), "MM/yyyy")

            ElseIf DMonth.Text = "ກຸມພາ" Then
                Dim Day As String
                Dim MM As Date
                Dim Fromm As Date
                MdStartDate = Format(CDate("01/02/" & Year(Myy.Value)), "dd-MM-yyyy")
                Fromm = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
                MM = Format(CDate("01/3/" & Year(MdStartDate)), "dd-MM-yyyy")
                Day = DateDiff(DateInterval.Day, Fromm, MM)
                MdToDate = Format(CDate(Day & "/02" & "/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ກຸມພາ"
                DMonth.SelectedIndex = 1
                Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
                dpMonthPrev.Value = DateAdd("m", -1, MdStartDate)

                MdStartDate_MM = Format(CDate("01/01/" & Year(MdStartDate)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/01/" & Year(MdStartDate)), "dd-MM-yyyy")

            ElseIf DMonth.Text = "ມີນາ" Then
                MdStartDate = Format(CDate("01/03/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/03/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ມີນາ"
                DMonth.SelectedIndex = 2
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                Dim Day As String
                Dim MM As Date
                Dim Fromm As Date
                Fromm = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
                MM = Format(CDate("01/3/" & Year(MdStartDate)), "dd-MM-yyyy")
                Day = DateDiff(DateInterval.Day, Fromm, MM)
                MdStartDate_MM = Format(CDate("01/02/" & Year(dpMonthPrev.Value)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate(Day & "/02" & "/" & Year(dpMonthPrev.Value)), "dd-MM-yyyy")

            ElseIf DMonth.Text = "ເມສາ" Then
                MdStartDate = Format(CDate("01/04/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/04/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ເມສາ"
                DMonth.SelectedIndex = 3
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)

                MdStartDate_MM = Format(CDate("01/03/" & Year(MdStartDate)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/03/" & Year(MdStartDate)), "dd-MM-yyyy")


            ElseIf DMonth.Text = "ພຶດສະພາ" Then
                MdStartDate = Format(CDate("01/05/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/05/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ພຶດສະພາ"
                DMonth.SelectedIndex = 4
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                MdStartDate_MM = Format(CDate("01/04/" & Year(MdStartDate)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("30/04/" & Year(MdStartDate)), "dd-MM-yyyy")

            ElseIf DMonth.Text = "ມິຖຸນາ" Then
                MdStartDate = Format(CDate("01/06/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/06/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ມິຖຸນາ"
                DMonth.SelectedIndex = 5
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                MdStartDate_MM = Format(CDate("01/05/" & Year(MdStartDate)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/05/" & Year(MdStartDate)), "dd-MM-yyyy")


            ElseIf DMonth.Text = "ກໍລະກົດ" Then
                MdStartDate = Format(CDate("01/07/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/07/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ກໍລະກົດ"
                DMonth.SelectedIndex = 6
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                MdStartDate_MM = Format(CDate("01/06/" & Year(MdStartDate)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("30/06/" & Year(MdStartDate)), "dd-MM-yyyy")

            ElseIf DMonth.Text = "ສິງຫາ" Then
                MdStartDate = Format(CDate("01/08/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/08/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ສິງຫາ"
                DMonth.SelectedIndex = 7
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                MdStartDate_MM = Format(CDate("01/07/" & Year(MdStartDate)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/07/" & Year(MdStartDate)), "dd-MM-yyyy")


            ElseIf DMonth.Text = "ກັນຍາ" Then
                MdStartDate = Format(CDate("01/09/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/09/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ກັນຍາ"
                DMonth.SelectedIndex = 8
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                MdStartDate_MM = Format(CDate("01/08/" & Year(MdStartDate)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/08/" & Year(MdStartDate)), "dd-MM-yyyy")

            ElseIf DMonth.Text = "ຕຸລາ" Then
                MdStartDate = Format(CDate("01/10/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/10/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ຕຸລາ"
                DMonth.SelectedIndex = 9
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                MdStartDate_MM = Format(CDate("01/09/" & Year(MdStartDate)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("30/09/" & Year(MdStartDate)), "dd-MM-yyyy")

            ElseIf DMonth.Text = "ພະຈິກ" Then
                MdStartDate = Format(CDate("01/11/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/11/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ພະຈິກ"
                DMonth.SelectedIndex = 10
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                MdStartDate_MM = Format(CDate("01/10/" & Year(MdStartDate)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/10/" & Year(MdStartDate)), "dd-MM-yyyy")

            ElseIf DMonth.Text = "ທັນວາ" Then
                MdStartDate = Format(CDate("01/12/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/12/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "ທັນວາ"
                DMonth.SelectedIndex = 11
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                MdStartDate_MM = Format(CDate("01/11/" & Year(MdStartDate)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("30/11/" & Year(MdStartDate)), "dd-MM-yyyy")

            End If
            'Lb.Text = "ສຳລັບວັນທີ " & (MdToDate.Day) & " " & MonthLetter1 & " " & Year(MdToDate)
            Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
            INNM = "ທ້າຍເດືອນ " & CDbl(DMonth.SelectedIndex) + 1 & "/" & Year(MdToDate)
        Else

            If DMonth.Text = "January" Then
                MdStartDate = Format(CDate("01/01/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/01/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "January"
                DMonth.SelectedIndex = 0
            ElseIf DMonth.Text = "February" Then
                Dim Day As String
                Dim MM As Date
                Dim Fromm As Date
                MdStartDate = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
                Fromm = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
                MM = Format(CDate("01/3/" & Year(MdStartDate)), "dd-MM-yyyy")
                Day = DateDiff(DateInterval.Day, Fromm, MM)
                MdToDate = Format(CDate(Day & "/02" & "/" & Year(MdStartDate)), "dd-MM-yyyy")
                MonthLetter1 = "February"
                DMonth.SelectedIndex = 1
                Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
            ElseIf DMonth.Text = "March" Then
                MdStartDate = Format(CDate("01/03/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/03/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "March"
                DMonth.SelectedIndex = 2
            ElseIf DMonth.Text = "April" Then
                MdStartDate = Format(CDate("01/04/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/04/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "April"
                DMonth.SelectedIndex = 3
            ElseIf DMonth.Text = "May" Then
                MdStartDate = Format(CDate("01/05/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/05/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "May"
                DMonth.SelectedIndex = 4
            ElseIf DMonth.Text = "June" Then
                MdStartDate = Format(CDate("01/06/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/06/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "June"
                DMonth.SelectedIndex = 5
            ElseIf DMonth.Text = "July" Then
                MdStartDate = Format(CDate("01/07/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/07/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "July"
                DMonth.SelectedIndex = 6
            ElseIf DMonth.Text = "August" Then
                MdStartDate = Format(CDate("01/08/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/08/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "August"
                DMonth.SelectedIndex = 7
            ElseIf DMonth.Text = "September" Then
                MdStartDate = Format(CDate("01/09/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/09/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "September"
                DMonth.SelectedIndex = 8
            ElseIf DMonth.Text = "October" Then
                MdStartDate = Format(CDate("01/10/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/10/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "October"
                DMonth.SelectedIndex = 9
            ElseIf DMonth.Text = "November" Then
                MdStartDate = Format(CDate("01/11/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/11/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "November"
                DMonth.SelectedIndex = 10
            ElseIf DMonth.Text = "December" Then
                MdStartDate = Format(CDate("01/12/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/12/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "December"
                DMonth.SelectedIndex = 11
            End If
            Lb.Text = "For the Month Ended " & (MdToDate.Day) & " " & MonthLetter1 & " " & Year(MdToDate)
        End If

        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub

    Private Sub LoadPeriod()
        If Period.Text = "ໄຕມາດ 1" Then
            MdStartDate = Format(CDate("01/01/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/03/" & Year(Pyy.Value)), "dd-MM-yyyy")

            MdStartDate_PRV = Format(CDate("01/10/" & Year(Pyy.Value) - 1), "dd-MM-yyyy")
            MdToDate_PRV = Format(CDate("31/12/" & Year(Pyy.Value) - 1), "dd-MM-yyyy")


            Lb.Text = "ປະຈຳໄຕມາດ " & "1" & " ປີ " & Pyy.Text
        ElseIf Period.Text = "ໄຕມາດ 2" Then
            MdStartDate = Format(CDate("01/04/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/06/" & Year(Pyy.Value)), "dd-MM-yyyy")

            MdStartDate_PRV = Format(CDate("01/01/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate_PRV = Format(CDate("31/03/" & Year(Pyy.Value)), "dd-MM-yyyy")

            Lb.Text = "ປະຈຳໄຕມາດ " & "2" & " ປີ " & Pyy.Text
        ElseIf Period.Text = "ໄຕມາດ 3" Then
            MdStartDate = Format(CDate("01/07/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/09/" & Year(Pyy.Value)), "dd-MM-yyyy")

            MdStartDate_PRV = Format(CDate("01/04/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate_PRV = Format(CDate("30/06/" & Year(Pyy.Value)), "dd-MM-yyyy")

            Lb.Text = "ປະຈຳໄຕມາດ " & "3" & " ປີ " & Pyy.Text
        ElseIf Period.Text = "ໄຕມາດ 4" Then
            MdStartDate = Format(CDate("01/10/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/12/" & Year(Pyy.Value)), "dd-MM-yyyy")

            MdStartDate_PRV = Format(CDate("01/07/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate_PRV = Format(CDate("30/09/" & Year(Pyy.Value)), "dd-MM-yyyy")

            Lb.Text = "ປະຈຳໄຕມາດ " & "4" & " ປີ " & Pyy.Text
        End If
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳ" & Period.Text & " ປີ " & Year(MdToDate)
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub

    Private Sub LoadYear()
        MdStartDate = Format(CDate("01/1/" & Year(yy.Value)), "dd-MM-yyyy")
        MdToDate = Format(CDate("31/12/" & Year(Toyy.Value)), "dd-MM-yyyy")
        Lb.Text = "ປະຈຳປີ " & yy.Text
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd/MM/yyyy") & " ສີ້ນສຸດວັນທີ " & Format(MdToDate, "dd/MM/yyyy")
        L5.Text = MdStartDate & " => " & MdToDate
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
        'LngId = "7057" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
        LngId = "7095" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & " " & DTDATE02 & "' As Crl_RptName ,"
        LngId = "7002" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDP ,"
        LngId = "7059" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_PP ,"
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
        LngId = "7088" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Amt	 ,"
        LngId = "7039" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_TotalAmt ,"
        LngId = "7040" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Balance ,"
        LngId = "7043" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RemAmt ,"
        'LngId = "7048" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"


        'LngId = "7097" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"

        
        If RM.Checked = True Then
            LngId = "7096" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
            LngId = "7060" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7061" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        ElseIf RP.Checked = True Then
            LngId = "7112" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
            LngId = "7062" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7063" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"

        ElseIf RT.Checked = True Then
            LngId = "7112" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
            If Ct.SelectedIndex = 0 Then
                LngId = "7078" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
                LngId = "7079" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            Else
                LngId = "7079" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
                LngId = "7078" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
            End If


        ElseIf RY.Checked = True Then
            LngId = "7113" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
            LngId = "7064" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7065" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        End If

        'If CMB_Curr.Text = "EQVL" Then
        '    LngId = "7106" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        'ElseIf CMB_Curr.Text = "LAK" Then
        '    LngId = "7097" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        'ElseIf CMB_Curr.Text = "USD" Then
        '    LngId = "7108" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        'End If
        If CMB_Curr.Text = "EQVL" Then
            If CheckBox6.Checked = True Then
                LngId = "7121" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"

            Else
                LngId = "7106" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"

            End If
        ElseIf CMB_Curr.Text = "LAK" Then
            LngId = "7097" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        ElseIf CMB_Curr.Text = "USD" Then
            LngId = "7108" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        End If

        SLF = "SELECT   N'" & MuOffDep & "'  as RptSjoff_Dep  ,  " & mformat & "  as mformat  ,   " & MuLngRpt & "   *   FROM Ap_Rpt_Cashflow  "
        Call LoadLoGO()
        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            .Open("" & SLF & "where Rpt_Id <>'' " & RPT_ID & " " & r & "order by grp, Rpt_Id asc  ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim Rpt As New CryCashflow_statement2
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
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("Curr"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = LngStr
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("Text5"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = INNM
        'myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("Text2"), CrystalDecisions.CrystalReports.Engine.TextObject)
        'myText2.Text = Lb.Text
        myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("Text6"), CrystalDecisions.CrystalReports.Engine.TextObject)
        myText2.Text = CURR01
        Rpt.SetDataSource(Rs)
        FrmPreview.ReportViewer.ReportSource = Rpt
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
        LngId = "7057" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & " " & DTDATE02 & "' As Crl_RptName ,"
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
        'LngId = "7048" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
        If RM.Checked = True Then
            LngId = "7060" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7061" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        ElseIf RP.Checked = True Then
            LngId = "7062" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7063" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        ElseIf RY.Checked = True Then
            LngId = "7064" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
            LngId = "7065" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LastMounth ,"
        End If

        If CMB_Curr.Text = "EQVL" Then 
            LngId = "7106" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ," 
        ElseIf CMB_Curr.Text = "LAK" Then
            LngId = "7097" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        ElseIf CMB_Curr.Text = "USD" Then
            LngId = "7108" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"
        End If

        SLF = "SELECT  " & MuLngRpt & "  * ,  N'" & MuOffDep & "'  as RptSjoff_Dep FROM Ap_Rpt_Cashflow_Detail  "
        Call LoadLoGO()
        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            .Open(SLF, CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

            '.Open("SELECT *  ,N'" & txtReport_name & "' as txtReport_name  FROM Ap_Rpt_Cashflow where Amount_in_million_Kip <>0  ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim Rpt As New CryRpt_Cashflow_Item
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
        'myText2 = CType(Rpt.ReportDefinition.ReportObjects.Item("Text2"), CrystalDecisions.CrystalReports.Engine.TextObject)
        'myText2.Text = Lb.Text 
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
        LoadMonth()
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
        CMB_Curr_SelectedIndexChanged(sender, e)
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
        LoadDay()
    End Sub

    Private Sub Dt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Dt.ValueChanged
        LoadDay()
    End Sub

    Private Sub yy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yy.ValueChanged
        Call LoadYear()
        CMB_Curr_SelectedIndexChanged(sender, e)
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





        FmCashflow_Item1.ShowDialog()
        FmCashflow_Item1.Focus()
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
                LL6.Text = "ໄຕມາດ " & Period.SelectedIndex
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
    Private Sub BLS_PreV()
        Call Office()
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        'Call ChangBalance()
        BLNEW_Prev()
        CNN.Execute("update  Ap_Rpt_BLS_Item_Old set Last_amt_dr  =  " & CDbl(0) & " , Last_amt_cr  = " & CDbl(0) & " ,amt_dr  =  " & CDbl(0) & " , amt_cr  = " & CDbl(0) & "  ")
        CNN.Execute("update Ap_Rpt_BLS_Old set Last_Amt = 0 , Amt  =  " & CDbl(0) & "  ")
        CNN.Execute("DELETE FROM Ap_Rpt_BLS_Detail ")
        SelcectIn_BLS()
        UpdateIIn_BLS()
        SelectOut_BLS()
        UpdateOut_BLS()
        Update_Sum_BLS()
    End Sub
    Private Sub BLNEW_Prev()


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
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
               " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amount_Dr)as amt_dr , sum(amount_Cr)as amt_cr  from gen_jn  WHERE  1=1  and Curr=N'LAK'   and   gen_jn.date_work   BETWEEN '" & Format(MdStartDate_MM, "yyyy-MM-dd") & "' AND '" & Format(MdToDate_MM, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")

        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
   " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amount_Dr)* " & CDbl(txtRate.Text) & " as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & " as amt_cr  from gen_jn  WHERE  1=1  and Curr=N'USD'   and   gen_jn.date_work   BETWEEN '" & Format(MdStartDate_MM, "yyyy-MM-dd") & "' AND '" & Format(MdToDate_MM, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")

        Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate_MM)
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code , sum(amount_Dr)as amt_dr , sum(amount_Cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE  1=1   and Curr=N'LAK'   and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate_MM, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        Dim KK As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
      " select ac_code , sum(amount_Dr)* " & CDbl(txtRate.Text) & " as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & " as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE  1=1   and Curr=N'USD'   and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate_MM, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code"
        CNN.Execute(KK)


        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code  , sum(amount_Dr) as amt_dr , sum(amount_Cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1  and Curr=N'LAK' and     date_work='" & "1-1-" & Format(MdStartDate_MM, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        '       CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        '" select ac_code  , sum(amount_Dr)* " & CDbl(txtRate.Text) & "  as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & "  as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1  and Curr=N'USD'  and date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
" select ac_code  , sum(amount_Dr)  as amt_dr , sum(amount_Cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1  and Curr=N'USD'  and date_work='" & "1-1-" & Format(MdStartDate_MM, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

        CNN.Execute("UPDATE Ap_balance_6 set Ac_Code = left(Ac_Code,7) ")


        CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr)as open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")
        '    If CMB_Curr.SelectedIndex = 0 Then
        '        CNN.Execute(" insert into Ap_balance_6_col(ac_code, ac_name, open_amt_dr, open_amt_cr,  amt_dr, amt_cr,   Rem_dr, Rem_cr,Status) " & _
        '" Select '2382120.00.0000',N'ຄູ່ມູນຄ່າຖານະ ແລກປ່ຽນເງິນຕາ ຕ່າງປະເທດ',sum(open_amt_dr),sum(open_amt_cr),sum(amt_dr),sum(amt_cr),sum(Rem_dr),sum(Rem_cr),1   " & _
        '" from Ap_balance_6_col where left(ac_code,7)='2382120' group by  left(ac_code,7) ")
        '        CNN.Execute("delete  Ap_balance_6_col  where left(ac_code,7)='2382120' and Status is null ")
        '    End If
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
    Private Sub Chang_Incom12()
        If MDACC00 = 0 Then
            New_Code = New_Code

            Insr = "delete  Ap_balance_6  " & _
             "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr)   select '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr) from Ap_balance_6_col where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "' " & _
    "update  Ap_balance_6 set open_amt_dr = 0 where open_amt_dr  is null  " & _
    "update  Ap_balance_6 set open_amt_cr = 0 where open_amt_cr  is null  " & _
    "update  Ap_balance_6 set amt_dr = 0 where amt_dr  is null  " & _
    "update  Ap_balance_6 set amt_cr = 0 where amt_cr  is null   " & _
     "Update  Ap_balance_6 set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr   " & _
    "Update  Ap_balance_6 set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr " & _
    "Update  Ap_balance_6 set   amt_dr = amt_dr  - amt_cr , amt_cr=0  where amt_dr  >= amt_cr   " & _
    "Update  Ap_balance_6 set   amt_cr = amt_cr  - amt_dr , amt_dr=0  where amt_cr  >= amt_dr " & _
       "delete  Ap_balance_6_col  where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "'   " & _
         "  insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr ,status )  " & _
" select  '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr),1 from Ap_balance_6_col where Ac_Code =  '" & New_Code & "' " & _
  "       delete  Ap_balance_6_col where Ac_Code =  '" & New_Code & "' " & _
"  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , sum(open_amt_dr) , sum(open_amt_cr) , sum(amt_dr) , sum(amt_cr)  from Ap_balance_6 group by Ac_Code "

            CNN.Execute(Insr)
            If Month(MdStartDate) = 12 Then
                Insr = "delete  Ap_balance_6  " & _
              "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr)   select '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr) from Ap_balance_6_col where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "' " & _
     "update  Ap_balance_6 set open_amt_dr = 0 where open_amt_dr  is null  " & _
     "update  Ap_balance_6 set open_amt_cr = 0 where open_amt_cr  is null  " & _
     "update  Ap_balance_6 set amt_dr = 0 where amt_dr  is null  " & _
     "update  Ap_balance_6 set amt_cr = 0 where amt_cr  is null   " & _
      "Update  Ap_balance_6 set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr   " & _
     "Update  Ap_balance_6 set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr " & _
     "Update  Ap_balance_6 set   amt_dr = amt_dr  - amt_cr , amt_cr=0  where amt_dr  >= amt_cr   " & _
     "Update  Ap_balance_6 set   amt_cr = amt_cr  - amt_dr , amt_dr=0  where amt_cr  >= amt_dr " & _
        "delete  Ap_balance_6_col  where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "'   " & _
          "  insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr ,status )  " & _
 " select  '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr),1 from Ap_balance_6_col where Ac_Code =  '" & New_Code & "' " & _
   "       delete  Ap_balance_6_col where Ac_Code =  '" & New_Code & "' " & _
 "  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , sum(open_amt_dr) , sum(open_amt_cr) , sum(amt_dr) , sum(amt_cr)  from Ap_balance_6 group by Ac_Code "

                CNN.Execute(Insr)
            End If



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

    End Sub
    Private Sub BLS_LAST()
        Call Office()
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        'Call ChangBalance()
        BLNEW_LAST()
        CNN.Execute("update  Ap_Rpt_BLS_Item_Old set Last_amt_dr  =  " & CDbl(0) & " , Last_amt_cr  = " & CDbl(0) & " ,amt_dr  =  " & CDbl(0) & " , amt_cr  = " & CDbl(0) & "  ")
        CNN.Execute("update Ap_Rpt_BLS_Old set Last_Amt = 0 , Amt  =  " & CDbl(0) & "  ")
        CNN.Execute("DELETE FROM Ap_Rpt_BLS_Detail ")
        SelcectIn_BLS()
        UpdateIIn_BLS()
        SelectOut_BLS()
        UpdateOut_BLS()
        Update_Sum_BLS()
    End Sub
    Private Sub BLNEW_LAST()


        New_Code = "3901000"
        New_Code4 = "00.3901000"
        New_Code = "3901000"
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

        '       Dim GGG As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        '             " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(Amount_Dr)as amt_dr , sum(Amount_cr)as amt_cr  from gen_jn  WHERE 1=1  and Curr=N'LAK'   and gen_jn.date_work   BETWEEN '" & Format(MdStartDate_MM, "yyyy-MM-dd") & "' AND '" & Format(MdToDate2, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code "
        '       CNN.Execute(GGG)

        '       Dim USD As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        '" select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amount_Dr)* " & CDbl(txtRate.Text) & "  as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & " as amt_cr  from gen_jn  WHERE 1=1 and Curr=N'USD'  and gen_jn.date_work   BETWEEN '" & Format(MdStartDate_MM, "yyyy-MM-dd") & "' AND '" & Format(MdToDate2, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code "
        '       CNN.Execute(USD)

        Dim S As Date = MdStartDate_PRV : S = DateAdd("d", CDbl(-1), MdStartDate)


        'CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        '" select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        '=======LAK===
        Dim PPP As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select ac_code , sum(Amount_Dr)as amt_dr , sum(Amount_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE 1=1  and Curr=N'LAK'   and gen_jn.date_work    BETWEEN '" & Format(MdStartDate_PRV, "yyyy-MM-dd") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code"
        CNN.Execute(PPP)
        Dim PPPUSD As String = "INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
" select ac_code , sum(amount_Dr)* " & CDbl(txtRate.Text) & " as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & " as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE 1=1  and Curr=N'USD'  and gen_jn.date_work   BETWEEN '" & Format(MdStartDate_PRV, "yyyy-MM-dd") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code"
        CNN.Execute(PPPUSD)

        '        '=======LAK===
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code  , sum(Amount_Dr) as amt_dr , sum(Amount_cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1   and Curr=N'LAK'  and date_work='" & "1-1-" & Format(MdStartDate_PRV, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")


        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
   " select ac_code  , sum(amount_Dr)* " & CDbl(txtRate.Text) & "  as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & "  as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE  1=1  and Curr=N'USD'  and   date_work='" & "1-1-" & Format(MdStartDate_PRV, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")


        CNN.Execute("UPDATE Ap_balance_6 set Ac_Code = left(Ac_Code,7) ")


        CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr)as open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")
        '    If CMB_Curr.SelectedIndex = 0 Then
        '        CNN.Execute(" insert into Ap_balance_6_col(ac_code, ac_name, open_amt_dr, open_amt_cr,  amt_dr, amt_cr,   Rem_dr, Rem_cr,Status) " & _
        '" Select '2382120.00.0000',N'ຄູ່ມູນຄ່າຖານະ ແລກປ່ຽນເງິນຕາ ຕ່າງປະເທດ',sum(open_amt_dr),sum(open_amt_cr),sum(amt_dr),sum(amt_cr),sum(Rem_dr),sum(Rem_cr),1   " & _
        '" from Ap_balance_6_col where left(ac_code,7)='2382120' group by  left(ac_code,7) ")
        '        CNN.Execute("delete  Ap_balance_6_col  where left(ac_code,7)='2382120' and Status is null ")
        '    End If
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

    Private Sub RT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RT.CheckedChanged
        selectLoad()
    End Sub

    Private Sub yyt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yyt.ValueChanged
        selectLoad()
    End Sub

    Private Sub Ct_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ct.SelectedIndexChanged
        selectLoad()
        CMB_Curr_SelectedIndexChanged(sender, e)
    End Sub
End Class