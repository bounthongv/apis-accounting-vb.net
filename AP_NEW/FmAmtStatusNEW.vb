Public Class FmAmtStatusNEW
    Dim MdStartDate2, MdToDate2 As Date
    Dim r As String
    Dim CLT_Str, CLT_Last_Str As String
    Dim bls1 As String
    Dim MonthLetter1 As String
    Dim MdStartDate As Date
    Dim MdToDate As Date
    Dim MdStartDate_MM As Date
    Dim MdToDate_MM As Date
    Dim Month_IN_MM As String
    Dim MdQuarter As Date
    Dim MdStartDate_Last As Date
    Dim MdToDate_Last As Date

    Dim MonthLetter_Last As String
    Dim Month_IN As String
    Dim Month_Last As String

    Dim ny, ly, n_L_y As String

    Dim sql As String
    Dim AmtOpenDR, AmtOpenCR, AmtOpenMonthDR, AmtOpenMonthCR As Double
    Dim VCode1, VCode2, VCode3, VCode4, VCode5, VCode6, VCode7, VCode8, VCode9 As String
    'Dim MdQuarter As Date
    Dim RsOpen As New ADODB.Recordset
    Dim RsOpenMonth As New ADODB.Recordset
    Dim RsRpt As New ADODB.Recordset
    Dim VOpenDate As Date
    Dim RptNme As String
    Dim RSC12 As New ADODB.Recordset
    Dim RSCIn_M As New ADODB.Recordset
    Private Sub HeaDer()
        LoadSqlData("SELECT * FROM Header WHERE ID=N'B04' ", RSC)
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
            LoadSqlData("SELECT * FROM Header WHERE ID=N'B04' ", RSC)
            If RSC.RecordCount = 0 Then
                CNN.Execute("INSERT INTO Header(ID,Nm,S1,S2,S3,S4,PP) " & _
                            " values('B04',N'" & TxtHeader.Text & "',N'" & TxtS1.Text & "',N'" & TxtS2.Text & "',N'" & TxtS3.Text & "',N'" & TxtS4.Text & "',N'" & TxtPP.Text & "') ")
            Else
                CNN.Execute("UPDATE Header set Nm=N'" & TxtHeader.Text & "',S1=N'" & TxtS1.Text & "',S2=N'" & TxtS2.Text & "',S3=N'" & TxtS3.Text & "',S4=N'" & TxtS4.Text & "',PP=N'" & TxtPP.Text & "' " & _
                            " where ID='B04' ")
            End If
        Else
            LoadSqlData("SELECT * FROM Header WHERE ID=N'B04' ", RSC)
            If RSC.RecordCount = 0 Then
                CNN.Execute("INSERT INTO Header(ID,Nm,S1e,S2e,S3e,S4e,PPe) " & _
                            " values('B04',N'" & TxtHeader.Text & "',N'" & TxtS1.Text & "',N'" & TxtS2.Text & "',N'" & TxtS3.Text & "',N'" & TxtS4.Text & "',N'" & TxtPP.Text & "') ")
            Else
                CNN.Execute("UPDATE Header set Nm=N'" & TxtHeader.Text & "',S1e=N'" & TxtS1.Text & "',S2e=N'" & TxtS2.Text & "',S3e=N'" & TxtS3.Text & "',S4e=N'" & TxtS4.Text & "',PPe=N'" & TxtPP.Text & "' " & _
                            " where ID='B04' ")
            End If
        End If

    End Sub
    Private Sub FmAmtStatus_statement_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        If MuLng = "L" Then

            Label10.Text = "ລາຍເຊັນ1"
            Label14.Text = "ລາຍເຊັນ2"
            Label13.Text = "ລາຍເຊັນ3"
            Label12.Text = "ລາຍເຊັນ4"
            Label11.Text = "ທີ່"
            Ct.Items.Clear()
            Ct.Items.Add("6 ເດືອນຕົ້ນປີ")
            Ct.Items.Add("6 ເດືອນທ້າຍປີ")

        Else
            Ct.Items.Clear()
            Ct.Items.Add("First half year")
            Ct.Items.Add("Second half year")

            Label10.Text = "Signature1"
            Label14.Text = "Signature2"
            Label13.Text = "Signature3"
            Label12.Text = "Signature4"
            Label11.Text = "Location"


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
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr      ) " & _
        " select  ac_code  ,  0 , 0   , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  , 0 , 0   from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook & "   group BY ac_code")
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
           "" & _
            "  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr ,   amt_cr , amt_dr from Ap_balance_6"
            CNN.Execute(Insr)
        Else

            New_Code = New_Code4

            Insr = "delete  Ap_balance_6  " & _
                "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) values ('" & New_Code & "' ,0,0,0,0) " & _
             "update Ap_balance_6 set  open_amt_Dr =  " & _
             "(select top 1  (select  (Sum(open_amt_dr))-( Sum(open_amt_cr)) As open_amt_dr from Ap_balance_6_col where  left(Ac_Code,4) ='" & Code_Dr1 & "'  )  As Dr " & _
             "from Ap_balance_6_col )  where  Ac_Code ='" & New_Code & "'  " & _
         "update Ap_balance_6 set  open_amt_cr =  " & _
         "(select top 1  (select  (Sum(open_amt_cr))-( Sum(open_amt_dr)) As open_amt_dr from Ap_balance_6_col where   left(Ac_Code,4)='" & Code_Cr1 & "'  )  As Cr " & _
          "from Ap_balance_6_col ) where  Ac_Code ='" & New_Code & "'   " & _
         "update Ap_balance_6 set  amt_Dr = " & _
         "(select top 1  (select  (Sum(amt_dr))-( Sum(amt_cr)) As amt_dr from Ap_balance_6_col where left(Ac_Code,4) ='" & Code_Dr1 & "'  )  As Dr " & _
         "from Ap_balance_6_col )  where  Ac_Code ='" & New_Code & "'  " & _
         "update Ap_balance_6 set  amt_cr =  " & _
         "(select top 1  (select  (Sum(amt_cr))-( Sum(amt_dr)) As amt_dr from Ap_balance_6_col where   left(Ac_Code,4)='" & Code_Cr1 & "' )  As Cr " & _
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
           "" & _
            "  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr ,   amt_cr , amt_dr from Ap_balance_6"
            CNN.Execute(Insr)
        End If


    End Sub
    Private Sub MMM()
        CNN.Execute("DELETE  Ap_balance_6_col ")
        CNN.Execute("DELETE FROM Ap_balance_6 ")

        Dim ST As Integer = Month(MdStartDate) - 1

        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
        " select ac_code  ,  0 , 0  , 0  , 0  , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate_MM, "yyyy-MM-dd") & "' AND '" & Format(MdToDate_MM, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")
        Dim S As Date = MdStartDate_MM : S = DateAdd("d", CDbl(-1), MdStartDate_MM)
        'Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr      ) " & _
        " select  ac_code  ,  0 , 0   , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  , 0 , 0   from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate_MM, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        CNN.Execute("INSERT INTO Ap_balance_6 (  ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
        " select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr   ,  0 , 0  , 0 , 0  from Open_jn WHERE    date_work='" & "1-1-" & Format(MdStartDate_MM, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr  ) " & _
        " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr) as open_amt_cr , sum(Amt_Last_M_Dr) as Amt_Last_M_Dr , sum(Amt_Last_M_Cr) as Amt_Last_M_Cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        New_Code = "3901000"
        Code_Dr = "4"
        Code_Cr = "5"
        '==============
        New_Code = "3901000"
        New_Code4 = "00.3901000"
        Code_Dr = "4"
        Code_Dr1 = "00.4"
        Code_Cr = "5"
        Code_Cr1 = "00.5"

        If MDACC00 = 0 Then
            New_Code = New_Code
        Else

            New_Code = New_Code4
        End If

        Call Chang_Incom()

        CNN.Execute("update  Ap_balance_6_col set Amt_Last_M_dr = 0 where Amt_Last_M_dr  is null")
        CNN.Execute("update  Ap_balance_6_col set Amt_Last_M_cr = 0 where Amt_Last_M_cr  is null")

        CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + Amt_Last_M_Dr + amt_dr) - (open_amt_cr + Amt_Last_M_Cr + amt_cr) where (open_amt_dr + Amt_Last_M_Dr + amt_dr) >= (open_amt_cr + Amt_Last_M_Cr + amt_cr) ")
        CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr +  Amt_Last_M_Cr + amt_cr) - (open_amt_dr + Amt_Last_M_Dr + amt_dr) where (open_amt_cr + Amt_Last_M_Cr + amt_cr) >= (open_amt_dr + Amt_Last_M_Dr + amt_dr) ")
        CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        'Call ChangBalance()

    End Sub
    Private Sub MMM22()
        CNN.Execute("DELETE  Ap_balance_6_col ")
        CNN.Execute("DELETE FROM Ap_balance_6 ")
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
        " select ac_code  ,  0 , 0  , 0  , 0  , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")
        Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr      ) " & _
        " select  ac_code  ,  0 , 0   , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  , 0 , 0   from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        CNN.Execute("INSERT INTO Ap_balance_6 (  ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
        " select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr   ,  0 , 0  , 0 , 0  from Open_jn WHERE    date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr  ) " & _
        " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr) as open_amt_cr , sum(Amt_Last_M_Dr) as Amt_Last_M_Dr , sum(Amt_Last_M_Cr) as Amt_Last_M_Cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        New_Code = "3901000"
        Code_Dr = "4"
        Code_Cr = "5"
        '==============
        New_Code = "3901000"
        New_Code4 = "00.3901000"
        Code_Dr = "4"
        Code_Dr1 = "00.4"
        Code_Cr = "5"
        Code_Cr1 = "00.5"

        If MDACC00 = 0 Then
            New_Code = New_Code
        Else

            New_Code = New_Code4
        End If

        Call Chang_Incom()

        CNN.Execute("update  Ap_balance_6_col set Amt_Last_M_dr = 0 where Amt_Last_M_dr  is null")
        CNN.Execute("update  Ap_balance_6_col set Amt_Last_M_cr = 0 where Amt_Last_M_cr  is null")

        CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + Amt_Last_M_Dr + amt_dr) - (open_amt_cr + Amt_Last_M_Cr + amt_cr) where (open_amt_dr + Amt_Last_M_Dr + amt_dr) >= (open_amt_cr + Amt_Last_M_Cr + amt_cr) ")
        CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr +  Amt_Last_M_Cr + amt_cr) - (open_amt_dr + Amt_Last_M_Dr + amt_dr) where (open_amt_cr + Amt_Last_M_Cr + amt_cr) >= (open_amt_dr + Amt_Last_M_Dr + amt_dr) ")
        CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        'Call ChangBalance()

    End Sub

    Private Sub BtnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        AddHeader()
        'Click_Last()
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        CNN.Execute("DELETE  Ap_balance_6_col ")
        CNN.Execute("DELETE FROM Ap_balance_6 ")
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
        " select ac_code  ,  0 , 0  , 0  , 0  , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")
        Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
        CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr      ) " & _
        " select  ac_code  ,  0 , 0   , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  , 0 , 0   from gen_jn  WHERE gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        CNN.Execute("INSERT INTO Ap_balance_6 (  ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr   ) " & _
        " select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr   ,  0 , 0  , 0 , 0  from Open_jn WHERE    date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        CNN.Execute("INSERT INTO Ap_balance_6_col ( ac_code  , open_amt_dr, open_amt_cr , Amt_Last_M_Dr , Amt_Last_M_Cr , amt_dr , amt_cr  ) " & _
        " select ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr) as open_amt_cr , sum(Amt_Last_M_Dr) as Amt_Last_M_Dr , sum(Amt_Last_M_Cr) as Amt_Last_M_Cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY ac_code")
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        New_Code = "3901000"
        Code_Dr = "4"
        Code_Cr = "5"
        '==============
        New_Code = "3901000"
        New_Code4 = "00.3901000"
        Code_Dr = "4"
        Code_Dr1 = "00.4"
        Code_Cr = "5"
        Code_Cr1 = "00.5"

        If MDACC00 = 0 Then
            New_Code = New_Code
        Else 
            New_Code = New_Code4
        End If

        Call Chang_Incom()

        CNN.Execute("update  Ap_balance_6_col set Amt_Last_M_dr = 0 where Amt_Last_M_dr  is null")
        CNN.Execute("update  Ap_balance_6_col set Amt_Last_M_cr = 0 where Amt_Last_M_cr  is null")

        CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + Amt_Last_M_Dr + amt_dr) - (open_amt_cr + Amt_Last_M_Cr + amt_cr) where (open_amt_dr + Amt_Last_M_Dr + amt_dr) >= (open_amt_cr + Amt_Last_M_Cr + amt_cr) ")
        CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr +  Amt_Last_M_Cr + amt_cr) - (open_amt_dr + Amt_Last_M_Dr + amt_dr) where (open_amt_cr + Amt_Last_M_Cr + amt_cr) >= (open_amt_dr + Amt_Last_M_Dr + amt_dr) ")
        CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        Call ChangBalance()
        'Call MPREV()

        CNN.Execute("update AP_Rpt_Amt_Status set Amt1=0 ,amt2=0 , Amt3=0 , Amt4=0  ,Amt5=0 ,Amt6=0")
        CNN.Execute("DELETE TEST_ABC ")
        CNN.Execute("INSERT INTO TEST_ABC(Rpt_ID,Name,amt) select Rpt_ID,descriptione,Amt1 from AP_Rpt_Amt_Status ")

        Call M_10()
        Dim ds As Date
        ds = DateAdd(DateInterval.Year, 0, MdStartDate)

        '========== - ຜົນກະທົບຈາກອັດຕາແລກປ່ຽນເງິນຕາຕ່າງປະເທດ
        Dim rs As New ADODB.Recordset
        Dim MDRate_Last As String
        MDRate_Last = " and rate_dt<='" & Format(dpMonthPrev.Value, "yyyy-MM-dd") & "'  "
        MDRate_Last = " and month(rate_dt)<='" & Month(dpMonthPrev.Value) - 1 & "' and  year(rate_dt)='" & Year(dpMonthPrev.Value) & "' "


        Call LoadSqlData("select top 1  * from AP_Rate_history where 1=1  " & MDRate_Last & "  and curr='USD'  ORDER BY rate_dt DESC ", rs)
        If rs.RecordCount > 0 Then
            MD_Rate = (rs.Fields("Rate").Value)
        End If

        Dim KKq As String = "update TEST_ABC set Amt=Amt+  (select (sum(amount_dr)-sum(amount_cr))* " & CDbl(MD_Rate) & " from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)<='" & Month(ds) - 2 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='15'"
        CNN.Execute(KKq)

        'Dim KK2qq As String = "update TEST_ABC set Amt=Amt+(select (sum(amount_dr)-sum(amount_cr))* " & CDbl(MD_Rate) & "  from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)<='" & Month(ds) - 2 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='15'"
        'CNN.Execute(KK2qq)
        '========== - Currrmmmmm
        MDRate_DT = " and rate_dt<='" & Format(MdToDate, "yyyy-MM-dd") & "'  "
        MDRate_DT = " and month(rate_dt)<='" & Month(MdToDate) - 1 & "' and  year(rate_dt)='" & Year(MdToDate) & "' "

        Call LoadSqlData("select top 1  * from AP_Rate_history where 1=1  " & MDRate_DT & "  and curr='USD'  ORDER BY rate_dt DESC ", rs)
        If rs.RecordCount > 0 Then
            MD_Rate = (rs.Fields("Rate").Value)
        End If
        '========== - ຜົນກະທົບຈາກອັດຕາແລກປ່ຽນເງິນຕາຕ່າງປະເທດ
        Dim KK2w As String = "update TEST_ABC set Amt=Amt-(select (sum(amount_dr)-sum(amount_cr))* " & CDbl(MD_Rate) & "  from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)<='" & Month(ds) - 2 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='15'"
        CNN.Execute(KK2w)

        'MsgBox(DisT)
        If MDACC00 = 0 Then
            If Month(MdStartDate) <> 1 Then
                Call MMM()
                CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '310%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '310%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Cr- Amt_Dr) from Open_jn where Ac_Code Like '380%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")

                If Month(MdStartDate) > 2 Then 
                    If RP.Checked = True Then
                        If Period.SelectedIndex = 1 Then
                            'Dim AA1 As String = "update AP_Rpt_Amt_Status set Amt4= Amt4+ (select sum(Amt_cr)-sum(Amt_dr) from gen_jn where   (left(ac_code,1)='4' or left(ac_code,1)='5')  and Month(Date_Work)<'" & Month(ds) & "' and Year(Date_Work)='" & Year(ds) & "' )  where rpt_id='01'"
                            'CNN.Execute(AA1)
                        ElseIf Period.SelectedIndex = 2 Then
                            'Dim AA2 As String = "update AP_Rpt_Amt_Status set Amt4= Amt4+ (select sum(Amt_cr)-sum(Amt_dr) from gen_jn where   (left(ac_code,1)='4' or left(ac_code,1)='5') and Month(Date_Work)<'" & Month(ds) & "' and Year(Date_Work)='" & Year(ds) & "' )  where rpt_id='01'"
                            'CNN.Execute(AA2)
                            Dim AA3 As String = "update AP_Rpt_Amt_Status set Amt4= Amt4+ (select sum(Amt_cr)-sum(Amt_dr) from gen_jn where   (left(ac_code,1)='4' or left(ac_code,1)='5') and Month(Date_Work)<'" & Month(ds) - 3 & "' and Year(Date_Work)='" & Year(ds) & "' )  where rpt_id='01'"
                            CNN.Execute(AA3)
                        ElseIf Period.SelectedIndex = 3 Then
                            'Dim AA2 As String = "update AP_Rpt_Amt_Status set Amt4= Amt4+ (select sum(Amt_cr)-sum(Amt_dr) from gen_jn where   (left(ac_code,1)='4' or left(ac_code,1)='5') and Month(Date_Work)<'" & Month(ds) & "' and Year(Date_Work)='" & Year(ds) & "' )  where rpt_id='01'"
                            'CNN.Execute(AA2)
                            Dim AA3 As String = "update AP_Rpt_Amt_Status set Amt4= Amt4+ (select sum(Amt_cr)-sum(Amt_dr) from gen_jn where   (left(ac_code,1)='4' or left(ac_code,1)='5') and Month(Date_Work)<'" & Month(ds) - 3 & "' and Year(Date_Work)='" & Year(ds) & "' )  where rpt_id='01'"
                            CNN.Execute(AA3)
                            'Dim AA3 As String = "update AP_Rpt_Amt_Status set Amt4= Amt4+ (select sum(Amt_cr)-sum(Amt_dr) from gen_jn where   (left(ac_code,1)='4' or left(ac_code,1)='5') and Month(Date_Work)='" & Month(ds) & "' and Year(Date_Work)='" & Year(ds) & "' )  where rpt_id='01'"
                            'CNN.Execute(AA3)
                        End If


                    Else
                        'Dim AA As String = "update AP_Rpt_Amt_Status set Amt4= Amt4+ (select sum(Amt_cr)-sum(Amt_dr) from gen_jn where   (left(ac_code,1)='4' or left(ac_code,1)='5')  and Month(Date_Work)<'" & Month(ds) & "' and Year(Date_Work)='" & Year(ds) & "' )  where rpt_id='01'"
                        'CNN.Execute(AA)

                        'Dim AA2 As String = "update AP_Rpt_Amt_Status set Amt4= Amt4+ (select sum(Amt_cr)-sum(Amt_dr) from gen_jn where   (left(ac_code,1)='4' or left(ac_code,1)='5')  and Month(Date_Work)<'" & Month(ds) - 1 & "' and Year(Date_Work)='" & Year(ds) & "' )  where rpt_id='01'"
                        'CNN.Execute(AA2)
                        If RM.Checked = True Then
                            Dim AA As String = "update AP_Rpt_Amt_Status set Amt4= Amt4+ (select isnull(sum(Amt_cr)-sum(Amt_dr),0) from gen_jn where   (left(ac_code,1)='4' or left(ac_code,1)='5')  and Month(Date_Work)<'" & Month(ds) - 1 & "' and Year(Date_Work)='" & Year(ds) & "' )  where rpt_id='01'"
                            CNN.Execute(AA)
                        Else

                            Dim AA As String = "update AP_Rpt_Amt_Status set Amt4= Amt4+ (select isnull(sum(Amt_cr)-sum(Amt_dr),0) from gen_jn where   (left(ac_code,1)='4' or left(ac_code,1)='5')  and Month(Date_Work)<'" & Month(ds) & "' and Year(Date_Work)='" & Year(ds) & "' )  where rpt_id='01'"
                            CNN.Execute(AA)
                        End If
                    End If

                End If
                'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_dr)-sum(Amt_cr) from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)<='" & Month(ds) - 1 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")
                ''========== - ຜົນກະທົບຈາກອັດຕາແລກປ່ຽນເງິນຕາຕ່າງປະເທດ
                'Dim KK As String = "update AP_Rpt_Amt_Status set Amt1=Amt1-(select (sum(amount_dr)-sum(amount_cr))*9576  from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)<='" & Month(ds) - 1 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'"
                'CNN.Execute(KK)
                'Dim Pq As String = "update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)='" & Month(ds) - 1 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='10'"
                'CNN.Execute(Pq)

                CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '3202%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")

                CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '320%'  And Ac_Code<> '3202' And Year(Date_Work)=" & Year(ds) & ")  where rpt_id='01'")

                'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Cr- Amt_Dr) from Open_jn where Ac_Code Like '390%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")
                '            -  - ທຶນຈົດທະບຽນ ທີ່ໄດ້ຮັບ
                'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select (sum(amount_cr))* " & CDbl(MD_Rate) & " from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)='" & Month(ds) - 1 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='10'")
                If RM.Checked = True Then
                    CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select (sum(amount_cr))* " & CDbl(MD_Rate) & " from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)='" & Month(ds) - 1 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='10'")

                ElseIf RP.Checked = True Then
                    MdStartDate2 = DateAdd(DateInterval.Month, -3, MdStartDate)
                    MdToDate2 = DateAdd(DateInterval.Day, -1, MdStartDate)
                    MDRate_Last = " and month(rate_dt)<='" & Month(MdToDate2) & "' and  year(rate_dt)='" & Year(MdToDate2) & "' "

                    Call LoadSqlData("select top 1  * from AP_Rate_history where 1=1  " & MDRate_Last & "  and curr='USD'  ORDER BY rate_dt DESC ", rs)
                    If rs.RecordCount > 0 Then
                        MD_Rate = (rs.Fields("Rate").Value)
                    End If

                    'Dim LAST As String = "update Ap_Rpt_Cashflow set  Last_amt=(   select sum(amt_cr)-sum(amt_dr) from gen_jn   where (left(ac_code,1)='4' or left(ac_code,1)='5' )  and date_work BETWEEN '" & Format(MdStartDate2, "yyyy-MM-dd") & "' AND '" & Format(MdToDate2, "yyyy-MM-dd") & "'  " & MULook2 & " ) where Rpt_Id='02'"
                    'CNN.Execute(LAST)
                    'Dim PK2 As String = "update AP_Rpt_Amt_Status set Amt1=  (select (sum(amount_cr))* " & CDbl(MD_Rate) & " from Gen_jn where (Ac_Code Like '310%' )  And  Date_Work BETWEEN '" & Format(MdStartDate2, "yyyy-MM-dd") & "' AND '" & Format(MdToDate2, "yyyy-MM-dd") & "'  )  where rpt_id='10'"
                    'CNN.Execute(PK2)
                    Dim PP As String = "update AP_Rpt_Amt_Status set Amt1=  (select (sum(amount_cr)* " & CDbl(MD_Rate) & " - sum(amount_dr)* " & CDbl(MD_Rate) & ") from Gen_jn where curr='USD' and  (Ac_Code Like '310%' )  And  Date_Work   BETWEEN '" & Format(MdStartDate2, "yyyy-MM-dd") & "' AND '" & Format(MdToDate2, "yyyy-MM-dd") & "'  )  where rpt_id='10'"
                    CNN.Execute(PP)
                    Dim PP1 As String = "update AP_Rpt_Amt_Status set Amt1= Amt1+ (select isnull(sum(amount_cr)-sum(amount_dr) ,0) from Gen_jn where  curr='LAK' and  (Ac_Code Like '310%' )  And  Date_Work   BETWEEN '" & Format(MdStartDate2, "yyyy-MM-dd") & "' AND '" & Format(MdToDate2, "yyyy-MM-dd") & "'  )  where rpt_id='10'"
                    CNN.Execute(PP1)
                Else


                End If
                'Dim xa As String = "update TEST_ABC set Amt=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)='" & Month(ds) - 2 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='10'"
                'CNN.Execute(xa)

                'Dim Pq As String = "update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)='" & Month(ds) - 1 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='10'"
                'CNN.Execute(Pq)

                'Dim Pq2 As String = "update AP_Rpt_Amt_Status set Amt1= Amt1+ (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)='" & Month(ds) - 1 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'"
                'CNN.Execute(Pq2)

                'Dim Pq3 As String = "update AP_Rpt_Amt_Status set Amt1= Amt1+ (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)='" & Month(ds) - 1 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'"
                'CNN.Execute(Pq3)

                CNN.Execute("update AP_Rpt_Amt_Status set Amt4=   (select sum(Amt_dr-Amt_cr) from Ap_balance_6 where (Ac_Code Like '390%'))   where rpt_id='11'")
                '========== - ຜົນກະທົບຈາກອັດຕາແລກປ່ຽນເງິນຕາຕ່າງປະເທດ

                MDRate_Last = " and rate_dt<='" & Format(dpMonthPrev.Value, "yyyy-MM-dd") & "'  "
                MDRate_Last = " and month(rate_dt)<='" & Month(dpMonthPrev.Value) & "' and  year(rate_dt)='" & Year(dpMonthPrev.Value) & "' "


                Call LoadSqlData("select top 1  * from AP_Rate_history where 1=1  " & MDRate_Last & "  and curr='USD'  ORDER BY rate_dt DESC ", rs)
                If rs.RecordCount > 0 Then
                    MD_Rate = (rs.Fields("Rate").Value)
                End If
                 
                Dim KK As String = "update AP_Rpt_Amt_Status set Amt1=  (select (sum(amount_dr)-sum(amount_cr))* " & CDbl(MD_Rate) & " from Gen_jn where (Ac_Code Like '310%' )  and curr='USD'  And Month(Date_Work)<='" & Month(ds) - 1 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='15'"
                CNN.Execute(KK)

                '========== - Currrmmmmm
                MDRate_DT = " and rate_dt<='" & Format(MdToDate, "yyyy-MM-dd") & "'  "
                Call LoadSqlData("select top 1  * from AP_Rate_history where 1=1  " & MDRate_DT & "  and curr='USD'  ORDER BY rate_dt DESC ", rs)
                If rs.RecordCount > 0 Then
                    MD_Rate = (rs.Fields("Rate").Value)
                End If
                '========== - ຜົນກະທົບຈາກອັດຕາແລກປ່ຽນເງິນຕາຕ່າງປະເທດ
                Dim KK2 As String = "update AP_Rpt_Amt_Status set Amt1=Amt1-(select (sum(amount_dr)-sum(amount_cr))* " & CDbl(MD_Rate) & "  from Gen_jn where (Ac_Code Like '310%' )   and curr='USD' And Month(Date_Work)<='" & Month(ds) - 1 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='15'"
                CNN.Execute(KK2)


                If RP.Checked = True Then
                    MdStartDate2 = DateAdd(DateInterval.Month, -3, MdStartDate)
                    MdToDate2 = DateAdd(DateInterval.Day, -1, MdStartDate)
                    MDRate_Last = " and month(rate_dt)<='" & Month(MdToDate2) & "' and  year(rate_dt)='" & Year(MdToDate2) & "' "


     
                    Call LoadSqlData("select top 1  * from AP_Rate_history where 1=1  " & MDRate_Last & "  and curr='USD'  ORDER BY rate_dt DESC ", rs)
                    If rs.RecordCount > 0 Then
                        MD_Rate = (rs.Fields("Rate").Value)
                    End If


                    Dim KK3 As String = "update AP_Rpt_Amt_Status set Amt1=  (select (sum(amount_dr)-sum(amount_cr))* " & CDbl(MD_Rate) & " from Gen_jn where     (Ac_Code Like '310%' )  and curr='USD' And  Date_Work   BETWEEN '" & Format(MdStartDate2, "yyyy-MM-dd") & "' AND '" & Format(MdToDate2, "yyyy-MM-dd") & "'  )  where rpt_id='15'"
                    CNN.Execute(KK3)

                    '========== - Currrmmmmm
                    MDRate_DT = " and rate_dt<='" & Format(MdToDate, "yyyy-MM-dd") & "'  "
                    Call LoadSqlData("select top 1  * from AP_Rate_history where 1=1  " & MDRate_DT & "  and curr='USD'  ORDER BY rate_dt DESC ", rs)
                    If rs.RecordCount > 0 Then
                        MD_Rate = (rs.Fields("Rate").Value)
                    End If
                    '========== - ຜົນກະທົບຈາກອັດຕາແລກປ່ຽນເງິນຕາຕ່າງປະເທດ
                    Dim KK4 As String = "update AP_Rpt_Amt_Status set Amt1=Amt1+(select (sum(amount_dr)-sum(amount_cr))* " & CDbl(MD_Rate) & "  from Gen_jn where (Ac_Code Like '310%' )   and curr='USD' And  Date_Work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  )  where rpt_id='15'"
                    CNN.Execute(KK4)

                End If

                If Month(MdStartDate) > 1 Then
                    CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_Cr-Amt_Dr) from Open_jn where (Ac_Code Like '330%' or Ac_Code Like '340%'or Ac_Code Like '350%' or Ac_Code Like '360%'or Ac_Code Like '370%' )  And Year(Date_Work)='" & Year(ds) & "' )  where rpt_id='01'")

                End If
                '==============NEW=====
                'CNN.Execute("UPDATE TEST_ABC set amt=0 where amt is null")
                'Dim aq As String = "UPDATE AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.Amt1=AP_Rpt_Amt_Status.Amt1+(select sum(amt) from TEST_ABC where  (rpt_id='10' or rpt_id='15') ) where  rpt_id='01' "
                'CNN.Execute(aq)

                'CNN.Execute("INSERT INTO TEST_MM(Rpt_ID,Amt,MM) select '01',Amt1,'" & Format(CDate(MdStartDate), "yyyy-MM-dd") & "' from AP_Rpt_Amt_Status where rpt_id='16' ")

                If RP.Checked = True Then
                    CNN.Execute("update AP_Rpt_Amt_Status set Amt1=(select sum(Amt1) from AP_Rpt_Amt_Status where   rpt_id<13)  where rpt_id='13' ")
                    CNN.Execute("update AP_Rpt_Amt_Status set Amt2=(select sum(Amt2) from AP_Rpt_Amt_Status where  rpt_id<13)  where rpt_id='13' ")
                    CNN.Execute("update AP_Rpt_Amt_Status set Amt3=(select sum(Amt3) from AP_Rpt_Amt_Status where   rpt_id<13)  where rpt_id='13' ")
                    CNN.Execute("update AP_Rpt_Amt_Status set Amt4=(select sum(Amt4) from AP_Rpt_Amt_Status where   rpt_id<13)  where rpt_id='13' ")
                    CNN.Execute("update AP_Rpt_Amt_Status set Amt5=(select sum(Amt5) from AP_Rpt_Amt_Status where   rpt_id<13)  where rpt_id='13' ")
                Else
                    '==============NEW=====
                    CNN.Execute("UPDATE TEST_ABC set amt=0 where amt is null")
                    Dim aq As String = "UPDATE AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.Amt1=AP_Rpt_Amt_Status.Amt1+(select sum(amt) from TEST_ABC where  (rpt_id='10' or rpt_id='15') ) where  rpt_id='01' "
                    CNN.Execute(aq)

                    CNN.Execute("INSERT INTO TEST_MM(Rpt_ID,Amt,MM) select '01',Amt1,'" & Format(CDate(MdStartDate), "yyyy-MM-dd") & "' from AP_Rpt_Amt_Status where rpt_id='16' ")

                    Dim PK As String = "update AP_Rpt_Amt_Status set Amt1=(select Amt from TEST_MM where   rpt_id='01'  and Month(MM)='" & Month(ds) - 1 & "' And Year(MM)='" & Year(ds) & "')  where rpt_id='01' "
                    CNN.Execute(PK)
                    CNN.Execute("update AP_Rpt_Amt_Status set Amt1=(select sum(Amt1) from AP_Rpt_Amt_Status where   rpt_id<13)  where rpt_id='13' ")
                    CNN.Execute("update AP_Rpt_Amt_Status set Amt2=(select sum(Amt2) from AP_Rpt_Amt_Status where  rpt_id<13)  where rpt_id='13' ")
                    CNN.Execute("update AP_Rpt_Amt_Status set Amt3=(select sum(Amt3) from AP_Rpt_Amt_Status where   rpt_id<13)  where rpt_id='13' ")
                    CNN.Execute("update AP_Rpt_Amt_Status set Amt4=(select sum(Amt4) from AP_Rpt_Amt_Status where   rpt_id<13)  where rpt_id='13' ")
                    CNN.Execute("update AP_Rpt_Amt_Status set Amt5=(select sum(Amt5) from AP_Rpt_Amt_Status where   rpt_id<13)  where rpt_id='13' ")
                End If





                'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=(select sum(Amt4) from AP_Rpt_Amt_Status where   rpt_id<13)  where rpt_id='14' ")

                'CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '3201%'  And Ac_Code Like '3206%' And  Year(Date_Work)='" & Year(ds) & "')  where rpt_id='03'")

                'CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '3202000%'  And  Year(Date_Work)='" & Year(ds) & "')  where rpt_id='04'")
                'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '3202000%' And  Year(Date_Work)='" & Year(ds) & "') * -1 where rpt_id='04'")

                'CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '3203%'  And Ac_Code <> '3202%' And  Year(Date_Work)='" & Year(ds) & "')  where rpt_id='05'")
                'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum( Amt_cr) from Gen_jn where Ac_Code Like '3203%'  And Ac_Code <> '3202%' And Year(Date_Work)='" & Year(ds) & "') * -1  where rpt_id='05'")

                'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum( Amt_cr) from Gen_jn where Ac_Code Like '3108210%'  And Ac_Code <> '3202%' And Year(Date_Work)='" & Year(ds) & "') * -1  where rpt_id='06'")
                ''CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Cr) from Open_jn where Ac_Code Like '3108210%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='07'")

                'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Dr) from Gen_jn where Ac_Code Like '3908000%' And Year(Date_Work)='" & Year(MdStartDate) & "')  where rpt_id='07'")

                'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status  where rpt_id='01' Or rpt_id='06' )  where rpt_id='08'")
                'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status  where rpt_id='01' Or rpt_id='03' Or rpt_id='04' Or rpt_id='05' Or rpt_id='07' )  where rpt_id='08'")

                'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status  where rpt_id='08' )  where rpt_id='10'")
                'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status  where rpt_id='08' )  where rpt_id='10'")

                'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status  where rpt_id='10' Or  rpt_id='11' )  where rpt_id='12'")
                'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status  where rpt_id='10'  Or  rpt_id='11' )  where rpt_id='12'")
            End If
            '============
            MMM22()
            ds = DateAdd(DateInterval.Year, 0, MdStartDate)
            Dim KAS As String = "update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '310%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='16'"
            CNN.Execute(KAS)

            CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '3202%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='16'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '320%'  And Ac_Code<> '3202' And Year(Date_Work)=" & Year(ds) & ")  where rpt_id='16'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Cr- Amt_Dr) from Open_jn where Ac_Code Like '380%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='16'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_Cr- Amt_Dr) from Open_jn where (Ac_Code Like '330%' or Ac_Code Like '340%'or Ac_Code Like '350%' or Ac_Code Like '360%'or Ac_Code Like '370%' )  And Year(Date_Work)='" & Year(ds) & "' )  where rpt_id='16'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=Amt1+  (select sum(Amt_Cr+ Amt_Dr) from gen_jn where Ac_Code Like '310%' And month(Date_Work)='" & Month(ds) & "'  And Year(Date_Work)='" & Year(ds) & "' )  where rpt_id='01'")
            'Dim PKa As String = "update AP_Rpt_Amt_Status set Amt1=(select Amt from TEST_MM where   rpt_id='01'  and Month(MM)='" & Month(ds) - 1 & "' And Year(MM)='" & Year(ds) & "')  where rpt_id='01' "
            'CNN.Execute(PKa)

            If Month(MdStartDate) <> 1 Then
                CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status where rpt_id='13' )  where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt2) from AP_Rpt_Amt_Status where rpt_id='13' )  where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt3) from AP_Rpt_Amt_Status where rpt_id='13' )  where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status where rpt_id='13' )  where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt5) from AP_Rpt_Amt_Status where rpt_id='13' )  where rpt_id='16'")

            End If

            If Month(MdStartDate) > 2 Then
                CNN.Execute("update AP_Rpt_Amt_Status set Amt1= 0    where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt2= 0 where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt3= 0 where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt4=0 where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt5= 0 where rpt_id='16'")

                CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status where rpt_id>12 and  rpt_id<16 )  where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt2) from AP_Rpt_Amt_Status where rpt_id>12  and  rpt_id<16 )  where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt3) from AP_Rpt_Amt_Status where rpt_id>12  and  rpt_id<16 )  where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status where rpt_id>12  and  rpt_id<16 )  where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt5) from AP_Rpt_Amt_Status where rpt_id>12  and  rpt_id<16)  where rpt_id='16'")

            End If
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='14'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt2) from AP_Rpt_Amt_Status where    rpt_id<13 )  where rpt_id='14'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt3) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='14'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='14'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt5) from AP_Rpt_Amt_Status where  rpt_id<13)  where rpt_id='14'")

            ' - ສ່ວນຜິດດ່ຽງຈາກການຕີມູນຄ່າຊັບສິນ
            CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '3201%' or Ac_Code Like '3206%')  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='19'")
            '    - ທຶນຊ່ວຍໜູນແລະ ທຶນທີ່ໄດ້ຮັບຈັດສັນ
            CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '3701%' or Ac_Code Like '3702%')  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='20'")
            '       - ໂອນທຶນຊ່ວຍໜູນກໍ່ສ້າງພື້ນຖານເຂົ້າບັນຊີລາຍຮັບ
            CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '37012%' )  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='21'")
            '         - ຊຳລະຄືນທຶນທີ່ໄດ້ຮັບຈັດສັນ 
            CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '3702%' )  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='22'")
            '           - ໂອນກຳໄລສຸດທິເຂົ້າຄັງສຳຮອງຕາມກົດໝາຍ 
            CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '3202%' )  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='24'")
            '            - ໂອນກຳໄລສຸດທິເຂົ້າຄັງສຳຮອງທົ່ວໄປ 
            CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '3203%' or Ac_Code Like '3204%' or Ac_Code Like '3205%' or Ac_Code Like '3208%' or Ac_Code Like '350%' )  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='25'")
            '            -  - ທຶນຈົດທະບຽນ ທີ່ໄດ້ຮັບ
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '310%' )  And Month(Date_Work)='" & Month(ds) & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='24'")
            'Dim PP As String = "update AP_Rpt_Amt_Status set Amt1=  (select (sum(amount_cr))* " & CDbl(MD_Rate) & " from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)='" & Month(ds) & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='25'"
            'CNN.Execute(PP)


            If RM.Checked = True Then
                Dim PP As String = "update AP_Rpt_Amt_Status set Amt1=  (select (sum(amount_cr))* " & CDbl(MD_Rate) & " from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)='" & Month(ds) & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='25'"
                CNN.Execute(PP)
            Else

                Dim PP As String = "update AP_Rpt_Amt_Status set Amt1=  (select (sum(amount_cr)* " & CDbl(MD_Rate) & " - sum(amount_dr)* " & CDbl(MD_Rate) & ") from Gen_jn where curr='USD' and  (Ac_Code Like '310%' )  And  Date_Work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  )  where rpt_id='25'"
                CNN.Execute(PP)
                Dim PP1 As String = "update AP_Rpt_Amt_Status set Amt1= Amt1+ (select  isnull(sum(amount_cr)-sum(amount_dr) ,0)  from Gen_jn where  curr='LAK' and  (Ac_Code Like '310%' )  And  Date_Work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  )  where rpt_id='25'"
                CNN.Execute(PP1)
            End If


            '             - ກຳໄລ/ຂາດທຶນໃນປີ   
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '390%' )  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='26'")
            If RY.Checked = True Then
                CNN.Execute("update AP_Rpt_Amt_Status set Amt4=   ( select sum(Amt_cr-Amt_dr) from gen_jn where Year(Date_Work)='" & Year(ds) & "' and (left(ac_code,1)='4' or left(ac_code,1)='5' ))   where rpt_id='26'")
            Else
                'CNN.Execute(" update AP_Rpt_Amt_Status set Amt4=   (select sum(Amt_dr-Amt_cr) from Ap_balance_6 where (Ac_Code Like '390%'))   where rpt_id='26'")
                Dim yearamt As String
                yearamt = " update AP_Rpt_Amt_Status set Amt4=   ( select sum(amount_cr)-sum(amount_dr) from gen_jn   where (left(ac_code,1)='4' or left(ac_code,1)='5' )  and month(date_work)='" & Month(MdStartDate) & "'  and year(date_work)='" & Year(MdStartDate) & "'  " & MULook2 & " )   where rpt_id='26' "
                CNN.Execute(yearamt) 
           End If


            '               ຄັງສະສົມອື່ນໆ 
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '340%' )  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='27'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt1=(select sum(Amt1) from AP_Rpt_Amt_Status where rpt_id>15 and  rpt_id<28)  where rpt_id='28'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt2=(select sum(Amt2) from AP_Rpt_Amt_Status where rpt_id>15 and  rpt_id<28)  where rpt_id='28'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt3=(select sum(Amt3) from AP_Rpt_Amt_Status where rpt_id>15 and  rpt_id<28)  where rpt_id='28'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt4=(select sum(Amt4) from AP_Rpt_Amt_Status where rpt_id>15 and  rpt_id<28)  where rpt_id='28'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt5=(select sum(Amt5) from AP_Rpt_Amt_Status where rpt_id>15 and  rpt_id<28)  where rpt_id='28'")



            'CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '3202000%'  And  Year(Date_Work)='" & Year(ds) & "')  where rpt_id='15'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '3202000%'  And  Year(Date_Work)='" & Year(ds) & "') * -1 where rpt_id='15'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '3203%'  And Ac_Code <> '3202%' And  Year(Date_Work)='" & Year(ds) & "') where rpt_id='15'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum( Amt_cr) from Gen_jn where Ac_Code Like '3203%'  And Ac_Code <> '3202%' And  Year(Date_Work)='" & Year(ds) & "') * -1  where rpt_id='16'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum( Amt_cr) from Gen_jn where Ac_Code Like '3108210%'  And Ac_Code <> '3202%' And  Year(Date_Work)='" & Year(ds) & "') where rpt_id='17'")

            ''CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Dr) from Gen_jn where Ac_Code Like '3901000%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='18'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Dr) from Gen_jn where Ac_Code Like '00.3901000%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='25'")



            '    'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '3108210%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")
            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '310%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")

            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '3202%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")
            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '320%'  And Ac_Code<> '3202' And Year(Date_Work)=" & Year(ds) & ")  where rpt_id='01'")
            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '3810000%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")

            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '23626%'  And Ac_Code Like '23636%' And  Year(Date_Work)='" & Year(MdStartDate) & "')  where rpt_id='03'")
            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '3202000%'  And  Year(Date_Work)='" & Year(MdStartDate) & "')  where rpt_id='04'")
            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '3202000%' And  Year(Date_Work)='" & Year(MdStartDate) & "') * -1 where rpt_id='04'")

            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '3203%'  And Ac_Code <> '3202%' And  Year(Date_Work)='" & Year(MdStartDate) & "')  where rpt_id='05'")
            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum( Amt_cr) from Gen_jn where Ac_Code Like '3203%'  And Ac_Code <> '3202%' And Year(Date_Work)='" & Year(MdStartDate) & "') * -1  where rpt_id='05'")

            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum( Amt_cr) from Gen_jn where Ac_Code Like '3108210%'  And Ac_Code <> '3202%' And Year(Date_Work)='" & Year(MdStartDate) & "') * -1  where rpt_id='06'")
            '    'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Cr) from Open_jn where Ac_Code Like '3108210%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='07'")

            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Dr) from Gen_jn where Ac_Code Like '3908000%' And Year(Date_Work)='" & Year(MdStartDate) & "')  where rpt_id='07'")

            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status  where rpt_id='01' Or rpt_id='06' )  where rpt_id='08'")
            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status  where rpt_id='01' Or rpt_id='03' Or rpt_id='04' Or rpt_id='05' Or rpt_id='07' )  where rpt_id='08'")

            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status  where rpt_id='08' )  where rpt_id='10'")
            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status  where rpt_id='08' )  where rpt_id='10'")

            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status  where rpt_id='10' Or  rpt_id='11' )  where rpt_id='12'")
            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status  where rpt_id='10'  Or  rpt_id='11' )  where rpt_id='12'")
            '    '============
            '    ds = DateAdd(DateInterval.Year, 0, MdStartDate)
            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '3108210%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='12'")
            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '3202%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='12'")
            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '320%'  And Ac_Code<> '3202' And Year(Date_Work)=" & Year(ds) & ")  where rpt_id='12'")
            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '3810000%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='12'")

            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '23626%'   And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='14'")
            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '3202000%'  And  Year(Date_Work)='" & Year(ds) & "')  where rpt_id='15'")
            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '3202000%'  And  Year(Date_Work)='" & Year(ds) & "') * -1 where rpt_id='15'")

            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '3203%'  And Ac_Code <> '3202%' And  Year(Date_Work)='" & Year(ds) & "') where rpt_id='15'")
            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum( Amt_cr) from Gen_jn where Ac_Code Like '3203%'  And Ac_Code <> '3202%' And  Year(Date_Work)='" & Year(ds) & "') * -1  where rpt_id='16'")

            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum( Amt_cr) from Gen_jn where Ac_Code Like '3108210%'  And Ac_Code <> '3202%' And  Year(Date_Work)='" & Year(ds) & "') where rpt_id='17'")

            '    'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Dr) from Gen_jn where Ac_Code Like '3901000%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='18'")
            '    CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Dr) from Gen_jn where Ac_Code Like '00.3901000%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='25'")
        Else
            CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select isnull(sum(Amt_Cr+ Amt_Dr),0) from Open_jn where Ac_Code Like '00.310%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select isnull(sum(Amt_Cr+ Amt_Dr),0) from Open_jn where Ac_Code Like '00.3202%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select isnull(sum(Amt_Cr+ Amt_Dr),0) from Open_jn where (Ac_Code Like '00.3203%' or Ac_Code Like '00.3204%' or Ac_Code Like '00.3205%' or Ac_Code Like '00.3208%' or Ac_Code Like '00.350%' )  And Year(Date_Work)=" & Year(ds) & ")  where rpt_id='01'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select isnull(sum(Amt_Cr+ Amt_Dr),0) from Open_jn where (Ac_Code Like '00.380%' or Ac_Code Like '00.3908%' or Ac_Code Like '00.3901%' ) And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select isnull(sum(Amt_Cr+ Amt_Dr),0) from Open_jn where Ac_Code Like '00.3201%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")
            ' - ສ່ວນຜິດດ່ຽງຈາກການຕີມູນຄ່າຊັບສິນ ++++
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '00.3201%' or Ac_Code Like '00.3206%')  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='03'")
            '    - ທຶນຊ່ວຍໜູນແລະ ທຶນທີ່ໄດ້ຮັບຈັດສັນ
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '00.3701%' or Ac_Code Like '00.3702%')  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='03'")
            '       - ໂອນທຶນຊ່ວຍໜູນກໍ່ສ້າງພື້ນຖານເຂົ້າບັນຊີລາຍຮັບ
            CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '00.37012%' )   And Month(Date_Work)='" & Month(ds) & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='05'")
            '         - ຊຳລະຄືນທຶນທີ່ໄດ້ຮັບຈັດສັນ 
            CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '00.3702%' )   And Month(Date_Work)='" & Month(ds) & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='06'")
            '           - ໂອນກຳໄລສຸດທິເຂົ້າຄັງສຳຮອງຕາມກົດໝາຍ 
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '00.3202%' )  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='23'")
            '            - ໂອນກຳໄລສຸດທິເຂົ້າຄັງສຳຮອງທົ່ວໄປ 
            CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '00.3203%' or Ac_Code Like '00.3204%' or Ac_Code Like '00.3205%' or Ac_Code Like '00.3208%' or Ac_Code Like '00.350%')   And Month(Date_Work)='" & Month(ds) & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='09'")
            '            - ໂອນກຳໄລສຸດທິເຂົ້າຄັງສຳຮອງທົ່ວໄປ 
            CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '00.310%' )  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='10'")
            '             - ກຳໄລ/ຂາດທຶນໃນປີ     380+3908+3901+5-4
            CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_cr) from Open_jn where (Ac_Code Like '00.380%' or Ac_Code Like '00.3908%' or Ac_Code Like '00.3901%' )  And Year(Date_Work)='" & Year(ds) & "'  )  where rpt_id='11'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4= Amt4+ (select isnull(sum(Amt_cr-Amt_dr),0) from Gen_jn where (Ac_Code Like '00.380%' or Ac_Code Like '00.3908%' or Ac_Code Like '00.3901%' )  And Date_Work<'" & (ds) & "'   And Year(Date_Work)='" & Year(ds) & "' )  where rpt_id='11'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt4= Amt4+ (select isnull(sum(Amt_cr-Amt_dr),0) from Gen_jn where (Ac_Code Like '00.380%' or Ac_Code Like '00.3908%' or Ac_Code Like '00.3901%' )  And Month(Date_Work)='" & Month(ds) & "'   And Year(Date_Work)='" & Year(ds) & "'  )  where rpt_id='11'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4= Amt4+ (select sum(Amt_cr-Amt_dr) from Gen_jn where left(Ac_Code,4)='00.5'  And Year(Date_Work)='" & Year(ds) & "'  And Month(Date_Work)='" & Month(ds) & "'  )  where rpt_id='11'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4= Amt4- (select sum(Amt_dr-Amt_cr) from Gen_jn where left(Ac_Code,4)='00.4'  And Year(Date_Work)='" & Year(ds) & "'  And Month(Date_Work)='" & Month(ds) & "'  )  where rpt_id='11'")

            '               ຄັງສະສົມອື່ນໆ  
            CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '00.340%' )  And Month(Date_Work)='" & Month(ds) & "'  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='12'")

            '               4.- ຍອດເຫຼືອທ້າຍປີ N ( ປີນີ້ )
            CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select  isnull(sum(Amt_cr-Amt_dr),0)  from Gen_jn where (Ac_Code Like '00.310%' )   And Month(Date_Work)='" & Month(ds) & "'  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='13'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select  isnull(sum(Amt_cr-Amt_dr),0) from Gen_jn where (Ac_Code Like '00.3202%' )   And Month(Date_Work)='" & Month(ds) & "'  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='13'")
            '===3203+3204+3205+3208+350   380+3908+3901+5-4
            CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select  isnull(sum(Amt_cr-Amt_dr),0)  from Gen_jn where (Ac_Code Like '00.3203%' or Ac_Code Like '00.3204%' or Ac_Code Like '00.3205%' or Ac_Code Like '00.3208%' or Ac_Code Like '00.350%')  And Month(Date_Work)='" & Month(ds) & "'  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='13'")

            '             - ກຳໄລ/ຂາດທຶນໃນປີ     380+3908+3901+5-4
            CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select isnull(sum(Amt_cr),0) from Open_jn where (Ac_Code Like '00.380%' or Ac_Code Like '00.3908%' or Ac_Code Like '00.3901%' )  And Year(Date_Work)='" & Year(ds) & "'  )  where rpt_id='13'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt4= Amt4+ (select  isnull(sum(Amt_cr-Amt_dr),0) from Gen_jn where (Ac_Code Like '00.380%' or Ac_Code Like '00.3908%' or Ac_Code Like '00.3901%' ) And Month(Date_Work)='" & Month(ds) & "'   And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='13'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt4= Amt4+ (select isnull(sum(Amt_cr-Amt_dr),0) from Gen_jn where left(Ac_Code,4)='00.5'  And Month(Date_Work)='" & Month(ds) & "'  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='13'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt4= Amt4- (select isnull(sum(Amt_dr-Amt_cr),0) from Gen_jn where left(Ac_Code,4)='00.4'  And Month(Date_Work)='" & Month(ds) & "'   And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='13'")

            CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_cr+Amt_dr) from Gen_jn where (Ac_Code Like '00.3201%' or Ac_Code Like '00.3206%' or Ac_Code Like '00.37011%' )   And Month(Date_Work)='" & Month(ds) & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='13'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt5= Amt5- (select sum(Amt_cr+Amt_dr) from Gen_jn where (Ac_Code Like '00.37012%' or Ac_Code Like '00.3702%' or Ac_Code Like '00.340%'or Ac_Code Like '00.360%' )   And Month(Date_Work)='" & Month(ds) & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='13'")


            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Cr) from Open_jn where Ac_Code Like '3108210%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='07'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Dr) from Gen_jn where Ac_Code Like '00.3908000%' And Year(Date_Work)='" & Year(MdStartDate) & "')  where rpt_id='07'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status  where rpt_id='01' Or rpt_id='06' )  where rpt_id='08'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status  where rpt_id='01' Or rpt_id='03' Or rpt_id='04' Or rpt_id='05' Or rpt_id='07' )  where rpt_id='08'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status  where rpt_id='08' )  where rpt_id='10'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status  where rpt_id='08' )  where rpt_id='10'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status  where rpt_id='10' Or  rpt_id='11' )  where rpt_id='12'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status  where rpt_id='10'  Or  rpt_id='11' )  where rpt_id='12'")
            '============
            ds = DateAdd(DateInterval.Year, 0, MdStartDate)
            CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select isnull(sum(Amt_Cr+ Amt_Dr),0) from Open_jn where Ac_Code Like '00.310%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='16'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select isnull(sum(Amt_Cr+ Amt_Dr),0) from Open_jn where Ac_Code Like '00.3202%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='16'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select isnull(sum(Amt_Cr+ Amt_Dr),0) from Open_jn where (Ac_Code Like '00.3203%' or Ac_Code Like '00.3204%' or Ac_Code Like '00.3205%' or Ac_Code Like '00.3208%' or Ac_Code Like '00.350%' )  And Year(Date_Work)=" & Year(ds) & ")  where rpt_id='16'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select isnull(sum(Amt_Cr+ Amt_Dr),0) from Open_jn where (Ac_Code Like '00.380%' or Ac_Code Like '00.3908%' or Ac_Code Like '00.3901%' ) And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='16'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select isnull(sum(Amt_Cr+ Amt_Dr),0) from Open_jn where Ac_Code Like '00.3201%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='16'")
            ' - ສ່ວນຜິດດ່ຽງຈາກການຕີມູນຄ່າຊັບສິນ ++++
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '00.3201%' or Ac_Code Like '00.3206%')  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='18'")
            '    - ທຶນຊ່ວຍໜູນແລະ ທຶນທີ່ໄດ້ຮັບຈັດສັນ
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '00.3701%' or Ac_Code Like '00.3702%')  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='19'")
            '       - ໂອນທຶນຊ່ວຍໜູນກໍ່ສ້າງພື້ນຖານເຂົ້າບັນຊີລາຍຮັບ
            CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '00.37012%' )   And Month(Date_Work)='" & Month(ds) & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='20'")
            '         - ຊຳລະຄືນທຶນທີ່ໄດ້ຮັບຈັດສັນ 
            CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '00.3702%' )   And Month(Date_Work)='" & Month(ds) & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='21'")
            '           - ໂອນກຳໄລສຸດທິເຂົ້າຄັງສຳຮອງຕາມກົດໝາຍ 
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '00.3202%' )  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='23'")
            '            - ໂອນກຳໄລສຸດທິເຂົ້າຄັງສຳຮອງທົ່ວໄປ 
            CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '00.3203%' or Ac_Code Like '00.3204%' or Ac_Code Like '00.3205%' or Ac_Code Like '00.3208%' or Ac_Code Like '00.350%')   And Month(Date_Work)='" & Month(ds) & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='24'")
            '            - ໂອນກຳໄລສຸດທິເຂົ້າຄັງສຳຮອງທົ່ວໄປ 
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '00.310%' )  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='25'")
            '             - ກຳໄລ/ຂາດທຶນໃນປີ     380+3908+3901+5-4
            CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_cr) from Open_jn where (Ac_Code Like '00.380%' or Ac_Code Like '00.3908%' or Ac_Code Like '00.3901%' )  And Year(Date_Work)='" & Year(ds) & "'  )  where rpt_id='26'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4= Amt4+ (select isnull(sum(Amt_cr-Amt_dr),0) from Gen_jn where (Ac_Code Like '00.380%' or Ac_Code Like '00.3908%' or Ac_Code Like '00.3901%' )  And Date_Work<'" & (ds) & "'   And Year(Date_Work)='" & Year(ds) & "' )  where rpt_id='26'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt4= Amt4+ (select isnull(sum(Amt_cr-Amt_dr),0) from Gen_jn where (Ac_Code Like '00.380%' or Ac_Code Like '00.3908%' or Ac_Code Like '00.3901%' )  And Month(Date_Work)='" & Month(ds) & "'   And Year(Date_Work)='" & Year(ds) & "'  )  where rpt_id='26'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4= Amt4+ (select sum(Amt_cr-Amt_dr) from Gen_jn where left(Ac_Code,4)='00.5'  And Year(Date_Work)='" & Year(ds) & "'  And Month(Date_Work)='" & Month(ds) & "'  )  where rpt_id='26'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4= Amt4- (select sum(Amt_dr-Amt_cr) from Gen_jn where left(Ac_Code,4)='00.4'  And Year(Date_Work)='" & Year(ds) & "'  And Month(Date_Work)='" & Month(ds) & "'  )  where rpt_id='26'")

            '               ຄັງສະສົມອື່ນໆ  
            CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '00.340%' )  And Month(Date_Work)='" & Month(ds) & "'  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='27'")

            '               4.- ຍອດເຫຼືອທ້າຍປີ N ( ປີນີ້ )
            CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select  isnull(sum(Amt_cr-Amt_dr),0)  from Gen_jn where (Ac_Code Like '00.310%' )   And Month(Date_Work)='" & Month(ds) & "'  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='28'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select  isnull(sum(Amt_cr-Amt_dr),0) from Gen_jn where (Ac_Code Like '00.3202%' )   And Month(Date_Work)='" & Month(ds) & "'  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='28'")
            '===3203+3204+3205+3208+350   380+3908+3901+5-4
            CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select  isnull(sum(Amt_cr-Amt_dr),0)  from Gen_jn where (Ac_Code Like '00.3203%' or Ac_Code Like '00.3204%' or Ac_Code Like '00.3205%' or Ac_Code Like '00.3208%' or Ac_Code Like '00.350%')  And Month(Date_Work)='" & Month(ds) & "'  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='28'")

            '             - ກຳໄລ/ຂາດທຶນໃນປີ     380+3908+3901+5-4
            CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select isnull(sum(Amt_cr),0) from Open_jn where (Ac_Code Like '00.380%' or Ac_Code Like '00.3908%' or Ac_Code Like '00.3901%' )  And Year(Date_Work)='" & Year(ds) & "'  )  where rpt_id='28'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt4= Amt4+ (select  isnull(sum(Amt_cr-Amt_dr),0) from Gen_jn where (Ac_Code Like '00.380%' or Ac_Code Like '00.3908%' or Ac_Code Like '00.3901%' ) And Month(Date_Work)='" & Month(ds) & "'   And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='28'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt4= Amt4+ (select isnull(sum(Amt_cr-Amt_dr),0) from Gen_jn where left(Ac_Code,4)='00.5'  And Month(Date_Work)='" & Month(ds) & "'  And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='28'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt4= Amt4- (select isnull(sum(Amt_dr-Amt_cr),0) from Gen_jn where left(Ac_Code,4)='00.4'  And Month(Date_Work)='" & Month(ds) & "'   And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='28'")

            CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt_cr+Amt_dr) from Gen_jn where (Ac_Code Like '00.3201%' or Ac_Code Like '00.3206%' or Ac_Code Like '00.37011%' )   And Month(Date_Work)='" & Month(ds) & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='28'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt5= Amt5- (select sum(Amt_cr+Amt_dr) from Gen_jn where (Ac_Code Like '00.37012%' or Ac_Code Like '00.3702%' or Ac_Code Like '00.340%'or Ac_Code Like '00.360%' )   And Month(Date_Work)='" & Month(ds) & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='28'")


            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '00.3108210%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '00.3202%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '00.320%'  And Ac_Code<> '00.3202' And Year(Date_Work)=" & Year(ds) & ")  where rpt_id='01'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '00.3810000%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '00.23626%'  And Ac_Code Like '00.23636%' And  Year(Date_Work)='" & Year(ds) & "')  where rpt_id='03'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '00.3202000%'  And  Year(Date_Work)='" & Year(ds) & "')  where rpt_id='04'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '00.3202000%' And  Year(Date_Work)='" & Year(ds) & "') * -1 where rpt_id='04'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '00.3203%'  And Ac_Code <> '00.3202%' And  Year(Date_Work)='" & Year(ds) & "')  where rpt_id='05'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum( Amt_cr) from Gen_jn where Ac_Code Like '00.3203%'  And Ac_Code <> '00.3202%' And Year(Date_Work)='" & Year(ds) & "') * -1  where rpt_id='05'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum( Amt_cr) from Gen_jn where Ac_Code Like '00.3108210%'  And Ac_Code <> '00.3202%' And Year(Date_Work)='" & Year(ds) & "') * -1  where rpt_id='06'")
            ''CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Cr) from Open_jn where Ac_Code Like '3108210%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='07'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Dr) from Gen_jn where Ac_Code Like '00.3908000%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='07'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status  where rpt_id='01' Or rpt_id='06' )  where rpt_id='08'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status  where rpt_id='01' Or rpt_id='03' Or rpt_id='04' Or rpt_id='05' Or rpt_id='07' )  where rpt_id='08'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status  where rpt_id='08' )  where rpt_id='10'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status  where rpt_id='08' )  where rpt_id='10'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status  where rpt_id='10' Or  rpt_id='11' )  where rpt_id='12'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status  where rpt_id='10'  Or  rpt_id='11' )  where rpt_id='12'")
            ''============
            'ds = DateAdd(DateInterval.Year, 0, MdStartDate)
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '00.3108210%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='12'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '00.3202%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='12'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '00.320%'  And Ac_Code<> '3202' And Year(Date_Work)=" & Year(ds) & ")  where rpt_id='12'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '00.3810000%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='12'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '00.23626%'   And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='14'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '00.3202000%'  And  Year(Date_Work)='" & Year(ds) & "')  where rpt_id='15'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '00.3202000%'  And  Year(Date_Work)='" & Year(ds) & "') * -1 where rpt_id='15'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt_cr) from Gen_jn where Ac_Code Like '00.3203%'  And Ac_Code <> '00.3202%' And  Year(Date_Work)='" & Year(ds) & "') where rpt_id='15'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum( Amt_cr) from Gen_jn where Ac_Code Like '00.3203%'  And Ac_Code <> '00.3202%' And  Year(Date_Work)='" & Year(ds) & "') * -1  where rpt_id='16'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum( Amt_cr) from Gen_jn where Ac_Code Like '00.3108210%'  And Ac_Code <> '00.3202%' And  Year(Date_Work)='" & Year(ds) & "') where rpt_id='17'")

            ''CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Dr) from Gen_jn where Ac_Code Like '3901000%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='18'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Dr) from Gen_jn where Ac_Code Like '00.3901000%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='25'")

        End If
        'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status  where rpt_id='12' Or rpt_id='17' )  where rpt_id='19'")
        'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status  where rpt_id='16' )  where rpt_id='28'")
        'CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt2) from AP_Rpt_Amt_Status  where rpt_id='12' Or rpt_id='15' )  where rpt_id='19'")
        'CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt3) from AP_Rpt_Amt_Status  where rpt_id='12' Or rpt_id='16'  )  where rpt_id='19'")
        'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status  where rpt_id='12' Or rpt_id='14'  Or rpt_id='15' Or rpt_id='15' Or rpt_id='16' Or rpt_id='18')  where rpt_id='19'")


        CNN.Execute("update AP_Rpt_Amt_Status set Amt1=0 where amt1 Is null")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt2=0 where amt2 Is null")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt3=0 where amt3 Is null")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt4=0 where amt4 Is null")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt5=0 where amt5 Is null")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt6=0 where amt6 Is null")

        CNN.Execute("DELETE from TEST_MM where month(MM)='" & Month(MdStartDate) & "' and year(MM)='" & Year(MdStartDate) & "'  and rpt_id='01' ")
        CNN.Execute("INSERT INTO TEST_MM(Rpt_ID,Amt,MM) select '01',Amt1,'" & Format(CDate(MdStartDate), "yyyy-MM-dd") & "' from AP_Rpt_Amt_Status where rpt_id='16' ")

        CNN.Execute("DELETE from AP_Rpt_Amt_Status_MM where month(MM)='" & Month(MdStartDate) & "' and year(MM)='" & Year(MdStartDate) & "'   ")
        Dim KKa As String = " INSERT INTO AP_Rpt_Amt_Status_MM(MM,Rpt_ID, Grp, Grp_Nme, Description, Descriptione, Chart_of_Accounts_Codes, Oject, RPT_Type, Fnt, Clor, x, Udln, Lck, clt_Str, clt_Amt, NL, Amt1, Amt2, Amt3, Amt4, Amt5, Amt6, Remark,  Ac_Code, RemShow) " & _
        " select  '" & Format(CDate(MdStartDate), "yyyy-MM-dd") & "',Rpt_ID, Grp, Grp_Nme, Description, Descriptione, Chart_of_Accounts_Codes, Oject, RPT_Type, Fnt, Clor, x, Udln, Lck, clt_Str, clt_Amt, NL, Amt1, Amt2, Amt3, Amt4, Amt5, Amt6, Remark,  Ac_Code, RemShow  from AP_Rpt_Amt_Status "
        CNN.Execute(KKa)
        If RY.Checked = True Then

            CNN.Execute("DELETE from AP_Rpt_Amt_Status_YY where year(MM)='" & Year(MdStartDate) & "'   ")
            Dim KKay As String = " INSERT INTO AP_Rpt_Amt_Status_YY(MM,Rpt_ID, Grp, Grp_Nme, Description, Descriptione, Chart_of_Accounts_Codes, Oject, RPT_Type, Fnt, Clor, x, Udln, Lck, clt_Str, clt_Amt, NL, Amt1, Amt2, Amt3, Amt4, Amt5, Amt6, Remark,  Ac_Code, RemShow) " & _
            " select  '" & Format(CDate(MdStartDate), "yyyy-MM-dd") & "',Rpt_ID, Grp, Grp_Nme, Description, Descriptione, Chart_of_Accounts_Codes, Oject, RPT_Type, Fnt, Clor, x, Udln, Lck, clt_Str, clt_Amt, NL, Amt1, Amt2, Amt3, Amt4, Amt5, Amt6, Remark,  Ac_Code, RemShow  from AP_Rpt_Amt_Status "
            CNN.Execute(KKay)
        End If

        If RP.Checked = True Then

            CNN.Execute("DELETE from AP_Rpt_Amt_Status_PP where year(MM)='" & Year(MdStartDate) & "' and PP=N'" & Period.SelectedIndex & "'  ")
            Dim KKay As String = " INSERT INTO AP_Rpt_Amt_Status_PP(PP,MM,Rpt_ID, Grp, Grp_Nme, Description, Descriptione, Chart_of_Accounts_Codes, Oject, RPT_Type, Fnt, Clor, x, Udln, Lck, clt_Str, clt_Amt, NL, Amt1, Amt2, Amt3, Amt4, Amt5, Amt6, Remark,  Ac_Code, RemShow) " & _
            " select '" & Period.SelectedIndex & "', '" & Format(CDate(MdStartDate), "yyyy-MM-dd") & "',Rpt_ID, Grp, Grp_Nme, Description, Descriptione, Chart_of_Accounts_Codes, Oject, RPT_Type, Fnt, Clor, x, Udln, Lck, clt_Str, clt_Amt, NL, Amt1, Amt2, Amt3, Amt4, Amt5, Amt6, Remark,  Ac_Code, RemShow  from AP_Rpt_Amt_Status "
            CNN.Execute(KKay)
        End If

        'If RT.Checked = True Then

        '    CNN.Execute("DELETE from AP_Rpt_Amt_Status_6M where year(MM)='" & Year(MdStartDate) & "' and PP=N'" & Ct.SelectedIndex & "'  ")
        '    Dim KKay As String = " INSERT INTO AP_Rpt_Amt_Status_6M(PP,MM,Rpt_ID, Grp, Grp_Nme, Description, Descriptione, Chart_of_Accounts_Codes, Oject, RPT_Type, Fnt, Clor, x, Udln, Lck, clt_Str, clt_Amt, NL, Amt1, Amt2, Amt3, Amt4, Amt5, Amt6, Remark,  Ac_Code, RemShow) " & _
        '    " select '" & Ct.SelectedIndex & "', '" & Format(CDate(MdStartDate), "yyyy-MM-dd") & "',Rpt_ID, Grp, Grp_Nme, Description, Descriptione, Chart_of_Accounts_Codes, Oject, RPT_Type, Fnt, Clor, x, Udln, Lck, clt_Str, clt_Amt, NL, Amt1, Amt2, Amt3, Amt4, Amt5, Amt6, Remark,  Ac_Code, RemShow  from AP_Rpt_Amt_Status "
        '    CNN.Execute(KKay)
        'End If



        If RY.Checked = True Then
            'CNN.Execute("  update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt1=(select AP_Rpt_Amt_Status_MM.amt1 from  AP_Rpt_Amt_Status_MM where AP_Rpt_Amt_Status_MM.rpt_id='16' and  month(MM)='12' and year(MM)='" & Year(MdStartDate) - 1 & "' ) where    AP_Rpt_Amt_Status.rpt_id='01'  ")
            ''CNN.Execute("update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt4=(select AP_Rpt_Amt_Status_MM.amt4 from  AP_Rpt_Amt_Status_MM where AP_Rpt_Amt_Status_MM.rpt_id='16' and  month(MM)='12' and year(MM)='" & Year(MdStartDate) - 1 & "'  ) where    AP_Rpt_Amt_Status.rpt_id='01'  ")
            'CNN.Execute("update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt5=(select AP_Rpt_Amt_Status_MM.amt5 from  AP_Rpt_Amt_Status_MM where AP_Rpt_Amt_Status_MM.rpt_id='16' and  month(MM)='12' and year(MM)='" & Year(MdStartDate) - 1 & "'  ) where    AP_Rpt_Amt_Status.rpt_id='01'  ")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=   ( select sum(Amt_cr-Amt_dr) from gen_jn where Year(Date_Work)='" & Year(ds) - 1 & "' and (left(ac_code,1)='4' or left(ac_code,1)='5' ))   where rpt_id='11'")
            CNN.Execute("  update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt1=(select AP_Rpt_Amt_Status_YY.amt1 from  AP_Rpt_Amt_Status_YY where AP_Rpt_Amt_Status_YY.rpt_id='16'  and year(MM)='" & Year(MdStartDate) - 1 & "' ) where    AP_Rpt_Amt_Status.rpt_id='01'  ")
            CNN.Execute("update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt4=(select AP_Rpt_Amt_Status_YY.amt4 from  AP_Rpt_Amt_Status_YY where AP_Rpt_Amt_Status_YY.rpt_id='16'  and year(MM)='" & Year(MdStartDate) - 1 & "'  ) where    AP_Rpt_Amt_Status.rpt_id='01'  ")
            CNN.Execute("update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt5=(select AP_Rpt_Amt_Status_YY.amt5 from  AP_Rpt_Amt_Status_YY where AP_Rpt_Amt_Status_YY.rpt_id='16'   and year(MM)='" & Year(MdStartDate) - 1 & "'  ) where    AP_Rpt_Amt_Status.rpt_id='01'  ")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt4=   ( select sum(Amt_cr-Amt_dr) from gen_jn where Year(Date_Work)='" & Year(ds) - 1 & "' and (left(ac_code,1)='4' or left(ac_code,1)='5' ))   where rpt_id='11'")


            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='13'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt2) from AP_Rpt_Amt_Status where    rpt_id<13 )  where rpt_id='13'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt3) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='13'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='13'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt5) from AP_Rpt_Amt_Status where  rpt_id<13)  where rpt_id='13'")

            CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='14'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt2) from AP_Rpt_Amt_Status where    rpt_id<13 )  where rpt_id='14'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt3) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='14'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='14'")
            CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt5) from AP_Rpt_Amt_Status where  rpt_id<13)  where rpt_id='14'")
        ElseIf RP.Checked = True Then
            If Period.SelectedIndex = 0 Then
                CNN.Execute("  update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt1=(select AP_Rpt_Amt_Status_PP.amt1 from  AP_Rpt_Amt_Status_PP where AP_Rpt_Amt_Status_PP.rpt_id='16'  and pp='3' and year(MM)='" & Year(MdStartDate) - 1 & "' ) where    AP_Rpt_Amt_Status.rpt_id='01'  ")
                CNN.Execute("update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt4=(select AP_Rpt_Amt_Status_PP.amt4 from  AP_Rpt_Amt_Status_PP where AP_Rpt_Amt_Status_PP.rpt_id='16'  and pp='3'  and year(MM)='" & Year(MdStartDate) - 1 & "'  ) where    AP_Rpt_Amt_Status.rpt_id='01'  ")
                CNN.Execute("update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt5=(select AP_Rpt_Amt_Status_PP.amt5 from  AP_Rpt_Amt_Status_PP where AP_Rpt_Amt_Status_PP.rpt_id='16'   and pp='3'  and year(MM)='" & Year(MdStartDate) - 1 & "'  ) where    AP_Rpt_Amt_Status.rpt_id='01'  ")
                CNN.Execute("update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt1=(select AP_Rpt_Amt_Status_PP.amt1 from  AP_Rpt_Amt_Status_PP where AP_Rpt_Amt_Status_PP.rpt_id='25'   and pp='3'  and year(MM)='" & Year(MdStartDate) - 1 & "'  ) where    AP_Rpt_Amt_Status.rpt_id='10'  ")

                'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=   ( select sum(Amt_cr-Amt_dr) from gen_jn where Year(Date_Work)='" & Year(ds) - 1 & "' and (left(ac_code,1)='4' or left(ac_code,1)='5' ))   where rpt_id='11'")
                Dim A1, A2 As Date
                A1 = Year(MdStartDate) - 1 & "-10-01"
                A2 = Year(MdStartDate) - 1 & "-12-31"
                Dim PP As String = "update AP_Rpt_Amt_Status set Amt4=   ( select sum(Amt_cr-Amt_dr) from gen_jn where     Date_Work   BETWEEN '" & Format(A1, "yyyy-MM-dd") & "' AND '" & Format(A2, "yyyy-MM-dd") & "'  and (left(ac_code,1)='4' or left(ac_code,1)='5' ))   where rpt_id='11'"
                CNN.Execute(PP)

                CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='14'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt2) from AP_Rpt_Amt_Status where    rpt_id<13 )  where rpt_id='14'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt3) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='14'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='14'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt5) from AP_Rpt_Amt_Status where  rpt_id<13)  where rpt_id='14'")


            End If
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='14'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt2) from AP_Rpt_Amt_Status where    rpt_id<13 )  where rpt_id='14'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt3) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='14'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='14'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt5) from AP_Rpt_Amt_Status where  rpt_id<13)  where rpt_id='14'")
        ElseIf RT.Checked = True Then
            If Ct.SelectedIndex = 0 Then
                CNN.Execute("  update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt1=(select AP_Rpt_Amt_Status_6M.amt1 from  AP_Rpt_Amt_Status_6M where AP_Rpt_Amt_Status_6M.rpt_id='16'  and pp='1' and year(MM)='" & Year(MdStartDate) - 1 & "' ) where    AP_Rpt_Amt_Status.rpt_id='01'  ")
                Dim po As String = "update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt4=(select AP_Rpt_Amt_Status_6M.amt4 from  AP_Rpt_Amt_Status_6M where AP_Rpt_Amt_Status_6M.rpt_id='16'  and pp='1'  and year(MM)='" & Year(MdStartDate) - 1 & "'  ) where    AP_Rpt_Amt_Status.rpt_id='01'  "
                CNN.Execute(po)
                CNN.Execute("update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt5=(select AP_Rpt_Amt_Status_6M.amt5 from  AP_Rpt_Amt_Status_6M where AP_Rpt_Amt_Status_6M.rpt_id='16'   and pp='1'  and year(MM)='" & Year(MdStartDate) - 1 & "'  ) where    AP_Rpt_Amt_Status.rpt_id='01'  ")
                CNN.Execute("update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt1=(select AP_Rpt_Amt_Status_6M.amt1 from  AP_Rpt_Amt_Status_6M where AP_Rpt_Amt_Status_6M.rpt_id='25'   and pp='1'  and year(MM)='" & Year(MdStartDate) - 1 & "'  ) where    AP_Rpt_Amt_Status.rpt_id='10'  ")

                'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=   ( select sum(Amt_cr-Amt_dr) from gen_jn where Year(Date_Work)='" & Year(ds) - 1 & "' and (left(ac_code,1)='4' or left(ac_code,1)='5' ))   where rpt_id='11'")
                Dim A1, A2 As Date
                A1 = Year(MdStartDate) - 1 & "-07-01"
                A2 = Year(MdStartDate) - 1 & "-12-31"
                Dim PP As String = "update AP_Rpt_Amt_Status set Amt4=   ( select sum(Amt_cr-Amt_dr) from gen_jn where     Date_Work   BETWEEN '" & Format(A1, "yyyy-MM-dd") & "' AND '" & Format(A2, "yyyy-MM-dd") & "'  and (left(ac_code,1)='4' or left(ac_code,1)='5' ))   where rpt_id='11'"
                CNN.Execute(PP)

                CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='14'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt2) from AP_Rpt_Amt_Status where    rpt_id<13 )  where rpt_id='14'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt3) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='14'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='14'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt5) from AP_Rpt_Amt_Status where  rpt_id<13)  where rpt_id='14'")

            Else
                CNN.Execute("  update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt1=(select AP_Rpt_Amt_Status_6M.amt1 from  AP_Rpt_Amt_Status_6M where AP_Rpt_Amt_Status_6M.rpt_id='16'  and pp='0' and year(MM)='" & Year(MdStartDate) & "' ) where    AP_Rpt_Amt_Status.rpt_id='01'  ")
                CNN.Execute("update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt4=(select AP_Rpt_Amt_Status_6M.amt4 from  AP_Rpt_Amt_Status_6M where AP_Rpt_Amt_Status_6M.rpt_id='16'  and pp='0'  and year(MM)='" & Year(MdStartDate) & "'  ) where    AP_Rpt_Amt_Status.rpt_id='01'  ")
                CNN.Execute("update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt5=(select AP_Rpt_Amt_Status_6M.amt5 from  AP_Rpt_Amt_Status_6M where AP_Rpt_Amt_Status_6M.rpt_id='16'   and pp='0'  and year(MM)='" & Year(MdStartDate) & "'  ) where    AP_Rpt_Amt_Status.rpt_id='01'  ")
                'CNN.Execute("update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt1=(select AP_Rpt_Amt_Status_6M.amt1 from  AP_Rpt_Amt_Status_6M where AP_Rpt_Amt_Status_6M.rpt_id='25'   and pp='0'  and year(MM)='" & Year(MdStartDate) & "'  ) where    AP_Rpt_Amt_Status.rpt_id='10'  ")

                'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=   ( select sum(Amt_cr-Amt_dr) from gen_jn where Year(Date_Work)='" & Year(ds) - 1 & "' and (left(ac_code,1)='4' or left(ac_code,1)='5' ))   where rpt_id='11'")
                'Dim A1, A2 As Date
                'A1 = Year(MdStartDate) - 1 & "-01-01"
                'A2 = Year(MdStartDate) - 1 & "-06-30"
                'Dim PP As String = "update AP_Rpt_Amt_Status set Amt4=   ( select sum(Amt_cr-Amt_dr) from gen_jn where     Date_Work   BETWEEN '" & Format(A1, "yyyy-MM-dd") & "' AND '" & Format(A2, "yyyy-MM-dd") & "'  and (left(ac_code,1)='4' or left(ac_code,1)='5' ))   where rpt_id='11'"
                'CNN.Execute(PP)

                CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status where   rpt_id<12 )  where rpt_id='13'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt2) from AP_Rpt_Amt_Status where    rpt_id<12 )  where rpt_id='13'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt3) from AP_Rpt_Amt_Status where   rpt_id<12 )  where rpt_id='13'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status where   rpt_id<12 )  where rpt_id='13'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt5) from AP_Rpt_Amt_Status where  rpt_id<12)  where rpt_id='13'")

                CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt2) from AP_Rpt_Amt_Status where    rpt_id<13 )  where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt3) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt5) from AP_Rpt_Amt_Status where  rpt_id<13)  where rpt_id='16'")
            End If

        ElseIf RM.Checked = True Then
            If Month(MdStartDate) = 1 Then
                CNN.Execute("  update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt1=(select AP_Rpt_Amt_Status_MM.amt1 from  AP_Rpt_Amt_Status_MM where AP_Rpt_Amt_Status_MM.rpt_id='16' and  month(MM)='12' and year(MM)='" & Year(MdStartDate) - 1 & "' ) where    AP_Rpt_Amt_Status.rpt_id='01'  ")
                CNN.Execute("update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt4=(select AP_Rpt_Amt_Status_MM.amt4 from  AP_Rpt_Amt_Status_MM where AP_Rpt_Amt_Status_MM.rpt_id='16' and  month(MM)='12' and year(MM)='" & Year(MdStartDate) - 1 & "'  ) where    AP_Rpt_Amt_Status.rpt_id='01'  ")
                CNN.Execute("update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt5=(select AP_Rpt_Amt_Status_MM.amt5 from  AP_Rpt_Amt_Status_MM where AP_Rpt_Amt_Status_MM.rpt_id='16' and  month(MM)='12' and year(MM)='" & Year(MdStartDate) - 1 & "'  ) where    AP_Rpt_Amt_Status.rpt_id='01'  ")
                '==========
                CNN.Execute("  update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt1=(select AP_Rpt_Amt_Status_MM.amt1 from  AP_Rpt_Amt_Status_MM where AP_Rpt_Amt_Status_MM.rpt_id='26' and  month(MM)='12' and year(MM)='" & Year(MdStartDate) - 1 & "' ) where    AP_Rpt_Amt_Status.rpt_id='11'  ")
                CNN.Execute("update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt4=(select AP_Rpt_Amt_Status_MM.amt4 from  AP_Rpt_Amt_Status_MM where AP_Rpt_Amt_Status_MM.rpt_id='26' and  month(MM)='12' and year(MM)='" & Year(MdStartDate) - 1 & "'  ) where    AP_Rpt_Amt_Status.rpt_id='11'  ")
                CNN.Execute("update AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.amt5=(select AP_Rpt_Amt_Status_MM.amt5 from  AP_Rpt_Amt_Status_MM where AP_Rpt_Amt_Status_MM.rpt_id='26' and  month(MM)='12' and year(MM)='" & Year(MdStartDate) - 1 & "'  ) where    AP_Rpt_Amt_Status.rpt_id='11'  ")

                CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='13'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt2) from AP_Rpt_Amt_Status where    rpt_id<13 )  where rpt_id='13'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt3) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='13'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='13'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt5) from AP_Rpt_Amt_Status where  rpt_id<13)  where rpt_id='13'")

                CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='14'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt2) from AP_Rpt_Amt_Status where    rpt_id<13 )  where rpt_id='14'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt3) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='14'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status where   rpt_id<13 )  where rpt_id='14'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt5) from AP_Rpt_Amt_Status where  rpt_id<13)  where rpt_id='14'")
            Else
                CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status where   (rpt_id>12 and rpt_id<16) )  where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt2=  (select sum(Amt2) from AP_Rpt_Amt_Status where    (rpt_id>12 and rpt_id<16) )  where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt3=  (select sum(Amt3) from AP_Rpt_Amt_Status where    (rpt_id>12 and rpt_id<16) )  where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status where    (rpt_id>12 and rpt_id<16) )  where rpt_id='16'")
                CNN.Execute("update AP_Rpt_Amt_Status set Amt5=  (select sum(Amt5) from AP_Rpt_Amt_Status where  (rpt_id>12 and rpt_id<16) )  where rpt_id='16'")
            End If
        End If
        If RT.Checked = True Then

            CNN.Execute("DELETE from AP_Rpt_Amt_Status_6M where year(MM)='" & Year(MdStartDate) & "' and PP=N'" & Ct.SelectedIndex & "'  ")
            Dim KKay As String = " INSERT INTO AP_Rpt_Amt_Status_6M(PP,MM,Rpt_ID, Grp, Grp_Nme, Description, Descriptione, Chart_of_Accounts_Codes, Oject, RPT_Type, Fnt, Clor, x, Udln, Lck, clt_Str, clt_Amt, NL, Amt1, Amt2, Amt3, Amt4, Amt5, Amt6, Remark,  Ac_Code, RemShow) " & _
            " select '" & Ct.SelectedIndex & "', '" & Format(CDate(MdStartDate), "yyyy-MM-dd") & "',Rpt_ID, Grp, Grp_Nme, Description, Descriptione, Chart_of_Accounts_Codes, Oject, RPT_Type, Fnt, Clor, x, Udln, Lck, clt_Str, clt_Amt, NL, Amt1, Amt2, Amt3, Amt4, Amt5, Amt6, Remark,  Ac_Code, RemShow  from AP_Rpt_Amt_Status "
            CNN.Execute(KKay)
        End If

        CNN.Execute("update AP_Rpt_Amt_Status set Amt1=(select sum(Amt1) from AP_Rpt_Amt_Status where rpt_id>15 and  rpt_id<28)  where rpt_id='28'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt2=(select sum(Amt2) from AP_Rpt_Amt_Status where rpt_id>15 and  rpt_id<28)  where rpt_id='28'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt3=(select sum(Amt3) from AP_Rpt_Amt_Status where rpt_id>15 and  rpt_id<28)  where rpt_id='28'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt4=(select sum(Amt4) from AP_Rpt_Amt_Status where rpt_id>15 and  rpt_id<28)  where rpt_id='28'")
        CNN.Execute("update AP_Rpt_Amt_Status set Amt5=(select sum(Amt5) from AP_Rpt_Amt_Status where rpt_id>15 and  rpt_id<28)  where rpt_id='28'")



        If CheckBox1.Checked = False Then
            Call LoadReport()
        Else
            'Call LoadReportItem()
        End If
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
                        'CNN.Execute(" Update  Ap_Rpt_Cashflow set Amt = " & CLT_Str & " , Last_Amt = " & CLT_Last_Str & " where  Rpt_ID =   '" & (RSC1.Fields("Rpt_ID").Value.ToString) & "' ")
                        'MsgBox((RSC1.Fields("Rpt_ID").Value.ToString) & "===> " & CLT_Str)
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

        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳ" & Ct.Text & " ປີ " & Year(MdToDate)
        'L5.Text = MdStartDate & " => " & MdToDate
    End Sub
    Private Sub LoadDay()
        MdStartDate = Format(CDate(Ds.Value), "dd-MM-yyyy")
        MdToDate = Format(CDate(Dt.Value), "dd-MM-yyyy")
        'Lb.Text = "ແຕ່ວັນທີ " & MdStartDate & " ຫາວັນທີ " & MdToDate
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາວັນທີ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        L5.Text = MdStartDate & " => " & MdToDate
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳວັນທີ"
    End Sub
    Private Sub LoadMonth()
        ''---------------------------------
        'If DMonth.Text = "01" Then
        '    MdStartDate = Format(CDate("01/01/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("31/01/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MonthLetter1 = "01"
        'ElseIf DMonth.Text = "02" Then
        '    Dim Day As String
        '    Dim MM As Date
        '    Dim Fromm As Date
        '    MdStartDate = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
        '    Fromm = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
        '    MM = Format(CDate("01/3/" & Year(MdStartDate)), "dd-MM-yyyy")
        '    Day = DateDiff(DateInterval.Day, Fromm, MM)
        '    MdToDate = Format(CDate(Day & "/02" & "/" & Year(MdStartDate)), "dd-MM-yyyy")
        '    MonthLetter1 = "02"
        '    Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
        'ElseIf DMonth.Text = "03" Then
        '    MdStartDate = Format(CDate("01/03/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("31/03/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MonthLetter1 = "03"
        'ElseIf DMonth.Text = "04" Then
        '    MdStartDate = Format(CDate("01/04/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("30/04/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MonthLetter1 = "04"
        'ElseIf DMonth.Text = "05" Then
        '    MdStartDate = Format(CDate("01/05/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("31/05/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MonthLetter1 = "05"
        'ElseIf DMonth.Text = "06" Then
        '    MdStartDate = Format(CDate("01/06/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("30/06/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MonthLetter1 = "06"
        'ElseIf DMonth.Text = "07" Then
        '    MdStartDate = Format(CDate("01/07/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("31/07/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MonthLetter1 = "07"
        'ElseIf DMonth.Text = "08" Then
        '    MdStartDate = Format(CDate("01/08/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("31/08/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MonthLetter1 = "08"
        'ElseIf DMonth.Text = "09" Then
        '    MdStartDate = Format(CDate("01/09/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("30/09/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MonthLetter1 = "09"
        'ElseIf DMonth.Text = "10" Then
        '    MdStartDate = Format(CDate("01/10/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("31/10/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MonthLetter1 = "10"
        'ElseIf DMonth.Text = "11" Then
        '    MdStartDate = Format(CDate("01/11/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("30/11/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MonthLetter1 = "11"
        'ElseIf DMonth.Text = "12" Then
        '    MdStartDate = Format(CDate("01/12/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MdToDate = Format(CDate("31/12/" & Year(Myy.Value)), "dd-MM-yyyy")
        '    MonthLetter1 = "12"
        'End If
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
            ElseIf DMonth.Text = "ກຸມພາ" Then
                Dim Day As String
                Dim MM As Date
                Dim Fromm As Date
                MdStartDate = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
                Fromm = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
                MM = Format(CDate("01/3/" & Year(MdStartDate)), "dd-MM-yyyy")
                Day = DateDiff(DateInterval.Day, Fromm, MM)
                MdToDate = Format(CDate(Day & "/02" & "/" & Year(MdStartDate)), "dd-MM-yyyy")

                MdStartDate_MM = Format(CDate("01/01/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/01/" & Year(Myy.Value)), "dd-MM-yyyy")

                MonthLetter1 = "ກຸມພາ"
                MonthLetter_Last = "ມັງກອນ"
                DMonth.SelectedIndex = 1
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
            ElseIf DMonth.Text = "ມີນາ" Then
                Dim Day As String
                Dim MM As Date
                Dim Fromm As Date

                MdStartDate = Format(CDate("01/03/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/03/" & Year(Myy.Value)), "dd-MM-yyyy")

                Fromm = Format(CDate("01/02/" & Year(MdStartDate)), "dd-MM-yyyy")
                MM = Format(CDate("01/3/" & Year(MdStartDate)), "dd-MM-yyyy")
                Day = DateDiff(DateInterval.Day, Fromm, MM)
                'MdToDate_MM = Format(CDate(Day & "/02" & "/" & Year(MdStartDate)), "dd-MM-yyyy")
                MdStartDate_MM = Format(CDate("01/02/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate(Day & "/02" & "/" & Year(MdStartDate_MM)), "dd-MM-yyyy")

                MonthLetter1 = "ມີນາ"
                MonthLetter_Last = "ກຸມພາ"
                DMonth.SelectedIndex = 2
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "ເມສາ" Then
                MdStartDate = Format(CDate("01/04/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/04/" & Year(Myy.Value)), "dd-MM-yyyy")

                MdStartDate_MM = Format(CDate("01/03/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/03/" & Year(Myy.Value)), "dd-MM-yyyy")

                MonthLetter1 = "ເມສາ"
                MonthLetter_Last = "ມີນາ"
                DMonth.SelectedIndex = 3
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "ພຶດສະພາ" Then
                MdStartDate = Format(CDate("01/05/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/05/" & Year(Myy.Value)), "dd-MM-yyyy")

                MdStartDate_MM = Format(CDate("01/04/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("30/04/" & Year(Myy.Value)), "dd-MM-yyyy")

                MonthLetter1 = "ພຶດສະພາ"
                MonthLetter_Last = "ເມສາ"
                DMonth.SelectedIndex = 4
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)

            ElseIf DMonth.Text = "ມິຖຸນາ" Then
                MdStartDate = Format(CDate("01/06/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/06/" & Year(Myy.Value)), "dd-MM-yyyy")


                MdStartDate_MM = Format(CDate("01/05/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/05/" & Year(Myy.Value)), "dd-MM-yyyy")

                MonthLetter1 = "ມິຖຸນາ"
                MonthLetter_Last = "ພຶດສະພາ"
                DMonth.SelectedIndex = 5
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "ກໍລະກົດ" Then
                MdStartDate = Format(CDate("01/07/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/07/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdStartDate_MM = Format(CDate("01/06/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("30/06/" & Year(Myy.Value)), "dd-MM-yyyy")

                MonthLetter1 = "ກໍລະກົດ"
                MonthLetter_Last = "ມີຖຸນາ"
                DMonth.SelectedIndex = 6
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "ສິງຫາ" Then
                MdStartDate = Format(CDate("01/08/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/08/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdStartDate_MM = Format(CDate("01/07/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/07/" & Year(Myy.Value)), "dd-MM-yyyy")

                MonthLetter1 = "ສິງຫາ"
                MonthLetter_Last = "ກໍລະກົດ"
                DMonth.SelectedIndex = 7
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "ກັນຍາ" Then
                MdStartDate = Format(CDate("01/09/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/09/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdStartDate_MM = Format(CDate("01/08/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/08/" & Year(Myy.Value)), "dd-MM-yyyy")

                MonthLetter1 = "ກັນຍາ"
                MonthLetter_Last = "ສິງຫາ"
                DMonth.SelectedIndex = 8
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "ຕຸລາ" Then
                MdStartDate = Format(CDate("01/10/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/10/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdStartDate_MM = Format(CDate("01/09/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("30/09/" & Year(Myy.Value)), "dd-MM-yyyy")

                MonthLetter1 = "ຕຸລາ"
                MonthLetter_Last = "ກັນຍາ"
                DMonth.SelectedIndex = 9
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "ພະຈິກ" Then
                MdStartDate = Format(CDate("01/11/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/11/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdStartDate_MM = Format(CDate("01/10/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("31/10/" & Year(Myy.Value)), "dd-MM-yyyy")

                MonthLetter1 = "ພະຈິກ"
                MonthLetter_Last = "ຕຸລາ"
                DMonth.SelectedIndex = 10
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "ທັນວາ" Then
                MdStartDate = Format(CDate("01/12/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/12/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdStartDate_MM = Format(CDate("01/11/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate_MM = Format(CDate("30/11/" & Year(Myy.Value)), "dd-MM-yyyy")

                MonthLetter1 = "ທັນວາ"
                MonthLetter_Last = "ພະຈິກ"
                DMonth.SelectedIndex = 11
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            End If

            'Month_Last = Format(dpMonthPrev.Value, "MM/yyyy")
            Month_Last = "[" & MonthLetter_Last & " " & Format(dpMonthPrev.Value, "yyyy") & "]"
            Month_IN = "[" & MonthLetter1 & " " & Format(MdToDate, "yyyy") & "]"
            Lb.Text = "ສຳລັບວັນທີ " & (MdToDate.Day) & " " & MonthLetter1 & " " & Year(MdToDate)
            Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
        Else

            If DMonth.Text = "January" Then
                MdStartDate = Format(CDate("01/01/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/01/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "January"
                MonthLetter_Last = "December"
                DMonth.SelectedIndex = 0
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
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
                MonthLetter_Last = "January"
                DMonth.SelectedIndex = 1
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
                Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
            ElseIf DMonth.Text = "March" Then
                MdStartDate = Format(CDate("01/03/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/03/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "March"
                MonthLetter_Last = "February"
                DMonth.SelectedIndex = 2
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "April" Then
                MdStartDate = Format(CDate("01/04/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/04/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "April"
                MonthLetter_Last = "March"
                MonthLetter_Last = DMonth.SelectedIndex = 3
                DMonth.SelectedIndex = 3
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "May" Then
                MdStartDate = Format(CDate("01/05/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/05/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "May"
                MonthLetter_Last = "April"
                DMonth.SelectedIndex = 4
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "June" Then
                MdStartDate = Format(CDate("01/06/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/06/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "June"
                MonthLetter_Last = "May"
                DMonth.SelectedIndex = 5
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "July" Then
                MdStartDate = Format(CDate("01/07/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/07/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "July"
                MonthLetter_Last = "June"
                DMonth.SelectedIndex = 6
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "August" Then
                MdStartDate = Format(CDate("01/08/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/08/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "August"
                MonthLetter_Last = "July"
                DMonth.SelectedIndex = 7
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "September" Then
                MdStartDate = Format(CDate("01/09/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/09/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "September"
                MonthLetter_Last = "August"
                DMonth.SelectedIndex = 8
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "October" Then
                MdStartDate = Format(CDate("01/10/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/10/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "October"
                MonthLetter_Last = "September"
                DMonth.SelectedIndex = 9
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "November" Then
                MdStartDate = Format(CDate("01/11/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("30/11/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "November"
                MonthLetter_Last = "October"
                DMonth.SelectedIndex = 10
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            ElseIf DMonth.Text = "December" Then
                MdStartDate = Format(CDate("01/12/" & Year(Myy.Value)), "dd-MM-yyyy")
                MdToDate = Format(CDate("31/12/" & Year(Myy.Value)), "dd-MM-yyyy")
                MonthLetter1 = "December"
                MonthLetter_Last = "November"
                DMonth.SelectedIndex = 11
                dpMonthPrev.Value = DateAdd("m", -1, MdToDate)
            End If

            'dpMonthPrev.Value = DateAdd("m", -1, MdToDate)

            Month_Last = "[" & MonthLetter_Last & " " & Format(dpMonthPrev.Value, "yyyy") & "]"
            Month_IN = "[" & MonthLetter1 & " " & Format(MdToDate, "yyyy") & "]"
            Lb.Text = "For the Month Ended " & (MdToDate.Day) & " " & MonthLetter1 & " " & Year(MdToDate)
        End If

        '-----------------
        'Lb.Text = "ປະຈຳເດືອນ " & MonthLetter1 & "/" & Year(MdToDate)
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳເດືອນ " & MonthLetter1 & " ປີ " & Year(MdToDate)
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub

    Private Sub LoadPeriod()
        If Period.Text = "ໄຕມາດ 1" Then
            MdStartDate = Format(CDate("01/01/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/03/" & Year(Pyy.Value)), "dd-MM-yyyy")

            MdStartDate_MM = Format(CDate("01/10/" & Year(Pyy.Value) - 1), "dd-MM-yyyy")
            MdToDate_MM = Format(CDate("31/12/" & Year(Pyy.Value) - 1), "dd-MM-yyyy")

            Lb.Text = "ປະຈຳໄຕມາດ " & "1" & " ປີ " & Pyy.Text
        ElseIf Period.Text = "ໄຕມາດ 2" Then
            MdStartDate = Format(CDate("01/04/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/06/" & Year(Pyy.Value)), "dd-MM-yyyy")

            MdStartDate_MM = Format(CDate("01/01/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate_MM = Format(CDate("31/03/" & Year(Pyy.Value)), "dd-MM-yyyy")


            Lb.Text = "ປະຈຳໄຕມາດ " & "2" & " ປີ " & Pyy.Text
        ElseIf Period.Text = "ໄຕມາດ 3" Then
            MdStartDate = Format(CDate("01/07/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("30/09/" & Year(Pyy.Value)), "dd-MM-yyyy")

            MdStartDate_MM = Format(CDate("01/04/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate_MM = Format(CDate("30/06/" & Year(Pyy.Value)), "dd-MM-yyyy")


            Lb.Text = "ປະຈຳໄຕມາດ " & "3" & " ປີ " & Pyy.Text
        ElseIf Period.Text = "ໄຕມາດ 4" Then
            MdStartDate = Format(CDate("01/10/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/12/" & Year(Pyy.Value)), "dd-MM-yyyy")

            MdStartDate_MM = Format(CDate("01/07/" & Year(Pyy.Value)), "dd-MM-yyyy")
            MdToDate_MM = Format(CDate("30/09/" & Year(Pyy.Value)), "dd-MM-yyyy")


            Lb.Text = "ປະຈຳໄຕມາດ " & "4" & " ປີ " & Pyy.Text
        End If
        L5.Text = MdStartDate & " => " & MdToDate
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
        DTDATE02 = "ປະຈຳ" & Period.Text & " ປີ " & Year(MdToDate)
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub

    Private Sub LoadYear()
        MdStartDate = Format(CDate("01/1/" & Year(yy.Value)), "dd-MM-yyyy")
        MdToDate = Format(CDate("31/12/" & Year(Toyy.Value)), "dd-MM-yyyy")
        Lb.Text = "ປະຈຳປີ " & yy.Text
        'Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd/MM/yyyy") & " ສີ້ນສຸດວັນທີ " & Format(MdToDate, "dd/MM/yyyy")
        Lb.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ສີ້ນສຸດວັນທີ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
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
        'MuLngRpt = ""
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As    RptSjUd ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
        LngId = "7066" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & " " & DTDATE02 & "' As Crl_RptName ,"
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

        'LngId = "7096" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Mon ,"
        LngId = "7097" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Curr ,"

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
        'SLF = "SELECT  " & MuLngRpt & "  *   FROM Ap_Rpt_Amt_Status  "
        Call LoadLoGO()
        If RM.Checked = True Then
            CNN.Execute("UPDATE Ap_Rpt_Amt_Status set Grp=0 ")
        ElseIf RP.Checked = True Then
            CNN.Execute("UPDATE Ap_Rpt_Amt_Status set Grp=2 ")
        ElseIf RT.Checked = True Then
            CNN.Execute("UPDATE Ap_Rpt_Amt_Status set Grp=3 ")
        Else
            CNN.Execute("UPDATE Ap_Rpt_Amt_Status set Grp=1 ")
        End If



        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            Dim s As String = " SELECT " & MuLngRpt & "  * ,  N'" & MuOffDep & "'  as RptSjoff_Dep   FROM Ap_Rpt_Amt_Status Order by   Rpt_Id asc  "
            .Open(s, CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With

        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim Rpt As New CryAmt_Status
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
        myText2.Text = "ຫົວໜ່ວຍ : ກີບ"

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
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & Lb.Text & "' As    RptSjUd ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
        LngId = "7002" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDP ,"
        LngId = "7057" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_RptName ,"
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
        SLF = "SELECT  " & MuLngRpt & "  *   FROM Ap_Rpt_Cashflow_Detail  "
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
    Private Sub M_10()
        Dim ds As Date
        ds = DateAdd(DateInterval.Year, 0, MdStartDate)


        Dim rs As New ADODB.Recordset
        Dim MDRate_Last As String
        MDRate_Last = " and rate_dt<='" & Format(dpMonthPrev.Value, "yyyy-MM-dd") & "'  "
        MDRate_Last = " and month(rate_dt)<='" & Month(dpMonthPrev.Value) - 2 & "' and  year(rate_dt)='" & Year(dpMonthPrev.Value) & "' "


        Call LoadSqlData("select top 1  * from AP_Rate_history where 1=1  " & MDRate_Last & "  and curr='USD'  ORDER BY rate_dt DESC ", rs)
        If rs.RecordCount > 0 Then
            MD_Rate = (rs.Fields("Rate").Value)
        End If


        Dim KKq As String = "update TEST_ABC set Amt=  (select (sum(amount_dr)-sum(amount_cr))* " & CDbl(MD_Rate) & " from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)<='" & Month(ds) - 2 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='15'"
        CNN.Execute(KKq)

        'Dim KK2qq As String = "update TEST_ABC set Amt=Amt+(select (sum(amount_dr)-sum(amount_cr))* " & CDbl(MD_Rate) & "  from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)<='" & Month(ds) - 2 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='15'"
        'CNN.Execute(KK2qq)

        '========== - Currrmmmmm
        MDRate_DT = " and rate_dt<='" & Format(MdToDate, "yyyy-MM-dd") & "'  "
        MDRate_DT = " and month(rate_dt)<='" & Month(MdToDate) - 2 & "' and  year(rate_dt)='" & Year(MdToDate) & "' "

        Call LoadSqlData("select top 1  * from AP_Rate_history where 1=1  " & MDRate_DT & "  and curr='USD'  ORDER BY rate_dt DESC ", rs)
        If rs.RecordCount > 0 Then
            MD_Rate = (rs.Fields("Rate").Value)
        End If
        '========== - ຜົນກະທົບຈາກອັດຕາແລກປ່ຽນເງິນຕາຕ່າງປະເທດ
        Dim KK2w As String = "update TEST_ABC set Amt=Amt-(select (sum(amount_dr)-sum(amount_cr))* " & CDbl(MD_Rate) & "  from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)<='" & Month(ds) - 2 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='15'"
        CNN.Execute(KK2w)

    End Sub
    Private Sub Mprev()
        CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt1=0 ,amt2=0 , Amt3=0 , Amt4=0  ,Amt5=0 ,Amt6=0")
        CNN.Execute("DELETE TEST_ABC ")
        CNN.Execute("INSERT INTO TEST_ABC(Rpt_ID,Name,amt) select Rpt_ID,descriptione,Amt1 from AP_Rpt_Amt_Status ")

        Call M_10()
        Dim ds As Date
        ds = DateAdd(DateInterval.Year, 0, MdStartDate)

        '========== - ຜົນກະທົບຈາກອັດຕາແລກປ່ຽນເງິນຕາຕ່າງປະເທດ
        Dim rs As New ADODB.Recordset
        Dim MDRate_Last As String
        MDRate_Last = " and rate_dt<='" & Format(dpMonthPrev.Value, "yyyy-MM-dd") & "'  "
        MDRate_Last = " and month(rate_dt)<='" & Month(dpMonthPrev.Value) - 2 & "' and  year(rate_dt)='" & Year(dpMonthPrev.Value) & "' "


        Call LoadSqlData("select top 1  * from AP_Rate_history where 1=1  " & MDRate_Last & "  and curr='USD'  ORDER BY rate_dt DESC ", rs)
        If rs.RecordCount > 0 Then
            MD_Rate = (rs.Fields("Rate").Value)
        End If


        Dim KKq As String = "update TEST_ABC set Amt=Amt+  (select (sum(amount_dr)-sum(amount_cr))* " & CDbl(MD_Rate) & " from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)<='" & Month(ds) - 2 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='15'"
        CNN.Execute(KKq)

        'Dim KK2qq As String = "update TEST_ABC set Amt=Amt+(select (sum(amount_dr)-sum(amount_cr))* " & CDbl(MD_Rate) & "  from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)<='" & Month(ds) - 2 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='15'"
        'CNN.Execute(KK2qq)

        '========== - Currrmmmmm
        MDRate_DT = " and rate_dt<='" & Format(MdToDate, "yyyy-MM-dd") & "'  "
        MDRate_DT = " and month(rate_dt)<='" & Month(MdToDate) - 2 & "' and  year(rate_dt)='" & Year(MdToDate) & "' "

        Call LoadSqlData("select top 1  * from AP_Rate_history where 1=1  " & MDRate_DT & "  and curr='USD'  ORDER BY rate_dt DESC ", rs)
        If rs.RecordCount > 0 Then
            MD_Rate = (rs.Fields("Rate").Value)
        End If
        '========== - ຜົນກະທົບຈາກອັດຕາແລກປ່ຽນເງິນຕາຕ່າງປະເທດ
        Dim KK2w As String = "update TEST_ABC set Amt=Amt-(select (sum(amount_dr)-sum(amount_cr))* " & CDbl(MD_Rate) & "  from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)<='" & Month(ds) - 2 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='15'"
        CNN.Execute(KK2w)


        'MsgBox(DisT)

        If Month(MdStartDate) <> 1 Then
            Call MMM()

            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt1=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '310%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt1=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '310%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt4=  (select sum(Amt_Cr- Amt_Dr) from Open_jn where Ac_Code Like '380%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=Amt1+  (select sum(Amt_Cr+ Amt_Dr) from gen_jn where Ac_Code Like '310%' And month(Date_Work)='" & Month(ds) - 1 & "'  And Year(Date_Work)='" & Year(ds) & "' )  where rpt_id='01'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '310%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Cr- Amt_Dr) from Open_jn where Ac_Code Like '380%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")

            If Month(MdStartDate) > 2 Then
                Dim AA As String = "update AP_Rpt_Amt_Status_MM set Amt4= Amt4+ (select sum(Amt_cr)-sum(Amt_dr) from gen_jn where   (left(ac_code,1)='4' or left(ac_code,1)='5')  and Month(Date_Work)<'" & Month(ds) - 2 & "' and Year(Date_Work)='" & Year(ds) & "' )  where rpt_id='01'"
                CNN.Execute(AA)
            End If
            'CNN.Execute("update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_dr)-sum(Amt_cr) from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)<='" & Month(ds) - 1 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")
            ''========== - ຜົນກະທົບຈາກອັດຕາແລກປ່ຽນເງິນຕາຕ່າງປະເທດ
            'Dim KK As String = "update AP_Rpt_Amt_Status set Amt1=Amt1-(select (sum(amount_dr)-sum(amount_cr))*9576  from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)<='" & Month(ds) - 1 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'"
            'CNN.Execute(KK)
            'Dim Pq As String = "update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)='" & Month(ds) - 1 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='10'"
            'CNN.Execute(Pq)

            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt2=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '3202%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")

            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt3=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '320%'  And Ac_Code<> '3202' And Year(Date_Work)=" & Year(ds) & ")  where rpt_id='01'")

            'CNN.Execute("update AP_Rpt_Amt_Status set Amt4=  (select sum(Amt_Cr- Amt_Dr) from Open_jn where Ac_Code Like '390%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'")
            '            -  - ທຶນຈົດທະບຽນ ທີ່ໄດ້ຮັບ
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt1=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)='" & Month(ds) - 2 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='10'")
            'Dim xa As String = "update TEST_ABC set Amt=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)='" & Month(ds) - 2 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='10'"
            'CNN.Execute(xa)

            'Dim Pq As String = "update AP_Rpt_Amt_Status set Amt1=  (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)='" & Month(ds) - 1 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='10'"
            'CNN.Execute(Pq)

            'Dim Pq2 As String = "update AP_Rpt_Amt_Status set Amt1= Amt1+ (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)='" & Month(ds) - 1 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'"
            'CNN.Execute(Pq2)

            'Dim Pq3 As String = "update AP_Rpt_Amt_Status set Amt1= Amt1+ (select sum(Amt_cr) from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)='" & Month(ds) - 1 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='01'"
            'CNN.Execute(Pq3)

            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt4=   (select sum(Amt_dr-Amt_cr) from Ap_balance_6 where (Ac_Code Like '390%'))   where rpt_id='11'")
            '========== - ຜົນກະທົບຈາກອັດຕາແລກປ່ຽນເງິນຕາຕ່າງປະເທດ

            MDRate_Last = " and rate_dt<='" & Format(dpMonthPrev.Value, "yyyy-MM-dd") & "'  "
            MDRate_Last = " and month(rate_dt)<='" & Month(dpMonthPrev.Value) - 1 & "' and  year(rate_dt)='" & Year(dpMonthPrev.Value) & "' "


            Call LoadSqlData("select top 1  * from AP_Rate_history where 1=1  " & MDRate_Last & "  and curr='USD'  ORDER BY rate_dt DESC ", rs)
            If rs.RecordCount > 0 Then
                MD_Rate = (rs.Fields("Rate").Value)
            End If


            Dim KK As String = "update AP_Rpt_Amt_Status_MM set Amt1=  (select (sum(amount_dr)-sum(amount_cr))* " & CDbl(MD_Rate) & " from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)<='" & Month(ds) - 2 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='15'"
            CNN.Execute(KK)

            '========== - Currrmmmmm
            MDRate_DT = " and rate_dt<='" & Format(MdToDate, "yyyy-MM-dd") & "'  "
            Call LoadSqlData("select top 1  * from AP_Rate_history where 1=1  " & MDRate_DT & "  and curr='USD'  ORDER BY rate_dt DESC ", rs)
            If rs.RecordCount > 0 Then
                MD_Rate = (rs.Fields("Rate").Value)
            End If
            '========== - ຜົນກະທົບຈາກອັດຕາແລກປ່ຽນເງິນຕາຕ່າງປະເທດ
            Dim KK2 As String = "update AP_Rpt_Amt_Status_MM set Amt1=Amt1-(select (sum(amount_dr)-sum(amount_cr))* " & CDbl(MD_Rate) & "  from Gen_jn where (Ac_Code Like '3108%' )  And Month(Date_Work)<='" & Month(ds) - 2 & "' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='15'"
            CNN.Execute(KK2)

            If Month(MdStartDate) > 1 Then
                CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt5=  (select sum(Amt_Cr-Amt_Dr) from Open_jn where (Ac_Code Like '330%' or Ac_Code Like '340%'or Ac_Code Like '350%' or Ac_Code Like '360%'or Ac_Code Like '370%' )  And Year(Date_Work)='" & Year(ds) & "' )  where rpt_id='01'")

            End If

            '==============NEW=====
            'CNN.Execute("UPDATE TEST_ABC set amt=0 where amt is null")
            'Dim aq As String = "UPDATE AP_Rpt_Amt_Status set AP_Rpt_Amt_Status.Amt1=AP_Rpt_Amt_Status.Amt1+(select sum(amt) from TEST_ABC where  (rpt_id='10' or rpt_id='15') ) where  rpt_id='01' "
            'CNN.Execute(aq)


            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt1=(select sum(Amt1) from AP_Rpt_Amt_Status where   rpt_id<13)  where rpt_id='13' ")
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt2=(select sum(Amt2) from AP_Rpt_Amt_Status where  rpt_id<13)  where rpt_id='13' ")
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt3=(select sum(Amt3) from AP_Rpt_Amt_Status where   rpt_id<13)  where rpt_id='13' ")
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt4=(select sum(Amt4) from AP_Rpt_Amt_Status where   rpt_id<13)  where rpt_id='13' ")
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt5=(select sum(Amt5) from AP_Rpt_Amt_Status where   rpt_id<13)  where rpt_id='13' ")

        End If
        '============
        MMM22()
        ds = DateAdd(DateInterval.Year, 0, MdStartDate)
        CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt1=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '310%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='16'")
        CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt2=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '3202%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='16'")
        CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt3=  (select sum(Amt_Cr+ Amt_Dr) from Open_jn where Ac_Code Like '320%'  And Ac_Code<> '3202' And Year(Date_Work)=" & Year(ds) & ")  where rpt_id='16'")
        CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt4=  (select sum(Amt_Cr- Amt_Dr) from Open_jn where Ac_Code Like '380%' And Year(Date_Work)='" & Year(ds) & "')  where rpt_id='16'")
        CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt5=  (select sum(Amt_Cr- Amt_Dr) from Open_jn where (Ac_Code Like '330%' or Ac_Code Like '340%'or Ac_Code Like '350%' or Ac_Code Like '360%'or Ac_Code Like '370%' )  And Year(Date_Work)='" & Year(ds) & "' )  where rpt_id='16'")

        If Month(MdStartDate) <> 1 Then
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status_MM where rpt_id='13' )  where rpt_id='16'")
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt2=  (select sum(Amt2) from AP_Rpt_Amt_Status_MM where rpt_id='13' )  where rpt_id='16'")
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt3=  (select sum(Amt3) from AP_Rpt_Amt_Status_MM where rpt_id='13' )  where rpt_id='16'")
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status_MM where rpt_id='13' )  where rpt_id='16'")
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt5=  (select sum(Amt5) from AP_Rpt_Amt_Status_MM where rpt_id='13' )  where rpt_id='16'")

        End If

        If Month(MdStartDate) > 2 Then
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt1= 0    where rpt_id='16'")
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt2= 0 where rpt_id='16'")
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt3= 0 where rpt_id='16'")
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt4=0 where rpt_id='16'")
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt5= 0 where rpt_id='16'")

            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt1=  (select sum(Amt1) from AP_Rpt_Amt_Status_MM where rpt_id>12 and  rpt_id<16 )  where rpt_id='16'")
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt2=  (select sum(Amt2) from AP_Rpt_Amt_Status_MM where rpt_id>12  and  rpt_id<16 )  where rpt_id='16'")
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt3=  (select sum(Amt3) from AP_Rpt_Amt_Status_MM where rpt_id>12  and  rpt_id<16 )  where rpt_id='16'")
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt4=  (select sum(Amt4) from AP_Rpt_Amt_Status_MM where rpt_id>12  and  rpt_id<16 )  where rpt_id='16'")
            CNN.Execute("update AP_Rpt_Amt_Status_MM set Amt5=  (select sum(Amt5) from AP_Rpt_Amt_Status_MM where rpt_id>12  and  rpt_id<16)  where rpt_id='16'")

        End If
    End Sub

    Private Sub RT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RT.CheckedChanged
        selectLoad()
    End Sub

    Private Sub Ct_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ct.SelectedIndexChanged
        selectLoad()

    End Sub

    Private Sub yyt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yyt.ValueChanged
        selectLoad()
    End Sub
End Class