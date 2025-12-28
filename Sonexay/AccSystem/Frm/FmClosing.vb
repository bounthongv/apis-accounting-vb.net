Public Class FmClosing
    Dim mylock As Integer = 0
    Private Sub FmClosing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetControlText(Me)
        FG.FormatString = "^ດ/ລ |ສະກຸນເງິນ|ສະກຸນເງິນ (ຊື່ເຕັມ)                     |ອັດຕາແລກປ່ຽນ |"
        Label1.Text = "ອັດຕາໂດລາ"
        FG.Cols = 4
        LoadListFG()
        loadOffice_User()
    End Sub
    Private Sub loadOffice_User()
        Off_Usr.Items.Clear()
        LoadSqlData("select sub_id , off_add2  from  Ap_Office  Order by sub_id", RSC)
        With RSC
            Do Until .EOF = True
                Off_Usr.Items.Add((.Fields("sub_id").Value) & " " & (.Fields("off_add2").Value))
                .MoveNext()
            Loop
        End With
        Off_Usr.Text = FmLogin.Sub_Company.Text
    End Sub
    Public Sub LoadListFG()
        FG.Rows = 1
        With RSC
            Call LoadSqlData("SELECT * FROM  Ap_RateSeting ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("Curr").Value)) & _
                                 "" & vbTab & Trim(CStr(.Fields("Curr_Name").Value)) & _
                                 "" & vbTab & Trim(Format(CDbl(.Fields("Rate").Value), "##,##0.00")))
                    .MoveNext()
                End While
            Else
                FG.Rows = 2
            End If
            FG.Size = New System.Drawing.Size(457, 160)
        End With
        Call LoadSqlData("SELECT * FROM  Ap_RateSeting where Curr='USD' ", RSC)
        If RSC.RecordCount > 0 Then
            txtRate.Text = Format(CDbl(RSC.Fields("Rate").Value), "##,##0.00")
        Else
            txtRate.Text = 1
        End If
    End Sub
    Private Sub LockAcc()
        Dim RSC1 As New ADODB.Recordset
        With RSC1
            Call LoadSqlData("SELECT Date_work FROM  Open_jn where  My_Lock =1 And  year(date_Work) =  '" & CDbl(yy.Text) & "'   ", RSC1)
            If .RecordCount <> 0 Then
                MsgBox("ບັນຊີປີ '" & CDbl(yy.Text) & "' ໄດ້ລ໋ອດໄວ້ແລ້ວທ່ານຕ້ອງປົດລ໋ອດກ່ອນ")
                mylock = 1
            End If
        End With

    End Sub
    Private Sub BtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelete.Click

        mylock = 0 : LockAcc() : If mylock = 1 Then Exit Sub
        If MessageBox.Show("ທ່ານຕ້ອງການປິດບັນຊີປະຈຳປີ  " & yy.Text & " ແທ້ຫລືບໍ່?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            MdStartDate = "1/1/" & yy.Text
            MdToDate = "31/12/" & yy.Text
            Off_Find = Off_Usr.Text
            Dim OfUsr1 As String = Microsoft.VisualBasic.Left(Off_Find, 5)
            Dim OfUsr2 As String = Mid(Off_Find, 4, 2)
            Dim OfUsr3 As String = Microsoft.VisualBasic.Left(Off_Find, 2)
            CNN.Execute("update Ap_Office set Lck=0 ")
            If OfUsr1 = "00-00" Then
                CNN.Execute("update Ap_Office set Lck=1 where sub_id <> '00-00' And Substring(sub_id,4,2) <> '00' ")
            Else
                If OfUsr2 = "00" Then
                    CNN.Execute("update Ap_Office set Lck=1 where sub_id <> '00-00' And Substring(sub_id,4,2) <> '00' And Substring(sub_id,1,2)= '" & OfUsr3 & "' ")
                Else
                    CNN.Execute("update Ap_Office set Lck=1 where Substring(sub_id,1,5)= '" & OfUsr1 & "' ")
                End If
            End If
            With RSC
                Call LoadSqlData("SELECT sub_id FROM  Ap_Office where  Lck=1 ", RSC)
                If .RecordCount <> 0 Then
                    While Not .EOF

                        Off_Find = Trim(CStr(.Fields("sub_id").Value)) : MuTable = "" : Call Find_Company()
                        ChangInCom = 1
                        New_Code = "3901000.00.0000"
                        Code_Dr = "4"
                        Code_Cr = "5"
                        MuLeftAcCode = 0
                        Call ChangBalance_COLSE()

                        Call ClosingAc_Code()
                        .MoveNext()
                    End While
                End If
            End With
            Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
            If CheckBox1.Checked = True Then
                CNN.Execute("Update Gen_jn set Lock= '2' where  year(date_Work) =  '" & CDbl(yy.Text) & "' " & MULook2 & "")
            Else
                CNN.Execute("Update Gen_jn set Lock= '0' where  year(date_Work) =  '" & CDbl(yy.Text) & "' " & MULook2 & "")
            End If
            CNN.Execute("update  Open_jn set  Rate = '1' where  Curr='LAK'   and year(Open_jn.date_Work) =  '" & CDbl(yy.Text) + 1 & "'  And Open_jn.company='" & Off_Find & "' ")

            CNN.Execute("update  Open_jn set  Rate =  " & CDbl(txtRate.Text) & "  where  Curr='USD'    and year(Open_jn.date_Work) =  '" & CDbl(yy.Text) + 1 & "'  And Open_jn.company='" & Off_Find & "' ")

            CNN.Execute("update Open_jn set Amount_Dr  = Amt_Dr/rate ,   Amount_cr = Amt_cr/rate ,  Last_User = '" & MUserID & "'  where year(Open_jn.date_Work) =  '" & CDbl(yy.Text) + 1 & "'  And Open_jn.company='" & Off_Find & "' ")

            CNN.Execute("update  Open_jn set Curr='LAK' , Rate = '1' , Amount_Dr = Amt_Dr ,Amount_Cr = Amt_Cr  where curr Is null  ")
            CNN.Execute("update  Open_jn set Curr='LAK' , Rate = '1' , Amount_Dr = Amt_Dr ,Amount_Cr = Amt_Cr  where Ac_Code = '" & New_Code & "'  ")

            CNN.Execute("update  Open_jn set ac_name = Acc_Code.Name_L , ac_namee = Acc_Code.Name_E , ac_type = Acc_Code.Acc_Type ,   ac_typee = Acc_Code.Acc_Typee from Open_jn , Acc_Code where Open_jn.Ac_Code = Acc_Code.Ac_Code  ")
            CNN.Execute("  Update Open_jn set amount_cr =0 where amount_cr is null Update Open_jn set amount_Dr =0 where amount_Dr is null")
            MessageBox.Show("ບັນຊີປະຈຳປີ " & yy.Text & " ຖືກປິດຮຽບຮ້ຽຍແລ້ວ", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
    Public Sub ChangBalance_COLSE()
        CNN.Execute("DELETE  Ap_balance_6_col ")
        CNN.Execute("DELETE FROM Ap_balance_6 ")

 
        Dim GGG As String = "INSERT INTO Ap_balance_6 (curr, ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select curr,ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(Amount_Dr)as amt_dr , sum(Amount_cr)as amt_cr  from gen_jn  WHERE 1=1  and Curr=N'LAK'   and gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY curr,ac_code "
        CNN.Execute(GGG)



        Dim USD As String = "INSERT INTO Ap_balance_6 ( curr,ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
 " select curr,ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amount_Dr)* " & CDbl(txtRate.Text) & "  as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & " as amt_cr  from gen_jn  WHERE 1=1 and Curr=N'USD'  and gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY curr,ac_code "
        CNN.Execute(USD)

        Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)

        '=======LAK===
        Dim PPP As String = "INSERT INTO Ap_balance_6 ( curr,ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
                " select curr,ac_code , sum(Amount_Dr)as amt_dr , sum(Amount_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE 1=1  and Curr=N'LAK'   and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY curr,ac_code"
        CNN.Execute(PPP)
        Dim PPPUSD As String = "INSERT INTO Ap_balance_6 ( curr,ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
" select curr,ac_code , sum(amount_Dr)* " & CDbl(txtRate.Text) & " as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & " as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE 1=1  and Curr=N'USD'  and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY curr,ac_code"
        CNN.Execute(PPPUSD)

        '        '=======LAK===
        CNN.Execute("INSERT INTO Ap_balance_6 ( curr,ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
  " select curr,ac_code  , sum(Amount_Dr) as amt_dr , sum(Amount_cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE  1=1   and Curr=N'LAK'   and date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY curr,ac_code")

 
        CNN.Execute("INSERT INTO Ap_balance_6 ( curr,ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
  " select curr,ac_code  , sum(amount_Dr)* " & CDbl(txtRate.Text) & "  as amt_dr , sum(amount_cr)* " & CDbl(txtRate.Text) & "  as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE  1=1  and Curr=N'USD'  and   date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY curr,ac_code")


        CNN.Execute("INSERT INTO Ap_balance_6_col ( curr,ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        " select curr,ac_code , sum(open_amt_dr)as open_amt_dr , sum(open_amt_cr)as open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from Ap_balance_6   group BY curr,ac_code")
        '=================NEWWWWW==============
        'CNN.Execute("DELETE  Ap_balance_6_col ")
        'CNN.Execute("DELETE FROM Ap_balance_6 ")
        'Dim KKKa As String = " insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  from Ap_balance_TB where  date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'  " & Ac_Code & "  order by Ac_Code asc "
        'CNN.Execute(KKKa)

        Call Left_AcCode()
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr ")
        CNN.Execute("Update  Ap_balance_6_col set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr ")
        Call Chang_Incom()
        CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + amt_dr) - (open_amt_cr + amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr + amt_cr) ")
        CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr + amt_cr) - (open_amt_dr + amt_dr) where (open_amt_cr + amt_cr) >= (open_amt_dr + amt_dr) ")
        If MuLng = "L" Then
            CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        Else
            CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_E from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")
        End If
    End Sub
    Private Sub Chang_Incom()
        If ChangInCom = 1 Then
            '      Insr = "delete  Ap_balance_6  " & _
            '         "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  )   select '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr) from Ap_balance_6_col where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "'  or  Ac_Code =  '" & New_Code & "'  " & _
            '"update  Ap_balance_6 set open_amt_dr = 0 where open_amt_dr  is null  " & _
            '"update  Ap_balance_6 set open_amt_cr = 0 where open_amt_cr  is null  " & _
            '"update  Ap_balance_6 set amt_dr = 0 where amt_dr  is null  " & _
            '"update  Ap_balance_6 set amt_cr = 0 where amt_cr  is null   " & _
            ' "Update  Ap_balance_6 set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr   " & _
            '"Update  Ap_balance_6 set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr " & _
            '"Update  Ap_balance_6 set   amt_dr = amt_dr  - amt_cr , amt_cr=0  where amt_dr  >= amt_cr   " & _
            '"Update  Ap_balance_6 set   amt_cr = amt_cr  - amt_dr , amt_dr=0  where amt_cr  >= amt_dr " & _
            '   "delete  Ap_balance_6_col  where left(Ac_Code,1) ='" & Code_Dr & "' Or left(Ac_Code,1)='" & Code_Cr & "'  or  Ac_Code =  '" & New_Code & "'  " & _
            '     "  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  from Ap_balance_6"
            '      CNN.Execute(Insr)
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
" select  '" & New_Code & "'  , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr),1 from Ap_balance_6_col where Ac_Code =  '" & New_Code & "' " & _
"       delete  Ap_balance_6_col where Ac_Code =  '" & New_Code & "' " & _
"  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , sum(open_amt_dr) , sum(open_amt_cr) , sum(amt_dr) , sum(amt_cr)  from Ap_balance_6 group by Ac_Code "

            CNN.Execute(Insr)
        End If
    End Sub
    Private Sub Left_AcCode()
        Dim L As String
        If MuLeftAcCode > 0 Then
            L = CDbl(MuLeftAcCode) + 2
            Insr = "delete Ap_balance_6 " & _
           "Update Ap_balance_6_col set Acc_Parent = left(Ac_Code," & L & ") " & _
            "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  )  " & _
            " select Acc_Parent ,sum(open_amt_dr) As open_amt_dr ,sum(open_amt_cr) As open_amt_cr  ,sum(amt_dr) As amt_dr ,sum(amt_cr) As amt_cr  from Ap_balance_6_col group by  Acc_Parent " & _
           "  delete Ap_balance_6_col " & _
           "insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  )  " & _
         "select Ac_Code ,sum(open_amt_dr) As open_amt_dr ,sum(open_amt_cr) As open_amt_cr  ,sum(amt_dr) As amt_dr ,sum(amt_cr) As amt_cr  from Ap_balance_6  group by  Ac_Code"
            CNN.Execute(Insr)
        End If
    End Sub
    Private Sub ClosingAc_Code()
        'MsgBox(MUserID & "==" & MuSubOff)
        Dim d As Date = "1/1/" & CDbl(yy.Text) + 1
        CNN.Execute("delete Open_jn where date_Work =  '" & "1/1/" & CDbl(yy.Text) + 1 & "' And company='" & Off_Find & "' ")
        Dim KK As String = "insert into Open_jn (date_work, curr,ac_code , amt_dr , amt_cr  , company  ) " & _
            " select  '" & "1/1/" & CDbl(yy.Text) + 1 & "', curr,ac_code , sum(Rem_dr) , sum(Rem_cr) , '" & Off_Find & "' from   Ap_balance_6_col where    Rem_dr- Rem_Cr <> 0 group by curr,ac_code "
        CNN.Execute(KK)
        CNN.Execute("DELETE Open_jn where date_Work = '" & "1/1/" & CDbl(yy.Text) + 1 & "' And company='" & Off_Find & "' and (round(amt_dr,1)+round(amt_cr,1)) <=0 ")
        'MsgBox(Off_Find)


        CNN.Execute("delete Acc_Close_Rate")
        Dim s As String = "insert into Acc_Close_Rate (Ac_code,Cnt_MT)SELECT  ac_code, (SELECT TOP 1 cnt  FROM gen_jn AS B WHERE (ac_code = A.ac_code) ORDER BY cnt  desc) AS cnt  FROM gen_jn AS A  " & _
        " WHERE year(date_work)=" & yy.Text & " And company='" & Off_Find & "'  GROUP BY ac_code ORDER BY ac_code "
        CNN.Execute(s)

        CNN.Execute("update Acc_Close_Rate set Curr=gen_jn.curr  from  Acc_Close_Rate , gen_jn where Acc_Close_Rate.cnt_Mt=gen_jn.cnt")

        CNN.Execute("update Acc_Close_Rate set rate=Ap_RateSeting.rate from  Acc_Close_Rate , Ap_RateSeting where Acc_Close_Rate.curr=Ap_RateSeting.curr")

        'CNN.Execute("update Open_jn set curr=Acc_Close_Rate.curr,Rate=Acc_Close_Rate.Rate from Open_jn , Acc_Close_Rate  where Open_jn.Ac_Code=Acc_Close_Rate.Ac_Code And Open_jn.date_Work =  '" & "1/1/" & CDbl(yy.Text) + 1 & "' And Open_jn.company='" & Off_Find & "'  ")

        CNN.Execute("update Open_jn set  Rate=Acc_Close_Rate.Rate from Open_jn , Acc_Close_Rate  where Open_jn.Ac_Code=Acc_Close_Rate.Ac_Code And Open_jn.date_Work =  '" & "1/1/" & CDbl(yy.Text) + 1 & "' And Open_jn.company='" & Off_Find & "'  ")


        CNN.Execute("update  Open_jn set Code_Dr=Ac_Code , Code_cr =''  where Amt_Dr <>0 And date_Work =  '" & "1/1/" & CDbl(yy.Text) + 1 & "' And Open_jn.company='" & Off_Find & "' ")
        CNN.Execute("update  Open_jn set Code_cr=Ac_Code , Code_dr =''   where Amt_Cr <>0 And date_Work =  '" & "1/1/" & CDbl(yy.Text) + 1 & "'  And Open_jn.company='" & Off_Find & "' ")

        If CheckBox1.Checked = True Then
            CNN.Execute("update Open_jn set Amount_Dr  = Amt_Dr/rate ,   Amount_cr = Amt_cr/rate , My_Lock=1 , Last_Update = getdate() , Last_User = '" & MUserID & "'  where year(Open_jn.date_Work) =  '" & CDbl(yy.Text) & "'  And Open_jn.company='" & Off_Find & "' ")
        Else
            CNN.Execute("update Open_jn set Amount_Dr  = Amt_Dr/rate ,   Amount_cr = Amt_cr/rate , My_Lock=0 , Last_Update = getdate() , Last_User = '" & MUserID & "'  where year(Open_jn.date_Work) =  '" & CDbl(yy.Text) & "'  And Open_jn.company='" & Off_Find & "' ")
        End If
 
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        MdSearchDataList = "FmClosing"
        FmRate.ShowDialog()
    End Sub

    Private Sub yy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yy.ValueChanged

    End Sub

    Private Sub txtRate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRate.KeyPress
        If e.KeyChar = Chr(13) Then
            txtRate.Text = Format(CDbl(txtRate.Text), "#,##0.00")
        End If
    End Sub

    Private Sub txtRate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRate.TextChanged

    End Sub
End Class