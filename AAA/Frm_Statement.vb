Public Class Frm_Statement
    Dim ACCNO, AccNm, Curr As String
    Private Sub HeaDer()
        LoadSqlData("SELECT * FROM Header WHERE ID=N'ACST01' ", RSC)
        If RSC.RecordCount <> 0 Then
            TxtHeader.Text = Trim(RSC.Fields("Nm").Value.ToString)
            TxtS1.Text = Trim(RSC.Fields("S1").Value.ToString)
            TxtS2.Text = Trim(RSC.Fields("S2").Value.ToString)
            TxtS3.Text = Trim(RSC.Fields("S3").Value.ToString)
            TxtS4.Text = Trim(RSC.Fields("S4").Value.ToString)
            TxtPP.Text = Trim(RSC.Fields("pp").Value.ToString)
        End If
    End Sub
    Private Sub AddHeader()
        LoadSqlData("SELECT * FROM Header WHERE ID=N'ACST01' ", RSC)
        If RSC.RecordCount = 0 Then
            CNN.Execute("INSERT INTO Header(ID,Nm,S1,S2,S3,S4,PP) " & _
                        " values('ACST01',N'" & TxtHeader.Text & "',N'" & TxtS1.Text & "',N'" & TxtS2.Text & "',N'" & TxtS3.Text & "',N'" & TxtS4.Text & "',N'" & TxtPP.Text & "') ")
        Else
            CNN.Execute("UPDATE Header set Nm=N'" & TxtHeader.Text & "',S1=N'" & TxtS1.Text & "',S2=N'" & TxtS2.Text & "',S3=N'" & TxtS3.Text & "',S4=N'" & TxtS4.Text & "',PP=N'" & TxtPP.Text & "' " & _
                        " where ID='ACST01' ")
        End If
    End Sub
    Private Sub FrmCustomer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        HeaDer()
        SetupGrid()

        CmbCust.Items.Clear()
        Call load_Cmb(" SELECT Name  FROM Customer  ORDER BY cnt ", "Name", CmbCust)
        If CmbCust.Items.Count > 0 Then
            CmbCust.SelectedIndex = 0
        End If

        CmbSupp.Items.Clear()
        Call load_Cmb(" SELECT Name  FROM Supplier  ORDER BY cnt ", "Name", CmbSupp)
        If CmbSupp.Items.Count > 0 Then
            CmbSupp.SelectedIndex = 0
        End If

        ' Hide columns 7 and 8 (Balance Debit and Balance Credit)
        FG.Columns(7).Visible = False
        FG.Columns(8).Visible = False
    End Sub

    Private Sub SetupGrid()
        ' Clear and setup DataGridView columns
        FG.Columns.Clear()
        FG.Columns.Add("No", "No.")
        FG.Columns.Add("Date", "Date")
        FG.Columns.Add("VoucherNo", "Voucher No.")
        FG.Columns.Add("Description", "Description")
        FG.Columns.Add("CheckRef", "Check Ref.")
        FG.Columns.Add("Debit", "Debit")
        FG.Columns.Add("Credit", "Credit")
        FG.Columns.Add("BalanceDebit", "Balance Debit")
        FG.Columns.Add("BalanceCredit", "Balance Credit")
        FG.Columns.Add("Remain", "Remain")

        ' Set column widths
        FG.Columns(0).Width = 50
        FG.Columns(1).Width = 100
        FG.Columns(2).Width = 120
        FG.Columns(3).Width = 200
        FG.Columns(4).Width = 100
        FG.Columns(5).Width = 100
        FG.Columns(6).Width = 100
        FG.Columns(7).Width = 120
        FG.Columns(8).Width = 120
        FG.Columns(9).Width = 100

        ' Configure DataGridView properties
        FG.AllowUserToAddRows = False
        FG.ReadOnly = True
        FG.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG.MultiSelect = False
    End Sub
     
    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Me.Close()
    End Sub
 

    Private Sub BtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch.Click
        fmShartOfAccDetail.txtSty.Text = "Acc_Statement"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
        'FG.Focus()
        If CheckBox4.Checked = False Then
            Call LoadFG_AC_CODE()
        Else
            Call LoadFG()
        End If

    End Sub
    Private Sub LoadFG()
        MdStartDate = Format(CDate(Ds.Value), "dd-MM-yyyy")
        MdToDate = Format(CDate(Dt.Value), "dd-MM-yyyy")

        CNN.Execute("DELETE  Ap_PostedLedgers ")
        CNN.Execute("DELETE  Ap_Ope_PostedLedgers_Group ")
        CNN.Execute("DELETE FROM Ap_Open_PostedLedgers ")
        Dim S As Date
        S = MdStartDate
        S = DateAdd("d", CDbl(-1), MdStartDate)
        MuTable = "gen_jn." : Call Find_Company()
        Dim CUST_Supp As String
        If CheckBox4.Checked = True Then
            If RadioButton1.Checked = True Then
                CUST_Supp = " and CustID=N'" & TxtCustID.Text & "' "
            Else
                CUST_Supp = " and SuppID=N'" & TxtSuppID.Text & "' "
            End If 
        Else
            CUST_Supp = ""
        End If
        Dim ACCN As String
        If TxtAccCode.Text = "" Then
            ACCN = ""
        Else
            'ACCN = " and ac_code=N'" & TxtAccCode.Text & "'  "
            ACCN = "and left(ac_code,'" & Len(TxtAccCode.Text) & "')>='" & TxtAccCode.Text & "' and  left(ac_code,'" & Len(TxtAccCode.Text) & "')<='" & TxtAccCode.Text & "' "

        End If

        Dim Cur As String = ""
        Cur = " AND Curr=N'" & CMB_Curr.Text & "' "

        Dim s11 As String = "INSERT INTO Ap_Open_PostedLedgers( ac_code , Amount_Dr , Amount_Cr , Amt_Dr , Amt_Cr , Curr   ) " & _
         "select CustID+SuppID , sum(amount_dr) , sum(amount_cr) ,  sum(amt_dr) , sum(amt_cr) , Curr  from gen_jn WHERE  1=1 " & CUST_Supp & " " & Cur & "  " & ACCN & "  and gen_jn.date_work   BETWEEN '" & "" & Format(MdStartDate, "yyyy") & "-1-1' AND '" & Format(S, "yyyy-MM-dd") & "'   group BY CustID+SuppID , Curr "
        CNN.Execute(s11)
        Dim aa As String
        MuTable = "Open_jn." : Call Find_Company()
        aa = "INSERT INTO Ap_Open_PostedLedgers( ac_code , Amount_Dr , Amount_Cr , Amt_Dr , Amt_Cr  , Curr  ) " & _
         "select ac_code , amount_dr , amount_cr , amt_dr , amt_cr , Curr from Open_jn  WHERE  1=1  " & Cur & "  " & ACCN & "   and date_work='" & "" & Format(MdStartDate, "yyyy") & "-1-1'    "
        CNN.Execute(aa)
        MuTable = "gen_jn." : Call Find_Company()
        aa = " INSERT INTO Ap_PostedLedgers( ac_code , ac_name , date_work , descrip , Open_amount , open_amt , amount_dr , amount_cr , amt_dr , amt_cr , certify , remain , Status ,  curr ) " & _
               "select CustID+SuppID ,'' ,date_work, descrip ,0 ,0 ,amount_dr ,amount_cr , amt_dr , amt_cr ,certify, 0 , 0 ,  Curr  from gen_jn    WHERE   1=1 " & CUST_Supp & "  " & Cur & "  " & ACCN & "   and gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'   "
        CNN.Execute(aa)

        CNN.Execute("Delete Ap_Ope_PostedLedgers_Group ")
        CNN.Execute("Insert Into Ap_Ope_PostedLedgers_Group (Ac_Code  ,Amount, Amt  )" & _
                     " Select Ac_Code ,Sum(Amount_Dr - Amount_Cr)  ,Sum(Amt_Dr - Amt_Cr)  from Ap_Open_PostedLedgers group by Ac_Code ")
        CNN.Execute("Update Ap_PostedLedgers set  Ap_PostedLedgers.Open_Amt = Ap_Ope_PostedLedgers_Group.Amt  , Ap_PostedLedgers.Open_Amount = Ap_Ope_PostedLedgers_Group.Amount  From Ap_PostedLedgers , Ap_Ope_PostedLedgers_Group  Where Ap_PostedLedgers.Ac_Code = Ap_Ope_PostedLedgers_Group.Ac_Code   ")

        CNN.Execute("Delete Ap_Ope_PostedLedgers_Group ")
        CNN.Execute("Insert Into Ap_Ope_PostedLedgers_Group (Ac_Code , Amount  , Amt )" & _
                  " Select Ac_Code ,Sum(Amount_Dr - Amount_Cr)   ,Sum(Amt_Dr - Amt_Cr)   from Ap_Open_PostedLedgers group by Ac_Code ")
        CNN.Execute("Update Ap_PostedLedgers set Ap_PostedLedgers.Open_Amount = Ap_Ope_PostedLedgers_Group.Amount , Ap_PostedLedgers.Open_Amt = Ap_Ope_PostedLedgers_Group.Amt   From Ap_PostedLedgers , Ap_Ope_PostedLedgers_Group  Where Ap_PostedLedgers.Ac_Code = Ap_Ope_PostedLedgers_Group.Ac_Code  ")

        CNN.Execute(" delete Ap_PostedLedgers_Rem ")
        CNN.Execute(" insert into Ap_PostedLedgers_Rem  (ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work)  " & _
         " select ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work from Ap_PostedLedgers order by  Ac_Code,Date_Work , certify ,cnt Asc  ")
        CNN.Execute("delete Ap_PostedLedgers   ")


        Dim KK As String = "insert into Ap_PostedLedgers  (ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work)  " & _
   " select ac_code, '', '', 'Opening Balance', 0, 0, 0, 0, 0, 0, 0, 0, Open_amount, '', '', '','" & Format(MdStartDate, "yyyy-MM-dd") & "' from Ap_PostedLedgers_Rem group by ac_code,Open_amount  "
        CNN.Execute(KK)

        CNN.Execute("insert into Ap_PostedLedgers  (ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work)  " & _
         " select ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work from Ap_PostedLedgers_Rem order by cnt   ")

        CNN.Execute(" delete Ap_PostedLedgers_Rem ")
        CNN.Execute(" insert into Ap_PostedLedgers_Rem (Ac_Code , cnt_Mat ,remain )  " & _
          " select Ac_Code , cnt,Open_amt+(select SUM(Amt_Dr-Amt_cr) from Ap_PostedLedgers  where Ac_Code = x.ac_code    And cnt <= x.cnt )as Rem from Ap_PostedLedgers as x  order by  cnt Asc  ")


        CNN.Execute("Update Ap_PostedLedgers set Ap_PostedLedgers.remain = Ap_PostedLedgers_Rem.remain from Ap_PostedLedgers , Ap_PostedLedgers_Rem where  Ap_PostedLedgers.cnt = Ap_PostedLedgers_Rem.cnt_Mat  " & _
            "  update	 Ap_PostedLedgers set remain=0 where remain is null ")

        CNN.Execute(" Update Ap_PostedLedgers set Open_dr = 0 where  Open_dr is Null Update Ap_PostedLedgers set Open_cr = 0 where  Open_cr is Null Update Ap_PostedLedgers set Open_amount = 0 where  Open_amount is Null Update Ap_PostedLedgers set open_amt = 0 where  open_amt is Null Update Ap_PostedLedgers set amount_dr = 0 where  amount_dr is Null Update Ap_PostedLedgers set amount_cr = 0 where  amount_cr is Null Update Ap_PostedLedgers set amt_dr = 0 where  amt_dr is Null Update Ap_PostedLedgers set amt_cr = 0 where  amt_cr is Null Update Ap_PostedLedgers set remain = 0 where  remain is Null ")

        'If CheckBox1.Checked = True Then
        '    CNN.Execute(" delete Ap_PostedLedgers_Rem ")
        '    CNN.Execute(" insert into Ap_PostedLedgers_Rem (Ac_Code , cnt_Mat ,remain )  " & _
        '      " select Ac_Code , cnt,Open_amount+(select SUM(Amount_Dr-Amount_cr) from Ap_PostedLedgers  where Ac_Code = x.ac_code    And cnt <= x.cnt )as Rem_Curr from Ap_PostedLedgers as x  order by  cnt Asc  ")
        '    CNN.Execute("Update Ap_PostedLedgers set Ap_PostedLedgers.Rem_Curr = Ap_PostedLedgers_Rem.remain from Ap_PostedLedgers , Ap_PostedLedgers_Rem where  Ap_PostedLedgers.cnt = Ap_PostedLedgers_Rem.cnt_Mat  " & _
        '        "  update	 Ap_PostedLedgers set remain=0 where remain is null ")
        'End If


        MuTable = "Open_jn." : Call Find_Company()
        CNN.Execute("update Open_jn set Lck=0 where year(date_work)= " & Format(MdStartDate, "yyyy") & "  ")
        CNN.Execute("update Open_jn set Lck=1 from Open_jn, Ap_PostedLedgers   where year(Open_jn.date_work)= " & Format(MdStartDate, "yyyy") & " And  Open_jn.ac_code =  Ap_PostedLedgers.ac_code ")
        Dim s7 As String = " Insert into Ap_PostedLedgers (ac_code,Open_dr,Open_Cr,Open_amount,open_amt,amount_dr,amount_cr,amt_dr,amt_cr,Curr,remain,Rem_Curr,Status)  " & _
                    "select ac_code,0,0, Sum(amount_dr-amount_cr),Sum(amt_dr-amt_cr),0,0,0,0,Curr,0,0,0 from Open_jn where  1=1  " & Cur & "  and ac_code='" & TxtAccCode.Text & "' and Lck=0 And year(date_work)= " & Format(MdStartDate, "yyyy") & "  Group by ac_code , Curr"
        CNN.Execute(s7)

        CNN.Execute("Update Ap_PostedLedgers set open_amt = Ap_Ope_PostedLedgers_Group.Amt from Ap_PostedLedgers ,Ap_Ope_PostedLedgers_Group  where  Ap_PostedLedgers.Ac_Code = Ap_Ope_PostedLedgers_Group.Ac_Code")

        'CNN.Execute("update Ap_PostedLedgers set Ap_PostedLedgers.ac_name=Acc_Code.Name_L from Ap_PostedLedgers , Acc_Code where Ap_PostedLedgers.Ac_Code = Acc_Code.Ac_Code")
        CNN.Execute("update Ap_PostedLedgers set Ap_PostedLedgers.ac_name=gen_jn.ac_name from Ap_PostedLedgers , gen_jn where Ap_PostedLedgers.Ac_Code = gen_jn.Ac_Code and   Ap_PostedLedgers.certify =gen_jn.certify ")

        CNN.Execute("Update Ap_PostedLedgers set remain = open_amt where Descrip=N'Opening Balance' ")

        LoadListFG()
        'Call loadColor()
    End Sub
    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            Panel4.Enabled = True

        Else
            Panel4.Enabled = False
        End If
        RadioButton1.Checked = True
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            Dim Rs As New ADODB.Recordset
            With Rs
                Call LoadSqlData("select    * from Customer where 1=1  and Name=N'" & CmbCust.Text & "'   ", Rs)
                If .RecordCount > 0 Then
                    TxtCustID.Text = Trim(.Fields("Code").Value.ToString)
                Else
                    TxtCustID.Text = ""
                End If
            End With

        Else
            TxtCustID.Text = ""
        End If

        CmbSupp.Enabled = False
        CmbCust.Enabled = True
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            Dim Rs As New ADODB.Recordset
            With Rs
                Call LoadSqlData("select    * from Supplier where 1=1  and Name=N'" & CmbSupp.Text & "'   ", Rs)
                If .RecordCount > 0 Then
                    TxtSuppID.Text = Trim(.Fields("Code").Value.ToString)
                Else
                    TxtSuppID.Text = ""
                End If
            End With

        Else
            TxtSuppID.Text = ""
        End If

        CmbSupp.Enabled = True
        CmbCust.Enabled = False
    End Sub

    Private Sub CmbCust_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCust.SelectedIndexChanged
        If RadioButton1.Checked = True Then
            Dim Rs As New ADODB.Recordset
            With Rs
                Call LoadSqlData("select    * from Customer where 1=1  and Name=N'" & CmbCust.Text & "'   ", Rs)
                If .RecordCount > 0 Then
                    TxtCustID.Text = Trim(.Fields("Code").Value.ToString)
                Else
                    TxtCustID.Text = ""
                End If
            End With
            LoadFG()
        End If
    End Sub

    Private Sub CmbSupp_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbSupp.SelectedIndexChanged
        If RadioButton2.Checked = True Then
            Dim Rs As New ADODB.Recordset
            With Rs
                Call LoadSqlData("select    * from Supplier where 1=1  and Name=N'" & CmbSupp.Text & "'   ", Rs)
                If .RecordCount > 0 Then
                    TxtSuppID.Text = Trim(.Fields("Code").Value.ToString)
                Else
                    TxtSuppID.Text = ""
                End If
            End With
            LoadFG()
        End If
    End Sub
    Private Sub LoadFG_AC_CODE()
        MdStartDate = Format(CDate(Ds.Value), "dd-MM-yyyy")
        MdToDate = Format(CDate(Dt.Value), "dd-MM-yyyy")

        CNN.Execute("DELETE  Ap_PostedLedgers ")
        CNN.Execute("DELETE  Ap_Ope_PostedLedgers_Group ")
        CNN.Execute("DELETE FROM Ap_Open_PostedLedgers ")
        Dim S As Date
        S = MdStartDate
        S = DateAdd("d", CDbl(-1), MdStartDate)
        MuTable = "gen_jn." : Call Find_Company()
        Dim CUST_Supp As String
 
        Dim ACCN As String
        If TxtAccCode.Text = "" Then
            ACCN = ""
        Else
            'ACCN = " and ac_code=N'" & TxtAccCode.Text & "'  "
            ACCN = "and left(ac_code,'" & Len(TxtAccCode.Text) & "')>='" & TxtAccCode.Text & "' and  left(ac_code,'" & Len(TxtAccCode.Text) & "')<='" & TxtAccCode.Text & "' "

        End If

        Dim Cur As String = ""
        Cur = " AND Curr=N'" & CMB_Curr.Text & "' "

        Dim s11 As String = "INSERT INTO Ap_Open_PostedLedgers( ac_code , Amount_Dr , Amount_Cr , Amt_Dr , Amt_Cr , Curr   ) " & _
         "select Ac_Code , sum(amount_dr) , sum(amount_cr) ,  sum(amt_dr) , sum(amt_cr) , Curr  from gen_jn WHERE  1=1   " & Cur & "  " & ACCN & "  and gen_jn.date_work   BETWEEN '" & "" & Format(MdStartDate, "yyyy") & "-1-1' AND '" & Format(S, "yyyy-MM-dd") & "'   group BY Ac_Code , Curr "
        CNN.Execute(s11)
        Dim aa As String
        MuTable = "Open_jn." : Call Find_Company()
        aa = "INSERT INTO Ap_Open_PostedLedgers( ac_code , Amount_Dr , Amount_Cr , Amt_Dr , Amt_Cr  , Curr  ) " & _
         "select ac_code , amount_dr , amount_cr , amt_dr , amt_cr , Curr from Open_jn  WHERE  1=1  " & Cur & "  " & ACCN & "   and date_work='" & "" & Format(MdStartDate, "yyyy") & "-1-1'    "
        CNN.Execute(aa)
        MuTable = "gen_jn." : Call Find_Company()
        aa = " INSERT INTO Ap_PostedLedgers( ac_code , ac_name , date_work , descrip , Open_amount , open_amt , amount_dr , amount_cr , amt_dr , amt_cr , certify , remain , Status ,  curr ) " & _
               "select ac_code ,'' ,date_work, descrip ,0 ,0 ,amount_dr ,amount_cr , amt_dr , amt_cr ,certify, 0 , 0 ,  Curr  from gen_jn    WHERE   1=1  " & Cur & "  " & ACCN & "   and gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "'   "
        CNN.Execute(aa)
        CNN.Execute("UPDATE Ap_PostedLedgers set ac_code='" & TxtAccCode.Text & "' ")
        CNN.Execute("UPDATE Ap_Open_PostedLedgers set ac_code='" & TxtAccCode.Text & "' ")

        CNN.Execute("Delete Ap_Ope_PostedLedgers_Group ")
        CNN.Execute("Insert Into Ap_Ope_PostedLedgers_Group (Ac_Code  ,Amount, Amt  )" & _
                     " Select Ac_Code ,Sum(Amount_Dr - Amount_Cr)  ,Sum(Amt_Dr - Amt_Cr)  from Ap_Open_PostedLedgers group by Ac_Code ")
        CNN.Execute("Update Ap_PostedLedgers set  Ap_PostedLedgers.Open_Amt = Ap_Ope_PostedLedgers_Group.Amt  , Ap_PostedLedgers.Open_Amount = Ap_Ope_PostedLedgers_Group.Amount  From Ap_PostedLedgers , Ap_Ope_PostedLedgers_Group  Where Ap_PostedLedgers.Ac_Code = Ap_Ope_PostedLedgers_Group.Ac_Code   ")

        CNN.Execute("Delete Ap_Ope_PostedLedgers_Group ")
        CNN.Execute("Insert Into Ap_Ope_PostedLedgers_Group (Ac_Code , Amount  , Amt )" & _
                  " Select Ac_Code ,Sum(Amount_Dr - Amount_Cr)   ,Sum(Amt_Dr - Amt_Cr)   from Ap_Open_PostedLedgers group by Ac_Code ")
        CNN.Execute("Update Ap_PostedLedgers set Ap_PostedLedgers.Open_Amount = Ap_Ope_PostedLedgers_Group.Amount , Ap_PostedLedgers.Open_Amt = Ap_Ope_PostedLedgers_Group.Amt   From Ap_PostedLedgers , Ap_Ope_PostedLedgers_Group  Where Ap_PostedLedgers.Ac_Code = Ap_Ope_PostedLedgers_Group.Ac_Code  ")

        CNN.Execute(" delete Ap_PostedLedgers_Rem ")
        CNN.Execute(" insert into Ap_PostedLedgers_Rem  (ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work)  " & _
         " select ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work from Ap_PostedLedgers order by  Ac_Code,Date_Work , certify ,cnt Asc  ")
        CNN.Execute("delete Ap_PostedLedgers   ")


        Dim KK As String = "insert into Ap_PostedLedgers  (ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work)  " & _
   " select ac_code, '', '', 'Opening Balance', 0, 0, 0, 0, 0, 0, 0, 0, Open_amount, '', '', '','" & Format(MdStartDate, "yyyy-MM-dd") & "' from Ap_PostedLedgers_Rem group by ac_code,Open_amount  "
        CNN.Execute(KK)
        CNN.Execute("UPDATE Ap_PostedLedgers set ac_code='" & TxtAccCode.Text & "' ")

        CNN.Execute("insert into Ap_PostedLedgers  (ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work)  " & _
         " select ac_code, certify, ac_name, descrip, Open_dr, Open_cr, Open_amount, open_amt, amount_dr, amount_cr, amt_dr, amt_cr, remain, Status, curr, Lck, date_work from Ap_PostedLedgers_Rem order by cnt   ")

        'CNN.Execute("UPDATE Ap_PostedLedgers set ac_code='" & TxtAccCode.Text & "' ")

        CNN.Execute(" delete Ap_PostedLedgers_Rem ")
        CNN.Execute(" insert into Ap_PostedLedgers_Rem (Ac_Code , cnt_Mat ,remain )  " & _
          " select Ac_Code , cnt,Open_amt+(select SUM(Amt_Dr-Amt_cr) from Ap_PostedLedgers  where Ac_Code = x.ac_code    And cnt <= x.cnt )as Rem from Ap_PostedLedgers as x  order by  cnt Asc  ")


        CNN.Execute("Update Ap_PostedLedgers set Ap_PostedLedgers.remain = Ap_PostedLedgers_Rem.remain from Ap_PostedLedgers , Ap_PostedLedgers_Rem where  Ap_PostedLedgers.cnt = Ap_PostedLedgers_Rem.cnt_Mat  " & _
            "  update	 Ap_PostedLedgers set remain=0 where remain is null ")

        CNN.Execute(" Update Ap_PostedLedgers set Open_dr = 0 where  Open_dr is Null Update Ap_PostedLedgers set Open_cr = 0 where  Open_cr is Null Update Ap_PostedLedgers set Open_amount = 0 where  Open_amount is Null Update Ap_PostedLedgers set open_amt = 0 where  open_amt is Null Update Ap_PostedLedgers set amount_dr = 0 where  amount_dr is Null Update Ap_PostedLedgers set amount_cr = 0 where  amount_cr is Null Update Ap_PostedLedgers set amt_dr = 0 where  amt_dr is Null Update Ap_PostedLedgers set amt_cr = 0 where  amt_cr is Null Update Ap_PostedLedgers set remain = 0 where  remain is Null ")

        'If CheckBox1.Checked = True Then
        '    CNN.Execute(" delete Ap_PostedLedgers_Rem ")
        '    CNN.Execute(" insert into Ap_PostedLedgers_Rem (Ac_Code , cnt_Mat ,remain )  " & _
        '      " select Ac_Code , cnt,Open_amount+(select SUM(Amount_Dr-Amount_cr) from Ap_PostedLedgers  where Ac_Code = x.ac_code    And cnt <= x.cnt )as Rem_Curr from Ap_PostedLedgers as x  order by  cnt Asc  ")
        '    CNN.Execute("Update Ap_PostedLedgers set Ap_PostedLedgers.Rem_Curr = Ap_PostedLedgers_Rem.remain from Ap_PostedLedgers , Ap_PostedLedgers_Rem where  Ap_PostedLedgers.cnt = Ap_PostedLedgers_Rem.cnt_Mat  " & _
        '        "  update	 Ap_PostedLedgers set remain=0 where remain is null ")
        'End If


        MuTable = "Open_jn." : Call Find_Company()
        CNN.Execute("update Open_jn set Lck=0 where year(date_work)= " & Format(MdStartDate, "yyyy") & "  ")
        CNN.Execute("update Open_jn set Lck=1 from Open_jn, Ap_PostedLedgers   where year(Open_jn.date_work)= " & Format(MdStartDate, "yyyy") & " And  Open_jn.ac_code =  Ap_PostedLedgers.ac_code ")
        Dim s7 As String = " Insert into Ap_PostedLedgers (ac_code,Open_dr,Open_Cr,Open_amount,open_amt,amount_dr,amount_cr,amt_dr,amt_cr,Curr,remain,Rem_Curr,Status)  " & _
                    "select ac_code,0,0, Sum(amount_dr-amount_cr),Sum(amt_dr-amt_cr),0,0,0,0,Curr,0,0,0 from Open_jn where  1=1  " & Cur & "  and ac_code='" & TxtAccCode.Text & "' and Lck=0 And year(date_work)= " & Format(MdStartDate, "yyyy") & "  Group by ac_code , Curr"
        CNN.Execute(s7)
        CNN.Execute("UPDATE Ap_PostedLedgers set ac_code='" & TxtAccCode.Text & "' ")
        CNN.Execute("Update Ap_PostedLedgers set open_amt = Ap_Ope_PostedLedgers_Group.Amt from Ap_PostedLedgers ,Ap_Ope_PostedLedgers_Group  where  Ap_PostedLedgers.Ac_Code = Ap_Ope_PostedLedgers_Group.Ac_Code")

        'CNN.Execute("update Ap_PostedLedgers set Ap_PostedLedgers.ac_name=Acc_Code.Name_L from Ap_PostedLedgers , Acc_Code where Ap_PostedLedgers.Ac_Code = Acc_Code.Ac_Code")
        CNN.Execute("update Ap_PostedLedgers set Ap_PostedLedgers.ac_name=gen_jn.ac_name from Ap_PostedLedgers , gen_jn where Ap_PostedLedgers.Ac_Code = gen_jn.Ac_Code and   Ap_PostedLedgers.certify =gen_jn.certify ")

        CNN.Execute("Update Ap_PostedLedgers set remain = open_amt where Descrip=N'Opening Balance' ")

        LoadListFG_AC_Code()
        'Call loadColor()
    End Sub
    Private Sub BtnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        AddHeader()
        If CheckBox4.Checked = False Then
            'Dim RSN As New ADODB.Recordset
            LoadFG_AC_CODE()
            Call Office()

            If CMB_Curr.Text = "LAK" Then
                ACCNO = "00-" & TxtAccCode.Text
            ElseIf CMB_Curr.Text = "USD" Then
                ACCNO = "01-" & TxtAccCode.Text
            End If
            Curr = CMB_Curr.Text
            AccNm = TxtAccName.Text
            Call LoadLoGO()
            Dim AAA As String = " For the Period " & Ds.Text & " - " & Dt.Text
            SLF = "SELECT N'" & TxtS1.Text & "'  as TxtS1,N'" & TxtS2.Text & "'  as TxtS2,N'" & TxtS3.Text & "'  as TxtS3,N'" & TxtS4.Text & "'  as TxtS4,N'" & TxtPP.Text & "'  as TxtPP, N'" & ACCNO & "'  as ACCN  , N'" & Curr & "'  as Curr  , N'" & AccNm & "'  as AccNm  ,  N'" & AAA & "'  as DD  , N'" & MDSgn1 & "' as S1,N'" & MDSgn2 & "' as S2,N'" & MDSgn3 & "' as S3, N'" & MDSgn4 & "' as S4,  N'" & MDSgn5 & "' as S5,   N'" & MDSgn6 & "' as S6,  N'" & RptPro & "' as pp, " & mformat & "  as mformat  ,    *   FROM Ap_PostedLedgers Order by CNT asc "
            'Dim Rs As New ADODB.Recordset
            Dim RSN As New ADODB.Recordset
            With RSN
                If .State = ConnectionState.Open Then .Close()
                .Open(SLF, CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
                If .EOF Then Exit Sub
            End With
            Dim FrmPreview As New FmPreview : FrmClosing()
            Dim Rpt As New Object
            Rpt = New Acc_Statement_Ac_Code
            If MdShowLOGO = 1 Then
                Rpt.Subreports(0).SetDataSource(RsLOGO)
            End If
            '1388300:
            'Dim myText3 As CrystalDecisions.CrystalReports.Engine.TextObject
            'If CMB_Curr.Text = "LAK" Then
            '    myText3 = CType(Rpt.ReportDefinition.ReportObjects.Item("Acc"), CrystalDecisions.CrystalReports.Engine.TextObject)
            '    myText3.Text = "00-" & TxtAccCode.Text
            'ElseIf CMB_Curr.Text = "USD" Then
            '    myText3 = CType(Rpt.ReportDefinition.ReportObjects.Item("Acc"), CrystalDecisions.CrystalReports.Engine.TextObject)
            '    myText3.Text = "01-" & TxtAccCode.Text
            'End If


            'myText3 = CType(Rpt.ReportDefinition.ReportObjects.Item("Curr"), CrystalDecisions.CrystalReports.Engine.TextObject)
            'myText3.Text = CMB_Curr.Text
            'myText3 = CType(Rpt.ReportDefinition.ReportObjects.Item("AccNm"), CrystalDecisions.CrystalReports.Engine.TextObject)
            'myText3.Text = TxtAccName.Text

            Rpt.SetDataSource(RSN)
            Rpt.Refresh()
            FrmPreview.ReportViewer.ReportSource = Rpt
            FrmPreview.ReportViewer.DisplayGroupTree = False
            FrmPreview.MdiParent = FmMain
            FrmPreview.WindowState = FormWindowState.Maximized
            FrmPreview.Show()
            FrmPreview.Focus()
        Else

            'MsgBox("No") : Exit Sub
            'End If

            LoadFG()
            Call Office()
            'Dim AAA As String = " ວັນທີ່ " & Ds.Text & " ຫາ " & Dt.Text
            Dim AAA As String = " For the Period " & Ds.Text & " - " & Dt.Text
            'SLF = "SELECT   N'" & AAA & "'  as DD  ,N'" & MDSgn1 & "' as S1,N'" & MDSgn2 & "' as S2,N'" & MDSgn3 & "' as S3, N'" & MDSgn4 & "' as S4,  N'" & MDSgn5 & "' as S5,   N'" & MDSgn6 & "' as S6,  N'" & RptPro & "' as pp, " & mformat & "  as mformat  ,    *   FROM Ap_PostedLedgers Order by CNT asc "
            SLF = "SELECT N'" & TxtS1.Text & "'  as TxtS1,N'" & TxtS2.Text & "'  as TxtS2,N'" & TxtS3.Text & "'  as TxtS3,N'" & TxtS4.Text & "'  as TxtS4,N'" & TxtPP.Text & "'  as TxtPP, N'" & ACCNO & "'  as ACCN  , N'" & Curr & "'  as Curr  , N'" & AccNm & "'  as AccNm  ,  N'" & AAA & "'  as DD  , N'" & MDSgn1 & "' as S1,N'" & MDSgn2 & "' as S2,N'" & MDSgn3 & "' as S3, N'" & MDSgn4 & "' as S4,  N'" & MDSgn5 & "' as S5,   N'" & MDSgn6 & "' as S6,  N'" & RptPro & "' as pp, " & mformat & "  as mformat  ,    *   FROM Ap_PostedLedgers Order by CNT asc "
            Call LoadLoGO()
            Dim Rs As New ADODB.Recordset
            With Rs
                If .State = ConnectionState.Open Then .Close()
                .Open(SLF, CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
                If .EOF Then Exit Sub
            End With
            Dim FrmPreview As New FmPreview : FrmClosing()
            Dim Rpt As New Object
            Rpt = New Acc_Statement
            If MdShowLOGO = 1 Then
                Rpt.Subreports(0).SetDataSource(RsLOGO)
            End If
            Rpt.SetDataSource(Rs)
            FrmPreview.ReportViewer.ReportSource = Rpt
            FrmPreview.ReportViewer.DisplayGroupTree = False
            FrmPreview.MdiParent = FmMain
            FrmPreview.WindowState = FormWindowState.Maximized
            FrmPreview.Show()
            FrmPreview.Focus()
        End If
    End Sub
    Public Sub LoadListFG()
        ' Clear existing rows
        FG.Rows.Clear()

        Dim CUST_Supp As String

        If RadioButton1.Checked = True Then
            CUST_Supp = " and ac_code=N'" & TxtCustID.Text & "' "
        Else
            CUST_Supp = " and ac_code=N'" & TxtSuppID.Text & "' "
        End If

        With RSC
            'Call LoadSqlData("SELECT * FROM  Ap_PostedLedgers    WHERE  ac_code=N'" & TxtAccCode.Text & "' order by cnt ASC  ", RSC)
            Call LoadSqlData("SELECT * FROM  Ap_PostedLedgers    WHERE 1=1 " & CUST_Supp & " order by cnt ASC  ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.Rows.Add(.AbsolutePosition, _
                      Trim(CStr(.Fields("Date_Work").Value.ToString)), _
                      Trim(CStr(.Fields("Certify").Value.ToString)), _
                      Trim(CStr(.Fields("Descrip").Value.ToString)), _
                      "", _  ' Check Ref column
                      Format(CDbl(.Fields("amt_dr").Value), "##,##0.00"), _
                      Format(CDbl(.Fields("amt_Cr").Value), "##,##0.00"), _
                      Format(CDbl(.Fields("amt_dr").Value), "##,##0.00"), _
                      Format(CDbl(.Fields("amt_Cr").Value), "##,##0.00"), _
                      Format(CDbl(.Fields("remain").Value), "##,##0.00"))
                    .MoveNext()
                End While
            End If
        End With

        Call LoadSqlData("SELECT isnull(sum(amt_dr),0) as amt_dr, isnull(sum(amt_Cr),0) as amt_Cr FROM  Ap_PostedLedgers    WHERE 1=1 " & CUST_Supp & "  ", RSC)
        If RSC.RecordCount > 0 Then
            TxtDebit.Text = Format(CDbl(RSC.Fields("amt_dr").Value), "#,##0.00")
            TxtCredit.Text = Format(CDbl(RSC.Fields("amt_Cr").Value), "#,##0.00")
        Else
            TxtDebit.Text = "0.00"
            TxtCredit.Text = "0.00"
        End If
        Call LoadSqlData("SELECT remain FROM  Ap_PostedLedgers    WHERE 1=1 " & CUST_Supp & " and Descrip=N'Opening Balance'   ", RSC)
        If RSC.RecordCount > 0 Then
            TxtOpen.Text = Format(CDbl(RSC.Fields("remain").Value), "#,##0.00")
        Else
            TxtOpen.Text = "0.00"
        End If


        TxtEnd.Text = (CDbl(TxtOpen.Text) + CDbl(TxtDebit.Text)) - CDbl(TxtCredit.Text)
        TxtEnd.Text = Format(CDbl(TxtEnd.Text), "#,##0.00")

    End Sub

    Private Sub TxtAccCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtAccCode.KeyPress
        Dim rs As New ADODB.Recordset
        If CheckBox4.Checked = False Then
            If e.KeyChar = Chr(13) Then
                With rs
                    Dim sa As String = " SELECT    *  from Acc_Code   WHERE 1=1 AND  AC_CODE='" & TxtAccCode.Text & "' "
                    Call LoadSqlData(sa, rs)
                    If .RecordCount <> 0 Then
                        TxtAccCode.Text = (.Fields("AC_CODE").Value.ToString)
                        TxtAccName.Text = (.Fields("Name_L").Value.ToString) & " / " & (.Fields("Name_E").Value.ToString)
                    Else
                        TxtAccCode.Text = ""
                        TxtAccName.Text = ""
                    End If

                End With

                LoadFG_AC_CODE()
            End If
        Else
            If e.KeyChar = Chr(13) Then
                With rs
                    Dim sa As String = " SELECT    *  from Acc_Code   WHERE 1=1 AND  AC_CODE='" & TxtAccCode.Text & "' "
                    Call LoadSqlData(sa, rs)
                    If .RecordCount <> 0 Then
                        TxtAccCode.Text = (.Fields("AC_CODE").Value.ToString)
                        'TxtAccName.Text = (.Fields("Name_L").Value.ToString)
                        TxtAccName.Text = (.Fields("Name_L").Value.ToString) & " / " & (.Fields("Name_E").Value.ToString)
                    Else
                        TxtAccCode.Text = ""
                        TxtAccName.Text = ""
                    End If

                End With

                LoadFG()
            End If
        End If
      
    End Sub

    Private Sub TxtAccCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtAccCode.TextChanged

    End Sub

    Private Sub Ds_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Ds.LostFocus
        LoadFG()
    End Sub

    Private Sub Ds_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ds.ValueChanged

    End Sub

    Private Sub Dt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Dt.LostFocus
        LoadFG()
    End Sub
    Public Sub LoadListFG_AC_Code()
        ' Clear existing rows
        FG.Rows.Clear()
        Dim CUST_Supp As String

        'If RadioButton1.Checked = True Then
        '    CUST_Supp = " and ac_code=N'" & TxtCustID.Text & "' "
        'Else
        '    CUST_Supp = " and ac_code=N'" & TxtSuppID.Text & "' "
        'End If




        With RSC
            'Call LoadSqlData("SELECT * FROM  Ap_PostedLedgers    WHERE  ac_code=N'" & TxtAccCode.Text & "' order by cnt ASC  ", RSC)
            Call LoadSqlData("SELECT * FROM  Ap_PostedLedgers    WHERE 1=1 order by cnt ASC  ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.Rows.Add(.AbsolutePosition, _
                      Trim(CStr(.Fields("Date_Work").Value.ToString)), _
                      Trim(CStr(.Fields("Certify").Value.ToString)), _
                      Trim(CStr(.Fields("Descrip").Value.ToString)), _
                      "", _  ' Check Ref column
                      Format(CDbl(.Fields("amt_dr").Value), "##,##0.00"), _
                      Format(CDbl(.Fields("amt_Cr").Value), "##,##0.00"), _
                      Format(CDbl(.Fields("amt_dr").Value), "##,##0.00"), _
                      Format(CDbl(.Fields("amt_Cr").Value), "##,##0.00"), _
                      Format(CDbl(.Fields("remain").Value), "##,##0.00"))
                    .MoveNext()
                End While
            End If
        End With

        Call LoadSqlData("SELECT isnull(sum(amt_dr),0) as amt_dr, isnull(sum(amt_Cr),0) as amt_Cr FROM  Ap_PostedLedgers    WHERE 1=1    ", RSC)
        If RSC.RecordCount > 0 Then
            TxtDebit.Text = Format(CDbl(RSC.Fields("amt_dr").Value), "#,##0.00")
            TxtCredit.Text = Format(CDbl(RSC.Fields("amt_Cr").Value), "#,##0.00")
        Else
            TxtDebit.Text = "0.00"
            TxtCredit.Text = "0.00"
        End If
        Call LoadSqlData("SELECT remain FROM  Ap_PostedLedgers    WHERE 1=1  and Descrip=N'Opening Balance'   ", RSC)
        If RSC.RecordCount > 0 Then
            TxtOpen.Text = Format(CDbl(RSC.Fields("remain").Value), "#,##0.00")
        Else
            TxtOpen.Text = "0.00"
        End If


        TxtEnd.Text = (CDbl(TxtOpen.Text) + CDbl(TxtDebit.Text)) - CDbl(TxtCredit.Text)
        TxtEnd.Text = Format(CDbl(TxtEnd.Text), "#,##0.00")

    End Sub
End Class