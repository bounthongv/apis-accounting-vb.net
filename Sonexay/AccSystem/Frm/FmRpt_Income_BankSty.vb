Public Class FmRpt_Income_BankSty
    Dim bls1 As String
    Dim MonthLetter1 As String
    Dim MdStartDate As Date
    Dim MdToDate As Date
    Dim MdQuarter As Date

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

    Private Sub BtnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click



        '============


        CNN.Execute("update  Ap_Rpt_Income_Item set amt_Last_M_dr  =  " & CDbl(0) & " , amt_Last_M_Cr  =  " & CDbl(0) & " , amt_M_dr  =  " & CDbl(0) & " , amt_M_cr  = " & CDbl(0) & "  , amt_Q_dr  =  " & CDbl(0) & " , amt_Q_cr  = " & CDbl(0) & " , amt_y_dr  =  " & CDbl(0) & " , amt_y_cr  = " & CDbl(0) & "   ")
        CNN.Execute("update Ap_Rpt_Income set  Last_Month  =  " & CDbl(0) & "  , Current_Month  =  " & CDbl(0) & "  , Quarter_to_Date  =  " & CDbl(0) & " , Year_to_Date  =  " & CDbl(0) & "       ")
        CNN.Execute("DELETE FROM Ap_Rpt_Incon_Detail ")
        CNN.Execute("DELETE FROM Ap_balance_6_col WHERE cnt <> '" & "" & "'")
        CNN.Execute("DELETE FROM Ap_balance_6 WHERE cnt <> '" & "" & "'")
        LoadOpen_Jn1()
        LoadOpen_Jn2()
        LoadOpen_Jn3()
        LoadOpen_Jn6()
        LoadOpen_Jn7()
        LoadOpen_Jn11()
        LoadOpen_Jn12()
        LoadOpen_Jn12_11()
        LoadOpen_Jn14_1()
        LoadOpen_Jn15()

        SelcectIn()
        UpdateIIn()

        SelectOut()
        UpdateOut()

        Update_Sum()

        Call LoadReport()

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
    Private Sub Update_Sum()
        '        Dim InteIn_M_L As Double = 0
        '        Dim InteIn_M As Double = 0
        '        Dim InteIn_Q As Double = 0
        '        Dim InteIn_Y As Double = 0
        '        Dim xm1_L As Double = 0
        '        Dim xm2_L As Double = 0
        '        Dim xm1 As Double = 0
        '        Dim xm2 As Double = 0
        '        Dim xq1 As Double = 0
        '        Dim xq2 As Double = 0
        '        Dim xy1 As Double = 0
        '        Dim xy2 As Double = 0


        '        '==========5
        '        Dim GOI_M_L As Double = 0
        '        Dim GOI_M As Double = 0
        '        Dim GOI_Q As Double = 0
        '        Dim GOI_Y As Double = 0
        '        Dim GOIM3_L As Double = 0
        '        Dim GOIM4_L As Double = 0
        '        Dim GOIM3 As Double = 0
        '        Dim GOIM4 As Double = 0
        '        Dim GOIQ3 As Double = 0
        '        Dim GOIQ4 As Double = 0
        '        Dim GOIY3 As Double = 0
        '        Dim GOIY4 As Double = 0

        '        '==========7
        '        Dim NI_M_L As Double = 0
        '        Dim NI_M As Double = 0
        '        Dim NI_Q As Double = 0
        '        Dim NI_Y As Double = 0
        '        Dim NIM5_L As Double = 0
        '        Dim NIM6_L As Double = 0
        '        Dim NIM5 As Double = 0
        '        Dim NIM6 As Double = 0
        '        Dim NIQ5 As Double = 0
        '        Dim NIQ6 As Double = 0
        '        Dim NIY5 As Double = 0
        '        Dim NIY6 As Double = 0


        '        '==========10
        '        Dim PBT_M_L As Double = 0
        '        Dim PBT_M As Double = 0
        '        Dim PBT_Q As Double = 0
        '        Dim PBT_Y As Double = 0
        '        Dim PBTM7_L As Double = 0
        '        Dim PBTM8_L As Double = 0
        '        Dim PBTM9_L As Double = 0
        '        Dim PBTM7 As Double = 0
        '        Dim PBTM8 As Double = 0
        '        Dim PBTM9 As Double = 0
        '        Dim PBTQ7 As Double = 0
        '        Dim PBTQ8 As Double = 0
        '        Dim PBTQ9 As Double = 0
        '        Dim PBTY7 As Double = 0
        '        Dim PBTY8 As Double = 0
        '        Dim PBTY9 As Double = 0

        '        '==========12
        '        Dim NP_M_L As Double = 0
        '        Dim NP_M As Double = 0
        '        Dim NP_Q As Double = 0
        '        Dim NP_Y As Double = 0
        '        Dim NPM10_L As Double = 0
        '        Dim NPM11_L As Double = 0
        '        Dim NPM10 As Double = 0
        '        Dim NPM11 As Double = 0
        '        Dim NPQ10 As Double = 0
        '        Dim NPQ11 As Double = 0
        '        Dim NPY10 As Double = 0
        '        Dim NPY11 As Double = 0

        '        Dim RSC As New ADODB.Recordset
        '        LoadSqlData("select * from Ap_Rpt_Income  ", RSC)
        '        With RSC
        '            Do Until .EOF = True

        '                If (.Fields("Grp").Value) = "01" Then
        '                    xm1_L = xm1_L + CDbl((.Fields("Last_Month").Value))
        '                    xm1 = xm1 + CDbl((.Fields("Current_Month").Value))
        '                    xq1 = xq1 + CDbl((.Fields("Quarter_to_Date").Value))
        '                    xy1 = xy1 + CDbl((.Fields("Year_to_Date").Value))
        '                End If
        '                If (.Fields("Grp").Value) = "02" Then
        '                    xm2_L = xm2_L + CDbl((.Fields("Last_Month").Value))
        '                    xm2 = xm2 + CDbl((.Fields("Current_Month").Value))
        '                    xq2 = xq2 + CDbl((.Fields("Quarter_to_Date").Value))
        '                    xy2 = xy2 + CDbl((.Fields("Year_to_Date").Value))
        '                End If

        '                .MoveNext()
        '            Loop
        '        End With
        '        InteIn_M_L = xm1_L - xm2
        '        InteIn_M = xm1 - xm2
        '        InteIn_Q = xq1 - xq2
        '        InteIn_Y = xy1 - xy2
        '        CNN.Execute("Update Ap_Rpt_Income set " & _
        '                       "Last_Month =" & InteIn_M_L & " , " & _
        '                    "Current_Month =" & InteIn_M & " , " & _
        '                    " Quarter_to_Date =" & InteIn_Q & " , " & _
        '                    " Year_to_Date =" & InteIn_Y & " " & _
        '                    "where Rpt_ID='03'")
        '        '===========5
        '        LoadSqlData("select * from Ap_Rpt_Income  ", RSC)
        '        With RSC
        '            Do Until .EOF = True
        '                '==========5
        '                If (.Fields("Grp").Value) = "03" Then
        '                    GOIM3_L = GOIM3_L + CDbl((.Fields("Last_Month").Value))
        '                    GOIM3 = GOIM3 + CDbl((.Fields("Current_Month").Value))
        '                    GOIQ3 = GOIQ3 + CDbl((.Fields("Quarter_to_Date").Value))
        '                    GOIY3 = GOIY3 + CDbl((.Fields("Year_to_Date").Value))
        '                End If
        '                If (.Fields("Grp").Value) = "04" Then
        '                    GOIM4_L = GOIM4_L + CDbl((.Fields("Last_Month").Value))
        '                    GOIM4 = GOIM4 + CDbl((.Fields("Current_Month").Value))
        '                    GOIQ4 = GOIQ4 + CDbl((.Fields("Quarter_to_Date").Value))
        '                    GOIY4 = GOIY4 + CDbl((.Fields("Year_to_Date").Value))
        '                End If
        '                .MoveNext()
        '            Loop
        '        End With
        '        GOI_M_L = GOIM3_L + GOIM4_L
        '        GOI_M = GOIM3 + GOIM4
        '        GOI_Q = GOIQ3 + GOIQ4
        '        GOI_Y = GOIY3 + GOIY4
        '        CNN.Execute("Update Ap_Rpt_Income set " & _
        '                      "Last_Month =" & GOI_M_L & " , " & _
        '                    "Current_Month =" & GOI_M & " , " & _
        '                    " Quarter_to_Date =" & GOI_Q & " , " & _
        '                    " Year_to_Date =" & GOI_Y & " " & _
        '                    "where Rpt_ID='05'")

        '        '===========7



        '        LoadSqlData("select sum(Last_Month) As Last_Month , sum(Current_Month) As Current_Month , sum(Quarter_to_Date) As Quarter_to_Date  , sum(Year_to_Date) As Year_to_Date  from Ap_Rpt_Income where Rpt_ID Like '06.05%' ", RSC)
        '        If RSC.RecordCount <> 0 Then
        '            CNN.Execute("Update Ap_Rpt_Income set Last_Month=" & CDbl((RSC.Fields("Last_Month").Value)) & " , Current_Month=" & CDbl((RSC.Fields("Current_Month").Value)) & " , Quarter_to_Date=" & CDbl((RSC.Fields("Quarter_to_Date").Value)) & " , Year_to_Date=" & CDbl((RSC.Fields("Year_to_Date").Value)) & " where Rpt_ID='06.05'")
        '        End If

        '        '===========7


        '        LoadSqlData("select * from Ap_Rpt_Income  ", RSC)
        '        With RSC
        '            Do Until .EOF = True

        '                '==========7
        '                If (.Fields("Grp").Value) = "05" Then
        '                    NIM5_L = NIM5_L + CDbl((.Fields("Last_Month").Value))
        '                    NIM5 = NIM5 + CDbl((.Fields("Current_Month").Value))
        '                    NIQ5 = NIQ5 + CDbl((.Fields("Quarter_to_Date").Value))
        '                    NIY5 = NIY5 + CDbl((.Fields("Year_to_Date").Value))
        '                End If
        '                If (.Fields("Rpt_Id").Value) = "06.01" Or (.Fields("Rpt_Id").Value) = "06.02" Or (.Fields("Rpt_Id").Value) = "06.03" Or (.Fields("Rpt_Id").Value) = "06.04" Or (.Fields("Rpt_Id").Value) = "06.05" Or (.Fields("Rpt_Id").Value) = "06.06" Or (.Fields("Rpt_Id").Value) = "06.07" Or (.Fields("Rpt_Id").Value) = "06.08" Then

        '                    'MsgBox(NIM6)
        '                    NIM6_L = NIM6_L + CDbl((.Fields("Last_Month").Value))
        '                    NIM6 = NIM6 + CDbl((.Fields("Current_Month").Value))
        '                    NIQ6 = NIQ6 + CDbl((.Fields("Quarter_to_Date").Value))
        '                    NIY6 = NIY6 + CDbl((.Fields("Year_to_Date").Value))
        '                End If


        '                .MoveNext()
        '            Loop
        '        End With
        '        NI_M_L = NIM5_L - NIM6_L
        '        NI_M = NIM5 - NIM6
        '        NI_Q = NIQ5 - NIQ6
        '        NI_Y = NIY5 - NIY6

        '        CNN.Execute("Update Ap_Rpt_Income set " & _
        '                      "Last_Month =" & NI_M_L & " , " & _
        '                    "Current_Month =" & NI_M & " , " & _
        '                    " Quarter_to_Date =" & NI_Q & " , " & _
        '                    " Year_to_Date =" & NI_Y & " " & _
        '                    "where Rpt_ID='07'")

        '        '===========10
        '        LoadSqlData("select * from Ap_Rpt_Income  ", RSC)
        '        With RSC
        '            Do Until .EOF = True
        '                '==========10
        '                If (.Fields("Grp").Value) = "07" Then
        '                    PBTM7_L = PBTM7_L + CDbl((.Fields("Last_Month").Value))
        '                    PBTM7 = PBTM7 + CDbl((.Fields("Current_Month").Value))
        '                    PBTQ7 = PBTQ7 + CDbl((.Fields("Quarter_to_Date").Value))
        '                    PBTY7 = PBTY7 + CDbl((.Fields("Year_to_Date").Value))
        '                End If
        '                If (.Fields("Grp").Value) = "08" Then
        '                    PBTM8_L = PBTM8_L + CDbl((.Fields("Last_Month").Value))
        '                    PBTM8 = PBTM8 + CDbl((.Fields("Current_Month").Value))
        '                    PBTQ8 = PBTQ8 + CDbl((.Fields("Quarter_to_Date").Value))
        '                    PBTY8 = PBTY8 + CDbl((.Fields("Year_to_Date").Value))
        '                End If
        '                If (.Fields("Grp").Value) = "09" Then
        '                    PBTM9_L = PBTM9_L + CDbl((.Fields("Last_Month").Value))
        '                    PBTM9 = PBTM9 + CDbl((.Fields("Current_Month").Value))
        '                    PBTQ9 = PBTQ9 + CDbl((.Fields("Quarter_to_Date").Value))
        '                    PBTY9 = PBTY9 + CDbl((.Fields("Year_to_Date").Value))
        '                End If
        '                .MoveNext()
        '            Loop
        '        End With
        '        PBT_M_L = CDbl(PBTM7_L - PBTM8_L) + PBTM9_L
        '        PBT_M = CDbl(PBTM7 - PBTM8) + PBTM9
        '        PBT_Q = CDbl(PBTQ7 - PBTQ8) + PBTQ9
        '        PBT_Y = CDbl(PBTY7 - PBTY8) + PBTY9
        '        CNN.Execute("Update Ap_Rpt_Income set " & _
        '                         "Last_Month =" & PBT_M_L & " , " & _
        '                    "Current_Month =" & PBT_M & " , " & _
        '                    " Quarter_to_Date =" & PBT_Q & " , " & _
        '                    " Year_to_Date =" & PBT_Y & " " & _
        '                    "where Rpt_ID='10'")
        '        '===========12
        '        LoadSqlData("select * from Ap_Rpt_Income  ", RSC)
        '        With RSC
        '            Do Until .EOF = True
        '                '==========12
        '                If (.Fields("Grp").Value) = "10" Then
        '                    NPM10_L = NPM10_L + CDbl((.Fields("Last_Month").Value))
        '                    NPM10 = NPM10 + CDbl((.Fields("Current_Month").Value))
        '                    NPQ10 = NPQ10 + CDbl((.Fields("Quarter_to_Date").Value))
        '                    NPY10 = NPY10 + CDbl((.Fields("Year_to_Date").Value))
        '                End If
        '                If (.Fields("Grp").Value) = "11" Then
        '                    NPM11_L = NPM11_L + CDbl((.Fields("Last_Month").Value))
        '                    NPM11 = NPM11 + CDbl((.Fields("Current_Month").Value))
        '                    NPQ11 = NPQ11 + CDbl((.Fields("Quarter_to_Date").Value))
        '                    NPY11 = NPY11 + CDbl((.Fields("Year_to_Date").Value))
        '                End If
        '                .MoveNext()
        '            Loop
        '        End With
        '        NP_M_L = NPM10_L - NPM11_L
        '        NP_M = NPM10 - NPM11
        '        NP_Q = NPQ10 - NPQ11
        '        NP_Y = NPY10 - NPY11
        '        CNN.Execute("Update Ap_Rpt_Income set " & _
        '                          "Last_Month =" & NP_M_L & " , " & _
        '                    "Current_Month =" & NP_M & " , " & _
        '                    " Quarter_to_Date =" & NP_Q & " , " & _
        '                    " Year_to_Date =" & NP_Y & " " & _
        '                    "where Rpt_ID='12'")
        '        If Format(MdStartDate, "MM") = "01" Then
        '            CNN.Execute("Update Ap_Rpt_Income set Last_Month = 0 ")

        '        End If









        '        '==============


        '        LoadSqlData("SELECT    Left( Rpt_ID , 2) as Rpt_ID , sum( Last_Month) As Last_Month  , Sum(Current_Month) as Current_Month " & _
        '" , sum( Quarter_to_Date) As Quarter_to_Date  , sum( Year_to_Date) As Year_to_Date " & _
        '" FROM         Ap_Rpt_Income where (  Rpt_ID<>'06.05' and  left( Rpt_ID,2)<>'03' and  left( Rpt_ID,2)<>'05' and  left( Rpt_ID,2)<>'07' and  left( Rpt_ID,2)<>'10' and  left( Rpt_ID,2)<>'11' and  left( Rpt_ID,2)<>'12' and  left( Rpt_ID,2)<>'13' and  left( Rpt_ID,2)<>'14')   " & _
        '        " group BY Left( Rpt_ID , 2) ", RSC)
        '        With RSC
        '            Do Until .EOF = True
        '                CNN.Execute(" Update  Ap_Rpt_Income set " & _
        '                            " Last_Month = " & CDbl((.Fields("Last_Month").Value)) & " , " & _
        '                                " Current_Month = " & CDbl((.Fields("Current_Month").Value)) & " , " & _
        '                                    " Quarter_to_Date = " & CDbl((.Fields("Quarter_to_Date").Value)) & " , " & _
        '                                      " Year_to_Date = " & CDbl((.Fields("Year_to_Date").Value)) & "  " & _
        '                                                               " Where Rpt_ID = '" & CStr((.Fields("Rpt_ID").Value)) & "' ")
        '                'MsgBox(CStr((.Fields("Rpt_ID").Value)))

        '                .MoveNext()
        '            Loop
        '        End With





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
        LoadSqlData("select Rpt_ID         , sum(Amt_Last_M_Dr) As Amt_Last_M_Dr , sum(Amt_Last_M_Cr) As Amt_Last_M_Cr                      , sum(Amt_M_Dr) As Amt_M_Dr , sum(Amt_M_Cr) As Amt_M_Cr  , sum(Amt_Q_Dr) As Amt_Q_Dr ,  sum(Amt_Q_Cr) As Amt_Q_Cr , sum(Amt_y_Dr) As Amt_y_Dr , sum(Amt_y_Cr) As Amt_y_Cr  from Ap_Rpt_Income_Item  where  Rpt_Type = 'Out' group by Rpt_ID ", RSC)
        With RSC
            Do Until .EOF = True
                'MsgBox(CDbl((.Fields("Amt_M_Cr").Value)))
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


    Private Sub SelectOut()

        LoadSqlData("select * from Ap_Rpt_Income_Item where  Rpt_Type = 'Out'", RSCIn_M)
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
        If DMonth.Text = "ມັງກອນ" Then
            MdStartDate = Format(CDate("01/01/" & Year(Myy.Value)), "dd-MM-yyyy")
            MdToDate = Format(CDate("31/01/" & Year(Myy.Value)), "dd-MM-yyyy")
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
        MdToDate = Format(CDate("31/12/" & Year(Toyy.Value)), "dd-MM-yyyy")
        Lb.Text = "ປະຈຳປີ " & yy.Text
        L5.Text = MdStartDate & " => " & MdToDate
    End Sub
    Private Sub LoadReport()
        Dim RPT_ID As String
        RPT_ID = " "
        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()

            RptNme = "ໃບສະຫລຸບຊັບສົມບັດ(ຊັບສິນ-ໜີ້ສິນ)"

            Dim ny, ly, n_L_y As String
            ny = CDbl(Year(MdStartDate))
            ly = CDbl(Year(MdStartDate)) - 1
            'MsgBox(CDbl(Year(MdStartDate)))
            n_L_y = " N'" & ny & "' As Now_Year , N'" & ly & "' As Last_Year ,  "
            .Open("SELECT " & n_L_y & " N'" & RptNme & "' As Rpt_Name , N'" & Lb.Text & "' As    RptSjUd ," & RptSjOff & " * FROM Ap_Rpt_Income where grp<>'' order by Rpt_Id asc  ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)


            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With
        Dim FrmPreview As New FmPreview : FrmClosing()
        'Dim Rpt As New CryLOGO
        Dim Rpt As New CryRpt_Income
        Rpt.SetDataSource(Rs)
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.DisplayGroupTree = False
        'FrmPreview.MdiParent = FmMain
        FrmPreview.WindowState = FormWindowState.Maximized
        FrmPreview.Show()
        FrmPreview.Focus()
    End Sub

    Private Sub LoadReportItem()


        Dim Rs As New ADODB.Recordset
        With Rs
            If .State = ConnectionState.Open Then .Close()
            .Open("SELECT N'" & Lb.Text & "' As    RptSjUd ," & RptSjOff & " *  FROM Ap_Rpt_Income_Item where    Current_Month <>0  or Quarter_to_Date <>0  Year_to_Date <>0  ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
            If .EOF Then Exit Sub
        End With

        Dim FrmPreview As New FmPreview : FrmClosing()
        Dim Rpt As New CryRpt_Income_Itemxx
        Rpt.SetDataSource(Rs)
        FrmPreview.ReportViewer.ReportSource = Rpt
        FrmPreview.ReportViewer.DisplayGroupTree = False
        FrmPreview.MdiParent = FmMain
        FrmPreview.WindowState = FormWindowState.Maximized
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
        RD.Checked = True
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
        Call SelectLoad()
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
End Class