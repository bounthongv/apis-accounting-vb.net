Public Class FmRptPro_New
    Dim MuBTW, MuLFTD As String
    Dim CLT_Str, CLT_Last_Str As String
    Dim hh As String = 0
    Dim RSCIn_M As New ADODB.Recordset
    Private Sub HeaDer()
        LoadSqlData("SELECT * FROM Header WHERE ID=N'B09' ", RSC)
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
            LoadSqlData("SELECT * FROM Header WHERE ID=N'B09' ", RSC)
            If RSC.RecordCount = 0 Then
                CNN.Execute("INSERT INTO Header(ID,Nm,S1,S2,S3,S4,PP) " & _
                            " values('B09',N'" & TxtHeader.Text & "',N'" & TxtS1.Text & "',N'" & TxtS2.Text & "',N'" & TxtS3.Text & "',N'" & TxtS4.Text & "',N'" & TxtPP.Text & "') ")
            Else
                CNN.Execute("UPDATE Header set Nm=N'" & TxtHeader.Text & "',S1=N'" & TxtS1.Text & "',S2=N'" & TxtS2.Text & "',S3=N'" & TxtS3.Text & "',S4=N'" & TxtS4.Text & "',PP=N'" & TxtPP.Text & "' " & _
                            " where ID='B09' ")
            End If
        Else
            LoadSqlData("SELECT * FROM Header WHERE ID=N'B09' ", RSC)
            If RSC.RecordCount = 0 Then
                CNN.Execute("INSERT INTO Header(ID,Nm,S1e,S2e,S3e,S4e,PPe) " & _
                            " values('B09',N'" & TxtHeader.Text & "',N'" & TxtS1.Text & "',N'" & TxtS2.Text & "',N'" & TxtS3.Text & "',N'" & TxtS4.Text & "',N'" & TxtPP.Text & "') ")
            Else
                CNN.Execute("UPDATE Header set Nm=N'" & TxtHeader.Text & "',S1e=N'" & TxtS1.Text & "',S2e=N'" & TxtS2.Text & "',S3e=N'" & TxtS3.Text & "',S4e=N'" & TxtS4.Text & "',PPe=N'" & TxtPP.Text & "' " & _
                            " where ID='B09' ")
            End If
        End If

    End Sub
    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Close()
    End Sub
    Private Sub OptDuring_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RD.CheckedChanged
        SelectDate()
    End Sub
    Private Sub SelectDate()
        Dim s1, s2, s3 As String
        Dim s As Integer = 1
        Ds.Enabled = False
        Dt.Enabled = False
        Period.Enabled = False
        Pyy.Enabled = False
        M2.Enabled = False
        M1.Enabled = False
        Mtoyy.Enabled = False
        Ct.Enabled = False
        yyt.Enabled = False
        yy.Enabled = False
        Toyy.Enabled = False
        If RD.Checked = True Then
            Ds.Enabled = True
            Dt.Enabled = True
            MdStartDate = Ds.Value
            MdToDate = Dt.Value

            If Ds.Value = Dt.Value Then
                LngId = 2069 : CallLngStr()
                L1.Text = LngStr & " " & Ds.Text
            Else
                LngId = 2070 : CallLngStr() : s1 = LngStr
                LngId = 2054 : CallLngStr() : s2 = LngStr
                L1.Text = s1 & " " & Format(MdStartDate, "dd/MM/yyyy") & " " & s2 & " " & Format(MdToDate, "dd/MM/yyyy")
                L1.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
                DTDATE02 = "ປະຈຳວັນທີ"
            End If
        ElseIf RMto.Checked = True Then
            Dim sx As Integer = M1.SelectedIndex + 1
            Dim x1 As Integer = M2.SelectedIndex + 1
            MdStartDate = Format(CDate("01/" & sx & "/" & Year(Mtoyy.Value)), "dd/MM/yyyy")
            MdToDate = Format(CDate("01/" & x1 & "/" & Year(Mtoyy.Value)), "dd/MM/yyyy")
            Dim SM As Date = DateAdd("M", CDbl(1), MdToDate)
            Dim SD As Date = DateAdd("d", CDbl(-1), SM)
            MdToDate = Format(CDate(SD), "dd-MM-yyyy")
            If M2.SelectedIndex = -1 Then
                LoadDateNow()
            End If
            If M1.SelectedIndex = M2.SelectedIndex Then

                LngId = 2071 : CallLngStr() : s1 = LngStr
                LngId = 2049 : CallLngStr() : s3 = LngStr
                L1.Text = s1 & " " & M1.Text & "/" & Format(MdStartDate, "yyyy")
                Dim DTMonth1, DTMonth2 As String
                DTMonth1 = Format(MdStartDate, "MM")
                If Format(MdStartDate, "MM") = "01" Then
                    DTMonth1 = "ມັງກອນ"
                ElseIf Format(MdStartDate, "MM") = "02" Then
                    DTMonth1 = "ກຸມພາ"
                ElseIf Format(MdStartDate, "MM") = "03" Then
                    DTMonth1 = "ມີນາ"
                ElseIf Format(MdStartDate, "MM") = "04" Then
                    DTMonth1 = "ເມສາ"
                ElseIf Format(MdStartDate, "MM") = "05" Then
                    DTMonth1 = "ພຶດສະພາ"
                ElseIf Format(MdStartDate, "MM") = "06" Then
                    DTMonth1 = "ມີຖຸນາ"
                ElseIf Format(MdStartDate, "MM") = "07" Then
                    DTMonth1 = "ກໍລະກົດ"
                ElseIf Format(MdStartDate, "MM") = "08" Then
                    DTMonth1 = "ສິງຫາ"
                ElseIf Format(MdStartDate, "MM") = "09" Then
                    DTMonth1 = "ກັນຍາ"
                ElseIf Format(MdStartDate, "MM") = "10" Then
                    DTMonth1 = "ຕຸລາ"
                ElseIf Format(MdStartDate, "MM") = "11" Then
                    DTMonth1 = "ພະຈິກ"
                ElseIf Format(MdStartDate, "MM") = "12" Then
                    DTMonth1 = "ທັນວາ"
                End If
                L1.Text = s1 & "  " & DTMonth1 & "  ປີ " & Format(MdStartDate, "yyyy")
            Else
                LngId = 2072 : CallLngStr() : s1 = LngStr
                LngId = 2054 : CallLngStr() : s2 = LngStr
                LngId = 2049 : CallLngStr() : s3 = LngStr
                L1.Text = s1 & " " & Format(MdStartDate, "MM") & " " & s2 & " " & Format(MdToDate, "MM") & "/" & Format(Mtoyy.Value, "yyyy")
                Dim DTMonth1, DTMonth2 As String
                DTMonth1 = Format(MdStartDate, "MM")
                If Format(MdStartDate, "MM") = "01" Then
                    DTMonth1 = "ມັງກອນ"
                ElseIf Format(MdStartDate, "MM") = "02" Then
                    DTMonth1 = "ກຸມພາ"
                ElseIf Format(MdStartDate, "MM") = "03" Then
                    DTMonth1 = "ມີນາ"
                ElseIf Format(MdStartDate, "MM") = "04" Then
                    DTMonth1 = "ເມສາ"
                ElseIf Format(MdStartDate, "MM") = "05" Then
                    DTMonth1 = "ພຶດສະພາ"
                ElseIf Format(MdStartDate, "MM") = "06" Then
                    DTMonth1 = "ມີຖຸນາ"
                ElseIf Format(MdStartDate, "MM") = "07" Then
                    DTMonth1 = "ກໍລະກົດ"
                ElseIf Format(MdStartDate, "MM") = "08" Then
                    DTMonth1 = "ສິງຫາ"
                ElseIf Format(MdStartDate, "MM") = "09" Then
                    DTMonth1 = "ກັນຍາ"
                ElseIf Format(MdStartDate, "MM") = "10" Then
                    DTMonth1 = "ຕຸລາ"
                ElseIf Format(MdStartDate, "MM") = "11" Then
                    DTMonth1 = "ພະຈິກ"
                ElseIf Format(MdStartDate, "MM") = "12" Then
                    DTMonth1 = "ທັນວາ"
                End If
                If Format(MdToDate, "MM") = "01" Then
                    DTMonth2 = "ມັງກອນ"
                ElseIf Format(MdToDate, "MM") = "02" Then
                    DTMonth2 = "ກຸມພາ"
                ElseIf Format(MdToDate, "MM") = "03" Then
                    DTMonth2 = "ມີນາ"
                ElseIf Format(MdToDate, "MM") = "04" Then
                    DTMonth2 = "ເມສາ"
                ElseIf Format(MdToDate, "MM") = "05" Then
                    DTMonth2 = "ພຶດສະພາ"
                ElseIf Format(MdToDate, "MM") = "06" Then
                    DTMonth2 = "ມີຖຸນາ"
                ElseIf Format(MdToDate, "MM") = "07" Then
                    DTMonth2 = "ກໍລະກົດ"
                ElseIf Format(MdToDate, "MM") = "08" Then
                    DTMonth2 = "ສິງຫາ"
                ElseIf Format(MdToDate, "MM") = "09" Then
                    DTMonth2 = "ກັນຍາ"
                ElseIf Format(MdToDate, "MM") = "10" Then
                    DTMonth2 = "ຕຸລາ"
                ElseIf Format(MdToDate, "MM") = "11" Then
                    DTMonth2 = "ພະຈິກ"
                ElseIf Format(MdToDate, "MM") = "12" Then
                    DTMonth2 = "ທັນວາ"
                End If
                L1.Text = s1 & "  " & DTMonth1 & "  " & s2 & " " & DTMonth2 & " ປີ " & Format(Mtoyy.Value, "yyyy")

                L1.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
                DTDATE02 = "ປະຈຳເດືອນ " & DTMonth1 & " ປີ " & Year(MdToDate)

            End If
            M1.Enabled = True
            M2.Enabled = True
            Mtoyy.Enabled = True
        ElseIf RP.Checked = True Then
            Period.Enabled = True
            Pyy.Enabled = True
            s = 1
            If Period.SelectedIndex > 0 Then
                s = Period.SelectedIndex * 3 + 1
            End If
            MdStartDate = Format(CDate("01/" & s & "/" & Year(Pyy.Value)), "dd/MM/yyyy")
            MdToDate = Format(CDate("01/" & s & "/" & Year(Pyy.Value)), "dd/MM/yyyy")
            Dim SM As Date = DateAdd("M", CDbl(3), MdToDate)
            Dim SD As Date = DateAdd("d", CDbl(-1), SM)
            MdToDate = Format(CDate(SD), "dd-MM-yyyy")
            LngId = 2075 : CallLngStr() : s1 = LngStr
            LngId = 2049 : CallLngStr() : s3 = LngStr
            L1.Text = s1 & " " & Period.Text & " " & Format(Pyy.Value, "yyyy")
            L1.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
            DTDATE02 = "ປະຈຳ" & Period.Text & " ປີ " & Year(MdToDate)

        ElseIf RT.Checked = True Then
            Ct.Enabled = True
            yyt.Enabled = True
            s = 1
            LngId = 2073 : CallLngStr() : s1 = LngStr
            L1.Text = s1 & " " & Format(yyt.Value, "yyyy")
            If Ct.SelectedIndex > 0 Then
                s = Ct.SelectedIndex * 6 + 1
                LngId = 2074 : CallLngStr() : s1 = LngStr
                L1.Text = s1 & " " & Format(yyt.Value, "yyyy")
            End If
            MdStartDate = Format(CDate("01/" & s & "/" & Year(yyt.Value)), "dd/MM/yyyy")
            MdToDate = Format(CDate("01/" & s & "/" & Year(yyt.Value)), "dd/MM/yyyy")
            Dim SM As Date = DateAdd("M", CDbl(6), MdToDate)
            Dim SD As Date = DateAdd("d", CDbl(-1), SM)
            MdToDate = Format(CDate(SD), "dd-MM-yyyy")
            L1.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
            DTDATE02 = "ປະຈຳ" & Ct.Text & " ປີ " & Year(MdToDate)
        ElseIf RY.Checked = True Then
            yy.Enabled = True
            Toyy.Enabled = True
            MdStartDate = Format(CDate("01/01/" & Year(yy.Value)), "dd/MM/yyyy")
            MdToDate = Format(CDate("31/12/" & Year(Toyy.Value)), "dd/MM/yyyy")
            If yy.Text = Toyy.Text Then
                LngId = 2077 : CallLngStr() : s1 = LngStr
                L1.Text = s1 & " " & Format(yy.Value, "yyyy")
            Else
                LngId = 2076 : CallLngStr() : s1 = LngStr
                LngId = 2054 : CallLngStr() : s2 = LngStr
                L1.Text = s1 & " " & Format(yy.Value, "yyyy") & "/" & Format(Toyy.Value, "yyyy")
            End If
            L1.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd/MM/yyyy") & " ສີ້ນສຸດວັນທີ " & Format(MdToDate, "dd/MM/yyyy")
            L1.Text = "ແຕ່ວັນທີ " & Format(MdStartDate, "dd") & " / " & Format(MdStartDate, "MM") & " / " & Format(MdStartDate, "yyyy") & " ຫາ " & Format(MdToDate, "dd") & " / " & Format(MdToDate, "MM") & " / " & Format(MdToDate, "yyyy")
            DTDATE02 = "ປະຈຳປີ  " & Year(MdToDate)
        End If
        L2.Text = Format(MdStartDate, "dd/MM/yyyy") & " To " & Format(MdToDate, "dd/MM/yyyy")

    End Sub

    Private Sub RM_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SelectDate()
    End Sub

    Private Sub RP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RP.CheckedChanged
        SelectDate()
    End Sub

    Private Sub RMto_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RMto.CheckedChanged
        SelectDate()
    End Sub

    Private Sub RT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RT.CheckedChanged
        SelectDate()
    End Sub

    Private Sub RY_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RY.CheckedChanged
        SelectDate()
    End Sub

    Private Sub Ds_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ds.ValueChanged
        Dt.Value = Ds.Value
        SelectDate()
    End Sub

    Private Sub Dt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Dt.ValueChanged
        SelectDate()
    End Sub

    Private Sub DMonth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SelectDate()
    End Sub

    Private Sub Myy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SelectDate()
    End Sub

    Private Sub Period_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Period.SelectedIndexChanged
        SelectDate()
    End Sub

    Private Sub Pyy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Pyy.ValueChanged
        SelectDate()
    End Sub

    Private Sub M1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles M1.SelectedIndexChanged

        M2.SelectedIndex = M1.SelectedIndex

        SelectDate()
    End Sub

    Private Sub M2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles M2.SelectedIndexChanged
        If M2.SelectedIndex < M1.SelectedIndex Then
            M1.SelectedIndex = M2.SelectedIndex
        End If
        SelectDate()
    End Sub

    Private Sub Mtoyy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Mtoyy.ValueChanged
        SelectDate()
    End Sub

    Private Sub Ct_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ct.SelectedIndexChanged
        SelectDate()
    End Sub

    Private Sub yyt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yyt.ValueChanged
        SelectDate()
    End Sub

    Private Sub yy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yy.ValueChanged
        Toyy.Value = yy.Value
        SelectDate()

    End Sub

    Private Sub Toyy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Toyy.ValueChanged
        SelectDate()
    End Sub

    Private Sub BtnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        AddHeader()
        SelectDate()
        Call Office()
        '==============
        Dim MUTY2 As String = MUTY
        MUTY = "PRO4"
        Cacu()
        Update_Sum()

        If MUTY = "PRO4" Then
            CNN.Execute(" update  So_Rpt_Pro set Amtb=0 where RptID = '28' ")
            CNN.Execute(" update So_Rpt_Pro set Amtb = (amt)*AmtA/100 ")
            CNN.Execute(" update  So_Rpt_Pro set Amtb= (select SUM((amt)*AmtA/100) from So_Rpt_Pro where  RptType = 'PRO4') where RptType = 'PRO4' and RptID = '28' ")
            CNN.Execute(" update  So_Rpt_Pro set Amtb= (select SUM(amtb) from So_Rpt_Pro where  RptType = 'PRO4' And ( RptID = '09' or RptID = '08' or RptID = '10') ) where RptType = 'PRO4' and RptID = '07' ")
            CNN.Execute(" update  So_Rpt_Pro set Amtb= (select SUM(amtb) from So_Rpt_Pro where  RptType = 'PRO4' And ( RptID = '12' or RptID = '13' or RptID = '14'  or RptID = '15'  or RptID = '15') ) where RptType = 'PRO4' and RptID = '11' ")
            CNN.Execute(" update  So_Rpt_ProHelp set Amt= (select SUM((amt)*AmtA/100) from So_Rpt_Pro)   ")
            CNN.Execute(" update  So_Rpt_Pro set Amtb= (select SUM(amtb) from So_Rpt_Pro where  RptType = 'PRO4' And ( RptID = '11' or RptID = '07'   ) ) where RptType = 'PRO4' and RptID = '01' ")
        Else
        End If
        MUTY = MUTY2
        Cacu()

        CNN.Execute(" update   So_Rpt_Pro set Amt= (select amt  from So_Rpt_ProHelp ) where RptType = 'PRO1' and RptID = '02' ")
        Update_Sum()
        'CNN.Execute(" update  So_Rpt_Pro set Amtb=0 where RptID = '28' ")
        'CNN.Execute(" update So_Rpt_Pro set Amtb = (amt)*AmtA/100 ")
        'CNN.Execute(" update  So_Rpt_Pro set Amtb= (select SUM((amt)*AmtA/100) from So_Rpt_Pro where  RptType = 'PRO4') where RptType = 'PRO4' and RptID = '28' ")


        'CNN.Execute(" update  So_Rpt_ProHelp set Amt= (select SUM((amt)*AmtA/100) from So_Rpt_Pro)   ")

        '=============
        If MUTY = "PRO1" Then
            'CNN.Execute(" update  So_Rpt_Pro set Amt= ((select amt  from  So_Rpt_Pro where RptType = 'PRO1' and RptID = '07' )/(select amt  from So_Rpt_Pro where RptType = 'PRO1' and RptID = '15' ))*100  where RptType = 'PRO1' and RptID = '16' ")
            Dim KK As String = "  update  So_Rpt_Pro set Amt= ((select amt  from  So_Rpt_Pro where RptType = 'PRO1' and RptID = '07' and amt<>0 )/(select amt  from So_Rpt_Pro where RptType = 'PRO1' and RptID = '15' and amt<>0 ))*100  where RptType = 'PRO1' and RptID = '16' "
            CNN.Execute(KK)
        End If
        If MUTY = "PRO6" Then
            'CNN.Execute(" update  So_Rpt_Pro set Amt= ((select amt  from  So_Rpt_Pro where RptType = 'PRO1' and RptID = '07' )/(select amt  from So_Rpt_Pro where RptType = 'PRO1' and RptID = '15' ))*100  where RptType = 'PRO1' and RptID = '16' ")
            'Dim KK As String = "  update  So_Rpt_Pro set Amt= ((select amt  from  So_Rpt_Pro where RptType = 'PRO6' and RptID = '07' and amt<>0 )/(select amt  from So_Rpt_Pro where RptType = 'PRO6' and RptID = '15' and amt<>0 ))*100  where RptType = 'PRO6' and RptID = '16' "
            'CNN.Execute(KK)
            '=========I=======
            'CNN.Execute("  update  So_Rpt_Pro set Amt= ((select amt  from  So_Rpt_Pro where RptType = 'PRO6' and RptID = '03' )/(select amt  from So_Rpt_Pro where RptType = 'PRO6' and RptID = '03' ))  where RptType = 'PRO6' and RptID = '02' and amt <>0 ")
            '=========I=======
            'CNN.Execute("  update  So_Rpt_Pro set Amt= ((select amt  from  So_Rpt_Pro where RptType = 'PRO6' and RptID = '03' )/(select amt  from So_Rpt_Pro where RptType = 'PRO6' and RptID = '03' ))  where RptType = 'PRO6' and RptID = '06' and amt <>0 ")
            Call PEO_BLS06()
            Call PEO_InCom06()
            '=======InCom06()==
            CNN.Execute("  Update So_Rpt_Pro set amt=(select amt from Ap_Rpt_Income_Old where Rpt_ID='1.2.4.3') where RptID='03' and  (RptType = 'PRO6') ")
            CNN.Execute(" Update So_Rpt_Pro set amt=(select amt from Ap_Rpt_Income_Old where Rpt_ID='1.2.4.3') where RptID='07' and  (RptType = 'PRO6') ")
            CNN.Execute(" Update So_Rpt_Pro set amt=(select amt from Ap_Rpt_Income_Old where Rpt_ID='1.13') where RptID='08' and  (RptType = 'PRO6') ")
            '=======BLS06()==
            CNN.Execute("  Update So_Rpt_Pro set amt=(select amt from Ap_Rpt_BLS_Old where Rpt_ID='1.5') where RptID='04' and  (RptType = 'PRO6') ")
            CNN.Execute("  Update So_Rpt_Pro set amt=(select amt from Ap_Rpt_BLS_Old where Rpt_ID='1.5') where RptID='09' and  (RptType = 'PRO6') ")
            CNN.Execute("  Update So_Rpt_Pro set amt=(select amt from Ap_Rpt_BLS_Old where Rpt_ID='5') where RptID='10' and  (RptType = 'PRO6') ")
            CNN.Execute("  Update So_Rpt_Pro set amt=(select amt from Ap_Rpt_BLS_Old where Rpt_ID='2.1') where RptID='13' and  (RptType = 'PRO6') ")
            CNN.Execute("  Update So_Rpt_Pro set amt=(select sum(amt) as amt from Ap_Rpt_BLS_Old where (Rpt_ID='2.2' or Rpt_ID='2.3')) where RptID='14' and  (RptType = 'PRO6') ")
            CNN.Execute("  Update So_Rpt_Pro set amt=(select amt from Ap_Rpt_BLS_Old where Rpt_ID='5') where RptID='15' and  (RptType = 'PRO6') ")
            '===ໜີ້ສິນລວມ
            CNN.Execute("  Update So_Rpt_Pro set amt=(select amt from Ap_Rpt_BLS_Old where Rpt_ID='3') where RptID='18' and  (RptType = 'PRO6') ")
            '===ຊັບສິນລວມ
            CNN.Execute("  Update So_Rpt_Pro set amt=(select amt from Ap_Rpt_BLS_Old where Rpt_ID='1.5') where RptID='19' and  (RptType = 'PRO6') ")

            '===ອັດຕາຜົນຕອບແທນຈາກຊັບສິນທັງໝົດ (ROA)
            Dim Last As String = "delete AP_Sum "
            Last = Last & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select  '03'  ,amt,0 from So_Rpt_Pro     where rptID='03' and  (RptType = 'PRO6')  "
            Last = Last & "   insert into AP_Sum(Rpt_ID,amt2,amt3) select  '04'  ,0 ,amt from So_Rpt_Pro     where rptID='04'   and  (RptType = 'PRO6')   "
            Last = Last & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select  '02'  ,0 ,amt from So_Rpt_Pro     where rptID='02'   and  (RptType = 'PRO6')   "
            Last = Last & "  update AP_Sum set amt=(select sum(amt2)/sum(amt3) from AP_Sum where (rpt_ID='03' or rpt_ID='04')  ) where rpt_ID='02'     "
            Last = Last & "    delete  AP_Sum  where Rpt_ID<>'02'  "
            Last = Last & " Update So_Rpt_Pro set So_Rpt_Pro.amt=AP_Sum.amt from AP_Sum,So_Rpt_Pro    where So_Rpt_Pro.rptID=AP_Sum.rpt_ID   and  (So_Rpt_Pro.RptType = 'PRO6')   "
            CNN.Execute(Last)
            '====ອັດຕາຜົນຕອບແທນຈາກທຶນ (ROE)===========
            Dim ROE As String = "delete AP_Sum "
            ROE = ROE & " insert into AP_Sum(Rpt_ID,amt1,amt2,amt3,amt4) select  '07'  ,amt,0,0,0 from So_Rpt_Pro     where rptID='07' and  (RptType = 'PRO6')    "
            ROE = ROE & "     insert into AP_Sum(Rpt_ID,amt1,amt2,amt3,amt4) select  '08'  ,0 ,amt,0,0 from So_Rpt_Pro     where rptID='08'   and  (RptType = 'PRO6')     "
            ROE = ROE & "  insert into AP_Sum(Rpt_ID,amt1,amt2,amt3,amt4) select  '09'  ,0,0,amt,0 from So_Rpt_Pro     where rptID='09' and  (RptType = 'PRO6')    "
            ROE = ROE & "   insert into AP_Sum(Rpt_ID,amt1,amt2,amt3,amt4) select  '10'  ,0 ,0,0 ,amt from So_Rpt_Pro     where rptID='10'   and  (RptType = 'PRO6')   "
            ROE = ROE & "     insert into AP_Sum(Rpt_ID,amt1,amt2,amt3,amt4) select  '06'  ,0 ,0,0,0 from So_Rpt_Pro     where rptID='06'   and  (RptType = 'PRO6')      "
            ROE = ROE & " update AP_Sum set amt=(select sum(amt1)/sum(amt2)*sum(amt2)/sum(amt3) *sum(amt3)/sum(amt4) from AP_Sum where (rpt_ID='07' or rpt_ID='08'  or rpt_ID='09'  or rpt_ID='10' )  ) where rpt_ID='06'        "
            ROE = ROE & "    delete  AP_Sum  where Rpt_ID<>'06'  "
            ROE = ROE & " Update So_Rpt_Pro set So_Rpt_Pro.amt=AP_Sum.amt from AP_Sum,So_Rpt_Pro    where So_Rpt_Pro.rptID=AP_Sum.rpt_ID   and  (So_Rpt_Pro.RptType = 'PRO6')   "
            CNN.Execute(ROE)

            '===ອັດຕາສ່ວນໜີ້ສິນຕໍ່ສ່ວນຂອງຜູ້ຖືຮຸ້ນ (Debt to Equity Ratio)
            Dim Debt As String = "delete AP_Sum "
            Debt = Debt & " insert into AP_Sum(Rpt_ID,amt1,amt2,amt3,amt4) select  '13'  ,amt,0,0,0 from So_Rpt_Pro     where rptID='13' and  (RptType = 'PRO6')    "
            Debt = Debt & "     insert into AP_Sum(Rpt_ID,amt1,amt2,amt3,amt4) select  '14'  ,0 ,amt,0,0 from So_Rpt_Pro     where rptID='14'   and  (RptType = 'PRO6')     "
            Debt = Debt & "  insert into AP_Sum(Rpt_ID,amt1,amt2,amt3,amt4) select  '15'  ,0,0,amt,0 from So_Rpt_Pro     where rptID='15' and  (RptType = 'PRO6')    "
            Debt = Debt & "     insert into AP_Sum(Rpt_ID,amt1,amt2,amt3,amt4) select  '12'  ,0 ,0,0,0 from So_Rpt_Pro     where rptID='12'   and  (RptType = 'PRO6')      "
            Debt = Debt & " update AP_Sum set amt=(select (sum(amt1)+sum(amt2))/sum(amt3) from AP_Sum where (rpt_ID='13' or rpt_ID='14'  or rpt_ID='15')  ) where rpt_ID='12'        "
            Debt = Debt & "    delete  AP_Sum  where Rpt_ID<>'12'  "
            Debt = Debt & " Update So_Rpt_Pro set So_Rpt_Pro.amt=AP_Sum.amt from AP_Sum,So_Rpt_Pro    where So_Rpt_Pro.rptID=AP_Sum.rpt_ID   and  (So_Rpt_Pro.RptType = 'PRO6')   "
            CNN.Execute(Debt)

            '===ອັດຕາສ່ວນໜີ້ສິນຕໍ່ຊັບສິນລວມ (Debt to Total Asset Ratio)
            Dim Asset As String = "delete AP_Sum "
            Asset = Asset & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select  '18'  ,amt,0 from So_Rpt_Pro     where rptID='18' and  (RptType = 'PRO6')  "
            Asset = Asset & "   insert into AP_Sum(Rpt_ID,amt2,amt3) select  '19'  ,0 ,amt from So_Rpt_Pro     where rptID='19'   and  (RptType = 'PRO6')   "
            Asset = Asset & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select  '17'  ,0 ,amt from So_Rpt_Pro     where rptID='17'   and  (RptType = 'PRO6')   "
            Asset = Asset & "  update AP_Sum set amt=(select sum(amt2)/sum(amt3) from AP_Sum where (rpt_ID='18' or rpt_ID='19')  ) where rpt_ID='17'     "
            Asset = Asset & "    delete  AP_Sum  where Rpt_ID<>'17'"
            Asset = Asset & " Update So_Rpt_Pro set So_Rpt_Pro.amt=AP_Sum.amt from AP_Sum,So_Rpt_Pro    where So_Rpt_Pro.rptID=AP_Sum.rpt_ID   and  (So_Rpt_Pro.RptType = 'PRO6')   "
            CNN.Execute(Asset)

        End If
        If CheckBox1.Checked = False Then
            Report_Pro1()
        Else
            Report_ProDetail()
        End If

    End Sub
   

    Private Sub Report_Pro1()
        Dim RPT_ID As String
        RPT_ID = " "
        Call Office()
        MuLng = "L"
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & L1.Text & "' As    RptSjUd ,"
        MuLngRpt = MuLngRpt & "N'" & txtRptNme.Text & " " & DTDATE02 & "' As Crl_RptName ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
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
        LngId = "7048" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InMounth ,"
        LoadLoGO()
        Dim s1 As String = ""
        Dim s2 As String = ""
        SqlClient = ""
        SqlClient = vbCrLf & SqlClient & "  update  So_Rpt_Pro set Hid = 0 where RptType='" & MUTY & "' "
        CnnEdit()

        If MUTY = "PRO1" Then
            s2 = "RptID = '16' and  RptType='PRO1'"
            s1 = " (select Amt from So_Rpt_Pro where " & s2 & ") Amtss , ( select Des from So_Rpt_Pro where  " & s2 & ") Dess , "
            SqlClient = vbCrLf & SqlClient & "  update  So_Rpt_Pro set Hid = 1  where   " & s2 & " "
            CnnEdit()
        ElseIf MUTY = "PRO2" Then
            s2 = "RptID = '38' and  RptType='PRO2'"
            s1 = " (select Amt from So_Rpt_Pro where " & s2 & ") Amtss , ( select Des from So_Rpt_Pro where  " & s2 & ") Dess , "
            SqlClient = vbCrLf & SqlClient & "  update  So_Rpt_Pro set Hid = 1  where   " & s2 & " "
            CnnEdit()
        ElseIf MUTY = "PRO3" Then
            s2 = "RptID = '18' and  RptType='PRO3'"
            s1 = " (select Amt from So_Rpt_Pro where " & s2 & ") Amtss , ( select Des from So_Rpt_Pro where  " & s2 & ") Dess , "
            SqlClient = vbCrLf & SqlClient & "  update  So_Rpt_Pro set Hid = 1  where   " & s2 & " "
            CnnEdit()
        ElseIf MUTY = "PRO4" Then
            s2 = "RptID = '28' and  RptType='PRO4'"
            s1 = " (select Amt from So_Rpt_Pro where " & s2 & ") Amtss , ( select Des from So_Rpt_Pro where  " & s2 & ") Dess , "
            SqlClient = vbCrLf & SqlClient & "  update  So_Rpt_Pro set Hid = 1  where   " & s2 & " "
            CnnEdit()
        ElseIf MUTY = "PRO8" Then
            s2 = "RptID = '28' and  RptType='PRO4'"
            s1 = " (select Amt from So_Rpt_Pro where " & s2 & ") Amtss , ( select Des from So_Rpt_Pro where  " & s2 & ") Dess , "
            SqlClient = vbCrLf & SqlClient & "  update  So_Rpt_Pro set Hid = 1  where   " & s2 & " "
            CnnEdit()
        End If

        'CNN.Execute(" update So_Rpt_Pro set amta=amt")

        Dim ds As New DataSet
        Try
            ConnectCL()
            SqlClient = "Select " & s1 & "   " & mformat & "  as mformat  ,   " & MuLngRpt & " *,  N'" & MuOffDep & "'  as RptSjoff_Dep   from So_Rpt_Pro Where RptType = '" & MUTY & "' and Hid = 0  Order by RptId "
            'SqlClient = "Select  " & MuLngRpt & " * from So_Rpt_Pro Order by RptId "
            'SqlClient = "Select  'L' As Crl_Lng , *  from So_Rpt_Pro "
            ConnectCL()
            LoadCN()
            da.Fill(ds, "So_Rpt_Pro")

            If MUTY = "PRO1" Then
                Dim rp As New CryRpt_Pro1
                If MdShowLOGO = 1 Then
                    rp.Subreports(0).SetDataSource(RsLOGO)
                End If
                Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
                myText2 = CType(rp.ReportDefinition.ReportObjects.Item("Text1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = "ຫົວໜ່ວຍ : ກີບ"
                rp.SetDataSource(ds.Tables("So_Rpt_Pro"))
                rp.Refresh()
                Dim FmPreview As New FmPreview : FrmClosing()
                FmPreview.ReportViewer.ReportSource = rp
                FmPreview.ReportViewer.DisplayGroupTree = True
                FmPreview.WindowState = FormWindowState.Maximized
                FmPreview.Show()
            ElseIf MUTY = "PRO2" Then
                Dim rp As New CryRpt_Total_Assets
                If MdShowLOGO = 1 Then
                    rp.Subreports(0).SetDataSource(RsLOGO)
                End If
                Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
                myText2 = CType(rp.ReportDefinition.ReportObjects.Item("Text1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = "ຫົວໜ່ວຍ : ກີບ"
                rp.SetDataSource(ds.Tables("So_Rpt_Pro"))
                rp.Refresh()
                Dim FmPreview As New FmPreview : FrmClosing()
                FmPreview.ReportViewer.ReportSource = rp
                FmPreview.ReportViewer.DisplayGroupTree = False
                FmPreview.WindowState = FormWindowState.Maximized
                FmPreview.Show()
            ElseIf MUTY = "PRO3" Then
                Dim rp As New CryRpt_liabilities
                If MdShowLOGO = 1 Then
                    rp.Subreports(0).SetDataSource(RsLOGO)
                End If
                Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
                myText2 = CType(rp.ReportDefinition.ReportObjects.Item("Text1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = "ຫົວໜ່ວຍ : ກີບ"
                rp.SetDataSource(ds.Tables("So_Rpt_Pro"))
                rp.Refresh()
                Dim FmPreview As New FmPreview : FrmClosing()
                FmPreview.ReportViewer.ReportSource = rp
                FmPreview.ReportViewer.DisplayGroupTree = True
                FmPreview.WindowState = FormWindowState.Maximized
                FmPreview.Show()
            ElseIf MUTY = "PRO4" Then
                Dim rp As New CryRpt_Risk_rate
                If MdShowLOGO = 1 Then
                    rp.Subreports(0).SetDataSource(RsLOGO)
                End If
                Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
                myText2 = CType(rp.ReportDefinition.ReportObjects.Item("Text1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = "ຫົວໜ່ວຍ : ກີບ"
                rp.SetDataSource(ds.Tables("So_Rpt_Pro"))
                rp.Refresh()
                Dim FmPreview As New FmPreview : FrmClosing()
                FmPreview.ReportViewer.ReportSource = rp
                FmPreview.ReportViewer.DisplayGroupTree = True
                FmPreview.WindowState = FormWindowState.Maximized
                FmPreview.Show()
            ElseIf MUTY = "PRO6" Then
                Dim rp As New CryRpt_Pro6
                If MdShowLOGO = 1 Then
                    rp.Subreports(0).SetDataSource(RsLOGO)
                End If
                Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
                myText2 = CType(rp.ReportDefinition.ReportObjects.Item("S1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = TxtS1.Text
                myText2 = CType(rp.ReportDefinition.ReportObjects.Item("S2"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = TxtS2.Text
                myText2 = CType(rp.ReportDefinition.ReportObjects.Item("S3"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = TxtS3.Text
                myText2 = CType(rp.ReportDefinition.ReportObjects.Item("S4"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = TxtS4.Text
                myText2 = CType(rp.ReportDefinition.ReportObjects.Item("txtprint_user"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = MUserName
                myText2 = CType(rp.ReportDefinition.ReportObjects.Item("pp"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = TxtPP.Text

                myText2 = CType(rp.ReportDefinition.ReportObjects.Item("Text1"), CrystalDecisions.CrystalReports.Engine.TextObject)
                myText2.Text = "ຫົວໜ່ວຍ : ກີບ"
                rp.SetDataSource(ds.Tables("So_Rpt_Pro"))
                rp.Refresh()
                Dim FmPreview As New FmPreview : FrmClosing()
                FmPreview.ReportViewer.ReportSource = rp
                FmPreview.ReportViewer.DisplayGroupTree = True
                FmPreview.WindowState = FormWindowState.Maximized
                FmPreview.Show()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub Report_ProDetail()
        Dim RPT_ID As String
        RPT_ID = " "
        Call Office()
        MuLng = "L"
        MuLngRpt = RptSjOff
        MuLngRpt = MuLngRpt & "N'" & MuLng & "' As Crl_Lng,N'" & L1.Text & "' As    RptSjUd ,"
        MuLngRpt = MuLngRpt & "N'" & txtRptNme.Text & "( " & CheckBox1.Text & ") ' As Crl_RptName ,"
        LngId = "7001" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDR ,"
        LngId = "7002" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_LaoPDP ,"
        LngId = "7059" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_PP ,"
        LngId = "7004" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Descrip ,"
        LngId = "7005" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_No ,"
        LngId = "7008" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Dr ,"
        LngId = "7009" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Cr ,"
        LngId = "7011" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Sign ,"
        LngId = "7012" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_SignNme ,"
        LngId = "7026" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_InDate ,"
        LngId = "7031" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_User ,"
        LngId = "7042" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Open ,"
        LngId = "7048" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Period	 ,"
        LngId = "7043" : CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr & "' As Crl_Rem ,"

        Dim ds As New DataSet
        Try
            ConnectCL()
            SqlClient = "   SELECT  " & MuLngRpt & "   dbo.So_Rpt_ProDetail.RptID, dbo.So_Rpt_Pro.Des, dbo.So_Rpt_Pro.DesE, dbo.So_Rpt_ProDetail.AcCode, dbo.Acc_Code.Name_L, dbo.Acc_Code.Name_E,          dbo.So_Rpt_ProDetail.OpenDr, dbo.So_Rpt_ProDetail.OpenCr, dbo.So_Rpt_ProDetail.AmtDr, dbo.So_Rpt_ProDetail.AmtCr    FROM    dbo.So_Rpt_ProDetail LEFT OUTER JOIN        dbo.Acc_Code ON dbo.So_Rpt_ProDetail.AcCode = dbo.Acc_Code.Ac_Code LEFT OUTER JOIN        dbo.So_Rpt_Pro ON dbo.So_Rpt_ProDetail.RptID = dbo.So_Rpt_Pro.RptID where dbo.So_Rpt_Pro.RptType=  '" & MUTY & "'   "

            'SqlClient = "Select  'L' As Crl_Lng , *  from So_Rpt_Pro "
            ConnectCL()
            LoadCN()
            da.Fill(ds, "  from dbo.So_Rpt_ProDetail LEFT OUTER JOIN        dbo.Acc_Code ON dbo.So_Rpt_ProDetail.AcCode = dbo.Acc_Code.Ac_Code LEFT OUTER JOIN        dbo.So_Rpt_Pro ON dbo.So_Rpt_ProDetail.RptID = dbo.So_Rpt_Pro.RptID ")
            Dim rp As New CryRpt_Prodetail
            rp.SetDataSource(ds.Tables("  from dbo.So_Rpt_ProDetail LEFT OUTER JOIN        dbo.Acc_Code ON dbo.So_Rpt_ProDetail.AcCode = dbo.Acc_Code.Ac_Code LEFT OUTER JOIN        dbo.So_Rpt_Pro ON dbo.So_Rpt_ProDetail.RptID = dbo.So_Rpt_Pro.RptID "))
            Dim myText2 As CrystalDecisions.CrystalReports.Engine.TextObject
            myText2 = CType(rp.ReportDefinition.ReportObjects.Item("S1"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = TxtS1.Text
            myText2 = CType(rp.ReportDefinition.ReportObjects.Item("S2"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = TxtS2.Text
            myText2 = CType(rp.ReportDefinition.ReportObjects.Item("S3"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = TxtS3.Text
            myText2 = CType(rp.ReportDefinition.ReportObjects.Item("S4"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = TxtS4.Text

            myText2 = CType(rp.ReportDefinition.ReportObjects.Item("pp"), CrystalDecisions.CrystalReports.Engine.TextObject)
            myText2.Text = TxtPP.Text

            rp.Refresh()
            Dim FmPreview As New FmPreview : FrmClosing()
            FmPreview.ReportViewer.ReportSource = rp
            FmPreview.ReportViewer.DisplayGroupTree = True
            FmPreview.WindowState = FormWindowState.Maximized
            FmPreview.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub Cacu()
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        MuBTW = "BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "'  AND  '" & Format(MdToDate, "yyyy-MM-dd") & "' And RptType = '" & MUTY & "' " & MULook2 & " "
        'MuLFTD = "BETWEEN '" & Format(CDate(DateAdd(DateInterval.Day, -1, MdStartDate)), "yyyy") & "/01/01'  AND  '" & Format(CDate(DateAdd(DateInterval.Day, -1, MdStartDate)), "yyyy-MM-dd") & "'  And RptType = '" & MUTY & "' " & MULook2 & " "

        If Format(MdStartDate, "yyyy-MM-dd") = Format(CDate(DateAdd(DateInterval.Day, -1, MdStartDate)), "yyyy") & "-01-01" Then
            MuLFTD = "BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "'  AND  '" & Format(MdToDate, "yyyy-MM-dd") & "' And RptType = '" & MUTY & "' " & MULook2 & " "
        Else
            MuLFTD = "BETWEEN '" & Format(CDate(DateAdd(DateInterval.Day, -1, MdStartDate)), "yyyy") & "/01/01'  AND  '" & Format(CDate(DateAdd(DateInterval.Day, -1, MdStartDate)), "yyyy-MM-dd") & "'  And RptType = '" & MUTY & "' " & MULook2 & " "
        End If
        SqlClient = ""
        SqlClient = vbCrLf & SqlClient & "    delete So_Rpt_ProUpdate  Update So_Rpt_Pro set   OpenAmt = 0 , Amt=0 Update So_Rpt_Pro set   OpenAmt = 0 , Amt=0  Update So_Rpt_Proitems set OpenAmt=0 , Amt = 0  Delete So_Rpt_ProDetail  "
        SqlClient = vbCrLf & SqlClient & " Insert into So_Rpt_ProDetail (RptID , AcCode, AmtDr, AmtCr) Select  b.RptID , ac_code , isnull(sum(amt_Dr),0) ,isnull(sum(amt_Cr),0)  from gen_jn as a , So_Rpt_Proitems as b   where  a.ac_code=b.AcCode And b.CurrType = 0 And a.Date_Work " & MuBTW & " group by  a.ac_code ,  b.RptID   "
        SqlClient = vbCrLf & SqlClient & " Insert into So_Rpt_ProDetail (RptID , AcCode, OpenDr, OpenCr) Select  b.RptID , ac_code , isnull(sum(amt_Dr),0) ,isnull(sum(amt_Cr),0)  from gen_jn as a , So_Rpt_Proitems as b   where a.ac_code=b.AcCode And b.CurrType = 0 And a.Date_Work " & MuLFTD & " group by  a.ac_code ,  b.RptID  "

        SqlClient = vbCrLf & SqlClient & " Insert into So_Rpt_ProDetail (RptID , AcCode, OpenDr, OpenCr) Select  b.RptID , ac_code , isnull(sum(amt_Dr),0) ,isnull(sum(amt_Cr),0)  from Open_jn as a , So_Rpt_Proitems as b   where a.ac_code=b.AcCode And b.CurrType = 0 And a.Date_Work " & MuLFTD & " group by  a.ac_code ,  b.RptID  "
        SqlClient = vbCrLf & SqlClient & " Insert into So_Rpt_ProDetail (RptID , AcCode, AmtDr, AmtCr) Select  b.RptID , ac_code , isnull(sum(amount_Dr),0) ,isnull(sum(amount_Cr),0)  from gen_jn as a , So_Rpt_Proitems as b   where a.ac_code=b.AcCode And b.CurrType = 1 And a.Date_Work " & MuBTW & " group by  a.ac_code ,  b.RptID  "
        SqlClient = vbCrLf & SqlClient & " Insert into So_Rpt_ProDetail (RptID , AcCode, OpenDr, OpenCr) Select  b.RptID , ac_code , isnull(sum(amount_Dr),0) ,isnull(sum(amount_Cr),0)  from gen_jn as a , So_Rpt_Proitems as b   where a.ac_code=b.AcCode And b.CurrType = 1 And a.Date_Work " & MuLFTD & " group by  a.ac_code ,  b.RptID  "
        SqlClient = vbCrLf & SqlClient & " Insert into So_Rpt_ProDetail (RptID , AcCode, OpenDr, OpenCr) Select  b.RptID , ac_code , isnull(sum(amount_Dr),0) ,isnull(sum(amount_Cr),0)  from Open_jn as a , So_Rpt_Proitems as b   where a.ac_code=b.AcCode And b.CurrType = 1 And a.Date_Work " & MuLFTD & " group by  a.ac_code ,  b.RptID  "
        SqlClient = vbCrLf & SqlClient & " insert into So_Rpt_ProDetail (RptID , AcCode,OpenDr,OpenCr,AmtDr,AmtCr, LCk) Select RptID,AcCode, isnull(Sum(OpenDr),0), isnull(Sum(OpenCr),0) , isnull(Sum(AmtDr),0) , isnull(Sum(AmtCr),0),1  from So_Rpt_ProDetail Group by RptID,AcCode"

        SqlClient = vbCrLf & SqlClient & " delete So_Rpt_ProDetail where LCk=0  "
        SqlClient = vbCrLf & SqlClient & " Update So_Rpt_ProDetail set OpenDr=OpenDr-OpenCr , OpenCr=0 where  OpenDr>OpenCr"
        SqlClient = vbCrLf & SqlClient & " Update So_Rpt_ProDetail set OpenCr=OpenCr-OpenDr , OpenDr=0 where  OpenCr>OpenDr"
        SqlClient = vbCrLf & SqlClient & " Update So_Rpt_ProDetail set AmtDr=AmtDr-AmtCr , AmtCr =0 where  AmtDr>AmtCr"
        SqlClient = vbCrLf & SqlClient & " Update So_Rpt_ProDetail set AmtCr=AmtCr-AmtDr , AmtDr =0 where  AmtCr>AmtDr"
        SqlClient = vbCrLf & SqlClient & " Update So_Rpt_ProDetail set OpenDr=0, OpenCr=0 where  OpenDr=OpenCr"
        SqlClient = vbCrLf & SqlClient & " Update So_Rpt_ProDetail set AmtDr=0, AmtCr =0 where  AmtDr=AmtCr"
        SqlClient = vbCrLf & SqlClient & " Update  So_Rpt_ProDetail set OpenDr = 0 , OpenCr=0   from So_Rpt_ProDetail as a , So_Rpt_Proitems as b where a.RptID=b.RptID and a.acCode=b.AcCode and b.SelOpen=0"
        SqlClient = vbCrLf & SqlClient & " Update  So_Rpt_ProDetail set AmtDr = 0 , AmtCr=0   from So_Rpt_ProDetail as a , So_Rpt_Proitems as b where a.RptID=b.RptID and a.acCode=b.AcCode and b.SelAmt=0"
        SqlClient = vbCrLf & SqlClient & " Update So_Rpt_Proitems set OpenAmt= b.OpenDr-b.OpenCr, Amt =   b.amtDr-b.amtCr  from  So_Rpt_Proitems as a ,So_Rpt_ProDetail as b   " & _
       "  where a.RptID=b.RptID and a.AcCode=b.AcCode and a.RptStatus='Dr-Cr' and RptType = '" & MUTY & "' "

        SqlClient = vbCrLf & SqlClient & "    Update So_Rpt_Proitems set OpenAmt= b.OpenCr-b.OpenDr, Amt =   b.amtCr-b.amtDr  from  So_Rpt_Proitems as a ,So_Rpt_ProDetail as b   " & _
        " where a.RptID=b.RptID and a.AcCode=b.AcCode and a.RptStatus='Cr-Dr' and RptType ='" & MUTY & "' "

        SqlClient = vbCrLf & SqlClient & "  Update So_Rpt_Proitems set OpenAmt= b.OpenDr, Amt =   b.amtDr  from  So_Rpt_Proitems as a ,So_Rpt_ProDetail as b   " & _
         " where a.RptID=b.RptID and a.AcCode=b.AcCode and a.RptStatus='Dr' and RptType = '" & MUTY & "' "

        SqlClient = vbCrLf & SqlClient & "     Update So_Rpt_Proitems set OpenAmt= b.OpenCr, Amt =   b.amtCr  from  So_Rpt_Proitems as a ,So_Rpt_ProDetail as b   " & _
        " where a.RptID=b.RptID and a.AcCode=b.AcCode and a.RptStatus='Cr' and RptType = '" & MUTY & "' "


        SqlClient = vbCrLf & SqlClient & "   insert into So_Rpt_ProUpdate (RptID, OpenAmt, Amt)  select RptID,isnull(sum(OpenAmt),0),isnull(sum(Amt),0) from  So_Rpt_Proitems where    RptType = '" & MUTY & "' group by RptID "
        SqlClient = vbCrLf & SqlClient & "  Update So_Rpt_Pro set   OpenAmt= b.OpenAmt ,  Amt= b.Amt   from So_Rpt_Pro AS a,So_Rpt_ProUpdate As b   where  a.RptType = '" & MUTY & "'   And a.RptID=b.RptID "
        'SqlClient = vbCrLf & SqlClient & " Update So_Rpt_Pro set   " & _
        '" OpenAmt= (Select isnull(SUM(OpenAmt),0) from So_Rpt_Proitems where RptID=b.RptID ) ,   " & _
        '" Amt= (Select isnull(SUM(Amt),0) from So_Rpt_Proitems where RptID=b.RptID)   " & _
        '" from  So_Rpt_Pro as a , So_Rpt_Proitems as b where a.RptID=b.RptID "
        SqlClient = vbCrLf & SqlClient & " Update So_Rpt_Pro set Amt= Amt +OpenAmt "
        CnnEdit()
    End Sub

    Private Sub Update_Sum()
        SqlClient = ""
        SqlClient = vbCrLf & SqlClient & " Update Caculate_Rpt set  CLT_Amt  = CLT_Str ,  CLT_Last_Amt  = CLT_Str  where ( CLT_Str = '+' or CLT_Str = '-' or CLT_Str = '*' or CLT_Str = '+' or CLT_Str = '/' or CLT_Str = '(' or CLT_Str=')' Or CLT_Str<>'Cast(('   Or CLT_Str<>')As Float) '  )    And ( Rpt_Type = '" & MUTY & "'  ) "
        SqlClient = vbCrLf & SqlClient & " delete Caculate_Lock"
        SqlClient = vbCrLf & SqlClient & " delete Caculate_Start  "
        SqlClient = vbCrLf & SqlClient & "  Insert Into Caculate_Start (Rpt_Id , Rpt_Type , clt_Str , clt_Amt , Clt_Last_Amt ) select Rpt_Id , Rpt_Type , clt_Str , clt_Amt , Clt_Last_Amt from Caculate_Rpt where Rpt_Type = '" & MUTY & "'  Order by  Rpt_id ,cnt asc "
        SqlClient = vbCrLf & SqlClient & " update Caculate_Start set lck =0  "
        SqlClient = vbCrLf & SqlClient & " Insert into Caculate_Lock (cnt_Mt)  SELECT  (SELECT     TOP 1 cnt FROM Caculate_Start AS B WHERE(Rpt_Id = A.Rpt_Id   ) ORDER BY cnt desc) AS cnt FROM Caculate_Start  AS A  GROUP BY Rpt_Id ORDER BY Rpt_Id "
        SqlClient = vbCrLf & SqlClient & " update  Caculate_Start set lck=1 from Caculate_Start ,Caculate_Lock  where Caculate_Start.cnt=Caculate_Lock.cnt_MT  "
        SqlClient = vbCrLf & SqlClient & " Update Caculate_Start set Caculate_Start.Amt = So_Rpt_Pro.Amt , Caculate_Start.Last_Amt = So_Rpt_Pro.OpenAmt   from Caculate_Start , So_Rpt_Pro  where  Caculate_Start.CLT_Str  = So_Rpt_Pro.RptId and RptType = '" & MUTY & "'    "
        SqlClient = vbCrLf & SqlClient & " Update Caculate_Start set lck_Amt=0 "
        SqlClient = vbCrLf & SqlClient & " Update Caculate_Start set lck_Amt=1 where CLT_Str <> '+' And CLT_Str <> '-' And CLT_Str <> '*' And CLT_Str <> '+' And CLT_Str <> '/' And CLT_Str <> '(' And CLT_Str<>')' And CLT_Str<>'Cast(('   And CLT_Str<>')As Float)' "
        CnnEdit()


        'Dim ds As New DataSet
        'Try
        '    ConnectCL()
        '    SqlClient = "select *  from Caculate_Start where Rpt_Type = 'Pro1'  Order by  Rpt_id ,cnt  "
        '    LoadCN()
        '    da.Fill(ds, "gen_jn")
        '    For i = 0 To ds.Tables(0).Rows.Count - 1
        '        MsgBox(ds.Tables("Caculate_Start").Rows(i).Item("cnt"))
        '        Dim ss As String = ds.Tables("Caculate_Start").Rows(i).Item("cnt").ToString
        '        If (ds.Tables("Caculate_Start").Rows(i).Item("lck_Amt").ToString) = "1" Then
        '            CLT_Str = CLT_Str & ds.Tables("Caculate_Start").Rows(i).Item("cnt").ToString
        '            CLT_Last_Str = CLT_Last_Str & ds.Tables("Caculate_Start").Rows(i).Item("Last_Amt").ToString
        '        Else
        '            CLT_Str = CLT_Str & ds.Tables("Caculate_Start").Rows(i).Item("CLT_Amt").ToString
        '            CLT_Last_Str = CLT_Last_Str & ds.Tables("Caculate_Start").Rows(i).Item("CLT_Last_Amt").ToString
        '        End If
        '        If ds.Tables("Caculate_Start").Rows(i).Item("lck").ToString = "1" Then
        '            MsgBox(CLT_Last_Str)
        '            'checkhang()
        '        End If
        '    Next i
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message)
        'End Try





        Dim RSC1 As New ADODB.Recordset
        CLT_Str = ""
        CLT_Last_Str = ""
        With RSC1
            Call LoadSqlData("select *  from Caculate_Start where  Rpt_Type = '" & MUTY & "'  Order by  Rpt_id ,cnt asc", RSC1)
            If .RecordCount > 0 Then
                While Not .EOF()
                    'If (RSC1.Fields("Rpt_ID").Value.ToString) = "03" Then
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
                            'If (RSC1.Fields("Rpt_ID").Value.ToString) = "03" Then
                            Dim ss11 As String = " Update  So_Rpt_Pro set Amt = " & CLT_Str & " , OpenAmt = " & CLT_Last_Str & " where  RptID =   '" & (RSC1.Fields("Rpt_ID").Value.ToString) & "' "
                            CNN.Execute(ss11)
                            'End If
                        Else
                            Dim s As String = " Update  So_Rpt_Pro set Amt = " & CLT_Str & " , OpenAmt = " & CLT_Last_Str & " where  RptID =   '" & (RSC1.Fields("Rpt_ID").Value.ToString) & "' "
                            's = 0
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

    Private Sub FmRptPro_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        HeaDer()
        LoadDateNow()
        Me.Text = "FmRptPro(" & MUTY & ")"
        SetControlText(Me)
        RMto.Checked = True
        loadOffice_User()
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
        Off_Usr.SelectedIndex = 0
    End Sub
    Public Sub LoadDateNow()
        Dim MWorkSetting As Date = Date.Now
        Ds.Value = MWorkSetting
        Dt.Value = MWorkSetting
        MdStartDate = Ds.Value
        MdToDate = Dt.Value
        If CDbl(Month(MWorkSetting)) = 1 Then
            Period.SelectedIndex = 0
            Pyy.Value = MWorkSetting
        ElseIf CDbl(Month(MWorkSetting)) = 2 Then
            Period.SelectedIndex = 1
            Pyy.Value = MWorkSetting
        ElseIf CDbl(Month(MWorkSetting)) = 3 Then
            Period.SelectedIndex = 2
            Pyy.Value = MWorkSetting
        ElseIf CDbl(Month(MWorkSetting)) = 4 Then
            Period.SelectedIndex = 3
            Pyy.Value = MWorkSetting
        End If
        M1.SelectedIndex = CDbl(Month(MWorkSetting)) - 1
        M2.SelectedIndex = CDbl(Month(MWorkSetting)) - 1
        Mtoyy.Value = MWorkSetting
        If CDbl(Month(MWorkSetting)) < 7 Then
            Ct.SelectedIndex = 0
            yyt.Value = MWorkSetting
        Else
            Ct.SelectedIndex = 1
            yyt.Value = MWorkSetting
        End If
        yy.Value = MWorkSetting
        Toyy.Value = MWorkSetting
    End Sub
    Private Sub checkhang()
        On Error GoTo hang
hang:
        If Err.Number = 0 Then

        Else

            Exit Sub
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        FmRptProItem.ShowDialog()

    End Sub
    Private Sub PEO_BLS06()
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


    Private Sub UpdateOut_BLS()
        CNN.Execute("delete Ap_Rpt_BLS_Stock ")
        CNN.Execute(" insert into Ap_Rpt_BLS_Stock ( Rpt_ID , Last_Amt_Dr , Last_Amt_Cr, Amt_Dr , Amt_Cr)" & _
                     "  select Rpt_ID , sum(Last_Amt_Dr) As Last_Amt_Dr , sum(Last_Amt_Cr) As Last_Amt_Cr ,sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr   from Ap_Rpt_BLS_Item_Old  where  Rpt_Type = 'Out' group by Rpt_ID")
        CNN.Execute("Update Ap_Rpt_BLS_Old set Amt = Ap_Rpt_BLS_Stock.Amt_Dr-Ap_Rpt_BLS_Stock.Amt_Dr ,Last_Amt =Ap_Rpt_BLS_Stock.Last_Amt_dr-Ap_Rpt_BLS_Stock.Last_Amt_Cr  from Ap_Rpt_BLS_Old ,Ap_Rpt_BLS_Stock where  Ap_Rpt_BLS_Old.Rpt_ID=Ap_Rpt_BLS_Stock.Rpt_ID")
        CNN.Execute("Update Ap_Rpt_BLS_Old set Amt = Ap_Rpt_BLS_Stock.Amt_Cr-Ap_Rpt_BLS_Stock.Amt_Dr ,Last_Amt =Ap_Rpt_BLS_Stock.Last_Amt_Cr-Ap_Rpt_BLS_Stock.Last_Amt_Dr  from Ap_Rpt_BLS_Old ,Ap_Rpt_BLS_Stock where  Ap_Rpt_BLS_Old.Rpt_ID=Ap_Rpt_BLS_Stock.Rpt_ID")
        Dim RSC As New ADODB.Recordset

    End Sub


    Private Sub UpdateIIn_BLS()
        CNN.Execute("delete Ap_Rpt_BLS_Stock ")
        CNN.Execute(" insert into Ap_Rpt_BLS_Stock ( Rpt_ID , Last_Amt_Dr , Last_Amt_Cr, Amt_Dr , Amt_Cr)" & _
                     "  select Rpt_ID , sum(Last_Amt_Dr) As Last_Amt_Dr , sum(Last_Amt_Cr) As Last_Amt_Cr ,sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr   from Ap_Rpt_BLS_Item_Old  where  Rpt_Type = 'In' group by Rpt_ID")
        CNN.Execute("Update Ap_Rpt_BLS_Old set Amt = Ap_Rpt_BLS_Stock.Amt_Dr-Ap_Rpt_BLS_Stock.Amt_cr ,Last_Amt =Ap_Rpt_BLS_Stock.Last_Amt_dr-Ap_Rpt_BLS_Stock.Last_Amt_Cr  from Ap_Rpt_BLS_Old ,Ap_Rpt_BLS_Stock where  Ap_Rpt_BLS_Old.Rpt_ID=Ap_Rpt_BLS_Stock.Rpt_ID")
    End Sub
   


    Private Sub SelectOut_BLS()
        CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
          "where Ap_Rpt_BLS_Item_Old.ac_code = Ap_balance_6_col.ac_code And  Rpt_Type = 'Out'")

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
    
    Private Sub SelcectIn_BLS()

        CNN.Execute("Update Ap_Rpt_BLS_Item_Old set Last_amt_dr= Ap_balance_6_col.open_amt_dr , Last_amt_cr= Ap_balance_6_col.open_amt_cr, amt_dr= Ap_balance_6_col.Rem_Dr , amt_cr= Ap_balance_6_col.Rem_cr from Ap_Rpt_BLS_Item_Old , Ap_balance_6_col " & _
                "where Ap_Rpt_BLS_Item_Old.ac_code = Ap_balance_6_col.ac_code And  Rpt_Type = 'In'")

        CNN.Execute("Insert into Ap_Rpt_BLS_Detail (  Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type )" & _
         " select   Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type from Ap_Rpt_BLS_Item_Old where Rpt_Type = 'In' And ( Amt_Dr <>0 or Amt_Cr <>0  or Last_Amt_Dr <>0 or Last_Amt_Cr <>0 )")

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
    Private Sub BLNEW()
        New_Code = "3901000"
        Code_Dr = "4"
        Code_Cr = "5"
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
        'If CMB_Curr.SelectedIndex = 0 Then
        '    B_Curr = ""
        'Else
        '    B_Curr = " AND  Curr=N'" & CMB_Curr.Text & "' "
        'End If


        CNN.Execute("DELETE  Ap_balance_6_col ")
        CNN.Execute("DELETE FROM Ap_balance_6 ")
        'If CMB_Curr.SelectedIndex = 0 Then
            CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
  " select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr  from gen_jn  WHERE  1=1 " & B_Curr & " and   gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")
            Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
            CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            " select ac_code , sum(amt_dr)as amt_dr , sum(amt_cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE  1=1 " & B_Curr & " and gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
            CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
            " select ac_code  , sum(amt_dr) as amt_dr , sum(amt_cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE 1=1 " & B_Curr & " and     date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        '      Else
        '          CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        '" select ac_code  ,  0 As open_amt_dr, 0 As open_amt_cr , sum(Amount_Dr)as amt_dr , sum(Amount_Cr)as amt_cr  from gen_jn  WHERE  1=1 " & B_Curr & " and  gen_jn.date_work   BETWEEN '" & Format(MdStartDate, "yyyy-MM-dd") & "' AND '" & Format(MdToDate, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "  group BY ac_code ")
        '          Dim S As Date = MdStartDate : S = DateAdd("d", CDbl(-1), MdStartDate)
        '          CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        '          " select ac_code , sum(Amount_Dr)as amt_dr , sum(Amount_Cr)as amt_cr , 0 As amt_dr2 , 0 amt_cr2  from gen_jn  WHERE 1=1 " & B_Curr & " and  gen_jn.date_work   BETWEEN '" & "1-1-" & Format(MdStartDate, "yyyy") & "' AND '" & Format(S, "yyyy-MM-dd") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")
        '          CNN.Execute("INSERT INTO Ap_balance_6 ( ac_code  , open_amt_dr, open_amt_cr , amt_dr , amt_cr   ) " & _
        '          " select ac_code  , sum(Amount_Dr) as amt_dr , sum(Amount_Cr) as amt_cr  , 0 As amt_dr2 , 0 amt_cr2 from Open_jn WHERE  1=1 " & B_Curr & " and    date_work='" & "1-1-" & Format(MdStartDate, "yyyy") & "' " & Ac_Code & " " & MULook2 & "   group BY ac_code")

        '      End If

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
 
        If MDACC00 = 0 Then
            New_Code = New_Code
        Else

            New_Code = New_Code4
        End If

        Insr = "delete  Ap_balance_6  " & _
           "insert Into Ap_balance_6 (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  )   select '" & New_Code & "' , Sum(open_amt_dr), Sum(open_amt_cr), Sum(amt_Dr), Sum(amt_Cr) from Ap_balance_6_col where  left(Ac_Code,1) ='" & Code_Dr & "' or left(Ac_Code,4) ='" & Code_Dr1 & "'  Or left(Ac_Code,1)='" & Code_Cr & "' or left(Ac_Code,4)='" & Code_Cr1 & "'  or  Ac_Code =  '" & New_Code & "'  or  Ac_Code =  '" & New_Code4 & "'  " & _
  "update  Ap_balance_6 set open_amt_dr = 0 where open_amt_dr  is null  " & _
  "update  Ap_balance_6 set open_amt_cr = 0 where open_amt_cr  is null  " & _
  "update  Ap_balance_6 set amt_dr = 0 where amt_dr  is null  " & _
  "update  Ap_balance_6 set amt_cr = 0 where amt_cr  is null   " & _
   "Update  Ap_balance_6 set   open_amt_dr = open_amt_dr  - open_amt_cr , open_amt_cr=0  where open_amt_dr  >= open_amt_cr   " & _
  "Update  Ap_balance_6 set   open_amt_cr = open_amt_cr  - open_amt_dr , open_amt_dr=0  where open_amt_cr  >= open_amt_dr " & _
  "Update  Ap_balance_6 set   amt_dr = amt_dr  - amt_cr , amt_cr=0  where amt_dr  >= amt_cr   " & _
  "Update  Ap_balance_6 set   amt_cr = amt_cr  - amt_dr , amt_dr=0  where amt_cr  >= amt_dr " & _
     "delete  Ap_balance_6_col  where   left(Ac_Code,1) ='" & Code_Dr & "' or left(Ac_Code,4) ='" & Code_Dr1 & "'  Or  left(Ac_Code,1)='" & Code_Cr & "' or left(Ac_Code,4)='" & Code_Cr1 & "'  or  Ac_Code =  '" & New_Code & "'  or  Ac_Code =  '" & New_Code4 & "'  " & _
       "  insert Into Ap_balance_6_col (Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  ) select Ac_Code , open_amt_dr , open_amt_cr , amt_dr , amt_cr  from Ap_balance_6"
        CNN.Execute(Insr)

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
    End Sub
    Private Sub PEO_InCom06()

        Dim s As String = "update  Ap_Rpt_Income_Item_Old set Last_Amt_dr  =  0 , Last_Amt_Cr  =  0 , amt_dr  = 0 , amt_cr  = 0 "
        CNN.Execute(s)

        CNN.Execute("update Ap_Rpt_Income_Old set  Last_Amt  = 0 , Amt  = 0    ")
        CNN.Execute("DELETE FROM Ap_Rpt_Incon_Detail ")
        Off_Find = Off_Usr.Text : MuTable = "" : Call Find_Company()
        ChangBalance_InC()
        SelcectInLast()
        UpdateIInLast_InC()
        SelectOutLast_InC()
        UpdateOut_InC()
        Update_Sum_InC()
        '========1 =========
        'CNN.Execute("update Ap_Rpt_Income set amt=(select isnull(sum(amt),0) as amt  from Ap_Rpt_Income where (rpt_ID='1.1.1' or  rpt_ID='1.1.1.1'  or  rpt_ID='1.1.1.2'  or  rpt_ID='1.1.1.2.1'  or  rpt_ID='1.1.1.2.2' )) where  rpt_ID='1' ")
        CNN.Execute("update Ap_Rpt_Income_Old set amt=(select isnull(sum(amt),0) as amt  from Ap_Rpt_Income_Old where (rpt_ID='1.1.1.1.1' or  rpt_ID='1.1.1.1.2'  or  rpt_ID='1.1.1.1.3'  or  rpt_ID='1.1.1.1.4'  or  rpt_ID='1.1.1.1.5'   or  rpt_ID='1.1.1.1.6' )) where  rpt_ID='1.1.1.1' ")
        CNN.Execute("update Ap_Rpt_Income_Old set Last_Amt=(select isnull(sum(Last_Amt),0) as amt  from Ap_Rpt_Income_Old where (rpt_ID='1.1.1.1.1' or  rpt_ID='1.1.1.1.2'  or  rpt_ID='1.1.1.1.3'  or  rpt_ID='1.1.1.1.4'  or  rpt_ID='1.1.1.1.5'   or  rpt_ID='1.1.1.1.6' )) where  rpt_ID='1.1.1.1' ")
        '========1 =========
        CNN.Execute("update Ap_Rpt_Income_Old set amt=(select isnull(sum(amt),0) as amt  from Ap_Rpt_Income_Old where (rpt_ID='1.1.1.1.1' or  rpt_ID='1.1.1.1.2'  or  rpt_ID='1.1.1.1.3'  or  rpt_ID='1.1.1.1.4'  or  rpt_ID='1.1.1.1.5'   or  rpt_ID='1.1.1.1.6' )) where  rpt_ID='1.1.1' ")
        CNN.Execute("update Ap_Rpt_Income_Old set Last_Amt=(select isnull(sum(Last_Amt),0) as amt  from Ap_Rpt_Income_Old where (rpt_ID='1.1.1.1.1' or  rpt_ID='1.1.1.1.2'  or  rpt_ID='1.1.1.1.3'  or  rpt_ID='1.1.1.1.4'  or  rpt_ID='1.1.1.1.5'   or  rpt_ID='1.1.1.1.6' )) where  rpt_ID='1.1.1' ")
        '==========1.1.2
        CNN.Execute("update Ap_Rpt_Income_Old set amt=(select isnull(sum(amt),0) as amt  from Ap_Rpt_Income_Old where (rpt_ID='1.1.2.1' )) where  rpt_ID='1.1.2' ")
        CNN.Execute("update Ap_Rpt_Income_Old set Last_Amt=(select isnull(sum(Last_Amt),0) as amt  from Ap_Rpt_Income_Old where  (rpt_ID='1.1.2.1'))  where  rpt_ID='1.1.2' ")
        '==========1.1.3
        CNN.Execute("update Ap_Rpt_Income_Old set amt=(select isnull(sum(amt),0) as amt  from Ap_Rpt_Income_Old where (rpt_ID='1.1.3.5' )) where  rpt_ID='1.1.3' ")
        CNN.Execute("update Ap_Rpt_Income_Old set Last_Amt=(select isnull(sum(Last_Amt),0) as amt  from Ap_Rpt_Income_Old where  (rpt_ID='1.1.3.5'))  where  rpt_ID='1.1.3' ")

        '=dsfsdfdfd
        CNN.Execute("update Ap_Rpt_Income_Old set amt=(select isnull(sum(amt),0) as amt  from Ap_Rpt_Income_Old where (rpt_ID='1.1.1' or  rpt_ID='1.1.2'  or  rpt_ID='1.1.3')) where  rpt_ID='1.13' ")
        CNN.Execute("update Ap_Rpt_Income_Old set Last_Amt=(select isnull(sum(Last_Amt),0) as amt  from Ap_Rpt_Income_Old where  (rpt_ID='1.1.1' or  rpt_ID='1.1.2'  or  rpt_ID='1.1.3'))  where  rpt_ID='1.13' ")

        '==========1.2.1
        CNN.Execute("update Ap_Rpt_Income_Old set amt=(select isnull(sum(amt),0) as amt  from Ap_Rpt_Income_Old where (rpt_ID='1.2.1.1.1' or  rpt_ID='1.2.1.1.2')) where  rpt_ID='1.2.1' ")
        CNN.Execute("update Ap_Rpt_Income_Old set Last_Amt=(select isnull(sum(Last_Amt),0) as amt  from Ap_Rpt_Income_Old where  (rpt_ID='1.2.1.1.1' or  rpt_ID='1.2.1.1.2'))  where  rpt_ID='1.2.1' ")
        '==========1.2.2
        CNN.Execute("update Ap_Rpt_Income_Old set amt=(select isnull(sum(amt),0) as amt  from Ap_Rpt_Income_Old where (rpt_ID='1.2.2.2.1' or  rpt_ID='1.2.2.2.2'or  rpt_ID='1.2.2.2.3'or  rpt_ID='1.2.2.2.4')) where  rpt_ID='1.2.2' ")
        CNN.Execute("update Ap_Rpt_Income_Old set Last_Amt=(select isnull(sum(Last_Amt),0) as amt  from Ap_Rpt_Income_Old where   (rpt_ID='1.2.2.2.1' or  rpt_ID='1.2.2.2.2'or  rpt_ID='1.2.2.2.3'or  rpt_ID='1.2.2.2.4'))  where  rpt_ID='1.2.2' ")
        '==========1.2.3
        CNN.Execute("update Ap_Rpt_Income_Old set amt=(select isnull(sum(amt),0) as amt  from Ap_Rpt_Income_Old where (rpt_ID='1.2.3.3.1')) where  rpt_ID='1.2.3' ")
        CNN.Execute("update Ap_Rpt_Income_Old set Last_Amt=(select isnull(sum(Last_Amt),0) as amt  from Ap_Rpt_Income_Old where   (rpt_ID='1.2.3.3.1'))  where  rpt_ID='1.2.3' ")

        '=dsfsdfdfd
        CNN.Execute("update Ap_Rpt_Income_Old set amt=(select isnull(sum(amt),0) as amt  from Ap_Rpt_Income_Old where (rpt_ID='1.2.1' or  rpt_ID='1.2.2'  or  rpt_ID='1.2.3' or  rpt_ID='1.2.4')) where  rpt_ID='1.2.4' ")
        CNN.Execute("update Ap_Rpt_Income_Old set Last_Amt=(select isnull(sum(Last_Amt),0) as amt  from Ap_Rpt_Income_Old where  (rpt_ID='1.2.1' or  rpt_ID='1.2.2'  or  rpt_ID='1.2.3' or  rpt_ID='1.2.4'))  where  rpt_ID='1.2.4' ")

        '=ກຳໄລກ່ອນອາກອນ
        Dim Prof As String = "delete AP_Sum "
        Prof = Prof & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,amt,0 from Ap_Rpt_Income_Old     where rpt_ID='1.13'  "
        Prof = Prof & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '1.2.4.1'  ,0 ,amt from Ap_Rpt_Income_Old     where rpt_ID='1.2.4'  "
        Prof = Prof & "  update AP_Sum set amt=(select sum(amt2)-sum(amt3) from AP_Sum where (rpt_ID='1.13' or rpt_ID='1.2.4.1')  ) where rpt_ID='1.2.4.1'   "
        Prof = Prof & "  delete  AP_Sum  where rpt_ID='1.13'   "
        Prof = Prof & " Update Ap_Rpt_Income_Old set Ap_Rpt_Income_Old.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Income_Old    where Ap_Rpt_Income_Old.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(Prof)

        Dim Prof_Last As String = "delete AP_Sum "
        Prof_Last = Prof_Last & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,Last_amt,0 from Ap_Rpt_Income_Old     where rpt_ID='1.13'  "
        Prof_Last = Prof_Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '1.2.4.1'  ,0 ,Last_amt from Ap_Rpt_Income_Old     where rpt_ID='1.2.4'  "
        Prof_Last = Prof_Last & "  update AP_Sum set amt=(select sum(amt2)-sum(amt3) from AP_Sum where (rpt_ID='1.13' or rpt_ID='1.2.4.1')  ) where rpt_ID='1.2.4.1'   "
        Prof_Last = Prof_Last & "  delete  AP_Sum  where rpt_ID='1.13'   "
        Prof_Last = Prof_Last & " Update Ap_Rpt_Income_Old set Ap_Rpt_Income_Old.Last_amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Income_Old    where Ap_Rpt_Income_Old.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(Prof_Last)

        '=Profit 
        Dim MI As String = "delete AP_Sum "
        MI = MI & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,amt,0 from Ap_Rpt_Income_Old     where rpt_ID='1.13'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '1.2.4.3'  ,0 ,amt from Ap_Rpt_Income_Old     where rpt_ID='1.2.4'  "
        MI = MI & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '1.2.4.3'  ,0 ,amt from Ap_Rpt_Income_Old     where rpt_ID='1.2.4.2'  "
        MI = MI & "  update AP_Sum set amt=(select sum(amt2)-sum(amt3) from AP_Sum where (rpt_ID='1.13' or rpt_ID='1.2.4.3')  ) where rpt_ID='1.2.4.3'   "
        MI = MI & "  delete  AP_Sum  where rpt_ID='1.13'   "
        MI = MI & " Update Ap_Rpt_Income_Old set Ap_Rpt_Income_Old.amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Income_Old    where Ap_Rpt_Income_Old.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(MI)

        Dim Last As String = "delete AP_Sum "
        Last = Last & "  insert into AP_Sum(Rpt_ID,amt2,amt3) select rpt_ID ,Last_amt,0 from Ap_Rpt_Income_Old     where rpt_ID='1.13'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '1.2.4.3'  ,0 ,Last_amt from Ap_Rpt_Income_Old     where rpt_ID='1.2.4'  "
        Last = Last & " insert into AP_Sum(Rpt_ID,amt2,amt3) select  '1.2.4.3'  ,0 ,Last_amt from Ap_Rpt_Income_Old     where rpt_ID='1.2.4.2'  "
        Last = Last & "  update AP_Sum set amt=(select sum(amt2)-sum(amt3) from AP_Sum where (rpt_ID='1.13' or rpt_ID='1.2.4.3')  ) where rpt_ID='1.2.4.3'   "
        Last = Last & "  delete  AP_Sum  where rpt_ID='1.13'   "
        Last = Last & " Update Ap_Rpt_Income_Old set Ap_Rpt_Income_Old.Last_amt=AP_Sum.amt from AP_Sum,Ap_Rpt_Income_Old    where Ap_Rpt_Income_Old.rpt_ID=AP_Sum.rpt_ID "
        CNN.Execute(Last)

        CNN.Execute("Update Ap_Rpt_Income_Old set Amt=Amt+Last_Amt")
    End Sub

    Private Sub UpdateOut_InC()
        Dim RSC As New ADODB.Recordset
        LoadSqlData("select Rpt_ID, sum(Last_Amt_dr) As Last_Amt_Dr , sum(Last_Amt_cr) As Last_Amt_cr , sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr    from Ap_Rpt_Income_Item_Old  where  Rpt_Type = 'Out'   group by Rpt_ID ", RSC)
        With RSC
            Do Until .EOF = True
                CNN.Execute("Update Ap_Rpt_Income_Old set " & _
                    " Amt ='" & CDbl(CDbl((.Fields("Amt_dr").Value)) - CDbl((.Fields("Amt_cr").Value))) & "' " & _
                      " ,Last_Amt ='" & CDbl(CDbl((.Fields("Last_Amt_dr").Value)) - CDbl((.Fields("Last_Amt_cr").Value))) & "' " & _
                       " where Rpt_ID = '" & (.Fields("Rpt_ID").Value) & "' ")
                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub SelectOutLast_InC()

        LoadSqlData("select * from Ap_Rpt_Income_Item_Old where  Rpt_Type = 'Out'  ", RSCIn_M)
        With RSCIn_M
            Do Until .EOF = True
                Call UpdateOut_Item_InC()
                .MoveNext()
            Loop
        End With
        'If RSCIn_M.State = ConnectionState.Open Then RSCIn_M.Close()
    End Sub

    Private Sub UpdateOut_Item_InC()
        Dim RSCkk As New ADODB.Recordset
        LoadSqlData(" select * from Ap_balance_6_col   where ac_code =  '" & (RSCIn_M.Fields("Ac_Code").Value) & "' ", RSCkk)
        With RSCkk
            Do Until .EOF = True
                'MsgBox((RSCIn_M.Fields("Ac_Code").Value))
                CNN.Execute("Insert into Ap_Rpt_Incon_Detail (Ac_Code_Parent , Rpt_Id , Ac_Code , Ac_Name  ,  Last_Amt_Dr , Last_Amt_Cr, Amt_Dr , Amt_Cr , Rpt_Type ) values (  '" & CStr((RSCIn_M.Fields("Ac_Code").Value)) & "' , '" & CStr((RSCIn_M.Fields("Rpt_Id").Value)) & "' , '" & CStr((.Fields("Ac_Code").Value)) & "' , N'" & CStr((.Fields("Ac_Name").Value)) & "'  ,   " & CDbl((.Fields("Open_Amt_dr").Value)) & " , " & CDbl((.Fields("Open_Amt_cr").Value)) & " , " & CDbl((.Fields("Amt_dr").Value)) & " , " & CDbl((.Fields("Amt_cr").Value)) & " , 'Out' )")
                CNN.Execute("update  Ap_Rpt_Income_Item_Old set Last_Amt_Dr  =  Last_Amt_Dr+" & CDbl((.Fields("Open_Amt_dr").Value)) & " , Last_Amt_Cr  = Last_Amt_Cr+" & CDbl((.Fields("Open_Amt_cr").Value)) & "  , Amt_Dr  =  Amt_Dr+" & CDbl((.Fields("Amt_dr").Value)) & " , Amt_Cr  = Amt_Cr+" & CDbl((.Fields("Amt_cr").Value)) & "   where ac_code = '" & (RSCIn_M.Fields("ac_code").Value) & "' And  Rpt_Type = 'Out' ")

                .MoveNext()
            Loop
        End With
    End Sub
    Private Sub UpdateIInLast_InC()
        Dim RSC As New ADODB.Recordset
        LoadSqlData("select Rpt_ID , sum(Last_Amt_Dr) As Last_Amt_Dr , sum(Last_Amt_Cr) As Last_Amt_Cr ,sum(Amt_Dr) As Amt_Dr , sum(Amt_Cr) As Amt_Cr   from Ap_Rpt_Income_Item_Old  where  Rpt_Type = 'In'   group by Rpt_ID ", RSC)
        With RSC
            Do Until .EOF = True

                CNN.Execute("Update Ap_Rpt_Income_Old set " & _
                     " Amt = '" & CDbl(CDbl((.Fields("Amt_cr").Value)) - CDbl((.Fields("Amt_dr").Value))) & "' " & _
                       " , Last_Amt ='" & CDbl(CDbl((.Fields("Last_Amt_cr").Value)) - CDbl((.Fields("Last_Amt_dr").Value))) & "' " & _
                        " where Rpt_ID = '" & (.Fields("Rpt_ID").Value) & "' ")
                .MoveNext()
            Loop
        End With
    End Sub


    Private Sub SelcectInLast()
        LoadSqlData("select * from Ap_Rpt_Income_Item_Old where  Rpt_Type = 'In'  ", RSCIn_M)
        With RSCIn_M
            Do Until .EOF = True
                UpdateIIn_ItemLast_InC()
                .MoveNext()
            Loop
        End With
    End Sub
    Private Sub UpdateIIn_ItemLast_InC()
        Dim RSCkk As New ADODB.Recordset
        LoadSqlData(" select * from Ap_balance_6_col   where ac_code =  '" & (RSCIn_M.Fields("Ac_Code").Value) & "' ", RSCkk)
        With RSCkk
            Do Until .EOF = True
                CNN.Execute("Insert into Ap_Rpt_Incon_Detail (  Rpt_Id , Ac_Code , Ac_Name , Last_Amt_Dr , Last_Amt_Cr , Amt_Dr , Amt_Cr , Rpt_Type ) values ( '" & CStr((RSCIn_M.Fields("Rpt_Id").Value)) & "' , '" & CStr((.Fields("Ac_Code").Value)) & "' , N'" & CStr((.Fields("Ac_Name").Value)) & "'   , " & CDbl((.Fields("open_amt_dr").Value)) & " , " & CDbl((.Fields("open_amt_Cr").Value)) & "   , " & CDbl((.Fields("Amt_dr").Value)) & " , " & CDbl((.Fields("Amt_cr").Value)) & " , 'In')")
                CNN.Execute("update  Ap_Rpt_Income_Item_Old set  Last_amt_dr  =  Last_amt_dr+" & CDbl((.Fields("open_amt_dr").Value)) & " , Last_amt_cr  = Last_amt_cr+" & CDbl((.Fields("open_amt_Cr").Value)) & " , Amt_Dr  =  Amt_Dr+" & CDbl((.Fields("Amt_dr").Value)) & " , Amt_Cr  = Amt_Cr+" & CDbl((.Fields("Amt_cr").Value)) & "   where ac_code = '" & (RSCIn_M.Fields("ac_code").Value) & "' And  Rpt_Type = 'In' ")
                .MoveNext()
            Loop
        End With
    End Sub
    Private Sub Update_Sum_InC()
        CNN.Execute("update Ap_Rpt_Incon_Detail set  Rpt_Name=Ap_Rpt_Income_Old.Description from   Ap_Rpt_Incon_Detail , Ap_Rpt_Income_Old where Ap_Rpt_Incon_Detail.Rpt_Id = Ap_Rpt_Income_Old.Rpt_Id")
        CNN.Execute(" Update Caculate_Rpt set  CLT_Amt  = CLT_Str ,  CLT_Last_Amt  = CLT_Str where CLT_Str = '+' or CLT_Str = '-' or CLT_Str = '*' or CLT_Str = '+' or CLT_Str = '/' or CLT_Str = '(' or CLT_Str=')' Or CLT_Str<>'Cast(('   Or CLT_Str<>')As Float)'")
        CNN.Execute("delete Caculate_Lock")
        CNN.Execute("delete Caculate_Start")
        CNN.Execute(" Insert Into Caculate_Start (Rpt_Id , Rpt_Type , clt_Str , clt_Amt , Clt_Last_Amt ) select Rpt_Id , Rpt_Type , clt_Str , clt_Amt , Clt_Last_Amt from Caculate_Rpt where Rpt_Type = 'INC'  Order by  Rpt_id ,cnt asc ")
        CNN.Execute("update Caculate_Start set lck =0")
        CNN.Execute("Insert into Caculate_Lock (cnt_Mt)  SELECT  (SELECT     TOP 1 cnt FROM Caculate_Start AS B WHERE(Rpt_Id = A.Rpt_Id   ) ORDER BY cnt desc) AS cnt FROM Caculate_Start  AS A  GROUP BY Rpt_Id ORDER BY Rpt_Id")
        CNN.Execute("update  Caculate_Start set lck=1 from Caculate_Start ,Caculate_Lock  where Caculate_Start.cnt=Caculate_Lock.cnt_MT")
        CNN.Execute("  Update Caculate_Start set Caculate_Start.Amt = Ap_Rpt_Income_Old.Amt , Caculate_Start.Last_Amt = Ap_Rpt_Income_Old.Last_Amt   from Caculate_Start , Ap_Rpt_Income_Old  where  Caculate_Start.CLT_Str  = Ap_Rpt_Income_Old.Rpt_Id  ")
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
                            Dim s1 As String = " Update  Ap_Rpt_Income_Old set Amt = " & CLT_Str & " , Last_Amt = " & CLT_Last_Str & " where  Rpt_ID =   '" & (RSC1.Fields("Rpt_ID").Value.ToString) & "' "
                            CNN.Execute(s1)
                        Else
                            Dim s As String = " Update  Ap_Rpt_Income_Old set Amt = " & CLT_Str & " , Last_Amt = " & CLT_Last_Str & " where  Rpt_ID =   '" & (RSC1.Fields("Rpt_ID").Value.ToString) & "' "
                            Dim k As String = s
                            'CNN.Execute(k)

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
    Private Sub ChangBalance_InC()
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
        'Call Chang_Incom()
        CNN.Execute("Update  Ap_balance_6_col set Rem_cr=0 , Rem_dr= (open_amt_dr + amt_dr) - (open_amt_cr + amt_cr) where (open_amt_dr + amt_dr) >= (open_amt_cr + amt_cr) ")
        CNN.Execute("Update  Ap_balance_6_col set Rem_dr=0 , Rem_cr= (open_amt_cr + amt_cr) - (open_amt_dr + amt_dr) where (open_amt_cr + amt_cr) >= (open_amt_dr + amt_dr) ")
        CNN.Execute("update Ap_balance_6_col set Ap_balance_6_col.ac_name=Acc_Code.Name_L from Ap_balance_6_col , Acc_Code where Ap_balance_6_col.Ac_Code = Acc_Code.Ac_Code")

    End Sub
End Class