Module MdiContainer

    Dim Str, Str2 As String
    Public CvConsonant, MonthLetter As String
    Public Ac_Code As String
    Public New_Code, Code_Dr, Code_Cr, Insr As String
    Public New_Code4, Code_Dr1, Code_Cr1 As String
    Public Sub LoadInfo()
        Dim W As String = Format(CDate(MWorkSetting), "ddd")
        Dim d As String = Format(CDate(MWorkSetting), "dd/MM/yyyy")
        If W = "Mon" Then Str = "ວັນຈັນ"
        If W = "Tue" Then Str = "ວັນອັງຄານ"
        If W = "Wed" Then Str = "ວັນພຸດ"
        If W = "Thu" Then Str = "ວັນພະຫັດ"
        If W = "Fri" Then Str = "ວັນສຸກ"
        If W = "Sat" Then Str = "ວັນເສົາ"
        If W = "Sun" Then Str = "ວັນອາທິດ"
        If Format(CDate(MWorkSetting), "MM") = "01" Then MonthLetter = "ມັງກອນ"
        If Format(CDate(MWorkSetting), "MM") = "02" Then MonthLetter = "ກຸມພາ"
        If Format(CDate(MWorkSetting), "MM") = "03" Then MonthLetter = "ມີນາ"
        If Format(CDate(MWorkSetting), "MM") = "04" Then MonthLetter = "ເມສາ"
        If Format(CDate(MWorkSetting), "MM") = "05" Then MonthLetter = "ພຶດສະພາ"
        If Format(CDate(MWorkSetting), "MM") = "06" Then MonthLetter = "ມີຖຸນາ"
        If Format(CDate(MWorkSetting), "MM") = "07" Then MonthLetter = "ກໍລະກົດ"
        If Format(CDate(MWorkSetting), "MM") = "08" Then MonthLetter = "ສິງຫາ"
        If Format(CDate(MWorkSetting), "MM") = "09" Then MonthLetter = "ກັນຍາ"
        If Format(CDate(MWorkSetting), "MM") = "10" Then MonthLetter = "ຕຸລາ"
        If Format(CDate(MWorkSetting), "MM") = "11" Then MonthLetter = "ພະຈິກ"
        If Format(CDate(MWorkSetting), "MM") = "12" Then MonthLetter = "ທັນວາ"



        FmMain.TUserId.Text = "ລະຫັດຜູ້ໃຊ້: " & MUserID & "|"
        FmMain.TUserName.Text = "ຊື່ຜູ້ໃຊ້: " & MUserName & "|"
        FmMain.TPermision.Text = "ໃຊ້ໃນນາມ: " & MPermit & "|"
        FmMain.TCompanyName.Text = "ສາຂາ: " & MuSubOff & " *ເມືອງ:" & Mid(FmLogin.Sub_Company.Text, 6, CDbl(Len(FmLogin.Sub_Company.Text)) - 5) & ", ແຂວງ: " & Mid(FmLogin.cmbCompany.Text, 4, CDbl(Len(FmLogin.cmbCompany.Text)) - 3) & "  |"
        FmMain.TDate.Text = Str & ", ວັນທີ່: " & Format(CDate(MWorkSetting), "dd") & " " & MonthLetter & " ປີ: " & Format(CDate(MWorkSetting), "yyyy") & ","
        FmMain.Label17.Text = MPermit



        'FmMain.Label21.Text = Format(Now, "dd/MM/yyyy")
    End Sub
    Public Sub MdiCNum()

        FmMain.Text = "AP Banking System 10"


        Dim s As Integer = FmMain.CmbForm.SelectedIndex
        Dim x As Integer = FmMain.CmbForm.Items.Count

        Dim u As Integer = 0
        For i As Integer = 0 To (My.Application.OpenForms.Count - 1)
            If My.Application.OpenForms.Item(i).Name() <> "FmLogin" And My.Application.OpenForms.Item(i).Name() <> "FmShow" And My.Application.OpenForms.Item(i).Name() <> "FmMain" Then
                u = u + 1
            End If
        Next i


        If x > 1 Then
            FmMain.lblpage_total.Text = FmMain.CmbForm.SelectedIndex + 1 & "/" & u - 1
        End If

        If u = 1 Then
            FmMain.lblpage_total.Text = "0/0"
        End If
        If s > 0 Then
            FmMain.CmbForm.SelectedIndex = FmMain.CmbForm.SelectedIndex - 1
            FmMain.lblpage_total.Text = FmMain.CmbForm.SelectedIndex + 1 & "/" & u - 1
        End If

        If Application.OpenForms.Count = 4 Then
            FmMain.IsMdiContainer = False
            FmMain.Panel4.Visible = True
            FmMain.Text = "AP Banking System 10"
         

            'FmMain.lblpage_total.Text = "0/0"
            'FmMain.CmbForm.Items.Clear()
            'For i As Integer = 0 To (My.Application.OpenForms.Count - 1)
            '    If My.Application.OpenForms.Item(i).Name() <> "FmLogin" And My.Application.OpenForms.Item(i).Name() <> "FmShow" And My.Application.OpenForms.Item(i).Name() <> "FmMain" Then
            '        FmMain.CmbForm.Items.Add(My.Application.OpenForms.Item(i).Name())
            '    End If
            'Next i

            'If FmMain.CmbForm.SelectedIndex >= 0 Then
            '    FmMain.lblpage_total.Text = FmMain.CmbForm.SelectedIndex + 1 & "/" & FmMain.CmbForm.Items.Count
            'End If
            'FmMain.TabControl1.TabPages.Remove(FmMain.TabControl1.SelectedTab)
        End If
        'If Num > 0 Then
        '    FmMain.Panel4.Visible = False

        '    IsMdiContainerTrue()
        'ElseIf Num <= 0 Then
        '    FmMain.Panel4.Visible = True
        '    IsMdiContainerFalse()
        'End If
    End Sub

    Public Sub IsMdiContainerTrue()
        FmMain.IsMdiContainer = True
    End Sub

    Public Sub IsMdiContainerFalse()
        FmMain.IsMdiContainer = False
    End Sub
End Module
