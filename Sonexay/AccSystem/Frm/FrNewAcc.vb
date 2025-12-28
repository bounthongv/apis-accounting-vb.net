Public Class FrNewAcc
    Dim Finance As String
    Dim AcTypeLao, AcTypeEng As String
    Dim d, p As String

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        txtAc_code.Enabled = True
        fmShartOfAcc.StartLoadDataList()
        Close()
    End Sub
    Private Sub AddFinance()
        If CAs.Checked = True Then
            Finance = "As"
            'CNN.Execute("delete Ap_Rpt_BLS_Item_Old where Ac_Code like '" & txtAc_code.Text & "%' And Rpt_ID = '123' And Rpt_Type = 'In'  insert into Ap_Rpt_BLS_Item_Old (Rpt_ID , Ac_Code , Ac_Name, Rpt_Type) select N'123' ,  Ac_Code , Name_L , 'In' from Acc_Code where Ac_Code like '" & txtAc_code.Text & "%'  ")

        ElseIf CFi.Checked = True Then
            Finance = "Fi"
        ElseIf CIn.Checked = True Then
            Finance = "In"
        ElseIf CEx.Checked = True Then
            Finance = "Ex"
        End If
    End Sub



    Private Sub LoadFinance()
        If Finance = "As" Then
            CAs.Checked = True
            CFi.Checked = False
            CIn.Checked = False
            CEx.Checked = False
        ElseIf Finance = "Fi" Then
            CAs.Checked = False
            CFi.Checked = True
            CIn.Checked = False
            CEx.Checked = False
        ElseIf Finance = "In" Then
            CAs.Checked = False
            CFi.Checked = False
            CIn.Checked = True
            CEx.Checked = False
        ElseIf Finance = "Ex" Then
            CAs.Checked = False
            CFi.Checked = False
            CIn.Checked = False
            CEx.Checked = True
        End If
    End Sub
    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        Call AddFinance()
        '1 /**************************************
        If Trim(txtAc_code.Text) = "" Then MsgBox("ກະລຸນາໃສ່ ລະຫັດບັນຊີ ກ່ອນ!", MsgBoxStyle.OkOnly) : txtAc_code.Focus() : Exit Sub

        '2 /****************************************
        'For i = 1 To Len(Trim(txtAc_code.Text))
        '    If Mid(Trim(txtAc_code.Text), i, 1) = "." Then
        '        AcTypeLao = "ບັນຊີຍ່ອຍ (D)"
        '        AcTypeEng = "Detail Account"
        '        Exit For
        '    Else
        '        AcTypeLao = "ບັນຊີແມ່ (P)"
        '        AcTypeEng = "Parent Account"
        '    End If
        'Next i

        If Len(txtAc_code.Text) > "7" Then
            AcTypeLao = "ບັນຊີຍ່ອຍ (D)"
            AcTypeEng = "Detail Account"
        Else
            AcTypeLao = "ບັນຊີແມ່ (P)"
            AcTypeEng = "Parent Account"
        End If



        d = txtAc_code.Text

        For i = 1 To Len(d)
            'If Mid(d, i, 1) = "." Then
            '    p = Microsoft.VisualBasic.Left(d, i - 1)
            '    Exit Sub
            'Else
            '    p = d
            'End If

        Next i
        '3 /**************************************
        If txtAc_code.Enabled = True Then
            Call LoadSqlData("SELECT AC_CODE FROM Acc_Code WHERE AC_CODE = '" & Trim(txtAc_code.Text) & "'", RSC)
            If RSC.RecordCount > 0 Then
                MsgBox("ເລກລະຫັດ : " & Trim(txtAc_code.Text) & " ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
                txtAc_code.Focus()
                If RSC.State = ConnectionState.Open Then RSC.Close()
                Exit Sub
            End If
            '==============
            'If TxtWISE_Orginal.Text = "" Then
            '    TxtWISE_Orginal.Text = Trim(txtAc_code.Text)
            'End If
            'Call LoadSqlData("SELECT AC_CODE FROM Acc_Code WHERE Ac_Original=N'" & Trim(TxtWISE_Orginal.Text) & "'", RSC)
            'If RSC.RecordCount > 0 Then
            '    MsgBox("ເລກລະຫັດ : " & Trim(TxtWISE_Orginal.Text) & " ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
            '    TxtWISE_Orginal.Focus()
            '    If RSC.State = ConnectionState.Open Then RSC.Close()
            '    Exit Sub
            'End If
            '==========

            Call LoadSqlData("SELECT AC_CODE FROM Acc_Code WHERE AC_CODE = '" & Trim(txtAc_code.Text) & "'", RSC)
            If RSC.RecordCount = 0 Then
                CNN.Execute("INSERT INTO Acc_Code( AC_CODE,Ac_Original,  Name_L , Name_E ,Acc_Type ,Acc_TypeE ,Print_status ,last_user ,last_update , Acc_Parent , Finance ) " & _
                    "Values('" & txtAc_code.Text & "',N'" & TxtWISE_Orginal.Text & "', N'" & txtAccName.Text.Trim & "', N'" & txtAccName_E.Text.Trim & "', N'" & AcTypeLao & "', '" & AcTypeEng & "', '" & "0" & "', '" & MUserName & "', '" & Format(DtmDate.Value, "yyyy-MM-dd") & "' , '" & p & "' , '" & Finance & "')")
            Else

            End If
            If RSC.State = ConnectionState.Open Then RSC.Close()
            MsgBox("ການບັນທຶກສຳເລັດ!", MsgBoxStyle.OkOnly)
        End If
        '4 /**************************************
        If txtAc_code.Enabled = False Then
            Call LoadSqlData("SELECT AC_CODE FROM Acc_Code WHERE AC_CODE = '" & Trim(txtAc_code.Text) & "'", RSC)
            CNN.Execute("UPDATE Acc_Code SET Ac_Original= N'" & TxtWISE_Orginal.Text & "',name_l= N'" & txtAccName.Text & "',name_e='" & txtAccName_E.Text & "',last_user='" & MUserName & "',last_update='" & Format(DtmDate.Value, "yyyy-MM-dd") & "' , Acc_Parent='" & p & "' , Finance =  '" & Finance & "',Acc_Type=N'" & AcTypeLao & "',Acc_TypeE=N'" & AcTypeEng & "' WHERE AC_CODE = '" & txtAc_code.Text.Trim & "'")
            If RSC.State = ConnectionState.Open Then RSC.Close()
            MsgBox("ການແກ້ໄຂຮຽບຮ້ອນ !", MsgBoxStyle.OkOnly)
        End If
    End Sub

    Private Sub txtAccName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAccName.KeyPress
        If e.KeyChar = Chr(13) Then
            txtAccName_E.Focus()
        End If
    End Sub

    Private Sub FrNewAcc_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetControlText(Me)
        If txtAc_code.Enabled = True Then
            AddNew()
        Else
            LoadText()
        End If
        Label4.Text = "WISE Orginal"
    End Sub
    Private Sub LoadText()

        txtAccName.Text = ""
        txtAccName_E.Text = ""
        LoadSqlData("SELECT * FROM Acc_Code WHERE AC_CODE = '" & txtAc_code.Text & "'", RSC)
        With RSC
            Do Until .EOF = True
                txtAccName.Text = Trim(.Fields("Name_L").Value)
                txtAccName_E.Text = Trim(.Fields("Name_E").Value.ToString)
                Finance = Trim(.Fields("Finance").Value.ToString)
                TxtWISE_Orginal.Text = Trim(.Fields("Ac_Original").Value.ToString)
                .MoveNext()
            Loop
        End With
        'Call LoadFinance()
    End Sub
    Private Sub AddNew()
        TxtWISE_Orginal.Clear()
        txtAc_code.Clear()
        txtAccName.Clear()
        txtAccName_E.Clear()
    End Sub

    Private Sub txtAc_code_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAc_code.KeyPress
        If e.KeyChar = Chr(13) Then
            txtAccName.Focus()
        End If
    End Sub

    Private Sub BntNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BntNew.Click
        txtAc_code.Enabled = True
    End Sub

End Class