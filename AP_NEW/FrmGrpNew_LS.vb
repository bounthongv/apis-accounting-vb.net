Public Class FrmGrpNew_LS
    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        If txtNm.Text = "" Then MsgBox("ກະລຸນາໃສ່ຊື່ພາສາລາວກ່ອນ") : txtNm.Focus() : Exit Sub
        If txtNmE.Text = "" Then MsgBox("ກະລຸນາໃສ່ຊື່ພາສາອັງກິດກ່ອນ") : txtNmE.Focus() : Exit Sub
        Dim cRS As New ADODB.Recordset
        If txtID.Enabled = True Then
            'RunID()

            Call LoadSqlData("SELECT * FROM Groups_Asset WHERE Group_ID =N'" & txtID.Text & "'  ", RSC)
            If RSC.RecordCount <> 0 Then
                MsgBox("ລຫັດນີ້ມີແລ້ວ", MsgBoxStyle.Exclamation) : txtID.Focus() : Exit Sub
            End If

            CNN.Execute("INSERT INTO Groups_Asset(Group_ID, Grp_NO, Group_Nm, Group_NmE,AccountCodeAsDR,AccountCodeAsCR,Ac_Code,Dep_Code,AccountCodeBrokenDR) " & _
                        " VALUES('" & Trim(txtID.Text) & "','" & Trim(TxtGrp_No.Text) & "', N'" & Trim(txtNm.Text) & "',N'" & Trim(txtNmE.Text) & "','" & Trim(TxtCodeAsDR.Text) & "','" & Trim(TxtCodeAsCR.Text.ToString) & "','" & Trim(txtAcc.Text.ToString) & "','" & Trim(TxtDep_Code.Text.ToString) & "','" & Trim(txtCodeBrokenDR.Text.ToString) & "')")
        Else
            Dim ss As String
            ss = "Update Groups_Asset Set Group_Nm=N'" & Trim(txtNm.Text) & "',Grp_No=N'" & Trim(TxtGrp_No.Text) & "',Group_NmE=N'" & Trim(txtNmE.Text.ToString) & "', AccountCodeAsDR='" & Trim(TxtCodeAsDR.Text.ToString) & "',AccountCodeAsCR='" & Trim(TxtCodeAsCR.Text.ToString) & "',Ac_Code='" & Trim(txtAcc.Text.ToString) & "',Dep_Code='" & Trim(TxtDep_Code.Text.ToString) & "',AccountCodeBrokenDR='" & Trim(txtCodeBrokenDR.Text.ToString) & "' Where Group_ID='" & Trim(txtID.Text.ToString) & "' "
            CNN.Execute(ss)
        End If
        CNN.Execute(" delete Groups ")
        Dim DD As String
        DD = "INSERT INTO Groups ( Group_ID, Grp_NO, Group_Nm, Group_NmE,AccountCodeAsDR,AccountCodeAsCR,Ac_Code,Dep_Code,AccountCodeBrokenDR ) " & _
        " select Group_ID, Grp_NO, Group_Nm, Group_NmE,AccountCodeAsDR,AccountCodeAsCR,Ac_Code,Dep_Code,AccountCodeBrokenDR from Groups_Asset  "

        txtID.Enabled = False
        MsgBox("Finish")
    End Sub

    Private Sub RunID()
        Dim rRS As New ADODB.Recordset
        Dim mNum2 As String
        Call LoadSqlData("Select Top 1 Group_ID From Groups_Asset Order by Group_ID DESC ", rRS)
        If rRS.RecordCount <> 0 Then
            mNum2 = Val(rRS.Fields("Group_ID").Value) + 1
            If Len(CStr(mNum2).Trim) = 1 Then
                txtID.Text = "0" & CStr(mNum2)
            ElseIf Len(CStr(mNum2).Trim) = 2 Then
                txtID.Text = CStr(mNum2)
            End If
        Else
            txtID.Text = "01"
        End If
    End Sub

    Private Sub FrmAssetNew_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Call Loadlang()

        If mEdit = True Then
            Call loadAST()
            txtID.Enabled = False
        Else
            Call cmdNew_Click(sender, e)
        End If
        Label7.Text = "ລະຫັດຫຼັກ"
    End Sub
    Private Sub loadAST()
        Dim aRS As New ADODB.Recordset
        Call LoadSqlData("Select * from Groups_Asset where Group_ID='" & myTemp & "' ", aRS)
        With aRS
            If .RecordCount <> 0 Then
                txtID.Text = Trim(.Fields("Group_ID").Value.ToString)
                txtNm.Text = Trim(.Fields("Group_Nm").Value.ToString)
                txtNmE.Text = Trim(.Fields("Group_NmE").Value.ToString)
                txtAcc.Text = Trim(.Fields("Ac_Code").Value.ToString)
                TxtCodeAsDR.Text = Trim(.Fields("AccountCodeAsDR").Value.ToString)
                TxtCodeAsCR.Text = Trim(.Fields("AccountCodeAsCR").Value.ToString)
                TxtDep_Code.Text = Trim(.Fields("Dep_Code").Value.ToString)
                txtCodeBrokenDR.Text = Trim(.Fields("AccountCodeBrokenDR").Value.ToString)
                TxtGrp_No.Text = Trim(.Fields("Grp_No").Value.ToString)

                Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtCodeAsCR.Text & "' ", RSC)
                If RSC.RecordCount <> 0 Then
                    TxtCrNm.Text = Trim(RSC.Fields("Name_L").Value.ToString)
                End If

                Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtCodeAsDR.Text & "' ", RSC)
                If RSC.RecordCount <> 0 Then
                    TxtDrNm.Text = Trim(RSC.Fields("Name_L").Value.ToString)
                End If

            End If
        End With

    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub cmdNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNew.Click
        Call RunID()
        txtNm.Text = ""
        txtNmE.Text = ""
        TxtGrp_No.Text = ""
        txtAcc.Text = ""
        TxtCodeAsDR.Text = ""
        TxtCodeAsCR.Text = ""
        TxtDep_Code.Text = ""
        txtCodeBrokenDR.Text = ""

        TxtDrNm.Text = ""
        TxtCrNm.Text = ""
        txtID.Enabled = True
        txtNm.Focus()
    End Sub

    Private Sub BtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch.Click
        fmShartOfAccDetail.txtSty.Text = "FrmGrpNew_LS_DR"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        fmShartOfAccDetail.txtSty.Text = "FrmGrpNew_LS_CR"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
    End Sub

    Private Sub TxtCodeAsDR_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCodeAsDR.KeyPress
        If e.KeyChar = Chr(13) Then
            Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtCodeAsDR.Text & "' ", RSC)
            If RSC.RecordCount <> 0 Then
                TxtDrNm.Text = Trim(RSC.Fields("Name_L").Value.ToString)
            Else
                MsgBox("ເລກບັນຊີບໍ່ມີໃນລາລະບານ", MsgBoxStyle.Exclamation) : TxtCodeAsDR.Focus() : Exit Sub
            End If

            TxtCodeAsCR.Focus()
        End If
    End Sub

    Private Sub TxtCodeAsDR_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCodeAsDR.TextChanged

    End Sub

    Private Sub TxtCodeAsCR_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCodeAsCR.KeyPress
        If e.KeyChar = Chr(13) Then
            Call LoadSqlData("SELECT * FROM ACC_CODE WHERE AC_CODE=N'" & TxtCodeAsCR.Text & "' ", RSC)
            If RSC.RecordCount <> 0 Then
                TxtCrNm.Text = Trim(RSC.Fields("Name_L").Value.ToString)
            Else
                MsgBox("ເລກບັນຊີບໍ່ມີໃນລາລະບານ", MsgBoxStyle.Exclamation) : TxtCodeAsCR.Focus() : Exit Sub
            End If

            TxtCodeAsCR.Focus()
        End If
    End Sub

    Private Sub TxtCodeAsCR_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCodeAsCR.TextChanged

    End Sub
End Class