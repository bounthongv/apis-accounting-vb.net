Public Class FrmDepartmentNew
    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        If txtCompany.Text = "" Then
            MsgBox("ກະລຸນາເລືອກສໍານັກງານກ່ອນ", MsgBoxStyle.Exclamation) : Exit Sub
        End If

        If txtNm.Text = "" Then MsgBox("ກະລຸນາໃສ່ຊື່ພາສາລາວກ່ອນ") : txtNm.Focus() : Exit Sub
        If txtNmE.Text = "" Then MsgBox("ກະລຸນາໃສ່ຊື່ພາສາອັງກິດກ່ອນ") : txtNmE.Focus() : Exit Sub
        Dim cRS As New ADODB.Recordset
        If txtID.Enabled = True Then
            'RunID()
            Call LoadSqlData("SELECT * FROM Department WHERE DepartmentID =N'" & txtID.Text & "'  ", RSC)
            If RSC.RecordCount <> 0 Then
                MsgBox("ລຫັດນີ້ມີແລ້ວ", MsgBoxStyle.Exclamation) : txtID.Focus() : Exit Sub
            End If
            CNN.Execute("INSERT INTO Department(DepartmentID, DepartmentNm,DepartmentNmE, DepartmentRemark,Company) " & _
                        " VALUES('" & Trim(txtID.Text.ToString) & "', N'" & Trim(txtNm.Text.ToString) & "', N'" & Trim(txtNmE.Text.ToString) & "', N'" & Trim(txtRemark.Text.ToString) & "','" & Trim(txtCompany.Text) & "')")

        Else
            Dim ss As String
            ss = "Update Department Set DepartmentNm=N'" & Trim(txtNm.Text) & "',DepartmentNmE=N'" & Trim(txtNmE.Text.ToString) & "', DepartmentRemark=N'" & Trim(txtRemark.Text.ToString) & "',Company='" & Trim(txtCompany.Text) & "' Where DepartmentID='" & Trim(txtID.Text) & "' "
            CNN.Execute(ss)
            CNN.Execute("Update assets set company='" & Trim(txtCompany.Text) & "' where DepartmentID='" & Trim(txtID.Text) & "' ")
        End If
        txtID.Enabled = False
        MsgBox("Finish")
    End Sub

    Private Sub RunID()
        Dim rRS As New ADODB.Recordset
        Dim mNum2 As String
        Call LoadSqlData("Select Top 1 DepartmentID From Department Order by DepartmentID DESC ", rRS)
        If rRS.RecordCount <> 0 Then
            mNum2 = Val(rRS.Fields("DepartmentID").Value) + 1
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
        'SetControlText(Me)
        'ChgChildForm()
        LdCompany()
        If mEdit = True Then
            Call loadAST()
            txtID.Enabled = False
        Else
            Call cmdNew_Click(sender, e)
        End If

    End Sub
    Private Sub LdCompany()
        Dim gRS As New ADODB.Recordset
        CmbCompany.Items.Clear()
        Call LoadSqlData("Select * from AP_Office  where Off_ID<>'00'  Order by Off_ID", gRS)
        If gRS.RecordCount <> 0 Then
            While Not gRS.EOF
                CmbCompany.Items.Add(gRS.Fields("Off_Name").Value)
                gRS.MoveNext()
            End While
        End If
      
    End Sub
    Private Sub loadAST()
        Dim aRS As New ADODB.Recordset
        Call LoadSqlData("Select * from Department where DepartmentID='" & myTemp & "' ", aRS)
        With aRS
            If .RecordCount <> 0 Then
                txtID.Text = Trim(.Fields("DepartmentID").Value.ToString)
                txtNm.Text = Trim(.Fields("DepartmentNm").Value.ToString)
                txtNmE.Text = Trim(.Fields("DepartmentNmE").Value.ToString)
                txtRemark.Text = Trim(.Fields("DepartmentRemark").Value.ToString)
                txtCompany.Text = Trim(.Fields("Company").Value.ToString)
            End If
        End With
     
        Call LoadSqlData("SELECT * FROM AP_Office where Off_ID='" & Trim(txtCompany.Text) & "'", RSC)
        If RSC.RecordCount > 0 Then
            CmbCompany.Text = RSC.Fields("Off_Nm").Value.ToString
        End If
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub cmdNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNew.Click
        'Call RunID()
        txtID.Text = ""
        txtNm.Text = ""
        txtNmE.Text = ""
        txtID.Enabled = True
        txtID.Focus()
    End Sub

    Private Sub CmbCompany_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCompany.SelectedIndexChanged
        Dim dRS As New ADODB.Recordset
        Call LoadSqlData("Select * from AP_Office Where Off_Name = N'" & Trim(CmbCompany.Text) & "' ", dRS)
        If dRS.RecordCount <> 0 Then
            txtCompany.Text = Trim(dRS.Fields("Off_ID").Value.ToString)
        End If
 
    End Sub
End Class