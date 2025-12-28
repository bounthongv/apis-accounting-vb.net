Public Class FrmCustomer

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()

    End Sub

    Private Sub FrmCustomer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FG.FormatString = "^No. |< Code       |< Lao Name                             |< English Name                              |< Tel                   |< Fax                 |< Email                |< Website           "
        LoadListFG()
    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtFax.TextChanged

    End Sub

    Private Sub BtnAddNew2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNew2.Click
        Call AddNew()
    End Sub
    Private Sub AddNew()
        TxtCode.Text = ""
        TxtName.Text = ""
        TxtNameE.Text = ""
        TxtAdd.Text = ""
        TxtAddE.Text = ""
        TxtTel.Text = ""
        TxtFax.Text = ""
        TxtEmail.Text = ""
        TxtWebsite.Text = ""
        TxtOther.Text = ""
        TxtCode.Enabled = True
        TxtCode.Focus()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TxtCode.Text = "" Then MsgBox("ກະລຸນາເລືອກກ່ອນ!", MsgBoxStyle.OkOnly) : Exit Sub
        If MessageBox.Show("ທ່ານຕ້ອງລຶບລະຫັດ " & FG.get_TextMatrix(FG.Row, 1) & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            CNN.Execute("DELETE FROM Customer WHERE Code=N'" & FG.get_TextMatrix(FG.Row, 1) & "'")
        
            LoadListFG()
            Call AddNew()
        End If
    End Sub
    Public Sub LoadListFG()
        FG.Rows = 1
        With RSC
            Call LoadSqlData("SELECT * FROM  Customer order by Code ASC  ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("Code").Value)) & _
                      vbTab & Trim(CStr(.Fields("Name").Value.ToString)) & _
                       vbTab & Trim(CStr(.Fields("NameE").Value.ToString)) & _
                          vbTab & Trim(CStr(.Fields("Tel").Value.ToString)) & _
                             vbTab & Trim(CStr(.Fields("Fax").Value.ToString)) & _
                                vbTab & Trim(CStr(.Fields("Email").Value.ToString)) & _
                      "" & vbTab & ((.Fields("Other").Value.ToString)))
                    .MoveNext()
                End While
            Else
                FG.Rows = 2
            End If
        End With

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        If TxtCode.Text = "" Then MsgBox("", MsgBoxStyle.Exclamation) : TxtCode.Focus() : Exit Sub

        If TxtCode.Enabled = True Then
            Call LoadSqlData("SELECT * FROM Customer WHERE Code =N'" & Trim(TxtCode.Text) & "'", RSC)
            If RSC.RecordCount <> 0 Then
                MsgBox("ລະຫັດມີແລ້ວ!", MsgBoxStyle.Exclamation) : TxtCode.Focus() : Exit Sub
            End If
        End If


        Call LoadSqlData("SELECT * FROM Customer WHERE Code =N'" & Trim(TxtCode.Text) & "'", RSC)
        If RSC.RecordCount = 0 Then
            CNN.Execute("INSERT INTO Customer( Code, Name, NameE, Address, AddressE, Tel, Fax, Email, Other) " & _
                "Values(N'" & Trim(TxtCode.Text) & "',N'" & Trim(TxtName.Text) & "' , N'" & Trim(TxtNameE.Text) & "' ,N'" & Trim(TxtAdd.Text) & "',N'" & Trim(TxtAddE.Text) & "',N'" & Trim(TxtTel.Text) & "',N'" & Trim(TxtFax.Text) & "',N'" & Trim(TxtEmail.Text) & "',N'" & Trim(TxtOther.Text) & "')")
        Else
            CNN.Execute("UPDATE Customer SET  Name=N'" & Trim(TxtName.Text) & "' , NameE=N'" & Trim(TxtNameE.Text) & "' ,Address=N'" & Trim(TxtAdd.Text) & "',AddressE=N'" & Trim(TxtAddE.Text) & "'," & _
                        "Tel=N'" & Trim(TxtTel.Text) & "',Fax=N'" & Trim(TxtFax.Text) & "',Email=N'" & Trim(TxtEmail.Text) & "',Other=N'" & Trim(TxtOther.Text) & "'  " & _
                        " WHERE Code =N'" & Trim(TxtCode.Text) & "'")
        End If
        If RSC.State = ConnectionState.Open Then RSC.Close()
        MsgBox("ການບັນທຶກສຳເລັດ!", MsgBoxStyle.OkOnly)
        TxtCode.Focus()
        LoadListFG()
    End Sub

    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelChange
        TxtCode.Text = FG.get_TextMatrix(FG.Row, 1)
        TxtName.Text = FG.get_TextMatrix(FG.Row, 2)
        Call LoadText()
        TxtCode.Enabled = False
    End Sub
    Private Sub LoadText()
        Call LoadSqlData("SELECT * FROM Customer WHERE Code =N'" & Trim(TxtCode.Text) & "'", RSC)
        If RSC.RecordCount = 0 Then
            AddNew()
        Else
            TxtCode.Text = Trim(RSC.Fields("Code").Value.ToString)
            TxtName.Text = Trim(RSC.Fields("Name").Value.ToString)
            TxtNameE.Text = Trim(RSC.Fields("NameE").Value.ToString)
            TxtAdd.Text = Trim(RSC.Fields("Address").Value.ToString)
            TxtAddE.Text = Trim(RSC.Fields("AddressE").Value.ToString)
            TxtTel.Text = Trim(RSC.Fields("Tel").Value.ToString)
            TxtFax.Text = Trim(RSC.Fields("Fax").Value.ToString)
            TxtEmail.Text = Trim(RSC.Fields("Email").Value.ToString)
            TxtOther.Text = Trim(RSC.Fields("Other").Value.ToString)
        End If
    End Sub
End Class