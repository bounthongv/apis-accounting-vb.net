Public Class FrmSupplier

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()

    End Sub

    Private Sub FrmSupplier_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetupDataGridView()
        LoadListFG()
    End Sub

    Private Sub SetupDataGridView()
        ' Clear existing columns
        FG.Columns.Clear()

        ' Add columns with headers
        FG.Columns.Add("No", "No.")
        FG.Columns.Add("Code", "Code")
        FG.Columns.Add("LaoName", "Lao Name")
        FG.Columns.Add("EngName", "English Name")
        FG.Columns.Add("Tel", "Tel")
        FG.Columns.Add("Fax", "Fax")
        FG.Columns.Add("Email", "Email")
        FG.Columns.Add("Website", "Website")
        FG.Columns.Add("Other", "Other")

        ' Set column widths
        FG.Columns(0).Width = 50
        FG.Columns(1).Width = 100
        FG.Columns(2).Width = 150
        FG.Columns(3).Width = 150
        FG.Columns(4).Width = 100
        FG.Columns(5).Width = 100
        FG.Columns(6).Width = 150
        FG.Columns(7).Width = 100
        FG.Columns(8).Width = 100

        ' Configure DataGridView properties
        FG.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG.ReadOnly = True
        FG.AllowUserToAddRows = False
        FG.AllowUserToDeleteRows = False
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

        Dim selectedCode As String = ""
        If FG.CurrentRow IsNot Nothing Then
            selectedCode = FG.CurrentRow.Cells(1).Value.ToString()
        End If

        If MessageBox.Show("ທ່ານຕ້ອງລຶບລະຫັດ " & selectedCode & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            CNN.Execute("DELETE FROM Supplier WHERE Code=N'" & selectedCode & "'")

            LoadListFG()
            Call AddNew()
        End If
    End Sub

    Public Sub LoadListFG()
        ' Clear existing rows
        FG.Rows.Clear()

        With RSC
            Call LoadSqlData("SELECT * FROM  Supplier order by Code ASC  ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    ' Add a new row to the DataGridView
                    Dim rowIndex As Integer = FG.Rows.Add()
                    Dim currentRow As DataGridViewRow = FG.Rows(rowIndex)

                    ' Populate the cells with data
                    currentRow.Cells(0).Value = .AbsolutePosition
                    currentRow.Cells(1).Value = Trim(CStr(.Fields("Code").Value))
                    currentRow.Cells(2).Value = Trim(CStr(.Fields("Name").Value.ToString))
                    currentRow.Cells(3).Value = Trim(CStr(.Fields("NameE").Value.ToString))
                    currentRow.Cells(4).Value = Trim(CStr(.Fields("Tel").Value.ToString))
                    currentRow.Cells(5).Value = Trim(CStr(.Fields("Fax").Value.ToString))
                    currentRow.Cells(6).Value = Trim(CStr(.Fields("Email").Value.ToString))
                    currentRow.Cells(7).Value = "" ' Website column - not in the query
                    currentRow.Cells(8).Value = Trim(CStr(.Fields("Other").Value.ToString))

                    .MoveNext()
                End While
            Else
                ' Add an empty row if no data
                FG.Rows.Add()
            End If
        End With
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        If TxtCode.Text = "" Then MsgBox("", MsgBoxStyle.Exclamation) : TxtCode.Focus() : Exit Sub

        If TxtCode.Enabled = True Then
            Call LoadSqlData("SELECT * FROM Supplier WHERE Code =N'" & Trim(TxtCode.Text) & "'", RSC)
            If RSC.RecordCount <> 0 Then
                MsgBox("ລະຫັດມີແລ້ວ!", MsgBoxStyle.Exclamation) : TxtCode.Focus() : Exit Sub
            End If
        End If


        Call LoadSqlData("SELECT * FROM Supplier WHERE Code =N'" & Trim(TxtCode.Text) & "'", RSC)
        If RSC.RecordCount = 0 Then
            CNN.Execute("INSERT INTO Supplier( Code, Name, NameE, Address, AddressE, Tel, Fax, Email, Other) " & _
                "Values(N'" & Trim(TxtCode.Text) & "',N'" & Trim(TxtName.Text) & "' , N'" & Trim(TxtNameE.Text) & "' ,N'" & Trim(TxtAdd.Text) & "',N'" & Trim(TxtAddE.Text) & "',N'" & Trim(TxtTel.Text) & "',N'" & Trim(TxtFax.Text) & "',N'" & Trim(TxtEmail.Text) & "',N'" & Trim(TxtOther.Text) & "')")
        Else
            CNN.Execute("UPDATE Supplier SET  Name=N'" & Trim(TxtName.Text) & "' , NameE=N'" & Trim(TxtNameE.Text) & "' ,Address=N'" & Trim(TxtAdd.Text) & "',AddressE=N'" & Trim(TxtAddE.Text) & "'," & _
                        "Tel=N'" & Trim(TxtTel.Text) & "',Fax=N'" & Trim(TxtFax.Text) & "',Email=N'" & Trim(TxtEmail.Text) & "',Other=N'" & Trim(TxtOther.Text) & "'  " & _
                        " WHERE Code =N'" & Trim(TxtCode.Text) & "'")
        End If
        If RSC.State = ConnectionState.Open Then RSC.Close()
        MsgBox("ການບັນທຶກສຳເລັດ!", MsgBoxStyle.OkOnly)
        TxtCode.Focus()
        LoadListFG()
    End Sub

    Private Sub FG_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles FG.SelectionChanged
        If FG.CurrentRow IsNot Nothing Then
            TxtCode.Text = FG.CurrentRow.Cells(1).Value.ToString()
            TxtName.Text = FG.CurrentRow.Cells(2).Value.ToString()
            Call LoadText()
            TxtCode.Enabled = False
        End If
    End Sub

    Private Sub LoadText()
        Call LoadSqlData("SELECT * FROM Supplier WHERE Code =N'" & Trim(TxtCode.Text) & "'", RSC)
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