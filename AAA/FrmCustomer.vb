Public Class FrmCustomer

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()

    End Sub

    Private Sub FrmCustomer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        If MessageBox.Show("ທ່ານຕ້ອງລຶບລະຫັດ " & FG.CurrentRow.Cells(1).Value.ToString() & " ແທ້ຫລືບໍ່", "ຄຳຢືນຢັນ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            DbHelper.ExecuteNonQuery("DELETE FROM Customer WHERE Code=N'" & FG.CurrentRow.Cells(1).Value.ToString() & "'")

            LoadListFG()
            Call AddNew()
        End If
    End Sub
    Public Sub LoadListFG()
        ' Clear existing rows
        FG.Rows.Clear()

        Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM  Customer order by Code ASC")
        If dt.Rows.Count > 0 Then
            Dim rowNum As Integer = 1
            For Each row As DataRow In dt.Rows
                ' Add a new row to the DataGridView
                Dim rowIndex As Integer = FG.Rows.Add()
                Dim currentRow As DataGridViewRow = FG.Rows(rowIndex)

                ' Populate the cells with data
                currentRow.Cells(0).Value = rowNum
                currentRow.Cells(1).Value = DbHelper.GetStr(row("Code"))
                currentRow.Cells(2).Value = DbHelper.GetStr(row("Name"))
                currentRow.Cells(3).Value = DbHelper.GetStr(row("NameE"))
                currentRow.Cells(4).Value = DbHelper.GetStr(row("Tel"))
                currentRow.Cells(5).Value = DbHelper.GetStr(row("Fax"))
                currentRow.Cells(6).Value = DbHelper.GetStr(row("Email"))
                currentRow.Cells(7).Value = "" ' Website column - not in the query
                currentRow.Cells(8).Value = DbHelper.GetStr(row("Other"))

                rowNum += 1
            Next
        Else
            ' Add an empty row if no data
            FG.Rows.Add()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        If TxtCode.Text = "" Then MsgBox("", MsgBoxStyle.Exclamation) : TxtCode.Focus() : Exit Sub

        If TxtCode.Enabled = True Then
            Dim dtCheck As DataTable = DbHelper.GetDataTable("SELECT * FROM Customer WHERE Code =N'" & Trim(TxtCode.Text) & "'")
            If dtCheck.Rows.Count <> 0 Then
                MsgBox("ລະຫັດມີແລ້ວ!", MsgBoxStyle.Exclamation) : TxtCode.Focus() : Exit Sub
            End If
        End If


        Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM Customer WHERE Code =N'" & Trim(TxtCode.Text) & "'")
        If dt.Rows.Count = 0 Then
            DbHelper.ExecuteNonQuery("INSERT INTO Customer( Code, Name, NameE, Address, AddressE, Tel, Fax, Email, Other) " & _
                "Values(N'" & Trim(TxtCode.Text) & "',N'" & Trim(TxtName.Text) & "' , N'" & Trim(TxtNameE.Text) & "' ,N'" & Trim(TxtAdd.Text) & "',N'" & Trim(TxtAddE.Text) & "',N'" & Trim(TxtTel.Text) & "',N'" & Trim(TxtFax.Text) & "',N'" & Trim(TxtEmail.Text) & "',N'" & Trim(TxtOther.Text) & "')")
        Else
            DbHelper.ExecuteNonQuery("UPDATE Customer SET  Name=N'" & Trim(TxtName.Text) & "' , NameE=N'" & Trim(TxtNameE.Text) & "' ,Address=N'" & Trim(TxtAdd.Text) & "',AddressE=N'" & Trim(TxtAddE.Text) & "'," & _
                        "Tel=N'" & Trim(TxtTel.Text) & "',Fax=N'" & Trim(TxtFax.Text) & "',Email=N'" & Trim(TxtEmail.Text) & "',Other=N'" & Trim(TxtOther.Text) & "'  " & _
                        " WHERE Code =N'" & Trim(TxtCode.Text) & "'")
        End If
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
        Dim dt As DataTable = DbHelper.GetDataTable("SELECT * FROM Customer WHERE Code =N'" & Trim(TxtCode.Text) & "'")
        If dt.Rows.Count = 0 Then
            AddNew()
        Else
            TxtCode.Text = DbHelper.GetStr(dt.Rows(0)("Code"))
            TxtName.Text = DbHelper.GetStr(dt.Rows(0)("Name"))
            TxtNameE.Text = DbHelper.GetStr(dt.Rows(0)("NameE"))
            TxtAdd.Text = DbHelper.GetStr(dt.Rows(0)("Address"))
            TxtAddE.Text = DbHelper.GetStr(dt.Rows(0)("AddressE"))
            TxtTel.Text = DbHelper.GetStr(dt.Rows(0)("Tel"))
            TxtFax.Text = DbHelper.GetStr(dt.Rows(0)("Fax"))
            TxtEmail.Text = DbHelper.GetStr(dt.Rows(0)("Email"))
            TxtOther.Text = DbHelper.GetStr(dt.Rows(0)("Other"))
        End If
    End Sub
End Class