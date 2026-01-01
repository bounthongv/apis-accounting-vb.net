Public Class FmRestorData
    Dim sql As String
    Dim a, k As String
    Dim b As Integer
    Private Sub cmdrestore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdrestore.Click
        ConnectSQL()
        RestorData1()
        RestorData()
        ConnectSQL()
    End Sub

    Private Sub RestorData1()

        If txtSaveIn.Text = "" Then
            MsgBox("select File")
            Dim Fl As New OpenFileDialog
            Fl.ShowDialog()
            TextBox1.Text = Fl.SafeFileName
            txtSaveIn.Text = Fl.FileName
            Exit Sub
        End If
        Dim strSQL As String
        Dim strCon As String
        strCon = " Password = " & MDServerPassword & "; Integrated Security= True; " & _
     "User ID = " & MDServerUser & "; Initial Catalog = " & "master" & "; Data Source =" & MDServerName & ""
        Dim cmdRestore As SqlClient.SqlCommand = New SqlClient.SqlCommand
        Dim SqlConnection1 As SqlClient.SqlConnection = New SqlClient.SqlConnection
        SqlConnection1.ConnectionString = strCon
        SqlConnection1.Open()
        cmdRestore.Connection = SqlConnection1
        Cursor = Cursors.WaitCursor
        Try
            strSQL = "ALTER DATABASE " & MDDatabaName & " SET SINGLE_USER"
            cmdRestore.CommandText = strSQL
            cmdRestore.ExecuteNonQuery()
            strSQL = "RESTORE DATABASE " & MDDatabaName & "  "
            strSQL &= "FROM DISK = '" & txtSaveIn.Text & "' "
            cmdRestore.CommandText = strSQL
            cmdRestore.ExecuteNonQuery()
            'MsgBox("Restor databases finish")
        Catch ex As Exception
            'MsgBox("Restor databases Error")
        Finally
            strSQL = "ALTER DATABASE " & MDDatabaName & "  SET MULTI_USER"
            cmdRestore.CommandText = strSQL
            cmdRestore.ExecuteNonQuery()
        End Try
        Cursor = Cursors.Arrow
        SqlConnection1.Close()
        cmdRestore.Dispose()
        cmdRestore = Nothing
    End Sub


    Private Sub RestorData()

        If txtSaveIn.Text = "" Then
            MsgBox("select File")
            Dim Fl As New OpenFileDialog
            Fl.ShowDialog()
            TextBox1.Text = Fl.SafeFileName
            txtSaveIn.Text = Fl.FileName
            Exit Sub
        End If
        Dim strSQL As String
        Dim strCon As String
        strCon = " Password = " & MDServerPassword & "; Integrated Security= True; " & _
     "User ID = " & MDServerUser & "; Initial Catalog = " & "master" & "; Data Source =" & MDServerName & ""
  
        Dim cmdRestore As SqlClient.SqlCommand = New SqlClient.SqlCommand
        Dim SqlConnection1 As SqlClient.SqlConnection = New SqlClient.SqlConnection
        SqlConnection1.ConnectionString = strCon
        SqlConnection1.Open()
        cmdRestore.Connection = SqlConnection1
        Cursor = Cursors.WaitCursor
        Try
            strSQL = "ALTER DATABASE " & MDDatabaName & " SET SINGLE_USER"
            cmdRestore.CommandText = strSQL
            cmdRestore.ExecuteNonQuery()
            strSQL = "RESTORE DATABASE " & MDDatabaName & "  "
            strSQL &= "FROM DISK = '" & txtSaveIn.Text & "' "
            cmdRestore.CommandText = strSQL
            cmdRestore.ExecuteNonQuery()

            MsgBox("Restor databases finish")
        Catch ex As Exception
       
            MsgBox("Restor databases Error")
        Finally
            strSQL = "ALTER DATABASE " & MDDatabaName & "  SET MULTI_USER"
            cmdRestore.CommandText = strSQL
            cmdRestore.ExecuteNonQuery()
        End Try
        Cursor = Cursors.Arrow
        SqlConnection1.Close()
        cmdRestore.Dispose()
        cmdRestore = Nothing
    End Sub
    Private Sub FmRestorData_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetupGrid()
   
        a = My.Application.Info.DirectoryPath
        k = Mid(a, CDbl(Len(a)) - 8, 3)
        'MsgBox(k)
        If k = "bin" Then
            b = CDbl(Len(a)) - 10
        Else
            b = CDbl(Len(a)) - 6
        End If
        txtSaveIn.Text = Microsoft.VisualBasic.Left(a, b) & "\BackUp Data"
        'txtFileNane.Text = MDDatabaName & " BackUp" & "-" & DateString




    End Sub

    Private Sub cmdsearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdsearch.Click
        Dim Fl As New OpenFileDialog
        Fl.ShowDialog()
        TextBox1.Text = Fl.SafeFileName
        txtSaveIn.Text = Fl.FileName
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        LoadListFG()
    End Sub


    Private Sub LoadListFG()
        sql = ""
        'sql = " AND Back_UpData.File_Date   BETWEEN '" & Format(dts.Value, "MM-dd-yyyy") & "' AND '" & Format(dtt.Value, "MM-dd-yyyy") & "' "

        Dim R As New ADODB.Recordset
        FG.Rows.Clear()
        With R
            Call LoadAcData("SELECT * FROM  Back_UpData where File_Nme <>'""' " & sql & "  ", R)
            If .RecordCount <> 0 Then
                While Not .EOF
                    FG.Rows.Add(.AbsolutePosition, _
                                Format(CDate(.Fields("File_Date").Value.ToString), "dd/MM/yyyy"), _
                                (.Fields("File_Nme").Value.ToString))
                    .MoveNext()
                End While
            Else
                'FG.Rows.Clear() ' Not needed, as Clear() is already called, or if there's specific handling, add it.
            End If
        End With
    End Sub

    Private Sub FG_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelectionChanged
        If FG.CurrentRow IsNot Nothing Then
            TextBox1.Text = FG.CurrentRow.Cells(2).Value.ToString()
            txtSaveIn.Text = Microsoft.VisualBasic.Left(a, b) & "\BackUp Data\" & FG.CurrentRow.Cells(2).Value.ToString()
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            'cmdsearch.Enabled = True
            'dts.Enabled = False
            'dtt.Enabled = False
            'Button1.Enabled = False
            'FG.Editable=
        End If
    End Sub

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Close()
    End Sub

    Private Sub SetupGrid()
        FG.Columns.Clear()
        FG.Rows.Clear()
        FG.AllowUserToAddRows = False
        FG.AllowUserToDeleteRows = False
        FG.ReadOnly = True
        FG.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG.MultiSelect = False
        FG.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        FG.RowHeadersVisible = False

        Dim c As DataGridViewColumn

        c = New DataGridViewTextBoxColumn() : c.Name = "No" : c.HeaderText = "ລ/ດ" : c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter : FG.Columns.Add(c)
        c = New DataGridViewTextBoxColumn() : c.Name = "Date" : c.HeaderText = "ວັນທີ" : c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft : FG.Columns.Add(c)
        c = New DataGridViewTextBoxColumn() : c.Name = "FileName" : c.HeaderText = "ຊື່ເອກະສານ" : c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft : c.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill : FG.Columns.Add(c)

        UpdateColumnHeaders()
    End Sub

    Private Sub UpdateColumnHeaders()
        If FG.Columns.Count < 3 Then Exit Sub
        If MuLng = "L" Then
            FG.Columns("No").HeaderText = "ລ/ດ"
            FG.Columns("Date").HeaderText = "ວັນທີ"
            FG.Columns("FileName").HeaderText = "ຊື່ເອກະສານ"
        Else
            FG.Columns("No").HeaderText = "No"
            FG.Columns("Date").HeaderText = "Date"
            FG.Columns("FileName").HeaderText = "File Name"
        End If
    End Sub
End Class