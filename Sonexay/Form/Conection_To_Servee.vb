Option Explicit On
Option Strict On
Public Class Conection_To_Servee
    Dim rsProj As New ADODB.Recordset
    Dim sql As String
    Public editProj As Boolean
    Public conn As New ADODB.Connection
    Public Sub Connect()
        'Dim strConn As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =Conection.mdb;Persist Security Info=False"
        Dim strConn As String = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = Connection.mdb;Persist Security Info=False;Jet OLEDB:Database Password=2459428"
        Try
            If conn.State = ConnectionState.Open Then conn.Close()
            conn.Open(strConn)
        Catch ex As Exception
            MessageBox.Show("ບໍ່ສາມາດເຊື່ອມຕໍ່ກັບຖານຂໍ້ມູນໄດ້ຍ້ອນ: " & vbNewLine _
            & ex.ToString, "ການເຊື່ອມຕໍ່ກັບຖານຂໍ້ມູນຜິດພາດ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub loadrs(ByVal sql As String, ByVal rs As ADODB.Recordset)
        With rs
            If .State = ConnectionState.Open Then .Close()
            .ActiveConnection = conn
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(sql)
            .Requery()
        End With
    End Sub

    Private Sub btnConect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConect.Click
        Dim rs As New ADODB.Recordset
        Call loadrs("select * from Conect where SvID='" & "001" & "'", rs)

        conn.Execute("Update Conect Set ServerName ='" & txtServerName.Text.ToString & "', DataBaseName='" & txtDatabaseName.Text.ToString & "', UserName='" & txtServerUser.Text.ToString & "', UserPassword='" & txtServerPassword.Text.ToString & "' " & _
                     " WHERE SvID='" & "001" & "' ")

        MsgBox("Conect To Server Completed!")

        Me.Close()
     
    End Sub
    Private Sub LoadDatabaseServer()
        Call loadrs("Select * from Conect WHERE SvID='" & "001" & "' ", rsProj)
        With rsProj
            If .RecordCount <> 0 Then
                txtServerName.Text = (.Fields("ServerName").Value.ToString)
                txtDatabaseName.Text = (.Fields("DatabaseName").Value.ToString)
                txtServerUser.Text = (.Fields("UserName").Value.ToString)
                txtServerPassword.Text = (.Fields("UserPassword").Value.ToString)
            End If
        End With
    End Sub

    Private Sub Conection_To_Servee_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Connect()
        Call LoadDatabaseServer()
        LoadListFG()



        If CheckBox1.Checked = True Then
            TextBox1.Enabled = True
            txtSaveIn.Enabled = True
            cmdsearch.Enabled = True
            cmdrestore.Enabled = True
            'RdBackUp.Enabled = True
            'RdRestor.Enabled = True
        Else
            TextBox1.Enabled = False
            txtSaveIn.Enabled = False
            cmdsearch.Enabled = False
            cmdrestore.Enabled = False
            'RdBackUp.Enabled = False
            'RdRestor.Enabled = False

        End If
    End Sub
 
    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Close()
    End Sub

    Private Sub txtServerName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtServerName.KeyPress
        If e.KeyChar = Chr(13) Then
            CvConsonant = txtServerName.Text
            LoadCvConsonant()
            txtServerName.Text = CvConsonant
        End If
    End Sub

    Private Sub txtServerName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtServerName.LostFocus
        CvConsonant = txtServerName.Text
        LoadCvConsonant()
        txtServerName.Text = CvConsonant
    End Sub

    Private Sub cmdrestore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdrestore.Click
        If RdRestor.Checked = True Then
            Call RestorData()
        Else
            Call BackUpData()
        End If
    End Sub
    Private Sub BackUpData()

        If txtSaveIn.Text = "" Then
            MsgBox("select Folder")
            Dim Fl As New OpenFileDialog
            Fl.ShowDialog()
            TextBox1.Text = Fl.SafeFileName
            txtSaveIn.Text = Fl.FileName
            Exit Sub
        End If

        If TextBox1.Text = "" Then
            MsgBox("file Name")
            TextBox1.Focus()
            Exit Sub
        End If


        Dim con As SqlClient.SqlConnection = New SqlClient.SqlConnection("Data Source= " & txtServerName.Text & " ;Integrated Security=SSPI;Initial Catalog= " & txtDatabaseName.Text & " ")
        Dim cmd As SqlClient.SqlCommand = New SqlClient.SqlCommand("backup database " & txtDatabaseName.Text & " to disk='" & txtSaveIn.Text & "\" & TextBox1.Text & ".bak" & "'", con)
        con.Open()
        cmd.ExecuteNonQuery()
        con.Close()
        'conn.Execute("Insert inTo Back_UpData ( File_Nme , File_Add , File_Date ) Values ('" & TextBox1.Text & "' , '" & txtSaveIn.Text & "' , '" & Format(MWorkSetting, "dd/MM/yyyy") & "')")
        MsgBox("Restor finish")
    End Sub


    Private Sub LoadListFG()
        'sql = ""
        'Dim RSC As New ADODB.Recordset
        'FG.Rows = 1
        'With RSC
        '    Call loadrs("SELECT * FROM  Back_UpData  ", RSC)
        '    If .RecordCount > 0 Then
        '        While Not .EOF
        '            FG.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("File_Nme").Value)) & _
        '                         "" & vbTab & Trim(CStr(.Fields("File_Add").Value)) & _
        '                       "" & vbTab & Trim(CStr(.Fields("File_Date").Value)))
        '            .MoveNext()
        '        End While
        '    Else
        '        FG.Rows = 16
        '    End If
        'End With
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


        'Exit Sub
        'If txtSaveIn.Text = "" Then MsgBox("plese select folder") 
        Dim strSQL As String
        Dim strCon As String
        'strCon = "Data Source=NITHI;Initial Catalog=master;Integrated Security=True"
        strCon = " Password = " & txtServerPassword.Text & "; Integrated Security= True; " & _
     "User ID = " & txtServerUser.Text & "; Initial Catalog = " & "master" & "; Data Source =" & txtServerName.Text & ""
        Dim cmdRestore As SqlClient.SqlCommand = New SqlClient.SqlCommand
        Dim SqlConnection1 As SqlClient.SqlConnection = New SqlClient.SqlConnection
        SqlConnection1.ConnectionString = strCon
        SqlConnection1.Open()
        cmdRestore.Connection = SqlConnection1
        Cursor = Cursors.WaitCursor
        Try
            strSQL = "ALTER DATABASE " & txtDatabaseName.Text & " SET SINGLE_USER"
            cmdRestore.CommandText = strSQL
            cmdRestore.ExecuteNonQuery()
            strSQL = "RESTORE DATABASE " & txtDatabaseName.Text & "  "
            strSQL &= "FROM DISK = '" & txtSaveIn.Text & "' "
            cmdRestore.CommandText = strSQL
            cmdRestore.ExecuteNonQuery()
            MsgBox("Restor databases finish")
        Catch ex As Exception
            MsgBox("Restor databases Error")
        Finally
            strSQL = "ALTER DATABASE " & txtDatabaseName.Text & "  SET MULTI_USER"
            cmdRestore.CommandText = strSQL
            cmdRestore.ExecuteNonQuery()
        End Try
        Cursor = Cursors.Arrow
        SqlConnection1.Close()
        cmdRestore.Dispose()
        cmdRestore = Nothing
    End Sub

    Private Sub cmdsearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdsearch.Click

        If RdBackUp.Checked Then
            Dim dl As New FolderBrowserDialog
            dl.ShowDialog()
            txtSaveIn.Text = dl.SelectedPath
        Else
            Dim Fl As New OpenFileDialog
            Fl.ShowDialog()
            TextBox1.Text = Fl.SafeFileName
            txtSaveIn.Text = Fl.FileName
        End If

     
      
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox1.Enabled = True
            txtSaveIn.Enabled = True
            cmdsearch.Enabled = True
            cmdrestore.Enabled = True
            'RdBackUp.Enabled = True
            'RdRestor.Enabled = True
        Else
            TextBox1.Enabled = False
            txtSaveIn.Enabled = False
            cmdsearch.Enabled = False
            cmdrestore.Enabled = False
            'RdBackUp.Enabled = False
            'RdRestor.Enabled = False

        End If
    End Sub

    Private Sub RdBackUp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RdBackUp.CheckedChanged
        txtSaveIn.Clear()
        TextBox1.Clear()
    End Sub

    Private Sub txtServerName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtServerName.TextChanged

    End Sub
End Class
