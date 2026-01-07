Imports MySql.Data.MySqlClient
Imports System.Data.SqlClient
Module MDSQLConnection
    Public CNN As New ADODB.Connection
    Public RSC As New ADODB.Recordset
    Public Comm As ADODB.Command
    ' LEGACY SUPPORT - DO NOT REMOVE UNTIL PHASE 4
    Public sqlCNN As New SqlConnection
    Public Sub ConnectSQL()
        With CNN
            On Error GoTo hang
            If .State = ConnectionState.Open Then .Close()
hang:
            If Err.Number = 0 Then
                VSysError = False
                .ConnectionString = "Provider = SQLOLEDB.1; Password = " & MDServerPassword & "; Persist Security Info = True; " & _
            "User ID = " & MDServerUser & "; Initial Catalog = " & MDDatabaName & "; Data Source =" & MDServerName & ""
                .CommandTimeout = 32767
                .ConnectionTimeout = 32767
                .Open()
            Else
                VSysError = True
                FmLogin.Visible = False
                MessageBox.Show("ຕິດຕໍ່ຖານຂໍ້ມູນບໍ່ຄດ້")
                Conection_To_Servee.ShowDialog()
            End If
        End With
        
        ' Initialize modern ADO.NET connection
        ConnectSQLNET()
    End Sub
    
    Public Sub ConnectSQLNET()
        Try
            With sqlCNN
                If .State = ConnectionState.Open Then .Close()
                .ConnectionString = "Server=" & MDServerName & ";Database=" & MDDatabaName & ";User Id=" & MDServerUser & ";Password=" & MDServerPassword & ";"
                .Open()
            End With
        Catch ex As Exception
            VSysError = True
            FmLogin.Visible = False
            MessageBox.Show("Modern SQL Connection Error: " & ex.Message, "Connection Error")
            Conection_To_Servee.ShowDialog()
        End Try
    End Sub

    Public Sub LoadSqlData(ByVal StrSql As String, ByVal Rs As ADODB.Recordset)
        With Rs
            If .State = ConnectionState.Open Then .Close()
            .ActiveConnection = CNN
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
            .LockType = ADODB.LockTypeEnum.adLockReadOnly
  
            .Open(StrSql)
            .Requery()
        End With
    End Sub
    
    ' Modern ADO.NET helper methods
    Public Function GetDataTable(ByVal sql As String) As DataTable
        Try
            Using command As New SqlCommand(sql, sqlCNN)
                Using adapter As New SqlDataAdapter(command)
                    Dim dt As New DataTable()
                    adapter.Fill(dt)
                    Return dt
                End Using
            End Using
        Catch ex As Exception
            VSysError = True
            MessageBox.Show("Database Error: " & ex.Message, "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return New DataTable()
        End Try
    End Function
    
    Public Function ExecuteNonQuery(ByVal sql As String) As Integer
        Try
            Using command As New SqlCommand(sql, sqlCNN)
                Return command.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            VSysError = True
            MessageBox.Show("Database Error: " & ex.Message, "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return -1
        End Try
    End Function

    Public CNNMY As New MySqlConnection
    Public Sub ConnectMYSQL()
        With CNNMY
            '            On Error GoTo hang
            '            If .State = ConnectionState.Open Then .Close()
            'hang:
            '            If Err.Number = 0 Then
            '                VSysError = False
            '                ' .ConnectionString = "Server=apis.com.la;Database=apb_msp;Uid=admin;Pwd=Sql_admin@#2024;Port=3306;"
            '                .ConnectionString = "server=apis.com.la;user id=admin;password=Sql_admin@#2024;database=apb_msp"
            '                '.CommandTimeout = 32767
            '                '.ConnectionTimeout = 32767
            '                .Open()
            '            Else
            '                VSysError = True
            '                FmLogin.Visible = False
            '                MessageBox.Show("ຕິດຕໍ່ຖານຂໍ້ມູນບໍ່ໄດ້")
            '                Conection_To_Servee.ShowDialog()
            '            End If
            Try
                ' More explicit connection string with port, SSL mode and a short timeout to fail fast when unreachable
                .ConnectionString = "server=apis.com.la;port=3306;user id=admin;password=Sql_admin@#2024;database=apb_msp;SslMode=None;ConnectionTimeout=10"
                .Open()
            Catch ex As MySqlException
                ' Mark error and present a helpful message to the user with the connector error text
                VSysError = True
                FmLogin.Visible = False
                MessageBox.Show("ຕິດຕໍ່ຖານຂໍ້ມູນບໍ່ໄດ້: " & ex.Message, "MySQL Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Conection_To_Servee.ShowDialog()
            Catch ex As Exception
                VSysError = True
                FmLogin.Visible = False
                MessageBox.Show("Unexpected error while connecting to MySQL: " & ex.Message, "MySQL Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Conection_To_Servee.ShowDialog()
            End Try
        End With
    End Sub
    Public Sub LoadMySqlData(ByVal StrMYSql As String, ByVal Rs As MySqlConnection)
        'With Rs
        '    If .State = ConnectionState.Open Then .Close()
        '    .ActiveConnection = CNNMY
        '    .CursorLocation = ADODB.CursorLocationEnum.adUseClient
        '    .CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
        '    .LockType = ADODB.LockTypeEnum.adLockReadOnly
        '    .Open(StrMYSql)
        '    .Requery()
        'End With

        'Using conn As New MySqlConnection(connStr)
        '    conn.Open()
        '    Dim da As New MySqlDataAdapter("SELECT * FROM products", conn)
        '    Dim dt As New DataTable()
        '    da.Fill(dt)
        '    DataGridView1.DataSource = dt
        'End Using
    End Sub

End Module
