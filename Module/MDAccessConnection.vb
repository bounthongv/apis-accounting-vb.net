Option Explicit On
Option Strict On

Imports System.Data.OleDb
Module DBConnection
    Private conn As New OleDbConnection
    Public strConn, RunCnn As String
    Public Sub ConnectAccess()
        'Dim strConn As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = Connection.mdb;Persist Security Info=False;Jet OLEDB:Database Password=2459428"
        'strConn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = Connection.mdb;Persist Security Info=True;Jet OLEDB:Database Password=2459428"
        strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = Connection.mdb;Persist Security Info=False;Jet OLEDB:Database Password=2459428"

        Try
            If conn.State = ConnectionState.Open Then conn.Close()
            conn.ConnectionString = strConn
            conn.Open()
        Catch ex As Exception
            MessageBox.Show("ບໍ່ສາມາດບັນທຶກຄດ້")
            FmLogin.Close()
        End Try
    End Sub
    Public Sub ConnectPartition()
        Dim strConn As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = PTS.mdb;Persist Security Info=False;Jet OLEDB:Database Password=2459428"

        Try
            If conn.State = ConnectionState.Open Then conn.Close()
            conn.ConnectionString = strConn
            conn.Open()
        Catch ex As Exception
            MessageBox.Show("ບໍ່ສາມາດບັນທຶກຄດ້")
            FmLogin.Close()
        End Try
    End Sub
    
    ' Modern ADO.NET helper methods
    Public Function GetAccessDataTable(ByVal sql As String) As DataTable
        Try
            Using command As New OleDbCommand(sql, conn)
                Using adapter As New OleDbDataAdapter(command)
                    Dim dt As New DataTable()
                    adapter.Fill(dt)
                    Return dt
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Database Error: " & ex.Message, "Access Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return New DataTable()
        End Try
    End Function
    
    Public Function ExecuteAccessNonQuery(ByVal sql As String) As Integer
        Try
            Using command As New OleDbCommand(sql, conn)
                Return command.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            MessageBox.Show("Database Error: " & ex.Message, "Access Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return -1
        End Try
    End Function
    
    ' Legacy compatibility method - redirects to modern implementation
    Public Sub LoadAcData(ByVal sql As String, ByVal rs As Object)
        ' Legacy method - no longer used, keeping for backward compatibility
        ' Implementation replaced with GetAccessDataTable(sql)
        Dim dt As DataTable = GetAccessDataTable(sql)
        ' Note: This method should be removed from calling code
    End Sub
End Module
