Imports System.Data.SqlClient
Imports System.Data

Module DbHelper
    ' Connection String Helper
    Private Function GetConnectionString() As String
        ' Reusing global variables from MDDeclareLation/MDSQLConnection
        Dim connStr As String = "Data Source=" & MDServerName & ";Initial Catalog=" & MDDatabaName & ";User ID=" & MDServerUser & ";Password=" & MDServerPassword & ";Persist Security Info=True;TrustServerCertificate=True;"
        Return connStr
    End Function

    ' Execute NonQuery (Insert, Update, Delete)
    Public Function ExecuteNonQuery(ByVal sql As String) As Integer
        Using conn As New SqlConnection(GetConnectionString())
            Using cmd As New SqlCommand(sql, conn)
                Try
                    conn.Open()
                    Return cmd.ExecuteNonQuery()
                Catch ex As Exception
                    Throw New Exception("Error executing SQL: " & ex.Message & vbCrLf & "SQL: " & sql, ex)
                End Try
            End Using
        End Using
    End Function

    ' Execute Scalar (Get single value)
    Public Function ExecuteScalar(ByVal sql As String) As Object
        Using conn As New SqlConnection(GetConnectionString())
            Using cmd As New SqlCommand(sql, conn)
                Try
                    conn.Open()
                    Return cmd.ExecuteScalar()
                Catch ex As Exception
                    Throw New Exception("Error executing Scalar: " & ex.Message & vbCrLf & "SQL: " & sql, ex)
                End Try
            End Using
        End Using
    End Function

    ' Get DataTable (Select)
    Public Function GetDataTable(ByVal sql As String) As DataTable
        Using conn As New SqlConnection(GetConnectionString())
            Using cmd As New SqlCommand(sql, conn)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    Try
                        conn.Open()
                        da.Fill(dt)
                        Return dt
                    Catch ex As Exception
                        Throw New Exception("Error fetching data: " & ex.Message & vbCrLf & "SQL: " & sql, ex)
                    End Try
                End Using
            End Using
        End Using
    End Function

    ' Helper to get safe string from DB value
    Public Function GetStr(ByVal val As Object) As String
        If IsDBNull(val) OrElse val Is Nothing Then
            Return ""
        Else
            Return val.ToString()
        End If
    End Function
End Module
