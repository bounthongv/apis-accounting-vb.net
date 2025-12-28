Imports System.Data
Imports System.Data.SqlClient
Module ConnectClientServer
    Public cn As New SqlClient.SqlConnection
    Public cm As New SqlCommand
    Public da As SqlDataAdapter
    Public ds As New DataSet
    Public SqlClient As String = ""
    Public MuAcCode As String = ""
    '============================
    Public SrvrNme, DTNme, MUTY As String
    Public MDBSL As String
    Public daa, dd As SqlDataAdapter
    Public SQL_DG As String
 
    Public Sub LoadCN_DG()
        ConnectCL()
        da = New SqlDataAdapter(sql, cn)

        dd = New SqlDataAdapter(SQL_DG, cn)
        daa = New SqlDataAdapter(SQL_DG, cn)
    End Sub
    Public Sub LoadCN()
        ConnectCL()
        da = New SqlDataAdapter(sql, cn)
        da = New SqlDataAdapter(SqlClient, cn)

        dd = New SqlDataAdapter(SQL_DG, cn)
        daa = New SqlDataAdapter(SQL_DG, cn)
    End Sub


    Public Sub CnnEdit_DG()
        ConnectCL()
        Try
            With cm
                .CommandType = CommandType.Text
                .CommandText = sql
                .Connection = cn
                .ExecuteNonQuery()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub CnnEdit()
        ConnectCL()
        Try
            With cm
                .CommandType = CommandType.Text
                .CommandText = SqlClient
                .Connection = cn
                .ExecuteNonQuery()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Public Sub ConnectCL()
        Try
            With cn
                If .State = ConnectionState.Open Then .Close()
                .ConnectionString = "Data Source= " & MDServerName & " ;Initial Catalog= " & MDDatabaName & " ;User ID = " & MDServerUser & " ;Password= " & MDServerPassword & " "
                .Open()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Module
