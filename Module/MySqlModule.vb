Imports MySql.Data.MySqlClient

Module MySqlModule
    Public Function FillData(ByVal sql As String) As DataSet
        Dim da As MySqlDataAdapter
        Dim ds As New DataSet
        Dim connStr As String = "Server=apis.com.la;Database=apb_msp;Uid=admin;Pwd=Sql_admin@#2024;Port=3306;SslMode=None;"
        Dim conn As New MySqlConnection(connStr)
        ds = New DataSet
        da = New MySqlDataAdapter(sql, conn)
        da.Fill(ds)
        FillData = ds
    End Function
End Module
