Option Explicit On
Option Strict On
Module DBConnection
    Public conn As New ADODB.Connection
    Public strConn, RunCnn As String
    Public Sub ConnectAccess()
        'Dim strConn As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =Connection.mdb;Persist Security Info=False;Jet OLEDB:Database Password=2459428"
        'strConn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = Connection.mdb;Persist Security Info=False;Jet OLEDB:Database Password=2459428"
        Dim strConn As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =Connection.mdb;Persist Security Info=True;Jet OLEDB:Database Password=2459428"

        Try
            If conn.State = ConnectionState.Open Then conn.Close()

            conn.Open(strConn)
        Catch ex As Exception
            MessageBox.Show("ບໍ່ສາມາດເຊື່ອມຕໍ່ກັບຖານຂໍ້ມູນໄດ້: ")
            FmLogin.Close()
        End Try
    End Sub
    Public Sub ConnectPartition()
        Dim strConn As String = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = PTS.mdb;Persist Security Info=False;Jet OLEDB:Database Password=2459428"

        Try
            If conn.State = ConnectionState.Open Then conn.Close()
            conn.Open(strConn)
        Catch ex As Exception
            MessageBox.Show("ບໍ່ສາມາດເຊື່ອມຕໍ່ກັບຖານຂໍ້ມູນໄດ້: ")
            FmLogin.Close()
        End Try
    End Sub
    Public Sub LoadAcData(ByVal sql As String, ByVal rs As ADODB.Recordset)
        With rs
            If .State = ConnectionState.Open Then .Close()
            .ActiveConnection = conn
            .Open(sql)
            .Requery()
        End With
    End Sub
End Module
