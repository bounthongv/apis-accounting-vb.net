Public Class FrmData_server
    Dim rsProj As New ADODB.Recordset
    Public editProj As Boolean
    Dim Sql As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'If MDForMain <> "Kantana" Then
        '    FrmLogin.Close()
        'End If
        'If MDForMain = "Kantana" Then
        '    Me.Close()
        'End If
        Me.Close()
    End Sub
    Private Sub FrmData_server_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FmLogin.Hide()
        Dim Conn As New ADODB.Connection
        Call ConnectSQL()
        Call LoadDatabaseServer()

    End Sub
    Private Sub LoadDatabaseServer()
        Call LoadAcData("select * from Conect where svID='" & "001" & "'", rsProj)
        'Call LoadData("SELECT * FROM Conn WHERE ServerID<>'" & Sql & "' order by ServerID ", rsProj)
        With rsProj
            If .RecordCount <> 0 Then
                txtServerID.Text = (.Fields("svID").Value.ToString)
                txtServer.Text = (.Fields("ServerName").Value.ToString)
                txtData.Text = (.Fields("DatabaseName").Value.ToString)
                txtUserNm.Text = (.Fields("UserName").Value.ToString)
                txtPass.Text = (.Fields("UserPassword").Value.ToString)
            End If
        End With
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Dim rs As New ADODB.Recordset
            Call LoadAcData("select * from Conect where svID='" & txtServerID.Text & "'", rs)
            conn.Execute("Update Conect Set ServerName ='" & txtServer.Text.ToString & "', DatabaseName='" & txtData.Text.ToString & "', UserName='" & txtUserNm.Text.ToString & "', UserPassword='" & txtPass.Text.ToString & "' " & _
                         " WHERE svID='" & txtServerID.Text.ToString & "' ")
            'Me.Hide()
            'FrmLogin.Visible = True
            'Test Connection Database server
            With Conn
                If .State = ConnectionState.Open Then .Close()
                .ConnectionString = "Provider = SQLOLEDB.1; Password = " & txtPass.Text & "; Persist Security Info = True; " & _
                "User ID = " & txtUserNm.Text & "; Initial Catalog = " & txtData.Text & "; Data Source =" & txtServer.Text & ""
                .Open()
            End With
            MsgBox("Conect To Server Completed!", MsgBoxStyle.OkOnly)
            If MDForMain <> "Kantana" Then
                FmLogin.Visible = True
                Me.Close()
            End If
            If MDForMain = "Kantana" Then
                Call LoadDatabaseServer()
                'FrmAPCashier.Show()
                Me.Hide()
            End If
        Catch ex As Exception
            MsgBox("ຂໍ້ມູນໃນການຕິດຕໍ່ ຖານຂໍ້ມູນ ບໍ່ຖືກຕ້ອງ !!" & vbCrLf & ex.Message)
        End Try

    End Sub
End Class