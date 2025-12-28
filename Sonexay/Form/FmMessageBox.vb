Public Class FmMessageBox
    Dim s As Double = 0
    Private Sub FmMessageBox_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Visible = True
        s = 0
        Timer1.Enabled = True
    End Sub
    Public Sub LoadServer()
        Dim Conn As New ADODB.Connection
        Dim rsProj As New ADODB.Recordset
        Call LoadAcData("Select * from Conect ", rsProj)
        With rsProj
            If .RecordCount <> 0 Then
                MDServerName = (.Fields("ServerName").Value.ToString)
                MDDatabaName = (.Fields("DatabaseName").Value.ToString)
                MDServerUser = (.Fields("UserName").Value.ToString)
                MDServerPassword = (.Fields("UserPassword").Value.ToString)
                MDSeriaAccess = (.Fields("PartitionSeria").Value.ToString)
            End If
        End With
        'MsgBox(MDDatabaName)
    End Sub
    Private Sub LoadRunCnn()
        Dim Conn As New ADODB.Connection
        Dim rsProj As New ADODB.Recordset
        Call LoadAcData("Select RunCnn from Conect ", rsProj)
        With rsProj
            If .RecordCount <> 0 Then
                RunCnn = (.Fields("RunCnn").Value.ToString)
            End If
        End With
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'Me.Visible = False
        s = s + 1
        If s > 1000 Then
            Timer1.Enabled = False
            Call ConnectAccess()
            Call LoadServer()
            Call ConnectSQL()
        End If



        Static cntd As Integer = 1
        cntd = cntd + 1
        LoadRunCnn()
        If RunCnn = 0 Then

            If cntd > 15 Then
                Me.Visible = True
                If TextBox1.Text = "ກຳລັງເຊື່ອມຕໍ່ກັບຖານຂໍ້ມູນກະລຸນນາລໍຖ້າ......" Then
                    Timer1.Enabled = False
                    Call ConnectAccess()
                    Call LoadServer()
                    Call ConnectSQL()
                End If
            End If
        Else
            If cntd > 10 Then
                cntd = 1
            End If
        End If


     
    End Sub
End Class