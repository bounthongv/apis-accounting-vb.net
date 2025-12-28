Public Class FmOff_User

    Private Sub FmOff_User_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FG.FormatString = "^ລ/ດ |< ລະຫັດສາຂາ     |<ຊື່ສາຂາ             "
        LoadListFG()
        loadUser()
    End Sub

    Private Sub loadUser()
        CmbUsr.Items.Clear()
        LoadSqlData("select Usr_id from  AP_Users Order by cnt", RSC)
        With RSC
            Do Until .EOF = True
                CmbUsr.Items.Add((.Fields("Usr_id").Value))
                .MoveNext()
            Loop
        End With
        CmbUsr.SelectedIndex = 0
    End Sub

    Public Sub LoadListFG()
        FG.Rows = 1
        With RSC
            Call LoadSqlData("select Sub_Id , Off_Add2  from Ap_office order by cnt", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("Sub_Id").Value)) & _
                    "" & vbTab & Trim(CStr(.Fields("Off_Add2").Value)))
                    .MoveNext()

                End While
            Else
                FG.Rows = 2
            End If
        End With
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

    End Sub
End Class