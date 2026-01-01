Public Class FmOff_User

    Private Sub FmOff_User_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Setup FG Columns
        FG.Columns.Clear()
        FG.Columns.Add("No", "ລ/ດ")
        FG.Columns.Add("BranchID", "ລະຫັດສາຂາ")
        FG.Columns.Add("BranchName", "ຊື່ສາຂາ")

        FG.Columns(0).Width = 50
        FG.Columns(1).Width = 100
        FG.Columns(2).Width = 150

        FG.AllowUserToAddRows = False
        FG.ReadOnly = True
        FG.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG.MultiSelect = False
        
        ' Setup FG2 Columns (Assuming similar structure or empty for now as it was not used in legacy code snippet)
        FG2.Columns.Clear()
        FG2.Columns.Add("No", "ລ/ດ")
        FG2.Columns.Add("Col1", "Column 1")
        FG2.Columns.Add("Col2", "Column 2")
        
        FG2.AllowUserToAddRows = False
        FG2.ReadOnly = True
        FG2.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        FG2.MultiSelect = False

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
        If CmbUsr.Items.Count > 0 Then
            CmbUsr.SelectedIndex = 0
        End If
    End Sub

    Public Sub LoadListFG()
        FG.Rows.Clear()
        With RSC
            Call LoadSqlData("select Sub_Id , Off_Add2  from Ap_office order by cnt", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.Rows.Add(.AbsolutePosition, _
                                Trim(CStr(.Fields("Sub_Id").Value)), _
                                Trim(CStr(.Fields("Off_Add2").Value)))
                    .MoveNext()

                End While
            End If
        End With
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

    End Sub
End Class