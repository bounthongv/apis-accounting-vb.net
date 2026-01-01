Public Class Frm_F08Edit

    Private Sub Frm_F08Edit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetupGrid()
        Call AAA()
    End Sub

    Private Sub SetupGrid()
        FG.Columns.Clear()
        FG.Columns.Add("No", "ລ/ດ")
        FG.Columns.Add("ItemID", "ລະຫັດ")
        FG.Columns.Add("Itemnm", "ປະເພດສິນເຊື່ອ")
        FG.Columns.Add("Vala", "ໄລຍະເວລາ")
        FG.Columns.Add("Int", "ອັດຕາ%")
        FG.Columns.Add("SaveNm", "ປະເພດເງິນຝາກ")
        FG.Columns.Add("Save_Vala", "ໄລຍະເວລາ")
        FG.Columns.Add("Save_Int", "ອັດຕາ%")

        FG.Columns(0).Width = 50
        FG.Columns(1).Width = 80
        FG.Columns(2).Width = 200
        FG.Columns(3).Width = 120
        FG.Columns(4).Width = 100
        FG.Columns(5).Width = 200
        FG.Columns(6).Width = 120
        FG.Columns(7).Width = 100

        FG.Columns(0).ReadOnly = True
        FG.Columns(1).ReadOnly = False
        FG.Columns(2).ReadOnly = False
        FG.Columns(3).ReadOnly = False
        FG.Columns(4).ReadOnly = False
        FG.Columns(5).ReadOnly = False
        FG.Columns(6).ReadOnly = False
        FG.Columns(7).ReadOnly = False
        
        FG.AllowUserToAddRows = False
        FG.SelectionMode = DataGridViewSelectionMode.CellSelect
    End Sub

    Private Sub AAA()
        FG.Rows.Clear()
        With RSC
            Call LoadSqlData("SELECT * FROM RPT_F08 Order by ItemID asc ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.Rows.Add(.AbsolutePosition, _
                                Trim(CStr(.Fields("ItemID").Value)), _
                                Trim(CStr(.Fields("Itemnm").Value)), _
                                Trim(CStr(.Fields("Vala").Value)), _
                                ((.Fields("int").Value)), _
                                ((.Fields("SaveNm").Value)), _
                                ((.Fields("Save_Vala").Value)), _
                                ((.Fields("Save_Int").Value)))
                    .MoveNext()
                End While
            End If
        End With
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        FG.Rows.Add()
        For i = 0 To FG.Rows.Count - 1
            If FG.Rows(i).Cells(4).Value Is Nothing OrElse FG.Rows(i).Cells(4).Value.ToString() = "" Then
                FG.Rows(i).Cells(4).Value = Format(CDbl(0), "#,##0.00")
            End If
        Next i
    End Sub

    Private Sub FG_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles FG.CellDoubleClick
        If e.RowIndex < 0 Then Exit Sub
        
        Dim itemId As String = If(FG.Rows(e.RowIndex).Cells(1).Value Is Nothing, "", FG.Rows(e.RowIndex).Cells(1).Value.ToString())
        Dim itemNm As String = If(FG.Rows(e.RowIndex).Cells(2).Value Is Nothing, "", FG.Rows(e.RowIndex).Cells(2).Value.ToString())
        
        If MessageBox.Show("ທ່ານຕ້ອງການລືບລາຍການ'" & itemId & itemNm & "' ນີ້ ແທ້ ຫຼື ບໍ່ ?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            CNN.Execute(" DELETE RPT_F08 where  ItemID=N'" & itemId & "' ")
            FG.Rows.RemoveAt(e.RowIndex)
        End If
    End Sub

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        Save_item()
        MsgBox("Finish")
    End Sub

    Private Sub Save_item()
        Dim Rschk As New ADODB.Recordset
        Dim i As Integer
        
        For i = 0 To FG.Rows.Count - 1
            Dim itemId As String = If(FG.Rows(i).Cells(1).Value Is Nothing, "", FG.Rows(i).Cells(1).Value.ToString())
            If itemId = "" Then Continue For
            
            Dim sk As String = "Select * FROM RPT_F08 where  ItemID=N'" & itemId & "'   "
            Call LoadSqlData(sk, Rschk)
            
            Dim v1 As String = If(FG.Rows(i).Cells(1).Value Is Nothing, "", FG.Rows(i).Cells(1).Value.ToString())
            Dim v2 As String = If(FG.Rows(i).Cells(2).Value Is Nothing, "", FG.Rows(i).Cells(2).Value.ToString())
            Dim v3 As String = If(FG.Rows(i).Cells(3).Value Is Nothing, "", FG.Rows(i).Cells(3).Value.ToString())
            Dim v4 As String = If(FG.Rows(i).Cells(4).Value Is Nothing, "0", FG.Rows(i).Cells(4).Value.ToString())
            Dim v5 As String = If(FG.Rows(i).Cells(5).Value Is Nothing, "", FG.Rows(i).Cells(5).Value.ToString())
            Dim v6 As String = If(FG.Rows(i).Cells(6).Value Is Nothing, "", FG.Rows(i).Cells(6).Value.ToString())
            Dim v7 As String = If(FG.Rows(i).Cells(7).Value Is Nothing, "0", FG.Rows(i).Cells(7).Value.ToString())

            If Rschk.RecordCount = 0 Then
                Dim sa As String = "INSERT INTO RPT_F08 ( ItemID, ItemNM, Vala, Int,SaveNm, Save_Vala, Save_Int ) " & _
                " VALUES (  N'" & v1 & "'," & _
                 " N'" & v2 & "'," & _
                  " N'" & v3 & "'," & _
                    " N'" & v4 & "',N'" & v5 & "',N'" & v6 & "',N'" & v7 & "') "
                CNN.Execute(sa)
            Else
                Dim UPPP As String = " UPDATE RPT_F08 set ItemNM=N'" & v2 & "'," & _
                  " Vala=N'" & v3 & "'," & _
                    " Int=N'" & v4 & "',SaveNm=N'" & v5 & "',Save_Vala=N'" & v6 & "',Save_Int=N'" & v7 & "' " & _
                    " where ItemID=N'" & v1 & "' "
                CNN.Execute(UPPP)
            End If
        Next i
        
        CNN.Execute("DELETE FROM RPT_F08 where  ItemID=''  ") 
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class