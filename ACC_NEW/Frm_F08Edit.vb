Public Class Frm_F08Edit

    Private Sub Frm_F08Edit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FG.FormatString = "^ລ/ດ |<ລະຫັດ |< ປະເພດສິນເຊື່ອ                 |< ໄລຍະເວລາ         |< ອັດຕາ%           |< ປະເພດເງິນຝາກ                  |< ໄລຍະເວລາ       |< ອັດຕາ%        "
        Call AAA()
    End Sub
    Private Sub AAA()
        FG.Rows = 1
        With RSC

            Call LoadSqlData("SELECT * FROM RPT_F08 Order by ItemID asc ", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("ItemID").Value)) & vbTab & Trim(CStr(.Fields("Itemnm").Value)) & vbTab & Trim(CStr(.Fields("Vala").Value)) & _
                      "" & vbTab & ((.Fields("int").Value)) & vbTab & ((.Fields("SaveNm").Value)) & vbTab & ((.Fields("Save_Vala").Value)) & vbTab & ((.Fields("Save_Int").Value)))
                    .MoveNext()
                End While
            Else
                FG.Rows = 2
            End If

        End With
 
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        FG.Rows = FG.Rows + 1 
        For i = 1 To FG.Rows - 1
            If FG.get_TextMatrix(i, 4) = "" Then
                FG.set_TextMatrix(i, 4, Format(CDbl(0), "#,##0.00"))
            End If
        Next i
    End Sub

    Private Sub FG_DblClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles FG.DblClick
        'AccCD = FG.get_TextMatrix(FG.Row, 1) & FG.get_TextMatrix(FG.Row, 2)
        If MessageBox.Show("ທ່ານຕ້ອງການລືບລາຍການ'" & FG.get_TextMatrix(FG.Row, 1) & FG.get_TextMatrix(FG.Row, 2) & "' ນີ້ ແທ້ ຫຼື ບໍ່ ?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            CNN.Execute(" DELETE RPT_F08 where  ItemID=N'" & (FG.get_TextMatrix(FG.Row, 1)) & "' ")
            FG.RemoveItem()
        End If
    End Sub

    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelChange
        If FG.Col = 1 Or FG.Col = 2 Or FG.Col = 3 Or FG.Col = 4 Or FG.Col = 5 Or FG.Col = 6 Or FG.Col = 7 Then
            FG.Editable = VSFlex8U.EditableSettings.flexEDKbd
        Else
            FG.Editable = VSFlex8U.EditableSettings.flexEDNone
        End If
    End Sub

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        Save_item()
        MsgBox("Finish")
    End Sub
    Private Sub Save_item()
       Dim Rschk As New ADODB.Recordset
        Dim i As Integer
        With Rschk 
            For i = 1 To FG.Rows - 1
                Dim sk As String = "Select * FROM RPT_F08 where  ItemID=N'" & (FG.get_TextMatrix(i, 1)) & "'   "
                Call LoadSqlData(sk, Rschk)
                If Rschk.RecordCount = 0 Then
                    Dim sa As String = "INSERT INTO RPT_F08 ( ItemID, ItemNM, Vala, Int,SaveNm, Save_Vala, Save_Int ) " & _
                    " VALUES (  N'" & (FG.get_TextMatrix(i, 1)) & "'," & _
                     " N'" & (FG.get_TextMatrix(i, 2)) & "'," & _
                      " N'" & (FG.get_TextMatrix(i, 3)) & "'," & _
                        " N'" & (FG.get_TextMatrix(i, 4)) & "',N'" & (FG.get_TextMatrix(i, 5)) & "',N'" & (FG.get_TextMatrix(i, 6)) & "',N'" & (FG.get_TextMatrix(i, 7)) & "') "
                    CNN.Execute(sa)
                Else
                    Dim UPPP As String = " UPDATE RPT_F08 set ItemNM=N'" & (FG.get_TextMatrix(i, 2)) & "'," & _
                      " Vala=N'" & (FG.get_TextMatrix(i, 3)) & "'," & _
                        " Int=N'" & (FG.get_TextMatrix(i, 4)) & "',SaveNm=N'" & (FG.get_TextMatrix(i, 5)) & "',Save_Vala=N'" & (FG.get_TextMatrix(i, 6)) & "',Save_Int=N'" & (FG.get_TextMatrix(i, 7)) & "' " & _
                        " where ItemID=N'" & (FG.get_TextMatrix(i, 1)) & "' "
                    CNN.Execute(UPPP)
                End If
            Next i
        End With
        CNN.Execute("DELETE FROM RPT_F08 where  ItemID=''  ") 
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class