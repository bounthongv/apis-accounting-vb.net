Public Class FmReceipt_List
    Dim sql As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        loadSQL()
        LoadListFG()

    End Sub
    Private Sub SelectS()
        If RdDate.Checked = True Then
            DtmStartDate.Enabled = True
            DtmToDate.Enabled = True
            SearchId.Enabled = False
            SearchName.Enabled = False
            TAc_Bnk_Coode.Enabled = False
        End If
        If RdId.Checked = True Then
            DtmStartDate.Enabled = False
            DtmToDate.Enabled = False
            SearchId.Enabled = True
            SearchName.Enabled = False
            TAc_Bnk_Coode.Enabled = False
        End If
        If RdName.Checked = True Then
            DtmStartDate.Enabled = False
            DtmToDate.Enabled = False
            SearchId.Enabled = False
            SearchName.Enabled = True
            TAc_Bnk_Coode.Enabled = False
        End If
        If Ac_Bnk_Coode.Checked = True Then
            DtmStartDate.Enabled = False
            DtmToDate.Enabled = False
            SearchId.Enabled = False
            SearchName.Enabled = False
            TAc_Bnk_Coode.Enabled = True
        End If

    End Sub
    Public Sub loadSQL()

        If RdId.Checked = True Then
            'If SearchId.Text = "" Then MsgBox("ກະລຸນນາໃສ່ຂໍ້ມູນກ່ອນ") : SearchId.Focus() : Exit Sub
            If Rfull.Checked = True Then
                sql = ""
                sql = " AND Receipt_No = '" & SearchId.Text & "' "
            ElseIf RLeft.Checked = True Then
                sql = ""
                sql = " AND (Receipt_No  Like N'" & SearchName.Text.Trim & "%')"
            ElseIf RRight.Checked = True Then
                sql = ""
                sql = " AND (Receipt_No  Like N'%" & SearchName.Text.Trim & "')"
            ElseIf RPercent.Checked = True Then
                sql = ""
                sql = " AND (Receipt_No  Like N'%" & SearchName.Text.Trim & "%')"
            End If

        ElseIf RdName.Checked = True Then
            'If SearchName.Text = "" Then MsgBox("ກະລຸນນາໃສ່ຂໍ້ມູນກ່ອນ") : SearchName.Focus() : Exit Sub
            If Rfull.Checked = True Then
                sql = ""
                sql = " AND Bnk_Ac_Name = '" & SearchName.Text & "' "
            ElseIf RLeft.Checked = True Then
                sql = ""
                sql = " AND (Bnk_Ac_Name  Like N'" & SearchName.Text.Trim & "%')"
            ElseIf RRight.Checked = True Then
                sql = ""
                sql = " AND (Bnk_Ac_Name  Like N'%" & SearchName.Text.Trim & "')"
            ElseIf RPercent.Checked = True Then
                sql = ""
                sql = " AND (Bnk_Ac_Name  Like N'%" & SearchName.Text.Trim & "%')"
            End If
        ElseIf Ac_Bnk_Coode.Checked = True Then
            'If TAc_Bnk_Coode.Text = "" Then MsgBox("ກະລຸນນາໃສ່ຂໍ້ມູນກ່ອນ") : TAc_Bnk_Coode.Focus() : Exit Sub
            If Rfull.Checked = True Then
                sql = ""
                sql = " AND Bnk_Ac_Code = '" & TAc_Bnk_Coode.Text & "' "
            ElseIf RLeft.Checked = True Then
                sql = ""
                sql = " AND (Bnk_Ac_Code  Like N'" & TAc_Bnk_Coode.Text.Trim & "%')"
            ElseIf RRight.Checked = True Then
                sql = ""
                sql = " AND (Bnk_Ac_Code  Like N'%" & TAc_Bnk_Coode.Text.Trim & "')"
            ElseIf RPercent.Checked = True Then
                sql = ""
                sql = " AND (Bnk_Ac_Code  Like N'%" & TAc_Bnk_Coode.Text.Trim & "%')"
            End If
       
        ElseIf RdDate.Checked = True Then
            sql = ""
            sql = " AND Ap_Receipt.InDate    BETWEEN '" & Format(DtmStartDate.Value, "yyyy-MM-dd") & "' AND '" & Format(DtmToDate.Value, "yyyy-MM-dd") & "' "
      
        End If
        sql = sql & " AND status ='0' "
        sql = sql & " AND Receipt_Type =N'" & ComboBox1.Text & "' "
    End Sub
    Public Sub LoadListFG()
        FG.Rows = 1
        With RSC
            Call LoadSqlData("select *  from Ap_Receipt WHERE Receipt_No<>'' " & sql & " order by Receipt_No", RSC)
            If .RecordCount > 0 Then
                While Not .EOF
                    FG.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("Receipt_No").Value)) & _
                                "" & vbTab & Trim(CStr(.Fields("InDate").Value)) & _
                                "" & vbTab & Trim(CStr(.Fields("Bnk_Ac_Code").Value)) & _
                                "" & vbTab & Trim(CStr(.Fields("Bnk_Ac_Name").Value)) & _
                                 "" & vbTab & Trim(Format(CDbl(.Fields("Amt").Value), "##,##0.00")) & _
                                 "" & vbTab & ((.Fields("Curr").Value)))
                    .MoveNext()
                End While
            Else
                FG.Rows = 2
            End If
        End With
    End Sub

    Private Sub RdId_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RdId.CheckedChanged
        SelectS()
    End Sub

    Private Sub RdName_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RdName.CheckedChanged
        SelectS()
    End Sub

    Private Sub RdDate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RdDate.CheckedChanged
        SelectS()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub FG_DblClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles FG.DblClick
       

      

     
     
        Close()
    End Sub

    Private Sub FG_MouseUpEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseUpEvent) Handles FG.MouseUpEvent
        MDReceipt = FG.get_TextMatrix(FG.Row, 1)
    End Sub

    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelChange
        MDReceipt = FG.get_TextMatrix(FG.Row, 1)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Call MdiCNum()
        FmReceipt.ShowDialog()
        FmReceipt.ComboBox1.Text = ComboBox1.Text
    End Sub

    Private Sub FmReceipt_List_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RdDate.Checked = True
        Rfull.Checked = True
        SelectS()
        loadSQL()
        LoadListFG()
        FG.FormatString = "^ລ/ດ |< ເລກບິນ  |ວັນທີ       ||< ຊື່ບັນຊື             |< ຈຳນວນເງິນ        |< ສະກຸນເງິນ  "
        ComboBox1.Items.Clear()

        Call load_Cmb(" SELECT bookname FROM  books  where Type='RCTY' ", "bookname", ComboBox1)
        ComboBox1.SelectedIndex = 0
    End Sub

    Private Sub Ac_Bnk_Coode_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ac_Bnk_Coode.CheckedChanged
        SelectS()
    End Sub

    Private Sub SearchName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles SearchName.KeyPress
        If e.KeyChar = Chr(13) Then
            loadSQL()
            LoadListFG()
        End If
    End Sub

    Private Sub SearchName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SearchName.TextChanged

    End Sub

    Private Sub SearchId_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles SearchId.KeyPress
        If e.KeyChar = Chr(13) Then
            loadSQL()
            LoadListFG()
        End If
    End Sub

    Private Sub SearchId_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SearchId.TextChanged

    End Sub

    Private Sub DtmStartDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DtmStartDate.KeyPress
        If e.KeyChar = Chr(13) Then
            loadSQL()
            LoadListFG()
        End If
    End Sub

    Private Sub DtmStartDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DtmStartDate.ValueChanged

    End Sub

    Private Sub DtmToDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DtmToDate.KeyPress
        If e.KeyChar = Chr(13) Then
            loadSQL()
            LoadListFG()
        End If
    End Sub

    Private Sub DtmToDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DtmToDate.ValueChanged

    End Sub

    Private Sub TAc_Bnk_Coode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TAc_Bnk_Coode.KeyPress
        If e.KeyChar = Chr(13) Then
            loadSQL()
            LoadListFG()
        End If
    End Sub

    Private Sub TAc_Bnk_Coode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TAc_Bnk_Coode.TextChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        loadSQL()
        LoadListFG()
    End Sub

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Call Close()
    End Sub
End Class