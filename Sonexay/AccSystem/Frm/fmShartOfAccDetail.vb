Public Class fmShartOfAccDetail
    Dim SQL As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Panel1.Visible = True
    End Sub
    Public Sub LoadSQL()

        If RdName.Checked = True Then
            If ChbLang.Checked = True Then
                If Rdlasth.Checked = True Then
                    SQL = " AND left(Name_E, N'" & Len(txtSearchName.Text.Trim) & "')= N'" & txtSearchName.Text.Trim & "' "
                End If
                If similar.Checked = True Then
                    SQL = " AND (Name_E  Like N'%" & txtSearchName.Text.Trim & "%')"
                End If
            Else
                '********************
                If Rdlasth.Checked = True Then
                    SQL = " AND left(Name_L, N'" & Len(txtSearchName.Text.Trim) & "')= N'" & txtSearchName.Text.Trim & "' "
                End If
                If similar.Checked = True Then
                    SQL = " AND (Name_L  Like N'%" & txtSearchName.Text.Trim & "%')"
                End If
            End If
        End If
        If RdId.Checked = True Then
            SQL = " AND (left(AC_CODE, '" & Len(txtSearchId.Text.Trim) & "')= '" & txtSearchId.Text.Trim & "'  or AC_Original Like N'%" & txtSearchId.Text & "%') "
        End If
        SQL = SQL & " AND left(AC_CODE, '" & Len(MDSearchAcccode) & "')= '" & MDSearchAcccode & "' "
        If ChAll.Checked = False Then
            'SQL = SQL & " AND Len(AC_CODE) = 7"
        End If
    End Sub

    Private Sub fmShartOfAccDetail_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        FG.Focus()
        FG.Row = 1
        FG.Col = 1
    End Sub

    'Private Sub fmShartOfAccDetail_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    '    BtnExit.Enabled = True
    '    MDSearchAcccode = ""
    '    txtSearchId.Text = ""
    '    txtSearchName.Text = ""
    'End Sub

    Private Sub fmShartOfAccDetail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'FG.FormatString = "^ລ/ດ|< ເລກລະຫັດ   |< ຊື່ພາສາລາວ                                         |< ຊື່ພາສາອັງກິງ                  "

        'SQL = " AND left(AC_CODE, '" & Len(MDSearchAcccode) & "')= '" & MDSearchAcccode & "' "

        'LoadListFG()
        'Timer1.Enabled = True
        'FG.Row = 1
        'FG.Col = 2
        'If FG.get_TextMatrix(1, 1) = "" Then
        '    FG.Rows = 2
        '    Exit Sub
        'End If










        'MsgBox(MDSearchAcccode)

        HidBTClose()
        FG.FormatString = "^ລ/ດ|< ເລກລະຫັດ              |< ຊື່ພາສາລາວ                           |< ຊື່ພາສາອັງກິງ                        |< AC Original     "
       
        StartLoadDataList()
        Timer1.Enabled = True
        FG.Row = 1
        FG.Col = 2
        If FG.get_TextMatrix(1, 1) = "" Then
            FG.Rows = 2
            Exit Sub
        End If
        BtnExit.Enabled = True
    End Sub
  

    Private Sub RdId_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RdId.CheckedChanged
        If RdId.Checked = True Then
            txtSearchId.Enabled = True
            txtSearchName.Enabled = False
            ChbLang.Enabled = False
            txtSearchId.Focus()
        Else
            txtSearchId.Enabled = False
            txtSearchName.Enabled = True
            ChbLang.Enabled = True
            txtSearchName.Focus()
        End If
    End Sub

    Private Sub ChbLang_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChbLang.CheckedChanged
        If ChbLang.Checked = True Then
            ChbLang.Text = "Lao"
        Else
            ChbLang.Text = "Englisth"
        End If
    End Sub

    Private Sub RdName_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RdName.CheckedChanged
        If RdId.Checked = True Then
            txtSearchId.Enabled = True
            txtSearchName.Enabled = False
            ChbLang.Enabled = False
            txtSearchId.Focus()
        Else
            txtSearchId.Enabled = False
            txtSearchName.Enabled = True
            ChbLang.Enabled = True
            txtSearchName.Focus()
        End If
    End Sub

    Private Sub txtSearchId_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSearchId.KeyPress
        If e.KeyChar = Chr(13) Then

            StartLoadDataList()
            Panel1.Visible = False
        End If
    End Sub

    Private Sub txtSearchName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSearchName.KeyPress
        If e.KeyChar = Chr(13) Then

            StartLoadDataList()
            Panel1.Visible = False
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Panel1.Visible = False
    End Sub

    Private Sub BntNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BntNew.Click
        FrNewAcc.txtAc_code.Enabled = True
        FrNewAcc.ShowDialog()
    End Sub
    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        If txtSty.Text = "NsewJeneralJournal" Or txtSty.Text = "NsewJeneralJournal_DR" Then
            If FmNsewJeneralJournal.FG.get_TextMatrix(FmNsewJeneralJournal.FG.Row, 1) + FmNsewJeneralJournal.FG.get_TextMatrix(FmNsewJeneralJournal.FG.Row, 2) = "" Then
                If FmNsewJeneralJournal.FG.Row <> FmNsewJeneralJournal.FG.Rows - 1 Then
                    FmNsewJeneralJournal.FG.RemoveItem()
                    FmNsewJeneralJournal.SumAmountDr()
                    FmNsewJeneralJournal.Panel1.Visible = False
                    FmNsewJeneralJournal.BtnMove.Visible = False
                    FmNsewJeneralJournal.BtnSearch.Visible = False

                End If
            End If
            Close()
        End If
        If txtSty.Text = "NsewJeneralJournal_Adjust" Or txtSty.Text = "NsewJeneralJournal_Adjust_DR" Then
            If Convert.ToString(FmNsewJeneralJournal_Adjust.FG.CurrentRow.Cells(1).Value) + Convert.ToString(FmNsewJeneralJournal_Adjust.FG.CurrentRow.Cells(2).Value) = "" Then
                If FmNsewJeneralJournal_Adjust.FG.CurrentRow.Index <> FmNsewJeneralJournal_Adjust.FG.Rows.Count - 1 Then
                    FmNsewJeneralJournal_Adjust.FG.Rows.RemoveAt(FmNsewJeneralJournal_Adjust.FG.CurrentRow.Index)
                    FmNsewJeneralJournal_Adjust.SumAmountDr()
                    FmNsewJeneralJournal_Adjust.Panel1.Visible = False
                    FmNsewJeneralJournal_Adjust.BtnMove.Visible = False
                    FmNsewJeneralJournal_Adjust.BtnSearch.Visible = False

                End If
            End If
            Close()
        End If
        Close()
    End Sub
    Private Sub HidBTClose()
        ''If txtSty.Text = "NsewJeneralJournal" Then
        ''    BtnExit.Enabled = False
        ''ElseIf txtSty.Text = "NsewJeneralJournal_DR" Then

        ''    BtnExit.Enabled = False
        ''ElseIf txtSty.Text = "NsewJeneralJournal2" Then
        ''    BtnExit.Enabled = False

        ''ElseIf txtSty.Text = "NsewJeneralJournal_DR2" Then
        ''    BtnExit.Enabled = False
        ''End If
    End Sub

    Private Sub FG_DblClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles FG.DblClick
        If FG.get_TextMatrix(1, 1) = "" Then
            MessageBox.Show("ບໍ່ມີລະຫັດໃຫ້ເລືອກ")
            SQL = ""
            StartLoadDataList()
            FG.Row = 1
            FG.Col = 1
            Exit Sub
        End If

        If txtSty.Text = "NsewJeneralJournal" Then
            FmNsewJeneralJournal.FG.set_TextMatrix(R, L, FG.get_TextMatrix(FG.Row, 1))
            FmNsewJeneralJournal.FG.set_TextMatrix(R, 3, FG.get_TextMatrix(FG.Row, 2))
            FmNsewJeneralJournal.FG.set_TextMatrix(R, 4, FG.get_TextMatrix(FG.Row, 3))
            FmNsewJeneralJournal.LoadDesc()
            FmNsewJeneralJournal.AddAcc()
            'If FmNsewJeneralJournal.FG.get_TextMatrix(FmNsewJeneralJournal.FG.Row, 1) = "" Then

            'End If
            Close()
        End If
        If txtSty.Text = "NsewJeneralJournal_DR" Then
            Close()
            FmNsewJeneralJournal.FG.set_TextMatrix(R, L, FG.get_TextMatrix(FG.Row, 1))
            FmNsewJeneralJournal.FG.set_TextMatrix(R, 3, FG.get_TextMatrix(FG.Row, 2))
            FmNsewJeneralJournal.FG.set_TextMatrix(R, 4, FG.get_TextMatrix(FG.Row, 2))
            FmNsewJeneralJournal.LoadDesc()
            FmNsewJeneralJournal.AddAcc2()

        End If
        '==========adj=====



        If txtSty.Text = "NsewJeneralJournal_Adjust" Then
            FmNsewJeneralJournal_Adjust.FG.Rows(R).Cells(L).Value = FG.get_TextMatrix(FG.Row, 1)
            FmNsewJeneralJournal_Adjust.FG.Rows(R).Cells(3).Value = FG.get_TextMatrix(FG.Row, 2)
            FmNsewJeneralJournal_Adjust.FG.Rows(R).Cells(4).Value = FG.get_TextMatrix(FG.Row, 3)
            FmNsewJeneralJournal_Adjust.LoadDesc()
            FmNsewJeneralJournal_Adjust.AddAcc()
            'If FmNsewJeneralJournal.FG.get_TextMatrix(FmNsewJeneralJournal.FG.Row, 1) = "" Then

            'End If
            Close()
        End If
        If txtSty.Text = "NsewJeneralJournal_Adjust_DR" Then
            Close()
            FmNsewJeneralJournal_Adjust.FG.Rows(R).Cells(L).Value = FG.get_TextMatrix(FG.Row, 1)
            FmNsewJeneralJournal_Adjust.FG.Rows(R).Cells(3).Value = FG.get_TextMatrix(FG.Row, 2)
            FmNsewJeneralJournal_Adjust.FG.Rows(R).Cells(4).Value = FG.get_TextMatrix(FG.Row, 2)
            FmNsewJeneralJournal_Adjust.LoadDesc()
            FmNsewJeneralJournal_Adjust.AddAcc2()

        End If




        'If txtSty.Text = "NsewJeneralJournal" Then
        '    FmNsewJeneralJournal_Adjust.FG.set_TextMatrix(R, L, FG.get_TextMatrix(FG.Row, 1))
        '    FmNsewJeneralJournal_Adjust.FG.set_TextMatrix(R, 3, FG.get_TextMatrix(FG.Row, 2))
        '    FmNsewJeneralJournal_Adjust.FG.set_TextMatrix(R, 4, FG.get_TextMatrix(FG.Row, 3))
        '    FmNsewJeneralJournal_Adjust.LoadDesc()
        '    FmNsewJeneralJournal_Adjust.AddAcc()
        '    'If FmNsewJeneralJournal.FG.get_TextMatrix(FmNsewJeneralJournal.FG.Row, 1) = "" Then

        '    'End If
        '    Close()
        'End If
        'If txtSty.Text = "NsewJeneralJournal_DR" Then
        '    Close()
        '    FmNsewJeneralJournal_Adjust.FG.set_TextMatrix(R, L, FG.get_TextMatrix(FG.Row, 1))
        '    FmNsewJeneralJournal_Adjust.FG.set_TextMatrix(R, 3, FG.get_TextMatrix(FG.Row, 2))
        '    FmNsewJeneralJournal_Adjust.FG.set_TextMatrix(R, 4, FG.get_TextMatrix(FG.Row, 2))
        '    FmNsewJeneralJournal_Adjust.LoadDesc()
        '    FmNsewJeneralJournal_Adjust.AddAcc2()

        'End If







        '*****************For open_jn************
        If txtSty.Text = "NewOpen_jn_dr" Then
            FmNewOpen_jn.txtCode_dr.Text = FG.get_TextMatrix(FG.Row, 1)
            Close()
        End If
        If txtSty.Text = "NewOpen_jn_cr" Then
            FmNewOpen_jn.txtCode_cr.Text = FG.get_TextMatrix(FG.Row, 1)
            Close()
        End If

        If txtSty.Text = "FmBLS" Then
            FmLBS_Item.FG2.set_TextMatrix(FmLBS_Item.FG2.Row, 2, FG.get_TextMatrix(FG.Row, 1))
            FmLBS_Item.FG2.set_TextMatrix(FmLBS_Item.FG2.Row, 3, FG.get_TextMatrix(FG.Row, 2))
            FmLBS_Item.FG2.set_TextMatrix(FmLBS_Item.FG2.Row, 1, FmLBS_Item.TextBox1.Text)
            FmLBS_Item.FG2.set_TextMatrix(FmLBS_Item.FG2.Row, 4, FG.get_TextMatrix(FG.Row, 3))
            FmLBS_Item.FG2.set_TextMatrix(FmLBS_Item.FG2.Row, 0, CDbl(CDbl(FmLBS_Item.FG2.Row) - 1) + 1)
            'MsgBox(FmBankReportId.FG2.Rows)
            If FmLBS_Item.FG2.Row = FmLBS_Item.FG2.Rows - 1 Then
                FmLBS_Item.FG2.Rows = CDbl(FmLBS_Item.FG2.Row) + 2
            End If
            FmLBS_Item.Button2.Enabled = True
            Close()
        End If
        If txtSty.Text = "FmInCome" Then
            FmIncome_Old.FG2.set_TextMatrix(FmIncome_Old.FG2.Row, 2, FG.get_TextMatrix(FG.Row, 1))
            FmIncome_Old.FG2.set_TextMatrix(FmIncome_Old.FG2.Row, 3, FG.get_TextMatrix(FG.Row, 2))
            FmIncome_Old.FG2.set_TextMatrix(FmIncome_Old.FG2.Row, 1, FmIncome_Old.TextBox1.Text)
            FmIncome_Old.FG2.set_TextMatrix(FmIncome_Old.FG2.Row, 4, FG.get_TextMatrix(FG.Row, 3))
            FmIncome_Old.FG2.set_TextMatrix(FmIncome_Old.FG2.Row, 0, CDbl(CDbl(FmIncome_Old.FG2.Row) - 1) + 1)
            'MsgBox(FmBankReportId.FG2.Rows)
            If FmIncome_Old.FG2.Row = FmIncome_Old.FG2.Rows - 1 Then
                FmIncome_Old.FG2.Rows = CDbl(FmIncome_Old.FG2.Row) + 2
            End If
            FmIncome_Old.Button2.Enabled = True
            Close()
        End If
        If txtSty.Text = "FmAmtStatus_Item" Then

        End If
        '*****************For open_jn************
        If txtSty.Text = "Acc_Statement" Then
            Frm_Statement.TxtAccCode.Text = FG.get_TextMatrix(FG.Row, 1)
            Frm_Statement.TxtAccName.Text = FG.get_TextMatrix(FG.Row, 2)
            Close()
        End If
        '*****************FrmRpt_Group DR************
        If txtSty.Text = "FrmRpt_Group_DR" Then
            FrmRpt_Group.txtAcc.Text = FG.get_TextMatrix(FG.Row, 1)
            FrmRpt_Group.TxtDrNm.Text = FG.get_TextMatrix(FG.Row, 2)
            Close()
        End If
        '****************FrmRpt_Group CR***********
        If txtSty.Text = "FrmRpt_Group_CR" Then
            FrmRpt_Group.TxtLH.Text = FG.get_TextMatrix(FG.Row, 1)
            FrmRpt_Group.TxtCrNm.Text = FG.get_TextMatrix(FG.Row, 2)
            Close()
        End If
        '****************FrmRpt_Group CR***********
        If txtSty.Text = "FrmAssetNew" Then
            FrmAssetNew.txtAcc.Text = FG.get_TextMatrix(FG.Row, 1)
            FrmAssetNew.TxtCrNm.Text = FG.get_TextMatrix(FG.Row, 2)
            Close()
        End If

        '****************FrmAdjustment_List_Dr_Curr  ***********
        If txtSty.Text = "FrmAdjustment_List_Dr_Curr" Then
            FrmAdjustment_List.TxtDr_Curr.Text = FG.get_TextMatrix(FG.Row, 1)
            FrmAdjustment_List.TxtDrNm_Curr.Text = FG.get_TextMatrix(FG.Row, 2)
            Close()
        End If
        '****************FrmAdjustment_List_Cr _Curr ***********
        If txtSty.Text = "FrmAdjustment_List_Cr_Curr" Then
            FrmAdjustment_List.TxtCr_Curr.Text = FG.get_TextMatrix(FG.Row, 1)
            FrmAdjustment_List.TxtCrNm_Curr.Text = FG.get_TextMatrix(FG.Row, 2)
            Close()
        End If

        '****************FrmAdjustment_List_Dr ***********
        If txtSty.Text = "FrmAdjustment_List_Dr" Then
            FrmAdjustment_List.TxtDr.Text = FG.get_TextMatrix(FG.Row, 1)
            FrmAdjustment_List.TxtDrNm.Text = FG.get_TextMatrix(FG.Row, 2)
            Close()
        End If
        '****************FrmAdjustment_List_Cr ***********
        If txtSty.Text = "FrmAdjustment_List_Cr" Then
            FrmAdjustment_List.TxtCr.Text = FG.get_TextMatrix(FG.Row, 1)
            FrmAdjustment_List.TxtCrNm.Text = FG.get_TextMatrix(FG.Row, 2)
            Close()
        End If



        If txtSty.Text = "FrmRpt_Fixed_Assets_DR" Then
            'Dim DDDR, DDDRN As String
            DDDR = FG.get_TextMatrix(FG.Row, 1)
            DDDRN = FG.get_TextMatrix(FG.Row, 2)
            FrmRpt_Fixed_Assets.TxtDr.Text = FG.get_TextMatrix(FG.Row, 1)
            FrmRpt_Fixed_Assets.TxtDrNm.Text = FG.get_TextMatrix(FG.Row, 2)
            Me.Close()
        End If

        If txtSty.Text = "FrmRpt_Fixed_Assets_CR" Then
            DDDR = FG.get_TextMatrix(FG.Row, 1)
            DDDRN = FG.get_TextMatrix(FG.Row, 2)
            FrmRpt_Fixed_Assets.TxtCr.Text = FG.get_TextMatrix(FG.Row, 1)
            FrmRpt_Fixed_Assets.TxtCrNm.Text = FG.get_TextMatrix(FG.Row, 2)
            Me.Close()
        End If
        '=======Grp Asset=======
        If txtSty.Text = "FrmGrpNew_LS_DR" Then
            FrmGrpNew_LS.TxtCodeAsDR.Text = FG.get_TextMatrix(FG.Row, 1)
            FrmGrpNew_LS.TxtDrNm.Text = FG.get_TextMatrix(FG.Row, 2)
            Close()
        End If

        If txtSty.Text = "FrmGrpNew_LS_CR" Then
            FrmGrpNew_LS.TxtCodeAsCR.Text = FG.get_TextMatrix(FG.Row, 1)
            FrmGrpNew_LS.TxtCrNm.Text = FG.get_TextMatrix(FG.Row, 2)
            Close()
        End If
        '=======ສະສາງ======= 
        If txtSty.Text = "FrmBrokeNew_DR" Then
            FrmBrokeNew.TxtDr.Text = FG.get_TextMatrix(FG.Row, 1)
            FrmBrokeNew.TxtDrNm.Text = FG.get_TextMatrix(FG.Row, 2)
            Close()
        End If
        If txtSty.Text = "FrmBrokeNew_CR" Then
            FrmBrokeNew.TxtCr.Text = FG.get_TextMatrix(FG.Row, 1)
            FrmBrokeNew.TxtCrNm.Text = FG.get_TextMatrix(FG.Row, 2)
            Close()
        End If
        If txtSty.Text = "FrmBrokeNew_DR22" Then
            FrmBrokeNew.TxtDr22.Text = FG.get_TextMatrix(FG.Row, 1)
            FrmBrokeNew.TxtDrNm22.Text = FG.get_TextMatrix(FG.Row, 2)
            Close()
        End If
        If txtSty.Text = "FrmBrokeNew_DR33" Then
            FrmBrokeNew.TxtDr33.Text = FG.get_TextMatrix(FG.Row, 1)
            FrmBrokeNew.TxtDrNm33.Text = FG.get_TextMatrix(FG.Row, 2)
            Close()
        End If
        If txtSty.Text = "FrmBrokeNew_CR22" Then
            FrmBrokeNew.TxtCr22.Text = FG.get_TextMatrix(FG.Row, 1)
            FrmBrokeNew.TxtCrNm22.Text = FG.get_TextMatrix(FG.Row, 2)
            Close()
        End If
    End Sub

    Private Sub FG_EndAutoSearch(ByVal sender As Object, ByVal e As System.EventArgs) Handles FG.EndAutoSearch

    End Sub

    Private Sub FG_KeyPressEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_KeyPressEvent) Handles FG.KeyPressEvent
        If FG.get_TextMatrix(1, 1) = "" Then
            MessageBox.Show("ບໍ່ມີລະຫັດໃຫ້ເລືອກ")
            SQL = ""
            MDSearchAcccode = ""
            StartLoadDataList()
            FG.Row = 1
            FG.Col = 1
            Exit Sub
        End If
        If txtSty.Text = "NsewJeneralJournal" Then
            FmNsewJeneralJournal.FG.set_TextMatrix(R, L, FG.get_TextMatrix(FG.Row, 1))
            FmNsewJeneralJournal.FG.set_TextMatrix(R, 3, FG.get_TextMatrix(FG.Row, 2))
            FmNsewJeneralJournal.FG.set_TextMatrix(R, 4, FG.get_TextMatrix(FG.Row, 3))
            FmNsewJeneralJournal.AddAcc()
            Close()
        End If
        If txtSty.Text = "NsewJeneralJournal_DR" Then
            Close()
            FmNsewJeneralJournal.FG.set_TextMatrix(R, L, FG.get_TextMatrix(FG.Row, 1))
            FmNsewJeneralJournal.FG.set_TextMatrix(R, 3, FG.get_TextMatrix(FG.Row, 2))
            FmNsewJeneralJournal.FG.set_TextMatrix(R, 4, FG.get_TextMatrix(FG.Row, 3))
            FmNsewJeneralJournal.AddAcc2()

        End If
        '=====================Adj====
        If txtSty.Text = "NsewJeneralJournal" Then
            FmNsewJeneralJournal_Adjust.FG.Rows(R).Cells(L).Value = FG.get_TextMatrix(FG.Row, 1)
            FmNsewJeneralJournal_Adjust.FG.Rows(R).Cells(3).Value = FG.get_TextMatrix(FG.Row, 2)
            FmNsewJeneralJournal_Adjust.FG.Rows(R).Cells(4).Value = FG.get_TextMatrix(FG.Row, 3)
            FmNsewJeneralJournal_Adjust.AddAcc()
            Close()
        End If
        If txtSty.Text = "NsewJeneralJournal_DR" Then
            Close()
            FmNsewJeneralJournal_Adjust.FG.Rows(R).Cells(L).Value = FG.get_TextMatrix(FG.Row, 1)
            FmNsewJeneralJournal_Adjust.FG.Rows(R).Cells(3).Value = FG.get_TextMatrix(FG.Row, 2)
            FmNsewJeneralJournal_Adjust.FG.Rows(R).Cells(4).Value = FG.get_TextMatrix(FG.Row, 3)
            FmNsewJeneralJournal_Adjust.AddAcc2()

        End If



    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        SQL = ""
        StartLoadDataList()
    End Sub

    Private Sub FG_MouseUpEvent(ByVal sender As Object, ByVal e As AxVSFlex8U._IVSFlexGridEvents_MouseUpEvent) Handles FG.MouseUpEvent
        Acc_Code = FG.get_TextMatrix(FG.Row, 1)
    End Sub

    Private Sub FG_SelChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelChange

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        StartLoadDataList()
    End Sub

    Private Sub Button2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.DoubleClick

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Static cntd As Integer = 0
        Dim a As String
        If cntd > 0 Then
            a = cntd.ToString
        Else
            Timer1.Enabled = False
            FG.Focus()
        End If
    End Sub

    Private Sub txtSearchId_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchId.TextChanged

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
      call  StartLoadDataList()
    End Sub
    Public Sub StartLoadDataList()

        Call LoadDividePage()
        P = 1

        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page
        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        If CmbPage.SelectedIndex >= 0 Then
            CmbPage.SelectedIndex = 0
        End If
        If CmbPage.Items.Count > 0 Then
            CmbPage.SelectedIndex = 0
        End If
    End Sub
    Public Sub PageCnt(ByVal StrSQL As String, ByVal ConStr As String, ByVal PageNum As Long, ByVal RowPerPage As Integer)
        'Me.Enabled = False
        Dim RsLoad As New ADODB.Recordset
        Dim rssum As New ADODB.Recordset
        Dim i As Integer
        FG.Rows = 1
        Dim x As String
        x = " AC_CODE,AC_Original , Name_L , Name_E , code_dr "

        PageNum = PageNum - 1
        Dim PK As String = "select " & x & "  from Acc_Code WHERE cnt<>''  " & SQL & " and Ac_code like '%.%' order by Ac_Code ASC "
        LoadSqlData(PK, RSC)
        With RSC
            If .RecordCount <> 0 Then
                .MoveFirst()
                .Move(RowPerPage * PageNum)
                LbPage.Text = Int(.RecordCount)
                If Int(.RecordCount Mod RowPerPage) = 0 Then
                    Last_page = Int(.RecordCount / DividePage)
                Else
                    Last_page = Int(.RecordCount / DividePage) + 1
                    If P = Last_page Then RowPerPage = (.RecordCount Mod RowPerPage)
                End If
                FG.Redraw = False
                FG.Rows = 1

                For i = 0 To RowPerPage - 1


                    FG.AddItem(.AbsolutePosition & vbTab & Trim(CStr(.Fields("AC_CODE").Value)) & _
                                "" & vbTab & Trim(CStr(.Fields("Name_L").Value.ToString)) & _
                                     "" & vbTab & Trim(CStr(.Fields("Name_E").Value.ToString)) & _
                                            "" & vbTab & ((.Fields("AC_Original").Value.ToString)))
                    .MoveNext()
                Next i
                FG.Row = FG.Rows - 1
                FG.Redraw = True
                lblpage_total.Text = P & "/" & Int(Last_page)
            Else
                FG.Rows = 1
                FG.Rows = 2
                lblpage_total.Text = "0/0"
            End If
        End With

        FirstPage.Enabled = True
        BackPage.Enabled = True
        NextPage.Enabled = True
        LasthPage.Enabled = True
        If P = 1 Then
            FirstPage.Enabled = False
            BackPage.Enabled = False
            NextPage.Enabled = True
            LasthPage.Enabled = True
        ElseIf P = Last_page Then
            FirstPage.Enabled = True
            BackPage.Enabled = True
            NextPage.Enabled = False
            LasthPage.Enabled = False
        End If
        If FG.get_TextMatrix(1, 1) <> "" Then
            LbPage.Text = "Total = (" & LbPage.Text & "), " & " From (" & FG.get_TextMatrix(1, 0) & ") To (" & FG.get_TextMatrix(FG.Rows - 1, 0) & ")"
        Else
            LbPage.Text = "RecordTotal"
        End If

    End Sub
    Private Sub LoadDividePage()

        txtSC15.Enabled = False
        P15.ForeColor = Color.Black
        p25.ForeColor = Color.Black
        p50.ForeColor = Color.Black
        p250.ForeColor = Color.Black
        p500.ForeColor = Color.Black
        p1000.ForeColor = Color.Black
        p100.ForeColor = Color.Black
        If P15.Checked = True Then
            DividePage = txtSC15.Text
            txtSC15.Enabled = True
            P15.ForeColor = Color.Red : txtSC15.Focus() : txtSC15.SelectAll()
            FG.FormatString = "^ລ/ດ|< ເລກລະຫັດ   |< ຊື່ພາສາລາວ                                              |< ຊື່ພາສາອັງກິງ                "
        ElseIf p25.Checked = True Then
            DividePage = 25
            p25.ForeColor = Color.Red
        ElseIf p50.Checked = True Then
            DividePage = 50
            p50.ForeColor = Color.Red
        ElseIf p100.Checked = True Then
            DividePage = 100
            p100.ForeColor = Color.Red
        ElseIf p250.Checked = True Then
            DividePage = 250
            p250.ForeColor = Color.Red
        ElseIf p500.Checked = True Then
            DividePage = 500
            p500.ForeColor = Color.Red
        ElseIf p1000.Checked = True Then
            DividePage = 1000
            p1000.ForeColor = Color.Red
        End If

    End Sub

    Private Sub CmbPage_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbPage.SelectedIndexChanged
        If CmbPage.Text <> "" Then
            P = CDbl(CmbPage.Text)
            Call LoadDividePage()
            Call LoadSQL()
            Call PageCnt(StrSQL, ConString, P, DividePage)
            Me.lblpage_total.Text = P & "/" & Last_page
        End If
    End Sub

    Private Sub FirstPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FirstPage.Click
        Call LoadDividePage()
        P = 1
        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page

        If CmbPage.SelectedIndex >= 0 Then
            CmbPage.SelectedIndex = P - 1
        End If
        If CmbPage.Items.Count > 0 Then
            CmbPage.SelectedIndex = 0
        End If
    End Sub

    Private Sub BackPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackPage.Click
        Call LoadDividePage()
        If P = 1 Then Exit Sub
        P = P - 1

        Call LoadSQL()

        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = P & "/" & Last_page
        If CmbPage.SelectedIndex >= 0 Then
            CmbPage.SelectedIndex = P - 1
        End If

    End Sub

    Private Sub NextPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NextPage.Click
        Call LoadDividePage()
        If P >= Last_page Then Exit Sub
        P = P + 1
        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = P & "/" & Last_page
        If CmbPage.SelectedIndex >= 0 Then
            CmbPage.SelectedIndex = P - 1
        End If
    End Sub

    Private Sub LasthPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LasthPage.Click
        Call LoadDividePage()
        P = Last_page
        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = Last_page & "/" & Last_page
        If CmbPage.SelectedIndex >= 0 Then
            CmbPage.SelectedIndex = P - 1
        End If
    End Sub

    Private Sub P15_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles P15.CheckedChanged

    End Sub

    Private Sub SelectPageSty()
        Call LoadDividePage()
        P = 1
        Call LoadSQL()
        Call PageCnt(StrSQL, ConString, P, DividePage)
        Me.lblpage_total.Text = "1/" & Last_page
        CmbPage.Items.Clear()
        For i = 0 To Last_page - 1
            CmbPage.Items.Add(i + 1)
        Next i
        If CmbPage.Items.Count > 0 Then
            CmbPage.SelectedIndex = 0
        End If
    End Sub


    Private Sub p25_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles p25.CheckedChanged

    End Sub

    Private Sub p25_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles p25.KeyPress

    End Sub

    Private Sub p25_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p25.MouseClick
        Call SelectPageSty()
    End Sub

    Private Sub p50_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles p50.CheckedChanged

    End Sub

    Private Sub p50_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p50.MouseClick
        Call SelectPageSty()
    End Sub

    Private Sub p100_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles p100.CheckedChanged

    End Sub

    Private Sub p100_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p100.MouseClick
        Call SelectPageSty()
    End Sub

    Private Sub p250_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles p250.CheckedChanged

    End Sub

    Private Sub p500_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles p500.CheckedChanged

    End Sub

    Private Sub p1000_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles p1000.CheckedChanged

    End Sub

    Private Sub p1000_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p1000.MouseClick
        Call SelectPageSty()
    End Sub

    Private Sub P15_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles P15.MouseClick
        Call SelectPageSty()
    End Sub

    Private Sub p250_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p250.MouseClick
        Call SelectPageSty()
    End Sub

    Private Sub p500_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p500.MouseClick
        Call SelectPageSty()
    End Sub

    Private Sub txtSC15_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSC15.KeyPress
        If e.KeyChar = Chr(13) Then
            Call StartLoadDataList()
        End If
    End Sub

    Private Sub txtSC15_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSC15.TextChanged

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click

    End Sub
End Class