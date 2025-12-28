Imports System.IO

Module MuLangauge
    Public sFileName1 As String
    Public MuImgLct, MuImgNme As String
    'Public MuLngRpt As String
    Public MuLngRpt As String
    Public Lng As Long = 9999
    Public MuLng As String
    Public LngL(Lng) As String
    Public MsgSL As Integer = 0
    Public LngE(Lng) As String
    Public LngStr As String = ""
    Public LngId As String
    Dim mItem As ToolStripMenuItem
    Dim mSubItem As ToolStripMenuItem
    Dim ForNme As String
    Dim ForSize As String = 3
    Public Sub LoadRptLng()
        CallLngStr() : MuLngRpt = MuLngRpt & "N'" & LngStr
    End Sub
    Private Sub LoadImageLocation()
        Dim a, k As String
        Dim b As Integer
        a = My.Application.Info.DirectoryPath
        k = Mid(a, CDbl(Len(a)) - 8, 3)
        b = CDbl(Len(a)) - 6
        MuImgLct = Microsoft.VisualBasic.Left(a, b) & "\Image\"
    End Sub
    Private Sub LoadImage()

    End Sub

    Public Sub LoadLngnnn()
        Lng = 0
        LoadAcData("select LngID, LngL , LngE from  Langauge   Order by LngID", RSC)
        With RSC
            Do Until .EOF = True
                LngL(CLng((.Fields("LngID").Value))) = (.Fields("LngL").Value)
                LngE(CLng((.Fields("LngID").Value))) = (.Fields("LngE").Value)
                Lng = Lng + 1
                .MoveNext()
            Loop
        End With
    End Sub
    Public Function ChgeLang(ByVal LangValue As Long) As String
        ChgeLang = IIf(Lang = False, LngL(LangValue), LngE(LangValue))
    End Function
    Public Sub LoadlangMM()
        'Sub Menu........................................
        With FmMain
            '====System====
            .MnSystem.Text = ChgeLang(.MnSystem.Tag)
            .ToolStripMenuItem1.Text = ChgeLang(.ToolStripMenuItem1.Tag)
            .MnOff.Text = ChgeLang(.MnOff.Tag)
            .MnOffSub.Text = ChgeLang(.MnOffSub.Tag)
            .MnDateSeting.Text = ChgeLang(.MnDateSeting.Tag)
            .ຕງຄາສຳປະສດToolStripMenuItem.Text = ChgeLang(.ຕງຄາສຳປະສດToolStripMenuItem.Tag)
            .ToolStripMenuItem56.Text = ChgeLang(.ToolStripMenuItem56.Tag)
            .ToolStripMenuItem58.Text = ChgeLang(.ToolStripMenuItem58.Tag)
            .MnSystemUser.Text = ChgeLang(.MnSystemUser.Tag)
            .MnChangPsw.Text = ChgeLang(.MnChangPsw.Tag)
            .MnChangUser.Text = ChgeLang(.MnChangUser.Tag)
            .MnUser.Text = ChgeLang(.MnUser.Tag)
            .MnImgSize.Text = ChgeLang(.MnImgSize.Tag)
            .ToolStripMenuItem6.Text = ChgeLang(.ToolStripMenuItem6.Tag)
            .ຕງຄາອດຕາແລກປຽນToolStripMenuItem.Text = ChgeLang(.ຕງຄາອດຕາແລກປຽນToolStripMenuItem.Tag)
            .ການປຽນແປງຂອງເງນຕາToolStripMenuItem.Text = ChgeLang(.ການປຽນແປງຂອງເງນຕາToolStripMenuItem.Tag)
            .ToolStripMenuItem10.Text = ChgeLang(.ToolStripMenuItem10.Tag)
            .ToolStripMenuItem11.Text = ChgeLang(.ToolStripMenuItem11.Tag)
            .ToolStripMenuItem23.Text = ChgeLang(.ToolStripMenuItem23.Tag)
            .ToolStripMenuItem57.Text = ChgeLang(.ToolStripMenuItem57.Tag)
            .ToolStripMenuItem27.Text = ChgeLang(.ToolStripMenuItem27.Tag)
            .ສງປດທກລະບບToolStripMenuItem.Text = ChgeLang(.ສງປດທກລະບບToolStripMenuItem.Tag)
            .ປຽນຖານຂມນToolStripMenuItem.Text = ChgeLang(.ປຽນຖານຂມນToolStripMenuItem.Tag)
            .ກວດສອບແລະເຊຕກບຖານຂມນToolStripMenuItem.Text = ChgeLang(.ກວດສອບແລະເຊຕກບຖານຂມນToolStripMenuItem.Tag)
            .ToolStripMenuItem13.Text = ChgeLang(.ToolStripMenuItem13.Tag)
            .ToolStripMenuItem8.Text = ChgeLang(.ToolStripMenuItem8.Tag)
            '====Adjust====
            .ToolStripMenuItem65.Text = ChgeLang(.ToolStripMenuItem65.Tag)
            .ToolStripMenuItem66.Text = ChgeLang(.ToolStripMenuItem66.Tag)
            .ToolStripMenuItem68.Text = ChgeLang(.ToolStripMenuItem68.Tag)
            .ToolStripMenuItem69.Text = ChgeLang(.ToolStripMenuItem69.Tag)
            '====General===
            .MnAt.Text = ChgeLang(.MnAt.Tag)
            .ToolStripMenuItem9.Text = ChgeLang(.ToolStripMenuItem9.Tag)
            .ToolStripMenuItem18.Text = ChgeLang(.ToolStripMenuItem18.Tag)
            .ToolStripMenuItem7.Text = ChgeLang(.ToolStripMenuItem7.Tag)
            .ToolStripMenuItem46.Text = ChgeLang(.ToolStripMenuItem46.Tag)
            .ToolStripMenuItem47.Text = ChgeLang(.ToolStripMenuItem47.Tag)
            .ToolStripMenuItem53.Text = ChgeLang(.ToolStripMenuItem53.Tag)
            '====Acc Report===
            .MnAcReport.Text = ChgeLang(.MnAcReport.Tag)
            .KpkoToolStripMenuItem.Text = ChgeLang(.KpkoToolStripMenuItem.Tag)
            .ບນຊແຍກປະເພດToolStripMenuItem.Text = ChgeLang(.ບນຊແຍກປະເພດToolStripMenuItem.Tag)
            .ໃບດນດຽງບນຊສຳຮອງToolStripMenuItem.Text = ChgeLang(.ໃບດນດຽງບນຊສຳຮອງToolStripMenuItem.Tag)
            .ToolStripMenuItem59.Text = ChgeLang(.ToolStripMenuItem59.Tag)
            '====General Report===
            .ToolStripMenuItem29.Text = ChgeLang(.ToolStripMenuItem29.Tag)
            .ToolStripMenuItem30.Text = ChgeLang(.ToolStripMenuItem30.Tag)
            .ToolStripMenuItem31.Text = ChgeLang(.ToolStripMenuItem31.Tag)
            .ToolStripMenuItem32.Text = ChgeLang(.ToolStripMenuItem32.Tag)
            .ToolStripMenuItem33.Text = ChgeLang(.ToolStripMenuItem33.Tag)
            .ToolStripMenuItem48.Text = ChgeLang(.ToolStripMenuItem48.Tag)
            .ToolStripMenuItem50.Text = ChgeLang(.ToolStripMenuItem50.Tag)
            .ToolStripMenuItem51.Text = ChgeLang(.ToolStripMenuItem51.Tag)
            .ToolStripMenuItem54.Text = ChgeLang(.ToolStripMenuItem54.Tag)
            .ToolStripMenuItem55.Text = ChgeLang(.ToolStripMenuItem55.Tag)
            '==== Report===Aset
            .ToolStripMenuItem41.Text = ChgeLang(.ToolStripMenuItem41.Tag)
            .ToolStripMenuItem42.Text = ChgeLang(.ToolStripMenuItem42.Tag)
            .ToolStripMenuItem43.Text = ChgeLang(.ToolStripMenuItem43.Tag)
            .ToolStripMenuItem44.Text = ChgeLang(.ToolStripMenuItem44.Tag)
            .ToolStripMenuItem45.Text = ChgeLang(.ToolStripMenuItem45.Tag)

            .ToolStripMenuItem2.Text = ChgeLang(.ToolStripMenuItem2.Tag)
            .MuLngL.Text = ChgeLang(.MuLngL.Tag)
            .MuLngE.Text = ChgeLang(.MuLngE.Tag)
            .MnHelp.Text = ChgeLang(.MnHelp.Tag)

            .ToolStripMenuItem60.Text = ChgeLang(.ToolStripMenuItem60.Tag)
            .ToolStripMenuItem61.Text = ChgeLang(.ToolStripMenuItem61.Tag)
            .ToolStripMenuItem71.Text = ChgeLang(.ToolStripMenuItem71.Tag)
            .ToolStripMenuItem72.Text = ChgeLang(.ToolStripMenuItem72.Tag)
            .ToolStripMenuItem67.Text = ChgeLang(.ToolStripMenuItem67.Tag)
            '.ToolStripMenuItem1.Text = ChgeLang(.ToolStripMenuItem1.Tag)
            '.ToolStripMenuItem1.Text = ChgeLang(.ToolStripMenuItem1.Tag)
        End With
        '-------------------------------------------
    End Sub
    Public Sub LoadLng()
        Lng = 0
        LoadAcData("select LngID, LngL , LngE from  Langauge   Order by LngID", RSC)
        With RSC
            Do Until .EOF = True
                LngL(CLng((.Fields("LngID").Value))) = (.Fields("LngL").Value)
                LngE(CLng((.Fields("LngID").Value))) = (.Fields("LngE").Value)
                Lng = Lng + 1
                .MoveNext()
            Loop
        End With
        LoadImageLocation()
 
    End Sub
    Public Sub CallLngStr()
        Dim Lng As String
        If MuLng = "L" Then
            Lng = "LngL"
            LngStr = LngL(CLng(LngId))
        Else
            Lng = "LngE"
            LngStr = LngE(CLng(LngId))
        End If
    End Sub

    Public Sub Sc(ByVal frm As Form)
        For Each Ctl In frm.Controls
        Next
    End Sub

    Public Sub FormOpening()
        FmMain.CmbForm.Items.Clear()
        'Dim sw As New Stopwatch
        'sw.Start()
        For i As Integer = 0 To (My.Application.OpenForms.Count - 1)
            If My.Application.OpenForms.Item(i).Name() <> "FmLogin" And My.Application.OpenForms.Item(i).Name() <> "FmShow" And My.Application.OpenForms.Item(i).Name() <> "FmMain" Then
                FmMain.CmbForm.Items.Add(My.Application.OpenForms.Item(i).Name())
            End If
        Next i
        'FmMain.CmbForm.SelectedIndex = FmMain.CmbForm.SelectedItem
        'sw.Stop()
    End Sub
    Public Sub FormOpening2()
        FmMain.CmbForm.Items.Clear()
        'Dim sw As New Stopwatch
        'sw.Start()
        For i As Integer = 0 To (My.Application.OpenForms.Count - 1)
            If My.Application.OpenForms.Item(i).Name() <> "FmLogin" And My.Application.OpenForms.Item(i).Name() <> "FmShow" And My.Application.OpenForms.Item(i).Name() <> "FmMain" Then
                FmMain.CmbForm.Items.Add(My.Application.OpenForms.Item(i).Name())
            End If
        Next i
        FmMain.CmbForm.SelectedIndex = FmMain.CmbForm.Items.Count - 1
        'sw.Stop()
    End Sub
    Public Sub FrmClosing()
        For i As Integer = 0 To (My.Application.OpenForms.Count - 1)
            On Error GoTo hang
hang:
            If Err.Number = 0 Then
                If My.Application.OpenForms.Item(i).Name() = "FmPreview" Then
                    My.Application.OpenForms.Item(i).Close()
                End If
            Else
                FmPreview.Close()
                'Exit Sub
            End If
        Next i
    End Sub
    '    Public Sub SetControlText(ByVal frm As Form)
    '        If Lang = True Then
    '            ForNme = "Saysettha OT"
    '        Else
    '            ForNme = "Saysettha OT"
    '            'ForNme = "Times New Roman"
    '        End If
    '        ForNme = ""
    '        On Error GoTo ProcedureError
    '        Dim i As Integer = 0
    '        Dim j As Integer = 0
    '        Dim Ctl As Control
    '        Dim msg As MessageBox
    '        For Each Ctl In frm.Controls
    '            Select Case TypeName(Ctl)
    '                Case "MenuStrip", "MenuStrip.items"
    '                    Dim mn As Windows.Forms.MenuStrip
    '                    mn = Ctl
    '                    With mn
    '                        For i = 0 To .Items.Count - 1
    '                            If Lang = True Then
    '                                .Items(i).Text = LngE(CInt(.Items(i).Tag))
    '                            Else
    '                                .Items(i).Text = LngL(CInt(.Items(i).Tag))
    '                            End If
    '                        Next
    '                    End With

    '                    'Case "ListView"
    '                    '    Dim lv As Windows.Forms.ListView
    '                    '    lv = Ctl
    '                    '    With lv
    '                    '        For i = 0 To .Columns.Count - 1
    '                    '            If Lang = True Then
    '                    '                .Columns.Item(i).Text = LngL(CInt(.Columns.Item(i).Tag))
    '                    '            Else
    '                    '                .Columns.Item(i).Text = LngE(CInt(.Columns.Item(i).Tag))
    '                    '            End If
    '                    '        Next
    '                    '    End With

    '                Case "AxVSFlexGrid"
    '                    Dim MG As AxVSFlex8U.AxVSFlexGrid
    '                    MG = Ctl
    '                    For i = 0 To MG.Cols - 1
    '                        If Lang = True Then
    '                            'MG.FormatString = LngL(CInt(Ctl.Tag))
    '                            MG.set_TextMatrix(0, i, LngE(CInt(MG.Tag) + i))
    '                        Else
    '                            'MG.FormatString = LngE(CInt(Ctl.Tag))
    '                            MG.set_TextMatrix(0, i, LngL(CInt(MG.Tag) + i))
    '                        End If
    '                    Next
    '                Case "DataGridView"
    '                    Dim Dg As Windows.Forms.DataGridView
    '                    Dg = Ctl
    '                    With Dg
    '                        For i = 0 To .Columns.Count - 1
    '                            If Lang = True Then
    '                                .Columns.Item(i).HeaderText = LngE(CInt(.Columns.Item(i).Tag))
    '                            Else
    '                                .Columns.Item(i).HeaderText = LngL(CInt(.Columns.Item(i).Tag))
    '                            End If
    '                        Next
    '                    End With

    '                Case "Button"
    '                    If Ctl.GetType Is GetType(Button) Then
    '                        If Lang = True Then
    '                            CType(Ctl, Button).Text = LngE(CInt(Ctl.Tag))
    '                        Else
    '                            CType(Ctl, Button).Text = LngL(CInt(Ctl.Tag))
    '                        End If
    '                    End If

    '                Case "Label", "LinkLabel"
    '                    If Ctl.GetType Is GetType(Label) Then
    '                        If Lang = True Then
    '                            CType(Ctl, Label).Text = LngE(CInt(Ctl.Tag))
    '                        Else
    '                            CType(Ctl, Label).Text = LngL(CInt(Ctl.Tag))
    '                        End If
    '                    End If
    '                    If Ctl.GetType Is GetType(LinkLabel) Then
    '                        If Lang = True Then
    '                            CType(Ctl, LinkLabel).Text = LngE(CInt(Ctl.Tag))
    '                        Else
    '                            CType(Ctl, LinkLabel).Text = LngL(CInt(Ctl.Tag))
    '                        End If
    '                    End If
    '                Case "GroupBox"
    '                    If Ctl.GetType Is GetType(GroupBox) Then
    '                        If Lang = True Then
    '                            CType(Ctl, GroupBox).Text = LngE(CInt(Ctl.Tag))
    '                            For i = 0 To Ctl.Controls.Count - 1
    '                                If Ctl.Controls(i).GetType Is GetType(RadioButton) Then CType(Ctl.Controls(i), RadioButton).Text = LngE(CInt(Ctl.Tag))
    '                                If Ctl.Controls(i).GetType Is GetType(CheckBox) Then CType(Ctl.Controls(i), CheckBox).Text = LngE(CInt(Ctl.Tag))
    '                                If Ctl.Controls(i).GetType Is GetType(Label) Then CType(Ctl.Controls(i), Label).Text = LngE(CInt(Ctl.Tag))
    '                                If Ctl.Controls(i).GetType Is GetType(Button) Then CType(Ctl.Controls(i), Button).Text = LngE(CInt(Ctl.Tag))
    '                            Next
    '                        Else
    '                            CType(Ctl, GroupBox).Text = LngL(CInt(Ctl.Tag))
    '                            For i = 0 To Ctl.Controls.Count - 1
    '                                If Ctl.Controls(i).GetType Is GetType(RadioButton) Then CType(Ctl.Controls(i), RadioButton).Text = LngL(CInt(Ctl.Tag))
    '                                If Ctl.Controls(i).GetType Is GetType(CheckBox) Then CType(Ctl.Controls(i), CheckBox).Text = LngL(CInt(Ctl.Tag))
    '                                If Ctl.Controls(i).GetType Is GetType(Label) Then CType(Ctl.Controls(i), Label).Text = LngL(CInt(Ctl.Tag))
    '                                If Ctl.Controls(i).GetType Is GetType(Button) Then CType(Ctl.Controls(i), Button).Text = LngL(CInt(Ctl.Tag))
    '                            Next
    '                        End If
    '                    End If
    '                Case "RadioButton"
    '                    If Ctl.GetType Is GetType(RadioButton) Then
    '                        If Lang = True Then
    '                            CType(Ctl, RadioButton).Text = LngE(CInt(Ctl.Tag))
    '                        Else
    '                            CType(Ctl, RadioButton).Text = LngL(CInt(Ctl.Tag))
    '                        End If
    '                    End If
    '                Case "CheckBox"
    '                    If Ctl.GetType Is GetType(CheckBox) Then
    '                        If CType(Ctl, CheckBox).Tag <> "" Then
    '                            If ForNme <> "" Then
    '                                CType(Ctl, CheckBox).Font = New System.Drawing.Font(ForNme, CType(Ctl, CheckBox).Font.Size)
    '                            End If
    '                            If Lang = True Then
    '                                CType(Ctl, CheckBox).Text = LngE(CInt(Ctl.Tag))
    '                            Else
    '                                CType(Ctl, CheckBox).Text = LngL(CInt(Ctl.Tag))
    '                            End If
    '                        End If

    '                    End If
    '                Case "CheckBox"
    '                    If Ctl.GetType Is GetType(CheckBox) Then
    '                        If Lang = True Then
    '                            CType(Ctl, CheckBox).Text = LngE(CInt(Ctl.Tag))
    '                        Else
    '                            CType(Ctl, CheckBox).Text = LngL(CInt(Ctl.Tag))
    '                        End If
    '                    End If

    '            End Select
    '        Next
    'ProcedureExit:
    '        Exit Sub
    'ProcedureError:
    '        ' If ErrMsgBox("mDeclare.SetControlCaptionStrings") = vbRetry Then Resume Next
    '    End Sub
    Public Sub SetControlText(ByVal frm As Form)
        'FormOpening()
        FormOpening2()
        'If MuLng = "L" Then
        '    ForNme = "Saysettha OT"
        'Else
        '    ForNme = "Saysettha OT"
        '    'ForNme = "Times New Roman"
        'End If
        'ForNme = ""
        On Error GoTo ProcedureError
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim Ctl As Control
        Dim msg As MessageBox
        For Each Ctl In frm.Controls
            Select Case TypeName(Ctl)
                Case "MenuStrip", "MenuStrip.items"
                    'Case "MenuStrip", "MenuStrip.items"
                    Dim mn As Windows.Forms.MenuStrip
                    mn = Ctl
                    With mn
                        For i = 0 To FmMain.MenuStrip1.Items.Count - 1
                            If .Items(i).Tag <> "" Then
                                If ForNme <> "" Then
                                    FmMain.MenuStrip1.Items.Item(i).Font = New System.Drawing.Font(ForNme, .Font.Size)
                                End If

                                If MuLng = "E" Then
                                    FmMain.MenuStrip1.Items.Item(i).Text = LngE(CInt(.Items(i).Tag))
                                Else
                                    MsgBox(FmMain.MenuStrip1.Items.Item(i))
                                    FmMain.MenuStrip1.Items.Item(i).Text = LngL(CInt(.Items(i).Tag))
                                End If
                                mItem = CType(FmMain.MenuStrip1.Items.Item(i), ToolStripMenuItem)
                            End If
                            For j = 0 To mItem.DropDownItems.Count - 1
                                If mItem.DropDownItems(j).Tag <> "" Then
                                    If ForNme <> "" Then
                                        mItem.DropDownItems(j).Font = New System.Drawing.Font(ForNme, .Font.Size)
                                    End If
                                    If MuLng = "E" Then
                                        mItem.DropDownItems(j).Text = LngE(CInt(mItem.DropDownItems(j).Tag))
                                    Else
                                        mItem.DropDownItems(j).Text = LngL(CInt(mItem.DropDownItems(j).Tag))
                                    End If
                                    mSubItem = mItem.DropDownItems(j)
                                End If
                                For k = 0 To mSubItem.DropDownItems.Count - 1
                                    If mSubItem.DropDownItems(k).Tag <> "" Then
                                        If ForNme <> "" Then
                                            mSubItem.DropDownItems(k).Font = New System.Drawing.Font(ForNme, .Font.Size)
                                        End If
                                        If MuLng = "E" Then
                                            mSubItem.DropDownItems(k).Text = LngE(CInt(mSubItem.DropDownItems(k).Tag))
                                        Else
                                            mSubItem.DropDownItems(k).Text = LngL(CInt(mSubItem.DropDownItems(k).Tag))
                                        End If
                                    End If
                                Next k
                            Next j
                        Next i

                    End With
                Case "ListView"
                    Dim lv As Windows.Forms.ListView
                    lv = Ctl
                    With lv
                        For i = 0 To .Columns.Count - 1
                            'If ForNme <> "" Then
                            '    .Columns.Item(i).Font = New System.Drawing.Font(ForNme, .Font.Size)
                            'End If
                            If MuLng = "E" Then
                                .Columns.Item(i).Text = LngL(CInt(.Columns.Item(i).Tag))
                            Else
                                .Columns.Item(i).Text = LngE(CInt(.Columns.Item(i).Tag))
                            End If
                        Next
                    End With
                Case "AxVSFlexGrid"
                    Dim MG As AxVSFlex8U.AxVSFlexGrid
                    MG = Ctl
                    For i = 0 To MG.Cols - 1
                        If ForNme <> "" Then
                            MG.Font = New System.Drawing.Font(ForNme, MG.Font.Size)
                        End If
                        If MuLng = "L" Then
                            MG.FormatString = LngL(CInt(Ctl.Tag))
                            'MG.set_TextMatrix(0, i, LngE(CInt(MG.Tag) + i))
                        Else
                            MG.FormatString = LngE(CInt(Ctl.Tag))
                            'MG.set_TextMatrix(0, i, LngL(CInt(MG.Tag) + i))
                        End If
                    Next
                Case "DataGridView"
                    Dim Dg As Windows.Forms.DataGridView
                    Dg = Ctl
                    With Dg
                        For i = 0 To .Columns.Count - 1
                            If MuLng = "E" Then
                                .Columns.Item(i).HeaderText = LngE(CInt(.Columns.Item(i).Tag))
                            Else
                                .Columns.Item(i).HeaderText = LngL(CInt(.Columns.Item(i).Tag))
                            End If
                        Next
                    End With
                Case "CheckBox"
                    If Ctl.GetType Is GetType(CheckBox) Then
                        If ForNme <> "" Then
                            CType(Ctl, CheckBox).Font = New System.Drawing.Font(ForNme, CType(Ctl, CheckBox).Font.Size)
                        End If
                        If MuLng = "E" Then
                            CType(Ctl, CheckBox).Text = LngE(CInt(Ctl.Tag))
                        Else
                            CType(Ctl, CheckBox).Text = LngL(CInt(Ctl.Tag))
                        End If
                    End If
                Case "Button"

                    If Ctl.GetType Is GetType(Button) Then
                        MuImgNme = MuImgLct & CType(Ctl, Button).Tag & ".png"
                        Dim fFile1 As New FileInfo(MuImgNme)
                        If Not fFile1.Exists Then
                        Else
                            CType(Ctl, Button).Image = Image.FromFile(MuImgNme)
                        End If
                        If ForNme <> "" Then
                            CType(Ctl, Button).Font = New System.Drawing.Font(ForNme, CType(Ctl, Button).Font.Size)
                        End If
                        If MuLng = "E" Then
                            CType(Ctl, Button).Text = LngE(CInt(Ctl.Tag))
                        Else
                            CType(Ctl, Button).Text = LngL(CInt(Ctl.Tag))
                        End If
                    End If
                Case "Label", "LinkLabel"
                    If Ctl.GetType Is GetType(Label) Then
                        If ForNme <> "" Then
                            CType(Ctl, Label).Font = New System.Drawing.Font(ForNme, CType(Ctl, Label).Font.Size)
                        End If
                        If MuLng = "E" Then
                            If CType(Ctl, Label).Tag <> "" Then
                                CType(Ctl, Label).Text = LngE(CInt(Ctl.Tag))
                            End If

                        Else
                            If CType(Ctl, Label).Tag <> "" Then
                                CType(Ctl, Label).Text = LngL(CInt(Ctl.Tag))
                            End If

                        End If
                    End If
                    If Ctl.GetType Is GetType(LinkLabel) Then
                        If ForNme <> "" Then
                            CType(Ctl, LinkLabel).Font = New System.Drawing.Font(ForNme, CType(Ctl, LinkLabel).Font.Size)
                        End If
                        If MuLng = "E" Then
                            CType(Ctl, LinkLabel).Text = LngE(CInt(Ctl.Tag))
                        Else
                            CType(Ctl, LinkLabel).Text = LngL(CInt(Ctl.Tag))
                        End If
                    End If
                Case "TextBox"
                    If Ctl.GetType Is GetType(TextBox) Then
                        'CType(Ctl, TextBox).Text = MuLng
                    End If
                Case "Panel"
                    If Ctl.GetType Is GetType(Panel) Then
                        If MuLng = "L" Then
                            Dim t1 As String = CInt(Ctl.Tag)
                            If CInt(Ctl.Tag) <> 0 Then
                                For i = 0 To Ctl.Controls.Count - 1
                                    If Ctl.Controls(i).GetType Is GetType(Label) Then
                                        Dim t As String = CType(Ctl.Controls(i), Label).Tag()


                                        If ForNme <> "" Then
                                            CType(Ctl.Controls(i), Label).Font = New System.Drawing.Font(ForNme, CType(Ctl.Controls(i), Label).Font.Size)
                                        End If
                                        If CType(Ctl.Controls(i), Label).Tag <> "" Then
                                            If Ctl.Controls(i).GetType Is GetType(Label) Then CType(Ctl.Controls(i), Label).Text = LngL(t)
                                        End If

                                    End If
                                    If Ctl.Controls(i).GetType Is GetType(CheckBox) Then
                                        Dim t As String = CType(Ctl.Controls(i), CheckBox).Tag()
                                        If ForNme <> "" Then
                                            CType(Ctl.Controls(i), CheckBox).Font = New System.Drawing.Font(ForNme, CType(Ctl.Controls(i), CheckBox).Font.Size)
                                        End If
                                        If CType(Ctl.Controls(i), CheckBox).Tag <> "" Then
                                            If Ctl.Controls(i).GetType Is GetType(CheckBox) Then CType(Ctl.Controls(i), CheckBox).Text = LngL(t)
                                        End If

                                    End If
                                    If Ctl.Controls(i).GetType Is GetType(RadioButton) Then
                                        Dim t As String = CType(Ctl.Controls(i), RadioButton).Tag()
                                        If ForNme <> "" Then
                                            CType(Ctl.Controls(i), RadioButton).Font = New System.Drawing.Font(ForNme, CType(Ctl.Controls(i), RadioButton).Font.Size)
                                        End If
                                        If CType(Ctl.Controls(i), RadioButton).Tag <> "" Then
                                            If Ctl.Controls(i).GetType Is GetType(RadioButton) Then CType(Ctl.Controls(i), RadioButton).Text = LngL(t)
                                        End If

                                    End If
                                    If Ctl.Controls(i).GetType Is GetType(Button) Then

                                        MuImgNme = MuImgLct & CType(Ctl.Controls(i), Button).Tag & ".png"
                                        'MsgBox(MuImgNme)
                                        Dim fFile1 As New FileInfo(MuImgNme)
                                        If Not fFile1.Exists Then
                                        Else
                                            CType(Ctl.Controls(i), Button).Image = Image.FromFile(MuImgNme)
                                        End If

                                        Dim t As String = CType(Ctl.Controls(i), Button).Tag()
                                        If ForNme <> "" Then
                                            CType(Ctl.Controls(i), Button).Font = New System.Drawing.Font(ForNme, CType(Ctl.Controls(i), Button).Font.Size)
                                        End If
                                        If CType(Ctl.Controls(i), Button).Tag <> "" Then
                                            If Ctl.Controls(i).GetType Is GetType(Button) Then CType(Ctl.Controls(i), Button).Text = LngL(t)
                                        End If
                                    End If
                                Next
                                CType(Ctl, Panel).Text = LngL(Ctl.Tag)
                            End If
                        Else
                            Dim t1 As String = CInt(Ctl.Tag)
                            If CInt(Ctl.Tag) <> 0 Then

                                For i = 0 To Ctl.Controls.Count - 1
                                    If Ctl.Controls(i).GetType Is GetType(Label) Then
                                        Dim t As String = CType(Ctl.Controls(i), Label).Tag()
                                        If ForNme <> "" Then
                                            CType(Ctl.Controls(i), Label).Font = New System.Drawing.Font(ForNme, CType(Ctl.Controls(i), Label).Font.Size)
                                        End If
                                        If CType(Ctl.Controls(i), Label).Tag <> "" Then
                                            If Ctl.Controls(i).GetType Is GetType(Label) Then CType(Ctl.Controls(i), Label).Text = LngE(t)
                                        End If
                                    End If
                                    If Ctl.Controls(i).GetType Is GetType(CheckBox) Then
                                        Dim t As String = CType(Ctl.Controls(i), CheckBox).Tag()
                                        If ForNme <> "" Then
                                            CType(Ctl.Controls(i), CheckBox).Font = New System.Drawing.Font(ForNme, CType(Ctl.Controls(i), CheckBox).Font.Size)
                                        End If
                                        If CType(Ctl.Controls(i), CheckBox).Tag <> "" Then
                                            If Ctl.Controls(i).GetType Is GetType(CheckBox) Then CType(Ctl.Controls(i), CheckBox).Text = LngE(t)
                                        End If
                                    End If
                                    If Ctl.Controls(i).GetType Is GetType(RadioButton) Then
                                        Dim t As String = CType(Ctl.Controls(i), RadioButton).Tag()
                                        If ForNme <> "" Then
                                            CType(Ctl.Controls(i), RadioButton).Font = New System.Drawing.Font(ForNme, CType(Ctl.Controls(i), RadioButton).Font.Size)
                                        End If
                                        If CType(Ctl.Controls(i), RadioButton).Tag <> "" Then
                                            If Ctl.Controls(i).GetType Is GetType(RadioButton) Then CType(Ctl.Controls(i), RadioButton).Text = LngE(t)
                                        End If
                                    End If
                                    If Ctl.Controls(i).GetType Is GetType(Button) Then



                                        MuImgNme = MuImgLct & CType(Ctl.Controls(i), Button).Tag & ".png"
                                        Dim fFile1 As New FileInfo(MuImgNme)
                                        If Not fFile1.Exists Then
                                        Else
                                            CType(Ctl.Controls(i), Button).Image = Image.FromFile(MuImgNme)
                                        End If

                                        Dim t As String = CType(Ctl.Controls(i), Button).Tag()
                                        If ForNme <> "" Then
                                            CType(Ctl.Controls(i), Button).Font = New System.Drawing.Font(ForNme, CType(Ctl.Controls(i), Button).Font.Size)
                                        End If
                                        If CType(Ctl.Controls(i), Button).Tag <> "" Then
                                            If Ctl.Controls(i).GetType Is GetType(Button) Then CType(Ctl.Controls(i), Button).Text = LngE(t)
                                        End If

                                    End If
                                Next
                                CType(Ctl, Panel).Text = LngE(CInt(Ctl.Tag))
                            End If
                        End If
                    End If
                    ' ===
                Case "GroupBox"
                    If Ctl.GetType Is GetType(GroupBox) Then
                        If MuLng = "L" Then
                            Dim t1 As String = CInt(Ctl.Tag)
                            If CInt(Ctl.Tag) <> 0 Then
                                For i = 0 To Ctl.Controls.Count - 1
                                    If Ctl.Controls(i).GetType Is GetType(Label) Then
                                        Dim t As String = CType(Ctl.Controls(i), Label).Tag()
                                        If ForNme <> "" Then
                                            CType(Ctl.Controls(i), Label).Font = New System.Drawing.Font(ForNme, CType(Ctl.Controls(i), Label).Font.Size)
                                        End If
                                        If Ctl.Controls(i).GetType Is GetType(Label) Then CType(Ctl.Controls(i), Label).Text = LngL(t)
                                    End If
                                    If Ctl.Controls(i).GetType Is GetType(CheckBox) Then
                                        Dim t As String = CType(Ctl.Controls(i), CheckBox).Tag()
                                        If ForNme <> "" Then
                                            CType(Ctl.Controls(i), CheckBox).Font = New System.Drawing.Font(ForNme, CType(Ctl.Controls(i), CheckBox).Font.Size)
                                        End If
                                        If Ctl.Controls(i).GetType Is GetType(CheckBox) Then CType(Ctl.Controls(i), CheckBox).Text = LngL(t)
                                    End If
                                    If Ctl.Controls(i).GetType Is GetType(RadioButton) Then
                                        Dim t As String = CType(Ctl.Controls(i), RadioButton).Tag()
                                        If ForNme <> "" Then
                                            CType(Ctl.Controls(i), RadioButton).Font = New System.Drawing.Font(ForNme, CType(Ctl.Controls(i), RadioButton).Font.Size)
                                        End If
                                        If Ctl.Controls(i).GetType Is GetType(RadioButton) Then CType(Ctl.Controls(i), RadioButton).Text = LngL(t)
                                    End If
                                    If Ctl.Controls(i).GetType Is GetType(Button) Then


                                        MuImgNme = MuImgLct & CType(Ctl.Controls(i), Button).Tag & ".png"
                                        Dim fFile1 As New FileInfo(MuImgNme)
                                        If Not fFile1.Exists Then
                                        Else
                                            CType(Ctl.Controls(i), Button).Image = Image.FromFile(MuImgNme)
                                        End If

                                        Dim t As String = CType(Ctl.Controls(i), Button).Tag()
                                        If ForNme <> "" Then
                                            CType(Ctl.Controls(i), Button).Font = New System.Drawing.Font(ForNme, CType(Ctl.Controls(i), Button).Font.Size)
                                        End If
                                        If Ctl.Controls(i).GetType Is GetType(Button) Then CType(Ctl.Controls(i), Button).Text = LngL(t)
                                    End If
                                Next
                                CType(Ctl, GroupBox).Text = LngL(Ctl.Tag)
                            End If
                        Else
                            Dim t1 As String = CInt(Ctl.Tag)
                            If CInt(Ctl.Tag) <> 0 Then
                                For i = 0 To Ctl.Controls.Count - 1
                                    If Ctl.Controls(i).GetType Is GetType(Label) Then
                                        Dim t As String = CType(Ctl.Controls(i), Label).Tag()
                                        If ForNme <> "" Then
                                            CType(Ctl.Controls(i), Label).Font = New System.Drawing.Font(ForNme, CType(Ctl.Controls(i), Label).Font.Size)
                                        End If
                                        If Ctl.Controls(i).GetType Is GetType(Label) Then CType(Ctl.Controls(i), Label).Text = LngE(t)
                                    End If
                                    If Ctl.Controls(i).GetType Is GetType(CheckBox) Then
                                        Dim t As String = CType(Ctl.Controls(i), CheckBox).Tag()
                                        If ForNme <> "" Then
                                            CType(Ctl.Controls(i), CheckBox).Font = New System.Drawing.Font(ForNme, CType(Ctl.Controls(i), CheckBox).Font.Size)
                                        End If
                                        If Ctl.Controls(i).GetType Is GetType(CheckBox) Then CType(Ctl.Controls(i), CheckBox).Text = LngE(t)
                                    End If
                                    If Ctl.Controls(i).GetType Is GetType(RadioButton) Then
                                        Dim t As String = CType(Ctl.Controls(i), RadioButton).Tag()
                                        If ForNme <> "" Then
                                            CType(Ctl.Controls(i), RadioButton).Font = New System.Drawing.Font(ForNme, CType(Ctl.Controls(i), RadioButton).Font.Size)
                                        End If
                                        If Ctl.Controls(i).GetType Is GetType(RadioButton) Then CType(Ctl.Controls(i), RadioButton).Text = LngE(t)
                                    End If
                                    If Ctl.Controls(i).GetType Is GetType(Button) Then

                                        MuImgNme = MuImgLct & CType(Ctl.Controls(i), Button).Tag & ".png"
                                        Dim fFile1 As New FileInfo(MuImgNme)
                                        If Not fFile1.Exists Then
                                        Else
                                            CType(Ctl.Controls(i), Button).Image = Image.FromFile(MuImgNme)
                                        End If





                                        Dim t As String = CType(Ctl.Controls(i), Button).Tag()
                                        If ForNme <> "" Then
                                            CType(Ctl.Controls(i), Button).Font = New System.Drawing.Font(ForNme, CType(Ctl.Controls(i), Button).Font.Size)
                                        End If
                                        If Ctl.Controls(i).GetType Is GetType(Button) Then CType(Ctl.Controls(i), Button).Text = LngE(t)
                                    End If
                                Next
                                CType(Ctl, GroupBox).Text = LngE(CInt(Ctl.Tag))
                            End If
                        End If
                    End If
                Case "RadioButton"
                    If Ctl.GetType Is GetType(RadioButton) Then
                        If ForNme <> "" Then
                            CType(Ctl, RadioButton).Font = New System.Drawing.Font(ForNme, CType(Ctl, RadioButton).Font.Size)
                        End If
                        If MuLng = "E" Then
                            CType(Ctl, RadioButton).Text = LngE(CInt(Ctl.Tag))
                        Else
                            CType(Ctl, RadioButton).Text = LngL(CInt(Ctl.Tag))
                        End If
                    End If
            End Select
        Next
ProcedureExit:
        Exit Sub
ProcedureError:
        ' If ErrMsgBox("mDeclare.SetControlCaptionStrings") = vbRetry Then Resume Next
    End Sub
    Public Sub MsgQtin()
        Call CallLngStr()
        If MessageBox.Show(LngStr, "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            MsgSL = 1
        Else
            MsgSL = 0
        End If
    End Sub
    Public Sub MsgRpt()
        Call CallLngStr()
        MsgBox(LngStr)
    End Sub
    Public Function ErrMsgBox(ByVal Msg As String) As Integer
        ErrMsgBox = MsgBox("Error: " & Err.Number & ". " & Err.Description, vbRetryCancel + vbCritical, Msg)
    End Function
    Public Sub ChgChildForm()
        For Each ChildForm As Form In FmMain.MdiChildren
            SetControlText(ChildForm)
        Next
    End Sub


End Module
