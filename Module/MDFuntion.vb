
Option Explicit On
Option Strict On

Imports AxVSFlex8U
'MDSeriaCom = CDbl(Mid(MDPartitionSeria.ShowDriveInfo(Application.StartupPath), 2, Len(Trim(MDPartitionSeria.ShowDriveInfo(Application.StartupPath)) - 1)))

Module MDFuntion
    Public Function Apostrophe(ByVal sFieldString As String) As String
        If CBool(InStr(sFieldString, "'")) Then
            Dim iLen As Integer
            Dim i As Integer
            Dim apostr As String
            iLen = Len(sFieldString)
            i = 1
            Do While i <= iLen
                If Mid(sFieldString, i, 1) = "'" Then
                    apostr = CStr(i)
                    sFieldString = Left(sFieldString, CInt(apostr)) & "'" & Right(sFieldString, iLen - CInt(apostr))
                    iLen = Len(sFieldString)
                    i = i + 1
                End If
                i = i + 1
            Loop
        End If
        Apostrophe = sFieldString
    End Function
  
    Public Sub Check_Pair(ByVal mFG As AxVSFlexGrid, ByVal mCol As Integer)
        Dim Pair, Pair1, i As Integer
        Dim Plus As Boolean
        Plus = False
        Pair = 1
        For i = 1 To mFG.Rows - 2
            If mFG.get_TextMatrix(i, 0) <> "" Then
                Pair1 = Pair
                Plus = False
            Else
                If Plus = False Then
                    Pair = Pair + 1
                    Plus = True
                End If
            End If
            mFG.set_TextMatrix(i, mCol, CStr(Pair1))
        Next i
    End Sub
    Public Sub GridEdit(ByVal g As AxVSFlexGrid, ByVal C As TextBox, ByVal KeyAscii As Integer)
        '--------------------------------------------------
        ' prepare the control
        ' On Error Resume Next

        '--------------------------------------------------
        ' show it at the right place
        If g.FillStyle = 0 Then
            If g.Row <> g.RowSel Or g.Col <> g.ColSel Then
                g.RowSel = g.Row
                g.ColSel = g.Col
            End If
        End If

        ' and let it rip
        On Error Resume Next
        On Error GoTo 0

    End Sub
    Public Sub EditKeyCode(ByVal g As AxVSFlexGrid, ByVal C As TextBox, ByVal KeyCode%, ByVal Shift%)
        ' standard edit control processing
        Select Case KeyCode
            Case 27 ' esc
                C.Visible = False
                g.Focus()
            Case 13 ' enter
                g.Focus()
                Application.DoEvents()
            Case 40 ' down
                g.Focus()
                Application.DoEvents()
                If g.Row < g.Rows - 1 Then
                    g.Row = g.Row + 1
                End If
            Case 38 ' up
                g.Focus()
                Application.DoEvents()
                If g.Row > g.FixedRows Then
                    g.Row = g.Row - 1
                End If
        End Select
    End Sub
    Public Sub load_Cmb(ByVal sql As String, ByVal Str As String, ByVal CmbFg As ComboBox)
        Dim Rsgf As New ADODB.Recordset
        Dim Rsshop As New ADODB.Recordset
        With Rsgf
            If .State = ConnectionState.Open Then .Close()
            .ActiveConnection = CNN
            .CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(sql)

        End With

        While Not Rsgf.EOF
            CmbFg.BeginUpdate()
            CmbFg.Items.Add(IIf(IsDBNull(Rsgf.Fields(Str).Value) = True, "", Trim(CStr(Rsgf.Fields(Str).Value))))
            Rsgf.MoveNext()
            CmbFg.EndUpdate()
        End While
    End Sub

End Module
