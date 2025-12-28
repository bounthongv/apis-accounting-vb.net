Module MdLswin
    Declare Function InitLSWin Lib "LSWin32.dll" Alias "Initialize" () As Long
    Declare Function Romanize Lib "LSWin32.dll" (ByVal lpData As String, ByVal lpNewData As String, ByVal cbBuf As Long, ByVal lpFontName As String) As Long
    Declare Function ConvertText Lib "LSWin32.dll" (ByVal lpData As String, ByVal lpNewData As String, ByVal cbBuf As Long, ByVal lpFontName As String) As Long
    Declare Function SetKeyboard Lib "LSWin32.dll" (ByVal kbName As String) As Long
    Declare Function SortString Lib "LSWin32.dll" (ByVal Term As String, ByVal SortKey As String, ByVal BufSize As Long, ByVal FontName As String) As Long
    Declare Function WrapText Lib "LSWin32.dll" (ByVal OldData As String, ByVal NewData As String, ByVal BufSize As Long, ByVal FontName As String, ByVal Mode As Long) As Long
    Declare Function LSWinOptions Lib "LSWin32.dll" (ByVal hwnd As Long) As Long
    Public Function SetLaoOptions()
        If InitLSWin() Then
            SetLaoOptions = True
        Else
            SetLaoOptions = False
        End If
    End Function
    'Public Function SetLao(ByVal Mode)
    '    'On Error Resume Next
    '    'Dim rslt
    '    'SetLao = 0
    '    'If InitLSWin() = 0 Then Exit Function
    '    'Select Case Mode
    '    '    Case 0 'disable keyboard translation
    '    '        rslt = SetKeyboard("")
    '    '    Case 1 'standard LSWin (Saysettha OT) coding
    '    '        rslt = SetKeyboard("Unicode")
    '    '    Case 2 'standard LSWin (Saysettha Lao) coding
    '    '        rslt = SetKeyboard("LSwin")
    '    'End Select
    'End Function
End Module
