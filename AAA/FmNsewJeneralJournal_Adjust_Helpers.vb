
    Private Function GetValue(ByVal cellVal As Object) As Double
        If cellVal Is Nothing OrElse cellVal.ToString() = "" Then
            Return 0
        Else
            If IsNumeric(cellVal) Then
                Return CDbl(cellVal)
            End If
            Return 0
        End If
    End Function

    Private Function GetString(ByVal cellVal As Object) As String
        If cellVal Is Nothing Then Return ""
        Return cellVal.ToString()
    End Function
