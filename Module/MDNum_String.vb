Module MDNum_String
    Public curr As String
    Public RDB As String
    Public DividePage As Double = 0
    Public Rem_Dr, Rem_Cr As String
    Public Last_page, P As Integer
    Public StrSQL, ConString As String
    Public MdStartDate As Date
    Public MdToDate As Date
    Public mCompStr As String
    Public RptDt, sql As String
    Public GrpIDOrder As String
    Public GrpNm As String
    Public myTemp2 As String
    Public myTemp3 As String
    Public PCNM As String
    Public mClose As Boolean
    Public mEdit As String
    Public myTemp As String
    Public myTemp1 As String
 
    Public mStrEdit As Integer
    Public mDate As Date = Date.Today

    Public Function InwordsEng(ByVal str1 As String) As String
        Dim Istr As String
        Dim nos
        nos = CDbl(str1)
        'If Val(str1) = 0 Then
        'Exit Function
        'End If
        Istr = ""
        Select Case nos
            Case 0 To 19
                Istr = Choose(nos + 1, "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "seventeen", "Eightteen", "Nineteen")
            Case 20, 30, 40, 50, 60, 70, 80, 90
                Istr = Choose(nos / 10, "Kip", "Twenty", "Thirty", "Fourty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety") 'RAVI
            Case 21, 31, 41, 51, 61, 71, 81, 91
                '' Istr = Choose(nos / 10, "ກີບ", "ຊາວເອັດ", "ສາມສິບເອັດ", "ສີ່ສິບເອັດ", "ຫ້າສິບເອັດ", "ຫົກສິບເອັດ", "ເຈັດສິລເອັດ", "ແປດສິບເອັດ", "Ninetyone") 'RAVI
                Istr = Choose(nos / 10, "Kip", "Twentyone", "one", "Fourtyone", "Fiftyone", "Sixtyone", "Seventyone", "Eightyone", "Ninetyone") 'RAVI
            Case 22 To 30, 32 To 40, 42 To 50, 52 To 60, 62 To 70, 72 To 80, 82 To 99
                Istr = InwordsEng(Left(nos, 1) & 0) & InwordsEng(Right(nos, 1))
            Case 101 To 999
                If CInt(nos) <> 0 Then
                    Istr = InwordsEng(Left(nos, 1)) & "Hundred" & InwordsEng(Right(nos, 2))
                Else
                    Istr = InwordsEng(Left(nos, 1)) & "Hundred" & InwordsEng(Right(nos, 2))
                End If
            Case 1001 To 9999
                Istr = InwordsEng(Left(nos, 1)) & "Thousand" & InwordsEng(Right(nos, 3))
            Case 10000 To 99999
                Istr = InwordsEng(Mid(nos, 1, 2)) & "Thousand" & InwordsEng(Right(nos, 3))
            Case 100000 To 999999
                Istr = InwordsEng(Left(nos, 1)) & "HundredThousand" & InwordsEng(Right(nos, 5))
            Case 1000000 To 9999999
                Istr = InwordsEng(Left(nos, 1)) & "Million" & InwordsEng(Right(nos, 6))
            Case 10000000 To 99999999
                Istr = InwordsEng(Left(nos, 2)) & "Million" & InwordsEng(Right(nos, 6))
            Case 100000000 To 999999999
                Istr = InwordsEng(Left(nos, 3)) & "Million" & InwordsEng(Right(nos, 6))
            Case 1000000000 To 9999999999.0#
                Istr = InwordsEng(Left(nos, 4)) & "Million" & InwordsEng(Right(nos, 6))
            Case 10000000000.0# To 99999999999.0#
                Istr = InwordsEng(Left(nos, 5)) & "Million" & InwordsEng(Right(nos, 6))
            Case 100000000000.0# To 999999999999.0#
                Istr = InwordsEng(Left(nos, 6)) & "Million" & InwordsEng(Right(nos, 6))
            Case 100
                Istr = InwordsEng(Left(nos, 1)) & "Mundred"
            Case 1000
                Istr = InwordsEng(Left(nos, 1)) & "Thousand"
            Case 100000
                Istr = InwordsEng(Left(nos, 1)) & "HundredThousand"
            Case 10000000
                Istr = InwordsEng(Left(nos, 2)) & "Million"
        End Select
        InwordsEng = Istr
    End Function

    Public Function CMoneyEng(ByVal strr As String) As String
        Dim rs
        Dim i As Integer
        Dim A As String
        A = ""
        rs = (Right(strr, (Len(strr) - InStr(1, strr, "."))))
        'If Val(strr) = 0 Then
        'Exit Function
        'End If
        If InStr(1, strr, ".") <> 0 Then
            If Len(Right(strr, (Len(strr) - InStr(1, strr, ".")))) = 1 Then
                If Val(Mid(strr, 1, InStr(1, strr, "."))) = 0 Then
                    CMoneyEng = PiaseEng(rs * 10) & curr
                Else
                    'CMoneyEng = PiaseEng(Left(strr, InStr(1, strr, ".") - 1)) & " ຈຸດ" & PiaseEng(rs * 10) & curr
                    CMoneyEng = PiaseEng(Left(strr, InStr(1, strr, ".") - 1)) & " ," & PiaseEng(rs * 10) & curr
                End If
            Else
                '            If Val(Mid(strr, 1, InStr(1, strr, "."))) = 0 Then
                '                       CMoney = Piase(Right(strr, (Len(strr) - InStr(1, strr, ".")))) & " ¡ó®"
                '            Else
                For i = 1 To Len(Right(strr, (Len(strr) - InStr(1, strr, "."))))
                    A = A & PiaseEng(Mid(Right(strr, (Len(strr) - InStr(1, strr, "."))), i, 1))
                Next
                If Val(Right(strr, (Len(strr) - InStr(1, strr, ".")))) > 0 Then
                    If Val(Left(strr, InStr(1, strr, ".") - 1)) = 0 Then
                        ' CMoneyEng = "Zero" & " ຈຸດ" & A & curr
                        CMoneyEng = "Zero" & " ," & A & curr
                    Else
                        ' CMoneyEng = InwordsEng(Left(strr, InStr(1, strr, ".") - 1)) & " ຈຸດ" & A & curr
                        CMoneyEng = InwordsEng(Left(strr, InStr(1, strr, ".") - 1)) & " ," & A & curr
                    End If
                Else
                    CMoneyEng = InwordsEng(Left(strr, InStr(1, strr, ".") - 1)) & curr
                End If
                '       End If
            End If
        Else
            CMoneyEng = InwordsEng(strr) & curr
        End If
    End Function
    Public Function PiaseEng(ByVal pi As String) As String
        Dim Istr As String
        Istr = ""
        Dim nos
        nos = pi
        Select Case nos
            Case 0 To 9 : Istr = Choose(nos + 1, "Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine")
        End Select
        PiaseEng = Istr
    End Function
    '''aaaaaaaaaaaaaaaaaaaaaaaaaaaaa
    Public Function Inwords(ByVal str1 As String) As String
        Dim Istr As String
        Istr = ""
        Dim nos
        nos = CDbl(str1)
        'If Val(str1) = 0 Then
        'Exit Function
        'End If

        Select Case nos
            Case 0 To 19
                Istr = Choose(nos + 1, "", "ໜຶ່ງ", "ສອງ", "ສາມ´", "ສີ່", "ຫ້າ", "ຫົກ", "ເຈັດ", "ແປດ", "ເກົ້າ", "ສິບ", "ສິບເອັດ", "ສິບສອງ", "ສິບສາມ´", "ສິບສີ່", "ສິບຫ້າ", "ສິບຫົກ", "ສິບເຈັດ", "ສິບແປດ", "ສິບເກົ້າ")
            Case 20, 30, 40, 50, 60, 70, 80, 90
                Istr = Choose(nos / 10, "ກີບ", "ຊາວ", "ສາມສິບ", "ສີ່ສິບ", "ຫ້າສິບ", "ຫົກສິບ", "ເຈັດສິບ", "ແປດສິບ", "ເກົ້າສິບ") 'RAVI
            Case 21, 31, 41, 51, 61, 71, 81, 91
                Istr = Choose(nos / 10, "ກີບ", "ຊາວເອັດ", "ສາມສິບເອັດ", "ສີ່ສິບເອັດ", "ຫ້າສິບເອັດ", "ຫົກສິບເອັດ", "ເຈັດສິລເອັດ", "ແປດສິບເອັດ", "ເກົ້າສິບເອັດ") 'RAVI
            Case 22 To 30, 32 To 40, 42 To 50, 52 To 60, 62 To 70, 72 To 80, 82 To 99
                Istr = Inwords(Left(nos, 1) & 0) & Inwords(Right(nos, 1))
            Case 101 To 999
                If CInt(nos) <> 0 Then
                    Istr = Inwords(Left(nos, 1)) & "ຮ້ອຍ" & Inwords(Right(nos, 2))
                Else
                    Istr = Inwords(Left(nos, 1)) & "ຮ້ອຍ" & Inwords(Right(nos, 2))
                End If
            Case 1001 To 9999
                Istr = Inwords(Left(nos, 1)) & "ພັນ" & Inwords(Right(nos, 3))
            Case 10000 To 99999
                Istr = Inwords(Mid(nos, 1, 2)) & "ພັນ" & Inwords(Right(nos, 3))
            Case 100000 To 999999
                Istr = Inwords(Left(nos, 1)) & "ແສນ" & Inwords(Right(nos, 5))
            Case 1000000 To 9999999
                Istr = Inwords(Left(nos, 1)) & "ລ້ານ" & Inwords(Right(nos, 6))
            Case 10000000 To 99999999
                Istr = Inwords(Left(nos, 2)) & "ລ້ານ" & Inwords(Right(nos, 6))
            Case 100000000 To 999999999
                Istr = Inwords(Left(nos, 3)) & "ລ້ານ" & Inwords(Right(nos, 6))
            Case 1000000000 To 9999999999.0#
                Istr = Inwords(Left(nos, 4)) & "ລ້ານ" & Inwords(Right(nos, 6))
            Case 10000000000.0# To 99999999999.0#
                Istr = Inwords(Left(nos, 5)) & "ລ້ານ" & Inwords(Right(nos, 6))
            Case 100000000000.0# To 999999999999.0#
                Istr = Inwords(Left(nos, 6)) & "ລ້ານ" & Inwords(Right(nos, 6))
            Case 100
                Istr = Inwords(Left(nos, 1)) & "ຮ້ອຍ"
            Case 1000
                Istr = Inwords(Left(nos, 1)) & "ພັນ"
            Case 100000
                Istr = Inwords(Left(nos, 1)) & "ແສນ"
            Case 10000000
                Istr = Inwords(Left(nos, 2)) & "ລ້ານ"
        End Select
        Inwords = Istr
    End Function

    Public Function CMoney(ByVal strr As String) As String
        Dim rs
        Dim i As Integer
        Dim A As String
        A = ""
        rs = (Right(strr, (Len(strr) - InStr(1, strr, "."))))
        'If Val(strr) = 0 Then
        'Exit Function
        'End If
        If InStr(1, strr, ".") <> 0 Then
            If Len(Right(strr, (Len(strr) - InStr(1, strr, ".")))) = 1 Then
                If Val(Mid(strr, 1, InStr(1, strr, "."))) = 0 Then
                    CMoney = Piase(rs * 10) & curr
                Else
                    CMoney = Inwords(Left(strr, InStr(1, strr, ".") - 1)) & " ຈຸດ" & Piase(rs * 10) & curr
                End If
            Else
                '            If Val(Mid(strr, 1, InStr(1, strr, "."))) = 0 Then
                '                       CMoney = Piase(Right(strr, (Len(strr) - InStr(1, strr, ".")))) & " ¡ó®"
                '            Else
                For i = 1 To Len(Right(strr, (Len(strr) - InStr(1, strr, "."))))
                    A = A & Piase(Mid(Right(strr, (Len(strr) - InStr(1, strr, "."))), i, 1))
                Next
                If Val(Right(strr, (Len(strr) - InStr(1, strr, ".")))) > 0 Then
                    If Val(Left(strr, InStr(1, strr, ".") - 1)) = 0 Then
                        CMoney = "ສູນ" & " ຈຸດ" & A & curr
                    Else
                        CMoney = Inwords(Left(strr, InStr(1, strr, ".") - 1)) & " ຈຸດ" & A & curr
                    End If
                Else
                    CMoney = Inwords(Left(strr, InStr(1, strr, ".") - 1)) & curr
                End If
                '       End If
            End If
        Else
            CMoney = Inwords(strr) & curr
        End If
    End Function
    Public Function Piase(ByVal pi As String) As String
        Dim Istr As String
        Istr = ""
        Dim nos
        nos = pi
        Select Case nos
            Case 0 To 9 : Istr = Choose(nos + 1, "ສູນ", "ໜຶ່ງ", "ສອງ", "ສາມ", "ສີ່", "ຫ້າ", "ຫົກ", "ເຈັດ", "ແປດ", "ເກົ້າ")
        End Select
        Piase = Istr
    End Function

    Public Sub CloseInfo()
        FmMain.IsMdiContainer = True

        'FmMain.Label17.Visible = False
    End Sub

    Public Sub OpenInfo()

        'FmMain.IsMdiContainer = False
        'FmMain.Label17.Visible = True
    End Sub

End Module
