Module MDRateSetting
    Public MDJPY, MDLAK, MDTHB, MDUSD, MDEUR, MDUSD_LAK, MDTHB_LAK, MDEUR_LAK, MDEUR_THB, MDUSD_THB, MDEUR_USD As Double
    Public MDJPY_LAK, MDJPY_THB, MDJPY_USA As Double
    'Public MD_Rate As Double
    Public MDJPY1, MDLAK1, MDTHB1, MDUSD1, MDEUR1, MDUSD_LAK1, MDTHB_LAK1, MDEUR_LAK1, MDEUR_THB1, MDUSD_THB1, MDEUR_USD1 As Double
    Public MDRate_Curr, MDR_Curr As String
    ' Public StrDate As String
    Public StrDate As Date
    Public StrMM As Date
    Public Sub RateSetting()
        Dim Rs As New ADODB.Recordset

        MD_Rate = 1
        MDLAK = 1
        MDTHB = 1
        MDUSD = 1
        MDEUR = 1
        MDUSD_LAK = 1
        MDTHB_LAK = 1
        MDEUR_LAK = 1
        MDEUR_THB = 1
        MDUSD_THB = 1
        MDEUR_USD = 1

        With Rs
            Call LoadSqlData("select top 1  * from AP_Rate_history where 1=1  " & MDRate_DT & " " & SS_Curr & " " & MDR_Curr & "   ORDER BY rate_dt DESC ", Rs)
            If .RecordCount > 0 Then
                StrDate = Trim(.Fields("rate_dt").Value)
                MD_Rate = (.Fields("Rate").Value)
                MD_Rate2 = (.Fields("Rate2").Value)
                MDLAK = Trim(.Fields("Rate").Value)
                MDTHB = Trim(.Fields("Rate").Value)
                MDUSD = Trim(.Fields("Rate").Value)
                MDEUR = Trim(.Fields("Rate").Value)
                MDUSD_LAK = (.Fields("Rate").Value)
                MDTHB_LAK = (.Fields("Rate").Value)
                MDEUR_LAK = (.Fields("Rate").Value)
                MDEUR_THB = (.Fields("Rate").Value)
                MDUSD_THB = (.Fields("Rate").Value)
                MDEUR_USD = (.Fields("Rate").Value)
                MDRate_Curr = Trim(.Fields("Curr").Value)

            Else
                MD_Rate = 1
                Call LoadSqlData("select * from AP_Rate_history where com_id='00'  ORDER BY rate_dt DESC ", Rs)
                If Rs.RecordCount > 0 Then

                    MD_Rate = (.Fields("Rate").Value)
                    MD_Rate2 = (.Fields("Rate2").Value)
                    'Call LoadRs("select * from AP_Rate WHERE status=1", Rs)
                    MDLAK = Trim(.Fields("Rate").Value.ToString)
                    MDTHB = Trim(.Fields("Rate").Value)
                    MDUSD = Trim(.Fields("Rate").Value)
                    MDEUR = Trim(.Fields("Rate").Value)
                    MDUSD_LAK = (.Fields("Rate").Value)
                    MDTHB_LAK = (.Fields("Rate").Value)
                    MDEUR_LAK = (.Fields("Rate").Value)
                    MDEUR_THB = (.Fields("Rate").Value)
                    MDUSD_THB = (.Fields("Rate").Value)
                    MDEUR_USD = (.Fields("Rate").Value)
                    MDRate_Curr = Trim(.Fields("Curr").Value)

                End If

            End If
        End With
        '==============================================

        Dim Rss As New ADODB.Recordset
        With Rss


            Call LoadSqlData("select top 1  * from AP_Rate_history where 1=1  " & MDRate_DT & "  and curr='USD' ORDER BY rate_dt DESC ", Rss)
            If .RecordCount > 0 Then
                MDUSD_LAK = Trim(.Fields("Rate").Value)

            Else
                MDUSD_LAK = 1
            End If
        End With

        'If MDUSD_LAK = 0 Then
        '    MDUSD_LAK = 8000
        'End If
    End Sub
    Public Sub RateSetting1()
        Dim Rs As New ADODB.Recordset
        With Rs
            Call LoadSqlData("select * from APListCurrency ORDER BY TimeUpdate DESC ", Rs)
            If .RecordCount > 0 Then
                StrDate = Trim(.Fields("TimeUpdate").Value)
                MDUSD_LAK1 = Trim(.Fields("Rate_LAK").Value)
                .MoveNext()
            Else
                '            Call LoadRs("select * from AP_Rate WHERE status=1", Rs)
                '            MDLAK = Trim(.Fields("LAK").Value)
                '            MDTHB = Trim(.Fields("THB").Value)
                '            MDUSD = Trim(.Fields("USD").Value)
                '            MDEUR = Trim(.Fields("EUR").Value)
                '            MDJPY = Trim(.Fields("JPY").Value)
                '            MDUSD_LAK = Trim(.Fields("USD_LAK").Value)
                '            MDTHB_LAK = Trim(.Fields("THB_LAK").Value)
                '            MDEUR_LAK = Trim(.Fields("EUR_LAK").Value)
                '            MDJPY_LAK = Trim(.Fields("JPY_LAK").Value)
                '            MDEUR_THB = Trim(.Fields("EUR_THB").Value)
                '            MDUSD_THB = Trim(.Fields("USD_THB").Value)
                '            MDJPY_THB = Trim(.Fields("JPY_THB").Value)
                '            MDEUR_USD = Trim(.Fields("EUR_USD").Value)
                '            MDJPY_THB = Trim(.Fields("JPY_USD").Value)
                '            MDRate_Curr = Trim(.Fields("Curr").Value)
                '            .MoveNext()
            End If
        End With
    End Sub
End Module
