Imports System.Data.SqlClient
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

        Try
            Dim dt As DataTable = DbHelper.GetDataTable("select top 1  * from AP_Rate_history where 1=1  " & MDRate_DT & " " & SS_Curr & " " & MDR_Curr & "   ORDER BY rate_dt DESC ")
            If dt.Rows.Count > 0 Then
                Dim row As DataRow = dt.Rows(0)
                StrDate = Trim(row("rate_dt").ToString())
                MD_Rate = Convert.ToDouble(row("Rate"))
                MD_Rate2 = Convert.ToDouble(row("Rate2"))
                MDLAK = Convert.ToDouble(row("Rate"))
                MDTHB = Convert.ToDouble(row("Rate"))
                MDUSD = Convert.ToDouble(row("Rate"))
                MDEUR = Convert.ToDouble(row("Rate"))
                MDUSD_LAK = Convert.ToDouble(row("Rate"))
                MDTHB_LAK = Convert.ToDouble(row("Rate"))
                MDEUR_LAK = Convert.ToDouble(row("Rate"))
                MDEUR_THB = Convert.ToDouble(row("Rate"))
                MDUSD_THB = Convert.ToDouble(row("Rate"))
                MDEUR_USD = Convert.ToDouble(row("Rate"))
                MDRate_Curr = Trim(row("Curr").ToString())

            Else
                MD_Rate = 1
                Dim dt2 As DataTable = DbHelper.GetDataTable("select * from AP_Rate_history where com_id='00'  ORDER BY rate_dt DESC ")
                If dt2.Rows.Count > 0 Then
                    Dim row As DataRow = dt2.Rows(0)
                    MD_Rate = Convert.ToDouble(row("Rate"))
                    MD_Rate2 = Convert.ToDouble(row("Rate2"))
                    MDLAK = Convert.ToDouble(row("Rate"))
                    MDTHB = Convert.ToDouble(row("Rate"))
                    MDUSD = Convert.ToDouble(row("Rate"))
                    MDEUR = Convert.ToDouble(row("Rate"))
                    MDUSD_LAK = Convert.ToDouble(row("Rate"))
                    MDTHB_LAK = Convert.ToDouble(row("Rate"))
                    MDEUR_LAK = Convert.ToDouble(row("Rate"))
                    MDEUR_THB = Convert.ToDouble(row("Rate"))
                    MDUSD_THB = Convert.ToDouble(row("Rate"))
                    MDEUR_USD = Convert.ToDouble(row("Rate"))
                    MDRate_Curr = Trim(row("Curr").ToString())
                End If
            End If
        Catch ex As Exception
            VSysError = True
            MessageBox.Show("Database Error in RateSetting: " & ex.Message, "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        '==============================================

        Dim dtUSD As DataTable = DbHelper.GetDataTable("select top 1  * from AP_Rate_history where 1=1  " & MDRate_DT & "  and curr='USD' ORDER BY rate_dt DESC ")
        If dtUSD.Rows.Count > 0 Then
            MDUSD_LAK = Convert.ToDouble(Trim(dtUSD.Rows(0)("Rate").ToString()))

        Else
            MDUSD_LAK = 1
        End If

        'If MDUSD_LAK = 0 Then
        '    MDUSD_LAK = 8000
        'End If
    End Sub
    Public Sub RateSetting1()
        Try
            Dim dt As DataTable = DbHelper.GetDataTable("select * from APListCurrency ORDER BY TimeUpdate DESC ")
            If dt.Rows.Count > 0 Then
                Dim row As DataRow = dt.Rows(0)
                StrDate = Trim(row("TimeUpdate").ToString())
                MDUSD_LAK1 = Convert.ToDouble(row("Rate_LAK"))
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
        Catch ex As Exception
            VSysError = True
            MessageBox.Show("Database Error in RateSetting1: " & ex.Message, "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Module
