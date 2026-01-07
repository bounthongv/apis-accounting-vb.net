Imports System.Data.SqlClient
Module MDOffice

    ' Using DbHelper functions for database operations
    Private Function GetDataTable(sql As String) As DataTable
        Return DbHelper.GetDataTable(sql)
    End Function

    Private Function ExecuteNonQuery(sql As String) As Integer
        Return DbHelper.ExecuteNonQuery(sql)
    End Function
    Public Usr, AssetID, ASName As String
    Public COM, IMageID, SqlPrint As String
    Public MDWrite, MDDelete As Integer
    Public MuOffNEW, MDRegister, MDOffName, MDOffAdd, MDOffTel, MDOffPlace, MDSgn1, MDSgn2, MDSgn3, MDSgn4, MDSgn5, MDSgn6, MDACC00 As String
    Public HeaDSec, OffName, OffNameE, Off_strtl, Off_VillageL, Off_DistL, Off_ProVL, OffAddress1, OffAddress2, OffTel, OffFax, OffDepartment, Print, Print1, Sign1, Sign2, Sign3, Sign4, Sign5, Sign6, Cal_Sys, OffPlace As String
    Public PlaecL, PlaecE, Sign1e, Sign2e, Sign3e, Sign4e, Sign5e As String
    Public MDHead, MDSignal1, MDSignal2, MDSignal3, MDSignal4, MDSignal5, MDNm, CallNM, MDDM, MDDetail As String
    Public MDPlace As String
    Public MDSql As String
    Public CURR01 As String
    Public MuOff2, MuOff, MuOffDep, RptSjOff, RptPro As String
    Public MULOGO As PictureBox
    Public MdCertifyAuto As String
    Public Snt As String
    Public MdShowLOGO As String
    Public MyOff, AccCD, MDStrID, MDStrNM As String
    Public MDCurr As String = ""
    Public MDRate_DT, COMPUTER_NM, S_code As String
    Public SS_Curr As String
    Public SaleID As String
    Public DTDATE02 As String
    Public MD_Curr As String
    Public Lang As Boolean = False
    Public MDACC_CR As String = ""
    Public Acc_Code, Acc_Code2, Acc_NmL, Acc_NmE, SCT, SCT_ID, LOT, LOT_ID, MDOther, MD_OffID, MD_Remark As String
    Public MD_Rate As Double
    Public MD_Rate2 As Double
    Public MD_Pay_Advince As Double
    Public MD_AMTCR As Double
    Public MD_AMTDR As Double
    Public MD_AMTCR_LAK As Double
    Public MD_AMTDR_LAK As Double
    Public MDACC_DR As String = ""
    Public MDAMT_CR As Double
    Public MDAMT_DR As Double
    Public MD_ACCODE As String = ""
    Public MD_FROM As String = ""
    Public MD_ACC_NameL As String
    Public MD_ACC_NameE As String
    Public MD_KH As String = ""

    'Public MD_Rate As Double
    Public mformat As Integer = 0
    Public Edit_Pro As Integer = 0
    Public EditActive As Boolean = False

    Public Sub LoadLoGO()
        If MdShowLOGO = 1 Then
            Try
                Dim sql As String = "Select * from Ap_Image where Img_Id = 'a' And ImgType='LOGO'"
                Dim dt As DataTable = GetDataTable(sql)
                If dt.Rows.Count = 0 Then 
                    MsgBox("ບໍ່ມີຂໍ້ມູນ") 
                    Exit Sub
                End If
            Catch ex As Exception
                MsgBox("Error loading logo: " & ex.Message)
                Exit Sub
            End Try
        End If
    End Sub

    Public Sub CHKVNK()
        Try
            Dim W As String = Format(CDate(MWorkSetting), "ddd")
            If W = "Wed" Then
                System.Diagnostics.Process.Start(My.Application.Info.DirectoryPath & "\CHKVNK.exe")
            End If
            If W = "Thu" Then
                ExecuteNonQuery("Update MbDtUse Set Snt='1'")
            End If
        Catch ex As Exception
            MsgBox("Error in CHKVNK: " & ex.Message)
        End Try
    End Sub
    Public Sub Office()
        'Call LoadLoGO()
        Try
            If MuLng = "L" Then
                Dim sql As String = "SELECT ShowLogo, place, CertifyAuto, off_name, DepNm, DepNme, tel, fax, Place, Off_Add1, Sign1, Sign2, Sign3, Sign4, Sign5, Sign6, MD00 FROM Ap_office WHERE Sub_id ='" & MuSubOff & "' ORDER BY Sub_id"
                Dim dt As DataTable = GetDataTable(sql)
                If dt.Rows.Count = 0 Then Exit Sub
                
                Dim row As DataRow = dt.Rows(0)
                MdCertifyAuto = CDbl(DbHelper.GetStr(row("CertifyAuto")))
                MdShowLOGO = CDbl(DbHelper.GetStr(row("ShowLogo")))
                MuOff = Trim(DbHelper.GetStr(row("off_name"))) & vbCrLf & "" & _
                 "" & Trim(DbHelper.GetStr(row("Place"))) & vbCrLf & "" & _
                 "" & Trim(DbHelper.GetStr(row("tel"))) & Trim(DbHelper.GetStr(row("fax"))) & ""
                MuOff = Trim(DbHelper.GetStr(row("off_name")))
                MuOffDep = Trim(DbHelper.GetStr(row("DepNm")))
                'RptPro = Trim(.Fields("Off_Add1").Value) & " ວັນທີ : ......../......../............  , "
                RptPro = " Date : ......../......../............  , "
                MDSgn1 = Trim(DbHelper.GetStr(row("Sign1")))
                MDSgn2 = Trim(DbHelper.GetStr(row("Sign2")))
                MDSgn3 = Trim(DbHelper.GetStr(row("Sign3")))
                MDSgn4 = Trim(DbHelper.GetStr(row("Sign4")))
                MDSgn5 = Trim(DbHelper.GetStr(row("Sign5")))
                MDSgn6 = Trim(DbHelper.GetStr(row("Sign6")))
                MDACC00 = Trim(DbHelper.GetStr(row("MD00")))
                MuOffNEW = Trim(DbHelper.GetStr(row("off_name")))
                MDRegister = Trim(DbHelper.GetStr(row("tel")))
                MDOffAdd = Trim(DbHelper.GetStr(row("place")))
                MuOff2 = Trim(DbHelper.GetStr(row("off_name"))) & vbCrLf & "" & _
                 "" & Trim(DbHelper.GetStr(row("Place"))) & vbCrLf & "" & _
                 "" & Trim(DbHelper.GetStr(row("tel"))) & ""
                MuOff = Trim(DbHelper.GetStr(row("off_name")))
                MuOffDep = Trim(DbHelper.GetStr(row("DepNm")))

                RptPro = " Date : ......../......../............  , "
                MDSgn1 = Trim(DbHelper.GetStr(row("Sign1")))
                MDSgn2 = Trim(DbHelper.GetStr(row("Sign2")))
                MDSgn3 = Trim(DbHelper.GetStr(row("Sign3")))
                MDSgn4 = Trim(DbHelper.GetStr(row("Sign4")))
                MDSgn5 = Trim(DbHelper.GetStr(row("Sign5")))
                MDSgn6 = Trim(DbHelper.GetStr(row("Sign6")))
                MDACC00 = Trim(DbHelper.GetStr(row("MD00")))
            Else
                Dim sqlE As String = "SELECT CertifyAuto, off_namee, DepNm, DepNme, tel, fax, Placee, Off_Adde1, Sign1e, Sign2e, Sign3e, Sign4e, Sign5e, Sign6e, MD00 FROM Ap_office WHERE Sub_id ='" & MuSubOff & "' ORDER BY Sub_id"
                Dim dtE As DataTable = GetDataTable(sqlE)
                If dtE.Rows.Count = 0 Then Exit Sub
                
                Dim rowE As DataRow = dtE.Rows(0)
                MdCertifyAuto = CDbl(DbHelper.GetStr(rowE("CertifyAuto")))
                MuOff = Trim(DbHelper.GetStr(rowE("off_namee"))) & vbCrLf & "" & _
                 "" & Trim(DbHelper.GetStr(rowE("Placee"))) & vbCrLf & "" & _
                 "" & Trim(DbHelper.GetStr(rowE("tel"))) & Trim(DbHelper.GetStr(rowE("fax"))) & ""
                MuOff = Trim(DbHelper.GetStr(rowE("off_namee")))
                MuOffDep = Trim(DbHelper.GetStr(rowE("DepNmE")))

                RptPro = Trim(DbHelper.GetStr(rowE("Off_Adde1"))) & " Date : ......../......../............  , "
                MDSgn1 = Trim(DbHelper.GetStr(rowE("Sign1e")))
                MDSgn2 = Trim(DbHelper.GetStr(rowE("Sign2e")))
                MDSgn3 = Trim(DbHelper.GetStr(rowE("Sign3e")))
                MDSgn4 = Trim(DbHelper.GetStr(rowE("Sign4e")))
                MDSgn5 = Trim(DbHelper.GetStr(rowE("Sign5e")))
                MDSgn6 = Trim(DbHelper.GetStr(rowE("Sign6e")))
                MDACC00 = Trim(DbHelper.GetStr(rowE("MD00")))
            End If

            RptSjOff = "N'" & MuOff2 & "' As RptSjOff , N'" & MuOffDep & "' As RptSjDep , N'" & RptPro & "' As RptPro  , N'" & MDSgn1 & "' As RptSign1  , N'" & MDSgn2 & "' As RptSign2  , N'" & MDSgn3 & "' As RptSign3   , N'" & MDSgn4 & "' As RptSign4  , N'" & MDSgn5 & "' As RptSign5  , N'" & MDSgn6 & "' As RptSign6  , "

            'LoadLoGO()
            'CHKVNK()
        Catch ex As Exception
            MsgBox("Error loading office data: " & ex.Message)
        End Try
    End Sub



    Public RPT_GRP As String
    Public RPT_GRPID As Double

    Public Off_Find, Off_Find2, MuTable As String

    Public Sub Find_Company()
        Dim OfUsr1 As String = Microsoft.VisualBasic.Left(Off_Find, 5)
        Dim OfUsr2 As String = Mid(Off_Find, 4, 2)
        Dim OfUsr3 As String = Microsoft.VisualBasic.Left(Off_Find, 2)
        If OfUsr1 = "00-00" Then
            MULook2 = ""
        Else
            If OfUsr2 = "00" Then
                MULook2 = "  And  Left(" & MuTable & "company,2)= '" & OfUsr3 & "' "
            Else
                MULook2 = "  And " & MuTable & "company= '" & OfUsr1 & "' "
            End If
        End If
    End Sub
End Module
