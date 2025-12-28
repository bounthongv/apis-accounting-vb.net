Module MDOffice
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
    Public RsLOGO As New ADODB.Recordset
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
            With RsLOGO
                If .State = ConnectionState.Open Then .Close()
                .Open(" Select * from  Ap_Image  where Img_Id = 'a' And ImgType='LOGO' ", CNN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                If .EOF Then MsgBox("ບໍ່ມີຂໍ້ມູນ") : Exit Sub
                If .EOF Then Exit Sub
            End With
        End If
    End Sub

    Public Sub CHKVNK()

        Dim W As String = Format(CDate(MWorkSetting), "ddd")
        If W = "Wed" Then
            System.Diagnostics.Process.Start(My.Application.Info.DirectoryPath & "\CHKVNK.exe")
        End If
        If W = "Thu" Then
            conn.Execute("Update MbDtUse Set Snt='1'")
        End If
    End Sub
    Public Sub Office()
        'Call LoadLoGO()
        Dim Rs As New ADODB.Recordset
        With Rs
            If MuLng = "L" Then
                Call LoadSqlData("SELECT ShowLogo, place , CertifyAuto ,  off_name , DepNm, DepNme, tel , fax , Place , Off_Add1 , Sign1 , Sign2 , Sign3  , Sign4 , Sign5, Sign6,MD00 " & _
                  "FROM Ap_office " & _
                  " WHERE Sub_id ='" & MuSubOff & "' ORDER BY Sub_id ", Rs)
                If .RecordCount = 0 Then Exit Sub
                MdCertifyAuto = CDbl((.Fields("CertifyAuto").Value))
                MdShowLOGO = CDbl((.Fields("ShowLogo").Value))
                MuOff = Trim(.Fields("off_name").Value) & vbCrLf & "" & _
                 "" & Trim(.Fields("Place").Value) & vbCrLf & "" & _
                "" & Trim(.Fields("tel").Value) & Trim(.Fields("fax").Value) & ""
                MuOff = Trim(.Fields("off_name").Value.ToString)
                MuOffDep = Trim(.Fields("DepNm").Value.ToString)
                'RptPro = Trim(.Fields("Off_Add1").Value) & " ວັນທີ : ......../......../............  , "
                RptPro = " ທີ່......................................., ວັນທີ : ......../......../............  , "
                MDSgn1 = Trim(.Fields("Sign1").Value.ToString)
                MDSgn2 = Trim(.Fields("Sign2").Value.ToString)
                MDSgn3 = Trim(.Fields("Sign3").Value.ToString)
                MDSgn4 = Trim(.Fields("Sign4").Value.ToString)
                MDSgn5 = Trim(.Fields("Sign5").Value.ToString)
                MDSgn6 = Trim(.Fields("Sign6").Value.ToString)
                MDACC00 = Trim(.Fields("MD00").Value.ToString)
                MuOffNEW = Trim(.Fields("off_name").Value.ToString)
                MDRegister = Trim(.Fields("tel").Value.ToString)
                MDOffAdd = Trim(.Fields("place").Value.ToString)
                MuOff2 = Trim(.Fields("off_name").Value) & vbCrLf & "" & _
                 "" & Trim(.Fields("Place").Value) & vbCrLf & "" & _
                "" & Trim(.Fields("tel").Value) & ""
            Else
                Call LoadSqlData("SELECT CertifyAuto ,  off_namee ,  DepNm, DepNme,  tel , fax , Placee , Off_Adde1, Sign1e , Sign2e , Sign3e , Sign4e , Sign5e , Sign6e ,MD00  " & _
                            "FROM Ap_office " & _
                            " WHERE Sub_id ='" & MuSubOff & "' ORDER BY Sub_id ", Rs)
                If .RecordCount = 0 Then Exit Sub
                MdCertifyAuto = CDbl((.Fields("CertifyAuto").Value))
                MuOff = Trim(.Fields("off_namee").Value) & vbCrLf & "" & _
                 "" & Trim(.Fields("Placee").Value) & vbCrLf & "" & _
                "" & Trim(.Fields("tel").Value) & Trim(.Fields("fax").Value) & ""
                MuOff = Trim(.Fields("off_namee").Value.ToString)
                MuOffDep = Trim(.Fields("DepNmE").Value.ToString)

                RptPro = Trim(.Fields("Off_Adde1").Value) & " Date : ......../......../............  , "
                MDSgn1 = Trim(.Fields("Sign1e").Value.ToString)
                MDSgn2 = Trim(.Fields("Sign2e").Value.ToString)
                MDSgn3 = Trim(.Fields("Sign3e").Value.ToString)
                MDSgn4 = Trim(.Fields("Sign4e").Value.ToString)
                MDSgn5 = Trim(.Fields("Sign5e").Value.ToString)
                MDSgn6 = Trim(.Fields("Sign6e").Value.ToString)
                MDACC00 = Trim(.Fields("MD00").Value.ToString)
            End If

        End With
        RptSjOff = "N'" & MuOff2 & "' As RptSjOff , N'" & MuOffDep & "' As RptSjDep , N'" & RptPro & "' As RptPro  , N'" & MDSgn1 & "' As RptSign1  , N'" & MDSgn2 & "' As RptSign2  , N'" & MDSgn3 & "' As RptSign3   , N'" & MDSgn4 & "' As RptSign4  , N'" & MDSgn5 & "' As RptSign5  , N'" & MDSgn6 & "' As RptSign6 , "

        'LoadLoGO()
        'CHKVNK()
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
