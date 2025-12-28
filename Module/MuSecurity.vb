Module MuSecurity
    Public ScOwner As String
    Public ScPermitRecord As String
    Public ScPermitSave As String
    Public ScRecordUsing As String = 0
    Public ScSaving As Double
    Public ScLock, ScLock2 As String
    Public CloseAll As String
    Public Axt As String

    Public MULook, MULook2, MULookSelct As String




    Public Sub LoadAtoSaveRecor()
        Dim srNum As New ADODB.Recordset
        Call LoadSqlData("SELECT top 1 ScSaving FROM AgUng Order by ScSaving DESC", srNum)
        ScSaving = Val(srNum.Fields("ScSaving").Value.ToString)
        CNN.Execute("update AgUng set ScSaving =" & CDbl(ScSaving) + 1 & "")
        'LoadCheckRecor()
        'CNN.Execute("update AgUng set ScRecordUsing =" & CDbl(ScRecordUsing) & "")
    End Sub
 
    Public Sub Loadfind()
    
        MULook = ""
        If Mid(FmLogin.Sub_Company.Text, 1, 5) = "00-00" Then
            MULook = ""

        Else
            If Mid(FmLogin.Sub_Company.Text, 4, 2) = "00" Then
                MULook = "AND company  Like N'" & Mid(FmLogin.Sub_Company.Text, 1, 2) & "%'"
            Else
                MULook = "AND company = '" & Mid(FmLogin.Sub_Company.Text, 1, 5) & "' "
            End If

        End If

        If MPermit = "User" Then
            MULook = " AND Last_User = N'" & MUserName & "' "
        End If
    
    End Sub
    Public Sub LoadfindLogOut()


        MULook = ""
        If Mid(FmLogOut.Sub_Company.Text, 1, 5) = "00-00" Then
            MULook = ""

        Else

            If Mid(FmLogOut.Sub_Company.Text, 4, 2) = "00" Then
                MULook = "AND company  Like N'" & Mid(FmLogOut.Sub_Company.Text, 1, 2) & "%'"
            Else
                MULook = "AND company = '" & Mid(FmLogOut.Sub_Company.Text, 1, 5) & "' "
            End If

        End If

        If MPermit = "User" Then
            MULook = " AND Last_User = N'" & MUserName & "' "
        End If

    End Sub

    Public Sub LoadCheckRecor()
        Call LoadSqlData("SELECT lock FROM gen_jn ", RSC)
        With RSC
            Do Until .EOF = True
                ScRecordUsing = CDbl(ScRecordUsing) + 1
                .MoveNext()
            Loop
        End With
    End Sub
    Public Sub LockProgrome()
        CNN.Execute("update AgUng set Sclock ='1'")
    End Sub
    Public Sub UpdateExS()
        CNN.Execute("update AgUng set ScExS ='1'")
    End Sub
    Public Sub UpdateRemoveExS()
        CNN.Execute("update AgUng set ScExS ='0'")
    End Sub
    Public Sub LoadCheLock()
        Dim srNum As New ADODB.Recordset
        Call LoadSqlData("SELECT  ScLock FROM AgUng ", srNum)
        ScLock = Val(srNum.Fields("ScLock").Value.ToString)

    End Sub
    Public Sub LoadScExS()
        Dim srNum As New ADODB.Recordset
        Call LoadSqlData("SELECT  ScExS FROM AgUng ", srNum)
        CloseAll = Val(srNum.Fields("ScExS").Value.ToString)

    End Sub
    Public Sub LoadChecPermitSave()
        Dim srNum As New ADODB.Recordset
        Call LoadSqlData("SELECT  ScPermitSave FROM AgUng ", srNum)
        ScPermitSave = Val(srNum.Fields("ScPermitSave").Value.ToString)

    End Sub
    Public Sub LoadChecScSaving()
        Dim srNum As New ADODB.Recordset
        Call LoadSqlData("SELECT  ScSaving FROM AgUng ", srNum)
        ScSaving = Val(srNum.Fields("ScSaving").Value.ToString)

    End Sub
    Public Sub LoadChecOwner()
        Dim srNum As New ADODB.Recordset
        Call LoadSqlData("SELECT  ScOwner FROM AgUng ", srNum)
        ScOwner = Val(srNum.Fields("ScOwner").Value.ToString)
        'MsgBox(ScOwner)
    End Sub
    Public Sub LoadPermitRecord()
        Dim srNum As New ADODB.Recordset
        Call LoadSqlData("SELECT  ScPermitRecord FROM AgUng ", srNum)
        ScPermitRecord = Val(srNum.Fields("ScPermitRecord").Value.ToString)
        'MsgBox(ScPermitRecord)
    End Sub
End Module
