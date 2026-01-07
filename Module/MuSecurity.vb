Imports System.Data.SqlClient
Module MuSecurity

    ' Using DbHelper functions for database operations
    Private Function GetDataTable(sql As String) As DataTable
        Return DbHelper.GetDataTable(sql)
    End Function

    Private Function ExecuteNonQuery(sql As String) As Integer
        Return DbHelper.ExecuteNonQuery(sql)
    End Function
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
        Try
            Dim srNum As DataTable = GetDataTable("SELECT top 1 ScSaving FROM AgUng Order by ScSaving DESC")
            If srNum.Rows.Count > 0 Then
                ScSaving = Val(srNum.Rows(0)("ScSaving").ToString())
            End If
            ExecuteNonQuery("update AgUng set ScSaving =" & CDbl(ScSaving) + 1 & "")
            'LoadCheckRecor()
            'ExecuteNonQuery("update AgUng set ScRecordUsing =" & CDbl(ScRecordUsing) & "")
        Catch ex As Exception
            MsgBox("Error in LoadAtoSaveRecor: " & ex.Message)
        End Try
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
        Try
            Dim rs As DataTable = GetDataTable("SELECT lock FROM gen_jn ")
            For Each row As DataRow In rs.Rows
                ScRecordUsing = CDbl(ScRecordUsing) + 1
            Next
        Catch ex As Exception
            MsgBox("Error in LoadCheckRecor: " & ex.Message)
        End Try
    End Sub
    Public Sub LockProgrome()
        ExecuteNonQuery("update AgUng set Sclock ='1'")
    End Sub
    Public Sub UpdateExS()
        ExecuteNonQuery("update AgUng set ScExS ='1'")
    End Sub
    Public Sub UpdateRemoveExS()
        ExecuteNonQuery("update AgUng set ScExS ='0'")
    End Sub
    Public Sub LoadCheLock()
        Try
            Dim srNum As DataTable = GetDataTable("SELECT ScLock FROM AgUng ")
            If srNum.Rows.Count > 0 Then
                ScLock = Val(srNum.Rows(0)("ScLock").ToString())
            End If
        Catch ex As Exception
            MsgBox("Error in LoadCheLock: " & ex.Message)
        End Try
    End Sub
    Public Sub LoadScExS()
        Try
            Dim srNum As DataTable = GetDataTable("SELECT ScExS FROM AgUng ")
            If srNum.Rows.Count > 0 Then
                CloseAll = Val(srNum.Rows(0)("ScExS").ToString())
            End If
        Catch ex As Exception
            MsgBox("Error in LoadScExS: " & ex.Message)
        End Try
    End Sub
    Public Sub LoadChecPermitSave()
        Try
            Dim srNum As DataTable = GetDataTable("SELECT ScPermitSave FROM AgUng ")
            If srNum.Rows.Count > 0 Then
                ScPermitSave = Val(srNum.Rows(0)("ScPermitSave").ToString())
            End If
        Catch ex As Exception
            MsgBox("Error in LoadChecPermitSave: " & ex.Message)
        End Try
    End Sub
    Public Sub LoadChecScSaving()
        Try
            Dim srNum As DataTable = GetDataTable("SELECT ScSaving FROM AgUng ")
            If srNum.Rows.Count > 0 Then
                ScSaving = Val(srNum.Rows(0)("ScSaving").ToString())
            End If
        Catch ex As Exception
            MsgBox("Error in LoadChecScSaving: " & ex.Message)
        End Try
    End Sub
    Public Sub LoadChecOwner()
        Try
            Dim srNum As DataTable = GetDataTable("SELECT ScOwner FROM AgUng ")
            If srNum.Rows.Count > 0 Then
                ScOwner = Val(srNum.Rows(0)("ScOwner").ToString())
            End If
            'MsgBox(ScOwner)
        Catch ex As Exception
            MsgBox("Error in LoadChecOwner: " & ex.Message)
        End Try
    End Sub
    Public Sub LoadPermitRecord()
        Try
            Dim srNum As DataTable = GetDataTable("SELECT ScPermitRecord FROM AgUng ")
            If srNum.Rows.Count > 0 Then
                ScPermitRecord = Val(srNum.Rows(0)("ScPermitRecord").ToString())
            End If
            'MsgBox(ScPermitRecord)
        Catch ex As Exception
            MsgBox("Error in LoadPermitRecord: " & ex.Message)
        End Try
    End Sub
End Module
