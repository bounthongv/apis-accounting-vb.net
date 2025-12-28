Imports System.Data
Imports System.Windows.Forms
Imports System.Collections.Generic

' Note: The designer file for this form appears to be empty. 
' You may need to reconstruct the UI controls in Visual Studio.
Public Class FmJeneralJournal_List
    
    ' Method to fetch 'wait' records, insert to local DB, and update status to 'success'
    Public Sub LoadAndProcessData()
        Try
            ' 1. Fetch data from API (Status = 'wait')
            Dim dtMsp As DataTable = ApiClient.GetMspData("wait")
            
            If dtMsp Is Nothing OrElse dtMsp.Rows.Count = 0 Then
                MessageBox.Show("No pending transactions found (status='wait').", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            Dim successCount As Integer = 0
            Dim failCount As Integer = 0

            ' Ensure Local DB Connection (using global CNN from MDSQLConnection)
            ' If CNN is not open, you might need to call MDSQLConnection.ConnectSQL()
            ' If CNN.State <> ConnectionState.Open Then Call ConnectSQL()

            For Each row As DataRow In dtMsp.Rows
                Dim trnId As String = row("trn_id").ToString()
                Dim isLocalInsertSuccess As Boolean = False
                
                Try
                    ' --- Transaction Start (Optional but recommended) ---
                    ' CNN.BeginTrans() 

                    ' A. Insert Header (MSP)
                    ' Note: Adjust table names (STAGE_MSP vs msp) as per your local DB schema
                    Dim sqlMsp As String = "INSERT INTO STAGE_MSP (trn_id, trn_desc, currency, acc_book, status, bis_date, create_date, ex_rate) VALUES (" & _
                        "'" & trnId & "', " & _
                        "N'" & CleanSql(row("trn_desc")) & "', " & _
                        "'" & CleanSql(row("currency")) & "', " & _
                        "'" & CleanSql(row("acc_book")) & "', " & _
                        "'" & CleanSql(row("status")) & "', " & _
                        "'" & FormatDate(row("bis_date")) & "', " & _
                        "'" & FormatDate(row("create_date")) & "', " & _
                        Val(row("ex_rate").ToString()) & ")"
                    
                    CNN.Execute(sqlMsp)
                    
                    ' B. Insert Debits
                    Dim dtDr As DataTable = ApiClient.GetDetails("retrieve_dr_trn_id", trnId)
                    If dtDr IsNot Nothing Then
                        For Each dr As DataRow In dtDr.Rows
                            Dim sqlDr As String = "INSERT INTO STAGE_TBL_DR (trn_id, dr_ac, dr_amt, dr_amt_lak, dr_desc) VALUES (" & _
                                "'" & trnId & "', " & _
                                "'" & CleanSql(dr("dr_ac")) & "', " & _
                                Val(dr("dr_amt").ToString()) & ", " & _
                                Val(dr("dr_amt_lak").ToString()) & ", " & _
                                "N'" & CleanSql(dr("dr_desc")) & "')"
                            CNN.Execute(sqlDr)
                        Next
                    End If
                    
                    ' C. Insert Credits
                    Dim dtCr As DataTable = ApiClient.GetDetails("retrieve_cr_trn_id", trnId)
                    If dtCr IsNot Nothing Then
                        For Each cr As DataRow In dtCr.Rows
                             Dim sqlCr As String = "INSERT INTO STAGE_TBL_CR (trn_id, cr_ac, cr_amt, cr_amt_lak, cr_desc) VALUES (" & _
                                "'" & trnId & "', " & _
                                "'" & CleanSql(cr("cr_ac")) & "', " & _
                                Val(cr("cr_amt").ToString()) & ", " & _
                                Val(cr("cr_amt_lak").ToString()) & ", " & _
                                "N'" & CleanSql(cr("cr_desc")) & "')"
                            CNN.Execute(sqlCr)
                        Next
                    End If
                    
                    ' CNN.CommitTrans()
                    isLocalInsertSuccess = True

                Catch ex As Exception
                    ' CNN.RollbackTrans()
                    Console.WriteLine("Local DB Error for " & trnId & ": " & ex.Message)
                    failCount += 1
                End Try

                ' 3. Update Status to 'success' via API
                If isLocalInsertSuccess Then
                    Dim apiUpdated As Boolean = ApiClient.UpdateStatus(trnId, "success")
                    If apiUpdated Then
                        successCount += 1
                        Console.WriteLine("Successfully processed: " & trnId)
                    Else
                        Console.WriteLine("Local insert success, but API update failed for: " & trnId)
                        ' Optionally: Log this to a retry queue
                    End If
                End If
            Next
            
            MessageBox.Show("Process Completed." & vbCrLf & _
                            "Success: " & successCount & vbCrLf & _
                            "Failed: " & failCount, "Report", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Critical Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Helper to clean strings for SQL injection prevention (basic)
    Private Function CleanSql(ByVal val As Object) As String
        If IsDBNull(val) OrElse val Is Nothing Then Return ""
        Return val.ToString().Replace("'", "''")
    End Function

    ' Helper to format dates for SQL Server
    Private Function FormatDate(ByVal val As Object) As String
        If IsDBNull(val) OrElse val Is Nothing Then Return ""
        Try
            Return CDate(val).ToString("yyyy-MM-dd HH:mm:ss")
        Catch
            Return ""
        End Try
    End Function

End Class
