Imports System.Net
Imports System.Data
Imports System.Data.SqlClient
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ApiClient
    ' Configure your server details here
    Private Shared ReadOnly Property BaseUrl As String
        Get
            Return "http://" & MDServerName2 & ":8000/apis"
        End Get
    End Property
    Private Const ApiToken As String = "8c57a7c3dfe7307abf40c9e35d0508ba6d2e2c4dda27ae66567627b0da5d68ae"

    Private Shared Function CreateClient() As WebClient
        Dim client As New WebClient()
        ' Bypass system proxy to ensure connection to local/internal IP
        client.Proxy = Nothing
        client.Encoding = System.Text.Encoding.UTF8
        client.Headers.Add("Authorization", "Bearer " & ApiToken)
        client.Headers.Add("Content-Type", "application/json")
        Return client
    End Function

    Public Shared Function GetMspData(ByVal status As String) As DataTable
        ' Validation for Server Name
        If String.IsNullOrEmpty(MDServerName2) Then
            MsgBox("Error: MDServerName2 is not set. Please check your configuration.")
            Return Nothing
        End If

        Dim url As String = BaseUrl & "/retrieve_msp_status?status=" & status
        Console.WriteLine("DEBUG: Requesting URL: " & url)

        Using client As WebClient = CreateClient()
            Try
                Dim jsonResponse As String = client.DownloadString(url)
                Console.WriteLine("DEBUG: Response received: " & jsonResponse)

                If String.IsNullOrEmpty(jsonResponse) Then
                    MsgBox("Error: API returned empty response")
                    Return Nothing
                End If

                Dim rootObject As JObject = JObject.Parse(jsonResponse)

                ' Flexible check for "200"
                Dim respCode As String = ""
                If rootObject("code") IsNot Nothing Then respCode = rootObject("code").ToString().Trim()

                ' Extract data
                Dim dataToken As JToken = rootObject("data")
                If dataToken Is Nothing OrElse Not dataToken.HasValues Then
                    Return New DataTable() ' Return empty table if no records
                End If

                Dim dataRows As JArray = CType(dataToken, JArray)

                ' Deserialize to DataTable
                Return JsonConvert.DeserializeObject(Of DataTable)(dataRows.ToString())

            Catch ex As System.Net.WebException
                ' Capture server error
                Dim serverError As String = ""
                If ex.Response IsNot Nothing Then
                    Using reader As New System.IO.StreamReader(ex.Response.GetResponseStream())
                        serverError = reader.ReadToEnd()
                    End Using
                End If
                Dim errCommon As String = "WebException: " & ex.Message & vbCrLf & "Server Output: " & serverError
                Console.WriteLine(errCommon)
                MsgBox(errCommon)
                Return Nothing

            Catch ex As Exception
                ' Catch JSON parsing or other errors
                Dim errCommon As String = "General Error: " & ex.Message & vbCrLf & "Stack Trace: " & ex.StackTrace
                Console.WriteLine(errCommon)
                MsgBox(errCommon)
                Return Nothing
            End Try
        End Using
    End Function

    Public Shared Function GetDetails(ByVal endpoint As String, ByVal trnId As String) As DataTable
        Dim url As String = BaseUrl & "/" & endpoint & "?trn_id=" & trnId
        Using client As WebClient = CreateClient()
            Try
                Dim jsonResponse As String = client.DownloadString(url)

                If String.IsNullOrEmpty(jsonResponse) Then
                    MsgBox("Error: API returned empty response")
                    Return Nothing
                End If

                Dim rootObject As JObject = JObject.Parse(jsonResponse)

                ' Flexible check for "200"
                Dim respCode As String = ""
                If rootObject("code") IsNot Nothing Then respCode = rootObject("code").ToString().Trim()

                'If respCode <> "200" Then
                '    Dim msg As String = "Unknown Error"
                '    If rootObject("message") IsNot Nothing Then msg = rootObject("message").ToString()
                '    MsgBox("API Error (Code " & respCode & "): " & msg)
                '    Return Nothing
                'End If

                ' Extract data
                Dim dataToken As JToken = rootObject("data")
                If dataToken Is Nothing OrElse Not dataToken.HasValues Then
                    Return New DataTable() ' Return empty table if no records
                End If

                Dim dataRows As JArray = CType(dataToken, JArray)

                ' Deserialize to DataTable
                Return JsonConvert.DeserializeObject(Of DataTable)(dataRows.ToString())

            Catch ex As System.Net.WebException
                If ex.Response IsNot Nothing Then
                    Using reader As New System.IO.StreamReader(ex.Response.GetResponseStream())
                        Console.WriteLine("SERVER ERROR: " & reader.ReadToEnd())
                    End Using
                End If
                Console.WriteLine("WebException: " & ex.Message)
                Return Nothing
            Catch ex As Exception
                Console.WriteLine("General Error: " & ex.Message)
                Return Nothing
            End Try
        End Using
    End Function

    Public Shared Function UpdateStatus(ByVal trnId As String, ByVal status As String, Optional ByVal failReason As String = "") As Boolean
        Dim url As String = BaseUrl & "/update_status"
        Using client As WebClient = CreateClient()
            Try
                ' Create the payload
                Dim payload As New JObject()
                payload("trn_id") = trnId
                payload("status") = status
                If Not String.IsNullOrEmpty(failReason) Then
                    payload("fail_reason") = failReason
                End If

                ' PATCH is supported by UploadString by specifying the method parameter
                Dim response As String = client.UploadString(url, "PATCH", payload.ToString())

                Dim rootObject As JObject = JObject.Parse(response)
                Return rootObject("code").ToString() = "200"

            Catch ex As WebException
                Dim serverError As String = ""
                If ex.Response IsNot Nothing Then
                    Using reader As New System.IO.StreamReader(ex.Response.GetResponseStream())
                        serverError = reader.ReadToEnd()
                    End Using
                End If
                Console.WriteLine("UpdateStatus Error: " & ex.Message & vbCrLf & "Server Output: " & serverError)
                Return False
            Catch ex As Exception
                Console.WriteLine("UpdateStatus General Error: " & ex.Message)
                Return False
            End Try
        End Using
    End Function
End Class