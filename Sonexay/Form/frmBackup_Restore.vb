Imports System.Data.SqlClient
Imports System.Net
Imports System.IO
Public Class frmBackup_Restore
    Dim con, con1 As SqlConnection
    Dim cmd As SqlCommand
    Dim dread As SqlDataReader

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        server(MDServerName)
        ' server("apisserver\sqlexpress")
    End Sub
    Sub server(ByVal str As String)
        Try
            con = New SqlConnection("Server=" & str & ";Database=BCEL_TK;User ID=" & MDServerUser & ";Password= " & MDServerPassword & " ")
            con.Open()
            cmd = New SqlCommand("select *  from sysservers  where srvproduct='SQL Server'", con)
            dread = cmd.ExecuteReader
            While dread.Read
                cmbserver.Items.Add(dread(2))
            End While
            dread.Close()
        Catch ex As Exception
            '    MessageBox.Show(ex.Message)
        End Try

    End Sub
    Sub connection()
        Try
            con = New SqlConnection("SERVER=" & Trim(cmbserver.Text) & ";Database=BCEL_TK;User ID=" & MDServerUser & ";Password= " & MDServerPassword & "")
            con.Open()
            cmbdatabase.Items.Clear()
            cmd = New SqlCommand("select * from sysdatabases", con)
            dread = cmd.ExecuteReader
            While dread.Read
                cmbdatabase.Items.Add(dread(0))
            End While
            dread.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmbserver_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbserver.SelectedIndexChanged
        connection()
    End Sub
    Sub query(ByVal que As String)
        On Error Resume Next
        cmd = New SqlCommand(que, con)
        cmd.ExecuteNonQuery()
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If ProgressBar1.Value = 100 Then
            Timer1.Enabled = False
            ProgressBar1.Visible = False
            MsgBox("Successfully Done")
        Else
            ProgressBar1.Value = ProgressBar1.Value + 5
        End If
    End Sub
    Sub blank(ByVal str As String)
        Try
            Dim cnstr As String = "SERVER=" & Trim(cmbserver.Text) & ";Database=BCEL_TK;User ID=" & MDServerUser & ";Password= " & MDServerPassword & " "
            If cmbserver.Text = "" Or cmbdatabase.Text = "" Then
                MsgBox("Server Name & Database Blank Field")
                Exit Sub
            Else
                If str = "backup" Then
                    SaveFileDialog1.FileName = cmbdatabase.Text
                    SaveFileDialog1.ShowDialog()
                    Timer1.Enabled = True
                    ProgressBar1.Visible = True
                    Dim s As String
                    Dim f As FileInfo
                    s = SaveFileDialog1.FileName
                    f = New FileInfo(s)
                    Dim dd As String = Microsoft.VisualBasic.Right(Name, 4)
                    SQLBACKUP(cnstr, cmbdatabase.Text, "D:\Backup data\" & TextBox1.Text & "." & f.Name)
                    'DownloadFTP("ftp://192.168.1.3/SQLBACKUP/" & f.Name, s, "psm", "psm123")
                ElseIf str = "restore" Then
                    OpenFileDialog1.ShowDialog()
                    Timer1.Enabled = True
                    ProgressBar1.Visible = True
                    Dim f As FileInfo
                    f = New FileInfo(OpenFileDialog1.FileName)
                    'UPLOADFTP("ftp://192.168.1.3/SQLBACKUP/" & f.Name, OpenFileDialog1.FileName, "psm", "psm123")
                    SQLRESTORE(cnstr, cmbdatabase.Text, "D:\Backup data\" & f.Name)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub cmbbackup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbbackup.Click
        blank("backup")
    End Sub

    Private Sub cmdrestore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdrestore.Click
        blank("restore")
    End Sub

    Private Sub SQLBACKUP(ByVal cnstr As String, ByVal DBName As String, ByVal BKFile As String)
        Try
            Dim sqlConnectionString As String = cnstr
            Dim conn As New SqlConnection(sqlConnectionString)
            conn.Open()
            Dim cmd As New SqlCommand
            With cmd
                .CommandType = CommandType.Text
                .CommandText = "BACKUP DATABASE " & DBName & " TO DISK=N'" & BKFile & "' WITH COPY_ONLY"
                .Connection = conn
                .ExecuteNonQuery()

            End With
            conn.Close()
            '  MessageBox.Show("Successfull")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub SQLRESTORE(ByVal cnstr As String, ByVal DBName As String, ByVal BKFile As String)
        Try
            Dim sqlConnectionString As String = cnstr
            Dim conn As New SqlConnection(sqlConnectionString)
            conn.Open()
            Dim cmd As New SqlCommand
            With cmd
                .CommandType = CommandType.Text
                .CommandText = "RESTORE DATABASE " & DBName & " FROM DISK = N'" & BKFile & "' WITH REPLACE"
                .Connection = conn
                .ExecuteNonQuery()

            End With
            conn.Close()
            '  MessageBox.Show("Successfull")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub UPLOADFTP(ByVal RFN As String, ByVal LFN As String, ByVal ui As String, ByVal pw As String)
        Try
            Dim clsRequest As System.Net.FtpWebRequest = DirectCast(System.Net.WebRequest.Create(RFN), System.Net.FtpWebRequest)
            clsRequest.Credentials = New System.Net.NetworkCredential(ui, pw)
            clsRequest.Method = System.Net.WebRequestMethods.Ftp.UploadFile

            Dim bFile() As Byte = System.IO.File.ReadAllBytes(LFN)
            Dim clsStream As System.IO.Stream = _
                clsRequest.GetRequestStream()
            clsStream.Write(bFile, 0, bFile.Length)
            clsStream.Close()
            clsStream.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub DeleteFTP(ByVal SFN As String, ByVal ui As String, ByVal pw As String)
        Try
            Dim FTPDelReq As FtpWebRequest = WebRequest.Create(SFN)
            FTPDelReq.Credentials = New Net.NetworkCredential(ui, pw)
            FTPDelReq.Method = WebRequestMethods.Ftp.DeleteFile
            Dim FTPDelResp As FtpWebResponse = FTPDelReq.GetResponse
        Catch ex As Exception

        End Try
    End Sub
    Private Sub DownloadFTP(ByVal RFN As String, ByVal LFN As String, ByVal ui As String, ByVal pw As String)
        Try
            Dim ftp As FtpWebRequest = CType(FtpWebRequest.Create(RFN), FtpWebRequest)
            '   Console.WriteLine("Downloading: " & RFN)
            ftp.Credentials = New NetworkCredential(ui, pw)
            ftp.KeepAlive = False
            ftp.UseBinary = True
            ftp.Method = WebRequestMethods.Ftp.DownloadFile
            Using FtpResponse As FtpWebResponse = CType(ftp.GetResponse, FtpWebResponse)
                Using ResponseStream As IO.Stream = FtpResponse.GetResponseStream

                    Using fs As New IO.FileStream(LFN, FileMode.Create)
                        Dim buffer(2047) As Byte
                        Dim read As Integer = 0
                        Do
                            read = ResponseStream.Read(buffer, 0, buffer.Length)
                            fs.Write(buffer, 0, read)
                            '    Console.Write(".")
                        Loop Until read = 0
                        ResponseStream.Close()
                        fs.Flush()
                        fs.Close()
                        '   Log("")
                    End Using
                    ResponseStream.Close()
                End Using
            End Using
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub cmbdatabase_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbdatabase.SelectedIndexChanged

    End Sub
End Class
