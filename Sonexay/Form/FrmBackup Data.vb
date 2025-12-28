Imports System.Data.SqlClient 
Imports System.Net
Imports System.IO
Public Class FrmBackup_Data
    Dim MyProgress As Double
    Dim l, k As Double
    '===================BackUpData==========================
    Dim con As SqlConnection
    Dim cmd As SqlCommand

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCanCal.Click
        Me.Close()

    End Sub

    Function Factorial(ByVal Value As Double) As Double
        If (Value = 0) Then
            Factorial = 1.0
            System.Threading.Thread.Sleep(3000)
        Else
            Factorial = Value * Factorial(Math.Ceiling(Value - 1))
        End If
    End Function

    Private Sub BtnBackUpData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBackUpData.Click
        '===================
    

        '======================



        PictureBox2.Location = New System.Drawing.Point(65, 37)
        My.Computer.FileSystem.CreateDirectory("C:\BackUp Data")
        ProgressBar1.Style = ProgressBarStyle.Continuous
        ProgressBar1.Step = 1
        Timer1.Enabled = True
        PictureBox2.Visible = False
        'PictureBox2.Visible = True
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ProgressBar1.Value = MyProgress
        ProgressBar2.Value = 1
        PictureBox2.Location = New System.Drawing.Point(60 + k, 37)

        'If k < 100 Then

        '    If k < 96 Then
        k = k + 7.5
        If k >= 135 Then

            PictureBox2.Visible = False

        End If
        If k > 0 Then
            If k <= 150 Then
                PictureBox2.Visible = True
            End If


        End If
        If k >= 150 Then
            k = 0
        End If

        If l < 100 Then

            If l < 96 Then
                l = l + 32

            End If


            If l >= 96 Then

                l = l + 2
            End If

        End If
        If l = 100 Then
            l = 0
        End If
        'If l < 100 Then
        '    l = l + 20

        '    If l > 80 Then
        '        l = 0
        '    End If
        'End If
        'If MyProgress < 100 Then MyProgress = MyProgress + 5

        If MyProgress < 100 Then

            If MyProgress < 20 Then
                MyProgress = MyProgress + 1

            End If
            If MyProgress >= 20 Then
                If MyProgress < 70 Then

                    MyProgress = MyProgress + 5

                End If

            End If

            If MyProgress >= 70 Then
                MyProgress = MyProgress + 0.5
            End If

        End If

        If MyProgress < 99 Then
            LbStatus.Text = "BackUp Data= " & MyProgress & "%"

        End If

        If MyProgress >= 99 Then
            '
            '
            PictureBox2.Visible = False
            'ProgressBar2.Value = 100
            ProgressBar1.Value = 100
            PictureBox2.Location = New System.Drawing.Point(134, 37)
            'Label4.Location = New System.Drawing.Point(57, 20)
            LbStatus.Text = "BackUp Data " & "(" & "Complete" & ")"
        End If
        'Dim s As String
        'Dim f As FileInfo
        's = SaveFileDialog1.FileName
        'f = New FileInfo(s)

        If MyProgress = 100 Then
            'con = New SqlConnection("Data Source= " & MDServerName & " ;Integrated Security=SSPI;Initial Catalog= " & MDDatabaName & " ")
            'cmd = New SqlCommand("backup database " & MDDatabaName & " to disk='" & txtSaveIn.Text & "\" & txtFileNane.Text & ".bak" & "'", con)


            con = New SqlConnection("Data Source= " & MDServerName & " ;Persist Security Info=true;User ID = " & MDServerUser & "; Password = " & MDServerPassword & "; Initial Catalog= " & MDDatabaName & "  ")

            cmd = New SqlCommand("backup database " & MDDatabaName & " to disk='" & txtSaveIn.Text & "\" & txtFileNane.Text & ".bak" & "'", con)
            ''cmd = New SqlCommand("backup database " & MDDatabaName & " to disk='" & txtSaveIn.Text & "\" & txtFileNane.Text & ".bak" & "' WITH DIFFERENTIAL  ", con)
            'cmd = New SqlCommand("backup database " & MDDatabaName & " to disk='C:\BackUp Data\" & txtFileNane.Text & ".bak" & "' WITH COPY_ONLY ", con)
            Try
                DownloadFTP("BCEL_TK.bak", "ftp://10.10.10.30/BCEL_TK.bak", "admin", "123")
            Catch ex As Exception

            End Try

            'DownloadFTP("ftp://10.10.10.44/BackUp Data/YYYY.bak", "C:\\YYYY.bak", "admin", "123")

            con.Open()
            cmd.ExecuteNonQuery()
            con.Close()
            Timer1.Enabled = False
            conn.Execute(" delete from Back_UpData where File_Nme='" & txtFileNane.Text & "'")
            conn.Execute("Insert inTo Back_UpData ( File_Nme ,  File_Date ) Values ('" & txtFileNane.Text & ".bak' , '" & Format(MWorkSetting, "dd/MM/yyyy") & "')")
            MsgBox("ການສຳຮອງຂໍ້ມູນຮຽບຮ້ອຍແລ້ວ")
            ProgressBar2.Value = 0
            ProgressBar1.Value = 0
            Me.Close()
        End If
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
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub FrmBackup_Database_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LbStatus.Text = "         Start"
        PictureBox2.Location = New System.Drawing.Point(134, 37)
        Dim a, k As String
        Dim b As Integer
        a = My.Application.Info.DirectoryPath
        k = Mid(a, CDbl(Len(a)) - 8, 3)
        'MsgBox(k)
        'If k = "bin" Then
        '    b = CDbl(Len(a)) - 10
        'Else
        '    b = CDbl(Len(a)) - 6
        'End If
        b = CDbl(Len(a)) - 6
        PictureBox2.Visible = True
        txtSaveIn.Text = Microsoft.VisualBasic.Left(a, b) & "\BackUp Data"
        ProgressBar1.Value = 0
        MyProgress = 0
        txtFileNane.Text = MDDatabaName & " BackUp" & "-" & Format(CDate(MWorkSetting), "dd-MM-yyyy")
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBrowser.Click
        Dim dl As New FolderBrowserDialog
        dl.ShowDialog()
        txtSaveIn.Text = dl.SelectedPath
        ProgressBar1.Value = 0
        MyProgress = 0
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class