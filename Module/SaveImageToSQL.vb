Imports System.IO
Module SaveImageToSQL
    Public SUPD As Integer = 0
    Public ImageSlno As Integer
    Public con As New OleDb.OleDbConnection("Provider=SQLOLEDB;User id=" & MDServerUser & ";database=" & MDDatabaName & ";password=" & MDServerPassword & ";data source=" & MDServerName & "")
    Public b_x, b_y, g_x, g_y As Integer
    Public Sub LoadImgSize()
        LoadSqlData("SELECT * FROM Ap_SizeImg ", RSC)
        With RSC
            Do Until .EOF = True
                b_x = Trim(.Fields("b_x").Value)
                b_y = Trim(.Fields("b_y").Value)
                g_x = Trim(.Fields("g_x").Value)
                g_y = Trim(.Fields("g_y").Value)
                .MoveNext()
            Loop
        End With
    End Sub




    Public Sub LoadPhoto()
        Try
            Dim str As String = "SELECT Img FROM Ap_Image WHERE Img_Id = '" & Fm_Image.Img_ID.Text & "' And  ImgType = '" & Fm_Image.ImgType.Text & "' "
            con.Open()
            Dim cmd As New OleDb.OleDbCommand(str, con)
            Dim b() As Byte
            b = cmd.ExecuteScalar()
            con.Close()
            If (b.Length > 0) Then
                Dim stream As New MemoryStream(b, True)
                stream.Write(b, 0, b.Length)
                DrawToScale(New Bitmap(stream))
                stream.Close()
            End If
        Catch ex As Exception
            Fm_Image.PictureBox1.Image = Fm_Image.a123456789.Image
        End Try
    End Sub
    Private Sub DrawToScale(ByVal bmp As Image)
        Fm_Image.PictureBox1.Image = New Bitmap(bmp)
    End Sub

    Public Sub deleteImage()
        CNN.Execute("delete Ap_Image  WHERE Img_Id = '" & Fm_Image.Img_ID.Text & "' And  ImgType = '" & Fm_Image.ImgType.Text & "'")
    End Sub
    Public Sub Insert_Image2()
        'If (Fm_Image.PictureBox1.Image Is Nothing) Then
        '    MsgBox("No Image Is There ")
        ''    Exit Sub
        'End If
        'Try
        ''Dim st As New FileStream(Fm_Image.OpenFileDialog1.FileName, FileMode.Open, FileAccess.Read)
        'Dim mbr As BinaryReader = New BinaryReader(st)
        'Dim buffer(st.Length) As Byte
        'mbr.Read(buffer, 0, CInt(st.Length))
        'st.Close()
        con.Open()


        Dim Str As String = "delete Caculate_Start insert into Caculate_Start (Rpt_Id,clt_Str) select Rpt_Id , STUFF((  select ' '+b.CLT_Amt from Caculate_Rpt b   where b.Rpt_Id = a.Rpt_Id   order by b.cnt for xml path('a'), type).value('.','nvarchar(2000)'),1,1,'') As  CLT_Amt      from Caculate_Rpt a where CLT_Amt <>''group by Rpt_Id"
        Dim Cmd As New System.Data.OleDb.OleDbCommand(Str, con)
        'Cmd.Parameters.Add("", System.Data.OleDb.OleDbType.Binary, Buffer.Length).Value = Buffer
        Cmd.ExecuteNonQuery()
        con.Close()
        'Catch ex As Exception
        '    con.Close()
        '    MsgBox("ກະລຸນນາເລືອກຮູບກ000່ອນ", MsgBoxStyle.Critical, "")
        '    MsgBox(ex.ToString)
        'End Try
    End Sub
    Public Sub Insert_Image()
        If (Fm_Image.PictureBox1.Image Is Nothing) Then
            MsgBox("No Image Is There ")
            Exit Sub
        End If
        Try
            Dim st As New FileStream(Fm_Image.OpenFileDialog1.FileName, FileMode.Open, FileAccess.Read)
            Dim mbr As BinaryReader = New BinaryReader(st)
            Dim buffer(st.Length) As Byte
            mbr.Read(buffer, 0, CInt(st.Length))
            st.Close()
            con.Open()


            Dim Str As String = "insert into Ap_Image(img_id , ImgType ,img)  values( '" & Fm_Image.Img_ID.Text & "' ,'" & Fm_Image.ImgType.Text & "',?)"
            Dim Cmd As New System.Data.OleDb.OleDbCommand(Str, con)
            Cmd.Parameters.Add("@img", System.Data.OleDb.OleDbType.Binary, buffer.Length).Value = buffer
            Cmd.ExecuteNonQuery()
            con.Close()
        Catch ex As Exception
            con.Close()
            MsgBox("ກະລຸນນາເລືອກຮູບກ000່ອນ", MsgBoxStyle.Critical, "")
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Sub SelectImge()
        Try
            Fm_Image.OpenFileDialog1.Filter = "Bmp Files(*.bmp)|*.bmp|Gif Files(*.gif)|*.gif|Jpg Files(*.jpg)|*.jpg"
            Fm_Image.OpenFileDialog1.ShowDialog()
            Dim s As String
            s = Fm_Image.OpenFileDialog1.FileName
            Dim objImage As System.Drawing.Image = System.Drawing.Image.FromFile(s)
            LoadImgSize()
            'MsgBox(b_x & "+++" & b_y)
            If objImage.Width > g_x Or objImage.Height > g_y Then MsgBox("ຂະຫນາດຮູບ (" & objImage.Width & " x " & objImage.Height & ") ໃຫ່ຍເກີນຂະຫນາດ (" & g_x & " x " & g_y & ") ") : Fm_Image.PictureBox1.Image = Fm_Image.a123456789.Image : Exit Sub
            Fm_Image.PictureBox1.Image = Image.FromFile(Fm_Image.OpenFileDialog1.FileName)
            SUPD = 1
        Catch

        End Try
    End Sub


    Public Sub Update_Image()
        If (Fm_Image.PictureBox1.Image Is Nothing) Then
            'MsgBox("No Image Is There ")
            Exit Sub
        End If
        Update_Image(ImageSlno)
    End Sub


    Public Sub Update_Image(ByVal slno As Integer)
        Try
            Dim st As New FileStream(Fm_Image.OpenFileDialog1.FileName, FileMode.Open, FileAccess.Read)
            Dim mbr As BinaryReader = New BinaryReader(st)
            Dim buffer(st.Length) As Byte
            mbr.Read(buffer, 0, CInt(st.Length))
            st.Close()
            con.Open()
            Dim Str As String = "update Ap_Image set Img = ? WHERE Img_Id = '" & Fm_Image.Img_ID.Text & "' And  ImgType = '" & Fm_Image.ImgType.Text & "'"
            Dim Cmd As New System.Data.OleDb.OleDbCommand(Str, con)
            Cmd.Parameters.Add("@Img", System.Data.OleDb.OleDbType.Binary, buffer.Length).Value = buffer
            Cmd.ExecuteNonQuery()
            con.Close()
            'MsgBox("Image Updated Successfully")
        Catch ex As Exception
            con.Close()
            MsgBox(ex.ToString)
        End Try
    End Sub


End Module
