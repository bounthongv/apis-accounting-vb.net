Public Class Form1
    Dim s As String = 0
    Dim w As Integer = 0
    Dim k As String
    Dim sql As String
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
    Private Sub LoadSQL()
        sql = ""
    End Sub

    Private Sub LoadFirstRecord()
        Call LoadSQL()
        Call LoadSqlData("SELECT CuSId FROM  Table_1  where CuSId<>''  " & sql & " ", RSC)
        With RSC
            Record.Text = "1/" & .RecordCount
            w = 1
            If w = 1 Then
                NextRecord.Enabled = True
                LastRecord.Enabled = True
                BackRecord.Enabled = False
                FirstRecord.Enabled = False
            End If
        End With
        LoadSqlData("select top 1 CuSId from Table_1  where CuSId <>'' " & sql & " order by CuSId  asc ", RSC)
        With RSC
            Do Until .EOF = True
                k = (.Fields("CuSId").Value)
                Cust_ID.Text = (.Fields("CuSId").Value)
                s = k
                .MoveNext()
            Loop
        End With
        Call loadText()
    End Sub
 
 
    Private Sub LoadBackRecord()
        Call LoadSQL()
        Dim RSC1 As New ADODB.Recordset
        Dim RS, RS1 As New ADODB.Recordset
        Call LoadSqlData("SELECT CuSId FROM  Table_1  where CuSId<>'' " & sql & " ", RSC)
        With RSC
            w = w - 1
            Record.Text = w & "/" & .RecordCount
            If w = 1 Then
                NextRecord.Enabled = True
                LastRecord.Enabled = True
                BackRecord.Enabled = False
                FirstRecord.Enabled = False
            End If
        End With
        LoadSqlData("select top 1 CuSId from Table_1  " & _
        " where  CuSId <>'' and  CuSId <" & s & " " & sql & "  order by CuSId desc", RSC)
        With RSC
            Do Until .EOF = True
                Cust_ID.Text = (.Fields("CuSId").Value)
                k = Cust_ID.Text
                s = k
                .MoveNext()
            Loop
        End With
        Call loadText()
    End Sub
    Private Sub loadNextCount()
        Call LoadSQL()
        Dim RSC1 As New ADODB.Recordset
        Dim RS, RS1 As New ADODB.Recordset
        Call LoadSqlData("SELECT CuSId FROM  Table_1  where CuSId<>'' " & sql & " ", RSC)
        With RSC
            w = w + 1
            Record.Text = w & "/" & .RecordCount
            If w = .RecordCount Then
                NextRecord.Enabled = False
                LastRecord.Enabled = False
                BackRecord.Enabled = True
                FirstRecord.Enabled = True
            End If
        End With
        LoadSqlData("select top 1 CuSId from Table_1 " & _
        " where   CuSId <>'' and CuSId >" & s & " " & sql & "  order by CuSId asc", RSC)
        With RSC
            Do Until .EOF = True
                k = (.Fields("CuSId").Value)
                Cust_ID.Text = (.Fields("CuSId").Value)
                s = k
                .MoveNext()
            Loop
        End With
        Call loadText()
    End Sub

    Private Sub LoadLastRecord()
        Call LoadSQL()
        Dim RSC1 As New ADODB.Recordset
        Dim RS, RS1 As New ADODB.Recordset
        Call LoadSqlData("SELECT CuSId FROM  Table_1  where CuSId<>'' " & sql & " ", RSC)
        With RSC
            w = .RecordCount
            Record.Text = .RecordCount & "/" & .RecordCount
            If w = .RecordCount Then
                NextRecord.Enabled = False
                LastRecord.Enabled = False
                BackRecord.Enabled = True
                FirstRecord.Enabled = True
            End If
        End With
        LoadSqlData("select top 1 CuSId from Table_1  where CuSId <>'' " & sql & "   order by CuSId  desc ", RSC)
        With RSC
            Do Until .EOF = True
                Cust_ID.Text = (.Fields("CuSId").Value)
                k = Cust_ID.Text
                s = k
                .MoveNext()
            Loop
        End With
        Call loadText()
    End Sub
    Private Sub loadText()
        LoadSqlData("select * from Table_1 where CuSId = '" & Cust_ID.Text & "'", RSC)
        With RSC
            Do Until .EOF = True
                Cust_ID.Text = (.Fields("CuSId").Value)
                Nme.Text = (.Fields("Nme").Value)
                .MoveNext()
            Loop
        End With
        SUPD = 0
    End Sub

    Private Sub BackRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackRecord.Click
        Call LoadBackRecord()
   
    End Sub

    Private Sub NextRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NextRecord.Click

        Call loadNextCount()

    End Sub

    Private Sub FirstRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FirstRecord.Click
        Call LoadFirstRecord()

    End Sub
    Private Sub LastRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LastRecord.Click
        Call LoadLastRecord()
    End Sub
End Class