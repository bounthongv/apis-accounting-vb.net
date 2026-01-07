Imports System.Data.SqlClient

Public Class FmCalcu

    Dim sd As String
    Dim int As Integer = 0
    
    ' Setup DataGridView helper functions
    Private Function GetGridValue(ByVal grid As DataGridView, ByVal row As Integer, ByVal col As Integer) As String
        Try
            If grid.RowCount <= row OrElse row < 0 Then Return ""
            If grid.ColumnCount <= col OrElse col < 0 Then Return ""
            If grid.Rows(row).Cells(col).Value Is Nothing Then Return ""
            Return grid.Rows(row).Cells(col).Value.ToString()
        Catch ex As Exception
            Return ""
        End Try
    End Function
    
    Private Sub SetGridValue(ByVal grid As DataGridView, ByVal row As Integer, ByVal col As Integer, ByVal value As Object)
        Try
            If row < 0 Then Exit Sub
            While grid.RowCount <= row
                grid.Rows.Add()
            End While
            If col < grid.ColumnCount Then
                grid.Rows(row).Cells(col).Value = value
            End If
        Catch ex As Exception
            ' Ignore
        End Try
    End Sub
    
    Private Sub FmCalcu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            SetupGrid()
        Catch ex As Exception
            ' Ignore
        End Try
    End Sub
    
    Private Sub SetupGrid()
        Try
            FG.AllowUserToAddRows = True
            FG.AllowUserToDeleteRows = True
            FG.ReadOnly = False
            
            ' Setup columns for calculator grid
            FG.Columns.Clear()
            FG.Columns.Add("Col0", "#")
            FG.Columns.Add("Col1", "Label1")
            FG.Columns.Add("Col2", "Label2")
            FG.Columns.Add("Col3", "Label3")
            FG.Columns.Add("Col4", "Label4")
            FG.Columns.Add("Col5", "Label6")
            FG.Columns.Add("Col6", "Value1")
            FG.Columns.Add("Col7", "Operator")
            FG.Columns.Add("Col8", "Value2")
            FG.Columns.Add("Col9", "Result")
            
            ' Auto-size columns
            For Each col As DataGridViewColumn In FG.Columns
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            Next
            
            ' Set default grid size
            If FG.Rows.Count = 0 Then
                For i As Integer = 0 To 4
                    FG.Rows.Add()
                Next
            End If
        Catch ex As Exception
            ' Ignore
        End Try
    End Sub

Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        Calcu()
        Dim int2 As Integer
        int2 = int + 1
        int = int2
        'MsgBox(int2)
        If int >= 10 Then
            ListBox1.Items.Add(int2 & "... (" & TextBox1.Text & ") " & TextBox4.Text & " (" & TextBox2.Text & " )" & " =  " & TextBox3.Text)
        Else
            ListBox1.Items.Add("0" & int2 & "... (" & TextBox1.Text & ") " & TextBox4.Text & " (" & TextBox2.Text & " )" & " =  " & TextBox3.Text)
        End If
        Letter()
        
        ' Add new row to grid
        Dim rowIndex As Integer = If(FG.Rows.Count > 0, FG.Rows.Count - 1, 0)
        
        ' Set grid values
        SetGridValue(FG, rowIndex, 0, rowIndex + 1)
        SetGridValue(FG, rowIndex, 1, Label1.Text)
        SetGridValue(FG, rowIndex, 2, Label2.Text)
        SetGridValue(FG, rowIndex, 3, Label3.Text)
        SetGridValue(FG, rowIndex, 4, Label4.Text)
        SetGridValue(FG, rowIndex, 5, TextBox1.Text)
        SetGridValue(FG, rowIndex, 6, TextBox4.Text)
        SetGridValue(FG, rowIndex, 7, TextBox2.Text)
        SetGridValue(FG, rowIndex, 8, TextBox3.Text)
    End Sub
    Private Sub Calcu()
        If TextBox4.Text = "+" Then
            Label2.Text = "=> ບວກກັບ (+)"
            TextBox3.Text = CDbl(TextBox1.Text) + CDbl(TextBox2.Text)
        End If
        If TextBox4.Text = "-" Then
            Label2.Text = "=> ລົບອອກ (-)"
            TextBox3.Text = CDbl(TextBox1.Text) - CDbl(TextBox2.Text)
        End If
        If TextBox4.Text = "*" Then
            Label2.Text = "=> ຄູນກັບ (*)"
            TextBox3.Text = CDbl(TextBox1.Text) * CDbl(TextBox2.Text)
        End If
        If TextBox4.Text = "/" Then
            Label2.Text = "=> ຫານໃຫ້ (/)"
            TextBox3.Text = CDbl(TextBox1.Text) / CDbl(TextBox2.Text)
        End If
        Formata()




        Letter()


    End Sub
    Private Sub Formata()
        TextBox2.Text = Format(CDbl(TextBox2.Text), "##,##0.00")
        TextBox3.Text = Format(CDbl(TextBox3.Text), "##,##0.00")
        TextBox1.Text = Format(CDbl(TextBox1.Text), "##,##0.00")
    End Sub

    Private Sub Letter()
        If TextBox4.Text = "+" Then
            Label2.Text = "=> ບວກກັບ (+)"
        End If
        If TextBox4.Text = "-" Then
            Label2.Text = "=> ລົບອອກ (-)"
        End If
        If TextBox4.Text = "*" Then
            Label2.Text = "=> ຄູນກັບ (*)"
        End If
        If TextBox4.Text = "/" Then
            Label2.Text = "=> ຫານໃຫ້ (/)"
        End If
        Label1.Text = "=> " & Letter_amt1(TextBox1) & " (" & TextBox1.Text & ")"
        Label3.Text = "=> " & Letter_amt1(TextBox2) & " (" & TextBox2.Text & ")"
        Label4.Text = "=> " & Letter_amt1(TextBox3) & " (" & TextBox3.Text & ")"
        Label6.Text = "=> ເທົາກັບ (=)"


        If CDbl(TextBox1.Text) = 0 Then
            Label1.Text = "=> ສູນ"
        End If
        If CDbl(TextBox2.Text) = 0 Then
            Label3.Text = "=> ສູນ"
        End If
        If CDbl(TextBox3.Text) = 0 Then
            Label4.Text = "=> ສູນ"
        End If
  
      


    End Sub
    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown

        If e.KeyCode = 76 Then
            TextBox1.Text = TextBox3.Text
            TextBox2.Text = 0
            TextBox2.Focus()
            TextBox2.SelectAll()
        End If
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = Chr(13) Then
            Calcu()
            sd = 1
            Dim int2 As Integer
            int2 = int + 1
            int = int2
            'MsgBox(int2)
            If int >= 10 Then
                ListBox1.Items.Add(int2 & "... (" & TextBox1.Text & ") " & TextBox4.Text & " (" & TextBox2.Text & " )" & " =  " & TextBox3.Text)
            Else
                ListBox1.Items.Add("0" & int2 & "... (" & TextBox1.Text & ") " & TextBox4.Text & " (" & TextBox2.Text & " )" & " =  " & TextBox3.Text)
End If
            FG.Rows.Add()
            Dim rowIndex As Integer = FG.Rows.Count - 1
            SetGridValue(FG, rowIndex, 1, Label1.Text)
            SetGridValue(FG, rowIndex, 2, Label2.Text)
            SetGridValue(FG, rowIndex, 3, Label3.Text)
            SetGridValue(FG, rowIndex, 4, Label4.Text)
            SetGridValue(FG, rowIndex, 5, Label6.Text)
            SetGridValue(FG, rowIndex, 6, TextBox1.Text)
            SetGridValue(FG, rowIndex, 7, TextBox4.Text)
            SetGridValue(FG, rowIndex, 8, TextBox2.Text)
            SetGridValue(FG, rowIndex, 9, TextBox3.Text)




            TextBox1.Focus()
        End If


    
    End Sub
    Public Function Letter_amt1(ByVal Txt As TextBox, Optional ByVal CurrKIP As Boolean = False) As String
        Letter_amt1 = CMoney(Format(CDbl(Txt.Text), "##0.00"))
    End Function

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        'MsgBox(e.KeyCode)
        'If e.KeyCode = 76 Then
        '    TextBox1.Text = TextBox3.Text
        '    TextBox2.Text = 0
        '    TextBox2.Focus()
        'End If
    End Sub


    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Calcu()
            TextBox2.Focus()

        End If
    End Sub
    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

        If Microsoft.VisualBasic.Right(TextBox1.Text, 1) = "+" Then
            Dim s As String
            TextBox4.Text = Microsoft.VisualBasic.Right(TextBox1.Text, 1)
            s = Microsoft.VisualBasic.Left(TextBox1.Text, CDbl(Len(TextBox1.Text)) - 1)
            TextBox1.Text = s
            TextBox2.Focus()
            sd = 2
        End If
        If Microsoft.VisualBasic.Right(TextBox1.Text, 1) = "-" Then
            Dim s As String
            TextBox4.Text = Microsoft.VisualBasic.Right(TextBox1.Text, 1)
            s = Microsoft.VisualBasic.Left(TextBox1.Text, CDbl(Len(TextBox1.Text)) - 1)
            TextBox1.Text = s
            TextBox2.Focus()
            sd = 2
        End If
        If Microsoft.VisualBasic.Right(TextBox1.Text, 1) = "*" Then
            Dim s As String
            TextBox4.Text = Microsoft.VisualBasic.Right(TextBox1.Text, 1)
            s = Microsoft.VisualBasic.Left(TextBox1.Text, CDbl(Len(TextBox1.Text)) - 1)
            TextBox1.Text = s
            TextBox2.Focus()
            sd = 2
        End If
        If Microsoft.VisualBasic.Right(TextBox1.Text, 1) = "/" Then
            Dim s As String
            TextBox4.Text = Microsoft.VisualBasic.Right(TextBox1.Text, 1)
            s = Microsoft.VisualBasic.Left(TextBox1.Text, CDbl(Len(TextBox1.Text)) - 1)
            TextBox1.Text = s
            TextBox2.Focus()
            sd = 2
        End If

        If TextBox1.Text = "" Then
            TextBox1.Text = 0

        End If
        If IsNumeric(TextBox1.Text) = False Then
            TextBox1.Text = 0
            TextBox1.Focus()
            sd = 2
        End If
        'Letter()
    End Sub

    Private Sub TextBox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = Chr(13) Then
            Calcu()
            TextBox2.Focus()
            sd = 2
        End If
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Calcu()
        'TextBox2.Focus()
    End Sub

    Private Sub Form1_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        TextBox1.Focus()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetupGrid()
        Label1.Text = ""
        Label2.Text = ""
        Label3.Text = ""
        Label4.Text = ""
        Label6.Text = ""
        Button11_Click(sender, e)
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click

        TextBox1.Text = "0"
        TextBox2.Text = "0"
        TextBox3.Text = "0"
        TextBox4.Text = "+"

        ListBox1.Items.Clear()
        int = 0
        TextBox1.Focus()
        Letter()
        sd = 1
        FG.Rows.Clear()
        'Button11_Click(sender, e)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If sd = 1 Then
            If TextBox1.Text = 0 Then
                TextBox1.Text = 1
                Letter()

                Exit Sub
            End If

            TextBox1.Text = TextBox1.Text & "1"
        End If
        If sd = 2 Then
            If TextBox2.Text = 0 Then
                TextBox2.Text = 1

                Exit Sub
            End If
            TextBox2.Text = TextBox2.Text & "1"

        End If


        Letter()
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

        If IsNumeric(TextBox2.Text) = False Then

            TextBox2.Text = 0
            TextBox2.Focus()
        End If
        If TextBox1.Text = "" Then
            TextBox1.Text = 0

        End If
        'Letter()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If sd = 1 Then
            If TextBox1.Text = 0 Then
                TextBox1.Text = 2
                Letter()
                Exit Sub
            End If
            TextBox1.Text = TextBox1.Text & "2"
        End If
        If sd = 2 Then
            If TextBox2.Text = 0 Then
                TextBox2.Text = 2
                Exit Sub
            End If
            TextBox2.Text = TextBox2.Text & "2"
        End If
        Letter()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If sd = 1 Then
            If TextBox1.Text = 0 Then
                TextBox1.Text = 3
                Letter()
                Exit Sub
            End If
            TextBox1.Text = TextBox1.Text & "3"
        End If
        If sd = 2 Then
            If TextBox2.Text = 0 Then
                TextBox2.Text = 3
                Exit Sub
            End If
            TextBox2.Text = TextBox2.Text & "3"
        End If
        Letter()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If sd = 1 Then
            If TextBox1.Text = 0 Then
                TextBox1.Text = 4
                Letter()
                Exit Sub
            End If
            TextBox1.Text = TextBox1.Text & "4"
        End If
        If sd = 2 Then
            If TextBox2.Text = 0 Then
                TextBox2.Text = 4
                Exit Sub
            End If
            TextBox2.Text = TextBox2.Text & "4"
        End If
        Letter()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        If sd = 1 Then
            If TextBox1.Text = 0 Then
                TextBox1.Text = 5
                Letter()
                Exit Sub
            End If
            TextBox1.Text = TextBox1.Text & "5"
        End If
        If sd = 2 Then
            If TextBox2.Text = 0 Then
                TextBox2.Text = 5
                Exit Sub
            End If
            TextBox2.Text = TextBox2.Text & "5"
        End If
        Letter()








    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If sd = 1 Then
            If TextBox1.Text = 0 Then
                TextBox1.Text = 6
                Letter()
                Exit Sub
            End If
            TextBox1.Text = TextBox1.Text & "6"
        End If
        If sd = 2 Then
            If TextBox2.Text = 0 Then
                TextBox2.Text = 6
                Exit Sub
            End If
            TextBox2.Text = TextBox2.Text & "6"
        End If
        Letter()






    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        If sd = 1 Then
            If TextBox1.Text = 0 Then
                TextBox1.Text = 7
                Letter()
                Exit Sub
            End If
            TextBox1.Text = TextBox1.Text & "7"
        End If
        If sd = 2 Then
            If TextBox2.Text = 0 Then
                TextBox2.Text = 7
                Exit Sub
            End If
            TextBox2.Text = TextBox2.Text & "7"
        End If
        Letter()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click

        If sd = 1 Then
            If TextBox1.Text = 0 Then
                TextBox1.Text = 8
                Letter()
                Exit Sub
            End If
            TextBox1.Text = TextBox1.Text & "8"
        End If
        If sd = 2 Then
            If TextBox2.Text = 0 Then
                TextBox2.Text = 8
                Exit Sub
            End If
            TextBox2.Text = TextBox2.Text & "8"
        End If
        Letter()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click

        If sd = 1 Then
            If TextBox1.Text = 0 Then
                TextBox1.Text = 9
                Letter()
                Exit Sub
            End If
            TextBox1.Text = TextBox1.Text & "9"
        End If
        If sd = 2 Then
            If TextBox2.Text = 0 Then
                TextBox2.Text = 9
                Exit Sub
            End If
            TextBox2.Text = TextBox2.Text & "9"
        End If
        Letter()


    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        If sd = 1 Then
            TextBox1.Text = TextBox1.Text & "0"
        End If
        If sd = 2 Then
            TextBox2.Text = TextBox2.Text & "0"
        End If
        Letter()
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click

        If sd = 1 Then
            TextBox1.Text = TextBox1.Text & "00"
        End If
        If sd = 2 Then
            TextBox2.Text = TextBox2.Text & "00"
        End If
        Letter()
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        If sd = 1 Then
            TextBox1.Text = TextBox1.Text & "000"
        End If
        If sd = 2 Then
            TextBox2.Text = TextBox2.Text & "000"
        End If
        Letter()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        sd = 2
        TextBox4.Text = "+"
        TextBox2.Focus()
        Letter()
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        sd = 2
        TextBox4.Text = "-"
        Letter()
        TextBox2.Focus()

    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        sd = 2
        TextBox4.Text = "*"
        Letter()
        TextBox2.Focus()
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        sd = 2
        TextBox4.Text = "/"
        Letter()
        TextBox2.Focus()
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        If sd = 1 Then
            TextBox1.Text = TextBox1.Text & "."
        End If
        If sd = 2 Then
            TextBox2.Text = TextBox2.Text & "."
        End If
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        If sd = 1 Then
            If TextBox1.Text = "0" Then
                TextBox1.Text = 0
                Exit Sub
            End If
            Dim s As String
            s = Microsoft.VisualBasic.Left(TextBox1.Text, CDbl(Len(TextBox1.Text)) - 1)
            TextBox1.Text = s
        End If
        If sd = 2 Then
            If TextBox2.Text = "0" Then
                TextBox2.Text = 0
                Exit Sub
            End If
            Dim s As String
            s = Microsoft.VisualBasic.Left(TextBox2.Text, CDbl(Len(TextBox2.Text)) - 1)
            TextBox2.Text = s
        End If
        Letter()
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        Call MdiCNum()
        Close()
    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label8_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        TextBox1.Text = TextBox3.Text
        TextBox2.Text = 0
        TextBox2.Focus()
        'ListBox1.add()
        'ListBox1.Items.Add("aa")
        'ListBox1.Items.Clear()
        'ListView1.SelectedItems(0)
        Letter()
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click

        Me.Visible = False
    End Sub

    Private Sub ListBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.DoubleClick
        Dim rowIndex As Integer = ListBox1.SelectedIndex
        If rowIndex >= 0 AndAlso rowIndex < FG.Rows.Count Then
            TextBox1.Text = GetGridValue(FG, rowIndex, 6)
            TextBox4.Text = GetGridValue(FG, rowIndex, 7)
            TextBox2.Text = GetGridValue(FG, rowIndex, 8)
            TextBox3.Text = GetGridValue(FG, rowIndex, 9)
        End If
        TextBox1.Focus()
    End Sub

    Private Sub ListBox1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBox1.MouseDown
        Dim rowIndex As Integer = ListBox1.SelectedIndex
        If rowIndex >= 0 AndAlso rowIndex < FG.Rows.Count Then
            Label1.Text = GetGridValue(FG, rowIndex, 1)
            Label2.Text = GetGridValue(FG, rowIndex, 2)
            Label3.Text = GetGridValue(FG, rowIndex, 3)
            Label6.Text = GetGridValue(FG, rowIndex, 4)
            Label4.Text = GetGridValue(FG, rowIndex, 5)
        End If
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        'MsgBox(ListBox1.SelectedIndex)


       
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

End Class
