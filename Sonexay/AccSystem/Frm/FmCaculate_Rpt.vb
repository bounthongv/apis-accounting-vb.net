Public Class FmCaculate_Rpt
    Dim x As String

    Private Sub ComboBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            LoadFg2()
        End If
    End Sub
    Private Sub LoadFg2()
        FG2.Rows = 1
        If ComboBox1.Text = "INC" Then
            With RSC
                Call LoadSqlData("select * from Ap_Rpt_Income order by Rpt_Id asc", RSC)
                If .RecordCount > 0 Then
                    While Not .EOF()
                        FG2.AddItem(.AbsolutePosition & _
                        Chr(9) & (.Fields("Rpt_ID").Value.ToString) & _
                        Chr(9) & (.Fields("Description").Value.ToString) & _
                             Chr(9) & (.Fields("CLT_Str").Value.ToString) & _
                                  Chr(9) & (.Fields("Fnt").Value.ToString) & _
                                       Chr(9) & (.Fields("Clor").Value.ToString) & _
                        Chr(9) & (.Fields("Udln").Value.ToString))
                        .MoveNext()
                    End While
                End If
            End With
        ElseIf ComboBox1.Text = "BLS" Then
            With RSC
                Call LoadSqlData("select * from Ap_Rpt_BLS order by Rpt_Id asc", RSC)
                If .RecordCount > 0 Then
                    While Not .EOF()
                        FG2.AddItem(.AbsolutePosition & _
                          Chr(9) & (.Fields("Rpt_ID").Value.ToString) & _
                          Chr(9) & (.Fields("Description").Value.ToString) & _
                               Chr(9) & (.Fields("CLT_Str").Value.ToString) & _
                                    Chr(9) & (.Fields("Fnt").Value.ToString) & _
                                         Chr(9) & (.Fields("Clor").Value.ToString) & _
                          Chr(9) & (.Fields("Udln").Value.ToString))
                        .MoveNext()
                    End While
                End If
            End With
        ElseIf ComboBox1.Text = "CAF" Then
            With RSC
                Call LoadSqlData("select * from Ap_Rpt_Cashflow order by Rpt_Id asc", RSC)
                If .RecordCount > 0 Then
                    While Not .EOF()
                        FG2.AddItem(.AbsolutePosition & _
                           Chr(9) & (.Fields("Rpt_ID").Value.ToString) & _
                           Chr(9) & (.Fields("Description").Value.ToString) & _
                                Chr(9) & (.Fields("CLT_Str").Value.ToString) & _
                                     Chr(9) & (.Fields("Fnt").Value.ToString) & _
                                          Chr(9) & (.Fields("Clor").Value.ToString) & _
                           Chr(9) & (.Fields("Udln").Value.ToString))
                        .MoveNext()
                    End While
                End If
            End With
        ElseIf ComboBox1.SelectedIndex > 2 Then
            With RSC
                Call LoadSqlData("select * from So_Rpt_Pro Where RptType = '" & ComboBox1.Text & "' order by RptId asc ", RSC)
                If .RecordCount > 0 Then
                    While Not .EOF()
                        FG2.AddItem(.AbsolutePosition & _
                           Chr(9) & (.Fields("RptID").Value.ToString) & _
                           Chr(9) & (.Fields("Des").Value.ToString) & _
                                Chr(9) & (.Fields("StrCal").Value.ToString) & _
                                     Chr(9) & (.Fields("Fnb").Value.ToString) & _
                                          Chr(9) & (.Fields("Cor").Value.ToString) & _
                           Chr(9) & (.Fields("Und").Value.ToString))
                        .MoveNext()
                    End While
                End If
            End With
        End If



        For i = 1 To FG2.Rows - 1
            FG2.Col = 2
            FG2.Row = i
            If FG2.get_TextMatrix(i, 4) = "1" Then
                FG2.CellFontBold = True
            End If
            If FG2.get_TextMatrix(i, 5) = "1" Then
                FG2.CellForeColor = Color.Red
            End If
            If FG2.get_TextMatrix(i, 5) = "2" Then
                FG2.CellForeColor = Color.Blue
            End If
            If FG2.get_TextMatrix(i, 6) = "1" Then
                FG2.CellFontUnderline = True
            End If
        Next i


    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
       
      
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Close()
    End Sub

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        Dim CLT, Fnt, Clor, Udln As String
        Fnt = FontStype.SelectedIndex
        Clor = FontColor.SelectedIndex
        If CheckBox1.Checked = True Then
            Udln = 1
        Else
            Udln = 0
        End If
        CLT = ""
        CNN.Execute("Delete Caculate_Rpt where Rpt_Id = '" & TextBox1.Text & "' And Rpt_Type =  '" & ComboBox1.Text & "' ")
        For i = 1 To FG.Cols - 2
            CNN.Execute("Insert Into Caculate_Rpt (Rpt_Id , CLT_Str , Rpt_Type ) Values ('" & TextBox1.Text & "' , '" & FG.get_TextMatrix(1, i) & "' , '" & ComboBox1.Text & "') ")
            CLT = CLT & FG.get_TextMatrix(1, i)
        Next i
        If ComboBox1.SelectedIndex > 2 Then
            CNN.Execute("Update So_Rpt_Pro set StrCal = '" & CLT & "' , Fnb = '" & Fnt & "' , Cor = '" & Clor & "'  , Und = '" & Udln & "' where RptId =  '" & TextBox1.Text & "' and RptType = '" & ComboBox1.Text & "'  ")
        End If

        If ComboBox1.Text = "BLS" Then
            CNN.Execute("Update Ap_Rpt_BLS set CLT_Str = '" & CLT & "' , Fnt = '" & Fnt & "' , Clor = '" & Clor & "'  , Udln = '" & Udln & "' where Rpt_Id =  '" & TextBox1.Text & "'")

        ElseIf ComboBox1.Text = "INC" Then
            CNN.Execute("Update Ap_Rpt_Income set CLT_Str = '" & CLT & "' , Fnt = '" & Fnt & "' , Clor = '" & Clor & "'  , Udln = '" & Udln & "'  where Rpt_Id =  '" & TextBox1.Text & "'")
        ElseIf ComboBox1.Text = "CAF" Then
            CNN.Execute("Update Ap_Rpt_Cashflow set CLT_Str = '" & CLT & "' , Fnt = '" & Fnt & "' , Clor = '" & Clor & "'  , Udln = '" & Udln & "'  where Rpt_Id =  '" & TextBox1.Text & "'")

        End If
        MsgBox("ການບັນທຶກສຳເລັດຜົນ")
        CLock.Checked = False
        LoadFg2()
    End Sub

    Private Sub FG_CellEndEdit(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles FG.CellEndEdit

        ' Check if the last column has content and add a new column if needed
        If FG.Columns.Count > 0 AndAlso FG.Rows.Count > 1 AndAlso FG.Rows(1).Cells(FG.Columns.Count - 1).Value IsNot Nothing Then
            If FG.Rows(1).Cells(FG.Columns.Count - 1).Value.ToString() <> "" Then
                ' Add a new column if needed - though this is unusual for DataGridView
                ' Usually you wouldn't dynamically add columns like this
            End If
        End If
        ' The rest of the logic doesn't translate directly to DataGridView as columns aren't typically added dynamically
    End Sub

    Private Sub FG_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG.SelectionChanged
        ' DataGridView doesn't have an Editable property like VSFlexGrid
        ' The editability is typically controlled at the column or cell level
        If FG.CurrentRow IsNot Nothing Then
            If FG.CurrentRow.Index = 2 Then
                ' Make row 2 non-editable by setting ReadOnly property on cells
                For i As Integer = 0 To FG.Columns.Count - 1
                    FG.Rows(2).Cells(i).ReadOnly = True
                Next
            Else
                ' Make other rows editable by setting ReadOnly property on cells
                For i As Integer = 0 To FG.Columns.Count - 1
                    FG.Rows(FG.CurrentRow.Index).Cells(i).ReadOnly = False
                Next
            End If
        End If
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        'CNN.Execute("Insert Into Caculate_Rpt (Rpt_Id , CLT , Rpt_Type ) Values ('" & TextBox1.Text & "' , '" & FG.get_TextMatrix(1, i) & "' , '" & ComboBox1.Text & "') ")
      
        If e.KeyChar = Chr(13) Then
            x = 1
            FG.Cols = 2
            Try
                Dim dt As DataTable = DbHelper.GetDataTable("select * from Caculate_Rpt where Rpt_Id = '" & TextBox1.Text & "' And  Rpt_Type = '" & ComboBox1.Text & "' ")
                If dt.Rows.Count > 0 Then
                    For Each row As DataRow In dt.Rows
                        'MsgBox(row("CLT"))
                        FG.set_TextMatrix(1, x, row("CLT").ToString())
                        x = x + 1
                        FG.Cols = FG.Cols + 1
                    Next
                End If
            Catch ex As Exception
                VSysError = True
                MessageBox.Show("Database Error: " & ex.Message, "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            'FG.Cols = FG.Cols + 1
        End If
    End Sub




    Private Sub FG2_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FG2.SelectionChanged
        'MsgBox(FG2.get_TextMatrix(FG2.Row, 4))
  

        If CLock.Checked = False Then
            If FG2.Row > 0 Then

                'MsgBox(FG2.get_TextMatrix(FG.Row, 1))

                TextBox1.Text = FG2.get_TextMatrix(FG2.Row, 1)
                TextBox2.Text = FG2.get_TextMatrix(FG2.Row, 2)
                TextBox2.Text = FG2.get_TextMatrix(FG2.Row, 2)
                x = 1
                FG.Cols = 1
                FG.Cols = 2
                Try
                    Dim dt As DataTable = DbHelper.GetDataTable("select * from Caculate_Rpt where Rpt_Id = '" & TextBox1.Text & "' And  Rpt_Type = '" & ComboBox1.Text & "' Order by cnt ")
                    If dt.Rows.Count > 0 Then
                        For Each row As DataRow In dt.Rows
                            FG.set_TextMatrix(1, x, row("CLT_Str").ToString())
                            x = x + 1
                            FG.Cols = FG.Cols + 1
                        Next
                    End If
                Catch ex As Exception
                    VSysError = True
                    MessageBox.Show("Database Error: " & ex.Message, "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If

        Else
            '=========
            If FG2.get_TextMatrix(FG2.Row, 3) <> "" Then
                x = FG.Cols - 1
                'FG.Cols = 1
                'FG.Cols = 2
                With RSC
                    Call LoadSqlData("select * from Caculate_Rpt where Rpt_Id = '" & FG2.get_TextMatrix(FG2.Row, 1) & "' And  Rpt_Type = '" & ComboBox1.Text & "' ", RSC)
                    If .RecordCount > 0 Then
                        While Not .EOF()
                            FG.set_TextMatrix(1, x, (.Fields("CLT_Str").Value))
                            x = x + 1
                            FG.Cols = FG.Cols + 1
                            .MoveNext()
                        End While
                    End If
                End With

            Else

                FG.set_TextMatrix(1, FG.Cols - 1, FG2.get_TextMatrix(FG2.Row, 1))
                FG.Cols = FG.Cols + 1

            End If



        End If
        TextBox3.Text = ""
        For i = 1 To FG.Cols - 2
            TextBox3.Text = TextBox3.Text & FG.get_TextMatrix(1, i)
        Next i

        FontColor.SelectedIndex = FG2.get_TextMatrix(FG2.Row, 5)
        FontStype.SelectedIndex = FG2.get_TextMatrix(FG2.Row, 4)
        If FG2.get_TextMatrix(FG2.Row, 6) = "1" Then
            CheckBox1.Checked = True
        Else
            CheckBox1.Checked = False
        End If
    End Sub

    Private Sub FmCaculate_Rpt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FontColor.SelectedIndex = 0
        FontStype.SelectedIndex = 0
        FG.FormatString = "   |<             "
        FG2.Cols = 7
        FG2.Size = New System.Drawing.Size(725, 406)
        FG2.FormatString = "ລ/ດ|<ລະຫັດ       |<ເນື້ອໃນລາຍການ                                          |<ສູດຄິດໄລ່     |<Font|<Color|<Under"
    End Sub


    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        FG.set_TextMatrix(1, FG.Cols - 1, "(")
        FG.Cols = FG.Cols + 1
        TextBox3.Text = ""
        For i = 1 To FG.Cols - 2
            TextBox3.Text = TextBox3.Text & FG.get_TextMatrix(1, i)
        Next i
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        FG.set_TextMatrix(1, FG.Cols - 1, ")")
        FG.Cols = FG.Cols + 1
        TextBox3.Text = ""
        For i = 1 To FG.Cols - 2
            TextBox3.Text = TextBox3.Text & FG.get_TextMatrix(1, i)
        Next i
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        FG.set_TextMatrix(1, FG.Cols - 1, "+")
        FG.Cols = FG.Cols + 1
        TextBox3.Text = ""
        For i = 1 To FG.Cols - 2
            TextBox3.Text = TextBox3.Text & FG.get_TextMatrix(1, i)
        Next i
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        FG.set_TextMatrix(1, FG.Cols - 1, "-")
        FG.Cols = FG.Cols + 1
        TextBox3.Text = ""
        For i = 1 To FG.Cols - 2
            TextBox3.Text = TextBox3.Text & FG.get_TextMatrix(1, i)
        Next i
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        FG.set_TextMatrix(1, FG.Cols - 1, "*")
        FG.Cols = FG.Cols + 1
        TextBox3.Text = ""
        For i = 1 To FG.Cols - 2
            TextBox3.Text = TextBox3.Text & FG.get_TextMatrix(1, i)
        Next i
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        FG.set_TextMatrix(1, FG.Cols - 1, "/")
        FG.Cols = FG.Cols + 1
        TextBox3.Text = ""
        For i = 1 To FG.Cols - 2
            TextBox3.Text = TextBox3.Text & FG.get_TextMatrix(1, i)
        Next i
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        FG.Cols = 1
        FG.Cols = 2
        TextBox3.Text = ""
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        LoadFg2()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Call LoadFg2()
        CNN.Execute(" Update Caculate_Rpt set  CLT_Amt  = 0 ,  CLT_Last_Amt  = 0  ")
        CNN.Execute(" Update Caculate_Rpt set  CLT_Amt  = CLT_Str ,  CLT_Last_Amt  = CLT_Str where CLT_Str = '+' or CLT_Str = '-' or CLT_Str = '*' or CLT_Str = '+' or CLT_Str = '/' or CLT_Str = '(' or CLT_Str=')' ")
        Dim LSTR As String = ""
        Dim LSTR_Last As String = ""
        Dim LSTR_CLT As String = ""
        For i = 1 To FG2.Rows - 1
         

            If FG2.get_TextMatrix(i, 3) <> "" Then
                LSTR = ""
                LSTR_Last = ""
                LSTR_CLT = ""
                ''=====
                With RSC
                    Call LoadSqlData("select CLT_Amt,CLT_Last_Amt,clt_Str  from Caculate_Rpt where Rpt_Id = '" & FG2.get_TextMatrix(i, 1) & "' And Rpt_Type = '" & ComboBox1.Text & "' Order by cnt asc", RSC)
                    If .RecordCount > 0 Then
                        While Not .EOF()
                            LSTR = LSTR & (RSC.Fields("CLT_Amt").Value.ToString)
                            LSTR_Last = LSTR_Last & (RSC.Fields("CLT_Last_Amt").Value.ToString)
                            LSTR_CLT = LSTR_CLT & (RSC.Fields("clt_Str").Value.ToString)
                            .MoveNext()
                        End While
                    End If
                End With

                On Error GoTo hang
hang:
                If Err.Number = 0 Then
                    CNN.Execute(" Update  Caculate_Test set Amt = " & LSTR & "  ")
                    CNN.Execute(" Update  Caculate_Test set Amt = " & LSTR_Last & "  ")
                Else
                    MessageBox.Show("ສູດຄິດໄລ່ຂອງ " & FG2.get_TextMatrix(i, 1) & " =   " & LSTR_CLT & " ບໍ່ຖຶກຕ້ອງກະລຸນນາກວດສອບຄືນໃຫມ່")
                    FG2.Col = 3
                    FG2.Row = i
                    FG2.BackColorSel = Color.Red
                    Exit Sub
                End If

                '=======
            End If
        Next i
        TextBox3.Text = ""
        MsgBox("ການກວດສອບສູດຄິດໄລ່ສຳເລັດຜົນ (ຜ່ານ)")
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        If FG.Cols > 1 Then
            FG.Cols = FG.Cols - 2
            FG.Cols = FG.Cols + 1
            TextBox3.Text = ""
            For i = 1 To FG.Cols - 2
                TextBox3.Text = TextBox3.Text & FG.get_TextMatrix(1, i)
            Next i
        End If
     
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        FG.set_TextMatrix(1, FG.Cols - 1, Button12.Text)
        FG.Cols = FG.Cols + 1
        TextBox3.Text = ""
        For i = 1 To FG.Cols - 2
            TextBox3.Text = TextBox3.Text & FG.get_TextMatrix(1, i)
        Next i
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        FG.set_TextMatrix(1, FG.Cols - 1, Button13.Text)
        FG.Cols = FG.Cols + 1
        TextBox3.Text = ""
        For i = 1 To FG.Cols - 2
            TextBox3.Text = TextBox3.Text & FG.get_TextMatrix(1, i)
        Next i
    End Sub
End Class