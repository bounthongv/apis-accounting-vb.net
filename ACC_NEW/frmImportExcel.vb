
Imports System.Data.SqlClient
Public Class frmImportExcel
    Dim DtSet As System.Data.DataSet
    Private Sub cmdAdd_Click(sender As System.Object, e As System.EventArgs) Handles cmdAdd.Click
        Try
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim rvl As Boolean = False
            Dim fBrowse As New OpenFileDialog
            With fBrowse
                .Filter = "Excel files(*.xlsx)|*.xlsx|All files (*.*)|*.*"
                .FilterIndex = 1
                .Title = "Import data from Excel file"
            End With
            If fBrowse.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Dim fname As String
                fname = fBrowse.FileName
                'MyConnection = New System.Data.OleDb.OleDbConnection("Dsn=Excel Files;dbq=D:\LHSETEST\BCELAccStatement_093110000869828001_04-01-2019_04-01-2019.xlsx;defaultdir=D:\LHSETEST;driverid=1046;maxbuffersize=2048;pagetimeout=5")
                MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source='" & fname & " '; " & "Extended Properties=Excel 8.0;")
                MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection)
                '  MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from test", MyConnection)
                MyCommand.TableMappings.Add("Table", "Test")
                DtSet = New System.Data.DataSet
                If rvl = True Then DtSet.Tables(0).Rows.Clear()
                MyCommand.Fill(DtSet)
                If DtSet.Tables(0).Rows.Count > 0 Then
                    rvl = True
                Else
                    rvl = False
                End If
                MyConnection.Close()
                With DataGridView1
                    .DataSource = DtSet.Tables(0)
                    .Refresh()
                End With


                '    txtBankDate.Value = DtSet.Tables(0).Rows(17).Item(0)
                '    For Each Drr As DataRow In DtSet.Tables(0).Rows
                '   Drr(2).ToString()
                '   Execute_Local("INSERT INTO Excel(Name, Designation, Salary) VALUES ('" & Drr(0).ToString & "','" & Drr(1).ToString & "','" & Drr(2).ToString & "')")
                '  Next
                MsgBox("Load Successfully")

            End If
        Catch ex As Exception
            '    MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub frmImportExcel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub cmdDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDel.Click

    End Sub
End Class