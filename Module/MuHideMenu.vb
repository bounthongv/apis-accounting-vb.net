Imports System.Data.SqlClient
Module MuHideMenu
    Dim MuHide As String = ""
    Dim MuMENM As String = ""
    Dim mItem As ToolStripMenuItem
    Dim mSubItem As ToolStripMenuItem
    Dim Usr As String
    Dim NS1, NS2, NS3 As String
    Dim NM1, NM2, NM3 As Integer

    Public Sub HideMenu()
        If MUserID <> "a" Then
            Try
                Dim dt As DataTable = GetDataTable("select Menu_Nm from Ap_Section_Item where Ints = 0 And Usr_Id= '" & MUserID & "' Order by Menu_Nm")
                For Each row As DataRow In dt.Rows
                    MuHide = row("Menu_Nm").Value
                    Call LoadMenuItem()
                Next
            Catch ex As Exception
                MsgBox("Error loading menu: " & ex.Message)
            End Try
        End If
    End Sub

   
    Public Sub LoadMenuItem()
        For i = 0 To FmMain.MenuStrip1.Items.Count - 1
            If MuHide = FmMain.MenuStrip1.Items(i).Name Then
                FmMain.MenuStrip1.Items.Item(i).Visible = False
            End If
            mItem = CType(FmMain.MenuStrip1.Items.Item(i), ToolStripMenuItem)
            For j = 0 To mItem.DropDownItems.Count - 1
                If MuHide = mItem.DropDownItems(j).Name Then
                    mItem.DropDownItems(j).Visible = False
                End If

                If mItem.DropDownItems(j).Text <> "" Then
                    mSubItem = mItem.DropDownItems(j)
                    For a = 0 To mSubItem.DropDownItems.Count - 1
                        If MuHide = mSubItem.DropDownItems(a).Name Then
                            mSubItem.DropDownItems(a).Visible = False
                        End If
                    Next a
                End If

            Next j
        Next i
    End Sub
    Public Sub InsertMenuStrip1()
        NM1 = 0
        NM2 = 0
        NM3 = 0
        ExecuteNonQuery("Delete Ap_Section   Delete Ap_Section_AdNew Delete Ap_Section_Item ")
        Dim N As Integer = 0
        For i = 0 To FmMain.MenuStrip1.Items.Count - 1
            N = N + 1
            If FmMain.MenuStrip1.Items.Item(i).Text <> "" Then
                ExecuteNonQuery("Insert Into Ap_Section (Sec_ID , Sec_Nm , Menu_Nm ) VAlues (" & N & " ,N'" & FmMain.MenuStrip1.Items.Item(i).Text & "'   ,N'" & FmMain.MenuStrip1.Items.Item(i).Name & "' )")
            End If
        Next i
        Dim dt As DataTable = DbHelper.GetDataTable("select Usr_id  from  Ap_Users  Order by Usr_id")
        For Each row As DataRow In dt.Rows
            Usr = row("Usr_id").ToString()
            LoadUsrACC()
        Next
        ExecuteNonQuery("Update Ap_Section_AdNew set Ints=Ap_Section_Item.Ints from Ap_Section_AdNew , Ap_Section_Item where Ap_Section_AdNew.Menu_Nm = Ap_Section_Item.Menu_Nm And Ap_Section_AdNew.Usr_Id = Ap_Section_Item.Usr_Id")
        ExecuteNonQuery("delete Ap_Section_Item")
        ExecuteNonQuery("insert  into Ap_Section_Item (Sec_ID, Sec_Nm, Ints, Menu_Nm , Usr_Id) select Sec_ID, Sec_Nm, Ints, Menu_Nm , Usr_Id from Ap_Section_AdNew Order by cnt ")
        ExecuteNonQuery("Update Ap_Section_Item set Ints = 1 where  Usr_Id = 'a'  ")
        ExecuteNonQuery("Update Ap_Section_Item set Ints = 1 ")
    End Sub
    Public Sub InsertMenuStrip_Usr()
        Usr = FrmUser_DDC.txtUsr_id.Text
        ExecuteNonQuery(" Delete Ap_Section   Delete Ap_Section_AdNew  ")

        Dim N As Integer = 0
        For i = 0 To FmMain.MenuStrip1.Items.Count - 1
            N = N + 1
            If FmMain.MenuStrip1.Items.Item(i).Text <> "" Then
                ExecuteNonQuery("Insert Into Ap_Section (Sec_ID , Sec_Nm , Menu_Nm ) VAlues (" & N & " ,N'" & FmMain.MenuStrip1.Items.Item(i).Text & "'   ,N'" & FmMain.MenuStrip1.Items.Item(i).Name & "' )")
            End If
        Next i
        LoadUsrACC()
        'MsgBox(Usr)
        ExecuteNonQuery("Update Ap_Section_AdNew set Ints=Ap_Section_Item.Ints from Ap_Section_AdNew , Ap_Section_Item where Ap_Section_AdNew.Menu_Nm = Ap_Section_Item.Menu_Nm And Ap_Section_AdNew.Usr_Id = Ap_Section_Item.Usr_Id ")
        ExecuteNonQuery("delete Ap_Section_Item where Usr_Id='" & Usr & "'")
        ExecuteNonQuery("insert  into Ap_Section_Item (Sec_ID, Sec_Nm, Ints, Menu_Nm , Usr_Id) select Sec_ID, Sec_Nm, Ints, Menu_Nm , Usr_Id from Ap_Section_AdNew Order by cnt ")
        ExecuteNonQuery("Update Ap_Section_Item set Ints = 1 where Usr_Id='" & Usr & "'")
    End Sub

    Private Sub LoadUsrACC()

        Dim N As Integer = 0
        For i = 0 To FmMain.MenuStrip1.Items.Count - 1
            N = N + 1
            NM2 = 0
            mItem = CType(FmMain.MenuStrip1.Items.Item(i), ToolStripMenuItem)
            mItem = CType(FmMain.MenuStrip1.Items.Item(i), ToolStripMenuItem)
            ExecuteNonQuery("Insert Into Ap_Section_AdNew (Sec_ID , Sec_Nm ,Ints, Menu_Nm , Usr_Id) VAlues (" & N & " ,N'" & N & " " & FmMain.MenuStrip1.Items(i).Text & "'   ,1,N'" & FmMain.MenuStrip1.Items(i).Name & "' , '" & Usr & "' )")
            For j = 0 To mItem.DropDownItems.Count - 1
                'MsgBox(mItem.DropDownItems(j).Text)
                If mItem.DropDownItems(j).Text <> "" Then
                    ExecuteNonQuery("Insert Into Ap_Section_AdNew (Sec_ID , Sec_Nm ,Ints, Menu_Nm , Usr_Id) VAlues (" & N & " ,N'*  " & mItem.DropDownItems(j).Text & "'   ,0,N'" & mItem.DropDownItems(j).Name & "' , '" & Usr & "' )")

                Else
                    ExecuteNonQuery("Insert Into Ap_Section_AdNew (Sec_ID , Sec_Nm ,Ints, Menu_Nm , Usr_Id) VAlues (" & N & " ,N'----------------------------------------------------'   ,0,N'" & mItem.DropDownItems(j).Name & "' , '" & Usr & "' )")

                End If

                If mItem.DropDownItems(j).Text <> "" Then
                    mSubItem = mItem.DropDownItems(j)
                    For a = 0 To mSubItem.DropDownItems.Count - 1
                        NM3 = NM3 + 1
                        If mSubItem.DropDownItems(a).Text <> "" Then
                            ExecuteNonQuery("Insert Into Ap_Section_AdNew (Sec_ID , Sec_Nm ,Ints, Menu_Nm , Usr_Id) VAlues (" & N & " ,N' - " & mSubItem.DropDownItems(a).Text & "'   ,0,N'" & mSubItem.DropDownItems(a).Name & "' , '" & Usr & "' )")
                        Else
                            ExecuteNonQuery("Insert Into Ap_Section_AdNew (Sec_ID , Sec_Nm ,Ints, Menu_Nm , Usr_Id) VAlues (" & N & " ,N'   ------------------------------------------------'   ,0,N'" & mSubItem.DropDownItems(a).Name & "' , '" & Usr & "' )")

                        End If

                    Next a
                End If
            Next j
        Next i


    End Sub



    Public Sub KK1()
        For i = 0 To FmMain.MenuStrip1.Items.Count - 1
            mItem = CType(FmMain.MenuStrip1.Items.Item(i), ToolStripMenuItem)
            For j = 0 To mItem.DropDownItems.Count - 1

            Next j
        Next i

    End Sub
End Module
