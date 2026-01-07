Public Class Frm_F01Edit

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Frm_F01Edit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Loadtext()
    End Sub
    Private Sub Loadtext()
        Dim dt As DataTable = DbHelper.GetDataTable(" select * from RPT_F01 where ItemID='F01' ")
        If dt.Rows.Count > 0 Then
            txtNo1.Text = DbHelper.GetStr(dt.Rows(0)("No1"))
            txtNo2.Text = DbHelper.GetStr(dt.Rows(0)("No2"))
            txtNo3.Text = DbHelper.GetStr(dt.Rows(0)("No3"))
            '===========
            txtNm1.Text = DbHelper.GetStr(dt.Rows(0)("Nm1"))
            txtNm2.Text = DbHelper.GetStr(dt.Rows(0)("Nm2"))
            txtNm3.Text = DbHelper.GetStr(dt.Rows(0)("Nm3"))
            txtNm4.Text = DbHelper.GetStr(dt.Rows(0)("Nm4"))
            txtNm5.Text = DbHelper.GetStr(dt.Rows(0)("Nm5"))
            txtNm6.Text = DbHelper.GetStr(dt.Rows(0)("Nm6"))
            txtNm7.Text = DbHelper.GetStr(dt.Rows(0)("Nm7"))
            txtNm8.Text = DbHelper.GetStr(dt.Rows(0)("Nm8"))
            txtNm9.Text = DbHelper.GetStr(dt.Rows(0)("Nm9"))
            '===========
            txtAccNo1.Text = DbHelper.GetStr(dt.Rows(0)("AccNo1"))
            txtAccNo2.Text = DbHelper.GetStr(dt.Rows(0)("AccNo2"))
            txtAccNo3.Text = DbHelper.GetStr(dt.Rows(0)("AccNo3"))
            '===========
            TxtHolder1.Text = DbHelper.GetStr(dt.Rows(0)("Holder1"))
            TxtHolder2.Text = DbHelper.GetStr(dt.Rows(0)("Holder2"))
            TxtHolder3.Text = DbHelper.GetStr(dt.Rows(0)("Holder3"))
            TxtHolder4.Text = DbHelper.GetStr(dt.Rows(0)("Holder4"))
            TxtHolder5.Text = DbHelper.GetStr(dt.Rows(0)("Holder5"))
            TxtHolder6.Text = DbHelper.GetStr(dt.Rows(0)("Holder6"))
            TxtHolder7.Text = DbHelper.GetStr(dt.Rows(0)("Holder7"))
            TxtHolder8.Text = DbHelper.GetStr(dt.Rows(0)("Holder8"))
            TxtHolder9.Text = DbHelper.GetStr(dt.Rows(0)("Holder9"))
            TxtHolder10.Text = DbHelper.GetStr(dt.Rows(0)("Holder10"))
            TxtHolder11.Text = DbHelper.GetStr(dt.Rows(0)("Holder11"))
            TxtHolder12.Text = DbHelper.GetStr(dt.Rows(0)("Holder12"))
            TxtHolder13.Text = DbHelper.GetStr(dt.Rows(0)("Holder13"))
            TxtHolder14.Text = DbHelper.GetStr(dt.Rows(0)("Holder14"))
            TxtHolder15.Text = DbHelper.GetStr(dt.Rows(0)("Holder15"))
            '===========
            TxtBoard1.Text = DbHelper.GetStr(dt.Rows(0)("Board1"))
            TxtBoard2.Text = DbHelper.GetStr(dt.Rows(0)("Board2"))
            TxtBoard3.Text = DbHelper.GetStr(dt.Rows(0)("Board3"))
            TxtBoard4.Text = DbHelper.GetStr(dt.Rows(0)("Board4"))
            TxtBoard5.Text = DbHelper.GetStr(dt.Rows(0)("Board5"))
            TxtBoard6.Text = DbHelper.GetStr(dt.Rows(0)("Board6"))
            TxtBoard7.Text = DbHelper.GetStr(dt.Rows(0)("Board7"))
            TxtBoard8.Text = DbHelper.GetStr(dt.Rows(0)("Board8"))
            TxtBoard9.Text = DbHelper.GetStr(dt.Rows(0)("Board9"))
            TxtBoard10.Text = DbHelper.GetStr(dt.Rows(0)("Board10"))
            TxtBoard11.Text = DbHelper.GetStr(dt.Rows(0)("Board11"))
            TxtBoard12.Text = DbHelper.GetStr(dt.Rows(0)("Board12"))
            TxtBoard13.Text = DbHelper.GetStr(dt.Rows(0)("Board13"))
            TxtBoard14.Text = DbHelper.GetStr(dt.Rows(0)("Board14"))
            TxtBoard15.Text = DbHelper.GetStr(dt.Rows(0)("Board15"))
            '===========
            TxtDirector1.Text = DbHelper.GetStr(dt.Rows(0)("Director1"))
            TxtDirector2.Text = DbHelper.GetStr(dt.Rows(0)("Director2"))
            TxtDirector3.Text = DbHelper.GetStr(dt.Rows(0)("Director3"))
            TxtDirector4.Text = DbHelper.GetStr(dt.Rows(0)("Director4"))
            TxtDirector5.Text = DbHelper.GetStr(dt.Rows(0)("Director5"))
            TxtDirector6.Text = DbHelper.GetStr(dt.Rows(0)("Director6"))
            TxtDirector7.Text = DbHelper.GetStr(dt.Rows(0)("Director7"))
            '===========
            TxtAudit1.Text = DbHelper.GetStr(dt.Rows(0)("Audit1"))
            TxtAudit2.Text = DbHelper.GetStr(dt.Rows(0)("Audit2"))
            TxtAudit3.Text = DbHelper.GetStr(dt.Rows(0)("Audit3"))
            TxtAudit4.Text = DbHelper.GetStr(dt.Rows(0)("Audit4"))
            TxtAudit5.Text = DbHelper.GetStr(dt.Rows(0)("Audit5"))
            TxtAudit6.Text = DbHelper.GetStr(dt.Rows(0)("Audit6"))
            TxtAudit7.Text = DbHelper.GetStr(dt.Rows(0)("Audit7"))
            '===========
            TxtCredit1.Text = DbHelper.GetStr(dt.Rows(0)("Credit1"))
            TxtCredit2.Text = DbHelper.GetStr(dt.Rows(0)("Credit2"))
            TxtCredit3.Text = DbHelper.GetStr(dt.Rows(0)("Credit3"))
            TxtCredit4.Text = DbHelper.GetStr(dt.Rows(0)("Credit4"))
            TxtCredit5.Text = DbHelper.GetStr(dt.Rows(0)("Credit5"))
            TxtCredit6.Text = DbHelper.GetStr(dt.Rows(0)("Credit6"))
            TxtCredit7.Text = DbHelper.GetStr(dt.Rows(0)("Credit7"))
            '===========
            TxtBusiness1.Text = DbHelper.GetStr(dt.Rows(0)("Business1"))
            TxtBusiness2.Text = DbHelper.GetStr(dt.Rows(0)("Business2"))
            TxtBusiness3.Text = DbHelper.GetStr(dt.Rows(0)("Business3"))
            TxtBusiness4.Text = DbHelper.GetStr(dt.Rows(0)("Business4"))
            TxtBusiness5.Text = DbHelper.GetStr(dt.Rows(0)("Business5"))
            '===========
            txtregist1.Text = DbHelper.GetStr(dt.Rows(0)("regist1"))
            txtregist2.Text = DbHelper.GetStr(dt.Rows(0)("regist2"))
            txtregist3.Text = DbHelper.GetStr(dt.Rows(0)("regist3"))
            txtregist4.Text = DbHelper.GetStr(dt.Rows(0)("regist4"))
            txtregist5.Text = DbHelper.GetStr(dt.Rows(0)("regist5"))
            '===========
            TxtContact1.Text = DbHelper.GetStr(dt.Rows(0)("Contact1"))
            TxtContact2.Text = DbHelper.GetStr(dt.Rows(0)("Contact2"))
            TxtContact3.Text = DbHelper.GetStr(dt.Rows(0)("Contact3"))
            TxtContact4.Text = DbHelper.GetStr(dt.Rows(0)("Contact4"))
            TxtContact5.Text = DbHelper.GetStr(dt.Rows(0)("Contact5"))
            TxtContact6.Text = DbHelper.GetStr(dt.Rows(0)("Contact6"))
            TxtContact7.Text = DbHelper.GetStr(dt.Rows(0)("Contact7"))
        End If
    End Sub

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        Dim dt As DataTable = DbHelper.GetDataTable(" select * from RPT_F01 where ItemID='F01' ")
        If dt.Rows.Count = 0 Then
            DbHelper.ExecuteNonQuery("INSERT INTO RPT_F01(ItemID) values (N'F01') ")
        End If
        DbHelper.ExecuteNonQuery(" UPDATE RPT_F01 set " & _
                     " No1=N'" & txtNo1.Text & "', " & _
                     " No2=N'" & txtNo2.Text & "', " & _
                     " No3=N'" & txtNo3.Text & "', " & _
                     " Nm1=N'" & txtNm1.Text & "', " & _
                     " Nm2=N'" & txtNm2.Text & "', " & _
                     " Nm3=N'" & txtNm3.Text & "', " & _
                     " Nm4=N'" & txtNm4.Text & "', " & _
                     " Nm5=N'" & txtNm5.Text & "', " & _
                     " Nm6=N'" & txtNm6.Text & "', " & _
                     " Nm7=N'" & txtNm7.Text & "', " & _
                     " Nm8=N'" & txtNm8.Text & "', " & _
                     " Nm9=N'" & txtNm9.Text & "', " & _
           " AccNo1=N'" & txtAccNo1.Text & "', " & _
           " AccNo2=N'" & txtAccNo2.Text & "', " & _
           " AccNo3=N'" & txtAccNo3.Text & "', " & _
               " Holder1=N'" & TxtHolder1.Text & "', " & _
               " Holder2=N'" & TxtHolder2.Text & "', " & _
               " Holder3=N'" & TxtHolder3.Text & "', " & _
               " Holder4=N'" & TxtHolder4.Text & "', " & _
               " Holder5=N'" & TxtHolder5.Text & "', " & _
               " Holder6=N'" & TxtHolder6.Text & "', " & _
               " Holder7=N'" & TxtHolder7.Text & "', " & _
               " Holder8=N'" & TxtHolder8.Text & "', " & _
               " Holder9=N'" & TxtHolder9.Text & "', " & _
               " Holder10=N'" & TxtHolder10.Text & "', " & _
               " Holder11=N'" & TxtHolder11.Text & "', " & _
               " Holder12=N'" & TxtHolder12.Text & "', " & _
               " Holder13=N'" & TxtHolder13.Text & "', " & _
               " Holder14=N'" & TxtHolder14.Text & "', " & _
               " Holder15=N'" & TxtHolder15.Text & "', " & _
                        " Board1=N'" & TxtBoard1.Text & "', " & _
                        " Board2=N'" & TxtBoard2.Text & "', " & _
                        " Board3=N'" & TxtBoard3.Text & "', " & _
                        " Board4=N'" & TxtBoard4.Text & "', " & _
                        " Board5=N'" & TxtBoard5.Text & "', " & _
                        " Board6=N'" & TxtBoard6.Text & "', " & _
                        " Board7=N'" & TxtBoard7.Text & "', " & _
                        " Board8=N'" & TxtBoard8.Text & "', " & _
                        " Board9=N'" & TxtBoard9.Text & "', " & _
                        " Board10=N'" & TxtBoard10.Text & "', " & _
                        " Board11=N'" & TxtBoard11.Text & "', " & _
                        " Board12=N'" & TxtBoard12.Text & "', " & _
                        " Board13=N'" & TxtBoard13.Text & "', " & _
                        " Board14=N'" & TxtBoard14.Text & "', " & _
                        " Board15=N'" & TxtBoard15.Text & "', " & _
                                " Director1=N'" & TxtDirector1.Text & "', " & _
                                " Director2=N'" & TxtDirector2.Text & "', " & _
                                " Director3=N'" & TxtDirector3.Text & "', " & _
                                " Director4=N'" & TxtDirector4.Text & "', " & _
                                " Director5=N'" & TxtDirector5.Text & "', " & _
                                " Director6=N'" & TxtDirector6.Text & "', " & _
                                " Director7=N'" & TxtDirector7.Text & "', " & _
                           " Audit1=N'" & TxtAudit1.Text & "', " & _
                           " Audit2=N'" & TxtAudit2.Text & "', " & _
                           " Audit3=N'" & TxtAudit3.Text & "', " & _
                           " Audit4=N'" & TxtAudit4.Text & "', " & _
                           " Audit5=N'" & TxtAudit5.Text & "', " & _
                           " Audit6=N'" & TxtAudit6.Text & "', " & _
                           " Audit7=N'" & TxtAudit7.Text & "', " & _
                                       " Credit1=N'" & TxtCredit1.Text & "', " & _
                                       " Credit2=N'" & TxtCredit2.Text & "', " & _
                                       " Credit3=N'" & TxtCredit3.Text & "', " & _
                                       " Credit4=N'" & TxtCredit4.Text & "', " & _
                                       " Credit5=N'" & TxtCredit5.Text & "', " & _
                                       " Credit6=N'" & TxtCredit6.Text & "', " & _
                                       " Credit7=N'" & TxtCredit7.Text & "', " & _
                                   " Business1=N'" & TxtBusiness1.Text & "', " & _
                                   " Business2=N'" & TxtBusiness2.Text & "', " & _
                                   " Business3=N'" & TxtBusiness3.Text & "', " & _
                                   " Business4=N'" & TxtBusiness4.Text & "', " & _
                                   " Business5=N'" & TxtBusiness5.Text & "', " & _
                                         " regist1=N'" & txtregist1.Text & "', " & _
                                         " regist2=N'" & txtregist2.Text & "', " & _
                                         " regist3=N'" & txtregist3.Text & "', " & _
                                         " regist4=N'" & txtregist4.Text & "', " & _
                                         " regist5=N'" & txtregist5.Text & "', " & _
                                " Contact1=N'" & TxtContact1.Text & "', " & _
                                " Contact2=N'" & TxtContact2.Text & "', " & _
                                " Contact3=N'" & TxtContact3.Text & "', " & _
                                " Contact4=N'" & TxtContact4.Text & "', " & _
                                " Contact5=N'" & TxtContact5.Text & "', " & _
                                " Contact6=N'" & TxtContact6.Text & "', " & _
                                 " Contact7=N'" & TxtContact7.Text & "' " & _
                    " WHERE ItemID=N'F01' ")
        MsgBox("Finish")
    End Sub
End Class