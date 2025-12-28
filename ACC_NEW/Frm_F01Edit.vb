Public Class Frm_F01Edit

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Frm_F01Edit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Loadtext()
    End Sub
    Private Sub Loadtext()
        Call LoadSqlData(" select * from RPT_F01 where ItemID='F01' ", RSC)
        If RSC.RecordCount <> 0 Then
            txtNo1.Text = Trim(RSC.Fields("No1").Value.ToString)
            txtNo2.Text = Trim(RSC.Fields("No2").Value.ToString)
            txtNo3.Text = Trim(RSC.Fields("No3").Value.ToString)
            '===========
            txtNm1.Text = Trim(RSC.Fields("Nm1").Value.ToString)
            txtNm2.Text = Trim(RSC.Fields("Nm2").Value.ToString)
            txtNm3.Text = Trim(RSC.Fields("Nm3").Value.ToString)
            txtNm4.Text = Trim(RSC.Fields("Nm4").Value.ToString)
            txtNm5.Text = Trim(RSC.Fields("Nm5").Value.ToString)
            txtNm6.Text = Trim(RSC.Fields("Nm6").Value.ToString)
            txtNm7.Text = Trim(RSC.Fields("Nm7").Value.ToString)
            txtNm8.Text = Trim(RSC.Fields("Nm8").Value.ToString)
            txtNm9.Text = Trim(RSC.Fields("Nm9").Value.ToString)
            '===========
            txtAccNo1.Text = Trim(RSC.Fields("AccNo1").Value.ToString)
            txtAccNo2.Text = Trim(RSC.Fields("AccNo2").Value.ToString)
            txtAccNo3.Text = Trim(RSC.Fields("AccNo3").Value.ToString)
            '===========
            TxtHolder1.Text = Trim(RSC.Fields("Holder1").Value.ToString)
            TxtHolder2.Text = Trim(RSC.Fields("Holder2").Value.ToString)
            TxtHolder3.Text = Trim(RSC.Fields("Holder3").Value.ToString)
            TxtHolder4.Text = Trim(RSC.Fields("Holder4").Value.ToString)
            TxtHolder5.Text = Trim(RSC.Fields("Holder5").Value.ToString)
            TxtHolder6.Text = Trim(RSC.Fields("Holder6").Value.ToString)
            TxtHolder7.Text = Trim(RSC.Fields("Holder7").Value.ToString)
            TxtHolder8.Text = Trim(RSC.Fields("Holder8").Value.ToString)
            TxtHolder9.Text = Trim(RSC.Fields("Holder9").Value.ToString)
            TxtHolder10.Text = Trim(RSC.Fields("Holder10").Value.ToString)
            TxtHolder11.Text = Trim(RSC.Fields("Holder11").Value.ToString)
            TxtHolder12.Text = Trim(RSC.Fields("Holder12").Value.ToString)
            TxtHolder13.Text = Trim(RSC.Fields("Holder13").Value.ToString)
            TxtHolder14.Text = Trim(RSC.Fields("Holder14").Value.ToString)
            TxtHolder15.Text = Trim(RSC.Fields("Holder15").Value.ToString)
            '===========
            TxtBoard1.Text = Trim(RSC.Fields("Board1").Value.ToString)
            TxtBoard2.Text = Trim(RSC.Fields("Board2").Value.ToString)
            TxtBoard3.Text = Trim(RSC.Fields("Board3").Value.ToString)
            TxtBoard4.Text = Trim(RSC.Fields("Board4").Value.ToString)
            TxtBoard5.Text = Trim(RSC.Fields("Board5").Value.ToString)
            TxtBoard6.Text = Trim(RSC.Fields("Board6").Value.ToString)
            TxtBoard7.Text = Trim(RSC.Fields("Board7").Value.ToString)
            TxtBoard8.Text = Trim(RSC.Fields("Board8").Value.ToString)
            TxtBoard9.Text = Trim(RSC.Fields("Board9").Value.ToString)
            TxtBoard10.Text = Trim(RSC.Fields("Board10").Value.ToString)
            TxtBoard11.Text = Trim(RSC.Fields("Board11").Value.ToString)
            TxtBoard12.Text = Trim(RSC.Fields("Board12").Value.ToString)
            TxtBoard13.Text = Trim(RSC.Fields("Board13").Value.ToString)
            TxtBoard14.Text = Trim(RSC.Fields("Board14").Value.ToString)
            TxtBoard15.Text = Trim(RSC.Fields("Board15").Value.ToString)
            '===========
            TxtDirector1.Text = Trim(RSC.Fields("Director1").Value.ToString)
            TxtDirector2.Text = Trim(RSC.Fields("Director2").Value.ToString)
            TxtDirector3.Text = Trim(RSC.Fields("Director3").Value.ToString)
            TxtDirector4.Text = Trim(RSC.Fields("Director4").Value.ToString)
            TxtDirector5.Text = Trim(RSC.Fields("Director5").Value.ToString)
            TxtDirector6.Text = Trim(RSC.Fields("Director6").Value.ToString)
            TxtDirector7.Text = Trim(RSC.Fields("Director7").Value.ToString)
            '===========
            TxtAudit1.Text = Trim(RSC.Fields("Audit1").Value.ToString)
            TxtAudit2.Text = Trim(RSC.Fields("Audit2").Value.ToString)
            TxtAudit3.Text = Trim(RSC.Fields("Audit3").Value.ToString)
            TxtAudit4.Text = Trim(RSC.Fields("Audit4").Value.ToString)
            TxtAudit5.Text = Trim(RSC.Fields("Audit5").Value.ToString)
            TxtAudit6.Text = Trim(RSC.Fields("Audit6").Value.ToString)
            TxtAudit7.Text = Trim(RSC.Fields("Audit7").Value.ToString)
            '===========
            TxtCredit1.Text = Trim(RSC.Fields("Credit1").Value.ToString)
            TxtCredit2.Text = Trim(RSC.Fields("Credit2").Value.ToString)
            TxtCredit3.Text = Trim(RSC.Fields("Credit3").Value.ToString)
            TxtCredit4.Text = Trim(RSC.Fields("Credit4").Value.ToString)
            TxtCredit5.Text = Trim(RSC.Fields("Credit5").Value.ToString)
            TxtCredit6.Text = Trim(RSC.Fields("Credit6").Value.ToString)
            TxtCredit7.Text = Trim(RSC.Fields("Credit7").Value.ToString)
            '===========
            TxtBusiness1.Text = Trim(RSC.Fields("Business1").Value.ToString)
            TxtBusiness2.Text = Trim(RSC.Fields("Business2").Value.ToString)
            TxtBusiness3.Text = Trim(RSC.Fields("Business3").Value.ToString)
            TxtBusiness4.Text = Trim(RSC.Fields("Business4").Value.ToString)
            TxtBusiness5.Text = Trim(RSC.Fields("Business5").Value.ToString)
            '===========
            txtregist1.Text = Trim(RSC.Fields("regist1").Value.ToString)
            txtregist2.Text = Trim(RSC.Fields("regist2").Value.ToString)
            txtregist3.Text = Trim(RSC.Fields("regist3").Value.ToString)
            txtregist4.Text = Trim(RSC.Fields("regist4").Value.ToString)
            txtregist5.Text = Trim(RSC.Fields("regist5").Value.ToString)
            '===========
            TxtContact1.Text = Trim(RSC.Fields("Contact1").Value.ToString)
            TxtContact2.Text = Trim(RSC.Fields("Contact2").Value.ToString)
            TxtContact3.Text = Trim(RSC.Fields("Contact3").Value.ToString)
            TxtContact4.Text = Trim(RSC.Fields("Contact4").Value.ToString)
            TxtContact5.Text = Trim(RSC.Fields("Contact5").Value.ToString)
            TxtContact6.Text = Trim(RSC.Fields("Contact6").Value.ToString)
            TxtContact7.Text = Trim(RSC.Fields("Contact7").Value.ToString)
        End If
    End Sub

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        Call LoadSqlData(" select * from RPT_F01 where ItemID='F01' ", RSC)
        If RSC.RecordCount = 0 Then
            CNN.Execute("INSERT INTO RPT_F01(ItemID) values (N'F01') ")
        End If
        CNN.Execute(" UPDATE RPT_F01 set " & _
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