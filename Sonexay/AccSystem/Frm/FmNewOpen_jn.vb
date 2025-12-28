Public Class FmNewOpen_jn
    Dim MdCode_dr As String
    Dim MdCode_cr, MuSubOffOp As String


    Dim Amount_dr As Double
    Dim Amount_cr As Double
    Dim Amt_dr As Double
    Dim Amt_cr As Double
    Private Sub loadOffice_User()
        Off_Usr.Items.Clear()
        LoadSqlData("select sub_id , off_add2  from  Ap_office Where sub_id <> '00-00' And Substring(sub_id,4,2) <> '00'  Order by sub_id", RSC)
        With RSC
            Do Until .EOF = True
                Off_Usr.Items.Add((.Fields("sub_id").Value) & " " & (.Fields("off_add2").Value))
                .MoveNext()
            Loop
        End With
        Off_Usr.Text = FmLogin.Sub_Company.Text



        'Off_Usr.Enabled = True
        If txtCode_dr.Enabled = False Then
            'Off_Usr.Enabled = False
            LoadSqlData("select sub_id , off_add2  from  Ap_office Where sub_id = '" & FmOpen_jn_List.FG.get_TextMatrix(FmOpen_jn_List.FG.Row, 13) & "'", RSC)
            If RSC.RecordCount <> 0 Then
                Off_Usr.Text = (RSC.Fields("sub_id").Value) & " " & (RSC.Fields("off_add2").Value)
            End If
        End If


    End Sub
    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        If txtCode_dr.Enabled = True Then
            'Call LoadSqlData("SELECT ac_code FROM open_jn WHERE ac_code = '" & txtCode_dr.Text & txtCode_cr.Text & "'  and company = '" & Microsoft.VisualBasic.Left(Off_Usr.Text, 5) & "'  And date_Work = '" & Format(CDate("01/01/" & Year(DtmYearDate.Value)), "dd/MM/yyyy") & "' ", RSC)
            Call LoadSqlData("SELECT ac_code FROM open_jn WHERE Curr=N'" & Cmb.Text & "' and  ac_code =N'" & txtCode_dr.Text & txtCode_cr.Text & "'  and company = '" & Microsoft.VisualBasic.Left(Off_Usr.Text, 5) & "'  And date_Work = '" & Format(CDate("01/01/" & Year(DtmYearDate.Value)), "dd/MM/yyyy") & "' ", RSC)

            If RSC.RecordCount > 0 Then
                MsgBox("ເລກລະຫັດ : " & txtCode_dr.Text & txtCode_cr.Text & " " & Cmb.Text & " ມີໃນຖານຂໍ້ມູນແລ້ວ ກະລຸນາປ່ຽນ!", MsgBoxStyle.OkOnly)
                If RSC.State = ConnectionState.Open Then RSC.Close()
                Exit Sub
            End If
            SdaveData()
            MessageBox.Show("ການບັນທຶກຮຽບຮ້ອຍ")
        Else
            CNN.Execute("DELETE FROM Open_jn WHERE ac_code='" & FmOpen_jn_List.FG.get_TextMatrix(FmOpen_jn_List.FG.Row, 4) & "' And Company='" & LAbel9.Text & "' and cnt='" & FmOpen_jn_List.FG.get_TextMatrix(FmOpen_jn_List.FG.Row, 14) & "' ")
            SdaveData()
            MessageBox.Show("ການແກ້ໄຂຮຽບຮ້ອຍ")
        End If

    End Sub
    Private Sub SdaveData()
        If txtCode_dr.Text <> "" Then
            Amount_dr = txtAmount.Text
            Amt_dr = txtAmt.Text
            Amount_cr = "0.00"
            Amt_cr = "0.00"
        Else
            Amount_dr = "0.00"
            Amt_dr = "0.00"
            Amount_cr = txtAmount.Text
            Amt_cr = txtAmt.Text
        End If


        CNN.Execute("INSERT INTO open_jn( date_work,  code_dr , code_cr ,ac_code, ac_name, ac_namee  ,ac_type, ac_typee, amount_dr, amount_cr ,curr, rate , amt_dr , amt_cr ,my_lock,last_user ,last_update,company) " & _
                          "Values('" & Format(CDate("1/1/" & Year(DtmYearDate.Value)), "dd-MM-yyyy") & "', N'" & txtCode_dr.Text & "', N'" & txtCode_cr.Text & "', '" & txtCode_dr.Text & txtCode_cr.Text & "', N'" & txtAccName.Text & "', N'" & txtAccNameE.Text & "', N'" & txtAccType.Text & "', N'" & txtAccTypee.Text & "', '" & CDbl(Amount_dr) & "', '" & CDbl(Amount_cr) & "', '" & Cmb.Text & "', '" & CDbl(txtRate.Text) & "', '" & CDbl(Amt_dr) & "', '" & CDbl(Amt_cr) & "', '" & "0" & "', '" & MUserID & "', '" & Format(Now.Date, "yyyy/MM/dd") & "', '" & Microsoft.VisualBasic.Left(Off_Usr.Text, 5) & "')")



    End Sub

    Private Sub txtCode_cr_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCode_cr.KeyPress
        If e.KeyChar = Chr(13) Then
            txtAmount.Focus()
        End If
    End Sub

    Private Sub txtCode_cr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCode_cr.TextChanged
        txtCode_dr.Clear()
        txtAccType.Clear()
        txtAccTypee.Clear()
        txtAccName.Clear()
        txtAccNameE.Clear()
        LoadSqlData("SELECT * FROM Acc_Code WHERE ac_code = N'" & txtCode_cr.Text & "'", RSC)
        With RSC
            Do Until .EOF = True
                txtAccType.Text = Trim(.Fields("Acc_Type").Value)
                txtAccTypee.Text = Trim(.Fields("Acc_TypeE").Value)
                txtAccName.Text = Trim(.Fields("Name_L").Value)
                txtAccNameE.Text = Trim(.Fields("Name_E").Value)
                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub txtCode_dr_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCode_dr.KeyPress
        If e.KeyChar = Chr(13) Then
            txtAmount.Focus()
        End If
    End Sub



    Private Sub txtCode_dr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCode_dr.TextChanged
        txtCode_cr.Clear()
        txtAccType.Clear()
        txtAccTypee.Clear()
        txtAccName.Clear()
        txtAccNameE.Clear()
        LoadSqlData("SELECT * FROM Acc_Code WHERE ac_code = N'" & txtCode_dr.Text & "'", RSC)
        With RSC
            Do Until .EOF = True
                txtAccType.Text = Trim(.Fields("Acc_Type").Value)
                txtAccTypee.Text = Trim(.Fields("Acc_TypeE").Value)
                txtAccName.Text = Trim(.Fields("Name_L").Value)
                txtAccNameE.Text = Trim(.Fields("Name_E").Value)

                .MoveNext()
            Loop
        End With
    End Sub
    Private Sub Load_Open_jn()
        Dim RSK As New ADODB.Recordset

        Dim s As String = "SELECT * FROM Open_jn WHERE ac_code = N'" & txtCode_dr.Text & txtCode_cr.Text & "' And company = '" & LAbel9.Text & "' And Year(Date_Work) = " & Format(CDate(FmOpen_jn_List.FG.get_TextMatrix(FmOpen_jn_List.FG.Row, 1)), "yyyy") & " "
        LoadSqlData(s, RSK)
        With RSK
            Do Until .EOF = True
                DtmYearDate.Value = Trim(.Fields("Date_Work").Value)
                txtAccType.Text = Trim(.Fields("ac_type").Value)
                txtAccTypee.Text = Trim(.Fields("ac_typee").Value)
                txtAccName.Text = Trim(.Fields("ac_name").Value)
                txtAccNameE.Text = Trim(.Fields("ac_namee").Value)
                Off_Usr.Text = Trim(.Fields("Company").Value)
                'MsgBox(Trim(.Fields("Company").Value))
                Cmb.Text = Trim(.Fields("Curr").Value)
                If CDbl(Trim(.Fields("amount_dr").Value)) = 0 Then
                    txtAmount.Text = CDbl(Trim(.Fields("amount_cr").Value))
                Else
                    txtAmount.Text = CDbl(Trim(.Fields("amount_dr").Value))
                End If

                txtAmt.Text = CDbl(Trim(.Fields("amt_dr").Value))
                txtRate.Text = Trim(.Fields("rate").Value)
                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub FmNewOpen_jn_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        FmOpen_jn_List.LoadListFG()
        MdAtv = False
    End Sub
    Public Sub LoadCurr()
        Dim Comm As ADODB.Command
        Dim rsat As New ADODB.Recordset
        Comm = New ADODB.Command
        Comm.ActiveConnection = CNN
        Cmb.Items.Clear()
        Comm.CommandText = "SELECT Curr FROM Ap_RateSeting WHERE Curr <> '" & "" & " order by Curr'"
        rsat = Comm.Execute
        If rsat.RecordCount <> 0 Then
            While Not rsat.EOF()
                Cmb.Items.Add(Trim(rsat.Fields("Curr").Value))
                rsat.MoveNext()
            End While
        End If
        FormatText()
    End Sub
    Private Sub FmNewOpen_jn_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtAmt.ReadOnly = True
        BtnSave.Enabled = True
        BntNew.Enabled = True
        Call loadOffice_User()
        Call RateSetting()

        Cmb.Items.Clear()
        Call load_Cmb(" SELECT Curr  FROM Curr_For_Rate  ORDER BY cnt ", "Curr", Cmb)
        If Cmb.Items.Count > 0 Then
            Cmb.SelectedIndex = 0
        End If

        If txtCode_dr.Enabled = True Then
            AddNew()
            FormatText()
            'Call loadRate()
            'LoadCurr()
        Else
            Load_Open_jn()
            FormatText()
           
        End If

        SetControlText(Me)
   
    End Sub

    Private Sub loadRate()
        LoadSqlData("select * from Ap_RateSeting where Curr='" & Cmb.Text & "'", RSC)
        With RSC
            Do Until .EOF = True
                txtRate.Text = Trim(.Fields("Rate").Value)

                .MoveNext()
            Loop
        End With
    End Sub

    Private Sub FormatText()
        txtAmt.Text = CDbl(txtAmount.Text) * CDbl(txtRate.Text)
        txtAmt.Text = Format(CDbl(txtAmt.Text), "#,##0.00")
        txtRate.Text = Format(CDbl(txtRate.Text), "#,##0.00")
        txtAmount.Text = Format(CDbl(txtAmount.Text), "#,##0.00")

    End Sub






    Private Sub CmbCerrency_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmb.SelectedIndexChanged
        'loadRate()

        If txtAmt.Text <> "" Then
        Dim rs As New ADODB.Recordset
        Call LoadSqlData("Select * From Curr_For_Rate Where   Curr =N'" & Trim(Cmb.Text) & "'", rs)
        If rs.RecordCount > 0 Then
            txtcurr_name2.Text = Trim(rs("Curr_name").Value.ToString)
        End If

            MDRate_DT = " and rate_dt<='" & Format(DtmYearDate.Value, "yyyy-MM-dd") & "'  "
        SS_Curr = " and AP_Rate_history.Curr =N'" & Cmb.Text & "' "
        Call RateSetting()
        txtRate.Text = Format(MD_Rate, "#,##0.00")


            'Call loadRate()
            FormatText()
        End If
    End Sub

    Private Sub BntNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BntNew.Click
        Call AddNew()

    End Sub
    Private Sub AddNew()
        txtCode_dr.Enabled = True
        txtCode_cr.Enabled = True
        BtnSearch_dr.Enabled = True
        BtnSearch_cr.Enabled = True
        txtCode_dr.Clear()
        txtCode_cr.Clear()
        txtAccType.Clear()
        txtAccTypee.Clear()
        txtAccName.Clear()
        txtAccNameE.Clear()
        txtAmount.Text = "0.00"
        txtAmt.Text = "0.00"
        txtRate.Text = "0.00"
        DtmYearDate.Text = ""
        LAbel9.Text = ""
        DtmYearDate.Value = "01/01/" & DtmYearDate.Value.Year
        Call FormatText()
    End Sub
    Private Sub txtAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAmount.KeyPress
        If e.KeyChar = Chr(13) Then
            txtAmt.Text = CDbl(txtAmount.Text) * CDbl(txtRate.Text)
            Call FormatText()
        End If
    End Sub



    Private Sub BtnSearch_dr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch_dr.Click
        txtCode_cr.Clear()
        fmShartOfAccDetail.txtSty.Text = "NewOpen_jn_dr"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()

    End Sub

    Private Sub BtnSearch_cr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch_cr.Click
        txtCode_dr.Clear()
        fmShartOfAccDetail.txtSty.Text = "NewOpen_jn_cr"
        MDSearchAcccode = ""
        fmShartOfAccDetail.ShowDialog()
    End Sub




    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub txtAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAmount.LostFocus
        txtAmt.Text = CDbl(txtAmount.Text) * CDbl(txtRate.Text)
        Call FormatText()
    End Sub

    Private Sub txtAmount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmount.TextChanged

    End Sub

    Private Sub DtmStartDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtAccType_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAccType.TextChanged

    End Sub

    Private Sub Label27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label27.Click
        'DtmYearDate.Text = 2012
    End Sub

    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click
        txtAmt.ReadOnly = False
    End Sub
End Class