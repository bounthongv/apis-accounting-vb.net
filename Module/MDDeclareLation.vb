Module MDDeclareLation

    Public MuGen_jn As String = "SELECT    *,0 as txtName,0 as txtAdd1,0 as txtAdd2,0 as txtTel,0 as txtFax,0 as txtPeriod,0 as txtDesc,Acc_Code.Name_L AS ExAcc_Name FROM  gen_jn       INNER JOIN Acc_Code ON gen_jn.ac_code = Acc_Code.Ac_Code WHERE Book <>''  "
    'file.""
    Public AG As Integer
    Public sd, za As String
    Public et As Boolean
    Public MdCode As String
    Public MdDate As Date
    Public Certify_Id As String
    Public MDateSeting As String
    Public SLF As String
    Public FmNme As String = ""
    Public Key_Code, T2, t3, t4, t5, rate As String
    ''============Register======================
    Public MDSeriel As String
    Public MDStarDate As String
    Public MDSerielID As String
    Public MDSerielAge As String
    Public MDUsingDay As String
    Public MDAgeRegistrtion As String
    Public RightsByLaw As String
    Public RightsByLaws As String
    Public MuFmName As String
    '===========================
    Public MULockey As Boolean

    Public MDServerPassword2 As String
    Public MDServerUser2 As String
    Public MDDatabaName2 As String
    Public MDServerName2 As String

    Public MDSeriaAccess As String
    Public MDSeriaCom As String
    Public MDForMain As String
    Public MDServerPassword As String
    Public MDServerUser As String
    Public MDDatabaName As String
    Public MDServerName As String

    Public MWorkSetting As Date = Format(MWorkSetting, "dd/MM/yyyy")
    Public MuSubOff, MuSubOff2, INternet As String
    Public VSysError As Boolean
    Public MDEDIT As Integer
    Public MDSaving, MdSearchDataList As String
    Public MDDeposit As String
    Public MDWithdr As String
    Public MDCustID As String
    Public MDDepositID As String
    Public MDWithdrID As String
    Public MdDepositRemain As Double
    Public STFDate, STTDate, STCode, STOpen, STDeposit, STWithdr, Curr_Last, STRemain, MDDepositNm As String
    Public SRFDate, SRTDate, SRCode, SRCurr, SRRate As String
    Public OfficeID, OfficeNm, MDReceipt, MDTranfering As String
    '===========================================so
    Public R As Integer
    Public L As Integer
    Public MDInvoiceNo As String
    Public MDInvoice_RefNo As String
    Public DDDR, DDDRN As String
    Public MDInvoiceDT As String
    Public MDSearchAcccode As String
    Public AccName As String
    Public AccNamee As String
    Public AccId As String
    Public MDTHB As String
    Public MDUSD As String
    Public MDEUR As String
    Public MDoff_id As String
    Public MdAtv As Boolean

    Public FgR As Integer
    Public FgC As Integer
    Public MDReceiptType As String
    '============MdSearchDataList====
    'Public MdSearchDataList As String

    '============ConnectServer=======================
    'Public MDServerPassword As String
    'Public MDServerUser As String
    'Public MDDatabaName As String
    'Public MDServerName As String
    'Public VSysError As Boolean
    '============User=======================
    'Public MDUserName As String
    '============Compy=========
    'Public MDCompanyName As String
    Public MDComTel As String
    Public MDComAddress As String
    Public MDCompany, FormMain As String
    '============SaerchType=======================
    Public MDSearchType As String
    '====================Login====================
    Public MUserID, MUserName, MPws, MSection, MPermit, Mpermiss As String
    '=========================Security===========================
    Public MD1 As Integer
    Public Off_Id, Sub_ID, SUB_Nm As String
End Module
