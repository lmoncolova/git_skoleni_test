VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pSkpBankAccount As Long
Private pText_bank_account_number As String
Private pAmt_balance As Double
Private pAccount_type As String
Private pAccount_status As String
Private pAccount_wfstatus As String
Private pCode_currency As String
Private pAmtGovernmentHoldSum  As Double
Private pCnt_days_overdue_tol_A As Integer
Private pCnt_days_overdue_tol_B As Integer
Private pId_person As Long


Public Property Get skpBankAccount() As Long
     skpBankAccount = pSkpBankAccount
 End Property

Public Property Get id_person() As Long
     id_person = pId_person
 End Property
 
 Public Property Get Text_bank_account_number() As String
     Text_bank_account_number = pText_bank_account_number
 End Property
 
 Public Property Get Amt_balance() As Double
     Amt_balance = pAmt_balance
 End Property
 
 Public Property Get Account_type() As String
     Account_type = pAccount_type
 End Property
 
 Public Property Get Account_status() As String
     Account_status = pAccount_status
 End Property

 Public Property Get Account_wfstatus() As String
     Account_wfstatus = pAccount_wfstatus
 End Property
 
 Public Property Get Code_currency() As String
     Code_currency = pCode_currency
 End Property
  
 Public Property Get AmtGovernmentHoldSum() As Double
     AmtGovernmentHoldSum = pAmtGovernmentHoldSum
 End Property
 
 Public Property Get Cnt_days_overdue_tol_A() As Integer
     Cnt_days_overdue_tol_A = pCnt_days_overdue_tol_A
 End Property
 
 Public Property Get Cnt_days_overdue_tol_B() As Integer
     Cnt_days_overdue_tol_B = pCnt_days_overdue_tol_B
 End Property
 
 
Public Sub initialiaze(skp_bank_account)
Dim pSkp_bank_account As Variant
Dim sql As String

If cs_HDWMN.State = 1 Then
prihlasen:


    pSkp_bank_account = skp_bank_account
    
    sql = "select u.skp_bank_account, u.cislo_uctu, u.typ_uctu, u.date_period, u.amt_balance, u.code_currency, u.name_bank_account_status, u.name_bank_account_wfstatus, u.AMT_GOVERNMENT_HOLD_SUM, u.CNT_DAYS_OVERDUE_TOL_A, u.CNT_DAYS_OVERDUE_TOL_B from " & schema & ".ucet u where skp_bank_account = " & pSkp_bank_account
    
    
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
        rst1.MoveFirst
                       
        pText_bank_account_number = rst1![cislo_uctu]
        If IsNull(rst1![Amt_balance]) Then
        Else
            pAmt_balance = rst1![Amt_balance]
        End If
        pSkpBankAccount = rst1![skp_bank_account]
        pAccount_type = rst1![typ_uctu]
        pAccount_status = rst1![NAME_BANK_ACCOUNT_STATUS]
        pAccount_wfstatus = rst1![name_bank_account_wfstatus]
        pCode_currency = rst1![Code_currency]
        If IsNull(rst1![AMT_GOVERNMENT_HOLD_SUM]) Then
        Else: pAmtGovernmentHoldSum = rst1![AMT_GOVERNMENT_HOLD_SUM]
        End If
        If IsNull(rst1![Cnt_days_overdue_tol_A]) Then
        Else: pCnt_days_overdue_tol_A = rst1![Cnt_days_overdue_tol_A]
        End If
        If IsNull(rst1![Cnt_days_overdue_tol_B]) Then
        Else: pCnt_days_overdue_tol_B = rst1![Cnt_days_overdue_tol_B]
        End If
   End If
   rst1.Close
    
    

    If err_connection Then
     err_connection = False
     Call re_login(Main.username.text, Main.password.text)
     If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
    End If
    
Else

    Call re_login(Main.username.text, Main.password.text)
    If cs_HDWMN.State = 1 Then
        GoTo prihlasen
    Else: Call logout
    End If
End If
End Sub
