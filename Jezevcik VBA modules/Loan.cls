VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Loan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pSkp_loan As Long
'Private pAmt_balance_principal As Double
'Private pAmt_interest_overdue As Double
'Private pIr_standard As Double
'Private pIr_bonus_actual As Double
Private pLoan_status As String
Private pLoan_wfstatus As String
'Private pOu_total As Double
'Private pSu_total As Double
Private pAmt_instalment As Double
Private pCode_loan_number As String
Private pDate_loan_utilization As Date
Private pDate_payoff As Date
Private pDate_loan_close As Date
Private pDate_loan_write_off As Date
Private pAmt_loan_principal_paid As Double
Private pAmt_loan_principal_total As Double
Private pText_bank_account_number As String
Private pDtime_signature As Date
Private pText_appendix_number As String
Private Interests As Collection
' Private Interests_new As Collection   ' for debugging only..
Private pDATE_INSTALMENT_FIRST As Date
Private pDATE_INSTALMENT_LAST As Date
Private pDATE_LOAN_CLOSE_PLAN As Date
Private pId_person As Long

Private pRate_Interest_Standard As Double
Private pRate_Interest_Repo As Double 'RISK-726
Private pAmt_loan_limit As Double 'RISK-726
Private pRate_eir_standard As Double 'RISK-726
Private pFlag_payed_one_instalment As Boolean ' TRUE, pokud nebyla prvni splatka uhrazena v plne vysi  ' RISK-1237 ' jez.v.5.9



Public Property Get AmtInstalment() As Long
     AmtInstalment = pAmt_instalment
 End Property
Public Property Get skp_loan() As Long
     skp_loan = pSkp_loan
 End Property

Public Property Get id_person() As Long
     id_person = pId_person
 End Property
 
 Public Property Get code_loan_number() As String
     code_loan_number = pCode_loan_number
 End Property
'
' Public Property Get Amt_balance_principal() As Double
'     Amt_balance_principal = pAmt_balance_principal
' End Property
 
 Public Property Get Date_payoff() As Date
     Date_payoff = pDate_payoff
 End Property
 Public Property Get Date_loan_utilization() As Date
     Date_loan_utilization = pDate_loan_utilization
 End Property
 Public Property Get Date_loan_close() As Date
     Date_loan_close = pDate_loan_close
 End Property
 Public Property Get DateInstalmentLast() As Date
     DateInstalmentLast = pDATE_INSTALMENT_LAST
 End Property
  Public Property Get DateInstalmentFirst() As Date
     DateInstalmentFirst = pDATE_INSTALMENT_FIRST
 End Property
 Public Property Get DateLoanClosePlan() As Date
     DateLoanClosePlan = pDATE_LOAN_CLOSE_PLAN
 End Property
 Public Property Get Date_loan_write_off() As Date
     Date_loan_write_off = pDate_loan_write_off
 End Property
 Public Property Get Amt_loan_principal_total() As Double
     Amt_loan_principal_total = pAmt_loan_principal_total
 End Property
 Public Property Get Text_bank_account_number() As String
     Text_bank_account_number = pText_bank_account_number
 End Property
 Public Property Get Dtime_signature() As Date
      Dtime_signature = pDtime_signature
 End Property
 Public Property Get Text_appendix_number() As String
      Text_appendix_number = pText_appendix_number
 End Property
 Public Property Get Amt_loan_principal_paid() As Double
      Amt_loan_principal_paid = pAmt_loan_principal_paid
 End Property
 
 Public Property Get Amt_loan_principal_unpaid() As Double
      Amt_loan_principal_unpaid = pAmt_loan_principal_total - pAmt_loan_principal_paid
 End Property
 
 
' Public Property Get Amt_interest_overdue() As Double
'     Amt_interest_overdue = pAmt_interest_overdue
' End Property
 
' Public Property Get Ir_standard() As Double
'     Ir_standard = pIr_standard
' End Property
'
' Public Property Get Ir_bonus_actual() As Double
'     Ir_bonus_actual = pIr_bonus_actual
' End Property
 
 Public Property Get Loan_status() As String
     Loan_status = pLoan_status
 End Property

 Public Property Get Loan_wfstatus() As String
     Loan_wfstatus = pLoan_wfstatus
 End Property
 
Public Property Get Rate_Interest_Standard()
'Nacita sa az po zavolani property If pRate_Interest_Standard = 0 Then GetRatesInterests
    Rate_Interest_Standard = pRate_Interest_Standard
End Property

Public Property Get Rate_Interest_Repo() As Double
'Nacita sa az po zavolani property If pRate_Interest_Repo = 0 Then GetRatesInterests
   Rate_Interest_Repo = pRate_Interest_Repo
End Property

Public Property Get Amt_loan_limit() As Double
    Amt_loan_limit = pAmt_loan_limit
End Property

Public Property Get Rate_eir_standard() As Double
    Rate_eir_standard = pRate_eir_standard
End Property

Public Property Get blnUhrazena_alespon_jedna_splatka() As Boolean
    ' alias k property Flag_payed_one_instalment
    blnUhrazena_alespon_jedna_splatka = pFlag_payed_one_instalment
End Property

Public Property Get Flag_payed_one_instalment() As Boolean
    ' alias k property blnUhrazena_alespon_jedna_splatka
    Flag_payed_one_instalment = pFlag_payed_one_instalment
End Property

''''   END OF PROPERTIES



'Private Sub GetRatesInterests()
''Aktualna vyska obchodneho uroku k uctu.
'
'Dim sql8 As String, rst8 As Variant
'
'    sql8 = "select amt_loan_limit, rate_interest_standard, rate_interest_repo from " & schema & ".sp7 where code_loan_number  = '" & pCode_loan_number & "'"
'    Set rst8 = New ADODB.Recordset
'    rst8.Open sql8, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'        If Not rst8.EOF Then
'            'amt_loan_limit = rst8(0)
'            pRate_Interest_Standard = rst8![Rate_Interest_Standard]
'            pRate_Interest_Repo = rst8![Rate_Interest_Repo] 'RISK-726
'        End If
'    rst8.Close
'
'End Sub
 

Public Function GetInterests(ByVal index As Variant) As Interest
'odkazani se na jednotlive uvery

    Set GetInterests = Interests.item(index)

End Function
 

Public Function GetCntInterests() As Long
'kolik ma klient uveru

    GetCntInterests = Interests.count

End Function

 
Public Sub initialiaze(skp_loan As Long)

Dim sql2 As String
Dim sql_zadna_splatka As String
Dim rst_zadna_splatka As ADODB.Recordset

If cs_HDWMN.State = 1 Then
prihlasen:


    'pSkp_loan = Skp_loan
    
   '    sql1 = "select l.code_loan_number, l.amt_balance_principal, l.amt_interest_overdue, l.datum_zespl, l.rate_interest_standard ,l.rate_interest_bonus_actual ,l.name_loan_status, l.name_loan_wfstatus, l.ou_total, l.su_total from " & schema & ".intervaly l where sysdate <= dtime_valid_to and skp_loan = " & pSkp_loan
'
'
'
'   Set rst1 = New ADODB.Recordset
'   rst1.Open sql1, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'
'   If Not rst1.EOF Then
'        rst1.MoveFirst
'
'        pCode_loan_number = rst1![Code_loan_number]
'        If IsNull(rst1![Amt_balance_principal]) Then
'        Else
'            pAmt_balance_principal = rst1![Amt_balance_principal]
'        End If
'        If IsNull(rst1![Amt_interest_overdue]) Then
'        Else
'            pAmt_interest_overdue = rst1![Amt_interest_overdue]
'        End If
'        pDate_payoff = rst1![datum_zespl]
'        pIr_standard = rst1![rate_interest_standard]
'        If IsNull(rst1![rate_interest_bonus_actual]) Then
'        Else
'            pIr_bonus_actual = rst1![rate_interest_bonus_actual]
'        End If
'
'        pOu_total = rst1![Ou_total]
'        pSu_total = rst1![Su_total]
'
'   End If
'   rst1.Close
    
    sql2 = "select l.skp_loan, l.amt_instalment, l.name_loan_status, l.name_loan_wfstatus, " & _
            " L.Code_loan_number, trunc(l.DATE_LOAN_UTILIZATION) as DATE_LOAN_UTILIZATION, trunc(l.DATE_PAYOFF) as DATE_PAYOFF, trunc(l.DATE_LOAN_CLOSE) as DATE_LOAN_CLOSE, trunc(l.DATE_LOAN_WRITTEN_OFF) as DATE_LOAN_WRITTEN_OFF, nvl(l.AMT_LOAN_PRINCIPAL_TOTAL,0) as AMT_LOAN_PRINCIPAL_TOTAL, l.TEXT_BANK_ACCOUNT_NUMBER, trunc(l.DTIME_SIGNATURE) as DTIME_SIGNATURE, l.text_appendix_number, trunc(l.DATE_INSTALMENT_FIRST) as DATE_INSTALMENT_FIRST, trunc(l.DATE_INSTALMENT_LAST) as DATE_INSTALMENT_LAST, l.date_loan_close_plan " & _
            ", nvl(amt_loan_principal_paid, 0) as amt_loan_principal_paid " & _
            ", amt_loan_limit, rate_interest_standard, rate_eir_standard, rate_interest_repo, flag_payed_one_instalment " & _
            " from " & schema & ".loan l where skp_loan = " & skp_loan
    
    Set rst2 = New ADODB.Recordset
    rst2.Open sql2, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
        If Not rst2.EOF Then
            rst2.MoveFirst
            pSkp_loan = rst2![skp_loan]
            pLoan_status = "" & rst2![NAME_LOAN_STATUS] & ""
            pLoan_wfstatus = "" & rst2![name_loan_wfstatus] & ""
                           
            If IsNull(rst2![Amt_Instalment]) Then
            Else
                pAmt_instalment = rst2![Amt_Instalment]
            End If
            
            pCode_loan_number = "" & rst2![code_loan_number] & ""
            If Not IsNull(rst2![Date_loan_utilization]) Then pDate_loan_utilization = rst2![Date_loan_utilization]
            If Not IsNull(rst2![Date_payoff]) Then pDate_payoff = rst2![Date_payoff]
            If Not IsNull(rst2![Date_loan_close]) Then pDate_loan_close = rst2![Date_loan_close]
            If Not IsNull(rst2![Date_loan_written_off]) Then pDate_loan_write_off = rst2![Date_loan_written_off]
            pAmt_loan_principal_total = rst2![Amt_loan_principal_total]
            pText_bank_account_number = "" & rst2![Text_bank_account_number] & ""
            If Not IsNull(rst2![Dtime_signature]) Then pDtime_signature = rst2![Dtime_signature]
            pText_appendix_number = "" & rst2![Text_appendix_number] & ""
            pAmt_loan_principal_paid = rst2![Amt_loan_principal_paid]
            If Not IsNull(rst2![DATE_INSTALMENT_FIRST]) Then pDATE_INSTALMENT_FIRST = rst2![DATE_INSTALMENT_FIRST]
            If Not IsNull(rst2![date_instalment_last]) Then pDATE_INSTALMENT_LAST = rst2![date_instalment_last]
            If Not IsNull(rst2![date_loan_close_plan]) Then pDATE_LOAN_CLOSE_PLAN = rst2![date_loan_close_plan]
            
            If IsNull(rst2![Amt_loan_limit]) Or IsNull(rst2![Rate_Interest_Standard]) Or IsNull(rst2![Rate_eir_standard]) Or IsNull(rst2![Rate_Interest_Repo]) Then
               'MsgBox "V datab·zi chybÌ v˝öka ˙roku k ˙vÏru " & pCode_loan_number & ". JezevËÌk bu mÏl b˝t OK, ale öablony si s tÌm neporadÌ.", vbCritical, "ChybÌ ˙daj"
               'Debug.Assert 1 = 0 ' pokud se zastavi kvuli REPO, asi to je OK, na DWH evidujeme chybu. Proto nevyhazuji hlasku.
            End If
            If Not IsNull(rst2![Amt_loan_limit]) Then pAmt_loan_limit = rst2![Amt_loan_limit]
            If Not IsNull(rst2![Rate_Interest_Standard]) Then pRate_Interest_Standard = rst2![Rate_Interest_Standard]
            If Not IsNull(rst2![Rate_eir_standard]) Then pRate_eir_standard = rst2![Rate_eir_standard]
            If Not IsNull(rst2![Rate_Interest_Repo]) Then pRate_Interest_Repo = rst2![Rate_Interest_Repo]
            If Not IsNull(rst2![Flag_payed_one_instalment]) Then  ' RISK-1237
                If (rst2![Flag_payed_one_instalment] = "Y") Then
                    pFlag_payed_one_instalment = True
                Else
                    pFlag_payed_one_instalment = False
                End If
            End If
            
        End If
    rst2.Close
    
    
    ' Call find_interests(pSkp_loan)
    Call find_interests_new(pSkp_loan) ' RISK-726!!! Toto je podstatne
    
'''    ''' napocet flag_payed_one_instalment, pokud funguje z tabulky LOAN tak tohle vyhodit. jez.v.5.9
'''    sql_zadna_splatka = "SELECT SKP_LOAN FROM " & schema & ".V_NEUHRAZENA_ZADNA_SPLATKA a WHERE a.SKP_LOAN = " & pSkp_loan
'''    Set rst_zadna_splatka = New ADODB.Recordset
'''    rst_zadna_splatka.Open sql_zadna_splatka, cs_HDWMN, adOpenForwardOnly, adLockBatchOptimistic
'''        If rst_zadna_splatka.RecordCount > 0 Then
'''            pFlag_payed_one_instalment = True
'''        Else
'''            pFlag_payed_one_instalment = False
'''        End If
'''    rst_zadna_splatka.Close
    

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



Public Sub find_interests_new(skp_loan As Long)
' autor: FIlip Antol
' datum: 5.3.2018, RISK-726
' nov˝ n·poËet objekt˘ Interest - naËtou se vöechny spl·tky nar·z
' p¯edeöl˝ zp˘sob je nelogick˝, neefektivnÌ a oökliv˝ (probÌha spousta duplicitnÌch dotaz˘ do datab·ze
' pokud bude fungovat spravne, vyhodit puvodni find_interests. Interest.initialize bude taky zbytecny, ale nechal bych ho tam..

' nejd¯Ìve obchodnÌ a sankËnÌ ˙roky, pak kapitalizovan˝ zvl·öù


Dim sqlinterest As String
Dim rstInterest As ADODB.Recordset
Dim Interest As Object
Dim i As Integer


    Set Interests = New Collection
    
    
    
    
    ' OBCHODNI A SANKCNI UROKY (OU a SU)
    
    
    
    sqlinterest = "select skp_loan, AMT_BALANCE_PRINCIPAL, RATE_INTEREST, UROK_TYP, START_INTERVAL, KONEC_INTERVAL, AMT_UROK " & _
                 " FROM " & schema & ".v_interest a " & _
                 " WHERE a.Skp_loan = " & skp_loan & ""
                   
    Set rstInterest = New ADODB.Recordset
    rstInterest.Open sqlinterest, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
       
        If Not rstInterest.EOF Then
            
            rstInterest.MoveFirst
            
            For i = 1 To rstInterest.RecordCount ' poslednÌ OU je nulov˝..
            
                If Not rstInterest.EOF Then
                    
                    'vytvo¯Ì nov˝ objekt Prislusenstvi, zavol· init a p¯id· jej do kolekce
                    Set Interest = New Interest
                    
                    If Not (rstInterest![urok_typ] = "OU" Or rstInterest![urok_typ] = "SU") Then
                        Debug.Assert 1 = 0
                        MsgBox "Jez-error 564339+:" & vbCrLf & "Nespr·vna forma ˙roku, poöli prinstscreen analytik˘m.", vbCritical, "Chyba v datech"
                    End If
                    
                    
                    '  naplnime hodnoty do objekt˘ Interest
                    '  pokud OU, plnÌ se jenom Rate_interest_standard
                    '  pokud SU, plnÌ se jenom Rate_interest_repo
                    Call Interest.initialize_values(Amt_balance_principal:=rstInterest![Amt_balance_principal], _
                                                                  urok_typ:=rstInterest![urok_typ], _
                                                            start_interval:=rstInterest![start_interval], _
                                                            konec_interval:=rstInterest![konec_interval], _
                                                                  amt_urok:=rstInterest![amt_urok])
                                                                  
                    Interests.Add item:=Interest
                    
                    
                    rstInterest.MoveNext
                    
                End If
            
            Next i
        
        End If
        
    rstInterest.Close
    
    
    
    
    
    ' KAPITALIZOVANY UROK (KU)
    
    
    
    sqlinterest = " SELECT a.SKP_LOAN, a.AMT_BALANCE_PRINCIPAL, a.RATE_INTEREST_STANDARD, a.RATE_INTEREST_REPO, a.UROK_TYP, a.START_INTERVAL, a.KONEC_INTERVAL, a.AMT_UROK " & _
                    " FROM " & schema & ".v_interest_kapit_urok a " & _
                    " WHERE a.SKP_LOAN = " & skp_loan
                   
    Set rstInterest = New ADODB.Recordset
    rstInterest.Open sqlinterest, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
        Debug.Assert rstInterest.RecordCount < 2
        
        
        If Not rstInterest.EOF Then
        
            rstInterest.MoveFirst
            Set Interest = New Interest
            Call Interest.initialize_values(Amt_balance_principal:=rstInterest![Amt_balance_principal], _
                                                          urok_typ:=rstInterest![urok_typ], _
                                                    start_interval:=rstInterest![start_interval], _
                                                    konec_interval:=rstInterest![konec_interval], _
                                                          amt_urok:=rstInterest![amt_urok])
            Interests.Add item:=Interest
            
        End If
        
    rstInterest.Close


End Sub



'najde a vytvori objekty - p¯isluöenstvÌ k ˙vÏr˘m
Public Sub find_interests(skp_loan As Long)
Dim i As Integer
Dim sqlinterest As String
Dim rstInterest As ADODB.Recordset
Dim Interest As Object
'Dim a_skp_loan() As Long, cnt_loans As Integer
    Set Interests = New Collection
    Debug.Assert 1 = 0 ' FA tahle SUB by se uz nemela pouzivat. RISK-726
        
    '''  zjistime vsechny uroky
    sqlinterest = "SELECT skp_loan, AMT_BALANCE_PRINCIPAL, RATE_INTEREST_STANDARD, UROK_TYP, START_INTERVAL, KONEC_INTERVAL, AMT_UROK " & _
                " FROM ( " & _
                " select skp_loan, AMT_BALANCE_PRINCIPAL, RATE_INTEREST_STANDARD, 'OU' as UROK_TYP, start_Ou AS START_INTERVAL, konec_Ou AS KONEC_INTERVAL, OU AS AMT_UROK " & _
                " from " & schema & ".intervaly " & _
                " UNION " & _
                " select skp_loan, AMT_BALANCE_PRINCIPAL, RATE_INTEREST_STANDARD, 'SU' as UROK_TYP, start_Su AS START_INTERVAL, konec_Su AS KONEC_INTERVAL, SU AS AMT_UROK " & _
                " from " & schema & ".intervaly " & _
                " Union " & _
                " select l.skp_loan, i.AMT_BALANCE_PRINCIPAL, i.RATE_INTEREST_STANDARD,'KU' as UROK_TYP, nvl(p.last_full_pay, l.DATE_LOAN_UTILIZATION) as START_INTERVAL, l.date_payoff as KONEC_INTERVAL, i.amt_interest_overdue as AMT_UROK" & _
                " from " & schema & ".loan l " & _
                " left join " & schema & ".intervaly i on l.skp_loan = i.skp_loan " & _
                " left join ( " & _
                "     select skp_loan, max(datum_platby) as last_full_pay " & _
                "            from (" & _
                "                select skp_loan, uver_cislo, datum_splatky, amt_instalment_total, max(datum_platby) as datum_platby, sum(amt_payment) as amt_payment from " & schema & ".platby" & _
                "                group by skp_loan, uver_cislo, datum_splatky, amt_instalment_total " & _
                "                ) " & _
                "            WHERE AMT_INSTALMENT_TOTAL = amt_payment " & _
                "            group by skp_loan " & _
                "            ) p on p.skp_loan = l.skp_loan " & _
                "    WHERE i.dtime_valid_to > sysdate " & _
                "      )" & _
                "  WHERE skp_loan = " & skp_loan & " order by  urok_typ desc "
                
 
    Set rstInterest = New ADODB.Recordset
    rstInterest.Open sqlinterest, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
       If Not rstInterest.EOF Then
            
            rstInterest.MoveFirst
            For i = 1 To rstInterest.RecordCount - 1
                'vytvo¯Ì nov˝ objekt Prislusenstvi, zavol· init a p¯id· jej do kolekce
                Set Interest = New Interest
                Call Interest.initialize(i, skp_loan, "OU")
                Interests.Add item:=Interest
                rstInterest.MoveNext
            Next
            
            If i = rstInterest.RecordCount Then
                Set Interest = New Interest
                Call Interest.initialize(i, skp_loan, "KU") 'vytvo¯Ì nov˝ objekt Prislusenstvi, zavol· init a p¯id· jej do kolekce
                Interests.Add item:=Interest
                rstInterest.MoveNext
            End If
       End If
    rstInterest.Close
       
End Sub



