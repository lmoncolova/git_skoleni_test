VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Interest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iAmt_balance_principal As Double
' iRate_interest_standard aj iRate_interest_repo vyhadzujem z objektu Interest, pretoze sa nachadza v objekte Loan. Tu nema co robit. Jez.v.5.3, RISK-726
'Private iRate_interest_standard As Double
'Private iRate_interest_repo As Double 'RISK-726
Private iUrok_typ As String
Private iStart_interval As Date
Private iKonec_interval As Date
Private iAmt_urok As Double


Public Property Get AmtUrok() As Double
    AmtUrok = iAmt_urok
End Property
'Public Property Get RateInterestStandard() As Double
'    RateInterestStandard = iRate_interest_standard
'End Property
'Public Property Get RateInterestRepo() As Double
'    RateInterestRepo = iRate_interest_repo
'End Property
Public Property Get UrokTyp() As String
   UrokTyp = iUrok_typ
End Property
Public Property Get StartInterval() As Date
   StartInterval = iStart_interval
End Property
Public Property Get KonecInterval() As Date
   KonecInterval = iKonec_interval - 1
End Property
Public Property Get AmtBalancePrincipal() As Double
   AmtBalancePrincipal = iAmt_balance_principal
End Property
 


Public Sub initialize_values(ByVal Amt_balance_principal As Double, _
                             ByVal urok_typ As String, _
                             ByVal start_interval As Date, _
                             ByVal konec_interval As Date, _
                             ByVal amt_urok As Double)

                             'ByVal Rate_interest_standard As Double, _
                             'ByVal Rate_Interest_Repo As Double, _

' autor: Filip Antol
' datum: 5.3.2018, RISK-726
' dotaz do databáze probíha v objektu Loan, metoda fint_interests_new()
' tohle je jenom dosazení hodnot.. Šlo by taky zadefinovat pro každou promìnnou Public Property Let, ale takhle lze nastavit jenom všechny naráz, a to mi pøijde vhodnìjší (nelze omylem zmìnit jednu hodnotu..)


    iAmt_balance_principal = Amt_balance_principal
    'iRate_interest_standard = Rate_interest_standard
    'iRate_interest_repo = Rate_Interest_Repo
    iUrok_typ = urok_typ
    iStart_interval = start_interval
    iKonec_interval = konec_interval
    iAmt_urok = amt_urok

End Sub


Public Sub initialize(i As Integer, skp_loan As Long, typ As String)
Dim sqlinterest As String
Dim rstInterest As ADODB.Recordset

    
    If typ = "KU" And 0 = 1 Then ' používáno do verze 5.2, RISK-726
''        sqlinterest = " select l.skp_loan, i.AMT_BALANCE_PRINCIPAL, i.RATE_INTEREST_STANDARD, nvl(i.rate_interest_repo, 999999) as rate_interest_repo, 'KU' as UROK_TYP, nvl(p.last_full_pay, l.DATE_LOAN_UTILIZATION) as START_INTERVAL, l.date_payoff as konec_INTERVAL, nvl(i.amt_interest_overdue, 0) as AMT_UROK " & _
''                      "  from " & schema & ".loan l " & _
''                      "    left join " & schema & ".intervaly i on l.skp_loan = i.skp_loan " & _
''                      "    left join (  " & _
''                      "                   select skp_loan, max(datum_splatky) as last_full_pay  " & _
''                      "                   from (  " & _
''                      "                        select skp_loan, uver_cislo, datum_splatky, amt_instalment_total, max(datum_platby) as datum_platby, sum(amt_payment) as amt_payment from " & schema & ".platby  " & _
''                      "                        group by skp_loan, uver_cislo, datum_splatky, amt_instalment_total " & _
''                      "                        )  " & _
''                      "                    where AMT_INSTALMENT_TOTAL = amt_payment  " & _
''                      "                   group by skp_loan   " & _
''                      "                    ) p on p.skp_loan = l.skp_loan  " & _
''                      "    where i.dtime_valid_to > sysdate  " & _
''                      "  and l.skp_loan = " & skp_loan
''
''             Set rstInterest = New ADODB.Recordset
''               rstInterest.Open sqlinterest, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''                If Not rstInterest.EOF Then
''                    rstInterest.MoveFirst
''                    iAmt_balance_principal = rstInterest![amt_balance_principal]
''                    iRate_interest_standard = rstInterest![rate_interest_standard]
''                    iRate_interest_repo = rstInterest![rate_interest_repo] 'RISK-726
''                    iUrok_typ = rstInterest![Urok_typ]
''                    iStart_interval = rstInterest![Start_interval]
''                    iKonec_interval = rstInterest![Konec_interval]
''                    iAmt_urok = rstInterest![amt_urok] 'FAXXX error null, cuid 224469, uver c.20
''
''                End If
''                rstInterest.Close
    
    
    ElseIf typ = "KU" Then ' od verze 5.3, RISK-726
    
'        sqlinterest = " SELECT a.SKP_LOAN, a.AMT_BALANCE_PRINCIPAL, a.RATE_INTEREST_STANDARD, a.RATE_INTEREST_REPO, a.UROK_TYP, a.START_INTERVAL, a.KONEC_INTERVAL, a.AMT_UROK " & _
'                      " FROM " & schema & ".v_interest_kapit_urok a " & _
'                      " WHERE a.SKP_LOAN = " & skp_loan
        sqlinterest = " SELECT a.SKP_LOAN, a.AMT_BALANCE_PRINCIPAL, a.UROK_TYP, a.START_INTERVAL, a.KONEC_INTERVAL, a.AMT_UROK " & _
                      " FROM " & schema & ".v_interest_kapit_urok a " & _
                      " WHERE a.SKP_LOAN = " & skp_loan
                            
             Set rstInterest = New ADODB.Recordset
               rstInterest.Open sqlinterest, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
                If Not rstInterest.EOF Then
                    rstInterest.MoveFirst
                    iAmt_balance_principal = rstInterest![Amt_balance_principal]
                    'iRate_interest_standard = rstInterest![Rate_interest_standard]
                    'iRate_interest_repo = rstInterest![Rate_Interest_Repo] 'RISK-726
                    iUrok_typ = rstInterest![urok_typ]
                    iStart_interval = rstInterest![start_interval]
                    iKonec_interval = rstInterest![konec_interval]
                    iAmt_urok = rstInterest![amt_urok] 'FAXXX error null, cuid 224469, uver c.20
                    
                End If
                rstInterest.Close
    
    
    
    ElseIf 1 = 0 Then ' používáno do verze 5.2, RISK-726
    
''            'VYPSÁNÍ VŠECH UROKÙ + INTERVALY XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''            'OBCHODNÍ UROK
''            'øazení OU a SU dle RISK-392
''            'XX RISK-726 este som sem nedal
''            sqlinterest = "select skp_loan, AMT_BALANCE_PRINCIPAL, RATE_INTEREST_STANDARD, UROK_TYP, START_INTERVAL, KONEC_INTERVAL, AMT_UROK from(" & _
''                          "select rownum as rn, skp_loan, AMT_BALANCE_PRINCIPAL, RATE_INTEREST_STANDARD, UROK_TYP, START_INTERVAL, KONEC_INTERVAL, AMT_UROK from ( " & _
''                          "SELECT skp_loan, AMT_BALANCE_PRINCIPAL, RATE_INTEREST_STANDARD, UROK_TYP, START_INTERVAL, KONEC_INTERVAL, AMT_UROK " & _
''                          "  FROM ( " & _
''                          "            select skp_loan, AMT_BALANCE_PRINCIPAL, RATE_INTEREST_STANDARD, 'OU' as UROK_TYP, start_Ou AS START_INTERVAL, konec_Ou AS KONEC_INTERVAL, OU AS AMT_UROK " & _
''                          "            FROM " & schema & ".intervaly " & _
''                          "  Union " & _
''                          "       select skp_loan, AMT_BALANCE_PRINCIPAL, RATE_INTEREST_STANDARD, 'SU' as UROK_TYP, start_Su AS START_INTERVAL, konec_Su AS KONEC_INTERVAL, SU AS AMT_UROK " & _
''                          "            FROM " & schema & ".intervaly " & _
''                          "      ) " & _
''                          "  WHERE Skp_loan = " & skp_loan & "  ORDER BY AMT_BALANCE_PRINCIPAL DESC, START_INTERVAL ASC, KONEC_INTERVAL DESC, urok_typ " & _
''                          ")) where rn = " & i
''
''             Set rstInterest = New ADODB.Recordset
''               rstInterest.Open sqlinterest, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''                If Not rstInterest.EOF Then
''                    rstInterest.MoveFirst
''                    iAmt_balance_principal = rstInterest![amt_balance_principal]
''                    iRate_interest_standard = rstInterest![rate_interest_standard]
''                    iRate_interest_repo = rstInterest![rate_interest_repo] 'RISK-726
''                    iUrok_typ = rstInterest![Urok_typ]
''                    iStart_interval = rstInterest![Start_interval]
''                    iKonec_interval = rstInterest![Konec_interval]
''                    iAmt_urok = rstInterest![amt_urok]
''
''                End If
''                rstInterest.Close
    
    
    Else    ' od verze 5.3, RISK-726
    
            'VYPSÁNÍ VŠECH UROKÙ + INTERVALY XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
            'OBCHODNÍ UROK
            'øazení OU a SU dle RISK-392
            'XX RISK-726 - vytvorene view v_interest
'            sqlinterest = "select skp_loan, AMT_BALANCE_PRINCIPAL, RATE_INTEREST, UROK_TYP, START_INTERVAL, KONEC_INTERVAL, AMT_UROK " & _
'                          " FROM " & schema & ".v_interest a " & _
'                          " WHERE a.Skp_loan = " & skp_loan & " AND a.row_number = " & i
            sqlinterest = "select skp_loan, AMT_BALANCE_PRINCIPAL, UROK_TYP, START_INTERVAL, KONEC_INTERVAL, AMT_UROK " & _
                          " FROM " & schema & ".v_interest a " & _
                          " WHERE a.Skp_loan = " & skp_loan & " AND a.row_number = " & i
                            
             Set rstInterest = New ADODB.Recordset
               rstInterest.Open sqlinterest, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
                If Not rstInterest.EOF Then
                    rstInterest.MoveFirst
                    iAmt_balance_principal = rstInterest![Amt_balance_principal]
                    
                    
                    iUrok_typ = rstInterest![urok_typ]
                    
                    'RISK-726
'                    If iUrok_typ = "OU" Then
'                        iRate_interest_standard = rstInterest![rate_interest]
'                    ElseIf iUrok_typ = "SU" Then
'                        iRate_interest_repo = rstInterest![rate_interest]
'                    Else
'                        Debug.Assert 1 = 0
'                        MsgBox "Jez-error 564339+:" & vbCrLf & "Nesprávna forma úroku, pošli prinstscreen analytikùm."
'                    End If
                    
                    iStart_interval = rstInterest![start_interval]
                    iKonec_interval = rstInterest![konec_interval]
                    iAmt_urok = rstInterest![amt_urok]
                    
                End If
                rstInterest.Close
                
    End If

End Sub

