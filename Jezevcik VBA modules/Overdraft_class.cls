VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Overdraft_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private datum_blokace As Date
' Private datum_zesp As Date ' tento by mal byt na Klientovi
' private datum_zaloby as Date ' zaloba je tiez na klienta


''' takhle by se mohli jmenovat nektere public properties - jakozto ALIAS k nazvum sloupecku / promennych nize
'Private datum_uzavreni_smlouvy As Date
'Private datum_cerpani As Date
'Private amt_nesplacena_jistina_k_zesp As Double
'Private amt_nesplacena_jistina_k_zalobe As Double
'Private amt_nesplacena_jistina_dnes As Double



' sloupecky tabule OVERDRAFT  (nepotrebne se vyhodi)
Private skp_bank_account As Long
Private skp_bank_account_debit As Long
Private skp_overdraft As Long
Private code_overdraft_number As String
Private amt_overdraft_limit As Double
Private date_overdraft_open As Date
Private date_overdraft_utilization As Date
Private skp_overdraft_status As Long
Private name_overdraft_status As String
Private name_overdraft_wfstatus As String
Private amt_overdraft As Double
Private amt_overdraft_overdue As Double
Private amt_overdraft_interest_to_pay As Double
Private amt_interest_overdue As Double
Private amt_interest_paid As Double
Private amt_interest_paid_total As Double
Private amt_interest_total As Double
Private amt_interest_last As Double
Private rate_penalty As Double
Private rate_penalty_interest As Double
Private amt_payoff As Double
Private amt_payoff_principal As Double
Private amt_payoff_interest As Double
Private skp_currency As Long

Private amt_posledni_kladny_zustatek As Double ' z tabule OVERDRAFT_POSLEDNI_ZUSTATEK
