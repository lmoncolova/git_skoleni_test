VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCuid As Long
Private pskp_person As Long
Private pId_person As Long
Private pName As String
'Private rc As String
Private pAddress_C As String
Private pAddress_P As String
Private pUlice_PA As String
Private pCP_PA As String
Private pPSC_PA As String
Private pMesto_PA As String
Private pZeme_PA As String
Private pUlice_CA As String
Private pCP_CA As String
Private pPSC_CA As String
Private pMesto_CA As String
Private pZeme_CA As String
Private pPin As String
Private pGender As String
Private pDateBirth As Date
Private pDateDeath As Date
Private pTextTitleBefore As String
Private pTextTitleAfter As String
Private pNationality As String
Private pTextContractNumber As String 'RISK-398
Private pDatePayoff As Date 'RISK-398
Private pPrimaryPhone As String ' RISK-403
Private pPrimaryEmail As String ' RISK-403
Private pInsolvencyStatus As String 'RISK-468
Private collAccounts As Collection
Private collLoans As Collection
Private pExeTitul As enum_exekucni_titul
Private pExeTitul_spisova_znacka As String
Private pDate_Contract_Signed As String


Private pos As Variant
Private sql As String


Private pSOP As SOP_class ' RISK-1101
Private pSOUD As SOUD_type
Private pOverdraft As Overdraft_class


Public Property Get Debetar() As Boolean
    Dim i As Integer
    Dim dblSumaNaUctech As Double
    
    If collAccounts Is Nothing Then
        Me.find_accounts
    End If
    
    For i = 1 To c.GetCnt_Accounts
        dblSumaNaUctech = dblSumaNaUctech + c.GetAccounts(i).Amt_balance
    Next i
    
    If dblSumaNaUctech < 0 Then
        Debetar = True
    Else
        Debetar = False
    End If
End Property
Public Property Get Nazev_slozky() As String
    ' to same co property Nazev_slozky_bez_lomitka, akurat S lomitkem na konci (pro skladani cesty k souboru nebo slozce)
    Nazev_slozky = c.name & " " & c.CUID & "\"
End Property
Public Property Get Nazev_slozky_bez_lomitka() As String
    ' to same co property Nazev_slozky, akurat bez lomitka na konci (pouziva se v nazvech sablon)
    Nazev_slozky_bez_lomitka = c.name & " " & c.CUID
End Property
Public Property Get Exekucni_titul() As enum_exekucni_titul
    ' pokud chces nazev: get_enum_exekucni_titul_Name(c.Exekucni_titul)
    ' nazev ale NIKDY nepouzivej v kodu! at je na jedinem miste, a to ve zminene funkci
    ' podminku vzdy nastav na ENUM_...
    Exekucni_titul = pExeTitul
End Property

Public Property Get ExeTitul_spisova_znacka() As String
    ' protoze osoba muze mit i vice nez jeden exekucni titul, je potreba k nemu uchovat i cislo spisove znacky
    ExeTitul_spisova_znacka = pExeTitul_spisova_znacka
End Property

Public Property Get SOP() As SOP_class ' RISK-1101
    If pSOP Is Nothing Then
        Set pSOP = New SOP_class
    End If
    Set SOP = pSOP
End Property

Public Property Get SOUD() As SOUD_type
    If pSOUD.skp_soud = 0 Then
        pSOUD = NaplnSoud
    End If
    SOUD = pSOUD
End Property

Public Property Get Overdraft() As Overdraft_class
    If pOverdraft Is Nothing Then
        Set pOverdraft = New Overdraft_class
    End If
    Set Overdraft = pOverdraft
End Property

Public Property Get Name_first() As String
    pos = InStr(pName, " ")
    Name_first = Mid(pName, pos + 1)
    
End Property

Public Property Get Name_last() As String
    pos = InStr(pName, " ")
    Name_last = Left(name, pos - 1)
End Property
    

'nechci, aby slo menit atributy zvenci - muze potom dojit k nekonzistenci dat, proto pouzivam jenom get

Public Property Get name() As String
     name = pName
 End Property
' Public Property Let Name(Value As String)
'     pName = Value
' End Property
 
 
Public Property Get CUID() As String
     CUID = pCuid
 End Property
' Public Property Let Cuid(Value As String)
'     pCuid = Cuid
' End Property
 
 
Public Property Get id_person() As String
     id_person = pId_person
 End Property
 
 Public Property Get skp_person() As String
     skp_person = pskp_person
 End Property
 
 Public Property Get Address_P_Ulice() As String
     Address_P_Ulice = pUlice_PA
 End Property
 Public Property Get Address_P_CP() As String
     Address_P_CP = pCP_PA
 End Property
 Public Property Get Address_P_PSC() As String
     Address_P_PSC = pPSC_PA
 End Property
 Public Property Get Address_P_Mesto() As String
     Address_P_Mesto = pMesto_PA
 End Property
 Public Property Get Address_P_Zeme() As String
     Address_P_Zeme = pZeme_PA
 End Property
 Public Property Get Address_C_Ulice() As String
     Address_C_Ulice = pUlice_CA
 End Property
 Public Property Get Address_C_CP() As String
     Address_C_CP = pCP_CA
 End Property
 Public Property Get Address_C_PSC() As String
     Address_C_PSC = pPSC_CA
 End Property
 Public Property Get Address_C_Mesto() As String
     Address_C_Mesto = pMesto_CA
 End Property
 Public Property Get Address_C_Zeme() As String
     Address_C_Zeme = pZeme_CA
 End Property
Public Property Get Address_P() As String
     Address_P = pAddress_P
 End Property

Public Property Get Address_C() As String
     Address_C = pAddress_C
 End Property
 Public Property Get Pin() As String
    If pPin <> "XNA" Then
        Pin = pPin
    Else: Pin = ""
    End If
 End Property
 Public Property Get PinLom() As String
    If pPin <> "XNA" Then
       PinLom = Left(pPin, 6) & "/" & Mid(pPin, 7)
    Else: PinLom = ""
    End If
 End Property
 Public Property Get TextBirthNumber() As String
     TextBirthNumber = Pin  'Jez.v.5.3
 End Property
 Public Property Get TextBirthNumberLomitko() As String
     TextBirthNumberLomitko = PinLom  'Jez.v.5.3
 End Property
 
 
 Public Property Get Gender() As String
     Gender = pGender
 End Property
 Public Property Get Date_Birth() As Date
     Date_Birth = pDateBirth
 End Property
 Public Property Get Text_Title_before() As String
     If pTextTitleBefore <> "XNA" Then
     Text_Title_before = pTextTitleBefore & Space(1)
     Else: Text_Title_before = ""
     End If
 End Property
 Public Property Get Text_Title_After() As String
     If pTextTitleAfter <> "XNA" Then
     Text_Title_After = "," & Space(1) & pTextTitleAfter
     Else: Text_Title_After = ""
     End If
 End Property
 Public Property Get Nationality() As String
     Nationality = pNationality
 End Property
 Public Property Get TextContractNumber() As String 'RISK-398
     TextContractNumber = pTextContractNumber
 End Property
  Public Property Get DatePayoff() As Date 'RISK-398
     DatePayoff = pDatePayoff
 End Property
 Public Property Get PrimaryPhone() As String ' RISK-403
     PrimaryPhone = pPrimaryPhone
 End Property
 Public Property Get PrimaryEmail() As String ' RISK-403
     PrimaryEmail = pPrimaryEmail
 End Property
 
 Public Property Get InsolvencyStatus() As String ' RISK-468
     InsolvencyStatus = pInsolvencyStatus
 End Property
 
 Public Property Get DateDeath() As Date
     DateDeath = pDateDeath
 End Property
 
 Public Property Get Date_Contract_Signed() As String
    Date_Contract_Signed = pDate_Contract_Signed
 End Property

Public Sub test_me(id_person As Long)
    
    Call Me.initialize(id_person)
    'debug.Assert
    
End Sub


'
' Public Property Let Id_person(Value As String)
'     pId_person = Id_person
' End Property


'odkazani se na jednotlive ucty
Public Function GetAccounts(ByVal index As Variant) As Account

      Set GetAccounts = collAccounts.item(index)

End Function
 
'kolik ma klient uctu

Public Function GetCnt_Accounts() As Long

      GetCnt_Accounts = collAccounts.count

End Function


'odkazani se na jednotlive uvery
Public Function GetLoans(ByVal index As Variant, Optional code_loan_number As String) As Loan
' druhy argument pouzijeme, ak nevieme cislo uveru, ale jeho nazov "HU000..." od verze 5.3
' ak ideme cez nazov, tak do indexu treba dat -1
' Pozor, funkce je rekurzivni (pocas debuggingu -> na horni liste je View -> Call Stack)

Dim i As Integer

    If index = -1 And code_loan_number <> "" Then
        For i = 1 To c.GetCnt_Loans
            If c.GetLoans(i).code_loan_number = code_loan_number Then ' rekurze, tady vola funkce sebe sama
                index = i
                Exit For
            End If
        Next i
    End If
    Set GetLoans = collLoans.item(index)

End Function
 
'kolik ma klient uveru

Public Function GetCnt_Loans() As Long

      GetCnt_Loans = collLoans.count

End Function

'kolik ma klient puvodni soucet splatek na zesplatnenych uverech

Public Function GetAmt_InstalmentBeforePayoff() As Long 'RISK-442
Dim AmtInstalmentSumBeforePayoff As Double
Dim clc As Integer

 AmtInstalmentSumBeforePayoff = 0
 For clc = 1 To collLoans.count

  If collLoans(clc).Loan_status = "Aktivní" And (collLoans(clc).Loan_wfstatus = "Zesplatnìný" Or collLoans(clc).Loan_wfstatus = "K odpisu" Or collLoans(clc).Loan_wfstatus = "Odepsaný bankou") Then
    AmtInstalmentSumBeforePayoff = AmtInstalmentSumBeforePayoff + collLoans(clc).AmtInstalment
  End If
 Next clc
 
 GetAmt_InstalmentBeforePayoff = AmtInstalmentSumBeforePayoff

End Function


 
'inicializacni skript pro naplneni zakladnich dat klienta
Public Sub initialize(idP As Long)
Dim sql As String
Dim sqlDem As String
Dim sqlContact As String
Dim SQLADR As String
Dim sql_datum_RS As String
Dim rst_datum_RS As ADODB.Recordset


If cs_HDWMN.State = 1 Then
prihlasen:
   ' z PERSON
   'RISK-398
   pId_person = idP 'RISK-468 - doplnìna insolvence
   sql = "select Cuid as cu, skp_person ,Name as n, Address_C, Address_P, text_contract_number, date_payoff, CODE_INSOLVENCY_CASE_STATUS, DATE_DEATH" & _
          " FROM " & schema & ".v_person_base where id_person = " & pId_person

   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
       If Not rst1.EOF Then
    
            pCuid = rst1![cu]
            pskp_person = rst1![skp_person]
            pName = "" & rst1![n] & ""
            pAddress_C = "" & rst1![Address_C] & ""
            pAddress_P = "" & rst1![Address_P] & ""
            'RISK-398
            If IsNull(rst1![text_contract_number]) Then
            Else
                pTextContractNumber = rst1![text_contract_number]
            End If
            pDatePayoff = rst1![Date_payoff]
            pInsolvencyStatus = "" & rst1![CODE_INSOLVENCY_CASE_STATUS] & "" 'RISK-468
            If Not IsNull(rst1![DATE_DEATH]) Then pDateDeath = "" & rst1![DATE_DEATH] & "" ' Jez.v.5.3
            
       End If
   rst1.Close
   
   
   ' z tabulky PERSON_DEMOGRAPHY
   sqlDem = "select gender, date_birth, TEXT_TITLE_BEFORE, TEXT_TITLE_AFTER, nationality, nvl(text_birth_number, 'XNA') as pin from " & schema & ".PERSON_DEMOGRAPHY where id_person = " & pId_person

   Set rst2 = New ADODB.Recordset
   rst2.Open sqlDem, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
       
       If Not rst2.EOF Then
    
            pGender = rst2![Gender]
            If IsNull(rst2![Date_Birth]) Then ' MM Opraveno 28.12.2016 - predtim padalo, kdyz jsem neznali datum narozeni
            Else
                pDateBirth = rst2![Date_Birth]
            End If
            
            pTextTitleBefore = "" & rst2![Text_Title_before] & ""
            pTextTitleAfter = "" & rst2![Text_Title_After] & ""
            pNationality = "" & rst2![Nationality] & ""
            pPin = "" & rst2![Pin] & ""
            
       End If
   rst2.Close
   
   
   ' z tabulky PERSON_CONTACT 'RISK-403
   sqlContact = "select primary_phone, primary_email  from " & schema & ".PERSON_CONTACT where id_person = " & pId_person

   Set rst3 = New ADODB.Recordset
   rst3.Open sqlContact, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst3.EOF Then

        pPrimaryPhone = "" & rst3![PRIMARY_PHONE] & ""
        pPrimaryEmail = "" & rst3![PRIMARY_EMAIL] & ""

   End If
   rst3.Close
   
   
   ' z tabulky PERSON_ADDRESS
   SQLADR = "select NAME_PERMANENT_ADDRESS_STREET, TEXT_PERMANENT_ADDRESS_NUM, CODE_PERMANENT_ADDRESS_ZIP, NAME_PERMANENT_ADDRESS_TOWN, NAME_PERMANENT_ADDRESS_COUNTRY , " & _
            " NAME_REPORTS_ADDRESS_STREET, TEXT_REPORTS_ADDRESS_NUM, CODE_REPORTS_ADDRESS_ZIP, NAME_REPORTS_ADDRESS_TOWN, NAME_REPORTS_ADDRESS_COUNTRY  " & _
            " from " & schema & ".PERSON_ADDRESS where id_person = " & pId_person

   Set rst4 = New ADODB.Recordset
   rst4.Open SQLADR, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
       If Not rst4.EOF Then
    
            'naplnìní TRVALÉ adresy z tabukly PERSON_ADDRESS - PA (PermanentAddress)
            pUlice_PA = "" & rst4![name_permanent_address_street] & ""
            pCP_PA = "" & rst4![text_permanent_address_num] & ""
            pPSC_PA = "" & rst4![code_permanent_address_zip] & ""
            pMesto_PA = "" & rst4![name_permanent_address_town] & ""
            pZeme_PA = "" & rst4![NAME_PERMANENT_ADDRESS_COUNTRY] & ""
            'naplnìní KONTAKTNÍ adresy z tabukly PERSON_ADDRESS - PA (PermanentAddress)
            pUlice_CA = "" & rst4![NAME_REPORTS_ADDRESS_STREET] & ""
            pCP_CA = "" & rst4![TEXT_REPORTS_ADDRESS_NUM] & ""
            pPSC_CA = "" & rst4![CODE_REPORTS_ADDRESS_ZIP] & ""
            pMesto_CA = "" & rst4![NAME_REPORTS_ADDRESS_TOWN] & ""
            pZeme_CA = "" & rst4![NAME_REPORTS_ADDRESS_COUNTRY] & ""
       End If
   rst4.Close
   
   
   
    If Len(pTextContractNumber) > 1 Then
        sql_datum_RS = "select trunc(dtime_signature) as dtime_signature from " & schema & ".sablona_epr where id_person = " & pId_person & "  and code_completion_subject_type = 'RS' AND TEXT_CONTRACT_NUMBER = " & pTextContractNumber
        Set rst_datum_RS = New ADODB.Recordset
        rst_datum_RS.Open sql_datum_RS, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
            If Not rst_datum_RS.EOF Then
                pDate_Contract_Signed = "" & nvl(rst_datum_RS![Dtime_signature]) & ""
            End If
        rst_datum_RS.Close
    End If
   
   If KlientiSeZapornymKU Is Nothing Then
        Set KlientiSeZapornymKU = CreateObject("Scripting.Dictionary")
   End If
   
   If zapornyKapitalizovanyUrok(c.id_person) And Not KlientiSeZapornymKU.Exists(c.id_person) Then
        MsgBox "Tenhle klient má záporný kapitalizovaný úrok, Jezevèík s tím ale moc nenadìlá, šablony bude potøeba opravovat.", vbExclamation, "Klient se záporným KU"
        KlientiSeZapornymKU.Add CStr(c.id_person), c.id_person
   End If
   
   Call napln_exekucni_titul
   
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



Private Sub napln_exekucni_titul()
    ' dotaz do stavu klienta v nalezacim rizeni a do reportu PM pohledavek
    Dim sql As String
    Dim rst As ADODB.Recordset
    
    assertion c.id_person > 0, "jez-78357: chyba v inicializaci klienta." ' should never happen
    
    sql = "SELECT nvl(A.TEXT_EXEKUCNI_TITUL, 0) as titul, nvl(A.spisova_znacka, 0) as spisova_znacka FROM " & schema & ".V_EXEKUCNI_TITUL A WHERE A.ID_PERSON = " & c.id_person
    
    Set rst = New ADODB.Recordset
    rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        If Not rst.EOF Then
            If rst.RecordCount = 1 Then
            
                Select Case rst![titul]
                
                    Case get_enum_exekucni_titul_Name(exe_titul_EPR)
                        pExeTitul = exe_titul_EPR
                        
                    Case get_enum_exekucni_titul_Name(exe_titul_platebni_rozkaz)
                        pExeTitul = exe_titul_platebni_rozkaz
                        
                    Case get_enum_exekucni_titul_Name(exe_titul_rozsudek)
                        pExeTitul = exe_titul_rozsudek
                        
                    Case get_enum_exekucni_titul_Name(exe_titul_usneseni_C)
                        pExeTitul = exe_titul_usneseni_C
                        
                    Case get_enum_exekucni_titul_Name(exe_titul_usneseni_EPR)
                        pExeTitul = exe_titul_usneseni_EPR
                        
                    Case get_enum_exekucni_titul_Name(exe_titul_none)
                        pExeTitul = exe_titul_none ' velmi vynimocne. Defaultni hodnota, radek jenom for clarity
                        
                    Case Else
                        pExeTitul = exe_titul_none
                        ' napr. klient CUID 109977 pri inicializaci z jakehokoliv rizeni
                        ' assertion 1 = 0, "jez-78359: chyba v inicializaci klienta. Neznámá forma exekuèního titulu." ' should never happen
                        MsgBox "jez-78359: chyba v inicializaci klienta. Neznámá forma exekuèního titulu."
                        
                End Select
                
                pExeTitul_spisova_znacka = rst![spisova_znacka]
                
            Else
                pExeTitul = exe_titul_none
                ' assertion 1 = 0, "jez-78358: chyba v inicializaci klienta. Nalezeno více záznamù exekuèního titulu." ' should never happen
                MsgBox "jez-78358: chyba v inicializaci klienta. Nalezeno více záznamù exekuèního titulu." ' should never happen
            End If
        
        Else
            pExeTitul = exe_titul_none ' Defaultni hodnota, radek jenom for clarity
        End If
    rst.Close
    
End Sub


Private Function zapornyKapitalizovanyUrok(ByVal in_id_person As String) As Boolean
' vytvorene 15.2.2018
' chyba na DWH, asi 35 klientov ma zaporne hodnoty pre kapitalizovany urok, rychle upozornenie pri nacitani klienta
' tiket #HEL-5615
' V PRIPADE CHYBY VRATI TRUE

Dim sqlKU As String
Dim rstKU As New ADODB.Recordset

sqlKU = "SELECT * FROM " & schema & ".V_DATACHECK_CREDIT_A A WHERE A.KAPITALIZOVNAY_UROK < 0 AND A.ID_PERSON = " & in_id_person
rstKU.Open sqlKU, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
    If rstKU.RecordCount > 0 Then zapornyKapitalizovanyUrok = True
    
rstKU.Close


End Function

'najde a vytvori objekty - ucty
Public Sub find_accounts()
Dim Acc As Object

If cs_HDWMN.State = 1 Then
prihlasen:

    Set collAccounts = New Collection
    Dim a_skp_bank_account() As Long, cnt_accounts As Integer
        
        'zjistime vsechny ucty
   sql = "select skp_bank_account skp from " & schema & ".ucet where id_person = " & pId_person
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   cnt_accounts = rst1.RecordCount
   If Not rst1.EOF Then
        ReDim a_skp_bank_account(cnt_accounts)
        rst1.MoveFirst
        For i = 1 To rst1.RecordCount
            a_skp_bank_account(i) = rst1![SKP]
            rst1.MoveNext
        Next i
        
   End If
   rst1.Close
   
 
   'vytvorime vsechny ucty
   For i = 1 To cnt_accounts
        Set Acc = New Account
        Call Acc.initialiaze(a_skp_bank_account(i))
        collAccounts.Add item:=Acc, key:="" & a_skp_bank_account(i) & ""
   Next i
   
   

    
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





'najde a vytvori objekty - uvery
Public Sub find_loans()
Dim Loan As Object

If cs_HDWMN.State = 1 Then
prihlasen:

    Set collLoans = New Collection
    Dim a_skp_loan() As Long, cnt_loans As Integer
        
        'zjistime vsechny uvery
   sql = "select skp_loan skp from " & schema & ".loan where id_person = " & pId_person & " order by date_loan_utilization"
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   cnt_loans = rst1.RecordCount
   If Not rst1.EOF Then
        ReDim a_skp_loan(cnt_loans)
        rst1.MoveFirst
        For i = 1 To rst1.RecordCount
            a_skp_loan(i) = rst1![SKP]
            rst1.MoveNext
        Next i
        
   End If
   rst1.Close
   
 
   'vytvorime vsechny uvery
   For i = 1 To cnt_loans
        Set Loan = New Loan
        Call Loan.initialiaze(a_skp_loan(i))
        collLoans.Add item:=Loan, key:="" & a_skp_loan(i) & ""
   Next i
   
   

    
    If err_connection Then
        err_connection = False
        Call re_login(Main.username.text, Main.password.text)
        If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
    End If
    
Else

    Call re_login(Main.username.text, Main.password.text)
    If cs_HDWMN.State = 1 Then
        GoTo prihlasen
    Else
        Call logout
    End If
End If




End Sub



Private Function NaplnSoud() As SOUD_type
    ' funkce vrati promennu typu SOUD pro aktualne inicializovaneho klienta v public promenne "c"
    
    Dim sql As String
    Dim rst As ADODB.Recordset
    Dim new_soud As SOUD_type
    Dim lng_missing As Long
    
    lng_missing = 0
    
    
    sql = "SELECT nvl(s.SKP_SOUD, 0) as SKP_SOUD, s.OKRESNI_SOUD, s.KRAJSKY_SOUD, s.ADRESA, s.ADRESA_P1, s.ADRESA_P2, s.ADRESA_P3, s.ID_SOUD, s.MAIL, S.tel " & _
          " FROM " & schema & ".person ps " & _
          " JOIN " & schema & ".soudy s ON ps.SKP_SOUD = s.SKP_SOUD " & _
          " WHERE ps.ID_PERSON = " & c.id_person
    
    Set rst = New ADODB.Recordset
    
    rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
        If rst.RecordCount = 1 Then
            
            With new_soud
                .skp_soud = rst![skp_soud]
                .okresni_soud = Empty_string_highlighter(nvl(rst![okresni_soud]), cfg.cst_missing_string_value)
                .krajsky_soud = Empty_string_highlighter(nvl(rst![krajsky_soud]), cfg.cst_missing_string_value)
                .adresa = Empty_string_highlighter(nvl(rst![adresa]), cfg.cst_missing_string_value)
                .adresa_P1 = Empty_string_highlighter(nvl(rst![adresa_P1]), cfg.cst_missing_string_value)
                .adresa_P2 = Empty_string_highlighter(nvl(rst![adresa_P2]), cfg.cst_missing_string_value)
                .adresa_P3 = Empty_string_highlighter(nvl(rst![adresa_P3]), cfg.cst_missing_string_value)
                .ID_soud = Empty_string_highlighter(nvl(rst![ID_soud]), cfg.cst_missing_string_value)
                .mail = Empty_string_highlighter(nvl(rst![mail]), cfg.cst_missing_string_value)
                .tel = Empty_string_highlighter(nvl(rst![tel]), cfg.cst_missing_string_value)
            End With
        
        Else
        
            With new_soud
                .skp_soud = lng_missing
                .okresni_soud = cfg.cst_missing_string_value
                .krajsky_soud = cfg.cst_missing_string_value
                .adresa = cfg.cst_missing_string_value
                .adresa_P1 = cfg.cst_missing_string_value
                .adresa_P2 = cfg.cst_missing_string_value
                .adresa_P3 = cfg.cst_missing_string_value
                .ID_soud = cfg.cst_missing_string_value
                .mail = cfg.cst_missing_string_value
                .tel = cfg.cst_missing_string_value
            End With
            
        End If
    rst.Close
    
    NaplnSoud = new_soud
    
End Function
