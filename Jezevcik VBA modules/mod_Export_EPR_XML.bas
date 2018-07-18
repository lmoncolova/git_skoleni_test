Attribute VB_Name = "mod_Export_EPR_XML"
Option Explicit

Public XDoc As Object, root As Object 'XDoc public, pouziva sa aj v insolvencnej sablone IS_XML (sub naplnElement)
Public atribut As Object
Private castkaList As Object

Public ParentElement As Object, ChildElement As Object
Public NewAtribut As Object
Public cI As Long
Public cL As Long
Public cA As Long
Public Folder As String
Public cilElem As String

Private listNode As Variant





'http://analystcave.com/vba-xml-working-xml-files/
'RISK-392, RISK-389, RISK-391, RISK-390
Public Sub generovatEPR_XML()
Dim castD As String
Dim elem As Integer ' poèítání, kolikátı element pro nárok mám naplnit, plní se jen uvìry aktivní, zesplatnìné, k odpisu, odepsané, pøípadnì nepovolenı debet
Dim elemUrok As Integer 'pro poèítání, kolikátı element pro pøísluèenství mám naplnit (plní se jen nenulové pøíslušenství)
'Dim cI As Integer
Dim cilElem As String
Dim adresaCZ As String
Dim poplatek As Double
Dim rstEPRS As ADODB.Recordset
Dim KodEprSoud As String
Dim zesplDopis As Date
Dim fso
Dim sqleprs As String
Dim sqlLoxon As String
Dim rstLoxon As ADODB.Recordset
    
    strMsgbox = ""
    
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    
    'Naètení xml s elementy EPR s pøedvyplnìnımi konstantními údaji (alobce = AB, mail:michaelstefl@airbank.cz, úèet pro vraceni sop 30300389/3030,...)
    XDoc.Load (cfg.Cesta("vzorEprXml"))
    Set root = XDoc.DocumentElement
    
'    Set atribut = root.ChildNodes(0)

'    Debug.Print atribut.XML
    'Set atribut = root.ChildNodes(0).ChildNodes(0)
    'Set atribut = root.getElementsByTagName("Name")
    
    'doèasnì pro ladeni kodu, pozdeji se bude inicializovat v souvilosit s vıberem klienta
    'Call c.Initialize(2252037)
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'A) soud, kterému je návrh urèen
    ' je potøeba dotáhnout èíselník z PDF a doplnit do naší DB kódy soudù - EPR_SOUDY
'    cilElem = "//d:root/"
'    Call naplnElement(cilElem, "a_soud", "KSJIMBM")
    'Call naplnElement("//d:root/", "a_soud", "OSPHA10")
    
    ''''''' RISK-1098 sql upraveno
'''    sqleprs = "SELECT P.ID_PERSON, EPRS.KOD_EPR_SOUD FROM " & schema & ".PERSON P " & _
'''                "LEFT JOIN " & schema & ".person_soud PS ON PS.ID_PERSON = P.ID_PERSON " & _
'''                "LEFT JOIN " & schema & ".EPR_SOUDY EPRS ON UPPER(EPRS.NAZEV_EPR_SOUD) = UPPER(PS.OKRESNI_SOUD) or (EPRS.NAZEV_EPR_SOUD = 'Okresní soud ve Strakonicích' and PS.OKRESNI_SOUD = 'Okresní soud v Strakonicích') " & _
'''                "WHERE P.ID_PERSON = " & c.id_person
    sqleprs = "SELECT P.ID_PERSON, EPRS.KOD_EPR_SOUD FROM " & schema & ".PERSON P " & _
              " JOIN " & schema & ".soudy s ON s.SKP_SOUD = p.SKP_SOUD  " & _
              " JOIN " & schema & ".EPR_SOUDY EPRS ON UPPER(EPRS.NAZEV_EPR_SOUD) = UPPER(S.OKRESNI_SOUD) OR (EPRS.NAZEV_EPR_SOUD = 'Okresní soud ve Strakonicích' AND S.OKRESNI_SOUD = 'Okresní soud v Strakonicích') " & _
              " WHERE P.ID_PERSON = " & c.id_person
      Set rstEPRS = New ADODB.Recordset
        rstEPRS.Open sqleprs, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

           If Not rstEPRS.EOF Then
                
                If IsNull(rstEPRS![kod_epr_soud]) Then
                    HlaskaUzivateli ("Je tøeba doplnit soud! Ke klientovi není v DB Jezevèíka pøiøazen okresní soud! Teprve po pøiøazení soudu se objeví kolonka pro Skuteènosti vıznamné pro rozhodnutí - je tøeba upravit u úvìrù.")
                Else
                    Call naplnElement("//d:root/", "a_soud", rstEPRS![kod_epr_soud])
                End If


           End If
        rstEPRS.Close
        
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'B) Úèastnící a jejich zástupcI - ALOVANİ
    'NAPLNÌNÍ ÈÁSTI ALOVANİ - 1 _ Plní se PRO JEDNOHO ALOVANÉHO z modulu Client
    cilElem = "//d:root/b_zalovany_opak[0]/zalovany[0]/fyzicka_osoba[0]/"
    Call naplnElement(cilElem, "jmeno", c.Name_first)
    Call naplnElement(cilElem, "prijmeni", c.Name_last)
    If c.PinLom <> "" Then
        Call naplnElement(cilElem, "rodne_cislo", c.PinLom)
        
    End If
    Call naplnElement(cilElem, "datum_narozeni", format(c.Date_Birth, "yyyy-mm-dd"))
    Call naplnElement(cilElem, "titul_pred_jmenem", c.Text_Title_before)
    Call naplnElement(cilElem, "titul_za_jmenem", c.Text_Title_After)
    
    'pokud je trvalá adresa v èr, tak trvalé, jinak korespondenèní v èr, pokud by ani jedna nebyla v èr, zùstane prázdné
    If c.Address_P_Zeme = "Czech Republic" Then
       
        Call naplnElement(cilElem, "adresa_pevna[0]/ulice", c.Address_P_Ulice)
        If InStr(c.Address_P_CP, "/") = 0 Then 'pokud je v èp lomítko, vyplòuje se zvláš cp a ce
            Call naplnElement(cilElem, "adresa_pevna[0]/cp", c.Address_P_CP)
        Else
            Call naplnElement(cilElem, "adresa_pevna[0]/cp", Left(c.Address_P_CP, InStr(c.Address_P_CP, "/") - 1))
            Call naplnElement(cilElem, "adresa_pevna[0]/ce", Mid(c.Address_P_CP, InStr(c.Address_P_CP, "/") + 1))
        End If
        Call naplnElement(cilElem, "adresa_pevna[0]/obec", c.Address_P_Mesto)
        Call naplnElement(cilElem, "adresa_pevna[0]/psc", c.Address_P_PSC)
        Call naplnElement(cilElem, "adresa_pevna[0]/stat", "CZ")
    Else
        If c.Address_C_Zeme = "Czech Republic" Then 'Pokud trvalá není CZ a kontaktní je CZ, doplní se kontaktní
            Call naplnElement(cilElem, "adresa_pevna[0]/ulice", c.Address_C_Ulice)
            If InStr(c.Address_C_CP, "/") = 0 Then 'pokud je v èp lomítko, vyplòuje se zvláš cp a ce
                Call naplnElement(cilElem, "adresa_pevna[0]/cp", c.Address_C_CP)
            Else
                Call naplnElement(cilElem, "adresa_pevna[0]/cp", Left(c.Address_C_CP, InStr(c.Address_C_CP, "/") - 1))
                Call naplnElement(cilElem, "adresa_pevna[0]/ce", Mid(c.Address_C_CP, InStr(c.Address_C_CP, "/") + 1))
            End If
            Call naplnElement(cilElem, "adresa_pevna[0]/obec", c.Address_C_Mesto)
            Call naplnElement(cilElem, "adresa_pevna[0]/psc", c.Address_C_PSC)
            Call naplnElement(cilElem, "adresa_pevna[0]/stat", "CZ")
        'pokud ani trvalá ani kontaktní není CZ, zùstane prázdné a vyskoèí upozornìní
        Else
            HlaskaUzivateli ("Trvalá ani koresponenèní adresa není v ÈR, prosím doplò ruènì adresu alovaného do formuláøe.")
        End If
    End If
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'C) Pøedmìt - právo na zaplacení penìité èástky
    ' Nárok a) úvìr -> zaškrtnuto Pohledávka i Úrok
    '       b) nepovolenı debet -> Pohledávka
    'úrok najdu na kartì kontrola dat - co øádek a interval, to + úrok (i rozpad dle druhu)
    ' - kapitalizovanı - od = den zaplacení celé splátky / datum naèerpání úvìru (pokud nebyla zaplacena ani jedna celá splátka
    '                  - do = zaplacení / den zesplatnìní - 1 pokud je úroku více
    'skuteènosti úvìr = úvìr, debet = jiné
    'cilElem = "//d:root/c_naroky_opak"
    'Debug.Print c.GetCnt_Loans
     elem = -1
    If c.GetCnt_Loans > 0 Then
   
        For cL = 1 To c.GetCnt_Loans
              'tady vybrat jen poadované úvìry - ovìøit s PV jaké - zøejmì zesplatnìné + k odpisu + odepsané

            If c.GetLoans(cL).Loan_status = "Aktivní" And (c.GetLoans(cL).Loan_wfstatus = "Zesplatnìnı" Or c.GetLoans(cL).Loan_wfstatus = "K odpisu" Or c.GetLoans(cL).Loan_wfstatus = "Odepsanı bankou") Then
                elem = elem + 1
                Call pridejElement(0, "//d:root/c_naroky_opak", "narok")
                Call pridejElement(elem, "//d:root/c_naroky_opak[0]/narok", "pohledavka")
                Call pridejElement(0, "//d:root/c_naroky_opak[0]/narok[" & elem & "]/pohledavka", "castka")
                Call naplnElementNumber("//d:root/c_naroky_opak[0]/narok[" & elem & "]/pohledavka[0]/", "castka", c.GetLoans(cL).Amt_loan_principal_unpaid) 'doplnit jistinu NEBO NECO PO DOMLUVE S PV!!!
        
                
                Call pridejElement(0, "//d:root/c_naroky_opak[0]/narok[" & elem & "]/pohledavka", "mena")
                Call naplnElement("//d:root/c_naroky_opak[0]/narok[" & elem & "]/pohledavka[0]/", "mena", "KÈ")
                
                'èást, která plní uroky
                If c.GetLoans(cL).GetCntInterests > 0 Then
                        Call pridejElement(elem, "//d:root/c_naroky_opak[0]/narok", "prislusenstvi_opak")
                        elemUrok = -1
                        For cI = 1 To c.GetLoans(cL).GetCntInterests
                            
                            If c.GetLoans(cL).GetInterests(cI).AmtUrok > 0 Then
                                elemUrok = elemUrok + 1
                                Call pridejElementPrislusenstvi(elem, elemUrok, c.GetLoans(cL).GetInterests(cI), c.GetLoans(cL))  'další parametr bude elemUrok a plnìní všech vnitøních elementù bude ve fci pridejElementPrislusenstvi
                                'Debug.Print
                             End If
                         
                        Next cI
                End If
                
                Call pridejElement(elem, "//d:root/c_naroky_opak[0]/narok", "skutecnosti")
                Call pridejElement(0, "//d:root/c_naroky_opak[0]/narok[" & elem & "]/skutecnosti", "tyka_se_autor_zakona")
                Call naplnElement("//d:root/c_naroky_opak[0]/narok[" & elem & "]/skutecnosti[0]/", "tyka_se_autor_zakona", "smlouvy o pùjèce") '!!!tady to se do pdf podle toho, co je vyprané v poloce SOUD v èásti A!!!!
                Call pridejElement(elem, "//d:root/c_naroky_opak[0]/narok", "skutecnosti_a_dukazy")
                Call naplnElement("//d:root/c_naroky_opak[0]/narok[" & elem & "]/", "skutecnosti_a_dukazy", generuj_epr_skutecnosti_a_dukazy(c.GetLoans(cL), c.id_person))
                'Debug.Print ParentElement(0).XML
            
            End If
        Next cL
                
    

    End If
    'HlaskaUzivateli ("konec uver")
    'èást, kde se plní NEPOVOLENY DEBET
    If c.GetCnt_Accounts > 0 Then
    
    'vybrat jen poadované úèty - ty co mají nepovolenı debet. Pro JPK, spoøicí úèty a úèty v cizích mìnách bude vyskakovat upozornìní pro manuální doplnìní.
        For cA = 1 To c.GetCnt_Accounts
              'tady vybrat jen poadované úvìry - zesplatnìné + k odpisu + odepsané

              If c.GetAccounts(cA).Amt_balance < 0 And c.GetAccounts(cA).Account_type = "BÚ" And c.GetAccounts(cA).Code_currency = "CZK" Then
              
              
                elem = elem + 1
                Call pridejElement(0, "//d:root/c_naroky_opak", "narok")
                Call pridejElement(elem, "//d:root/c_naroky_opak[0]/narok", "pohledavka")
                Call pridejElement(0, "//d:root/c_naroky_opak[0]/narok[" & elem & "]/pohledavka", "castka")
                Call naplnElementNumber("//d:root/c_naroky_opak[0]/narok[" & elem & "]/pohledavka[0]/", "castka", -(c.GetAccounts(cA).Amt_balance))
                
                Call pridejElement(0, "//d:root/c_naroky_opak[0]/narok[" & elem & "]/pohledavka", "mena")
                Call naplnElement("//d:root/c_naroky_opak[0]/narok[" & elem & "]/pohledavka[0]/", "mena", "KÈ")
                
                
                Call pridejElement(elem, "//d:root/c_naroky_opak[0]/narok", "skutecnosti")
                Call pridejElement(0, "//d:root/c_naroky_opak[0]/narok[" & elem & "]/skutecnosti", "tyka_se_autor_zakona")
                Call naplnElement("//d:root/c_naroky_opak[0]/narok[" & elem & "]/skutecnosti[0]/", "tyka_se_autor_zakona", "jiné")
                Call pridejElement(elem, "//d:root/c_naroky_opak[0]/narok", "skutecnosti_a_dukazy")
                
                
                'Skuteènosti pro nepovolenı debet dle RISK-391
                              
                
                Call naplnElement("//d:root/c_naroky_opak[0]/narok[" & elem & "]/", "skutecnosti_a_dukazy", skutecnostiDebetText(c.GetAccounts(cA).skpBankAccount, c.id_person, "EPR"))
                'Debug.Print ParentElement(0).XML
                
                'Pokud je v debetu jinı úèet klienta ne Bìnı úèet, vyskoèí uppozornìní
                'pokud je v debetu cizomìnovı úèet, vyskoèí upozornìní viz RISK-391 komentáø 27.3.2017
              Else
                    If c.GetAccounts(cA).Amt_balance < 0 And c.GetAccounts(cA).Account_type <> "BÚ" Then
                        HlaskaUzivateli ("Klient má úèet se zápornım zùstatkem typu Spoøicí úèet nebo JPK. Je tøeba doplnit pro tento nepovolenı debet ruènì všechny údaje do EPR.")
                    Else
                        If c.GetAccounts(cA).Amt_balance < 0 And c.GetAccounts(cA).Code_currency <> "CZK" Then
                            HlaskaUzivateli ("Klient má úèet se zápornım zùstatkem v cizí mìnì. Je tøeba doplnit pro tento nepovolenı debet ruènì všechny údaje do EPR.")
                        End If
                    
                    End If
                
              End If
             Next cA
    End If
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'D)Náhrada nákladù øízení
    'ne
    'ano
    '+statickı text, kde se upraví 3 data (doplòuje se manuálnì - Michael Šteffl)
    'Soudní poplatek - poèítá pdf - ale a pøi zmìnì pole, ze kterého se vıpoèet prvádí
    'RISK-390
    'Za návrh na vydání elektronického platebního rozkazu, jeho pøedmìtem je penìité plnìní
    'a) do èástky 10000 Kè vèetnì 400 Kè
    'b) v èástce vyšší ne 10000 do 20000 Kè vèetnì 800 Kè
    'c) v èástce vyšší ne 20000 Kè 4 % z této èástky
    ' dle propoètu se to poèítá z jistiny a debetu a zaokrouhlí na koruny nahoru
    Set castkaList = XDoc.SelectNodes("//d:root/c_naroky_opak/narok/pohledavka/castka")
    poplatek = 0
    For Each listNode In castkaList
        poplatek = poplatek + Val(listNode.text)
    Next listNode
   
    Select Case WorksheetFunction.Ceiling(poplatek, 10) 'Základ poplatku se zaokrouhluje na celé desítky korun nahoru
        Case Is <= 10000
            Call naplnElement("//d:root/d_naklady_rizeni[0]/", "poplatek", "400")
        Case Is <= 20000
            Call naplnElement("//d:root/d_naklady_rizeni[0]/", "poplatek", "800")
        Case Is > 20000
            'Debug.Print poplatek
            Call naplnElement("//d:root/d_naklady_rizeni[0]/", "poplatek", WorksheetFunction.Ceiling(WorksheetFunction.Ceiling(poplatek, 10) * 0.04, 1)) 'Základ poplatku se zaokrouhluje na celé desítky korun nahoru a vypoètenı procentní poplatek ze základu se zaokrouhluje na celé koruny nahoru
    End Select
    
    '
    sqlLoxon = "select don as zesplLoxon from " & schema & ".sablona_epr12 where id_person = " & c.id_person
    Set rstLoxon = New ADODB.Recordset
        rstLoxon.Open sqlLoxon, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        If Not rstLoxon.EOF Then
         rstLoxon.MoveFirst
             zesplDopis = rstLoxon![zesplLoxon]
    End If
    rstLoxon.Close
    
    
    
    castD = "alovanému byl dne " & zesplDopis & " odeslán zesplatòující dopis - pøedalobní vızva na adresu bydlištì alovaného. " & Chr(10) & Chr(10) & _
            "Dùkaz:" & Chr(10) & "Zesplatòující dopis - pøedalobní vızva" & Chr(10) & _
            "Podací arch " & Chr(10) & "Doklad odeslané pošty"
    'Debug.Print castD
    
    Call naplnElement("//d:root/d_naklady_rizeni[0]/", "skutecnosti", castD)
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'E)Soudní poplatek
    ' ne vdy
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'F) Bankovní údaje
    '1 - BU klienta navázanı na HU, VS = èíslo RS, Pokud BÚ není zatíen exekuèní blokací, tak uvádìt èíslo BÚ. Pokud je zatíen exekuèní blokací, tak nechat prázdné.
        'Podle toho, jestli je vyplnìní èíslo úètu, tak ne/vyplòovat variabilní symbol
    '2 - vdy prázdné
    '3 - vdy 30300389/3030, vs = èíslo RS
    cilElem = "//d:root/e_bankovni_udaje[0]/"
        
    '1 - BU klienta navázanı na HU, VS = èíslo RS, Pokud BÚ není zatíen exekuèní blokací, tak uvádìt èíslo BÚ. Pokud je zatíen exekuèní blokací, tak nechat prázdné.
        'Podle toho, jestli je vyplnìní èíslo úètu, tak ne/vyplòovat variabilní symbol
    If c.GetCnt_Loans > 0 Then
            For i = 1 To c.GetCnt_Accounts
             If c.GetAccounts(i).Text_bank_account_number = c.GetLoans(1).Text_bank_account_number Then
                If c.GetAccounts(i).AmtGovernmentHoldSum = 0 Then
                    Call naplnElement(cilElem, "cu_zalovana_castka", c.GetAccounts(i).Text_bank_account_number & "/3030")
                    Call naplnElement(cilElem, "vs_zalovana_castka", c.TextContractNumber)
                End If
                
             End If
            Next i
            
   Else
    HlaskaUzivateli ("Èíslo úètu pro zaplacení alované èástky je tøeba doplnit! - Dle èísla BU, kterı je v debetu.")
   End If
   
   
   Call naplnElement(cilElem, "vs_vraceni_sop", c.TextContractNumber)
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'G) Petit - automaticky generovat - OK
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'H) Pøílohy - další fáze
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'Patièka - souhl = Ne, el. adresa vdy michael.stefl@airbank.cz, elektronickı podpis - automaticky v další fázi
 
    
'    Set atribut = XDoc.SelectNodes("//d:root/b_zalovany_opak[0]/zalovany[0]/fyzicka_osoba[0]/jmeno")
'    atribut(0).text = c.Name_first
   
    'Debug.Print atribut(0).text
    
    'Get first child ( same as ChildNodes(0) )
    
    Call VypisVsechnyHlasky
    
    'Save the XML file
    Call UlozSoubor(cfg.Cesta("EprPodaneZaloby") & c.Nazev_slozky_bez_lomitka, "BaseXmlEpr_" & c.Nazev_slozky_bez_lomitka & ".xml", cfg.Cesta("EprKlientiBezSlozky"))
    
    Set XDoc = Nothing
    
    Call VypisVsechnyHlasky
    
End Sub


Private Sub pridejElementPrislusenstvi(elemLoan As Integer, elemPrisl As Integer, Prislusenstvi As Interest, uver As Loan)

Dim druhUrok As String

    '<Prislusenstvi>
    Call pridejElement(0, "//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak", "prislusenstvi")
        '<Datum_splatnosti>
        Call pridejElement(elemPrisl, "//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi", "datum_splatnosti")
        Call naplnElement("//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/", "datum_splatnosti", format(c.GetLoans(cL).Date_payoff, "yyyy-mm-dd"))
        '<druh_prislusenstvi>
        Call pridejElement(elemPrisl, "//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi", "druh_prislusenstvi")
        
        Select Case Prislusenstvi.UrokTyp
            Case "KU"
                druhUrok = "kapitalizovany_urok"
            Case "SU"
                druhUrok = "zakonny_urok_z_prodleni"
            Case "OU"
                druhUrok = "urok"
        End Select
        
        '''''''''''''
        'Call naplnElement("//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/", "druh_prislusenstvi", druhUrok)
        Call naplnElement("//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/", "druh_prislusenstvi", IIf(druhUrok = "kapitalizovany_urok", druhUrok, "urok")) ' HOTFIX epodatelna, Jez.v.5.3, RISK-726
        ''''''''''''''
        
        
        
        '<urok>
        Call pridejElement(elemPrisl, "//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi", "urok")
                '<kapitalizovana_castka/>
                Call pridejElement(0, "//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]", "kapitalizovana_castka")
                If druhUrok = "kapitalizovany_urok" Then
                    Call naplnElementNumber("//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]/", "kapitalizovana_castka", Prislusenstvi.AmtUrok)
                    '<kapitalizovana_mena/>
                    Call pridejElement(0, "//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]", "kapitalizovana_mena")
                    Call naplnElement("//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]/", "kapitalizovana_mena", "KÈ")
               
                End If
                '<sazba/>
                 Call pridejElement(0, "//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]", "sazba")
                '-<casovy_interval>
                Call pridejElement(0, "//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]", "casovy_interval")
                    '<perioda>roènì</perioda>
                    Call pridejElement(0, "//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]/casovy_interval", "perioda")
                        '<jinak/>
                        Call pridejElement(0, "//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]/casovy_interval", "jinak")
                '<z_castky>z pohledávky/z èástky</z_castky>
                Call pridejElement(0, "//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]", "z_castky")
                '<castka/>
                Call pridejElement(0, "//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]", "castka")
                '<mena>KÈ</mena>
                Call pridejElement(0, "//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]", "mena")
                
                If druhUrok = "urok" Then
                    Call naplnElementNumber("//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]/", "sazba", uver.Rate_Interest_Standard) ' RISK-726 predtym Prislusenstvi.RateInterestStandard
                ElseIf druhUrok = "zakonny_urok_z_prodleni" Then
                    '''  ZAKONNY UROK, doplneno 22.2.2018 do nove verze EPR formulare 30-a, upraveno pak 5.3.2018 v ramci RISK-726
                    Call naplnElementNumber("//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]/", "sazba", sngZakonnyUrok + uver.Rate_Interest_Repo)
                End If
               
                
                'Pro kapitalizovanı urok se tato èást neplní
                If druhUrok <> "kapitalizovany_urok" Then
                    '<perioda>roènì</perioda>
                    Call naplnElement("//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]/casovy_interval[0]/", "perioda", "roènì")
            
                        If Round(uver.Amt_loan_principal_unpaid, 5) = Round(Prislusenstvi.AmtBalancePrincipal, 5) Then 'zaokrouhlení, protoe u CUID 202295 se bez zaokrouhlení vyhodnotilo nok
                            Call naplnElement("//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]/", "z_castky", "z pohledávky")
                        Else
                            Call naplnElement("//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]/", "z_castky", "z èástky")
                        End If
                   
                    Call naplnElementNumber("//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]/", "castka", Prislusenstvi.AmtBalancePrincipal)
                    Call naplnElement("//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]/", "mena", "KÈ")

                End If
                '<datum_od>
                Call pridejElement(0, "//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]", "datum_od")
                Call naplnElement("//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]/", "datum_od", format(Prislusenstvi.StartInterval, "yyyy-mm-dd"))
                '<do_zaplaceni>
                Call pridejElement(0, "//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]", "do_zaplaceni")
                '<datum_do xsi:nil="true" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"/>
                Call pridejElement(0, "//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]", "datum_do")
                
                If Round(uver.Amt_loan_principal_unpaid, 5) <> Round(Prislusenstvi.AmtBalancePrincipal, 5) Or druhUrok = "kapitalizovany_urok" Then
                    Call naplnElement("//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]/", "do_zaplaceni", "0")
                    Call naplnElement("//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]/", "datum_do", format(Prislusenstvi.KonecInterval, "yyyy-mm-dd"))
                Else
                    Call naplnElement("//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi[" & elemPrisl & "]/urok[0]/", "do_zaplaceni", "1")
                End If
                
                
                
        
        '<rozepsano_text>
        Call pridejElement(elemPrisl, "//d:root/c_naroky_opak[0]/narok[" & elemLoan & "]/prislusenstvi_opak[0]/prislusenstvi", "rozepsano_text")
        'Debug.Print root.ChildNodes(4).XML
End Sub






Public Function skutecnostiDebetText(skpAccount, skpPerson, ucel As String, Optional ByRef CelkovaVyskaDebetu As Double = 0) As String

' by Filip Antol; ucel je bud "EPR" nebo "IP"   'RISK-391
' vstupni argumenty: CelkovaVyskaDebetu - MUSI BYT BYREF, je to kontrolni vystup v pripade insolvencniho rizeni

Dim dukazy As String
Dim dukaz_cenik As String
Dim dukaz_odeslana_posta As String

Dim pridat_dukaz_vymahani As Boolean
Dim pridat_dukaz_upominka As Boolean
Dim pridat_dukaz_cenik As Boolean

Dim nacitat_sumu_prichozich As Boolean
Dim err_nevysvetlena_castka As Boolean
Dim err_nenaparovane_storna As Boolean

Dim sql_debet As String
Dim datum_trn As String
Dim datum_trn_accounted As String
Dim amt_trn As Single
Dim amt_prichozi_trn As Single
Dim amt_kladny_zustatek As Single
Dim kategorie As String
Dim flag_transaction_debit As String

Dim nevysvetlena_castka As Single
Dim nenaparovane_storna As Single
' Dim strMsgbox As String ' vyhozeno jez.v.5.9 v ramci RISK-1249
Dim kategorie_do_ponizeni As Variant
Dim kateg As Variant
Dim rst_debet As ADODB.Recordset


Dim textace As Object
Set textace = CreateObject("Scripting.dictionary")
Dim dluzne_castky As Object
Set dluzne_castky = CreateObject("scripting.dictionary")
Dim dluzne_datumy As Object
Set dluzne_datumy = CreateObject("scripting.dictionary")
Dim ponizeni As Object
Set ponizeni = CreateObject("scripting.dictionary")



''''' ........................................................................................................................................
'''''     NASTAVENI FIXNICH TEXTACI
'''  názov textace musí by zhodnı s názvom kategorie (nišie v skripte)!

With textace
    
    '''   rozne textace pre IP a EPR   (v Insolvenci neexistuje zalobce a zalovany, ale dluznik a veritel)
    
    If ucel = "IP" Then
        .Add "náklady na vymáhání", " odpovídá nezaplacenım nákladùm úèelnì vynaloenıch vìøitelem na vymáhání dluhu po dluníkovi, viz Oznámení o naúètování nákladù na vymáhání ze dne " '300 Kc
'        .Add "náklady na vymáhání_jeden", textace("náklady na vymáhání") ' RISK-1054
'        .Add "náklady na vymáhání_vice", textace("náklady na vymáhání")
'        .Add "náklady na vymáhání", " odpovídá nezaplacenım nákladùm úèelnì vynaloenıch alobcem na vymáhání dluhu po alovaném. alobce se od prvního dne prodlení" & _
'                                    " snail alovaného kontaktovat - upomínal alovaného opakovanì telefonicky, v jeho internetovém bankovnictví, e-mailem a SMS zprávami." & _
'                                    " V Podmínkách pro pouívání úvìru byla stranami sjednána monost alobce poadovat od alovaného náhradu úèelnì vynaloenıch nákladù" & _
'                                    " na vymáhání, viz èl. Splácení úvìru, odst. 7. alobce také pøedkládá soudu Pøehled nákladù úèelnì vynaloenıch bankou na vymáhání dluhu," & _
'                                    " kde uvádí pøehled nákladù, které mu s vymáháním dluhu vznikly." & _
'                                    " Z dùvodu neplnìní závazku byly alovanému tyto náklady ve vıši 300,- naúètovány dne " ' RISK-1054 zmena textace
'        .Add "náklady na vymáhání_závìr", ", viz Oznámení o naúètování nákladù na vymáhání ze dne "
        
        .Add "tarif_jeden", " odpovídá nezaplacenému poplatku za vedení bìného úètu dluníka s Velkım tarifem ze dne "
        .Add "tarif_vice", " odpovídá nezaplacenım poplatkùm za vedení bìného úètu dluníka s Velkım tarifem ze dne "
        .Add "offline", " odpovídá zápornému zùstatku na bìném úètu dluníka."
        .Add "pojištìní", " odpovídá dlunému pojistnému, které bylo z úètu dluníka po ukonèení pojištìní doplaceno pojišovnì."
    Else ' EPR
'        .Add "náklady na vymáhání", " odpovídá nezaplacenım nákladùm úèelnì vynaloenıch alobcem na vymáhání dluhu po alovaném, viz Oznámení o naúètování nákladù na vymáhání ze dne " '300 Kc
        .Add "náklady na vymáhání", " odpovídá nezaplacenım nákladùm úèelnì vynaloenıch alobcem na vymáhání dluhu po alovaném. alobce se od prvního dne prodlení" & _
                                    " snail alovaného kontaktovat - upomínal alovaného opakovanì telefonicky, v jeho internetovém bankovnictví, e-mailem a SMS zprávami." & _
                                    " V Podmínkách pro pouívání úvìru byla stranami sjednána monost alobce poadovat od alovaného náhradu úèelnì vynaloenıch nákladù" & _
                                    " na vymáhání, viz èl. Splácení úvìru, odst. 7. alobce také pøedkládá soudu Pøehled nákladù úèelnì vynaloenıch bankou na vymáhání dluhu," & _
                                    " kde uvádí pøehled nákladù, které mu s vymáháním dluhu vznikly." & _
                                    " Z dùvodu neplnìní závazku byly alovanému tyto náklady ve vıši 300,- naúètovány dne " ' RISK-1054 zmena textace
        .Add "náklady na vymáhání_závìr", ", viz Oznámení o naúètování nákladù na vymáhání ze dne "
        
        .Add "tarif_jeden", " odpovídá nezaplacenému poplatku za vedení bìného úètu alovaného s Velkım tarifem ze dne "
        .Add "tarif_vice", " odpovídá nezaplacenım poplatkùm za vedení bìného úètu alovaného s Velkım tarifem ze dne "
        .Add "offline", " odpovídá zápornému zùstatku na bìném úètu alovaného." 'bez datumu a s bodkou.
        .Add "pojištìní", " odpovídá dlunému pojistnému, které bylo z úètu alovaného po ukonèení pojištìní doplaceno pojišovnì." 'inak nez ostatne
    End If

    '''   rovnake textace pre IP a EPR
    .Add "debet", ""
    .Add "dc", "Dluná èástka "
    .Add "upomínka_jeden", " odpovídá nezaplacenému poplatku za upomínku ze dne "
    .Add "upomínka_vice", " odpovídá nezaplacenım poplatkùm za upomínky ze dne "
    .Add "sms_jeden", " odpovídá mìsíènímu vyúètování za odeslání SMS zprávy ze dne "
    .Add "sms_vice", " odpovídá mìsíènímu vyúètování za odeslání SMS zpráv ze dne "
    .Add "poplatek vıbìr_jeden", " odpovídá nezaplacenému poplatku za vıbìr hotovosti z bankomatu ze dne "
    .Add "poplatek vıbìr_vice", " odpovídá nezaplacenım poplatkùm za vıbìry hotovosti z bankomatù ze dne "
    .Add "vıpis_jeden", " odpovídá vyúètování za odeslání vıpisu ze dne "
    .Add "vıpis_vice", " odpovídá vyúètováním za odeslání vıpisù ze dne "
    .Add "poplatek sazka_jeden", " odpovídá vyúètování za vıbìr penìz u Sazky ze dne "
    .Add "poplatek sazka_vice", " odpovídá vyúètováním za vıbìry penìz u Sazky ze dne "
    .Add "exekuce", " Kè - NEZAØAZENA EXEKUCE - "
    .Add "nezaradene", " Kè - NEZAØAZENA TRANSAKCE - "
    .Add "nenaparovane_storna", " ,-Kè jsou STORNO TRANSAKCE nebo odúètování poplatkù, které se nepovedlo napárovat na pùvodní transakci - je potøeba dohledat v OBS a odpoèítat. Data: "
    .Add "debet_zaver", "Dle èl. 1 Rámcové smlouvy jsou nedílnou souèástí této smlouvy také Obchodní podmínky, Podmínky pro pouívání úvìru a Ceník. " & _
                       "S tìmito dokumenty se alovanı mìl monost prokazatelnì seznámit ještì pøed podpisem smlouvy na webovıch stránkách alobce, " & _
                       "k èemu byl také vyzván i v samotném znìní Rámcové smlouvy pøed jejím podpisem, kdy v èl. 1 alobce sám vyzıvá alovaného, aby v pøípadì, e nemìl monost se s tìmito dokumenty øádnì seznámit, a smlouvu nepodepisuje. " & _
                       "Na základì znìní Rámcové smlouvy ve spojení s vıše uvedenımi dokumenty poaduje alobce po alovaném úhradu dluhu." & Chr(10)
    .Add "insol_zaver", "Pøiloen mimoøádnı vıpis z bìného úètu." ' ke dni "
End With


dukaz_cenik = "Ceník Air Bank a.s. "
dukaz_odeslana_posta = "Doklad odeslané pošty."
dukazy = "Dùkazy: " & Chr(10) & _
    "Obchodní podmínky Air Bank a.s. " & Chr(10) & _
    "Podmínky pro pouívání úvìru Air Bank a.s. " & Chr(10) & _
    "Rámcová smlouva è. " & c.TextContractNumber & "." & Chr(10) & _
    "Mimoøádnı vıpis z úètu alovaného. " & Chr(10)

err_nevysvetlena_castka = True
err_nenaparovane_storna = True
nacitat_sumu_prichozich = True



''''' ........................................................................................................................................
'''''   NACTENI DAT

'''  nacitame zoradene poplatky a transakce klienta
sql_debet = "SELECT a.id_person, trunc(a.DATE_TRANSACTION_DUE) DATE_TRANSACTION_DUE, trunc(a.dtime_transaction_accounted) dtime_transaction_accounted, " & _
            " a.amt_transaction, a.kategorie, a.flag_transaction_debit, a.nevysvetlena_castka, a.nenaparovane_storna, a.suma_prichozich_trn " & _
            " FROM " & schema & ".epr_debety_all a " & _
            " WHERE id_person = " & skpPerson & " AND skp_bank_account = " & skpAccount & _
            " ORDER BY a.dtime_transaction_accounted , a.flag_transaction_debit ASC , CASE WHEN upper(a.kategorie) LIKE UPPER('%poplatek%') THEN 2 ELSE 1 END"
            
 Set rst_debet = New ADODB.Recordset
   rst_debet.Open sql_debet, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   If Not rst_debet.EOF Then
     rst_debet.MoveFirst
     
     Do Until rst_debet.EOF = True
     
        datum_trn = "" & rst_debet![date_transaction_due] & "" 'RISK-734. V textaci uvadzame datum provedeni vyberu, ne datum zauctovani. (pozor, v sql vyssie je stale potrebne ORDER BY datum zauctovani)
        datum_trn_accounted = "" & rst_debet![dtime_transaction_accounted] & "" 'kvoli pojisteni, tam potrebujeme tento datum namiesto date_transaction_due.
        amt_trn = rst_debet![AMT_TRANSACTION]
        kategorie = "" & rst_debet![kategorie] & "" 'uvodzovkami okolo excel nevyhodi error ak je tam NULL
        flag_transaction_debit = "" & rst_debet![flag_transaction_debit] & ""
        nevysvetlena_castka = rst_debet![nevysvetlena_castka]
        nenaparovane_storna = rst_debet![nenaparovane_storna]
        
        If nacitat_sumu_prichozich = True Then  'Tahle hodnota se nacita pouze jednou
          amt_prichozi_trn = rst_debet![suma_prichozich_trn]
          nacitat_sumu_prichozich = False
        End If
        
        ' jedina transakce ktora robi komplikuje flag_transaction_debit je Kreditní úrok.
        ' aby nemuseli byt vsade podmienky naviac, tu to prepisem, ale nesmie sa na to zabudnut.
        ' pozor aby tieto riadky boli PRED if (amt_trn < 0)
        If flag_transaction_debit = "N" Then
          amt_trn = -amt_trn
          flag_transaction_debit = "Y"
        End If
        
        
        'v tabulke mame stale riadky s nenaparovanymi stornami. U tych chceme skocit na dalsiu iteraciu, inak sa rozbiju vsetky cisla pri ponizovani.
        If amt_trn < 0 Then
            'do dluzne_castky pridavat NEchcem, pretoze cez tie ide loop na konci tejto funkcie
            If Not dluzne_datumy.Exists("nenaparovane_storna") Then
                dluzne_datumy.Add "nenaparovane_storna", datum_trn
            Else
                dluzne_datumy("nenaparovane_storna") = dluzne_datumy("nenaparovane_storna") & ", " & datum_trn
            End If
            GoTo NextIteration
        End If
        
        ' nenaparovane transakce alebo nesedi suma transakci
        If nevysvetlena_castka <> 0 And err_nevysvetlena_castka Then
          'HlaskaUzivateli "U tohohle klienta nesedí vıška debetu se sumou nalezenıch transakcí. Je nutné manuálnì dohledat rozdíl ve vıši " & nevysvetlena_castka & "."
          HlaskaUzivateli "> U tohohle klienta nesedí vıška debetu se sumou nalezenıch transakcí. Je nutné manuálnì dohledat rozdíl ve vıši " & nevysvetlena_castka & "." & Chr(10)
          err_nevysvetlena_castka = False 'aby sa vypisoval iba 1 krat
        End If
        If nenaparovane_storna <> 0 And err_nenaparovane_storna Then
          'HlaskaUzivateli "Klient má nìjaké storno transakce ve vıši " & nenaparovane_storna & ", které se nepovedlo napárovat na pùvodí transakce. Je potøeba dohledat pùvodní transakce v dané vıši."
          HlaskaUzivateli "> Klient má nìjaké storno transakce ve vıši " & nenaparovane_storna & ", které se nepovedlo napárovat na pùvodí transakce. Je potøeba dohledat pùvodní transakce v dané vıši." & Chr(10)
          err_nenaparovane_storna = False 'aby sa vypisoval iba 1 krat
        End If
        
          
        If (kategorie = "posledni kladny zustatek") Then
          amt_kladny_zustatek = amt_trn
          GoTo NextIteration
        End If
        
        
        'Ak klientova platba alebo posledni zustatek (alebo oboje) PLNE pokryje debetní transakci, nechceme ju vobec do textace.
        If amt_kladny_zustatek >= amt_trn Then
          amt_kladny_zustatek = Round(amt_kladny_zustatek - amt_trn, 2)
          GoTo NextIteration
        End If
        If amt_prichozi_trn >= amt_trn Then
          amt_prichozi_trn = Round(amt_prichozi_trn - amt_trn, 2)
          'amt_trn = 0 'zbytocne
          GoTo NextIteration
        End If
        If (amt_kladny_zustatek + amt_prichozi_trn) >= amt_trn Then
          amt_trn = Round(amt_trn - amt_kladny_zustatek, 2)
          amt_kladny_zustatek = 0 'je jedno ci znulujem zostatok alebo prichozi
          amt_prichozi_trn = Round(amt_prichozi_trn - amt_trn, 2)
          GoTo NextIteration
        End If

        
        
        Select Case kategorie
          Case "náklady na vymáhání"
            kategorie_do_ponizeni = kategorie
            'dukazy = dukazy & "Oznámení o naúètování nákladù na vymáhání ze dne " & datum_trn & "." & Chr(10) ' FA vyhozeno jez.v.5.9
            pridat_dukaz_vymahani = True
          
          Case "upomínka"
            kategorie_do_ponizeni = kategorie
            dukazy = dukazy & "Upomínka ze dne " & datum_trn & "." & Chr(10)
            pridat_dukaz_upominka = True
          
          Case "sms", "poplatek vıbìr", "vıpis", "poplatek sazka", "tarif" 'vsetky tie, ku ktorym sa priklada CENIK do dokazov
            kategorie_do_ponizeni = kategorie
            pridat_dukaz_cenik = True
          
          Case "pojištìní"
            kategorie_do_ponizeni = kategorie
            datum_trn = datum_trn_accounted 'u pojisteni to chceme takto. RISK-734
          
          Case "Karetní transakce", "Vıbìr z cizího bankomatu", "Daò z úroku", "Kreditní úrok", "Vıbìr z cizího bankomatu v cizinì v EU", "Vıbìr z cizího bankomatu v cizinì mimo EU"
            kategorie_do_ponizeni = "offline"
            'Kreditní úrok je na DWH NEdebetní a má obrácené znaménko, aby som nemusel prepisova všade podmienku, tak som ho (trochu vyssie) zmenil na debetnu trn a obrátil znamienko.
            'Pointa je e všetky èísla vo vısledku sedia, ale nekorešponduje to úplne s info na DWH a v podkladovıch tabu¾kách.
          
          Case "Splátka úvìru", "Pøevod", "Sazka kód", "Platba obchodního úroku", "Sankce za neplacení úvìru", "Vıbìr z vlastního bankomatu", "Zesplatnìní úvìru bankou"
            kategorie_do_ponizeni = "nezaradene" 'nezaradene su tu na 2 miestach. Rozdiel je v chybovej hlaske iba.
            HlaskaUzivateli "> Nìco je špatnì. U klienta je debetní transakce typu " & kategorie & ", touhle transakcí se však do debetu dostat nedá. Je potøeba zkontrolovat." & Chr(10)
          
          Case "Vıplata exekuce"
            kategorie_do_ponizeni = "exekuce"
            HlaskaUzivateli "> U klienta byla zaznamenána debetní transakce typu " & kategorie & _
                                    ". To znamená, e v textaci chybí èástka ve vıši " & amt_trn - (amt_kladny_zustatek + amt_prichozi_trn) & " Kè. Celková vıše pohledávky by mìla sedìt." & Chr(10) ' upraveno RISK-1031
            'HlaskaUzivateli "> U klienta je debetní transakce typu " & kategorie & " ve vıši " & amt_trn & " Kè. Je potøeba opravit textaci." & Chr(10)
            
            
          Case Else
            kategorie_do_ponizeni = "nezaradene" 'nezaradene su tu na 2 miestach. Rozdiel je v chybovej hlaske iba.
            HlaskaUzivateli "> U klienta je debetní transakce typu " & kategorie & " ve vıši " & amt_trn & " Kè. Je potøeba opravit textaci." & Chr(10)
          
        End Select
        
        
        'pre danu kategoriu pridame zaznam do dluzne_castky aj do dluzne_datumy.
        If Not dluzne_castky.Exists(kategorie_do_ponizeni) Then
            dluzne_castky.Add kategorie_do_ponizeni, amt_trn
            dluzne_datumy.Add kategorie_do_ponizeni, datum_trn
        Else
            dluzne_castky(kategorie_do_ponizeni) = Round(dluzne_castky(kategorie_do_ponizeni) + amt_trn, 2)
            dluzne_datumy(kategorie_do_ponizeni) = dluzne_datumy(kategorie_do_ponizeni) & ", " & datum_trn
        End If
        
        If ucel = "IP" Then
            strDatumPosledniDebetniTrn = datum_trn
            CelkovaVyskaDebetu = Round(CelkovaVyskaDebetu + amt_trn, 2)
        End If
        
        'ponizeni o prichozi transakci anebo kladny zustatek na uctu (pokud to nepokryje celou vysi transakce)
        
        If amt_kladny_zustatek > 0 Or amt_prichozi_trn > 0 Then
            ponizeni.Add "kateg", kategorie_do_ponizeni 'nefungovalo pre offline, vsetky nezaradene, exekuce (na konci nenajde zhodu
            ponizeni.Add "amt_zustatek", amt_kladny_zustatek 'inicializoval som nulu na zaciatku
            ponizeni.Add "amt_prichozi", amt_prichozi_trn    'inicializoval som nulu na zaciatku
            
            CelkovaVyskaDebetu = Round(CelkovaVyskaDebetu - amt_kladny_zustatek - amt_prichozi_trn, 2)
            amt_kladny_zustatek = 0
            amt_prichozi_trn = 0
        End If
            
        
NextIteration:

        rst_debet.MoveNext
        
     Loop
     
     
     
   End If
 rst_debet.Close



''''' ........................................................................................................................................
'''''   VYGENEROVANI TEXTACE


'dluzna castka xy ,-Kè je STORNO TRANSAKCE nebo odúètování poplatku, které se nepovedlo napárovat na pôvodní transakci - je potøeba dohledat v OBS a odpoèítat.

'NASLEDUJE vygenerovani textace. v dictionaries dluzne_castky, dluzne_datumy a ponizeni mam zaznamy ku vsetkym typom transakcii ktore sa vyskytli u klienta.
'Nasleduje loop cez tieto dictionaries a textace ktory vytvori cely text (okrem dokazov)
'najprv sa riesi pojisteni pretoze ma specialnu textaci a uvadza sa iba jeden datum.
'potom sa riesi offline, pretoze ma specialnu textaci.
'tak uz sa riesia vsetky ostatne - najprv tie co maju rozne textace pre 1 a viac zaznamov (poplatek vs. poplatky) a potom ostatne
'najprv hlavna cast textu s ciastkou, datumom, kategoriou
' a potom sa prida cast textu ak bola ciastka ponizena.
For Each kateg In dluzne_castky.keys 'v dictionary su iba tie kategorie, u ktorych bola nejaka trn
    
    If kateg = "exekuce" Then
        ' RISK-1031 - Vyplata Exekuce se nema vypisovat
        GoTo nextKateg
    End If
    
    If kateg = "pojištìní" Then 'pojisteni ma specialni textaci, jinak datum a jinak ponizeni
        If ponizeni("kateg") = kateg Then
            'BUG
'            textace("debet") = textace("debet") & _
'                               textace("dc") & Round(dluzne_castky(kateg) - ponizeni("amt_prichozi") - ponizeni("amt_zustatek"), 2) & _
'                               " ze dne " & Left(dluzne_datumy(kateg), InStr(dluzne_datumy(kateg), ",") - 1) & textace(kateg) & Chr(10) 'iba 1.datum treba sem dat
            textace("debet") = textace("debet") & _
                               textace("dc") & VlozCastku(dluzne_castky(kateg) - ponizeni("amt_prichozi") - ponizeni("amt_zustatek"))
        Else
            textace("debet") = textace("debet") & _
                               textace("dc") & VlozCastku(dluzne_castky(kateg))
        End If
        
        textace("debet") = textace("debet") & _
                            " ze dne " & _
                            IIf(InStr(dluzne_datumy(kateg), ",") <= 0, _
                                 dluzne_datumy(kateg), _
                                 Left(dluzne_datumy(kateg), WorksheetFunction.max(InStr(dluzne_datumy(kateg), ",") - 1, 0))) & _
                            textace(kateg) & Chr(10)  'iba 1.datum treba sem dat
        
    ElseIf kateg = "offline" Then 'offline ma specialni textaci, bez datumu a ponizeni je taky jinak
        If ponizeni("kateg") = kateg Then
            textace("debet") = textace("debet") & _
                               textace("dc") & VlozCastku(dluzne_castky(kateg) - ponizeni("amt_prichozi") - ponizeni("amt_zustatek")) & _
                               textace(kateg) & Chr(10) 'u offline neuvadzame datumy, bodka je v textaci
        Else
            textace("debet") = textace("debet") & _
                               textace("dc") & VlozCastku(dluzne_castky(kateg)) & _
                               textace(kateg) & Chr(10) 'u offline neuvadzame datumy, bodka je v textaci
        End If
        
        
    Else 'vetsina kategorii.
        
        
        Select Case kateg
            Case "upomínka", "sms", "poplatek vıbìr", "vıpis", "poplatek sazka", "tarif" 'vsetky tie, co maju odlisnu textaciu pre jednu a viacej transakcii v kategorii
                If Len(dluzne_datumy(kateg)) < 12 Then 'pokud je tam jen jeden datum (normalny datum ma delku 8-10, 12 je Bulharska aproximace)
                    textace("debet") = textace("debet") & _
                                        textace("dc") & VlozCastku(dluzne_castky(kateg)) & textace(kateg & "_jeden") & dluzne_datumy(kateg)
                Else
                    textace("debet") = textace("debet") & _
                                    textace("dc") & VlozCastku(dluzne_castky(kateg)) & textace(kateg & "_vice") & dluzne_datumy(kateg)
                End If
            
            Case "náklady na vymáhání" ' datum sa vklada do tela aj na koniec, RISK-1054
                textace("debet") = textace("debet") & _
                                   textace("dc") & VlozCastku(dluzne_castky(kateg)) & textace(kateg) & dluzne_datumy(kateg)
                '''   nasl. kousek textace vyhozen ve verzi jez.v.5.8
'                If ucel = "EPR" Then
'                    textace("debet") = textace("debet") & _
'                                       textace("náklady na vymáhání_závìr") & dluzne_datumy(kateg)
'                End If
            Case Else 'tie, co maju rovnaku textaciu pre jednu a viacej transakcii v kategorii
                textace("debet") = textace("debet") & _
                                   textace("dc") & VlozCastku(dluzne_castky(kateg)) & textace(kateg) & dluzne_datumy(kateg)
        End Select
        
        
        'ak bola ponizena transakce, tu sa vlozi textace (okrem offline a pojisteni)
        If ponizeni("kateg") = kateg Then
            If ponizeni("amt_prichozi") > 0 And ponizeni("amt_zustatek") > 0 Then
                textace("debet") = textace("debet") & _
                                   " poníena o kladnı zùstatek na bìném úètu ve vıši " & VlozCastku(ponizeni("amt_zustatek")) & _
                                   " a o pøíchozí platbu ve vıši " & VlozCastku(ponizeni("amt_prichozi")) & "." & Chr(10)
            ElseIf ponizeni("amt_prichozi") > 0 Then
                textace("debet") = textace("debet") & _
                                   " poníena o pøíchozí platbu ve vıši " & VlozCastku(ponizeni("amt_prichozi")) & "." & Chr(10)
            ElseIf ponizeni("amt_zustatek") > 0 Then
                textace("debet") = textace("debet") & _
                                   " poníena o kladnı zùstatek na bìném úètu ve vıši " & VlozCastku(ponizeni("amt_zustatek")) & "." & Chr(10)
            End If
            
        Else
            textace("debet") = textace("debet") & "." & Chr(10)
        End If
        
    End If
    
nextKateg:

Next

''''   Napokon pridame este riadok s nenaparovanymi stornami, aby na nich nezabudli. Musi to byt extra az pod CASEs lebo chceme vsetky datumy.
If dluzne_datumy.Exists("nenaparovane_storna") Then
    textace("debet") = textace("debet") & _
                       ">>>>>" & textace("dc") & nenaparovane_storna & textace("nenaparovane_storna") & dluzne_datumy("nenaparovane_storna") & ".<<<<<" & Chr(10)
End If




''''   do dukazu pridame polozky podle toho jake transakce u klienta nastali
If pridat_dukaz_vymahani Then
  dukazy = dukazy & _
           dukaz_odeslana_posta & Chr(10) & _
           "Pøehled nákladù úèelnì vynaloenıch bankou na vymáhání dluhu." & Chr(10)
End If

If pridat_dukaz_upominka Then
  dukazy = dukazy & _
           dukaz_odeslana_posta & Chr(10) & _
           dukaz_cenik
ElseIf 1 = 1 Or pridat_dukaz_cenik Then  ' RISK-1054 (koment) - cenik chteji do dukazu pridat vzdy
  dukazy = dukazy & _
           dukaz_cenik
End If



If ucel = "EPR" Then
    skutecnostiDebetText = textace("debet") & Chr(10) & textace("debet_zaver") & Chr(10) & dukazy
ElseIf ucel = "IP" Then
    'extra textace("insol_zaver") nastav a vloz
    If textace("debet") <> "" Then
        skutecnostiDebetText = textace("debet") & Chr(10) & textace("insol_zaver") '& Chr(10) & dukazy
    Else
        skutecnostiDebetText = textace("insol_zaver")
    End If
Else 'nikdy
    assertion False, "jezevèík error 63339" 'nikdy by nemal nastat, funkcia skutecnostiDebetText sa vola iba na dvoch miestach, sablona EPR alebo sablona IP.
    skutecnostiDebetText = textace("debet") & Chr(10)
End If
'''skutecnostiDebetText = textace("debet") & Chr(10) & textace("debet_zaver") & Chr(10) & dukazy
End Function




'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'                                                        vycházím z MainModule, export_sablona_muz
Private Function generuj_epr_skutecnosti_a_dukazy(Loan, SKP)
Dim sql As String, sql2 As String, sql3 As String, sql4 As String, sql7 As String, datum_p As String, skp_ba As String, skp_loan As String
Dim code As String, dodatek As String, dodatek_ucet As String, datum_ucet As String, contract_number As String
Dim ba As String, datum_nacerpani_pujcky As String, amt_loan As String
Dim text1 As String, dukaz1 As String
Dim text21 As String, text222 As String
Dim text_zak_urok_stred As String, text4 As String
Dim rate_interest As Double
Dim hu As String
Dim KuOd, KuDo, KuAmt, name_person
Dim datum_posl_zapl_splatky_po_zes As String
Dim rstFA As ADODB.Recordset
Dim pr As Long
Dim udalost As Variant


On Error GoTo connection_problem
hu = Loan.code_loan_number

'informace o kapitalizovaném úroku
For cI = 1 To Loan.GetCntInterests
    If Loan.GetInterests(cI).UrokTyp = "KU" Then
        KuOd = Loan.GetInterests(cI).StartInterval
        KuDo = Loan.GetInterests(cI).KonecInterval
        KuAmt = Loan.GetInterests(cI).AmtUrok
    End If
Next cI

Workbooks(ThisWorkbook.name).Activate

name_person = c.name


'--------------Struktura zaloby--------------------

'------------------------------------------------------------------------------------------------------------------------------------------------------
'--------------CAST I ---------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------------------
'nejdrive zjistime informace o pujce
sql = "select * from " & schema & ".sablona_epr where code_loan_number = '" & hu & "' and code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING')"
 Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst1.EOF Then
        datum_p = "" & rst1(1) & ""
        skp_ba = rst1(2)
        skp_loan = rst1(3)
        code = rst1(5)
        dodatek = rst1(7)
    End If
rst1.Close

List1.Select
Cells(1, 13).value = skp_ba

'kdy byl zalozen ucet
sql2 = "select * from " & schema & ".sablona_epr where skp_bank_account = " & skp_ba & " and code_completion_subject_type = 'CREATE_BU'"
 Set rst2 = New ADODB.Recordset
   rst2.Open sql2, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst2.EOF Then
        contract_number = rst2(0)
        datum_ucet = "" & rst2(1) & ""
        dodatek_ucet = rst2(7)
        ba = rst2(9)
    End If
rst2.Close


''''kdy byla zalozena RS  ' FA vyhozeno v jez.v.5.8, doplneno do objektu Client
'''sql3 = "select trunc(dtime_signature) as dtime_signature from " & schema & ".sablona_epr where id_person = " & SKP & "  and code_completion_subject_type = 'RS' AND TEXT_CONTRACT_NUMBER = " & c.TextContractNumber
''' Set rst1 = New ADODB.Recordset
'''   rst1.Open sql3, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''   If Not rst1.EOF Then
'''   datum_RS = "" & rst1![Dtime_signature] & ""
'''   End If
'''rst1.Close



datum_nacerpani_pujcky = c.GetLoans(-1, hu).Date_loan_utilization ' "" & rst1(0) & ""
amt_loan = c.GetLoans(-1, hu).Amt_loan_limit ' rst1(1)
rate_interest = c.GetLoans(-1, hu).Rate_Interest_Standard ' rst1(3)


text1 = "I." & Chr(10)
text1 = text1 & "alobce uzavøel s alovanım dne " & c.Date_Contract_Signed & " Rámcovou smlouvu è. " & contract_number & _
    " v souladu s Obchodními podmínkami Air Bank a.s. - èlánek Vznik rámcové smlouvy."

If dodatek = 0 Then
    '---------------------------RS, uver a ucet ve stejny den -------------------
    text1 = text1 & " Na základì Rámcové smlouvy byl alovanému " & _
    "aktivován bìnı úèet è. " & ba & "/3030 a alovanı jej mohl zaèít plnì vyuívat." & _
    " Podpisem Rámcové smlouvy alovanı souèasnì poádal alobce o úvìr ve vıši " & format(amt_loan, "#,###") & ",- Kè, viz " & _
    "Rámcová smlouva è. " & contract_number & " - Smlouva o úvìru è. " & hu & "."
    dukaz1 = "Rámcová smlouva è. " & contract_number & " - Smlouva o úvìru è. " & hu
Else
    If dodatek_ucet = 0 Then
    '---------------------------RS, a ucet ve stejny den, uver jindy -------------------
    text1 = text1 & " Na základì Rámcové smlouvy byl alovanému " & _
    "aktivován bìnı úèet è. " & ba & "/3030 a alovanı jej mohl zaèít plnì vyuívat." & _
    " Dne " & datum_p & " uzavøel alovanı s alobcem Dodatek è. " & dodatek & " k Rámcové smlouvì, jeho podpisem" & _
    " alovanı poádal o úvìr ve vıši " & format(amt_loan, "#,###") & ",- Kè, viz Smlouva o úvìru è. " & hu & "."
    dukaz1 = "Rámcová smlouva è. " & contract_number & Chr(10) & _
    "Dodatek è. " & dodatek & " k Rámcové smlouvì - Smlouva o úvìru è. " & hu
    ElseIf datum_ucet <> datum_p Then
    '---------------------------RS, uver a ucet ve jine dny-------------------
    text1 = text1 & " Dne " & datum_ucet & " uzavøel alovanı s alobcem Dodatek è. " & dodatek_ucet & " k Rámcové smlouvì," & _
    "na základì kterého mu byl zøízen úèet è. " & ba & "/3030 viz Dodatek è. " & dodatek_ucet & " k Rámcové smlouvì - Smlouva o úètu è. " & ba & _
    ". Dne " & datum_p & " uzavøel alovanı s alobcem Dodatek è. " & dodatek & " k Rámcové smlouvì, jeho podpisem" & _
    " alovanı poádal o úvìr ve vıši " & format(amt_loan, "#,###") & ",- Kè, viz Dodatek è. " & dodatek & " k Rámcové smlouvì - Smlouva o úvìru è. " & hu & "."
    dukaz1 = "Rámcová smlouva è. " & contract_number & Chr(10) & _
    "Dodatek è. " & dodatek_ucet & " k Rámcové smlouvì - Smlouva o úètu è. " & ba & Chr(10) & _
    "Dodatek è. " & dodatek & " k Rámcové smlouvì - Smlouva o úvìru è. " & hu
    ElseIf datum_ucet = datum_p Then
    '---------------------------RS v jiny den, ucet a uver stejny den-------------------
    text1 = text1 & " Dne " & datum_p & " uzavøel alovanı s alobcem Dodatek è. " & dodatek & " k Rámcové smlouvì, jeho podpisem" & _
            " alovanı poádal o úvìr ve vıši " & format(amt_loan, "#,###") & ",- Kè a na základì kterého si zøídil úèet è. " & ba & "/3030, viz Dodatek è. " & dodatek & " k Rámcové smlouvì - Smlouva o úvìru è. " & hu & "."
    dukaz1 = "Rámcová smlouva è. " & contract_number & Chr(10) & _
            "Dodatek è. " & dodatek & " k Rámcové smlouvì - Smlouva o úvìru è. " & hu
    End If
End If

'text1 = text1 & " Pøed poskytnutím úvìru alobce v souladu se zákonem posoudil úvìruschopnost alovaného. Dùkazy alobce doplní na vızvu soudu." 'RISK-887
text1 = text1 & " Protoe alovanı splnil všechny smluvní pøedpoklady pro poskytnutí úvìru v souladu s Obchodními podmínkami a Podmínkami pro pouívání úvìru" & _
        " spoleènosti Air Bank a.s., byl mu úvìr è. " & hu & " schválen a dne " & datum_nacerpani_pujcky & " vyplacen na bìnı úèet alovaného v plné vıši, viz pøiloenı vıpis z úètu. " & Chr(10)
dukaz1 = "Dùkaz: " & Chr(10) & dukaz1 & Chr(10) & _
"Obchodní podmínky Air Bank a.s." & Chr(10) & _
"Podmínky pro pouívání úvìru Air Bank a.s." & Chr(10) & _
"Mimoøádnı vıpis z úètu alovaného" & Chr(10) & _
"Formuláø pro standardní informace o spotøebitelském úvìru" & Chr(10)



'------------------------------------------------------------------------------------------------------------------------------------------------------
'--------------CAST II -------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------------
'dle RISK-574 upraveno ve verzi 4.3: 3) odstranit èást textu, kde uvádíme, kolik splátek je na úvìru

Dim text2 As String, dukaz2 As String, sql5 As String, datum_zrizeni_uver As String
Dim splatka1 As Double, d_prav_spl As String
Dim splatka2 As Double, d_prav_spl2 As String
Dim min_datum_splatky As String


sql5 = "select * from " & schema & ".sablona_epr5 where skp_loan = " & skp_loan & " and code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING') and dtime_valid_from = (select max(dtime_valid_from) maxx  from " & schema & ".sablona_epr4 where skp_loan = " & skp_loan & " and code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING') group by skp_loan)"
Set rst1 = New ADODB.Recordset
rst1.Open sql5, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst1.EOF Then
        rst1.MoveFirst
        splatka1 = rst1(2)
        d_prav_spl = rst1(5)
        min_datum_splatky = rst1![min_datum_splatky]
        
        'u nìkterıch loan se to nemusí dobøe napárovat viz RISK-420 - proto chybové upozornìní
    Else
        HlaskaUzivateli ("Nedotáhla se nìkterá data o smlouvì k úvìru è. " & Loan.code_loan_number & ". Je potøeba zkontrolovat a doplnit ruènì. Tıká se textu Skuteènosti (èást II).")
    End If
rst1.Close



'text2 = "II." & Chr(10) & "Podle podmínek úvìrové smlouvy byl alovanı zavázanı mìsíènì hradit pravidelnou konstantní splátku úvìru v celkové vıši " & Isinteger2(splatka1) & _
'        " Kè, skládající se ze splátky jistiny a úroku " & rate_interest & " % p. a., jeho vıše byla sjednaná Smlouvou o úvìru," & _
'        " a to vdy k " & d_prav_spl & ". dni mìsíce poèínaje dnem " & Loan.DateInstalmentFirst & " s tím, e pøedpokládané datum splacení celého úvìru bylo stanoveno" & _
'        " na " & Format(Loan.DateLoanClosePlan, "dd.mm.yyyy") & "." & Chr(10)  ', viz " & chr(10)
text2 = "II." & Chr(10) & "Podle podmínek úvìrové smlouvy byl alovanı zavázanı mìsíènì hradit pravidelnou konstantní splátku úvìru v celkové vıši " & Isinteger2(splatka1) & _
        " Kè, skládající se ze splátky jistiny a úroku " & rate_interest & " % p. a., jeho vıše byla sjednaná Smlouvou o úvìru," & _
        " a to vdy k " & d_prav_spl & ". dni mìsíce poèínaje dnem " & min_datum_splatky & " s tím, e pøedpokládané datum splacení celého úvìru bylo stanoveno" & _
        " na " & format(Loan.DateLoanClosePlan, "dd.mm.yyyy") & "." & Chr(10) & Chr(10)  ', viz " & chr(10)
        'FA RISK-622, FA RISK-574 hodnota Loan.DateInstalmentFirst bola nahradena hodnotou min_datum_splatky
        '

'- nyni je potreba zjistit, zda doslo pred prvni splatkou ke zmene vysi splatky

Dim sql6 As String, date_change_amt As String
sql6 = "select * from " & schema & ".sablona_epr5 where code_completion_subject_type =  'CHANGE_INST_AMOUNT' and dtime_valid_from < min_pujcky and skp_loan = " & skp_loan
Set rst1 = New ADODB.Recordset
   rst1.Open sql6, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   If Not rst1.EOF Then
   rst1.MoveFirst
   For i = 1 To rst1.RecordCount
   If rst1(2) <> splatka1 Then
    text2 = text2 & "alovanı si však pøed první splátkou úvìru zmìnil vıši splátky z " & Isinteger2(splatka1) & " Kè na " & Isinteger2(CDbl(rst1(2))) & _
    " Kè." & Chr(10) & Chr(10)
    If rst1(2) <= datum_nacerpani_pujcky Then
    HlaskaUzivateli "Pozor, datum dodatku je menší rovno naèerpání pùjèky, je moné, e budou chybné datumy v dodatcích! " & Loan.code_loan_number & "."
    End If
     rst1.MoveNext

    End If
    Next i
   End If
rst1.Close




'zde doplnime, pokud doslo ke zmene datumu splatek jeste pred prvni splatkou

Dim sql61 As String
Dim dukaz_dod As String
sql61 = "select * from " & schema & ".sablona_epr5 where code_completion_subject_type =  'CHANGE_INST_DATE' and dtime_valid_from < min_pujcky and skp_loan = " & skp_loan
Set rst1 = New ADODB.Recordset
   rst1.Open sql61, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   If Not rst1.EOF Then
   For i = 1 To rst1.RecordCount
   rst1.MoveFirst
        text2 = text2 & "Dne " & rst1(1) & " uzavøel alovanı s alobcem Dodatek è. " & rst1(9) & " k Rámcové smlouvì, na základì kterého " & _
        "byla provedena zmìna data splátek z " & d_prav_spl & ". dne v mìsíci na " & rst1(5) & ". den v mìsíci. První splátka s touto zmìnou probìhla " & rst1(3) & "."
         dukaz_dod = dukaz_dod & "Zmìnovı dodatek è. " & rst1(9) & ". k Rámcové smlouvì " & Chr(10)
        d_prav_spl = rst1(5)

        If rst1(2) <= datum_nacerpani_pujcky Then
            HlaskaUzivateli "Pozor, datum dodatku je menší rovno naèerpání pùjèky, je moné, e budou chybné datumy v dodatcích! " & Loan.code_loan_number & "."
        End If
        
        
        text2 = text2 & Chr(10) & Chr(10)
        
        rst1.MoveNext
        Next i
    End If
rst1.Close



Worksheets("Sablona_data").Activate
'vsechny dalsi zmeny, vyleju do tabulky, ze ktere to potom vezmu
sql7 = "select * from " & schema & ".sablona_epr5 where skp_loan = '" & skp_loan & "' and code_completion_subject_type in ('CHANGE_INST_AMOUNT','CHANGE_INST_DATE') and dtime_valid_from >= min_pujcky order by dtime_valid_from, text_appendix_number asc"
 Set rst7 = New ADODB.Recordset
    rst7.Open sql7, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst7.EOF Then
        rst7.MoveFirst
         For i = 1 To rst7.RecordCount 'radky
             For j = 1 To 11 'sloupce
                Cells(i, j).value = "" & rst7(j - 1) & ""
              Next j
              rst7.MoveNext
              
              If i > 1 Then
                If format(Cells(i, 8).value, "dd/mmm/yy") = format(Cells(i - 1, 8).value, "dd/mmm/yy") Then
                    HlaskaUzivateli "Pozor - existují dodatky se stejnım datumem u " & Loan.code_loan_number & ". Je moné, e nebudou správnì informace o splátkovém kalendáøi."
                
                End If
               'HlaskaUzivateli Format(Cells(i - 1, 8).Value, "dd/mmm/yy") & " " & Cells(i - 1, 8).Value
              End If


        Next i

   
        'Dne <datum podpisu dodatku>< uzavøela alovaná/uzavøel alovanı> s alobcem Dodatek è. <èíslo dodatku> k Rámcové smlouvì, na základì kterého byla provedena zmìnajeho podpisem <poádala/poádal> o zmìnu vıše pravidelné splátky z <pùvodní vıše splátky>Kè na <nová vıše splátky>Kè. První splátka s touto zmìnou probìhla <datum první splátky se zmìnou vıše splátky>.
   pr = rst7.RecordCount
   rst7.Close


    For i = 1 To pr
    udalost = Cells(i, 9).value
        If udalost = "CHANGE_INST_DATE" Then
            If i = 1 Then
                text2 = text2 & "Dne " & Cells(i, 2).value & " uzavøel alovanı s alobcem Dodatek è. " & Cells(i, 10).value & " k Rámcové smlouvì, na základì kterého " & _
                "byla provedena zmìna data splátek z " & d_prav_spl & ". dne v mìsíci na " & Cells(i, 6).value & ". den v mìsíci. První splátka s touto zmìnou probìhla " & Cells(i, 4).value & "."
                dukaz_dod = dukaz_dod & "Zmìnovı dodatek è. " & Cells(i, 10).value & ". k Rámcové smlouvì " & Chr(10)
          
                text2 = text2 & Chr(10) & Chr(10)
          
            Else
                text2 = text2 & "Dne " & Cells(i, 2).value & " uzavøel alovanı s alobcem Dodatek è. " & Cells(i, 10).value & " k Rámcové smlouvì, na základì kterého " & _
                "byla provedena zmìna data splátek z " & Cells(i - 1, 6).value & ". dne v mìsíci na " & Cells(i, 6).value & ". den v mìsíci. První splátka s touto zmìnou probìhla " & Cells(i, 4).value & "."
                dukaz_dod = dukaz_dod & "Zmìnovı dodatek è. " & Cells(i, 10).value & ". k Rámcové smlouvì " & Chr(10)
            
                text2 = text2 & Chr(10) & Chr(10)
            
            End If
        Else
            If i = 1 Then
                text2 = text2 & "Dne " & Cells(i, 2).value & " uzavøel alovanı s alobcem Dodatek è. " & Cells(i, 10).value & " k Rámcové smlouvì, na základì kterého byla provedena zmìna " & _
                "vıše pravidelné splátky z " & Isinteger2(splatka1) & " Kè na " & Isinteger2(Cells(i, 3).value) & " Kè. První splátka s touto zmìnou probìhla " & Cells(i, 4).value & "."
                
                text2 = text2 & Chr(10) & Chr(10)
                
                dukaz_dod = dukaz_dod & "Zmìnovı dodatek è. " & Cells(i, 10).value & ". k Rámcové smlouvì " & Chr(10)
            
            Else
                text2 = text2 & "Dne " & Cells(i, 2).value & " uzavøel alovanı s alobcem Dodatek è. " & Cells(i, 10).value & " k Rámcové smlouvì, na základì kterého byla provedena zmìna " & _
                "vıše pravidelné splátky z " & Isinteger2(Cells(i - 1, 3).value) & "Kè na " & Isinteger2(Cells(i, 3).value) & "Kè. První splátka s touto zmìnou probìhla " & Cells(i, 4).value & "."
                
                text2 = text2 & Chr(10) & Chr(10)
                
                dukaz_dod = dukaz_dod & "Zmìnovı dodatek è. " & Cells(i, 10).value & ". k Rámcové smlouvì " & Chr(10)
           
            End If
        End If
    Next i
    Range(Cells(1, 1), Cells(pr, 11)).Clear
End If
    
Worksheets("Sablona_data").Visible = False
'Worksheets("Sablona").Activate
'--------------CAST II.II -------------------------------------------------------------------------------------------------------------------------------



'zda a kdy prisla mimoradna splatka

Dim sql11 As String, text2_dukaz As String, text2ms() As String, datumy_ms() As Date
sql11 = "select * from " & schema & ".sablona_epr8 where skp_loan = " & skp_loan
Set rst1 = New ADODB.Recordset
rst1.Open sql11, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst1.EOF Then
        rst1.MoveFirst
        ReDim text2ms(rst1.RecordCount)
        ReDim datumy_ms(rst1.RecordCount)
        For i = 1 To rst1.RecordCount
            text2ms(i) = "Dne " & rst1(13) & " provedl alovanı mimoøádnou splátku úvìru ve vıši " & Isinteger2(CDbl(rst1(12))) & " Kè., viz Zmìnovı dodatek è. " & rst1(7) & " k Rámcové smlouvì. " & Chr(10) & Chr(10)
            text2_dukaz = "Zmìnovı dodatek è. " & rst1(7) & " k Rámcové smlouvì - Provedení mimoøádné splátky úvìru èíslo " & hu
            datumy_ms(i) = rst1(13)
        Next i
    End If
rst1.Close





'musime rozhodnout, zda platil vcas nebo splatky splatil s nekolikadennim zpozdenim
'zde pouze platby do zesplatneni

Dim sql8 As String, datum_posl_zapl_splatky As String, text2ss As String, datum_prvni_splatky As Date

sql8 = "select min(datum_splatky), max(datum_splatky), max(paid_in_time) from " & schema & ".sablona_epr6 where skp_loan = " & skp_loan & " and platba = splatka group by skp_loan"
Set rst8 = New ADODB.Recordset
   rst8.Open sql8, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
      If Not rst8.EOF Then
        rst8.MoveFirst
        
        If rst8(0) = rst8(1) Then
            text2ss = "Splátka " & rst8(0) & " byla splacena. "
        Else
            text2ss = "Splátky od " & rst8(0) & " do " & rst8(1) & " byly splaceny. "
        End If
        
'        If rst8(2) = 0 Then '1 je, pokud byly splatky se zpozdenim
'        text2ss = "Splátky od " & rst8(0) & " do " & rst8(1) & " byly splaceny. "
'        Else
'        text2ss = "Splátky od " & rst8(0) & " do " & rst8(1) & " byly splaceny. "  ' , i kdy ne vdy vèas." ', viz splátková kalendáø úvìru è. " & hu & " a pøehled plateb è. " & hu & "."
'        End If
        datum_posl_zapl_splatky = rst8(1)
        datum_prvni_splatky = rst8(0)
   End If
rst8.Close

'zjistime datum posledni kompletne zaplacene splatky (platby i po zesplatneni)
Dim sql8_1
sql8_1 = "select * from " & schema & ".sablona_epr14 where skp_loan = " & skp_loan
Set rst1 = New ADODB.Recordset
   rst1.Open sql8_1, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
      If Not rst1.EOF Then
        rst1.MoveFirst
        
        datum_posl_zapl_splatky_po_zes = rst1(0)
 
   End If
rst1.Close



'podivame se, kdy bylo datum mimoradne splatky - a podle toho zaradime

If text2_dukaz = "" Then
    text2 = text2 & text2ss
Else

    For i = 1 To UBound(text2ms)
    
        If datumy_ms(i) < datum_prvni_splatky Then
            text2 = text2 & text2ms(i) & text2ss
        Else: text2 = text2 & text2ss & text2ms(i)
        End If
    
    Next i
End If


'nyni se podivame, zda dosla cast nejake splatky pred zesplatnenim nebo ne


Dim sql9 As String, sql10 As String, tedy As String, radne As String, dku As Date  'dz As String,

sql9 = "select * from " & schema & ".sablona_epr7 where skp_loan = " & skp_loan & " order by datum_platby"
Set rst1 = New ADODB.Recordset
rst1.Open sql9, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst1.EOF Then
        rst1.MoveFirst
        For i = 1 To rst1.RecordCount
            text2 = text2 & "Na splátku dne " & rst1(1) & " bylo ze strany alovaného dne " & rst1(2) & " zaplaceno pouze " & Isinteger2(CDbl(rst1(5))) & " Kè, pøièem tuto èástku si alobce zapoèetl èástí ve vıši " & _
                    Isinteger2(CDbl(rst1(3))) & " Kè na úhradu jistiny splátky a èástí ve vıši " & Isinteger2(CDbl(rst1(4))) & " Kè na úhradu úroku splátky. "
                    ', viz splátkovı kalendáø úvìru è. " & hu & " a pøehled plateb è. " & hu & "." 'Následnì neevidujeme ádnou další platbu na úhradu splátek ze strany alované."
    
            If i = 1 Then
                dku = rst1(1)
            End If
                
            rst1.MoveNext
        Next i
        tedy = "tedy" 'slovicko tedy se vlozi, pokud je tam nejaka castecna uhrada splatky
        radne = " øádnì a vèas"
        
    Else 'pokud neprisla zadna castecna splatka
        rst1.Close

    
        sql10 = "select min(date_instalment) from " & schema & ".sp3 where code_loan_number = '" & hu & "' group by code_loan_number "
        Set rst1 = New ADODB.Recordset
        rst1.Open sql10, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
            If Not rst1.EOF Then
                rst1.MoveFirst
                dku = rst1(0)
            Else
                HlaskaUzivateli "Pozor, úvìr " & Loan.code_loan_number & " nemá zaplacenou ádnou splátku!!"
            End If
    End If
rst1.Close

text2 = text2 & "Zápoèet jednotlivıch splátek na jistinu a úrok úvìru viz pøehled plateb è. " & hu & ". "
text2 = text2 & "Poèínaje splátkou ke dni " & dku & " " & tedy & " nebyly splátky ze strany alovaného uhrazeny" & radne & ", viz splátkovı kalendáø úvìru è. " & hu & " a pøehled plateb è. " & hu & "."

text21 = "Následnì neevidujeme ádnou další platbu na úhradu splátek ze strany alovaného." 'pokud prisla nejaka platba, tak se tam toto psat nebude
    

'--------------CAST II.III -------------------------------------------------------------------------------------------------------------------------------
Dim sql12 As String, jistina_zesp As Double, urok_zesp As Double, dukaz_z_d  As String, text22 As String, sql12FA As String
Dim text_jistina_neobsahuje As String


text_jistina_neobsahuje = " (jistina neobsahuje ádné úroky, úroky z prodlení, pokuty, poplatky ani jiné poloky)"

Dim text23 As String, text24 As String
sql12 = "select skp_loan, id_person, nvl(jplatba,0), nvl(uplatba,0), nvl(jpredpis,0), nvl(upredpis,0), nvl(ip,0), nvl(ii,0), platba_den_po_zesplatneni from " & schema & ".sablona_epr9 where skp_loan = " & skp_loan
sql12FA = "select ih.skf_instalment_head from " & schema & ".INSTALMENT_HEAD ih " & _
            " WHERE ih.skp_loan = " & skp_loan & _
            " AND ih.CODE_STATUS = 'a' " & _
            " AND ih.NUM_LOAN_SCHEDULE_ORDER = 0 " & _
            " AND ih.NUM_INSTALMENT_NUMBER = 1 /* prvni splatka */ " & _
            " AND ((ih.amt_payment IS NULL) OR (ih.amt_payment < ih.amt_instalment)) /* nezaplacena (prvni splatka) */ "
Set rst1 = New ADODB.Recordset
rst1.Open sql12, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst1.EOF Then
        
        rst1.MoveFirst
        jistina_zesp = rst1(4) + rst1(6) - rst1(2)
        urok_zesp = rst1(5) + rst1(7) - rst1(3)
        
        If rst1(8) = "Y" Then 'prisla platba den po zesplatneni
            
            'text22 = "Vzhledem k závanému porušení smluvních povinností stranou alovanou bylo v souladu s Obchodními podmínkami a na základì Podmínek pro pouívání úvìru Air Bank a.s. - èl. Nìkolik bodù k ukonèení smlouvy, odst. 3 a " & _
            '" rozhodnuto ze strany alobce o zesplatnìní úvìru  z dùvodu neplnìní smluvních podmínek alovanou stranou, viz zesplatòující dopis - pøedalobní vızva. " ' ze dne " & Day(DataCheck.TextBox1.Text) & "." & Month(DataCheck.TextBox1.Text) & "." & Year(DataCheck.TextBox1.Text) & chr(10)
            ' uprava textu RISK-1054, Jez.v.5.3
            text22 = "Vzhledem k závanému porušení smluvních povinností stranou alovanou pøistoupil alobce v souladu s ujednáním Smlouvy o úvìru è. " & hu & _
                     " (èl. Ostatní ustanovení v závìru dokumentu) ve spojení s Podmínkami pro pouívání úvìru (èl. Nìkolik bodù k ukonèení smlouvy, odst. 3)" & _
                     " dne " & Loan.Date_payoff & " k zesplatnìní úvìru, viz zesplatòující dopis - pøedalobní vızva ze dne " & Loan.Date_payoff & "." & Chr(10)
            dukaz_z_d = "Zesplatòující dopis - pøedalobní vızva" 'musime take upravit podobu dukazu s ohledem na platbu den po zesplatneni
        
        Else
        
            dukaz_z_d = "Zesplatòující dopis - pøedalobní vızva"
            'text22 = "Vzhledem k závanému porušení smluvních povinností stranou alovanou bylo v souladu s Obchodními podmínkami a na základì Podmínek pro pouívání úvìru Air Bank a.s. - èl. Nìkolik bodù k ukonèení smlouvy, odst. 3 a - dne " & Loan.Date_payoff & _
            '" rozhodnuto ze strany alobce o zesplatnìní úvìru  z dùvodu neplnìní smluvních podmínek alovanou stranou, viz zesplatòující dopis - pøedalobní vızva ze dne " & Loan.Date_payoff & "." & Chr(10)
            ' uprava textu RISK-1054, Jez.v.5.3
            text22 = "Vzhledem k závanému porušení smluvních povinností stranou alovanou pøistoupil alobce v souladu s ujednáním Smlouvy o úvìru è. " & hu & _
                     " (èl. Ostatní ustanovení v závìru dokumentu) ve spojení s Podmínkami pro pouívání úvìru (èl. Nìkolik bodù k ukonèení smlouvy, odst. 3)" & _
                     " dne " & Loan.Date_payoff & " k zesplatnìní úvìru, viz zesplatòující dopis - pøedalobní vızva ze dne " & Loan.Date_payoff & "." & Chr(10)
            
            text24 = "Ke dni zesplatnìní tedy byla dluná pohledávka tvoøena nesplacenou jistinou ve vıši " & format(jistina_zesp, "#,##0.00") & " Kè" & _
                     " a nesplacenım smluvním úrokem z úvìru (kapitalizovanı úrok) ve vıši " & format(urok_zesp, "#,##0.00") & " Kè." 'toto se tam da, pokud byla platba po zesplatneni
            
            text23 = "Ke dni zesplatnìní tedy byla dluná pohledávka tvoøena nesplacenou jistinou ve vıši " & format(jistina_zesp, "#,##0.00") & " Kè" & _
                     text_jistina_neobsahuje & _
                     " a nesplacenım smluvním úrokem z úvìru (kapitalizovanı úrok) ve vıši " & format(urok_zesp, "#,##0.00") & " Kè" & _
                     ", kterı je souhrnem nesplacenıch smluvních úrokù ode dne " & _
                     IIf(c.GetLoans(skp_loan).blnUhrazena_alespon_jedna_splatka, "poslední øádnì zaplacené splátky", "naèerpání úvìru") & _
                     ", tj. ode dne " & datum_posl_zapl_splatky & " do dne pøedcházejícího dni zesplatnìní, tedy do dne " & Loan.Date_payoff - 1 & _
                     ", konkrétnì pak (poèet dnù prodlení/365 x aktuální jistina x " & rate_interest & "/100)."
            
'''            '''   zakomentovana cast vyhozena v jez.v.5.9, preslo se na atribut objektu LOAN: c.GetLoans(1).blnUhrazena_alespon_jedna_splatka
'''
'''            text23 = "Ke dni zesplatnìní tedy byla dluná pohledávka tvoøena nesplacenou jistinou ve vıši " & format(jistina_zesp, "#,##0.00") & " Kè" & _
'''                     text_jistina_neobsahuje & _
'''                     " a nesplacenım smluvním úrokem z úvìru (kapitalizovanı úrok) ve vıši " & format(urok_zesp, "#,##0.00") & " Kè" & _
'''                     ", kterı je souhrnem nesplacenıch smluvních úrokù ode dne poslední øádnì zaplacené splátky, tj. ode dne " & datum_posl_zapl_splatky & " do dne pøedcházejícího dni zesplatnìní, tedy do dne " & Loan.Date_payoff - 1 & ", konkrétnì pak (poèet dnù prodlení/365 x aktuální jistina x " & rate_interest & "/100)." & Chr(10) 'pokud nebyla platba po zesplatneni
'''
'''            'FA RISK-574 odtial
'''            'netyka sa pokial prisla platba po zesplatneni (tam sa nevyuziva text23)
'''            Set rstFA = New ADODB.Recordset
'''               rstFA.Open sql12FA, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''               If rstFA.RecordCount > 0 Then  'jiny text pokud klient nezaplatil ani prvni splatku, RISK-574 (bod 1)
'''                  text23 = "Ke dni zesplatnìní tedy byla dluná pohledávka tvoøena nesplacenou jistinou ve vıši " & format(jistina_zesp, "#,##0.00") & " Kè" & _
'''                           text_jistina_neobsahuje & _
'''                           " a nesplacenım smluvním úrokem z úvìru (kapitalizovanı úrok) ve vıši " & format(urok_zesp, "#,##0.00") & " Kè" & _
'''                           ", kterı je souhrnem nesplacenıch smluvních úrokù ode dne naèerpání úvìru, tj. ode dne " & datum_nacerpani_pujcky & " do dne pøedcházejícího dni zesplatnìní, tedy do dne " & Loan.Date_payoff - 1 & ", konkrétnì pak (poèet dnù prodlení/365 x aktuální jistina x " & rate_interest & "/100)." & Chr(10) 'pokud nebyla platba po zesplatneni
'''               End If
'''            rstFA.Close
'''            'FA RISK-574 potial
        
    
        End If
    End If
rst1.Close

Dim sql17 As String, sql13 As String, text25 As String
sql13 = "select * from " & schema & ".sp1 where skp_loan = " & skp_loan & " order by datum_platby, datum_splatky "
Set rst1 = New ADODB.Recordset
rst1.Open sql13, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
    If Not rst1.EOF Then
    
        rst1.MoveFirst
        For i = 1 To rst1.RecordCount
            If i = 1 Then
                text222 = "Po zesplatnìní úvìru uhradil alovanı dne " & rst1(2) & " èástku ve vıši " & Isinteger2(CDbl(rst1(13) + rst1(14))) & " Kè, kterou si alobce zapoèetl na úhradu dluné pohledávky, " & _
                          "pøièem tuto èástku si alobce zapoèetl èástí ve vıši " & Isinteger2(CDbl(rst1(13))) & " Kè na úhradu jistiny splátky ze dne " & rst1(1) & " a èástí ve vıši " & Isinteger2(CDbl(rst1(14))) & " Kè na úhradu úroku splátky ze dne " & rst1(1) & "."
            Else
                text222 = text222 & Chr(10) & Chr(10) & "Dále po zesplatnìní úvìru uhradil alovanı dne " & rst1(2) & " èástku ve vıši " & Isinteger2(rst1(13) + rst1(14)) & " Kè, kterou si alobce zapoèetl na úhradu dluné pohledávky, " & _
                          "pøièem tuto èástku si alobce zapoèetl èástí ve vıši " & Isinteger2(CDbl(rst1(13))) & " Kè na úhradu jistiny splátky ze dne " & rst1(1) & " a èástí ve vıši " & Isinteger2(CDbl(rst1(14))) & " Kè na úhradu úroku splátky ze dne " & rst1(1) & "."
            End If
            rst1.MoveNext
        Next i
           
        ''''  zjistime info o aktualni pohledavce
             
        sql17 = "select * from " & schema & ".balance_obs where skp_loan = " & skp_loan
        Set rst2 = New ADODB.Recordset
        rst2.Open sql17, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
            If Not rst2.EOF Then
                rst2.MoveFirst
                

                 
                If rst2![amt_payoff_principal] = 0 Then ' RISK-1249, bod 1
                    text25 = "Aktuální vıše pohledávky je tvoøena nesplacenım smluvním úrokem a zákonnım úrokem z prodlení."
                Else
                    text25 = text25 & "Aktuální vıše pohledávky je tvoøena nesplacenou jistinou ve vıši " & VlozCastku(rst2![amt_payoff_principal]) & _
                            text_jistina_neobsahuje & _
                            " a nesplacenım smluvním úrokem z úvìru do zesplatnìní" & _
                            " (kapitalizovanı úrok) ve vıši " & VlozCastku(rst2(3)) & ", kterı je souhrnem nesplacenıch smluvních úrokù ode dne " & _
                            IIf(IsNull(datum_posl_zapl_splatky_po_zes), "naèerpání úvìru", "poslední zaplacené splátky, tj. ode dne " & KuOd & ",") & _
                            " do dne pøedcházejícího dni zesplatnìní, tedy do dne " & Loan.Date_payoff - 1 & _
                            ", konkrétnì pak (poèet dnù prodlení/365 x aktuální jistina x " & rate_interest & "/100)."
                End If
                 
            End If
        rst2.Close
         
        text2 = text2 & Chr(10) & Chr(10) & _
                text22 & Chr(10) & _
                text24 & Chr(10) & Chr(10) & _
                text222 & Chr(10) & Chr(10) & _
                text25 & Chr(10) & Chr(10) & _
                text21 & Chr(10)
    
    Else   '''  nebyla platba po zesplatneni
    
        text2 = text2 & Chr(10) & Chr(10) & _
                text21 & Chr(10) & Chr(10) & _
                text22 & Chr(10) & _
                text23 & Chr(10)
              
    End If
rst1.Close

text2 = text2 & Chr(10) & "Dùkaz: " & Chr(10) & _
"Splátkovı kalendáø úvìru è. " & hu & Chr(10) & _
"Pøehled plateb úvìru è. " & hu & Chr(10)
If text2_dukaz <> "" Then
    text2 = text2 & text2_dukaz & Chr(10)
End If

text2 = text2 & dukaz_dod ' & Chr(10)


If dukaz_z_d <> "" Then
    text2 = text2 & dukaz_z_d & Chr(10)
End If

'text2 = text2 & "Mimoøádnı vıpis z úètu" & Chr(10) & Chr(10) & "III."
text2 = text2 & Chr(10) & "III."
    

'------------------------------------------------------------------------------------------------------------------------------------------------------
'--------------CAST III -------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------------

Dim text3 As String, t3 As Integer

text3 = ""
text3 = text3 & "O zesplatnìní úvìru informoval alobce alovaného prostøednictvím zesplatòujícího dopisu - pøedalobní vızvy ze dne " & Loan.Date_payoff & _
        ", dále pak také zprávami na e-mailovou adresu alovaného, do jeho internetového bankovnictví a SMS zprávou." & Chr(10) & Chr(10)

text3 = text3 & "Zesplatòujícím dopisem - pøedalobní vızvou byl alovanı vyzván k zaplacení celého dluhu, kterı byl tvoøen dlunou jistinou ve " & _
        "vıši " & format(jistina_zesp, "#,##0.00") & " Kè a nesplacenım smluvním úrokem z úvìru ve vıši " & format(urok_zesp, "#,##0.00") & " Kè, pøièem jejich vıše byla stanovena ke dni " & Loan.Date_payoff - 1 & ". Zesplatnìním se stal závazek splatnım v plné vıši spolu s pøíslušenstvím ke dni " & _
        Loan.Date_payoff & "." & Chr(10) & Chr(10) & "alobce alovanému v zesplatòujícím dopise - pøedalobní vızvì poskytl dodateènou " & _
        " lhùtu k doplacení, alovanı ji však ke splacení nevyuil. Kdyby alovanı v této lhùtì svùj dluh zaplatil, nepoadoval by alobce " & _
        "zaplacení smluvního úroku ani zákonného úroku z prodlení ode dne zesplatnìní do dne úhrady." & Chr(10) & Chr(10)

text3 = text3 & "Smluvní strany si podpisem Rámcové smlouvy sjednaly, viz Obchodní podmínky - èl. Doruèování dokumentù, e alobce doruèuje alovanému dokumenty " & _
        "na adresu alovaného. Dokumenty odeslané poštou se povaují  za " & _
        "doruèené desátım dnem ode dne jejich odeslání."
    

text3 = text3 & Chr(10) & Chr(10) & "Dùkaz: " & Chr(10) & _
dukaz_z_d & Chr(10) & _
"Podací arch " & Chr(10) & _
"Doklad odeslané pošty " & Chr(10)


text3 = text3 & "Obchodní podmínky Air Bank a.s." & Chr(10) & Chr(10) & "IV."

Dim text_smluvni_urok As String
Dim text_zak_urok_uvod As String
Dim sql15 As String
Dim sql16 As String
Dim bln_prvni_zak_urok As Boolean
Dim start_su As Date
Dim bln_do_zaplaceni As Boolean

text_zak_urok_stred = ""
text4 = ""

bln_prvni_zak_urok = True ' logika prepsana v jez.v.5.9, RISK-1249
sql15 = "select a.DATUM_PLATBY, a.DP, a.JISTINA, a.ID_PERSON, a.UVER_CISLO from " & schema & ".sablona_epr11 a where uver_cislo = '" & hu & "' order by datum_platby"
sql16 = "select min(start_su2) from " & schema & ".V_DATACHECK_CREDIT_I where uver_cislo = '" & hu & "' group by uver_cislo "

Set rst2 = New ADODB.Recordset
rst2.Open sql16, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst2.EOF Then
        rst2.MoveFirst
        start_su = rst2(0)
    End If
rst2.Close


Set rst1 = New ADODB.Recordset
rst1.Open sql15, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst1.EOF Then
        rst1.MoveFirst
        
        For i = 1 To rst1.RecordCount
            
            bln_do_zaplaceni = False
            
            '''   nastaveni konce predesleho intervalu pro obchodni i sankcni urok
            'Debug.Print hu, rst1![JISTINA]
            text_smluvni_urok = text_smluvni_urok & " do " & rst1![dp]
            If (rst1![DATUM_PLATBY] >= start_su) And bln_prvni_zak_urok Then
                text_zak_urok_uvod = text_zak_urok_uvod & " do " & rst1![dp]
            Else
                text_zak_urok_stred = text_zak_urok_stred & " do " & rst1![dp]
            End If
            
            '''   novy interval pro obchodni i sankcni urok
            If (rst1![JISTINA] > 0) Then
            
                bln_do_zaplaceni = True
                text_smluvni_urok = text_smluvni_urok & ", spolu se smluvním úrokem " & rate_interest & " % roènì z èástky " & format(rst1![JISTINA], "#,##0.00") & " Kè od " & rst1![DATUM_PLATBY]
                
                If rst1![DATUM_PLATBY] >= start_su Then
                    If bln_prvni_zak_urok Then
                        text_zak_urok_uvod = text_zak_urok_uvod & ", spolu se zákonnım úrokem z prodlení z èástky " & format(rst1![JISTINA], "#,##0.00") & " Kè od " & rst1![DATUM_PLATBY]
                        bln_prvni_zak_urok = False
                    Else
                        text_zak_urok_stred = text_zak_urok_stred & ", spolu se zákonnım úrokem z prodlení z èástky " & format(rst1![JISTINA], "#,##0.00") & " Kè od " & rst1![DATUM_PLATBY]
                    End If
                End If
                
            End If
            
            rst1.MoveNext
        Next i
        
        If bln_do_zaplaceni Then
            text_smluvni_urok = text_smluvni_urok & " do zaplacení"
            text_zak_urok_stred = text_zak_urok_stred & " do zaplacení."
        Else
            text_zak_urok_stred = text_zak_urok_stred & "."
        End If
        
    End If
rst1.Close



Dim sql155 As String, jistina_akt As Double, urok_akt As Double, celkem_akt As Double

sql155 = "select * from " & schema & ".v_datacheck_credit_a where uver_cislo = '" & hu & "'"
Set rst1 = New ADODB.Recordset
rst1.Open sql155, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst1.EOF Then
        jistina_akt = rst1(3)
        urok_akt = rst1(4)
        celkem_akt = jistina_akt + urok_akt
    End If
rst1.Close


'------------------------------------------------------------------------------------------------------------------------------------------------------
'--------------CAST IV -------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------------
Dim jistina_zesp2 As Double, sqlFA As String, tolerance_dni As Integer 'FA RISK-622 odtial
tolerance_dni = 10
sqlFA = "SELECT COUNT(*) FROM " & schema & ".sp1 WHERE Skp_loan = " & skp_loan & " AND datum_platby - date_payoff BETWEEN 0 AND " & tolerance_dni

Set rstFA = New ADODB.Recordset
    rstFA.Open sqlFA, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
    If rstFA(0) >= 1 Then 'mnozstvo splatok zaplatenych
        Dim sqlFA2 As String
        sqlFA2 = "select min(unpaid_principal_po) FROM " & schema & ".sp1 WHERE skp_loan = " & skp_loan & " AND datum_platby - date_payoff BETWEEN 0 AND " & tolerance_dni
        
        Set rst2 = New ADODB.Recordset
            rst2.Open sqlFA2, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
            jistina_zesp2 = rst2(0)
        rst2.Close
    Else
        jistina_zesp2 = jistina_zesp
    End If
rstFA.Close 'FA RISK-622 potial


text_smluvni_urok = "Celkovı dluh vyplıvající z úvìrové smlouvy è. " & hu & " ke dni podání aloby èiní na jistinì èástku ve vıši " & format(jistina_akt, "#,##0.00") & " Kè " & _
         "spolu se smluvním úrokem " & rate_interest & " % roènì z èástky " & format(jistina_zesp, "#,##0.00") & " Kè od " & Loan.Date_payoff & text_smluvni_urok
text_zak_urok_uvod = " a dále zákonnı úrok z prodlení z èástky " & format(jistina_zesp2, "#,##0.00") & " Kè od " & start_su & text_zak_urok_uvod


text4 = text_smluvni_urok & text_zak_urok_uvod & text_zak_urok_stred & " alobce poaduje zákonnı úrok z prodlení a od " & start_su & _
        ", nebo dle interních pravidel poskytuje alobce alovanému ve vızvì k plnìní dodateènou lhùtu."

If urok_akt > 0 Then
    text4 = text4 & " Dále alobce poaduje kapitalizovanı smluvní úrok ve vıši " & format(urok_akt, "#,##0.00") & " Kè za období od " & KuOd & " do " & KuDo & "."
End If


'_____________________________________________________________________________________________________________________________________________________

'uloení celeho textu do vıstupu fce
generuj_epr_skutecnosti_a_dukazy = text1 & Chr(10) & dukaz1 & Chr(10) & text2 & Chr(10) & text3 & Chr(10) & text4

SheetWelcome.Select

Exit Function
connection_problem:

If Mid(Err.Description, 1, 9) = "ORA-03114" Then
err_connection = True
Else
 MsgBox Err.Description
End If

End Function

