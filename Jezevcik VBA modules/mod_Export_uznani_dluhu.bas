Attribute VB_Name = "mod_Export_uznani_dluhu"
Option Explicit

Private elem As Integer
Private detailDluhuUver As String
Private detailDluhuDebet As String
'Private wdFindContinue, wdAlignParagraphJustify, wdAlignParagraphCenter, wdUnderlineSingle, wdFindAsk As Variant '???
'Private namech As Variant
Private fileName As String ', bezSlozky As String
Private fso As Variant



Public Sub export_uznani_dluhu(SKP As Variant)
On Error GoTo connection_problem
Dim uvery As String
Dim nepovoleneDebety As String
Dim name As String, address As String, text As String


'RISK-972 Oprava text se nepromazával a pokud nìkdo generoval vícekrát po sobì tak se poøád nabalovaly hodnoty.
uvery = vbNullString
nepovoleneDebety = vbNullString
text = vbNullString
detailDluhuUver = vbNullString ' Toto jsou private prommìnné
detailDluhuDebet = vbNullString ' Toto jsou private prommìnné
elem = 0 'Aby èíslování zaèínalo od 1

 If IsNull(c.Address_P) Then
        address = ""
        MsgBox "Je nutné doplnit adresu!!"
    Else
        address = c.Address_P_Ulice & " " & c.Address_P_CP & ", " & Left(c.Address_P_PSC, 3) & " " & Right(c.Address_P_PSC, 2) & " " & c.Address_P_Mesto
 End If


'MLI RISK-481
'''''''''''''''''''''
If c.GetCnt_Loans > 0 Then
   
        For cL = 1 To c.GetCnt_Loans
              'tady vybrat jen požadované úvìry - ovìøit s PV jaké - zøejmì zesplatnìné + k odpisu + odepsané

              If c.GetLoans(cL).Loan_status = "Aktivní" And (c.GetLoans(cL).Loan_wfstatus = "Zesplatnìný" Or c.GetLoans(cL).Loan_wfstatus = "K odpisu" Or c.GetLoans(cL).Loan_wfstatus = "Odepsaný bankou") Then
                elem = elem + 1
                uvery = uvery & Chr(10) & " - úvìr ve výši " & format(c.GetLoans(cL).Amt_loan_principal_total, "#,##0.00") & " Kè, který mi Air Bank " & c.GetLoans(cL).Date_loan_utilization & " poslala na mùj bìžný úèet èíslo " & c.GetLoans(cL).Text_bank_account_number & _
                "/3030 podle smlouvy o spotøebitelském úvìru èíslo " & c.GetLoans(cL).code_loan_number & ". Tu jsem s Air Bank uzavøel " & c.GetLoans(cL).Dtime_signature & ""
               
               detailDluhuUver = detailDluhuUver & Chr(10) & elem & ". u úvìru èíslo " & c.GetLoans(cL).code_loan_number & Chr(10) & _
                                    "- nesplacenou jistinu " & format(c.GetLoans(cL).Amt_loan_principal_unpaid, "#,##0.00") & " Kè" & Chr(10) & _
                                    "- nesplacený smluvní úrok do zesplatnìní " & format(c.GetLoans(cL).GetInterests(c.GetLoans(cL).GetCntInterests).AmtUrok, "#,##0.00") & " Kè" & Chr(10)
                                    
                'èást, která plní uroky
                If c.GetLoans(cL).GetCntInterests > 0 Then
                        
                       'pro všechny intervaly kromì kapitalizovaného
                        For cI = 1 To c.GetLoans(cL).GetCntInterests - 1
                            'pokud je úrok > 0
                            If c.GetLoans(cL).GetInterests(cI).AmtUrok > 0 Then
                                'rozlišení smluvního a zákonného úroku
                                Select Case c.GetLoans(cL).GetInterests(cI).UrokTyp()
                                    Case "SU"
                                        detailDluhuUver = detailDluhuUver & _
                                        "- zákonný úrok z prodlení " & sngZakonnyUrok + c.GetLoans(cL).Rate_Interest_Repo & " %" ' RISK-726
                                        
                                    Case "OU"
                                        detailDluhuUver = detailDluhuUver & _
                                        "- smluvní úrok " & c.GetLoans(cL).Rate_Interest_Standard & " %" ' RISK-726
                                End Select
                                'doplnìní èástky a zaèátku intervalu
                                detailDluhuUver = detailDluhuUver & " roènì z èástky " & format(c.GetLoans(cL).GetInterests(cI).AmtBalancePrincipal, "#,##0.00") & _
                                        " Kè od " & c.GetLoans(cL).GetInterests(cI).StartInterval & " do "
                                'doplnìní konce intervalu
                                If Round(c.GetLoans(cL).Amt_loan_principal_unpaid, 5) <> Round(c.GetLoans(cL).GetInterests(cI).AmtBalancePrincipal, 5) Then
                                    detailDluhuUver = detailDluhuUver & c.GetLoans(cL).GetInterests(cI).KonecInterval & Chr(10)
                                Else
                                    detailDluhuUver = detailDluhuUver & "zaplacení" & Chr(10)
                                
                                End If
                                
                                  
                             End If

                        Next cI
                End If
                
              
              End If
             Next cL
                
    

    End If

''''''''''''''''''''''''''''
'èást, kde se plní NEPOVOLENY DEBET
    If c.GetCnt_Accounts > 0 Then
    
    'vybrat jen požadované úèty - ty co mají nepovolený debet. Pro JPK, spoøicí úèty a úèty v cizích mìnách bude vyskakovat upozornìní pro manuální doplnìní.
        For cA = 1 To c.GetCnt_Accounts
              'tady vybrat jen požadované úèty

              If c.GetAccounts(cA).Amt_balance < 0 And c.GetAccounts(cA).Account_type = "BÚ" And c.GetAccounts(cA).Code_currency = "CZK" Then
                nepovoleneDebety = nepovoleneDebety & "- záporný zùstatek na bìžném úètu èíslo " & c.GetAccounts(cA).Text_bank_account_number & "/3030 ve výši " & format(c.GetAccounts(cA).Amt_balance * (-1), "#,##0.00") & " Kè" & Chr(10)
                elem = elem + 1
                detailDluhuDebet = detailDluhuDebet & elem & ". èástku " & format(c.GetAccounts(cA).Amt_balance * (-1), "#,##0.00") & " Kè, právì o tolik jsem na mém bìžném úètu èíslo " & c.GetAccounts(cA).Text_bank_account_number & _
                "/3030 v minusu" & Chr(10)
              Else
                    If c.GetAccounts(cA).Amt_balance < 0 And c.GetAccounts(cA).Account_type <> "BÚ" Then
                        MsgBox ("Klient má úèet se záporným zùstatkem typu Spoøicí úèet nebo JPK. Je tøeba doplnit pro tento nepovolený debet ruènì všechny údaje.")
                    Else
                        If c.GetAccounts(cA).Amt_balance < 0 And c.GetAccounts(cA).Code_currency <> "CZK" Then
                            MsgBox ("Klient má úèet se záporným zùstatkem v cizí mìnì. Je tøeba doplnit pro tento nepovolený debet ruènì všechny údaje.")
                        End If
                    
                    End If
                
              End If
             Next cA
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim dluh1 As String, dluh2 As String, dluh3 As String, dluh4 As String

If elem = 1 Then
    dluh1 = "UZNÁVÁM SVÙJ DLUH" & Chr(10)
    dluh2 = "svého dluhu"
    dluh3 = "Uznávám tento dluh:"
    dluh4 = "Mùj dluh, jehož výši i dùvody tímto prohlášením uznávám, a který se tímto zavazuji zaplatit, k dnešku zahrnuje:"
Else
    dluh1 = "UZNÁVÁM SVÉ DLUHY" & Chr(10)
    dluh2 = "svých dluhù"
    dluh3 = "Uznávám tyto dluhy:"
    dluh4 = "Mé dluhy, jejichž výši i dùvody tímto prohlášením uznávám, a které se tímto zavazuji zaplatit, k dnešku zahrnují:"
End If

text = dluh1 & Chr(10) & "Já, " & c.Name_first & Space(1) & c.Name_last

If c.Gender = "Z" Then
    text = text & ", narozená "
    Else: text = text & ", narozený "
End If

text = text & c.Date_Birth & ", trvale bytem " & address & ", uznávám dùvody i výši " & dluh2 & _
       ", které mám u Air Bank a.s., IÈ 290 45 371, se sídlem Evropská 2690/17, 160 00 Praha 6 (dále jen jako Air Bank), " & _
       "a to v souladu s ustanovením § 2053 a následujících zákona èíslo 89/2012 Sb.,  obèanského zákoníku." & _
       Chr(10) & Chr(10) & _
       dluh3

text = Chr(10) & Chr(10) & Chr(10) & text & uvery & Chr(10) & nepovoleneDebety & Chr(10) & dluh4 & Chr(10) & detailDluhuUver

text = text & Chr(10) & detailDluhuDebet & Chr(10) & "V .............................. dne ............................. " & Chr(10) & _
Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "    ........................................................." & _
Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "     " & c.Name_first & " " & c.Name_last & Chr(10) & _
Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "          " & "(úøednì ovìøený podpis)"


'-----------------------------------     Práce s Wordem a formátování

Dim wordapp As Object
Dim objDoc As Object
Dim objSelection As Object

Set wordapp = CreateObject("word.Application")
' upraveno v jez.v.5.8 'Set objDoc = wordapp.Documents.Open("W:\Vymahani\LATE\Vzory\Hlavièkový papír uznání dluhu NEMAZAT.dotx") 'oteviram tento dokument, protoze potrebuju zahlavi a zapati ' jez.v.5.7. NOVY (nahradene .docx za .dotx)
On Error GoTo chyba_otevreni_sablony
    Set objDoc = wordapp.Documents.Add(Template:=cfg.Cesta("vzorUznaniDluhu")) 'oteviram tento dokument, protoze potrebuju zahlavi a zapati ' jez.v.5.7. NOVY (nahradene .docx za .dotx)
On Error GoTo connection_problem
Set objSelection = wordapp.Selection

wordapp.Visible = True
wordapp.Activate

'problem - nejde pracovat z excelu ve wordu
'proc? nefunguje oznacit word jako Activedocument, obecne nelze oznacit word ani jako document. Potom nelze odkazovat do wordu pomoci range, nelze editovat text. Lze pouze vlozit pomoci selection
'EDIT: lze to, viz export_hlavickovy_papir


'''   vlozeni CELEHO textu sablony
Call ZarovnejCelyDokumentDoBloku(objDoc)
objSelection.text = text

'''   tady se nastavi BOLD pro vetu na prvnim/druhem radku (pro jednotne cislo)
    
    With objSelection.Find
        .text = "(úøednì ovìøený podpis)"
        .replacement.text = "(úøednì ovìøený podpis)"
        .replacement.Font.Italic = True
        .Forward = True
        .Wrap = cfg.wdFindContinue
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .execute Replace:=cfg.wdReplaceAll 'tento riadok tu musi byt aby to bolo ITALIC
    End With
    
    
'''   tady se nastavi BOLD pro vetu na prvnim/druhem radku (pro jednotne cislo)
        
    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find.replacement.Font
        .Size = 10
        .Bold = True
    End With
    With objSelection.Find.replacement.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
'        .Alignment = cfg.wdAlignParagraphCenter
    End With
    With objSelection.Find
        .text = "uznávám dùvody i výši svého dluhu, které mám u Air Bank a.s."
        .replacement.text = "uznávám dùvody i výši svého dluhu, které mám u Air Bank a.s."
        .Forward = True
        .Wrap = cfg.wdFindContinue
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll
        

'''   tady se nastavi BOLD pro vetu na prvnim/druhem radku (pro mnozne cislo)
    
    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find.replacement.Font
        .Size = 10
        .Bold = True
    End With
    With objSelection.Find.replacement.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .Alignment = cfg.wdAlignParagraphJustify
    End With
    With objSelection.Find
        .text = "uznávám dùvody i výši svých dluhù, které mám u Air Bank a.s."
        .replacement.text = "uznávám dùvody i výši svých dluhù, které mám u Air Bank a.s."
        .Forward = True
        .Wrap = cfg.wdFindContinue
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll
    
'''   tady se nenastavi nic, protoze promenna namech nebyla nikde naplnena (asi se brala z jineho modulu, je mi povedoma.)
'
'    objSelection.Find.ClearFormatting
'    objSelection.Find.replacement.ClearFormatting
'    With objSelection.Find.replacement.Font
'        .Size = 10
'        .Bold = True
'    End With
'    With objSelection.Find.replacement.ParagraphFormat
'        .SpaceBeforeAuto = False
'        .SpaceAfterAuto = False
'        .Alignment = cfg.wdAlignParagraphJustify
'    End With
'    With objSelection.Find
'        .text = namech
'        .replacement.text = namech
'        .Forward = True
'        .Wrap = cfg.wdFindContinue
'        .format = True
'        .MatchCase = False
'        .MatchWholeWord = False
'        .MatchWildcards = False
'        .MatchSoundsLike = False
'        .MatchAllWordForms = False
'    End With
'    objSelection.Find.execute Replace:=cfg.wdReplaceAll


'----------------------------------------------------------------------------------------------------- formatovani nadpisu


    objSelection.Find.ClearFormatting
    With objSelection.Find.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .Alignment = cfg.wdAlignParagraphCenter
    End With
    With objSelection.Find
        .text = "UZNÁVÁM SVÉ DLUHY"
        .replacement.text = "UZNÁVÁM SVÉ DLUHY"
        .Forward = True
        .Wrap = cfg.wdFindContinue
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find.replacement.Font
        .Size = 12
        .Bold = True
        '.Underline = cfg.wdUnderlineSingle
    End With
    With objSelection.Find.replacement.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .Alignment = cfg.wdAlignParagraphCenter
    End With
    With objSelection.Find
        .text = "UZNÁVÁM SVÉ DLUHY"
        .replacement.text = "UZNÁVÁM SVÉ DLUHY"
        .Forward = True
        .Wrap = cfg.wdFindContinue
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll


    objSelection.Find.ClearFormatting
    With objSelection.Find.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .Alignment = cfg.wdAlignParagraphCenter
    End With
    With objSelection.Find
        .text = "UZNÁVÁM SVÉ DLUHY"
        .replacement.text = "UZNÁVÁM SVÉ DLUHY"
        .Forward = True
        .Wrap = cfg.wdFindContinue
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting

    With objSelection.Find.replacement.Font
        .Size = 12
        .Bold = True
'        .Underline = cfg.wdUnderlineSingle
    End With
    With objSelection.Find.replacement.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .Alignment = cfg.wdAlignParagraphCenter
    End With
    With objSelection.Find
        .text = "UZNÁVÁM SVÉ DLUHY"
        .replacement.text = "UZNÁVÁM SVÉ DLUHY"
        .Forward = True
        .Wrap = cfg.wdFindContinue
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll

    '''   jednotne cislo

    objSelection.Find.ClearFormatting
    With objSelection.Find.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .Alignment = cfg.wdAlignParagraphCenter
    End With
    With objSelection.Find
        .text = "UZNÁVÁM SVÙJ DLUH"
        .replacement.text = "UZNÁVÁM SVÙJ DLUH"
        .Forward = True
        .Wrap = cfg.wdFindContinue
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find.replacement.Font
        .Size = 12
        .Bold = True
'        .Underline = cfg.wdUnderlineSingle
    End With
    With objSelection.Find.replacement.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .Alignment = cfg.wdAlignParagraphCenter
    End With
    With objSelection.Find
        .text = "UZNÁVÁM SVÙJ DLUH"
        .replacement.text = "UZNÁVÁM SVÙJ DLUH"
        .Forward = True
        .Wrap = cfg.wdFindContinue
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll


    objSelection.Find.ClearFormatting
    With objSelection.Find.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .Alignment = cfg.wdAlignParagraphCenter
    End With
    With objSelection.Find
        .text = "UZNÁVÁM SVÙJ DLUH"
        .replacement.text = "UZNÁVÁM SVÙJ DLUH"
        .Forward = True
        .Wrap = cfg.wdFindContinue
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find.replacement.Font
        .Size = 12
        .Bold = True
'        .Underline = cfg.wdUnderlineSingle
    End With
    With objSelection.Find.replacement.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .Alignment = cfg.wdAlignParagraphCenter
    End With
    With objSelection.Find
        .text = "UZNÁVÁM SVÙJ DLUH"
        .replacement.text = "UZNÁVÁM SVÙJ DLUH"
        .Forward = True
        .Wrap = cfg.wdFindContinue
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll
'--------------------------------------------------------------------------------- odsazeni u uveru
    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find.replacement.ParagraphFormat
        .LeftIndent = 12
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .FirstLineIndent = -6
        .CharacterUnitLeftIndent = 0
        .CharacterUnitFirstLineIndent = 0
    End With
    With objSelection.Find
        .text = "- smluvní úrok"
        .replacement.text = "- smluvní úrok"
        .Forward = True
        .Wrap = cfg.wdFindAsk
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll

    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find.replacement.ParagraphFormat
        .LeftIndent = 6
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .FirstLineIndent = 0
        .CharacterUnitLeftIndent = 0
        .CharacterUnitFirstLineIndent = 0
    End With
    With objSelection.Find
        .text = "- nesplacenou jistinu"
        .replacement.text = "- nesplacenou jistinu"
        .Forward = True
        .Wrap = cfg.wdFindAsk
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll

    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find.replacement.ParagraphFormat
        .LeftIndent = 6
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .FirstLineIndent = 0
        .CharacterUnitLeftIndent = 0
        .CharacterUnitFirstLineIndent = 0
    End With
    With objSelection.Find
        .text = "- nesplacený smluvní úrok"
        .replacement.text = "- nesplacený smluvní úrok"
        .Forward = True
        .Wrap = cfg.wdFindAsk
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll

    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find.replacement.ParagraphFormat
        .LeftIndent = 6
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .FirstLineIndent = 0
        .CharacterUnitLeftIndent = 0
        .CharacterUnitFirstLineIndent = 0
    End With
    With objSelection.Find
        .text = "- zákonný úrok z"
        .replacement.text = "- zákonný úrok z"
        .Forward = True
        .Wrap = cfg.wdFindAsk
        .format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll


'Call UlozSoubor(cfg.Cesta("EprPodaneZaloby") & c.Nazev_slozky_bez_lomitka, "Uznani dluhu klienta " & c.Nazev_slozky_bez_lomitka & ".docx", "W:\Vymahani\LATE\EPR\Klienti beze složky\")
'Save the XML file 'Pro testovacího jezevèíka se ukláda na W: do složky Testy
 Folder = cfg.Cesta("EprPodaneZaloby") & c.Nazev_slozky_bez_lomitka
 fileName = "Uznani dluhu klienta " & c.Nazev_slozky_bez_lomitka & ".docx"
 
 If schema = "APP_COLL_JEZEVCIK_TEST" Then
        objDoc.SaveAs cfg.Cesta("TestovaciSlozka") & fileName
        ' Call Shell("explorer.exe " & cfg.Cesta("TestovaciSlozka"), vbNormalFocus)
    Else ' na PROD se ukládá do nachystané složky a pokud neexistuje, do složky Klienti bez složky
      
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        If fso.FolderExists(Folder) Then
            objDoc.SaveAs (Folder & "\" & fileName)
            Call Shell("explorer.exe " & Folder, vbNormalFocus)
        Else
            MsgBox "Složka pro klienta " & c.name & " neexistuje." 'pokud složka klienta neexistuje
            objDoc.SaveAs (cfg.Cesta("EprKlientiBezSlozky") & fileName)
            Call Shell("explorer.exe " & cfg.Cesta("EprKlientiBezSlozky"), vbNormalFocus)
                
        End If
    End If


Exit Sub

connection_problem:

    ' V pøípadì, že se jedná o chybu z dùvodu nepøipojení k databázi, bude hodnota promìnné err_connection true a posléze bude proveden reconnect/relogin
    If Mid(Err.Description, 1, 9) = "ORA-03114" Then
        err_connection = True
    Else
        MsgBox Err.Description
    End If

Exit Sub

chyba_otevreni_sablony:

    If (Err.Number = 13 And Err.Description = "Type mismatch") Or (Err.Number = 5981) Or (Err.Description = "Application-defined or object-defined error") Then
        MsgBox "jez-00013: Vypadá to, že nemáš povolené zmìny ve vzoru šablony. Je potøeba šablonu otevøít a zakliknout povolení úprav (v žluté lištì navrchu). " & vbCrLf & _
               "Jedná se o tenhle soubor: " & cfg.Cesta("vzorUznaniDluhu") & vbCrLf & _
               "> v pøípadì, že šablona (.docx soubor) nelze spustit, je potøeba pøes správce úloh ukonèit všechny procesy WINWORD.EXE.", vbInformation, "Povolení úprav šablony"
    Else
        MsgBox "Došlo k neznámé chybì pøi naèítání vzoru šablony."
    End If


End Sub



