Attribute VB_Name = "mod_Export_Podani"
Option Explicit

'Sub export_hlavickovy_papir(skp As Variant)
'' az do verze 5.3 se takhle jmenovala fce export_podani()
'' FA prejmenoval jsem ji protoze mi to vice dava smysl, tohle je tady kvuli zpetne referenci..
'    Call export_podani(skp)
'End Sub



Sub export_podani(SKP As Variant)
' = tlacitko "export podání" na karte EPR (Zaloby)
' az do verze 5.3 se to jmenovalo: Sub export_hlavickovy_papir(skp As Variant)

' veskere podkladove data se berou z tabulek uznani dluhu. EDIT 6.1.2016 + tabulka o vysi dluhu + napojeni na soud

Dim sql1 As String
Dim ii As Integer
Dim objSelection As Object
Dim wdrng As Variant
Dim AAFolder As String, BBFolder As String ', CCFolder As String
Dim fso As Object
Dim name As String
Dim address As String
Dim text1 As String
Dim max As Integer
Dim CUID As Long
Dim commaposition As Integer
Dim Name_first As String
Dim Name_last As String
Dim namech As String
Dim uver As Boolean
Dim nazev_exportovaneho_souboru As String


On Error GoTo connection_problem






sql1 = "select * from " & schema & ".LOAN where NAME_LOAN_STATUS <> 'Ukonèený' and date_loan_utilization < sysdate and id_person = " & SKP & " order by code_loan_number "
Set rst1 = New ADODB.Recordset
rst1.Open sql1, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst1.EOF Then
        rst1.MoveFirst
        name = c.name
        namech = c.Name_first & " " & c.Name_last
        
        
        If IsNull(c.Address_P) Then
            address = ""
            MsgBox "Je nutné doplnit adresu!!"
        Else
            address = c.Address_P
        End If
        
        CUID = c.CUID
        max = rst1.RecordCount
        '        For i = 1 To rst1.RecordCount
        '        text1 = text1 & Chr(10) & " - úvìr ve výši " & Format(rst1(1), "#,##0.00") & " Kè, který mi Air Bank " & rst1(3) & " poslala na mùj bìžný úèet èíslo " & rst1(4) & _
        '        "/3030 podle smlouvy o spotøebitelském úvìru èíslo " & rst1(0) & ". Tu jsem s Air Bank uzavøel" & rst1(10) & " " & rst1(6) & "."
        '        rst1.MoveNext
        '    Next i
        uver = True
    Else
        '     MsgBox "Klient nemá nedoplacený úvìr.", vbInformation
        GoTo Export
        
           
        Exit Sub
    End If
rst1.Close


Dim skp_loan() As Double, sql2 As String, count As Integer, sql3 As String, sql4 As String

sql2 = "select distinct skp_loan, code_loan_number from " & schema & ".uznani_dluhu2 where id_person = " & SKP & " order by code_loan_number "
Set rst2 = New ADODB.Recordset
rst2.Open sql2, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst2.EOF Then
        rst2.MoveFirst
        count = rst2.RecordCount
        ReDim skp_loan(count)
        'MsgBox count
        For i = 1 To count
            skp_loan(i) = rst2(0)
            rst2.MoveNext
        Next i
    End If
rst2.Close

Dim amt_claim As Double, sql21 As String

sql21 = "select amt_claim from " & schema & ".information where id_person = " & SKP

Set rst1 = New ADODB.Recordset
rst1.Open sql21, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst1.EOF Then
        rst1.MoveFirst
        
        If IsNull(rst1(0)) Then
            amt_claim = 0
        Else
            amt_claim = rst1(0)
        End If
    End If
rst1.Close

'cislo jednaci

Dim text As String, textjis As String, textou As String, textsu As String, textku As String, text_sou As String, textou_k As String, textsu_k As String, NextLetter As String
Dim n As Integer, text_part As String, Date_payoff As Date

NextLetter = "a"



'Dulezita poznamka! Sankcni urok je znacen jako textou a obchodni urok je textsu

For i = 1 To count 'pro kazdy uver se udela toto
    sql3 = "select * from " & schema & ".uznani_dluhu2 where skp_loan = " & skp_loan(i)
    Set rst3 = New ADODB.Recordset
    rst3.Open sql3, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst3.EOF Then
        rst3.MoveFirst

        textjis = NextLetter & ") èástku " & format(rst3(3), "#,##0.00") & " Kè"
        NextLetter = Chr(Asc(NextLetter) + 1)
        
        textou = "Zákonný úrok z prodlení ve výši " & sngZakonnyUrok + rst3![Rate_Interest_Repo] & " % roènì z èástky " & format(rst3(13), "#,##0.00") & " Kè od " & rst3(12) ' RISK-726: bolo tu fixne 8,05
        Date_payoff = rst3(11)
        textsu = "Úrok " & rst3(10) & " % roènì z èástky " & format(rst3(13), "#,##0.00") & " Kè od " & rst3(11)
        textku = "Kapitalizovaný úrok ve výši " & format(rst3(4), "#,##0.00") & " Kè"
                
        textou_k = " do zaplacení"
        textsu_k = " do zaplacení"
        
    End If
    rst3.Close
    
    sql4 = "select * from " & schema & ".uznani_dluhu2 where skp_loan = " & skp_loan(i) & " and jis > 0 order by datum_platby"
    Set rst4 = New ADODB.Recordset
    rst4.Open sql4, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst4.EOF Then
        rst4.MoveFirst
        'If rst4.RecordCount > 1 Then
            For ii = 1 To rst4.RecordCount 'pro kazdou platbu po zesplatneni
            
            If IsNull(rst4(9)) Then 'pokud neni datum platby, rovnou to uzavreme
                
                textou = textou & textou_k
                textsu = textsu & textsu_k
                text_sou = textsu & Chr(10) & textou & Chr(10)
                
            Else
                '------------------------------------------------------------------------
                If Date_payoff + 10 > rst4(9) - 1 Then 'pokud je datum zesplatneni vetsi, nez datum platby, nechceme zakonny urok z prodleni
                    textou = ""
                    Else
                    textou = Chr(10) & textou & " do " & rst4(9) - 1 & ""
                End If
                
                textsu = textsu & " do " & rst4(9) - 1 & ""
               
              
                text_sou = text_sou & Chr(10) & textsu & textou
            
                '------------------------------------------------------------------------
            
                textou = "Zákonný úrok z prodlení ve výši " & sngZakonnyUrok + rst4![Rate_Interest_Repo] & " % roènì z èástky " & format(rst4(8), "#,##0.00") & " Kè od " & rst4(9) 'FA XX zakonny urok 'RISK-726 ' RISK-726: bolo tu fixne 8,05
            
                textsu = "Úrok " & rst4(10) & " % roènì z èástky " & format(rst4(8), "#,##0.00") & " Kè od " & rst4(9)
    
                '------------------------------------------------------------------------
    
                If ii = rst4.RecordCount Then
    
                    If Date_payoff + 10 > rst4(9) - 1 Then 'pokud uzavirame cely loan a platba je pred 10denni lhutou, dame datum uroceni do data zesplatneni
                        textou = "Zákonný úrok z prodlení ve výši " & sngZakonnyUrok + rst4![Rate_Interest_Repo] & " % roènì z èástky " & format(rst4(8), "#,##0.00") & " Kè od " & Date_payoff + 10 & " do zaplacení" 'FA XX zakonny urok 'RISK-726 ' RISK-726: bolo tu fixne 8,05

                        Else
                        textou = textou & textou_k
                    End If
                    textsu = textsu & textsu_k
                    text_sou = text_sou & Chr(10) & textsu & Chr(10) & textou & Chr(10)
                End If
            
            'textou = textou & " do " & rst4(9) - 1 & ", dále smluvní úrok  " & rst4(10) & " % roènì z èástky " & Format(rst4(8), "#,##0.00") & " Kè od " & rst4(9)
            'textsu = textsou & " do " & rst4(9) - 1 & ", dále smluvní úrok 8,05 % roènì z èástky " & Format(rst4(8), "#,##0.00") & " Kè od " & rst4(9)

            
            End If
            
            rst4.MoveNext
            Next ii
       ' End If
    End If
    rst4.Close
    text = text & Chr(10) & Chr(10) & textjis & Chr(10) & text_sou & textku
    
    text_sou = ""
Next i


text = text & VlozDebet(NextLetter)

text = text & Chr(10) & Chr(10) & Chr(Asc(NextLetter)) & ") náhradu nákladù øízení"




GoTo Export

Export:


Dim sqlj As String, cj As String


sqlj = "select code from " & schema & ".information where id_person = " & SKP
Set rst1 = New ADODB.Recordset
rst1.Open sqlj, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst1.EOF Then
        rst1.MoveFirst
        If IsNull(rst1(0)) Then
            cj = ""
        Else
            cj = rst1(0)
        End If
    End If
rst1.Close


''''''' RISK-1098 sql nahrazeno za objekt
''''adresa soudu
'''Dim sql_soud As String, adr1 As String, adr2 As String, adr3 As String, ID_soud As String
'''sql_soud = "select  so.* from " & schema & ".person_soud ss join (select  distinct okresni_soud, adresa_p1, adresa_p2, adresa_p3, id_soud from " & schema & ".soudy  ) so on so.okresni_soud = ss.okresni_soud where ss.id_person = " & SKP
'''Set rst1 = New adodb.Recordset
'''rst1.Open sql_soud, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''    If Not rst1.EOF Then
'''        rst1.MoveFirst
'''        adr1 = c.SOUD.adresa_P1
'''        adr2 = c.SOUD.adresa_P2
'''        adr3 = c.SOUD.adresa_P3
'''        ID_soud = c.SOUD.ID_soud
'''    End If
'''rst1.Close


'--------------------------------------------------------otevreni sablony----------------------------------------------------------------

Dim wordapp, objDoc
Set wordapp = CreateObject("word.Application")

' upraveno v jez.v.5.8 'Set objDoc = wordapp.Documents.Open("W:\Vymahani\LATE\Vzory\Hlavièkový papír a petit šablona jezevèík NEMAZAT.docx", ReadOnly:=False)
On Error GoTo chyba_otevreni_sablony
    Set objDoc = wordapp.Documents.Add(Template:=cfg.Cesta("vzorPodani"))
On Error GoTo connection_problem

wordapp.Visible = True
wordapp.Activate


If uver = False Then 'pokud nema uver
address = c.Address_P
namech = c.Name_first & " " & c.Name_last
CUID = c.CUID
address = c.Address_P
'    sqlj = "select name, cuid, address_c from " & schema & ".person where id_person = " & skp
'    Set rst1 = New ADODB.Recordset
'       rst1.Open sqlj, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'       If Not rst1.EOF Then
'        rst1.MoveFirst
'        Name = rst1(0)
'        commaposition = InStr(Name, " ")
'        Name_first = Mid(Name, commaposition + 1)
'        Name_last = Left(Name, commaposition - 1)
'        namech = Name_first & " " & Name_last
'        Cuid = rst1(1)
'        address = c.Address_P
'        End If
'    rst1.Close

End If


   Set objSelection = wordapp.Selection


'vyse pohledavky
If amt_claim > 0 Then
    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find
        .text = "o zaplacení èástky …. Kè s pøíslušenstvím"
        .replacement.text = "o zaplacení èástky " & amt_claim & " Kè s pøíslušenstvím"
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll
End If

'adresa soudu
If c.SOUD.adresa_P1 <> "" Then

    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find
        .text = "Adresa1"
        .replacement.text = c.SOUD.adresa_P1
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll

    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find
        .text = "Adresa2"
        .replacement.text = c.SOUD.adresa_P2
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll
    
        objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find
        .text = "Adresa3"
        .replacement.text = c.SOUD.adresa_P3
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll
    
        objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find
        .text = "ID: ……."
        .replacement.text = "ID: " & c.SOUD.ID_soud
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll
    
Else 'pokud adresa neni dostupna, nahradime teckami


    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find
        .text = "Adresa1"
        .replacement.text = "………………………"
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll


    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find
        .text = "Adresa2"
        .replacement.text = "……………….."
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll


    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find
        .text = "Adresa3"
        .replacement.text = "……………"
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll



End If



    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find
        .text = "V Brnì dne ………"
        .replacement.text = "V Brnì dne " & format(date, "Short Date")
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll

If cj <> "" Then

    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find
        .text = "Sp.zn.: ……………….."
        .replacement.text = "Sp.zn.: " & cj
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll
    
End If

If c.Gender = "Z" Then

    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find
        .text = "Žalované/mu: ……………………, bytem ……………………………………"
        .replacement.text = "Žalované: " & namech & ", bytem " & address
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll
    
    

    
    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find.replacement.Font

        .Bold = True

    End With
    With objSelection.Find
        .text = namech
        .replacement.text = namech
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll
    
    
    Else
    
    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find
        .text = "Žalované/mu: ……………………, bytem ……………………………………"
        .replacement.text = "Žalovanému: " & namech & ", bytem " & address
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll



    objSelection.Find.ClearFormatting
    objSelection.Find.replacement.ClearFormatting
    With objSelection.Find.replacement.Font

        .Bold = True

    End With
    With objSelection.Find
        .text = namech
        .replacement.text = namech
    End With
    objSelection.Find.execute Replace:=cfg.wdReplaceAll
End If

'vkladam primo do sablony

'Set wdRng = objDoc.Range(287 + 27 + Len(cj) - Len("……………….."), 288 + 27 + Len(cj) - Len("……………….."))

If uver = False Then 'rozdeleni, pokud ma uver

    text = Chr(10) & "a) èástku .... Kè" & Chr(10) & Chr(10) & "b) náhradu nákladù øízení"


End If
    Set wdrng = objDoc.Range( _
     Start:=objDoc.Paragraphs(25).Range.Start, _
     End:=objDoc.Paragraphs(26).Range.End)
    
    
    
    wdrng.Select
    
      wdrng.Collapse
      wdrng.InsertAfter (text)


'----------------------------------ulozeni - stejne jako EPR


'
' AFolder = cfg.Cesta("EprPodaneZaloby") & name & " " & Klient.Label2.Caption
' BFolder = cfg.Cesta("EprNachystane") & name & " " & Klient.Label2.Caption
' CFolder = "W:\Vymahani\LATE\EPR\NACHYSTANÉ BØEZEN 2015\" & name & " " & Klient.Label2.Caption

 AAFolder = cfg.Cesta("EprPodaneZaloby") & name & " " & Klient.Label2.Caption & "\"
 BBFolder = cfg.Cesta("EprNachystane") & name & " " & Klient.Label2.Caption & "\"  ' do verze jez.v.5.8: "W:\Vymahani\LATE\EPR\NACHYSTANÉ\"
 'CCFolder = "W:\Vymahani\LATE\EPR\NACHYSTANÉ BØEZEN 2015\" & name & " " & Klient.Label2.Caption
 nazev_exportovaneho_souboru = "Podání " & name & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".docx"

Set fso = CreateObject("Scripting.FileSystemObject")

If schema = "APP_COLL_JEZEVCIK_TEST" Then
    objDoc.SaveAs (cfg.Cesta("TestovaciSlozka") & nazev_exportovaneho_souboru)
    'Call Shell("explorer.exe " & cfg.Cesta("TestovaciSlozka"), vbNormalFocus) 'otvirani slozky na testu
Else
    If fso.FolderExists(AAFolder) Then
        objDoc.SaveAs (AAFolder & nazev_exportovaneho_souboru)
        Call Shell("explorer.exe " & AAFolder, vbNormalFocus)
    ElseIf fso.FolderExists(BBFolder) Then
        objDoc.SaveAs (BBFolder & nazev_exportovaneho_souboru)
        Call Shell("explorer.exe " & BBFolder, vbNormalFocus)
'    ElseIf fso.FolderExists(CCFolder) Then
'        objDoc.SaveAs (CCFolder & "\Podání " & name & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".docx")
'        Call Shell("explorer.exe " & CCFolder, vbNormalFocus)
    Else
        MsgBox "Složka pro klienta " & c.name & " neexistuje." ', vbInformation, "Složka nenalezena!"
        objDoc.SaveAs (cfg.Cesta("EprKlientiBezSlozky") & nazev_exportovaneho_souboru)
        Call Shell("explorer.exe " & cfg.Cesta("EprKlientiBezSlozky"), vbNormalFocus)
    End If
End If
 

        



Set wordapp = Nothing
Set objDoc = Nothing


Exit Sub
connection_problem:

    If Mid(Err.Description, 1, 9) = "ORA-03114" Then
    err_connection = True
    Else
     MsgBox Err.Description
    End If

Exit Sub
chyba_otevreni_sablony:

    If (Err.Number = 13 And Err.Description = "Type mismatch") Or (Err.Number = 5981) Or (Err.Description = "Application-defined or object-defined error") Then
        MsgBox "jez-00013: Vypadá to, že nemáš povolené zmìny ve vzoru šablony. Je potøeba šablonu otevøít a zakliknout povolení úprav (v žluté lištì navrchu). " & vbCrLf & _
               "Jedná se o tenhle soubor: " & cfg.Cesta("vzorZadostDolozka") & vbCrLf & _
               "> v pøípadì, že šablona (.docx soubor) nelze spustit, je potøeba pøes správce úloh ukonèit všechny procesy WINWORD.EXE.", vbInformation, "Povolení úprav šablony"
    Else
        MsgBox "Došlo k neznámé chybì pøi naèítání vzoru šablony."
    End If


End Sub



Private Function VlozDebet(ByRef NextLetter As String) As String
    Dim result As String
    
    If c.GetCnt_Accounts > 0 Then
    
        'vybrat jen požadované úèty - ty co mají nepovolený debet. Pro JPK, spoøicí úèty a úèty v cizích mìnách bude vyskakovat upozornìní pro manuální doplnìní.
        For cA = 1 To c.GetCnt_Accounts
            'tady vybrat jen požadované úvìry - zesplatnìné + k odpisu + odepsané

            If c.GetAccounts(cA).Amt_balance < 0 And c.GetAccounts(cA).Account_type = "BÚ" And c.GetAccounts(cA).Code_currency = "CZK" Then
                
                result = result & Chr(10) & Chr(10) & Chr(Asc(NextLetter)) & ") èástku " & _
                         VlozCastku(-c.GetAccounts(cA).Amt_balance)
                NextLetter = Chr(Asc(NextLetter) + 1)
                
            Else
                If c.GetAccounts(cA).Amt_balance < 0 And c.GetAccounts(cA).Account_type <> "BÚ" Then
                    MsgBox ("Klient má úèet se záporným zùstatkem typu Spoøicí úèet nebo JPK. Je tøeba doplnit pro tento nepovolený debet ruènì všechny údaje do EPR.")
                Else
                    If c.GetAccounts(cA).Amt_balance < 0 And c.GetAccounts(cA).Code_currency <> "CZK" Then
                        MsgBox ("Klient má úèet se záporným zùstatkem v cizí mìnì. Je tøeba doplnit pro tento nepovolený debet ruènì všechny údaje do EPR.")
                    End If
                
                End If
            End If
        Next cA
    
    Else
        
        result = Chr(10) & Chr(10) & Chr(Asc(NextLetter)) & ") èástku .... Kè"
        NextLetter = Chr(Asc(NextLetter) + 1)
        
    End If
    
    VlozDebet = result
End Function


