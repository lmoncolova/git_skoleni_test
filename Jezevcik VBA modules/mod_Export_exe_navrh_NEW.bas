Attribute VB_Name = "mod_Export_exe_navrh_NEW"
Option Explicit

' modul pro sablonu Exekucniho Navrhu, dle RISK-977 (a dalsi mensi tikety, predtym to bol RISK-398)
' vytvoril FA, 15.5.2018 do verze jez.v.5.6


' pro primou praci s objekty Microsoft Word je potrebne pouzivat tzv. Late Binding cez metodu CreateObject() (google it..)


'''  Jak modul/sablona funguje (technicky pohled)
' metody "Napln...", "Nacitej..." a pod. se dotazuji na data do databaze
' tyhle data jsou ulozeny ve vlastnich nadeklarovanych typech (napr. data_pro_sablonu)
' metody "VytvorText..." pak nactene data "vyukladaji" do hezkeho textu
'   - vetsina dekoracnich zmen dle pozadavku uzivatelu je mozne delat tady, pridanim dalsiho formatu do enum "ScitanieCiastok"
' vsechny fixne textace jsou ulozeny v globalni promenne "Textace"
'   - vetsina dekoracnich zmen v textacich by mela byt na tehle urovni
' metoda export_exe_navrh_NEW je ustredni metoda ktera to vsechno zavola



'''   Logika exekucnich navrhu
' ak mame EPR, tak to znamena ze nejaky uradnik dal razitko na nas dokument
' ak mame rozsudok, tak to je novy dokument od sudcu, v ktorom si jednotlive ciastky preusporiada a poscita tak ako sa jemu paci.
' Prave kvoli tomu, ze v rozsudku mozu byt jednotlive pohledavky poscitane roznymi sposobmi, je v pripade rozsudku vyskakovacie okno na uzivatela, kde si vyberie jednu z moznosti z enum ScitanieCiastok
' je to napisane tak, ze pridanie noveho formatu do ScitanieCiastok by nemal byt problem
' u uverov je dolezite, aby sli v spravnom poradi (od nejstarsiho), a zaroven aby bolo poradie zachovane aj v clankoch II a V. - riesene cez ORDER BY vo VIEWs.



'''   nedostatky
' v NaplnDataProSablonu pri plneni uverov beriem do uvahy iba prve EPR. Nie je tam vobec cyklus, ale natvrdo EPR(1)! - preto som nahradil konstantou PoradoveCisloEPR na urovni modulu
'  - az se na to narazi, je moznost vyzkouset upravit konstantu a vygenerovat data k dalsimu EPR. Ale sablona na to nebyla delana.
'  - neresilo se to, protoze jde o historickou vec a tech klientu je zopar a ubyva.
' sablona vubec neumi DEBET. (V casti I se sice nacte z reportu, ale prichozi platby na nej v casti II a tim padem ani cast V proste neumi. Nidky sme to neresili databazove.)


'''   TODO
' presunut zopar stringov do textace()
' eventuelne presunut vsetky textace do Config
' format scitania Scit_Soucet_jistin_a_KU_a_debet v casti II
' VytvorTextCast_I_uvod a VytvorTextCast_V_uvod su hodne podobne a skarede..
' RISK-1096 - search: "rozsu
'    > najsikovnejsie asi bude urobit vlastny type (alebo Class) exekucni_titul, ktory bude obsahovat info zo selectov cast_I_uvod a cast_V_uvod, tiez info z v_exekucni_titul ktore je teraz na klientovi (a k tomu aj genitiv, inak to clovek nenajde)
'    > Array exekucnich titulov sa moze vlozit: do objektu Client
'                                               do data_pro_sablonu (nahradi sa nejak EPR() ktore tam je teraz)
'    > nasledne uz sa dotazujem iba do tejto premennej.
'    > zaroven s tym vyriesit problem:  cl.I  - vykonatelnost datum - ak je rovnake ako datum pravni moci, tak ho neuvadet. Detto cast V.



Private Const PoradoveCisloEPR As Long = 1 ' zatim cela sablona pocita s jedinym EPR na klienta. Hodnota jina nez 1 asi nebyla testovana, takze opatrne..
'Private Const PathVzorEXENavrh As String = "W:\Vymahani\LATE\Vzory\Exekuèní návrh šablona jezevèík_2018_04_13.dotx" 'Pøidat do configu

Private Enum CastExeNavrhu
    CastExeNavrhu_none  ' default
    castexenavrhu_I
    CastExeNavrhu_II
    CastExeNavrhu_V
End Enum


Private Type data_pro_sablonu
    ''' zoznam veci na ktore sa potrebujeme opytat do databaze
    ''' pointa je presunut SQL dotazy na jedno miesto a oddelit tak ziskavanie dat od prace s nimi
    
    flag_stejne As Boolean                    ' from v_exp_bank_claim
    pocet_uveru As Integer
    nazvy_uveru() As String
    poradove_pismeno_pohledavky As Scripting.Dictionary
    poradove_cislo_pohledavky As Scripting.Dictionary
    poradovy_obch_urok_uveru As Scripting.Dictionary
    namech As String
    address As String
    cislo_uctu As String
    klient_udaje As String
    pocetZalob As Long
    ' dblPausal As Double
    EPR() As String   ' EPR po jednom
    epr_all As String ' vsechny EPR za sebou
    datumyZaplacenePohledavky() As Date         ' plni se ve fci VytvorTextCast_II_uvod
    PocetPlatebCast_II As Integer               ' plni se ve fci VytvorTextCast_II_uvod
    castkyZaplacenePohledavky() As Double       ' plni se ve fci VytvorTextCast_II_uvod
End Type


Private Enum ScitanieCiastok
    '''   ak mame rozsudok, tento typ zadefinuje akym sposobom sa vypisu a scitaju ciastky
    Scit_Soucet_none ' defaultni hodnota
    Scit_bez_scitani
    Scit_Soucet_jistin
    Scit_Soucet_jistin_a_KU
    Scit_Soucet_jistin_a_KU_a_debet
End Enum


Private Enum TypPohledavky
    TypPohl_none ' defaultni hodnota
    TypPohl_uver
    TypPohl_debet
    TypPohl_naklady_rizeni
End Enum

Private Enum TypCastky
    '''   U uveru se rozlisuje jistina, obchodni, zakonny(sankcni) a kapitalizovany urok
    TypCastky_none ' defaultni hodnota
    TypCastky_jistina
    TypCastky_KU
    TypCastky_OU
    TypCastky_ZU
End Enum


Private Type Pohledavka
    '''   datovy typ (kolektivni promenna) pro jednotlive pohledavky v casti I a casti V exekucniho navrhu
    enumTypPohl As TypPohledavky
    nazevUveru As String            ' HUxxxxxxx (u jineho typu pohl. se muze nechat "")
    enumTypCastky As TypCastky
    castka As Double
    vyska_uroku As Double
    z_castky As Double
    datum_od As Date
    datum_do As Date
    datum_do_as_string As String    ' 1.1.3000 -> "do zaplaceni"
End Type

Private Type Pohl_zaplacena
    '''   datovy typ (kolektivni promenna) pro jednotlive zaplacene pohledavky v casti II exekucniho navrhu
    enumTypPohl As TypPohledavky
    nazevUveru As String            ' HUxxxxxxx (u jineho typu pohl. se muze nechat "")
    castka_jistina As Double
    castka_urok As Double
    castka_debet As Double
    castka_sop As Double
    datum_zaplaceni As Date
End Type


Private textace As Scripting.Dictionary 'Object
Private ZvolenaFormaScitani As ScitanieCiastok








''''''     MAIN BODY




Public Sub export_exe_navrh_NEW()
    
    '''   hlavni metoda exekucniho navrhu
    '''   nejdrive naplni globalni promenne, pak postupne vola jednotlive casti exekucniho navrhu a vklada je do bookmarks v sablnoe
    
    
    Dim wordapp As Object 'Word.Application
    Dim objDoc As Object 'Word.Document
    Dim arrPohledavky() As Pohledavka
    Dim arrPohl_Zaplacene() As Pohl_zaplacena
    Dim arrPohl_Petit() As Pohledavka
    Dim intPocetPlatebCast_II As Integer
    Dim sablonaData As data_pro_sablonu
    Dim textCast_II As String
    
    
    On Error GoTo connection_problem
    If Not Mame_PM_rozhodnuti(CStr(c.CUID)) Then Exit Sub
    
    
    '''''    inicializace promennych
    
    strMsgbox = ""
    Set textace = NaplnTextace()
    sablonaData = NaplnDataProSablonu
    
    
    '''''     otevreni Word dokumentu
    
    Set wordapp = CreateObject("word.Application")
    On Error GoTo chyba_otevreni_sablony
    Set objDoc = wordapp.Documents.Add(Template:=cfg.Cesta("vzorExeNavrh"))
    On Error GoTo connection_problem
    
    
    '''''     vlozeni udaju o klientovi
    
    Call NaplnBookmark("RamcovaSmlouva", Empty_string_highlighter(c.TextContractNumber), objDoc)
    Call NaplnBookmark("RamcovaSmlouva2", Empty_string_highlighter(c.TextContractNumber), objDoc)
    Call NaplnBookmark("Povinny_NameFull", Empty_string_highlighter(c.name), objDoc, blnBold:=True)
    Call NaplnBookmark("Povinny_narozen_a_adresa", "nar. " & Empty_string_highlighter(c.Date_Birth) & ", trvale bytem " & sablonaData.address, objDoc)
    Call NaplnBookmark("Povereni_cislo_uctu", sablonaData.cislo_uctu, objDoc, blnBold:=True)
    ' podpis zamestnanca
    Call NaplnBookmark("Podpis", JezUserLoggedIn.strSignature, objDoc)
    
    
    '''''     Scitanie Ciastok
    
    ' Pokud mame rozsudek, zeptame se uzivatele jestli chce poscitat jistiny/uroky/...
    If Left(Trim(sablonaData.EPR(PoradoveCisloEPR)), 3) <> "EPR" Then
        ZvolenaFormaScitani = DotazUzivatel_scitani()
        If ZvolenaFormaScitani = Scit_Soucet_none Then Exit Sub
    Else
        ZvolenaFormaScitani = Scit_bez_scitani
    End If
    
    
    '''''     CAST I (PM pohledavka)
    
    arrPohledavky() = NacitejPohledavkyCasti_I_nebo_V(sablonaData, castexenavrhu_I)
    Call NaplnBookmark("CastI", VytvorTextCast_I_uvod(sablonaData) & textace("radek") & VytvorTextCast_I_nebo_V(arrPohledavky, ZvolenaFormaScitani, sablonaData) & ".", objDoc)
    
    
    '''''     Cast II (prichozi platby)
    
    textCast_II = VytvorTextCast_II_uvod(sablonaData)  ' promenna intPocetPlatebCast_II se zde naplni
    arrPohl_Zaplacene() = NacitejPohledavkyCast_II(sablonaData, ZvolenaFormaScitani) ' predtim se musi zavolat VytvorTextCast_II_uvod - kvuli 2 promennym v sablonaData
    Call NaplnBookmark("CastII", textCast_II & VytvorTextCast_II(arrPohl_Zaplacene, ZvolenaFormaScitani, sablonaData), objDoc)
    
    
    '''''    CAST V   (Petit / Povereni)
    
    arrPohl_Petit() = NacitejPohledavkyCasti_I_nebo_V(sablonaData, CastExeNavrhu_V)
    Call NaplnBookmark("Povereni", VytvorTextCast_V_uvod(sablonaData) & textace("radek") & _
                                   VytvorTextCast_I_nebo_V(arrPohl_Petit, ZvolenaFormaScitani, sablonaData) & _
                                   textace("radek"), _
                        objDoc)
    
    
    
    '''''     ULOZENI SABLONY
    
    Call PridajOdrazky(objDoc)
    Call VymazBookmarks(objDoc)
    Call ZarovnejCelyDokumentDoBloku(objDoc) ' jez.v.5.8
    Call UlozDokument(objDoc)
    wordapp.Visible = True
    
    
    '''''     Zahozeni objektu
    
    Set wordapp = Nothing
    Set objDoc = Nothing
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Call VypisVsechnyHlasky
    Exit Sub
    
connection_problem:
    
    ' V pøípadì, že se jedná o chybu z dùvodu nepøipojení k databázi, bude hodnota promìnné err_connection true a posléze bude proveden reconnect/relogin
    If Mid(Err.Description, 1, 9) = "ORA-03114" Then
        err_connection = True
    Else
        MsgBox Err.Description
    End If
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

chyba_otevreni_sablony:

    If (Err.Number = 13 And Err.Description = "Type mismatch") Or (Err.Number = 5981) Or (Err.Description = "Application-defined or object-defined error") Then
        MsgBox "jez-00013: Vypadá to, že nemáš povolené zmìny ve vzoru šablony. Je potøeba šablonu otevøít a zakliknout povolení úprav (v žluté lištì navrchu). " & vbCrLf & _
               "Jedná se o tenhle soubor: " & cfg.Cesta("vzorExeNavrh") & vbCrLf & _
               "> v pøípadì, že šablona (.docx soubor) nelze spustit, je potøeba pøes správce úloh ukonèit všechny procesy WINWORD.EXE.", vbInformation, "Povolení úprav šablony"
    Else
        MsgBox "Došlo k neznámé chybì pøi naèítání vzoru šablony."
    End If
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub






'''''''''''  END OF MAIN BODY



Private Function VlozCastku(ByVal value) As String
    ' pozor, tohle je format jenom pro tuhle sablonu
    ' existuje PUBLIC funkce s rovnakym jmenem ktera generuji trochu jinej format

    VlozCastku = format(value, "#####,##0.00") & " Kè"
    
End Function




Private Function NaplnTextace() As Scripting.Dictionary
    ' sem vsetky textace
    
    Set NaplnTextace = New Scripting.Dictionary
    
    With NaplnTextace
        
        .Add "-", " -   "
        .Add "radek", Chr(10) '& Chr(12)
        
        '.Add "tab", "    "
        
                
        '''  CAST I - uvod
        .Add "I_zaklad", "Na základì "
        .Add "I_zaklad_dal", "Dále na základì "
        .Add "I_vydal", ", který vydal "
        .Add "dne", " dne "
        .Add "dnem", "dnem "
        .Add "I_nabyti_PM", ", který nabyl právní moci dne "
        .Add "I_vykon", " a je vykonatelný "
        .Add "I_uloz_povinnost", " (dále jen exekuèní titul), byla povinnému uložena povinnost zaplatit oprávnìnému:"
        
        
        '''  CAST I
        .Add "I_castka", "èástku "
        .Add "I_urok_zak", "zákonný úrok z prodlení "
        .Add "I_urok_obch", "úrok "
        .Add "I_urok_kap", "kapitalizovaný úrok "
        .Add "I_uroceni", " % roènì z èástky "
        .Add "I_naklady_NR", "náklady nalézacího øízení "
        ' Kè, od, do vložím až v texte..
        
        
        ''' CAST II
        .Add "II_uhradil", "Povinný dobrovolnì uhradil do dne podání tohoto návrhu pouze èást povinnosti uložené pøedmìtným vykonatelným exekuèním titulem, a to:"
        .Add "II_nesplnil", "Povinný dobrovolnì nesplnil do dne podání tohoto návrhu svou povinnost dle pøedmìtného vykonatelného exekuèního titulu, a to ani èásteènì."
        .Add "II_pouze_cast", "Oprávnìný se rozhodl vymáhat pouze èást povinnosti uložené pøedmìtným vykonatelným exekuèním titulem, pøièemž její výše je uvedená v èlánku V. tohoto návrhu."
        .Add "II_castky_zapocet", "Uvedené èástky zapoèetl oprávnìný na úhradu dluhu následovnì: "
        .Add "II_zapocet", " zapoèetl oprávnìný"
        .Add "II_cast", "èástí "
        .Add "II_cast+", " èástí "
        .Add "II_cast_pokr", " a " & .item("II_cast")
        .Add "II_pohl_pism", " pohledávky pod písm. "
        .Add "II_pohl_cislo", " pohledávky è. "
        .Add "II_pohl_urok", " pohledávky úroèené úrokem "
        .Add "II_pohl_urok_rocne", " % roènì a zákonným úrokem z prodlení"
        .Add ")", ") "
        .Add "ze_dne", " ze dne "
        .Add "na_jistinu", " na jistinu"
        .Add "na_urok", " na úrok"
        .Add "na_kap_urok", " na kapitalizovaný úrok"
        .Add "na_naklady_NR", " na náklady nalézacího øízení"
        .Add "na_debet", " na pohledávku pod písm. "
        
        
        '''  CAST V
        .Add "Povereni_uvod", "§ 43a odst. 3 zákona è. 120/2001 Sb., o soudních exekutorech a exekuèní èinnosti (exekuèní øád) " & _
                                "se provedením exekuce na návrh oprávnìného Air Bank a.s., IÈ 29045371, se sídlem Evropská 2690/17, 160 00 Praha 6, proti povinnému "
        
        


    End With
    
End Function




Private Sub NaplnPocetZalob()
'''   naplneni zalob do promenne EPR() 'RISK-398
Dim sql4 As String
Dim rst4 As ADODB.Recordset
Dim lngPocetZalob As Long 'povodne na urovni modulu

    sql4 = "select distinct epr from ap_coll.vs_import_pm_all pm join " & schema & ".person p on p.cuid = pm.cuid and p.flag_valid = 'Y' where p.id_person = " & id_person
    Set rst4 = New ADODB.Recordset
        
    rst4.Open sql4, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        If Not rst4.EOF Then
            
            lngPocetZalob = rst4.RecordCount 'pocet zalob
            ReDim EPR(lngPocetZalob) 'zmena dimenze podle poctu epr
            rst4.MoveFirst
            
            For i = 1 To lngPocetZalob
                EPR(i) = rst4(0)
                rst4.MoveNext
            Next i
            
        End If
    rst4.Close



End Sub






Private Function VytvorTextCast_V_uvod_OLD_beforeEnumTitul(ByRef sd As data_pro_sablonu) As String
    ' vystupem funkce je textovy retezec, ktery je na zacatku Petitu, teda clanku V exekucniho navrhu
    ' text je kombinace fixniho textu, info o klientovi ktere mame v datech pro sablonu, a pak jemna variaci toho co bylo v uvodu casti I
    ' TODO extrahovat hodnoty do NaplnTextace() a odvolavat se pres promennou Textace()
    
    Dim text As String
    Dim i As Integer
    Dim sql11 As String
    Dim rst1 As ADODB.Recordset
    Dim typ_rozsudku As String
    
    
    text = textace("radek")
    
    For i = 1 To sd.pocetZalob 'pro kazde epr doplnime info z rozsudku do hlavicky
    
        If i = 1 Then
            text = text & "Podle "
        Else
            text = text & "Dále podle "
        End If
        
        text = text & textace("Povereni_uvod") & sd.klient_udaje
        
        sql11 = "select soud, epr, datum_vydani, datum_pm, datum_vykon, typ from ap_coll.vs_import_pm_rozsudek_all where epr = '" & sd.EPR(i) & "'"
        Set rst1 = New ADODB.Recordset
        
        rst1.Open sql11, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
            If Not rst1.EOF Then
    
                typ_rozsudku = "" & rst1(5) & ""
                
                If typ_rozsudku = "rozsudek" Then
                    typ_rozsudku = "rozsudku"
                    text = text & ", dle " & typ_rozsudku & " è.j. " & rst1(1) & ", který vydal " & rst1(0) & " dne " & rst1(2) & ", který nabyl právní moci dne " & rst1(3) & " a" & Chr(160) & "je vykonatelný dnem " & rst1(4) & ", uhradit oprávnìnému:" & Chr(10)
                    
                ElseIf typ_rozsudku = "elektronický platební rozkaz" Then
                    typ_rozsudku = "elektronického platebního rozkazu"
                    text = text & ", dle " & typ_rozsudku & " è.j. " & rst1(1) & ", který vydal " & rst1(0) & " dne " & rst1(2) & ", který nabyl právní moci dne " & rst1(3) & " a" & Chr(160) & "je vykonatelný, uhradit oprávnìnému:" & Chr(10)
                
                ''  nasledujuce 2 su iba kvoli preklepom - napr. klient cuid 619818 mal "elekronický platební rozkaz" (chyba tam T)
                ''  pridane ako ElseIf na koniec, pretoze ak recordset nie je prazdny tak by tam mala byt jedna z prvych dvoch moznosti.. tieto sa testuju iba ak dojde k failu zjavnemu
                ElseIf InStr(typ_rozsudku, "rozs") Then
                    typ_rozsudku = "rozsudku"
                    text = text & ", dle " & typ_rozsudku & " è.j. " & rst1(1) & ", který vydal " & rst1(0) & " dne " & rst1(2) & ", který nabyl právní moci dne " & rst1(3) & " a" & Chr(160) & "je vykonatelný dnem " & rst1(4) & ", uhradit oprávnìnému:" & Chr(10)
                
                ElseIf InStr(typ_rozsudku, "elek") Then
                    typ_rozsudku = "elektronického platebního rozkazu"
                    text = text & ", dle " & typ_rozsudku & " è.j. " & rst1(1) & ", který vydal " & rst1(0) & " dne " & rst1(2) & ", který nabyl právní moci dne " & rst1(3) & " a" & Chr(160) & "je vykonatelný, uhradit oprávnìnému:" & Chr(10)
                
                Else
                    HlaskaUzivateli "V petitu chybí èást textu k EPR/rozsudku. Nejspíš pøeklep v datech."
                End If
            
            End If
        rst1.Close
    
    Next i
    
    VytvorTextCast_V_uvod_OLD_beforeEnumTitul = text

End Function







Private Function NaplnDataProSablonu() As data_pro_sablonu
    ' sem vsetky SQL statements (alebo aspon tie, u ktorych sa iba naplni 1-2 hodnoty)
    Dim sql As String
    Dim rst As ADODB.Recordset
        
        With NaplnDataProSablonu
            
            
            '''  kontrola a naplneni udaju o klientovi
            .namech = c.Name_first & " " & c.Name_last
            If IsNull(c.TextContractNumber) Then HlaskaUzivateli "Klient nemá vyplnìné èíslo rámcové smlouvy"
            If IsNull(c.Address_P) Then
                HlaskaUzivateli "Klient nemá vyplnìnou adresu"
            Else
                .address = c.Address_P_Ulice & " " & c.Address_P_CP & ", " & Left(c.Address_P_PSC, 3) & " " & Right(c.Address_P_PSC, 2) & " " & c.Address_P_Mesto
            End If
            .klient_udaje = c.name & ", nar. " & c.Date_Birth & ", trvale bytem " & .address ' vyuziva se ve funkcich co generuji text pro jednotlive casti
    
            
            '''  cislo uctu
            
            sql = "select text_bank_account_number from " & schema & ".exe_account where id_person = " & id_person
            Set rst = New ADODB.Recordset
            rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
                If Not rst.EOF Then
                    .cislo_uctu = "" & rst![Text_bank_account_number] & "/3030"
                End If
            rst.Close

            
                        
            ''' EPR
            'Call NaplnPocetZalob
            
            sql = "select distinct epr from ap_coll.vs_import_pm_all pm join " & schema & ".person p on p.cuid = pm.cuid and p.flag_valid = 'Y' where p.id_person = " & id_person
            Set rst = New ADODB.Recordset
            rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
                
                If Not rst.EOF Then
                    
                    .pocetZalob = rst.RecordCount 'pocet zalob
                    ReDim .EPR(1 To .pocetZalob) 'zmena dimenze podle poctu epr
                    rst.MoveFirst
                    
                    For i = 1 To .pocetZalob
                        .EPR(i) = rst![EPR]
                        rst.MoveNext
                    Next i
                    
                End If
            rst.Close
            
            
            .epr_all = "'" & .EPR(PoradoveCisloEPR) & "'"
            For i = 2 To .pocetZalob
                .epr_all = .epr_all & ", '" & .EPR(i) & "'"
            Next i
            
            
            
            '''  pocet uveru, nazvy uveru a taky jejich poradi
            ' XX v nasl. sql je .EPR(PoradoveCisloEPR), asi treba pridat cyklus okolo
            sql = "select hu, MAX(ou_sazba)*100 as ou_sazba from ap_coll.vs_import_pm_all where typ = 'HU'  and epr = '" & .EPR(PoradoveCisloEPR) & "' and hu is not null GROUP BY hu order by hu" 'zjistime, kolik pro dane epr hu a jak se jmenuji
            Set rst = New ADODB.Recordset
            rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
                
                Set .poradove_pismeno_pohledavky = New Scripting.Dictionary
                Set .poradove_cislo_pohledavky = New Scripting.Dictionary
                Set .poradovy_obch_urok_uveru = New Scripting.Dictionary
                
                If Not rst.EOF Then
                    'With NaplnDataProSablonu
                        .pocet_uveru = rst.RecordCount
                        
                        ReDim .nazvy_uveru(.pocet_uveru)
                        For i = 1 To .pocet_uveru
                            
                            .nazvy_uveru(i) = rst![hu]
                            .poradove_cislo_pohledavky.Add key:=.nazvy_uveru(i), item:=i
                            .poradove_pismeno_pohledavky.Add key:=.nazvy_uveru(i), item:=Chr(96 + i) ' chr(97) = "a", chr(98) = "b", ...
                            .poradovy_obch_urok_uveru.Add key:=.nazvy_uveru(i), item:=CStr(rst![ou_sazba])
                            
                            rst.MoveNext
                        Next i
                    'End With
                End If
            rst.Close
            
            
            '''   Pausal
'            .dblPausal = PausalVyska(c.id_person)
            
        End With
        
        
        
        '''   V pripade, ze se z jakehokoliv duvodu nektera hodnota nenacetla, je potreba to uzivateli zvyraznit
        
        With NaplnDataProSablonu
            .namech = Empty_string_highlighter(.namech, cfg.cst_missing_string_value)
            .address = Empty_string_highlighter(.address, cfg.cst_missing_string_value)
            .cislo_uctu = Empty_string_highlighter(.cislo_uctu, cfg.cst_missing_string_value)
            .klient_udaje = Empty_string_highlighter(.klient_udaje, cfg.cst_missing_string_value)
        End With
        
        
        
End Function



Private Function DotazUzivatel_scitani() As ScitanieCiastok

    Dim Msg As String
    Dim Title As String
    Dim MyInput As Integer
    
        Msg = "Jezevèík zaregistroval rozsudek, chceš posèítat èástky? " & vbNewLine & vbNewLine _
            & "Napiš 1 pokud ne (standardní možnost)," & vbNewLine _
            & "napiš 2 pro sèítání jistin." & vbNewLine _
            & "napiš 3 pro sèítání jistin a sèítání kapitalizovaných úrokù." & vbNewLine _
            & "napiš 4 pro sèítání všeho (jistiny + kap. úroky + debet)." & vbNewLine '_
            '& "Generování šablony ukonèíš klávesou ESC."
        Title = "Volba sèítaní èástek dle rozsudku"
        
        MyInput = Application.InputBox(Msg, Title, Type:=1, Default:=1)
        
        Select Case MyInput
            Case 0
                DotazUzivatel_scitani = Scit_Soucet_none ' tahle moznost ukonci generovani sablony
            Case 1
                DotazUzivatel_scitani = Scit_bez_scitani
            Case 2
                DotazUzivatel_scitani = Scit_Soucet_jistin
            Case 3
                DotazUzivatel_scitani = Scit_Soucet_jistin_a_KU
            Case 4
                DotazUzivatel_scitani = Scit_Soucet_jistin_a_KU_a_debet
            Case Else
                MsgBox "Neplatný vstup. Šablona se vygeneruje bez sèítání.", vbInformation, "Je mi líto"
                DotazUzivatel_scitani = Scit_bez_scitani
        End Select

End Function


Private Function Mame_PM_rozhodnuti(ByVal CUID As String) As Boolean
    ' funkce zkontroluje jestli mame PM rozhodnuti k danemu klientovi
    ' pokud PM rozhodnuti nemame, fce vrati FALSE !!! (v hlavni fci se ukonci export)
    ' pouziva se fce HlaskaUzivateli
    
    ' Dim sql0 As String
    Dim sql As String
    
    
    
'''    '' zahazuji jez.v.5.7
'''    sql0 = "select * from ap_coll.pb_040_summary where cuid = " & CUID
'''
'''    Set rst1 = New ADODB.Recordset 'pokud neni epr evidovane pro daneho klienta, vypneme makro
'''    rst1.Open sql0, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''        If rst1.EOF Then
'''            MsgBox "Klient nemá evidované pravomocné rozhodnutí", vbInformation
'''            Mame_PM_rozhodnuti = False
'''            Exit Function
'''        Else
'''            rst1.MoveFirst
'''            For i = 1 To rst1.RecordCount 'kontrola, zda je vypocet v poradku
'''                If rst1(7) <> "OK" Then
'''                    HlaskaUzivateli "Pozor, podle automatické kontroly je ve výpoètu PM pohledávky chyba nebo je chybnì evidovaná. Problém: " & rst1(7)
'''                    Exit For
'''                End If
'''                rst1.MoveNext
'''
'''            Next i
'''        End If
'''    rst1.Close
    
    

    sql = "select distinct epr from ap_coll.vs_import_pm_all where cuid = " & CUID
    Set rst1 = New ADODB.Recordset
    
    rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
       If rst1.EOF Then
            MsgBox "Klient nemá evidované PM rozhodutí"
            Mame_PM_rozhodnuti = False
            Exit Function
        Else
            Mame_PM_rozhodnuti = True
        End If
    rst1.Close


End Function




Private Function NacitejPohledavkyCasti_I_nebo_V(ByRef sd As data_pro_sablonu, ByVal castNavrhu As CastExeNavrhu) As Pohledavka()
    
    '''   funkce nacte pohledavky bud do casti I exekucniho navrhu, nebo do casti V exekucniho navrhu
    
    Dim arrPohl() As Pohledavka
    Dim intPocetPohledavek As Integer
    Dim rst As ADODB.Recordset, sql As String
    Dim i As Integer
    Dim ii As Integer
    Dim qi As Integer

    
    assertion (castNavrhu <> CastExeNavrhu_none), "Chyba pøi naèítání hodnot exekuèního návrhu."
    assertion (castNavrhu <> CastExeNavrhu_II), "Chyba pøi naèítání hodnot do èástí II exekuèního návrhu."
    
    
    intPocetPohledavek = 300
    ReDim arrPohl(1 To intPocetPohledavek)
    i = 1
    
    
    ''''''   naplnit jednotlive hodnoty
    
    For ii = 1 To sd.pocet_uveru ' loop pres uvery
        
        
        '''   jistina a kapitalizovany urok
        
        Set rst = New ADODB.Recordset
        
        If castNavrhu = castexenavrhu_I Then
            sql = "select nvl(jistina,0) jistina, nvl(ku,0) ku, su_sazba, su_castka, su_dtime_valid_from, ou_sazba, ou_castka, ou_dtime_valid_from, hu" & _
                    " from ap_coll.vs_import_pm_all where typ = 'HU' and hu = '" & sd.nazvy_uveru(ii) & "' and ou_dtime_valid_to > sysdate"
        ElseIf castNavrhu = CastExeNavrhu_V Then
            sql = "SELECT a.HU, nvl(a.JISTINA, 0) AS JISTINA, nvl(a.KU, 0) AS KU, a.OU_SAZBA, a.OU_CASTKA, a.OU_DTIME_VALID_FROM, a.SU_SAZBA, a.SU_CASTKA, a.SU_DTIME_VALID_FROM" & _
                  " FROM " & schema & ".EXE_NAVRH_CAST_V A " & _
                  " WHERE A.TYP = 'HU' " & _
                  " AND A.HU = '" & sd.nazvy_uveru(ii) & "'" & _
                  " AND a.ou_dtime_valid_to > sysdate" ''& " ORDER BY a.hu, A.OU_DTIME_VALID_FROM"
        Else
            assertion 1 = 0, "Jez-44714: chyba ètení dat do èástí V exekuèního návrhu."
        End If
        
        rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic 'doplnime jistinu a k.urok
            If Not rst.EOF Then
                
                With arrPohl(i)
                    .enumTypPohl = TypPohl_uver
                    .nazevUveru = rst![hu]
                    .enumTypCastky = TypCastky_jistina
                    .castka = rst![JISTINA]
                    .vyska_uroku = 0
                    '.z_castky = 0
                    '.datum_od =
                    '.datum_do =
                    '.datum_do_as_string =
                End With
                
                i = i + 1
                
                With arrPohl(i)
                    .enumTypPohl = TypPohl_uver
                    .nazevUveru = rst![hu]
                    .enumTypCastky = TypCastky_KU
                    .castka = rst![ku]
                    .vyska_uroku = 0
                    '.z_castky = 0
                    '.datum_od =
                    '.datum_do =
                    '.datum_do_as_string =
                End With
                
                i = i + 1
            End If
        rst.Close
        
        
        
        
        '''   obchodne a sankcne uroky
        
        If castNavrhu = castexenavrhu_I Then
            sql = "select jistina, ku, su_sazba, su_castka, su_dtime_valid_from, ou_sazba, ou_castka, ou_dtime_valid_from, su_dtime_valid_to, ou_dtime_valid_to " & _
                    " from ap_coll.vs_import_pm_all where typ = 'HU' and hu = '" & sd.nazvy_uveru(ii) & "' order by interval"
        ElseIf castNavrhu = CastExeNavrhu_V Then
            sql = "SELECT a.HU, nvl(a.JISTINA, 0) AS JISTINA, nvl(a.KU, 0) AS KU, a.OU_SAZBA, a.OU_CASTKA, a.OU_DTIME_VALID_FROM, a.OU_DTIME_VALID_TO, a.SU_SAZBA, a.SU_CASTKA, a.SU_DTIME_VALID_FROM, a.SU_DTIME_VALID_TO" & _
                  " FROM " & schema & ".EXE_NAVRH_CAST_V A " & _
                  " WHERE A.TYP = 'HU' " & _
                  " AND A.HU = '" & sd.nazvy_uveru(ii) & "'" & _
                  " order by interval"
        Else
            assertion 1 = 0, "Jez-44715: chyba ètení dat do èástí V exekuèního návrhu."
        End If
        
        
        
        Set rst = New ADODB.Recordset
        rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
            
            If Not rst.EOF Then
                
                For qi = 1 To rst.RecordCount
                    
                    '  obchodni urok
                    If Not (IsNull(rst![ou_castka]) Or IsNull(rst![ou_sazba]) Or IsNull(rst![JISTINA]) Or IsNull(rst![ou_dtime_valid_from]) Or IsNull(rst![ou_dtime_valid_to])) Then
                        With arrPohl(i)
                            .enumTypPohl = TypPohl_uver
                            .nazevUveru = sd.nazvy_uveru(ii)
                            .enumTypCastky = TypCastky_OU
                            .castka = rst![ou_castka]
                            .vyska_uroku = rst![ou_sazba] * 100
                            .z_castky = rst![JISTINA]
                            .datum_od = rst![ou_dtime_valid_from]
                            .datum_do = rst![ou_dtime_valid_to]
                            .datum_do_as_string = IIf(rst![ou_dtime_valid_to] = "1.1.3000", "zaplacení", rst![ou_dtime_valid_to]) ' 1.1.3000 -> "do zaplaceni"
                        End With
                        
                        i = i + 1
                    End If
                    
                    
                    '  zakonny urok z prodleni / sankcni urok
                    If Not (IsNull(rst![su_castka]) Or IsNull(rst![su_sazba]) Or IsNull(rst![JISTINA]) Or IsNull(rst![su_dtime_valid_from]) Or IsNull(rst![su_dtime_valid_to])) Then
                    
                        With arrPohl(i)
                            .enumTypPohl = TypPohl_uver
                            .nazevUveru = sd.nazvy_uveru(ii)
                            .enumTypCastky = TypCastky_ZU
                            .castka = rst![su_castka]
                            .vyska_uroku = rst![su_sazba] * 100
                            .z_castky = rst![JISTINA]
                            .datum_od = rst![su_dtime_valid_from]
                            .datum_do = rst![su_dtime_valid_to]
                            .datum_do_as_string = IIf(rst![su_dtime_valid_to] = "1.1.3000", "zaplacení", rst![su_dtime_valid_to]) ' 1.1.3000 -> "do zaplaceni"
                        End With
                        
                        i = i + 1
                    End If
                    
                    rst.MoveNext
                Next qi
            End If
        rst.Close
        
    
    Next ii
    
    
    
    '''   DEBET a SOP
    
    If castNavrhu = castexenavrhu_I Then
        sql = "select nvl(sum(debet),0) debet, nvl(sum(sop),0) sop from  ap_coll.vs_import_pm_all where epr = '" & sd.EPR(PoradoveCisloEPR) & "'"
    ElseIf castNavrhu = CastExeNavrhu_V Then
            sql = "SELECT nvl(sum(debet),0) debet, nvl(sum(sop),0) sop " & _
                  " FROM " & schema & ".EXE_NAVRH_CAST_V A " & _
                  " WHERE A.TYP in('SOP', 'DEBET') " & _
                  " AND A.EPR = '" & sd.EPR(PoradoveCisloEPR) & "'"
    Else
        assertion 1 = 0, "Jez-44716: chyba ètení dat do èástí V exekuèního návrhu."
    End If
    
    
    Set rst = New ADODB.Recordset
    rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        If Not rst.EOF Then
            
            '  debet
            If rst![debet] > 0 Then
                With arrPohl(i)
                    .enumTypPohl = TypPohl_debet
                    .castka = rst![debet]
                End With
                i = i + 1
            End If
            
            '  sop
            If rst![SOP] > 0 Then
                
                With arrPohl(i)
                    .enumTypPohl = TypPohl_naklady_rizeni
                    
                    '  Pausal v casti V navrhu nevymahame
                    If (castNavrhu = CastExeNavrhu_V) And (c.SOP.Pausal > 0) Then
                        If rst![SOP] > c.SOP.Pausal Then
                            .castka = Round(rst![SOP] - c.SOP.Pausal)
                        Else
                            HlaskaUzivateli "V Petitu nesedí èástky pro náklady øízení (" & rst![SOP] & "), od které po odeètení Paušálu (" & c.SOP.Pausal & ") vznikne záporná pohledávka."
                            .castka = rst![SOP]
                        End If
                    Else
                        .castka = rst![SOP]
                    End If
                    
                    
                End With
                i = i + 1
                
            End If
        End If
    rst.Close


    ReDim Preserve arrPohl(1 To i) 'vyhodim nepotrebne zaznamy (spravne by melo byt (i-1), nechavam jeden navic jako rezervu)
                                   'klicovy je keyword PRESERVE :D guess why
    
    NacitejPohledavkyCasti_I_nebo_V = arrPohl
End Function



Private Function NacitejPohledavkyCast_II(ByRef sd As data_pro_sablonu, ByVal format_scitania As ScitanieCiastok) As Pohl_zaplacena()
    ' naplni se pohledavky, ktere (alespon castecne) byli uhrazeny splatkami.
    ' jde o to, jak byli rozdelene jednotlive platby (co slo na jistinu/urok/.. ktereho uveru)
    ' jeste pred touhle funkci se musi zavolat VytvorTextCast_II_uvod, ve ktere se plni sd.VytvorTextCast_II_uvod a sd.datumyZaplacenePohledavky
    
    
    Dim arrPohlZapl() As Pohl_zaplacena
    Dim intPocetPohledavek As Integer
    Dim rst As ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    Dim pocetPohl As Integer
    Dim blnMameDebet As Boolean
    
    ReDim arrPohlZapl(1 To 300) ' na konci se zkrati
    
    sql = "SELECT a.ID_PERSON, a.DATUM_PLATBY_TRUNC, a.UVER_CISLO, a.amt_payment_jistina, a.amt_payment_urok " & _
          " FROM " & schema & ".V_EXE_NAVRH_UVERY_GROUPED a " & _
          " WHERE a.id_person = " & c.id_person & _
          " ORDER BY a.DATUM_PLATBY_TRUNC, a.UVER_CISLO "
    
    Set rst = New ADODB.Recordset
    rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
        If Not rst.EOF Then
            
            rst.MoveFirst
            
            For i = 1 To rst.RecordCount
                
                ' v tabulce by se NULL nemel vyskytnout, jsou tam nuly
                ' Call assertion(IsNull(rst![amt_payment_jistina]), "jez-94474: hodnota NULL v datech") ' v podkladove tabulce V_EXE_NAVRH_UVERY_GROUPED by zadny NULL nemel byt
                ' Call assertion(IsNull(rst![amt_payment_jistina]), "jez-94475: hodnota NULL v datech")
                Call assertion((nvl(rst![amt_payment_jistina], 0) > 0) Or (nvl(rst![amt_payment_urok], 0) > 0), "jez-94476: nulová pohledávka")
                
                
                With arrPohlZapl(i)
                    .nazevUveru = rst![uver_cislo]
                    .datum_zaplaceni = rst![DATUM_PLATBY_TRUNC]
                    .castka_jistina = rst![amt_payment_jistina]
                    .castka_urok = rst![amt_payment_urok]
                    .castka_debet = 0 ' for clarity
                    .castka_sop = 0   ' for clarity
                End With
                
                rst.MoveNext
                pocetPohl = i
            Next i
        End If
    rst.Close
    
    
    
            
    '''   debet
        
        
        blnMameDebet = False
        If (format_scitania <> Scit_Soucet_jistin_a_KU_a_debet) Then
            
            ' todo XX
            
            blnMameDebet = True ' tohle je potreba zjistit a zapodminkovat ' XX todo
            
            If blnMameDebet Then
                ' kod by mel fungovat, akurat upravit IF podmienku na "kdyz mame debet"
                ' kravina, debet bude mit poradove cislo v zavislosti od formatu scitania
                sd.poradove_cislo_pohledavky.Add "DEBET", sd.pocet_uveru + 1
                sd.poradove_pismeno_pohledavky.Add "DEBET", item:=Chr(96 + sd.pocet_uveru + 1) ' chr(97) = "a", chr(98) = "b", ...
                
            End If
        End If
        
    
    
    
    '''   naklady rizeni
    sql = "SELECT a.ID_PERSON, a.SKP_BANK_ACCOUNT, a.dtime_valid_from_trunc, a.DATUM_OD, a.amt_paid " & _
          " FROM " & schema & ".v_exe_navrh_sopy_grouped a " & _
          " WHERE a.id_person = " & c.id_person & _
          " ORDER BY a.dtime_valid_from_trunc"
    
    Set rst = New ADODB.Recordset
    rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
        If Not rst.EOF Then
            
            rst.MoveFirst
            
            For i = 1 To rst.RecordCount
            
                pocetPohl = pocetPohl + 1
                With arrPohlZapl(pocetPohl)
                    .datum_zaplaceni = rst![dtime_valid_from_trunc]
                    .castka_sop = rst![amt_paid]
                End With
                
                rst.MoveNext
                
            Next i
            
            ' nasledujici 2 radky neni nikdy potreba, sop se necisluje
            sd.poradove_cislo_pohledavky.Add "SOP", sd.pocet_uveru + IIf(blnMameDebet, 2, 1)
            sd.poradove_pismeno_pohledavky.Add "SOP", item:=Chr(96 + sd.pocet_uveru + IIf(blnMameDebet, 2, 1)) ' chr(97) = "a", chr(98) = "b", ...
            
        End If
    rst.Close


    If pocetPohl > 0 Then
        ReDim Preserve arrPohlZapl(1 To pocetPohl)
    End If
    NacitejPohledavkyCast_II = arrPohlZapl

End Function





Private Function VytvorTextCast_I_uvod(ByRef sd As data_pro_sablonu) As String

    Dim i As Integer
    Dim sql As String
    Dim rst1 As ADODB.Recordset
    Dim text1 As String
    Dim text_vydani As String
    
    
    For i = 1 To sd.pocetZalob
    'dynamicky definovane pole, ktere se meni podle poctu epr
        
        '''   Text na zaklade rozsudku/EPR
        
        sql = "select soud, epr, datum_vydani, datum_pm, datum_vykon, typ from ap_coll.vs_import_pm_rozsudek_all where epr = '" & sd.EPR(i) & "'"
        Set rst1 = New ADODB.Recordset
        
        rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
            If Not rst1.EOF Then
            
                                
                If i > 1 Then
                   text1 = text1 & textace("radek") & textace("I_zaklad_dal")
                Else
                   text1 = text1 & textace("I_zaklad")
                End If
                
                text_vydani = textace("I_vydal") & rst1![SOUD] & " è.j. " & rst1![EPR] & " dne " & rst1![datum_vydani] & textace("I_nabyti_PM") & rst1![datum_pm] & textace("I_vykon")
                
                
                ' rozlisujeme formu exekucniho titulu  ' RISK-1096
                If sd.EPR(i) = c.ExeTitul_spisova_znacka Then
                    
                    text1 = text1 & get_enum_exekucni_titul_Name(c.Exekucni_titul, blnGenitiv:=True) _
                                  & text_vydani
                    
                    If (rst1![datum_vykon] <> rst1![datum_pm]) _
                        And ((c.Exekucni_titul = exe_titul_platebni_rozkaz) _
                                Or (c.Exekucni_titul = exe_titul_rozsudek) _
                                Or (c.Exekucni_titul = exe_titul_usneseni_C) _
                            ) Then
                        
                        text1 = text1 & textace("dnem") & rst1![datum_vykon]
                        
                    End If
                    
                    
                Else
                
                    If rst1![typ] = "rozsudek" Then
                        text1 = text1 & "rozsudku" & _
                                text_vydani
                        If (rst1![datum_vykon] <> rst1![datum_pm]) Then
                            text1 = text1 & textace("dnem") & rst1![datum_vykon]
                        End If
                        
                    Else
                        text1 = text1 & "elektronického platebního rozkazu" & _
                                text_vydani
        
                    End If
                
                    text_vydani = ""
                    
                End If
                
                text1 = text1 & textace("I_uloz_povinnost") & textace("radek")
            
            End If
        rst1.Close
    Next i
    
    VytvorTextCast_I_uvod = text1
    
End Function







Private Function VytvorTextCast_I_nebo_V(ByRef arrPohledavky() As Pohledavka, format As ScitanieCiastok, sd As data_pro_sablonu) As String ', ByVal textace As Object)
    ' Funkcia ktora vygeneruje textaciu vsetkych ciastok do navrhu podla zvoleneho formatu (vratane roznych akumulacii)
    ' (v pripade, ze mame rozsudok, sudcovia si casto zmenia format a napr. scitaju vsetky jistiny dokopy)
    ' tato funkcia vezme datovu strukturu a vygeneruje textak podla akehokolvek formatu
    ' nemal by byt problem pridat lubovolny novy format (do ENUM statement a potom nadefinovat v tejto fcii) - zvysok kodu ostane nezmeneny
    ' funkcia pouziva textace pre Export Navrhu (premenna na urovni modulu)
    
    
    ' FA nechceme robit zmeny v premennej arrPohledavky, preto je vytvorena kopia "arrPohl" pre ucely tejto fcie
    ' pozor NEPOUZIVAT premennu "arrPohledavky" - dovod je ten, ze je to referencia na premennu ktora sa este moze pouzivat neskor
    ' normalne by stacilo hodit do deklarace funkce "BYVAL arrPohledavky() ..." , problem je ze u array nelze pouzit ByVal
    
    
    
    
    Dim i As Integer
    Dim text As String
    Dim text_kap_urok As String
    Dim text_zaver As String
    Dim nazev_uveru As String
    Dim blnTextCarka As Boolean
    Dim pohlSoucetJistin As Pohledavka
    Dim pohlSoucetKU As Pohledavka
    Dim pohlSoucetVseho As Pohledavka
    Dim arrPohl() As Pohledavka ' viz komentar v uvodu funkce
    
    ReDim arrPohl(LBound(arrPohledavky) To UBound(arrPohledavky))
    arrPohl = arrPohledavky
        
        
    
        ' text_uvod = VytvorTextCast_I_uvod(sd.EPR) & textace("radek")
        
        If format = Scit_bez_scitani Then
                
                For i = LBound(arrPohl) To UBound(arrPohl)
                    With arrPohl(i)
                        
                        Select Case .enumTypPohl
                            
                            Case TypPohl_uver
                                
                                    If i > 1 And blnTextCarka Then text = text & textace("radek") ' & ","
                                    If i > 1 And .nazevUveru <> nazev_uveru Then
    '                                    If (text_kap_urok = "") And (InStr(Right(text, 2), ",") > 0) Then
    '                                        ' ve specialnim pripade, pokud uver nema zadny kapitalizovany urok, je potreba odstranit posledni carku
    '                                        text = Left(text, Len(text) - 2) & textace("radek")
    '                                    End If
                                        If text_kap_urok <> "" Then
                                            text = text & text_kap_urok & textace("radek")  ' & ","
                                        End If
                                        text = text & textace("radek")
                                        text_kap_urok = ""
                                    End If
                                    
                                    nazev_uveru = .nazevUveru 'abychom mohli sledovat az se prejde na dalsi uver
                                    blnTextCarka = True
                                    
                                    Select Case .enumTypCastky
                                        ' pozor na poradi zalezi
                    
                                        Case TypCastky_jistina
                                            text = text & textace("-") & textace("I_castka") & VlozCastku(.castka)
                    
                                        Case TypCastky_OU
                                            text = text & textace("-") & textace("I_urok_obch") & .vyska_uroku & textace("I_uroceni") & _
                                                   VlozCastku(.z_castky) & " od " & .datum_od & " do " & .datum_do_as_string
                    
                                        Case TypCastky_ZU
                                            text = text & textace("-") & textace("I_urok_zak") & .vyska_uroku & textace("I_uroceni") & _
                                                   VlozCastku(.z_castky) & " od " & .datum_od & " do " & .datum_do_as_string
                    
                                        Case TypCastky_KU
                                            text_kap_urok = text_kap_urok & textace("-") & textace("I_urok_kap") & VlozCastku(.castka)
                                            blnTextCarka = False
                    
                                    End Select
                            
                            
                            Case TypPohl_debet
                                    
                                    'carky pryc: 'If Len(text_zaver) > 1 Then text_zaver = text_zaver & ","
                                    text_zaver = text_zaver & textace("-") & textace("I_castka") & VlozCastku(.castka)
                                    blnTextCarka = False
                                    
                                    
                            Case TypPohl_naklady_rizeni
                                    
                                    If Len(text_zaver) > 1 Then text_zaver = text_zaver & textace("radek") & textace("radek")  ' & ","
                                    text_zaver = text_zaver & textace("-") & textace("I_naklady_NR") & VlozCastku(.castka)
                                    blnTextCarka = False
                                                                
                        End Select
                        
                    End With
                Next i
                
                If text_kap_urok <> "" Then
                    text = text & textace("radek") ' & ","
                    text = text & text_kap_urok & textace("radek") & textace("radek") ' & ","
                    text_kap_urok = ""
                End If
        
        
        
        ElseIf format = Scit_Soucet_jistin Then
            
            'soud seète jistiny všech pohledávek (úvìrù a k nim i debet pokud tam je) do jediné èástky - ta je na zaèátku výroku rozsudku
            'pøíslušenství soud vypíše po jednotlivých úvìrech (1. úrok, 2. úrok z prodlení, 3. kapitalizovaný úrok
            '      - poøadí mùže prohodit nebo dát k sobì smluvní úroky všech úvìrù a úroky z prodlení všech úvìrù - viz. CUID 189888 níže - to je ale jen o pøesunutí jednotlivých øádkù a nejèastìji soudci poøadí zachovají)
            'pøíklad: rozsudek u Leštinová Michaela, CUID 255756
    
            
            For i = LBound(arrPohl) To UBound(arrPohl)
                With arrPohl(i)
                    
                    Select Case .enumTypPohl
                    
                            Case TypPohl_uver
                                
                                    If i > 1 And blnTextCarka Then text = text & textace("radek") ' & ","
                                    If i > 1 And .nazevUveru <> nazev_uveru Then
                                        text = text & text_kap_urok & textace("radek") & textace("radek") ' & ","
                                        text_kap_urok = ""
                                    End If
                                    
                                    nazev_uveru = .nazevUveru 'abychom mohli sledovat az se prejde na dalsi uver
                                    blnTextCarka = True
                                    
                                    Select Case .enumTypCastky
                    
                                        Case TypCastky_jistina
                                            'text = text & textace("-") & textace("I_castka") & VlozCastku(.castka)
                                            pohlSoucetJistin.castka = pohlSoucetJistin.castka + .castka
                                            blnTextCarka = False
                                            
                                        Case TypCastky_OU
                                            text = text & textace("-") & textace("I_urok_obch") & .vyska_uroku & textace("I_uroceni") & _
                                                   VlozCastku(.z_castky) & " od " & .datum_od & " do " & .datum_do_as_string
                    
                                        Case TypCastky_ZU
                                            text = text & textace("-") & textace("I_urok_zak") & .vyska_uroku & textace("I_uroceni") & _
                                                   VlozCastku(.z_castky) & " od " & .datum_od & " do " & .datum_do_as_string
                    
                                        Case TypCastky_KU
                                            text_kap_urok = text_kap_urok & textace("-") & textace("I_urok_kap") & VlozCastku(.castka)
                                            blnTextCarka = False
                    
                                    End Select
                            
                            
                            Case TypPohl_debet
                                    
                                    'text_zaver = text_zaver & textace("-") & textace("I_castka") & VlozCastku(.castka)
                                    pohlSoucetJistin.castka = pohlSoucetJistin.castka + .castka
                                    blnTextCarka = False
                                    
                                    
                            Case TypPohl_naklady_rizeni
                                    
                                    text_zaver = text_zaver & textace("-") & textace("I_naklady_NR") & VlozCastku(.castka)
                                    blnTextCarka = False
                                                   
                    End Select
                                    
                End With
            Next i
            
            If pohlSoucetJistin.castka > 0 Then
                text = textace("-") & textace("I_castka") & VlozCastku(pohlSoucetJistin.castka) & textace("radek") & textace("radek") _
                        & text  ' & ","
            End If
            
            If text_kap_urok <> "" Then
                text = text & textace("radek") & _
                       text_kap_urok & textace("radek") & textace("radek")  ' & "," ' & ","
                text_kap_urok = ""
            End If
            
            
        ElseIf format = Scit_Soucet_jistin_a_KU Then
            
            ' soud seète jistiny všech pohledávek (úvìrù a k nim i debet pokud tam je) do jediné èástky - ta je na zaèátku výroku rozsudku
            ' soud seète kapitalizované úroky všech úvìrù do jediné èástky - bývá hned na druhém místì za souètem jistin
            ' zbývající pøíslušenství soud vypíše po jednotlivých úvìrech (1. úrok, 2. úrok z prodlení - poøadí mùže prohodit nebo dát k sobì smluvní úroky všech úvìrù a úroky z prodlení všech úvìrù)
            ' pøíklad: Fejerèák Petr, CUID 189888 - ještì nepravomocný, PM pohledávka není evidována, dnes byl teprve doruèen, tak o nìm vím že je to tento typ
            ' (prikald predosleho je CUID 255756)
            
            text_zaver = textace("radek")
            
            For i = LBound(arrPohl) To UBound(arrPohl)
                With arrPohl(i)
                    
                    Select Case .enumTypPohl
                    
                            Case TypPohl_uver
                                
                                    If i > 1 And blnTextCarka Then text = text & textace("radek") ' & ","
                                    If i > 1 And .nazevUveru <> nazev_uveru Then
                                        text = text & textace("radek")
                                    End If
                                    
                                    nazev_uveru = .nazevUveru 'abychom mohli sledovat az se prejde na dalsi uver
                                    blnTextCarka = True
                                    
                                    Select Case .enumTypCastky
                    
                                        Case TypCastky_jistina
                                            'text = text & textace("-") & textace("I_castka") & VlozCastku(.castka)
                                            pohlSoucetJistin.castka = pohlSoucetJistin.castka + .castka
                                            blnTextCarka = False
                                            
                                        Case TypCastky_OU
                                            text = text & textace("-") & textace("I_urok_obch") & .vyska_uroku & textace("I_uroceni") & _
                                                   VlozCastku(.z_castky) & " od " & .datum_od & " do " & .datum_do_as_string
                    
                                        Case TypCastky_ZU
                                            text = text & textace("-") & textace("I_urok_zak") & .vyska_uroku & textace("I_uroceni") & _
                                                   VlozCastku(.z_castky) & " od " & .datum_od & " do " & .datum_do_as_string
                    
                                        Case TypCastky_KU
                                            pohlSoucetKU.castka = pohlSoucetKU.castka + .castka
                                            ' text_kap_urok = text_kap_urok & textace("-") & textace("I_urok_kap") & VlozCastku(.castka)
                                            blnTextCarka = False
                    
                                    End Select
                            
                            
                            Case TypPohl_debet
                                    
                                    'text_zaver = text_zaver & textace("-") & textace("I_castka") & VlozCastku(.castka)
                                    pohlSoucetJistin.castka = pohlSoucetJistin.castka + .castka
                                    blnTextCarka = False
                                    
                                    
                            Case TypPohl_naklady_rizeni
                                    
                                    text_zaver = text_zaver & textace("radek") & textace("-") & textace("I_naklady_NR") & VlozCastku(.castka)
                                    blnTextCarka = False
                                                   
                    End Select
                                    
                End With
            Next i
            
            
            ' zadne carky: 'If Len(text) > 1 And Mid(text, Len(text)) <> "," Then text = text & ","
            
            If pohlSoucetJistin.castka > 0 And pohlSoucetKU.castka > 0 Then
                text = textace("-") & textace("I_castka") & VlozCastku(pohlSoucetJistin.castka) & textace("radek") & textace("radek") & _
                       textace("-") & textace("I_urok_kap") & VlozCastku(pohlSoucetKU.castka) & textace("radek") & textace("radek") & _
                       text  ' & "," ' & ","
            ElseIf pohlSoucetJistin.castka > 0 Then
                text = textace("-") & textace("I_castka") & VlozCastku(pohlSoucetJistin.castka) & textace("radek") & textace("radek") & _
                       text ' & ","
            ElseIf pohlSoucetKU.castka > 0 Then
                text = textace("-") & textace("I_urok_kap") & VlozCastku(pohlSoucetKU.castka) & textace("radek") & textace("radek") & _
                       text ' & ","
            End If
            
            
            
            
        ElseIf format = Scit_Soucet_jistin_a_KU_a_debet Then
        
        
            ' soud seète všechny èísla, které tam jsou - jistiny, kapitalizované úroky, debet - do jediné èástky - ta je na zaèátku rozsudku
            ' zbývající pøíslušenství soud vypíše po jednotlivých úvìrech (1. úrok, 2. úrok z prodlení - poøadí mùže prohodit nebo dát k sobì smluvní úroky všech úvìrù a úroky z prodlení všech úvìrù)
            ' pøíklad: nemám
    
            text_zaver = textace("radek")
            
            For i = LBound(arrPohl) To UBound(arrPohl)
                With arrPohl(i)
                    
                    Select Case .enumTypPohl
                    
                            Case TypPohl_uver
                                
                                    If i > 1 And blnTextCarka Then text = text & textace("radek")  ' & ","
                                    If i > 1 And .nazevUveru <> nazev_uveru Then
                                        text = text & textace("radek")
                                    End If
                                    
                                    nazev_uveru = .nazevUveru 'abychom mohli sledovat az se prejde na dalsi uver
                                    blnTextCarka = True
                                    
                                    Select Case .enumTypCastky
                    
                                        Case TypCastky_jistina
                                            'text = text & textace("-") & textace("I_castka") & VlozCastku(.castka)
                                            pohlSoucetJistin.castka = pohlSoucetJistin.castka + .castka
                                            blnTextCarka = False
                                            
                                        Case TypCastky_OU
                                            text = text & textace("-") & textace("I_urok_obch") & .vyska_uroku & textace("I_uroceni") & _
                                                   VlozCastku(.z_castky) & " od " & .datum_od & " do " & .datum_do_as_string
                    
                                        Case TypCastky_ZU
                                            text = text & textace("-") & textace("I_urok_zak") & .vyska_uroku & textace("I_uroceni") & _
                                                   VlozCastku(.z_castky) & " od " & .datum_od & " do " & .datum_do_as_string
                    
                                        Case TypCastky_KU
                                            pohlSoucetKU.castka = pohlSoucetKU.castka + .castka
                                            ' text_kap_urok = text_kap_urok & textace("-") & textace("I_urok_kap") & VlozCastku(.castka)
                                            blnTextCarka = False
                    
                                    End Select
                            
                            
                            Case TypPohl_debet
                                    
                                    'text_zaver = text_zaver & textace("I_castka") & VlozCastku(.castka)
                                    pohlSoucetVseho.castka = pohlSoucetVseho.castka + .castka
                                    blnTextCarka = False
                                    
                                    
                            Case TypPohl_naklady_rizeni
                                    
                                    text_zaver = text_zaver & textace("radek") & textace("-") & textace("I_naklady_NR") & VlozCastku(.castka)
                                    blnTextCarka = False
                                                   
                    End Select
                                    
                End With
            Next i
            
            
            
            pohlSoucetVseho.castka = pohlSoucetVseho.castka + pohlSoucetJistin.castka + pohlSoucetKU.castka
            
            If Len(text) > 1 And Mid(text, Len(text)) <> "," Then text = text  ' & ","
            
            If pohlSoucetVseho.castka > 0 Then
                text = textace("-") & textace("I_castka") & VlozCastku(pohlSoucetVseho.castka) & textace("radek") & textace("radek") & _
                       text  ' & ","
            End If
            
            
            
        End If
    
    VytvorTextCast_I_nebo_V = text & text_zaver

End Function








Private Function VytvorTextCast_II_uvod(sd As data_pro_sablonu) As String
    '''   pozor, fce taky naplni sd.PocetPlatebCast_II
    
    Dim sql As String
    Dim rst As ADODB.Recordset
    Dim text2 As String
    
    
    sql = "SELECT a.ID_PERSON, a.DATUM_PLATBY_TRUNC, a.amt_paid " & _
          " FROM " & schema & ".v_exe_navrh_cast_II_platby a " & _
          " WHERE a.id_person = " & c.id_person
    
    Set rst = New ADODB.Recordset
    rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
        If Not rst.EOF Then
            
            rst.MoveFirst
            
            ReDim sd.datumyZaplacenePohledavky(1 To rst.RecordCount)
            ReDim sd.castkyZaplacenePohledavky(1 To rst.RecordCount)
            sd.PocetPlatebCast_II = rst.RecordCount
            text2 = textace("II_uhradil") & textace("radek") & textace("radek")
            
            For i = 1 To rst.RecordCount
                
                text2 = text2 & textace("-") & "dne " & format(rst![DATUM_PLATBY_TRUNC], "Short Date") & " " & textace("I_castka") & VlozCastku(rst![amt_paid]) & textace("radek") 'puvodne: format$(rst![amt_paid], "#,##0.00") & " Kè," & vbCrLf
                
                sd.datumyZaplacenePohledavky(i) = format(rst![DATUM_PLATBY_TRUNC], "Short Date")
                sd.castkyZaplacenePohledavky(i) = rst![amt_paid]
                
                rst.MoveNext
            Next i
            text2 = text2 & textace("radek") & textace("II_castky_zapocet") '& textace("radek")
            
        Else
            
            text2 = text2 & textace("II_nesplnil")
            
        End If
    rst.Close
    
    VytvorTextCast_II_uvod = text2

End Function




Private Function OdkazNaPoradiPohledavky(sd As data_pro_sablonu, nazev_uveru As String, Optional ByVal blnPouziUrok As Boolean = False) As String
    ' vystupem je neco typu "na pohledávku pod písm. c ) EPR "
    ' vystupem je text, ktery je na konci vetsiny radku v casti II exekucniho navrhu
    
    If blnPouziUrok Then
        Debug.Assert nazev_uveru <> "SOP"
        Debug.Assert nazev_uveru <> "DEBET"
        
        OdkazNaPoradiPohledavky = textace("II_pohl_urok") & Empty_string_highlighter(sd.poradovy_obch_urok_uveru(nazev_uveru)) & textace("II_pohl_urok_rocne")
    Else
        
        If (Left(sd.EPR(PoradoveCisloEPR), 3) = "EPR") Then
            ' u EPR se vklada poradove pismeno
            OdkazNaPoradiPohledavky = textace("II_pohl_pism") & Empty_string_highlighter(sd.poradove_pismeno_pohledavky(nazev_uveru)) & textace(")") & "EPR"
        Else
            ' u rozsudku se vklada poradove cislo
            OdkazNaPoradiPohledavky = textace("II_pohl_cislo") & Empty_string_highlighter(sd.poradove_cislo_pohledavky(nazev_uveru)) & textace(")") & "rozsudku"
        End If
        
    End If
End Function





Private Function VytvorTextCast_II(Pohl_Zaplacene() As Pohl_zaplacena, ByVal format_scitania As ScitanieCiastok, sd As data_pro_sablonu) As String
    
    ' funkce se nazyva "VytvorText...", protoze Vytvori text podle zvolene formy scitani
    ' LOGIKA FUNKCE:
    ' texty sa NEgeneruju zaradom ako ich je vidiet vo vyslednej sablone.
    ' Najprv sa spravi loop cez vsetky prichodzie platby a ku kazdej sa vytvori uvodny kus textacie
    ' Potom sa spravi loop cez Pohl_Zaplacene, a ku kazdej platbe sa text doplni podla toho, ci je tam jedina alebo viac platieb
    ' Az potom sa to spoji do jedneho textu
    ' Na konci sa pripise extra veta ak mame nejaky Pausal
    
    
    Dim arrText() As String
    
    Dim i As Integer
    Dim j As Integer
    Dim blnJistinaVlozena As Boolean
    Dim dblCastkaSoucetKU As Double
    Dim dblSoucetDebetu As Double
    Dim dblSoucetSopu As Double
    Dim pohlSoucetVseho As Pohl_zaplacena
    Dim blnCela_Castka_Sop As Boolean
    Dim blnCela_Castka_Debet As Boolean
    Dim blnAktualni_radek_sop_nebo_debet As Boolean
    Dim blnAktualni_radek_sel_do_souctu As Boolean
    
    
    If format_scitania = Scit_Soucet_jistin_a_KU_a_debet Then
        Debug.Assert False ' tohle je jenom docasne kym se napise kod k tehle variante
        format_scitania = Scit_Soucet_jistin_a_KU
    End If
    
    
    If sd.PocetPlatebCast_II > 0 Then
        ReDim arrText(1 To sd.PocetPlatebCast_II)
        
        
        '''   Uvodna cast textu pre kazdu platbu
        For i = 1 To sd.PocetPlatebCast_II
            arrText(i) = textace("I_castka") & VlozCastku(sd.castkyZaplacenePohledavky(i)) & _
                        textace("ze_dne") & sd.datumyZaplacenePohledavky(i) & textace("II_zapocet")
        Next i
        
        
        '''   Zvysok textu pre kazdu platbu, rozdelenie na jednotlive ucely
        '''   Dle zvolene formy scitani
        
        If format_scitania = Scit_bez_scitani Then
            
            ' jedine v tehle verzi muze byt EPR
            
            For i = 1 To sd.PocetPlatebCast_II
            
                For j = 1 To UBound(Pohl_Zaplacene)
                    
                    With Pohl_Zaplacene(j)
                    
                        If .datum_zaplaceni = sd.datumyZaplacenePohledavky(i) Then
                            
                            blnCela_Castka_Sop = False
                            blnCela_Castka_Debet = False
                            blnAktualni_radek_sop_nebo_debet = False
                            blnJistinaVlozena = False
                            
                            '''   Pripady, kdy cela platba sla na jedinu pohledavku (jistina, urok, naklady rizeni nebo debet)
                            
                            If .castka_jistina = sd.castkyZaplacenePohledavky(i) Then

                                    arrText(i) = arrText(i) & textace("na_jistinu")
                                                 
                            ElseIf .castka_urok = sd.castkyZaplacenePohledavky(i) Then
                            
                                    arrText(i) = arrText(i) & textace("na_kap_urok")
                                                 
                            ElseIf .castka_sop = sd.castkyZaplacenePohledavky(i) Then
                            
                                    arrText(i) = arrText(i) & textace("na_naklady_NR")
                                    blnCela_Castka_Sop = True
                                                 
                            ElseIf .castka_debet = sd.castkyZaplacenePohledavky(i) Then
                            
                                    arrText(i) = arrText(i) & textace("na_debet")
                                    blnCela_Castka_Debet = True
                                             
                                             
                            
                            
                            Else   '''   Else: platba byla rozdelena na nekolik ucelu (probehne vicekrat blok kodu)
                            
                                
                                
                                If Not ((.castka_jistina + .castka_urok + .castka_sop + .castka_debet) = sd.castkyZaplacenePohledavky(i)) Then
                                    ' platba sla na vice pohledavek
                                    arrText(i) = arrText(i) & textace("radek") & textace("-")
                                End If
                                
                                
                                '''  vytvor text kolik slo na jistinu a na urok uveru
                                
                                If .castka_jistina > 0 Then
                                
                                    arrText(i) = arrText(i) & _
                                                 textace("II_cast+") & VlozCastku(.castka_jistina) & textace("na_jistinu")
                                    blnJistinaVlozena = True
                                    
                                ElseIf .castka_urok > 0 Then
                                    
                                    ' JENOM urok
                                    arrText(i) = arrText(i) & _
                                                 textace("II_cast") & VlozCastku(.castka_urok) & textace("na_kap_urok")
                                                 
                                End If
                                
                                
                                If blnJistinaVlozena And .castka_urok > 0 Then
                                    arrText(i) = arrText(i) & _
                                                 textace("II_cast_pokr") & VlozCastku(.castka_urok) & textace("na_kap_urok")
                                End If
                                
                                
                                ''' SOP a debet
                                
                                If .castka_sop > 0 Then
                                    dblSoucetSopu = dblSoucetSopu + .castka_sop
                                    blnAktualni_radek_sop_nebo_debet = True
                                End If
                                
                                If .castka_debet > 0 Then
                                    dblSoucetDebetu = dblSoucetDebetu + .castka_debet
                                    blnAktualni_radek_sop_nebo_debet = True
                                End If
                                
                            End If
                        
                            '''   text (EPR nebo rozsudek) na konec radku
                            If blnCela_Castka_Sop Then
                                arrText(i) = arrText(i) & OdkazNaPoradiPohledavky(sd, "SOP")
                            ElseIf blnCela_Castka_Debet Then
                                arrText(i) = arrText(i) & OdkazNaPoradiPohledavky(sd, "DEBET")
                            ElseIf blnAktualni_radek_sop_nebo_debet = False Then
                                arrText(i) = arrText(i) & OdkazNaPoradiPohledavky(sd, Pohl_Zaplacene(j).nazevUveru)
                            End If
                            
                        End If
                    End With
                    
                Next j
                
                ' Debet a naklady rizeni (SOP) az po uverech
                If dblSoucetDebetu > 0 Then
                    arrText(i) = arrText(i) & textace("radek") & textace("-") & _
                                 textace("II_cast") & VlozCastku(dblSoucetDebetu) & textace("na_debet") & _
                                 OdkazNaPoradiPohledavky(sd, "DEBET")
                End If
                
                If dblSoucetSopu > 0 Then
                    arrText(i) = arrText(i) & textace("radek") & textace("-") & _
                                 textace("II_cast") & VlozCastku(dblSoucetSopu) & textace("na_naklady_NR")
                End If
                
                
            Next i
        
        
        ElseIf format_scitania = Scit_Soucet_jistin Then
            
            
            ' textace("II_pohl_urok"), II_pohl_urok_rocne
            For i = 1 To sd.PocetPlatebCast_II
            
                For j = 1 To UBound(Pohl_Zaplacene)
                    
                    With Pohl_Zaplacene(j)
                    
                        If .datum_zaplaceni = sd.datumyZaplacenePohledavky(i) Then
                            
                            blnCela_Castka_Sop = False
                            blnCela_Castka_Debet = False
                            blnAktualni_radek_sop_nebo_debet = False
                            blnJistinaVlozena = False
                            
                            '''   Pripady, kdy cela platba sla na jedinu pohledavku (jistina, urok, naklady rizeni nebo debet)
                            
                            If .castka_jistina = sd.castkyZaplacenePohledavky(i) Then

                                    arrText(i) = arrText(i) & textace("na_jistinu")
                                                 
                            ElseIf .castka_urok = sd.castkyZaplacenePohledavky(i) Then
                            
                                    arrText(i) = arrText(i) & textace("na_kap_urok")
                                                 
                            ElseIf .castka_sop = sd.castkyZaplacenePohledavky(i) Then
                            
                                    arrText(i) = arrText(i) & textace("na_naklady_NR")
                                    blnCela_Castka_Sop = True
                                                 
                            ElseIf .castka_debet = sd.castkyZaplacenePohledavky(i) Then
                            
                                    arrText(i) = arrText(i) & textace("na_debet")
                                    blnCela_Castka_Debet = True
                                             
                                             
                            Else  '''   Else: platba byla rozdelena na nekolik ucelu (probehne vicekrat blok kodu)
                            
                            
                                If Not ((.castka_jistina + .castka_urok + .castka_sop + .castka_debet) = sd.castkyZaplacenePohledavky(i)) Then
                                    ' platba sla na vice pohledavek
                                    arrText(i) = arrText(i) & textace("radek") & textace("-")
                                End If
                                
                                
                                '''  vytvor text kolik slo na jistinu a na urok uveru
                                
                                
                                If .castka_jistina > 0 Then
                                
                                    arrText(i) = arrText(i) & _
                                                 textace("II_cast") & VlozCastku(.castka_jistina) & textace("na_jistinu")
                                    blnJistinaVlozena = True
                                    
                                ElseIf .castka_urok > 0 Then
                                    
                                    ' JENOM urok
                                    arrText(i) = arrText(i) & _
                                                 textace("II_cast") & VlozCastku(.castka_urok) & textace("na_kap_urok")
                                                 
                                End If
                                
                                
                                If blnJistinaVlozena And .castka_urok > 0 Then
                                    arrText(i) = arrText(i) & _
                                                 textace("II_cast_pokr") & VlozCastku(.castka_urok) & textace("na_kap_urok")
                                End If
                                
                                ''' SOP a debet
                                
                                If .castka_sop > 0 Then
                                    dblSoucetSopu = dblSoucetSopu + .castka_sop
                                    blnAktualni_radek_sop_nebo_debet = True
                                End If
                                
                                If .castka_debet > 0 Then
                                    dblSoucetDebetu = dblSoucetDebetu + .castka_debet
                                    blnAktualni_radek_sop_nebo_debet = True
                                End If
                                
                                
                            End If
                            
                            '''   text na konec radku
                            ' u tehle varianty se vklada jenom vyska uroku daneho uveru (jinak nelze z rozsudku rozlisit)
                            If (blnCela_Castka_Sop = False) And (blnCela_Castka_Debet = False) And (blnAktualni_radek_sop_nebo_debet = False) Then
                                arrText(i) = arrText(i) & OdkazNaPoradiPohledavky(sd, Pohl_Zaplacene(j).nazevUveru, blnPouziUrok:=True)
                            End If
                            
                        End If
                    End With
                    
                Next j
                
                ' Debet a naklady rizeni (SOP) az po uverech
                If dblSoucetDebetu > 0 Then
                    arrText(i) = arrText(i) & textace("radek") & textace("-") & _
                                 textace("II_cast") & VlozCastku(dblSoucetDebetu) & textace("na_debet") & _
                                 OdkazNaPoradiPohledavky(sd, "DEBET")
                End If
                
                If dblSoucetSopu > 0 Then
                    arrText(i) = arrText(i) & textace("radek") & textace("-") & _
                                 textace("II_cast") & VlozCastku(dblSoucetSopu) & textace("na_naklady_NR")
                End If
                
            Next i
        
        
        
        ElseIf format_scitania = Scit_Soucet_jistin_a_KU Then
        
              
            
            
            ' textace("II_pohl_urok"), II_pohl_urok_rocne
            For i = 1 To sd.PocetPlatebCast_II
            
                For j = 1 To UBound(Pohl_Zaplacene)
                    
                    With Pohl_Zaplacene(j)
                    
                        If .datum_zaplaceni = sd.datumyZaplacenePohledavky(i) Then
                            
                            blnCela_Castka_Sop = False
                            blnCela_Castka_Debet = False
                            blnAktualni_radek_sop_nebo_debet = False
                            blnJistinaVlozena = False
                            blnAktualni_radek_sel_do_souctu = False
                            
                            '''   Pripady, kdy cela platba sla na jedinu pohledavku (jistina, urok, naklady rizeni nebo debet)
                            
                            If .castka_jistina = sd.castkyZaplacenePohledavky(i) Then

                                    arrText(i) = arrText(i) & textace("na_jistinu")
                                                 
                            ElseIf .castka_urok = sd.castkyZaplacenePohledavky(i) Then
                            
                                    arrText(i) = arrText(i) & textace("na_kap_urok")
                                                 
                            ElseIf .castka_sop = sd.castkyZaplacenePohledavky(i) Then
                            
                                    arrText(i) = arrText(i) & textace("na_naklady_NR")
                                    blnCela_Castka_Sop = True
                                                 
                            ElseIf .castka_debet = sd.castkyZaplacenePohledavky(i) Then
                            
                                    arrText(i) = arrText(i) & textace("na_debet")
                                    blnCela_Castka_Debet = True
                                             
                                             
                                             
                            Else  '''   Else: platba byla rozdelena na nekolik ucelu (probehne vicekrat blok kodu)
                            
                            
                                blnJistinaVlozena = False
                                If Not ((.castka_jistina + .castka_urok + .castka_sop + .castka_debet) = sd.castkyZaplacenePohledavky(i)) Then
                                    ' platba sla na vice pohledavek
                                    arrText(i) = arrText(i) & textace("radek") & textace("-")
                                End If
                                
                                
                                '''  vytvor text kolik slo na jistinu a na urok uveru
                                
                                
                                If .castka_jistina > 0 Then
                                
                                    arrText(i) = arrText(i) & _
                                                 textace("II_cast") & VlozCastku(.castka_jistina) & textace("na_jistinu")
                                    blnJistinaVlozena = True
                                    
                                ElseIf .castka_urok > 0 Then
                                    
                                    ' JENOM urok
'                                    arrText(i) = arrText(i) & _
'                                                 textace("II_cast") & VlozCastku(.castka_urok) & textace("na_kap_urok")
                                    dblCastkaSoucetKU = dblCastkaSoucetKU + .castka_urok
                                    blnAktualni_radek_sel_do_souctu = True
                                                 
                                End If
                                
                                
                                If blnJistinaVlozena And .castka_urok > 0 Then
'                                    arrText(i) = arrText(i) & _
'                                                 textace("II_cast_pokr") & VlozCastku(.castka_urok) & textace("na_kap_urok")
                                    dblCastkaSoucetKU = dblCastkaSoucetKU + .castka_urok
                                    blnAktualni_radek_sel_do_souctu = True
                                End If
                                
                                ''' SOP a debet
                                
                                If .castka_sop > 0 Then
                                    dblSoucetSopu = dblSoucetSopu + .castka_sop
                                    blnAktualni_radek_sop_nebo_debet = True
                                End If
                                
                                If .castka_debet > 0 Then
                                    dblSoucetDebetu = dblSoucetDebetu + .castka_debet
                                    blnAktualni_radek_sop_nebo_debet = True
                                End If
                                
                                
                            End If
                            
                            '''   text na konec radku
                            ' u tehle varianty se vklada jenom vyska uroku daneho uveru (jinak nelze z rozsudku rozlisit)
                            If (blnCela_Castka_Sop = False) And (blnCela_Castka_Debet = False) And (blnAktualni_radek_sop_nebo_debet = False) _
                                And (blnAktualni_radek_sel_do_souctu = False) Then
                                arrText(i) = arrText(i) & OdkazNaPoradiPohledavky(sd, Pohl_Zaplacene(j).nazevUveru, blnPouziUrok:=True)
                            End If
                            
                        End If
                    End With
                    
                Next j
                
                
                ' Soucet kapitalizovanych uroku
                If dblCastkaSoucetKU > 0 Then
                    arrText(i) = arrText(i) & _
                                 textace("radek") & textace("-") & _
                                 textace("II_cast") & VlozCastku(dblCastkaSoucetKU) & textace("na_kap_urok")
                End If
                dblCastkaSoucetKU = 0
                
                
                ' Debet a naklady rizeni (SOP) az po uverech
                If dblSoucetDebetu > 0 Then
                    arrText(i) = arrText(i) & textace("radek") & textace("-") & _
                                 textace("II_cast") & VlozCastku(dblSoucetDebetu) & textace("na_debet") & _
                                 OdkazNaPoradiPohledavky(sd, "DEBET")
                End If
                
                If dblSoucetSopu > 0 Then
                    arrText(i) = arrText(i) & textace("radek") & textace("-") & _
                                 textace("II_cast") & VlozCastku(dblSoucetSopu) & textace("na_naklady_NR")
                End If
                
            Next i
        
                                
        
        
        ElseIf format_scitania = Scit_Soucet_jistin_a_KU_a_debet Then
        
            
                ''' TODO XX
        
        
        Else
            
            
            
        End If
    
    
    
        '''   Spojime riadky do jedineho stringu
        For i = LBound(arrText) To UBound(arrText)
            VytvorTextCast_II = VytvorTextCast_II & textace("radek") & textace("radek") & arrText(i)
        Next i
    
    End If
    
    If Len(VytvorTextCast_II) > 1 Then VytvorTextCast_II = VytvorTextCast_II & "."
    
    '''   Pokud ma klient nejaky pausal, tak ten NEvymahame a do cl.II se vlozi veta navic
    
    
    If c.SOP.Pausal > 0 Then
        VytvorTextCast_II = VytvorTextCast_II & textace("radek") & textace("radek") & textace("II_pouze_cast") '& _
                            '" (XXXXXX Jezevèík tady mùže uvést i èástku: " & c.SOP.Pausal & ".)" '& textace("radek")
    End If
    
    
    
End Function








Private Function VytvorTextCast_V_uvod(ByRef sd As data_pro_sablonu) As String
    ' vystupem funkce je textovy retezec, ktery je na zacatku Petitu, teda clanku V exekucniho navrhu
    ' text je kombinace fixniho textu, info o klientovi ktere mame v datech pro sablonu, a pak jemna variaci toho co bylo v uvodu casti I
    ' TODO extrahovat hodnoty do NaplnTextace() a odvolavat se pres promennou Textace()
    
    Dim text As String
    Dim i As Integer
    Dim sql11 As String
    Dim rst1 As ADODB.Recordset
    Dim typ_rozsudku As String
    
    
    text = textace("radek")
    
    For i = 1 To sd.pocetZalob 'pro kazde epr doplnime info z rozsudku do hlavicky
    
        If i = 1 Then
            text = text & "Podle "
        Else
            text = text & "Dále podle "
        End If
        
        text = text & textace("Povereni_uvod") & sd.klient_udaje
        
        sql11 = "select soud, epr, datum_vydani, datum_pm, datum_vykon, typ from ap_coll.vs_import_pm_rozsudek_all where epr = '" & sd.EPR(i) & "'"
        Set rst1 = New ADODB.Recordset
        
        rst1.Open sql11, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
            If Not rst1.EOF Then
                
                
                ' RISK-1096
                If sd.EPR(i) = c.ExeTitul_spisova_znacka Then
                    
                    text = text & ", dle " & get_enum_exekucni_titul_Name(c.Exekucni_titul, blnGenitiv:=True) & _
                                  " è.j. " & c.ExeTitul_spisova_znacka & ", který vydal " & rst1(0) & " dne " & rst1(2) & _
                                  ", který nabyl právní moci dne " & rst1(3) & _
                                  " a je vykonatelný"
                    
                    If (rst1![datum_vykon] <> rst1![datum_pm]) _
                        And ((c.Exekucni_titul = exe_titul_platebni_rozkaz) _
                                Or (c.Exekucni_titul = exe_titul_rozsudek) _
                                Or (c.Exekucni_titul = exe_titul_usneseni_C) _
                            ) Then
                        
                        text = text & " dnem " & rst1![datum_vykon]
                        
                    End If
                    
                    
                Else
                    text = text & ", dle "
                    typ_rozsudku = "" & rst1![typ] & ""
                    
                    If typ_rozsudku = "rozsudek" Then
                        text = text & "rozsudku"
                        
                    ElseIf typ_rozsudku = "elektronický platební rozkaz" Then
                        text = text & "elektronického platebního rozkazu"
                        
                    ''  nasledujuce 2 su iba kvoli preklepom - napr. klient cuid 619818 mal "elekronický platební rozkaz" (chyba tam T)
                    
                    ElseIf InStr(typ_rozsudku, "rozs") > 0 Then
                        text = text & "rozsudku"
                        
                    ElseIf InStr(typ_rozsudku, "elek") > 0 Then
                        text = text & "elektronického platebního rozkazu"
                        
                    Else
                        HlaskaUzivateli "V petitu chybí èást textu k EPR/rozsudku. Nejspíš pøeklep v datech."
                        
                    End If
                    
                    
                    text = text & " è.j. " & rst1(1) & ", který vydal " & rst1(0) & " dne " & rst1(2) & _
                                  ", který nabyl právní moci dne " & rst1(3) & " a" & Chr(160) & "je vykonatelný"
                    
                    
                    If (rst1![datum_vykon] <> rst1![datum_pm]) _
                        And typ_rozsudku = "rozsudek" Then
                        
                        text = text & " dnem " & rst1![datum_vykon]
                        
                    End If
                    
                End If
                
                text = text & ", uhradit oprávnìnému:" & textace("radek")
                
            End If
        rst1.Close
    
    Next i
    
    VytvorTextCast_V_uvod = text

End Function






Private Sub UlozDokument(ByRef WordDocument As Variant)
    
    Dim AFolder As String
    Dim BFolder As String
    Dim fileName As String
    Dim fso As Object
    
    
    fileName = "Exekuèní návrh klienta " & c.name & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".docx"
    AFolder = cfg.Cesta("ExePodane") & c.Nazev_slozky
    BFolder = cfg.Cesta("ExePodane_BezSlozky")                                      ' "W:\Vymahani\LATE\EXEKUCE\PODANÉ\Klienti bez složky\"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If schema = "APP_COLL_JEZEVCIK_TEST" Then
        WordDocument.SaveAs (cfg.Cesta("TestovaciSlozka") & fileName)
        'Call Shell("explorer.exe " & cfg.Cesta("TestovaciSlozka"), vbNormalFocus) 'otvirani slozky na testu
    Else
        If fso.FolderExists(AFolder) Then
            WordDocument.SaveAs (AFolder & fileName)
            Call Shell("explorer.exe " & AFolder, vbNormalFocus)
        ElseIf fso.FolderExists(BFolder) Then
            MsgBox "Složka pro klienta " & c.name & " neexistuje." ', vbInformation, "Složka nenalezena!"
            WordDocument.SaveAs (BFolder & fileName)
            Call Shell("explorer.exe " & BFolder, vbNormalFocus)
        Else
            MsgBox "Dokument se neuložil, neexistuje cesta " & BFolder, "Dokument neuložen", vbInformation
            'WordDocument.SaveAs (BFolder & fileName)
            'Call Shell("explorer.exe " & BFolder, vbNormalFocus)
        End If
    
    End If
    


End Sub


