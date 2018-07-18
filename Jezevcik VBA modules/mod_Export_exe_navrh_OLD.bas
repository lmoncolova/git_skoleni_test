Attribute VB_Name = "mod_Export_exe_navrh_OLD"
'Option Explicit
'
'' FA: 12.1.2018 modul cely upraven v ramci RISK-977  (predtim RISK-398 )
'' do References som pridal Microsoft Word, ale nefunguje to medzi Office balikmi 2010 a 2013 - preto je potrebne pouzivat tzv. Late Binding cez metodu CreateObject()
'
'
'' ak mame EPR, tak to znamena ze nejaky uradnik dal peciatku na nas dokument
'' ak mame rozsudok, tak to je novy dokument od sudcu, v ktorom si jednotlive ciastky preusporiada a poscita tak ako sa jemu paci
'' v pripade rozsudku sa spravi vyskakovacie okno na uzivatela, kde si vyberie jednu z moznosti z enum ScitanieCiastok
'
'' data sa v povodnej sablone vkladali iba ako textaky do wordu
'' po novom chceme strukturu ktora umozni ich presuvanie a agregaciu takym sposobom, ze pridanie noveho formatu do ScitanieCiastok by nemal byt problem
'
'
'Private Enum ScitanieCiastok ' ak mame rozsudok, tento typ zadefinuje akym sposobom sa vypisu a scitaju ciastky
'    Scit_Soucet_jistin
'    Scit_Soucet_jistin_a_KU
'    Scit_Soucet_jistin_a_KU_a_debet
'End Enum
'
'
'Private EPR As Variant ' FA bug - tejto premennej vobec nerozumiem ako a kde ma byt zadefinovana
'Private klient_udaje As String
'Private ii As Integer, qq As Integer, qi As Integer, ij As Integer, iii As Integer, q As Integer, qe As Integer
'Private lngPocetZalob As Long
'
'
'
'Private Sub NaplnPocetZalob()
''''   naplneni zalob do promenne EPR() 'RISK-398
'Dim sql4 As String
'Dim rst4 As ADODB.Recordset
'
'    sql4 = "select distinct epr from ap_coll.vs_import_pm_all pm join " & schema & ".person p on p.cuid = pm.cuid and p.flag_valid = 'Y' where p.id_person = " & id_person
'    Set rst4 = New ADODB.Recordset
'
'    rst4.Open sql4, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'        If Not rst4.EOF Then
'
'            lngPocetZalob = rst4.RecordCount 'pocet zalob
'            ReDim EPR(lngPocetZalob) 'zmena dimenze podle poctu epr
'            rst4.MoveFirst
'
'            For i = 1 To lngPocetZalob
'                EPR(i) = rst4(0)
'                rst4.MoveNext
'            Next i
'
'        End If
'    rst4.Close
'
'End Sub
'
'
'
'
'''''''     MAIN BODY
'
'
'
'
'Sub export_exe_navrh()
''jdou zde exportovat vsechny epr z evidence, pokud maji KO ve vypoctu nebo chybu v evidenci, vyhodi se upozorneni
'
'Dim namech As String, address As String
''Dim wordapp As Word.Application
''Dim objDoc As Word.Document
'Dim wordapp As Object
'Dim objDoc As Object
'Dim sqlexe As String
'
'On Error GoTo connection_problem
'
'' vypne cely export pokud neni nalezeno PM rozhodnuti
'If Not Mame_PM_rozhodnuti(CStr(c.CUID)) Then Exit Sub
'strMsgbox = ""
'
'
''''  kontrola a naplneni udaju o klientovi
'
'namech = c.Name_first & " " & c.Name_last
'If IsNull(c.TextContractNumber) Then HlaskaUzivateli "Klient nemá vyplnìné èíslo rámcové smlouvy"
'If IsNull(c.Address_P) Then
'    HlaskaUzivateli "Klient nemá vyplnìnou adresu"
'Else
'    address = c.Address_P_Ulice & " " & c.Address_P_CP & ", " & Left(c.Address_P_PSC, 3) & " " & Right(c.Address_P_PSC, 2) & " " & c.Address_P_Mesto
'End If
'klient_udaje = c.name & ", nar. " & c.Date_Birth & ", trvale bytem " & address ' vyuziva se ve funkcich co generuji text pro jednotlive casti
'
'
'
''''   otevreni DOCu
'
'Set wordapp = CreateObject("word.Application")
''Set objDoc = wordapp.Documents.Add(Template:="W:\Vymahani\LATE\Vzory\Exekuèní návrh šablona jezevèík_201702_02.dotx")  'predesla verze
''Set objDoc = wordapp.Documents.Add(Template:="W:\Vymahani\LATE\Vzory\Exekuèní návrh šablona jezevèík_2018_01_12.dotx")
'Set objDoc = wordapp.Documents.Add(Template:="W:\Vymahani\LATE\Vzory\Exekuèní návrh šablona jezevèík_2018_02_22.dotx")
''wordapp.Visible = True
'
'
''''   vlozeni udaju o klientovi
'
'Call NaplnBookmark("RamcovaSmlouva", c.TextContractNumber, objDoc)
'Call NaplnBookmark("RamcovaSmlouva2", c.TextContractNumber, objDoc)
'Call NaplnBookmark("Povinny_NameFull", c.name, objDoc, blnBold:=True)
'Call NaplnBookmark("Povinny_narozen_a_adresa", "nar. " & c.Date_Birth & ", trvale bytem " & address, objDoc)
'
'sqlexe = "select text_bank_account_number from " & schema & ".exe_account where id_person = " & id_person
'Set rst1 = New ADODB.Recordset
'rst1.Open sqlexe, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'    If Not rst1.EOF Then
'        Call NaplnBookmark("Povereni_cislo_uctu", "" & rst1(0) & "/3030", objDoc, blnBold:=True)
'    End If
'rst1.Close
'
'Call NaplnPocetZalob
'
''--------------------------------DOPLNENI INFO Z PM ROZSUDKU do sekce Exekucni navrh (I.)----------------------------------------
'' FA tu som zmazal vela komentaru
'
''RISK-977
'Call NaplnBookmark("CastI", GenerujTextCast_I_OLD(), objDoc) 'XXX
''Call NaplnBookmark("CastI", GenerujTextCast_I(), objDoc) 'XXX
'
'
''--------------------------------DOPLNENI INFO do casti II.----------------------------------------
'
'Call NaplnBookmark("CastII", GenerujTextCast_II(), objDoc)
'
'
'
''--------------------------------aktualni info do sekce Povereni (V.)----------------------------------------
'
'Call NaplnBookmark("Povereni", GenerujTextPovereni_V(), objDoc)
'
'
'
''''   ULOZENI SABLONY
'
''Sub komentar_pred_ulozenim '''asi 250 riadkov zakomentovanych, presunute na koniec tohto modulu
'Call PridajOdrazky(objDoc)
'Call VymazBookmarks(objDoc)
'Call UlozDokument(objDoc)
'wordapp.Visible = True
'
'Set wordapp = Nothing
'Set objDoc = Nothing
'Application.EnableEvents = True
'Application.ScreenUpdating = True
'Application.DisplayAlerts = True
'
'Call VypisVsechnyHlasky
'Exit Sub
'
'
'connection_problem:
'
'    ' V pøípadì, že se jedná o chybu z dùvodu nepøipojení k databázi, bude hodnota promìnné err_connection true a posléze bude proveden reconnect/relogin
'    If Mid(Err.Description, 1, 9) = "ORA-03114" Then
'        err_connection = True
'    Else
'        MsgBox Err.Description
'    End If
'
'
'End Sub
'
'
'
'
'
'
'
'
'
'Private Function Mame_PM_rozhodnuti(ByVal CUID As String) As Boolean
'' funkce zkontroluje jestli mame PM rozhodnuti k danemu klientovi
'' pokud PM rozhodnuti nemame, fce vrati FALSE !!! (v hlavni fci se ukonci export)
'' pouziva se fce HlaskaUzivateli
'
'Dim sql0 As String
'Dim sql000 As String
'
'Mame_PM_rozhodnuti = True
'
'sql0 = "select * from ap_coll.pb_040_summary where cuid = " & CUID 'tabulka se plni na hristi ap_coll
'
'Set rst1 = New ADODB.Recordset 'pokud neni epr evidovane pro daneho klienta, vypneme makro
'rst1.Open sql0, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'    If rst1.EOF Then
'        MsgBox "Klient nemá evidované pravomocné rozhodnutí", vbInformation
'        Mame_PM_rozhodnuti = False
'        Exit Function
'    Else
'        rst1.MoveFirst
'        For i = 1 To rst1.RecordCount 'kontrola, zda je vypocet v poradku
'            If rst1(7) <> "OK" Then
'                HlaskaUzivateli "Pozor, podle automatické kontroly je ve výpoètu PM pohledávky chyba nebo je chybnì evidovaná. Problém: " & rst1(7)
'                Exit For
'            End If
'            rst1.MoveNext
'
'        Next i
'    End If
'rst1.Close
'
'
'If Mame_PM_rozhodnuti Then
'    sql000 = "select distinct epr from ap_coll.vs_import_pm_all where cuid = " & CUID 'tabulka se plni na ap_coll
'    Set rst1 = New ADODB.Recordset
'
'    rst1.Open sql000, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'       If rst1.EOF Then
'            MsgBox "Klient nemá evidované PM rozhodutí"
'            Mame_PM_rozhodnuti = False
'            Exit Function
'       End If
'    rst1.Close
'End If
'
'End Function
'
'
'
'
'
'Private Function GenerujTextPovereni_V() As String
'' vystupem funkce je textovy retezec (ktery se vlozi do Word Dokumentu do casti V - Povereni)
'' v ramci funkce je hlavnim textakem promenna text5
'
'
'Dim cnt_loan As Integer, cnt_loan_akt As Integer, sql12 As String, amt_principal_akt As Double, amt_interest_akt As Double, loan_akt() As String
'Dim sql13 As String, sql11 As String, sql14 As String, debet_akt As Double, sop_akt As Double, sql15 As String, sql16 As String, sql17 As String, sql18 As String, sql100 As String
'Dim cnt_sankce As Integer, cnt_urok As Integer, sankce_akt As String, urok_akt As String, ainc As Integer, ainc2 As Integer, ainc_h As Integer
'Dim sql00 As String, comp As Boolean, text5 As String, text5su As String, text5ou As String, text5ku As String
'Dim typ_roz As String
'Dim qq As Integer, ii As Integer
'
'
'
'    For i = 1 To lngPocetZalob 'pro kazde epr doplnime info z rozsudku do hlavicky
'
'        If i = 1 Then
'            text5 = text5 & "Podle "
'        Else
'            text5 = text5 & "Dále podle "
'        End If
'
'        text5 = text5 & "§ 43a odst. 3 zákona è. 120/2001 Sb., o soudních exekutorech a exekuèní èinnosti (exekuèní øád) " & _
'                        "se provedením exekuce na návrh oprávnìného Air Bank a.s., IÈ 29045371, se sídlem Evropská 2690/17, 160 00 Praha 6, proti povinnému " & _
'                        klient_udaje
'
'
'        sql11 = "select soud, epr, datum_vydani, datum_pm, datum_vykon, typ from ap_coll.vs_import_pm_rozsudek_all where epr = '" & EPR(i) & "'" 'tabulka se plni na ap_coll hristi
'        Set rst1 = New ADODB.Recordset
'
'        rst1.Open sql11, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'            If Not rst1.EOF Then
'
'                typ_roz = "" & rst1(5) & ""
'
'                If typ_roz = "rozsudek" Then
'                    typ_roz = "rozsudku"
'                    text5 = text5 & ", dle " & typ_roz & " è.j. " & rst1(1) & ", který vydal " & rst1(0) & " dne " & rst1(2) & ", který nabyl právní moci dne " & rst1(3) & " a" & Chr(160) & "je vykonatelný dnem " & rst1(4) & ", uhradit oprávnìnému:" & Chr(10)
'
'                ElseIf typ_roz = "elektronický platební rozkaz" Then
'                    typ_roz = "elektronického platebního rozkazu"
'                    text5 = text5 & ", dle " & typ_roz & " è.j. " & rst1(1) & ", který vydal " & rst1(0) & " dne " & rst1(2) & ", který nabyl právní moci dne " & rst1(3) & " a" & Chr(160) & "je vykonatelný, uhradit oprávnìnému:" & Chr(10)
'
'                ''  nasledujuce 2 su iba kvoli preklepom - napr. klient cuid 619818 mal "elekronický platební rozkaz" (chyba tam T)
'                ''  pridane ako ElseIf na koniec, pretoze ak recordset nie je prazdny tak by tam mala byt jedna z prvych dvoch moznosti.. tieto sa testuju iba ak dojde k failu zjavnemu
'                ElseIf InStr(typ_roz, "rozs") Then
'                    typ_roz = "rozsudku"
'                    text5 = text5 & ", dle " & typ_roz & " è.j. " & rst1(1) & ", který vydal " & rst1(0) & " dne " & rst1(2) & ", který nabyl právní moci dne " & rst1(3) & " a" & Chr(160) & "je vykonatelný dnem " & rst1(4) & ", uhradit oprávnìnému:" & Chr(10)
'
'                ElseIf InStr(typ_roz, "elek") Then
'                    typ_roz = "elektronického platebního rozkazu"
'                    text5 = text5 & ", dle " & typ_roz & " è.j. " & rst1(1) & ", který vydal " & rst1(0) & " dne " & rst1(2) & ", který nabyl právní moci dne " & rst1(3) & " a" & Chr(160) & "je vykonatelný, uhradit oprávnìnému:" & Chr(10)
'
'                Else
'                    HlaskaUzivateli "V petitu chybí èást textu k EPR/rozsudku. Nejspíš pøeklep v datech."
'                End If
'
'            End If
'        rst1.Close
'
'
'
'
'        sql12 = "select distinct typ_id from ap_coll.pb_040_pohledavka_unpaid where typ like 'HU%' and epr = '" & EPR(i) & "' order by typ_id" 'pro kazde epr zjistime pocet hu
'
'        Set rst1 = New ADODB.Recordset
'
'        rst1.Open sql12, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'        If Not rst1.EOF Then
'            cnt_loan_akt = rst1.RecordCount
'            ReDim loan_akt(cnt_loan_akt)
'
'            rst1.MoveFirst
'            For qq = 1 To cnt_loan_akt
'                loan_akt(qq) = rst1(0)
'                rst1.MoveNext
'            Next qq
'
'        End If
'        rst1.Close
'
'    '    'pokud nema nedoplacenou zadnou cast pohledavky, tuto cast kodu preskocime
'    '    If IsNull(cnt_loan_akt) Then
'    '        MsgBox "Pohledávka je doplacená"
'    '        GoTo skip_pm_akt
'    '    End If
'
'        If (IsNull(cnt_loan_akt) Or cnt_loan_akt = 0) And InStr(strMsgbox, "má doplacenou jistinu a kap. úrok") = 0 Then
'            ' posledna podmienka je aby sa hlaska nevypisovala duplicitne
'            HlaskaUzivateli "Klient má doplacenou jistinu a kap. úrok u EPR: " & EPR(1) & _
'            ". Je potøeba provést kontrolu, zda nemá nedoplacené obchodní nebo sankèní úroky. "
'        End If
'
'
'        '---------------------hu
'        For ii = 1 To cnt_loan_akt
'
'            'musime zjistit, zda psat jistinu a kap. urok dohormady. Pravidlo je, ze pokud obch a sankcni urok se uroci ze stejne castky, pise se to zvlast, jinak dohromady (stare zaloby)
'            sql100 = "select comp from " & schema & ".v_exp_comp where hu = '" & loan_akt(ii) & "'"
'
'            Set rst1 = New ADODB.Recordset
'            rst1.Open sql100, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'                If Not rst1.EOF Then
'                    comp = rst1(0) 'pro dane hu zjistime, zda psat dohromady castky, pokud true, piseme dohromady
'                End If
'            rst1.Close
'
'
'            sql13 = "select sum(amt) from ap_coll.pb_040_pohledavka_unpaid where typ_id = '" & loan_akt(ii) & "' and typ like 'HU - Úrok%'"
'
'            Set rst1 = New ADODB.Recordset
'            rst1.Open sql13, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'                If Not rst1.EOF Then
'                    If IsNull(rst1(0)) Then
'                        amt_interest_akt = 0
'                    Else
'                        amt_interest_akt = rst1(0) 'pro dane epr pocitame vysi uroku
'                    End If
'                End If
'            rst1.Close
'
'            Set rst1 = New ADODB.Recordset
'            sql14 = "select sum(amt) from ap_coll.pb_040_pohledavka_unpaid where typ_id = '" & loan_akt(ii) & "' and typ like 'HU - Jis%'"
'
'            rst1.Open sql14, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'                If Not rst1.EOF Then
'                    If IsNull(rst1(0)) Then
'                        amt_principal_akt = 0 'pro dane epr pocitame vysi jistiny
'                    Else
'                        amt_principal_akt = rst1(0) 'pro dane epr pocitame vysi uroku
'                    End If
'                End If
'            rst1.Close
'
'
'            text5 = text5 & Chr(10)
'
'
'            If comp Then
'                text5 = text5 & vbCrLf & " -   èástku ve výši " & format$(amt_principal_akt + amt_interest_akt, "#,##0.00") & " Kè, " & Chr(10)
'            Else
'                text5 = text5 & vbCrLf & " -   èástku ve výši " & format$((amt_principal_akt), "#,##0.00") & " Kè, " & Chr(10)
'                text5ku = " -   s kapitalizovaným úrokem ve výši " & format$(amt_interest_akt, "#,##0.00") & " Kè, " '& Chr(10)
'            End If
'
'            'sankcni a obchodni urok
'            sql17 = "select * from ap_coll.pb_040_intervaly where code_loan_number = '" & loan_akt(ii) & "'  order by  dtime_valid_from, typ asc" 'tabulka zustava na ap_coll
'            Set rst2 = New ADODB.Recordset
'            rst2.Open sql17, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'                If Not rst2.EOF Then
'                    cnt_sankce = rst2.RecordCount
'                    rst2.MoveFirst
'
'                    For qi = 1 To cnt_sankce
'
'                        If qi < cnt_sankce - 1 Then
'                            If rst2(5) = "ourok" Then
'                                text5 = text5 & " -   spolu s úrokem ve výši " & rst2(4) * 100 & " % roènì z èástky " & format$(rst2(1), "#,##0.00") & " Kè od " & rst2(2) & " do " & rst2(3) - 1 & "," & Chr(10)
'                            Else
'                                text5 = text5 & " -   se zákonným úrokem z prodlení ve výši " & rst2(4) * 100 & " % roènì z èástky " & format$(rst2(1), "#,##0.00") & " Kè od " & rst2(2) & " do " & rst2(3) - 1 & "," & Chr(10)
'
'                            End If
'
'                        Else
'
'                            If rst2(5) = "ourok" Then
'                                text5 = text5 & " -   spolu s úrokem ve výši " & rst2(4) * 100 & " % roènì z èástky " & format$(rst2(1), "#,##0.00") & " Kè od " & rst2(2) & " do zaplacení," & Chr(10)
'                            Else
'                                text5 = text5 & " -   se zákonným úrokem z prodlení ve výši " & rst2(4) * 100 & " % roènì z èástky " & format(rst2(1), "#,##0.00") & " Kè od " & rst2(2) & " do zaplacení," & Chr(10)
'
'                            End If
'
'                         End If
'                        'sankce_akt = sankce_akt & rst2(2) - 1 & ", dále úrok ve výši " & rst2(4) * 100 & " % z èástky" & Format(rst2(1), "### ### ##0.00") & " od " & rst2(2) & " do "
'                        rst2.MoveNext
'
'                    Next qi
'                    'text5ou = text5ou & "zaplacení," & Chr(10)
'
'                End If
'            rst2.Close
'            text5 = text5 & text5ku
'
'        Next ii
'
'
'        '''   dale zjistime pro kazde epr debet a sop
'
'        'ainc2 = ainc
'        Set rst1 = New ADODB.Recordset
'        sql15 = "select sum(amt) from ap_coll.pb_040_pohledavka_unpaid where typ = 'DEBET' and epr = '" & EPR(i) & "'" 'pro kazde epr zjistime debet
'
'        rst1.Open sql15, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'            If Not rst1.EOF Then
'                If Not IsNull(rst1(0)) Then
'                    debet_akt = rst1(0)
'                Else
'                    debet_akt = 0
'                End If
'            End If
'        rst1.Close
'
'        Set rst1 = New ADODB.Recordset
'        sql16 = "select sum(amt) from ap_coll.pb_040_pohledavka_unpaid where typ = 'SOP' and epr = '" & EPR(i) & "'" 'pro kazde epr zjistime sop
'
'        rst1.Open sql16, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'            If Not rst1.EOF Then
'                If Not IsNull(rst1(0)) Then
'                    sop_akt = rst1(0)
'                Else
'                    sop_akt = 0
'                End If
'            End If
'        rst1.Close
'
'
'        If debet_akt > 0 Then
'            text5 = text5 & vbCrLf & vbCrLf & " -   èástku ve výši " & format$(debet_akt, "#,##0.00") & " Kè"
'        End If
'
'        If sop_akt > 0 Then
'
'            If debet_akt > 0 Then
'                text5 = text5 & ", "
'            End If
'            text5 = text5 & vbCrLf & vbCrLf & " -   náklady nalézacího øízení ve výši " & format$(sop_akt, "#,##0.00") & " Kè" & Chr(10)
'        End If
'
'        text5 = text5
'
'
'    Next i
'
'
'    '------------------------epr
'
'GenerujTextPovereni_V = text5
'
'End Function
'
'
'
'
'Private Function GenerujTextCast_II() As String
'
'Dim sql20 As String, platby As String, platby_rozpis() As String 'string je prilis dlouhy pro nahrazeni, proto udelame jako pole
'Dim cnt_platby As Integer, sql21 As String, amt_paid() As Double, date_p() As Date, platby_par() As String
'Dim text2 As String, flag_platby As Boolean, epr_all As String
'
'
''''   musime zjistit, zda neco uhradil na nekterem z epr
'
'epr_all = "'" & EPR(1) & "'"
'For i = 2 To lngPocetZalob
'    epr_all = epr_all & ", '" & EPR(i) & "'"
'Next i
'
'sql20 = "select date_payment, sum(amt_paid) as amt_paid from (select trunc(date_payment) date_payment, epr, amt_paid from ap_coll.pb_040_uhrady_pred_parovanim) where epr in (" & epr_all & ") group by date_payment order by date_payment asc"
'
'Set rst1 = New ADODB.Recordset
'rst1.Open sql20, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'
'    If Not rst1.EOF Then
'
'        flag_platby = True 'pokud neco uhradil
'        rst1.MoveFirst
'        ReDim platby_rozpis(rst1.RecordCount)
'        ReDim amt_paid(rst1.RecordCount)
'        ReDim date_p(rst1.RecordCount)
'        ReDim platby_par(rst1.RecordCount)
'        cnt_platby = rst1.RecordCount
'        text2 = "Povinný dobrovolnì uhradil do dne podání tohoto návrhu pouze èást povinnosti uložené pøedmìtným vykonatelným exekuèním titulem, a to:" & vbCrLf & vbCrLf
'        For i = 1 To rst1.RecordCount
'
'            text2 = text2 & " -   dne " & format(rst1![date_payment], "Short Date") & " èástku ve výši " & format$(rst1![amt_paid], "#,##0.00") & " Kè," & vbCrLf
'            'FA som vyhodil nejaky zbytocny if
'            amt_paid(i) = rst1![amt_paid]
'            date_p(i) = format(rst1![date_payment], "Short Date")
'
'            rst1.MoveNext
'        Next i
'        text2 = text2 & vbCrLf & "Uvedené èástky zapoèetl oprávnìný na úhradu dluhu následovnì:" & vbCrLf
'
'    Else
'
'        text2 = text2 & "Povinný dobrovolnì nesplnil do dne podání tohoto návrhu svou povinnost dle pøedmìtného vykonatelného exekuèního titulu, a to ani èásteènì." '& vbCrLf & vbCrLf
'
'    End If
'rst1.Close
'
'
'
'
'Dim cnt_pl_par As Integer
'
''ted zjistime, kam se naparovaly jednotlive platby
'If (flag_platby) Then
'    For i = 1 To cnt_platby 'pro kazdou platbu udelame zvlast dotaz
'
'        Set rst1 = New ADODB.Recordset
'        sql21 = "select * from " & schema & ".v_exp_exe_par where epr in (" & epr_all & ") and date_payment = to_date('" & format(date_p(i), "Short Date") & "') order by kategorie_por"
'        rst1.Open sql21, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'            If Not rst1.EOF Then
'
'                cnt_pl_par = rst1.RecordCount
'                rst1.MoveFirst
'
'                text2 = text2 & vbCrLf
'                text2 = text2 & "- èástku ve výši " & format$(amt_paid(i), "#,##0.00") & " Kè ze dne " & date_p(i) & " zapoèetl oprávnìný èástí ve výši "
'
'                For ii = 1 To cnt_pl_par
'
'                    text2 = text2 & format$(rst1![amt_paid], "#,##0.00") & " Kè na úhradu " & rst1![kategorie] & ", "
'                    rst1.MoveNext
'
'                Next ii
'
'                'text2 = text2 & vbCrLf
'
'            End If
'        rst1.Close
'    Next i
'End If
'
'GenerujTextCast_II = text2
'
'
''''   Pokud ma klient nejaky pausal, tak ten NEvymahame a do cl.II se vlozi veta navic
'
''dblPausal = PausalVyska(c.id_person)
'If c.SOP.Pausal > 0 Then
'    'GenerujTextCast_II = GenerujTextCast_II & vbCrLf & "Blah blah blah Paušál ve výši " & c.SOP.Pausal & " blah blah."
'    GenerujTextCast_II = GenerujTextCast_II & vbCrLf & vbCrLf & _
'                        "Oprávnìný se rozhodl vymáhat pouze èást povinnosti uložené pøedmìtným vykonatelným exekuèním titulem, pøièemž její výše je uvedená v èlánku V. tohoto návrhu." & _
'                        " (Mùžeme tady uvést i èástku: " & c.SOP.Pausal & ".)" '& vbCrLf
'
'End If
'
'
'
''''''      FA v povodnom kode sa vsetky dalsie hodnoty naplnili ale uz sa nikam nevlozili!
'
''''   musime nyni zjistit, zda vymahame mene (zda pridat treti variantu), to se ale poji pouze s 1. variantou
'Dim sql22 As String, flag_stejne As Boolean
'
'If Not flag_platby Then
'
'    sql22 = "select min(stejne) from " & schema & ".v_exp_bank_claim where epr in (" & epr_all & ") group by epr"
'
'    Set rst1 = New ADODB.Recordset
'    rst1.Open sql22, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'        If Not rst1.EOF Then
'
'            flag_stejne = rst1(0)
'
'        End If
'    rst1.Close
'End If
'
'
'If Not flag_stejne And Not flag_platby Then
'    text2 = text2 & Chr(10) & Chr(10) & "Oprávnìný se rozhodl vymáhat pouze èást povinnosti uložené pøedmìtným vykonatelným exekuèním titulem, pøièemž její výše je uvedená v èlánku V. tohoto návrhu."
'End If
'
'
'End Function
'
'
'
'Private Function GenerujTextCast_I() As String
'Dim sql0 As String, sql66 As String, sql7 As String, sql100 As String, sql1 As String, sql2 As String
'Dim sql_he As String, sql_h As String, Loan() As String, cnt_int As Integer, amt_principal As Double, amt_interest As Double
'Dim su_sazba() As Double, su_castka() As Double, ou_sazba() As Double, ou_castka() As Double
'Dim su_dtime_valid_to() As Date, ou_dtime_valid_to() As Date, su_dtime_valid_from() As Date, ou_dtime_valid_from() As Date
'Dim inc2 As Integer, inc3 As Integer, sql_re As String, debet As Double, SOP As Double
'Dim text1 As String, sankce() As String, ourok() As String
'Dim debet_text As String
'Dim sop_text As String
'Dim inc As Variant
'Dim sql9 As String
'Dim text1ku As String
'Dim cnt_loan As Integer
'Dim comp As Boolean
'
''-------------------------------------------------------------EPR---------------------
'For i = 1 To lngPocetZalob
''dynamicky definovane pole, ktere se meni podle poctu epr
'    amt_principal = 0 'pro kazde epr zacneme napocitavat jistinu a urok od nuly
'    amt_interest = 0
'
'    sql9 = "select soud, epr, datum_vydani, datum_pm, datum_vykon, typ from ap_coll.vs_import_pm_rozsudek_all where epr = '" & EPR(i) & "'"
'    Set rst1 = New ADODB.Recordset
'
'    rst1.Open sql9, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'        If Not rst1.EOF Then
'
'            If i > 1 Then
'               text1 = text1 & Chr(10) & "Dále na základì "
'            Else
'               text1 = text1 & "Na základì "
'            End If
'
'            If rst1![typ] = "rozsudek" Then
'                text1 = text1 & "rozsudku"
'                text1 = text1 & ", který vydal " & rst1(0) & " è.j. " & rst1(1) & " dne " & rst1(2) & ", který nabyl právní moci dne " & _
'                rst1(3) & " a" & Chr(160) & "je vykonatelný dnem " & rst1(4) & " (dále jen exekuèní titul), byla povinnému uložena povinnost  zaplatit oprávnìnému:" & Chr(10)
'
'            Else
'                text1 = text1 & "elektronického platebního rozkazu"
'                text1 = text1 & ", který vydal " & rst1(0) & " è.j. " & rst1(1) & " dne " & rst1(2) & ", který nabyl právní moci dne " & _
'                rst1(3) & " a" & Chr(160) & "je vykonatelný (dále jen exekuèní titul), byla povinnému uložena povinnost zaplatit oprávnìnému:" & Chr(10)
'
'            End If
'
'            'text1 = text1 & ", který vydal " & rst1(0) & " è.j. " & rst1(1) & " dne " & rst1(2) & ", který nabyl právní moci dne " & _
'            'rst1(3) & " a je vykonatelný dnem " & rst1(4) & " (dále jen exekuèní titul), byla povinnému uložena povinnost  zaplatit oprávnìnému:" & Chr(10)
'
'
'        End If
'    rst1.Close
'
'
'    sql_he = "select distinct hu from ap_coll.vs_import_pm_all where typ = 'HU'  and epr = '" & EPR(i) & "' and hu is not null order by hu" 'zjistime, kolik pro dane epr hu a jak se jmenuji
'    Set rst5 = New ADODB.Recordset
'    rst5.Open sql_he, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'
'        cnt_loan = 0
'        ReDim Loan(cnt_loan)
'
'        If Not rst5.EOF Then
'            cnt_loan = rst5.RecordCount
'            ReDim Loan(cnt_loan)
'            For ij = 1 To cnt_loan
'                Loan(ij) = rst5(0)
'                rst5.MoveNext
'            Next ij
'        End If
'
'    rst5.Close
'
'        '-------------------------------------------------------------hu------------
'    For ii = 1 To cnt_loan 'podle poctu hu budu dale vkladat intervaly uroceni
'
'        sql100 = "select comp from " & schema & ".v_exp_comp where hu = '" & Loan(ii) & "'" 'zjistime, zda psat castky dohromady nebo ne
'
'        Set rst1 = New ADODB.Recordset
'        rst1.Open sql100, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'            If Not rst1.EOF Then
'                comp = rst1(0) 'pro dane hu zjistime, zda psat dohromady castky, pokud true, piseme dohromady
'            End If
'        rst1.Close
'
'        Set rst7 = New ADODB.Recordset
'        sql7 = "select jistina, ku, su_sazba, su_castka, su_dtime_valid_from, ou_sazba, ou_castka, ou_dtime_valid_from from ap_coll.vs_import_pm_all where typ = 'HU' and hu = '" & Loan(ii) & "' and ou_dtime_valid_to > sysdate"
'        rst7.Open sql7, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic 'doplnime jistinu a k.urok
'            If Not rst7.EOF Then
'
'                If Not IsNull(rst7(0)) Then
'
'                    amt_principal = rst7(0)
'                Else
'                    amt_principal = 0
'                End If
'
'                If Not IsNull(rst7(1)) Then
'
'                    amt_interest = rst7(1)
'                Else
'                    amt_interest = 0
'                End If
'
'            End If
'        rst7.Close
'
'        text1ku = ""
'        If comp Then
'
'            text1 = text1 & Chr(10) & " -   èástku ve výši " & format$(amt_principal + amt_interest, "#,##0.00") & " Kè, " & Chr(10)
'
'        Else
'            text1 = text1 & Chr(10) & " -   èástku ve výši " & format$(amt_principal, "#,##0.00") & " Kè, " & Chr(10)
'            text1ku = " -   s kapitalizovaným úrokem ve výši " & format$(amt_interest, "#,##0.00") & " Kè, " & Chr(10)
'
'        End If
'
'
'        sql66 = "select jistina, ku, su_sazba, su_castka, su_dtime_valid_from, ou_sazba, ou_castka, ou_dtime_valid_from, su_dtime_valid_to, ou_dtime_valid_to " & _
'                " from ap_coll.vs_import_pm_all where typ = 'HU' and hu = '" & Loan(ii) & "' order by interval"
'        Set rst1 = New ADODB.Recordset
'        rst1.Open sql66, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic 'ke kazdemu hu ziskame udaje pocet intervalu
'            If Not rst1.EOF Then
'                cnt_int = rst1.RecordCount 'pocet intervalu k danemu hu
'                ReDim su_sazba(cnt_int)
'                ReDim su_castka(cnt_int)
'                ReDim su_dtime_valid_from(cnt_int)
'                ReDim ou_castka(cnt_int)
'                ReDim ou_sazba(cnt_int)
'                ReDim ou_dtime_valid_from(cnt_int)
'                ReDim ou_dtime_valid_to(cnt_int)
'                ReDim su_dtime_valid_to(cnt_int)
'                ReDim sankce(cnt_int)
'                ReDim ourok(cnt_int)
'
'
'                For iii = 1 To cnt_int
'
'                    If Not IsNull(rst1(2)) Then su_sazba(iii) = 100 * rst1![su_sazba] '(2)
'                    If Not IsNull(rst1(3)) Then su_castka(iii) = rst1![su_castka] '(3)
'                    If Not IsNull(rst1(4)) Then su_dtime_valid_from(iii) = rst1![su_dtime_valid_from] '(4)
'                    If Not IsNull(rst1(5)) Then ou_sazba(iii) = 100 * rst1![ou_sazba] '(5)
'                    If Not IsNull(rst1(6)) Then ou_castka(iii) = rst1![ou_castka] '(6)
'                    If Not IsNull(rst1(7)) Then ou_dtime_valid_from(iii) = rst1![ou_dtime_valid_from] '(7)
'                    If Not IsNull(rst1(8)) Then su_dtime_valid_to(iii) = rst1![su_dtime_valid_to] '(8)
'                    If Not IsNull(rst1(9)) Then ou_dtime_valid_to(iii) = rst1![ou_dtime_valid_to] '(9)
'
'                    rst1.MoveNext
'                Next iii
'
'            End If
'        rst1.Close
'
'        'poskladame zneni zakonneho a sankcniho uroku
'
'
'        ourok(1) = " -   s úrokem ve výši " & ou_sazba(1) & " % roènì z èástky " & format$(ou_castka(1), "#,##0.00") & " Kè od " & ou_dtime_valid_from(1)
'        sankce(1) = " -   se zákonným úrokem z prodlení ve výši " & format$(su_sazba(1), "#,##0.00") & " % roènì z èástky " & format$(su_castka(1), "#,##0.00") & " Kè od " & su_dtime_valid_from(1)
'
'        For q = 2 To cnt_int
'            sankce(q - 1) = sankce(q - 1) & " do " & su_dtime_valid_from(q) - 1 & "," & Chr(10)
'            sankce(q) = " -   se zákonným úrokem z prodlení ve výši " & su_sazba(q) & " % roènì z èástky " & format$(su_castka(q), "#,##0.00") & " Kè od " & su_dtime_valid_from(q)
'            ourok(q - 1) = ourok(q - 1) & " do " & su_dtime_valid_from(q) - 1 & "," & Chr(10)
'            ourok(q) = " -   s úrokem ve výši " & ou_sazba(q) & " % roènì z èástky " & format$(ou_castka(q), "#,##0.00") & " Kè od " & ou_dtime_valid_from(q)
'        Next q
'
'        sankce(cnt_int) = sankce(cnt_int) & " do zaplacení," & Chr(10)
'        ourok(cnt_int) = ourok(cnt_int) & " do zaplacení," & Chr(10)
'
'
'        For qe = 1 To cnt_int
'            text1 = text1 & ourok(qe) & sankce(qe)
'        Next qe
'
'        text1 = text1 & text1ku
'        inc2 = inc 'definujeme zde, protoze ve vyberu oblasti to chceme az u dalsiho epr (v dalsim kole cyklu)
'
'    Next ii
'     '-------------------------------------------------------------HU-----
'
'
'
'    'na urovni celeho epr zmenime jistinu, k. urok a sop
'    'musime zjistit sop a debet
'
'    sql_re = "select nvl(sum(debet),0), nvl(sum(sop),0) from  ap_coll.vs_import_pm_all where epr = '" & EPR(i) & "'"
'
'    debet = 0
'    SOP = 0
'    Set rst8 = New ADODB.Recordset
'    rst8.Open sql_re, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'    If Not rst8.EOF Then
'        debet = rst8(0)
'        SOP = rst8(1)
'    End If
'    rst8.Close
'
'    debet_text = ""
'    If debet > 0 Then
'        debet_text = Chr(10) & " -   èástku ve výši " & format$(debet, "#,##0.00") & " Kè,"
'
'        inc = inc + 1
'        inc2 = inc2 + 1
'    End If
'
'    sop_text = ""
'    If SOP > 0 Then
'        sop_text = vbCrLf & vbCrLf & " -   náklady nalézacího øízení ve výši " & format$(SOP, "#,##0.00") & " Kè"
'        'sop_text = vbNewLine & " -   náklady nalézacího øízení ve výši " & Format$(sop, "#,##0.00") & " Kè"
'    End If
'
'    text1 = text1 & debet_text & sop_text & "."
'
'Next i
'
'GenerujTextCast_I = text1
'
'End Function
'
'
'
'
'Private Function GenerujTextCast_I_OLD() As String
'Dim sql0 As String, sql66 As String, sql7 As String, sql100 As String, sql1 As String, sql2 As String
'Dim sql_he As String, sql_h As String, Loan() As String, cnt_int As Integer, amt_principal As Double, amt_interest As Double
'Dim su_sazba() As Double, su_castka() As Double, ou_sazba() As Double, ou_castka() As Double
'Dim su_dtime_valid_to() As Date, ou_dtime_valid_to() As Date, su_dtime_valid_from() As Date, ou_dtime_valid_from() As Date
'Dim inc2 As Integer, inc3 As Integer, sql_re As String, debet As Double, SOP As Double
'Dim text1 As String, sankce() As String, ourok() As String
'Dim debet_text As String
'Dim sop_text As String
'Dim inc As Variant
'Dim sql9 As String
'Dim text1ku As String
'Dim cnt_loan As Integer
'Dim comp As Boolean
'
''-------------------------------------------------------------EPR---------------------
'For i = 1 To lngPocetZalob
''dynamicky definovane pole, ktere se meni podle poctu epr
'    amt_principal = 0 'pro kazde epr zacneme napocitavat jistinu a urok od nuly
'    amt_interest = 0
'
'    sql9 = "select soud, epr, datum_vydani, datum_pm, datum_vykon, typ from ap_coll.vs_import_pm_rozsudek_all where epr = '" & EPR(i) & "'"
'    Set rst1 = New ADODB.Recordset
'
'    rst1.Open sql9, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'        If Not rst1.EOF Then
'
'            If i > 1 Then
'               text1 = text1 & Chr(10) & "Dále na základì "
'            Else
'               text1 = text1 & "Na základì "
'            End If
'
'            If rst1![typ] = "rozsudek" Then
'                text1 = text1 & "rozsudku"
'                text1 = text1 & ", který vydal " & rst1(0) & " è.j. " & rst1(1) & " dne " & rst1(2) & ", který nabyl právní moci dne " & _
'                rst1(3) & " a" & Chr(160) & "je vykonatelný dnem " & rst1(4) & " (dále jen exekuèní titul), byla povinnému uložena povinnost  zaplatit oprávnìnému:" & Chr(10)
'
'            Else
'                text1 = text1 & "elektronického platebního rozkazu"
'                text1 = text1 & ", který vydal " & rst1(0) & " è.j. " & rst1(1) & " dne " & rst1(2) & ", který nabyl právní moci dne " & _
'                rst1(3) & " a" & Chr(160) & "je vykonatelný (dále jen exekuèní titul), byla povinnému uložena povinnost zaplatit oprávnìnému:" & Chr(10)
'
'            End If
'
'            'text1 = text1 & ", který vydal " & rst1(0) & " è.j. " & rst1(1) & " dne " & rst1(2) & ", který nabyl právní moci dne " & _
'            'rst1(3) & " a je vykonatelný dnem " & rst1(4) & " (dále jen exekuèní titul), byla povinnému uložena povinnost  zaplatit oprávnìnému:" & Chr(10)
'
'
'        End If
'    rst1.Close
'
'
'    sql_he = "select distinct hu from ap_coll.vs_import_pm_all where typ = 'HU'  and epr = '" & EPR(i) & "' and hu is not null order by hu" 'zjistime, kolik pro dane epr hu a jak se jmenuji
'    Set rst5 = New ADODB.Recordset
'    rst5.Open sql_he, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'
'        cnt_loan = 0
'        ReDim Loan(cnt_loan)
'
'        If Not rst5.EOF Then
'            cnt_loan = rst5.RecordCount
'            ReDim Loan(cnt_loan)
'            For ij = 1 To cnt_loan
'                Loan(ij) = rst5(0)
'                rst5.MoveNext
'            Next ij
'        End If
'
'    rst5.Close
'
'        '-------------------------------------------------------------hu------------
'    For ii = 1 To cnt_loan 'podle poctu hu budu dale vkladat intervaly uroceni
'
'        sql100 = "select comp from " & schema & ".v_exp_comp where hu = '" & Loan(ii) & "'" 'zjistime, zda psat castky dohromady nebo ne
'
'        Set rst1 = New ADODB.Recordset
'        rst1.Open sql100, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'            If Not rst1.EOF Then
'                comp = rst1(0) 'pro dane hu zjistime, zda psat dohromady castky, pokud true, piseme dohromady
'            End If
'        rst1.Close
'
'        Set rst7 = New ADODB.Recordset
'        sql7 = "select jistina, ku, su_sazba, su_castka, su_dtime_valid_from, ou_sazba, ou_castka, ou_dtime_valid_from from ap_coll.vs_import_pm_all where typ = 'HU' and hu = '" & Loan(ii) & "' and ou_dtime_valid_to > sysdate"
'        rst7.Open sql7, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic 'doplnime jistinu a k.urok
'            If Not rst7.EOF Then
'
'                If Not IsNull(rst7(0)) Then
'
'                    amt_principal = rst7(0)
'                Else
'                    amt_principal = 0
'                End If
'
'                If Not IsNull(rst7(1)) Then
'
'                    amt_interest = rst7(1)
'                Else
'                    amt_interest = 0
'                End If
'
'            End If
'        rst7.Close
'
'        text1ku = ""
'        If comp Then
'
'            text1 = text1 & Chr(10) & " -   èástku ve výši " & format$(amt_principal + amt_interest, "#,##0.00") & " Kè, " & Chr(10)
'
'        Else
'            text1 = text1 & Chr(10) & " -   èástku ve výši " & format$(amt_principal, "#,##0.00") & " Kè, " & Chr(10)
'            text1ku = " -   s kapitalizovaným úrokem ve výši " & format$(amt_interest, "#,##0.00") & " Kè, " & Chr(10)
'
'        End If
'
'
'        sql66 = "select jistina, ku, su_sazba, su_castka, su_dtime_valid_from, ou_sazba, ou_castka, ou_dtime_valid_from, su_dtime_valid_to, ou_dtime_valid_to " & _
'                " from ap_coll.vs_import_pm_all where typ = 'HU' and hu = '" & Loan(ii) & "' order by interval"
'        Set rst1 = New ADODB.Recordset
'        rst1.Open sql66, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic 'ke kazdemu hu ziskame udaje pocet intervalu
'            If Not rst1.EOF Then
'                cnt_int = rst1.RecordCount 'pocet intervalu k danemu hu
'                ReDim su_sazba(cnt_int)
'                ReDim su_castka(cnt_int)
'                ReDim su_dtime_valid_from(cnt_int)
'                ReDim ou_castka(cnt_int)
'                ReDim ou_sazba(cnt_int)
'                ReDim ou_dtime_valid_from(cnt_int)
'                ReDim ou_dtime_valid_to(cnt_int)
'                ReDim su_dtime_valid_to(cnt_int)
'                ReDim sankce(cnt_int)
'                ReDim ourok(cnt_int)
'
'
'                For iii = 1 To cnt_int
'
'                    If Not IsNull(rst1(2)) Then su_sazba(iii) = 100 * rst1![su_sazba] '(2)
'                    If Not IsNull(rst1(3)) Then su_castka(iii) = rst1![su_castka] '(3)
'                    If Not IsNull(rst1(4)) Then su_dtime_valid_from(iii) = rst1![su_dtime_valid_from] '(4)
'                    If Not IsNull(rst1(5)) Then ou_sazba(iii) = 100 * rst1![ou_sazba] '(5)
'                    If Not IsNull(rst1(6)) Then ou_castka(iii) = rst1![ou_castka] '(6)
'                    If Not IsNull(rst1(7)) Then ou_dtime_valid_from(iii) = rst1![ou_dtime_valid_from] '(7)
'                    If Not IsNull(rst1(8)) Then su_dtime_valid_to(iii) = rst1![su_dtime_valid_to] '(8)
'                    If Not IsNull(rst1(9)) Then ou_dtime_valid_to(iii) = rst1![ou_dtime_valid_to] '(9)
'
'                    rst1.MoveNext
'                Next iii
'
'            End If
'        rst1.Close
'
'        'poskladame zneni zakonneho a sankcniho uroku
'
'
'        ourok(1) = " -   s úrokem ve výši " & ou_sazba(1) & " % roènì z èástky " & format$(ou_castka(1), "#,##0.00") & " Kè od " & ou_dtime_valid_from(1)
'        sankce(1) = " -   se zákonným úrokem z prodlení ve výši " & format$(su_sazba(1), "#,##0.00") & " % roènì z èástky " & format$(su_castka(1), "#,##0.00") & " Kè od " & su_dtime_valid_from(1)
'
'        For q = 2 To cnt_int
'            sankce(q - 1) = sankce(q - 1) & " do " & su_dtime_valid_from(q) - 1 & "," & Chr(10)
'            sankce(q) = " -   se zákonným úrokem z prodlení ve výši " & su_sazba(q) & " % roènì z èástky " & format$(su_castka(q), "#,##0.00") & " Kè od " & su_dtime_valid_from(q)
'            ourok(q - 1) = ourok(q - 1) & " do " & su_dtime_valid_from(q) - 1 & "," & Chr(10)
'            ourok(q) = " -   s úrokem ve výši " & ou_sazba(q) & " % roènì z èástky " & format$(ou_castka(q), "#,##0.00") & " Kè od " & ou_dtime_valid_from(q)
'        Next q
'
'        sankce(cnt_int) = sankce(cnt_int) & " do zaplacení," & Chr(10)
'        ourok(cnt_int) = ourok(cnt_int) & " do zaplacení," & Chr(10)
'
'
'        For qe = 1 To cnt_int
'            text1 = text1 & ourok(qe) & sankce(qe)
'        Next qe
'
'        text1 = text1 & text1ku
'        inc2 = inc 'definujeme zde, protoze ve vyberu oblasti to chceme az u dalsiho epr (v dalsim kole cyklu)
'
'    Next ii
'     '-------------------------------------------------------------HU-----
'
'
'
'    'na urovni celeho epr zmenime jistinu, k. urok a sop
'    'musime zjistit sop a debet
'
'    sql_re = "select nvl(sum(debet),0), nvl(sum(sop),0) from  ap_coll.vs_import_pm_all where epr = '" & EPR(i) & "'"
'
'    debet = 0
'    SOP = 0
'    Set rst8 = New ADODB.Recordset
'    rst8.Open sql_re, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'    If Not rst8.EOF Then
'        debet = rst8(0)
'        SOP = rst8(1)
'    End If
'    rst8.Close
'
'    debet_text = ""
'    If debet > 0 Then
'        debet_text = Chr(10) & " -   èástku ve výši " & format$(debet, "#,##0.00") & " Kè,"
'
'        inc = inc + 1
'        inc2 = inc2 + 1
'    End If
'
'    sop_text = ""
'    If SOP > 0 Then
'        sop_text = vbCrLf & vbCrLf & " -   náklady nalézacího øízení ve výši " & format$(SOP, "#,##0.00") & " Kè"
'        'sop_text = vbNewLine & " -   náklady nalézacího øízení ve výši " & Format$(sop, "#,##0.00") & " Kè"
'    End If
'
'    text1 = text1 & debet_text & sop_text & "."
'
'Next i
'
'GenerujTextCast_I_OLD = text1
'
'End Function
'
'
'
'Private Sub UlozDokument(ByRef WordDocument As Variant)
'
'Dim AFolder As String, BFolder As String
'Dim fso As Object
'Dim sql3 As String
'
'sql3 = "select * from " & schema & ".exe_status_data where skp_exe_status = 1 and id_person = " & id_person
'Set rst3 = New ADODB.Recordset
'rst3.Open sql3, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'    If Not rst3.EOF Then
'        AFolder = cfg.Cesta("ExePodane") & c.Nazev_slozky_bez_lomitka
'        BFolder = "W:\Vymahani\LATE\EXEKUCE\PODANÉ\Klienti bez složky"
'    Else
'        AFolder = "W:\Vymahani\LATE\EXEKUCE\PØIPRAVENÉ\" & c.Nazev_slozky_bez_lomitka
'        BFolder = "W:\Vymahani\LATE\EXEKUCE\PØIPRAVENÉ\Klienti bez složky"
'    End If
'rst3.Close
'
'
'Set fso = CreateObject("Scripting.FileSystemObject")
'
'If schema = "APP_COLL_JEZEVCIK_TEST" Then
'    WordDocument.SaveAs (cfg.Cesta("TestovaciSlozka") & "\Exekuèní návrh klienta " & c.name & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".docx")
'    'Call Shell("explorer.exe " & cfg.Cesta("TestovaciSlozka"), vbNormalFocus) 'otvirani slozky na testu
'Else
'    If fso.FolderExists(AFolder) Then
'        WordDocument.SaveAs (AFolder & "\Exekuèní návrh klienta " & c.name & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".docx")
'        Call Shell("explorer.exe " & AFolder, vbNormalFocus)
'    ElseIf fso.FolderExists(BFolder) Then
'        WordDocument.SaveAs (BFolder & "\Exekuèní návrh klienta " & c.name & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".docx")
'        Call Shell("explorer.exe " & BFolder, vbNormalFocus)
'    Else
'        MsgBox "Složka pro klienta " & c.name & " neexistuje." ', vbInformation, "Složka nenalezena!"
'        WordDocument.SaveAs (BFolder & "\Exekuèní návrh klienta " & c.name & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".docx")
'    End If
'
'End If
'
'
'
'End Sub
'
'
'
'Private Sub komentar_pred_ulozenim()
'' 200 riadkov zakomentovaneho kodu
'' najdes v jezevciku verze 5.1
'End Sub
'
'
'
'
'
'
'
''''''''''''''''
''''''''''''''''
''''''''''''''''
''''''''''''''''
''''''''''''''''
''''''''''''''''
''''''''''''''''
''''''''''''''''
''''''''''''''''
''''''''''''''''  odtialto bordel
''''''''''''''''
''''''''''''''''
''''''''''''''''
''''''''''''''''
''''''''''''''''
''''''''''''''''
''''''''''''''''
''''''''''''''''
''''''''''''''''
''''''''''''''''
'
'
'
'
'
'
'
'
'
'
''Option Explicit
''
''Private Function VytvorTextCast_II_old(Pohl_Zaplacene() As Pohl_zaplacena, sd As data_pro_sablonu) As String
''    ' ?? meni sa format podla rozsudku?? (ked tak sa prida argument..)
''
''    ' LOGIKA FUNKCE:
''    ' texty sa NEgeneruju zaradom ako ich je vidiet vo vyslednej sablone.
''    ' Najprv sa spravi loop cez vsetky prichodzie platby a ku kazdej sa vytvori uvodny kus textacie
''    ' Potom sa spravi loop cez Pohl_Zaplacene, a ku kazdej platbe sa text doplni podla toho, ci je tam jedina alebo viac platieb
''    ' Az potom sa to spoji do jedneho textu
''    ' Na konci sa pripise extra veta ak mame nejaky Pausal
''
''
''    Dim arrText() As String
''
''    Dim i As Integer
''    Dim j As Integer
''
''    If sd.PocetPlatebCast_II > 0 Then
''        ReDim arrText(1 To sd.PocetPlatebCast_II)
''
''
''
''        '''   Uvodna cast textu pre kazdu platbu
''        For i = 1 To sd.PocetPlatebCast_II
'''            arrText(i) = textace("-") & textace("I_castka") & VlozCastku(sd.castkyZaplacenePohledavky(i)) & _
'''                       " ze dne " & sd.datumyZaplacenePohledavky(i) & " zapoèetl oprávnìný"
''            arrText(i) = textace("I_castka") & VlozCastku(sd.castkyZaplacenePohledavky(i)) & _
''                        " ze dne " & sd.datumyZaplacenePohledavky(i) & " zapoèetl oprávnìný"
''        Next i
''
''
''
''        '''   Zvysok textu pre kazdu platbu, rozdelenie na jednotlive ucely
''        For i = 1 To sd.PocetPlatebCast_II
''
''            For j = 1 To UBound(Pohl_Zaplacene)
''
''                If (Left(sd.EPR(PoradoveCisloEPR), 3) = "EPR") Then
''
''                    If Pohl_Zaplacene(j).datum_zaplaceni = sd.datumyZaplacenePohledavky(i) Then
''                        If Pohl_Zaplacene(j).castka = sd.castkyZaplacenePohledavky(i) Then
''                            ' cela platba sla na jediny ucel
''                            arrText(i) = arrText(i) & " na " & Pohl_Zaplacene(j).kategorie_akuzativ & textace("II_pohl_pism") & _
''                                         sd.poradove_pismeno_uveru(Pohl_Zaplacene(j).nazevUveru) & textace(")") & "EPR"
''                        Else
''                            ' platba byla rozdelena na nekolik ucelu (probehne vicekrat nasl. radek)
''
''                            'If Pohl_Zaplacene(j - i).nazevUveru = Pohl_Zaplacene(j).nazevUveru And Pohl_Zaplacene(j - 1).datum_zaplaceni = Pohl_Zaplacene(j).datum_zaplaceni Then
''                                ' platba z rovnakeho data na rovnaky uver jako predesla
''                            arrText(i) = arrText(i) & textace("radek") & textace("-") & textace("II_cast") & VlozCastku(Pohl_Zaplacene(j).castka) & " na " & _
''                                         Pohl_Zaplacene(j).kategorie_akuzativ & textace("II_pohl_pism") & sd.poradove_pismeno_uveru(Pohl_Zaplacene(j).nazevUveru) & textace(")") & "EPR"
''                        End If
''                    End If
''
''                Else   ' rozsudek
''
''                    ' misto poradoveho pismena EPR se vklada poradove cislo rozsudku
''
''                    If Pohl_Zaplacene(j).datum_zaplaceni = sd.datumyZaplacenePohledavky(i) Then
''                        If Pohl_Zaplacene(j).castka = sd.castkyZaplacenePohledavky(i) Then
''                            ' cela platba sla na jediny ucel
''                            arrText(i) = arrText(i) & " na " & Pohl_Zaplacene(j).kategorie_akuzativ & textace("II_pohl_cislo") & _
''                                         sd.poradove_cislo_uveru(Pohl_Zaplacene(j).nazevUveru) & textace(")") & "rozsudku"
''                        Else
''                            ' platba byla rozdelena na nekolik ucelu (probehne vicekrat nasl. radek)
''
''                            arrText(i) = arrText(i) & textace("radek") & textace("-") & textace("II_cast") & VlozCastku(Pohl_Zaplacene(j).castka) & " na " & _
''                                         Pohl_Zaplacene(j).kategorie_akuzativ & textace("II_pohl_cislo") & sd.poradove_cislo_uveru(Pohl_Zaplacene(j).nazevUveru) & textace(")") & "rozsudku"
''                        End If
''                    End If
''
''                End If
''
''
''
''
''
''                ''' odtialto starsi verze
''''                If Pohl_Zaplacene(j).datum_zaplaceni = sd.datumyZaplacenePohledavky(i) Then
''''                    If Pohl_Zaplacene(j).castka = sd.castkyZaplacenePohledavky(i) Then
''''                        ' cela platba sla na jediny ucel
''''                        arrText(i) = arrText(i) & " na " & Pohl_Zaplacene(j).kategorie_akuzativ & textace("II_pohl_pism") & sd.poradove_pismeno_uveru(Pohl_Zaplacene(j).nazevUveru) & textace(")") & _
''''                                     IIf(Left(sd.EPR(PoradoveCisloEPR), 3) = "EPR", "EPR", "rozsudku") '& sd.EPR ' FAXX podla coho zistim pismeno? a ze je to EPR? cislo pohladavky? (detto o 3 riadky nizsie)
''''                    Else
''''                        ' platba byla rozdelena na nekolik ucelu (probehne vicekrat nasl. radek)
''''                        arrText(i) = arrText(i) & textace("radek") & textace("-") & textace("II_cast") & VlozCastku(Pohl_Zaplacene(j).castka) & _
''''                                     " na " & Pohl_Zaplacene(j).kategorie_akuzativ & textace("II_pohl_pism") & sd.poradove_pismeno_uveru(Pohl_Zaplacene(j).nazevUveru) & textace(")") & _
''''                                     IIf(Left(sd.EPR(PoradoveCisloEPR), 3) = "EPR", "EPR", "rozsudku")
''''                    End If
''''
''''                End If
''
''            Next j
''
''        Next i
''
''
''        '''   Spojime to do jedineho stringu
''        For i = LBound(arrText) To UBound(arrText)
''            VytvorTextCast_II = VytvorTextCast_II & textace("radek") & textace("radek") & arrText(i)
''        Next i
''
''    End If
''
''    '''   Pokud ma klient nejaky pausal, tak ten NEvymahame a do cl.II se vlozi veta navic
''
''
''    If c.SOP.Pausal > 0 Then
''        VytvorTextCast_II = VytvorTextCast_II & textace("radek") & textace("radek") & textace("II_pouze_cast") '& _
''                            '" (XXXXXX Jezevèík tady mùže uvést i èástku: " & c.SOP.Pausal & ".)" '& textace("radek")
''    End If
''
''
''    '''''      FA v povodnom kode tu este bol kusok (podla mna) zbytocneho kodu
''
''
''End Function
''
''
''
''
''
''
''
''Private Function VytvorTextCast_II_uvod_old(sd As data_pro_sablonu) As String
''' pozor, fce taky naplni sd.PocetPlatebCast_II
''
''Dim sql As String
''Dim rst As adodb.Recordset
''Dim text2 As String
''
''
''    sql = "select date_payment, sum(amt_paid) as amt_paid from (select trunc(date_payment) date_payment, epr, amt_paid from ap_coll.pb_040_uhrady_pred_parovanim) where epr in (" & sd.epr_all & ") group by date_payment order by date_payment asc"
''
''    Set rst = New adodb.Recordset
''    rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''
''        If Not rst.EOF Then
''
''            rst.MoveFirst
''
''            ' flag_platby = True 'pokud neco uhradil
''            ' ReDim platby_rozpis(rst.RecordCount)
''            ' ReDim amt_paid(rst.RecordCount)
''            ' ReDim date_p(rst.RecordCount)
''            ' ReDim platby_par(rst.RecordCount)
''
''            ReDim sd.datumyZaplacenePohledavky(1 To rst.RecordCount)
''            ReDim sd.castkyZaplacenePohledavky(1 To rst.RecordCount)
''            sd.PocetPlatebCast_II = rst.RecordCount
''            text2 = textace("II_uhradil") & textace("radek") & textace("radek")
''            For i = 1 To rst.RecordCount
''
''                text2 = text2 & textace("-") & "dne " & format(rst![date_payment], "Short Date") & " " & textace("I_castka") & VlozCastku(rst![amt_paid]) & textace("radek") 'puvodne: format$(rst![amt_paid], "#,##0.00") & " Kè," & vbCrLf
''
''                ' amt_paid(i) = rst![amt_paid]
''                ' date_p(i) = format(rst![date_payment], "Short Date")
''                sd.datumyZaplacenePohledavky(i) = format(rst![date_payment], "Short Date")
''                sd.castkyZaplacenePohledavky(i) = rst![amt_paid]
''
''                rst.MoveNext
''            Next i
''            text2 = text2 & textace("radek") & textace("II_castky_zapocet") '& textace("radek")
''
''        Else
''
''            text2 = text2 & textace("II_nesplnil")
''
''        End If
''    rst.Close
''
''    VytvorTextCast_II_uvod_old = text2
''
''End Function
''
''
''
''
''
''Private Function NacitejPohledavkyCast_II_old(ByRef sd As data_pro_sablonu) As Pohl_zaplacena()
''    ' naplni se pohledavky, ktere (alespon castecne) byli uhrazeny splatkami.
''    ' jde o to, jak byli rozdelene jednotlive platby (co slo na jistinu/urok/.. ktereho uveru)
''    ' jeste pred touhle funkci se musi zavolat VytvorTextCast_II_uvod, ve ktere se plni sd.VytvorTextCast_II_uvod a sd.datumyZaplacenePohledavky
''
''
''    Dim arrPohlZapl() As Pohl_zaplacena
''    Dim intPocetPohledavek As Integer
''    Dim rst As adodb.Recordset, sql As String
''    Dim i As Integer
''    Dim cisloPohl As Integer
''
''    cisloPohl = 0
''    ReDim arrPohlZapl(1 To 300) ' na konci se zkrati
''
''    sql = "SELECT a.ID_PERSON, a.DATUM_PLATBY_TRUNC, a.UVER_CISLO, a.amt_payment_jistina, a.amt_payment_urok " & _
''          " FROM " & schema & ".V_EXE_NAVRH_UVERY_GROUPED a " & _
''          " WHERE a.id_person = " & c.id_person & _
''          " ORDER BY a.DATUM_PLATBY_TRUNC, a.UVER_CISLO "
''
''    Set rst = New adodb.Recordset
''    rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''
''        If Not rst.EOF Then
''
''            rst.MoveFirst
''
''            For i = 1 To rst.RecordCount
''
''                '''   pohledavka Jistina
''                If Not IsNull(rst![amt_payment_jistina]) Then
''                    If rst![amt_payment_jistina] > 0 Then
''                        cisloPohl = cisloPohl + 1
''
''                        With arrPohlZapl(cisloPohl)
''                            .castka = rst![amt_payment_jistina]
''                            .datum_zaplaceni = rst![DATUM_PLATBY_TRUNC]
''                            .enumTypCastky = TypCastky_jistina
''                            .enumTypPohl = TypPohl_uver
''                            .kategorie_akuzativ = "jistinu"
''                            .nazevUveru = rst![UVER_CISLO]
''                        End With
''
''                    End If
''                End If
''
''
''                '''   pohledavka Urok
''                If Not IsNull(rst![amt_payment_urok]) Then
''                    If rst![amt_payment_urok] > 0 Then
''                        cisloPohl = cisloPohl + 1
''
''                        With arrPohlZapl(cisloPohl)
''                            .castka = rst![amt_payment_urok]
''                            .datum_zaplaceni = rst![DATUM_PLATBY_TRUNC]
''                            .enumTypCastky = TypCastky_OU
''                            .enumTypPohl = TypPohl_uver
''                            .kategorie_akuzativ = "úrok"
''                            .nazevUveru = rst![UVER_CISLO]
''                        End With
''
''                    End If
''                End If
''
''                rst.MoveNext
''
''            Next i
''        End If
''    rst.Close
''
''
''    '''   naklady rizeni
''        ' todo
''
''    '''   debet
''        ' todo
''
''    If cisloPohl > 0 Then
''        ReDim Preserve arrPohlZapl(1 To cisloPohl)
''    End If
''    NacitejPohledavkyCast_II_old = arrPohlZapl
''
''End Function
''
''
''
''
''
''Private Function NacitejPohledavkyCast_II_ap_coll(ByRef sd As data_pro_sablonu) As Pohl_zaplacena()
''' naplni se pohledavky, ktere (alespon castecne) byli uhrazeny splatkami.
''' jde o to, jak byli rozdelene jednotlive platby (co slo na jistinu/urok/.. ktereho uveru)
''' jeste pred touhle funkci se musi zavolat VytvorTextCast_II_uvod, ve ktere se plni sd.VytvorTextCast_II_uvod a sd.datumyZaplacenePohledavky
''
''
''Dim arrPohlZapl() As Pohl_zaplacena
''Dim intPocetPohledavek As Integer
''Dim rst As adodb.Recordset, sql As String
''Dim i As Integer, ii As Integer
''Dim cisloPohl As Integer
''
''    cisloPohl = 1
''    ReDim arrPohlZapl(1 To 300) ' na konci se zkrati
''    For i = 1 To sd.PocetPlatebCast_II 'pro kazdou platbu udelame zvlast dotaz
''
''        Set rst = New adodb.Recordset
''        ' tabulka do ktere se diva view: AP_COLL.PB_040_PAROVANI
''        sql = "select a.EPR, a.KATEGORIE, a.AMT_PAID, a.DATE_PAYMENT, a.KATEGORIE_POR " & _
''                " FROM " & schema & ".v_exp_exe_par a " & _
''                " WHERE epr in (" & sd.epr_all & ") and date_payment = to_date('" & format(sd.datumyZaplacenePohledavky(i), "Short Date") & "') order by kategorie_por"
''        rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''
''            If Not rst.EOF Then
''
''                rst.MoveFirst
''                For ii = 1 To rst.RecordCount
''
''                    ' 'kapitalizovaneho uroku',  'jistiny',  'naklady rizeni',  'sankcniho a obchodniho uroku',  'debetu',
''                    With arrPohlZapl(cisloPohl)
''                        .castka = rst![amt_paid]
''                        .datum_zaplaceni = rst![date_payment]
''
''                        Select Case rst![kategorie]
''
''                            Case "kapitalizovaneho uroku", "kapitalizovaný úrok"
''                                .enumTypPohl = TypPohl_uver
''                                .enumTypCastky = TypCastky_KU
''                                .kategorie_akuzativ = "kapitalizovaný úrok"
''
''                            Case "jistiny", "jistinu"
''                                .enumTypPohl = TypPohl_uver
''                                .enumTypCastky = TypCastky_jistina
''                                .kategorie_akuzativ = "jistinu"
''
''                            Case "sankcniho a obchodniho uroku", "sankèní a obchodní úrok"
''                                .enumTypPohl = TypPohl_uver
''                                .enumTypCastky = TypCastky_OU ' OR TypCastky_ZU, je to jedno..
''                                .kategorie_akuzativ = "sankèní a obchodní úrok"
''
''                            Case "naklady rizeni", "náklady øízení"
''                                .enumTypPohl = TypPohl_naklady_rizeni
''                                .kategorie_akuzativ = "náklady øízení"
''
''                            Case "debetu", "debet"
''                                .enumTypPohl = TypPohl_debet
''                                .kategorie_akuzativ = "debet"
''
''                            Case Else
''                                Debug.Assert False
''                                MsgBox "Jez-error  J522331", vbCritical
''                        End Select
''
''                        ' .nazevUveru
''
''                    End With
''
''                    cisloPohl = cisloPohl + 1
''                    rst.MoveNext
''
''                Next ii
''            End If
''        rst.Close
''    Next i
''
''    ReDim Preserve arrPohlZapl(1 To cisloPohl)
''    NacitejPohledavkyCast_II_ap_coll = arrPohlZapl
''
''End Function
''
''
''
''
''
''
''
''''''
''''''Private Function VytvorTextCast_I_uvod_old(nazvy_EPR() As String) As String
'''''''''   Text na zaklade rozsudku/EPR
''''''' cislo_epr = sd.EPR(i)
''''''Dim text1 As String
''''''Dim sql As String
''''''Dim rst As ADODB.Recordset
''''''
''''''
''''''    For i = LBound(nazvy_EPR) To UBound(nazvy_EPR)
''''''
''''''        sql = "select soud, epr, datum_vydani, datum_pm, datum_vykon, typ from ap_coll.vs_import_pm_rozsudek_all where epr = '" & nazvy_EPR(i) & "'"
''''''        Set rst = New ADODB.Recordset
''''''
''''''        rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''''''            If Not rst.EOF Then
''''''
''''''                If i > 1 Then
''''''                   text1 = text1 & Chr(10) & "Dále na základì "
''''''                Else
''''''                   text1 = text1 & "Na základì "
''''''                End If
''''''
''''''                If rst![typ] = "rozsudek" Then
''''''                    text1 = text1 & "rozsudku"
''''''                    text1 = text1 & ", který vydal " & rst![Soud] & " è.j. " & rst![EPR] & " dne " & rst![datum_vydani] & ", který nabyl právní moci dne " & _
''''''                    rst![datum_pm] & " a" & Chr(160) & "je vykonatelný dnem " & rst![datum_vykon] & " (dále jen exekuèní titul), byla povinnému uložena povinnost  zaplatit oprávnìnému:" & Chr(10)
''''''
''''''                Else
''''''                    text1 = text1 & "elektronického platebního rozkazu"
''''''                    text1 = text1 & ", který vydal " & rst![Soud] & " è.j. " & rst![EPR] & " dne " & rst![datum_vydani] & ", který nabyl právní moci dne " & _
''''''                    rst![datum_pm] & " a" & Chr(160) & "je vykonatelný (dále jen exekuèní titul), byla povinnému uložena povinnost zaplatit oprávnìnému:" & Chr(10)
''''''
''''''                End If
''''''
''''''                'text1 = text1 & ", který vydal " & rst![soud] & " è.j. " & rst![epr] & " dne " & rst![datum_vydani] & ", který nabyl právní moci dne " & _
''''''                'rst![datum_pm] & " a je vykonatelný dnem " & rst![datum_vykon] & " (dále jen exekuèní titul), byla povinnému uložena povinnost  zaplatit oprávnìnému:" & Chr(10)
''''''
''''''
''''''            End If
''''''        rst.Close
''''''
''''''    Next i
''''''
''''''VytvorTextCast_I_uvod = text1
''''''End Function
''
''
''
''
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
''
''Private Function GenerujTextCast_I(ByRef sd As data_pro_sablonu) As String
''Dim sql0 As String, sql66 As String, sql7 As String, sql100 As String, sql1 As String, sql2 As String
''Dim sql_he As String, sql_h As String, Loan() As String, cnt_int As Integer, amt_principal As Double, amt_interest As Double
''Dim su_sazba() As Double, su_castka() As Double, ou_sazba() As Double, ou_castka() As Double
''Dim su_dtime_valid_to() As Date, ou_dtime_valid_to() As Date, su_dtime_valid_from() As Date, ou_dtime_valid_from() As Date
''Dim inc2 As Integer, inc3 As Integer, sql_re As String, debet As Double, SOP As Double
''Dim text1 As String, sankce() As String, ourok() As String
''Dim debet_text As String
''Dim sop_text As String
''Dim inc As Variant
''Dim sql9 As String
''Dim text1ku As String
''Dim cnt_loan As Integer
''Dim comp As Boolean
''
'''-------------------------------------------------------------EPR---------------------
''For i = 1 To sd.pocetZalob
'''dynamicky definovane pole, ktere se meni podle poctu epr
''    amt_principal = 0 'pro kazde epr zacneme napocitavat jistinu a urok od nuly
''    amt_interest = 0
''
''
''    '''   Text na zaklade rozsudku/EPR
''
''    sql9 = "select soud, epr, datum_vydani, datum_pm, datum_vykon, typ from ap_coll.vs_import_pm_rozsudek_all where epr = '" & sd.EPR(i) & "'"
''    Set rst1 = New adodb.Recordset
''
''    rst1.Open sql9, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''        If Not rst1.EOF Then
''
''            If i > 1 Then
''               text1 = text1 & Chr(10) & "Dále na základì "
''            Else
''               text1 = text1 & "Na základì "
''            End If
''
''            If rst1![typ] = "rozsudek" Then
''                text1 = text1 & "rozsudku"
''                text1 = text1 & ", který vydal " & rst1(0) & " è.j. " & rst1(1) & " dne " & rst1(2) & ", který nabyl právní moci dne " & _
''                rst1(3) & " a" & Chr(160) & "je vykonatelný dnem " & rst1(4) & " (dále jen exekuèní titul), byla povinnému uložena povinnost zaplatit oprávnìnému:" & Chr(10)
''
''            Else
''                text1 = text1 & "elektronického platebního rozkazu"
''                text1 = text1 & ", který vydal " & rst1(0) & " è.j. " & rst1(1) & " dne " & rst1(2) & ", který nabyl právní moci dne " & _
''                rst1(3) & " a" & Chr(160) & "je vykonatelný (dále jen exekuèní titul), byla povinnému uložena povinnost zaplatit oprávnìnému:" & Chr(10)
''
''            End If
''
''            'text1 = text1 & ", který vydal " & rst1(0) & " è.j. " & rst1(1) & " dne " & rst1(2) & ", který nabyl právní moci dne " & _
''            'rst1(3) & " a je vykonatelný dnem " & rst1(4) & " (dále jen exekuèní titul), byla povinnému uložena povinnost  zaplatit oprávnìnému:" & Chr(10)
''
''
''        End If
''    rst1.Close
''
''
''
''    '''  pocet uveru a jejich nazvy (FA nazvy zbytocne)
''    sql_he = "select distinct hu from ap_coll.vs_import_pm_all where typ = 'HU'  and epr = '" & sd.EPR(i) & "' and hu is not null order by hu" 'zjistime, kolik pro dane epr hu a jak se jmenuji
''    Set rst5 = New adodb.Recordset
''    rst5.Open sql_he, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''
''        cnt_loan = 0
''        'ReDim Loan(cnt_loan) 'FA zbytocne
''
''        If Not rst5.EOF Then
''            cnt_loan = rst5.RecordCount
''            ReDim Loan(cnt_loan)
''            For ij = 1 To cnt_loan
''                Loan(ij) = rst5![hu]
''                rst5.MoveNext
''            Next ij
''        End If
''
''    rst5.Close
''
''        '-------------------------------------------------------------hu------------
''    For ii = 1 To cnt_loan 'podle poctu hu budu dale vkladat intervaly uroceni
''
''
''        'XX FA odtialto zbytocne - v podstate format pisania ciastok
''        'sql100, citalo z v_exp_comp
'''            sql100 = "select comp from " & schema & ".v_exp_comp where hu = '" & Loan(ii) & "'" 'zjistime, zda psat castky dohromady nebo ne
'''
'''            Set rst1 = New ADODB.Recordset
'''            rst1.Open sql100, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''                If Not rst1.EOF Then
'''                    comp = rst1(0) 'pro dane hu zjistime, zda psat dohromady castky, pokud true, piseme dohromady
'''                End If
'''            rst1.Close
''        'XX fa potialto zbytocne
''
''        Set rst7 = New adodb.Recordset
''        sql7 = "select nvl(jistina,0) jistina, nvl(ku,0) ku, su_sazba, su_castka, su_dtime_valid_from, ou_sazba, ou_castka, ou_dtime_valid_from" & _
''                " from ap_coll.vs_import_pm_all where typ = 'HU' and hu = '" & Loan(ii) & "' and ou_dtime_valid_to > sysdate"
''        rst7.Open sql7, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic 'doplnime jistinu a k.urok
''            If Not rst7.EOF Then
''                amt_principal = rst7![jistina]
''                amt_interest = rst7![ku]
''            End If
''        rst7.Close
''
''        'FA prec nasl. par riadkov
''        'text1ku = ""
''        'If comp Then
''        '
''        '    text1 = text1 & Chr(10) & " -   èástku ve výši " & Format$(amt_principal + amt_interest, "#,##0.00") & " Kè, " & Chr(10)
''        '
''        'Else
''            text1 = text1 & Chr(10) & " -   èástku ve výši " & format$(amt_principal, "#,##0.00") & " Kè, " & Chr(10)
''            text1ku = " -   s kapitalizovaným úrokem ve výši " & format$(amt_interest, "#,##0.00") & " Kè, " & Chr(10)
''
''        'End If
''
''
''        sql66 = "select jistina, ku, su_sazba, su_castka, su_dtime_valid_from, ou_sazba, ou_castka, ou_dtime_valid_from, su_dtime_valid_to, ou_dtime_valid_to " & _
''                " from ap_coll.vs_import_pm_all where typ = 'HU' and hu = '" & Loan(ii) & "' order by interval"
''        Set rst1 = New adodb.Recordset
''        rst1.Open sql66, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic 'ke kazdemu hu ziskame udaje pocet intervalu
''            If Not rst1.EOF Then
''                cnt_int = rst1.RecordCount 'pocet intervalu k danemu hu
''                ReDim su_sazba(cnt_int)
''                ReDim su_castka(cnt_int)
''                ReDim su_dtime_valid_from(cnt_int)
''                ReDim ou_castka(cnt_int)
''                ReDim ou_sazba(cnt_int)
''                ReDim ou_dtime_valid_from(cnt_int)
''                ReDim ou_dtime_valid_to(cnt_int)
''                ReDim su_dtime_valid_to(cnt_int)
''                ReDim sankce(cnt_int)
''                ReDim ourok(cnt_int)
''
''
''                For qi = 1 To cnt_int
''
''                    If Not IsNull(rst1(2)) Then su_sazba(qi) = 100 * rst1![su_sazba] '(2)
''                    If Not IsNull(rst1(3)) Then su_castka(qi) = rst1![su_castka] '(3)
''                    If Not IsNull(rst1(4)) Then su_dtime_valid_from(qi) = rst1![su_dtime_valid_from] '(4)
''                    If Not IsNull(rst1(5)) Then ou_sazba(qi) = 100 * rst1![ou_sazba] '(5)
''                    If Not IsNull(rst1(6)) Then ou_castka(qi) = rst1![ou_castka] '(6)
''                    If Not IsNull(rst1(7)) Then ou_dtime_valid_from(qi) = rst1![ou_dtime_valid_from] '(7)
''                    If Not IsNull(rst1(8)) Then su_dtime_valid_to(qi) = rst1![su_dtime_valid_to] '(8)
''                    If Not IsNull(rst1(9)) Then ou_dtime_valid_to(qi) = rst1![ou_dtime_valid_to] '(9)
''
''                    rst1.MoveNext
''                Next qi
''
''            End If
''        rst1.Close
''
''
''
''        ''  FA odtialto vec uz sa bude generovat v zvlast funkcii podla formatu
''
''            'poskladame zneni zakonneho a sankcniho uroku
''
''
''            ourok(1) = " -   s úrokem ve výši " & ou_sazba(1) & " % roènì z èástky " & format$(ou_castka(1), "#,##0.00") & " Kè od " & ou_dtime_valid_from(1)
''            sankce(1) = " -   se zákonným úrokem z prodlení ve výši " & format$(su_sazba(1), "#,##0.00") & " % roènì z èástky " & format$(su_castka(1), "#,##0.00") & " Kè od " & su_dtime_valid_from(1)
''
''            For qi = 2 To cnt_int
''                sankce(qi - 1) = sankce(qi - 1) & " do " & su_dtime_valid_from(qi) - 1 & "," & Chr(10)
''                sankce(qi) = " -   se zákonným úrokem z prodlení ve výši " & su_sazba(qi) & " % roènì z èástky " & format$(su_castka(qi), "#,##0.00") & " Kè od " & su_dtime_valid_from(qi)
''                ourok(qi - 1) = ourok(qi - 1) & " do " & su_dtime_valid_from(qi) - 1 & "," & Chr(10)
''                ourok(qi) = " -   s úrokem ve výši " & ou_sazba(qi) & " % roènì z èástky " & format$(ou_castka(qi), "#,##0.00") & " Kè od " & ou_dtime_valid_from(qi)
''            Next qi
''
''            sankce(cnt_int) = sankce(cnt_int) & " do zaplacení," & Chr(10)
''            ourok(cnt_int) = ourok(cnt_int) & " do zaplacení," & Chr(10)
''
''
''            For qi = 1 To cnt_int
''                text1 = text1 & ourok(qi) & sankce(qi)
''            Next qi
''
''            text1 = text1 & text1ku
''            inc2 = inc 'definujeme zde, protoze ve vyberu oblasti to chceme az u dalsiho epr (v dalsim kole cyklu)
''
''        ''  FA potial
''
''    Next ii
''     '-------------------------------------------------------------HU-----
''
''
''
''    'na urovni celeho epr zmenime jistinu, k. urok a sop
''    'musime zjistit sop a debet
''
''    sql_re = "select nvl(sum(debet),0), nvl(sum(sop),0) from  ap_coll.vs_import_pm_all where epr = '" & sd.EPR(i) & "'"
''
''    debet = 0
''    SOP = 0
''    Set rst8 = New adodb.Recordset
''    rst8.Open sql_re, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''    If Not rst8.EOF Then
''        debet = rst8(0)
''        SOP = rst8(1)
''    End If
''    rst8.Close
''
''    debet_text = ""
''    If debet > 0 Then
''        debet_text = Chr(10) & textace("I_castka") & format$(debet, "#,##0.00") & " Kè,"
''
''        inc = inc + 1
''        inc2 = inc2 + 1
''    End If
''
''    sop_text = ""
''    If SOP > 0 Then
''        sop_text = vbCrLf & vbCrLf & textace("I_naklady_NR") & format$(SOP, "#,##0.00") & " Kè"
''        'sop_text = vbNewLine & " -   náklady nalézacího øízení ve výši " & Format$(sop, "#,##0.00") & " Kè"
''    End If
''
''    text1 = text1 & debet_text & sop_text & "."
''
''Next i
''
''GenerujTextCast_I = text1
''
''End Function
''
''
''
''
''
''
''
''
''
''
''Private Function GenerujTextCast_II(ByRef sd As data_pro_sablonu) As String
''
''Dim sql20 As String, platby As String, platby_rozpis() As String 'string je prilis dlouhy pro nahrazeni, proto udelame jako pole
''Dim cnt_platby As Integer, sql21 As String, amt_paid() As Double, date_p() As Date, platby_par() As String
''Dim text2 As String, flag_platby As Boolean, epr_all As String
''
''
'''''   musime zjistit, zda neco uhradil na nekterem z epr
''
''epr_all = "'" & sd.EPR(PoradoveCisloEPR) & "'"
''For i = 2 To sd.pocetZalob
''    epr_all = epr_all & ", '" & sd.EPR(i) & "'"
''Next i
''
''sql20 = "select date_payment, sum(amt_paid) as amt_paid from (select trunc(date_payment) date_payment, epr, amt_paid from ap_coll.pb_040_uhrady_pred_parovanim) where epr in (" & epr_all & ") group by date_payment order by date_payment asc"
''
''Set rst1 = New adodb.Recordset
''rst1.Open sql20, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''
''    If Not rst1.EOF Then
''
''        flag_platby = True 'pokud neco uhradil
''        rst1.MoveFirst
''        ReDim platby_rozpis(rst1.RecordCount)
''        ReDim amt_paid(rst1.RecordCount)
''        ReDim date_p(rst1.RecordCount)
''        ReDim platby_par(rst1.RecordCount)
''        cnt_platby = rst1.RecordCount
''        text2 = textace("II_uhradil") & vbCrLf & vbCrLf
''        For i = 1 To rst1.RecordCount
''
''            text2 = text2 & " -   dne " & format(rst1![date_payment], "Short Date") & " èástku ve výši " & format$(rst1![amt_paid], "#,##0.00") & " Kè," & vbCrLf
''            'FA som vyhodil nejaky zbytocny if
''            amt_paid(i) = rst1![amt_paid]
''            date_p(i) = format(rst1![date_payment], "Short Date")
''
''            rst1.MoveNext
''        Next i
''        text2 = text2 & vbCrLf & textace("II_castky_zapocet") & vbCrLf
''
''    Else
''
''        text2 = text2 & textace("II_nesplnil") '& vbCrLf & vbCrLf
''
''    End If
''rst1.Close
''
''
''
''
''Dim cnt_pl_par As Integer
''
'''ted zjistime, kam se naparovaly jednotlive platby
''If (flag_platby) Then
''    For i = 1 To cnt_platby 'pro kazdou platbu udelame zvlast dotaz
''
''        Set rst1 = New adodb.Recordset
''        sql21 = "select * from " & schema & ".v_exp_exe_par where epr in (" & epr_all & ") and date_payment = to_date('" & format(date_p(i), "Short Date") & "') order by kategorie_por"
''        rst1.Open sql21, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''            If Not rst1.EOF Then
''
''                cnt_pl_par = rst1.RecordCount
''                rst1.MoveFirst
''
''                text2 = text2 & vbCrLf
''                text2 = text2 & "- èástku ve výši " & format$(amt_paid(i), "#,##0.00") & " Kè ze dne " & date_p(i) & " zapoèetl oprávnìný èástí ve výši "
''
''                For ii = 1 To cnt_pl_par
''
''                    text2 = text2 & format$(rst1![amt_paid], "#,##0.00") & " Kè na úhradu " & rst1![kategorie] & ", "
''                    rst1.MoveNext
''
''                Next ii
''
''                'text2 = text2 & vbCrLf
''
''            End If
''        rst1.Close
''    Next i
''End If
''
''GenerujTextCast_II = text2
''
''
'''''   Pokud ma klient nejaky pausal, tak ten NEvymahame a do cl.II se vlozi veta navic
''
''
''If c.SOP.Pausal > 0 Then
''    'GenerujTextCast_II = GenerujTextCast_II & vbCrLf & "Blah blah blah Paušál ve výši " & c.SOP.Pausal & " blah blah."
''    GenerujTextCast_II = GenerujTextCast_II & vbCrLf & vbCrLf & textace("II_pouze_cast") '& _
''                        '" (XXXXXX Jezevèík tady mùže uvést i èástku: " & c.SOP.Pausal & ".)" '& vbCrLf
''
''End If
''
''
''
'''''''      FA v povodnom kode sa vsetky dalsie hodnoty naplnili ale uz sa nikam nevlozili
''
'''''   musime nyni zjistit, zda vymahame mene (zda pridat treti variantu), to se ale poji pouze s 1. variantou
''Dim sql22 As String, flag_stejne As Boolean
''
''If Not flag_platby Then
''
''    sql22 = "select min(stejne) from " & schema & ".v_exp_bank_claim where epr in (" & epr_all & ") group by epr"
''
''    Set rst1 = New adodb.Recordset
''    rst1.Open sql22, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''        If Not rst1.EOF Then
''
''            flag_stejne = rst1(0)
''
''        End If
''    rst1.Close
''End If
''
''
''If Not flag_stejne And Not flag_platby Then
''    'XX FA u klientky CUID 443414 (Pong  ) se textace   generuje, ale nemela by
''    'XX FA u klienta  CUID 224469 (Melich) se textace negeneruje, ale   mela by
''    ' obidva pripady su zle kvoli podmienke flag_stejne (opacna hodnota)
''    text2 = text2 & Chr(10) & Chr(10) & textace("II_pouze_cast")
''End If
''
''
''End Function
''
''
''
''
''
''Private Function GenerujTextPovereni_V(ByRef sd As data_pro_sablonu) As String
''' vystupem funkce je textovy retezec (ktery se vlozi do Word Dokumentu do casti V - Povereni)
''' v ramci funkce je hlavnim textakem promenna text5
''
''
''Dim cnt_loan As Integer, cnt_loan_akt As Integer, sql12 As String, amt_principal_akt As Double, amt_interest_akt As Double, loan_akt() As String
''Dim sql13 As String, sql11 As String, sql14 As String, debet_akt As Double, sop_akt As Double, sql15 As String, sql16 As String, sql17 As String, sql18 As String, sql100 As String
''Dim cnt_sankce As Integer, cnt_urok As Integer, sankce_akt As String, urok_akt As String, ainc As Integer, ainc2 As Integer, ainc_h As Integer
''Dim sql00 As String, comp As Boolean, text5 As String, text5su As String, text5ou As String, text5ku As String
''Dim typ_roz As String
''
''
''
''    For i = 1 To sd.pocetZalob 'pro kazde epr doplnime info z rozsudku do hlavicky
''
''        If i = 1 Then
''            text5 = text5 & "Podle "
''        Else
''            text5 = text5 & "Dále podle "
''        End If
''
''        text5 = text5 & textace("Povereni_uvod") & sd.klient_udaje
''
''        sql11 = "select soud, epr, datum_vydani, datum_pm, datum_vykon, typ from ap_coll.vs_import_pm_rozsudek_all where epr = '" & sd.EPR(i) & "'" 'tabulka se plni na ap_coll hristi
''        Set rst1 = New adodb.Recordset
''
''        rst1.Open sql11, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''            If Not rst1.EOF Then
''
''                typ_roz = "" & rst1(5) & ""
''
''                If typ_roz = "rozsudek" Then
''                    typ_roz = "rozsudku"
''                    text5 = text5 & ", dle " & typ_roz & " è.j. " & rst1(1) & ", který vydal " & rst1(0) & " dne " & rst1(2) & ", který nabyl právní moci dne " & rst1(3) & " a" & Chr(160) & "je vykonatelný dnem " & rst1(4) & ", uhradit oprávnìnému:" & Chr(10)
''
''                ElseIf typ_roz = "elektronický platební rozkaz" Then
''                    typ_roz = "elektronického platebního rozkazu"
''                    text5 = text5 & ", dle " & typ_roz & " è.j. " & rst1(1) & ", který vydal " & rst1(0) & " dne " & rst1(2) & ", který nabyl právní moci dne " & rst1(3) & " a" & Chr(160) & "je vykonatelný, uhradit oprávnìnému:" & Chr(10)
''
''                ''  nasledujuce 2 su iba kvoli preklepom - napr. klient cuid 619818 mal "elekronický platební rozkaz" (chyba tam T)
''                ''  pridane ako ElseIf na koniec, pretoze ak recordset nie je prazdny tak by tam mala byt jedna z prvych dvoch moznosti.. tieto sa testuju iba ak dojde k failu zjavnemu
''                ElseIf InStr(typ_roz, "rozs") Then
''                    typ_roz = "rozsudku"
''                    text5 = text5 & ", dle " & typ_roz & " è.j. " & rst1(1) & ", který vydal " & rst1(0) & " dne " & rst1(2) & ", který nabyl právní moci dne " & rst1(3) & " a" & Chr(160) & "je vykonatelný dnem " & rst1(4) & ", uhradit oprávnìnému:" & Chr(10)
''
''                ElseIf InStr(typ_roz, "elek") Then
''                    typ_roz = "elektronického platebního rozkazu"
''                    text5 = text5 & ", dle " & typ_roz & " è.j. " & rst1(1) & ", který vydal " & rst1(0) & " dne " & rst1(2) & ", který nabyl právní moci dne " & rst1(3) & " a" & Chr(160) & "je vykonatelný, uhradit oprávnìnému:" & Chr(10)
''
''                Else
''                    HlaskaUzivateli "V petitu chybí èást textu k EPR/rozsudku. Nejspíš pøeklep v datech."
''                End If
''
''            End If
''        rst1.Close
''
''
''
''
''        sql12 = "select distinct typ_id from ap_coll.pb_040_pohledavka_unpaid where typ like 'HU%' and epr = '" & sd.EPR(i) & "' order by typ_id" 'pro kazde epr zjistime pocet hu
''
''        Set rst1 = New adodb.Recordset
''
''        rst1.Open sql12, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''        If Not rst1.EOF Then
''            cnt_loan_akt = rst1.RecordCount
''            ReDim loan_akt(cnt_loan_akt)
''
''            rst1.MoveFirst
''            For ii = 1 To cnt_loan_akt
''                loan_akt(ii) = rst1(0)
''                rst1.MoveNext
''            Next ii
''
''        End If
''        rst1.Close
''
''    '    'pokud nema nedoplacenou zadnou cast pohledavky, tuto cast kodu preskocime
''    '    If IsNull(cnt_loan_akt) Then
''    '        MsgBox "Pohledávka je doplacená"
''    '        GoTo skip_pm_akt
''    '    End If
''
''        If (IsNull(cnt_loan_akt) Or cnt_loan_akt = 0) And InStr(strMsgbox, "má doplacenou jistinu a kap. úrok") = 0 Then
''            ' posledna podmienka je aby sa hlaska nevypisovala duplicitne
''            HlaskaUzivateli "Klient má doplacenou jistinu a kap. úrok u EPR: " & sd.EPR(PoradoveCisloEPR) & _
''            ". Je potøeba provést kontrolu, zda nemá nedoplacené obchodní nebo sankèní úroky. "
''        End If
''
''
''        '---------------------hu
''        For ii = 1 To cnt_loan_akt
''
''            'FA XX toto asi zbytocne odtial
''                'musime zjistit, zda psat jistinu a kap. urok dohormady. Pravidlo je, ze pokud obch a sankcni urok se uroci ze stejne castky, pise se to zvlast, jinak dohromady (stare zaloby)
''                sql100 = "select comp from " & schema & ".v_exp_comp where hu = '" & loan_akt(ii) & "'"
''
''                Set rst1 = New adodb.Recordset
''                rst1.Open sql100, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''                    If Not rst1.EOF Then
''                        comp = rst1(0) 'pro dane hu zjistime, zda psat dohromady castky, pokud true, piseme dohromady
''                    End If
''                rst1.Close
''            'FA XX potial
''
''            sql13 = "select sum(amt) from ap_coll.pb_040_pohledavka_unpaid where typ_id = '" & loan_akt(ii) & "' and typ like 'HU - Úrok%'"
''
''            Set rst1 = New adodb.Recordset
''            rst1.Open sql13, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''                If Not rst1.EOF Then
''                    If IsNull(rst1(0)) Then
''                        amt_interest_akt = 0
''                    Else
''                        amt_interest_akt = rst1(0) 'pro dane epr pocitame vysi uroku
''                    End If
''                End If
''            rst1.Close
''
''            Set rst1 = New adodb.Recordset
''            sql14 = "select sum(amt) from ap_coll.pb_040_pohledavka_unpaid where typ_id = '" & loan_akt(ii) & "' and typ like 'HU - Jis%'"
''
''            rst1.Open sql14, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''                If Not rst1.EOF Then
''                    If IsNull(rst1(0)) Then
''                        amt_principal_akt = 0 'pro dane epr pocitame vysi jistiny
''                    Else
''                        amt_principal_akt = rst1(0) 'pro dane epr pocitame vysi uroku
''                    End If
''                End If
''            rst1.Close
''
''
''
'''            If comp Then ' NESCITAT
'''                text5 = text5 & vbCrLf & " -   èástku ve výši " & format$(amt_principal_akt + amt_interest_akt, "#,##0.00") & " Kè, " & Chr(10)
'''            Else
''''                text5 = text5 & vbCrLf & " -   èástku ve výši " & format$((amt_principal_akt), "#,##0.00") & " Kè, " & Chr(10)
''''                text5ku = " -   s kapitalizovaným úrokem ve výši " & format$(amt_interest_akt, "#,##0.00") & " Kè, " & Chr(10)
''                text5 = text5 & vbCrLf & " -   èástku ve výši " & VlozCastku(amt_principal_akt) & ", " & Chr(10)
''                text5ku = " -   s kapitalizovaným úrokem ve výši " & VlozCastku(amt_interest_akt) & ", " & Chr(10)
'''            End If
''
''            'sankcni a obchodni urok
''            ' FA: tabulka v nasledujucom selecte vychadza z reportov nejak, sileny skript ktory Vlastik iba prebral od Peta
''            sql17 = "select * from ap_coll.pb_040_intervaly where code_loan_number = '" & loan_akt(ii) & "'  order by  dtime_valid_from, typ asc" 'tabulka zustava na ap_coll
''            Set rst2 = New adodb.Recordset
''            rst2.Open sql17, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''                If Not rst2.EOF Then
''                    cnt_sankce = rst2.RecordCount
''                    rst2.MoveFirst
''
''                    For qi = 1 To cnt_sankce
''
''                        If qi < cnt_sankce - 1 Then
''                            If rst2(5) = "ourok" Then
''''                                text5 = text5 & " -   spolu s úrokem ve výši " & rst2(4) * 100 & " % roènì z èástky " & format$(rst2(1), "#,##0.00") & " Kè od " & rst2(2) & " do " & rst2(3) - 1 & "," & Chr(10)
''                                text5 = text5 & " -   spolu s úrokem ve výši " & rst2(4) * 100 & " % roènì z èástky " & VlozCastku(rst2(1)) & " od " & rst2(2) & " do " & rst2(3) - 1 & "," & Chr(10)
''                            Else
''                                text5 = text5 & " -   se zákonným úrokem z prodlení ve výši " & rst2(4) * 100 & " % roènì z èástky " & VlozCastku(rst2(1)) & " od " & rst2(2) & " do " & rst2(3) - 1 & "," & Chr(10)
''
''                            End If
''
''                        Else
''
''                            If rst2(5) = "ourok" Then
''                                text5 = text5 & " -   spolu s úrokem ve výši " & rst2(4) * 100 & " % roènì z èástky " & VlozCastku(rst2(1)) & " od " & rst2(2) & " do zaplacení," & Chr(10)
''                            Else
''                                text5 = text5 & " -   se zákonným úrokem z prodlení ve výši " & rst2(4) * 100 & " % roènì z èástky " & VlozCastku(rst2(1)) & " od " & rst2(2) & " do zaplacení," & Chr(10)
''
''                            End If
''
''                         End If
''                        'sankce_akt = sankce_akt & rst2(2) - 1 & ", dále úrok ve výši " & rst2(4) * 100 & " % z èástky" & Format(rst2(1), "### ### ##0.00") & " od " & rst2(2) & " do "
''                        rst2.MoveNext
''
''                    Next qi
''                    'text5ou = text5ou & "zaplacení," & Chr(10)
''
''                End If
''            rst2.Close
''            text5 = text5 & text5ku
''
''        Next ii
''
''
''        '''   dale zjistime pro kazde epr debet a sop
''
''        'ainc2 = ainc
''        Set rst1 = New adodb.Recordset
''        sql15 = "select sum(amt) from ap_coll.pb_040_pohledavka_unpaid where typ = 'DEBET' and epr = '" & sd.EPR(i) & "'" 'pro kazde epr zjistime debet
''
''        rst1.Open sql15, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''            If Not rst1.EOF Then
''                If Not IsNull(rst1(0)) Then
''                    debet_akt = rst1(0)
''                Else
''                    debet_akt = 0
''                End If
''            End If
''        rst1.Close
''
''        Set rst1 = New adodb.Recordset
''        sql16 = "select sum(amt) from ap_coll.pb_040_pohledavka_unpaid where typ = 'SOP' and epr = '" & sd.EPR(i) & "'" 'pro kazde epr zjistime sop
''
''        rst1.Open sql16, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''            If Not rst1.EOF Then
''                If Not IsNull(rst1(0)) Then
''                    sop_akt = rst1(0)
''                Else
''                    sop_akt = 0
''                End If
''            End If
''        rst1.Close
''
''
''        If debet_akt > 0 Then
''            text5 = text5 & vbCrLf & " -   èástku ve výši " & VlozCastku(debet_akt) & "," & vbCrLf
''        End If
''
''        If sop_akt > 0 Then
''            'text5 = text5 & vbCrLf & " -   náklady nalézacího øízení ve výši " & format$(sop_akt, "#,##0.00") & " Kè" & Chr(10)\
''            text5 = text5 & vbCrLf & " -   náklady nalézacího øízení ve výši " & VlozCastku(sop_akt - c.SOP.Pausal) & Chr(10)
''            If (sop_akt - c.SOP.Pausal) < 0 Then HlaskaUzivateli "V èásti V je v nákladech øízení záporná èástka. Jezevèík nechápe."
''        End If
''
''
''    Next i
''
''
''    '------------------------epr
''
''GenerujTextPovereni_V = text5
''
''End Function
''
''
'
'
'
'
'
'
'
'
'
'
'
'
'''Private Function VytvorTextCast_II_ehh(Pohl_Zaplacene As Pohl_zaplacena, sd As data_pro_sablonu) As String
'''''  tuto mozes vyhodit ak funguje verzia bez "ehh" v nazve
'''
'''' ?? meni sa format podla rozsudku??
'''' podla coho pismeno a EPR v casti 2 na konci zistim?
'''
'''Dim text2 As String
'''Dim dblPausal As Double
'''
'''Dim sql As String
'''Dim rst As ADODB.Recordset
'''Dim i As Integer
'''Dim ii As Integer
'''
'''
'''For i = 1 To sd.PocetPlatebCast_II 'pro kazdou platbu udelame zvlast dotaz
'''
'''    Set rst = New ADODB.Recordset
'''    sql = "select * from " & schema & ".v_exp_exe_par where epr in (" & sd.epr_all & ") and date_payment = to_date('" & format(sd.datumyZaplacenePohledavky(i), "Short Date") & "') order by kategorie_por"
'''    rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''        If Not rst.EOF Then
'''
'''
'''            rst.MoveFirst
'''
'''            text2 = text2 & textace("radek") & textace("-") & textace("I_castka") & format$(sd.castkyZaplacenePohledavky(i), "#,##0.00") & _
'''                    " ze dne " & sd.datumyZaplacenePohledavky(i) & " zapoèetl oprávnìný" ' èástí ve výši "
'''
'''            If rst.RecordCount = 1 Then
'''                text2 = text2 & " na " & rst![kategorie] & textace("II_pohl_pism") '& "" & sd.EPR ' FAXX podla coho zistim pismeno? a ze je to EPR? cislo pohladavky? (detto o 3 riadky nizsie)
'''            Else
'''                For ii = 1 To rst.RecordCount
'''
'''                    text2 = text2 & textace("radek") & textace("II_cast") & VlozCastku(rst![amt_paid]) & " na " & rst![kategorie] & textace("II_pohl_pism") '& "" & sd.EPR
'''
'''                    ' povodna verzia:
'''                        'text2 = text2 & VlozCastku(rst![amt_paid]) & " na úhradu " & rst![kategorie]
'''                        '' povotext2 = text2 & format$(rst![amt_paid], "#,##0.00") & " Kè na úhradu " & rst![kategorie] & ", "
'''
'''                    rst.MoveNext
'''                Next ii
'''            End If
'''
'''        End If
'''    rst.Close
'''Next i
'''
'''
'''
'''
''''''   Pokud ma klient nejaky pausal, tak ten NEvymahame a do cl.II se vlozi veta navic
'''
'''dblPausal = PausalVyska(c.id_person)
'''
'''If dblPausal > 0 Then
'''    text2 = text2 & textace("radek") & textace("radek") & textace("II_pouze_cast") '& _
'''                '" (XXXXXX Jezevèík tady mùže uvést i èástku: " & dblPausal & ".)" '& textace("radek")
'''
'''End If
'''
'''
'''VytvorTextCast_II = text2
'''
'''
''''''''      FA v povodnom kode tu este bol kusok (podla mna) zbytocneho kodu
'''
'''
'''
'''End Function
'
'
'
'
'
