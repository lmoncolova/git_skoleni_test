Attribute VB_Name = "mod_Export_IP_XML"
Option Explicit

Public Const sngZakonnyUrok = 8 ' + REPO sazba vzdy a vsade sa muzi pripocitat
Public strDatumPosledniDebetniTrn As String 'public, pretoze sa plni v epr_xml



Private Ins As Insolvency 'XX este overenie nejake ci uz neexistuje objekt, netusim ako sa to spravat bude. Nieco je napisane tu: http://www.cpearson.com/excel/classes.aspx
Private elemIP As Integer
Private strCisloUctu As String, text As String, StrCisloJednaci As String, text_vykonatelnost_dle As String
Private objSelection As Variant
Private sql_report As String, rst_report As ADODB.Recordset, strReportChybiciUdaje As String, strReportChybiciUvery As String
Private dblDebet As Double
'private dblSop As Double'vyhazuji jez.v.5.8
Private blnReport As Boolean, blnReportDebet As Boolean, blnReportSop As Boolean, blnEpr As Boolean
Private strSoudReport As String, strDatumVydani As String, strDatumPm As String, strDatumVykon As String, strReportBankAccountNumber As String
' Private blnPlatbyPoZesplatneni As Boolean ' FA vyhazuji do verze jez.v.5.7. Souvisi s napoctem AP_COLL.PB_040_SUMMARY, a RISK-849

'Private XDoc As Object 'XDoc je PUBLIC v module EPR_XML, odtial sa vyuzivaju funkcie naplnElement a pod. Pre oddelenie modulov by trebalo pridat argument XDOC do funkcie pridajElement.
'Private cL As Long
'Private cA As Long






''''''     MAIN BODY






Public Sub generovatIP_XML()
'Sub generovatIP_XML(id_person As Long)'FA toto chceme pouzit

'Logika kodu pro generovani prihlasky do insolvencniho rizeni je nasledovna:
'nejdrive se inicializuje objekt Insolvency do promenne Ins.
'pak se postupne naplni cast dluznik, pohledavky k uverum (jeden nebo vice), pak pohledavka k debetu (pokud nejaky je) a sopu (pokud nejaky je) - pro kazdou kategorii je zvlast Sub()
'vzdy se nejdrive plni promenne, vkladani do XML elementu probiha az ke konci kazde Sub()
'mnoho veci v sablone zavisi od toho, ci mame ku klientovi PM rozhodnuti, a ci je klient evidovany v reporte PM pohledavek
Dim root As Object


Call VynulujPremenne

Set XDoc = CreateObject("MSXML2.DOMDocument")
XDoc.async = False: XDoc.validateOnParse = False

'Naètení xml s elementy EPR s pøedvyplnìnými konstantními údaji (žalobce = AB, mail:michaelstefl@airbank.cz, úèet pro vraceni sop 30300389/3030,...)
XDoc.Load (cfg.Cesta("vzorInsXMLBase"))
Set root = XDoc.DocumentElement



'''''      inicializace objektu insolvence


'XX presunute na uroven Modulu: Dim Ins As Insolvency 'XX tu by to chcelo asi nejak overit ci uz taky objekt nemame, netusim ako sa to spravat bude. Nieco je napisane tu: http://www.cpearson.com/excel/classes.aspx
Set Ins = New Insolvency
Call Ins.initialize(Val(c.id_person))



'''''      NAPLNÌNÍ ÈÁSTI Dlužník


'pokud je trvalá adresa v èr, tak trvalé, jinak korespondenèní v èr, pokud by ani jedna nebyla v èr, zùstane prázdné a vyskoèí upozornìní
Call NaplnCastDluznik



'''''      Prihlasky pohledavek. HU -> debet -> SOP
'Call ZkontrolujPrichoziPlatby(Val(c.CUID)) 'musi jit pred prihlaskaPohledavky_uver kvuli promenne blnPlatbyPoZesplatneni
Call prihlaskaPohledavky_uver(Val(c.id_person))

If Ins.blnPm Then
    Call nacteniDebetuASopuZReportu(Val(c.id_person))
End If
If text_vykonatelnost_dle = "" Then
    NaplnTextVykonatelnost
End If

Call prihlaskaPohledavky_debet(Val(c.id_person)) 'az po nacteniDebetuASopuZReportu()

If Ins.blnPm Then 'SOP prihlasujeme pouze pokud mame PM
    Call prihlaskaPohledavky_sop(Val(c.id_person)) 'az po nacteniDebetuASopuZReportu()
End If


If Ins.blnPmZamitnut Then HlaskaUzivateli "PM Rozsudek neuznal všechny pohledávky -> intervaly k úvìrùm klienta byli naplnìné podle reportu PM pohl (ne podle OBS)."
'If Not Ins.blnPmZamitnut Then HlaskaUzivateli "PM Rozsudek uznal všechny pohledávky -> intervaly k úvìrùm klienta byli naplnìné podle OBS (ne podle reportu PM pohl)."
Call VypisVsechnyHlasky


Call UlozSoubor(cfg.Cesta("InsPripravene") & c.Nazev_slozky_bez_lomitka, "BaseXmlIp_" & c.Nazev_slozky_bez_lomitka & ".xml", cfg.Cesta("InsPripravene"))

Call VypisVsechnyHlasky

End Sub


Private Sub NaplnCastDluznik()
Dim cilElem As String

cilElem = "//d:root/dluznik[0]/fyzicka_osoba[0]/"
Call naplnElement(cilElem, "osoba[0]/jmeno", c.Name_first)
Call naplnElement(cilElem, "osoba[0]/prijmeni", c.Name_last)
Call naplnElement(cilElem, "osoba[0]/titul_pred", c.Text_Title_before)
Call naplnElement(cilElem, "osoba[0]/titul_za", c.Text_Title_After)

If c.PinLom <> "" Then
    Call naplnElement(cilElem, "rodne_cislo", c.PinLom)
End If

Call naplnElement(cilElem, "datum_narozeni", format(c.Date_Birth, "yyyy-mm-dd"))

'pokud je trvalá adresa v èr, tak trvalé, jinak korespondenèní v èr, pokud by ani jedna nebyla v èr, zùstane prázdné
If c.Address_P_Zeme = "Czech Republic" Then
   
    Call naplnElement(cilElem, "adresa[0]/ulice", c.Address_P_Ulice)
    
    If InStr(c.Address_P_CP, "/") = 0 Then 'pokud je v èp lomítko, vyplòuje se zvláš cp a ce
        Call naplnElement(cilElem, "adresa[0]/c_popisne", c.Address_P_CP)
    Else
        Call naplnElement(cilElem, "adresa[0]/c_popisne", Left(c.Address_P_CP, InStr(c.Address_P_CP, "/") - 1))
        Call naplnElement(cilElem, "adresa[0]/c_orientacni", Mid(c.Address_P_CP, InStr(c.Address_P_CP, "/") + 1))
    End If
    
    Call naplnElement(cilElem, "adresa[0]/obec", c.Address_P_Mesto)
    Call naplnElement(cilElem, "adresa[0]/psc", c.Address_P_PSC)
    Call naplnElement(cilElem, "adresa[0]/stat", "CZ")
Else
    If c.Address_C_Zeme = "Czech Republic" Then 'Pokud trvalá není CZ a kontaktní je CZ, doplní se kontaktní
        Call naplnElement(cilElem, "adresa[0]/ulice", c.Address_C_Ulice)
        
        If InStr(c.Address_C_CP, "/") = 0 Then 'pokud je v èp lomítko, vyplòuje se zvláš cp a ce
            Call naplnElement(cilElem, "adresa[0]/c_popisne", c.Address_C_CP)
        Else
            Call naplnElement(cilElem, "adresa[0]/c_popisne", Left(c.Address_C_CP, InStr(c.Address_C_CP, "/") - 1))
            Call naplnElement(cilElem, "adresa[0]/c_orientacni", Mid(c.Address_C_CP, InStr(c.Address_C_CP, "/") + 1))
        End If
        
        Call naplnElement(cilElem, "adresa[0]/obec", c.Address_C_Mesto)
        Call naplnElement(cilElem, "adresa[0]/psc", c.Address_C_PSC)
        Call naplnElement(cilElem, "adresa[0]/stat", "CZ")
    'pokud ani trvalá ani kontaktní není CZ, zùstane prázdné a vyskoèí upozornìní
    Else
        HlaskaUzivateli "Trvalá ani koresponenèní adresa není v ÈR, prosím doplò ruènì adresu žalovaného do formuláøe."
    End If
    
End If
Call naplnElement("//d:root/podpisem_stvrzuje/", "dne", VlozDatumDoXML(date))




'''''     soud + sp. znacka


Call naplnElement("//d:root/", "soud", Ins.strKodInsSoud)
Call naplnElement("//d:root/", "bc", Mid(Ins.strInsCode, _
                                         InStr(1, Ins.strInsCode, " ") + 1, _
                                         InStr(1, Ins.strInsCode, "/") - InStr(1, Ins.strInsCode, " ") - 1))
Call naplnElement("//d:root/", "rocnik", Right(Ins.strInsCode, 4))



'''''     cislo uctu klienta


strCisloUctu = Ins.strVu & "/3030" & " " & "VS: " & c.TextContractNumber
'Debug.Print strCisloUctu
Call naplnElement("//d:root/veritel[0]/pravnicka_osoba[0]/adresa[0]/", "c_uctu", strCisloUctu)


End Sub



Private Sub prihlaskaPohledavky_uver(ByVal id_person As Variant) ' variant?


Dim strHu As String, strOu_dtime_valid_from As String, strOu_dtime_valid_to As String, strSu_dtime_valid_from As String, strSu_dtime_valid_to As String
Dim dblCelkem As Double, dblJistina As Double, dblKu As Double, dblOu As Double, dblSu As Double, dblOu_sazba As Double, dblOu_castka As Double, dblSu_sazba As Double, dblSu_castka As Double
Dim dblJistinaKeDniZaloby As Double, dblKapUrokKeDniZaloby As Double, dblKapUrok As Double, dblSplatnaCastka As Double

Dim dtmPodpis As Date, dtmCerpani As Date, cLn As String, strDodatek As String ', amt_loan_limit As Double, ucet As String
'Dim rate_interest As single 'rate_interest = 14.5 'takto sa to plnilo nizsie v kode,je tam pri tom poznamka nejaka FA_XX --ide o smluvni urok, z niekade ho dotiahnut treba.

Dim sql_inter As String, cnt_int As Integer, ii As Integer, amt_uroky As Double
Dim rst_inter As ADODB.Recordset
Dim text_pohl As String, text_pohl_uvod As String

Dim strUrokDruh As String
Dim intInterval As Integer


'FA NIEKTORE ZMENY PREMENNYCH oproti povodnej sablone:
'pm(bool)              -> Ins.boolPm
'Dim kr_soud As String -> Ins.strKrajSoud  -> c.soud.krajsky_soud
'okr_soud As String    -> Ins.strOkresSoud -> c.soud.okresni_soud
'sp As String          -> Ins.strInsCode
'amt_sop               -> Ins.dblAmtSop
'                      -> Ins.strKodInsSoud
'vu                    -> ins.strVu
'code_epr              -> Ins.strSpisZnacka
'amt_jis_akt           -> dblJistina
'amt_jis               -> dblJistinaKeDniZaloby
'amt_kap_urok          -> dblKapUrok




'''''    info o pohledavce - UVER


'  zjistime vsechny loany

If c.GetCnt_Loans > 0 Then

    For cL = 1 To c.GetCnt_Loans
        'XX tady vybrat jen požadované úvìry - ovìøit s PV jaké - zøejmì zesplatnìné + k odpisu + odepsané

        If c.GetLoans(cL).Loan_status = "Aktivní" And _
            (c.GetLoans(cL).Loan_wfstatus = "Zesplatnìný" Or c.GetLoans(cL).Loan_wfstatus = "K odpisu" Or c.GetLoans(cL).Loan_wfstatus = "Odepsaný bankou") Then
            
            elemIP = elemIP + 1
            
            'podle dodatku zjistime, zda s ramcovou smlouvu nebo ne
                   
            'cislo uveru, dodatek, ucet
            cLn = c.GetLoans(cL).code_loan_number
            strDodatek = c.GetLoans(cL).Text_appendix_number
            dtmCerpani = "" & c.GetLoans(cL).Date_loan_utilization & ""
            dtmPodpis = c.GetLoans(cL).Dtime_signature
            strUrokDruh = "Smluvní úrok"
            
            
            'cely sql_report bude treba nejak zapodmienkovat, teraz to bude hladat pri kazdom uvere..
            'obmedzene na uvery (debet a SOP zvlast nizsie su. Tiez cez loop, SOP sa moze opakovat urcite.)
            'sql prikaz je dlhsi, lebo v nom pocitame ciastky k intervalom a dotahuje sa aj datum upadku extra kvoli tomu.
            
            'if blnreport then
                sql_report = "select * from " & schema & ".V_INS_INTERVALY_REPORT a " & _
                             "WHERE a.cuid = " & c.CUID & " AND a.epr LIKE '%" & Ins.strSpisZnacka & "%' AND a.HU = '" & cLn & "' " & _
                             "ORDER BY a.interval ASC"

                
                Set rst_report = New ADODB.Recordset
                rst_report.Open sql_report, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
            
                    If Not rst_report.EOF Then
                        
                        '   logika kodu: najprv ulozime premenne ktore su rovnake pre vsetky riadky, potom nizsie cyklus cez ostatne a zrataju sa intervaly
                        
                        rst_report.MoveFirst
                        'Debug.Print rst_report![EPR]
                        blnReport = True
                        cnt_int = rst_report.RecordCount
                        
                        
                        '   nacitanie dat do premennych z reportu PM pohledavek z sheetu Souhrn.
                        If IsNull(rst_report![EPR]) And IsNull(rst_report![SOUD]) And IsNull(rst_report![datum_vydani]) And IsNull(rst_report![datum_pm]) And IsNull(rst_report![datum_vykon]) Then
                            'Ak sú všetky NULLs, spojím chybové hlášky do jednej.
                            HlaskaUzivateli "Z reportu se nenahrál soud, èíslo jednací, datum vydání, datum pm ani datum výkonu. Je potøeba doplnit ruènì. (" & c.GetLoans(cL).code_loan_number & ")."
                        Else
                        
                            NactiPromennu StrCisloJednaci, rst_report![EPR], "èj"
                            NactiPromennu strSoudReport, rst_report![SOUD], "soud"
                            NactiPromennu strDatumVydani, rst_report![datum_vydani], "datum vydání"
                            NactiPromennu strDatumPm, rst_report![datum_pm], "datum PM"
                            NactiPromennu strDatumVykon, rst_report![datum_vykon], "datum výkonu"
                            
                        End If
                        
                        
                        '   upozorneni pokud je v reportu niektory OPRAVENY datum (datum_opr_vydani, datum_opr_pm, datum_opr_vykon)
                        
                        If Not (IsNull(rst_report![datum_opr_vydani]) And IsNull(rst_report![datum_opr_pm]) And IsNull(rst_report![datum_opr_vykon])) Then
                            HlaskaUzivateli "Klient má v reportu PM pohledávek vyplnìné nìjaké opravené datum. Jezevèík jej ignoruje, bohužel to musíš doplnit ruènì."
                        End If
                        
                        
                        '   naplnenie obchodnych a sankcnych urokov do premennych z ktorych sa nizsie plni XML.
                        ' do XML sa vkladaju az ovela nizsie v casti 9.
                        
                        amt_uroky = 0
                        
                        
                        
                        '''''    INTERVALY
                        ' v tento moment su naprogramovane 3 moznosti plnenia intervalov, pricom moznost 2 a 3 robia vacsinou to iste:
                        '   1. Podla Reportu PM Pohledavek = ap_coll.vs_import_pm_all                 (problem je ze sa ignoruju platby po zesplatneni)
                        '   2. Podla OBS / DWH skrz tabulku app_coll_jezevcik(test).intervaly         (problem je ze sa ignoruje PM Rozsudek A DALSI problem ze napocitava az do dneska (spravne by malo iba do datumu zesplatneni), tiez napr. cuid 600794 sa objavil uplne odveci datum u posledneho intervalu.
                        '   3. Podla OBS / DWH skrz tabulku ap_coll.sr_coll_014_intervaly             (problem je ze sa ignoruje PM Rozsudek)
                        '  (4. skrz objekty Interest v Jezevcikovi, e.g. v XML pro EPR search: c.GetLoans(cL).GetInterests(cI).AmtUrok)
                        '   5. Podla OBS / DWH skrz view V_INS_INTERVALY                              - diva sa na intervaly v JEZ, ale ukoncuje podla datumu zesplatneni. Menej chyb nez predosle...
                        '
                        '
                        ' napocet ktory by to bral vsetko do uvahy este nie je. Je potrebne si uvedomit ze PM Rozsudek nam nieco zamietne iba v MALOM pocte pripadov a tie su lahko identifikovatelne
                        ' podla stavu (PM rozsudek, PM rozsudek - castecne zamitnuto, a pod.)
                        '
                        ' v tento moment su tu ponechane vsetky kody, aby eventuelne u problemoveho klienta boli vsetky moznosti k dispozicii.
                        ' Sablona sa pozrie na evidovane stavy u klienta, ak mal niekedy stav "PM..zamitnut..", tak sa vygeneruju intervaly podla reportu.
                        ' V opacnom pripade (teda vacsinou) sa vygeneruju moznostou cislo 3.
                        '
                        '
                        '
                        ' update od Jez.v.5.3  -  ak zafunguje, tak predosle napocty uz nebudu potrebne  -  nove view V_INS_INTERVALY na hristi Jez
                        '                      -  pouziva sa v pripade, ak nam PM rozsudek nic nezamietol
                        '
                        '
                        '
                        '''''
                        
                        
                        
                        '''  tu sa naplnia intervaly JEDINE v pripade, ze nam PM Rozsudek nieco zamietol
                        If (Ins.blnPmZamitnut) Then
                        'If (Ins.blnPmZamitnut) Or (Not blnPlatbyPoZesplatneni) Then
                        
                            For ii = 1 To cnt_int
                                
                                'dblCelkem = Round(rst_report![celkem], 2) 'nevyuziva se nikde
                                dblJistina = Round(rst_report![JISTINA], 2) 'neskor este dblJistinaKeDniZaloby sa bude pouzivat
                                dblKu = Round(rst_report![ku], 2)
                                dblOu = Round(rst_report![ou], 2)
                                dblOu_sazba = rst_report![ou_sazba] 'urcite nezaokrouhlovat
                                dblOu_castka = Round(rst_report![ou_castka], 2)
                                strOu_dtime_valid_from = "" & rst_report![ou_dtime_valid_from] & ""
                                strOu_dtime_valid_to = "" & rst_report![ou_dtime_valid_to] & ""
                                dblSu = Round(rst_report![su], 2)
                                If Not IsNull(rst_report![su_sazba]) Then dblSu_sazba = rst_report![su_sazba] 'urcite nezaokrouhlovat '2.2.2018 MM Pøidána podmínka, padal na NULL
                                If Not IsNull(rst_report![su_castka]) Then dblSu_castka = Round(rst_report![su_castka], 2) '2.2.2018 MM Pøidána podmínka, padal na NULL
                                strSu_dtime_valid_from = "" & rst_report![su_dtime_valid_from] & ""
                                strSu_dtime_valid_to = "" & rst_report![su_dtime_valid_to] & ""
                                intInterval = rst_report![interval]
                                'Format(rst_report![ou_castka], "#,##0.00")
                                
                                
                                '''''   intervaly
                                
                                'zapamatujeme si jistinu a kap. urok v den zesplatneni
                                If intInterval = 1 Then
                                    dblJistinaKeDniZaloby = dblOu_castka 'rst_inter![Amt_balance_principal] 'dblSu_castka alebo dblOu_castka z prveho intervalu
                                    'amt_uroky = dblSu_castka 'rst_inter![interest_overdue]
                                End If
                                If intInterval = cnt_int Then
                                    dblKapUrok = dblKu 'dblKu z posledneho intervalu (inde su nulls) 'rst_inter![interest_overdue]
                                    'XX nasledujuci riadok este over
                                    dblKapUrokKeDniZaloby = dblKu
                                    
                                    'text_pohl_uvod = "Nezaplacený smluvní úrok do zesplatnìní úvìru ve výši " & VlozCastkuText(dblKu) & "." '& Chr(10) 'dblKu po novom, rst_inter![interest_overdue] po starom, 'povodna testace, v maily 22.11.2017 chceli zmenit
                                    text_pohl_uvod = "Nezaplacený smluvní úrok " & format(c.GetLoans(cL).Rate_Interest_Standard(), "#,##0.0#") & " % v celkové výši " & VlozCastkuText(dblKu) & _
                                                    ", který je souhrnem nesplacených smluvních úrokù ode dne " & _
                                                    IIf(c.GetLoans(cL).blnUhrazena_alespon_jedna_splatka, "poslední øádnì zaplacené splátky", "naèerpání úvìru") & _
                                                    ", tj. ode dne " & DatumPosledniZaplaceneSplatky(c.GetLoans(cL).skp_loan) & _
                                                    ", do dne pøedcházejícího dni zesplatnìní, tedy do dne " & Ins.dtmDatumDenPredZesplatnenim & _
                                                    ", konkrétnì pak (poèet dnù prodlení/365 x aktuální jistina x " & format(c.GetLoans(cL).Rate_Interest_Standard(), "#,##0.0#") & "/100)."
                                End If
                                
                                
                                
                                '''''   obchodni urok
                                
                                ''''text_pohl = text_pohl & Chr(10) & "Smluvní úrok ve výši " & rate_interest & " % roènì z pohledávky " & Format(rst_inter(1), "#,##0.00") & " Kè od " & rst_inter![dtime_valid_from_ou] & " do " & rst_inter![dtime_valid_to_ou] & " v celkové výši " & Format(rst_inter![ou], "#,##0.00") & " Kè."
                                If dblOu > 0 Then
                                    text_pohl = text_pohl & Chr(10) & "Smluvní úrok ve výši " & format(dblOu_sazba * 100, "#,##0.0#") & " % roènì z pohledávky " & VlozCastkuText(dblOu_castka) & " od " & _
                                                strOu_dtime_valid_from & " do " & strOu_dtime_valid_to & " v celkové výši " & VlozCastkuText(dblOu) & "."
                                    amt_uroky = amt_uroky + dblOu
                                End If
                                
                                '''''   zakonny urok z prodleni
                                
                                'FA neni to uplne dobre v originale, prvy interval ma datum FROM vacsi nez datum TO a ciastka je nulova..
                                If dblSu > 0 Then
                                    text_pohl = text_pohl & Chr(10) & "Zákonný úrok z prodlení ve výši " & format(dblSu_sazba * 100, "#,##0.0#") & " % roènì z pohledávky " & VlozCastkuText(dblSu_castka) & " od " & _
                                                strSu_dtime_valid_from & " do " & strSu_dtime_valid_to & " v celkové výši " & VlozCastkuText(dblSu) & "."
                                    strUrokDruh = "Smluvní úrok a úrok z prodlení"
                                    amt_uroky = amt_uroky + dblSu
                                End If
                                
                                
                                
                                
                                rst_report.MoveNext
                            Next ii
                            
                            ' do celkovej vysky vymahanych urokov este priratame kapitalizovany
                            amt_uroky = amt_uroky + dblKapUrok
                        
                        End If 'if Ins.blnPmZamitnut
                        
                    Else 'XX hlasku uprav
                        If strReportChybiciUvery = "" Then
                            strReportChybiciUvery = c.GetLoans(cL).code_loan_number
                        Else
                            strReportChybiciUvery = strReportChybiciUvery & ", " & c.GetLoans(cL).code_loan_number
                        End If
                        'HlaskaUzivateli se vlozi az za cyklem pres uvery
                    End If
                rst_report.Close
            
            
            
            
            'Else '
            'If 1 = 0 And
            If ((Not blnReport) Or (Not Ins.blnPmZamitnut)) Then
            'If ((Not blnReport) Or ((Not Ins.blnPmZamitnut) And blnPlatbyPoZesplatneni)) Then
                '''   plnenie premennych k urokom ak nie su v reporte, alebo ak nam PM Rozsudek uznal vsetky pohladavky a zaroven mame platby po zesplatneni
                ' v kode som zachoval nazvy premennych ako uz boli nizsie (e.g. dblSu_sazba, ..) a prisposobil som ich plnenie druhej tabulke
                ' plneni dle "OBS" jak je popsano v RISK-849
                '
'                sql_inter = "SELECT a.amt_balance_principal, a.amt_interest_overdue, a.rate_interest_standard, a.rate_interest_repo, a.start_ou, a.konec_ou - 1 as konec_ou, a.ou, a.start_su, a.konec_su - 1 as konec_su, a.su " & _
'                            "FROM " & schema & ".intervaly a WHERE a.skp_loan = " & c.GetLoans(cL).skp_loan & " AND a.code_loan_number = '" & cLn & "' ORDER BY a.start_ou ASC"
                
                
                ' FA 1.3.2018 update: vytvoril som VIEWcko V_INS_INTERVALY ktore pouzijeme namiesto tabulky INTERVALY
                ' V_INS_INTERVALY by malo mat posledny interval skrateny tak ako to u insolvencnych klientov potrebujeme
                ' nasadene na prod s verziou 5.3, niekedy v marci 2018 - predtym fungoval napocet nizsie ktory ide z AP_COLL
                sql_inter = "SELECT a.amt_balance_principal, a.amt_interest_overdue, a.rate_interest_standard, a.rate_interest_repo, a.start_ou, a.konec_ou - 1 as konec_ou, a.ou, a.start_su, a.konec_su - 1 as konec_su, a.su " & _
                            "FROM " & schema & ".V_INS_INTERVALY a WHERE a.skp_loan = " & c.GetLoans(cL).skp_loan & " AND a.code_loan_number = '" & cLn & "' ORDER BY a.start_ou ASC"
                
                
                Set rst_inter = New ADODB.Recordset
                rst_inter.Open sql_inter, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
                    
                    If Not rst_inter.EOF Then
                        cnt_int = rst_inter.RecordCount
                        
                        For ii = 1 To cnt_int ' Z TABULKY V JEZEVCIKOVI
                            
                            'dblCelkem =
                            dblJistina = Round(rst_inter![Amt_balance_principal], 2) 'neskor este dblJistinaKeDniZaloby sa bude pouzivat
                            '' este dopln dblKu  - nacitat treba iba u posledneho zaznamu
                            dblKu = Round(rst_inter![amt_interest_overdue], 2)
                            dblOu = Round(rst_inter![ou], 2)
                            dblOu_sazba = rst_inter![Rate_Interest_Standard] 'urcite nezaokrouhlovat
                            ' dblOu_castka >>> nahradene za dblJistina
                            ' dblSu_castka >>> nahradene za dblJistina
                            strOu_dtime_valid_from = "" & rst_inter![start_ou] & ""
                            strOu_dtime_valid_to = "" & rst_inter![konec_ou] & ""
                            dblSu = Round(rst_inter![su], 2)
                            dblSu_sazba = sngZakonnyUrok + rst_inter![Rate_Interest_Repo]
                            strSu_dtime_valid_from = "" & rst_inter![start_su] & ""
                            strSu_dtime_valid_to = "" & rst_inter![konec_su] & ""
                            intInterval = ii ' mozno zbytocne
                            
                            
                            '''''   intervaly
                    
                            'zapamatujeme si jistinu a kap. urok v den zesplatneni
                            If intInterval = 1 Then
                                dblJistinaKeDniZaloby = dblJistina
                                amt_uroky = 0
                            End If
                            
                            If intInterval = cnt_int Then
                                dblKapUrok = dblKu
                                dblKapUrokKeDniZaloby = dblKu
                                
                                'text_pohl_uvod = "Nezaplacený smluvní úrok do zesplatnìní úvìru ve výši " & VlozCastkuText(dblKu) & "." '& Chr(10) 'dblKu po novom, rst_inter![interest_overdue] po starom, 'povodna testace, v maily 22.11.2017 chceli zmenit
                                text_pohl_uvod = "Nezaplacený smluvní úrok " & format(c.GetLoans(cL).Rate_Interest_Standard(), "#,##0.0#") & " % v celkové výši " & VlozCastkuText(dblKu) & _
                                                ", který je souhrnem nesplacených smluvních úrokù ode dne " & _
                                                IIf(c.GetLoans(cL).blnUhrazena_alespon_jedna_splatka, "poslední øádnì zaplacené splátky", "naèerpání úvìru") & _
                                                ", tj. ode dne " & DatumPosledniZaplaceneSplatky(c.GetLoans(cL).skp_loan) & _
                                                ", do dne pøedcházejícího dni zesplatnìní, tedy do dne " & Ins.dtmDatumDenPredZesplatnenim & _
                                                ", konkrétnì pak (poèet dnù prodlení/365 x aktuální jistina x " & format(c.GetLoans(cL).Rate_Interest_Standard(), "#,##0.0#") & "/100)."
                            End If
                            
                            
                            '''   obchodni urok
                            If dblOu > 0 Then
                                text_pohl = text_pohl & Chr(10) & "Smluvní úrok ve výši " & format(dblOu_sazba, "#,##0.0#") & " % roènì z pohledávky " & VlozCastkuText(dblJistina) & " od " & _
                                            strOu_dtime_valid_from & " do " & strOu_dtime_valid_to & " v celkové výši " & VlozCastkuText(dblOu) & "."
                                amt_uroky = amt_uroky + dblOu
                            End If
                            
                            '''   zakonny urok z prodleni
                            If dblSu > 0 Then
                                text_pohl = text_pohl & Chr(10) & "Zákonný úrok z prodlení ve výši " & format(dblSu_sazba, "#,##0.0#") & " % roènì z pohledávky " & VlozCastkuText(dblJistina) & " od " & _
                                            strSu_dtime_valid_from & " do " & strSu_dtime_valid_to & " v celkové výši " & VlozCastkuText(dblSu) & "."
                                strUrokDruh = "Smluvní úrok a úrok z prodlení"
                                amt_uroky = amt_uroky + dblSu
                            End If
                                                     
                                                    
                                                       
                            rst_inter.MoveNext
                            
                        Next ii
                        
                        ' do celkovej vysky vymahanych urokov este priratame kapitalizovany
                        amt_uroky = amt_uroky + dblKapUrok
                        
                    End If
                    
                rst_inter.Close
            End If
                        
            
            
            'If 1 = 0 and Then
            If 1 = 0 And ((Not blnReport) Or (Not Ins.blnPmZamitnut)) Then
            ' tohle je treti zpusob generovani intervalu k uverum.
            '''   plnenie premennych k urokom ak nie su v reporte, alebo ak nam PM Rozsudek uznal vsetky pohladavky
            
                '' vyhoda tejto tabulky oproti jezevcikovskej je, ze ma aj interest_overdue aj AMT_interest_overdue. Jezevcik iba tu druhu ma. Obcas ale potrebujeme aj opacnu... Trebalo by presunut napocet do Jez.
                ' NEBUDE FUNGOVAT, su_sazba nebyla nikdy naimplementovana do tabulky.. Pouziva se verze vyse, je tam k tomu VIEW v jezevcikovi
                sql_inter = "select interest_overdue, amt_balance_principal, amt_interest_overdue, dtime_valid_from_ou , dtime_valid_to_ou -1 as dtime_valid_to_ou, " & _
                            "dtime_valid_from_su, dtime_valid_to_su -1 as dtime_valid_to_su, ou, su, su_sazba, " & _
                            "CASE WHEN ('" & Ins.dtmDatumZesplatneni & "' BETWEEN dtime_valid_from AND dtime_valid_to) THEN AMT_INTEREST_OVERDUE ELSE 0 END As kap_urok_ke_dni_zaloby " & _
                            "from ap_coll.sr_coll_014_intervaly where skp_loan = " & c.GetLoans(cL).skp_loan & _
                            " order by dtime_valid_from_ou" 'tabulka se plni na ap_coll hristi
                Set rst_inter = New ADODB.Recordset
                rst_inter.Open sql_inter, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
                
                    If Not rst_inter.EOF Then
                        
                        
                        '' FA 16.2.2018 oprava chybnych castek odtial
                        ' V nekterych pripadech potrebujeme amt_interest_overdue, v nekterych interest_overdue
                        ' nechal som to obecne, v pripade ze sa odlisuju tak si uzivatel moze vybrat.
                        ' skarede riesenie, ale na obecnu logiku kedy je ktory spravny som neprisiel
                        
                        Dim intVolbaInterestOverdue As Integer
                        Dim dblInterestOverdue As Double
                        
                        If rst_inter![amt_interest_overdue] <> rst_inter![interest_overdue] And intVolbaInterestOverdue = 0 Then
                            'intVolbaInterestOverdue = (MsgBox("Použi druhú variantu pre kapitalizovaný úrok?", vbYesNo) = vbYes)
                            intVolbaInterestOverdue = MsgBox("Použi druhú variantu pre kapitalizovaný úrok?", vbYesNo) - 5 ' vbYes = 6, vbNo = 7
                            'dblInterestOverdue = Choose(blnVolbaInterestOverdue + 2, rst_inter![amt_interest_overdue], rst_inter![interest_overdue])
                        End If
                        
                        If intVolbaInterestOverdue > 0 Then
                            dblInterestOverdue = Choose(intVolbaInterestOverdue, rst_inter![amt_interest_overdue], rst_inter![interest_overdue])
                        Else
                            dblInterestOverdue = rst_inter![interest_overdue]
                        End If
                        
                        
''                         '''' povodny kod nizsie
''                        cnt_int = rst_inter.RecordCount
''                        rst_inter.MoveFirst
''                        'aktualni kap. urok
''                        'text_pohl_uvod = "Nezaplacený smluvní úrok do zesplatnìní úvìru ve výši " & VlozCastkuText(rst_inter![interest_overdue]) & "." '& Chr(10) 'dblKu po novom'povodna testace, v maily 22.11.2017 chceli zmenit
''                        text_pohl_uvod = "Nezaplacený smluvní úrok " & Format(c.GetLoans(cL).rate_interest(), "#,##0.0#") & " % v celkové výši " & VlozCastkuText(rst_inter![amt_interest_overdue]) & ", který je souhrnem nesplacených smluvních úrokù " & _
''                                        "ode dne poslední øádnì zaplacené splátky, tj. ode dne " & DatumPosledniZaplaceneSplatky(c.GetLoans(cL).skp_loan) & _
''                                        ", do dne pøedcházejícího dni zesplatnìní, tedy do dne " & Ins.dtmDatumDenPredZesplatnenim & _
''                                        ", konkrétnì pak (poèet dnù prodlení/365 x aktuální jistina x " & Format(c.GetLoans(cL).rate_interest(), "#,##0.0#") & "/100)."
''                        amt_uroky = rst_inter![interest_overdue] 'celkova castka za uroky, postupne se navysuje
''
                        
                        '' FA oprava potial
                        
                        
                        cnt_int = rst_inter.RecordCount
                        rst_inter.MoveFirst
                        'aktualni kap. urok
                        'text_pohl_uvod = "Nezaplacený smluvní úrok do zesplatnìní úvìru ve výši " & VlozCastkuText(rst_inter![interest_overdue]) & "." '& Chr(10) 'dblKu po novom'povodna testace, v maily 22.11.2017 chceli zmenit
                        text_pohl_uvod = "Nezaplacený smluvní úrok " & format(c.GetLoans(cL).Rate_Interest_Standard(), "#,##0.0#") & " % v celkové výši " & VlozCastkuText(dblInterestOverdue) & _
                                        ", který je souhrnem nesplacených smluvních úrokù ode dne " & _
                                        IIf(c.GetLoans(cL).blnUhrazena_alespon_jedna_splatka, "poslední øádnì zaplacené splátky", "naèerpání úvìru") & _
                                        ", tj. ode dne " & DatumPosledniZaplaceneSplatky(c.GetLoans(cL).skp_loan) & _
                                        ", do dne pøedcházejícího dni zesplatnìní, tedy do dne " & Ins.dtmDatumDenPredZesplatnenim & _
                                        ", konkrétnì pak (poèet dnù prodlení/365 x aktuální jistina x " & format(c.GetLoans(cL).Rate_Interest_Standard(), "#,##0.0#") & "/100)."
                        amt_uroky = dblInterestOverdue 'celkova castka za uroky, postupne se navysuje
                            
                        'zapamatujeme si jistinu a kap. urok v den zesplatneni
                        dblJistinaKeDniZaloby = rst_inter![Amt_balance_principal]
                        dblKapUrokKeDniZaloby = rst_inter![kap_urok_ke_dni_zaloby]
                        'dblSplatnaCastka =
                        dblKapUrok = rst_inter![amt_interest_overdue] 'chceme amt z prveho intervalu (rozdiel oproti interest_overdue ak mal klient platby po zesplatneni)
                        
                        For ii = 1 To cnt_int
                            
                            '''''    obchodni urok

                            If rst_inter![ou] > 0 Then
                                
                                text_pohl = text_pohl & Chr(10) & "Smluvní úrok ve výši " & format(c.GetLoans(cL).Rate_Interest_Standard(), "#,##0.0#") & " % roènì z pohledávky " & VlozCastkuText(rst_inter(1)) & " od " & _
                                            rst_inter![dtime_valid_from_ou] & " do " & rst_inter![dtime_valid_to_ou] & " v celkové výši " & VlozCastkuText(rst_inter![ou]) & "."
                                amt_uroky = amt_uroky + rst_inter![ou]
                                
                            End If
                            
                            
                            '''''    zakonny urok z prodleni
                            
                            'FA povodne tu toto bolo: text_pohl = text_pohl & Chr(10) & "Zákonný úrok z prodlení ve výši " & sngZakonnyUrok & " % roènì z pohledávky " & Format(rst_inter![amt_balance_principal], "#,##0.00") & " Kè od " & rst_inter![dtime_valid_from_su] & " do " & rst_inter![dtime_valid_to_su] & " v celkové výši " & Format(rst_inter![su], "#,##0.00") & " Kè."
                            If rst_inter![su] > 0 Then
                                
                                text_pohl = text_pohl & Chr(10) & "Zákonný úrok z prodlení ve výši " & format(sngZakonnyUrok + rst_inter![su_sazba], "#,##0.0#") & " % roènì z pohledávky " & _
                                            VlozCastkuText(rst_inter![Amt_balance_principal]) & " od " & _
                                            rst_inter![dtime_valid_from_su] & " do " & rst_inter![dtime_valid_to_su] & " v celkové výši " & VlozCastkuText(rst_inter![su]) & "."
                                
                                strUrokDruh = "Smluvní úrok a úrok z prodlení"
                                amt_uroky = amt_uroky + rst_inter![su]
                                
                            End If
                            
                            dblJistina = rst_inter![Amt_balance_principal] 'aktualni vyse pohledavky 'FA povodne amt_jis_akt, som premenoval
                            
                            rst_inter.MoveNext
                        Next ii
                        
                    End If
                rst_inter.Close
            End If
            
            
            
            
            ''''''      PLNENIE XML ELEMENTOV     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            
            
            
            
            Call pridejElementyPohledavka(elemIP, True) 'blnZahrnPrislusenstvi = True
            
            
            ''''''    Výše jistiny
            
            'Call naplnElementNumber("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/", "vyse_jistiny", c.GetLoans(cL).Amt_loan_principal_unpaid)
            'Debug.Print "dblJistinaKeDniZaloby: " & dblJistinaKeDniZaloby & ", dblJistina: " & dblJistina & ", dblJistina: " & dblJistina
            Call naplnElementNumber("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/", "vyse_jistiny", dblJistina)
            
            
            ''''''    06 Dùvod vzniku
            
            text = ""
            text = "Poskytnutí úvìru è. " & cLn & " na èástku " & VlozCastkuTextSkracenou(c.GetLoans(cL).Amt_loan_principal_total) & " ze dne " & dtmPodpis
            If strDodatek = 0 Then
                text = text & " na základì podpisu Rámcové smlouvy è. " & c.TextContractNumber & "."
            Else
                text = text & " na základì podpisu dodatku è. " & strDodatek & " k Rámcové smlouvì è. " & c.TextContractNumber & "."
            End If
             
            
            If Ins.blnPm Then
                
                ' pokud neni soud v reportu, vlozime okresni
                text = text & " Vìøitel svùj nárok na zaplacení pohledávky uplatnil u " & VysklonujSoudGenitiv(VlozSoud()) & ". Dlužníkovi byla "
                
                'EPR / rozsudek
                If UCase(Left(Ins.strSpisZnacka, 1)) = "E" Then
                    text = text & "elektronickým platebním rozkazem è.j. "
                    blnEpr = 1
                Else
                    text = text & "rozsudkem è.j. "
                End If
                
                'pokud neni spisova znacka, vlozime cislo jednaci
                text = text & VlozSpisovouZnacku() & " ze dne " & strDatumVydani & ", který nabyl právní moci dne " & strDatumPm & ", uložena povinnost vìøiteli pohledávku zaplatit."
                
            End If
            
            
            Call naplnElement("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/", "duvod_vzniku", text)
            text = ""
            
            
            
            ''''''     07 - dalsi okolnosti
            
            
            text = "Finanèní prostøedky ve výši " & VlozCastkuTextSkracenou(c.GetLoans(cL).Amt_loan_principal_total) & " byly dne " & dtmCerpani & " naèerpány na bìžný úèet klienta - è. úètu: " & _
                    c.GetLoans(cL).Text_bank_account_number & "/3030. Vzhledem k závažnému porušení smluvních povinností ze strany dlužníka bylo v souladu s Obchodními podmínkami a na základì " & _
                    "Podmínek pro používání pùjèky Air Bank a.s. - str. 8, bod 3a - dne " & c.GetLoans(cL).Date_payoff & " rozhodnuto ze strany Air Bank a.s. o zesplatnìní úvìru."
        
            Call naplnElement("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/", "dalsi_okolnosti", text)
            text = ""
            
            ''''''     08 - vykonatelnost  - zahrneme pouze pokud PM
            
            If Ins.blnPm Then
            
                Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi", "vykonatelnost")
                Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/vykonatelnost", "pro_castku")
                Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/vykonatelnost", "dle")


                Call naplnElementNumber("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/vykonatelnost[0]/", "pro_castku", Round(dblJistina + amt_uroky, 2)) 'Round(dblJistina + amt_uroky, 2)
                
                NaplnTextVykonatelnost 'premenna text_vykonatelnost_dle sa vyuziva aj pre debety / SOPy
                Call naplnElement("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/vykonatelnost[0]/", "dle", text_vykonatelnost_dle)
                
            End If
            
            
            ''''''    9 Prislusenstvi
            
            Call naplnElement("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/prislusenstvi/", "druh", strUrokDruh)
            Call naplnElementNumber("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/prislusenstvi/", "vyse", Round(amt_uroky, 2))
            Call naplnElement("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/prislusenstvi/", "zpusob_vypoctu", text_pohl_uvod & text_pohl)
            text_pohl_uvod = "" 'uz nepotrebne
            text_pohl = ""
            
            
            ''''''    10 Vlastnosti pohledavky
            
            Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/vlastnosti_pohledavky/splatna_opak", "splatna")
            Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/vlastnosti_pohledavky/splatna_opak/splatna", "splatna_od")
            Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/vlastnosti_pohledavky/splatna_opak/splatna", "v_castce")
            
            Call naplnElement("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/vlastnosti_pohledavky[0]/splatna_opak[0]/splatna/", "splatna_od", VlozDatumDoXML(c.DatePayoff))
            'Debug.Print "dblJistinaKeDniZaloby: " & dblJistinaKeDniZaloby & ", dblJistina: " & dblJistina & ", dblKapUrok: " & dblKapUrok
            
            'v nasledujucom elemente ma byt VZDY zustatek ke dni zesplatneni (pozor, to znamena ze to nemusi byt suma inych ciastok s ktorymi pracujeme)
            ' doplneno Jez.v.5.3  '''msgbox o par riadkov nizsie zakomentovany
            Call naplnElementNumber("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/vlastnosti_pohledavky/splatna_opak/splatna/", "v_castce", ZustatekKeDniZesplatneni(c.GetLoans(cL).skp_loan)) ' FA 20.12.2017 - mala by sa plnit ciastka rovnako ako v sablone EPR.
            'Debug.Print c.GetLoans(cL).skp_loan, ZustatekKeDniZesplatneni(c.GetLoans(cL).skp_loan)
            ' v_castce: Povodne zadanie: Jistina + kapitalizovaný úrok ke dni zesplatnìní (je jedno, jestli probìhly nìjaké platby po zesplatnìní, nebo neprobìhly)
            ' v_castce: EPR sablona to plni spravne. V tom hlavnom texte, v casti II., je veta typu (CUID 543335)::  Ke dni zesplatnìní tedy byla dlužná pohledávka tvoøena nesplacenou jistinou ve výši 45 881,17 Kè a nesplaceným smluvním úrokem z úvìru (kapitalizovaný úrok) ve výši 2 779,81 Kè.
            ' v_castce: My chceme SUMU tychto dvoch ciastok z predosleho riadku, pritom ale jednotlivo sa v XML nikde vyskytnut nemaju.
            
        End If
        
    Next cL
    
    ' HlaskaUzivateli "U úvìrù v èásti 10 - Vlastnosti pohledávky - chybí èástka. Nutno doplnit ruènì." ' Jez.v.5.3

End If




If strReportChybiciUvery <> "" Then
    HlaskaUzivateli "V reportu PM pohledávek není ani zmínka k tìmto úvìrùm: " & strReportChybiciUvery & "."
End If

If Len(strReportChybiciUdaje) > 1 Then
    'HlaskaUzivateli "Z reportu se nenahráli tyto údaje: " & strReportChybiciUdaje & ". Je potøeba je doplnit ruènì."
    HlaskaUzivateli "V reportu PM pohledávek chybí tyhle konkrétní údaje k úverùm: " & strReportChybiciUdaje & ". Je potøeba je doplnit ruènì."
End If

End Sub



Private Sub nacteniDebetuASopuZReportu(id_person As Variant)
'naplnia sa premenne dblSop, dblDebet, blnReportSop, blnReportDebet, strReportBankAccountNumber (u debetu)

Dim strDebetOrSop As String
blnReportDebet = False
blnReportSop = False


'sql pre debety aj sopy
sql_report = "SELECT a.cuid, a.epr, a.typ, a.bank_account_number, nvl(a.debet, 0) debet, nvl(a.sop, 0) sop, a.travel_expenses, a.other_expenses, a.note, a.special " & _
             "FROM ap_coll.vs_import_pm_all a " & _
             "WHERE upper(a.typ) IN ('DEBET', 'SOP') AND a.cuid = " & c.CUID & " AND a.epr LIKE '%" & Ins.strSpisZnacka & "%' ORDER BY a.typ ASC"

blnReport = False
Set rst_report = New ADODB.Recordset
rst_report.Open sql_report, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
    If Not rst_report.EOF Then
        rst_report.MoveFirst
     
        Do Until rst_report.EOF = True
        
            blnReport = True
            strDebetOrSop = rst_report![typ]
            
            'nacitame to podstatne
            Select Case strDebetOrSop
            
                Case "DEBET"
                    
                    'chce to nejaky vhodny iterovatelny format pre ukladanie viacerych debetov alebo SOPov, zatial iba msgbox
                    If blnReportDebet Then
                        HlaskaUzivateli "Tenhle klient má v reportu PM pohledávek víc než jeden debet. To jezevèíka zatím nikdo nenauèil, musíš na to mrknout sám."
                    End If
                    blnReportDebet = True
                    dblDebet = Round(rst_report![debet], 2)
                    
                    If IsNull(rst_report![bank_account_number]) Then
                        HlaskaUzivateli "V reportu PM pohledávek chybí bank_account_number."
                    Else
                        strReportBankAccountNumber = rst_report![bank_account_number]
                    End If
                    
                    
                Case "SOP"
                    
                    If blnReportSop Then
                        HlaskaUzivateli "Tenhle klient má v reportu PM pohledávek víc než jeden SOP. To jezevèíka zatím nikdo nenauèil, musíš na to mrknout sám."
                    End If
                    blnReportSop = True
                    'dblSop = Round(rst_report![SOP], 2)
                
                
                Case Else 'nikdy
                    HlaskaUzivateli "Jezevèík error J-84472." 'nemal by sa nikdy stat lebo debet / sop sme obmedzili uz v SQL statement
            
            End Select
            
            
            'Hlaska pro hodnoty v posl. 4 sloupcich
            Call HlaskaPokudNeniNull(rst_report![TRAVEL_EXPENSES], "TRAVEL_EXPENSES")
            Call HlaskaPokudNeniNull(rst_report![OTHER_EXPENSES], "OTHER_EXPENSES")
            Call HlaskaPokudNeniNull(rst_report![note], "NOTE")
            Call HlaskaPokudNeniNull(rst_report![SPECIAL], "SPECIAL")
            
            
            rst_report.MoveNext
        Loop
        
        
    
    Else
        HlaskaUzivateli "V reportu PM pohledávek není ani zmínka po debetu nebo sopu."
    
    End If
rst_report.Close

End Sub


Private Sub prihlaskaPohledavky_debet(ByVal id_person As Variant)

''''
'''''      DEBET
''''

Dim blnChybiDatumUpadku As Boolean
Dim dblVyskaDebetuVOkolnostech As Double

If (InStr(UCase(Ins.strDatumDenPredUpadkem), "DOPLN") <> 0) Or IsNull(Ins.strDatumDenPredUpadkem) Then 'DOPLN sa nastavuje ako hodnota ak datum upadku je NULL
    blnChybiDatumUpadku = True
    HlaskaUzivateli "Nelze nahrát výšku debetu: v jezevèíkovi není zadán datum úpadku klienta." ' Nahraj datum do jezevèíka a zkus to zítra (až se naètou èerstvá data)."
End If

'Debug.Print skutecnostiDebetText(c.GetAccounts(3).skpBankAccount, c.id_person, "IP")

If c.GetCnt_Accounts > 0 Then

'vybrat jen požadované úèty - ty co mají nepovolený debet. Pro JPK, spoøicí úèty a úèty v cizích mìnách bude vyskakovat upozornìní pro manuální doplnìní.
    For cA = 1 To c.GetCnt_Accounts
        'tady vybrat jen požadované úvìry - zesplatnìné + k odpisu + odepsané
        
        'oproti EPR sa berie do uvahy este dblDebet (pracujeme totiz so zostatkom zo dna pred upadkom)
        If blnReportDebet = False Then
            'ak sa dblDebet nenaplnil z reportu, naplnime ho z jezevcika
            dblDebet = Round(Ins.ZjistiDebetPredUpadkem(c.id_person, c.GetAccounts(cA).skpBankAccount), 2)
        End If
        
        
        If dblDebet > 0 And c.GetAccounts(cA).Amt_balance < 0 And c.GetAccounts(cA).Account_type = "BÚ" And c.GetAccounts(cA).Code_currency = "CZK" Then
            'pozor, ono zaroven ta podmienka riesi ak ma klient 10 beznych uctov, aby sa debet vypisoval iba k tomu jednemu co je v zapore
            'teraz je to spravene tak, ze nas zaujima iba ak ten ucet je ESTE AJ DNES v debete
            
            
            '  pridani vsech elementu do XML
            elemIP = elemIP + 1
            Call pridejElementyPohledavka(elemIP, False) 'blnZahrnPrislusenstvi = False 'DEBET nema prislusenstvi

            ''''''    Vyse jistiny
            
            
            If Not blnChybiDatumUpadku Then
                Call naplnElementNumber("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/", "vyse_jistiny", Round(dblDebet, 2))
            Else
                Call naplnElementNumber("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/", "vyse_jistiny", -0)
            End If
            

            ''''''    06 Dùvod vzniku
            
            
            text = ""
            text = "Vìøitel uzavøel s dlužníkem dne " & Ins.strDatumRamcovka & " Rámcovou smlouvu è. " & c.TextContractNumber & " v souladu s Obchodními podmínkami Air Bank a.s. " & _
                    "- èlánek Vznik rámcové smlouvy. Na základì Rámcové smlouvy byl dlužníkovi aktivován bìžný úèet è. " & c.GetAccounts(cA).Text_bank_account_number & "/3030."
            Call naplnElement("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/", "duvod_vzniku", text)
            text = ""
            
            
            ''''''    07 - dalsi okolnosti

            'XX?
            'ak je to ten isty ucet co v reporte, parada (napln element alebo nerobi nic)
            'ak je to iny, vyhod hlasku ze mozno klient ma este iny debetny ucet
            dblVyskaDebetuVOkolnostech = 0
            Call naplnElement("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/", "dalsi_okolnosti" _
                                , skutecnostiDebetText(c.GetAccounts(cA).skpBankAccount, c.id_person, "IP", dblVyskaDebetuVOkolnostech))
            
            If dblVyskaDebetuVOkolnostech <> dblDebet Then
                HlaskaUzivateli "Výška debetu v okolnostech (" & dblVyskaDebetuVOkolnostech & ") se liší od výšky debetu ve výši jistiny (" & dblDebet & "). Radìji na to mrkni."
            End If
            
            
            ''''''    08 - vykonatelnost  - zahrneme pouze pokud PM   - castka a dle
    
    
            If Ins.blnPm Then
            
                Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi", "vykonatelnost")
                Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/vykonatelnost", "pro_castku")
                Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/vykonatelnost", "dle")
                
                Call naplnElementNumber("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/vykonatelnost[0]/", "pro_castku", Round(dblDebet, 2))
                Call naplnElement("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/vykonatelnost[0]/", "dle", text_vykonatelnost_dle)
            End If
            
            
            ''''''    9 Prislusenstvi
            'nic se nevyplnuje
            
            
            ''''''    10 Vlastnosti pohledavky
            
            Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/vlastnosti_pohledavky/splatna_opak", "splatna")
            Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/vlastnosti_pohledavky/splatna_opak/splatna", "splatna_od")
            Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/vlastnosti_pohledavky/splatna_opak/splatna", "v_castce")
            
            Call naplnElement("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/vlastnosti_pohledavky[0]/splatna_opak[0]/splatna/", "splatna_od", VlozDatumDoXML(strDatumPosledniDebetniTrn)) ' "_POSL.TRN_")
            
            If Not blnChybiDatumUpadku Then 'InStr(UCase(Ins.strDatumDenPredUpadkem), "DOPLN") = 0 Then 'DOPLN sa nastavi ako hodnota ak NULL
                Call naplnElementNumber("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/vlastnosti_pohledavky/splatna_opak/splatna/", "v_castce", Round(dblDebet, 2))
            Else
                'sablona(debet) je nepouzitelna bez datumu upadku
                Call naplnElementNumber("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/vlastnosti_pohledavky/splatna_opak/splatna/", "v_castce", -0)
            End If
            
            
        Else
        
            'JPK a cizí mìna
            'FA_XX FA XX  ´- Lucka - chceme nasledujuce hlasky nechat ako su, alebo obmedzit na pripady ked mali debet AJ dnes AJ ku dnu pred zesplatnenim? zmena textace mozno?
            If c.GetAccounts(cA).Amt_balance < 0 And c.GetAccounts(cA).Account_type <> "BÚ" Then
                HlaskaUzivateli ("Klient má úèet se záporným zùstatkem (ke dnešku) typu Spoøicí úèet nebo JPK. Je tøeba doplnit pro tento nepovolený debet ruènì všechny údaje.")
            Else
                'If c.GetAccounts(cA).Amt_balance < 0 And c.GetAccounts(cA).Code_currency <> "CZK" Then
                If c.GetAccounts(cA).Amt_balance < 0 And c.GetAccounts(cA).Code_currency <> "CZK" Then
                    HlaskaUzivateli ("Klient má úèet se záporným zùstatkem (ke dnešku) v cizí mìnì. Je tøeba doplnit pro tento nepovolený debet ruènì všechny údaje.")
                End If
            
            End If
                
        End If
        
    Next cA

End If


End Sub

Private Sub prihlaskaPohledavky_sop(ByVal id_person As Variant)

''''
'''''      SOP
''''

If blnReportSop Or (Round(Ins.dblAmtSop, 2) > 0) Then 'ta druha podminka neni dle zadani
    
    elemIP = elemIP + 1
    Call pridejElementyPohledavka(elemIP, False) 'SOP je bez prisulsenstvi
    
    
    
    ''''''    Vyse jistiny
    
    
    'výše SOPu podle jezevèíka, èást že mùže být zaplacena
    
'    If blnReportSop Then
'        Call naplnElementNumber("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/", "vyse_jistiny", Round(dblSop - c.SOP.Pausal, 2))
'    Else
        Call naplnElementNumber("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/", "vyse_jistiny", Round(Ins.dblAmtSop - Ins.dblAmtSopPaid, 2))
'    End If
    
    
    ''''''    06 Dùvod vzniku
    
    text = ""
    text = "Povinnost nahradit vìøiteli náklady obèanskoprávního øízení uložená dlužníkovi "
    If blnEpr Then
        text = text & "elektronickým platebním rozkazem vydaným "
    Else
        text = text & "rozsudkem vydaným " 'è.j. " & strCisloJednaci & " || " & Ins.strSpisZnacka
    End If
    
    text = text & VysklonujSoudInstrumental(VlozSoud())
    text = text & " è.j. " & VlozSpisovouZnacku() & " ze dne " & strDatumVydani & ", který nabyl právní moci dne " & strDatumPm & "."
    
    Call naplnElement("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/", "duvod_vzniku", text)
    text = ""
    

    
    
    
    ''''''    07 - dalsi okolnosti
    
    ' nic se nevyplnuje
    
    
    ''''''    08 - vykonatelnost - vzdy - castka a dle
    
    Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi", "vykonatelnost")
    Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/vykonatelnost", "pro_castku")
    Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/vykonatelnost", "dle")
    
    
    ''   castka
    
'    If blnReportSop Then
'        Call naplnElementNumber("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/vykonatelnost[0]/", "pro_castku", Round(dblSop - c.SOP.Pausal, 2))
'    Else
        Call naplnElementNumber("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/vykonatelnost[0]/", "pro_castku", Round(Ins.dblAmtSop - Ins.dblAmtSopPaid, 2))
'    End If
    
    Call naplnElement("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi[0]/vykonatelnost[0]/", "dle", text_vykonatelnost_dle)
    
    
    
    ''''''    9 Prislusenstvi
    
    'nic se nevyplnuje
    
    
    ''''''    10 Vlastnosti pohledavky
    
    Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/vlastnosti_pohledavky/splatna_opak", "splatna")
    Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/vlastnosti_pohledavky/splatna_opak/splatna", "splatna_od")
    Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/vlastnosti_pohledavky/splatna_opak/splatna", "v_castce")
    
    '  splatna od
    Call naplnElement("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/vlastnosti_pohledavky[0]/splatna_opak[0]/splatna/", "splatna_od", VlozDatumDoXML(strDatumPm))
    If InStr(UCase(strDatumPm), "DATUM") <> 0 Then
        HlaskaUzivateli "V SOPu chybí datum splatnosti."
    End If
    
    '  v castce
'    If blnReportSop Then
'        Call naplnElementNumber("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/vlastnosti_pohledavky/splatna_opak/splatna/", "v_castce", Round(dblSop - c.SOP.Pausal, 2))
'    Else
        Call naplnElementNumber("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/vlastnosti_pohledavky/splatna_opak/splatna/", "v_castce", Round(Ins.dblAmtSop, 2))
'    End If

End If

'Exit Sub
'connection_problem:
'
'If Mid(Err.Description, 1, 9) = "ORA-03114" Then
'err_connection = True
'Else
' MsgBox Err.Description
'End If

End Sub


Private Sub VynulujPremenne()
 
strMsgbox = ""
strDatumPosledniDebetniTrn = "__DATUM__"

Set Ins = Nothing
Set rst_report = Nothing
elemIP = -1
strCisloUctu = ""
text = ""
text_vykonatelnost_dle = ""
objSelection = Empty
sql_report = ""

strReportChybiciUdaje = ""
strReportChybiciUvery = ""

dblDebet = 0
'dblSop = 0
blnReport = False
blnReportDebet = False
blnReportSop = False
blnEpr = False
'blnPlatbyPoZesplatneni = False

StrCisloJednaci = "__0__"
strSoudReport = "___SOUD___"
strDatumVydani = "___DATUM VYDANI___"
strDatumPm = "___DATUM PM___"
strDatumVykon = "___DATUM VYKONU___"

strReportBankAccountNumber = ""

End Sub




Private Function ZustatekKeDniZesplatneni(in_skp_loan As String) As Double
' RISK-885, Jez.v.5.3
' na konci prihlaskaPohledavky_uver
' kod je prevzaty z EPR_XML, CAST II.III


    'v nasledujucom elemente ma byt VZDY zustatek ke dni zesplatneni (pozor, to znamena ze to nemusi byt suma inych ciastok s ktorymi pracujeme)
    'nie je dokoncene. Aktualne nefunguje, pretoze v tabulke intervalov na AP_COLL nesedi hodnota...
    'idealne by bolo prejst kompletne na nejaku jezevcikovsku tabulku
    'msgbox o par riadkov nizsie vyhod az to bude ok
    
    'Call naplnElementNumber("//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & elemIP & "]/nezajistena_jednotlive[0]/vlastnosti_pohledavky/splatna_opak/splatna/", "v_castce", Round(dblJistinaKeDniZaloby + dblKapUrokKeDniZaloby, 2)) ' FA 20.12.2017 - mala by sa plnit ciastka rovnako ako v sablone EPR.
    ' v_castce: Povodne zadanie: Jistina + kapitalizovaný úrok ke dni zesplatnìní (je jedno, jestli probìhly nìjaké platby po zesplatnìní, nebo neprobìhly)
    ' v_castce: EPR sablona to plni spravne. V tom hlavnom texte, v casti II., je veta typu (CUID 543335)::  Ke dni zesplatnìní tedy byla dlužná pohledávka tvoøena nesplacenou jistinou ve výši 45 881,17 Kè a nesplaceným smluvním úrokem z úvìru (kapitalizovaný úrok) ve výši 2 779,81 Kè.
    ' v_castce: My chceme SUMU tychto dvoch ciastok z predosleho riadku, pritom ale jednotlivo sa v XML nikde vyskytnut nemaju.


Dim sql12 As String, jistina_zesp As Double, urok_zesp As Double

sql12 = "select skp_loan, id_person, nvl(jplatba,0), nvl(uplatba,0), nvl(jpredpis,0), nvl(upredpis,0), nvl(ip,0), nvl(ii,0), platba_den_po_zesplatneni, " & _
        "NVL(JPREDPIS, 0) + NVL(IP, 0) - NVL(JPLATBA, 0) AS jistina_zesp, " & _
        "NVL(UPREDPIS, 0) + NVL(II, 0) - NVL(UPLATBA, 0) AS  urok_zesp " & _
        "FROM " & schema & ".sablona_epr9 where skp_loan = " & in_skp_loan

Set rst1 = New ADODB.Recordset
rst1.Open sql12, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst1.EOF Then
        rst1.MoveFirst
        ' jez.v.5.7
        ' jistina_zesp = rst1(4) + rst1(6) - rst1(2)
        ' urok_zesp = rst1(5) + rst1(7) - rst1(3)
        jistina_zesp = rst1![jistina_zesp]
        urok_zesp = rst1![urok_zesp]
    End If
rst1.Close

ZustatekKeDniZesplatneni = jistina_zesp + urok_zesp

End Function



Private Sub ZkontrolujPrichoziPlatby(ByVal CUID As Long)
    
    ''' SUB vyhazuji ve verzi jez.v.5.7. Je to posledni odkaz na nefunkcni napocet (tabulke PB_040_SUMMARY) v JEzevcikovi.
    ''' Marketa Simkova mi potvrdila, ze hlasku o tom ze klient neco uhradil po evidenci PM pohledavky obcas vidi, odklikava ji, ale pak nic nemeni.
'''
'''    Dim sql As String
'''    Dim rst As adodb.Recordset
'''    Dim cnt_epr As Integer
'''
'''    sql = "SELECT a.cuid, a.epr, nvl(a.kontrola, '') kontrola, nvl(a.suma_uhrady, 0) suma_uhrady FROM AP_COLL.PB_040_SUMMARY a WHERE a.cuid = " & CUID
'''    '--AND a.epr = '26 C 202/2015 - 22'
'''
'''    Set rst = New adodb.Recordset
'''    rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''
'''        If Not rst.EOF Then
'''
'''            rst.MoveFirst
'''            cnt_epr = rst.RecordCount
'''
'''
'''            If cnt_epr > 1 Then
'''                HlaskaUzivateli ">> Klient má více než 1 pravomocní rozhodnutí (v Reportu PM pohledávek, karta Souhrn) - to jezevèík bohužel neumí.<<<" 'vygeneruje sa nejake chaoticke XML
'''            End If
'''
'''            'ak <> OK tak hlaska
'''            If rst![kontrola] <> "OK" Then
'''                HlaskaUzivateli "Poznámka z kontroly v Reportu PM pohledávek, karta Souhrn: " & rst![kontrola]
'''            End If
'''
'''            'ak suma_uhrady nie je 0 tak hlaska
'''            If rst![suma_uhrady] <> 0 Then
'''                'HlaskaUzivateli "Tenhle klient už nìco uhradil po evidenci PM pohledávky. To jezevèík zatím neumí."
'''                blnPlatbyPoZesplatneni = True
'''            End If
'''
'''            'rst.MoveNext
'''
'''        End If
'''
'''    rst.Close
    

End Sub




Private Function VysklonujSoudInstrumental(ByVal SOUD As String) As String
'  skloòujeme soud z reportu
Dim polohaPredlozky As Integer

If strSoudReport <> "" And SOUD <> "" And (InStr(strSoudReport, "SOUD_") = 0) And (strSoudReport <> cfg.cst_missing_string_value) Then 'pouze pokud se soud doplnil (nechceme duplicitu hlasku)
    
    If InStr(LCase(SOUD), "okres") <> 0 Then
        VysklonujSoudInstrumental = "Okresním soudem "
    ElseIf InStr(LCase(SOUD), "kraj") <> 0 Then
        VysklonujSoudInstrumental = "Krajským soudem "
    ElseIf InStr(LCase(SOUD), "obvod") <> 0 Then
        VysklonujSoudInstrumental = "Obvodním soudem "
    ElseIf InStr(LCase(SOUD), "mìst") <> 0 Then
        VysklonujSoudInstrumental = "Mìstským soudem "
    Else
        HlaskaUzivateli ("Z nìjakého dùvodu se nepovedlo vyskloòovat soud z reportu (v èásti SOP)")
    End If
    
    'polohaPredlozky = WorksheetFunction.max(InStr(Soud, " v"), InStr(Soud, " pro"), InStr(strSoudReport, " Plzeò"))
    polohaPredlozky = InStr(SOUD, Split(SOUD, " ")(2)) - 1
    VysklonujSoudInstrumental = VysklonujSoudInstrumental & Right(SOUD, Len(SOUD) - polohaPredlozky)

Else
    
    VysklonujSoudInstrumental = SOUD
    
End If

End Function


Private Sub pridejElementyPohledavka(ByVal cisloElem As Integer, ByVal blnZahrnPrislusenstvi As Boolean) 'elemIP As Integer, blnZahrnPrislusenstvi As Boolean)

Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak", "pohledavka")

Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & cisloElem & "]", "nezajistena_jednotlive")
Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & cisloElem & "]/nezajistena_jednotlive[0]", "pohledavka_zahlavi")
Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & cisloElem & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi", "vyse_jistiny")
Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & cisloElem & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi", "duvod_vzniku")
Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & cisloElem & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi", "dalsi_okolnosti")

If (blnZahrnPrislusenstvi) Then 'prislusenstvi se u DEBETu nevyplnuje
    Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & cisloElem & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi", "prislusenstvi")
    Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & cisloElem & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi/prislusenstvi", "druh")
    Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & cisloElem & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi/prislusenstvi", "vyse")
    Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & cisloElem & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi/prislusenstvi", "zpusob_vypoctu")
End If

Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & cisloElem & "]/nezajistena_jednotlive[0]/pohledavka_zahlavi", "celkova_vyse")
Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & cisloElem & "]/nezajistena_jednotlive[0]", "vlastnosti_pohledavky")
Call pridejElement(0, "//d:root/prihlaska_pohledavky[0]/pohledavka_opak[0]/pohledavka[" & cisloElem & "]/nezajistena_jednotlive[0]/vlastnosti_pohledavky", "splatna_opak")
'Call UlozSoubor("W:\Vymahani\LATE\IP????\" & c.Nazev_slozky_bez_lomitka, "BaseXmlIp_" & c.Nazev_slozky_bez_lomitka & ".xml") 'doplnit požadované ukládání soubor na PROD

End Sub


Private Sub NactiPromennu(ByRef promenna As Variant, ByRef rst_stlpec As Variant, ByVal colName As String)

If IsNull(rst_stlpec) Then
    If strReportChybiciUdaje = "" Then
        strReportChybiciUdaje = colName & "(" & c.GetLoans(cL).code_loan_number & ")"
    Else
        strReportChybiciUdaje = strReportChybiciUdaje & ", " & colName & " (" & c.GetLoans(cL).code_loan_number & ")"
    End If
    
    'HlaskaUzivateli "Z reportu se nenahrál " & colname & ", je potøeba doplnit ruènì (" & c.GetLoans(cL).Code_loan_number & ")."
    promenna = "___" & UCase(colName) & "___"
Else
    promenna = rst_stlpec
End If

End Sub

Private Sub HlaskaPokudNeniNull(ByRef rst_stlpec As Variant, ByVal colName As String)

If Not IsNull(rst_stlpec) Then
    HlaskaUzivateli "V reportu PM pohledávek je záznam pro " & colName & ", jezevèík jej nijak nebere do úvahy."
End If


End Sub





Private Sub NaplnTextVykonatelnost()

If blnEpr Then
    text_vykonatelnost_dle = "EPR "
Else
    text_vykonatelnost_dle = "rozsudku "
End If

'pokud neni soud v reportu, vlozime Okresni soud
'pokud neni spisova znacka, vlozime cislo jednaci
text_vykonatelnost_dle = text_vykonatelnost_dle & VysklonujSoudGenitiv(VlozSoud()) & _
                         ", è.j. " & VlozSpisovouZnacku() & _
                         " ze dne " & strDatumVydani & ", který nabyl právní moci dne " & strDatumPm

End Sub


Private Function VlozSpisovouZnacku() As String
'pokud neni spisova znacka, vlozime cislo jednaci
If InStr(StrCisloJednaci, "__") = 0 Then
    VlozSpisovouZnacku = StrCisloJednaci
Else
    VlozSpisovouZnacku = Ins.strSpisZnacka 'spisova znacka
End If

End Function

Private Function VlozSoud()
'pokud neni soud v reportu, vlozime okresni soud
If strSoudReport <> "" And InStr(strSoudReport, "SOUD_") = 0 Then
    VlozSoud = strSoudReport
Else
    VlozSoud = c.SOUD.okresni_soud
End If

End Function

Public Sub testCastky()
Dim castka As Double

castka = 11

Debug.Print "  " & VlozCastkuText(castka)
Debug.Print "  " & VlozCastkuTextSkracenou(castka)
Debug.Print "  " & VlozCastku(castka)

End Sub

Private Function VlozCastkuText(ByVal value) As String
    ''ked vkladame castky do textaci, chceme zaokruhlit a pouzit desatinnu ciarku
    'value = CStr(Round(value, 2))
    'VlozCastkuText = Replace(value, ".", ",")
    
    '''  puvodni kod do verze 5.3
    'new, naviac este oddelit tisicky chceme
    'VlozCastkuText = format(value, "Currency")
    
    ''' novy kod od verze 5.3
    VlozCastkuText = VlozCastku(value)
End Function

Private Function VlozCastkuTextSkracenou(ByVal value) As String
    VlozCastkuTextSkracenou = format(value, "#,##0") & " Kè"
End Function



Private Sub naplnElementNumber(cilElement As String, nameElement As String, value As Variant)
    Dim atribut As Object
    Set atribut = XDoc.SelectNodes(cilElement & nameElement)
    
    'If IsNumeric(value) Then
    '    value = CStr(Round(value, 2))
    'End If
    value = CStr(Round(value, 2))
    atribut(0).text = Replace(value, ",", ".")
End Sub


