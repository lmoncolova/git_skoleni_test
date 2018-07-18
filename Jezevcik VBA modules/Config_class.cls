VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' class created by Filip Antol 26.4.2018
' konfiguracni soubor (zmes hodnot ktere se muzou kdykoliv menit pokud si uzivatel rekne)
' tady hazet vsechny veci, ktere by jinak koncili jako PUBLIC promenne
' misto pro fixne hodnoty, ktere v kodu nemaji co delat

Private pcstDEFAULT_AMT_SOP_VALUE As Long
Private Const pcst_missing_string_value As String = "_____XXXXXXX_____"
Private Const pcst_ucet_522 As String = "30300522" ' chci jako string


Public Enum enum_exekucni_titul
    ' dle stavu v nalezacim rizeni a dle nazvu spisove znacky rozlisujeme 4 varianty. RISK-1096, jez.v.5.6
    exe_titul_none  ' default
    exe_titul_EPR
    exe_titul_platebni_rozkaz
    exe_titul_rozsudek
    exe_titul_usneseni_EPR
    exe_titul_usneseni_C
End Enum


Public Enum barva
    ' barvy primarne pro obarveni historie klienta na kartach
    barvaModra = 16711680        ' RGB(0, 0, 255)    ' Stavy
    barvaZelenaSvetla = 5921280  ' RGB(0, 90, 90)    ' Klient
    barvaSeda = 5263440          ' RGB(80, 80, 80)   ' Poznamky
    barvaRuda = 255              ' RGB(255, 0, 0)    ' Upozorneni
    barvaFialova = 12583104      ' RGB(192, 0, 192)  ' SOPy
    barvaRudaSvetla = 3289855    ' RGB(255, 50, 50)  '
    barvaCerna = 0               ' RGB(0, 0, 0)
    barvaZelenaAirBank = 4046997 ' RGB(149, 192, 61)
    barvaLososova = 12632319     ' RGB(255, 192, 192)
    barvaSedaSvetla = 15790320   ' RGB(240, 240, 240)
    barvaOranzova = 42495        ' RGB(255, 165, 0)
    barvaBila = 16777215         ' RGB(255, 255, 255)
End Enum


Public Enum stavAutorizace
    stavAuth_None
    stavAuth_Autorizovano
    stavAuth_Ceka_na_autorizaci
End Enum


Public Enum typRizeni
    typRizeni_none ' default
    typRizeni_EPR
    typRizeni_EXE
    typRizeni_INS
    typRizeni_TREST
    typRizeni_DED
    typRizeni_HEIR
End Enum


Private pWbJez As Workbook
Private pSlovnikUzivatelskychNazvu_SOP As Dictionary ' vyuziva se vsude v objektu SOP_class (Klient.listview4 a Klient_exe.listview4)
Private pCesta As Dictionary 'RISK-1114 - zatím se využívá pouze v mod_auto_tvorba_slozek_exekuce - èasem by se mìlo vyyužívat úplnì všude
''Private pFormatBunek As Scripting.Dictionary
Private pSlovnikPrehledPlateb_KTK As Dictionary

' WordoveKonstanty - v projektu neni reference, proto jako cisla: https://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
Private Const pwdFindContinue As Long = 1
Private Const pwdAlignParagraphJustify As Long = 3
Private Const pwdAlignParagraphCenter As Long = 1
Private Const pwdUnderlineSingle As Long = 1
Private Const pwdFindAsk As Long = 2
Private Const pwdReplaceAll As Long = 2

''Public Property Get formaty_Bunek() As Scripting.Dictionary
''    Set formaty_Bunek = pFormatBunek
''End Property

Public Property Get wb_Jez() As Workbook
    Set wb_Jez = pWbJez
End Property

Public Property Get cstDEFAULT_AMT_SOP_VALUE() As Long
    ' konstanta, ktera se vyuziva napric celym SOP_class, pri autorizacich a hodnotach nalezaciho rizeni
    ' je to hodnota, se kterou jezevcik vsude pracuje, pokud byl v databazi NULL
    ' nesmi to byt 0, protoze 0 taky vyzaduje autorizaci
    cstDEFAULT_AMT_SOP_VALUE = -141 ' bulharska konstanta, na hodnote by nemelo zalezet, jakekoliv ZAPORNE cislo
End Property

Public Property Get wdFindContinue() As Long
    wdFindContinue = pwdFindContinue
End Property

Public Property Get wdAlignParagraphJustify() As Long
    wdAlignParagraphJustify = pwdAlignParagraphJustify
End Property

Public Property Get wdAlignParagraphCenter() As Long
    wdAlignParagraphCenter = pwdAlignParagraphCenter
End Property

Public Property Get wdUnderlineSingle() As Long
    wdUnderlineSingle = pwdUnderlineSingle
End Property

Public Property Get wdFindAsk() As Long
    wdFindAsk = pwdFindAsk
End Property

Public Property Get wdReplaceAll() As Long
    wdReplaceAll = pwdReplaceAll
End Property


Public Property Get cst_missing_string_value() As String
    cst_missing_string_value = pcst_missing_string_value
End Property

Public Property Get cst_ucet_522() As String
    cst_ucet_522 = pcst_ucet_522
End Property

Public Property Get SlovnikPrehledPlateb_KTK() As Dictionary
    Set SlovnikPrehledPlateb_KTK = pSlovnikPrehledPlateb_KTK
End Property

Public Property Get SlovnikUzivatelskychNazvu_SOP() As Dictionary
    Set SlovnikUzivatelskychNazvu_SOP = pSlovnikUzivatelskychNazvu_SOP
End Property

Public Property Get Cesta() As Dictionary
    Set Cesta = pCesta
End Property


Private Sub Class_Initialize()

    Set pWbJez = ThisWorkbook
    '''   Inicializace jednotlivych slovniku
    Call NaplnSlovnikSOPy
    Call NaplnCestyKeSlozkam
'    Call NaplnFormatyBunek
    Call naplnSlovnikPrehledPlateb_KTK
End Sub


Private Sub NaplnSlovnikSOPy()
    ' slovnik ktery prevede nazvy sloupecku v tabuce SOP na jejich nazvy pro uzivatele
    ' vyuziva se zejmena na kartach klient.listview4 a klient_exe.listview4
    
    If pSlovnikUzivatelskychNazvu_SOP Is Nothing Then
        
        Set pSlovnikUzivatelskychNazvu_SOP = New Dictionary
        
        With pSlovnikUzivatelskychNazvu_SOP
        
            ' .Add "SKP_SOP", "SKP_SOP"
            ' .Add "ID_PERSON", "ID_PERSON"
            .Add "AMT_SOP1", "Výše SOP1"
            .Add "DATE_SOP1", "Datum SOP1"
            .Add "SOP_ADDITIONAL_FEE", "Paušál"
            .Add "AMT_ADDITIONAL_SOP1", "Výše doplatku SOP1"
            .Add "DATE_ADDITIONAL_SOP1", "Datum doplatku SOP1"
            .Add "AMT_RETURN_SOP1", "Pøeplatek (SOP1)"
            .Add "DATE_INFO_RETURN_SOP1", "Datum sdìlení pøeplatku (SOP1)"
            .Add "AMT_RETURN_REAL_SOP1", "Bylo vráceno (SOP1)"
            .Add "DATE_RETURN_REAL_SOP1", "Datum vrácení pøeplatku (SOP1)"
            .Add "AMT_SOP1_COURT", "Výše SOP1 urèena soudem"
            .Add "AMT_SOP2", "Výše SOP2"
            .Add "DATE_SOP2", "Datum SOP2"
            .Add "AMT_ADDITIONAL_SOP2", "Výše doplatku SOP2"
            .Add "DATE_ADDITIONAL_SOP2", "Datum doplatku SOP2"
            .Add "AMT_RETURN_SOP2", "Pøeplatek (SOP2)"
            .Add "DATE_INFO_RETURN_SOP2", "Datum sdìlení pøeplatku (SOP2)"
            .Add "AMT_RETURN_REAL_SOP2", "Bylo vráceno (SOP2)"
            .Add "DATE_RETURN_REAL_SOP2", "Datum vrácení pøeplatku (SOP2)"
            .Add "AMT_SOP2_COURT", "Výše SOP2 urèena soudem"
            .Add "SOP_ADDITIONAL_PAID", "Nenapárovaná platba SOP"
            .Add "TRAVEL_EXPENSES", "Cestovní náklady"
            .Add "DTIME_MODIFICATION", "Datum poslední modifikace"
            .Add "EMPLOYEE", "Provedl"
            
        End With
        
    
    Else
    
    End If
    
    'Call generate_SOP_columns
    
End Sub

Private Sub NaplnCestyKeSlozkam()
    ' Nechceme nikdy v Jezevèíku plnit cesty na pøímo, budeme to dìlat pøes tento slovník
    ' vsechny cesty k souborum a slozkam, se kterymi jezevcik pracuje
    ' vsechny cesty by meli koncit znakem "\", tedy pokud to neni primo nejakej soubor (vzory sablon)
    
    If pCesta Is Nothing Then
        
        Set pCesta = New Dictionary
        
        With pCesta
            
            .Add "Late", "W:\Vymahani\LATE\"
            
                '''   EPR
                .Add "EPR", .item("Late") & "EPR\"
                    .Add "EprPodaneZaloby", .item("EPR") & "PODANÉ ŽALOBY\"
                    .Add "EprNachystane", .item("EPR") & "NACHYSTANÉ\"
                    .Add "EprKlientiBezSlozky", .item("EPR") & "Klienti beze složky\"
                    .Add "EprDebety", .item("EPR") & "DEBETY\"
                    .Add "EprVzorPDF", .item("EPR") & "Vzor EPR\Nové vzory\EPR 5.4.2018.pdf"
                    
                    
                '''   EXE
                .Add "EXE", .item("Late") & "EXEKUCE\"
                    .Add "ExePodane", .item("EXE") & "PODANÉ\"
                        .Add "ExePodane_BezSlozky", .item("ExePodane") & "Klienti bez složky\"
                    
                    .Add "ExePripravene", .item("EXE") & "PØIPRAVENÉ\"
                        .Add "ExePripravene_BezSlozky", .item("ExePripravene") & "Klienti bez složky\"
                
                
                '''   INS
                .Add "INS", .item("Late") & "Insolvence\"
                    '.Add "InsNachystane", .item("INS") & "\"
                    .Add "InsPripravene", .item("INS") & "PØIPRAVENÉ\"
                    .Add "InsPrihlasene", .item("INS") & "PØIHLÁŠENÉ\"
                    .Add "InsVzory", .item("INS") & "Vzory\"
                        .Add "InsPrilohy", .item("InsVzory") & "Pøílohy\"
                        .Add "vzorInsPrihlaska", .item("InsVzory") & "nové vzory\Insolvenèní pøhláška od 26.10.2017.pdf"
                
                '''   DED
                .Add "DED", .item("Late") & "DÌDICKÉ\"
                .Add "DEDICOVE", .item("Late") & "DÌDICOVÉ\"
                
                '''   Vzory
                .Add "Vzory", .item("Late") & "Vzory\"
                    .Add "vzorInsXMLBase", .item("Vzory") & "XmlIpBase.xml"
                    .Add "vzorExeNavrh", .item("Vzory") & "Exekuèní návrh šablona jezevèík_2018_04_13.dotx"
                    .Add "vzorZadostDolozka", .item("Vzory") & "Žádost o doložku PM sablona NEMAZAT.dotx"
                    .Add "vzorPrihlaskaDoDR", .item("Vzory") & "vzor Pøihláška do DØ.dotx"
                    .Add "vzorPodani", .item("Vzory") & "Hlavièkový papír a petit šablona jezevèík NEMAZAT.docx"
                    .Add "vzorUznaniDluhu", .item("Vzory") & "Hlavièkový papír uznání dluhu NEMAZAT.dotx"
                    .Add "vzorEprXml", .item("Vzory") & "XmlEprBase.xml"
                    .Add "vzorSonikNBU", .item("Vzory") & "Soník NBÚ_VZOR_base NEMAZAT.dotx"
                    .Add "vzorSonikAutoOdpoved", .item("Vzory") & "Soník automatická odpovìï NEMAZAT.dotx"
                    
                '''   zesplatnujici dopis
                .Add "ZespDopis", .item("Late") & "Termination letters\" ' Pøedžalobní\ zatim se nevyuziva
                
                '''   OP + PPU + CENIK
                .Add "OP_PPU_CENIK", .item("Late") & "OP + PPU + Ceník\"
            
            
            .Add "TestovaciSlozka", NastavTestovaciSlozku()                ' adresa slozky, kam se budou ukladat vsechny sablony, pokud jsi na testu (odlisna pro Vymahani / Risk)
            
            
            
            
        End With

    
    End If


End Sub


Private Function NastavTestovaciSlozku() As String
    ' pointa je, ze chceme aby TESTovaci jezevcik ukladal sablony do jine slozky nez produkcni
    ' ta se ale jeste lisi pokud testujeme na risku a pokud testuji vymahaci
    
    If schema = "APP_COLL_JEZEVCIK_TEST" Then
        If InStr(UCase(Application.ActiveWorkbook.Path), UCase("W:\RISK")) > 0 Then
            NastavTestovaciSlozku = "W:\RISK\JEZEVÈÍK\Test\"
        ElseIf InStr(UCase(Application.ActiveWorkbook.Path), UCase("W:\Vymahani")) > 0 Then
            NastavTestovaciSlozku = Me.Cesta("Late") & "Testování Jezevèíka\sablony\"
        Else
            MsgBox "Jezevèík error jez-32123: Jezevèík nemùže uložit šablonu. Kontaktuj analytiky.", vbCritical
        End If
    End If

End Function


''Private Sub NaplnFormatyBunek()
''    ' Nektere sablony (zejmena export plateb pro uvery, KTK (casem i hypoteky)) vyuzivaji excel Sheets
''    ' formatovani bunek sheetu je ulozeno tady
''
''
''    'Dim PrehledPlatebRadekZesplatneni As CellFormat ' tohle je kravina. Cellformat neni objekt por Range ale pod Excel Application. Range nema jediny objekt s formatem, viz: https://forums.windowssecrets.com/showthread.php/76914-Store-a-range-of-formatting-in-a-variable-(Excel-VBA)
''
''
''    If pFormatBunek Is Nothing Then
''
''        Set pFormatBunek = New Scripting.Dictionary
''
''        With PrehledPlatebRadekZesplatneni
''            .Interior.ThemeColor = xlThemeColorLight1
''            .Interior.TintAndShade = 0.349986266670736
''            .Font.Color = vbWhite
''        End With
''
''        With pFormatBunek
''            .Add "Prehled_plateb_radek_zesp", PrehledPlatebRadekZesplatneni
''        End With
''
''    Else
''
''    End If
''
''End Sub


Private Sub naplnSlovnikPrehledPlateb_KTK()
    ' Tady by meli byt vsechny fixne textace ktere se tykaji Prehledu Plateb pro Kontokorent
    
    
    If pSlovnikPrehledPlateb_KTK Is Nothing Then
    
        Set pSlovnikPrehledPlateb_KTK = New Dictionary
            
        With pSlovnikPrehledPlateb_KTK
            
            ''' uvod - zakladni informace
            ' - levy sloupec
            .Add "info_klient", "Klient: "
            .Add "info_datum_uzavreni", "Datum uzavøení smlouvy/dodatku o Kontokorentu: "
            .Add "info_sazba", "Úroková sazba: "
            .Add "info_datum_cerpani", "Datum èerpání kontokorentu: "
            
            ' - pravy sloupec
            .Add "info_limit", "Výše limitu kontokorentu: "
            .Add "info_jistina_zesp", "Celková nesplacená jistina ke dni zesplatnìní: "
            .Add "info_jistina_zaloba", "Celková nesplacená jistina ke dni podání žaloby: "
            .Add "info_jistina_dnes", "Celková nesplacená jistina ke dni " & date & ": "
            
            
            ''' Telo tabulky
            .Add "tab_blokace", "Zùstatek ke dni blokace, který se skládá z: "
            .Add "tab_zesp", "Zùstatek ke dni zesplatnìní, který se skládá z: "
            .Add "tab_zaloba", "Zùstatek ke dni podání žaloby, který se skládá z: "
            .Add "tab_aktualni", "Aktuální zùstatek: "
            
        End With
        
    End If
    
End Sub
