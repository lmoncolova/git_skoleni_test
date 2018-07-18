VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Heirs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Hlavny objekt = pcolHeirs = je tzv. Strongly Typed Collection  -  to znamena, ze je to objekt Collection, v ktorom kazdy element MUSI byt objekt typu Heir.
' priklad o co ide: http://productivebytes.blogspot.cz/2012/04/how-to-create-strongly-typed-collection.html
' v tomto objekte je kolekcia odkazov na vsetkych dedicov vybraneho zemreleho, odkaz na notara k dedickemu rizeni a spisova znacka

' promenne pristupne pres public properties
Private pcolHeirs As Collection
Private plngId_person_dead As Long ''Private pobjZemrely As Client ' nebudeme vytvarat objekt, zbytocne, iba ulozime id_person
Private pobjZemrely As Client ' pridano Jez.v.5.3, 'RISK-973
Private pStrSpisovaZnacka As String
Private plngSkp_Inheritance As Long
Private pobjNotar As Notary
Private pdblTotalReceivable As Double
Private pdblTotalReceivablePaid As Double
' interni promenne pouze pro ucely modulu
Private plngSkp_notary As Long ' tahle promenna ma public property get az z objektu NOTARY - nechceme to na dvou mistach
'Private pblnNotaryFound As Boolean


Public Property Get colHeirs()
    Set colHeirs = pcolHeirs
End Property

Public Property Get lngId_person_dead()
    lngId_person_dead = plngId_person_dead
End Property

Public Property Get objZemrely()
    If pobjZemrely Is Nothing Then
        Set pobjZemrely = New Client
        Call pobjZemrely.initialize(plngId_person_dead)
    End If
    Set objZemrely = pobjZemrely
End Property

Public Property Get StrSpisovaZnacka()
    StrSpisovaZnacka = pStrSpisovaZnacka
End Property

Public Property Get lngSkp_Inheritance()
    lngSkp_Inheritance = plngSkp_Inheritance
End Property

Public Property Get dblTotalReceivable()
    dblTotalReceivable = pdblTotalReceivable
End Property

Public Property Get dblTotalReceivablePaid()
    dblTotalReceivablePaid = pdblTotalReceivablePaid
End Property



Public Property Get objNotar() As Notary
    ' pri prvnim dotazu na tenhle objekt se nainicializuje objekt Notary pod nim
    If pobjNotar Is Nothing Then ' 'true
        Set pobjNotar = New Notary
        Call pobjNotar.NacitajNotara(plngSkp_notary)
    End If
    
    ' if not pobjNotar.lngSkp_notary = 0
    ' if not plngSkp_notary = 0
    ' If pblnNotaryFound Then
    Set objNotar = pobjNotar
    'End If
End Property


Private Sub Class_Initialize()
    Set pcolHeirs = New Collection
    'Set pobjNotar = New Notary
End Sub

Private Sub Class_terminate()
    Set pcolHeirs = Nothing
    Set pobjNotar = Nothing
End Sub


'''''
'''''  nasledujuce 4 metody su zakladne metody pre collection
'''''

Property Get item(index As Variant) As Heir
' dedice ziskame bud jako dedici.item(cstr(skp_heir)), anebo dedici.item(1)
    Set item = pcolHeirs.item(index)
End Property

' nasledujuca SUB je klicova v tom, ze vstupni parametr je item AS HEIR  -- to je duvod proc se tomu rika strongly typed collection
Public Sub Add(item As Heir, Optional key As Variant)
    Call pcolHeirs.Add(item, key)
End Sub

Public Function count() As Long
    count = pcolHeirs.count
End Function

Public Sub remove(index As Variant)
    pcolHeirs.remove (index)
End Sub



Public Sub NajdiVsechnyDedice(id_person_dead As Long)
' hlavni SUB pro nacteni dedicu
Dim sql_heirs As String
Dim id_person_heir As Variant
Dim rst_person_heir As New ADODB.Recordset
Dim i As Integer
Dim dedic As Heir

'''   pseudokod:
' connection do tabulky dedicu where skp_heir
' pro kazdeho dedice (radek) se vytvori / inicializuje objekt HEIR (tam hlbsie sa budu inicializovat objekty Client)
' cez public property let asi nastavim odtialto aj hodnoty dalsich stlpcov v hlavnej tabulke skp_heir pre kazdeho dedice (aby nebol rovnaky dotaz aj tam zbytocne)
' kazdy takto vytvoreny objekt HEIR sa prida do pcolHeirs  (easy, iba pcolHeirs.add(objekt Heir))
' vysledok(vystup) by mal byt jeden objekt Heirs cez ktoreho bude pristup ku vsetkym dedicom klienta. V podstate jednoducha collection.


'pblnNotaryFound = False
plngId_person_dead = id_person_dead

'''   inicializace objektu Client pro KAZDEHO dedice
sql_heirs = "SELECT a.skp_heir, a.id_person_dead, a.Id_person_heir, a.amt_receivable, a.amt_receivable_paid, a.variable_symbol, a.flag_informovan_o_nahravani " & _
            "FROM " & schema & ".heirs a WHERE a.flag_deleted = 'N' AND a.id_person_dead = " & id_person_dead

Set rst1 = New ADODB.Recordset
rst_person_heir.Open sql_heirs, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
    If Not rst_person_heir.EOF Then
        
        For i = 1 To rst_person_heir.RecordCount
            
            ' overeni podminek
            If IsNull(rst_person_heir![id_person_heir]) Or IsNull(rst_person_heir![id_person_dead]) Then
                MsgBox "Jezev49k fatal error jez-551239" ' this should never happen
                Exit Sub
            End If
            
            ' Plneni promennych
            'plngId_person_dead = rst_person_heir![id_person_dead] ' KONINA JAK TRAKY! id_person_dead je vstupni argument, tady se hodi na 0 pokud nema dedice..
            
            Set dedic = New Heir
            
            If Not IsNull(rst_person_heir![amt_receivable]) Then dedic.dblAmtReceivable = rst_person_heir![amt_receivable]
            If Not IsNull(rst_person_heir![amt_receivable_paid]) Then dedic.dblAmtReceivablePaid = rst_person_heir![amt_receivable_paid]
            If Not IsNull(rst_person_heir![variable_symbol]) Then dedic.strVariabilniSymbol = rst_person_heir![variable_symbol]
            If Not IsNull(rst_person_heir![flag_informovan_o_nahravani]) Then dedic.strInformovan_o_nahravani = rst_person_heir![flag_informovan_o_nahravani]

            dedic.lngId_person_heir = rst_person_heir![id_person_heir]
            dedic.lngSkp_heir = rst_person_heir![skp_heir]
            dedic.strStavAktualni = ZjistiStavDedice(dedic.lngSkp_heir)
            
            ' !!!
            Call pcolHeirs.Add(dedic, CStr(rst_person_heir![skp_heir]))
            ' !!!
            
            rst_person_heir.MoveNext
        Next i
        
    Else
        'MsgBox "Žádný dìdic nenalezen."
    End If

Call NajdiNotare(plngId_person_dead)
Call VypocitejSpolecneCastky

rst_person_heir.Close



End Sub

Private Sub NajdiNotare(id_person_dead As Long)
' tahle SUB se dotazuje do tabulky INHERITANCE a naplni hodnoty plngSkp_notary a CODE
' samotny objekt notare se vytvori az pri prvnim dotazu (viz "public property get objNotar")

Dim sql_notar As String
Dim rst_notar As ADODB.Recordset

sql_notar = "SELECT i.skp_notary, i.code, i.skp_inheritance FROM " & schema & ".inheritance i " & _
            "WHERE i.flag_deleted != 'Y' AND i.id_person_dead = " & id_person_dead

Set rst_notar = New ADODB.Recordset
rst_notar.Open sql_notar, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
    If rst_notar.RecordCount = 1 Then
        
        If Not IsNull(rst_notar![skp_notary]) Then plngSkp_notary = rst_notar![skp_notary]
        pStrSpisovaZnacka = "" & rst_notar![code] & ""
        plngSkp_Inheritance = rst_notar![skp_inheritance]
        'pblnNotaryFound = True
        
    ElseIf rst_notar.RecordCount > 1 Then
        
        MsgBox "jez-91227: Bylo nalezeno více záznamù o notáøovi nebo spisové znaèce k tomuhle øízení. To není ok, pošli prosím printscreen analytikùm.", vbInformation, "chyba v databázi"
    
    Else
    
        'MsgBox "Notáø k tomuhle øízení nebyl nalezen."
        'pblnNotaryFound = False
        
    End If
    
rst_notar.Close



End Sub

Private Function ZjistiStavDedice(ByVal skp_heir As Variant) As String
' vstup: skp_heir
' vystup: ??? - zatim som dal text - nazov stavu
Dim sql_stav As String
Dim rst_stav As ADODB.Recordset

sql_stav = "SELECT a.name_heir_status FROM " & schema & ".v_heir_status_current a WHERE a.SKP_HEIR = " & CStr(skp_heir)

Set rst_stav = New ADODB.Recordset
rst_stav.Open sql_stav, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
    If rst_stav.RecordCount = 1 Then
        
        If Not IsNull(rst_stav![name_heir_status]) Then
            ZjistiStavDedice = rst_stav![name_heir_status]
        Else
            MsgBox "Jezevèík potøebuje operaci slepého støeva. Pošli printscreen analytikùm." 'nemelo by se nikdy stat
        End If
        
    ElseIf rst_stav.RecordCount = 0 Then
        ZjistiStavDedice = "(žádný stav)"
        
    Else
        MsgBox "Jezevèík potøebuje nový kardiostimulátor. Pošli printscreen analytikùm." 'nemelo by se nikdy stat
        
    End If
    
rst_stav.Close

End Function

Private Sub VypocitejSpolecneCastky()
'Spoèíta èástky které jsou sumou pøes dìdice

pdblTotalReceivable = 0
pdblTotalReceivablePaid = 0

For i = 1 To dedici.count
    
    pdblTotalReceivable = pdblTotalReceivable + dedici.item(i).dblAmtReceivable
    pdblTotalReceivablePaid = pdblTotalReceivablePaid + dedici.item(i).dblAmtReceivablePaid

Next i

End Sub

Public Sub Refresh()
' znovu nainicializuje vsetkych dedicov
' podstatne je aby pri tom zachovala hodnoty public premennych c,d,id_person,dedic, ...


'
'    Set dedici = New Heirs
'    Call dedici.NajdiVsechnyDedice(c.id_person)
End Sub


