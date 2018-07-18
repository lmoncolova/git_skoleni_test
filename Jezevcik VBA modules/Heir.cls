VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Heir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' objekt Heir = vazba mezi 2 osobami: dedicem a zemrelym. (a notarem)
' Jeden objekt Heir odpoveda jednemu skp_heir v hlavnej tabulke heir.
' tenhle excelovy objekt by idealne mel zrkadlit strukturu tabulek v databazi. T.j. obsahuje pouze odkazy na dedice, zemreleho, notare.


' vsetky premenne zacinaju skratkou pre datovy typ pouzity, je to good practice. Tabulka skratiek: https://msdn.microsoft.com/en-us/library/aa240858(v=vs.60).aspx

Private pobjDedic As Client
'Private pobjZemrely As Client ' odkaz na zemreleho tady zadny neni, pouze v objektu HeirS
Private plngSkp_heir As Long
Private plngId_person_heir As Long


' tady jsou pouze informace o dedicovi ktere nejsou v objektech klienta - duplikaci totiz nechceme
Private pdblAmtReceivable As Double
Private pdblAmtReceivablePaid As Double
Private pstrVariabilniSymbol As String
Private pstrInformovan_o_nahravani As String
Private pstrStavAktualni As String


' uzivatel otvori dedicke rizeni
' dblclick na zemreleho, nic extra sa nedeje
' az potom sa zavola:
'    Dim dedici As Heirs
'    Set dedici = New Heirs
'    Call dedici.NajdiVsechnyDedice(id_person_dead)
' Heirs.NajdiVsechnyDedice spravi nasledujuce:
'    pozrie sa do hlavnej tabulky HEIRS
'    pre daneho zemreleho najde vsetkych dedicov
'    pre kazdeho dedica vytvori objekt Heir a naplni v nom hodnoty (VS, pohledavka, ...)
'    objekt Heir prida do seba (collection).
' Vysledkom je, ze mame otvorenu kartu Dedicke rizeni a klienta v nej + mame objekt dedici (Heirs), ktory obsahuje agregovane info ku vsetkym dedicom toho klienta


Public Property Get objDedic() As Client
    ' pri prvnim dotazu na tenhle objekt se nainicializuje objekt Client pod nim
    If pobjDedic Is Nothing Then
        Call Me.InicializujObjektKlienta
    End If
    
    Set objDedic = pobjDedic
End Property



Public Property Get dblAmtReceivable() As Double
    dblAmtReceivable = pdblAmtReceivable
End Property

Public Property Get dblAmtReceivablePaid() As Double
    dblAmtReceivablePaid = pdblAmtReceivablePaid
End Property

Public Property Get strVariabilniSymbol() As String
    strVariabilniSymbol = pstrVariabilniSymbol
End Property

Public Property Get lngId_person_heir() As Long
    lngId_person_heir = plngId_person_heir
End Property

Public Property Get lngSkp_heir() As Long
    lngSkp_heir = plngSkp_heir
End Property

Public Property Get strInformovan_o_nahravani() As String
    strInformovan_o_nahravani = pstrInformovan_o_nahravani
End Property

Public Property Get strStavAktualni() As String
    strStavAktualni = pstrStavAktualni
End Property

Public Sub InicializujObjektKlienta()
    ' funkce nainicializuje objekt Client skryty pod dedicem
    ' pak me.objDedic predstavuje objekt klienta tak jak se vsude v jezevcikovi vyuziva pod skratkou "c"
    
    Set pobjDedic = New Client
    pobjDedic.initialize (Me.lngId_person_heir)
    
End Sub

''''''    PROPERTIES LET
' property LET nastavi hodnotu vybranej premennej

Public Property Let dblAmtReceivable(ByVal AmtReceivable As Double)
    pdblAmtReceivable = AmtReceivable
End Property

Public Property Let dblAmtReceivablePaid(ByVal dblAmtReceivablePaid As Double)
    pdblAmtReceivablePaid = dblAmtReceivablePaid
End Property

Public Property Let strVariabilniSymbol(ByVal strVariabilniSymbol As String)
    pstrVariabilniSymbol = strVariabilniSymbol
End Property

Public Property Let lngId_person_heir(ByVal lngId_person_heir As Long)
    plngId_person_heir = lngId_person_heir
End Property

Public Property Let lngSkp_heir(ByVal lngSkp_heir As Long)
    plngSkp_heir = lngSkp_heir
End Property

Public Property Let strInformovan_o_nahravani(ByVal strInformovan_o_nahravani As String)
    pstrInformovan_o_nahravani = strInformovan_o_nahravani
End Property

Public Property Let strStavAktualni(ByVal strStavAktualni As String)
    pstrStavAktualni = strStavAktualni
End Property




'''
'''Public Sub NajdiInfoKDedici(Id_person_heir As Variant)
'''Dim sql_heir As String
'''' NEPOTREBNE tak to vyzera, moze sa vyhodit..
'''' tohle je hlavni funkce (initialize nemuze mit vstupni argument)
'''' k dedicovi sa nacitaju informacie a hlavnej tabulky HEIRS a nainicializuje sa objekt
'''
''''''   pseudokod
'''' dotahni info k dedici a to skrz objekt Client a skrz jeden sql dotaz (pre premenne ktore nie su na klientovi)
'''' uloz to vsetko do objDedic
''''
'''
'''sql_heir = "SELECT a.skp_heir, a.id_person_dead, a.Id_person_heir, a.skp_notary, a.vyse_pohledavky, a.variabilni_symbol, a.flag_informovan_o_nahravani " & _
'''            "FROM " & schema & ".heirs a WHERE a.flag_deleted = 'N' AND a.Id_person_heir = " & Id_person_heir
'''
'''Set rst1 = New ADODB.Recordset
'''
'''rst1.Open sql_heir, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''    If rst1.RecordCount = 1 Then
'''        plngSkp_heir = rst1![skp_heir]
'''
'''    Else
'''        ''' nema zaznam nebo ma vice zaznamu
'''        MsgBox "87965454873345"
'''    End If
'''rst1.Close
'''
'''
'''End Sub


