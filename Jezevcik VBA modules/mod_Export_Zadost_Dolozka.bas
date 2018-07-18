Attribute VB_Name = "mod_Export_Zadost_Dolozka"
Option Explicit


Public Sub ExportZadostDolozka()

'RISK-983 - volá se z karty Klient.CommandButtonVyznaceniDolozky
Dim wordapp As Object
Dim objDoc As Object

Dim sql As String
'Dim StrOkresniSoud As String, Dim StrSoudAdresaUlice As String, StrSoudAdresaPscMesto As String, StrSoudIDSchranky As String
Dim StrSpisovaZnacka As String
Dim DblAmtClaim As Double
Dim i As Integer

Set wordapp = CreateObject("word.Application")
On Error GoTo chyba_otevreni_sablony
Set objDoc = wordapp.Documents.Add(Template:=cfg.Cesta("vzorZadostDolozka"))
On Error GoTo 0
wordapp.Visible = True



'Tyto atributy umím naplnit bez dodateèného sql dotazu nebo listview dotazu
Call NaplnBookmark("DatumNarozeni", c.Date_Birth, objDoc)
Call NaplnBookmark("Dnes", date, objDoc)
Call NaplnBookmark("Jmeno", c.name, objDoc, blnBold:=True)
Call NaplnBookmark("JmenoSpecialisty", JezUserLoggedIn.StrCivilName, objDoc)
Call NaplnBookmark("TrvalaAdresa", c.Address_P, objDoc)




''''''' RISK-1098 sql nahrazeno za objekt
''''OkresniSoud najdu v Klient.ListView6
'''For i = 1 To Klient.ListView6.ListItems.count
'''    Select Case Klient.ListView6.ListItems(i)
'''
'''        Case "Okresni soud"
'''            If Klient.ListView6.ListItems(i).SubItems(1) <> vbNullString Then
'''                StrOkresniSoud = Klient.ListView6.ListItems(i).SubItems(1)
'''            End If
'''
'''    End Select
'''Next i



If c.SOUD.okresni_soud = vbNullString Then
    HlaskaUzivateli "Klient nemá pøiøazen okresní soud"
End If


Call NaplnBookmark("OkresniSoud1", c.SOUD.okresni_soud, objDoc)
'Call NaplnBookmark("OkresniSoud2", c.SOUD.okresni_soud, objDoc) 'Vymazano na dodatecnou zadost Evy Rajsiglove

'Spisovou znaèku a amt claim najdu v Klient.ListView5

For i = 1 To Klient.ListView5.ListItems.count
    Select Case Klient.ListView5.ListItems(i)
    
        Case "Spisová znaèka"
            If Klient.ListView5.ListItems(i).SubItems(1) <> vbNullString Then
                StrSpisovaZnacka = Klient.ListView5.ListItems(i).SubItems(1)
            End If
        Case "Výše pohledávky"
            If Klient.ListView5.ListItems(i).SubItems(1) <> vbNullString Then
                DblAmtClaim = Klient.ListView5.ListItems(i).SubItems(1)
            End If
    End Select
Next i

If StrSpisovaZnacka = vbNullString Then
    HlaskaUzivateli "Klient nemá vyplnìnou spisovou znaèku"
End If

If DblAmtClaim = 0 Then
    HlaskaUzivateli "Klient nemá vyplnìnou výši pohledávky"
End If



Call NaplnBookmark("SpisZn1", StrSpisovaZnacka, objDoc)
'Call NaplnBookmark("SpisZn2", StrSpisovaZnacka, objDoc) 'Vymazano na dodatecnou zadost Evy Rajsiglove
Call NaplnBookmark("Castka", format(DblAmtClaim, "#,##0.00"), objDoc)


''''''' RISK-1098 sql nahrazeno za objekt
''''Údaje o soudu si dotáhnu na základì názvu soudu a napojení do tabulky SOUDY
'''sql = "select DISTINCT ADRESA_P2, ADRESA_P3, ID_SOUD from " & schema & ".SOUDY WHERE OKRESNI_SOUD = '" & c.SOUD.okresni_soud & "'"
'''Set rst1 = New adodb.Recordset
'''rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''    If Not rst1.EOF Then
'''        StrSoudAdresaUlice = "" & rst1![adresa_P2] & ""
'''        StrSoudAdresaPscMesto = "" & rst1![adresa_P3] & ""
'''        StrSoudIDSchranky = "" & rst1![ID_soud] & ""
'''
'''    Else
'''
'''    HlaskaUzivateli "Adresa soudu nenalezena"
'''
'''    End If
'''rst1.Close

Call NaplnBookmark("SoudAdresaUlice", c.SOUD.adresa_P2, objDoc)
Call NaplnBookmark("SoudAdresaPscMesto", c.SOUD.adresa_P3, objDoc)
Call NaplnBookmark("SoudIDSchranky", c.SOUD.ID_soud, objDoc)


Call VymazBookmarks(objDoc)
Call ZarovnejCelyDokumentDoBloku(objDoc) ' jez.v.5.8
Call UlozDokument(objDoc)

Set wordapp = Nothing
Set objDoc = Nothing
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True

Call VypisVsechnyHlasky

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



Private Sub UlozDokument(ByRef WordDocument As Variant)
'RISK-983

Dim AFolder As String, BFolder As String
Dim fso As Object

Dim BlnCistyDebetar As Boolean, sql As String 'True znaèí, že se má uložit do Podaných Žalob

'Pokud se jedná o èistého debetáøe, tak uložím do jedné složky, pokud má i úvìr tak do jiné, pokud nenajdu složku tak jen otevøu

sql = " SELECT b.NAME_PERSON_STATUS " & _
      " from " & schema & ".person_status_data a " & _
      " JOIN " & schema & ".person_status b ON a.SKP_PERSON_STATUS = b.SKP_PERSON_STATUS " & _
                                            " AND a.flag_valid = 'Y' AND a.flag_deleted = 'N' and a.id_person = " & c.id_person  ' RISK-985 SKP_PERSON_STATUS
 
 Set rst1 = New ADODB.Recordset
 rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
 
 If Not rst1.EOF Then
    If rst1![NAME_PERSON_STATUS] = "DEBET" Then
        BlnCistyDebetar = True
    End If
 End If
 rst1.Close

AFolder = cfg.Cesta("EprPodaneZaloby") & c.Nazev_slozky_bez_lomitka 'Úvìráø
BFolder = cfg.Cesta("EprDebety") & c.Nazev_slozky_bez_lomitka 'Èistý debetáø

Set fso = CreateObject("Scripting.FileSystemObject")

If schema = "APP_COLL_JEZEVCIK_TEST" Then
    WordDocument.SaveAs (cfg.Cesta("TestovaciSlozka") & "\Žádost o doložku PM " & c.name & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".docx")
Else
    
    If (Not BlnCistyDebetar) And fso.FolderExists(AFolder) Then
        WordDocument.SaveAs (AFolder & "\Žádost o doložku PM " & c.name & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".docx")
        Call Shell("explorer.exe " & AFolder, vbNormalFocus)
    ElseIf BlnCistyDebetar And fso.FolderExists(BFolder) Then
        WordDocument.SaveAs (BFolder & "\Žádost o doložku PM " & c.name & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".docx")
        Call Shell("explorer.exe " & BFolder, vbNormalFocus)
    Else
        MsgBox "Složka pro klienta " & c.name & " neexistuje. Ulož si prosím dokument svépomocí.", vbInformation
    End If

End If


End Sub
