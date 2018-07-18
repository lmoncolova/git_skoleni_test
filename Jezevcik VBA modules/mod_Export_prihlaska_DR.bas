Attribute VB_Name = "mod_Export_prihlaska_DR"
Option Explicit
Private textace As Scripting.Dictionary


Public Sub vytvor_prihlasku_do_DR()
    
    Dim wordapp As Object
    Dim objDoc As Object
    
    Call VynulujPremenne
    
    Set wordapp = CreateObject("word.Application")
    
    Debug.Assert Not cfg Is Nothing ' call jj()
    Debug.Assert Len(cfg.Cesta("vzorPrihlaskaDoDR")) > 2
    
    On Error GoTo chyba_otevreni_sablony
        Set objDoc = wordapp.Documents.Add(Template:=cfg.Cesta("vzorPrihlaskaDoDR"))
    On Error GoTo 0
    
    Set textace = NaplnTextace()
    
    wordapp.Visible = True
    wordapp.Activate
    
    
    ''''   hlavni kod pro plneni sablony
    
    Call VypisVsechnyBookmarks(objDoc) ' notar_ID, notar_radek1, notar_radek2, notar_radek3, podpis_vymahace, uvod_spisova_znacka, vazeny_pane
'    Call NaplnBookmark("RamcovaSmlouva", Empty_string_highlighter(c.TextContractNumber), objDoc)
'    Call NaplnBookmark("Povinny_NameFull", Empty_string_highlighter(c.name), objDoc, blnBold:=True)
    
    NaplnBookmark "notar_ID", dedici.objNotar.strIdDS, objDoc
    NaplnBookmark "notar_radek1", dedici.objNotar.strName, objDoc
    NaplnBookmark "notar_radek2", Left(dedici.objNotar.strAddress, InStr(dedici.objNotar.strAddress, ",") - 1), objDoc
    NaplnBookmark "notar_radek3", Trim(Right(dedici.objNotar.strAddress, Len(dedici.objNotar.strAddress) - InStr(dedici.objNotar.strAddress, ","))), objDoc
    NaplnBookmark "uvod_spisova_znacka", dedici.StrSpisovaZnacka, objDoc
    NaplnBookmark "podpis_vymahace", JezUserLoggedIn.strSignature, objDoc
    
    Call VyplnTeloPrihlasky(objDoc)
    
    '''   ulozeni
    
    Call VymazBookmarks(objDoc)
    Call ZarovnejCelyDokumentDoBloku(objDoc) ' jez.v.5.8
    Call UlozPrihlaskuDR(objDoc)
    wordapp.Visible = True
    
    Call VypisVsechnyHlasky
    
    Set objDoc = Nothing
    Set wordapp = Nothing
    Exit Sub
    
chyba_otevreni_sablony:

    If (Err.Number = 13 And Err.Description = "Type mismatch") Or (Err.Number = 5981) Or (Err.Description = "Application-defined or object-defined error") Then
        MsgBox "jez-00013: Vypadá to, že nemáš povolené zmìny ve vzoru šablony. Je potøeba šablonu otevøít a zakliknout povolení úprav (v žluté lištì navrchu). " & vbCrLf & _
               "Jedná se o tenhle soubor: " & cfg.Cesta("vzorPrihlaskaDoDR") & vbCrLf & _
               "> v pøípadì, že šablona (.docx soubor) nelze spustit, je potøeba pøes správce úloh ukonèit všechny procesy WINWORD.EXE.", vbInformation, "Povolení úprav šablony"
    Else
        MsgBox "Došlo k neznámé chybì pøi naèítání vzoru šablony."
    End If
    
    Set objDoc = Nothing
    Set wordapp = Nothing
        
End Sub


Private Sub UlozPrihlaskuDR(ByRef WordDocument As Variant)
    ' Sub ulozi Prihlasku do slozky klienta nebo do testovaci slozky
    ' pokud slozka klienta neexistuje, vytvori se
    
    Dim AFolder As String
    Dim fileName As String
    
    
    fileName = "Pøihláška do DØ klienta " & c.name & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".docx"
    AFolder = cfg.Cesta("DED") & c.Nazev_slozky
    
    
    If schema = "APP_COLL_JEZEVCIK_TEST" Then
        WordDocument.SaveAs (cfg.Cesta("TestovaciSlozka") & fileName)
        'Call Shell("explorer.exe " & cfg.Cesta("TestovaciSlozka"), vbNormalFocus) 'otvirani slozky na testu
    Else
        If CreatePath(AFolder) Then
            WordDocument.SaveAs (AFolder & fileName)
            Call Shell("explorer.exe " & AFolder, vbNormalFocus)
        Else
            'MsgBox "Dokument se neuložil, neexistuje cesta " & BFolder, "Dokument neuložen", vbInformation
            assertion 1 = 0, "jez-88232: neznámý problém s uložením souboru: " & AFolder & fileName
            'WordDocument.SaveAs (BFolder & fileName)
            'Call Shell("explorer.exe " & BFolder, vbNormalFocus)
        End If
    
    End If
    


End Sub


Private Function NaplnTextace() As Scripting.Dictionary
    
    ' tady jsou vsechny stringy se kterymi pak kod pracuje
    ' rimske cislice oznacuji odsek ve kterem se nachazeji
    Dim klient_full_name As String
    
    Set NaplnTextace = New Scripting.Dictionary
    klient_full_name = IIf(c.Gender = "Z", "pí.", "p.") & c.Text_Title_before & c.name & c.Text_Title_After
    
    With NaplnTextace
        
        '.Add "-", " -   "
        .Add "radek", Chr(10) '& Chr(12)
        
        '''   textace ktere se nikdy nemeni
        .Add "osloveni", IIf(InStr(UCase(dedici.objNotar.strName), UCase("Mgr")) > 0, "Vážený pane magistøe", "Vážený pane doktore")
        
        .Add "odst_I", "ve vìci dìdictví po " & klient_full_name & _
                               ", r.è. " & c.TextBirthNumberLomitko & ", " & _
                               IIf(c.Gender = "Z", "zemøelé", "zemøelém") & " dne " & c.DateDeath & ", " & _
                               "poslední známá adresa " & c.Address_P & ", " & _
                               "pøihlašujeme do dìdického øízení pohledávku spoleènosti Air Bank a.s., IÈ 29045371, se sídlem Evropská 2690/17, 160 00 Praha 6."
                               
        .Add "odst_II", IIf(c.Gender = "Z", "Zemøelá ", "Zemøelý ") & klient_full_name & " a Air Bank a.s. uzavøeli dne " & c.Date_Contract_Signed & _
                            " Rámcovou smlouvu è. " & c.TextContractNumber & ", na základì které byl " & IIf(c.Gender = "Z", "klientce ", "klientovi ") & _
                            "následnì aktivován " & vyjmenuj_ucty_klienta() & ". Podpisem Rámcové smlouvy byla souèasnì uzavøena smlouva o úvìru " & _
                            "_______......._______"
        
        .Add "odst_III_pohl_uvod", "Pohledávka vyplývající ze smlouvy o hotovostním úvìru è. "
        .Add "odst_III_pohl_uvod", "Pohledávka vyplývající ze smlouvy o hotovostním úvìru è. "
        
        
        
        
        If c.Gender = "Z" Then
            
        Else
            
        End If
        
        
    End With
    
End Function

Public Function vyjmenuj_ucty_klienta() As String
    ' funkce udela loop pres vsechny ucty klienta a vrati jediny string, kde jsou vyjmenovane po poradi
    ' konkretni forma textace byla puvodne nastavena pro pro sablonu Prihlaska do Dedickeho Rizeni
    ' TODO: presunut do Public_functions, pripadne primo jak property objektu Client
    
    Const str_BU As String = "BÚ"
    Const str_SU As String = "SÚ"
    Const text_BU As String = " bìžný úèet è. "
    Const text_SU As String = " spoøící úèet è. "
    Const str_za_cislem_uctu As String = "/3030 vedený u Air Bank a.s."
    Dim result As String
    Dim str_spojka As String
    Dim i As Integer
    
    
    '''   podminky
    
    Debug.Assert Not c Is Nothing
    
'    If (c.GetAccounts Is Nothing) Then
'        Call c.find_accounts
'    End If
    
    
    '''   samotnej loop ktery nahaze ucty do jednoho stringu
    str_spojka = ""
    
    If c.GetCnt_Accounts > 0 Then
        For i = 1 To c.GetCnt_Accounts
            
            '''   nastaveni spojky pokud je uctu vic
            If i > 1 Then
                If i = c.GetCnt_Accounts Then
                    str_spojka = " a"
                Else
                    str_spojka = ","
                End If
            End If
            
            '''   vyplneni cisla uctu
            If c.GetAccounts(i).Account_type = str_BU Then
                result = result & str_spojka & text_BU & c.GetAccounts(i).Text_bank_account_number & str_za_cislem_uctu
            ElseIf c.GetAccounts(i).Account_type = str_SU Then
                result = result & str_spojka & text_SU & c.GetAccounts(i).Text_bank_account_number & str_za_cislem_uctu
            Else
                HlaskaUzivateli "Klient má kromì bìžných a spoøících úètù i úèet typu " & c.GetAccounts(i).Account_type & ". Ten se do druhého odstavce nepropsal."
            End If
        Next i
    End If
    
    vyjmenuj_ucty_klienta = result
    
End Function

Private Sub VyplnTeloPrihlasky(ByRef WordDocument As Variant)
    
    'NaplnBookmark "vazeny_pane", dedici.objNotar.strName, WordDocument
    
End Sub



Private Sub VynulujPremenne()
    
    strMsgbox = ""
    
    
End Sub
