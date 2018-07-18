Attribute VB_Name = "mod_PDF_Sendkeys"
Option Explicit

Private PDFViewerApp As Variant
'Const cst_EPR_Vzor_path As String = "W:\Vymahani\LATE\EPR\Vzor EPR\Nové vzory\EPR 5.4.2018 – kopie.pdf"

'Const cst_EPR_XML_Vzor_path As String = "eprXMLVzor.XML"
Const cstStrPathToAttachments As String = "W:\RISK\JEZEVÈÍK\Vzory\test prilohy"
Const cstStrPathAdobeReaderDC As String = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"


' musi byt anglicka klavesnice nastavena jako defaultni
' bylo by fajn znemoznit, ze uzivatel nadela bordel v jezevcikovi.. Odinicializovat klienta, nebo ho vyhodit z dwh na chvili, nebo alespon otevrit a rollbacknout transakci.
' podobne by nebylo vubec na skodu zavrit jine aplikace (mail, obs, vsecko..)
' VZDY PRACOVAT JENOM S NEPODEPSANYMI EPR - aby se vec nemohla odeslat

' TADY borec resi problem (prepinani mezi aplikacemi): https://www.mrexcel.com/forum/excel-questions/961020-after-activating-another-window-macro-stops-running-when-required-appactivate-excel-true-until-excel-window-manually-clicked-then-continues.html

' uzivatel klikne na button v Jez
' vygeneruje sa XML
' potom sa otvori vzor EPR (public premenna s cestou)
' ulozi sa do klientovej zlozky s jasnym nazvom ze to je automaticky generovane + dtime
'  1a. vlozi vsetky subory v priecinku Prilohy
'  1b. opyta sa uzivatela na prilohy ktore chce vlozit
' ulozit



Public Sub test_odhlaseni_a_prihlaseni()
    logout
    Main.CommandButton1_Click
End Sub


Public Sub Vytvor_PDF_komplet(ByVal rizeni As typRizeni)
    ' HLAVNI SUB
    ' tady se postupne volaji funkce, ktere otevrou Adobe, naplni prilohy, nahrajou XML a ulozi jej.
    ' tady se nastavuji asi vsechny paths
    ' vychytane je zatim pouze pro EPR.
    ' Insolvencni sablonu by nemel byt problem zcela jednoduse doplnit, nicmene je potreba znovu vychytat pocet {TAB} pro vkladani priloh a pod.
    
    

    Dim strClientFolder As String
    Dim strCestaPDFSablona As String
    Dim strNazevXML As String
    Dim strNazevSlozkyPrilohy As String
    Dim fso As Scripting.FileSystemObject
    
    
    '''   Nastavi se slozka klienta dle toho, v jakem sme rizeni
    
    Select Case rizeni
        
        Case typRizeni_EPR
            
            strCestaPDFSablona = cfg.Cesta("EprVzorPDF")
            If schema = "APP_COLL_JEZEVCIK" Then
                strClientFolder = cfg.Cesta("EprPodaneZaloby")
            Else
                strClientFolder = cfg.Cesta("TestovaciSlozka")
            End If
            
        Case typRizeni_INS
        
            strCestaPDFSablona = cfg.Cesta("vzorInsXMLBase")
            If schema = "APP_COLL_JEZEVCIK" Then
                strClientFolder = cfg.Cesta("InsPripravene") '' xx asi spatna slozka
                Debug.Assert False ' jaka slozka tam ma byt?
            Else
                strClientFolder = cfg.Cesta("TestovaciSlozka")
            End If
                
        Case Else
            assertion 1 = 0, "jez-77369: neznámý typ øízení"
            Exit Sub
        
    End Select

    strClientFolder = strClientFolder & c.Nazev_slozky
    
    Set fso = New FileSystemObject
    If Not fso.FolderExists(strClientFolder) Then
        MsgBox "Nepovedlo se najít klientovu složku: " & strClientFolder & "." & vbCrLf & "Konec magie."
        Exit Sub
    End If
    
    
    '''   tady se zacina chaos, vytvari se antihmota a deli se nulou
    Call TransBegin
        
        Call otevri_PDF_Viewer(strPathToPDFfile:=strCestaPDFSablona)
        
        
        '''   IMPORT XML
        
        strNazevXML = FindFileInFolder(strClientFolder, ".xml")
        If fso.FileExists(fso.BuildPath(strClientFolder, strNazevXML)) Then
            Call importuj_XML(strClientFolder, strNazevXML)
        '    Call importuj_XML("W:\Vymahani\LATE\Testování Jezevèíka", _
        '                      "eprXMLVzor.xml")
        End If
        ' ulozi sa do klientovej zlozky s jasnym nazvom ze to je automaticky generovane + dtime
        
        
        '''   PRILOHY
        
        '  1a. vlozi vsetky subory v priecinku Prilohy
        '  1b. opyta sa uzivatela na prilohy ktore chce vlozit
        
        strNazevSlozkyPrilohy = strClientFolder & "Pøílohy\"
        If fso.FolderExists(strNazevSlozkyPrilohy) Then
            Call VlozPrilohyDoEPR(strPathToAttachmentsFolder:=strNazevSlozkyPrilohy)
        Else
            'HlaskaUzivateli "Nenalezena složka s pøílohami: " & strNazevSlozkyPrilohy
            ' zatial nepouzivam Hlasky..
        End If
        
        
        '''   ULOZENI
        
        Call Uloz_PDF_jako_2(strClientFolder, "EPR auto " & c.CUID & " " & Replace(Replace(Now, ":", "-"), " ", "_") & ".pdf")
        
        
        ' ulozit zmeny
        ' Debug.Assert False
        ' SendKeys "^s"
        MsgBox "Hotovo. Konec magie."
        
    Call TransRollback
    
End Sub


Private Sub aktivuj_PDF_Viewer(Optional ByVal viewerName As String = "Adobe Acrobat Reader DC")
    
    AppActivate Application.Caption, False ' tady kdyz se da True, tak to dela bordel (kod se zasekne a spusti az pri kliknuti na excel)
    AppActivate viewerName, True
    Application.Wait Now + #12:00:01 AM#
    ' SendKeys "%sj", True ' pro test
End Sub


Private Sub otevri_PDF_Viewer(ByVal strPathToPDFfile As String)
' Sub otevre PDF soubor v defaultnim prohlizeci (testovano pouze pro Adobe Acrobat Reader DC)
' pokud je PDF viewer jiz otevren, pouze se aktivuje
' vstup: cesta k PDF souboru

    On Error GoTo ViewerNotFound
        Call aktivuj_PDF_Viewer
        Exit Sub
    
ViewerNotFound:
    On Error GoTo 0
    
    If 1 = 0 Then
        PDFViewerApp = Shell(cstStrPathAdobeReaderDC, vbNormalFocus)
    Else
        ActiveWorkbook.FollowHyperlink (strPathToPDFfile)
    End If
    
End Sub




Private Sub VlozPrilohyDoEPR(ByVal strPathToAttachmentsFolder As String)
    Dim i As Integer
    Dim fso As FileSystemObject
''    Dim arrPrilohy() As Scripting.file
    Dim Priloha As Scripting.file
    Dim strPrilohaName As String
    Dim strPrilohaParentName As String
    Dim strPathAttFolder As String
    strPathAttFolder = strPathToAttachmentsFolder
    
    Set fso = New FileSystemObject
    
    Call aktivuj_PDF_Viewer
    
    
    For i = 1 To 20
        SendKeys "{TAB}", True
    Next i
    
    
    ''Debug.Print PDFViewerApp
    
    ' Enter na vlozenie priloh, ak mame viac ako 0
    SendKeys "{ENTER}", True
    SendKeys "{TAB}", True
    SendKeys "{TAB}", True
    
''    ReDim arrPrilohy(0 To fso.GetFolder(strPathAttFolder).Files.count)
''    arrPrilohy() = Array(fso.GetFolder(strPathAttFolder).Files)
    'For Each Priloha In fso.GetFolder(cstStrPathToAttachments).Files
    'For Each Priloha In fso.GetFolder(strPathAttFolder).Files
    For Each Priloha In fso.GetFolder(strPathAttFolder).Files
        SendKeys "{ENTER}", True
        strPrilohaName = Priloha.name
        strPrilohaParentName = Priloha.ParentFolder
        Application.Wait Now + #12:00:02 AM# ' tady 1 vterina nestaci
        ''SendKeys "Info pøílohy.pdf", True
        SendKeys strPrilohaName, True
        
        SendKeys "{F4}", True
        SendKeys "^a", True
        SendKeys strPrilohaParentName, True
        Application.Wait Now + 1 / (24 * 60 * 60# * 2)
        SendKeys "{ENTER}", True
'        For i = 1 To 6
'            SendKeys "{TAB}", True
'        Next i
        Application.Wait Now + #12:00:01 AM#
        SendKeys "{ENTER}", True
        SendKeys "{ENTER}", True
        SendKeys "{ENTER}", True
        Application.Wait Now + #12:00:01 AM#
        
        
'''        SendKeys "{ENTER}", True
'''        Application.Wait Now + #12:00:01 AM#
'''        SendKeys "{F4}", False
'''        SendKeys "^a", False
'''        SendKeys "W:\Vymahani\LATE\EPR\Vzor EPR", False
'''        SendKeys "{ENTER}", True
'''        For i = 1 To 6
'''            SendKeys "{TAB}", True
'''        Next i
'''        Application.Wait Now + #12:00:01 AM#
'''        SendKeys "Info pøílohy.pdf", False
'''        SendKeys "{ENTER}", True
        
        
    Next Priloha
    
    SendKeys "{ENTER}", True
    SendKeys "{ESC}", True
    SendKeys "{ESC}", True
    SendKeys "{ESC}", True
    SendKeys "{ESC}", True
    SendKeys "{ESC}", True
    
End Sub





Private Sub importuj_XML(ByVal Cesta As String, ByVal nazev As String)
    ' asi nefunguje pokud je v ceste znak "_"
    Dim i As Integer
    Dim strCesta As String
    Dim strNazev As String
    
    strCesta = Cesta
    strNazev = nazev
    
'    On Error Resume Next
'        Call aktivuj_PDF_Viewer
'    On Error GoTo 0
    
    Application.Wait Now + #12:00:01 AM#
    SendKeys "{ESCAPE}", True
    SendKeys "{ESCAPE}", True
    Application.Wait Now + #12:00:01 AM#
    
    SendKeys "{ESCAPE}", True
    
    SendKeys "%p", True
    For i = 1 To 4
        SendKeys "{UP}", True
    Next i
    SendKeys "{RIGHT}", True
    SendKeys "{ENTER}", True
    
    SendKeys strNazev, True
    Application.Wait Now + #12:00:01 AM#
    Call ZadejCestuANazevSouboru(strCesta, strNazev)
    Application.Wait Now + #12:00:02 AM# ''' nacteni XML muze trvat..

End Sub




Private Sub Uloz_PDF_jako_2(ByVal Cesta As String, ByVal nazev As String)
    Dim strCesta As String
    Dim strNazev As String
    Dim i As Integer
    
    strCesta = Cesta
    strNazev = nazev
    
    SendKeys "%s", True
    Application.Wait Now + #12:00:01 AM#
    SendKeys "j", True
    Application.Wait Now + #12:00:01 AM#
    
    Call ZadejCestuANazevSouboru(strCesta, strNazev)
    
End Sub




Private Sub ZadejCestuANazevSouboru(ByVal Cesta As String, ByVal nazev As String)
    
    Dim i As Integer
    Dim strCesta As String
    Dim strNazev As String
    
    strCesta = Cesta ' je potreba, jinak se stringy rozsypou
    strNazev = nazev ' je potreba, jinak se stringy rozsypou
    
    ''' set path to folder
    SendKeys "{F4}", True
    SendKeys "^a", True
    SendKeys "^a", True
    '"W:\Vymahani\LATE\Testování Jezevèíka\EPR_auto"  '  priecinok klienta kam sa ma ulozit EPR
    SendKeys strCesta, True
    Application.Wait Now + #12:00:01 AM#
    'Application.Wait Now + 1 / (24 * 60 * 60# * 2)
    SendKeys "{ENTER}", True


    For i = 1 To 6
        SendKeys "{TAB}", True
    Next i
    Application.Wait Now + #12:00:01 AM#
    SendKeys strNazev, True

    
    SendKeys "{ENTER}", True
    Application.Wait Now + #12:00:01 AM#
    SendKeys "{ENTER}", True

End Sub





Private Sub vyberSuboryManualne()
    '''   ASI SE NEVYUZIVA
    
    Dim myFiles As Variant
    Dim file As Scripting.file

'    myFiles = Application.GetOpenFilename(MultiSelect:=True)
'
'    For Each file In myFiles
'        Debug.Print file
'    Next file
'
    
    For Each file In Application.GetOpenFilename(MultiSelect:=True)
        Debug.Print file.name
        Debug.Print file.ParentFolder
    Next file
    
End Sub


