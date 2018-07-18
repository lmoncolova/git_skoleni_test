Attribute VB_Name = "mod_Auto_Tvorba_Slozek_EXE_INS"
Option Explicit

Private Const cstExeTitulPripona1 As String = "s doložkou"
Private Const cstExeTitulPripona2 As String = "Výpis ze seznamu pohledávek"
Private Const cstExeTitulPMOpravne As String = "PM opravné"
Private Const cstExeTitulPMOpravne_negativ As String = "NePM"
Private Const cstStrCenik As String = "Ceník"
Private Const cstTerminationLetterName As String = "Zesplatòující dopis"
Private Const cstInsolPrihlaskaName As String = "Insolvenèní prihláška "
Private Const cst_max_pocet_CUIDu As Long = 120



Public Sub GenerateClientFoldersEXE()
'RISK-1114

'Volá se z Update_stav_all

'Pøi zadání urèitého stavu na klientovi v Jezevèíkovi - konkrétnì viz public funkce ExeFolderToBeCreated

'       na nalézací kartì jde o stavy èíslo 16, 17, 59, 68, 70, 73, 84, 87
'       na trestní kartì stavy èíslo 15, 16
'       na Insolvenèní kartì èíslo 22,
'       na kartì Dìdicové èíslo 66, 71, 74, 75, 76, 78, 80

'dojde k automatickému vytvoøení složky v LATE/EXEKUCE/PODANÉ, která bude mít v názvu PøíjmenímezeraJménomezeraCUID klienta.
'Do této novì vytvoøené složky se souèasnì automaticky nakopíruje exekuèní titul (rozhodnutí soudu s doložkou) ze složky klienta v LATE/EPR/Podané.
'V názvu souboru s exekuèním titulem je vždy text "s doložkou" anebo "Výpis ze seznamu pohledávek".


Dim strDestPath As String, StrSourcePath As String, StrLookUpFileSubstr As String
Dim fso As FileSystemObject
Dim strFileToBeCopied As String
Dim blnFileFound As Boolean

Set fso = New FileSystemObject

'Nastavím cestu, kde se má složka vytvoøit
If schema = "APP_COLL_JEZEVCIK" Then
    strDestPath = cfg.Cesta("ExePodane") & c.Nazev_slozky
Else
    strDestPath = cfg.Cesta("TestovaciSlozka") & c.Nazev_slozky
End If


' Zavolám funkci, která složku vytvoøí
'Pokud už složka existuje, tak nechci dìlat vùbec nic
If Not fso.FolderExists(strDestPath) Then

    If CreatePath(strDestPath) Then 'Funkce vrátí True, pokud folder existuje (bez ohledu na to jestli byl prave vytvoren nebo jiz existoval)
    
        
        StrSourcePath = cfg.Cesta("EprPodaneZaloby") & c.Nazev_slozky
        
        
        '''   dokument 1 (cstExeTitulPripona1 = s dolozkou (ALE NE OPRAVNE))
        
        strFileToBeCopied = FindFileInFolderOrSubfolder(srcPath:=StrSourcePath, FileNameSubstr:=cstExeTitulPripona1, FileNameSubstr_negative:=cstExeTitulPMOpravne)
        If strFileToBeCopied <> "" Then
            If Not CopyFile(strFileToBeCopied, StrSourcePath, strDestPath) Then
                MsgBox "Jezevèík error jez-9631268: " & vbCrLf & " Nepovedlo se zkopírovat soubor " & strFileToBeCopied & " do složky", vbExclamation
            End If
        End If
        strFileToBeCopied = ""
        
        
        '''   dokument 2 (cstExeTitulPripona2 = vypis ze seznamu pohl)
        
        strFileToBeCopied = FindFileInFolderOrSubfolder(srcPath:=StrSourcePath, FileNameSubstr:=cstExeTitulPripona2)
        If strFileToBeCopied <> "" Then
            If Not CopyFile(strFileToBeCopied, StrSourcePath, strDestPath) Then
                MsgBox "Jezevèík error jez-9631268: " & vbCrLf & " Nepovedlo se zkopírovat soubor " & strFileToBeCopied & " do složky", vbExclamation
            End If
        End If
        strFileToBeCopied = ""
        
        
        '''   dokument 3 (cstExeTitulPMOpravne)
        
        strFileToBeCopied = FindFileInFolderOrSubfolder(srcPath:=StrSourcePath, FileNameSubstr:=cstExeTitulPMOpravne, FileNameSubstr_negative:=cstExeTitulPMOpravne_negativ)
        If strFileToBeCopied <> "" Then
            If Not CopyFile(strFileToBeCopied, StrSourcePath, strDestPath) Then
                MsgBox "Jezevèík error jez-9631268: " & vbCrLf & " Nepovedlo se zkopírovat soubor " & strFileToBeCopied & " do složky", vbExclamation
            End If
        End If
        
        
        
    End If
    
End If

End Sub

Public Function ExeFolderToBeCreated(strRizeni As String, StrNameStatus As String) As Boolean
    'RISK-1114
    'Funkce øíká, jestli se mají pro klienta vytvoøit exekuèní složky
    'Na základì nadefinovaných stavù na rùzných øízeních
    
            '       na nalézací kartì jde o stavy èíslo 16, 17, 59, 68, 70, 73, 84, 87
            '       na trestní kartì stavy èíslo 15, 16
            '       na Insolvenèní kartì èíslo 22,
            '       na kartì Dìdicové èíslo 66, 71, 74, 75, 76, 78, 80
                        
    'Volá se z formu Update_stav_all a pokud je vyhodnocena jako true, tak volá proceduru GenerateClientFoldersEXE
    If strRizeni = "epr" Then
    
        Select Case StrNameStatus
            Case "PM Platební rozkaz s doložkou", "PM Rozsudek s doložkou", "PM rozsudek – SK s doložkou", "PM Øízení zastaveno - s doložkou", _
                 "PM rozsudek  - žaloba zamítnuta s doložkou", "PM Usnesení o schválení smíru s doložkou", "PM Rozsudek - èásteènì zamítnuto s doložkou", _
                 "PM Rozsudek - SK èásteènì zamítnuto s doložkou"
                 'SKPs 16, 17, 59, 68, 70, 73, 84, 87
                ExeFolderToBeCreated = True
        End Select
    
    ElseIf strRizeni = "trest" Then
    
        Select Case StrNameStatus
            Case "PM Trestní pøíkaz s doložkou", "PM Rozsudek s doložkou"
                'SKPs 15, 16
                ExeFolderToBeCreated = True
        End Select
    
    ElseIf strRizeni = "ins" Then
    
        Select Case StrNameStatus
            Case "Výpis ze seznamu pohledávek"
                'SKP 22
                ExeFolderToBeCreated = True
        End Select
    
    ElseIf strRizeni = "heir" Then
        
        Select Case StrNameStatus
            Case "PM Platební rozkaz s doložkou", "PM Rozsudek - èásteènì zamítnuto s doložkou", "PM Rozsudek - SK èásteènì zamítnuto s doložkou", _
                 "PM rozsudek – SK s doložkou", "PM Rozsudek s doložkou", "PM Øízení zastaveno - s doložkou", "PM Usnesení o schválení smíru s doložkou"
                'SKPs 66, 71, 74, 75, 76, 78, 80
                ExeFolderToBeCreated = True
        End Select
    
    Else
    
        ExeFolderToBeCreated = False
    
    End If

End Function


Public Sub Generate_folders_for_listed_clients(ByVal rizeni As typRizeni)
    ' RISK-1114
    ' Vytvoril Filip Antol 25.5.2018 do verze jez.v.5.7
    ' Tahle funkce vytvori exekucni nebo insolvenèní slozky (podle argumentu) pro kazdeho klienta ze seznamu, ktery je jeste nema
    ' Jedna se pouze o cyklus, ktery postupne klienty inicializuje a vola GenerateClientFoldersEXE
    ' vstupem je RANGE obsahujici CUIDy klientu, ke kterym se maji slozky vytvorit (funkce se zepta)
    ' preventivne jsem omezil maximalni pocet bunek na 120 (cst_max_pocet_CUIDu) - pokud nekdo nahodou zada milion bunek, aby to cele nespadlo. Nebo aby se nevytvorilo absurdni mnozstvi slozek..
    
    
    Dim rng_seznam_CUIDu As Range
    Dim current_cell As Range
    Dim current_CUID As Long
    Dim current_ID_person As Long
    
    
    ''' zadani vstupnich hodnot
    
    On Error GoTo konec_generovani
    Set rng_seznam_CUIDu = Application.InputBox("Vyber buòky s CUIDmi klientù", "Vyber CUID", Type:=8)
    On Error GoTo 0
    
    
    ''' overeni vstupnich hodnot
    
    If rng_seznam_CUIDu.count > cst_max_pocet_CUIDu Then
        MsgBox "Musíš zadat nejvýše " & cst_max_pocet_CUIDu & " CUIDù." & vbCrLf & _
                "(pak èekej až vyskoèí hláška ""Hotovo"", mùže to trvat nìjakou minutu.)", vbCritical
        Exit Sub
    End If
    
    
    ''' inicializace klientu a generovani slozek
    
    strMsgbox = ""
    
    For Each current_cell In rng_seznam_CUIDu
        
        ' MsgBox "jedu na klienta " & current_cell
        current_CUID = current_cell.value
        
        If Len(current_CUID) > 1 Then
        
            current_ID_person = ConvertCUIDtoId(current_CUID)
            Set c = New Client
            c.initialize (current_ID_person)
            'Debug.Print c.name
            
            
            If rizeni = typRizeni_EXE Then
                Call GenerateClientFoldersEXE
            ElseIf rizeni = typRizeni_INS Then
                GenerateClientFoldersINS
                ' RISK-1159
            End If
            
        End If
        ' MsgBox "dokoncil jsem klienta " & current_cell
        
    Next current_cell

    HlaskaUzivateli "Vytváøení složek ukonèeno."
    VypisVsechnyHlasky
    
    'MsgBox "jsem na konci generovani"
konec_generovani:
    
End Sub





Public Sub GenerateClientFoldersINS()
    ' ' RISK-1159
    ' msgboxy se vyhazuji pres HlaskaUzivateli. T.j. pak se musi zavolat VypisVsechnyHlasky
    
    Dim strDestFolder As String
    Dim DestFolder As Scripting.Folder
    Dim strSourcePrilohy As String
    Dim SourceFolderPrilohy As Scripting.Folder ' XX vyhod "Folder" z nazvu pokud nebudou dalsi
    Dim strSourcePMrozhodnuti As String
    Dim strSourceZespDopis As String
    Dim strClientSubfolder As String
    Dim strFileToCopy As String
    Dim FileToCopy As Scripting.file
    Dim fso As FileSystemObject
    Dim Ins As Insolvency
    Dim strNechciHlaskyZInsolvence As String
    Dim strInsolPrihlaskaName As String
    
    
    'Dim SourceFolderPMrozhodnuti As scripting.Folder
    'Dim SourceFolderZespDopis As scripting.Folder
    'Dim fileToBeCopied As String
    'Dim blnFileFound As Boolean
    
    
    '''  nastaveni pseudo-konstant
    Set fso = New FileSystemObject
    strClientSubfolder = c.Nazev_slozky
    strSourcePMrozhodnuti = cfg.Cesta("EprPodaneZaloby") & strClientSubfolder
    strSourceZespDopis = cfg.Cesta("ZespDopis")
    strInsolPrihlaskaName = cstInsolPrihlaskaName & c.Name_last
    
    
    
'    ''''  pseudokod
'    ' nadefinuj zdrojovou slozku
'        ' pro kazdy soubor ve zdrojove slozce (eventuelne po jednom, ale necham takhle)
'            ' zjisti jestli existuje
'            ' pokud neexistuje
'                ' nakopiruj
'    ''''  pseudokod
    
    
    '''   nastaveni cilove slozky
    If schema = "APP_COLL_JEZEVCIK" Then
        strDestFolder = cfg.Cesta("InsPripravene") & strClientSubfolder & "Pøílohy\"
    Else
        strDestFolder = cfg.Cesta("TestovaciSlozka") & strClientSubfolder & "Pøílohy\"
    End If
    
    
    
    ''''   Kopiruji soubory jenom kdyz slozka neexistuje
    If Not fso.FolderExists(strDestFolder) Then
        
        If CreatePath(strDestFolder) Then 'Funkce vrátí True, pokud folder existuje (bez ohledu na to jestli byl prave vytvoren nebo jiz existoval)
            
            Set DestFolder = fso.GetFolder(strDestFolder)
            
            ''' 1) Vzorove prilohy
            
            If fso.FolderExists(cfg.Cesta("InsPrilohy")) Then
                Set SourceFolderPrilohy = fso.GetFolder(cfg.Cesta("InsPrilohy"))
                
                ' kazdy soubor ve zdrojove slozce zkopiruj do cilove slozky (slo by zkopirovat i celou slozku)
                For Each FileToCopy In SourceFolderPrilohy.Files
                    
''                    If ((InStr(FileToCopy, cstStrCenik) > 0) And (c.Debetar = False)) Then ' chceme aby se Cenik nahraval, pouze pokud ma klient debet. Nevim ale jak zjistit to druhe... Debetari nam uplne nefunguji.
''                        ' cenik pouze pokud ma debet
                    If Not fso.FileExists(fso.BuildPath(DestFolder.Path, FileToCopy.name)) Then
                        FileToCopy.Copy fso.BuildPath(DestFolder.Path, FileToCopy.name), overwritefiles:=False
                    End If
                    
                Next FileToCopy
            End If
            
            
            ''' 2) PM rozhodnuti
            
            If fso.FolderExists(strSourcePMrozhodnuti) Then
                
                'Set SourceFolderPMrozhodnuti = fso.GetFolder(strSourcePMrozhodnuti)
                strFileToCopy = FindFileInFolder(strSourcePMrozhodnuti, cstExeTitulPripona1)
                If strFileToCopy <> "" Then
                    If Not CopyFile(strFileToCopy, strSourcePMrozhodnuti, strDestFolder) Then
                        HlaskaUzivateli "Jezevèík error jez-9631268: " & vbCrLf & " Nepovedlo se zkopírovat soubor " & strFileToCopy & " do složky"
                    End If
                End If
                
            End If
            
            
            ''' 3) Vzor insolvencni prihlasky
            
'            if fso.FolderExists()
'            strFileToCopy = Cesta("vzorInsPrihlaska")
'            if not copyfile(Cesta("vzorInsPrihlaska"),
'
            If fso.FileExists(cfg.Cesta("vzorInsPrihlaska")) Then
                Set FileToCopy = fso.GetFile(cfg.Cesta("vzorInsPrihlaska"))
                If Not fso.FileExists(fso.BuildPath(DestFolder.ParentFolder, FileToCopy.name)) Then
                    FileToCopy.Copy fso.BuildPath(DestFolder.ParentFolder, strInsolPrihlaskaName & "." & fso.GetExtensionName(FileToCopy.Path)), overwritefiles:=False ' FileToCopy.name
                End If
                        
            Else
                HlaskaUzivateli "error jez-36988: Nenalezen vzor insolvenèní pøihlášky: " & cfg.Cesta("vzorInsPrihlaska")
            End If
            
            
            
            ''' 4) Zesplatnujici dopis
            
            '''   podminka ze datumy se musi lisit
            strNechciHlaskyZInsolvence = strMsgbox
            Set Ins = New Insolvency
            Call Ins.initialize(Val(c.id_person))
            strMsgbox = strNechciHlaskyZInsolvence
            ' test cuid 392464
            
            If Ins.strDatumUpadku <> CStr(Ins.dtmDatumZesplatneni) Then
                If fso.FolderExists(strSourceZespDopis) Then
                    ' Set SourceFolderZespDopis = fso.GetFolder(strSourceZespDopis)
                    
                    strFileToCopy = FindFileInFolderOrSubfolder(strSourceZespDopis, c.CUID, ".pdf", blnPrintSearchedFolders:=False)
''''                    ' zakomentovavam, asi zbytecne a trva to dloouho
''''                    If strFileToCopy = "" Then
''''                        strFileToCopy = FindFileInFolderOrSubfolder(strSourceZespDopis, c.CUID, ".doc", blnPrintSearchedFolders:=False)
''''                    End If

                    ' debug.print fso.GetFileName(strFileToCopy), fso.GetParentFolderName(strFileToCopy)
''                    If strFileToCopy <> "" Then
''                        If Not CopyFile(srcFile:=fso.GetFileName(strFileToCopy), _
''                                        srcPath:=fso.BuildPath(strSourceZespDopis, fso.GetParentFolderName(strFileToCopy)) & "\", _
''                                        destPath:=strDestFolder) Then
''                            HlaskaUzivateli "Jezevèík error jez-9631268: " & vbCrLf & " Nepovedlo se zkopírovat soubor " & strFileToCopy & " do složky"
''                        End If
''                    End If
                    
                    
                    If fso.FileExists(strSourceZespDopis & strFileToCopy) Then
                        Set FileToCopy = fso.GetFile(strSourceZespDopis & strFileToCopy)
                        If Not fso.FileExists(fso.BuildPath(DestFolder.ParentFolder, cstTerminationLetterName & "." & fso.GetExtensionName(FileToCopy.Path))) Then
                            FileToCopy.Copy fso.BuildPath(DestFolder, cstTerminationLetterName & "." & fso.GetExtensionName(FileToCopy.Path)), overwritefiles:=False ' FileToCopy.name
                        End If
                    Else
                        HlaskaUzivateli "Zesplatòující dopis ke klientovi CUID " & c.CUID & " nebyl nalezen."
                    End If
                    
                    
                End If
            End If
            
            
        End If

    End If
    
    
End Sub



