Attribute VB_Name = "Public_Functions"
'Tento modul slouûÌ jakoûto koteËek pro vöechny public funkce
Option Explicit

Public strAdmins As String ' pouzivat pouze pro nedulezite veci, tohle neni bezpecne overeni uzivatele! plni se v Sub NaplnAdminov()

Public Declare Function sndPlaySound32 _
    Lib "winmm.dll" _
    Alias "sndPlaySoundA" ( _
        ByVal lpszSoundName As String, _
        ByVal uFlags As Long) As Long 'potrebne pro Sub PlayTheSound()

Public PrankAlertTime As Date

Sub RibbonSchovej()
    ''' uzivatelovi skryje hlavny excel ribbon (a dalsi detaily jako dolni lista)
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""RIBBON"", False)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    ActiveWindow.DisplayWorkbookTabs = False
        
End Sub

Sub RibbonUkaz()
    ''' uzivatelovi naspat ukaze excel ribbon (a dalsi detaily jako dolni lista)
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""RIBBON"", True)"
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    
End Sub


Public Sub NajdiAdresuMapyCZ(ByVal adresa As String)
' Sub upravi adresu na URL pro MAPY.CZ ktere chceme vhodit do prohlizece

Dim URL As String
    
    adresa = Replace(adresa, " ", "%20")
    URL = "https://mapy.cz/zakladni?source=addr&q= " & adresa
    Call OtvorWebStranku(URL)
    
End Sub


Public Sub Mail_small_Text_Outlook()
''' Jezevcik prebral z http://www.rondebruin.nl/win/s1/outlook/bmail4.htm

'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Working in Office 2000-2016
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Const URL_doodle As String = "https://doodle.com/poll/5gpib8uhs9b2t84t"

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    strbody = "Ahoj PavlÌku" & vbNewLine & vbNewLine & _
              "Zahlasuj prosÌm v doodli k ˙nikovke: " & URL_doodle & "." & vbNewLine & _
              "" & vbNewLine & _
              "DÌky," & vbNewLine & _
              "F."

    On Error Resume Next
    With OutMail
        .To = "Pavel.Hruska@airbank.cz"
        '.To = "filip.antol@airbank.cz"
        .CC = ""
        .BCC = ""
        .Subject = "⁄nikovka"
        .Body = strbody
        'You can add a file like this
        '.Attachments.Add ("C:\test.txt")
        .send   'or use .Display
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub





Public Sub OtvorWebStranku(ByVal WebUrl As String)
'' sub, ktora otvori internet explorer na danej stranke
'' zdroj: https://social.msdn.microsoft.com/Forums/en-US/42fbea0d-3977-480f-a4c3-6093e34f469d/open-a-web-page-with-vba?forum=isvvba
Dim IEapp As Object
'Dim WebUrl As String
    
    Set IEapp = CreateObject("InternetExplorer.Application") 'Set IEapp = InternetExplorer
    'WebUrl = "http://www.google.com"
    
    'You can't really automate the passing of variables (ie. Login Name and Passwords) using Chrome or FireFox, but the below code will open the desired page
    'Shell ("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe -url " & WebUrl)  'Opens URL in Chrome if installed, replace chrome.exe with FireFox
                                                                                            
    With IEapp
        .Silent = True 'No Pop-ups
        .Visible = True 'Set InternetExplorer to Visible
        .Navigate WebUrl 'Load web page

        'Run and Wait, if you intend on passing variables at a later stage
        Do While .Busy
            DoEvents
        Loop
        
        Do While .readyState <> 4
            DoEvents
        Loop
    
    End With
           
End Sub



Public Sub Listview_Proper_Sort(Sorted_Listview As ListView, ColumnHeader As MSComctlLib.ColumnHeader, Optional ByVal blnAscendingOrder = False, Optional ByVal blnDescendingOrder = False)
    'MM: Public subroutine pro ¯azenÌ vöech listview, funguje i pro datumy a ËÌsla jako texù·ky
    'http://www.vbaexpress.com/forum/showthread.php?54357-Listview-sort-when-column-header-clicked
    'MM: Nefunguje p¯i ¯azenÌ z·porn˝ch a kladn˝ch ËÌsel
    
    
    If Sorted_Listview.ListItems.count = 0 Then Exit Sub

     
    Dim strItem As String 'List item
    Dim intMaxItemLength As Integer 'Max. length of list items
    Dim Intcounter As Integer 'Counter
    Dim i As Integer 'Counter for first non-null value
     
    With Sorted_Listview
         'Set sorted property of listview
        .Sorted = False
         'Add temporary column to sort
        .ColumnHeaders.Add , , , 0
         
         'Select first item from list view
         'If first column
        i = 1
        If ColumnHeader.index = 1 Then
               'MM ˙prava: byla tam chyba pokud jsem ¯adil a prvnÌ hodnota byla null, tak se nerozpoznal typ. UdÏl·m tedy loop aû najdu prvnÌ nenulovou a podle toho urËÌm typ promÏnnÈ
                Do Until Sorted_Listview.ListItems(i) <> vbNullString
                    i = i + 1
                    If i > Sorted_Listview.ListItems.count Then
                        ' v pripade, ze vsechny radky jsou NULL, tak v nasledujicim loopu by spadl cely excel
                        Exit Sub
                    End If
                Loop
            strItem = .ListItems(i)
             'If another column
        Else
                'MM ˙prava: byla tam chyba pokud jsem ¯adil a prvnÌ hodnota byla null, tak se nerozpoznal typ. UdÏl·m tedy loop aû najdu prvnÌ nenulovou a podle toho urËÌm typ promÏnnÈ
                Do Until Sorted_Listview.ListItems(i).SubItems(ColumnHeader.index - 1) <> vbNullString
                    i = i + 1
                    If i > Sorted_Listview.ListItems.count Then
                        ' v pripade, ze vsechny radky jsou NULL, tak v nasledujicim loopu by spadl cely excel
                        Exit Sub
                    End If
                Loop
            strItem = .ListItems(i).SubItems(ColumnHeader.index - 1)
        End If
    End With
     
     'Date type?
    If IsDate(strItem) Then
        With Sorted_Listview
             'If first column
            If ColumnHeader.index = 1 Then
                For Intcounter = 1 To .ListItems.count
                     'Get list item
                    strItem = Trim(.ListItems(Intcounter))
                     'Add list item to last column in new format to sort
                    .ListItems(Intcounter).SubItems(.ColumnHeaders.count - 1) = format(strItem, "yyyymmddHHMMSS")
                Next
                 'If another column
            Else
                For Intcounter = 1 To .ListItems.count
                     'Get sub item
                    strItem = Trim(.ListItems(Intcounter).SubItems(ColumnHeader.index - 1))
                     'Add sub item to last column in new format to sort
                    .ListItems(Intcounter).SubItems(.ColumnHeaders.count - 1) = format(strItem, "yyyymmddHHMMSS")
                Next
            End If
             'Set sort key property (temporary column)
            .SortKey = .ColumnHeaders.count - 1
        End With
         'Numberic type?
    ElseIf IsNumeric(strItem) Then
        With Sorted_Listview
             'Set start up value for maximum length of items
            intMaxItemLength = 0
             'If first column
            If ColumnHeader.index = 1 Then
                 'Find max. length of items
                For Intcounter = 1 To .ListItems.count
                    strItem = Trim(Replace(Replace(.ListItems(Intcounter), ",", ""), ".", ""))
                    If Len(strItem) > intMaxItemLength Then
                        intMaxItemLength = Len(strItem)
                    End If
                Next
                For Intcounter = 1 To .ListItems.count
                     'Get sub item
                    strItem = Trim(Replace(Replace(.ListItems(Intcounter), ",", ""), ".", ""))
                     'Add sub item to last column in new format to sort
                    .ListItems(Intcounter).SubItems(.ColumnHeaders.count - 1) = "0" & String(intMaxItemLength - Len(strItem), "0") & strItem
                Next
                 'If another column
            Else
                 'Find max. length of items
                For Intcounter = 1 To .ListItems.count
                    strItem = Trim(Replace(Replace(.ListItems(Intcounter).SubItems(ColumnHeader.index - 1), ",", ""), ".", ""))
                    If Len(strItem) > intMaxItemLength Then
                        intMaxItemLength = Len(strItem)
                    End If
                Next
                For Intcounter = 1 To .ListItems.count
                     'Get sub item
                    strItem = Trim(Replace(Replace(.ListItems(Intcounter).SubItems(ColumnHeader.index - 1), ",", ""), ".", ""))
                     'Add sub item to last column in new format to sort
                    .ListItems(Intcounter).SubItems(.ColumnHeaders.count - 1) = "0" & String(intMaxItemLength - Len(strItem), "0") & strItem
                Next
            End If
             'Set sort key property (temporary column)
            .SortKey = .ColumnHeaders.count - 1
        End With
         'Another type?
    Else
        Sorted_Listview.SortKey = ColumnHeader.index - 1
    End If
     
     'Set properties of list view
    With Sorted_Listview
         'Change the sort order
        .SortOrder = 1 - .SortOrder
        If blnAscendingOrder Then .SortOrder = lvwAscending   ' = 0
        If blnDescendingOrder Then .SortOrder = lvwDescending ' = 1
         'Set sorted property
        .Sorted = True
         'Remove temporary column
        .ColumnHeaders.remove .ColumnHeaders.count
    End With
     
End Sub



Public Sub HlaskaUzivateli(text As String) ', byref strMsgbox As String)
'pracuje s globalni promennou strMsgbox a se Sub VypisVsechnyHlasky
'opatrne se zmenami, funkce se pouziva i v jinych modulech
If Left(Trim(text), 1) = ">" Then
    strMsgbox = strMsgbox & Trim(text) & Chr(10)
Else
    strMsgbox = strMsgbox & "> " & Trim(text) & Chr(10)
End If
    
End Sub


Public Sub VypisVsechnyHlasky()
'opatrne se zmenami, funkce se pouziva i v jinych modulech
'pracuje s globalni promennou strMsgbox a se Sub HlaskaUzivateli
If strMsgbox <> "" Then
    
    'MsgBox je obmedzeny na cca 1024 znakov, preto ak je dlhsi tak rozdelime na viacero
    MsgBox Left(strMsgbox, 999), vbExclamation, "Varov·nÌ (pro zkopÌrov·nÌ zpr·vy pouûij ctrl+c)"
    
    If Len(strMsgbox) > 999 Then
        MsgBox Mid(strMsgbox, 1000, 999), vbExclamation, "Varov·nÌ (pro zkopÌrov·nÌ zpr·vy pouûij ctrl+c)"
    End If
    
    If Len(strMsgbox) > 1999 Then
        MsgBox Mid(strMsgbox, 2000, 999), vbExclamation, "Varov·nÌ (pro zkopÌrov·nÌ zpr·vy pouûij ctrl+c)"
    End If
    
    If Len(strMsgbox) > 2999 Then '
        MsgBox Mid(strMsgbox, 3000, 999), vbExclamation, "Varov·nÌ (pro zkopÌrov·nÌ zpr·vy pouûij ctrl+c)"
    End If
    
    
    strMsgbox = ""
End If

End Sub


Public Sub naplnElement(cilElement As String, nameElement As String, value As String)
    Set atribut = XDoc.SelectNodes(cilElement & nameElement)
    atribut(0).text = value
End Sub

Public Sub naplnElementNumber(cilElement As String, nameElement As String, value As String)
    Set atribut = XDoc.SelectNodes(cilElement & nameElement)
    atribut(0).text = Replace(value, ",", ".")
End Sub

Public Sub pridejElement(i As Integer, cilElement As String, nameElement As String)
        Set ParentElement = XDoc.SelectNodes(cilElement)
        Set ChildElement = XDoc.createElement(nameElement)
        ParentElement(i).appendChild ChildElement
End Sub

Public Sub UlozSoubor(Folder As String, fileName As String, bezSlozky As String)
Dim fso As Object
'Save the XML file 'Pro testovacÌho jezevËÌka se ukl·da na W: do sloûky Testy
If schema = "APP_COLL_JEZEVCIK_TEST" Then
    XDoc.Save cfg.Cesta("TestovaciSlozka") & fileName
    'Call Shell("explorer.exe " & cfg.Cesta("TestovaciSlozka"), vbNormalFocus) 'otvirani slozky na testu
Else ' na PROD se ukl·d· do nachystanÈ sloûky a pokud neexistuje, do sloûky Klienti bez sloûky
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    
    
    If fso.FolderExists(Folder) Then
        XDoc.Save (Folder & "\" & fileName)
        Call Shell("explorer.exe " & Folder, vbNormalFocus)
    Else: MsgBox "Sloûka pro klienta " & c.name & " neexistuje." 'pokud sloûka klienta neexistuje
        XDoc.Save (bezSlozky & fileName)
        Call Shell("explorer.exe " & bezSlozky, vbNormalFocus)
        
    End If
End If
'RISK-391
End Sub



Public Sub jz()
'funkce pro rychle restartovani jezevcika z Immediate window - staci napsat "jz" do prvni radky
'Reset
    Call ThisWorkbook.Workbook_Open
    'Main.CommandButton1.SetFocus 'nefunguje
    'Application.run "Main.CommandButton1_Click()"
    Call Main.CommandButton1_Click
End Sub

Public Sub jj()
' lol :D
Call jz
End Sub
Public Sub jez()
' lol :D
Call jz
End Sub
Public Sub jy()
' lol :D
Call jz
End Sub
Public Sub zj()
' lol :D
Call jz
End Sub

Public Sub jji(Optional ByVal CUID As Long = 413982)
' jj + initialize
    ' CUID ins: 606969
    jj
    Set c = New Client
    c.initialize (ConvertCUIDtoId(CUID))
End Sub



Public Sub Discoo()
'na karte Main budu blikat podsvietenia kariet
'striedat sa to bude asi kazdu sekundu, nedoporucujem davat menej, excel to nezvlada
    If blnPrank Then
        Dim sngRandom(1 To 5) As Single
        Dim i As Integer
        
        For i = 1 To 5
            sngRandom(i) = Round(Rnd()) ' * 3 + 0.5, 0)
        Next i
        
        With Main
            .CommandButton2.ForeColor = barvaCerna
            .CommandButton6.ForeColor = barvaCerna
            .CommandButton7.ForeColor = barvaCerna
            .CommandButton8.ForeColor = barvaCerna
            .CommandButton9.ForeColor = barvaCerna
            .CommandButton10.ForeColor = barvaCerna
            .CommandButton11.ForeColor = barvaCerna
            .CommandButton12.ForeColor = barvaCerna
            .CommandButton13.ForeColor = barvaCerna
            .CommandButton_Sonik.ForeColor = barvaCerna
        End With
            
        If sngRandom(1) = 1 Then
            With Main
                .CommandButton2.ForeColor = barvaZelenaAirBank
                .CommandButton6.ForeColor = barvaZelenaAirBank
            End With
        End If
        
        If sngRandom(2) = 1 Then
            With Main
                .CommandButton7.ForeColor = barvaZelenaAirBank
                .CommandButton8.ForeColor = barvaZelenaAirBank
            End With
        End If
        
        If sngRandom(3) = 1 Then
            With Main
                .CommandButton9.ForeColor = barvaZelenaAirBank
                .CommandButton10.ForeColor = barvaZelenaAirBank
            End With
        End If
        
        If sngRandom(4) = 1 Then
            With Main
                .CommandButton11.ForeColor = barvaZelenaAirBank
                .CommandButton12.ForeColor = barvaZelenaAirBank
            End With
        End If
        
        If sngRandom(5) = 1 Then
            With Main
                .CommandButton13.ForeColor = barvaZelenaAirBank
                .CommandButton_Sonik.ForeColor = barvaZelenaAirBank
            End With
        End If
        
        PrankAlertTime = Now + TimeValue("00:00:02")
        Application.OnTime PrankAlertTime, "Discoo"
    End If
End Sub


Public Function NahodnyUzivatel()
'blnPrank
Dim jmena(1 To 13) As Variant
Dim i As Single
'Application.UserName
'Main.username
'jmena = Array("word1", "word2", "word3")
jmena(1) = "MacÌËek Martin"
jmena(2) = "Rajsiglov· Eva"
jmena(3) = "Doubkov· MarkÈta"
jmena(4) = "Halcin Marek"
jmena(5) = "Kaöparov· Lucie"
jmena(6) = "Starobov· Petra"
jmena(7) = "ätefl Michael"
jmena(8) = "Tomeökov· Lucie"
jmena(9) = "ZapletalÌkov· Lenka"
jmena(10) = "Seidlov· Martina"
jmena(11) = "Severa Vlastimil"
jmena(12) = "Broûina VÌt"
jmena(13) = "Broûina VÌt" 'VÌùa je tam dvakr·t, nech mu to pad· Ëastejöie :D


i = Round(Rnd() * (UBound(jmena) - 1) + 1, 0)

NahodnyUzivatel = jmena(i)


End Function

'hrani zvuku
Sub PlayTheSound(ByVal WhatSound As String)
'zdroj: https://stackoverflow.com/questions/9443801/how-to-play-a-wave-file-or-sound-file-using-vba
    If Dir(WhatSound, vbNormal) = "" Then
        ' WhatSound is not a file. Get the file named by
        ' WhatSound from the Windows\Media directory.
        'WhatSound = Environ("SystemRoot") & "\Media\" & WhatSound 'C:\Windows\Media
        WhatSound = "W:\Vymahani\LATE\JezevËÌk\xx\" & WhatSound
        If InStr(1, WhatSound, ".") = 0 Then
            ' if WhatSound does not have a .wav extension,
            ' add one.
            WhatSound = WhatSound & ".wav"
        End If
        If Dir(WhatSound, vbNormal) = vbNullString Then
            Beep            ' Can't find the file. Do a simple Beep.
            Exit Sub
        End If
    Else
        ' WhatSound is a file. Use it.
    End If

    sndPlaySound32 WhatSound, 0&    ' Finally, play the sound.
End Sub



Public Sub CopyToClipboard(ByVal value As Variant, Optional karta As String)
'pri kliknuti (napr. na label CUID na karte klienta) se zkopiruje do schranky
'doporuceno nastavit karta = Me.name

Call text_information_hide(karta)

If value <> vbNullString Then
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject
    clipboard.SetText value
    clipboard.PutInClipboard
End If


Select Case karta
    
    Case "Klient"
        Klient.Label27.Visible = True
        Klient.Label29.Visible = True
    
    Case "Klient_dedictvi"
        Klient_dedictvi.Label27.Visible = True
        Klient_dedictvi.Label29.Visible = True
    
    Case "Klient_exe"
        Klient_exe.Label27.Visible = True
        Klient_exe.Label29.Visible = True
    
    Case "Klient_insol"
        Klient_insol.Label27.Visible = True
        Klient_insol.Label29.Visible = True
    
    Case "Klient_trest"
        Klient_trest.Label27.Visible = True
        Klient_trest.Label29.Visible = True
    
End Select


End Sub

Public Sub ZabarviRadkyTextove(ByRef LV As ListView, ByVal sloupec As Integer, ByVal text_pro_zabarveni As String, Optional ByVal barvicka As barva = barvaRudaSvetla)
'SUB zacervena radky listview, ktere maji hodnotu v SLOUPCI = text_pro_zabarveni
'LV je listview napr: Souhrn.LV_KontrolaPM
'sloupec je poradove cislo sloupecku podle ktereho zabarvit 'pro prvni sloupec staci dat 0
Dim i As Integer, j As Integer
Dim itmX As ListItem, lvSI As ListSubItem



If sloupec = 0 Then
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i) = text_pro_zabarveni Then 'XXXX pouze a jenom pro presnou zhodu
            Set itmX = LV.ListItems(i)
            itmX.ForeColor = barvicka
            For j = 1 To LV.ColumnHeaders.count - 1
                Set lvSI = itmX.ListSubItems(j)
                lvSI.ForeColor = barvicka
            Next j
        End If
    Next i
ElseIf sloupec > 0 Then
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).SubItems(sloupec) = text_pro_zabarveni Then 'XXXX pouze a jenom pro presnou zhodu
            Set itmX = LV.ListItems(i)
            itmX.ForeColor = barvicka
            For j = 1 To LV.ColumnHeaders.count - 1
                Set lvSI = itmX.ListSubItems(j)
                lvSI.ForeColor = barvicka
            Next j
        End If
    Next i
Else
    Debug.Assert 1 = 0
    MsgBox "Jez-error J54474 v ZabarviRadkyTextove", vbCritical, "a jÈjda"
End If

End Sub


Public Sub ZabarviRadky(ByRef LV As ListView, ByVal sloupec As Integer, Optional ByVal barvicka As barva = barvaRudaSvetla)
'SUB zacervena radky listview, ktere maji hodnotu v SLOUPCI <= 0
'LV je listview napr: Souhrn.LV_KontrolaPM
'sloupec je poradove cislo sloupecku podle ktereho zabarvit
Dim i As Integer, j As Integer
Dim itmX As ListItem, lvSI As ListSubItem


If sloupec = 0 Then
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i) <= 0 Then
            Set itmX = LV.ListItems(i)
            itmX.ForeColor = barvicka
            For j = 1 To LV.ColumnHeaders.count - 1
                Set lvSI = itmX.ListSubItems(j)
                lvSI.ForeColor = barvicka
            Next j
        End If
    Next i
ElseIf sloupec > 0 Then
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).SubItems(sloupec) <= 0 Then
            Set itmX = LV.ListItems(i)
            itmX.ForeColor = barvicka
            For j = 1 To LV.ColumnHeaders.count - 1
                Set lvSI = itmX.ListSubItems(j)
                lvSI.ForeColor = barvicka
            Next j
        End If
    Next i
Else
    Debug.Assert 1 = 0
    MsgBox "Jez-error J54474 v ZabarviRadky", vbCritical, "a jÈjda"
End If

End Sub


Public Sub NaplnAssigneeSeznam(ByRef YourComboBox As combobox)
'tahle Sub naplni seznam hodnot do rolovatka Assignee pri vkladani nebo editovani upozorneni
'ma privlastek PUBLIC, protoze se pouziva na karte Upozorneni i na karte Souhrn
'Sub zaroven nastavi default value na prihlaseneho uzivatele
Dim sql As String
Dim rst As Object 'adodb.Recordset
Dim i As Integer
Dim username As String


If (InStr(schema, "TEST") > 0) Then
    sql = "SELECT a.username FROM " & schema & ".jez_users a WHERE a.flag_valid = 'Y' "
Else
    sql = "SELECT a.username FROM " & schema & ".jez_users a WHERE a.flag_valid = 'Y' AND a.username NOT IN ('AIR_ANTOL', 'AIR_MACICEKM', 'AIR_SEVERA', 'AIR_GRUZA', 'AIR_TOMESKOVA') ORDER BY a.username ASC"
End If

YourComboBox.Clear
YourComboBox.AddItem ""

Set rst = New ADODB.Recordset
rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst.EOF Then
        rst.MoveFirst
        For i = 1 To rst.RecordCount
            username = Right(rst![username], Len(rst![username]) - 4) 'bez AIR_
            
            YourComboBox.AddItem username
            
        rst.MoveNext
        Next i
    End If
rst.Close

End Sub



Public Sub NastavKliceListview(ByRef listviewToSet As ListView)
'pre vsetky stlpce v danom listview nastavi ich column keys
'ide o to, ze mame columnheader.TEXT vsade nastaveny, ale aby sme mohli zavolat stlpec cez jeho meno, musi mat nastaveny columnheader.KEY
'tato SUB nastavi vsetky KEYS aby boli zhodne s columnheaders
'staci spustit raz pre kazdy listview, po tom ako sa nastavia vsetky columnheaders
'zatial nefunguje :D



'nasledujuce funguje az ked naplnime hodnoty do listview
'Debug.Print Souhrn.ListView3.ListItems(1).SubItems(Souhrn.ListView3.ColumnHeaders.Item("Klient").index - 1)

Dim colhead As ColumnHeader

For Each colhead In listviewToSet.ColumnHeaders
    'Debug.Print colhead.text
    If colhead.key = "" Then
        colhead.key = colhead.text 'key musi byt unique
    End If
Next

End Sub

Public Function ListvZiskejHodnotu(ByRef LV As ListView, ByVal colName As String, Optional ByRef rowNumber As Integer = -2) As String
'funkce vrati hodnotu v danem radku a sloupci vybraneho Listview
'pokud neni zadan rowNumber, funkce pracuje se selectedItem
'rowNumber = item number = cislo radku
'colname je nazov stlpca ktory chceme ziskat


'ak mam string "Klient", a mam nastaveny KEY, k cislu sa dostanem takto
'Debug.Print Souhrn.ListView3.ColumnHeaders.item("Klient")
'Debug.Print Souhrn.ListView3.ColumnHeaders.item("Klient").index

If rowNumber > 0 Then

    If LV.ColumnHeaders.item(colName).index = 1 Then
        ListvZiskejHodnotu = LV.ListItems(rowNumber).text
    Else
        ListvZiskejHodnotu = LV.ListItems(rowNumber).SubItems(LV.ColumnHeaders.item(colName).index - 1)
    End If
    
Else
    
    If LV.ColumnHeaders.item(colName).index = 1 Then
        ListvZiskejHodnotu = LV.SelectedItem.text
    Else
        ListvZiskejHodnotu = LV.SelectedItem.SubItems(LV.ColumnHeaders.item(colName).index - 1)
    End If
    
End If

End Function

Public Sub ListvVlozHodnotu(ByRef LV As ListView, ByVal colName As String, ByVal newValue As String, Optional ByRef rowNumber As Integer = -2)
'funkce nastavi hodnotu v danem radku a sloupci vybraneho Listview
'pokud neni zadan rowNumber, funkce pracuje se selectedItem
'rowNumber = item number = cislo radku
'colname je nazov stlpca do ktoreho chceme vlozit


If rowNumber > 0 Then

    If LV.ColumnHeaders.item(colName).index = 1 Then
        LV.ListItems(rowNumber).text = newValue
    Else
        LV.ListItems(rowNumber).SubItems(LV.ColumnHeaders.item(colName).index - 1) = newValue
    End If
    
Else
    
    If LV.ColumnHeaders.item(colName).index = 1 Then
        LV.SelectedItem.text = newValue
    Else
        LV.SelectedItem.SubItems(LV.ColumnHeaders.item(colName).index - 1) = newValue
    End If
    
End If

End Sub


Public Function KliceKListview(ByRef LV As ListView) As Dictionary
' funkce se zavola pred naplnenim libovolneho userformu
' vystup je DICTIONARY ktery ke kazdemu sloupecku vraci jeho poradove cislo
' vyhoda je ze pri plneni sloupecku uz neni potreba myslet na jejich poradi

''''' PRIKLAD:
' puvodni kod:
'               itmX.SubItems(1) = "" & rst1![cuid] & ""
' novy kod:
'               Dim keys As Dictionary
'               Set keys = KliceKListview(Souhrn.ListView3)
'               itmX.SubItems(keys("CUID")) = "" & rst1![cuid] & ""

Dim colhead As ColumnHeader
Set KliceKListview = New Dictionary

For Each colhead In LV.ColumnHeaders
    'Debug.Print colhead.text
    If colhead.key = "" Then
        MsgBox "jez-error J80012: ListViewKey not set"
    Else
        KliceKListview.Add colhead.key, (colhead.index - 1)
    End If
    
Next

End Function


''Public Sub NaplnAdminov()
''' na niekolkych miestach v jezevcikovi by sa hodilo mat poruke boolean aby to bolo videtelne iba pre adminov (napr. veci ktore su iba pre veduceho PV)
''' nechceme to ale spravit prilis naivne..
''
''Dim sql As String
''
''sql = "SELECT a.username FROM " & schema & ".jez_users a WHERE a.user_role = 'ADMIN'"
''Set rst1 = New ADODB.Recordset
''rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''    If Not rst1.EOF Then
''        rst1.MoveFirst1
''        For i = 1 To rst1.RecordCount
''            username = Right(rst1![username], Len(rst1![username]) - 4) 'bez AIR_
''
''            YourComboBox.AddItem username
''
''        rst1.MoveNext
''        Next i
''    End If
''rst1.Close
''
''
''End Sub



Public Sub Exportuj_Recordset_Do_Noveho_Workbooku(Optional ByRef rst As ADODB.Recordset, _
                                                  Optional ByRef sql As String, _
                                                  Optional strWorkbookName As String)
' Jakykoliv recordset se proste vypise do noveho workbooku ve forme tabulky
' dlhe nazvy jsou COOOOOL!!!
' vstupem je bud primo recordset, nebo jenom SQL dotaz
' pokud zadas nazev strWorkbookName, tak se excel ulozi uzivateli na plochu
' neni to nijak extra vychytana funkce, jenom jednoducha pomucka

Dim wbExport As Workbook 'Workbook
Dim rst1 As ADODB.Recordset
'Dim strWorkbookName As String ' puvodne vstupni argument, ale nelze ukladat

    If rst Is Nothing And sql = "" Then
        Debug.Assert 1 = 0
        Exit Sub
    End If

    Set wbExport = Workbooks.Add
    With wbExport
        .Title = "export " & strWorkbookName & " " & Now
        .Subject = "export " & strWorkbookName & " " & Now
        If Len(strWorkbookName) > 0 Then .SaveAs fileName:=CreateObject("WScript.Shell").specialfolders("Desktop") & "\export " & strWorkbookName & " " & Replace(Now, ":", "-") & ".xlsx"   'path to Desktop, regardless of user
    End With
    
    If Not rst Is Nothing Then
        Call VlozRecordsetDoWorkbooku(wbExport, rst)
    ElseIf Not sql = "" Then
        
        Set rst1 = New ADODB.Recordset
        rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
            rst1.MoveFirst
            If Not rst1.EOF Then
                ' vlozenie hodnot do excelu
                Call VlozRecordsetDoWorkbooku(wbExport, rst1)
            End If
        rst1.Close
    End If
    
    If Len(strWorkbookName) > 0 Then wbExport.Save
    
End Sub


Public Sub VlozRecordsetDoWorkbooku(ByRef WB As Workbook, ByRef rstX As ADODB.Recordset)
Dim WS As Worksheet
Dim intColIndex As Integer

Set WS = WB.Worksheets(1)

' vyplnime nazvy sloupecku do excelu
For intColIndex = 0 To rstX.Fields.count - 1
    WS.Range("A1").Offset(0, intColIndex).value = rstX.Fields(intColIndex).name
Next

' vyplnime hodnoty
WS.Range("A2").CopyFromRecordset rstX, 10000, 100 'max 10000 rows and 100 columns

' spravime z toho tabulku
Call SelectAllMakeTable(WS)

End Sub

Public Sub SelectAllMakeTable(ByRef WS As Worksheet)
' Sub vytvori excelovsku tabulku zo vsetkych dat vlozenych v danom worksheet
' malo by to byt to iste ako ked v exceli kliknes na bunku A1 a stlacis ctrl+T (skratka pre vytvorenie tabulky)
' nijak som ale netestoval chybajucie hodnoty, pocita sa s jednoduchou tabulkou zacinajucou v A1
    Dim tbl As ListObject
    Dim rng As Range

    Set rng = Range(WS.Range("A1"), WS.Range("A1").SpecialCells(xlLastCell))
    Set tbl = WS.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.TableStyle = "TableStyleMedium2"
End Sub

Public Function OmezPocetRadkuVQuery(ByVal sql As String, Optional pocetRadku As Integer = intOmezitPoctyKlientu) As String
' funkce zmeni sql dotaz tak, aby vracel maximalne zvoleny pocet radku
' pocet radku je defaultne nastaven na globalni konstantu intOmezitPoctyKlientu
' priklad vyuziti, kde sql je jakykoliv sql dotaz (string), staci zadat:
' sql = OmezPocetRadkuVQuery(sql)
' sql = OmezPocetRadkuVQuery(sql, 100)

OmezPocetRadkuVQuery = "SELECT * FROM (" & sql & ") WHERE ROWNUM <= " & pocetRadku

End Function

Public Sub NaplnAdminov()
' pouzivat pouze pro nedulezite veci, tohle neni bezpecne overeni uzivatele!
' pro overeni pouzit nasledujici postup:
'
' If strAdmins = "" Then
'    Call NaplnAdminov
' End If
'
' if InStr(strAdmins, UCase(Main.username)) > 0 then
'     ....
'

Dim sql As String
Dim i As Integer
sql = "SELECT a.username FROM " & schema & ".jez_users a WHERE a.user_role = 'ADMIN'"

strAdmins = ""

Set rst1 = New ADODB.Recordset
rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst1.EOF Then
        For i = 1 To rst1.RecordCount
            strAdmins = strAdmins & rst1![username] & ", "
            rst1.MoveNext
        Next i
    End If
    
rst1.Close

End Sub




Public Sub HighlightDuplicateCells(Values As Range)
'Funkce slouûÌ k podbarvenÌ (ûlutÏ) duplicitnÌch hodnot v oblasti bunÏk - nap¯. Range("A3", "A11")
Dim Cell
 
For Each Cell In Values
    If WorksheetFunction.CountIf(Values, Cell.value) > 1 Then
        Cell.Interior.ColorIndex = 6
    End If
 
Next Cell
End Sub


'RISK_468
Public Function ZiskejDatumRC(RodneCislo As String) As String
' zadej rodne cislo, funkce ti vrati datum narozeni
Dim rok As Integer, mesic As Integer, den As Integer, DatumNarozeni As String

rok = Left(RodneCislo, 2)
mesic = Mid(RodneCislo, 3, 2)
den = Mid(RodneCislo, 5, 2)
If mesic > 50 Then mesic = mesic - 50

If rok > 20 Then
    rok = rok + 1900
Else
    rok = rok + 2000
End If

ZiskejDatumRC = den & "." & mesic & "." & rok
End Function


Public Function NajdiSkpKDedicovi(ByVal id_person_heir As Variant, ByVal column_name_to_return As String) As Long
' zadej id_person_heir a funkce ti vrati id_person_dead nebo skp_heir (pro jinou moznost zatim jenom zasteka)
' predpoklada se (assertion) ze dedic ma nejvyse jednoho zemreleho
' zasteka pokud nema zadneho
' pokud by ses nudil, urcite to pujde nejak hezky zobecnit
' asi se nikde nepouziva

Dim sql_dead As String
Dim rst_dead As ADODB.Recordset
'NajdiSkpKDedicovi = 0 'inicialni hodnota pokud by fce nedobehla do konce

If column_name_to_return <> "id_person_dead" And column_name_to_return <> "skp_heir" Then
    MsgBox "chacha, nÏco si nedomyslel"
    Exit Function
End If

sql_dead = "SELECT h." & column_name_to_return & " FROM " & schema & ".heirs h WHERE h.flag_deleted = 'N' AND h.id_person_heir = '" & id_person_heir & "'"

Set rst_dead = New ADODB.Recordset
rst_dead.Open sql_dead, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
    'vic zemrelych k 1 dedicovi neni naimplementovano
    If rst_dead.RecordCount > 1 Then MsgBox "Tenhle dÏdic m· vÌce neû jednÈho zem¯elÈho. Automatick˝ proklik vybral jednoho n·hodnÏ.", vbInformation
    
    rst_dead.MoveFirst
    If rst_dead.RecordCount < 1 Then
        MsgBox "Nepovedlo se najÌt z·znam k dÏdicovi. Jez-error 00112"
        Exit Function
    End If
    If IsNull(rst_dead(0)) Then
        MsgBox "Nepovedlo se najÌt hodnotu k dÏdicovi (NULL). Jez-error 00113"
        Exit Function
    End If
    
    NajdiSkpKDedicovi = rst_dead(0)
    
rst_dead.Close
'?NajdiSkpZemrelehoKDedicovi("-30")
End Function



Public Sub OtevriKartuDedice(ByRef LV As ListView)
' Slouzi pro prokliky na konkretniho dedice na karte Klient_heir
' Samostatna sub, protoze na pozadi se nejdrive otevre Dedicke rizeni, az pak Heir rizeni (teda konkretni dedic)
' V dobe vytvoreni (20.12.2017) se vyuziva pri prokliku z Dedictvi.Listview1_DblClick a taky Souhrn.Listview3_DblClick


' XXXX nezabudni nastavit zavolanie SUB aj z dedictvi.listview1
Dim blnOK As Boolean
Dim skp_heir As String

    blnOK = False

    If (LV.name = "ListView1" And LV.Parent.name = "Dedictvi") Then

        '''   nejdrive si musime ulozit skp_heir na konkretneho dedice
        'stary zpusob: ' skp_heir = NajdiSkpKDedicovi(ListView1.SelectedItem.SubItems(1), "skp_heir") 'stary zpusob nefunguje pokud ma jeden dedic vic zemrelych (vybere nahodne)
        skp_heir = ListvZiskejHodnotu(LV, UCase("skp_heir"))

        '''   nastavime id_person na ZEMRELEHO (v listview je dedic!)
        'stary zpusob: 'id_person = NajdiSkpKDedicovi(ListView1.SelectedItem.SubItems(1), "id_person_dead") 'stary zpusob nefunguje pokud ma jeden dedic vic zemrelych (vybere nahodne)
        id_person = ListvZiskejHodnotu(LV, UCase("id_person_dead"))

        Call set_filter(Dedictvi.name, 1)
        Dedictvi.Hide
        blnOK = True
    
    ElseIf LV.name = "ListView3" And LV.Parent.name = "Souhrn" Then

        skp_heir = NajdiSkpKDedicovi(ListvZiskejHodnotu(LV, UCase("id_person")), "skp_heir") 'tenhle zpusob nefunguje pokud ma jeden dedic vic zemrelych (vybere nahodne)
        id_person = NajdiSkpKDedicovi(ListvZiskejHodnotu(LV, UCase("id_person")), "id_person_dead") 'stary zpusob nefunguje pokud ma jeden dedic vic zemrelych (vybere nahodne)
        blnOK = True
        
    Else
    
        MsgBox "JezevËÌk error jez-59901", vbCritical, "Ooooops!"
        blnOK = False
        
    End If

    If blnOK Then
    
        Klient_dedictvi.Show
        Call Klient_dedictvi.UserForm_Activate ' melo by se spustit samo s .Show, ale nespoustelo se, proto extra line

        'Tady musÌm nainicializovat dÏdice, rovnako jako pri DblClick z karty zemreleho
        Set c = New Client
        Set c = dedici.item(CStr(skp_heir)).objDedic 'XXXX Tady to p¯estalo fungovat, p¯i prokliku ze seznamu dÏdic˘... GRRR
        Set d = dedici.item(CStr(skp_heir))
        Klient_heir.Show
    
    End If


End Sub


Public Sub debug_risk_1145(Optional ByVal CUID As Long = 835900)
    ' cuid 835900 ma KTK (skp_overdraft 10)
    ' klient musi mat zadany nejaky EXE stav, inak to spadne (nenajde ho)
    ' KTK klienti cuid in ('835900', '61425', '389243', '238996', '52572', '336606', '429952', '208', '788279', '645705', '541242')
    If c Is Nothing Then
    
spustiZnovu:
        Call jj
        'Call Main.CommandButton1_Click
        With Exekuce
            .Show
            .TextBox1.text = CUID
            '.CommandButton11_Click
            .test_me (CUID)
            .ListView1_DblClick
        End With
        
        With Klient_exe
            Call Klient_exe.UserForm_Activate
        End With
        
    ElseIf c.CUID <> CUID Then
        
        GoTo spustiZnovu
    
    End If
    
    
    ' spusteni sablony
    Call exportuj_platby_KTK
    
    
End Sub



Public Sub debug_risk_977(Optional ByVal CUID As Long = 686212)
' spusti Jezevcika a na karte exekuce otvori testovaciho klienta a spusti sablonu
' je to moc prakticke, ale vetsinu pouzitych subs jsem musel predelat na PUBLIC, coz asi neni OK...


' Nejdrive posleme mejlik Pavlikovi, protoze nam kasle na ankety
' Call Mail_small_Text_Outlook

' povodny default klient: cuid 289544
' klient ma vela toho v casti II :  CUIDs: 457370, 583546, starsi nejaky: 55907
'                                          378900 (TK)
' Marta: 686212


If c Is Nothing Then

spustiZnovu:
    Call jj
    'Call Main.CommandButton1_Click
    With Exekuce
        .Show
        .TextBox1.text = CUID
        '.CommandButton11_Click
        .test_me (CUID)
        .ListView1_DblClick
    End With
    
    With Klient_exe
        Call Klient_exe.UserForm_Activate
    End With
    
ElseIf c.CUID <> CUID Then
    
    GoTo spustiZnovu

End If


' spusteni sablony
Call Klient_exe.CommandButton31_Click

End Sub



Public Sub debug_risk_1135(Optional ByVal CUID As Long = 673460)
' spusti Jezevcika a na karte exekuce otvori testovaciho klienta a pak klikne na Kontrolu Dat
' je to moc prakticke, ale vetsinu pouzitych subs jsem musel predelat na PUBLIC, coz asi neni OK...


    Call jj
    'Call Main.CommandButton1_Click
    With Exekuce
        .Show
        .TextBox1.text = CUID
        '.CommandButton11_Click
        .test_me (CUID)
        .ListView1_DblClick
    End With
    
    If c Is Nothing Then
        MsgBox "klient asi neni v exekucnim rizeni, pridej mu nejaky stav nebo si to hezky vyklikej."
        End
    End If
    
    With Klient_exe
        Call .UserForm_Activate
        Call .CommandButton25_Click
    End With

End Sub


Public Sub NaplnBookmark(ByVal strBookmarkName As String, ByVal strText As String, ByRef WordDocument As Variant, Optional ByVal blnBold As Boolean = False)
' Sub na ulahcenie vkladania hodnot do Wordovych sablon v roznych exportoch
' Sub vlozi textovy retazec do vybraneho bookmarku v danom Word dokumente
' (proste vlozi text kam treba...)
Dim rg As Range

    If WordDocument.bookmarks.Exists(strBookmarkName) Then
        With WordDocument.bookmarks(strBookmarkName).Range
            .text = strText
            If blnBold Then
                .Font.Bold = True
            End If
        End With
    Else
        Debug.Print "Bookmark not found: " & strBookmarkName
    End If
    
    
    'WordDocument.Bookmarks(strBookmarkName).Range.text = strText  ' v jedinom riadku
    
    '' povodna dvojriadkova verzia:
    'objDoc.Bookmarks("CastI").Select
    'wordapp.Selection.text = text1
    
End Sub


Public Sub VymazBookmarks(ByRef WordDocument As Object) 'Word.Document)
' vymaze vsechny bookmarks z daneho Word dokumentu

For i = WordDocument.bookmarks.count To 1 Step -1
    WordDocument.bookmarks(i).Delete
Next i

End Sub

Public Sub VypisVsechnyBookmarks(ByRef WordDocument As Object)
    ' funkce vypise nazvy vsech bookmarks v danem Wordovskem dokumentu (aby je programator nehledal jako blazen)
    
    Dim bookmark As Object
    Dim result As String
    
    For Each bookmark In WordDocument.bookmarks
        result = result & bookmark.name
        result = result & ", "
    Next bookmark
    
    Debug.Print Left(result, Len(result) - 2)
    
End Sub

Public Sub ZarovnejCelyDokumentDoBloku(ByRef WordDocument As Object) 'Object) 'Word.Document)
    ' metoda oznaci cely text ve Word dokumentu a zarovna jej do bloku (t.j. jako kdyz stlacis ctrl+a a ctrl+j)
    'dim WordDocument as word.docu
    
    WordDocument.Content.ParagraphFormat.Alignment = cfg.wdAlignParagraphJustify
    
End Sub

Public Sub PridajOdrazky(ByRef objDoc As Object) 'Word.Document)
' Sub vezme Wordovy dokument, prejde vsetky riadky a tie ktore zacinaju znakom "-" preformatuje na odrazky (bullet points / bullets)
' ignoruju sa medzery pred prvym znakom
' je to naivny pristup.
' pozor, je to fascinujuco POMALE! jedna strana zaberie niekolko sekund

Dim singleLine As Object 'Paragraph
Dim lineText As String

    For Each singleLine In objDoc.Paragraphs
        lineText = singleLine.Range.text
        If InStr(Left(Trim(singleLine), 1), "-") > 0 Then
            'MsgBox lineText 'vyzera ze funguje
            singleLine.Range.listformat.ApplyBulletDefault
'        ElseIf Left(singleLine, 4) = "    " Then
'            'snaha o odsazeni radku v casti II, ale nechce to delat to co ja chci..
'            singleLine = Replace(singleLine, "    ", "")
'            'singleLine.Range.listformat.listindent
'            singleLine.indent
'            singleLine.indent
        End If
    Next singleLine
    
End Sub


Public Sub OtevriPredchoziKartu(ByRef aktualniUsrFrm As Object)
' idealne sa pri kliknuti na tlacitko Navrat z ktorehokolvek rizeni vzdy zavola TATO sub
' vstup: userform z ktoreho sa vola
'        public premenna userform_prv


'Dim aktualniUsrFrm As UserForm
'Set aktualniUsrFrm = New Klient


Call aktualniUsrFrm.Hide 'nope

End Sub


Public Sub testInputBox()
Dim volbaInterestOverdue  As Boolean
'volbaInterestOverdue = Application.InputBox("Please select 1 to 3" & vbLf & vbLf & "1. Mary" & vbLf & "2. May" & vbLf & "3. Susan", "Enter a value", 1)
volbaInterestOverdue = (MsgBox("Pouûiù druh˙ variantu pre kapitalizovan˝ ˙rok?", vbYesNo) = vbYes)
Debug.Print volbaInterestOverdue
End Sub



Private Sub testRecordset()
Dim rs As New ADODB.Recordset
Dim sql As String
Dim data As Range
Dim arrData(0 To 2) As Variant
Dim sloupecky(0 To 2) As String
Dim i As Integer

'Set data = Sheets("DataSheet").Range("A1").CurrentRegion
'ReDim arrData(1 To 3, 1 To data.count / 3)
'arrData(1, ) = Sheets("DataSheet").Range("A1:A36").value
'arrData = data.value

For i = 0 To 2
    sloupecky(i) = Choose(i + 1, LCase("SKP_IP_SOUD"), LCase("KOD_IP_SOUD"), LCase("NAZEV_IP_SOUD"))
    arrData(i) = Choose(i + 1, 30, "FSNM", "Filipuv soud na Marsu")
Next i



sql = "select * from " & schema & ".IP_SOUDY t where 1=0"

Call jj

'On Error GoTo connection_problem
'On Error GoTo 0
blnTransactionStatus = cs_HDWMN.BeginTrans
    rs.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        'rs.AddNew sloupecky, arrData
        'rs.AddNew Array(LCase("SKP_IP_SOUD"), LCase("KOD_IP_SOUD"), LCase("NAZEV_IP_SOUD")), _
        '          Array(Array(300, 301), Array("FSNN", "FSNM"), Array("Filipuv soud na Marsu", "Filipuv soud na Marsu"))
'        rs.AddNew Array(LCase("SKP_IP_SOUD"), LCase("KOD_IP_SOUD"), LCase("NAZEV_IP_SOUD")), _
'                  Array(Array(301, "FSNM", "Filipuv soud na Marsu"), _
'                        Array(302, "FSNM", "Filipuv soud na Marsu"))
        rs.AddNew Array(LCase("SKP_IP_SOUD"), LCase("KOD_IP_SOUD"), LCase("NAZEV_IP_SOUD")), _
                  Array(Array(301, "FSNM", "Filipuv soud na Marsu"))
        rs.UpdateBatch
        
    rs.Close

'blnTransactionStatus =
cs_HDWMN.CommitTrans


'connection_problem:
    'Debug.Print "s"
    execute (sql)

End Sub

Public Sub TransBegin()
    ' funkce TransBegin, TransCommit a TransRollback by se meli pouzivat kvuli tomu, aby programator nezapomnel na promennou blnTransactionStatus
    ' promenna blnTransactionStatus je podstatna, na zaklade ni se pri ORA errorech spousti Rollback transakce
    blnTransactionStatus = cs_HDWMN.BeginTrans
End Sub

Public Sub TransCommit()
    ' funkce TransBegin, TransCommit a TransRollback by se meli pouzivat kvuli tomu, aby programator nezapomnel na promennou blnTransactionStatus

    cs_HDWMN.CommitTrans
    blnTransactionStatus = False
End Sub

Public Sub TransRollback()
    ' funkce TransBegin, TransCommit a TransRollback by se meli pouzivat kvuli tomu, aby programator nezapomnel na promennou blnTransactionStatus
    If blnTransactionStatus Then
        cs_HDWMN.RollbackTrans
        blnTransactionStatus = False
    End If
End Sub

Public Function ConvertSkptoId(ByVal SKP As Long) As Long

Dim sql As String
Dim rst As New ADODB.Recordset

sql = "select id_person from " & schema & ".person t where skp_person = " & SKP

    rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
        If rst.RecordCount = 1 Then
            ConvertSkptoId = rst![id_person]
        Else
            assertion 1 = 0, "jez-54584: konverze SKP do ID se nepovedla, nalezeno " & rst.RecordCount & " klient˘ s CUIDem " & SKP & "."
        End If
        
    rst.Close

End Function

Public Function ConvertCUIDtoId(ByVal CUID As Long) As Long
    
    Dim sql As String
    Dim rst As New ADODB.Recordset
    
    sql = "select id_person from " & schema & ".person t where CUID = " & CUID
    
    rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
        If rst.RecordCount = 1 Then
            ConvertCUIDtoId = rst![id_person]
        Else
            assertion 1 = 0, "jez-54574: konverze CUID do ID se nepovedla, nalezeno " & rst.RecordCount & " klient˘ s CUIDem " & CUID & "."
        End If
        
    rst.Close

End Function


Public Function VlozCastku(ByVal value) As String
    ' CIL: JEDNOTNY FORMAT VKLADANI CASTEK NAPRIC SABLONAMI
    
    'VlozCastku = format(value, "#,##0")
    'taky by slo:
    'VlozCastku = format(value, "currency")
    
    '''    VlozCastku = format(value, "#####,##0.##")
    '''    If value = Round(value, 0) Then
    '''        VlozCastku = VlozCastku & "- KË"
    '''    Else
    '''        VlozCastku = VlozCastku & " KË"
    '''    End If
    
    If value = Round(value, 0) Then
        VlozCastku = format(value, "#####,##0.##") & "- KË"
    Else
        VlozCastku = format(value, "#####,##0.00") & " KË"
    End If
    
End Function



'Public Function PausalVyska(ByVal in_id_person As String) As Double
'    ' funkce vrati evidovanu vysku pausalu ke klientovi
'    Dim sql As String
'    Dim rstPau As ADODB.Recordset
'
'        sql = "SELECT s.sop_additional_fee FROM " & schema & ".sop s WHERE s.id_person = " & in_id_person
'        Set rstPau = New ADODB.Recordset
'        rstPau.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'
'            If rstPau.RecordCount < 1 Then
'                HlaskaUzivateli "Ke klientovi nebyli nalezeny û·dnÈ z·znamy k SOPu nebo pauö·lu."
'                Exit Function
'            Else
'                If Not IsNull(rstPau![SOP_ADDITIONAL_FEE]) Then
'                    PausalVyska = Round(rstPau![SOP_ADDITIONAL_FEE], 2)
'                End If
'            End If
'
'        rstPau.Close
'
'End Function







Public Function VlozDatumDoXML(ByVal value As Variant) As String
    'value should be date or string
    'XMl vyzaduje datetime format, naoko to vezme cokoliv, ale pak jsou problemy u podepisovani nebo pri vymazani nejake prilohy
    VlozDatumDoXML = format(value, "yyyy-mm-dd")
End Function



Public Function VysklonujSoudGenitiv(strSoudReport) As String
    'pozor, fce pracuje s HlaskaUzivateli
    Dim polohaPredlozky As Integer
    
    If strSoudReport <> "" And strSoudReport <> "" And (InStr(strSoudReport, "SOUD_") = 0) And (strSoudReport <> cfg.cst_missing_string_value) Then 'pouze pokud se soud doplnil (nechceme duplicitu hlasku)
    
        If InStr(LCase(strSoudReport), "okres") <> 0 Then
            VysklonujSoudGenitiv = "OkresnÌho soudu "
        ElseIf InStr(LCase(strSoudReport), "kraj") <> 0 Then
            VysklonujSoudGenitiv = "KrajskÈho soudu "
        ElseIf InStr(LCase(strSoudReport), "obvod") <> 0 Then
            VysklonujSoudGenitiv = "ObvodnÌho soudu "
        ElseIf InStr(LCase(strSoudReport), "mÏst") <> 0 Then
            VysklonujSoudGenitiv = "MÏstskÈho soudu "
        Else
            HlaskaUzivateli ("Z nÏjakÈho d˘vodu se nepovedlo vyskloÚovat soud z reportu")
        End If
        
        'polohaPredlozky = WorksheetFunction.max(InStr(strSoudReport, " v"), InStr(strSoudReport, " pro"), InStr(strSoudReport, " PlzeÚ"))
        polohaPredlozky = InStr(strSoudReport, Split(strSoudReport, " ")(2)) - 1
        VysklonujSoudGenitiv = VysklonujSoudGenitiv & Right(strSoudReport, Len(strSoudReport) - polohaPredlozky)
        
    Else
    
        VysklonujSoudGenitiv = strSoudReport
    
    End If

End Function




Public Function DatumPosledniZaplaceneSplatky(ByVal skp_loan As Variant) As String
    ' funkce vrati DATUM POSLEDNI ZAPLACENE SPLATKY DANEHO UVERU
    ' nasledujici kus kodu byl testovan jenom zlehka..
    
    '''' alternativny sposob ale nefungoval vzdy
    '''Dim sql_posl_splatka As String
    '''Dim rst1 As ADODB.Recordset
    '''
    '''sql_posl_splatka = "select min(datum_splatky), max(datum_splatky) max_dtm_splatky, max(paid_in_time) from " & schema & ".sablona_epr6 where skp_loan = " & skp_loan & " and platba = splatka group by skp_loan"
    '''Set rst1 = New ADODB.Recordset
    '''rst1.Open sql_posl_splatka, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    '''
    '''    If Not rst1.EOF Then
    '''         rst1.MoveFirst
    '''         DatumPosledniZaplaceneSplatky = rst1![max_dtm_splatky]
    '''         'datum_prvni_splatky = rst1(0)
    '''    End If
    '''
    '''rst1.Close
    
    
    '''   nasledujuci kus kodu je prebraty z EPR_XML  -  vyuziva objekt Interest
    
    DatumPosledniZaplaceneSplatky = cfg.cst_missing_string_value
    
    If c.GetLoans(cL).GetCntInterests > 0 Then
            'Call pridejElement(elem, "//d:root/c_naroky_opak[0]/narok", "prislusenstvi_opak")
            'elemUrok = -1
            For cI = 1 To c.GetLoans(cL).GetCntInterests
                
                If c.GetLoans(cL).GetInterests(cI).UrokTyp = "KU" Then
                    DatumPosledniZaplaceneSplatky = c.GetLoans(cL).GetInterests(cI).StartInterval
                    Exit Function
                End If
             
            Next cI
    End If

End Function



Public Sub ExportWithHeaderFooter(ByRef WS As Worksheet) 'Kdyû nevÌö co, tak dej activesheet jako argument
'RISK-1103 'Funkce slouûÌ k vytvo¯enÌ AirBank hlaviËky a patiËky na EXCEL dokumentu.
'Abych ale vidÏl v˝sledek, tak si musÌm kliknout na VIEW a PAGE_LAYOUT

Const StrPicturePath = "W:\Vymahani\LATE\Vzory"

With WS.PageSetup

        .LeftHeader = "&G"
        .LeftHeaderPicture.fileName = StrPicturePath & "\HeaderAB.jpg" 'VloûenÌ air bank hlaviËky do Headeru
        
        .LeftFooter = "&G"
        .LeftFooterPicture.fileName = StrPicturePath & "\FooterAB.jpg" 'VloûenÌ air bank patiËky do Footeru

        .TopMargin = Application.InchesToPoints(1.87165354330709) 'MÌsto pro hlaviËku
        .BottomMargin = Application.InchesToPoints(0.984251968503937) 'MÌsto pro patiËku
        

End With

End Sub


Public Function nvl(ByVal value As Variant, Optional ByVal value_if_null As Variant = "") As Variant
    If IsNull(value) Then
        nvl = value_if_null
'    ElseIf value = "" Then
'        nvl = value_if_null
    Else
        nvl = value
    End If
End Function

'Public Sub assertion(ByVal a As Variant, ByVal b As Variant)
'    ' metoda debug.assert nedela nic, pokud uzivatel nema odheslovanej VBA projekt
'    ' proto vlastni metoda, ktera jezevce zhodi i na produkci
'    Debug.Assert a = b
'
'    If a = b Then
'        Exit Sub
'    Else
'        MsgBox "JezevËÌk error jez-84885. JezevËÌk se po odkliknutÌ okna zav¯e, poöli printscreen analytik˘m." & vbCrLf & a & vbCrLf & b, vbCritical, "JEZ-84885"
'        End
'    End If
'
'End Sub

Public Sub assertion(ByVal expression As Boolean, Optional ByVal error_text As String = "jez-84885: nesplnÏn nezn·m˝ p¯edpoklad.")
    ' metoda debug.assert nedela nic, pokud uzivatel nema odheslovanej VBA projekt
    ' proto vlastni metoda, ktera jezevce zhodi i na produkci
    
    Debug.Assert expression
    
    If expression Then
        Exit Sub
    Else
        
        ' uprava error_text
        If Not Right(error_text, 1) = "." Then
            error_text = error_text & "."
        End If
        
        
        MsgBox "JezevËÌk assertion error:" & vbCrLf & error_text & vbCrLf & _
               "JezevËÌk se po odkliknutÌ okna zav¯e, poöli printscreen analytik˘m.", vbCritical, "nesplnÏn nezn·m˝ p¯edpoklad"
        
        ' shozeni jezevcika
        End
        
    End If
    
End Sub


Public Function code_generator(ByVal colName As String) As String
' pro programatora, aby nemusel vypisovat sloupecky kdyz se plni z recordsetu
    code_generator = "p" & colName & " = nvl(rst![" & colName & "])"
End Function

Public Function Empty_string_highlighter(ByVal text As String, Optional ByVal highlighter As String = "__XXX__") As String
    ' pokud je vstupni string prazdny, funkce jej konvertuje na "XXXXXX"
    If highlighter = "__XXX__" Then highlighter = cfg.cst_missing_string_value  '  defaultni hodnota, nelze ale vlozit primo do definice
    
    If text = "" Then
        Empty_string_highlighter = highlighter
    Else
        Empty_string_highlighter = text
    End If
    
End Function

Public Function Empty_string_for_default_value(value As Variant, Optional ByVal default_numeric_value As Double = 0) As Variant
' defaultni hodnoty ruznych datovych typu (e.g. 0) konvertuje do prazdneho stringu ""
' ne-defaultni hodnoty konvertuje na STRING !!!

    Select Case LCase(TypeName(value))
    
        Case "string"
            Empty_string_for_default_value = value
            Exit Function
            
        Case "long", "double", "integer", "single"
            If value = default_numeric_value Then
                Empty_string_for_default_value = ""
            Else
                Empty_string_for_default_value = CStr(value)
            End If
            
        Case "date"
            If value = #12:00:00 AM# Then
                Empty_string_for_default_value = ""
            Else
                Empty_string_for_default_value = value
            End If
            
        Case Else
            Debug.Assert 1 = 0
            assertion 1 = 0
    
    End Select
End Function




Public Function ZabarviRadekListview(ByRef LV As ListView, ByVal rowTextToSearch As String, ByVal barvicka As barva) '  ByVal colName As String,
    
    ' funkce u daneho listview zabarvi na cerveno radky, ktere maji v prvnim sloupecku hodnotu rovnu rowTextToSearch
    
    ' test: ?ZabarviRadekListview(Klient.ListView4, "⁄daj", "Datum SOP1", barvaRudaSvetla)
    '       ?ZabarviRadekListview(Klient.ListView4, "⁄daj", "P¯eplatek (SOP1)", barvaRudaSvetla)
    
    
    Dim radek As Variant
    Dim i As Variant
    
    For Each radek In LV.ListItems
        ' Debug.Print radek & " ||| " & radek.SubItems(1)
        If Trim(radek.text) = (rowTextToSearch) Then
            
            radek.ForeColor = barvicka
            For i = 1 To LV.ListItems(1).ListSubItems.count
                radek.ListSubItems(i).ForeColor = barvicka
            Next i
            
        End If
    Next radek
    
End Function




Public Function CopyFile(srcFile As String, srcPath As String, destPath As String) As Boolean
'RISK-1114
'Funkce zkopÌruje soubor libovolnÈho form·tu z jednÈ sloûky do druhÈ
'Pokud nastane nÏjak· chyba - neplatn˝ soubor, sloûky nebo cokoliv jinÈho, tak vyhodÌ chybovou hl·öku
    
    If srcFile = "" Then GoTo ErrHandler
    On Error GoTo ErrHandler
    
    
    FileCopy srcPath & srcFile, destPath & srcFile
    
    CopyFile = True
    
    'MsgBox srcFile & " byl p¯esunut do " & destPath, vbInformation
    
    Exit Function
    
ErrHandler:
    
    CopyFile = False
    
    If Err.Number = 53 Or Err.Number = 58 Or srcFile = "" Then
        MsgBox "Pozor, nastala nÏjak· chyba p¯i kopÌrov·nÌ souboru " & srcFile, vbExclamation
        Exit Function
    End If

End Function



Public Function FindFileInFolder(srcPath As String, FileNameSubstr As String) As String
'RISK-1114
'Funkce vr·tÌ cestu k nalezenÈmu souboru
' update, vyhodi ho jenom pres HlaskuUzivateli (musi se zobrazit pozdeji hromadne pres VypisVsechnyHlasky) Pokud nenajde sloûku, nebo soubor, tak vyhodÌ upozornÏnÌ
' existuje taky funkce FindFileInFolderOrSubfolder

Dim fso As FileSystemObject, f As Object

Set fso = New FileSystemObject

If fso.FolderExists(srcPath) Then

    For Each f In fso.GetFolder(srcPath).Files
        If InStr(UCase(f.name), UCase(FileNameSubstr)) > 0 Then
            FindFileInFolder = f.name
            Exit For
        End If
    Next f

    ' If FindFileInFolder = "" Then MsgBox "Pozor, soubor " & FileNameSubstr & " nebyl nalezen", vbInformation ' FA vyhazuji, spam

Else
    
    HlaskaUzivateli "Sloûka nenalezena:" & vbCrLf & srcPath

End If

End Function



Public Function FindFileInFolderOrSubfolder(srcPath As String, FileNameSubstr As String, Optional ByVal FileNameSubstr2 As String = "", Optional ByVal blnPrintSearchedFolders = False, Optional ByVal FileNameSubstr_negative As String = "") As String
    'RISK-1159
    'Funkce vr·tÌ cestu k prvnÌmu nalezenÈmu souboru v cÌlovÈ sloûce NEBO KTER…KOLIV JEJÕ PODSLOéCE
    'Pokud nenajde sloûku, nebo soubor, tak vyhodÌ upozornÏnÌ
    ' existuje taky funkce FindFileInFolder
    ' cestu od nazvu souboru oddelist takhle:
    '   debug.print fso.GetFileName(strFile), fso.GetParentFolderName(strFile)
    
    Dim fso As FileSystemObject
    Dim f As Object
    Dim subfolder As Scripting.Folder

    Set fso = New FileSystemObject
    
    'If InStr(srcPath, "31122015") > 0 Then Debug.Assert False
    If blnPrintSearchedFolders Then Debug.Print "new call to function, srcPath: " & srcPath
    
    If fso.FolderExists(srcPath) Then
        
        For Each subfolder In fso.GetFolder(srcPath).SubFolders
            ' tady rekurze. Funkce skonci pri prvnim nalezenem souboru
            FindFileInFolderOrSubfolder = FindFileInFolderOrSubfolder(subfolder.Path, FileNameSubstr, FileNameSubstr2, blnPrintSearchedFolders, FileNameSubstr_negative)
            If FindFileInFolderOrSubfolder <> "" Then
                FindFileInFolderOrSubfolder = fso.BuildPath(subfolder.name, FindFileInFolderOrSubfolder)
                Exit Function
            End If
            
        Next subfolder
        
        For Each f In fso.GetFolder(srcPath).Files
            'If (InStr(UCase(f.name), UCase(FileNameSubstr)) > 0) And (InStr(UCase(f.name), UCase(FileNameSubstr2)) > 0) Then
            If (InStr(UCase(f.name), UCase(FileNameSubstr)) > 0) Then
                If (InStr(UCase(f.name), UCase(FileNameSubstr2)) > 0) Then
                    If (FileNameSubstr_negative = "") Or ((InStr(UCase(f.name), UCase(FileNameSubstr_negative)) = 0)) Then
                        FindFileInFolderOrSubfolder = f.name
                        Exit For
                    End If
                End If
            End If
        Next f
    
        ' If FindFileInFolder = "" Then MsgBox "Pozor, soubor " & FileNameSubstr & " nebyl nalezen", vbInformation ' FA vyhazuji, spam
    
    Else
        
        MsgBox "Sloûka nenalezena:" & vbCrLf & srcPath, vbExclamation
    
    End If

End Function



Public Function CreatePath(strPath As String) As Boolean
    ' nasledujuca fce spravi rekurzivni volani
    Dim fso As New FileSystemObject
    'Debug.Print "vytvaram pre cestu: " & strPath
        
        ' kontrola jestli nevytvarime z Testovacy slozky na PRODukci
        If (InStr(schema, "TEST") > 0) And (InStr(strPath, cfg.Cesta("TestovaciSlozka")) = 0) Then
            If Not MsgBox("UrËitÏ chceö vytv·¯et sloûky na produkci z testu? (bl·zne)", vbYesNo, "dotÏrn˝ dotaz") Then
                Call assertion(1 = 0, "Storno vytv·¯enÌ produkËnÌch sloûek z testu.")
            End If
        End If
        
    
        If fso.FolderExists(strPath) Then
            CreatePath = True
            Exit Function
        ElseIf Len(strPath) < 2 Then
            CreatePath = False
            Exit Function
        Else
            If Not fso.FolderExists(fso.GetParentFolderName(strPath)) Then
                If Not CreatePath(fso.GetParentFolderName(strPath)) Then
                    CreatePath = False
                    Exit Function
                End If
            End If
            fso.CreateFolder strPath
            ' Debug.Print fso.GetParentFolderName(strPath)
        End If
        
        If fso.FolderExists(strPath) Then CreatePath = True
End Function



Public Function get_enum_exekucni_titul_Name(exeTitul As enum_exekucni_titul, Optional ByVal blnGenitiv As Boolean = False)
    ' funkce vrati textovy retezec jako nazev daneho exekucniho titulu
    ' musi se shodovat s nazvami v podkladovem V_EXEKUCNI_TITUL v databazi
    ' tohle by melo byt jedine misto v excelu, kde jsou ulozeny presne stringy nazvu jednotlivych titulu
    ' pokud by se menily ve V_EXEKUCNI_TITUL, ze ktereho data vychazeji, melo by stacit upravit to zde
    
    Dim result As String
    
    Select Case exeTitul
    
        Case exe_titul_EPR
            result = IIf(blnGenitiv, "elektronickÈho platebnÌho rozkazu", "EPR")
        Case exe_titul_platebni_rozkaz
            result = IIf(blnGenitiv, "platebnÌho rozkazu", "PlatebnÌ rozkaz")
        Case exe_titul_rozsudek
            result = IIf(blnGenitiv, "rozsudku", "Rozsudek")
        Case exe_titul_usneseni_EPR
            result = IIf(blnGenitiv, "usnesenÌ", "UsnesenÌ-EPR")
        Case exe_titul_usneseni_C
            result = IIf(blnGenitiv, "usnesenÌ", "UsnesenÌ-C")
        Case exe_titul_none
            result = "" ' mozna zmenit na neco vhodnejsiho
        Case Else
            assertion 1 = 0, "jez-99101: Chyba v ËÌselnÌku exekuËnÌch titul˘." ' should never happen
    End Select
    
    get_enum_exekucni_titul_Name = result
    
End Function


Public Function SklonujSZ2Pad(ByVal StrOdesilatelDZ As String)
Dim StrMeziVypocet As String
'RISK-978
'Tato funkce se pouûÌv· v öablon·ch SonÌka a skloÚuje odesÌlatele datovÈ zpr·vy do druhÈho p·du - "na z·kladÏ û·dosti OkresnÌho st·tnÌho zastupitelstvÌ v PrachaticÌch"
'OdesÌlatel je v SonÌku uloûen v tabulce sonik_data_message a v public promÏnnÈ Odesilatel_datove_zpravy

If StrOdesilatelDZ = "N·rodnÌ centr·la proti organizovanÈmu zloËinu SKPV (PolicejnÌ prezidium »eskÈ republiky)" Then
    SklonujSZ2Pad = "N·rodnÌ centr·ly proti organizovanÈmu zloËinu SKPV (PolicejnÌho prezidia »eskÈ republiky)"
Else

    StrMeziVypocet = StrOdesilatelDZ 'UloûÌm si to do pomocnÈ promÏnnÈ aù nemÏnÌm p˘vodnÌ hodnotu
    'SELECT DISTINCT DATA_MESSAGE_SENDER from sonik_data_message order by DATA_MESSAGE_SENDER
    StrMeziVypocet = Replace(StrMeziVypocet, "KrajskÈ", "KrajskÈho")
    StrMeziVypocet = Replace(StrMeziVypocet, "MÏstskÈ", "MÏstskÈho")
    StrMeziVypocet = Replace(StrMeziVypocet, "MÏstsk˝", "MÏstskÈho")
    StrMeziVypocet = Replace(StrMeziVypocet, "N·rodnÌ", "N·rodnÌho")
    StrMeziVypocet = Replace(StrMeziVypocet, "ObvodnÌ", "ObvodnÌho")
    StrMeziVypocet = Replace(StrMeziVypocet, "OkresnÌ", "OkresnÌho")
    StrMeziVypocet = Replace(StrMeziVypocet, "VrchnÌ", "VrchnÌho")
    StrMeziVypocet = Replace(StrMeziVypocet, "NejvyööÌ", "NejvyööÌho")
    
    StrMeziVypocet = Replace(StrMeziVypocet, "st·tnÌ", "st·tnÌho")
    StrMeziVypocet = Replace(StrMeziVypocet, "soud", "soudu")
    StrMeziVypocet = Replace(StrMeziVypocet, "bezpeËnostnÌ", "bezpeËnostnÌho")
    StrMeziVypocet = Replace(StrMeziVypocet, "˙¯ad", "˙¯adu")
    
    SklonujSZ2Pad = StrMeziVypocet

End If

End Function



Public Sub PrintAllNamedRanges()
    ' do Immediate Window se vypisou vsechny pojmenovane oblasti
    
    Dim named_range As Variant
    Dim i As Long
    
    i = 0
    
    For Each named_range In ActiveWorkbook.Names
        i = i + 1
        Debug.Print i & ": " & named_range.name & Chr(10) & "   " & named_range
    Next named_range
    
End Sub




Public Sub exportuj_sheet(ByRef sheet_to_export As Worksheet, ByVal bln_export_do_PDF As Boolean, ByVal fileName As String, ByVal folder1 As String, _
                          Optional ByVal range_to_export As Range, Optional ByVal folder2 As String, Optional ByVal folder3 As String)
    
    ' obecna metoda, ktera exportuje cely sheet do PDF nebo EXCELu a ulozi do nektere ze specifikovanych slozek
    ' created by Filip Antol, 24.5.2018 do verze Jez.v.5.7
    
    ' bln_export_do_PDF = bud jako PDF, nebo jako novy excel workbook
    ' fileName          = nazov souboru (PDF nebo XML)
    ' folder1           = folder kam se ma ulozit
    ' range_to_export   = (optional) pokud se generuje PDF, tak se nevygeneruje cely worksheet ale jenom tenhle range (omezeni zbytecnych stranek navic)
    ' folder2           = (optional) folder kam se ma ulozit, pokud neexistuje folder1
    ' folder3           = (optional) folder kam se ma ulozit, pokud neexistuje folder1 ani folder2
    
    
    Dim cstPocetSlozek As Integer
    Dim fso As Object
    Dim currFolder As String
    Dim folder_final As String
    Dim wb_jez_name As String
    Dim bln_zobrazit_slozku As Boolean
    
    cstPocetSlozek = 3
    bln_zobrazit_slozku = True
    
    
    '''''     nastaveni slozky pro ulozeni
    
    
    If InStr(schema, "TEST") > 0 Then
        
        folder_final = cfg.Cesta("TestovaciSlozka")
        bln_zobrazit_slozku = False
        
    Else
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        For i = 1 To cstPocetSlozek
            
            currFolder = Choose(i, folder1, folder2, folder3)
            If fso.FolderExists(currFolder) Then
                
                folder_final = currFolder
                Exit For
            
            Else
                
                If i = cstPocetSlozek Then
                    
                    MsgBox "Sloûka pro klienta " & c.name & " neexistuje. V˝stup byl uloûen do " & cfg.Cesta("TestovaciSlozka") ', vbInformation, "Sloûka nenalezena!"
                    folder_final = cfg.Cesta("TestovaciSlozka")
                    
                End If
                
            End If
            
        Next i
        
    End If
    
    
    
    '''''   samotny export
    
    
    If bln_export_do_PDF Then
    
        ' metoda ExportAsFixedFormat funguje iba pre VISIBLE sheets
        sheet_to_export.Visible = xlSheetVisible
        If range_to_export Is Nothing Then
            sheet_to_export.ExportAsFixedFormat fileName:=folder_final & fileName & ".pdf", Type:=xlTypePDF, _
                                                Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        Else
            range_to_export.ExportAsFixedFormat fileName:=folder_final & fileName & ".pdf", Type:=xlTypePDF, _
                                                Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        End If
        'sheet_to_export.Visible = xlSheetHidden
    Else
        
        If Len(fileName) > 30 Then
            ' bohuzel excel nepovoli delsi nazev
            fileName = Right(fileName, 30)
        End If
        
        
        If wb_export_plateb Is Nothing Then
            ' create workbook for XML
            ' v pripade generovani plateb pro HU se workbook nainicializuje na karte Datacheck pri kliknuti na export
            Set wb_export_plateb = Workbooks.Add
            
            With wb_export_plateb
                .Title = fileName
                .Subject = fileName
                '.SaveAs fileName:=CreateObject("WScript.Shell").specialfolders("Desktop") & "\" & strNazev & ".xlsx" 'path to Desktop, regardless of user
                .SaveAs fileName:=folder_final & fileName & ".xlsx"
            End With
        End If
        
        ' Copy the worksheet
        wb_jez_name = ActiveWorkbook.name
        sheet_to_export.Visible = True
        sheet_to_export.Activate
        sheet_to_export.Select
        sheet_to_export.Copy after:=wb_export_plateb.Sheets(wb_export_plateb.Sheets.count)
        wb_export_plateb.Sheets(Sheets.count).name = fileName
        wb_export_plateb.Save
        
        Workbooks(wb_jez_name).Activate
        
    End If
    
    
    If bln_zobrazit_slozku Then
        Call Shell("explorer.exe " & folder_final, vbNormalFocus)
    End If
    
    
    
End Sub



Public Function SonikFormatDate(ByVal Datum As Variant) As String
'RISK-1131 'Funkce p¯evede vstup ve formÏ datumu (aù uû jako String nebo Date) do Stringu ve form·tu "d. m. yyyy"
'DobrÈ je, ûe funguje i p¯i¯azenÌ v˝stupu tÈto funkce do promÏnnÈ typu Date

'D·v·m variant protoûe chci, aby vstup mohl b˝t jak String tak Date

  
    If Datum = "" Then
        SonikFormatDate = "" 'Pojistka, kdyby tam padl nÏjak˝ pr·zdn˝ string
    Else
        SonikFormatDate = format$(Datum, "d. m. yyyy")
    End If
    
End Function




Public Sub prototype_error_nacteni_sablony()
    ' chci zachytit error abych vypsal trefnou hlasku pokud uzivatel nema povolene zmeny v nejake sablone
    ' nepovedlo se mi dostat error.
    
    Dim wordapp As Object
    Dim objDoc As Object
    Dim objSelection As Object
    
    
    Set wordapp = CreateObject("word.Application")
    
    '''   metoda Documents.Open zadny error nehaze.
    'Set objDoc = wordapp.Documents.Open("W:\Vymahani\LATE\Vzory\test_prototyp.dotx")
    
    On Error GoTo chyba
    Set objDoc = wordapp.Documents.Add(Template:="W:\Vymahani\LATE\Vzory\test_prototyp.dotx")
    On Error GoTo 0
    Set objSelection = wordapp.Selection
    
    
    wordapp.Visible = True
    wordapp.Activate
    
    objSelection.text = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    objDoc.Close
        
    Exit Sub
    
chyba:
    'MsgBox Err.Number, Err.Description
    If Err.Number = 13 And Err.Description = "Type mismatch" Then
        MsgBox "jez-00013: Vypad· to, ûe nem·ö povolenÈ zmÏny ve vzoru öablony. Je pot¯eba öablonu otev¯Ìt a zakliknout povolenÌ ˙prav (v ûlutÈ liötÏ navrchu). " & vbCrLf & _
               "Jedn· se o tenhle soubor: " & "W:\Vymahani\LATE\Vzory\test_prototyp.dotx" & vbCrLf & _
               "> v p¯ÌpadÏ, ûe öablona (.docx soubor) nelze spustit, je pot¯eba p¯es spr·vce ˙loh ukonËit vöechny procesy WINWORD.EXE.", vbInformation, "PovolenÌ ˙prav öablony"
    Else
        MsgBox "Doölo k nezn·mÈ chybÏ p¯i naËÌt·nÌ vzoru öablony."
    End If
End Sub


Public Function SonikSignature() As String
'RISK-1217
'Do SonÌka se nepodepisuji vûdy jen jeden ËlovÏk, ale vûdycky dva, podle toho, kdo je zrovna p¯ihl·öen˝
'Funkce nem· parametr, ale vezme z objektu JezUser p¯ihl·öenÈho uûivatele

Select Case JezUserLoggedIn.StrUserName
    
    'äablona Lucka s MarkÈtou - Lucka je na mate¯skÈ
    Case "AIR_KASPAROVA"
        SonikSignature = JezUserLoggedIn.StrCivilName & " / Specialistka souËinnosti s externÌmi subjekty" & vbCrLf & "Air Bank a.s. / Ëlen skupiny PPF" & vbCrLf & JezUserLoggedIn.StrPhoneNumber & " / soucinnost@airbank.cz" & vbCrLf & vbCrLf & "MarkÈta Doubkov· / Specialistka souËinnosti s externÌmi subjekty" & vbCrLf & "Air Bank a.s. / Ëlen skupiny PPF" & vbCrLf & "+420 547 134 820 / soucinnost@airbank.cz"

    'äablona MarkÈta s Markem
    Case "AIR_DOUBKOVA"
        SonikSignature = JezUserLoggedIn.StrCivilName & " / Specialistka souËinnosti s externÌmi subjekty" & vbCrLf & "Air Bank a.s. / Ëlen skupiny PPF" & vbCrLf & JezUserLoggedIn.StrPhoneNumber & " / soucinnost@airbank.cz" & vbCrLf & vbCrLf & "Marek Halcin / Specialista souËinnosti s externÌmi subjekty" & vbCrLf & "Air Bank a.s. / Ëlen skupiny PPF" & vbCrLf & "+420 547 134 855 / soucinnost@airbank.cz"
    
    'äablona Marka s MarkÈtou:
    Case "AIR_HALCIN"
        SonikSignature = JezUserLoggedIn.StrCivilName & " / Specialista souËinnosti s externÌmi subjekty" & vbCrLf & "Air Bank a.s. / Ëlen skupiny PPF" & vbCrLf & JezUserLoggedIn.StrPhoneNumber & " / soucinnost@airbank.cz" & vbCrLf & vbCrLf & "MarkÈta Doubkov· / Specialistka souËinnosti s externÌmi subjekty" & vbCrLf & "Air Bank a.s. / Ëlen skupiny PPF" & vbCrLf & "+420 547 134 820 / soucinnost@airbank.cz"
        
    'a ostatnÌ pouûÌvajÌ tento podpis 'äablona MarkÈta s Markem
    Case Else
        SonikSignature = "MarkÈta Doubkov· / Specialistka souËinnosti s externÌmi subjekty" & vbCrLf & "Air Bank a.s. / Ëlen skupiny PPF" & vbCrLf & "+420 547 134 820 / soucinnost@airbank.cz" & vbCrLf & vbCrLf & "Marek Halcin / Specialista souËinnosti s externÌmi subjekty" & vbCrLf & "Air Bank a.s. / Ëlen skupiny PPF" & vbCrLf & "+420 547 134 855 / soucinnost@airbank.cz"
        
End Select





End Function



Public Sub PrintAllControlsNamesOnUserform()

Dim Ctrl As Variant
Dim VstupniForm As UserForm

Set VstupniForm = Sonik_Karta_Klienta 'Tady si vyber co chces

For Each Ctrl In VstupniForm.Controls
        If Left(Ctrl.name, 4) = "List" Then
        
        Debug.Print Ctrl.name

        End If
Next Ctrl
End Sub

