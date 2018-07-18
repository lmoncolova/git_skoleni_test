VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Klient_dedictvi 
   Caption         =   "Jezevèík - Dìdické øízení"
   ClientHeight    =   11775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   22815
   OleObjectBlob   =   "Klient_dedictvi.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Klient_dedictvi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit

Private Sub OverJestliNejsiNaDedicovi()
' Vsechny karty v Jezevcikovi pracuji s globalnim objektem c (Client), ve kterem je nacteny aktualni klient
' Pokud se pres zemreleho dostaneme na dedice, tak se do objektu c nainicializuje dedic, a cely Jezevcik funguje normalne... S VYNIMKOU KARTY DEDICTVI!
' Rekurze (Dedic zemre a dostane se do dedictvi a bude mit vlastni dedice...) NENI implementovana

End Sub

Sub history_filter()

If cs_HDWMN.State = 1 Then

prihlasen:
        Call text_information_hide(Me.name)
        
        Call query("", Me.name & "History")
     
    If err_connection Then
        err_connection = False
        Call re_login(Main.username.text, Main.password.text)
        If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
    End If

Else
    Call re_login(Main.username.text, Main.password.text)
    If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If
End Sub

Private Sub CommandButton_ded_prihlaska_Click()
    
    Call vytvor_prihlasku_do_DR
    
End Sub

'RISK-468
Private Sub CommandButton_Insert_Heir_Click()
Dim IntMaxPocetDedicu As Integer

IntMaxPocetDedicu = 7

If dedici.count >= IntMaxPocetDedicu Then
    MsgBox "Je mi líto, ale na zemøelém lze evidovat maximálnì " & IntMaxPocetDedicu & " dìdicù.", vbExclamation, "Nelze víc!"
    Exit Sub
End If

Update_klient_heir.ZadaniNovyDedic = True

Update_klient_heir.Show

End Sub

Private Sub CommandButton_Insert_Notary_Click()
    Update_notary.Show
End Sub

Private Sub CommandButton_Sonik_Click()
'RISK-721
 Call Client_From_Jez_To_Sonik(Klient_dedictvi.name, 3) ' RISK-985 odkomentovat po hromadnom REPLACE
End Sub

Private Sub CommandButton_Upravit_Spisovku_Click()
' FA: rozhodol som sa to dat extra (nie pod notara), pretoze to logicky dava viac zmysel.
' eventuelne to je vec ktora by mohla ist pod button 'Upravit informace o klientovi'

' nejdrive se vlozi novy zaznam do tabule INHERITANCE
' pak se nastavi flag_deleted v zaznamu ktery platil doted (pokud nejaky takovy je)

' TODO nejaku poznamku na klienta vlozit na konci 'xxxx

Dim strNoveCisloSpisovky As String
Dim sql_spis As String
Dim sqlMaxDedNote As String
Dim skp_ded_note As Long
Dim sql As String
Dim vntNotar As Variant ' bud cislo nebo null

If prava("select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'INHERITANCE' and PRIVILEGE = 'INSERT'") Then
    
    
    strNoveCisloSpisovky = Application.InputBox("Zadej nové èíslo spisové znaèky", "Zmìna spisové znaèky", dedici.StrSpisovaZnacka, Type:=2)
    
    '''   pokud da Cancel, nechceme nic mìnit
    If strNoveCisloSpisovky = "False" Then
        Exit Sub
    End If
    
    If dedici.lngSkp_Inheritance > 0 Then ' UZ MA ZAZNAM
        sql_spis = "Update " & schema & ".inheritance SET code = '" & strNoveCisloSpisovky & "' WHERE flag_deleted = 'N' AND skp_inheritance = " & dedici.lngSkp_Inheritance
    Else
        
        ' pokud se vklada spisovka ale jeste neni notar, chceme vlozit NULL
        If dedici.objNotar.lngSkp_notary > 0 Then
            vntNotar = dedici.objNotar.lngSkp_notary
        Else
            vntNotar = "NULL"
        End If
        
        sql_spis = "INSERT INTO " & schema & ".inheritance (skp_inheritance, code, id_person_dead,  skp_notary, flag_deleted, employee_inserted, dtime_inserted) " & _
                    "VALUES (" & schema & ".inheritance_seq.nextval, '" & strNoveCisloSpisovky & "', " & id_person & ", " & vntNotar & ", 'N', '" & Main.username & "', SYSDATE)"
    End If
    execute (sql_spis)
    sql_spis = ""
    
    'vlozit poznamku o zmene spisove znaèky - pouze pokud ale pøedtím již nìjaká byla
    
    If dedici.StrSpisovaZnacka <> vbNullString Then
    
         sqlMaxDedNote = "select max(skp_Ded_note) as max_skp_Ded_note from " & schema & ".Ded_NOTE"
    
         Set rst1 = New ADODB.Recordset
         rst1.Open sqlMaxDedNote, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
         
         If Not IsNull(rst1![max_skp_Ded_note]) Then
             skp_ded_note = rst1![max_skp_Ded_note] + 1
         Else
             skp_ded_note = 1
         End If
        
         rst1.Close
    
         sql = "insert into " & schema & ".ded_note (SKP_DED_NOTE, SKP_DED_NOTE_TYPE, ID_PERSON, TEXT_NOTE, DTIME_NOTE_CREATED, EMPLOYEE)" & _
             " values (" & skp_ded_note & ", 3, " & c.id_person & _
         ", 'Zmìna spisové znaèky: døíve " & dedici.StrSpisovaZnacka & " novì " & strNoveCisloSpisovky & "', sysdate, '" & UCase(Main.username) & "')"
         
         execute (sql)
    
    End If
    
    
    Application.ScreenUpdating = False
    Klient_dedictvi.Hide
    Klient_dedictvi.Show
    Application.ScreenUpdating = True
    
Else
    MsgBox "Je mi líto, ale nemáte práva na zmìnu spisové znaèky", vbExclamation, "Nedostateèná práva"
    
End If

End Sub

Private Sub ButtonZmenitStav_Click()
If cs_HDWMN.State = 1 Then
prihlasen:
 
If prava("select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'EPR_STATUS_DATA' and PRIVILEGE = 'INSERT'") Then
   Update_stav_all.VstupnyUserform = Klient_dedictvi
   Update_stav_all.Show
Else
    Call text_information_hide(Me.name)
    
    Label29.Caption = "Nemáte práva pro zmìnu stavu"
    Label28.Visible = True
    Label29.Visible = True
End If


If err_connection Then
    err_connection = False
    Call re_login(Main.username.text, Main.password.text)
    If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If
Else
    Call re_login(Main.username.text, Main.password.text)
    If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If
End Sub

Private Sub CommandButton16_Click()
Me.Hide

'If odkaz_Klient = False Then

Select Case userform_prv
    Case "Zaloby"
        Klient.Show
        odkaz_Klient = True
    Case "Exekuce"
        Klient_exe.Show
        odkaz_Klient = True
    Case "Dedictvi", "Klient_heir"
        Dedictvi.Show
        odkaz_Klient = False 'oprava MLI 8.2.2017
    Case "Insol"
        Klient_insol.Show
        odkaz_Klient = True
    Case "Trest"
        Klient_trest.Show
        Case "Zesplatneni"
        Zesplatneni.Show
        odkaz_Klient = False
    Case "Nezesplatneni"
        Nezesplatneni.Show
        odkaz_Klient = False
    Case "SouhrnNotice", "SouhrnAuth", "Souhrn"
        Souhrn.Show
        odkaz_Klient = False
End Select

'End If
'odkaz_Klient = False

TextBox1.text = ""
TextBox2.text = ""

End Sub



Private Sub CommandButton19_Click()
Dim BlnSpustitNaProdukci As Boolean
'RISK-468

BlnSpustitNaProdukci = True 'Dokud se nevyøeší úkol se sjednocením rodných èísel(RISK-517), tak nenasazovat. Jinak sice uloží RÈ, ale bude to k nièemu, protože pro klienty se pøemaže s denním loadem. A jen to vyvolá nekonzistenci mezi tabulkami PERSON a PERSON_DEMOGRAPHY
If BlnSpustitNaProdukci Then
Update_klient_ded.Show
Else: MsgBox "Provádìní zmìn na zemøelém je z mravních dùvodù momentálnì zakázáno", vbInformation, "Je mi líto"
End If
End Sub

Private Sub CommandButton20_Click()
'Call CommandButton24_Click 'does not exist
End Sub

Private Sub ButtonVlozPoznamku_Click()

If prava("select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'EPR_NOTE' and PRIVILEGE = 'INSERT'") Then
    
    Poznamka_all.VstupnyUserform = Klient_dedictvi
    Poznamka_all.Show
    
Else
    Call text_information_hide(Me.name)
    
    Label29.Caption = "Nemáte práva pro pøidání poznámky"
    Label28.Visible = True
    Label29.Visible = True
End If

End Sub


Private Sub CommandButton25_Click()
If cs_HDWMN.State = 1 Then
prihlasen:

' zobrazení informací

DataCheck.Show

 If err_connection Then
  err_connection = False
  Call re_login(Main.username.text, Main.password.text)
  If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
 End If
Else
 Call re_login(Main.username.text, Main.password.text)
 If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If
End Sub

Private Sub CommandButton26_Click()


'In this Example I am Opening a Folder ("C:\Temp")

 Dim fso
 Dim AFolder As String
 AFolder = cfg.Cesta("Late") & "DÌDICKÉ\" & c.Nazev_slozky_bez_lomitka
 'BFolder = cfg.Cesta("Late") & "EPR\NACHYSTANÉ\" & c.Nazev_slozky_bez_lomitka
 
Set fso = CreateObject("Scripting.FileSystemObject")

If Not fso.FolderExists(AFolder) Then
 'If Not FSO.FolderExists(BFolder) Then
  MsgBox "Složka pro klienta " & Label3.Caption & " neexistuje", vbInformation, "Složka nenalezena!"
 'Else
  'Call Shell("explorer.exe " & BFolder, vbNormalFocus)
 'End If
Else
 Call Shell("explorer.exe " & AFolder, vbNormalFocus)
End If
 
End Sub

Private Sub CommandButton27_Click()
odkaz_Klient = True

Klient.Show
End Sub

Private Sub CommandButton28_Click()
odkaz_Klient = True

Klient_trest.Show

End Sub

Private Sub CommandButton29_Click()
odkaz_Klient = True

Klient_insol.Show
End Sub

Private Sub CommandButton30_Click()
odkaz_Klient = True

Klient_exe.Show
End Sub


Private Sub CommandButtonVlozUpo_Click()
If cs_HDWMN.State = 1 Then
prihlasen:

If prava("select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'NOTICE' and PRIVILEGE = 'INSERT'") Then
    
    Upozorneni_all.VstupnyUserform = Klient_dedictvi
    Upozorneni_all.VstupNoveUpo = True
    Upozorneni_all.Show
    
Else
    Call text_information_hide(Me.name)
    
    Label29.Caption = "Nemáte práva pro pøidávání upozornìní"
    Label28.Visible = True
    Label29.Visible = True
End If
 
If err_connection Then
    err_connection = False
    Call re_login(Main.username.text, Main.password.text)
    If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If

Else 'ne-prihlasen:
    Call re_login(Main.username.text, Main.password.text)
    If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If
End Sub

Private Sub CheckBox1_Click()
Call history_filter
End Sub

Private Sub CheckBox2_Click()
Call history_filter
End Sub

Private Sub CheckBox3_Click()
Call history_filter
End Sub

Private Sub CheckBox4_Click()
Call history_filter
End Sub

Private Sub CheckBox5_Click()
Call history_filter
End Sub



'Private Sub Label2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Call text_information_hide(Me.name)
''Call CopyToClipboard(Label2.Caption, Me.name) tu si skoncil
'Dim clipboard As MSForms.DataObject
' Set clipboard = New MSForms.DataObject
' clipboard.SetText Label2.Caption
' clipboard.PutInClipboard
'
' Label29.Caption = "CUID byl zkopírovaný do schránky"
' Label27.Visible = True
' Label29.Visible = True
'End Sub
'
'Private Sub Label91_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Call text_information_hide(Me.name)
'
'Dim clipboard As MSForms.DataObject
' Set clipboard = New MSForms.DataObject
' clipboard.SetText Label91.Caption
' clipboard.PutInClipboard
'
' Label29.Caption = "Èíslo rámcové smlouvy bylo zkopírovano do schránky"
' Label27.Visible = True
' Label29.Visible = True
'End Sub
'
'Private Sub Label3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Call text_information_hide(Me.name)
'
'Dim clipboard As MSForms.DataObject
' Set clipboard = New MSForms.DataObject
' clipboard.SetText Label3.Caption
' clipboard.PutInClipboard
'
' Label29.Caption = "Jméno bylo zkopírovano do schránky"
' Label27.Visible = True
' Label29.Visible = True
'End Sub
'
'Private Sub Label4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Call text_information_hide(Me.name)
'
'Dim clipboard As MSForms.DataObject
' Set clipboard = New MSForms.DataObject
' clipboard.SetText Label4.Caption
' clipboard.PutInClipboard
'
' Label29.Caption = "Rodné èíslo bylo zkopírované do schránky"
' Label27.Visible = True
' Label29.Visible = True
'End Sub
'
'Private Sub Label5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Call text_information_hide(Me.name)
'
'Dim clipboard As MSForms.DataObject
' Set clipboard = New MSForms.DataObject
' clipboard.SetText Label5.Caption
' clipboard.PutInClipboard
'
' Label29.Caption = "Trvalá adresa byla zkopírovaná do schránky"
' Label27.Visible = True
' Label29.Visible = True
'End Sub
'
'Private Sub Label6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Call text_information_hide(Me.name)
'
'Dim clipboard As MSForms.DataObject
' Set clipboard = New MSForms.DataObject
' clipboard.SetText Label6.Caption
' clipboard.PutInClipboard
'
' Label29.Caption = "Kontaktní adresa byla zkopírovaná do schránky"
' Label27.Visible = True
' Label29.Visible = True
'End Sub
'
'
'Private Sub Label95_DblClick(ByVal Cancel As MSForms.ReturnBoolean) 'RISK-403
'Call text_information_hide(Me.name)
'
'Dim clipboard As MSForms.DataObject
' Set clipboard = New MSForms.DataObject
' clipboard.SetText Label95.Caption
' clipboard.PutInClipboard
'
' Label29.Caption = "Telefonní èíslo bylo zkopírováno do schránky"
' Label27.Visible = True
' Label29.Visible = True
'End Sub
'
'Private Sub Label96_DblClick(ByVal Cancel As MSForms.ReturnBoolean) 'RISK-403
'Call text_information_hide(Me.name)
'
'Dim clipboard As MSForms.DataObject
' Set clipboard = New MSForms.DataObject
' clipboard.SetText Label96.Caption
' clipboard.PutInClipboard
'
' Label29.Caption = "E-mailová adresa byla zkopírovaná do schránky"
' Label27.Visible = True
' Label29.Visible = True
'End Sub

'RISK-468 - odstranìno, kvùli vymazání pùvodního listview
'Private Sub ListView5_Click()
'
'TextBox2.text = ""
'
'If ListView5.ListItems.count > 0 Then
'For i = 1 To ListView5.ListItems.count
' If ListView5.ListItems(i).Selected = True Then
' TextBox2.text = ListView5.ListItems(i)
' End If
'Next i
'
'TextBox2.text = TextBox2.text & ": " & vbNewLine & vbNewLine & ListView5.SelectedItem.SubItems(1)
'
' End If
'End Sub

Private Sub ListView_Notary_Click()

TextBox2.text = ""

If ListView_Notary.ListItems.count > 0 Then
For i = 1 To ListView_Notary.ListItems.count
 If ListView_Notary.ListItems(i).Selected = True Then
 TextBox2.text = ListView_Notary.ListItems(i)
 End If
Next i

TextBox2.text = TextBox2.text & ":" & vbNewLine & vbNewLine & ListView_Notary.SelectedItem.SubItems(1)

 End If
End Sub

'RISK-468
Private Sub ListView_Heirs_Click()

TextBox2.text = ""

If ListView_Heirs.ListItems.count > 0 Then

    TextBox2.text = ListvZiskejHodnotu(ListView_Heirs, "Dìdic")
    TextBox2.text = TextBox2.text & ":" & vbNewLine & vbNewLine & ListvZiskejHodnotu(ListView_Heirs, "Aktuální stav")

End If

End Sub


'RISK-468
Private Sub ListView_Heirs_DblClick()
Dim skp_heir As String

If ListView_Heirs.ListItems.count > 0 Then
    skp_heir = ListvZiskejHodnotu(ListView_Heirs, "SKP_HEIR")

    'Tady musím nainicializovat dìdice
    Set c = New Client
    Set c = dedici.item(CStr(skp_heir)).objDedic
    
    Set d = dedici.item(CStr(skp_heir))
    
    
    Klient_heir.Show
    Klient_dedictvi.Hide ' FA 17.1.2018 - my nechceme aby bylo mozne se na dedice dostat posunutim okna navrchu, ale vzdy jenom pres nejaky CommandButton

End If

End Sub



Private Sub ListView1_Click()
 Call text_information_hide(Me.name)
TextBox1.text = ""

If ListView1.ListItems.count > 0 Then

 TextBox1.text = ListView1.SelectedItem.SubItems(3)
 Label29.Caption = "Zobrazena poznámka z pøehledu"
 Label27.Visible = True
 Label29.Visible = True
 End If
End Sub

Private Sub ListView1_DblClick()
Dim sql As String


If cs_HDWMN.State = 1 Then
prihlasen:
TextBox1.text = ""
 Call text_information_hide(Me.name)
If ListView1.ListItems.count > 0 Then
    If ((ListView1.SelectedItem.SubItems(4) = "Ano") Or (Len(ListView1.SelectedItem.SubItems(3)) > 130)) Then
 
     
        sql = "select a.text_note from " & schema & ".ded_status_data a join " & schema & ".ded_note b on a.skp_ded_status_data = b.skp_ded_status_data and b.skp_ded_note = " & ListView1.SelectedItem
         
        Call query(sql, "Note_ded")
        
         
        Label29.Caption = "Poznámka zobrazena"
        Label27.Visible = True
        Label29.Visible = True
    
     

    
    ElseIf ListView1.SelectedItem.SubItems(4) = "Ne" Then
        Label29.Caption = "K poznámce nejsou další informace"
        Label28.Visible = True
        Label29.Visible = True
    Else
        Label29.Caption = "Není pro co zobrazovat"
        Label28.Visible = True
        Label29.Visible = True
    End If
End If
 
 If err_connection Then
  err_connection = False
  Call re_login(Main.username.text, Main.password.text)
  If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
 End If
Else
 Call re_login(Main.username.text, Main.password.text)
 If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If
End Sub


Private Sub ListView2_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
If (KeyCode = 67 And Shift = 2) Then

Call text_information_hide(Me.name)

Dim clipboard As MSForms.DataObject
Set clipboard = New MSForms.DataObject
clipboard.SetText ListView2.SelectedItem
clipboard.PutInClipboard

 Label29.Caption = "Èíslo smlouvy bylo zkopírovano do schránky"
 Label27.Visible = True
 Label29.Visible = True

End If
End Sub

Private Sub ListView3_DblClick()

If cs_HDWMN.State = 1 Then
prihlasen:

Call text_information_hide(Me.name)
If ListView3.ListItems.count > 0 Then

    Upozorneni_all.VstupnyUserform = Klient_dedictvi
    Upozorneni_all.VstupNoveUpo = False
    Upozorneni_all.Show

Else
    Label29.Caption = "Není pro co zobrazovat"
    Label28.Visible = True
    Label29.Visible = True
End If

If err_connection Then
    err_connection = False
    Call re_login(Main.username.text, Main.password.text)
    If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If
Else
    Call re_login(Main.username.text, Main.password.text)
    If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If

End Sub



Private Sub ListView_Notary_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
If (KeyCode = 67 And Shift = 2) Then

Call text_information_hide(Me.name)

Dim clipboard As MSForms.DataObject
Set clipboard = New MSForms.DataObject
clipboard.SetText ListView_Notary.SelectedItem.SubItems(1)
clipboard.PutInClipboard

 Label29.Caption = "Hodnota '" & ListView_Notary.SelectedItem & "' byla zkopírovana do schránky"
 Label27.Visible = True
 Label29.Visible = True

End If
End Sub

Private Sub NaplnListView_Notary()
Dim itmX As ListItem
 Klient_dedictvi.ListView_Notary.ListItems.Clear

 Set itmX = Klient_dedictvi.ListView_Notary.ListItems.Add()
 itmX.text = "Spisová znaèka"
 itmX.SubItems(1) = dedici.StrSpisovaZnacka

 Set itmX = Klient_dedictvi.ListView_Notary.ListItems.Add()
 itmX.text = "Jméno notáøe"
 itmX.SubItems(1) = dedici.objNotar.strName

  Set itmX = Klient_dedictvi.ListView_Notary.ListItems.Add()
 itmX.text = "Adresa notáøe"
 itmX.SubItems(1) = dedici.objNotar.strAddress

  Set itmX = Klient_dedictvi.ListView_Notary.ListItems.Add()
 itmX.text = "Telefonní èíslo notáøe"
 itmX.SubItems(1) = dedici.objNotar.StrPhoneNumber
 
  Set itmX = Klient_dedictvi.ListView_Notary.ListItems.Add()
 itmX.text = "ID datové schránky"
 itmX.SubItems(1) = dedici.objNotar.strIdDS

End Sub

' nasledujuca sub je Public, pretoze sa vola extra v Dedictvi.Listview1_DblClick (pri dvojkliku na dedice (ne zemreleho))
Public Sub UserForm_Activate()

' Variables for ListView component
Dim clmx As ColumnHeader
Dim itmX As ListItem
Dim sql1 As String

If cs_HDWMN.State = 1 Then
prihlasen:

userform_prv = "Dedictvi"

'----------------------- inicializace klienta pres objekt---------------
If odkaz_Klient = False Then
    'Dim c As Client
    Set c = New Client
    Call c.initialize(Val(id_person))
    Call c.find_accounts
    Call c.find_loans
End If



'RISK-468
Set dedici = New Heirs

Call dedici.NajdiVsechnyDedice(c.id_person)

'RISK-468
Call table_prepare(Me.name)


 'RISK-468
  If Klient_dedictvi.ListView_Notary.ColumnHeaders.count = 0 Then
 
   Set clmx = Klient_dedictvi.ListView_Notary.ColumnHeaders.Add(, , "Údaj", 100)
   Set clmx = Klient_dedictvi.ListView_Notary.ColumnHeaders.Add(, , "Hodnota", 152)
 
  End If
  
 If Klient_dedictvi.ListView_Heirs.ColumnHeaders.count = 0 Then
 
    Set clmx = Klient_dedictvi.ListView_Heirs.ColumnHeaders.Add(, , "SKP_HEIR", 0)
    Set clmx = Klient_dedictvi.ListView_Heirs.ColumnHeaders.Add(, , "ID_PERSON", 0)
    Set clmx = Klient_dedictvi.ListView_Heirs.ColumnHeaders.Add(, , "Dìdic", 132)
    Set clmx = Klient_dedictvi.ListView_Heirs.ColumnHeaders.Add(, , "Aktuální stav", 120)
 
 End If
 
 Call NastavKliceListview(Klient_dedictvi.ListView_Heirs)
 
 'RISK-468
Dim keys As Dictionary
Set keys = KliceKListview(Klient_dedictvi.ListView_Heirs)
 
 
Klient_dedictvi.ListView_Heirs.ListItems.Clear
 
For i = 1 To dedici.count
 
    Set itmX = Klient_dedictvi.ListView_Heirs.ListItems.Add()
         itmX.text = dedici.item(i).lngSkp_heir
         itmX.SubItems(keys("Dìdic")) = dedici.item(i).objDedic.name
         itmX.SubItems(keys("Aktuální stav")) = dedici.item(i).strStavAktualni
         itmX.SubItems(keys("ID_PERSON")) = dedici.item(i).lngId_person_heir
    
 Next i


'RISK-468
Call NaplnListView_Notary

' info o vyske SOPu

Me.Label_sop_value.Visible = False
Me.Label_sop_text.Visible = False

sql1 = "select priznan_sop, sop_paid from " & schema & ".v_sop_balance where id_person = " & id_person
Set rst1 = New ADODB.Recordset
rst1.Open sql1, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
    If Not rst1.EOF Then
        If Not IsNull(rst1![priznan_sop]) Then
            'If rst1![priznan_sop] > 0 Then
                
                Me.Label_sop_value.Visible = True
                Me.Label_sop_text.Visible = True
                Me.Label_sop_value = Round(rst1![priznan_sop] - rst1![sop_paid], 2) & " Kè"
                
            'End If
        End If
    End If
rst1.Close




'----------------------- inicializace klienta pres objekt---------------

 Call text_information_hide(Me.name)
 
' CommandButton20.Visible = prava("select * from USER_TAB_PRIVS where owner = '" & schema & "' and table_name = 'AUTHORIZATION' and PRIVILEGE = 'INSERT'")
 
 Call query("", Me.name & "History")
 Call query("", Me.name & "Other")
 
If ListView1.ListItems.count > 0 Then
 Label29.Caption = "Nalezeno " & ListView1.ListItems.count & " záznamù"
 Label27.Visible = True
 Label29.Visible = True
Else
 Label29.Caption = "Nalezeno " & ListView1.ListItems.count & " záznamù"
 Label28.Visible = True
 Label29.Visible = True
End If

 
 If userform_prv = "SouhrnAuth" Then
  userform_prv = "Souhrn"
  'Call CommandButton24_Click
 End If
 
 
Else
    Call re_login(Main.username.text, Main.password.text)
    If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If

'RISK-721
 ' RISK-985 odkomentovat po hromadnom REPLACE
If Sonik_User Then
CommandButton_Sonik.Enabled = True
Else: CommandButton_Sonik.Enabled = False
End If
''' ' RISK-985 odkomentovat po hromadnom REPLACE

End Sub



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Me.Hide

'If odkaz_Klient = False Then
Select Case userform_prv
 Case "Zaloby"
  Klient.Show
    odkaz_Klient = True
Case "Exekuce"
  Klient_exe.Show
      odkaz_Klient = True
 Case "Dedictvi"
  Dedictvi.Show
      odkaz_Klient = False
Case "Insol"
  Klient_insol.Show
      odkaz_Klient = True
Case "Trest"
  Klient_trest.Show
  odkaz_Klient = True
  Case "Zesplatneni"
  Zesplatneni.Show
  odkaz_Klient = False
 Case "Nezesplatneni"
  Nezesplatneni.Show
  odkaz_Klient = False
 Case "SouhrnNotice", "SouhrnAuth", "Souhrn"
  Souhrn.Show
  odkaz_Klient = False
End Select
'End If

'odkaz_Klient = False

TextBox1.text = ""
TextBox2.text = ""

End Sub


Private Sub Label1_Click()
Call CopyToClipboard(Label1.Caption, Me.name)
Label29.Caption = "ID_PERSON bylo zkopírovano do schránky"
End Sub

Private Sub Label2_Click()
Call CopyToClipboard(Label2.Caption, Me.name)
Label29.Caption = "CUID byl zkopírovaný do schránky"
End Sub

Private Sub Label3_Click()
Call CopyToClipboard(Label3.Caption, Me.name)
Label29.Caption = "Jméno bylo zkopírovano do schránky"
End Sub

Private Sub Label4_Click()
Call CopyToClipboard(Label4.Caption, Me.name)
Label29.Caption = "Rodné èíslo bylo zkopírované do schránky"
End Sub

Private Sub Label5_Click()
Call CopyToClipboard(Label5.Caption, Me.name)
Label29.Caption = "Trvalá adresa byla zkopírovaná do schránky"
End Sub

Private Sub Label6_Click()
Call CopyToClipboard(Label6.Caption, Me.name)
Label29.Caption = "Kontaktní adresa byla zkopírovaná do schránky"
End Sub

Private Sub Label91_Click()
Call CopyToClipboard(Label91.Caption, Me.name)
Label29.Caption = "Èíslo rámcové smlouvy bylo zkopírovano do schránky"
End Sub


Private Sub Label95_Click()
Call CopyToClipboard(Label95.Caption, Me.name)
Label29.Caption = "Telefonní èíslo bylo zkopírováno do schránky"
End Sub

Private Sub Label96_Click()
Call CopyToClipboard(Label96.Caption, Me.name)
Label29.Caption = "E-mailová adresa byla zkopírovaná do schránky"
End Sub



Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton19.ForeColor = barvaCerna
CommandButtonVlozUpo.ForeColor = barvaCerna
ButtonVlozPoznamku.ForeColor = barvaCerna
End Sub


Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
'CommandButton20.ForeColor = barvaCerna
End Sub

Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton16.ForeColor = barvaCerna
End Sub


Private Sub CommandButton16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton16.ForeColor = barvaZelenaAirBank
End Sub


Private Sub CommandButton19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton19.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButtonVlozUpo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButtonVlozUpo.ForeColor = barvaZelenaAirBank
End Sub

Private Sub ButtonVlozPoznamku_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
ButtonVlozPoznamku.ForeColor = barvaZelenaAirBank
End Sub



