VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Klient_heir 
   Caption         =   "Jezevèík - Karta Dìdice"
   ClientHeight    =   11775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   22815
   OleObjectBlob   =   "Klient_heir.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Klient_heir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit
'Private Const blnInfoKZemrelemu As Boolean = False ' obsah listview vlevo dolu (notar / zemrely)


Private Sub history_filter()
'RISK-468 novì vytaženo ze støev jezevèíka sem

Dim sql As String
Dim itmX As ListItem
Dim keys As Dictionary
Dim lvSI As ListSubItem


   Klient_heir.ListView_Heir_History.ListItems.Clear
   
   
   sql = "select skp_heir_note, dtime_note_created, type, text_note, detail, employee from " & schema & ".v_history_heir where skp_heir = " & d.lngSkp_heir & " and (1=0 "
   
   If Klient_heir.CheckBox_Stavy.value Then sql = sql & " or type = 'Stav' "
   If Klient_heir.CheckBox_Info.value Then sql = sql & " or type = 'Klient' "
   If Klient_heir.CheckBox_SOPy.value Then sql = sql & " or type = 'SOP' "
   If Klient_heir.CheckBox_Upo.value Then sql = sql & " or type = 'Upozornìní' "
   If Klient_heir.CheckBox_Pozn.value Then sql = sql & " or type = 'Poznámka' "
   
   sql = sql & ") order by dtime_note_created DESC, skp_heir_note DESC"
   
   Set keys = KliceKListview(Klient_heir.ListView_Heir_History)
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Klient_heir.ListView_Heir_History.ListItems.Add()
          itmX.text = "" & rst1![skp_heir_note] & ""
          itmX.SubItems(keys("Datum události")) = "" & rst1![dtime_note_created] & ""
          itmX.SubItems(keys("Typ události")) = "" & rst1![Type] & ""
          itmX.SubItems(keys("Událost")) = "" & rst1![TEXT_NOTE] & ""
          itmX.SubItems(keys("Detail")) = "" & rst1![detail] & ""
          itmX.SubItems(keys("Provedl")) = "" & rst1![EMPLOYEE] & ""
    rst1.MoveNext
    Next i
   End If
   
   rst1.Close
   
    Call ZabarviRadkyTextove(Klient_heir.ListView_Heir_History, keys("Typ události"), "Stav", barvaModra)
    Call ZabarviRadkyTextove(Klient_heir.ListView_Heir_History, keys("Typ události"), "Klient", barvaZelenaSvetla)
    Call ZabarviRadkyTextove(Klient_heir.ListView_Heir_History, keys("Typ události"), "Poznámka", barvaSeda)
    Call ZabarviRadkyTextove(Klient_heir.ListView_Heir_History, keys("Typ události"), "Upozornìní", barvaRuda)
    Call ZabarviRadkyTextove(Klient_heir.ListView_Heir_History, keys("Typ události"), "SOP", barvaFialova)


End Sub


Private Sub CommandButton_OpenFolder_Click()
'RISK-1001
'In this Example I am Opening a Folder ("C:\Temp")

 Dim fso
 Dim AFolder As String
 AFolder = cfg.Cesta("DEDICOVE") & Label_Name.Caption & " " & Label_CUID.Caption
 
Set fso = CreateObject("Scripting.FileSystemObject")

If Not fso.FolderExists(AFolder) Then
 'If Not FSO.FolderExists(BFolder) Then
  MsgBox "Složka pro klienta " & Label_Name.Caption & " neexistuje", vbInformation, "Složka nenalezena!"
 'Else
  'Call Shell("explorer.exe " & BFolder, vbNormalFocus)
 'End If
Else
 Call Shell("explorer.exe " & AFolder, vbNormalFocus)
End If

End Sub

Private Sub CommandButton_Uznani_Dluhu_Click()



If cs_HDWMN.State = 1 Then
prihlasen:


'Pøi úvodní iniciliaziaci dìdicù nehledám jejich loany ani úèty. Není to totiž potøeba
'Tady to ale musím udìlat, jinak procka export_uznani_dluhu spadne
'Dìdic sice nemá žádný svùj úvìr, ale stejnì to nainicializuji, vytvoøí se pak šablona jen se statickým textem.

Call c.find_accounts
Call c.find_loans

On Error GoTo chyba


   Application.EnableEvents = False
   Application.ScreenUpdating = False
   Application.DisplayAlerts = False


Call export_uznani_dluhu(id_person)



   Application.EnableEvents = True
   Application.ScreenUpdating = True
   Application.DisplayAlerts = True
   
' tady vypise pouze splatky, u kterych prisla platba do zesplatneni
chyba:

' pokud nastala chyba pøipojením, tak ...
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

Private Sub CommandButtonVlozUpo_Click()
If cs_HDWMN.State = 1 Then
prihlasen:

If prava("select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'NOTICE' and PRIVILEGE = 'INSERT'") Then

    Upozorneni_all.VstupnyUserform = Klient_heir
    Upozorneni_all.VstupNoveUpo = True
    Upozorneni_all.Show
    
Else
    Call text_information_hide(Me.name)
    
    Label_Text_Poznamka_Uzivateli.Caption = "Nemáte práva pro pøidávání upozornìní"
    Label_No.Visible = True
    Label_Text_Poznamka_Uzivateli.Visible = True
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


Private Sub ButtonZmenitStav_Click()
If cs_HDWMN.State = 1 Then
prihlasen:
 
If prava("select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'HEIR_STATUS_DATA' and PRIVILEGE = 'INSERT'") Then
    Update_stav_all.VstupnyUserform = Klient_heir
    Update_stav_all.Show
Else
    Call text_information_hide(Me.name)
    
    Label_Text_Poznamka_Uzivateli.Caption = "Nemáte práva pro zmìnu stavu"
    Label_No.Visible = True
    Label_Text_Poznamka_Uzivateli.Visible = True
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


Private Sub CommandButton_Back_Click()
Me.Hide
' FA fungovanie tychto preklikov je potrebne zjednotit napriec celym Jezevcikom (neviem ako je na tom Sonik)
' v tento moment to tak nejak funguje, ale je to nahoda
' nerozumiem preco odkaz_Klient sa zmeni az PO userform.Show
' rozne userforms maju zcela rozne podmienky, + vacsinou to maju duplicitne pre tlacitko navrat a pre Cancel Button
' idealne by bolo vytvorit jednu PUBLIC SUB, ktora by bola riadne napodmienkovana a na vsetkych kartach by uz sa volala iba ta SUB

Select Case userform_prv
    Case "Exekuce"
        Exekuce.Show
        odkaz_Klient = False
    Case "Zaloby"
        Klient.Show
        odkaz_Klient = True
    Case "Insol"
        Klient_insol.Show
        odkaz_Klient = True
    Case "Dedictvi", "Klient_heir"
        id_person = dedici.lngId_person_dead
        odkaz_Klient = False 'XXXX Bugííík
        Klient_dedictvi.Show
    Case "Zesplatneni"
        Zesplatneni.Show
        odkaz_Klient = True ' False
    Case "Nezesplatneni"
        Nezesplatneni.Show
        odkaz_Klient = False
    Case "Trest"
        Klient_trest.Show
        odkaz_Klient = True
    Case "SouhrnNotice", "SouhrnAuth", "Souhrn"
        Souhrn.Show
        odkaz_Klient = False
End Select

'XXXX - dodìlat prev form HEIR

TextBox1.text = ""

End Sub



Private Sub OptionButtonNotar_Click()
    With Klient_heir
        .ListView_Person_Dead.Top = .ListView_Notary.Top
        .ListView_Person_Dead.Visible = False
        .ListView_Notary.Visible = True
        .Label71 = "Notáø"
    End With
End Sub
Private Sub OptionButtonZemrely_Click()
    With Klient_heir
        .ListView_Notary.Visible = False
        .ListView_Person_Dead.Visible = True
        .Label71 = "Zemøelý"
    End With
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call CommandButton_Back_Click
    Cancel = True
End Sub




Private Sub ButtonVlozPoznamku_Click()

If prava("select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'HEIR_NOTE' and PRIVILEGE = 'INSERT'") Then
    
    Poznamka_all.VstupnyUserform = Klient_heir
    Poznamka_all.Show
    
Else
    Call text_information_hide(Me.name)
    
    Label_Text_Poznamka_Uzivateli.Caption = "Nemáte práva pro pøidání poznámky"
    Label_No.Visible = True
    Label_Text_Poznamka_Uzivateli.Visible = True
End If

End Sub



Private Sub CommandButton_Ded_Click()
'Zároveò nastavím globální promìnnou abych skoèil zpìt na zemøelého
    id_person = dedici.lngId_person_dead 'XXXX F: opraveno, melo by fungovat, netestoval jsem
    Klient_dedictvi.Show
    
End Sub

Private Sub CommandButton_Trest_Click()
'Zároveò nastavím globální promìnnou id_person abych se zbavil zemøelého a skoèil na dìdice pøi dalších proklicích
    Debug.Assert id_person = d.lngId_person_heir
    Klient_trest.Show
End Sub

Private Sub CommandButton_Insol_Click()
'Zároveò nastavím globální promìnnou id_person abych se zbavil zemøelého a skoèil na dìdice pøi dalších proklicích
    Debug.Assert id_person = d.lngId_person_heir
    Klient_insol.Show
End Sub

Private Sub CommandButton_Nalez_Click()
'Zároveò nastavím globální promìnnou id_person abych se zbavil zemøelého a skoèil na dìdice pøi dalších proklicích
    Debug.Assert id_person = d.lngId_person_heir

    Klient.Show
    
End Sub

Private Sub CommandButton_Exe_Click()
'Zároveò nastavím globální promìnnou id_person abych se zbavil zemøelého a skoèil na dìdice pøi dalších proklicích
    Debug.Assert id_person = d.lngId_person_heir
    Klient_exe.Show
End Sub





Private Sub CheckBox_Heir_Informed_Click()
'Uživatel mùže kliknout, pouze, pokud ještì u klienta rozhodnutí nepadlo a mùže dát YES nebo NO. V momentì, kdy se ale formuláø znovu naète, tak už mu editace nebude umožnìna
Dim sql As String

If d.strInformovan_o_nahravani = "X" Then 'Dovolím klik, nebo zmìnu hodnoty pouze tehdy, pokud ještì nebylo rozhodnuto - stav X

    If CheckBox_Heir_Informed Then 'Dávám tady tu podmínku, aby pøi prvním zadání mìl uživatel možnost ještì zmìnit rozhodnutí. Po reloadu obrazovky už to nepùjde
        sql = "update " & schema & ".heirs set FLAG_INFORMOVAN_O_NAHRAVANI = 'Y' where skp_heir = " & d.lngSkp_heir
    Else
        sql = "update " & schema & ".heirs set FLAG_INFORMOVAN_O_NAHRAVANI = 'N' where skp_heir = " & d.lngSkp_heir
    End If
    
    execute (sql)

End If

End Sub


Private Sub CheckBox_Stavy_Click()
Call history_filter
End Sub

Private Sub CheckBox_Info_Click()
Call history_filter
End Sub

Private Sub CheckBox_SOPy_Click()
Call history_filter
End Sub

Private Sub CheckBox_Upo_Click()
Call history_filter
End Sub

Private Sub CheckBox_Pozn_Click()
Call history_filter
End Sub




'Zachovávám pùvodní jméno, kvùli potenciálnímu problému s komplatibilitou u ostatních formù
Private Sub ListView_Heir_History_Click()
 Call text_information_hide(Me.name)
TextBox1.text = ""

If ListView_Heir_History.ListItems.count > 0 Then

 TextBox1.text = ListView_Heir_History.SelectedItem.SubItems(3)
 Label_Text_Poznamka_Uzivateli.Caption = "Zobrazena poznámka z pøehledu"
 Label_Yes.Visible = True
 Label_Text_Poznamka_Uzivateli.Visible = True
 
 End If
End Sub


Private Sub ListView_Heir_History_DblClick()
Dim sql As String


If cs_HDWMN.State = 1 Then
prihlasen:
TextBox1.text = ""
Call text_information_hide(Me.name)

If (ListView_Heir_History.ListItems.count > 0) Then
    If (ListvZiskejHodnotu(Klient_heir.ListView_Heir_History, "Detail") = "Ano" Or Len(ListvZiskejHodnotu(Klient_heir.ListView_Heir_History, "Událost")) > 130) Then
 
        sql = "select a.text_note from " & schema & ".heir_status_data a join " & schema & ".heir_note b on a.skp_heir_status_data = b.skp_heir_status_data and b.skp_heir_note = " & ListvZiskejHodnotu(Klient_heir.ListView_Heir_History, "ID")
        
        Set rst1 = New ADODB.Recordset
        rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
        Klient_heir.TextBox1.text = ""
        
        If Not rst1.EOF Then
            Klient_heir.TextBox1.text = "" & rst1![TEXT_NOTE] & ""
        End If
        
        rst1.Close
    
      
        Label_Text_Poznamka_Uzivateli.Caption = "Poznámka zobrazena"
        Label_Yes.Visible = True
        Label_Text_Poznamka_Uzivateli.Visible = True
    ElseIf ListvZiskejHodnotu(Klient_heir.ListView_Heir_History, "Detail") = "Ne" Then
        Label_Text_Poznamka_Uzivateli.Caption = "K poznámce nejsou další informace"
        Label_No.Visible = True
        Label_Text_Poznamka_Uzivateli.Visible = True
    Else
        Label_Text_Poznamka_Uzivateli.Caption = "Není pro co zobrazovat"
        Label_No.Visible = True
        Label_Text_Poznamka_Uzivateli.Visible = True
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



Private Sub ListView_Heir_Notice_DblClick()

If cs_HDWMN.State = 1 Then
prihlasen:

Call text_information_hide(Me.name)
If ListView_Heir_Notice.ListItems.count > 0 Then

    Upozorneni_all.VstupnyUserform = Klient_heir
    Upozorneni_all.VstupNoveUpo = False
    Upozorneni_all.Show

Else
    Label_Text_Poznamka_Uzivateli.Caption = "Není pro co zobrazovat"
    Label_Yes.Visible = True
    Label_Text_Poznamka_Uzivateli.Visible = True
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


Private Sub CommandButton_Chance_Import_Click()

Zadani_Sance.VstupnyUserform = Klient_heir
Zadani_Sance.Show

'XXXX Tabulka importu šance se bude muset rozšíøit o sloupec SKP_HEIR - aby to u dìdicù ukazovalo pouze šanci na daném dìdici ne na osobì. Ve výsledku se to ale stejnì bude vyhodnocovat na osobu.
'Bude také tøeba upravit kód na formu Zadani_Sance

'XXXX pak to nìjak padá, podívat se na to.

End Sub



Private Sub ZmenDedice(ByVal skp_heir As String)
'Nainicializovat klienta (na kterého jsem klikl - ten kdo je vyplnìný v daném optionButtonu a zavolat znovu Userform_activate())

Set c = New Client
Set c = dedici.item(skp_heir).objDedic

Set d = New Heir
Set d = dedici.item(skp_heir)
id_person = d.lngId_person_heir

Call UserForm_Activate

End Sub

Private Sub OptionButton_Heir1_Click()
Call ZmenDedice(CStr(OptionButton_Heir1.Tag)) ' skp_heir ulozeno na TAGu pod optionbuttonem OptionButton_Heir1.Tag
End Sub

Private Sub OptionButton_Heir2_Click()
Call ZmenDedice(CStr(OptionButton_Heir2.Tag))
End Sub

Private Sub OptionButton_Heir3_Click()
Call ZmenDedice(CStr(OptionButton_Heir3.Tag))
End Sub

Private Sub OptionButton_Heir4_Click()
Call ZmenDedice(CStr(OptionButton_Heir4.Tag))
End Sub

Private Sub OptionButton_Heir5_Click()
Call ZmenDedice(CStr(OptionButton_Heir5.Tag))
End Sub

Private Sub OptionButton_Heir6_Click()
Call ZmenDedice(CStr(OptionButton_Heir6.Tag))
End Sub

Private Sub OptionButton_Heir7_Click()
Call ZmenDedice(CStr(OptionButton_Heir7.Tag))
End Sub



' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()


'RISK-468 - nedržím se logiky Jezevèíka døíve - ale vše naèítám rovnou na této stránce

Dim i As Integer
Dim CB As control
Dim itmX As ListItem, clmx As ColumnHeader
Dim sql As String
Dim keys As Dictionary

userform_prv = Me.name
id_person = d.lngId_person_heir



If cs_HDWMN.State = 1 Then
prihlasen:
'----------------------- inicializace klienta pres objekt---------------
If odkaz_Klient = False Then
    Set c = New Client
    Call c.initialize(Val(id_person))
    'Call c.find_accounts
    'Call c.find_loans
End If
'----------------------- inicializace klienta pres objekt---------------
 

'PLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEK
' Nastavení základních údajù o osobì
Label_CUID = c.CUID
'Debug.Print "d " & c.name
Label_ID_PERSON = c.id_person
Label_RS = c.TextContractNumber
Label_Name = c.name
Label_RC = c.Pin
Label_Phone = c.PrimaryPhone
Label_Mail = c.PrimaryEmail
Label_ADRK = c.Address_C
Label_ADRT = c.Address_P
Label_Heir_Name = c.name
Label_Heir_Insolvency_Status.Visible = False
If Not c.InsolvencyStatus = vbNullString Then
    Label_Heir_Insolvency_Status.Visible = True
    Label_Heir_Insolvency_Status = "Dìdic v insolvenci (" & c.InsolvencyStatus & ")"
End If
Label_VS = d.strVariabilniSymbol

'PLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEK
' Nastavení Listview s pohledávkami

Klient_heir.ListView_Receivable.ListItems.Clear
If Klient_heir.ListView_Receivable.ColumnHeaders.count = 0 Then

Set clmx = Klient_heir.ListView_Receivable.ColumnHeaders.Add(, "Description", "Údaj", 180)
Set clmx = Klient_heir.ListView_Receivable.ColumnHeaders.Add(, "Value", "Hodnota", 80)
 
End If
 
Call NastavKliceListview(Klient_heir.ListView_Receivable)
 


Set keys = KliceKListview(Klient_heir.ListView_Receivable)
 

ListView_Receivable.ColumnHeaders(2).Alignment = lvwColumnRight

 Set itmX = Klient_heir.ListView_Receivable.ListItems.Add()
      itmX.text = "Celková výše pohledávky, kterou má dìdic zaplatit"
      itmX.SubItems(keys("Value")) = format(d.dblAmtReceivable, "### ### ##0.00") & " Kè"

 Set itmX = Klient_heir.ListView_Receivable.ListItems.Add()
      itmX.text = "Celková výše pohledávky, kterou již dìdic zaplatil"
      itmX.SubItems(keys("Value")) = format(d.dblAmtReceivablePaid, "### ### ##0.00") & " Kè"
 
 Set itmX = Klient_heir.ListView_Receivable.ListItems.Add()
      itmX.text = "Kolik má ještì dìdic doplatit"
      itmX.SubItems(keys("Value")) = format(d.dblAmtReceivable - d.dblAmtReceivablePaid, "### ### ##0.00") & " Kè"
      
Set itmX = Klient_heir.ListView_Receivable.ListItems.Add()

Set itmX = Klient_heir.ListView_Receivable.ListItems.Add()
      itmX.text = "Celková výše pohledávky"
      itmX.SubItems(keys("Value")) = format(dedici.dblTotalReceivable, "### ### ##0.00") & " Kè"

 Set itmX = Klient_heir.ListView_Receivable.ListItems.Add()
      itmX.text = "Celková výše pohledávky doplacená všemi dìdici "
      itmX.SubItems(keys("Value")) = format(dedici.dblTotalReceivablePaid, "### ### ##0.00") & " Kè"
 
 Set itmX = Klient_heir.ListView_Receivable.ListItems.Add()
      itmX.text = "Kolik mají dìdicové ještì doplatit"
      itmX.SubItems(keys("Value")) = format(dedici.dblTotalReceivable - dedici.dblTotalReceivablePaid, "### ### ##0.00") & " Kè"


'PLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEK
' Nastavení Listview s notáøi
'Potom se musí transformovat na nové struktury

Klient_heir.ListView_Notary.ListItems.Clear
If Klient_heir.ListView_Notary.ColumnHeaders.count = 0 Then
    
    Set clmx = Klient_heir.ListView_Notary.ColumnHeaders.Add(, "Description", "Údaj", 100)
    Set clmx = Klient_heir.ListView_Notary.ColumnHeaders.Add(, "Value", "Hodnota", 150)
     
End If
 
Call NastavKliceListview(Klient_heir.ListView_Notary)
 
 
 'RISK-468
Set keys = KliceKListview(Klient_heir.ListView_Notary)

Set itmX = Klient_heir.ListView_Notary.ListItems.Add()
    itmX.text = "Spisová znaèka"
    itmX.SubItems(1) = dedici.StrSpisovaZnacka

Set itmX = Klient_heir.ListView_Notary.ListItems.Add()
    itmX.text = "Jméno notáøe"
    itmX.SubItems(1) = dedici.objNotar.strName

Set itmX = Klient_heir.ListView_Notary.ListItems.Add()
    itmX.text = "Adresa notáøe"
    itmX.SubItems(1) = dedici.objNotar.strAddress

Set itmX = Klient_heir.ListView_Notary.ListItems.Add()
    itmX.text = "Telefonní èíslo notáøe"
    itmX.SubItems(1) = dedici.objNotar.StrPhoneNumber

Set itmX = Klient_heir.ListView_Notary.ListItems.Add()
    itmX.text = "ID datové schránky"
    itmX.SubItems(1) = dedici.objNotar.strIdDS


'PLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEK
' Nastavení Listview se zemøelým
' RISK-973

With Klient_heir.ListView_Person_Dead

    .ListItems.Clear
    If .ColumnHeaders.count = 0 Then
        Set clmx = .ColumnHeaders.Add(, "Description", "Údaj", 100)
        Set clmx = .ColumnHeaders.Add(, "Value", "Hodnota", 150)
    End If
    
    Call NastavKliceListview(Klient_heir.ListView_Person_Dead)
    
    
    
    Set keys = KliceKListview(Klient_heir.ListView_Person_Dead)
    
    Set itmX = .ListItems.Add()
        itmX.text = "Jméno zemøelého"
        itmX.SubItems(1) = dedici.objZemrely.name
    
    Set itmX = .ListItems.Add()
        itmX.text = "Datum narození"
        itmX.SubItems(1) = dedici.objZemrely.Date_Birth
    
    Set itmX = .ListItems.Add()
        itmX.text = "RÈ"
        itmX.SubItems(1) = dedici.objZemrely.TextBirthNumber
    
    Set itmX = .ListItems.Add()
        itmX.text = "Trvalá adresa"
        itmX.SubItems(1) = dedici.objZemrely.Address_P
    
    Set itmX = .ListItems.Add()
        itmX.text = "Datum úmrtí"
        itmX.SubItems(1) = dedici.objZemrely.DateDeath
    
    Set itmX = .ListItems.Add()
        itmX.text = "CUID zemøelého"
        itmX.SubItems(1) = dedici.objZemrely.CUID
    
End With



 


'PLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEK
' Nastavení Listview s historií

'Naplnìní hlavièek
Klient_heir.ListView_Heir_History.ListItems.Clear

If Klient_heir.ListView_Heir_History.ColumnHeaders.count = 0 Then

    Set clmx = Klient_heir.ListView_Heir_History.ColumnHeaders.Add(, , "ID", 30)
    Set clmx = Klient_heir.ListView_Heir_History.ColumnHeaders.Add(, , "Datum události", 90)
    Set clmx = Klient_heir.ListView_Heir_History.ColumnHeaders.Add(, , "Typ události", 60)
    Set clmx = Klient_heir.ListView_Heir_History.ColumnHeaders.Add(, , "Událost", 500)
    Set clmx = Klient_heir.ListView_Heir_History.ColumnHeaders.Add(, , "Detail", 30)
    Set clmx = Klient_heir.ListView_Heir_History.ColumnHeaders.Add(, , "Provedl", 100)
 
End If
    
Call NastavKliceListview(Klient_heir.ListView_Heir_History)
 
'Naplnìní dat
Set keys = KliceKListview(Klient_heir.ListView_Heir_History)
 
 
sql = "select skp_heir_note, dtime_note_created, type, text_note, detail, employee from " & schema & ".v_history_heir where skp_heir = " & d.lngSkp_heir
 
Set rst1 = New ADODB.Recordset
rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
     
    If Not rst1.EOF Then
      rst1.MoveFirst
      For i = 0 To rst1.RecordCount - 1
        Set itmX = Klient_heir.ListView_Heir_History.ListItems.Add()
              itmX.text = "" & rst1![skp_heir_note] & ""
              itmX.SubItems(keys("Datum události")) = "" & rst1![dtime_note_created] & ""
              itmX.SubItems(keys("Typ události")) = "" & rst1![Type] & ""
              itmX.SubItems(keys("Událost")) = "" & rst1![TEXT_NOTE] & ""
              itmX.SubItems(keys("Detail")) = "" & rst1![detail] & ""
              itmX.SubItems(keys("Provedl")) = "" & rst1![EMPLOYEE] & ""
       rst1.MoveNext
       Next i
    End If
     
rst1.Close

Call ZabarviRadkyTextove(Klient_heir.ListView_Heir_History, keys("Typ události"), "Stav", barvaModra)
Call ZabarviRadkyTextove(Klient_heir.ListView_Heir_History, keys("Typ události"), "Klient", barvaZelenaSvetla)
Call ZabarviRadkyTextove(Klient_heir.ListView_Heir_History, keys("Typ události"), "Poznámka", barvaSeda)
Call ZabarviRadkyTextove(Klient_heir.ListView_Heir_History, keys("Typ události"), "Upozornìní", barvaRuda)
Call ZabarviRadkyTextove(Klient_heir.ListView_Heir_History, keys("Typ události"), "SOP", barvaFialova)



'PLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEK
' Nastavení Listview s upozornìními

'Naplnìní hlavièek
Klient_heir.ListView_Heir_Notice.ListItems.Clear

If Klient_heir.ListView_Heir_Notice.ColumnHeaders.count = 0 Then

 Set clmx = Klient_heir.ListView_Heir_Notice.ColumnHeaders.Add(, , "ID", 30)
 Set clmx = Klient_heir.ListView_Heir_Notice.ColumnHeaders.Add(, , "Upozornìní", 480)
 Set clmx = Klient_heir.ListView_Heir_Notice.ColumnHeaders.Add(, , "Deadline", 100)
 Set clmx = Klient_heir.ListView_Heir_Notice.ColumnHeaders.Add(, , "Zbývá", 100)
 Set clmx = Klient_heir.ListView_Heir_Notice.ColumnHeaders.Add(, , "Assignee", 100)
 
End If
    
Call NastavKliceListview(Klient_heir.ListView_Heir_Notice)
 
'Naplnìní dat

Set keys = KliceKListview(Klient_heir.ListView_Heir_Notice)
 

sql = "select skp_notice, text_notice, dtime_notice_deadline, dtime_notice_deadline - trunc(sysdate) as days_left, assignee from " & schema & ".notice where code_status = 'a' and id_person = " & d.lngId_person_heir & " order by days_left"

Set rst1 = New ADODB.Recordset
rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

If Not rst1.EOF Then
 rst1.MoveFirst
 For i = 0 To rst1.RecordCount - 1
     Set itmX = Klient_heir.ListView_Heir_Notice.ListItems.Add()
     itmX.text = "" & rst1![skp_notice] & ""
     itmX.SubItems(keys("Upozornìní")) = "" & rst1![text_notice] & ""
     itmX.SubItems(keys("Deadline")) = "" & rst1![dtime_notice_deadline] & ""
     itmX.SubItems(keys("Zbývá")) = "" & rst1![days_left] & ""
     itmX.SubItems(keys("Assignee")) = "" & rst1![Assignee] & ""
 rst1.MoveNext
 Next i
End If
rst1.Close




'PLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEK
' Nastavení zaškrtávátka jestli byl klient informován o nahrávání

sql = "select FLAG_INFORMOVAN_O_NAHRAVANI from " & schema & ".HEIRS where SKP_HEIR = " & d.lngSkp_heir


Set rst1 = New ADODB.Recordset
rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
 
If Not rst1.EOF Then 'Oèekávám pouze jeden záznam
    If d.strInformovan_o_nahravani = "Y" Then
        CheckBox_Heir_Informed.value = True 'zdá se, že tento pøíkaz zavolá i proceduru Checkbox_Heir_click ale je ošetøeno podmínkou tam  :(...
        CheckBox_Heir_Informed.Enabled = False
'    ElseIf d.strInformovan_o_nahravani = "N" Then
'        CheckBox_Heir_Informed.value = False
'        CheckBox_Heir_Informed.Enabled = False
    Else
        CheckBox_Heir_Informed.value = False 'd.strInformovan_o_nahravani = "X"
        CheckBox_Heir_Informed.Enabled = True
    End If
End If
 
rst1.Close
 
 
 
 
 
'PLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEK
' Nastavení buttonù s ostatními dìdici nahoøe
For Each CB In Me.Controls
    If Left(CB.name, 17) = "OptionButton_Heir" Then
        CB.Visible = False
    End If
Next CB


If dedici.count > 0 Then

    For i = 1 To dedici.count


        For Each CB In Me.Controls
        
            If Left(CB.name, 17) = "OptionButton_Heir" And Right(CB.name, 1) = i Then
                CB.Visible = True
                CB.Caption = dedici.item(i).objDedic.name
                CB.Tag = dedici.item(i).lngSkp_heir 'Do tagu si vložím SKP_HEIR a potom když klient klikne na jméno co vidí, tak se odkážu pøes tag na objekt
                If CB.Tag = d.lngSkp_heir Then ' tady iniciálnì zaškrtnu vybraného dìdice
                    CB.value = True
                End If
            End If
        Next CB

    Next i

End If
 
 
'PLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEK
' Nastavení tlaèítek dole na kartì, podle toho jestli má v daných øízeních klient nìjaké stavy tak buï dám enabled nebo disabled






'PLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEK
'PLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEK

'PLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEK
'PLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEKPLÙTEK
' Starý kus kódu


''
'' HeisenBug - presunute na zacatek Sub
''If cs_HDWMN.State = 1 Then
''prihlasen:
'''----------------------- inicializace klienta pres objekt---------------
''If odkaz_Klient = False Then
''    'Dim c As Client
''    Set c = New Client
''    Call c.initialize(Val(id_person))
''    Call c.find_accounts
''    Call c.find_loans
''End If
'''----------------------- inicializace klienta pres objekt---------------
''
 
'CommandButton20.Visible = prava("select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'AUTHORIZATION' and PRIVILEGE = 'INSERT'")
 
 Call query1("", Me.name & "History")
 Call query1("", Me.name & "Other")
 
If ListView_Heir_History.ListItems.count > 0 Then
    Label_Text_Poznamka_Uzivateli.Caption = "Nalezeno " & ListView_Heir_History.ListItems.count & " záznamù"
    Label_Yes.Visible = True
    Label_Text_Poznamka_Uzivateli.Visible = True
    Label_No.Visible = False
Else
    Label_Text_Poznamka_Uzivateli.Caption = "Nalezeno " & ListView_Heir_History.ListItems.count & " záznamù"
    Label_No.Visible = True
    Label_Text_Poznamka_Uzivateli.Visible = True
End If
 
 
If userform_prv = "SouhrnAuth" Then
    userform_prv = "Souhrn"
    'Call CommandButton24_Click
End If


' VBA bug, listview haze do leveho horniho rohu, tohle je workaround: https://social.msdn.microsoft.com/Forums/en-US/9a65fc23-cdd4-4f34-b4c6-9528916afeb1/excel-2007-listview-positioning-problem?forum=isvvba
    Call OptionButtonZemrely_Click
    Call OptionButtonNotar_Click
' konec VBA bugu
If OptionButtonZemrely Then Call OptionButtonZemrely_Click ' Jez.v.5.3, RISK-973
If OptionButtonNotar Then Call OptionButtonNotar_Click

Klient_dedictvi.Hide ' FA 17.1.2018
 
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



Private Sub CommandButton_Edit_Heir_Click()
'RISK-468
'Zatím nefunguje, každopádnì pokud má kladný CUID tak nesmím dovolit úpravu údajù o klientovi - pouze výši pohledávky a variabilní symbol

Update_klient_heir.ZadaniNovyDedic = False
Update_klient_heir.Show

End Sub


Private Sub Label_ID_PERSON_Click()
Call CopyToClipboard(Label_ID_PERSON.Caption, Me.name)
Label_Text_Poznamka_Uzivateli.Caption = "ID_PERSON bylo zkopírovano do schránky"
End Sub

Private Sub Label_CUID_Click()
Call CopyToClipboard(Label_CUID.Caption, Me.name)
Label_Text_Poznamka_Uzivateli.Caption = "CUID byl zkopírovaný do schránky"
End Sub

Private Sub Label_Name_Click()
Call CopyToClipboard(Label_Name.Caption, Me.name)
Label_Text_Poznamka_Uzivateli.Caption = "Jméno bylo zkopírovano do schránky"
End Sub

Private Sub Label_RC_Click()
Call CopyToClipboard(Label_RC.Caption, Me.name)
Label_Text_Poznamka_Uzivateli.Caption = "Rodné èíslo bylo zkopírované do schránky"
End Sub

Private Sub Label_ADRT_Click()
Call CopyToClipboard(Label_ADRT.Caption, Me.name)
Label_Text_Poznamka_Uzivateli.Caption = "Trvalá adresa byla zkopírovaná do schránky"
End Sub

Private Sub Label_ADRK_Click()
Call CopyToClipboard(Label_ADRK.Caption, Me.name)
Label_Text_Poznamka_Uzivateli.Caption = "Kontaktní adresa byla zkopírovaná do schránky"
End Sub

Private Sub Label_RS_Click()
Call CopyToClipboard(Label_RS.Caption, Me.name)
Label_Text_Poznamka_Uzivateli.Caption = "Èíslo rámcové smlouvy bylo zkopírovano do schránky"
End Sub

Private Sub Label_Phone_Click()
Call CopyToClipboard(Label_Phone.Caption, Me.name)
Label_Text_Poznamka_Uzivateli.Caption = "Telefonní èíslo bylo zkopírováno do schránky"
End Sub

Private Sub Label_Mail_Click()
Call CopyToClipboard(Label_Mail.Caption, Me.name)
Label_Text_Poznamka_Uzivateli.Caption = "E-mailová adresa byla zkopírovaná do schránky"
End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
ButtonZmenitStav.ForeColor = barvaCerna
CommandButton_Edit_Heir.ForeColor = barvaCerna
CommandButton_Chance_Import.ForeColor = barvaCerna
CommandButton_Templates.ForeColor = barvaCerna
CommandButtonVlozUpo.ForeColor = barvaCerna
ButtonVlozPoznamku.ForeColor = barvaCerna
CommandButton_Nalez.ForeColor = barvaCerna
CommandButton_Trest.ForeColor = barvaCerna
CommandButton_Insol.ForeColor = barvaCerna
CommandButton_Exe.ForeColor = barvaCerna
CommandButton_Ded.ForeColor = barvaCerna

End Sub


Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Back.ForeColor = barvaCerna
End Sub

Private Sub CommandButton_Back_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Back.ForeColor = barvaZelenaAirBank
End Sub

Private Sub ButtonZmenitStav_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
ButtonZmenitStav.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Edit_Heir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Edit_Heir.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Chance_Import_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Chance_Import.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Templates_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Templates.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButtonVlozUpo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButtonVlozUpo.ForeColor = barvaZelenaAirBank
End Sub

Private Sub ButtonVlozPoznamku_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
ButtonVlozPoznamku.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Nalez_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Nalez.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Trest_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Trest.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Insol_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Insol.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Exe_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Exe.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Ded_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Ded.ForeColor = barvaZelenaAirBank
End Sub



