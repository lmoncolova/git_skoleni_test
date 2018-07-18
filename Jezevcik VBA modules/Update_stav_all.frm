VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Update_stav_all 
   Caption         =   "UserForm1"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17655
   OleObjectBlob   =   "Update_stav_all.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Update_stav_all"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit
' UserForm zavedeny 11/2017, zjednotenie 5 userformov do jedneho univerzalneho pre vsetky rizeni
' Userform funguje podobne ako Poznamka_all a Upozorneni_all
' pred zavolanim userformu je potrebne nastavit kartu z ktorej sa zavola, t.j. Klient / Klient_dedictvi / Klient_insol / ...
' pokud prijde pozadavek od vymahacu, ze chteji aby se nejaky stav propagoval mezi rizenimi, lze to docela jednoduse pridat - viz vyuziti Sub ExtraEPR
'
'
'
''    Premenovanie controls oproti povodnym 5 userforms:
''CommandButton20 -> ButtonUlozit
''CommandButton12 -> ButtonZavrit
''CommandButton21 -> ButtonSmazatStav
''CommandButton16 -> ButtonNavrat
''CheckBox2 -> CheckBoxZobrazitSmazane
''CheckBox1 -> CheckBoxAktualni
''ComboBox1 -> ComboBoxStav
''DTPicker1 -> DTPickerPlatnyOd
''TextBox1 -> TextBoxPoznamka
''ListView1 -> ListViewHistorieStavu
''start_query -> blnZalobaPodana
''
''    Vyhodene controls:
''DTPicker2
''Label2
''Label11
''Label12
''Label13
''IDs
'
'
'    TODO
' povyhadzovat treba nasledujuce pre vsetky zjednotene userforms:
    ' z mainmodule vyhodit z query veci spolu s vyhodenim userformov (cize Update_stav, Update_stavProces, ...)
    ' detto definiciu listview columnheaders vyhodit
    ' detto v MainModule.text_information_hide()
' mozno vsade kde sa opusta modul tak vynuluj nastav objVstupnyUserform na prazdno

Private objVstupnyUserform As Object 'objekt bude userform Klient / Klient_dedictvi / Klient_...  'tady MUSI byt object, ne userform! (jinak objekt nebude mit metodu Show)

'  ostatne premenne
Private strRizeni As String 'jedna z hodnot "epr", "ded", "exe", "ins", "trest", "heir" ' RISK-468
'Private intRizeni As Integer 'pri vkladani upozorneni
' xxxx nasledujuce premenne boli zavedene pre rizeni HEIR, totiz hlavny identifikator v tabulkach uz nie vzdy je id_person,  ale moze to byt aj skp_heir
Private strSkpKey As String ' bud id_person alebo skp_heir
Private strSkpValue As String ' konkretna hodnota pre id_person alebo skp_heir
Private strSkpIdentificator As String ' rozsirenie predosleho o konkretnu hodnotu do podmienok WHERE: bud "id_person = ... " anebo "skp_heir = ... "



Public Property Let VstupnyUserform(ByRef VstupnyUserform As Object) ' tady MUSI byt object, ne userform! (jinak objekt nebude mit metodu Show)
    Set objVstupnyUserform = VstupnyUserform
    
    Select Case objVstupnyUserform.name
        Case "Klient"
        strRizeni = "epr"
        'intRizeni = 3
    Case "Klient_dedictvi"
        strRizeni = "ded"
        'intRizeni = 1
    Case "Klient_exe"
        strRizeni = "exe"
        'intRizeni = 2
    Case "Klient_insol"
        strRizeni = "ins"
        'intRizeni = 4
    Case "Klient_trest"
        strRizeni = "trest"
        'intRizeni = 5
    Case "Klient_heir"
        strRizeni = "heir"
        'intRizeni = 6
    Case Else
        assertion 1 = 0, "jez-45174: neznámé øízení"
    End Select
    
End Property


Private Sub ComboBoxStav_DropButtonClick()
If ComboBoxStav.text = "Žaloba podána" Then 'And strRizeni = "epr" "ded" "ins" "trest"
    Call text_information_hide(Me.name)
    Label30.Visible = True
    Label29.Visible = True
    Label29.Caption = "V dalších informací u klienta bude vyplnìno datum žaloby hodnotou " & Mid(DTPickerPlatnyOd.value, 1, 10)
End If
End Sub

Private Sub DTPickerPlatnyOd_Change()
If ComboBoxStav.text = "Žaloba podána" Then 'And strRizeni = "epr" "ded" "ins" "trest"
    Call text_information_hide(Me.name)
    Label30.Visible = True
    Label29.Visible = True
    Label29.Caption = "V dalších informací u klienta bude vyplnìno datum žaloby hodnotou " & Mid(DTPickerPlatnyOd.value, 1, 10)
End If

End Sub

Private Sub ButtonZavrit_Click()
    Call Me.Hide
End Sub

Private Sub ButtonNavrat_Click()
    Call Me.Hide
End Sub

Private Sub ButtonUlozit_Click()
' najprv sa ukonci platnost vsetkych stavov klienta na danej karte, ktore su prave platne
' potom sa vlozi novy stav
' potom sa zavola extra SUB podla vstupneho rizeni
' nakoniec poznamka o novom stave

Dim sql As String
Dim id As Variant ' integer nebo "NULL"


If strRizeni = "heir" Then
    Debug.Assert Not d Is Nothing ' d je public promenna - objekt Heir
    If d Is Nothing Then
        MsgBox "jezevcik error 008704", vbCritical
        End
    End If
    strSkpKey = "skp_heir"
    strSkpValue = d.lngSkp_heir
Else
    strSkpKey = "id_person"
    strSkpValue = id_person
End If
strSkpIdentificator = strSkpKey & " = " & strSkpValue

If cs_HDWMN.State = 1 Then
prihlasen:
 
If MsgBox("Opravdu chcete pøidat stav na " & ComboBoxStav.text & "?", vbExclamation + vbYesNo, "Zmìna stavu") = vbYes Then 'RISK-676x nova textace
    
    Call TransBegin
        
        '''   updatuje se stav ktery platil do ted
        sql = "update " & schema & "." & strRizeni & "_status_data " & _
                "set dtime_valid_to = (to_date('" & DTPickerPlatnyOd.value & "','dd.mm.yyyy hh24:mi:ss')-1/(24*60*60)), dtime_modification = sysdate, employee = '" & Main.username & _
                "' where flag_deleted = 'N' and " & strSkpIdentificator & _
                " and to_date('" & DTPickerPlatnyOd.value & "','dd.mm.yyyy hh24:mi:ss') between dtime_valid_from and dtime_valid_to"
        Call execute(sql)
        sql = ""
      
        
        '''   vlozi se novy stav
        id = max_skp("select max(skp_" & strRizeni & "_status_data) from " & schema & "." & strRizeni & "_status_data")
        
        
        sql = "insert into " & schema & "." & strRizeni & "_status_data " & _
                "(skp_" & strRizeni & "_status_data, " & strSkpKey & ", skp_" & strRizeni & "_status, dtime_valid_from, dtime_valid_to, text_note, employee, flag_deleted, dtime_modification) " & _
                "select " & id & ", " & strSkpValue & ", skp_" & strRizeni & "_status, to_date('" & DTPickerPlatnyOd.value & "','dd.mm.yyyy hh24:mi:ss'), to_date('1.1.3000','dd.mm.yyyy'), '" & _
                TextBoxPoznamka.text & "', '" & Main.username & "', 'N', sysdate " & _
                "from " & schema & "." & strRizeni & "_status where name_" & strRizeni & "_status = '" & ComboBoxStav.text & "' and dtime_valid_to > sysdate"
        
        Call execute(sql)
        sql = ""
        
        
        ' specialne poziadavky pre rizeni
        If strRizeni = "epr" Then
            Call ExtraEPR
        ElseIf strRizeni = "ins" Then
            Call ExtraINS
        ElseIf strRizeni = "exe" Then
            Call ExtraEXE
        End If
        
        '
        If ComboBoxStav.text = "Vymáhání ukonèeno" And Not strRizeni = "heir" Then
            Call ExtraVymahaniUkonceno(objVstupnyUserform, id_person)
        End If
        
        ' XX nasledujuce moze nastat iba u EPR a EXE a HEIR
        If ComboBoxStav.text = "Žaloba podána" And Not strRizeni = "heir" Then
            Call PodaniZaloby
        End If
        
        
        '''  vlozi sa poznamka
        If Len(Trim(TextBoxPoznamka.text)) < 1 Then
            id = "NULL"
        End If
        
        sql = "insert into " & schema & "." & strRizeni & "_note (skp_" & strRizeni & "_note, " & strSkpKey & ", skp_" & strRizeni & "_note_type, skp_" & strRizeni & "_status_data, text_note, dtime_note_created, employee) " & _
                    "values (" & max_skp("select max(skp_" & strRizeni & "_note) from " & schema & "." & strRizeni & "_note") & ", " & strSkpValue & ", 1, " & id & _
                    ", 'Pøidán nový stav: " & ComboBoxStav.text & ". Platnost od " & DTPickerPlatnyOd.value & "' , sysdate, '" & Main.username & "')"
        Call execute(sql)
        sql = ""
        
    Call TransCommit
    
    'RISK-1114
    'Pokud se jedná o jeden ze stavù navozující vytvoøení složky, tak budu vytváøet složky
    'MM 5.6.2018 zakomentováno - bude se generovat po stistku tlaèítka na exekuèní kartì
    'If ExeFolderToBeCreated(strRizeni, ComboBoxStav.text) Then Call GenerateClientFoldersEXE
    
    Call UserForm_Activate
  
End If ' if msgbox na potvrzeni od zacatku


If err_connection Then
    err_connection = False
    Call re_login(Main.username.text, Main.password.text)
    If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If

Else ' if prihlasen - na zacatku SUB
    Call re_login(Main.username.text, Main.password.text)
    If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If
End Sub

Private Sub ExtraINS()
' specialne poziadavky ktore sa vykonaju iba ak vstupny userform je INS
Dim sql As String

If ComboBoxStav.text = "Pøihláška podána" Then
    sql = "update " & schema & ".ins_info set datum_prihlaseni = to_date('" & DTPickerPlatnyOd.value & "','dd.mm.yyyy hh24:mi:ss') ," & _
        " employee = '" & Main.username & "' , DTIME_MODIFICATION = sysdate  where id_person  =  " & id_person & " "

    Call execute(sql)
End If

End Sub

Private Sub ExtraEXE()
' specialne poziadavky ktore sa vykonaju iba ak vstupny userform je EXE
Dim sql As String
Dim id As Long

' prechod z NR do ER
If ComboBoxStav.text = "Exekuèní návrh" Then
    
    ' ukonci se stav v nalezacim rizeni
    sql = "update " & schema & ".epr_status_data set dtime_valid_to = (to_date('" & DTPickerPlatnyOd.value & "','dd.mm.yyyy hh24:mi:ss')-1/(24*60*60)), dtime_modification = sysdate, employee = '" & Main.username & _
          "' where flag_deleted = 'N' and id_person = " & id_person & _
          " and to_date('" & DTPickerPlatnyOd.value & "','dd.mm.yyyy hh24:mi:ss') between dtime_valid_from and dtime_valid_to"
  
    Call execute(sql)
    sql = ""
    
    
    ' do nalezaciho rizeni se vlozi stav predano do exekucniho rizeni
    id = max_skp("select max(skp_epr_status_data) from " & schema & ".epr_status_data")
    
    sql = "insert into " & schema & ".epr_status_data select " & id & _
          ", " & id_person & ", 81, to_date('" & DTPickerPlatnyOd.value & "','dd.mm.yyyy hh24:mi:ss'), to_date('1.1.3000','dd.mm.yyyy'), '" & TextBoxPoznamka.text & "', '" & Main.username & _
          "', 'N', sysdate from dual " '" & schema & ".exe_status where name_exe_status = '" & ComboBox1.Text & "' and dtime_valid_to > sysdate"
    Call execute(sql)
    sql = ""
    
    ' a jeste poznamka do nalezaciho o zmene stavu
    sql = "insert into " & schema & ".epr_note values (" & max_skp("select max(skp_epr_note) from " & schema & ".epr_note") & ", " & id_person & _
          ", 1, " & id & ", 'Pøidán nový stav: Pøedáno do exekuèního øízení. Platnost od " & DTPickerPlatnyOd.value & "' , sysdate, '" & Main.username & "')"
    Call execute(sql)
    sql = ""

End If


End Sub

Private Sub ExtraEPR()
' specialne poziadavky ktore sa vykonaju iba ak vstupny userform je EPR
' zvlast SUB aby bol prehladnejsi kod
Dim max_skp_notice As Long
Dim sql As String
Dim id As Integer


'''   Vloženie upozornenia pri nepm rozsudku
If (ComboBoxStav.value = "NePM Rozsudek - èásteènì zamítnuto" Or ComboBoxStav.value = "NePM Rozsudek - SK èásteènì zamítnuto" Or ComboBoxStav.value = "NePM Usnesení o schválení smíru" Or ComboBoxStav.value = "NePM Rozsudek" Or ComboBoxStav.value = "NePM Rozsudek - SK" Or ComboBoxStav.value = "Øízení zastaveno" Or ComboBoxStav.value = "Žaloba odmítnuta" Or ComboBoxStav.value = "NePM rozsudek  - žaloba zamítnuta") Then

    max_skp_notice = max_skp("select max(skp_notice) from " & schema & ".notice")
    
    sql = "insert into " & schema & ".notice t (t.skp_notice, t.id_person, t.text_notice, t.dtime_notice_created, t.dtime_notice_deadline, t.code_status, t.employee) " & _
            " values (" & max_skp_notice & " , " & id_person & " , 'Kontrola PM Rozhodnutí' , sysdate , trunc(sysdate + 60) , 'a' , 'sys')"
    
    Call execute(sql)

ElseIf ComboBoxStav.value = "NePM Platební rozkaz" Then 'pokud se jedna o platebni rozkaz, je pozadovano, aby se upozorneni nastavilo az za 90 dni

    max_skp_notice = max_skp("select max(skp_notice) from " & schema & ".notice")
    
    sql = "insert into " & schema & ".notice t (t.skp_notice, t.id_person, t.text_notice, t.dtime_notice_created, t.dtime_notice_deadline, t.code_status, t.employee) " & _
            "values (" & max_skp_notice & " , " & id_person & " , 'Kontrola PM Rozhodnutí' , sysdate , trunc(sysdate + 90) , 'a' , 'sys')"
    
    Call execute(sql)

End If

'''  pri predexucni vyzve to vlozi stav do execniho rizeni

If ComboBoxStav.value = "Pøedexekuèní výzva" Then
    
    sql = "update " & schema & ".exe_status_data set dtime_valid_to = (to_date('" & DTPickerPlatnyOd.value & "','dd.mm.yyyy hh24:mi:ss')-1/(24*60*60)), dtime_modification = sysdate, employee = '" & Main.username & _
            "' where flag_deleted = 'N' and id_person = " & id_person & _
            " and to_date('" & DTPickerPlatnyOd.value & "','dd.mm.yyyy hh24:mi:ss') between dtime_valid_from and dtime_valid_to"

    Call execute(sql)

    sql = "insert into " & schema & ".exe_note values (" & max_skp("select max(skp_exe_note) from " & schema & ".exe_note") & ", " & id_person & _
        ", 1, " & id & ", 'Pøidán nový stav: " & ComboBoxStav.text & ". Platnost od " & DTPickerPlatnyOd.value & "', sysdate, '" & Main.username & "')"
    Call execute(sql)

    id = max_skp("select max(skp_exe_status_data) from " & schema & ".exe_status_data")
    
    'vlozeni stavu do ER
    
    sql = "insert into " & schema & ".exe_status_data select " & id & _
         ", " & id_person & ", 4, to_date('" & DTPickerPlatnyOd.value & "','dd.mm.yyyy hh24:mi:ss'), to_date('1.1.3000','dd.mm.yyyy'), '" & TextBoxPoznamka.text & "', '" & Main.username & _
         "', 'N', sysdate from dual"
    
    Call execute(sql)

End If


End Sub

Private Sub PodaniZaloby()
' SUB vlozi datum podani zaloby do information
Dim sql As String

If prava("select * from " & UCase(schema) & ".information where id_person = " & id_person) Then
    sql = "update " & schema & ".information set date_action = trunc(to_date('" & DTPickerPlatnyOd.value & "','dd.mm.yyyy hh24:mi:ss')), " & _
          "dtime_modification = sysdate, employee = '" & Main.username & "'  where id_person = " & id_person
    Call execute(sql)
Else
    sql = "insert into " & schema & ".information (id_person, date_action, dtime_modification, employee) values (" & _
          id_person & ", trunc(to_date('" & DTPickerPlatnyOd.value & "','dd.mm.yyyy hh24:mi:ss')), sysdate, '" & Main.username & "')"
    Call execute(sql)
End If

sql = "insert into " & schema & ".epr_note values (" & max_skp("select max(skp_epr_note) from " & schema & ".epr_note") & ", " & id_person & _
    ", 2, NULL, 'Zmìna údajù klienta: datum žaloby (" & Mid(DTPickerPlatnyOd.value, 1, 10) & ")' , sysdate, '" & Main.username & "')"
Call execute(sql)

End Sub

Private Sub ButtonSmazatStav_Click()

'If strRizeni = "heir" Then
'    MsgBox "zatím není naimplementováno" 'xxxx tu som zatial nezmenil nic. Bude treba prisposobit metodu Del_stav
'Else
    Del_stav (strRizeni) ' najde posledny stav na klientovi v rizeni a nastavi flag_deleted
'End If
Call UserForm_Activate

End Sub


Private Sub CheckBoxZobrazitSmazane_Click()
    Call UserForm_Activate
End Sub


Private Sub ListViewHistorieStavu_DblClick()

    If ListViewHistorieStavu.ListItems.count > 0 Then
        
        Call text_information_hide(Me.name)
        TextBoxPoznamka.text = ListViewHistorieStavu.SelectedItem.SubItems(6)
    
    End If

End Sub

Private Sub TextBoxPoznamka_Change()
    Label32.Caption = 200 - Len(TextBoxPoznamka.text)
End Sub


' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()
Dim sql As String

If Len(strRizeni) < 2 Then MsgBox "Jezevèík dostal pásomnici. ((je to bug))" 'akoze kontrola jestli bylo zadano rizeni

If cs_HDWMN.State = 1 Then
prihlasen:

Call text_information_hide(Me.name)



TextBoxPoznamka.text = ""

Call NaplnListViewHistorieStavu ' podle rizeni a klienta
Call NaplnComboBoxStav ' podle rizeni


'''  vyhozeno 5.6.2018, jez.v.5.7., RISK-1176
'''  minimalny dtime nech je o jedna vacsi nez posledny stav
'''If ListViewHistorieStavu.ListItems.count <= 0 Then
'''    DTPickerPlatnyOd.MinDate = format("23.2.1989 06:30:55", "dd.MM.yyyy HH:mm:ss") 'XX Bulhar = datum narozeni Lucky Kasparovej ako pamiatka na to, ze odhalila tento bug. Yeeaah.
'''Else
'''    For i = 1 To ListViewHistorieStavu.ListItems.count
'''        If ListViewHistorieStavu.ListItems(i).SubItems(3) = "Aktivni" Then 'you could also use the LIKE operator
'''            DTPickerPlatnyOd.MinDate = DateAdd("s", 1, ListViewHistorieStavu.ListItems(i).SubItems(4)) 'k datumu zacatku platnosti poslendiho stavu priratame 1 sekundu
'''            Exit For
'''        End If
'''    Next
'''End If
DTPickerPlatnyOd.CustomFormat = "dd.MM.yyyy HH:mm:ss" 'DTPickerPlatnyOd.format = dtpCustom
DTPickerPlatnyOd.MinDate = Now
DTPickerPlatnyOd.MaxDate = DateAdd("s", 10, Now) 'bulharskych 10 sekund naviac
DTPickerPlatnyOd.value = Now


If ListViewHistorieStavu.ListItems.count > 0 Then
    Label27.Visible = True
    Label29.Visible = True
    Label29.Caption = "Nalezeno " & ListViewHistorieStavu.ListItems.count & " záznamù"
Else
    Label28.Visible = True
    Label29.Visible = True
    Label29.Caption = "Nalezeno " & ListViewHistorieStavu.ListItems.count & " záznamù"
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

Private Sub NaplnListViewHistorieStavu()
' tato SUB zmaze a naplni historiu stavov na klientovi do listview
' (povodne, este ked boli 5 Userformov, sa to robilo v MainModule vo funkcii query(), resp. query1() )
' najprv sa nastavi sql podla rizeni
' potom sa cez loop naplni listview

Dim sql As String
Dim itmX As ListItem

ListViewHistorieStavu.ListItems.Clear

If strRizeni = "heir" Then
    sql = "select * from " & schema & ".v_" & strRizeni & "_status where skp_heir = " & d.lngSkp_heir
Else
    sql = "select * from " & schema & ".v_" & strRizeni & "_status where id_person = " & id_person
End If

'''   nastavi se i smazane stavy
If Not CheckboxZobrazitSmazane.value Then
    sql = sql & " and flag_deleted is null "
End If
'''   ORDER BY
sql = sql & " order by dtime_valid_from desc"


'''   NAPLNIME LISTVIEW
Set rst1 = New ADODB.Recordset
rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

    If Not rst1.EOF Then
        rst1.MoveFirst
        For i = 0 To rst1.RecordCount - 1
            Set itmX = ListViewHistorieStavu.ListItems.Add()
            itmX.text = "" & rst1(1) & ""
            itmX.SubItems(1) = "" & rst1(2) & ""
            itmX.SubItems(2) = "" & rst1(10) & ""
            itmX.SubItems(3) = "" & rst1(5) & ""
            itmX.SubItems(4) = "" & rst1(3) & ""
            itmX.SubItems(5) = "" & rst1(4) & ""
            itmX.SubItems(6) = "" & rst1(9) & ""
            itmX.SubItems(7) = "" & rst1(7) & ""
            itmX.SubItems(8) = "" & rst1(8) & ""
            itmX.SubItems(9) = "" & rst1(6) & ""
            rst1.MoveNext
        Next i
    End If
rst1.Close

Call ZabarviRadkyTextove(ListViewHistorieStavu, 9, "Smazano")


End Sub


Private Sub NaplnComboBoxStav()
' tato SUB zmaze a naplni zoznam stavov do ComboBoxStav
' (povodne, este ked boli 5 Userformov, sa to robilo v MainModule vo funkcii query(), resp. query1() )
' zoznam sa meni podla rizeni, takze sa najprv vytvori sql dotaz podla rizeni
' potom sa cez loop naplnia hodnoty do rolovatka a nastavi sa hodnota prvej z nich

Dim sql As String

ComboBoxStav.Clear

' vytvorime sql dotaz podle rizeni.
Select Case strRizeni

    Case "epr"
        
        sql = "select s.name_epr_status as name_status from " & schema & ".epr_status s where s.dtime_valid_to > sysdate " & _
                    "ORDER BY CASE WHEN lower(s.name_epr_status) < LOWER('PM Rozsudek') THEN 1 " & _
                            "WHEN LOWER(s.name_epr_status) = LOWER('PM Rozsudek') THEN 2 " & _
                            "WHEN LOWER(s.name_epr_status) = LOWER('PM Rozsudek s doložkou') THEN 3 " & _
                            "WHEN (LOWER(s.name_epr_status) > LOWER('PM Rozsudek')  AND  LOWER(s.name_epr_status) < LOWER('PM Rozsudek - SK')) THEN 4 " & _
                            "WHEN LOWER(s.name_epr_status) = LOWER('PM Rozsudek - SK') THEN 5 " & _
                            "WHEN LOWER(s.name_epr_status) = LOWER('PM rozsudek – SK s doložkou') THEN 6 ELSE 99 End, s.name_epr_status"
        'old: 'sql = "select s.name_epr_status from " & schema & ".epr_status s where s.dtime_valid_to > sysdate order by order_item nulls last"
        'old: 'sql = "select s.name_epr_status from " & schema & ".epr_status s where s.dtime_valid_to > sysdate order by s.name_epr_status"

    Case "ded"
        sql = "select s.name_ded_status as name_status from " & schema & ".ded_status s where s.dtime_valid_to > sysdate order by s.name_ded_status nulls last "
    Case "exe"
        sql = "select s.name_exe_status as name_status from " & schema & ".exe_status s where s.dtime_valid_to > sysdate order by s.name_exe_status nulls last"
    Case "ins"
        sql = "select s.name_ins_status as name_status from " & schema & ".ins_status s where s.dtime_valid_to > sysdate order by s.order_item nulls last "
    Case "trest"
        sql = "select s.name_trest_status as name_status from " & schema & ".trest_status s where s.dtime_valid_to > sysdate order by s.name_trest_status nulls last "
    Case "heir"
        sql = "select s.name_heir_status as name_status FROM " & schema & ".heir_status s where s.dtime_valid_to > sysdate order by s.ORDER_ITEM, s.name_heir_status nulls last "
    Case Else
        MsgBox "Urval si jezevèíkovi nohu."

End Select


' dedictvi v mainmodule ma nejaky kod naviac tam -- search: Case "Update_stav_ded"

Set rst1 = New ADODB.Recordset
rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst1.EOF Then
        rst1.MoveFirst
        For i = 0 To rst1.RecordCount - 1
            ComboBoxStav.AddItem "" & rst1![name_status] & ""
            rst1.MoveNext
        Next i
    End If
rst1.Close
sql = ""


If ComboBoxStav.ListCount > 0 Then
   ComboBoxStav.ListIndex = 0
End If


End Sub

Private Sub UserForm_Initialize()
Call table_prepare(Me.name)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, _
                            CloseMode As Integer)
Dim ans
If CloseMode = vbFormControlMenu Then
    Cancel = True
    Call Me.Hide
End If
End Sub


Private Sub ExtraVymahaniUkonceno(ByRef VstupnyUserform As Object, ByVal id_person As Variant)
' RISK-419  -  pokud se v nekterem rizeni zada stav vymahani ukonceno, chceme zobrazit poznamku na vsech kartach klienta
'do kolekce strRizeniVloz dostaneme vsechny rizeni, kam se ma poznamka vlozit (tedy krome jednoho)
'pak loop pres rizeni ktery tu poznamku vlozi
' ZATIM SE IGNORUJE HEIR


Dim strRizeniVloz As Collection
Dim r As Variant
Dim dedic As Heir
'Dim initial_client As Client
Dim strSkpIdentify As String
Dim strRizeni As String
Dim strSkpKey As String   ' tady vzdy SKP_HEIR
Dim strSkpValue As String
Dim id As Long

Set strRizeniVloz = New Collection
Dim sql As String


Debug.Assert VstupnyUserform.name <> "heir"

With strRizeniVloz
    .Add "epr", "epr"
    .Add "ded", "ded"
    .Add "exe", "exe"
    .Add "ins", "ins"
    .Add "trest", "trest"
End With


Select Case VstupnyUserform.name
    Case "Klient"
        strRizeniVloz.remove "epr"
    Case "Klient_dedictvi"
        strRizeniVloz.remove "ded"
    Case "Klient_exe"
        strRizeniVloz.remove "exe"
    Case "Klient_insol"
        strRizeniVloz.remove "ins"
    Case "Klient_trest"
        strRizeniVloz.remove "trest"
    Case Else
        MsgBox "BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG " & _
                "BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG " & _
                "BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG "
End Select



For Each r In strRizeniVloz
    sql = "insert into " & schema & "." & r & "_note n (N.SKP_" & r & "_NOTE, N.ID_PERSON, N.SKP_" & r & "_NOTE_TYPE, N.SKP_" & r & "_STATUS_DATA, N.TEXT_NOTE, N.DTIME_NOTE_CREATED, N.EMPLOYEE) " & _
                " values (" & max_skp("select max(skp_" & r & "_note) from " & schema & "." & r & "_note") & ", " & id_person & _
                ", 3, NULL, 'Vymáhání ukonèeno', sysdate, '" & Main.username & "')"

    Call execute(sql)
Next r


'''   RISK-1176  Propisování stavu Vymáhání ukonceno na karty dedicu

If VstupnyUserform.name = "Klient_dedictvi" Then
    
    strRizeni = "HEIR"
    strSkpKey = "SKP_HEIR"
    Set dedic = New Heir
    
    'Transakce už je v nadøazeném modulu

    For Each dedic In dedici.colHeirs
                
        strSkpValue = dedic.lngSkp_heir
        strSkpIdentify = strSkpKey & " = " & strSkpValue
        
        '''   updatuje se stav ktery platil do ted
        
        sql = "update " & schema & "." & strRizeni & "_status_data " & _
                "set dtime_valid_to = (to_date('" & DTPickerPlatnyOd.value & "','dd.mm.yyyy hh24:mi:ss')-1/(24*60*60)), dtime_modification = sysdate, employee = '" & Main.username & _
                "' where flag_deleted = 'N' and " & strSkpIdentify & _
                " and to_date('" & DTPickerPlatnyOd.value & "','dd.mm.yyyy hh24:mi:ss') between dtime_valid_from and dtime_valid_to"
        Call execute(sql)
        sql = ""
        
        
        '''   vlozi se novy stav
        
        id = max_skp("select max(skp_" & strRizeni & "_status_data) from " & schema & "." & strRizeni & "_status_data")
        
        sql = "insert into " & schema & "." & strRizeni & "_status_data " & _
                "(skp_" & strRizeni & "_status_data, " & strSkpKey & ", skp_" & strRizeni & "_status, dtime_valid_from, dtime_valid_to, text_note, employee, flag_deleted, dtime_modification) " & _
                "select " & id & ", " & strSkpValue & ", skp_" & strRizeni & "_status, to_date('" & DTPickerPlatnyOd.value & "','dd.mm.yyyy hh24:mi:ss'), to_date('1.1.3000','dd.mm.yyyy'), '" & _
                TextBoxPoznamka.text & "', '" & Main.username & "', 'N', sysdate " & _
                "from " & schema & "." & strRizeni & "_status where name_" & strRizeni & "_status = '" & ComboBoxStav.text & "' and dtime_valid_to > sysdate"
        
        Call execute(sql)
        sql = ""
        
        
        '''  vlozi se poznamka
        
        sql = "insert into " & schema & "." & strRizeni & "_note (skp_" & strRizeni & "_note, " & strSkpKey & ", skp_" & strRizeni & "_note_type, skp_" & strRizeni & "_status_data, text_note, dtime_note_created, employee) " & _
                    "values (" & max_skp("select max(skp_" & strRizeni & "_note) from " & schema & "." & strRizeni & "_note") & ", " & strSkpValue & ", 1, NULL" & _
                    ", 'Pøidán nový stav: " & ComboBoxStav.text & ". Platnost od " & DTPickerPlatnyOd.value & "' , sysdate, '" & Main.username & "')"
        Call execute(sql)
        sql = ""
        
        
    Next dedic




End If




End Sub


Private Sub VymazPoznamkuVymahaniUkonceno(ByRef VstupnyUserform As Object, ByVal id_person As Variant)
' pokud se nekde zmaze stav Vymahani Ukonceno, je potreba smazat poznamku na vsech kartach klienta
' ((pricinou je RISK-419  -  pokud se v nekterem rizeni zada stav vymahani ukonceno, zobrazi se poznamka na vsech kartach klienta))
'do kolekce strRizeniVymaz dostaneme vsechny rizeni, odkud se ma poznamka smazat
'pak loop pres rizeni ktery tu poznamku smaze
' ZATIM SE IGNORUJE HEIR (e.g. RISK-1176)

Dim strRizeniVymaz As Collection
Dim r As Variant

Set strRizeniVymaz = New Collection
Dim sql As String

Debug.Assert Not VstupnyUserform.name = "heir"

With strRizeniVymaz
    .Add "epr", "epr"
    .Add "ded", "ded"
    .Add "exe", "exe"
    .Add "ins", "ins"
    .Add "trest", "trest"
End With


Select Case VstupnyUserform.name
    Case "Klient"
        strRizeniVymaz.remove "epr"
    Case "Klient_dedictvi"
        strRizeniVymaz.remove "ded"
    Case "Klient_exe"
        strRizeniVymaz.remove "exe"
    Case "Klient_insol"
        strRizeniVymaz.remove "ins"
    Case "Klient_trest"
        strRizeniVymaz.remove "trest"
    Case Else
        MsgBox "BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG " & _
                "BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG " & _
                "BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG BUG "
End Select



For Each r In strRizeniVymaz
    
    ' najdi cas zadania stavu povodneho
    ' najdi na ostatnych kartach ID poznamok, daj tam podmienku na cas nech to je v rozmedzi par hodin (stacilo by aj minut..)
    ' musi najst 0 alebo 1 IDcok ku kazdemu rizeni, ak niekde viac tak msgbox
    ' debug.assert ze tych IDcok na zmazanie spolu nebude viac ako 5
    ' vymazat poznamky ktore maju dane ID..
    ' pak jeste extra pripad zemreli a dedicove
    
    sql = "insert into " & schema & "." & r & "_note n (N.SKP_" & r & "_NOTE, N.ID_PERSON, N.SKP_" & r & "_NOTE_TYPE, N.SKP_" & r & "_STATUS_DATA, N.TEXT_NOTE, N.DTIME_NOTE_CREATED, N.EMPLOYEE) " & _
                " values (" & max_skp("select max(skp_" & r & "_note) from " & schema & "." & r & "_note") & ", " & id_person & _
                ", 3, NULL, 'Vymáhání ukonèeno', sysdate, '" & Main.username & "')"

    Call execute(sql)
    
Next r

End Sub



' Zmìna barvy po najetí myši **************************************************************************************************************************
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
ButtonZavrit.ForeColor = barvaCerna
ButtonUlozit.ForeColor = barvaCerna

End Sub

Private Sub ButtonZavrit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
ButtonZavrit.ForeColor = barvaZelenaAirBank
End Sub

Private Sub ButtonUlozit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
ButtonUlozit.ForeColor = barvaZelenaAirBank
End Sub

Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
ButtonNavrat.ForeColor = barvaCerna
End Sub

Private Sub ButtonNavrat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
ButtonNavrat.ForeColor = barvaZelenaAirBank
End Sub





