VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Zadani_Sance 
   Caption         =   "Zadání Šance"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13785
   OleObjectBlob   =   "Zadani_Sance.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Zadani_Sance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit

'RISK-778
Private objVstupnyUserform As Object 'objekt bude userform Klient / Klient_dedictvi / Klient_...  'tady MUSI byt object, ne userform! (jinak objekt nebude mit metodu Show)
'  ostatne premenne
Private strRizeni As String 'jedna z hodnot "epr", "ded", "exe", "ins", "trest", "heir"

Private SKP_CHANCE_IMPORT As Long, SKP_PROCESS As Long  'Pouziji pak pro delete a update a také rozhodování o tom, jak má vypadat ta stránka

Public Property Let VstupnyUserform(ByRef VstupnyUserform As Object) ' tady MUSI byt object, ne userform! (jinak objekt nebude mit metodu Show)
    Set objVstupnyUserform = VstupnyUserform
       
        Select Case objVstupnyUserform.name
        Case "Klient"
            strRizeni = "epr"
        Case "Klient_dedictvi"
            strRizeni = "ded"
        Case "Klient_heir" 'RISK-468
            strRizeni = "heir"
        Case "Klient_exe"
            strRizeni = "exe"
        Case "Klient_insol"
            strRizeni = "ins"
        Case "Klient_trest"
            strRizeni = "trest"
    End Select
       
End Property


Private Sub ComboBox_Chance_Type_Change()
'Kontrola, pokud už bude nìjaký platný záznam šance daného typu, tak nedovolím zmìnit na dané šanci datum udìlení a uznání dluhu
'RISK-919
Dim i As Integer
MonthView_Chance_Given.Enabled = True
ComboBox_Uznani_Dluhu.Enabled = True

'Chci to zablokovat pouze v pøípadì, kdy pøidává novou, ne když dìlá update
If SKP_CHANCE_IMPORT = 0 Then

For i = 1 To Zadani_Sance.ListView_Chance_Overview.ListItems.count

    If ListvZiskejHodnotu(Zadani_Sance.ListView_Chance_Overview, "Chance_Type", i) = ComboBox_Chance_Type.value Then
        MonthView_Chance_Given = ListvZiskejHodnotu(Zadani_Sance.ListView_Chance_Overview, "Chance_Date", i)
        MonthView_Chance_Given.Enabled = False
        ComboBox_Uznani_Dluhu.value = ListvZiskejHodnotu(Zadani_Sance.ListView_Chance_Overview, "Uznani_Dluhu", i)
        ComboBox_Uznani_Dluhu.Enabled = False
        Exit For
    End If
        
Next i

End If

End Sub



Private Sub CommandButton_Add_Chance_Click()

MonthView_Chance_Given = date
MonthView_Instalment = date

'Vyèistím
SKP_CHANCE_IMPORT = 0

Zadani_Sance.Height = 383.25
'A oddìlám ostatní tlaèítka
CommandButton_Change_Chance.Enabled = False
CommandButton_Delete_Chance.Enabled = False

Call ComboBox_Chance_Type_Change 'Volám proto, aby se mi pøedvyplnila hodnota do comboboxu uznání dluhu z hodnot které už znám. Zároveò to zablokuje uživateli možnost hodnotu zmìnit.

End Sub

Private Sub CommandButton_Delete_Chance_Click()
Dim sql As String

If Zadani_Sance.ListView_Chance_Overview.ListItems.count = 0 Then
    MsgBox "Na klientovi není žádná šance evidována, nelze nic mazat", vbExclamation, "Klient bez šance"
    Exit Sub
End If

If MsgBox("Opravdu chcete smazat oznaèenou šanci?", vbExclamation + vbYesNo, "Smazání šance") = vbYes Then

    SKP_CHANCE_IMPORT = ListvZiskejHodnotu(Zadani_Sance.ListView_Chance_Overview, "SKP_CHANCE_IMPORT")
    sql = "update " & schema & ".chance_import" & _
          " set flag_deleted = 'Y'" & _
          "    ,employee_modified = '" & UCase(Main.username) & "'" & _
          "    ,dtime_modified = sysdate" & _
          " where skp_chance_import = " & SKP_CHANCE_IMPORT

    execute (sql)
    
    MsgBox "Šance byla úspìšnì smazána"
    Call UserForm_Activate

End If

'Vyèistím
SKP_CHANCE_IMPORT = 0

End Sub

Private Sub CommandButton_Change_Chance_Click()
Dim sql As String

If Zadani_Sance.ListView_Chance_Overview.ListItems.count = 0 Then
    MsgBox "Na klientovi není žádná šance evidována, nelze nic mìnit", vbExclamation, "Klient bez šance"
    Exit Sub
End If

Zadani_Sance.Height = 383.25
'A oddìlám ostatní tlaèítka
CommandButton_Add_Chance.Enabled = False
CommandButton_Delete_Chance.Enabled = False

'Pomocí tento promìnné potom zmìním, když dá uložit
SKP_CHANCE_IMPORT = ListvZiskejHodnotu(Zadani_Sance.ListView_Chance_Overview, "SKP_CHANCE_IMPORT")

'Inicializace boxíkù
ComboBox_Chance_Type.value = ListvZiskejHodnotu(Zadani_Sance.ListView_Chance_Overview, "Chance_Type")
ComboBox_Instalment_Type.value = ListvZiskejHodnotu(Zadani_Sance.ListView_Chance_Overview, "Inst_Type")
'RISK-919
ComboBox_Uznani_Dluhu.value = ListvZiskejHodnotu(Zadani_Sance.ListView_Chance_Overview, "Uznani_Dluhu")
TextBox_Amt_Inst.value = ListvZiskejHodnotu(Zadani_Sance.ListView_Chance_Overview, "Inst_Amt")

TextBox_Note.value = ListvZiskejHodnotu(Zadani_Sance.ListView_Chance_Overview, "Note")

MonthView_Chance_Given.value = ListvZiskejHodnotu(Zadani_Sance.ListView_Chance_Overview, "Chance_Date")
MonthView_Instalment.value = ListvZiskejHodnotu(Zadani_Sance.ListView_Chance_Overview, "Inst_Date")

End Sub

Private Sub CommandButton_Zrus_Click()
   'Vyèistím
   SKP_CHANCE_IMPORT = 0


ComboBox_Chance_Type.value = vbNullString
ComboBox_Instalment_Type.value = vbNullString
ComboBox_Uznani_Dluhu.value = vbNullString 'RISK-919
TextBox_Amt_Inst.value = vbNullString
TextBox_Note.value = vbNullString

CommandButton_Add_Chance.Enabled = True
CommandButton_Change_Chance.Enabled = True
CommandButton_Delete_Chance.Enabled = True

Zadani_Sance.Height = 210.75
End Sub

Private Sub CommandButton_Uloz_Click()
Dim sql As String, sql_chance As String
Dim Amt_Instalment As String
Dim Poznamka As String
Dim SKP_CHANCE As Long, Bln_New_Chance As Boolean 'Používá se jako nová hodnota se sekvence, pokud klient ještì nemá záznam stejného typu šance, jinak se vybere hodnota, co už tam má
Dim Str_Uznani_Dluhu As String 'RISK-919


Bln_New_Chance = True
   
If TextBox_Amt_Inst.value = vbNullString Or Not IsNumeric(TextBox_Amt_Inst.value) Or TextBox_Amt_Inst.value < 0 Then
    MsgBox "Do výše splátky nebyla zadána èíselná hodnota, zkus to znovu", vbExclamation, "Chybné zadání hodnot"
    Exit Sub
End If
'RISK-919
If ComboBox_Uznani_Dluhu.value = vbNullString Then
    MsgBox "Nebylo zadáno uznání dluhu, zkus to znovu", vbExclamation, "Chybné zadání hodnot"
    Exit Sub
ElseIf ComboBox_Uznani_Dluhu.value = "Ano" Then
    Str_Uznani_Dluhu = "Y"
Else: Str_Uznani_Dluhu = "N"
    
End If

If ComboBox_Chance_Type.value = vbNullString Then
    MsgBox "Nebyl zadán typ šance, zkus to znovu", vbExclamation, "Chybné zadání hodnot"
    Exit Sub
End If

If ComboBox_Instalment_Type.value = vbNullString Then
    MsgBox "Nebyl zadán typ splátky, zkus to znovu", vbExclamation, "Chybné zadání hodnot"
    Exit Sub
End If


    Amt_Instalment = Replace(TextBox_Amt_Inst.value, ",", ".")
   
   If SKP_CHANCE_IMPORT >= 1 Then 'Update
        
        'RISK-919
        sql = "update " & schema & ".chance_import" & _
              " set typ_sk = '" & ComboBox_Chance_Type.value & "'" & _
              "    ,date_sk = to_date('" & MonthView_Chance_Given.value & "','dd.mm.yyyy')" & _
              "    ,date_instalment = to_date('" & MonthView_Instalment.value & "','dd.mm.yyyy')" & _
              "    ,amt_instalment = " & Amt_Instalment & _
              "    ,type_instalment = '" & ComboBox_Instalment_Type.value & "'" & _
              "    ,employee_modified = '" & UCase(Main.username) & "'" & _
              "    ,dtime_modified = sysdate" & _
              "    ,note = '" & TextBox_Note.value & "'" & _
              "    ,flag_uznani_dluhu = '" & Str_Uznani_Dluhu & "'" & _
              " where skp_chance_import = " & SKP_CHANCE_IMPORT
       
        execute (sql)
        MsgBox "Šance byla úspìšnì upravena"
        
   Else: 'Insert
        
        
        'Podívám se, jestli už klient má záznam stejného typu šance a podle toho pøiøadím buï nové, nebo stávájící SKP_CHANCE
        'Nemùžu se podívat pouze do listview, protože chci i smazané šance
        sql_chance = "select distinct skp_chance, typ_sk from " & schema & ".chance_import where id_person = " & c.id_person
        
        Set rst1 = New ADODB.Recordset
           rst1.Open sql_chance, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
           If Not rst1.EOF Then
            rst1.MoveFirst
            For i = 1 To rst1.RecordCount
              
              If ComboBox_Chance_Type.value = rst1![typ_sk] Then
                SKP_CHANCE = rst1![SKP_CHANCE]
                Bln_New_Chance = False
              End If
            rst1.MoveNext
            Next i
           End If
        
        rst1.Close
        
        
        'Uložit do recordsetu, pak projet pøes daný typ šance a podle toho urèit promìnnou Bln_New_Chance
        
        If ComboBox_Chance_Type.value = "dìdická šance" Then 'oprava - objekt d.lngSkp_heir není nadefinován u normálních šancí a proto to spadne
        
               'RISK-919 RISK-468 RISK-1068
                If Bln_New_Chance Then 'Pokud zadává novou šanci tak hodnota ze sekvenci
                    sql = "insert into " & schema & ".chance_import (skp_chance_import, skp_chance ,id_person, typ_sk, date_sk, date_instalment, amt_instalment, type_instalment, employee_modified, dtime_modified, note, flag_deleted, skp_process, flag_uznani_dluhu, skp_heir, dtime_inserted)" & _
                          " values " & _
                          " (" & schema & ".chance_import_seq.nextval," & schema & ".chance_seq.nextval, " & c.id_person & ",'" & ComboBox_Chance_Type.value & "', to_date('" & MonthView_Chance_Given.value & "','dd.mm.yyyy'), to_date('" & MonthView_Instalment.value & "','dd.mm.yyyy'), " & Amt_Instalment & ",'" & ComboBox_Instalment_Type.value & "', '" & UCase(Main.username) & "', sysdate,'" & TextBox_Note.value & "','N'," & SKP_PROCESS & ",'" & Str_Uznani_Dluhu & "'," & d.lngSkp_heir & ", sysdate)"
                    
                Else: 'Pokud jen doplòuje k šanci, tak už existující
                    sql = "insert into " & schema & ".chance_import (skp_chance_import, skp_chance ,id_person, typ_sk, date_sk, date_instalment, amt_instalment, type_instalment, employee_modified, dtime_modified, note, flag_deleted, skp_process, flag_uznani_dluhu, skp_heir, dtime_inserted)" & _
                          " values " & _
                          " (" & schema & ".chance_import_seq.nextval," & SKP_CHANCE & ", " & c.id_person & ",'" & ComboBox_Chance_Type.value & "', to_date('" & MonthView_Chance_Given.value & "','dd.mm.yyyy'), to_date('" & MonthView_Instalment.value & "','dd.mm.yyyy'), " & Amt_Instalment & ",'" & ComboBox_Instalment_Type.value & "', '" & UCase(Main.username) & "', sysdate,'" & TextBox_Note.value & "','N'," & SKP_PROCESS & ",'" & Str_Uznani_Dluhu & "'," & d.lngSkp_heir & ", sysdate)"
                    
                
                End If
        
        Else 'standardní šance - ne dìdická
        
                
                If Bln_New_Chance Then 'Pokud zadává novou šanci tak hodnota ze sekvenci
                    sql = "insert into " & schema & ".chance_import (skp_chance_import, skp_chance ,id_person, typ_sk, date_sk, date_instalment, amt_instalment, type_instalment, employee_modified, dtime_modified, note, flag_deleted, skp_process, flag_uznani_dluhu)" & _
                          " values " & _
                          " (" & schema & ".chance_import_seq.nextval," & schema & ".chance_seq.nextval, " & c.id_person & ",'" & ComboBox_Chance_Type.value & "', to_date('" & MonthView_Chance_Given.value & "','dd.mm.yyyy'), to_date('" & MonthView_Instalment.value & "','dd.mm.yyyy'), " & Amt_Instalment & ",'" & ComboBox_Instalment_Type.value & "', '" & UCase(Main.username) & "', sysdate,'" & TextBox_Note.value & "','N'," & SKP_PROCESS & ",'" & Str_Uznani_Dluhu & "')"
                    
                Else: 'Pokud jen doplòuje k šanci, tak už existující
                    sql = "insert into " & schema & ".chance_import (skp_chance_import, skp_chance ,id_person, typ_sk, date_sk, date_instalment, amt_instalment, type_instalment, employee_modified, dtime_modified, note, flag_deleted, skp_process, flag_uznani_dluhu)" & _
                          " values " & _
                          " (" & schema & ".chance_import_seq.nextval," & SKP_CHANCE & ", " & c.id_person & ",'" & ComboBox_Chance_Type.value & "', to_date('" & MonthView_Chance_Given.value & "','dd.mm.yyyy'), to_date('" & MonthView_Instalment.value & "','dd.mm.yyyy'), " & Amt_Instalment & ",'" & ComboBox_Instalment_Type.value & "', '" & UCase(Main.username) & "', sysdate,'" & TextBox_Note.value & "','N'," & SKP_PROCESS & ",'" & Str_Uznani_Dluhu & "')"
                    
                
                End If
        
        
        End If
            
        execute (sql)
        MsgBox "Šance byla úspìšnì vložena"
   End If
   
   
   
   'Vyèistím
   SKP_CHANCE_IMPORT = 0
   SKP_CHANCE = 0
ComboBox_Chance_Type.value = vbNullString
ComboBox_Instalment_Type.value = vbNullString
ComboBox_Uznani_Dluhu.value = vbNullString
TextBox_Amt_Inst.value = vbNullString
TextBox_Note.value = vbNullString
Zadani_Sance.Height = 210.75

CommandButton_Add_Chance.Enabled = True
CommandButton_Change_Chance.Enabled = True
CommandButton_Delete_Chance.Enabled = True

'Pøi vložení nové šance se zároveò propíše nový stav na nalézací (pøípadnì na dìdické) øízení.
'Nenastává u šance typu SK urèeno soudem.
'!! Plní se databázovým triggerem STATUS_AFTER_INSERT_CHANCE
  

Call UserForm_Activate
 
End Sub



Private Sub MonthView_Instalment_DateClick(ByVal DateClicked As Date)
'Kontrola, pokud je datum splátky menší než datum zadání šance, tak vyhodím varování...
If MonthView_Instalment.value < MonthView_Chance_Given.value Then
    MsgBox "Datum splátky pøedchází datu zadání šance, zkontroluj si to, prosím. ", vbExclamation, "Varování"
End If

End Sub

Private Sub MonthView_Chance_Given_DateClick(ByVal DateClicked As Date)
'Kontrola, pokud je datum splátky menší než datum zadání šance, tak vyhodím varování...
If MonthView_Instalment.value < MonthView_Chance_Given.value Then
    MsgBox "Datum splátky pøedchází datu zadání šance, zkontroluj si to, prosím. ", vbExclamation, "Varování"
End If

End Sub

Private Sub TextBox_Note_Change()
Label_Note_Zbyva.Caption = 300 - Len(TextBox_Note.text)
End Sub

' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()
Dim keys As Dictionary
Dim clmx As ColumnHeader
Dim itmX As ListItem
Dim sql As String

MonthView_Chance_Given.Enabled = True
CommandButton_Add_Chance.Enabled = True
CommandButton_Change_Chance.Enabled = True
CommandButton_Delete_Chance.Enabled = True
Zadani_Sance.Height = 210.75

Zadani_Sance.TextBox_Amt_Inst.ControlTipText = "Pokud chceš i haléøe, tak zadej výši ve formátu 1234,56"

'Naètení listview se zadanými šancemi

If Zadani_Sance.ListView_Chance_Overview.ColumnHeaders.count = 0 Then

 Set clmx = Zadani_Sance.ListView_Chance_Overview.ColumnHeaders.Add(, , "SKP_CHANCE_IMPORT", 0)
 Set clmx = Zadani_Sance.ListView_Chance_Overview.ColumnHeaders.Add(, "SKP_CHANCE", "ID Šance", 30)
 Set clmx = Zadani_Sance.ListView_Chance_Overview.ColumnHeaders.Add(, "Chance_Type", "Typ šance", 100)
 Set clmx = Zadani_Sance.ListView_Chance_Overview.ColumnHeaders.Add(, "Chance_Date", "Datum udìlení šance/ nabytí PM", 70)
 Set clmx = Zadani_Sance.ListView_Chance_Overview.ColumnHeaders.Add(, "Inst_Date", "Datum splátky", 60)
 Set clmx = Zadani_Sance.ListView_Chance_Overview.ColumnHeaders.Add(, "Inst_Amt", "Výše splátky", 60)
 Set clmx = Zadani_Sance.ListView_Chance_Overview.ColumnHeaders.Add(, "Inst_Type", "Druh splátky", 80)
 Set clmx = Zadani_Sance.ListView_Chance_Overview.ColumnHeaders.Add(, "Uznani_Dluhu", "Uznání dluhu", 80) 'RISK-919
 Set clmx = Zadani_Sance.ListView_Chance_Overview.ColumnHeaders.Add(, "Employee", "Kdo šanci zadal", 80)
 Set clmx = Zadani_Sance.ListView_Chance_Overview.ColumnHeaders.Add(, "Dtime_Modified", "Datum zadání", 70)
 Set clmx = Zadani_Sance.ListView_Chance_Overview.ColumnHeaders.Add(, "Note", "Poznámka", 150)
 Set clmx = Zadani_Sance.ListView_Chance_Overview.ColumnHeaders.Add(, , "SKP_HEIR", 0) 'RISK-468
 
 Call NastavKliceListview(Zadani_Sance.ListView_Chance_Overview)

End If
'RISK-919
sql = "select skp_chance_import, skp_chance ,typ_sk, date_sk, date_instalment, amt_instalment, type_instalment, employee_modified, dtime_modified, note, decode(flag_uznani_dluhu,'Y','Ano','N','Ne',flag_uznani_dluhu) as flag_uznani_dluhu, skp_heir from " & schema & ".chance_import where flag_deleted = 'N' and id_person = " & c.id_person & " order by typ_sk, date_instalment"


Set keys = KliceKListview(Zadani_Sance.ListView_Chance_Overview)
   
  Zadani_Sance.ListView_Chance_Overview.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
        rst1.MoveFirst
        For i = 1 To rst1.RecordCount
            Set itmX = Zadani_Sance.ListView_Chance_Overview.ListItems.Add()
            itmX.text = "" & rst1![SKP_CHANCE_IMPORT] & ""
            itmX.SubItems(keys("Chance_Type")) = "" & rst1![typ_sk] & ""
            itmX.SubItems(keys("Chance_Date")) = "" & rst1![date_sk] & ""
            itmX.SubItems(keys("Inst_Date")) = "" & rst1![date_instalment] & ""
            itmX.SubItems(keys("Inst_Amt")) = "" & rst1![Amt_Instalment] & ""
            itmX.SubItems(keys("Inst_Type")) = "" & rst1![type_instalment] & ""
            itmX.SubItems(keys("Employee")) = "" & rst1![employee_modified] & ""
            itmX.SubItems(keys("Dtime_Modified")) = "" & rst1![dtime_modified] & ""
            itmX.SubItems(keys("Note")) = "" & rst1![note] & ""
            itmX.SubItems(keys("SKP_CHANCE")) = "" & rst1![SKP_CHANCE] & ""
            itmX.SubItems(keys("Uznani_Dluhu")) = "" & rst1![flag_uznani_dluhu] & "" 'RISK-919
            itmX.SubItems(keys("SKP_HEIR")) = "" & rst1![skp_heir] & "" 'RISK-468
        rst1.MoveNext
        Next i
   End If

   rst1.Close


'Naplnìní komboboxù


ComboBox_Chance_Type.Clear

If strRizeni = "heir" Then
    ComboBox_Chance_Type.AddItem "dìdická šance"
   ' Call ComboBox_Chance_Type_Change
    SKP_PROCESS = 6 'HEIR øízení
    ComboBox_Chance_Type.ListIndex = 0
Else
    
     Select Case strRizeni 'RISK-468
        Case "ded"
            SKP_PROCESS = 1
        Case "exe"
            SKP_PROCESS = 2
        Case "epr"
            SKP_PROCESS = 3
        Case "ins"
            SKP_PROCESS = 4
        Case "trest"
            SKP_PROCESS = 5
     End Select

    ComboBox_Chance_Type.AddItem "druhá šance"
    ComboBox_Chance_Type.AddItem "poslední šance"
    ComboBox_Chance_Type.AddItem "SK urèený soudem"
    ComboBox_Chance_Type.AddItem "obživlá šance"
End If
'ComboBox_Chance_Type.ListIndex = 0

ComboBox_Instalment_Type.Clear
ComboBox_Instalment_Type.AddItem "stejná výše splátky až do splacení"
ComboBox_Instalment_Type.AddItem "mimoøádná splátka"
ComboBox_Instalment_Type.ListIndex = 0

'RISK-919

'Pouze pro pøípad dìdické šance: Chci to zablokovat pouze v pøípadì, kdy pøidává novou, ne když dìlá update

If strRizeni = "heir" Then

    If SKP_CHANCE_IMPORT = 0 Then
    
        For i = 1 To Zadani_Sance.ListView_Chance_Overview.ListItems.count
        
            If ListvZiskejHodnotu(Zadani_Sance.ListView_Chance_Overview, "Chance_Type", i) = "dìdická šance" Then
                MonthView_Chance_Given = ListvZiskejHodnotu(Zadani_Sance.ListView_Chance_Overview, "Chance_Date", i)
                MonthView_Chance_Given.Enabled = False
                ComboBox_Uznani_Dluhu.value = ListvZiskejHodnotu(Zadani_Sance.ListView_Chance_Overview, "Uznani_Dluhu", i)
                ComboBox_Uznani_Dluhu.Enabled = False
                Exit For
            End If
                
        Next i
    
    End If

End If

ComboBox_Uznani_Dluhu.Clear
ComboBox_Uznani_Dluhu.AddItem "Ano"
ComboBox_Uznani_Dluhu.AddItem "Ne"

End Sub



Private Sub UserForm_QueryClose(Cancel As Integer, _
                            CloseMode As Integer)
Dim ans
If CloseMode = vbFormControlMenu Then
    Cancel = True
    'Vyèistím
    SKP_CHANCE_IMPORT = 0
    
    
    ComboBox_Chance_Type.value = vbNullString
    ComboBox_Instalment_Type.value = vbNullString
    ComboBox_Uznani_Dluhu.value = vbNullString 'RISK-919
    TextBox_Amt_Inst.value = vbNullString
    TextBox_Note.value = vbNullString

    
    Call Me.Hide
End If
End Sub


' Zmìna barvy po najetí myši **************************************************************************************************************************
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Zrus.ForeColor = barvaCerna
CommandButton_Uloz.ForeColor = barvaCerna
CommandButton_Add_Chance.ForeColor = barvaCerna
CommandButton_Change_Chance.ForeColor = barvaCerna
CommandButton_Delete_Chance.ForeColor = barvaCerna
End Sub

Private Sub CommandButton_Zrus_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Zrus.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Uloz_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Uloz.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Add_Chance_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Add_Chance.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Change_Chance_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Change_Chance.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Delete_Chance_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Delete_Chance.ForeColor = barvaZelenaAirBank
End Sub




