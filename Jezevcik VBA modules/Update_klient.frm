VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Update_klient 
   Caption         =   "Jezevèík - Nalézací øízení"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11265
   OleObjectBlob   =   "Update_klient.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Update_klient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

Private Const cst_form_height_short As Double = 296.25  ' 224.25
Private Const cst_form_height_tall As Double = 447      '375 ' rozdiel
Private skp_soud_selected As Long                       ' v tehle promenne se uchovava posledni soud, ktery uzivatel zvolil v comboboxe


Private Sub ComboBox_okres_soud_Change()
    ' vzdy kdyz uzivatel vybere nejaky okresni soud z nabidnuteho seznamu, automaticky se propise i krajsky soud a predevsim se ulozi hodnota SKP_SOUD do modulove promenne
    
    Dim sql As String
    Dim rst As ADODB.Recordset
    
    Set rst = New ADODB.Recordset
    
    If (Len(Trim(ComboBox_okres_soud.value)) > 0) Then
    
        sql = "SELECT s.SKP_SOUD, s.KRAJSKY_SOUD FROM " & schema & ".soudy s WHERE s.okresni_soud LIKE '%" & ComboBox_okres_soud.value & "'"
            
        Set rst = New ADODB.Recordset
        rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
            If Not rst.EOF Then
            
                rst.MoveFirst
                If rst.RecordCount = 1 Then
                    
                    skp_soud_selected = rst![skp_soud]
                    ComboBox_kraj_soud.value = rst![krajsky_soud]
                    
                ElseIf rst.RecordCount > 1 Then
                    assertion False, "jez-64617: Bylo nalezeno více krajských soudù k jednomu okresnímu. Pošli prosím printscreen analytikùm."
                Else
                    assertion False, "jez-64618: K okresnímu soudù nebyl nalezen žádný krajský. Pošli prosím printscreen analytikùm."
                End If
            
            Else
                assertion False, "jez-64619: K okresnímu soudù nebyl nalezen žádný krajský. Pošli prosím printscreen analytikùm."
            End If
        rst.Close
    
        ' MsgBox "ComboBox_okres_soud_Change:  " & skp_soud_selected
    
    End If
    
End Sub

Private Sub CommandButton12_Click()
    Call Me.Hide
End Sub

Private Sub CommandButton20_Click()
    Dim BeforeValue As String
    Dim AfterValue As String
    Dim id As Integer
    
    Dim sql As String
    Dim sql_tab_PERSON As String
    Dim sql_tab_INFO As String
    Dim sql_tab_DEMOGRAPHY As String
    Dim sql_tab_PAIRED_MANUALLY As String
    Dim zmena As String
    
    Dim bln_zmena_tab_PERSON As Boolean
    Dim bln_zmena_tab_INFORMATION As Boolean
    Dim bln_zmena_tab_DEMOGRAPHY As Boolean
    Dim bln_zmena_tab_PAIRED_MANUALLY As Boolean
    
    
    If cs_HDWMN.State = 1 Then
prihlasen:
    
    ''''''      inicializace hodnot
    
    Call text_information_hide(Me.name)
    
    bln_zmena_tab_PERSON = False
    bln_zmena_tab_INFORMATION = False
    bln_zmena_tab_DEMOGRAPHY = False
    bln_zmena_tab_PAIRED_MANUALLY = False
    
    zmena = ""
    sql_tab_PERSON = "update " & schema & ".person set "
    sql_tab_DEMOGRAPHY = "update " & schema & ".person_demography set " 'RISK-517 DONE
    sql_tab_INFO = "update " & schema & ".information set "
    
    
    ''''''      nastaveni booleans a sql strings dle toho co zadal uzivatel
    
    '''   Jmeno klienta
    
    If Len(Trim(TextBox1.text)) > 0 And CheckBox1.value And TextBox1.text <> Label1.Caption Then
        sql_tab_PERSON = sql_tab_PERSON & "name = '" & TextBox1.text & "', "
        bln_zmena_tab_PERSON = True
        zmena = zmena & "jméno (" & TextBox1.text & "), "
    End If
    sql_tab_PERSON = sql_tab_PERSON & " dtime_modification = sysdate, employee = '" & Main.username & "' where id_person = " & id_person
    
    '''   Soud
    
    If (Len(Trim(ComboBox_okres_soud.text)) > 0) And (CheckBox_okres_soud.value) And (ComboBox_okres_soud.text <> Label_okres_soud.Caption) And (skp_soud_selected > 0) Then
        sql_tab_PAIRED_MANUALLY = "insert into " & schema & ".SOUDY_PAIRED_MANUALLY (ID_PERSON, SKP_SOUD, FLAG_VALID, DTIME_INSERTED) " & _
                                  " VALUES (" & c.id_person & ", " & skp_soud_selected & ", 'Y', SYSDATE)"  '  na tabulku je napojenej trigger ktery updatuje PERSON
        bln_zmena_tab_PAIRED_MANUALLY = True
        zmena = zmena & "soudy: (" & ComboBox_okres_soud.text & " (SKP_SOUD: " & skp_soud_selected & ")), "
    End If
    
    
    
    
    '''   Rodne cislo
    
    If (Len(Trim(TextBox2.text)) > 0) And (CheckBox2.value) And (TextBox2.text <> Label2.Caption) Then
        
        'Kontrola na správné zadání RÈ
        If TextBox2.value <> vbNullString And (Not IsNumeric(TextBox2.value) Or TextBox2.value <= 0 Or Len(TextBox2.value) < 9 Or Len(TextBox2.value) > 10) Then
            MsgBox "Rodné èíslo není vyplnìno správnì.", vbExclamation, "Chybné zadání hodnot"
            Exit Sub
        End If
        
        sql_tab_DEMOGRAPHY = sql_tab_DEMOGRAPHY & "text_birth_number = '" & TextBox2.text & "' " 'RISK-517 DONE
        bln_zmena_tab_DEMOGRAPHY = True
        zmena = zmena & "rè (" & TextBox2.text & "), "
    End If
    sql_tab_DEMOGRAPHY = sql_tab_DEMOGRAPHY & " where id_person = " & id_person
    
    
    
    If prava("select * from " & schema & ".information where id_person = " & id_person) Then
        
        '''   Datum zaloby
        
        If CheckBox6.value And DTPicker1.value <> Label6.Caption Then
            sql_tab_INFO = sql_tab_INFO & "date_action = to_date('" & DTPicker1.value & "','dd.mm.yyyy'), "
            bln_zmena_tab_INFORMATION = True
            zmena = zmena & "datum žaloby (" & DTPicker1.value & "), "
        End If
        
        '''   Spisova znacka
        
        If Len(Trim(TextBox5.text)) > 0 And CheckBox5.value And TextBox5.text <> Label5.Caption Then
            sql_tab_INFO = sql_tab_INFO & "code  = '" & TextBox5.text & "', "
            bln_zmena_tab_INFORMATION = True
            zmena = zmena & "Sp.zn.(" & TextBox5.text & "), "
        End If
        
        '''   Vyse pohledavky
        
        If CheckBox7.value And TextBox7.text <> Label7.Caption Then
            If Len(Trim(TextBox7.text)) > 0 And IsNumeric(TextBox7.text) Then
                sql_tab_INFO = sql_tab_INFO & "amt_claim  = " & Replace(Replace(TextBox7.text, ",", "."), " ", "") & ", "
                bln_zmena_tab_INFORMATION = True
                zmena = zmena & "pohledávka (" & TextBox7.text & "), "
            Else
                GoTo chyba
            End If
        End If
        
        ' If CheckBox8.Value And DTPicker2.Value <> Label8.Caption Then
        '  sql_tab_INFO = sql_tab_INFO & "date_calculation_claim = to_date('" & DTPicker2.Value & "','dd.mm.yyyy'), "
        '  bln_zmena_tab_INFORMATION = True
        '  zmena = zmena & "datum pohledávky (" & DTPicker2.Value & "), "
        ' End If
        '
        ' If Len(Trim(TextBox9.Text)) > 0 And CheckBox9.Value And TextBox9.Text <> Label9.Caption Then
        '  sql_tab_INFO = sql_tab_INFO & "id_inbox  = '" & TextBox9.Text & "', "
        '  bln_zmena_tab_INFORMATION = True
        '  zmena = zmena & "ID Datové schránky(" & TextBox9.Text & "), "
        ' End If
        
        sql_tab_INFO = sql_tab_INFO & " dtime_modification = sysdate, employee = '" & Main.username & "' where id_person = " & id_person
        
    Else
        
        BeforeValue = "id_person, "
        AfterValue = "(" & id_person & ", "
        sql_tab_INFO = "insert into " & schema & ".information ("
        
        '''   Datum zaloby
        
        If CheckBox6.value And DTPicker1.value <> Label6.Caption Then
            BeforeValue = BeforeValue & "date_action, "
            AfterValue = AfterValue & "to_date('" & DTPicker1.value & "','dd.mm.yyyy'), "
            bln_zmena_tab_INFORMATION = True
            zmena = zmena & "datum žaloby (" & DTPicker1.value & "), "
        End If
    
        '''   Spisova znacka
        
        If Len(Trim(TextBox5.text)) > 0 And CheckBox5.value And TextBox5.text <> Label5.Caption Then
            BeforeValue = BeforeValue & "code, "
            AfterValue = AfterValue & "'" & TextBox5.text & "', "
            bln_zmena_tab_INFORMATION = True
            zmena = zmena & "Sp.zn. (" & TextBox5.text & "), "
        End If
    
        '''   Vyse pohledavky
        
        If CheckBox7.value And TextBox7.text <> Label7.Caption Then
            If Len(Trim(TextBox7.text)) > 0 And IsNumeric(TextBox7.text) Then
                BeforeValue = BeforeValue & "amt_claim, "
                AfterValue = AfterValue & Replace(Replace(TextBox7.text, ",", "."), " ", "") & ", "
                bln_zmena_tab_INFORMATION = True
                zmena = zmena & "pohledávka (" & TextBox7.text & "), "
            Else
                GoTo chyba
            End If
        End If
        
        ' If CheckBox8.Value And DTPicker2.Value <> Label8.Caption Then
        '  BeforeValue = BeforeValue & "date_calculation_claim, "
        '  AfterValue = AfterValue & "to_date('" & DTPicker2.Value & "','dd.mm.yyyy'), "
        '  bln_zmena_tab_INFORMATION = True
        '  zmena = zmena & "datum pohledávky (" & DTPicker2.Value & "), "
        ' End If
        '
        ' If Len(Trim(TextBox9.Text)) > 0 And CheckBox9.Value And TextBox9.Text <> Label9.Caption Then
        '  BeforeValue = BeforeValue & "id_inbox, "
        '  AfterValue = AfterValue & "'" & TextBox9.Text & "', "
        '  bln_zmena_tab_INFORMATION = True
        '  zmena = zmena & "ID Datové schránky (" & TextBox9.Text & "), "
        ' End If
        
        
        BeforeValue = BeforeValue & "dtime_modification, employee)"
        AfterValue = AfterValue & "sysdate, '" & Main.username & "')"
        sql_tab_INFO = sql_tab_INFO & BeforeValue & " values " & AfterValue
        
    End If
    
    
    If bln_zmena_tab_PERSON Or bln_zmena_tab_INFORMATION Or bln_zmena_tab_DEMOGRAPHY Or bln_zmena_tab_PAIRED_MANUALLY Then 'RISK-517 DONE
    
        If MsgBox("Opravdu chcete uložit?", vbYesNo, "Uložit") = vbYes Then
            
            TransBegin
                
                If bln_zmena_tab_PERSON Then
                    execute (sql_tab_PERSON)
                End If
        
                If bln_zmena_tab_DEMOGRAPHY Then
                    execute (sql_tab_DEMOGRAPHY) 'RISK-517 DONE
                End If
          
                If bln_zmena_tab_INFORMATION Then
                    execute (sql_tab_INFO)
                End If
                
                If bln_zmena_tab_PAIRED_MANUALLY Then
                    execute (sql_tab_PAIRED_MANUALLY)
                End If
                
                sql = "insert into " & schema & ".epr_note values (" & max_skp("select max(skp_epr_note) from " & schema & ".epr_note") & ", " & id_person & _
                      ", 2, NULL, 'Zmìna údajù klienta: " & Mid(zmena, 1, Len(zmena) - 2) & "', sysdate, '" & Main.username & "')"
        
                execute (sql)
            
            TransCommit
            
            Me.Hide
            Klient.Show
            odkaz_Klient = False
        Else
            Label29.Visible = True
            Label28.Visible = True
            Label29.Caption = "Zrušeno"
        End If
    
    Else
        Label29.Visible = True
        Label28.Visible = True
        Label29.Caption = "Žádné zmìny"
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
    
    Exit Sub
    
chyba:

    Label29.Visible = True
    Label28.Visible = True
    Label29.Caption = "Chybný formát vstupu"
    
End Sub



Private Sub CheckBox1_Click()
Label1.Visible = Not CheckBox1.value
TextBox1.Visible = CheckBox1.value

TextBox1.text = Label1.Caption
End Sub

Private Sub CheckBox2_Click()
Label2.Visible = Not CheckBox2.value
TextBox2.Visible = CheckBox2.value

TextBox2.text = Label2.Caption
End Sub


Private Sub CheckBox5_Click()
Label5.Visible = Not CheckBox5.value
TextBox5.Visible = CheckBox5.value

TextBox5.text = Label5.Caption
End Sub

Private Sub CheckBox6_Click()
Label6.Visible = Not CheckBox6.value
DTPicker1.Visible = CheckBox6.value

If CheckBox6.value And Label6.Caption <> "" Then
DTPicker1 = Label6.Caption
ElseIf CheckBox6.value Then
DTPicker1 = date
End If

Me.Repaint
End Sub


Private Sub CheckBox7_Click()
Label7.Visible = Not CheckBox7.value
TextBox7.Visible = CheckBox7.value

TextBox7.text = Label7.Caption
End Sub



Private Sub CheckBox_okres_soud_Click()
    Label_okres_soud.Visible = Not CheckBox_okres_soud.value
    ComboBox_okres_soud.Visible = CheckBox_okres_soud.value
    If (Label_okres_soud.Caption <> cfg.cst_missing_string_value) Then
        ComboBox_okres_soud.value = Label_okres_soud.Caption  '  spadne pokud "_____XXXXXXX_____"
    End If
    ' ciselnik je ve skutecnosti cielnikem OKRESNICH soudu, krajsky nelze priradit pokud neni v ciselniku prirazenej k nejakemu okresnimu..
    CheckBox_kraj_soud.value = Not CheckBox_kraj_soud.value
    Call CheckBox_kraj_soud_Click
    
    If Me.ComboBox_okres_soud.ListCount < 1 Then
        Call NaplnSeznamOkresniSoudy(Me.ComboBox_okres_soud)
    End If
    
    Call ComboBox_okres_soud_Change
    ' MsgBox "CheckBox_okres_soud_Click:  " & skp_soud_selected
End Sub


Private Sub CheckBox_kraj_soud_Click()
    ' uzivatel tohle nezavola, vzdy jedine naraz s okresnim soudem.
    ' ve skutecnosti k nicemu neslouzi, je to jenom vizualni informace
    
    Label_kraj_soud.Visible = Not CheckBox_kraj_soud.value
    ComboBox_kraj_soud.Visible = CheckBox_kraj_soud.value
    If (Label_kraj_soud.Caption <> cfg.cst_missing_string_value) Then
        ComboBox_kraj_soud.value = Label_kraj_soud.Caption  '  spadne pokud "_____XXXXXXX_____"
    End If
    
    Me.Repaint
End Sub


Private Sub NaplnSeznamOkresniSoudy(ByRef YourComboBox As combobox)
    ' tahle sub naplni seznam hodnot do rolovatka Okresni Soud
    ' inicializace probiha na karte Klient pri kliknuti na button (search> '''  XXX RISK-1197 - tady je potreba nainicializovat labels pro okresni a krajsky soud)
    ' malinko to nedava smysl, chci totiz nejspis naplnit i ten druhy combobox (krajsky soud)
    
    Dim sql As String
    Dim rst As Object 'adodb.Recordset
    Dim i As Integer
    Dim str_okr_soud As String
    
    sql = "SELECT s.SKP_SOUD, s.OKRESNI_SOUD, s.KRAJSKY_SOUD FROM " & schema & ".soudy s ORDER BY s.OKRESNI_SOUD"
    
    YourComboBox.Clear
    
    Set rst = New ADODB.Recordset
    rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        If Not rst.EOF Then
            rst.MoveFirst
            For i = 1 To rst.RecordCount
            
                str_okr_soud = rst![okresni_soud]
                YourComboBox.AddItem str_okr_soud
                
            rst.MoveNext
            Next i
        End If
    rst.Close

    
End Sub



'' Vypnuté atributy

'Private Sub CheckBox8_Click()
'Label8.visible = Not CheckBox8.Value
'DTPicker2.visible = CheckBox8.Value
'
'If CheckBox8.Value And Label8.Caption <> "" Then
'DTPicker2 = Label8.Caption
'ElseIf CheckBox8.Value Then
'DTPicker2 = date
'End If
'
'Me.Repaint
'End Sub
'
'
'Private Sub CheckBox9_Click()
'Label9.visible = Not CheckBox9.Value
'TextBox9.visible = CheckBox9.Value
'
'TextBox9.Text = Label9.Caption
'End Sub

'*******************************************************************************

' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()
'RISK-517 RISK-952
Dim sql As String
Dim control As control
Dim keys As Dictionary
Dim itmX As ListItem

Me.Height = cst_form_height_short ' Iniciálnì nezobrazím Listview z párováním neklientù


Call text_information_hide(Me.name)

For Each control In Me.Controls
    'Debug.Print control.name & ",   tag: " & control.Tag & ",   visible before: " & control.Visible
    If control.Tag = "1" Then control.Visible = True
    If control.Tag = "2" Then control.Visible = False
    If control.Tag = "3" Then control.value = False
    'Debug.Print control.name & ",   tag: " & control.Tag & ",   visible after: " & control.Visible
Next

Label_okres_soud = Replace(c.SOUD.okresni_soud, cfg.cst_missing_string_value, "") ' c.SOUD.okresni_soud
Label_kraj_soud = Replace(c.SOUD.krajsky_soud, cfg.cst_missing_string_value, "") ' c.SOUD.krajsky_soud
If Me.ComboBox_okres_soud.ListCount < 2 Then
    Call NaplnSeznamOkresniSoudy(Me.ComboBox_okres_soud)
End If

'RISK-517 - dovolím zmìnu Rodného èísla pouze na neklientech - zároveò pouze u neklientù zobrazím párovací tabulku na klienty
If c.CUID <= 0 Then
        'CheckBox1.Enabled = True 'Jméno
        CheckBox2.Enabled = True 'RÈ
        'Frame3.ControlTipText = "" 'Frame se jménem
        Frame4.ControlTipText = "" 'Frame se RÈ
        
        
        'RISK-952 BEGIN -----------------------------------------------------------------------------------------------------------------------------------------------------------
        'Nahraju všechny ještì nenapárované páry neklienta
        sql = "SELECT SKP_NON_CLIENT_MATCH, ID_PERSON, CUID_DWH, NAME_LAST_DWH, NAME_FIRST_DWH, DATE_BIRTH_DWH, TEXT_BIRTH_NUMBER_DWH, MATCH_TYPE, DTIME_INSERTED " & _
        " FROM " & schema & ".NON_CLIENT_MATCH " & _
        " WHERE NAME_CONFIRMATION_STATUS IS NULL AND ID_PERSON = " & c.id_person
        
        Set rst1 = New ADODB.Recordset
        rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
        If Not rst1.EOF Then
                rst1.MoveFirst
                
                If Update_klient.ListView_Potvrzeni.ColumnHeaders.count = 0 Then
                
                        With Update_klient.ListView_Potvrzeni.ColumnHeaders
                                
                                .Add , , "SKP_NON_CLIENT_MATCH", 0 'Primární klíè v tabulce párù
                                .Add , , "ID_PERSON", 0
                                .Add , , "Potvrzení", 70
                                .Add , "CUID_DWH", "CUID", 50
                                .Add , "NAME_DWH", "Jméno", 100
                                .Add , "DATE_BIRTH_DWH", "Dat. nar.", 60
                                .Add , "TEXT_BIRTH_NUMBER_DWH", "Rodné èíslo", 90
                                .Add , "MATCH_TYPE", "Typ napárování", 110
                                .Add , "DTIME_INSERTED", "Datum napárování", 100
                        
                        End With
                        
                        Call NastavKliceListview(ListView_Potvrzeni)
                
                End If
                
                
                Update_klient.ListView_Potvrzeni.ListItems.Clear
                Set keys = KliceKListview(Update_klient.ListView_Potvrzeni)
                
                rst1.MoveFirst
                
                For i = 1 To rst1.RecordCount
                        Set itmX = Update_klient.ListView_Potvrzeni.ListItems.Add()
                        itmX.text = rst1![SKP_NON_CLIENT_MATCH] & ""
                        itmX.SubItems(keys("ID_PERSON")) = rst1![id_person] & ""
                        
                        itmX.SubItems(keys("Potvrzení")) = "Nepotvrzeno" 'Iniciálnì, hodnota co se bude mìnit klikem uživatele
                        
                        itmX.SubItems(keys("CUID_DWH")) = rst1![CUID_DWH] & ""
                        itmX.SubItems(keys("NAME_DWH")) = rst1![NAME_LAST_DWH] & " " & rst1![NAME_FIRST_DWH] & ""
                        itmX.SubItems(keys("DATE_BIRTH_DWH")) = rst1![DATE_BIRTH_DWH] & ""
                        itmX.SubItems(keys("TEXT_BIRTH_NUMBER_DWH")) = rst1![TEXT_BIRTH_NUMBER_DWH] & ""
                        itmX.SubItems(keys("MATCH_TYPE")) = rst1![MATCH_TYPE] & ""
                        itmX.SubItems(keys("DTIME_INSERTED")) = rst1![DTIME_INSERTED] & ""
                        
                        rst1.MoveNext
                Next i
                
                Me.Height = cst_form_height_tall 'Zobrazím Listview s páry
                
        End If
        rst1.Close
        'RISK-952 END -----------------------------------------------------------------------------------------------------------------------------------------------------------
        
Else
        'CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        'Frame3.ControlTipText = "Jméno lze mìnit pouze u neklientù - u klientù pøes OBS"
        Frame4.ControlTipText = "Rodné èíslo lze mìnit pouze u neklientù - u klientù pøes OBS"
End If
 
End Sub



Private Sub ListView_Potvrzeni_DblClick()
'RISK-952
'Dovolím klikání, jenom pokud tam je na co klikat, jinak to spadne
If ListView_Potvrzeni.ListItems.count > 0 Then


        'RISK-876 bohužel pro update hodnoty musím nechat tak jak je
        If ListvZiskejHodnotu(ListView_Potvrzeni, "Potvrzení") = "Nepotvrzeno" Then
           Call ListvVlozHodnotu(ListView_Potvrzeni, "Potvrzení", "Potvrzeno")
        
        ElseIf ListvZiskejHodnotu(ListView_Potvrzeni, "Potvrzení") = "Potvrzeno" Then
           Call ListvVlozHodnotu(ListView_Potvrzeni, "Potvrzení", "Zamítnuto")
        
        ElseIf ListvZiskejHodnotu(ListView_Potvrzeni, "Potvrzení") = "Zamítnuto" Then
           Call ListvVlozHodnotu(ListView_Potvrzeni, "Potvrzení", "Potvrzeno")
        
        End If

End If

End Sub



Private Sub CommandButton_Potvrzeni_Click()
'RISK-952
Dim i As Integer, IntCounterConfirmed As Integer, IntCounterNonDecided As Integer
'Prvnì vybere pouze jednoho klienta, toho potvrdí a když dá potvrdit tak listview se schová a data se pøedvyplní
'V pøípadì pøedvyplnìných datech již nesmím povolit editaci
Dim StrSKPConfirmedPair As String
Dim sql As String

If ListView_Potvrzeni.ListItems.count > 0 Then
    
    
    For i = 1 To ListView_Potvrzeni.ListItems.count
        If ListvZiskejHodnotu(ListView_Potvrzeni, "Potvrzení", i) = "Potvrzeno" Then
            IntCounterConfirmed = IntCounterConfirmed + 1
            StrSKPConfirmedPair = ListvZiskejHodnotu(ListView_Potvrzeni, "SKP_NON_CLIENT_MATCH", i)
        
        ElseIf ListvZiskejHodnotu(ListView_Potvrzeni, "Potvrzení", i) = "Nepotvrzeno" Then
            IntCounterNonDecided = IntCounterNonDecided + 1
            
        End If
        
    Next i


    'Nesmí být potvrzený více než jeden, jinak neuložím
    If IntCounterConfirmed > 1 Then
        MsgBox "Pro uložení nesmí být potrvzeno více subjektù jeden subjekt", vbExclamation, "Nadbyteèné potvrzení hodnot"
        Exit Sub
    ElseIf IntCounterNonDecided > 0 Then
        MsgBox "O osudu nìkterých párù nebylo rozhodnuto. Potvrï prosím nebo zamítni všechny.", vbExclamation, "Nedostateèné potvrzení hodnot"
        Exit Sub
    End If
    
    
    'Update dat v tabulce
    
    'Pokud došlo pouze k zamítnutí tak
    If StrSKPConfirmedPair = "" Then StrSKPConfirmedPair = "-99946134" 'nic se nepotvrdí - jakákoliv záporná hodnota
    
    sql = "update " & schema & ".NON_CLIENT_MATCH " & _
         " set NAME_CONFIRMATION_STATUS = CASE WHEN SKP_NON_CLIENT_MATCH = " & StrSKPConfirmedPair & " then 'Potvrzeno' else 'Zamítnuto' end " & _
         " , USER_CONFIRMED = '" & UCase(JezUserLoggedIn.StrUserName) & "'" & _
         " , DTIME_CONFIRMED = sysdate" & _
         " where name_confirmation_status is null and id_person = " & c.id_person
         
    execute (sql)
    
    If StrSKPConfirmedPair > 0 Then
            MsgBox "Neklient " & c.name & " CUID: " & c.CUID & " byl úspìšnì napárován na klienta.", vbInformation, "Transformace probìhla úspìšnì."
    End If
        
    Me.Height = cst_form_height_short 'Skryju Listview
        
    assertion (c.id_person = id_person)
        
    'Na konci znovu nainicializuji klienta, aby mìl již nový CUID a SKP_PERSON
    Call c.initialize(Val(id_person))
        
End If


End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, _
                            CloseMode As Integer)
Dim ans
If CloseMode = vbFormControlMenu Then
    Cancel = True
    Call Me.Hide
End If
End Sub


' Zmìna barvy po najetí myši **************************************************************************************************************************
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton12.ForeColor = barvaCerna
CommandButton20.ForeColor = barvaCerna

End Sub

Private Sub CommandButton12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton12.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton20.ForeColor = barvaZelenaAirBank
End Sub


