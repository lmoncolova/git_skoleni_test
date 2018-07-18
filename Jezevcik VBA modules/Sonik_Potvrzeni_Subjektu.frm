VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sonik_Potvrzeni_Subjektu 
   Caption         =   "Soník - Zadávání subjektù"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19830
   OleObjectBlob   =   "Sonik_Potvrzeni_Subjektu.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sonik_Potvrzeni_Subjektu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit

Private potvrzeni As String
Private id_cuid As String
Private Date_Birth As Variant
Private Name_first As String
Private Name_last As String
Private TEXT_BIRTH_NUMBER As String


'RISK-992
Private Type DispOwner
    LngSKP_PERSON_DISPONENT As Long
    StrID_CUID_DISP As String
    StrNAME_FIRST_DISP As String
    StrNAME_LAST_DISP As String
    DtDATE_BIRTH_DISP As Date
    StrTEXT_BIRTH_NUMBER_DISP As String
End Type

Private Type DispDisponent
    LngSKP_PERSON_OWNER As Long
    StrID_CUID_OWNER As String
    StrNAME_FIRST_OWNER As String
    StrNAME_LAST_OWNER As String
    DtDATE_BIRTH_OWNER As Date
    StrTEXT_BIRTH_NUMBER_OWNER As String
End Type


'Ruèní pøidání nenalezeného klienta

Private Sub CommandButton_Pridej_Click()

Sonik_Rucni_Pridani_Klienta.Show

End Sub




'Ruèní pøidání karty neklienta

Private Sub CommandButton_Karta_Click()

Sonik_Rucni_Pridani_Karty.Show

End Sub




'Potvrzování subjektu
Private Sub ListView_Potvrzeni_DblClick()


'Dovolím klikání, jenom pokud tam je na co klikat, jinak to spadne
If ListView_Potvrzeni.ListItems.count > 0 Then


        'RISK-876 bohužel pro update hodnoty musím nechat tak jak je
        If ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Potvrzení") = "Nepotvrzeno" Then
           Call ListvVlozHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Potvrzení", "Potvrzeno")
        
        ElseIf ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Potvrzení") = "Potvrzeno" Then
           Call ListvVlozHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Potvrzení", "Zamítnuto")
        
        ElseIf ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Potvrzení") = "Zamítnuto" Then
           Call ListvVlozHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Potvrzení", "Potvrzeno")
        
        End If

End If

End Sub


Private Sub CommandButton_Reject_All_Click()

If MsgBox("Opravdu chcete zamítnout všechny dosud nepotvzené subjekty?", vbYesNo, "Hromadné zamítnutí subjektù") = vbYes Then

    If Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ListItems.count > 0 Then
    
        For i = 1 To Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ListItems.count
            'RISK-876
            If ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Potvrzení", i) = "Nepotvrzeno" Then
            
                Call ListvVlozHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Potvrzení", "Zamítnuto", i)
            
            End If
        
        Next i
       
    End If

End If

End Sub



'Negativní souèinnost - pokud nikoho nenajdou, tak se zadá negativka

Private Sub CommandButton_Negativka_Click()

Dim skp_person As Long 'RISK-985ZAP

'Pokud už byl ruènì pøidán klient pøes CUID, tak nedovolím udìlat negativku

If Flag_Existuje_Rucne_Pridany_Klient = 1 Then
MsgBox "Klient byl již ruènì pøidán, proto nelze zadat negativní souèinnost", vbExclamation, "Klient byl již ruènì pøidán pøes CUID"
Exit Sub
End If



'Pokud je tam ještì nìjaký nalezený subjekt, tak chci aby jej uživatel zamítnul
If Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ListItems.count > 0 Then

    For i = 1 To Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ListItems.count
    
        If ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Potvrzení", i) <> "Zamítnuto" Then
        
        MsgBox "Pro uložení negativní souèinnosti je tøeba zamítnout všechny subjekty", vbExclamation, "Nedostateèné potvrzení hodnot"
        
        Exit Sub
        
        End If
    
    Next i
   
End If


'Ulozeni negativky do parovaci tabulky

'Zjistim maximalni skp

'sql_max_skp = "select max(skp_confirm_pair) as max_skp from " & schema & ".sonik_confirm_pair"

'Dim max_skp As String
Dim sql_insert_negative As String
'
'   Set rst1 = New ADODB.Recordset
'   rst1.Open sql_max_skp, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'
'   If Not rst1.EOF Then
'    rst1.MoveFirst
'    max_skp = rst1![max_skp]
'    End If
'
'   rst1.Close
'
''Jenom pojistka pro první spuštìní
'If max_skp = "" Then
'max_skp = 0
'End If



'For I = 1 To Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ListItems.count
'
''Najdu si maximum z skp_subject toho listview - dìlám pro pøípad, že bych tam pøidal nìjaký ruènì a mìl by nové skp jakožto nové hledání
'If SKP_Subject_Negative < Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ListItems(I).text Then
'SKP_Subject_Negative = Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ListItems(I).text
'End If
'
'Next I

'Vùbec nedovolím pustit negativku, pokud bude ruèní pøidání
'sql_insert_negative = "insert into  " & schema & ".sonik_confirm_pair values (" & max_skp + 1 & "," & SKP_Subject & "," & SKP_datove_zpravy & "," & "null" & "," & "null" & "," & "null" & "," & "null" & "," & "null" & "," & "null" & "," & "sysdate" & "," & "'" & UCase(Main.username) & "'" & "," & "'Negativní'" & ")"

'RISK-992 úprava insertu
sql_insert_negative = "insert into  " & schema & ".sonik_confirm_pair (skp_confirm_pair,skp_subject,skp_data_message,skp_person,id_cuid,name_first,name_last,date_birth,text_birth_number,dtime_inserted,name_employee_inserted,name_confirmation_status) " & _
                      " values (" & schema & ".sonik_confirm_pair_seq.nextval, " & SKP_Subject & "," & SKP_datove_zpravy & "," & "null" & "," & "null" & "," & "null" & "," & "null" & "," & "null" & "," & "null" & "," & "sysdate" & "," & "'" & UCase(Main.username) & "'" & "," & "'Negativní'" & ")"


'Pouziji funkci z Jezevèíka a nainzertuju negativku
Call execute(sql_insert_negative)



'XXXX Pak tady musi byt kod, ktery ulozi ostatni vyhledane - zamitnute - subjekty


'Ulozeni do parovaci tabulky

'Zjistim maximalni skp

'sql_max_skp = "select max(skp_confirm_pair) as max_skp from " & schema & ".sonik_confirm_pair"

Dim sql_insert_pairs As String
'
'   Set rst1 = New ADODB.Recordset
'   rst1.Open sql_max_skp, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'
'   If Not rst1.EOF Then
'    rst1.MoveFirst
'    max_skp = rst1![max_skp]
'    End If
'
'   rst1.Close
'
''Jenom pojistka pro první spuštìní
'If max_skp = "" Then
'max_skp = 0
'End If


For i = 1 To Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ListItems.count

    'Pouziji funkci z Jezevèíka
    'RISK-876
    SKP_Subject = ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "ID_Subjektu", i)
    potvrzeni = ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Potvrzení", i)
    id_cuid = ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "ID_CUID", i)
    skp_person = ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "skp_person", i) 'RISK-985ZAP
    Date_Birth = ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Datum narození", i)
    Name_first = ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Jméno", i)
    Name_last = ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Pøíjmení", i)
    TEXT_BIRTH_NUMBER = ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Rodné èíslo", i)
    
    
    
    'XXXX- casem asi neni treba toto ukladat...... staci tam dat cuid a skp_person do tabulky confirm_pair
    
    'sql_insert_pairs = "insert into  " & schema & ".sonik_confirm_pair values (" & max_skp + i & "," & SKP_Subject & "," & SKP_datove_zpravy & "," & skp_person & "," & "'" & ID_CUID & "'" & "," & "'" & NAME_FIRST & "'" & "," & "'" & NAME_LAST & "'" & "," & "to_date('" & DATE_BIRTH & "','DD.MM.YYYY')" & "," & "'" & TEXT_BIRTH_NUMBER & "'" & "," & "sysdate" & "," & "'" & UCase(Main.username) & "'" & "," & "'" & potvrzeni & "'" & ")"
    'RISK-992 úprava insertu
    sql_insert_pairs = "insert into  " & schema & ".sonik_confirm_pair (skp_confirm_pair,skp_subject,skp_data_message,skp_person,id_cuid,name_first,name_last,date_birth,text_birth_number,dtime_inserted,name_employee_inserted,name_confirmation_status) " & _
                      " values (" & schema & ".sonik_confirm_pair_seq.nextval, " & SKP_Subject & "," & SKP_datove_zpravy & "," & skp_person & "," & "'" & id_cuid & "'" & "," & "'" & Name_first & "'" & "," & "'" & Name_last & "'" & "," & "to_date('" & Date_Birth & "','DD.MM.YYYY')" & "," & "'" & TEXT_BIRTH_NUMBER & "'" & "," & "sysdate" & "," & "'" & UCase(Main.username) & "'" & "," & "'" & potvrzeni & "'" & ")"
    
    
    
    Call execute(sql_insert_pairs)

Next i








'Teïka když potvrdím, tak refresh dat v celkovém listview, a pomocí selectu do DWH a refresh
Call nacteni_dat_Sonik_Datova_Zprava


'Promažu
Sonik_Datova_Zprava.TextBox_CUID.text = ""
Sonik_Datova_Zprava.TextBox_Jmeno.text = ""
Sonik_Datova_Zprava.TextBox_Prijmeni.text = ""
Sonik_Datova_Zprava.TextBox_DateBirth.text = ""
Sonik_Datova_Zprava.TextBox_RC.text = ""
Sonik_Datova_Zprava.TextBox_BankAccount.text = ""
Sonik_Datova_Zprava.TextBox_Card.text = ""
Sonik_Datova_Zprava.TextBox_Loan.text = ""


Me.Hide

'Promennou inicializuji zpet na nulu
Flag_Existuje_Rucne_Pridany_Klient = 0

End Sub







'Potvrzení a zavøení

Private Sub CommandButton_Uloz_Click()

Dim skp_person As Long 'RISK-985ZAP

Dim pocet_zamitnuto As Integer
Const sql_conf_pair_part As String = "(SKP_CONFIRM_PAIR,SKP_SUBJECT,SKP_DATA_MESSAGE,SKP_PERSON,ID_CUID,NAME_FIRST, NAME_LAST,DATE_BIRTH,TEXT_BIRTH_NUMBER,DTIME_INSERTED,NAME_EMPLOYEE_INSERTED,NAME_CONFIRMATION_STATUS)"

'Pokud nenajdu nic, tak bych mìl jít do negativky
If Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ListItems.count = 0 Then
MsgBox "Žádný subjekt nebyl nalezen, zvažte možnost zadání negativní souèinnosti", vbExclamation, "Žádný subjekt nebyl nalezen"

Exit Sub
End If

'Pokud uživatel zamítne úplnì vše, tak musí pustit negativku
'RISK-876
For i = 1 To Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ListItems.count
If ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Potvrzení", i) = "Zamítnuto" Then
pocet_zamitnuto = pocet_zamitnuto + 1
End If
Next i

If pocet_zamitnuto = Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ListItems.count Then

MsgBox "Všechny subjekty byly zamítnuty, zvažte možnost zadání negativní souèinnosti", vbExclamation, "Všechny subjekty zamítnuty"

Exit Sub
End If


For i = 1 To Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ListItems.count
'RISK-876
If ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Potvrzení", i) = "Nepotvrzeno" Then

MsgBox "Pro uložení je tøeba potvrdit nebo zamítnout všechny subjekty", vbExclamation, "Nedostateèné potvrzení hodnot"
Exit Sub

End If

Next i

'Ulozeni do parovaci tabulky

'Zjistim maximalni skp

'sql_max_skp = "select max(skp_confirm_pair) as max_skp from " & schema & ".sonik_confirm_pair"
'
'Dim max_skp As String
Dim potvrzeni As String
Dim sql_insert_pairs As String
Dim id_cuid As String

'   Set rst1 = New ADODB.Recordset
'   rst1.Open sql_max_skp, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'
'   If Not rst1.EOF Then
'    rst1.MoveFirst
'    max_skp = rst1![max_skp]
'    End If
'
'   rst1.Close
'
'
''Jenom pojistka pro první spuštìní
'If max_skp = "" Then
'max_skp = 0
'End If


For i = 1 To Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ListItems.count

    'Pouziji funkci z Jezevèíka
    'RISK-876
    SKP_Subject = ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "ID_Subjektu", i)
    potvrzeni = ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Potvrzení", i)
    id_cuid = ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "ID_CUID", i)
    skp_person = ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "skp_person", i)
    Date_Birth = ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Datum narození", i)
    Name_first = ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Jméno", i)
    Name_last = ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Pøíjmení", i)
    TEXT_BIRTH_NUMBER = ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "Rodné èíslo", i)
    
    sql_insert_pairs = "insert into  " & schema & ".sonik_confirm_pair (skp_confirm_pair,skp_subject,skp_data_message,skp_person,id_cuid,name_first,name_last,date_birth,text_birth_number,dtime_inserted,name_employee_inserted,name_confirmation_status) " & _
                       " values (" & schema & ".sonik_confirm_pair_seq.nextval, " & SKP_Subject & "," & SKP_datove_zpravy & "," & skp_person & "," & "'" & id_cuid & "'" & "," & "'" & Name_first & "'" & "," & "'" & Name_last & "'" & "," & "to_date('" & Date_Birth & "','DD.MM.YYYY')" & "," & "'" & TEXT_BIRTH_NUMBER & "'" & "," & "sysdate" & "," & "'" & UCase(Main.username) & "'" & "," & "'" & potvrzeni & "'" & ")"
    
    
    Call execute(sql_insert_pairs)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'RISK-992
    
    'V tomto bloku se podívám jestli neexistují disponenti, nebo klient není disponent a vyzvu uživatele jestli je chce.
    'Pokud ano, tak nahraju do potvrzených párù s poznámkou o co se jedná
    
    
    Dim sqlDispOwner As String, sqlDispDisponent As String
    Dim j As Integer
    
    'Vytvoøení instancí private typù definovaných nahoøe
    Dim DOW As DispOwner, DD As DispDisponent
    
    If potvrzeni = "Potvrzeno" Then 'Disponenty a majitele budu hledat pouze k potvrzeným subjektùm
    
        'Najdu k hledanému subjektu všechny jeho disponenty, kteøí ještì nejsou jako potvrzení v dané datové zprávì
        sqlDispOwner = "select SKP_PERSON_DISPONENT, ID_CUID_DISP, NAME_FIRST_DISP, NAME_LAST_DISP, DATE_BIRTH_DISP, TEXT_BIRTH_NUMBER_DISP from " & schema & ".sonik_v_rel_owner_disp " & _
                       " where skp_person_disponent not in (select skp_person from " & schema & ".sonik_confirm_pair WHERE name_confirmation_status = 'Potvrzeno' AND skp_data_message = " & SKP_datove_zpravy & ")" & _
                       " and skp_person_owner = " & skp_person
        
        Set rst1 = New ADODB.Recordset
        rst1.Open sqlDispOwner, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
        If Not rst1.EOF Then
            rst1.MoveFirst
            For j = 1 To rst1.RecordCount
                DOW.LngSKP_PERSON_DISPONENT = rst1![SKP_PERSON_DISPONENT]
                DOW.StrID_CUID_DISP = rst1![ID_CUID_DISP]
                DOW.StrNAME_FIRST_DISP = rst1![NAME_FIRST_DISP]
                DOW.StrNAME_LAST_DISP = rst1![NAME_LAST_DISP]
                DOW.DtDATE_BIRTH_DISP = rst1![DATE_BIRTH_DISP]
                DOW.StrTEXT_BIRTH_NUMBER_DISP = rst1![TEXT_BIRTH_NUMBER_DISP]
                
                'Dotaz uživateli jestli má zájem, pokud to bude otravovat, tak se dá vypnout
                If MsgBox("Potvrzený subjekt " & Name_first & " " & Name_last & "(" & id_cuid & ") má disponenta " & DOW.StrNAME_FIRST_DISP & " " & DOW.StrNAME_LAST_DISP & "(" & DOW.StrID_CUID_DISP & "). Chceš tohoto disponenta pøiøadit do potvrzených subjektù?", vbYesNo, "Disponent k pøiøazení") = vbYes Then
                
                    sql_insert_pairs = "insert into  " & schema & ".sonik_confirm_pair (skp_confirm_pair,skp_subject,skp_data_message,skp_person,id_cuid,name_first,name_last,date_birth,text_birth_number,dtime_inserted,name_employee_inserted,name_confirmation_status,DISPONENT_RELATION) " & _
                                       " values (" & schema & ".sonik_confirm_pair_seq.nextval, " & SKP_Subject & "," & SKP_datove_zpravy & "," & DOW.LngSKP_PERSON_DISPONENT & "," & "'" & DOW.StrID_CUID_DISP & "'" & "," & "'" & DOW.StrNAME_FIRST_DISP & "'" & "," & "'" & DOW.StrNAME_LAST_DISP & "'" & "," & "to_date('" & DOW.DtDATE_BIRTH_DISP & "','DD.MM.YYYY')" & "," & "'" & DOW.StrTEXT_BIRTH_NUMBER_DISP & "'" & "," & "sysdate" & "," & "'" & UCase(Main.username) & "'" & "," & "'" & potvrzeni & "'" & ", 'DISPONENT')"
                    
                    Call execute(sql_insert_pairs)
                    
                End If
                
                
            rst1.MoveNext
            Next j
        End If
        
        rst1.Close
        
        
        'Najdu k hledanému klienty, kde je hledaný disponentem, a kteøí ještì nejsou jako potvrzení v dané datové zprávì
        sqlDispDisponent = "select SKP_PERSON_OWNER, ID_CUID_OWNER, NAME_FIRST_OWNER, NAME_LAST_OWNER, DATE_BIRTH_OWNER, TEXT_BIRTH_NUMBER_OWNER from " & schema & ".sonik_v_rel_owner_disp" & _
        " where skp_person_owner not in (select skp_person from " & schema & ".sonik_confirm_pair WHERE name_confirmation_status = 'Potvrzeno' AND skp_data_message = " & SKP_datove_zpravy & ")" & _
                       " and skp_person_disponent = " & skp_person
    
        
        Set rst1 = New ADODB.Recordset
        rst1.Open sqlDispDisponent, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
        If Not rst1.EOF Then
            rst1.MoveFirst
            For j = 1 To rst1.RecordCount
                DD.LngSKP_PERSON_OWNER = rst1![skp_person_owner]
                DD.StrID_CUID_OWNER = rst1![ID_CUID_OWNER]
                DD.StrNAME_FIRST_OWNER = rst1![NAME_FIRST_OWNER]
                DD.StrNAME_LAST_OWNER = rst1![NAME_LAST_OWNER]
                DD.DtDATE_BIRTH_OWNER = rst1![DATE_BIRTH_OWNER]
                DD.StrTEXT_BIRTH_NUMBER_OWNER = rst1![TEXT_BIRTH_NUMBER_OWNER]
                
                'Dotaz uživateli jestli má zájem, pokud to bude otravovat, tak se dá vypnout
                If MsgBox("Potvrzený subjekt " & Name_first & " " & Name_last & "(" & id_cuid & ") je disponentem k " & DD.StrNAME_FIRST_OWNER & " " & DD.StrNAME_LAST_OWNER & "(" & DD.StrID_CUID_OWNER & "). Chceš tohoto vlastíka pøiøadit do potvrzených subjektù?", vbYesNo, "Vlastík k pøiøazení") = vbYes Then
                
                    sql_insert_pairs = "insert into  " & schema & ".sonik_confirm_pair (skp_confirm_pair,skp_subject,skp_data_message,skp_person,id_cuid,name_first,name_last,date_birth,text_birth_number,dtime_inserted,name_employee_inserted,name_confirmation_status,DISPONENT_RELATION) " & _
                                       " values (" & schema & ".sonik_confirm_pair_seq.nextval, " & SKP_Subject & "," & SKP_datove_zpravy & "," & DD.LngSKP_PERSON_OWNER & "," & "'" & DD.StrID_CUID_OWNER & "'" & "," & "'" & DD.StrNAME_FIRST_OWNER & "'" & "," & "'" & DD.StrNAME_LAST_OWNER & "'" & "," & "to_date('" & DD.DtDATE_BIRTH_OWNER & "','DD.MM.YYYY')" & "," & "'" & DD.StrTEXT_BIRTH_NUMBER_OWNER & "'" & "," & "sysdate" & "," & "'" & UCase(Main.username) & "'" & "," & "'" & potvrzeni & "'" & ", 'OWNER')"
                    
                    Call execute(sql_insert_pairs)
                    
                End If
                
                
            rst1.MoveNext
            Next j
        End If
        
        rst1.Close
        
    
    End If
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    

Next i

Me.Hide


'Teïka když potvrdím, tak refresh dat v celkovém listview, a pomocí selectu do DWH a refresh
Call nacteni_dat_Sonik_Datova_Zprava


'Promažu
Sonik_Datova_Zprava.TextBox_CUID.text = ""
Sonik_Datova_Zprava.TextBox_Jmeno.text = ""
Sonik_Datova_Zprava.TextBox_Prijmeni.text = ""
Sonik_Datova_Zprava.TextBox_DateBirth.text = ""
Sonik_Datova_Zprava.TextBox_RC.text = ""
Sonik_Datova_Zprava.TextBox_BankAccount.text = ""
Sonik_Datova_Zprava.TextBox_Card.text = ""
Sonik_Datova_Zprava.TextBox_Loan.text = ""

'Promennou inicializuji zpet na nulu
Flag_Existuje_Rucne_Pridany_Klient = 0

End Sub


' ListView sorting
Private Sub ListView_Potvrzeni_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_Potvrzeni, ColumnHeader)
End Sub



' Zavøení formuláøe
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Call set_filter(Me.name, 1)
Me.Hide
Sonik_Datova_Zprava.Show

'Promennou inicializuji zpet na nulu
Flag_Existuje_Rucne_Pridany_Klient = 0
End Sub


' Back button
Private Sub CommandButton_Navrat_Click()
Call text_information_hide(formular)
Call UserForm_QueryClose(1, 0)

'Promennou inicializuji zpet na nulu
Flag_Existuje_Rucne_Pridany_Klient = 0
End Sub

















' Zmìna barvy po najetí myši **************************************************************************************************************************

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Uloz.ForeColor = barvaCerna
CommandButton_Pridej.ForeColor = barvaCerna
End Sub

Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Navrat.ForeColor = barvaCerna
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Uloz.ForeColor = barvaCerna
CommandButton_Pridej.ForeColor = barvaCerna
End Sub

Private Sub CommandButton_Navrat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Navrat.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Uloz_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Uloz.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Pridej_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Pridej.ForeColor = barvaZelenaAirBank
End Sub
