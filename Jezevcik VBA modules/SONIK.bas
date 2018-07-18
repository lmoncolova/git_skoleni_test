Attribute VB_Name = "SONIK"
Option Explicit

Dim clmx As ColumnHeader
Dim itmX As ListItem
Dim i As Integer
Dim j As Integer
Dim sql As String

'RISK-317


'Nadeklarovani globalnich promennych

Public SKP_datove_zpravy As Long
Public Cislo_datove_zpravy As String
Public Datum_doruceni_datove_zpravy As Date
Public Odesilatel_datove_zpravy As String
Public Cislo_Jednaci_datove_zpravy As String

Public SKP_Subject As Long
Public Rucne_Pridani_Cuid As Long
Public Rucne_Pridana_Karta As Long
Public Hl_Card As String
Public Spisova_Znacka_datove_zpravy As String

Public DM_Import As Workbook
Public Workbook_Open As Integer
Public CntRows_Import As Long

Public Flag_Existuje_Rucne_Pridany_Klient As Long 'Pomocn· promÏnn·, slouûÌ pro zabr·nÏnÌ zaloûenÌ negativnÌ souËinnosti, pokud jiû uûivatel ruËnÏ p¯idal klienta

Public Flag_Nove_Upozorneni_DZ As Long 'Pomocn· promÏnn·, slouûÌ pro rozliöenÌ p¯i zad·v·nÌ upozornÏnÌ jestli chci upravit st·vajÌcÌ nebo vytvo¯it novÈ


Public skp_person_KL As Long 'Pro kartu klienta
Public ID_CUID_KL As String 'Pro kartu klienta
Public NAME_KL As String 'Pro kartu klienta
Public NUM_Card_KL As String 'Pro kartu klienta
Public SKP_DAMAGE As Long 'Pro zadani vyse skody

'Public Data_Export As Workbook 'RISK-1235 - p¯esunuto jako promÏnn· v na danÈm formu a potÈ vol·no jako parametr funkce


'Vytvo¯Ìm kolekce, do kter˝ch si budu ukl·dat to co uû uûivatel vybral na str·nce, kde vybÌr· informace o datovÈ zpr·vÏ
Public coll_provided_info As New Collection
Public coll_req_info As New Collection

Public listview_datove_karty As String 'pouziva se pro vyber karet, ktere chci exportovat do excelu


Public Sonik_User As Boolean 'Pouziva se pri prihlaseni do JezevËÌka


Public SonikDM As SonikDataMessage_class 'RISK-1235


Sub nacteni_dat_Sonik_Datove_Zpravy_Prehled(Bln_Display_Sample As Boolean)

Dim sql_dz As String

'Tiptext - najedou myöÌ a zobrazÌ se n·povÏda

Sonik_Datove_Zpravy_Prehled.TextBox_DateDelivery.ControlTipText = "Zadej datum ve form·tu 17.3.2017."
Sonik_Datove_Zpravy_Prehled.TextBox_DateBirth.ControlTipText = "Zadej datum ve form·tu 17.3.2017."

With Sonik_Datove_Zpravy_Prehled.ListView_Datove_Zpravy.ColumnHeaders

    If .count = 0 Then
    
    'Nastaveni listview datovych zprav
        .Add , , "ID", 1
        .Add , , "»Ìslo datovÈ zpr·vy", 80
        .Add , , "Datum doruËenÌ", 70
        .Add , , "OdesÌlatel zpr·vy", 200
        .Add , , "Spisov· znaËka", 120
        .Add , , "»Ìslo jednacÌ", 120
        .Add , , "Datum zaevidov·nÌ v SonÌku", 120
        .Add , , "Zaevidoval", 60
        .Add , , "Stav datovÈ zpr·vy", 90
    
    End If

End With


'NaËtenÌ datov˝ch zpr·v
sql_dz = "select SKP_DATA_MESSAGE, DATA_MESSAGE_NUMBER, date_data_message_delivery, DATA_MESSAGE_SENDER, TEXT_FILE_NUMBER, TEXT_REFERENCE_NUMBER, DTIME_INSERTED, NAME_EMPLOYEE_INSERTED, NAME_DATA_MESSAGE_STATUS from " & schema & ".sonik_v_data_message"

If Bln_Display_Sample Then 'nebudou se naËÌtat vöechny datovky, ale jen p·r. Budou mÌt moûnost si zobrazit vöechny
    sql_dz = OmezPocetRadkuVQuery(sql_dz)
End If


Call query_datove_zpravy(sql_dz)


'BarviËky podle stav˘

With Sonik_Datove_Zpravy_Prehled.ListView_Datove_Zpravy

    If .ListItems.count > 0 Then
    
        For i = 1 To .ListItems.count
    
            If (.ListItems(i).SubItems(8) = "Naimportovan·" Or .ListItems(i).SubItems(8) = "Naimportov·no") And .ListItems(i).SubItems(5) <> vbNullString Then
            'Pokud je naimportovan·, ale uû bylo zad·no ËÌslo jednacÌ
            
            For j = 1 To .ColumnHeaders.count - 1
                .ListItems(i).ListSubItems.item(j).ForeColor = RGB(255, 140, 0)
            Next j
            
            ElseIf .ListItems(i).SubItems(8) = "Naimportovan·" Or .ListItems(i).SubItems(8) = "Naimportov·no" Then
            
            For j = 1 To .ColumnHeaders.count - 1
                .ListItems(i).ListSubItems.item(j).ForeColor = barvaRuda
            Next j
            
            ElseIf .ListItems(i).SubItems(8) = "SouËinnost zad·na" Or .ListItems(i).SubItems(8) = "Zaevidov·no" Then
            
            For j = 1 To .ColumnHeaders.count - 1
                .ListItems(i).ListSubItems.item(j).ForeColor = RGB(0, 204, 102)
            Next j
    
            ElseIf .ListItems(i).SubItems(8) = "»ek· na zpracov·nÌ" Or .ListItems(i).SubItems(8) = "»·steËnÏ zpracov·no" Then
            
            For j = 1 To .ColumnHeaders.count - 1
                .ListItems(i).ListSubItems.item(j).ForeColor = barvaModra
            Next j
            
            ElseIf .ListItems(i).SubItems(8) = "PoskytnutÌ info" Or .ListItems(i).SubItems(8) = "Zpracov·no" Then
            
            For j = 1 To .ColumnHeaders.count - 1
                .ListItems(i).ListSubItems.item(j).ForeColor = barvaCerna
            Next j
            
            End If
    
        Next i
    
    End If

End With

'NaËtenÌ comboboxu pro vyhled·v·nÌ dle stav˘

'' NastavenÌ comboboxu
 Sonik_Datove_Zpravy_Prehled.ComboBox_Stav_DZ.Clear
 
 'Pro zacatek vlozim prazdny ¯·dek, aby se podle nÏj nehledalo
 Sonik_Datove_Zpravy_Prehled.ComboBox_Stav_DZ.AddItem ""
 
 sql = "select NAME_DATA_MESSAGE_STATUS from " & schema & ".sonik_cl_data_message_status" 'chci i dnes neplatnÈ stavy
 
 Set rst1 = New ADODB.Recordset
 rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

 If Not rst1.EOF Then
  rst1.MoveFirst
  For i = 0 To rst1.RecordCount - 1
    Sonik_Datove_Zpravy_Prehled.ComboBox_Stav_DZ.AddItem "" & rst1![NAME_DATA_MESSAGE_STATUS] & ""
    rst1.MoveNext
  Next i
 End If
 rst1.Close

 If Sonik_Datove_Zpravy_Prehled.ComboBox_Stav_DZ.ListCount > 0 Then
    Sonik_Datove_Zpravy_Prehled.ComboBox_Stav_DZ.ListIndex = 0
 End If



End Sub




Sub nacteni_dat_Sonik_Datova_Zprava()

Dim sql_dz_history As String
Dim sql_req_info As String
Dim sql_provided_info As String
Dim sql_confirmed_pairs As String
Dim sql_negativky As String
Dim sql_dz_notice As String


Sonik_Datova_Zprava.TextBox_DateBirth.ControlTipText = "Zadej datum ve form·tu 13.7.2017."

With Sonik_Datova_Zprava

    .Label_ID_DZ.Caption = SKP_datove_zpravy
    .Label_Dz_Number.Caption = Cislo_datove_zpravy
    .Label_DZ_Date_Delivery.Caption = Datum_doruceni_datove_zpravy
    .Label_DZ_Sender.Caption = Odesilatel_datove_zpravy
    .Label_DZ_Reference_NUM.Caption = Cislo_Jednaci_datove_zpravy
    .Label_DZ_File_NUM.Caption = Spisova_Znacka_datove_zpravy

End With


With Sonik_Datova_Zprava.ListView_DZ_History.ColumnHeaders

    'Nastaveni listview stavu a poznamek
    If .count = 0 Then
     
     .Add , , "DATA_MESSAGE_STATUS_ID", 0
     .Add , , "DATA_MESSAGE_NOTE_ID", 0
     .Add , , "DATA_MESSAGE_NOTICE_ID", 0
     .Add , , "Datum ud·losti", 80
     .Add , , "Typ ud·losti", 60
     .Add , , "Ud·lost", 250
     .Add , , "Provedl", 70
     .Add , , "Detail", 40 'Flag, jestli existuje dodateËn· pozn·mka
     .Add , , "DodateËn· pozn·mka", 0 'Nebudu ji re·lnÏ zobrazovat ale jen si ji uloûÌm a pak zobrazÌm v novÈm oknÏ
    
    End If

End With

'NaËtenÌ historie na datovÈ zpr·vÏ
sql_dz_history = "select DATA_MESSAGE_STATUS_ID, DATA_MESSAGE_NOTE_ID, DATA_MESSAGE_NOTICE_ID, DTIME_INSERTED, EVENT_TYPE, EVENT, NAME_EMPLOYEE_INSERTED, ADDITIONAL_NOTE from " & schema & ".sonik_v_data_message_history where skp_data_message = " & SKP_datove_zpravy & " order by DTIME_INSERTED desc"
Call query_dz_history(sql_dz_history)



'Nastaveni listview pozadovanych informaci
With Sonik_Datova_Zprava.ListView_pozad_info.ColumnHeaders
    If .count = 0 Then
    
     .Add , , "ID", 1
     .Add , , "N·zev poûadovanÈ informace", 90
     .Add , , "»as vloûenÌ", 100
     .Add , , "Vloûil", 60
    
    End If
End With
'NaËtenÌ vÏcÌ, na co se policie pt·
sql_req_info = "select a.id, b.name_req_prov_info, a.dtime_inserted, a.name_employee_inserted " & _
" from " & schema & ".sonik_req_info a" & _
" join " & schema & ".sonik_cl_req_prov_info b on a.skp_req_info = b.skp_req_prov_info" & _
" where a.skp_data_message = " & SKP_datove_zpravy

Call query_pozad_info(sql_req_info)




'Nastaveni listview poskytnutych informaci
With Sonik_Datova_Zprava.ListView_Poskyt_Info.ColumnHeaders
    If .count = 0 Then
    
        .Add , , "ID", 1
        .Add , , "N·zev poskytovanÈ informace", 90
        .Add , , "»as vloûenÌ", 100
        .Add , , "Vloûil", 60
    
    End If
End With

'NaËtenÌ vÏcÌ, na co se policie pt·
sql_provided_info = "select a.id, b.name_req_prov_info, a.dtime_inserted, a.name_employee_inserted  " & _
" from " & schema & ".sonik_provided_info a" & _
" join " & schema & ".sonik_cl_req_prov_info b on a.skp_provided_info = b.skp_req_prov_info" & _
" where a.skp_data_message = " & SKP_datove_zpravy

Call query_poskyt_info(sql_provided_info)






'Nastaveni listview naparovanych subjektu
'RISK-876 - LU jako LOOKUP - hledan· hodnota
With Sonik_Datova_Zprava.ListView_Paired_Subjects.ColumnHeaders
    If .count = 0 Then
    
     .Add , , "ID", 1
     .Add , , "ID Subjektu", 1
     .Add , "JmÈno LU", "JmÈno", 50
     .Add , "P¯ÌjmenÌ LU", "P¯ÌjmenÌ", 50
     .Add , "Datum narozenÌ LU", "Datum narozenÌ", 50
     .Add , "RodnÈ ËÌslo LU", "RodnÈ ËÌslo", 80
     .Add , "»Ìslo ˙Ëtu LU", "»Ìslo ˙Ëtu", 50
     .Add , "»Ìslo karty LU", "»Ìslo karty", 50
     .Add , "»Ìslo ˙vÏru LU", "»Ìslo ˙vÏru", 50
     .Add , "»Ìslo telefonu LU", "»Ìslo telefonu", 50
     .Add , , "SKP_PERSON", 0
     .Add , , "CUID", 60
     .Add , , "Klient", 90
     .Add , , "Datum narozenÌ", 50
     .Add , , "RodnÈ ËÌslo", 80
     .Add , , "Datum potvrzenÌ", 100
     .Add , , "Potvrzeno", 80
     .Add , , "DISPONENT_RELATION", 0 'RISK-992
    
    
    'RISK-876 - p¯id·nÌ klÌËe pro moûnost zpÏtnÈho odk·z·nÌ se
     Call NastavKliceListview(Sonik_Datova_Zprava.ListView_Paired_Subjects)
    
    End If
End With

'RISK-992 - p¯id·nÌ DISPONENT_RELATION sloupce
'RISK-876
'NaËtenÌ potvrzen˝ch p·r˘ k danÈ souËinnosti ' lu jako zkratka lookup
sql_confirmed_pairs = "select b.SKP_CONFIRM_PAIR, b.SKP_SUBJECT, " & _
" a.name_first as name_first_lu, a.name_last as name_last_lu, nullif(a.date_birth,date'3000-01-01') as date_birth_lu, a.text_birth_number as text_birth_number_lu, a.text_bank_account_number as text_bank_account_number_lu, a.text_card_number as text_card_number_lu, a.code_loan_number as code_loan_number_lu, a.text_phone_number as text_phone_number_lu," & _
" b.skp_person,b.id_cuid , b.NAME_FIRST ||' '|| b.NAME_LAST as PERSON_NAME, b.DATE_BIRTH, b.TEXT_BIRTH_NUMBER ,b.DTIME_INSERTED, b.NAME_EMPLOYEE_INSERTED, b.DISPONENT_RELATION" & _
" from " & schema & ".sonik_subject a " & _
" join " & schema & ".sonik_confirm_pair b on a.skp_subject = b.skp_subject " & _
" where b.NAME_CONFIRMATION_STATUS = 'Potvrzeno' and b.SKP_DATA_MESSAGE = " & SKP_datove_zpravy

Call query_confirmed_pairs(sql_confirmed_pairs)




'Nastaveni listview negativek
With Sonik_Datova_Zprava.ListView_Negativky.ColumnHeaders
    If .count = 0 Then
    
     .Add , , "ID", 1
     .Add , , "ID Subjektu", 1
     .Add , , "JmÈno", 50
     .Add , , "P¯ÌjmenÌ", 50
     .Add , , "Datum narozenÌ", 50
     .Add , , "RodnÈ ËÌslo", 80
     .Add , , "»Ìslo ˙Ëtu", 50
     .Add , , "»Ìslo karty", 50
     .Add , , "»Ìslo ˙vÏru", 50
     .Add , , "»Ìslo telefonu", 80
     .Add , , "Datum potvrzenÌ", 100
     .Add , , "Potvrzeno", 80
     
    
    'RISK-876 - p¯id·nÌ klÌËe pro moûnost zpÏtnÈho odk·z·nÌ se
    Call NastavKliceListview(Sonik_Datova_Zprava.ListView_Negativky)
    
    End If
End With

'RISK-876
'NaËtenÌ potvrzen˝ch p·r˘ k danÈ souËinnosti
sql_negativky = "select b.SKP_CONFIRM_PAIR, b.SKP_SUBJECT, " & _
" a.name_first, a.name_last, nullif(a.date_birth,date'3000-01-01') as date_birth, a.text_birth_number, a.text_bank_account_number, a.text_card_number, a.code_loan_number, a.text_phone_number, " & _
" b.DTIME_INSERTED, b.NAME_EMPLOYEE_INSERTED" & _
" from " & schema & ".sonik_subject a " & _
" join " & schema & ".sonik_confirm_pair b on a.skp_subject = b.skp_subject " & _
" where b.NAME_CONFIRMATION_STATUS = 'NegativnÌ' and b.SKP_DATA_MESSAGE = " & SKP_datove_zpravy

Call query_negativky(sql_negativky)





'Nastaveni listview upozornÏnÌ
With Sonik_Datova_Zprava.ListView_DZ_Upozorneni.ColumnHeaders
    If .count = 0 Then
     
       .Add , , "ID", 0
       .Add , , "UpozornÏnÌ", 200
       .Add , , "Deadline", 60
       .Add , , "Zb˝v·", 60
    
    End If
End With






'NaËtenÌ historie na datovÈ zpr·vÏ
sql_dz_notice = "select SKP_DATA_MESSAGE_NOTICE, TEXT_NOTICE, DTIME_NOTICE_DEADLINE from " & schema & ".sonik_data_message_notice where skp_data_message = " & SKP_datove_zpravy & " and code_status = 'AktivnÌ' order by DTIME_NOTICE_DEADLINE"
Call query_dz_notice(sql_dz_notice)


'Pokud je nÏjakÈ upozornÏnÌ po termÌnu, nebo se blÌûÌ, tak vybarvÌm

If Sonik_Datova_Zprava.ListView_DZ_Upozorneni.ListItems.count > 0 Then

    For i = 1 To Sonik_Datova_Zprava.ListView_DZ_Upozorneni.ListItems.count

        If Sonik_Datova_Zprava.ListView_DZ_Upozorneni.ListItems(i).SubItems(3) <= 0 Then
        
        For j = 1 To Sonik_Datova_Zprava.ListView_DZ_Upozorneni.ColumnHeaders.count - 1
        Sonik_Datova_Zprava.ListView_DZ_Upozorneni.ListItems(i).ListSubItems.item(j).ForeColor = barvaRuda
        Next j

        ElseIf Sonik_Datova_Zprava.ListView_DZ_Upozorneni.ListItems(i).SubItems(3) <= 5 Then
        
        For j = 1 To Sonik_Datova_Zprava.ListView_DZ_Upozorneni.ColumnHeaders.count - 1
        Sonik_Datova_Zprava.ListView_DZ_Upozorneni.ListItems(i).ListSubItems.item(j).ForeColor = RGB(255, 100, 0)
        Next j
        
        End If

    Next i

End If



End Sub


Sub nacteni_dat_Sonik_Potvrzeni_Subjektu()
 
With Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ColumnHeaders
    If .count = 0 Then
    'RISK-876
    'Nastaveni listview
     Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ListItems.Clear
     .Add , , "ID_Subjektu", 1
     .Add , , "PotvrzenÌ", 80
     .Add , , "ID_CUID", 60
     .Add , , "JmÈno", 80
     .Add , , "P¯ÌjmenÌ", 80
     .Add , , "Datum narozenÌ", 70
     .Add , , "RodnÈ ËÌslo", 60
     .Add , , "Trval· adresa", 150
     .Add , , "»Ìsla ˙Ët˘", 100
     .Add , , "»Ìsla ˙vÏr˘", 100
     .Add , , "»Ìsla telefon˘", 100
     .Add , , "skp_person", 0
    
    'RISK-876 - p¯id·nÌ klÌËe pro moûnost zpÏtnÈho odk·z·nÌ se
    Call NastavKliceListview(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni)
    
    End If
End With
End Sub


Sub nacteni_dat_Info_DZ()

Dim sql_dz_cl_req_info As String
Dim sql_dz_cl_provided_info As String


If Sonik_Rucni_Pridani_Info_DZ.ListView_Provided_Info.ColumnHeaders.count = 0 Then
'Nastaveni listview
 Sonik_Rucni_Pridani_Info_DZ.ListView_Provided_Info.ListItems.Clear
 Sonik_Rucni_Pridani_Info_DZ.ListView_Provided_Info.ColumnHeaders.Add , , "Poskytnuto", 120
End If

If Sonik_Rucni_Pridani_Info_DZ.ListView_Req_Info.ColumnHeaders.count = 0 Then
'Nastaveni listview
 Sonik_Rucni_Pridani_Info_DZ.ListView_Req_Info.ListItems.Clear
 Sonik_Rucni_Pridani_Info_DZ.ListView_Req_Info.ColumnHeaders.Add , , "Poûadov·no", 120
End If




'NaËtenÌ ciselniku pozadovane info
sql_dz_cl_req_info = "select NAME_REQ_PROV_INFO from " & schema & ".sonik_cl_req_prov_info where flag_valid = 'Y'"
Call query_cl_Req_Info_dz(sql_dz_cl_req_info)



'NaËtenÌ ciselniku poskytnute info - nep˘jde z ËÌselnÌku, ale z toho, co uû bylo vybr·no jako poûadov·no - RISK-768
sql_dz_cl_provided_info = "select NAME_REQ_PROV_INFO from " & schema & ".sonik_cl_req_prov_info where flag_valid = 'Y'"
Call query_cl_Provided_Info_dz(sql_dz_cl_provided_info)


End Sub





Sub nacteni_dat_Vybrane_Info_DZ()

If Sonik_Rucni_Pridani_Info_DZ.ListBox_Info_Selected.ColumnHeaders.count = 0 Then
'Nastaveni listview
 Sonik_Rucni_Pridani_Info_DZ.ListBox_Info_Selected.ListItems.Clear
 Sonik_Rucni_Pridani_Info_DZ.ListBox_Info_Selected.ColumnHeaders.Add , , "Vybr·no", 120
End If

End Sub



Sub nacteni_dat_Sonik_Update_Stavu_DZ()

Dim sql_historie_stavu_DZ As String

With Sonik_Update_Stav_DZ.ListView_Historie_Stavu.ColumnHeaders
    If .count = 0 Then
     
    'Nastaveni listview
     Sonik_Update_Stav_DZ.ListView_Historie_Stavu.ListItems.Clear
     .Add , , "ID", 1
     .Add , , "Stav", 70
     .Add , , "Pozn·mka", 120
     .Add , , "Platnost od", 90
     .Add , , "Platnost do", 90
     .Add , , "Stav p¯idal", 80
    
    End If
End With

sql_historie_stavu_DZ = "select s.id,cl.name_data_message_status, n.text_note ,s.dtime_valid_from, s.dtime_valid_to, s.name_employee_inserted " & _
" from " & schema & ".sonik_data_message_status s " & _
" join " & schema & ".sonik_cl_data_message_status cl on s.skp_data_message_status = cl.skp_data_message_status " & _
" left join " & schema & ".sonik_data_message_note n on s.skp_data_message = n.skp_data_message and s.id = n.id_data_message_status " & _
" where s.skp_data_message = " & SKP_datove_zpravy & _
" order by s.dtime_valid_to desc "

Call query_historie_stavu_dz(sql_historie_stavu_DZ)


' NastavenÌ comboboxu
 Sonik_Update_Stav_DZ.ComboBox_Stavy.Clear
 
 sql = "select NAME_DATA_MESSAGE_STATUS from " & schema & ".sonik_cl_data_message_status where flag_valid = 'Y'"
 
 Set rst1 = New ADODB.Recordset
 rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
 
 If Not rst1.EOF Then
  rst1.MoveFirst
  For i = 0 To rst1.RecordCount - 1
    Sonik_Update_Stav_DZ.ComboBox_Stavy.AddItem "" & rst1![NAME_DATA_MESSAGE_STATUS] & ""
    rst1.MoveNext
  Next i
 End If
 rst1.Close
 
 If Sonik_Update_Stav_DZ.ComboBox_Stavy.ListCount > 0 Then
    Sonik_Update_Stav_DZ.ComboBox_Stavy.ListIndex = 0
 End If




End Sub




Sub nacteni_dat_Karta_Klienta()


Dim sql_kk_client_soucinnosti As String
Dim sql_kk_data_client_actual As String
Dim sql_kk_data_client_gc As String
Dim sql_kk_data_client_Birth_Number_Change As String
Dim sql_kk_data_client_Name_Change As String
Dim sql_kk_data_client_Change_Mail_Phone As String
Dim sql_kk_data_client_Change_Address As String
Dim sql_kk_data_client_login_ib As String
Dim sql_kk_data_client_accounts As String
Dim sql_kk_data_client_cards As String
Dim sql_kk_data_client_transactions As String
Dim sql_kk_data_client_loan_app As String
Dim sql_kk_data_client_mort_app As String
Dim sql_kk_data_client_manual_verification As String
Dim sql_kk_data_client_amg_r8 As String
Dim sql_kk_data_client_amg_note As String
Dim sql_kk_data_client_amg_calls As String
Dim sql_kk_data_client_brki As String
Dim sql_kk_data_client_communication As String
Dim sql_kk_data_transactions_homer As String
Dim sql_kk_data_transactions_homer_agg As String
Dim sql_kk_data_disponent_rights As String 'RISK-992
Dim sql_kk_data_application_form_data As String 'RISK-884










'Pokud jdu z JezevËÌka, tak uz mam nacteno
'RISK-721
If userform_prv = "Sonik_Datova_Zprava" Then
'RISK-876
NUM_Card_KL = ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Paired_Subjects, "»Ìslo karty LU")
skp_person_KL = ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Paired_Subjects, "SKP_PERSON")
ID_CUID_KL = ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Paired_Subjects, "CUID")
NAME_KL = ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Paired_Subjects, "Klient")


End If


'Nacteni vsech soucinnosti, ve kter˝ch klient figuruje

With Sonik_Karta_Klienta.ListView_Soucinnosti_Klienta.ColumnHeaders
    If .count = 0 Then
     
     
    'Nastaveni listview
     Sonik_Karta_Klienta.ListView_Soucinnosti_Klienta.ListItems.Clear
     .Add , , "SKP_DATA_MESSAGE", 0
     .Add , , "»Ìslo datovÈ zpr·vy", 70
     .Add , , "Datum doruËenÌ", 70
     .Add , , "OdesÌlatel zpr·vy", 70
     .Add , , "Spisov· znaËka", 70
     .Add , , "»Ìslo jednacÌ", 70
     .Add , , "Stav", 70
     .Add , , "SKP_DAMAGE", 0
     .Add , , "V˝öe ökody", 80
     .Add , , "Pozn·mka ke ökodÏ", 80
     
    
    
    End If
End With

If skp_person_KL = -1 Then

sql_kk_client_soucinnosti = "select a.skp_data_message, a.data_message_number, a.date_data_message_delivery, a.data_message_sender, a.text_file_number, a.text_reference_number, c.name_data_message_status, e.skp_damage ,e.amt_damage, e.text_note" & _
" from " & schema & ".sonik_data_message a" & _
" join " & schema & ".sonik_data_message_status b on a.skp_data_message = b.skp_data_message and b.flag_valid = 'Y'" & _
" join " & schema & ".sonik_cl_data_message_status c on b.skp_data_message_status = c.skp_data_message_status" & _
" join " & schema & ".sonik_confirm_pair d on a.skp_data_message = d.skp_data_message and d.name_confirmation_status = 'Potvrzeno'" & _
" join " & schema & ".sonik_subject s on s.skp_subject = d.skp_subject" & _
" left join " & schema & ".sonik_damage e on a.skp_data_message = e.skp_data_message and d.skp_person = e.skp_person" & _
" where s.text_card_number = " & NUM_Card_KL

Else
'RISK-721
sql_kk_client_soucinnosti = "select a.skp_data_message, a.data_message_number, a.date_data_message_delivery, a.data_message_sender, a.text_file_number, a.text_reference_number, c.name_data_message_status, e.skp_damage ,e.amt_damage, e.text_note" & _
" from " & schema & ".sonik_data_message a" & _
" join " & schema & ".sonik_data_message_status b on a.skp_data_message = b.skp_data_message and b.flag_valid = 'Y'" & _
" join " & schema & ".sonik_cl_data_message_status c on b.skp_data_message_status = c.skp_data_message_status" & _
" join " & schema & ".sonik_confirm_pair d on a.skp_data_message = d.skp_data_message and d.name_confirmation_status = 'Potvrzeno'" & _
" left join " & schema & ".sonik_damage e on a.skp_data_message = e.skp_data_message and d.skp_person = e.skp_person" & _
" where a.skp_data_message > 0 and d.skp_person = " & skp_person_KL

End If

Call query_kk_client_soucinnosti(sql_kk_client_soucinnosti)




'Nacteni jednotlivych datovych oken na karte kliena

With Sonik_Karta_Klienta.ListView_Client_Actual.ColumnHeaders
'Data o klientovi aktu·lnÌ
    If .count = 0 Then
     
    'Nastaveni listview
     Sonik_Karta_Klienta.ListView_Client_Actual.ListItems.Clear
     .Add , , "JmÈno", 70
     .Add , , "P¯ÌjmenÌ", 70
     .Add , , "Titul p¯ed", 70
     .Add , , "Titul za", 70
     .Add , , "RodnÈ ËÌslo", 70
     .Add , , "Dat. nar.", 70
     '.Add, , "N·rodnost", 70
     .Add , , "Trv. adr. (Ulice, ».p., PS», MÏsto, St·t )", 70
     .Add , , "Kor. adr. (Ulice, ».p., PS», MÏsto, St·t )", 70
     .Add , , "Prim·rnÌ telefon", 70
     .Add , , "Prim·rnÌ email", 70
     .Add , , "PohlavÌ", 70
     .Add , , "Datum ˙mrtÌ", 70
    
    End If
End With

sql_kk_data_client_actual = "select NAME_FIRST , NAME_LAST, Text_Title_before, Text_Title_After, TEXT_BIRTH_NUMBER, DATE_BIRTH, /*NAME_NATIONALITY,*/ TRVALA_ADRESA, KORESPONDENCNI_ADRESA, TEXT_MOBILE_PHONE, TEXT_EMAIL_ADDRESS, NAME_GENDER, DATUM_UMRTI from " & schema & ".sonik_klient_aktual where skp_person = " & skp_person_KL

Call query_kk_data_client_actual(sql_kk_data_client_actual)




With Sonik_Karta_Klienta.ListView_Client_GC.ColumnHeaders
'Data o klientovi - RS
    If .count = 0 Then
     
     
    'Nastaveni listview
     Sonik_Karta_Klienta.ListView_Client_GC.ListItems.Clear
     .Add , , "Stav RS", 70
     .Add , , "Podpis RS", 70
     .Add , , "»Ìslo RS", 70
     .Add , , "Platnost stavu RS od", 70
     .Add , , "Platnost stavu RS do", 70
    
    End If
End With
sql_kk_data_client_gc = "select NAME_CONTRACT_STATUS, NAME_CONTRACT_WFSTATUS, DTIME_CONTRACT_SIGNATURE, TEXT_CONTRACT_NUMBER, PLATNOST_OD, PLATNOST_DO from " & schema & ".sonik_klient_ab_rs where skp_person = " & skp_person_KL & " order by PLATNOST_DO desc nulls first"

Call query_kk_data_client_gc(sql_kk_data_client_gc)


With Sonik_Karta_Klienta.ListView_Birth_Number_Change.ColumnHeaders
'Data o klientovi zmÏna R»
    If .count = 0 Then
     
    'Nastaveni listview
     Sonik_Karta_Klienta.ListView_Birth_Number_Change.ListItems.Clear
     .Add , , "RodnÈ ËÌslo", 70
     .Add , , "Platnost R» od", 70
     .Add , , "Platnost R» do", 70
    
    End If
End With

sql_kk_data_client_Birth_Number_Change = "select TEXT_ATTRIBUTE_VALUE, PLATNOST_OD, PLATNOST_DO from " & schema & ".sonik_klient_rc where skp_person = " & skp_person_KL & "  order by PLATNOST_DO desc nulls first"

Call query_kk_data_client_Birth_Number_Change(sql_kk_data_client_Birth_Number_Change)




'Data o klientovi zmÏna jmena, prijmeni
With Sonik_Karta_Klienta.ListView_Name_Change.ColumnHeaders
    If .count = 0 Then
     
     
    'Nastaveni listview
     Sonik_Karta_Klienta.ListView_Name_Change.ListItems.Clear
     .Add , , "JmÈno", 70
     .Add , , "P¯ÌjmenÌ", 70
     .Add , , "Titul p¯ed", 70
     .Add , , "Titul za", 70
     .Add , , "Dat. nar.", 70
     .Add , , "N·rodnost", 70
     .Add , , "UpdatujÌcÌ oper·tor", 70
     .Add , , "Platnost od", 70
     .Add , , "Platnost do", 70
     .Add , , "Platnost stavu z·znamu", 70
     
    
    End If
End With

sql_kk_data_client_Name_Change = "select NAME_FIRST, NAME_LAST, TEXT_TITLE_BEFORE, TEXT_TITLE_AFTER, DATE_BIRTH, NAME_COUNTRY, CODE_MODIFIED_BY, PLATNOST_OD, PLATNOST_DO, FLAG_CURRENT from " & schema & ".sonik_klient_jmeno where skp_person = " & skp_person_KL & "  order by PLATNOST_DO desc nulls first"

Call query_kk_data_client_Name_Change(sql_kk_data_client_Name_Change)




'Data o klientovi zmÏna Email, Mobil
If Sonik_Karta_Klienta.ListView_Change_Mail_Phone.ColumnHeaders.count = 0 Then
 
 
'Nastaveni listview
 Sonik_Karta_Klienta.ListView_Change_Mail_Phone.ListItems.Clear
 Set clmx = Sonik_Karta_Klienta.ListView_Change_Mail_Phone.ColumnHeaders.Add(, , "JmÈno", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Change_Mail_Phone.ColumnHeaders.Add(, , "P¯ÌjmenÌ", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Change_Mail_Phone.ColumnHeaders.Add(, , "Typ kontaktu", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Change_Mail_Phone.ColumnHeaders.Add(, , "Kontakt", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Change_Mail_Phone.ColumnHeaders.Add(, , "Platnost stavu z·znamu", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Change_Mail_Phone.ColumnHeaders.Add(, , "Aktualizace z·znamu", 70)
 

End If

sql_kk_data_client_Change_Mail_Phone = "select NAME_FIRST, NAME_LAST, CODE_CONTACT_TYPE, TEXT_PERSON_CONTACT, CODE_CONTACT_STATUS, AKTUALIZACE from " & schema & ".sonik_klient_kontakt where skp_person = " & skp_person_KL & "  order by to_date(aktualizace,'YYYY-MM-DD HH24:MI:SS') desc nulls first"

Call query_kk_data_client_Change_Mail_Phone(sql_kk_data_client_Change_Mail_Phone)



'Data o klientovi zmÏna adresy
If Sonik_Karta_Klienta.ListView_Address_Change.ColumnHeaders.count = 0 Then
 
 
'Nastaveni listview
 Sonik_Karta_Klienta.ListView_Address_Change.ListItems.Clear
 Set clmx = Sonik_Karta_Klienta.ListView_Address_Change.ColumnHeaders.Add(, , "JmÈno", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Address_Change.ColumnHeaders.Add(, , "P¯ÌjmenÌ", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Address_Change.ColumnHeaders.Add(, , "Typ adresy", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Address_Change.ColumnHeaders.Add(, , "Ulice", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Address_Change.ColumnHeaders.Add(, , "»Ìslo popisnÈ", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Address_Change.ColumnHeaders.Add(, , "SmÏrovacÌ ËÌslo", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Address_Change.ColumnHeaders.Add(, , "MÏsto", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Address_Change.ColumnHeaders.Add(, , "St·t", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Address_Change.ColumnHeaders.Add(, , "Platnost adresy od", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Address_Change.ColumnHeaders.Add(, , "Platnost stavu z·znamu", 70)

 

End If

sql_kk_data_client_Change_Address = "select JMENO, PRIJMENI, NAME_ADDRESS_TYPE, ULICE, TEXT_HOUSE_NUM, CODE_ZIP, NAME_TOWN, NAME_COUNTRY, CAS_ZMENY, FLAG_CURRENT from " & schema & ".sonik_klient_adresa where skp_person = " & skp_person_KL & "  order by CAS_ZMENY desc nulls first"

Call query_kk_data_client_Change_Address(sql_kk_data_client_Change_Address)



'Data o klientovi login do IB
If Sonik_Karta_Klienta.ListView_Login_IB.ColumnHeaders.count = 0 Then
 
 
'Nastaveni listview
 Sonik_Karta_Klienta.ListView_Login_IB.ListItems.Clear
 Set clmx = Sonik_Karta_Klienta.ListView_Login_IB.ColumnHeaders.Add(, , "JmÈno", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Login_IB.ColumnHeaders.Add(, , "P¯ÌjmenÌ", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Login_IB.ColumnHeaders.Add(, , "UûivatelskÈ jmÈno", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Login_IB.ColumnHeaders.Add(, , "IP Adresa", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Login_IB.ColumnHeaders.Add(, , "»as zaloûenÌ", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Login_IB.ColumnHeaders.Add(, , "»as odhl·öenÌ", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Login_IB.ColumnHeaders.Add(, , "Kan·l", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Login_IB.ColumnHeaders.Add(, , "⁄spÏönÈ provedenÌ poûadavku", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Login_IB.ColumnHeaders.Add(, , "KÛd operace", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Login_IB.ColumnHeaders.Add(, , "Info ke kÛdu", 70)


 

End If

sql_kk_data_client_login_ib = "select NAME_FIRST, NAME_LAST, TEXT_USER_NAME, IP_ADDRESS, DTIME_CREATED, DTIME_LOGGED_OUT, NAME_COMMON_CHANNEL, FLAG_REQUEST_SUCCESSFUL, CODE_OPERATION, CODE_INFO   from " & schema & ".sonik_login_ib where skp_person = " & skp_person_KL & "  order by dtime_created desc nulls first"

Call query_kk_data_client_login_ib(sql_kk_data_client_login_ib)


'Data o klientovych uctech
If Sonik_Karta_Klienta.ListView_Accounts.ColumnHeaders.count = 0 Then
 
'RISK-992 odmaz·nÌ nepot¯ebn˝ch sloupeËk˘
'Nastaveni listview
 Sonik_Karta_Klienta.ListView_Accounts.ListItems.Clear
 Set clmx = Sonik_Karta_Klienta.ListView_Accounts.ColumnHeaders.Add(, , "Typ ˙Ëtu", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Accounts.ColumnHeaders.Add(, , "MÏna", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Accounts.ColumnHeaders.Add(, , "»Ìslo ˙Ëtu", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Accounts.ColumnHeaders.Add(, , "Stav ˙Ëtu", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Accounts.ColumnHeaders.Add(, , "Datum aktivace ˙Ëtu", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Accounts.ColumnHeaders.Add(, , "ZaloûenÌ û·dosti", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Accounts.ColumnHeaders.Add(, , "Kan·l û·dosti", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Accounts.ColumnHeaders.Add(, , "Datum podpisu", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Accounts.ColumnHeaders.Add(, , "Kan·l podpisu", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Accounts.ColumnHeaders.Add(, , "Zp˘sob podpisu", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Accounts.ColumnHeaders.Add(, , "PoboËka podpisu", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Accounts.ColumnHeaders.Add(, , "ID zamÏstnance p¯i podpisu", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Accounts.ColumnHeaders.Add(, , "⁄ËetnÌ z˘statek", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Accounts.ColumnHeaders.Add(, , "DisponibilnÌ z˘statek", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Accounts.ColumnHeaders.Add(, , "IP adresa ze û·dosti", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Accounts.ColumnHeaders.Add(, , "Datum ukonËenÌ ˙Ëtu", 70)




 

End If

'RISK-992
sql_kk_data_client_accounts = "select BANK_ACCOUNT_TYPE, CODE_CURRENCY, TEXT_BANK_ACCOUNT_NUMBER, NAME_BANK_ACCOUNT_STATUS, DATE_BANK_ACCOUNT_ACTIVATION, DTIME_OF_APPLICATION, APP_CHANNEL_TYPE_REQUEST, DTIME_SIGNATURE, SIGNATURE_CHANNEL, SIGNATURE_TYPE, BRANCH_SIGN, NAME_PERSON_SIGN, UCETNI_ZUSTATEK, DISP_ZUSTATEK, IP_ADDRESS_APPR_SIGN, DATE_BANK_ACCOUNT_TERMINATION from " & schema & ".sonik_ucty where skp_person = " & skp_person_KL

Call query_kk_data_client_accounts(sql_kk_data_client_accounts)




'RISK-992
'Data o disponentsk˝ch pr·vech
If Sonik_Karta_Klienta.ListView_Disponent_Rights.ColumnHeaders.count = 0 Then
    'RISK-992
    'Nastaveni listview
    
    With Sonik_Karta_Klienta.ListView_Disponent_Rights
        .ListItems.Clear
        .ColumnHeaders.Add , "PERSON_ACCOUNT_REL", "Vztah", 70
        .ColumnHeaders.Add , "CUID_OWNER", "CUID vlastnÌka", 70
        .ColumnHeaders.Add , "NAME_FIRST_OWNER", "JmÈno vlastnÌka", 70
        .ColumnHeaders.Add , "NAME_LAST_OWNER", "P¯ÌjmenÌ vlastnÌka", 70
        .ColumnHeaders.Add , "CUID_DISPONENT", "CUID disponenta", 70
        .ColumnHeaders.Add , "NAME_FIRST_DISPONENT", "JmÈno disponenta", 70
        .ColumnHeaders.Add , "NAME_LAST_DISPONENT", "P¯ÌjmenÌ disponenta", 70
        .ColumnHeaders.Add , "CODE_RIGHT_TYPE", "Typ disponentskÈho pr·va", 70
        .ColumnHeaders.Add , "TEXT_BANK_ACCOUNT_NUMBER", "»Ìslo ˙Ëtu", 70
        .ColumnHeaders.Add , "DTIME_RIGHT_VALID_FROM", "Pr·vo od", 100
        .ColumnHeaders.Add , "DTIME_RIGHT_VALID_TO", "Pr·vo do", 100
    End With
End If



Call NastavKliceListview(Sonik_Karta_Klienta.ListView_Disponent_Rights)

sql_kk_data_disponent_rights = "SELECT PERSON_ACCOUNT_REL, CUID_OWNER, NAME_FIRST_OWNER, NAME_LAST_OWNER, CUID_DISPONENT, NAME_FIRST_DISPONENT, NAME_LAST_DISPONENT, CODE_RIGHT_TYPE, TEXT_BANK_ACCOUNT_NUMBER, DTIME_RIGHT_VALID_FROM, DTIME_RIGHT_VALID_TO from " & schema & ".sonik_disponents where skp_person = " & skp_person_KL

Call query_kk_data_disponent_rights(sql_kk_data_disponent_rights)




'Data o klientovych kartach
If Sonik_Karta_Klienta.ListView_Cards.ColumnHeaders.count = 0 Then
 
 
'Nastaveni listview
 Sonik_Karta_Klienta.ListView_Cards.ListItems.Clear
 Set clmx = Sonik_Karta_Klienta.ListView_Cards.ColumnHeaders.Add(, , "»Ìslo ˙Ëtu", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Cards.ColumnHeaders.Add(, , "»Ìslo karty", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Cards.ColumnHeaders.Add(, , "»as û·dosti o kartu", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Cards.ColumnHeaders.Add(, , "Kan·l û·dosti", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Cards.ColumnHeaders.Add(, , "Platnost karty od", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Cards.ColumnHeaders.Add(, , "Platnost karty do", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Cards.ColumnHeaders.Add(, , "»as aktivace", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Cards.ColumnHeaders.Add(, , "»as blokace karty", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Cards.ColumnHeaders.Add(, , "»as ukonËenÌ karty", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Cards.ColumnHeaders.Add(, , "»as odblokov·nÌ karty", 70) 'RISK-868
 Set clmx = Sonik_Karta_Klienta.ListView_Cards.ColumnHeaders.Add(, , "Stav karty", 70) 'RISK-943
 Set clmx = Sonik_Karta_Klienta.ListView_Cards.ColumnHeaders.Add(, , "Technologie", 70) 'RISK-943



 

End If

'RISK-868 'RISK-943
sql_kk_data_client_cards = "select TEXT_BANK_ACCOUNT_NUMBER, TEXT_CARD_NUMBER, DTIME_CARD_APPLICATION, CARD_CHANNEL_TYPE_REQ, DTIME_CARD_VALID_FROM, DTIME_CARD_VALID_TO, DTIME_CARD_ACTIVATION, DTIME_CARD_BLOCKED, DTIME_CARD_CLOSED,DTIME_CARD_UNBLOCKED, NAME_CARD_STATUS, NAME_TECHNOLOGY from " & schema & ".sonik_karty where skp_person = " & skp_person_KL

Call query_kk_data_client_cards(sql_kk_data_client_cards)



'Data o klientovych karetnich transakcich
If Sonik_Karta_Klienta.ListView_Transactions.ColumnHeaders.count = 0 Then
 
 
'Nastaveni listview
 Sonik_Karta_Klienta.ListView_Transactions.ListItems.Clear
 Set clmx = Sonik_Karta_Klienta.ListView_Transactions.ColumnHeaders.Add(, , "»Ìslo ˙Ëtu", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Transactions.ColumnHeaders.Add(, , "»Ìslo karty", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Transactions.ColumnHeaders.Add(, , "ProvedenÌ transakce", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Transactions.ColumnHeaders.Add(, , "Za˙Ëtov·nÌ transakce", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Transactions.ColumnHeaders.Add(, , "Blokovan· Ë·stka", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Transactions.ColumnHeaders.Add(, , "Za˙Ëtovan· Ë·stka", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Transactions.ColumnHeaders.Add(, , "MÏna", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Transactions.ColumnHeaders.Add(, , "Za˙Ëtovan· transakce", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Transactions.ColumnHeaders.Add(, , "Typ transakce", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Transactions.ColumnHeaders.Add(, , "Kategorie transakce", 70)


 

End If

sql_kk_data_client_transactions = "select TEXT_BANK_ACCOUNT_NUMBER, TEXT_CARD_NUMBER, DTIME_BLOCKING, DATE_ACCOUNTING, AMT_BLOCKED, AMT_CONFIRMED, CODE_CURRENCY, FLAG_TRANSACTION_ACCOUNTED, TYP_TRANSAKCE, KATEGORIE_TRANSAKCE from " & schema & ".sonik_karty_trans where skp_person = " & skp_person_KL

Call query_kk_data_client_transactions(sql_kk_data_client_transactions)


'Data o klientovych û·dostech o ˙vÏr
If Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.count = 0 Then
 'RISK-955
 
 
'RISK-884 - chtÏli zmÏnit i ˙daje na z·loûce û·dost o ˙vÏr
'Nastaveni listview
 Sonik_Karta_Klienta.ListView_Loan_Application.ListItems.Clear
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "»as 2.scoringu", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "»as odesl·nÌ do banky", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "Typ û·dosti", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "»Ìslo û·dosti", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "»Ìslo ˙vÏru", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "⁄vÏrov˝ r·mec", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "Stav û·dosti", 70)
 
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "Kan·l û·dosti", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "IP Adresa - kan·l û·dosti", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "Kan·l dokonËenÌ", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "IP Adresa - kan·l dokonËenÌ", 70)
 
  Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "Kan·l podpisu", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "Dat. podpisu", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "PoboËka podpisu", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "ID zamÏstnance", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "D˘vod zamÌtnutÌ", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "Datum kompletace", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "Stav kompletace", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "D˘vod zamÌtnutÌ kompletace", 70)
' RISK-884 P¯esouv· se do aplikaËnÌch dat
' Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "Spl·tky jin˝ch ˙vÏr˘", 70)
' Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "Typ p¯Ìjmu", 70)
' Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "N·zev zamÏstnavatele", 70)
' Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "I»/DI»", 70)
' Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "PracovnÌ pomÏr", 70)
' Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "ZamÏstn·n od", 70)
' Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "ZamÏstn·n do", 70)
' Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "V˝öe ËistÈho p¯Ìjmu", 70)
' Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "Typ vedl. p¯Ìjmu", 70)
' Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "N·zev zamÏstnavatele - vedl. p¯Ìjem", 70)
' Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "I»/DI» zamÏstnavatele - vedl. p¯Ìjem", 70)
' Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "PracovnÌ pomÏr - vedl. p¯Ìjem", 70)
' Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "ZamÏstn·n od - vedl. p¯Ìjem", 70)
' Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "ZamÏstn·n do - vedl. p¯Ìjem", 70)
' Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "V˝öe ËistÈho p¯Ìjmu - vedl. p¯Ìjem", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "Datum  naËerp·nÌ", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "Datum ukonËenÌ", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "Datum storna", 70)
 Set clmx = Sonik_Karta_Klienta.ListView_Loan_Application.ColumnHeaders.Add(, , "Nesplacen· jistina", 70)



End If
'RISK-955 RISK-884
sql_kk_data_client_loan_app = "select CAS2SCORING, CAS_ODESLANI_ZIB, DRUH_ZADOSTI, CISLO_ZADOSTI, CISLO_UVERU, VYSE_UVERU, VYSLEDEK_ZADOSTI, REQ_IP_ADDRESS, REQ_C_TYPE, FIN_IP_ADDRESS, FIN_C_TYPE, KANAL_PODPISU_PUJCKY, DATUM_PODPISU_PUJCKY, POBOCKA_PUJCKA, ID_ZAMESTNANCA_PUJCKA, DUVOD_ZAMITNUTI_CELKOVO, COMPL_FI_1, STAV_KOMPL_UVER_PRISLIB, DUVOD_ZAMITNUTI_CLIENT_KOMPL, DATUM_NACERPANI, DATUM_UKONCENI, DATUM_STORNA, NESPLACENA_JISTINA from " & schema & ".sonik_apl_data_uver where skp_person = " & skp_person_KL

Call query_kk_data_client_loan_app(sql_kk_data_client_loan_app)




'RISK-884
'Data o klientovych û·dostech o hypotÈku
If Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.count = 0 Then
 
 
'Nastaveni listview
 Sonik_Karta_Klienta.ListView_Mortgage_Application.ListItems.Clear
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "Klient", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "»as 2.scoringu", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "»as odesl·nÌ do banky", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "Typ û·dosti", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "»Ìslo û·dosti", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "»Ìslo ˙vÏru", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "⁄vÏrov˝ r·mec", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "Stav û·dosti", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "Kan·l û·dosti", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "IP Adresa", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "Kan·l podpisu - HypoteËnÌ p¯Ìslib", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "Dat. podpisu - HypoteËnÌ p¯Ìslib", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "PoboËka podpisu - HypoteËnÌ p¯Ìslib", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "ID zamÏstnance - HypoteËnÌ p¯Ìslib", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "Kan·l podpisu - HypoteËnÌ smlouva", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "Dat. podpisu - HypoteËnÌ smlouva", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "PoboËka podpisu - HypoteËnÌ smlouva", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "ID zamÏstnance - HypoteËnÌ smlouva", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "D˘vod zamÌtnutÌ", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "Datum kompletace - HypoteËnÌ p¯Ìslib", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "Stav kompletace - HypoteËnÌ p¯Ìslib", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "D˘vod zamÌtnutÌ kompletace - HypoteËnÌ p¯Ìslib", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "Datum kompletace - HypoteËnÌ smlouva", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "Stav kompletace - HypoteËnÌ smlouva", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "D˘vod zamÌtnutÌ kompletace - HypoteËnÌ smlouva", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "Datum  naËerp·nÌ", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "Datum ukonËenÌ", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "Datum storna", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Mortgage_Application.ColumnHeaders.Add(, , "Nesplacen· jistina", 70)



 

End If



'RISK-884
sql_kk_data_client_mort_app = "select JMENO||' '||PRIJMENI||' ('||DATUM_NAROZENI||')' as KLIENT" & _
", CAS2SCORING, CAS_ODESLANI_ZIB " & _
", DRUH_ZADOSTI, CISLO_ZADOSTI, CISLO_UVERU, VYSE_PUJCKY, VYSLEDEK_ZADOSTI, KANAL_ZADOSTI_APPROVE, KANAL_PODPISU_HYPO_PRISLIB, DATUM_PODPISU_HYPO_PRISLIB, POBOCKA_HYPOTECNI_PRISLIB, ID_ZAMESTNANCA_HYP_PRISLIB, KANAL_PODPISU_HYPO_SMLOUVA, DATUM_PODPISU_HYPO_SMLOUVA, POBOCKA_HYPOTECNI_SMLOUVA, ID_ZAMESTNANCA_HYP_SMLOUVA, DUVOD_ZAMITNUTI_CELKOVO, IP_ADRESA, COMPL_FI_1, STAV_KOMPL_UVER_PRISLIB, DUVOD_ZAMITNUTI_CLIENT_KOMPL, COMPL_FI_2, STAV_KOMPL_HYP_SMLOUVA, DUVOD_ZAMITNUTI_CLIENT_HYPSML " & _
",  DATUM_NACERPANI, DATUM_UKONCENI, DATUM_STORNA, NESPLACENA_JISTINA" & _
" from " & schema & ".sonik_apl_data_hypoteka where skp_person = " & skp_person_KL

Call query_kk_data_client_mort_app(sql_kk_data_client_mort_app)



'Data LAP CC - jeötÏ se dohodne, co s nimi

If Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.count = 0 Then
 
'RISK-862
'Nastaveni listview
 Sonik_Karta_Klienta.ListView_Manual_Verification.ListItems.Clear
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "Datum zaloûenÌ û·dosti", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "ZaË·tek provÏrky", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "Konec provÏrky", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "Workflow", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "V˝sledek aktivity", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "Aktivita", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "V˝sledek oper·tor", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "Popisek informace", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "Oper·tor", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "»Ìslo û·dosti", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "ID Ob·lky", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "Hodnota", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, "Further On From AMG", "<- LAP CC / AMG ->", 110)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "Typ û·dosti", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "KÛd UW aktivity", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "Role ûadatele", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "Typ p¯Ìjmu", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "V˝sledek provÏrky ARES", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "Rozsah ovÏ¯enÌ", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "ZamÏstnavatel", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "PracovnÌ vztah", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "PracovnÌ pomÏr", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "Doba urËit· do", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "ProdlouûenÌ DU", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "Datum n·stupu", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "ZkuöebnÌ doba", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "V˝povÏdnÌ lh˘ta", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "Mate¯sk· dovolen·", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "RodiËovsk· dovolen·", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "Dlouhodob· nemoc", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "Sr·ûky, exekuce", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "V˝öe p¯Ìjmu", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "OvÏ¯enÈ I»", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "Mzda vypl·cena", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "Rizikov˝ p¯Ìpad", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Manual_Verification.ColumnHeaders.Add(, , "Stav vyplnÏnÌ", 70)
End If

Call NastavKliceListview(Sonik_Karta_Klienta.ListView_Manual_Verification)

sql_kk_data_client_manual_verification = "SELECT DATUM_ZALOZENI_ZADOSTI, ZACATEK_PROVERKY, KONEC_PROVERKY, WF, VYSLEDEK_AKTIVITY, AKTIVITA, VYSLEDEK_OPERATOR, POPISEK_INFORMACE, OPERATOR, CISLO_ZADOSTI, ID_ENVELOPE, COALESCE(to_char(HODNOTA_DATUM,'DD.MM.YYYY HH24:MI:SS'), to_char(CISELNA_HODNOTA), nullif(TEXTOVA_HODNOTA,'XNA')) AS HODNOTA, TYP_ZADOSTI, NAME_UNDERWRITING_TASK_TYPE, ROLE_ZADATELE, TYP_PRIJMU, VYSLEDEK_PROVERKY_ARES, ROZSAH_OVERENI, ZAMESTNAVATEL, PRACOVNI_VZTAH, PRACOVNI_POMER, DOBA_URCITA_DO, PRODLOUZENI_DU, DATUM_NASTUPU, ZKUSEBNI_DOBA, VYPOVEDNI_LHUTA, MATERSKA_DOVOLENA, RODICOVSKA_DOVOLENA, DLOUHODOBA_NEMOC, SRAZKY_EXEKUCE, VYSE_PRIJMU, OVERENE_IC, MZDA_VYPLACENA, RIZIKOVY_PRIPAD, STAV_VYPLNENI from " & schema & ".sonik_manual_verification where skp_person = " & skp_person_KL

Call query_kk_data_client_manual_verification(sql_kk_data_client_manual_verification)




'Data AMG R8

If Sonik_Karta_Klienta.ListView_AMGR8.ColumnHeaders.count = 0 Then
 
 
'Nastaveni listview
 Sonik_Karta_Klienta.ListView_AMGR8.ListItems.Clear

Set clmx = Sonik_Karta_Klienta.ListView_AMGR8.ColumnHeaders.Add(, , "JmÈno", 70)
Set clmx = Sonik_Karta_Klienta.ListView_AMGR8.ColumnHeaders.Add(, , "P¯ÌjmenÌ", 70)
Set clmx = Sonik_Karta_Klienta.ListView_AMGR8.ColumnHeaders.Add(, , "»as zaloûenÌ û·dosti", 70)
Set clmx = Sonik_Karta_Klienta.ListView_AMGR8.ColumnHeaders.Add(, , "»as odesl·nÌ k posouzenÌ bankou", 70)
Set clmx = Sonik_Karta_Klienta.ListView_AMGR8.ColumnHeaders.Add(, , "ZaË·tek aktivity", 70)
Set clmx = Sonik_Karta_Klienta.ListView_AMGR8.ColumnHeaders.Add(, , "Konec aktivity", 70)
Set clmx = Sonik_Karta_Klienta.ListView_AMGR8.ColumnHeaders.Add(, , "ID û·dosti", 70)
Set clmx = Sonik_Karta_Klienta.ListView_AMGR8.ColumnHeaders.Add(, , "Typ û·dosti", 70)
Set clmx = Sonik_Karta_Klienta.ListView_AMGR8.ColumnHeaders.Add(, , "Typ aktivity", 70)
Set clmx = Sonik_Karta_Klienta.ListView_AMGR8.ColumnHeaders.Add(, , "D˘vod zamÌtnutÌ RISK - internÌ", 70)
Set clmx = Sonik_Karta_Klienta.ListView_AMGR8.ColumnHeaders.Add(, , "D˘vod zamÌtnutÌ pro klienta", 70)
Set clmx = Sonik_Karta_Klienta.ListView_AMGR8.ColumnHeaders.Add(, , "Stav aktivity", 70)


End If


sql_kk_data_client_amg_r8 = "select JMENO, PRIJMENI, CAS_ZALOZENI, CAS_2SCORING, ZACATEK_ACTIVITY, KONEC_ACTIVITY, ID_ZADOSTI, NAME_APPLICATION_TYPE, TASK_QUEUE, REJECT_REASON, REJECT_REASON_CLIENT, STATUS_PROVERKY from " & schema & ".sonik_verifikace_amg where skp_person = " & skp_person_KL

Call query_kk_data_client_amg_r8(sql_kk_data_client_amg_r8)







'Data AMG Note

If Sonik_Karta_Klienta.ListView_AMG_Note.ColumnHeaders.count = 0 Then
 
 
'Nastaveni listview
 Sonik_Karta_Klienta.ListView_AMG_Note.ListItems.Clear

Set clmx = Sonik_Karta_Klienta.ListView_AMG_Note.ColumnHeaders.Add(, , "JmÈno", 70)
Set clmx = Sonik_Karta_Klienta.ListView_AMG_Note.ColumnHeaders.Add(, , "P¯ÌjmenÌ", 70)
Set clmx = Sonik_Karta_Klienta.ListView_AMG_Note.ColumnHeaders.Add(, , "ID Envelope", 70)
Set clmx = Sonik_Karta_Klienta.ListView_AMG_Note.ColumnHeaders.Add(, , "Datum pozn·mky", 70)
Set clmx = Sonik_Karta_Klienta.ListView_AMG_Note.ColumnHeaders.Add(, , "Pozn·mka", 70)
Set clmx = Sonik_Karta_Klienta.ListView_AMG_Note.ColumnHeaders.Add(, , "Priv·tnÌ riskov· pozn·mka", 70)



End If


sql_kk_data_client_amg_note = "select JMENO, PRIJMENI, ID_APP_ENVELOPE, MODIF_TIME, NOTE, NOTE_PRIVATE from " & schema & ".sonik_verifikace_amg_note where skp_person = " & skp_person_KL

Call query_kk_data_client_amg_note(sql_kk_data_client_amg_note)





'Data AMG Calls_R8

If Sonik_Karta_Klienta.ListView_Calls_R8.ColumnHeaders.count = 0 Then
 
 
'Nastaveni listview
 Sonik_Karta_Klienta.ListView_Calls_R8.ListItems.Clear

Set clmx = Sonik_Karta_Klienta.ListView_Calls_R8.ColumnHeaders.Add(, , "Grid KontaktnÌ historie (OBS): SystÈmovÈ ID", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Calls_R8.ColumnHeaders.Add(, , "⁄Ëel kontaktu", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Calls_R8.ColumnHeaders.Add(, , "SmÏr hovoru", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Calls_R8.ColumnHeaders.Add(, , "»as zaloûenÌ kontaktu", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Calls_R8.ColumnHeaders.Add(, , "P¯edvolba", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Calls_R8.ColumnHeaders.Add(, , "VolanÈ tel. ËÌslo", 70)
Set clmx = Sonik_Karta_Klienta.ListView_Calls_R8.ColumnHeaders.Add(, , "Typ kontaktu", 70)




End If


sql_kk_data_client_amg_calls = "select IDENTIFIKATOR_OBS, UCEL_KONTAKTU, SMER_HOVORU, CAS_ZALOZENI_CONTACTU, PREDVOLBA, VOLANE_TELEFONNI_CISLO, TYP_CONTACTU from " & schema & ".sonik_amg_calls where skp_person = " & skp_person_KL

Call query_kk_data_client_amg_calls(sql_kk_data_client_amg_calls)




'Data BRKI Dotazy

If Sonik_Karta_Klienta.ListView_BRKI.ColumnHeaders.count = 0 Then
 
 
'Nastaveni listview
 Sonik_Karta_Klienta.ListView_BRKI.ListItems.Clear

Set clmx = Sonik_Karta_Klienta.ListView_BRKI.ColumnHeaders.Add(, , "JmÈno", 70)
Set clmx = Sonik_Karta_Klienta.ListView_BRKI.ColumnHeaders.Add(, , "P¯ÌjmenÌ", 70)
Set clmx = Sonik_Karta_Klienta.ListView_BRKI.ColumnHeaders.Add(, , "Datum narozenÌ", 70)
Set clmx = Sonik_Karta_Klienta.ListView_BRKI.ColumnHeaders.Add(, , "Datum a Ëas", 70)
Set clmx = Sonik_Karta_Klienta.ListView_BRKI.ColumnHeaders.Add(, , "ID û·dosti", 70)
Set clmx = Sonik_Karta_Klienta.ListView_BRKI.ColumnHeaders.Add(, , "V˝sledek dotazu", 70)
Set clmx = Sonik_Karta_Klienta.ListView_BRKI.ColumnHeaders.Add(, , "Popis chybovÈ hl·öky", 70)





End If


sql_kk_data_client_brki = "select NAME_FIRST, NAME_LAST, DATE_BIRTH, REQUEST_TIME, ID_ZADOSTI, RESPONSE_ERROR_VYSVETLENI, ERRORDESCRIPTION from " & schema & ".sonik_brki_dotazy where skp_person = " & skp_person_KL

Call query_kk_data_client_brki(sql_kk_data_client_brki)




'Komunikace s klientem

If Sonik_Karta_Klienta.ListView_Client_Communication.ColumnHeaders.count = 0 Then
 
 
'Nastaveni listview
 Sonik_Karta_Klienta.ListView_Client_Communication.ListItems.Clear
 
'RISK-864
Set clmx = Sonik_Karta_Klienta.ListView_Client_Communication.ColumnHeaders.Add(, , "Kan·l komunikace", 90)
Set clmx = Sonik_Karta_Klienta.ListView_Client_Communication.ColumnHeaders.Add(, , "Datum odesl·nÌ zpr·vy", 110)
Set clmx = Sonik_Karta_Klienta.ListView_Client_Communication.ColumnHeaders.Add(, , "Odesl·no na", 90)
Set clmx = Sonik_Karta_Klienta.ListView_Client_Communication.ColumnHeaders.Add(, , "HlaviËka mailu", 100)
Set clmx = Sonik_Karta_Klienta.ListView_Client_Communication.ColumnHeaders.Add(, , "Zpr·va", 500)
Set clmx = Sonik_Karta_Klienta.ListView_Client_Communication.ColumnHeaders.Add(, , "JmÈno zamÏstnance", 90)
Set clmx = Sonik_Karta_Klienta.ListView_Client_Communication.ColumnHeaders.Add(, , "OddÏlenÌ", 90)
Set clmx = Sonik_Karta_Klienta.ListView_Client_Communication.ColumnHeaders.Add(, "ID_CONTACT", "SystÈmovÈ ID", 90)

End If


Call NastavKliceListview(Sonik_Karta_Klienta.ListView_Client_Communication)


sql_kk_data_client_communication = "select NAME_COMMON_CHANNEL, DTIME_SEND ,nvl(nullif(TEXT_PHONE_NUMBER,'XNA'),nullif(TEXT_EMAIL_ADDRESS,'XNA')) as TEXT_RECEIVER, nullif(TEXT_HEADER,'XNA') as TEXT_HEADER, TEXT_MESSAGE, NAME_EMPLOYEE, CODE_EMPLOYEE_DEPARTMENT, ID_CONTACT from " & schema & ".sonik_sms_ib_email where skp_person = " & skp_person_KL & "order by DTIME_SEND desc"

Call query_kk_data_client_communication(sql_kk_data_client_communication)




'RISK-868

If Sonik_Karta_Klienta.ListView_Transactions_Homer.ColumnHeaders.count = 0 Then
 
 
'Nastaveni listview
 Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems.Clear

Set clmx = Sonik_Karta_Klienta.ListView_Transactions_Homer.ColumnHeaders.Add(, , "»Ìslo karty", 130)
Set clmx = Sonik_Karta_Klienta.ListView_Transactions_Homer.ColumnHeaders.Add(, , "Datum a Ëas", 90)
Set clmx = Sonik_Karta_Klienta.ListView_Transactions_Homer.ColumnHeaders.Add(, , "Platba", 90)
Set clmx = Sonik_Karta_Klienta.ListView_Transactions_Homer.ColumnHeaders.Add(, , "Stav", 90)
Set clmx = Sonik_Karta_Klienta.ListView_Transactions_Homer.ColumnHeaders.Add(, , "»·stka v CZK", 90) 'RISK-1239
'Set clmx = Sonik_Karta_Klienta.ListView_Transactions_Homer.ColumnHeaders.Add(, , "MÏna", 90) 'RISK-1239
Set clmx = Sonik_Karta_Klienta.ListView_Transactions_Homer.ColumnHeaders.Add(, , "ObchodnÌk", 130)
Set clmx = Sonik_Karta_Klienta.ListView_Transactions_Homer.ColumnHeaders.Add(, , "Adresa", 130)
Set clmx = Sonik_Karta_Klienta.ListView_Transactions_Homer.ColumnHeaders.Add(, , "MÏsto", 90)
Set clmx = Sonik_Karta_Klienta.ListView_Transactions_Homer.ColumnHeaders.Add(, , "ZemÏ", 50)
Set clmx = Sonik_Karta_Klienta.ListView_Transactions_Homer.ColumnHeaders.Add(, , "St·t", 50)
Set clmx = Sonik_Karta_Klienta.ListView_Transactions_Homer.ColumnHeaders.Add(, , "ID Termin·lu", 80)
Set clmx = Sonik_Karta_Klienta.ListView_Transactions_Homer.ColumnHeaders.Add(, , "MCC", 50) 'RISK-999



End If


'RISK-999 'RISK-1239
sql_kk_data_transactions_homer_agg = "select CARDNUMBER, DTIME_AUTHORIZATION, PLATBA, ODEZVA_AUTORIZACE, CASTKA_V_CZK, OBCHODNI_MISTO, TEXT_CARD_ACCEPTOR_STREET, MESTO, ZEME, STAT, TERMINAL, MCC from " & schema & ".sonik_karty_trans_homer_agg where skp_person = " & skp_person_KL & " order by DTIME_AUTHORIZATION"

Call query_kk_data_transactions_homer_Agg(sql_kk_data_transactions_homer_agg)




'RISK-884

If Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.count = 0 Then
 
 
'Nastaveni listview
 Sonik_Karta_Klienta.ListView_Application_Form.ListItems.Clear
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "»as 2. scoringu", 70) '
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "»as odesl·nÌ do banky", 70) '
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "Typ û·dosti", 70) '
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "»Ìslo û·dosti", 70) '
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "»Ìslo ˙vÏru", 70) '
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "Poûadovan˝ ˙vÏrov˝ r·mec", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "Stav û·dosti", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "Typ bydlenÌ / Od", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "VzdÏl·nÌ", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "Rodinn˝ stav", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "PoËet vyûivovan˝ch osob", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "Spl·tky uvedenÈ klientem (Suma spl·tek / v˝öe r·mc˘)", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "V˝daje dom·cnosti (BydlenÌ / lÈky-jÌdlo-doprava / spl·tky ostatnÌch Ëlen˘)", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "P¯Ìjmy ostat. Ëlen˘ dom·cnosti", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "ZamÏstnavatel (hl. p¯Ìjem)", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "ZamÏstn·n od", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "PracovnÌ pomÏr", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "V˝öe hl. p¯Ìjmu", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "Profese hl. p¯Ìjmu", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "Adresa zamÏstnavatele - hl. p¯Ìjem", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "ZamÏstnavatel (vedl. p¯Ìjem)", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "ZamÏstn·n od (vedl. p¯Ìjem)", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "PracovnÌ pomÏr (vedl. p¯Ìjem)", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "V˝öe vedl. p¯Ìjmu", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "Profese vedl. p¯Ìjmu", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "Adresa zamÏstnavatele - vedl. p¯Ìjem", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "Prim·rnÌ telefon", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "Druh˝ telefon", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "Telefon do hl. zamÏstn·nÌ", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "Telefon do vedl. zamÏstn·nÌ", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "Prim·rnÌ e-mail", 70)
    Set clmx = Sonik_Karta_Klienta.ListView_Application_Form.ColumnHeaders.Add(, , "Preferovan˝ typ druhÈho dokladu", 70)
End If
Call NastavKliceListview(Sonik_Karta_Klienta.ListView_Application_Form)

'RISK-884
sql_kk_data_application_form_data = "select dtime_scoring, dtime_offer_sent_from_ib, name_application_type, text_application_number, code_loan_number, creditamount, name_application_status " & _
" , name_housing_type, date_residence_from, name_education, name_marital_status, num_dependent_person, amt_loan_payment, amt_credits_sum " & _
" , amt_accommodation_exp, amt_med_food_trans_exp, amt_other_house_mem_inst_exp, amt_other_house_mem_income, ico_main, name_employer_main, dtime_employed_from_main, dtime_agreement_until_main, amt_net_income_main, name_economical_status_main, name_employment_type_main, name_job_profession_main " & _
" , name_job_industry_main, name_country_main, name_district_main, name_locality_main, name_town_main, name_street_main, text_house_num_main, code_zip_main, ico_add, name_employer_add, dtime_employed_from_add, dtime_agreement_until_add, amt_net_income_add, name_economical_status_add " & _
" , name_employment_type_add, name_job_profession_add, name_job_industry_add, name_country_add, name_district_add, name_locality_add, name_town_add, name_street_add, text_house_num_add, code_zip_add, primary_phone_extension, primary_phone, other_phone_extension, other_phone, main_empl_phone_extension " & _
" , main_empl_phone, add_empl_phone_extension, add_empl_phone, primary_email, second_id_type_prefered " & _
" from " & schema & ".sonik_app_form_data where skp_person = " & skp_person_KL & " order by dtime_inserted"

Call query_kk_data_application_form_data(sql_kk_data_application_form_data)


End Sub









Sub query_dz_history(ByVal sql As String)

   Sonik_Datova_Zprava.ListView_DZ_History.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Datova_Zprava.ListView_DZ_History.ListItems.Add()
      itmX.text = "" & rst1![DATA_MESSAGE_STATUS_ID] & ""
      itmX.SubItems(1) = "" & rst1![DATA_MESSAGE_NOTE_ID] & ""
      itmX.SubItems(2) = "" & rst1![DATA_MESSAGE_NOTICE_ID] & ""
      itmX.SubItems(3) = "" & rst1![DTIME_INSERTED] & ""
      itmX.SubItems(4) = "" & rst1![EVENT_TYPE] & ""
      itmX.SubItems(5) = "" & rst1![Event] & ""
      itmX.SubItems(6) = "" & rst1![name_employee_inserted] & ""
      If rst1![ADDITIONAL_NOTE] <> vbNullString Then
      itmX.SubItems(7) = "Ano" 'Flag, jestli existuje dodateËn· pozn·mka k zobrazenÌ
      Else: itmX.SubItems(7) = "Ne"
      End If
      itmX.SubItems(8) = "" & rst1![ADDITIONAL_NOTE] & ""
    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub


Sub query_dz_notice(ByVal sql As String)

   Sonik_Datova_Zprava.ListView_DZ_Upozorneni.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Datova_Zprava.ListView_DZ_Upozorneni.ListItems.Add()
      itmX.text = "" & rst1![skp_data_message_notice] & ""
      itmX.SubItems(1) = "" & rst1![text_notice] & ""
      itmX.SubItems(2) = "" & rst1![dtime_notice_deadline] & ""
      itmX.SubItems(3) = "" & Round(rst1![dtime_notice_deadline] - date) & ""

    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub


'Sub query_ciselnik_pozad_info(ByVal sql As String)
''NaËtenÌ ËÌslenÌku moûn˝ch vÏcÌ, na co se policie pt·
'
'   Sonik_Datova_Zprava.ListView_pozad_info.ListItems.Clear
'
'   Set rst1 = New ADODB.Recordset
'   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'
'   If Not rst1.EOF Then
'    rst1.MoveFirst
'    For I = 1 To rst1.RecordCount
'      Set itmX = Sonik_Datova_Zprava.ListView_pozad_info.ListItems.Add()
'      itmX.text = "" & rst1![SKP_REQ_INFO] & ""
'      itmX.SubItems(1) = "" & rst1![NAME_REQ_INFO] & ""
'    rst1.MoveNext
'    Next I
'   End If
'
'   rst1.Close
'
'End Sub


Sub query_pozad_info(ByVal sql As String)
'NaËtenÌ vÏcÌ, na co se policie pt·
   
   Sonik_Datova_Zprava.ListView_pozad_info.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Datova_Zprava.ListView_pozad_info.ListItems.Add()
      itmX.text = "" & rst1![id] & ""
      itmX.SubItems(1) = "" & rst1![name_req_prov_info] & ""
      itmX.SubItems(2) = "" & rst1![DTIME_INSERTED] & ""
      itmX.SubItems(3) = "" & rst1![name_employee_inserted] & ""
    rst1.MoveNext
    Next i
   End If

   rst1.Close

End Sub



Sub query_poskyt_info(ByVal sql As String)
'NaËtenÌ vÏcÌ, kterÈ poskytujeme
   
   Sonik_Datova_Zprava.ListView_Poskyt_Info.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Datova_Zprava.ListView_Poskyt_Info.ListItems.Add()
      itmX.text = "" & rst1![id] & ""
      itmX.SubItems(1) = "" & rst1![name_req_prov_info] & ""
      itmX.SubItems(2) = "" & rst1![DTIME_INSERTED] & ""
      itmX.SubItems(3) = "" & rst1![name_employee_inserted] & ""
    rst1.MoveNext
    Next i
   End If

   rst1.Close

End Sub



Sub query_datove_zpravy(ByVal sql As String)
   
   Sonik_Datove_Zpravy_Prehled.ListView_Datove_Zpravy.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Datove_Zpravy_Prehled.ListView_Datove_Zpravy.ListItems.Add()
      itmX.text = "" & rst1![SKP_DATA_MESSAGE] & ""
      itmX.SubItems(1) = "" & rst1![data_message_number] & ""
      itmX.SubItems(2) = "" & rst1![date_data_message_delivery] & ""
      itmX.SubItems(3) = "" & rst1![data_message_sender] & ""
      itmX.SubItems(4) = "" & rst1![text_file_number] & ""
      itmX.SubItems(5) = "" & rst1![text_reference_number] & ""
      itmX.SubItems(6) = "" & rst1![DTIME_INSERTED] & ""
      itmX.SubItems(7) = "" & rst1![name_employee_inserted] & ""
      itmX.SubItems(8) = "" & rst1![NAME_DATA_MESSAGE_STATUS] & ""
    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub



Sub query_confirmed_pairs(ByVal sql As String)
   
'RISK-876
Dim keys As Dictionary
Set keys = KliceKListview(Sonik_Datova_Zprava.ListView_Paired_Subjects)
   
   Sonik_Datova_Zprava.ListView_Paired_Subjects.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    'RISK-876
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Datova_Zprava.ListView_Paired_Subjects.ListItems.Add()
      itmX.text = "" & rst1![SKP_CONFIRM_PAIR] & ""
      itmX.SubItems(keys("ID Subjektu")) = "" & rst1![SKP_Subject] & ""
      itmX.SubItems(keys("JmÈno LU")) = "" & rst1![NAME_FIRST_lu] & ""
      itmX.SubItems(keys("P¯ÌjmenÌ LU")) = "" & rst1![NAME_LAST_lu] & ""
      itmX.SubItems(keys("Datum narozenÌ LU")) = "" & rst1![DATE_BIRTH_lu] & ""
      itmX.SubItems(keys("RodnÈ ËÌslo LU")) = "" & rst1![TEXT_BIRTH_NUMBER_lu] & ""
      itmX.SubItems(keys("»Ìslo ˙Ëtu LU")) = "" & rst1![Text_bank_account_number_lu] & ""
      itmX.SubItems(keys("»Ìslo karty LU")) = "" & rst1![TEXT_CARD_NUMBER_lu] & ""
      itmX.SubItems(keys("»Ìslo ˙vÏru LU")) = "" & rst1![CODE_LOAN_NUMBER_lu] & ""
      itmX.SubItems(keys("SKP_PERSON")) = "" & rst1![skp_person] & ""
      itmX.SubItems(keys("CUID")) = "" & rst1![id_cuid] & ""
      itmX.SubItems(keys("Klient")) = "" & rst1![Person_Name] & ""
      itmX.SubItems(keys("Datum narozenÌ")) = "" & rst1![Date_Birth] & ""
      itmX.SubItems(keys("RodnÈ ËÌslo")) = "" & rst1![TEXT_BIRTH_NUMBER] & ""
      itmX.SubItems(keys("Datum potvrzenÌ")) = "" & rst1![DTIME_INSERTED] & ""
      itmX.SubItems(keys("Potvrzeno")) = "" & rst1![name_employee_inserted] & ""
      itmX.SubItems(keys("»Ìslo telefonu LU")) = "" & rst1![text_phone_number_lu] & ""
    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub


Sub query_negativky(ByVal sql As String)
   
'RISK-876
Dim keys As Dictionary
Set keys = KliceKListview(Sonik_Datova_Zprava.ListView_Negativky)
   
   Sonik_Datova_Zprava.ListView_Negativky.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    'RISK-876
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Datova_Zprava.ListView_Negativky.ListItems.Add()
      itmX.text = "" & rst1![SKP_CONFIRM_PAIR] & ""
      itmX.SubItems(keys("ID Subjektu")) = "" & rst1![SKP_Subject] & ""
      itmX.SubItems(keys("JmÈno")) = "" & rst1![Name_first] & ""
      itmX.SubItems(keys("P¯ÌjmenÌ")) = "" & rst1![Name_last] & ""
      itmX.SubItems(keys("Datum narozenÌ")) = "" & rst1![Date_Birth] & ""
      itmX.SubItems(keys("RodnÈ ËÌslo")) = "" & rst1![TEXT_BIRTH_NUMBER] & ""
      itmX.SubItems(keys("»Ìslo ˙Ëtu")) = "" & rst1![Text_bank_account_number] & ""
      itmX.SubItems(keys("»Ìslo karty")) = "" & rst1![TEXT_CARD_NUMBER] & ""
      itmX.SubItems(keys("»Ìslo ˙vÏru")) = "" & rst1![code_loan_number] & ""
      itmX.SubItems(keys("Datum potvrzenÌ")) = "" & rst1![DTIME_INSERTED] & ""
      itmX.SubItems(keys("Potvrzeno")) = "" & rst1![name_employee_inserted] & ""
      itmX.SubItems(keys("»Ìslo telefonu")) = "" & rst1![TEXT_PHONE_NUMBER] & ""
    rst1.MoveNext
    Next i
   End If

   rst1.Close





End Sub



'RISK-813
Sub query_hledani_subjektu_v_dtb(ByVal sql As String)
   
Dim i As Integer, j As Integer, k As Integer, Bln_Show_Error_Message As Boolean, Str_Hlaska As String
'RISK-876
Dim keys As Dictionary
Set keys = KliceKListview(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni)
   
  Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   'RISK-876
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ListItems.Add()
      itmX.text = SKP_Subject 'Vezmu p¯Ìmo z excelu
      itmX.SubItems(keys("ID_CUID")) = "" & rst1![id_cuid] & ""
      itmX.SubItems(keys("JmÈno")) = "" & rst1![Name_first] & ""
      itmX.SubItems(keys("P¯ÌjmenÌ")) = "" & rst1![Name_last] & ""
      itmX.SubItems(keys("Datum narozenÌ")) = "" & rst1![Date_Birth] & ""
      itmX.SubItems(keys("RodnÈ ËÌslo")) = "" & rst1![TEXT_BIRTH_NUMBER] & ""
      itmX.SubItems(keys("Trval· adresa")) = "" & rst1![PERMANENT_ADDRESS] & ""
      itmX.SubItems(keys("»Ìsla ˙Ët˘")) = "" & rst1![TEXT_BANK_ACCOUNT_NUMBER_AGG] & ""
      itmX.SubItems(keys("»Ìsla ˙vÏr˘")) = "" & rst1![CODE_LOAN_NUMBER_AGG] & ""
      itmX.SubItems(keys("PotvrzenÌ")) = "Nepotvrzeno"
      itmX.SubItems(keys("skp_person")) = "" & rst1![skp_person] & ""
      itmX.SubItems(keys("»Ìsla telefon˘")) = "" & rst1![TEXT_PHONE_NUMBER_AGG] & ""
    rst1.MoveNext
    Next i
   End If

   rst1.Close


Bln_Show_Error_Message = False
'RISK-813
'Projdu sestavu vöech jiû potvrzen˝ch subjekt˘ na datovÈ zpr·vÏ a srovn·m s pr·vÏ vyhled·van˝mi klienty
'Pokud jiû nÏkter˝ hledan˝ klient ve zpr·vÏ potvrzen je, tak podbarvÌm ¯·dek ËervenÏ, d·m implicitnÏ na zamÌtnuto a vypÌöu hl·öku uûivateli
    For i = 1 To Sonik_Datova_Zprava.ListView_Paired_Subjects.ListItems.count
        For j = 1 To Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ListItems.count
            'RISK-876
            
            If ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Paired_Subjects, "SKP_PERSON", i) = ListvZiskejHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "skp_person", j) Then
            
                For k = 1 To Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ColumnHeaders.count - 1
                    Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ListItems(j).ListSubItems.item(k).ForeColor = barvaRudaSvetla 'Podbarvi
                    Call ListvVlozHodnotu(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni, "PotvrzenÌ", "ZamÌtnuto", j) 'RISK-876
                    Bln_Show_Error_Message = True
                Next k
            
            End If
            
        Next j
    Next i
      

If Bln_Show_Error_Message Then

    Str_Hlaska = "NÏkter˝ z vyhled·van˝ch klient˘ byl jiû v tÈto zpr·vÏ jednou p¯id·n (ËervenÏ oznaËeno)"
    Call HlaskaUzivateli(Str_Hlaska)
    Str_Hlaska = "ZamÌtni jej prosÌm aù se tam neobjevÌ dvakr·t"
    Call HlaskaUzivateli(Str_Hlaska)
        
    Call VypisVsechnyHlasky
End If


End Sub




Sub query_historie_stavu_dz(ByVal sql As String)
   
   Sonik_Update_Stav_DZ.ListView_Historie_Stavu.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Update_Stav_DZ.ListView_Historie_Stavu.ListItems.Add()
      itmX.text = "" & rst1![id] & ""
      itmX.SubItems(1) = "" & rst1![NAME_DATA_MESSAGE_STATUS] & ""
      itmX.SubItems(2) = "" & rst1![TEXT_NOTE] & ""
      itmX.SubItems(3) = "" & rst1![DTIME_VALID_FROM] & ""
      itmX.SubItems(4) = "" & rst1![DTIME_VALID_TO] & ""
      itmX.SubItems(5) = "" & rst1![name_employee_inserted] & ""
    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub





Sub query_cl_Provided_Info_dz(ByVal sql As String)
   
   Sonik_Rucni_Pridani_Info_DZ.ListView_Provided_Info.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Rucni_Pridani_Info_DZ.ListView_Provided_Info.ListItems.Add()
      itmX.text = "" & rst1![name_req_prov_info] & ""
    rst1.MoveNext
    Next i
   End If

   rst1.Close

End Sub


Sub query_cl_Req_Info_dz(ByVal sql As String)
   
   Sonik_Rucni_Pridani_Info_DZ.ListView_Req_Info.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Rucni_Pridani_Info_DZ.ListView_Req_Info.ListItems.Add()
      itmX.text = "" & rst1![name_req_prov_info] & ""
    rst1.MoveNext
    Next i
   End If

   rst1.Close

End Sub




Sub query_kk_client_soucinnosti(ByVal sql As String)
   
   Sonik_Karta_Klienta.ListView_Soucinnosti_Klienta.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_Soucinnosti_Klienta.ListItems.Add()
      itmX.text = "" & rst1![SKP_DATA_MESSAGE] & ""
      itmX.SubItems(1) = "" & rst1![data_message_number] & ""
      itmX.SubItems(2) = SonikFormatDate("" & rst1![date_data_message_delivery] & "") 'RISK-1131
      itmX.SubItems(3) = "" & rst1![data_message_sender] & ""
      itmX.SubItems(4) = "" & rst1![text_file_number] & ""
      itmX.SubItems(5) = "" & rst1![text_reference_number] & ""
      itmX.SubItems(6) = "" & rst1![NAME_DATA_MESSAGE_STATUS] & ""
      itmX.SubItems(7) = "" & rst1![SKP_DAMAGE] & ""
      itmX.SubItems(8) = "" & rst1![amt_damage] & ""
      itmX.SubItems(9) = "" & rst1![TEXT_NOTE] & ""
      
    rst1.MoveNext
    Next i
   End If

   rst1.Close



End Sub



Sub query_kk_data_client_actual(ByVal sql As String)
   
   Sonik_Karta_Klienta.ListView_Client_Actual.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_Client_Actual.ListItems.Add()
      itmX.text = "" & rst1![Name_first] & ""
      itmX.SubItems(1) = "" & rst1![Name_last] & ""
      itmX.SubItems(2) = "" & rst1![Text_Title_before] & ""
      itmX.SubItems(3) = "" & rst1![Text_Title_After] & ""
      itmX.SubItems(4) = "" & rst1![TEXT_BIRTH_NUMBER] & ""
      itmX.SubItems(5) = SonikFormatDate("" & rst1![Date_Birth] & "") 'RISK-1131
      'itmX.SubItems(6) = "" & rst1![NAME_NATIONALITY] & ""
      itmX.SubItems(6) = "" & rst1![TRVALA_ADRESA] & ""
      itmX.SubItems(7) = "" & rst1![KORESPONDENCNI_ADRESA] & ""
      itmX.SubItems(8) = "" & rst1![text_mobile_phone] & ""
      itmX.SubItems(9) = "" & rst1![text_email_address] & ""
      itmX.SubItems(10) = "" & rst1![NAME_GENDER] & ""
      itmX.SubItems(11) = SonikFormatDate("" & rst1![DATUM_UMRTI] & "") 'RISK-1131
      
      
      
    rst1.MoveNext
    Next i
   End If

   rst1.Close



End Sub




Sub query_kk_data_client_gc(ByVal sql As String)
   
   Sonik_Karta_Klienta.ListView_Client_GC.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_Client_GC.ListItems.Add()
      itmX.text = "" & rst1![NAME_CONTRACT_STATUS] & "" & " / " & rst1![NAME_CONTRACT_WFSTATUS] & ""
      itmX.SubItems(1) = "" & rst1![DTIME_CONTRACT_SIGNATURE] & ""
      itmX.SubItems(2) = "" & rst1![text_contract_number] & ""
      itmX.SubItems(3) = "" & rst1![PLATNOST_OD] & ""
      itmX.SubItems(4) = "" & rst1![PLATNOST_DO] & ""
      
      
    rst1.MoveNext
    Next i
   End If

   rst1.Close



End Sub



Sub query_kk_data_client_Birth_Number_Change(ByVal sql As String)
   
   Sonik_Karta_Klienta.ListView_Birth_Number_Change.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_Birth_Number_Change.ListItems.Add()
      itmX.text = "" & rst1![TEXT_ATTRIBUTE_VALUE] & ""
      itmX.SubItems(1) = "" & rst1![PLATNOST_OD] & ""
      itmX.SubItems(2) = "" & rst1![PLATNOST_DO] & ""
      
      
    rst1.MoveNext
    Next i
   End If

   rst1.Close

End Sub



Sub query_kk_data_client_Name_Change(ByVal sql As String)
   
   Sonik_Karta_Klienta.ListView_Name_Change.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_Name_Change.ListItems.Add()
      itmX.text = "" & rst1![Name_first] & ""
      itmX.SubItems(1) = "" & rst1![Name_last] & ""
      itmX.SubItems(2) = "" & rst1![Text_Title_before] & ""
      itmX.SubItems(3) = "" & rst1![Text_Title_After] & ""
      itmX.SubItems(4) = SonikFormatDate("" & rst1![Date_Birth] & "") 'RISK-1131
      itmX.SubItems(5) = "" & rst1![NAME_COUNTRY] & ""
      itmX.SubItems(6) = "" & rst1![CODE_MODIFIED_BY] & ""
      itmX.SubItems(7) = "" & rst1![PLATNOST_OD] & ""
      itmX.SubItems(8) = "" & rst1![PLATNOST_DO] & ""
      itmX.SubItems(9) = "" & rst1![FLAG_CURRENT] & ""
      
      
    rst1.MoveNext
    Next i
   End If

   rst1.Close

End Sub


Sub query_kk_data_client_Change_Mail_Phone(ByVal sql As String)
   
   Sonik_Karta_Klienta.ListView_Change_Mail_Phone.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_Change_Mail_Phone.ListItems.Add()
      itmX.text = "" & rst1![Name_first] & ""
      itmX.SubItems(1) = "" & rst1![Name_last] & ""
      itmX.SubItems(2) = "" & rst1![CODE_CONTACT_TYPE] & ""
      itmX.SubItems(3) = "" & rst1![TEXT_PERSON_CONTACT] & ""
      itmX.SubItems(4) = "" & rst1![CODE_CONTACT_STATUS] & ""
      itmX.SubItems(5) = "" & rst1![AKTUALIZACE] & ""
          
    rst1.MoveNext
    Next i
   End If

   rst1.Close

End Sub


Sub query_kk_data_client_Change_Address(ByVal sql As String)
   
   Sonik_Karta_Klienta.ListView_Address_Change.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_Address_Change.ListItems.Add()
      itmX.text = "" & rst1![jmeno] & ""
      itmX.SubItems(1) = "" & rst1![PRIJMENI] & ""
      itmX.SubItems(2) = "" & rst1![NAME_ADDRESS_TYPE] & ""
      itmX.SubItems(3) = "" & rst1![ULICE] & ""
      itmX.SubItems(4) = "" & rst1![TEXT_HOUSE_NUM] & ""
      itmX.SubItems(5) = "" & rst1![CODE_ZIP] & ""
      itmX.SubItems(6) = "" & rst1![NAME_TOWN] & ""
      itmX.SubItems(7) = "" & rst1![NAME_COUNTRY] & ""
      itmX.SubItems(8) = "" & rst1![CAS_ZMENY] & ""
      itmX.SubItems(9) = "" & rst1![FLAG_CURRENT] & ""
          
    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub



Sub query_kk_data_client_login_ib(ByVal sql As String)
   
   Sonik_Karta_Klienta.ListView_Login_IB.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_Login_IB.ListItems.Add()
      itmX.text = "" & rst1![Name_first] & ""
      itmX.SubItems(1) = "" & rst1![Name_last] & ""
      itmX.SubItems(2) = "" & rst1![TEXT_USER_NAME] & ""
      itmX.SubItems(3) = "" & rst1![IP_ADDRESS] & ""
      itmX.SubItems(4) = "" & rst1![DTIME_CREATED] & ""
      itmX.SubItems(5) = "" & rst1![DTIME_LOGGED_OUT] & ""
      itmX.SubItems(6) = "" & rst1![NAME_COMMON_CHANNEL] & ""
      itmX.SubItems(7) = "" & rst1![FLAG_REQUEST_SUCCESSFUL] & ""
      itmX.SubItems(8) = "" & rst1![CODE_OPERATION] & ""
      itmX.SubItems(9) = "" & rst1![CODE_INFO] & ""
          
          
    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub



Sub query_kk_data_client_accounts(ByVal sql As String)
   
   Sonik_Karta_Klienta.ListView_Accounts.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   'RISK-992 odmaz·nÌ jiû nepot¯ebn˝ch sloupc˘
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_Accounts.ListItems.Add()
      itmX.text = "" & rst1![bank_account_type] & ""
      itmX.SubItems(1) = "" & rst1![Code_currency] & ""
      itmX.SubItems(2) = "" & rst1![Text_bank_account_number] & ""
      itmX.SubItems(3) = "" & rst1![NAME_BANK_ACCOUNT_STATUS] & ""
      itmX.SubItems(4) = "" & rst1![DATE_BANK_ACCOUNT_ACTIVATION] & ""
      itmX.SubItems(5) = "" & rst1![dtime_of_application] & ""
      itmX.SubItems(6) = "" & rst1![APP_CHANNEL_TYPE_REQUEST] & ""
      itmX.SubItems(7) = SonikFormatDate("" & rst1![Dtime_signature] & "") 'RISK-1131
      itmX.SubItems(8) = "" & rst1![signature_channel] & ""
      itmX.SubItems(9) = "" & rst1![SIGNATURE_TYPE] & ""
      itmX.SubItems(10) = "" & rst1![BRANCH_SIGN] & ""
      itmX.SubItems(11) = "" & rst1![NAME_PERSON_SIGN] & ""
      itmX.SubItems(12) = "" & rst1![ucetni_zustatek] & ""
      itmX.SubItems(13) = "" & rst1![DISP_ZUSTATEK] & ""
      itmX.SubItems(14) = "" & rst1![IP_ADDRESS_APPR_SIGN] & ""
      itmX.SubItems(15) = "" & rst1![DATE_BANK_ACCOUNT_TERMINATION] & ""
      
          
    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub



Sub query_kk_data_disponent_rights(ByVal sql As String)
   
'RISK-992
Dim keys As Dictionary
Set keys = KliceKListview(Sonik_Karta_Klienta.ListView_Disponent_Rights)
   
   Sonik_Karta_Klienta.ListView_Disponent_Rights.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_Disponent_Rights.ListItems.Add()
      itmX.text = "" & rst1![PERSON_ACCOUNT_REL] & ""
      itmX.SubItems(keys("CUID_OWNER")) = "" & rst1![CUID_OWNER] & ""
      itmX.SubItems(keys("NAME_FIRST_OWNER")) = "" & rst1![NAME_FIRST_OWNER] & ""
      itmX.SubItems(keys("NAME_LAST_OWNER")) = "" & rst1![NAME_LAST_OWNER] & ""
      itmX.SubItems(keys("CUID_DISPONENT")) = "" & rst1![CUID_DISPONENT] & ""
      itmX.SubItems(keys("NAME_FIRST_DISPONENT")) = "" & rst1![NAME_FIRST_DISPONENT] & ""
      itmX.SubItems(keys("NAME_LAST_DISPONENT")) = "" & rst1![NAME_LAST_DISPONENT] & ""
      itmX.SubItems(keys("CODE_RIGHT_TYPE")) = "" & rst1![code_right_type] & ""
      itmX.SubItems(keys("TEXT_BANK_ACCOUNT_NUMBER")) = "" & rst1![Text_bank_account_number] & ""
      itmX.SubItems(keys("DTIME_RIGHT_VALID_FROM")) = "" & rst1![DTIME_RIGHT_VALID_FROM] & ""
      If rst1![DTIME_RIGHT_VALID_TO] = "1.1.3000" Then itmX.SubItems(keys("DTIME_RIGHT_VALID_TO")) = "souËasnost" Else itmX.SubItems(keys("DTIME_RIGHT_VALID_TO")) = "" & rst1![DTIME_RIGHT_VALID_TO] & ""
    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub


Sub query_kk_data_client_cards(ByVal sql As String)
   
   Sonik_Karta_Klienta.ListView_Cards.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_Cards.ListItems.Add()
      itmX.text = "" & rst1![Text_bank_account_number] & ""
      itmX.SubItems(1) = "" & rst1![TEXT_CARD_NUMBER] & ""
      itmX.SubItems(2) = "" & rst1![DTIME_CARD_APPLICATION] & ""
      itmX.SubItems(3) = "" & rst1![CARD_CHANNEL_TYPE_REQ] & ""
      itmX.SubItems(4) = "" & rst1![DTIME_CARD_VALID_FROM] & ""
      itmX.SubItems(5) = "" & rst1![DTIME_CARD_VALID_TO] & ""
      itmX.SubItems(6) = "" & rst1![DTIME_CARD_ACTIVATION] & ""
      itmX.SubItems(7) = "" & rst1![DTIME_CARD_BLOCKED] & ""
      itmX.SubItems(8) = "" & rst1![DTIME_CARD_CLOSED] & ""
      itmX.SubItems(9) = "" & rst1![DTIME_CARD_UNBLOCKED] & "" 'RISK-868
      itmX.SubItems(10) = "" & rst1![NAME_CARD_STATUS] & "" 'RISK-943
      itmX.SubItems(11) = "" & rst1![NAME_TECHNOLOGY] & "" 'RISK-943

    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub


Sub query_kk_data_client_transactions(ByVal sql As String)
   
   Sonik_Karta_Klienta.ListView_Transactions.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_Transactions.ListItems.Add()
      itmX.text = "" & rst1![Text_bank_account_number] & ""
      itmX.SubItems(1) = "" & rst1![TEXT_CARD_NUMBER] & ""
      itmX.SubItems(2) = "" & rst1![DTIME_BLOCKING] & ""
      itmX.SubItems(3) = SonikFormatDate("" & rst1![DATE_ACCOUNTING] & "") 'RISK-1131
      itmX.SubItems(4) = "" & rst1![AMT_BLOCKED] & ""
      itmX.SubItems(5) = "" & rst1![AMT_CONFIRMED] & ""
      itmX.SubItems(6) = "" & rst1![Code_currency] & ""
      itmX.SubItems(7) = "" & rst1![FLAG_TRANSACTION_ACCOUNTED] & ""
      itmX.SubItems(8) = "" & rst1![TYP_TRANSAKCE] & ""
      itmX.SubItems(9) = "" & rst1![KATEGORIE_TRANSAKCE] & ""
    
    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub


Sub query_kk_data_client_loan_app(ByVal sql As String)
   
   'RISK-955
   Sonik_Karta_Klienta.ListView_Loan_Application.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_Loan_Application.ListItems.Add()
       itmX.text = "" & rst1![CAS2SCORING] & ""
        itmX.SubItems(1) = "" & rst1![CAS_ODESLANI_ZIB] & ""
        itmX.SubItems(2) = "" & rst1![DRUH_ZADOSTI] & ""
        itmX.SubItems(3) = "" & rst1![CISLO_ZADOSTI] & ""
        itmX.SubItems(4) = "" & rst1![CISLO_UVERU] & ""
        itmX.SubItems(5) = "" & rst1![vyse_uveru] & ""
        itmX.SubItems(6) = "" & rst1![VYSLEDEK_ZADOSTI] & ""
        'RISK-955
        itmX.SubItems(7) = "" & rst1![REQ_C_TYPE] & ""
        itmX.SubItems(8) = "" & rst1![REQ_IP_ADDRESS] & ""
        itmX.SubItems(9) = "" & rst1![FIN_C_TYPE] & ""
        itmX.SubItems(10) = "" & rst1![FIN_IP_ADDRESS] & ""
        'RISK-955
        itmX.SubItems(11) = "" & rst1![KANAL_PODPISU_PUJCKY] & ""
        itmX.SubItems(12) = SonikFormatDate("" & rst1![datum_podpisu_pujcky] & "") 'RISK-1131
        itmX.SubItems(13) = "" & rst1![POBOCKA_PUJCKA] & ""
        itmX.SubItems(14) = "" & rst1![ID_ZAMESTNANCA_PUJCKA] & ""
        itmX.SubItems(15) = "" & rst1![DUVOD_ZAMITNUTI_CELKOVO] & ""
        itmX.SubItems(16) = "" & rst1![COMPL_FI_1] & ""
        itmX.SubItems(17) = "" & rst1![STAV_KOMPL_UVER_PRISLIB] & ""
        itmX.SubItems(18) = "" & rst1![DUVOD_ZAMITNUTI_CLIENT_KOMPL] & ""
'RISK-884 p¯esouv· se do aplikaËnÌch dat
'        itmX.SubItems(19) = "" & rst1![AMT_LOAN_PAYMENT] & ""
'        itmX.SubItems(20) = "" & rst1![ZAMESTNANI_HLAVNI_PRIJEM] & ""
'        itmX.SubItems(21) = "" & rst1![ZAMESTNAVATEL_HLAVNI_PRIJEM] & ""
'        itmX.SubItems(22) = "" & rst1![ic_dic_hlavni_prijem] & ""
'        itmX.SubItems(23) = "" & rst1![PRACOVNI_POMER_HLAVNI_PRIJEM] & ""
'        itmX.SubItems(24) = "" & rst1![zamestnan_od_hlavni_prijem] & ""
'        itmX.SubItems(25) = "" & rst1![zamestnan_do_hlavni_prijem] & ""
'        itmX.SubItems(26) = "" & rst1![VYSE_CISTEHO_PRIJMU_HLAVNI] & ""
'        itmX.SubItems(27) = "" & rst1![ZAMESTNANI_VEDLEJSI_PRIJEM] & ""
'        itmX.SubItems(28) = "" & rst1![ZAMESTNAVATEL_VEDLEJSI_PRIJEM] & ""
'        itmX.SubItems(29) = "" & rst1![ic_dic_vedlejsi_prijem] & ""
'        itmX.SubItems(30) = "" & rst1![PRACOVNI_POMER_VEDLEJSI_PRIJEM] & ""
'        itmX.SubItems(31) = "" & rst1![zamestnan_od_vedlejsi_prijem] & ""
'        itmX.SubItems(32) = "" & rst1![ZAMESTNAN_DO_VEDLEJSI_PRIJEM] & ""
'        itmX.SubItems(33) = "" & rst1![VYSE_CISTEHO_PRIJMU_VEDLEJSI] & ""
        itmX.SubItems(19) = SonikFormatDate("" & rst1![DATUM_NACERPANI] & "") 'RISK-1131
        itmX.SubItems(20) = SonikFormatDate("" & rst1![DATUM_UKONCENI] & "") 'RISK-1131
        itmX.SubItems(21) = "" & rst1![DATUM_STORNA] & ""
        itmX.SubItems(22) = "" & rst1![NESPLACENA_JISTINA] & ""


    
    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub




Sub query_kk_data_client_mort_app(ByVal sql As String)
'RISK-884
   Sonik_Karta_Klienta.ListView_Mortgage_Application.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_Mortgage_Application.ListItems.Add()
      itmX.text = "" & rst1![Klient] & ""
        itmX.SubItems(1) = "" & rst1![CAS2SCORING] & ""
        itmX.SubItems(2) = "" & rst1![CAS_ODESLANI_ZIB] & ""
        itmX.SubItems(3) = "" & rst1![DRUH_ZADOSTI] & ""
        itmX.SubItems(4) = "" & rst1![CISLO_ZADOSTI] & ""
        itmX.SubItems(5) = "" & rst1![CISLO_UVERU] & ""
        itmX.SubItems(6) = "" & rst1![vyse_pujcky] & ""
        itmX.SubItems(7) = "" & rst1![VYSLEDEK_ZADOSTI] & ""
        itmX.SubItems(8) = "" & rst1![KANAL_ZADOSTI_APPROVE] & ""
        itmX.SubItems(9) = "" & rst1![IP_ADRESA] & ""
        itmX.SubItems(10) = "" & rst1![KANAL_PODPISU_HYPO_PRISLIB] & ""
        itmX.SubItems(11) = SonikFormatDate("" & rst1![DATUM_PODPISU_HYPO_PRISLIB] & "") 'RISK-1131
        itmX.SubItems(12) = "" & rst1![POBOCKA_HYPOTECNI_PRISLIB] & ""
        itmX.SubItems(13) = "" & rst1![ID_ZAMESTNANCA_HYP_PRISLIB] & ""
        itmX.SubItems(14) = "" & rst1![KANAL_PODPISU_HYPO_SMLOUVA] & ""
        itmX.SubItems(15) = SonikFormatDate("" & rst1![DATUM_PODPISU_HYPO_SMLOUVA] & "") 'RISK-1131
        itmX.SubItems(16) = "" & rst1![POBOCKA_HYPOTECNI_SMLOUVA] & ""
        itmX.SubItems(17) = "" & rst1![ID_ZAMESTNANCA_HYP_SMLOUVA] & ""
        itmX.SubItems(18) = "" & rst1![DUVOD_ZAMITNUTI_CELKOVO] & ""
        itmX.SubItems(19) = "" & rst1![COMPL_FI_1] & ""
        itmX.SubItems(20) = "" & rst1![STAV_KOMPL_UVER_PRISLIB] & ""
        itmX.SubItems(21) = "" & rst1![DUVOD_ZAMITNUTI_CLIENT_KOMPL] & ""
        itmX.SubItems(22) = "" & rst1![COMPL_FI_2] & ""
        itmX.SubItems(23) = "" & rst1![STAV_KOMPL_HYP_SMLOUVA] & ""
        itmX.SubItems(24) = "" & rst1![DUVOD_ZAMITNUTI_CLIENT_HYPSML] & ""
        itmX.SubItems(25) = SonikFormatDate("" & rst1![DATUM_NACERPANI] & "") 'RISK-1131
        itmX.SubItems(26) = SonikFormatDate("" & rst1![DATUM_UKONCENI] & "") 'RISK-1131
        itmX.SubItems(27) = "" & rst1![DATUM_STORNA] & ""
        itmX.SubItems(28) = "" & rst1![NESPLACENA_JISTINA] & ""



    
    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub


'RISK-862
Sub query_kk_data_client_manual_verification(ByVal sql As String)
   
Dim keys As Dictionary
      
Set keys = KliceKListview(Sonik_Karta_Klienta.ListView_Manual_Verification)
   
  Sonik_Karta_Klienta.ListView_Manual_Verification.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_Manual_Verification.ListItems.Add()
      itmX.text = "" & rst1![DATUM_ZALOZENI_ZADOSTI] & ""
        itmX.SubItems(keys("ZaË·tek provÏrky")) = "" & rst1![ZACATEK_PROVERKY] & ""
        itmX.SubItems(keys("Konec provÏrky")) = "" & rst1![KONEC_PROVERKY] & ""
        itmX.SubItems(keys("Workflow")) = "" & rst1![WF] & ""
        itmX.SubItems(keys("V˝sledek aktivity")) = "" & rst1![VYSLEDEK_AKTIVITY] & ""
        itmX.SubItems(keys("Aktivita")) = "" & rst1![AKTIVITA] & ""
        itmX.SubItems(keys("V˝sledek oper·tor")) = "" & rst1![VYSLEDEK_OPERATOR] & ""
        itmX.SubItems(keys("Popisek informace")) = "" & rst1![POPISEK_INFORMACE] & ""
        itmX.SubItems(keys("Oper·tor")) = "" & rst1![Operator] & ""
        itmX.SubItems(keys("»Ìslo û·dosti")) = "" & rst1![CISLO_ZADOSTI] & ""
        itmX.SubItems(keys("ID Ob·lky")) = "" & rst1![ID_ENVELOPE] & ""
        itmX.SubItems(keys("Hodnota")) = "" & rst1![HODNOTA] & ""
        itmX.SubItems(keys("Typ û·dosti")) = "" & rst1![TYP_ZADOSTI] & ""
        itmX.SubItems(keys("KÛd UW aktivity")) = "" & rst1![NAME_UNDERWRITING_TASK_TYPE] & ""
        itmX.SubItems(keys("Role ûadatele")) = "" & rst1![ROLE_ZADATELE] & ""
        itmX.SubItems(keys("Typ p¯Ìjmu")) = "" & rst1![TYP_PRIJMU] & ""
        itmX.SubItems(keys("V˝sledek provÏrky ARES")) = "" & rst1![VYSLEDEK_PROVERKY_ARES] & ""
        itmX.SubItems(keys("Rozsah ovÏ¯enÌ")) = "" & rst1![ROZSAH_OVERENI] & ""
        itmX.SubItems(keys("ZamÏstnavatel")) = "" & rst1![ZAMESTNAVATEL] & ""
        itmX.SubItems(keys("PracovnÌ vztah")) = "" & rst1![PRACOVNI_VZTAH] & ""
        itmX.SubItems(keys("PracovnÌ pomÏr")) = "" & rst1![PRACOVNI_POMER] & ""
        itmX.SubItems(keys("Doba urËit· do")) = SonikFormatDate("" & rst1![DOBA_URCITA_DO] & "") 'RISK-1131
        itmX.SubItems(keys("ProdlouûenÌ DU")) = "" & rst1![PRODLOUZENI_DU] & ""
        itmX.SubItems(keys("Datum n·stupu")) = SonikFormatDate("" & rst1![DATUM_NASTUPU] & "") 'RISK-1131
        itmX.SubItems(keys("ZkuöebnÌ doba")) = "" & rst1![ZKUSEBNI_DOBA] & ""
        itmX.SubItems(keys("V˝povÏdnÌ lh˘ta")) = "" & rst1![VYPOVEDNI_LHUTA] & ""
        itmX.SubItems(keys("Mate¯sk· dovolen·")) = "" & rst1![MATERSKA_DOVOLENA] & ""
        itmX.SubItems(keys("RodiËovsk· dovolen·")) = "" & rst1![RODICOVSKA_DOVOLENA] & ""
        itmX.SubItems(keys("Dlouhodob· nemoc")) = "" & rst1![DLOUHODOBA_NEMOC] & ""
        itmX.SubItems(keys("Sr·ûky, exekuce")) = "" & rst1![SRAZKY_EXEKUCE] & ""
        itmX.SubItems(keys("V˝öe p¯Ìjmu")) = "" & rst1![VYSE_PRIJMU] & ""
        itmX.SubItems(keys("OvÏ¯enÈ I»")) = "" & rst1![OVERENE_IC] & ""
        itmX.SubItems(keys("Mzda vypl·cena")) = "" & rst1![MZDA_VYPLACENA] & ""
        itmX.SubItems(keys("Rizikov˝ p¯Ìpad")) = "" & rst1![RIZIKOVY_PRIPAD] & ""
        itmX.SubItems(keys("Stav vyplnÏnÌ")) = "" & rst1![STAV_VYPLNENI] & ""

    rst1.MoveNext
    Next i
   End If

   rst1.Close
   
End Sub










Sub query_kk_data_client_amg_r8(ByVal sql As String)
   
   Sonik_Karta_Klienta.ListView_AMGR8.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_AMGR8.ListItems.Add()
      itmX.text = "" & rst1![jmeno] & ""
        itmX.SubItems(1) = "" & rst1![PRIJMENI] & ""
        itmX.SubItems(2) = "" & rst1![cas_zalozeni] & ""
        itmX.SubItems(3) = "" & rst1![cas_2scoring] & ""
        itmX.SubItems(4) = "" & rst1![zacatek_activity] & ""
        itmX.SubItems(5) = "" & rst1![PRIJMENI] & ""
        itmX.SubItems(6) = "" & rst1![cas_zalozeni] & ""
        itmX.SubItems(7) = "" & rst1![cas_2scoring] & ""
        itmX.SubItems(8) = "" & rst1![zacatek_activity] & ""
        itmX.SubItems(9) = "" & rst1![konec_activity] & ""
        itmX.SubItems(10) = "" & rst1![ID_ZADOSTI] & ""
        itmX.SubItems(11) = "" & rst1![name_application_type] & ""

    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub


Sub query_kk_data_client_amg_note(ByVal sql As String)
   
   Sonik_Karta_Klienta.ListView_AMG_Note.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_AMG_Note.ListItems.Add()
      itmX.text = "" & rst1![jmeno] & ""
        itmX.SubItems(1) = "" & rst1![PRIJMENI] & ""
        itmX.SubItems(2) = "" & rst1![ID_APP_ENVELOPE] & ""
        itmX.SubItems(3) = "" & rst1![MODIF_TIME] & ""
        itmX.SubItems(4) = "" & rst1![note] & ""
        itmX.SubItems(5) = "" & rst1![PRIJMENI] & ""


    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub



Sub query_kk_data_client_amg_calls(ByVal sql As String)
   
   Sonik_Karta_Klienta.ListView_Calls_R8.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_Calls_R8.ListItems.Add()
      itmX.text = "" & rst1![IDENTIFIKATOR_OBS] & ""
        itmX.SubItems(1) = "" & rst1![UCEL_KONTAKTU] & ""
        itmX.SubItems(2) = "" & rst1![SMER_HOVORU] & ""
        itmX.SubItems(3) = "" & rst1![CAS_ZALOZENI_CONTACTU] & ""
        itmX.SubItems(4) = "" & rst1![PREDVOLBA] & ""
        itmX.SubItems(5) = "" & rst1![UCEL_KONTAKTU] & ""
        itmX.SubItems(6) = "" & rst1![SMER_HOVORU] & ""



    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub



Sub query_kk_data_client_brki(ByVal sql As String)
   
   Sonik_Karta_Klienta.ListView_BRKI.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_BRKI.ListItems.Add()
      itmX.text = "" & rst1![Name_first] & ""
        itmX.SubItems(1) = "" & rst1![Name_last] & ""
        itmX.SubItems(2) = SonikFormatDate("" & rst1![Date_Birth] & "") 'RISK-1131
        itmX.SubItems(3) = "" & rst1![REQUEST_TIME] & ""
        itmX.SubItems(4) = "" & rst1![ID_ZADOSTI] & ""
        itmX.SubItems(5) = "" & rst1![RESPONSE_ERROR_VYSVETLENI] & ""
        itmX.SubItems(6) = "" & rst1![ERRORDESCRIPTION] & ""




    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub




Sub query_kk_data_client_communication(ByVal sql As String)
   
Dim keys As Dictionary

   
   'RISK-864
   Sonik_Karta_Klienta.ListView_Client_Communication.ListItems.Clear
   
Set keys = KliceKListview(Sonik_Karta_Klienta.ListView_Client_Communication)
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_Client_Communication.ListItems.Add()
      itmX.text = "" & rst1![NAME_COMMON_CHANNEL] & ""
        itmX.SubItems(keys("Datum odesl·nÌ zpr·vy")) = "" & rst1![DTIME_SEND] & ""
        itmX.SubItems(keys("Odesl·no na")) = "" & rst1![TEXT_RECEIVER] & ""
        itmX.SubItems(keys("HlaviËka mailu")) = "" & rst1![TEXT_HEADER] & ""
        itmX.SubItems(keys("Zpr·va")) = "" & rst1![TEXT_MESSAGE] & ""
        itmX.SubItems(keys("JmÈno zamÏstnance")) = "" & rst1![NAME_EMPLOYEE] & ""
        itmX.SubItems(keys("OddÏlenÌ")) = "" & rst1![CODE_EMPLOYEE_DEPARTMENT] & ""
        itmX.SubItems(keys("ID_CONTACT")) = "" & rst1![ID_CONTACT] & ""




    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub


'RISK-868
Sub query_kk_data_transactions_homer_Agg(ByVal sql As String)
   
   Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems.Add()
      itmX.text = "" & rst1![CARDNUMBER] & ""
        itmX.SubItems(1) = "" & rst1![DTIME_AUTHORIZATION] & ""
        itmX.SubItems(2) = "" & rst1![PLATBA] & ""
        itmX.SubItems(3) = "" & rst1![ODEZVA_AUTORIZACE] & ""
        itmX.SubItems(4) = "" & rst1![CASTKA_V_CZK] & "" 'RISK-1239
        itmX.SubItems(5) = "" & rst1![OBCHODNI_MISTO] & ""
        itmX.SubItems(6) = "" & rst1![TEXT_CARD_ACCEPTOR_STREET] & ""
        itmX.SubItems(7) = "" & rst1![MESTO] & ""
        itmX.SubItems(8) = "" & rst1![ZEME] & ""
        itmX.SubItems(9) = "" & rst1![Stat] & ""
        itmX.SubItems(10) = "" & rst1![TERMINAL] & ""
        itmX.SubItems(11) = "" & rst1![MCC] & "" 'RISK-999

    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub



'RISK-884
Sub query_kk_data_application_form_data(ByVal sql As String)
   
   Dim keys As Dictionary

    Dim StrHousing As String
    Dim StrIBPayments As String
    Dim StrHHExpenses As String
    Dim StrMainEmployer As String
    Dim StrMainEmploymentData As String
    Dim StrMainEconomicalData As String
    Dim StrMainEmployerAddress As String
    Dim StrAddEmployer As String
    Dim StrAddIncome As String
    Dim StrAddEmployedFrom As String
    Dim StrAddEmploymentData As String
    Dim StrAddEconomicalData As String
    Dim StrAddEmployerAddress As String
    Dim StrPrimaryPhone As String
    Dim StrOtherPhone As String
    Dim StrMainEmplPhone As String
    Dim StrAddEmplPhone As String
    Dim StrSecondIDCard As String
    Dim StrHHIncomes As String
    

    Dim BlnAddEmployer As Boolean





   
      
   Set keys = KliceKListview(Sonik_Karta_Klienta.ListView_Application_Form)
   
   Sonik_Karta_Klienta.ListView_Application_Form.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      
        BlnAddEmployer = False
        
        If rst1![NAME_EMPLOYER_ADD] <> "" And rst1![NAME_EMPLOYER_ADD] <> "XNA" Then BlnAddEmployer = True
      
        StrHousing = rst1![NAME_HOUSING_TYPE] & " / " & SonikFormatDate(rst1![DATE_RESIDENCE_FROM]) 'RISK-1131
        StrIBPayments = rst1![AMT_LOAN_PAYMENT] & " / " & rst1![AMT_CREDITS_SUM] & ""
        If Not IsNull(rst1![AMT_ACCOMMODATION_EXP]) Then StrHHExpenses = rst1![AMT_ACCOMMODATION_EXP] & " / " & rst1![amt_med_food_trans_exp] & " / " & rst1![AMT_OTHER_HOUSE_MEM_INST_EXP] & ""
        
        If IsNull(rst1![NAME_EMPLOYER_MAIN]) Or rst1![NAME_EMPLOYER_MAIN] = "" Or rst1![NAME_EMPLOYER_MAIN] = "XNA" Then
                StrMainEmployer = ""
        Else:   StrMainEmployer = rst1![NAME_EMPLOYER_MAIN] & ", I»O: " & rst1![ICO_MAIN]
        End If
        
        If rst1![DTIME_AGREEMENT_UNTIL_MAIN] = "1.1.3000" Then
            StrMainEmploymentData = "Na dobu neurËitou"
        Else: StrMainEmploymentData = "Doba urËit· do " & SonikFormatDate(rst1![DTIME_AGREEMENT_UNTIL_MAIN]) 'RISK-1131
        End If
            
        StrMainEconomicalData = rst1![NAME_ECONOMICAL_STATUS_MAIN] & " / " & rst1![NAME_EMPLOYMENT_TYPE_MAIN] & " / " & rst1![NAME_JOB_PROFESSION_MAIN] & " / " & rst1![NAME_JOB_INDUSTRY_MAIN]
        If Not IsNull(rst1![NAME_STREET_MAIN]) Then StrMainEmployerAddress = rst1![NAME_STREET_MAIN] & " " & rst1![TEXT_HOUSE_NUM_MAIN] & ", " & rst1![CODE_ZIP_MAIN] & " " & rst1![NAME_TOWN_MAIN] 'Dukelsk· 893, 441 01 Podbo¯any.
        If Not BlnAddEmployer Then StrAddEmployer = "" Else StrAddEmployer = rst1![NAME_EMPLOYER_ADD] & ", I»O: " & rst1![ICO_ADD]
        If Not BlnAddEmployer Then StrAddIncome = "" Else StrAddIncome = "" & rst1![AMT_NET_INCOME_ADD] & ""
        If Not BlnAddEmployer Then StrAddEmployedFrom = "" Else StrAddEmployedFrom = "" & SonikFormatDate(rst1![DTIME_EMPLOYED_FROM_ADD]) 'RISK-1131
        If Not BlnAddEmployer Then
                StrAddEmploymentData = ""
        Else
                If rst1![DTIME_AGREEMENT_UNTIL_ADD] = "1.1.3000" Then
                    StrAddEmploymentData = "Na dobu neurËitou"
                Else: StrAddEmploymentData = "Doba urËit· do " & SonikFormatDate(rst1![DTIME_AGREEMENT_UNTIL_ADD]) 'RISK-1131
                End If
        
        End If
        
        If Not BlnAddEmployer Then StrAddEconomicalData = "" Else StrAddEconomicalData = rst1![NAME_ECONOMICAL_STATUS_ADD] & " / " & rst1![NAME_EMPLOYMENT_TYPE_ADD] & " / " & rst1![NAME_JOB_PROFESSION_ADD] & " / " & rst1![NAME_JOB_INDUSTRY_ADD]
        If Not BlnAddEmployer Then StrAddEmployerAddress = "" Else StrAddEmployerAddress = rst1![NAME_STREET_ADD] & " " & rst1![TEXT_HOUSE_NUM_ADD] & ", " & rst1![CODE_ZIP_ADD] & " " & rst1![NAME_TOWN_ADD] 'Dukelsk· 893, 441 01 Podbo¯any.
        StrPrimaryPhone = rst1![PRIMARY_PHONE_EXTENSION] & "" & rst1![PRIMARY_PHONE]
        If IsNull(rst1![OTHER_PHONE]) Or rst1![OTHER_PHONE] = "XNA" Then StrOtherPhone = "" Else StrOtherPhone = rst1![OTHER_PHONE_EXTENSION] & "" & rst1![OTHER_PHONE]
        StrMainEmplPhone = rst1![MAIN_EMPL_PHONE_EXTENSION] & "" & rst1![MAIN_EMPL_PHONE]
        If Not BlnAddEmployer Then StrAddEmplPhone = "" Else StrAddEmplPhone = rst1![ADD_EMPL_PHONE_EXTENSION] & "" & rst1![ADD_EMPL_PHONE]
        
        If Not IsNull(rst1![AMT_OTHER_HOUSE_MEM_INCOME]) Then StrHHIncomes = rst1![AMT_OTHER_HOUSE_MEM_INCOME] & "" Else StrHHIncomes = ""
        
        Select Case rst1![SECOND_ID_TYPE_PREFERED]
            Case "ID_CARD"
                StrSecondIDCard = "ObËansk˝ pr˘kaz"
            Case "PASSPORT"
                StrSecondIDCard = "CestovnÌ pas"
            Case "BIRTH"
                StrSecondIDCard = "Rodn˝ list"
            Case "STAY_PERMIT"
                StrSecondIDCard = "PovolenÌ k trvalÈmu pobytu"
            Case "DRIVE"
                StrSecondIDCard = "ÿidiËsk˝ pr˘kaz"
            Case "ID_CARD_SUBSTITUTED"
                StrSecondIDCard = "N·hradnÌ identifikaËnÌ pr˘kaz"
            Case Else
                StrSecondIDCard = ""
        End Select
                
                
      
      Set itmX = Sonik_Karta_Klienta.ListView_Application_Form.ListItems.Add()

        itmX.text = "" & rst1![dtime_scoring]
        itmX.SubItems(keys("»as odesl·nÌ do banky")) = "" & rst1![dtime_offer_sent_from_ib] & ""
        itmX.SubItems(keys("Typ û·dosti")) = "" & rst1![name_application_type] & ""
        itmX.SubItems(keys("»Ìslo û·dosti")) = "" & rst1![TEXT_APPLICATION_NUMBER]
        itmX.SubItems(keys("»Ìslo ˙vÏru")) = "" & rst1![code_loan_number]
        itmX.SubItems(keys("Poûadovan˝ ˙vÏrov˝ r·mec")) = "" & rst1![creditamount] & ""
        itmX.SubItems(keys("Stav û·dosti")) = "" & rst1![name_application_status]
        
        itmX.SubItems(keys("Typ bydlenÌ / Od")) = StrHousing
        itmX.SubItems(keys("VzdÏl·nÌ")) = "" & rst1![NAME_EDUCATION]
        itmX.SubItems(keys("Rodinn˝ stav")) = "" & rst1![NAME_MARITAL_STATUS]
        itmX.SubItems(keys("PoËet vyûivovan˝ch osob")) = "" & rst1![NUM_DEPENDENT_PERSON]
        itmX.SubItems(keys("Spl·tky uvedenÈ klientem (Suma spl·tek / v˝öe r·mc˘)")) = StrIBPayments
        itmX.SubItems(keys("V˝daje dom·cnosti (BydlenÌ / lÈky-jÌdlo-doprava / spl·tky ostatnÌch Ëlen˘)")) = StrHHExpenses
        itmX.SubItems(keys("P¯Ìjmy ostat. Ëlen˘ dom·cnosti")) = StrHHIncomes
        
        itmX.SubItems(keys("ZamÏstnavatel (hl. p¯Ìjem)")) = StrMainEmployer
        itmX.SubItems(keys("ZamÏstn·n od")) = "" & SonikFormatDate(rst1![DTIME_EMPLOYED_FROM_MAIN]) 'RISK-1131
        itmX.SubItems(keys("PracovnÌ pomÏr")) = StrMainEmploymentData
        itmX.SubItems(keys("V˝öe hl. p¯Ìjmu")) = "" & rst1![AMT_NET_INCOME_MAIN]
        itmX.SubItems(keys("Profese hl. p¯Ìjmu")) = StrMainEconomicalData
        itmX.SubItems(keys("Adresa zamÏstnavatele - hl. p¯Ìjem")) = StrMainEmployerAddress
        
        itmX.SubItems(keys("ZamÏstnavatel (vedl. p¯Ìjem)")) = StrAddEmployer
        itmX.SubItems(keys("ZamÏstn·n od (vedl. p¯Ìjem)")) = StrAddEmployedFrom
        itmX.SubItems(keys("PracovnÌ pomÏr (vedl. p¯Ìjem)")) = StrAddEmploymentData
        itmX.SubItems(keys("V˝öe vedl. p¯Ìjmu")) = StrAddIncome
        itmX.SubItems(keys("Profese vedl. p¯Ìjmu")) = StrAddEconomicalData
        itmX.SubItems(keys("Adresa zamÏstnavatele - vedl. p¯Ìjem")) = StrAddEmployerAddress
        
        itmX.SubItems(keys("Prim·rnÌ telefon")) = StrPrimaryPhone
        itmX.SubItems(keys("Druh˝ telefon")) = StrOtherPhone
        itmX.SubItems(keys("Telefon do hl. zamÏstn·nÌ")) = StrMainEmplPhone
        itmX.SubItems(keys("Telefon do vedl. zamÏstn·nÌ")) = StrAddEmplPhone
        itmX.SubItems(keys("Prim·rnÌ e-mail")) = "" & rst1![PRIMARY_EMAIL]
        
        itmX.SubItems(keys("Preferovan˝ typ druhÈho dokladu")) = StrSecondIDCard
        
    rst1.MoveNext
    Next i
   End If

   rst1.Close


End Sub








Sub status_change_notices(new_status As String)

Dim skp_data_message_notice As Long
Dim sql_insert_notice As String
Dim sql_update_notice As String
Dim Time_Offset As Integer
Dim sql As String 'RISK-958

'Kontrola jestli je klient v trestu, pokud ano, tak u nÏkter˝ch stav˘ budu zapisovat upozornÏnÌ na jeho trestnÌ kartu


Select Case new_status

       Case "Zpracov·no"
       'automaticky smazat veöker· upozornÏnÌ, kter· byla nastavena automaticky - a d·m dokonËeno - dle toho, kdo je vloûil sys
       
            For i = 1 To Sonik_Datova_Zprava.ListView_DZ_Upozorneni.ListItems.count 'V tomto listview uû jsou pouze aktivnÌ
           
    
                skp_data_message_notice = Sonik_Datova_Zprava.ListView_DZ_Upozorneni.ListItems(i).text
                    
                sql_update_notice = "update " & schema & ".sonik_data_message_notice" & _
                                    " set code_status = 'DokonËeno', DTIME_MODIFIED = sysdate, NAME_EMPLOYEE_MODIFIED = 'sys'" & _
                                    " where name_employee_inserted = 'sys' and skp_data_message_notice = " & skp_data_message_notice
                Call execute(sql_update_notice)
            
            Next i
       
       
       Case "»·steËnÏ zpracov·no"
       'novÈ automatickÈ upozornÏnÌ v sonÌku, doba na zpracov·nÌ 5 dnÌ

            Time_Offset = 5
            sql_insert_notice = "insert into " & schema & ".sonik_data_message_notice values (" & schema & ".sonik_data_message_notice_seq.nextval, " & SKP_datove_zpravy & ",'DokonËi zpracov·nÌ souËinnosti',sysdate,sysdate+ " & Time_Offset & " ,'AktivnÌ','sys', null, null)"
            Call execute(sql_insert_notice)
       

       'Pozn·mka v Trestu - klient ale musÌ v trestu b˝t - zavolam specialni proceduru, ktera se podiva, jestli je klient v trestu a pokud ano, tak vyplni prislusnou poznamku
       'RISK-1060 - rozdÏlenÌ stavu P¯ipojenÌ se ökodou na t¯i
       Case "P¯ipojenÌ se ökodou - PK", "P¯ipojenÌ se ökodou - ˙vÏr", "P¯ipojenÌ se ökodou - AB", "VyrozumÏnÌ - prostudov·nÌ spisu", "Zah·jenÌ trestnÌho stÌh·nÌ", "P¯Ìpad odloûen", "Pod·na ûaloba", "Pod·na obûaloba", "VyrozumÏnÌ o pod·nÌ obûaloby", "N·vrh na potrest·nÌ", "ÿÌzenÌ zastaveno", "PodmÌnÏnÈ zastavenÌ Tÿ", "VyrozumenÌ - nenÌ T»", "VyrozumÏnÌ - nenÌ T»", "OsvÏdËenÌ/NeosvÏdËenÌ"
            Call insert_trest_note_change_status(new_status)
            Call insert_trest_notice_change_status(new_status, 0) 'Druh˝ parametr ¯Ìk·, kolik dnÌ je na zpracov·nÌ upozornÏnÌ
            Call AutUkonceniDatovky(new_status)
        
       'RISK-1022 - tyto stavy majÌ ukonËit datovky, ale nemajÌ dÏlat nic v trestu
       Case "VyrozumÏnÌ o p¯edloûenÌ vÏci", "Vyd·nÌ/vr·cenÌ vÏci"
            Call AutUkonceniDatovky(new_status)


End Select


End Sub


Private Sub AutUkonceniDatovky(new_status As String)

Dim sql As String

        'RISK-958 - automatickÈ ukonËenÌ zpr·vy (nastavenÌ stavu 'Zpracov·no') pokud se klient p¯ed·v· do Trestu - viz stavy naho¯e v CASE
        
        
        If MsgBox("Zad·nÌ stavu " & new_status & " vyvol· automatickÈ ukonËenÌ datovky. Opravdu si p¯ejeö datovku ukonËit?", vbYesNo, "AutomatickÈ ukonËenÌ datovky") = vbYes Then
        
                'UkonËenÌ vöech aktivnÌch upozornÏnÌ
                sql = "UPDATE " & schema & ".sonik_data_message_notice" & _
                     " SET code_status = 'DokonËeno', dtime_modified = SYSDATE, name_employee_modified = 'sys'" & _
                     " WHERE code_status = 'AktivnÌ' AND skp_data_message = " & SKP_datove_zpravy
   
                Call execute(sql)
                
                'UkonËenÌ st·vajÌcÌho stavu - d·v·m plus 1 vte¯inu, aby se to nebylo z d¯Ìve zadan˝m stavem
                sql = "UPDATE " & schema & ".sonik_data_message_status" & _
                     " SET DTIME_VALID_TO = SYSDATE + 1/24/60/60, FLAG_VALID = 'N'" & _
                     " WHERE FLAG_VALID = 'Y' AND skp_data_message = " & SKP_datove_zpravy
                
                Call execute(sql)
        
                'ZaloûenÌ novÈho stavu - zpracov·no - ukonËenÌ datovky - d·v·m plus 2 vte¯iny, aby se to nebylo z d¯Ìve zadan˝m stavem
                        sql = "INSERT INTO  " & schema & ".sonik_data_message_status (ID, SKP_DATA_MESSAGE, SKP_DATA_MESSAGE_STATUS, DTIME_INSERTED, NAME_EMPLOYEE_INSERTED, DTIME_VALID_FROM, DTIME_VALID_TO, FLAG_VALID)" & _
                              " VALUES (" & schema & ".sonik_data_message_status_seq.nextval," & SKP_datove_zpravy & ",4/*Zpracov·no*/,SYSDATE + 2/24/60/60,'sys',SYSDATE + 2/24/60/60,DATE'3000-01-01','Y')"
                
                Call execute(sql)
                
                
        
         End If
        'RISK-958 - Konec

End Sub



Sub insert_trest_note_change_status(new_status As String)

'PouûÌv· se v procedu¯e status_change_notices
'NaËtenÌ dat o p¯Ìtomnosti klienta v trestu
'MusÌm udÏlat loop p¯es vöechny klienty co tam jsou v potvrzen˝ch subjektech

'RISK-985 SonÌk

Dim sql_client_trest As String
Dim Flag_client_in_trest As Boolean
Dim skp_person_for_trest As Long
Dim id_person_for_trest As Long 'RISK-985 SonÌk
Dim max_skp_trest As Long
Dim sql_max_skp_trest As String
Dim sql_insert_trest_note As String

    If Sonik_Datova_Zprava.ListView_Paired_Subjects.ListItems.count > 0 Then
    
        For i = 1 To Sonik_Datova_Zprava.ListView_Paired_Subjects.ListItems.count
        
           Flag_client_in_trest = False
           skp_person_for_trest = ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Paired_Subjects, "SKP_PERSON", i)
           sql_client_trest = "select t.SKP_TREST_STATUS_DATA, p.id_person from " & schema & ".Trest_Status_Data t" & _
                              " join " & schema & ".person p on p.id_person = t.id_person and p.flag_valid = 'Y'" & _
                              " where p.skp_person = " & skp_person_for_trest
           
            Set rst1 = New ADODB.Recordset
             rst1.Open sql_client_trest, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
              If Not rst1.EOF Then
               Flag_client_in_trest = True
               id_person_for_trest = rst1![id_person]
              End If
            rst1.Close
            
            If Flag_client_in_trest Then
            
                sql_max_skp_trest = "select max(skp_trest_note) as max_skp_trest_note from " & schema & ".TREST_NOTE"
                
                
                Set rst2 = New ADODB.Recordset
                 rst2.Open sql_max_skp_trest, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
                 If Not rst2.EOF Then
                 rst2.MoveFirst
                  For j = 1 To rst2.RecordCount
                   max_skp_trest = "" & rst2![max_skp_trest_note] & ""
                   rst2.MoveNext
                  Next j
                 End If
                rst2.Close
                
                'Ta trojka je trest_note_type pozn·mka - p¯id·m jednou pozn·mku
                sql_insert_trest_note = "insert into " & schema & ".TREST_NOTE values" & _
                                    " (" & max_skp_trest + 1 & ", 3, " & id_person_for_trest & ", null, 'SONÕK: (ËÌslo DZ:" & Cislo_datove_zpravy & ") p¯id·n nov˝ stav: " & new_status & ". Platnost od ' || to_char(sysdate,'DD.MM.YYYY HH24:MI:SS'), sysdate, '" & UCase(Main.username) & "')"
             
                Call execute(sql_insert_trest_note)
                'P¯id·m jeötÏ pozn·mku, ve kterÈ se pÌöe, ûe jsem p¯idal upozornÏnÌ... to je na hlavu, kdo to tak vymyslel...
                sql_insert_trest_note = "insert into " & schema & ".TREST_NOTE values" & _
                                    " (" & max_skp_trest + 2 & ", 4, " & id_person_for_trest & ", null, 'P¯id·no upozornÏnÌ: SONÕK: (ËÌslo DZ:" & Cislo_datove_zpravy & ") p¯id·n nov˝ stav: " & new_status & ". Platnost od ' || to_char(sysdate,'DD.MM.YYYY HH24:MI:SS'), sysdate, '" & UCase(Main.username) & "')"
             
                Call execute(sql_insert_trest_note)
                
                
             End If
        Next i
       
    End If


End Sub







Sub insert_trest_notice_change_status(new_status As String, Time_Offset As Integer)

'PouûÌv· se v procedu¯e status_change_notices
'NaËtenÌ dat o p¯Ìtomnosti klienta v trestu
'MusÌm udÏlat loop p¯es vöechny klienty co tam jsou v potvrzen˝ch subjektech

'RISK-985 SonÌk

Dim sql_client_trest As String
Dim Flag_client_in_trest As Boolean
Dim skp_person_for_trest As Long
Dim id_person_for_trest As Long 'RISK-985 SonÌk
Dim max_skp_notice As Long
Dim sql_max_skp_notice As String
Dim sql_insert_notice As String

    If Sonik_Datova_Zprava.ListView_Paired_Subjects.ListItems.count > 0 Then
    
        For i = 1 To Sonik_Datova_Zprava.ListView_Paired_Subjects.ListItems.count
        
           Flag_client_in_trest = False
           skp_person_for_trest = ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Paired_Subjects, "SKP_PERSON", i)
           sql_client_trest = "select t.SKP_TREST_STATUS_DATA, p.id_person from " & schema & ".Trest_Status_Data t" & _
                              " join " & schema & ".person p on p.id_person = t.id_person and p.flag_valid = 'Y'" & _
                              " where p.skp_person = " & skp_person_for_trest
           
           
           
           
            Set rst1 = New ADODB.Recordset
            rst1.Open sql_client_trest, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
                If Not rst1.EOF Then
                    Flag_client_in_trest = True
                    id_person_for_trest = rst1![id_person]
                End If
            rst1.Close
            
            If Flag_client_in_trest Then
            
                sql_max_skp_notice = "select max(skp_notice) as max_skp_notice from " & schema & ".NOTICE"
                
                
                Set rst2 = New ADODB.Recordset
                rst2.Open sql_max_skp_notice, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
                    If Not rst2.EOF Then
                        rst2.MoveFirst
                        For j = 1 To rst2.RecordCount
                            max_skp_notice = "" & rst2![max_skp_notice] & ""
                            rst2.MoveNext
                        Next j
                    End If
                rst2.Close
                
                
                sql_insert_notice = "insert into " & schema & ".NOTICE (skp_notice, id_person, text_notice, dtime_notice_created, dtime_notice_deadline, code_status, employee) " & _
                                    "values (" & max_skp_notice + 1 & ", " & id_person_for_trest & ", 'SONÕK: (ËÌslo DZ:" & Cislo_datove_zpravy & ") p¯id·n nov˝ stav: " & new_status & ". Platnost od ' || to_char(sysdate,'DD.MM.YYYY HH24:MI:SS'), sysdate, trunc(sysdate) + " & Time_Offset & ", 'a','" & UCase(Main.username) & "')"
             
                Call execute(sql_insert_notice)
             End If
        Next i
       
    End If


End Sub











Sub Client_From_Jez_To_Sonik(UserFormName As String, Shift_months As Integer) 'NÏjak nefunguje odkazat se na promÏnnou typu UserForm a vzÌt jejÌ jmÈno, proto takhle na dvakr·t


'RISK-721
'JeötÏ se musÌ po¯eöit p¯Ìstupy uûivatel˘m
'ÿeöeno p¯i inicializaci danÈ karty v JEZ

'Sjednotit userform_prev, at to sedi na jezevcika

Dim sql_insert_client As String, sql As String
Dim sql_max_skp As String, max_skp As Long, SKP_Subject As Long 'Pro vkl·d·nÌ subjektu
Dim i As Integer
Dim Bln_Client_In_Sonik As Boolean, Bln_Client_Message_In_Progress As Boolean, Bln_Client_Already_Added As Boolean, Bln_Client_Added_Calculated As Boolean
Dim sql_conf_pair_part As String


'Pokud se jedn· o neklienta tak nic
If c.skp_person < 1 Then
MsgBox "Je mi lÌto, ale neklienta takto do SonÌka dostat neumÌm.", vbExclamation, "Chyba - neklient"
Exit Sub
End If

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'Inicializace


Bln_Client_In_Sonik = False 'Je-li klient v nÏjakÈ norm·lnÌ datovÈ zpr·vÏ
Bln_Client_Message_In_Progress = False 'Je-li ta zpr·va v neukonËenÈm stavu
Bln_Client_Already_Added = False 'Pokud byl uû klient takto ruËnÏ p¯id·n
Bln_Client_Added_Calculated = False 'Pokud se pro nÏj st·le napoËÌt·vajÌ data dle tabulky client_addition



'NaËtenÌ dat o p¯Ìtomnosti klienta v SonÌku
 sql = "select dm.skp_data_message, cms.name_data_message_status" & _
    " from " & schema & ".sonik_data_message dm" & _
    " join " & schema & ".sonik_data_message_status ms on dm.skp_data_message = ms.skp_data_message and ms.flag_valid = 'Y'" & _
    " join " & schema & ".sonik_cl_data_message_status cms on ms.skp_data_message_status = cms.skp_data_message_status" & _
    " join " & schema & ".sonik_confirm_pair cp on dm.skp_data_message = cp.skp_data_message" & _
    " where dm.skp_data_message > 0 and cp.skp_person = " & c.skp_person

    
 Set rst1 = New ADODB.Recordset
 rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
 If Not rst1.EOF Then
  Bln_Client_In_Sonik = True 'Pokud aspoÚ nÏco najdu, tak klient v SonÌku je
  rst1.MoveFirst
  For i = 0 To rst1.RecordCount - 1
     If rst1![NAME_DATA_MESSAGE_STATUS] <> "Zpracov·no" Then 'Pokud je v nÏjakÈ neukonËenÈ zpr·vÏ
     Bln_Client_Message_In_Progress = True
     End If
     rst1.MoveNext
  Next i
 End If
 rst1.Close


'NaËtenÌ dat o p¯id·v·nÌ klienta do SonÌka z JezevËÌka
sql = "select SKP_CLIENT_ADDITION, case when sysdate between dtime_calculate_from and dtime_calculate_to then 1 else 0 end as FLAG_CLIENT_CALCULATED" & _
   " from " & schema & ".sonik_client_addition" & _
   " where skp_person = " & c.skp_person

    
 Set rst1 = New ADODB.Recordset
 rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
 If Not rst1.EOF Then
  Bln_Client_Already_Added = True 'Pokud aspoÚ nÏco najdu, tak jiû byl klient ruËnÏ do JezevËÌka p¯id·n
  rst1.MoveFirst
  For i = 0 To rst1.RecordCount - 1
     If rst1![FLAG_CLIENT_CALCULATED] = 1 Then 'Pokud se pro nÏj st·le napoËÌt·vajÌ data dle tabulky client_addition
     Bln_Client_Added_Calculated = True
     End If
     rst1.MoveNext
  Next i
 End If
 rst1.Close
 
 
 'Zjistim maximalni skp_subject
sql_max_skp = "select nvl(max(skp_subject),0) as max_skp from " & schema & ".sonik_subject"

   Set rst1 = New ADODB.Recordset
   rst1.Open sql_max_skp, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    
        If rst1![max_skp] >= 1 Then
           max_skp = rst1![max_skp]
        End If
    End If

   rst1.Close
    
    'Jenom pojistka pro prvnÌ spuötÏnÌ
    If max_skp >= 1 Then
    Else
    max_skp = 0
    End If
    SKP_Subject = max_skp + 1
 
 
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


'SamotnÈ akce co se m· st·t

If Bln_Client_Message_In_Progress Then 'Je v SonÌku a v platnÈ datovÈ zpr·vÏ
    'Pouze proklik, nic vÌc.. nechci nic prodluûovat
    

ElseIf Bln_Client_Added_Calculated Then 'Aktu·lnÏ se poËÌt· bokem
    'Pouze proklik, nic vÌc.. nechci nic prodluûovat



ElseIf Not Bln_Client_In_Sonik And Not Bln_Client_Already_Added Then  'V˘bec nenÌ v SonÌku ani v Client Addition

   
    If MsgBox("P¯ejeö si pro tohoto klienta po dobu n·sledujÌcÌch " & Shift_months & " mÏsÌc˘ nechat poËÌtat data v SonÌku?", vbYesNo, "Data klienta v SonÌku") = vbYes Then
        
        
        sql_conf_pair_part = "(SKP_CONFIRM_PAIR,SKP_SUBJECT,SKP_DATA_MESSAGE,SKP_PERSON,ID_CUID,NAME_FIRST, NAME_LAST,DATE_BIRTH,TEXT_BIRTH_NUMBER,DTIME_INSERTED,NAME_EMPLOYEE_INSERTED,NAME_CONFIRMATION_STATUS)"

        
        'VloûenÌ do CUID ADDITION
        sql = "insert into " & schema & ".sonik_client_addition values ( " & schema & ".sonik_client_addition_seq.nextval," & c.skp_person & ",'" & c.CUID & "', '" & UserFormName & "',sysdate,'" & UCase(Main.username) & "',trunc(sysdate), add_months(trunc(sysdate)," & Shift_months & "))"
        Call execute(sql)
        
        'VloûenÌ novÈho subjektu - vloûÌm pouze CUID - RISK-876
        sql = "insert into " & schema & ".sonik_subject values (" & SKP_Subject & ",  -1 ," & "'" & c.CUID & "',null,null,null,null,null,null,null," & "sysdate" & "," & "'" & UCase(Main.username) & "',null)"
        Call execute(sql)
        
        'RISK-992 - ˙prava insertu
        
        'VloûenÌ novÈho potvrzenÈho p·ru
        sql = "insert into  " & schema & ".sonik_confirm_pair (skp_confirm_pair,skp_subject,skp_data_message,skp_person,id_cuid,name_first,name_last,date_birth,text_birth_number,dtime_inserted,name_employee_inserted,name_confirmation_status) " & _
              " values (" & schema & ".sonik_confirm_pair_seq.nextval, " & SKP_Subject & ", -1 ," & c.skp_person & "," & "'" & c.CUID & "'" & "," & "'" & c.Name_first & "'" & "," & "'" & c.Name_last & "'" & "," & "to_date('" & c.Date_Birth & "','DD.MM.YYYY')" & "," & "'" & c.Pin & "'" & "," & "sysdate" & "," & "'" & UCase(Main.username) & "'" & "," & "'Potvrzeno'" & ")"
        Call execute(sql)
    End If
    

    
    
ElseIf Not Bln_Client_In_Sonik And Bln_Client_Already_Added And Not Bln_Client_Added_Calculated Then  'V˘bec nenÌ v SonÌku je v Client Addition ale ne platn˝

       
    If MsgBox("P¯ejeö si pro tohoto klienta po dobu n·sledujÌcÌch " & Shift_months & " mÏsÌc˘ nechat poËÌtat data v SonÌku?", vbYesNo, "Data klienta v SonÌku") = vbYes Then
        
        'Update CUID ADDITION
        sql = "update " & schema & ".sonik_client_addition" & _
          " set dtime_inserted = sysdate" & _
          "    ,name_jezevcik_card = '" & UserFormName & "'" & _
          "    ,name_employee_inserted = '" & UCase(Main.username) & "'" & _
          "    ,dtime_calculate_from = trunc(sysdate) " & _
          "    ,dtime_calculate_to = add_months(trunc(sysdate)," & Shift_months & ")" & _
          " where skp_person = " & c.skp_person
        Call execute(sql)

    End If
       

       
    
ElseIf Bln_Client_In_Sonik And Not Bln_Client_Message_In_Progress And Not Bln_Client_Already_Added Then 'Je v SonÌku ale ne v platnÈ zpr·vÏ a ani v Client Addition


    If MsgBox("P¯ejeö si pro tohoto klienta po dobu n·sledujÌcÌch " & Shift_months & " mÏsÌc˘ nechat poËÌtat data v SonÌku?", vbYesNo, "Data klienta v SonÌku") = vbYes Then
        
        'VloûenÌ do CUID ADDITION
        sql = "insert into " & schema & ".sonik_client_addition values ( " & schema & ".sonik_client_addition_seq.nextval," & c.skp_person & ",'" & c.CUID & "', '" & UserFormName & "',sysdate,'" & UCase(Main.username) & "',trunc(sysdate), add_months(trunc(sysdate)," & Shift_months & "))"
        Call execute(sql)

    End If





ElseIf Bln_Client_In_Sonik And Not Bln_Client_Message_In_Progress And Bln_Client_Already_Added And Not Bln_Client_Added_Calculated Then 'Je v SonÌku ale ne v platnÈ zpr·vÏ a je v Client Addition ale ne platn˝


    If MsgBox("P¯ejeö si pro tohoto klienta po dobu n·sledujÌcÌch " & Shift_months & " mÏsÌc˘ nechat poËÌtat data v SonÌku?", vbYesNo, "Data klienta v SonÌku") = vbYes Then
        
        'Update CUID ADDITION
        sql = "update " & schema & ".sonik_client_addition" & _
          " set dtime_inserted = sysdate" & _
          "    ,name_jezevcik_card = '" & UserFormName & "'" & _
          "    ,name_employee_inserted = '" & UCase(Main.username) & "'" & _
          "    ,dtime_calculate_from = trunc(sysdate) " & _
          "    ,dtime_calculate_to = add_months(trunc(sysdate)," & Shift_months & "))" & _
          " where skp_person = " & c.skp_person
        Call execute(sql)

    End If




End If


     

     

'Pokud se poda¯Ì najÌt zp˘sob, jak to udÏlat p¯Ìmo na odkaz Userform.Hide, tak odstranit tento divn˝ case
'userform prev plnim takto blbe, protoze to uz v Jezevcikovi tak je...
Select Case UserFormName

   Case "Klient"
       Klient.Hide
   Case "Klient_dedictvi"
       Klient_dedictvi.Hide
   Case "Klient_exe"
       Klient_exe.Hide
   Case "Klient_insol"
       Klient_insol.Hide
   Case "Klient_trest"
       Klient_trest.Hide
End Select


     'nastavenÌ minulÈho formul·¯e pro zpÏtn· tlaËÌtka - nebudu dÏlat, nech·m si ten p¯edtÌm, tento k niËemu nepot¯ebuju
     'userform_prv = UserFormName
     
    
     'Inicializace pro kartu klienta
     ID_CUID_KL = c.CUID
     NAME_KL = c.name
     skp_person_KL = c.skp_person

     Sonik_Karta_Klienta.Show
     Call nacteni_dat_Karta_Klienta
     Load Sonik_Karta_Klienta
     
End Sub




'RISK-1094
Public Sub FillAssigneSonikToDo(ByRef YourComboBox As combobox)

Dim sql As String

'NaplnÌm boxÌky - prvnÏ p¯ihl·öenÈho uûivatele, a pak vöechny ostatnÌ z datab·ze
YourComboBox.Clear
YourComboBox.AddItem UCase(JezUserLoggedIn.StrUserName)


sql = "select username from " & schema & ".SONIK_USERS where upper(USERNAME) <> '" & UCase(JezUserLoggedIn.StrUserName) & "' and FLAG_VALID = 'Y' AND FLAG_DISPLAY_TODO = 'Y' ORDER BY username ASC"


Set rst1 = New ADODB.Recordset
rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst1.EOF Then
        rst1.MoveFirst
        For i = 1 To rst1.RecordCount
            
            YourComboBox.AddItem rst1![username]
            
        rst1.MoveNext
        Next i
    End If
rst1.Close

YourComboBox.ListIndex = 0

End Sub


'RISK-1094
Public Sub SonikSaveToDoListChange(SKPDataMessage As Long, Assignee As String, ToDoNote As String)

Dim StrSKPSonikUser As String
Dim sql As String

sql = "select skp_sonik_user from " & schema & ".SONIK_USERS where upper(USERNAME) = '" & Assignee & "'"
'Pot¯ebuji zjisti SKP ASSIGNEE
Set rst1 = New ADODB.Recordset
rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst1.EOF Then 'OËek·v·m jeden z·znam
        StrSKPSonikUser = "" & rst1![skp_sonik_user]
    End If
rst1.Close

If StrSKPSonikUser = "" Or IsNull(StrSKPSonikUser) Then StrSKPSonikUser = "null"

Call TransBegin
    'PrvnÏ ukonËÌm platnost souËasnÈho z·znamu
    sql = "update " & schema & ".SONIK_DATA_MESSAGE_TODO_LIST " & _
          " set flag_valid = 'N', dtime_valid_to = sysdate - 1/24/60/60 " & _
          " where flag_valid = 'Y' and skp_data_message = " & SKPDataMessage
    execute (sql)

    
    sql = "INSERT INTO " & schema & ".sonik_data_message_todo_list (skp_data_message_todo_list, skp_data_message, skp_sonik_user, user_inserted ,text_note, dtime_valid_from, dtime_valid_to, flag_valid, flag_deleted) " & _
         " VALUES (" & schema & ".SONIK_DM_TODO_LIST_SEQ.NEXTVAL, " & SKPDataMessage & ", " & StrSKPSonikUser & ",'" & UCase(JezUserLoggedIn.StrUserName) & "','" & ToDoNote & "', sysdate, date'3000-01-01', 'Y', 'N')"
    execute (sql)

Call TransCommit

End Sub
