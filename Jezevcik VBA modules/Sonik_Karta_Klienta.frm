VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sonik_Karta_Klienta 
   Caption         =   "Soník - karta klienta"
   ClientHeight    =   11415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18120
   OleObjectBlob   =   "Sonik_Karta_Klienta.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sonik_Karta_Klienta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit


Private Data_Export As Workbook 'RISK-1235




Private Sub CommandButton_ATM_Branch_Detail_Click()

Sonik_Navstevy_ATM_Branch.Show

End Sub

Private Sub CommandButton_Damage_Click()


'Dovolím klikání, jenom pokud tam je na co klikat, jinak to spadne
If ListView_Soucinnosti_Klienta.ListItems.count > 0 Then

    If ListView_Soucinnosti_Klienta.SelectedItem.SubItems(7) = vbNullString Then
    
    Sonik_Update_Vyse_Skody.Label_Vyse_Skody_Nadpis.Caption = "Pøidání výše škody"
    
    Else
    
    Sonik_Update_Vyse_Skody.Label_Vyse_Skody_Nadpis.Caption = "Úprava výše škody"
    
    End If
    
    
Sonik_Update_Vyse_Skody.Label_Dz_Number = ListView_Soucinnosti_Klienta.SelectedItem.SubItems(1)
Sonik_Update_Vyse_Skody.Label_Name = NAME_KL & " (" & ID_CUID_KL & ")"
Sonik_Update_Vyse_Skody.TextBox_Amt_Damage.text = ListView_Soucinnosti_Klienta.SelectedItem.SubItems(8)
Sonik_Update_Vyse_Skody.TextBox_Poznamka_DZ.text = ListView_Soucinnosti_Klienta.SelectedItem.SubItems(9)
Sonik_Update_Vyse_Skody.Label_ID_DZ = ListView_Soucinnosti_Klienta.SelectedItem.text
If ListView_Soucinnosti_Klienta.SelectedItem.SubItems(7) <> vbNullString Then
SKP_DAMAGE = ListView_Soucinnosti_Klienta.SelectedItem.SubItems(7)
End If
    
    Sonik_Update_Vyse_Skody.Show


End If

End Sub

Private Sub CommandButton_Export_Current_Card_Click()
'RISK-1235 - refaktoring celé funkce. Funkce teï zaškrtne checkbox "za uživatele"

Dim IntPageOrder As Integer
Dim jmeno_karty_klienta As String
Dim Chbox As Variant
Dim BlnChBoxFound As Boolean

IntPageOrder = Sonik_Karta_Klienta.MultiPage1.value
jmeno_karty_klienta = Sonik_Karta_Klienta.MultiPage1.Pages(IntPageOrder).name

If IntPageOrder = 0 Then Exit Sub

'Prvnì musím všechny odškrtnout
CheckBox_Data_All.value = True
CheckBox_Data_All.value = False

BlnChBoxFound = False

'Projedu všechny checkbox a zaškrtnu ten, jehož název se rovná názvu karty. Ten se potom vyexportuje
For Each Chbox In Sonik_Karta_Klienta.Controls
        If TypeName(Chbox) = "CheckBox" Then
        
            If Replace(Chbox.name, "CheckBox", "") = Replace(jmeno_karty_klienta, "Page", "") Then
                Chbox.value = True
                BlnChBoxFound = True
                Exit For
            End If

        End If
Next Chbox


If Not BlnChBoxFound Then
    MsgBox "Nastala chyba pøi exportu aktuální karty JEZ-ERR: 423145671", vbExclamation, "Karta nenalezena"
Else
    Call PripravaExportu
End If

End Sub

Private Sub CommandButton_Export_Data_Click()
'RISK-1235 - rozdeleni na dve funkce, vytvoreni nove funkce PripravaExportu

'V momentì, kdy je uzivatel nic nevybral, tak nebudu nic exportovat
Dim CB As control
Dim sum_generovat As Integer


sum_generovat = 0
For Each CB In Sonik_Karta_Klienta.Controls
        If Left(CB.name, 8) = "CheckBox" Then
         If CB.value Then
          sum_generovat = sum_generovat + 1
         End If
        End If
Next CB

If sum_generovat = 0 Then

MsgBox "Nebylo vybráno nic pro export", vbExclamation, "Chybné potvrzení hodnot"

Else

Call PripravaExportu

End If



End Sub


Private Sub PripravaExportu()
'RISK-1235 - funkce vyjmuta z kodu buttonu CommandButton_Export_Data_Click()
'Vola se jeste z druheho buttonu CommandButton_Export_Current_Card_Click

Dim CheckBox As Variant
Dim StrNadpisDokumentu As String 'RISK-1103
Dim LV As ListView 'pouziva se pro dynamicky odkaz na listview - kartu klienta

Application.ScreenUpdating = False

Set Data_Export = Workbooks.Add

For Each CheckBox In Sonik_Karta_Klienta.Controls
            If TypeName(CheckBox) = "CheckBox" Then
                If CheckBox = True And CheckBox.Caption <> "Nebo oznaè vše" Then
                    
                    'RISK-1103 nìkde chtìli na nadpisu exportu jiný název než je název checkboxu jinde zùstane stejný
                    StrNadpisDokumentu = CheckBox.Caption 'Iniciálnì
        

                    'naplnìní, které karty se mají exportovat
                    If CheckBox.Caption = "Data o klientovi aktuální" Then
                        Set LV = Sonik_Karta_Klienta.ListView_Client_Actual
                    
                    ElseIf CheckBox.Caption = "Data o klientovi - RS" Then
                        Set LV = Sonik_Karta_Klienta.ListView_Client_GC
                        StrNadpisDokumentu = "Data o klientovi v Rámcové smlouvì"
                    
                    ElseIf CheckBox.Caption = "Klient - zmìna RÈ" Then
                        Set LV = Sonik_Karta_Klienta.ListView_Birth_Number_Change
                        StrNadpisDokumentu = "Zmìna údajù klienta"
                                        
                    ElseIf CheckBox.Caption = "Klient - zmìna Jméno, Pøíjmení" Then
                        Set LV = Sonik_Karta_Klienta.ListView_Name_Change
                        StrNadpisDokumentu = "Zmìna údajù klienta"
                    
                    ElseIf CheckBox.Caption = "Klient - zmìna Email, Mobil" Then
                        Set LV = Sonik_Karta_Klienta.ListView_Change_Mail_Phone
                        StrNadpisDokumentu = "Zmìna kontaktních údajù"
                        
                    ElseIf CheckBox.Caption = "Klient - zmìna adresy" Then
                        Set LV = Sonik_Karta_Klienta.ListView_Address_Change
                        StrNadpisDokumentu = "Zmìna adresy klienta"

                    
                    
                    
                    ElseIf CheckBox.Caption = "Login IB" Then
                        Set LV = Sonik_Karta_Klienta.ListView_Login_IB
                        StrNadpisDokumentu = "Logovací soubory a IP adresy"
                    
                    ElseIf CheckBox.Caption = "Bìžný a spoøící úèet" Then
                        Set LV = Sonik_Karta_Klienta.ListView_Accounts
                        StrNadpisDokumentu = "Pøehled žádostí o úèet"
                        
                    ElseIf CheckBox.Caption = "Disponentská práva" Then 'RISK-992
                        Set LV = Sonik_Karta_Klienta.ListView_Disponent_Rights
                        StrNadpisDokumentu = "Pøehled disponentských práv"
                        
                    
                    ElseIf CheckBox.Caption = "Platební karty" Then
                        Set LV = Sonik_Karta_Klienta.ListView_Cards
                    
                    ElseIf CheckBox.Caption = "Platební karty - transakce" Then
                        Set LV = Sonik_Karta_Klienta.ListView_Transactions
                        StrNadpisDokumentu = "Transakce platební kartou"
                    
                    'RISK-884
                    ElseIf CheckBox.Caption = "Aplikaèní data" Then
                        Set LV = Sonik_Karta_Klienta.ListView_Application_Form
                        StrNadpisDokumentu = "Aplikaèní data"
                    
                    ElseIf CheckBox.Caption = "Žádost_Úvìr" Then
                        Set LV = Sonik_Karta_Klienta.ListView_Loan_Application
                        StrNadpisDokumentu = "Pøehled žádostí o úvìr"
                    
                    ElseIf CheckBox.Caption = "Žádost_Hypotéka" Then
                        Set LV = Sonik_Karta_Klienta.ListView_Mortgage_Application
                        StrNadpisDokumentu = "Pøehled žádostí o hypotéku"
                    
                    ElseIf CheckBox.Caption = "Manuální verifikace" Then
                        Set LV = Sonik_Karta_Klienta.ListView_Manual_Verification
                    
                    ElseIf CheckBox.Caption = "Manuální schvalování - AMG R8" Then
                        Set LV = Sonik_Karta_Klienta.ListView_AMGR8
                        StrNadpisDokumentu = "Manuální schvalování AMG R8"
                    
                    ElseIf CheckBox.Caption = "Man. schvalování - AMG Note R8" Then
                        Set LV = Sonik_Karta_Klienta.ListView_AMG_Note
                        StrNadpisDokumentu = "Manuální schvalování AMG Note R8"
                    
                    ElseIf CheckBox.Caption = "Man. schvalování - Calls R8" Then
                        Set LV = Sonik_Karta_Klienta.ListView_Calls_R8
                        StrNadpisDokumentu = "Manuální schvalování Calls R8"

                    ElseIf CheckBox.Caption = "BRKI dotazy" Then
                        Set LV = Sonik_Karta_Klienta.ListView_BRKI
                    
                    ElseIf CheckBox.Caption = "Komunikace s klientem" Then
                        Set LV = Sonik_Karta_Klienta.ListView_Client_Communication
                    'RISK-868
                    ElseIf CheckBox.Caption = "Transakce z Homéra" Then
                        Set LV = Sonik_Karta_Klienta.ListView_Transactions_Homer
                        StrNadpisDokumentu = "Pøehled plateb platební kartou"
                    
                    'POZOR Pokud se bude nasazovat s RISK-884 tak pøidat i RISK-884 - Formuláø aplikaèní data
                    'a Také RISK-992 a RISK-978 - Formuláø disponenti
                    
                    End If
                    
                    
                    Call export_dat_kk(CheckBox.Caption, LV, StrNadpisDokumentu, Data_Export)
    
                End If
            End If
Next CheckBox



Application.DisplayAlerts = False
Data_Export.Sheets(1).Delete
Application.DisplayAlerts = True

Application.ScreenUpdating = True


End Sub



Private Sub export_dat_kk(jmeno_datove_karty As String, LV As ListView, NadpisKarty As String, Data_Export As Workbook) 'RISK-1235 - pøidán parametr
'RISK-1235 - funkce pøesunuta do formu Sonik_Karta_Klienta a pøedìlána na private
'RISK-1103
Dim ncol As Integer
Dim nrow As Integer
Dim StrNazevSloupce As String 'RISK-959 - oprava exportu èísla karty - pøevod na text
Dim IntRowShift As Integer 'RISK-1103 - kde mají zaèít být plnìna data

'RISK-1158
Dim BlnOpenApiCommentDisplay As Boolean

    Sheets.Add(after:=Sheets(Sheets.count)).name = jmeno_datove_karty
    
    ncol = LV.ColumnHeaders.count
    nrow = LV.ListItems.count
    
    For i = 1 To ncol
    IntRowShift = 4
    Data_Export.Sheets(jmeno_datove_karty).Cells(1 + IntRowShift, i) = LV.ColumnHeaders(i).text
                
    'RISK-959 - vyberu název sloupeèku
 
    StrNazevSloupce = LV.ColumnHeaders(i).text

    'RISK-1158 - inicialne false pro jistotu...
    BlnOpenApiCommentDisplay = False
                
                
                   Data_Export.Sheets(jmeno_datove_karty).Cells(1 + IntRowShift, i).Font.Bold = True
                   Data_Export.Sheets(jmeno_datove_karty).Cells(1 + IntRowShift, i).Interior.Color = 5296274 'Nìjaká AB zelená
                     
    
                    With Data_Export.Sheets(jmeno_datove_karty).Cells(1 + IntRowShift, i).borders(xlEdgeBottom)
                        .LineStyle = xlDouble
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThick
                    End With
               'RISK-1103
               Data_Export.Sheets(jmeno_datove_karty).Range("A1").value = NadpisKarty
               Data_Export.Sheets(jmeno_datove_karty).Range("A1").Font.Bold = True
               Data_Export.Sheets(jmeno_datove_karty).Range("A3").value = NAME_KL & ", nar. " & Sonik_Karta_Klienta.Label_Date_Birth
               Data_Export.Sheets(jmeno_datove_karty).Range("A3").Font.Bold = True
                
               Call ExportWithHeaderFooter(ActiveSheet)
               ActiveSheet.PageSetup.Orientation = xlLandscape 'Uživatelé to vždycky obracejí, takže to mùžu udìlat za nì. Vleze se tam potom více informací.
               'IntRowShift = 11
               For j = 1 To nrow  'RISK-1103 posunu až na dolù
               

                'Protože na jednu A4 stranu se vejde 50 øádkù a hlavièka zabere asi 6, tak je potøeba na každé nové stránce opìt odskoèit text.
                'If (j + 1 + IntRowShift) Mod 50 = 0 Then IntRowShift = IntRowShift + 7
                
                
                If i = 1 Then
                
                
                    
                   
                
                    If StrNazevSloupce = "Èíslo karty" Then
                        Data_Export.Sheets(jmeno_datove_karty).Cells(j + 1 + IntRowShift, i).NumberFormat = "@"
                        Data_Export.Sheets(jmeno_datove_karty).Cells(j + 1 + IntRowShift, i) = CStr(LV.ListItems(j).text)
                    Else: Data_Export.Sheets(jmeno_datove_karty).Cells(j + 1 + IntRowShift, i) = LV.ListItems(j).text
                    End If
                Else
 
                    
                    If StrNazevSloupce = "Èíslo karty" Then
                        Data_Export.Sheets(jmeno_datove_karty).Cells(j + 1 + IntRowShift, i).NumberFormat = "@"
                        Data_Export.Sheets(jmeno_datove_karty).Cells(j + 1 + IntRowShift, i) = CStr(LV.ListItems(j).SubItems(i - 1))
                    Else: Data_Export.Sheets(jmeno_datove_karty).Cells(j + 1 + IntRowShift, i) = LV.ListItems(j).SubItems(i - 1)
                          'RISK-1158 -pokud nìkde najdu text Open API, tak na konec dokumentu doplním vysvìtlení
                          If Data_Export.Sheets(jmeno_datove_karty).Cells(j + 1 + IntRowShift, i).value = "Open API" Then BlnOpenApiCommentDisplay = True
                    End If
                End If
              Next j
    Next i
    Data_Export.Sheets(jmeno_datove_karty).Columns("A:AZ").EntireColumn.AutoFit

    'RISK-1158 Textaci umístím o 4 øádky pod poslední øádek
    If BlnOpenApiCommentDisplay Then Data_Export.Sheets(jmeno_datove_karty).Cells(j + 1 + IntRowShift + 4, 1).value = "Open API = Rozhraní pro pøihlášení klienta do internetového bankovnictví prostøednictvím aplikace tøetí strany. Viz Evropská smìrnice o platební styku PSD2."
    

End Sub










Private Sub CommandButton_Export_NBU_Click()

Call Sonik_Export_sablona_NBU

End Sub


Private Sub CommandButton_Export_PK_Click()

Dim i As Integer

'RISK-868
If Sonik_Karta_Klienta.ListView_Cards.ListItems.count = 0 Then
    MsgBox "Soník ke klientovi neeviduje žádnou platební kartu.", vbExclamation, "Chyba - klient nemá kartu"
    Exit Sub
End If

Sonik_Vstupy_SablonaPK.Show

End Sub


Private Sub CommandButton_Export_Soucinnosti_Click()
Call Sonik_AO_Export_automaticka_odpoved

End Sub

Private Sub CommandButton_Overview_Click()

'Pøehodí mì na první stránku multipage
Sonik_Karta_Klienta.MultiPage1.value = 0

End Sub

Private Sub CommandButton_Trest_Click()

'Pøeskoèím do trestu pro daného klienta

userform_prv = "Sonik"

'If cs_HDWMN.State = 1 Then
'prihlasen:
  'skp_person = skp_person_KL 'RISK-985ZAP
  id_person = ConvertSkptoId(skp_person_KL) 'Tady byl pøedtím kód, který incializovat klienta ze soníka do trestu pøes SKP. Napsána nová funkce, která konvertuje SKP na ID person
  Call set_filter(Trest.name, 1)
  
  Me.Hide
  Klient_trest.Show

' If err_connection Then
'  err_connection = False
'  Call re_login(Main.username.text, Main.password.text)
'  If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
' End If
'
'Else
' Call re_login(Main.username.text, Main.password.text)
' If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
'End If

End Sub



Private Sub Label_Jmeno_Klienta_Click()

Dim lastSpace As Integer, lastName As String, firstSpace As Integer, firstName As String
Dim clipboard As Variant




If skp_person_KL <> -1 Then

         firstSpace = InStrRev(NAME_KL, " ")
         firstName = Left(NAME_KL, firstSpace - 1)
            
         lastSpace = InStrRev(NAME_KL, " ")
         lastName = Mid(NAME_KL, lastSpace + 1)

         Set clipboard = New MSForms.DataObject
         clipboard.SetText lastName & " " & firstName & "_" & ID_CUID_KL
         clipboard.PutInClipboard
        
         Label_Copy_Values.Caption = "Údaje pro složky klienta zkopírovány"
         Label_Flag_Copy_Values.Visible = True
         Label_Copy_Values.Visible = True
End If




End Sub

Private Sub ListView_Soucinnosti_Klienta_DblClick()


If ListView_Soucinnosti_Klienta.ListItems.count > 0 Then


'Výbìr datové zprávy


'Toto by mely byt public promenne nadefinovane v sub SONIK
SKP_datove_zpravy = ListView_Soucinnosti_Klienta.SelectedItem.text
Cislo_datove_zpravy = ListView_Soucinnosti_Klienta.SelectedItem.SubItems(1)
Datum_doruceni_datove_zpravy = ListView_Soucinnosti_Klienta.SelectedItem.SubItems(2)
Odesilatel_datove_zpravy = ListView_Soucinnosti_Klienta.SelectedItem.SubItems(3)
Cislo_Jednaci_datove_zpravy = ListView_Soucinnosti_Klienta.SelectedItem.SubItems(4)
Spisova_Znacka_datove_zpravy = ListView_Soucinnosti_Klienta.SelectedItem.SubItems(5)


Me.Hide
Sonik_Datova_Zprava.Show


Call nacteni_dat_Sonik_Datova_Zprava

End If

End Sub



Private Sub UserForm_Activate()
Dim sql_kk_data_client As String
Dim sql_kk_client_trest As String
Dim sql_kk_client_actual As String
Dim sql_kk_client_ATM As String
Dim sql_kk_client_BRANCH As String


Application.Cursor = xlWait

 Label_Flag_Copy_Values.Visible = False
 Label_Copy_Values.Visible = False
 
 


Sonik_Karta_Klienta.MultiPage1.value = 0
'Nìjak blblo zobrazování první karty, zkusím takovouto fintu
Sonik_Karta_Klienta.MultiPage1.value = 1
'Pøehodí mì na první stránku multipage
Sonik_Karta_Klienta.MultiPage1.value = 0

Label_CUID.Caption = ID_CUID_KL
Label_Jmeno_Klienta.Caption = NAME_KL & " (" & ID_CUID_KL & ")"


'Naètení základních dat o klientovi

sql_kk_data_client = "select NAME_FIRST , NAME_LAST, Text_Title_before, Text_Title_After ,TEXT_BIRTH_NUMBER, DATE_BIRTH, TRVALA_ADRESA, KORESPONDENCNI_ADRESA, TEXT_MOBILE_PHONE, TEXT_EMAIL_ADDRESS from " & schema & ".sonik_klient_aktual where cuid = '" & ID_CUID_KL & "'"


Set rst1 = New ADODB.Recordset
   rst1.Open sql_kk_data_client, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      
      Label_Name.Caption = "" & rst1![Text_Title_before] & " " & rst1![Name_first] & " " & rst1![Name_last] & " " & rst1![Text_Title_After] & ""
      Label_RC.Caption = "" & rst1![TEXT_BIRTH_NUMBER] & ""
      Label_Date_Birth.Caption = "" & rst1![Date_Birth] & ""
      Label_Adress_Permanent.Caption = "" & rst1![TRVALA_ADRESA] & ""
      Label_Adress_Temporary.Caption = "" & rst1![KORESPONDENCNI_ADRESA] & ""
      Label_Phone_Number.Caption = "" & rst1![text_mobile_phone] & ""
      Label_Email.Caption = "" & rst1![text_email_address] & ""
      
      
  
    rst1.MoveNext
    Next i
   
   CommandButton_Export_Data.Enabled = True
   CommandButton_Export_Current_Card.Enabled = True
   CommandButton_Export_NBU.Enabled = True
   CommandButton_Export_Soucinnosti.Enabled = True
   
   'V pøípadì, že data ještì nejsou dostupná
   Else
   
      Label_Name.Caption = ""
      Label_RC.Caption = ""
      Label_Date_Birth.Caption = ""
      Label_Adress_Permanent.Caption = ""
      Label_Adress_Temporary.Caption = ""
      Label_Phone_Number.Caption = ""
      Label_Email.Caption = ""
   
   CommandButton_Export_Data.Enabled = False
   CommandButton_Export_Current_Card.Enabled = False
   CommandButton_Export_NBU.Enabled = False
   CommandButton_Export_Soucinnosti.Enabled = False
   
   End If

   rst1.Close


'Naètení dat o pøítomnosti klienta v trestu

    sql_kk_client_trest = "select t.SKP_TREST_STATUS_DATA, cl.name_trest_status, t.dtime_valid_from, t.dtime_valid_to, t.dtime_modification ,t.text_note, t.flag_deleted " & _
    " from " & schema & ".Trest_Status_Data t" & _
    " join " & schema & ".Person p on t.id_person = p.id_person and p.flag_valid = 'Y'" & _
    " join " & schema & ".Trest_Status cl on t.skp_trest_status = cl.skp_trest_status" & _
    " where p.skp_person = " & skp_person_KL

'Iniciální nastavení
Label_Flag_Trest_Podano_TO.Caption = "Trestní oznámení na klienta nebylo podáno."
CommandButton_Trest.Enabled = True

Set rst1 = New ADODB.Recordset
   rst1.Open sql_kk_client_trest, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      

      
       If rst1![DTIME_VALID_TO] = "1.1.3000" Then
       
       Label_Flag_Trest.Caption = "Klient je aktuálnì veden v trestu ve stavu " & rst1![NAME_TREST_STATUS] & " a to od " & rst1![DTIME_VALID_FROM] & "."
       
       
       End If
       
       If rst1![NAME_TREST_STATUS] = "Podáno TO" Then
       
       Label_Flag_Trest_Podano_TO.Caption = "Na klienta bylo podáno trestní oznámení " & rst1![DTIME_VALID_FROM] & "."
       
       
       End If
       
    rst1.MoveNext
    Next i
   
   'Pokud vùbec není v trestu tak nenapíšu nic a zakážu button
   Else
   
       Label_Flag_Trest.Caption = "Klient není evidován v Trestu."
       Label_Flag_Trest_Podano_TO.Caption = ""
       
       CommandButton_Trest.Enabled = False
   
   End If

   rst1.Close


'Má klient aktuální data?

    If skp_person_KL = -1 Then
        sql_kk_client_actual = "select dcb.text_card_number, st.endtime from " & schema & ".sonik_data_calc_base_card dcb join " & schema & ".sonik_tracking st on trunc(st.endtime) = trunc(sysdate) and st.result = 'SUCCEDED' where dcb.text_card_number = '" & NUM_Card_KL & "'"
    Else
        sql_kk_client_actual = "select dcb.cuid, st.endtime from " & schema & ".sonik_data_calc_base dcb join " & schema & ".sonik_tracking st on trunc(st.endtime) = trunc(sysdate) and st.result = 'SUCCEDED' where dcb.cuid = '" & ID_CUID_KL & "'"
    End If

Set rst1 = New ADODB.Recordset
   rst1.Open sql_kk_client_actual, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then

        Label_Data_Nactena.Caption = ""
        Label_Missing_Data.Visible = False
    
   Else
      
        Label_Data_Nactena.Caption = "Data ke klientovi dnes ještì nejsou dostupná"
        Label_Missing_Data.Visible = True
      
   End If

   rst1.Close

'Naètení dat o návštìvách ATM a poboèek


'Iniciální nastavení
Label_ATM.Caption = "Data o klientových výbìrech z bankomatu nejsou k dispozici."
Label_Branch.Caption = "Soník neví nic o tom, jestli se klient objevil na poboèce nebo ne."
CommandButton_ATM_Branch_Detail.Enabled = False
Label_ATM_NOK.Visible = False
Label_ATM_OK.Visible = False
Label_BRANCH_NOK.Visible = False
Label_BRANCH_OK.Visible = False




'Má se zafixovat k datu doruèení datové zprávy, pokud je na klientovi datových zpráv více, tak kterou vzít?
'Zatím vezmu tu poslední
Dim datum_zadosti_o_soucinnost As Date
Dim pocet_vyberu_v_toleranci As Integer
Dim pocet_navstev_v_toleranci As Integer
Dim Flag_vyber As Integer
Dim Flag_navsteva As Integer
Dim page As Variant

    'Vyberu maximální datum doruceni zpravy, ze vsech zprav klienta
    For i = 1 To ListView_Soucinnosti_Klienta.ListItems.count
      If datum_zadosti_o_soucinnost < ListView_Soucinnosti_Klienta.ListItems(i).SubItems(2) Then
      datum_zadosti_o_soucinnost = ListView_Soucinnosti_Klienta.ListItems(i).SubItems(2)
      End If
    Next i


'Naètení dat o ATM
If skp_person_KL = -1 Then

    sql_kk_client_ATM = "select TEXT_CARD_NUMBER, DTIME_ATM_TRANSACTION, NAME_ATM_TRANSACTION_TYPE, AMT_TRANSACTION, ID_ATM_TERMINAL, ATM_ADDRESS from " & schema & ".sonik_visit_ATM" & _
    " where TEXT_CARD_NUMBER = " & NUM_Card_KL
    

Else

    sql_kk_client_ATM = "select TEXT_CARD_NUMBER, DTIME_ATM_TRANSACTION, NAME_ATM_TRANSACTION_TYPE, AMT_TRANSACTION, ID_ATM_TERMINAL, ATM_ADDRESS from " & schema & ".sonik_visit_ATM" & _
    " where skp_person = " & skp_person_KL

End If

Set rst1 = New ADODB.Recordset
   rst1.Open sql_kk_client_ATM, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      
       If rst1![DTIME_ATM_TRANSACTION] >= datum_zadosti_o_soucinnost - 60 Then
       
       pocet_vyberu_v_toleranci = pocet_vyberu_v_toleranci + 1
       

       End If
       
    rst1.MoveNext
    Next i
    
   'Pomocny flag, pokud mám nìjaká data k dispozici, nezávisle na toleranci, tak zobrazím tlaèítko
   Flag_vyber = 1
    
   End If

   rst1.Close

If pocet_vyberu_v_toleranci >= 1 And Label_Missing_Data.Visible = False Then
Label_ATM.Caption = "Klient se otøel o náš bankomat v posledních 60 dnech od doruèení datové zprávy"
Label_ATM_OK.Visible = True

ElseIf pocet_vyberu_v_toleranci = 0 And Label_Missing_Data.Visible = False Then
Label_ATM.Caption = "Klient se na náš bankomat v posledních 60 dnech od doruèení datové zprávy ani nepodíval"
Label_ATM_NOK.Visible = True

Else
Label_ATM.Caption = "Data o klientových výbìrech z bankomatu dnes ještì nejsou k dispozici."
End If


       


'Naètení dat o poboèkách
sql_kk_client_BRANCH = "select DTIME_CONTACT_REALIZED, NAME_EMPLOYEE, NAME_BRANCH_OFFICE, BRANCH_ADDRESS from " & schema & ".sonik_visit_branch" & _
" where skp_person = " & skp_person_KL

Set rst1 = New ADODB.Recordset
   rst1.Open sql_kk_client_BRANCH, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      
       If rst1![DTIME_CONTACT_REALIZED] >= datum_zadosti_o_soucinnost - 30 Then
       
       pocet_navstev_v_toleranci = pocet_navstev_v_toleranci + 1
       

       End If
       
    rst1.MoveNext
    Next i
    
   'Pomocny flag, pokud mám nìjaká data k dispozici, nezávisle na toleranci, tak zobrazím tlaèítko
   Flag_navsteva = 1
    
   End If

   rst1.Close

If pocet_navstev_v_toleranci >= 1 And Label_Missing_Data.Visible = False Then
Label_Branch.Caption = "Klient si na naší poboèce v posledních 30 dnech od doruèení datové zprávy dával kafe a bonbón"
Label_BRANCH_OK.Visible = True

ElseIf pocet_navstev_v_toleranci = 0 And Label_Missing_Data.Visible = False Then
Label_Branch.Caption = "Klient o naši poboèku v posledních 30 dnech ani pohledem nezavadil."
Label_BRANCH_NOK.Visible = True

Else
Label_Branch.Caption = "Dnes ještì Soník neví nic o tom, jestli se klient objevil na poboèce nebo ne."
End If


If Flag_navsteva = 1 Or Flag_vyber = 1 Then
CommandButton_ATM_Branch_Detail.Enabled = True
Else: CommandButton_ATM_Branch_Detail.Enabled = False

End If



'Pokud se jedná o neklienta, tak chci trošku jiný vzhled karty
If skp_person_KL = -1 Then

 Sonik_Karta_Klienta.Caption = "Soník - karta neklienta"
 Sonik_Karta_Klienta.Label_Nadpis.Caption = "Karta Neklienta:"
 Sonik_Karta_Klienta.Label_Jmeno_Klienta.Caption = "Èíslo karty: " & NUM_Card_KL
 Sonik_Karta_Klienta.Label_Soucinnosti_Klienta.Caption = "Evidence neklienta v Soníku"
 Sonik_Karta_Klienta.CommandButton_Damage.Enabled = False
 
 For Each page In MultiPage1.Pages
    If page.Caption <> "Pøehled" Then
       page.Visible = False
    End If
 Next page
 
 
'Pokud je to ale klient, tak to vrátím zpátky, ono by to totiž pak nefungovalo, když se pøekliknu z neklienta na klienta
Else:

 Sonik_Karta_Klienta.Caption = "Soník - karta klienta"
 Sonik_Karta_Klienta.Label_Nadpis.Caption = "Karta Klienta:"
 Sonik_Karta_Klienta.Label_Jmeno_Klienta.Caption = NAME_KL & " (" & ID_CUID_KL & ")"
 Sonik_Karta_Klienta.Label_Soucinnosti_Klienta.Caption = "Evidence klienta v Soníku"
 Sonik_Karta_Klienta.CommandButton_Damage.Enabled = True
 
 For Each page In MultiPage1.Pages
    page.Visible = True
 Next page
 
End If


'RISK-721

'Dovolím klikání, jenom pokud tam je na co klikat, jinak to spadne
If ListView_Soucinnosti_Klienta.ListItems.count = 0 Then
    CommandButton_Export_NBU.Enabled = False
    CommandButton_Export_Soucinnosti.Enabled = False
End If


'Promazu

datum_zadosti_o_soucinnost = 0
pocet_vyberu_v_toleranci = 0
pocet_navstev_v_toleranci = 0
Flag_vyber = 0
Flag_navsteva = 0

Application.Cursor = xlDefault

End Sub




Private Sub CheckBox_Data_All_Click()

Dim CB As control

'V momentì, kdy je prázdný a kliknu na nìj, tak už je tam hodnota True
If CheckBox_Data_All.value Then

        For Each CB In Sonik_Karta_Klienta.Controls
        If Left(CB.name, 8) = "CheckBox" Then
        CB.value = True
        End If
        Next CB
        
Else

        For Each CB In Sonik_Karta_Klienta.Controls
        If Left(CB.name, 8) = "CheckBox" Then
        CB.value = False
        End If
        Next CB
End If

End Sub






'Kopírování textu

Private Sub Label_CUID_Click()

If Label_CUID.Caption <> vbNullString Then

Dim clipboard As MSForms.DataObject
 Set clipboard = New MSForms.DataObject
 clipboard.SetText Label_CUID.Caption
 clipboard.PutInClipboard
 
 Label_Copy_Values.Caption = "CUID zkopírován do schránky"
 Label_Flag_Copy_Values.Visible = True
 Label_Copy_Values.Visible = True
 
End If
 
End Sub

Private Sub Label_Name_Click()

If Label_Name.Caption <> vbNullString Then

Dim clipboard As MSForms.DataObject
 Set clipboard = New MSForms.DataObject
 clipboard.SetText Label_Name.Caption
 clipboard.PutInClipboard
 
 Label_Copy_Values.Caption = "Jméno zkopírováno do schránky"
 Label_Flag_Copy_Values.Visible = True
 Label_Copy_Values.Visible = True
 
End If
 
End Sub

Private Sub Label_RC_Click()

If Label_RC.Caption <> vbNullString Then

Dim clipboard As MSForms.DataObject
 Set clipboard = New MSForms.DataObject
 clipboard.SetText Label_RC.Caption
 clipboard.PutInClipboard
 
 Label_Copy_Values.Caption = "Rodné èíslo zkopírováno do schránky"
 Label_Flag_Copy_Values.Visible = True
 Label_Copy_Values.Visible = True
 
End If
 
End Sub

Private Sub Label_Date_Birth_Click()

If Label_Date_Birth.Caption <> vbNullString Then

Dim clipboard As MSForms.DataObject
 Set clipboard = New MSForms.DataObject
 clipboard.SetText Label_Date_Birth.Caption
 clipboard.PutInClipboard
 
 Label_Copy_Values.Caption = "Datum narození zkopírováno do schránky"
 Label_Flag_Copy_Values.Visible = True
 Label_Copy_Values.Visible = True
 
End If
 
End Sub

Private Sub Label_Adress_Permanent_Click()

If Label_Adress_Permanent.Caption <> vbNullString Then

Dim clipboard As MSForms.DataObject
 Set clipboard = New MSForms.DataObject
 clipboard.SetText Label_Adress_Permanent.Caption
 clipboard.PutInClipboard
 
 Label_Copy_Values.Caption = "Trvalá adresa zkopírována do schránky"
 Label_Flag_Copy_Values.Visible = True
 Label_Copy_Values.Visible = True
 
End If
 
End Sub

Private Sub Label_Adress_Temporary_Click()

If Label_Adress_Temporary.Caption <> vbNullString Then

Dim clipboard As MSForms.DataObject
 Set clipboard = New MSForms.DataObject
 clipboard.SetText Label_Adress_Temporary.Caption
 clipboard.PutInClipboard
 
 Label_Copy_Values.Caption = "Kontaktní adresa zkopírována do schránky"
 Label_Flag_Copy_Values.Visible = True
 Label_Copy_Values.Visible = True
 
End If
 
End Sub

Private Sub Label_Phone_Number_Click()

If Label_Phone_Number.Caption <> vbNullString Then

Dim clipboard As MSForms.DataObject
 Set clipboard = New MSForms.DataObject
 clipboard.SetText Label_Phone_Number.Caption
 clipboard.PutInClipboard
 
 Label_Copy_Values.Caption = "Tel. èíslo zkopírováno do schránky"
 Label_Flag_Copy_Values.Visible = True
 Label_Copy_Values.Visible = True
 
End If
 
End Sub

Private Sub Label_Email_Click()

If Label_Email.Caption <> vbNullString Then

Dim clipboard As MSForms.DataObject
 Set clipboard = New MSForms.DataObject
 clipboard.SetText Label_Email.Caption
 clipboard.PutInClipboard
 
 Label_Copy_Values.Caption = "E-mail zkopírován do schránky"
 Label_Flag_Copy_Values.Visible = True
 Label_Copy_Values.Visible = True
 
End If
 
End Sub










Private Sub CommandButton_Back_Click()
Me.Hide

Select Case userform_prv

'RISK-721
 Case "Zaloby"
  Klient.Show
  odkaz_Klient = False
Case "Exekuce"
  Klient_exe.Show
odkaz_Klient = True
  
 Case "Dedictvi"
  Klient_dedictvi.Show
odkaz_Klient = True
 
Case "Insol"
  Klient_insol.Show
 odkaz_Klient = True
 
Case "Trest"
  Klient_trest.Show
odkaz_Klient = True
   
  Case "Zesplatneni"
  Klient.Show
odkaz_Klient = False

 Case "Nezesplatneni"
  Klient.Show
odkaz_Klient = False

 Case "SouhrnNotice", "SouhrnAuth", "Souhrn"
  Souhrn.Show
odkaz_Klient = False

'RISK-721 Po sem
                
Case "Sonik_Datova_Zprava"
  Sonik_Datova_Zprava.Show

Case "Sonik"
  Sonik_Datova_Zprava.Show

End Select

'RISK-721 -  zapoznámkováno - nechám to co tam je
'userform_prv = "Sonik"


End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Me.Hide
Sonik_Datova_Zprava.Show
End Sub

' Zmìna barvy po najetí myši **************************************************************************************************************************

Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Back.ForeColor = barvaCerna
End Sub



Private Sub CommandButton_Back_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Back.ForeColor = barvaZelenaAirBank
End Sub




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'LISTVIEW SORTING

Private Sub ListView_Client_Actual_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_Client_Actual, ColumnHeader)
End Sub

Private Sub ListView_Client_GC_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_Client_GC, ColumnHeader)
End Sub

Private Sub ListView_Birth_Number_Change_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_Birth_Number_Change, ColumnHeader)
End Sub

Private Sub ListView_Name_Change_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_Name_Change, ColumnHeader)
End Sub

Private Sub ListView_Change_Mail_Phone_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_Change_Mail_Phone, ColumnHeader)
End Sub

Private Sub ListView_Address_Change_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_Address_Change, ColumnHeader)
End Sub

Private Sub ListView_Login_IB_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_Login_IB, ColumnHeader)
End Sub

Private Sub ListView_Accounts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_Accounts, ColumnHeader)
End Sub

Private Sub ListView_Cards_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_Cards, ColumnHeader)
End Sub

Private Sub ListView_Transactions_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_Transactions, ColumnHeader)
End Sub

Private Sub ListView_Loan_Application_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_Loan_Application, ColumnHeader)
End Sub

Private Sub ListView_Mortgage_Application_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_Mortgage_Application, ColumnHeader)
End Sub

Private Sub ListView_LAPCC_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_Manual_Verification, ColumnHeader)
End Sub

Private Sub ListView_AMGR8_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_AMGR8, ColumnHeader)
End Sub

Private Sub ListView_AMG_Note_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_AMG_Note, ColumnHeader)
End Sub

Private Sub ListView_Calls_R8_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_Calls_R8, ColumnHeader)
End Sub

Private Sub ListView_BRKI_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_BRKI, ColumnHeader)
End Sub

Private Sub ListView_Client_Communication_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_Client_Communication, ColumnHeader)
End Sub


Private Sub ListView_Transactions_Homer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_Transactions_Homer, ColumnHeader)
End Sub


















