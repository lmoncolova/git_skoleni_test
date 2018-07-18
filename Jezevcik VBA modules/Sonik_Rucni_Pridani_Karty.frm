VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sonik_Rucni_Pridani_Karty 
   Caption         =   "Soník - Zadávání karet"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6750
   OleObjectBlob   =   "Sonik_Rucni_Pridani_Karty.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sonik_Rucni_Pridani_Karty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit


Private Sub UserForm_Activate()

TextBox_Rucni_Karty.value = Sonik_Datova_Zprava.TextBox_Card.value 'Hl_Card

End Sub


'Potvrzení a zavøení


Private Sub CommandButton_Uloz_Kartu_Click()

Dim sql_max_skp As String
Dim sql_rucne_Hl_Subj_Karta As String
Dim itmX As ListItem

'Kontroly, jestli uživatel zadal správný formát - èíslo tam kam mìl
'Èíslo karty
If Not IsNumeric(TextBox_Rucni_Karty.value) And TextBox_Rucni_Karty.value <> vbNullString Then
         MsgBox "Zkus to prosím znovu, zadej do èísla karty pouze èíselné hodnoty", vbExclamation, "Chybné zadání hodnot"
         TextBox_Rucni_Karty.value = vbNullString
Exit Sub
ElseIf Len(TextBox_Rucni_Karty.value) <> 16 And TextBox_Rucni_Karty.value <> vbNullString Then
         MsgBox "Zkus to prosím znovu, èíslo karty musí obsahovat právì 16 znakù", vbExclamation, "Chybné zadání hodnot"
         TextBox_Rucni_Karty.value = vbNullString

Exit Sub
End If

'MsgBox "Vykoèí okno s informací, zda karta byla, èi nebyla použita na našich bankomatech. Tlaèítko OK -> èíslo karty se uloží do databáze a zaènou se dotahovat data, jako pro klienta, dokud se DZ neuzavøe."
'
'TextBox_Rucni_Karty.value = vbNullString
'Hl_Card = vbNullString

'Ulozim do globální promìnné
'Rucne_Pridana_Karta = TextBox_Rucni_Karty.value


'XXX Musím to teï vložit zpìt do potvrzovací tabulky

'Než zaènu hledat, tak uložím subjekt do DTB


'Zjistim maximalni skp - skute
sql_max_skp = "select nvl(max(skp_subject),0) as max_skp from " & schema & ".sonik_subject"

Dim max_skp As String
Dim sql_insert_subject As String

   Set rst1 = New ADODB.Recordset
   rst1.Open sql_max_skp, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    max_skp = "" & rst1![max_skp] & ""
    End If

   rst1.Close


'Jenom pojistka pro první spuštìní
If max_skp >= 1 Then
Else
max_skp = 0
End If
SKP_Subject = max_skp + 1




'Nechavam skript protoze ty ostatni atributy budou prazdne...

sql_insert_subject = "insert into  " & schema & ".sonik_subject values (" & SKP_Subject & ", " & SKP_datove_zpravy & ",null,null,null," & "to_date('01.01.3000','DD.MM.YY')" & ",null,null,'" & TextBox_Rucni_Karty.value & "',null," & "sysdate" & "," & "'" & UCase(Main.username) & "',null)"

Call execute(sql_insert_subject)

sql_rucne_Hl_Subj_Karta = "select skp_person, ID_CUID, NAME_FIRST, NAME_LAST, DATE_BIRTH, TEXT_BIRTH_NUMBER, name_permanent_address_street ||' '|| text_permanent_address_num ||', '|| name_permanent_address_town ||', '|| code_permanent_address_zip as PERMANENT_ADDRESS, TEXT_BANK_ACCOUNT_NUMBER_AGG, CODE_LOAN_NUMBER_AGG, TEXT_PHONE_NUMBER_AGG from " & schema & ".sonik_ab_klient_dtb where skp_person = -1"

'RISK-876
Dim keys As Dictionary
Set keys = KliceKListview(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni)

'chci jen pridat do tabulky novy cuid
   Set rst1 = New ADODB.Recordset
   rst1.Open sql_rucne_Hl_Subj_Karta, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   'RISK-876
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni.ListItems.Add()
      itmX.text = SKP_Subject 'Vezmu pøímo z excelu
      itmX.SubItems(keys("ID_CUID")) = "" & rst1![id_cuid] & ""
      itmX.SubItems(keys("Jméno")) = "" & rst1![Name_first] & ""
      itmX.SubItems(keys("Pøíjmení")) = "" & rst1![Name_last] & ""
      itmX.SubItems(keys("Datum narození")) = "" & rst1![Date_Birth] & ""
      itmX.SubItems(keys("Rodné èíslo")) = "" & rst1![TEXT_BIRTH_NUMBER] & ""
      itmX.SubItems(keys("Trvalá adresa")) = "" & rst1![PERMANENT_ADDRESS] & ""
      itmX.SubItems(keys("Èísla úètù")) = "" & rst1![TEXT_BANK_ACCOUNT_NUMBER_AGG] & ""
      itmX.SubItems(keys("Èísla úvìrù")) = "" & rst1![CODE_LOAN_NUMBER_AGG] & ""
      itmX.SubItems(keys("Potvrzení")) = "Nepotvrzeno"
      itmX.SubItems(keys("skp_person")) = "" & rst1![skp_person] & ""
    rst1.MoveNext
    Next i
   End If

   rst1.Close




'Naplnim tady takovy flag z toho dùvodu, že v momentì kdy už ruènì pøidám klienta pøes CUID, tak mi to nedovolí zadat negativku - pak totiž neumím pracovat s klíèi a stejnì je to k nièemu
Flag_Existuje_Rucne_Pridany_Klient = 1

'Po zadání promažu
Me.TextBox_Rucni_Karty.text = ""
Sonik_Rucni_Pridani_Karty.Hide

End Sub



' Zavøení formuláøe
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Call set_filter(Me.name, 1)
Me.Hide
Sonik_Potvrzeni_Subjektu.Show
End Sub


' Back button
Private Sub CommandButton_Karty_Navrat_Click()
Call text_information_hide(formular)
Call UserForm_QueryClose(1, 0)
End Sub














' Zmìna barvy po najetí myši **************************************************************************************************************************

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Uloz_Kartu.ForeColor = barvaCerna
End Sub

Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Karty_Navrat.ForeColor = barvaCerna
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Uloz_Kartu.ForeColor = barvaCerna
End Sub

Private Sub CommandButton_Karty_Navrat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Karty_Navrat.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Uloz_Kartu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Uloz_Kartu.ForeColor = barvaZelenaAirBank
End Sub


