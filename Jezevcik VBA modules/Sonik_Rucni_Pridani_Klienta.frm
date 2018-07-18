VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sonik_Rucni_Pridani_Klienta 
   Caption         =   "Soník - Zadávání subjektù"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6750
   OleObjectBlob   =   "Sonik_Rucni_Pridani_Klienta.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sonik_Rucni_Pridani_Klienta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit


'Potvrzení a zavøení

Private Sub CommandButton_Uloz_Click()
Dim sql_rucne_Hl_Subj_CUID As String
Dim itmX As ListItem

'Kontroly, jestli uživatel zadal správný formát - èíslo tam kam mìl
'CUID
If Not IsNumeric(TextBox_Rucni_CUID.value) And TextBox_Rucni_CUID.value <> vbNullString Then
         MsgBox "Zkus to prosím znovu, zadej do cuidu pouze èíselné hodnoty", vbExclamation, "Chybné zadání hodnot"
         TextBox_Rucni_CUID.value = vbNullString
Exit Sub
End If


'XXX Musím to teï vložit zpìt do potvrzovací tabulky

'Než zaènu hledat, tak uložím subjekt do DTB
Dim sql_max_skp As String

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
'RISK-876 XXXX
sql_insert_subject = "insert into  " & schema & ".sonik_subject values (" & SKP_Subject & ", " & SKP_datove_zpravy & "," & "'" & TextBox_Rucni_CUID.value & "'" & ",null,null," & "to_date('01.01.3000','DD.MM.YY')" & ",null,null,null,null," & "sysdate" & "," & "'" & UCase(Main.username) & "',null)"


Call execute(sql_insert_subject)
'RISK-876
sql_rucne_Hl_Subj_CUID = "select skp_person, ID_CUID, NAME_FIRST, NAME_LAST, DATE_BIRTH, TEXT_BIRTH_NUMBER, name_permanent_address_street ||' '|| text_permanent_address_num ||', '|| name_permanent_address_town ||', '|| code_permanent_address_zip as PERMANENT_ADDRESS, TEXT_BANK_ACCOUNT_NUMBER_AGG, CODE_LOAN_NUMBER_AGG, TEXT_PHONE_NUMBER_AGG from " & schema & ".sonik_ab_klient_dtb where ID_CUID = '" & TextBox_Rucni_CUID.value & "'"

'RISK-876
Dim keys As Dictionary
Set keys = KliceKListview(Sonik_Potvrzeni_Subjektu.ListView_Potvrzeni)

'chci jen pridat do tabulky novy cuid
   Set rst1 = New ADODB.Recordset
   rst1.Open sql_rucne_Hl_Subj_CUID, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
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
Me.TextBox_Rucni_CUID.text = ""
Sonik_Rucni_Pridani_Klienta.Hide

End Sub




' Zavøení formuláøe
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Call set_filter(Me.name, 1)
Me.Hide
Sonik_Potvrzeni_Subjektu.Show
End Sub


' Back button
Private Sub CommandButton_Navrat_Click()
Call text_information_hide(formular)
Call UserForm_QueryClose(1, 0)
End Sub

















' Zmìna barvy po najetí myši **************************************************************************************************************************

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Uloz.ForeColor = barvaCerna
End Sub

Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Navrat.ForeColor = barvaCerna
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Uloz.ForeColor = barvaCerna
End Sub

Private Sub CommandButton_Navrat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Navrat.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Uloz_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Uloz.ForeColor = barvaZelenaAirBank
End Sub

