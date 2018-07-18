VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sonik_Update_Stav_DZ 
   Caption         =   "Soník - Datová zpráva"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10290
   OleObjectBlob   =   "Sonik_Update_Stav_DZ.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sonik_Update_Stav_DZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit




Private Sub CommandButton_Zrus_Click()
    Call Me.Hide
End Sub

Private Sub CommandButton_Back_Click()
    Call Me.Hide
End Sub

Private Sub CommandButton_Uloz_Click()
Dim sql_max_id_dz As String
Dim sql_max_id_status As String

'Tady tento kus kódu musí kromì pøidání nového stavu i ukonèit platnost toho minulého


'Prvnì stavy - potrebuji nejvyssi skp na dane soucinnosti - kvuli updateu

'Zjistim maximalni id pro skp DZ
sql_max_id_dz = "select nvl(max(id),0) as max_id_dz from " & schema & ".sonik_data_message_status where skp_data_message = " & SKP_datove_zpravy

Dim max_id_status_dz As String

Dim sql_update_status As String
Dim sql_insert_status As String
Dim sql_note As String

   Set rst1 = New ADODB.Recordset
   rst1.Open sql_max_id_dz, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    max_id_status_dz = "" & rst1![max_id_dz] & ""
    End If

   rst1.Close




'Jenom pojistka pro první spuštìní
If max_id_status_dz >= 1 Then
Else
max_id_status_dz = 0
End If




sql_update_status = "update " & schema & ".sonik_data_message_status" & _
" set dtime_valid_to = sysdate - 1/86400, flag_valid = 'N' where id = " & max_id_status_dz


sql_insert_status = "insert into " & schema & ".sonik_data_message_status" & _
" select " & schema & ".sonik_data_message_status_seq.nextval ," & SKP_datove_zpravy & ", SKP_DATA_MESSAGE_STATUS,sysdate,'" & UCase(Main.username) & "',sysdate,date'3000-01-01','Y'" & _
" from " & schema & ".sonik_cl_data_message_status where NAME_DATA_MESSAGE_STATUS = '" & Me.ComboBox_Stavy.value & "'"



'Updateuju ten poslední, pokud tam žádný není, tak není co updateovat
If max_id_status_dz > 0 Then
  Call execute(sql_update_status)
End If

  Call execute(sql_insert_status)
  
  
'Automatická upozornìní

Call status_change_notices(Me.ComboBox_Stavy.value)



'Teï ještì uložím tu poznámku - pokud tam tedy nìjaká je...


If TextBox_Poznamka.text <> vbNullString Then


'Zjistim maximalni skp pro navazaní na ten status, který se právì nahrál...
'Tady to opravdu musim zjistit, abych spravne navazal poznamku na id zmeny stavu, nejde to pres sequence
sql_max_id_status = "select nvl(max(id),0) as max_id_status from " & schema & ".sonik_data_message_status"

Dim max_id_status As String

Set rst1 = New ADODB.Recordset
   rst1.Open sql_max_id_status, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    max_id_status = "" & rst1![max_id_status] & ""
    End If

   rst1.Close

 'Jenom pojistka pro první spuštìní
If max_id_status >= 1 Then
Else
max_id_status = 0
End If
  
  

 sql_note = "insert into " & schema & ".sonik_data_message_note values  (" & schema & ".sonik_data_message_note_seq.nextval, " & SKP_datove_zpravy & "," & max_id_status & ", '" & TextBox_Poznamka.text & "',sysdate,'" & UCase(Main.username) & "')"


Call execute(sql_note)

End If
  
  
  




  'Aktualizuju hlavni stranku DZ
  Call nacteni_dat_Sonik_Datova_Zprava
  
TextBox_Poznamka.text = ""

Me.Hide

End Sub





Private Sub TextBox_Poznamka_Change()
Label32.Caption = 500 - Len(TextBox_Poznamka.text)
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
CommandButton_Zrus.ForeColor = barvaCerna
CommandButton_Uloz.ForeColor = barvaCerna

End Sub

Private Sub CommandButton_Zrus_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Zrus.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Uloz_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Uloz.ForeColor = barvaZelenaAirBank
End Sub

Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Back.ForeColor = barvaCerna
End Sub

Private Sub CommandButton_Back_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Back.ForeColor = barvaZelenaAirBank
End Sub


