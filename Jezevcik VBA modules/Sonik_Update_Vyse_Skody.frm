VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sonik_Update_Vyse_Skody 
   Caption         =   "Soník - Zadání výše škody"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6465
   OleObjectBlob   =   "Sonik_Update_Vyse_Skody.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sonik_Update_Vyse_Skody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit






Private Sub CommandButton_Zrus_Click()
    Call Me.Hide
End Sub

Private Sub CommandButton_Uloz_Click()
Dim sql_damage As String
Dim sql_insert_trest_note As String
Dim sql_max_skp_trest As String
Dim max_skp_trest As String
Dim id_person_for_trest As Long, sql_id_person As String 'RISK-985 Soník

Dim sql_kk_client_soucinnosti As String


 'Musi tu byt kontrolni podminky na to, co uzivatel vyplnuje
 
 
 If Not IsNumeric(TextBox_Amt_Damage.value) Or Len(Trim(TextBox_Amt_Damage.text)) = 0 Then
         MsgBox "Zkus to prosím znovu, zadej do výše škody èíselné hodnoty (haléøe oddìl desetinnou èárkou)", vbExclamation, "Chybné zadání hodnot"
 Exit Sub
 
 End If
 
  If MsgBox("Opravdu chcete uložit?", vbYesNo, Sonik_Update_Vyse_Skody.Label_Vyse_Skody_Nadpis.Caption) = vbYes Then
  
  


    If Sonik_Update_Vyse_Skody.Label_Vyse_Skody_Nadpis.Caption = "Pøidání výše škody" Then
    
    sql_damage = "insert into  " & schema & ".sonik_damage values (" & schema & ".sonik_damage_seq.nextval," & skp_person_KL & "," & Label_ID_DZ.Caption & "," & Replace(TextBox_Amt_Damage.text, ",", ".") & ", '" & TextBox_Poznamka_DZ.text & "',sysdate,'" & UCase(Main.username) & "',null,null)"
            
    Else
    
    sql_damage = "update " & schema & ".sonik_damage" & _
                        " set amt_damage = " & Replace(TextBox_Amt_Damage.text, ",", ".") & "," & _
                        " dtime_modified = sysdate," & _
                        " text_note = '" & TextBox_Poznamka_DZ.text & "'," & _
                        " name_employee_modified = '" & UCase(Main.username) & "'" & _
                        " where skp_damage = " & SKP_DAMAGE
    
    End If
    
    Call execute(sql_damage)
    
    'RISK-705 - pokud je klient v trestu, tak hned zapisu jako poznamku do trestu
    If Sonik_Karta_Klienta.CommandButton_Trest.Enabled Then
    
       sql_max_skp_trest = "select max(skp_trest_note) as max_skp_trest_note from " & schema & ".TREST_NOTE"
                
                
        Set rst1 = New ADODB.Recordset
        rst1.Open sql_max_skp_trest, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
            If Not rst1.EOF Then
                rst1.MoveFirst
                For j = 1 To rst1.RecordCount
                    max_skp_trest = "" & rst1![max_skp_trest_note] & ""
                    rst1.MoveNext
                Next j
            End If
        rst1.Close
       
       
       'RISK-985 Soník - bylo potøeba pøidat nový dotaz do tabulky person a dotáhnout si ID_PERSON
       
       sql_id_person = "select id_person  from " & schema & ".person where flag_valid = 'Y' and skp_person = " & skp_person_KL
                
                
        Set rst1 = New ADODB.Recordset
        rst1.Open sql_max_skp_trest, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
            If Not rst1.EOF Then
                rst1.MoveFirst
                For j = 1 To rst1.RecordCount
                    id_person_for_trest = rst1![id_person]
                    rst1.MoveNext
                Next j
            End If
        rst1.Close
      
      
       'RISK-985 Soník
      
        sql_insert_trest_note = "insert into  " & schema & ".trest_note values (" & max_skp_trest + 1 & ", 3 /* Poznámka */," & id_person_for_trest & ", null /* žádný stav */," & _
                               " 'SONÍK: " & TextBox_Poznamka_DZ.text & ". Èíslo datové zprávy: " & Cislo_datove_zpravy & ". Výška škody: " & Replace(TextBox_Amt_Damage.text, ",", ".") & " Kè. Platnost od ' || sysdate," & _
                               " sysdate,  '" & UCase(Main.username) & "')"
    
        Call execute(sql_insert_trest_note)
    End If

    
    
    
    
    sql_kk_client_soucinnosti = "select a.skp_data_message, a.data_message_number, a.date_data_message_delivery, a.data_message_sender, a.text_reference_number, a.text_file_number, c.name_data_message_status, e.skp_damage ,e.amt_damage, e.text_note" & _
    " from " & schema & ".sonik_data_message a" & _
    " join " & schema & ".sonik_data_message_status b on a.skp_data_message = b.skp_data_message and b.flag_valid = 'Y'" & _
    " join " & schema & ".sonik_cl_data_message_status c on b.skp_data_message_status = c.skp_data_message_status" & _
    " join " & schema & ".sonik_confirm_pair d on a.skp_data_message = d.skp_data_message and d.name_confirmation_status = 'Potvrzeno'" & _
    " left join " & schema & ".sonik_damage e on a.skp_data_message = e.skp_data_message and d.skp_person = e.skp_person" & _
    " where d.skp_person = " & skp_person_KL
    
    'A obnovim okno
    Call query_kk_client_soucinnosti(sql_kk_client_soucinnosti)
    

   End If
 
  
   Me.Hide
 


End Sub


Private Sub TextBox_Poznamka_DZ_Change()
Label32.Caption = 500 - Len(TextBox_Poznamka_DZ.text)
End Sub

' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()
'
'
'
'If cs_HDWMN.State = 1 Then
'prihlasen:
'
'
'
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



Private Sub UserForm_QueryClose(Cancel As Integer, _
                            CloseMode As Integer)
Dim ans
If CloseMode = vbFormControlMenu Then
    Cancel = True
    Call Me.Hide
End If
End Sub

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



