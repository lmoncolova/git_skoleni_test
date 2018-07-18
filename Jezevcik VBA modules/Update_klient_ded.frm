VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Update_klient_ded 
   Caption         =   "Jezevèík - Dìdické øízení"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6345
   OleObjectBlob   =   "Update_klient_ded.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Update_klient_ded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit

'RISK-468


Private Sub CommandButton12_Click()
    Call Me.Hide
End Sub

Private Sub CommandButton20_Click()
Dim sql As String
Dim sqlPERSON As String, BlnPerson As Boolean
Dim sqlDEMOGRAPHY As String, BlnDemography As Boolean
Dim zmena As String



If cs_HDWMN.State = 1 Then
prihlasen:

Call text_information_hide(Me.name)

BlnPerson = False
BlnDemography = False
zmena = ""


sqlPERSON = "update " & schema & ".person set "
sqlDEMOGRAPHY = "update " & schema & ".person_demography set "

If Len(Trim(TextBox1.text)) > 0 And CheckBox1.value Then
 sqlPERSON = sqlPERSON & "name = '" & TextBox1.text & "', "
 BlnPerson = True
 zmena = zmena & "jméno (" & TextBox1.text & "), "
End If

If Len(Trim(TextBox2.text)) > 0 And CheckBox2.value Then
 
     'Kontrola na správné zadání RÈ
    If TextBox2.value <> vbNullString And (Not IsNumeric(TextBox2.value) Or TextBox2.value <= 0 Or Len(TextBox2.value) < 9 Or Len(TextBox2.value) > 10) Then
            MsgBox "Rodné èíslo není vyplnìno správnì.", vbExclamation, "Chybné zadání hodnot"
            Exit Sub
    End If
 
 sqlDEMOGRAPHY = sqlDEMOGRAPHY & "text_birth_number = '" & TextBox2.text & "' "
 BlnDemography = True
 zmena = zmena & "rè (" & TextBox2.text & "), "
End If


sqlPERSON = sqlPERSON & " dtime_modification = sysdate, employee = '" & Main.username & "' where id_person = " & id_person
sqlDEMOGRAPHY = sqlDEMOGRAPHY & " where id_person = " & id_person
 
If BlnPerson Or BlnDemography Then
 If MsgBox("Opravdu chcete uložit?", vbYesNo, "Uložit") = vbYes Then
  If BlnPerson Then
   execute (sqlPERSON)
  End If
  
  If BlnDemography Then
   execute (sqlDEMOGRAPHY)
  End If

  
  sql = "insert into " & schema & ".ded_note values (" & max_skp("select max(skp_ded_note) from " & schema & ".ded_note") & ", 2, " & id_person & _
      ", NULL, 'Zmìna údajù klienta: " & Mid(zmena, 1, Len(zmena) - 2) & "', sysdate, '" & Main.username & "')"
  
  execute (sql)
  
  Me.Hide
  Klient_dedictvi.Show
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
TextBox1.Visible = CheckBox1.value

TextBox1.text = c.name
End Sub

Private Sub CheckBox2_Click()
TextBox2.Visible = CheckBox2.value

TextBox2.text = c.Pin
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
Dim sql As String
Dim control As control

If cs_HDWMN.State = 1 Then
prihlasen:

Call text_information_hide(Me.name)

For Each control In Me.Controls
 If control.Tag = "1" Then control.Visible = True
 If control.Tag = "2" Then control.Visible = False
 If control.Tag = "3" Then control.value = False
Next


'RISK-517 - dovolím zmìnu Rodného èísla pouze na neklientech
If c.CUID <= 0 Then
    'CheckBox1.Enabled = True 'Jméno
    CheckBox2.Enabled = True 'RÈ
    'Frame3.ControlTipText = "" 'Frame se jménem
    Frame4.ControlTipText = "" 'Frame se RÈ
    
Else: 'CheckBox1.Enabled = False
      CheckBox2.Enabled = False
      'Frame3.ControlTipText = "Jméno lze mìnit pouze u neklientù - u klientù pøes OBS"
      Frame4.ControlTipText = "Rodné èíslo lze mìnit pouze u neklientù - u klientù pøes OBS"
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


