VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Update_klient_insol 
   Caption         =   "Jezevèík - Insolvenèní øízení"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11220
   OleObjectBlob   =   "Update_klient_insol.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Update_klient_insol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit






Private Sub CommandButton12_Click()
    Call Me.Hide
End Sub

Private Sub CommandButton20_Click()
Dim sql As String
Dim sqlBASE As String
Dim sqlINFO As String
'Dim sqlINFO2 As String
Dim BeforeValue As String
Dim AfterValue As String
'Dim BeforeValue2 As String
'Dim AfterValue2 As String
Dim id As Integer
Dim zmena As String
Dim base As Boolean
Dim info As Boolean
'Dim info2 As Boolean
'Dim max_heir As Integer
'Dim max_notary As Integer
Dim max_ins_info As Integer

Dim sqlDEMOGRAPHY As String, BlnDemography As Boolean

If cs_HDWMN.State = 1 Then
prihlasen:

Call text_information_hide(Me.name)

base = False
info = False
'info2 = False
zmena = ""


sqlBASE = "update " & schema & ".person set "
sqlDEMOGRAPHY = "update " & schema & ".person_demography set " 'RISK-517 DONE

If Len(Trim(TextBox1.text)) > 0 And CheckBox1.value And TextBox1.text <> Label1.Caption Then
 sqlBASE = sqlBASE & "name = '" & TextBox1.text & "', "
 base = True
 zmena = zmena & "jméno (" & TextBox1.text & "), "
End If

If Len(Trim(TextBox2.text)) > 0 And CheckBox2.value And TextBox2.text <> Label2.Caption Then
    
    'Kontrola na správné zadání RÈ
    If TextBox2.value <> vbNullString And (Not IsNumeric(TextBox2.value) Or TextBox2.value <= 0 Or Len(TextBox2.value) < 9 Or Len(TextBox2.value) > 10) Then
            MsgBox "Rodné èíslo není vyplnìno správnì.", vbExclamation, "Chybné zadání hodnot"
            Exit Sub
    End If
    
    sqlDEMOGRAPHY = sqlDEMOGRAPHY & "text_birth_number = '" & TextBox2.text & "' " 'RISK-517 DONE
    BlnDemography = True
    zmena = zmena & "rè (" & TextBox2.text & "), "
End If

sqlDEMOGRAPHY = sqlDEMOGRAPHY & " where id_person = " & id_person

sqlBASE = sqlBASE & " dtime_modification = sysdate, employee = '" & Main.username & "' where id_person = " & id_person



If prava("select * from " & UCase(schema) & ".INS_INFO where id_person = " & id_person) Then
sqlINFO = "update " & schema & ".ins_info set "


 If CheckBox6.value And DTPicker6.value <> Label6.Caption Then
  sqlINFO = sqlINFO & "datum_upadku = to_date('" & DTPicker6.value & "','dd.mm.yyyy'), "
  info = True
  zmena = zmena & "datum úpadku (" & DTPicker6.value & "), "
 End If
 
  If CheckBox7.value And DTPicker7.value <> Label7.Caption Then
  sqlINFO = sqlINFO & "datum_prihlaseni = to_date('" & DTPicker7.value & "','dd.mm.yyyy'), "
  info = True
  zmena = zmena & "datum pøihlašení (" & DTPicker7.value & "), "
 End If
 

sqlINFO = sqlINFO & " dtime_modification = sysdate, employee = '" & Main.username & "' where id_person = " & id_person


Else

MsgBox ("Nìkde nastal problém. Informujte vývojáøe!")

'max_ins_info = max_skp("select max(skp_ins_info) from " & schema & ".ins_info")
'
' BeforeValue = " skp_ins_info ,id_person, "
' AfterValue = "(" & max_ins_info & " ," & id_person & ", "
' sqlINFO = "insert into " & schema & ".ins_info ("
'
'' If CheckBox6.Value And DTPicker1.Value <> Label6.Caption Then
''  BeforeValue = BeforeValue & "date_action, "
''  AfterValue = AfterValue & "to_date('" & DTPicker1.Value & "','dd.mm.yyyy'), "
''  info = True
''  zmena = zmena & "datum žaloby (" & DTPicker1.Value & "), "
'' End If
'
'   If Len(Trim(TextBox11.Text)) > 0 And CheckBox11.Value And TextBox11.Text <> Label11.Caption Then
'  BeforeValue = BeforeValue & "code, "
'  AfterValue = AfterValue & "'" & TextBox11.Text & "', "
'  info = True
'  zmena = zmena & "sp.zn.(" & TextBox11.Text & "), "
' End If
'
'
'BeforeValue = BeforeValue & "dtime_modification, employee)"
'AfterValue = AfterValue & "sysdate, '" & Main.username & "')"
'sqlINFO = sqlINFO & BeforeValue & " values " & AfterValue

' If CheckBox7.Value And TextBox7.Text <> Label7.Caption Then
'  If Len(Trim(TextBox7.Text)) > 0 And IsNumeric(TextBox7.Text) Then
'   BeforeValue = BeforeValue & "amt_claim, "
'   AfterValue = AfterValue & Replace(Replace(TextBox7.Text, ",", "."), " ", "") & ", "
'   info = True
'   zmena = zmena & "pohledávka (" & TextBox7.Text & "), "
'  Else
'   GoTo chyba
'  End If
' End If

' If CheckBox8.Value And DTPicker2.Value <> Label8.Caption Then
'  BeforeValue = BeforeValue & "date_calculation_claim, "
'  AfterValue = AfterValue & "to_date('" & DTPicker2.Value & "','dd.mm.yyyy'), "
'  info = True
'  zmena = zmena & "datum pohledávky (" & DTPicker2.Value & "), "
' End If
'
' If Len(Trim(TextBox9.Text)) > 0 And CheckBox9.Value And TextBox9.Text <> Label9.Caption Then
'  BeforeValue = BeforeValue & "id_inbox, "
'  AfterValue = AfterValue & "'" & TextBox9.Text & "', "
'  info = True
'  zmena = zmena & "ID Datové schránky (" & TextBox9.Text & "), "
' End If

End If



If base Or info Or BlnDemography Then 'RISK-517 DONE

'
'a = InputBox("", "", sqlINFO)
'a = InputBox("", "", sqlINFO2)

 If MsgBox("Opravdu chcete uložit?", vbYesNo, "Uložit") = vbYes Then
  If base Then
   execute (sqlBASE)
  End If

  If BlnDemography Then
   execute (sqlDEMOGRAPHY) 'RISK-517 DONE
  End If

  If info Then
   execute (sqlINFO)
  End If
  
  
  
  sql = "insert into " & schema & ".ins_note values (" & max_skp("select max(skp_ins_note) from " & schema & ".ins_note") & ", 2, " & id_person & _
      ", NULL, 'Zmìna údajù klienta: " & Mid(zmena, 1, Len(zmena) - 2) & "', sysdate, '" & Main.username & "')"
  
  execute (sql)
  
  Me.Hide
  Klient_insol.Show
      odkaz_Klient = True
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


Private Sub CheckBox6_Click()
Label6.Visible = Not CheckBox6.value
DTPicker6.Visible = CheckBox6.value

If CheckBox6.value And Label6.Caption <> "" Then
DTPicker6 = Label6.Caption
ElseIf CheckBox6.value Then
DTPicker6 = date
End If

Me.Repaint
End Sub


Private Sub CheckBox7_Click()
Label7.Visible = Not CheckBox7.value
DTPicker7.Visible = CheckBox7.value

If CheckBox7.value And Label7.Caption <> "" Then
DTPicker7 = Label7.Caption
ElseIf CheckBox7.value Then
DTPicker7 = date
End If

Me.Repaint
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

Me.Frame7.ControlTipText = "Pro úpravu datumu úpadku musíš napsat svému oblíbenému analytikovi." 'RISK-1038 (na formu enabled = false)


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


