VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Update_klient_trest 
   Caption         =   "Jezevèík - Trestní øízení"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11220
   OleObjectBlob   =   "Update_klient_trest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Update_klient_trest"
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
Dim max_trest_info As Integer

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



If prava("select * from " & UCase(schema) & ".trest_info where id_person = " & id_person) Then
sqlINFO = "update " & schema & ".trest_info set "

  If Len(Trim(TextBox10.text)) > 0 And CheckBox10.value And TextBox10.text <> Label10.Caption Then
  sqlINFO = sqlINFO & "amt_damage  = " & Replace(Replace(TextBox10.text, ",", "."), " ", "") & ", "
  info = True
  zmena = zmena & "výše škody(" & Replace(Replace(TextBox10.text, ".", ","), " ", "") & "), "
 End If
 
  If Len(Trim(TextBox11.text)) > 0 And CheckBox11.value And TextBox11.text <> Label11.Caption Then
  sqlINFO = sqlINFO & "code  = '" & TextBox11.text & "', "
  info = True
  zmena = zmena & "sp.zn.(" & TextBox11.text & "), "
 End If
 
 If Len(Trim(TextBox12.text)) > 0 And CheckBox12.value And TextBox12.text <> Label12.Caption Then
  sqlINFO = sqlINFO & "octr  = '" & TextBox12.text & "', "
  info = True
  zmena = zmena & "oètø(" & TextBox12.text & "), "
 End If


sqlINFO = sqlINFO & " dtime_modification = sysdate, employee = '" & Main.username & "' where id_person = " & id_person


Else

max_trest_info = max_skp("select max(skp_trest_info) from " & schema & ".trest_info")

 BeforeValue = " skp_trest_info ,id_person, "
 AfterValue = "(" & max_trest_info & " ," & id_person & ", "
 sqlINFO = "insert into " & schema & ".trest_info ("
 
' If CheckBox6.Value And DTPicker1.Value <> Label6.Caption Then
'  BeforeValue = BeforeValue & "date_action, "
'  AfterValue = AfterValue & "to_date('" & DTPicker1.Value & "','dd.mm.yyyy'), "
'  info = True
'  zmena = zmena & "datum žaloby (" & DTPicker1.Value & "), "
' End If
 
   If Len(Trim(TextBox10.text)) > 0 And CheckBox10.value And TextBox10.text <> Label10.Caption Then
  BeforeValue = BeforeValue & "amt_damage, "
  AfterValue = AfterValue & Replace(Replace(TextBox10.text, ",", "."), " ", "") & ", "
  info = True
  zmena = zmena & "výše škody(" & Replace(Replace(TextBox10.text, ".", ","), " ", "") & "), "
 End If
 
   If Len(Trim(TextBox11.text)) > 0 And CheckBox11.value And TextBox11.text <> Label11.Caption Then
  BeforeValue = BeforeValue & "code, "
  AfterValue = AfterValue & "'" & TextBox11.text & "', "
  info = True
  zmena = zmena & "sp.zn.(" & TextBox11.text & "), "
 End If
 
    If Len(Trim(TextBox12.text)) > 0 And CheckBox12.value And TextBox12.text <> Label12.Caption Then
  BeforeValue = BeforeValue & "octr, "
  AfterValue = AfterValue & "'" & TextBox12.text & "', "
  info = True
  zmena = zmena & "oètø(" & TextBox12.text & "), "
 End If
 

BeforeValue = BeforeValue & "dtime_modification, employee)"
AfterValue = AfterValue & "sysdate, '" & Main.username & "')"
sqlINFO = sqlINFO & BeforeValue & " values " & AfterValue

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
    
    Call TransBegin
    
        If base Then
            execute (sqlBASE)
        End If
        
        If BlnDemography Then
            execute (sqlDEMOGRAPHY) 'RISK-517 DONE
        End If
        
        If info Then
            execute (sqlINFO)
        End If
        
        
        
        sql = "insert into " & schema & ".trest_note values (" & max_skp("select max(skp_trest_note) from " & schema & ".trest_note") & ", 2, " & id_person & _
            ", NULL, 'Zmìna údajù klienta: " & Mid(zmena, 1, Len(zmena) - 2) & "', sysdate, '" & Main.username & "')"
        
        execute (sql)
        
    Call TransCommit
    
    Me.Hide
    Klient_trest.Show
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
''
''Private Sub CheckBox3_Click()
''Label3.Visible = Not CheckBox3.value
''TextBox3.Visible = CheckBox3.value
''
''TextBox3.text = Label3.Caption
''End Sub
''
''Private Sub CheckBox4_Click()
''Label4.Visible = Not CheckBox4.value
''TextBox4.Visible = CheckBox4.value
''
''TextBox4.text = Label4.Caption
''End Sub
''
''Private Sub CheckBox5_Click()
''Label5.Visible = Not CheckBox5.value
''TextBox5.Visible = CheckBox5.value
''
''TextBox5.text = Label5.Caption
''End Sub
''
''Private Sub CheckBox6_Click()
''Label6.Visible = Not CheckBox6.value
''TextBox6.Visible = CheckBox6.value
''
''TextBox6.text = Label6.Caption
''
''Me.Repaint
''End Sub
''
''
''Private Sub CheckBox7_Click()
''Label7.Visible = Not CheckBox7.value
''TextBox7.Visible = CheckBox7.value
''
''TextBox7.text = Label7.Caption
''End Sub
''
''Private Sub CheckBox8_Click()
''Label8.Visible = Not CheckBox8.value
''TextBox8.Visible = CheckBox8.value
''
''TextBox8.text = Label8.Caption
''End Sub
''
''Private Sub CheckBox9_Click()
''Label9.Visible = Not CheckBox9.value
''TextBox9.Visible = CheckBox9.value
''
''TextBox9.text = Label9.Caption
''End Sub

Private Sub CheckBox10_Click()
Label10.Visible = Not CheckBox10.value
TextBox10.Visible = CheckBox10.value

TextBox10.text = Label10.Caption
End Sub

Private Sub CheckBox11_Click()
Label11.Visible = Not CheckBox11.value
TextBox11.Visible = CheckBox11.value

TextBox11.text = Label11.Caption
End Sub

Private Sub CheckBox12_Click()
Label12.Visible = Not CheckBox12.value
TextBox12.Visible = CheckBox12.value

TextBox12.text = Label12.Caption
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


