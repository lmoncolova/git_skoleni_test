VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Update_auth 
   Caption         =   "Jezevèík v 1.0 - Autorizace"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12840
   OleObjectBlob   =   "Update_auth.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Update_auth"
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
Dim BeforeValue As String
Dim AfterValue As String
Dim id As Integer
Dim zmena As String
Dim start_query As Boolean

If cs_HDWMN.State = 1 Then
prihlasen:

Call text_information_hide(Me.name)

start_query = False
zmena = ""

sql = "update " & schema & ".sop set "
 
 If CheckBox1.value And CheckBox1.Enabled Then
   sql = sql & "DATE_SOP1 = to_date('" & DTPicker1.value & "','dd.mm.yyyy'), "
   start_query = True
   zmena = zmena & "SOP1 (" & Label1.Caption & "/" & DTPicker1.value & "), "
 End If
 
 If CheckBox2.value And CheckBox2.Enabled Then
   sql = sql & "DATE_ADDITIONAL_SOP1 = to_date('" & DTPicker2.value & "','dd.mm.yyyy'), "
   start_query = True
   zmena = zmena & "doplatek SOP1 (" & Label2.Caption & "/" & DTPicker2.value & "), "
 End If

 If CheckBox3.value And CheckBox3.Enabled Then
  If IsNumeric(TextBox3.text) Or Len(Trim(TextBox3.text)) > 0 Then
   sql = sql & "AMT_RETURN_REAL_SOP1 = " & Replace(TextBox3.text, ",", ".") & ", DATE_RETURN_REAL_SOP1 = to_date('" & DTPicker3.value & "','dd.mm.yyyy'), "
   start_query = True
   zmena = zmena & "pøeplatek SOP1 (" & Label31.Caption & "/" & TextBox3.text & "," & Label32.Caption & "/" & DTPicker3.value & "),  "
  Else
   GoTo chyba
  End If
 End If

  If CheckBox4.value And CheckBox4.Enabled Then
   sql = sql & "DATE_SOP2 = to_date('" & DTPicker4.value & "','dd.mm.yyyy'), "
   start_query = True
   zmena = zmena & "SOP2 (" & Label4.Caption & "/" & DTPicker4.value & "), "
 End If
 
 If CheckBox5.value And CheckBox5.Enabled Then
   sql = sql & "DATE_ADDITIONAL_SOP2 = to_date('" & DTPicker5.value & "','dd.mm.yyyy'), "
   start_query = True
   zmena = zmena & "doplatek SOP2 (" & Label5.Caption & "/" & DTPicker5.value & "), "
 End If

 If CheckBox6.value And CheckBox6.Enabled Then
  If IsNumeric(TextBox6.text) Or Len(Trim(TextBox6.text)) > 0 Then
   sql = sql & "AMT_RETURN_REAL_SOP2 = " & Replace(TextBox6.text, ",", ".") & ", DATE_RETURN_REAL_SOP2 = to_date('" & DTPicker6.value & "','dd.mm.yyyy'), "
   start_query = True
   zmena = zmena & "pøeplatek SOP2 (" & Label61.Caption & "/" & TextBox6.text & "," & Label62.Caption & "/" & DTPicker6.value & "),  "
  Else
   GoTo chyba
  End If
 End If
 
 
 sql = sql & " dtime_modification = sysdate, employee = '" & Main.username & "' where id_person = " & id_person


If start_query Then
 If True Then ' If MsgBox("Opravdu chcete autorizovat?", vbYesNo, "Autorizovat") = vbYes Then
    TransBegin
        execute (sql)
        sql = "insert into " & schema & ".epr_note values (" & max_skp("select max(skp_epr_note) from " & schema & ".epr_note") & ", " & id_person & _
            ", 5, NULL, 'Autorizace dat: " & Mid(zmena, 1, Len(zmena) - 2) & "', sysdate, '" & Main.username & "')"
        execute (sql)
    TransCommit
    
    c.SOP.re_initialize
    Me.Hide
    Klient.Show
    
 Else
    Label28.Visible = True
    Label29.Visible = True
    Label29.Caption = "Zrušeno"
 End If
 
Else
 Label28.Visible = True
 Label29.Visible = True
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
 Label28.Visible = True
 Label29.Visible = True
 Label29.Caption = "Chybný formát vstupních dat"
End Sub





Private Sub CheckBox1_Click()
 If CheckBox1.value Then
  Update_auth.Frame1.Height = 60
 Else
  Update_auth.Frame1.Height = 24
 End If
End Sub

Private Sub CheckBox2_Click()
 If CheckBox2.value Then
  Update_auth.Frame2.Height = 60
 Else
  Update_auth.Frame2.Height = 24
 End If
End Sub


Private Sub CheckBox3_Click()
 If CheckBox3.value Then
  Update_auth.Frame3.Height = 108
 Else
  Update_auth.Frame3.Height = 48
 End If
End Sub

Private Sub CheckBox4_Click()
 If CheckBox4.value Then
  Update_auth.Frame4.Height = 60
 Else
  Update_auth.Frame4.Height = 24
 End If
End Sub

Private Sub CheckBox5_Click()
 If CheckBox5.value Then
  Update_auth.Frame5.Height = 60
 Else
  Update_auth.Frame5.Height = 24
 End If
End Sub

Private Sub CheckBox6_Click()
 If CheckBox6.value Then
  Update_auth.Frame6.Height = 108
 Else
  Update_auth.Frame6.Height = 48
 End If
End Sub




' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()
Dim sql As String
Dim control As control

If cs_HDWMN.State = 1 Then
prihlasen:

Update_auth.Caption = "Jezevèík v." & version & " - autorizace"

Call text_information_hide(Me.name)

If prava("select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'AUTHORIZATION' and PRIVILEGE = 'INSERT'") Then

 CommandButton20.Enabled = True

Else
 
 CommandButton20.Enabled = False
 
 If Update_auth.Frame1.BackColor = barvaLososova Then
  CheckBox1.Enabled = False
 End If
 If Update_auth.Frame2.BackColor = barvaLososova Then
  CheckBox2.Enabled = False
 End If
 If Update_auth.Frame3.BackColor = barvaLososova Then
  CheckBox3.Enabled = False
 End If
 If Update_auth.Frame4.BackColor = barvaLososova Then
  CheckBox4.Enabled = False
 End If
 If Update_auth.Frame5.BackColor = barvaLososova Then
  CheckBox5.Enabled = False
 End If
 If Update_auth.Frame6.BackColor = barvaLososova Then
  CheckBox6.Enabled = False
 End If
  
  
 Label28.Visible = True
 Label29.Visible = True
 Label29.Caption = "Nemáte autorizaèní práva"
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



