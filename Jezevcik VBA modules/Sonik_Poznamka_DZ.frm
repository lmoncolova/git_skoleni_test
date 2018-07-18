VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sonik_Poznamka_DZ 
   Caption         =   "Soník - Datová zpráva"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6465
   OleObjectBlob   =   "Sonik_Poznamka_DZ.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sonik_Poznamka_DZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit










Private Sub CommandButton_Zrus_Click()
    Call Me.Hide
End Sub

Private Sub CommandButton_Uloz_Click()
Dim sql_note As String
Dim id As Integer


If cs_HDWMN.State = 1 Then
prihlasen:
 If Len(Trim(TextBox_Poznamka_DZ.text)) > 0 Then
  If MsgBox("Opravdu chcete pøidat poznámku?", vbYesNo, "Pøidání poznámky") = vbYes Then
  
  
  
  sql_note = "insert into " & schema & ".sonik_data_message_note values  (" & schema & ".sonik_data_message_note_seq.nextval, " & SKP_datove_zpravy & ",null,'" & TextBox_Poznamka_DZ.text & "',sysdate,'" & UCase(Main.username) & "')"

  Call execute(sql_note)
  
  'Aktualizuju hlavni stranku DZ
  Call nacteni_dat_Sonik_Datova_Zprava
  
   Me.Hide
   
  End If
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

Private Sub TextBox_Poznamka_DZ_Change()
Label32.Caption = 500 - Len(TextBox_Poznamka_DZ.text)
End Sub

' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()
Dim sql As String


If cs_HDWMN.State = 1 Then
prihlasen:

TextBox_Poznamka_DZ.text = ""

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



