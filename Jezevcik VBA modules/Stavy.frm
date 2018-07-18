VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Stavy 
   Caption         =   "Jezevèík v 1.0 - Seznam stavù"
   ClientHeight    =   10770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13785
   OleObjectBlob   =   "Stavy.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Stavy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit
'RISK-468 RISK-949

Private Sub OptionButton1_Click()

Call UserForm_Activate

End Sub
Private Sub OptionButton2_Click()

Call UserForm_Activate

End Sub
Private Sub OptionButton3_Click()

Call UserForm_Activate

End Sub
Private Sub OptionButton4_Click()

Call UserForm_Activate

End Sub

Private Sub OptionButton5_Click()

Call UserForm_Activate

End Sub

Private Sub OptionButton6_Click()

Call UserForm_Activate

End Sub

Private Sub OptionButton7_Click()

Call UserForm_Activate

End Sub


Private Sub CommandButton11_Click()
If cs_HDWMN.State = 1 Then
prihlasen:

Stavy_update.ComboBox1.Locked = False
Stavy_update.ComboBox1.Enabled = True

 If prava("select * from USER_TAB_PRIVS where owner = '" & schema & "' and table_name = 'EPR_STATUS' and PRIVILEGE = 'INSERT'") Then
  Stavy_update.CommandButton20.Caption = "Uložit"
  Stavy_update.Caption = "Pøidání nového stavu"
  Stavy_update.Label26.Caption = "Pøidání nového stavu"

  Stavy_update.TextBox1.Enabled = True
  Stavy_update.TextBox1 = ""
  Stavy_update.CheckBox1.Visible = False
  Stavy_update.CheckBox1.value = False

  Stavy_update.Show
  
 ' Call UserForm_Activate
 Else
  Call text_information_hide(Me.name)
  
  Label29.Caption = "Nemáte práva pro pøidání nového stavu"
  Label28.Visible = True
  Label29.Visible = True
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

Private Sub CommandButton16_Click()
Me.Hide
Main.Show
End Sub


Private Sub CheckBox1_Click()

Call UserForm_Activate

End Sub

Private Sub ListView1_DblClick()
If cs_HDWMN.State = 1 Then
prihlasen:

Call text_information_hide(Me.name)

Stavy_update.ComboBox1.Locked = True
Stavy_update.ComboBox1.Enabled = False

If ListView1.ListItems.count > 0 Then
 If ListView1.SelectedItem.SubItems(4) = "1.1.3000" Then
   
 If prava("select * from USER_TAB_PRIVS where owner = '" & schema & "' and table_name = 'EPR_STATUS' and PRIVILEGE = 'INSERT'") Then
  Stavy_update.CommandButton20.Caption = "Upravit"
  Stavy_update.Caption = "Úprava stávajícího stavu"
  Stavy_update.Label26.Caption = "Úprava stávajícího stavu"
  Stavy_update.CheckBox1.Visible = True
  Stavy_update.CheckBox1.value = False
  Stavy_update.TextBox1.text = ListView1.SelectedItem.SubItems(1)
  Stavy_update.ComboBox1 = ListView1.SelectedItem.SubItems(2)
  Stavy_update.Show
 Else
  
  Label29.Caption = "Nemáte práva pro úpravu stavu"
  Label28.Visible = True
  Label29.Visible = True
 End If
 
 
 
 Else
  Label29.Caption = "Nelze upravovat døíve odstranìné stavy"
  Label28.Visible = True
  Label29.Visible = True
 End If
 
Else
  Label29.Caption = "Není co vybrat"
  Label28.Visible = True
  Label29.Visible = True
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

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView1, ColumnHeader)
End Sub


' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()
Dim sql As String
'RISK-468 RISK-949

If cs_HDWMN.State = 1 Then
prihlasen:

Stavy.Caption = "Jezevèík v." & version & " - Stavy"

Call text_information_hide(Me.name)
ListView1.Sorted = 0
userform_prv = Me.name

If OptionButton1.value = False And OptionButton2.value = False And OptionButton3.value = False And OptionButton4.value = False And OptionButton5.value = False And OptionButton6.value = False And OptionButton7.value = False Then
    OptionButton1.value = True
End If

'RISK-468 'RISK-949
sql = "select skp_process , name_process from " & schema & ".process" ' where skp_process in (1,2,3,5,4)"
Call query(sql, Me.name & "Proces")

sql = "select * from " & schema & ".v_status_epr b where 1 = 1"
Call query(sql, Me.name)
 
 Me.Repaint

If ListView1.ListItems.count > 0 Then
 Label29.Caption = "Nalezeno " & ListView1.ListItems.count & " záznamù"
 Label27.Visible = True
 Label29.Visible = True
Else
 Label29.Caption = "Nalezeno " & ListView1.ListItems.count & " záznamù"
 Label28.Visible = True
 Label29.Visible = True
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



Private Sub UserForm_Initialize()
    Call table_prepare(Me.name)
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Me.Hide
Main.Show
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton11.ForeColor = barvaCerna
End Sub

Private Sub CommandButton11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton11.ForeColor = barvaZelenaAirBank
End Sub


Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton16.ForeColor = barvaCerna
End Sub

Private Sub CommandButton16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton16.ForeColor = barvaZelenaAirBank
End Sub


