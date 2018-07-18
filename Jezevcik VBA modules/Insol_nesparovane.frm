VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Insol_nesparovane 
   Caption         =   "Jezevèík - Insolvenèní øízení"
   ClientHeight    =   11415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19785
   OleObjectBlob   =   "Insol_nesparovane.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Insol_nesparovane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit


Private Sub CheckBox3_Click()
Call stav_rizeni
End Sub

Private Sub CheckBox4_Click()
Call stav_rizeni
End Sub

Private Sub CheckBox5_Click()
Call stav_rizeni
End Sub

Private Sub CheckBox6_Click()
Call stav_rizeni
End Sub

Private Sub CheckBox7_Click()
Call stav_rizeni
End Sub

Sub stav_rizeni()

If cs_HDWMN.State = 1 Then
prihlasen:
 Call text_information_hide(Me.name)
 
 Call UserForm_Activate
 
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

' ListView sorting
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView1, ColumnHeader)
End Sub



' Back button
Private Sub CommandButton16_Click()
Call text_information_hide(formular)
Me.Hide
Insol.Show
End Sub



Private Sub ListView1_DblClick()
Dim sql As String
Dim ans As Integer

If cs_HDWMN.State = 1 Then
prihlasen:
 If ListView1.ListItems.count > 0 Then
  
  ans = MsgBox("Je klient uvedený v insolvenèním øízení klientem AirBank?", vbYesNoCancel)
   
    If ans = vbYes Then
        
        sql = "insert into " & schema & ".ins_client_list values (" & ListView1.SelectedItem & ",'" & Main.username & "' , '" & ListView1.SelectedItem.SubItems(6) & "' , sysdate)"
        Call execute(sql)
        
    ElseIf ans = vbNo Then
        
        sql = "insert into " & schema & ".ins_unconfirmed values (" & ListView1.SelectedItem & ", '" & ListView1.SelectedItem.SubItems(6) & "','" & Main.username & "' , sysdate )"
        Call execute(sql)
    
        
    ElseIf ans = vbCancel Then
    
        ' ??
    
    End If
    

 End If
 
 Call UserForm_Activate

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



' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()
Dim sql As String


If cs_HDWMN.State = 1 Then
prihlasen:

Call text_information_hide(Me.name)
ListView1.Sorted = 0



 ListView1.ListItems.Clear
 
 ' Define of sql query
 sql = "select distinct * from " & schema & ".v_ins_unpaired where 1=1 "
  
 sql = sql & " and (1=0 "
   
   If Insol_nesparovane.CheckBox4.value Then sql = sql & " or ic_stav_rizeni = 'KONKURS' "
   If Insol_nesparovane.CheckBox5.value Then sql = sql & " or ic_stav_rizeni = 'NEVYRIZENA' "
   If Insol_nesparovane.CheckBox6.value Then sql = sql & " or ic_stav_rizeni = 'ODDLUŽENÍ' "
   If Insol_nesparovane.CheckBox7.value Then sql = sql & " or ic_stav_rizeni = 'ÚPADEK' "
   If Insol_nesparovane.CheckBox3.value Then sql = sql & " or ic_stav_rizeni not in ('KONKURS','NEVYRIZENA','ODDLUŽENÍ','ÚPADEK') or ic_stav_rizeni is null "
   
 sql = sql & ") order by deadline desc nulls first "
 
  
Call query_insolvence_nesparovane(sql)

' If ToggleButton1.Value Then
'  Call ToggleButton1_Click
' End If
'
 
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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

Cancel = True
Me.Hide

End Sub


Private Sub UserForm_Initialize()
 Call table_prepare(Me.name)

End Sub



Private Sub query_insolvence_nesparovane(sql As String)
Dim itmX As ListItem, lvSI As ListSubItem

   Insol_nesparovane.ListView1.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Insol_nesparovane.ListView1.ListItems.Add()
      itmX.text = "" & rst1(0) & ""
      itmX.SubItems(1) = "" & rst1(1) & ""
      itmX.SubItems(2) = "" & rst1(2) & ""
      itmX.SubItems(3) = "" & rst1(3) & ""
      itmX.SubItems(4) = "" & rst1(4) & ""
      itmX.SubItems(5) = "" & rst1(5) & ""
      itmX.SubItems(6) = "" & rst1(6) & ""
      itmX.SubItems(7) = "" & rst1(7) & ""
      itmX.SubItems(8) = "" & rst1(8) & ""
      itmX.SubItems(9) = "" & rst1(9) & ""
      itmX.SubItems(10) = "" & rst1(10) & ""
    rst1.MoveNext
    Next i
   End If
   
   rst1.Close
   
   

    Call ZabarviRadkyTextove(Insol_nesparovane.ListView1, 3, "ÚPADEK", barvaFialova)
    Call ZabarviRadkyTextove(Insol_nesparovane.ListView1, 3, "KONKURS", barvaSeda)
    Call ZabarviRadkyTextove(Insol_nesparovane.ListView1, 3, "NEVYRIZENA", barvaZelenaSvetla)
    Call ZabarviRadkyTextove(Insol_nesparovane.ListView1, 3, "ODDLUŽENÍ", barvaRuda)
    
    
   
End Sub






'
'
'' Zmìna barvy po najetí myši **************************************************************************************************************************
'Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'CommandButton18.ForeColor = barvaCerna
'End Sub
'
'Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'CommandButton11.ForeColor = barvaCerna
'CommandButton12.ForeColor = barvaCerna
'ToggleButton1.ForeColor = barvaCerna
'End Sub
'
Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton16.ForeColor = barvaCerna
End Sub
'
'Private Sub CommandButton11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'CommandButton11.ForeColor = barvaZelenaAirBank
'End Sub
'
'Private Sub CommandButton12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'CommandButton12.ForeColor = barvaZelenaAirBank
'End Sub
'
'Private Sub CommandButton18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'CommandButton18.ForeColor = barvaZelenaAirBank
'End Sub
'
Private Sub CommandButton16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton16.ForeColor = barvaZelenaAirBank
End Sub
'
'Private Sub ToggleButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'ToggleButton1.ForeColor = barvaZelenaAirBank
'End Sub
'

