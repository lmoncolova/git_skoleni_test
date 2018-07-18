VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Insol 
   Caption         =   "Jezevèík - Insolvenèní øízení"
   ClientHeight    =   11415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   22560
   OleObjectBlob   =   "Insol.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Insol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit



Private Sub ComboBox2_Click()

Call query("", Me.name & "VhodnyCiselnik")

End Sub




Private Sub btn_generuj_ins_slozky_Click()
    Call Generate_folders_for_listed_clients(typRizeni_INS)
End Sub

Private Sub CheckBox2_Click()

 Call UserForm_Activate

End Sub

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


Private Sub CheckBox8_Click()
Call stav_rizeni
End Sub


Private Sub CheckBox9_Click()
Call stav_rizeni
End Sub


Private Sub CheckBox10_Click()
Call stav_rizeni
End Sub


Private Sub CheckBox11_Click()
Call stav_rizeni
End Sub


Private Sub CheckBox12_Click()
Call stav_rizeni
End Sub

' ListView sorting
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView1, ColumnHeader)
End Sub

' Export to new Workbook
Private Sub CommandButton18_Click()
Call text_information_hide(Me.name)
If cs_HDWMN.State = 1 Then
prihlasen:
If Me.ListView1.ListItems.count > 0 Then

Application.EnableEvents = False
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Call Export(Me.name)
    
Call export_novy(Me.name)

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True


 Label27.Visible = True
  Label29.Caption = "Dokonèeno"
 Label29.Visible = True
Else
 Label29.Caption = "Není co exportovat"
 Label29.Visible = True
 Label28.Visible = True


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




' Seach button (Adding condition to sql query)
Private Sub CommandButton11_Click()
Call set_filter(Me.name, 1)
Call UserForm_Activate
End Sub

' Search - Cancel button (All set to default settings)
Private Sub CommandButton12_Click()
Call set_filter(Me.name, 0)
Call UserForm_Activate
End Sub

' FA zakomentovano 2.11.2017
' CommandButton13 neexistuje a tahle SUB se nikde nevola

'''' Export to new Sheet
'''Private Sub CommandButton13_Click()
'''Call text_information_hide(formular)
'''If cs_HDWMN.State = 1 Then
'''prihlasen:
'''
'''If form.ListView1.ListItems.count > 0 Then
'''
'''Application.EnableEvents = False
'''Application.ScreenUpdating = False
'''Application.DisplayAlerts = False
'''
'''Call Export("Form")
'''
'''Application.EnableEvents = True
'''Application.ScreenUpdating = True
'''Application.DisplayAlerts = True
'''
'''
''' Label27.Visible = True
'''  Label29.Caption = "Dokonèeno"
''' Label29.Visible = True
'''Else
''' Label29.Caption = "Není co exportovat"
''' Label29.Visible = True
''' Label28.Visible = True
'''
'''
'''End If
'''
''' If err_connection Then
'''  err_connection = False
'''  Call re_login(Main.username.text, Main.password.text)
'''  If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
''' End If
'''
'''Else
''' Call re_login(Main.username.text, Main.password.text)
''' If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
'''End If
'''
'''End Sub



' Back button
Private Sub CommandButton16_Click()
Call text_information_hide(formular)
Call set_filter(Me.name, 1)
Me.Hide
Main.Show
End Sub

Private Sub CommandButton19_Click()

Insol_nesparovane.Show

End Sub


Private Sub CheckBox1_Click()
Label10.Visible = CheckBox1.value
Label11.Visible = CheckBox1.value
DTPicker1.Visible = CheckBox1.value
DTPicker2.Visible = CheckBox1.value
Frame1.Repaint
End Sub

Private Sub CheckBoxVsechnyZaznamy_Change()
 Call set_filter(Me.name, 1)
 Call UserForm_Activate
End Sub



Private Sub ListView1_DblClick()
Dim sql As String

If cs_HDWMN.State = 1 Then
prihlasen:
 If ListView1.ListItems.count > 0 Then
  id_person = ListView1.SelectedItem.SubItems(1)
  Call set_filter(Me.name, 1)
  
  Me.Hide
  Klient_insol.Show
  odkaz_Klient = False
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





Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then
 Call CommandButton11_Click
 TextBox1.SetFocus
End If
End Sub



Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then
 Call CommandButton11_Click
 TextBox2.SetFocus
End If
End Sub

Private Sub TextBox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then
 Call CommandButton11_Click
 TextBox3.SetFocus
End If
End Sub

Private Sub TextBox4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then
 Call CommandButton11_Click
 TextBox4.SetFocus
End If
End Sub
'

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



' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()
Dim sql As String


If cs_HDWMN.State = 1 Then
prihlasen:

Call text_information_hide(Me.name)
ListView1.Sorted = 0
userform_prv = Me.name

Call query("", Me.name & "Enumeration")
TextBox1.SetFocus 'FA klik - aby clovek nemusel klikat pro zadani CUID

If ins_filtr.set = True Then
 TextBox1.text = ins_filtr.CUID
 TextBox2.text = ins_filtr.jmeno
 TextBox3.text = ins_filtr.rc
 TextBox4.text = ins_filtr.spis
 CheckBox1.value = True
 DTPicker1 = ins_filtr.od
 DTPicker2 = ins_filtr.do
 CheckBox1.value = ins_filtr.date
 ComboBox1.ListIndex = 0
 'ComboBox2.ListIndex = 0
 
'  For i = 0 To ComboBox2.ListCount - 1
'  If ComboBox2.List(i) = zal_filtr.proces Then
'   ComboBox2.ListIndex = i
'   Exit For
'  End If
' Next i
  
 For i = 0 To ComboBox1.ListCount - 1
  If ComboBox1.List(i) = ins_filtr.stav Then
   ComboBox1.ListIndex = i
   Exit For
  End If
 Next i
 

 
Else
 ComboBox1.text = "Vše"
 'ComboBox2.Text = "Vše"
End If

 ListView1.ListItems.Clear
 
 ' Define of sql query
 sql = "select distinct * from " & schema & ".v_ins where 1=1 "
 
 If (Len(Trim(TextBox1.text)) > 0 And IsNumeric(TextBox1)) Then sql = sql & " and cuid = '" & Trim(TextBox1.text) & "'"
 If (Len(Trim(TextBox1.text)) > 0 And Not IsNumeric(TextBox1)) Then sql = sql & " and UPPER(convert(to_char(name), 'US7ASCII')) like UPPER(CONVERT('%" & Trim(TextBox1.text) & "%', 'US7ASCII')) " 'FA vyhladavanie pri zadani CUID a mena do nespraveho pola, ignoruje diakritiku
 If (Len(Trim(TextBox2.text)) > 0 And IsNumeric(TextBox2)) Then sql = sql & " and cuid = '" & Trim(TextBox2.text) & "'"
 If (Len(Trim(TextBox2.text)) > 0 And Not IsNumeric(TextBox2)) Then sql = sql & " and UPPER(convert(to_char(name), 'US7ASCII')) like UPPER(CONVERT('%" & Trim(TextBox2.text) & "%', 'US7ASCII')) " 'FA vyhladavanie pri zadani CUID a mena do nespraveho pola, ignoruje diakritiku
 If Len(Trim(TextBox3.text)) > 0 Then sql = sql & " and text_birth_number like '%" & Trim(TextBox3.text) & "%'"
 If Len(Trim(TextBox4.text)) > 0 Then sql = sql & " and upper(code) like '%" & UCase(Trim(TextBox4.text)) & "%'"
 If ComboBox1.text <> "Vše" Then sql = sql & " and name_ins_status  =  '" & ComboBox1.text & "'"
 'If ComboBox2.Text <> "Vše" Then sql = sql & " and name_process  =  '" & ComboBox2.Text & "'"
 If CheckBox1.value Then sql = sql & " and trunc(datum_zmeny_stavu) between to_date('" & DTPicker1.value & "','dd.mm.yyyy') and to_date('" & DTPicker2.value & "','dd.mm.yyyy')"
 If Not CheckBox2.value Then sql = sql & " and (dtime_valid_to > sysdate or dtime_valid_to is null) "

 sql = sql & " and (1=0 "
   
   If Insol.CheckBox4.value Then sql = sql & " or ic_stav_rizeni = 'KONKURS' "
   If Insol.CheckBox5.value Then sql = sql & " or ic_stav_rizeni = 'NEVYRIZENA' "
   If Insol.CheckBox6.value Then sql = sql & " or ic_stav_rizeni = 'ODDLUŽENÍ' "
   If Insol.CheckBox7.value Then sql = sql & " or ic_stav_rizeni = 'ÚPADEK' "
   If Insol.CheckBox3.value Then sql = sql & " or ic_stav_rizeni not in ('KONKURS','NEVYRIZENA','ODDLUŽENÍ','ÚPADEK') or ic_stav_rizeni is null "
   
 sql = sql & ")   and (1=0 OR vztah is null "
   
   If Insol.CheckBox8.value Then sql = sql & " or vztah = 'žadatel o úvìr    ' "
   If Insol.CheckBox9.value Then sql = sql & " or vztah = 'žadatel o hypotéku' "
   If Insol.CheckBox10.value Then sql = sql & " or vztah = 'debetáø  ' "
   If Insol.CheckBox11.value Then sql = sql & " or vztah = 'zástavce ' "
   If Insol.CheckBox12.value Then sql = sql & " or vztah = 'spoludlužník' "

   
   
 sql = sql & ") order by deadline desc "
 
 If Not CheckBoxVsechnyZaznamy Then
    sql = OmezPocetRadkuVQuery(sql)
 End If
 

Call query_insolvence(sql)

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



Private Sub UserForm_Initialize()
 Call table_prepare(Me.name)
If userform_prv <> Me.name Then
 Call set_filter(Me.name, 0)
End If
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Call set_filter(Me.name, 1)
Me.Hide
Main.Show
End Sub


Private Sub query_insolvence(sql As String)
Dim itmX As ListItem, lvSI As ListSubItem

   Insol.ListView1.ListItems.Clear

   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Insol.ListView1.ListItems.Add()
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
      itmX.SubItems(11) = "" & rst1(11) & ""
    rst1.MoveNext
    
    If blnPrank Then
        If i < 50 And i > 15 And Rnd() < 0.015 Then
            Set itmX = Insol.ListView1.ListItems.Add()
            itmX.text = 1052
            itmX.SubItems(1) = 2063639
            itmX.SubItems(2) = 411265
            itmX.SubItems(3) = "Macíèek Martin"
            itmX.SubItems(4) = "8509146969"
            itmX.SubItems(5) = "diletant"
            itmX.SubItems(6) = "Èeká na vyhlášení"
            itmX.SubItems(7) = "INS 16987/2017"
            itmX.SubItems(8) = "ODPRUŽENÍ"
            itmX.SubItems(9) = "9.6.2017 13:36:19"
            itmX.SubItems(10) = "Ano"
            itmX.SubItems(11) = "16"
        End If
    End If
    
    
    Next i
   End If

   rst1.Close


    Call ZabarviRadkyTextove(Insol.ListView1, 8, "ÚPADEK", barvaFialova)
    Call ZabarviRadkyTextove(Insol.ListView1, 8, "KONKURS", barvaSeda)
    Call ZabarviRadkyTextove(Insol.ListView1, 8, "NEVYRIZENA", barvaZelenaSvetla)
    Call ZabarviRadkyTextove(Insol.ListView1, 8, "ODDLUŽENÍ", barvaRuda)
    

End Sub



' Zmìna barvy po najetí myši **************************************************************************************************************************
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton18.ForeColor = barvaCerna
CommandButton19.ForeColor = barvaCerna
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton11.ForeColor = barvaCerna
CommandButton12.ForeColor = barvaCerna
End Sub

Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton16.ForeColor = barvaCerna
End Sub

Private Sub CommandButton11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton11.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton12.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton18.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton16.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton19.ForeColor = barvaZelenaAirBank
End Sub


