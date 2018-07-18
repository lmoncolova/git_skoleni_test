VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Exekuce 
   Caption         =   "Jezevèík - Exekuèní øízení"
   ClientHeight    =   11415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   22560
   OleObjectBlob   =   "Exekuce.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Exekuce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit
Private sqlExportRecordset As String


Private Sub ComboBox2_Click()

Call query1("", Me.name & "VhodnyCiselnik")

End Sub
    
Private Sub btn_generuj_exe_slozky_Click()
    Call Generate_folders_for_listed_clients(typRizeni_EXE)
End Sub

Private Sub CheckBox2_Click()
 Call set_filter(Me.name, 1)
 Call UserForm_Activate
End Sub

Private Sub CheckBoxVsechnyZaznamy_Change()
 Call set_filter(Me.name, 1)
 Call UserForm_Activate
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

Public Sub test_me(CUID As Long)
Dim i As Integer

On Error GoTo ERROR_HANDLER
    'Me.Show
    
    Me.TextBox1.text = CUID
    Call CommandButton11_Click 'hledani klienta
    For i = 1 To 2
        Call CheckBox1_Click
        Call CheckBox2_Click
        Call CheckBoxVsechnyZaznamy_Change
    Next i
    
    
    Me.ListView1_DblClick
    Exit Sub

ERROR_HANDLER:
    Debug.Assert 1 = 0
    Debug.Print "Spustil se error handler v objektu " & Me.name
    
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
'''
'''' Export to new Sheet
'''Private Sub CommandButton13_Click()
'''Call text_information_hide(formular)
'''If cs_HDWMN.State = 1 Then
'''prihlasen:
'''
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



Private Sub CheckBox1_Click()
Label10.Visible = CheckBox1.value
Label11.Visible = CheckBox1.value
DTPicker1.Visible = CheckBox1.value
DTPicker2.Visible = CheckBox1.value
Frame1.Repaint
End Sub



Public Sub ListView1_DblClick()
Dim sql As String

If cs_HDWMN.State = 1 Then
prihlasen:
 If ListView1.ListItems.count > 0 Then
  id_person = ListView1.SelectedItem.SubItems(1)
  Call set_filter(Me.name, 1)
  
  Me.Hide
  Klient_exe.Show
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

'Private Sub ToggleButton1_Click()
'Dim sql As String
'Dim Graf As Chart
'
'If cs_HDWMN.State = 1 Then
'prihlasen:
' If ToggleButton1.value = True Then
'  Frame3.Visible = True
'   sql = "select trunc(den, 'mm') mounth, sum(cnt_persons) cnt_person, sum (cnt_deals) cnt_deals  from " & schema & ".v_zaloby_souhrn where 1=1"
'
'
'    If ComboBox1.text <> "Vše" Then sql = sql & " and name_epr_status  =  '" & ComboBox1.text & "'"
'    'If ComboBox2.Text <> "Vše" Then sql = sql & " and name_process = '" & ComboBox2.Text & "'"
'    If CheckBox1.value Then sql = sql & " and den between to_date('" & DTPicker1.value & "','dd.mm.yyyy') and to_date('" & DTPicker2.value & "','dd.mm.yyyy')"
'
'    sql = sql & " group by trunc(den, 'mm') order by trunc(den, 'mm')"
'    'ass = InputBox("", "", sql)
'    'Call query(sql, Me.name & "Souhrn")
'
'    'Set Graf = Worksheets(Me.name & "Souhrn").ChartObjects(1).Chart
'    'Dateiname = "graf1.gif"
'    'Graf.Export fileName:=Dateiname, FilterName:="GIF"
'    'Zaloby.Image2.Picture = LoadPicture(Dateiname)
'    '
'    'Worksheets(Me.name & "Souhrn").Visible = False
'
'
' Else
'  Frame3.Visible = False
' End If
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
'End Sub

Private Sub ButtonExport_Click()
    
    Call Exportuj_Recordset_Do_Noveho_Workbooku(, sqlExportRecordset)
    
End Sub



' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()
Dim sql As String


If cs_HDWMN.State = 1 Then
prihlasen:

Call text_information_hide(Me.name)
ListView1.Sorted = 0
userform_prv = Me.name

Call query1("", Me.name & "Enumeration")
TextBox1.SetFocus 'FA klik - aby clovek nemusel klikat pro zadani CUID

If exe_filtr.set = True Then
 TextBox1.text = exe_filtr.CUID
 TextBox2.text = exe_filtr.jmeno
 TextBox3.text = exe_filtr.rc
 TextBox4.text = exe_filtr.spis
 CheckBox1.value = True
 DTPicker1 = exe_filtr.od
 DTPicker2 = exe_filtr.do
 CheckBox1.value = exe_filtr.date
 ComboBox1.ListIndex = 0
 'ComboBox2.ListIndex = 0
 
'  For i = 0 To ComboBox2.ListCount - 1
'  If ComboBox2.List(i) = zal_filtr.proces Then
'   ComboBox2.ListIndex = i
'   Exit For
'  End If
' Next i
  
 For i = 0 To ComboBox1.ListCount - 1
  If ComboBox1.List(i) = exe_filtr.stav Then
   ComboBox1.ListIndex = i
   Exit For
  End If
 Next i
 
 
''' button pouze pro vedouciho PV
If UCase(Main.username) = strVedouciPV Or schema = "APP_COLL_JEZEVCIK_TEST" Or InStr(strAdmins, UCase(Main.username)) > 0 Then
    Me.btn_generuj_exe_slozky.Visible = True
Else
    Me.btn_generuj_exe_slozky.Visible = False
End If

 
Else
 ComboBox1.text = "Vše"
 'ComboBox2.Text = "Vše"
End If

 ListView1.ListItems.Clear
 
 ' Definition of sql query
 sql = "select distinct * from " & schema & ".v_exe where 1=1 "
 

 If (Len(Trim(TextBox1.text)) > 0 And IsNumeric(TextBox1)) Then sql = sql & " and cuid = '" & Trim(TextBox1.text) & "'"
 If (Len(Trim(TextBox1.text)) > 0 And Not IsNumeric(TextBox1)) Then sql = sql & " and UPPER(convert(to_char(name), 'US7ASCII')) like UPPER(CONVERT('%" & Trim(TextBox1.text) & "%', 'US7ASCII')) " 'FA vyhladavanie pri zadani CUID a mena do nespraveho pola, ignoruje diakritiku
 If (Len(Trim(TextBox2.text)) > 0 And IsNumeric(TextBox2)) Then sql = sql & " and cuid = '" & Trim(TextBox2.text) & "'"
 If (Len(Trim(TextBox2.text)) > 0 And Not IsNumeric(TextBox2)) Then sql = sql & " and UPPER(convert(to_char(name), 'US7ASCII')) like UPPER(CONVERT('%" & Trim(TextBox2.text) & "%', 'US7ASCII')) " 'FA vyhladavanie pri zadani CUID a mena do nespraveho pola, ignoruje diakritiku
 If Len(Trim(TextBox3.text)) > 0 Then sql = sql & " and text_birth_number like '%" & Trim(TextBox3.text) & "%'"
 If Len(Trim(TextBox4.text)) > 0 Then sql = sql & " and upper(code) like '%" & UCase(Trim(TextBox4.text)) & "%'"
 If ComboBox1.text <> "Vše" Then sql = sql & " and name_exe_status  =  '" & ComboBox1.text & "'"
 'If ComboBox2.Text <> "Vše" Then sql = sql & " and name_process  =  '" & ComboBox2.Text & "'"
 If CheckBox1.value Then sql = sql & " and trunc(dtime_valid_from) between to_date('" & DTPicker1.value & "','dd.mm.yyyy') and to_date('" & DTPicker2.value & "','dd.mm.yyyy')"
 If Not CheckBox2.value Then sql = sql & " and dtime_valid_to > sysdate "
  
 If Not CheckBoxVsechnyZaznamy Then
    sql = OmezPocetRadkuVQuery(sql)
 End If
  
 Call query1(sql, Me.name)
 sqlExportRecordset = sql ' pro tlaèítko export
' If ToggleButton1.Value Then
'  Call ToggleButton1_Click
' End If
 
 
 
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

' Zmìna barvy po najetí myši **************************************************************************************************************************
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton18.ForeColor = barvaCerna
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton11.ForeColor = barvaCerna
CommandButton12.ForeColor = barvaCerna
ToggleButton1.ForeColor = barvaCerna
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

Private Sub ToggleButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
ToggleButton1.ForeColor = barvaZelenaAirBank
End Sub


