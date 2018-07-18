VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Zesplatneni 
   Caption         =   "Jezevèík - Zesplatnìné smlouvy"
   ClientHeight    =   11475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   22560
   OleObjectBlob   =   "Zesplatneni.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Zesplatneni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit




Private Sub CommandButton11_Click()
Call set_filter(Me.name, 1)
Call UserForm_Activate
End Sub

Private Sub CommandButton19_Click()

Pridanie_klienta.Show

End Sub

Private Sub CheckBox1_Click()
Label10.Visible = CheckBox1.value
Label11.Visible = CheckBox1.value
DTPicker1.Visible = CheckBox1.value
DTPicker2.Visible = CheckBox1.value
Frame1.Repaint
End Sub

Private Sub CheckBox2_Click()
Call set_filter(Me.name, 1)
Call UserForm_Activate
'zes_filtr.set = False
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

' Search - Cancel button (All set to default settings)
Private Sub CommandButton12_Click()
Call set_filter(Me.name, 0)
Call UserForm_Activate
End Sub

' Export to new Sheet
'''' zakomentovano 8.11.2017, commandbutton13 neexistuje
''''Private Sub CommandButton13_Click()
''''
''''If cs_HDWMN.State = 1 Then
''''prihlasen:
''''Call text_information_hide(formular)
''''If form.ListView1.ListItems.count > 0 Then
''''
''''Application.EnableEvents = False
''''Application.ScreenUpdating = False
''''Application.DisplayAlerts = False
''''
''''Call Export("Form")
''''
''''Application.EnableEvents = True
''''Application.ScreenUpdating = True
''''Application.DisplayAlerts = True
''''
''''
'''' Label27.Visible = True
''''  Label29.Caption = "Dokonèeno"
'''' Label29.Visible = True
''''Else
'''' Label29.Caption = "Není co exportovat"
'''' Label29.Visible = True
'''' Label28.Visible = True
''''
''''
''''End If
''''
'''' If err_connection Then
''''  err_connection = False
''''  Call re_login(Main.username.text, Main.password.text)
''''  If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
'''' End If
''''
''''Else
'''' Call re_login(Main.username.text, Main.password.text)
'''' If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
''''End If
''''
''''End Sub

' Back button
Private Sub CommandButton16_Click()
Call text_information_hide(formular)
Call UserForm_QueryClose(1, 0)
End Sub





Private Sub ListView1_DblClick()
Dim sql As String

If cs_HDWMN.State = 1 Then
prihlasen:
 If ListView1.ListItems.count > 0 Then
  id_person = ListView1.SelectedItem
  Call set_filter(Me.name, 1)
  
  Me.Hide
  Pridanie_klienta.Hide
  Klient.Show
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

Private Sub ToggleButton1_Click()
Dim sql As String
Dim Graf As Variant
Dim Dateiname As String

If cs_HDWMN.State = 1 Then
prihlasen:
 If ToggleButton1.value = True Then
  Frame3.Visible = True
   
  
   sql = "select trunc(den, 'mm') mounth, sum(cnt_persons) cnt_person, sum (cnt_deals) cnt_deals  from " & schema & ".v_zesplatneni_souhrn where 1=1"
 
   If CheckBox1.value Then sql = sql & " and den between to_date('" & DTPicker1.value & "','dd.mm.yyyy') and to_date('" & DTPicker2.value & "','dd.mm.yyyy')"
   
   If ComboBox1.text <> "Vše" Then sql = sql & " and payoff_reason  =  '" & ComboBox1.text & "'"
   If ComboBox2.text = "Žádný" Then
    sql = sql & " and code_insolvency_case_status is null "
   ElseIf ComboBox2.text <> "Vše" Then
    sql = sql & " and code_insolvency_case_status = '" & ComboBox2.text & "'"
   End If
   
   sql = sql & " group by trunc(den, 'mm') order by trunc(den, 'mm')"
   'ass = InputBox("", "", sql)
   Call query(sql, Me.name & "Souhrn")
   
   'If Worksheets(Me.name & "Souhrn").Range("A2") <> "" Then
   
   Set Graf = Worksheets(Me.name & "Souhrn").ChartObjects(1).Chart
   Dateiname = "graf1.gif"
   Graf.Export fileName:=Dateiname, FilterName:="GIF"
   Zesplatneni.Image2.Picture = LoadPicture(Dateiname)
   
'   Else
'    Call text_information_hide(Me.name)
'    Label29.Caption = "Nejsou hodnoty pro zobrazení"
'    Label28.visible = True
'    Label29.visible = True
'   End If
   
   
   Worksheets(Me.name & "Souhrn").Visible = False
      
   Application.EnableEvents = True
   Application.ScreenUpdating = True
   Application.DisplayAlerts = True

 
 
 Else
  Frame3.Visible = False
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

' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()
Dim sql As String


If cs_HDWMN.State = 1 Then
prihlasen:

Call text_information_hide(Me.name)
ListView1.Sorted = 0
userform_prv = Me.name

Call table_prepare(Me.name & "Combo")
TextBox1.SetFocus 'FA klik - aby clovek nemusel klikat pro zadani CUID

If zes_filtr.set = True Then
 TextBox1.text = zes_filtr.CUID
 TextBox2.text = zes_filtr.jmeno
 TextBox3.text = zes_filtr.rc
 TextBox4.text = zes_filtr.smlouva
 CheckBox1.value = True
 DTPicker1 = zes_filtr.od
 DTPicker2 = zes_filtr.do
 CheckBox1.value = zes_filtr.date
 ComboBox1.text = zes_filtr.duvod
 ComboBox2.text = zes_filtr.code
 
End If

 ListView1.ListItems.Clear
 
 ' Define of sql query
 sql = "select distinct * from " & schema & ".v_zesplatneni where 1=1"
 
 If (Len(Trim(TextBox1.text)) > 0 And IsNumeric(TextBox1)) Then sql = sql & " and cuid = '" & Trim(TextBox1.text) & "'"
 If (Len(Trim(TextBox1.text)) > 0 And Not IsNumeric(TextBox1)) Then sql = sql & " and UPPER(convert(to_char(name), 'US7ASCII')) like UPPER(CONVERT('%" & Trim(TextBox1.text) & "%', 'US7ASCII')) " 'FA vyhladavanie pri zadani CUID a mena do nespraveho pola, ignoruje diakritiku
 If (Len(Trim(TextBox2.text)) > 0 And IsNumeric(TextBox2)) Then sql = sql & " and cuid = '" & Trim(TextBox2.text) & "'"
 If (Len(Trim(TextBox2.text)) > 0 And Not IsNumeric(TextBox2)) Then sql = sql & " and UPPER(convert(to_char(name), 'US7ASCII')) like UPPER(CONVERT('%" & Trim(TextBox2.text) & "%', 'US7ASCII')) " 'FA vyhladavanie pri zadani CUID a mena do nespraveho pola, ignoruje diakritiku
 If Len(Trim(TextBox3.text)) > 0 Then sql = sql & " and text_birth_number like '%" & Trim(TextBox3.text) & "%'"
 If Len(Trim(TextBox4.text)) > 0 Then sql = sql & " and upper(client_contracts) like '%" & UCase(Trim(TextBox4.text)) & "%'"
 If CheckBox1.value Then sql = sql & " and trunc(date_payoff) between to_date('" & DTPicker1.value & "','dd.mm.yyyy') and to_date('" & DTPicker2.value & "','dd.mm.yyyy')"
 If CheckBox2.value Then sql = sql & " and evidovano = 'Ne'"
 If ComboBox1.text <> "Vše" Then sql = sql & " and payoff_reason  =  '" & ComboBox1.text & "'"
 If ComboBox2.text = "Žádný" Then
  sql = sql & " and code_insolvency_case_status is null "
 ElseIf ComboBox2.text <> "Vše" Then
  sql = sql & " and code_insolvency_case_status = '" & ComboBox2.text & "'"
 End If
 
 If Not CheckBoxVsechnyZaznamy Then
    sql = OmezPocetRadkuVQuery(sql)
 End If
 
 Application.ScreenUpdating = False ' melo by to trochu zrychlit kod
 Call query_zesplatneni(sql)
 Application.ScreenUpdating = True
 
 If ToggleButton1.value Then
  Call ToggleButton1_Click
 End If
 
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
' First opening of UserForm
Call table_prepare(Me.name)

If userform_prv <> Me.name Then
 Call set_filter(Me.name, 0)
End If

End Sub


Private Sub query_zesplatneni(sql As String)
Dim clmx As ColumnHeader
Dim itmX As ListItem
Dim lvSI As ListSubItem

   Zesplatneni.ListView1.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    rst1.Sort = "DATE_PAYOFF"
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Zesplatneni.ListView1.ListItems.Add()
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
    rst1.MoveNext
    
    If blnPrank Then
        If i < 50 And i > 15 And Rnd() < 0.015 Then
            Set itmX = Zesplatneni.ListView1.ListItems.Add()
            itmX.text = 2063639
            itmX.SubItems(1) = 411265
            itmX.SubItems(2) = "Brožina Vít"
            itmX.SubItems(3) = "8509146969"
            itmX.SubItems(4) = "HU00009099"
            itmX.SubItems(5) = "19.8.2013"
            itmX.SubItems(6) = "Demencee"
            itmX.SubItems(7) = ""
            itmX.SubItems(8) = "N"
            itmX.SubItems(9) = "zásoba"
        End If
     End If
    
    Next i
   End If
   
   rst1.Close
   
   
    Call ZabarviRadkyTextove(Zesplatneni.ListView1, 9, "Ne") ' jez.v.5.3
    Call ZabarviRadkyTextove(Zesplatneni.ListView1, 9, "Ne (pod 1500)") ' jez.v.5.3

End Sub




Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Call set_filter(Me.name, 1)
Me.Hide
Pridanie_klienta.Hide
Main.Show
End Sub

' Zmìna barvy po najetí myši **************************************************************************************************************************
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton18.ForeColor = barvaCerna
CommandButton19.ForeColor = barvaCerna
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

Private Sub CommandButton19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton19.ForeColor = barvaZelenaAirBank
End Sub
