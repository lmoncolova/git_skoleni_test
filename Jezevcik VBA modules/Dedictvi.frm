VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dedictvi 
   Caption         =   "Jezevèík - Dìdické øízení"
   ClientHeight    =   11415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   22560
   OleObjectBlob   =   "Dedictvi.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dedictvi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit



Private Sub ComboBox2_Click()

Call query("", Me.name & "VhodnyCiselnik")

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

' Export to new Sheet
' FA zakomentovano 2.11.2017, CommandButton13 neexistuje
'''
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



Private Sub CheckBox1_Click()
Label10.Visible = CheckBox1.value
Label11.Visible = CheckBox1.value
DTPicker1.Visible = CheckBox1.value
DTPicker2.Visible = CheckBox1.value
Frame1.Repaint
End Sub


Private Sub ListView1_DblClick()
' v pripade ze je zaskrtnuty optionbutton pro DEDICE, tak se v pozadi udela to same,
' jako kdyby uzivatel otevrel nejdrive zemreleho a pak dblClick na listview s vybranym dedicem

Dim skp_heir As String

If cs_HDWMN.State = 1 Then
prihlasen:
If ListView1.ListItems.count > 0 Then
    
    If OptionButtonZemreli Then
        id_person = ListView1.SelectedItem.SubItems(1)
        Call set_filter(Me.name, 1)
        
        Me.Hide
        Klient_dedictvi.Show
        
    Else
        
        If 1 = 0 Then
            '''   nejdrive si musime ulozit skp_heir na konkretneho dedice
            'stary zpusob: ' skp_heir = NajdiSkpKDedicovi(ListView1.SelectedItem.SubItems(1), "skp_heir") 'stary zpusob nefunguje pokud ma jeden dedic vic zemrelych (vybere nahodne)
            skp_heir = ListvZiskejHodnotu(Dedictvi.ListView1, UCase("skp_heir"))
            
            '''   nastavime id_person na ZEMRELEHO (v listview je dedic!)
            'stary zpusob: 'id_person = NajdiSkpKDedicovi(ListView1.SelectedItem.SubItems(1), "id_person_dead") 'stary zpusob nefunguje pokud ma jeden dedic vic zemrelych (vybere nahodne)
            id_person = ListvZiskejHodnotu(Dedictvi.ListView1, UCase("id_person_dead"))
            
            Call set_filter(Me.name, 1)
            Me.Hide
            Klient_dedictvi.Show
            Klient_dedictvi.UserForm_Activate ' melo by se spustit samo s .Show, ale nespoustelo se, proto extra line
            
            'Tady musím nainicializovat dìdice, rovnako jako pri DblClick z karty zemreleho
            Set c = New Client
            Set c = dedici.item(CStr(skp_heir)).objDedic 'XXXX Tady to pøestalo fungovat, pøi prokliku ze seznamu dìdicù... GRRR
            Set d = dedici.item(CStr(skp_heir))
            Klient_heir.Show
        Else
            Call OtevriKartuDedice(Me.ListView1)
        End If
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





Private Sub OptionButtonDedicove_Click()
' pro dedice neni naimplementovano vyhledavani podle stavu
    
    ComboBox1.Enabled = False
    CheckBox1.Enabled = False
    Call CommandButton11_Click

End Sub

Private Sub OptionButtonZemreli_Click()
' pro dedice neni naimplementovano vyhledavani podle stavu
    
    ComboBox1.Enabled = True
    CheckBox1.Enabled = True
    Call CommandButton11_Click

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
'Private Sub ToggleButton1_Click()
'Dim sql As String
'
'If cs_HDWMN.State = 1 Then
'prihlasen:
' If ToggleButton1.Value = True Then
'  Frame3.Visible = True
'   sql = "select trunc(den, 'mm') mounth, sum(cnt_persons) cnt_person, sum (cnt_deals) cnt_deals  from " & schema & ".v_zaloby_souhrn where 1=1"
'
'
'   If ComboBox1.Text <> "Vše" Then sql = sql & " and name_epr_status  =  '" & ComboBox1.Text & "'"
'   'If ComboBox2.Text <> "Vše" Then sql = sql & " and name_process = '" & ComboBox2.Text & "'"
'   If CheckBox1.Value Then sql = sql & " and den between to_date('" & DTPicker1.Value & "','dd.mm.yyyy') and to_date('" & DTPicker2.Value & "','dd.mm.yyyy')"
'
'   sql = sql & " group by trunc(den, 'mm') order by trunc(den, 'mm')"
'   'ass = InputBox("", "", sql)
'   Call query(sql, Me.name & "Souhrn")
'
'   Set Graf = Worksheets(Me.name & "Souhrn").ChartObjects(1).Chart
'   Dateiname = "graf1.gif"
'   Graf.export Filename:=Dateiname, FilterName:="GIF"
'   Zaloby.Image2.Picture = LoadPicture(Dateiname)
'
'   Worksheets(Me.name & "Souhrn").Visible = False
'
'   Application.EnableEvents = True
'   Application.ScreenUpdating = True
'   Application.DisplayAlerts = True
' Else
'  Frame3.Visible = False
' End If
'
' If err_connection Then
'  err_connection = False
'  Call re_login(Main.username.Text, Main.password.Text)
'  If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
' End If
'
'Else
' Call re_login(Main.username.Text, Main.password.Text)
' If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
'End If
'End Sub

Private Sub Listview1_Napln(ByVal sql As String)
Dim clmx As ColumnHeaders
Dim keys As Dictionary
Dim itmX As ListItem

    '''   nejdrive se nastavi nazvy a velikosti sloupecku LISTVIEW
    
    Set clmx = Dedictvi.ListView1.ColumnHeaders
    If clmx.count < 3 Then
        clmx.Clear
        Call clmx.Add(, , "ID", 30)
        Call clmx.Add(, , "ID_PERSON", 65)
        Call clmx.Add(, , "CUID", 80)
        Call clmx.Add(, , "Klient", 120)
        Call clmx.Add(, , "SKP_HEIR", 1)
        Call clmx.Add(, , "ID_PERSON_DEAD", 1)
        Call clmx.Add(, , "Zemøelý", 90)
        Call clmx.Add(, , "Rodné èíslo", 80)
        Call clmx.Add(, , "Status", 160)
        Call clmx.Add(, , "Proces", 100)
        Call clmx.Add(, , "Spisová znaèka", 100)
        Call clmx.Add(, , "Platný od", 100)
        Call clmx.Add(, , "Platný do", 100)
        Call clmx.Add(, , "Nežádoucnost", 65)
        Call NastavKliceListview(Dedictvi.ListView1)
    End If
    
    '''   nalijou se data
    
    Dedictvi.ListView1.ListItems.Clear
    Set keys = KliceKListview(Dedictvi.ListView1)
    
    Set rst1 = New ADODB.Recordset
    'asss = InputBox("", "", sql)
    rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
    
    If Not rst1.EOF Then
        rst1.Sort = "ID_PERSON, DTIME_VALID_FROM DESC"
        rst1.MoveFirst
        For i = 1 To rst1.RecordCount
            Set itmX = Dedictvi.ListView1.ListItems.Add()
            itmX.text = "" & rst1![id] & ""
            itmX.SubItems(keys("ID_PERSON")) = "" & rst1![id_person] & ""
            itmX.SubItems(keys("CUID")) = "" & rst1![CUID] & ""
            itmX.SubItems(keys("Klient")) = "" & rst1![name] & ""
            itmX.SubItems(keys("SKP_HEIR")) = "" & rst1![skp_heir] & ""
            itmX.SubItems(keys("ID_PERSON_DEAD")) = "" & rst1![id_person_dead] & ""
            itmX.SubItems(keys("Zemøelý")) = "" & rst1![name_person_dead] & ""
            itmX.SubItems(keys("Rodné èíslo")) = "" & rst1![TEXT_BIRTH_NUMBER] & ""
            itmX.SubItems(keys("Status")) = "" & rst1![name_status] & ""
            itmX.SubItems(keys("Proces")) = "" & rst1![name_process] & ""
            itmX.SubItems(keys("Spisová znaèka")) = "" & rst1![code] & ""
            itmX.SubItems(keys("Platný od")) = "" & rst1![DTIME_VALID_FROM] & ""
            itmX.SubItems(keys("Platný do")) = "" & rst1![DTIME_VALID_TO] & ""
            itmX.SubItems(keys("Nežádoucnost")) = "" & rst1![flag_undesirability] & ""
            
            ''' PRED POUZITIM JE POTREBA AKTUALIZOVAT SLOUPECKY
'            If blnPrank Then
'                If i < 50 And Rnd() < 0.01 Then
'                    Set itmX = Dedictvi.ListView1.ListItems.Add()
'                    itmX.text = 1052
'                    itmX.SubItems(1) = 2063639
'                    itmX.SubItems(2) = 411265
'                    itmX.SubItems(3) = "Brožina Vít"
'                    itmX.SubItems(4) = "8509146969"
'                    itmX.SubItems(5) = "ukradl pejska z poboèky, dluží 9400" '"Vyjádøení vìøitele"
'                    itmX.SubItems(6) = "Dìdické zøízení"
'                    itmX.SubItems(7) = "10 D 89/2017"
'                    itmX.SubItems(8) = "9.6.2017 13:36:19"
'                    itmX.SubItems(9) = "1.1.3000"
'                    itmX.SubItems(10) = "N"
'                End If
'            End If
            
            rst1.MoveNext
        Next i
    End If
    
    rst1.Close

    
    Call ZabarviRadkyTextove(Dedictvi.ListView1, 0, "") ' jez.v.5.3

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

If ded_filtr.set = True Then
    TextBox1.text = ded_filtr.CUID
    TextBox2.text = ded_filtr.jmeno
    TextBox3.text = ded_filtr.rc
    TextBox4.text = ded_filtr.spis
    CheckBox1.value = True
    DTPicker1 = ded_filtr.od
    DTPicker2 = ded_filtr.do
    CheckBox1.value = ded_filtr.date
    ComboBox1.ListIndex = 0
     'ComboBox2.ListIndex = 0
     
    '  For i = 0 To ComboBox2.ListCount - 1
    '  If ComboBox2.List(i) = zal_filtr.proces Then
    '   ComboBox2.ListIndex = i
    '   Exit For
    '  End If
    ' Next i
  
    For i = 0 To ComboBox1.ListCount - 1
        If ComboBox1.List(i) = ded_filtr.stav Then
            ComboBox1.ListIndex = i
            Exit For
        End If
    Next i
    
Else
    ComboBox1.text = "Vše"
    'ComboBox2.Text = "Vše"
End If

ListView1.ListItems.Clear
 

 
'''''    sql query zavisi od toho, jestli chceme zemrelych nebo dedice

' vybereme view ze ktereho budeme cerpat
If OptionButtonZemreli Then
    sql = "select distinct d.id, d.id_person, d.cuid, d.name, d.skp_heir, d.id_person_dead, d.name_person_dead, d.text_birth_number, d.name_ded_status AS name_status, d.name_process, d.code, " & _
            "d.dtime_valid_from, d.dtime_valid_to, d.flag_undesirability FROM " & schema & ".v_dedictvi d where 1=1 "
Else
    sql = "SELECT distinct h.ID, h.id_person, h.cuid, h.name, h.skp_heir, h.id_person_dead, h.name_person_dead, h.text_birth_number, h.name_heir_status AS name_status, h.name_process, h.code, " & _
            "h.DTIME_VALID_FROM, h.DTIME_VALID_TO, h.flag_undesirability FROM " & schema & ".V_HEIRS h where 1=1 "
End If

' TextBoxy 1-4 jsou spolecne pro vyhledavanich zemrelych i dedicu
If (Len(Trim(TextBox1.text)) > 0 And IsNumeric(TextBox1)) Then sql = sql & " and cuid = '" & Trim(TextBox1.text) & "'"
If (Len(Trim(TextBox1.text)) > 0 And Not IsNumeric(TextBox1)) Then sql = sql & " and UPPER(convert(to_char(name), 'US7ASCII')) like UPPER(CONVERT('%" & Trim(TextBox1.text) & "%', 'US7ASCII')) " 'FA vyhladavanie pri zadani CUID a mena do nespraveho pola, ignoruje diakritiku
If (Len(Trim(TextBox2.text)) > 0 And IsNumeric(TextBox2)) Then sql = sql & " and cuid = '" & Trim(TextBox2.text) & "'"
If (Len(Trim(TextBox2.text)) > 0 And Not IsNumeric(TextBox2)) Then sql = sql & " and UPPER(convert(to_char(name), 'US7ASCII')) like UPPER(CONVERT('%" & Trim(TextBox2.text) & "%', 'US7ASCII')) " 'FA vyhladavanie pri zadani CUID a mena do nespraveho pola, ignoruje diakritiku
If Len(Trim(TextBox3.text)) > 0 Then sql = sql & " and text_birth_number like '%" & Trim(TextBox3.text) & "%'"
If Len(Trim(TextBox4.text)) > 0 Then sql = sql & " and upper(code) like '%" & UCase(Trim(TextBox4.text)) & "%'"

' stavy a obdobi nejsou pro dedice jeste naimplementovany
If OptionButtonZemreli Then
    If ComboBox1.text <> "Vše" Then sql = sql & " and name_ded_status  =  '" & ComboBox1.text & "'"
    'If ComboBox2.Text <> "Vše" Then sql = sql & " and name_process  =  '" & ComboBox2.Text & "'"
    If CheckBox1.value Then sql = sql & " and trunc(dtime_valid_from) between to_date('" & DTPicker1.value & "','dd.mm.yyyy') and to_date('" & DTPicker2.value & "','dd.mm.yyyy')"
End If

If Not CheckBox2.value Then sql = sql & " and nvl(dtime_valid_to, to_date('1.1.3000')) > sysdate"
If Not CheckBoxVsechnyZaznamy Then sql = OmezPocetRadkuVQuery(sql)

 
'''   z MainModule sa moze vyhodit kod pre Dedictvi.Listview1  -> COlumnHeaders (Sub Table_prepare) aj plnenie(sub Query) - bolo to presunute do tohto UserFormu ' RISK-468
'Call query(sql, Me.name)
Call Listview1_Napln(sql)

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
    ' Call table_prepare(Me.name) ' RISK-468
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


