VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Souhrn 
   Caption         =   "Jezevèík v 3.0 - Souhrn"
   ClientHeight    =   10800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18600
   OleObjectBlob   =   "Souhrn.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Souhrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit




Private Sub ButtonExportUpo_Click_old()
' vytvori sa excel, vlozi sa donho kompletny vystup z view V_NOTICE_EXPORT a ulozi sa uzivatelovi na plochu
' Zobecnena funkcia bola vytvorena: Exportuj_Recordset_Do_Noveho_Workbooku()
Dim wbExportUpo As Workbook 'Workbook
Dim sql As String

If UCase(Main.username) = strVedouciPV Or schema = "APP_COLL_JEZEVCIK_TEST" Or InStr(strAdmins, UCase(Main.username)) > 0 Then
    ' Vytvorenie excelu
    Set wbExportUpo = Workbooks.Add
    With wbExportUpo
        .Title = "export upozornìní " & date
        .Subject = "export upozornìní " & date
        .SaveAs fileName:=CreateObject("WScript.Shell").specialfolders("Desktop") & "\export upozornìní " & date & ".xlsx"   'path to Desktop, regardless of user
    End With
    
    ' nacitanie hodnot
    sql = "SELECT * FROM " & schema & ".V_NOTICE_EXPORT"
    Set rst1 = New ADODB.Recordset
    rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        If Not rst1.EOF Then
            ' vlozenie hodnot do excelu
            Call VlozRecordsetDoWorkbooku(wbExportUpo, rst1)
        Else
            MsgBox "Neoèekávaná chyba, pošli printscreen analytikùm.", vbExclamation, "jez-error J62889"
        End If
    rst1.Close
    
    
Else
    MsgBox "Pokoušíš se dìlat vìci, na které nemáš práva. Ale Jezevèík je chytrý a odhalil tì, klidnì to zkoušej dál.", vbCritical, "Pokus o nepovolené vniknuntí"
End If

End Sub

Private Sub ButtonExportUpo_Click()
' vytvori sa excel, vlozi sa donho kompletny vystup z view V_NOTICE_EXPORT a ulozi sa uzivatelovi na plochu

'Dim wbExportUpo As Workbook 'Workbook
Dim sql As String

If UCase(Main.username) = strVedouciPV Or schema = "APP_COLL_JEZEVCIK_TEST" Or InStr(strAdmins, UCase(Main.username)) > 0 Then
    
    sql = "SELECT * FROM " & schema & ".V_NOTICE_EXPORT"
    Call Exportuj_Recordset_Do_Noveho_Workbooku(, sql, " upozornìní")
    
Else
    MsgBox "Pokoušíš se dìlat vìci, na které nemáš práva. Ale Jezevèík je chytrý a odhalil tì, klidnì to zkoušej dál.", vbCritical, "Pokus o nepovolené vniknuntí"
End If

End Sub

Private Sub cmdExportSonik_Click()

'vytvor novy workbook
'kopiruj hodnoty do noveho workbooku

Dim wb_export As Object, i As Integer, j As Integer, k As Integer

If LV_Sonik.ListItems.count > 0 Then

    Set wb_export = Workbooks.Add
    With wb_export
        .Title = "export ze Souhrnu Soníka " '& c.Name
        .Subject = "export ze Souhrnu Soníka " '& c.Name
        .SaveAs fileName:=CreateObject("WScript.Shell").specialfolders("Desktop") & "\export souhrn Soník " & format(Now, "yyyy/mm/dd hh/nn/ss") & ".xlsx" 'path to Desktop, regardless of user
    End With
    
    For i = 1 To (LV_Sonik.ListItems.count)
        
        If LV_Sonik.ListItems(i).Selected = True Then
            
            For j = 2 To 6

                Debug.Print LV_Sonik.ListItems(i).SubItems(j)
                'text = text & Chr(9) & LV_Sonik.SelectedItem.SubItems(j)
                wb_export.Sheets(1).Range("A2").Offset(k, j - 2) = LV_Sonik.ListItems(i).SubItems(j)
            Next j
            k = k + 1
            
        End If
        
    Next i
    
    'nazvy stlpcov do prveho riadku
    For j = 2 To 6
        wb_export.Sheets(1).Range("A1").Offset(0, j - 2) = LV_Sonik.ColumnHeaders(j)
    Next j
    
    wb_export.Sheets(1).Columns("A:I").AutoFit
    wb_export.Save
    
'    With New MSForms.DataObject
'        .SetText "http://zombo.com"
'        .PutInClipboard
'    End With
    
End If



End Sub

Private Sub ComboBoxAssignee_click()
Dim sql As String
Dim strUser As String

If Len(ComboBoxAssignee.value) > 0 Then
    strUser = ComboBoxAssignee.value
Else
    strUser = "nikoho"
End If


If prava("select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'NOTICE' and PRIVILEGE = 'INSERT'") Then
    
    If MsgBox("Opravdu chcete pøeøadit tohle upozornìní na " & strUser & "?", vbYesNo, "Úprava upozornìní") = vbYes Then
        sql = "update " & schema & ".notice set " & _
                "employee = '" & Main.username & "'" & _
                ", assignee = '" & ComboBoxAssignee.value & "'" & _
                " where skp_notice = " & Me.ListView3.SelectedItem.SubItems(8)
        Call execute(sql)
        sql = ""
    End If

Else
    MsgBox "Na zmìnu upozornìní nemáš práva.", vbExclamation, "To nepùjde"
End If

ComboBoxAssignee.value = ""
Call Listview3_Napln

End Sub

Private Sub ComboBoxFilterAssignee_Change()

Call Listview3_Napln

End Sub

Private Sub CommandButton16_Click()
Me.Hide
Main.Show
End Sub

' FA  - chcel som nastavit ze backspace ta vrati do menu hlavneho, ale nezistil som ako na to..
'Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'If KeyCode = 8 Then MsgBox ("YOLO") 'Unload Me 'backspace
''https://msdn.microsoft.com/en-us/library/aa243025(v=vs.60).aspx
'End Sub




'Private Sub CommandButton18_Click() ' RISK-599 button ide prec by Filip Antol 24.5.2017 v ramci zmeny vzhladu celej karty Souhrn
'
'
'
'If cs_HDWMN.State = 1 Then
'prihlasen:
'
''If ListView1.ListItems.count > 0 Then
''
''Application.EnableEvents = False
''Application.ScreenUpdating = False
''Application.DisplayAlerts = False
''
''Call export_novy(Me.name)
''
''Application.EnableEvents = True
''Application.ScreenUpdating = True
''Application.DisplayAlerts = True
''
''End If
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
'
'End Sub



Private Sub ListView4_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView4, ColumnHeader)
End Sub


Private Sub ListView4_DblClick()
Dim sql As String


If cs_HDWMN.State = 1 Then
prihlasen:

If ListView4.ListItems.count > 0 Then
 id_person = ListView4.SelectedItem
 userform_prv = Me.name & "Auth"
 Souhrn.Hide
 Klient.Show
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


Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView3, ColumnHeader)
End Sub


Private Sub ListView3_DblClick()
Dim sql As String


If cs_HDWMN.State = 1 Then
prihlasen:

    If ListView3.ListItems.count > 0 Then
        id_person = ListView3.SelectedItem
        userform_prv = Me.name & "Notice"
        Souhrn.Hide
        'RISK-130
        If ListView3.SelectedItem.SubItems(6) = vbNullString Then
            Klient.Show
        ElseIf ListView3.SelectedItem.SubItems(6) = 1 Then
            Klient_dedictvi.Show
        ElseIf ListView3.SelectedItem.SubItems(6) = 2 Then
            Klient_exe.Show
        ElseIf ListView3.SelectedItem.SubItems(6) = 3 Then
            Klient.Show
        ElseIf ListView3.SelectedItem.SubItems(6) = 4 Then
            Klient_insol.Show
        ElseIf ListView3.SelectedItem.SubItems(6) = 5 Then
            Klient_trest.Show
        ElseIf ListView3.SelectedItem.SubItems(6) = 6 Then 'RISK-468
            Call OtevriKartuDedice(Souhrn.ListView3)
            'MsgBox "Filip na prokliku do dìdice maká", vbInformation
            'Souhrn.Show
        Else: Klient.Show 'Pojistka kdyby nìco, ale nemìlo by se stát
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


Private Sub LV_KontrolaPM_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(LV_KontrolaPM, ColumnHeader)
End Sub


' FA pokus 30.5.2017
'Private Sub LV_KontrolaPM_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'If KeyCode = 13 Then
' Call LV_KontrolaPM_DblClick
' 'TextBox1.SetFocus
'End If
'End Sub

Private Sub LV_KontrolaPM_DblClick()
Dim sql As String


If cs_HDWMN.State = 1 Then
prihlasen:

If LV_KontrolaPM.ListItems.count > 0 Then
    id_person = LV_KontrolaPM.SelectedItem
    userform_prv = Me.name & "Notice"
    Souhrn.Hide
    Klient.Show
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

Private Sub LV_Sonik_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(LV_Sonik, ColumnHeader)
End Sub

Private Sub LV_Sonik_DblClick()
Dim sql As String
Dim rst_dm As Variant


If cs_HDWMN.State = 1 Then
prihlasen:

''' ' RISK-985 odkomentovat po hromadnom REPLACE
If LV_Sonik.ListItems.count > 0 Then
 SKP_datove_zpravy = LV_Sonik.SelectedItem.SubItems(1)

 sql = "select data_message_number, date_data_message_delivery, data_message_sender, text_file_number, text_reference_number FROM " & schema & ".sonik_data_message WHERE skp_data_message = " & SKP_datove_zpravy
 Set rst_dm = New ADODB.Recordset
 rst_dm.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

    If Not rst_dm.EOF Then

        rst_dm.MoveFirst

        Cislo_datove_zpravy = "" & rst_dm![data_message_number] & ""
        Datum_doruceni_datove_zpravy = "" & rst_dm![date_data_message_delivery] & ""
        Odesilatel_datove_zpravy = "" & rst_dm![data_message_sender] & ""
        Spisova_Znacka_datove_zpravy = "" & rst_dm![text_file_number] & ""
        Cislo_Jednaci_datove_zpravy = "" & rst_dm![text_reference_number] & ""

    End If
 rst_dm.Close

 If Sonik_User Then
    Call nacteni_dat_Sonik_Datova_Zprava
 Else
    MsgBox "Milý uživateli, chtìl ses dostat do Soníka oklikou, ale vìz, že se ti to nepodaøí.", vbCritical, "Pokus o nepovolené vniknutí!"
    Exit Sub
 End If

 userform_prv = Me.name & "Notice"
 Souhrn.Hide
 Sonik_Datova_Zprava.Show
End If
''' ' RISK-985 odkomentovat po hromadnom REPLACE
 
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



Private Sub Listview3_Napln()
Dim user As String
user = ComboBoxFilterAssignee.value

If Len(user) > 1 Then
    Call query("select n.id_person, n.cuid, n.name, n.text_notice, n.dtime_notice_deadline, n.days_left, n.skp_process, n.assignee, n.skp_notice from " & schema & ".v_notice n WHERE text_notice NOT LIKE '%Kontrola PM%' AND upper(n.assignee) IN ('" & UCase(user) & "', '" & UCase(Right(user, Len(user) - 4)) & "')", Me.name & "Upozorneni") '
Else
    Call query("select n.id_person, n.cuid, n.name, n.text_notice, n.dtime_notice_deadline, n.days_left, n.skp_process, n.assignee, n.skp_notice from " & schema & ".v_notice n WHERE text_notice NOT LIKE '%Kontrola PM%'", Me.name & "Upozorneni")
End If

End Sub

'Private Sub ShowChart() ' RISK-599 - ShowChart ide prec by Filip Antol 24.5.2017 v ramci zmeny vzhladu celej karty Souhrn
'On Error GoTo chyba
'    Application.EnableEvents = False
'    Application.ScreenUpdating = False
'    Application.DisplayAlerts = False
'
'   wa.Worksheets("SouhrnGraf").Visible = True
'   wa.Worksheets("SouhrnGraf").Activate
'
'   Range("TabulkaGraf[#All]").ClearContents
'
''   ActiveSheet.Range("A1") = Souhrn.ListView1.ColumnHeaders(2)
''   ActiveSheet.Range("B1") = Souhrn.ListView1.ColumnHeaders(3)
''
''    For i = 1 To Souhrn.ListView1.ListItems.count
''      Range("A1").Offset(i, 0) = Souhrn.ListView1.ListItems(i).SubItems(1)
''      Range("B1").Offset(i, 0) = Souhrn.ListView1.ListItems(i).SubItems(2)
''    Next i
'
'    Range("A1:B1").Select
'
'    Range(Selection, Selection.End(xlDown)).Select
'
'    Selection.WrapText = False
'
'
'    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).name = "TabulkaGraf"
'    Range("TabulkaGraf[#All]").Select
'    ActiveSheet.ListObjects("TabulkaGraf").TableStyle = "TableStyleMedium18"
'
'    Columns("A:B").EntireColumn.AutoFit
'
'
'    ActiveSheet.ChartObjects("Graf 3").Activate
'    ActiveChart.SetSourceData Source:=Range("TabulkaGraf[#All]")
'
'
'    Set Graf = Worksheets(Me.name & "Graf").ChartObjects(1).Chart
'    Dateiname = "graf1.jpg"
'    Graf.Export fileName:=Dateiname, FilterName:="JPG"
'    Souhrn.Image2.Picture = LoadPicture(Dateiname)
'
'    Worksheets(Me.name & "Graf").Visible = False
'
'    Application.EnableEvents = True
'    Application.ScreenUpdating = True
'    Application.DisplayAlerts = True
'
'Exit Sub
'chyba:
' CommandButton18.Visible = False
'End Sub


' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()
Dim sql As String


If cs_HDWMN.State = 1 Then
prihlasen:

Souhrn.Caption = "Jezevèík v." & version & " - Souhrn"

If strAdmins = "" Then
    Call NaplnAdminov
End If
If UCase(Main.username) = strVedouciPV Or schema = "APP_COLL_JEZEVCIK_TEST" Or InStr(strAdmins, UCase(Main.username)) > 0 Then
    ' miesto podmienky na schema by to chcelo podmienku na Adminov: UCase(Main.username) in (SELECT * FROM jez_users a WHERE a.user_role = 'ADMIN')
    ButtonExportUpo.Visible = True
Else
    ButtonExportUpo.Visible = False
End If

userform_prv = Me.name

' XX ListView1 XX Call query("select b.skp_epr_status, b.name_epr_status, count(distinct a.id_person) as pocet from " & schema & ".epr_status_data a join " & schema & ".epr_status b on b.skp_epr_status = a.skp_epr_status where a.dtime_valid_to > sysdate and a.flag_deleted = 'N' group by b.skp_epr_status, b.name_epr_status order by pocet desc", Me.name & "Pocet")
' Call query("select * from " & schema & ".v_epr_status_interval", Me.name & "Doba")
'Call query("select * from " & schema & ".v_notice", Me.name & "Upozorneni") 'Risk-599 by FA
'Call query("select * from " & schema & ".v_sop_auth", Me.name & "Autorizace")
'RISK-130
Listview3_Napln ' query pre upozorneni 'Call query("select n.id_person, n.cuid, n.name, n.text_notice, n.dtime_notice_deadline, n.days_left, n.skp_process from " & schema & ".v_notice n WHERE text_notice NOT LIKE '%Kontrola PM%'", Me.name & "Upozorneni")

''''''' RISK-1098 sql upraveno na nove VIEW
'''Call query("SELECT a.id_person, p.cuid, a.name, a.text, t.okresni_soud, " & _
'''           "COALESCE(s.AMT_RETURN_SOP2, s.AMT_ADDITIONAL_SOP2, s.AMT_SOP2, s.AMT_RETURN_SOP1, s.AMT_ADDITIONAL_SOP1, s.AMT_SOP1) AS Castka, " & _
'''           "COALESCE(S.DATE_INFO_RETURN_SOP2, S.DATE_ADDITIONAL_SOP2, S.DATE_SOP2, S.DATE_INFO_RETURN_SOP1, S.DATE_ADDITIONAL_SOP1, S.DATE_SOP1) AS DATUM_K_CASTCE " & _
'''           "FROM " & schema & ".v_sop_auth a JOIN " & schema & ".PERSON_SOUD t ON a.id_person = t.id_person JOIN " & schema & ".sop s ON a.id_person = s.ID_PERSON JOIN " & schema & ".person p ON p.id_person = a.id_person", Me.name & "Autorizace") 'FA RISK-599
Call query("SELECT a.id_person, a.cuid, a.name, a.text, a.okresni_soud, a.castka, a.datum_k_castce " & _
           "FROM " & schema & ".v_souhrn_autorizace a" _
           , Me.name & "Autorizace")

Call query("select sum(amt_sop_balance) as suma1, sum(priznan_sop) as suma2, sum(sop_paid) as suma3 from " & schema & ".v_sop_balance", Me.name & "SOP")
Call query1("select n.id_person, n.cuid, n.name, n.text_notice, n.dtime_notice_deadline, n.days_left from " & schema & ".v_notice n WHERE text_notice LIKE '%Kontrola PM%'", Me.name & "KontrolaPM")
Call query1("SELECT s.skp_data_message_notice, s.skp_data_message, nvl(s.cuids, '') cuids, nvl(s.jmena, '') jmena, s.text_notice, trunc(s.dtime_notice_deadline) dtime_notice_deadline, s.days_left, s.data_message_number  FROM " & schema & ".SONIK_V_NOTICE s", Me.name & "Upozorneni_Sonik")

 
 If Me.ComboBoxAssignee.ListCount < 1 Then
    Call NaplnAssigneeSeznam(Souhrn.ComboBoxAssignee)
 End If
 ComboBoxAssignee.value = ""
 If Me.ComboBoxFilterAssignee.ListCount < 1 Then
    Call NaplnAssigneeSeznam(Souhrn.ComboBoxFilterAssignee)
 End If
 'ComboBoxFilterAssignee.value = ""
 
 
 Me.Repaint

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


'Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'CommandButton18.ForeColor = barvaCerna
'End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Call set_filter(Me.name, 1)
Me.Hide
Main.Show
End Sub


Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton16.ForeColor = barvaCerna
End Sub


Private Sub CommandButton16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton16.ForeColor = barvaZelenaAirBank
End Sub

'Private Sub CommandButton18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'CommandButton18.ForeColor = barvaZelenaAirBank
'End Sub

