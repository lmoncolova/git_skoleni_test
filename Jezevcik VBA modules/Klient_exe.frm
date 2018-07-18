VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Klient_exe 
   Caption         =   "Jezevèík - Exekuèní øízení"
   ClientHeight    =   11775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   22815
   OleObjectBlob   =   "Klient_exe.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Klient_exe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit


Public Sub test_me(ByVal id_person_in As Long)

On Error GoTo ERROR_HANDLER

    id_person = id_person_in
    ' Call UserForm_Initialize ' ak error, search "xtstereehcvgb".
    Call UserForm_Activate
    Me.Show
    ' sablony
    Call CommandButton31_Click
    
    
    ''' prokliky
'    Call CommandButton27_Click
'    Call CommandButton28_Click
'    Call CommandButton29_Click
'    Call CommandButton30_Click
'    ''' Call ButtonZmenitStav_Click ' okno je MODALNI, zastavi exekuci
'    Call CommandButton19_Click
'    Call CommandButtonVlozUpo_Click
'    Call CommandButton25_Click
    
    
NORMAL_FLOW:
    Exit Sub

ERROR_HANDLER:
    Debug.Assert 1 = 0
    Debug.Print "Spustil se error handler v objektu " & Me.name
    
    
End Sub


Private Sub history_filter()
If cs_HDWMN.State = 1 Then
prihlasen:
 Call text_information_hide(Me.name)
 
 Call query1("", Me.name & "History")
 
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

Private Sub CommandButton_Sonik_Click()
'RISK-721
 Call Client_From_Jez_To_Sonik(Klient_exe.name, 3) ' RISK-985 odkomentovat po hromadnom REPLACE
End Sub

Private Sub CommandButtonVlozUpo_Click()
If cs_HDWMN.State = 1 Then
prihlasen:

If prava("select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'NOTICE' and PRIVILEGE = 'INSERT'") Then

    Upozorneni_all.VstupnyUserform = Klient_exe
    Upozorneni_all.VstupNoveUpo = True
    Upozorneni_all.Show
    
Else
    Call text_information_hide(Me.name)
    
    Label29.Caption = "Nemáte práva pro pøidávání upozornìní"
    Label28.Visible = True
    Label29.Visible = True
End If
 
If err_connection Then
    err_connection = False
    Call re_login(Main.username.text, Main.password.text)
    If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If

Else 'ne-prihlasen:
    Call re_login(Main.username.text, Main.password.text)
    If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If
End Sub


Private Sub ButtonZmenitStav_Click()
If cs_HDWMN.State = 1 Then
prihlasen:
 
If prava("select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'EPR_STATUS_DATA' and PRIVILEGE = 'INSERT'") Then
   Update_stav_all.VstupnyUserform = Klient_exe
   Update_stav_all.Show
Else
    Call text_information_hide(Me.name)
    
    Label29.Caption = "Nemáte práva pro zmìnu stavu"
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
'If odkaz_Klient = False Then

Select Case userform_prv
Case "Exekuce"
    Exekuce.Show
    odkaz_Klient = False
Case "Zaloby"
    Klient.Show
    odkaz_Klient = True
Case "Insol"
    Klient_insol.Show
    odkaz_Klient = True
Case "Dedictvi"
    Klient_dedictvi.Show
    odkaz_Klient = True
Case "Zesplatneni"
    Zesplatneni.Show
    odkaz_Klient = True
Case "Nezesplatneni"
    Nezesplatneni.Show
    odkaz_Klient = False
Case "SouhrnNotice", "SouhrnAuth", "Souhrn"
    Souhrn.Show
    odkaz_Klient = False
Case "Klient_heir"
    Klient_heir.Show
End Select


'End If

TextBox1.text = ""

End Sub

Private Sub CommandButton18_Click()
If cs_HDWMN.State = 1 Then
prihlasen:

Call text_information_hide(Me.name)

   Application.EnableEvents = False
   Application.ScreenUpdating = False
   Application.DisplayAlerts = False
   
   
   Worksheets(Me.name).Visible = True
   Worksheets(Me.name).Activate
   Application.GoTo Reference:="obsah"
   Selection.ClearContents
   Worksheets(Me.name).Range("D8:I9").ClearContents
   Worksheets(Me.name).Range("D10:I11").ClearContents
   Worksheets(Me.name).Range("D13:I14").ClearContents
   
   Range("D4") = Label2.Caption
   Range("D5") = Label3.Caption
   Range("D6") = Label4.Caption
   Range("D8") = Label5.Caption
   Range("D10") = Label6.Caption
   
   Range("D13") = ListView2.ListItems(1)
   For i = 2 To ListView2.ListItems.count
    Range("D13") = Range("D13") & ", " & ListView2.ListItems(i)
   Next i
   
   For i = 1 To ListView5.ListItems.count - 2
    Range("B23").Offset(i - 1, 0) = ListView5.ListItems(i)
    Range("D23").Offset(i - 1, 0) = ListView5.ListItems(i).SubItems(1)
   Next i
   
   For i = 1 To ListView4.ListItems.count - 2
    If i <= 9 Then
     Range("D39").Offset(i - 1, 0) = ListView4.ListItems(i).SubItems(1)
    Else
     Range("H39").Offset(i - 10, 0) = ListView4.ListItems(i).SubItems(1)
    End If
   Next i
   
   Range("H50") = Label77.Caption
   Range("H51") = Label78.Caption
   
   For i = 1 To ListView1.ListItems.count
     Range("J5").Offset(i - 1, 0) = ListView1.ListItems(i).SubItems(1)
     Range("K5").Offset(i - 1, 0) = ListView1.ListItems(i).SubItems(3)
     Range("L5").Offset(i - 1, 0) = ListView1.ListItems(i).SubItems(5)

   Next i
   
   
   
   
   Range("J5").Offset(i + 1, 0) = date & " " & Time
   Range("K5").Offset(i + 1, 0) = "Vyexportováno"
   Range("L5").Offset(i + 1, 0) = Main.username
   
   
   Call query("select b.name_epr_status, a.dtime_valid_from from " & schema & ".epr_status_data a join " & schema & ".epr_status b on a.skp_epr_status = b.skp_epr_status where a.flag_deleted = 'N' and a.dtime_valid_to > sysdate and id_person = " & id_person, Me.name & "Export")
   
   
   Call export_pdf(Me.name)
   
   Application.EnableEvents = True
   Application.ScreenUpdating = True
   Application.DisplayAlerts = True

   
 Label29.Caption = "Vyexportováno (uloženo do Dokumentù)"
 Label27.Visible = True
 Label29.Visible = True
   
 
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

Private Sub CommandButton19_Click()

If cs_HDWMN.State = 1 Then
prihlasen:

 If prava("select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'INFORMATION' and PRIVILEGE = 'INSERT'") Then
 
  Update_klient_exe.Label1.Caption = Label3.Caption
  Update_klient_exe.Label2.Caption = Label4.Caption

 
  For i = 1 To ListView5.ListItems.count
   Select Case ListView5.ListItems(i)
    Case "Datum žaloby"
     Update_klient_exe.Label6.Caption = ListView5.ListItems(i).SubItems(1)
    Case "Exekuèní znaèka"
     Update_klient_exe.Label5.Caption = ListView5.ListItems(i).SubItems(1)
    Case "Výše pohledávky"
     Update_klient_exe.Label7.Caption = ListView5.ListItems(i).SubItems(1)
'    Case "Datum vyèíslení pohledávky"
'     Update_klient.Label8.Caption = ListView5.ListItems(i).SubItems(1)
' Case "ID Datové schránky"
'   Update_klient.Label9.Caption = ListView5.ListItems(i).SubItems(1)
   End Select
  Next i
 
  Update_klient_exe.Show
 
 Else
  Call text_information_hide(Me.name)

  Label29.Caption = "Nemáte práva pro zmìnu informací o klientovi"
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

Private Sub ButtonVlozPoznamku_Click()

If prava("select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'EPR_NOTE' and PRIVILEGE = 'INSERT'") Then
    
    Poznamka_all.VstupnyUserform = Klient_exe
    Poznamka_all.Show
    
Else
    Call text_information_hide(Me.name)
    
    Label29.Caption = "Nemáte práva pro pøidání poznámky"
    Label28.Visible = True
    Label29.Visible = True
End If

End Sub

Public Sub CommandButton25_Click()
' public kvuli debug_risk_1135
If cs_HDWMN.State = 1 Then
prihlasen:

' zobrazení informací

DataCheck.Show

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

Private Sub CommandButton26_Click()

    Dim fso
    Dim AFolder As String, BFolder As String
    AFolder = cfg.Cesta("ExePodane") & c.Nazev_slozky_bez_lomitka
 
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(AFolder) Then
        
        If Not fso.FolderExists(BFolder) Then
            
            ' pokud slozka neexistuje, vytvorim ji
            
            If MsgBox("Složka pro klienta " & Label3.Caption & " neexistuje. Má ji Jezevèík vytvoøit?", vbYesNo, "Složka nenalezena!") = vbYes Then
                
                strMsgbox = ""
                Call GenerateClientFoldersEXE
                VypisVsechnyHlasky
                
                If fso.FolderExists(BFolder) Then
                    Shell "explorer.exe " & BFolder, vbNormalFocus
                End If
                
            End If
            
        Else
            Call Shell("explorer.exe " & BFolder, vbNormalFocus)
        End If
        
    Else
        Call Shell("explorer.exe " & AFolder, vbNormalFocus)
    End If
 
End Sub


Private Sub CommandButton27_Click()
'odkaz_Klient_exe = True
Klient_dedictvi.Show
End Sub

Private Sub CommandButton28_Click()
'odkaz_Klient_exe = True
Klient_trest.Show
End Sub

Private Sub CommandButton29_Click()
'odkaz_Klient_exe = True
Klient_insol.Show
End Sub

Private Sub CommandButton30_Click()
'odkaz_Klient_exe = True
Klient.Show
End Sub

Public Sub CommandButton31_Click()
If cs_HDWMN.State = 1 Then
prihlasen:

On Error GoTo chyba
    'Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' RISK-977, jez.v.5.3
'    If InStr(schema, "TEST") > 0 Then
        Call export_exe_navrh_NEW
'    Else
'        Call export_exe_navrh
'    End If
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

chyba:

' pokud nastala chyba pøipojením, tak ...
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

Private Sub CheckBox1_Click()
Call history_filter
End Sub

Private Sub CheckBox2_Click()
Call history_filter
End Sub

Private Sub CheckBox3_Click()
Call history_filter
End Sub

Private Sub CheckBox4_Click()
Call history_filter
End Sub

Private Sub CheckBox5_Click()
Call history_filter
End Sub



'Private Sub Label2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Call text_information_hide(Me.name)
'
'Dim clipboard As MSForms.DataObject
' Set clipboard = New MSForms.DataObject
' clipboard.SetText Label2.Caption
' clipboard.PutInClipboard
'
' Label29.Caption = "CUID byl zkopírovaný do schránky"
' Label27.Visible = True
' Label29.Visible = True
'
'End Sub
'
'Private Sub Label91_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Call text_information_hide(Me.name)
'
'Dim clipboard As MSForms.DataObject
' Set clipboard = New MSForms.DataObject
' clipboard.SetText Label91.Caption
' clipboard.PutInClipboard
'
' Label29.Caption = "Èíslo rámcové smlouvy bylo zkopírovano do schránky"
' Label27.Visible = True
' Label29.Visible = True
'
'End Sub
'
'
'
'Private Sub Label3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Call text_information_hide(Me.name)
'
'Dim clipboard As MSForms.DataObject
' Set clipboard = New MSForms.DataObject
' clipboard.SetText Label3.Caption
' clipboard.PutInClipboard
'
' Label29.Caption = "Jméno bylo zkopírovano do schránky"
' Label27.Visible = True
' Label29.Visible = True
'
'End Sub
'
'
'Private Sub Label4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Call text_information_hide(Me.name)
'
'Dim clipboard As MSForms.DataObject
' Set clipboard = New MSForms.DataObject
' clipboard.SetText Label4.Caption
' clipboard.PutInClipboard
'
' Label29.Caption = "Rodné èíslo bylo zkopírované do schránky"
' Label27.Visible = True
' Label29.Visible = True
'
'End Sub
'
'Private Sub Label5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Call text_information_hide(Me.name)
'
'Dim clipboard As MSForms.DataObject
' Set clipboard = New MSForms.DataObject
' clipboard.SetText Label5.Caption
' clipboard.PutInClipboard
'
' Label29.Caption = "Trvalá adresa byla zkopírovaná do schránky"
' Label27.Visible = True
' Label29.Visible = True
'
'End Sub
'
'Private Sub Label6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Call text_information_hide(Me.name)
'
'Dim clipboard As MSForms.DataObject
' Set clipboard = New MSForms.DataObject
' clipboard.SetText Label6.Caption
' clipboard.PutInClipboard
'
' Label29.Caption = "Kontaktní adresa byla zkopírovaná do schránky"
' Label27.Visible = True
' Label29.Visible = True
'
'End Sub
'
'Private Sub Label95_DblClick(ByVal Cancel As MSForms.ReturnBoolean) 'RISK-403
'Call text_information_hide(Me.name)
'
'Dim clipboard As MSForms.DataObject
' Set clipboard = New MSForms.DataObject
' clipboard.SetText Label95.Caption
' clipboard.PutInClipboard
'
' Label29.Caption = "Telefonní èíslo bylo zkopírováno do schránky"
' Label27.Visible = True
' Label29.Visible = True
'
'End Sub
'
'Private Sub Label96_DblClick(ByVal Cancel As MSForms.ReturnBoolean) 'RISK-403
'Call text_information_hide(Me.name)
'
'Dim clipboard As MSForms.DataObject
' Set clipboard = New MSForms.DataObject
' clipboard.SetText Label96.Caption
' clipboard.PutInClipboard
'
' Label29.Caption = "E-mailová adresa byla zkopírovaná do schránky"
' Label27.Visible = True
' Label29.Visible = True
'
'End Sub

Private Sub ListView1_Click()
 Call text_information_hide(Me.name)
TextBox1.text = ""

If ListView1.ListItems.count > 0 Then

 TextBox1.text = ListView1.SelectedItem.SubItems(3)
 Label29.Caption = "Zobrazena poznámka z pøehledu"
 Label27.Visible = True
 Label29.Visible = True
 
 End If
End Sub

Private Sub ListView1_DblClick()
Dim sql As String


If cs_HDWMN.State = 1 Then
prihlasen:
TextBox1.text = ""
 Call text_information_hide(Me.name)
If (ListView1.ListItems.count > 0) Then
    
    If (ListView1.SelectedItem.SubItems(4) = "Ano" Or Len(ListView1.SelectedItem.SubItems(3)) > 130) Then
     
        sql = "select a.text_note from " & schema & ".exe_status_data a join " & schema & ".exe_note b on a.skp_exe_status_data = b.skp_exe_status_data and b.skp_exe_note = " & ListView1.SelectedItem
         
        Call query1(sql, "Note_exe")
        
         
        Label29.Caption = "Poznámka zobrazena"
        Label27.Visible = True
        Label29.Visible = True

    ElseIf ListView1.SelectedItem.SubItems(4) = "Ne" Then
        Label29.Caption = "K poznámce nejsou další informace"
        Label28.Visible = True
        Label29.Visible = True
    Else
        Label29.Caption = "Není pro co zobrazovat"
        Label28.Visible = True
        Label29.Visible = True
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


Private Sub ListView2_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
If (KeyCode = 67 And Shift = 2) Then

Call text_information_hide(Me.name)

Dim clipboard As MSForms.DataObject
Set clipboard = New MSForms.DataObject
clipboard.SetText ListView2.SelectedItem
clipboard.PutInClipboard

 Label29.Caption = "Èíslo smlouvy bylo zkopírovano do schránky"
 Label27.Visible = True
 Label29.Visible = True

End If
End Sub


Private Sub ListView3_DblClick()

If cs_HDWMN.State = 1 Then
prihlasen:

Call text_information_hide(Me.name)
If ListView3.ListItems.count > 0 Then

    Upozorneni_all.VstupnyUserform = Klient_exe
    Upozorneni_all.VstupNoveUpo = False
    Upozorneni_all.Show

Else
    Label29.Caption = "Není pro co zobrazovat"
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



Private Sub ListView4_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
If (KeyCode = 67 And Shift = 2) Then

Call text_information_hide(Me.name)

Dim clipboard As MSForms.DataObject
Set clipboard = New MSForms.DataObject
clipboard.SetText ListView4.SelectedItem.SubItems(1)
clipboard.PutInClipboard

 Label29.Caption = "Hodnota '" & ListView4.SelectedItem & "' byla zkopírovana do schránky"
 Label27.Visible = True
 Label29.Visible = True

End If
End Sub

Private Sub ListView5_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
If (KeyCode = 67 And Shift = 2) Then

Call text_information_hide(Me.name)

Dim clipboard As MSForms.DataObject
Set clipboard = New MSForms.DataObject
clipboard.SetText ListView5.SelectedItem.SubItems(1)
clipboard.PutInClipboard

 Label29.Caption = "Hodnota '" & ListView5.SelectedItem & "' byla zkopírovana do schránky"
 Label27.Visible = True
 Label29.Visible = True

End If
End Sub

' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Public Sub UserForm_Activate()
If cs_HDWMN.State = 1 Then
prihlasen:
'----------------------- inicializace klienta pres objekt---------------
If odkaz_Klient = False Then
    'Dim c As Client
    Set c = New Client
    Call c.initialize(Val(id_person))
    Call c.find_accounts
    Call c.find_loans
End If
'----------------------- inicializace klienta pres objekt---------------
 
 Call text_information_hide(Me.name)
 
  
 Call query1("", Me.name & "History")
 Call query1("", Me.name & "Other")
 
If ListView1.ListItems.count > 0 Then
 Label29.Caption = "Nalezeno " & ListView1.ListItems.count & " záznamù"
 Label27.Visible = True
 Label29.Visible = True
 Label28.Visible = False
Else
 Label29.Caption = "Nalezeno " & ListView1.ListItems.count & " záznamù"
 Label28.Visible = True
 Label29.Visible = True
End If
 
 
 If userform_prv = "SouhrnAuth" Then
  userform_prv = "Souhrn"
  'Call CommandButton24_Click
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

'RISK-721
 ' RISK-985 odkomentovat po hromadnom REPLACE
If Sonik_User Then
CommandButton_Sonik.Enabled = True
Else: CommandButton_Sonik.Enabled = False
End If
 ' RISK-985 odkomentovat po hromadnom REPLACE

End Sub

Private Sub UserForm_Initialize()
Call table_prepare(Me.name)
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'If odkaz_Klient = False Then

Me.Hide

Select Case userform_prv
 Case "Zaloby"
  Klient.Show
    odkaz_Klient = True
Case "Exekuce"
  Exekuce.Show
    odkaz_Klient = False
 Case "Dedictvi"
  Klient_dedictvi.Show
  odkaz_Klient = True
Case "Insol"
  Klient_insol.Show
  odkaz_Klient = True
Case "Trest"
  Klient_trest.Show
    odkaz_Klient = True
  Case "Zesplatneni"
  Zesplatneni.Show
    odkaz_Klient = False
 Case "Nezesplatneni"
  Nezesplatneni.Show
    odkaz_Klient = False
 Case "SouhrnNotice", "SouhrnAuth", "Souhrn"
  Souhrn.Show
      odkaz_Klient = False
End Select

'Else

Cancel = True

'End If

TextBox1.text = ""

End Sub


Private Sub Label1_Click()
Call CopyToClipboard(Label1.Caption, Me.name)
Label29.Caption = "ID_PERSON bylo zkopírovano do schránky"
End Sub

Private Sub Label2_Click()
Call CopyToClipboard(Label2.Caption, Me.name)
Label29.Caption = "CUID byl zkopírovaný do schránky"
End Sub

Private Sub Label3_Click()
Call CopyToClipboard(Label3.Caption, Me.name)
Label29.Caption = "Jméno bylo zkopírovano do schránky"
End Sub

Private Sub Label4_Click()
Call CopyToClipboard(Label4.Caption, Me.name)
Label29.Caption = "Rodné èíslo bylo zkopírované do schránky"
End Sub

Private Sub Label5_Click()
Call CopyToClipboard(Label5.Caption, Me.name)
Label29.Caption = "Trvalá adresa byla zkopírovaná do schránky"
End Sub

Private Sub Label6_Click()
Call CopyToClipboard(Label6.Caption, Me.name)
Label29.Caption = "Kontaktní adresa byla zkopírovaná do schránky"
End Sub

Private Sub Label91_Click()
Call CopyToClipboard(Label91.Caption, Me.name)
Label29.Caption = "Èíslo rámcové smlouvy bylo zkopírovano do schránky"
End Sub

Private Sub Label95_Click()
Call CopyToClipboard(Label95.Caption, Me.name)
Label29.Caption = "Telefonní èíslo bylo zkopírováno do schránky"
End Sub

Private Sub Label96_Click()
Call CopyToClipboard(Label96.Caption, Me.name)
Label29.Caption = "E-mailová adresa byla zkopírovaná do schránky"
End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton19.ForeColor = barvaCerna
CommandButton18.ForeColor = barvaCerna
CommandButtonVlozUpo.ForeColor = barvaCerna
ButtonVlozPoznamku.ForeColor = barvaCerna
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton20.ForeColor = barvaCerna
End Sub

Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton16.ForeColor = barvaCerna
End Sub

Private Sub CommandButton16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton16.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton18.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton19.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton20.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButtonVlozUpo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButtonVlozUpo.ForeColor = barvaZelenaAirBank
End Sub

Private Sub ButtonVlozPoznamku_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
ButtonVlozPoznamku.ForeColor = barvaZelenaAirBank
End Sub



