VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Klient 
   Caption         =   "Jezevèík - Nalézací øízení"
   ClientHeight    =   11775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   22815
   OleObjectBlob   =   "Klient.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Klient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit


Public Sub test_me(ByVal id_person_in As Long)
    
On Error GoTo ERROR_HANDLER
    
    id_person = id_person_in
    ' Call UserForm_Initialize
    Call UserForm_Activate
    Me.Show
    
    ' sablony
    Call CommandButton4_Click
    Call CommandButton31_Click
    Call CommandButtonXmlEpr_Click
    Call CommandButtonVyznaceniDolozky_Click
    ' Call CommandButton18_Click
    
NORMAL_FLOW:
    Exit Sub

ERROR_HANDLER:
    Debug.Assert 1 = 0
    Debug.Print "Spustil se error handler v objektu " & Me.name & vbNewLine & _
                 Err.Number & ": " & Err.Description
    
End Sub

Sub history_filter()
If cs_HDWMN.State = 1 Then
prihlasen:
 Call text_information_hide(Me.name)
 
 Call query("", Me.name & "History")
 
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

Private Sub button_Vytvor_EPR_PDF_Click()
    If schema = "APP_COLL_JEZEVCIK_TEST" Or (InStr(strAdmins, UCase(Main.username)) > 0) Or (InStr(UCase(ThisWorkbook.Path), UCase("test")) > 0) Then
        Call Vytvor_PDF_komplet(rizeni:=typRizeni_EPR)
    End If
End Sub

Private Sub CommandButton_Chance_Import_Click()
'RISK-778
'Nastavím si z jakého formuláøe se šance volá - je to property na userformu Zadani_Sance
Zadani_Sance.VstupnyUserform = Klient
Zadani_Sance.Show
End Sub

Private Sub CommandButton_Sonik_Click()
' RISK-721
 Call Client_From_Jez_To_Sonik(Klient.name, 3) ' RISK-985 odkomentovat po hromadnom REPLACE
End Sub

Private Sub CommandButtonVlozUpo_Click()
If cs_HDWMN.State = 1 Then
prihlasen:

If prava("select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'NOTICE' and PRIVILEGE = 'INSERT'") Then

    Upozorneni_all.VstupnyUserform = Klient
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
   Update_stav_all.VstupnyUserform = Klient
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
Case "Zaloby"
    Zaloby.Show
    odkaz_Klient = False
    
Case "Exekuce"
    Klient_exe.Show
    odkaz_Klient = True
    
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
    
Case "Klient_heir"
    Klient_heir.Show

Case Else
    Main.Show
    
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
 
  Update_klient.Label1.Caption = Label3.Caption
  Update_klient.Label2.Caption = Label4.Caption
  
  '''  XXX RISK-1197 - tady je potreba nainicializovat labels pro okresni a krajsky soud
 
  For i = 1 To ListView5.ListItems.count
   Select Case ListView5.ListItems(i)
    Case "Datum žaloby"
     Update_klient.Label6.Caption = ListView5.ListItems(i).SubItems(1)
    Case "Spisová znaèka"
     Update_klient.Label5.Caption = ListView5.ListItems(i).SubItems(1)
    Case "Výše pohledávky"
     Update_klient.Label7.Caption = ListView5.ListItems(i).SubItems(1)
'    Case "Datum vyèíslení pohledávky"
'     Update_klient.Label8.Caption = ListView5.ListItems(i).SubItems(1)
' Case "ID Datové schránky"
'   Update_klient.Label9.Caption = ListView5.ListItems(i).SubItems(1)
   End Select
  Next i
 
  Update_klient.Show
 
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

'Private Sub CommandButton20_Click()
'Call CommandButton24_Click
'End Sub

Private Sub ButtonVlozPoznamku_Click()

If prava("select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'EPR_NOTE' and PRIVILEGE = 'INSERT'") Then
    
    Poznamka_all.VstupnyUserform = Klient
    Poznamka_all.Show
    
Else
    Call text_information_hide(Me.name)
    
    Label29.Caption = "Nemáte práva pro pøidání poznámky"
    Label28.Visible = True
    Label29.Visible = True
End If

End Sub


Private Sub CommandButton23_Click()
If cs_HDWMN.State = 1 Then
prihlasen:


    If prava("select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'SOP' and PRIVILEGE = 'INSERT'") Then
        
        With Update_SOP
            .Label1.Caption = Empty_string_for_default_value(c.SOP.AMT_SOP1, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)               '  Klient.ListView4.ListItems(1).SubItems(1)
            .Label3.Caption = Empty_string_for_default_value(c.SOP.AMT_ADDITIONAL_SOP1, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)    '  Klient.ListView4.ListItems(4).SubItems(1)
            .Label5.Caption = Empty_string_for_default_value(c.SOP.AMT_RETURN_SOP1, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)        '  Klient.ListView4.ListItems(6).SubItems(1)
            .Label6.Caption = Empty_string_for_default_value(c.SOP.DATE_INFO_RETURN_SOP1, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)  '  Klient.ListView4.ListItems(7).SubItems(1)
        
            .Label46.Caption = Empty_string_for_default_value(c.SOP.AMT_SOP1_COURT, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)        '  Klient.ListView4.ListItems(10).SubItems(1)
            .Label9.Caption = Empty_string_for_default_value(c.SOP.AMT_SOP2, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)               '  Klient.ListView4.ListItems(11).SubItems(1)
            .Label11.Caption = Empty_string_for_default_value(c.SOP.AMT_ADDITIONAL_SOP2, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)   '  Klient.ListView4.ListItems(13).SubItems(1)
            .Label13.Caption = Empty_string_for_default_value(c.SOP.AMT_RETURN_SOP2, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)       '  Klient.ListView4.ListItems(15).SubItems(1)
            .Label14.Caption = Empty_string_for_default_value(c.SOP.DATE_INFO_RETURN_SOP2, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE) '  Klient.ListView4.ListItems(16).SubItems(1)
        
            .Label47.Caption = Empty_string_for_default_value(c.SOP.AMT_SOP2_COURT, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)        '  Klient.ListView4.ListItems(19).SubItems(1)
            .Label50.Caption = Empty_string_for_default_value(c.SOP.SOP_ADDITIONAL_PAID, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)   '  Klient.ListView4.ListItems(20).SubItems(1)
            .Label34.Caption = Empty_string_for_default_value(c.SOP.TRAVEL_EXPENSES, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)       '  Klient.ListView4.ListItems(21).SubItems(1)
            .Label51.Caption = Empty_string_for_default_value(c.SOP.SOP_ADDITIONAL_FEE, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)    '  Klient.ListView4.ListItems(3).SubItems(1) ' pausal
        End With
        
        Update_SOP.Show
    Else
        Call text_information_hide(Me.name)
        
        Label29.Caption = "Nemáte práva pro zmìnu soudních poplatkù"
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

Private Sub CommandButton20_Click()
If cs_HDWMN.State = 1 Then
prihlasen:

' RISK-1101, jez.v.5.4.1 - je potreba komplet prepsat, hodit sem vlastni SQL select

Update_auth.Label1.Caption = Empty_string_for_default_value(c.SOP.AMT_SOP1, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)                ' Vyse SOP1                    ' Klient.ListView4.ListItems(1).SubItems(1)
Update_auth.Label2.Caption = Empty_string_for_default_value(c.SOP.AMT_ADDITIONAL_SOP1, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)     ' Vyse doplatku SOP 1          ' Klient.ListView4.ListItems(3).SubItems(1)
Update_auth.Label31.Caption = Empty_string_for_default_value(c.SOP.AMT_RETURN_SOP1, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)        ' vyse preplatku SOP1          ' Klient.ListView4.ListItems(5).SubItems(1)
Update_auth.Label32.Caption = Empty_string_for_default_value(c.SOP.DATE_INFO_RETURN_SOP1, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)  ' datum sdeleni preplatku SOP1 ' Klient.ListView4.ListItems(6).SubItems(1)

Update_auth.Label4.Caption = Empty_string_for_default_value(c.SOP.AMT_SOP2, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)                ' Vyse SOP1                    '  Klient.ListView4.ListItems(10).SubItems(1)
Update_auth.Label5.Caption = Empty_string_for_default_value(c.SOP.AMT_ADDITIONAL_SOP2, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)     ' Vyse doplatku SOP 1          '  Klient.ListView4.ListItems(12).SubItems(1)
Update_auth.Label61.Caption = Empty_string_for_default_value(c.SOP.AMT_RETURN_SOP2, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)        ' vyse preplatku SOP1          '  Klient.ListView4.ListItems(14).SubItems(1)
Update_auth.Label62.Caption = Empty_string_for_default_value(c.SOP.DATE_INFO_RETURN_SOP2, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)  ' datum sdeleni preplatku SOP1 '  Klient.ListView4.ListItems(15).SubItems(1)

' Autorizace SOP 1: Výše --------------------------------------------------------------------------
' If Klient.ListView4.ListItems(2).SubItems(1) <> "" Then ' RISK-1101
' If c.sop.DATE_SOP1 <> #12:00:00 AM# Then
If c.SOP.Auth_SOP1 = stavAuth_Autorizovano Then
    ' Autorizováno
    Update_auth.Frame1.BackColor = barvaZelenaAirBank
    Update_auth.Frame1.Height = 60
    
    Update_auth.CheckBox1.value = True
    Update_auth.CheckBox1.Enabled = False
    Update_auth.CheckBox1.Caption = "Autorizováno"
    
    Update_auth.DTPicker1 = c.SOP.DATE_SOP1 ' Klient.ListView4.ListItems(2).SubItems(1)
    Update_auth.DTPicker1.Enabled = False
'ElseIf Klient.ListView4.ListItems(1).SubItems(1) <> "" Then ' RISK-1101
' ElseIf c.sop.AMT_SOP1 <> 0 Then
ElseIf c.SOP.Auth_SOP1 = stavAuth_Ceka_na_autorizaci Then
    ' Nutno autorizovat
    Update_auth.Frame1.BackColor = barvaLososova
    Update_auth.Frame1.Height = 24
    
    Update_auth.CheckBox1.value = False
    Update_auth.CheckBox1.Enabled = True
    Update_auth.CheckBox1.Caption = "Autorizovat"
    
    Update_auth.DTPicker1 = date
    Update_auth.DTPicker1.Enabled = True
Else
    ' No action
    Update_auth.Frame1.BackColor = barvaSedaSvetla
    Update_auth.Frame1.Height = 24
    
    Update_auth.CheckBox1.value = False
    Update_auth.CheckBox1.Enabled = False
    Update_auth.CheckBox1.Caption = "Autorizovat"
    
    Update_auth.DTPicker1 = date
    Update_auth.DTPicker1.Enabled = False
End If

' Autorizace SOP 1: Doplatek --------------------------------------------------------------------------
'If Klient.ListView4.ListItems(4).SubItems(1) <> "" Then ' RISK-1101
' If c.sop.DATE_ADDITIONAL_SOP1 <> #12:00:00 AM# Then
If c.SOP.Auth_Additional_SOP1 = stavAuth_Autorizovano Then
    ' Autorizováno
    Update_auth.Frame2.BackColor = barvaZelenaAirBank
    Update_auth.Frame2.Height = 60
    Update_auth.CheckBox2.value = True
    Update_auth.CheckBox2.Enabled = False
    Update_auth.CheckBox2.Caption = "Autorizováno"
    
    Update_auth.DTPicker2 = c.SOP.DATE_ADDITIONAL_SOP1 ' Klient.ListView4.ListItems(4).SubItems(1)
    Update_auth.DTPicker2.Enabled = False
' ElseIf Klient.ListView4.ListItems(3).SubItems(1) <> "" Then ' RISK-1101
' ElseIf c.sop.AMT_ADDITIONAL_SOP1 <> 0 Then
ElseIf c.SOP.Auth_Additional_SOP1 = stavAuth_Ceka_na_autorizaci Then
    ' Nutno autorizovat
    Update_auth.Frame2.BackColor = barvaLososova
    Update_auth.Frame2.Height = 24
    Update_auth.CheckBox2.value = False
    Update_auth.CheckBox2.Enabled = True
    Update_auth.CheckBox2.Caption = "Autorizovat"
    
    Update_auth.DTPicker2 = date
    Update_auth.DTPicker2.Enabled = True
Else
    ' No action
    Update_auth.Frame2.BackColor = barvaSedaSvetla
    Update_auth.Frame2.Height = 24
    Update_auth.CheckBox2.value = False
    Update_auth.CheckBox2.Enabled = False
    Update_auth.CheckBox2.Caption = "Autorizovat"
    
    Update_auth.DTPicker2 = date
    Update_auth.DTPicker2.Enabled = False
End If

' Autorizace SOP 1: Pøeplatek --------------------------------------------------------------------------
' If Klient.ListView4.ListItems(7).SubItems(1) = Klient.ListView4.ListItems(5).SubItems(1) And Klient.ListView4.ListItems(7).SubItems(1) <> "" Then ' RISK-1101
' If c.sop.AMT_RETURN_REAL_SOP1 = c.sop.AMT_RETURN_SOP1 And c.sop.AMT_RETURN_REAL_SOP1 <> 0 Then
If c.SOP.Auth_Return_SOP1 = stavAuth_Autorizovano Then
    ' Autorizováno
    Update_auth.Frame3.BackColor = barvaZelenaAirBank
    Update_auth.Frame3.Height = 108
    Update_auth.CheckBox3.value = True
    Update_auth.CheckBox3.Enabled = False
    Update_auth.CheckBox3.Caption = "Autorizováno"
    
    Update_auth.TextBox3.text = c.SOP.AMT_RETURN_SOP1   ' Klient.ListView4.ListItems(5).SubItems(1)
    Update_auth.TextBox3.Enabled = False
    Update_auth.DTPicker3 = c.SOP.DATE_INFO_RETURN_SOP1 ' Klient.ListView4.ListItems(6).SubItems(1)
    Update_auth.DTPicker3.Enabled = False
' upppIf Klient.ListView4.ListItems(7).SubItems(1) = Klient.ListView4.ListItems(5).SubItems(1)  And Klient.ListView4.ListItems(7).SubItems(1) <> "" Then 'porovnavam s IF ktery je o 10 radku vyse
' ElseIf Klient.ListView4.ListItems(7).SubItems(1) <> Klient.ListView4.ListItems(5).SubItems(1) And Klient.ListView4.ListItems(5).SubItems(1) <> "" Then ' RISK-1101
' ElseIf c.sop.AMT_RETURN_REAL_SOP1 <> c.sop.AMT_RETURN_SOP1 And c.sop.AMT_RETURN_SOP1 <> 0 Then ' RISK-1101
ElseIf c.SOP.Auth_Return_SOP1 = stavAuth_Ceka_na_autorizaci Then
    ' Nutno autorizovat
    Update_auth.Frame3.BackColor = barvaLososova
    Update_auth.Frame3.Height = 48
    Update_auth.CheckBox3.value = False
    Update_auth.CheckBox3.Enabled = True
    Update_auth.CheckBox3.Caption = "Autorizovat"
    Update_auth.TextBox3.Enabled = True
    Update_auth.TextBox3.text = ""
    Update_auth.DTPicker3 = date
    Update_auth.DTPicker3.Enabled = True
Else
    ' No action
    Update_auth.Frame3.BackColor = barvaSedaSvetla
    Update_auth.Frame3.Height = 48
    Update_auth.CheckBox3.value = False
    Update_auth.CheckBox3.Enabled = False
    Update_auth.CheckBox3.Caption = "Autorizovat"
     Update_auth.TextBox3.Enabled = True
    Update_auth.TextBox3.text = ""
    Update_auth.DTPicker3 = date
    Update_auth.DTPicker3.Enabled = False
End If

' Autorizace SOP 2: Výše --------------------------------------------------------------------------
' If Klient.ListView4.ListItems(11).SubItems(1) <> "" Then ' RISK-1101
' If c.sop.DATE_SOP2 <> #12:00:00 AM# Then ' RISK-1101
If c.SOP.Auth_SOP2 = stavAuth_Autorizovano Then
    ' Autorizováno
    Update_auth.Frame4.BackColor = barvaZelenaAirBank
    Update_auth.Frame4.Height = 60
    
    Update_auth.CheckBox4.value = True
    Update_auth.CheckBox4.Enabled = False
    Update_auth.CheckBox4.Caption = "Autorizováno"
    
    Update_auth.DTPicker4 = c.SOP.DATE_SOP2 'Klient.ListView4.ListItems(11).SubItems(1)
    Update_auth.DTPicker4.Enabled = False
' ElseIf Klient.ListView4.ListItems(10).SubItems(1) <> "" Then ' RISK-1101
' ElseIf c.sop.AMT_SOP2 <> 0 Then
ElseIf c.SOP.Auth_SOP2 = stavAuth_Ceka_na_autorizaci Then
    ' Nutno autorizovat
    Update_auth.Frame4.BackColor = barvaLososova
    Update_auth.Frame4.Height = 24
    
    Update_auth.CheckBox4.value = False
    Update_auth.CheckBox4.Enabled = True
    Update_auth.CheckBox4.Caption = "Autorizovat"
    
    Update_auth.DTPicker4 = date
    Update_auth.DTPicker4.Enabled = True
Else
    ' No action
    Update_auth.Frame4.BackColor = barvaSedaSvetla
    Update_auth.Frame4.Height = 24
    
    Update_auth.CheckBox4.value = False
    Update_auth.CheckBox4.Enabled = False
    Update_auth.CheckBox4.Caption = "Autorizovat"
    
    Update_auth.DTPicker4 = date
    Update_auth.DTPicker4.Enabled = False
End If

' Autorizace SOP 2: Doplatek --------------------------------------------------------------------------
' If Klient.ListView4.ListItems(13).SubItems(1) <> "" Then ' RISK-1101
' If c.sop.DATE_ADDITIONAL_SOP2 <> #12:00:00 AM# Then
If c.SOP.Auth_Additional_SOP2 = stavAuth_Autorizovano Then
    '' XX FA tu si skoncil RISK-1101
     ' Autorizováno
     
     Update_auth.Frame5.BackColor = barvaZelenaAirBank
     Update_auth.Frame5.Height = 60
     Update_auth.CheckBox5.value = True
     Update_auth.CheckBox5.Enabled = False
     Update_auth.CheckBox5.Caption = "Autorizováno"
     
     Update_auth.DTPicker5 = c.SOP.DATE_ADDITIONAL_SOP2 ' Klient.ListView4.ListItems(12).SubItems(1) ' datum doplatku SOP2 ' RISK-1101 nejak mi to nesedi! ukazuje to na amt_additional_sop2
     Update_auth.DTPicker5.Enabled = False
'ElseIf Klient.ListView4.ListItems(12).SubItems(1) <> "" Then ' RISK-1101
' ElseIf c.sop.AMT_ADDITIONAL_SOP2 <> 0 Then
ElseIf c.SOP.Auth_Additional_SOP2 = stavAuth_Ceka_na_autorizaci Then
    ' Nutno autorizovat
    Update_auth.Frame5.BackColor = barvaLososova
    Update_auth.Frame5.Height = 24
    Update_auth.CheckBox5.value = False
    Update_auth.CheckBox5.Enabled = True
    Update_auth.CheckBox5.Caption = "Autorizovat"
    
    Update_auth.DTPicker5 = date
    Update_auth.DTPicker5.Enabled = True
Else
    ' No action
    Update_auth.Frame5.BackColor = barvaSedaSvetla
    Update_auth.Frame5.Height = 24
    Update_auth.CheckBox5.value = False
    Update_auth.CheckBox5.Enabled = False
    Update_auth.CheckBox5.Caption = "Autorizovat"
    
    Update_auth.DTPicker5 = date
    Update_auth.DTPicker5.Enabled = False
End If

' Autorizace SOP 2: Pøeplatek --------------------------------------------------------------------------

'If Klient.ListView4.ListItems(16).SubItems(1) = Klient.ListView4.ListItems(14).SubItems(1) And Klient.ListView4.ListItems(16).SubItems(1) <> "" Then ' RISK-1101
' If c.sop.AMT_RETURN_REAL_SOP2 = c.sop.AMT_RETURN_SOP2 And c.sop.AMT_RETURN_REAL_SOP2 <> 0 Then
If c.SOP.Auth_Return_SOP2 = stavAuth_Autorizovano Then
    ' Autorizováno
    Update_auth.Frame6.BackColor = barvaZelenaAirBank
    Update_auth.Frame6.Height = 108
    Update_auth.CheckBox6.value = True
    Update_auth.CheckBox6.Enabled = False
    Update_auth.CheckBox6.Caption = "Autorizováno"
    
    Update_auth.TextBox6.text = c.SOP.AMT_RETURN_REAL_SOP2   ' Klient.ListView4.ListItems(15).SubItems(1) ' RISK-1101 prohodil jsem hodnotu s tou nasledujici
    Update_auth.TextBox6.Enabled = False
    Update_auth.DTPicker6 = c.SOP.DATE_INFO_RETURN_SOP2      ' Klient.ListView4.ListItems(16).SubItems(1) ' RISK-1101 prohodil jsem hodnotu s tou predeslou
    Update_auth.DTPicker6.Enabled = False
' upupIf Klient.ListView4.ListItems(16).SubItems(1) = Klient.ListView4.ListItems(14).SubItems(1)  And Klient.ListView4.ListItems(16).SubItems(1) <> "" Then ' RISK-1101
' ElseIf Klient.ListView4.ListItems(16).SubItems(1) <> Klient.ListView4.ListItems(14).SubItems(1) And Klient.ListView4.ListItems(14).SubItems(1) <> "" Then '  RISK-1101
' ElseIf c.sop.AMT_RETURN_REAL_SOP2 <> c.sop.AMT_RETURN_SOP2 And c.sop.AMT_RETURN_SOP2 <> 0 Then
ElseIf c.SOP.Auth_Return_SOP2 = stavAuth_Ceka_na_autorizaci Then
    ' Nutno autorizovat
    Update_auth.Frame6.BackColor = barvaLososova
    Update_auth.Frame6.Height = 48
    Update_auth.CheckBox6.value = False
    Update_auth.CheckBox6.Enabled = True
    Update_auth.CheckBox6.Caption = "Autorizovat"
    Update_auth.TextBox6.Enabled = True
    Update_auth.TextBox6.text = ""
    Update_auth.DTPicker6 = date
    Update_auth.DTPicker6.Enabled = True
Else
    ' No action
    Update_auth.Frame6.BackColor = barvaSedaSvetla
    Update_auth.Frame6.Height = 48
    Update_auth.CheckBox6.value = False
    Update_auth.CheckBox6.Enabled = False
    Update_auth.CheckBox6.Caption = "Autorizovat"
    Update_auth.TextBox6.Enabled = False
    Update_auth.TextBox6.text = ""
    Update_auth.DTPicker6 = date
    Update_auth.DTPicker6.Enabled = False
End If

 Update_auth.Show
 
 
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


Private Sub CommandButton25_Click()
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
    
    AFolder = cfg.Cesta("EprPodaneZaloby") & c.Nazev_slozky_bez_lomitka
    BFolder = cfg.Cesta("EprNachystane") & c.Nazev_slozky_bez_lomitka
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(AFolder) Then
        If Not fso.FolderExists(BFolder) Then
            MsgBox "Složka pro klienta " & Label3.Caption & " neexistuje", vbInformation, "Složka nenalezena!"
        Else
            Call Shell("explorer.exe " & BFolder, vbNormalFocus)
        End If
    Else
        Call Shell("explorer.exe " & AFolder, vbNormalFocus)
    End If
 
End Sub


Private Sub CommandButton27_Click()

odkaz_Klient = True

Klient_dedictvi.Show

End Sub

Private Sub CommandButton28_Click()

odkaz_Klient = True

Klient_trest.Show

End Sub

Private Sub CommandButton29_Click()

odkaz_Klient = True

Klient_insol.Show

End Sub

Private Sub CommandButton30_Click()
odkaz_Klient = True

Klient_exe.Show
End Sub

Private Sub CommandButton31_Click()

If cs_HDWMN.State = 1 Then
prihlasen:

On Error GoTo chyba




   Application.EnableEvents = False
   Application.ScreenUpdating = False
   Application.DisplayAlerts = False
   
'Dim sql712 As String, skp As Variant
'
'Set rst712 = New ADODB.Recordset
''zjistime skp klienta
'   sql712 = "select id_person from " & schema & ".INTERVALY where code_loan_number = '" & ListView1.SelectedItem & "'"
'   rst712.Open sql712, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'
'   If Not rst712.EOF Then
'    skp = rst712(0)
'   End If
'rst712.Close


Call export_podani(id_person)



   Application.EnableEvents = True
   Application.ScreenUpdating = True
   Application.DisplayAlerts = True
   
' tady vypise pouze splatky, u kterych prisla platba do zesplatneni
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

Private Sub CommandButton4_Click()

If cs_HDWMN.State = 1 Then
prihlasen:

On Error GoTo chyba




   Application.EnableEvents = False
   Application.ScreenUpdating = False
   Application.DisplayAlerts = False
   
   
' nove se jde primo z karty klienta (drive bylo tlacitko umistene na kontrole dat), neni teda potreba hledat id_person
   
'Dim sql712 As String, skp As Variant

'Set rst712 = New ADODB.Recordset
''zjistime skp klienta
'   sql712 = "select id_person from " & schema & ".INTERVALY where code_loan_number = '" & ListView1.SelectedItem & "'"
'   rst712.Open sql712, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'
'   If Not rst712.EOF Then
'    skp = rst712(0)
'   End If
'rst712.Close

Call export_uznani_dluhu(id_person)



   Application.EnableEvents = True
   Application.ScreenUpdating = True
   Application.DisplayAlerts = True
   
' tady vypise pouze splatky, u kterych prisla platba do zesplatneni
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

Private Sub CommandButtonVyznaceniDolozky_Click()
Call ExportZadostDolozka 'RISK-983
End Sub

Private Sub CommandButtonXmlEpr_Click()
Call generovatEPR_XML
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
'End Sub

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
'End Sub


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
'End Sub

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
'End Sub

'Private Sub Label5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Call text_information_hide(Me.name)
'Call CopyToClipboard(Label5.Caption)
'Dim clipboard As MSForms.DataObject
' Set clipboard = New MSForms.DataObject
' clipboard.SetText Label5.Caption
' clipboard.PutInClipboard
'
' Label29.Caption = "Trvalá adresa byla zkopírovaná do schránky"
' Label27.Visible = True
' Label29.Visible = True
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
' Label29.Caption = "Telefonní èíslo bylo zkopírováno do schránky"
' Label27.Visible = True
' Label29.Visible = True
'End Sub
'
'Private Sub Label98_DblClick(ByVal Cancel As MSForms.ReturnBoolean) 'RISK-403
'Call text_information_hide(Me.name)
'
'Dim clipboard As MSForms.DataObject
' Set clipboard = New MSForms.DataObject
' clipboard.SetText Label98.Caption
' clipboard.PutInClipboard
'
' Label29.Caption = "E-mailová adresa byla zkopírovaná do schránky"
' Label27.Visible = True
' Label29.Visible = True
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
   
        sql = "select a.text_note from " & schema & ".epr_status_data a join " & schema & ".epr_note b on a.skp_epr_status_data = b.skp_epr_status_data and b.skp_epr_note = " & ListView1.SelectedItem
         
        Call query(sql, "Note")
        
         
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

    Upozorneni_all.VstupnyUserform = Klient
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
Private Sub UserForm_Activate()
If cs_HDWMN.State = 1 Then
prihlasen:
TextBox1.value = ""
'----------------------- inicializace klienta pres objekt---------------
If odkaz_Klient = False Then
    'Dim c As Client
    Set c = New Client
    Call c.initialize(Val(id_person))
    Call c.find_accounts
    Call c.find_loans
'MsgBox c.Cuid
End If
'----------------------- inicializace klienta pres objekt---------------


 Call text_information_hide(Me.name)
 
 CommandButton20.Visible = prava("select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'AUTHORIZATION' and PRIVILEGE = 'INSERT'")
 
 Call query("", Me.name & "History")
 Call query("", Me.name & "Other")
 Call query1("", Me.name & "Other") 'funkce query je omezena a nelze vlozit dalsi prikazy na nacteni dalsich dat
If ListView1.ListItems.count > 0 Then
 Label29.Caption = "Nalezeno " & ListView1.ListItems.count & " záznamù"
 Label27.Visible = True
 Label29.Visible = True
Else
 Label29.Caption = "Nalezeno " & ListView1.ListItems.count & " záznamù"
 Label28.Visible = True
 Label29.Visible = True
End If
 
 
 If userform_prv = "SouhrnAuth" Then
  userform_prv = "Souhrn"
  Call CommandButton20_Click
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
''' ' RISK-985 odkomentovat po hromadnom REPLACE
If Sonik_User Then
CommandButton_Sonik.Enabled = True
Else: CommandButton_Sonik.Enabled = False
End If
 ' RISK-985 odkomentovat po hromadnom REPLACE

End Sub

Private Sub UserForm_Initialize()
TextBox1.value = ""
Call table_prepare(Me.name)
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'If odkaz_Klient = False Then

Me.Hide
Select Case userform_prv
 Case "Zaloby"
  Zaloby.Show
  odkaz_Klient = False
Case "Exekuce"
  Klient_exe.Show
  odkaz_Klient = True
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

'Cancel = True

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

Private Sub Label96_Click()
Call CopyToClipboard(Label96.Caption, Me.name)
Label29.Caption = "Telefonní èíslo bylo zkopírováno do schránky"
End Sub

Private Sub Label98_Click()
Call CopyToClipboard(Label98.Caption, Me.name)
Label29.Caption = "E-mailová adresa byla zkopírovaná do schránky"
End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    CommandButton19.ForeColor = barvaCerna
    CommandButtonVlozUpo.ForeColor = barvaCerna
    ButtonVlozPoznamku.ForeColor = barvaCerna
    CommandButton23.ForeColor = barvaCerna
    CommandButton25.ForeColor = barvaCerna
    CommandButton26.ForeColor = barvaCerna
    CommandButton4.ForeColor = barvaCerna
    CommandButton31.ForeColor = barvaCerna
    CommandButtonXmlEpr.ForeColor = barvaCerna
    CommandButton_Chance_Import.ForeColor = barvaCerna 'RISK-778
    CommandButtonVyznaceniDolozky.ForeColor = barvaCerna 'RISK-983
    ButtonZmenitStav.ForeColor = barvaCerna
    Me.button_Vytvor_EPR_PDF.ForeColor = barvaCerna
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

Private Sub CommandButton23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton23.ForeColor = barvaZelenaAirBank
End Sub


Private Sub CommandButton31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton31.ForeColor = barvaZelenaAirBank
End Sub
Private Sub CommandButton4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton4.ForeColor = barvaZelenaAirBank
End Sub
Private Sub CommandButton25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton25.ForeColor = barvaZelenaAirBank
End Sub
Private Sub CommandButton26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton26.ForeColor = barvaZelenaAirBank
End Sub
Private Sub CommandButtonXmlEpr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButtonXmlEpr.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButtonVlozUpo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButtonVlozUpo.ForeColor = barvaZelenaAirBank
End Sub

Private Sub ButtonVlozPoznamku_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
ButtonVlozPoznamku.ForeColor = barvaZelenaAirBank
End Sub

'RISK-778
Private Sub CommandButton_Chance_Import_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Chance_Import.ForeColor = barvaZelenaAirBank
End Sub

'RISK-983
Private Sub CommandButtonVyznaceniDolozky_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButtonVyznaceniDolozky.ForeColor = barvaZelenaAirBank
End Sub

Private Sub ButtonZmenitStav_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
ButtonZmenitStav.ForeColor = barvaZelenaAirBank
End Sub

Private Sub button_Vytvor_EPR_PDF_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
button_Vytvor_EPR_PDF.ForeColor = barvaZelenaAirBank
End Sub

