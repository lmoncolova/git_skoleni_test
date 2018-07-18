VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Upozorneni_all 
   Caption         =   "UserForm1"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6720
   OleObjectBlob   =   "Upozorneni_all.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Upozorneni_all"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit
''''''   LOGIKA MODULU
'ciel je, aby sa modul dal nacitat zo vsetkych kariet a nemuseli existovat zvlast userforms pre kazde rizeni
'riesenie je mat premenne na urovni modulu, a k nim PUBLIC PROPERTY LET metodu ktorou sa da nastavit ich hodnota
'na karte Klient (/Klient_ded / ...) sa najprv nastavia premenne blnNoveUpozorneni a objVstupnyUserform (docasne aj strRizeni) a potom sa zavola userform.show



'  VSTUPY ktore treba nastavit PRED spustenim userform (vyuzit na to treba "public property LET" zadefinovany)
Private blnNoveUpozorneni As Boolean 'TRUE pro nove upozorneni, FALSE pro upravu existujiciho upo
Private objVstupnyUserform As Object 'objekt bude userform Klient / Klient_dedictvi / Klient_...  'tady MUSI byt object, ne userform! (jinak objekt nebude mit metodu Show)

'  ostatne premenne
Private strRizeni As String 'jedna z hodnot "epr", "ded", "exe", "ins", "trest", "heir" ' RISK-468
Private intRizeni As Integer 'pri vkladani upozorneni
Private strAssigneeDefault As String
Private lvUpozorneni As ListView


Public Property Let VstupnyUserform(ByRef VstupnyUserform As Object) ' tady MUSI byt object, ne userform! (jinak objekt nebude mit metodu Show)
    Set objVstupnyUserform = VstupnyUserform
End Property

Public Property Let VstupNoveUpo(ByVal NoveUpo As Boolean)
    blnNoveUpozorneni = NoveUpo
End Property


Private Sub NaplnPromennePodleRizeni()
'tahle Sub naplni promenne strRizeni a intRizeni, podle toho se pak vkladaji upo i poznamky

Select Case objVstupnyUserform.name '.Name?
    Case "Klient"
        strRizeni = "epr"
        intRizeni = 3
    Case "Klient_dedictvi"
        strRizeni = "ded"
        intRizeni = 1
    Case "Klient_exe"
        strRizeni = "exe"
        intRizeni = 2
    Case "Klient_insol"
        strRizeni = "ins"
        intRizeni = 4
    Case "Klient_trest"
        strRizeni = "trest"
        intRizeni = 5
    Case "Klient_heir"
        strRizeni = "heir"
        intRizeni = 6
    Case Else
        MsgBox "Nìco si nedomyslel. Jez-error 88559"
End Select

If strRizeni = "heir" Then
    'Debug.Print Klient_Heir.ListView_Heir_Notice.SelectedItem.SubItems(2)
    Set lvUpozorneni = Klient_heir.ListView_Heir_Notice
Else
    Set lvUpozorneni = objVstupnyUserform.ListView3
End If

End Sub



Private Sub CommandButton_Zrus_Click()
    Call Me.Hide
End Sub

Private Sub CommandButton_Uloz_Click()
Dim sql As String
Dim id As Integer
Dim strSkpKey As String ' bud id_person alebo skp_heir
Dim strSkpValue As String ' konkretna hodnota pre id_person alebo skp_heir
Dim strSkpPerson As String ' promenna tak trochu navic, jde o to, ze globalni promenna id_person se nenastavuje pro heir, proto ji tady lokalni nahradime za strSkpPerson

If Len(Trim(TextBox1.text)) = 0 Then
    MsgBox "Musíš zadat text upozornìní", vbExclamation, "Chybné zadání hodnot"
    Exit Sub
End If
Debug.Assert (strRizeni = "epr" Or strRizeni = "exe" Or strRizeni = "trest" Or strRizeni = "ded" Or strRizeni = "ins" Or strRizeni = "heir") 'kontrola hodnot, ci nie je niekde preklep


If strRizeni = "heir" Then
    Debug.Assert Not d Is Nothing ' d je public promenna - objekt Heir
    strSkpKey = "skp_heir"
    strSkpValue = d.lngSkp_heir
    strSkpPerson = d.objDedic.id_person
Else
    strSkpKey = "id_person"
    strSkpValue = id_person
    strSkpPerson = id_person
End If


' vlozit nove upozorneni

If blnNoveUpozorneni Then
    'If MsgBox("Opravdu chcete pøidat upozornìní?", vbYesNo, "Pøidání upozornìní") = vbYes Then
        
        Call TransBegin
            
            '   VLOZENI POZNAMKY NA KARTU KLIENTA
            
            sql = "insert into " & schema & "." & strRizeni & "_note t (t.skp_" & strRizeni & "_note, t." & strSkpKey & ", t.skp_" & strRizeni & "_note_type, t.skp_" & strRizeni & "_status_data, t.text_note, t.dtime_note_created, t.employee) " & _
                  " values (" & max_skp("select max(skp_" & strRizeni & "_note) from " & schema & "." & strRizeni & "_note") & ", " & strSkpValue & _
                  ", 4, NULL, 'Pøidáno upozornìní: " & TextBox1.text & " (deadline:  " & MonthView1.value & ")', sysdate, '" & Main.username & "')"
            
            Call execute(sql)
            sql = ""
            
            
            '   VLOZENI SAMOTNEHO UPOZORNENI
            
            'RISK-130
            'Novì je požadováno, aby se upozornìní otevíralo na zadané kartì.
            'Proto byla rozšíøena struktura tabulky notice o klíè SKP_PROCESS. Pomocí toho se pak pøi prokliku na souhrnu dostanu na kartu, kde bylo upozornìní zadané.
            sql = "insert into " & schema & ".notice (skp_notice, id_person, text_notice, dtime_notice_created, dtime_notice_deadline, code_status, employee, skp_process, assignee) " & _
                    " values (" & max_skp("select max(skp_notice) from " & schema & ".notice") & ", " & strSkpPerson & _
                    ",'" & TextBox1.text & "', sysdate, to_date('" & MonthView1.value & "','dd.mm.yyyy'), 'a', '" & Main.username & "', " & intRizeni & ", '" & ComboBoxAssignee.value & "')"
            'stav se zadava natvrdo jako aktivni
            
            Call execute(sql)
            sql = ""
            
        Call TransCommit
        
    'End If


' upravit existujici upozorneni

Else
    ' If MsgBox("Opravdu chcete upravit upozornìní?", vbYesNo, "Úprava upozornìní") = vbYes Then
        
        Call TransBegin
        
            '   VLOZENI POZNAMKY NA KARTU KLIENTA
            
            sql = ""
            If ComboBox1.value = "Storno/Smazat" Or ComboBox1.value = "Dokonèeno" Then
                
                sql = "insert into " & schema & "." & strRizeni & "_note t (t.skp_" & strRizeni & "_note, t." & strSkpKey & ", t.skp_" & strRizeni & "_note_type, t.skp_" & strRizeni & "_status_data, t.text_note, t.dtime_note_created, t.employee) " & _
                        " values (" & max_skp("select max(skp_" & strRizeni & "_note) from " & schema & "." & strRizeni & "_note") & ", " & strSkpValue & _
                        ", 4, NULL, '" & ComboBox1.value & ": " & TextBox1.text & " (deadline: " & MonthView1.value & ")', sysdate, '" & Main.username & "')"
            
            End If
            
            'If MonthView1.value <> lvUpozorneni.SelectedItem.SubItems(2) Then MsgBox "zmenil sa datum"
            'If TextBox1.text <> lvUpozorneni.SelectedItem.SubItems(1) Then MsgBox "zmenil sa textbox"
            
            If sql <> "" Then
                Call execute(sql)
                sql = ""
            End If
            
            
            '   VLOZENI SAMOTNEHO UPOZORNENI
            
            
            sql = "update " & schema & ".notice set " & _
                    " code_status = '" & LCase(Left(ComboBox1.value, 1)) & "'" & _
                    ", employee = '" & Main.username & "'" & _
                    ", assignee = '" & ComboBoxAssignee.value & "'" & _
                    ", text_notice = '" & TextBox1.text & "'" & _
                    ", dtime_notice_deadline = to_date('" & MonthView1.value & "','dd.mm.yyyy') " & _
                    " where skp_notice = " & lvUpozorneni.SelectedItem
            Call execute(sql)
            sql = ""
            
        Call TransCommit
        
    ' End If
End If


Me.Hide
objVstupnyUserform.Show
Set objVstupnyUserform = Nothing


End Sub




Private Sub TextBox1_Change()
    Label32.Caption = 400 - Len(TextBox1.text)
End Sub



Private Sub UserForm_Activate()
Dim sql As String
Dim intComboItem As Integer

If cs_HDWMN.State = 1 Then
prihlasen:


NaplnPromennePodleRizeni

' pri prvnim spusteni userformu se nactou hodnoty pro comboboxy (lidove nazyvane rolovatka)
If ComboBoxAssignee.ListCount < 1 Then
    Call NaplnAssigneeSeznam(Me.ComboBoxAssignee)
End If
If ComboBox1.ListCount < 1 Then
    NaplnCombobox1
End If

'pokud lze, nastavi se (pri prvnim spusteni) default assignee na toho kdo je prihlasen
If Len(Trim(strAssigneeDefault)) < 1 Then
    For intComboItem = 0 To Me.ComboBoxAssignee.ListCount - 1
        If UCase(Me.ComboBoxAssignee.List(intComboItem)) = UCase(Right(Main.username, Len(Main.username) - 4)) Or UCase(Me.ComboBoxAssignee.List(intComboItem)) = UCase(Main.username) Then  'alternativne: If InStr(UCase(username), UCase(Left(Application.username, InStr(Application.username, " ") - 1))) > 0 Then
            strAssigneeDefault = Me.ComboBoxAssignee.List(intComboItem)
        End If
        
    Next
End If


' nastavi se vetsina hodnot a poli na userformu podle toho, jestli vkladame nove nebo upravujeme existujici upozorneni

If blnNoveUpozorneni Then

    Label26.Caption = "Nové upozornìní"
    Me.Caption = "Nové upozornìní"
    TextBox1.text = ""
    MonthView1 = date
    ComboBox1.Enabled = False
    ComboBox1.ListIndex = 0 ' aktivni
    If Len(strAssigneeDefault) > 1 Then
        ComboBoxAssignee.value = strAssigneeDefault ' pri novom upozorneni se nastavi assignee na toho kdo je prihlasen (pokud lze)
    Else
        ComboBoxAssignee.value = ""
    End If
    
Else
    
    Label26.Caption = "Zmìna upozornìní"
    Me.Caption = "Zmìna upozornìní"
    TextBox1.text = lvUpozorneni.SelectedItem.SubItems(1)
    MonthView1 = lvUpozorneni.SelectedItem.SubItems(2)
    ComboBox1.Enabled = True
    ComboBox1.ListIndex = 0 ' aktivni
    ComboBoxAssignee.value = lvUpozorneni.SelectedItem.SubItems(4)
    
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


Private Sub NaplnCombobox1()
    ComboBox1.Clear
    ComboBox1.AddItem "Aktivní"
    ComboBox1.AddItem "Storno/Smazat"
    ComboBox1.AddItem "Dokonèeno"
    ComboBox1.ListIndex = 0
End Sub




Private Sub UserForm_QueryClose(Cancel As Integer, _
                            CloseMode As Integer)
Dim ans
If CloseMode = vbFormControlMenu Then
    Cancel = True
    Call Me.Hide
End If
End Sub


' Zmìna barvy po najetí myši **************************************************************************************************************************
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




