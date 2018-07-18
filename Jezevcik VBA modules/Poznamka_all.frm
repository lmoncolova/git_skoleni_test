VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Poznamka_all 
   Caption         =   "UserForm1"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6465
   OleObjectBlob   =   "Poznamka_all.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Poznamka_all"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

'  VSTUPY ktore treba nastavit PRED spustenim userform (vyuzit na to treba "public property LET" zadefinovany)
Private objVstupnyUserform As Object 'objekt bude userform Klient / Klient_dedictvi / Klient_...  'tady MUSI byt object, ne userform! (jinak objekt nebude mit metodu Show)

'  ostatne premenne
Private strRizeni As String 'jedna z hodnot "epr", "ded", "exe", "ins", "trest", "heir" ' RISK-468


Public Property Let VstupnyUserform(ByRef VstupnyUserform As Object) ' tady MUSI byt object, ne userform! (jinak objekt nebude mit metodu Show)
    
    Set objVstupnyUserform = VstupnyUserform
    
    Select Case objVstupnyUserform.name
        Case "Klient"
            strRizeni = "epr"
        Case "Klient_dedictvi"
            strRizeni = "ded"
        Case "Klient_exe"
            strRizeni = "exe"
        Case "Klient_insol"
            strRizeni = "ins"
        Case "Klient_trest"
            strRizeni = "trest"
        Case "Klient_heir"
            strRizeni = "heir"
        Case Else
            MsgBox "Nìco si nedomyslel. Jez-error 88558"
    End Select

End Property


Private Sub ButtonZrusit_Click()
    Call Me.Hide
End Sub


Private Sub UlozPoznamku()
Dim sql As String
Dim strSkpKey As String ' bud id_person alebo skp_heir
Dim strSkpValue As String ' konkretna hodnota pre id_person alebo skp_heir
Dim strSkpPerson As String ' promenna tak trochu navic, jde o to, ze globalni promenna id_person se nenastavuje pro heir, proto ji tady lokalni nahradime za strSkpPerson

If Len(Trim(TextBox1.text)) = 0 Then
    MsgBox "Musíš zadat text poznámky", vbExclamation, "Chybné zadání hodnot"
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


If Len(Trim(TextBox1.text)) > 0 Then
    If MsgBox("Opravdu chcete pøidat poznámku?", vbYesNo, "Pøidání poznámky") = vbYes Then
        
        sql = "insert into " & schema & "." & strRizeni & "_note n (N.SKP_" & strRizeni & "_NOTE, N." & strSkpKey & ", N.SKP_" & strRizeni & "_NOTE_TYPE, N.SKP_" & strRizeni & "_STATUS_DATA, N.TEXT_NOTE, N.DTIME_NOTE_CREATED, N.EMPLOYEE) " & _
                " values (" & max_skp("select max(skp_" & strRizeni & "_note) from " & schema & "." & strRizeni & "_note") & ", " & strSkpValue & _
                ", 3, NULL, '" & TextBox1.text & "', sysdate, '" & Main.username & "')"

        Call execute(sql)
        
    End If
End If

End Sub


Private Sub ButtonUlozit_Click()

If cs_HDWMN.State = 1 Then
prihlasen:

Call UlozPoznamku

Me.Hide
objVstupnyUserform.Show
If strRizeni = "epr" Then
    odkaz_Klient = True
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

Private Sub ButtonUlozitUpo_Click()

prihlasen:

Call UlozPoznamku
Me.Hide

If prava("select * from USER_TAB_PRIVS where owner = '" & schema & "' and table_name = 'NOTICE' and PRIVILEGE = 'INSERT'") Then
    Upozorneni_all.VstupnyUserform = objVstupnyUserform ' upozorneni se spusti z rovnakeho rizeni, z jakeho se spustilo vlozeni poznamky
    Upozorneni_all.VstupNoveUpo = True
    Upozorneni_all.Show
Else
    Call text_information_hide(Me.name)
    MsgBox "Nemáte práva pro pøidávání upozornìní"
    objVstupnyUserform.Show
End If



If err_connection Then
    err_connection = False
    Call re_login(Main.username.text, Main.password.text)
    If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If


End Sub

Private Sub TextBox1_Change()
    Label32.Caption = 500 - Len(TextBox1.text)
End Sub

' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()


If cs_HDWMN.State = 1 Then
prihlasen:

TextBox1.text = ""


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

If CloseMode = vbFormControlMenu Then
    Cancel = True
    Call Me.Hide
End If

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
ButtonZrusit.ForeColor = barvaCerna
ButtonUlozit.ForeColor = barvaCerna
ButtonUlozitUpo.ForeColor = barvaCerna
End Sub

Private Sub ButtonZrusit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
ButtonZrusit.ForeColor = barvaZelenaAirBank
End Sub

Private Sub ButtonUlozit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
ButtonUlozit.ForeColor = barvaZelenaAirBank
End Sub

Private Sub ButtonUlozitUpo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
ButtonUlozitUpo.ForeColor = barvaZelenaAirBank
End Sub




