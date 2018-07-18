VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Main 
   Caption         =   "Jezevèík v 3.0 - Hlavní menu"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8085
   OleObjectBlob   =   "Main.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit


Private Sub cmdClientInfo_Click()
' pokud je pole TextBoxClientInfo prazdne, fce nedela nic
' pokud tam je neco co se da konvertovat do cisla, hleda ve view ke klientum podle CUID, jinak podle jmena

Dim frmPoznBlok As Poznamkovy_blok

Dim strInput As String
Dim blnCUID As Boolean
Dim sql As String
Dim rstClient As ADODB.Recordset
Dim textSuhrn As String
Dim col As Integer
Dim strWarningInfo As String
Dim i As Integer

    strInput = Trim(Me.TextBoxClientInfo.text)
    
    
    '''   otvorenie prazdneho poznamkoveho bloku
    If Len(strInput) < 1 Then
        
        Set frmPoznBlok = New Poznamkovy_blok
        frmPoznBlok.Show
        Exit Sub
        
    End If
    
    
    
    If IsNumeric(strInput) Then blnCUID = True
    
    '''  nastaveni SQL dotazu
    sql = "SELECT DISTINCT c.name, c.cuid, c.id_person, c.SKP_PERSON, c.date_payoff, c.posl_stav_EPR, c.posl_stav_EXE, c.posl_stav_INS, c.posl_stav_TREST," & _
            " c.posl_stav_DED, c.posl_stav_Dedic" & _
            " FROM " & schema & ".v_client_status c WHERE c.flag_valid = 'Y'"
    
    If blnCUID Then
        sql = sql & " AND c.cuid = " & strInput
    Else
        sql = sql & " AND " & schema & ".IGNORUJ_DIAK(c.name) LIKE " & schema & ".IGNORUJ_DIAK('%' || '" & strInput & "' || '%')"
    End If
    
    
    '''  NAPLNENI DAT
    Set rstClient = New ADODB.Recordset
    rstClient.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
        '''   nalezeni prilis malo nebo mnoho klientu -> msgbox
        If rstClient.RecordCount < 1 Then
            MsgBox "Žádný klient v databázi Jezevèíka nebyl nenalezen.", vbInformation, "Klient nenalezen"
            Exit Sub
        End If
        
        '''   nalezen jediny klient
        
        ' FA - rad by som urobil UserForm ClientInfo, kde by boli iba rozpisane tieto informacie, ale zvladal by aj viac klientov (prepinanie nejakymi sipkami)
        ' zaroven by ten userform nijak nekomunikoval so zvyskom Jezevcika (nenastavovali by sa public premenne c pre klienta ani nic podobne
        ' vo vysledku vzdy pri kliknuti na lupu (t.j. tato sub = cmdClientInfo_Click), tak by sa otvorila nova instancia userform ClientInfo.
        ' ->> uzivatelia by ich tak mohli mat otvorenych viacero naraz. Boli by vbModeless, takze si ich len nechaju niekde na pozadi.
        ' chcem iba vyskust pracu s userForms as collection(), v zmysle viacerych instancii jedneho userForm
        '' zatial to iba vyhodi jednoduchy userForm
        
'        If rstClient.RecordCount > 1 Then
'            strWarningInfo = "Bylo nalezeno více klientù. To je jako kdybys hodil Jezevèíkovi více klackù najednou. On teï chudák vùbec neví, pro který z nich bìžet." & vbCrLf & vbCrLf & _
'                             "Poèet nalezených klientù: " & rstClient.RecordCount & ". Zobrazuji prvního z nich: " & vbCrLf & vbCrLf
'        End If
        
        Set frmPoznBlok = New Poznamkovy_blok
        frmPoznBlok.numberOfPages = rstClient.RecordCount
        
        For i = 1 To rstClient.RecordCount
            For col = 0 To (rstClient.Fields.count - 1)
                textSuhrn = textSuhrn & vbCrLf & _
                            rstClient.Fields(col).name & ":  " & rstClient(col)
                'Debug.Print textSuhrn
            Next col
            frmPoznBlok.VlozTextNaStranu i, strWarningInfo & textSuhrn
            textSuhrn = ""
            rstClient.MoveNext
        Next i
        
        frmPoznBlok.Show
        
        
    rstClient.Close
    
End Sub

' Login button
Public Sub CommandButton1_Click()
Dim control As control
Dim sql As String
Dim ok As Boolean


If Len(username.text) > 0 And Len(password.text) > 0 Then
  Call login(username.text, password.text)
  If cs_HDWMN.State = 1 Then
    ' Hide Login form and show buttons menu + LogOut botton
    For Each control In Main.Controls
         If control.Tag = "1" Then control.Visible = False
         If control.Tag = "2" Then control.Visible = True
    Next
    
    ' Check, if client has SELECT FROM privileges to specific table
    sql = "select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'PERSON' and PRIVILEGE = 'SELECT'"
    ok = prava(sql)
    
    CommandButton2.Enabled = ok
    CommandButton6.Enabled = ok
    CommandButton7.Enabled = ok
    CommandButton8.Enabled = ok
    CommandButton10.Enabled = ok
    CommandButton12.Enabled = ok
    CommandButton13.Enabled = ok
    Frame2.Enabled = ok
    Frame1.Enabled = ok
    CommandButton9.Enabled = ok
    CommandButton11.Enabled = ok
    
    Main.Frame1.BackColor = barvaZelenaAirBank
    
       'RISK-983
    
    Set JezUserLoggedIn = New JezUser
    
    Call JezUserLoggedIn.NactiUzivatele(username.text)
    
    Call UserForm_Activate
  End If
End If
End Sub


Private Sub CommandButton11_Click()
If cs_HDWMN.State = 1 Then
prihlasen:
 Main.Hide
' nastavení formuláøe probíhá pøi jeho aktivaci
 Nezesplatneni.Show

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

Private Sub CommandButton12_Click()
If cs_HDWMN.State = 1 Then
prihlasen:
 Main.Hide
' nastavení formuláøe probíhá pøi jeho aktivaci
 Insol.Show

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

Private Sub CommandButton13_Click()
If cs_HDWMN.State = 1 Then
prihlasen:
 Main.Hide
' nastavení formuláøe probíhá pøi jeho aktivaci
 Exekuce.Show

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

' Logout button
Private Sub CommandButton5_Click()
Call logout
End Sub



' Action after selection of button (Next UserForm)
Private Sub CommandButton2_Click()
If cs_HDWMN.State = 1 Then
prihlasen:
 Main.Hide
' nastavení formuláøe probíhá pøi jeho aktivaci
 Zesplatneni.Show

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



Private Sub CommandButton6_Click()
If cs_HDWMN.State = 1 Then
prihlasen:
 Main.Hide
' nastavení formuláøe probíhá pøi jeho aktivaci
 Zaloby.Show

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

Private Sub CommandButton7_Click()
If cs_HDWMN.State = 1 Then
prihlasen:
 Main.Hide
' nastavení formuláøe probíhá pøi jeho aktivaci
 Souhrn.Show

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

Private Sub CommandButton8_Click()
If cs_HDWMN.State = 1 Then
prihlasen:
 Main.Hide
' nastavení formuláøe probíhá pøi jeho aktivaci
 Stavy.Show

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


Private Sub CommandButton9_Click()
If cs_HDWMN.State = 1 Then
prihlasen:
 Main.Hide
' nastavení formuláøe probíhá pøi jeho aktivaci
 Dedictvi.Show

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


Private Sub CommandButton10_Click()
If cs_HDWMN.State = 1 Then
prihlasen:
 Main.Hide
' nastavení formuláøe probíhá pøi jeho aktivaci
 Trest.Show

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

'RISK-317
Private Sub CommandButton_Sonik_Click()

            If cs_HDWMN.State = 1 Then
prihlasen:
             Main.Hide
            ' nastavení formuláøe probíhá pøi jeho aktivaci
             
              Sonik_Datove_Zpravy_Prehled.Show ' RISK-985 odkomentovat po hromadnom REPLACE
            
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

Private Sub Frame2_Click()
Call CommandButton7_Click
End Sub

Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
Label34.ForeColor = barvaCerna
Label35.ForeColor = barvaCerna
Label36.ForeColor = barvaCerna
Label37.ForeColor = barvaCerna
Label38.ForeColor = barvaCerna
Label39.ForeColor = barvaCerna
Label_Sonik.ForeColor = barvaCerna
Label_Sonik_upo.ForeColor = barvaCerna
End Sub


Private Sub Image_Air_Test_Click()
    If MsgBox("Wanna party? Wanna run Full test?", vbYesNo, "Just do it!") = vbYes Then
        Call Jez_full_test
    End If
End Sub

Private Sub Label34_Click()
Call CommandButton7_Click
End Sub
Private Sub Label35_Click()
Call CommandButton7_Click
End Sub
Private Sub Label36_Click()
Call CommandButton7_Click
End Sub
Private Sub Label37_Click()
Call CommandButton7_Click
End Sub
Private Sub Label38_Click()
Call CommandButton7_Click
End Sub
Private Sub Label39_Click()
Call CommandButton7_Click
End Sub
Private Sub Label_Sonik_Click()
Call CommandButton7_Click
End Sub
Private Sub Label_Sonik_upo_Click()
Call CommandButton7_Click
End Sub


' SetFocus on OK button after press Enter
Private Sub password_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 9 Or KeyCode = 13 Then
    CommandButton1.SetFocus
End If
End Sub


Private Sub UserForm_Activate()
Dim sql As String

Main.Caption = "Jezevèík v." & version & " - Hlavní menu"
Main.Label3.Caption = "Verze " & version


If Not username.Visible Then
If cs_HDWMN.State = 1 Then
prihlasen:

 
'RISK-317
 Sonik_Datove_Zpravy_Prehled.Caption = "Soník - Pøehled datových zpráv" ' RISK-985 odkomentovat po hromadnom REPLACE
'Nadefinování uživatelù Soníka - nestaèí to ale tady, musí se grantovat i tabulky v DTB

Dim sql_sonik_user As String
Set rst1 = New ADODB.Recordset
'zjistime jestli má uživatel právo na soníka
   sql_sonik_user = "select USERNAME from " & schema & ".sonik_users where FLAG_VALID = 'Y' and USERNAME = '" & UCase(Main.username) & "'"
   rst1.Open sql_sonik_user, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
    Sonik_User = False ' RISK-985 odkomentovat po hromadnom REPLACE
   
   If Not rst1.EOF Then
     Sonik_User = True ' RISK-985 odkomentovat po hromadnom REPLACE
   End If
rst1.Close

''' ' RISK-985 odkomentovat po hromadnom REPLACE
If Sonik_User Then
    Main.CommandButton_Sonik.Enabled = True
Else
    Main.CommandButton_Sonik.Enabled = False
End If


'pocty na dolnu listu
If Sonik_User Then
    Call query("select pocet, po, auth, sonik from (select count(*) as pocet from " & schema & ".v_notice), (select count(*) as po from " & schema & ".v_notice where days_left < 0), (select count(*) as auth from " & schema & ".v_sop_auth), (SELECT COUNT(*) AS sonik FROM " & schema & ".SONIK_V_NOTICE)", Me.name)
Else
    Call query("select pocet, po, auth, sonik from (select count(*) as pocet from " & schema & ".v_notice), (select count(*) as po from " & schema & ".v_notice where days_left < 0), (select count(*) as auth from " & schema & ".v_sop_auth), (SELECT 'N/A' AS sonik FROM dual)", Me.name)
End If
''' ' RISK-985 odkomentovat po hromadnom REPLACE
 
'RISK-679 'Automatické generování složek - vždy pøi prvním denním spuštìní Evou Rajsiglovou
Call generate_client_folders


'RISK-434
Dim sqlda As String, date_load As Date
Set rst1 = New ADODB.Recordset
'zjistime jmeno klienta
sqlda = "select max(trunc(endtime)) as date_load_last from " & schema & ".chain_tracking where result = 'SUCCEDED'"
rst1.Open sqlda, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
    If Not rst1.EOF Then
        rst1.MoveFirst
        date_load = rst1![date_load_last]
    End If
rst1.Close
 
 If date_load = date Then
 Main.Label40.Visible = False
 Main.Label41.Visible = False
 Else
 Main.Label40.Visible = True
 Main.Label40.Caption = "V Jezevèíkovi dnes ještì nejsou èerstvá data"
 Main.Label41.Visible = True
 End If


 If err_connection Then
  err_connection = False
  Call re_login(Main.username.text, Main.password.text)
  If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
 End If
 
 
 If blnPrank Then
    PlayTheSound "Bloom.mp3" 'treba .wav subor
    PrankAlertTime = Now + TimeValue("00:00:02")
    Application.OnTime PrankAlertTime, "Discoo"
 End If

Else
 Call re_login(Main.username.text, Main.password.text)
 If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If

End If
End Sub





Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Image2.Visible = Not Image2.Visible
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 ThisWorkbook.Close SaveChanges:=False
End Sub




' Change color after MouseMove action **************************************************************************************************************************
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton2.ForeColor = barvaCerna
CommandButton6.ForeColor = barvaCerna
CommandButton7.ForeColor = barvaCerna
CommandButton8.ForeColor = barvaCerna
CommandButton9.ForeColor = barvaCerna
CommandButton10.ForeColor = barvaCerna
CommandButton11.ForeColor = barvaCerna
CommandButton12.ForeColor = barvaCerna
CommandButton13.ForeColor = barvaCerna
CommandButton_Sonik.ForeColor = barvaCerna
Label34.ForeColor = barvaBila
Label35.ForeColor = barvaBila
Label36.ForeColor = barvaBila
Label37.ForeColor = barvaBila
Label38.ForeColor = barvaBila
Label39.ForeColor = barvaBila
Label_Sonik.ForeColor = barvaBila
Label_Sonik_upo.ForeColor = barvaBila
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton1.ForeColor = barvaCerna
CommandButton5.ForeColor = barvaCerna
End Sub

Private Sub CommandButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton1.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton2.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton5.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton6.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton7.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton8.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton9.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    CommandButton10.ForeColor = barvaZelenaAirBank
    If blnPrank Then
        zamen_buttony
    End If
End Sub

Private Sub CommandButton11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton11.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    CommandButton12.ForeColor = barvaZelenaAirBank
    If blnPrank Then
        zamen_buttony
    End If
End Sub

Private Sub CommandButton13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton13.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Sonik_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Sonik.ForeColor = barvaZelenaAirBank
End Sub


Public Sub zamen_buttony()
    ' v ramci srandy, kdyz uzivatel najede kursorem na button, zavola sa tahle SUB ktera jej vymeni s jinym buttonem
    ' zatim zamenuje fixne jenom 2 konkretni buttons mezi sebou (10 a 12, t.j. trest a insolvence)
    
    Dim b1_left As Long
    Dim b1_top As Long
    
    
    If blnPrank Then
        
        ' ulozim pozici prvniho buttonu a presunu ho na druhej
        With CommandButton10
            b1_left = .Left
            b1_top = .Top
            
            .Left = CommandButton12.Left
            .Top = CommandButton12.Top
        End With
        
        ' druhej button presunu na ulozenou pozici prvniho
        With CommandButton12
            .Left = b1_left
            .Top = b1_top
        End With
        
        
    End If
End Sub

