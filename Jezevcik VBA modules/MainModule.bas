Attribute VB_Name = "MainModule"
Option Explicit

'''   Nastavenie PUBLIC premenn˝ch, ktorÈ sa menia medzi verziami jezevËÌka a TEST vs. PROD
'''   PromÏnn·, kter· definuje h¯iötÏ - zda se jedn· o testovaËku nebo produkci
Public Const schema = "APP_COLL_JEZEVCIK_TEST" ' je potreba psat velke 'APP_COLL_JEZEVCIK nebo APP_COLL_JEZEVCIK_TEST
Public Const version = "5.9 - TEST"
Public Const DWHserver = "PROD" 'PROD nebo PREPROD


Public Const blnPrank = False      'TRUE pro prvni april, narozeniny jezevcika, narozeniny oddeleni pravniho vymahani a pod..
Public Const intOmezitPoctyKlientu As Integer = 50 ' omezeni poctu radku ktere se nacitaji na hlavnich kartach rizeni - kvuli zrychleni jezevcika. Zavedeno ~ 13.11.2017
Public strMsgbox As String

Public export_chosen_option As Integer 'premenna bude mat hodnoty 1-4, pre export pdf/xml, aktualny/vsetky uvery.  'FA_export
Public wb_export_plateb As Object
Public JezUserLoggedIn As JezUser 'RISK-983
Public KlientiSeZapornymKU As Object ' vyuziva se v Client.initialize - bug 16.2.2018
Public err_connection As Boolean  ' PromÏnn· pro evidenci, zda jde o chybu z d˘vodu chybnÈho p¯ipojenÌ

Public c As Client                ' Ustredni promenna pro nacteneho klienta
Public dedici As Heirs            'RISK-468
Public d As Heir                  ' konkretni dedic (d je instance objektu Heir, podobne jak c je instance objektu Client) 'RISK-468
Public odkaz_Klient As Boolean    ' Promenna pro odkazovani sa z zalobnej karty  klienta na inu kartu klienta
Public skp_proces As String
Public id_person As String        ' PromÏnn· pro uchov·nÌ informace, o kterÈho klienta se jedn·
Public userform_prv As String     ' PromÏnn·, kter· uchov·v· informaci, kter˝ formul·¯ byl otev¯en p¯edtÌm
Public formular As String         ' Variable for evidence of opened/opening UserForm
Public cfg As Config_class

' Variable for Connection to Database and managing tables
Public cs_HDWMN As ADODB.Connection
Private dbConnectStr As String
Public blnTransactionStatus As Boolean ' True, pokud je otevrena transakce


' Typ pro zapamatov·nÌ filtru na formul·¯ich Insolvence, Trest, Dedictvi, Zaloby, Nezesplatneni, Exekuce, Zesplatneni
Type filtr
    set As Boolean
    CUID As String
    jmeno As String
    rc As String
    smlouva As String
    date As Boolean
    od As Date
    do As Date
    duvod As String
    code As String
    stav As String
    spis As String
    'code_ex As String
    'proces As String '
End Type


' Nadeklarov·nÌ promÏnnÈ typu filtru
Public zes_filtr As filtr
Public zal_filtr As filtr
Public exe_filtr As filtr
Public ded_filtr As filtr
Public trest_filtr As filtr
Public ins_filtr As filtr
Public nezes_filtr As filtr



Public i As Integer
Public j As Integer
Public rst1 As ADODB.Recordset
Public rst2 As ADODB.Recordset
Public rst3 As ADODB.Recordset
Public rst4 As ADODB.Recordset
Public rst5 As ADODB.Recordset
Public rst6 As ADODB.Recordset
Public rst7 As ADODB.Recordset
Public rst8 As ADODB.Recordset
'Private vbWhi As Variant 'LOL
'Private objSelection As Object
'Private fso As Variant



Public Type SOUD_type
    skp_soud As Long
    okresni_soud As String
    krajsky_soud As String
    adresa As String
    adresa_P1 As String
    adresa_P2 As String ' StrSoudAdresaUlice
    adresa_P3 As String ' StrSoudAdresaPscMesto
    ID_soud As String   ' StrSoudIDSchranky
    mail As String
    tel As String
End Type





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''     Connection subs     '''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




Sub login(username As String, password As String)
Dim strServerAddress As String
On Error GoTo chyba
  Set cs_HDWMN = Nothing
  Set cs_HDWMN = New ADODB.Connection
  
  If DWHserver = "PROD" Then
        strServerAddress = "HDWMN.BANKA.HCI"
  ElseIf DWHserver = "PREPROD" Then
        strServerAddress = "MNPPDW.BANKA.HCI" 'MNPPDW.BANKA.HCI
  Else
        MsgBox "tak prod nebo preprod??"
        Exit Sub
  End If

  dbConnectStr = "Provider = MSDAORA.1; Data source = " & strServerAddress & "; User ID = '" + username + "'; Password = '" + password + "';"
  cs_HDWMN.ConnectionString = dbConnectStr
  cs_HDWMN.Mode = adModeReadWrite
  cs_HDWMN.CursorLocation = adUseClient
  cs_HDWMN.CommandTimeout = 300
  cs_HDWMN.ConnectionTimeout = 20
  cs_HDWMN.Open

Exit Sub
chyba:
   
    MsgBox "Chyba: " & Err.Description
End Sub


Sub re_login(username As String, password As String)
Dim strServerAddress As String
On Error GoTo chyba
  cs_HDWMN.Close
chyba:

  Set cs_HDWMN = Nothing
  Set cs_HDWMN = New ADODB.Connection
  
  If DWHserver = "PROD" Then
        strServerAddress = "HDWMN.BANKA.HCI"
  ElseIf DWHserver = "PREPROD" Then
        strServerAddress = "MNPPDW.BANKA.HCI"
  Else
        MsgBox "tak prod nebo preprod??"
        Exit Sub
  End If

  dbConnectStr = "Provider = MSDAORA.1; Data source = " & strServerAddress & "; User ID = '" + username + "'; Password = '" + password + "';"
  
  cs_HDWMN.ConnectionString = dbConnectStr
  cs_HDWMN.Mode = adModeReadWrite
  cs_HDWMN.CursorLocation = adUseClient
  cs_HDWMN.CommandTimeout = 300
  cs_HDWMN.ConnectionTimeout = 20
  cs_HDWMN.Open

End Sub


Public Sub logout()
Dim control As control

'Form.Hide
Main.Show
Main.Frame1.BackColor = barvaCerna
   For Each control In Main.Controls
    If control.Tag = "1" Then control.Visible = True
    If control.Tag = "2" Then control.Visible = False
   Next

Main.Label27.Caption = "Byl jste odhl·öen." & Chr(10) & "Pro pr·ci s touto aplikacÌ je nutnÈ se p¯Ìhl·sit"
'RISK-434
Main.Label40.Visible = False
Main.Label41.Visible = False

'On Error GoTo continue
cs_HDWMN.Close

'continue:

End Sub






''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''       SQL  subs       '''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''





' Execution of script on database
Sub execute(dotaz As String)

    On Error GoTo connection_problem
    
    cs_HDWMN.execute dotaz
    Exit Sub
    
    
connection_problem:
    
    ''  Rollback (funkce overi jestli je otevrena transakce)
    Call TransRollback
    
    ' V p¯ÌpadÏ, ûe se jedn· o chybu z d˘vodu nep¯ipojenÌ k datab·zi, bude hodnota promÏnnÈ err_connection true a poslÈze bude proveden reconnect/relogin
    If Mid(Err.Description, 1, 9) = "ORA-03114" Then
        err_connection = True
    Else
        'MsgBox "JezevËÌk upadl a zav¯e se, poöli prosÌm printscreen oblÌbenÈmu analytikovi. " & vbNewLine & Err.Description, vbCritical, "Kritick· chyba"
        MsgBox "Doölo k chybÏ p¯i nahr·v·nÌ zmÏn do datab·ze. JezevËÌk se zav¯e. Zkus prosÌm znovu vykonat poslednÌ akci, a pokud by se chyba zopakovala tak prosÌm poöli printscreen oblÌbenÈmu analytikovi. " & vbNewLine & Err.Description, vbCritical, "Kritick· chyba"
        Debug.Print dotaz
        Debug.Assert (1 = 0) 'pokud ma uzivatel otevreny VBA editor tak se to tady zasekne
        End 'End nenÌ end IFu ale ukonËÌ macro - break Ëtverec naho¯e v liötÏ
    End If

End Sub





' Find next ID (SKP_number) for define table
Function max_skp(dotaz As String) As Long
Dim rst_skp As ADODB.Recordset
On Error GoTo connection_problem

Set rst_skp = New ADODB.Recordset
rst_skp.Open dotaz, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

    ' Najde n·sledujÌcÌ id/skp v tabulce (pokud bude hodnota null, zaËÌn· od 1)
    If Not IsNull(rst_skp(0)) Then
     max_skp = rst_skp(0) + 1
    Else
     max_skp = 1
    End If

rst_skp.Close

Exit Function
connection_problem:

' V p¯ÌpadÏ, ûe se jedn· o chybu z d˘vodu nep¯ipojenÌ k datab·zi, bude hodnota promÏnnÈ err_connection true a poslÈze bude proveden reconnect/relogin
If Mid(Err.Description, 1, 9) = "ORA-03114" Then
err_connection = True
Else
 MsgBox Err.Description
End If


End Function



Function prava(dotaz As String) As Boolean
' Check, if user have privileges to SELECT FROM/INSERT INTO table
' Kontrola toho, jestli dotaz vracÌ nÏjakÈ hodnoty (pouûitÌ p¯edevöÌm pro kontrolu pr·v)

On Error GoTo connection_problem
 Set rst1 = New ADODB.Recordset
 rst1.Open dotaz, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
 
 If Not rst1.EOF Then
  prava = True
 Else
  prava = False
 End If
 rst1.Close
 
Exit Function

connection_problem:

If Mid(Err.Description, 1, 9) = "ORA-03114" Then
err_connection = True
Else
 MsgBox Err.Description
End If

 
End Function



Public Sub Del_stav(cil_tab As String) '" & cil_tab & "
'procedura na smazani stavu, pridane v ramci RISK-35
' FA 5.1.2018 upravovano aby fungovalo i pro dedice (RISK-468)
' Sub vyuziva globalni id_person

Dim sql As String
Dim name_epr_stdata As String
Dim dtValidFrom As Date
Dim arraySel() As String
Dim strSkpStatusData As String
Dim StrNameStatus As String
Dim strSkpKey As String ' bud id_person alebo skp_heir
Dim strSkpValue As String ' konkretna hodnota pre id_person alebo skp_heir
Dim strSkpPerson As String ' promenna tak trochu navic, jde o to, ze globalni promenna id_person se nenastavuje pro heir, proto ji tady lokalni nahradime za strSkpPerson

    
    dtValidFrom = Empty
    name_epr_stdata = Empty
    
    '''   vybr·nÌ poslednÌho stavu ke smaz·nÌ pomocÌ fce Posl_stav
    arraySel = posl_stav(cil_tab)
    If arraySel(0) = "" Then
        MsgBox ("Neexistuje stav ke smaz·nÌ.")
        Exit Sub
    End If
    strSkpStatusData = arraySel(0)
    dtValidFrom = CDate(arraySel(1))
    StrNameStatus = arraySel(2)
    
        
    '''   nastaveni SKP klicu aby fungovali pro PERSON i HEIR (strSkpValue uz bolo zjednotene vo funkcii posl_stav, ktora sa volala o par riadkov vyssie)
    If cil_tab = "heir" Then
        Debug.Assert Not d Is Nothing ' d je public promenna - objekt Heir
        If d Is Nothing Then
            MsgBox "jezevcik error 008705", vbCritical
            End
        End If
        strSkpKey = "skp_heir"
        strSkpValue = d.lngSkp_heir
        strSkpPerson = d.objDedic.id_person
    Else
        strSkpKey = "id_person"
        strSkpValue = id_person
        strSkpPerson = id_person
    End If
    

    If MsgBox("Opravdu chcete smazat poslednÌ stav " & StrNameStatus & " na klientovi platn˝ od " & dtValidFrom & "? ", vbExclamation + vbYesNo, "Smaz·nÌ stavu") = vbYes Then
        sql = "update " & schema & "." & cil_tab & "_STATUS_DATA set FLAG_DELETED='Y', dtime_valid_to = sysdate, employee = '" & Main.username & "' where SKP_" & cil_tab & "_STATUS_DATA = " & strSkpStatusData
        execute (sql)
        'uprava dtime_valid_to na poslednÌm platnÈm ¯·dku
        sql = "update " & schema & "." & cil_tab & "_STATUS_DATA set dtime_valid_to = TO_DATE('01.01.3000', 'dd.mm.yyyy') where DTIME_VALID_TO = to_date('" & dtValidFrom & "','dd.mm.yyyy hh24:mi:ss')-1/86400 " & _
                    "AND " & strSkpKey & " = " & strSkpValue
        execute (sql)
        'vloûenÌ pozn·mky do _note
        sql = "insert into " & schema & "." & cil_tab & "_note (skp_" & cil_tab & "_note, " & strSkpKey & ", skp_" & cil_tab & "_note_type, skp_" & cil_tab & "_status_data, text_note, dtime_note_created, employee)" & _
                    " VALUES (" & max_skp("select max(skp_" & cil_tab & "_note) from " & schema & "." & cil_tab & "_note") & ", " & strSkpValue & _
                    ", 1, " & strSkpStatusData & ", 'Smaz·n poslednÌ stav " & StrNameStatus & " na klientovi platn˝ od " & dtValidFrom & "', sysdate, '" & Main.username & "')"
        execute (sql)
        
        ' EXTRA podminky pro nektere stavy
        If StrNameStatus = "Vym·h·nÌ ukonËeno" Then
            MsgBox "Klientovi je pot¯eba vymazat pozn·mku Vym·h·nÌ UkonËeno ze vöech ¯ÌzenÌ. Hezky popros Filipa - krotitele JezevËÌka (p¯ÌpadnÏ jinÈho analytika)." ' RISK-468
            ' funkce rozpracovana na userforme update_stav_all, vola sa VymazPoznamkuVymahaniUkonceno
        End If
        
        If StrNameStatus = "P¯edexekuËnÌ v˝zva" And cil_tab = "epr" Then
            'MsgBox ("Je t¯eba manu·lnÏ vymazat tento stav i z tabulky EXE - Formul·¯ Exekuce!!!")
            Del_stav ("exe")
        End If
    End If


End Sub

Private Function posl_stav(cil_tab As String) As String()
' funkce je vyuzivana v ramci mazani stavu - viz Sub Del_stav
' FA 5.1.2018 upravovano aby fungovalo i pro dedice (RISK-468)
' Sub vyuziva globalni id_person

Dim arraySel(0 To 2) As String
Dim sql As String

Dim strSkpKey As String ' bud id_person alebo skp_heir
Dim strSkpValue As String ' konkretna hodnota pre id_person alebo skp_heir
Dim strSkpPerson As String ' promenna tak trochu navic, jde o to, ze globalni promenna id_person se nenastavuje pro heir, proto ji tady lokalni nahradime za strSkpPerson

'''   nastaveni SKP klicu a hodnot aby fungovali pro PERSON i HEIR
If cil_tab = "heir" Then
    Debug.Assert Not d Is Nothing ' d je public promenna - objekt Heir
    If d Is Nothing Then
        MsgBox "jezevcik error 008704", vbCritical
        End
    End If
    strSkpKey = "skp_heir"
    strSkpValue = d.lngSkp_heir
    strSkpPerson = d.objDedic.id_person
Else
    strSkpKey = "id_person"
    strSkpValue = id_person
    strSkpPerson = id_person
End If



    'vyberu poslednÌ stav na klientovi
    sql = "select a.SKP_" & cil_tab & "_STATUS_DATA, a.dtime_valid_from, b.name_" & cil_tab & "_status from " & schema & "." & cil_tab & "_STATUS_DATA a left join  " _
                & schema & "." & cil_tab & "_STATUS b on  a.skp_" & cil_tab & "_status =  b.skp_" & cil_tab & "_status" & _
                " WHERE FLAG_DELETED='N' AND " & strSkpKey & " = " & strSkpValue & " AND a.DTIME_VALID_TO = TO_DATE('01.01.3000', 'dd.mm.yyyy')"
                '" WHERE FLAG_DELETED='N' AND id_person = " & id_person & " and a.DTIME_VALID_TO = TO_DATE('01.01.3000', 'dd.mm.yyyy')"
    
    Set rst1 = New ADODB.Recordset
    rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
        If Not rst1.EOF Then
            For i = 0 To 2
                arraySel(i) = rst1(i)
            Next
        End If
        
    rst1.Close
    posl_stav = arraySel
    
End Function













''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''   Subs pre naËÌtanie UserForms    '''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






' Prepare ListView component to insert data
' P¯Ìprava tabulek (vÏtöinou na ud·lost UserForm_Initialize): Vytvo¯enÌ sloupc˘ v listview/DoplnÏnÌ fixnÌch hodnot do Comboboxu
Sub table_prepare(table As String)

' Variable for ListView component
Dim clmx As ColumnHeader
Dim itmX As ListItem
Dim lvSI As ListSubItem


Select Case table

Case "Zaloby"

    Zaloby.ListView1.ColumnHeaders.Clear
    With Zaloby.ListView1.ColumnHeaders
    
        Call .Add(, , "ID", 50)
        Call .Add(, , "ID_PERSON", 65)
        Call .Add(, , "CUID", 80)
        Call .Add(, , "Klient", 150)
        Call .Add(, , "RodnÈ ËÌslo", 80)
        Call .Add(, , "Status", 208)
        Call .Add(, , "Proces", 100)
        Call .Add(, , "Spisov· znaËka", 100)
        Call .Add(, , "Platn˝ od", 100)
        Call .Add(, , "Platn˝ do", 100)
        Call .Add(, , "Neû·doucnost", 65)
        
    End With
 
 
Case "Exekuce"
 Exekuce.ListView1.ColumnHeaders.Clear
 Set clmx = Exekuce.ListView1.ColumnHeaders.Add(, , "ID", 50)
 Set clmx = Exekuce.ListView1.ColumnHeaders.Add(, , "ID_PERSON", 65)
 Set clmx = Exekuce.ListView1.ColumnHeaders.Add(, , "CUID", 80)
 Set clmx = Exekuce.ListView1.ColumnHeaders.Add(, , "Klient", 150)
 Set clmx = Exekuce.ListView1.ColumnHeaders.Add(, , "RodnÈ ËÌslo", 80)
 Set clmx = Exekuce.ListView1.ColumnHeaders.Add(, , "Status", 208)
 Set clmx = Exekuce.ListView1.ColumnHeaders.Add(, , "Proces", 100)
 Set clmx = Exekuce.ListView1.ColumnHeaders.Add(, , "ExekuËnÌ znaËka", 100)
 Set clmx = Exekuce.ListView1.ColumnHeaders.Add(, , "Platn˝ od", 100)
 Set clmx = Exekuce.ListView1.ColumnHeaders.Add(, , "Platn˝ do", 100)
  Set clmx = Exekuce.ListView1.ColumnHeaders.Add(, , "Neû·doucnost", 65)
 
 
 
'''   zakomentovano RISK-468
'''Case "Dedictvi"
''' Dedictvi.ListView1.ColumnHeaders.Clear
''' Set clmx = Dedictvi.ListView1.ColumnHeaders.Add(, , "ID", 50)
''' Set clmx = Dedictvi.ListView1.ColumnHeaders.Add(, , "ID_PERSON", 65)
''' Set clmx = Dedictvi.ListView1.ColumnHeaders.Add(, , "CUID", 80)
''' Set clmx = Dedictvi.ListView1.ColumnHeaders.Add(, , "Klient", 150)
''' Set clmx = Dedictvi.ListView1.ColumnHeaders.Add(, , "RodnÈ ËÌslo", 80)
''' Set clmx = Dedictvi.ListView1.ColumnHeaders.Add(, , "Status", 208)
''' Set clmx = Dedictvi.ListView1.ColumnHeaders.Add(, , "Proces", 100)
''' Set clmx = Dedictvi.ListView1.ColumnHeaders.Add(, , "Spisov· znaËka", 100)
''' Set clmx = Dedictvi.ListView1.ColumnHeaders.Add(, , "Platn˝ od", 100)
''' Set clmx = Dedictvi.ListView1.ColumnHeaders.Add(, , "Platn˝ do", 100)
''' Set clmx = Dedictvi.ListView1.ColumnHeaders.Add(, , "Neû·doucnost", 65)
 
Case "Trest"
 Trest.ListView1.ColumnHeaders.Clear
 Set clmx = Trest.ListView1.ColumnHeaders.Add(, , "ID", 50)
 Set clmx = Trest.ListView1.ColumnHeaders.Add(, , "ID_PERSON", 65)
 Set clmx = Trest.ListView1.ColumnHeaders.Add(, , "CUID", 80)
 Set clmx = Trest.ListView1.ColumnHeaders.Add(, , "Klient", 150)
 Set clmx = Trest.ListView1.ColumnHeaders.Add(, , "RodnÈ ËÌslo", 80)
 Set clmx = Trest.ListView1.ColumnHeaders.Add(, , "Status", 208)
 Set clmx = Trest.ListView1.ColumnHeaders.Add(, , "Proces", 100)
 Set clmx = Trest.ListView1.ColumnHeaders.Add(, , "Spisov· znaËka", 100)
 Set clmx = Trest.ListView1.ColumnHeaders.Add(, , "Platn˝ od", 100)
 Set clmx = Trest.ListView1.ColumnHeaders.Add(, , "Platn˝ do", 100)
 Set clmx = Trest.ListView1.ColumnHeaders.Add(, , "Neû·doucnost", 65)
 
Case "Insol"
 Insol.ListView1.ColumnHeaders.Clear
 Set clmx = Insol.ListView1.ColumnHeaders.Add(, , "ID", 50)
 Set clmx = Insol.ListView1.ColumnHeaders.Add(, , "ID_PERSON", 65)
 Set clmx = Insol.ListView1.ColumnHeaders.Add(, , "CUID", 60)
 Set clmx = Insol.ListView1.ColumnHeaders.Add(, , "Klient", 150)
 Set clmx = Insol.ListView1.ColumnHeaders.Add(, , "RodnÈ ËÌslo", 80)
 Set clmx = Insol.ListView1.ColumnHeaders.Add(, , "Vztah k bance", 70)
 Set clmx = Insol.ListView1.ColumnHeaders.Add(, , "Aktu·lnÌ status", 208)
 Set clmx = Insol.ListView1.ColumnHeaders.Add(, , "Spisov· znaËka", 100)
 Set clmx = Insol.ListView1.ColumnHeaders.Add(, , "Stav ¯ÌzenÌ", 90)
 Set clmx = Insol.ListView1.ColumnHeaders.Add(, , "ZmÏna stavu ¯ÌzenÌ", 110)
 Set clmx = Insol.ListView1.ColumnHeaders.Add(, , "ZesplatnÏnÌ", 55)
 Set clmx = Insol.ListView1.ColumnHeaders.Add(, , "Deadline", 60)
 
Case "Insol_nesparovane"
 Insol_nesparovane.ListView1.ColumnHeaders.Clear
 Set clmx = Insol_nesparovane.ListView1.ColumnHeaders.Add(, , "ID_PERSON", 65)
 Set clmx = Insol_nesparovane.ListView1.ColumnHeaders.Add(, , "CUID", 60)
 Set clmx = Insol_nesparovane.ListView1.ColumnHeaders.Add(, , "JmÈno", 60)
 Set clmx = Insol_nesparovane.ListView1.ColumnHeaders.Add(, , "RodnÈ ËÌslo", 100)
 Set clmx = Insol_nesparovane.ListView1.ColumnHeaders.Add(, , "Stav ¯ÌzenÌ", 100)
 Set clmx = Insol_nesparovane.ListView1.ColumnHeaders.Add(, , "D·tum zaloûenÌ spisu", 110)
 Set clmx = Insol_nesparovane.ListView1.ColumnHeaders.Add(, , "Spisov· znaËka", 110)
 Set clmx = Insol_nesparovane.ListView1.ColumnHeaders.Add(, , "D·tum zmeny stavu", 110)
 Set clmx = Insol_nesparovane.ListView1.ColumnHeaders.Add(, , "Vzùah", 100)
 Set clmx = Insol_nesparovane.ListView1.ColumnHeaders.Add(, , "ZesplatnÏnÌ", 100)
 Set clmx = Insol_nesparovane.ListView1.ColumnHeaders.Add(, , "Deadline", 80)


Case "Zesplatneni"
 Zesplatneni.ListView1.ColumnHeaders.Clear
 Set clmx = Zesplatneni.ListView1.ColumnHeaders.Add(, , "ID_PERSON", 65)
 Set clmx = Zesplatneni.ListView1.ColumnHeaders.Add(, , "CUID", 55)
 Set clmx = Zesplatneni.ListView1.ColumnHeaders.Add(, , "Klient", 125)
 Set clmx = Zesplatneni.ListView1.ColumnHeaders.Add(, , "RodnÈ ËÌslo", 60)
 Set clmx = Zesplatneni.ListView1.ColumnHeaders.Add(, , "Smlouvy klienta", 80)
 Set clmx = Zesplatneni.ListView1.ColumnHeaders.Add(, , "Datum zesplatnÏnÌ", 90)
 Set clmx = Zesplatneni.ListView1.ColumnHeaders.Add(, , "D˘vod zesplacenÌ", 100)
 Set clmx = Zesplatneni.ListView1.ColumnHeaders.Add(, , "KÛd insolvence", 75)
 Set clmx = Zesplatneni.ListView1.ColumnHeaders.Add(, , "Neû·doucnost", 65)
 Set clmx = Zesplatneni.ListView1.ColumnHeaders.Add(, , "Evidov·no", 65)
 
Case "Nezesplatneni"
 Nezesplatneni.ListView1.ColumnHeaders.Clear
 Set clmx = Nezesplatneni.ListView1.ColumnHeaders.Add(, , "ID_PERSON", 65)
 Set clmx = Nezesplatneni.ListView1.ColumnHeaders.Add(, , "CUID", 55)
 Set clmx = Nezesplatneni.ListView1.ColumnHeaders.Add(, , "Klient", 125)
 Set clmx = Nezesplatneni.ListView1.ColumnHeaders.Add(, , "RodnÈ ËÌslo", 60)
 Set clmx = Nezesplatneni.ListView1.ColumnHeaders.Add(, , "Smlouvy klienta", 80)
 Set clmx = Nezesplatneni.ListView1.ColumnHeaders.Add(, , "Datum zesplatnÏnÌ", 90)
 Set clmx = Nezesplatneni.ListView1.ColumnHeaders.Add(, , "D˘vod zesplacenÌ", 100)
 Set clmx = Nezesplatneni.ListView1.ColumnHeaders.Add(, , "KÛd insolvence", 75)
 Set clmx = Nezesplatneni.ListView1.ColumnHeaders.Add(, , "Neû·doucnost", 65)
 Set clmx = Nezesplatneni.ListView1.ColumnHeaders.Add(, , "Evidov·no", 65)

 
Case "ZesplatneniCombo"
    Zesplatneni.ComboBox1.Clear
    Zesplatneni.ComboBox2.Clear
    
    Zesplatneni.ComboBox1.AddItem "Vöe"
    Zesplatneni.ComboBox1.AddItem "AML"
    Zesplatneni.ComboBox1.AddItem "Delikvence"
    Zesplatneni.ComboBox1.AddItem "Insolvence"
    Zesplatneni.ComboBox1.AddItem "TrestnÌ ¯ÌzenÌ/podvod"
    Zesplatneni.ComboBox1.AddItem "⁄mrtÌ"
    Zesplatneni.ComboBox1.AddItem "XNA"
    
    Zesplatneni.ComboBox2.AddItem "Vöe"
    Zesplatneni.ComboBox2.AddItem "é·dn˝"
    Zesplatneni.ComboBox2.AddItem "K-PO ZRUä."
    Zesplatneni.ComboBox2.AddItem "KONKURS"
    Zesplatneni.ComboBox2.AddItem "MARATORIUM"
    Zesplatneni.ComboBox2.AddItem "MYSLN› Z¡P."
    Zesplatneni.ComboBox2.AddItem "NEVYR-POST"
    Zesplatneni.ComboBox2.AddItem "NEVYRIZENA"
    Zesplatneni.ComboBox2.AddItem "OBZIVLA"
    Zesplatneni.ComboBox2.AddItem "ODDLUéENÕ"
    Zesplatneni.ComboBox2.AddItem "ODSKRTNUTA"
    Zesplatneni.ComboBox2.AddItem "PRAVOMOCNA"
    Zesplatneni.ComboBox2.AddItem "REORGANIZ"
    Zesplatneni.ComboBox2.AddItem "⁄PADEK"
    Zesplatneni.ComboBox2.AddItem "VYRIZENA"
    Zesplatneni.ComboBox2.AddItem "ZRUäENO VS"
    Zesplatneni.ComboBox2.AddItem "XAP"
    Zesplatneni.ComboBox2.AddItem "XNA"
 
 Case "NezesplatneniCombo"
    Nezesplatneni.ComboBox1.Clear
    Nezesplatneni.ComboBox2.Clear
    
    Nezesplatneni.ComboBox1.AddItem "Vöe"
    Nezesplatneni.ComboBox1.AddItem "AML"
    Nezesplatneni.ComboBox1.AddItem "Delikvence"
    Nezesplatneni.ComboBox1.AddItem "Insolvence"
    Nezesplatneni.ComboBox1.AddItem "TrestnÌ ¯ÌzenÌ/podvod"
    Nezesplatneni.ComboBox1.AddItem "⁄mrtÌ"
    Nezesplatneni.ComboBox1.AddItem "XNA"
    
    Nezesplatneni.ComboBox2.AddItem "Vöe"
    Nezesplatneni.ComboBox2.AddItem "é·dn˝"
    Nezesplatneni.ComboBox2.AddItem "K-PO ZRUä."
    Nezesplatneni.ComboBox2.AddItem "KONKURS"
    Nezesplatneni.ComboBox2.AddItem "MARATORIUM"
    Nezesplatneni.ComboBox2.AddItem "MYSLN› Z¡P."
    Nezesplatneni.ComboBox2.AddItem "NEVYR-POST"
    Nezesplatneni.ComboBox2.AddItem "NEVYRIZENA"
    Nezesplatneni.ComboBox2.AddItem "OBZIVLA"
    Nezesplatneni.ComboBox2.AddItem "ODDLUéENÕ"
    Nezesplatneni.ComboBox2.AddItem "ODSKRTNUTA"
    Nezesplatneni.ComboBox2.AddItem "PRAVOMOCNA"
    Nezesplatneni.ComboBox2.AddItem "REORGANIZ"
    Nezesplatneni.ComboBox2.AddItem "⁄PADEK"
    Nezesplatneni.ComboBox2.AddItem "VYRIZENA"
    Nezesplatneni.ComboBox2.AddItem "ZRUäENO VS"
    Nezesplatneni.ComboBox2.AddItem "XAP"
    Nezesplatneni.ComboBox2.AddItem "XNA"
    

Case "Klient"
 Klient.TextBox1.text = ""
 Set clmx = Klient.ListView1.ColumnHeaders.Add(, , "ID", 30)
 Set clmx = Klient.ListView1.ColumnHeaders.Add(, , "Datum ud·losti", 90)
 Set clmx = Klient.ListView1.ColumnHeaders.Add(, , "Typ ud·losti", 60)
 Set clmx = Klient.ListView1.ColumnHeaders.Add(, , "Ud·lost", 500)
 Set clmx = Klient.ListView1.ColumnHeaders.Add(, , "Detail", 30)
 Set clmx = Klient.ListView1.ColumnHeaders.Add(, , "Provedl", 100)
 
 Set clmx = Klient.ListView2.ColumnHeaders.Add(, , "»Ìslo smlouvy", 132)
 
 Set clmx = Klient.ListView3.ColumnHeaders.Add(, , "ID", 30)
 Set clmx = Klient.ListView3.ColumnHeaders.Add(, , "UpozornÏnÌ", 480)
 Set clmx = Klient.ListView3.ColumnHeaders.Add(, , "Deadline", 100)
 Set clmx = Klient.ListView3.ColumnHeaders.Add(, , "Zb˝v·", 100)
 Set clmx = Klient.ListView3.ColumnHeaders.Add(, , "Assignee", 100)
 
' Set clmx = Klient.ListView4.ColumnHeaders.Add(, , "⁄daj", 152) ' RISK-1101
' Set clmx = Klient.ListView4.ColumnHeaders.Add(, , "Hodnota", 100)

 Set clmx = Klient.ListView5.ColumnHeaders.Add(, , "Data", 152)
 Set clmx = Klient.ListView5.ColumnHeaders.Add(, , "Hodnota", 100)
 
 Set clmx = Klient.ListView6.ColumnHeaders.Add(, , "⁄daj", 152)
 Set clmx = Klient.ListView6.ColumnHeaders.Add(, , "Hodnota", 150)
 
 
' Set clmX = Klient.ListView6.ColumnHeaders.Add(, , "OkresnÌ soud", 100)
' Set clmX = Klient.ListView6.ColumnHeaders.Add(, , "Krajsk˝ soud", 100)
' Set clmX = Klient.ListView6.ColumnHeaders.Add(, , "Kontrola v˝poËtu", 100)
' Set clmX = Klient.ListView6.ColumnHeaders.Add(, , "OBS pohled·vka", 100)
' Set clmX = Klient.ListView6.ColumnHeaders.Add(, , "PM pohled·vka", 100)
' Set clmX = Klient.ListView6.ColumnHeaders.Add(, , "»Ìslo jednacÌ", 100)
'   Set itmX = Klient.ListView6.ListItems.Add()
'  itmX.text = "OkresnÌ soud"
'     Set itmX = Klient.ListView6.ListItems.Add()
'  itmX.text = "Krajsk˝ soud"
'     Set itmX = Klient.ListView6.ListItems.Add()
'  itmX.text = "Kontrola v˝poËtu"
'     Set itmX = Klient.ListView6.ListItems.Add()
'  itmX.text = "OBS pohled·vka"
'
'
Case "Klient_exe"
 Klient_exe.TextBox1.text = ""

 Set clmx = Klient_exe.ListView1.ColumnHeaders.Add(, , "ID", 30)
 Set clmx = Klient_exe.ListView1.ColumnHeaders.Add(, , "Datum ud·losti", 90)
 Set clmx = Klient_exe.ListView1.ColumnHeaders.Add(, , "Typ ud·losti", 60)
 Set clmx = Klient_exe.ListView1.ColumnHeaders.Add(, , "Ud·lost", 500)
 Set clmx = Klient_exe.ListView1.ColumnHeaders.Add(, , "Detail", 30)
 Set clmx = Klient_exe.ListView1.ColumnHeaders.Add(, , "Provedl", 100)
 
 Set clmx = Klient_exe.ListView2.ColumnHeaders.Add(, , "»Ìslo smlouvy", 132)
 
 Set clmx = Klient_exe.ListView3.ColumnHeaders.Add(, , "ID", 30)
 Set clmx = Klient_exe.ListView3.ColumnHeaders.Add(, , "UpozorninÌ", 480)
 Set clmx = Klient_exe.ListView3.ColumnHeaders.Add(, , "Deadline", 100)
 Set clmx = Klient_exe.ListView3.ColumnHeaders.Add(, , "Zb˝v·", 100)
 Set clmx = Klient_exe.ListView3.ColumnHeaders.Add(, , "Assignee", 100)
 
' Set clmx = Klient_exe.ListView4.ColumnHeaders.Add(, , "⁄daj", 152) ' RISK-1101
' Set clmx = Klient_exe.ListView4.ColumnHeaders.Add(, , "Hodnota", 100)

 Set clmx = Klient_exe.ListView5.ColumnHeaders.Add(, , "Data", 152)
 Set clmx = Klient_exe.ListView5.ColumnHeaders.Add(, , "Hodnota", 100)



Case "Klient_dedictvi"
     Klient_dedictvi.TextBox1.text = ""
    
    'RISK-468
    If Klient_dedictvi.ListView1.ColumnHeaders.count = 0 Then
    
        Set clmx = Klient_dedictvi.ListView1.ColumnHeaders.Add(, , "ID", 30)
        Set clmx = Klient_dedictvi.ListView1.ColumnHeaders.Add(, , "Datum ud·losti", 90)
        Set clmx = Klient_dedictvi.ListView1.ColumnHeaders.Add(, , "Typ ud·losti", 60)
        Set clmx = Klient_dedictvi.ListView1.ColumnHeaders.Add(, , "Ud·lost", 500)
        Set clmx = Klient_dedictvi.ListView1.ColumnHeaders.Add(, , "Detail", 30)
        Set clmx = Klient_dedictvi.ListView1.ColumnHeaders.Add(, , "Provedl", 100)
     
    End If
     
    If Klient_dedictvi.ListView2.ColumnHeaders.count = 0 Then
        Set clmx = Klient_dedictvi.ListView2.ColumnHeaders.Add(, , "»Ìslo smlouvy", 132)
        Set clmx = Klient_dedictvi.ListView2.ColumnHeaders.Add(, , "Status", 120)
    End If
     
    If Klient_dedictvi.ListView3.ColumnHeaders.count = 0 Then
     
        Set clmx = Klient_dedictvi.ListView3.ColumnHeaders.Add(, , "ID", 30)
        Set clmx = Klient_dedictvi.ListView3.ColumnHeaders.Add(, , "UpozornÏnÌ", 480)
        Set clmx = Klient_dedictvi.ListView3.ColumnHeaders.Add(, , "Deadline", 100)
        Set clmx = Klient_dedictvi.ListView3.ColumnHeaders.Add(, , "Zb˝v·", 100)
        Set clmx = Klient_dedictvi.ListView3.ColumnHeaders.Add(, , "Assignee", 100)
     
    End If
    
     'RISK-468 - odstranÏno, kv˘li vymaz·nÌ p˘vodnÌho listview
    'Set clmx = Klient_dedictvi.ListView5.ColumnHeaders.Add(, , "⁄daj", 152)
    'Set clmx = Klient_dedictvi.ListView5.ColumnHeaders.Add(, , "Hodnota", 100)
     

 
 Case "Klient_trest"
 Klient_trest.TextBox1.text = ""

 Set clmx = Klient_trest.ListView1.ColumnHeaders.Add(, , "ID", 30)
 Set clmx = Klient_trest.ListView1.ColumnHeaders.Add(, , "Datum ud·losti", 90)
 Set clmx = Klient_trest.ListView1.ColumnHeaders.Add(, , "Typ ud·losti", 60)
 Set clmx = Klient_trest.ListView1.ColumnHeaders.Add(, , "Ud·lost", 500)
 Set clmx = Klient_trest.ListView1.ColumnHeaders.Add(, , "Detail", 30)
 Set clmx = Klient_trest.ListView1.ColumnHeaders.Add(, , "Provedl", 100)
 
 Set clmx = Klient_trest.ListView2.ColumnHeaders.Add(, , "»Ìslo smlouvy", 132)
 Set clmx = Klient_trest.ListView2.ColumnHeaders.Add(, , "Status", 120)
 
 Set clmx = Klient_trest.ListView3.ColumnHeaders.Add(, , "ID", 30)
 Set clmx = Klient_trest.ListView3.ColumnHeaders.Add(, , "UpozornÏnÌ", 480)
 Set clmx = Klient_trest.ListView3.ColumnHeaders.Add(, , "Deadline", 100)
 Set clmx = Klient_trest.ListView3.ColumnHeaders.Add(, , "Zb˝v·", 100)
 Set clmx = Klient_trest.ListView3.ColumnHeaders.Add(, , "Assignee", 100)
 
 Set clmx = Klient_trest.ListView4.ColumnHeaders.Add(, , "⁄daj", 152)
 Set clmx = Klient_trest.ListView4.ColumnHeaders.Add(, , "Hodnota", 100)


 Case "Klient_insol"
 Klient_insol.TextBox1.text = ""

 Set clmx = Klient_insol.ListView1.ColumnHeaders.Add(, , "ID", 30)
 Set clmx = Klient_insol.ListView1.ColumnHeaders.Add(, , "Datum ud·losti", 90)
 Set clmx = Klient_insol.ListView1.ColumnHeaders.Add(, , "Typ ud·losti", 60)
 Set clmx = Klient_insol.ListView1.ColumnHeaders.Add(, , "Ud·lost", 500)
 Set clmx = Klient_insol.ListView1.ColumnHeaders.Add(, , "Detail", 30)
 Set clmx = Klient_insol.ListView1.ColumnHeaders.Add(, , "Provedl", 100)
 
 Set clmx = Klient_insol.ListView2.ColumnHeaders.Add(, , "»Ìslo smlouvy", 132)
 Set clmx = Klient_insol.ListView2.ColumnHeaders.Add(, , "Status", 120)
 
 Set clmx = Klient_insol.ListView3.ColumnHeaders.Add(, , "ID", 30)
 Set clmx = Klient_insol.ListView3.ColumnHeaders.Add(, , "UpozornÏnÌ", 480)
 Set clmx = Klient_insol.ListView3.ColumnHeaders.Add(, , "Deadline", 100)
 Set clmx = Klient_insol.ListView3.ColumnHeaders.Add(, , "Zb˝v·", 100)
 Set clmx = Klient_insol.ListView3.ColumnHeaders.Add(, , "Assignee", 100)
 
 Set clmx = Klient_insol.ListView4.ColumnHeaders.Add(, , "⁄daj", 152)
 Set clmx = Klient_insol.ListView4.ColumnHeaders.Add(, , "Hodnota", 100)


' Case "SOP"       ' presunuto do objektu SOP_class v ramci 'RISK-1101 jez.v.5.4.1
' Case "SOP_exe"   ' presunuto do objektu SOP_class v ramci 'RISK-1101 jez.v.5.4.1
 
  
Case "Souhrn"
 
 Set clmx = Souhrn.ListView3.ColumnHeaders.Add(, , "ID_PERSON", 0)
 Set clmx = Souhrn.ListView3.ColumnHeaders.Add(, , "CUID", 40)
 Set clmx = Souhrn.ListView3.ColumnHeaders.Add(, , "Klient", 90)
 Set clmx = Souhrn.ListView3.ColumnHeaders.Add(, , "UpozornÏnÌ", 146)
 Set clmx = Souhrn.ListView3.ColumnHeaders.Add(, , "Deadline", 55)
 Set clmx = Souhrn.ListView3.ColumnHeaders.Add(, , "Zb˝v·", 40)
 Set clmx = Souhrn.ListView3.ColumnHeaders.Add(, , "SKP_PROCESS", 0) 'RISK-130
 Set clmx = Souhrn.ListView3.ColumnHeaders.Add(, , "Assignee", 55)
 Set clmx = Souhrn.ListView3.ColumnHeaders.Add(, , "SKP_NOTICE", 0)
 
 Call NastavKliceListview(Souhrn.ListView3)
 
 With Souhrn.ListView4
    Call .ColumnHeaders.Add(, , "ID_PERSON", 0)
    Call .ColumnHeaders.Add(, , "CUID", 50)
    Call .ColumnHeaders.Add(, , "Klient", 100)
    Call .ColumnHeaders.Add(, , "NutnÈ autorizovat", 90)
    Call .ColumnHeaders.Add(, , "OkresnÌ soud", 70)
    Call .ColumnHeaders.Add(, , "»·stka", 40)
    Call .ColumnHeaders.Add(, , "Datum k Ë·stce", 80)
 End With
 
 Set clmx = Souhrn.LV_KontrolaPM.ColumnHeaders.Add(, , "ID_PERSON", 0)
 Set clmx = Souhrn.LV_KontrolaPM.ColumnHeaders.Add(, , "CUID", 70)
 Set clmx = Souhrn.LV_KontrolaPM.ColumnHeaders.Add(, , "Klient", 100)
 Set clmx = Souhrn.LV_KontrolaPM.ColumnHeaders.Add(, , "UpozornÏnÌ", 146)
 Set clmx = Souhrn.LV_KontrolaPM.ColumnHeaders.Add(, , "Deadline", 70)
 Set clmx = Souhrn.LV_KontrolaPM.ColumnHeaders.Add(, , "Zb˝v·", 40)
 
 Set clmx = Souhrn.LV_Sonik.ColumnHeaders.Add(, , "SKP_DATA_MESSAGE_NOTICE", 0)
 Set clmx = Souhrn.LV_Sonik.ColumnHeaders.Add(, , "SKP_DATA_MESSAGE", 0)
 Set clmx = Souhrn.LV_Sonik.ColumnHeaders.Add(, , "CUIDy", 100)
 Set clmx = Souhrn.LV_Sonik.ColumnHeaders.Add(, , "Klienti", 70)
 Set clmx = Souhrn.LV_Sonik.ColumnHeaders.Add(, , "UpozornÏnÌ", 96)
 Set clmx = Souhrn.LV_Sonik.ColumnHeaders.Add(, , "Deadline", 60)
 Set clmx = Souhrn.LV_Sonik.ColumnHeaders.Add(, , "Zb˝v·", 40) 'suma 426
 Set clmx = Souhrn.LV_Sonik.ColumnHeaders.Add(, , "»Ìslo DZ", 60)
 
Case "Stavy"
 Set clmx = Stavy.ListView1.ColumnHeaders.Add(, , "ID", 20)
 Set clmx = Stavy.ListView1.ColumnHeaders.Add(, , "N·zev stavu", 195)
 Set clmx = Stavy.ListView1.ColumnHeaders.Add(, , "N·zev procesu", 90)
 Set clmx = Stavy.ListView1.ColumnHeaders.Add(, , "Platnost od", 85)
 Set clmx = Stavy.ListView1.ColumnHeaders.Add(, , "Platnost do", 88)
 Set clmx = Stavy.ListView1.ColumnHeaders.Add(, , "PoslednÌ zmÏna", 85)
 Set clmx = Stavy.ListView1.ColumnHeaders.Add(, , "Provedl", 100)

Case "Update_stav_all"
 Set clmx = Update_stav_all.ListViewHistorieStavu.ColumnHeaders.Add(, , "ID", 30)
 Set clmx = Update_stav_all.ListViewHistorieStavu.ColumnHeaders.Add(, , "N·zev stavu", 195)
 Set clmx = Update_stav_all.ListViewHistorieStavu.ColumnHeaders.Add(, , "Proces", 50)
 Set clmx = Update_stav_all.ListViewHistorieStavu.ColumnHeaders.Add(, , "Status", 50)
 Set clmx = Update_stav_all.ListViewHistorieStavu.ColumnHeaders.Add(, , "Platnost od", 88)
 Set clmx = Update_stav_all.ListViewHistorieStavu.ColumnHeaders.Add(, , "Platnost do", 88)
 Set clmx = Update_stav_all.ListViewHistorieStavu.ColumnHeaders.Add(, , "Pozn·mka", 140)
 Set clmx = Update_stav_all.ListViewHistorieStavu.ColumnHeaders.Add(, , "PoslednÌ zmÏna", 88)
 Set clmx = Update_stav_all.ListViewHistorieStavu.ColumnHeaders.Add(, , "Provedl", 80)
 Set clmx = Update_stav_all.ListViewHistorieStavu.ColumnHeaders.Add(, , "Smaz·no", 50)
 
 Call NastavKliceListview(Update_stav_all.ListViewHistorieStavu)

'' nasledujici DataCheck vyhozen ve verzj jez.v.5.7
'Case "DataCheck"
' Set clmx = DataCheck.ListView1.ColumnHeaders.Add(, , "»Ìslo smlouvy/˙Ëtu", 70)
' Set clmx = DataCheck.ListView1.ColumnHeaders.Add(, , "Status", 70)
' Set clmx = DataCheck.ListView1.ColumnHeaders.Add(, , "Typ", 60)
'
' Set clmx = DataCheck.ListView2.ColumnHeaders.Add(, , "Platnost od (O⁄)", 80)
' Set clmx = DataCheck.ListView2.ColumnHeaders.Add(, , "Platnost do (O⁄)", 80)
' Set clmx = DataCheck.ListView2.ColumnHeaders.Add(, , "V˝öe O⁄", 80)
' Set clmx = DataCheck.ListView2.ColumnHeaders.Add(, , "Platnost od (S⁄)", 80)
' Set clmx = DataCheck.ListView2.ColumnHeaders.Add(, , "Platnost do (S⁄)", 80)
' Set clmx = DataCheck.ListView2.ColumnHeaders.Add(, , "V˝öe S⁄", 100)
' Set clmx = DataCheck.ListView2.ColumnHeaders.Add(, , "V˝öe jistiny", 100)
'
' Set clmx = DataCheck.ListView3.ColumnHeaders.Add(, , "Datum platby", 100)
' Set clmx = DataCheck.ListView3.ColumnHeaders.Add(, , "Datum spl·tky", 100)
' Set clmx = DataCheck.ListView3.ColumnHeaders.Add(, , "V˝öe spl·tky", 90)
' Set clmx = DataCheck.ListView3.ColumnHeaders.Add(, , "Typ p¯edpisu", 130)
' Set clmx = DataCheck.ListView3.ColumnHeaders.Add(, , "V˝öe p¯edpisu", 90)
' Set clmx = DataCheck.ListView3.ColumnHeaders.Add(, , "V˝öe platby", 90)
'
' Set clmx = DataCheck.ListView4.ColumnHeaders.Add(, , "Datum vygenerov·nÌ transakce ", 100)
' Set clmx = DataCheck.ListView4.ColumnHeaders.Add(, , "Datum za˙Ëtov·nÌ", 100)
' Set clmx = DataCheck.ListView4.ColumnHeaders.Add(, , "N·zev transakce", 250)
' Set clmx = DataCheck.ListView4.ColumnHeaders.Add(, , "V˝öe transakce", 100)


Case "Pridanie_klienta"
 Set clmx = Pridanie_klienta.ListView1.ColumnHeaders.Add(, , "ID", 30)
 Set clmx = Pridanie_klienta.ListView1.ColumnHeaders.Add(, , "CUID", 50)
 Set clmx = Pridanie_klienta.ListView1.ColumnHeaders.Add(, , "P¯id·n", 40)
 Set clmx = Pridanie_klienta.ListView1.ColumnHeaders.Add(, , "Zadal", 70)
 Set clmx = Pridanie_klienta.ListView1.ColumnHeaders.Add(, , "Datum poûadavku", 92)
 
End Select

End Sub





Sub text_information_hide(page As String)
' Hiding information Labels
' SkrytÌ doplÚujÌcÌch informacÌ na str·nce

Select Case page

Case "Zaloby"
 Zaloby.Label27.Visible = False
 Zaloby.Label28.Visible = False
 Zaloby.Label29.Visible = False
 
Case "Exekuce"
 Exekuce.Label27.Visible = False
 Exekuce.Label28.Visible = False
 Exekuce.Label29.Visible = False
 
 Case "Dedictvi"
 Dedictvi.Label27.Visible = False
 Dedictvi.Label28.Visible = False
 Dedictvi.Label29.Visible = False

Case "Zesplatneni"
 Zesplatneni.Label27.Visible = False
 Zesplatneni.Label28.Visible = False
 Zesplatneni.Label29.Visible = False
 
Case "Nezesplatneni"
 Nezesplatneni.Label27.Visible = False
 Nezesplatneni.Label28.Visible = False
 Nezesplatneni.Label29.Visible = False

Case "Klient"
 Klient.Label27.Visible = False
 Klient.Label28.Visible = False
 Klient.Label29.Visible = False
 

Case "Klient_exe"
 Klient.Label27.Visible = False
 Klient.Label28.Visible = False
 Klient.Label29.Visible = False
  
 
Case "Trest"
 Trest.Label27.Visible = False
 Trest.Label28.Visible = False
 Trest.Label29.Visible = False
 
Case "Insol"
 Insol.Label27.Visible = False
 Insol.Label28.Visible = False
 Insol.Label29.Visible = False

Case "Insol_nesparovane"
 Insol_nesparovane.Label27.Visible = False
 Insol_nesparovane.Label28.Visible = False
 Insol_nesparovane.Label29.Visible = False
 
Case "Klient_dedictvi"
 Klient_dedictvi.Label27.Visible = False
 Klient_dedictvi.Label28.Visible = False
 Klient_dedictvi.Label29.Visible = False
 
 
 
Case "Klient_trest"
 Klient_trest.Label27.Visible = False
 Klient_trest.Label28.Visible = False
 Klient_trest.Label29.Visible = False
 
Case "Klient_insol"
 Klient_insol.Label27.Visible = False
 Klient_insol.Label28.Visible = False
 Klient_insol.Label29.Visible = False
 
Case "Update_klient"
 Update_klient.Label28.Visible = False
 Update_klient.Label29.Visible = False
 
 Case "Update_klient_exe"
 Update_klient.Label28.Visible = False
 Update_klient.Label29.Visible = False
 
 Case "Update_klient_ded"
 Update_klient_ded.Label28.Visible = False
 Update_klient_ded.Label29.Visible = False
 
Case "Update_klient_trest"
 Update_klient_trest.Label28.Visible = False
 Update_klient_trest.Label29.Visible = False
 
Case "Update_klient_insol"
 Update_klient_insol.Label28.Visible = False
 Update_klient_insol.Label29.Visible = False
 
Case "Update_SOP"
 Update_SOP.Label27.Visible = False
 Update_SOP.Label28.Visible = False
 Update_SOP.Label29.Visible = False

 
Case "Update_auth"
 Update_auth.Label28.Visible = False
 Update_auth.Label29.Visible = False
 
Case "Stavy"
 Stavy.Label27.Visible = False
 Stavy.Label28.Visible = False
 Stavy.Label29.Visible = False
 
Case "Stavy_update"
 Stavy.Label27.Visible = False
 Stavy.Label28.Visible = False
 Stavy.Label29.Visible = False
 
Case "Update_stav_all"
 Update_stav_all.Label27.Visible = False
 Update_stav_all.Label28.Visible = False
 Update_stav_all.Label29.Visible = False
 Update_stav_all.Label30.Visible = False

Case "DataCheck"
 DataCheck.Label27.Visible = False
 DataCheck.Label28.Visible = False
 DataCheck.Label29.Visible = False

End Select


End Sub





' Nastavov·nÌ filtru
' Moûnosti:
'  - 0 - Vynulov·nÌ filtru (PrvnÌ otev¯enÌ, zruöenÌ filtr˘ pomocÌ tlaËÌtka)
'  - 1 - NastavenÌ filtru (TlaËÌtko vyhledat, zav¯enÌ formul·¯e)
Sub set_filter(form As String, code As Long)

Select Case form
 Case "Zesplatneni"
  Select Case code
   Case 0
    zes_filtr.set = True
    zes_filtr.CUID = ""
    zes_filtr.jmeno = ""
    zes_filtr.rc = ""
    zes_filtr.smlouva = ""
    zes_filtr.date = False
    zes_filtr.od = date
    zes_filtr.do = date
    zes_filtr.duvod = "Vöe"
    zes_filtr.code = "é·dn˝"
   Case 1
    zes_filtr.set = True
    zes_filtr.CUID = Zesplatneni.TextBox1.text
    zes_filtr.jmeno = Zesplatneni.TextBox2.text
    zes_filtr.rc = Zesplatneni.TextBox3.text
    zes_filtr.smlouva = Zesplatneni.TextBox4.text
    zes_filtr.date = Zesplatneni.CheckBox1.value
    Zesplatneni.CheckBox1.value = True
    zes_filtr.od = Zesplatneni.DTPicker1
    zes_filtr.do = Zesplatneni.DTPicker2
    Zesplatneni.CheckBox1.value = zes_filtr.date
    
    If Zesplatneni.ComboBox1.text <> "" Then
     zes_filtr.duvod = Zesplatneni.ComboBox1.text
    Else
     zes_filtr.duvod = "Vöe"
    End If
    
    If Zesplatneni.ComboBox2.text <> "" Then
     zes_filtr.code = Zesplatneni.ComboBox2.text
    Else
     zes_filtr.code = "é·dn˝"
    End If
  
  End Select
  
 Case "Nezesplatneni"
  Select Case code
   Case 0
    nezes_filtr.set = True
    nezes_filtr.CUID = ""
    nezes_filtr.jmeno = ""
    nezes_filtr.rc = ""
    nezes_filtr.smlouva = ""
    nezes_filtr.date = False
    nezes_filtr.od = date
    nezes_filtr.do = date
    nezes_filtr.duvod = "Vöe"
    nezes_filtr.code = "é·dn˝"
   Case 1
    nezes_filtr.set = True
    nezes_filtr.CUID = Nezesplatneni.TextBox1.text
    nezes_filtr.jmeno = Nezesplatneni.TextBox2.text
    nezes_filtr.rc = Nezesplatneni.TextBox3.text
    nezes_filtr.smlouva = Nezesplatneni.TextBox4.text
    nezes_filtr.date = Nezesplatneni.CheckBox1.value
    Nezesplatneni.CheckBox1.value = True
    nezes_filtr.od = Nezesplatneni.DTPicker1
    nezes_filtr.do = Nezesplatneni.DTPicker2
    Nezesplatneni.CheckBox1.value = nezes_filtr.date
    
    If Nezesplatneni.ComboBox1.text <> "" Then
     nezes_filtr.duvod = Nezesplatneni.ComboBox1.text
    Else
     nezes_filtr.duvod = "Vöe"
    End If
    
    If Nezesplatneni.ComboBox2.text <> "" Then
     nezes_filtr.code = Nezesplatneni.ComboBox2.text
    Else
     nezes_filtr.code = "é·dn˝"
    End If
  
  End Select
  
 Case "Exekuce"
  Select Case code
   Case 0
    exe_filtr.set = True
    exe_filtr.CUID = ""
    exe_filtr.jmeno = ""
    exe_filtr.rc = ""
    exe_filtr.stav = "Vöe"
    'zal_filtr.proces = "Vöe" '
    exe_filtr.date = False
    exe_filtr.od = date
    exe_filtr.do = date
    exe_filtr.spis = ""
   Case 1
    exe_filtr.set = True
    exe_filtr.CUID = Exekuce.TextBox1.text
    exe_filtr.jmeno = Exekuce.TextBox2.text
    exe_filtr.rc = Exekuce.TextBox3.text
    exe_filtr.stav = Exekuce.ComboBox1.text
    'zal_filtr.proces = Zaloby.ComboBox2.Text  '
    exe_filtr.date = Exekuce.CheckBox1.value
    Exekuce.CheckBox1.value = True
    exe_filtr.od = Exekuce.DTPicker1
    exe_filtr.do = Exekuce.DTPicker2
    Exekuce.CheckBox1.value = exe_filtr.date
    exe_filtr.spis = Zaloby.TextBox4.text
  End Select
    
  
 Case "Zaloby"
  Select Case code
   Case 0
    zal_filtr.set = True
    zal_filtr.CUID = ""
    zal_filtr.jmeno = ""
    zal_filtr.rc = ""
    zal_filtr.stav = "Vöe"
    'zal_filtr.proces = "Vöe" '
    zal_filtr.date = False
    zal_filtr.od = date
    zal_filtr.do = date
    zal_filtr.spis = ""
   Case 1
    zal_filtr.set = True
    zal_filtr.CUID = Zaloby.TextBox1.text
    zal_filtr.jmeno = Zaloby.TextBox2.text
    zal_filtr.rc = Zaloby.TextBox3.text
    zal_filtr.stav = Zaloby.ComboBox1.text
    'zal_filtr.proces = Zaloby.ComboBox2.Text  '
    zal_filtr.date = Zaloby.CheckBox1.value
    Zaloby.CheckBox1.value = True
    zal_filtr.od = Zaloby.DTPicker1
    zal_filtr.do = Zaloby.DTPicker2
    Zaloby.CheckBox1.value = zal_filtr.date
    zal_filtr.spis = Zaloby.TextBox4.text
  End Select
  
'   Case "Stavy"
'  Select Case code
'   Case 0
'    filtr_proces = ""
'   Case 1
'    filtr_proces = Stavy.ComboBox1.Text
'  End Select
  
Case "Dedictvi"
  Select Case code
   Case 0
    ded_filtr.set = True
    ded_filtr.CUID = ""
    ded_filtr.jmeno = ""
    ded_filtr.rc = ""
    ded_filtr.stav = "Vöe"
    'ded_filtr.proces = "Vöe" '
    ded_filtr.date = False
    ded_filtr.od = date
    ded_filtr.do = date
    ded_filtr.spis = ""
   Case 1
    ded_filtr.set = True
    ded_filtr.CUID = Dedictvi.TextBox1.text
    ded_filtr.jmeno = Dedictvi.TextBox2.text
    ded_filtr.rc = Dedictvi.TextBox3.text
    ded_filtr.stav = Dedictvi.ComboBox1.text
    'ded_filtr.proces = Dedictvi.ComboBox2.Text  '
    ded_filtr.date = Dedictvi.CheckBox1.value
    Dedictvi.CheckBox1.value = True
    ded_filtr.od = Dedictvi.DTPicker1
    ded_filtr.do = Dedictvi.DTPicker2
    Dedictvi.CheckBox1.value = ded_filtr.date
    ded_filtr.spis = Dedictvi.TextBox4.text
  End Select
  
Case "Trest"
  Select Case code
   Case 0
    trest_filtr.set = True
    trest_filtr.CUID = ""
    trest_filtr.jmeno = ""
    trest_filtr.rc = ""
    trest_filtr.stav = "Vöe"
    'trest_filtr.proces = "Vöe" '
    trest_filtr.date = False
    trest_filtr.od = date
    trest_filtr.do = date
    trest_filtr.spis = ""
   Case 1
    trest_filtr.set = True
    trest_filtr.CUID = Trest.TextBox1.text
    trest_filtr.jmeno = Trest.TextBox2.text
    trest_filtr.rc = Trest.TextBox3.text
    trest_filtr.stav = Trest.ComboBox1.text
    'trest_filtr.proces = Trest.ComboBox2.Text  '
    trest_filtr.date = Trest.CheckBox1.value
    Trest.CheckBox1.value = True
    trest_filtr.od = Trest.DTPicker1
    trest_filtr.do = Trest.DTPicker2
    Trest.CheckBox1.value = trest_filtr.date
    trest_filtr.spis = Trest.TextBox4.text
  End Select
  
Case "Insol"
  Select Case code
   Case 0
    ins_filtr.set = True
    ins_filtr.CUID = ""
    ins_filtr.jmeno = ""
    ins_filtr.rc = ""
    ins_filtr.stav = "Vöe"
    'trest_filtr.proces = "Vöe" '
    ins_filtr.date = False
    ins_filtr.od = date
    ins_filtr.do = date
    ins_filtr.spis = ""
   Case 1
    ins_filtr.set = True
    ins_filtr.CUID = Insol.TextBox1.text
    ins_filtr.jmeno = Insol.TextBox2.text
    ins_filtr.rc = Insol.TextBox3.text
    ins_filtr.stav = Insol.ComboBox1.text
    'trest_filtr.proces = Trest.ComboBox2.Text  '
    ins_filtr.date = Insol.CheckBox1.value
    Insol.CheckBox1.value = True
    ins_filtr.od = Insol.DTPicker1
    ins_filtr.do = Insol.DTPicker2
    Insol.CheckBox1.value = ins_filtr.date
    ins_filtr.spis = Insol.TextBox4.text
  End Select
  
End Select

End Sub








''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''   Subs pre menöie exporty    '''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''







' Export to PDF
Sub export_pdf(oblast As String)
' FA 8.3.2018: Tlacitko "export do PDF" na kartach Klient_...
' > tlacitko je zasedle a nepouziva sa (ale nekontroloval som ci na vsetkych kartach)


Select Case oblast
Case "Klient"

    ActiveSheet.ExportAsFixedFormat fileName:="Export_" & Range("D4") & "_" & Range("D5") & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".pdf", Type:=xlTypePDF, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True

Worksheets(oblast).Visible = False

Case "Klient_dedictvi"

    ActiveSheet.ExportAsFixedFormat fileName:="Export_" & Range("D4") & "_" & Range("D5") & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".pdf", Type:=xlTypePDF, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True

Worksheets(oblast).Visible = False

Case "Klient_insol"

    ActiveSheet.ExportAsFixedFormat fileName:="Export_" & Range("D4") & "_" & Range("D5") & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".pdf", Type:=xlTypePDF, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True

Worksheets(oblast).Visible = False

Case "Klient_trest"

    ActiveSheet.ExportAsFixedFormat fileName:="Export_" & Range("D4") & "_" & Range("D5") & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".pdf", Type:=xlTypePDF, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True

Worksheets(oblast).Visible = False

End Select

End Sub


Sub export_novy(oblast As String)
' Export to new WorkBook
Dim name As String
    
If oblast = "Souhrn" Then
    cfg.wb_Jez.Activate
    Worksheets("SouhrnGraf").Visible = True
    Sheets("SouhrnGraf").Select
    Sheets("SouhrnGraf").Copy
    name = "Export_" & Replace(Replace(Replace(CStr(Now()), " ", ""), ":", ""), ".", "")
    ActiveSheet.name = name
    cfg.wb_Jez.Worksheets("SouhrnGraf").Visible = False
Else

    cfg.wb_Jez.Activate
    Sheets("Export").Select
    Sheets("Export").Copy
    name = "Export_" & formular & "_" & Replace(Replace(Replace(CStr(Now()), " ", ""), ":", ""), ".", "")
    ActiveSheet.name = name
    
    cfg.wb_Jez.Worksheets("Export").Visible = False
End If
End Sub


Sub Export(oblast As String)
' Export to new Sheet (from ListView component)
    cfg.wb_Jez.Activate
    Sheets.Add after:=Sheets(Sheets.count)

On Error GoTo no_exist
    Sheets(Sheets.count).name = "Exp"
    Sheets("Export").Visible = True
    Sheets("Export").Select
    ActiveWindow.SelectedSheets.Delete


no_exist:

    Sheets("Exp").name = "Export"
    Sheets("Export").Select


Select Case oblast

    Case "Zesplatneni"
         ActiveSheet.Range("A1") = Zesplatneni.ListView1.ColumnHeaders(2)
         ActiveSheet.Range("B1") = Zesplatneni.ListView1.ColumnHeaders(3)
         ActiveSheet.Range("C1") = Zesplatneni.ListView1.ColumnHeaders(4)
         ActiveSheet.Range("D1") = Zesplatneni.ListView1.ColumnHeaders(5)
         ActiveSheet.Range("E1") = Zesplatneni.ListView1.ColumnHeaders(6)
         ActiveSheet.Range("F1") = Zesplatneni.ListView1.ColumnHeaders(7)
         ActiveSheet.Range("G1") = Zesplatneni.ListView1.ColumnHeaders(8)
         ActiveSheet.Range("H1") = Zesplatneni.ListView1.ColumnHeaders(9)
         ActiveSheet.Range("I1") = Zesplatneni.ListView1.ColumnHeaders(10)

         For i = 2 To Zesplatneni.ListView1.ListItems.count + 1
          ActiveSheet.Range("A" & i) = Zesplatneni.ListView1.ListItems.item(i - 1).SubItems(1)
          ActiveSheet.Range("B" & i) = Zesplatneni.ListView1.ListItems.item(i - 1).SubItems(2)
          ActiveSheet.Range("C" & i) = Zesplatneni.ListView1.ListItems.item(i - 1).SubItems(3)
          ActiveSheet.Range("D" & i) = Zesplatneni.ListView1.ListItems.item(i - 1).SubItems(4)
          ActiveSheet.Range("E" & i) = Zesplatneni.ListView1.ListItems.item(i - 1).SubItems(5)
          ActiveSheet.Range("F" & i) = Zesplatneni.ListView1.ListItems.item(i - 1).SubItems(6)
          ActiveSheet.Range("G" & i) = Zesplatneni.ListView1.ListItems.item(i - 1).SubItems(7)
          ActiveSheet.Range("H" & i) = Zesplatneni.ListView1.ListItems.item(i - 1).SubItems(8)
          ActiveSheet.Range("I" & i) = Zesplatneni.ListView1.ListItems.item(i - 1).SubItems(9)
         Next i

        ' It depends on the number of columns that I want to Export
        Range("A1:I1").Select

        Range(Selection, Selection.End(xlDown)).Select

        Selection.WrapText = False

        ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).name = "Tabulka"
        Range("Tabulka[#All]").Select
        ActiveSheet.ListObjects("Tabulka").TableStyle = "TableStyleMedium18"

        Columns("A:H").EntireColumn.AutoFit

    Case "Nezesplatneni"
         ActiveSheet.Range("A1") = Nezesplatneni.ListView1.ColumnHeaders(2)
         ActiveSheet.Range("B1") = Nezesplatneni.ListView1.ColumnHeaders(3)
         ActiveSheet.Range("C1") = Nezesplatneni.ListView1.ColumnHeaders(4)
         ActiveSheet.Range("D1") = Nezesplatneni.ListView1.ColumnHeaders(5)
         ActiveSheet.Range("E1") = Nezesplatneni.ListView1.ColumnHeaders(6)
         ActiveSheet.Range("F1") = Nezesplatneni.ListView1.ColumnHeaders(7)
         ActiveSheet.Range("G1") = Nezesplatneni.ListView1.ColumnHeaders(8)
         ActiveSheet.Range("H1") = Nezesplatneni.ListView1.ColumnHeaders(9)
         ActiveSheet.Range("I1") = Nezesplatneni.ListView1.ColumnHeaders(10)

         For i = 2 To Nezesplatneni.ListView1.ListItems.count + 1
          ActiveSheet.Range("A" & i) = Nezesplatneni.ListView1.ListItems.item(i - 1).SubItems(1)
          ActiveSheet.Range("B" & i) = Nezesplatneni.ListView1.ListItems.item(i - 1).SubItems(2)
          ActiveSheet.Range("C" & i) = Nezesplatneni.ListView1.ListItems.item(i - 1).SubItems(3)
          ActiveSheet.Range("D" & i) = Nezesplatneni.ListView1.ListItems.item(i - 1).SubItems(4)
          ActiveSheet.Range("E" & i) = Nezesplatneni.ListView1.ListItems.item(i - 1).SubItems(5)
          ActiveSheet.Range("F" & i) = Nezesplatneni.ListView1.ListItems.item(i - 1).SubItems(6)
          ActiveSheet.Range("G" & i) = Nezesplatneni.ListView1.ListItems.item(i - 1).SubItems(7)
          ActiveSheet.Range("H" & i) = Nezesplatneni.ListView1.ListItems.item(i - 1).SubItems(8)
          ActiveSheet.Range("I" & i) = Nezesplatneni.ListView1.ListItems.item(i - 1).SubItems(9)
         Next i

        ' It depends on the number of columns that I want to Export
        Range("A1:I1").Select

        Range(Selection, Selection.End(xlDown)).Select

        Selection.WrapText = False

        ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).name = "Tabulka"
        Range("Tabulka[#All]").Select
        ActiveSheet.ListObjects("Tabulka").TableStyle = "TableStyleMedium18"

        Columns("A:H").EntireColumn.AutoFit

    Case "Exekuce"
         ActiveSheet.Range("A1") = Exekuce.ListView1.ColumnHeaders(1)
         ActiveSheet.Range("B1") = Exekuce.ListView1.ColumnHeaders(3)
         ActiveSheet.Range("C1") = Exekuce.ListView1.ColumnHeaders(4)
         ActiveSheet.Range("D1") = Exekuce.ListView1.ColumnHeaders(5)
         ActiveSheet.Range("E1") = Exekuce.ListView1.ColumnHeaders(6)
         ActiveSheet.Range("F1") = Exekuce.ListView1.ColumnHeaders(7)
         ActiveSheet.Range("G1") = Exekuce.ListView1.ColumnHeaders(8)

         For i = 2 To Zaloby.ListView1.ListItems.count + 1
          ActiveSheet.Range("A" & i) = Exekuce.ListView1.ListItems.item(i - 1)
          ActiveSheet.Range("B" & i) = Exekuce.ListView1.ListItems.item(i - 1).SubItems(2)
          ActiveSheet.Range("C" & i) = Exekuce.ListView1.ListItems.item(i - 1).SubItems(3)
          ActiveSheet.Range("D" & i) = Exekuce.ListView1.ListItems.item(i - 1).SubItems(4)
          ActiveSheet.Range("E" & i) = Exekuce.ListView1.ListItems.item(i - 1).SubItems(5)
          ActiveSheet.Range("F" & i) = Exekuce.ListView1.ListItems.item(i - 1).SubItems(6)
          ActiveSheet.Range("G" & i) = Exekuce.ListView1.ListItems.item(i - 1).SubItems(7)
         Next i

        ' It depends on the number of columns that I want to Export
        Range("A1:G1").Select

        Range(Selection, Selection.End(xlDown)).Select

        Selection.WrapText = False

        ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).name = "Tabulka"
        Range("Tabulka[#All]").Select
        ActiveSheet.ListObjects("Tabulka").TableStyle = "TableStyleMedium18"

        Columns("A:G").EntireColumn.AutoFit


    Case "Zaloby"
         ActiveSheet.Range("A1") = Zaloby.ListView1.ColumnHeaders(1)
         ActiveSheet.Range("B1") = Zaloby.ListView1.ColumnHeaders(3)
         ActiveSheet.Range("C1") = Zaloby.ListView1.ColumnHeaders(4)
         ActiveSheet.Range("D1") = Zaloby.ListView1.ColumnHeaders(5)
         ActiveSheet.Range("E1") = Zaloby.ListView1.ColumnHeaders(6)
         ActiveSheet.Range("F1") = Zaloby.ListView1.ColumnHeaders(7)
         ActiveSheet.Range("G1") = Zaloby.ListView1.ColumnHeaders(8)

         For i = 2 To Zaloby.ListView1.ListItems.count + 1
          ActiveSheet.Range("A" & i) = Zaloby.ListView1.ListItems.item(i - 1)
          ActiveSheet.Range("B" & i) = Zaloby.ListView1.ListItems.item(i - 1).SubItems(2)
          ActiveSheet.Range("C" & i) = Zaloby.ListView1.ListItems.item(i - 1).SubItems(3)
          ActiveSheet.Range("D" & i) = Zaloby.ListView1.ListItems.item(i - 1).SubItems(4)
          ActiveSheet.Range("E" & i) = Zaloby.ListView1.ListItems.item(i - 1).SubItems(5)
          ActiveSheet.Range("F" & i) = Zaloby.ListView1.ListItems.item(i - 1).SubItems(6)
          ActiveSheet.Range("G" & i) = Zaloby.ListView1.ListItems.item(i - 1).SubItems(7)
         Next i

        ' It depends on the number of columns that I want to Export
        Range("A1:G1").Select

        Range(Selection, Selection.End(xlDown)).Select

        Selection.WrapText = False

        ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).name = "Tabulka"
        Range("Tabulka[#All]").Select
        ActiveSheet.ListObjects("Tabulka").TableStyle = "TableStyleMedium18"

        Columns("A:G").EntireColumn.AutoFit

    Case "Trest"
         ActiveSheet.Range("A1") = Trest.ListView1.ColumnHeaders(1)
         ActiveSheet.Range("B1") = Trest.ListView1.ColumnHeaders(3)
         ActiveSheet.Range("C1") = Trest.ListView1.ColumnHeaders(4)
         ActiveSheet.Range("D1") = Trest.ListView1.ColumnHeaders(5)
         ActiveSheet.Range("E1") = Trest.ListView1.ColumnHeaders(6)
         ActiveSheet.Range("F1") = Trest.ListView1.ColumnHeaders(7)
         ActiveSheet.Range("G1") = Trest.ListView1.ColumnHeaders(8)

         For i = 2 To Trest.ListView1.ListItems.count + 1
          ActiveSheet.Range("A" & i) = Trest.ListView1.ListItems.item(i - 1)
          ActiveSheet.Range("B" & i) = Trest.ListView1.ListItems.item(i - 1).SubItems(2)
          ActiveSheet.Range("C" & i) = Trest.ListView1.ListItems.item(i - 1).SubItems(3)
          ActiveSheet.Range("D" & i) = Trest.ListView1.ListItems.item(i - 1).SubItems(4)
          ActiveSheet.Range("E" & i) = Trest.ListView1.ListItems.item(i - 1).SubItems(5)
          ActiveSheet.Range("F" & i) = Trest.ListView1.ListItems.item(i - 1).SubItems(6)
          ActiveSheet.Range("G" & i) = Trest.ListView1.ListItems.item(i - 1).SubItems(7)
         Next i

        ' It depends on the number of columns that I want to Export
        Range("A1:G1").Select

        Range(Selection, Selection.End(xlDown)).Select

        Selection.WrapText = False

        ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).name = "Tabulka"
        Range("Tabulka[#All]").Select
        ActiveSheet.ListObjects("Tabulka").TableStyle = "TableStyleMedium18"

        Columns("A:G").EntireColumn.AutoFit

    Case "Dedictvi"
         ActiveSheet.Range("A1") = Dedictvi.ListView1.ColumnHeaders(1)
         ActiveSheet.Range("B1") = Dedictvi.ListView1.ColumnHeaders(3)
         ActiveSheet.Range("C1") = Dedictvi.ListView1.ColumnHeaders(4)
         ActiveSheet.Range("D1") = Dedictvi.ListView1.ColumnHeaders(5)
         ActiveSheet.Range("E1") = Dedictvi.ListView1.ColumnHeaders(6)
         ActiveSheet.Range("F1") = Dedictvi.ListView1.ColumnHeaders(7)
         ActiveSheet.Range("G1") = Dedictvi.ListView1.ColumnHeaders(8)

         For i = 2 To Dedictvi.ListView1.ListItems.count + 1
          ActiveSheet.Range("A" & i) = Dedictvi.ListView1.ListItems.item(i - 1)
          ActiveSheet.Range("B" & i) = Dedictvi.ListView1.ListItems.item(i - 1).SubItems(2)
          ActiveSheet.Range("C" & i) = Dedictvi.ListView1.ListItems.item(i - 1).SubItems(3)
          ActiveSheet.Range("D" & i) = Dedictvi.ListView1.ListItems.item(i - 1).SubItems(4)
          ActiveSheet.Range("E" & i) = Dedictvi.ListView1.ListItems.item(i - 1).SubItems(5)
          ActiveSheet.Range("F" & i) = Dedictvi.ListView1.ListItems.item(i - 1).SubItems(6)
          ActiveSheet.Range("G" & i) = Dedictvi.ListView1.ListItems.item(i - 1).SubItems(7)
         Next i

        ' It depends on the number of columns that I want to Export
        Range("A1:G1").Select

        Range(Selection, Selection.End(xlDown)).Select

        Selection.WrapText = False

        ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).name = "Tabulka"
        Range("Tabulka[#All]").Select
        ActiveSheet.ListObjects("Tabulka").TableStyle = "TableStyleMedium18"

        Columns("A:G").EntireColumn.AutoFit

    Case "Insol"
         ActiveSheet.Range("A1") = Insol.ListView1.ColumnHeaders(1)
         ActiveSheet.Range("B1") = Insol.ListView1.ColumnHeaders(3)
         ActiveSheet.Range("C1") = Insol.ListView1.ColumnHeaders(4)
         ActiveSheet.Range("D1") = Insol.ListView1.ColumnHeaders(5)
         ActiveSheet.Range("E1") = Insol.ListView1.ColumnHeaders(6)
         ActiveSheet.Range("F1") = Insol.ListView1.ColumnHeaders(7)
         ActiveSheet.Range("G1") = Insol.ListView1.ColumnHeaders(8)

         For i = 2 To Insol.ListView1.ListItems.count + 1
          ActiveSheet.Range("A" & i) = Insol.ListView1.ListItems.item(i - 1)
          ActiveSheet.Range("B" & i) = Insol.ListView1.ListItems.item(i - 1).SubItems(2)
          ActiveSheet.Range("C" & i) = Insol.ListView1.ListItems.item(i - 1).SubItems(3)
          ActiveSheet.Range("D" & i) = Insol.ListView1.ListItems.item(i - 1).SubItems(4)
          ActiveSheet.Range("E" & i) = Insol.ListView1.ListItems.item(i - 1).SubItems(5)
          ActiveSheet.Range("F" & i) = Insol.ListView1.ListItems.item(i - 1).SubItems(6)
          ActiveSheet.Range("G" & i) = Insol.ListView1.ListItems.item(i - 1).SubItems(7)
         Next i

        ' It depends on the number of columns that I want to Export
        Range("A1:G1").Select

        Range(Selection, Selection.End(xlDown)).Select

        Selection.WrapText = False

        ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).name = "Tabulka"
        Range("Tabulka[#All]").Select
        ActiveSheet.ListObjects("Tabulka").TableStyle = "TableStyleMedium18"

        Columns("A:G").EntireColumn.AutoFit

End Select

End Sub



Public Function Isinteger2(value As Double) As String
    'vrati navic ,-, pokud je to cele cislo
    'function to check, whether value is integer or not
    
    If value - Round(value) = 0 Then
        Isinteger2 = "" & format(value, "### ### ##0") & ",-"
    Else
        Isinteger2 = "" & format(value, "### ### ##0.00") & ""
    End If
    
    Isinteger2 = Trim(Isinteger2)

End Function

