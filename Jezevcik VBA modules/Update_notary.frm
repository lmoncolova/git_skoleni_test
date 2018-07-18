VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Update_notary 
   Caption         =   "Notáøi"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12615
   OleObjectBlob   =   "Update_notary.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Update_notary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






'RISK-468


Option Explicit
Private BlnNovyNotar As Boolean
Private skp_notary As Long
Private strHledanyVyraz As String
' k 4.12. neni v tomhle modulu naimplementovan FLAG_DELETED, zaznamy se pouze ignoruji v hlavnim listview.


Public Property Let ZadaniNovyNotar(ByRef NovyNotar As Boolean)
    BlnNovyNotar = NovyNotar
    'Tuto promìnnou naplním podle toho, odkud se klikne a podle toho zmìním vzhled formuláøe
End Property


'Private Sub TextBox_Search_Enter()
'    Call CommandButton_Search_Click
'End Sub

Private Sub TextBox_Search_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call CommandButton_Search_Click
End Sub



Private Sub CommandButton_Search_Click()
' vyhledavat bude ve vsech polich
Dim sql_notar As String

strHledanyVyraz = Trim(TextBox_Search)
'sql = sql & " and UPPER(convert(to_char(name), 'US7ASCII')) like UPPER(CONVERT('%" & Trim(TextBox1.text) & "%', 'US7ASCII')) " 'FA vyhladavanie pri zadani CUID a mena do nespraveho pola, ignoruje diakritiku
sql_notar = "SELECT a.skp_notary, a.name, a.address, a.phone_number, a.id_ds, a.flag_deleted FROM " & schema & ".notary a WHERE a.flag_deleted = 'N' and 1 = 1"

If Len(strHledanyVyraz) > 0 Then
    
    sql_notar = sql_notar & " AND ( UPPER(convert(to_char(name), 'US7ASCII')) like UPPER(CONVERT('%" & strHledanyVyraz & "%', 'US7ASCII')) " & _
                                    "OR UPPER(convert(to_char(address), 'US7ASCII')) like UPPER(CONVERT('%" & strHledanyVyraz & "%', 'US7ASCII')) " & _
                                    "OR UPPER(ID_DS) like UPPER('%" & strHledanyVyraz & "%')" & _
                                    "OR UPPER(phone_number) like '%" & strHledanyVyraz & "%'" & _
                                    "OR UPPER(skp_notary) like '%" & strHledanyVyraz & "%'" & _
                                ")"
    sql_notar = sql_notar & " AND flag_deleted != 'Y'"

End If

Call NaplnListview_Potvrzeni(sql_notar)

End Sub

Private Sub VytvorListview_Potvrzeni()
Dim objColHeads As ColumnHeaders

Set objColHeads = Me.ListView_Potvrzeni.ColumnHeaders

objColHeads.Add , , "SKP_NOTARY", 1
objColHeads.Add , , "Potvrzení", 70
objColHeads.Add , , "Jméno", 120
objColHeads.Add , , "Adresa", 200
objColHeads.Add , , "Telefon", 70
objColHeads.Add , , "ID_DS", 50

Call NastavKliceListview(Me.ListView_Potvrzeni)

End Sub


Private Sub NaplnListview_Potvrzeni(ByVal sql_notar As String)
Dim rst_notar As ADODB.Recordset
Dim i As Integer
Dim itmX As ListItem
Dim LV As ListView
Dim keys As Dictionary

Set LV = Me.ListView_Potvrzeni
Set keys = KliceKListview(LV)

LV.ListItems.Clear

Set rst_notar = New ADODB.Recordset
rst_notar.Open sql_notar, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
    If Not rst_notar.EOF Then
        
        'a.skp_notary, a.name, a.address, a.phone_number, a.id_ds, a.flag_deleted
        rst_notar.MoveFirst
        For i = 1 To rst_notar.RecordCount
            Set itmX = LV.ListItems.Add()
            itmX.text = rst_notar![skp_notary]
            itmX.SubItems(keys("Jméno")) = "" & rst_notar![name] & ""
            itmX.SubItems(keys("Adresa")) = "" & rst_notar![address] & ""
            itmX.SubItems(keys("Telefon")) = "" & rst_notar![PHONE_NUMBER] & ""
            itmX.SubItems(keys("ID_DS")) = "" & rst_notar![id_ds] & ""
            itmX.SubItems(keys("Potvrzení")) = "Nepotvrzeno"
            
            rst_notar.MoveNext
        Next i
        
    End If
    
rst_notar.Close
Me.Repaint

End Sub

Private Sub CommandButton_Uloz_Click()
' SUB: overi se ze nejde o duplicitu
'      vlozi se zaznam do ciselniku notaru
'      pak zmizi cast userformu pro vlozeni noveho notare a je potreba potvrdit shodu

Dim strJmeno As String
Dim strTelefon As String
Dim strAdresa As String
Dim strID_DS As String
Dim sql_notar As String


'''''   overenie prav uzivatela na insert
If Not prava("select * from USER_TAB_PRIVS where owner = '" & UCase(schema) & "' and table_name = 'NOTARY' and PRIVILEGE = 'INSERT'") Then
    MsgBox "Na vkládaní notáøù nemáš práva."
    Exit Sub
End If

'''''   nacitame hodnoty textboxov do premennych
strJmeno = "'" & Trim(TextBox_Name.text) & "'"
strTelefon = "'" & Trim(TextBox_Phone.text) & "'"
strAdresa = "'" & Trim(TextBox_Address_Full.text) & "'"
strID_DS = "'" & Trim(TextBox_ID_DS.text) & "'"

If strJmeno = "''" Then strJmeno = "NULL"
If strTelefon = "''" Then strTelefon = "NULL"
If strAdresa = "''" Then strAdresa = "NULL"
If strID_DS = "''" Then strID_DS = "NULL"

'''''    kontrola hodnot, pocet nevyplnenych hodnot, vyhladanie podobneho notara
'
' TO BE DONE XXXX
'
'
'''''

' If MsgBox("Opravdu si pøeješ vložit nového notáøe?", vbYesNo) = vbYes Then
    
    sql_notar = "INSERT INTO " & schema & ".notary (skp_notary, name, address, phone_number, id_ds, flag_deleted) " & _
                "VALUES (" & schema & ".notary_seq.nextval, " & strJmeno & ", " & strAdresa & ", " & strTelefon & ", " & strID_DS & ", 'N')"
    
    execute (sql_notar)
    
    BlnNovyNotar = False
    Call NastavVyskuUserformu
    TextBox_Search.text = Trim(TextBox_Name.text)
    Call CommandButton_Search_Click
    
    MsgBox "Notáø pøidán do seznamu. Pro pøirazení k zemøelému jej ještì potvrï ve vyhledávacím menu."
    
' End If

End Sub


Private Sub CommandButton_Smazat_Click()
Call promazHodnoty
End Sub


Private Sub CommandButton_Vlozit_Notare_Click()

BlnNovyNotar = True
Call UserForm_Activate
CommandButton_Vlozit_Notare.Enabled = False

End Sub

Private Sub ListView_Potvrzeni_DblClick()

'Dovolím klikání, jenom pokud tam je na co klikat
If ListView_Potvrzeni.ListItems.count > 0 Then

        If ListvZiskejHodnotu(ListView_Potvrzeni, "Potvrzení") = "Nepotvrzeno" Then
           Call ListvVlozHodnotu(ListView_Potvrzeni, "Potvrzení", "Potvrzeno")
        
        ElseIf ListvZiskejHodnotu(ListView_Potvrzeni, "Potvrzení") = "Potvrzeno" Then
           Call ListvVlozHodnotu(ListView_Potvrzeni, "Potvrzení", "Zamítnuto")
        
        ElseIf ListvZiskejHodnotu(ListView_Potvrzeni, "Potvrzení") = "Zamítnuto" Then
           Call ListvVlozHodnotu(ListView_Potvrzeni, "Potvrzení", "Potvrzeno")
        
        End If

End If

End Sub

Private Sub CommandButton_Potvrzeni_Click()
Dim i As Integer, Intcounter As Integer, IntPozice As Integer
Dim sql As String, sqlMaxDedNote As String
Dim BlnInheritanceExists As Boolean, BlnNotaryExists As Boolean
Dim LngSkpNotary As Long, StrNameNewNotary As String
Dim skp_ded_note As Long

'Iniciálnì
BlnInheritanceExists = False
BlnNotaryExists = False

If ListView_Potvrzeni.ListItems.count > 0 Then
    
    Intcounter = 0
    
    For i = 1 To ListView_Potvrzeni.ListItems.count
    'RISK-876
        If ListvZiskejHodnotu(ListView_Potvrzeni, "Potvrzení", i) = "Potvrzeno" Then
            Intcounter = Intcounter + 1
            IntPozice = i ' Uložím si pozici potvrzeného øádku
        End If
    Next i

    'Musí být potvrzený právì jeden, jinak neuložím
    If Intcounter <> 1 Then
        MsgBox "Pro uložení je tøeba potvrdit právì jeden subjekt. (Subjekt se potvrdzuje dvojitým kliknutím na øádek.)", vbExclamation, "Nedostateèné potvrzení hodnot"
        Exit Sub
    End If
    

    
    'Rozdìlím na dvì vìtve - pokud už notáø ke klientovi existuje, tak se zeptám, jestli chce zmìnit notáøe. Pokud ne, tak vloží nového
    
    If dedici.lngSkp_Inheritance <> vbNullString And dedici.lngSkp_Inheritance > 0 Then
        BlnInheritanceExists = True
        
        If dedici.objNotar.lngSkp_notary > 0 Then
            BlnNotaryExists = True
        End If
    
    End If
    
    
    LngSkpNotary = ListvZiskejHodnotu(ListView_Potvrzeni, "SKP_NOTARY", IntPozice)
    StrNameNewNotary = ListvZiskejHodnotu(ListView_Potvrzeni, "Jméno", IntPozice)
    
    
    'Vazba notáøe na zemøelého již existuje - update
    If BlnNotaryExists Then
        
        If MsgBox("Notáø na tomto øízení již existuje, pøeješ si zmìnit pøiøazení notáøe?", vbExclamation + vbYesNo, "Výmìna notáøe") = vbYes Then
                
            Call TransBegin
                
                sql = "update " & schema & ".inheritance set " & _
                     " skp_notary = " & LngSkpNotary & _
                     " where SKP_INHERITANCE = " & dedici.lngSkp_Inheritance
                    
                execute (sql)
                
                
                'vlozit poznamku o zmene notare
        
                sqlMaxDedNote = "select max(skp_Ded_note) as max_skp_Ded_note from " & schema & ".Ded_NOTE"
                
                Set rst1 = New ADODB.Recordset
                rst1.Open sqlMaxDedNote, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
                    
                    If Not IsNull(rst1![max_skp_Ded_note]) Then
                        skp_ded_note = rst1![max_skp_Ded_note] + 1
                    Else
                        skp_ded_note = 1
                    End If
                
                rst1.Close
                
                sql = "insert into " & schema & ".ded_note (SKP_DED_NOTE, SKP_DED_NOTE_TYPE, ID_PERSON, TEXT_NOTE, DTIME_NOTE_CREATED, EMPLOYEE)" & _
                    " values (" & skp_ded_note & ", 3, " & c.id_person & _
                ", 'Zmìna notáøe: døíve " & dedici.objNotar.strName & " novì " & StrNameNewNotary & "', sysdate, '" & UCase(Main.username) & "')"
                
                execute (sql)
            
            Call TransCommit
            
        End If
    
    Else  'Notáø ještì dìdickému øízení není pøiøazen - pokud tam už øádek je ale skp_notary 0 tak update - pokud tam ještì není tak insert nového øádku
    
        'If MsgBox("Opravdu si pøeješ pøiøadit notáøe?", vbExclamation + vbYesNo, "Pøiøazení notáøe") = vbYes Then
         
            If BlnInheritanceExists Then
            'pokud tam už øádek je ale skp_notary 0 tak update
        
                sql = "update " & schema & ".inheritance set " & _
                    " skp_notary = " & LngSkpNotary & _
                    " where SKP_INHERITANCE = " & dedici.lngSkp_Inheritance
                
                execute (sql)
            
            Else 'Øádek tam ještì vùbec není
            
                sql = "INSERT INTO " & schema & ".inheritance (skp_inheritance,  id_person_dead,  skp_notary, flag_deleted, employee_inserted, dtime_inserted) " & _
                      "VALUES (" & schema & ".inheritance_seq.nextval, " & id_person & ", " & LngSkpNotary & ", 'N', '" & Main.username & "', SYSDATE)"
        
                execute (sql)
        
            End If
            
        'End If
    
    End If
    
    
    
    
    
    
End If


Call CommandButton_Navrat_Click


End Sub



Private Sub UserForm_Initialize()

Call NastavVyskuUserformu
Call promazHodnoty
CommandButton_Vlozit_Notare.Enabled = True

Call VytvorListview_Potvrzeni ' pouze v userform_initialize!
Call CommandButton_Search_Click ' naplni listview

End Sub

Private Sub UserForm_Activate()

Call NastavVyskuUserformu
Call promazHodnoty

CommandButton_Vlozit_Notare.Enabled = True

Call CommandButton_Search_Click ' naplni listview

End Sub

Private Sub NastavVyskuUserformu()

If BlnNovyNotar Then
    Me.Height = 406
Else
    Me.Height = 288
End If
End Sub

Private Sub promazHodnoty()
Dim cCont As control

For Each cCont In Me.Controls
    If TypeName(cCont) = "TextBox" Then
        cCont = ""
    End If
Next cCont
End Sub

Private Sub CommandButton_Navrat_Click()
'Aby se to pøi pøíštím naètení nerozrolovalo
BlnNovyNotar = False
Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim ans

'Aby se to pøi pøíštím naètení nerozrolovalo
BlnNovyNotar = False

If CloseMode = vbFormControlMenu Then
    Cancel = True
    export_chosen_option = 0
    Call CommandButton_Navrat_Click
End If

End Sub

''Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
''Call CommandButton_Navrat_Click
''End Sub


