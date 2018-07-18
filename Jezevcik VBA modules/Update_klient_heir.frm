VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Update_klient_heir 
   Caption         =   "Jezevèík - Dìdické øízení"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12465
   OleObjectBlob   =   "Update_klient_heir.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Update_klient_heir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






'RISK-468

Option Explicit
Private BlnNovyDedic As Boolean
Private id_person_Dedic As Long
Private skp_person_Dedic As Long
Private BlnKlientHledan As Boolean



Public Property Let ZadaniNovyDedic(ByRef NovyDedic As Boolean)
BlnNovyDedic = NovyDedic
'Tuto promìnnou naplním podle toho, odkud se klikne a podle toho zmìním vzhled formuláøe
       
End Property



Private Sub cmdNajdiPSC_trvala_Click()
' Sub otevre Internet Explorer na adrese (odtam uz si snadno skopiruji PSC)
Dim strAdress As String
    
    strAdress = TextBox_Permanent_Address_Street.text & ", " & TextBox_Permanent_Address_Num.text & ", " & TextBox_Permanent_Address_Town.text
    Call NajdiAdresuMapyCZ(strAdress)

End Sub

Private Sub cmdNajdiPSC_prechodni_Click()
' Sub otevre Internet Explorer na adrese (odtam uz si snadno skopiruji PSC)
Dim strAdress As String
    
    strAdress = TextBox_Reports_Address_Street.text & ", " & TextBox_Reports_Address_Num.text & ", " & TextBox_Reports_Address_Town.text
    Call NajdiAdresuMapyCZ(strAdress)

End Sub

Private Sub CommandButton_Hledej_Click()
Dim sql As String, sql_part As String, id_cuid As Long
Dim clmx As ColumnHeader
Dim itmX As ListItem
Dim keys As Dictionary
Dim CB As control

If Not BlnNovyDedic Then Exit Sub


'iniciálnì
id_person_Dedic = 0
skp_person_Dedic = 0

'Inicální kontroly na správné zadání


If TextBox_RC.value = vbNullString And TextBox_CUID.value = vbNullString Then
    If TextBox_Name.value = vbNullString Or TextBox_Surname.value = vbNullString Or TextBox_Date_Birth.value = vbNullString Then
        MsgBox "Nezadal jsi ani rodné èíslo nebo CUID, ani plnou kombinaci jména, pøíjmení a data narození. Zkus to znovu.", vbExclamation, "Chybné zadání hodnot"
        Exit Sub
    End If
End If


'CUID

If Not IsNumeric(TextBox_CUID.value) And TextBox_CUID.value <> vbNullString Then
         MsgBox "Zkus to prosím znovu, zadej do CUIDu pouze èíselné hodnoty", vbExclamation, "Chybné zadání hodnot"
         Exit Sub
End If

'RC
If Not IsNumeric(TextBox_RC.value) And TextBox_RC.value <> vbNullString Then
         MsgBox "Zkus to prosím znovu, zadej do rodného èísla pouze èíselné hodnoty", vbExclamation, "Chybné zadání hodnot"
         Exit Sub
ElseIf Not (Len(TextBox_RC.value) = 9 Or Len(TextBox_RC.value) = 10) And TextBox_RC.value <> vbNullString Then
         MsgBox "Zkus to prosím znovu, rodné èíslo musí obsahovat 9 nebo 10 znakù", vbExclamation, "Chybné zadání hodnot"
         Exit Sub
End If

'Date Birth
If TextBox_RC.value = vbNullString And TextBox_CUID.value = vbNullString And Not IsDate(TextBox_Date_Birth.value) Then
         MsgBox "Zkus to prosím znovu, datum narození nebylo zadáno ve správném formátu", vbExclamation, "Chybné zadání hodnot"
         Exit Sub
End If


'Podle toho co již bylo vyplnìno, tak se bude hledat klient. Pokud bude rodné èíslo, tak hledám dle RÈ. Pokud bude jméno pøíjmení a datum narození, tak podle toho." & _
'Mùžou nastat tøi situace: 1) Buï to nenajde nikoho a potom se zahlásí, že klient nenalezen. 2) Najde právì jednoho, pak se ihned pøedvyplní data. 3) najde více, pak bude muset uživatel vybrat, který to je - nebo také žádný")

'RISK-985
sql_part = "k.SKP_PERSON, p.ID_PERSON, k.ID_CUID, k.NAME_FIRST, k.NAME_LAST, k.DATE_BIRTH, decode(k.TEXT_BIRTH_NUMBER,'XNA',NULL,k.TEXT_BIRTH_NUMBER) as TEXT_BIRTH_NUMBER, k.NAME_REPORTS_ADDRESS_STREET, k.TEXT_REPORTS_ADDRESS_NUM, k.CODE_REPORTS_ADDRESS_ZIP, k.NAME_REPORTS_ADDRESS_TOWN, k.NAME_REPORTS_ADDRESS_COUNTRY, k.NAME_PERMANENT_ADDRESS_STREET, k.TEXT_PERMANENT_ADDRESS_NUM, k.CODE_PERMANENT_ADDRESS_ZIP, k.NAME_PERMANENT_ADDRESS_TOWN, k.NAME_PERMANENT_ADDRESS_COUNTRY, k.PRIMARY_PHONE, k.PRIMARY_EMAIL, k.CODE_GENDER"

If TextBox_CUID.value <> vbNullString Then
    sql = "select " & sql_part & " from " & schema & ".sonik_ab_klient_dtb k " & _
                                 " left join " & schema & ".person p on k.skp_person = p.skp_person " & _
         " where k.ID_CUID = '" & TextBox_CUID.value & "'"
ElseIf TextBox_RC.value <> vbNullString Then
    sql = "select " & sql_part & " from " & schema & ".sonik_ab_klient_dtb k " & _
                                 " left join " & schema & ".person p on k.skp_person = p.skp_person " & _
         " where k.TEXT_BIRTH_NUMBER = '" & TextBox_RC.value & "'"
Else
    sql = "select " & sql_part & " from " & schema & ".sonik_ab_klient_dtb k " & _
                                 " left join " & schema & ".person p on k.skp_person = p.skp_person " & _
         " where replace(lower(convert(NAME_FIRST,'US7ASCII')),' ','') like replace(lower(convert('%" & TextBox_Name.value & "%','US7ASCII')),' ','') and replace(lower(convert(NAME_LAST,'US7ASCII')),' ','') like replace(lower(convert('%" & TextBox_Surname.value & "%','US7ASCII')),' ','') and DATE_BIRTH = to_date('" & TextBox_Date_Birth.value & "','DD.MM.YYYY')"
End If


'Samotné hledání
   
BlnKlientHledan = True
   
Set rst1 = New ADODB.Recordset
rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
    If Not rst1.EOF Then
    
        'V pøípadì pøedvyplnìných datech již nesmím povolit editaci
        
        For Each CB In Me.Controls
            If (Left(CB.name, 7) = "TextBox" Or Left(CB.name, 8) = "ComboBox") And CB.name <> "TextBox_Receivable" And CB.name <> "TextBox_Payment_Ref_Number" And CB.name <> "CheckBox_Address_Match" Then
                CB.Enabled = False
            End If
        Next CB
        
        CheckBox_Address_Match.Visible = False
    
        If rst1.RecordCount = 1 Then
            'Našel jsem právì jednoho, pøedvyplním data
            
            
            TextBox_RC.value = "" & rst1![TEXT_BIRTH_NUMBER] & ""
            TextBox_CUID.value = "" & rst1![id_cuid] & ""
            TextBox_Name.value = "" & rst1![Name_first] & ""
            TextBox_Surname.value = "" & rst1![Name_last] & ""
            TextBox_Date_Birth.value = "" & rst1![Date_Birth] & ""
            ComboBox_Gender.value = "" & rst1![CODE_GENDER] & ""
            TextBox_Phone.value = "" & rst1![PRIMARY_PHONE] & ""
            TextBox_Email.value = "" & rst1![PRIMARY_EMAIL] & ""
            TextBox_Permanent_Address_Street.value = "" & rst1![name_permanent_address_street] & ""
            TextBox_Permanent_Address_Num.value = "" & rst1![text_permanent_address_num] & ""
            TextBox_Permanent_Address_Town.value = "" & rst1![name_permanent_address_town] & ""
            TextBox_Permanent_Address_ZIP.value = "" & rst1![code_permanent_address_zip] & ""
            TextBox_Reports_Address_Street.value = "" & rst1![NAME_REPORTS_ADDRESS_STREET] & ""
            TextBox_Reports_Address_Num.value = "" & rst1![TEXT_REPORTS_ADDRESS_NUM] & ""
            TextBox_Reports_Address_Town.value = "" & rst1![NAME_REPORTS_ADDRESS_TOWN] & ""
            TextBox_Reports_Address_ZIP.value = "" & rst1![CODE_REPORTS_ADDRESS_ZIP] & ""
            
            
            If rst1![NAME_PERMANENT_ADDRESS_COUNTRY] = "Czech Republic" Or rst1![NAME_PERMANENT_ADDRESS_COUNTRY] = "Slovakia" Then
                ComboBox_Permanent_Address_Country = "" & rst1![NAME_PERMANENT_ADDRESS_COUNTRY] & ""
            Else: ComboBox_Permanent_Address_Country = "Jiný"
                  TextBox_Permanent_Address_Country = "" & rst1![NAME_PERMANENT_ADDRESS_COUNTRY] & ""
            End If
            
            If rst1![NAME_REPORTS_ADDRESS_COUNTRY] = "Czech Republic" Or rst1![NAME_REPORTS_ADDRESS_COUNTRY] = "Slovakia" Then
                ComboBox_Reports_Address_Country = "" & rst1![NAME_REPORTS_ADDRESS_COUNTRY] & ""
            Else: ComboBox_Reports_Address_Country = "Jiný"
                  TextBox_Reports_Address_Country = "" & rst1![NAME_REPORTS_ADDRESS_COUNTRY] & ""
            End If
            
            If Not IsNull(rst1![skp_person]) Then skp_person_Dedic = rst1![skp_person] ' RISK-985
            If Not IsNull(rst1![id_person]) Then id_person_Dedic = rst1![id_person] ' RISK-985
            
            
        Else
            'Našel jsem více, uživatel musí vybrat, který to je - zobrazím do uživateli do listview
            
            BlnKlientHledan = False 'Na True se pøepíše až pøi potvrzení subjektu
            
            If Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.count = 0 Then
    
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, , "SKP_PERSON", 1)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, , "Potvrzení", 70)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, "ID_CUID", "CUID", 50)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, "NAME_FIRST", "Jméno", 50)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, "NAME_LAST", "Pøíjmení", 50)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, "TEXT_BIRTH_NUMBER", "Rodné èíslo", 50)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, "DATE_BIRTH", "Datum nar.", 50)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, "PRIMARY_PHONE", "Telefonní èíslo", 50)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, "PRIMARY_EMAIL", "E-mail", 50)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, "PERMANENT_ADDRESS", "Trvalá adresa", 50)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, "REPORTS_ADDRESS", "Kontaktní adresa", 50)
                'Pomocné sloupce pro uložení
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, , "CODE_GENDER", 0)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, , "NAME_PERMANENT_ADDRESS_STREET", 0)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, , "TEXT_PERMANENT_ADDRESS_NUM", 0)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, , "NAME_PERMANENT_ADDRESS_TOWN", 0)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, , "CODE_PERMANENT_ADDRESS_ZIP", 0)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, , "NAME_PERMANENT_ADDRESS_COUNTRY", 0)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, , "NAME_REPORTS_ADDRESS_STREET", 0)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, , "TEXT_REPORTS_ADDRESS_NUM", 0)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, , "NAME_REPORTS_ADDRESS_TOWN", 0)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, , "CODE_REPORTS_ADDRESS_ZIP", 0)
                Set clmx = Update_klient_heir.ListView_Potvrzeni.ColumnHeaders.Add(, , "NAME_REPORTS_ADDRESS_COUNTRY", 0)
                
                Call NastavKliceListview(ListView_Potvrzeni)
            
            End If
            
            
            Update_klient_heir.ListView_Potvrzeni.ListItems.Clear
            Set keys = KliceKListview(Update_klient_heir.ListView_Potvrzeni)
            
            rst1.MoveFirst
            For i = 1 To rst1.RecordCount
                Set itmX = Update_klient_heir.ListView_Potvrzeni.ListItems.Add()
                itmX.text = SKP_Subject 'Vezmu pøímo z excelu ' ' RISK-985 odkomentovat po hromadnom REPLACE
                itmX.SubItems(keys("ID_CUID")) = "" & rst1![id_cuid] & ""
                itmX.SubItems(keys("NAME_FIRST")) = "" & rst1![Name_first] & ""
                itmX.SubItems(keys("NAME_LAST")) = "" & rst1![Name_last] & ""
                itmX.SubItems(keys("TEXT_BIRTH_NUMBER")) = "" & rst1![TEXT_BIRTH_NUMBER] & ""
                itmX.SubItems(keys("DATE_BIRTH")) = "" & rst1![Date_Birth] & ""
                itmX.SubItems(keys("PRIMARY_PHONE")) = "" & rst1![PRIMARY_PHONE] & ""
                itmX.SubItems(keys("PRIMARY_EMAIL")) = "" & rst1![PRIMARY_EMAIL] & ""
                itmX.SubItems(keys("PERMANENT_ADDRESS")) = "" & rst1![name_permanent_address_street] & "" & rst1![text_permanent_address_num] & "" & rst1![name_permanent_address_town] & " " & rst1![code_permanent_address_zip] & ""
                itmX.SubItems(keys("REPORTS_ADDRESS")) = "" & rst1![NAME_REPORTS_ADDRESS_STREET] & "" & rst1![TEXT_REPORTS_ADDRESS_NUM] & " " & rst1![NAME_REPORTS_ADDRESS_TOWN] & " " & rst1![CODE_REPORTS_ADDRESS_ZIP] & ""
                
                itmX.SubItems(keys("Potvrzení")) = "Nepotvrzeno"
                'Pomocné sloupce, nezobrazuji
                itmX.SubItems(keys("CODE_GENDER")) = "" & rst1![CODE_GENDER] & ""
                itmX.SubItems(keys("NAME_PERMANENT_ADDRESS_STREET")) = "" & rst1![name_permanent_address_street] & ""
                itmX.SubItems(keys("TEXT_PERMANENT_ADDRESS_NUM")) = "" & rst1![text_permanent_address_num] & ""
                itmX.SubItems(keys("NAME_PERMANENT_ADDRESS_TOWN")) = "" & rst1![name_permanent_address_town] & ""
                itmX.SubItems(keys("CODE_PERMANENT_ADDRESS_ZIP")) = "" & rst1![code_permanent_address_zip] & ""
                itmX.SubItems(keys("NAME_PERMANENT_ADDRESS_COUNTRY")) = "" & rst1![NAME_PERMANENT_ADDRESS_COUNTRY] & ""
                itmX.SubItems(keys("NAME_REPORTS_ADDRESS_STREET")) = "" & rst1![NAME_REPORTS_ADDRESS_STREET] & ""
                itmX.SubItems(keys("TEXT_REPORTS_ADDRESS_NUM")) = "" & rst1![TEXT_REPORTS_ADDRESS_NUM] & ""
                itmX.SubItems(keys("NAME_REPORTS_ADDRESS_TOWN")) = "" & rst1![NAME_REPORTS_ADDRESS_TOWN] & ""
                itmX.SubItems(keys("CODE_REPORTS_ADDRESS_ZIP")) = "" & rst1![CODE_REPORTS_ADDRESS_ZIP] & ""
                itmX.SubItems(keys("NAME_REPORTS_ADDRESS_COUNTRY")) = "" & rst1![NAME_REPORTS_ADDRESS_COUNTRY] & ""
                rst1.MoveNext
            
                Me.Height = 520.5 'Zobrazím Listview
            
            Next i
        End If
        
    
        
        
        
    Else
        
        MsgBox "Hledaný subjekt není klientem AB nebo nebyl podle zadaných paramterù dohledán", vbInformation, "Klient nedohledán"
        ' RISK-985 - tu naplnit id_person_dedic (pridavame noveho neklienta) (sekvence.nextval)
    
    End If

rst1.Close




End Sub


Private Sub ListView_Potvrzeni_DblClick()

'Dovolím klikání, jenom pokud tam je na co klikat, jinak to spadne
If ListView_Potvrzeni.ListItems.count > 0 Then


        'RISK-876 bohužel pro update hodnoty musím nechat tak jak je
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
'Prvnì vybere pouze jednoho klienta, toho potvrdí a když dá potvrdit tak listview se schová a data se pøedvyplní
'V pøípadì pøedvyplnìných datech již nesmím povolit editaci




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
        MsgBox "Pro uložení je tøeba potvrdit právì jeden subjekt", vbExclamation, "Nedostateèné potvrzení hodnot"
        Exit Sub
    End If


        'Uživatel potvrdil právì jednoho, pøedvyplním data
        TextBox_RC.value = ListvZiskejHodnotu(ListView_Potvrzeni, "TEXT_BIRTH_NUMBER", IntPozice)
        TextBox_CUID.value = ListvZiskejHodnotu(ListView_Potvrzeni, "ID_CUID", IntPozice)
        TextBox_Name.value = ListvZiskejHodnotu(ListView_Potvrzeni, "NAME_FIRST", IntPozice)
        TextBox_Surname.value = ListvZiskejHodnotu(ListView_Potvrzeni, "NAME_LAST", IntPozice)
        TextBox_Date_Birth.value = ListvZiskejHodnotu(ListView_Potvrzeni, "DATE_BIRTH", IntPozice)
        ComboBox_Gender.value = ListvZiskejHodnotu(ListView_Potvrzeni, "CODE_GENDER", IntPozice)
        TextBox_Phone.value = ListvZiskejHodnotu(ListView_Potvrzeni, "PRIMARY_PHONE", IntPozice)
        TextBox_Email.value = ListvZiskejHodnotu(ListView_Potvrzeni, "PRIMARY_EMAIL", IntPozice)
        TextBox_Permanent_Address_Street.value = ListvZiskejHodnotu(ListView_Potvrzeni, "NAME_PERMANENT_ADDRESS_STREET", IntPozice)
        TextBox_Permanent_Address_Num.value = ListvZiskejHodnotu(ListView_Potvrzeni, "TEXT_PERMANENT_ADDRESS_NUM", IntPozice)
        TextBox_Permanent_Address_Town.value = ListvZiskejHodnotu(ListView_Potvrzeni, "NAME_PERMANENT_ADDRESS_TOWN", IntPozice)
        TextBox_Permanent_Address_ZIP.value = ListvZiskejHodnotu(ListView_Potvrzeni, "CODE_PERMANENT_ADDRESS_ZIP", IntPozice)
        TextBox_Reports_Address_Street.value = ListvZiskejHodnotu(ListView_Potvrzeni, "NAME_REPORTS_ADDRESS_STREET", IntPozice)
        TextBox_Reports_Address_Num.value = ListvZiskejHodnotu(ListView_Potvrzeni, "TEXT_REPORTS_ADDRESS_NUM", IntPozice)
        TextBox_Reports_Address_Town.value = ListvZiskejHodnotu(ListView_Potvrzeni, "NAME_REPORTS_ADDRESS_TOWN", IntPozice)
        TextBox_Reports_Address_ZIP.value = ListvZiskejHodnotu(ListView_Potvrzeni, "CODE_REPORTS_ADDRESS_ZIP", IntPozice)


        If ListvZiskejHodnotu(ListView_Potvrzeni, "NAME_PERMANENT_ADDRESS_COUNTRY", IntPozice) = "Czech Republic" Or ListvZiskejHodnotu(ListView_Potvrzeni, "NAME_PERMANENT_ADDRESS_COUNTRY", IntPozice) = "Slovakia" Then
            ComboBox_Permanent_Address_Country = ListvZiskejHodnotu(ListView_Potvrzeni, "NAME_PERMANENT_ADDRESS_COUNTRY", IntPozice)
        Else: ComboBox_Permanent_Address_Country = "Jiný"
              TextBox_Permanent_Address_Country = ListvZiskejHodnotu(ListView_Potvrzeni, "NAME_PERMANENT_ADDRESS_COUNTRY", IntPozice)
        End If
        
        If ListvZiskejHodnotu(ListView_Potvrzeni, "NAME_REPORTS_ADDRESS_COUNTRY", IntPozice) = "Czech Republic" Or ListvZiskejHodnotu(ListView_Potvrzeni, "NAME_REPORTS_ADDRESS_COUNTRY", IntPozice) = "Slovakia" Then
            ComboBox_Reports_Address_Country = ListvZiskejHodnotu(ListView_Potvrzeni, "NAME_REPORTS_ADDRESS_COUNTRY", IntPozice)
        Else: ComboBox_Reports_Address_Country = "Jiný"
              TextBox_Reports_Address_Country = ListvZiskejHodnotu(ListView_Potvrzeni, "NAME_REPORTS_ADDRESS_COUNTRY", IntPozice)
        End If

        id_person_Dedic = ListvZiskejHodnotu(ListView_Potvrzeni, "ID_PERSON", IntPozice)

        Me.Height = 364.5 'Skryju Listview
        
        BlnKlientHledan = True
End If


End Sub


Private Sub CommandButton_Uloz_Click()
Dim CB As control
Dim sql As String
Dim BlnClientValidInJez As Boolean
Dim skp_demography_dedic As Long, skp_address_dedic As Long, skp_contact_dedic As Long, cuid_dedic As Long
Dim StrReportsAddressCountry As String, StrPermanentAddressCountry As String
Dim StrZmena As String
Dim sqlMaxHeirNote As String, skp_heir_note As Long
Dim StrAmtReceivable As String
Dim rst19 As ADODB.Recordset

'XXXX Doøešit situaci, když je nìkde XNA a pøedvyplní se to...
'XXXX Doplnit rodná èísla do demography client_ab_DTB ve správných formátech

'Validace

'Kontrola jestli vùbec klienta vyhledávali

If Not BlnKlientHledan Then
        MsgBox "Pøed uložením musíš zkusit klienta vyhledat. Bez toho to nepùjde.", vbExclamation, "Neoprávnìné vniknutí.)"
        Exit Sub
End If


'Chci vždy plnou sadu dat, i když jde o update - tam bude pøedvyplnìno
'Toto se musí ještì doladit - ne všechno musí být vyplnìno - RÈ, adresy, telefon atd...
'Spíše vybrat pouze ty, které se musí kontrolovat
If TextBox_Name.value = vbNullString Or TextBox_Surname.value = vbNullString Or TextBox_Date_Birth.value = vbNullString Or ComboBox_Gender.value = vbNullString Then
        MsgBox "Nebyly vyplnìny základní identifikaèní údaje o klientovi.", vbExclamation, "Chybné zadání hodnot.)"
        Exit Sub
End If


'For Each CB In Me.Controls 'Pokud nìco není vyplnìno
'    If (Left(CB.name, 7) = "TextBox" Or Left(CB.name, 8) = "ComboBox") And CB.name <> "TextBox_Permanent_Address_Country" And CB.name <> "TextBox_Reports_Address_Country" Then
'        If CB.value = vbNullString Then
'            MsgBox "Nìkterá z položek nebyla vyplnìna.", vbExclamation, "Chybné zadání hodnot"
'            Exit Sub
'        End If
'    End If
'Next CB

'Kontrola na CUID
If TextBox_CUID.value <> vbNullString And Not IsNumeric(TextBox_CUID.value) Then
        MsgBox "CUID není vyplnìn správnì.", vbExclamation, "Chybné zadání hodnot"
        Exit Sub
End If

'Kontrola na RÈ
If TextBox_RC.value <> vbNullString And (Not IsNumeric(TextBox_RC.value) Or TextBox_RC.value <= 0 Or Len(TextBox_RC.value) < 9 Or Len(TextBox_RC.value) > 10) Then
        MsgBox "Rodné èíslo není vyplnìno správnì.", vbExclamation, "Chybné zadání hodnot"
        Exit Sub
End If

'Kontrola na datum narození
If Not IsDate(TextBox_Date_Birth.value) Then
        MsgBox "Datum narození není vyplnìno správnì.", vbExclamation, "Chybné zadání hodnot"
        Exit Sub
End If

'Kontrola na PSÈ
If TextBox_Permanent_Address_ZIP.value <> vbNullString And (Not IsNumeric(TextBox_Permanent_Address_ZIP.value) Or TextBox_Permanent_Address_ZIP.value <= 0) Then
        MsgBox "PSÈ trvalé adresy není vyplnìno správnì.", vbExclamation, "Chybné zadání hodnot"
        Exit Sub
End If
If TextBox_Reports_Address_ZIP.value <> vbNullString And (Not IsNumeric(TextBox_Reports_Address_ZIP.value) Or TextBox_Reports_Address_ZIP.value <= 0) Then
        MsgBox "PSÈ kontaktní adresy není vyplnìno správnì.", vbExclamation, "Chybné zadání hodnot"
        Exit Sub
End If

'Kontrola na telefon
If TextBox_Phone.value <> vbNullString And (Not IsNumeric(TextBox_Phone.value) Or Len(TextBox_Phone.value) > 9) Then
        MsgBox "Telefonní èíslo není vyplnìno správnì.", vbExclamation, "Chybné zadání hodnot"
        Exit Sub
End If

'Kontrola na výši pohledávky
If Not IsNumeric(TextBox_Receivable.value) Or TextBox_Receivable.value < 0 Then
        MsgBox "Výše pohledávky není vyplnìna správnì.", vbExclamation, "Chybné zadání hodnot"
        Exit Sub
End If
StrAmtReceivable = Replace(TextBox_Receivable.value, ",", ".") 'Pokud se zadává s desetinnou èárkou, tak by na konci spadlo...

'Kontrola na VS
If Not IsNumeric(TextBox_Payment_Ref_Number.value) Or TextBox_Payment_Ref_Number.value <= 0 Then
        MsgBox "Variabilní symbol není vyplnìn správnì.", vbExclamation, "Chybné zadání hodnot"
        Exit Sub
End If

'Pøipravení si státu do jedné promìnné
If ComboBox_Reports_Address_Country.value = "Jiný" Then
    StrReportsAddressCountry = TextBox_Reports_Address_Country.value
Else
    StrReportsAddressCountry = ComboBox_Reports_Address_Country.value
End If
        
If ComboBox_Permanent_Address_Country.value = "Jiný" Then
    StrPermanentAddressCountry = TextBox_Permanent_Address_Country.value
Else
    StrPermanentAddressCountry = ComboBox_Permanent_Address_Country.value
End If


If BlnNovyDedic Then 'Insert nového dìdice

    If TextBox_CUID.value = vbNullString Then
        cuid_dedic = 0
    Else
        cuid_dedic = TextBox_CUID.value 'Iniciálnì - pokud není klient a hodnota je null, tak pøiøadím pozdìji záporné èíslo skp_person a novou hodnotu id_person_dedic
    End If





    'Musím zkontrolovat všechna pole, že jsou vyplnìna
    For Each CB In Me.Controls
        If Left(CB.name, 8) = "CheckBox" And CB.name <> "CheckBox_Address_Match" Then  'Reálnì by nemìlo nastat - pojistka
            If CB.value = False Then
                MsgBox "Pøi vkládání nového klienta musí být všechna pole zaškrtnutá.", vbExclamation, "Chybné zadání hodnot"
                Exit Sub
            End If
        End If
    Next CB

    'Teï už mám ovìøená data, mùžu provést kontroly na existenci klienta a uložit do JEZ.
    
    'Je už klient v Jezevèíkovi? RISK-985 jdu pøes ID_PERSON
    sql = "select id_person, flag_valid from " & schema & ".PERSON where id_person = " & id_person_Dedic
    BlnClientValidInJez = False
    
    Set rst19 = New ADODB.Recordset
    rst19.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
        If Not rst19.EOF Then 'Osoba je v Jezevèíkovi
            
            'Kontrola, že dìdic už na daném zemøelém existuje - pokud ano, tak nebudu dìlat nic
            For i = 1 To dedici.count
                If dedici.item(i).lngId_person_heir = id_person_Dedic Then
                    MsgBox "Tento dìdic již v tomto øízení figuruje, nelze jej znovu pøidat.", vbExclamation, "Chybné zadání hodnot"
                    Exit Sub
                End If
            Next i
            
            If rst19![flag_valid] = "Y" Then
                BlnClientValidInJez = True
            End If
    
            If BlnClientValidInJez Then
                'Na tomto místì není potøeba žádná další akce, pouze poté pøidání do HEIR
                MsgBox "Osoba už v Jezevèíkovi je.", vbExclamation, "Upozornìní"
            Else
                'Update v tabulce PERSON
                Call TransBegin
                    sql = "update " & schema & ".PERSON " & _
                          " set flag_valid = 'Y', dtime_valid_from = sysdate, dtime_valid_to = date'3000-01-01' " & _
                          " where id_person = " & id_person_Dedic
                    execute (sql)
                Call TransCommit
                
                MsgBox "Osoba už v Jezevèíkovi byla.", vbExclamation, "Upozornìní"
            End If
        
            'rst19.Close
        
        Else ' Osoba neni v Jezevcikovi
       
            'Pokud se zadává neklient - tedy nenalezli žádného klienta - tak musím nastavit skp_person_dedic na hodnotu min - 1 a ostantní skp dìlám do mínusu
            
            If skp_person_Dedic = 0 Then 'Osoba není klient banky - nenalezena v SONIK_AB_KLIENT_DTB
                'RISK-985
                sql = "select min(skp_person) as min_skp_person from " & schema & ".PERSON"
                
                Set rst1 = New ADODB.Recordset
                rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
                    If Not rst1.EOF Then
                        skp_person_Dedic = rst1![min_skp_person]
                        skp_person_Dedic = skp_person_Dedic - 1
                    End If
                rst1.Close
                
                'cuid do minusu do dalších tabulek
                cuid_dedic = skp_person_Dedic
                
            End If
            
                        
            sql = "select " & schema & ".person_seq.nextval as id_person_seq_nextval from dual"
            
            Set rst1 = New ADODB.Recordset
            rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
                If Not rst1.EOF Then
                    id_person_Dedic = rst1![id_person_seq_nextval] 'Nové ID_PERSON pro dìdice
                End If
            rst1.Close
            
            If skp_person_Dedic < 0 Then ' neklient
                skp_address_dedic = skp_person_Dedic
                skp_contact_dedic = skp_person_Dedic
                skp_demography_dedic = skp_person_Dedic
            Else
                skp_address_dedic = max_skp("select max(skp_person_address) as max_skp_person_address from " & schema & ".PERSON_ADDRESS")
                skp_contact_dedic = max_skp("select max(skp_person_contact) as max_skp_person_contact from " & schema & ".PERSON_CONTACT")
                skp_demography_dedic = max_skp("select max(skp_person_demography) as max_skp_person_demography from " & schema & ".PERSON_DEMOGRAPHY")
            End If
            
'''''            ' NASLEDUJICI KOMENT BYL BUG VE VERZI 5.4 - MUZE SE VYHODIT KOMPLET
''''            If skp_person_Dedic = 0 Then 'Osoba není klient banky - nenalezena v SONIK_AB_KLIENT_DTB
''''                'RISK-985
''''                sql = "select min(skp_person) as min_skp_person from " & schema & ".PERSON"
''''
''''                Set rst1 = New ADODB.Recordset
''''                rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''''                    If Not rst1.EOF Then
''''                        skp_person_Dedic = rst1![min_skp_person]
''''                        skp_person_Dedic = skp_person_Dedic - 1
''''                    End If
''''                rst1.Close
''''
''''
''''                sql = "select " & schema & ".person_seq.nextval as id_person_seq_nextval from dual"
''''
''''                Set rst1 = New ADODB.Recordset
''''                rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''''                    If Not rst1.EOF Then
''''                        id_person_Dedic = rst1![id_person_seq_nextval] 'Nové ID_PERSON pro dìdice
''''                    End If
''''                rst1.Close
''''
''''
''''
''''                'cuid do minusu do dalších tabulek
''''                cuid_dedic = skp_person_Dedic
''''
''''            End If
''''
''''
''''             Else 'Osoba je klient 'skp_person_Dedic > 0
''''
''''                    sql = "select " & schema & ".person_seq.nextval as id_person_seq_nextval  from dual"
''''
''''                    Set rst1 = New ADODB.Recordset
''''                    rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''''                        If Not rst1.EOF Then
''''                            id_person_Dedic = rst1![id_person_seq_nextval] 'Nové ID_PERSON pro dìdice
''''                        End If
''''                    rst1.Close
''''
''''                    sql = "select max(skp_person_address) as max_skp_person_address from " & schema & ".PERSON_ADDRESS"
''''
''''                    Set rst1 = New ADODB.Recordset
''''                    rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''''                        If Not rst1.EOF Then
''''                            skp_address_dedic = rst1![max_skp_person_address]
''''                            skp_address_dedic = skp_address_dedic + 1
''''                        End If
''''                    rst1.Close
''''
''''                    sql = "select max(skp_person_contact) as max_skp_person_contact from " & schema & ".PERSON_CONTACT"
''''
''''                    Set rst1 = New ADODB.Recordset
''''                    rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''''                        If Not rst1.EOF Then
''''                            skp_contact_dedic = rst1![max_skp_person_contact]
''''                            skp_contact_dedic = skp_contact_dedic + 1
''''                        End If
''''                    rst1.Close
''''
''''            End If
            
            
            
            
            'tabulka PERSON
            'FLAG_PRINCIPAL_PAID_OBS, FLAG_INTEREST_PAID_OBS, FLAG_UNDESIRABILITY - zatím nastavím na N - v denním loadu by se mìlo pøepoèítávat
            'TEXT_CONTRACT_NUMBER se nebude vkládat pøímo odtud, protože data nemám jednodušše k dispozici ale dotáhne se na druhý den v proceduøe z DWH
            
            Call TransBegin
                
                ' RISK-985 NEDORIESENE
                Debug.Assert id_person_Dedic <> 0
                sql = "insert into " & schema & ".PERSON (ID_PERSON, CUID,  NAME, DATE_PAYOFF, DTIME_MODIFICATION, EMPLOYEE, PAYOFF_REASON, FLAG_PRINCIPAL_PAID_OBS, FLAG_INTEREST_PAID_OBS, FLAG_UNDESIRABILITY,  DTIME_VALID_FROM, DTIME_VALID_TO, FLAG_VALID, SKP_PERSON) " & _
                      " values (" & id_person_Dedic & ",'" & cuid_dedic & "','" & TextBox_Surname.value & " " & TextBox_Name.value & "', date'3000-01-01', sysdate, '" & UCase(Main.username) & "', 'XNA', 'N','N','N', sysdate, date'3000-01-01','Y', " & skp_person_Dedic & ")"
                execute (sql) 'RISK-517 - DONE
                
                'tabulka PERSON_DEMOGRAPHY
                
                Debug.Assert skp_demography_dedic <> 0
                sql = "insert into " & schema & ".PERSON_DEMOGRAPHY (skp_person_demography, ID_PERSON, GENDER, DATE_BIRTH, TEXT_BIRTH_NUMBER) " & _
                      " values (" & skp_demography_dedic & "," & id_person_Dedic & ",'" & ComboBox_Gender.value & "', to_date('" & TextBox_Date_Birth.value & "','dd.mm.yyyy'),'" & TextBox_RC.value & "')"
                execute (sql)
                
                'tabulka PERSON_ADDRESS
                
                Debug.Assert skp_address_dedic <> 0
                sql = "insert into " & schema & ".PERSON_ADDRESS (skp_person_address, ID_PERSON, NAME_REPORTS_ADDRESS_STREET, TEXT_REPORTS_ADDRESS_NUM, CODE_REPORTS_ADDRESS_ZIP, NAME_REPORTS_ADDRESS_TOWN, NAME_REPORTS_ADDRESS_COUNTRY, NAME_PERMANENT_ADDRESS_STREET, TEXT_PERMANENT_ADDRESS_NUM, CODE_PERMANENT_ADDRESS_ZIP, NAME_PERMANENT_ADDRESS_TOWN, NAME_PERMANENT_ADDRESS_COUNTRY) " & _
                      " values (" & skp_address_dedic & "," & id_person_Dedic & ",'" & TextBox_Reports_Address_Street.value & "','" & TextBox_Reports_Address_Num.value & "','" & TextBox_Reports_Address_ZIP.value & "','" & TextBox_Reports_Address_Town.value & "','" & StrReportsAddressCountry & "','" & TextBox_Permanent_Address_Street.value & "','" & TextBox_Permanent_Address_Num.value & "','" & TextBox_Permanent_Address_ZIP.value & "','" & TextBox_Permanent_Address_Town.value & "','" & StrPermanentAddressCountry & "')"
                execute (sql)
                
                'tabulka PERSON_CONTACT
                
                Debug.Assert skp_contact_dedic <> 0
                sql = "insert into " & schema & ".PERSON_CONTACT (skp_person_contact, ID_PERSON, PRIMARY_PHONE, PRIMARY_EMAIL) " & _
                    " values (" & skp_contact_dedic & "," & id_person_Dedic & ",'" & TextBox_Phone.value & "','" & TextBox_Email.value & "')"
                execute (sql)
                
            Call TransCommit
            MsgBox "Osoba byla úspìšnì pøidána do Jezevèíka.", vbInformation, "Informaèní zpráva"
        
        End If
         
    rst19.Close
    
    'tabulka HEIR - tam vkládám vždy pokud insertuju nového dìdice
    
    Call TransBegin
        sql = "insert into " & schema & ".HEIRS (SKP_HEIR, id_person_dead, id_person_heir,  AMT_RECEIVABLE, VARIABLE_SYMBOL, FLAG_INFORMOVAN_O_NAHRAVANI, FLAG_DELETED, DTIME_INSERTED, EMPLOYEE_INSERTED) " & _
              " values (" & schema & ".heirs_seq.nextval," & c.id_person & "," & id_person_Dedic & "," & StrAmtReceivable & ",'" & TextBox_Payment_Ref_Number.value & "','X','N',sysdate,'" & UCase(Main.username) & "')"
        execute (sql)
    Call TransCommit
    
    MsgBox "Dìdic byl úspìšnì nahrán.", vbInformation, "Informaèní zpráva"


Else 'Update stávajících dat - pouze pokud je CUID záporný - musí být nastaveno na buttonu
    If MsgBox("Opravdu chcete uložit zmìnu?", vbYesNo, "Uložit") = vbYes Then
        
        Call TransBegin
            'Iniciální texty, na které budu nabalovat
            StrZmena = vbNullString
            
            'tabulka PERSON
            If CheckBox_RC Or CheckBox_Name Or CheckBox_Surname Then
                sql = "update " & schema & ".person set "
                
        '            If CheckBox_RC Then
        '                sql = sql & " TEXT_BIRTH_NUMBER = '" & TextBox_RC.value & "'," 'RISK-517 - DONE
        '                StrZmena = StrZmena & "rè (" & TextBox_RC.value & "), "
        '            End If
                If CheckBox_Name Or CheckBox_Surname Then
                    sql = sql & " NAME = '" & TextBox_Surname.value & " " & TextBox_Name.value & "',"
                    StrZmena = StrZmena & "jméno (" & TextBox_Surname.value & " " & TextBox_Name.value & "), "
                End If
                
                'Dám vždy na konci èárky, a tu poslední co je navíc tak promažu
                sql = Left(sql, Len(sql) - 1)
                sql = sql & ", dtime_modification = sysdate, employee = '" & UCase(Main.username) & "'"
                sql = sql & " where id_person = " & id_person_Dedic
                execute (sql)
            End If
      
            'tabulka PERSON_DEMOGRAPHY
            If CheckBox_RC Or CheckBox_Date_Birth Or CheckBox_Gender Then
                sql = "update " & schema & ".person_demography set "
                
                If CheckBox_RC Then
                    sql = sql & " TEXT_BIRTH_NUMBER = '" & TextBox_RC.value & "',"
                    StrZmena = StrZmena & "rè (" & TextBox_RC.value & "), "  'RISK-517 - DONE
                End If
                If CheckBox_Date_Birth Then
                    sql = sql & " DATE_BIRTH = to_date('" & TextBox_Date_Birth.value & "','DD.MM.YYYY'),"
                    StrZmena = StrZmena & "dat. nar. (" & TextBox_Date_Birth.value & "), "
                End If
                If CheckBox_Gender Then
                    sql = sql & " GENDER = '" & ComboBox_Gender.value & "',"
                    StrZmena = StrZmena & "pohlaví (" & ComboBox_Gender.value & "), "
                End If
                
                sql = Left(sql, Len(sql) - 1)
                sql = sql & " where id_person = " & id_person_Dedic
                execute (sql)
            End If
        
        
            'tabulka PERSON_ADDRESS
            If CheckBox_Permanent_Address Or CheckBox_Reports_Address Then
                sql = "update " & schema & ".person_address set "
                
                If CheckBox_Permanent_Address Then
                    sql = sql & " NAME_PERMANENT_ADDRESS_STREET = '" & TextBox_Permanent_Address_Street.value & "', TEXT_PERMANENT_ADDRESS_NUM = '" & TextBox_Permanent_Address_Num.value & "', CODE_PERMANENT_ADDRESS_ZIP = '" & TextBox_Permanent_Address_ZIP.value & "', NAME_PERMANENT_ADDRESS_TOWN = '" & TextBox_Permanent_Address_Town.value & "', NAME_PERMANENT_ADDRESS_COUNTRY = '" & StrPermanentAddressCountry & "',"
                StrZmena = StrZmena & "nová trvalá adresa, "
                End If
                If CheckBox_Reports_Address Then
                    sql = sql & " NAME_REPORTS_ADDRESS_STREET = '" & TextBox_Reports_Address_Street.value & "', TEXT_REPORTS_ADDRESS_NUM = '" & TextBox_Reports_Address_Num.value & "', CODE_REPORTS_ADDRESS_ZIP = '" & TextBox_Reports_Address_ZIP.value & "', NAME_REPORTS_ADDRESS_TOWN = '" & TextBox_Reports_Address_Town.value & "', NAME_REPORTS_ADDRESS_COUNTRY = '" & StrReportsAddressCountry & "',"
                StrZmena = StrZmena & "nová kontaktní adresa, "
                End If
                
                sql = Left(sql, Len(sql) - 1)
                sql = sql & " where id_person = " & id_person_Dedic
                execute (sql)
            End If
        
        
            'tabulka PERSON_CONTACT
            If CheckBox_Phone Or CheckBox_Email Then
                sql = "update " & schema & ".person_contact set "
                
                If CheckBox_Phone Then
                    sql = sql & " PRIMARY_PHONE = '" & TextBox_Phone.value & "',"
                    StrZmena = StrZmena & "telefon (" & TextBox_Phone.value & "), "
                End If
                If CheckBox_Email Then
                    sql = sql & " PRIMARY_EMAIL = '" & TextBox_Email.value & "',"
                    StrZmena = StrZmena & "e-mail (" & TextBox_Email.value & "), "
                End If
            
                sql = Left(sql, Len(sql) - 1)
                sql = sql & " where id_person = " & id_person_Dedic
                execute (sql)
            End If
        
        
            'tabulka HEIR
            If CheckBox_Receivable Or CheckBox_Payment_Ref_Number Then
                sql = "update " & schema & ".heirs set "
                
                If CheckBox_Receivable Then
                    sql = sql & " AMT_RECEIVABLE = " & StrAmtReceivable & ","
                    StrZmena = StrZmena & "Výše pohledávky (" & StrAmtReceivable & "), "
                End If
                If CheckBox_Payment_Ref_Number Then
                    sql = sql & " VARIABLE_SYMBOL = '" & TextBox_Payment_Ref_Number.value & "',"
                    StrZmena = StrZmena & "VS byl zmìnìn, døíve " & d.strVariabilniSymbol & " novì " & TextBox_Payment_Ref_Number.value & ", "
                End If
            
                sql = Left(sql, Len(sql) - 1)
                sql = sql & " where skp_heir = " & d.lngSkp_heir
                execute (sql)
            End If
        
        
        
            If StrZmena <> vbNullString Then
              
                sqlMaxHeirNote = "select max(skp_heir_note) as max_skp_heir_note from " & schema & ".HEIR_NOTE"
            
                Set rst1 = New ADODB.Recordset
                rst1.Open sqlMaxHeirNote, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
                
                   If Not IsNull(rst1![max_skp_heir_note]) Then
                       skp_heir_note = rst1![max_skp_heir_note] + 1
                   Else
                       skp_heir_note = 1
                   End If
                  
                rst1.Close
                
                
                sql = "insert into " & schema & ".heir_note (SKP_HEIR_NOTE, SKP_HEIR_NOTE_TYPE, SKP_HEIR, TEXT_NOTE, DTIME_NOTE_CREATED, EMPLOYEE)" & _
                      " values (" & skp_heir_note & ", 2, " & d.lngSkp_heir & _
                      ", 'Zmìna údajù dìdice: " & Mid(StrZmena, 1, Len(StrZmena) - 2) & "', sysdate, '" & UCase(Main.username) & "')"
                
                execute (sql)
            End If
        
        Call TransCommit
        
        If StrZmena <> vbNullString Then
            MsgBox "Údaje o dìdici byly úspìšnì zmìnìny. " & Mid(StrZmena, 1, Len(StrZmena) - 2), vbInformation, "Informaèní zpráva"
        Else
            MsgBox "Údaje o dìdici nebyly  zmìnìny. Zøejmì jsi zapomnìl/(a) odškrnout zaškrtávátko.", vbExclamation, "Chybné zadání hodnot"
            Exit Sub
        End If
        
        
        
        
    Else
        Exit Sub 'Pokud klikne, že nechce uložit zmìnu tak hold nic
    End If
End If


Me.Hide

'XXXX Musím udìlat nìjaký ten repaint...
If Not BlnNovyDedic Then
    ' call heirs.refresh
    Klient_heir.Show
End If


End Sub

Private Sub CommandButton_Zrus_Click()
    
Dim CB As control

Me.Height = 364.5

If BlnNovyDedic Then

    'Iniciálnì všechny buttony povolím editovat
        For Each CB In Me.Controls
            If Left(CB.name, 7) = "TextBox" Or Left(CB.name, 8) = "ComboBox" Then
                CB.value = vbNullString
                CB.Enabled = True
            End If
        Next CB
        
    CheckBox_Address_Match.Visible = True
    CheckBox_Address_Match.Enabled = True
    CheckBox_Address_Match.value = False
    
Else
    Call UserForm_Activate
End If
End Sub













'Private Sub CheckBox1_Click()
'Label1.Visible = Not CheckBox1.value
'TextBox1.Visible = CheckBox1.value
'
'TextBox1.text = Label1.Caption
'End Sub



'*******************************************************************************

' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()
Dim CB As control

'Na zaèátku zobrazím pouze rodné èíslo - jednoznaèný identifikátor
Me.Height = 364.5

'Iniciální nastavení
Label_Update_Help.Visible = False 'Nápovìda pouze pøi updateu informací o klientovi
BlnKlientHledan = False

For Each CB In Me.Controls
    If Left(CB.name, 7) = "TextBox" Or Left(CB.name, 8) = "ComboBox" Then
        CB.value = vbNullString
        CB.Enabled = True
    End If
Next CB

CheckBox_Address_Match.Visible = True
CheckBox_Address_Match.Enabled = True
CheckBox_Address_Match.value = False


'Inicálnì schovám, objeví se, až pokud vyberou jiný stát
TextBox_Permanent_Address_Country.Visible = False
TextBox_Reports_Address_Country.Visible = False
'Inicální naplnìní hodnot do comboboxù

ComboBox_Gender.Clear
ComboBox_Gender.AddItem "M"
ComboBox_Gender.AddItem "Z"


ComboBox_Permanent_Address_Country.Clear
ComboBox_Permanent_Address_Country.AddItem "Czech Republic"
ComboBox_Permanent_Address_Country.AddItem "Slovakia"
ComboBox_Permanent_Address_Country.AddItem "Jiný"
ComboBox_Permanent_Address_Country.ListIndex = 0

ComboBox_Reports_Address_Country.Clear
ComboBox_Reports_Address_Country.AddItem "Czech Republic"
ComboBox_Reports_Address_Country.AddItem "Slovakia"
ComboBox_Reports_Address_Country.AddItem "Jiný"
ComboBox_Reports_Address_Country.ListIndex = 0

'Tip texty
TextBox_RC.ControlTipText = "Rodné èíslo zadávej bez lomítek a mezer"
TextBox_Date_Birth.ControlTipText = "Zadej datum ve formátu 17.3.1986."
TextBox_Phone.ControlTipText = "Telefonní èíslo zadávej ve formátu 606095623"
CommandButton_Hledej.ControlTipText = "Klienta hledej podle rodného èísla. Pokud jej neznáš, zadej jméno, pøíjmení a datum narození"






If BlnNovyDedic Then
    Label_Form_Name.Caption = "Zadání informací o novém dìdici"
    
    'Odškrtnu všechny checkboxy za uživatele a nenechám zmìnit
    For Each CB In Me.Controls
            If Left(CB.name, 8) = "CheckBox" And CB.name <> "CheckBox_Address_Match" Then
                CB.value = True
                CB.Enabled = False
            End If
    Next CB

   

    
Else: 'Nechám uživateli volnost, to co zaškrtne potom ovlivní to, co se ve skuteènosti updatuje
      
    Label_Form_Name.Caption = "Úprava informací o stávajícím dìdici"
    Label_Update_Help.Visible = True 'Nápovìda pouze pøi updateu informací o klientovi
    For Each CB In Me.Controls
            If Left(CB.name, 8) = "CheckBox" Then
                CB.value = False
                CB.Enabled = True
            End If
    Next CB
    
    BlnKlientHledan = True 'Promìnná slouží když hledají nového dìdice, tady upravuji stávajícího tedy byl již hledán
   
    'Pøedvyplním údaje o dìdici z objektù
        TextBox_RC.value = c.Pin
        TextBox_CUID.value = c.CUID
        TextBox_Name.value = "" & c.Name_first & ""
        TextBox_Surname.value = "" & c.Name_last & ""
        TextBox_Date_Birth.value = format$(c.Date_Birth, "d.m.yyyy")
        ComboBox_Gender.value = c.Gender
        TextBox_Phone.value = c.PrimaryPhone
        TextBox_Email.value = c.PrimaryEmail
        TextBox_Permanent_Address_Street.value = c.Address_P_Ulice
        TextBox_Permanent_Address_Num.value = c.Address_P_CP
        TextBox_Permanent_Address_Town.value = c.Address_P_Mesto
        TextBox_Permanent_Address_ZIP.value = c.Address_P_PSC
        TextBox_Reports_Address_Street.value = c.Address_C_Ulice
        TextBox_Reports_Address_Num.value = c.Address_C_CP
        TextBox_Reports_Address_Town.value = c.Address_C_Mesto
        TextBox_Reports_Address_ZIP.value = c.Address_C_PSC
        TextBox_Receivable = d.dblAmtReceivable
        TextBox_Payment_Ref_Number = d.strVariabilniSymbol
        
        
        If c.Address_P_Zeme = "Czech Republic" Or c.Address_P_Zeme = "Slovakia" Then
            ComboBox_Permanent_Address_Country = c.Address_P_Zeme
        Else: ComboBox_Permanent_Address_Country = "Jiný"
              TextBox_Permanent_Address_Country = c.Address_P_Zeme
        End If
        
        If c.Address_C_Zeme = "Czech Republic" Or c.Address_C_Zeme = "Slovakia" Then
            ComboBox_Reports_Address_Country = c.Address_C_Zeme
        Else: ComboBox_Reports_Address_Country = "Jiný"
              TextBox_Reports_Address_Country = c.Address_C_Zeme
        End If
        
        id_person_Dedic = c.id_person
    
    
        'nesmí mít možnost editovat CUID... nikdy nikdy
        CheckBox_CUID.value = False
        CheckBox_CUID.Enabled = False
        TextBox_CUID.Enabled = False
        
        'Pokud bude CUID kladný, tak nesmí editovat ani údaje o klientovi - pouze pohledávku a variabilní symbol
        If c.CUID > 0 Then
                For Each CB In Me.Controls 'Nepovolím zmìnit checkboxy
                    If Left(CB.name, 8) = "CheckBox" And Right(CB.name, 10) <> "Receivable" And Right(CB.name, 10) <> "Ref_Number" Then
                        CB.value = False
                        CB.Enabled = False
                    End If
                Next CB
                
                'Nastavení ostatních hodnot
                CheckBox_Address_Match.Visible = False
                TextBox_RC.Enabled = False
                TextBox_CUID.Enabled = False
                TextBox_Name.Enabled = False
                TextBox_Surname.Enabled = False
                TextBox_Date_Birth.Enabled = False
                ComboBox_Gender.Enabled = False
                TextBox_Phone.Enabled = False
                TextBox_Email.Enabled = False
                TextBox_Permanent_Address_Street.Enabled = False
                TextBox_Permanent_Address_Num.Enabled = False
                TextBox_Permanent_Address_Town.Enabled = False
                TextBox_Permanent_Address_ZIP.Enabled = False
                TextBox_Reports_Address_Street.Enabled = False
                TextBox_Reports_Address_Num.Enabled = False
                TextBox_Reports_Address_Town.Enabled = False
                TextBox_Reports_Address_ZIP.Enabled = False
                ComboBox_Permanent_Address_Country.Enabled = False
                ComboBox_Reports_Address_Country.Enabled = False
                
                
        End If
End If



End Sub


'Automatické validace:

Private Sub TextBox_RC_Change()

If Len(TextBox_RC.value) = 0 Then
    TextBox_Date_Birth.Enabled = True
    ComboBox_Gender.Enabled = True
    Exit Sub 'Aby to pokaždé zbyteènì neøvalo
End If

If Not IsNumeric(TextBox_RC.value) Then
    MsgBox "Zadávej prosím pouze èíselné hodnoty.", vbExclamation, "Chybné zadání hodnot"
    Exit Sub
End If

If CheckBox_RC And Len(TextBox_RC.value) >= 6 Then 'Pokud už mám dostatek hodnot k rozhodnutí, tak naplním

    If Mid(TextBox_RC.value, 3, 1) >= 5 Then
        ComboBox_Gender.value = "Z"
       TextBox_Date_Birth.value = ZiskejDatumRC(TextBox_RC.value)
    Else: ComboBox_Gender.value = "M"
       TextBox_Date_Birth.value = ZiskejDatumRC(TextBox_RC.value)
    End If
    
    TextBox_Date_Birth.Enabled = False
    ComboBox_Gender.Enabled = False
    
Else: 'Jinak vyèistím
    ComboBox_Gender.value = ""
    TextBox_Date_Birth.value = ""

    TextBox_Date_Birth.Enabled = True
    ComboBox_Gender.Enabled = True
    
End If

End Sub


Private Sub ComboBox_Permanent_Address_Country_Change()

If CheckBox_Permanent_Address And ComboBox_Permanent_Address_Country.value = "Jiný" Then
    TextBox_Permanent_Address_Country.Visible = True
    TextBox_Permanent_Address_Country.ControlTipText = "Zde zadej stát, pokud možno anglicky :)"
Else: TextBox_Permanent_Address_Country.Visible = False
End If

'Aby nemohl zadat jiné hodnoty, pokud zašktrl, že dává stejné
If CheckBox_Address_Match And ComboBox_Permanent_Address_Country <> vbNullString Then
    ComboBox_Reports_Address_Country = ComboBox_Permanent_Address_Country
End If

End Sub

Private Sub ComboBox_Reports_Address_Country_Change()

If CheckBox_Reports_Address And ComboBox_Reports_Address_Country.value = "Jiný" Then
    TextBox_Reports_Address_Country.Visible = True
    TextBox_Reports_Address_Country.ControlTipText = "Zde zadej stát, pokud možno anglicky :)"
Else: TextBox_Reports_Address_Country.Visible = False
End If

End Sub


Private Sub CheckBox_Address_Match_Click()

If CheckBox_Address_Match Then
    TextBox_Reports_Address_Street = TextBox_Permanent_Address_Street
    TextBox_Reports_Address_Num = TextBox_Permanent_Address_Num
    TextBox_Reports_Address_Town = TextBox_Permanent_Address_Town
    TextBox_Reports_Address_ZIP = TextBox_Permanent_Address_ZIP
    ComboBox_Reports_Address_Country = ComboBox_Permanent_Address_Country
    TextBox_Reports_Address_Country = TextBox_Permanent_Address_Country
    ' A dám disable
    TextBox_Reports_Address_Street.Enabled = False
    TextBox_Reports_Address_Num.Enabled = False
    TextBox_Reports_Address_Town.Enabled = False
    TextBox_Reports_Address_ZIP.Enabled = False
    ComboBox_Reports_Address_Country.Enabled = False
    TextBox_Reports_Address_Country.Enabled = False

Else:
    TextBox_Reports_Address_Street.Enabled = True
    TextBox_Reports_Address_Num.Enabled = True
    TextBox_Reports_Address_Town.Enabled = True
    TextBox_Reports_Address_ZIP.Enabled = True
    ComboBox_Reports_Address_Country.Enabled = True
    TextBox_Reports_Address_Country.Enabled = True

End If

End Sub


Private Sub TextBox_Permanent_Address_Street_Change()
'Aby nemohl zadat jiné hodnoty, pokud zašktrl, že dává stejné
If CheckBox_Address_Match Then
    TextBox_Reports_Address_Street = TextBox_Permanent_Address_Street
End If

End Sub

Private Sub TextBox_Permanent_Address_Num_Change()
'Aby nemohl zadat jiné hodnoty, pokud zašktrl, že dává stejné
If CheckBox_Address_Match Then
    TextBox_Reports_Address_Num = TextBox_Permanent_Address_Num
End If

End Sub

Private Sub TextBox_Permanent_Address_Town_Change()
'Aby nemohl zadat jiné hodnoty, pokud zašktrl, že dává stejné
If CheckBox_Address_Match Then
    TextBox_Reports_Address_Town = TextBox_Permanent_Address_Town
End If

End Sub

Private Sub TextBox_Permanent_Address_ZIP_Change()
'Aby nemohl zadat jiné hodnoty, pokud zašktrl, že dává stejné
If CheckBox_Address_Match Then
    TextBox_Reports_Address_ZIP = TextBox_Permanent_Address_ZIP
End If

End Sub


Private Sub TextBox_Permanent_Address_Country_Change()
'Aby nemohl zadat jiné hodnoty, pokud zašktrl, že dává stejné
If CheckBox_Address_Match Then
    TextBox_Reports_Address_Country = TextBox_Permanent_Address_Country
End If

End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, _
                            CloseMode As Integer)

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


