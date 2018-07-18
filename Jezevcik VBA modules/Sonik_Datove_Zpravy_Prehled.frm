VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sonik_Datove_Zpravy_Prehled 
   Caption         =   "Soník - Pøehled datových zpráv"
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   22395
   OleObjectBlob   =   "Sonik_Datove_Zpravy_Prehled.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sonik_Datove_Zpravy_Prehled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Option Explicit



'RISK-1094







Private Sub Label_Display_All_DZ_Click()

If ListView_Datove_Zpravy.ListItems.count > 0 Then

Call nacteni_dat_Sonik_Datove_Zpravy_Prehled(False) 'Naèti všechny

End If

End Sub


'RISK-1094
Private Sub ListViewToDo_Click()
'Podle toho na kterou datovku klikne, tak se bude pøedvyplòovat ASSIGNEE a NOTE v comboboxech

'Dovolím klikání, jenom pokud tam je na co klikat, jinak to spadne
If ListViewToDo.ListItems.count > 0 Then

        ComboBoxAssignee.text = ListvZiskejHodnotu(ListViewToDo, "Øeší")
        TextBox_ToDoNote.text = ListvZiskejHodnotu(ListViewToDo, "Poznámky")
        
End If

End Sub

'RISK-1094
Private Sub ListViewToDo_DblClick()
'Pokud udìlají dvojklik na datovku, tak se nasimuluje dvojklik na pøehledu na listview datovek

Dim sql As String

'Dovolím klikání, jenom pokud tam je na co klikat, jinak to spadne
If ListViewToDo.ListItems.count > 0 Then

    'Toto by mely byt public promenne nadefinovane v sub SONIK
    SKP_datove_zpravy = ListvZiskejHodnotu(ListViewToDo, "SKP_DATA_MESSAGE")
    
    sql = "select DATA_MESSAGE_NUMBER, DATE_DATA_MESSAGE_DELIVERY, DATA_MESSAGE_SENDER, TEXT_FILE_NUMBER, TEXT_REFERENCE_NUMBER from " & schema & ".SONIK_DATA_MESSAGE where SKP_DATA_MESSAGE = " & SKP_datove_zpravy
    
    Set rst1 = New ADODB.Recordset
    rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        If Not rst1.EOF Then 'Oèekávám jeden záznam
    
            Cislo_datove_zpravy = "" & rst1![data_message_number]
            Datum_doruceni_datove_zpravy = "" & rst1![date_data_message_delivery]
            Odesilatel_datove_zpravy = "" & rst1![data_message_sender]
            Spisova_Znacka_datove_zpravy = "" & rst1![text_file_number]
            Cislo_Jednaci_datove_zpravy = "" & rst1![text_reference_number]
        
        End If
    rst1.Close
    

    Me.Hide
    Sonik_Datova_Zprava.Show
    
    Call nacteni_dat_Sonik_Datova_Zprava
        
End If

End Sub

'RISK-1094
' ListView sorting
Private Sub ListViewToDo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListViewToDo, ColumnHeader)
End Sub



'RISK-1094
Private Sub CommandButton_Uloz_Click()

Call SonikSaveToDoListChange(ListvZiskejHodnotu(ListViewToDo, "SKP_DATA_MESSAGE"), UCase(ComboBoxAssignee.text), TextBox_ToDoNote.text)

'Refresh to do listu
Call query_ToDoList("ROWNUM <= 200 and NAME_DATA_MESSAGE_STATUS <> 'Zpracováno'")

End Sub

Private Sub UserForm_Activate()

Dim i As Integer, username As String 'RISK-1094


Call nacteni_dat_Sonik_Datove_Zpravy_Prehled(True)

'Vynulování promìnné pro spolupráci s importovacícm workbookem
Workbook_Open = 0

'Naètení dat pro ToDoList


'RISK-1094
If Sonik_Datove_Zpravy_Prehled.ListViewToDo.ColumnHeaders.count = 0 Then

'Nastaveni listview
 Sonik_Datove_Zpravy_Prehled.ListViewToDo.ListItems.Clear
Call Sonik_Datove_Zpravy_Prehled.ListViewToDo.ColumnHeaders.Add(, , "SKP_DATA_MESSAGE_TODO_LIST", 0)
Call Sonik_Datove_Zpravy_Prehled.ListViewToDo.ColumnHeaders.Add(, , "SKP_DATA_MESSAGE", 0)
Call Sonik_Datove_Zpravy_Prehled.ListViewToDo.ColumnHeaders.Add(, , "Èíslo DZ", 60)
Call Sonik_Datove_Zpravy_Prehled.ListViewToDo.ColumnHeaders.Add(, , "CUID", 70)
Call Sonik_Datove_Zpravy_Prehled.ListViewToDo.ColumnHeaders.Add(, , "Jméno klienta", 70)
Call Sonik_Datove_Zpravy_Prehled.ListViewToDo.ColumnHeaders.Add(, , "Týká se", 100)
Call Sonik_Datove_Zpravy_Prehled.ListViewToDo.ColumnHeaders.Add(, , "Deadline", 50)
Call Sonik_Datove_Zpravy_Prehled.ListViewToDo.ColumnHeaders.Add(, , "Zbývá ke zpracování", 50)
Call Sonik_Datove_Zpravy_Prehled.ListViewToDo.ColumnHeaders.Add(, , "Øeší", 70)
Call Sonik_Datove_Zpravy_Prehled.ListViewToDo.ColumnHeaders.Add(, , "Stav žádosti Soník", 70)
Call Sonik_Datove_Zpravy_Prehled.ListViewToDo.ColumnHeaders.Add(, , "Poznámky", 200)
End If

Call NastavKliceListview(Sonik_Datove_Zpravy_Prehled.ListViewToDo)

'Naètu jen prvních 200 datovek
Call query_ToDoList("ROWNUM <= 200 and NAME_DATA_MESSAGE_STATUS <> 'Zpracováno'")

'Naplním boxíky - prvnì pøihlášeného uživatele, a pak všechny ostatní z databáze

'Iniciálnì
ComboBoxAssignee.text = ""
'ComboBoxAssignee.AddItem ""
TextBox_ToDoNote.text = ""
ComboBoxAssigneeLookUp.text = ""
ComboBoxDMStatusLookUp.text = ""

Call FillAssigneSonikToDo(ComboBoxAssigneeLookUp)
Call FillDMStatusToDo


Call FillAssigneSonikToDo(ComboBoxAssignee)

'Seøazení ListView podle deadline - jde o to, že jim to nefungovalo, když se vracely z jiných karet
Call Listview_Proper_Sort(Sonik_Datove_Zpravy_Prehled.ListViewToDo, Sonik_Datove_Zpravy_Prehled.ListViewToDo.ColumnHeaders.item("Deadline"), blnAscendingOrder:=True)


End Sub



'RISK-1094
Private Sub FillDMStatusToDo()

Dim sql As String

'Naplním boxíky - prvnì pøihlášeného uživatele, a pak všechny ostatní z databáze
ComboBoxDMStatusLookUp.Clear


ComboBoxDMStatusLookUp.AddItem ""
ComboBoxDMStatusLookUp.AddItem "Všechny nezpracované"

sql = "select NAME_DATA_MESSAGE_STATUS from " & schema & ".SONIK_CL_DATA_MESSAGE_STATUS"


Set rst1 = New ADODB.Recordset
rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst1.EOF Then
        rst1.MoveFirst
        For i = 1 To rst1.RecordCount
            
            ComboBoxDMStatusLookUp.AddItem rst1![NAME_DATA_MESSAGE_STATUS]
            
        rst1.MoveNext
        Next i
    End If
rst1.Close
ComboBoxDMStatusLookUp.ListIndex = 0

End Sub



'RISK-1094
Private Sub LabelDisplayAllToDo_Click()

If ListViewToDo.ListItems.count > 0 Then

Call query_ToDoList 'Naèti všechny

End If

End Sub

'RISK-1094
Private Sub CommandButtonFilterToDoLookup_Click()

Dim StrAssigneeCondition As String, StrDMStatusCondition As String

'Kde nic není ani smrt nebere
If ComboBoxAssigneeLookUp.text = "" And ComboBoxDMStatusLookUp.text = "" Then Exit Sub

If ComboBoxAssigneeLookUp.text = "" Then
    StrAssigneeCondition = " 1 = 1 "
Else: StrAssigneeCondition = " USERNAME = '" & ComboBoxAssigneeLookUp.text & "' "
End If

If ComboBoxDMStatusLookUp.text = "" Then
    StrDMStatusCondition = " 1 = 1 "
ElseIf ComboBoxDMStatusLookUp.text = "Všechny nezpracované" Then
    StrDMStatusCondition = " NAME_DATA_MESSAGE_STATUS <> 'Zpracováno' "
Else: StrDMStatusCondition = " NAME_DATA_MESSAGE_STATUS = '" & ComboBoxDMStatusLookUp.text & "' "
End If

'Pokud podmínku nehledá, tak je tam 1=1 - vždy splnìno
Call query_ToDoList(StrAssigneeCondition & " AND " & StrDMStatusCondition)

End Sub


'RISK-1094
Private Sub CommandButtonVymazToDoLookUp_Click()

'Naètu jen prvních 200 datovek
Call query_ToDoList("ROWNUM <= 200 and NAME_DATA_MESSAGE_STATUS <> 'Zpracováno'")

ComboBoxAssigneeLookUp.text = ""
ComboBoxDMStatusLookUp.text = ""

End Sub


'RISK-1094 'Dìlám jako zvláštní proceduru, protože ji budu plnit vícekrát a podle rùzných filtrù
Sub query_ToDoList(Optional ByVal StrCondition As String)
   
Dim keys As Dictionary, itmX As ListItem, sql As String
     
Set keys = KliceKListview(Sonik_Datove_Zpravy_Prehled.ListViewToDo)
     
sql = "SELECT t.skp_data_message_todo_list, v.skp_data_message, v.data_message_number, v.name_data_message_status, v.date_deadline, v.cuids_string, v.names_string, v.info_string " & _
      ", su.username, t.text_note " & _
      " FROM " & schema & ".sonik_v_data_message_todo_list v " & _
      " LEFT JOIN " & schema & ".sonik_data_message_todo_list t ON v.skp_data_message = t.skp_data_message  " & _
      "                                                                AND t.flag_valid = 'Y'  " & _
      "                                                                AND t.flag_deleted = 'N'  " & _
      " LEFT JOIN " & schema & ".sonik_users su ON t.skp_sonik_user = su.skp_sonik_user " & _
      " ORDER BY v.date_deadline - trunc(sysdate) "
Sonik_Datove_Zpravy_Prehled.ListViewToDo.ListItems.Clear
   
If StrCondition = "" Then StrCondition = " 1 = 1 "
    
sql = "select * from (" & sql & ") where " & StrCondition

   
Set rst1 = New ADODB.Recordset
rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
        Set itmX = Sonik_Datove_Zpravy_Prehled.ListViewToDo.ListItems.Add()
        itmX.text = "" & rst1![skp_data_message_todo_list] & ""
        itmX.SubItems(keys("SKP_DATA_MESSAGE")) = "" & rst1![SKP_DATA_MESSAGE] & ""
        itmX.SubItems(keys("Èíslo DZ")) = "" & rst1![data_message_number] & ""
        itmX.SubItems(keys("CUID")) = "" & rst1![cuids_string] & ""
        itmX.SubItems(keys("Jméno klienta")) = "" & rst1![names_string] & ""
        itmX.SubItems(keys("Týká se")) = "" & rst1![info_string] & ""
        itmX.SubItems(keys("Deadline")) = "" & rst1![date_deadline] & ""
        If Not IsNull(rst1![date_deadline]) Then itmX.SubItems(keys("Zbývá ke zpracování")) = rst1![date_deadline] - date Else itmX.SubItems(keys("Zbývá ke zpracování")) = ""
        itmX.SubItems(keys("Øeší")) = "" & rst1![username] & ""
        itmX.SubItems(keys("Stav žádosti Soník")) = "" & rst1![NAME_DATA_MESSAGE_STATUS] & ""
        itmX.SubItems(keys("Poznámky")) = "" & rst1![TEXT_NOTE] & ""
        
        rst1.MoveNext
    Next i
End If

rst1.Close
   
   
   
'Barvièky podle toho kolik zbývá ke zpracování

'0 –5 dní zbývá ke zpracování èervenì
'6 –13 dní zbývá ke zpracování oranžovì
'14 a déle dní zbývá ke zpracování mùže zùstat èernì

'RISK-1212
If Sonik_Datove_Zpravy_Prehled.ListViewToDo.ListItems.count > 0 Then

    For i = 1 To Sonik_Datove_Zpravy_Prehled.ListViewToDo.ListItems.count

        If ListvZiskejHodnotu(ListViewToDo, "Stav žádosti Soník", i) = "Zpracováno" Then
        
            For j = 1 To Sonik_Datove_Zpravy_Prehled.ListViewToDo.ColumnHeaders.count - 1
            Sonik_Datove_Zpravy_Prehled.ListViewToDo.ListItems(i).ListSubItems.item(j).ForeColor = barvaCerna  'Eerni
            Next j
        
        ElseIf ListvZiskejHodnotu(ListViewToDo, "Zbývá ke zpracování", i) = "" Then
        
            For j = 1 To Sonik_Datove_Zpravy_Prehled.ListViewToDo.ColumnHeaders.count - 1
            Sonik_Datove_Zpravy_Prehled.ListViewToDo.ListItems(i).ListSubItems.item(j).ForeColor = barvaRuda 'Èervená
            Next j
        
        
        ElseIf ListvZiskejHodnotu(ListViewToDo, "Zbývá ke zpracování", i) <= 5 Then
        
            For j = 1 To Sonik_Datove_Zpravy_Prehled.ListViewToDo.ColumnHeaders.count - 1
            Sonik_Datove_Zpravy_Prehled.ListViewToDo.ListItems(i).ListSubItems.item(j).ForeColor = barvaRuda 'Èervená
            Next j
        
        
        ElseIf ListvZiskejHodnotu(ListViewToDo, "Zbývá ke zpracování", i) <= 13 Then
        
            For j = 1 To Sonik_Datove_Zpravy_Prehled.ListViewToDo.ColumnHeaders.count - 1
            Sonik_Datove_Zpravy_Prehled.ListViewToDo.ListItems(i).ListSubItems.item(j).ForeColor = barvaOranzova 'Oranžová
            Next j
        

        
        Else: 'ListvZiskejHodnotu(ListViewToDo, "Zbývá ke zpracování", i) > 13
        
            For j = 1 To Sonik_Datove_Zpravy_Prehled.ListViewToDo.ColumnHeaders.count - 1
            Sonik_Datove_Zpravy_Prehled.ListViewToDo.ListItems(i).ListSubItems.item(j).ForeColor = RGB(0, 204, 102) 'Zelená
            Next j
        
        End If

    Next i

End If
   
   
'Seøazení ListView podle deadline - jde o to, že jim to nefungovalo, když se vracely z jiných karet
Call Listview_Proper_Sort(Sonik_Datove_Zpravy_Prehled.ListViewToDo, Sonik_Datove_Zpravy_Prehled.ListViewToDo.ColumnHeaders.item("Deadline"), blnAscendingOrder:=True)
   

End Sub




'Výbìr datové zprávy
Private Sub ListView_Datove_Zpravy_DblClick()


If ListView_Datove_Zpravy.ListItems.count > 0 Then


'Toto by mely byt public promenne nadefinovane v sub SONIK
SKP_datove_zpravy = ListView_Datove_Zpravy.SelectedItem.text
Cislo_datove_zpravy = ListView_Datove_Zpravy.SelectedItem.SubItems(1)
Datum_doruceni_datove_zpravy = ListView_Datove_Zpravy.SelectedItem.SubItems(2)
Odesilatel_datove_zpravy = ListView_Datove_Zpravy.SelectedItem.SubItems(3)
Spisova_Znacka_datove_zpravy = ListView_Datove_Zpravy.SelectedItem.SubItems(4)
Cislo_Jednaci_datove_zpravy = ListView_Datove_Zpravy.SelectedItem.SubItems(5)

Me.Hide
Sonik_Datova_Zprava.Show

Call nacteni_dat_Sonik_Datova_Zprava

End If


End Sub



' ListView sorting
Private Sub ListView_Datove_Zpravy_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_Datove_Zpravy, ColumnHeader)
End Sub



'Pøidání nového dopisu
Private Sub CommandButton_Zaevidovani_Dopis_Click()
'RISK-877

Sonik_Zaevidovani_Dopisu.Show

End Sub

'Zaevidování nových datových zpráv
Private Sub CommandButton_Zaevidovani_DZ_Click()

    'MsgBox "Otevøe se nový excel, kam si data uloží. Do SQL tabulky sonik_data_message nahrajeme jen potøebné sloupce, následnì se refreshne stránka a uvidí tam novou zprávu"
    If Workbook_Open = 1 Then
       Call Import_Workbook_DM
    Else
       Call Open_Workbook_DM
    End If

End Sub

'Otevøení excelu pro vložení dávky datových zpráv
Sub Open_Workbook_DM()

    Set DM_Import = Workbooks.Add
    Workbook_Open = 1

End Sub

'Nahrání dávky datových zpráv a zavøení importního excelu
Sub Import_Workbook_DM()
Dim ans As Variant

CntRows_Import = DM_Import.ActiveSheet.Cells(DM_Import.ActiveSheet.Rows.count, "C").End(xlUp).Row
'RISK-795
ans = MsgBox("Chcete vložit nové záznamy?", vbYesNo + vbQuestion, "Import dat do databáze")
    
Select Case ans

       Case vbNo
       Exit Sub
       Case vbYes
       Call Import_Workbook_DM_Yes

End Select
    
End Sub

Sub Import_Workbook_DM_Yes()
'RISK-795

Dim i As Long, j As Long, lRow As Long
Dim data_message_number As Long, data_message_delivery As Date, data_message_sender As String, text_file_number As String
Dim pom_zprava_vyctena As Long, CntRows_Import As Long
Dim max_skp_DM As Long, sql_insert_DM As String, sql_insert_DMS As String, sql_insert_DM_Notice_Automatic As String
Dim sql_max_skp_DM As String
Dim Time_Offset As Long
    
On Error GoTo chyba
    
'Prvnì naètu všechny datovky do listview, abych mohl hledat jednodušše, jestli už tam není
If ListView_Datove_Zpravy.ListItems.count > 0 Then
    Call nacteni_dat_Sonik_Datove_Zpravy_Prehled(False) 'Naèti všechny
End If
    
    
'Find the last non-blank cell in column A(1)
lRow = DM_Import.ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row

'Zjistim maximalni skp
            
sql_max_skp_DM = "select max(skp_data_message) as max_skp from " & schema & ".sonik_data_message"

Set rst1 = New ADODB.Recordset
rst1.Open sql_max_skp_DM, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

If Not rst1.EOF Then
 rst1.MoveFirst
     If IsNull(rst1![max_skp]) Then
     max_skp_DM = 0
     Else: max_skp_DM = rst1![max_skp]
     End If
 End If

rst1.Close

'Jenom pojistka pro první spuštìní
 If max_skp_DM >= 1 Then
 Else
    max_skp_DM = 0
 End If


'Projdu všechny buòky v prvním sloupci excelu
For i = 1 To lRow


    'Hledám èíslo datové zprávy, pokud jej ještì nemám a zároveò se tam objeví èíslo vìtší než 10000 - tak by to mìlo být právì ono
    If IsNumeric(DM_Import.ActiveSheet.Cells(i, "A").value) And DM_Import.ActiveSheet.Cells(i, "A").value >= 10000 And data_message_number = 0 Then
        data_message_number = Trim(DM_Import.ActiveSheet.Cells(i, "A").value)
        pom_zprava_vyctena = pom_zprava_vyctena + 1
    End If
    
    'Odesílatele poznám jen tak, že následuje po èísle datové zprávy
    If DM_Import.ActiveSheet.Cells(i, "A").value <> vbNullString And data_message_sender = vbNullString And data_message_delivery > 0 Then 'Následuje až po èísle DZ
        data_message_sender = DM_Import.ActiveSheet.Cells(i, "A").value
        pom_zprava_vyctena = pom_zprava_vyctena + 1
    End If
    
    'Datum doruèení poznávám podle toho, že je tam napsáno doruèeno
    If InStr(DM_Import.ActiveSheet.Cells(i, "A").value, "Doruèeno") > 0 And data_message_delivery = 0 Then
        data_message_delivery = Right(DM_Import.ActiveSheet.Cells(i, "A").value, Len(DM_Import.ActiveSheet.Cells(i, "A").value) - 8)
        pom_zprava_vyctena = pom_zprava_vyctena + 1
    End If
    
    'Spisová znaèka by mìla být ta hodnota následující po "O / è.j" - to je devìt znakù
    If Left(DM_Import.ActiveSheet.Cells(i, "A").value, 7) = "O / è.j" And text_file_number = vbNullString Then
        text_file_number = Mid(DM_Import.ActiveSheet.Cells(i, "A").value, 9)
        
        'Dodateèné oèištìní...
        If InStr(text_file_number, " - null") > 0 Then 'Pokud se v tom textu vyskytuje null
            text_file_number = Left(text_file_number, InStr(text_file_number, " - null") - 1) 'Tak vezmu èást pøed pomlèkou
        ElseIf InStr(text_file_number, " / ") > 0 And Left(text_file_number, 4) = Mid(text_file_number, InStr(text_file_number, " / ") + 3, 4) Then 'Pokud se tam vyskytuje / s mezerami okolo a Pokud je text pøed a po lomítku stejný (staèí zaèátek), tak vezmu jen to pøed lomítkem
            text_file_number = Left(text_file_number, InStr(text_file_number, " / ") - 1) 'Tak vezmu èást pøed lomítkem
        End If

        
        pom_zprava_vyctena = pom_zprava_vyctena + 1
    End If
    
    
    'Vždy si uložím do promìnné, pokud najdu jednu vlastnost DZ, když už mám všechny 4, tak zaèínám ukládat do databáze
    If pom_zprava_vyctena = 4 Then
     
     
           'Kontrola na datové formáty
           
           
           
               'Kontroly, jestli uživatel zadal správný formát - èíslo tam kam mìl
               
                'Datum pøijetí zprávy - promìnná je typu date, takže by to kdyžtak mìlo spadnout už døíve....
'                If Not IsDate(data_message_delivery) Then
'                         MsgBox "Neplatný formát data obdržení zprávy na " & i & ". øádku.", vbExclamation, "Chybnì zadaná data"
'                Exit Sub
'                End If
                
                'ID datové zprávy - promìnná je typu Long, takže by to kdyžtak mìlo spadnout už døíve....
                If Not IsNumeric(data_message_number) Then
                    MsgBox "Neplatný formát ID datové zprávy na " & i & ". øádku.", vbExclamation, "Chybnì zadaná data"
                    Exit Sub
                End If
                
                'Odesílatel
                If Not Application.IsText(data_message_sender) Then
                    MsgBox "Neplatný formát odesílatele na " & i & ". øádku.", vbExclamation, "Chybnì zadaná data"
                    Exit Sub
                End If
           
     
                'Kontroly, jestli náhodou datová zpráva už jednou zaevidována nebyla - pro každou novì zadávanou zprávu projedu listview s už zaevidovanými a pokud tam už je tak shodím proceduru
                 If Sonik_Datove_Zpravy_Prehled.ListView_Datove_Zpravy.ListItems.count > 0 Then
                     
                     For j = 1 To Sonik_Datove_Zpravy_Prehled.ListView_Datove_Zpravy.ListItems.count
                     
                             If IsNumeric(Sonik_Datove_Zpravy_Prehled.ListView_Datove_Zpravy.ListItems(j).SubItems(1)) Then
                                If Sonik_Datove_Zpravy_Prehled.ListView_Datove_Zpravy.ListItems(j).SubItems(1) = data_message_number Then
                                     MsgBox "Datová zpráva s èíslem " & data_message_number & " již byla døíve zaevidována.", vbExclamation, "Chybnì zadaná data"
                                     Exit Sub
                                End If
                             End If
                     Next j
                 End If

     
     
     'Ulož do DTB
     
            'Ulozeni do databáze - Vkládám jak táznam samotné DZ, tak i její status do pøíslušné tabule
            
                sql_insert_DM = "insert into  " & schema & ".sonik_data_message values (" & schema & ".sonik_data_message_seq.nextval," & data_message_number & ",to_date('" & data_message_delivery & "','DD.MM.YYYY'),'" & data_message_sender & "','" & text_file_number & "', null, sysdate, '" & UCase(Main.username) & "')"
                sql_insert_DMS = "insert into  " & schema & ".sonik_data_message_status values (" & schema & ".sonik_data_message_status_seq.nextval," & max_skp_DM + CntRows_Import + 1 & ", 1, sysdate,'" & UCase(Main.username) & "', sysdate, '1.1.3000', 'Y')"
                
                
                'Nahrání upozornìní na zpracování 30 dní
                Time_Offset = 30
                sql_insert_DM_Notice_Automatic = "insert into " & schema & ".sonik_data_message_notice values (" & schema & ".sonik_data_message_notice_seq.nextval," & max_skp_DM + CntRows_Import + 1 & ",'Termín zpracování souèinnosti',sysdate,sysdate+ " & Time_Offset & " ,'Aktivní','sys',null,null)"
                
                
                Call execute(sql_insert_DM)
                
                'Zatím mi pro tyto testy staèí vkládat jen do data message
                Call execute(sql_insert_DMS)
                Call execute(sql_insert_DM_Notice_Automatic)



     
     'Promaž a hledej další zprávy
     CntRows_Import = CntRows_Import + 1
     pom_zprava_vyctena = 0
     data_message_number = 0
     data_message_delivery = 0
     data_message_sender = vbNullString
     text_file_number = vbNullString
     
    End If

Next i


    
DM_Import.Close SaveChanges:=False
Workbook_Open = 0




MsgBox "Do databáze bylo vloženo " & CntRows_Import & " datových zpráv."

Me.Show

Call nacteni_dat_Sonik_Datove_Zpravy_Prehled(True)

Exit Sub 'Ukonèení sub kvùli tomu, aby to nedošlo do error handleru

chyba:

        MsgBox "Jezevèík upadl a zavøe se, pošli prosím printscreen oblíbenému analytikovi. " & vbNewLine & Err.Description, vbCritical, "Kritická chyba"
        Debug.Assert (1 = 0) 'pokud ma uzivatel otevreny VBA editor tak se to tady zasekne
        'cs_HDWMN.RollbackTrans 'FA: ADODB nema property, kterou by slo zjistit jestli je otevrena nejaka transakce. Takhle dojde k erroru pokud zadna TRN neni otevrena, to asi nechceme.
        End 'End není end IFu ale ukonèí macro - break ètverec nahoøe v lištì


End Sub


'Pøidání nebo úprava spisové znaèky
Private Sub CommandButton_Cislo_Jednaci_Click()


If ListView_Datove_Zpravy.ListItems.count > 0 Then

SKP_datove_zpravy = ListView_Datove_Zpravy.SelectedItem.text
Sonik_Rucni_Cislo_Jednaci_DZ.Show
        
End If




End Sub


Private Sub TextBox_Cislo_DZ_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'Pokud stisknu enter, když jsem ve vyhledávacím políèku, hned se spustí vyhledávání (nemusím klikat na OK)

If KeyCode = 13 Then
    Call CommandButton_Hledej_Click
    TextBox_Cislo_DZ.SetFocus
End If

End Sub

Private Sub TextBox_Spisova_Znacka_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'Pokud stisknu enter, když jsem ve vyhledávacím políèku, hned se spustí vyhledávání (nemusím klikat na OK)

If KeyCode = 13 Then
    Call CommandButton_Hledej_Click
    TextBox_Spisova_Znacka.SetFocus
End If

End Sub

Private Sub TextBox_Cislo_Jednaci_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'Pokud stisknu enter, když jsem ve vyhledávacím políèku, hned se spustí vyhledávání (nemusím klikat na OK)

If KeyCode = 13 Then
    Call CommandButton_Hledej_Click
    TextBox_Cislo_Jednaci.SetFocus
End If

End Sub

Private Sub TextBox_DateDelivery_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'Pokud stisknu enter, když jsem ve vyhledávacím políèku, hned se spustí vyhledávání (nemusím klikat na OK)

If KeyCode = 13 Then
    Call CommandButton_Hledej_Click
    TextBox_DateDelivery.SetFocus
End If

End Sub

Private Sub TextBox_CUID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'Pokud stisknu enter, když jsem ve vyhledávacím políèku, hned se spustí vyhledávání (nemusím klikat na OK)

If KeyCode = 13 Then
    Call CommandButton_Hledej_Click
    TextBox_CUID.SetFocus
End If

End Sub

Private Sub TextBox_DateBirth_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'Pokud stisknu enter, když jsem ve vyhledávacím políèku, hned se spustí vyhledávání (nemusím klikat na OK)

If KeyCode = 13 Then
    Call CommandButton_Hledej_Click
    TextBox_DateBirth.SetFocus
End If

End Sub



Private Sub CommandButton_Hledej_Click()

'Kontroly, jestli uživatel zadal správný formát - èíslo tam kam mìl
'CUID
If Not IsNumeric(TextBox_CUID.value) And TextBox_CUID.value <> vbNullString Then
         MsgBox "Zkus to prosím znovu, zadej do cuidu pouze èíselné hodnoty", vbExclamation, "Chybné zadání hodnot"
         TextBox_CUID.value = vbNullString
Exit Sub
End If


'Samotné hledání

Dim Hl_DZ_Cislo_Datove_Zpravy As String
Dim Hl_DZ_Cislo_Jednaci As String
Dim Hl_DZ_Spisova_Znacka As String
Dim Hl_DZ_Datum_Doruceni As Date

Dim Hl_DZ_Datum_Narozeni As Date
Dim Hl_DZ_CUID As String

Dim Hl_DZ_Note As String 'RISK-1014

Dim sql_dz As String

Hl_DZ_Cislo_Datove_Zpravy = Trim(TextBox_Cislo_DZ.value)
Hl_DZ_Cislo_Jednaci = TextBox_Cislo_Jednaci.value
Hl_DZ_Spisova_Znacka = TextBox_Spisova_Znacka.value
 
If IsDate(TextBox_DateDelivery.value) Then
Hl_DZ_Datum_Doruceni = TextBox_DateDelivery.value
Else: Hl_DZ_Datum_Doruceni = "1.1.3000"
End If
If IsDate(TextBox_DateBirth.value) Then
Hl_DZ_Datum_Narozeni = TextBox_DateBirth.value
Else: Hl_DZ_Datum_Narozeni = "1.1.3000"
End If

Hl_DZ_CUID = TextBox_CUID.value


'Neudìlám nic, pokud uživatel nezadá nic
If Hl_DZ_Cislo_Datove_Zpravy = vbNullString And Hl_DZ_Cislo_Jednaci = vbNullString And Hl_DZ_Spisova_Znacka = vbNullString And Hl_DZ_Datum_Doruceni = "1.1.3000" And Hl_DZ_CUID = vbNullString And Hl_DZ_Datum_Narozeni = "1.1.3000" And ComboBox_Stav_DZ.text = vbNullString And TextBox_Note.text = vbNullString Then 'RISK-1014
   MsgBox "Zkus to prosím znovu, žádný atribut k hledání nebyl zadán", vbExclamation, "Chybné zadání hodnot"
Exit Sub
End If




If Hl_DZ_Cislo_Datove_Zpravy <> vbNullString Then
   sql_dz = "select SKP_DATA_MESSAGE, DATA_MESSAGE_NUMBER, date_data_message_delivery, DATA_MESSAGE_SENDER, TEXT_FILE_NUMBER, TEXT_REFERENCE_NUMBER, DTIME_INSERTED, NAME_EMPLOYEE_INSERTED, NAME_DATA_MESSAGE_STATUS from " & schema & ".sonik_v_data_message where DATA_MESSAGE_NUMBER like '%" & Hl_DZ_Cislo_Datove_Zpravy & "%'"

   ElseIf Hl_DZ_Cislo_Jednaci <> vbNullString Then
   sql_dz = "select SKP_DATA_MESSAGE, DATA_MESSAGE_NUMBER, date_data_message_delivery, DATA_MESSAGE_SENDER, TEXT_FILE_NUMBER, TEXT_REFERENCE_NUMBER, DTIME_INSERTED, NAME_EMPLOYEE_INSERTED, NAME_DATA_MESSAGE_STATUS from " & schema & ".sonik_v_data_message where lower(convert(TEXT_REFERENCE_NUMBER,'US7ASCII')) like lower(convert('%" & Hl_DZ_Cislo_Jednaci & "%','US7ASCII'))"

   ElseIf Hl_DZ_Spisova_Znacka <> vbNullString Then
   sql_dz = "select SKP_DATA_MESSAGE, DATA_MESSAGE_NUMBER, date_data_message_delivery, DATA_MESSAGE_SENDER, TEXT_FILE_NUMBER, TEXT_REFERENCE_NUMBER, DTIME_INSERTED, NAME_EMPLOYEE_INSERTED, NAME_DATA_MESSAGE_STATUS from " & schema & ".sonik_v_data_message where lower(convert(TEXT_FILE_NUMBER,'US7ASCII')) like lower(convert('%" & Hl_DZ_Spisova_Znacka & "%','US7ASCII'))"

   ElseIf Hl_DZ_Datum_Doruceni <> "1.1.3000" Then
   sql_dz = "select SKP_DATA_MESSAGE, DATA_MESSAGE_NUMBER, date_data_message_delivery, DATA_MESSAGE_SENDER, TEXT_FILE_NUMBER, TEXT_REFERENCE_NUMBER, DTIME_INSERTED, NAME_EMPLOYEE_INSERTED, NAME_DATA_MESSAGE_STATUS from " & schema & ".sonik_v_data_message where trunc(date_data_message_delivery) = to_date('" & Hl_DZ_Datum_Doruceni & "','DD.MM.YYYY')"

   ElseIf Hl_DZ_CUID <> vbNullString Then
   sql_dz = "select SKP_DATA_MESSAGE, DATA_MESSAGE_NUMBER, date_data_message_delivery, DATA_MESSAGE_SENDER, TEXT_FILE_NUMBER" & _
            " , TEXT_REFERENCE_NUMBER, DTIME_INSERTED, NAME_EMPLOYEE_INSERTED, NAME_DATA_MESSAGE_STATUS" & _
            " from " & schema & ".sonik_v_data_message" & _
            " where skp_data_message in (select skp_data_message from " & schema & ".sonik_confirm_pair where NAME_CONFIRMATION_STATUS = 'Potvrzeno' and ID_CUID = '" & Hl_DZ_CUID & "')"
   
   ElseIf Hl_DZ_Datum_Narozeni <> "1.1.3000" Then
   sql_dz = "select SKP_DATA_MESSAGE, DATA_MESSAGE_NUMBER, date_data_message_delivery, DATA_MESSAGE_SENDER, TEXT_FILE_NUMBER" & _
            " , TEXT_REFERENCE_NUMBER, DTIME_INSERTED, NAME_EMPLOYEE_INSERTED, NAME_DATA_MESSAGE_STATUS" & _
            " from " & schema & ".sonik_v_data_message" & _
            " where skp_data_message in (select skp_data_message from " & schema & ".sonik_confirm_pair where NAME_CONFIRMATION_STATUS = 'Potvrzeno' and DATE_BIRTH = to_date('" & Hl_DZ_Datum_Narozeni & "','DD.MM.YYYY'))"
  
  'Dle stavu datové zprávy
   ElseIf ComboBox_Stav_DZ.text <> vbNullString Then
   sql_dz = "select SKP_DATA_MESSAGE, DATA_MESSAGE_NUMBER, date_data_message_delivery, DATA_MESSAGE_SENDER, TEXT_FILE_NUMBER, TEXT_REFERENCE_NUMBER, DTIME_INSERTED, NAME_EMPLOYEE_INSERTED, NAME_DATA_MESSAGE_STATUS from " & schema & ".sonik_v_data_message where NAME_DATA_MESSAGE_STATUS = '" & ComboBox_Stav_DZ.text & "'"
   
   'Dle poznámky v datové zprávì
   ElseIf TextBox_Note.text <> vbNullString Then 'RISK-1014
   sql_dz = "select SKP_DATA_MESSAGE, DATA_MESSAGE_NUMBER, date_data_message_delivery, DATA_MESSAGE_SENDER, TEXT_FILE_NUMBER, TEXT_REFERENCE_NUMBER, DTIME_INSERTED, NAME_EMPLOYEE_INSERTED, NAME_DATA_MESSAGE_STATUS from " & schema & ".sonik_v_data_message" & _
            " WHERE skp_data_message IN (SELECT skp_data_message FROM " & schema & ".sonik_data_message_note" & _
                                       " WHERE  " & schema & ".IGNORUJ_DIAK(text_note) LIKE " & schema & ".IGNORUJ_DIAK('%' || '" & TextBox_Note.text & "' || '%'))"

   


End If









'Pokud jsem až do teï nebyl schopnej naplnit hledací select, tak napíšu chybu a ukonèím proceduru
If sql_dz = vbNullString Then
MsgBox "Zkus to prosím znovu, nezadal jsi dostateèný poèet hodnot k vyhledávání datové zprávy", vbExclamation, "Chybné zadání hodnot"
Exit Sub
End If

'Jinak spustim query
Call query_datove_zpravy(sql_dz)


End Sub








Private Sub CommandButton_Zrus_Hledani_Click()
'Tlaèítko zruš hledání

Me.TextBox_Cislo_DZ.text = ""
Me.TextBox_Cislo_Jednaci.text = ""
Me.TextBox_Spisova_Znacka.text = ""
Me.TextBox_DateDelivery.value = ""
Me.TextBox_DateBirth.value = ""
Me.TextBox_CUID.text = ""
Me.TextBox_Note.text = "" 'RISK-1014
Me.ComboBox_Stav_DZ.text = ""


'A obnovim listview

Call nacteni_dat_Sonik_Datove_Zpravy_Prehled(True)

'   sql_dz = "select SKP_DATA_MESSAGE, DATA_MESSAGE_NUMBER, date_data_message_delivery, DATA_MESSAGE_SENDER, TEXT_FILE_NUMBER, TEXT_REFERENCE_NUMBER, DTIME_INSERTED, NAME_EMPLOYEE_INSERTED, NAME_DATA_MESSAGE_STATUS from " & schema & ".sonik_v_data_message"
'   Call query_datove_zpravy(sql_dz)

End Sub




' Zavøení formuláøe
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Call set_filter(Me.name, 1)
Me.Hide
Main.Show
End Sub


' Back button
Private Sub CommandButton_Navrat_Click()
Call text_information_hide(formular)
Call UserForm_QueryClose(1, 0)
End Sub















'
'
'' Zmìna barvy po najetí myši **************************************************************************************************************************
'
'Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'CommandButton_Hledej.ForeColor = barvaCerna
'CommandButton_Zrus_Hledani.ForeColor = barvaCerna
'End Sub
'
'Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'CommandButton_Navrat.ForeColor = barvaCerna
'End Sub
'
'Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'CommandButton_Hledej.ForeColor = barvaCerna
'CommandButton_Zrus_Hledani.ForeColor = barvaCerna
'End Sub
'
'Private Sub CommandButton_Navrat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'CommandButton_Navrat.ForeColor = barvaZelenaAirBank
'End Sub
'
'Private Sub CommandButton_Hledej_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'CommandButton_Hledej.ForeColor = barvaZelenaAirBank
'End Sub
'
'Private Sub CommandButton_Zrus_Hledani_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'CommandButton_Zrus_Hledani.ForeColor = RGB(219, 41, 41)
'End Sub
