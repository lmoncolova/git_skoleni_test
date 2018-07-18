VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sonik_Datova_Zprava 
   Caption         =   "Soník - Datová zpráva"
   ClientHeight    =   11475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   22560
   OleObjectBlob   =   "Sonik_Datova_Zprava.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sonik_Datova_Zprava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Option Explicit


Private Sub CommandButton_Missing_Aggreement_Click()

Dim sablona As String
Dim podpis As String
Dim cislojednaci As String

    Dim objWord As Object
    Dim objWordDocument As Object
    Dim objStory As Object
    Dim objToc As Object

sablona = "W:\Vymahani\LATE\Vzory\Soník Chybìjící souhlas státního zastupitelství NEMAZAT.dotx"

''RISK-641 - úprava podpisu podle toho, kdo generuje šablonu
''Šablona Lucka:
'If UCase(Main.username) = "AIR_KASPAROVA" Then
'    podpis = "Lucie Kašparová / Specialistka souèinnosti s externími subjekty" & vbCrLf & "Air Bank a.s. / èlen skupiny PPF" & vbCrLf & "+420 547 134 816 / soucinnost@airbank.cz" & vbCrLf & vbCrLf & "Markéta Doubková / Specialistka souèinnosti s externími subjekty" & vbCrLf & "Air Bank a.s. / èlen skupiny PPF" & vbCrLf & "+420 547 134 820 / soucinnost@airbank.cz"
'
''Šablona Markéta:
'ElseIf UCase(Main.username) = "AIR_DOUBKOVA" Then
'    podpis = "Markéta Doubková / Specialistka souèinnosti s externími subjekty" & vbCrLf & "Air Bank a.s. / èlen skupiny PPF" & vbCrLf & "+420 547 134 820 / soucinnost@airbank.cz" & vbCrLf & vbCrLf & "Lucie Kašparová / Specialistka souèinnosti s externími subjekty" & vbCrLf & "Air Bank a.s. / èlen skupiny PPF" & vbCrLf & "+420 547 134 816 / soucinnost@airbank.cz"
'
''Šablona Marka:
'ElseIf UCase(Main.username) = "AIR_HALCIN" Then
'    podpis = "Marek Halcin / Specialista souèinnosti s externími subjekty" & vbCrLf & "Air Bank a.s. / èlen skupiny PPF" & vbCrLf & "+420 547 134 855 / soucinnost@airbank.cz" & vbCrLf & vbCrLf & "Lucie Kašparová / Specialistka souèinnosti s externími subjekty" & vbCrLf & "Air Bank a.s. / èlen skupiny PPF" & vbCrLf & "+420 547 134 816 / soucinnost@airbank.cz"
'
'Else: podpis = "Lucie Kašparová / Specialistka souèinnosti s externími subjekty" & vbCrLf & "Air Bank a.s. / èlen skupiny PPF" & vbCrLf & "+420 547 134 816 / soucinnost@airbank.cz" & vbCrLf & vbCrLf & "Markéta Doubková / Specialistka souèinnosti s externími subjekty" & vbCrLf & "Air Bank a.s. / èlen skupiny PPF" & vbCrLf & "+420 547 134 820 / soucinnost@airbank.cz"
'
'
'End If
'RISK-1217
podpis = SonikSignature

cislojednaci = Label_DZ_Reference_NUM.Caption



    Set objWord = CreateObject("Word.Application")

    With objWord

        'otevøení dokumentu
        'Set objWordDocument = .Documents.Open(ThisWorkbook.Path & _
            "\Dokument.docx")

        'otevøení dokumentu postaveném na šablonì
        On Error GoTo chyba_otevreni_sablony
        Set objWordDocument = .Documents.Add(Template:=sablona)
        On Error GoTo 0
        

       With objWordDocument
            .bookmarks("Podpisy").Select
        End With
        .Selection.text = podpis
        
        With objWordDocument
            .bookmarks("cislo_jednaci").Select
        End With
        .Selection.text = cislojednaci
        
        
        With objWordDocument
            .bookmarks("dnesni_datum").Select
        End With
        .Selection.text = format$(date, "d. m. yyyy")
        
        'videlná aplikace Word
        .Visible = True
        'pøedání fokusu
        .Activate

        'uvolnìní z pamìti
        Set objWordDocument = Nothing
        Set objWord = Nothing

    End With
    Exit Sub
    
chyba_otevreni_sablony:

    If (Err.Number = 13 And Err.Description = "Type mismatch") Or (Err.Number = 5981) Or (Err.Description = "Application-defined or object-defined error") Then
        MsgBox "jez-00013: Vypadá to, že nemáš povolené zmìny ve vzoru šablony. Je potøeba šablonu otevøít a zakliknout povolení úprav (v žluté lištì navrchu). " & vbCrLf & _
               "Jedná se o tenhle soubor: " & sablona & vbCrLf & _
               "> v pøípadì, že šablona (.docx soubor) nelze spustit, je potøeba pøes správce úloh ukonèit všechny procesy WINWORD.EXE.", vbInformation, "Povolení úprav šablony"
    Else
        MsgBox "Došlo k neznámé chybì pøi naèítání vzoru šablony."
    End If
    
End Sub

Private Sub CommandButton_Negativni_Sablona_Click()

Dim text_osoby As String
Dim text_osoby_word As String

If Sonik_Datova_Zprava.ListView_Negativky.ListItems.count = 0 Then
 
MsgBox "Nenalezen žádný negativní subjekt souèinnosti - šablonu nelze vygenerovat", vbExclamation, "Šablonu nelze vygenerovat"
 
Exit Sub
 
End If
 
 
 


Dim sablona As String
Dim podpis As String
Dim cislojednaci As String
Dim spisovaznacka As String

    Dim objWord As Object
    Dim objWordDocument As Object
    Dim objStory As Object
    Dim objToc As Object

sablona = "W:\Vymahani\LATE\Vzory\Soník Negativní souèinnost NEMAZAT.dotx"

''RISK-641 - úprava podpisu podle toho, kdo generuje šablonu
''Šablona Lucka:
'If UCase(Main.username) = "AIR_KASPAROVA" Then
'    podpis = "Lucie Kašparová / Specialistka souèinnosti s externími subjekty" & vbCrLf & "Air Bank a.s. / èlen skupiny PPF" & vbCrLf & "+420 547 134 816 / soucinnost@airbank.cz" & vbCrLf & vbCrLf & "Markéta Doubková / Specialistka souèinnosti s externími subjekty" & vbCrLf & "Air Bank a.s. / èlen skupiny PPF" & vbCrLf & "+420 547 134 820 / soucinnost@airbank.cz"
'
''Šablona Markéta:
'ElseIf UCase(Main.username) = "AIR_DOUBKOVA" Then
'    podpis = "Markéta Doubková / Specialistka souèinnosti s externími subjekty" & vbCrLf & "Air Bank a.s. / èlen skupiny PPF" & vbCrLf & "+420 547 134 820 / soucinnost@airbank.cz" & vbCrLf & vbCrLf & "Lucie Kašparová / Specialistka souèinnosti s externími subjekty" & vbCrLf & "Air Bank a.s. / èlen skupiny PPF" & vbCrLf & "+420 547 134 816 / soucinnost@airbank.cz"
'
''Šablona Marka:
'ElseIf UCase(Main.username) = "AIR_HALCIN" Then
'    podpis = "Marek Halcin / Specialista souèinnosti s externími subjekty" & vbCrLf & "Air Bank a.s. / èlen skupiny PPF" & vbCrLf & "+420 547 134 855 / soucinnost@airbank.cz" & vbCrLf & vbCrLf & "Lucie Kašparová / Specialistka souèinnosti s externími subjekty" & vbCrLf & "Air Bank a.s. / èlen skupiny PPF" & vbCrLf & "+420 547 134 816 / soucinnost@airbank.cz"
'
'Else: podpis = "Lucie Kašparová / Specialistka souèinnosti s externími subjekty" & vbCrLf & "Air Bank a.s. / èlen skupiny PPF" & vbCrLf & "+420 547 134 816 / soucinnost@airbank.cz" & vbCrLf & vbCrLf & "Markéta Doubková / Specialistka souèinnosti s externími subjekty" & vbCrLf & "Air Bank a.s. / èlen skupiny PPF" & vbCrLf & "+420 547 134 820 / soucinnost@airbank.cz"
'
'
'End If
'RISK-1217
podpis = SonikSignature

spisovaznacka = Label_DZ_File_NUM.Caption
cislojednaci = Label_DZ_Reference_NUM.Caption


For i = 1 To Sonik_Datova_Zprava.ListView_Negativky.ListItems.count
    'RISK-876
    If ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "ID Subjektu", i) <> vbNullString Then 'SKP_SUBJECT
        text_osoby = text_osoby & "      " & Chr(150) & "    "
    End If
    
    If ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Jméno", i) <> vbNullString And ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Pøíjmení", i) <> vbNullString Then 'Jméno a pøíjmení
        text_osoby = text_osoby & ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Jméno", i) & " " & ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Pøíjmení", i)
    End If
    
    If ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Datum narození", i) <> vbNullString Then 'datum narození
        If text_osoby <> "      " & Chr(150) & "    " Then ' Už je nìjaká informace vyplnìna
              text_osoby = text_osoby & ", datum narození " & ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Datum narození", i)
        Else: text_osoby = text_osoby & "Datum narození " & ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Datum narození", i)
        End If
    End If
    
    If ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Rodné èíslo", i) <> vbNullString Then 'rodné èíslo
        If text_osoby <> "      " & Chr(150) & "    " Then ' Už je nìjaká informace vyplnìna
              text_osoby = text_osoby & ", rodné èíslo " & ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Rodné èíslo", i)
        Else: text_osoby = text_osoby & "Rodné èíslo " & ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Rodné èíslo", i)
        End If
    End If
    
    If ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Èíslo úètu", i) <> vbNullString Then 'èíslo úètu
        If text_osoby <> "      " & Chr(150) & "    " Then ' Už je nìjaká informace vyplnìna
              text_osoby = text_osoby & ", èíslo úètu " & ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Èíslo úètu", i) & "/3030"
        Else: text_osoby = text_osoby & "Èíslo úètu " & ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Èíslo úètu", i) & "/3030"
        End If
    End If
    
    If ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Èíslo karty", i) <> vbNullString Then 'èíslo karty
        If text_osoby <> "      " & Chr(150) & "    " Then ' Už je nìjaká informace vyplnìna
              text_osoby = text_osoby & ", èíslo karty " & ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Èíslo karty", i)
        Else: text_osoby = text_osoby & "Èíslo karty " & ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Èíslo karty", i)
        End If
    End If
    
    If ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Èíslo úvìru", i) <> vbNullString Then 'èíslo úvìru
        If text_osoby <> "      " & Chr(150) & "    " Then ' Už je nìjaká informace vyplnìna
              text_osoby = text_osoby & ", èíslo úvìru " & ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Èíslo úvìru", i)
        Else: text_osoby = text_osoby & "Èíslo úvìru " & ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Èíslo úvìru", i)
        End If
    End If
    
    If ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Èíslo telefonu", i) <> vbNullString Then 'telefonní èíslo
        If text_osoby <> "      " & Chr(150) & "    " Then ' Už je nìjaká informace vyplnìna
              text_osoby = text_osoby & ", èííslo telefonu " & ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Èíslo telefonu", i)
        Else: text_osoby = text_osoby & "Èíslo telefonu " & ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Negativky, "Èíslo telefonu", i)
        End If
    End If
    
    text_osoby_word = text_osoby_word & text_osoby & "" & Chr(10)
    
    text_osoby = vbNullString
Next i




    Set objWord = CreateObject("Word.Application")

    With objWord

        'otevøení dokumentu
        'Set objWordDocument = .Documents.Open(ThisWorkbook.Path & _
            "\Dokument.docx")

        'otevøení dokumentu postaveném na šablonì
         
        On Error GoTo chyba_otevreni_sablony
        Set objWordDocument = .Documents.Add(Template:=sablona)
        On Error GoTo 0
        

       With objWordDocument
            .bookmarks("Podpisy").Select
        End With
        .Selection.text = podpis
        
       With objWordDocument
            .bookmarks("osoby").Select
        End With
        .Selection.text = text_osoby_word
        
       With objWordDocument
            .bookmarks("zadost_od").Select
        End With
        .Selection.text = SklonujSZ2Pad(Odesilatel_datove_zpravy) 'RISK-978
        
        With objWordDocument
            .bookmarks("cislo_jednaci").Select
        End With
        .Selection.text = cislojednaci
        
        With objWordDocument
            .bookmarks("cislo_jednaci_2").Select
        End With
        .Selection.text = spisovaznacka
        
        With objWordDocument
            .bookmarks("dnesni_datum").Select
        End With
        .Selection.text = format$(date, "d. m. yyyy")
        
        'videlná aplikace Word
        .Visible = True
        'pøedání fokusu
        .Activate

        'uvolnìní z pamìti
        Set objWordDocument = Nothing
        Set objWord = Nothing

    End With
    Exit Sub

chyba_otevreni_sablony:

    If (Err.Number = 13 And Err.Description = "Type mismatch") Or (Err.Number = 5981) Or (Err.Description = "Application-defined or object-defined error") Then
        MsgBox "jez-00013: Vypadá to, že nemáš povolené zmìny ve vzoru šablony. Je potøeba šablonu otevøít a zakliknout povolení úprav (v žluté lištì navrchu). " & vbCrLf & _
               "Jedná se o tenhle soubor: " & sablona & vbCrLf & _
               "> v pøípadì, že šablona (.docx soubor) nelze spustit, je potøeba pøes správce úloh ukonèit všechny procesy WINWORD.EXE.", vbInformation, "Povolení úprav šablony"
    Else
        MsgBox "Došlo k neznámé chybì pøi naèítání vzoru šablony."
    End If

End Sub



'RISK-1094
Private Sub CommandButton_UlozToDo_Click()

Dim sql As String, StrToDoNote As String

'Naètení poznámky
'Jde o to, že na této kartì mìním pouze uživatele a ne poznámku, takže ji potøebuji odnìkud naèíst, abych mohl udìlat insert
sql = "SELECT  t.text_note " & _
      " FROM " & schema & ".sonik_data_message_todo_list t " & _
      " WHERE t.flag_valid = 'Y' AND t.flag_deleted = 'N'  " & _
      " AND t.SKP_DATA_MESSAGE = " & SKP_datove_zpravy

 
Set rst1 = New ADODB.Recordset
rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

If Not rst1.EOF Then 'Oèekávám jen jeden øádek nebo null
        StrToDoNote = "" & rst1![TEXT_NOTE] & ""
Else: StrToDoNote = ""
 
End If

rst1.Close

Call SonikSaveToDoListChange(SKP_datove_zpravy, UCase(ComboBoxAssignee.text), StrToDoNote) 'Assignee je nová mìnìná hodnota a ToDoNote je pùvodní, na této kartì se nemùže mìnit



End Sub

Private Sub Label_Reference_Num_Pointer_Click()

userform_prv = "Sonik_Datova_Zprava"

Sonik_Rucni_Cislo_Jednaci_DZ.Show

End Sub



Private Sub ListView_Paired_Subjects_DblClick()

'Dovolím klikání, jenom pokud tam je na co klikat, jinak to spadne
If ListView_Paired_Subjects.ListItems.count > 0 Then

userform_prv = "Sonik_Datova_Zprava"

Me.Hide

'RISK-721
Sonik_Karta_Klienta.Show
Call nacteni_dat_Karta_Klienta
Load Sonik_Karta_Klienta

End If





End Sub



Private Sub UserForm_Activate()
'Call nacteni_dat_Sonik_Datova_Zprava



Me.TextBox_Poznamka = ""

 Label_Flag_Copy_Values.Visible = False
 Label_Copy_Values.Visible = False


'Barvièky pokud se v napárovaných subjektech jedná o neklienta
'RISK-992 barvièky, pokud byl klient pøidán jako disponent nebo majitel úètu, ke kterému je hledaný disponentem

If Sonik_Datova_Zprava.ListView_Paired_Subjects.ListItems.count > 0 Then

    For i = 1 To Sonik_Datova_Zprava.ListView_Paired_Subjects.ListItems.count

        If ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Paired_Subjects, "SKP_PERSON", i) = -1 Then
        
            For j = 2 To Sonik_Datova_Zprava.ListView_Paired_Subjects.ColumnHeaders.count - 1
            Sonik_Datova_Zprava.ListView_Paired_Subjects.ListItems(i).ListSubItems.item(j - 1).ForeColor = RGB(125, 125, 125) 'RISK-992 šedá
            Next j
        
        End If
        
        If ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Paired_Subjects, "DISPONENT_RELATION", i) = "DISPONENT" Then
        
            For j = 2 To Sonik_Datova_Zprava.ListView_Paired_Subjects.ColumnHeaders.count - 1
            Sonik_Datova_Zprava.ListView_Paired_Subjects.ListItems(i).ListSubItems.item(j - 1).ForeColor = RGB(188, 177, 76) 'RISK-992 žlutá
            Next j
        
        ElseIf ListvZiskejHodnotu(Sonik_Datova_Zprava.ListView_Paired_Subjects, "DISPONENT_RELATION", i) = "OWNER" Then
        
            For j = 2 To Sonik_Datova_Zprava.ListView_Paired_Subjects.ColumnHeaders.count - 1
            Sonik_Datova_Zprava.ListView_Paired_Subjects.ListItems(i).ListSubItems.item(j - 1).ForeColor = RGB(76, 177, 188) 'RISK-992 tyrkysová
            Next j
        
        End If
       

    Next i

End If


Sonik_Datova_Zprava.TextBox_Phone.ControlTipText = "Telefonní èíslo zadávej ve formátu 606095623"

'RISK-992
Sonik_Datova_Zprava.ListView_Paired_Subjects.ControlTipText = "Význam barvièek:  šedá - cizí karta, žlutá - pøidán jako disponent, tyrkysová - pøidán jako vlastík"


'RISK-1094
Dim sql As String

'Naètení hodnot do rolovátka
Call FillAssigneSonikToDo(ComboBoxAssignee)

'Naètení aktuálního øešitele
sql = "SELECT su.USERNAME " & _
      " FROM " & schema & ".sonik_data_message_todo_list t " & _
      " JOIN " & schema & ".sonik_users su ON t.skp_sonik_user = su.skp_sonik_user  " & _
      "                                    AND t.flag_valid = 'Y'  " & _
      "                                    AND t.flag_deleted = 'N'  " & _
      " WHERE t.SKP_DATA_MESSAGE = " & SKP_datove_zpravy

 
Set rst1 = New ADODB.Recordset
rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

If Not rst1.EOF Then 'Oèekávám jen jeden øádek nebo null
        ComboBoxAssignee.text = "" & rst1![username] & ""
Else: ComboBoxAssignee.text = ""
 
End If

rst1.Close


End Sub



'Private Sub TextBox_CUID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
''Pokud stisknu enter, když jsem ve vyhledávacím políèku, hned se spustí vyhledávání (nemusím klikat na OK)
'
'If KeyCode = 13 Then
'    Call CommandButton_Hledej_Click
'    TextBox_CUID.SetFocus
'End If
'
'End Sub
'
'Private Sub TextBox_Jmeno_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
''Pokud stisknu enter, když jsem ve vyhledávacím políèku, hned se spustí vyhledávání (nemusím klikat na OK)
'
'If KeyCode = 13 Then
'    Call CommandButton_Hledej_Click
'    TextBox_Jmeno.SetFocus
'End If
'
'End Sub
'
'Private Sub TextBox_Prijmeni_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
''Pokud stisknu enter, když jsem ve vyhledávacím políèku, hned se spustí vyhledávání (nemusím klikat na OK)
'
'If KeyCode = 13 Then
'    Call CommandButton_Hledej_Click
'    TextBox_Prijmeni.SetFocus
'End If
'
'End Sub
'
'Private Sub TextBox_DateBirth_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
''Pokud stisknu enter, když jsem ve vyhledávacím políèku, hned se spustí vyhledávání (nemusím klikat na OK)
'
'If KeyCode = 13 Then
'    Call CommandButton_Hledej_Click
'    TextBox_DateBirth.SetFocus
'End If
'
'End Sub
'
'Private Sub TextBox_RC_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
''Pokud stisknu enter, když jsem ve vyhledávacím políèku, hned se spustí vyhledávání (nemusím klikat na OK)
'
'If KeyCode = 13 Then
'    Call CommandButton_Hledej_Click
'    TextBox_RC.SetFocus
'End If
'
'End Sub
'
'Private Sub TextBox_BankAccount_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
''Pokud stisknu enter, když jsem ve vyhledávacím políèku, hned se spustí vyhledávání (nemusím klikat na OK)
'
'If KeyCode = 13 Then
'    Call CommandButton_Hledej_Click
'    TextBox_BankAccount.SetFocus
'End If
'
'End Sub
'
'Private Sub TextBox_Card_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
''Pokud stisknu enter, když jsem ve vyhledávacím políèku, hned se spustí vyhledávání (nemusím klikat na OK)
'
'If KeyCode = 13 Then
'    Call CommandButton_Hledej_Click
'    TextBox_Card.SetFocus
'End If
'
'End Sub
'
'Private Sub TextBox_Loan_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
''Pokud stisknu enter, když jsem ve vyhledávacím políèku, hned se spustí vyhledávání (nemusím klikat na OK)
'
'If KeyCode = 13 Then
'    Call CommandButton_Hledej_Click
'    TextBox_Loan.SetFocus
'End If
'
'End Sub




'Tlaèítko hledej subjekty v DTB
Private Sub CommandButton_Hledej_Click()

'RISK-908
Dim Bln_Show_Error_Message As Boolean, Str_Hlaska As String, Sql_Employee As String

'Kontroly, jestli uživatel zadal správný formát - èíslo tam kam mìl
'CUID
If Not IsNumeric(TextBox_CUID.value) And TextBox_CUID.value <> vbNullString Then
         MsgBox "Zkus to prosím znovu, zadej do cuidu pouze èíselné hodnoty", vbExclamation, "Chybné zadání hodnot"
         TextBox_CUID.value = vbNullString
Exit Sub
End If

'RC
If Not IsNumeric(TextBox_RC.value) And TextBox_RC.value <> vbNullString Then
         MsgBox "Zkus to prosím znovu, zadej do rodného èísla pouze èíselné hodnoty", vbExclamation, "Chybné zadání hodnot"
         TextBox_RC.value = vbNullString
Exit Sub
ElseIf Not (Len(TextBox_RC.value) = 9 Or Len(TextBox_RC.value) = 10) And TextBox_RC.value <> vbNullString Then
         MsgBox "Zkus to prosím znovu, rodné èíslo musí obsahovat 9 nebo 10 znakù", vbExclamation, "Chybné zadání hodnot"
         TextBox_RC.value = vbNullString

Exit Sub
End If

'Cislo uctu
If Not IsNumeric(TextBox_BankAccount.value) And TextBox_BankAccount.value <> vbNullString Then
         MsgBox "Zkus to prosím znovu, zadej do èísla úètu pouze èíselné hodnoty", vbExclamation, "Chybné zadání hodnot"
         TextBox_BankAccount.value = vbNullString
Exit Sub
ElseIf Len(TextBox_BankAccount.value) <> 10 And TextBox_BankAccount.value <> vbNullString Then
         MsgBox "Zkus to prosím znovu, èíslo úètu musí obsahovat právì 10 znakù", vbExclamation, "Chybné zadání hodnot"
         TextBox_BankAccount.value = vbNullString

Exit Sub
End If


'Cislo karty
If Not IsNumeric(TextBox_Card.value) And TextBox_Card.value <> vbNullString Then
         MsgBox "Zkus to prosím znovu, zadej do èísla karty pouze èíselné hodnoty", vbExclamation, "Chybné zadání hodnot"
         TextBox_Card.value = vbNullString
Exit Sub
ElseIf Len(TextBox_Card.value) <> 16 And TextBox_Card.value <> vbNullString Then
         MsgBox "Zkus to prosím znovu, èíslo karty musí obsahovat právì 16 znakù", vbExclamation, "Chybné zadání hodnot"
         TextBox_Card.value = vbNullString

Exit Sub
End If
'Telefonni cislo RISK-876
If Not IsNumeric(TextBox_Phone.value) And TextBox_Phone.value <> vbNullString Then
         MsgBox "Zkus to prosím znovu, zadej do telefonního èísla pouze èíselné hodnoty", vbExclamation, "Chybné zadání hodnot"
         TextBox_Phone.value = vbNullString
Exit Sub
End If


'Pokud projde pøes kontroly tak zaènu hledat

'Prvnì zkusím jednoduchý pøípad hledání pøes CUID, pak se musí øešit kombinace hledání

Dim Hl_Subj_CUID As String
Dim Hl_name_first As String
Dim Hl_name_last As String
Dim Hl_date_birth As Date
Dim Hl_RC As String
Dim Hl_Card As String
Dim Hl_Loan As String
Dim Hl_Bankaccount As String
'RISK-876
Dim Hl_Phone As String
Dim sql_dtb_klient_part As String



Hl_Subj_CUID = TextBox_CUID.value
Hl_name_first = StrConv(TextBox_Jmeno.value, 3)
Hl_name_last = StrConv(TextBox_Prijmeni.value, 3)
 
If IsDate(TextBox_DateBirth) Then
Hl_date_birth = TextBox_DateBirth.value
Else
Hl_date_birth = "1.1.3000"
End If

Hl_RC = TextBox_RC.value
Hl_Bankaccount = TextBox_BankAccount.value
Hl_Card = TextBox_Card.value
Hl_Loan = TextBox_Loan.value
Hl_Phone = Replace(TextBox_Phone.value, " ", "")


'Neudìlám nic, pokud uživatel nezadá nic
If Hl_Subj_CUID = vbNullString And (Hl_name_first = vbNullString And Hl_name_last = vbNullString And Hl_date_birth = "1.1.3000" And Hl_RC = vbNullString And Hl_Bankaccount = vbNullString And Hl_Card = vbNullString And Hl_Loan = vbNullString And Hl_Phone = vbNullString) Then
   MsgBox "Zkus to prosím znovu, CUID nebo jiný atribut k hledání nebyl zadán", vbExclamation, "Chybné zadání hodnot"
Exit Sub
End If

Dim sql_max_skp As String

'Než zaènu hledat, tak uložím subjekt do DTB

'Zjistim maximalni skp - skute
sql_max_skp = "select nvl(max(skp_subject),0) as max_skp from " & schema & ".sonik_subject"

Dim max_skp As Long
Dim sql_insert_subject As String

   Set rst1 = New ADODB.Recordset
   rst1.Open sql_max_skp, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    
        If rst1![max_skp] >= 1 Then
           max_skp = rst1![max_skp]
        End If
    End If

   rst1.Close


'Jenom pojistka pro první spuštìní
If max_skp >= 1 Then
Else
max_skp = 0
End If
SKP_Subject = max_skp + 1

'RISK-876
sql_insert_subject = "insert into  " & schema & ".sonik_subject values (" & SKP_Subject & ", " & SKP_datove_zpravy & "," & "'" & Hl_Subj_CUID & "'" & "," & "'" & Hl_name_first & "'" & "," & "'" & Hl_name_last & "'" & "," & "to_date('" & Hl_date_birth & "','DD.MM.YYYY')" & "," & "'" & Hl_RC & "'" & "," & "'" & Hl_Bankaccount & "'" & "," & "'" & Hl_Card & "'" & "," & "'" & Hl_Loan & "'" & "," & "sysdate" & "," & "'" & UCase(Main.username) & "','" & Hl_Phone & "')"

Call execute(sql_insert_subject)


'RISK-876 - pouziju proto, abych nemusel kopirovat do vsech selectu kdyz se prida sloupec
sql_dtb_klient_part = "skp_person, ID_CUID, NAME_FIRST, NAME_LAST, DATE_BIRTH, TEXT_BIRTH_NUMBER, name_permanent_address_street ||' '|| text_permanent_address_num ||', '|| name_permanent_address_town ||', '|| code_permanent_address_zip as PERMANENT_ADDRESS, TEXT_BANK_ACCOUNT_NUMBER_AGG, CODE_LOAN_NUMBER_AGG, TEXT_PHONE_NUMBER_AGG"

Dim sql_hl_subj As String

'Vìtev hledání napøímo dle CUIDu
If Hl_Subj_CUID <> vbNullString Then
'RISK-876
sql_hl_subj = "select " & sql_dtb_klient_part & " from " & schema & ".sonik_ab_klient_dtb where ID_CUID = '" & Hl_Subj_CUID & "'"


End If


'-------------------------------------------------------------------------------------

'Vìtev hledání podle ostatních jednozaèných atributù



'RISK-876
If Hl_Subj_CUID = vbNullString And (Hl_Bankaccount <> vbNullString Or Hl_Card <> vbNullString Or Hl_RC <> vbNullString Or Hl_Loan <> vbNullString Or Hl_Phone <> vbNullString) Then
   
   'Èíslo úètu - jednoznaèný identifikátor
   If Hl_Bankaccount <> vbNullString Then
   sql_hl_subj = "select " & sql_dtb_klient_part & " from " & schema & ".sonik_ab_klient_dtb where TEXT_BANK_ACCOUNT_NUMBER_AGG like '%" & Hl_Bankaccount & "%'"

   'Èíslo karty - jednoznaèný identifikátor
   ElseIf Hl_Card <> vbNullString Then
   sql_hl_subj = "select " & sql_dtb_klient_part & " from " & schema & ".sonik_ab_klient_dtb where TEXT_CARD_NUMBER_AGG like '%" & Hl_Card & "%'"
   
   'Rodné èíslo - jednoznaèný identifikátor
   ElseIf Hl_RC <> vbNullString Then
   sql_hl_subj = "select " & sql_dtb_klient_part & " from " & schema & ".sonik_ab_klient_dtb where TEXT_BIRTH_NUMBER = '" & Hl_RC & "'"
   
   'Èíslo úvìru - jednoznaèný identifikátor
   ElseIf Hl_Loan <> vbNullString Then
   sql_hl_subj = "select " & sql_dtb_klient_part & " from " & schema & ".sonik_ab_klient_dtb where CODE_LOAN_NUMBER_AGG like upper('%" & Hl_Loan & "%')"
   
   'RISK-876 Telefonní - není sice jednoznaèný identifikátor ale pøedpokládám, že pokud budou mít èíslo, tak nebudou hledat nic jiného
   ElseIf Hl_Phone <> vbNullString Then
   sql_hl_subj = "select " & sql_dtb_klient_part & " from " & schema & ".sonik_ab_klient_dtb where TEXT_PHONE_NUMBER_AGG like '%" & Hl_Phone & "%'"
   
   End If



End If



'-------------------------------------------------------------------------------------

'Vìtev hledání podle nejednozaèných atributù jméno pøíjmení a datum narození

'Je 2^3 - 1 možností, které mohou nastat, ale reálnì mnohem míò, nìkteré možnosti vùbec nepøipouštím:

'RISK-876
If Hl_Subj_CUID = vbNullString And (Hl_name_first <> vbNullString Or Hl_name_last <> vbNullString Or Hl_date_birth <> "1.1.3000") Then

    'Vyplní všechno
    If Hl_name_first <> vbNullString And Hl_name_last <> vbNullString And Hl_date_birth <> "1.1.3000" Then
    sql_hl_subj = "select " & sql_dtb_klient_part & " from " & schema & ".sonik_ab_klient_dtb where replace(lower(convert(NAME_FIRST,'US7ASCII')),' ','') like replace(lower(convert('%" & Hl_name_first & "%','US7ASCII')),' ','') and replace(lower(convert(NAME_LAST,'US7ASCII')),' ','') like replace(lower(convert('%" & Hl_name_last & "%','US7ASCII')),' ','') and DATE_BIRTH = to_date('" & Hl_date_birth & "','DD.MM.YYYY')"

    'Vyplní Jméno a Pøíjmení
    ElseIf Hl_name_first <> vbNullString And Hl_name_last <> vbNullString And Hl_date_birth = "1.1.3000" Then
    sql_hl_subj = "select " & sql_dtb_klient_part & " from " & schema & ".sonik_ab_klient_dtb where replace(lower(convert(NAME_FIRST,'US7ASCII')),' ','') like replace(lower(convert('%" & Hl_name_first & "%','US7ASCII')),' ','') and replace(lower(convert(NAME_LAST,'US7ASCII')),' ','') like replace(lower(convert('%" & Hl_name_last & "%','US7ASCII')),' ','')"

    'Vyplní pøíjmení a datum narození
    ElseIf Hl_name_first = vbNullString And Hl_name_last <> vbNullString And Hl_date_birth <> "1.1.3000" Then
    sql_hl_subj = "select " & sql_dtb_klient_part & " from " & schema & ".sonik_ab_klient_dtb where replace(lower(convert(NAME_LAST,'US7ASCII')),' ','') like replace(lower(convert('%" & Hl_name_last & "%','US7ASCII')),' ','') and DATE_BIRTH = to_date('" & Hl_date_birth & "','DD.MM.YYYY')"

    'Vyplní jen datum narození
    ElseIf Hl_name_first = vbNullString And Hl_name_last = vbNullString And Hl_date_birth <> "1.1.3000" Then
    sql_hl_subj = "select " & sql_dtb_klient_part & " from " & schema & ".sonik_ab_klient_dtb where DATE_BIRTH = to_date('" & Hl_date_birth & "','DD.MM.YYYY')"
    
    'Vyplní jen pøíjmení
    ElseIf Hl_name_first = vbNullString And Hl_name_last <> vbNullString And Hl_date_birth = "1.1.3000" Then
    sql_hl_subj = "select " & sql_dtb_klient_part & " from " & schema & ".sonik_ab_klient_dtb where replace(lower(convert(NAME_LAST,'US7ASCII')),' ','') like replace(lower(convert('%" & Hl_name_last & "%','US7ASCII')),' ','')"
    


    End If




End If

'Pokud jsem až do teï nebyl schopnej naplnit hledací select, tak napíšu chybu a ukonèím proceduru
If sql_hl_subj = vbNullString Then
MsgBox "Zkus to prosím znovu, nezadal jsi dostateèný poèet hodnot k vyhledávání subjektu", vbExclamation, "Chybné zadání hodnot"
Exit Sub
End If


'RISK-908 - pokud je hledaný klient dle jména a pøíjmení mezi zamìstnanci, tak vyhodím varovací hlášku
If Hl_name_first <> vbNullString And Hl_name_last <> vbNullString Then

    Bln_Show_Error_Message = False 'Iniciálnì
    
    Sql_Employee = "select 1 as FLAG_EMPLOYEE from " & schema & ".employee_ab where name_first = '" & Hl_name_first & "' and name_last = '" & Hl_name_last & "'"

    Set rst1 = New ADODB.Recordset
    rst1.Open Sql_Employee, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

    If Not rst1.EOF Then
        Bln_Show_Error_Message = True 'Pokud nìkoho najdu, tak vyhodím hlášku
    End If

    rst1.Close

    If Bln_Show_Error_Message Then
    
        Str_Hlaska = "Dej si pozor, vyhledáváný subjekt " & Hl_name_first & " " & Hl_name_last & " byl nalezen mezi zamìstnanci Air Bank. Je tedy možné, že jej Soník nebude schopen najít."
        Call HlaskaUzivateli(Str_Hlaska)
    
        Call VypisVsechnyHlasky
    
    End If
End If

'Jinak spustim hledani
Call nacteni_dat_Sonik_Potvrzeni_Subjektu
Call query_hledani_subjektu_v_dtb(sql_hl_subj & "order by NAME_LAST, NAME_FIRST,ID_CUID")

Sonik_Potvrzeni_Subjektu.Show

End Sub






'Tlaèítko zruš hledání
Private Sub CommandButton_Zrus_Hledani_Click()
Me.TextBox_CUID.text = ""
Me.TextBox_Jmeno.text = ""
Me.TextBox_Prijmeni.text = ""
Me.TextBox_DateBirth.value = ""
Me.TextBox_RC.text = ""
Me.TextBox_BankAccount.text = ""
Me.TextBox_RC.text = ""
Me.TextBox_Loan.text = ""
End Sub







Private Sub CommandButton_Add_Info_Click()





Sonik_Rucni_Pridani_Info_DZ.TextBox_Add_CL.value = ""
Sonik_Rucni_Pridani_Info_DZ.TextBox_Add_CL.ControlTipText = "Pokud chceš pøidat novou informaci do èíselníku, tak zadej a zmáèkni tlaèítko rozšiø èíselník"

Call nacteni_dat_Info_DZ
Call nacteni_dat_Vybrane_Info_DZ
Sonik_Rucni_Pridani_Info_DZ.Show

End Sub


Private Sub CommandButton_Add_Status_Click()

Call nacteni_dat_Sonik_Update_Stavu_DZ

Sonik_Update_Stav_DZ.Show

End Sub

Private Sub CommandButton_Add_Note_Click()

Sonik_Poznamka_DZ.Show

End Sub

Private Sub CommandButton_Add_Notice_Click()

Flag_Nove_Upozorneni_DZ = 1
Sonik_Upozorneni_DZ.Show

End Sub


Private Sub ListView_DZ_Upozorneni_DblClick()

Flag_Nove_Upozorneni_DZ = 0
Sonik_Upozorneni_DZ.Show



End Sub




'Plneni dodateèné poznámky
Private Sub ListView_DZ_History_Click()

'Dovolím klikání, jenom pokud tam je na co klikat, jinak to spadne
If ListView_DZ_History.ListItems.count > 0 Then

   TextBox_Poznamka.text = ListView_DZ_History.SelectedItem.SubItems(8)

End If

End Sub






' Zavøení formuláøe
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Call set_filter(Me.name, 1)

'Promažu
Sonik_Datova_Zprava.TextBox_CUID.text = ""
Me.Hide

If userform_prv = "SouhrnNotice" Then
    Souhrn.Show
Else
    Sonik_Datove_Zpravy_Prehled.Show
End If
End Sub


' Back button
Private Sub CommandButton_Navrat_Click()
'Promažu
Sonik_Datova_Zprava.TextBox_CUID.text = ""
Call text_information_hide(formular)
Call UserForm_QueryClose(1, 0)
End Sub





'Kopírování textu

Private Sub Label_ID_DZ_Click()

If Label_ID_DZ.Caption <> vbNullString Then

Dim clipboard As MSForms.DataObject
 Set clipboard = New MSForms.DataObject
 clipboard.SetText Label_ID_DZ.Caption
 clipboard.PutInClipboard
 
 Label_Copy_Values.Caption = "Interní ID zkopírováno do schránky"
 Label_Flag_Copy_Values.Visible = True
 Label_Copy_Values.Visible = True
 
End If
 
End Sub

Private Sub Label_DZ_Number_Click()

If Label_Dz_Number.Caption <> vbNullString Then

Dim clipboard As MSForms.DataObject
 Set clipboard = New MSForms.DataObject
 clipboard.SetText Label_Dz_Number.Caption
 clipboard.PutInClipboard
 
 Label_Copy_Values.Caption = "Èíslo datové zprávy zkopírováno do schránky"
 Label_Flag_Copy_Values.Visible = True
 Label_Copy_Values.Visible = True
 
End If
 
End Sub

Private Sub Label_DZ_Date_Delivery_Click()

If Label_DZ_Date_Delivery.Caption <> vbNullString Then

Dim clipboard As MSForms.DataObject
 Set clipboard = New MSForms.DataObject
 clipboard.SetText Label_DZ_Date_Delivery.Caption
 clipboard.PutInClipboard
 
 Label_Copy_Values.Caption = "Datum doruèení zkopírováno do schránky"
 Label_Flag_Copy_Values.Visible = True
 Label_Copy_Values.Visible = True
 
End If
 
End Sub

Private Sub Label_DZ_Sender_Click()

If Label_DZ_Sender.Caption <> vbNullString Then

Dim clipboard As MSForms.DataObject
 Set clipboard = New MSForms.DataObject
 clipboard.SetText Label_DZ_Sender.Caption
 clipboard.PutInClipboard
 
 Label_Copy_Values.Caption = "Odesílatel zkopírován do schránky"
 Label_Flag_Copy_Values.Visible = True
 Label_Copy_Values.Visible = True
 
End If
 
End Sub

Private Sub Label_DZ_File_NUM_Click()

If Label_DZ_File_NUM.Caption <> vbNullString Then

Dim clipboard As MSForms.DataObject
 Set clipboard = New MSForms.DataObject
 clipboard.SetText Label_DZ_File_NUM.Caption
 clipboard.PutInClipboard
 
 Label_Copy_Values.Caption = "Spisová znaèka zkopírována do schránky"
 Label_Flag_Copy_Values.Visible = True
 Label_Copy_Values.Visible = True
 
End If
 
End Sub

Private Sub Label_DZ_Reference_NUM_Click()

If Label_DZ_Reference_NUM.Caption <> vbNullString Then

Dim clipboard As MSForms.DataObject
 Set clipboard = New MSForms.DataObject
 clipboard.SetText Label_DZ_Reference_NUM.Caption
 clipboard.PutInClipboard
 
 Label_Copy_Values.Caption = "Jednací èíslo zkopírováno do schránky"
 Label_Flag_Copy_Values.Visible = True
 Label_Copy_Values.Visible = True
 
End If
 
End Sub




' Zmìna barvy po najetí myši **************************************************************************************************************************

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Hledej.ForeColor = barvaCerna
CommandButton_Zrus_Hledani.ForeColor = barvaCerna
End Sub

Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Navrat.ForeColor = barvaCerna
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Hledej.ForeColor = barvaCerna
CommandButton_Zrus_Hledani.ForeColor = barvaCerna
End Sub

Private Sub CommandButton_Navrat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Navrat.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Hledej_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Hledej.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Zrus_Hledani_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Zrus_Hledani.ForeColor = RGB(219, 41, 41)
End Sub

