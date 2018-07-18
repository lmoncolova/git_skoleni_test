VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sonik_Rucni_Pridani_Info_DZ 
   Caption         =   "Soník - Zadávání subjektù"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11010
   OleObjectBlob   =   "Sonik_Rucni_Pridani_Info_DZ.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sonik_Rucni_Pridani_Info_DZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit
Private itmX As Variant
Private Bln_Show_Date_Od As Boolean

Private Sub CommandButton_Add_CL_Click()

If TextBox_Add_CL.value = vbNullString Then

MsgBox "Musíš zadat text nové informace", vbExclamation, "Chybné zadání hodnot"

Exit Sub

End If

Dim sql_insert_Cl As String

If MsgBox("Opravdu chcete pøidat novou hodnotu do èíselníku?", vbYesNo, "Rozšíøení èíselníku hodnot") = vbYes Then

'Uložení nové hodnoty do èíselníku
sql_insert_Cl = "insert into  " & schema & ".sonik_cl_req_prov_info values (" & schema & ".sonik_cl_req_prov_info_seq.nextval ,'" & TextBox_Add_CL.value & "', sysdate, '" & UCase(Main.username) & "', sysdate, date'3000-01-01', 'Y')"

Call execute(sql_insert_Cl)

Call nacteni_dat_Info_DZ
Call nacteni_dat_Vybrane_Info_DZ
Sonik_Rucni_Pridani_Info_DZ.Show

End If

Sonik_Rucni_Pridani_Info_DZ.TextBox_Add_CL.value = ""



End Sub

Private Sub CommandButton_Add_Provided_Info_Click()

'Dodìlat tu to, že chci pøidat všechny vybrané
'Nahrát to v Loopu jedno po druhém do kolekce
'
Dim x As Long


   For x = 1 To Sonik_Rucni_Pridani_Info_DZ.ListView_Provided_Info.ListItems.count
      If Sonik_Rucni_Pridani_Info_DZ.ListView_Provided_Info.ListItems(x).Selected = True Then
         'I can do What I want in here
         
         coll_provided_info.Add Sonik_Rucni_Pridani_Info_DZ.ListView_Provided_Info.ListItems(x).text
         
         Set itmX = Sonik_Rucni_Pridani_Info_DZ.ListBox_Info_Selected.ListItems.Add()
         itmX.text = Sonik_Rucni_Pridani_Info_DZ.ListView_Provided_Info.ListItems(x).text
         
      End If
   Next x


'Vytvoøil jsem si kolekci, kterou potom pøi uložení nahraju
'coll_provided_info.Add Me.ComboBox_Provided_Info.value
'
'
'Set itmX = Sonik_Rucni_Pridani_Info_DZ.ListBox_Info_Selected.ListItems.Add()
'      itmX.text = Me.ComboBox_Provided_Info.value




'Tady jsem chtel udelat pokus, aby se to vybrane uz z Comboboxu vymazalo, ale nejak to nefunguje
'CPI_Count = coll_provided_info.count
'
'    For intComboItem = 0 To Me.ComboBox_Provided_Info.ListCount - 1
'        'If Me.ComboBox_Provided_Info(intComboItem).value = coll_provided_info(CPI_Count) Then
'        If Me.ComboBox_Provided_Info(intComboItem)
'        Me.ComboBox_Provided_Info.RemoveItem (intComboItem)
'        End If
'    Next intComboItem

End Sub


Private Sub CommandButton_Add_Req_Info_Click()


Dim x As Long
'RISK-869
Dim Req_Info As Variant
Dim sql_insert_req_info As String


   For x = 1 To Sonik_Rucni_Pridani_Info_DZ.ListView_Req_Info.ListItems.count
      If Sonik_Rucni_Pridani_Info_DZ.ListView_Req_Info.ListItems(x).Selected = True Then
         'I can do What I want in here
         
         coll_req_info.Add Sonik_Rucni_Pridani_Info_DZ.ListView_Req_Info.ListItems(x).text
         
         Set itmX = Sonik_Rucni_Pridani_Info_DZ.ListBox_Info_Selected.ListItems.Add()
         itmX.text = Sonik_Rucni_Pridani_Info_DZ.ListView_Req_Info.ListItems(x).text
         
      End If
   Next x

'Vytvoøil jsem si kolekci, kterou potom pøi uložení nahraju
'coll_req_info.Add Me.ComboBox_Req_Info.value
'
'Set itmX = Sonik_Rucni_Pridani_Info_DZ.ListBox_Info_Selected.ListItems.Add()
'      itmX.text = Me.ComboBox_Req_Info.value

'RISK-869
'Pokud bude v kolekci pozadovaneho infa jedna z "Platební karta" nebo "Kamerové záznamy" potom odkryju pole pro zadání, od kdy je informace požadována.
'Na základì toho se potom na druhý den v proceduøe vygeneruje ke klientùm upozornìní - musí se to projet loopem až po naètení dat pro klienta a vyhodit do upozornìní štrùdl klientù, kterých se to týká


'Required Info


For Each Req_Info In coll_req_info

    'RISK-945
    If CStr(Req_Info) = "Platební karta" Then ' Or CStr(Req_Info) = "Kamerové záznamy" Then
        Bln_Show_Date_Od = True
        Exit For
    End If


Next Req_Info


If Bln_Show_Date_Od Then
    Label_Datum_Od.Visible = True
    TextBox_Datum_Od.Visible = True
    TextBox_Datum_Od.ControlTipText = "Zadej datum požadovaných kamerových záznamù ve formátu DD.MM.YYYY"
End If


End Sub


'Potvrzení a zavøení

Private Sub CommandButton_Uloz_Click()


'RISK-850

Dim Bln_update_system_notice As Boolean
Bln_update_system_notice = False

'Provided Info
Dim Prov_Info As Variant
Dim sql_insert_prov_info As String

'RISK-869
Dim Date_Datum_Od As Date

'RISK-869
If Bln_Show_Date_Od Then

    If TextBox_Datum_Od.value = vbNullString Then 'RISK-945
        MsgBox "Vybral/a jsi v požadovaných informacích Platební kartu. Vyplò tedy prosím kolonku Datum od.", vbExclamation, "Chybné zadání hodnot"
        Exit Sub
    ElseIf Not IsDate(TextBox_Datum_Od.value) Then
        MsgBox "Do kolonky Datum od jsi nezadal správný formát datumu. Zkus to znovu", vbExclamation, "Chybné zadání hodnot"
        Exit Sub
    Else: Date_Datum_Od = TextBox_Datum_Od.value
    End If
End If


For Each Prov_Info In coll_provided_info

    'Uložení provided info
    sql_insert_prov_info = "insert into  " & schema & ".sonik_provided_info" & _
    " select " & schema & ".sonik_provided_info_seq.nextval," & SKP_datove_zpravy & ", a.skp_req_prov_info, sysdate, '" & UCase(Main.username) & "'" & _
    " from " & schema & ".sonik_cl_req_prov_info a" & _
    " left join " & schema & ".sonik_provided_info b on a.skp_req_prov_info = b.skp_provided_info and b.skp_data_message = " & SKP_datove_zpravy & _
    " where a.name_req_prov_info = '" & Prov_Info & "'" & _
    " and b.skp_provided_info is null /* Ještì tam tato hodnota není */"
    
    Call execute(sql_insert_prov_info)
    
    


Next Prov_Info







'Required Info
Dim Req_Info As Variant
Dim sql_insert_req_info As String

For Each Req_Info In coll_req_info

    'RISK-869
    'Uložení required info to_date('" & Datum_doruceni_datove_zpravy & "','DD.MM.YYYY')
    If (Bln_Show_Date_Od And CStr(Req_Info) = "Platební karta") Or CStr(Req_Info) = "Kamerové záznamy" Then
    
        'RISK-945
        If CStr(Req_Info) = "Kamerové záznamy" Then Date_Datum_Od = date - 60 'Pro kamerové záznamy novì datum zadávat nebudou, budu sledovat celou historii - Z poboèky jsou kamerové záznamy uchovávány 30 dní a z bankomatù 60 dní.
        
        sql_insert_req_info = "insert into  " & schema & ".sonik_req_info" & _
        " select " & schema & ".sonik_req_info_seq.nextval," & SKP_datove_zpravy & ", a.skp_req_prov_info, sysdate, '" & UCase(Main.username) & "',to_date('" & Date_Datum_Od & "','DD.MM.YYYY')" & _
        " from " & schema & ".sonik_cl_req_prov_info a" & _
        " left join " & schema & ".sonik_req_info b on a.skp_req_prov_info = b.skp_req_info and b.skp_data_message = " & SKP_datove_zpravy & _
        " where a.name_req_prov_info = '" & Req_Info & "'" & _
        " and b.skp_req_info is null /* Ještì tam tato hodnota není */"
    Else:
        sql_insert_req_info = "insert into  " & schema & ".sonik_req_info" & _
        " select " & schema & ".sonik_req_info_seq.nextval," & SKP_datove_zpravy & ", a.skp_req_prov_info, sysdate, '" & UCase(Main.username) & "',null" & _
        " from " & schema & ".sonik_cl_req_prov_info a" & _
        " left join " & schema & ".sonik_req_info b on a.skp_req_prov_info = b.skp_req_info and b.skp_data_message = " & SKP_datove_zpravy & _
        " where a.name_req_prov_info = '" & Req_Info & "'" & _
        " and b.skp_req_info is null /* Ještì tam tato hodnota není */"
    End If
    Call execute(sql_insert_req_info)
    
    
    
    'RISK-850
    If CStr(Req_Info) = "Platební karta" Or CStr(Req_Info) = "Vyèíslení škody" Then
        Bln_update_system_notice = True
    End If


Next Req_Info


'Kolekci promažu
Set coll_provided_info = Nothing
Set coll_req_info = Nothing





'Promazu pomocnou tabulku s vybranymi hodnotami
Sonik_Rucni_Pridani_Info_DZ.ListBox_Info_Selected.ListItems.Clear


'RISK-850
If Bln_update_system_notice Then
'Loop pøes aktivní upozornìní, najdu to, kde je sys a a textace Termín zpracování souèinnosti
'Vezmu si dané ID a udìlám update v tabulce sonik_data_message_notice na DATE_DATA_MESSAGE_DELIVERED + 14, pokud tam ale již není míò - sloupec zbývá
    
    Dim i As Integer
    Dim Int_ID_Upozorneni As Long
    Dim Str_sql_update_notice As String
    Dim Time_Offset As Integer
    
    Time_Offset = 14
    
    For i = 1 To Sonik_Datova_Zprava.ListView_DZ_Upozorneni.ListItems.count
      If Sonik_Datova_Zprava.ListView_DZ_Upozorneni.ListItems(i).SubItems(1) = "Termín zpracování souèinnosti" And Sonik_Datova_Zprava.ListView_DZ_Upozorneni.ListItems(i).SubItems(3) > 14 Then
        Int_ID_Upozorneni = Sonik_Datova_Zprava.ListView_DZ_Upozorneni.ListItems(i).text
      End If
    Next i

    'Teï ten update
    Str_sql_update_notice = "update " & schema & ".sonik_data_message_notice" & _
                            " set DTIME_NOTICE_DEADLINE = to_date('" & Datum_doruceni_datove_zpravy & "','DD.MM.YYYY')+ " & Time_Offset & _
                            "    ,DTIME_MODIFIED = sysdate" & _
                            "    ,NAME_EMPLOYEE_MODIFIED = 'sys'" & _
                            " where SKP_DATA_MESSAGE_NOTICE = " & Int_ID_Upozorneni
    Call execute(Str_sql_update_notice)

End If

'RISK-869
TextBox_Datum_Od.value = ""

'A obnovim vypis dat
Call nacteni_dat_Sonik_Datova_Zprava

Me.Hide






End Sub




' Zavøení formuláøe
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Call set_filter(Me.name, 1)

'Promazu pomocnou tabulku s vybranymi hodnotami
Sonik_Rucni_Pridani_Info_DZ.ListBox_Info_Selected.ListItems.Clear

'RISK-869
TextBox_Datum_Od.value = ""

'Kolekci promažu
Set coll_provided_info = Nothing
Set coll_req_info = Nothing

Me.Hide
End Sub


' Back button
Private Sub CommandButton_Navrat_Click()

'Promazu pomocnou tabulku s vybranymi hodnotami
Sonik_Rucni_Pridani_Info_DZ.ListBox_Info_Selected.ListItems.Clear

'RISK-869
TextBox_Datum_Od.value = ""

'Kolekci promažu
Set coll_provided_info = Nothing
Set coll_req_info = Nothing

Call text_information_hide(formular)
Call UserForm_QueryClose(1, 0)
End Sub




Private Sub UserForm_Activate()
'RISK-869
'Iniciální nastavení
Bln_Show_Date_Od = False
Label_Datum_Od.Visible = False
TextBox_Datum_Od.Visible = False

End Sub












' Zmìna barvy po najetí myši **************************************************************************************************************************

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Uloz.ForeColor = barvaCerna
End Sub

Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Navrat.ForeColor = barvaCerna
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Uloz.ForeColor = barvaCerna
End Sub

Private Sub CommandButton_Navrat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Navrat.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Uloz_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Uloz.ForeColor = barvaZelenaAirBank
End Sub


