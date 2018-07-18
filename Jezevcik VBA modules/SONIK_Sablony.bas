Attribute VB_Name = "SONIK_Sablony"
Option Explicit


'RISK-1235 - prepis na private, tam kde to kompilatoru stacilo
Private sonik_name1 As String
Private sonik_name2 As String
Private sonik_dateBirth As Date
Private sonik_DPD As Long
Private sonik_FFPD As String
Private sonik_trvala_adresa As String 'RISK-992


Public Sonik_AO_Adresa As String
Public Sonik_AO_IDds As String
Public Sonik_AO_Subjekt As String
Public Sonik_AO_SpisZn As String 'Spisova znacka
Public Sonik_AO_SpisZn2 As String 'Toto je realne cislo jednaci
Public Sonik_AO_ZadostOd As String
Public Sonik_AO_ZadostDen As Date
Public Sonik_AO_KontrolaCastek As String
Public Sonik_AO_OdPohybFin As String
Public Sonik_AO_Tel_Cislo_Adresata As String
Public Sonik_AO_rst_address As Recordset

Private Sonik_AO_PM As Boolean
Private Sonik_AO_Ins_dne As Date
Private Sonik_AO_ins_SZ As String
Private Sonik_AO_zalobaPodana As Boolean
Private Sonik_AO_zespl As Date
Private Sonik_AO_Zamitnuta As Long



Private Sonik_pr As Integer 'posledni radek




Public Sub Sonik_Export_sablona_NBU()

Dim sqlLoans As String, sqlMortgages As String 'RISK-1106
Dim sql_extra_instalments As String
Dim sqlRejectedApps As String 'RISK-1106
Dim MainRange As Range
Dim sonik_dateBirth As Date
Dim i As Integer
Dim j As Integer
Dim podpis As String
Dim cisloPozadavku As String
Dim datumPozadavku As Date

'RISK-868
Dim JezevecBook As Workbook

Set JezevecBook = ThisWorkbook
JezevecBook.Worksheets("Sonik_NBU").Activate

Worksheets("Sonik_NBU").Visible = False

Application.ScreenUpdating = False

Worksheets("Sonik_NBU").Visible = True



sonik_DPD = 0
sonik_FFPD = "N"


sonik_name1 = Sonik_Karta_Klienta.ListView_Client_Actual.ListItems(1).text
sonik_name2 = Sonik_Karta_Klienta.ListView_Client_Actual.ListItems(1).SubItems(1)
sonik_dateBirth = SonikFormatDate(Sonik_Karta_Klienta.ListView_Client_Actual.ListItems(1).SubItems(5)) 'RISK-1131


'cisloPozadavku = Application.InputBox("Zadej èíslo jednací", "Vstupní hodnoty pro vygenerování dokumentu", "-NBÚ/P")
'datumPozadavku = Application.InputBox("Zadej datum požadavku", "Vstupní hodnoty pro vygenerování dokumentu", Format$(Now(), "d. m. yyyy"))

datumPozadavku = SonikFormatDate(Sonik_Karta_Klienta.ListView_Soucinnosti_Klienta.SelectedItem.SubItems(2)) 'RISK-1131
cisloPozadavku = Sonik_Karta_Klienta.ListView_Soucinnosti_Klienta.SelectedItem.SubItems(5)


Worksheets("Sonik_NBU").UsedRange.ClearContents
Worksheets("Sonik_NBU").UsedRange.ClearFormats

 
'RISK-1106 obohacení o ukonèené/doplacené úvìry
sqlLoans = "select adu.cislo_uveru " & _
     ", adu.vyse_uveru " & _
     ", adu.datum_nacerpani " & _
     ", adu.cislo_uctu " & _
     ", adu.amt_instalment " & _
     ", adu.nesplacena_jistina " & _
     ", adu.cnt_days_past_due_tol_k_max " & _
     ", adu.flag_1st_payment_def " & _
     ", adu.cnt_bonus_level " & _
     ", adu.datum_podpisu_pujcky " & _
     ", adu.name_loan_wfstatus " & _
     " from " & schema & ".sonik_apl_data_uver adu " & _
     " where adu.cislo_uveru is not null and adu.datum_nacerpani < sysdate " & _
     " and adu.skp_person = " & skp_person_KL


'RISK-1106 Hypotéky
sqlMortgages = "select adh.cislo_uveru " & _
     ", adh.vyse_pujcky " & _
     ", adh.datum_nacerpani " & _
     ", adh.cislo_uctu " & _
     ", adh.amt_instalment " & _
     ", adh.nesplacena_jistina " & _
     " from " & schema & ".sonik_apl_data_hypoteka adh " & _
     " where adh.cislo_uveru is not null and adh.datum_nacerpani < sysdate " & _
     " and adh.skp_person = " & skp_person_KL


'Mimoøádné splátky
sql_extra_instalments = "select ms.code_loan_number " & _
     ", ms.amt_transaction " & _
     ", ms.dtime_transaction_generation " & _
     " from " & schema & ".sonik_mimorad_splatky ms " & _
     " where ms.skp_person = " & skp_person_KL

'Poèet zamítnutých žádostí klienta o úvìrové produkty
sqlRejectedApps = "SELECT count(*) as CNT_REJECTED_APPS " & _
    " from " & schema & ".sonik_v_rejected_apps" & _
    " WHERE skp_person = " & skp_person_KL & _
    " group by skp_person "

Worksheets("Sonik_NBU").Cells(1, 1).value = "Informace jsou aktuální ke dni " & SonikFormatDate(Now() - 1) 'RISK-1131
Worksheets("Sonik_NBU").Cells(2, 1).value = sonik_name1 & " " & sonik_name2 & ", nar. " & SonikFormatDate(sonik_dateBirth) 'RISK-1131
Worksheets("Sonik_NBU").Cells(2, 1).Font.Bold = True


Call Sonik_NBU_ucty("Bìžný úèet")
Call Sonik_NBU_ucty("Spoøicí úèet")
Call Sonik_NBU_uvery(sqlLoans) 'RISK-1106
Call Sonik_NBU_rejectedApps(sqlRejectedApps) 'RISK-1106
Call Sonik_NBU_mimSplatky(sql_extra_instalments)
'RISK-1106
Call Sonik_NBU_Hypo(sqlMortgages)


Sheets("Sonik_NBU").Select
If Sheets("Sonik_NBU").Range("A" & ActiveSheet.Rows.count).End(xlUp).Row = 2 Then
Sheets("Sonik_NBU").Range("A3") = "Zjištìná osoba v souèasné dobì nemá u Air Bank zøízen žádný úèet ani žádné jiné bankovní produkty."
Else
Range("A" & Sheets("Sonik_NBU").Range("A" & ActiveSheet.Rows.count).End(xlUp).Row + 1) = "Zjištìná osoba nemá u Air Bank zøízen žádný další úèet a nevyužívá žádné jiné bankovní produkty."

'RISK-1106 doplnìna teèka na konec vìty
Range("A" & Sheets("Sonik_NBU").Range("A" & ActiveSheet.Rows.count).End(xlUp).Row + 1) = "Air Bank nevede pro klienta depozitní ani jiné schránky a nespravuje ani neuchovává klientovi majetek, vìci, listiny ani zaknihované cenné papíry."

'Range("A" & Sheets("Sonik_NBU").Range("A" & ActiveSheet.Rows.count).End(xlUp).Row + 2) = "Air Bank nevede pro klienta depozitní ani jiné schránky a nespravuje ani neuchovává klientovi "
'Range("A" & Sheets("Sonik_NBU").Range("A" & ActiveSheet.Rows.count).End(xlUp).Row + 1) = "majetek, vìci, listiny ani zaknihované cenné papíry."

End If


Worksheets("Sonik_NBU").Range("A" & Worksheets("Sonik_NBU").Rows.count).End(xlUp).Select
Range(Selection, Cells(1, 6)).Select


With Selection
    .Font.name = "Arial"
    .Font.Size = 10
    '.HorizontalAlignment = xlLeft
    .Orientation = 0
    .AddIndent = False
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With
Set MainRange = Range(Selection, Cells(1, 6))


'RISK-1217
'RISK-641 - úprava podpisu podle toho, kdo generuje šablonu
'Šablona Lucka:
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

podpis = SonikSignature

'RISK-854
'Nastaveni sirky sloupcu
Columns("A").ColumnWidth = 11.71
Columns("B").ColumnWidth = 13.86
Columns("C").ColumnWidth = 14.14
Columns("D").ColumnWidth = 16#
Columns("E:R").ColumnWidth = 8.43


Call Sonik_ExportDoWordu(cfg.Cesta("vzorSonikNBU"), "NBU", cisloPozadavku, cisloPozadavku, datumPozadavku, "HlavnioddilSablonaNBU", MainRange, podpis)



Worksheets("Sonik_NBU").UsedRange.ClearContents
Worksheets("Sonik_NBU").UsedRange.ClearFormats
Worksheets("Sonik_NBU").Visible = False


Application.ScreenUpdating = True


End Sub







Public Sub Sonik_AO_Export_automaticka_odpoved() 'zobrazení formuláøe pro zadání vtupních informací

Dim sql_address As String
Dim i As Integer
'Nahraju data pro výbìr Adresy, kam mají poslat


sql_address = "select SKP_ANSWER_RECIPIENT, RECIPIENT_NAME, RECIPIENT_ADDRESS, TEXT_PHONE_NUMBER, ID_DATA_BOX from " & schema & ".sonik_cl_answer_recipient where flag_valid = 'Y'"


'naplnìní zákl. udajù o osobì
 Set Sonik_AO_rst_address = New ADODB.Recordset
   Sonik_AO_rst_address.Open sql_address, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
  
   
   If Not Sonik_AO_rst_address.EOF Then
    Sonik_AO_rst_address.MoveFirst
    For i = 1 To Sonik_AO_rst_address.RecordCount
    Sonik_VyzadaniSoucinnosti.ComboSubjekt.AddItem "" & Sonik_AO_rst_address![RECIPIENT_NAME] & ""
    Sonik_AO_rst_address.MoveNext
    Next i
    
   End If
 ' zavøít, až dá tlaèítko ok, odeslat...
 'Sonik_AO_rst_address.Close
 
  If Sonik_VyzadaniSoucinnosti.ComboSubjekt.ListCount > 0 Then
    Sonik_VyzadaniSoucinnosti.ComboSubjekt.ListIndex = 0
 End If



   Sonik_VyzadaniSoucinnosti.TextSubjAdresa.value = ""
   Sonik_VyzadaniSoucinnosti.TextSubjAdresa.ControlTipText = "Jako oddìlovaè použij støedník, bude nahrazen koncem øádku."
   Sonik_VyzadaniSoucinnosti.TextIDds.value = ""
   'Sonik_VyzadaniSoucinnosti.TextZadostDen.value = Format$(Now(), "d. m. yyyy")
   'Sonik_VyzadaniSoucinnosti.TextZadostOd.text = "<konkrétního státního zastupitelství>"
   'Sonik_VyzadaniSoucinnosti.TextSpisZn.text = "sp. zn. / è. j."
   Sonik_VyzadaniSoucinnosti.TextKontrolaCastek.value = ""
   Sonik_VyzadaniSoucinnosti.TextKontrolaCastek.ControlTipText = "Pokud pole zùstane nevyplnìné, tak se tato èást v šablonì neplní. Èíslo musí být ve formátu bez mezer."
   Sonik_VyzadaniSoucinnosti.TextOdPohyb.value = ""
   'Sonik_VyzadaniSoucinnosti.TextSpisZn02.value = ""
   'Sonik_VyzadaniSoucinnosti.TextSpisZn.ControlTipText = "Vyplò èíslo jednací doplòeného za nadpis Vyžádání souèinnosti."
   'Sonik_VyzadaniSoucinnosti.TextSpisZn02.ControlTipText = "Vyplò èíslo jednací doplòeného za oznaèení státního zastupitelství."
   Sonik_VyzadaniSoucinnosti.Show

End Sub
'RISK-415, RISK-526, RISK-641
Public Sub Sonik_AO_GenerujAutOdpoved()
Dim sql0 As String
Dim sql1 As String
Dim sql2 As String
Dim sql3 As String
Dim sql4 As String
Dim sql5 As String
Dim sql6 As String
Dim sql7 As String
Dim sql8 As String
Dim sql9 As String
Dim sql10 As String
Dim sql11 As String
Dim sqlMaDispo As String
Dim sqlJeDispo As String
Dim MainRange As Range
Dim i As Integer

Dim podpis As String
Dim cisloPozadavku As String
Dim datumPozadavku As Date

'RISK-868
Dim JezevecBook As Workbook

Set JezevecBook = ThisWorkbook
JezevecBook.Worksheets("Sonik_NBU").Activate


sonik_name1 = ""
sonik_name2 = ""
sonik_trvala_adresa = "" 'RISK-992


Sonik_AO_PM = False
Sonik_AO_zalobaPodana = False
Sonik_AO_Ins_dne = DateSerial(2000, 1, 1)
sonik_DPD = 0
sonik_FFPD = "N"
Sonik_AO_Zamitnuta = 0


'PERSON
sql0 = "select abk.skp_person, abk.name_first, abk.name_last, abk.date_birth, abk.name_permanent_address_street, abk.text_permanent_address_num, abk.name_permanent_address_town, abk.code_permanent_address_zip from " & schema & ".sonik_ab_klient_dtb abk where abk.SKP_Person = " & skp_person_KL


Worksheets("Sonik_NBU").UsedRange.ClearContents
Worksheets("Sonik_NBU").UsedRange.ClearFormats


'naplnìní zákl. udajù o osobì
 Set rst1 = New ADODB.Recordset
   rst1.Open sql0, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
        sonik_name1 = rst1![Name_first]
        sonik_name2 = rst1![Name_last]
        sonik_dateBirth = SonikFormatDate(rst1![Date_Birth]) 'RISK-1131
        sonik_trvala_adresa = "" & rst1![name_permanent_address_street] & " " & rst1![text_permanent_address_num] & ", " & format(rst1![code_permanent_address_zip], "### ##") & " " & rst1![name_permanent_address_town] 'RISK-992
    rst1.MoveNext
    Next i
    
   End If
 
rst1.Close



'BU 'RISK-992 odmazání account_relationship
sql1 = "select u.text_bank_account_number || '/3030' as CisloUctu, u.date_bank_account_activation, u.code_currency, u.ucetni_zustatek, u.name_bank_account_status, u.dtime_of_application, u.APP_CHANNEL_TYPE_REQUEST" & _
       " from " & schema & ".sonik_ucty u" & _
       " where u.bank_account_type = 'Bìžný úèet'" & _
       " and u.SKP_Person = " & skp_person_KL




'SU 'RISK-992 odmazání account_relationship
sql2 = "select u.text_bank_account_number || '/3030' as CisloUctu, u.date_bank_account_activation, u.code_currency, u.ucetni_zustatek, u.name_bank_account_status, u.dtime_of_application, u.APP_CHANNEL_TYPE_REQUEST" & _
       " from " & schema & ".sonik_ucty u" & _
       " where u.bank_account_type = 'Spoøicí úèet'" & _
       " and u.SKP_Person = " & skp_person_KL



 
'Aktivní úvìry !!!!! je NOK
sql3 = "select adu.cislo_uveru, adu.vyse_uveru, adu.datum_nacerpani, adu.cislo_uctu, adu.amt_instalment, adu.nesplacena_jistina, adu.cnt_days_past_due_tol_k_max, adu.flag_1st_payment_def, adu.cnt_bonus_level, adu.datum_podpisu_pujcky, adu.name_loan_wfstatus" & _
       " from " & schema & ".sonik_apl_data_uver adu" & _
       " where adu.name_loan_status = 'Aktivní' and adu.skp_person = " & skp_person_KL



'Žádosti
sql4 = "select adu.cislo_uveru, adu.cislo_zadosti, adu.vyse_uveru, adu.cas2scoring, adu.VYSLEDEK_ZADOSTI, adu.duvod_zamitnuti_celkovo " & _
       " from " & schema & ".sonik_apl_data_uver adu " & _
       " where adu.cas2scoring is not null and adu.skp_person = " & skp_person_KL
 



'Doklady - chybí propojení v tabulkách, proto nelze rozumì dotáhnout platnost do
sql5 = "select kd.name_document_type, kd.date_document_issue from " & schema & ".sonik_klient_doklady kd where kd.skp_person = " & skp_person_KL



'Insolvence - podaná pøihláška
       sql6 = "select p.skp_person, i.code, i.datum_prihlaseni from " & schema & ".ins_info i" & _
             " join " & schema & ".person p on i.id_person = p.id_person and p.flag_valid = 'Y' " & _
      " where i.datum_prihlaseni Is Not Null And p.skp_person = " & skp_person_KL
Set rst1 = New ADODB.Recordset
   rst1.Open sql6, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   If Not rst1.EOF Then
      Sonik_AO_Ins_dne = SonikFormatDate(rst1![datum_prihlaseni]) 'RISK-1131
      Sonik_AO_ins_SZ = rst1![code]
    End If
rst1.Close


'zjistime, zda ma klient pravomocne rozhodnuti - ovlnivni to par veci v sabolne
      sql7 = "select 1, nvl(pm.pm,0) as pm, p.date_payoff from " & schema & ".epr_status_data e" & _
      " join  " & schema & ".epr_status es on e.skp_epr_status = es.skp_epr_status join " & schema & ".PERSON p on p.ID_PERSON = e.id_person " & _
      " left join (select id_person, 1 as pm from " & schema & ".epr_status_data e join " & schema & ".epr_status es on e.skp_epr_status = es.skp_epr_status where name_epr_status like 'PM%') pm on pm.id_person = e.id_person " & _
      " where es.name_epr_status like 'Žaloba podána%' and p.skp_person = " & skp_person_KL
Set rst1 = New ADODB.Recordset
   rst1.Open sql7, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   If Not rst1.EOF Then
      Sonik_AO_zalobaPodana = True
      Sonik_AO_PM = rst1![pm]
      Sonik_AO_zespl = SonikFormatDate(rst1![Date_payoff]) 'RISK-1131
    End If
rst1.Close


'RISK-992 odmazání account_relationship - nutno doøešit v rámci RISK-978
'Zatím dávám distinct, protože tam má vlastík v tabulce nìjaké zdvojení
'Klient je majitel a sleduji jeho disponenty
sqlMaDispo = "SELECT DISTINCT d.TEXT_BANK_ACCOUNT_NUMBER, d.DTIME_RIGHT_VALID_FROM, d.DTIME_RIGHT_VALID_TO, k.name_first, k.name_last, k.date_birth " & _
           " , k.name_permanent_address_street, k.text_permanent_address_num, k.code_permanent_address_zip ,k.name_permanent_address_town " & _
           " from " & schema & ".sonik_disponents d " & _
           " JOIN " & schema & ".sonik_ab_klient_dtb k ON d.skp_person_disponent = k.skp_person " & _
           "                                                   AND d.code_right_type = 'Úèet' " & _
           " WHERE d.skp_person_owner = " & skp_person_KL & _
           " order by d.TEXT_BANK_ACCOUNT_NUMBER, d.DTIME_RIGHT_VALID_FROM "
            
'RISK-992 odmazání account_relationship - nutno doøešit v rámci RISK-978
'Zatím dávám distinct, protože tam má vlastík v tabulce nìjaké zdvojení
'Klient je disponentem u jiných klientù
sqlJeDispo = "SELECT DISTINCT d.TEXT_BANK_ACCOUNT_NUMBER, d.DTIME_RIGHT_VALID_FROM, d.DTIME_RIGHT_VALID_TO " & _
           " from " & schema & ".sonik_disponents d " & _
           " WHERE d.code_right_type = 'Úèet' AND d.skp_person_disponent = " & skp_person_KL & _
           " order by d.TEXT_BANK_ACCOUNT_NUMBER, d.DTIME_RIGHT_VALID_FROM "


''''''''''''''''''''''''''''''''''''''''''''''''''''''

'kontrola èástek vložených pøes bankomat
sql10 = "select vatm.date_transaction_due, vatm.text_bank_account_number, vatm.amt_transaction, vatm.code_currency" & _
        " from " & schema & ".sonik_vklady_atm vatm" & _
        " where vatm.amt_transaction >= " & Sonik_AO_KontrolaCastek & _
        " and vatm.skp_person = " & skp_person_KL & "" & _
        " order by vatm.date_transaction_due"



'doplnìní kontaktních a personálních údajù RISK-526
sql11 = "select ab.name_first, ab.name_last, ab.date_birth, ab.primary_phone, ab.primary_email, ab.name_reports_address_street, ab.text_reports_address_num, ab.name_reports_address_town, ab.code_reports_address_zip, ab.name_permanent_address_street, ab.text_permanent_address_num, ab.name_permanent_address_town, ab.code_permanent_address_zip" & _
        " from " & schema & ".sonik_ab_klient_dtb ab where ab.skp_person = " & skp_person_KL
  
 




'zde se zaèínají plnit informace do šablony
Worksheets("Sonik_NBU").Cells(1, 1).value = "Níže uvedené informace jsou aktuální ke dni " & SonikFormatDate(Now() - 1) & "            "  'RISK-1131
Worksheets("Sonik_NBU").Cells(2, 1).value = sonik_name1 & " " & sonik_name2 & ", nar. " & SonikFormatDate(sonik_dateBirth) & "            " 'RISK-1131
Worksheets("Sonik_NBU").Cells(2, 1).Font.Bold = True


Call Sonik_AO_ucty(sql1, "Bìžné úèty")
Call Sonik_AO_ucty(sql2, "Spoøicí úèty")
Call Sonik_AO_MaDispo(sqlMaDispo) 'RISK-992
Call Sonik_AO_JeDispo(sqlJeDispo) 'RISK-992
Call Sonik_AO_uvery(sql3)
Call Sonik_AO_Zadosti(sql4)
Call Sonik_AO_Doklady(sql5)
Call Sonik_AO_Kontakty(sql11)
Range("A" & ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row + 2) = "Zjištìná osoba nemá u Air Bank zøízen žádný další úèet a nevyužívá žádné jiné bankovní produkty."



If Sonik_AO_KontrolaCastek <> "" Then
    Call Sonik_AO_Transakce(sql10)
End If



Application.ScreenUpdating = False

'Když jsem jej nezviditelnil, tak to nefungovalo
Worksheets("Sonik_NBU").Visible = True

Worksheets("Sonik_NBU").Activate
If ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row = 2 Then
    Range("A3") = "Zjištìná osoba v souèasné dobì nemá u Air Bank zøízen žádný úèet ani žádné jiné bankovní produkty."
Else
    Range("A" & ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row + 1) = "Uvedené osobì byla k úètu zøízena služba internetového bankovnictví, v pøíloze zasíláme logovací soubory a IP adresy."
    
    Range("A" & ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row + 1) = "Air Bank nevede pro klienta depozitní ani jiné schránky a nespravuje ani neuchovává klientovi majetek, vìci, listiny ani zaknihované cenné papíry."

End If


ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Select
Range(Selection, Cells(1, 7)).Select

With Selection
    .Font.name = "Arial"
    .Font.Size = 10
    '.HorizontalAlignment = xlLeft
    .Orientation = 0
    .AddIndent = False
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With
Set MainRange = Range(Selection, Cells(1, 7))

'RISK-1217
'RISK-641 - úprava podpisu podle toho, kdo generuje šablonu
'Šablona Lucka:
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
podpis = SonikSignature
 
'RISK-854
'Nastaveni sirky sloupcu
Columns("A").ColumnWidth = 11.71
Columns("B:R").ColumnWidth = 8.43
 
 
Call Sonik_ExportDoWordu(cfg.Cesta("vzorSonikAutoOdpoved"), Sonik_AO_Adresa, Sonik_AO_SpisZn, "sp. zn./è. j. " & Sonik_AO_SpisZn2, Sonik_AO_ZadostDen, "Hlavnioddil", MainRange, podpis)

Application.ScreenUpdating = True

Worksheets("Sonik_NBU").Visible = False


End Sub














Private Sub Sonik_NBU_ucty(typ As String)
'Pøepsáno oproti souèinnostníku, který se díval do DTB, tady se podívám pøímo do karty klienta, abych nemusel dìlat znova SQL select
  
Dim i As Integer
Dim j As Integer
Dim pocet_zaznamu_ucet As Integer
  
If Sonik_Karta_Klienta.ListView_Accounts.ListItems.count > 0 Then
 
 'Prvnì proloopuju, jestli má klient vùbec nìjaký úèet daného typu, pak totiž nechci vyplòovat ani tu hlavièku
    'RISK-992 - po odmazání nìkterých sloupeèkù je tøeba pøemapovat...
    For i = 1 To Sonik_Karta_Klienta.ListView_Accounts.ListItems.count
        If Sonik_Karta_Klienta.ListView_Accounts.ListItems(i).text = typ And SonikFormatDate(Sonik_Karta_Klienta.ListView_Accounts.ListItems(i).SubItems(4)) <> "1.1.3000" Then 'RISK-1131
            pocet_zaznamu_ucet = pocet_zaznamu_ucet + 1
        End If
    Next i
    
    If pocet_zaznamu_ucet = 0 Then
        Exit Sub
    End If


    Worksheets("Sonik_NBU").Activate
    Sonik_pr = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
    
    If typ = "Bìžný úèet" Then
        Cells(Sonik_pr + 1, 1).value = "Bìžné úèty"
    Else
        Cells(Sonik_pr + 1, 1).value = "Spoøicí úèty"
    End If
    
    Cells(Sonik_pr + 2, 1).value = "Èíslo"
    Cells(Sonik_pr + 2, 2).value = "Datum založení"
    Cells(Sonik_pr + 2, 3).value = "Majitel úètu"
    Cells(Sonik_pr + 2, 4).value = "Mìna"
    Cells(Sonik_pr + 2, 5).value = "Zùstatek"
    Cells(Sonik_pr + 2, 6).value = "Stav úètu"
    Range(Cells(Sonik_pr + 1, 1), Cells(Sonik_pr + 2, 6)).Font.Bold = True
    

    j = 1
    For i = 1 To Sonik_Karta_Klienta.ListView_Accounts.ListItems.count
     
     
        If Sonik_Karta_Klienta.ListView_Accounts.ListItems(i).text = typ And SonikFormatDate(Sonik_Karta_Klienta.ListView_Accounts.ListItems(i).SubItems(4)) <> "1.1.3000" Then 'RISK-1131
            'RISK-992 - po odmazání nìkterých sloupeèkù je tøeba pøemapovat...
            Cells(Sonik_pr + 2 + j, 1).value = Sonik_Karta_Klienta.ListView_Accounts.ListItems(i).SubItems(2) & "/3030"
            Cells(Sonik_pr + 2 + j, 2).value = SonikFormatDate(Sonik_Karta_Klienta.ListView_Accounts.ListItems(i).SubItems(4)) 'RISK-1131
            Cells(Sonik_pr + 2 + j, 3).value = sonik_name1 & " " & sonik_name2
            Cells(Sonik_pr + 2 + j, 4).value = Sonik_Karta_Klienta.ListView_Accounts.ListItems(i).SubItems(1)
            Cells(Sonik_pr + 2 + j, 5).value = format$(Sonik_Karta_Klienta.ListView_Accounts.ListItems(i).SubItems(12), "#,##0.00 " & Sonik_Karta_Klienta.ListView_Accounts.ListItems(i).SubItems(1))
            Cells(Sonik_pr + 2 + j, 6).value = Sonik_Karta_Klienta.ListView_Accounts.ListItems(1).SubItems(3)

        
            j = j + 1
        
        End If
        
    Next i

    Range(Cells(Sonik_pr + 2, 1), Cells(Sonik_pr + 1 + j, 6)).Select
    Call Sonik_borders
    'RISK-1106 - vložení vìty pod tabulku s úèty
    Cells(Sonik_pr + 2 + i, 1).value = "Pohyb finanèních prostøedkù na úètu/úètech za období od do souèasnosti zasíláme v pøíloze."
    Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekcí
End If

End Sub



Private Sub Sonik_NBU_uvery(sql As String)


Dim i As Integer
Dim uver1 As String
Dim uver2 As String

Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   If Not rst1.EOF Then
    rst1.MoveFirst
    Worksheets("Sonik_NBU").Activate
    Sonik_pr = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
    Cells(Sonik_pr + 1, 1).value = "Úvìry" 'RISK-1106
    Cells(Sonik_pr + 2, 1).value = "Èíslo"
    Cells(Sonik_pr + 2, 2).value = "Výše úvìru"
    Cells(Sonik_pr + 2, 3).value = "Datum naèerpání"
    Cells(Sonik_pr + 2, 4).value = "Èíslo úètu, kam byl úvìr naèerpán"
    Cells(Sonik_pr + 2, 5).value = "Výše pravidelné mìsíèní splátky"
    Cells(Sonik_pr + 2, 6).value = "Zùstatek jistiny"
    Range(Cells(Sonik_pr + 1, 1), Cells(Sonik_pr + 2, 6)).Font.Bold = True
    For i = 1 To rst1.RecordCount
        Cells(Sonik_pr + 2 + i, 1).value = rst1![CISLO_UVERU]
        Cells(Sonik_pr + 2 + i, 2).value = format$(rst1![vyse_uveru], "#,##0.00 CZK")
        Cells(Sonik_pr + 2 + i, 3).value = SonikFormatDate(rst1![DATUM_NACERPANI]) 'RISK-1131
        Cells(Sonik_pr + 2 + i, 4).value = rst1![cislo_uctu]
        Cells(Sonik_pr + 2 + i, 5).value = format$(rst1![Amt_Instalment], "#,##0.00 CZK")
        Cells(Sonik_pr + 2 + i, 6).value = format$(rst1![NESPLACENA_JISTINA], "#,##0.00 CZK")
        'promìnná pro DPD a FPD
        sonik_DPD = Application.max(rst1![cnt_days_past_due_tol_k_max], sonik_DPD)
        If rst1![flag_1st_payment_def] = "Y" Then sonik_FFPD = rst1![flag_1st_payment_def]
        'End If
        rst1.MoveNext
    Next i
    
    Range(Cells(Sonik_pr + 2, 1), Cells(Sonik_pr + 1 + i, 6)).Select
    Call Sonik_borders
    
    If i = 2 Then
        uver1 = "úvìru"
        uver2 = "úvìr"
    Else
        uver1 = "úvìrù"
        uver2 = "úvìry"
    End If
    
     'RISK-1106 doplnìno u pøed u úvìru
        If ((sonik_DPD > 5 And sonik_FFPD = "N") Or (sonik_DPD > 11 And sonik_FFPD = "Y")) Then
            Cells(Sonik_pr + 2 + i, 1).value = sonik_name1 & " " & sonik_name2 & " je u " & uver1 & " žadatelem a " & uver2 & " nesplácí øádnì a vèas dohodnutým zpùsobem."
        Else
            If sonik_DPD = 0 Then
                Cells(Sonik_pr + 2 + i, 1).value = sonik_name1 & " " & sonik_name2 & " je u " & uver1 & " žadatelem a " & uver2 & " splácí øádnì a vèas dohodnutým zpùsobem."
            Else
                Cells(Sonik_pr + 2 + i, 1).value = sonik_name1 & " " & sonik_name2 & " je u " & uver1 & " žadatelem a " & uver2 & " splácí øádnì, avšak nìkdy s nìkolikadenním zpoždìním."
            End If
        End If
   
   
        Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekci
   
   End If
rst1.Close



End Sub



Private Sub Sonik_NBU_rejectedApps(sql As String)
'RISK-1106

Dim IntCntRejectedApps As Integer

Sonik_pr = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row

Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   If Not rst1.EOF Then

        IntCntRejectedApps = rst1![CNT_REJECTED_APPS]

   End If

    'RISK-1106 doplnìno u pøed u úvìru
    If IntCntRejectedApps = 0 Then
        Cells(Sonik_pr + 1, 1).value = "U klienta neevidujeme žádné zamítnuté žádosti o úvìrové produkty."
    ElseIf IntCntRejectedApps = 1 Then
        Cells(Sonik_pr + 1, 1).value = "U klienta evidujeme další žádost o úvìr. Po posouzení všech údajù byla tato žádost ze strany banky zamítnuta a úvìr nebyl poskytnut. Všechny žádosti klientù o úvìry jsou vždy posuzovány zcela individuálnì a kritéria rozhodování se v prùbìhu èasu mìní"
    ElseIf IntCntRejectedApps < 5 Then
        Cells(Sonik_pr + 1, 1).value = "U klienta evidujeme další " & IntCntRejectedApps & " žádosti o úvìr. Po posouzení všech údajù byly tyto žádosti ze strany banky zamítnuty a úvìry nebyly poskytnuty. Všechny žádosti klientù o úvìry jsou vždy posuzovány zcela individuálnì a kritéria rozhodování se v prùbìhu èasu mìní"
    Else
        Cells(Sonik_pr + 1, 1).value = "U klienta evidujeme dalších " & IntCntRejectedApps & " žádostí o úvìr. Po posouzení všech údajù byly tyto žádosti ze strany banky zamítnuty a úvìry nebyly poskytnuty. Všechny žádosti klientù o úvìry jsou vždy posuzovány zcela individuálnì a kritéria rozhodování se v prùbìhu èasu mìní"
    End If

   Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekci

rst1.Close



End Sub


Private Sub Sonik_NBU_Hypo(sql As String)
'RISK-1106

Dim i As Integer

Set rst1 = New ADODB.Recordset
rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst1.EOF Then
        rst1.MoveFirst
        Worksheets("Sonik_NBU").Activate
        Sonik_pr = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
        Cells(Sonik_pr + 1, 1).value = "Hypoteèní úvìry"
        Cells(Sonik_pr + 2, 1).value = "Èíslo úvìru"
        Cells(Sonik_pr + 2, 2).value = "Úvìrový rámec"
        Cells(Sonik_pr + 2, 3).value = "Datum naèerpání"
        Cells(Sonik_pr + 2, 4).value = "Èíslo úètu, kam byl úvìr naèerpán"
        Cells(Sonik_pr + 2, 5).value = "Výše splátky"
        Cells(Sonik_pr + 2, 6).value = "Zùstatek jistiny"
        Range(Cells(Sonik_pr + 1, 1), Cells(Sonik_pr + 2, 6)).Font.Bold = True
        
        For i = 1 To rst1.RecordCount
            Cells(Sonik_pr + 2 + i, 1).value = rst1![CISLO_UVERU]
            Cells(Sonik_pr + 2 + i, 2).value = format$(rst1![vyse_pujcky], "#,##0.00 CZK")
            Cells(Sonik_pr + 2 + i, 3).value = SonikFormatDate(rst1![DATUM_NACERPANI]) 'RISK-1131
            Cells(Sonik_pr + 2 + i, 4).value = rst1![cislo_uctu]
            Cells(Sonik_pr + 2 + i, 5).value = format$(rst1![Amt_Instalment], "#,##0.00 CZK")
            Cells(Sonik_pr + 2 + i, 6).value = format$(rst1![NESPLACENA_JISTINA], "#,##0.00 CZK")
        
            rst1.MoveNext
        Next i
        
        Range(Cells(Sonik_pr + 2, 1), Cells(Sonik_pr + 1 + i, 6)).Select
        Call Sonik_borders
        
        
        Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekci
        
        End If
rst1.Close



End Sub



Private Sub Sonik_NBU_mimSplatky(sql As String)
Dim i As Integer
Set rst1 = New ADODB.Recordset
rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst1.EOF Then
        rst1.MoveFirst
        Worksheets("Sonik_NBU").Activate
        Sonik_pr = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
        Cells(Sonik_pr + 1, 1).value = "Mimoøádné splátky"
        Cells(Sonik_pr + 2, 1).value = "Èíslo úvìru"
        Cells(Sonik_pr + 2, 2).value = "Výše splátky"
        Cells(Sonik_pr + 2, 3).value = "Datum splátky"
        Range(Cells(Sonik_pr + 1, 1), Cells(Sonik_pr + 2, 3)).Font.Bold = True
        
        For i = 1 To rst1.RecordCount
            Cells(Sonik_pr + 2 + i, 1).value = rst1![code_loan_number]
            Cells(Sonik_pr + 2 + i, 2).value = format$(rst1![AMT_TRANSACTION], "#,##0.00 CZK")
            Cells(Sonik_pr + 2 + i, 3).value = SonikFormatDate(rst1![dtime_transaction_generation]) 'RISK-1131
            rst1.MoveNext
        Next i
        
        Range(Cells(Sonik_pr + 2, 1), Cells(Sonik_pr + 1 + i, 3)).Select
        Call Sonik_borders
        Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekcí
    End If
rst1.Close
End Sub
















Private Sub Sonik_AO_ucty(sql As String, typ As String)

Dim i As Integer

Set rst1 = New ADODB.Recordset
rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   If Not rst1.EOF Then
        rst1.MoveFirst
        Worksheets("Sonik_NBU").Activate
        Sonik_pr = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
        Cells(Sonik_pr + 1, 1).value = typ
        Cells(Sonik_pr + 2, 1).value = "Èíslo úètu"
        Cells(Sonik_pr + 2, 2).value = "Zùstatek úètu"
        Cells(Sonik_pr + 2, 3).value = "Datum podání žádosti"
        Cells(Sonik_pr + 2, 4).value = "Kanál podání žádosti"
        Cells(Sonik_pr + 2, 5).value = "Datum aktivace"
        Cells(Sonik_pr + 2, 6).value = "Stav úètu"
        Cells(Sonik_pr + 2, 7).value = "Mìna"
        Range(Cells(Sonik_pr + 1, 1), Cells(Sonik_pr + 2, 7)).Font.Bold = True
        For i = 1 To rst1.RecordCount
            Cells(Sonik_pr + 2 + i, 1).value = rst1![cislouctu]
            Cells(Sonik_pr + 2 + i, 2).value = format$(rst1![ucetni_zustatek], "#,##0.00")
            Cells(Sonik_pr + 2 + i, 3).value = SonikFormatDate(rst1![dtime_of_application]) 'RISK-1131
            Cells(Sonik_pr + 2 + i, 4).value = rst1![APP_CHANNEL_TYPE_REQUEST]
            If rst1![NAME_BANK_ACCOUNT_STATUS] <> "Stornovaný" Then
            Cells(Sonik_pr + 2 + i, 5).value = SonikFormatDate(rst1![DATE_BANK_ACCOUNT_ACTIVATION]) 'RISK-1131
            End If
            Cells(Sonik_pr + 2 + i, 6).value = rst1![NAME_BANK_ACCOUNT_STATUS]
            Cells(Sonik_pr + 2 + i, 7).value = rst1![Code_currency]
           
           rst1.MoveNext
        Next i
        Range(Cells(Sonik_pr + 2, 1), Cells(Sonik_pr + 1 + i, 7)).Select
        Call Sonik_borders
        
        'RISK-992
        If i = 1 Then
             Cells(Sonik_pr + 3 + i, 1).value = "Úèet založil majitel úètu " & sonik_name1 & " " & sonik_name2 & ", nar. " & SonikFormatDate(sonik_dateBirth) & ", trvale bytem " & sonik_trvala_adresa & "." 'RISK-1131
        ElseIf i > 1 Then
             Cells(Sonik_pr + 3 + i, 1).value = "Úèty založil majitel úètu " & sonik_name1 & " " & sonik_name2 & ", nar. " & SonikFormatDate(sonik_dateBirth) & ", trvale bytem " & sonik_trvala_adresa & "." 'RISK-1131
        End If
        
        Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekcí
  
   End If
rst1.Close
End Sub
Private Sub Sonik_AO_uvery(sql As String)

Dim i As Integer
Dim uver1 As String
Dim uver2 As String

Set rst1 = New ADODB.Recordset
    rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst1.EOF Then
        rst1.MoveFirst
        Worksheets("Sonik_NBU").Activate
        Sonik_pr = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
        Cells(Sonik_pr + 1, 1).value = "Aktivní úvìry"
        Cells(Sonik_pr + 2, 1).value = "Èíslo úvìru"
        Cells(Sonik_pr + 2, 2).value = "Výše úvìru"
        Cells(Sonik_pr + 2, 3).value = "Datum naèerpání"
        Cells(Sonik_pr + 2, 4).value = "Èíslo úètu - naèerpání úvìru"
        Cells(Sonik_pr + 2, 5).value = "Datum podpisu"
        Cells(Sonik_pr + 2, 6).value = "Stav úvìru"
        Range(Cells(Sonik_pr + 1, 1), Cells(Sonik_pr + 2, 6)).Font.Bold = True
        
        For i = 1 To rst1.RecordCount
            Cells(Sonik_pr + 2 + i, 1).value = rst1![CISLO_UVERU]
            Cells(Sonik_pr + 2 + i, 2).value = format$(rst1![vyse_uveru], "#,##0.00 CZK")
            Cells(Sonik_pr + 2 + i, 3).value = SonikFormatDate(rst1![DATUM_NACERPANI]) 'RISK-1131
            Cells(Sonik_pr + 2 + i, 4).value = rst1![cislo_uctu]
            Cells(Sonik_pr + 2 + i, 5).value = SonikFormatDate(rst1![datum_podpisu_pujcky]) 'RISK-1131
            Cells(Sonik_pr + 2 + i, 6).value = rst1![name_loan_wfstatus]
            'promìnná pro DPD a FPD
            sonik_DPD = Application.max(rst1![cnt_days_past_due_tol_k_max], sonik_DPD)
            If rst1![flag_1st_payment_def] = "Y" Then sonik_FFPD = rst1![flag_1st_payment_def]
            'End If
            rst1.MoveNext
        Next i
        
        Range(Cells(Sonik_pr + 2, 1), Cells(Sonik_pr + 1 + i, 6)).Select
        Call Sonik_borders
        
        If i = 2 Then
            uver1 = "úvìru"
            uver2 = "úvìr"
        Else
            uver1 = "úvìrù"
            uver2 = "úvìry"
        End If
        
        If Sonik_AO_Ins_dne <> DateSerial(2000, 1, 1) Then
            Cells(Sonik_pr + 3 + i, 1).value = sonik_name1 & " " & sonik_name2 & " nesplácí " & uver2 & " øádnì a vèas dohodnutým zpùsobem. Vzhledem k tomu, že bylo schváleno oddlužení,"
            Cells(Sonik_pr + 4 + i, 1).value = "pøihlásili jsme dne " & SonikFormatDate(Sonik_AO_Ins_dne) & " pohledávku do insolvenèního øízení, které je vedené u KONKRÉTNÍ SOUD" 'RISK-1131
            Cells(Sonik_pr + 5 + i, 1).value = "pod sp. zn.: " & Sonik_AO_ins_SZ & "."
                
            Else
                    If Sonik_AO_zalobaPodana And Sonik_AO_PM Then  'Vìtu pøidávat v pøípadì, kdy byla již žaloba podána. "je" - ještì nemáme PM rozhodnutí, "bylo" - již máme PM rozhodnutí. Pokud se vygeneruje vìta o insolvenci, tato vìta se již negeneruje.
                            'Cells(Sonik_pr + 3 + i, 1).Value = sonik_name1 & " " & sonik_name2 & " nesplácí " & uver2 & " øádnì a vèas dohodnutým zpùsobem. Z toho dùvodu byly dané úvìry dne " & Format$(zespl, "d. m. yyyy") & " zesplatnìny. Ve vìci bylo vedeno soudní øízení. Dluh je vymáhán exekuènì."
                            Cells(Sonik_pr + 3 + i, 1).value = sonik_name1 & " " & sonik_name2 & " nesplácí " & uver2 & " øádnì a vèas dohodnutým zpùsobem. Z toho dùvodu byly dané úvìry dne"
                            Cells(Sonik_pr + 4 + i, 1).value = SonikFormatDate(Sonik_AO_zespl) & " zesplatnìny. Ve vìci bylo vedeno soudní øízení. Dluh je vymáhán exekuènì." 'RISK-1131
                        Else
                            If Sonik_AO_zalobaPodana Then
                                Cells(Sonik_pr + 3 + i, 1).value = sonik_name1 & " " & sonik_name2 & " nesplácí " & uver2 & " øádnì a vèas dohodnutým zpùsobem. Z toho dùvodu byly dané úvìry dne"
                                Cells(Sonik_pr + 4 + i, 1).value = SonikFormatDate(Sonik_AO_zespl) & " zesplatnìny. Ve vìci je vedeno soudní øízení. Dluh je vymáhán exekuènì" 'RISK-1131
                                Else
                                    If ((sonik_DPD > 5 And sonik_FFPD = "N") Or (sonik_DPD > 11 And sonik_FFPD = "Y")) Then
                                        Cells(Sonik_pr + 3 + i, 1).value = sonik_name1 & " " & sonik_name2 & " je " & uver1 & " žadatelem a " & uver2 & " nesplácí øádnì a vèas dohodnutým zpùsobem."
                                        Else
                                            If sonik_DPD = 0 Then
                                                Cells(Sonik_pr + 3 + i, 1).value = sonik_name1 & " " & sonik_name2 & " je " & uver1 & " žadatelem a " & uver2 & " splácí øádnì a vèas dohodnutým zpùsobem."
                                                Else: Cells(Sonik_pr + 3 + i, 1).value = sonik_name1 & " " & sonik_name2 & " je " & uver1 & " žadatelem a " & uver2 & " splácí øádnì, avšak nìkdy s nìkolikadenním zpoždìním."
                                            End If
                                    End If
                             
                             
                             End If
                    End If
        
            
        End If
    
    
    
   Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekci
   
   End If
rst1.Close



End Sub
Private Sub Sonik_AO_Zadosti(sql As String)

Dim i As Integer

Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   If Not rst1.EOF Then
    rst1.MoveFirst
    Worksheets("Sonik_NBU").Activate
    Sonik_pr = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
    Cells(Sonik_pr + 1, 1).value = "Žádosti o úvìr: "
    Cells(Sonik_pr + 2, 1).value = "Èíslo žádosti"
    Cells(Sonik_pr + 2, 2).value = "Èíslo úvìru"
    Cells(Sonik_pr + 2, 3).value = "Výše úvìru"
    Cells(Sonik_pr + 2, 4).value = "Datum podání žádosti"
    Cells(Sonik_pr + 2, 5).value = "Kanál žádosti"
    Cells(Sonik_pr + 2, 6).value = "Stav žádosti"
    Range(Cells(Sonik_pr + 1, 1), Cells(Sonik_pr + 2, 6)).Font.Bold = True
    For i = 1 To rst1.RecordCount
        Cells(Sonik_pr + 2 + i, 1).value = rst1![CISLO_ZADOSTI]
        Cells(Sonik_pr + 2 + i, 2).value = rst1![CISLO_UVERU]
        Cells(Sonik_pr + 2 + i, 3).value = format$(rst1![vyse_uveru], "#,##0.00 CZK")
        Cells(Sonik_pr + 2 + i, 4).value = SonikFormatDate(rst1![CAS2SCORING]) 'RISK-1131
        ' Cells(Sonik_pr + 2 + i, 5).value = rst1![KANAL_ZADOSTI_APPROVE] ' zakomentovano do Jezevcika verze 5.2
        Cells(Sonik_pr + 2 + i, 6).value = rst1![VYSLEDEK_ZADOSTI]
        If rst1![VYSLEDEK_ZADOSTI] = "Zamítnutá" Then
                Sonik_AO_Zamitnuta = Sonik_AO_Zamitnuta + 1
        End If
    rst1.MoveNext
    Next i
   Range(Cells(Sonik_pr + 2, 1), Cells(Sonik_pr + 1 + i, 6)).Select
   Call Sonik_borders
    
   End If
rst1.Close

If Sonik_AO_Zamitnuta > 0 Then 'Všechny žádosti jsou posuzovány individuálnì a kritéria rozhodování o poskytnutí úvìru se v prùbìhu
        Cells(Sonik_pr + 2 + i, 1).value = "Všechny žádosti jsou posuzovány individuálnì a kritéria rozhodování o poskytnutí úvìru se v prùbìhu èasu"
            If Sonik_AO_Zamitnuta = i - 1 Then
            Cells(Sonik_pr + 3 + i, 1).value = "mìní. Po posouzení všech dostupných údajù byly všechny žádosti zamítnuty a dané úvìry nebyly poskytnuty. "
            Else: Cells(Sonik_pr + 3 + i, 1).value = "mìní. Po posouzení všech dostupných údajù byly nìkteré žádosti zamítnuty a dané úvìry nebyly poskytnuty. "
            End If
End If
Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekcí

End Sub

Private Sub Sonik_AO_Doklady(sql As String)

Dim i As Integer

Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   If Not rst1.EOF Then
    rst1.MoveFirst
    Worksheets("Sonik_NBU").Activate
    Sonik_pr = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
    Cells(Sonik_pr + 1, 1).value = "Ze strany žadatele byly doloženy následující doklady: "
    For i = 1 To rst1.RecordCount
        Cells(Sonik_pr + 1 + i, 1).value = " - " & rst1![name_document_type] & ", doklad byl vydán dne: " & SonikFormatDate(rst1![date_document_issue]) & "            " 'RISK-1131
    rst1.MoveNext
    Next i
   Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekcí
   End If
rst1.Close
End Sub
''pokud má klient dispo právo k jiným úètùm - jiných klientù
Private Sub Sonik_AO_JeDispo(sql As String)
'RISK-978
'Klient je disponentem u jiných klientù
Dim i As Integer

Dim StrRightFrom As String, StrRightTo As String, BlnRightExpired As Boolean, StrMaMel As String

Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   Worksheets("Sonik_NBU").Activate
    Sonik_pr = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
   If Not rst1.EOF Then
    rst1.MoveFirst


    For i = 1 To rst1.RecordCount
    
        BlnRightExpired = False 'inciálnì
    
        StrRightFrom = rst1![DTIME_RIGHT_VALID_FROM]
        
        If rst1![DTIME_RIGHT_VALID_TO] < date Then
            StrRightTo = rst1![DTIME_RIGHT_VALID_TO]
            BlnRightExpired = True
        End If
        
        If BlnRightExpired Then
            Cells(Sonik_pr + i, 1).value = sonik_name1 & " " & sonik_name2 & "  mìl(a) od " & SonikFormatDate(StrRightFrom) & " do " & StrRightTo & " zøízeno dispozièní právo k úètu è. " & rst1![Text_bank_account_number] & "/3030." 'RISK-1131
        Else: Cells(Sonik_pr + i, 1).value = sonik_name1 & " " & sonik_name2 & " má od " & SonikFormatDate(StrRightFrom) & " zøízeno dispozièní právo k úètu è. " & rst1![Text_bank_account_number] & "/3030." 'RISK-1131
        End If

    rst1.MoveNext
    Next i
   Else
    If rst1.RecordCount < 1 Then
        Cells(Sonik_pr + 1, 1).value = sonik_name1 & " " & sonik_name2 & " nemìl/nemìla a nemá zøízeno dispozièní právo k úètùm jiných osob."
    End If
   End If
   Cells(Sonik_pr + 2 + i, 1).value = "Pohyb finanèních prostøedkù na úètu/úètech za období od " & SonikFormatDate(Sonik_AO_OdPohybFin) & " do souèasnosti zasíláme v pøíloze." 'RISK-1131
   'Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekcí
rst1.Close
End Sub
'pokud k úètùm klienta má dispo právo nìkdo další
'RISK-577



Private Sub Sonik_AO_MaDispo(sql As String)
'RISK-978
'Klient je majitel a sleduji jeho disponenty

Dim radek As Integer
Dim i As Integer


Dim StrRightFrom As String, StrRightTo As String, BlnRightExpired As Boolean, StrDateBirth As String, StrPermAddress As String, StrClientName As String, StrMaMel As String

Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   Worksheets("Sonik_NBU").Activate
   Sonik_pr = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
   If Not rst1.EOF Then
    rst1.MoveFirst
    
    
    For i = 1 To rst1.RecordCount
    
    
        BlnRightExpired = False 'inciálnì
        
        StrClientName = rst1![Name_first] & " " & rst1![Name_last]
    
        StrDateBirth = SonikFormatDate(rst1![Date_Birth]) 'RISK-1131
        StrPermAddress = "" & rst1![name_permanent_address_street] & " " & rst1![text_permanent_address_num] & ", " & rst1![code_permanent_address_zip] & " " & rst1![name_permanent_address_town]
        
        StrRightFrom = SonikFormatDate(rst1![DTIME_RIGHT_VALID_FROM]) 'RISK-1131
        
        StrMaMel = "má" 'iniciálnì
        
        
        If rst1![DTIME_RIGHT_VALID_TO] < date Then
            StrRightTo = SonikFormatDate(rst1![DTIME_RIGHT_VALID_TO]) 'RISK-1131
            BlnRightExpired = True
        Else: StrRightTo = "souèasnosti"
        End If
        
        
        
        
        If BlnRightExpired Then StrMaMel = "mìl(a)"
        
        
        Sonik_pr = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
        If i = 1 Then Sonik_pr = Sonik_pr + 1 'Pro první disponentské právo odskoèím text od tabulky.
        Cells(Sonik_pr + 1, 1).value = "K úètu è. " & rst1![Text_bank_account_number] & "/3030 " & StrMaMel & " od " & StrRightFrom & " do " & StrRightTo & "  zøízeno dispozièní právo " & StrClientName & " nar. " & _
                                      StrDateBirth & ", trvale bytem " & StrPermAddress & "."
        radek = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
        

        
        
    rst1.MoveNext
    Next i
   Else
    
    
    If rst1.RecordCount < 1 Then
        Cells(Sonik_pr + 2, 1).value = "K uvedenému úètu/úètùm nemìla a nemá zøízeno dispozièní právo žádná osoba."
    End If
   End If
      'Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekcí
rst1.Close
End Sub




Private Sub Sonik_AO_Transakce(sql As String)

Dim i As Integer

Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   Worksheets("Sonik_NBU").Activate
    Sonik_pr = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
   If Not rst1.EOF Then
    rst1.MoveFirst
    Cells(Sonik_pr + 1, 1).value = "U klienta evidujeme následující vklady finanèních prostøedkù v èástkách nad " & format$(Sonik_AO_KontrolaCastek, "#,##0.00") & " CZK. "
    Cells(Sonik_pr + 2, 1).value = "Datum splatnosti"
    Cells(Sonik_pr + 2, 2).value = "Èíslo úètu"
    Cells(Sonik_pr + 2, 3).value = "Èástka"
    Cells(Sonik_pr + 2, 4).value = "Mìna"
    Range(Cells(Sonik_pr + 2, 1), Cells(Sonik_pr + 2, 4)).Font.Bold = True
    For i = 1 To rst1.RecordCount
        Cells(Sonik_pr + 2 + i, 1).value = SonikFormatDate(rst1![date_transaction_due]) 'RISK-1131
        Cells(Sonik_pr + 2 + i, 2).value = rst1![Text_bank_account_number] & "/3030" 'RISK-978
        Cells(Sonik_pr + 2 + i, 3).value = format$(rst1![AMT_TRANSACTION], "#,##0.00")
        Cells(Sonik_pr + 2 + i, 4).value = rst1![Code_currency]
    rst1.MoveNext
    Next i
   Range(Cells(Sonik_pr + 2, 1), Cells(Sonik_pr + 1 + i, 4)).Select
   Call Sonik_borders
   Else
    If rst1.RecordCount < 1 Then
        Cells(Sonik_pr + 1, 1).value = "U klienta neevidujeme žádné vklady finanèních prostøedkù v èástkách nad " & format$(Sonik_AO_KontrolaCastek, "#,##0.00") & " CZK. "
    End If
    End If
   Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekcí
rst1.Close
End Sub

Private Sub Sonik_AO_Kontakty(sql As String)
 Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   Worksheets("Sonik_NBU").Activate
   Sonik_pr = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
   If Not rst1.EOF Then
    Cells(Sonik_pr + 1, 1).value = "Kontaktní údaje:                                           "
    Cells(Sonik_pr + 2, 1).value = " - " & "Jméno a pøíjmení: " & sonik_name1 & " " & sonik_name2 & "              "
    Cells(Sonik_pr + 3, 1).value = " - " & "Datum narození: " & SonikFormatDate(rst1![Date_Birth]) 'RISK-1131
    Cells(Sonik_pr + 4, 1).value = " - " & "Trvalá adresa: " & rst1![name_permanent_address_street] & " " & rst1![text_permanent_address_num] & ", " & format(rst1![code_permanent_address_zip], "### ##") & " " & rst1![name_permanent_address_town] & "              "
    Cells(Sonik_pr + 5, 1).value = " - " & "Korespondenèní adresa: " & rst1![NAME_REPORTS_ADDRESS_STREET] & " " & rst1![TEXT_REPORTS_ADDRESS_NUM] & ", " & format(rst1![CODE_REPORTS_ADDRESS_ZIP], "### ##") & " " & rst1![NAME_REPORTS_ADDRESS_TOWN] & "              "
    Cells(Sonik_pr + 6, 1).value = " - " & "Primární telefon: " & rst1![PRIMARY_PHONE]
    Cells(Sonik_pr + 7, 1).value = " - " & "Primární e-mail: " & rst1![PRIMARY_EMAIL] & "              "
                              
   End If
   Cells(Sonik_pr + 1, 1).Font.Bold = True
  Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekcí
rst1.Close
End Sub

















Private Sub Sonik_borders()
'
'ohranièení bunìk
'
    'RISK-992
    With Selection.borders()
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Private Sub Sonik_ExportDoWordu(sablona As String, adr As String, cislo As String, cislo2 As String, Datum As Date, zalozka As String, obsah As Range, podpis As String)

    Dim objWord As Object
    Dim objWordDocument As Object
    Dim objStory As Object
    Dim objToc As Object

    Set objWord = CreateObject("Word.Application")

    With objWord

        'otevøení dokumentu
        'Set objWordDocument = .Documents.Open(ThisWorkbook.Path & _
            "\Dokument.docx")

        'otevøení dokumentu postaveném na šablonì
        On Error GoTo chyba_otevreni_sablony
            Set objWordDocument = .Documents.Add(Template:=sablona)
        On Error GoTo 0
        
        If adr <> "NBU" Then
         With objWordDocument
            .bookmarks("Adresat").Select
            End With
            .Selection.text = adr
            
            'RISK-868 - pokud se jedná o šablonu na platební karty, tak nechci vyplòovat název odesílatele
            If InStr(sablona, "Soník Šablona PK Nemazat.dotx") = 0 Then
                With objWordDocument
                    .bookmarks("ZadostOd").Select
                End With
                .Selection.text = SklonujSZ2Pad(Sonik_AO_ZadostOd) 'RISK-978
            End If
        End If
        
        With objWordDocument
            .bookmarks(zalozka).Select
        End With

        
            obsah.Copy
    
            'propojená tabulka?
            'pøebírat formátování Wordu?
            'kopie ve formátu HTML?
            .Selection.PasteExcelTable False, False, False

        
       With objWordDocument
            .bookmarks("Podpisy").Select
        End With
        .Selection.text = podpis
        With objWordDocument
            .bookmarks("CisloJednaci0").Select
        End With
        .Selection.text = cislo
        With objWordDocument
            .bookmarks("CisloJednaci").Select
        End With
        .Selection.text = cislo2
        With objWordDocument
            .bookmarks("DatumPozadavku").Select
        End With
            .Selection.text = SonikFormatDate(Datum) 'RISK-1131
            
        
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



Public Sub Sonik_Export_Sablona_PK(SelectedCards As Collection, DateCardFrom As Date)
'RISK-868


Dim i As Integer, j As Integer, podpis As String, sablona As String, cisloPozadavku As String, datumPozadavku As Date ', Str_textace As String
'Cislo karty nepotrebuji, kdyz mam kolekci
Dim Var_Cislo_Karty As Variant, Bln_Card_Used As Boolean, Str_Gender As String, Str_Text_Pohlavi As String, Bln_Any_Card_Used As Boolean
Dim MainRange As Range, Bln_Card_Blocked As Boolean, Bln_Card_Cancelled As Boolean, Bln_Card_Unblocked As Boolean

Dim Date_Card_Blocked As Date, Date_Card_Unblocked As Date, Date_Card_Cancelled As Date, Date_Card_Operation As Date
'Pro úèel formátování do šablony
Dim Str_Date_Card_Blocked As String, Str_Date_Card_Unblocked As String, Str_Date_Card_Cancelled As String, Str_Date_Card_Operation As String, Str_DateBirth As String, Str_DateCardFrom As String
Dim Str_Date_Card_Valid_From As String, Str_Date_Card_Valid_To As String, Str_Date_Card_Activation As String
'Pro úèel pøenesení se úplnì jinam v pomocném listu, hodím si to nìkam dolù do rohu
Dim Int_PB_Row As Integer, Int_PB_Row_Init As Integer, Int_PB_Col As Integer, Int_PB_Col_Init As Integer
Dim JezevecBook As Workbook, NewBook As Workbook
Dim Str_Cislo_Karty As String

Set JezevecBook = ThisWorkbook
JezevecBook.Worksheets("Sonik_NBU").Activate
sablona = "W:\Vymahani\LATE\Vzory\Soník Šablona PK Nemazat.dotx"

sonik_name1 = Sonik_Karta_Klienta.ListView_Client_Actual.ListItems(1).text
sonik_name2 = Sonik_Karta_Klienta.ListView_Client_Actual.ListItems(1).SubItems(1)
Str_Gender = Sonik_Karta_Klienta.ListView_Client_Actual.ListItems(1).SubItems(10)




Str_DateCardFrom = format$(DateCardFrom, "d. m. yyyy HH:MM")

Str_DateBirth = SonikFormatDate(Sonik_Karta_Klienta.ListView_Client_Actual.ListItems(1).SubItems(5)) 'RISK-1131

datumPozadavku = SonikFormatDate(Sonik_Karta_Klienta.ListView_Soucinnosti_Klienta.SelectedItem.SubItems(2)) 'RISK-1131
cisloPozadavku = Sonik_Karta_Klienta.ListView_Soucinnosti_Klienta.SelectedItem.SubItems(5)


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
'End If
'RISK-1217
podpis = SonikSignature

'Nastavení pomocného excel sheetu
Worksheets("Sonik_NBU").Visible = False
Application.ScreenUpdating = False
Worksheets("Sonik_NBU").Visible = True
Worksheets("Sonik_NBU").Activate
Worksheets("Sonik_NBU").UsedRange.ClearContents
Worksheets("Sonik_NBU").UsedRange.ClearFormats

Sonik_pr = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row

'Èást jméno, pøíjmení dat. narození tuènì
Cells(Sonik_pr, 1).value = sonik_name1 & " " & sonik_name2 & ", nar. " & Str_DateBirth
'Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekcí


'Všechno musí být v loopu pøes posloupnost karet, které uživatel vybral

Bln_Any_Card_Used = False


For Each Var_Cislo_Karty In SelectedCards
        
        JezevecBook.Worksheets("Sonik_NBU").Activate
        
        Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekcí
        Sonik_pr = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
        
        'Zjistim, jestli byla karta v daném období použita, podívám se do listview z raw daty z Homéra
        Bln_Card_Used = False
        'Pokud daná karta byla ve vybraném období alespoò jednou použita, tak pøiøadím pomocnou promìnnou, podle toho dìlám dále rùznou textaci
        For i = 1 To Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems.count
          Date_Card_Operation = Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems(i).SubItems(1)
          If Date_Card_Operation >= DateCardFrom And Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems(i).text = Var_Cislo_Karty Then
            Bln_Card_Used = True
          End If
        Next i
        
        
        'Podívám se, jestli byla daná karta zablokována nebo zrušena
        Bln_Card_Blocked = False
        Bln_Card_Unblocked = False
        Bln_Card_Cancelled = False
        
        'Projedu danou kartou a pokud najdu datum blokace, tak ulozim do pomocne promenne
        For i = 1 To Sonik_Karta_Klienta.ListView_Cards.ListItems.count
          If Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(1) = Var_Cislo_Karty And Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(7) <> vbNullString Then
            Bln_Card_Blocked = True
                'Tady udìlám maximální datum.. projedu všechno a pokud ta hodnota, kterou právì držím, bude vìtší než ta kterou již mám tak uložím. Na zaèátku bude platit.
                If Date_Card_Blocked < CDate(Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(7)) Then
                    Date_Card_Blocked = Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(7)
                    Str_Date_Card_Blocked = SonikFormatDate(Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(7)) 'RISK-1131
                End If
          End If
          
          If Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(1) = Var_Cislo_Karty And Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(9) <> vbNullString Then
            Bln_Card_Unblocked = True
                 'Tady udìlám maximální datum.. projedu všechno a pokud ta hodnota, kterou právì držím, bude vìtší než ta kterou již mám tak uložím. Na zaèátku bude platit.
                 If Date_Card_Unblocked < CDate(Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(9)) Then
                     Date_Card_Unblocked = Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(9)
                     Str_Date_Card_Unblocked = SonikFormatDate(Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(9)) 'RISK-1131
                 End If
          End If
          
          If Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(1) = Var_Cislo_Karty And Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(8) <> vbNullString Then
            Bln_Card_Cancelled = True
            'Tady udìlám maximální datum.. projedu všechno a pokud ta hodnota, kterou právì držím, bude vìtší než ta kterou již mám tak uložím. Na zaèátku bude platit.
                 If Date_Card_Cancelled < CDate(Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(8)) Then
                     Date_Card_Cancelled = Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(8)
                     Str_Date_Card_Cancelled = SonikFormatDate(Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(8)) 'RISK-1131
                 End If
          End If
          
          If Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(1) = Var_Cislo_Karty Then
            Str_Cislo_Karty = Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(1)
            Str_Date_Card_Valid_From = SonikFormatDate(Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(4)) 'RISK-1131
            Str_Date_Card_Valid_To = SonikFormatDate(Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(5)) 'RISK-1131
            Str_Date_Card_Activation = SonikFormatDate(Sonik_Karta_Klienta.ListView_Cards.ListItems(i).SubItems(6)) 'RISK-1131
          End If
        Next i
        
        
        If Not Bln_Card_Cancelled Then
            Str_Text_Pohlavi = "má"
        ElseIf Str_Gender = vbNullString Then
            Str_Text_Pohlavi = "mìl(a)"
        ElseIf Str_Gender = "Muž" Then
            Str_Text_Pohlavi = "mìl"
        Else:
            Str_Text_Pohlavi = "mìla"
        End If
                
        
        
        
        If DateCardFrom = "1.1.1000" Then
        
            If Bln_Card_Used Then
                Bln_Any_Card_Used = True
                Cells(Sonik_pr + 1, 1).value = sonik_name1 & " " & sonik_name2 & " " & Str_Text_Pohlavi & " platební kartu Mastercard è. " & Var_Cislo_Karty & ", která do dnešního dne byla použita, viz pøiložený Pøehled plateb."
            Else:
                Cells(Sonik_pr + 1, 1).value = sonik_name1 & " " & sonik_name2 & " " & Str_Text_Pohlavi & " platební kartu Mastercard è. " & Var_Cislo_Karty & ", která do dnešního dne nebyla vyzkoušena ani jinak zneužita."
            End If
        
        Else:
            
            If Bln_Card_Used Then
                Bln_Any_Card_Used = True
                Cells(Sonik_pr + 1, 1).value = sonik_name1 & " " & sonik_name2 & " " & Str_Text_Pohlavi & " platební kartu Mastercard è. " & Var_Cislo_Karty & ", která za období od " & Str_DateCardFrom & " hod. do dnešního dne byla použita, viz pøiložený Pøehled plateb."
            Else:
                Cells(Sonik_pr + 1, 1).value = sonik_name1 & " " & sonik_name2 & " " & Str_Text_Pohlavi & " platební kartu Mastercard è. " & Var_Cislo_Karty & ", která za období od " & Str_DateCardFrom & " hod. do dnešního dne nebyla vyzkoušena ani jinak zneužita."
            End If
        
        End If
        
        
        
        
        If Bln_Card_Blocked And Bln_Card_Cancelled Then
            If Date_Card_Blocked = Date_Card_Cancelled Then
                Cells(Sonik_pr + 2, 1).value = "Tato karta byla dne " & Str_Date_Card_Blocked & " zablokována a následnì zrušena"
                Else: Cells(Sonik_pr + 2, 1).value = "Tato karta byla dne " & Str_Date_Card_Blocked & " zablokována a následnì dne " & Str_Date_Card_Cancelled & " zrušena."
            End If
        
        ElseIf Not Bln_Card_Blocked And Bln_Card_Cancelled Then
            Cells(Sonik_pr + 2, 1).value = "Tato karta byla dne " & Str_Date_Card_Cancelled & " zrušena."
        
        ElseIf Bln_Card_Blocked And Not Bln_Card_Cancelled Then
        
                If Bln_Card_Unblocked Then
                      Cells(Sonik_pr + 2, 1).value = "Tato karta byla dne " & Str_Date_Card_Blocked & " zablokována a dne " & Str_Date_Card_Unblocked & " opìt aktivována."
                Else: Cells(Sonik_pr + 2, 1).value = "Tato karta byla dne " & Str_Date_Card_Blocked & " zablokována."
                End If
            
        ElseIf Not Bln_Card_Blocked And Not Bln_Card_Cancelled Then
            Cells(Sonik_pr + 2, 1).value = "Tato karta je stále aktivní a nikdy nebyla blokována."
        End If



'Plnìní pomocných excelù - hodím se na listì nìkam dolù do rohu, abych si nepokazil to co už tam mám
    Int_PB_Row_Init = 500
    Int_PB_Col_Init = 50

If Bln_Card_Used And Date_Card_Operation >= DateCardFrom Then
    

    
    'Pøesun na sheetu nìkam daleko, kde nic jiného nikdy nebude
    

    
    Int_PB_Row = Int_PB_Row_Init
    Int_PB_Col = Int_PB_Col_Init
    
    'Vyplnìní hlavièky ke kartì
    
    Cells(Int_PB_Row, Int_PB_Col).value = "Èíslo karty"
    'Cells(Int_PB_Row, Int_PB_Col + 1).value = "Platnost od/do"
    'Cells(Int_PB_Row, Int_PB_Col + 2).value = "Aktivace karty"
    Cells(Int_PB_Row, Int_PB_Col + 1).value = "Blokace karty"
    Cells(Int_PB_Row, Int_PB_Col + 2).value = "Zrušení karty"
    
    Int_PB_Row = Int_PB_Row + 1
    Cells(Int_PB_Row, Int_PB_Col).NumberFormat = "@"
    Cells(Int_PB_Row, Int_PB_Col).value = "'" & Str_Cislo_Karty 'Èíslo karty se zobrazovalo nìkdy špatnì, jako 5E+15...
    'Cells(Int_PB_Row, Int_PB_Col + 1).value = Str_Date_Card_Valid_From & " / " & Str_Date_Card_Valid_To
    'Cells(Int_PB_Row, Int_PB_Col + 2).value = Str_Date_Card_Activation
    Cells(Int_PB_Row, Int_PB_Col + 1).value = Str_Date_Card_Blocked
    Cells(Int_PB_Row, Int_PB_Col + 2).value = Str_Date_Card_Cancelled


    'Vyplnìní hlavièky k transakcím
    Int_PB_Row = Int_PB_Row + 3
    
    Cells(Int_PB_Row, Int_PB_Col).value = "Datum a èas provedení"
    Cells(Int_PB_Row, Int_PB_Col + 1).value = "Platba"
    Cells(Int_PB_Row, Int_PB_Col + 2).value = "Stav"
    Cells(Int_PB_Row, Int_PB_Col + 3).value = "Èástka v CZK" 'RISK-1239
    'Cells(Int_PB_Row, Int_PB_Col + 4).value = "Mìna" 'RISK-1239
    Cells(Int_PB_Row, Int_PB_Col + 4).value = "Obchodník"
    Cells(Int_PB_Row, Int_PB_Col + 5).value = "Adresa"
    Cells(Int_PB_Row, Int_PB_Col + 6).value = "Mìsto"
    Cells(Int_PB_Row, Int_PB_Col + 7).value = "Zemì"
    Cells(Int_PB_Row, Int_PB_Col + 8).value = "Stát"
    Cells(Int_PB_Row, Int_PB_Col + 9).value = "ID Terminálu"
    Cells(Int_PB_Row, Int_PB_Col + 10).value = "Help Date To delete"
    
    'Samotné vyplnìní tabulky - podívám se do listviewèka, kde to mám nahrané
    'Pokud ale øazení v excelu nebude fungovat, potom musím otevøít recordset a seøadit pomocí SQL
    For i = 1 To Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems.count
          j = 1
          If Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems(i).text = Var_Cislo_Karty And CDate(Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems(i).SubItems(1)) >= DateCardFrom Then 'RISK-1131
          
            Int_PB_Row = Int_PB_Row + j
            
            Cells(Int_PB_Row, Int_PB_Col).value = Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems(i).SubItems(1)
            Cells(Int_PB_Row, Int_PB_Col + 1).value = Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems(i).SubItems(2)
            Cells(Int_PB_Row, Int_PB_Col + 2).value = Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems(i).SubItems(3)
            
            Cells(Int_PB_Row, Int_PB_Col + 3).value = Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems(i).SubItems(4)

                        
            
            Cells(Int_PB_Row, Int_PB_Col + 4).value = Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems(i).SubItems(5)
            Cells(Int_PB_Row, Int_PB_Col + 5).value = Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems(i).SubItems(6)
            Cells(Int_PB_Row, Int_PB_Col + 6).value = Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems(i).SubItems(7)
            Cells(Int_PB_Row, Int_PB_Col + 7).value = Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems(i).SubItems(8)
            Cells(Int_PB_Row, Int_PB_Col + 8).value = Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems(i).SubItems(9)
            Cells(Int_PB_Row, Int_PB_Col + 9).value = Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems(i).SubItems(10)
            'Upravené datum pro pozdìjší seøazení, aby fungovalo i v EXCEL 2010
            Cells(Int_PB_Row, Int_PB_Col + 10).value = format(Sonik_Karta_Klienta.ListView_Transactions_Homer.ListItems(i).SubItems(1), "yyyymmddHHMMSS") 'RISK-1239
            j = j + 1
          End If
    
    Next i
    
    

    
    
    'Pøehraju do nového excelu a promažu abych uvolnil místo dalším kartám

    
    
    
    Range(Cells(Int_PB_Row_Init, Int_PB_Col), Cells(Int_PB_Row, Int_PB_Col + 10)).Copy 'RISK-1239
    Set NewBook = Workbooks.Add
'    With NewBook
'        .SaveAs fileName:="Book2.xls"
'    End With
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
    
    'Nastavím formát sloupce na datum a seøadím pomocí pomocného sloupeèku
    Dim lastrow As Long
    lastrow = Cells(Rows.count, 1).End(xlUp).Row
    Range("A6:A" & lastrow).NumberFormat = "[$-F800]dddd, mmmm dd, yyyy"
    Range("A6:K" & lastrow).Sort key1:=Range("K6:K" & lastrow), order1:=xlAscending, Header:=xlNo 'RISK-1239
    
    'Pokud se mi stane, že tam budou dva øádky transakcí ve stejný èas, tak je to asi chyba v napárování - potom buòky zvýrazním
    Call HighlightDuplicateCells(Range("A6:A" & lastrow))


    'Barvièky
    With Range("A1:C1").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    'Barvièky
    With Range("A5:J5").Interior 'RISK-1239
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    
    'Vyrovnání sloupèekù
    Range("A:J").Columns.AutoFit 'RISK-1239
    
    'Vymazání pomocného sloupeèku
    Range("K5:K" & lastrow).ClearContents 'RISK-1239
    
    
    Dim r As Variant
    
    For Each r In Range("D6:D" & lastrow)
    If IsNumeric(r) Then
       r.value = Round(CSng(r.value), 2)
       
        'Pokud je èástka bez haléøù tak vypsat ve formátu 2100, pokud s haléøy tak 2100.99
        If Round(r.value) = r.value Then
            r.NumberFormat = "General"
        Else: r.NumberFormat = "0.00"
        End If

    End If
    Next
     
     

     
    
    'Vyèištìní pomocného listu
    JezevecBook.Worksheets("Sonik_NBU").Activate
    Range(Cells(Int_PB_Row_Init, Int_PB_Col), Cells(Int_PB_Row_Init + 1, Int_PB_Col + 9)).ClearContents 'RISK-1239
    
    
       
    
    
End If




Next Var_Cislo_Karty

Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekcí

Sonik_pr = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row



If Bln_Any_Card_Used Then

Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekcí
'Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekcí

Cells(Sonik_pr + 1, 1).value = "Pøílohy:"

Cells(Sonik_pr + 2, 1).value = "     -     Pøehled plateb"
Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekcí
'Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekcí
'Cells(Sonik_pr + 1, 1).EntireRow.Insert 'vložení øádku pøed celou sekcí
End If

Worksheets("Sonik_NBU").Range("A" & Worksheets("Sonik_NBU").Rows.count).End(xlUp).Select

Worksheets("Sonik_NBU").Cells(1, 1).Font.Bold = True
Worksheets("Sonik_NBU").Cells(Sonik_pr + 2, 1).Font.Bold = True
Range(Selection, Cells(1, 1)).Select

With Selection
        .Font.name = "Arial"
        .Font.Size = 10
        '.HorizontalAlignment = xlLeft
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
Set MainRange = Range(Selection, Cells(1, 1))


Call Sonik_ExportDoWordu(sablona, Sonik_AO_Adresa, cisloPozadavku, cisloPozadavku, datumPozadavku, "Hlavnioddil", MainRange, podpis)


'Závìreèné vyèištìní pomocného excel sheetu
Worksheets("Sonik_NBU").UsedRange.ClearContents
Worksheets("Sonik_NBU").UsedRange.ClearFormats
JezevecBook.Worksheets("Sonik_NBU").Activate
JezevecBook.Worksheets("Sonik_NBU").Range("A1").Select
Worksheets("Sonik_NBU").Visible = False


Application.ScreenUpdating = True

End Sub





