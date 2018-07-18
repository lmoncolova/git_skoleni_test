Attribute VB_Name = "mod_Export_text_zaloby_OLD"
'''Option Explicit
'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''
'''' FA 12.3.2018: Jez.v.5.3 - button ktery vola sablonu jsem zneaktivnil, protoze stejne nefungoval
'''' FA 11.7.2018: Jez.v.5.9:
''''  - kod je tady temer celej dvakrat, a pak je jeste jednou na tretim miste (Sub export_sablona_epr v mod_Export_EPR_XML), kde se i vyuziva
''''  - tenhle kod se uz dloouhou dobu nikde nevyuziva, tudiz asi nebude fungovat jak je potreba
''''  - taky je napsat jeste starym stylem, neortogonalni, duplicitni, necitelni, ...
''''  - a proto nevidim duvod jej drzet
''''  - celej modul proto slavnostne zakomentovavam. Jéééé
'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''
'''
'''
'''Private rst61 As ADODB.Recordset
'''Private sql7 As String
'''Private pr As Integer
'''Private udalost As String
'''Private text21 As String
'''Private objSelection As Object
'''    Private datum_posl_zapl_splatky_po_zes As String
'''    Private dz_1 As Variant
'''    Private dpp As Variant
'''    'Private text45 As String
'''    'Private text4 As String
'''    Private AFolder As String
'''    Private BFolder As String
'''    Private CFolder As String
'''    Private AAFolder As String
'''    Private BBFolder As String
'''    Private CCFolder As String
'''Private fso As Variant
'''Private ii As Integer, qq As Integer, qi As Integer, ij As Integer, iii As Integer, q As Integer, qe As Integer
'''
'''
'''Sub export_text_zaloby_zena(hu As String, SKP As Variant)
'''' do verze 5.3: Sub export_sablona_zena(hu As String, skp As Variant)  v MainModule
'''' FA prejmenoval jsem to aby nazev nebyl tak vagni
'''
'''Dim sql As String, sql2 As String, sql3 As String, sql4 As String
'''Dim datum_p As String
'''Dim skp_ba As String
'''Dim skp_loan As String
'''Dim code As String, dodatek As String, dodatek_ucet As String, datum_ucet As String, contract_number As String
'''Dim ba As String, datum_nacerpani_pujcky As String, amt_loan As String
'''Dim text1 As String, dukaz1 As String
'''Dim rate_interest As Double
'''
'''Dim AFolder As String
'''Dim BFolder As String
'''Dim CFolder As String
'''Dim AAFolder As String
'''Dim BBFolder As String
'''Dim CCFolder As String
'''Dim fso As Variant
'''Dim nazev_exportovaneho_souboru As String
'''
'''On Error GoTo connection_problem
'''
''''name = ThisWorkbook.name
''''MsgBox ThisWorkbook.name
'''Workbooks(ThisWorkbook.name).Activate
'''
''''Worksheets("Sablona").Visible = True
'''
'''
''''Worksheets("Sablona").Activate
'''
'''
'''
'''
'''
'''
'''
'''
''''--------------Struktura zaloby--------------------
'''Dim t1 As Integer, d1 As Integer
'''
'''
'''
''''------------------------------------------------------------------------------------------------------------------------------------------------------
''''--------------CAST I ---------------------------------------------------------------------------------------------------
''''------------------------------------------------------------------------------------------------------------------------------------------------------
'''Worksheets("Sablona").Activate
'''
''''nejdrive zjistime informace o pujce
'''sql = "select * from " & schema & ".sablona_epr where code_loan_number = '" & hu & "' and code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING')"
''' Set rst1 = New ADODB.Recordset
'''   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''   If Not rst1.EOF Then
'''   datum_p = "" & rst1(1) & ""
'''   skp_ba = rst1(2)
'''   skp_loan = rst1(3)
'''   code = rst1(5)
'''   dodatek = rst1(7)
'''   End If
'''rst1.Close
'''
'''Cells(1, 13).value = skp_ba
'''
''''kdy byl zalozen ucet
'''sql2 = "select * from " & schema & ".sablona_epr where skp_bank_account = " & skp_ba & " and code_completion_subject_type = 'CREATE_BU'"
''' Set rst2 = New ADODB.Recordset
'''   rst2.Open sql2, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''   If Not rst2.EOF Then
'''   contract_number = rst2(0)
'''   datum_ucet = "" & rst2(1) & ""
'''   dodatek_ucet = rst2(7)
'''    ba = rst2(9)
'''   End If
'''rst2.Close
'''
'''
'''''''kdy byla zalozena RS  ' FA vyhozeno v jez.v.5.8, doplneno do objektu Client
''''''sql3 = "select dtime_signature from " & schema & ".sablona_epr where id_person = " & SKP & "  and code_completion_subject_type = 'RS' AND TEXT_CONTRACT_NUMBER = " & c.TextContractNumber
'''''''sql3 = "select case when dtime_signature is null then dtime_completion_finished else dtime_signature end as dr from " & schema & ".sablona_epr where id_person = " & skp & " and code_completion_subject_type = 'RS'"
'''''' Set rst3 = New ADODB.Recordset
''''''   rst3.Open sql3, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''''''   If Not rst3.EOF Then
''''''   datum_RS = "" & rst3![Dtime_signature] & ""
''''''   End If
''''''rst3.Close
'''
'''
'''
''''''    FA - nasledujici sql dotaz vyhozen v ramci RISK-726 do Jez.v.5.3, informace se nenahravi z tabulky SP7 ktera je zbytecna, ale z objektu Loan
''''''sql4 = " select * from " & schema & ".sp7 where code_loan_number = '" & hu & "'"
''''''Set rst4 = New ADODB.Recordset
''''''   rst4.Open sql4, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''''''   If Not rst4.EOF Then
''''''    rst4.MoveFirst
'''        datum_nacerpani_pujcky = c.GetLoans(-1, hu).Date_loan_utilization ' "" & rst4(0) & ""
'''        amt_loan = c.GetLoans(-1, hu).Amt_loan_limit ' rst4(1)
'''        rate_interest = c.GetLoans(-1, hu).Rate_Interest_Standard ' rst4(3)
'''''' '   Cells(r + 2, 3).Value = "Úroková sazba:  " & Format(rst7(3), "#,##0.00") & " % p. a."
'''''''Cells(r + 3, 3).Value = "RPSN:  " & Format(rst7(4), "#,##0.00") & " % p. a."
''''''    End If
''''''rst4.Close
'''
'''text1 = "I." & Chr(10)
'''text1 = text1 & "alobce uzavøel s alovanou dne " & c.Date_Contract_Signed & " Rámcovou smlouvu è. " & contract_number & _
'''    " v souladu s Obchodními podmínkami Air Bank a.s. - èlánek Vznik rámcové smlouvy."
'''
'''If dodatek = 0 Then
'''    '---------------------------RS, uver a ucet ve stejny den -------------------
'''    text1 = text1 & " Na základì Rámcové smlouvy byl alované " & _
'''    "aktivován bìnı úèet è. " & ba & "/3030 a alovaná jej mohla zaèít plnì vyuívat." & _
'''    " Podpisem Rámcové smlouvy alovaná souèasnì poádala alobce o úvìr ve vıši " & format(amt_loan, "#,###") & ",- Kè, viz. " & _
'''    "Rámcová smlouva è. " & contract_number & " - Smlouva o úvìru è. " & hu & "."
'''    dukaz1 = "Rámcová smlouva è. " & contract_number & " - Smlouva o úvìru è. " & hu
'''Else
'''    If dodatek_ucet = 0 Then
'''    '---------------------------RS, a ucet ve stejny den, uver jindy -------------------
'''    text1 = text1 & " Na základì Rámcové smlouvy byl alované " & _
'''    "aktivován bìnı úèet è. " & ba & "/3030 a alovaná jej mohla zaèít plnì vyuívat." & _
'''    " Dne " & datum_p & " uzavøela alovaná s alobcem Dodatek è. " & dodatek & " k Rámcové smlouvì, jeho podpisem" & _
'''    " alovaná poádala o úvìr ve vıši " & format(amt_loan, "#,###") & ",- Kè, viz. Smlouva o úvìru è. " & hu & "."
'''    dukaz1 = "Rámcová smlouva è. " & contract_number & Chr(10) & _
'''    "Dodatek è. " & dodatek & " k Rámcové smlouvì - Smlouva o úvìru è. " & hu
'''    ElseIf datum_ucet <> datum_p Then
'''    '---------------------------RS, uver a ucet ve jine dny-------------------
'''    text1 = text1 & " Dne " & datum_ucet & " uzavøela alovaná s alobcem Dodatek è. " & dodatek_ucet & " k Rámcové smlouvì," & _
'''    "na základì kterého jí byl zøízen úèet è. " & ba & "/3030 viz. Dodatek è. " & dodatek_ucet & " k Rámcové smlouvì - Smlouva o úètu è. " & ba & _
'''    ". Dne " & datum_p & " uzavøela alovaná s alobcem Dodatek è. " & dodatek & " k Rámcové smlouvì, jeho podpisem" & _
'''    " alovaná poádala o úvìr ve vıši " & format(amt_loan, "#,###") & ",- Kè, viz. Dodatek è. " & dodatek & " k Rámcové smlouvì - Smlouva o úvìru è. " & hu & "."
'''    dukaz1 = "Rámcová smlouva è. " & contract_number & Chr(10) & _
'''    "Dodatek è. " & dodatek_ucet & " k Rámcové smlouvì - Smlouva o úètu è. " & ba & Chr(10) & _
'''    "Dodatek è. " & dodatek & " k Rámcové smlouvì - Smlouva o úvìru è. " & hu
'''
'''    ElseIf datum_ucet = datum_p Then
'''    '---------------------------RS v jiny den, ucet a uver stejny den-------------------
'''    text1 = text1 & " Dne " & datum_p & " uzavøela alovaná s alobcem Dodatek è. " & dodatek & " k Rámcové smlouvì, jeho podpisem" & _
'''    " alovaná poádala o úvìr ve vıši " & format(amt_loan, "#,###") & ",- Kè a na základì kterého si zøídil úèet è. " & ba & "/3030, viz. Dodatek è. " & dodatek & " k Rámcové smlouvì - Smlouva o úvìru è. " & hu & "."
'''    dukaz1 = "Rámcová smlouva è. " & contract_number & Chr(10) & _
'''    "Dodatek è. " & dodatek & " k Rámcové smlouvì - Smlouva o úvìru è. " & hu
'''    End If
'''End If
'''text1 = text1 & " Protoe alovaná splnila všechny smluvní pøedpoklady pro poskytnutí úvìru v souladu s Obchodními podmínkami a Podmínkami pro pouívání úvìru" & _
'''        " spoleènosti Air Bank a.s., byl jí úvìr è. " & hu & " schválen a dne " & datum_nacerpani_pujcky & " vyplacen na bìnı úèet alované v plné vıši, viz. pøiloenı vıpis z úètu. " & Chr(10)
'''dukaz1 = "Dùkaz: " & Chr(10) & dukaz1 & Chr(10) & _
'''"Obchodní podmínky Air Bank a.s." & Chr(10) & _
'''"Podmínky pro pouívání úvìru Air Bank a.s." & Chr(10) & _
'''"Mimoøádnı vıpis z úètu alované" & Chr(10) & _
'''"Formuláø pro standardní informace o spotøebitelském úvìru" & Chr(10)
'''
'''
'''
''''------------------------------------------------------------------------------------------------------------------------------------------------------
''''--------------CAST II -------------------------------------------------------------------------------------------------------------------------------
''''------------------------------------------------------------------------------------------------------------------------------------------------------
'''
'''Dim text2 As String, dukaz2 As String, sql5 As String, datum_zrizeni_uver As String
'''Dim splatka1 As Double, d_prv_s As String, d_pos_spl As String, d_prav_spl As String, pocet As Double
'''Dim splatka2 As Double, d_prv_s2 As String, d_pos_spl2 As String, d_prav_spl2 As String, pocet2 As Double
'''
'''
'''sql5 = "select * from " & schema & ".sablona_epr5 where skp_loan = " & skp_loan & " and code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING') and dtime_valid_from = (select max(dtime_valid_from) maxx  from " & schema & ".sablona_epr4 where skp_loan = " & skp_loan & " and code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING') group by skp_loan)"
''' Set rst5 = New ADODB.Recordset
'''   rst5.Open sql5, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''   If Not rst5.EOF Then
'''   rst5.MoveFirst
'''   splatka1 = rst5(2)
'''   d_prv_s = rst5(3)
'''   d_pos_spl = rst5(4)
'''   d_prav_spl = rst5(5)
'''   pocet = rst5(6)
'''   End If
'''rst5.Close
'''
'''
'''
'''text2 = "II." & Chr(10) & "Podle podmínek úvìrové smlouvy byla alovaná zavázána mìsíènì hradit pravidelnou konstantní splátku úvìru v celkové vıši " & Isinteger2(splatka1) & _
'''        " Kè, skládající se ze splátky jistiny a úroku " & rate_interest & " % p. a., jeho vıše byla sjednaná Smlouvou o úvìru," & _
'''        " a to vdy k " & d_prav_spl & ". dni mìsíce poèínaje dnem " & d_prv_s & " s tím, e pøedpokládané datum splacení celého úvìru bylo stanoveno" & _
'''        " na " & d_pos_spl & "." & Chr(10) ', viz. " & chr(10)
'''
'''
'''
''''If dodatek = 0 Then
''''text2 = text2 & "Rámcová smlouva è. " & contract_number & " - Smlouva o úvìru è. " & hu & chr(10)
''''Else
''''text2 = text2 & "Dodatek è. " & dodatek & " k Rámcové smlouvì - Smlouva o úvìru è. " & hu & chr(10)
''''End If
'''
'''
''''- nyni je potreba zjistit, zda doslo pred prvni splatkou ke zmene vysi splatky
'''
'''Dim sql6 As String, date_change_amt As String
'''sql6 = "select * from " & schema & ".sablona_epr5 where code_completion_subject_type =  'CHANGE_INST_AMOUNT' and dtime_valid_from < min_pujcky and skp_loan = " & skp_loan
'''Set rst6 = New ADODB.Recordset
'''   rst6.Open sql6, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''   If Not rst6.EOF Then
'''   rst6.MoveFirst
'''   For i = 1 To rst6.RecordCount
'''   If rst6(2) <> splatka1 Then
'''    text2 = text2 & "alovaná si však pøed první splátkou úvìru zmìnila vıši splátky z " & Isinteger2(splatka1) & " Kè na " & Isinteger2(CDbl(rst6(2))) & _
'''    " Kè a tím došlo ke zmìnì délky splácení úvìru z délka splácení úvìru z pùvodních " & pocet & " mìsícù na " & rst6(6) & " mìsícù." & Chr(10)
'''
'''    If rst6(2) <= datum_nacerpani_pujcky Then
'''    MsgBox "Pozor, datum dodatku je menší rovno naèerpání pùjèky, je moné, e budou chybné datumy v dodatcích!!"
'''
'''    End If
'''    rst6.MoveNext
'''    End If
'''    Next i
'''   End If
'''rst6.Close
'''
'''
'''
'''
''''zde doplnime, pokud doslo ke zmene datumu splatek jeste pred prvni splatkou
'''
'''Dim sql61 As String, dukaz_dod As String
'''sql61 = "select * from " & schema & ".sablona_epr5 where code_completion_subject_type =  'CHANGE_INST_DATE' and dtime_valid_from < min_pujcky and skp_loan = " & skp_loan
'''Set rst61 = New ADODB.Recordset
'''   rst61.Open sql61, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''   If Not rst61.EOF Then
'''   For i = 1 To rst61.RecordCount
'''   rst61.MoveFirst
'''        text2 = text2 & " Dne " & rst61(1) & " uzavøela alovaná s alobcem Dodatek è. " & rst61(9) & " k Rámcové smlouvì, na základì kterého " & _
'''        "byla provedena zmìna data splátek z " & d_prav_spl & ". dne v mìsíci na " & rst61(5) & ". den v mìsíci. První splátka s touto zmìnou probìhla " & rst61(3) & "."
'''        d_prav_spl = rst61(5)
'''                 dukaz_dod = dukaz_dod & "Zmìnovı dodatek è. " & rst61(9) & ". k Rámcové smlouvì " & Chr(10)
'''
'''
'''    If rst61(2) <= datum_nacerpani_pujcky Then
'''    MsgBox "Pozor, datum dodatku je menší rovno naèerpání pùjèky, je moné, e budou chybné datumy v dodatcích!!"
'''    End If
'''        If rst61(6) <> pocet Then
'''    text2 = text2 & " Zmìnila se i délka splácení úvìru z pùvodních " & pocet & " mìsícù na " & rst61(6) & " mìsícù." & Chr(10)
'''    pocet = rst61(6)
'''    Else: text2 = text2 & Chr(10)
'''    End If
'''
'''
'''
'''    rst61.MoveNext
'''    Next i
'''
'''   End If
'''rst61.Close
'''
'''
'''
'''
'''
'''Worksheets("Sablona_data").Activate
''''vsechny dalsi zmeny, vyleju do tabulky, ze ktere to potom vezmu
'''sql7 = "select * from " & schema & ".sablona_epr5 where skp_loan = '" & skp_loan & "' and code_completion_subject_type in ('CHANGE_INST_AMOUNT','CHANGE_INST_DATE') and dtime_valid_from >= min_pujcky order by dtime_valid_from, text_appendix_number asc"
''' Set rst7 = New ADODB.Recordset
'''    rst7.Open sql7, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''    If Not rst7.EOF Then
'''        rst7.MoveFirst
'''         For i = 1 To rst7.RecordCount 'radky
'''
'''             For j = 1 To 11 'sloupce
'''                Cells(i, j).value = "" & rst7(j - 1) & ""
'''              Next j
'''              rst7.MoveNext
'''            If i > 1 Then
'''                If format(Cells(i, 8).value, "dd/mmm/yy") = format(Cells(i - 1, 8).value, "dd/mmm/yy") Then
'''            End If
'''
'''              MsgBox "Pozor - existují dodatky se stejnım datumem. je moné, e nebudou správnì informace o splátkovém kalendáøi.", vbExclamation
'''              End If
'''
'''        Next i
'''
'''        'If rst7(8) = "CHANGE_INST_DATE" Then 'zmena splatky
'''
'''        'text2 = text2 & "Dne " & rst7(7) & "< uzavøela alovaná/uzavøel alovanı> s alobcem Dodatek è. rst7(9) k Rámcové smlouvì, jeho podpisemna základì kterého " & _
'''                "byla provedena o zmìnu data splátek z rst7(5). dne v mìsíci na <nové datum splátky>. den v mìsíci. První splátka s touto zmìnou probìhla <datum první splátky se zmìnou data splácení>."
'''
'''
'''
'''       ' Else
'''
'''        'Dne <datum podpisu dodatku>< uzavøela alovaná/uzavøel alovanı> s alobcem Dodatek è. <èíslo dodatku> k Rámcové smlouvì, na základì kterého byla provedena zmìnajeho podpisem <poádala/poádal> o zmìnu vıše pravidelné splátky z <pùvodní vıše splátky>Kè na <nová vıše splátky>Kè. První splátka s touto zmìnou probìhla <datum první splátky se zmìnou vıše splátky>.
'''   pr = rst7.RecordCount
'''   rst7.Close
'''
'''
'''    For i = 1 To pr
'''    udalost = Cells(i, 9).value
'''        If udalost = "CHANGE_INST_DATE" Then
'''            If i = 1 Then
'''            text2 = text2 & " Dne " & Cells(i, 2).value & " uzavøela alovaná s alobcem Dodatek è. " & Cells(i, 10).value & " k Rámcové smlouvì, na základì kterého " & _
'''                "byla provedena zmìna data splátek z " & d_prav_spl & ". dne v mìsíci na " & Cells(i, 6).value & ". den v mìsíci. První splátka s touto zmìnou probìhla " & Cells(i, 4).value & "."
'''            dukaz_dod = dukaz_dod & "Zmìnovı dodatek è. " & Cells(i, 10).value & ". k Rámcové smlouvì " & Chr(10)
'''
'''                If pocet <> Cells(i, 7).value Then
'''                text2 = text2 & " Zmìnila se i délka splácení úvìru z pùvodních " & pocet & " mìsícù na " & Cells(i, 7).value & " mìsícù." & Chr(10)
'''                Else: text2 = text2 & Chr(10)
'''                End If
'''
'''            Else
'''            text2 = text2 & " Dne " & Cells(i, 2).value & " uzavøela alovaná s alobcem Dodatek è. " & Cells(i, 10).value & " k Rámcové smlouvì, na základì kterého " & _
'''                "byla provedena zmìna data splátek z " & Cells(i - 1, 6).value & ". dne v mìsíci na " & Cells(i, 6).value & ". den v mìsíci. První splátka s touto zmìnou probìhla " & Cells(i, 4).value & "."
'''            dukaz_dod = dukaz_dod & "Zmìnovı dodatek è. " & Cells(i, 10).value & ". k Rámcové smlouvì " & Chr(10)
'''                If Cells(i - 1, 7).value <> Cells(i, 7).value Then
'''                text2 = text2 & " Zmìnila se i délka splácení úvìru z pùvodních " & Cells(i - 1, 7).value & " mìsícù na " & Cells(i, 7).value & " mìsícù." & Chr(10)
'''                Else: text2 = text2 & Chr(10)
'''                End If
'''
'''            End If
'''        Else
'''            If i = 1 Then
'''            text2 = text2 & " Dne " & Cells(i, 2).value & " uzavøela alovaná s alobcem Dodatek è. " & Cells(i, 10).value & " k Rámcové smlouvì, na základì kterého byla provedena zmìna " & _
'''            "vıše pravidelné splátky z " & Isinteger2(splatka1) & " Kè na " & Isinteger2(Cells(i, 3).value) & " Kè. První splátka s touto zmìnou probìhla " & Cells(i, 4).value & "."
'''                If pocet <> Cells(i, 7).value Then
'''                text2 = text2 & " Zmìnila se i délka splácení úvìru z pùvodních " & pocet & " mìsícù na " & Cells(i, 7).value & " mìsícù." & Chr(10)
'''                Else: text2 = text2 & Chr(10)
'''                End If
'''            dukaz_dod = dukaz_dod & "Zmìnovı dodatek è. " & Cells(i, 10).value & ". k Rámcové smlouvì " & Chr(10)
'''
'''            Else
'''            text2 = text2 & " Dne " & Cells(i, 2).value & " uzavøela alovaná s alobcem Dodatek è. " & Cells(i, 10).value & " k Rámcové smlouvì, na základì kterého byla provedena zmìna " & _
'''            "vıše pravidelné splátky z " & Isinteger2(Cells(i - 1, 3).value) & "Kè na " & Isinteger2(Cells(i, 3).value) & "Kè. První splátka s touto zmìnou probìhla " & Cells(i, 4).value & "."
'''                If Cells(i - 1, 7).value <> Cells(i, 7).value Then
'''                text2 = text2 & " Zmìnila se i délka splácení úvìru z pùvodních " & Cells(i - 1, 7).value & " mìsícù na " & Cells(i, 7).value & " mìsícù." & Chr(10)
'''                Else: text2 = text2 & Chr(10)
'''                End If
'''            dukaz_dod = dukaz_dod & "Zmìnovı dodatek è. " & Cells(i, 10).value & ". k Rámcové smlouvì " & Chr(10)
'''
'''            End If
'''        End If
'''    Next i
'''    'Range(Cells(1, 1), Cells(pr, 11)).Clear
'''End If
'''
'''Worksheets("Sablona_data").Visible = False
''''Worksheets("Sablona").Activate
''''--------------CAST II.II -------------------------------------------------------------------------------------------------------------------------------
'''
''''musime rozhodnout, zda platil vcas nebo splatky splatil s nekolikadennim zpozdenim
'''
'''Dim sql11 As String, text2_dukaz As String, text2ms() As String, datumy_ms() As Date ', rst11 As ADODB.Recordset
'''sql11 = "select * from " & schema & ".sablona_epr8 where skp_loan = " & skp_loan
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql11, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''      If Not rst1.EOF Then
'''        rst1.MoveFirst
'''        ReDim text2ms(rst1.RecordCount)
'''        ReDim datumy_ms(rst1.RecordCount)
'''            For i = 1 To rst1.RecordCount
'''            text2ms(i) = "Dne " & rst1(13) & " provedla alovaná mimoøádnou splátku úvìru ve vıši " & Isinteger2(CDbl(rst1(12))) & " Kè., viz Zmìnovı dodatek è. " & rst1(7) & " k Rámcové smlouvì. "
'''            text2_dukaz = "Zmìnovı dodatek è. " & rst1(7) & " k Rámcové smlouvì - Provedení mimoøádné splátky úvìru èíslo " & hu
'''            datumy_ms(i) = rst1(13)
'''            Next i
'''    End If
'''rst1.Close
'''
'''
'''
'''
'''
''''musime rozhodnout, zda platil vcas nebo splatky splatil s nekolikadennim zpozdenim
'''
'''Dim sql8 As String, datum_posl_zapl_splatky As String, text2ss As String, datum_prvni_splatky As String
'''sql8 = "select min(datum_splatky), max(datum_splatky), max(paid_in_time) from " & schema & ".sablona_epr6 where skp_loan = " & skp_loan & " and platba = splatka group by skp_loan"
'''Set rst8 = New ADODB.Recordset
'''   rst8.Open sql8, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''      If Not rst8.EOF Then
'''        rst8.MoveFirst
'''
'''        If rst8(0) = rst8(1) Then
'''            text2ss = "Splátka " & rst8(0) & " byla splacena. "
'''        Else
'''            text2ss = "Splátky od " & rst8(0) & " do " & rst8(1) & " byly splaceny. "
'''        End If
'''
''''        If rst8(2) = 0 Then '1 je, pokud byly splatky se zpozdenim
''''        text2ss = "Splátky od " & rst8(0) & " do " & rst8(1) & " byly splaceny. "
''''        Else
''''        text2ss = "Splátky od " & rst8(0) & " do " & rst8(1) & " byly splaceny. "  ' , i kdy ne vdy vèas." ', viz. splátková kalendáø úvìru è. " & hu & " a pøehled plateb è. " & hu & "."
''''        End If
'''        datum_posl_zapl_splatky = rst8(1)
'''        datum_prvni_splatky = rst8(0)
'''   End If
'''rst8.Close
'''
'''Dim sql8_1 As String, datum_posl_zapl_splatky_po_zes As String
'''
''''zjistime datum posledni kompletne zaplacene splatky (platby i po zesplatneni)
'''sql8_1 = "select * from " & schema & ".sablona_epr14 where skp_loan = " & skp_loan
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql8_1, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''      If Not rst1.EOF Then
'''        rst1.MoveFirst
'''
'''
'''        datum_posl_zapl_splatky_po_zes = rst1(0)
'''
'''   End If
'''rst1.Close
'''
''''podivame se, kdy bylo datum mimoradne splatky - a podle toho zaradime
'''If text2_dukaz = "" Then
'''text2 = text2 & text2ss
'''Else
'''    For i = 1 To UBound(text2ms)
'''
'''        If datumy_ms(i) < datum_prvni_splatky Then
'''            text2 = text2 & text2ms(i) & text2ss
'''        Else: text2 = text2 & text2ss & text2ms(i)
'''        End If
'''
'''    Next i
'''
'''End If
''''nyni se podivame, zda dosla cast nejake splatky pred zesplatnenim nebo ne
'''
'''
'''Dim sql9 As String, sql10 As String, dz As String, dz_1 As Variant, tedy As String, radne As String, dku As Date
'''sql9 = "select * from " & schema & ".sablona_epr7 where skp_loan = " & skp_loan & " order by datum_platby"
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql9, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''    If Not rst1.EOF Then
'''        rst1.MoveFirst
'''            For i = 1 To rst1.RecordCount
'''            text2 = text2 & "Na splátku dne " & rst1(1) & " bylo ze strany alované dne " & rst1(2) & " zaplaceno pouze " & Isinteger2(CDbl(rst1(5))) & " Kè, pøièem tuto èástku si alobce zapoèetl èástí ve vıši " & _
'''            Isinteger2(CDbl(rst1(3))) & " Kè na úhradu jistiny splátky a èástí ve vıši " & Isinteger2(CDbl(rst1(4))) & " Kè na úhradu úroku splátky. " ' , viz. splátkovı kalendáø úvìru è. " & hu & " a pøehled plateb è. " & hu & ". " 'Následnì neevidujeme ádnou další platbu na úhradu splátek ze strany alované."
'''
'''            If i = 1 Then
'''            dku = rst1(1)
'''
'''           End If
'''            rst1.MoveNext
'''            Next i
'''            tedy = "tedy "
'''            radne = " øádnì a vèas"
'''            rst1.Close
'''    Else 'pokud neprisla zadna castecna splatka
'''    rst1.Close
'''    sql10 = "select min(date_instalment) from " & schema & ".sp3 where code_loan_number = '" & hu & "' group by code_loan_number "
'''    Set rst1 = New ADODB.Recordset
'''    rst1.Open sql10, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''      If Not rst1.EOF Then
'''        rst1.MoveFirst
'''
'''        dku = rst1(0)
'''     Else
'''     MsgBox "Pozor, úvìr nemá zaplacenou ádnou splátku!!"
'''      End If
'''    rst1.Close
'''End If
'''
'''text2 = text2 & "Poèínaje splátkou ke dni " & dku & " " & tedy & " nebyly splátky ze strany alované uhrazeny" & radne & ", viz. splátkovı kalendáø úvìru è. " & hu & " a pøehled plateb è. " & hu & "."
'''
'''text21 = "Následnì neevidujeme ádnou další platbu na úhradu splátek ze strany alované."  'pokud prisla nejaka platba, tak se tam toto psat nebude
'''
'''
''''--------------CAST II.III -------------------------------------------------------------------------------------------------------------------------------
'''Dim sql12 As String, sql14 As String, sql144 As String, jistina_zesp As Double, urok_zesp As Double, dukaz_z_d  As String, text22 As String, don As String
'''
'''
'''
''''potrebujeme zjistit, kdy bylo datum zesplatneni a den pred zesplatnenim
'''
'''sql14 = "select * from " & schema & ".sablona_epr10 where skp_loan = " & skp_loan
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql14, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''   If Not rst1.EOF Then
'''    rst1.MoveFirst
'''    dz = rst1(2)
'''    dz_1 = rst1(3)
'''    End If
'''rst1.Close
'''
'''
'''sql144 = "select * from " & schema & ".sablona_epr12 where id_person = " & SKP
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql144, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''   If Not rst1.EOF Then
'''    rst1.MoveFirst
'''        don = rst1(0)
'''    End If
'''rst1.Close
'''
'''Dim text23 As String, text24 As String
'''sql12 = "select skp_loan, id_person, nvl(jplatba,0), nvl(uplatba,0), nvl(jpredpis,0), nvl(upredpis,0), nvl(ip,0), nvl(ii,0), platba_den_po_zesplatneni from " & schema & ".sablona_epr9 where skp_loan = " & skp_loan
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql12, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''      If Not rst1.EOF Then
'''        rst1.MoveFirst
'''        jistina_zesp = rst1(4) + rst1(6) - rst1(2)
'''        urok_zesp = rst1(5) + rst1(7) - rst1(3)
'''          If rst1(8) = "Y" Then 'prisla platba den po zesplatneni
'''            text22 = "Vzhledem k závanému porušení smluvních povinností stranou alovanou bylo v souladu s Obchodními podmínkami a na základì Podmínek pro pouívání úvìru Air Bank a.s. - èl. Nìkolik bodù k ukonèení smlouvy, odst. 3 a " & _
'''            " rozhodnuto ze strany alobce o zesplatnìní úvìru  z dùvodu neplnìní smluvních podmínek alovanou stranou, viz. zesplatòující dopis - pøedalobní vızva. " ' ze dne " & Day(DataCheck.TextBox1.Text) & "." & Month(DataCheck.TextBox1.Text) & "." & Year(DataCheck.TextBox1.Text) & chr(10)
'''           dukaz_z_d = "Zesplatòující dopis - pøedalobní vızva" 'musime take upravit podobu dukazu s ohledem na platbu den po zesplatneni
'''          Else
'''          dukaz_z_d = "Zesplatòující dopis - pøedalobní vızva"
'''            text22 = "Vzhledem k závanému porušení smluvních povinností stranou alovanou bylo v souladu s Obchodními podmínkami a na základì Podmínek pro pouívání úvìru Air Bank a.s. - èl. Nìkolik bodù k ukonèení smlouvy, odst. 3 a - dne " & dz & _
'''            " rozhodnuto ze strany alobce o zesplatnìní úvìru  z dùvodu neplnìní smluvních podmínek alovanou stranou, viz. zesplatòující dopis - pøedalobní vızva ze dne " & dz & "." & Chr(10)
'''            text24 = "Ke dni zesplatnìní tedy byla dluná pohledávka tvoøena nesplacenou jistinou ve vıši " & format(jistina_zesp, "#,##0.00") & " Kè a nesplacenım smluvním úrokem z úvìru (kapitalizovanı úrok) ve vıši " & format(urok_zesp, "#,##0.00") & " Kè." 'toto se tam da, pokud byla platba po zesplatneni
'''
'''            text23 = "Ke dni zesplatnìní tedy byla dluná pohledávka tvoøena nesplacenou jistinou ve vıši " & format(jistina_zesp, "#,##0.00") & " Kè a nesplacenım smluvním úrokem z úvìru (kapitalizovanı úrok) ve vıši " & format(urok_zesp, "#,##0.00") & " Kè" & _
'''            ", kterı je souhrnem nesplacenıch smluvních úrokù ode dne poslední øádne zaplacené splátky, tj. ode dne " & datum_posl_zapl_splatky & " do dne pøedcházejícího dni zesplatnìní, tedy do dne " & dz_1 & ", konkrétnì pak (poèet dnù prodlení/365 x aktuální jistina x " & rate_interest & "/100)." & Chr(10) 'pokud nebyla platba po zesplatneni
'''
'''        End If
'''    End If
'''rst1.Close
'''
'''Dim sql17 As String, sql13 As String, text25 As String, text222 As String
'''sql13 = "select * from " & schema & ".sp1 where skp_loan = " & skp_loan & " order by datum_platby, datum_splatky "
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql13, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''      If Not rst1.EOF Then
'''        rst1.MoveFirst
'''        For i = 1 To rst1.RecordCount
'''            If i = 1 Then
'''                  text222 = "Po zesplatnìní úvìru uhradila alovaná dne " & rst1(2) & " èástku ve vıši " & Isinteger2(CDbl(rst1(13) + rst1(14))) & " Kè, kterou si alobce zapoèetl na úhradu dluné pohledávky, " & _
'''                      "pøièem tuto èástku si alobce zapoèetl èástí ve vıši " & Isinteger2(CDbl(rst1(13))) & " Kè na úhradu jistiny splátky ze dne " & rst1(1) & " a èástí ve vıši " & Isinteger2(CDbl(rst1(14))) & " Kè na úhradu úroku splátky ze dne " & rst1(1) & "."
'''            Else
'''                  text222 = text222 & Chr(10) & Chr(10) & "Dále po zesplatnìní úvìru uhradila alovaná dne " & rst1(2) & " èástku ve vıši " & rst1(13) + rst1(14) & " Kè, kterou si alobce zapoèetl na úhradu dluné pohledávky, " & _
'''                  "pøièem tuto èástku si alobce zapoèetl èástí ve vıši " & Isinteger2(CDbl(rst1(13))) & " Kè na úhradu jistiny splátky ze dne " & rst1(1) & " a èástí ve vıši " & Isinteger2(CDbl(rst1(14))) & " Kè na úhradu úroku splátky ze dne " & rst1(1) & "."
'''            End If
'''            rst1.MoveNext
'''        Next i
'''       rst1.Close
'''        'zjistime info o aktualni pohledavce
'''
'''
'''
'''    sql17 = "select * from " & schema & ".balance_obs where skp_loan = " & skp_loan
'''    Set rst2 = New ADODB.Recordset
'''    rst2.Open sql17, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''        If Not rst2.EOF Then
'''            rst2.MoveFirst
'''            If IsNull(datum_posl_zapl_splatky_po_zes) Then
'''                text25 = "Aktuální vıše pohledávky je tvoøena nesplacenou jistinou ve vıši " & format(rst2(2), "#,##0.00") & " Kè a nesplacenım smluvním úrokem z úvìru do zesplatnìní (kapitalizovanı úrok) ve vıši " & format(rst2(3), "#,##0.00") & " Kè, kterı je souhrnem nesplacenıch smluvních úrokù ode dne naèerpání úvìru do dne pøedcházejícího dni zesplatnìní, tedy do dne " & dz_1 & _
'''                         ", konkrétnì pak (poèet dnù prodlení/365 x aktuální jistina x " & rate_interest & "/100)."
'''            Else
'''                text25 = "Aktuální vıše pohledávky je tvoøena nesplacenou jistinou ve vıši " & format(rst2(2), "#,##0.00") & " Kè a nesplacenım smluvním úrokem z úvìru do zesplatnìní (kapitalizovanı úrok) ve vıši " & format(rst2(3), "#,##0.00") & " Kè, kterı je souhrnem nesplacenıch smluvních úrokù ode dne poslední zaplacené splátky, tj. ode dne " & datum_posl_zapl_splatky_po_zes & ", do dne pøedcházejícího dni zesplatnìní, tedy do dne " & dz_1 & _
'''                         ", konkrétnì pak (poèet dnù prodlení/365 x aktuální jistina x " & rate_interest & "/100)."
'''            End If
'''        End If
'''    rst2.Close
'''
'''
'''     text2 = text2 & Chr(10) & Chr(10) & text22 & Chr(10) & text24 & Chr(10) & Chr(10) & text222 & Chr(10) & Chr(10) & text25 & Chr(10) & Chr(10) & text21
'''
'''    Else 'nebyla platba po zesplatneni
'''        rst1.Close
'''        text2 = text2 & Chr(10) & Chr(10) & text21 & Chr(10) & Chr(10) & text22 & Chr(10) & Chr(10) & text23
'''    End If
'''
'''
'''text2 = text2 & Chr(10) & "Dùkaz: " & Chr(10) & _
'''"Splátkovı kalendáø úvìru è. " & hu & Chr(10) & _
'''"Pøehled plateb úvìru è. " & hu & Chr(10)
'''If text2_dukaz <> "" Then
'''    text2 = text2 & text2_dukaz & Chr(10)
'''End If
'''
'''If dukaz_dod <> "" Then
'''    text2 = text2 & dukaz_dod & Chr(10)
'''End If
'''
'''If dukaz_z_d <> "" Then
'''    text2 = text2 & dukaz_z_d & Chr(10)
'''End If
'''text2 = text2 & "Mimoøádnı vıpis z úètu" & Chr(10) & Chr(10) & "III."
'''
'''
''''------------------------------------------------------------------------------------------------------------------------------------------------------
''''--------------CAST III -------------------------------------------------------------------------------------------------------------------------------
''''------------------------------------------------------------------------------------------------------------------------------------------------------
'''
'''Dim text3 As String, text4 As String, dz1 As String, t3 As Integer
'''
'''dz1 = Day(DataCheck.TextBox1.text) - 1 & "." & Month(DataCheck.TextBox1.text) & "." & Year(DataCheck.TextBox1.text)
'''text3 = "Zesplatòujícím dopisem - pøedalobní vızvou byla alovaná vyzvána k zaplacení celého dluhu, kterı byl tvoøen dlunou jistinou ve " & _
'''"vıši " & format(jistina_zesp, "#,##0.00") & " Kè a nesplacenım smluvním úrokem z úvìru ve vıši " & format(urok_zesp, "#,##0.00") & " Kè, pøièem jejich vıše byla stanovena ke dni " & dz_1 & ". Zesplatnìním se stal závazek splatnım v plné vıši spolu s pøíslušenstvím ke dni " & _
'''dz & "." & Chr(10) & Chr(10) & "alobce alované v zesplatòujícím dopise - pøedalobní vızvì poskytl dodateènou " & _
'''" lhùtu k doplacení, alovaná ji však ke splacení nevyuila. Kdyby alovaná v této lhùtì svùj dluh zaplatila, nepoadoval by alobce " & _
'''"zaplacení smluvního úroku ani zákonného úroku z prodlení ode dne zesplatnìní do dne úhrady." & Chr(10) & Chr(10)
'''
''''uprava 15.2.2017 RISK-505, 420
'''
'''    text3 = text3 & "Smluvní strany si podpisem Rámcové smlouvy sjednaly, viz. Obchodní podmínky - èl. Doruèování dokumentù, e alobce doruèuje alované dokumenty " & _
'''    "na adresu alované. Dokumenty odeslané poštou se povaují  za " & _
'''    "doruèené desátım dnem ode dne jejich odeslání."
'''
'''
'''text3 = text3 & Chr(10) & Chr(10) & "Dùkaz: " & Chr(10) & _
'''dukaz_z_d & Chr(10) & _
'''"Podací arch " & Chr(10) & _
'''"Doklad odeslané pošty " & Chr(10)
'''
'''
'''text3 = text3 & "Obchodní podmínky Air Bank a.s." & Chr(10) & Chr(10) & "IV."
'''
''''"Doklad o odeslání/doruèení zesplatòujícího dopisu - pøedalobní vızvy ze dne " & don & Chr(10) & _
'''
'''
'''Dim text41 As String, text42 As String, text45 As String, sql15 As String, sql16 As String, qq As Boolean, start_su As Date, dpp As Variant
'''qq = False
'''sql15 = "select * from " & schema & ".sablona_epr11 where uver_cislo = '" & hu & "' order by datum_platby"
'''sql16 = "select min(start_su2) from " & schema & ".V_DATACHECK_CREDIT_I where uver_cislo = '" & hu & "' group by uver_cislo "
'''
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql16, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''      If Not rst1.EOF Then
'''        rst1.MoveFirst
'''        start_su = rst1(0)
'''    End If
'''rst1.Close
'''
'''
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql15, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''      If Not rst1.EOF Then
'''        rst1.MoveFirst
'''        dpp = rst1(1)
'''        For i = 1 To rst1.RecordCount
'''        text41 = text41 & " do " & rst1(1) & ", spolu se smluvním úrokem " & rate_interest & " % roènì z èástky " & format(rst1(2), "#,##0.00") & " Kè od " & rst1(0)
'''            If rst1(0) >= start_su Then
'''                If qq = False Then
'''                text42 = text42 & " do " & rst1(1) & " spolu se zákonnım úrokem z prodlení z èástky " & format(rst1(2), "#,##0.00") & " Kè od " & rst1(0)
'''                qq = True
'''                Else
'''                text45 = text45 & " do " & rst1(1) & ", spolu se zákonnım úrokem z prodlení z èástky " & format(rst1(2), "#,##0.00") & " Kè od " & rst1(0)
'''                End If
'''            End If
'''        rst1.MoveNext
'''        Next i
'''    End If
'''rst1.Close
'''
'''Dim sql155 As String, jistina_akt As Double, urok_akt As Double, celkem_akt As Double
'''
'''sql155 = "select * from " & schema & ".v_datacheck_credit_a where uver_cislo = '" & hu & "'"
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql155, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''      If Not rst1.EOF Then
'''        jistina_akt = rst1(3)
'''        urok_akt = rst1(4)
'''        celkem_akt = jistina_akt + urok_akt
'''    End If
'''rst1.Close
'''
'''
'''
'''
'''
'''
'''text41 = "Celkovı dluh vyplıvající z úvìrové smlouvy è. " & hu & " ke dni podání aloby èiní na jistinì èástku ve vıši " & format(jistina_akt, "#,##0.00") & " Kè spolu se smluvním úrokem " & rate_interest & _
'''" % roènì z èástky " & format(jistina_zesp, "#,##0.00") & " Kè od " & dz & text41
'''text42 = " a dále zákonnı úrok z prodlení z èástky " & format(jistina_zesp, "#,##0.00") & " Kè od " & start_su & text42
'''
'''text41 = text41 & " do zaplacení"
'''text42 = text42 & text45 & " do zaplacení."
'''
'''text4 = text41 & text42 & " alobce poaduje zákonnı úrok z prodlení a od " & start_su & ", nebo dle interních pravidel poskytuje alobce alované ve vızvì k plnìní dodateènou lhùtu. " & _
'''"Dále alobce poaduje kapitalizovanı smluvní úrok ve vıši " & format(urok_akt, "#,##0.00") & " Kè za období od " & dz_1 & " do " & dz1 & "."
'''
'''
'''
'''Dim text As String
'''
'''
'''
'''text = text1 & Chr(10) & dukaz1 & Chr(10) & text2 & Chr(10) & text3 & Chr(10) & text4
'''
'''
'''    Dim objWord
'''
'''    Dim objDoc
'''
'''    Set objWord = CreateObject("Word.Application")
'''
'''    Set objDoc = objWord.Documents.Add
'''
'''    objWord.Visible = True
'''    Set objSelection = objWord.Selection
'''    objSelection.text = text
'''
'''    '   objDoc.SaveAs ("D:\Users\pboril1374ab\Documents\al.docx")
'''    'Worksheets("Sablona").Cells(1, 1).Value = ""
'''    'Worksheets("Sablona").Cells(1, 1).Value = Text
'''    '
'''    '
'''     AFolder = cfg.Cesta("EprPodaneZaloby") & c.Nazev_slozky
'''     BFolder = cfg.Cesta("EprNachystane") & c.Nazev_slozky  ' "W:\Vymahani\LATE\EPR\NACHYSTANÉ\"
'''     CFolder = "W:\Vymahani\LATE\EPR\NACHYSTANÉ BØEZEN 2015\" & c.Nazev_slozky
'''
'''     AAFolder = cfg.Cesta("EprPodaneZaloby") & c.Nazev_slozky & "Pøílohy\"
'''     BBFolder = cfg.Cesta("EprNachystane") & c.Nazev_slozky & "Pøílohy\"
'''     CCFolder = "W:\Vymahani\LATE\EPR\NACHYSTANÉ BØEZEN 2015\" & c.Nazev_slozky & "Pøílohy\"
'''     nazev_exportovaneho_souboru = "aloba úvìru è. " & hu & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".docx"
'''
'''Set fso = CreateObject("Scripting.FileSystemObject")
'''
'''If fso.FolderExists(AAFolder) Then
'''    objDoc.SaveAs (AAFolder & nazev_exportovaneho_souboru)
'''    Call Shell("explorer.exe " & AAFolder, vbNormalFocus)
'''
'''ElseIf fso.FolderExists(AFolder) Then
'''    objDoc.SaveAs (AFolder & nazev_exportovaneho_souboru)
'''    Call Shell("explorer.exe " & AFolder, vbNormalFocus)
'''
'''ElseIf fso.FolderExists(BBFolder) Then
'''    objDoc.SaveAs (BBFolder & nazev_exportovaneho_souboru)
'''    Call Shell("explorer.exe " & BBFolder, vbNormalFocus)
'''
'''ElseIf fso.FolderExists(BFolder) Then
'''    objDoc.SaveAs (BFolder & nazev_exportovaneho_souboru)
'''    Call Shell("explorer.exe " & BFolder, vbNormalFocus)
'''
'''ElseIf fso.FolderExists(CCFolder) Then
'''    objDoc.SaveAs (CCFolder & nazev_exportovaneho_souboru)
'''    Call Shell("explorer.exe " & CCFolder, vbNormalFocus)
'''
'''ElseIf fso.FolderExists(CFolder) Then
'''    objDoc.SaveAs (CFolder & nazev_exportovaneho_souboru)
'''    Call Shell("explorer.exe " & CFolder, vbNormalFocus)
'''
'''Else
'''    MsgBox "Sloka pro klienta " & c.name & " neexistuje." ', vbInformation, "Sloka nenalezena!"
'''    'Range(Cells(1, 1), Cells(150, 9)).ExportAsFixedFormat Filename:="aloba úvìru è." & hu & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".pdf", Type:=xlTypePDF, _
'''            Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
'''            :=False, OpenAfterPublish:=True
'''    objDoc.SaveAs (cfg.Cesta("EprKlientiBezSlozky") & nazev_exportovaneho_souboru)
'''    Call Shell("explorer.exe " & cfg.Cesta("EprKlientiBezSlozky"), vbNormalFocus)
'''
'''End If
'''
'''
'''Worksheets("Sablona").Visible = False
'''Exit Sub
'''connection_problem:
'''
'''If Mid(Err.Description, 1, 9) = "ORA-03114" Then
'''err_connection = True
'''Else
''' MsgBox Err.Description
'''End If
'''
'''End Sub
'''
'''Sub export_text_zaloby_muz(hu As String, SKP As Variant)
'''' do verze 5.3: Sub export_sablona_muz(hu As String, skp As Variant)  v MainModule
'''' FA prejmenoval jsem to aby nazev nebyl tak vagni
'''
'''Dim sql As String, sql2 As String, sql3 As String, sql4 As String, datum_p As String, skp_ba As String, skp_loan As String
'''Dim code As String, dodatek As String, dodatek_ucet As String, datum_ucet As String, contract_number As String
'''Dim ba As String, datum_nacerpani_pujcky As String, amt_loan As String
'''Dim text1 As String, dukaz1 As String
'''Dim rate_interest As Double
'''Dim text45 As String
'''Dim text4 As String
'''
'''On Error GoTo connection_problem
'''
''''name = ThisWorkbook.name
''''MsgBox ThisWorkbook.name
'''Workbooks(ThisWorkbook.name).Activate
'''
''''Worksheets("Sablona").Visible = True
'''
'''
'''Worksheets("Sablona").Activate
'''
'''
'''
'''
'''
'''
'''
'''
''''--------------Struktura zaloby--------------------
'''
''''------------------------------------------------------------------------------------------------------------------------------------------------------
''''--------------CAST I ---------------------------------------------------------------------------------------------------
''''------------------------------------------------------------------------------------------------------------------------------------------------------
''''Worksheets("Sablona").Activate
'''
''''nejdrive zjistime informace o pujce
'''sql = "select * from " & schema & ".sablona_epr where code_loan_number = '" & hu & "' and code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING')"
''' Set rst1 = New ADODB.Recordset
'''   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''   If Not rst1.EOF Then
'''   datum_p = "" & rst1(1) & ""
'''   skp_ba = rst1(2)
'''   skp_loan = rst1(3)
'''   code = rst1(5)
'''   dodatek = rst1(7)
'''   End If
'''rst1.Close
'''
'''Cells(1, 13).value = skp_ba
'''
''''kdy byl zalozen ucet
'''sql2 = "select * from " & schema & ".sablona_epr where skp_bank_account = " & skp_ba & " and code_completion_subject_type = 'CREATE_BU'"
''' Set rst2 = New ADODB.Recordset
'''   rst2.Open sql2, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''   If Not rst2.EOF Then
'''   contract_number = rst2(0)
'''   datum_ucet = "" & rst2(1) & ""
'''   dodatek_ucet = rst2(7)
'''    ba = rst2(9)
'''   End If
'''rst2.Close
'''
'''
'''''''kdy byla zalozena RS   ' FA vyhozeno v jez.v.5.8, doplneno do objektu Client
''''''sql3 = "select dtime_signature from " & schema & ".sablona_epr where id_person = " & SKP & "  and code_completion_subject_type = 'RS' AND TEXT_CONTRACT_NUMBER = " & c.TextContractNumber
'''''''sql3 = "select max(case when dtime_signature is null then dtime_completion_finished else dtime_signature end) as dr from " & schema & ".sablona_epr where id_person = " & skp & " and code_completion_subject_type = 'RS'"
'''''' Set rst3 = New ADODB.Recordset
''''''   rst3.Open sql3, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''''''   If Not rst3.EOF Then
''''''   datum_RS = "" & rst3![Dtime_signature] & ""
''''''   End If
''''''rst3.Close
'''
'''
''''''    FA - nasledujici sql dotaz vyhozen v ramci RISK-726 do Jez.v.5.3, informace se nenahravi z tabulky SP7 ktera je zbytecna, ale z objektu Loan
''''''sql4 = " select * from " & schema & ".sp7 where code_loan_number = '" & hu & "'"
''''''Set rst4 = New ADODB.Recordset
''''''   rst4.Open sql4, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
''''''   If Not rst4.EOF Then
''''''    rst4.MoveFirst
'''        datum_nacerpani_pujcky = c.GetLoans(-1, hu).Date_loan_utilization ' "" & rst4(0) & ""
'''        amt_loan = c.GetLoans(-1, hu).Amt_loan_limit '  rst4(1)
'''        rate_interest = c.GetLoans(-1, hu).Rate_Interest_Standard ' rst4(3)
'''''' '   Cells(r + 2, 3).Value = "Úroková sazba:  " & Format(rst7(3), "#,##0.00") & " % p. a."
'''''''Cells(r + 3, 3).Value = "RPSN:  " & Format(rst7(4), "#,##0.00") & " % p. a."
''''''    End If
''''''rst4.Close
'''
'''text1 = "I." & Chr(10)
'''text1 = text1 & "alobce uzavøel s alovanım dne " & c.Date_Contract_Signed & " Rámcovou smlouvu è. " & contract_number & _
'''    " v souladu s Obchodními podmínkami Air Bank a.s. - èlánek Vznik rámcové smlouvy."
'''
'''If dodatek = 0 Then
'''    '---------------------------RS, uver a ucet ve stejny den -------------------
'''    text1 = text1 & " Na základì Rámcové smlouvy byl alovanému " & _
'''    "aktivován bìnı úèet è. " & ba & "/3030 a alovanı jej mohl zaèít plnì vyuívat." & _
'''    " Podpisem Rámcové smlouvy alovanı souèasnì poádal alobce o úvìr ve vıši " & format(amt_loan, "#,###") & ",- Kè, viz. " & _
'''    "Rámcová smlouva è. " & contract_number & " - Smlouva o úvìru è. " & hu & "."
'''    dukaz1 = "Rámcová smlouva è. " & contract_number & " - Smlouva o úvìru è. " & hu
'''Else
'''    If dodatek_ucet = 0 Then
'''    '---------------------------RS, a ucet ve stejny den, uver jindy -------------------
'''    text1 = text1 & " Na základì Rámcové smlouvy byl alovanému " & _
'''    "aktivován bìnı úèet è. " & ba & "/3030 a alovanı jej mohl zaèít plnì vyuívat." & _
'''    " Dne " & datum_p & " uzavøel alovanı s alobcem Dodatek è. " & dodatek & " k Rámcové smlouvì, jeho podpisem" & _
'''    " alovanı poádal o úvìr ve vıši " & format(amt_loan, "#,###") & ",- Kè, viz. Smlouva o úvìru è. " & hu & "."
'''    dukaz1 = "Rámcová smlouva è. " & contract_number & Chr(10) & _
'''    "Dodatek è. " & dodatek & " k Rámcové smlouvì - Smlouva o úvìru è. " & hu
'''    ElseIf datum_ucet <> datum_p Then
'''    '---------------------------RS, uver a ucet ve jine dny-------------------
'''    text1 = text1 & " Dne " & datum_ucet & " uzavøel alovanı s alobcem Dodatek è. " & dodatek_ucet & " k Rámcové smlouvì," & _
'''    "na základì kterého mu byl zøízen úèet è. " & ba & "/3030 viz. Dodatek è. " & dodatek_ucet & " k Rámcové smlouvì - Smlouva o úètu è. " & ba & _
'''    ". Dne " & datum_p & " uzavøel alovanı s alobcem Dodatek è. " & dodatek & " k Rámcové smlouvì, jeho podpisem" & _
'''    " alovanı poádal o úvìr ve vıši " & format(amt_loan, "#,###") & ",- Kè, viz. Dodatek è. " & dodatek & " k Rámcové smlouvì - Smlouva o úvìru è. " & hu & "."
'''    dukaz1 = "Rámcová smlouva è. " & contract_number & Chr(10) & _
'''    "Dodatek è. " & dodatek_ucet & " k Rámcové smlouvì - Smlouva o úètu è. " & ba & Chr(10) & _
'''    "Dodatek è. " & dodatek & " k Rámcové smlouvì - Smlouva o úvìru è. " & hu
'''    ElseIf datum_ucet = datum_p Then
'''    '---------------------------RS v jiny den, ucet a uver stejny den-------------------
'''    text1 = text1 & " Dne " & datum_p & " uzavøel alovanı s alobcem Dodatek è. " & dodatek & " k Rámcové smlouvì, jeho podpisem" & _
'''    " alovanı poádal o úvìr ve vıši " & format(amt_loan, "#,###") & ",- Kè a na základì kterého si zøídil úèet è. " & ba & "/3030, viz. Dodatek è. " & dodatek & " k Rámcové smlouvì - Smlouva o úvìru è. " & hu & "."
'''    dukaz1 = "Rámcová smlouva è. " & contract_number & Chr(10) & _
'''    "Dodatek è. " & dodatek & " k Rámcové smlouvì - Smlouva o úvìru è. " & hu
'''    End If
'''End If
'''text1 = text1 & " Protoe alovanı splnil všechny smluvní pøedpoklady pro poskytnutí úvìru v souladu s Obchodními podmínkami a Podmínkami pro pouívání úvìru" & _
'''        " spoleènosti Air Bank a.s., byl mu úvìr è. " & hu & " schválen a dne " & datum_nacerpani_pujcky & " vyplacen na bìnı úèet alovaného v plné vıši, viz. pøiloenı vıpis z úètu. " & Chr(10)
'''dukaz1 = "Dùkaz: " & Chr(10) & dukaz1 & Chr(10) & _
'''"Obchodní podmínky Air Bank a.s." & Chr(10) & _
'''"Podmínky pro pouívání úvìru Air Bank a.s." & Chr(10) & _
'''"Mimoøádnı vıpis z úètu alovaného" & Chr(10) & _
'''"Formuláø pro standardní informace o spotøebitelském úvìru" & Chr(10)
'''
'''
'''
''''------------------------------------------------------------------------------------------------------------------------------------------------------
''''--------------CAST II -------------------------------------------------------------------------------------------------------------------------------
''''------------------------------------------------------------------------------------------------------------------------------------------------------
'''
'''Dim text2 As String, dukaz2 As String, sql5 As String, datum_zrizeni_uver As String
'''Dim splatka1 As Double, d_prv_s As String, d_pos_spl As String, d_prav_spl As String, pocet As Double
'''Dim splatka2 As Double, d_prv_s2 As String, d_pos_spl2 As String, d_prav_spl2 As String, pocet2 As Double
'''
'''
'''sql5 = "select * from " & schema & ".sablona_epr5 where skp_loan = " & skp_loan & " and code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING') and dtime_valid_from = (select max(dtime_valid_from) maxx  from " & schema & ".sablona_epr4 where skp_loan = " & skp_loan & " and code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING') group by skp_loan)"
''' Set rst5 = New ADODB.Recordset
'''   rst5.Open sql5, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''   If Not rst5.EOF Then
'''   rst5.MoveFirst
'''   splatka1 = rst5(2)
'''   d_prv_s = rst5(3)
'''   d_pos_spl = rst5(4)
'''   d_prav_spl = rst5(5)
'''   pocet = rst5(6)
'''   End If
'''rst5.Close
'''
'''
'''
'''text2 = "II." & Chr(10) & "Podle podmínek úvìrové smlouvy byl alovanı zavázanı mìsíènì hradit pravidelnou konstantní splátku úvìru v celkové vıši " & Isinteger2(splatka1) & _
'''        " Kè, skládající se ze splátky jistiny a úroku " & rate_interest & " % p. a., jeho vıše byla sjednaná Smlouvou o úvìru," & _
'''        " a to vdy k " & d_prav_spl & ". dni mìsíce poèínaje dnem " & d_prv_s & " s tím, e pøedpokládané datum splacení celého úvìru bylo stanoveno" & _
'''        " na " & d_pos_spl & "." & Chr(10) ', viz. " & chr(10)
'''
'''
'''
''''If dodatek = 0 Then
''''text2 = text2 & "Rámcová smlouva è. " & contract_number & " - Smlouva o úvìru è. " & hu & chr(10)
''''Else
''''text2 = text2 & "Dodatek è. " & dodatek & " k Rámcové smlouvì - Smlouva o úvìru è. " & hu & chr(10)
''''End If
'''
'''
''''- nyni je potreba zjistit, zda doslo pred prvni splatkou ke zmene vysi splatky
'''
'''Dim sql6 As String, date_change_amt As String
'''sql6 = "select * from " & schema & ".sablona_epr5 where code_completion_subject_type =  'CHANGE_INST_AMOUNT' and dtime_valid_from < min_pujcky and skp_loan = " & skp_loan
'''Set rst6 = New ADODB.Recordset
'''   rst6.Open sql6, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''   If Not rst6.EOF Then
'''   rst6.MoveFirst
'''   For i = 1 To rst6.RecordCount
'''   If rst6(2) <> splatka1 Then
'''    text2 = text2 & "alovanı si však pøed první splátkou úvìru zmìnil vıši splátky z " & Isinteger2(splatka1) & " Kè na " & Isinteger2(CDbl(rst6(2))) & _
'''    " Kè a tím došlo ke zmìnì délky splácení úvìru z délka splácení úvìru z pùvodních " & pocet & " mìsícù na " & rst6(6) & " mìsícù." & Chr(10)
'''    If rst6(2) <= datum_nacerpani_pujcky Then
'''    MsgBox "Pozor, datum dodatku je menší rovno naèerpání pùjèky, je moné, e budou chybné datumy v dodatcích!!"
'''    End If
'''     rst6.MoveNext
'''
'''    End If
'''    Next i
'''   End If
'''rst6.Close
'''
'''
'''
'''
''''zde doplnime, pokud doslo ke zmene datumu splatek jeste pred prvni splatkou
'''
'''Dim sql61 As String, dukaz_dod As String
'''sql61 = "select * from " & schema & ".sablona_epr5 where code_completion_subject_type =  'CHANGE_INST_DATE' and dtime_valid_from < min_pujcky and skp_loan = " & skp_loan
'''Set rst61 = New ADODB.Recordset
'''   rst61.Open sql61, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''   If Not rst61.EOF Then
'''   For i = 1 To rst61.RecordCount
'''   rst61.MoveFirst
'''        text2 = text2 & " Dne " & rst61(1) & " uzavøel alovanı s alobcem Dodatek è. " & rst61(9) & " k Rámcové smlouvì, na základì kterého " & _
'''        "byla provedena zmìna data splátek z " & d_prav_spl & ". dne v mìsíci na " & rst61(5) & ". den v mìsíci. První splátka s touto zmìnou probìhla " & rst61(3) & "."
'''         dukaz_dod = dukaz_dod & "Zmìnovı dodatek è. " & rst61(9) & ". k Rámcové smlouvì " & Chr(10)
'''        d_prav_spl = rst61(5)
'''
'''        If rst61(2) <= datum_nacerpani_pujcky Then
'''    MsgBox "Pozor, datum dodatku je menší rovno naèerpání pùjèky, je moné, e budou chybné datumy v dodatcích!!"
'''    End If
'''
'''    If rst61(6) <> pocet Then
'''    text2 = text2 & " Zmìnila se i délka splácení úvìru z pùvodních " & pocet & " mìsícù na " & rst61(6) & " mìsícù." & Chr(10)
'''    pocet = rst61(6)
'''    Else: text2 = text2 & Chr(10)
'''    End If
'''     rst61.MoveNext
'''    Next i
'''
'''   End If
'''rst61.Close
'''
'''
'''
'''
'''
'''Worksheets("Sablona_data").Activate
''''vsechny dalsi zmeny, vyleju do tabulky, ze ktere to potom vezmu
'''sql7 = "select * from " & schema & ".sablona_epr5 where skp_loan = '" & skp_loan & "' and code_completion_subject_type in ('CHANGE_INST_AMOUNT','CHANGE_INST_DATE') and dtime_valid_from >= min_pujcky order by dtime_valid_from, text_appendix_number asc"
''' Set rst7 = New ADODB.Recordset
'''    rst7.Open sql7, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''    If Not rst7.EOF Then
'''        rst7.MoveFirst
'''         For i = 1 To rst7.RecordCount 'radky
'''             For j = 1 To 11 'sloupce
'''                Cells(i, j).value = "" & rst7(j - 1) & ""
'''              Next j
'''              rst7.MoveNext
'''
'''              If i > 1 Then
'''                If format(Cells(i, 8).value, "dd/mmm/yy") = format(Cells(i - 1, 8).value, "dd/mmm/yy") Then
'''                    MsgBox "Pozor - existují dodatky se stejnım datumem. je moné, e nebudou správnì informace o splátkovém kalendáøi.", vbExclamation
'''
'''                End If
'''               'MsgBox Format(Cells(i - 1, 8).Value, "dd/mmm/yy") & " " & Cells(i - 1, 8).Value
'''              End If
'''
'''
'''        Next i
'''
'''        'If rst7(8) = "CHANGE_INST_DATE" Then 'zmena splatky
'''
'''        'text2 = text2 & "Dne " & rst7(7) & "< uzavøela alovaná/uzavøel alovanı> s alobcem Dodatek è. rst7(9) k Rámcové smlouvì, jeho podpisemna základì kterého " & _
'''                "byla provedena o zmìnu data splátek z rst7(5). dne v mìsíci na <nové datum splátky>. den v mìsíci. První splátka s touto zmìnou probìhla <datum první splátky se zmìnou data splácení>."
'''
'''
'''
'''       ' Else
'''
'''        'Dne <datum podpisu dodatku>< uzavøela alovaná/uzavøel alovanı> s alobcem Dodatek è. <èíslo dodatku> k Rámcové smlouvì, na základì kterého byla provedena zmìnajeho podpisem <poádala/poádal> o zmìnu vıše pravidelné splátky z <pùvodní vıše splátky>Kè na <nová vıše splátky>Kè. První splátka s touto zmìnou probìhla <datum první splátky se zmìnou vıše splátky>.
'''   pr = rst7.RecordCount
'''   rst7.Close
'''
'''
'''    For i = 1 To pr
'''    udalost = Cells(i, 9).value
'''        If udalost = "CHANGE_INST_DATE" Then
'''            If i = 1 Then
'''                text2 = text2 & " Dne " & Cells(i, 2).value & " uzavøel alovanı s alobcem Dodatek è. " & Cells(i, 10).value & " k Rámcové smlouvì, na základì kterého " & _
'''                "byla provedena zmìna data splátek z " & d_prav_spl & ". dne v mìsíci na " & Cells(i, 6).value & ". den v mìsíci. První splátka s touto zmìnou probìhla " & Cells(i, 4).value & "."
'''                dukaz_dod = dukaz_dod & "Zmìnovı dodatek è. " & Cells(i, 10).value & ". k Rámcové smlouvì "
'''
'''                If pocet <> Cells(i, 7).value Then
'''                    text2 = text2 & " Zmìnila se i délka splácení úvìru z pùvodních " & pocet & " mìsícù na " & Cells(i, 7).value & " mìsícù." & Chr(10)
'''                    Else: text2 = text2 & Chr(10)
'''                End If
'''
'''            Else
'''                text2 = text2 & " Dne " & Cells(i, 2).value & " uzavøel alovanı s alobcem Dodatek è. " & Cells(i, 10).value & " k Rámcové smlouvì, na základì kterého " & _
'''                "byla provedena zmìna data splátek z " & Cells(i - 1, 6).value & ". dne v mìsíci na " & Cells(i, 6).value & ". den v mìsíci. První splátka s touto zmìnou probìhla " & Cells(i, 4).value & "."
'''                dukaz_dod = dukaz_dod & "Zmìnovı dodatek è. " & Cells(i, 10).value & ". k Rámcové smlouvì "
'''
'''                If Cells(i - 1, 7).value <> Cells(i, 7).value Then
'''                    text2 = text2 & " Zmìnila se i délka splácení úvìru z pùvodních " & Cells(i - 1, 7).value & " mìsícù na " & Cells(i, 7).value & " mìsícù." & Chr(10)
'''                    Else: text2 = text2 & Chr(10)
'''                End If
'''
'''            End If
'''        Else
'''            If i = 1 Then
'''                text2 = text2 & " Dne " & Cells(i, 2).value & " uzavøel alovanı s alobcem Dodatek è. " & Cells(i, 10).value & " k Rámcové smlouvì, na základì kterého byla provedena zmìna " & _
'''                "vıše pravidelné splátky z " & Isinteger2(splatka1) & " Kè na " & Isinteger2(Cells(i, 3).value) & " Kè. První splátka s touto zmìnou probìhla " & Cells(i, 4).value & "."
'''                If pocet <> Cells(i, 7).value Then
'''                 text2 = text2 & " Zmìnila se i délka splácení úvìru z pùvodních " & pocet & " mìsícù na " & Cells(i, 7).value & " mìsícù." & Chr(10)
'''                    Else: text2 = text2 & Chr(10)
'''                End If
'''                dukaz_dod = dukaz_dod & "Zmìnovı dodatek è. " & Cells(i, 10).value & ". k Rámcové smlouvì " & Chr(10)
'''
'''                Else
'''                text2 = text2 & " Dne " & Cells(i, 2).value & " uzavøel alovanı s alobcem Dodatek è. " & Cells(i, 10).value & " k Rámcové smlouvì, na základì kterého byla provedena zmìna " & _
'''                "vıše pravidelné splátky z " & Isinteger2(Cells(i - 1, 3).value) & "Kè na " & Isinteger2(Cells(i, 3).value) & "Kè. První splátka s touto zmìnou probìhla " & Cells(i, 4).value & "."
'''                If Cells(i - 1, 7).value <> Cells(i, 7).value Then
'''                    text2 = text2 & " Zmìnila se i délka splácení úvìru z pùvodních " & Cells(i - 1, 7).value & " mìsícù na " & Cells(i, 7).value & " mìsícù." & Chr(10)
'''                    Else: text2 = text2 & Chr(10)
'''                End If
'''            dukaz_dod = dukaz_dod & "Zmìnovı dodatek è. " & Cells(i, 10).value & ". k Rámcové smlouvì " & Chr(10)
'''
'''            End If
'''        End If
'''    Next i
'''    Range(Cells(1, 1), Cells(pr, 11)).Clear
'''End If
'''
'''Worksheets("Sablona_data").Visible = False
''''Worksheets("Sablona").Activate
''''--------------CAST II.II -------------------------------------------------------------------------------------------------------------------------------
'''
'''
'''
''''zda a kdy prisla mimoradna splatka
'''
'''Dim sql11 As String, text2_dukaz As String, text2ms() As String, datumy_ms() As Date
'''sql11 = "select * from " & schema & ".sablona_epr8 where skp_loan = " & skp_loan
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql11, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''      If Not rst1.EOF Then
'''        rst1.MoveFirst
'''        ReDim text2ms(rst1.RecordCount)
'''        ReDim datumy_ms(rst1.RecordCount)
'''            For i = 1 To rst1.RecordCount
'''            text2ms(i) = "Dne " & rst1(13) & " provedl alovanı mimoøádnou splátku úvìru ve vıši " & Isinteger2(CDbl(rst1(12))) & " Kè., viz Zmìnovı dodatek è. " & rst1(7) & " k Rámcové smlouvì. "
'''            text2_dukaz = "Zmìnovı dodatek è. " & rst1(7) & " k Rámcové smlouvì - Provedení mimoøádné splátky úvìru èíslo " & hu
'''            datumy_ms(i) = rst1(13)
'''            Next i
'''    End If
'''rst1.Close
'''
'''
'''
'''
'''
''''musime rozhodnout, zda platil vcas nebo splatky splatil s nekolikadennim zpozdenim
''''zde pouze platby do zesplatneni
'''
'''Dim sql8 As String, sql8_1 As String, datum_posl_zapl_splatky As String, text2ss As String, datum_prvni_splatky As Date
'''
'''sql8 = "select min(datum_splatky), max(datum_splatky), max(paid_in_time) from " & schema & ".sablona_epr6 where skp_loan = " & skp_loan & " and platba = splatka group by skp_loan"
'''Set rst8 = New ADODB.Recordset
'''   rst8.Open sql8, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''      If Not rst8.EOF Then
'''        rst8.MoveFirst
'''
'''        If rst8(0) = rst8(1) Then
'''            text2ss = "Splátka " & rst8(0) & " byla splacena. "
'''        Else
'''            text2ss = "Splátky od " & rst8(0) & " do " & rst8(1) & " byly splaceny. "
'''        End If
'''
''''        If rst8(2) = 0 Then '1 je, pokud byly splatky se zpozdenim
''''        text2ss = "Splátky od " & rst8(0) & " do " & rst8(1) & " byly splaceny. "
''''        Else
''''        text2ss = "Splátky od " & rst8(0) & " do " & rst8(1) & " byly splaceny. "  ' , i kdy ne vdy vèas." ', viz. splátková kalendáø úvìru è. " & hu & " a pøehled plateb è. " & hu & "."
''''        End If
'''        datum_posl_zapl_splatky = rst8(1)
'''        datum_prvni_splatky = rst8(0)
'''   End If
'''rst8.Close
'''
''''zjistime datum posledni kompletne zaplacene splatky (platby i po zesplatneni)
'''
'''sql8_1 = "select * from " & schema & ".sablona_epr14 where skp_loan = " & skp_loan
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql8_1, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''      If Not rst1.EOF Then
'''        rst1.MoveFirst
'''
'''        datum_posl_zapl_splatky_po_zes = rst1(0)
'''
'''   End If
'''rst1.Close
'''
'''
'''
''''podivame se, kdy bylo datum mimoradne splatky - a podle toho zaradime
'''
'''If text2_dukaz = "" Then
'''text2 = text2 & text2ss
'''Else
'''
'''    For i = 1 To UBound(text2ms)
'''
'''        If datumy_ms(i) < datum_prvni_splatky Then
'''            text2 = text2 & text2ms(i) & text2ss
'''        Else: text2 = text2 & text2ss & text2ms(i)
'''        End If
'''
'''    Next i
'''End If
'''
'''
''''nyni se podivame, zda dosla cast nejake splatky pred zesplatnenim nebo ne
'''
'''
'''Dim sql9 As String, sql10 As String, dz As String, tedy As String, radne As String, dku As Date
'''
'''sql9 = "select * from " & schema & ".sablona_epr7 where skp_loan = " & skp_loan & " order by datum_platby"
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql9, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''    If Not rst1.EOF Then
'''        rst1.MoveFirst
'''            For i = 1 To rst1.RecordCount
'''            text2 = text2 & "Na splátku dne " & rst1(1) & " bylo ze strany alovaného dne " & rst1(2) & " zaplaceno pouze " & Isinteger2(CDbl(rst1(5))) & " Kè, pøièem tuto èástku si alobce zapoèetl èástí ve vıši " & _
'''            Isinteger2(CDbl(rst1(3))) & " Kè na úhradu jistiny splátky a èástí ve vıši " & Isinteger2(CDbl(rst1(4))) & " Kè na úhradu úroku splátky. " ', viz. splátkovı kalendáø úvìru è. " & hu & " a pøehled plateb è. " & hu & "." 'Následnì neevidujeme ádnou další platbu na úhradu splátek ze strany alované."
'''
'''            If i = 1 Then
'''            dku = rst1(1)
'''            End If
'''
'''            rst1.MoveNext
'''            Next i
'''            tedy = "tedy" 'slovicko tedy se vlozi, pokud je tam nejaka castecna uhrada splatky
'''                        radne = " øádnì a vèas"
'''    rst1.Close
'''    Else 'pokud neprisla zadna castecna splatka
'''    rst1.Close
'''
'''
'''sql10 = "select min(date_instalment) from " & schema & ".sp3 where code_loan_number = '" & hu & "' group by code_loan_number "
'''  Set rst1 = New ADODB.Recordset
'''    rst1.Open sql10, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''      If Not rst1.EOF Then
'''        rst1.MoveFirst
'''              dku = rst1(0)
'''           Else
'''     MsgBox "Pozor, úvìr nemá zaplacenou ádnou splátku!!"
'''      End If
'''
'''    rst1.Close
'''End If
'''        text2 = text2 & " Poèínaje splátkou ke dni " & dku & " " & tedy & " nebyly splátky ze strany alovaného uhrazeny" & radne & ", viz. splátkovı kalendáø úvìru è. " & hu & " a pøehled plateb è. " & hu & "."
'''
'''text21 = "Následnì neevidujeme ádnou další platbu na úhradu splátek ze strany alovaného."  'pokud prisla nejaka platba, tak se tam toto psat nebude
'''
'''
''''--------------CAST II.III -------------------------------------------------------------------------------------------------------------------------------
'''Dim sql12 As String, sql14 As String, sql144 As String, jistina_zesp As Double, urok_zesp As Double, dukaz_z_d  As String, text22 As String, don As String
'''
'''
'''
''''potrebujeme zjistit, kdy bylo datum zesplatneni a den pred zesplatnenim
'''
'''sql14 = "select * from " & schema & ".sablona_epr10 where skp_loan = " & skp_loan
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql14, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''   If Not rst1.EOF Then
'''    rst1.MoveFirst
'''    dz = rst1(2)
'''    dz_1 = rst1(3)
'''    End If
'''rst1.Close
'''
'''
'''sql144 = "select * from " & schema & ".sablona_epr12 where id_person = " & SKP
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql144, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''   If Not rst1.EOF Then
'''    rst1.MoveFirst
'''        don = rst1(0)
'''    End If
'''rst1.Close
'''
'''
'''
'''Dim text23 As String, text24 As String
'''sql12 = "select skp_loan, id_person, nvl(jplatba,0), nvl(uplatba,0), nvl(jpredpis,0), nvl(upredpis,0), nvl(ip,0), nvl(ii,0), platba_den_po_zesplatneni from " & schema & ".sablona_epr9 where skp_loan = " & skp_loan
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql12, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''      If Not rst1.EOF Then
'''        rst1.MoveFirst
'''        jistina_zesp = rst1(4) + rst1(6) - rst1(2)
'''        urok_zesp = rst1(5) + rst1(7) - rst1(3)
'''          If rst1(8) = "Y" Then 'prisla platba den po zesplatneni
'''            text22 = "Vzhledem k závanému porušení smluvních povinností stranou alovanou bylo v souladu s Obchodními podmínkami a na základì Podmínek pro pouívání úvìru Air Bank a.s. - èl. Nìkolik bodù k ukonèení smlouvy, odst. 3 a " & _
'''            " rozhodnuto ze strany alobce o zesplatnìní úvìru  z dùvodu neplnìní smluvních podmínek alovanou stranou, viz. zesplatòující dopis - pøedalobní vızva. " ' ze dne " & Day(DataCheck.TextBox1.Text) & "." & Month(DataCheck.TextBox1.Text) & "." & Year(DataCheck.TextBox1.Text) & chr(10)
'''           dukaz_z_d = "Zesplatòující dopis - pøedalobní vızva" 'musime take upravit podobu dukazu s ohledem na platbu den po zesplatneni
'''          Else
'''          dukaz_z_d = "Zesplatòující dopis - pøedalobní vızva"
'''            text22 = "Vzhledem k závanému porušení smluvních povinností stranou alovanou bylo v souladu s Obchodními podmínkami a na základì Podmínek pro pouívání úvìru Air Bank a.s. - èl. Nìkolik bodù k ukonèení smlouvy, odst. 3 a - dne " & dz & _
'''            " rozhodnuto ze strany alobce o zesplatnìní úvìru  z dùvodu neplnìní smluvních podmínek alovanou stranou, viz. zesplatòující dopis - pøedalobní vızva ze dne " & dz & "." & Chr(10)
'''            text24 = "Ke dni zesplatnìní tedy byla dluná pohledávka tvoøena nesplacenou jistinou ve vıši " & format(jistina_zesp, "#,##0.00") & " Kè a nesplacenım smluvním úrokem z úvìru (kapitalizovanı úrok) ve vıši " & format(urok_zesp, "#,##0.00") & " Kè." 'toto se tam da, pokud byla platba po zesplatneni
'''
'''            text23 = "Ke dni zesplatnìní tedy byla dluná pohledávka tvoøena nesplacenou jistinou ve vıši " & format(jistina_zesp, "#,##0.00") & " Kè a nesplacenım smluvním úrokem z úvìru (kapitalizovanı úrok) ve vıši " & format(urok_zesp, "#,##0.00") & " Kè" & _
'''            ", kterı je souhrnem nesplacenıch smluvních úrokù ode dne poslední øádnì zaplacené splátky, tj. ode dne " & datum_posl_zapl_splatky & " do dne pøedcházejícího dni zesplatnìní, tedy do dne " & dz_1 & ", konkrétnì pak (poèet dnù prodlení/365 x aktuální jistina x " & rate_interest & "/100)." & Chr(10) 'pokud nebyla platba po zesplatneni
'''
'''        End If
'''    End If
'''rst1.Close
'''
'''Dim sql17 As String, text25 As String, sql13 As String, text222 As String
'''sql13 = "select * from " & schema & ".sp1 where skp_loan = " & skp_loan & " order by datum_platby, datum_splatky "
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql13, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''      If Not rst1.EOF Then
'''        rst1.MoveFirst
'''        For i = 1 To rst1.RecordCount
'''          If i = 1 Then
'''            text222 = "Po zesplatnìní úvìru uhradil alovanı dne " & rst1(2) & " èástku ve vıši " & Isinteger2(CDbl(rst1(13) + rst1(14))) & " Kè, kterou si alobce zapoèetl na úhradu dluné pohledávky, " & _
'''                "pøièem tuto èástku si alobce zapoèetl èástí ve vıši " & Isinteger2(CDbl(rst1(13))) & " Kè na úhradu jistiny splátky ze dne " & rst1(1) & " a èástí ve vıši " & Isinteger2(CDbl(rst1(14))) & " Kè na úhradu úroku splátky ze dne " & rst1(1) & "."
'''            Else
'''                text222 = text222 & Chr(10) & Chr(10) & "Dále po zesplatnìní úvìru uhradil alovanı dne " & rst1(2) & " èástku ve vıši " & Isinteger2(rst1(13) + rst1(14)) & " Kè, kterou si alobce zapoèetl na úhradu dluné pohledávky, " & _
'''                "pøièem tuto èástku si alobce zapoèetl èástí ve vıši " & Isinteger2(CDbl(rst1(13))) & " Kè na úhradu jistiny splátky ze dne " & rst1(1) & " a èástí ve vıši " & Isinteger2(CDbl(rst1(14))) & " Kè na úhradu úroku splátky ze dne " & rst1(1) & "."
'''          End If
'''        rst1.MoveNext
'''        Next i
'''    rst1.Close
'''        'zjistime info o aktualni pohledavce
'''
'''    sql17 = "select * from " & schema & ".balance_obs where skp_loan = " & skp_loan
'''    Set rst2 = New ADODB.Recordset
'''    rst2.Open sql17, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''    If Not rst2.EOF Then
'''        rst2.MoveFirst
'''        If IsNull(datum_posl_zapl_splatky_po_zes) Then
'''            text25 = "Aktuální vıše pohledávky je tvoøena nesplacenou jistinou ve vıši " & format(rst2(2), "#,##0.00") & " Kè a nesplacenım smluvním úrokem z úvìru do zesplatnìní (kapitalizovanı úrok) ve vıši " & format(rst2(3), "#,##0.00") & " Kè, kterı je souhrnem nesplacenıch smluvních úrokù ode dne naèerpání úvìru do dne pøedcházejícího dni zesplatnìní, tedy do dne " & dz_1 & _
'''            ", konkrétnì pak (poèet dnù prodlení/365 x aktuální jistina x " & rate_interest & "/100)."
'''            Else
'''            text25 = "Aktuální vıše pohledávky je tvoøena nesplacenou jistinou ve vıši " & format(rst2(2), "#,##0.00") & " Kè a nesplacenım smluvním úrokem z úvìru do zesplatnìní (kapitalizovanı úrok) ve vıši " & format(rst2(3), "#,##0.00") & " Kè, kterı je souhrnem nesplacenıch smluvních úrokù ode dne poslední zaplacené splátky, tj. ode dne " & datum_posl_zapl_splatky_po_zes & ", do dne pøedcházejícího dni zesplatnìní, tedy do dne " & dz_1 & _
'''            ", konkrétnì pak (poèet dnù prodlení/365 x aktuální jistina x " & rate_interest & "/100)."
'''        End If
'''    End If
'''    rst2.Close
'''
'''     text2 = text2 & Chr(10) & Chr(10) & text22 & Chr(10) & text24 & Chr(10) & Chr(10) & text222 & Chr(10) & Chr(10) & text25 & Chr(10) & Chr(10) & text21
'''
'''    Else 'nebyla platba po zesplatneni
'''    rst1.Close
'''     text2 = text2 & Chr(10) & Chr(10) & text21 & Chr(10) & Chr(10) & text22 & Chr(10) & Chr(10) & text23
'''    End If
'''
'''
'''text2 = text2 & Chr(10) & "Dùkaz: " & Chr(10) & _
'''"Splátkovı kalendáø úvìru è. " & hu & Chr(10) & _
'''"Pøehled plateb úvìru è. " & hu & Chr(10)
'''If text2_dukaz <> "" Then
'''    text2 = text2 & text2_dukaz & Chr(10)
'''End If
'''
'''If dukaz_dod <> "" Then
'''    text2 = text2 & dukaz_dod & Chr(10)
'''End If
'''
'''If dukaz_z_d <> "" Then
'''    text2 = text2 & dukaz_z_d & Chr(10)
'''End If
'''
'''text2 = text2 & "Mimoøádnı vıpis z úètu" & Chr(10) & Chr(10) & "III."
'''
'''
''''------------------------------------------------------------------------------------------------------------------------------------------------------
''''--------------CAST III -------------------------------------------------------------------------------------------------------------------------------
''''------------------------------------------------------------------------------------------------------------------------------------------------------
'''
'''Dim text3 As String, dz1 As String, t3 As Integer
'''
'''dz1 = Day(DataCheck.TextBox1.text) - 1 & "." & Month(DataCheck.TextBox1.text) & "." & Year(DataCheck.TextBox1.text)
'''text3 = "Zesplatòujícím dopisem - pøedalobní vızvou byl alovanı vyzván k zaplacení celého dluhu, kterı byl tvoøen dlunou jistinou ve " & _
'''"vıši " & format(jistina_zesp, "#,##0.00") & " Kè a nesplacenım smluvním úrokem z úvìru ve vıši " & format(urok_zesp, "#,##0.00") & " Kè, pøièem jejich vıše byla stanovena ke dni " & dz_1 & ". Zesplatnìním se stal závazek splatnım v plné vıši spolu s pøíslušenstvím ke dni " & _
'''dz & "." & Chr(10) & Chr(10) & "alobce alovanému v zesplatòujícím dopise - pøedalobní vızvì poskytl dodateènou " & _
'''" lhùtu k doplacení, alovanı ji však ke splacení nevyuil. Kdyby alovanı v této lhùtì svùj dluh zaplatil, nepoadoval by alobce " & _
'''"zaplacení smluvního úroku ani zákonného úroku z prodlení ode dne zesplatnìní do dne úhrady." & Chr(10) & Chr(10)
'''
'''text3 = text3 & "Smluvní strany si podpisem Rámcové smlouvy sjednaly, viz. Obchodní podmínky - èl. Doruèování dokumentù, e alobce doruèuje alovanému dokumenty " & _
'''    "na adresu alovaného. Dokumenty odeslané poštou se povaují  za " & _
'''    "doruèené desátım dnem ode dne jejich odeslání."
'''
'''
'''text3 = text3 & Chr(10) & Chr(10) & "Dùkaz: " & Chr(10) & _
'''dukaz_z_d & Chr(10) & _
'''"Podací arch " & Chr(10) & _
'''"Doklad odeslané pošty " & Chr(10)
'''
'''
'''text3 = text3 & "Obchodní podmínky Air Bank a.s." & Chr(10) & Chr(10) & "IV."
'''
''''"Doklad o odeslání/doruèení zesplatòujícího dopisu - pøedalobní vızvy ze dne " & don & Chr(10) & _
'''
'''
'''Dim text41 As String, text42 As String, sql15 As String, qq As Boolean, start_su As Date, sql16 As String
'''qq = False
'''sql15 = "select * from " & schema & ".sablona_epr11 where uver_cislo = '" & hu & "' order by datum_platby"
'''sql16 = "select min(start_su2) from " & schema & ".V_DATACHECK_CREDIT_I where uver_cislo = '" & hu & "' group by uver_cislo "
'''
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql16, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''      If Not rst1.EOF Then
'''        rst1.MoveFirst
'''        start_su = rst1(0)
'''    End If
'''rst1.Close
'''
'''
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql15, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''      If Not rst1.EOF Then
'''        rst1.MoveFirst
'''        dpp = rst1(1)
'''        For i = 1 To rst1.RecordCount
'''        text41 = text41 & " do " & rst1(1) & ", spolu se smluvním úrokem " & rate_interest & " % roènì z èástky " & format(rst1(2), "#,##0.00") & " Kè od " & rst1(0)
'''            If rst1(0) >= start_su Then
'''                If qq = False Then
'''                text42 = text42 & " do " & rst1(1) & " spolu se zákonnım úrokem z prodlení z èástky " & format(rst1(2), "#,##0.00") & " Kè od " & rst1(0)
'''                qq = True
'''                Else
'''                text45 = text45 & " do " & rst1(1) & ", spolu se zákonnım úrokem z prodlení z èástky " & format(rst1(2), "#,##0.00") & " Kè od " & rst1(0)
'''                End If
'''            End If
'''        rst1.MoveNext
'''        Next i
'''    End If
'''rst1.Close
'''
'''
'''
'''Dim sql155 As String, jistina_akt As Double, urok_akt As Double, celkem_akt As Double
'''
'''sql155 = "select * from " & schema & ".v_datacheck_credit_a where uver_cislo = '" & hu & "'"
'''Set rst1 = New ADODB.Recordset
'''   rst1.Open sql155, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''      If Not rst1.EOF Then
'''        jistina_akt = rst1(3)
'''        urok_akt = rst1(4)
'''        celkem_akt = jistina_akt + urok_akt
'''    End If
'''rst1.Close
'''
'''
''''------------------------------------------------------------------------------------------------------------------------------------------------------
''''--------------CAST IV -------------------------------------------------------------------------------------------------------------------------------
''''------------------------------------------------------------------------------------------------------------------------------------------------------
'''
'''text41 = "Celkovı dluh vyplıvající z úvìrové smlouvy è. " & hu & " ke dni podání aloby èiní na jistinì èástku ve vıši " & format(jistina_akt, "#,##0.00") & " Kè spolu se smluvním úrokem " & rate_interest & _
'''" % roènì z èástky " & format(jistina_zesp, "#,##0.00") & " Kè od " & dz & text41
'''text42 = " a dále zákonnı úrok z prodlení z èástky " & format(jistina_zesp, "#,##0.00") & " Kè od " & start_su & text42
'''
'''text41 = text41 & " do zaplacení"
'''text42 = text42 & text45 & " do zaplacení."
'''
'''text4 = text41 & text42 & " alobce poaduje zákonnı úrok z prodlení a od " & start_su & ", nebo dle interních pravidel poskytuje alobce alovanému ve vızvì k plnìní dodateènou lhùtu. " & _
'''"Dále alobce poaduje kapitalizovanı smluvní úrok ve vıši " & format(urok_akt, "#,##0.00") & " Kè za období od " & dz_1 & " do " & dz1 & "."
''''FA_bug
'''Dim text As String
'''
'''
'''
'''text = text1 & Chr(10) & dukaz1 & Chr(10) & text2 & Chr(10) & text3 & Chr(10) & text4
'''
'''   Dim objWord
'''
'''   Dim objDoc
'''
'''   Set objWord = CreateObject("Word.Application")
'''
'''   Set objDoc = objWord.Documents.Add
'''
'''   objWord.Visible = True
'''   Set objSelection = objWord.Selection
'''   objSelection.text = text
''''   objDoc.SaveAs ("D:\Users\pboril1374ab\Documents\al.docx")
''''Worksheets("Sablona").Cells(1, 1).Value = ""
''''Worksheets("Sablona").Cells(1, 1).Value = Text
''''
''''
''' AFolder = cfg.Cesta("EprPodaneZaloby") & c.name & " " & Klient.Label2.Caption
''' BFolder = cfg.Cesta("EprNachystane") & c.name & " " & Klient.Label2.Caption
''' CFolder = "W:\Vymahani\LATE\EPR\NACHYSTANÉ BØEZEN 2015\" & c.name & " " & Klient.Label2.Caption
'''
''' AAFolder = cfg.Cesta("EprPodaneZaloby") & c.name & " " & Klient.Label2.Caption & "\Pøílohy"
''' BBFolder = cfg.Cesta("EprNachystane") & c.name & " " & Klient.Label2.Caption & "\Pøílohy"
''' CCFolder = "W:\Vymahani\LATE\EPR\NACHYSTANÉ BØEZEN 2015\" & c.name & " " & Klient.Label2.Caption & "\Pøílohy"
'''
'''
'''Set fso = CreateObject("Scripting.FileSystemObject")
'''
'''If fso.FolderExists(AAFolder) Then
''''Range(Cells(1, 1), Cells(150, 9)).ExportAsFixedFormat Filename:=AAFolder & "\aloba úvìru è. " & hu & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".pdf", Type:=xlTypePDF, _
'''        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
'''        :=False, OpenAfterPublish:=True
'''   objDoc.SaveAs (AAFolder & "\aloba úvìru è. " & hu & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".docx")
'''
''' Call Shell("explorer.exe " & AAFolder, vbNormalFocus)
'''
'''ElseIf fso.FolderExists(AFolder) Then
''''Range(Cells(1, 1), Cells(150, 9)).ExportAsFixedFormat Filename:=AFolder & "\aloba úvìru è. " & hu & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".pdf", Type:=xlTypePDF, _
'''        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
'''        :=False, OpenAfterPublish:=True
'''   objDoc.SaveAs (AFolder & "\aloba úvìru è. " & hu & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".docx")
'''
''' Call Shell("explorer.exe " & AFolder, vbNormalFocus)
'''
'''ElseIf fso.FolderExists(BBFolder) Then
''''  Range(Cells(1, 1), Cells(150, 9)).ExportAsFixedFormat Filename:=BBFolder & "\aloba úvìru è. " & hu & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".pdf", Type:=xlTypePDF, _
'''        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
'''        :=False, OpenAfterPublish:=True
'''   objDoc.SaveAs (BBFolder & "\aloba úvìru è. " & hu & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".docx")
'''
'''  Call Shell("explorer.exe " & BBFolder, vbNormalFocus)
'''
'''ElseIf fso.FolderExists(BFolder) Then
''''  Range(Cells(1, 1), Cells(150, 9)).ExportAsFixedFormat Filename:=BFolder & "\aloba úvìru è. " & hu & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".pdf", Type:=xlTypePDF, _
'''        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
'''        :=False, OpenAfterPublish:=True
'''   objDoc.SaveAs (BFolder & "\aloba úvìru è. " & hu & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".docx")
'''
'''  Call Shell("explorer.exe " & BFolder, vbNormalFocus)
'''
'''  ElseIf fso.FolderExists(CCFolder) Then
''''Range(Cells(1, 1), Cells(150, 9)).ExportAsFixedFormat Filename:=CCFolder & "\aloba úvìru è. " & hu & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".pdf", Type:=xlTypePDF, _
'''        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
'''        :=False, OpenAfterPublish:=True
'''   objDoc.SaveAs (CCFolder & "\aloba úvìru è. " & hu & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".docx")
'''
'''  Call Shell("explorer.exe " & CCFolder, vbNormalFocus)
'''
'''  ElseIf fso.FolderExists(CFolder) Then
''''Range(Cells(1, 1), Cells(150, 9)).ExportAsFixedFormat Filename:=CFolder & "\aloba úvìru è. " & hu & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".pdf", Type:=xlTypePDF, _
'''        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
'''        :=False, OpenAfterPublish:=True
'''   objDoc.SaveAs (CFolder & "\aloba úvìru è. " & hu & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".docx")
'''
'''  Call Shell("explorer.exe " & CFolder, vbNormalFocus)
'''
'''Else: MsgBox "Sloka pro klienta " & c.name & " neexistuje." ', vbInformation, "Sloka nenalezena!"
''''Range(Cells(1, 1), Cells(150, 9)).ExportAsFixedFormat Filename:="aloba úvìru è." & hu & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".pdf", Type:=xlTypePDF, _
'''        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
'''        :=False, OpenAfterPublish:=True
'''   objDoc.SaveAs (cfg.Cesta("EprKlientiBezSlozky") & "aloba úvìru è." & hu & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".docx")
'''        Call Shell("explorer.exe " & cfg.Cesta("EprKlientiBezSlozky"), vbNormalFocus)
'''
'''End If
'''
'''
'''
'''Worksheets("Sablona").Visible = False
'''Exit Sub
'''connection_problem:
'''
'''If Mid(Err.Description, 1, 9) = "ORA-03114" Then
'''err_connection = True
'''Else
''' MsgBox Err.Description
'''End If
'''
'''End Sub
'''
'''
'''
