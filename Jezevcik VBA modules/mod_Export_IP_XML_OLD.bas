Attribute VB_Name = "mod_Export_IP_XML_OLD"
Option Explicit

Private ii As Integer, qq As Integer, qi As Integer, ij As Integer, iii As Integer, q As Integer, qe As Integer
Private SOUD As Variant
Private objSelection As Object


Sub export_insolvence_old(id_person As Variant)
'RISK-539 - export insolvence do XML bol nahradeny novou sablonou


Dim pm As Boolean, sql1 As String, sql2 As String, sql3 As String, sql9 As String, text As String
'zjistime, zda ma klient pravomocne rozhodnuti - ovlnivni to par veci v sabolne
   sql1 = "select 1 from " & schema & ".epr_status_data e join " & schema & ".epr_status es on e.skp_epr_status = es.skp_epr_status where name_epr_status like 'PM%' and id_person = " & id_person
    
Set rst1 = New ADODB.Recordset
   rst1.Open sql1, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   If Not rst1.EOF Then
     
     pm = True
    
    End If
rst1.Close


' bod c.1 - soud + sp. znacka

'Dim kr_soud As String, okr_soud As String
Dim sp As String


''''''' RISK-1098 sql nahrazeno za objekt
'vyber soudu
'''sql2 = "select okresni_soud, krajsky_soud from " & schema & ".person_soud where id_person = " & id_person
'''Set rst2 = New adodb.Recordset
'''   rst2.Open sql2, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''   If Not rst2.EOF Then
'''
'''     kr_soud = c.SOUD.krajsky_soud
'''     okr_soud = c.SOUD.okresni_soud
'''
'''    End If
'''rst2.Close

'spisova znacka
sql3 = "select code from " & schema & ".ins_info where id_person = " & id_person
Set rst3 = New ADODB.Recordset
   rst3.Open sql3, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   If Not rst3.EOF Then
   
     sp = rst3(0)
   
    
    End If
rst3.Close



'bod c.2

Dim rs As String, vu As String, dp As Date

'RISK-398
rs = c.TextContractNumber
dp = c.DatePayoff

Dim sql5 As String

'dalsi potrebne udaje, ktere jsou v specialni tabulce
'tahle tabulka se napocitava v air_boril.pb_037_sablona
sql5 = "select TEXT_BANK_ACCOUNT_NUMBER from " & schema & ".ins_sablona where id_person = " & id_person 'owner_dwh.f_bank_account_ad.SKP_BANK_ACCOUNT_type = 21270 - vnitøní úèet insolvence
Set rst5 = New ADODB.Recordset
   rst5.Open sql5, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   If Not rst5.EOF Then
   
        
        If IsNull(rst5![Text_bank_account_number]) Then
            vu = "..."
            Else
        
            vu = "" & rst5![Text_bank_account_number] & ""
        End If
    
    End If
rst5.Close

Dim skl() As Long, sql6 As String
Dim cnt_loan As Integer, i As Integer

'text = c.SOUD.krajsky_soud & ", " & sp & Chr(10)
text = text & c.Text_Title_before & c.name & c.Text_Title_After & Chr(10) & "RÈ: " & c.Pin & Chr(10) & "Datum narození: " & c.Date_Birth & Chr(10) & c.Address_C & Chr(10) & c.Nationality & Chr(10) & vu & "/3030" & Chr(10) & "VS: " & rs & Chr(10)

'Dim skl() As Integer

'bod c.3 info o pohledavce
'zjistime vsechny loany
sql6 = "select skp_loan from " & schema & ".balance_obs where id_person = " & id_person & " order by skp_loan"
Set rst6 = New ADODB.Recordset
   rst6.Open sql6, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   If Not rst6.EOF Then
            
        cnt_loan = rst6.RecordCount
        ReDim skl(cnt_loan)
        rst6.MoveFirst
        For i = 1 To cnt_loan
            skl(i) = rst6![skp_loan]
            rst6.MoveNext
        Next i
        
            
    End If
rst6.Close

Dim sql7 As String, sql8 As String, datum_podpisu As Date, datum_cerpani As Date, cL As String, dod As String, Amt_loan_limit As Double, ucet As String, code_epr As String
Dim rate_interest As Double
Dim Rate_Interest_Repo As Double


'sql8 = "select amt_loan_limit from " & schema & ".sp7 where code_loan_number  = '" & cl & "'"
'Set rst8 = New ADODB.Recordset
'   rst8.Open sql8, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'   If Not rst8.EOF Then
'        amt_loan_limit = rst8(0)
'    End If
'rst8.Close



Dim sql10 As String, cnt_int As Integer, amt_jis As Double, amt_kap_urok As Double, amt_uroky As Double, sql77 As String
Dim text_pohl As String, amt_jis_akt As Double, flag_rs As Boolean, flag_epr As Boolean, text_pohl1 As String
'flag_rs = False
'pro kazdy loan zjistime informace !!! až bude kompletní objekt LOAN a v nìm všechny potøebné infomrmace, mùže se upravit
For i = 1 To cnt_loan
    'podle dodatku zjistime, zda s ramcovou smlouvu nebo ne
    
   
    'nyni zjistime cislo uveru, dodatek, ucet
    sql7 = "select code_loan_number, text_appendix_number, text_bank_account_number, trunc(dtime_signature) as dtime_signature, trunc(date_loan_utilization) as date_loan_utilization from " & schema & ".sablona_epr where date_loan_utilization is not null and skp_loan = " & skl(i)
    Set rst7 = New ADODB.Recordset
       rst7.Open sql7, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
       If Not rst7.EOF Then
              cL = rst7![code_loan_number]
              dod = rst7![Text_appendix_number]
              ucet = "" & rst7![Text_bank_account_number] & "" 'RISK-398 oprava
              datum_cerpani = "" & rst7![Date_loan_utilization] & ""
              If IsNull(rst7![Dtime_signature]) Then 'tady bych dala asi spíš podmínku if rst7![text_appendix_number] = 0
                flag_rs = True
                Else
                flag_rs = False
                datum_podpisu = rst7![Dtime_signature]
              End If
        End If
    rst7.Close
    
    'pokud vznikla smlouva spolecne s ramcovou smlouvou, musime se podivat na datum zalozeni RS
    If flag_rs Then
    
        sql77 = "select trunc(dtime_signature) from " & schema & ".sablona_epr where code_completion_subject_type = 'RS' and dtime_signature is not null and id_person = " & id_person
        Set rst1 = New ADODB.Recordset
           rst1.Open sql77, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
           If Not rst1.EOF Then
                datum_podpisu = rst1(0)
            End If
        rst1.Close
    
    End If
    
    
    
    
'''    FA - nasledujici sql dotaz vyhozen v ramci RISK-726 do Jez.v.5.3, informace se nenahravi z tabulky SP7 ktera je zbytecna, ale z objektu Loan
'''    sql8 = "select amt_loan_limit, rate_interest_standard, rate_interest_repo from " & schema & ".sp7 where code_loan_number  = '" & cL & "'"
'''    Set rst8 = New ADODB.Recordset
'''       rst8.Open sql8, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'''       If Not rst8.EOF Then
            Amt_loan_limit = c.GetLoans(-1, cL).Amt_loan_limit ' rst8(0)
            rate_interest = c.GetLoans(-1, cL).Rate_Interest_Standard ' rst8(1)
            Rate_Interest_Repo = c.GetLoans(-1, cL).Rate_Interest_Repo ' rst8![Rate_Interest_Repo]
'''        End If
'''    rst8.Close
    ''''''''''''''''''''''''''
    ' Výše jistiny úvìru     '
    ''''''''''''''''''''''''''
    ' body 08-10 - pocitani pohledavky z jiz napocitane tabulky insolvenci
    
    
    sql10 = "select interest_overdue, amt_balance_principal, amt_interest_overdue, dtime_valid_from_ou , dtime_valid_to_ou -1,  dtime_valid_from_su, dtime_valid_to_su -1, ou, su from ap_coll.sr_coll_014_intervaly where skp_loan = " & skl(i) & " order by dtime_valid_from_ou" 'tabulka se plni na ap_coll hristi

    Set rst1 = New ADODB.Recordset
       rst1.Open sql10, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
       If Not rst1.EOF Then
               cnt_int = rst1.RecordCount
                rst1.MoveFirst
                'aktualni kap. urok
                text_pohl = "Nezaplacený smluvní úrok do zesplatnìní úvìru ve výši " & format(rst1(0), "#,##0.00") & " Kè"
                
                'zapamatujeme si jistinu a kap. urok v den zesplatneni
                amt_jis = rst1(1)
                amt_kap_urok = rst1![interest_overdue]
                
                'celkova vyse vymahanych uroku
                amt_uroky = rst1(0)
                For ii = 1 To cnt_int
                    
                    'obchodni urok
                    text_pohl = text_pohl & Chr(10) & "Smluvní úrok ve výši " & rate_interest & " % roènì z pohledávky " & format(rst1(1), "#,##0.00") & " Kè od " & rst1(3) & " do " & rst1(4) & " v celkové výši " & format(rst1![ou], "#,##0.00") & " Kè."
                    'zakonny urok z prodleni
                    text_pohl = text_pohl & Chr(10) & "Zákonný úrok z prodlení ve výši " & sngZakonnyUrok + Rate_Interest_Repo & " % roènì z pohledávky " & format(rst1(1), "#,##0.00") & " Kè od " & rst1(5) & " do " & rst1(6) & " v celkové výši " & format(rst1![su], "#,##0.00") & " Kè." ' RISK-726: bolo tu fixne 8,05

                    
                    amt_uroky = amt_uroky + rst1![ou] + rst1![su]
                    
                    amt_jis_akt = rst1(1) 'aktualni vyse pohledavky
                    rst1.MoveNext
                Next ii
        End If
    rst1.Close
    
    text = text & i & ". ÚVÌR" & Chr(10)
    text = text & "Výše jistiny: " & amt_jis_akt & " Kè" & Chr(10)
    
    If dod = 0 Then
        text = text & "Poskytnutí úvìru è. " & cL & " na èástku " & format(Amt_loan_limit, "#,##0") & " Kè ze dne " & datum_podpisu & " na základì podpisu Rámcové smlouvy è. " & _
        rs & "."
        Else
        text = text & "Poskytnutí úvìru è. " & cL & " na èástku " & format(Amt_loan_limit, "#,##0") & " Kè ze dne " & datum_podpisu & " na základì podpisu dodatku è. " & _
        dod & " k Rámcové smlouvì è. " & rs & "."
    End If
    
    
    sql9 = "  select  i.code from  " & schema & ".information i where id_person =" & id_person
    'TADY TO PADÁ PRO NÌKTERÉ KLIENTY!!!!!!!!!!!!!!!!!!!!!!!!!!
    Set rst1 = New ADODB.Recordset
       rst1.Open sql9, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
       If Not rst1.EOF Then
            code_epr = "" & rst1![code] & ""
        End If
    rst1.Close
    

    
    If pm Then 'ma pm titul
        
        text = text & " Vìøitel svùj nárok na zaplacení pohledávky uplatnil u " & c.SOUD.okresni_soud & ". Dlužníkovi byla "
        
        If Left(code_epr, 1) = "E" Or Left(code_epr, 1) = "e" Then
            text = text & "elektronickým platebním rozkazem è.j. "
            flag_epr = 1
            Else
            text = text & "rozsudkem è.j. "
        End If
        'rozsudkem è.j. " &
        text = text & code_epr & " ze dne ..., který nabyl právní moci dne ..., uložena povinnost vìøiteli" & _
        " pohledávku zaplatit."
        
    End If


    '07 - dalsi okolonosti

    text = text & Chr(10) & "Finanèní prostøedky ve výši " & format(Amt_loan_limit, "#,##0") & " Kè byly dne " & datum_cerpani & " naèerpány na bìžný úèet klienta - è. úètu: " & ucet & "/3030. Vzhledem k závažnému porušení smluvních povinností " & _
    "ze strany dlužníka bylo v souladu s Obchodními podmínkami a na základì Podmínek pro používání pùjèky Air Bank a.s. - str. 8, bod 3a - dne " & dp & " rozhodnuto ze strany Air Bank a.s. o zesplatnìní úvìru." & Chr(10)

    

    '''''''''''''''''''''''''
    '  blok 8-10 - odstranen'
    '''''''''''''''''''''''''
    
    If pm Then 'pokud pm, zahrneme 08 - vykonavatelnost
    'zakomentováno kuli text_pohl1
        text_pohl1 = Chr(10) & "Výše jistiny: " & format(amt_jis_akt, "#,##0.00") & " Kè" & Chr(10) & _
        "pro èástku: " & amt_jis_akt + amt_uroky & " Kè "

        If flag_epr Then
            text_pohl1 = text_pohl1 & "EPR "
            Else
            text_pohl1 = text_pohl1 & "rozsudku "
        End If
        'Rozsduku " & "

        text_pohl1 = text_pohl1 & c.SOUD.okresni_soud & " è.j., " & code_epr & " ze dne ..., který nabyl právní moci dne..." & Chr(10) & _
        "Smluvní úrok a úrok z prodlení: " & amt_uroky & Chr(10) & _
        text_pohl & Chr(10) & _
        "Celková èástka pohledávky: " & format(amt_jis_akt + amt_uroky, "#,##0.00") & Chr(10) & _
        "Splatná od: " & dp & " v èástce: " & amt_jis + amt_kap_urok

        Else

         text_pohl1 = "Smluvní úrok a úrok z prodlení: " & amt_uroky & Chr(10) & _
        text_pohl & Chr(10) & _
        "Celková èástka pohledávky: " & format(amt_jis_akt + amt_uroky, "#,##0.00") & Chr(10) & _
        "Splatná od: " & dp & " v èástce: " & amt_jis + amt_kap_urok

    End If
    text = text & text_pohl1 & Chr(10)
Next i


'_________________________________________________________________________________________________________
'text_pohl = Chr(10) & "Výše jistiny: " & Format(amt_jis_akt, "#,##0.00") & " Kè" & Chr(10) & _
'        "pro èástku: " & Format(amt_jis_akt + amt_uroky, "#,##0.00") & " Kè dle "
'
'        If flag_epr Then
'            text_pohl1 = text_pohl1 & "elektronického platebního rozkazu "
'            Else
'            text_pohl1 = text_pohl1 & "rozsudku "
'        End If
'        'Rozsduku " & "
'
'        text_pohl = text_pohl1 & c.SOUD.okresni_soud & " è.j., " & code_epr & " ze dne ..., který nabyl právní moci dne..." & Chr(10) & _
'        "Smluvní úrok a úrok z prodlení: " & Format(amt_uroky, "#,##0.00") & Chr(10) & _
'        text_pohl & Chr(10) & _
'        "Celková èástka pohledávky: " & Format(amt_jis_akt + amt_uroky, "#,##0.00") & Chr(10) & _
'        "Splatná od: " & dp & " v èástce: " & Format(amt_jis + amt_kap_urok, "#,##0.00")
'
'        Else
'
'         text_pohl = "Výše jistiny: " & Format(amt_jis_akt, "#,##0.00") & " Kè" & Chr(10) & _
'        "Smluvní úrok a úrok z prodlení: " & Format(amt_uroky, "#,##0.00") & Chr(10) & _
'        text_pohl & Chr(10) & _
'        "Celková èástka pohledávky: " & amt_jis_akt + amt_uroky & Chr(10) & _
'        "Splatná od: " & dp & " v èástce: " & Format(amt_jis + amt_kap_urok, "#,##0.00")
'
'    End If
'Next i
'
'text = text & text_pohl
'_______________________________________________________________________________________________________
If pm Then
'pokud je pm pohledavka, doplnime naklady rizeni - další pohledávka - SOP
    Dim sql11 As String, amt_sop As Double
    sql11 = "select naklady_rizeni from " & schema & ".v_sop_balance where id_person = " & id_person

    Set rst1 = New ADODB.Recordset
       rst1.Open sql11, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
       If Not rst1.EOF Then
            amt_sop = rst1(0)
        End If
    rst1.Close

    text = text & Chr(10) & Chr(10) & "Další pohledávka - SOP" & Chr(10) & _
    "Výše pohledávky: " & format(amt_sop, "#,##0.00") & Chr(10) & _
    "Dùvod vzniku: Nezaplacený soudní poplatek na základì "
    
        If flag_epr Then
            text = text & " elektronického platebního rozkazu "
            Else
            text = text & "rozsudku "
        End If
    
    text = text & SOUD & " è.j.: " & code_epr & " ze dne ..., který nabyl právní moci dne ..." & Chr(10) & _
    "Vykonavatelnost pro èástku: " & amt_sop
    
        If flag_epr Then
            text = text & " EPR "
            Else
            text = text & "rozsudku "
        End If
        
    text = text & c.SOUD.okresni_soud & " è.j.: " & code_epr & " ze dne ..., který nabyl právní moci dne ..." & Chr(10) & _
    "Splatná od ... v èástce: " & format(amt_sop, "#,##0.00")
    
    
End If

'Debug.Print text
'MsgBox "hotovo"
   Dim objWord

   Dim objDoc

   Set objWord = CreateObject("Word.Application")

   Set objDoc = objWord.Documents.Add

   objWord.Visible = True
   Set objSelection = objWord.Selection
   objSelection.text = text




Exit Sub
connection_problem:

If Mid(Err.Description, 1, 9) = "ORA-03114" Then
err_connection = True
Else
 MsgBox Err.Description
End If

End Sub





