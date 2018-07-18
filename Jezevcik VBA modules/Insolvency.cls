VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Insolvency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'objekt Insolvency vytvoreny ku sablone prihlasky pohledavky do insolvencniho rizeni - module IP_XML
'data nizsie sa nacitaju ku konkretnemu klientovi

'vsetky premenne zacinaju skratkou pre datovy typ pouzity, je to good practice. Tabulka tu: https://msdn.microsoft.com/en-us/library/aa240858(v=vs.60).aspx
Private pstrInsCode As String
Private pstrKodInsSoud As String
'Private pstrOkresSoud As String  ' v celem modulu odstraneno, RISK-1098
'Private pstrKrajSoud As String   ' v celem modulu odstraneno, RISK-1098
Private pblnPm As Boolean
Private pblnPmZamitnut As Boolean
Private pstrVu As String
Private pstrSpisZnacka As String
Private pdblAmtSop As Double
Private pdblAmtSopPaid As Double
'Private psngNakladyRizeni As Single
Private pstrDatumUpadku As String
Private pstrDatumDenPredUpadkem As String
Private pstrDatumRamcovka As String
Private pstrDatumPosledniDebetTrn As String
Private pdtmDatumZesplatneni As Date 'asi nebude treba

'premenne potrebne iba vo vnutri modulu
Private strChybiUdaj As String, strDoplnUdaj As String ' 'chybove hlasky vypiseme naraz v jednom msgboxe na konci


Public Property Get strInsCode() As String
    strInsCode = pstrInsCode
End Property

Public Property Get strKodInsSoud() As String
    strKodInsSoud = pstrKodInsSoud
End Property

'''Public Property Get strOkresSoud() As String
'''    strOkresSoud = pstrOkresSoud
'''End Property
'''
'''Public Property Get strKrajSoud() As String
'''    strKrajSoud = pstrKrajSoud
'''End Property

Public Property Get blnPm() As Boolean
    blnPm = pblnPm
End Property

Public Property Get blnPmZamitnut() As Boolean
    blnPmZamitnut = pblnPmZamitnut
End Property

Public Property Get strVu() As String
    strVu = pstrVu
End Property

Public Property Get strSpisZnacka() As String
    strSpisZnacka = pstrSpisZnacka
End Property

Public Property Get dblAmtSop() As Double
    dblAmtSop = pdblAmtSop
End Property

Public Property Get dblAmtSopPaid() As Double
    dblAmtSopPaid = pdblAmtSopPaid
End Property

'Public Property Get sngNakladyRizeni() As Single
'    sngNakladyRizeni = psngNakladyRizeni
'End Property

Public Property Get strDatumUpadku() As String
    strDatumUpadku = pstrDatumUpadku
End Property

Public Property Get strDatumDenPredUpadkem() As String
    strDatumDenPredUpadkem = pstrDatumDenPredUpadkem
End Property

Public Property Get strDatumRamcovka() As String
    strDatumRamcovka = pstrDatumRamcovka
End Property

Public Property Get dtmDatumZesplatneni() As Date
    dtmDatumZesplatneni = pdtmDatumZesplatneni
End Property

Public Property Get dtmDatumDenPredZesplatnenim() As Date
    dtmDatumDenPredZesplatnenim = DateAdd("d", -1, pdtmDatumZesplatneni)
End Property




Public Sub initialize(id_person As Variant) 'XX  nejaky jednoznacny identifikator k INSOL (person asi ani netreba) napr. skp_insolvence abo tak

Dim sql_main As String, rst_main As Variant, sql_pm As String, rst_pm As Variant, sql_cislo_uctu As String, rst_cislo_uctu As Variant
'Dim strStavExtra As String
'strStavExtra = "Výpis ze seznamu pohledávek"
Dim sql_datum_upadku As String, rst_datum_upadku As Variant
Dim sql_ramcovka As String, rst_ramcovka As Variant


If cs_HDWMN.State = 1 Then

prihlasen:
    
    

    
    
    ''''   Datum upadku
    'nahrame z Jezevcika (vkladaju ho vymahaci)
    sql_datum_upadku = "SELECT i.id_person, trunc(i.datum_upadku) as datum_upadku, trunc(i.datum_upadku - 1) as den_pred_upadkem FROM " & schema & ".INS_INFO i where i.id_person = " & id_person

    Set rst_datum_upadku = New ADODB.Recordset
    rst_datum_upadku.Open sql_datum_upadku, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

        If Not rst_datum_upadku.EOF Then
            If Not IsNull(rst_datum_upadku![datum_upadku]) Then

                pstrDatumUpadku = "" & rst_datum_upadku![datum_upadku] & ""
                pstrDatumDenPredUpadkem = "" & rst_datum_upadku![den_pred_upadkem] & ""

            Else
                HlaskaUzivateli "Chybí datum úpadku. Je potøeba ho nejdøíve nahrát do Jezevèíka na kartu klienta, pak spusti šablonu znovu."
                pstrDatumUpadku = "___DOPLN DATUM UPADKU___"
                pstrDatumDenPredUpadkem = "___DOPLN DATUM UPADKU___"
            End If
        Else
            HlaskaUzivateli "Chybí datum úpadku. Je potøeba ho nejdøíve nahrát do Jezevèíka na kartu klienta, pak spusti šablonu znovu."
            pstrDatumUpadku = "___DOPLN DATUM UPADKU___"
            pstrDatumDenPredUpadkem = "___DOPLN DATUM UPADKU___"
        End If
    rst_datum_upadku.Close
    
    
    
    
    '''''''  nejdrive zjistime, zda ma klient pravomocne rozhodnuti - ovlivni to par veci v sablone
    sql_pm = "select name_epr_status from " & schema & ".epr_status_data e join " & schema & ".epr_status es on e.skp_epr_status = es.skp_epr_status where name_epr_status like 'PM%' AND e.flag_deleted = 'N' and id_person = " & id_person
    Set rst_pm = New ADODB.Recordset
    rst_pm.Open sql_pm, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        'rst_pm.MoveFirst 'XX ak prazdny recordset tak sposobi error a spadne to cele, pozor na to
        
        If Not rst_pm.EOF Then
            
            Select Case rst_pm![name_epr_status]
                
                Case "Výpis ze seznamu pohledávek"
                    HlaskaUzivateli "Máme výpis ze seznamu pohledávek, pøihlášku vyplò ruènì."
                    
                Case "PM øízení zastaveno - s doložkou"
                    'XX sem nejaku podmienku na vysku SOPu vraj treba..
                    'alebo boolean premennu ze ide o tento stav a tam kde sa nacita sop tak IF sop>0 and booleanXY then blnPm = true else blnPm = false end if
                    
                Case Else
                    pblnPm = True
                    
            End Select
            
        Else
            pblnPm = False
        End If
    rst_pm.Close
    
    
    
    '''''''  Zjistime, jestli klient mal jeden ze stavu "PM..zamitnut..". Podle toho se pak generuji intervaly k uverum.
    Dim sql_zamit As String
    Dim rst_zamit As ADODB.Recordset
    
    sql_zamit = "select count(name_epr_status) AS pocet_stavov from " & schema & ".epr_status_data e join " & schema & ".epr_status es on e.skp_epr_status = es.skp_epr_status WHERE lower(name_epr_status) LIKE '%pm%zamítnut%' AND e.flag_deleted = 'N' and id_person = " & id_person
    
    pblnPmZamitnut = False
    Set rst_zamit = New ADODB.Recordset
    rst_zamit.Open sql_zamit, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        If Not rst_zamit.EOF Then
            If rst_zamit![pocet_stavov] >= 1 Then
                pblnPmZamitnut = True
            End If
        End If
    rst_zamit.Close
    
    
    '''''''  Cislo vnitrniho uctu klienta pro insolvence
    sql_cislo_uctu = "SELECT a.cislo_uctu FROM " & schema & ".ucet a WHERE typ_uctu = 'IÚ' AND id_person = " & id_person
    Set rst_cislo_uctu = New ADODB.Recordset
    rst_cislo_uctu.Open sql_cislo_uctu, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
        'ak nahodou viacero uctov ku klientovi, vyhodi to hlasku uzivatelovi a nasledne sa vlozi prvy z nich(v podstate nahodne)
        If rst_cislo_uctu.RecordCount > 1 Then
            HlaskaUzivateli "Ke klientovi je v Jezevèíkovi více vnitøních (šestkových) úètù pro insolvence. Šablona se vyplní prvním z nich, nutno zkontrolovat!"
            rst_cislo_uctu.MoveFirst
        End If
        
        If Not rst_cislo_uctu.EOF Then
            If Len(Trim(rst_cislo_uctu![cislo_uctu])) > 0 Then
                pstrVu = "" & rst_cislo_uctu![cislo_uctu] & ""
            Else
                pstrVu = cfg.cst_ucet_522
            End If
        Else
            pstrVu = cfg.cst_ucet_522
        End If
    rst_cislo_uctu.Close
    
    
    
    '''''''  Vacsina ostatnych info
    
'''    ''' sql dotaz upraven (vytvoreno view v_insolvency_object) v ramci RISK-1098
'''    sql_main = "SELECT DISTINCT P.ID_PERSON, sd.okresni_soud, sd.krajsky_soud, IPS.KOD_IP_SOUD, info.code, nvl(bal.priznan_sop, 0) priznan_sop, bal.sop_paid, ins_info.code ins_code, nvl(p.date_payoff, '') AS datum_zesp " & _
'''                "FROM " & schema & ".PERSON P " & _
'''                "LEFT JOIN " & schema & ".person_soud PS ON PS.ID_PERSON = P.ID_PERSON " & _
'''                "LEFT JOIN " & schema & ".soudy sd ON sd.SKP_SOUD = PS.SKP_SOUD " & _
'''                "LEFT JOIN " & schema & ".ip_soudy ips ON UPPER(ips.nazev_ip_soud) = UPPER(sd.krajsky_soud) OR (sd.krajsky_soud = 'Krajský souv v Praze' AND ips.nazev_ip_soud = 'Krajský soud v Praze') " & _
'''                "LEFT JOIN " & schema & ".v_sop_balance bal ON bal.id_person = p.id_person " & _
'''                "LEFT JOIN " & schema & ".information info ON info.id_person = p.id_person " & _
'''                "LEFT JOIN " & schema & ".ins_info ins_info ON ins_info.id_person = p.id_person " & _
'''                "WHERE  p.flag_valid = 'Y' AND ps.flag_valid = 'Y' AND p.id_person = " & id_person

    sql_main = "SELECT a.ID_PERSON, a.OKRESNI_SOUD, a.KRAJSKY_SOUD, a.KOD_IP_SOUD, a.CODE, a.PRIZNAN_SOP, a.SOP_PAID, a.INS_CODE, a.DATUM_ZESP " & _
               " FROM " & schema & ".v_insolvency_object a WHERE a.ID_PERSON = " & id_person
    
    Set rst_main = New ADODB.Recordset
    rst_main.Open sql_main, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
        If rst_main.RecordCount > 1 Then
            HlaskaUzivateli "Ke klientovi existuje více záznamù a jezevèík je z toho úplnì hotovej, potøebuješ psovoda."
        ElseIf rst_main.RecordCount < 1 Then
            HlaskaUzivateli "Ke klientovi neexistuje insolvenèní záznam. jez_err 6655689. Napíš analytikom." 'XX jez_err?? :D
        Else
            rst_main.MoveFirst
            pstrInsCode = "" & rst_main![ins_code] & ""
            pstrKodInsSoud = "" & rst_main![KOD_IP_SOUD] & ""
            'pstrOkresSoud = "" & rst_main![okresni_soud] & ""
            'pstrKrajSoud = "" & rst_main![krajsky_soud] & ""
            pdtmDatumZesplatneni = "" & rst_main![datum_zesp] & ""
            
            pstrSpisZnacka = "" & rst_main![code] & ""
            'pdblAmtSop = "" & rst_main![priznan_sop] & ""
            pdblAmtSop = rst_main![priznan_sop]
            pdblAmtSopPaid = "" & rst_main![sop_paid] & ""
            
        End If
    rst_main.Close
    
    
    
    ' Datum podpisu ramcove smlouvy
    
    sql_ramcovka = "select trunc(dtime_signature) as dtime_signature from " & schema & ".sablona_epr where id_person = " & id_person & "  and code_completion_subject_type = 'RS' AND TEXT_CONTRACT_NUMBER = " & c.TextContractNumber
    Set rst_ramcovka = New ADODB.Recordset
    rst_ramcovka.Open sql_ramcovka, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
        If Not rst_ramcovka.EOF Then
            pstrDatumRamcovka = "" & rst_ramcovka![Dtime_signature] & ""
        Else
            HlaskaUzivateli "Nepovedlo se dotáhnout datum podpisu rámcové smlouvy."
            pstrDatumRamcovka = "__DATUM__"
        End If
        
    rst_ramcovka.Close
    
    

'NAVOD KU KONTROLE CI BOLI HODNOTY ZADEFINOVANE
'| Type                                 | Test                            | Test2
'| Numeric (Long, Integer, Double etc.) | If obj.Property = 0 Then        |
'| Boolen (True/False)                  | If Not obj.Property Then        | If obj.Property = False Then
'| Object                               | If obj.Property Is Nothing Then |
'| String                               | If obj.Property = "" Then       | If LenB(obj.Property) = 0 Then
'| Variant                              | If obj.Property = Empty Then    |

    'pokud chybi nektery udaj, vyhodime hlasku uzivateli
    strChybiUdaj = "> Ke klientovi chybí "
    strDoplnUdaj = ", bude se muset doplnit manuálnì."
    If blnPm = True Then
        HlaskaUzivateli "Ke klientovi máme PM rozhodnutí"
        Call ZkontrolujHodnotuInt(Me.dblAmtSop, "výška SOPu")
    End If
    Call ZkontrolujHodnotuStr(Me.strInsCode, "kód insolvence")
    Call ZkontrolujHodnotuStr(Me.strKodInsSoud, "kód insolvenèního soudu")
'''    Call ZkontrolujHodnotuStr(Me.strOkresSoud, "okresní soud")
'''    Call ZkontrolujHodnotuStr(Me.strKrajSoud, "krajský soud")
    'Call ZkontrolujHodnotuInt(Me.sngNakladyRizeni, "výška nákladù øízení")
    If Me.strVu = cfg.cst_ucet_522 Then
        HlaskaUzivateli strChybiUdaj & "èíslo vnitøního insolvenèního úètu. Dle Markétina pøání byl doplnìn úèet " & cfg.cst_ucet_522 & ". Pokud chceš jiný" & strDoplnUdaj
    Else
        Call ZkontrolujHodnotuStr(Me.strVu, "èíslo vnitøního insolvenèního úètu")
    End If
    
    
    If err_connection Then
        err_connection = False
        Call re_login(Main.username.text, Main.password.text)
        If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
    End If
        
Else

    Call re_login(Main.username.text, Main.password.text)
    If cs_HDWMN.State = 1 Then
        GoTo prihlasen
    Else
        Call logout
    End If
End If

End Sub

Private Sub ZkontrolujHodnotuStr(ByVal var As String, ByVal varName As String)

If var = "" Then
    HlaskaUzivateli strChybiUdaj & varName & strDoplnUdaj
End If

End Sub

Private Sub ZkontrolujHodnotuInt(ByVal var As Integer, ByVal varName As String)

If var = 0 Then
    HlaskaUzivateli strChybiUdaj & varName & strDoplnUdaj
End If

End Sub

Public Function ZjistiDebetPredUpadkem(id_person As Long, skp_bank_account As Long) As Double
'samostatna funkcia, protoze neni na urovni klienta ale klienta A UCTU
'tam kde sa vola fcia daj podmienku ze musi byt datum upadku! (Ins.strDatumDenPredUpadkem)

Dim sql_zustatek_pred_zespl As String, rst_zustatek As Variant

sql_zustatek_pred_zespl = "SELECT ip.amt_balance_pred_upadkem FROM " & schema & _
                          ".ip_vhodni_klienti_help ip WHERE ip.id_person = " & id_person & " AND ip.skp_bank_account = " & skp_bank_account

Set rst_zustatek = New ADODB.Recordset
rst_zustatek.Open sql_zustatek_pred_zespl, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

    If Not rst_zustatek.EOF Then
        
        If Not IsNull(rst_zustatek![amt_balance_pred_upadkem]) Then
            ZjistiDebetPredUpadkem = (-1) * rst_zustatek![amt_balance_pred_upadkem] 'vyska debetu opacne znamenko nez balans na uctu
        Else
            HlaskaUzivateli "Nepovedlo se naèíst výšku debetu ke dni pøed úpadkem."
        End If
'    Else
'        HlaskaUzivateli "Nepovedlo se naèíst výšku debetu. eh."
'        'pro jine nez klientske ucty (BU, SU, JPK) nejsou zaznamy v f_account_balance_pd  - ani je nepotrebujeme
    End If
rst_zustatek.Close

End Function

