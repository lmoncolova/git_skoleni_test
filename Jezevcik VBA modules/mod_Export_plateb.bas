Attribute VB_Name = "mod_Export_plateb"
Option Explicit
' FA: 20.12.2017 sem byla vytahnuta Sub export_seznam_plateb z MainModule
' ma 600 radku, tak aby se v tom lepe pracovalo
Private lastrow As Variant

Private Const dblRadek_prvni_uplne As Double = 15 ' K1
Private dblRadek_posl_platba_pred_zesp As Double  ' K2
Private dblRadek_prvni_po_zesp As Double          ' K3
Private dblRadek_uprostred_zaloby As Double       ' K4
Private dblRadek_posledni_pred_zalobou As Double  ' K5
Private dblRadek_prvni_po_zalobe As Double        ' K6
Private dblRadek_posledni_po_zalobe As Double     ' K66
Private dblRadek_posledni_uplne As Double         ' K7

Private rst71 As ADODB.Recordset
Private name_person As String
Private r As Long
Private S As Long
Private nes_jis_zal As Variant
Private kap_ur_zal As Variant
Private kap_ur_akt As Double
Private lngPos As Variant



' specialni query na seznam plateb
Sub export_seznam_plateb(sql As String, sql2 As String, sql3 As String, sql4 As String, sql5 As String, sql6 As String, sql8 As String, _
                         hu As String, Optional ByVal volba_exportu As Integer = 1, Optional ByVal zobrazit_slozku_klienta As Boolean = True) 'FA_export posledne 2 premenne v ramci risk-710
Dim max As Double, nes_jis As Double, kap_ur As Double, nes_jis_akt As Double, lr As Double, sql7 As String
Dim AFolder As String
Dim BFolder As String
Dim CFolder As String
Dim AAFolder As String
Dim BBFolder As String
Dim CCFolder As String
Dim currFolder As String ' current Folder
Dim wb_jez_name As String
Dim blnSlozkaNeexistujeShown As Boolean
Dim file_name As String
Dim str_klient_folder As String
Dim dz As String
Dim ws_Splatky As Worksheet

Dim fso


On Error GoTo connection_problem

 
' ********************************************************************** Vygeneruje nove pdf se splatkacema klientu

'Application.EnableEvent = False
'Workbook("Žaloby v.2.3 (Jezevèík) ").Activate

'name = ThisWorkbook.name
'MsgBox ThisWorkbook.name
cfg.wb_Jez.Activate

Set ws_Splatky = cfg.wb_Jez.Worksheets("Splatky")
ws_Splatky.Visible = True


ws_Splatky.Activate
lastrow = ws_Splatky.Cells(Rows.count, "C").End(xlUp).Row
ws_Splatky.Range("B15:M" & lastrow + 5).MergeCells = True
ws_Splatky.Range("B15:M" & lastrow + 5).Clear
   
   'DataCheck.ListView3.ListItems.Clear
   Set rst1 = New ADODB.Recordset
   'asss = InputBox("", "", sql)
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    
     
    rst1.MoveFirst

    'platby ke splatkam pred zesplatnenim
    
    For i = 1 To rst1.RecordCount
    ws_Splatky.Activate
    Cells(dblRadek_prvni_uplne + i, 3).value = "" & rst1(25) & "" 'datum splatky
    Cells(dblRadek_prvni_uplne + i, 4).value = rst1(22)
    Cells(dblRadek_prvni_uplne + i, 5).value = rst1(23)
    Cells(dblRadek_prvni_uplne + i, 6).value = rst1(24)
    Cells(dblRadek_prvni_uplne + i, 8).value = "" & rst1(2) & "" 'datum uhrady
    Cells(dblRadek_prvni_uplne + i, 10).value = rst1(10)
    Cells(dblRadek_prvni_uplne + i, 11).value = rst1(11)
    Cells(dblRadek_prvni_uplne + i, 13).value = rst1(15)
    rst1.MoveNext
    Next i
   End If
   dblRadek_posl_platba_pred_zesp = rst1.RecordCount + dblRadek_prvni_uplne
    rst1.Close

   Set rst2 = New ADODB.Recordset
   'nyni splatky, ktere byly pred zesplatnenim, ale nejsou v prvni tabulce (ani koruna neprisla)(sp3)
   rst2.Open sql2, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst2.EOF Then
    rst2.MoveFirst
    
    For i = 1 To rst2.RecordCount
    Cells(dblRadek_posl_platba_pred_zesp + i, 3).value = "" & rst2(0) & ""
    Cells(dblRadek_posl_platba_pred_zesp + i, 4).value = rst2(1)
    Cells(dblRadek_posl_platba_pred_zesp + i, 5).value = rst2(2)
    Cells(dblRadek_posl_platba_pred_zesp + i, 6).value = rst2(3)
    rst2.MoveNext
    Next i
   End If
   dblRadek_prvni_po_zesp = rst2.RecordCount + dblRadek_posl_platba_pred_zesp + 2
    
    rst2.Close
    
   Set rst3 = New ADODB.Recordset
   'splatky, ktere prisly po zesplatnenim (sp1)
   rst3.Open sql3, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
    If Not rst3.EOF Then
        rst3.MoveFirst
        
        For i = 1 To rst3.RecordCount
            Cells(dblRadek_prvni_po_zesp + i, 3).value = "" & rst3(23) & ""
            Cells(dblRadek_prvni_po_zesp + i, 5).value = rst3(27)
            Cells(dblRadek_prvni_po_zesp + i, 6).value = rst3(28)
            Cells(dblRadek_prvni_po_zesp + i, 4).value = rst3(29)
            Cells(dblRadek_prvni_po_zesp + i, 8).value = "" & rst3(30) & ""
            Cells(dblRadek_prvni_po_zesp + i, 10).value = rst3(31)
            Cells(dblRadek_prvni_po_zesp + i, 11).value = rst3(32)
            Cells(dblRadek_prvni_po_zesp + i, 13).value = rst3(33)
            rst3.MoveNext
        Next i
    End If
   dblRadek_uprostred_zaloby = rst3.RecordCount + dblRadek_prvni_po_zesp
    rst3.Close

 Set rst4 = New ADODB.Recordset
   'splatky, u kterych platby neprisly po zesplatneni (pro vypocet zespl. castky) (sp4)
    rst4.Open sql4, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
        If Not rst4.EOF Then
            rst4.MoveFirst
            
            For i = 1 To rst4.RecordCount
                Cells(dblRadek_uprostred_zaloby + i, 3).value = "" & rst4(1) & ""
                Cells(dblRadek_uprostred_zaloby + i, 4).value = rst4(2)
                Cells(dblRadek_uprostred_zaloby + i, 5).value = rst4(3)
                Cells(dblRadek_uprostred_zaloby + i, 6).value = rst4(4)
                rst4.MoveNext
            Next i
        End If
        dblRadek_posledni_pred_zalobou = dblRadek_uprostred_zaloby + rst4.RecordCount
    rst4.Close

dblRadek_prvni_po_zalobe = Application.WorksheetFunction.max(dblRadek_uprostred_zaloby, dblRadek_posledni_pred_zalobou) + 2

Set rst5 = New ADODB.Recordset
   'splatky, ktere prisly po zalobe (sp5)
   rst5.Open sql5, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst5.EOF Then
    rst5.MoveFirst
    
    For i = 1 To rst5.RecordCount
    Cells(dblRadek_prvni_po_zalobe + i, 3).value = "" & rst5(23) & ""
    Cells(dblRadek_prvni_po_zalobe + i, 5).value = rst5(27)
    Cells(dblRadek_prvni_po_zalobe + i, 6).value = rst5(28)
    Cells(dblRadek_prvni_po_zalobe + i, 4).value = rst5(29)
    Cells(dblRadek_prvni_po_zalobe + i, 8).value = "" & rst5(12) & ""
    Cells(dblRadek_prvni_po_zalobe + i, 10).value = rst5(13)
    Cells(dblRadek_prvni_po_zalobe + i, 11).value = rst5(14)
    Cells(dblRadek_prvni_po_zalobe + i, 13).value = rst5(16)
    rst5.MoveNext
    Next i
   End If
   dblRadek_posledni_po_zalobe = rst5.RecordCount + dblRadek_prvni_po_zalobe
    rst5.Close

Set rst8 = New ADODB.Recordset
   'splatky, u kterych platby neprisly po zesplatneni (ani po zalobe), pro vypocet, kolik dluzil pri zalobe
   rst8.Open sql8, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst8.EOF Then
    rst8.MoveFirst
    
    For i = 1 To rst8.RecordCount
    Cells(dblRadek_posledni_po_zalobe + i, 3).value = "" & rst8(1) & ""
    Cells(dblRadek_posledni_po_zalobe + i, 4).value = rst8(2)
    Cells(dblRadek_posledni_po_zalobe + i, 5).value = rst8(3)
    Cells(dblRadek_posledni_po_zalobe + i, 6).value = rst8(4)
    rst8.MoveNext
    Next i
   End If
   dblRadek_posledni_uplne = dblRadek_posledni_po_zalobe + rst8.RecordCount
    rst8.Close


Set rst6 = New ADODB.Recordset
'zjistime, kdy doslo k podani zaloby (sp6)
   rst6.Open sql6, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst6.EOF Then
    rst6.MoveFirst
    
    For i = 1 To rst6.RecordCount
    dz = "" & rst6(1) & ""
    rst6.MoveNext
    Next i
   End If
    rst6.Close

Dim sql71 As String
Set rst71 = New ADODB.Recordset
'zjistime jmeno klienta
sql71 = "select id_person from " & schema & ".sp2 where uver_cislo = '" & hu & "'"
rst71.Open sql71, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    If Not rst71.EOF Then
        id_person = rst71(0)
    End If
rst71.Close


'RISK-398
'zjistime jmeno klienta
 name_person = c.name

'dodatecne informace


'dodatecne informace



'celkova castka na zesplatneni-------------------------------------------------------------


Cells(dblRadek_prvni_po_zesp - 1, 3).value = "Zùstatek ke dni zesplatnìní, který se skládá z:"
Range(Cells(dblRadek_prvni_po_zesp - 1, 3), Cells(dblRadek_prvni_po_zesp - 1, 6)).MergeCells = True
Cells(dblRadek_prvni_po_zesp - 1, 8).value = "Platby po zesplatnìní"
Range(Cells(dblRadek_prvni_po_zesp - 1, 8), Cells(dblRadek_prvni_po_zesp - 1, 11)).MergeCells = True
Cells(dblRadek_prvni_po_zesp, 3).value = DataCheck.TextBox1.text
Cells(dblRadek_prvni_po_zesp, 4).value = Application.sum(Range(Cells(dblRadek_prvni_po_zesp + 1, 4), Cells(dblRadek_posledni_pred_zalobou, 4)))
nes_jis = Application.sum(Range(Cells(dblRadek_prvni_po_zesp + 1, 5), Cells(dblRadek_posledni_pred_zalobou, 5)))
kap_ur = Application.sum(Range(Cells(dblRadek_prvni_po_zesp + 1, 6), Cells(dblRadek_posledni_pred_zalobou, 6)))
Cells(dblRadek_prvni_po_zesp, 5).value = nes_jis
Cells(dblRadek_prvni_po_zesp, 6).value = kap_ur



' hlavicka -------------------------------------------------------------------------------
r = 3
S = 8
Cells(r + 1, 3).value = "Èíslo úvìru:  " & hu 'ListView1.SelectedItem
Cells(r, 3).value = "Klient:  " & name_person 'Klient.Label3.Caption
' najdu datum nacerpani a nacerpanou pujcku



    Cells(r + 4, 3).value = "Datum naèerpání pùjèky:  " & c.GetLoans(-1, hu).Date_loan_utilization
    Cells(r, S).value = "Naèerpaná jistina:  " & format(c.GetLoans(-1, hu).Amt_loan_limit, "#,##0.00") & " CZK"
    Cells(dblRadek_prvni_uplne, 13).value = c.GetLoans(-1, hu).Amt_loan_limit
    Cells(r + 2, 3).value = "Úroková sazba:  " & format(c.GetLoans(-1, hu).Rate_Interest_Standard, "#,##0.00") & " % p. a."
    Cells(r + 3, 3).value = "RPSN:  " & format(c.GetLoans(-1, hu).Rate_eir_standard, "#,##0.00") & " % p. a."



Cells(r + 1, S) = "Celková nesplacená jistina ke dni zesplatnìní:  " & format(nes_jis, "#,##0.00") & " CZK"
Cells(r + 4, S) = "Kapitalizovaný úrok ke dni zesplatnìní:  " & format(kap_ur, "#,##0.00") & " CZK"
nes_jis_zal = Application.WorksheetFunction.Min(Range(Cells(dblRadek_prvni_po_zesp + 1, 13), Cells(dblRadek_posledni_pred_zalobou, 13)), nes_jis)
kap_ur_zal = kap_ur - Application.sum(Range(Cells(dblRadek_prvni_po_zesp + 1, 11), Cells(dblRadek_posledni_pred_zalobou, 11)))
Cells(r + 2, S) = "Celková nesplacená jistina ke dni podání žaloby:  " & format(nes_jis_zal, "#,##0.00") & " CZK"
Cells(r + 5, S) = "Kapitalizovaný úrok ke dni podání žaloby:  " & format(kap_ur_zal, "#,##0.00") & " CZK"
lr = Application.WorksheetFunction.max(dblRadek_uprostred_zaloby, dblRadek_posledni_pred_zalobou)
nes_jis_akt = nes_jis_zal - Application.sum(Range(Cells(lr + 2, 10), Cells(dblRadek_posledni_uplne, 10)))
kap_ur_akt = kap_ur_zal - Application.sum(Range(Cells(lr + 2, 11), Cells(dblRadek_posledni_uplne, 11)))
Cells(r + 3, S) = "Celková nesplacená jistina ke dni " & date & ":  " & format(nes_jis_akt, "#,##0.00") & " CZK"
Cells(r + 6, S) = "Kapitalizovaný úrok ke dni " & date & ":  " & format(kap_ur_akt, "#,##0.00") & " CZK"




Range("B3:I12").Font.Bold = False
'tucne zvyraznene casti textu
lngPos = InStr(Cells(r, 3).value, ":")
With Cells(r, 3).Characters(Start:=1, Length:=lngPos).Font
    .FontStyle = "Bold"
End With

lngPos = InStr(Cells(r + 1, 3).value, ":")
With Cells(r + 1, 3).Characters(Start:=1, Length:=lngPos).Font
    .FontStyle = "Bold"
End With

lngPos = InStr(Cells(r + 2, 3).value, ":")
With Cells(r + 2, 3).Characters(Start:=1, Length:=lngPos).Font
    .FontStyle = "Bold"
End With

lngPos = InStr(Cells(r + 3, S).value, ":")
With Cells(r + 3, S).Characters(Start:=1, Length:=lngPos).Font
    .FontStyle = "Bold"
End With

lngPos = InStr(Cells(r + 4, S).value, ":")
With Cells(r + 4, S).Characters(Start:=1, Length:=lngPos).Font
    .FontStyle = "Bold"
End With

lngPos = InStr(Cells(r, S).value, ":")
With Cells(r, S).Characters(Start:=1, Length:=lngPos).Font
    .FontStyle = "Bold"
End With

lngPos = InStr(Cells(r + 1, S).value, ":")
With Cells(r + 1, S).Characters(Start:=1, Length:=lngPos).Font
    .FontStyle = "Bold"
End With

lngPos = InStr(Cells(r + 2, S).value, ":")
With Cells(r + 2, S).Characters(Start:=1, Length:=lngPos).Font
    .FontStyle = "Bold"
End With

lngPos = InStr(Cells(r + 3, 3).value, ":")
With Cells(r + 3, 3).Characters(Start:=1, Length:=lngPos).Font
    .FontStyle = "Bold"
End With

lngPos = InStr(Cells(r + 4, 3).value, ":")
With Cells(r + 4, 3).Characters(Start:=1, Length:=lngPos).Font
    .FontStyle = "Bold"
End With

lngPos = InStr(Cells(r + 5, S).value, ":")
With Cells(r + 5, S).Characters(Start:=1, Length:=lngPos).Font
    .FontStyle = "Bold"
End With

lngPos = InStr(Cells(r + 6, S).value, ":")
With Cells(r + 6, S).Characters(Start:=1, Length:=lngPos).Font
    .FontStyle = "Bold"
End With




'radek k podani zaloby-----------------------------------------------------------
Cells(lr + 1, 3).value = "Zùstatek ke dni podání žaloby, který se skládá z:"
Range(Cells(lr + 1, 3), Cells(lr + 1, 6)).MergeCells = True
Cells(lr + 1, 8).value = "Platby po podání žaloby"
Range(Cells(lr + 1, 8), Cells(lr + 1, 11)).MergeCells = True
Cells(lr + 2, 3).value = dz
Cells(lr + 2, 3).HorizontalAlignment = xlLeft
Cells(lr + 2, 5).value = nes_jis_zal
Cells(lr + 2, 6).value = kap_ur_zal
Cells(lr + 2, 4).value = Cells(lr + 2, 5).value + Cells(lr + 2, 6).value
'Range(Cells(lr + 1, 2), Cells(lr + 1, 6)).BorderAround _
' ColorIndex:=1, Weight:=xlMedium
'Range(Cells(lr + 1, 2), Cells(lr + 1, 6)).Interior.ColorIndex = 15

'radek k aktualnimu datu -------------------------------------------------------
'nes_jis_akt = nes_jis_zal - Application.Sum(Range(Cells(lr + 2, 9), Cells(dblRadek_posledni_uplne, 9)))
'kap_ur_akt = kap_ur_zal - Application.Sum(Range(Cells(lr + 2, 10), Cells(dblRadek_posledni_uplne, 10)))
Cells(dblRadek_posledni_uplne + 1, 3).value = "Aktuální zùstatek:"
Range(Cells(dblRadek_posledni_uplne + 1, 3), Cells(dblRadek_posledni_uplne + 1, 6)).MergeCells = True
Cells(dblRadek_posledni_uplne + 2, 3).value = date
Cells(dblRadek_posledni_uplne + 2, 3).HorizontalAlignment = xlRight
Cells(dblRadek_posledni_uplne + 2, 5).value = nes_jis_akt
Cells(dblRadek_posledni_uplne + 2, 6).value = kap_ur_akt
Cells(dblRadek_posledni_uplne + 2, 4).value = Cells(dblRadek_posledni_uplne + 2, 5).value + Cells(dblRadek_posledni_uplne + 2, 6).value


'Cells(dblRadek_posledni_pred_zalobou + 1, 5).Value = Cells(dblRadek_posledni_pred_zalobou, 10)
'formatovani bunek-------------------------------------------------------------
Range(Cells(dblRadek_prvni_uplne, 3), Cells(dblRadek_posledni_uplne + 2, 3)).HorizontalAlignment = xlRight
Range(Cells(dblRadek_prvni_uplne, 3), Cells(dblRadek_posledni_uplne + 2, 13)).VerticalAlignment = xlCenter
Range(Cells(dblRadek_prvni_uplne, 8), Cells(dblRadek_posledni_uplne + 2, 8)).HorizontalAlignment = xlRight
Range(Cells(dblRadek_prvni_po_zesp - 1, 8), Cells(dblRadek_prvni_po_zesp - 1, 11)).HorizontalAlignment = xlCenter
Range(Cells(dblRadek_prvni_po_zesp - 1, 3), Cells(dblRadek_prvni_po_zesp - 1, 6)).HorizontalAlignment = xlCenter
Range(Cells(lr + 1, 8), Cells(lr + 1, 13)).HorizontalAlignment = xlCenter
Range(Cells(lr + 1, 3), Cells(lr + 1, 6)).HorizontalAlignment = xlCenter
Range(Cells(dblRadek_posledni_uplne + 1, 3), Cells(dblRadek_posledni_uplne + 1, 6)).HorizontalAlignment = xlCenter
Range(Cells(15, 4), Cells(dblRadek_posledni_uplne + 2, 6)).NumberFormat = "#,##0.00"
Range(Cells(15, 9), Cells(dblRadek_posledni_uplne + 2, 13)).NumberFormat = "#,##0.00"

Range(Cells(5, 8), Cells(9, 8)).NumberFormat = "#,##0.00"


Range(Cells(dblRadek_prvni_uplne, 2), Cells(dblRadek_posledni_uplne + 2, 13)).Font.Size = 11
Range(Cells(dblRadek_prvni_uplne, 3), Cells(dblRadek_posledni_uplne + 2, 13)).Interior.ThemeColor = xlThemeColorDark1
Range(Cells(dblRadek_prvni_uplne, 3), Cells(dblRadek_posledni_uplne + 2, 13)).Interior.TintAndShade = -0.149998474074526
With Range(Cells(dblRadek_prvni_uplne, 3), Cells(dblRadek_posledni_uplne + 2, 13)).borders
   .LineStyle = xlContinuous
   .Weight = xlMedium
   .ColorIndex = 2
End With
 
 ' jak je definovany format u radku: zustatek ke dni zesplatneni....
Range(Cells(dblRadek_prvni_po_zesp, 3), Cells(dblRadek_prvni_po_zesp, 6)).BorderAround _
 ColorIndex:=2, Weight:=xlMedium
Range(Cells(dblRadek_prvni_po_zesp, 3), Cells(dblRadek_prvni_po_zesp, 6)).Interior.ThemeColor = xlThemeColorLight1
Range(Cells(dblRadek_prvni_po_zesp, 3), Cells(dblRadek_prvni_po_zesp, 6)).Interior.TintAndShade = 0.349986266670736
Range(Cells(dblRadek_prvni_po_zesp, 3), Cells(dblRadek_prvni_po_zesp, 6)).Font.Color = vbWhite

Cells(dblRadek_prvni_po_zesp - 1, 3).BorderAround _
 ColorIndex:=2, Weight:=xlMedium
Cells(dblRadek_prvni_po_zesp - 1, 3).Interior.ThemeColor = xlThemeColorLight1
Cells(dblRadek_prvni_po_zesp - 1, 3).Interior.TintAndShade = 0.349986266670736
Cells(dblRadek_prvni_po_zesp - 1, 3).Font.Color = vbWhite

Cells(dblRadek_prvni_po_zesp - 1, 8).HorizontalAlignment = xlCenter
Cells(dblRadek_prvni_po_zesp - 1, 8).BorderAround _
 ColorIndex:=2, Weight:=xlMedium
Cells(dblRadek_prvni_po_zesp - 1, 8).Interior.ThemeColor = xlThemeColorLight1
Cells(dblRadek_prvni_po_zesp - 1, 8).Interior.TintAndShade = 0.349986266670736
Cells(dblRadek_prvni_po_zesp - 1, 8).Font.Color = vbWhite


Cells(dblRadek_prvni_po_zesp - 1, 13).HorizontalAlignment = xlCenter
Cells(dblRadek_prvni_po_zesp - 1, 13).BorderAround _
 ColorIndex:=2, Weight:=xlMedium
Cells(dblRadek_prvni_po_zesp - 1, 13).Interior.ThemeColor = xlThemeColorLight1
Cells(dblRadek_prvni_po_zesp - 1, 13).Interior.TintAndShade = 0.349986266670736
Cells(dblRadek_prvni_po_zesp - 1, 13).Font.Color = vbWhite

Cells(lr + 1, 3).BorderAround ColorIndex:=2, Weight:=xlMedium
Cells(lr + 1, 3).Interior.ThemeColor = xlThemeColorLight1
Cells(lr + 1, 3).Interior.TintAndShade = 0.349986266670736
Cells(lr + 1, 3).Font.Color = vbWhite

Cells(lr + 1, 8).BorderAround _
 ColorIndex:=2, Weight:=xlMedium
Cells(lr + 1, 8).Interior.ThemeColor = xlThemeColorLight1
Cells(lr + 1, 8).Interior.TintAndShade = 0.349986266670736
Cells(lr + 1, 8).Font.Color = vbWhite


Cells(lr + 1, 13).BorderAround _
 ColorIndex:=2, Weight:=xlMedium
Cells(lr + 1, 13).Interior.ThemeColor = xlThemeColorLight1
Cells(lr + 1, 13).Interior.TintAndShade = 0.349986266670736
Cells(lr + 1, 13).Font.Color = vbWhite


Cells(dblRadek_posledni_uplne + 1, 3).BorderAround _
 ColorIndex:=2, Weight:=xlMedium
Cells(dblRadek_posledni_uplne + 1, 3).Interior.ThemeColor = xlThemeColorLight1
Cells(dblRadek_posledni_uplne + 1, 3).Interior.TintAndShade = 0.349986266670736
Cells(dblRadek_posledni_uplne + 1, 3).Font.Color = vbWhite



Range(Cells(lr + 2, 3), Cells(lr + 2, 6)).BorderAround _
 ColorIndex:=2, Weight:=xlMedium
Range(Cells(lr + 2, 3), Cells(lr + 2, 6)).Interior.ThemeColor = xlThemeColorLight1
Range(Cells(lr + 2, 3), Cells(lr + 2, 6)).Interior.TintAndShade = 0.349986266670736
Range(Cells(lr + 2, 3), Cells(lr + 2, 6)).Font.Color = vbWhite

Range(Cells(dblRadek_posledni_uplne + 2, 3), Cells(dblRadek_posledni_uplne + 2, 6)).BorderAround _
 ColorIndex:=2, Weight:=xlMedium
Range(Cells(dblRadek_posledni_uplne + 2, 3), Cells(dblRadek_posledni_uplne + 2, 6)).Interior.ThemeColor = xlThemeColorLight1
Range(Cells(dblRadek_posledni_uplne + 2, 3), Cells(dblRadek_posledni_uplne + 2, 6)).Interior.TintAndShade = 0.349986266670736
Range(Cells(dblRadek_posledni_uplne + 2, 3), Cells(dblRadek_posledni_uplne + 2, 6)).Font.Color = vbWhite
Range(Cells(dblRadek_prvni_uplne, 7), Cells(dblRadek_posledni_uplne + 2, 7)).Interior.ColorIndex = 2
Range(Cells(dblRadek_prvni_uplne, 12), Cells(dblRadek_posledni_uplne + 2, 12)).Interior.ColorIndex = 2

For i = dblRadek_prvni_po_zesp To dblRadek_posledni_uplne 'v tabulce sp 4 je obcas radek navic (pokud u splatky doslo k platbe po datumu podani zaloby), proto radeji umazu radek...
    If Application.WorksheetFunction.max(Range(Cells(i, 3), Cells(i, 13))) > 0 Then
    ElseIf Cells(i, 3).value = "" Then
        Rows(i).EntireRow.Delete
    End If
    
Next i

For i = dblRadek_prvni_uplne + 1 To dblRadek_posledni_uplne
    
    If Cells(i, 10).value + Cells(i, 11).value = 0 Then
        Cells(i, 9).value = ""
    Else
        Cells(i, 9).value = Cells(i, 10).value + Cells(i, 11).value
    End If
Next i


file_name = "Pøehled plateb úvìru è. " & hu & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-")
str_klient_folder = name_person & " " & Klient.Label2.Caption

Call exportuj_sheet(sheet_to_export:=cfg.wb_Jez.Sheets("Splatky"), _
                bln_export_do_PDF:=((volba_exportu = 1 Or volba_exportu = 3)), _
                fileName:=file_name, _
                folder1:=cfg.Cesta("EprPodaneZaloby") & str_klient_folder & "\Pøílohy\", _
                range_to_export:=Range(Cells(1, 1), Cells(dblRadek_posledni_uplne + 2, 13)), _
                folder2:=cfg.Cesta("EprPodaneZaloby") & str_klient_folder, _
                folder3:=cfg.Cesta("EprNachystane") & str_klient_folder & "\Pøílohy\")
                
                
                
''' ''' odtial presunute do fce exportuj_sheet
'''
'''If schema = "APP_COLL_JEZEVCIK_TEST" Then
'''    ' na testu se uklada do jine slozky nez na produkci
'''
'''    Range(Cells(1, 1), Cells(dblRadek_posledni_uplne + 2, 13)).ExportAsFixedFormat fileName:=cfg.Cesta("TestovaciSlozka") & "Pøehled plateb úvìru è. " & hu & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".pdf", Type:=xlTypePDF, _
'''        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
'''        :=False, OpenAfterPublish:=True
'''    'Call Shell("explorer.exe " & cfg.Cesta("TestovaciSlozka"), vbNormalFocus) 'otvirani slozky na testu
'''
'''Else ' na PROD se ukládá do nachystané složky a pokud neexistuje, do složky Klienti bez složky
'''    'vygenerovani pdf ---------------------------------------------------------
'''    AFolder = cfg.Cesta("EprPodaneZaloby") & name_person & " " & Klient.Label2.Caption
'''    BFolder = cfg.Cesta("EprNachystane") & name_person & " " & Klient.Label2.Caption
'''    CFolder = "W:\Vymahani\LATE\EPR\NACHYSTANÉ BØEZEN 2015\" & name_person & " " & Klient.Label2.Caption
'''
'''    AAFolder = cfg.Cesta("EprPodaneZaloby") & name_person & " " & Klient.Label2.Caption & "\Pøílohy"
'''    BBFolder = cfg.Cesta("EprNachystane") & name_person & " " & Klient.Label2.Caption & "\Pøílohy"
'''    CCFolder = "W:\Vymahani\LATE\EPR\NACHYSTANÉ BØEZEN 2015\" & name_person & " " & Klient.Label2.Caption & "\Pøílohy"
'''
'''
'''    Set fso = CreateObject("Scripting.FileSystemObject")
'''
'''    For i = 1 To 6
'''
'''        currFolder = Choose(i, AAFolder, AFolder, BBFolder, BFolder, CCFolder, CFolder)
'''        If fso.FolderExists(currFolder) Then
'''
'''            Range(Cells(1, 1), Cells(dblRadek_posledni_uplne + 2, 13)).ExportAsFixedFormat fileName:=currFolder & "\Pøehled plateb úvìru è. " & hu & "_" & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".pdf", Type:=xlTypePDF, _
'''                Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
'''                :=False, OpenAfterPublish:=True
'''            If zobrazit_slozku_klienta = True Then Call Shell("explorer.exe " & currFolder, vbNormalFocus)
'''            Exit For
'''
'''        Else
'''            If i = 6 Then
'''
'''                If Not blnSlozkaNeexistujeShown Then MsgBox "Složka pro klienta " & name_person & " neexistuje." ', vbInformation, "Složka nenalezena!"
'''                blnSlozkaNeexistujeShown = True
'''
'''                Range(Cells(1, 1), Cells(dblRadek_posledni_uplne + 2, 13)).ExportAsFixedFormat fileName:="Export_" & hu & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".pdf", Type:=xlTypePDF, _
'''                    Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
'''                    :=False, OpenAfterPublish:=True
'''
'''            End If
'''
'''        End If
'''
'''    Next i
'''
'''End If
'''
'''
''''Range(Cells(1, 1), Cells(lr + 1, 10)).ExportAsFixedFormat Filename:="D:\Users\pboril1374ab\Documents\seznam plateb pro jezevcika\Export_" & hu & Replace(Replace(CStr(Now()), " ", "_"), ":", "-") & ".pdf", Type:=xlTypePDF, _
''' '       Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
'''  '      :=False, OpenAfterPublish:=True
'''
'''
''''FA export XML odtial (risk-710)
'''If (volba_exportu = 2 Or volba_exportu = 4) Then
'''    wb_jez_name = ActiveWorkbook.name
'''    Sheets("Splatky").Visible = True
'''    Sheets("Splatky").Select
'''    Sheets("Splatky").Copy After:=wb_export_plateb.Sheets(wb_export_plateb.Sheets.count)
'''    wb_export_plateb.Sheets(Sheets.count).name = Right(Range("C4").value, 10)
'''    wb_export_plateb.Save
'''
'''    Workbooks(wb_jez_name).Activate
'''End If
''''FA export XML potial (risk-710)
'''
''' ''' potial presunute do fce exportuj_sheet



ws_Splatky.Visible = False

Exit Sub
connection_problem:

If Mid(Err.Description, 1, 9) = "ORA-03114" Then
err_connection = True
Else
 MsgBox Err.Description
End If
End Sub

