Attribute VB_Name = "mod_Export_plateb_KTK"
Option Explicit

'''   stuff to do
' ve vsech existujicich sablonach prejdi na metodu "exportuj_sheet", pokud to lze udelat
' otazka plneni jednotlivych poli
' otazka vyhozeni / pridani radku do worksheetu dle konkretniho klienta
' otazka defaultni formy sablony (necham tam posledniho klienta nebo se furt premaze?)

' pro testovani muzes pouzit metodu  "debug_risk_1145"



Private Enum Format_bunky
    format_bunky_tabulka_zvyrazneny
    format_bunky_tabulka_obycejny
    format_bunky_info_klient
End Enum


Private Type data_pro_sablonu
    ' uvidime jestli bude potreba
    neco As Variant
End Type





''' promenne ke specialnim pojmenovanym bunkam v listu

Private rngUvod_Vlevo_Prvni As Range
Private rngUvod_Vpravo_Prvni As Range
Private rngTabulka_prvni_bunka As Range

    Private rngSloupec_Typ_transakce As Range
    Private rngSloupec_Datum_transakce As Range
    Private rngSloupec_Cerpani_splaceni As Range
    Private rngSloupec_Cerpana_jistina As Range
    
    Private rngSloupec_Predpis_OU As Range
    Private rngSloupec_Predpis_SU As Range
    Private rngSloupec_Zaplaceno_datum As Range
    Private rngSloupec_Zaplaceno_OU As Range
    Private rngSloupec_Zaplaceno_SU As Range
    
    Private rngSloupce_tabulky As Range

Private rngTabulka_bunka_blokace As Range
Private rngTabulka_bunka_zesplatneni As Range
Private rngTabulka_bunka_zaloba As Range
Private rngTabulka_bunka_posledni As Range


Private sablona_data As data_pro_sablonu
Private dummy_cislo_ktk As String ' nahrad az budou data
Private dummy_range As Range
Private Const skp_overdraft_dummy As Long = 10



Public Sub test_me()
    Call NastavRanges
    ' rngUvod_Vlevo_Prvni.value = "Klient:  mentol"
    ' Debug.Print rngUvod_Vlevo_Prvni.BorderAround
End Sub


Public Sub exportuj_platby_KTK(Optional ByVal bln_export_do_excelu As Boolean = False)
    ' hlavni metoda, ktera se zavola zvenci
    ' bln_export_do_excelu = false -> exportuje PDF
    '                      = true  -> exportuje excel worksheet
    
    Dim file_name As String
    file_name = "KTK " & c.CUID & " " & Replace(Replace(CStr(Now()), " ", "_"), ":", "-")
    
    
    Call NastavRanges
    Call Promaz_sablonu ' (vraceni sheetu do defaultniho stavu)
    
    Call Vypln_info_v_uvodu
    
    Call Vypln_telo_tabulky
    Set dummy_range = SheetPrehledKTK.Range("A1:M30")
    Set dummy_range = Range(SheetPrehledKTK.Range("A1"), SheetPrehledKTK.Cells(rngTabulka_bunka_posledni.Row, rngSloupec_Zaplaceno_SU.Column))
    Set dummy_range = SheetPrehledKTK.Range("A1", SheetPrehledKTK.Cells(rngTabulka_bunka_posledni.Row, rngSloupec_Zaplaceno_SU.Column))
    Call Udelej_tabulku_hezkou
    
    ' pri exportu do excelu musi mit nazev mene nez 31 znaku :/
    
    If bln_export_do_excelu Then
        '''   .xlsx
        Call exportuj_sheet(sheet_to_export:=SheetPrehledKTK, _
                            bln_export_do_PDF:=False, _
                            fileName:=file_name, _
                            folder1:=CreateObject("WScript.Shell").specialfolders("Desktop") & "\", _
                            folder2:="", _
                            folder3:="")

    Else
        '''   .pdf
        Call exportuj_sheet(sheet_to_export:=SheetPrehledKTK, _
                            bln_export_do_PDF:=True, _
                            fileName:=file_name, _
                            folder1:=CreateObject("WScript.Shell").specialfolders("Desktop") & "\", _
                            range_to_export:=dummy_range, _
                            folder2:="", _
                            folder3:="")
    End If
    
    Call Promaz_sablonu
End Sub



Private Sub NastavRanges()
    ' tohle by melo byt jedine misto v modulu, kde jsou specificky pojmenovane radky / sloupecky
    ' zbytek sablony by uz mel pracovat jenom s range promennymi, takze zmeny je potreba delat tady.
    


    
    With SheetPrehledKTK
        
        '''   pojmenovane oblasti ve worksheet  (vypises je pomoci vlastni funkce PrintAllNamedRanges)
        Set rngUvod_Vlevo_Prvni = Range(.Names("Uvod_radek1_vlevo"))
        Set rngUvod_Vpravo_Prvni = Range(.Names("Uvod_radek1_vpravo"))
        Set rngTabulka_prvni_bunka = Range(.Names("tabulka_prvni_radek"))
        
        
        '''   deklarace sloupecku do kterych se budou zadavat hodnoty
        ' xx predelat na const
        Set rngSloupec_Typ_transakce = .Columns("C")
        Set rngSloupec_Datum_transakce = .Columns("D")
        Set rngSloupec_Cerpani_splaceni = .Columns("E")
        Set rngSloupec_Cerpana_jistina = .Columns("F")
        
        Set rngSloupec_Predpis_OU = .Columns("H")
        Set rngSloupec_Predpis_SU = .Columns("I")
        
        Set rngSloupec_Zaplaceno_datum = .Columns("K")
        Set rngSloupec_Zaplaceno_OU = .Columns("L")
        Set rngSloupec_Zaplaceno_SU = .Columns("M")

    End With
            
    Set rngSloupce_tabulky = Application.Union(rngSloupec_Typ_transakce, _
                                               rngSloupec_Datum_transakce, _
                                               rngSloupec_Cerpani_splaceni, _
                                               rngSloupec_Cerpana_jistina, _
                                               rngSloupec_Predpis_OU, _
                                               rngSloupec_Predpis_SU, _
                                               rngSloupec_Zaplaceno_datum, _
                                               rngSloupec_Zaplaceno_OU, _
                                               rngSloupec_Zaplaceno_SU)
    
    
''    With SheetPrehledKTK
''        Set rngSloupce_tabulky = Application.Union(.Columns("C:F"), .Columns("H:I"), .Columns("K:M"))
''    End With
    
    
    ' vyuziti:
'    Debug.Print Application.Intersect(rngSloupce_tabulky, Rows(rngTabulka_prvni_bunka.Row)).address
'    Application.Intersect(rngSloupce_tabulky, Rows(rngTabulka_prvni_bunka.Row)).Select
End Sub


Private Sub Promaz_sablonu()
    ' vymazou se vsechny radky ktere tvori telo tabulky
    ' vymazou se info v uvodu
    Dim nrows_to_delete As Integer
    Dim rngToDelete As Range
    
    
    
    nrows_to_delete = 300
    Set rngToDelete = SheetPrehledKTK.Range(rngTabulka_prvni_bunka.Offset(1, 0).EntireRow, rngTabulka_prvni_bunka.Offset(nrows_to_delete, 0).EntireRow)
    
    Call Vypln_info_v_uvodu(blnPromazat:=True)
    ' vymaz tak 500 radku od prvni bunky pod rngTabulka_prvni_bunka
    'Debug.Assert False ' mazanie riadkov zatial neotestovane
    rngToDelete.Delete
    
End Sub


Private Sub Vypln_info_v_uvodu(Optional ByRef blnPromazat As Boolean = False)
    ' tahle metoda vyplni informace o klientovi, zustatcich a pod. v uvodu
    Dim intPromazat As Integer
    
    If blnPromazat Then ' abych to mohl pouzit ve fci CHOOSE(), potrebuji ziskat 1 a 2
        intPromazat = 2
    Else
        intPromazat = 1
    End If
    
    With rngUvod_Vlevo_Prvni
        .value = cfg.SlovnikPrehledPlateb_KTK("info_klient") & Choose(intPromazat, c.name, "")
        .Offset(1, 0).value = cfg.SlovnikPrehledPlateb_KTK("info_datum_uzavreni") & Choose(intPromazat, "", "")
        '.Offset(2, 0).value = cfg.SlovnikPrehledPlateb_KTK("info_sazba") & Choose(intPromazat, "", "")
        .Offset(2, 0).value = cfg.SlovnikPrehledPlateb_KTK("info_datum_cerpani") & Choose(intPromazat, "", "")
    End With
    
    With rngUvod_Vpravo_Prvni
        .value = cfg.SlovnikPrehledPlateb_KTK("info_limit") & Choose(intPromazat, "", "")
        .Offset(1, 0).value = cfg.SlovnikPrehledPlateb_KTK("info_jistina_zesp") & Choose(intPromazat, "", "")
        .Offset(2, 0).value = cfg.SlovnikPrehledPlateb_KTK("info_jistina_zaloba") & Choose(intPromazat, "", "")
        .Offset(3, 0).value = cfg.SlovnikPrehledPlateb_KTK("info_jistina_dnes") & Choose(intPromazat, "", "")
    End With
    
    
End Sub

Private Sub Vypln_telo_tabulky()
    ' doplni se transakce a castky
    
    Dim rst As ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    Dim current_row As Range
    Dim strLastFlagDay As String
    
    
    sql = "SELECT a.SKP_BANK_ACCOUNT, a.SKP_OVERDRAFT, a.ID_PERSON, a.AMT_TRANSACTION_CLEAN, a.DTIME_TRANSACTION_ACCOUNTED, TRUNC(a.DTIME_TRANSACTION_ACCOUNTED) as datum_trunc, a.TEXT_NOTE, " & _
          "       a.DATUM_POSLEDNI_ZANORENI , a.POSLEDNI_KLADNY_ZUSTATEK, a.AMT_OVERDRAFT_CUMULATIVE, a.FLAG_SUMA_ZA_DEN_OK, a.FLAG_SUMA_CELKOVA_OK " & _
          "  FROM " & schema & ".V_OVERDRAFT_PRUBEH a WHERE a.ID_PERSON = " & c.id_person
    ' dotaz presunut do objektu? alebo miesto ID_PERSON pouzit SKP_OVERDRAFT?
    
    Set rst = New ADODB.Recordset
    rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        For i = 1 To rst.RecordCount
            
            '''   nastavi se aktualni radek
            Set current_row = rngTabulka_prvni_bunka.Offset(i, 0).EntireRow
            
            '''   zkotroluji se FLAGy
            If (i = 1) And (nvl(rst![FLAG_SUMA_CELKOVA_OK]) <> "Y") Then
                HlaskaUzivateli "Celková suma transakcí neodpovídá výšce naèerpané jistiny. Nìkde v šablonì je chyba nebo chybí nìjaká transakce."
            End If
            
            If (strLastFlagDay <> nvl(rst![datum_trunc])) And (nvl(rst![FLAG_SUMA_ZA_DEN_OK]) <> "Y") Then
                HlaskaUzivateli "Celková suma transakcí za den " & nvl(rst![datum_trunc]) & " neodpovídá zmìne naèerpané jistiny za ten den. Nìkde v šablonì je chyba nebo chybí nìjaká transakce."
                strLastFlagDay = nvl(rst![datum_trunc])
            End If
            
            
            '''   Naplneni hodnot do tabulky
            ' If: sa jedna o specialny riadok (blokace, zesplatneni, podani zaloby, dnes) - zistis cez atributy objektu Overdraft
                ' uloz si jeho poziciu do konkretnej premennej
                ' vloz specialne hodnoty (nejake dynamicke kumulativne sumy? alebo extra select? Spravne riesenie je asi Extra VIEW k tomu + kontrola cez jednoduche sumy spravnych buniek)
            ' Else
                ' vloz hodnoty do stlpcov spravnych (vkladaj vzdy do prieniku aktualneho riadku a stlpca ulozeneho v range premennej )
                Application.Intersect(current_row, rngSloupec_Typ_transakce).value = Left(nvl(rst![TEXT_NOTE]), 40)
                Application.Intersect(current_row, rngSloupec_Datum_transakce).value = nvl(rst![datum_trunc])
                Application.Intersect(current_row, rngSloupec_Cerpani_splaceni).value = VlozCastku((rst![AMT_TRANSACTION_CLEAN]))
                Application.Intersect(current_row, rngSloupec_Cerpana_jistina).value = VlozCastku((rst![AMT_OVERDRAFT_CUMULATIVE]))
                
                'Application.Intersect(current_row, rngSloupec_Predpis_OU).value = nvl(rst![TEXT_NOTE])
                'Application.Intersect(current_row, rngSloupec_Predpis_SU).value = nvl(rst![TEXT_NOTE])
                'Application.Intersect(current_row, rngSloupec_Zaplaceno_datum).value = nvl(rst![TEXT_NOTE])
                'Application.Intersect(current_row, rngSloupec_Zaplaceno_OU).value = nvl(rst![TEXT_NOTE])
                'Application.Intersect(current_row, rngSloupec_Zaplaceno_SU).value = nvl(rst![TEXT_NOTE])
                
            ' End if
            rst.MoveNext
        Next i
        Set rngTabulka_bunka_posledni = Application.Intersect(current_row, rngSloupec_Typ_transakce)
    rst.Close
    
    
    
End Sub


Private Sub Udelej_tabulku_hezkou()
    ' tady se nastavi barvicky a ohraniceni vsech bunek tabulky
    ' taky vyska radku
    ' ... a pod...
    
    Dim rngTeloTabulky As Range
    
    
    Set rngTeloTabulky = Application.Intersect(Range(rngTabulka_prvni_bunka.Offset(1, 0).EntireRow, rngTabulka_bunka_posledni.EntireRow) _
                                               , rngSloupce_tabulky)
    
    rngTeloTabulky.RowHeight = 17
    Application.Intersect(rngSloupec_Cerpana_jistina, rngTeloTabulky).HorizontalAlignment = xlHAlignRight
    Application.Intersect(rngSloupec_Cerpani_splaceni, rngTeloTabulky).HorizontalAlignment = xlHAlignRight
    
    '''   nastaveni nazvu sloupecku aby se opakovali na kazde strance
    SheetPrehledKTK.PageSetup.PrintTitleRows = Range(rngTabulka_prvni_bunka.EntireRow, rngTabulka_prvni_bunka.Offset(-2, 0).EntireRow).address
    
    '''   formaty bunek
    If (Not rngTeloTabulky Is Nothing) Then Call Zvyrazni_bunky(rngTeloTabulky, format_bunky_tabulka_obycejny)
    If (Not rngTabulka_bunka_blokace Is Nothing) Then Call Zvyrazni_bunky(rngTabulka_bunka_blokace, format_bunky_tabulka_zvyrazneny)
    If (Not rngTabulka_bunka_zesplatneni Is Nothing) Then Call Zvyrazni_bunky(rngTabulka_bunka_zesplatneni, format_bunky_tabulka_zvyrazneny)
    If (Not rngTabulka_bunka_zaloba Is Nothing) Then Call Zvyrazni_bunky(rngTabulka_bunka_zaloba, format_bunky_tabulka_zvyrazneny)

    
End Sub



Private Sub Zvyrazni_bunky(ByRef bunky As Range, zvoleny_format As Format_bunky)
    ' pro nastaveni preddefinovaneho formatu pro nejakou oblast
    ' pozor na spravnou syntax! Musi se vlozit "Zvyrazni_bunky xxx" nebo "Call Zvyrazni_bunky (xxx)" -> pokud das zavorky v prvnim pripade, dostanes runtime error 424
    ' tohle je jedine misto, kde by meli byt nadefinovane jednotlive formaty
    
    ' test: call mod_Export_plateb_KTK.Zvyrazni_bunky(SheetPrehledKTK1.Range("E29"), format_bunky_tabulka_zvyrazneny)
    
    Select Case zvoleny_format
        
        Case format_bunky_tabulka_zvyrazneny
            With bunky
                .HorizontalAlignment = xlCenter
                .BorderAround ColorIndex:=2, Weight:=xlMedium
                .Interior.ThemeColor = xlThemeColorLight1
                .Interior.TintAndShade = 0.349986266670736
                .Font.Color = vbWhite
            End With
                
        Case format_bunky_tabulka_obycejny
            With bunky
                .Font.Size = 11
                .Interior.ThemeColor = xlThemeColorDark1
                .Interior.TintAndShade = -0.149998474074526
                
                .borders.LineStyle = xlContinuous
                .borders.Weight = xlMedium
                .borders.ColorIndex = 2
            End With
            
        Case format_bunky_info_klient
            ' todo
            
        Case Else
            assertion 1 = 0, "jez-91211: neexistující formát bunìk" ' should never happen
            
    End Select
    
    ' bunky.height nastavit nejak...
    
End Sub


Private Sub uloz_sablonu()

End Sub
