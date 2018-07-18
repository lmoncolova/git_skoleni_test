VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataCheck 
   Caption         =   "Jezevèík v 2.6 - Kontrola hodnot"
   ClientHeight    =   12285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12615
   OleObjectBlob   =   "DataCheck.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit
Private Listview1_last_clicked_item As ListItem



Private Sub button_export_chytre_rezervy_Click()
    ' RISK-1135
    
    
    
End Sub

Sub button_export_plateb_Click() 'FA_export, button v userform pribudol.
    Dim strNazev As String
    
    If cs_HDWMN.State = 1 Then
prihlasen:
    
    On Error GoTo chyba
       'Application.EnableEvents = False
       Application.ScreenUpdating = False
       Application.DisplayAlerts = False
    
    'vo formulari export_plateb si uzivatel vyberie moznost 1-4: 1) PDF k tomuto uveru, 2) EXCEL k tomuto uveru, 3) PDF ku vsetkym uverom klienta, 4) XML ku vsetkym uverom klienta
    If export_chosen_option = 0 Then Export_plateb.Show
    If export_chosen_option = 0 Then 'Cancel Button
        'export_chosen_option = 0
        Exit Sub 'cancel button
    End If
    
    Dim sql_vsetky_uvery As String, rst_vsetky_uvery As Variant, uver_n As String, zobrazit_slozku_klienta As Boolean
    zobrazit_slozku_klienta = True
    
    'pre export excelu si pripravime workbook
    If (export_chosen_option = 2 Or export_chosen_option = 4) Then
        'create workbook for XML
        Set wb_export_plateb = Workbooks.Add
        strNazev = "export plateb " & c.name & " " & Replace(Replace(CStr(Now()), " ", "_"), ":", "-")
        
        With wb_export_plateb
            '.Sheets("List1").Delete 'pozor, u excelu 2013 ma novy workbook iba 1 sheet, (naviac to moze byt anglicky)... Radsej nemazat.
            .Title = strNazev
            .Subject = strNazev '"export plateb " & c.name
            .SaveAs fileName:=cfg.Cesta("TestovaciSlozka") & strNazev & ".xlsx"
            '.SaveAs fileName:=CreateObject("WScript.Shell").specialfolders("Desktop") & "\" & strNazev & ".xlsx" 'path to Desktop, regardless of user
        End With
        
    End If
    
    
    '''   export plateb k jednemu uveru
    
    If (export_chosen_option = 1 Or export_chosen_option = 2) Then
        If (InStr(Listview1_last_clicked_item.text, "KTK") > 0) Then
            Call exportuj_platby_KTK(bln_export_do_excelu:=Choose(export_chosen_option, False, True))
        Else
            Call export_seznam_plateb("select * from " & schema & ".sp2 where uver_cislo = '" & Listview1_last_clicked_item.text & "' order by datum_splatky, datum_platby ", _
                                      "select * from " & schema & ".sp3 where code_loan_number = '" & Listview1_last_clicked_item.text & "' order by date_instalment", _
                                      "select * from " & schema & ".sp1 where uver_cislo = '" & Listview1_last_clicked_item.text & "' order by datum_splatky, datum_platby ", _
                                      "select * from " & schema & ".sp4 where uver_cislo = '" & Listview1_last_clicked_item.text & "' order by date_instalment ", _
                                      "select * from " & schema & ".sp5 where uver_cislo = '" & Listview1_last_clicked_item.text & "' order by datum_splatky, datum_platby ", _
                                      "select * from " & schema & ".sp6 where code_loan_number = '" & Listview1_last_clicked_item.text & "'", _
                                      "select * from " & schema & ".sp8 where uver_cislo = '" & Listview1_last_clicked_item.text & "' order by date_instalment ", _
                                      Listview1_last_clicked_item.text, volba_exportu:=export_chosen_option)
        End If
    
    '''   export plateb ku vsetkym uverom klienta
    
    ElseIf (export_chosen_option = 3 Or export_chosen_option = 4) Then
        'ten isty SQL dotaz co listview1 (t.j. zoznam uctov navrchu v Kontrole dat), akurat ze vyberieme iba HU
        
        sql_vsetky_uvery = "select cislo, status_2, typ_smlouvy from " & schema & ".v_datacheck_contracts a where id_person = " & id_person & " and typ_smlouvy IN ('HU', 'KTK') AND a.STATUS_2 NOT IN ('Stornovaný', 'Ukonèený')"
        
        Set rst_vsetky_uvery = New ADODB.Recordset
            rst_vsetky_uvery.Open sql_vsetky_uvery, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
          
            If Not rst_vsetky_uvery.EOF Then
                rst_vsetky_uvery.MoveFirst
                
                Do Until rst_vsetky_uvery.EOF = True
                    uver_n = "" & rst_vsetky_uvery![cislo] & ""
                    
                    If (InStr(uver_n, "KTK") > 0) Then
                        Call exportuj_platby_KTK(bln_export_do_excelu:=Choose(export_chosen_option - 2, False, True))
                    Else
                        Call export_seznam_plateb("select * from " & schema & ".sp2 where uver_cislo = '" & uver_n & "' order by datum_splatky, datum_platby ", _
                                                    "select * from " & schema & ".sp3 where code_loan_number = '" & uver_n & "' order by date_instalment", _
                                                    "select * from " & schema & ".sp1 where uver_cislo = '" & uver_n & "' order by datum_splatky, datum_platby ", _
                                                    "select * from " & schema & ".sp4 where uver_cislo = '" & uver_n & "' order by date_instalment ", _
                                                    "select * from " & schema & ".sp5 where uver_cislo = '" & uver_n & "' order by datum_splatky, datum_platby ", _
                                                    "select * from " & schema & ".sp6 where code_loan_number = '" & uver_n & "'", _
                                                    "select * from " & schema & ".sp8 where uver_cislo = '" & uver_n & "' order by date_instalment ", _
                                                    uver_n, volba_exportu:=export_chosen_option, zobrazit_slozku_klienta:=zobrazit_slozku_klienta)   'posledny parameter podla ktoreho mi vytvori pdf / excel
                    End If
                    
                    zobrazit_slozku_klienta = False 'pouze jednou
                    rst_vsetky_uvery.MoveNext
                Loop
            
            End If
        rst_vsetky_uvery.Close
    
    End If
    
    
    If (export_chosen_option = 2 Or export_chosen_option = 4) Then
        MsgBox "Export ukonèen. Excel soubor je ve složce " & cfg.Cesta("TestovaciSlozka") 'plose? to je spravne??
    End If
    
    export_chosen_option = 0 'aby sa znova vyhodilo okno export_plateb
    
    Set wb_export_plateb = Nothing
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Exit Sub


chyba:
     Set wb_export_plateb = Nothing
    ' pokud nastala chyba pøipojením, tak ...
     If err_connection Then
      err_connection = False
      Call re_login(Main.username.text, Main.password.text)
      If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
     End If
    
    Else
     Call re_login(Main.username.text, Main.password.text)
     If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
    End If
    MsgBox Err.Number & ": " & Err.Description
End Sub





Private Sub CommandButton3_Click()
If cs_HDWMN.State = 1 Then
prihlasen:

On Error GoTo chyba



'MsgBox "Pozor! Zatím pouze ve zkušebním provozu"
   Application.EnableEvents = False
   Application.ScreenUpdating = False
   Application.DisplayAlerts = False
   

'muz ci zena
Dim sql0 As String

Dim sql712 As String, SKP As Variant

Set rst1 = New ADODB.Recordset
'zjistime jmeno klienta
sql712 = "select id_person from " & schema & ".INTERVALY where code_loan_number = '" & Listview1_last_clicked_item.text & "'"
rst1.Open sql712, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
    If Not rst1.EOF Then
        SKP = rst1(0)
    End If
    
rst1.Close


'''   3 radky zakomentovany ve verzi jez.v.5.9 spolu s celym modulem. Duvody jsou uvedeny v modulu.
Select Case c.Gender
    Case "Z"
        ' Call export_text_zaloby_zena(Listview1_last_clicked_item.text, SKP)

    Case "M"
        ' Call export_text_zaloby_muz(Listview1_last_clicked_item.text, SKP)
    
    Case Else
        ' Call export_text_zaloby_muz(Listview1_last_clicked_item.text, SKP)
End Select

   Application.EnableEvents = True
   Application.ScreenUpdating = True
   Application.DisplayAlerts = True
   
' tady vypise pouze splatky, u kterych prisla platba do zesplatneni
chyba:

' pokud nastala chyba pøipojením, tak ...
  If err_connection Then
  err_connection = False
  Call re_login(Main.username.text, Main.password.text)
  If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
 End If

Else
 Call re_login(Main.username.text, Main.password.text)
 If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If


End Sub



'Private Sub CommandButton4_Click()
'If cs_HDWMN.State = 1 Then
'prihlasen:
'
'On Error GoTo chyba
'
'
'
'
'   Application.EnableEvents = False
'   Application.ScreenUpdating = False
'   Application.DisplayAlerts = False
'
'Dim sql712 As String, skp As Variant
'
'Set rst712 = New ADODB.Recordset
''zjistime skp klienta
'   sql712 = "select id_person from " & schema & ".INTERVALY where code_loan_number = '" & Listview1_last_clicked_item.text & "'"
'   rst712.Open sql712, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'
'   If Not rst712.EOF Then
'    skp = rst712(0)
'   End If
'rst712.Close
'
'Call uznani_dluhu(Listview1_last_clicked_item.text, skp)
'
'
'
'   Application.EnableEvents = True
'   Application.ScreenUpdating = True
'   Application.DisplayAlerts = True
'
'' tady vypise pouze splatky, u kterych prisla platba do zesplatneni
'chyba:
'
'' pokud nastala chyba pøipojením, tak ...
'  If err_connection Then
'  err_connection = False
'  Call re_login(Main.username.text, Main.password.text)
'  If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
' End If
'
'Else
' Call re_login(Main.username.text, Main.password.text)
' If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
'End If
'
'End Sub

Private Sub Label73_Click()

End Sub

' vybrání nìkterého z èísla úètu/úvìru
Private Sub ListView1_DblClick()
Dim sql As String
Dim control As control

' Pokud jsem pøipojen
If cs_HDWMN.State = 1 Then
prihlasen:

Call text_information_hide(Me.name)
' pøíprava formuláøe na data
ListView2.ListItems.Clear
TextBox1.text = ""
TextBox2.text = ""
TextBox3.text = ""
TextBox4.text = ""
TextBox5.text = ""
TextBox6.text = ""


If ListView1.ListItems.count > 0 Then

    Set Listview1_last_clicked_item = ListView1.SelectedItem
    '''   HU a HR
    If ((InStr(UCase(Listview1_last_clicked_item.text), UCase("HU")) > 0) Or (InStr(UCase(Listview1_last_clicked_item.text), UCase("HR")) > 0)) Then
        For Each control In Me.Controls
            If control.Tag = "1" Then control.Visible = True 'podle teto hodnoty se nastavi, zda ma byt objekt viditelny nebo ne
            If control.Tag = "2" Then control.Visible = False
        Next
        
        
        With Me.button_export_chytre_rezervy
            If (InStr(UCase(Listview1_last_clicked_item.text), UCase("HR")) > 0) Then
                .Enabled = True
                .ControlTipText = "wof"
            Else
                .Enabled = False
                .ControlTipText = "Export chytré rezervy funguje pouze pro hypotéky."
            End If
        End With
        
        Call query1("select * from " & schema & ".v_datacheck_credit_A where id_person = " & id_person & " and uver_cislo = '" & Listview1_last_clicked_item.text & "'", Me.name & "CreditA")
        Call query("select * from " & schema & ".v_datacheck_credit_I where id_person = " & id_person & " and uver_cislo = '" & Listview1_last_clicked_item.text & "' order by 3 DESC", Me.name & "CreditI")
        
        Call query1("select * from " & schema & ".v_datacheck_platby where id_person = " & id_person & " and uver_cislo = '" & Listview1_last_clicked_item.text & "'", Me.name & "Platby")
        
        Label41.Caption = Listview1_last_clicked_item
        Label70.Caption = Listview1_last_clicked_item
        
        If ListView2.ListItems.count > 0 Then
            ListView2.ListItems.item(1).Selected = True
            Call ListView2_DblClick
        End If
  
        
        Label27.Top = 588
        Label28.Top = 588
        Label29.Top = 588
        Label27.Left = 6
        Label28.Left = 6
        Label29.Left = 24
        Me.Height = 628
        
        ' Nefungoval mi tu Me.Repaint, tak jsem použil "znovuvykreslení" pomocí visible
        ListView2.Visible = Not ListView2.Visible
        ListView2.Visible = Not ListView2.Visible
        ListView3.Visible = Not ListView3.Visible
        ListView3.Visible = Not ListView3.Visible
        
    
    '''   KTK
    ElseIf (InStr(UCase(Listview1_last_clicked_item), UCase("KTK")) > 0) Then
    
        For Each control In Me.Controls
            If control.Tag = "1" Then control.Visible = True 'podle teto hodnoty se nastavi, zda ma byt objekt viditelny nebo ne
            If control.Tag = "2" Then control.Visible = False
            Me.Frame5.Visible = False
        Next
        
        
        Call NaplnListviewsKontokorentem

        Label41.Caption = Listview1_last_clicked_item
        Label70.Caption = Listview1_last_clicked_item
        
        If ListView2.ListItems.count > 0 Then
            ListView2.ListItems.item(1).Selected = True
            Call ListView2_DblClick
        End If
  
        
        Label27.Top = 588
        Label28.Top = 588
        Label29.Top = 588
        Label27.Left = 6
        Label28.Left = 6
        Label29.Left = 24
        Me.Height = 628
        
        ' Nefungoval mi tu Me.Repaint, tak jsem použil "znovuvykreslení" pomocí visible
        ListView2.Visible = Not ListView2.Visible
        ListView2.Visible = Not ListView2.Visible
        ListView3.Visible = Not ListView3.Visible
        ListView3.Visible = Not ListView3.Visible
    
    
    '''   ucty
    Else
    
        For Each control In Me.Controls
            If control.Tag = "1" Then control.Visible = False
            If control.Tag = "2" Then control.Visible = True
        Next
  
        Call query("select amt_balance, code_currency from " & schema & ".v_datacheck_debit " & _
                   "where id_person = " & id_person & " and cislo_uctu = '" & Listview1_last_clicked_item.text & "'" _
                   , Me.name & "Debit")
        Call query1("select DTIME_TRANSACTION_GENERATION, DTIME_TRANSACTION_ACCOUNTED, amt_transaction, text_note " & _
                    " FROM " & schema & ".v_datacheck_platby_ucet where id_person = " & c.id_person & " and cislo_uctu = '" & Listview1_last_clicked_item.text & "'" _
                    , Me.name & "Debit")
        ' uprava RISK-602
  
  
         Label27.Visible = True
         Label29.Visible = True
         'Label29.Caption = "Další informace nejsou k bankovním úètùm evidovány"
         Label102.Caption = Listview1_last_clicked_item
        
         Label27.Top = 185
         Label28.Top = 185
         Label29.Top = 185
         Label27.Left = 6 + 400
         Label28.Left = 6 + 400
         Label29.Left = 24 + 400
         Me.Height = 400
         
 
    End If

End If



' pokud nastala chyba pøipojením, tak ...
 If err_connection Then
  err_connection = False
  Call re_login(Main.username.text, Main.password.text)
  If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
 End If





Else
 Call re_login(Main.username.text, Main.password.text)
 If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If
End Sub


Private Sub NaplnListviewsKontokorentem()
    
    Dim itmX As ListItem
    Dim sql_prubeh As String
    
    sql_prubeh = "SELECT a.TEXT_NOTE, TRUNC(a.DTIME_TRANSACTION_ACCOUNTED) AS datum_trunc, a.AMT_TRANSACTION_CLEAN, a.AMT_OVERDRAFT_CUMULATIVE " & _
                 " FROM " & schema & ".v_overdraft_prubeh a WHERE a.ID_PERSON = " & c.id_person
    
'    Call query1("select * from " & schema & ".v_datacheck_credit_A where id_person = " & id_person & " and uver_cislo = '" & Listview1_last_clicked_item.text & "'", Me.name & "CreditA")
'    Call query("select * from " & schema & ".v_datacheck_credit_I where id_person = " & id_person & " and uver_cislo = '" & Listview1_last_clicked_item.text & "' order by 3 DESC", Me.name & "CreditI")
'    Call query1("select * from " & schema & ".v_datacheck_platby where id_person = " & id_person & " and uver_cislo = '" & Listview1_last_clicked_item.text & "'", Me.name & "Platby")
    
    
    '''   CreditA
'        'Case "DataCheckCreditA" Credit Aktualni
'        Dim SKP As Double, sql2 As String
'        Set rst1 = New ADODB.Recordset
'        rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'        If Not rst1.EOF Then
'           DataCheck.TextBox1.text = "" & rst1(2) & ""
'           DataCheck.TextBox2.text = "" & rst1(3) & ""
'           DataCheck.TextBox3.text = "" & rst1(4) & ""
'           DataCheck.TextBox4.text = "" & rst1(5) & ""
'           DataCheck.TextBox5.text = "" & rst1(6) & ""
'           DataCheck.TextBox10.text = "" & rst1(7) & ""
'           DataCheck.TextBox11.text = rst1(5) + rst1(4) + rst1(3) + rst1(6) 'soucet pro uver
'           SKP = rst1(0)
'        End If
'        rst1.Close
'
'         If SKP > 0 Then
'        Set rst2 = New ADODB.Recordset
'        sql2 = "select sum(amt_balance_principal) + sum(kapitalizovnay_urok) + sum(ou_total)+sum(su_total) from " & schema & ".v_datacheck_credit_A where id_person = " & SKP & " group by id_person "
'
'        rst2.Open sql2, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'
'        If Not rst2.EOF Then
'           DataCheck.TextBox12.text = "" & rst2(0) & ""
'        End If
'        End If
'
'
'    '''   Platby
'
'        'Case "DataCheckPlatby"
'        DataCheck.ListView3.ListItems.Clear
'
'        Set rst1 = New ADODB.Recordset
'        'asss = InputBox("", "", sql)
'        rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'
'            If Not rst1.EOF Then
'
'
'                rst1.MoveFirst
'                For i = 1 To rst1.RecordCount
'                    Set itmX = DataCheck.ListView3.ListItems.Add()
'                    itmX.text = "" & rst1(2) & ""
'                    itmX.SubItems(1) = "" & rst1(3) & ""
'                    itmX.SubItems(2) = "" & rst1(4) & ""
'                    itmX.SubItems(3) = "" & rst1(5) & ""
'                    itmX.SubItems(4) = "" & rst1(6) & ""
'                    itmX.SubItems(5) = "" & rst1(7) & ""
'                    rst1.MoveNext
'                Next i
'            End If
'
'        rst1.Close
'
'        Call Listview_Proper_Sort(DataCheck.ListView3, DataCheck.ListView3.ColumnHeaders(1), blnAscendingOrder:=True)   ' "Datum platby"
'
'    '''   CreditI
'        '- Credit: intervaly s daty
'        'Case "DataCheckCreditI"
        
        
        
        With ListView2.ColumnHeaders
            .Clear
            ' spolu sirka 600
            .Add , , "Typ transakce", 180 '80
            .Add , , "Datum transakce", 80
            .Add , , "Èerpání/Splácení", 80
            .Add , , "Èerpaná jistina", 80
            .Add , , "(...)", 80 'Platnost do (SÚ)
            .Add , , "(...)", 0 '100 'Výše SÚ
            .Add , , "(...)", 100 'Výše jistiny
        End With
        
        With ListView2
            .ListItems.Clear
            .Height = 155
        End With

        Set rst1 = New ADODB.Recordset
        'asss = InputBox("", "", sql)
        rst1.Open sql_prubeh, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
            If Not rst1.EOF Then
    
                rst1.MoveFirst
                For i = 1 To rst1.RecordCount
                    Set itmX = ListView2.ListItems.Add()
                    itmX.text = "" & rst1(0) & ""
                    itmX.SubItems(1) = "" & rst1(1) & ""
                    itmX.SubItems(2) = "" & rst1(2) & ""
                    itmX.SubItems(3) = "" & rst1(3) & ""
'                    itmX.SubItems(4) = "" & rst1(6) & ""
'                    itmX.SubItems(5) = "" & rst1(7) & ""
'                    itmX.SubItems(6) = "" & rst1(8) & ""
                    rst1.MoveNext
                Next i
            End If
        rst1.Close

End Sub


Private Sub ListView1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)

If (KeyCode = 67 And Shift = 2) Then
    Call text_information_hide(Me.name)
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject
    clipboard.SetText ListView1.SelectedItem
    clipboard.PutInClipboard
    
    Label27.Visible = True
    Label29.Visible = True
    Label29.Caption = "Èíslo bylo zkopírováno do schránky"
End If

End Sub

Private Sub ListView2_DblClick()
Dim sql As String
Dim control As control

If cs_HDWMN.State = 1 Then
prihlasen:

'''   nasledujici kus kodu chceme pouze pro uvery a hypoteky. Slouceno Listview2 a Listview4 ve verzi jez.v.5.7
If ((InStr(UCase(Listview1_last_clicked_item.text), UCase("HU")) > 0) Or (InStr(UCase(Listview1_last_clicked_item.text), UCase("HR")) > 0)) Then
    Call text_information_hide(Me.name)
    
    TextBox7.text = ""
    TextBox8.text = ""
    TextBox9.text = ""
    
    Label1.Caption = ""
    Label2.Caption = ""
    Label3.Caption = ""
    
    
    
    If ListView2.ListItems.count > 0 Then
    
        Label42.Caption = " od " & ListView2.SelectedItem & " do " & ListView2.SelectedItem.SubItems(1)
        
        TextBox7.text = ListView2.SelectedItem.SubItems(2) ' Obchodní úrok
        Label1.Caption = "od " & ListView2.SelectedItem & " do " & ListView2.SelectedItem.SubItems(1)
        TextBox8.text = ListView2.SelectedItem.SubItems(5) ' Sankèní úrok
        Label2.Caption = "od " & ListView2.SelectedItem.SubItems(3) & " do " & ListView2.SelectedItem.SubItems(4)
        TextBox9.text = ListView2.SelectedItem.SubItems(6) ' Jistina
        Label3.Caption = "od " & ListView2.SelectedItem & " do " & ListView2.SelectedItem.SubItems(1)
    End If
End If



 If err_connection Then
  err_connection = False
  Call re_login(Main.username.text, Main.password.text)
  If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
 End If

Else
 Call re_login(Main.username.text, Main.password.text)
 If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If
End Sub


Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call Listview_Proper_Sort(ListView2, ColumnHeader)
End Sub

Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call Listview_Proper_Sort(ListView3, ColumnHeader)
End Sub

'Private Sub ListView4_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    Call Listview_Proper_Sort(ListView4, ColumnHeader)
'End Sub

Private Sub ListView3_ItemClick(ByVal item As MSComctlLib.ListItem)


' Pokud jsem pøipojen
If cs_HDWMN.State = 1 Then
prihlasen:

Call text_information_hide(Me.name)

If ListView1.ListItems.count > 0 Then
 Label28.Visible = True
 Label29.Visible = True
 Call query1("select id_person, uver_cislo, datum_platby, sum(amt_payment) from " & schema & ".v_datacheck_platby where id_person = " & id_person & " and uver_cislo = '" & Listview1_last_clicked_item.text & "' and datum_platby = '" & ListView3.SelectedItem & "' group by id_person, uver_cislo, datum_platby", Me.name & "PlatbyCelkem")

End If

' pokud nastala chyba pøipojením, tak ...
 If err_connection Then
  err_connection = False
  Call re_login(Main.username.text, Main.password.text)
  If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
 End If

Else
 Call re_login(Main.username.text, Main.password.text)
 If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If


End Sub



' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()
Dim sql As String
Dim control As control

If cs_HDWMN.State = 1 Then
prihlasen:

DataCheck.Caption = "Jezevèík v." & version & " - Kontrola hodnot"


Call text_information_hide(Me.name)

    DataCheck.Label81.Caption = ""
    DataCheck.Label82.Caption = ""
    DataCheck.Label83.Caption = ""
    DataCheck.Label84.Caption = ""
    DataCheck.Label85.Caption = ""


Call query("select cislo, status_2, typ_smlouvy from " & schema & ".v_datacheck_contracts where id_person = " & id_person, Me.name)



If ListView1.ListItems.count > 0 Then
 ListView1.ListItems.item(1).Selected = True
 Call ListView1_DblClick
  Me.Repaint
End If


 If err_connection Then
  err_connection = False
  Call re_login(Main.username.text, Main.password.text)
  If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
 End If

Else
 Call re_login(Main.username.text, Main.password.text)
 If cs_HDWMN.State = 1 Then GoTo prihlasen Else Call logout
End If
End Sub


Private Sub UserForm_Initialize()
    'Call table_prepare(Me.name)
    export_chosen_option = 0 'FA_export
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, _
                            CloseMode As Integer)
Dim ans
If CloseMode = vbFormControlMenu Then
    Cancel = True
    Me.Height = 628
    Call Me.Hide
End If
End Sub




Private Sub button_export_plateb_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
button_export_plateb.ForeColor = barvaZelenaAirBank
End Sub
Private Sub CommandButton3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton3.ForeColor = barvaZelenaAirBank
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
button_export_plateb.ForeColor = barvaCerna
CommandButton3.ForeColor = barvaCerna
End Sub

Private Sub UserForm_MouseMove3(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton3.ForeColor = barvaCerna
End Sub

