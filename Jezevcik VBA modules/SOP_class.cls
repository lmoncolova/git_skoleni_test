VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SOP_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Created by Filip Antol 20.4.2018 do verze Jez.v.5.5
' tento objekt sprostredkuje Jezevcikovi atributy ku klientovi z tabulky SOP
' objekt Client obsahuje property SOP, ktora vracia odkaz na tento objekt
' vyuzitie: Klient.listview4, Klient_exe.listview4, button Upravit Udaje NR, button Autorizace
' objekt vyzaduje, aby byl nainicializovan klient
' vetsina kodu je repetitivni a byla generovana pomoci funkce generate_SOP_columns
' v Class_initialize sa pouziva c.id_person (nie globalna id_person)





'''   Promenne ktere se nacitaji z tabulky SOP
Private pSKP_SOP As Long
Private pAMT_SOP1 As Double
Private pDATE_SOP1 As Date
Private pSOP_ADDITIONAL_FEE As Double

Private pAMT_ADDITIONAL_SOP1 As Double
Private pDATE_ADDITIONAL_SOP1 As Date
Private pAMT_RETURN_SOP1  As Double
Private pDATE_INFO_RETURN_SOP1 As Date
Private pAMT_RETURN_REAL_SOP1 As Double
Private pDATE_RETURN_REAL_SOP1 As Date
Private pAMT_SOP1_COURT As Double

Private pAMT_SOP2 As Double
Private pDATE_SOP2 As Date
Private pAMT_ADDITIONAL_SOP2 As Double
Private pDATE_ADDITIONAL_SOP2
Private pAMT_RETURN_SOP2 As Double
Private pDATE_INFO_RETURN_SOP2
Private pAMT_RETURN_REAL_SOP2 As Double
Private pDATE_RETURN_REAL_SOP2
Private pAMT_SOP2_COURT As Double

Private pSOP_ADDITIONAL_PAID As Double
Private pTRAVEL_EXPENSES As Double
Private pDTIME_MODIFICATION As Date
Private pEMPLOYEE As String



'''   Ostatni promenne

' Autorizace - je 6 polozek, u ktorych je potrebna autorizace (vzdy za nejakych podminek)
'            - jsou to SOP1, preplatek SOP1, doplatek SOP1 + to same pro SOP2
' (btw, SOP1 sa tyka okresniho soudu. Vynimocne se airbank odvola proti rozsudku, jede to na krajskej soud a pak se resi SOP2)

' Private pblnPovolitAutorizaci As Boolean ' extra duvody mimo autorizaci
Private pcolSloupeckyPodbarvit As Collection

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''   PUBLIC Properties pro autorizaci
'''   jenom properties, nemaji vlastni promenne (je to zamerne, kvuli ortogonalizaci kodu)
'''   Pro SOP2 staci zkopirovat SOP1 a nahradit "SOP1" za "SOP2"

Public Property Get Auth_SOP1() As stavAutorizace
    
    If Me.DATE_SOP1 <> #12:00:00 AM# Then
        Auth_SOP1 = stavAuth_Autorizovano
    ElseIf Me.AMT_SOP1 <> cfg.cstDEFAULT_AMT_SOP_VALUE Then
        Auth_SOP1 = stavAuth_Ceka_na_autorizaci
    Else
        Auth_SOP1 = stavAuth_None
    End If
    
End Property

Public Property Get Auth_Additional_SOP1() As stavAutorizace
    
    If Me.DATE_ADDITIONAL_SOP1 <> #12:00:00 AM# Then
        Auth_Additional_SOP1 = stavAuth_Autorizovano
    ElseIf Me.AMT_ADDITIONAL_SOP1 <> cfg.cstDEFAULT_AMT_SOP_VALUE Then
        Auth_Additional_SOP1 = stavAuth_Ceka_na_autorizaci
    Else
        Auth_Additional_SOP1 = stavAuth_None
    End If
    
End Property

Public Property Get Auth_Return_SOP1() As stavAutorizace
    
    If Me.AMT_RETURN_REAL_SOP1 = Me.AMT_RETURN_SOP1 And Me.AMT_RETURN_REAL_SOP1 <> cfg.cstDEFAULT_AMT_SOP_VALUE Then
            Auth_Return_SOP1 = stavAuth_Autorizovano
    ElseIf (Me.AMT_RETURN_REAL_SOP1 <> Me.AMT_RETURN_SOP1 And Me.AMT_RETURN_REAL_SOP1 <> cfg.cstDEFAULT_AMT_SOP_VALUE) _
        Or (Me.DATE_RETURN_REAL_SOP1 = #12:00:00 AM# And pDATE_INFO_RETURN_SOP1 <> #12:00:00 AM#) Then
            Auth_Return_SOP1 = stavAuth_Ceka_na_autorizaci
    Else
            Auth_Return_SOP1 = stavAuth_None
    End If
    
End Property


'''   odsud SOP2 (kopie SOP1)

Public Property Get Auth_SOP2() As stavAutorizace
    
    If Me.DATE_SOP2 <> #12:00:00 AM# Then
        Auth_SOP2 = stavAuth_Autorizovano
    ElseIf Me.AMT_SOP2 <> cfg.cstDEFAULT_AMT_SOP_VALUE Then
        Auth_SOP2 = stavAuth_Ceka_na_autorizaci
    Else
        Auth_SOP2 = stavAuth_None
    End If
    
End Property

Public Property Get Auth_Additional_SOP2() As stavAutorizace
    
    If Me.DATE_ADDITIONAL_SOP2 <> #12:00:00 AM# Then
        Auth_Additional_SOP2 = stavAuth_Autorizovano
    ElseIf Me.AMT_ADDITIONAL_SOP2 <> cfg.cstDEFAULT_AMT_SOP_VALUE Then
        Auth_Additional_SOP2 = stavAuth_Ceka_na_autorizaci
    Else
        Auth_Additional_SOP2 = stavAuth_None
    End If
    
End Property

Public Property Get Auth_Return_SOP2() As stavAutorizace
    
    If Me.AMT_RETURN_REAL_SOP2 = Me.AMT_RETURN_SOP2 And Me.AMT_RETURN_REAL_SOP2 <> cfg.cstDEFAULT_AMT_SOP_VALUE Then
            Auth_Return_SOP2 = stavAuth_Autorizovano
    ElseIf (Me.AMT_RETURN_REAL_SOP2 <> Me.AMT_RETURN_SOP2 And Me.AMT_RETURN_REAL_SOP2 <> cfg.cstDEFAULT_AMT_SOP_VALUE) _
        Or (Me.DATE_RETURN_REAL_SOP2 = #12:00:00 AM# And pDATE_INFO_RETURN_SOP2 <> #12:00:00 AM#) Then
            Auth_Return_SOP2 = stavAuth_Ceka_na_autorizaci
    Else
            Auth_Return_SOP2 = stavAuth_None
    End If
    
End Property


Public Property Get blnPovolitAutorizaci() As Boolean
    ' Je potreba neco autorizovat?
    
    '
    If (Me.Auth_Additional_SOP1 = stavAuth_Ceka_na_autorizaci) Or _
        (Me.Auth_Additional_SOP2 = stavAuth_Ceka_na_autorizaci) Or _
        (Me.Auth_Return_SOP1 = stavAuth_Ceka_na_autorizaci) Or _
        (Me.Auth_Return_SOP2 = stavAuth_Ceka_na_autorizaci) Or _
        (Me.Auth_SOP1 = stavAuth_Ceka_na_autorizaci) Or _
        (Me.Auth_SOP2 = stavAuth_Ceka_na_autorizaci) Then
       
             blnPovolitAutorizaci = True
             
    End If
    
    
End Property


Public Property Get Pausal() As Double
    ' alias pro pausal
   Pausal = pSOP_ADDITIONAL_FEE
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''      Properties auto-generated via generate_SOP_columns:

Public Property Get SKP_SOP() As Long
    SKP_SOP = pSKP_SOP
End Property
Public Property Get AMT_SOP1() As Double
    AMT_SOP1 = pAMT_SOP1
End Property
Public Property Get DATE_SOP1() As Date
    DATE_SOP1 = pDATE_SOP1
End Property
Public Property Get SOP_ADDITIONAL_FEE() As Double
   SOP_ADDITIONAL_FEE = pSOP_ADDITIONAL_FEE
End Property
Public Property Get AMT_ADDITIONAL_SOP1() As Double
   AMT_ADDITIONAL_SOP1 = pAMT_ADDITIONAL_SOP1
End Property
Public Property Get DATE_ADDITIONAL_SOP1() As Date
   DATE_ADDITIONAL_SOP1 = pDATE_ADDITIONAL_SOP1
End Property
Public Property Get AMT_RETURN_SOP1() As Double
   AMT_RETURN_SOP1 = pAMT_RETURN_SOP1
End Property
Public Property Get DATE_INFO_RETURN_SOP1() As Date
   DATE_INFO_RETURN_SOP1 = pDATE_INFO_RETURN_SOP1
End Property
Public Property Get AMT_RETURN_REAL_SOP1() As Double
   AMT_RETURN_REAL_SOP1 = pAMT_RETURN_REAL_SOP1
End Property
Public Property Get DATE_RETURN_REAL_SOP1() As Date
   DATE_RETURN_REAL_SOP1 = pDATE_RETURN_REAL_SOP1
End Property
Public Property Get AMT_SOP1_COURT() As Double
   AMT_SOP1_COURT = pAMT_SOP1_COURT
End Property
Public Property Get AMT_SOP2() As Double
   AMT_SOP2 = pAMT_SOP2
End Property
Public Property Get DATE_SOP2() As Date
   DATE_SOP2 = pDATE_SOP2
End Property
Public Property Get AMT_ADDITIONAL_SOP2() As Double
   AMT_ADDITIONAL_SOP2 = pAMT_ADDITIONAL_SOP2
End Property
Public Property Get DATE_ADDITIONAL_SOP2() As Date
   DATE_ADDITIONAL_SOP2 = pDATE_ADDITIONAL_SOP2
End Property
Public Property Get AMT_RETURN_SOP2() As Double
   AMT_RETURN_SOP2 = pAMT_RETURN_SOP2
End Property
Public Property Get DATE_INFO_RETURN_SOP2() As Date
   DATE_INFO_RETURN_SOP2 = pDATE_INFO_RETURN_SOP2
End Property
Public Property Get AMT_RETURN_REAL_SOP2() As Double
   AMT_RETURN_REAL_SOP2 = pAMT_RETURN_REAL_SOP2
End Property
Public Property Get DATE_RETURN_REAL_SOP2() As Date
   DATE_RETURN_REAL_SOP2 = pDATE_RETURN_REAL_SOP2
End Property
Public Property Get AMT_SOP2_COURT() As Double
   AMT_SOP2_COURT = pAMT_SOP2_COURT
End Property
Public Property Get SOP_ADDITIONAL_PAID() As Double
   SOP_ADDITIONAL_PAID = pSOP_ADDITIONAL_PAID
End Property
Public Property Get TRAVEL_EXPENSES() As Double
   TRAVEL_EXPENSES = pTRAVEL_EXPENSES
End Property
Public Property Get DTIME_MODIFICATION() As Date
   DTIME_MODIFICATION = pDTIME_MODIFICATION
End Property
Public Property Get EMPLOYEE() As String
   EMPLOYEE = pEMPLOYEE
End Property


''''''      END Properties auto-generated via generate_SOP_columns:

'Public Property Get XXX() As YYY
'    XXX = pXXX
'End Property
'Public Property Get XXX() As YYY
'    XXX = pXXX
'End Property





Public Sub generate_SOP_columns()
' generator repetitivniho kodu
' tohle chces vyuzit

    Dim key As Variant
    Dim i As Integer
    i = 0
    
    ' Debug.Print vbCrLf
    Debug.Print "    '''   auto-generated via generate_SOP_columns:"
    
    For Each key In cfg.SlovnikUzivatelskychNazvu_SOP.keys
        
        
        If key = "SKP_SOP" Then GoTo Next_iteration
        i = i + 1
        ''' generate sop columns
        If 1 = 0 Then
            Debug.Print code_generator(key)
        End If
        
        ''' generate_get_properties
        If 1 = 0 Then
            Debug.Print "Public Property Get " & key & "() As double" & vbCrLf & _
                        "   " & key & " = p" & key & vbCrLf & _
                        "End Property"
        End If
        
        ''' generate_get_item
        If 1 = 0 Then
            Debug.Print "Case " & i & ": Item = Me." & key ' & vbCrLf
        End If
        
        ''' generate listview order
   
        If 1 = 1 Then
            Debug.Print "    IF (Empty_string_for_default_value_SOPclass (Me." & key & ") <> """") THEN " & vbCrLf & _
                        "        Set itmX = .ListItems.Add()" & vbCrLf & _
                        "            itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP(""" & key & """)" & vbCrLf & _
                        "            itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me." & key & ")" & vbCrLf & _
                        "    END IF"
            
            If i Mod 5 = 0 Then Debug.Print vbCrLf ' prazdny radek kazdych 5 iteraci
        End If
        
        ''' nastaveni defaultnich hodnot
        If 1 = 0 Then
            
            If InStr(key, "DATE") Then
                Debug.Print "p" & key & " = #12:00:00 AM#"
            ElseIf InStr(key, "AMT") Then
                Debug.Print "p" & key & " = cfg.cstDEFAULT_AMT_SOP_VALUE"
            Else
                Debug.Print "p" & key & " = cfg.cstDEFAULT_AMT_SOP_VALUE   ?????"
            End If
            
        End If
        
        
Next_iteration:
    Next key
    
    Debug.Print "    '''   END auto-generated via generate_SOP_columns"
    
End Sub

Public Sub re_initialize()
    Call Class_Initialize
End Sub

Private Sub Class_Initialize()
    Dim sql As String
    Dim rst As ADODB.Recordset
    
    assertion c.id_person > 0
    
    
    sql = "select SKP_SOP, id_person, AMT_SOP1, DATE_SOP1, sop_additional_fee, " & _
            "AMT_ADDITIONAL_SOP1, DATE_ADDITIONAL_SOP1, AMT_RETURN_SOP1, DATE_INFO_RETURN_SOP1, AMT_RETURN_REAL_SOP1, " & _
            "DATE_RETURN_REAL_SOP1, AMT_SOP1_COURT, AMT_SOP2, DATE_SOP2, AMT_ADDITIONAL_SOP2, " & _
            "DATE_ADDITIONAL_SOP2, AMT_RETURN_SOP2, DATE_INFO_RETURN_SOP2, AMT_RETURN_REAL_SOP2, DATE_RETURN_REAL_SOP2, " & _
            "AMT_SOP2_COURT, sop_additional_paid , travel_expenses , DTIME_MODIFICATION, EMPLOYEE " & _
          " FROM " & schema & ".sop where id_person = " & c.id_person
    Set rst = New ADODB.Recordset
    rst.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
        If Not rst.EOF Then
            Call assertion(rst.RecordCount = 1, "Chyba v inicializaci SOP_class (v1)")
            Call assertion(CStr(rst![id_person]) = CStr(id_person), "Chyba v inicializaci SOP_class (v2)")
            Call assertion(CStr(c.id_person) = CStr(id_person), "Chyba v inicializaci SOP_class (v3)")
            
'            pSKP_SOP = nvl(rst![Skp_sop])
'            pAMT_SOP1 = nvl(rst![Amt_sop1])
'            pDATE_SOP1 = nvl(rst![Date_sop1])
            
            '''   auto-generated via generate_SOP_columns:
            pSKP_SOP = nvl(rst![SKP_SOP])
            'pID_PERSON = nvl(rst![id_person])
            pAMT_SOP1 = nvl(rst![AMT_SOP1], cfg.cstDEFAULT_AMT_SOP_VALUE)
            pDATE_SOP1 = nvl(rst![DATE_SOP1], #12:00:00 AM#)
            pSOP_ADDITIONAL_FEE = nvl(rst![SOP_ADDITIONAL_FEE], cfg.cstDEFAULT_AMT_SOP_VALUE)
            pAMT_ADDITIONAL_SOP1 = nvl(rst![AMT_ADDITIONAL_SOP1], cfg.cstDEFAULT_AMT_SOP_VALUE)
            pDATE_ADDITIONAL_SOP1 = nvl(rst![DATE_ADDITIONAL_SOP1], #12:00:00 AM#)
            pAMT_RETURN_SOP1 = nvl(rst![AMT_RETURN_SOP1], cfg.cstDEFAULT_AMT_SOP_VALUE)
            pDATE_INFO_RETURN_SOP1 = nvl(rst![DATE_INFO_RETURN_SOP1], #12:00:00 AM#)
            pAMT_RETURN_REAL_SOP1 = nvl(rst![AMT_RETURN_REAL_SOP1], cfg.cstDEFAULT_AMT_SOP_VALUE)
            pDATE_RETURN_REAL_SOP1 = nvl(rst![DATE_RETURN_REAL_SOP1], #12:00:00 AM#)
            pAMT_SOP1_COURT = nvl(rst![AMT_SOP1_COURT], cfg.cstDEFAULT_AMT_SOP_VALUE)
            pAMT_SOP2 = nvl(rst![AMT_SOP2], cfg.cstDEFAULT_AMT_SOP_VALUE)
            pDATE_SOP2 = nvl(rst![DATE_SOP2], #12:00:00 AM#)
            pAMT_ADDITIONAL_SOP2 = nvl(rst![AMT_ADDITIONAL_SOP2], cfg.cstDEFAULT_AMT_SOP_VALUE)
            pDATE_ADDITIONAL_SOP2 = nvl(rst![DATE_ADDITIONAL_SOP2], #12:00:00 AM#)
            pAMT_RETURN_SOP2 = nvl(rst![AMT_RETURN_SOP2], cfg.cstDEFAULT_AMT_SOP_VALUE)
            pDATE_INFO_RETURN_SOP2 = nvl(rst![DATE_INFO_RETURN_SOP2], #12:00:00 AM#)
            pAMT_RETURN_REAL_SOP2 = nvl(rst![AMT_RETURN_REAL_SOP2], cfg.cstDEFAULT_AMT_SOP_VALUE)
            pDATE_RETURN_REAL_SOP2 = nvl(rst![DATE_RETURN_REAL_SOP2], #12:00:00 AM#)
            pAMT_SOP2_COURT = nvl(rst![AMT_SOP2_COURT], cfg.cstDEFAULT_AMT_SOP_VALUE)
            pSOP_ADDITIONAL_PAID = nvl(rst![SOP_ADDITIONAL_PAID], cfg.cstDEFAULT_AMT_SOP_VALUE)
            pTRAVEL_EXPENSES = nvl(rst![TRAVEL_EXPENSES], cfg.cstDEFAULT_AMT_SOP_VALUE)
            pDTIME_MODIFICATION = nvl(rst![DTIME_MODIFICATION], #12:00:00 AM#)
            pEMPLOYEE = nvl(rst![EMPLOYEE])
            '''   END auto-generated via generate_SOP_columns:
            
        Else
            Call NastavDefaultniHodnoty
        End If
        
    rst.Close
    
    
    
    Call NajdiSloupeckyProPodbarveni
    
End Sub



Private Sub NastavDefaultniHodnoty()
    ' vsechny hodnty je potreba nastavit na nasi defaultni hodnotu (ne 0)

    '''   auto-generated via generate_SOP_columns:
        pAMT_SOP1 = cfg.cstDEFAULT_AMT_SOP_VALUE
        pDATE_SOP1 = #12:00:00 AM#
        pSOP_ADDITIONAL_FEE = cfg.cstDEFAULT_AMT_SOP_VALUE
        pAMT_ADDITIONAL_SOP1 = cfg.cstDEFAULT_AMT_SOP_VALUE
        pDATE_ADDITIONAL_SOP1 = #12:00:00 AM#
        pAMT_RETURN_SOP1 = cfg.cstDEFAULT_AMT_SOP_VALUE
        pDATE_INFO_RETURN_SOP1 = #12:00:00 AM#
        pAMT_RETURN_REAL_SOP1 = cfg.cstDEFAULT_AMT_SOP_VALUE
        pDATE_RETURN_REAL_SOP1 = #12:00:00 AM#
        pAMT_SOP1_COURT = cfg.cstDEFAULT_AMT_SOP_VALUE
        pAMT_SOP2 = cfg.cstDEFAULT_AMT_SOP_VALUE
        pDATE_SOP2 = #12:00:00 AM#
        pAMT_ADDITIONAL_SOP2 = cfg.cstDEFAULT_AMT_SOP_VALUE
        pDATE_ADDITIONAL_SOP2 = #12:00:00 AM#
        pAMT_RETURN_SOP2 = cfg.cstDEFAULT_AMT_SOP_VALUE
        pDATE_INFO_RETURN_SOP2 = #12:00:00 AM#
        pAMT_RETURN_REAL_SOP2 = cfg.cstDEFAULT_AMT_SOP_VALUE
        pDATE_RETURN_REAL_SOP2 = #12:00:00 AM#
        pAMT_SOP2_COURT = cfg.cstDEFAULT_AMT_SOP_VALUE
        pSOP_ADDITIONAL_PAID = cfg.cstDEFAULT_AMT_SOP_VALUE
        pTRAVEL_EXPENSES = cfg.cstDEFAULT_AMT_SOP_VALUE
        pDTIME_MODIFICATION = #12:00:00 AM#
        pEMPLOYEE = ""
    '''   END auto-generated via generate_SOP_columns


End Sub



Private Sub NajdiSloupeckyProPodbarveni()
    '''''
    '''''   Do kolekce pcolSloupeckyPodbarvit si ulozime radky Listview, ktere chceme podbarvit
    '''''   vetsinou ty, ktere cekaji na autorizaci
    
    
    Set pcolSloupeckyPodbarvit = New Collection
    
    
    '''   SOPy
    
    If Me.Auth_SOP1 = stavAuth_Ceka_na_autorizaci Then
        pcolSloupeckyPodbarvit.Add "AMT_SOP1"
        pcolSloupeckyPodbarvit.Add "DATE_SOP1"
    End If
    
    If Me.Auth_SOP2 = stavAuth_Ceka_na_autorizaci Then                'If pDATE_SOP2 = #12:00:00 AM# And pAMT_SOP2 > 0 Then
        pcolSloupeckyPodbarvit.Add "AMT_SOP2"
        pcolSloupeckyPodbarvit.Add "DATE_SOP2"
    End If
    
'''    If pAMT_ADDITIONAL_SOP1 = 0 And pDATE_SOP1 > #12:00:00 AM# Then ' tohle je konina asi?
'''        pblnPovolitAutorizaci = True
'''        pcolSloupeckyPodbarvit.Add "AMT_SOP1"
'''    End If

    
    
    '''   doplatky
    
    
    If Me.Auth_Additional_SOP1 = stavAuth_Ceka_na_autorizaci Then     'If pDATE_ADDITIONAL_SOP1 = #12:00:00 AM# And pAMT_ADDITIONAL_SOP1 <> 0 Then
        ' Debug.Assert False
        ' pcolSloupeckyPodbarvit.Add "DATE_SOP1" ' asi konina
        pcolSloupeckyPodbarvit.Add "DATE_ADDITIONAL_SOP1"
        pcolSloupeckyPodbarvit.Add "AMT_ADDITIONAL_SOP1"
    End If
    
    If Me.Auth_Additional_SOP2 = stavAuth_Ceka_na_autorizaci Then     'If pDATE_ADDITIONAL_SOP2 = #12:00:00 AM# And pAMT_ADDITIONAL_SOP2 > 0 Then
        ' Debug.Assert False
        ' pcolSloupeckyPodbarvit.Add "DATE_SOP2" ' asi konina
        pcolSloupeckyPodbarvit.Add "DATE_ADDITIONAL_SOP2"
        pcolSloupeckyPodbarvit.Add "AMT_ADDITIONAL_SOP2"
    End If


    '''  preplatky
    
    If Me.Auth_Return_SOP1 = stavAuth_Ceka_na_autorizaci Then           '    If pAMT_RETURN_REAL_SOP1 = 0 And pAMT_RETURN_SOP1 <> 0 Then
        pcolSloupeckyPodbarvit.Add "AMT_RETURN_SOP1"
        pcolSloupeckyPodbarvit.Add "DATE_INFO_RETURN_SOP1"
        pcolSloupeckyPodbarvit.Add "AMT_RETURN_REAL_SOP1"
        pcolSloupeckyPodbarvit.Add "DATE_RETURN_REAL_SOP1"
    End If
    
    If Me.Auth_Return_SOP2 = stavAuth_Ceka_na_autorizaci Then           '     If pAMT_RETURN_REAL_SOP2 = 0 And pAMT_RETURN_SOP2 <> 0 Then
        pcolSloupeckyPodbarvit.Add "AMT_RETURN_SOP2"
        pcolSloupeckyPodbarvit.Add "DATE_INFO_RETURN_SOP2"
        pcolSloupeckyPodbarvit.Add "AMT_RETURN_REAL_SOP2"
        pcolSloupeckyPodbarvit.Add "DATE_RETURN_REAL_SOP2"
    End If
    
End Sub

Private Function Empty_string_for_default_value_SOPclass(value As Variant) As Variant
    ' wrapper k funkci Empty_string_for_default_value pro SOP_class
    ' jedine co funkce dela je, ze automaticky pridava volitelny argument
    
    Empty_string_for_default_value_SOPclass = Empty_string_for_default_value(value, default_numeric_value:=cfg.cstDEFAULT_AMT_SOP_VALUE)
End Function

Public Sub Vytvor_listview_prehled_SOPu(ByRef LV As ListView)
    ' pro plneni listview na karte Klient a Klient_exe se zavola tahle funkce
    ' VYBER atributu je urcen tady
    ' PORADI atributu je urceno tady
    ' pokud by se to melo lisit na ruznych kartach, tak pridej do funkce optional parametr
    ' defaultni hodnoty se vynechavaji - nemusi se pak scrollovat v listview (SOP2 vyuzivame jenom zridkakdy)


    Dim itmX As ListItem
    
    LV.ListItems.Clear
    LV.ColumnHeaders.Clear
    
    With LV
        
        Call .ColumnHeaders.Add(, , "Údaj", 152)
        Call .ColumnHeaders.Add(, , "Hodnota", 100)
        
                
        '''   auto-generated via generate_SOP_columns:
            
            If (Empty_string_for_default_value_SOPclass(Me.AMT_SOP1) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("AMT_SOP1")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.AMT_SOP1)
            End If
            If (Empty_string_for_default_value_SOPclass(Me.DATE_SOP1) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("DATE_SOP1")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.DATE_SOP1)
            End If
            If (Empty_string_for_default_value_SOPclass(Me.SOP_ADDITIONAL_FEE) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("SOP_ADDITIONAL_FEE")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.SOP_ADDITIONAL_FEE)
            End If
            If (Empty_string_for_default_value_SOPclass(Me.AMT_ADDITIONAL_SOP1) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("AMT_ADDITIONAL_SOP1")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.AMT_ADDITIONAL_SOP1)
            End If
            If (Empty_string_for_default_value_SOPclass(Me.DATE_ADDITIONAL_SOP1) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("DATE_ADDITIONAL_SOP1")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.DATE_ADDITIONAL_SOP1)
            End If
        
        
            If (Empty_string_for_default_value_SOPclass(Me.AMT_RETURN_SOP1) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("AMT_RETURN_SOP1")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.AMT_RETURN_SOP1)
            End If
            If (Empty_string_for_default_value_SOPclass(Me.DATE_INFO_RETURN_SOP1) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("DATE_INFO_RETURN_SOP1")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.DATE_INFO_RETURN_SOP1)
            End If
            If (Empty_string_for_default_value_SOPclass(Me.AMT_RETURN_REAL_SOP1) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("AMT_RETURN_REAL_SOP1")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.AMT_RETURN_REAL_SOP1)
            End If
            If (Empty_string_for_default_value_SOPclass(Me.DATE_RETURN_REAL_SOP1) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("DATE_RETURN_REAL_SOP1")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.DATE_RETURN_REAL_SOP1)
            End If
            If (Empty_string_for_default_value_SOPclass(Me.AMT_SOP1_COURT) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("AMT_SOP1_COURT")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.AMT_SOP1_COURT)
            End If
        
        
            If (Empty_string_for_default_value_SOPclass(Me.AMT_SOP2) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("AMT_SOP2")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.AMT_SOP2)
            End If
            If (Empty_string_for_default_value_SOPclass(Me.DATE_SOP2) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("DATE_SOP2")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.DATE_SOP2)
            End If
            If (Empty_string_for_default_value_SOPclass(Me.AMT_ADDITIONAL_SOP2) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("AMT_ADDITIONAL_SOP2")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.AMT_ADDITIONAL_SOP2)
            End If
            If (Empty_string_for_default_value_SOPclass(Me.DATE_ADDITIONAL_SOP2) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("DATE_ADDITIONAL_SOP2")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.DATE_ADDITIONAL_SOP2)
            End If
            If (Empty_string_for_default_value_SOPclass(Me.AMT_RETURN_SOP2) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("AMT_RETURN_SOP2")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.AMT_RETURN_SOP2)
            End If
        
        
            If (Empty_string_for_default_value_SOPclass(Me.DATE_INFO_RETURN_SOP2) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("DATE_INFO_RETURN_SOP2")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.DATE_INFO_RETURN_SOP2)
            End If
            If (Empty_string_for_default_value_SOPclass(Me.AMT_RETURN_REAL_SOP2) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("AMT_RETURN_REAL_SOP2")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.AMT_RETURN_REAL_SOP2)
            End If
            If (Empty_string_for_default_value_SOPclass(Me.DATE_RETURN_REAL_SOP2) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("DATE_RETURN_REAL_SOP2")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.DATE_RETURN_REAL_SOP2)
            End If
            If (Empty_string_for_default_value_SOPclass(Me.AMT_SOP2_COURT) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("AMT_SOP2_COURT")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.AMT_SOP2_COURT)
            End If
            If (Empty_string_for_default_value_SOPclass(Me.SOP_ADDITIONAL_PAID) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("SOP_ADDITIONAL_PAID")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.SOP_ADDITIONAL_PAID)
            End If
        
        
            If (Empty_string_for_default_value_SOPclass(Me.TRAVEL_EXPENSES) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("TRAVEL_EXPENSES")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.TRAVEL_EXPENSES)
            End If
            If (Empty_string_for_default_value_SOPclass(Me.DTIME_MODIFICATION) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("DTIME_MODIFICATION")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.DTIME_MODIFICATION)
            End If
            If (Empty_string_for_default_value_SOPclass(Me.EMPLOYEE) <> "") Then
                Set itmX = .ListItems.Add()
                    itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("EMPLOYEE")
                    itmX.SubItems(1) = Empty_string_for_default_value_SOPclass(Me.EMPLOYEE)
            End If
        '''   END auto-generated via generate_SOP_columns

        
        ''    Set itmX = .ListItems.Add()
        ''        itmX.text = cfg.SlovnikUzivatelskychNazvu_SOP("XXXXXX")
        ''        itmX.SubItems(1) = pXXXXXX

    End With
    
    Call ZabarviRadky(LV, barvaRudaSvetla, pcolSloupeckyPodbarvit)
    
    LV.Refresh

End Sub


Private Sub ZabarviRadky(ByRef LV As ListView, ByVal barvicka As barva, radky_k_zafarbeni As Collection)
    ' funkce obarvi vsechny radky v kolekci radky_k_zafarbeni
    ' funkce neni obecna kvuli cfg.SlovnikUzivatelskychNazvu_SOP

    Dim atribut As Variant
    
    For Each atribut In radky_k_zafarbeni
        Call ZabarviRadekListview(LV, cfg.SlovnikUzivatelskychNazvu_SOP(atribut), barvicka)
        ' Debug.Print atribut
    Next atribut
    
End Sub

