VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Poznamkovy_blok 
   Caption         =   "Pozn·mkov˝ blok"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5400
   OleObjectBlob   =   "Poznamkovy_blok.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Poznamkovy_blok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

' FA vytvorene 17.1.2017 pre zobrazovanie info o klientovi
' podstatna vlastnost tohto UserFormu je ze neni modalni! tzv. vbModeless windows
' (alternativa je vbModal. Modalni okno je takove, ktere neumoznuje delat nic na pozadi (je otravnejsi), napr. MsgBox je vzdy modalni. Tohle si uzivatel posune na kraj a pracuje v Jez. dal)
' to ale znamena ze na pozadi se muzou dit zmeny, proto Modeless userform ma smysl pouzivat pro nedulezite veci, jako treba zobrazeni info o klientovi (ktere ale nic krome zobrazeni info nedela a hlavne nijak nekomunikuje se zbytkem aplikace)

'' TODO este treba dorobit sub changePage(xlRight, xlLeft), ktora posunie na nasledujuceho
''      na samotny userform treba pridat dva buttons (sipky) + label vedla nich na ktorej strane z kolkych sme

Private arrTextMain() As String
Private intNumberOfPages As Integer ' pocet klientov
Private intCurrentPage As Integer

Enum smerOtoceniaStrany
    smerDolava = 1
    smerDoprava = 2
End Enum

Public Property Let numberOfPages(cislo As Integer)
    If cislo > 0 Then
        intNumberOfPages = cislo
        ReDim Preserve arrTextMain(1 To intNumberOfPages)
    Else
        MsgBox "jez-error, poËet str·n pozn·mkovÈho bloku musÌ byù kladn˝ integer.", , "jez-error"
    End If
End Property

Public Property Let currentPage(cislo As Integer)
    If cislo > 0 Then
        intCurrentPage = cislo
    Else
        MsgBox "jez-error, zvolen· strana pozn·mkovÈho bloku musÌ byù kladn˝ integer.", , "jez-error"
    End If
End Property

Public Sub VlozTextNaStranu(cisloStrany As Integer, text As String)
    If intNumberOfPages < 1 Then Exit Sub
    arrTextMain(cisloStrany) = text
End Sub

Private Sub CommandButton_Zavrit_Click()
    Unload Me
End Sub


Private Sub UserForm_Activate()
    Call UkazText
End Sub

Private Sub UserForm_Initialize()
' by default se nastavi array s jedinym prvkom a nahlad na prvu stranu

    If intCurrentPage < 1 Then
        intCurrentPage = 1
        intNumberOfPages = 1
    End If
    'ReDim Preserve arrTextMain(1 To 1)
    Me.numberOfPages = 1
    Call UkazText
    Call NastavHlavnyCaption
    
End Sub

Private Sub UkazText()
    Me.TextBox_Main.text = arrTextMain(intCurrentPage)
    Call UkazCisloStrany
End Sub

Private Sub UkazCisloStrany()
    Me.LabelStrana = intCurrentPage & " z " & intNumberOfPages
End Sub

Private Sub NastavHlavnyCaption()
' nastavi text v zelenom poli hore, na PREHLED RIZENI alebo POZNAMKOVY BLOK
    If Len(Me.TextBox_Main.text) > 5 Then
        Me.Label26.Caption = "P¯ehled ¯ÌzenÌ ke klientovi"
    Else
        Me.Label26.Caption = "Pozn·mky"
    End If
End Sub

Private Sub CommandButtonTurnLeft_Click()
    Call OtocStranu(smerDolava)
End Sub


Private Sub CommandButtonTurnRight_Click()
    Call OtocStranu(smerDoprava)
End Sub



Private Sub OtocStranu(smerOtocenia As smerOtoceniaStrany)

    
    Call UlozAktualniStranu
    
    
    '''   otoc na novu stranu (pokud nejsme na prvni nebo posledni strane)
    If smerOtocenia = smerDolava Then
        If intCurrentPage <= LBound(arrTextMain) Then Exit Sub
        intCurrentPage = intCurrentPage - 1
    ElseIf smerOtocenia = smerDoprava Then
        If intCurrentPage >= UBound(arrTextMain) Then Exit Sub
        intCurrentPage = intCurrentPage + 1
    Else
        MsgBox "jez-error J59866", vbCritical ' should never happen
    End If
    
    Call UkazText
    Call UkazCisloStrany
    
End Sub


Private Sub UlozAktualniStranu()
    arrTextMain(intCurrentPage) = Me.TextBox_Main.text
End Sub

Private Sub TextBox_Main_Change()
    Call UlozAktualniStranu
End Sub

'Private Sub TextBox_Main_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    Call UlozAktualniStranu
'End Sub
