VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JezUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'RISK-983
'Objekt uživatelù Jezevèíka, standardnì se inicializuje pøi pøihlášení Main.CommandButton1_Click

Private Const strEmailPravniVymahani As String = "pravnivymahani@airbank.cz"
Private Const strNazevOddeleni As String = "Právní vymáhání"

Private pStrUserName As String
Private pStrUserRole As String
Private pStrCivilName As String
Private pStrNameVocative As String
Private pStrEmailAddress As String
Private pStrPhoneNumber As String
Private pStrSignature As String
Private pblnBrigadnik As Boolean
Private pblnZkracenyPodpis As Boolean


Public Property Get StrUserName() As String
    StrUserName = pStrUserName
End Property
Public Property Get StrNameVocative() As String
    StrNameVocative = pStrNameVocative
End Property
Public Property Get StrUserRole() As String
    StrUserRole = pStrUserRole
End Property
Public Property Get StrCivilName() As String
    StrCivilName = pStrCivilName
End Property
Public Property Get StrEmailAddress() As String
    StrEmailAddress = pStrEmailAddress
End Property
Public Property Get StrPhoneNumber() As String
    StrPhoneNumber = pStrPhoneNumber
End Property
Public Property Get strSignature() As String
    If pStrSignature = "" Then pStrSignature = GenerujPodpis
    strSignature = pStrSignature
End Property
Public Property Get blnBrigadnik() As Boolean
    blnBrigadnik = pblnBrigadnik
End Property
Public Property Get blnZkracenyPodpis() As Boolean
    pblnZkracenyPodpis = blnZkracenyPodpis
End Property
 
Public Sub NactiUzivatele(ByVal username As String)
' sql dotaz do seznamu uzivatelu Jezevèíka

Dim sql_user As String
Dim rst_user As ADODB.Recordset

    
    sql_user = "SELECT USERNAME, USER_ROLE, NAME, EMAIL_ADDRESS, PHONE_NUMBER, NAME_VOCATIVE, FLAG_BRIGADNIK, FLAG_ZKRACENY_PODPIS " & _
                "FROM " & schema & ".jez_users WHERE flag_valid = 'Y' AND USERNAME = '" & UCase(username) & "'"
    
    Set rst_user = New ADODB.Recordset
    rst_user.Open sql_user, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
        If rst_user.RecordCount = 1 Then
            
            If Not IsNull(rst_user![username]) Then pStrUserName = rst_user![username]
            If Not IsNull(rst_user![USER_ROLE]) Then pStrUserRole = rst_user![USER_ROLE]
            If Not IsNull(rst_user![name]) Then pStrCivilName = rst_user![name]
            If Not IsNull(rst_user![EMAIL_ADDRESS]) Then pStrEmailAddress = LCase(rst_user![EMAIL_ADDRESS]) ' lowercase
            If Not IsNull(rst_user![PHONE_NUMBER]) Then pStrPhoneNumber = rst_user![PHONE_NUMBER]
            If Not IsNull(rst_user![NAME_VOCATIVE]) Then pStrNameVocative = rst_user![NAME_VOCATIVE]
            If Not IsNull(rst_user![FLAG_BRIGADNIK]) Then If rst_user![FLAG_BRIGADNIK] = "Y" Then pblnBrigadnik = True
            If Not IsNull(rst_user![FLAG_ZKRACENY_PODPIS]) Then If rst_user![FLAG_ZKRACENY_PODPIS] = "Y" Then pblnZkracenyPodpis = True
            
        Else
            
            MsgBox "Nepovedlo se naèíst uživatele. Napiš analytikùm, protože tì neevidují. Po odkliknutí se Jezevèík zavøe."
            Debug.Assert 1 = 0
            End
        
        End If
        
    rst_user.Close

End Sub

Private Function GenerujPodpis() As String
    ' k danemu uzivateli vytvori podpis - asi 4 radky textu, ktere pravnici a vymahaci vkladaji do zaveru sablon
    ' pokud to nepujde udelat skrz info v objektu, mozna by stalo za to to natvrdo napsat pro kazdeho z nich.
    
    '' priklad:
    '' Za Air Bank a.s. v povìøení
    '' Lenka Zapletalíková / Právní vymáhání
    '' Air Bank a.s. / èlen skupiny PPF
    '' P.O.Box 625 / 661 25 Brno
    '' telefon +420 547 134 714
    '' e -mail: lenka.zapletalikova@ airbank.cz
    
    
    
    
    '    '''   puvodni jednoducha verze
    '    GenerujPodpis = "Za Air Bank a.s. v povìøení" & vbCrLf & _
    '                    pStrCivilName & " / " & strNazevOddeleni & vbCrLf & _
    '                    "Air Bank a.s. / èlen skupiny PPF" & vbCrLf & _
    '                    "P.O.Box 625 / 661 25 Brno" & vbCrLf & _
    '                    "telefon: " & pStrPhoneNumber & vbCrLf & _
    '                    "e-mail: " & pStrEmailAddress
    
    
    ' XX este treba vyriesit ze u brigadnikov sa nema ICH email ukazovat, ale obecny

    GenerujPodpis = "Za Air Bank a.s. v povìøení" & vbCrLf & _
                    pStrCivilName & " / " & strNazevOddeleni & vbCrLf & _
                    "Air Bank a.s. / èlen skupiny PPF" & vbCrLf & _
                    "P.O.Box 625 / 661 25 Brno" & vbCrLf
    
    If Len(Trim(pStrPhoneNumber)) > 0 And Not pblnBrigadnik And Not pblnZkracenyPodpis Then
        GenerujPodpis = GenerujPodpis & _
                        "telefon: " & pStrPhoneNumber & vbCrLf
    End If
    
    
    If pblnBrigadnik Or pblnZkracenyPodpis Then
        GenerujPodpis = GenerujPodpis & _
                        "e-mail: " & strEmailPravniVymahani
    Else
        GenerujPodpis = GenerujPodpis & _
                        "e-mail: " & pStrEmailAddress
    End If

End Function

