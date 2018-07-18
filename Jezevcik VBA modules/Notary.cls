VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Notary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' objekt bude odpovedat jednemu notarovi z hlavnej tabulky (ciselniku notarov) NOTARY

Private plngSkp_notary As Long
Private pstrAddress As String
Private pstrName As String
Private pStrPhoneNumber As String
Private pstrIdDS As String

'Public Property Get XXXYYY()
'    XXXYYY = pXXXYYY
'End Property

Public Property Get lngSkp_notary()
    lngSkp_notary = plngSkp_notary
End Property

Public Property Get strName()
    strName = pstrName
End Property

Public Property Get strAddress()
    strAddress = pstrAddress
End Property


Public Property Get StrPhoneNumber()
    StrPhoneNumber = pStrPhoneNumber
End Property

Public Property Get strIdDS()
    strIdDS = pstrIdDS
End Property

Public Sub NacitajNotara(ByVal skp_notary As Long)
' sql dotaz do ciselniku notarov a naplnenie premennych

Dim sql_notar As String
Dim rst_notar As ADODB.Recordset


plngSkp_notary = skp_notary

sql_notar = "SELECT n.skp_notary, n.name, n.address, n.phone_number, n.id_ds " & _
            "FROM " & schema & ".notary n WHERE n.flag_deleted != 'Y' AND n.skp_notary = " & skp_notary

Set rst_notar = New ADODB.Recordset
rst_notar.Open sql_notar, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
    If rst_notar.RecordCount = 1 Then
        
        If Not IsNull(rst_notar![name]) Then pstrName = rst_notar![name]
        If Not IsNull(rst_notar![address]) Then pstrAddress = rst_notar![address]
        If Not IsNull(rst_notar![PHONE_NUMBER]) Then pStrPhoneNumber = rst_notar![PHONE_NUMBER]
        If Not IsNull(rst_notar![id_ds]) Then pstrIdDS = rst_notar![id_ds]
        
    Else
        'MsgBox "Notáø k dìdicovi nenalezen"
        
    End If
    
rst_notar.Close

End Sub



