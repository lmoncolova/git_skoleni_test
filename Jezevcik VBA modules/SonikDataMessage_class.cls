VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SonikDataMessage_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'RISK-1235


Private pLngSkpDataMessage As Long
Private pStrCisloDatoveZpravy As String
Private pDtDatumdoruceni As Date
Private pStrOdesilatel As String
Private pStrCisloJednaci As String
Private pStrSpisovaZnacka As String



Public Property Get LngSkpDataMessage() As Long
     LngSkpDataMessage = pLngSkpDataMessage
End Property

Public Property Get StrCisloDatoveZpravy() As String
     'Jediné mùže být null, doplòují až po zaevidování datovky
     StrCisloDatoveZpravy = pStrCisloDatoveZpravy
End Property

Public Property Get DtDatumdoruceni() As Date
     DtDatumdoruceni = pDtDatumdoruceni
End Property

Public Property Get StrOdesilatel() As String
     StrOdesilatel = pStrOdesilatel
End Property

Public Property Get StrCisloJednaci() As String
     StrCisloJednaci = pStrCisloJednaci
End Property

Public Property Get StrSpisovaZnacka() As String
     StrSpisovaZnacka = pStrSpisovaZnacka
End Property



'inicializacni skript pro naplneni zakladnich dat o datové zprávì
Public Sub initialize(skpDM As Long)

Dim sql As String
Dim rst1 As ADODB.Recordset


pLngSkpDataMessage = skpDM

sql = "select data_message_number, date_data_message_delivery, data_message_sender, text_file_number, text_reference_number" & _
      " FROM " & schema & ".sonik_data_message " & _
      " WHERE skp_data_message = " & pLngSkpDataMessage


   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
       If (rst1.RecordCount = 1) Then
    
            pStrCisloDatoveZpravy = rst1![data_message_number]
            pDtDatumdoruceni = rst1![date_data_message_delivery]
            pStrOdesilatel = rst1![data_message_sender]
            pStrSpisovaZnacka = rst1![text_file_number]
            pStrCisloJednaci = "" & rst1![text_reference_number] 'Jediné mùže být null, doplòují až po zaevidování datovky
         
       Else
            Call assertion(1 = 0, "jez-543747 : nepodaøilo se naèíst datovou zprávu")
       End If
   
   rst1.Close


End Sub















