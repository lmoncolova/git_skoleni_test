VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sonik_VyzadaniSoucinnosti 
   Caption         =   "Vyžádání souèinnosti"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7425
   OleObjectBlob   =   "Sonik_VyzadaniSoucinnosti.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sonik_VyzadaniSoucinnosti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit



Private Sub CB_generovat_Click()

'Zavøu recordset, už jej nepotøebuji
Sonik_AO_rst_address.Close


Sonik_AO_Adresa = Replace(Sonik_VyzadaniSoucinnosti.TextSubjAdresa.value, "; ", "" & Chr(10) & "")
Sonik_AO_Subjekt = Sonik_VyzadaniSoucinnosti.ComboSubjekt.value
If Sonik_AO_Subjekt <> "Jiný" Then
    Sonik_AO_Adresa = Sonik_AO_Subjekt & Chr(10) & Sonik_AO_Adresa
End If
Sonik_AO_IDds = Sonik_VyzadaniSoucinnosti.TextIDds.value
If Sonik_AO_IDds <> "" Then
    Sonik_AO_Adresa = Sonik_AO_Adresa & Chr(10) & "ID " & Sonik_AO_IDds
End If
Sonik_AO_SpisZn = Sonik_Karta_Klienta.ListView_Soucinnosti_Klienta.SelectedItem.SubItems(5)
Sonik_AO_ZadostOd = Sonik_Karta_Klienta.ListView_Soucinnosti_Klienta.SelectedItem.SubItems(3)
Sonik_AO_ZadostDen = format$(Sonik_Karta_Klienta.ListView_Soucinnosti_Klienta.SelectedItem.SubItems(2), "d. m. yyyy")
Sonik_AO_KontrolaCastek = Sonik_VyzadaniSoucinnosti.TextKontrolaCastek.value
Sonik_AO_OdPohybFin = Sonik_VyzadaniSoucinnosti.TextOdPohyb.value
'Toto je realne cislo jednaci
Sonik_AO_SpisZn2 = Sonik_Karta_Klienta.ListView_Soucinnosti_Klienta.SelectedItem.SubItems(4)


Sonik_VyzadaniSoucinnosti.Hide

Call Sonik_AO_GenerujAutOdpoved
End Sub

Private Sub ComboSubjekt_Change()

'Vybiraji z dat, ktera jsem si predtim ulozil z ciselniku na DWH do Public recordsetu


   Sonik_AO_rst_address.MoveFirst
    For i = 1 To Sonik_AO_rst_address.RecordCount
    
    If Sonik_AO_rst_address![RECIPIENT_NAME] = Sonik_VyzadaniSoucinnosti.ComboSubjekt.text Then
    
    Sonik_VyzadaniSoucinnosti.TextSubjAdresa.value = "" & Sonik_AO_rst_address![RECIPIENT_ADDRESS] & ""
    Sonik_VyzadaniSoucinnosti.TextIDds.value = "" & Sonik_AO_rst_address![id_data_box] & ""
    Sonik_AO_Tel_Cislo_Adresata = "" & Sonik_AO_rst_address![TEXT_PHONE_NUMBER] & ""
    
    End If
    
    Sonik_AO_rst_address.MoveNext
    Next i
    




End Sub





Private Sub CommandButton_SameAsSender_Click()
'RISK-978
Dim sql As String
Dim StrSelectDataMessageSender As String ' Podívám se, která datovka klienta je oznaèena na jeho kartì... a od tam vyberu


StrSelectDataMessageSender = Sonik_Karta_Klienta.ListView_Soucinnosti_Klienta.SelectedItem.SubItems(3)


sql = "SELECT prosecutor_name,prosecutor_address,id_data_box from " & schema & ".sonik_cl_prosecutor" & _
" where replace(prosecutor_name,' ') = '" & Replace(StrSelectDataMessageSender, " ", "") & "'"

Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount

        ComboSubjekt.text = rst1![prosecutor_name]
        TextSubjAdresa.text = rst1![prosecutor_address]
        TextIDds.text = rst1![id_data_box]


    rst1.MoveNext
    Next i

   Else: MsgBox "Je mi líto, ale údaje k zastupitelství se nepodaøilo najít", vbInformation, "Data nenalezena"

   End If

   rst1.Close

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, _
                            CloseMode As Integer)
Dim ans
If CloseMode = vbFormControlMenu Then
    Cancel = True
    'Zavøu recordset, už jej nepotøebuji
    Sonik_AO_rst_address.Close
    Call Me.Hide
End If
End Sub

