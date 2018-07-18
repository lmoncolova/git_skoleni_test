VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sonik_Rucni_Cislo_Jednaci_DZ 
   Caption         =   "Soník - Zadávání subjektù"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6750
   OleObjectBlob   =   "Sonik_Rucni_Cislo_Jednaci_DZ.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sonik_Rucni_Cislo_Jednaci_DZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit




'Potvrzení a zavøení

Private Sub CommandButton_Uloz_Click()


Dim sql_update_file_num As String

sql_update_file_num = "update " & schema & ".sonik_data_message set text_reference_number = '" & TextBox_Cislo_Jednaci.value & "' where skp_data_message = " & SKP_datove_zpravy

Call execute(sql_update_file_num)




'A obnovim vypis dat

'Pokud to dìlám pøímo klikem z datové zprávy, tak to tam napoprvé vložím natvrdo, pøi obnovení, už se naète z DTB
If userform_prv = "Sonik_Datova_Zprava" Then
Sonik_Datova_Zprava.Label_DZ_Reference_NUM.Caption = TextBox_Cislo_Jednaci.value
Else
Call nacteni_dat_Sonik_Datove_Zpravy_Prehled(True)
End If

'Po zadání promažu
Me.TextBox_Cislo_Jednaci.text = ""

Me.Hide

End Sub




' Zavøení formuláøe
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Call set_filter(Me.name, 1)
Me.Hide
End Sub


' Back button
Private Sub CommandButton_Navrat_Click()
Call text_information_hide(formular)
Call UserForm_QueryClose(1, 0)
End Sub

















' Zmìna barvy po najetí myši **************************************************************************************************************************

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Uloz.ForeColor = barvaCerna
End Sub

Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Navrat.ForeColor = barvaCerna
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Uloz.ForeColor = barvaCerna
End Sub

Private Sub CommandButton_Navrat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Navrat.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Uloz_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Uloz.ForeColor = barvaZelenaAirBank
End Sub

