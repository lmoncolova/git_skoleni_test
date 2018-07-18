VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Export_plateb 
   Caption         =   "Export plateb"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6000
   OleObjectBlob   =   "Export_plateb.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Export_plateb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit
'Static export_chosen_option As Integer 'premenna bude mat hodnoty 1-4, pre export pdf/xml, aktualny/vsetky uvery
'GroupName je pdf_OR_xml
''FA_export - cely userform

Private Sub Button_export_Click()
    Unload Me
End Sub

Private Sub PDF_jeden_Click()
    export_chosen_option = 1
End Sub

Private Sub PDF_vsechny_Click()
    export_chosen_option = 2
End Sub

Private Sub XML_jeden_Click()
    export_chosen_option = 3
End Sub

Private Sub XML_vsechny_Click()
    export_chosen_option = 4
End Sub



Private Sub UserForm_Activate()

    PDF_jeden.value = True

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, _
                            CloseMode As Integer)
Dim ans
If CloseMode = vbFormControlMenu Then
    Cancel = True
    export_chosen_option = 0
    Call Me.Hide
End If
End Sub

