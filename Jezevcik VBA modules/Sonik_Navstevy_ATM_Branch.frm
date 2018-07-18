VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sonik_Navstevy_ATM_Branch 
   Caption         =   "Soník - Návštìvy klienta na bankomatech a poboèkách AB"
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13320
   OleObjectBlob   =   "Sonik_Navstevy_ATM_Branch.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sonik_Navstevy_ATM_Branch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit





Private Sub UserForm_Activate()

Dim sql_atm As String
Dim sql_branch As String
Dim clmx As Variant
Dim itmX As Variant

If Sonik_Navstevy_ATM_Branch.ListView_ATM.ColumnHeaders.count = 0 Then


 Set clmx = Sonik_Navstevy_ATM_Branch.ListView_ATM.ColumnHeaders.Add(, , "Èíslo karty", 120)
 Set clmx = Sonik_Navstevy_ATM_Branch.ListView_ATM.ColumnHeaders.Add(, , "Datum transakce", 90)
 Set clmx = Sonik_Navstevy_ATM_Branch.ListView_ATM.ColumnHeaders.Add(, , "Typ transakce", 100)
 Set clmx = Sonik_Navstevy_ATM_Branch.ListView_ATM.ColumnHeaders.Add(, , "Výše transakce", 80)
 Set clmx = Sonik_Navstevy_ATM_Branch.ListView_ATM.ColumnHeaders.Add(, , "ID Bankomatu", 80)
 Set clmx = Sonik_Navstevy_ATM_Branch.ListView_ATM.ColumnHeaders.Add(, , "Adresa bankomatu", 150)

End If

If skp_person_KL = -1 Then

sql_atm = "select TEXT_CARD_NUMBER, DTIME_ATM_TRANSACTION, NAME_ATM_TRANSACTION_TYPE, AMT_TRANSACTION, ID_ATM_TERMINAL, ATM_ADDRESS from " & schema & ".sonik_visit_ATM" & _
 " where TEXT_CARD_NUMBER = " & NUM_Card_KL & " order by TEXT_CARD_NUMBER, DTIME_ATM_TRANSACTION desc"
    
Else

sql_atm = "select TEXT_CARD_NUMBER, DTIME_ATM_TRANSACTION, NAME_ATM_TRANSACTION_TYPE, AMT_TRANSACTION, ID_ATM_TERMINAL, ATM_ADDRESS from " & schema & ".sonik_visit_ATM" & _
 " where skp_person = " & skp_person_KL & " order by TEXT_CARD_NUMBER, DTIME_ATM_TRANSACTION desc"

End If

   Sonik_Navstevy_ATM_Branch.ListView_ATM.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql_atm, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Navstevy_ATM_Branch.ListView_ATM.ListItems.Add()
      itmX.text = "" & rst1![TEXT_CARD_NUMBER] & ""
      itmX.SubItems(1) = "" & rst1![DTIME_ATM_TRANSACTION] & ""
      itmX.SubItems(2) = "" & rst1![NAME_ATM_TRANSACTION_TYPE] & ""
      itmX.SubItems(3) = "" & rst1![AMT_TRANSACTION] & ""
      itmX.SubItems(4) = "" & rst1![ID_ATM_TERMINAL] & ""
      itmX.SubItems(5) = "" & rst1![ATM_ADDRESS] & ""
    rst1.MoveNext
    Next i
   End If

   rst1.Close




If Sonik_Navstevy_ATM_Branch.ListView_Branch.ColumnHeaders.count = 0 Then


 Set clmx = Sonik_Navstevy_ATM_Branch.ListView_Branch.ColumnHeaders.Add(, , "Datum kontaktu", 120)
 Set clmx = Sonik_Navstevy_ATM_Branch.ListView_Branch.ColumnHeaders.Add(, , "Jméno bankovního asistenta", 150)
 Set clmx = Sonik_Navstevy_ATM_Branch.ListView_Branch.ColumnHeaders.Add(, , "Poboèka", 120)
 Set clmx = Sonik_Navstevy_ATM_Branch.ListView_Branch.ColumnHeaders.Add(, , "Adresa poboèky", 150)

End If



sql_branch = "select DTIME_CONTACT_REALIZED, NAME_EMPLOYEE, NAME_BRANCH_OFFICE, BRANCH_ADDRESS from " & schema & ".sonik_visit_branch" & _
 " where skp_person = " & skp_person_KL & " order by DTIME_CONTACT_REALIZED desc"


   Sonik_Navstevy_ATM_Branch.ListView_Branch.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql_branch, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Sonik_Navstevy_ATM_Branch.ListView_Branch.ListItems.Add()
      itmX.text = "" & rst1![DTIME_CONTACT_REALIZED] & ""
      itmX.SubItems(1) = "" & rst1![NAME_EMPLOYEE] & ""
      itmX.SubItems(2) = "" & rst1![NAME_BRANCH_OFFICE] & ""
      itmX.SubItems(3) = "" & rst1![BRANCH_ADDRESS] & ""
    rst1.MoveNext
    Next i
   End If

   rst1.Close





End Sub




' Zavøení formuláøe
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Call set_filter(Me.name, 1)


Me.Hide
Sonik_Karta_Klienta.Show
End Sub


' Back button
Private Sub CommandButton_Navrat_Click()


'Call text_information_hide(formular)
Call UserForm_QueryClose(1, 0)
Sonik_Karta_Klienta.Show
End Sub



' ListView sorting
Private Sub ListView_ATM_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_ATM, ColumnHeader)
End Sub


' ListView sorting
Private Sub ListView_Branch_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView_Branch, ColumnHeader)
End Sub






' Zmìna barvy po najetí myši **************************************************************************************************************************



Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Navrat.ForeColor = barvaCerna
End Sub

Private Sub CommandButton_Navrat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Navrat.ForeColor = barvaZelenaAirBank
End Sub


