VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Pridanie_klienta 
   Caption         =   "Jezevèík"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6465
   OleObjectBlob   =   "Pridanie_klienta.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Pridanie_klienta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

Private Sub CommandButton12_Click()
    Call Me.Hide

Select Case userform_prv
 Case "Nezesplatneni"
  Nezesplatneni.Show
 Case "Zesplatneni"
  Zesplatneni.Show
End Select

End Sub

Private Sub CommandButton20_Click()
Dim sql As String
Dim id As Integer


If cs_HDWMN.State = 1 Then
prihlasen:

If IsNumeric(TextBox1.text) And InStr(1, TextBox1.text, " ") = 0 Then

    'If MsgBox("Opravdu chcete klienta pøidat?", vbYesNo, "Pøidání poznámky") = vbYes Then
  
         sql = "insert into " & schema & ".cuid_addition values (" & max_skp("select max(skp_cuid_addition) from " & schema & ".cuid_addition") & _
              " , '" & Replace(TextBox1.text, " ", "") & "', sysdate, '" & Main.username & "' , 'N')"
              
              'a = InputBox("", "", sql)
              
            Call execute(sql)
        
          Call UserForm_Activate
          
    'End If
    
    Else
    
    MsgBox ("Zadaný CUID je v nesprávném formátu!")
    
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


' ListView sorting
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call Listview_Proper_Sort(ListView1, ColumnHeader)
End Sub

' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()
Dim sql As String


If cs_HDWMN.State = 1 Then
prihlasen:

Me.TextBox1.text = ""

Me.ListView1.ListItems.Clear
Me.ListView1.ColumnHeaders.Clear

Call table_prepare(Me.name)

Call query1("select * from " & schema & ".v_cuid_addition", Me.name)


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



Private Sub UserForm_QueryClose(Cancel As Integer, _
                            CloseMode As Integer)
Dim ans
If CloseMode = vbFormControlMenu Then
    Cancel = True
    Call Me.Hide

Select Case userform_prv
 Case "Nezesplatneni"
  Nezesplatneni.Show
 Case "Zesplatneni"
  Zesplatneni.Show
End Select

End If
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton12.ForeColor = barvaCerna
CommandButton20.ForeColor = barvaCerna

End Sub

Private Sub CommandButton12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton12.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton20.ForeColor = barvaZelenaAirBank
End Sub



