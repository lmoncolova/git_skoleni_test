VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sonik_Upozorneni_DZ 
   Caption         =   "Soník - Pøidat upozornìní"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6465
   OleObjectBlob   =   "Sonik_Upozorneni_DZ.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sonik_Upozorneni_DZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit


Private Sub CommandButton_Zrus_Click()
    Call Me.Hide
End Sub

Private Sub CommandButton_Uloz_Click()
Dim sql As String
Dim id As Integer

  If Len(Trim(TextBox1.text)) = 0 Then
  
           MsgBox "Musíš zadat text upozornìní", vbExclamation, "Chybné zadání hodnot"
         
   Exit Sub
  End If

' Insert
If Me.Caption = "Nové upozornìní" Then

  If MsgBox("Opravdu chcete pøidat upozornìní?", vbYesNo, "Pøidání upozornìní") = vbYes Then
  

        Dim sql_insert_notice As String

        sql_insert_notice = "insert into " & schema & ".sonik_data_message_notice values (" & schema & ".sonik_data_message_notice_seq.nextval, " & SKP_datove_zpravy & ",'" & TextBox1.text & "', sysdate, to_date('" & MonthView1.value & "','dd.mm.yyyy'),'" & ComboBox1.text & "','" & UCase(Main.username) & "', null, null)"
        Call execute(sql_insert_notice)


  End If

 
 
'Update
Else

   If MsgBox("Opravdu chcete upravit upozornìní?", vbYesNo, "Úprava upozornìní") = vbYes Then
     
        Dim sql_update_notice As String
     

        sql_update_notice = "update " & schema & ".sonik_data_message_notice set TEXT_NOTICE = '" & TextBox1.text & "',DTIME_NOTICE_DEADLINE = to_date('" & MonthView1.value & "','dd.mm.yyyy hh24:mi:ss'), CODE_STATUS = '" & ComboBox1.text & "' , DTIME_MODIFIED = sysdate, NAME_EMPLOYEE_MODIFIED = '" & UCase(Main.username) & "' where SKP_DATA_MESSAGE_NOTICE = " & Sonik_Datova_Zprava.ListView_DZ_Upozorneni.SelectedItem.text
        Call execute(sql_update_notice)


  End If
End If

   Me.Hide
   Call nacteni_dat_Sonik_Datova_Zprava
   Sonik_Datova_Zprava.Show

End Sub



Private Sub TextBox1_Change()
Label32.Caption = 200 - Len(TextBox1.text)
End Sub

' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()
Dim sql As String


If cs_HDWMN.State = 1 Then
prihlasen:


ComboBox1.Clear
ComboBox1.AddItem "Aktivní"
ComboBox1.AddItem "Storno/Smazat"
ComboBox1.AddItem "Dokonèeno"
ComboBox1.ListIndex = 0

If Flag_Nove_Upozorneni_DZ = 0 And Sonik_Datova_Zprava.ListView_DZ_Upozorneni.ListItems.count > 0 Then
 
 Label26.Caption = "Zmìna upozornìní"
 Me.Caption = "Zmìna upozornìní"
 
 MonthView1 = Sonik_Datova_Zprava.ListView_DZ_Upozorneni.SelectedItem.SubItems(2)
 TextBox1.text = Sonik_Datova_Zprava.ListView_DZ_Upozorneni.SelectedItem.SubItems(1)
Else
 Label26.Caption = "Nové upozornìní"
 Me.Caption = "Nové upozornìní"
 TextBox1.text = ""
 MonthView1 = date
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



Private Sub UserForm_QueryClose(Cancel As Integer, _
                            CloseMode As Integer)
Dim ans
If CloseMode = vbFormControlMenu Then
    Cancel = True
    Call Me.Hide
End If
End Sub


' Zmìna barvy po najetí myši **************************************************************************************************************************
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Zrus.ForeColor = barvaCerna
CommandButton_Uloz.ForeColor = barvaCerna

End Sub

Private Sub CommandButton_Zrus_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Zrus.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Uloz_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Uloz.ForeColor = barvaZelenaAirBank
End Sub


