VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sonik_Zaevidovani_Dopisu 
   Caption         =   "Soník - Zaevidování dopisu"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7740
   OleObjectBlob   =   "Sonik_Zaevidovani_Dopisu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sonik_Zaevidovani_Dopisu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit

'RISK-877
Private Str_Sender As String, Str_File_Number As String, Str_Reference_Number As String, Date_Delivery As Date


Private Sub CommandButton_Cancel_Click()
    Call Me.Hide
End Sub



Private Sub CommandButton_Save_Click()
    
Dim sql_max_skp_DM As String, sql_insert_DM As String, sql_insert_DMS As String, sql_insert_DM_Notice_Automatic As String, rst1 As ADODB.Recordset, max_skp_DM As Long, Time_Offset As Integer
    
If MsgBox("Opravdu chceš vložit nový dopis?", vbYesNo, "Potvrzení vložení") = vbYes Then
    
    Str_Sender = Me.TextBox_Sender.value
    Str_File_Number = Me.TextBox_File_Number.value
    Str_Reference_Number = Me.TextBox_Reference_Number.value
    Date_Delivery = Me.MonthView_Delivery.value
    
    'Kontroly vyplnìnosti
    If Str_Sender = vbNullString Then
        MsgBox "Odesílatel zprávy nebyl vyplnìn", vbExclamation, "Chybné zadání dat"
        Exit Sub
    End If
    
    If Str_File_Number = vbNullString Then
        MsgBox "Spisová znaèka nebyla vyplnìna", vbExclamation, "Chybné zadání dat"
        Exit Sub
    End If
    
    If Str_Reference_Number = vbNullString Then
        MsgBox "Èíslo jednací nebylo vyplnìno", vbExclamation, "Chybné zadání dat"
        Exit Sub
    End If
    
    If Date_Delivery > date Then
        MsgBox "Datum doruèení je v budoucnosti", vbExclamation, "Chybné zadání dat"
        Exit Sub
    End If
    
    
    
   'Naètu si maximum z datových zpráv - nejde pøes sekvenci, protože musím udìlat insert i do tabulky stavù a notice
   sql_max_skp_DM = "select nvl(max(skp_data_message),0) as max_skp from " & schema & ".sonik_data_message"

   Set rst1 = New ADODB.Recordset
   rst1.Open sql_max_skp_DM, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    max_skp_DM = rst1![max_skp]
    End If

   rst1.Close

   'Jenom pojistka pro první spuštìní
    If max_skp_DM >= 1 Then
    Else
    max_skp_DM = 0
    End If
    
    
    
    
    sql_insert_DM = "insert into  " & schema & ".sonik_data_message values (" & schema & ".sonik_data_message_seq.nextval, 'P'||" & schema & ".sonik_data_message_letter_seq.nextval,to_date('" & Date_Delivery & "','DD.MM.YYYY'),'" & Str_Sender & "','" & Str_File_Number & "', '" & Str_Reference_Number & "', sysdate, '" & UCase(Main.username) & "')"
    sql_insert_DMS = "insert into  " & schema & ".sonik_data_message_status values (" & schema & ".sonik_data_message_status_seq.nextval," & max_skp_DM + 1 & ", 1, sysdate,'" & UCase(Main.username) & "', sysdate, '1.1.3000', 'Y')"
    
    
    'Nahrání upozornìní na zpracování 30 dní
    Time_Offset = 30
    sql_insert_DM_Notice_Automatic = "insert into " & schema & ".sonik_data_message_notice values (" & schema & ".sonik_data_message_notice_seq.nextval," & max_skp_DM + 1 & ",'Termín zpracování souèinnosti',sysdate,to_date('" & Date_Delivery & "','DD.MM.YYYY')+ " & Time_Offset & " ,'Aktivní','sys',null,null)"
    
    
    
    Call execute(sql_insert_DM)
    Call execute(sql_insert_DMS)
    Call execute(sql_insert_DM_Notice_Automatic)
    
    'Promazání formuláøe
    Me.TextBox_Sender.value = vbNullString
    Me.TextBox_File_Number.value = vbNullString
    Me.TextBox_Reference_Number.value = vbNullString
    Me.MonthView_Delivery.value = date
    
    
    
    
    Call Me.Hide
End If
End Sub



' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()
MonthView_Delivery = date
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
CommandButton_Cancel.ForeColor = barvaCerna
CommandButton_Save.ForeColor = barvaCerna

End Sub

Private Sub CommandButton_Cancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Cancel.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton_Save_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton_Save.ForeColor = barvaZelenaAirBank
End Sub

