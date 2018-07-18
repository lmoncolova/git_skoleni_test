VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Stavy_update 
   Caption         =   "Jezevèík v 1.0 - Nový stav"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6465
   OleObjectBlob   =   "Stavy_update.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Stavy_update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit


Private Sub CommandButton12_Click()
    Call Me.Hide
'    Call Stavy.Hide
'    Call Stavy.Show
End Sub

Private Sub CommandButton20_Click()
Dim sql As String
Dim id As Integer


If cs_HDWMN.State = 1 Then
prihlasen:
 If Me.Caption = "Pøidání nového stavu" Then
 
        For i = 1 To Stavy.ListView1.ListItems.count
            If Trim(TextBox1.text) = Stavy.ListView1.ListItems(i).SubItems(1) Then
                GoTo rovnost
            End If
        Next i
 
If Len(Trim(TextBox1.text)) > 0 And Len(Trim(ComboBox1.text)) > 0 Then
   If MsgBox("Opravdu chcete pøidat nový stav?", vbYesNo, "Pøidání nového stavu") = vbYes Then
       
        If Me.ComboBox1.Column(1) = 3 Then
       
          sql = "insert into " & schema & ".epr_status values (" & max_skp("select max(skp_epr_status) from " & schema & ".epr_status") & ", '" & Trim(TextBox1.text) & _
            "', sysdate, to_date('1.1.3000','dd.mm.yyyy'),sysdate, '" & Main.username & "' , " & Me.ComboBox1.Column(1) & " , null)"
    
       ElseIf Me.ComboBox1.Column(1) = 2 Then
       
          sql = "insert into " & schema & ".exe_status values (" & max_skp("select max(skp_exe_status) from " & schema & ".exe_status") & ", '" & Trim(TextBox1.text) & _
            "', sysdate, to_date('1.1.3000','dd.mm.yyyy'),sysdate, '" & Main.username & "' , " & Me.ComboBox1.Column(1) & " , null)"
    
    
        ElseIf Me.ComboBox1.Column(1) = 1 Then
        
          sql = "insert into " & schema & ".ded_status values (" & max_skp("select max(skp_ded_status) from " & schema & ".ded_status") & ", '" & Trim(TextBox1.text) & _
            "', sysdate, to_date('1.1.3000','dd.mm.yyyy'),sysdate, '" & Main.username & "' , " & Me.ComboBox1.Column(1) & ", null)"
        
        ElseIf Me.ComboBox1.Column(1) = 5 Then
        
          sql = "insert into " & schema & ".trest_status values (" & max_skp("select max(skp_trest_status) from " & schema & ".trest_status") & ", '" & Trim(TextBox1.text) & _
            "', sysdate, to_date('1.1.3000','dd.mm.yyyy'),sysdate, '" & Main.username & "' , " & Me.ComboBox1.Column(1) & ",null)"
            
        ElseIf Me.ComboBox1.Column(1) = 4 Then
        
          sql = "insert into " & schema & ".ins_status values (" & max_skp("select max(skp_ins_status) from " & schema & ".ins_status") & ", '" & Trim(TextBox1.text) & _
            "', sysdate, to_date('1.1.3000','dd.mm.yyyy'),sysdate, '" & Main.username & "' , " & Me.ComboBox1.Column(1) & ",null)"
    
        ElseIf Me.ComboBox1.Column(1) = 6 Then
            'RISK-468 RISK-949
            sql = "insert into " & schema & ".heir_status values (" & max_skp("select max(skp_heir_status) from " & schema & ".heir_status") & ", '" & Trim(TextBox1.text) & _
                "', sysdate, to_date('1.1.3000','dd.mm.yyyy'),sysdate, '" & Main.username & "' , " & Me.ComboBox1.Column(1) & ",null)"
        
        End If
        
      Call execute(sql)
      
       Me.Hide
       Stavy.Show
       
   
   
  End If
  
  Else
rovnost:
    MsgBox ("Nelze zadat prázdnou hodnotu nebo hodnotu, která již existuje.")
  
 End If
 

 Else
  
  If CheckBox1.value Then
   If MsgBox("Opravdu chcete zneaktivnit stávající stav?", vbYesNo, "Zneaktivnìní stavu") = vbYes Then
   
    If Me.ComboBox1.Column(1) = 3 Then
    
        sql = "update " & schema & ".epr_status set dtime_valid_to = sysdate - 1/(24*60), dtime_modification = sysdate, employee = '" & Main.username & "' where skp_epr_status = " & Stavy.ListView1.SelectedItem

    ElseIf Me.ComboBox1.Column(1) = 1 Then
        
        sql = "update " & schema & ".ded_status set dtime_valid_to = sysdate - 1/(24*60), dtime_modification = sysdate, employee = '" & Main.username & "' where skp_ded_status = " & Stavy.ListView1.SelectedItem
    
    ElseIf Me.ComboBox1.Column(1) = 2 Then
        
        sql = "update " & schema & ".exe_status set dtime_valid_to = sysdate - 1/(24*60), dtime_modification = sysdate, employee = '" & Main.username & "' where skp_exe_status = " & Stavy.ListView1.SelectedItem
    
    
    ElseIf Me.ComboBox1.Column(1) = 5 Then
    
        sql = "update " & schema & ".trest_status set dtime_valid_to = sysdate - 1/(24*60), dtime_modification = sysdate, employee = '" & Main.username & "' where skp_trest_status = " & Stavy.ListView1.SelectedItem
    
    ElseIf Me.ComboBox1.Column(1) = 4 Then
    
        sql = "update " & schema & ".ins_status set dtime_valid_to = sysdate - 1/(24*60), dtime_modification = sysdate, employee = '" & Main.username & "' where skp_ins_status = " & Stavy.ListView1.SelectedItem
    
    
    ElseIf Me.ComboBox1.Column(1) = 6 Then
        'RISK-468 RISK-949
        sql = "update " & schema & ".heir_status set dtime_valid_to = sysdate - 1/(24*60), dtime_modification = sysdate, employee = '" & Main.username & "' where skp_heir_status = " & Stavy.ListView1.SelectedItem
    
    End If

    Call execute(sql)
  
    Me.Hide
   Stavy.Show
   
  End If
  
  Else
  If Len(Trim(TextBox1.text)) > 0 Then
    
    MsgBox "Pro zmìnu stavu napište svému oblíbenému analytikovi.", vbCritical
'     If MsgBox("Opravdu chcete upravit stávající stav?", vbYesNo, "Úprava stávajícího stavu") = vbYes Then
'
'          If Me.ComboBox1.Column(1) = 3 Then
'
'           sql = "update " & schema & ".epr_status set name_epr_status = '" & Trim(TextBox1.text) & "', dtime_modification = sysdate, employee = '" & Main.username & "' , skp_process = " & Me.ComboBox1.Column(1) & " where skp_epr_status = " & Stavy.ListView1.SelectedItem
'
'          ElseIf Me.ComboBox1.Column(1) = 1 Then
'
'           sql = "update " & schema & ".ded_status set name_ded_status = '" & Trim(TextBox1.text) & "', dtime_modification = sysdate, employee = '" & Main.username & "' , skp_process = " & Me.ComboBox1.Column(1) & " where skp_ded_status = " & Stavy.ListView1.SelectedItem
'
'          ElseIf Me.ComboBox1.Column(1) = 2 Then
'
'           sql = "update " & schema & ".exe_status set name_exe_status = '" & Trim(TextBox1.text) & "', dtime_modification = sysdate, employee = '" & Main.username & "' , skp_process = " & Me.ComboBox1.Column(1) & " where skp_exe_status = " & Stavy.ListView1.SelectedItem
'
'
'          ElseIf Me.ComboBox1.Column(1) = 5 Then
'
'           sql = "update " & schema & ".trest_status set name_trest_status = '" & Trim(TextBox1.text) & "', dtime_modification = sysdate, employee = '" & Main.username & "' , skp_process = " & Me.ComboBox1.Column(1) & " where skp_trest_status = " & Stavy.ListView1.SelectedItem
'
'          ElseIf Me.ComboBox1.Column(1) = 4 Then
'
'           sql = "update " & schema & ".ins_status set name_ins_status = '" & Trim(TextBox1.text) & "', dtime_modification = sysdate, employee = '" & Main.username & "' , skp_process = " & Me.ComboBox1.Column(1) & " where skp_ins_status = " & Stavy.ListView1.SelectedItem
'
'          'RISK-468 RISK-949
'          ElseIf Me.ComboBox1.Column(1) = 6 Then
'
'               sql = "update " & schema & ".heir_status set name_heir_status = '" & Trim(TextBox1.text) & "', dtime_modification = sysdate, employee = '" & Main.username & "' , skp_process = " & Me.ComboBox1.Column(1) & " where skp_heir_status = " & Stavy.ListView1.SelectedItem
'
'          End If
'
'    Call execute(sql)
'
'     Me.Hide
'     Stavy.Show
'
'    End If
 End If
 End If
 
 
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



Private Sub CheckBox1_Click()
 If Me.Caption = "Úprava stávajícího stavu" Then
  TextBox1.text = Stavy.ListView1.SelectedItem.SubItems(1)
  TextBox1.Enabled = Not CheckBox1.value
  ComboBox1.Enabled = Not CheckBox1.value
 End If
End Sub

' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()
Dim sql As String

Call set_filter("Stavy", 1)

If cs_HDWMN.State = 1 Then
prihlasen:

'Stavy_update.Caption = "Jezevèík v." & version & " - Zmìna stavu "

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
CommandButton12.ForeColor = barvaCerna
CommandButton20.ForeColor = barvaCerna

End Sub

Private Sub CommandButton12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton12.ForeColor = barvaZelenaAirBank
End Sub

Private Sub CommandButton20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
CommandButton20.ForeColor = barvaZelenaAirBank
End Sub

