VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Update_SOP 
   Caption         =   "Jezevèík - Nalézací øízení"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10815
   OleObjectBlob   =   "Update_SOP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Update_SOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit


Private Sub CommandButton12_Click()
    Call Me.Hide
End Sub

Private Sub CommandButton20_Click()
    Dim sql As String
    
    Dim BeforeValue As String
    Dim AfterValue As String
    Dim id As Integer
    Dim zmena As String
    Dim start_query As Boolean
    Dim preplatekA1 As Boolean
    Dim preplatekB1 As Boolean
    Dim preplatekA2 As Boolean
    Dim preplatekB2 As Boolean
    
    
    
    If cs_HDWMN.State = 1 Then
prihlasen:
    
    Call text_information_hide(Me.name)
    
    start_query = False
    preplatekA1 = False
    preplatekB1 = False
    
    zmena = ""
    
    If prava("select * from " & UCase(schema) & ".sop where id_person = " & id_person) Then
        sql = "update " & schema & ".sop set "
        
        If CheckBox1.value And TextBox1.text <> Label1.Caption Then
            If IsNumeric(TextBox1.text) And Len(Trim(TextBox1.text)) > 0 Then
                sql = sql & "amt_sop1 = " & Replace(Replace(TextBox1.text, ",", "."), " ", "") & ", date_sop1 = NULL, "
                start_query = True
                zmena = zmena & "SOP1 (" & TextBox1.text & "), "
        Else
            GoTo chyba
        End If
    End If
     
    ' doplnene
    If CheckBox15.value And TextBox14.text <> Label46.Caption Then
        If IsNumeric(TextBox14.text) And Len(Trim(TextBox14.text)) > 0 Then
            sql = sql & "amt_sop1_court = " & Replace(Replace(TextBox14.text, ",", "."), " ", "") & ",  "
            start_query = True
            zmena = zmena & "soudem urèen SOP1 (" & TextBox14.text & "), "
        Else
            GoTo chyba
        End If
    End If
    '----------
     
      ' doplnene
      If CheckBox50.value And TextBox50.text <> Label50.Caption Then
      If IsNumeric(TextBox50.text) And Len(Trim(TextBox50.text)) > 0 Then
       sql = sql & "sop_additional_paid = " & Replace(Replace(TextBox50.text, ",", "."), " ", "") & ",  "
       start_query = True
       zmena = zmena & "Nenapárovaná platba SOP (" & TextBox50.text & "), "
      Else
       GoTo chyba
      End If
     End If
     '----------
     
      ' doplnene
      If CheckBox51.value And TextBox51.text <> Label51.Caption Then
      If IsNumeric(TextBox51.text) And Len(Trim(TextBox51.text)) > 0 Then
       sql = sql & "sop_additional_fee = " & Replace(Replace(TextBox51.text, ",", "."), " ", "") & ",  "
       start_query = True
       zmena = zmena & "Paušál (" & TextBox51.text & "), "
      Else
       GoTo chyba
      End If
     End If
     '----------
     
     If CheckBox3.value And TextBox3.text <> Label3.Caption Then
      If IsNumeric(TextBox3.text) And Len(Trim(TextBox3.text)) > 0 Then
       sql = sql & "AMT_ADDITIONAL_SOP1 = " & Replace(Replace(TextBox3.text, ",", "."), " ", "") & ", DATE_ADDITIONAL_SOP1 = NULL, "
       start_query = True
       zmena = zmena & "doplatek SOP1 (" & TextBox3.text & "), "
      Else
       GoTo chyba
      End If
     End If
    
     If CheckBox5.value And TextBox5.text <> Label5.Caption Then
      If IsNumeric(TextBox5.text) And Len(Trim(TextBox5.text)) > 0 Then
       sql = sql & "AMT_RETURN_SOP1 = " & Replace(Replace(TextBox5.text, ",", "."), " ", "") & ", "
       start_query = True
       preplatekA1 = True
       zmena = zmena & "pøeplatek SOP1 (" & TextBox5.text & "), "
      Else
       GoTo chyba
      End If
     End If
    
    
     If CheckBox6.value And DTPicker6 <> Label6.Caption Then
       sql = sql & "DATE_INFO_RETURN_SOP1 = to_date('" & DTPicker6.value & "','dd.mm.yyyy'), "
       start_query = True
       preplatekB1 = True
       zmena = zmena & "datum sdìlení pøeplatku SOP1 (" & DTPicker6.value & "), "
     End If
     
     
     If preplatekA1 Or preplatekB1 Then
       sql = sql & "AMT_RETURN_REAL_SOP1 = NULL, DATE_RETURN_REAL_SOP1 = NULL, "
     End If
     
    
    ' ---------------------------
    
     If CheckBox9.value And TextBox9.text <> Label9.Caption Then
      If IsNumeric(TextBox9.text) And Len(Trim(TextBox9.text)) > 0 Then
       sql = sql & "amt_sop2 = " & Replace(Replace(TextBox9.text, ",", "."), " ", "") & ", date_sop2 = NULL, "
       start_query = True
       zmena = zmena & "SOP2 (" & TextBox9.text & "), "
      Else
       GoTo chyba
      End If
     End If
     
     ' doplnene
     If CheckBox16.value And TextBox15.text <> Label47.Caption Then
      If IsNumeric(TextBox15.text) And Len(Trim(TextBox15.text)) > 0 Then
       sql = sql & "amt_sop2_court = " & Replace(Replace(TextBox15.text, ",", "."), " ", "") & ", "
       start_query = True
       zmena = zmena & "soudem urèen SOP2 (" & TextBox15.text & "), "
      Else
       GoTo chyba
      End If
     End If
     '----------
    
     If CheckBox11.value And TextBox11.text <> Label11.Caption Then
      If IsNumeric(TextBox11.text) And Len(Trim(TextBox11.text)) > 0 Then
       sql = sql & "AMT_ADDITIONAL_SOP2 = " & Replace(Replace(TextBox11.text, ",", "."), " ", "") & ", DATE_ADDITIONAL_SOP2 = NULL, "
       start_query = True
       zmena = zmena & "doplatek SOP2 (" & TextBox11.text & "), "
      Else
       GoTo chyba
      End If
     End If
    
     If CheckBox13.value And TextBox13.text <> Label13.Caption Then
      If IsNumeric(TextBox13.text) And Len(Trim(TextBox13.text)) > 0 Then
       sql = sql & "AMT_RETURN_SOP2 = " & Replace(Replace(TextBox13.text, ",", "."), " ", "") & ", "
       start_query = True
       preplatekA2 = True
       zmena = zmena & "pøeplatek SOP2 (" & TextBox13.text & "), "
      Else
       GoTo chyba
      End If
     End If
    
    
     If CheckBox14.value And DTPicker14 <> Label14.Caption Then
       sql = sql & "DATE_INFO_RETURN_SOP2 = to_date('" & DTPicker14.value & "','dd.mm.yyyy'), "
       start_query = True
       preplatekB2 = True
       zmena = zmena & "datum sdìlení pøeplatku SOP2 (" & DTPicker14.value & "), "
     End If
     
      ' doplnene
     If CheckBox34.value And TextBox34.text <> Label34.Caption Then
      If IsNumeric(TextBox34.text) And Len(Trim(TextBox34.text)) > 0 Then
       sql = sql & "travel_expenses = " & Replace(Replace(TextBox34.text, ",", "."), " ", "") & ", "
       start_query = True
       zmena = zmena & "cestovní (" & TextBox34.text & "), "
      Else
       GoTo chyba
      End If
     End If
     '----------
     
     If preplatekA2 Or preplatekB2 Then
       sql = sql & "AMT_RETURN_REAL_SOP2 = NULL, DATE_RETURN_REAL_SOP2 = NULL, "
     End If
     
     
     sql = sql & " dtime_modification = sysdate, employee = '" & Main.username & "' where id_person = " & id_person
    
    Else
    
     BeforeValue = "skp_sop, id_person, "
     AfterValue = "(" & max_skp("select max (skp_sop) from " & schema & ".sop") & ", " & id_person & ", "
     sql = "insert into " & schema & ".sop ("
     
     
     If CheckBox1.value And Len(Trim(TextBox1.text)) > 0 And TextBox1.text <> Label1.Caption Then
      If IsNumeric(TextBox1.text) Then
       BeforeValue = BeforeValue & "amt_sop1, "
       AfterValue = AfterValue & Replace(Replace(TextBox1.text, ",", "."), " ", "") & ", "
       start_query = True
       zmena = zmena & "SOP1 (" & TextBox1.text & "), "
      Else
       GoTo chyba
      End If
     End If
     
      ' doplnene
      If CheckBox15.value And Len(Trim(TextBox14.text)) > 0 And TextBox14.text <> Label46.Caption Then
      If IsNumeric(TextBox14.text) Then
       BeforeValue = BeforeValue & "amt_sop1_court, "
       AfterValue = AfterValue & Replace(Replace(TextBox14.text, ",", "."), " ", "") & ", "
       start_query = True
       zmena = zmena & "soudem urèen SOP1 (" & TextBox14.text & "), "
      Else
       GoTo chyba
      End If
     End If
     
       ' doplnene
      If CheckBox50.value And Len(Trim(TextBox50.text)) > 0 And TextBox50.text <> Label50.Caption Then
      If IsNumeric(TextBox50.text) Then
       BeforeValue = BeforeValue & "sop_additional_paid, "
       AfterValue = AfterValue & Replace(Replace(TextBox50.text, ",", "."), " ", "") & ", "
       start_query = True
       zmena = zmena & "Nenapárovaná platba SOP (" & TextBox50.text & "), "
      Else
       GoTo chyba
      End If
     End If
     
     'doplnena zmena pausalu
    If CheckBox51.value And Len(Trim(TextBox51.text)) > 0 And TextBox51.text <> Label51.Caption Then
      If IsNumeric(TextBox51.text) Then
       BeforeValue = BeforeValue & "sop_additional_fee , "
       AfterValue = AfterValue & Replace(Replace(TextBox51.text, ",", "."), " ", "") & ", "
       start_query = True
       zmena = zmena & "Nenapárovaná platba SOP (" & TextBox51.text & "), "
      Else
       GoTo chyba
      End If
     End If
     
     
     If CheckBox3.value And TextBox3.text <> Label3.Caption Then
      If IsNumeric(TextBox3.text) And Len(Trim(TextBox3.text)) > 0 Then
       BeforeValue = BeforeValue & "AMT_ADDITIONAL_SOP1, "
       AfterValue = AfterValue & Replace(Replace(TextBox3.text, ",", "."), " ", "") & ", "
       start_query = True
       zmena = zmena & "doplatek SOP1 (" & TextBox3.text & "), "
      Else
       GoTo chyba
      End If
     End If
    
     If CheckBox5.value And TextBox5.text <> Label5.Caption Then
      If IsNumeric(TextBox5.text) And Len(Trim(TextBox5.text)) > 0 Then
       BeforeValue = BeforeValue & "AMT_RETURN_SOP1, "
       AfterValue = AfterValue & Replace(Replace(TextBox5.text, ",", "."), " ", "") & ", "
       start_query = True
       zmena = zmena & "pøeplatek SOP1 (" & TextBox5.text & "), "
      Else
       GoTo chyba
      End If
     End If
    
    
     If CheckBox6.value And DTPicker6 <> Label6.Caption Then
       BeforeValue = BeforeValue & "DATE_INFO_RETURN_SOP1, "
       AfterValue = AfterValue & "to_date('" & DTPicker6.value & "','dd.mm.yyyy'), "
       start_query = True
       zmena = zmena & "datum sdìlení pøeplatku SOP1 (" & DTPicker6.value & "), "
     End If
     
    
    ' ---------------------------
    
     If CheckBox9.value And TextBox9.text <> Label9.Caption Then
      If IsNumeric(TextBox9.text) And Len(Trim(TextBox9.text)) > 0 Then
       BeforeValue = BeforeValue & "amt_SOP2, "
       AfterValue = AfterValue & Replace(Replace(TextBox9.text, ",", "."), " ", "") & ", "
       start_query = True
       zmena = zmena & "SOP2 (" & TextBox9.text & "), "
      Else
       GoTo chyba
      End If
     End If
     
      ' doplnene
      If CheckBox16.value And TextBox15.text <> Label47.Caption Then
      If IsNumeric(TextBox15.text) And Len(Trim(TextBox15.text)) > 0 Then
       BeforeValue = BeforeValue & "amt_SOP2_court, "
       AfterValue = AfterValue & Replace(Replace(TextBox15.text, ",", "."), " ", "") & ", "
       start_query = True
       zmena = zmena & "soudem urèen SOP2 (" & TextBox15.text & "), "
      Else
       GoTo chyba
      End If
     End If
     
     If CheckBox11.value And TextBox11.text <> Label11.Caption Then
      If IsNumeric(TextBox11.text) And Len(Trim(TextBox11.text)) > 0 Then
       BeforeValue = BeforeValue & "AMT_ADDITIONAL_SOP2, "
       AfterValue = AfterValue & Replace(Replace(TextBox11.text, ",", "."), " ", "") & ", "
       start_query = True
       zmena = zmena & "doplatek SOP2 (" & TextBox11.text & "), "
      Else
       GoTo chyba
      End If
     End If
    
     If CheckBox13.value And TextBox13.text <> Label13.Caption Then
      If IsNumeric(TextBox13.text) And Len(Trim(TextBox13.text)) > 0 Then
       BeforeValue = BeforeValue & "AMT_RETURN_SOP2, "
       AfterValue = AfterValue & Replace(Replace(TextBox13.text, ",", "."), " ", "") & ", "
       start_query = True
       zmena = zmena & "pøeplatek SOP2 (" & TextBox13.text & "), "
      Else
       GoTo chyba
      End If
     End If
     
       ' doplnene
      If CheckBox34.value And Len(Trim(TextBox34.text)) > 0 And TextBox34.text <> Label34.Caption Then
      If IsNumeric(TextBox34.text) Then
       BeforeValue = BeforeValue & "travel_expenses, "
       AfterValue = AfterValue & Replace(Replace(TextBox34.text, ",", "."), " ", "") & ", "
       start_query = True
       zmena = zmena & "cestovní (" & TextBox34.text & "), "
      Else
       GoTo chyba
      End If
     End If
    
    
     If CheckBox14.value And DTPicker6 <> Label14.Caption Then
       BeforeValue = BeforeValue & "DATE_INFO_RETURN_SOP2, "
       AfterValue = AfterValue & "to_date('" & DTPicker14.value & "','dd.mm.yyyy'), "
       start_query = True
       zmena = zmena & "datum sdìlení pøeplatku SOP2 (" & DTPicker14.value & "), "
     End If
      
    
    
    BeforeValue = BeforeValue & "dtime_modification, employee)"
    AfterValue = AfterValue & "sysdate, '" & Main.username & "')"
    sql = sql & BeforeValue & " values " & AfterValue
    End If
    
    
    
    
    If start_query Then
        If True Then ' MsgBox("Opravdu chcete uložit?", vbYesNo, "Uložit") = vbYes Then
            
            Call TransBegin
            
                execute (sql)
                
                sql = "insert into " & schema & ".epr_note values (" & max_skp("select max(skp_epr_note) from " & schema & ".epr_note") & ", " & id_person & _
                  ", 5, NULL, 'Zmìna údajù SOP: " & Mid(zmena, 1, Len(zmena) - 2) & "', sysdate, '" & Main.username & "')"
                
                execute (sql)
            
            Call TransCommit
            
            c.SOP.re_initialize
            Me.Hide
            Klient.Show
            odkaz_Klient = True
        Else
            Label29.Visible = True
            Label28.Visible = True
            Label29.Caption = "Zrušeno"
        End If
    Else
        Label29.Visible = True
        Label28.Visible = True
        Label29.Caption = "Žádné zmìny"
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
    Exit Sub

chyba:
    Label29.Visible = True
    Label28.Visible = True
    Label29.Caption = "Chybný formát vstupních dat"

End Sub



Private Sub CheckBox1_Click()
Label1.Visible = Not CheckBox1.value
TextBox1.Visible = CheckBox1.value
TextBox1.text = Label1.Caption

Label27.Visible = CheckBox1.value Or CheckBox3.value Or CheckBox5.value Or CheckBox6.value Or CheckBox9.value Or CheckBox11.value Or CheckBox13.value Or CheckBox14.value
Label29.Caption = "Po úpravì hodnot bude nutná jejich autorizace"
Label29.Visible = Label27.Visible
Label28.Visible = False

End Sub



Private Sub CheckBox3_Click()

Label3.Visible = Not CheckBox3.value
TextBox3.Visible = CheckBox3.value
TextBox3.text = Label3.Caption

Label27.Visible = CheckBox1.value Or CheckBox3.value Or CheckBox5.value Or CheckBox6.value Or CheckBox9.value Or CheckBox11.value Or CheckBox13.value Or CheckBox14.value
Label29.Caption = "Po úpravì hodnot bude nutná jejich autorizace"
Label29.Visible = Label27.Visible
Label28.Visible = False

End Sub



Private Sub CheckBox5_Click()
Call text_information_hide(Me.name)
Label5.Visible = Not CheckBox5.value
TextBox5.Visible = CheckBox5.value
TextBox5.text = Label5.Caption

Label27.Visible = CheckBox1.value Or CheckBox3.value Or CheckBox5.value Or CheckBox6.value Or CheckBox9.value Or CheckBox11.value Or CheckBox13.value Or CheckBox14.value
Label29.Caption = "Po úpravì hodnot bude nutná jejich autorizace"
Label29.Visible = Label27.Visible
Label28.Visible = False

' Preplatek se vzdy musi zadavat i s datem sdeleni
CheckBox6.value = CheckBox5.value
Call CheckBox6_Click
End Sub


Private Sub CheckBox6_Click()
Call text_information_hide(Me.name)
Label6.Visible = Not CheckBox6.value
DTPicker6.Visible = CheckBox6.value

If CheckBox6.value And Label6.Caption <> "" Then
DTPicker6 = Label6.Caption
ElseIf CheckBox6.value Then
DTPicker6 = date
End If

Me.Repaint

Label27.Visible = CheckBox1.value Or CheckBox3.value Or CheckBox5.value Or CheckBox6.value Or CheckBox9.value Or CheckBox11.value Or CheckBox13.value Or CheckBox14.value
Label29.Caption = "Po úpravì hodnot bude nutná jejich autorizace"
Label29.Visible = Label27.Visible
Label28.Visible = False

' Preplatek se vzdy musi zadavat i s datem sdeleni
CheckBox5.value = CheckBox6.value
End Sub
' doplnene
Private Sub CheckBox15_Click()

Label46.Visible = Not CheckBox15.value
TextBox14.Visible = CheckBox15.value
TextBox14.text = Label46.Caption

Label28.Visible = False

End Sub
' doplnene
Private Sub CheckBox50_Click()

Label50.Visible = Not CheckBox50.value
TextBox50.Visible = CheckBox50.value
TextBox50.text = Label50.Caption

Label28.Visible = False

End Sub


Private Sub CheckBox51_Click()

Label51.Visible = Not CheckBox50.value
TextBox51.Visible = CheckBox51.value
TextBox51.text = Label51.Caption

Label51.Visible = False

End Sub


Private Sub CheckBox9_Click()
Call text_information_hide(Me.name)
Label9.Visible = Not CheckBox9.value
TextBox9.Visible = CheckBox9.value
TextBox9.text = Label9.Caption

Label27.Visible = CheckBox1.value Or CheckBox3.value Or CheckBox5.value Or CheckBox6.value Or CheckBox9.value Or CheckBox11.value Or CheckBox13.value Or CheckBox14.value
Label29.Caption = "Po úpravì hodnot bude nutná jejich autorizace"
Label29.Visible = Label27.Visible
Label28.Visible = False

End Sub




Private Sub CheckBox11_Click()
Call text_information_hide(Me.name)
Label11.Visible = Not CheckBox11.value
TextBox11.Visible = CheckBox11.value
TextBox11.text = Label11.Caption

Label27.Visible = CheckBox1.value Or CheckBox3.value Or CheckBox5.value Or CheckBox6.value Or CheckBox9.value Or CheckBox11.value Or CheckBox13.value Or CheckBox14.value
Label29.Caption = "Po úpravì hodnot bude nutná jejich autorizace"
Label29.Visible = Label27.Visible
Label28.Visible = False

End Sub



Private Sub CheckBox13_Click()
Call text_information_hide(Me.name)
Label13.Visible = Not CheckBox13.value
TextBox13.Visible = CheckBox13.value
TextBox13.text = Label13.Caption

Label27.Visible = CheckBox1.value Or CheckBox3.value Or CheckBox5.value Or CheckBox6.value Or CheckBox9.value Or CheckBox11.value Or CheckBox13.value Or CheckBox14.value
Label29.Caption = "Po úpravì hodnot bude nutná jejich autorizace"
Label29.Visible = Label27.Visible
Label28.Visible = False

' Preplatek se vzdy musi zadavat i s datem sdeleni
CheckBox14.value = CheckBox13.value
Call CheckBox14_Click
End Sub

Private Sub CheckBox14_Click()
Call text_information_hide(Me.name)
Label14.Visible = Not CheckBox14.value
DTPicker14.Visible = CheckBox14.value

If CheckBox14.value And Label14.Caption <> "" Then
DTPicker14 = Label14.Caption
ElseIf CheckBox14.value Then
DTPicker14 = date
End If

Me.Repaint

Label27.Visible = CheckBox1.value Or CheckBox3.value Or CheckBox5.value Or CheckBox6.value Or CheckBox9.value Or CheckBox11.value Or CheckBox13.value Or CheckBox14.value
Label29.Caption = "Po úpravì hodnot bude nutná jejich autorizace"
Label29.Visible = Label27.Visible
Label28.Visible = False

' Preplatek se vzdy musi zadavat i s datem sdeleni
CheckBox13.value = CheckBox14.value
End Sub
' doplnene
Private Sub CheckBox16_Click()

Label47.Visible = Not CheckBox16.value
TextBox15.Visible = CheckBox16.value
TextBox15.text = Label47.Caption

Label28.Visible = False

End Sub

' doplnene
Private Sub CheckBox34_Click()

Label34.Visible = Not CheckBox34.value
TextBox34.Visible = CheckBox34.value
TextBox34.text = Label34.Caption

Label28.Visible = False

End Sub



' pøipravení formuláøe podle toho, co se bude ve formuláøi zadávat / co bude zobrazovat
Private Sub UserForm_Activate()
Dim sql As String
Dim control As control

If cs_HDWMN.State = 1 Then
prihlasen:

Call text_information_hide(Me.name)

For Each control In Me.Controls
 If control.Tag = "1" Then control.Visible = True
 If control.Tag = "2" Then control.Visible = False
 If control.Tag = "3" Then control.value = False
Next

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




