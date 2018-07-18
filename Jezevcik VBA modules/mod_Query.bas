Attribute VB_Name = "mod_Query"
Option Explicit


Sub query(sql As String, page As String)
' Loading data from database

Dim clmx As ColumnHeader
Dim itmX As ListItem
Dim lvSI As ListSubItem
Dim helpsum As String




On Error GoTo connection_problem

Select Case page
 
' ********************************************************************** Formuláø klient_ded - historie
Case "Klient_dedictviHistory"
   
   
   Klient_dedictvi.ListView1.ListItems.Clear
   
   
   sql = "select * from " & schema & ".v_history_ded where id_person = " & id_person & " and (1=0 "
   
   If Klient_dedictvi.CheckBox1.value Then sql = sql & " or type = 'Stav' "
   If Klient_dedictvi.CheckBox3.value Then sql = sql & " or type = 'Klient' "
   If Klient_dedictvi.CheckBox2.value Then sql = sql & " or type = 'Poznámka' "
   If Klient_dedictvi.CheckBox4.value Then sql = sql & " or type = 'Upozornìní' "
   'If Klient_dedictvi.CheckBox5.Value Then sql = sql & " or type = 'SOP' "
   
   sql = sql & ") order by dtime_note_created DESC, skp_ded_note DESC"
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
    If Not rst1.EOF Then
     rst1.MoveFirst
     For i = 1 To rst1.RecordCount
       Set itmX = Klient_dedictvi.ListView1.ListItems.Add()
       itmX.text = "" & rst1(1) & ""
       itmX.SubItems(1) = "" & rst1(2) & ""
       itmX.SubItems(2) = "" & rst1(3) & ""
       itmX.SubItems(3) = "" & rst1(4) & ""
       itmX.SubItems(4) = "" & rst1(5) & ""
       itmX.SubItems(5) = "" & rst1(6) & ""
     rst1.MoveNext
     Next i
    End If
   
   rst1.Close
   
       
    Call ZabarviRadkyTextove(Klient_dedictvi.ListView1, 2, "Stav", barvaModra)
    Call ZabarviRadkyTextove(Klient_dedictvi.ListView1, 2, "Klient", barvaZelenaSvetla)
    Call ZabarviRadkyTextove(Klient_dedictvi.ListView1, 2, "Poznámka", barvaSeda)
    Call ZabarviRadkyTextove(Klient_dedictvi.ListView1, 2, "Upozornìní", barvaRuda)
    Call ZabarviRadkyTextove(Klient_dedictvi.ListView1, 2, "SOP", barvaFialova)
   


' ********************************************************************** Formuláø klient
Case "Klient_dedictviOther"
' ********************************************************************** Formuláø klient - klient BASE
  
     Klient_dedictvi.Label1.Caption = ""
     Klient_dedictvi.Label2.Caption = ""
     Klient_dedictvi.Label3.Caption = ""
     Klient_dedictvi.Label4.Caption = ""
     Klient_dedictvi.Label5.Caption = ""
     Klient_dedictvi.Label6.Caption = ""
     Klient_dedictvi.Label7.Caption = ""
     Klient_dedictvi.Label8.Caption = ""
     Klient_dedictvi.Label91.Caption = ""
     Klient_dedictvi.Label95.Caption = "" 'RISK-403
     Klient_dedictvi.Label96.Caption = "" 'RISK-403
  
  sql = "select * from " & schema & ".v_person_base where id_person = " & id_person
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
     Klient_dedictvi.Label1.Caption = "" & rst1(0) & ""
     Klient_dedictvi.Label2.Caption = "" & rst1(1) & ""
     Klient_dedictvi.Label3.Caption = "" & rst1(2) & ""
     Klient_dedictvi.Label4.Caption = "" & rst1(3) & ""
     Klient_dedictvi.Label5.Caption = Trim(Replace("" & rst1(4) & "", "XNA", ""))
     Klient_dedictvi.Label6.Caption = Trim(Replace("" & rst1(5) & "", "XNA", ""))
     Klient_dedictvi.Label95.Caption = c.PrimaryPhone 'RISK-403
     Klient_dedictvi.Label96.Caption = c.PrimaryEmail 'RISK-403
     Klient_dedictvi.Label7.Caption = "" & rst1(8) & ""
     Klient_dedictvi.Label8.Caption = "" & rst1(9) & ""
     Klient_dedictvi.Label91.Caption = "" & rst1(12) & ""
     
        If "" & rst1(6) & "" = "Y" Then
            Klient_dedictvi.Label82.Visible = False
            Klient_dedictvi.Label83.Visible = True
        Else
            Klient_dedictvi.Label82.Visible = True
            Klient_dedictvi.Label83.Visible = False
        End If
        
        If "" & rst1(7) & "" = "Y" Then
            Klient_dedictvi.Label85.Visible = False
            Klient_dedictvi.Label84.Visible = True
        Else
            Klient_dedictvi.Label85.Visible = True
            Klient_dedictvi.Label84.Visible = False
        End If
        
        If "" & rst1(10) & "" = "Y" Then
            Klient_dedictvi.Label90.Visible = True
            Klient_dedictvi.Label89.Visible = True
            Klient_dedictvi.Label89.Caption = "" & rst1(11) & ""
        Else
            Klient_dedictvi.Label90.Visible = False
            Klient_dedictvi.Label89.Visible = False
        End If
     
    rst1.MoveNext
    Next i
   End If
   
   rst1.Close
' ********************************************************************** Formuláø klient - èísla smluv
 Klient_dedictvi.ListView2.ListItems.Clear
 
 sql = "select cislo, status_2 from " & schema & ".v_datacheck_contracts where id_person = " & id_person & " order by cislo"
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Klient_dedictvi.ListView2.ListItems.Add()
      itmX.text = "" & rst1(0) & ""
      itmX.SubItems(1) = "" & rst1(1) & ""
    rst1.MoveNext
    Next i
   End If
   
   rst1.Close

' ********************************************************************** Formuláø klient - další informace
 
 'RISK-468 - odstranìno, kvùli vymazání pùvodního listview
' Klient_dedictvi.ListView5.ListItems.Clear
' sql = "select * from " & schema & ".v_heir where id_person = " & id_person
' Set rst1 = New ADODB.Recordset
'   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'
'
'   ' rst1.MoveFirst
'    For i = 1 To rst1.Fields.count - 1
'      Set itmX = Klient_dedictvi.ListView5.ListItems.Add()
'      itmX.text = "" & rst1.Fields(i).name & ""
'    Next i
'
'    If Not rst1.EOF Then
'
'     For i = 1 To rst1.Fields.count - 1
'      Set itmX = Klient_dedictvi.ListView5.ListItems(i)
'      itmX.SubItems(1) = "" & rst1(i) & ""
'     Next i
'    End If
'
'   rst1.Close
   
 ' ********************************************************************** Formuláø klient - SOP
 
 ' RISK-468
' Klient_dedictvi.ListView4.ListItems.Clear
'
' sql = "select * from " & schema & ".MP_017_V_NOTARY where id_person = " & id_person ' MP_0017_V_NOTARY
' Set rst1 = New ADODB.Recordset
' rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'
'
'    For i = 1 To rst1.Fields.count - 1
'      Set itmX = Klient_dedictvi.ListView4.ListItems.Add()
'      itmX.text = "" & rst1.Fields(i).name & ""
'    Next i
'
'    If Not rst1.EOF Then
'
'     For i = 1 To rst1.Fields.count - 1
'      Set itmX = Klient_dedictvi.ListView4.ListItems(i)
'      itmX.SubItems(1) = "" & rst1(i) & ""
'     Next i
'    End If
'
' rst1.Close
  
 
  ' ********************************************************************** Formuláø klient - upozornìní
 
 Klient_dedictvi.ListView3.ListItems.Clear
 
 'Code status:
 
 'a - active
 'd - done
 's - storno/deleted
 
 sql = "select skp_notice, text_notice, dtime_notice_deadline, dtime_notice_deadline - trunc(sysdate) as days_left, assignee from " & schema & ".notice where code_status = 'a' and id_person = " & id_person & " order by days_left"
 Set rst1 = New ADODB.Recordset
 rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
 
 If Not rst1.EOF Then
  rst1.MoveFirst
  For i = 0 To rst1.RecordCount - 1
      Set itmX = Klient_dedictvi.ListView3.ListItems.Add()
      itmX.text = "" & rst1![skp_notice] & ""
      itmX.SubItems(1) = "" & rst1![text_notice] & ""
      itmX.SubItems(2) = "" & rst1![dtime_notice_deadline] & ""
      itmX.SubItems(3) = "" & rst1![days_left] & ""
      itmX.SubItems(4) = "" & rst1![Assignee] & ""
  rst1.MoveNext
  Next i
 End If
 rst1.Close
   

' ********************************************************************** Formuláø klient_trest - historie
Case "Klient_trestHistory"
   
   Klient_trest.ListView1.ListItems.Clear
   
   sql = "select * from " & schema & ".v_history_trest where id_person = " & id_person & " and (1=0 "
   
   If Klient_trest.CheckBox1.value Then sql = sql & " or type = 'Stav' "
   If Klient_trest.CheckBox3.value Then sql = sql & " or type = 'Klient' "
   If Klient_trest.CheckBox2.value Then sql = sql & " or type = 'Poznámka' "
   If Klient_trest.CheckBox4.value Then sql = sql & " or type = 'Upozornìní' "
   'If Klient_dedictvi.CheckBox5.Value Then sql = sql & " or type = 'SOP' "
   
   sql = sql & ") order by dtime_note_created DESC, skp_trest_note DESC"
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Klient_trest.ListView1.ListItems.Add()
      itmX.text = "" & rst1(1) & ""
      itmX.SubItems(1) = "" & rst1(2) & ""
      itmX.SubItems(2) = "" & rst1(3) & ""
      itmX.SubItems(3) = "" & rst1(4) & ""
      itmX.SubItems(4) = "" & rst1(5) & ""
      itmX.SubItems(5) = "" & rst1(6) & ""
    rst1.MoveNext
    Next i
   End If
   
   rst1.Close
   
    Call ZabarviRadkyTextove(Klient_trest.ListView1, 2, "Stav", barvaModra)
    Call ZabarviRadkyTextove(Klient_trest.ListView1, 2, "Klient", barvaZelenaSvetla)
    Call ZabarviRadkyTextove(Klient_trest.ListView1, 2, "Poznámka", barvaSeda)
    Call ZabarviRadkyTextove(Klient_trest.ListView1, 2, "Upozornìní", barvaRuda)
    Call ZabarviRadkyTextove(Klient_trest.ListView1, 2, "SOP", barvaFialova)
   


' ********************************************************************** Formuláø klient trest
Case "Klient_trestOther"
' ********************************************************************** Formuláø klient trest - klient BASE
  
     Klient_trest.Label1.Caption = ""
     Klient_trest.Label2.Caption = ""
     Klient_trest.Label3.Caption = ""
     Klient_trest.Label4.Caption = ""
     Klient_trest.Label5.Caption = ""
     Klient_trest.Label6.Caption = ""
     Klient_trest.Label7.Caption = ""
     Klient_trest.Label8.Caption = ""
     Klient_trest.Label91.Caption = ""
     Klient_trest.Label95.Caption = "" 'RISK-403
     Klient_trest.Label96.Caption = "" 'RISK-403
  
  sql = "select * from " & schema & ".v_person_base where id_person = " & id_person
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
     Klient_trest.Label1.Caption = "" & rst1(0) & ""
     Klient_trest.Label2.Caption = "" & rst1(1) & ""
     Klient_trest.Label3.Caption = "" & rst1(2) & ""
     Klient_trest.Label4.Caption = "" & rst1(3) & ""
     Klient_trest.Label5.Caption = Trim(Replace("" & rst1(4) & "", "XNA", ""))
     Klient_trest.Label6.Caption = Trim(Replace("" & rst1(5) & "", "XNA", ""))
     Klient_trest.Label95.Caption = c.PrimaryPhone 'RISK-403
     Klient_trest.Label96.Caption = c.PrimaryEmail 'RISK-403
     Klient_trest.Label7.Caption = "" & rst1(8) & ""
     Klient_trest.Label8.Caption = "" & rst1(9) & ""
     Klient_trest.Label91.Caption = "" & rst1(12) & ""
     
        If "" & rst1(6) & "" = "Y" Then
            Klient_trest.Label82.Visible = False
            Klient_trest.Label83.Visible = True
        Else
            Klient_trest.Label82.Visible = True
            Klient_trest.Label83.Visible = False
        End If
        
        If "" & rst1(7) & "" = "Y" Then
            Klient_trest.Label85.Visible = False
            Klient_trest.Label84.Visible = True
        Else
            Klient_trest.Label85.Visible = True
            Klient_trest.Label84.Visible = False
        End If
        
        If "" & rst1(10) & "" = "Y" Then
            Klient_trest.Label90.Visible = True
            Klient_trest.Label89.Visible = True
            Klient_trest.Label89.Caption = "" & rst1(11) & ""
        Else
            Klient_trest.Label90.Visible = False
            Klient_trest.Label89.Visible = False
        End If
     
    rst1.MoveNext
    Next i
   End If
   
   rst1.Close
' ********************************************************************** Formuláø klient - èísla smluv
 Klient_trest.ListView2.ListItems.Clear
 
 sql = "select cislo, status_2 from " & schema & ".v_datacheck_contracts where id_person = " & id_person & " order by cislo"
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Klient_trest.ListView2.ListItems.Add()
      itmX.text = "" & rst1(0) & ""
      itmX.SubItems(1) = "" & rst1(1) & ""
    rst1.MoveNext
    Next i
   End If
   
   rst1.Close

   
 ' ********************************************************************** Formuláø klient trest - info
 
 Klient_trest.ListView4.ListItems.Clear
 
 sql = "select * from " & schema & ".v_trest_info where id_person = " & id_person
 Set rst1 = New ADODB.Recordset
 rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
 

    For i = 1 To rst1.Fields.count - 1
      Set itmX = Klient_trest.ListView4.ListItems.Add()
      itmX.text = "" & rst1.Fields(i).name & ""
    Next i
    
    If Not rst1.EOF Then
   
     For i = 1 To rst1.Fields.count - 1
      Set itmX = Klient_trest.ListView4.ListItems(i)
      itmX.SubItems(1) = "" & rst1(i) & ""
     Next i
    End If
 
 rst1.Close
  
 
  ' ********************************************************************** Formuláø klient - upozornìní
 
 Klient_trest.ListView3.ListItems.Clear
 
 'Code status:
 
 'a - active
 'd - done
 's - storno/deleted
 
 sql = "select skp_notice, text_notice, dtime_notice_deadline, dtime_notice_deadline - trunc(sysdate) as days_left, assignee from " & schema & ".notice where code_status = 'a' and id_person = " & id_person & " order by days_left"
 Set rst1 = New ADODB.Recordset
 rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
 
 If Not rst1.EOF Then
  rst1.MoveFirst
  For i = 0 To rst1.RecordCount - 1
      Set itmX = Klient_trest.ListView3.ListItems.Add()
      itmX.text = "" & rst1(0) & ""
      itmX.SubItems(1) = "" & rst1(1) & ""
      itmX.SubItems(2) = "" & rst1(2) & ""
      itmX.SubItems(3) = "" & rst1(3) & ""
      itmX.SubItems(4) = "" & rst1![Assignee] & ""
  rst1.MoveNext
  Next i
 End If
 rst1.Close
   
   
' ********************************************************************** Formuláø klient_trest - historie
Case "Klient_insolHistory"
   
   Klient_insol.ListView1.ListItems.Clear
   
   sql = "select * from " & schema & ".v_history_ins where id_person = " & id_person & " and (1=0 "
   
   If Klient_insol.CheckBox1.value Then sql = sql & " or type = 'Stav' "
   If Klient_insol.CheckBox3.value Then sql = sql & " or type = 'Klient' "
   If Klient_insol.CheckBox2.value Then sql = sql & " or type = 'Poznámka' "
   If Klient_insol.CheckBox4.value Then sql = sql & " or type = 'Upozornìní' "
   'If Klient_dedictvi.CheckBox5.Value Then sql = sql & " or type = 'SOP' "
   
   sql = sql & ") order by dtime_note_created DESC, skp_ins_note DESC"
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Klient_insol.ListView1.ListItems.Add()
      itmX.text = "" & rst1(1) & ""
      itmX.SubItems(1) = "" & rst1(2) & ""
      itmX.SubItems(2) = "" & rst1(3) & ""
      itmX.SubItems(3) = "" & rst1(4) & ""
      itmX.SubItems(4) = "" & rst1(5) & ""
      itmX.SubItems(5) = "" & rst1(6) & ""
    rst1.MoveNext
    Next i
   End If
   
   rst1.Close
   
    Call ZabarviRadkyTextove(Klient_insol.ListView1, 2, "Stav", barvaModra)
    Call ZabarviRadkyTextove(Klient_insol.ListView1, 2, "Klient", barvaZelenaSvetla)
    Call ZabarviRadkyTextove(Klient_insol.ListView1, 2, "Poznámka", barvaSeda)
    Call ZabarviRadkyTextove(Klient_insol.ListView1, 2, "Upozornìní", barvaRuda)
    Call ZabarviRadkyTextove(Klient_insol.ListView1, 2, "SOP", barvaFialova)
   

' ********************************************************************** Formuláø klient insolvence
Case "Klient_insolOther"
' ********************************************************************** Formuláø klient insolvence - klient BASE
  
     Klient_insol.Label1.Caption = ""
     Klient_insol.Label2.Caption = ""
     Klient_insol.Label3.Caption = ""
     Klient_insol.Label4.Caption = ""
     Klient_insol.Label5.Caption = ""
     Klient_insol.Label6.Caption = ""
     Klient_insol.Label7.Caption = ""
     Klient_insol.Label8.Caption = ""
     Klient_insol.Label91.Caption = ""
     Klient_insol.Label95.Caption = "" 'RISK-403
     Klient_insol.Label96.Caption = "" 'RISK-403
  
  sql = "select * from " & schema & ".v_person_base where id_person = " & id_person
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
     Klient_insol.Label1.Caption = "" & rst1(0) & ""
     Klient_insol.Label2.Caption = "" & rst1(1) & ""
     Klient_insol.Label3.Caption = "" & rst1(2) & ""
     Klient_insol.Label4.Caption = "" & rst1(3) & ""
     Klient_insol.Label5.Caption = Trim(Replace("" & rst1(4) & "", "XNA", ""))
     Klient_insol.Label6.Caption = Trim(Replace("" & rst1(5) & "", "XNA", ""))
     Klient_insol.Label95.Caption = c.PrimaryPhone 'RISK-403
     Klient_insol.Label96.Caption = c.PrimaryEmail 'RISK-403
     Klient_insol.Label7.Caption = "" & rst1(8) & ""
     Klient_insol.Label8.Caption = "" & rst1(9) & ""
     Klient_insol.Label91.Caption = "" & rst1(12) & ""
     
        If "" & rst1(6) & "" = "Y" Then
            Klient_insol.Label82.Visible = False
            Klient_insol.Label83.Visible = True
        Else
            Klient_insol.Label82.Visible = True
            Klient_insol.Label83.Visible = False
        End If
        
        If "" & rst1(7) & "" = "Y" Then
            Klient_insol.Label85.Visible = False
            Klient_insol.Label84.Visible = True
        Else
            Klient_insol.Label85.Visible = True
            Klient_insol.Label84.Visible = False
        End If
        
        If "" & rst1(10) & "" = "Y" Then
            Klient_insol.Label90.Visible = True
            Klient_insol.Label89.Visible = True
            Klient_insol.Label89.Caption = "" & rst1(11) & ""
        Else
            Klient_insol.Label90.Visible = False
            Klient_insol.Label89.Visible = False
        End If
     
    rst1.MoveNext
    Next i
   End If
   
   rst1.Close
' ********************************************************************** Formuláø klient - èísla smluv
 Klient_insol.ListView2.ListItems.Clear
 
 sql = "select cislo, status_2 from " & schema & ".v_datacheck_contracts where id_person = " & id_person & " order by cislo"
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Klient_insol.ListView2.ListItems.Add()
      itmX.text = "" & rst1(0) & ""
      itmX.SubItems(1) = "" & rst1(1) & ""
    rst1.MoveNext
    Next i
   End If
   
   rst1.Close

   
 ' ********************************************************************** Formuláø klient trest - info
 
 Klient_insol.ListView4.ListItems.Clear
 
 sql = "select * from " & schema & ".v_ins_info where id_person = " & id_person
 Set rst1 = New ADODB.Recordset
 rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
 

    For i = 1 To rst1.Fields.count - 1
      Set itmX = Klient_insol.ListView4.ListItems.Add()
      itmX.text = "" & rst1.Fields(i).name & ""
    Next i
    
    If Not rst1.EOF Then
   
     For i = 1 To rst1.Fields.count - 1
      Set itmX = Klient_insol.ListView4.ListItems(i)
      itmX.SubItems(1) = "" & rst1(i) & ""
     Next i
    End If
 
 rst1.Close
  
 
  ' ********************************************************************** Formuláø klient - upozornìní
 
 Klient_insol.ListView3.ListItems.Clear
 
 'Code status:
 
 'a - active
 'd - done
 's - storno/deleted
 
 sql = "select skp_notice, text_notice, dtime_notice_deadline, dtime_notice_deadline - trunc(sysdate) as days_left, assignee from " & schema & ".notice where code_status = 'a' and id_person = " & id_person & " order by days_left"
 Set rst1 = New ADODB.Recordset
 rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
 
 If Not rst1.EOF Then
  rst1.MoveFirst
  For i = 0 To rst1.RecordCount - 1
      Set itmX = Klient_insol.ListView3.ListItems.Add()
      itmX.text = "" & rst1(0) & ""
      itmX.SubItems(1) = "" & rst1(1) & ""
      itmX.SubItems(2) = "" & rst1(2) & ""
      itmX.SubItems(3) = "" & rst1(3) & ""
      itmX.SubItems(4) = "" & rst1![Assignee] & ""
  rst1.MoveNext
  Next i
 End If
 rst1.Close
   

' ********************************************************************** Formuláø klient - historie
Case "KlientHistory"
   
   
   Klient.ListView1.ListItems.Clear
   
   
   sql = "select * from " & schema & ".v_history where id_person = " & id_person & " and (1=0 "
   
   If Klient.CheckBox1.value Then sql = sql & " or type = 'Stav' "
   If Klient.CheckBox3.value Then sql = sql & " or type = 'Klient' "
   If Klient.CheckBox2.value Then sql = sql & " or type = 'Poznámka' "
   If Klient.CheckBox4.value Then sql = sql & " or type = 'Upozornìní' "
   If Klient.CheckBox5.value Then sql = sql & " or type = 'SOP' "
   
   sql = sql & ") order by dtime_note_created DESC, skp_epr_note DESC"
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Klient.ListView1.ListItems.Add()
      itmX.text = "" & rst1(1) & ""
      itmX.SubItems(1) = "" & rst1(2) & ""
      itmX.SubItems(2) = "" & rst1(3) & ""
      itmX.SubItems(3) = "" & rst1(4) & ""
      itmX.SubItems(4) = "" & rst1(5) & ""
      itmX.SubItems(5) = "" & rst1(6) & ""
    rst1.MoveNext
    Next i
    
    
    
   End If
   rst1.Close
   
   
    Call ZabarviRadkyTextove(Klient.ListView1, 2, "Stav", barvaModra)
    Call ZabarviRadkyTextove(Klient.ListView1, 2, "Klient", barvaZelenaSvetla)
    Call ZabarviRadkyTextove(Klient.ListView1, 2, "Poznámka", barvaSeda)
    Call ZabarviRadkyTextove(Klient.ListView1, 2, "Upozornìní", barvaRuda)
    Call ZabarviRadkyTextove(Klient.ListView1, 2, "SOP", barvaFialova)
    


' ********************************************************************** Formuláø klient
Case "KlientOther"
' ********************************************************************** Formuláø klient - klient BASE
Klient.Label1.Caption = ""
     Klient.Label2.Caption = ""
     Klient.Label3.Caption = ""
     Klient.Label4.Caption = ""
     Klient.Label5.Caption = ""
     Klient.Label6.Caption = ""
     Klient.Label91.Caption = ""
     'Klient.Label96.Caption = ""
     'Klient.Label97.Caption = ""
     Klient.Label87.Caption = ""
     Klient.Label96.Caption = "" 'RISK-403
     Klient.Label98.Caption = "" 'RISK-403
  
  
  sql = "select * from " & schema & ".v_person_base where id_person = " & id_person
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
     Klient.Label1.Caption = "" & rst1(0) & ""
     Klient.Label2.Caption = "" & rst1(1) & ""
     Klient.Label3.Caption = "" & rst1(2) & ""
     Klient.Label4.Caption = "" & rst1(3) & ""
     Klient.Label5.Caption = Trim(Replace("" & c.Address_P & "", "XNA", ""))
     Klient.Label6.Caption = Trim(Replace("" & c.Address_C & "", "XNA", ""))
     Klient.Label96.Caption = c.PrimaryPhone 'RISK-403
     Klient.Label98.Caption = c.PrimaryEmail 'RISK-403
     Klient.Label91.Caption = "" & rst1(12) & ""
     
        If "" & rst1(6) & "" = "Y" Then
            Klient.Label82.Visible = False
            Klient.Label83.Visible = True
        Else
            Klient.Label82.Visible = True
            Klient.Label83.Visible = False
        End If
        
        If "" & rst1(7) & "" = "Y" Then
            Klient.Label85.Visible = False
            Klient.Label84.Visible = True
        Else
            Klient.Label85.Visible = True
            Klient.Label84.Visible = False
        End If
        
        If "" & rst1(10) & "" = "Y" Then
            Klient.Label86.Visible = True
            Klient.Label87.Visible = True
            Klient.Label87.Caption = "" & rst1(11) & ""
        Else
            Klient.Label87.Visible = False
            Klient.Label86.Visible = False
        End If
     
    rst1.MoveNext
    Next i
   End If
   
   rst1.Close
' ********************************************************************** Formuláø klient - èísla smluv
 Klient.ListView2.ListItems.Clear
 
 'sql = "select id_person, code_loan_number from " & schema & ".loan where id_person = " & id_person & " order by code_loan_number" ' FA: 19.3.2018 nahrazena tabulka DEAL za LOAN
 sql = "SELECT a.ID_PERSON, a.CISLO " & _
       "FROM " & schema & ".v_datacheck_contracts a " & _
       "WHERE id_person = " & id_person & " AND a.TYP_SMLOUVY IN ('HU', 'HR', 'KTK') " & _
       "ORDER BY a.TYP_SMLOUVY"
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Klient.ListView2.ListItems.Add()
      itmX.text = "" & rst1![cislo] & ""
    rst1.MoveNext
    Next i
   End If
   
   rst1.Close

' ********************************************************************** Formuláø klient - Pravomocnì pøiznaná pohledávka
 
 Klient.ListView6.ListItems.Clear
 sql = "select * from " & schema & ".v_pm_infobox where id_person = " & id_person
 Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
  

   ' rst1.MoveFirst
    For i = 1 To rst1.Fields.count - 1
      Set itmX = Klient.ListView6.ListItems.Add()
      itmX.text = "" & rst1.Fields(i).name & ""
    Next i
    
    If Not rst1.EOF Then
   
     For i = 1 To rst1.Fields.count - 1
     
     If (i = 5 Or i = 6) And IsNull(rst1(i)) = False Then
         Set itmX = Klient.ListView6.ListItems(i)
         itmX.SubItems(1) = "" & format(rst1(i), "### ### ##0.00") & " Kè"
        Else
         Set itmX = Klient.ListView6.ListItems(i)
         itmX.SubItems(1) = "" & rst1(i) & ""
     End If
     
     If i = 3 And rst1(i) <> date Then
       Set itmX = Klient.ListView6.ListItems(i)
       itmX.ListSubItems(1).ForeColor = barvaRuda
     End If
     Next i
    End If
   
   rst1.Close
   
' ********************************************************************** Formuláø klient - další informace
 
  Klient.ListView5.ListItems.Clear
 
 'RISK-442


 sql = "select * from " & schema & ".v_information where id_person = " & id_person
 Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic


      Set itmX = Klient.ListView5.ListItems.Add()
      itmX.text = "" & rst1.Fields(1).name & ""

      Set itmX = Klient.ListView5.ListItems.Add()
      'itmX.text = "Suma splátek akt. úvìrù pøi zesplatnìní"
      itmX.text = "Souèet pùvodních splátek"

    For i = 2 To rst1.Fields.count - 1
      Set itmX = Klient.ListView5.ListItems.Add()
      itmX.text = "" & rst1.Fields(i).name & ""
    Next i

    If Not rst1.EOF Then

      Set itmX = Klient.ListView5.ListItems(1)
      itmX.SubItems(1) = "" & rst1(1) & ""

      If c.GetAmt_InstalmentBeforePayoff = 0 Then
      helpsum = ""
      Else: helpsum = "" & c.GetAmt_InstalmentBeforePayoff & " Kè"
      End If

      Set itmX = Klient.ListView5.ListItems(2)
      itmX.SubItems(1) = helpsum

     For i = 2 To rst1.Fields.count - 1
      Set itmX = Klient.ListView5.ListItems(i + 1)
      itmX.SubItems(1) = "" & rst1(i) & ""
     Next i
    End If

   rst1.Close
   
 ' ********************************************************************** Formuláø klient - SOP
 
 Klient.ListView4.ListItems.Clear
    
    Call c.SOP.Vytvor_listview_prehled_SOPu(Klient.ListView4)
    Klient.CommandButton20.Enabled = c.SOP.blnPovolitAutorizaci
    
 
    sql = "select * from " & schema & ".v_sop_balance where id_person = " & id_person
    Set rst1 = New ADODB.Recordset
    rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
        If Not rst1.EOF Then
         Klient.Label77.Caption = rst1(4) & " Kè"
         Klient.Label78.Caption = rst1(2) & " Kè"
        End If
     
    rst1.Close
 
 

  ' ********************************************************************** Formuláø klient - upozornìní
 
 Klient.ListView3.ListItems.Clear
 
 'Code status:
 
 'a - active
 'd - done
 's - storno/deleted
 
 sql = "select skp_notice, text_notice, dtime_notice_deadline, dtime_notice_deadline - trunc(sysdate) as days_left, assignee from " & schema & ".notice where code_status = 'a' and id_person = " & id_person & " order by days_left"
 Set rst1 = New ADODB.Recordset
 rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
 
 If Not rst1.EOF Then
  rst1.MoveFirst
  For i = 0 To rst1.RecordCount - 1
      Set itmX = Klient.ListView3.ListItems.Add()
      itmX.text = "" & rst1(0) & ""
      itmX.SubItems(1) = "" & rst1(1) & ""
      itmX.SubItems(2) = "" & rst1(2) & ""
      itmX.SubItems(3) = "" & rst1(3) & ""
      itmX.SubItems(4) = "" & rst1![Assignee] & ""
  rst1.MoveNext
  Next i
 End If
 rst1.Close
   
  ' ********************************************************************** Formuláø klient - upozorneni, pokud se lisi datum zesplatneni na dwh a OBS, chyba HD 21459, nacita se v proc pb_036_uznani_dluhu
 sql = "select id_person from " & schema & ".wrong_dp where id_person = " & id_person
 Set rst1 = New ADODB.Recordset
 rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
 
 If Not rst1.EOF Then
            Klient.Label93.Visible = True
            Klient.Label92.Visible = True
            Else
            Klient.Label93.Visible = False
            Klient.Label92.Visible = False
           
 End If
 rst1.Close
 
 'Formuláø klient - zjistime, zda se jedna o debetare, propise se do hlavniho labelu na karte klienta
 
sql = "select SKP_PERSON_STATUS from " & schema & ".person_status_data where id_person = " & id_person & " and flag_deleted = 'N' and flag_valid = 'Y'"  ' RISK-985 SKP_PERSON_STATUS
 
 Set rst1 = New ADODB.Recordset
 rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
 
 If Not rst1.EOF Then

    If rst1(0) = 3 Then
        Klient.Label26.Caption = "Detail klienta (nalézací øízení) - DEBETÁØ"
    Else
        Klient.Label26.Caption = "Detail klienta (nalézací øízení)"
    End If
    
 End If
 rst1.Close
  
   
   
' ********************************************************************** Formuláø zobrazení poznámky
Case "Note"
 Set rst1 = New ADODB.Recordset
 rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
 
 Klient.TextBox1.text = ""
 
 If Not rst1.EOF Then
  Klient.TextBox1.text = "" & rst1(0) & ""
 End If
 rst1.Close
' ********************************************************************** Formuláø zobrazení poznámky

Case "Note_exe"
 Set rst1 = New ADODB.Recordset
 rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
 
 Klient_exe.TextBox1.text = ""
 
 If Not rst1.EOF Then
  Klient_exe.TextBox1.text = "" & rst1(0) & ""
 End If
 rst1.Close
 
' ********************************************************************** Formuláø zobrazení poznámky

Case "Note_ded"
 Set rst1 = New ADODB.Recordset
 rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
 
 Klient_dedictvi.TextBox1.text = ""
 
 If Not rst1.EOF Then
  Klient_dedictvi.TextBox1.text = "" & rst1(0) & ""
 End If
 rst1.Close
 
 ' ********************************************************************** Formuláø zobrazení poznámky
Case "Note_trest"
 Set rst1 = New ADODB.Recordset
 rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
 
 Klient_trest.TextBox1.text = ""
 
 If Not rst1.EOF Then
  Klient_trest.TextBox1.text = "" & rst1(0) & ""
 End If
 rst1.Close
 
  ' ********************************************************************** Formuláø zobrazení poznámky
Case "Note_insol"
 Set rst1 = New ADODB.Recordset
 rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
 
 Klient_insol.TextBox1.text = ""
 
 If Not rst1.EOF Then
  Klient_insol.TextBox1.text = "" & rst1(0) & ""
 End If
 rst1.Close
  
' ********************************************************************** Formuláø žaloby
Case "Zaloby"
   
   Zaloby.ListView1.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   'asss = InputBox("", "", sql)
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    rst1.Sort = "ID_PERSON, DTIME_VALID_FROM DESC"
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Zaloby.ListView1.ListItems.Add()
      itmX.text = "" & rst1(0) & ""
      itmX.SubItems(1) = "" & rst1(1) & ""
      itmX.SubItems(2) = "" & rst1(2) & ""
      itmX.SubItems(3) = "" & rst1(3) & ""
      itmX.SubItems(4) = "" & rst1(4) & ""
      itmX.SubItems(5) = "" & rst1(5) & ""
      itmX.SubItems(6) = "" & rst1(9) & ""
      itmX.SubItems(7) = "" & rst1(6) & ""
      itmX.SubItems(8) = "" & rst1(7) & ""
      itmX.SubItems(9) = "" & rst1(8) & ""
      itmX.SubItems(10) = "" & rst1(10) & ""
    rst1.MoveNext
    Next i
   End If
   
   rst1.Close
  
  
  ' ********************************************************************** Formuláø žaloby: èíselník
Case "ZalobyEnumeration"
      
  
   
Call query("", "ZalobyVhodnyCiselnik")



  ' ********************************************************************** Formuláø žaloby: èíselník
   
Case "ZalobyVhodnyCiselnik"

'If Zaloby.ComboBox2.ListCount > 0 Then

sql = "select name_epr_status from " & schema & ".epr_status where dtime_valid_to > sysdate and skp_process = 3 order by order_item nulls last"

  Zaloby.ComboBox1.Clear

   Zaloby.ComboBox1.AddItem "Vše"

   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Zaloby.ComboBox1.AddItem "" & rst1(0) & ""
    rst1.MoveNext
    Next i
   End If

    rst1.Close

   Zaloby.ComboBox1.ListIndex = 0


'End If

' ********************************************************************** Formuláø dedictvi
''' zakomentovano RISK-468
'''Case "Dedictvi"

' ********************************************************************** Formuláø trest
Case "Trest"

   Trest.ListView1.ListItems.Clear

   Set rst1 = New ADODB.Recordset
   'asss = InputBox("", "", sql)
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.Sort = "ID_PERSON, DTIME_VALID_FROM DESC"
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Trest.ListView1.ListItems.Add()
      itmX.text = "" & rst1(0) & ""
      itmX.SubItems(1) = "" & rst1(1) & ""
      itmX.SubItems(2) = "" & rst1(2) & ""
      itmX.SubItems(3) = "" & rst1(3) & ""
      itmX.SubItems(4) = "" & rst1(4) & ""
      itmX.SubItems(5) = "" & rst1(5) & ""
      itmX.SubItems(6) = "" & rst1(6) & ""
      itmX.SubItems(7) = "" & rst1(7) & ""
      itmX.SubItems(8) = "" & rst1(8) & ""
      itmX.SubItems(9) = "" & rst1(9) & ""
      itmX.SubItems(10) = "" & rst1(10) & ""
    rst1.MoveNext
    Next i
   End If

   rst1.Close


    Call ZabarviRadkyTextove(Trest.ListView1, 0, "")


  ' ********************************************************************** Formuláø dedictvi: èíselník
Case "DedictviEnumeration"

sql = "select name_ded_status from " & schema & ".ded_status where dtime_valid_to > sysdate and skp_process = 1 order by order_item nulls last"
      
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   Dedictvi.ComboBox1.Clear

   Dedictvi.ComboBox1.AddItem "Vše"
'   dedictvi.ComboBox2.Column(1, 0) = 0

   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then

    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Dedictvi.ComboBox1.AddItem "" & rst1(0) & ""
'      Zaloby.ComboBox2.Column(1, i) = rst1(0)
    rst1.MoveNext
    Next i
   End If

   rst1.Close


   Dedictvi.ComboBox1.ListIndex = 0
   
   
  ' ********************************************************************** Formuláø Trest: èíselník
Case "TrestEnumeration"

sql = "select name_trest_status from " & schema & ".trest_status where dtime_valid_to > sysdate and skp_process = 5 order by order_item nulls last"
      
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   Trest.ComboBox1.Clear

   Trest.ComboBox1.AddItem "Vše"
'   dedictvi.ComboBox2.Column(1, 0) = 0

   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then

    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Trest.ComboBox1.AddItem "" & rst1(0) & ""
'      Zaloby.ComboBox2.Column(1, i) = rst1(0)
    rst1.MoveNext
    Next i
   End If

   rst1.Close


   Trest.ComboBox1.ListIndex = 0

  ' ********************************************************************** Formuláø Trest: èíselník
Case "InsolEnumeration"

sql = "select name_ins_status from " & schema & ".ins_status where dtime_valid_to > sysdate and skp_process = 4 order by order_item nulls last"
      
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   Insol.ComboBox1.Clear

   Insol.ComboBox1.AddItem "Vše"
'   dedictvi.ComboBox2.Column(1, 0) = 0

   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then

    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Insol.ComboBox1.AddItem "" & rst1(0) & ""
'      Zaloby.ComboBox2.Column(1, i) = rst1(0)
    rst1.MoveNext
    Next i
   End If

   rst1.Close


   Insol.ComboBox1.ListIndex = 0
   
  ' ********************************************************************** Formuláø žaloby: souhrn
Case "ZalobySouhrn"
   
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   Application.EnableEvents = False
   Application.ScreenUpdating = False
   Application.DisplayAlerts = False
   
   cfg.wb_Jez.Activate
   
   Worksheets(page).Visible = True
   Worksheets(page).Activate
   
   
   Range("A2:C2").Select
   Range(Selection, Selection.End(xlDown)).Select
   Selection.Clear
   
   ActiveSheet.ListObjects("TabulkaZalobySouhrn").Resize Range("$A$1:$C$2")
   
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Range("A" & i + 1) = "" & rst1(0) & ""
      Range("B" & i + 1) = "" & rst1(1) & ""
      Range("C" & i + 1) = "" & rst1(2) & ""
    rst1.MoveNext
    Next i
   End If
   

   
   rst1.Close
   
  
  ' ********************************************************************** Formuláø Zesplatneni: souhrn
Case "ZesplatneniSouhrn"
   
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   Application.EnableEvents = False
   Application.ScreenUpdating = False
   Application.DisplayAlerts = False
   
   cfg.wb_Jez.Activate
   
   
   Worksheets(page).Visible = True
   Worksheets(page).Activate
   
   
   Range("A2:C2").Select
   Range(Selection, Selection.End(xlDown)).Select
   Selection.Clear
   
   ActiveSheet.ListObjects("TabulkaZesplatneniSouhrn").Resize Range("$A$1:$C$2")
   
   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Range("A" & i + 1) = "" & rst1(0) & ""
      Range("B" & i + 1) = "" & rst1(1) & ""
      Range("C" & i + 1) = "" & rst1(2) & ""
    rst1.MoveNext
    Next i
   End If
   

   
   rst1.Close
  

  
' ********************************************************************** Formuláø souhrn - poèet
'FA RISK-599: commented out 24.5.2017 by Filip Antol v ramci zmeny celej karty Souhrn
'Case "SouhrnPocet"
'
'   Souhrn.ListView1.ListItems.Clear
'
'   Set rst1 = New ADODB.Recordset
'   'asss = InputBox("", "", sql)
'   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'
'   If Not rst1.EOF Then
'
'    rst1.MoveFirst
'    For i = 1 To rst1.RecordCount
'      Set itmX = Souhrn.ListView1.ListItems.Add()
'      itmX.text = "" & rst1(0) & ""
'      itmX.SubItems(1) = "" & rst1(1) & ""
'      itmX.SubItems(2) = "" & rst1(2) & ""
'    rst1.MoveNext
'    Next i
'   End If
'
'   rst1.Close
  
  
' ********************************************************************** Formuláø souhrn - interval
'FA RISK-599: commented out 24.5.2017 by Filip Antol v ramci zmeny celej karty Souhrn
'Case "SouhrnDoba"
'
'   Souhrn.ListView2.ListItems.Clear
'
'   Set rst1 = New ADODB.Recordset
'   'asss = InputBox("", "", sql)
'   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
'
'   If Not rst1.EOF Then
'
'    rst1.MoveFirst
'    For i = 1 To rst1.RecordCount
'      Set itmX = Souhrn.ListView2.ListItems.Add()
'      itmX.text = "" & rst1(0) & ""
'      itmX.SubItems(1) = "" & rst1(1) & ""
'      itmX.SubItems(2) = "" & rst1(2) & ""
'    rst1.MoveNext
'    Next i
'   End If
'
'   rst1.Close
  
' ********************************************************************** Formuláø souhrn - upozornìní
Case "SouhrnUpozorneni"
   
   Souhrn.ListView3.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   'asss = InputBox("", "", sql)
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then

    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Souhrn.ListView3.ListItems.Add()
      itmX.text = "" & rst1![id_person] & ""
      itmX.SubItems(1) = "" & rst1![CUID] & ""
      itmX.SubItems(2) = "" & rst1![name] & ""
      itmX.SubItems(3) = "" & rst1![text_notice] & ""
      itmX.SubItems(4) = "" & rst1![dtime_notice_deadline] & ""
      itmX.SubItems(5) = "" & rst1![days_left] & ""
      itmX.SubItems(7) = "" & rst1![Assignee] & ""
      itmX.SubItems(8) = "" & rst1![skp_notice] & ""
      'RISK-130
      'Novì je požadováno, aby se upozornìní otevíralo na zadané kartì.
      'Proto byla rozšíøena struktura tabulky notice o klíè SKP_PROCESS. Pomocí toho se pak pøi prokliku na souhrnu dostanu na kartu, kde bylo upozornìní zadané.
      itmX.SubItems(6) = "" & rst1![SKP_PROCESS] & ""
    rst1.MoveNext
    Next i
   End If
   
   
   rst1.Close
    
   ' Oèervenìní øádkù, které mají deadline k dnešnímu datu a starší
   Call ZabarviRadky(Souhrn.ListView3, 5)
    
    
' ********************************************************************** Formuláø souhrn - autorizace
Case "SouhrnAutorizace"
   
   Souhrn.ListView4.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   'asss = InputBox("", "", sql)
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
       
       If Not rst1.EOF Then
    
            rst1.MoveFirst
            For i = 1 To rst1.RecordCount
                Set itmX = Souhrn.ListView4.ListItems.Add()
                itmX.text = "" & rst1![id_person] & ""
                itmX.SubItems(1) = "" & rst1![CUID] & ""
                itmX.SubItems(2) = "" & rst1![name] & ""
                itmX.SubItems(3) = "" & rst1![text] & ""
                If Not IsNull(rst1![okresni_soud]) Then
                    If InStr(rst1![okresni_soud], "soud") > 0 Then
                        itmX.SubItems(4) = "" & Trim(Split(rst1![okresni_soud], "soud")(1)) & "" ' rst1![okresni_soud] & ""
                    End If
                End If
                itmX.SubItems(5) = "" & rst1![castka] & ""
                itmX.SubItems(6) = "" & rst1![DATUM_K_CASTCE] & ""
            
                rst1.MoveNext
            Next i
       End If
   
   rst1.Close
    
' ********************************************************************** Formuláø souhrn - kontrolaPM
' zakomentovane a pridane nizsie v skripte, pretoze tento query uz bol prilis dlhy (error pri spusteni excelu)
' --> query1

' ********************************************************************** Formuláø souhrn - SOP
Case "SouhrnSOP"
   
   
   Set rst1 = New ADODB.Recordset
   'asss = InputBox("", "", sql)
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
      Souhrn.Label33.Caption = "" & format(rst1(0), "### ### ##0.00") & " Kè"
      Souhrn.Label36.Caption = "" & format(rst1(1), "### ### ##0.00") & " Kè"
      Souhrn.Label37.Caption = "" & format(rst1(2), "### ### ##0.00") & " Kè"
 
   End If
   
   rst1.Close
  
  
' ********************************************************************** Formuláø souhrn - autorizace
Case "Stavy"

'RISK-468 RISK-949
If Stavy.OptionButton1.value = True Then
    skp_proces = "1,2,3,4,5,6"
ElseIf Stavy.OptionButton2.value = True Then
    skp_proces = "3"
ElseIf Stavy.OptionButton3.value = True Then
    skp_proces = "1"
ElseIf Stavy.OptionButton4.value = True Then
    skp_proces = "5"
ElseIf Stavy.OptionButton5.value = True Then
    skp_proces = "4"
ElseIf Stavy.OptionButton6.value = True Then
    skp_proces = "2"
ElseIf Stavy.OptionButton7.value = True Then
    skp_proces = "6"
End If

   
       If Not Stavy.CheckBox1.value Then sql = sql & " and dtime_valid_to > sysdate "

        sql = sql & " and skp_process in (" & skp_proces & ") "
        
        sql = sql & " order by name_process, skp_epr_status"
   
   Stavy.ListView1.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   'asss = InputBox("", "", sql)
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then

    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Stavy.ListView1.ListItems.Add()
      itmX.text = "" & rst1(0) & ""
      itmX.SubItems(1) = "" & rst1(1) & ""
      itmX.SubItems(2) = "" & rst1(2) & ""
      itmX.SubItems(3) = "" & rst1(3) & ""
      itmX.SubItems(4) = "" & rst1(4) & ""
      itmX.SubItems(5) = "" & rst1(5) & ""
      itmX.SubItems(6) = "" & rst1(6) & ""
    rst1.MoveNext
    Next i
   End If
   
   rst1.Close
   
' ********************************************************************** Formuláø Stavy - proces
Case "StavyProces"
   
   
Stavy_update.ComboBox1.Clear
   
   Set rst1 = New ADODB.Recordset
   'asss = InputBox("", "", sql)
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then

    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Stavy_update.ComboBox1.AddItem "" & rst1(1) & ""
      Stavy_update.ComboBox1.Column(1, i - 1) = rst1(0)
    rst1.MoveNext
    Next i
   End If
   
   rst1.Close
   
  
' ********************************************************************** Data pro export do PDF
Case "KlientExport"
   

   
   Set rst1 = New ADODB.Recordset
   'asss = InputBox("", "", sql)
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    Range("D17") = "" & rst1(0) & ""
    Range("D18") = "" & rst1(1) & ""
   Else
    Range("D17") = "Žádný status"
   End If
   
   rst1.Close

' ********************************************************************** Data pro export do PDF
Case "Klient_dedictviExport"
   
   Set rst1 = New ADODB.Recordset
   'asss = InputBox("", "", sql)
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    Range("D17") = "" & rst1(0) & ""
    Range("D18") = "" & rst1(1) & ""
   Else
    Range("D17") = "Žádný status"
   End If
   
   rst1.Close
   
' ********************************************************************** Data pro export do PDF
Case "Klient_trestExport"
      
   Set rst1 = New ADODB.Recordset
   'asss = InputBox("", "", sql)
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    Range("D17") = "" & rst1(0) & ""
    Range("D18") = "" & rst1(1) & ""
   Else
    Range("D17") = "Žádný status"
   End If
   
   rst1.Close

' ********************************************************************** Hlavní menu - upozornìní
Case "Main"

   
   Set rst1 = New ADODB.Recordset
   'asss = InputBox("", "", sql)
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    Main.Label35.Caption = "" & rst1![pocet] & ""
    Main.Label36.Caption = "" & rst1![po] & ""
    Main.Label39.Caption = "" & rst1![auth] & ""
    Main.Label_Sonik_upo.Caption = "" & rst1![SONIK] & ""
   End If
   
   rst1.Close
   
   
   ' Pokud  nejsou žádné údaje pro zobrazení, upraví se vzhled a funkcionalita frame
   If Main.Label35.Caption = "0" And Main.Label39.Caption = "0" Then
    Main.Frame2.Enabled = False
    Main.Frame2.MousePointer = fmMousePointerDefault
    Main.Frame2.BackColor = barvaZelenaAirBank
   Else
    Main.Frame2.Enabled = True
    Main.Frame2.MousePointer = fmMousePointerCustom
    Main.Frame2.BackColor = barvaRuda
   End If


' ********************************************************************** DataCheck
' **********************************************************************  - Naplnìní èísel úvìrù/úètù
Case "DataCheck"
   
    With DataCheck.ListView1.ColumnHeaders
        .Clear
        Call .Add(, , "Èíslo smlouvy/úètu", 70)
        Call .Add(, , "Status", 70)
        Call .Add(, , "Typ", 60)
    End With
   
   DataCheck.ListView1.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then

    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = DataCheck.ListView1.ListItems.Add()
      itmX.text = "" & rst1(0) & ""
      itmX.SubItems(1) = "" & rst1(1) & ""
      itmX.SubItems(2) = "" & rst1(2) & ""
    rst1.MoveNext
    Next i
   End If
   
   rst1.Close
   
     DataCheck.Label73.Caption = ""
   
   sql = "select dtime_modified from " & schema & ".v_datacheck_debit where id_person = " & id_person
   
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then

        DataCheck.Label73.Caption = "" & rst1(0)
    
   End If
   
   rst1.Close
   
   
    sql = "select * from " & schema & ".v_datacheck_pohl where id_person = " & id_person
   
    Set rst1 = New ADODB.Recordset
    rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
        If Not rst1.EOF Then
        
            DataCheck.Label81.Caption = "" & format(rst1(1), "### ### ##0.00") & ""
            DataCheck.Label82.Caption = "" & format(rst1(2), "### ### ##0.00") & ""
            DataCheck.Label83.Caption = "" & format(rst1(3), "### ### ##0.00") & ""
            DataCheck.Label84.Caption = "" & format(rst1(4), "### ### ##0.00") & ""
            DataCheck.Label85.Caption = "" & format(rst1(5), "### ### ##0.00") & ""
         
        End If
        
    rst1.Close
   

' **********************************************************************  - Debit: výše debetního zùstatku
Case "DataCheckDebit"
   
   Set rst1 = New ADODB.Recordset
   'asss = InputBox("", "", sql)
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
      DataCheck.TextBox6.text = "" & rst1(0) & ""
      DataCheck.Label68.Caption = "" & rst1(1) & ""
   End If
   
   rst1.Close


' **********************************************************************  - Credit: intervaly s daty
Case "DataCheckCreditI"
    
    With DataCheck.ListView2.ColumnHeaders
        .Clear
        .Add , , "Platnost od (OÚ)", 80
        .Add , , "Platnost do (OÚ)", 80
        .Add , , "Výše OÚ", 80
        .Add , , "Platnost od (SÚ)", 80
        .Add , , "Platnost do (SÚ)", 80
        .Add , , "Výše SÚ", 100
        .Add , , "Výše jistiny", 100
    End With
    
    With DataCheck.ListView2
        .ListItems.Clear
        .Height = 84
    End With
   
   Set rst1 = New ADODB.Recordset
   'asss = InputBox("", "", sql)
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then

    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = DataCheck.ListView2.ListItems.Add()
      itmX.text = "" & rst1(2) & ""
      itmX.SubItems(1) = "" & rst1(3) & ""
      itmX.SubItems(2) = "" & rst1(4) & ""
      itmX.SubItems(3) = "" & rst1(5) & ""
      itmX.SubItems(4) = "" & rst1(6) & ""
      itmX.SubItems(5) = "" & rst1(7) & ""
      itmX.SubItems(6) = "" & rst1(8) & ""
    rst1.MoveNext
    Next i
   End If
   rst1.Close

End Select

Exit Sub

connection_problem:

If Mid(Err.Description, 1, 9) = "ORA-03114" Then
err_connection = True
Else
 MsgBox Err.Description
End If

End Sub






''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''  XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX  XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX  ''''''''''''''''''''''''
''''''''''''''''''''''''  XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX  XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX  ''''''''''''''''''''''''
''''''''''''''''''''''''  XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX  XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX  ''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



' puvodni funkce query byla prilis velka, proto nova funkce
Sub query1(sql As String, page As String)

Dim clmx As ColumnHeader
Dim itmX As ListItem
Dim lvSI As ListSubItem
Dim helpsum As String


On Error GoTo connection_problem

Select Case page

' ********************************************************************** Formuláø souhrn - kontrolaPM
Case "SouhrnKontrolaPM"

   Souhrn.LV_KontrolaPM.ListItems.Clear

   Set rst1 = New ADODB.Recordset
   'asss = InputBox("", "", sql)
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then

    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Souhrn.LV_KontrolaPM.ListItems.Add()
      itmX.text = "" & rst1![id_person] & ""
      itmX.SubItems(1) = "" & rst1![CUID] & ""
      itmX.SubItems(2) = "" & rst1![name] & ""
      itmX.SubItems(3) = "" & rst1![text_notice] & ""
      itmX.SubItems(4) = "" & rst1![dtime_notice_deadline] & ""
      itmX.SubItems(5) = "" & rst1![days_left] & ""
    rst1.MoveNext
    Next i
   End If


   rst1.Close

   ' Oèervenìní øádkù, které mají deadline k dnešnímu datu a starší
   Call ZabarviRadky(Souhrn.LV_KontrolaPM, 5)


' ********************************************************************** Formuláø souhrn - Sonik
Case "SouhrnUpozorneni_Sonik"

   Souhrn.LV_Sonik.ListItems.Clear

   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then

    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Souhrn.LV_Sonik.ListItems.Add()
      itmX.text = "" & rst1![skp_data_message_notice] & ""
      itmX.SubItems(1) = "" & rst1![SKP_DATA_MESSAGE] & ""
      itmX.SubItems(2) = "" & rst1![cuids] & ""
      itmX.SubItems(3) = "" & rst1![jmena] & ""
      itmX.SubItems(4) = "" & rst1![text_notice] & ""
      itmX.SubItems(5) = "" & rst1![dtime_notice_deadline] & ""
      itmX.SubItems(6) = "" & rst1![days_left] & ""
      itmX.SubItems(7) = "" & rst1![data_message_number] & ""
    rst1.MoveNext
    Next i
   End If


   rst1.Close

   ' Oèervenìní øádkù, které mají deadline k dnešnímu datu a starší
   Call ZabarviRadky(Souhrn.LV_Sonik, 6)

  
' **********************************************************************  - Datacheck: vypis z uct
Case "DataCheckDebit"

    With DataCheck.ListView2.ColumnHeaders
        .Clear
        Call .Add(, , "Datum vygenerování transakce ", 100)
        Call .Add(, , "Datum zaúètování", 100)
        Call .Add(, , "Název transakce", 250)
        Call .Add(, , "Výše transakce", 100)
    End With
    
    With DataCheck.ListView2
        .ListItems.Clear
        .Height = 155
    End With
    
    Set rst1 = New ADODB.Recordset
    rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
       
       
        If Not rst1.EOF Then
       
    
       
            For i = 1 To rst1.RecordCount
                Set itmX = DataCheck.ListView2.ListItems.Add()
                itmX.text = "" & rst1(0) & ""
                itmX.SubItems(1) = "" & rst1(1) & ""
                itmX.SubItems(2) = "" & rst1(3) & ""
                itmX.SubItems(3) = "" & rst1(2) & ""
        
                rst1.MoveNext
            Next i
       
        End If
    rst1.Close

' **********************************************************************  - Credit: Suma platby
    
  
' **********************************************************************  - Credit: aktuální data
Case "DataCheckCreditA"
   Dim SKP As Double, sql2 As String
   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   If Not rst1.EOF Then
      DataCheck.TextBox1.text = "" & rst1(2) & ""
      DataCheck.TextBox2.text = "" & rst1(3) & ""
      DataCheck.TextBox3.text = "" & rst1(4) & ""
      DataCheck.TextBox4.text = "" & rst1(5) & ""
      DataCheck.TextBox5.text = "" & rst1(6) & ""
      DataCheck.TextBox10.text = "" & rst1(7) & ""
      DataCheck.TextBox11.text = rst1(5) + rst1(4) + rst1(3) + rst1(6) 'soucet pro uver
      SKP = rst1(0)
   End If
   rst1.Close

    If SKP > 0 Then
   Set rst2 = New ADODB.Recordset
   sql2 = "select sum(amt_balance_principal) + sum(kapitalizovnay_urok) + sum(ou_total)+sum(su_total) from " & schema & ".v_datacheck_credit_A where id_person = " & SKP & " group by id_person "
   
   rst2.Open sql2, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst2.EOF Then
      DataCheck.TextBox12.text = "" & rst2(0) & ""
   End If
   End If
' **********************************************************************  - Credit: Suma platby
  
 Case "DataCheckPlatbyCelkem"
   
   Set rst1 = New ADODB.Recordset
   'asss = InputBox("", "", sql)
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
      DataCheck.Label29.Caption = "Platba dne " & DataCheck.ListView3.SelectedItem & " byla ve výši " & rst1(3) & "Kè"
   End If
   
   rst1.Close
  
' **********************************************************************  - Credit: Platby
Case "DataCheckPlatby"
   
    With DataCheck.ListView3.ColumnHeaders
        .Clear
        Call .Add(, , "Datum platby", 100)
        Call .Add(, , "Datum splátky", 100)
        Call .Add(, , "Výše splátky", 90)
        Call .Add(, , "Typ pøedpisu", 130)
        Call .Add(, , "Výše pøedpisu", 90)
        Call .Add(, , "Výše platby", 90)
    End With
    
    DataCheck.ListView3.ListItems.Clear
    
    Set rst1 = New ADODB.Recordset
    'asss = InputBox("", "", sql)
    rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
        If Not rst1.EOF Then
         
         
            rst1.MoveFirst
            For i = 1 To rst1.RecordCount
                Set itmX = DataCheck.ListView3.ListItems.Add()
                itmX.text = "" & rst1(2) & ""
                itmX.SubItems(1) = "" & rst1(3) & ""
                itmX.SubItems(2) = "" & rst1(4) & ""
                itmX.SubItems(3) = "" & rst1(5) & ""
                itmX.SubItems(4) = "" & rst1(6) & ""
                itmX.SubItems(5) = "" & rst1(7) & ""
                rst1.MoveNext
            Next i
        End If
    
    rst1.Close
    
    Call Listview_Proper_Sort(DataCheck.ListView3, DataCheck.ListView3.ColumnHeaders(1), blnAscendingOrder:=True)   ' "Datum platby"
   
' **********************************************************************  - Pridanie_klienta
Case "Pridanie_klienta"
   
   Set rst1 = New ADODB.Recordset
   'asss = InputBox("", "", sql)
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
      
      rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Pridanie_klienta.ListView1.ListItems.Add()
      itmX.text = "" & rst1(0) & ""
      itmX.SubItems(1) = "" & rst1(1) & ""
      itmX.SubItems(2) = "" & rst1(2) & ""
      itmX.SubItems(3) = "" & rst1(3) & ""
      itmX.SubItems(4) = "" & rst1(4) & ""
    rst1.MoveNext
    Next i
      
   End If
   
   rst1.Close
  
  
  
Case "Klient_exeOther"
' ********************************************************************** Formuláø Klient_exe - Klient_exe BASE

     Klient_exe.Label1.Caption = ""
     Klient_exe.Label2.Caption = ""
     Klient_exe.Label3.Caption = ""
     Klient_exe.Label4.Caption = ""
     Klient_exe.Label5.Caption = ""
     Klient_exe.Label6.Caption = ""
     Klient_exe.Label91.Caption = ""
     Klient_exe.Label95.Caption = "" 'RISK-403
     Klient_exe.Label96.Caption = "" 'RISK-403
     
     Klient_exe.Label87.Caption = ""
      
      
      sql = "select * from " & schema & ".v_person_base where id_person = " & id_person
       
       Set rst1 = New ADODB.Recordset
       rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
       
       If Not rst1.EOF Then
        rst1.MoveFirst
        For i = 1 To rst1.RecordCount
         Klient_exe.Label1.Caption = "" & rst1(0) & ""
         Klient_exe.Label2.Caption = "" & rst1(1) & ""
         Klient_exe.Label3.Caption = "" & rst1(2) & ""
         Klient_exe.Label4.Caption = "" & rst1(3) & ""
         Klient_exe.Label5.Caption = Trim(Replace("" & c.Address_P & "", "XNA", ""))
         Klient_exe.Label6.Caption = Trim(Replace("" & c.Address_C & "", "XNA", ""))
         Klient_exe.Label95.Caption = c.PrimaryPhone 'RISK-403
         Klient_exe.Label96.Caption = c.PrimaryEmail 'RISK-403
         Klient_exe.Label91.Caption = "" & rst1(12) & ""
         
            If "" & rst1(6) & "" = "Y" Then
                Klient_exe.Label82.Visible = False
                Klient_exe.Label83.Visible = True
            Else
                Klient_exe.Label82.Visible = True
                Klient_exe.Label83.Visible = False
            End If
            
            If "" & rst1(7) & "" = "Y" Then
                Klient_exe.Label85.Visible = False
                Klient_exe.Label84.Visible = True
            Else
                Klient_exe.Label85.Visible = True
                Klient_exe.Label84.Visible = False
            End If
            
            If "" & rst1(10) & "" = "Y" Then
                Klient_exe.Label86.Visible = True
                Klient_exe.Label87.Visible = True
                Klient_exe.Label87.Caption = "" & rst1(11) & ""
            Else
                Klient_exe.Label87.Visible = False
                Klient_exe.Label86.Visible = False
            End If
         
        rst1.MoveNext
        Next i
       End If
       
       rst1.Close
    ' ********************************************************************** Formuláø Klient_exe - èísla smluv
     Klient_exe.ListView2.ListItems.Clear
     
     'sql = "select id_person, code_loan_number from " & schema & ".loan where id_person = " & id_person & " order by code_loan_number" ' FA: 19.3.2018 nahrazena tabulka DEAL za LOAN
    sql = "SELECT a.ID_PERSON, a.CISLO " & _
          "FROM " & schema & ".v_datacheck_contracts a " & _
          "WHERE id_person = " & id_person & " AND a.TYP_SMLOUVY IN ('HU', 'HR', 'KTK') " & _
          "ORDER BY a.TYP_SMLOUVY"

       Set rst1 = New ADODB.Recordset
       rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
       
       If Not rst1.EOF Then
        rst1.MoveFirst
        For i = 1 To rst1.RecordCount
          Set itmX = Klient_exe.ListView2.ListItems.Add()
          itmX.text = "" & rst1![cislo] & ""
        rst1.MoveNext
        Next i
       End If
       
       rst1.Close
    
    
      ' ********************************************************************** Formuláø klient_exe - upozorneni, pokud se lisi datum zesplatneni na dwh a OBS - kod zkopirovan z karty klienta
     sql = "select id_person from " & schema & ".wrong_dp where id_person = " & id_person
     Set rst1 = New ADODB.Recordset
     rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
     
     If Not rst1.EOF Then
                Klient_exe.Label97.Visible = True
                Klient_exe.Label92.Visible = True
     Else
                Klient_exe.Label97.Visible = False
                Klient_exe.Label92.Visible = False
               
     End If
     rst1.Close
     
    
    ' ********************************************************************** Formuláø Klient_exe - další informace
     
     Klient_exe.ListView5.ListItems.Clear
     sql = "select * from " & schema & ".v_information_exe where id_person = " & id_person
     
     'RISK-442
     
       Set rst1 = New ADODB.Recordset
       rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
      
    
       ' rst1.MoveFirst , RISK - 442
        
          Set itmX = Klient_exe.ListView5.ListItems.Add()
          itmX.text = "" & rst1.Fields(1).name & ""
          
          Set itmX = Klient_exe.ListView5.ListItems.Add()
          'itmX.text = "Suma splátek akt. úvìrù pøi zesplatnìní"
          itmX.text = "Souèet pùvodních splátek"
        
        For i = 2 To rst1.Fields.count - 1
          Set itmX = Klient_exe.ListView5.ListItems.Add()
          itmX.text = "" & rst1.Fields(i).name & ""
        Next i
        
        If Not rst1.EOF Then
       
          Set itmX = Klient_exe.ListView5.ListItems(1)
          itmX.SubItems(1) = "" & rst1(1) & ""
          
         If c.GetAmt_InstalmentBeforePayoff = 0 Then
          helpsum = ""
          Else: helpsum = "" & c.GetAmt_InstalmentBeforePayoff & " Kè"
          End If
          
          Set itmX = Klient_exe.ListView5.ListItems(2)
          itmX.SubItems(1) = helpsum
       
         For i = 2 To rst1.Fields.count - 1
          Set itmX = Klient_exe.ListView5.ListItems(i + 1)
          itmX.SubItems(1) = "" & rst1(i) & ""
         Next i
        End If
        
       
       rst1.Close
    
       
     ' ********************************************************************** Formuláø Klient_exe - SOP
     
     Klient_exe.ListView4.ListItems.Clear
     
    Call c.SOP.Vytvor_listview_prehled_SOPu(Klient_exe.ListView4)
    Klient_exe.CommandButton20.Enabled = c.SOP.blnPovolitAutorizaci
     
     
     sql = "select * from " & schema & ".v_sop_balance where id_person = " & id_person
     Set rst1 = New ADODB.Recordset
     rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
        If Not rst1.EOF Then
            Klient_exe.Label77.Caption = rst1(4) & " Kè"
            Klient_exe.Label78.Caption = rst1(2) & " Kè"
        End If
         
     rst1.Close
    
      ' ********************************************************************** Formuláø Klient_exe - upozornìní
     
     Klient_exe.ListView3.ListItems.Clear
     
     'Code status:
     
     'a - active
     'd - done
     's - storno/deleted
     
     sql = "select skp_notice, text_notice, dtime_notice_deadline, dtime_notice_deadline - trunc(sysdate) as days_left, assignee from " & schema & ".notice where code_status = 'a' and id_person = " & id_person & " order by days_left"
     Set rst1 = New ADODB.Recordset
     rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
     
     If Not rst1.EOF Then
      rst1.MoveFirst
      For i = 0 To rst1.RecordCount - 1
          Set itmX = Klient_exe.ListView3.ListItems.Add()
          itmX.text = "" & rst1(0) & ""
          itmX.SubItems(1) = "" & rst1(1) & ""
          itmX.SubItems(2) = "" & rst1(2) & ""
          itmX.SubItems(3) = "" & rst1(3) & ""
          itmX.SubItems(4) = "" & rst1![Assignee] & ""
      rst1.MoveNext
      Next i
     End If
     rst1.Close
       
   
'************************************************************************************************
  
Case "Exekuce"
   
   Exekuce.ListView1.ListItems.Clear
   
   Set rst1 = New ADODB.Recordset
   'asss = InputBox("", "", sql)
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
   
   If Not rst1.EOF Then
    rst1.Sort = "ID_PERSON, DTIME_VALID_FROM DESC"
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Set itmX = Exekuce.ListView1.ListItems.Add()
      itmX.text = "" & rst1(0) & ""
      itmX.SubItems(1) = "" & rst1(1) & ""
      itmX.SubItems(2) = "" & rst1(2) & ""
      itmX.SubItems(3) = "" & rst1(3) & ""
      itmX.SubItems(4) = "" & rst1(4) & ""
      itmX.SubItems(5) = "" & rst1(5) & ""
      itmX.SubItems(6) = "" & rst1(9) & ""
      itmX.SubItems(7) = "" & rst1(6) & ""
      itmX.SubItems(8) = "" & rst1(7) & ""
      itmX.SubItems(9) = "" & rst1(8) & ""
      itmX.SubItems(10) = "" & rst1(10) & ""
    rst1.MoveNext
    
    If blnPrank Then
        If i < 50 And i > 15 And Rnd() < 0.015 Then
            Set itmX = Exekuce.ListView1.ListItems.Add()
            itmX.text = 1052
            itmX.SubItems(1) = 2063639
            itmX.SubItems(2) = 411265
            itmX.SubItems(3) = "Brožina Vít"
            itmX.SubItems(4) = "8509146969"
            itmX.SubItems(5) = "Øízení zaplaveno" '"Vyjádøení vìøitele"
            itmX.SubItems(6) = "Exekuèní zaøízení"
            itmX.SubItems(7) = "137 Ex 674/17"
            itmX.SubItems(8) = "9.6.2017 13:36:19"
            itmX.SubItems(9) = "1.1.3000"
            itmX.SubItems(10) = "N"
        End If
    End If
      
    
    Next i
   End If
   
   rst1.Close
  
  
  
  ' ********************************************************************** Formuláø žaloby: èíselník
Case "ExekuceEnumeration"
      
  
   
Call query1("", "ExekuceVhodnyCiselnik")



  ' ********************************************************************** Formuláø žaloby: èíselník
   
Case "ExekuceVhodnyCiselnik"

'If Zaloby.ComboBox2.ListCount > 0 Then

sql = "select name_exe_status from " & schema & ".exe_status where dtime_valid_to > sysdate and skp_process = 2 order by order_item nulls last"

  Exekuce.ComboBox1.Clear

   Exekuce.ComboBox1.AddItem "Vše"

   Set rst1 = New ADODB.Recordset
   rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

   If Not rst1.EOF Then
    rst1.MoveFirst
    For i = 1 To rst1.RecordCount
      Exekuce.ComboBox1.AddItem "" & rst1(0) & ""
    rst1.MoveNext
    Next i
   End If

    rst1.Close

   Exekuce.ComboBox1.ListIndex = 0



' **********************************************************************

Case "Klient_exeHistory"
   
    
    Klient_exe.ListView1.ListItems.Clear
    
    
    sql = "select * from " & schema & ".v_history_exe where id_person = " & id_person & " and (1=0 "
   
    If Klient.CheckBox1.value Then sql = sql & " or type = 'Stav' "
    If Klient.CheckBox3.value Then sql = sql & " or type = 'Klient' "
    If Klient.CheckBox2.value Then sql = sql & " or type = 'Poznámka' "
    If Klient.CheckBox4.value Then sql = sql & " or type = 'Upozornìní' "
    If Klient.CheckBox5.value Then sql = sql & " or type = 'SOP' "
    
    sql = sql & ") order by dtime_note_created DESC, skp_epr_note DESC"
    
    Set rst1 = New ADODB.Recordset
    rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
    
        If Not rst1.EOF Then
             rst1.MoveFirst
             For i = 1 To rst1.RecordCount
                 Set itmX = Klient_exe.ListView1.ListItems.Add()
                 itmX.text = "" & rst1(1) & ""
                 itmX.SubItems(1) = "" & rst1(2) & ""
                 itmX.SubItems(2) = "" & rst1(3) & ""
                 itmX.SubItems(3) = "" & rst1(4) & ""
                 itmX.SubItems(4) = "" & rst1(5) & ""
                 itmX.SubItems(5) = "" & rst1(6) & ""
                 rst1.MoveNext
             Next i
        End If
        
    rst1.Close
    
    Call ZabarviRadkyTextove(Klient_exe.ListView1, 2, "Stav", barvaModra)
    Call ZabarviRadkyTextove(Klient_exe.ListView1, 2, "Klient_exe", barvaZelenaSvetla)
    Call ZabarviRadkyTextove(Klient_exe.ListView1, 2, "Poznámka", barvaSeda)
    Call ZabarviRadkyTextove(Klient_exe.ListView1, 2, "Upozornìní", barvaRuda)
    Call ZabarviRadkyTextove(Klient_exe.ListView1, 2, "SOP", barvaFialova)
    
' **********************************************************************




End Select

Exit Sub

connection_problem:

If Mid(Err.Description, 1, 9) = "ORA-03114" Then
err_connection = True
Else
 MsgBox Err.Description
End If

End Sub


