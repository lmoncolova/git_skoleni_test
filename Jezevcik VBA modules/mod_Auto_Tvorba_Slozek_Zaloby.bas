Attribute VB_Name = "mod_Auto_Tvorba_Slozek_Zaloby"
'RISK-679 automatická tvorba složek pro klienty vhodné k žalobì
Option Explicit
Public Const strVedouciPV As String = "AIR_RAJSIGLOVA" '"AIR_ANTOL" '"AIR_RAJSIGLOVA"

Private auto_generate_folder_update_needed As Boolean, auto_generate_final_file As String

Sub generate_client_folders()
Dim destPath As String, sourcePath As String, name As String, CUID As String


If UCase(Main.username) = strVedouciPV Then 'vždy pøi prvním denním spuštìní Evou Rajsiglovou (vedoucim pravniho vymahani)

    Call check_todays_update 'Funkce, která se podívá na již zaevidované klienty, a zjistí, jestli je tøeba dnes znova generovat složky
    If auto_generate_folder_update_needed Then
    
    
        Dim sql As String, rst1 As Recordset, i As Integer
        
        'Sql k vytažení vhodných klientù k založení složky a základních informací o nich
        sql = "SELECT a.ID_PERSON, a.CUID, a.NAME, a.DATE_PAYOFF, a.DATE_SIGNATURE_LOAN, a.DATE_CONTRACT_SIGNATURE, a.DATE_BANK_ACCOUNT_ACTIVATION " & _
              " FROM " & schema & ".v_auto_client_folders a"
    
         Set rst1 = New ADODB.Recordset
         rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic
        
         If Not rst1.EOF Then
          rst1.MoveFirst
          
          'Kontrola uživatelem, jestli chce vytvoøit nové složky.
          
          If MsgBox("Dobré ráno " & JezUserLoggedIn.StrNameVocative & ", pøeješ si vytvoøit složky " & rst1.RecordCount & " novým klientùm vhodným k žalobì?", vbYesNo, "Vytvoøení nových složek") = vbYes Then
          
              For i = 0 To rst1.RecordCount - 1
                
                
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'Samotný kód pro tvorbu složek zde:
                            
                    If schema = "APP_COLL_JEZEVCIK" Then
                        destPath = cfg.Cesta("EprPodaneZaloby")
                    Else: destPath = cfg.Cesta("TestovaciSlozka")
                    End If
                    
                    sourcePath = cfg.Cesta("OP_PPU_CENIK")
                    
                    name = rst1![name]
                    CUID = rst1![CUID]
                    
                    
                    'Pokud jsem pro klienta vytvoøil složku, tak mùžu pokraèovat dál
                    ' Pokud jsem na testu, tak funkce sama bude ukládat do bordelu, ne do produkèních složek
                    destPath = destPath & name & " " & CUID & "\Pøílohy\"
                    If CreatePath(destPath) Then 'Funkce vrátí True, pokud složka ještì neexistovala a byla novì vytvoøena, jinak vrátí false
                        
                        
                        
                        
                        '' 5 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'budou se ukládat Obchodní podmínky ke dni uzavøení rámcové smlouvy a k datu zesplatnìní (jestliže klient má více HU a ty jsou zesplatnìna v rùzná data, bereme první datum zesplatnìní, tzn. to nejstarší)
                        
                        If rst1![DATE_CONTRACT_SIGNATURE] <> vbNullString Then
                        
                        
                        Call LoopThroughFiles(sourcePath, "OP", rst1![DATE_CONTRACT_SIGNATURE])
                            'A pošlu jej klientovi do složky
                            Call CopyFile(auto_generate_final_file, sourcePath, destPath)
                        
                            'Pokud má navíc zesplatnìné HU, tak dotáhnu i podmínky k datu zesplatnìní HU
                            If rst1![Date_payoff] <> vbNullString Then
                                'Zjistím si, který soubor podle data mám dotáhnout
                                Call LoopThroughFiles(sourcePath, "OP", rst1![Date_payoff])
                                'A pošlu jej klientovi do složky
                                Call CopyFile(auto_generate_final_file, sourcePath, destPath)
                               
                            End If
                        
                        
                        End If
                        ''   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
                        
                        
                        
                        
                        '' 6 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'budou se ukládat Podmínky pro používání úvìru k datu uzavøení 1. úvìru klienta a k datu zesplatnìní (jestliže klient má více HU a ty jsou zesplatnìna v rùzná data, bereme první datum zesplatnìní, tzn. to nejstarší)
                        
                        If rst1![DATE_SIGNATURE_LOAN] <> vbNullString Then
                        
                        
                        Call LoopThroughFiles(sourcePath, "PP", rst1![DATE_SIGNATURE_LOAN])
                            'A pošlu jej klientovi do složky
                            Call CopyFile(auto_generate_final_file, sourcePath, destPath)
                        
                            'Pokud má navíc zesplatnìné HU, tak dotáhnu i podmínky k datu zesplatnìní HU
                            If rst1![Date_payoff] <> vbNullString Then
                                'Zjistím si, který soubor podle data mám dotáhnout
                                Call LoopThroughFiles(sourcePath, "PP", rst1![Date_payoff])
                                'A pošlu jej klientovi do složky
                                Call CopyFile(auto_generate_final_file, sourcePath, destPath)
                               
                            End If
                        
                        
                        End If
                        ''   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
                        '' 7,8 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'jestliže má klient také záporný zùstatek na úètu,
                        If rst1![DATE_BANK_ACCOUNT_ACTIVATION] <> vbNullString Then 'Date bank account activation se ve skriptu nahoøe plní jen v pøípadì, že je klient v debetu
                            
    
                        
                            'automaticky se do složky nahraje ceník k datu aktivace úètu a k datu zesplatnìní HU
                            'Zjistím si, který soubor podle data mám dotáhnout
                            Call LoopThroughFiles(sourcePath, "Ceník", rst1![DATE_BANK_ACCOUNT_ACTIVATION])
                            'A pošlu jej klientovi do složky
                            Call CopyFile(auto_generate_final_file, sourcePath, destPath)
                        
                            'Pokud má navíc zesplatnìné HU, tak dotáhnu i ceník k datu zesplatnìní HU
                            If rst1![Date_payoff] <> vbNullString Then
                                'Zjistím si, který soubor podle data mám dotáhnout
                                Call LoopThroughFiles(sourcePath, "Ceník", rst1![Date_payoff])
                                'A pošlu jej klientovi do složky
                                Call CopyFile(auto_generate_final_file, sourcePath, destPath)
                               
                            End If
                        
                            'automaticky se do složky nahraje Pøehled nákladù úèelnì vynaložených bankou na vymáhání dluhu.pdf
                            Call CopyFile("Pøehled nákladù úèelnì vynaložených bankou na vymáhání dluhu.pdf", sourcePath, destPath)
                                            
                        
                        End If
                        ''   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
                        
                        '' 9 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'do všech složek se automaticky nahrají dvì povìøení - Povìøení Štefl.pdf a Povìøení Seidlová.pdf
                        Call CopyFile("Povìøení Štefl.pdf", sourcePath, destPath)
                        Call CopyFile("Povìøení Seidlová.pdf", sourcePath, destPath)
                        ''   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
                        
                    'Na úplný závìr uložím insertem do databáze, co jsem všechno provedl
                    sql = "insert into  " & schema & ".automatic_client_folders values (" & rst1![id_person] & ",'" & rst1![CUID] & "','" & rst1![name] & "',sysdate," & "'" & UCase(Main.username) & "')"
                    Call execute(sql)
                        
                    End If
                    
                    
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                
    
                
                
                
                rst1.MoveNext
              Next i
              
              
              MsgBox ("Složky byly vytvoøeny")
              
          Else: rst1.Close
                Exit Sub
                
              
          End If
         End If
         rst1.Close
    
    
    'Ještì uložím to, že jsem provedl update - i když by tam žádný klient nebyl
     sql = "insert into  " & schema & ".automatic_client_folders_upd values (sysdate," & "'" & UCase(Main.username) & "')"
     Call execute(sql)
    
    End If

End If


End Sub



Sub check_todays_update()

Dim sql As String, rst1 As Recordset

 sql = "select nvl(trunc(max(dtime_update)),date'1000-01-01') as max_update_date from " & schema & ".automatic_client_folders_upd"
 auto_generate_folder_update_needed = False
 
 Set rst1 = New ADODB.Recordset
 rst1.Open sql, cs_HDWMN, adOpenDynamic, adLockBatchOptimistic

 If Not rst1.EOF Then
    rst1.MoveFirst
    If rst1![max_update_date] <> date Then
    auto_generate_folder_update_needed = True
    End If
         
    rst1.MoveNext
  
 End If
 rst1.Close

End Sub



Sub LoopThroughFiles(srcPath As String, File_Type As String, Max_Date As Date)
    Dim MyObj As Object, MySource As Object, file As Variant, local_max_date As Date
    
    file = Dir(srcPath)
    
    While (file <> "") ' Hledám dokument s platností pro daného klienta, napø. Platil v dobì, kdy klient podepsal úvìr. - Dìlám loop pøes všechny co tam jsou a vezmu ten s maximálním datem, menším než datum podepsání úvìru
        If InStr(file, File_Type) > 0 And Mid(file, InStrRev(file, " ") + 1, Len(file) - InStrRev(file, " ") - 4) <= Max_Date And Mid(file, InStrRev(file, " ") + 1, Len(file) - InStrRev(file, " ") - 4) > local_max_date Then
            local_max_date = Mid(file, InStrRev(file, " ") + 1, Len(file) - InStrRev(file, " ") - 4)
            auto_generate_final_file = file
        End If
    file = Dir
  Wend

End Sub







