Attribute VB_Name = "mod_Testovani"
Option Explicit

' Vytvoril FA 27.3.2018 ve verzi Jez.5.4
' modul pro exekuci testovacích skriptù
' Cilem je, aby (napr. pred nasazenim nove verze) bylo mozne rychle a jednoduse vyzkouset vetsinu beznych funkcionalit Jezevcika
' Vystupem by byl nejaky log (textak) jestli je vse OK - zaznam ke kazdemu kroku.
' Melo by zachytit alespon nektere elementarni chyby:
'   > chybejici grant na tabuli
'   > obecne chyba v SQL query
'   > chyby pri nacitani userformu
'   > chyby pri stlaceni buttonu
'
'
' Taky to umozni jednoduche testovani pred releasem pro ostatni uzivatele - staci ze se prihlasi do Jezevcika a spusti test (nejaky button..)
'     (dulezite kvuli grantum, pri testovani bude do DWH prihlaseny jiny uzivatel)
'
' Bude idealni vytvorit vhodneho testovaciho (ne)klienta na produkci.
' Jednotlive testovaci funkce by asi idealne meli byt soucasti modulu, ktery testuji.. T.j. Test vytvoreni objektu Client by mel byt nadefinovan v modulu Client, tady se pouze zavola.
' Rozhodne supr by bylo mit vsechny dotazy do SQL pohromade, a tady by se otestovala jejich validita. To by ale museli byt nejak zgrupene, bud jako public dictionary, nebo alespon na urovni jednotlivych modulu.
'



'''  generovani moznosti
' > vyextrahovat vsechny SQL dotazy na jedno misto
' > klikatko priamo pred ocami
' > Test_me
' > KlikBot = bot, ktery to naklika. zatial co test_me() budu testovat funkcionalitu, KlikBot by testoval funkcnost preklikavacich buttons na kartach a navraty.
' > najmout vietnamce (e.g. Vasek)
'
'
'
'
'
'






' NEXT: FA: skoncil som pri tom, ze nechcem kvoli testovaniu nastavit vsetky mozne buttons na public, to by bola kravina.
'           v Exekuce je Sub Test_me, ktora by bola public ale ostatne by mohli ostat private
'           > tak jak to je teraz ale Test_me nedava uplne zmysel (iba som vytvoril alias..), treba vygooglit ako to korektne spravit..
'           > kazdopadne myslienka sub Test_me v kazdom module mam pocit ze je spravny smer.


' TEST_ME
'
'  > ako vyriesit to, ze jeden form / objekt pri svojom testovani asi bude volat ine objekty / testy? (e.g. Client.test_me bude volat Loan.test_me?)
'  > ak napr. karta Klient chce skusit vlozenie stavu, zavola sa update_stav_all.test_me(nazov_karty / rizeni, klient)
'  > podstata kazdej test_me metody je, ze nielen spravi x, ale zaroven pritom otestuje dalsich par veci (nielen vlozi stav, ale zaroven skusi vlozit 2 ine stavy s poznamkou a hned ich aj zmaze...)


' VYSTUP:
'
'  > bolo by super zaistit jednoduchy LOG subor
'  > chceme jednoduchu funkciu, ktorej vstup bude string a ten sa vlozi na novy riadok a ulozi
'  > asi moze byt otvoreny ten textacik, lebo preco nie..


Public Sub Jez_full_test(Optional ByVal CUID As Long = -11)

    Call jj
    
    
    
    If schema = "APP_COLL_JEZEVCIK_TEST" Then
                
        Exekuce.test_me (CUID)
        Klient_exe.test_me id_person_in:=CUID ' ak error, search "xtstereehcvgb"
        Klient.test_me id_person_in:=CUID
        Debug.Print c.name
        
        
    Else
        ' SEM poistku ze na PROD se da pouze pro pana Testa
        MsgBox "Funguje zatím jenom na testu"
    End If
    

End Sub
