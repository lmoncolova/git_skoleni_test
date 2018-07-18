Attribute VB_Name = "Komentare"
' Prostor pro uzitecne komentare





''''''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''''''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''''''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''''''///////    Prace s VBA

' autor: Filip Antol
' posledni editace: 5.3.2018


' DEKLARACE PROMENNYCH!
' na zacatku kazdeho modulu mit napsano OPTION EXPLICIT, co znamena ze kazda promenna v tom modulu se pak musi nadeklarovat.
' velmi dulezite pro eliminaci zbytocnych bugu, taky to obcas usetri hodiny hledani jednoho preklepu.
' google it...
' lze zapnout automaticky pro nove moduly (ne pro stavajici): klikni v horni liste na Tools -> Options -> Require Variable Declaration


' syntax error hlaska
' lze vypnout otravne vyskakovaci okno (editor stale podbarvi error cervene): klikni v horni liste na Tools -> Options -> Require Variable Declaration


' compilace projektu
' horni lista: Debug -> Compile Jezevcik
' vba zkompiluje projekt a upozorni na nektere typy chyb (treba kdyz se vola neexistujici fce, nebo neni nadeklarovana promenna)
' doporucuji vzdy pred ulozenim
' sikovne pri programovani, ulehci hodne prace


' Immediate window
' zapne se pres View -> Immediate window
' pokud si v debug modu, lze zjistovat hodnoty promennych (e.g. debug.print(c.id_person), nebo ?c.id_person)
' lze zavolat Subs nebo Functions, e.g. "jj" (spusti jezevcika) nebo ?ucase("spam")  (ucase je funkce jako kazda jina)


' Locals window
' zapne se pres View -> Locals window
' ukazuje hodnoty vsech nadeklarovanych promennych


' Bookmarks
' lze pridat na horni listu, velmi to zjednodusi opakovane preskakovani mezi ruznymi misty v kodu







''''''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''''''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''''''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''''''///////    Programovani obecne - dobre programatorske praktiky

' autor: Filip Antol
' posledni editace: 5.3.2018


' upozorneni: Jezevcik prede mnou vyvijelo nekolik lidi a vetsina kodu se tymito pravidly neridi.
' zaroven pisu jenom tolik co ja sam umim a co se mi vyplatilo...
' Nejsem zadnej expert, internet je bohatej, muzes si overit jestli tyhle myslenky davaji smysl.



''' odsazovani kodu
' VBA editor je VELMI archaicka vec, a neumi automaticke odsazovani. Taky neumi ukazat END IF, pokud oznacis IF a pod.
' proto je hodne uzitocne kod radne odsazovat. Psani neodsazeneho nebo nespravne odsazeneho kodu je podle mne svinarna, nelze to cist a melo by to byt ilegalni. (autor: Filip Antol)
' doplnim, ze cim hezceji a intuitivneji napsan kod, tim vic se programator muze soustredit na cteni te myslenky v nem (ne formy)
' ja pouzivam tabulator, ktery odsazuje 4 mezery. Je to hodne, ale pracuje se mi s tim vyborne a kod mi prijde prehledny.


''' pokud to lze, nepsat SQL dotazy ve VBA. VZDY se dotazovat jenom do view (velka cast Jez. tohle jeste nedela) ze 2 duvodu:
'    1. dotazy napsane ve VBA se tezko udrzuji a debuguji
'    2. aby mel databazovy objekt referenci - u tabulky hned vidim, ze je na ni napojene nejake view (nemusim otevirat excel a hledat)


''' psani komentaru
' kod bez komentaru je zhruba rovnako spatny jako kod s prilisnym mnozstvim komentaru, jedine co je jeste horsi je kod s nepravdivymi komentari...
' kod samotny rika JAK se veci delaji, do komentaru je lepsi napsat PROC, vyssi logiku cele veci, kterou programator nevycte z par radku..
' plati DRY princip (popsan nize) -> nepis do komentaru co je zjevne z kodu samotneho
' podpis se i s datem (vcetne roku)



''' zopar programatorskych pravidel, kterymi ma smysl se ridit:
' pokud ma funkce nebo metoda(sub) vice nez 200 radku, tak je neco spatne...
' the DRY principle = DON'T REPEAT YOURSELF!
'     > kazda myslenka by mela mit jedinou reprezentaci v danem systemu
'     > moje babicka tvrdila, ze v dobrem kodu se radny radek nevyskytne dvakrat...
'     > dedecek zasa rikaval, ze nejrychlejsi zpusob jak spoznat amatersky napsanej kod je mnozstvi duplicit v nem.
' make your code shy = pis hanblivy kod. Mysli se tim, ze promenne deklaruji na nejnizsi uroven na jakou lze (pouze pro funkci, pokud musim tak pro modul, a jenom pokud musim tak pro celej projekt...)
'     > plati pro vsechny metody(sub a funkce), promenne, ...
' Kod podleha entropii = s casem se rozpada pokud se neudrzuje.
'     > pokud vidis spatnej kod, zastav se pri nem a oprav ho, stoji to za to.
' VHODNE NAZVY - vzdy stoji za to se zamyslet nad vhodnym nazvem uplne hned pri prvnim vytvareni promenne/funkce/...
' kod bude napsanej jednou, ale ctenej mnohokrat -> proto ho pis citelne, nebud linej a misto cc pojmenuj promennu CilovaCastka
' make your code easy to reuse
' ak kod produkuje ten vystup ktory ma, tak to ... NEstaci! Dobry kod musi splnat este mnoho dalsieho.



''' Proc/kdy vytvarime funkce? zaludna to otazka, ma vicero moznosti
'   > pokud nejakou akci delam opakovane (treba pokud se to bude menit, aby byla zmena na jednem miste)
'   > v ramci DRY principu  (....zmena na jedinem miste)
'   > aby sem si ulehcil debugging v budoucnosti - overit funkcnost kratsi funkce je milionkrat jednodussi nez prechazet 200 radkovy modul a hledat to v nem..
'   > ... a nebo ... (dramaticka pauza pred nejlepsi odpovedi)
'   > abych pojmenoval kus kodu.













''''''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''''''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''''''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''''''///////    Popis zopar konkretnich problemu ktere se vyskytli


'Historizace èástek v Jezevèíkovi

'Ahoj,
'
'posílám struèné info k tomu, jak se zaznamenávají zmìny týkající se èástek v Jezevèíkovi.
'
'Uživatelé Jezevèíka mohou editovat èástky týkající se nákladù øízení (soudní poplatky, pøeplatek/doplatek, cestovné, úhradu poplatkù). Veškeré zmìny, vèetnì autorizace tìchto údajù vedoucím oddìlení PV, jsou zaznamenány na kartì klienta.
'
'Dalším editovatelným atributem v Jezevèíkovi je výše pohledávky v info boxu - klienta. Tento údaj se nyní již nevyužívá a sloužil k zaznamenání výši pohledávky k datu podání žaloby. I tak se zmìny propíšou na kartu klienta.
'
'Všechny další údaje týkající se èástek jsou brány z dwh, pøepoèítavají se každý den a v Jezevèíkovi je není možné editovat.
'
'Aby byl výèet kompletní, musím dodat, že veškeré údaje, které se v Jezevèíkovi vyskytují, lze velmi jednoduše pøepsat/smazat SQL dotazem od analytika s pøístupem na AP_COLL, aniž by to bylo kdekoliv v Jezevèíkovi zaznamenáno.
'
'Petr






'Popis: Pomoc s instalací doplòku do excelu
'
'ERROR: could not load some objects...

'Dobrý den,
'mùžete nám prosím pomoct s instalací doplòku do excelu? Nefunguje nám soubor s makry a aktivními prvky (konkrétnì jde zøejmì o Microsoft Date and Time Picker Control), pro které je potøeba registrovat soubor MSCOMCT2.OCX, postup je popsán zde
'http://www.mrexcel.com/forum/excel-questions/788407-excel-2013-64-bit-missing-microsoft-date-time-picker-control.html
'Kolega na pozici pøede mnou to tu jednou øešil, nevím teda s kým.
'Díky za pomoc



'Øešení , poznámky:
'12.11.2015: Problém zpùsobily nekorektnì nainstalované klony PC a NB s Windows 7 Pro x64 a chybìjící systémové soubory a knihovny.
'Pokud nastane tento problém, je tøeba pøedevším doplnit soubor MSCOMCTL.OCX do cesty C:\Windows\System32, soubor MSCOMCT2.OCX do cesty C:\Windows\System32 i do C:\Windows\SysWOW64 a do této složky ho korektnì zaregistrovat.
'
'




'Nìjaké základní info:
'
'Jezevèík je sada  propojených UserFormù (formuláøù) a maker naprogramovaných v VBA. Aplikace není samostatnì spustitelná,  bìží s excelem na pozadí.
'
'Pro práci v aplikace je nutné pøihlásit se do databáze (pomocí ADODB.Connection). Uživatelé pomocí dotazù komunikují  s databází. Celá skupina tabulek,
'která tvoøí datovou základnu pro tuto aplikaci je uložena na datovém høišti AP_COLL.  Dùležité tabulky jsou propojeny primárními a cizími klíèi, oindexovány a s
'louží jako podkladové tabulky pro reportování na PV. Dvakrát týdnì se provádí záloha všech tabulek (tabulky se vždy pøepíšou). Zmìny záznamù zpùsobené uživateli jsou èásteènì historizovány.
'
'Data jsou na denní bázi aktualizovány z replik na dwh pomocí nìkolika procedur - automaticky se vloží napø. zesplatnìný klient, ale kontroluje se i zmìna údajù u
'klientù (napø. insolvence, datum úmrtí atd..). Údaje od uživatelù aplikace se propíšou do zdrojových tabulek okamžitì. Nedokáže komunikovat s ostatními systémy v bance (napø. OBS).







'Chyba Error Object Library Feature Not Supportted

'RESENI
'I followed your advice, and dereferenced OLE Automation (I am using Microsoft Office 14.0 Object Library so it was left on) and compiling passed. Than, I re-referenced OLE Automation, and still works.
'
'Just to be sure, I was thinking maybe order of references are the in question, so I put OLE Automation up in the list (as it was when the error occurred) - but error did not reappear.
'
'So, my conclusion is: just de-reference, then re-reference OLE Automation - it worked for me.

