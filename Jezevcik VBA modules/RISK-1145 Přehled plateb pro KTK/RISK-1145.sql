-- RISK-1145        


-- Martin 11.7.2018: tabulku overdraft_posledni_zustatek zrus, detto v_overdraft_datum_zanoreni - pridaj ako atributy do hlavnej tabulky overdraft

/*PLAN

NAPOCET TESTOVACICH DAT
    > klienti - procka load_overdraft_clients - pridat do pckg okolo riadku 100 (vysvetlene v tele procky)
    > potom zavolat PCKG
    > potom naplnit tabule zo skriptu "plneni tabulek do procky.sql"



> [DONE] create table overdraft
> [DONE] create table overdraft_balance_pd
> [DONE] pridat constraints a komentare k tabulkam novym
> [DONE] upravit v_datacheck_contracts, v_nezesplatneni, v_zesplatneni, a pod. (doplnit k joinu na LOAN join na OVERDRAFT)
> zistit ktore dalsie VIEWs potrebuju update - otazka na uzivatelov, kde vsade to chcu vidiet - aasiii iba v ramci datacheck?
> dblClick na KTK na karte Kontrola Dat -> doplnit veci podobne ako u HUciek

> [DONE] V_OVERDRAFT_PRUBEH
     > hlavne view pre jistinu
     > i pro uroky (pripadne se udela view or table: predpisy a uhrady uroku)


PCKG:
> [DONE - skript "skript "plneni tabulek do procky.sql"]   plnenie overdraft
> [DONE - skript "skript "plneni tabulek do procky.sql"]   plnenie overdraft_balance_pd
> [DONE - skript "skript "plneni tabulek do procky.sql"]   plnenie OVERDRAFT_POSLEDNI_ZUSTATEK
> do tabulky UCET chceme dostat vnitrni ucet ku KTK. Dnes sa nenaplni, sice ma SKP_PERSON osoby, ale ma ID_BANK_PRODUCT = 0 a v plneni je podmienka na 1.
   > [DONE - skript "PCKG - pridani KTK uctu do tabule UCET.sql"] pridat skp_bank_account_type 21355 do skupiny exekucnich, insolvencnich, jpk uctu. (je tam RISK-1145 label)



EXCEL:
> [DONE] listview na karte klienta vlavo
> [DONE] vrchne listview na kontrole dat
> dblClick na KTK na karte Kontrola Dat -> co vsetko a ako tam chcu? 
> objekt Overdraft_class
  > (zatial cakam, co vsetko vlastne bude treba - vsetko stlpce by bol overkill)
  > 


------------------------
>>>   SABLONA   <<<

( riesi skript "RISK-1145 - napocet prehledu plateb (VIEWs).sql")
( TU SI SKONCIL - vymyslanie urokov v druhom skripte)
(prve KTK v EARLY vymahani cakame zaciatkom ZARI, preto docasne odlozene vsetko aby som mal poriadne data na DWH)

*/                                

SELECT * FROM owner_dwh.F_OVERDRAFT_AT;
SELECT * FROM owner_dwh.F_OVERDRAFT_BALANCE_PD;

SELECT * FROM app_coll_jezevcik_test.OVERDRAFT;
SELECT * FROM app_coll_jezevcik_test.OVERDRAFT_BALANCE_PD;
SELECT * FROM app_coll_jezevcik_test.OVERDRAFT_POSLEDNI_ZUSTATEK;

SELECT * FROM app_coll_jezevcik_test.V_OVERDRAFT_PRUBEH;








--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
--/--/--/--/          CREATE TABLE OVERDRAFT


SELECT * FROM app_coll_jezevcik_test.OVERDRAFT;

/*
DROP TABLE app_coll_jezevcik_test.OVERDRAFT_POSLEDNI_ZUSTATEK;
DROP TABLE app_coll_jezevcik_test.OVERDRAFT_BALANCE_PD;
DROP TABLE app_coll_jezevcik_test.OVERDRAFT;
*/

/*
-- riesenie pre ORA-03114 + ORA-03113:
    alter session set query_rewrite_enabled=false;
    alter session set "_query_rewrite_vop_cleanup"=false;
    alter session set "_complex_view_merging"=false;
    alter session set optimizer_features_enable = "11.1.0.7";
*/

CREATE TABLE app_coll_jezevcik_test.overdraft AS
SELECT DISTINCT 
       p.ID_PERSON
     , ov.SKP_BANK_ACCOUNT
     , ov.SKP_BANK_ACCOUNT_DEBIT
     , ov.SKP_OVERDRAFT
     , ov.CODE_OVERDRAFT_NUMBER
     , ov.AMT_OVERDRAFT_LIMIT
     , ov.DATE_OVERDRAFT_OPEN
     , ov.DATE_OVERDRAFT_UTILIZATION
     , ov.DATE_OVERDRAFT_UTIL_CURRENT                -- vsichni 1.1.3000, proto na konci jeste sloupecek DATUM_POSLEDNI_ZANORENI
     , ov.SKP_OVERDRAFT_STATUS
     , st.NAME_OVERDRAFT_STATUS
     , st.NAME_OVERDRAFT_WFSTATUS
     , ov.AMT_OVERDRAFT
     , ov.AMT_OVERDRAFT_OVERDUE
     , ov.AMT_OVERDRAFT_INTEREST_TO_PAY
     , ov.AMT_INTEREST_OVERDUE
     , ov.AMT_INTEREST_PAID
     , ov.AMT_INTEREST_PAID_TOTAL
     , ov.AMT_INTEREST_TOTAL
     , ov.AMT_INTEREST_LAST
     , ov.RATE_PENALTY
     , ov.RATE_PENALTY_INTEREST
     , ov.AMT_PAYOFF
     , ov.AMT_PAYOFF_PRINCIPAL
     , ov.AMT_PAYOFF_INTEREST
     , ov.SKP_CURRENCY
     , CASE   -- posledni zustatek je NULL pokud klient neni zanoreny
         WHEN ov.AMT_OVERDRAFT > 0 
           THEN max(ab.AMT_BALANCE)  KEEP (dense_rank FIRST ORDER BY bal.date_period DESC) OVER (PARTITION BY ov.SKP_OVERDRAFT)
         ELSE
           NULL
       END AS POSLEDNI_KLADNY_ZUSTATEK
     , CASE   -- datum je NULL pokud klient neni zanoreny
         WHEN ov.AMT_OVERDRAFT > 0 
           THEN MAX(bal.date_period) OVER (PARTITION BY ov.skp_overdraft ORDER BY bal.date_period DESC)
         ELSE
           NULL
       END AS DATUM_POSLEDNI_ZANORENI
       
FROM owner_dwh.F_OVERDRAFT_AT ov
      JOIN app_coll_jezevcik_test.person p
           ON ov.skp_person = p.SKP_PERSON
      JOIN owner_dwh.CL_OVERDRAFT_STATUS st
           ON ov.SKP_OVERDRAFT_STATUS = st.SKP_OVERDRAFT_STATUS
      LEFT JOIN owner_dwh.F_OVERDRAFT_BALANCE_PD bal    -- pro datum posledniho zanoreni
           ON bal.SKP_OVERDRAFT = ov.SKP_OVERDRAFT
           AND bal.AMT_OVERDRAFT <=0
      LEFT JOIN owner_dwh.f_account_balance_pd ab       -- pro posledni kladny zustatek
           ON ab.SKP_BANK_ACCOUNT = ov.SKP_BANK_ACCOUNT 
           AND ab.DATE_PERIOD = TRUNC(bal.date_period)
           AND ab.FLAG_DELETED = 'N'                    -- musi byt tady kvuli partitionu
WHERE ov.FLAG_DELETED = 'N'
      AND bal.FLAG_DELETED = 'N'
      AND st.NAME_OVERDRAFT_STATUS = 'Aktivní'
--      AND p.ID_PERSON = 12222
;
--SELECT * FROM v_client_status WHERE id_person = &id_skp_cuid OR cuid = &id_skp_cuid OR skp_person = &id_skp_cuid; 


    ALTER TABLE app_coll_jezevcik_test.OVERDRAFT
      ADD CONSTRAINT overdraft_primary_key PRIMARY KEY (skp_overdraft);
    ALTER TABLE app_coll_jezevcik_test.OVERDRAFT
      ADD CONSTRAINT fk_overdraft_person FOREIGN KEY (ID_PERSON) REFERENCES app_coll_jezevcik_test.person (ID_PERSON);
    -- pokud tabule ucet nema zadny primarni klic, tady ho nastaviiim
      ALTER TABLE app_coll_jezevcik_test.ucet 
        ADD CONSTRAINT pk_ucet_na_osobu PRIMARY KEY (SKP_BANK_ACCOUNT, ID_PERSON);
      ALTER TABLE app_coll_jezevcik_test.ucet
        ADD CONSTRAINT fk_ucet_person FOREIGN KEY (ID_PERSON) REFERENCES app_coll_jezevcik_test.person (ID_PERSON);
      -- pozor, v procce je potreba nahradit truncate za delete:
        -- (You cannot truncate a table with an enabled foreign key that points to it. Truncate does not fire any triggers, does not validate any constraints. It does not care of the child table is empty or not)
        -- old: execute immediate 'truncate table  app_coll_jezevcik_test.ucet drop storage'; -- nelze po pridani foreign key, RISK-1145
        -- new: DELETE FROM app_coll_jezevcik_test.ucet;
    -- ted uz muzu napojit i overdraft
    ALTER TABLE app_coll_jezevcik_test.OVERDRAFT
      ADD CONSTRAINT fk_overdraft_ucet FOREIGN KEY (SKP_BANK_ACCOUNT, ID_PERSON) REFERENCES app_coll_jezevcik_test.ucet (SKP_BANK_ACCOUNT, ID_PERSON);

    COMMENT ON TABLE app_coll_jezevcik_test.OVERDRAFT IS 'Hlavní tabulka kontokorentù.';
            
    SELECT listagg(username, ',') WITHIN GROUP (ORDER BY username) AS usernames FROM jez_users WHERE dtime_valid_to > SYSDATE; 
    GRANT SELECT ON app_coll_jezevcik_test.OVERDRAFT TO AIR_ANTOL,AIR_DOUBKOVA,AIR_HALCIN,AIR_KLOUDOVAE,AIR_MACICEKM,AIR_RAJSIGLOVA,AIR_SEIDLOVA,AIR_SEVERA,AIR_SIMKOVA,AIR_TOMESKOVA;
    GRANT ALL ON app_coll_jezevcik_test.OVERDRAFT TO air_macicekm, air_severa, air_antol;
    


SELECT * FROM app_coll_jezevcik_test.OVERDRAFT;

-- generator komentaru z DWH tabulky
(SELECT (' IS ' || '''' || comments || '''' || ';') AS comment_string FROM DBA_COL_COMMENTS a WHERE a.table_name = 'F_OVERDRAFT_AT' AND a.column_name = 'SKP_CURRENCY');

skript "generate COLUMN comments.sql"







--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
--/--/--/--/           CREATE TABLE OVERDRAFT_BALANCE_PD


SELECT * FROM owner_dwh.F_OVERDRAFT_BALANCE_PD;
SELECT * FROM owner_dwh.CL_OVERDRAFT_STATUS;


--DROP TABLE app_coll_jezevcik_test.OVERDRAFT_BALANCE_PD;
CREATE TABLE app_coll_jezevcik_test.OVERDRAFT_BALANCE_PD AS
SELECT ov.ID_PERSON
     , bal.SKP_OVERDRAFT
     , bal.DATE_EFFECTIVE
     , bal.SKP_CONTRACT
     , bal.SKP_OVERDRAFT_STATUS
     , st.NAME_OVERDRAFT_STATUS
     , st.NAME_OVERDRAFT_WFSTATUS
    /* , bal.SKP_OVERDRAFT_STATUS*/
    /* , bal.DTIME_INSERTED*/
    /* , bal.DTIME_MODIFIED*/
     , bal.DATE_PERIOD
     , bal.AMT_OVERDRAFT
     , bal.AMT_OVERDRAFT_LIMIT
     , bal.AMT_OVERDRAFT_OVERDUE
     , bal.AMT_PRINCIPAL_OVERDUE
     , bal.AMT_INTEREST_OVERDUE
     , bal.AMT_PRINCIPAL_FUTURE
     , bal.AMT_INTEREST_DUE
     , bal.AMT_INTEREST_PAID
     , bal.AMT_PRINCIPAL_PAID
     , bal.AMT_INTEREST_PAID_TOTAL
     , bal.AMT_PRINCIPAL_PAID_TOTAL
     , bal.AMT_INTEREST
     , bal.AMT_INTEREST_TOTAL
     , bal.AMT_INTEREST_LAST 
     , bal.AMT_PENALTY_CURRENT 
     , bal.AMT_PENALTY_INTEREST_CURRENT  
     , bal.AMT_PRINCIPAL_DRAW  
     , bal.RATE_PENALTY  
     , bal.RATE_PENALTY_INTEREST
     , bal.date_payoff
     , bal.AMT_PAYOFF
     , bal.AMT_PAYOFF_PRINCIPAL
     , bal.AMT_PAYOFF_INTEREST
FROM owner_dwh.F_OVERDRAFT_BALANCE_PD bal
JOIN app_coll_jezevcik_test.OVERDRAFT ov ON bal.skp_overdraft = ov.skp_overdraft
JOIN owner_dwh.CL_OVERDRAFT_STATUS st
  ON bal.SKP_OVERDRAFT_STATUS = st.SKP_OVERDRAFT_STATUS
WHERE bal.FLAG_DELETED = 'N'
  --AND st.NAME_OVERDRAFT_STATUS = 'Aktivní'
;

    ALTER TABLE app_coll_jezevcik_test.OVERDRAFT_BALANCE_PD
      ADD CONSTRAINT overdraft_balance_pk PRIMARY KEY (skp_overdraft, date_period);
    ALTER TABLE app_coll_jezevcik_test.OVERDRAFT_BALANCE_PD
      ADD CONSTRAINT fk_bal_skp_overdraft FOREIGN KEY (SKP_OVERDRAFT) REFERENCES app_coll_jezevcik_test.OVERDRAFT (SKP_OVERDRAFT);
    ALTER TABLE app_coll_jezevcik_test.OVERDRAFT_BALANCE_PD
      ADD CONSTRAINT fk_overdraft_bal_person FOREIGN KEY (ID_PERSON) REFERENCES app_coll_jezevcik_test.person (ID_PERSON);

    COMMENT ON TABLE app_coll_jezevcik_test.OVERDRAFT_BALANCE_PD IS 'Tabulka zachycuje denní balance kontokorentù.';
            
    SELECT listagg(username, ',') WITHIN GROUP (ORDER BY username) AS usernames FROM jez_users WHERE dtime_valid_to > SYSDATE; 
    GRANT SELECT ON app_coll_jezevcik_test.OVERDRAFT_BALANCE_PD TO AIR_ANTOL,AIR_DOUBKOVA,AIR_HALCIN,AIR_KLOUDOVAE,AIR_MACICEKM,AIR_RAJSIGLOVA,AIR_SEIDLOVA,AIR_SEVERA,AIR_SIMKOVA,AIR_TOMESKOVA;
    GRANT ALL ON app_coll_jezevcik_test.OVERDRAFT_BALANCE_PD TO air_macicekm, air_severa, air_antol;
            


SELECT * FROM app_coll_jezevcik_test.OVERDRAFT_BALANCE_PD;

-- generator komentaru z DWH tabulky

skript "generate COLUMN comments.sql"








--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
--/--/--/--/           CREATE TABLE OVERDRAFT_POSLEDNI_ZUSTATEK

-- TABULE SE ZAHAZUJE, NENI POTREBNA (STACI SLOUPECEK V TABULI OVERDRAFT)

/*
    -- TABULKA POSLEDNI ZUSTATEK
      CREATE TABLE app_coll_jezevcik_test.OVERDRAFT_POSLEDNI_ZUSTATEK AS
        SELECT o.SKP_OVERDRAFT, ab.AMT_BALANCE AS posledni_kladny_zustatek
        FROM app_coll_jezevcik_test.OVERDRAFT o
        JOIN app_coll_jezevcik_test.V_OVERDRAFT_DATUM_ZANORENI dtm ON dtm.SKP_OVERDRAFT = o.SKP_OVERDRAFT
        JOIN owner_dwh.f_account_balance_pd ab ON ab.SKP_BANK_ACCOUNT = o.SKP_BANK_ACCOUNT AND ab.DATE_PERIOD = TRUNC(dtm.datum_posledni_zanoreni)
        WHERE ab.FLAG_DELETED = 'N'
        ;
        
        ALTER TABLE app_coll_jezevcik_test.OVERDRAFT_POSLEDNI_ZUSTATEK
          ADD CONSTRAINT overdraft_zustatek_pk PRIMARY KEY (SKP_OVERDRAFT);
        ALTER TABLE app_coll_jezevcik_test.OVERDRAFT_POSLEDNI_ZUSTATEK
          ADD CONSTRAINT fk_skp_overdraft FOREIGN KEY (SKP_OVERDRAFT) REFERENCES app_coll_jezevcik_test.OVERDRAFT (SKP_OVERDRAFT);
        COMMENT ON TABLE app_coll_jezevcik_test.overdraft_posledni_zustatek IS 'Tabulka zachycuje poslední kladný zùstatek na úètu klienta pøed zanoøením.';
        COMMENT ON COLUMN app_coll_jezevcik_test.overdraft_posledni_zustatek.POSLEDNI_KLADNY_ZUSTATEK IS 'poslední kladný zùstatek na úètu klienta pøed zanoøením.';
        
        SELECT listagg(username, ',') WITHIN GROUP (ORDER BY username) AS usernames FROM jez_users WHERE dtime_valid_to > SYSDATE; 
        GRANT SELECT ON app_coll_jezevcik_test.OVERDRAFT_POSLEDNI_ZUSTATEK TO AIR_ANTOL,AIR_DOUBKOVA,AIR_HALCIN,AIR_KLOUDOVAE,AIR_MACICEKM,AIR_RAJSIGLOVA,AIR_SEIDLOVA,AIR_SEVERA,AIR_STAROBOVA,AIR_TOMESKOVA;
        GRANT ALL ON app_coll_jezevcik_test.OVERDRAFT_POSLEDNI_ZUSTATEK TO air_macicekm, air_severa, air_antol;
        
*/




--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
--/--/--/--/           UPDATE VIEWs v_zesplatneni

-- v_nezesplatneni -- CAJK, netyka sa KTK
-- SELECT * FROM v_zaloby;   -- CAJK
-- SELECT * FROM V_TREST;    -- CAJK
-- SELECT * FROM V_DEDICTVI; -- CAJK
-- SELECT * FROM v_exe;      -- CAJK
-- SELECT * FROM v_ins;      -- CAJK

SELECT * FROM app_coll_jezevcik_test.V_ZESPLATNENI;         -- with konstrukce pribudla
SELECT * FROM app_coll_jezevcik_test.v_datacheck_contracts; -- posledny union pribudol


-- nahlad
SELECT * FROM app_coll_jezevcik_test.v_datacheck_contracts a WHERE EXISTS (SELECT id_person FROM overdraft ov WHERE ov.ID_PERSON = a.ID_PERSON);
-- CUID 495598
SELECT * FROM app_coll_jezevcik_test.ucet u WHERE u.ID_PERSON = 11293;
SELECT * FROM app_coll_jezevcik_test.OVERDRAFT ov WHERE ov.ID_PERSON = 11293;
SELECT * FROM owner_dwh.F_BANK_ACCOUNT_AD a WHERE a.skp_bank_account IN (17619172,17628275); -- ten prvni, 17619172, je BU klienta
SELECT * FROM owner_dwh.CL_BANK_ACCOUNT_TYPE a WHERE a.SKP_BANK_ACCOUNT_TYPE = 21355; -- Splatkovy ucet, (ten druhy, 17628275)


-- EXCEL - dotaz do Klient.listview2 (mod_Query)
SELECT a.ID_PERSON, a.CISLO FROM app_coll_jezevcik_test.v_datacheck_contracts a WHERE a.TYP_SMLOUVY IN ('HU', 'HR', 'KTK');
SELECT typ_smlouvy, COUNT(typ_smlouvy) FROM app_coll_jezevcik_test.V_DATACHECK_CONTRACTS GROUP BY typ_smlouvy;






--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
--/--/--/--/          PLNENI novych tabulek v procke pres MERGE INTO


--  iba prepis create table na MERGE INTO (insert into) 
-- > overdraft
-- > overdraft balance


skript "plneni tabulek do procky.sql"




--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
--/--/--/--/          
