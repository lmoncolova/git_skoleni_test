-- RISK-1135
-- Pøehled plateb pro hypotéky

-- faze 1 je vytvoreni struktury pro chytrou rezervu (OTB = Open To Buy)
-- Jezevcik by ji pak vyexportil do zvlast excelu




--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
--/--/--/--/          PPP

[DONE] najst DATA na dwh
vytvorit v Jezevcovi tabulku pre pohyby chytre rezervy
doplnit do PCKG napocet tejto tabulky
? mal BY sa OTB ucet dotahovat do tabulky ucet? (logicky hej, je TO predsa na klienta..)


EXCEL:
na karte Kontrola Dat pridat tlacitko pre export chytrej rezervy do excelu (ENABLED iba pre HR)








--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
--/--/--/--/          nahlady

    -- klientka zo zadania
      SELECT * FROM v_client_status WHERE cuid = 673460;
      &id_skp_cuid OR cuid = &id_skp_cuid OR skp_person = &id_skp_cuid; 
      SELECT * FROM loan u WHERE u.ID_PERSON = 635;
    SELECT * FROM owner_dwh.f_loan_at a WHERE a.skp_loan = 12888223;
    SELECT * FROM owner_dwh.f_loan_balance_pm flb WHERE flb.skp_loan = 12888223;
    
    SELECT * FROM  owner_dwh.f_account_loan_rel_at a WHERE a.SKP_LOAN = 12888223;
    SELECT DISTINCT fab.SKP_BANK_ACCOUNT, fab.SKP_BANK_ACCOUNT_type FROM owner_dwh.f_account_balance_td fab WHERE fab.skp_person = 4458842;

--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
--/--/--/--/          MAIN: pohyby chytre rezervy k hypotece


CREATE TABLE app_coll_jezevcik_test.LOAN_OTB_BALANCE AS
    SELECT p.ID_PERSON
         , l.SKP_LOAN
         , l.CODE_LOAN_NUMBER            -- duplicita pro nasi citelnost
         , falr.SKP_BANK_ACCOUNT         -- ucet OTB
         , fab.DATE_BALANCE_VALID_FROM
         , fab.DATE_BALANCE_VALID_TO
         , fab.AMT_BALANCE
         , fab.AMT_BALANCE_AVAILABLE
         , fab.AMT_CREDIT
         , fab.AMT_DEBIT
         , (fab.AMT_CREDIT - fab.AMT_DEBIT) AS zmena_rezervy
         --, fab.*
    FROM app_coll_jezevcik_test.person p
     JOIN app_coll_jezevcik_test.loan l 
       ON p.ID_PERSON = l.ID_PERSON 
      AND l.CODE_LOAN_NUMBER LIKE 'HR%'                         -- neni potrebne jenom pro ctivost (staci nize join na ucet 21290, coz je ucet OTB)
     JOIN owner_dwh.f_account_loan_rel_at falr   
       ON falr.SKP_LOAN = l.SKP_LOAN
     JOIN owner_dwh.f_account_balance_td fab 
       ON fab.SKP_PERSON = p.SKP_PERSON 
      AND fab.SKP_BANK_ACCOUNT_TYPE = 21290 
      AND fab.SKP_BANK_ACCOUNT = falr.SKP_BANK_ACCOUNT 
    WHERE fab.FLAG_DELETED = 'N'
      AND falr.FLAG_DELETED = 'N'
    --  AND l.skp_loan = 12888223
    -- ORDER BY fab.DATE_BALANCE_VALID_TO DESC
    ;
    
-- KEYS:
    ALTER TABLE app_coll_jezevcik_test.LOAN_OTB_BALANCE
      ADD CONSTRAINT otb_balance_pk PRIMARY KEY (ID_PERSON, SKP_LOAN, DATE_BALANCE_VALID_FROM);
    ALTER TABLE app_coll_jezevcik_test.LOAN_OTB_BALANCE
      ADD CONSTRAINT otb_person_fk FOREIGN KEY (ID_PERSON) REFERENCES app_coll_jezevcik_test.person (ID_PERSON);
    ALTER TABLE app_coll_jezevcik_test.LOAN_OTB_BALANCE
      ADD CONSTRAINT otb_loan_fk FOREIGN KEY (SKP_LOAN) REFERENCES app_coll_jezevcik_test.loan (SKP_LOAN);
    cizi klic pro SKP_BANK_ACCOUNT nevytvoren, protoze tyhle technicke ucty tam nemame (TO je otazka, jestli je tam pridavat jenom kvuli tomu abych TO pak provazal, prijde mi TO zbytecne)
      - jedna klientka ma klidne i 20 nejakych uctu.

-- komentare:
    COMMENT ON TABLE app_coll_jezevcik_test.LOAN_OTB_BALANCE IS 'Tabulka balancí pro chytrou rezervu k hypotékam.';
    comment ON COLUMN ... IS '';

-- granty:
    SELECT listagg(username, ',') WITHIN GROUP (ORDER BY username) AS usernames FROM jez_users WHERE dtime_valid_to > SYSDATE; 
    GRANT SELECT ON app_coll_jezevcik_test.LOAN_OTB_BALANCE TO AIR_ANTOL,AIR_DOUBKOVA,AIR_HALCIN,AIR_KLOUDOVAE,AIR_MACICEKM,AIR_RAJSIGLOVA,AIR_SEIDLOVA,AIR_SEVERA,AIR_SIMKOVA,AIR_STAROBOVA,AIR_TOMESKOVA;
    GRANT ALL ON app_coll_jezevcik_test.LOAN_OTB_BALANCE TO AIR_ANTOL,air_macicekm, air_severa;

-- PCKG:
    v pripade FOREIGN KEY na tabuli UCET:  line 1956 PRED promazani tabulky UCET:  DELETE FROM app_coll_jezevcik_test.LOAN_OTB_BALANCE;  -- RISK-1135 , foreign key
    napocet niekam dat - asi cez MERGE INTO


-- vysledok:
    ;
    SELECT * FROM app_coll_jezevcik_test.LOAN_OTB_BALANCE;
    SELECT * FROM v_client_status WHERE id_person = &id_skp_cuid OR cuid = &id_skp_cuid OR skp_person = &id_skp_cuid; 





--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
------------                POVODNY SELECT Z DWH
    SELECT A.SKP_PERSON
         , a.SKP_LOAN
         , falr.SKP_BANK_ACCOUNT -- ucet OTB
         , fab.DATE_BALANCE_VALID_FROM
         , fab.DATE_BALANCE_VALID_TO
         , fab.AMT_BALANCE
         , fab.AMT_BALANCE_AVAILABLE
         , fab.AMT_CREDIT
         , fab.AMT_DEBIT
         , fab.AMT_CREDIT - fab.AMT_DEBIT AS zmena_rezervy
         --, fab.*
    FROM owner_dwh.f_loan_at a 
     JOIN owner_dwh.f_account_loan_rel_at falr   
       ON falr.SKP_LOAN = a.SKP_LOAN
     JOIN owner_dwh.f_account_balance_td fab 
       ON fab.SKP_PERSON = a.SKP_PERSON 
      AND fab.SKP_BANK_ACCOUNT_TYPE = 21290 
      AND fab.SKP_BANK_ACCOUNT = falr.SKP_BANK_ACCOUNT 
    WHERE a.skp_loan = 12888223
      AND A.FLAG_DELETED = 'N'
      AND falr.FLAG_DELETED = 'N'
      AND fab.FLAG_DELETED = 'N'
    ORDER BY fab.DATE_BALANCE_VALID_TO DESC
    ;
