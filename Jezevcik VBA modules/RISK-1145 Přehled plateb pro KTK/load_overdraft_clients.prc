create or replace procedure load_overdraft_clients is
begin
  
-- vytvoril Filip Antol 10.7.2018 v ramci RISK-1145
-- cilem je dostat do testovaciho Jezevcika kontokorenty, tahle procedura to udela tak, ze da vstupenku do testovaciho jezevcika vsem klientum s kontokorentem.
-- k nacteni KTK pak je potreba jeste spustit package
-- procedura slouzi pouze k vyvoji !!!
-- zavoalat je potreba v PCKG nekde kolem radku 100:
    /*
    BEGIN
      load_overdraft_clients; -- RISK-1145
    END;
    */



INSERT INTO app_coll_jezevcik_test.OVERDRAFT_ZES (SKP_PERSON,
                                                  SKP_OVERDRAFT,
                                                  DATE_PAYOFF,
                                                  NAME_PAYOFF_REASON,
                                                  NAME_OVERDRAFT_STATUS,
                                                  NAME_OVERDRAFT_WFSTATUS)
    SELECT a.skp_person
         , a.SKP_OVERDRAFT
         , a.DATE_PAYOFF
         , pr.NAME_PAYOFF_REASON
         , os.NAME_OVERDRAFT_STATUS
         , os.NAME_OVERDRAFT_WFSTATUS
    FROM owner_dwh.F_OVERDRAFT_AT a
    JOIN owner_dwh.CL_PAYOFF_REASON pr ON a.SKP_PAYOFF_REASON = pr.SKP_PAYOFF_REASON
    JOIN owner_dwh.CL_OVERDRAFT_STATUS os ON a.SKP_OVERDRAFT_STATUS = os.SKP_OVERDRAFT_STATUS
    WHERE os.NAME_OVERDRAFT_STATUS = 'Aktivní'
      -- !!!!!!!!!!!!!!
      -- AND os.NAME_OVERDRAFT_WFSTATUS = 'Zesplatnìný' -- RISK-1145  odkomentovat na PROD !!!!!!!!!
      -- !!!!!!!!!!!!!!
;


-- pojistka, kdyby nekdo omylem vytvoril a spustil proceduru na produkci
IF nazev_aktualniho_hriste = 'APP_COLL_JEZEVCIK_TEST' THEN
  COMMIT;
ELSE
  ROLLBACK;  
END IF;



end load_overdraft_clients;
/
