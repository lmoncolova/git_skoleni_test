CREATE OR REPLACE VIEW V_OVERDRAFT_DATUM_ZANORENI AS
SELECT a.SKP_OVERDRAFT
           -- , MIN(a.DATE_PERIOD) AS datum_prvni_zanoreni -- toto je zle, treba cez KEEP (z minimalnych datumov chcem ten najvacsi). Ale netreba mi to.
           , MAX(a.DATE_PERIOD) AS datum_posledni_zanoreni
      FROM app_coll_jezevcik_test.OVERDRAFT_BALANCE_PD a
      WHERE a.AMT_OVERDRAFT <=0
      GROUP BY a.SKP_OVERDRAFT
      ORDER BY a.SKP_OVERDRAFT DESC
;
