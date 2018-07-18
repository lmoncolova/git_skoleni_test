CREATE OR REPLACE VIEW V_OVERDRAFT_PRUBEH AS
SELECT A.SKP_BANK_ACCOUNT
      ----   hlavni view pro generovani prehledu plateb ke kontokorentu
      ----   ORDER BY klauzule - potreba aby tam krome datumu byl i nejaky unique sloupecek - aby vysledek byl vzdy jednoznacny
      ,O.SKP_OVERDRAFT
      ,A.ID_PERSON
      ,A.AMT_TRANSACTION
      ,CASE row_number() OVER (PARTITION BY o.skp_overdraft ORDER BY A.DTIME_TRANSACTION_ACCOUNTED ASC, a.AMT_TRANSACTION DESC, a.SKF_REALIZED_TRANSACTION)
         WHEN 1 THEN A.AMT_TRANSACTION + o.POSLEDNI_KLADNY_ZUSTATEK
         ELSE A.AMT_TRANSACTION
       END AS AMT_TRANSACTION_CLEAN
      /* -- kontroly jestli partitions delaji to co maji
      ,row_number() OVER (PARTITION BY o.skp_overdraft ORDER BY A.DTIME_TRANSACTION_ACCOUNTED ASC, a.AMT_TRANSACTION DESC, a.SKF_REALIZED_TRANSACTION) AS poradi_transakce
      ,CASE row_number() OVER (PARTITION BY o.skp_overdraft ORDER BY A.DTIME_TRANSACTION_ACCOUNTED ASC, a.AMT_TRANSACTION DESC, a.SKF_REALIZED_TRANSACTION)
         WHEN 1 THEN 'prvyy'
         ELSE 'neprvyy'
       END AS AMT_TRANSACTION_CLEAN*/

      ,A.DTIME_TRANSACTION_ACCOUNTED
      --,A.TEXT_NOTE
      ,CASE /*a.TEXT_NOTE*/
         WHEN (a.TEXT_NOTE LIKE '%Obchodní úrok kontokorentu%')
           THEN 'Obchodní úrok (' || GREATEST(LAST_DAY(ADD_MONTHS(a.DTIME_TRANSACTION_ACCOUNTED,-2))+1, o.datum_posledni_zanoreni) || ' - ' || LAST_DAY(ADD_MONTHS(a.DTIME_TRANSACTION_ACCOUNTED,-1)) || ')'
         ELSE to_char(a.TEXT_NOTE)
       END AS text_note

      ,o.DATUM_POSLEDNI_ZANORENI
      ,o.POSLEDNI_KLADNY_ZUSTATEK
      ,bal.AMT_OVERDRAFT

      ,(-1) * (o.POSLEDNI_KLADNY_ZUSTATEK + SUM(a.AMT_TRANSACTION)
         OVER (PARTITION BY o.skp_overdraft ORDER BY A.DTIME_TRANSACTION_ACCOUNTED ASC, a.AMT_TRANSACTION DESC, a.SKF_REALIZED_TRANSACTION)) AS AMT_OVERDRAFT_CUMULATIVE
      /*CASE row_number() OVER (PARTITION BY o.skp_overdraft, TRUNC(a.dtime_transaction_accounted) ORDER BY A.DTIME_TRANSACTION_ACCOUNTED ASC, a.AMT_TRANSACTION DESC, a.SKF_REALIZED_TRANSACTION)
        WHEN 1 THEN*/

      -- kontrola jestli suma transakci v dany den je rovna zmene amt_overdraft za dany den
        ,CASE
           WHEN (trunc(A.DTIME_TRANSACTION_ACCOUNTED) > trunc(o.datum_posledni_zanoreni) +1  -- vsechny dny krome prvniho
                 AND SUM(nvl(a.AMT_TRANSACTION,0)) OVER (PARTITION BY a.SKP_BANK_ACCOUNT, TRUNC(a.DTIME_TRANSACTION_ACCOUNTED)) = NVL((bal.AMT_PRINCIPAL_PAID - bal.amt_principal_draw),0)) THEN 'Y'
           WHEN (trunc(A.DTIME_TRANSACTION_ACCOUNTED) = trunc(o.datum_posledni_zanoreni) +1   -- v prvni den musime brat v potaz zustatek na uctu
                 AND SUM(nvl(a.AMT_TRANSACTION,0)) OVER (PARTITION BY a.SKP_BANK_ACCOUNT, TRUNC(a.DTIME_TRANSACTION_ACCOUNTED)) = NVL((bal.AMT_PRINCIPAL_PAID - bal.amt_principal_draw - o.POSLEDNI_KLADNY_ZUSTATEK),0)) THEN 'Y'
           ELSE 'N'
         END AS FLAG_SUMA_ZA_DEN_OK

      -- kontrola jestli suma transakci od posledniho zanoreni je rovna aktualnimu amt_overdraft
        ,CASE
           WHEN SUM(nvl(a.AMT_TRANSACTION,0)) OVER (PARTITION BY a.SKP_BANK_ACCOUNT) = ((-1) * (nvl(o.AMT_OVERDRAFT,0) + o.POSLEDNI_KLADNY_ZUSTATEK)) THEN 'Y'
           ELSE 'N'
         END AS FLAG_SUMA_CELKOVA_OK

        /*-- promenne ke kontrole kontrol
        ,SUM(nvl(a.AMT_TRANSACTION,0)) OVER (PARTITION BY a.SKP_BANK_ACCOUNT, TRUNC(a.DTIME_TRANSACTION_ACCOUNTED)) AS suma_transakci_za_den
        ,NVL((bal.AMT_PRINCIPAL_PAID - bal.amt_principal_draw),0) AS amt_principal_zmena

        ,SUM(nvl(a.AMT_TRANSACTION,0)) OVER (PARTITION BY a.SKP_BANK_ACCOUNT) suma_vsech_plateb
        ,nvl(o.AMT_OVERDRAFT,0) + ps.POSLEDNI_KLADNY_ZUSTATEK suma_ktk_a_zustatek
        */

  FROM APP_COLL_JEZEVCIK_TEST.PLATBY_UCET A
       JOIN app_coll_jezevcik_test.OVERDRAFT o 
         ON o.SKP_BANK_ACCOUNT = a.SKP_BANK_ACCOUNT                      -- tenhle join neni uplne spravne, protoze klient mohl mit vice KTK na jeden ucet. Nicmene to nevadi ze 2 duvodu: 1. kazda platba ma jednoznacny datum a ke kazdemu datumu mohl byt jenom jeden KTK, 2. ve vymahani nas vsechno zajima az od posledniho zanoreni, takze se to zamotat nemuze
       JOIN app_coll_jezevcik_test.OVERDRAFT_BALANCE_PD bal 
         ON bal.SKP_OVERDRAFT = o.SKP_OVERDRAFT 
        AND bal.DATE_PERIOD = TRUNC(a.DTIME_TRANSACTION_ACCOUNTED)
 WHERE trunc(a.DTIME_TRANSACTION_ACCOUNTED) > trunc(o.datum_posledni_zanoreni)                          -- podstatni podminka ktera zaruci, ze se platby k druhemu KTK klienta nezapocitaji k jeho prvnimu KTK (v pripade ze jich mel vic)
   --AND o.SKP_OVERDRAFT = 10
   /*AND ucet_bal.FLAG_DELETED = 'N'*/
   --AND A.SKP_BANK_ACCOUNT = 15270448
 ORDER BY A.DTIME_TRANSACTION_ACCOUNTED ASC, a.AMT_TRANSACTION DESC, a.SKF_REALIZED_TRANSACTION
;
