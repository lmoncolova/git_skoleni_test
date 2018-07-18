SELECT * FROM app_coll_jezevcik_test.V_OVERDRAFT_PRUBEH a WHERE a.SKP_OVERDRAFT = 10;
SELECT * FROM app_coll_jezevcik_test.OVERDRAFT a WHERE a.SKP_OVERDRAFT = 10;

SELECT * FROM person WHERE cuid in ('835900', '61425', '389243', '238996', '52572', '336606', '429952', '208', '788279', '645705', '541242');

select * from APP_COLL_JEZEVCIK_TEST.loan where id_person = 11340 order by date_loan_utilization;
SELECT * FROM APP_COLL_JEZEVCIK_TEST.loan a WHERE a.CODE_LOAN_NUMBER = 'HU00508800';

SELECT * FROM app_coll_jezevcik_test.V_DATACHECK_CONTRACTS a WHERE a.ID_PERSON = 11340;

/*(SELECT DISTINCT A.ID_PERSON
                 ,A.CISLO_UCTU                 AS CISLO
                 ,A.NAME_BANK_ACCOUNT_STATUS   AS STATUS_1
                 ,A.NAME_BANK_ACCOUNT_WFSTATUS AS STATUS_2
                 ,A.TYP_UCTU                   AS TYP_SMLOUVY
    FROM APP_COLL_JEZEVCIK_TEST.UCET A
    WHERE a.TYP_UCTU <> 'KTK'
      AND ID_PERSON = 11340
    )

UNION*/

 (SELECT DISTINCT B.ID_PERSON
                 ,B.CODE_LOAN_NUMBER
                 ,B.NAME_LOAN_STATUS             AS STATUS_1
                 ,B.NAME_LOAN_WFSTATUS           AS STATUS_2
                 ,SUBSTR(CODE_LOAN_NUMBER, 1, 2) AS TYP_SMLOUVY
    FROM APP_COLL_JEZEVCIK_TEST.INTERVALY B 
   WHERE ID_PERSON = 11340)
;

SELECT * FROM v_client_status WHERE id_person = &id_skp_cuid OR cuid = &id_skp_cuid OR skp_person = &id_skp_cuid; 
