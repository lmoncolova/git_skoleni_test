    -- V případě, kdy je na klientovi uložen jeden ze 3 uvedených stavů a zároveň neexistuje na klientovi zadaná šance s typem SK určený soudem,
    -- vyskočí v jezevčíkovi upozornění (klidně až na druhý den) s poznámkou "u klienta není zaevidovaný SK"
    insert /*+ APPEND */ into notice t (t.ID_COLLECTION_CASE, t.TEXT_NOTICE, t.DTIME_NOTICE_CREATED, t.DTIME_NOTICE_DEADLINE, t.CODE_STATUS,
                                                 t.EMPLOYEE, t.ASSIGNEE, t.SKP_PROCESS, t.SKP_NOTICE_TYPE)
      (SELECT sd.ID_COLLECTION_CASE
           , nt.text_notice
           , trunc(sysdate)
           , trunc(sysdate) + nt.DAYS_TO_DEADLINE
           , 'a'
           , 'sys'
           , nt.assignee
           , nt.skp_process
           , nt.skp_notice_type

          FROM epr_status_data sd
          JOIN epr_status st
            ON sd.skp_epr_status = st.skp_epr_status
              JOIN NOTICE_TYPE nt
                ON nt.CODE_NOTICE_TYPE = 'CLIENT_MISSING_SK'
          JOIN collection_case cc
            ON sd.id_collection_case = cc.id_collection_case
           AND cc.typ_klienta = 'AB'
          LEFT JOIN (select *
                       from chance_import
                       where TYP_SK = 'SK určený soudem'
                         and flag_deleted = 'N'
                     ) chance
                 ON chance.ID_COLLECTION_CASE = sd.ID_COLLECTION_CASE
          LEFT JOIN (SELECT DISTINCT esd.id_collection_case
                       FROM exe_status_data esd
                       JOIN collection_case cc
                         ON esd.id_collection_case = cc.id_collection_case
                        AND cc.typ_klienta = 'AB'
                       JOIN exe_status es
                         ON esd.skp_exe_status = es.skp_exe_status
                        AND es.name_exe_status = 'Exekuční návrh'
                        AND esd.flag_deleted = 'N'
                       ) exe
             ON cc.id_collection_case = exe.id_collection_case
          WHERE st.name_epr_status IN ('PM rozsudek – SK s doložkou', 'PM Rozsudek - SK částečně zamítnuto s doložkou', 'PM Usnesení o schválení smíru s doložkou')
            AND st.dtime_valid_to > SYSDATE
            AND sd.dtime_valid_to > SYSDATE
            AND sd.flag_deleted = 'N'
            AND chance.ID_COLLECTION_CASE IS NULL
            AND cc.code_insolvency_case_status IS NULL
            AND exe.id_collection_case IS NULL
            AND NOT EXISTS (SELECT * -- RISK-930, nechceme aby upozorneni vyskakovalo opakovane
                            FROM notice x
                             JOIN collection_case cc
                               ON x.id_collection_case = cc.id_collection_case
                              AND cc.typ_klienta = 'AB'
                            WHERE x.ID_COLLECTION_CASE = sd.ID_COLLECTION_CASE
                                  AND x.SKP_NOTICE_TYPE = nt.SKP_NOTICE_TYPE
                              AND x.code_status <> 's'
                              AND x.dtime_notice_created > SYSDATE - 60
                           )
          )
    ;
    COMMIT;
