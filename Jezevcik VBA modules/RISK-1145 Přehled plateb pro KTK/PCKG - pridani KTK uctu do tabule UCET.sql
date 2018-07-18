----    pckg - pridani KTK uctu do tabule UCET
--  kod v package potrebuje jenom malinko upravit - tam kde jsou komentare RISK-1145

--pridani exekucnich, insolvencnich, jpk a KTK uctu, RISK-276
select /*+ parallel(bb 4) use_hash(ba s) leading(p ba) */ p.id_person,
       ba.SKP_BANK_ACCOUNT,
       ba.text_bank_account_number,
       CASE ba.SKP_BANK_ACCOUNT_type
         WHEN 21130 THEN 'EÚ'
         when 21312 THEN 'JPK'
         WHEN 21355 THEN 'KTK'  -- RISK-1145
         ELSE 'IÚ'
       end as typ_uctu,
       trunc(sysdate) - 1,
       case
         when ba.SKP_BANK_ACCOUNT_TYPE = 21312 then
          - (bb.AMT_BALANCE_AVAILABLE + bb.AMT_GOVERNMENT_HOLD_SUM)
         else
          (bb.AMT_BALANCE_AVAILABLE + bb.AMT_GOVERNMENT_HOLD_SUM)
       end as amt_balance,
       'CZK',
       sysdate,
       s.name_bank_account_status,
       s.name_bank_account_wfstatus,
        bb.AMT_GOVERNMENT_HOLD_SUM, --RISK-389
       ba.CNT_DAYS_OVERDUE_TOL_A,
       ba.CNT_DAYS_OVERDUE_TOL_B,
       case
         when ba.SKP_BANK_ACCOUNT_TYPE = 21312 then
          - (bb.AMT_BALANCE_AVAILABLE)
         else
          (bb.AMT_BALANCE_AVAILABLE)
       end as amt_balance_available,
       case
         when ba.SKP_BANK_ACCOUNT_TYPE = 21312 then
          - (bb.AMT_BALANCE)
         else
          (bb.AMT_BALANCE)
       end as amt_balance_ucetni,
       ba.DATE_BANK_ACCOUNT_ACTIVATION
  from owner_dwh.f_bank_account_ad ba
  left join owner_dwh.f_account_balance_td bb
    on bb.SKP_BANK_ACCOUNT = ba.SKP_BANK_ACCOUNT
   and bb.date_balance_valid_to > sysdate
   and bb.FLAG_DELETED = 'N'
  left join owner_dwh.cl_bank_account_status s
    on ba.skp_bank_account_status = s.skp_bank_account_status
  join APP_COLL_JEZEVCIK_TEST.person p
    on p.skp_person = ba.skp_person
   and bb.skp_bank_account_type in (21130, 21270, 21312, 21355) -- 21355 = Splátkový úèet (INSTALLMENT) pro KTK  --RISK-1145
   and p.flag_valid = 'Y';
