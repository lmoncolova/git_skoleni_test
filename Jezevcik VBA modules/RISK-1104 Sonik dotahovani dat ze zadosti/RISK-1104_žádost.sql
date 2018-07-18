create table app_coll_jezevcik.sonik_apl_data_ktk as 

select fa.skp_application
     , fo.SKP_OVERDRAFT
     , sp.skp_person
     , sp.cuid
     , dp.name_first as Jmeno
     , dp.name_last as Prijmeni  
     , dp.date_birth as Datum_narozeni
     , to_date(bf.dtime_started, 'yyyy-mm-dd hh24:mi:ss') as cas2scoring
     , cs.dtime_completion_subj_inserted as cas_odeslani_zIB
     , cat.name_application_type as Druh_zadosti
     , fa.text_application_number as Cislo_zadosti 
     , fo.CODE_OVERDRAFT_NUMBER as Cislo_KTK
     , nvl(fo.AMT_OVERDRAFT_LIMIT, fa.amt_overdraft_limit) as povolena_vyse_KTK
     , case when (fo.DATE_ACTIVATION is not null and fo.DATE_ACTIVATION <> to_date('1.1.3000'))                       then 'Schválená'
            when cas.name_application_status = 'Schválená' and (cls.name_completion_subject_status = 'Ke kompletaci') then 'Dosud finálnì nerozhodnuto'
            when cas.name_application_status = 'Schválená' and (cls.name_completion_subject_status is null)           then 'Dosud finálnì nerozhodnuto'
            when cas.name_application_status = 'Schválená' and (cls.name_completion_subject_status = 'Zkompletovaný') and os.NAME_OVERDRAFT_STATUS = 'Stornovaný' 
                 then 'Stornovaná'
            when cas.name_application_status = 'Schválená' and (cls.name_completion_subject_status = 'Zkompletovaný') 
                 and (fo.date_cancel= to_date('1.1.2000') or fo.date_cancel is null or fo.date_cancel = to_date('1.1.3000')) then 'Dosud finálnì nerozhodnuto'
            when cas.name_application_status = 'Schválená' and (cls.name_completion_subject_status = 'Stornovaný')    then 'Stornovaná'
            when cas.name_application_status = 'Schválená' and (cls.name_completion_subject_status = 'Zamítnutý')     then 'Zamítnutá'
            when cas.name_application_status not in ('Schválená', 'Stornovaná', 'Zamítnutá')                          then 'Dosud finálnì nerozhodnuto' 
            else cas.name_application_status
        end as vysledek_zadosti
     , decode(bp.signature_channel, 'BRANCH', 'Poboèka'
                                  , 'COURIER', 'Kurýr'
                                  , 'POST', 'Pošta'
                                  , 'IB', 'Internetové bankovnictví'
                                  , bp.signature_channel) as Kanal_podpisu
     , bp.dtime_signature as Datum_podpisu
     , bp.branch as Pobocka
     , case when bp.signature_channel in ('BRANCH','COURIER') then bp.id_person||' '||bp.jmeno||' '||bp.prijmeni 
            else null 
        end as ID_zamestnance 
     -- RISK-955
     , case when substr(chr2.req_ip_address,0,3) = '10.' or substr(chr2.req_ip_address,0,8) = '192.168.' or substr(chr2.req_ip_address,0,7) = '172.20.'
              or chr2.req_ip_address = 'unknown'
                 then null
            else chr2.req_ip_address 
        end as IP_adresa_req --IP adresa request
     , chr2.req_c_type --Kanál žádosti request
     , case when substr(chr2.fin_ip_address,0,3) = '10.' or substr(chr2.fin_ip_address,0,8) = '192.168.' or substr(chr2.fin_ip_address,0,7) = '172.20.'
              or chr2.fin_ip_address = 'unknown'
                 then null
            else chr2.fin_ip_address 
        end as IP_adresa_fin --IP adresa final
     , chr2.fin_c_type --Kanál žádosti final
     
     , case when rc.name_reject_reason_client is not null and rc.name_reject_reason_client<>'XNA'     then rc.name_reject_reason_client 
            when crrc.name_reject_reason_client is not null and crrc.name_reject_reason_client<>'XNA' then crrc.name_reject_reason_client 
            else null 
        end as duvod_zamitnuti
     , cs.dtime_completion_finished as compl_fi_1
     , cls.name_completion_subject_status as stav_kompl
     , case when crrc.name_reject_reason_client = 'XNA' then null 
            else crrc.name_reject_reason_client 
        end as duvod_zamitnuti_client_kompl

     , case when fo.DATE_OVERDRAFT_UTILIZATION = to_date('1.1.3000') or fo.DATE_OVERDRAFT_UTILIZATION = to_date('1.1.2000') then null 
            else fo.DATE_OVERDRAFT_UTILIZATION 
        end as Datum_nacerpani
     , case when fo.DATE_OVERDRAFT_CLOSE = to_date('1.1.3000') or fo.DATE_OVERDRAFT_CLOSE = to_date('1.1.2000') then null 
            else fo.DATE_OVERDRAFT_CLOSE 
        end as datum_ukonceni
     , case when fo.date_cancel = to_date('1.1.3000') or fo.date_cancel = to_date('1.1.2000') then null 
            else fo.date_cancel 
        end as datum_storna
     , os.NAME_OVERDRAFT_STATUS -- Pro automatickou odpovìï + øádky níže
     , os.NAME_OVERDRAFT_WFSTATUS 
     , ba.TEXT_BANK_ACCOUNT_NUMBER||'/3030' as Cislo_Uctu 
     , fo.CNT_DAYS_PAST_DUE_TOL_K_MAX 
     
     , fo.AMT_OVERDRAFT as KTK_nacerpano
     , decode(fo.FLAG_REPAYMENT_CONDITION_MET,'Y','Ano','N','Ne',null) as podm_vynoreni
     , decode(fo.FLAG_CREDIT_CONDITION_MET,'Y','Ano','N','Ne',null) as podm_kredit_minul_mesic
     , decode(fo.FLAG_TURNOVER_CURRENT,'Y','Ano','N','Ne',null) as podm_kredit_akt_mesic
     , decode(fo.FLAG_SUSPENSION,'Y','Ano','N','Ne',null) as KTK_blokace
     , cosr.name_overdraft_susp_reason as duvod_blokace
     , fo.AMT_CREDIT_TURNOVER_REQUIRED as podm_kredit_pozadovano
     , fo.AMT_CREDIT_TURNOVER_CURRENT as podm_kredit_splneno
     , fo.AMT_OVERDRAFT_INTEREST_TO_PAY as urok_k_zaplaceni

  from app_coll_jezevcik.sonik_data_calc_base sp -- Zmìnit na platné klienty
  
  join owner_dwh.dc_person dp
    on sp.skp_person = dp.SKP_PERSON
   and dp.FLAG_DELETED <> 'Y'
  join owner_dwh.f_application_at fa 
    on dp.skp_person = fa.skp_person
   and fa.flag_deleted <> 'Y'
  join OWNER_DWH.cl_application_type cat 
    on fa.skp_application_type = cat.skp_application_type
   and cat.CODE_APPLICATION_TYPE in ('OVERDRAFT')
  join owner_dwh.cl_application_status cas 
    on fa.skp_application_status = cas.skp_application_status
   and cas.name_application_status != 'DEMO' 
  left join ap_risk.js_base_final bf 
    on bf.skp_application = fa.skp_application
    
  left join owner_dwh.f_overdraft_at fo
    on fo.SKP_APPLICATION = fa.SKP_APPLICATION
   and fo.FLAG_DELETED <> 'Y'
  left join owner_dwh.cl_overdraft_status os
    on os.SKP_OVERDRAFT_STATUS = fo.SKP_OVERDRAFT_STATUS  
  left join owner_dwh.f_overdraft_suspension_tt fost
    on fost.SKP_OVERDRAFT = fo.SKP_OVERDRAFT
   and fost.FLAG_DELETED <> 'Y'
   and fost.FLAG_VALID <> 'N'
   and fost.DTIME_VALID_TO > sysdate
  left join owner_dwh.cl_overdraft_susp_reason cosr
    on cosr.skp_overdraft_susp_reason = fost.SKP_OVERDRAFT_SUSP_REASON

  left join owner_dwh.f_applicant_financial_data_tt ffd 
    on ffd.skp_applicant = fa.skp_applicant 
   and ffd.flag_current <> 'N'
   and ffd.flag_deleted <> 'Y'
  left join owner_dwh.f_bank_account_ad ba
    on ba.SKP_BANK_ACCOUNT = fo.SKP_BANK_ACCOUNT
   and ba.FLAG_DELETED <> 'Y'
  

 --Hlavni prijem
  left join OWNER_DWH.f_applicant_employment_tt fae 
    on fa.skp_applicant = fae.skp_applicant 
   and fae.flag_current <> 'N'
   and fae.FLAG_DELETED <> 'Y' 
   and fae.flag_main_income = 'Y' 
   and fae.flag_removed_wo_confirmation = 'N'
  left join OWNER_DWH.cl_economical_status ces 
    on fae.skp_economical_status = ces.skp_economical_status
  left join OWNER_DWH.cl_employment_type cet 
    on fae.skp_employment_type = cet.skp_employment_type
    
  --Vedlejsi prijem
  left join OWNER_DWH.f_applicant_employment_tt fav 
    on fa.skp_applicant = fav.skp_applicant 
   and fav.flag_current <> 'N'
   and fav.FLAG_DELETED <> 'Y' 
   and fav.flag_main_income = 'N' 
   and fav.flag_removed_wo_confirmation = 'N'
  left join OWNER_DWH.cl_economical_status cev 
    on fav.skp_economical_status = cev.skp_economical_status
  left join OWNER_DWH.cl_employment_type cetv 
    on fav.skp_employment_type = cetv.skp_employment_type
  left join app_coll_jezevcik.sonik_approve_channel chr2 
    on fa.skp_application = chr2.skp_application --APPROVE_CHANNEL
  left join OWNER_DWH.cl_reject_reason_client rc 
    on rc.skp_reject_reason_client = fa.skp_reject_reason_client
  left join ap_risk.bk_podpisy bp
    on bp.cashloanidrequest = fa.text_application_number
   and bp.code_completion_subject_type in ('CREATE_OVERDRAFT')
  left join OWNER_DWH.f_completion_subject_at cs  
    on cs.skp_application = fa.skp_application
   and cs.code_completion_subject_type in ('CREATE_OVERDRAFT')
   and cs.FLAG_DELETED <> 'Y'
  left join OWNER_DWH.cl_completion_subject_status cls 
    on cls.skp_completion_subject_status = cs.skp_completion_subject_status
  left join OWNER_DWH.cl_reject_reason_client crrc 
    on crrc.code_reject_reason_client = cs.code_reject_reason_for_client
;


comment on table app_coll_jezevcik.sonik_apl_data_ktk is 'Datová tabule pro Soníka - Žádosti o kontokorent.'; 
