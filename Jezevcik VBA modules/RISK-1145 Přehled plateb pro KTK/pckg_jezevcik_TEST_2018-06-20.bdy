create or replace package body pckg_jezevcik is

  db_prostredi VARCHAR2(20);
  
----------------------------------------------------------------------------------------------------
/* Proc doùahuje do Jez zesplatnÏn˝ch klientov; PlnÌ sa tabulka sr_coll_019_zes */
----------------------------------------------------------------------------------------------------
procedure ZESPLATNENI IS

begin
  
  -- globalni promenne pro cely package
  -- PROD/PREPROD
  SELECT sys_context('USERENV','DB_NAME') INTO db_prostredi FROM dual;
  
  
  
--merime cas, jak dlouho bezi cely chain  
insert into chain_tracking
values
  ('CHAIN_D1',
   sysdate,
   null,
   null,
   null,
   case when to_char(sysdate, 'd') in (6, 7) then 'N' else 'Y' end);


-------------------------------------------------------------------------------------------------------------------------------------------------------
execute immediate('truncate table APP_COLL_JEZEVCIK_TEST.coll_019_zes drop storage');
execute immediate('truncate table APP_COLL_JEZEVCIK_TEST.overdraft_zes drop storage');
-------------------------------------------------------------------------------------------------------------------------------------------------------
insert into APP_COLL_JEZEVCIK_TEST.coll_019_zes
select /*+ parallel(d 4)  no_index(d) */
a.skp_loan , b.skp_person, b.id_cuid , a.code_loan_number , b.name_first , b.name_last , a.cnt_days_past_due_tol_j Actual_dpd ,
 dd.date_payoff, /*d.date_past_due_tol_j , d.date_past_due_tol_j_paid  ,*/ trunc(max(a.date_payoff - d.date_past_due_tol_j)) dpd_payoff ,
 --case when min(a.date_payoff) = '1/1/3000' then null else trunc(max(a.date_payoff - d.date_past_due_tol_j)) end as dpd_payoff,
   a.date_loan_utilization ,a.amt_payoff , p.name_payoff_reason/*, c.name_loan_status , c.name_loan_wfstatus */ 
from owner_dwh.f_loan_at a
     join (select skp_person, max(date_payoff) as date_payoff --zjistime datum zesplatneni. pokud je uver zesplatnen z duvodu umrti, bere se planovane, jinak realizovane
            from (select l.skp_person, --v tabulce f_loan_at je planovane, proto skutecne beru z tabulky f_loan_status_tt
                         case
                           when cp.NAME_PAYOFF_REASON = '⁄mrtÌ' or cp.NAME_PAYOFF_REASON = 'Insolvence' then
                            trunc(l.date_payoff)
                           else
                            trunc(ls.dtime_valid_from)
                         end as date_payoff
                    from owner_dwh.f_loan_status_tt ls
                    join owner_dwh.cl_loan_status c
                      on c.SKP_LOAN_STATUS = ls.SKP_LOAN_STATUS
                     and c.NAME_LOAN_WFSTATUS = 'ZesplatnÏn˝'
                    join owner_dwh.f_loan_at l
                      on l.skp_loan = ls.skp_loan
                    join owner_dwh.cl_payoff_reason cp
                      on cp.SKP_PAYOFF_REASON = l.SKP_PAYOFF_REASON
                   where ls.flag_deleted = 'N')
           group by skp_person
     ) dd on a.skp_person = dd.skp_person
     join owner_dwh.dc_person b on a.skp_person = b.skp_person
     join owner_dwh.cl_loan_status c on a.skp_loan_status = c.skp_loan_status
     left join owner_dwh.cl_payoff_reason p on a.skp_payoff_reason = p.skp_payoff_reason
     left join owner_dwh.f_instalment_head_ad d on a.skp_loan = d.skp_loan            
     -- left join APP_COLL_JEZEVCIK_TEST.instalment_head d on a.skp_loan = d.skp_loan          -- RISK-723 instalment_head sa plni az v chain_step_4, tato procka je step_1
          and a.date_payoff <= d.date_past_due_tol_j_paid  and a.date_payoff <> '1/1/3000'--between d.date_past_due_tol_j and d.date_past_due_tol_j_paid
     
where ((c.code_loan_status = '1' --AktivnÌ  
      and c.code_loan_wfstatus = '7') ) --ZesplatnÏn˝
group by a.skp_loan, b.skp_person, b.id_cuid , a.code_loan_number , b.name_first , b.name_last , a.cnt_days_past_due_tol_j ,
 dd.date_payoff, /*d.date_past_due_tol_j , d.date_past_due_tol_j_paid  ,  dpd_payoff */
   a.date_loan_utilization ,a.amt_payoff , p.name_payoff_reason /*, c.name_loan_status , c.name_loan_wfstatus */
--optimalizace 19.1.2016 order by dpd_payoff , dd.date_payoff
;


begin  dbms_stats.gather_table_stats(ownname => nazev_aktualniho_hriste,tabname => 'COLL_019_ZES',
   estimate_percent => dbms_stats.auto_sample_size,  cascade => true, no_invalidate => false); end;
   

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
    WHERE os.NAME_OVERDRAFT_STATUS = 'AktivnÌ'
      -- !!!!!!!!!!!!!!
      AND os.NAME_OVERDRAFT_WFSTATUS = 'ZesplatnÏn˝' -- RISK-1145  odkomentovat na PROD !!!!!!!!!
      -- !!!!!!!!!!!!!!
;

----------------------------------------------------------------------------------------------------------------------

/*
insert into coll_019_dluhy

with amts as (

select a.*, b.dtime_valid_from, b.dtime_valid_to , b.dtime_inserted,b.date_payoff as d_payoff, b.amt_payoff as a_payoff,
       b.amt_payoff_principal , b.amt_payoff_interest,
       b.amt_payoff_penalty , b.amt_penalty_interest_overdue , b.amt_penalty_interest_current, b.amt_penalty_overdue
from coll_019_zes a
     left join owner_dwh.f_loan_balance_at b on a.skp_loan = b.skp_loan
          and b.amt_payoff is not null and  trunc(dtime_valid_from) <= trunc(dtime_valid_to) -- pretoze su v tabulke divne hodnoty, ktore splnaju opak
                                                                                             -- a s Martinom sme sa zhodli, ze ich mame vynechat
order by a.skp_person , a.skp_loan, b.dtime_valid_from , b.dtime_inserted   )

select a.* , rownum as riadok
from amts a
;



begin  dbms_stats.gather_table_stats(ownname => nazev_aktualniho_hriste,tabname => 'COLL_019_DLUHY',
   estimate_percent => dbms_stats.auto_sample_size,  cascade => true, no_invalidate => false); end;*/

----------------------------------------------------------------------------------------------------------------------
/*

--hodnota dlhu k datumu zosplatnenia
insert into coll_019_dluh_zes

with riadky as (

select  min(riadok) as riadok
from coll_019_dluhy a
group by  SKP_LOAN, SKP_PERSON, ID_CUID, CODE_LOAN_NUMBER, NAME_FIRST, NAME_LAST, ACTUAL_DPD, DATE_PAYOFF, DPD_PAYOFF, DATE_LOAN_UTILIZATION, AMT_PAYOFF
order by a.skp_person , a.skp_loan
)
select a.*
from coll_019_dluhy a
     join riadky b on a.riadok = b.riadok
;
*/

----------------------------------------------------------------------------------------------------------------------

/*
--hodnoty k dnesnimu dnu
insert into coll_019_dluh_dnes

with dnes_riadky as (

select   max(riadok) as riadok
from coll_019_dluhy a
group by  SKP_LOAN, SKP_PERSON, ID_CUID, CODE_LOAN_NUMBER, NAME_FIRST, NAME_LAST, ACTUAL_DPD, DATE_PAYOFF, DPD_PAYOFF, DATE_LOAN_UTILIZATION, AMT_PAYOFF
order by a.skp_person , a.skp_loan
)
select a.*
from coll_019_dluhy a
     join dnes_riadky b on a.riadok = b.riadok
;

*/

----------------------------------------------------------------------------------------------------------------------
-- insolvence
/*edit 13-10-20166 - insolvence jiz nebudeme napocitavat zde, ale pouze v insol_jez, je to zde duplicitne

insert into APP_COLL_JEZEVCIK_TEST.coll_019_ins1
select aaa.id_cuid, a.skp_person, a.skp_loan, a.code_loan_number, b.name_full || ' ' || b.name_first as SR_jmeno,
       b.dtime_valid_from as SR_dtime_valid_from, e.code_insol_subject_role_type as IS_role_subj_v_ins_rizeni,
       g.code_insolvency_case_status as IC_stav_rizeni, g.dtime_publication as IC_datum_zalozeni_spisu, --g.dtime_cancellation as IC_date_cancellation,
       g.text_case_ident_number as IC_spisova_znacka,--, h.dtime_publication, h.name_insolvency_case_item_type
       h.dtime_publication as datum_zmeny_stavu
from owner_dwh.f_loan_at a
     join owner_dwh.cl_loan_status aa on a.skp_loan_status = aa.skp_loan_status
     join owner_dwh.dc_person aaa on a.skp_person = aaa.skp_person
     join owner_dwh.f_insol_subj_pers_relation_tt b on a.skp_person = b.skp_person
                                                     and b.flag_current = 'Y'
     join owner_dwh.cl_insolv_reg_match_type c on c.skp_insolv_reg_match_type = b.skp_insolv_reg_match_type
                                          and c.code_insolv_reg_match_type in  ('A','C','D') --vyloucime firmy
     left join owner_dwh.f_insolvency_subject_at d on d.skf_insolvency_subject = b.skf_insolvency_subject
     left join owner_dwh.cl_insol_subject_role_type e on e.skp_insol_subject_role_type = d.skp_ins_subject_role_type
     left join owner_dwh.cl_insolvency_subject_type f on d.skp_insolvency_subject_type = f.skp_insolvency_subject_type
     left join owner_dwh.dc_insolvency_case g on g.skp_insolvency_case = d.skp_insolvency_case
     left join owner_dwh.f_insolvency_case_status_tt h on g.skp_insolvency_case = h.skp_insolvency_case
                                                       and h.code_insolvency_case_status = g.code_insolvency_case_status
where aa.code_loan_status = '1'
      and e.code_insol_subject_role_type = 'DLUéNÕK'
;
----------------------------------------------------------------------------------------------------------------------
-- insolvence_beta
insert into APP_COLL_JEZEVCIK_TEST.coll_019_ins_beta
select aaa.id_cuid, a.skp_person, a.skp_loan, a.code_loan_number, b.name_last || ' ' || b.name_first as PERSON_jmeno, b.text_birth_number,
       e.code_insol_subject_role_type as IS_role_subj_v_ins_rizeni,
       g.code_insolvency_case_status as IC_stav_rizeni, g.dtime_publication as IC_datum_zalozeni_spisu, --g.dtime_cancellation as IC_date_cancellation,
       g.text_case_ident_number as IC_spisova_znacka,
       h.dtime_publication as datum_zmeny_stavu
from owner_dwh.f_loan_at a
     join owner_dwh.cl_loan_status aa on a.skp_loan_status = aa.skp_loan_status
     join owner_dwh.dc_person aaa on a.skp_person = aaa.skp_person
     join owner_dwh.dc_person b on a.skp_person = b.skp_person
     left join owner_dwh.f_insolvency_subject_at d on b.name_first = d.name_first
                                                       and b.name_last = d.name_full
                                                       and b.date_birth = d.date_birth
     left join owner_dwh.cl_insol_subject_role_type e on e.skp_insol_subject_role_type = d.skp_ins_subject_role_type
     left join owner_dwh.cl_insolvency_subject_type f on d.skp_insolvency_subject_type = f.skp_insolvency_subject_type
     left join owner_dwh.dc_insolvency_case g on g.skp_insolvency_case = d.skp_insolvency_case
     left join owner_dwh.f_insolvency_case_status_tt h on g.skp_insolvency_case = h.skp_insolvency_case
                                                       and h.code_insolvency_case_status = g.code_insolvency_case_status
where aa.code_loan_status = '1'
      and e.code_insol_subject_role_type = 'DLUéNÕK'
      and a.skp_loan not in (select skp_loan from APP_COLL_JEZEVCIK_TEST.coll_019_ins1)
;
---------------------------------------------------------------------------------------------------------------------
-- insolvence na osobu
insert into APP_COLL_JEZEVCIK_TEST.coll_019_ins2
select distinct skp_person  , id_cuid , sr_jmeno , first_value(ic_stav_rizeni) over (partition by skp_person order by ic_datum_zalozeni_spisu desc) as insolvency_case_status
from (
select skp_person  , id_cuid, sr_jmeno , ic_stav_rizeni , ic_datum_zalozeni_spisu
from APP_COLL_JEZEVCIK_TEST.coll_019_ins1
union
select skp_person , id_cuid , person_jmeno as sr_jmeno , ic_stav_rizeni , ic_datum_zalozeni_spisu
from APP_COLL_JEZEVCIK_TEST.coll_019_ins_beta
)
;
*/
---------------------------------------------------------------------------------------------------------------------
execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.sr_coll_019_zes drop storage');
insert into APP_COLL_JEZEVCIK_TEST.sr_coll_019_zes (SKP_PERSON,
                                                    SKP_LOAN,
                                                    ID_CUID,
                                                    CODE_LOAN_NUMBER,
                                                    NAME_LAST,
                                                    NAME_FIRST,
                                                    DATUM_UMRTI,
                                                    ACTUAL_DPD,
                                                    DATE_PAYOFF,
                                                    DPD_PAYOFF,
                                                    DATE_LOAN_UTILIZATION,
                                                    DATE_PAYOFF_MAX,
                                                    NAME_PAYOFF_REASON)
select a.SKP_PERSON , a.skp_loan , a.id_cuid , a.code_loan_number , a.name_last , a.name_first,
       /*case when c.skp_person is not null then 'ano' else 'ne' end as insolvence,*/
       case when d.date_death = to_date('1.1.3000') then null else d.date_death end as datum_umrti,
       a.actual_dpd ,
       a.date_payoff,
       a.dpd_payoff ,
       a.date_loan_utilization ,  --a.a_payoff as amt_payoff , a.amt_payoff_principal , a.amt_payoff_interest , a.amt_penalty_overdue,
       --b.a_payoff as dluh_dnes , b.amt_payoff_principal as jistina_po_splatnosti ,
       --b.amt_payoff_interest as urok_po_splatnosti, b.amt_penalty_overdue as sankce_po_splatnosti ,
       max(a.date_payoff) over (partition by a.SKP_PERSON ) as date_payoff_max ,
       a.name_payoff_reason  /*c.insolvency_case_status ,*/ 
from APP_COLL_JEZEVCIK_TEST.coll_019_zes a
     --join APP_COLL_JEZEVCIK_TEST.coll_019_dluh_dnes b on a.skp_loan = b.skp_loan
     join owner_dwh.dc_person d on a.skp_person = d.skp_person
     --left join APP_COLL_JEZEVCIK_TEST.coll_019_ins2 c on a.skp_person = c.skp_person
;


-- zespl info na klienta
-- RISK-KTK
-- Tabulka se vyuûÌv· pouze v collections reportech - p¯esunuto na AP_COLL - zakomentov·no 15.3.2018 - 
/*execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.zes_person drop storage');
insert into APP_COLL_JEZEVCIK_TEST.zes_person
select id_cuid , name_last, name_first , max(max_actual_dpd) as max_actual_dpd ,
       min(date_payoff) as min_date_payoff , listagg(code_loan_number, '; ')  within group (order by code_loan_number) code_loan_number ,
       listagg(name_payoff_reason, ', ') within group (order by name_payoff_reason) name_payoff_reason ,
       to_char(min(date_payoff), 'yyyy-mm') as mesic_zes,
       listagg(taskname, ', ') within group (order by taskname) taskname -- RISK-490 P¯id·nÌ "podd˘vod˘" zesplatnÏnÌ z LCS v p¯ÌpadÏ delikvence
       from
(
select  b.skp_person, b.id_cuid , listagg(a.code_loan_number, '; ')  within group (order by a.code_loan_number) code_loan_number  , b.name_first , b.name_last ,
        max(a.cnt_days_past_due_tol_j) max_actual_dpd ,
        min(a.date_payoff) as date_payoff,
   --a.date_loan_utilization ,a.amt_payoff ,
   p.name_payoff_reason
   , nvl(ltt.taskname,'AutoTerminationDPD') taskname
from owner_dwh.f_loan_at a
     join owner_dwh.dc_person b on a.skp_person = b.skp_person
     join owner_dwh.cl_loan_status c on a.skp_loan_status = c.skp_loan_status
     left join owner_dwh.cl_payoff_reason p on a.skp_payoff_reason = p.skp_payoff_reason
     left join owner_int.v_lcs_deal ld on ld.contractnr = a.CODE_LOAN_NUMBER
     left join owner_int.v_lcs_classifiable lc on lc.t_deal_1 = ld.t_deal_id
                                         and lc.ctype = 'DEAL'
     left join owner_int.v_lcs_taskcalendar ltc on ltc.t_classifiable_1 = lc.t_classifiable_id 
      and ltc.deleted = 0 and ltc.executed = 1 and ltc.isvalid = 1
      and ltc.t_tasktype_1 in (11,14,15,16,17,19,20,21,610)
     left join owner_int.v_lcs_tasktype ltt on ltt.t_tasktype_id = ltc.t_tasktype_1
where c.code_loan_status = '1'
      and c.code_loan_wfstatus = '7'
group by b.skp_person, b.id_cuid , b.name_first , b.name_last ,p.name_payoff_reason, ltt.taskname
)
group by skp_person , id_cuid , name_last, name_first
--optimalizace 19.1.2016 order by min_date_payoff desc
;*/
commit;

-------------------------------------------------------------------------------------------------------------------------------------------------------
execute immediate('truncate table APP_COLL_JEZEVCIK_TEST.coll_019_zes drop storage');
-------------------------------------------------------------------------------------------------------------------------------------------------------




commit;

END;


----------------------------------------------------------------------------------------------------
/* Proc doùahuje do Jez zhody insolvencÌ z ISIRu; PlnÌ sa tabulky insol_list (kopie DWH), ins_client_list (uchov·v· uûÌvatelskÈ vstupy) */
----------------------------------------------------------------------------------------------------

procedure INSOL IS
  
begin


/*
--- >>>  logika insolvenci k 23.2.2018  <<< ---

existuju 2 hlavne tabulky: 
 1. insol_list      = vymaze a naplni sa podla DWH kazdy den
 2. ins_client_list = Jezevcik si uchovava hodnoty. Plni sa z insol_list, alebo uzivatel potvrdi zhodu z Jezevcika.

u klientov ktori sa dostanu do view v_ins_unpaired, je na uzivatelovi aby rozhodol ci je to:
    1. zhoda   -> vlozi sa do ins_client_list
    2. nezhoda -> vlozi sa do ins_unconfirmed

--- >>>  logika plnenia tabuliek v procedure INSOL k 23.2.2018  <<< ---

plnenie TABULKA  |               FROM                         |                pozn.
insol            | owner_dwh                                  | naparovani klienti pres DWH
insol2           | insol union owner_dwh                      | insol + naparovani pres borgskou kostku (ctverec v kodu)
insol_loan       | insol2 join owner_dwh                      | 'ûadatel o hypotÈku', 'ûadatel o ˙vÏr    '
insol_overdraft  | insol2 join owvner_dwh                     | 'ûadatel o kontokorent'
insol_codeb      |(insol2 join owner_dwh) not in (insol_loan) | 'spoludluûnÌk'
insol_collateral |(insol2 join owner_dwh) not in (loan,codeb) | 'z·stavce'
insol_debet      |(insol2 join owner_dwh) not in (..+collat.) | 'debet·¯'

insol3           | UNION(loan,codeb,collateral,debet)         | iba select * ..
insol4           | insol3 union (insol3 join owner_dwh)       | ?? tomuto fakt nerozumiem..  owner_dwh: f_bank_account_ad, cl_bank_account_status
ins_contracts    | insol4 left join owner_dwh                 | ?? iba 5 stlpcov, mozno redundantne zbytocne vsetko
insol_list       | distinct insol4                            | 12 z 15 stlpcov

-----PROCEDURA UPDATE_INSOL_LIST
INS_CLIENT_LIST  | insert new from INSOL_LIST                 | daju sa pridavat aj manualne z Jezevcika
*/






execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.insol  drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.insol2  drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.insol_loan  drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.insol_overdraft  drop storage'; --RISK-1119
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.insol_codeb  drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.insol_collateral  drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.insol_debet  drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.insol3  drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.insol4  drop storage';


insert into APP_COLL_JEZEVCIK_TEST.insol
select  /*+ parallel(b 4)  */
        aaa.id_cuid, aaa.skp_person , --a.skp_person, a.skp_loan, a.code_loan_number,
       b.name_full || ' ' || b.name_first as SR_jmeno, 
       b.dtime_valid_from as SR_dtime_valid_from, e.code_insol_subject_role_type as IS_role_subj_v_ins_rizeni,
       g.code_insolvency_case_status as IC_stav_rizeni, g.dtime_publication as IC_datum_zalozeni_spisu, --g.dtime_cancellation as IC_date_cancellation,
       g.text_case_ident_number as IC_spisova_znacka,--, h.dtime_publication, h.name_insolvency_case_item_type
       h.dtime_publication as datum_zmeny_stavu , 'Y' flag_naparovana
from /*clients a
     join*/ owner_dwh.dc_person aaa --on a.skp_person = aaa.skp_person
     join owner_dwh.f_insol_subj_pers_relation_tt b on aaa.skp_person = b.skp_person
                                                     and b.flag_current = 'Y'
     join owner_dwh.cl_insolv_reg_match_type c on c.skp_insolv_reg_match_type = b.skp_insolv_reg_match_type
                                          and c.code_insolv_reg_match_type in  ('A','C','D') --vyloucime firmy
     left join owner_dwh.f_insolvency_subject_at d on d.skf_insolvency_subject = b.skf_insolvency_subject
     left join owner_dwh.cl_insol_subject_role_type e on e.skp_insol_subject_role_type = d.skp_ins_subject_role_type
     left join owner_dwh.cl_insolvency_subject_type f on d.skp_insolvency_subject_type = f.skp_insolvency_subject_type
     left join owner_dwh.dc_insolvency_case g on g.skp_insolvency_case = d.skp_insolvency_case
     left join owner_dwh.f_insolvency_case_status_tt h on g.skp_insolvency_case = h.skp_insolvency_case
                                                       and h.code_insolvency_case_status = g.code_insolvency_case_status
where  e.code_insol_subject_role_type = 'DLUéNÕK' and g.FLAG_DELETED = 'N'
       AND d.dtime_subject_cancelled = to_date('1.1.1000') -- FA pridal 29.8.2017
      --and g.code_insolvency_case_status = 'ODDLUéENÕ'
;
/*
insert into APP_COLL_JEZEVCIK_TEST.meric
select sysdate, 'insol_jez: 2' from dual;
*/

insert into APP_COLL_JEZEVCIK_TEST.insol2
select * from
APP_COLL_JEZEVCIK_TEST.insol
union
select b.id_cuid, b.skp_person, b.name_last || ' ' || b.name_first as SR_jmeno,  null ,
       e.code_insol_subject_role_type as IS_role_subj_v_ins_rizeni,
       g.code_insolvency_case_status as IC_stav_rizeni, g.dtime_publication as IC_datum_zalozeni_spisu, --g.dtime_cancellation as IC_date_cancellation,
       g.text_case_ident_number as IC_spisova_znacka,
       h.dtime_publication as datum_zmeny_stavu , 'N' as flag_naparovana
from owner_dwh.dc_person b
     left join owner_dwh.f_insolvency_subject_at d on                                               replace(
                                                       replace(replace(replace(
                                                       replace(utl_raw.cast_to_varchar2(nlssort(lower(b.name_first), 'nls_sort=binary_ai')),''''),'°'),'¥'),'y','i')
                                                       ,' ')
                                                       =
                                                       replace(
                                                       replace(replace(replace(
                                                       replace(utl_raw.cast_to_varchar2(nlssort(lower(d.name_first), 'nls_sort=binary_ai')),''''),'°'),'¥'),'y','i')
                                                       ,' ')

                                                       and

                                                       replace(
                                                       replace(replace(replace(
                                                       replace(utl_raw.cast_to_varchar2(nlssort(lower(b.name_last), 'nls_sort=binary_ai')),''''),'°'),'¥'),'y','i')
                                                       ,' ')
                                                       =
                                                       replace(
                                                       replace(replace(replace(
                                                       replace(utl_raw.cast_to_varchar2(nlssort(lower(d.name_full), 'nls_sort=binary_ai')),''''),'°'),'¥'),'y','i')
                                                       ,' ')

                                                       and
                                                       b.date_birth = d.date_birth
     left join owner_dwh.cl_insol_subject_role_type e on e.skp_insol_subject_role_type = d.skp_ins_subject_role_type
     left join owner_dwh.cl_insolvency_subject_type f on d.skp_insolvency_subject_type = f.skp_insolvency_subject_type
     left join owner_dwh.dc_insolvency_case g on g.skp_insolvency_case = d.skp_insolvency_case
     left join owner_dwh.f_insolvency_case_status_tt h on g.skp_insolvency_case = h.skp_insolvency_case
                                                      and h.code_insolvency_case_status = g.code_insolvency_case_status
where e.code_insol_subject_role_type = 'DLUéNÕK'
      and b.skp_person not in (select skp_person from APP_COLL_JEZEVCIK_TEST.insol) and g.FLAG_DELETED = 'N'
      AND d.dtime_subject_cancelled = to_date('1.1.1000') -- FA pridal 29.8.2017
;
begin  dbms_stats.gather_table_stats(ownname => nazev_aktualniho_hriste,tabname => 'insol2',
   estimate_percent => dbms_stats.auto_sample_size,  cascade => true, no_invalidate => false); end;
/*insert into APP_COLL_JEZEVCIK_TEST.meric
select sysdate, 'insol_jez: 3' from dual;
*/
insert into APP_COLL_JEZEVCIK_TEST.insol_loan (ID_CUID,
                                               SKP_PERSON,
                                               SR_JMENO,
                                               SR_DTIME_VALID_FROM,
                                               IS_ROLE_SUBJ_V_INS_RIZENI,
                                               IC_STAV_RIZENI,
                                               IC_DATUM_ZALOZENI_SPISU,
                                               IC_SPISOVA_ZNACKA,
                                               DATUM_ZMENY_STAVU,
                                               FLAG_NAPAROVANA,
                                               VZTAH,
                                               CONTRACTNR,
                                               NAME_LOAN_STATUS,
                                               NAME_LOAN_WFSTATUS,
                                               DATE_PAYOFF)
select t.id_cuid,
       t.skp_person,
       t.sr_jmeno,
       t.sr_dtime_valid_from,
       t.is_role_subj_v_ins_rizeni,
       t.ic_stav_rizeni,
       t.ic_datum_zalozeni_spisu,
       t.ic_spisova_znacka,
       t.datum_zmeny_stavu,
       t.flag_naparovana ,

       case when aty.CODE_APPLICATION_TYPE in ('MORTGAGE_REF') then 'ûadatel o hypotÈku' else 'ûadatel o ˙vÏr   ' end as vztah, --skp_product u hypoteky pouze refinanc, pokud se jendou budou nabizet hypoteky primo, je potreba to zmenit
       l.code_loan_number as contractnr,
       ls.name_loan_status , ls.name_loan_wfstatus ,
       l.date_payoff
from
APP_COLL_JEZEVCIK_TEST.insol2 t

join owner_dwh.f_loan_at l on t.skp_person = l.skp_person
join owner_dwh.cl_loan_status ls on l.skp_loan_status = ls.skp_loan_status
join owner_dwh.dc_application a on a.SKP_APPLICATION = l.SKP_APPLICATION
join owner_dwh.cl_application_type aty on aty.SKP_APPLICATION_TYPE = a.SKP_APPLICATION_TYPE -- RISK-554
where ls.code_loan_status = '1'

;


-- RISK-1119
INSERT INTO APP_COLL_JEZEVCIK_TEST.INSOL_OVERDRAFT (ID_CUID,
                                                    SKP_PERSON,
                                                    SR_JMENO,
                                                    SR_DTIME_VALID_FROM,
                                                    IS_ROLE_SUBJ_V_INS_RIZENI,
                                                    IC_STAV_RIZENI,
                                                    IC_DATUM_ZALOZENI_SPISU,
                                                    IC_SPISOVA_ZNACKA,
                                                    DATUM_ZMENY_STAVU,
                                                    FLAG_NAPAROVANA,
                                                    VZTAH,
                                                    CONTRACTNR,
                                                    STATUS,
                                                    WFSTATUS,
                                                    DATE_PAYOFF)
SELECT t.ID_CUID,
       t.SKP_PERSON,
       t.SR_JMENO,
       t.SR_DTIME_VALID_FROM,
       t.IS_ROLE_SUBJ_V_INS_RIZENI,
       t.IC_STAV_RIZENI,
       t.IC_DATUM_ZALOZENI_SPISU,
       t.IC_SPISOVA_ZNACKA,
       t.DATUM_ZMENY_STAVU,
       t.FLAG_NAPAROVANA,
       'ûadatel o kontokorent' AS vztah,
       ov.CODE_OVERDRAFT_NUMBER,
       os.NAME_OVERDRAFT_STATUS,
       os.NAME_OVERDRAFT_WFSTATUS,
       ov.DATE_PAYOFF
FROM APP_COLL_JEZEVCIK_TEST.insol2 t
JOIN owner_dwh.F_OVERDRAFT_AT ov ON t.SKP_PERSON = ov.SKP_PERSON
JOIN owner_dwh.CL_OVERDRAFT_STATUS os ON ov.SKP_OVERDRAFT_STATUS = os.SKP_OVERDRAFT_STATUS
WHERE os.NAME_OVERDRAFT_STATUS = 'AktivnÌ'
  AND t.skp_person not in (select skp_person from APP_COLL_JEZEVCIK_TEST.insol_loan)
;




/*
insert into APP_COLL_JEZEVCIK_TEST.meric
select sysdate, 'insol_jez: 4' from dual;
*/
insert into APP_COLL_JEZEVCIK_TEST.insol_codeb (ID_CUID,
                                                SKP_PERSON,
                                                SR_JMENO,
                                                SR_DTIME_VALID_FROM,
                                                IS_ROLE_SUBJ_V_INS_RIZENI,
                                                IC_STAV_RIZENI,
                                                IC_DATUM_ZALOZENI_SPISU,
                                                IC_SPISOVA_ZNACKA,
                                                DATUM_ZMENY_STAVU,
                                                FLAG_NAPAROVANA,
                                                VZTAH,
                                                CONTRACTNR,
                                                NAME_LOAN_STATUS,
                                                NAME_LOAN_WFSTATUS,
                                                DATE_PAYOFF)
select t.id_cuid,
       t.skp_person,
       t.sr_jmeno,
       t.sr_dtime_valid_from,
       t.is_role_subj_v_ins_rizeni,
       t.ic_stav_rizeni,
       t.ic_datum_zalozeni_spisu,
       t.ic_spisova_znacka,
       t.datum_zmeny_stavu,
       t.flag_naparovana ,

       'spoludluûnÌk' as vztah,
       l.code_loan_number  as contractnr ,
       ls.name_loan_status , ls.name_loan_wfstatus ,
       l.date_payoff
from
APP_COLL_JEZEVCIK_TEST.insol2 t
join owner_dwh.f_loan_codebtor_relation_tt lcr on t.skp_person = lcr.skp_person_codebtor
join owner_dwh.f_loan_at l on lcr.skp_loan = l.skp_loan
join owner_dwh.cl_loan_status ls on l.skp_loan_status = ls.skp_loan_status
left join owner_dwh.f_bank_account_ad ba on lcr.skp_person = ba.skp_person and ba.id_bank_product = '1'
where ls.code_loan_status = '1' and lcr.FLAG_RELATION_VALID = 'Y'
      and t.skp_person not in (select skp_person from APP_COLL_JEZEVCIK_TEST.insol_loan)
      AND t.skp_person not in (select skp_person from APP_COLL_JEZEVCIK_TEST.INSOL_OVERDRAFT)  --RISK-1119
;
/*insert into APP_COLL_JEZEVCIK_TEST.meric
select sysdate, 'insol_jez: 5' from dual;
*/
insert into APP_COLL_JEZEVCIK_TEST.insol_collateral (ID_CUID,
                                                     SKP_PERSON,
                                                     SR_JMENO,
                                                     SR_DTIME_VALID_FROM,
                                                     IS_ROLE_SUBJ_V_INS_RIZENI,
                                                     IC_STAV_RIZENI,
                                                     IC_DATUM_ZALOZENI_SPISU,
                                                     IC_SPISOVA_ZNACKA,
                                                     DATUM_ZMENY_STAVU,
                                                     FLAG_NAPAROVANA,
                                                     VZTAH,
                                                     CONTRACTNR,
                                                     NAME_LOAN_STATUS,
                                                     NAME_LOAN_WFSTATUS,
                                                     DATE_PAYOFF)
    select t.id_cuid,
           t.skp_person,
           t.sr_jmeno,
           t.sr_dtime_valid_from,
           t.is_role_subj_v_ins_rizeni,
           t.ic_stav_rizeni,
           t.ic_datum_zalozeni_spisu,
           t.ic_spisova_znacka,
           t.datum_zmeny_stavu,
           t.flag_naparovana ,

           'z·stavce' as vztah,
           l.code_loan_number  as contractnr ,
           ls.name_loan_status , ls.name_loan_wfstatus ,
           l.date_payoff
    from
    APP_COLL_JEZEVCIK_TEST.insol2 t
    join owner_dwh.f_ownship_collateral_real_at oc on oc.skp_person = t.skp_person
    join owner_dwh.f_loan_collateral_tt lc on oc.skp_loan_collateral = lc.skp_loan_collateral -- RISK-866
    --JOIN (SELECT NULL AS skp_loan, NULL AS skp_loan_collateral FROM dual) lc on oc.skp_loan_collateral = lc.skp_loan_collateral -- RISK-866
    join owner_dwh.cl_ownership_status cos on oc.skp_ownership_status = cos.skp_ownership_status and cos.code_ownership_status = 'A' -- PlatnÈ -- RISK-866
    join owner_dwh.f_loan_at l on lc.skp_loan = l.skp_loan
    join owner_dwh.cl_loan_status ls on l.skp_loan_status = ls.skp_loan_status
    where ls.code_loan_status = '1'
      AND SYSDATE BETWEEN oc.date_valid_from AND oc.date_valid_to                    -- RISK-866
      AND oc.name_mortgager_role = 'Aktu·lnÌ z·stavce'                               -- RISK-866
      AND lc.flag_current = 'Y'                                                      -- RISK-866
      AND lc.FLAG_DELETED = 'N'                                                      -- RISK-866
          and t.skp_person not in (select skp_person from APP_COLL_JEZEVCIK_TEST.insol_loan)
          and t.skp_person not in (select skp_person from APP_COLL_JEZEVCIK_TEST.insol_codeb)
          AND t.skp_person not in (select skp_person from APP_COLL_JEZEVCIK_TEST.INSOL_OVERDRAFT)  --RISK-1119
;
/*insert into APP_COLL_JEZEVCIK_TEST.meric
select sysdate, 'insol_jez: 6' from dual;
*/
insert into APP_COLL_JEZEVCIK_TEST.insol_debet (ID_CUID,
                                                SKP_PERSON,
                                                SR_JMENO,
                                                SR_DTIME_VALID_FROM,
                                                IS_ROLE_SUBJ_V_INS_RIZENI,
                                                IC_STAV_RIZENI,
                                                IC_DATUM_ZALOZENI_SPISU,
                                                IC_SPISOVA_ZNACKA,
                                                DATUM_ZMENY_STAVU,
                                                FLAG_NAPAROVANA,
                                                VZTAH,
                                                CONTRACTNR,
                                                NAME_LOAN_STATUS,
                                                NAME_LOAN_WFSTATUS,
                                                DATE_PAYOFF)
with ucet_522 as ( --nove pridano, abychom evidovali vyvedeny dluh u debetaru select insolvenci na uctu 30300522, RISK-199
select a.skp_person, - r.AMT_TRANSACTION_REALIZED as amt_balance
from owner_dwh.f_required_transaction_tt r 
join owner_dwh.f_contract_at a on a.TEXT_CONTRACT_NUMBER = r.TEXT_VARIABLE_SYMBOL
where skp_bank_account = 7533428 and a.FLAG_DELETED = 'N')


select t.* ,

       'debet·¯' as vztah , null as contractnr ,
       null as name_loan_status ,
       null as name_loan_wfstatus ,
       DATE'3000-01-01' as date_payoff
from
APP_COLL_JEZEVCIK_TEST.insol2 t
left join owner_dwh.f_bank_account_ad ba on t.skp_person = ba.skp_person
left join owner_dwh.cl_bank_account_status bas on ba.skp_bank_account_status = bas.skp_bank_account_status
left join owner_dwh.f_currency_rate_pd cr on cr.skp_currency_source = ba.skp_currency and trunc(sysdate)-1 = cr.DATE_RATE_EFFECTIVE
left join ucet_522 u on u.skp_person = t.skp_person
where (ba.id_bank_product = '1' or ba.ID_BANK_PRODUCT is null)
      and (u.skp_person is not null or ba.skp_person is not null)
      --and ba.amt_balance <= -100
      and t.skp_person not in (select skp_person from APP_COLL_JEZEVCIK_TEST.insol_loan)
      and t.skp_person not in (select skp_person from APP_COLL_JEZEVCIK_TEST.insol_codeb)
      and t.skp_person not in (select skp_person from APP_COLL_JEZEVCIK_TEST.insol_collateral)
      AND t.skp_person not in (select skp_person from APP_COLL_JEZEVCIK_TEST.INSOL_OVERDRAFT) --RISK-1119
group by t.ID_CUID, t.skp_person, t.SR_JMENO, t.SR_DTIME_VALID_FROM, t.IS_ROLE_SUBJ_V_INS_RIZENI,
         t.IC_STAV_RIZENI, t.IC_DATUM_ZALOZENI_SPISU, t.IC_SPISOVA_ZNACKA, t.DATUM_ZMENY_STAVU , t.flag_naparovana
having (sum(nvl(ba.amt_balance*cr.rate_exchange_middle , ba.amt_balance)) <= -1500 or sum(nvl(u.amt_balance,0)) <= -1500);



insert into APP_COLL_JEZEVCIK_TEST.insol3
    select * FROM APP_COLL_JEZEVCIK_TEST.insol_loan
  UNION
    SELECT * FROM APP_COLL_JEZEVCIK_TEST.INSOL_OVERDRAFT --RISK-1119
  UNION
    select * FROM APP_COLL_JEZEVCIK_TEST.insol_codeb
  union
    select * FROM APP_COLL_JEZEVCIK_TEST.insol_collateral
  union
    select id_cuid, skp_person, sr_jmeno, sr_dtime_valid_from, is_role_subj_v_ins_rizeni, ic_stav_rizeni, ic_datum_zalozeni_spisu, ic_spisova_znacka, datum_zmeny_stavu, flag_naparovana, vztah,
    null as contractnr, null as name_loan_status, null as name_loan_wf_status, date_payoff
    FROM APP_COLL_JEZEVCIK_TEST.insol_debet
;
/*
insert into APP_COLL_JEZEVCIK_TEST.meric
select sysdate, 'insol_jez: 7' from dual;

*/
insert into APP_COLL_JEZEVCIK_TEST.insol4 (ID_CUID,
                                           SKP_PERSON,
                                           SR_JMENO,
                                           SR_DTIME_VALID_FROM,
                                           IS_ROLE_SUBJ_V_INS_RIZENI,
                                           IC_STAV_RIZENI,
                                           IC_DATUM_ZALOZENI_SPISU,
                                           IC_SPISOVA_ZNACKA,
                                           DATUM_ZMENY_STAVU,
                                           FLAG_NAPAROVANA,
                                           VZTAH,
                                           CONTRACTNR,
                                           STATUS,
                                           WFSTATUS,
                                           DATE_PAYOFF)
    select i3.id_cuid,
           i3.skp_person,
           i3.sr_jmeno,
           i3.sr_dtime_valid_from,
           i3.is_role_subj_v_ins_rizeni,
           i3.ic_stav_rizeni,
           i3.ic_datum_zalozeni_spisu,
           i3.ic_spisova_znacka,
           i3.datum_zmeny_stavu,
           i3.flag_naparovana,
           i3.vztah,
           i3.contractnr,
           i3.status,
           i3.wfstatus,
           i3.date_payoff 
    FROM APP_COLL_JEZEVCIK_TEST.insol3 i3
  union
    select a.ID_CUID, a.SKP_PERSON, a.SR_JMENO,  a.SR_DTIME_VALID_FROM,
           IS_ROLE_SUBJ_V_INS_RIZENI, IC_STAV_RIZENI, IC_DATUM_ZALOZENI_SPISU,
           IC_SPISOVA_ZNACKA, DATUM_ZMENY_STAVU, FLAG_NAPAROVANA, VZTAH ,
           b.text_bank_account_number as contractnr ,
           c.name_bank_account_status , c.name_bank_account_wfstatus ,
           a.date_payoff
    from
      APP_COLL_JEZEVCIK_TEST.insol3 a
      left join owner_dwh.f_bank_account_ad b on b.skp_person = a.skp_person  and b.id_bank_product = '1'
      left join owner_dwh.cl_bank_account_status c on b.skp_bank_account_status = c.skp_bank_account_status
    group by a.ID_CUID, a.SKP_PERSON, a.SR_JMENO,  a.SR_DTIME_VALID_FROM,
           IS_ROLE_SUBJ_V_INS_RIZENI, IC_STAV_RIZENI, IC_DATUM_ZALOZENI_SPISU,
           IC_SPISOVA_ZNACKA, DATUM_ZMENY_STAVU, FLAG_NAPAROVANA, VZTAH ,a.date_payoff ,
           b.text_bank_account_number , c.name_bank_account_status , c.name_bank_account_wfstatus
;

-- ins_contracts sa vyuziva jedine vo V_DATACHECK_CONTRACTS, ovsem prinos do VIEW je minimalny (vacsina info je duplicitna v inych tabulkach) (k 20.3.2018 zaznamy ku 3 klientom).
execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.ins_contracts drop storage');
insert into APP_COLL_JEZEVCIK_TEST.ins_contracts (SKP_PERSON,
                                                  CONTRACTNR,
                                                  STATUS,
                                                  WFSTATUS,
                                                  TYP_UCTU)
  select distinct q.skp_person , trim(contractnr) as contractnr , status , wfstatus, 
          case 
            when bs.NAME_BANK_ACCOUNT_TYPE = 'BÏûn˝ ˙Ëet' then 'B⁄'
            when bs.NAME_BANK_ACCOUNT_TYPE = 'Spo¯icÌ ˙Ëet' then 'S⁄'
            when bs.NAME_BANK_ACCOUNT_TYPE = 'Vnit¯nÌ ˙Ëet' then 'V⁄'
            when substr(contractnr,1,2) in ('HU', 'HR') then substr(contractnr,1,2)
            else bs.NAME_BANK_ACCOUNT_TYPE 
          end as typ_uctu
  from  APP_COLL_JEZEVCIK_TEST.insol4 q
  left join owner_dwh.f_bank_account_ad b on b.TEXT_BANK_ACCOUNT_NUMBER = q.contractnr
  left join owner_dwh.cl_bank_account_type bs on bs.SKP_BANK_ACCOUNT_type = b.SKP_BANK_ACCOUNT_type
  where contractnr is not null
;
execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.insol_list drop storage');
insert into APP_COLL_JEZEVCIK_TEST.insol_list (ID_CUID,
                                               SKP_PERSON,
                                               SR_JMENO,
                                               SR_DTIME_VALID_FROM,
                                               IS_ROLE_SUBJ_V_INS_RIZENI,
                                               IC_STAV_RIZENI,
                                               IC_DATUM_ZALOZENI_SPISU,
                                               IC_SPISOVA_ZNACKA,
                                               DATUM_ZMENY_STAVU,
                                               FLAG_NAPAROVANA,
                                               VZTAH,
                                               DATE_PAYOFF)
select distinct ID_CUID, SKP_PERSON, SR_JMENO, SR_DTIME_VALID_FROM,
       IS_ROLE_SUBJ_V_INS_RIZENI, IC_STAV_RIZENI, IC_DATUM_ZALOZENI_SPISU,
       rtrim(ltrim(IC_SPISOVA_ZNACKA)) as IC_SPISOVA_ZNACKA, DATUM_ZMENY_STAVU, FLAG_NAPAROVANA, VZTAH , date_payoff
from
insol4
;
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.insol  drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.insol2  drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.insol_loan  drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.insol_overdraft  drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.insol_codeb  drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.insol_collateral  drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.insol_debet  drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.insol3  drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.insol4  drop storage';
commit
;
end;



----------------------------------------------------------------------------------------------------
-- Aktualizace insolvenËn˝ch prÌpadov; vklad· novÈ zhody do tabulky ins_client_list
----------------------------------------------------------------------------------------------------
procedure UPDATE_INSOL_LIST IS

begin

insert into APP_COLL_JEZEVCIK_TEST.ins_client_list (SKP_PERSON,
                                                    EMPLOYEE,
                                                    CODE,
                                                    DTIME_MODIFIED)
select skp_person , 'sys' , ic_spisova_znacka as code , sysdate as dtime_modified from
APP_COLL_JEZEVCIK_TEST.insol_list il
where il.flag_naparovana = 'Y'
      and (il.skp_person , il.ic_spisova_znacka ) not in (select skp_person , code from APP_COLL_JEZEVCIK_TEST.ins_client_list)
;

-- klienti sa do tejto tabulky mÙûu prid·vaù aj manu·lne z jezevËÌka


begin
        dbms_stats.gather_table_stats(ownname => nazev_aktualniho_hriste,tabname => 'INS_CLIENT_LIST',
                                                  estimate_percent => dbms_stats.auto_sample_size,
                                                  cascade => true,
                                                  no_invalidate => false);
end;


commit;
END;



----------------------------------------------------------------------------------------------------
-- VloûenÌ nov˝ch klient˘ a jejich smluv do zdrojov˝ch tabulek aplikace: 
-- person; balance_obs; loan; instalment_head; person_address; person_demography; person_contact
----------------------------------------------------------------------------------------------------
procedure LOAD_PAYOFF_DATA IS
  skp INT;
  /*CURSOR PERSON IS
  select id_person,
         listagg(text_contract_number, ', ') within group(order by id_person) as smlouvy
    from update_deal
   group by id_person
  ; --jenom TEST RISK-985 */ 
  CURSOR NOVE_UVERY_ZESP IS -- RISK-1071
      SELECT l.ID_PERSON, l.CODE_LOAN_NUMBER
      FROM APP_COLL_JEZEVCIK_TEST.LOAN l
      JOIN APP_COLL_JEZEVCIK_TEST.person p ON p.ID_PERSON = l.ID_PERSON
      LEFT JOIN APP_COLL_JEZEVCIK_TEST.LOAN_YESTERDAY ly ON ly.SKP_LOAN = l.SKP_LOAN
      WHERE ly.SKP_LOAN IS NULL
        AND trunc(p.DTIME_VALID_FROM) < trunc(SYSDATE);

BEGIN


execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_person1 drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_person2 drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_person3 drop storage';



--RISK-985
-- RISK-1119 - v ramci update_person1 se taky dotahuji zesplatnene kontokorenty
insert into update_person1
SELECT main.skp_person,
                dp.id_cuid,
                dp.name_last || ' ' || dp.name_first as name,
                max(main.date_payoff) as date_payoff,
                listagg(main.name_payoff_reason,'; ') within group (order by name_payoff_reason) as payoff_reason,
                i.code,
                cat.text_contract_number
           from (SELECT a.SKP_PERSON, a.DATE_PAYOFF, a.NAME_PAYOFF_REASON FROM APP_COLL_JEZEVCIK_TEST.sr_coll_019_zes a
                 UNION
                 SELECT b.SKP_PERSON, b.DATE_PAYOFF, b.NAME_PAYOFF_REASON FROM APP_COLL_JEZEVCIK_TEST.OVERDRAFT_ZES b
                 ) main
           JOIN owner_dwh.dc_person dp ON main.SKP_PERSON = dp.SKP_PERSON
           left join (select distinct il.skp_person,
                              first_value(il.ic_stav_rizeni) over(partition by il.skp_person order by il.ic_datum_zalozeni_spisu DESC) code
                       from APP_COLL_JEZEVCIK_TEST.insol_list il
                      ) i 
                   on i.skp_person = main.skp_person 
           left join owner_dwh.f_contract_at cat 
                  on cat.skp_person = main.skp_person
                 and cat.flag_deleted <> 'Y'
                 and cat.skp_contract_type = 3
                 and cat.code_contract_status = '2'
                 and cat.dtime_contract_termination > sysdate
group by  main.skp_person,
          dp.id_cuid,
          dp.name_last || ' ' || dp.name_first ,
          --main.INSOLVENCY_CASE_STATUS  ,
          i.code,
          cat.text_contract_number
;





--RISK-985
insert into update_person2 -- tabulky manualne pridanych klientov
select DISTINCT p.skp_person , a.cuid , p.name_last ||' '|| p.name_first as name ,
                DATE'3000-01-01' as date_payoff , 
                'XNA' as payoff_reason ,
                --case when p.skp_person < -100 then '00000000000000' else null end as CODE_INSOLVENCY_CASE_STATUS ,
                max(cat.text_contract_number) as text_contract_number
          FROM APP_COLL_JEZEVCIK_TEST.cuid_addition a
               join owner_dwh.dc_person p on a.cuid = p.id_cuid
               left join owner_dwh.f_contract_at cat on cat.skp_person = p.skp_person
                      and cat.flag_deleted <> 'Y'
                      and cat.skp_contract_type = 3
                      and cat.code_contract_status = '2'
                      and cat.dtime_contract_termination > sysdate
         WHERE a.flag_addition <> 'Y'
               AND p.skp_person not in (select skp_person from update_person1)
      GROUP BY p.skp_person , a.cuid , p.name_last, p.name_first
;

-- vklad·m klientov , ktor˝ s˙ nap·rovan˝, prÌpadne manu·lne potvrden˝

--RISK-985
insert into update_person3
select distinct p.skp_person , p.id_cuid as cuid , p.name_last ||' '|| p.name_first as name ,
                max(date_payoff) as date_payoff , 'XNA' as payoff_reason ,
                --case when p.skp_person < -100 then '00000000000000' else null end as CODE_INSOLVENCY_CASE_STATUS ,
               max(cat.text_contract_number) as text_contract_number
          FROM APP_COLL_JEZEVCIK_TEST.ins_client_list a
               join APP_COLL_JEZEVCIK_TEST.insol_list b on a.skp_person = b.skp_person and a.code = b.ic_spisova_znacka
               join owner_dwh.dc_person p on b.skp_person  = p.skp_person
               left join owner_dwh.f_contract_at cat on cat.skp_person = p.skp_person
                      and cat.flag_deleted <> 'Y'
                      and cat.skp_contract_type = 3
                      and cat.code_contract_status = '2'
                      and cat.dtime_contract_termination > sysdate
         WHERE p.skp_person not in (select skp_person from update_person1)
                       and p.skp_person not in (select skp_person from update_person2)
      GROUP BY p.skp_person , p.id_cuid , p.name_last, p.name_first
;








/* RISK-398 */
/* Zmena stavu na aktivni */
-- Jen u tech, ktere bych tam nove chtel nalit, ale uz v mp_17_person jsou ale jako neaktivni


--RISK-985
UPDATE APP_COLL_JEZEVCIK_TEST.person
SET dtime_valid_from = sysdate,
    dtime_valid_to = date'3000-01-01',
    FLAG_VALID = 'Y'
WHERE skp_person in (select main.skp_person from (select SKP_PERSON from update_person1
                                                  union
                                                  select SKP_PERSON from update_person2
                                                  union
                                                  select SKP_PERSON from update_person3
                                                  ) main
                                            join APP_COLL_JEZEVCIK_TEST.person p
                                                 on p.skp_person = main.skp_person and p.flag_valid = 'N'
                     )
;

-- Archivace zmÏn se ¯eöÌ pomocÌ triggeru PERSON_AFTER_UPDATE


/* Pokud klient v tabulce PERSON jeste vubec neni, at uz jako aktivni nebo ne, tak jej tam nalej - stavajicim skriptem */
-- Z·roveÚ nastav platnosti od sysdate do 1.1.3000 a nastav flag valid na Y



/* RISK-398 */
/* Pokud klient v tabulce PERSON jeste vubec neni, at uz jako aktivni nebo ne, tak jej tam nalej - stavajicim skriptem */
-- Z·roveÚ nastav platnosti od sysdate do 1.1.3000 a nastav flag valid na Y

-- naplnÏnÌ tabulky klient˘ nov˝mi klienty, kte¯Ì majÌ na smlouvÏ evidovan˝ stav "ZesplatnÏn·"

--RISK-985
insert into APP_COLL_JEZEVCIK_TEST.person (ID_PERSON,
                                           CUID,
                                           NAME,
                                           DATE_PAYOFF,
                                           DTIME_MODIFICATION,
                                           EMPLOYEE,
                                           PAYOFF_REASON,
                                           CODE_INSOLVENCY_CASE_STATUS,
                                           DATE_DEATH,
                                           FLAG_PRINCIPAL_PAID_OBS,
                                           FLAG_INTEREST_PAID_OBS,
                                           FLAG_UNDESIRABILITY,
                                           TEXT_UNDESIRABILITY_NOTE,
                                           TEXT_CONTRACT_NUMBER,
                                           DTIME_VALID_FROM,
                                           DTIME_VALID_TO,
                                           FLAG_VALID,
                                           SKP_PERSON)
with
-- pouze ti klienti, kte¯Ì jeötÏ nejsou v tabulce klient˘ evidov·ni a nebo jsou neaktivni
sub_new as (
        select main.* from
        (select SKP_PERSON, ID_CUID, NAME, DATE_PAYOFF, PAYOFF_REASON, CODE_INSOLVENCY_CASE_STATUS , text_contract_number from
                update_person1

                union

                select SKP_PERSON, CUID, NAME, DATE_PAYOFF, PAYOFF_REASON, null , text_contract_number from
                update_person2

                union

                select SKP_PERSON, CUID, NAME, DATE_PAYOFF, PAYOFF_REASON, null , text_contract_number from
                update_person3
             ) main
                 left join APP_COLL_JEZEVCIK_TEST.person p
                           on p.skp_person = main.skp_person
           where p.skp_person is null
)
-- dotaûenÌ dalöÌch informacÌ ke klientovi (rodnÈ ËÌslo, adresy, ..)
--u cizincu, kteri nemaji RC, se nedotahne adresa

, distinkt AS  (
select DISTINCT 
                  new.id_cuid as CUID,
                  new.name,
                  date_payoff,
                  sysdate as dtime_modification,
                  'sys' as employee ,
                  g.payoff_reasons,
                  new.CODE_INSOLVENCY_CASE_STATUS,
                  nullif(p.date_death, to_date('01.01.3000','dd.mm.yyyy')) as date_death ,
                  'N' as flag_principal_paid_obs  ,
                  'N' as flag_interest_paid_obs ,
                  'Y' as flag_undesaribility ,
                  null as text_undesaribility_note ,
                  new.text_contract_number,
                  sysdate as dtime_valid_from, 
                  date'3000-01-01' as dtime_valid_to, 
                  'Y' as FLAG_VALID,
                  new.skp_person
    from sub_new new
    -- v p¯ÌpadÏ, ûe je vÌce d˘vodu zesplatnÏnÌ, jsou vyps·ny s Ë·rkou za sebou
         join (SELECT skp_person,
         listagg(to_char(payoff_reason), ', '||chr(10)) within group (order by payoff_reason) as payoff_reasons
        FROM sub_new
  GROUP BY skp_person) g on g.skp_person = new.skp_person
        join owner_dwh.dc_person p
             on p.skp_person = new.skp_person
)

SELECT person_seq.nextval, d.*
FROM distinkt d
;


begin
        dbms_stats.gather_table_stats(ownname => nazev_aktualniho_hriste,tabname => 'PERSON',
                                                  estimate_percent => dbms_stats.auto_sample_size,
                                                  cascade => true,
                                                  no_invalidate => false);
end;

-- MM 19.3.2018 p¯id·na podmÌnka, - flag se updateuje pouze tehdy, ûe tam klient skuteËnÏ p¯id·n byl - je v tabulce PERSON
UPDATE APP_COLL_JEZEVCIK_TEST.cuid_addition ca
set ca.flag_addition = 'Y'
WHERE EXISTS (SELECT p.CUID 
              FROM APP_COLL_JEZEVCIK_TEST.PERSON P
              WHERE p.CUID = ca.CUID)
;



-- RISK-952 - hled·nÌ neklient˘, kte¯Ì by se mohli st·t klienty
-- KÛd do procedury, plnÌ se inkrement·lnÏ - nikdy nemazat z·znamy
INSERT INTO  APP_COLL_JEZEVCIK_TEST.non_client_match (skp_non_client_match,
                                                     id_person,
                                                     cuid_jez,
                                                     name_jez,
                                                     date_birth_jez,
                                                     text_birth_number_jez,
                                                     skp_person_dwh,
                                                     cuid_dwh,
                                                     name_last_dwh,
                                                     name_first_dwh,
                                                     date_birth_dwh,
                                                     text_birth_number_dwh,
                                                     match_type,
                                                     dtime_inserted,
                                                     user_inserted
                                                     )
with neklienti as ( -- b·ze vöech neklient˘

select jp.id_person 
     , jp.cuid
     , jp.name
     , jpd.date_birth
     , jpd.text_birth_number
  
  from  APP_COLL_JEZEVCIK_TEST.person jp
  join  APP_COLL_JEZEVCIK_TEST.person_demography jpd
    on jpd.id_person = jp.id_person
  
 where jp.skp_person < 0
   and jp.flag_valid = 'Y'
   -- jeötÏ nebyl potvrzen nebo zamÌtnut
   AND jp.id_person NOT IN (SELECT id_person FROM  APP_COLL_JEZEVCIK_TEST.non_client_match WHERE name_confirmation_status IN ('Potvrzeno')) -- podminka jenom pro jistotu, klient uz stejne bude mit skp_person > 0
   
)

, rodne_cislo as ( -- p·rov·nÌ p¯es rodnÈ ËÌslo, pokud m· v JezevËÌku vyplnÏno

select j.id_person as id_person
     , j.cuid as cuid_jez
     , j.name as name_jez
     , j.date_birth as date_birth_jez
     , j.text_birth_number as text_birth_number_jez
     , p.SKP_PERSON as skp_person_dwh
     , p.ID_CUID as cuid_dwh
     , p.NAME_LAST as name_last_dwh
     , p.NAME_FIRST as name_first_dwh
     , p.date_birth as date_birth_dwh
     , p.text_birth_number as text_birth_number_dwh
     , 'RodnÈ ËÌslo' as match_type

  from neklienti j
  
  join owner_dwh.dc_person p
    on p.TEXT_BIRTH_NUMBER = j.text_birth_number
    
 where j.text_birth_number <> 'XNA'
)

, jmeno_narozeni as ( -- p·rov·nÌ p¯es jmÈno a datum narozenÌ, pokud rodnÈ ËÌslo nem· v JezevËÌku vyplnÏno

select j.id_person as id_person
     , j.cuid as cuid_jez
     , j.name as name_jez
     , j.date_birth as date_birth_jez
     , j.text_birth_number as text_birth_number_jez
     , p.SKP_PERSON as skp_person_dwh
     , p.ID_CUID as cuid_dwh
     , p.NAME_LAST as name_last_dwh
     , p.NAME_FIRST as name_first_dwh
     , p.date_birth as date_birth_dwh
     , p.text_birth_number as text_birth_number_dwh
     , 'JmÈno a datum narozenÌ' as match_type
  
  from neklienti j
  
  join owner_dwh.dc_person p
    on ignoruj_diak(j.name) like '%' || ignoruj_diak(p.NAME_FIRST) || '%'
   and ignoruj_diak(j.name) like '%' || ignoruj_diak(p.NAME_LAST) || '%'
   and j.date_birth = p.DATE_BIRTH
    
 where j.text_birth_number = 'XNA'
)

, match AS (
-- d·v·m dohromady oba typy shody 
   select ID_PERSON, CUID_JEZ, NAME_JEZ, DATE_BIRTH_JEZ, TEXT_BIRTH_NUMBER_JEZ, SKP_PERSON_DWH, CUID_DWH, NAME_LAST_DWH, NAME_FIRST_DWH, DATE_BIRTH_DWH, TEXT_BIRTH_NUMBER_DWH, MATCH_TYPE 
   from rodne_cislo
UNION 
   SELECT ID_PERSON, CUID_JEZ, NAME_JEZ, DATE_BIRTH_JEZ, TEXT_BIRTH_NUMBER_JEZ, SKP_PERSON_DWH, CUID_DWH, NAME_LAST_DWH, NAME_FIRST_DWH, DATE_BIRTH_DWH, TEXT_BIRTH_NUMBER_DWH, MATCH_TYPE 
   from jmeno_narozeni
 )

 SELECT  APP_COLL_JEZEVCIK_TEST.non_client_match_seq.nextval
        ,ID_PERSON, CUID_JEZ, NAME_JEZ, DATE_BIRTH_JEZ, TEXT_BIRTH_NUMBER_JEZ, SKP_PERSON_DWH, CUID_DWH, NAME_LAST_DWH, NAME_FIRST_DWH, DATE_BIRTH_DWH, TEXT_BIRTH_NUMBER_DWH, MATCH_TYPE 
        ,SYSDATE, 'sys'
        
 FROM match m  
 -- dan˝ p·r tam jeötÏ nenÌ
 WHERE id_person||SKP_PERSON_DWH NOT IN (SELECT id_person||SKP_PERSON_DWH FROM  APP_COLL_JEZEVCIK_TEST.non_client_match)
 ;


execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_person1 drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_person2 drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_person3 drop storage';


commit;

-- p¯id·nÌ nov˝ch smluv (z tabulky reportu) do tabulky, kter· eviduje ËÌsla ˙vÏr˘ ke klient˘m


--RISK-985  -  DEAL existuje jenom na PROD
/* -- zakomentovano v RISK-1071
insert into APP_COLL_JEZEVCIK_TEST.deal p
  select new.id_person,
         new.code_loan_number as text_contract_number
    from (select distinct prso.id_person,
                          main.skp_loan,
                          main.code_loan_number
            from APP_COLL_JEZEVCIK_TEST.sr_coll_019_zes main
            JOIN APP_COLL_JEZEVCIK_TEST.person prso ON main.skp_person = prso.skp_person
                 left join APP_COLL_JEZEVCIK_TEST.deal p
                           on p.id_person = prso.id_person
           where p.id_person is null) new --pokud klient uz je jednou v deal, nedostaneme se k nemu nova zesplatnena smlouva
   ;


--zde nahrajeme nove zesplatnene smlouvy ke klientum, kteri jiz jsou v tabulce deal

insert into APP_COLL_JEZEVCIK_TEST.update_deal
  select new.id_person,
         new.code_loan_number as text_contract_number
    from (select distinct prso.id_person,
                          main.skp_loan,
                          main.code_loan_number
            from APP_COLL_JEZEVCIK_TEST.sr_coll_019_zes main
            JOIN APP_COLL_JEZEVCIK_TEST.person prso ON main.skp_person = prso.skp_person
                 left join APP_COLL_JEZEVCIK_TEST.deal p
                           on p.text_contract_number = main.code_loan_number
                 join APP_COLL_JEZEVCIK_TEST.deal p on p.id_person = prso.id_person
           where p.id_person is null) new
   ;

commit;



insert into APP_COLL_JEZEVCIK_TEST.deal
select * from APP_COLL_JEZEVCIK_TEST.update_deal;

--vlozeni upozorneni na pridani dalsi smlouvy

  FOR J IN PERSON LOOP
    -- zÌsk·nÌ dalöÌho id pro novÈ z·znamy do tabulky upozornÏnÌ
    select max(skp_notice) into skp from APP_COLL_JEZEVCIK_TEST.notice;
    if skp is null then
      skp := 1;
    else
      skp := skp + 1 ;
    END IF;

    -- vloûenÌ upozornÏnÌ
    insert into APP_COLL_JEZEVCIK_TEST.notice t (t.skp_notice, t.id_person, t.text_notice, t.dtime_notice_created, t.dtime_notice_deadline, t.code_status, t.employee)
    values (skp, J.id_person, 'UpozornÏnÌ: P¯id·na novÏ zesplatnÏn· smlouva ' || J.SMLOUVY|| '.', trunc(sysdate), trunc(sysdate) + 30, 'a', 'sys');



  END LOOP;
  COMMIT;

execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_deal drop storage';
*/




/* RISK-442 */
/* RISK-481 27.2.2017 */
-- Tabulka bude slouzit pro dodavani aktualnich dat do objektu klienta-loan
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.loan_yesterday drop storage';

-- Tabulka LOAN_YESTERDAY obsahuje informace z LOAN opoûdenÈ o 1 den (vÌce pokud nefungoval chain).
-- vytvorena v RISK-1071 kvuli upozorneni 'P¯id·na novÏ zesplatnÏn· smlouva ... ' 
INSERT INTO APP_COLL_JEZEVCIK_TEST.LOAN_YESTERDAY (SKP_LOAN,
                                                   ID_PERSON,
                                                   CODE_LOAN_NUMBER,
                                                   AMT_INSTALMENT,
                                                   NAME_LOAN_STATUS,
                                                   NAME_LOAN_WFSTATUS,
                                                   DATE_LOAN_UTILIZATION,
                                                   DATE_PAYOFF,
                                                   DATE_LOAN_CLOSE,
                                                   DATE_LOAN_WRITTEN_OFF,
                                                   AMT_LOAN_PRINCIPAL_TOTAL,
                                                   TEXT_BANK_ACCOUNT_NUMBER,
                                                   DTIME_SIGNATURE,
                                                   TEXT_APPENDIX_NUMBER,
                                                   AMT_LOAN_PRINCIPAL_PAID,
                                                   DATE_INSTALMENT_FIRST,
                                                   DATE_INSTALMENT_LAST,
                                                   DATE_LOAN_CLOSE_PLAN,
                                                   AMT_LOAN_LIMIT,
                                                   RATE_INTEREST_STANDARD,
                                                   RATE_EIR_STANDARD,
                                                   RATE_INTEREST_REPO)
SELECT l.SKP_LOAN,
       l.ID_PERSON,
       l.CODE_LOAN_NUMBER,
       l.AMT_INSTALMENT,
       l.NAME_LOAN_STATUS,
       l.NAME_LOAN_WFSTATUS,
       l.DATE_LOAN_UTILIZATION,
       l.DATE_PAYOFF,
       l.DATE_LOAN_CLOSE,
       l.DATE_LOAN_WRITTEN_OFF,
       l.AMT_LOAN_PRINCIPAL_TOTAL,
       l.TEXT_BANK_ACCOUNT_NUMBER,
       l.DTIME_SIGNATURE,
       l.TEXT_APPENDIX_NUMBER,
       l.AMT_LOAN_PRINCIPAL_PAID,
       l.DATE_INSTALMENT_FIRST,
       l.DATE_INSTALMENT_LAST,
       l.DATE_LOAN_CLOSE_PLAN,
       l.AMT_LOAN_LIMIT,
       l.RATE_INTEREST_STANDARD,
       l.RATE_EIR_STANDARD,
       l.RATE_INTEREST_REPO 
FROM APP_COLL_JEZEVCIK_TEST.loan l;

execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.loan drop storage';  -- VYMAZAT AZ PO NAPLNENI LOAN_YESTERDAY

insert into APP_COLL_JEZEVCIK_TEST.loan (SKP_LOAN,   -- NAPLNIT AZ PO NAPLNENI LOAN_YESTERDAY
                                         ID_PERSON,
                                         CODE_LOAN_NUMBER,
                                         AMT_INSTALMENT,
                                         NAME_LOAN_STATUS,
                                         NAME_LOAN_WFSTATUS,
                                         DATE_LOAN_UTILIZATION,
                                         DATE_PAYOFF,
                                         DATE_LOAN_CLOSE,
                                         DATE_LOAN_WRITTEN_OFF,
                                         AMT_LOAN_PRINCIPAL_TOTAL,
                                         TEXT_BANK_ACCOUNT_NUMBER,
                                         DTIME_SIGNATURE,
                                         TEXT_APPENDIX_NUMBER,
                                         AMT_LOAN_PRINCIPAL_PAID,
                                         DATE_INSTALMENT_FIRST,
                                         DATE_INSTALMENT_LAST,
                                         DATE_LOAN_CLOSE_PLAN,
                                         AMT_LOAN_LIMIT,
                                         RATE_INTEREST_STANDARD,
                                         RATE_EIR_STANDARD,
                                         RATE_INTEREST_REPO)
select l.SKP_LOAN, p.id_person, l.CODE_LOAN_NUMBER, l.AMT_INSTALMENT, cls.NAME_LOAN_STATUS, cls.NAME_LOAN_WFSTATUS
       , l.DATE_LOAN_UTILIZATION --RISK-940
       , l.DATE_PAYOFF, l.DATE_LOAN_CLOSE, l.DATE_LOAN_WRITE_OFF, l.AMT_LOAN_PRINCIPAL_TOTAL, ba.TEXT_BANK_ACCOUNT_NUMBER, PCS.DTIME_SIGNATURE
       , PC.text_appendix_number, l.AMT_LOAN_PRINCIPAL_PAID, l.DATE_INSTALMENT_FIRST, l.DATE_INSTALMENT_LAST, l.DATE_LOAN_CLOSE_PLAN
       , amt_loan_limit, rate_interest_standard, rate_eir_standard, rate_interest_repo -- RISK-726, nahrazeni tabulky SP7
from APP_COLL_JEZEVCIK_TEST.person p
join owner_dwh.f_loan_at l on p.skp_person = l.skp_person and p.flag_valid = 'Y' and l.DATE_LOAN_UTILIZATION <> date'3000-01-01' -- RISK-940
join owner_dwh.cl_loan_status cls on l.SKP_LOAN_STATUS = cls.SKP_LOAN_STATUS and cls.flag_deleted <> 'Y'
join owner_dwh.f_bank_account_ad ba on ba.skp_bank_account = l.skp_bank_account
JOIN (SELECT DISTINCT lst.SKP_LOAN                                                            -- RISK-1071 FA: 19.3.2018: jenom uvery, ktere byli kdysi zesplatnene (nezajimaji nas davno ukoncene)
        FROM owner_dwh.f_loan_status_tt lst
        JOIN owner_dwh.cl_loan_status ls ON ls.SKP_LOAN_STATUS = lst.SKP_LOAN_STATUS
        WHERE lst.FLAG_DELETED = 'N'
          AND ls.NAME_LOAN_WFSTATUS = 'ZesplatnÏn˝'
      ) zesp ON zesp.skp_loan = l.skp_loan
--DATUM PODPISU UVERU °°°°°°°°°°°°°°°°°°°°°
left JOIN   --kompletace (v ppf_completionsignature lze dohledat pouze hierarchicky nejvyööÌ idcompletion -> nutno vytvo¯it strukturu viz. nÌûe)
      (--je û·d·no pouze o ˙vÏr
      select PC.SKP_LOAN,  pc.skp_completion_subject , pc.id_source_application as cashLoanIdRequest, pc.skp_completion_subject as idCompletionParentFinal, pc.text_appendix_number
      from OWNER_DWH.f_completion_subject_at pc
      where pc.code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING','RHY_LOAN_COMMITMENT') and pc.skp_completion_parent ='-1'
      union
      --˙vÏr + BU
      select PC.SKP_LOAN,  pc.skp_completion_subject , pc.id_source_application as cashLoanIdRequest, pc1.skp_completion_subject as idCompletionParentFinal, pc.text_appendix_number
      from OWNER_DWH.f_completion_subject_at pc
      join OWNER_DWH.f_completion_subject_at  pc1 on (pc1.skp_completion_subject = pc.skp_completion_parent and pc1.skp_completion_parent ='-1')
      where pc.code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING','RHY_LOAN_COMMITMENT') and pc.skp_completion_parent <>'-1'
      union
      --˙vÏr + BU + RS
      select PC.SKP_LOAN,  pc.skp_completion_subject , pc.id_source_application as cashLoanIdRequest, pc1.skp_completion_parent as idCompletionParentFinal, pc.text_appendix_number
      from OWNER_DWH.f_completion_subject_at pc
      join OWNER_DWH.f_completion_subject_at pc1 on (pc1.skp_completion_subject = pc.SKP_COMPLETION_PARENT and pc1.skp_completion_parent <>'-1')
      where pc.code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING','RHY_LOAN_COMMITMENT') and pc.SKP_COMPLETION_PARENT <>'-1' ) pc ON PC.SKP_LOAN = L.SKP_LOAN
left join OWNER_DWH.f_completion_signature_tt pcs on (pcs.skp_completion_subject = pc.idCompletionParentFinal and pcs.dtime_signature_valid_to > sysdate ) /*kv˘li moûnÈmu p¯egenerov·nÌ smlouvy*/
--DATUM PODPISU UVERU ^^^^^^^^^^^^^^^^^
;

commit;




-- VLOZENI UPOZORNENI O NOVYCH ZESPLATNENYCH SMLOUVACH  -- RISK-1071

  FOR J IN NOVE_UVERY_ZESP LOOP
    
    -- zÌsk·nÌ dalöÌho id pro novÈ z·znamy do tabulky upozornÏnÌ
    select max(skp_notice) into skp from APP_COLL_JEZEVCIK_TEST.notice;
    if skp is null then
      skp := 1;
    else
      skp := skp + 1 ;
    END IF;

    -- vloûenÌ upozornÏnÌ
    insert into APP_COLL_JEZEVCIK_TEST.notice t (t.skp_notice, t.id_person, t.text_notice, t.dtime_notice_created, t.dtime_notice_deadline, t.code_status, t.employee)
    values (skp, J.id_person, 'UpozornÏnÌ: P¯id·na novÏ zesplatnÏn· smlouva ' || J.CODE_LOAN_NUMBER|| '.', trunc(sysdate), trunc(sysdate) + 30, 'a', 'sys');

  END LOOP;
  COMMIT;




-- tabulka s aktualnou pohledavkou dle obs
execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.balance_obs drop storage');

-- RISK-985, DEAL jenom na PROD
insert into APP_COLL_JEZEVCIK_TEST.balance_obs (ID_PERSON,
                                                SKP_LOAN,
                                                AMT_PAYOFF_PRINCIPAL,
                                                AMT_INTEREST_OVERDUE,
                                                AMT_PAYOFF_INTEREST,
                                                AMT_PAYOFF_PENALTY,
                                                AMT_PAYOFF)

/*       nvl(c.amt_payoff_interest,0) - nvl(c.amt_interest_overdue,0) as amt_payoff_interest,
       nvl(c.amt_payoff_penalty,0) as amt_payoff_penalty , --puvodni verze
*/

SELECT b.id_person , b.skp_loan , nvl(c.amt_payoff_principal,0)  as amt_payoff_principal,
-- MM: Oprava 31.7.2017 - HOTFIX chyby v f_loan_balance_at -- select a.skp_person , b.skp_loan , nvl(c.amt_balance_principal,0)  as amt_payoff_principal,
-- FA: vratene na povodny skript 29.8.2017

       nvl(c.amt_interest_overdue,0) as amt_interest_overdue,

/*pokud je uhrazena jistina a komplet obchodni urok (vcetne aktualniho v mesici), na dwh i v obs se po zaplaceni do konce mesice nevynuluje akt. obch. a sankcni urok,
proto je potreba to pocitat trochu jinak
*/

      case when amt_penalty_paid_late >= d.amt_penalty_overdue and amt_interest_overdue = 0 and amt_penalty_interest_current > 0 then 
             nvl(c.amt_penalty_interest_overdue, 0) else
             nvl(c.amt_penalty_interest_overdue, 0)  +
             nvl(c.amt_penalty_interest_current, 0) end as amt_payoff_interest,
             
           case when amt_penalty_paid_late >= d.amt_penalty_overdue and amt_interest_overdue = 0 and amt_penalty_interest_current > 0 then
             nvl(c.amt_penalty_overdue, 0) else       
             nvl(c.amt_penalty_overdue, 0) +
             nvl(c.amt_penalty_current, 0) end as amt_payoff_penalty,
--zmena owner_dwh 11.7.2016, tohle by mel byt spravny napocet obchodniho a sankcniho uroku

  /*           case when c.amt_penalty_interest_overdue = 0 and c.amt_penalty_overdue > 0 then
             nvl(c.amt_penalty_interest_overdue, 0) else
             nvl(c.amt_penalty_interest_overdue, 0)  +
             nvl(c.amt_penalty_interest_current, 0) end as amt_payoff_interest,

             case when c.amt_penalty_interest_overdue = 0 and c.amt_penalty_overdue > 0 then
             nvl(c.amt_penalty_overdue, 0) else
             nvl(c.amt_penalty_overdue, 0) +
             nvl(c.amt_penalty_current, 0) end as amt_payoff_penalty,*/

       /* nvl(c.amt_penalty_interest_overdue,0) + nvl(c.amt_penalty_interest_current,0) as amt_payoff_interest,
        nvl(c.amt_penalty_overdue,0) + nvl(c.amt_penalty_current,0) as  amt_payoff_penalty,*/
        /*nvl(c.amt_payoff,0)*/ -- Nahrazeno za radek nize - u zopar loanu ostavala v amt_payoff castka z konce minuleho mesice > nahlasena DWH chyba LOANu -- FA 7.3.2018 -- BUG_amt_payoff
        (nvl(c.amt_payoff_principal,0) + nvl(c.amt_payoff_penalty,0) + nvl(c.amt_payoff_interest,0) + nvl(c.AMT_PENALTY_CURRENT,0) + nvl(c.AMT_PENALTY_INTEREST_CURRENT,0)) AS amt_payoff


from APP_COLL_JEZEVCIK_TEST.loan b
     join owner_dwh.f_loan_balance_at c on b.skp_loan = c.skp_loan
      left join (select skp_loan , max(amt_penalty_overdue) amt_penalty_overdue 
      from owner_dwh.f_loan_balance_at 
      group by skp_loan) d on d.skp_loan = b.skp_loan
WHERE c.flag_deleted <> 'Y'
      and c.flag_valid = 'Y'
;



-- MM: Oprava 31.7.2017 - HOTFIX chyby v f_loan_balance_at
-- FA vratene(zakomentovane) 29.8.2017 - bug na DWH opraven
/*update APP_COLL_JEZEVCIK_TEST.balance_obs a
set a.amt_payoff = a.amt_payoff_principal + a.amt_interest_overdue + a.amt_payoff_interest + a.amt_payoff_penalty
;*/


begin
        dbms_stats.gather_table_stats(ownname => nazev_aktualniho_hriste,tabname => 'balance_obs',
                                                  estimate_percent => dbms_stats.auto_sample_size,
                                                  cascade => true,
                                                  no_invalidate => false);
end;


commit;




execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.INSTALMENT_HEAD drop storage';
commit;
Insert into APP_COLL_JEZEVCIK_TEST.INSTALMENT_HEAD (SKF_INSTALMENT_HEAD,
                                               FLAG_DELETED,
                                               SKP_LOAN,
                                               ID_PERSON,
                                               SKF_LOAN_SCHEDULE,
                                               CODE_STATUS,
                                               DATE_INSTALMENT,
                                               DATE_PAST_DUE_TOL_J,
                                               DATE_PAST_DUE_TOL_J_PAID,
                                               NUM_INSTALMENT_NUMBER,
                                               NUM_INSTALMENT_VERSION,
                                               NUM_LOAN_SCHEDULE_ORDER,
                                               AMT_INSTALMENT,
                                               AMT_INSTALMENT_PRINCIPAL,
                                               AMT_INSTALMENT_INTEREST,
                                               AMT_PAYMENT,
                                               AMT_PAYMENT_PRINCIPAL,
                                               AMT_PAYMENT_INTEREST,
                                               FLAG_VALID,
                                               DTIME_VALID_FROM,
                                               DTIME_VALID_TO,
                                               DTIME_INSERTED,
                                               DTIME_MODIFIED,
                                               DATE_PAY_INSTALMENT,
                                               DATE_PAY_LAST)
select  
            i.SKF_INSTALMENT_HEAD, 
            i.FLAG_DELETED, 
            i.SKP_LOAN, 
            p.id_person, 
            /*i.SKP_LOAN_SCHEDULE_TYPE, */
            i.SKF_LOAN_SCHEDULE,
            /*i.SKP_APPLICATION,*/
            i.CODE_STATUS,
            /*i.TEXT_LOAN_PRINCIPAL, */
            i.DATE_INSTALMENT, 
            i.DATE_PAST_DUE_TOL_J,
            i.DATE_PAST_DUE_TOL_J_PAID, 
            /*i.NUM_INSTALMENT_ORDER, */
            i.NUM_INSTALMENT_NUMBER, 
            i.NUM_INSTALMENT_VERSION, 
            i.NUM_LOAN_SCHEDULE_ORDER, 
            i.AMT_INSTALMENT, 
            i.AMT_INSTALMENT_PRINCIPAL, 
            i.AMT_INSTALMENT_INTEREST, 
            i.AMT_PAYMENT,
            i.AMT_PAYMENT_PRINCIPAL,
            i.AMT_PAYMENT_INTEREST, 
            /*i.AMT_PRINCIPAL_UNPAID,
            i.AMT_INTEREST_UNPAID,
            i.CNT_DAYS_PAST_DUE_TOL_J, */
            i.FLAG_VALID,
            i.DTIME_VALID_FROM, 
            i.DTIME_VALID_TO, 
            i.DTIME_INSERTED,
            i.DTIME_MODIFIED, 
            i.DATE_PAY_INSTALMENT,
            i.DATE_PAY_LAST
from APP_COLL_JEZEVCIK_TEST.person p
join owner_dwh.f_loan_at l on l.skp_person = p.skp_person
join owner_dwh.f_instalment_head_ad i on l.skp_loan = i.skp_loan and i.skp_person = p.skp_person
where (trunc(i.DTIME_INSERTED) >= trunc(l.DATE_LOAN_UTILIZATION))  -- RISK-940
and i.date_instalment <= sysdate --omezit budoucÌ spl·tky
-- and  num_loan_schedule_order = 0
and p.flag_valid = 'Y'
--AND i.FLAG_DELETED = 'N' -- podminky i.FLAG_DELETED = 'N', i.CODE_STATUS = 'a' musime vypustit kvoli sposobu, akym sa pocitaju tabulky sablona_epr 4 a sablona_epr 5 - bol tam problem parovania.
--AND i.CODE_STATUS = 'a'
;
commit;




-- RISK-1145 odtial
  execute immediate 'TRUNCATE TABLE app_coll_jezevcik_test.OVERDRAFT drop storage';
  INSERT INTO app_coll_jezevcik_test.OVERDRAFT (ID_PERSON, SKP_BANK_ACCOUNT, SKP_BANK_ACCOUNT_DEBIT, SKP_OVERDRAFT, CODE_OVERDRAFT_NUMBER,
                                                AMT_OVERDRAFT_LIMIT, DATE_OVERDRAFT_OPEN, DATE_OVERDRAFT_UTILIZATION, SKP_OVERDRAFT_STATUS, NAME_OVERDRAFT_STATUS,
                                                NAME_OVERDRAFT_WFSTATUS, AMT_OVERDRAFT, AMT_OVERDRAFT_OVERDUE, AMT_OVERDRAFT_INTEREST_TO_PAY, AMT_INTEREST_OVERDUE,
                                                AMT_INTEREST_PAID, AMT_INTEREST_PAID_TOTAL, AMT_INTEREST_TOTAL, AMT_INTEREST_LAST, RATE_PENALTY,
                                                RATE_PENALTY_INTEREST, AMT_PAYOFF, AMT_PAYOFF_PRINCIPAL, AMT_PAYOFF_INTEREST, SKP_CURRENCY)
  SELECT p.ID_PERSON , ov.SKP_BANK_ACCOUNT , ov.SKP_BANK_ACCOUNT_DEBIT , ov.SKP_OVERDRAFT , ov.CODE_OVERDRAFT_NUMBER
       , ov.AMT_OVERDRAFT_LIMIT , ov.DATE_OVERDRAFT_OPEN , ov.DATE_OVERDRAFT_UTILIZATION , ov.SKP_OVERDRAFT_STATUS , st.NAME_OVERDRAFT_STATUS
       , st.NAME_OVERDRAFT_WFSTATUS , ov.AMT_OVERDRAFT , ov.AMT_OVERDRAFT_OVERDUE , ov.AMT_OVERDRAFT_INTEREST_TO_PAY , ov.AMT_INTEREST_OVERDUE
       , ov.AMT_INTEREST_PAID , ov.AMT_INTEREST_PAID_TOTAL , ov.AMT_INTEREST_TOTAL , ov.AMT_INTEREST_LAST , ov.RATE_PENALTY
       , ov.RATE_PENALTY_INTEREST , ov.AMT_PAYOFF , ov.AMT_PAYOFF_PRINCIPAL , ov.AMT_PAYOFF_INTEREST , ov.SKP_CURRENCY
  FROM owner_dwh.F_OVERDRAFT_AT ov
  JOIN app_coll_jezevcik_test.person p
    ON ov.skp_person = p.SKP_PERSON
  JOIN owner_dwh.CL_OVERDRAFT_STATUS st
    ON ov.SKP_OVERDRAFT_STATUS = st.SKP_OVERDRAFT_STATUS
  WHERE ov.FLAG_DELETED = 'N'
  ;


  execute immediate 'TRUNCATE TABLE app_coll_jezevcik_test.OVERDRAFT_BALANCE_PD drop storage';
  INSERT INTO app_coll_jezevcik_test.OVERDRAFT_BALANCE_PD (ID_PERSON, SKP_OVERDRAFT, DATE_EFFECTIVE, SKP_CONTRACT, SKP_OVERDRAFT_STATUS,
                                                           NAME_OVERDRAFT_STATUS, NAME_OVERDRAFT_WFSTATUS, DATE_PERIOD, AMT_OVERDRAFT, AMT_OVERDRAFT_LIMIT,
                                                           AMT_OVERDRAFT_OVERDUE, AMT_PRINCIPAL_OVERDUE, AMT_INTEREST_OVERDUE, AMT_PRINCIPAL_FUTURE, AMT_INTEREST_DUE,
                                                           AMT_INTEREST_PAID, AMT_PRINCIPAL_PAID, AMT_INTEREST_PAID_TOTAL, AMT_PRINCIPAL_PAID_TOTAL, AMT_INTEREST,
                                                           AMT_INTEREST_TOTAL, AMT_INTEREST_LAST, AMT_PENALTY_CURRENT, AMT_PENALTY_INTEREST_CURRENT, AMT_PRINCIPAL_DRAW,
                                                           RATE_PENALTY, RATE_PENALTY_INTEREST, DATE_PAYOFF, AMT_PAYOFF, AMT_PAYOFF_PRINCIPAL, AMT_PAYOFF_INTEREST)
  SELECT p.ID_PERSON , bal.SKP_OVERDRAFT , bal.DATE_EFFECTIVE , bal.SKP_CONTRACT , bal.SKP_OVERDRAFT_STATUS
       , st.NAME_OVERDRAFT_STATUS , st.NAME_OVERDRAFT_WFSTATUS , bal.DATE_PERIOD , bal.AMT_OVERDRAFT , bal.AMT_OVERDRAFT_LIMIT
       , bal.AMT_OVERDRAFT_OVERDUE , bal.AMT_PRINCIPAL_OVERDUE , bal.AMT_INTEREST_OVERDUE , bal.AMT_PRINCIPAL_FUTURE , bal.AMT_INTEREST_DUE
       , bal.AMT_INTEREST_PAID , bal.AMT_PRINCIPAL_PAID , bal.AMT_INTEREST_PAID_TOTAL , bal.AMT_PRINCIPAL_PAID_TOTAL , bal.AMT_INTEREST
       , bal.AMT_INTEREST_TOTAL , bal.AMT_INTEREST_LAST  , bal.AMT_PENALTY_CURRENT  , bal.AMT_PENALTY_INTEREST_CURRENT   , bal.AMT_PRINCIPAL_DRAW  
       , bal.RATE_PENALTY   , bal.RATE_PENALTY_INTEREST , bal.date_payoff , bal.AMT_PAYOFF , bal.AMT_PAYOFF_PRINCIPAL , bal.AMT_PAYOFF_INTEREST
  FROM owner_dwh.F_OVERDRAFT_BALANCE_PD bal
  JOIN app_coll_jezevcik_test.person p ON bal.skp_person = p.SKP_PERSON
  JOIN owner_dwh.CL_OVERDRAFT_STATUS st
    ON bal.SKP_OVERDRAFT_STATUS = st.SKP_OVERDRAFT_STATUS
  WHERE bal.FLAG_DELETED = 'N'
  ;
  
  EXECUTE IMMEDIATE 'TRUNCATE TABLE app_coll_jezevcik_test.overdraft_posledni_zustatek drop storage';
  INSERT INTO app_coll_jezevcik_test.overdraft_posledni_zustatek (SKP_OVERDRAFT, POSLEDNI_KLADNY_ZUSTATEK)
    SELECT o.SKP_OVERDRAFT, ab.AMT_BALANCE AS posledni_kladny_zustatek
    FROM app_coll_jezevcik_test.OVERDRAFT o
    JOIN app_coll_jezevcik_test.V_OVERDRAFT_DATUM_ZANORENI dtm ON dtm.SKP_OVERDRAFT = o.SKP_OVERDRAFT
    JOIN owner_dwh.f_account_balance_pd ab ON ab.SKP_BANK_ACCOUNT = o.SKP_BANK_ACCOUNT AND ab.DATE_PERIOD = TRUNC(dtm.datum_posledni_zanoreni)
    WHERE ab.FLAG_DELETED = 'N'
    ;
  
  COMMIT;
-- RISK-1145 potial



-- RISK 385
-- Pridani novych adres, vzdy aktualnich ke vcerejsku pro vsechny klienty. Odebrano z PERSON a pridano do ADDRESS
--execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.person_address drop storage'; --RISK-468
DELETE FROM APP_COLL_JEZEVCIK_TEST.person_address a WHERE a.id_person IN (SELECT id_person FROM person p WHERE p.skp_person > 0);

insert into APP_COLL_JEZEVCIK_TEST.person_address (SKP_PERSON_ADDRESS,
                                                   ID_PERSON,
                                                   NAME_REPORTS_ADDRESS_STREET,
                                                   TEXT_REPORTS_ADDRESS_NUM,
                                                   CODE_REPORTS_ADDRESS_ZIP,
                                                   NAME_REPORTS_ADDRESS_TOWN,
                                                   NAME_REPORTS_ADDRESS_COUNTRY,
                                                   NAME_PERMANENT_ADDRESS_STREET,
                                                   TEXT_PERMANENT_ADDRESS_NUM,
                                                   CODE_PERMANENT_ADDRESS_ZIP,
                                                   NAME_PERMANENT_ADDRESS_TOWN,
                                                   NAME_PERMANENT_ADDRESS_COUNTRY)
select rownum as skp_person_address,
       jez.id_person,
       decode(dpa.NAME_REPORTS_ADDRESS_STREET,
              'XNA',
              dpa.NAME_REPORTS_ADDRESS_LOCALITY,
              dpa.NAME_REPORTS_ADDRESS_STREET) as NAME_REPORTS_ADDRESS_STREET,
       dpa.TEXT_REPORTS_ADDRESS_NUM,
       dpa.CODE_REPORTS_ADDRESS_ZIP,
       dpa.NAME_REPORTS_ADDRESS_TOWN,
       dpa.NAME_REPORTS_ADDRESS_COUNTRY,
       decode(dpa.NAME_PERMANENT_ADDRESS_STREET,
              'XNA',
              dpa.NAME_PERMANENT_ADDRESS_LOCALIT,
              dpa.NAME_PERMANENT_ADDRESS_STREET) as NAME_PERMANENT_ADDRESS_STREET,
       dpa.TEXT_PERMANENT_ADDRESS_NUM,
       dpa.CODE_PERMANENT_ADDRESS_ZIP,
       dpa.NAME_PERMANENT_ADDRESS_TOWN,
       dpa.NAME_PERMANENT_ADDRESS_COUNTRY

  from owner_dwh.DC_PERSON_ADDRESS dpa
  join APP_COLL_JEZEVCIK_TEST.person jez
    on dpa.SKP_PERSON = jez.skp_person
   and dpa.flag_deleted = 'N'
   and jez.flag_valid = 'Y';

commit;





COMMIT;
--MLI nasazenÌ 14.12.2016 RISK-395
--execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.person_demography drop storage'; --RISK-468
DELETE FROM APP_COLL_JEZEVCIK_TEST.person_demography a WHERE a.id_person IN (SELECT id_person FROM person p WHERE p.skp_person > 0);

insert into APP_COLL_JEZEVCIK_TEST.PERSON_DEMOGRAPHY (SKP_PERSON_DEMOGRAPHY,
                                                      ID_PERSON,
                                                      GENDER,
                                                      DATE_BIRTH,
                                                      TEXT_TITLE_BEFORE,
                                                      TEXT_TITLE_AFTER,
                                                      NATIONALITY,
                                                      TEXT_BIRTH_NUMBER)
    select rownum as skp_person_demography,
   -- p.CUID,
           p.ID_person,
           dpd.code_gender,
           d.DATE_BIRTH,
           d.TEXT_TITLE_BEFORE,
           d.TEXT_TITLE_AFTER,
           dpd.NAME_NATIONALITY,
           decode(d.TEXT_BIRTH_NUMBER,'*','XNA',d.TEXT_BIRTH_NUMBER) as TEXT_BIRTH_NUMBER -- RISK-517
          
      from owner_dwh.dc_person d
      join owner_dwh.dc_person_demography dpd
        on dpd.SKP_PERSON_DEMOGRAPHY = d.skp_person_demography
      join APP_COLL_JEZEVCIK_TEST.person p
        on p.skp_person = d.skp_person and p.flag_valid = 'Y'
     
        
;commit;


begin
        dbms_stats.gather_table_stats(ownname => nazev_aktualniho_hriste,tabname => 'person_address',
                                                  estimate_percent => dbms_stats.auto_sample_size,
                                                  cascade => true,
                                                  no_invalidate => false);
end;


/* RISK-403 */
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.person_contact drop storage';

insert into APP_COLL_JEZEVCIK_TEST.person_contact (SKP_PERSON_CONTACT,
                                                   ID_PERSON,
                                                   PRIMARY_PHONE,
                                                   PRIMARY_EMAIL)
select rownum as skp_person_contact,a.id_person, ph.text_person_contact as PRIMARY_PHONE,  pem.text_person_contact as PRIMARY_EMAIL
-- select count(*), count(distinct a.skp_person)
from APP_COLL_JEZEVCIK_TEST.person a
left join owner_dwh.f_person_contact_tt ph on a.skp_person = ph.skp_person
                       and ph.skp_contact_type = 3 -- »Ìslo mobilnÌho telefonu v »R (prim·rnÌ telefon)
                       and ph.code_contact_status = 'a'
                       and ph.flag_current = 'Y'
left join owner_dwh.f_person_contact_tt pem on a.skp_person = pem.skp_person
                       and pem.skp_contact_type = 1 -- E-mailov· adresa (prim·rnÌ e-mail pro ofici·lnÌ komunikaci s klientem)
                       and pem.code_contact_status = 'a'
                       and pem.flag_current = 'Y' 
where a.flag_valid = 'Y'
;

COMMIT;


begin
        dbms_stats.gather_table_stats(ownname => nazev_aktualniho_hriste,tabname => 'person_contact',
                                                  estimate_percent => dbms_stats.auto_sample_size,
                                                  cascade => true,
                                                  no_invalidate => false);
end;


END;


----------------------------------------------------------------------------------------------------
-- Aktualizace data ˙mrtÌ klienta
----------------------------------------------------------------------------------------------------
procedure UPDATE_DATE_DEATH IS
  skp INT;
  -- kurzor pro proch·zenÌ klient˘, u kter˝ch je pot¯eba zmÏnit datum ˙mrtÌ
  CURSOR PERSON IS
  SELECT a.id_person, a.date_death FROM APP_COLL_JEZEVCIK_TEST.update_help_date_death a;
BEGIN

  -- vloûenÌ klient˘ a jejich nov˝mi daty ˙mrtÌ
  insert into APP_COLL_JEZEVCIK_TEST.update_help_date_death (ID_PERSON, DATE_DEATH)
  select distinct
         g.id_person,
         nullif(main.date_death, to_date('01.01.3000','dd.mm.yyyy')) as date_death
    from APP_COLL_JEZEVCIK_TEST.person g
         join owner_dwh.dc_person main
              on g.skp_person = main.skp_person and g.flag_valid = 'Y'
   where nvl(main.date_death, '01.01.3000') <> nvl(g.date_death, '01.01.3000');

-- aktualizace dat v tabulce klient˘
update APP_COLL_JEZEVCIK_TEST.person p
   set p.date_death =
       (select distinct date_death
          from APP_COLL_JEZEVCIK_TEST.update_help_date_death b
          where p.id_person = b.id_person)
where exists(select 1 from update_help_date_death h where p.id_person = h.id_person);

-- vymaz·nÌ z pomocnÈ tabulky ty klienty, kte¯Ì jeötÏ nemajÌ evidovan˝ û·dn˝ stav ûaloby (nem· zaloûenou kartu klienta)
delete from APP_COLL_JEZEVCIK_TEST.update_help_date_death where id_person not in (select id_person from APP_COLL_JEZEVCIK_TEST.epr_status_data);


  FOR J IN PERSON LOOP
    -- zÌsk·nÌ dalöÌho id pro novÈ z·znamy do tabulky upozornÏnÌ
    select max(skp_notice) into skp from APP_COLL_JEZEVCIK_TEST.notice;
    if skp is null then
      skp := 1;
    else
      skp := skp + 1 ;
    END IF;

    -- vloûenÌ upozornÏnÌ
    insert into APP_COLL_JEZEVCIK_TEST.notice t (t.skp_notice, t.id_person, t.text_notice, t.dtime_notice_created, t.dtime_notice_deadline, t.code_status, t.employee)
    values (skp, J.id_person, '⁄mrtÌ klienta dne:  ' || J.DATE_DEATH, sysdate, trunc(sysdate), 'a', 'sys');

    -- zÌsk·nÌ dalöÌho id pro novÈ z·znamy do tabulky pozn·mek
    select max(skp_epr_note) into skp from APP_COLL_JEZEVCIK_TEST.epr_note;
    if skp is null then
      skp := 1;
    else
      skp := skp + 1 ;
    END IF;

    -- vloûenÌ novÈ pozn·mky typu zmÏna ˙daj˘ klienta
    insert into APP_COLL_JEZEVCIK_TEST.epr_note (SKP_EPR_NOTE,
                                                 ID_PERSON,
                                                 SKP_EPR_NOTE_TYPE,
                                                 SKP_EPR_STATUS_DATA,
                                                 TEXT_NOTE,
                                                 DTIME_NOTE_CREATED,
                                                 EMPLOYEE)
    values (skp, j.id_person, 2, NULL, 'ZmÏna ˙daj˘ klienta: datum ˙mrtÌ (' || J.date_death || ')', sysdate, 'sys');

    skp := skp +1;

    -- vloûenÌ novÈ pozn·mky typu upozornÏnÌ
    insert into APP_COLL_JEZEVCIK_TEST.epr_note (SKP_EPR_NOTE,
                                                 ID_PERSON,
                                                 SKP_EPR_NOTE_TYPE,
                                                 SKP_EPR_STATUS_DATA,
                                                 TEXT_NOTE,
                                                 DTIME_NOTE_CREATED,
                                                 EMPLOYEE) 
    values (skp, j.id_person, 4, NULL, 'UpozornÏnÌ: ⁄mrtÌ klienta (' || J.date_death || ', deadline: ' || trunc(sysdate)  || ')', sysdate, 'sys');

  END LOOP;
  COMMIT;

begin
        dbms_stats.gather_table_stats(ownname => nazev_aktualniho_hriste,tabname => 'PERSON',
                                                  estimate_percent => dbms_stats.auto_sample_size,
                                                  cascade => true,
                                                  no_invalidate => false);
end;


execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_help_date_death drop storage';
commit;
END;




----------------------------------------------------------------------------------------------------
-- (coll_019_intervaly); interest_paid_obs; intervaly; ucet; platby; platby_ucet
----------------------------------------------------------------------------------------------------
PROCEDURE ZES_POHL IS
  
--sazba_su number := 0.0805; nepouziva se nikde
begin
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.coll_019_intervaly drop storage';
insert into APP_COLL_JEZEVCIK_TEST.coll_019_intervaly (ID_PERSON,
                                                       SKP_LOAN,
                                                       CODE_LOAN_NUMBER,
                                                       AMT_BALANCE_PRINCIPAL,
                                                       AMT_INTEREST_OVERDUE,
                                                       DATUM_ZESPL,
                                                       DTIME_VALID_FROM,
                                                       DTIME_VALID_TO,
                                                       RATE_INTEREST_STANDARD,
                                                       RATE_INTEREST_REPO,
                                                       NAME_LOAN_STATUS,
                                                       NAME_LOAN_WFSTATUS,
                                                       START_OU,
                                                       KONEC_OU,
                                                       START_SU,
                                                       KONEC_SU)
with this as (
select distinct per.id_person , a.skp_loan , c.code_loan_number , --a.amt_payoff_principal as zesplatnena_jistina ,
       a.amt_balance_principal ,
       first_value(a.amt_interest_overdue) over (partition by a.skp_loan , a.amt_balance_principal order by a.dtime_valid_from desc) as amt_interest_overdue ,
       per.date_payoff as  datum_zespl ,
       first_value(a.dtime_valid_from) over (partition by a.skp_loan , a.amt_balance_principal order by a.dtime_valid_from) as dtime_valid_from  ,
       first_value(a.dtime_valid_to) over (partition by a.skp_loan , a.amt_balance_principal order by a.dtime_valid_from desc) as dtime_valid_to ,
       c.rate_interest_standard , nvl(c.rate_interest_repo, 0.05) /*RISK-726 - u ukoncenych uverov su REPO NULLs.. toto je docasny fix. */,
       d.name_loan_status , d.name_loan_wfstatus
from APP_COLL_JEZEVCIK_TEST.person per
     join owner_dwh.f_loan_balance_at a on per.skp_person = a.skp_person and per.flag_valid = 'Y'
     join owner_dwh.f_loan_at c on a.skp_loan = c.skp_loan and a.skp_person = c.skp_person
     join owner_dwh.cl_loan_status d on c.skp_loan_status = d.skp_loan_status
                                           and d.flag_deleted <> 'Y'
where /*a.skp_person = &sp
      and */ d.code_loan_status in ('1','2')
      and a.dtime_valid_from < a.dtime_valid_to
      and (trunc(a.dtime_valid_from) >= trunc(c.date_payoff) or trunc(c.date_payoff) between a.dtime_valid_from and a.dtime_valid_to)
--order by per.skp_person , a.skp_loan , dtime_valid_from
)
select apl.* ,
       case when apl.datum_zespl > apl.dtime_valid_from then trunc(apl.datum_zespl) else trunc(apl.dtime_valid_from) end as start_ou ,
         case when apl.dtime_valid_to > sysdate then trunc(sysdate+1) else trunc(apl.dtime_valid_to) +1  end as konec_ou ,
       case when apl.datum_zespl +10 > apl.dtime_valid_from then trunc(apl.datum_zespl) + 10 else trunc(apl.dtime_valid_from) end as start_su ,
         case when apl.dtime_valid_to > sysdate then trunc(sysdate+1) else trunc(apl.dtime_valid_to) +1  end as konec_su
from this apl
;
begin  dbms_stats.gather_table_stats(ownname => nazev_aktualniho_hriste,tabname => 'coll_019_intervaly',
   estimate_percent => dbms_stats.auto_sample_size,  cascade => true, no_invalidate => false); end;
   

execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.interest_paid_obs drop storage';
insert into interest_paid_obs (typ_uroku,
                               amt_paid,
                               skp_loan)
with
transakce_pujcky as
 (select fr.*
    from owner_dwh.f_realized_transaction_tt fr
   where skp_loan in (select skp_loan
                        from owner_dwh.f_loan_at
                       where date_payoff < sysdate)),
transakce_pujcky2 as
 (select case
           when skp_transaction_type in (71, 110736, 110738) then
            'su'
           else
            'ou'
         end as typ_uroku,
         r.skp_loan,
         sum(r.amt_transaction) as amt_paid,
         trunc(dtime_transaction_generation) date_payment
    from transakce_pujcky r
   where ((r.skp_transaction_type in (110736, 110738) and
         dtime_transaction_generation < '26.5.2016') --110738 - Platba odepsanÈho sankËnÌho ˙roku mimorozvaha, 110736 - Platba sankËnÌho ˙roku mimorozvaha (podle toho, zda je odepsany nebo ne)
         or (r.skp_transaction_type = 71 and
         dtime_transaction_generation >= '26.5.2016') --Sankce za neplacenÌ ˙vÏru - po R8.5 doölo ke zmene uctovani na technickych uctech
         or r.skp_transaction_type = 110745)
     and flag_transaction_debit = 'Y'
   group by skp_loan,
            trunc(dtime_transaction_generation),
            case
              when skp_transaction_type in (71, 110736, 110738) then
               'su'
              else
               'ou'
            end)

select typ_uroku, sum(amt_paid) as amt_paid, skp_loan
  from transakce_pujcky2
 group by typ_uroku, skp_loan

;
commit;


begin  dbms_stats.gather_table_stats(ownname => nazev_aktualniho_hriste,tabname => 'interest_paid_obs',
   estimate_percent => dbms_stats.auto_sample_size,  cascade => true, no_invalidate => false); end;
   
commit;

execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.intervaly drop storage');
insert into APP_COLL_JEZEVCIK_TEST.intervaly (ID_PERSON,
                                              SKP_LOAN,
                                              CODE_LOAN_NUMBER,
                                              AMT_BALANCE_PRINCIPAL,
                                              AMT_INTEREST_OVERDUE,
                                              DATUM_ZESPL,
                                              DTIME_VALID_FROM,
                                              DTIME_VALID_TO,
                                              RATE_INTEREST_STANDARD,
                                              RATE_INTEREST_REPO,
                                              NAME_LOAN_STATUS,
                                              NAME_LOAN_WFSTATUS,
                                              START_OU,
                                              KONEC_OU,
                                              START_SU,
                                              KONEC_SU,
                                              POCET_DNI_OU,
                                              POCET_DNI_SU,
                                              OU,
                                              SU,
                                              START_SU2,
                                              KONEC_SU2,
                                              OU_TOTAL,
                                              SU_TOTAL)
select a.id_person,
       a.skp_loan,
       a.code_loan_number,
       a.amt_balance_principal,
       a.amt_interest_overdue,
       a.datum_zespl,
       a.dtime_valid_from,
       a.dtime_valid_to,
       a.rate_interest_standard,
       a.rate_interest_repo,
       a.name_loan_status,
       a.name_loan_wfstatus,
       a.start_ou,
       a.konec_ou,
       a.start_su,
       a.konec_su ,
       greatest(a.konec_ou - a.start_ou , 0) as pocet_dni_ou ,
       greatest(a.konec_su - a.start_su , 0) as  pocet_dni_su,
       round(greatest(a.konec_ou - a.start_ou , 0)/365 * a.amt_balance_principal * ( a.rate_interest_standard / 100) , 2 ) as ou ,
       round(greatest(a.konec_su - a.start_su , 0) / 365 * a.amt_balance_principal * (a.rate_interest_repo + r.penalty_interest_rate /*RISK-726*/) /100 /*RISK-726*/, 2) as  su ,
       case when start_su > konec_su then null else start_su end as start_su2 ,
         case when start_su > konec_su then null else konec_su end as konec_su2 ,
       round(sum(greatest(a.konec_ou - a.start_ou , 0)/365 * a.amt_balance_principal * ( a.rate_interest_standard / 100) ) over (partition by a.skp_loan ) - nvl(ou.amt_paid,0),2) as ou_total ,
       round(sum(greatest(a.konec_su - a.start_su , 0) / 365 * a.amt_balance_principal * (a.rate_interest_repo + r.penalty_interest_rate /*RISK-726*/) /100 /*RISK-726*/) over (partition by a.skp_loan ) - nvl(su.amt_paid,0),2) as su_total
from APP_COLL_JEZEVCIK_TEST.coll_019_intervaly a
join penalty_interest_rate r on r.dtime_valid_from <= a.datum_zespl and r.dtime_valid_to >= a.datum_zespl --pridani 27.6.2016 pro rozliseni sankcni urokove sazby
left join interest_paid_obs su on su.skp_loan = a.skp_loan and su.typ_uroku = 'su'
left join interest_paid_obs ou on ou.skp_loan = a.skp_loan and ou.typ_uroku = 'ou'
--28.6.2016 pridan
--order by a.skp_person , a.skp_loan , a.dtime_valid_from
;
/*
delete from APP_COLL_JEZEVCIK_TEST.urok;
insert into APP_COLL_JEZEVCIK_TEST.urok
with this as (
select a.skp_person  , a.skp_loan  , b.code_loan_number as uver_cislo , a.date_instalment as datum_splatky , a.date_pay_instalment as splatenie_splatky,
       a.date_pay_last as posledni_platba, a.amt_instalment,
       a.amt_instalment_principal , a.amt_instalment_interest ,
       a.amt_payment , a.amt_payment_principal , a.amt_payment_interest
from APP_COLL_JEZEVCIK_TEST.person per
     join owner_dwh.f_loan_at b on b.skp_person = per.skp_person
     join owner_dwh.f_instalment_head_ad a on a.skp_loan = b.skp_loan and a.skp_person = b.skp_person
     left join owner_dwh.cl_loan_status d on b.skp_loan_status = d.skp_loan_status
where --a.skp_person   = 1710
      code_status = 'a'
      and num_loan_schedule_order = 0
      and a.flag_deleted <> 'Y'
      and b.flag_deleted <> 'Y'
      and d.code_loan_status in ('1','2')
--order by per.skp_person , a.skp_loan , date_instalment
)
select a.skp_person , a.skp_loan , a.uver_cislo  , sum(nvl(a.amt_instalment_interest,0) - nvl(a.amt_payment_interest,0)) as kapitalizovnay_urok
from this a
group by a.skp_person , a.skp_loan , a.uver_cislo
*/


execute immediate 'truncate table  APP_COLL_JEZEVCIK_TEST.ucet drop storage';
insert into APP_COLL_JEZEVCIK_TEST.ucet (ID_PERSON,
                                         SKP_BANK_ACCOUNT,
                                         CISLO_UCTU,
                                         TYP_UCTU,
                                         DATE_PERIOD,
                                         AMT_BALANCE,
                                         CODE_CURRENCY,
                                         DTIME_MODIFIED,
                                         NAME_BANK_ACCOUNT_STATUS,
                                         NAME_BANK_ACCOUNT_WFSTATUS,
                                         AMT_GOVERNMENT_HOLD_SUM,
                                         CNT_DAYS_OVERDUE_TOL_A,
                                         CNT_DAYS_OVERDUE_TOL_B,
                                         AMT_BALANCE_AVAILABLE,
                                         AMT_BALANCE_UCETNI,
                                         DATE_BANK_ACCOUNT_ACTIVATION)
with ucet_522 as 
( --nove pridano, abychom evidovali vyvedeny dluh u debetaru s insolvenci na uctu 30300522, upraveno pri RISK-303
  select a.skp_person, sum( case when r.FLAG_TRANSACTION_DEBIT = 'N' then r.AMT_TRANSACTION_REALIZED else - r.AMT_TRANSACTION_REALIZED  end) as amt_balance, r.SKP_BANK_ACCOUNT
  from owner_dwh.f_required_transaction_tt r 
  join owner_dwh.f_contract_at a on a.TEXT_CONTRACT_NUMBER = r.TEXT_VARIABLE_SYMBOL
  where skp_bank_account = 7533428 and a.FLAG_DELETED = 'N'
  group by a.skp_person , r.SKP_BANK_ACCOUNT
)

select /*+ parallel(bb 4) */  per.ID_person , b.skp_bank_account , b.text_bank_account_number as cislo_uctu , 
        case 
          when bs.NAME_BANK_ACCOUNT_TYPE = 'BÏûn˝ ˙Ëet' then 'B⁄'
          when bs.NAME_BANK_ACCOUNT_TYPE = 'Spo¯icÌ ˙Ëet' then 'S⁄'
          when bs.NAME_BANK_ACCOUNT_TYPE = 'Vnit¯nÌ ˙Ëet' then 'V⁄'
          else bs.NAME_BANK_ACCOUNT_TYPE 
        end as typ_uctu,
       trunc(sysdate - 1) as date_period ,
       bb.AMT_BALANCE_AVAILABLE + bb.AMT_GOVERNMENT_HOLD_SUM + bb.AMT_ENVELOPE as amt_balance, -- MM: 11.5.2018 - naöli jsme jednoho zem¯elÈho, co mÏl naspo¯eno na ob·lce na SU. skp_bank_account = 7072159
       c.code_currency ,
       sysdate as dtime_modified ,
       s.name_bank_account_status,
       s.name_bank_account_wfstatus,
       bb.AMT_GOVERNMENT_HOLD_SUM, --RISK-389
       b.CNT_DAYS_OVERDUE_TOL_A,
       b.CNT_DAYS_OVERDUE_TOL_B,
       bb.AMT_BALANCE_AVAILABLE,
       bb.AMT_BALANCE AS amt_balance_ucetni,
       b.DATE_BANK_ACCOUNT_ACTIVATION
from APP_COLL_JEZEVCIK_TEST.person per
     left join owner_dwh.f_bank_account_ad b on per.skp_person = b.skp_person
                                                   and b.flag_deleted <> 'Y'
     left join owner_dwh.cl_bank_account_type bs on bs.SKP_BANK_ACCOUNT_TYPE = b.SKP_BANK_ACCOUNT_TYPE
     left join owner_dwh.cl_bank_account_status s on b.skp_bank_account_status = s.skp_bank_account_status
     left join owner_dwh.f_account_balance_td bb on bb.SKP_BANK_ACCOUNT = b.SKP_BANK_ACCOUNT and bb.date_balance_valid_to > sysdate and bb.FLAG_DELETED = 'N'
     left join owner_dwh.cl_currency c on b.skp_currency = c.skp_currency
where b.id_bank_product = 1 and per.flag_valid = 'Y'

union
--pridame transakce, ktere jsou na ucte 522 a patri klientum v Jez, viz RISK-199
select per.id_person , u.skp_bank_account , '30300522' as cislo_uctu , 'V⁄' as typ_uctu ,
       trunc(sysdate)-1 ,
       u.amt_balance ,
       'CZK' ,
       sysdate as dtime_modified ,
       'AktivnÌ',
       'AktivnÌ',
        null, --RISK-389
        NULL,
        NULL,
        NULL,
        NULL,
        NULL
from APP_COLL_JEZEVCIK_TEST.person per
join ucet_522 u on u.skp_person = per.skp_person and per.flag_valid = 'Y'

union


--pridani exekucnich, insolvencnich, jpk a KTK uctu, RISK-276
select /*+ parallel(bb 4) use_hash(ba s) leading(p ba) */ p.id_person,
       ba.SKP_BANK_ACCOUNT,
       ba.text_bank_account_number,
       CASE ba.SKP_BANK_ACCOUNT_type
         WHEN 21130 THEN 'E⁄'
         when 21312 THEN 'JPK'
         WHEN 21355 THEN 'KTK'  -- RISK-1145
         ELSE 'I⁄'
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
   and bb.skp_bank_account_type in (21130, 21270, 21312, 21355) -- 21355 = Spl·tkov˝ ˙Ëet (INSTALLMENT) pro KTK  --RISK-1145
   and p.flag_valid = 'Y';




execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.platby drop storage');
insert into APP_COLL_JEZEVCIK_TEST.platby (ID_PERSON,
                                           SKP_LOAN,
                                           UVER_CISLO,
                                           DATUM_PLATBY,
                                           DATUM_SPLATKY,
                                           AMT_INSTALMENT_TOTAL,
                                           AMT_INSTALMENT,
                                           AMT_PAYMENT,
                                           NAME_INSTALMENT_LINE_TYPE,
                                           SKF_REQUIRED_TRANSACTION,
                                           SKF_REALIZED_TRANSACTION) -- RISK-977 doplnene posledni 2 sloupecky
select prso.id_person , r.skp_loan , r.code_loan_number as uver_cislo  , p.dtime_payment as datum_platby,
       p.date_instalment as datum_splatky , p.amt_instalment_total , p.amt_instalment , p.amt_payment , q.name_instalment_line_type,
       p.SKF_REQUIRED_TRANSACTION, p.SKF_REALIZED_TRANSACTION
from owner_dwh.f_loan_at r
JOIN APP_COLL_JEZEVCIK_TEST.person prso ON prso.skp_person = r.skp_person
       join owner_dwh.f_instalment_payment_ad p on r.skp_loan = p.skp_loan
                                                      and p.flag_deleted <> 'Y'
       left join owner_dwh.cl_instalment_line_type q on p.skp_instalment_line_type = q.skp_instalment_line_type
where r.skp_loan in (select distinct skp_loan from APP_COLL_JEZEVCIK_TEST.coll_019_intervaly)
      and code_status = 'a'
      and num_loan_schedule_order = 0
      --and r.date_payoff <= p.dtime_payment
;

--pridani vypisu z uctu, viz RISK-38
--tuhle tabulku nemazat!!
-- FA upraveno v ramci risk-609, pridana podminka na zhodu variabilniho symbolu s cislom ramcove smlouvy (join na f_contract_at)
insert into APP_COLL_JEZEVCIK_TEST.platby_ucet (ID_PERSON,
                                                SKP_BANK_ACCOUNT,
                                                CISLO_UCTU,
                                                DTIME_TRANSACTION_GENERATION,
                                                DTIME_TRANSACTION_ACCOUNTED,
                                                AMT_TRANSACTION,
                                                TEXT_NOTE,
                                                SKF_REALIZED_TRANSACTION,
                                                DTIME_INSERTED)
select DISTINCT /*+  no_index(@SEL$2 T@SEL$2) */ -- DISTINCT pridano v ramci RISK-746 (problem byl ze z f_contract_at sa platba parovala aj na stornovanu ramcovku (cize duplikovala)
       prso.id_person,
       p.skp_bank_account,
       p.cislo_uctu,
       t.dtime_transaction_generation,
       t.dtime_transaction_accounted,
case when t.FLAG_TRANSACTION_DEBIT = 'Y' then - t.AMT_TRANSACTION else t.AMT_TRANSACTION end as amt_transaction,
case when t.TEXT_NOTE_PART1 <> 'XNA' then tt.NAME_TRANSACTION_TYPE || ' - ' || t.TEXT_NOTE_PART1 else
tt.NAME_TRANSACTION_TYPE end as text_note,
                         t.skf_realized_transaction , sysdate
  from owner_dwh.f_realized_transaction_tt t
  join owner_dwh.cl_transaction_type tt
    on t.SKP_TRANSACTION_TYPE = tt.SKP_TRANSACTION_TYPE
  join owner_dwh.f_bank_account_ad b
    on b.SKP_BANK_ACCOUNT = t.SKP_BANK_ACCOUNT
  join APP_COLL_JEZEVCIK_TEST.ucet p
    on p.skp_bank_account = b.SKP_BANK_ACCOUNT
  JOIN APP_COLL_JEZEVCIK_TEST.person prso ON prso.id_person = p.id_person
  JOIN owner_dwh.f_contract_at con
    ON con.SKP_PERSON = prso.skp_person
   AND con.TEXT_CONTRACT_NUMBER = CASE WHEN p.skp_bank_account = 7533428 THEN t.TEXT_VARIABLE_SYMBOL ELSE con.TEXT_CONTRACT_NUMBER END
 left join APP_COLL_JEZEVCIK_TEST.platby_ucet pp
    on pp.skf_realized_transaction = t.SKF_REALIZED_TRANSACTION    
where TRUNC(t.DTIME_TRANSACTION_ACCOUNTED) > TRUNC(sysdate - 30)
      and pp.skf_realized_transaction is null   
--where TRUNC(t.date_effective) > trunc(sysdate -2) -- between TRUNC(sysdate - 90) and TRUNC(sysdate - 3) 
   and t.code_transaction_status not in ('25', '28')
   --and t.SKF_REALIZED_TRANSACTION not in (select u.skf_required_transaction from platby_ucet u)
   and t.FLAG_DELETED = 'N'
   and b.FLAG_DELETED = 'N' and t.DTIME_TRANSACTION_GENERATION >= TRUNC(sysdate - 60) ;
   
   
-- execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.coll_019_intervaly drop storage';
commit;
end;






----------------------------------------------------------------------------------------------------
-- Aktualizace data zesplatnÏnÌ
----------------------------------------------------------------------------------------------------
procedure UPDATE_DATE_PAYOFF IS
BEGIN

execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_help_date_payoff drop storage';


--RISK-985
  -- vloûenÌ klient˘ a jejich nov˝mi daty zesplatnÏnÌ
  -- RISK-1119 Rozsireno o klienty se zesplatnenym kontokorentem
  insert into APP_COLL_JEZEVCIK_TEST.update_help_date_payoff (ID_PERSON, DATE_PAYOFF)
  select g.id_person,
         -- vyrazuji 1.1.3000 z pocitani maxima 
         nvl( max(CASE WHEN main.DATE_PAYOFF = DATE'3000-01-01' /* nezesplatneni - nesmi se ale stat */ THEN NULL
                       ELSE main.DATE_PAYOFF
                  END)
             , DATE'3000-01-01') AS date_payoff        
  
    from APP_COLL_JEZEVCIK_TEST.person g
    join (SELECT a.SKP_PERSON, a.DATE_PAYOFF FROM APP_COLL_JEZEVCIK_TEST.sr_coll_019_zes a
          UNION
          SELECT b.SKP_PERSON, b.DATE_PAYOFF FROM APP_COLL_JEZEVCIK_TEST.OVERDRAFT_ZES b  
          ) main
      on g.skp_person = main.skp_person and g.flag_valid = 'Y'
   GROUP BY g.id_person, g.date_payoff
   HAVING nvl(max(main.date_payoff), '01.01.1111') <> nvl(g.date_payoff, '01.01.1111')
   ;
   
-- aktualizace dat v tabulce klient˘
update APP_COLL_JEZEVCIK_TEST.person p
   set p.date_payoff =
       (select distinct date_payoff
          from APP_COLL_JEZEVCIK_TEST.update_help_date_payoff b
         where p.id_person = b.id_person)
where exists(select 1 from update_help_date_payoff h where p.id_person = h.id_person);

-- na zmÏnu data nebude vytv·¯ena û·dn· pozn·mka ani upozornÏnÌ v JezevËÌkovi
begin
        dbms_stats.gather_table_stats(ownname => nazev_aktualniho_hriste,tabname => 'PERSON',
                                                  estimate_percent => dbms_stats.auto_sample_size,
                                                  cascade => true,
                                                  no_invalidate => false);
end;



commit;
END;


----------------------------------------------------------------------------------------------------
--k novÏ dotaûen˝m klient˘m v JEZ p¯id·me info o soudu
----------------------------------------------------------------------------------------------------
procedure update_soud is

BEGIN

-- RISK-1098 odtial
  
/*
 insert into APP_COLL_JEZEVCIK_TEST.person_soud (SKP_PERSON_SOUD,
                                                 ID_PERSON,
                                                 PSC,
                                                 OKRESNI_SOUD,
                                                 KRAJSKY_SOUD,
                                                 ODVOLACI_SOUD,
                                                 FLAG_VALID,
                                                 DTIME_MODIFICATION,
                                                 EMPLOYEE)
    with this as
     (select distinct count(distinct okresni_soud) over(partition by psc) as co,
                      psc,
                      okresni_soud,
                      krajsky_soud,
                      nazev_obce
        from APP_COLL_JEZEVCIK_TEST.soudy),
    this2 as
     (select distinct p.id_person,
                      --b.code_gender,
                      so.psc,
                      so.okresni_soud,
                      so.krajsky_soud,
                      NULL as odvolaci_soud,
                      'Y' as flag_valid,
                      sysdate as dtime_modification,
                      'sys' as employee
        from APP_COLL_JEZEVCIK_TEST.person p
        left join person_soud ps
          on p.id_person = ps.id_person
        join APP_COLL_JEZEVCIK_TEST.person_address d
          on d.id_person = p.id_person
        left join this so
          on (so.psc = d.code_permanent_address_zip and
             d.name_permanent_address_country = 'Czech Republic' and
             so.co = 1) --pokud k jednomu psc pouze jeden soud, parujeme podle psc (rozliseni na cizince vs. domaci - trvala vs. prechodna adresa)
          or (so.psc = d.code_reports_address_zip and
             d.name_permanent_address_country <> 'Czech Republic' and
             so.co = 1)
          or (so.psc = d.code_permanent_address_zip and
             so.nazev_obce = d.name_permanent_address_town and
             d.name_permanent_address_country = 'Czech Republic' and
             so.co = 2) -- pokud vicero soudu k psc, hledame podle jmena
          or (so.psc = d.code_reports_address_zip and
             so.nazev_obce = d.name_reports_address_town and
             d.name_permanent_address_country <> 'Czech Republic' and
             so.co = 2)
       where ps.id_person is null and p.flag_valid = 'Y'-- pridame klienty, kteri tam jeste nejsou
      )
    --zde jeste pridame skp_person_soud
    select (select max(skp_person_soud) from APP_COLL_JEZEVCIK_TEST.person_soud) +
           rownum as skp_person_soud,
           t.*
      from this2 t;
  commit;*/
  
  /*
  -- update pry nechceme, dle komentare v RISK-418
  <<update_existujicich_osob>>
  BEGIN
    
      UPDATE app_coll_jezevcik_test.person_soud a
           SET a.DTIME_MODIFICATION = SYSDATE
             , a.EMPLOYEE = 'sys'
             , a.SKP_SOUD = ( SELECT sp.SKP_SOUD from APP_COLL_JEZEVCIK_TEST.person_address d
                                                 join app_coll_jezevcik_test.SOUD_PAIR sp
                                                    on (sp.psc = d.code_permanent_address_zip and
                                                       d.name_permanent_address_country = 'Czech Republic' and
                                                       sp.POCET_SOUDU_K_PSC = 1) --pokud k jednomu psc pouze jeden soud, parujeme podle psc (rozliseni na cizince vs. domaci - trvala vs. prechodna adresa)
                                                    or (sp.psc = d.code_reports_address_zip and
                                                       d.name_permanent_address_country <> 'Czech Republic' and
                                                       sp.POCET_SOUDU_K_PSC = 1)
                                                    or (sp.psc = d.code_permanent_address_zip and
                                                       sp.nazev_obce = d.name_permanent_address_town and
                                                       d.name_permanent_address_country = 'Czech Republic' and
                                                       sp.POCET_SOUDU_K_PSC = 2) -- pokud vicero soudu k psc, hledame podle jmena
                                                    or (sp.psc = d.code_reports_address_zip and
                                                       sp.nazev_obce = d.name_reports_address_town and
                                                       d.name_permanent_address_country <> 'Czech Republic' and
                                                       sp.POCET_SOUDU_K_PSC = 2)
                              WHERE d.ID_PERSON = a.ID_PERSON
                             )
         WHERE EXISTS (SELECT sp.SKP_SOUD from APP_COLL_JEZEVCIK_TEST.person_address d
                                               join app_coll_jezevcik_test.SOUD_PAIR sp
                                                  on (sp.psc = d.code_permanent_address_zip and
                                                     d.name_permanent_address_country = 'Czech Republic' and
                                                     sp.POCET_SOUDU_K_PSC = 1) --pokud k jednomu psc pouze jeden soud, parujeme podle psc (rozliseni na cizince vs. domaci - trvala vs. prechodna adresa)
                                                  or (sp.psc = d.code_reports_address_zip and
                                                     d.name_permanent_address_country <> 'Czech Republic' and
                                                     sp.POCET_SOUDU_K_PSC = 1)
                                                  or (sp.psc = d.code_permanent_address_zip and
                                                     sp.nazev_obce = d.name_permanent_address_town and
                                                     d.name_permanent_address_country = 'Czech Republic' and
                                                     sp.POCET_SOUDU_K_PSC = 2) -- pokud vicero soudu k psc, hledame podle jmena
                                                  or (sp.psc = d.code_reports_address_zip and
                                                     sp.nazev_obce = d.name_reports_address_town and
                                                     d.name_permanent_address_country <> 'Czech Republic' and
                                                     sp.POCET_SOUDU_K_PSC = 2)
                        WHERE d.ID_PERSON = a.ID_PERSON
                       )
             AND NOT EXISTS (SELECT spm.ID_PERSON FROM app_coll_jezevcik_test.SOUDY_PAIRED_MANUALLY spm WHERE spm.ID_PERSON = a.ID_PERSON) -- manualne naparovane klienty updatovat nechceme!
        ;
        
    END update_existujicich_osob;
    */
    
    <<vlozeni_novych_osob>>
    BEGIN
         
          MERGE INTO app_coll_jezevcik_test.person p
          USING (                       SELECT d.id_person, sp.SKP_SOUD
                                        FROM APP_COLL_JEZEVCIK_TEST.person_address d
                                        join app_coll_jezevcik_test.SOUD_PAIR sp
                                          on (-- pokud k jednomu psc pouze jeden soud, parujeme podle psc (rozliseni na cizince vs. domaci - trvala vs. prechodna adresa)
                                                 sp.psc = d.code_permanent_address_zip and
                                                 d.name_permanent_address_country = 'Czech Republic' and
                                                 sp.POCET_SOUDU_K_PSC = 1) 
                                              or (sp.psc = d.code_reports_address_zip and
                                                 d.name_permanent_address_country <> 'Czech Republic' and
                                                 sp.POCET_SOUDU_K_PSC = 1)
                                                       
                                              -- pokud vicero soudu k psc, hledame podle jmena
                                              or (sp.psc = d.code_permanent_address_zip and
                                                 sp.nazev_obce = d.name_permanent_address_town and
                                                 d.name_permanent_address_country = 'Czech Republic' and
                                                 sp.POCET_SOUDU_K_PSC = 2) 
                                              or (sp.psc = d.code_reports_address_zip and
                                                 sp.nazev_obce = d.name_reports_address_town and
                                                 d.name_permanent_address_country <> 'Czech Republic' and
                                                 sp.POCET_SOUDU_K_PSC = 2)
                   ) s
                       
                   ON (p.id_person = s.id_person)
          WHEN MATCHED THEN 
          UPDATE SET p.skp_soud = s.skp_soud   
          WHERE p.skp_soud IS NULL -- update pouze, pokud klient jeste soud prirazen nema
          ;

    
    END vlozeni_novych_osob;
  
-- RISK-1098 potial

end;




----------------------------------------------------------------------------------------------------
-- Aktualizace duvody zesplatnÏnÌ
----------------------------------------------------------------------------------------------------
procedure update_reason_zes IS
BEGIN

execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_payoff_reason drop storage';

  -- vloûenÌ klient˘ a jejich nov˝mi duvodmi zesplatnÏnÌ

--RISK-985
--RISK-1119
insert into APP_COLL_JEZEVCIK_TEST.update_payoff_reason (ID_PERSON, PAYOFF_REASON)
with
sub_reason as 
(
  select distinct  skp_person , name_payoff_reason  as payoff_reason
  FROM (SELECT skp_person , name_payoff_reason FROM APP_COLL_JEZEVCIK_TEST.sr_coll_019_zes a
        UNION
        SELECT skp_person , name_payoff_reason FROM APP_COLL_JEZEVCIK_TEST.OVERDRAFT_ZES b
        )
),
sub_reason2 as 
(
  select skp_person , listagg(payoff_reason, ', '||chr(10)) within group (order by payoff_reason) as payoff_reason
  FROM sub_reason
  group by skp_person
)
  select g.id_person,
         main.payoff_reason
    from APP_COLL_JEZEVCIK_TEST.person g
         join sub_reason2 main
              on g.skp_person = main.skp_person and g.flag_valid = 'Y'
   WHERE main.payoff_reason <> g.payoff_reason
;

-- aktualizace dat v tabulce klient˘
update APP_COLL_JEZEVCIK_TEST.person p
   set p.payoff_reason =
       (select distinct payoff_reason
          from APP_COLL_JEZEVCIK_TEST.update_payoff_reason b
         where p.id_person = b.id_person)
where exists(select 1 from update_payoff_reason h where p.id_person = h.id_person);

-- na zmÏnu data nebude vytv·¯ena û·dn· pozn·mka ani upozornÏnÌ v JezevËÌkovi
begin
        dbms_stats.gather_table_stats(ownname => nazev_aktualniho_hriste,tabname => 'PERSON',
                                                  estimate_percent => dbms_stats.auto_sample_size,
                                                  cascade => true,
                                                  no_invalidate => false);
end;



commit;
END;



----------------------------------------------------------------------------------------------------
-- Kontrola zaplacenÌ soudnÌch poplatku klientem.
------------------------------------------------------------------------------------
procedure update_sop_paid is

skp INT;
skp2 int;

-- kurzor
CURSOR PERSON IS
SELECT a.id_person,
       a.flag_paid,
       a.flag_in_sop,
       a.amt_paid 
 FROM APP_COLL_JEZEVCIK_TEST.update_help_sop_paid a
;


CURSOR PERSON2 IS
select a.id_person,
       a.flag_paid,
       a.flag_in_sop,
       a.amt_paid
from APP_COLL_JEZEVCIK_TEST.update_help_sop_paid a
where flag_in_sop = 0
;

BEGIN

execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.platby_SOP drop storage');
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_help_sop_paid drop storage';

-- klienti se zaplacenym soudnim poplatkem
-- flag_paid : platba na sedmiËkovom ˙Ëte
-- flag_in_sop : m· z·znam v tabulke SOPu



INSERT INTO APP_COLL_JEZEVCIK_TEST.PLATBY_SOP (ID_PERSON,
                                               AMT_TRANSACTION_REALIZED,
                                               FLAG_TRANSACTION_DEBIT,
                                               SKP_BANK_ACCOUNT,
                                               DATE_EFFECTIVE,
                                               DATE_TRANSACTION_DUE,
                                               DATE_TRANS_REALIZATION,
                                               DTIME_TRANS_STATUS_VALID_FROM,
                                               SKP_BANK_ACCOUNT_OPPOSITE,
                                               ID_TRANSACTION_UNIQUE_IDENT,
                                               SKF_REQUIRED_TRANSACTION)
(
    --stare
    select p.ID_PERSON --, sum(rt.amt_transaction_realized) as amt_paid
         , rt.AMT_TRANSACTION_REALIZED
         , rt.FLAG_TRANSACTION_DEBIT
         , rt.SKP_BANK_ACCOUNT
         , rt.DATE_EFFECTIVE
         , rt.DATE_TRANSACTION_DUE
         , rt.DATE_TRANS_REALIZATION
         , rt.DTIME_TRANS_STATUS_VALID_FROM
         , rt.SKP_BANK_ACCOUNT_OPPOSITE
         , rt.ID_TRANSACTION_UNIQUE_IDENT
         , rt.SKF_REQUIRED_TRANSACTION
    FROM owner_dwh.f_required_transaction_tt rt
        join owner_dwh.f_bank_account_ad ba 
          on rt.skp_bank_account_opposite = ba.skp_bank_account
        JOIN APP_COLL_JEZEVCIK_TEST.person p
          ON ba.skp_person = p.SKP_PERSON
    where rt.skp_bank_account = 4977766 -- t.j. 30300506
          and rt.flag_transaction_debit = 'N'
          and rt.flag_deleted <> 'Y'
          and ba.flag_deleted <> 'Y'
          and rt.code_req_transact_status not in ('5')
          and rt.DATE_TRANS_REALIZATION < DATE'2016-09-07'
          and rt.skp_bank_account_opposite <> 379922  --!!! RISK-427
    --group by ba.skp_person

    union all

    --nove, viz risk-306
    select p.ID_PERSON --, sum(case when rt.FLAG_TRANSACTION_DEBIT = 'Y' then - rt.amt_transaction_realized else rt.amt_transaction_realized end) as amt_paid
         , rt.AMT_TRANSACTION_REALIZED
         , rt.FLAG_TRANSACTION_DEBIT
         , rt.SKP_BANK_ACCOUNT
         , rt.DATE_EFFECTIVE
         , rt.DATE_TRANSACTION_DUE
         , rt.DATE_TRANS_REALIZATION
         , rt.DTIME_TRANS_STATUS_VALID_FROM
         , rt.SKP_BANK_ACCOUNT_OPPOSITE
         , rt.ID_TRANSACTION_UNIQUE_IDENT
         , rt.SKF_REQUIRED_TRANSACTION
    FROM owner_dwh.f_required_transaction_tt rt
        join owner_dwh.f_bank_account_ad ba 
          ON rt.TEXT_VARIABLE_SYMBOL = ba.TEXT_BANK_ACCOUNT_NUMBER
        JOIN APP_COLL_JEZEVCIK_TEST.person p
          ON ba.skp_person = p.SKP_PERSON
    where rt.skp_bank_account = 4977766 -- t.j. 30300506
          and (rt.flag_transaction_debit = 'N' or text_note_part1 like '%Vr·cenÌ smluvnÌch ˙rok˘%'  or text_note_part1 like '%Vr·cenÌ z·konn˝ch ˙rok˘%')
          and lower(rt.TEXT_NOTE_PART1) not like '%od˙Ëtov·nÌ p¯eplatku%' -- RISK-615 - kv˘li chybnÏ za˙ËtovanÈ transakci a n·slednÈmu od˙Ëtov·nÌ, odebÌr·me z evidence zaplacen˝ch poplatk˘ - nenÌ re·lnÏ zaplaceno
          and rt.flag_deleted <> 'Y'
          and ba.flag_deleted <> 'Y'
          -- and rt.code_req_transact_status not in ('5') --FA odkomentovano v souvislosti s klientem Centovsky(cuid 511195), 9.6.2017
          and rt.DATE_TRANS_REALIZATION >= DATE'2016-09-07'
          and rt.skp_bank_account_opposite <> 379922  --!!! RISK-427
          AND rt.ID_TRANSACTION_UNIQUE_IDENT != 457509672 -- TRN byla stornovana opacnou transakci, 15.9.2017, kl. Nov·k -- ak by sa opakovalo, mozno nastavit na text_note_part1
    --group by ba.skp_person
);




insert into APP_COLL_JEZEVCIK_TEST.update_help_sop_paid (ID_PERSON,
                                                         FLAG_PAID,
                                                         FLAG_IN_SOP,
                                                         AMT_PAID)
WITH this as 
  (
    select id_person , sum(case when ps.FLAG_TRANSACTION_DEBIT = 'Y' then - ps.amt_transaction_realized else ps.amt_transaction_realized end) as amt_paid
    from APP_COLL_JEZEVCIK_TEST.PLATBY_SOP ps
    group by id_person 
  )

/*  !!! -- vypocet z vynosoveho uctu. Uû sa nepouûÌva , pretoûe dost·val klientov priamo do debitu
select \* a.skp_bank_account , a.text_bank_account_number ,
       b.skf_realized_transaction , b.dtime_transaction_accounted ,
       b.amt_transaction , b.code_currency , b.text_opposite_bank_acc_number ,
       c.skp_person , d.name_last , d.name_first , b.text_variable_symbol , d.id_cuid ,b.flag_transaction_debit*\
       d.skp_person , sum(b.amt_transaction) as amt_paid
from
owner_dwh.f_bank_account_ad a
                            join owner_dwh.f_realized_transaction_tt b on a.skp_bank_account = b.skp_bank_account
                            join owner_dwh.f_bank_account_ad c on b.skp_bank_account_opposite = c.skp_bank_account
                            join owner_dwh.dc_person d on c.skp_person = d.skp_person
where a.skp_bank_account = 1629690 --a.text_bank_account_number = '070009040501'
      and a.flag_deleted <> 'Y'
      and b.flag_deleted <> 'Y'
      and c.flag_deleted <> 'Y'
      and d.flag_deleted <> 'Y'
      and b.flag_transaction_debit = 'N'
group by d.skp_person
*/



select per.id_person , nvl2(t.id_person,1,0) as flag_paid , nvl2(s.id_person,1,0) as flag_in_sop  , nvl(t.amt_paid,0) amt_paid
from APP_COLL_JEZEVCIK_TEST.person per
     left join this t on t.id_person = per.id_person
     left join APP_COLL_JEZEVCIK_TEST.sop s on per.id_person = s.id_person
where (t.amt_paid <> nvl(s.sop_paid, 0)                                                       --!!! RISK-427
      or (nvl(t.amt_paid, 0) <> nvl(s.sop_paid, 0) and nvl(s.sop_paid, 0)>nvl(t.amt_paid, 0)) --!!! RISK-427 --tato podmÌnka je tam proto, kdyby z nÏjakÈho d˘vodu byly novÏ napoËÌtanÈ zaplacenÈ SOP < nez zaplacenÈ SOP v tabulce SOP. nap¯. z d˘vodu, kdy se urËitÈ platby do SOP p¯estaly zapoËÌt·vat.
      )                                                                                       --!!! RISK-427      
      and per.flag_valid = 'Y' 
;


-- aktualizace dat v tabulke SOP
--pokud klient ma zaznam v tabulce, nevkladame radek, ale pouze update hodnot
update APP_COLL_JEZEVCIK_TEST.sop s
set s.sop_paid =
    (select b.amt_paid
     from APP_COLL_JEZEVCIK_TEST.update_help_sop_paid b
     where s.id_person = b.id_person
     ),
     s.dtime_modification = sysdate ,
     s.employee = 'sys'
where exists(select 1 from APP_COLL_JEZEVCIK_TEST.update_help_sop_paid c where c.id_person = s.id_person and c.flag_in_sop = 1)
;

--zde vkladame nove zaznamy
FOR J IN PERSON2 LOOP
    -- zÌsk·nÌ dalöÌho id pro novÈ z·znamy do tabulky upozornÏnÌ
    select max(skp_sop) into skp2 from APP_COLL_JEZEVCIK_TEST.sop;
    if skp2 is null then
      skp2 := 1;
    else
      skp2 := skp2 + 1 ;
    END IF;

    -- vloûenÌ dat do tabulky SOP
    insert into APP_COLL_JEZEVCIK_TEST.sop (skp_sop , id_person , dtime_modification , employee , sop_paid)
    values (skp2, J.id_person , sysdate , 'sys' , J.amt_paid)
    ;

end loop;


 --vzdy vlozime upozorneni (bez ohledu na to, zda byl v tabulce nebo ne)
  FOR J IN PERSON LOOP
    -- zÌsk·nÌ dalöÌho id pro novÈ z·znamy do tabulky upozornÏnÌ
    select max(skp_notice) into skp from APP_COLL_JEZEVCIK_TEST.notice;
    if skp is null then
      skp := 1;
    else
      skp := skp + 1 ;
    END IF;

    -- vloûenÌ upozornÏnÌ
    insert into APP_COLL_JEZEVCIK_TEST.notice t (t.skp_notice, t.id_person, t.text_notice, t.dtime_notice_created, t.dtime_notice_deadline, t.code_status, t.employee)
    values (skp, J.id_person, 'Suma zaplacen˝ch soudnÌch poplatk˘: ' || J.amt_paid || ' KË', sysdate, trunc(sysdate), 'a', 'sys');

    -- zÌsk·nÌ dalöÌho id pro novÈ z·znamy do tabulky pozn·mek
    select max(skp_epr_note) into skp from APP_COLL_JEZEVCIK_TEST.epr_note;
    if skp is null then
      skp := 1;
    else
      skp := skp + 1 ;
    END IF;

    -- vloûenÌ novÈ pozn·mky typu zmÏna ˙daj˘ klienta
    insert into APP_COLL_JEZEVCIK_TEST.epr_note (SKP_EPR_NOTE,
                                                 ID_PERSON,
                                                 SKP_EPR_NOTE_TYPE,
                                                 SKP_EPR_STATUS_DATA,
                                                 TEXT_NOTE,
                                                 DTIME_NOTE_CREATED,
                                                 EMPLOYEE) 
    values (skp, j.id_person, 5, NULL, 'ZmÏna ˙daj˘ SOP: zaplacenÌ SOP (' || J.amt_paid|| ' KË)', sysdate, 'sys');

    skp := skp +1;

    -- vloûenÌ novÈ pozn·mky typu upozornÏnÌ
    insert into APP_COLL_JEZEVCIK_TEST.epr_note (SKP_EPR_NOTE,
                                                 ID_PERSON,
                                                 SKP_EPR_NOTE_TYPE,
                                                 SKP_EPR_STATUS_DATA,
                                                 TEXT_NOTE,
                                                 DTIME_NOTE_CREATED,
                                                 EMPLOYEE)
    values (skp, j.id_person, 4, NULL, 'UpozornÏnÌ: Zaplacen˝ SOP (' || J.amt_paid || ' KË, deadline: ' || trunc(sysdate)  || ')', sysdate, 'sys');

  END LOOP;

begin
        dbms_stats.gather_table_stats(ownname => nazev_aktualniho_hriste,tabname => 'SOP',
                                                  estimate_percent => dbms_stats.auto_sample_size,
                                                  cascade => true,
                                                  no_invalidate => false);
end;


execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_help_sop_paid drop storage';

commit;
end;


----------------------------------------------------------------------------------------------------
-- Aktualizace dat klienta (doplaceni jistiny/uroku) - flagy v tabulce PERSON
---------------------------------------------------------------------------------------------------
procedure update_paid_claim IS
  skp INT;
  -- kurzor pro proch·zenÌ klient˘
  CURSOR PERSON IS
  SELECT a.id_person FROM APP_COLL_JEZEVCIK_TEST.update_paid_principal a;


  CURSOR PERSON2 IS
  SELECT a.id_person FROM APP_COLL_JEZEVCIK_TEST.update_paid_interest a;
  
  CURSOR PERSON3 IS
  SELECT a.id_person FROM APP_COLL_JEZEVCIK_TEST.update_paid_DEBET a; -- PRIDANI V RAMCI RISK-254
    
BEGIN


-----------------------------------------------------------
-- Vytvorenie tabuliek s klientami ktorym sa zmeni flag doplatenia jistiny alebo uroku

execute immediate 'truncate table update_paid_principal drop storage';
execute immediate 'truncate table update_paid_interest drop storage';
execute immediate 'truncate table update_paid_debet drop storage';

--RISK-985   DEAL jenom na PROD
insert into update_paid_principal (id_person)
(
select b.ID_person
from APP_COLL_JEZEVCIK_TEST.loan b
     join owner_dwh.f_loan_balance_at c on b.skp_loan = c.skp_loan
where c.flag_valid = 'Y'
      and c.flag_deleted <> 'Y'
group BY b.id_person
having sign(sum(c.amt_payoff_principal)) = 0
-- MM: Oprava 31.7.2017 - HOTFIX chyby v f_loan_balance_at -- having sign(sum(c.AMT_BALANCE_PRINCIPAL)) = 0
-- FA: vratene na povodny skript 29.8.2017

intersect

select id_person
from person p
where p.flag_principal_paid_obs = 'N' and p.flag_valid = 'Y'
group by id_person
)
;



--RISK-985   DEAL jenom na PROD
insert into update_paid_interest (id_person)
(
select b.id_person
from APP_COLL_JEZEVCIK_TEST.loan b
     join owner_dwh.f_loan_balance_at c on b.skp_loan = c.skp_loan
where c.flag_valid = 'Y'
      and c.flag_deleted <> 'Y'
group by b.id_person
having sign(sum(c.amt_payoff_interest + c.amt_payoff_penalty)) = 0

intersect

select id_person
from person p
where p.flag_interest_paid_obs = 'N' and p.flag_valid = 'Y'
group by id_person
)
;

       
insert into  update_paid_debet (id_person) --risk 254, pridani upozorneni na doplaceni debetaru
    select distinct e.id_person
      from person_status_data e
      join (select sum(amt_balance) amt_min, id_person
              from ucet
             group by id_person
            having sum(amt_balance) >= 0) a
        on a.id_person = e.id_person
     left join person p on e.id_person = p.id_person -- and p.code_insolvency_case_status in ('KONKURS','⁄PADEK','ODDLUéENÕ') -- RISK-728
     where e.skp_person_status = 3
       and e.flag_deleted = 'N'
       and e.flag_collected = 'Y'
       and e.flag_valid = 'Y'
       AND p.flag_valid != 'N' -- AND (P.skp_person IS NULL or p.flag_valid = 'N')/* XXXXX - bud tam neni, nebo je nevalidni*/  RISK-728
;

-----------------------------------------------------------------
-- Update flagov urcujucich doplaceni
update person
set flag_principal_paid_obs = 'Y'
where id_person in (select a.id_person from update_paid_principal a)
;


update person
set flag_interest_paid_obs = 'Y'
where id_person in (select a.id_person from update_paid_interest a)
;


-- Cyklus plneni poznamek a upozorneni
---------------------------------------------------
  FOR J IN PERSON LOOP
    -- zÌsk·nÌ dalöÌho id pro novÈ z·znamy do tabulky upozornÏnÌ
    select max(skp_notice) into skp from APP_COLL_JEZEVCIK_TEST.notice;
    if skp is null then
      skp := 1;
    else
      skp := skp + 1 ;
    END IF;

    -- vloûenÌ upozornÏnÌ
    insert into APP_COLL_JEZEVCIK_TEST.notice t (t.skp_notice, t.id_person, t.text_notice, t.dtime_notice_created, t.dtime_notice_deadline, t.code_status, t.employee)
    values (skp, J.id_person, 'DoplacenÌ jistiny (OBS)', sysdate, trunc(sysdate), 'a', 'sys');

    -- zÌsk·nÌ dalöÌho id pro novÈ z·znamy do tabulky pozn·mek
    select max(skp_epr_note) into skp from APP_COLL_JEZEVCIK_TEST.epr_note;
    if skp is null then
      skp := 1;
    else
      skp := skp + 1 ;
    END IF;

    -- vloûenÌ novÈ pozn·mky typu zmÏna ˙daj˘ klienta
    insert into APP_COLL_JEZEVCIK_TEST.epr_note (SKP_EPR_NOTE,
                                                 ID_PERSON,
                                                 SKP_EPR_NOTE_TYPE,
                                                 SKP_EPR_STATUS_DATA,
                                                 TEXT_NOTE,
                                                 DTIME_NOTE_CREATED,
                                                 EMPLOYEE) 
    values (skp, j.id_person, 2, NULL, 'ZmÏna ˙daj˘ klienta: doplacenÌ jistiny (OBS)', sysdate, 'sys');

    skp := skp +1;

    -- vloûenÌ novÈ pozn·mky typu upozornÏnÌ
    insert into APP_COLL_JEZEVCIK_TEST.epr_note (SKP_EPR_NOTE,
                                                 ID_PERSON,
                                                 SKP_EPR_NOTE_TYPE,
                                                 SKP_EPR_STATUS_DATA,
                                                 TEXT_NOTE,
                                                 DTIME_NOTE_CREATED,
                                                 EMPLOYEE) 
    values (skp, j.id_person, 4, NULL, 'UpozornÏnÌ: Doplacen· jistina podle OBS (deadline: ' || trunc(sysdate)  || ')', sysdate, 'sys');

  END LOOP;
  COMMIT;


    FOR J IN PERSON2 LOOP
    -- zÌsk·nÌ dalöÌho id pro novÈ z·znamy do tabulky upozornÏnÌ
    select max(skp_notice) into skp from APP_COLL_JEZEVCIK_TEST.notice;
    if skp is null then
      skp := 1;
    else
      skp := skp + 1 ;
    END IF;

    -- vloûenÌ upozornÏnÌ
    insert into APP_COLL_JEZEVCIK_TEST.notice t (t.skp_notice, t.id_person, t.text_notice, t.dtime_notice_created, t.dtime_notice_deadline, t.code_status, t.employee)
    values (skp, J.id_person, 'DoplacenÌ ˙rok˘ (OBS)', sysdate, trunc(sysdate), 'a', 'sys');

    -- zÌsk·nÌ dalöÌho id pro novÈ z·znamy do tabulky pozn·mek
    select max(skp_epr_note) into skp from APP_COLL_JEZEVCIK_TEST.epr_note;
    if skp is null then
      skp := 1;
    else
      skp := skp + 1 ;
    END IF;

    -- vloûenÌ novÈ pozn·mky typu zmÏna ˙daj˘ klienta
    insert into APP_COLL_JEZEVCIK_TEST.epr_note (SKP_EPR_NOTE,
                                                 ID_PERSON,
                                                 SKP_EPR_NOTE_TYPE,
                                                 SKP_EPR_STATUS_DATA,
                                                 TEXT_NOTE,
                                                 DTIME_NOTE_CREATED,
                                                 EMPLOYEE) 
    values (skp, j.id_person, 2, NULL, 'ZmÏna ˙daj˘ klienta: doplacenÌ ˙rok˘ (OBS)', sysdate, 'sys');

    skp := skp +1;

    -- vloûenÌ novÈ pozn·mky typu upozornÏnÌ
    insert into APP_COLL_JEZEVCIK_TEST.epr_note (SKP_EPR_NOTE,
                                                 ID_PERSON,
                                                 SKP_EPR_NOTE_TYPE,
                                                 SKP_EPR_STATUS_DATA,
                                                 TEXT_NOTE,
                                                 DTIME_NOTE_CREATED,
                                                 EMPLOYEE)
    values (skp, j.id_person, 4, NULL, 'UpozornÏnÌ: DoplacenÌ ˙rok˘ podle OBS (deadline: ' || trunc(sysdate)  || ')', sysdate, 'sys');

  END LOOP;
  COMMIT;

  --debetari
    FOR J IN PERSON3 LOOP
    -- zÌsk·nÌ dalöÌho id pro novÈ z·znamy do tabulky upozornÏnÌ
    select max(skp_notice) into skp from APP_COLL_JEZEVCIK_TEST.notice;
    if skp is null then
      skp := 1;
    else
      skp := skp + 1 ;
    END IF;

    -- vloûenÌ upozornÏnÌ
    insert into APP_COLL_JEZEVCIK_TEST.notice t (t.skp_notice, t.id_person, t.text_notice, t.dtime_notice_created, t.dtime_notice_deadline, t.code_status, t.employee)
    values (skp, J.id_person, 'DoplacenÌ debetu (OBS)', sysdate, trunc(sysdate), 'a', 'sys');

    -- zÌsk·nÌ dalöÌho id pro novÈ z·znamy do tabulky pozn·mek
    select max(skp_epr_note) into skp from APP_COLL_JEZEVCIK_TEST.epr_note;
    if skp is null then
      skp := 1;
    else
      skp := skp + 1 ;
    END IF;

    -- vloûenÌ novÈ pozn·mky typu zmÏna ˙daj˘ klienta
    insert into APP_COLL_JEZEVCIK_TEST.epr_note (SKP_EPR_NOTE,
                                                 ID_PERSON,
                                                 SKP_EPR_NOTE_TYPE,
                                                 SKP_EPR_STATUS_DATA,
                                                 TEXT_NOTE,
                                                 DTIME_NOTE_CREATED,
                                                 EMPLOYEE) 
    values (skp, j.id_person, 2, NULL, 'ZmÏna ˙daj˘ klienta: doplacenÌ debetu (OBS)', sysdate, 'sys');

    skp := skp +1;

    -- vloûenÌ novÈ pozn·mky typu upozornÏnÌ
    insert into APP_COLL_JEZEVCIK_TEST.epr_note (SKP_EPR_NOTE,
                                                 ID_PERSON,
                                                 SKP_EPR_NOTE_TYPE,
                                                 SKP_EPR_STATUS_DATA,
                                                 TEXT_NOTE,
                                                 DTIME_NOTE_CREATED,
                                                 EMPLOYEE) 
    values (skp, j.id_person, 4, NULL, 'UpozornÏnÌ: DoplacenÌ debetu OBS (deadline: ' || trunc(sysdate)  || ')', sysdate, 'sys');

  END LOOP;
  COMMIT;

begin
        dbms_stats.gather_table_stats(ownname => nazev_aktualniho_hriste,tabname => 'PERSON',
                                                  estimate_percent => dbms_stats.auto_sample_size,
                                                  cascade => true,
                                                  no_invalidate => false);
end;


execute immediate 'truncate table update_paid_principal drop storage';
execute immediate 'truncate table update_paid_interest drop storage';
execute immediate 'truncate table update_paid_debet drop storage';

commit;
END;


----------------------------------------------------------------------------------------------------
-- Update neû·doucnosti u klientov - FLAG v tabulce PERSON
----------------------------------------------------------------------------------------------------
procedure update_undesirability is

begin


execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_undesirability drop storage';


UPDATE person
set flag_undesirability = 'N',
    text_undesirability_note = null
;


insert into APP_COLL_JEZEVCIK_TEST.update_undesirability (ID_PERSON, TEXT_UNDESIRABILITY_NOTE)
select g.id_person,
         listagg(main.text_undesirability_note , '; ') within group (order by main.text_undesirability_note) as text_undesirability_note
from APP_COLL_JEZEVCIK_TEST.person g
     join owner_dwh.f_person_undesirability_tt main
              on g.skp_person = main.skp_person and g.flag_valid = 'Y'
where main.flag_deleted <> 'Y' and main.flag_current = 'Y' and main.code_status = 'a'
group by g.id_person
;

-- aktualizace dat v tabulce klient˘
update APP_COLL_JEZEVCIK_TEST.person p
   set p.flag_undesirability = 'Y',
       p.text_undesirability_note = (select uu.text_undesirability_note
                                     from  APP_COLL_JEZEVCIK_TEST.update_undesirability uu
                                     where p.id_person = uu.id_person )
where exists(select 1 from update_undesirability h where p.id_person = h.id_person);

-- na zmÏnu data nebude vytv·¯ena û·dn· pozn·mka ani upozornÏnÌ v JezevËÌkovi
begin
        dbms_stats.gather_table_stats(ownname => nazev_aktualniho_hriste,tabname => 'PERSON',
                                                  estimate_percent => dbms_stats.auto_sample_size,
                                                  cascade => true,
                                                  no_invalidate => false);
end;




commit;
end;



-------------------------------------
-- Aktualizace insolvenËnÌho statusu - FLAG v tabulce PERSON
-------------------------------------
procedure UPDATE_INSOL_STATUS IS
  skp INT;
  -- kurzor pro proch·zenÌ klient˘, u kter˝ch je pot¯eba zmÏnit status insolvence
  CURSOR PERSON IS
  SELECT a.id_person, a.code FROM APP_COLL_JEZEVCIK_TEST.update_help_insolvency a;
BEGIN

  -- vloûenÌ klient˘ a jejich nov˝mi statusy do pomocnÈ tabulky
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.update_help_insolvency drop storage');


--RISK-985
  insert into APP_COLL_JEZEVCIK_TEST.update_help_insolvency (ID_PERSON, CODE)
  select distinct g.id_person, main.code
  from (select distinct prso.id_person,
                        first_value(main.ic_stav_rizeni) over(partition by prso.id_person 
                          ORDER BY DECODE(ics.CODE_INSOLVENCY_CASE_STATUS, 'ODSKRTNUTA', 2, 1), main.ic_datum_zalozeni_spisu DESC ) CODE
          from APP_COLL_JEZEVCIK_TEST.insol_list main
          JOIN APP_COLL_JEZEVCIK_TEST.person prso ON prso.skp_person = main.skp_person
          join ins_client_list icl 
            on icl.skp_person = prso.skp_person 
            and icl.code = main.ic_spisova_znacka
          JOIN owner_dwh.dc_insolvency_case ic
            ON ic.TEXT_CASE_IDENT_NUMBER = main.ic_spisova_znacka
          JOIN owner_dwh.f_insolvency_case_status_tt ics
            ON ics.SKP_INSOLVENCY_CASE = ic.SKP_INSOLVENCY_CASE
          WHERE ics.FLAG_DELETED = 'N'
            AND ics.FLAG_CURRENT = 'Y'
        ) main
  join APP_COLL_JEZEVCIK_TEST.person g
    on g.id_person = main.id_person and g.flag_valid = 'Y'
 where nvl(g.code_insolvency_case_status,' ') <> nvl(main.code,' ');
 
 
 /* RISK-618 - novy select vyssie
  
  insert into APP_COLL_JEZEVCIK_TEST.update_help_insolvency
  select distinct main.skp_person, main.code
    from (select distinct main.skp_person,
                          first_value(main.ic_stav_rizeni) over(partition by main.skp_person order by main.ic_datum_zalozeni_spisu DESC) code
            from APP_COLL_JEZEVCIK_TEST.insol_list main
                 join ins_client_list icl on icl.skp_person = main.skp_person and icl.code = main.ic_spisova_znacka) main
                 join APP_COLL_JEZEVCIK_TEST.person g
                      on g.skp_person = main.skp_person and g.flag_valid = 'Y'
   where nvl(g.code_insolvency_case_status,' ') <> nvl(main.code,' '); */

-- aktualizace dat v tabulce klient˘
update APP_COLL_JEZEVCIK_TEST.person p
   set p.code_insolvency_case_status =
       (select distinct code
          from APP_COLL_JEZEVCIK_TEST.update_help_insolvency b
         where p.id_person = b.id_person)
where exists(select 1 from update_help_insolvency h where p.id_person = h.id_person);

/* zakomentov·no 11.1.2017 MLI - mail Raisigov· Eva Sent: Wednesday, January 11, 2017 2:01 PM, RE: Kostiha - insolvence, nedobytnost
-- vymaz·nÌ z pomocnÈ tabulky ty klienty, kte¯Ì jeötÏ nemajÌ evidovan˝ û·dn˝ stav ûaloby (nem· zaloûenou kartu klienta)
delete from APP_COLL_JEZEVCIK_TEST.update_help_insolvency where skp_person not in (select skp_person from APP_COLL_JEZEVCIK_TEST.epr_status_data);
*/
-- proch·zenÌ zbyl˝ch klient˘ a vytv·¯enÌ pozn·mek a upozornÏnÌ na jejich kartÏ klienta
  FOR J IN PERSON LOOP
    
    IF J.code NOT IN ('⁄PADEK', 'ODDLUéENÕ', 'KONKURS', 'NEVYR-POST' , 'NEVYRIZENA') THEN -- RISK-807, RISK-891 pro tyhle nechceme UPO
      
        -- zÌsk·nÌ dalöÌho id pro novÈ z·znamy do tabulky upozornÏnÌ
        select max(skp_notice) into skp from APP_COLL_JEZEVCIK_TEST.notice;
        if skp is null then
          skp := 1;
        else
          skp := skp + 1 ;
        END IF;

        -- vloûenÌ upozornÏnÌ
        insert into APP_COLL_JEZEVCIK_TEST.notice t (t.skp_notice, t.id_person, t.text_notice, t.dtime_notice_created, t.dtime_notice_deadline, t.code_status, t.employee) 
        values (skp, J.id_person, 'ZmÏna v insolvenËnÌm statusu (' || J.code || ')', sysdate, trunc(sysdate), 'a', 'sys');

        -- zÌsk·nÌ dalöÌho id pro novÈ z·znamy do tabulky pozn·mek
        select max(skp_epr_note) into skp from APP_COLL_JEZEVCIK_TEST.epr_note;
        if skp is null then
          skp := 1;
        else
          skp := skp + 1 ;
        END IF;

        -- vloûenÌ novÈ pozn·mky typu upozornÏnÌ
        insert into APP_COLL_JEZEVCIK_TEST.epr_note (SKP_EPR_NOTE,
                                                     ID_PERSON,
                                                     SKP_EPR_NOTE_TYPE,
                                                     SKP_EPR_STATUS_DATA,
                                                     TEXT_NOTE,
                                                     DTIME_NOTE_CREATED,
                                                     EMPLOYEE) 
        values (skp, j.id_person, 4, NULL, 'UpozornÏnÌ: ZmÏna stavu insolvence (' || J.code || ', deadline: ' || trunc(sysdate)  || ')', sysdate, 'sys');
    
    END IF;



-- zÌsk·nÌ dalöÌho id pro novÈ z·znamy do tabulky ins pozn·mek
    select max(skp_ins_note) into skp from APP_COLL_JEZEVCIK_TEST.ins_note;
    if skp is null then
      skp := 1;
    else
      skp := skp + 1 ;
    END IF;
-- vloûenÌ novÈ pozn·mky typu upozornÏnÌ
    insert into APP_COLL_JEZEVCIK_TEST.ins_note (SKP_INS_NOTE,
                                                 SKP_INS_NOTE_TYPE,
                                                 ID_PERSON,
                                                 SKP_INS_STATUS_DATA,
                                                 TEXT_NOTE,
                                                 DTIME_NOTE_CREATED,
                                                 EMPLOYEE) 
    values (skp, 4, j.id_person, NULL, 'UpozornÏnÌ: ZmÏna stavu insolvence (' || J.code || ', deadline: ' || trunc(sysdate)  || ')', sysdate, 'sys');

  END LOOP;
  COMMIT;

begin
        dbms_stats.gather_table_stats(ownname => nazev_aktualniho_hriste,tabname => 'PERSON',
                                                  estimate_percent => dbms_stats.auto_sample_size,
                                                  cascade => true,
                                                  no_invalidate => false);
end;

execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_help_insolvency drop storage';
commit;
END;




----------------------------------------------------------------------------------------------------
-- Pridanie a update informacii k insolvenciam 
-- - tabulka ins_info; 
----------------------------------------------------------------------------------------------------
procedure update_ins_code is

-- kurzor pouûit˝ v tretej Ëasti pri zmene ËÌsla ˙Ëtu
skp INT;
  -- kurzor pro proch·zenÌ klient˘, u kter˝ch je pot¯eba zmÏnit status insolvence
  CURSOR PERSON IS
  SELECT a.id_person,
         a.skp_bank_account,
         a.text_bank_account_number 
    FROM APP_COLL_JEZEVCIK_TEST.update_ins_info3 a;


BEGIN




--***** insert do tabule  *****--
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_ins_info1 drop storage';

--RISK-985
insert into APP_COLL_JEZEVCIK_TEST.update_ins_info1 (ID_PERSON,
                                                     CODE,
                                                     SKP_BANK_ACCOUNT,
                                                     TEXT_BANK_ACCOUNT_NUMBER)
select distinct prso.id_person ,
       first_value(code) over (partition by a.skp_person order by b.ic_datum_zalozeni_spisu desc ) as code,
       first_value(ba.skp_bank_account) over (partition by a.skp_person order by ba.date_bank_account_activation desc) as skp_bank_account,
       first_value(ba.text_bank_account_number) over (partition by a.skp_person order by ba.date_bank_account_activation desc) as text_bank_account_number
from ins_client_list a
JOIN APP_COLL_JEZEVCIK_TEST.person prso ON prso.skp_person = a.skp_person
     join insol_list b on prso.skp_person = b.skp_person and a.code = b.ic_spisova_znacka
     left join owner_dwh.f_bank_account_ad ba on ba.skp_person = prso.skp_person and ba.code_bank_account_type = 'IINSOLVENCY'
where (prso.id_person) not in (select id_person from ins_info)
;

--priradime exekucni ucty
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.exe_account drop storage';
insert into APP_COLL_JEZEVCIK_TEST.exe_account (CUID,
                                                ID_PERSON,
                                                TEXT_BANK_ACCOUNT_NUMBER)
select p.cuid, p.id_person , ba.text_bank_account_number 
from owner_dwh.f_bank_account_ad ba
join APP_COLL_JEZEVCIK_TEST.person p on p.skp_person = ba.skp_person and p.flag_valid = 'Y'
and skp_bank_account_type = 21130 and skp_bank_account_status = 1;


-- k tabulke (primarnemu klucu) je vytvorena sekvencia, takûe sa do tabulky nemozu pridavat riadky bez uvazenia
insert into APP_COLL_JEZEVCIK_TEST.ins_info (SKP_INS_INFO,
                                             ID_PERSON,
                                             CODE,
                                             SKP_INS_ACCOUNT,
                                             TEXT_INS_ACCOUNT,
                                             DTIME_MODIFICATION,
                                             EMPLOYEE,
                                             DATUM_UPADKU,
                                             DATUM_PRIHLASENI)
select APP_COLL_JEZEVCIK_TEST.ins_info_seq.nextval ,
       id_person ,
       code ,
       skp_bank_account ,
       text_bank_account_number ,
       sysdate as dtime_modified,
       'sys' as employee,
       null as datum_upadku ,
       null as datum_prihlaseni
from APP_COLL_JEZEVCIK_TEST.update_ins_info1
;

execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_ins_info1 drop storage';

---------------------------------------------------------------------------------------
--***** update spisovej znacky  *****--


execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_ins_info2 drop storage';

-- klienti a nove aktualne spisove znaËky

--RISK-985
insert into update_ins_info2 (id_person, code)
select * from (
       select DISTINCT prso.id_person ,
              first_value(code) over (partition by a.skp_person order by b.ic_datum_zalozeni_spisu desc ) as code
              from ins_client_list a
              JOIN APP_COLL_JEZEVCIK_TEST.person prso ON prso.skp_person = a.skp_person
                   join insol_list b on prso.skp_person = b.skp_person and a.code = b.ic_spisova_znacka
              )
where (id_person , code) not in (select id_person , code from APP_COLL_JEZEVCIK_TEST.ins_info);



-- aktualizace dat v tabulce ins_info
update APP_COLL_JEZEVCIK_TEST.ins_info ii
   set ii.code =
       (select distinct code
          from APP_COLL_JEZEVCIK_TEST.update_ins_info2 b
         where ii.id_person = b.id_person)
where exists(select 1 from update_ins_info2 h where ii.id_person = h.id_person);



execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_ins_info2 drop storage';


---------------------------------------------------------------------------------------
--***** update uctu insolvence  *****--


execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_ins_info3 drop storage';


--RISK-985
insert into APP_COLL_JEZEVCIK_TEST.update_ins_info3 (ID_PERSON,
                                                     SKP_BANK_ACCOUNT,
                                                     TEXT_BANK_ACCOUNT_NUMBER)
select distinct prso.id_person ,
       first_value(ba.skp_bank_account) over (partition by a.skp_person order by ba.date_bank_account_activation , ba.skp_bank_account desc) as skp_bank_account,
       first_value(ba.text_bank_account_number) over (partition by a.skp_person order by ba.date_bank_account_activation , ba.skp_bank_account desc) as text_bank_account_number
from ins_client_list a
JOIN APP_COLL_JEZEVCIK_TEST.person prso ON prso.skp_person = a.skp_person
join insol_list b ON prso.skp_person = b.skp_person and a.code = b.ic_spisova_znacka
join owner_dwh.f_bank_account_ad ba on ba.skp_person = prso.skp_person and ba.code_bank_account_type = 'IINSOLVENCY'
     AND ba.FLAG_BANK_ACCOUNT_BLOCKED = 'N'
where (prso.id_person , skp_bank_account ) not in (select nvl(id_person,0) , nvl(skp_ins_account,0)  from ins_info)
;


-- aktualizace dat v tabulce ins_info
update APP_COLL_JEZEVCIK_TEST.ins_info ii
   set ii.skp_ins_account =
       (select distinct skp_bank_account
          from APP_COLL_JEZEVCIK_TEST.update_ins_info3 b
         where ii.id_person = b.id_person),
       ii.text_ins_account =
       (select distinct b.text_bank_account_number
          from APP_COLL_JEZEVCIK_TEST.update_ins_info3 b
         where ii.id_person = b.id_person)
where exists(select 1 from update_ins_info3 h where ii.id_person = h.id_person);



for J IN PERSON LOOP
-- zÌsk·nÌ dalöÌho id pro novÈ z·znamy do tabulky ins pozn·mek
    select max(skp_ins_note) into skp from APP_COLL_JEZEVCIK_TEST.ins_note;
    if skp is null then
      skp := 1;
    else
      skp := skp + 1 ;
    END IF;
-- vloûenÌ novÈ pozn·mky typu upozornÏnÌ
    insert into APP_COLL_JEZEVCIK_TEST.ins_note (SKP_INS_NOTE,
                                                 SKP_INS_NOTE_TYPE,
                                                 ID_PERSON,
                                                 SKP_INS_STATUS_DATA,
                                                 TEXT_NOTE,
                                                 DTIME_NOTE_CREATED,
                                                 EMPLOYEE) 
    values (skp, 2, j.id_person, NULL, 'UpozornÏnÌ: ZmÏna insolvenËnÌho ˙Ëtu (' || J.text_bank_account_number || ', deadline: ' || trunc(sysdate)  || ')', sysdate, 'sys');

end loop;


-- vyplneni datumu upadku  (od verze 5.3, RISK-1038)
-- POUZE klientum, kteri jej jeste vyplneny nemaji! (existuje moznost manualniho data-change u konkretnich klientu)
-- Zdrojem je tabulka f_insolvency_case_status_tt na DWH. Vezme se vûdy to nejstaröÌ datum z ('⁄PADEK', 'KONKURS', 'ODDLUéENÕ')
-- Zaroven byla uzivatelum odebrana moznost editace (karta update_klient_insol)
UPDATE APP_COLL_JEZEVCIK_TEST.INS_INFO A
   SET A.DATUM_UPADKU =
       (SELECT min(TRUNC(H.DTIME_PUBLICATION))
          FROM OWNER_DWH.DC_INSOLVENCY_CASE G
          JOIN OWNER_DWH.F_INSOLVENCY_CASE_STATUS_TT H
            ON G.SKP_INSOLVENCY_CASE = H.SKP_INSOLVENCY_CASE
           AND H.FLAG_DELETED <> 'Y'
           AND H.DTIME_CANCELLATION = DATE '1000-01-01'
         WHERE H.CODE_INSOLVENCY_CASE_STATUS IN('⁄PADEK', 'KONKURS', 'ODDLUéENÕ') -- datum upadku je prvni datum kterehokoliv z nich
           AND G.FLAG_DELETED <> 'Y'
           AND G.TEXT_CASE_IDENT_NUMBER = A.CODE
           GROUP BY h.SKP_INSOLVENCY_CASE
        ),
       a.dtime_modification = SYSDATE,
       a.employee = 'sys'
 WHERE A.DATUM_UPADKU IS NULL; -- myslenka podminky byla pro pripad, ze by bylo nektere datum manualne upraveno uzivateli (v tom pripade jej nechceme propsat)



begin
        dbms_stats.gather_table_stats(ownname => nazev_aktualniho_hriste,tabname => 'INS_INFO',
                                                  estimate_percent => dbms_stats.auto_sample_size,
                                                  cascade => true,
                                                  no_invalidate => false);
end;


execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_ins_info3 drop storage';



commit;
end;


--------------------------------------------------------------------
-- Vytvo¯enÌ pozn·mky k datu zesplatnÏnÌ klienta + 2 roky a 9 mesicu roky na promlceci lhutu
--------------------------------------------------------------------
procedure update_payoff_notice IS

  skp INT;
  CURSOR PERSON IS
  SELECT a.id_person, a.date_payoff FROM APP_COLL_JEZEVCIK_TEST.update_payoff_nt a;
  
  -- RISK-920
  cnt_days_lhuta INTEGER;
  flag_uznani_dluhu INTEGER; 
  
BEGIN

execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_payoff_nt drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_payoff_ntt drop storage';
execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.update_payoff_del drop storage';


-- RISK-920


-- vloûenÌ vhodn˝ch klient˘ s datem zesplatnÏnÌ, kte¯Ì jeötÏ nejsou v tabulce upozornÏnÌ
insert into APP_COLL_JEZEVCIK_TEST.update_payoff_nt (ID_PERSON, DATE_PAYOFF)
  select distinct p.id_person, p.date_payoff
    from APP_COLL_JEZEVCIK_TEST.person p
    left join APP_COLL_JEZEVCIK_TEST.notice n
      on n.id_person = p.id_person
     and n.text_notice like '%PromlËecÌ lh˘ta%'
   where p.flag_valid = 'Y' and p.id_person not in
         (select distinct id_person
            from APP_COLL_JEZEVCIK_TEST.epr_status_data s
           where s.skp_epr_status = 2
          minus
          select distinct id_person
            from APP_COLL_JEZEVCIK_TEST.epr_status_data s
           where s.skp_epr_status = 30
             and dtime_valid_to > sysdate)
     and p.skp_person > 0
     and p.id_person not in
         (select distinct id_person
            from APP_COLL_JEZEVCIK_TEST.epr_status_data
           where skp_epr_status in (65, 63)

          union (select distinct id_person
                  from APP_COLL_JEZEVCIK_TEST.ins_status_data
                 where skp_ins_status in (2)
                minus
                select distinct id_person
                  from APP_COLL_JEZEVCIK_TEST.ins_status_data
                 where skp_ins_status in (14)
                   and dtime_valid_to > sysdate))
     and n.id_person is null
     and p.date_death is null
     and (p.code_insolvency_case_status is null or
         p.code_insolvency_case_status not in
         ('ODDLUéENÕ', 'KONKURS', '⁄PADEK'))
     AND p.date_payoff < sysdate;



  FOR J IN PERSON LOOP
    -- zÌsk·nÌ dalöÌho id pro novÈ z·znamy do tabulky upozornÏnÌ
    select max(skp_notice) into skp from APP_COLL_JEZEVCIK_TEST.notice;
    if skp is null then
      skp := 1;
    else
      skp := skp + 1 ;
    END IF;
     
     -- RISK-920 zjistÌm jestli klient p¯i öanci uznal sv˘j dluh
     select MAX(nvl2(a.id_person,1,0)) into flag_uznani_dluhu -- pokud alespoÚ jedna jedniËka tak to znamen·, ûe uznal dluh
     from APP_COLL_JEZEVCIK_TEST.person p
     left join APP_COLL_JEZEVCIK_TEST.Chance_Import a on p.id_person = a.id_person
                                                              AND a.flag_uznani_dluhu = 'Y'
                                                              and a.flag_deleted = 'N'
     where p.id_person = J.id_person
     ; 


    IF flag_uznani_dluhu = 1 THEN
                           cnt_days_lhuta := 10*365;
    ELSE                   cnt_days_lhuta := 3*365;
    END IF;

    -- vloûenÌ upozornÏnÌ - pokud uznal dluh tak 10 let pokud neuznal tak pouze 3 roky
    insert into APP_COLL_JEZEVCIK_TEST.notice t (t.skp_notice, t.id_person, t.text_notice, t.dtime_notice_created, t.dtime_notice_deadline, t.code_status, t.employee)
    values (skp, J.id_person, 'UpozornÏnÌ: PromlËecÌ lh˘ta (' || to_char(trunc(J.date_payoff) + cnt_days_lhuta) || ')', trunc(sysdate), trunc(J.date_payoff) + cnt_days_lhuta - 90, 'a', 'sys');



  END LOOP;
  COMMIT;

--nechceme klienty, kteri maji stav zaloba podana nebo vyplnene datum umrti nebo konkurs, oddluzeni, upadek

insert into APP_COLL_JEZEVCIK_TEST.update_payoff_ntt (ID_PERSON)
  select distinct p.id_person
    from APP_COLL_JEZEVCIK_TEST.person p
   where p.flag_valid = 'Y' and ((p.id_person in
         (select distinct id_person
              from APP_COLL_JEZEVCIK_TEST.epr_status_data s
             where s.skp_epr_status = 2
            minus
            select distinct id_person
              from APP_COLL_JEZEVCIK_TEST.epr_status_data s
             where s.skp_epr_status = 30
               and dtime_valid_to > sysdate)) or
         (p.date_death is not null) or (p.code_insolvency_case_status in
         ('ODDLUéENÕ', 'KONKURS', '⁄PADEK')) or
         (p.id_person in (select distinct id_person
                              from APP_COLL_JEZEVCIK_TEST.epr_status_data
                             where skp_epr_status in (65, 63) --vymaze z tabulky notice klienty, kteri maji status doplaceno nebo vymahani ukonceno
                            /*union
                            select distinct id_person from APP_COLL_JEZEVCIK_TEST.trest_status_data
                            where skp_trest_status in (12)*/
                            union (select distinct id_person
                                    from APP_COLL_JEZEVCIK_TEST.ins_status_data
                                   where skp_ins_status in (2)
                                  minus
                                  select distinct id_person
                                    from APP_COLL_JEZEVCIK_TEST.ins_status_data
                                   where skp_ins_status in (14)
                                     and dtime_valid_to > sysdate) --vymaze klienty, kteri maji prihlaska podana, ale necha tam IR skonceno - podat zalobu
                            --komentar Hanky ñ sledovala bych promlËenÌ jen do stavu P¯ihll·öka pod·na, pak p¯eruöit; znovu aktivovat od stavu IR skonËeno ñ podat ûalobu
                            /*union
                            select distinct id_person from APP_COLL_JEZEVCIK_TEST.ded_status_data
                            where skp_ded_status in (18,22)*/
                            )))

     AND p.date_payoff < sysdate;

update notice
   set code_status = 's'
 where text_notice like '%PromlËecÌ lh˘ta%'
   and id_person in (select a.id_person from APP_COLL_JEZEVCIK_TEST.update_payoff_ntt a);

/*delete from notice n
 where text_notice like '%PromlËecÌ lh˘ta%'
   and skp_person in (select * from APP_COLL_JEZEVCIK_TEST.update_payoff_ntt);*/

------------------------------------------------------ 2.cast - mazu upozorneni kontrola pravomocnosti u klientu, kteri maji nasledne stav EPR - zruseno ------------------------

insert into APP_COLL_JEZEVCIK_TEST.update_payoff_del (ID_PERSON)
  select id_person
    from APP_COLL_JEZEVCIK_TEST.notice
   where text_notice = 'Kontrola PM RozhodnutÌ'
   and code_status = 'a'
     and id_person in
         (select distinct es.id_person
            from APP_COLL_JEZEVCIK_TEST.epr_status_data es
            join (select count(*) count, e.id_person
                   from APP_COLL_JEZEVCIK_TEST.epr_status_data e --nechceme tam klienty s vicero EPR
                   join APP_COLL_JEZEVCIK_TEST.person p
                     on p.id_person = e.id_person and p.flag_valid = 'Y'
                  where skp_epr_status = 2
                    and flag_deleted <> 'Y'
                  group by e.id_person, p.cuid) epr
              on epr.id_person = es.id_person
            left join (select id_person --podminka, aby se jednalo o upozorneni z NR a ne z jineho rizeni (pri tvorbe upozorneni v Jez se do teto tabulky prpise zaznam)
                   from APP_COLL_JEZEVCIK_TEST.trest_status_data) ee              on ee.id_person = es.id_person
           where es.skp_epr_status IN (13,18,19) -- Hanka: BÏh lh˘ty pro upozornÏnÌ na kontrolu PM rozhodnutÌ zastav prosÌm jakmile se zad· stav EPR zruöen ñ nedoruËeno, dale odvolani zalobce a zalovaneho (RISK-114).
             and epr.count < 2 --podminka, aby mel pouze 1 zalobu
             and ee.id_person is null -- nechceme lidi z trestneho rizeni
             and es.id_person not in
                 (select distinct e.id_person -----------------------------
                    from APP_COLL_JEZEVCIK_TEST.epr_status_data e
                    left join (select id_person, dtime_valid_from
                                from APP_COLL_JEZEVCIK_TEST.epr_status_data
                               where skp_epr_status IN (13,18,19)) t
                      on t.id_person = e.id_person
                     and e.dtime_valid_from < t.dtime_valid_from -- platnost stavu EPR zruseno je vetsi, nez platnost puvodniho NePM stavu
                   where e.skp_epr_status in
                         (8, 9, 29, 57, 85, 82, 71, 52, 30,15)
                     and t.id_person is null
                     and e.flag_deleted <> 'Y') --tahle skupina klientu by upozorneni mela mit, u nich to nemazu (nemaji EPR - nedoruËeno po Nepm stavu)

          )
          ;

--update upozorneni na s- storno
update notice
   set code_status = 's'
 where text_notice like 'Kontrola PM RozhodnutÌ'
   and code_status = 'a'
   and id_person in (select a.id_person from APP_COLL_JEZEVCIK_TEST.update_payoff_del a);

--

insert into APP_COLL_JEZEVCIK_TEST.epr_note (SKP_EPR_NOTE,
                                             ID_PERSON,
                                             SKP_EPR_NOTE_TYPE,
                                             SKP_EPR_STATUS_DATA,
                                             TEXT_NOTE,
                                             DTIME_NOTE_CREATED,
                                             EMPLOYEE)
  select (select max(skp_epr_note) from APP_COLL_JEZEVCIK_TEST.epr_note) + rownum,
         d.id_person,
         4,
         NULL,
         'Storno/Smazat: Kontrola PM Rozsudku (deadline: ' ||
         n.dtime_notice_deadline || ')',
         sysdate,
         'sys'
    from APP_COLL_JEZEVCIK_TEST.update_payoff_del d
    join notice n
      on d.id_person = n.id_person
     and text_notice like 'Kontrola PM RozhodnutÌ';


commit;

END;


-------------------------------------------------------------------------------------
-- Plni se tabulky splatkac_jezevcik_XY; sp_XZ
-------------------------------------------------------------------------------------
PROCEDURE SEZNAM_PLATEB
  /*authid current_user*/
 is
begin
  --bere standardni urokovou sazbu (do hlavicky dokumentu)
  --DO BUDOUCNA: pohlidat si, az se budou menit zesplatnujici dopisy (budou aktualni), aby sedelo datum prijeti platby vs. datum zesplatneni (asi se to bude na porovnavat na vteriny presne)
  --pokud platba v den zesplatneni, je zapocitana do casti tabulky pred zesplatnenim
  --pokud platba v den podani zaloby, je zapocitana az do casti tabulky po podani zaloby
  --je potreba spoustet az po dobehnuti coll_019
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik2 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik3 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik4 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik44 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik5 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik6 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik6_1 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik6_2 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik66 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik66_1 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik8 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik6N drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik7N drop storage');
  
  execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.sp1 drop storage';
  execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.sp2 drop storage';
  execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.sp3 drop storage';
  execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.sp4 drop storage';
  execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.sp5 drop storage';
  execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.sp6 drop storage';
  execute immediate 'truncate table APP_COLL_JEZEVCIK_TEST.sp8 drop storage';
 



 insert into APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik (ID_PERSON,
                                                       DATUM_PLATBY,
                                                       DATUM_SPLATKY,
                                                       UVER_CISLO,
                                                       JISTINA_SPLATKA,
                                                       UROK_SPLATKA,
                                                       JISTINA_PLATBA,
                                                       UROK_PLATBA,
                                                       DATE_PAYOFF,
                                                       AMT_LOAN_LIMIT,
                                                       SKP_LOAN)
    select t.id_person,
           t.datum_platby,
           t.datum_splatky,
           t.uver_cislo,
           nvl(t.jistina_splatka, 0) + nvl(t.jistina_splatka_1, 0) as jistina_splatka,
           nvl(t.urok_splatka_1, 0) + nvl(t.urok_splatka, 0) as urok_splatka,
           nvl(t.jistina_platba, 0) + nvl(t.jistina_platba_1, 0) as jistina_platba,
           nvl(t.urok_platba_1, 0) + nvl(t.urok_platba, 0) as urok_platba,
           trunc(l.date_payoff),
           l.amt_loan_limit,
           l.skp_loan
      from (select id_person,
                   datum_platby,
                   datum_splatky,
                   uver_cislo, --z jezevcikove tabulky, ze 2 radku na jednu platbu udela 1
                   max(case
                         when name_instalment_line_type = 'Jistina' then
                          amt_instalment
                       end) Jistina_splatka, --pokud dojdou 2 platby v jeden den odpovidajici 1 splatce, je treba mit pouze jednu z hodnot (ne soucet)
                   sum(case
                         when name_instalment_line_type = 'Jistina' then
                          amt_payment
                       end) Jistina_platba,
                   max(case
                         when name_instalment_line_type = '⁄rok' then
                          amt_instalment
                       end) urok_splatka,
                   sum(case
                         when name_instalment_line_type = '⁄rok' then
                          amt_payment
                       end) urok_platba,
                   max(case
                         when name_instalment_line_type =
                              'Mimor·dn· spl·tka jistiny' then
                          amt_instalment
                       end) Jistina_splatka_1,
                   sum(case
                         when name_instalment_line_type =
                              'Mimor·dn· spl·tka jistiny' then
                          amt_payment
                       end) Jistina_platba_1,
                   max(case
                         when name_instalment_line_type =
                              'Mimor·dn· spl·tka ˙roku' then
                          amt_instalment
                       end) urok_splatka_1,
                   sum(case
                         when name_instalment_line_type =
                              'Mimor·dn· spl·tka ˙roku' then
                          amt_payment
                       end) urok_platba_1 
              from APP_COLL_JEZEVCIK_TEST.v_datacheck_platby
             group by id_person, datum_platby, datum_splatky, uver_cislo
             /* optimalizace 19.1.2016 order by datum_splatky asc*/) t
      left join owner_dwh.f_loan_at l
        on t.uver_cislo = l.code_loan_number;
        
        
  insert into APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik2 (ID_PERSON,
                                                         DATUM_SPLATKY,
                                                         DATUM_PLATBY,
                                                         UVER_CISLO,
                                                         AMT_LOAN_LIMIT,
                                                         DATE_PAYOFF,
                                                         JISTINA_SPLATKA,
                                                         UROK_SPLATKA,
                                                         SKP_LOAN,
                                                         DATUM_PLATBY_PRED,
                                                         JISTINA_PLATBA_PRED,
                                                         UROK_PLATBA_PRED,
                                                         DATUM_PLATBY_PO,
                                                         JISTINA_PLATBA_PO,
                                                         UROK_PLATBA_PO)
    select id_person,
           datum_splatky,
           datum_platby,
           uver_cislo,
           amt_loan_limit,
           date_payoff,
           jistina_splatka,
urok_splatka,
           skp_loan,
           case
             when datum_platby <= date_payoff then
              datum_platby
           end datum_platby_PRED, -- rozdeleni podle data zesplatneni
           case
             when datum_platby <= date_payoff then
              jistina_platba
           end jistina_platba_PRED,
           case
             when datum_platby <= date_payoff then
              urok_platba
           end urok_platba_PRED,
           case
             when datum_platby > date_payoff then
              datum_platby
           end datum_platby_PO,
           case
             when datum_platby > date_payoff then
              jistina_platba
           end jistina_platba_PO,
           case
             when datum_platby > date_payoff then
              urok_platba
           end urok_platba_PO
      from APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik
     --optimalizace 19.1.2016 order by datum_platby asc
     ;
  
  
  
  insert into APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik3 (ID_PERSON,
                                                         DATUM_SPLATKY,
                                                         DATUM_PLATBY,
                                                         UVER_CISLO,
                                                         AMT_LOAN_LIMIT,
                                                         DATE_PAYOFF,
                                                         JISTINA_SPLATKA,
                                                         UROK_SPLATKA,
                                                         SKP_LOAN,
                                                         DATUM_PLATBY_PRED,
                                                         JISTINA_PLATBA_PRED,
                                                         UROK_PLATBA_PRED,
                                                         DATUM_PLATBY_PO,
                                                         JISTINA_PLATBA_PO,
                                                         UROK_PLATBA_PO,
                                                         UNPAID_PRINCIPAL_PRED,
                                                         UNPAID_PRINCIPAL_PO,
                                                         POCET)
    select s.id_person,
           s.datum_splatky,
           s.datum_platby,
           s.uver_cislo,
           s.amt_loan_limit,
           s.date_payoff,
           s.jistina_splatka,
           s.urok_splatka,
           s.skp_loan,
           s.datum_platby_pred,
           s.jistina_platba_pred,
           s.urok_platba_pred,
           s.datum_platby_po,
           s.jistina_platba_po,
           s.urok_platba_po,
           s.unpaid_principal_PRED,
           s.unpaid_principal_PO,
           s.pocet
      from --vypocitani nesplacenych castek
           (select s.id_person,
                   s.datum_splatky,
                   s.datum_platby,
                   s.uver_cislo,
                   s.amt_loan_limit,
                   s.date_payoff,
                   s.jistina_splatka,
                   s.urok_splatka,
                   s.skp_loan,
                   s.datum_platby_pred,
                   s.jistina_platba_pred,
                   s.urok_platba_pred,
                   s.datum_platby_po,
                   s.jistina_platba_po,
                   s.urok_platba_po,
                   amt_loan_limit -
                   SUM(jistina_platba_pred /*+ urok_platba_pred*/) OVER(PARTITION BY uver_cislo  ORDER BY datum_platby, datum_splatky) unpaid_principal_PRED,
                   amt_loan_limit -
                   SUM(nvl(jistina_platba_pred, 0) + /*urok_platba_pred +*/
                       nvl(jistina_platba_po, 0) /*+ urok_platba_po*/) OVER(PARTITION BY uver_cislo  ORDER BY datum_platby, datum_splatky) unpaid_principal_PO,
                   COUNT(datum_splatky /*+ urok_platba_pred*/) OVER(PARTITION BY uver_cislo, datum_splatky ORDER BY datum_splatky, datum_platby) pocet --kolikrat je tam dana splatka (info o splatce chceme pouze u jedne, ne u vsech), rozdilny partition, nez u kumulativnim souctu, protoze tady nas zajima, kolikrat je tam dany predpis splatky
              from APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik2 s) s
      --optimalizace 19.1.2016 order by datum_platby, datum_splatky
     ;
  --je treba pridat vyse predpisu z tabulky f_instalment_head (u zesplatnujici castky neni urok)
  --navic zde zjistime datum zaloby
  
  
  
  insert into APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik44 (SKF_INSTALMENT_HEAD,
                                                          FLAG_DELETED,
                                                          SKP_LOAN,
                                                          ID_PERSON,
                                                          SKF_LOAN_SCHEDULE,
                                                          CODE_STATUS,
                                                          DATE_INSTALMENT,
                                                          DATE_PAST_DUE_TOL_J,
                                                          DATE_PAST_DUE_TOL_J_PAID,
                                                          NUM_INSTALMENT_NUMBER,
                                                          NUM_INSTALMENT_VERSION,
                                                          NUM_LOAN_SCHEDULE_ORDER,
                                                          AMT_INSTALMENT,
                                                          AMT_INSTALMENT_PRINCIPAL,
                                                          AMT_INSTALMENT_INTEREST,
                                                          AMT_PAYMENT,
                                                          AMT_PAYMENT_PRINCIPAL,
                                                          AMT_PAYMENT_INTEREST,
                                                          FLAG_VALID,
                                                          DTIME_VALID_FROM,
                                                          DTIME_VALID_TO,
                                                          DTIME_INSERTED,
                                                          DTIME_MODIFIED,
                                                          DATE_PAY_INSTALMENT,
                                                          DATE_PAY_LAST) --chceme pouze klienty, kteri jsou v mp_person, ne vsechny
                                                          
    select h.skf_instalment_head,
           h.flag_deleted,
           h.skp_loan,
           h.id_person,
           h.skf_loan_schedule,
           h.code_status,
           h.date_instalment,
           h.date_past_due_tol_j,
           h.date_past_due_tol_j_paid,
           h.num_instalment_number,
           h.num_instalment_version,
           h.num_loan_schedule_order,
           h.amt_instalment,
           h.amt_instalment_principal,
           h.amt_instalment_interest,
           h.amt_payment,
           h.amt_payment_principal,
           h.amt_payment_interest,
           h.flag_valid,
           h.dtime_valid_from,
           h.dtime_valid_to,
           h.dtime_inserted,
           h.dtime_modified,
           h.date_pay_instalment,
           h.date_pay_last
      -- from owner_dwh.f_instalment_head_ad                           -- risk-723 fa_instalment
      from APP_COLL_JEZEVCIK_TEST.INSTALMENT_HEAD h
     where id_person in
           (select distinct id_person from APP_COLL_JEZEVCIK_TEST.person p where p.flag_valid = 'Y')
       and code_status = 'a'
       and NUM_LOAN_SCHEDULE_ORDER = 0
       and flag_valid = 'Y'
      ;
       
  insert into APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik4 (ID_PERSON,
                                                         DATUM_SPLATKY,
                                                         DATUM_PLATBY,
                                                         UVER_CISLO,
                                                         AMT_LOAN_LIMIT,
                                                         DATE_PAYOFF,
                                                         JISTINA_SPLATKA,
                                                         UROK_SPLATKA,
                                                         SKP_LOAN,
                                                         DATUM_PLATBY_PRED,
                                                         JISTINA_PLATBA_PRED,
                                                         UROK_PLATBA_PRED,
                                                         DATUM_PLATBY_PO,
                                                         JISTINA_PLATBA_PO,
                                                         UROK_PLATBA_PO,
                                                         UNPAID_PRINCIPAL_PRED,
                                                         UNPAID_PRINCIPAL_PO,
                                                         POCET,
                                                         AMT_INSTALMENT,
                                                         AMT_INSTALMENT_PRINCIPAL,
                                                         AMT_INSTALMENT_INTEREST,
                                                         DATUM_ZALOBY)
    select j.id_person,
           j.datum_splatky,
           j.datum_platby,
           j.uver_cislo,
           j.amt_loan_limit,
           j.date_payoff,
           j.jistina_splatka,
           j.urok_splatka,
           j.skp_loan,
           j.datum_platby_pred,
           j.jistina_platba_pred,
           j.urok_platba_pred,
           j.datum_platby_po,
           j.jistina_platba_po,
           j.urok_platba_po,
           j.unpaid_principal_pred,
           j.unpaid_principal_po,
           j.pocet,
           i.amt_instalment,
           i.amt_instalment_principal,
           i.amt_instalment_interest,
           case
             when e.dtime_valid_from is null then
              trunc(sysdate)
             else
              e.dtime_valid_from
           end as datum_zaloby
      from APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik3 j
      join APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik44 i
        on j.skp_loan = i.skp_loan
       and j.DATUM_SPLATKY = i.date_instalment
      left join (select id_person,
                        skp_epr_status,
                        max(dtime_valid_from) dtime_valid_from
                   from APP_COLL_JEZEVCIK_TEST.epr_status_data --problem u klientu, kteri maji vice zalob - vzdy se bere ta nejaktualnejsi
                  where skp_epr_status = 2
                    and flag_deleted = 'N'
                  group by id_person, skp_epr_status) e
        on e.id_person = j.id_person;
 

 --ke kazde splatce vyberu vzdy pouze u jeden predpis (u stejne splatky je vice plateb)
  insert into APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik5 (ID_PERSON,
                                                         DATUM_SPLATKY,
                                                         DATUM_PLATBY,
                                                         UVER_CISLO,
                                                         AMT_LOAN_LIMIT,
                                                         DATE_PAYOFF,
                                                         JISTINA_SPLATKA,
                                                         UROK_SPLATKA,
                                                         SKP_LOAN,
                                                         DATUM_PLATBY_PRED,
                                                         JISTINA_PLATBA_PRED,
                                                         UROK_PLATBA_PRED,
                                                         DATUM_PLATBY_PO,
                                                         JISTINA_PLATBA_PO,
                                                         UROK_PLATBA_PO,
                                                         UNPAID_PRINCIPAL_PRED,
                                                         UNPAID_PRINCIPAL_PO,
                                                         POCET,
                                                         AMT_INSTALMENT,
                                                         AMT_INSTALMENT_PRINCIPAL,
                                                         AMT_INSTALMENT_INTEREST,
                                                         DATUM_ZALOBY,
                                                         SPLATKA_PREDPIS,
                                                         JISTINA_PREDPIS,
                                                         UROK_PREDPIS,
                                                         DATUM_SPLATKY_SPRAVNE)
    select s.id_person,
           s.datum_splatky,
           s.datum_platby,
           s.uver_cislo,
           s.amt_loan_limit,
           s.date_payoff,
           s.jistina_splatka,
           s.urok_splatka,
           s.skp_loan,
           s.datum_platby_pred,
           s.jistina_platba_pred,
           s.urok_platba_pred,
           s.datum_platby_po,
           s.jistina_platba_po,
           s.urok_platba_po,
           s.unpaid_principal_pred,
           s.unpaid_principal_po,
           s.pocet,
           s.amt_instalment,
           s.amt_instalment_principal,
           s.amt_instalment_interest,
           s.datum_zaloby,
           case
             when pocet = 1 then
              max(s.AMT_INSTALMENT)
              over(partition by uver_cislo, datum_splatky)
             else
              null
           end splatka_predpis,
           case
             when pocet = 1 then
              max(s.AMT_INSTALMENT_PRINCIPAL)
              over(partition by uver_cislo, datum_splatky)
             else
              null
           end jistina_predpis,
           case
             when pocet = 1 then
              max(s.AMT_INSTALMENT_INTEREST)
              over(partition by uver_cislo, datum_splatky)
             else
              null
           end urok_predpis,
           case
             when s.pocet = 1 then
              s.datum_splatky
             else
              null
           end datum_splatky_spravne
      from APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik4 s
     --optimalizace 19.1.2016 order by datum_platby, datum_splatky
     ;
  --nyni rozdelime tabulku c. 5 na tri tabulky - podle toho, kdy zaplatil (zda po zesplatneni nebo pred)
  --kvuli formatu tabulky v jezevcikovi
 
 
 
  insert into APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik6 (ID_PERSON,
                                                         DATUM_SPLATKY,
                                                         DATUM_PLATBY,
                                                         UVER_CISLO,
                                                         AMT_LOAN_LIMIT,
                                                         DATE_PAYOFF,
                                                         JISTINA_SPLATKA,
                                                         UROK_SPLATKA,
                                                         SKP_LOAN,
                                                         DATUM_PLATBY_PRED,
                                                         JISTINA_PLATBA_PRED,
                                                         UROK_PLATBA_PRED,
                                                         DATUM_PLATBY_PO,
                                                         JISTINA_PLATBA_PO,
                                                         UROK_PLATBA_PO,
                                                         UNPAID_PRINCIPAL_PRED,
                                                         UNPAID_PRINCIPAL_PO,
                                                         POCET,
                                                         AMT_INSTALMENT,
                                                         AMT_INSTALMENT_PRINCIPAL,
                                                         AMT_INSTALMENT_INTEREST,
                                                         DATUM_ZALOBY,
                                                         POCET2)
    select s.id_person,
           s.datum_splatky,
           s.datum_platby,
           s.uver_cislo,
           s.amt_loan_limit,
           s.date_payoff,
           s.jistina_splatka,
           s.urok_splatka,
           s.skp_loan,
           s.datum_platby_pred,
           s.jistina_platba_pred,
           s.urok_platba_pred,
           s.datum_platby_po,
           s.jistina_platba_po,
           s.urok_platba_po,
           s.unpaid_principal_pred,
           s.unpaid_principal_po,
           s.pocet,
           s.amt_instalment,
           s.amt_instalment_principal,
           s.amt_instalment_interest,
           s.datum_zaloby,
           COUNT(datum_splatky /*+ urok_platba_pred*/) OVER(PARTITION BY uver_cislo, datum_splatky  ORDER BY datum_splatky, datum_platby) pocet2 --je to nova tabulka, chceme novy pocet (pro spodni cast tabulky dole v excelu)
      from APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik4 s
     where s.datum_platby > s.date_payoff;
  
  
  
  insert INTO APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik6_1 (ID_PERSON,
                                                           DATUM_SPLATKY,
                                                           DATUM_PLATBY,
                                                           UVER_CISLO,
                                                           AMT_LOAN_LIMIT,
                                                           DATE_PAYOFF,
                                                           JISTINA_SPLATKA,
                                                           UROK_SPLATKA,
                                                           SKP_LOAN,
                                                           DATUM_PLATBY_PRED,
                                                           JISTINA_PLATBA_PRED,
                                                           UROK_PLATBA_PRED,
                                                           DATUM_PLATBY_PO,
                                                           JISTINA_PLATBA_PO,
                                                           UROK_PLATBA_PO,
                                                           UNPAID_PRINCIPAL_PRED,
                                                           UNPAID_PRINCIPAL_PO,
                                                           POCET,
                                                           AMT_INSTALMENT,
                                                           AMT_INSTALMENT_PRINCIPAL,
                                                           AMT_INSTALMENT_INTEREST,
                                                           DATUM_ZALOBY,
                                                           POCET2,
                                                           DATUM_SPLATKY_SPRAVNE2,
                                                           SPLATKA_PREDPIS,
                                                           JISTINA_PREDPIS,
                                                           UROK_PREDPIS)
    select s.id_person,
           s.datum_splatky,
           s.datum_platby,
           s.uver_cislo,
           s.amt_loan_limit,
           s.date_payoff,
           s.jistina_splatka,
           s.urok_splatka,
           s.skp_loan,
           s.datum_platby_pred,
           s.jistina_platba_pred,
           s.urok_platba_pred,
           s.datum_platby_po,
           s.jistina_platba_po,
           s.urok_platba_po,
           s.unpaid_principal_pred,
           s.unpaid_principal_po,
           s.pocet,
           s.amt_instalment,
           s.amt_instalment_principal,
           s.amt_instalment_interest,
           s.datum_zaloby,
           s.pocet2,
           case
             when s.pocet2 = 1 then
              s.datum_splatky
             else
              null
           end datum_splatky_spravne2,
           case
             when s.pocet2 = 1 then
              max(s.AMT_INSTALMENT)
              over(partition by uver_cislo, datum_splatky)
             else
              null
           end splatka_predpis,
           case
             when s.pocet2 = 1 then
              max(s.AMT_INSTALMENT_PRINCIPAL)
              over(partition by uver_cislo, datum_splatky)
             else
              null
           end jistina_predpis,
           case
             when s.pocet2 = 1 then
              max(s.AMT_INSTALMENT_INTEREST)
              over(partition by uver_cislo, datum_splatky)
             else
              null
           end urok_predpis
      from APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik6 s;
 


 --je treba odecist, co uz zaplatil
  insert into APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik6_2 (ID_PERSON,
                                                           DATUM_SPLATKY,
                                                           DATUM_PLATBY,
                                                           UVER_CISLO,
                                                           AMT_LOAN_LIMIT,
                                                           DATE_PAYOFF,
                                                           JISTINA_SPLATKA,
                                                           UROK_SPLATKA,
                                                           SKP_LOAN,
                                                           DATUM_PLATBY_PRED,
                                                           JISTINA_PLATBA_PRED,
                                                           UROK_PLATBA_PRED,
                                                           DATUM_PLATBY_PO,
                                                           JISTINA_PLATBA_PO,
                                                           UROK_PLATBA_PO,
                                                           UNPAID_PRINCIPAL_PRED,
                                                           UNPAID_PRINCIPAL_PO,
                                                           POCET,
                                                           AMT_INSTALMENT,
                                                           AMT_INSTALMENT_PRINCIPAL,
                                                           AMT_INSTALMENT_INTEREST,
                                                           DATUM_ZALOBY,
                                                           POCET2,
                                                           DATUM_SPLATKY_SPRAVNE2,
                                                           SPLATKA_PREDPIS,
                                                           JISTINA_PREDPIS,
                                                           UROK_PREDPIS,
                                                           JISTINA_NEZAPLACENO,
                                                           UROK_NEZAPLACENO,
                                                           SPLATKA_NEZAPLACENO) -- jsou zde vsechny platby po zesplatneni
    select s.id_person,
           s.datum_splatky,
           s.datum_platby,
           s.uver_cislo,
           s.amt_loan_limit,
           s.date_payoff,
           s.jistina_splatka,
           s.urok_splatka,
           s.skp_loan,
           s.datum_platby_pred,
           s.jistina_platba_pred,
           s.urok_platba_pred,
           s.datum_platby_po,
           s.jistina_platba_po,
           s.urok_platba_po,
           s.unpaid_principal_pred,
           s.unpaid_principal_po,
           s.pocet,
           s.amt_instalment,
           s.amt_instalment_principal,
           s.amt_instalment_interest,
           s.datum_zaloby,
           s.pocet2,
           s.datum_splatky_spravne2,
           s.splatka_predpis,
           s.jistina_predpis,
           s.urok_predpis,
           s.jistina_predpis - nvl(t.jp, 0) jistina_nezaplaceno,
           s.urok_predpis - nvl(t.up, 0) urok_nezaplaceno,
           splatka_predpis - nvl(t.cp, 0) as splatka_nezaplaceno
      from APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik6_1 s
      left join (select datum_splatky,
                        uver_cislo,
                        sum(jistina_platba) jp,
                        sum(urok_platba) up,
                        sum(jistina_platba) + sum(urok_platba) cp --kolik zaplatil na splatkach, do zesplatneni, odecteme
                   from APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik
                  where datum_platby < = date_payoff
                  group by datum_splatky, uver_cislo) t
        on t.uver_cislo = s.uver_cislo
       and t.datum_splatky = s.datum_splatky
     --optimalizace 19.1.2016 order by s.DATUM_SPLATKY
     ;
  
  
  
  
  --nyni rozlisime podle toho, zda prisla po nebo pred podanim zaloby
  insert into APP_COLL_JEZEVCIK_TEST.sp1 (ID_PERSON,
                                          DATUM_SPLATKY,
                                          DATUM_PLATBY,
                                          UVER_CISLO,
                                          AMT_LOAN_LIMIT,
                                          DATE_PAYOFF,
                                          JISTINA_SPLATKA,
                                          UROK_SPLATKA,
                                          SKP_LOAN,
                                          DATUM_PLATBY_PRED,
                                          JISTINA_PLATBA_PRED,
                                          UROK_PLATBA_PRED,
                                          DATUM_PLATBY_PO,
                                          JISTINA_PLATBA_PO,
                                          UROK_PLATBA_PO,
                                          UNPAID_PRINCIPAL_PRED,
                                          UNPAID_PRINCIPAL_PO,
                                          POCET,
                                          AMT_INSTALMENT,
                                          AMT_INSTALMENT_PRINCIPAL,
                                          AMT_INSTALMENT_INTEREST,
                                          DATUM_ZALOBY,
                                          POCET2,
                                          DATUM_SPLATKY_SPRAVNE2,
                                          SPLATKA_PREDPIS,
                                          JISTINA_PREDPIS,
                                          UROK_PREDPIS,
                                          JISTINA_NEZAPLACENO,
                                          UROK_NEZAPLACENO,
                                          SPLATKA_NEZAPLACENO,
                                          DATUM_PLATBY_MEZI,
                                          JISTINA_PLATBA_MEZI,
                                          UROK_PLATBA_MEZI,
                                          UNPAID_PRINCIPA_MEZI)
    select s.id_person,
           s.datum_splatky,
           s.datum_platby,
           s.uver_cislo,
           s.amt_loan_limit,
           s.date_payoff,
           s.jistina_splatka,
           s.urok_splatka,
           s.skp_loan,
           s.datum_platby_pred,
           s.jistina_platba_pred,
           s.urok_platba_pred,
           s.datum_platby_po,
           s.jistina_platba_po,
           s.urok_platba_po,
           s.unpaid_principal_pred,
           s.unpaid_principal_po,
           s.pocet,
           s.amt_instalment,
           s.amt_instalment_principal,
           s.amt_instalment_interest,
           s.datum_zaloby,
           s.pocet2,
           s.datum_splatky_spravne2,
           s.splatka_predpis,
           s.jistina_predpis,
           s.urok_predpis,
           s.jistina_nezaplaceno,
           s.urok_nezaplaceno,
           s.splatka_nezaplaceno,
           case
             when (s.datum_platby < s.datum_zaloby) then
              s.datum_platby_po
             else
              null
           end as datum_platby_mezi, --pokud platba ve stejny den, jako zaloba, tak se nepocita (zapise se az do radku po zalobe)
           case
             when (s.datum_platby < s.datum_zaloby) then
              s.jistina_platba_po
             else
              null
           end as jistina_platba_mezi,
           case
             when (s.datum_platby < s.datum_zaloby) then
              s.urok_platba_po
             else
              null
           end as urok_platba_mezi,
           case
             when (s.datum_platby < s.datum_zaloby) then
              s.unpaid_principal_po
             else
              null
           end as unpaid_principa_mezi
      from APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik6_2 s;
  
  
  
  insert into APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik66 (ID_PERSON,
                                                          DATUM_SPLATKY,
                                                          DATUM_PLATBY,
                                                          UVER_CISLO,
                                                          AMT_LOAN_LIMIT,
                                                          DATE_PAYOFF,
                                                          JISTINA_SPLATKA,
                                                          UROK_SPLATKA,
                                                          SKP_LOAN,
                                                          DATUM_PLATBY_PRED,
                                                          JISTINA_PLATBA_PRED,
                                                          UROK_PLATBA_PRED,
                                                          DATUM_PLATBY_PO,
                                                          JISTINA_PLATBA_PO,
                                                          UROK_PLATBA_PO,
                                                          UNPAID_PRINCIPAL_PRED,
                                                          UNPAID_PRINCIPAL_PO,
                                                          POCET,
                                                          AMT_INSTALMENT,
                                                          AMT_INSTALMENT_PRINCIPAL,
                                                          AMT_INSTALMENT_INTEREST,
                                                          DATUM_ZALOBY,
                                                          POCET2) --nove 25.5., platby po podani zaloby, chceme tam ale u predpisu splatek odecteno, kolik prislo mezi zesplatnenim a zalobou
    select s.id_person,
           s.datum_splatky,
           s.datum_platby,
           s.uver_cislo,
           s.amt_loan_limit,
           s.date_payoff,
           s.jistina_splatka,
           s.urok_splatka,
           s.skp_loan,
           s.datum_platby_pred,
           s.jistina_platba_pred,
           s.urok_platba_pred,
           s.datum_platby_po,
           s.jistina_platba_po,
           s.urok_platba_po,
           s.unpaid_principal_pred,
           s.unpaid_principal_po,
           s.pocet,
           s.amt_instalment,
           s.amt_instalment_principal,
           s.amt_instalment_interest,
           s.datum_zaloby,
           COUNT(datum_splatky /*+ urok_platba_pred*/) OVER(PARTITION BY uver_cislo, datum_splatky ORDER BY datum_splatky, datum_platby) pocet2
      from APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik4 s
      left join (select id_person,
                        skp_epr_status,
                        max(dtime_valid_from) datum_zaloby
                   from APP_COLL_JEZEVCIK_TEST.epr_status_data --problem u klientu, kteri maji vice zalob - vzdy se bere ta nejaktualnejsi
                  where skp_epr_status = 2
                    and flag_deleted = 'N'
                  group by id_person, skp_epr_status) e
        on e.id_person = s.id_person
       and e.skp_epr_status = 2
     where s.datum_platby > e.datum_zaloby;
  
  
  
  insert into APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik66_1 (ID_PERSON,
                                                            DATUM_SPLATKY,
                                                            DATUM_PLATBY,
                                                            UVER_CISLO,
                                                            AMT_LOAN_LIMIT,
                                                            DATE_PAYOFF,
                                                            JISTINA_SPLATKA,
                                                            UROK_SPLATKA,
                                                            SKP_LOAN,
                                                            DATUM_PLATBY_PRED,
                                                            JISTINA_PLATBA_PRED,
                                                            UROK_PLATBA_PRED,
                                                            DATUM_PLATBY_PO,
                                                            JISTINA_PLATBA_PO,
                                                            UROK_PLATBA_PO,
                                                            UNPAID_PRINCIPAL_PRED,
                                                            UNPAID_PRINCIPAL_PO,
                                                            POCET,
                                                            AMT_INSTALMENT,
                                                            AMT_INSTALMENT_PRINCIPAL,
                                                            AMT_INSTALMENT_INTEREST,
                                                            DATUM_ZALOBY,
                                                            POCET2,
                                                            DATUM_SPLATKY_SPRAVNE2,
                                                            SPLATKA_PREDPIS,
                                                            JISTINA_PREDPIS,
                                                            UROK_PREDPIS) --nove 25.5. po podani zaloby
    select s.id_person,
           s.datum_splatky,
           s.datum_platby,
           s.uver_cislo,
           s.amt_loan_limit,
           s.date_payoff,
           s.jistina_splatka,
           s.urok_splatka,
           s.skp_loan,
           s.datum_platby_pred,
           s.jistina_platba_pred,
           s.urok_platba_pred,
           s.datum_platby_po,
           s.jistina_platba_po,
           s.urok_platba_po,
           s.unpaid_principal_pred,
           s.unpaid_principal_po,
           s.pocet,
           s.amt_instalment,
           s.amt_instalment_principal,
           s.amt_instalment_interest,
           s.datum_zaloby,
           s.pocet2,
           case
             when s.pocet2 = 1 then
              s.datum_splatky
             else
              null
           end datum_splatky_spravne2,
           case
             when s.pocet2 = 1 then
              max(s.AMT_INSTALMENT)
              over(partition by uver_cislo, datum_splatky)
             else
              null
           end splatka_predpis,
           case
             when s.pocet2 = 1 then
              max(s.AMT_INSTALMENT_PRINCIPAL)
              over(partition by uver_cislo, datum_splatky)
             else
              null
           end jistina_predpis,
           case
             when s.pocet2 = 1 then
              max(s.AMT_INSTALMENT_INTEREST)
              over(partition by uver_cislo, datum_splatky)
             else
              null
           end urok_predpis
      from APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik66 s;
  
  
  
  insert into APP_COLL_JEZEVCIK_TEST.sp5 (ID_PERSON,
                                          DATUM_SPLATKY,
                                          DATUM_PLATBY,
                                          UVER_CISLO,
                                          AMT_LOAN_LIMIT,
                                          DATE_PAYOFF,
                                          JISTINA_SPLATKA,
                                          UROK_SPLATKA,
                                          SKP_LOAN,
                                          DATUM_PLATBY_PRED,
                                          JISTINA_PLATBA_PRED,
                                          UROK_PLATBA_PRED,
                                          DATUM_PLATBY_PO,
                                          JISTINA_PLATBA_PO,
                                          UROK_PLATBA_PO,
                                          UNPAID_PRINCIPAL_PRED,
                                          UNPAID_PRINCIPAL_PO,
                                          POCET,
                                          AMT_INSTALMENT,
                                          AMT_INSTALMENT_PRINCIPAL,
                                          AMT_INSTALMENT_INTEREST,
                                          DATUM_ZALOBY,
                                          POCET2,
                                          DATUM_SPLATKY_SPRAVNE2,
                                          SPLATKA_PREDPIS,
                                          JISTINA_PREDPIS,
                                          UROK_PREDPIS,
                                          JISTINA_NEZAPLACENO,
                                          UROK_NEZAPLACENO,
                                          SPLATKA_NEZAPLACENO) --sp5 nove vytovreno 25.5.
    select s.id_person,
           s.datum_splatky,
           s.datum_platby,
           s.uver_cislo,
           s.amt_loan_limit,
           s.date_payoff,
           s.jistina_splatka,
           s.urok_splatka,
           s.skp_loan,
           s.datum_platby_pred,
           s.jistina_platba_pred,
           s.urok_platba_pred,
           s.datum_platby_po,
           s.jistina_platba_po,
           s.urok_platba_po,
           s.unpaid_principal_pred,
           s.unpaid_principal_po,
           s.pocet,
           s.amt_instalment,
           s.amt_instalment_principal,
           s.amt_instalment_interest,
           s.datum_zaloby,
           s.pocet2,
           s.datum_splatky_spravne2,
           s.splatka_predpis,
           s.jistina_predpis,
           s.urok_predpis,
           s.jistina_predpis - nvl(t.jp, 0) jistina_nezaplaceno,
           s.urok_predpis - nvl(t.up, 0) urok_nezaplaceno,
           splatka_predpis - nvl(t.cp, 0) as splatka_nezaplaceno
      from APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik66_1 s
--vsechny platby po podani zaloby
      left join (select j.datum_splatky,
                        j.uver_cislo,
                        sum(j.jistina_platba) jp,
                        sum(j.urok_platba) up,
                        sum(j.jistina_platba) + sum(j.urok_platba) cp --kolik zaplatil na vsech splatkach, do datumu zaloby, odecteme
                   from APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik j
                   left join (select id_person,
                                    skp_epr_status,
                                    max(dtime_valid_from) datum_zaloby
                               from APP_COLL_JEZEVCIK_TEST.epr_status_data --problem u klientu, kteri maji vice zalob - vzdy se bere ta nejaktualnejsi
                              where skp_epr_status = 2
                                and flag_deleted = 'N'
                              group by id_person, skp_epr_status) e
                     on e.id_person = j.id_person
                    and e.skp_epr_status = 2
                  where j.datum_platby < e.datum_zaloby --vsechny splatky do datumu zaloby
                  group by j.datum_splatky, j.uver_cislo) t
        on t.uver_cislo = s.uver_cislo
       and t.datum_splatky = s.datum_splatky
     --optimalizace 19.1.2016 order by s.DATUM_platby, s.datum_splatky
     ;
  
  
  
  --splatky, u kterych dosla platba pred zesplatnenim
  insert into APP_COLL_JEZEVCIK_TEST.sp2 (ID_PERSON,
                                          DATUM_SPLATKY,
                                          DATUM_PLATBY,
                                          UVER_CISLO,
                                          AMT_LOAN_LIMIT,
                                          DATE_PAYOFF,
                                          JISTINA_SPLATKA,
                                          UROK_SPLATKA,
                                          SKP_LOAN,
                                          DATUM_PLATBY_PRED,
                                          JISTINA_PLATBA_PRED,
                                          UROK_PLATBA_PRED,
                                          DATUM_PLATBY_PO,
                                          JISTINA_PLATBA_PO,
                                          UROK_PLATBA_PO,
                                          UNPAID_PRINCIPAL_PRED,
                                          UNPAID_PRINCIPAL_PO,
                                          POCET,
                                          AMT_INSTALMENT,
                                          AMT_INSTALMENT_PRINCIPAL,
                                          AMT_INSTALMENT_INTEREST,
                                          DATUM_ZALOBY,
                                          SPLATKA_PREDPIS,
                                          JISTINA_PREDPIS,
                                          UROK_PREDPIS,
                                          DATUM_SPLATKY_SPRAVNE)
    select sj.id_person,
           sj.datum_splatky,
           sj.datum_platby,
           sj.uver_cislo,
           sj.amt_loan_limit,
           sj.date_payoff,
           sj.jistina_splatka,
           sj.urok_splatka,
           sj.skp_loan,
           sj.datum_platby_pred,
           sj.jistina_platba_pred,
           sj.urok_platba_pred,
           sj.datum_platby_po,
           sj.jistina_platba_po,
           sj.urok_platba_po,
           sj.unpaid_principal_pred,
           sj.unpaid_principal_po,
           sj.pocet,
           sj.amt_instalment,
           sj.amt_instalment_principal,
           sj.amt_instalment_interest,
           sj.datum_zaloby,
           sj.splatka_predpis,
           sj.jistina_predpis,
           sj.urok_predpis,
           sj.datum_splatky_spravne
      from APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik5 sj
     where datum_platby <= date_payoff;
  
  
  
  --navic je potreba overit, jestli tam nechybi nejaka splatka (neni v jezevcikovi, protoze neprobehla platba, ale jeji predpis je pred datem zesplateni)
  --lepsi napojit az tady tuto tabulku, protoze v tabulce 7 muze chybet splatka do zespl. (vypadla v situaci, ze platba prisla po zesplatneni)
  insert into APP_COLL_JEZEVCIK_TEST.sp3 (DATE_INSTALMENT,
                                          AMT_INSTALMENT,
                                          AMT_INSTALMENT_PRINCIPAL,
                                          AMT_INSTALMENT_INTEREST,
                                          CODE_LOAN_NUMBER)
    select i.date_instalment,
           i.AMT_INSTALMENT,
           i.AMT_INSTALMENT_PRINCIPAL,
           i.AMT_INSTALMENT_interest,
           l.code_loan_number
      from APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik44 i
      join owner_dwh.f_loan_at l
        on i.skp_loan = l.skp_loan
       and l.amt_payoff > 0
      left join APP_COLL_JEZEVCIK_TEST.sp2 j
        on i.skp_loan = j.skp_loan
       and j.DATUM_SPLATKY = i.DATE_INSTALMENT
    --where i.NUM_LOAN_SCHEDULE_ORDER = 0 and  i.flag_valid = 'Y' and i.flag_deleted = 'N'
     where j.uver_cislo is null
     --optimalizace 19.1.2016 order by date_instalment
     ; --platby nejsou v tabulce jezevcika
 


 --mezi zespl a datum podanim zaloby: splatky, ze kterych je zesplatnena castka, ale nebyla z nich ani cast zaplaceno po zesplatneni
  insert into APP_COLL_JEZEVCIK_TEST.sp4 (UVER_CISLO,
                                          DATE_INSTALMENT,
                                          NS,
                                          NJ,
                                          NU)
    select l.code_loan_number as uver_cislo,
           i.date_instalment,
           i.AMT_INSTALMENT - nvl(i.AMT_PAYMENT, 0) NS,
           i.AMT_INSTALMENT_PRINCIPAL - nvl(i.AMT_PAYMENT_PRINCIPAL, 0) as NJ,
           i.AMT_INSTALMENT_INTEREST - nvl(i.AMT_PAYMENT_INTEREST, 0) as NU
      from APP_COLL_JEZEVCIK_TEST.INSTALMENT_HEAD i -- povodne APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik44 i -- zmenene v RISK-953
      join owner_dwh.f_loan_at l
        on i.skp_loan = l.skp_loan
      left join APP_COLL_JEZEVCIK_TEST.sp1 j
        on j.uver_cislo = l.code_loan_number
       and j.DATUM_SPLATKY = i.DATE_INSTALMENT
    --where i.NUM_LOAN_SCHEDULE_ORDER = 0 and  i.flag_valid = 'Y' and i.flag_deleted = 'N' --and l.code_loan_number = 'HU00001140'
     where nvl(i.amt_payment, 0) < i.amt_instalment
       AND NOT (i.amt_payment IS NULL AND i.date_PAY_INSTALMENT < SYSDATE) -- RISK-953 -- podminka odstrani zaznamy, ktore sa chybne naparovali kvoli hodnote NULL v amt_payment (DWH chybka)
       and j.DATUM_SPLATKY is NULL
       AND i.id_person in
           (select distinct id_person from APP_COLL_JEZEVCIK_TEST.person p where p.flag_valid = 'Y')
       and code_status = 'a'
       and NUM_LOAN_SCHEDULE_ORDER = 0
       and flag_valid = 'Y'
     --optimalizace 19.1.2016 order by i.date_instalment
     ;
 
 
 
  insert into APP_COLL_JEZEVCIK_TEST.sp8 (UVER_CISLO,
                                          DATE_INSTALMENT,
                                          NS,
                                          NJ,
                                          NU) -- je to jine, nez sp4 (v podmince pouze vynechame platby, ktere prisly po podani zaloby (vcetne))
    select l.code_loan_number as uver_cislo,
           i.date_instalment,
           i.AMT_INSTALMENT - nvl(i.AMT_PAYMENT, 0) NS,
           i.AMT_INSTALMENT_PRINCIPAL - nvl(i.AMT_PAYMENT_PRINCIPAL, 0) as NJ,
           i.AMT_INSTALMENT_INTEREST - nvl(i.AMT_PAYMENT_INTEREST, 0) as NU
      from APP_COLL_JEZEVCIK_TEST.INSTALMENT_HEAD i -- povodne APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik44 i  -- zmenene v RISK-953
      join owner_dwh.f_loan_at l
        on i.skp_loan = l.skp_loan
      left join APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik6 j
        on j.uver_cislo = l.code_loan_number
       and j.DATUM_SPLATKY = i.DATE_INSTALMENT
       and j.datum_platby >= j.datum_zaloby --pouze platby po podani zaloby (vcetne), ne mezi zespl. a podanim zaloby
    --where i.NUM_LOAN_SCHEDULE_ORDER = 0 and  i.flag_valid = 'Y' and i.flag_deleted = 'N' --and l.code_loan_number = 'HU00001140'
     where nvl(i.amt_payment, 0) < i.amt_instalment
       AND NOT (i.amt_payment IS NULL AND i.date_PAY_INSTALMENT < SYSDATE) -- RISK-953 -- podminka odstrani zaznamy, ktore sa chybne naparovali kvoli hodnote NULL v amt_payment (DWH chybka)
       and j.DATUM_SPLATKY is null
       AND i.id_person in
           (select distinct id_person from APP_COLL_JEZEVCIK_TEST.person p where p.flag_valid = 'Y')
       and code_status = 'a'
       and NUM_LOAN_SCHEDULE_ORDER = 0
       and flag_valid = 'Y'
     --optimalizace 19.1.2016 order by i.date_instalment
     ;
  
  ---- >>>>>>>>>   odtialto   RISK-953
  --  z tabuliek sp4 a sp8 odstranime zaznamy ktere tam nemeji co delat (nahlasena DWH chyba)
  DELETE FROM APP_COLL_JEZEVCIK_TEST.sp4 sp4
  WHERE sp4.uver_cislo || to_char(sp4.date_instalment) IN (
                                              SELECT sp4.uver_cislo || to_char(sp4.date_instalment)
                                              FROM APP_COLL_JEZEVCIK_TEST.sp4 sp4
                                              JOIN APP_COLL_JEZEVCIK_TEST.loan l
                                                ON l.code_loan_number = sp4.uver_cislo
                                              WHERE sp4.date_instalment < l.date_payoff - 65
                                                AND l.code_loan_number LIKE 'HU%'
                                                AND l.date_payoff < SYSDATE + 1
                                                AND sp4.nu = 0
                                                )
  ;

  DELETE FROM APP_COLL_JEZEVCIK_TEST.sp8 sp
  WHERE sp.uver_cislo || to_char(sp.date_instalment) IN (
                                              SELECT sp.uver_cislo || to_char(sp.date_instalment)
                                              FROM APP_COLL_JEZEVCIK_TEST.sp8 sp
                                              JOIN APP_COLL_JEZEVCIK_TEST.loan l
                                                ON l.code_loan_number = sp.uver_cislo
                                              WHERE sp.date_instalment < l.date_payoff - 65
                                                AND l.code_loan_number LIKE 'HU%'
                                                AND l.date_payoff < SYSDATE + 1
                                                AND sp.nu = 0
                                                )
  ;
  
  ---- <<<<<<<<<   potialto   RISK-953
  
  
  
  
  --jeste je potreba zjistit, kdy doslo k podani zaloby - nova tabulka, ostatni tabule tam obcas jsou a obcas ne, proto nova
  insert into APP_COLL_JEZEVCIK_TEST.sp6 (CODE_LOAN_NUMBER, DATUM_ZALOBY)
    select l.code_loan_number,
           case
             when trunc(e.dtime_valid_from) is null then
              trunc(sysdate)
             else
              trunc(e.dtime_valid_from)
           end as datum_zaloby
      -- from owner_dwh.f_loan_at l -- nahrazeno za nas LOAN v RISK-985
      FROM APP_COLL_JEZEVCIK_TEST.loan l
      left join (select id_person,
                        skp_epr_status,
                        max(dtime_valid_from) dtime_valid_from
                   from APP_COLL_JEZEVCIK_TEST.epr_status_data --problem u klientu, kteri maji vice zalob - vzdy se bere ta nejaktualnejsi
                  where skp_epr_status = 2
                    and flag_deleted = 'N'
                  group by id_person, skp_epr_status) e
        on e.id_person = l.id_person
       and e.SKP_EPR_STATUS = 2;
 
 
  
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik2 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik3 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik4 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik44 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik5 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik6 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik6_1 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik6_2 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik66 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik66_1 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik8 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.splatkac_jezevcik6N drop storage');
  commit;
end;



----------------------------------------------------------------------------------------------------
-- Plni se tabuly uznani_dluhu2; wrong_dp
----------------------------------------------------------------------------------------------------

PROCEDURE UZNANI_DLUHU is
begin
  --delete from APP_COLL_JEZEVCIK_TEST.uznani_dluhu; --RISK-481
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.uznani_dluhu2 drop storage');
  --delete from APP_COLL_JEZEVCIK_TEST.uznani_dluhu3;
  
 /* insert into APP_COLL_JEZEVCIK_TEST.uznani_dluhu --RISK-481
    with klienti as
     (select p.cuid, p.skp_person, p.name, p.date_payoff
        from APP_COLL_JEZEVCIK_TEST.person p
        where p.flag_valid = 'Y'
     )
    select l.code_loan_number,
           l.AMT_LOAN_PRINCIPAL_TOTAL,
           l.skp_person,
           l.date_loan_utilization,
           ba.text_bank_account_number,
           kk.name,
           k.dtime_signature,
           --pp.date_birth,
           kk.date_payoff,
           
           kk.cuid
      from owner_dwh.f_loan_at l
      join owner_dwh.f_completion_subject_at j
        on j.skp_person = l.skp_person
       and l.skp_loan = j.skp_loan
      left join owner_dwh.f_completion_signature_tt k
        on j.skp_completion_subject = k.skp_completion_subject
       and k.dtime_signature_valid_to > sysdate
      left join owner_dwh.f_bank_account_ad ba
        on ba.skp_bank_account = l.skp_bank_account
      join klienti kk
        on kk.skp_person = l.skp_person
     
     where j.code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING')
       and l.date_loan_utilization < sysdate
       and l.skp_loan_status not in (14, 15, 16, 740, 3, 13, 135);
  */   
  
  insert into APP_COLL_JEZEVCIK_TEST.uznani_dluhu2 (CODE_LOAN_NUMBER,
                                                    ID_PERSON,
                                                    SKP_LOAN,
                                                    AMT_PAYOFF_PRINCIPAL,
                                                    AMT_INTEREST_OVERDUE,
                                                    AMT_PAYOFF_INTEREST,
                                                    AMT_PAYOFF_PENALTY,
                                                    AMT_PAYOFF,
                                                    JIS,
                                                    DATUM_PLATBY,
                                                    RATE_INTEREST_STANDARD,
                                                    DATE_PAYOFF,
                                                    DP,
                                                    APP,
                                                    AMT_SU_START,
                                                    RATE_INTEREST_REPO)
    select l.code_loan_number,
           o.id_person,
           o.skp_loan,
           o.amt_payoff_principal,
           o.amt_interest_overdue,
           o.amt_payoff_interest,
           o.amt_payoff_penalty,
           o.amt_payoff,
           case
             when t.jis is null then
              o.amt_payoff_principal
             else
              t.jis
           end as jis, --pokud neprisla zadna platba (tablka sp1, tak se bere castka pri zesplatneni)
           t.datum_platby,
           l.rate_interest_standard,
           l.DATE_PAYOFF,
           l.date_payoff + 10 as dp,
           l.AMT_PAYOFF_PRINCIPAL app,
           case
             when su.suu is not null then
              su.suu
             else
              l.AMT_PAYOFF_PRINCIPAL
           end as amt_su_start,
           l.rate_interest_repo
      from APP_COLL_JEZEVCIK_TEST.balance_obs o
      left join (select min(unpaid_principal_po) jis,
                        skp_loan,
                        uver_cislo,
                        datum_platby
                   from APP_COLL_JEZEVCIK_TEST.sp1
                  where date_payoff < datum_platby
                  group by skp_loan, uver_cislo, datum_platby) t
        on t.skp_loan = o.skp_loan
      join owner_dwh.f_loan_at l
        on l.skp_loan = o.skp_loan
      left join (select max(unpaid_principal_po) suu,
                        min(datum_platby) dpp,
                        skp_loan
                   from APP_COLL_JEZEVCIK_TEST.sp1
                  where date_payoff + 10 > datum_platby
                  group by skp_loan) su
        on su.skp_loan = t.skp_loan
     where l.skp_loan_status not in (14, 15, 16, 740, 3, 13, 135);
     
/*  insert into APP_COLL_JEZEVCIK_TEST.uznani_dluhu3
    select sr.skp_person,
           ba.text_bank_account_number,
           sr.amt_balance,
           sr.code_currency
      from APP_COLL_JEZEVCIK_TEST.ucet sr
      join owner_dwh.f_bank_account_ad ba
        on ba.SKP_BANK_ACCOUNT = sr.SKP_BANK_ACCOUNT
     where sr.amt_balance < 0;
  commit;*/
  ---- wrong_dp se tyka nacteni dat pro info, zda klient nema odlisne data zesplatneni v obs a jez (nechtelo se mi delat novou proceduru pro jednu tabulku :))

  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.wrong_dp drop storage');
  
  insert into APP_COLL_JEZEVCIK_TEST.wrong_dp (ID_PERSON,
                                               DATE_PAYOFF,
                                               DATUM_POZADAVKU_NA_ZESPLATNENI)
    select prso.id_person,
           l.date_payoff,
           ll.dtime_valid_from datum_pozadavku_na_zesplatneni /*, ll.skp_employee_status_changed, ls.name_payoff_reason, e.name_last*/
      from owner_dwh.f_loan_at l
      JOIN APP_COLL_JEZEVCIK_TEST.person prso ON prso.skp_person = l.skp_person
      join (select skp_loan, dtime_valid_from, skp_employee_status_changed
              from OWNER_DWH.f_loan_status_tt
             where code_loan_wfstatus = 7) ll
        on l.skp_loan = ll.skp_loan
       and trunc(l.date_payoff) <> trunc(ll.dtime_valid_from)
      join owner_dwh.cl_payoff_reason ls
        on ls.skp_payoff_reason = l.skp_payoff_reason
     where ls.name_payoff_reason = 'Delikvence';
     
   COMMIT;
   
end;



-------------------------------------------------------------------------------------
-- PodkladovÈ data pro öablonu EPR; plnÌ se tabulky sablona_epr_XY
-------------------------------------------------------------------------------------
PROCEDURE SABLONA is
begin
  --select * from OWNER_DWH.DC_PERSON_DEMOGRAPHY;
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.sablona_epr drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.sablona_epr2 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.sablona_epr3 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.sablona_epr4 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.sablona_epr5 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.sablona_epr6 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.sablona_epr7 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.sablona_epr8 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.sablona_epr9 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.sablona_epr10 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.sablona_epr11 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.sablona_epr12 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.sablona_epr14 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.sablona_epr15 drop storage');
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.sablona_epr drop storage');
  
  
  insert into APP_COLL_JEZEVCIK_TEST.sablona_epr (TEXT_CONTRACT_NUMBER,
                                                  DTIME_SIGNATURE,
                                                  SKP_BANK_ACCOUNT,
                                                  SKP_LOAN,
                                                  ID_PERSON,
                                                  CODE_COMPLETION_SUBJECT_TYPE,
                                                  DTIME_COMPLETION_FINISHED,
                                                  TEXT_APPENDIX_NUMBER,
                                                  CODE_LOAN_NUMBER,
                                                  TEXT_BANK_ACCOUNT_NUMBER,
                                                  DTIME_COMPLETION_SUBJ_INSERTED,
                                                  DATE_LOAN_UTILIZATION)
    with tt as
     (select distinct p.skp_person,
                      l.skp_loan,
                      l.code_loan_number,
                      l.DATE_LOAN_UTILIZATION  -- RISK-940
        from APP_COLL_JEZEVCIK_TEST.person p
        join owner_dwh.f_loan_at l
          on p.skp_person = l.skp_person and p.flag_valid = 'Y'
      )
    select c.text_contract_number,
           k.dtime_signature,
           j.skp_bank_account,
           j.skp_loan,
           prso.id_person,
           j.code_completion_subject_type,
           j.dtime_completion_finished,
           j.text_appendix_number,
           tt.code_loan_number,
           ba.text_bank_account_number,
           j.dtime_completion_subj_inserted,
           case
             when j.code_completion_subject_type in
                  ('CREATE_LOAN', 'CREATE_REFINANCING') then
              tt.DATE_LOAN_UTILIZATION
             else
              null
           end as DATE_LOAN_UTILIZATION
      from --,j.* , l.* , k.* from
           owner_dwh.f_completion_subject_at j
      join owner_dwh.f_contract_at c
        on j.skp_contract = c.skp_contract
      JOIN APP_COLL_JEZEVCIK_TEST.person prso 
        ON prso.skp_person = c.skp_person
        AND prso.flag_valid = 'Y'
      left join tt
        on tt.skp_person = c.skp_person
       and tt.skp_loan = j.skp_loan
      join owner_dwh.cl_completion_subject_type l
        on j.skp_completion_subject_type = l.skp_completion_subject_type
      left join owner_dwh.f_completion_signature_tt k
        on j.skp_completion_subject = k.skp_completion_subject
       and k.dtime_signature_valid_to > sysdate
      left join owner_dwh.f_bank_account_ad ba
        on ba.skp_bank_account = j.skp_bank_account;
           
        
  insert into APP_COLL_JEZEVCIK_TEST.sablona_epr2 (FLAG_DELETED,
                                                   SKP_LOAN,
                                                   CODE_STATUS,
                                                   DATE_INSTALMENT,
                                                   AMT_INSTALMENT,
                                                   DTIME_VALID_FROM,
                                                   DTIME_VALID_TO,
                                                   DTIME_MODIFIED,
                                                   DTIME_INSERTED,
                                                   NUM_INSTALMENT_NUMBER,
                                                   SKF_LOAN_SCHEDULE,
                                                   NUM_INSTALMENT_VERSION)
with vyjimky as 
(--seznam vyjimek, u kterych nevadi, ze maji vygenerovany SK pred nacerpanim - v den nacerpani maji totiz dodatek
  select loan.skp_loan 
  from sablona_epr loan
  join sablona_epr dod on dod.skp_loan = loan.skp_loan and trunc(dod.dtime_completion_subj_inserted) = trunc(loan.date_loan_utilization)
  where loan.code_completion_subject_type = 'CREATE_LOAN' and dod.code_completion_subject_type in ('CHANGE_INST_AMOUNT','CHANGE_INST_DATE')
)
    select /*+ parallel(i 4) */
           i.flag_deleted,
           i.skp_loan,
           i.code_status,
           i.date_instalment,
           i.amt_instalment,
           i.dtime_valid_from,
           i.dtime_valid_to,
           i.dtime_modified,
           i.dtime_inserted,
           i.num_instalment_number,
           i.skf_loan_schedule,
           i.num_instalment_version
      from APP_COLL_JEZEVCIK_TEST.person p
      join owner_dwh.f_loan_at l  on p.skp_person = l.skp_person
      -- join owner_dwh.f_instalment_head_ad i on l.skp_loan = i.skp_loan and l.skp_person = i.skp_person -- RISK-723 FA_instalment, zakomentovane 3.8.2017
      join APP_COLL_JEZEVCIK_TEST.instalment_head i on l.skp_loan = i.skp_loan and p.id_person = i.id_person
      
     where p.flag_valid = 'Y'
       and i.num_loan_schedule_order = 0
       and i.DATE_INSTALMENT > i.dtime_inserted
       and i.date_instalment <= sysdate --omezit budoucÌ spl·tky  - do tabulky epr3 se n·slednÏ napoËÌt·v· datum poslednÌ spl·tky
       and i.dtime_valid_from < i.dtime_valid_to
      and ((trunc(i.DTIME_INSERTED) >= trunc(l.DATE_LOAN_UTILIZATION)) -- RISK-940
           or(l.skp_loan in (select skp_loan from vyjimky))) ;
        --podminka, aby se SK vytvoril az po nacerpani
     --optimalizace 19.1.2016 order by i.date_instalment, i.dtime_valid_from
      
    /*    zaloha 7.2.2017 - odstranit, pokud 8.2.2017 probÏhne n·poËet ok
      owner_dwh.f_instalment_head_ad i
      join owner_dwh.f_loan_at l
        on l.skp_loan = i.skp_loan
     where num_loan_schedule_order = 0
       and i.DATE_INSTALMENT > i.dtime_inserted
       and i.skp_person in
           (select distinct skp_person from APP_COLL_JEZEVCIK_TEST.person where flag_valid = 'Y')
       and dtime_valid_from < dtime_valid_to
       and ((trunc(i.DTIME_INSERTED) >= trunc(l.date_loan_utilization))
           or(l.skp_loan in (select skp_loan from vyjimky))) 
        --podminka, aby se SK vytvoril az po nacerpani
     --optimalizace 19.1.2016 order by i.date_instalment, i.dtime_valid_from
     ;*/

  insert into APP_COLL_JEZEVCIK_TEST.sablona_epr3 (SKP_LOAN,
                                                   DTIME_VALID_FROM,
                                                   SPLATKA,
                                                   MIN_DATUM_SPLATKY,
                                                   MAX_DATUM_SPLATKY,
                                                   DEN,
                                                   POCET) --souhrny o vsech splatkacich
    select skp_loan,
           dtime_valid_from,
           max(amt_instalment) splatka,
           min(date_instalment) min_datum_splatky,
           max(date_instalment) max_datum_splatky,
           EXTRACT(day FROM min(date_instalment)) den,
           max(num_instalment_number) pocet
      from APP_COLL_JEZEVCIK_TEST.sablona_epr2
     group by skp_loan, dtime_valid_from, num_instalment_version;
     
  insert into APP_COLL_JEZEVCIK_TEST.sablona_epr4 (SKP_LOAN,
                                                   DTIME_VALID_FROM,
                                                   SPLATKA,
                                                   MIN_DATUM_SPLATKY,
                                                   MAX_DATUM_SPLATKY,
                                                   DEN,
                                                   POCET,
                                                   DTIME_SIGNATURE,
                                                   CODE_COMPLETION_SUBJECT_TYPE,
                                                   TEXT_APPENDIX_NUMBER)
  /*
    select s.*,
           ss.dtime_signature,
           ss.code_completion_subject_type,
           ss.text_appendix_number --, sss.min_pujcky
      from APP_COLL_JEZEVCIK_TEST.sablona_epr3 s
      join APP_COLL_JEZEVCIK_TEST.sablona_epr ss
        on s.skp_loan = ss.skp_loan
       and ((trunc(s.dtime_valid_from) = trunc(ss.dtime_signature)) or
           (trunc(s.dtime_valid_from) =
           trunc(ss.dtime_completion_subj_inserted)) or
           (trunc(s.dtime_valid_from) = trunc(ss.date_loan_utilization)));
 */
 with ttt as (
     select s.*,
           ss.dtime_signature,
           ss.code_completion_subject_type,
           ss.text_appendix_number --, sss.min_pujcky
      from APP_COLL_JEZEVCIK_TEST.sablona_epr3 s
      join APP_COLL_JEZEVCIK_TEST.sablona_epr ss
        on s.skp_loan = ss.skp_loan
       and ((trunc(s.dtime_valid_from) = trunc(ss.dtime_signature)) or
           (trunc(s.dtime_valid_from) =
           trunc(ss.dtime_completion_subj_inserted)) or
           (trunc(s.dtime_valid_from) = trunc(ss.date_loan_utilization))))
 
     , t as
     (select epr.*,
             count(skp_loan) over(partition by skp_loan, code_completion_subject_type order by dtime_valid_from) cnt
        from ttt epr),
    tt as (select t.skp_loan, t.dtime_valid_from ,t.splatka, t.min_datum_splatky, t.max_datum_splatky, t.den, t.pocet, t.dtime_signature, t.code_completion_subject_type, t.text_appendix_number
      from t t
     where (code_completion_subject_type = 'CREATE_LOAN' and cnt = 1)
        or code_completion_subject_type <> 'CREATE_LOAN')
        select * from tt;
        
        
                   
  insert into APP_COLL_JEZEVCIK_TEST.sablona_epr5 (SKP_LOAN,
                                                   DTIME_VALID_FROM,
                                                   SPLATKA,
                                                   MIN_DATUM_SPLATKY,
                                                   MAX_DATUM_SPLATKY,
                                                   DEN,
                                                   POCET,
                                                   DTIME_SIGNATURE,
                                                   CODE_COMPLETION_SUBJECT_TYPE,
                                                   TEXT_APPENDIX_NUMBER,
                                                   MIN_PUJCKY)
    select s.skp_loan,
           s.dtime_valid_from,
           s.splatka,
           s.min_datum_splatky,
           s.max_datum_splatky,
           s.den,
           s.pocet,
           s.dtime_signature,
           s.code_completion_subject_type,
           s.text_appendix_number, 
           sss.min_pujcky
      from APP_COLL_JEZEVCIK_TEST.sablona_epr4 s
      join (select ss.min_datum_splatky min_pujcky, ss.skp_loan
              from APP_COLL_JEZEVCIK_TEST.sablona_epr4 ss
              join (select max(dtime_valid_from) maxx, skp_loan
                     from APP_COLL_JEZEVCIK_TEST.sablona_epr4
                    where code_completion_subject_type in
                          ('CREATE_LOAN', 'CREATE_REFINANCING')
                    group by skp_loan) s
                on s.maxx = ss.dtime_valid_from
               and s.skp_loan = ss.skp_loan
             where ss.code_completion_subject_type in
                   ('CREATE_LOAN', 'CREATE_REFINANCING')) sss
        on sss.skp_loan = s.skp_loan;
  ----------splatky pred zesplatnenim
  insert into APP_COLL_JEZEVCIK_TEST.sablona_epr6 (DATUM_SPLATKY,
                                                   DATUM_PLATBY,
                                                   PLATBA,
                                                   SPLATKA,
                                                   SKP_LOAN,
                                                   PAID_IN_TIME)
    select datum_splatky,
           max(datum_platby) datum_platby,
           sum(jistina_platba_pred) + sum(urok_platba_pred) platba,
           sum(splatka_predpis) splatka,
           skp_loan,
           case
             when datum_splatky <> max(datum_platby) then
              1
             else
              0
           end as paid_in_time
      from APP_COLL_JEZEVCIK_TEST.sp2
     group by datum_splatky, skp_loan;
     
  --castecne platby pred zesplatnenim
  insert into APP_COLL_JEZEVCIK_TEST.sablona_epr7 (SKP_LOAN,
                                                   DATUM_SPLATKY,
                                                   DATUM_PLATBY,
                                                   JISTINA_PLATBA_PRED,
                                                   UROK_PLATBA_PRED,
                                                   PLATBA,
                                                   SPLATKA_PREDPIS)
    select s.skp_loan,
           s.datum_splatky,
           s.datum_platby,
           s.jistina_platba_pred,
           s.urok_platba_pred,
           s.jistina_platba_pred + s.urok_platba_pred platba,
           s.splatka_predpis
      from APP_COLL_JEZEVCIK_TEST.sp2 s
      left join (select  a.datum_splatky,
                         a.datum_platby,
                         a.platba,
                         a.splatka,
                         a.skp_loan,
                         a.paid_in_time
                   from APP_COLL_JEZEVCIK_TEST.sablona_epr6 a
                  where platba = splatka) se
        on se.datum_splatky = s.datum_splatky
       and se.skp_loan = s.skp_loan
     where se.datum_splatky is null;
     
     
  insert into APP_COLL_JEZEVCIK_TEST.sablona_epr8 (TEXT_CONTRACT_NUMBER,
                                                   DTIME_SIGNATURE,
                                                   SKP_BANK_ACCOUNT,
                                                   SKP_LOAN,
                                                   ID_PERSON,
                                                   CODE_COMPLETION_SUBJECT_TYPE,
                                                   DTIME_COMPLETION_FINISHED,
                                                   TEXT_APPENDIX_NUMBER,
                                                   CODE_LOAN_NUMBER,
                                                   TEXT_BANK_ACCOUNT_NUMBER,
                                                   DTIME_COMPLETION_SUBJ_INSERTED,
                                                   DATE_LOAN_UTILIZATION,
                                                   AMT_INSTALMENT,
                                                   DATE_INST)
    select s.text_contract_number,
           s.dtime_signature,
           s.skp_bank_account,
           s.skp_loan,
           s.id_person,
           s.code_completion_subject_type,
           s.dtime_completion_finished,
           s.text_appendix_number,
           s.code_loan_number,
           s.text_bank_account_number,
           s.dtime_completion_subj_inserted,
           s.date_loan_utilization, 
           f.amt_instalment, 
           trunc(f.date_instalment) date_inst
      from APP_COLL_JEZEVCIK_TEST.sablona_epr s
      join owner_dwh.f_instalment_payment_ad f
        on s.skp_loan = f.skp_loan
       and trunc(s.dtime_completion_finished) = f.date_instalment
     where s.code_completion_subject_type = 'PARTIAL_REPAYMENT'
       and f.code_status = 'a'
       and f.skp_instalment_line_type = 8;
  
  --vypocitame pohledavku ke dni zesplatneni
  insert into APP_COLL_JEZEVCIK_TEST.sablona_epr9 (SKP_LOAN,
                                                   ID_PERSON,
                                                   JPLATBA,
                                                   UPLATBA,
                                                   JPREDPIS,
                                                   UPREDPIS,
                                                   IP,
                                                   II,
                                                   PLATBA_DEN_PO_ZESPLATNENI)
    select p.skp_loan,
           p.id_person,
           sp2.jplatba,
           sp2.uplatba,
           sp2.jpredpis,
           sp2.upredpis,
           sp3.ip,
           sp3.ii,
           case
             when q.min_d_p - 1 = q.date_payoff then
              'Y'
             else
              'N'
           end platba_den_po_zesplatneni --zda prisla platba den po zesplatneni
      from APP_COLL_JEZEVCIK_TEST.balance_obs p
      left join (select skp_loan,
                        sum(jistina_platba_pred) as jplatba,
                        sum(urok_platba_pred) as uplatba,
                        sum(jistina_predpis) as jpredpis,
                        sum(urok_predpis) as upredpis
                   from APP_COLL_JEZEVCIK_TEST.sp2
                  group by skp_loan) sp2
        on sp2.skp_loan = p.skp_loan
      join owner_dwh.f_loan_at l
        on l.skp_loan = p.skp_loan
      left join (select code_loan_number,
                        sum(amt_instalment_principal) as ip,
                        sum(amt_instalment_interest) as ii
                   from APP_COLL_JEZEVCIK_TEST.sp3
                  group by code_loan_number) sp3
        on sp3.code_loan_number = l.code_loan_number
      left join (select skp_loan, date_payoff, min(datum_platby) as min_d_p
                   from APP_COLL_JEZEVCIK_TEST.sp1
                  group by date_payoff, skp_loan) q
        on q.skp_loan = p.skp_loan;
        
        
  insert into APP_COLL_JEZEVCIK_TEST.sablona_epr10 (SKP_LOAN,
                                                    ID_PERSON,
                                                    DATE_PAYOFF,
                                                    DP)
    select l.skp_loan,
           p.id_person,
           trunc(p.date_payoff) date_payoff,
           trunc(p.date_payoff - 1) as dp
      from APP_COLL_JEZEVCIK_TEST.person p
      join owner_dwh.f_loan_at l
        on l.skp_person = p.skp_person and p.flag_valid = 'Y';
        
        
        
  insert into APP_COLL_JEZEVCIK_TEST.sablona_epr11 (DATUM_PLATBY,
                                                    DP,
                                                    JISTINA,
                                                    ID_PERSON,
                                                    UVER_CISLO)
    select datum_platby,
           datum_platby - 1 as dp,
           min(unpaid_principa_mezi) jistina,
           id_person,
           uver_cislo
      from APP_COLL_JEZEVCIK_TEST.sp1
     group by datum_platby, datum_platby - 1, id_person, uver_cislo;
     
     
  insert into APP_COLL_JEZEVCIK_TEST.sablona_epr12 (DON, ID_PERSON)
    select trunc(tc.executiondate) don, p.id_person /*tc.EXECUTIONDATE*/
      from OWNER_INT.V_LCS_ACTOR a
      left join OWNER_INT.V_LCS_CLASSIFIABLE c
        on a.T_ACTOR_ID = c.T_ACTOR_1
       and c.ctype = 'ACTOR'
      left join OWNER_INT.V_LCS_TASKCALENDAR tc
        on tc.T_CLASSIFIABLE_1 = c.T_CLASSIFIABLE_ID
      left join owner_int.v_lcs_tasktype tt
        on tt.T_TASKTYPE_ID = tc.T_TASKTYPE_1
      join APP_COLL_JEZEVCIK_TEST.person p
        on p.cuid = a.customeridentifier and p.flag_valid = 'Y'
     WHERE lower(tt.taskname) like '%termination letter%'
       and tc.deleted = 0 and tc.executed = 1;
  
  
  execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.ins_sablona drop storage');
  commit;
  
  
  --zjistime datum posledni kompletne zaplacene splatky (platby i po zesplatneni)
  insert into APP_COLL_JEZEVCIK_TEST.sablona_epr14 (DATUM_SPLATKY, SKP_LOAN)
    with t as
     (select skp_loan,
             uver_cislo,
             p.datum_splatky,
             p.amt_instalment,
             sum(p.amt_payment) platby
        from v_datacheck_platby p
        join owner_dwh.f_loan_at l
          on l.code_loan_number = p.uver_cislo
       group by uver_cislo, datum_splatky, p.amt_instalment, skp_loan
      )
    select max(datum_splatky) datum_splatky, skp_loan
      from t
     where platby = amt_instalment
     group by skp_loan;


  
  insert into APP_COLL_JEZEVCIK_TEST.ins_sablona (ID_PERSON, TEXT_BANK_ACCOUNT_NUMBER)
    select p.id_person,
           ba.TEXT_BANK_ACCOUNT_NUMBER
      from APP_COLL_JEZEVCIK_TEST.person p
      left join owner_dwh.f_bank_account_ad ba
        on ba.skp_person = p.skp_person
       and ba.SKP_BANK_ACCOUNT_type = 21270
   where p.flag_valid = 'Y';
  commit;

--tabulka na dotazeni informace, zda klient mel v den zesplatneni exekucni blokaci

insert into sablona_epr15 (skp_loan)
  select l.skp_loan
    from owner_dwh.f_loan_at l
    join owner_dwh.f_account_balance_td p
      on p.SKP_BANK_ACCOUNT = l.SKP_BANK_ACCOUNT
     and l.date_payoff between p.DATE_BALANCE_VALID_FROM and
         p.DATE_BALANCE_VALID_TO
   where p.FLAG_DELETED = 'N'
     and p.AMT_GOVERNMENT_HOLD_SUM > 0
     and l.date_payoff < sysdate;  
  
end;


-------------------------------------------------------------------------------------
--cilem teto procedury je rozdelit klienty podle toho, zda jsou vymahani nebo ne a na jakem produktu (hyref, hu, debet)
-------------------------------------------------------------------------------------
procedure update_person_status is 
--RISK 252
  
begin
  -------------------------------------------------NOVI KLIENTI----------------------------------------------------
  --nahrani nove pridanych klientu do Jezevcika, rozdeleni podle stavu (hu/hyref a debet) a podle toho, zda jsou collected
  insert into person_status_data (skp_person_status_data,
                                  id_person,
                                  skp_person_status,
                                  flag_collected,
                                  dtime_valid_from,
                                  dtime_valid_to,
                                  dtime_modification,
                                  flag_valid,
                                  flag_deleted,
                                  employee)
    with novy_klienti as
     (select p.id_person, p.date_payoff
        from person p
        left join person_status_data ps
          on p.id_person = ps.id_person
       where ps.id_person is null and p.flag_valid = 'Y'
      )
    
    select rownum +
           (select max(skp_person_status_data) from person_status_data),
           nk.id_person,
           case
             when hr.id_person is not null then
              2
             when a.id_person is not null then
              1
             else
              3
           end, --o jaky typ klieta se jedna
           case
             when nk.date_payoff < sysdate and
                  (a.id_person is not null or hr.id_person is not null) then
              'Y'
             when a.id_person is null and hr.id_person is null and
                  am.amt_min < 0 then
              'Y'
             else
              'N'
           end, --zda je vymahan nebo ne ,
           sysdate,
           DATE'3000-01-01',
           sysdate,
           'Y',
           'N',
           'sys'
      from novy_klienti nk
      left join (select distinct id_person
                   from intervaly
                  where code_loan_number like 'HU%'
                    and name_loan_status = 'AktivnÌ') a
        on a.id_person = nk.id_person -- zjistime, jestli ma klient aktivni uver. Pokud nebude mit, je to debetar
      left join (select min(amt_balance) amt_min, id_person
                   from ucet
                  group by id_person) am
        on am.id_person = nk.id_person --podle toho, zda budou mit debet nebo to budou aktivne/neaktivne vym. deb. 
      left join (select distinct id_person
                   from intervaly
                  where code_loan_number like 'HR%'
                    and name_loan_status = 'AktivnÌ') hr
        on hr.id_person = nk.id_person --zjistime, zda nema HYREF
    ;
---------------------------------------------UPRAVA PRODUKTU, PODLE KTEREHO JSOU KLIENTI VYMAHANI-------------------------
  ------------------------------------------------------------------------------------------------------------------------
  --zmena statusu z debet na hyref nebo hu
  --klienti, kteri jsou vedeni jako debetari a maji uver
  execute immediate ('truncate table update_person_status_dl drop storage');
  insert into update_person_status_dl (id_person, flag_collected)
    select p.id_person, p.flag_collected
      from person_status_data p
      join (select distinct id_person
              from intervaly
             where code_loan_number like 'HU%'
               and name_loan_status = 'AktivnÌ') l
        on l.id_person = p.id_person
     where p.skp_person_status = 3
       and p.flag_valid = 'Y'
       and p.flag_deleted = 'N';

  update person_status_data
     set dtime_valid_to     = trunc(sysdate) - 1 / (60 * 60 * 24),
         flag_valid         = 'N',
         dtime_modification = sysdate
   where id_person in
         (select id_person from update_person_status_dl)
     and flag_valid = 'Y';

  insert into person_status_data
    select rownum +
           (select max(skp_person_status_data) from person_status_data),
           id_person,
           1,
           flag_collected,
           trunc(sysdate),
           DATE'3000-01-01',
           sysdate,
           'Y',
           'N',
           'sys'
      from update_person_status_dl;

  --zmena statusu z debet na hyref

  execute immediate ('truncate table update_person_status_dh drop storage');

  insert into update_person_status_dh (id_person,
                                       flag_deleted,
                                       flag_collected)
    select p.id_person, p.flag_deleted, p.flag_collected
      from person_status_data p
      join (select distinct id_person
              from intervaly
             where code_loan_number like 'HR%'
               and name_loan_status = 'AktivnÌ') l
        on l.id_person = p.id_person
     where p.skp_person_status = 3
       and p.flag_valid = 'Y'
       and p.flag_deleted = 'N';

  update person_status_data
     set dtime_valid_to     = trunc(sysdate) - 1 / (60 * 60 * 24),
         flag_valid         = 'N',
         dtime_modification = sysdate
   where id_person in
         (select id_person from update_person_status_dh)
     and flag_valid = 'Y';

  insert into person_status_data (skp_person_status_data,
                                  id_person,
                                  skp_person_status,
                                  flag_collected,
                                  dtime_valid_from,
                                  dtime_valid_to,
                                  dtime_modification,
                                  flag_valid,
                                  flag_deleted,
                                  employee)
    select rownum +
           (select max(skp_person_status_data) from person_status_data),
           id_person,
           2,
           flag_collected,
           trunc(sysdate),
           DATE'3000-01-01',
           sysdate,
           'Y',
           'N',
           'sys'
      from update_person_status_dh;

  --zmena z hu na hyref

  execute immediate ('truncate table update_person_status_lh drop storage');

  insert into update_person_status_lh (id_person,
                                       flag_deleted,
                                       flag_collected)
    select p.id_person, p.flag_deleted, p.flag_collected
      from person_status_data p
      join (select distinct id_person
              from intervaly
             where code_loan_number like 'HR%'
               and name_loan_status = 'AktivnÌ') l
        on l.id_person = p.id_person
     where p.skp_person_status = 1
       and p.flag_valid = 'Y'
       and p.flag_deleted = 'N';

  update person_status_data
     set dtime_valid_to     = trunc(sysdate) - 1 / (60 * 60 * 24),
         flag_valid         = 'N',
         dtime_modification = sysdate
   where id_person in
         (select id_person from update_person_status_lh)
     and flag_valid = 'Y';

  insert into person_status_data (skp_person_status_data,
                                  id_person,
                                  skp_person_status,
                                  flag_collected,
                                  dtime_valid_from,
                                  dtime_valid_to,
                                  dtime_modification,
                                  flag_valid,
                                  flag_deleted,
                                  employee)
    select rownum +
           (select max(skp_person_status_data) from person_status_data),
           id_person,
           2,
           flag_collected,
           trunc(sysdate),
           DATE'3000-01-01',
           sysdate,
           'Y',
           'N',
           'sys'
      from update_person_status_lh;
  ---------------------------------------------UPRAVA FLAG_COLLECTED U VSECH KATEGORII-------------------------------------------
  --zmena collected a no/collected  ----------------------------------------------------------------------------------------------
  --u HU a HYREF pouze zmena stavu z vymahanych na nevymahane
  -- u debetaru

  --nejdrive HU, vymahani urcujeme podle jednotilvch stavu v Jez

  /*jake jsou prechody mezi akt/neakt vymahanim
  ï éaloby: Vym·h·nÌ ukonËeno, Doplaceno p¯ed pod·nÌm ûaloby
  ï ExekuËnÌ ¯ÌzenÌ: Vym·h·nÌ ukonËeno
  ï Trest: P¯Ìpad odloûen, ÿÌzenÌ zastaveno, PM Rozsudek s doloûkou, PM trestnÌ p¯Ìkaz s doloûkou
  ï DÏdictvÌ: vym·h·nÌ ukonËeno
  ï Insolvence: Nep¯ihl·öeno, P¯ihl·öka odmÌtnuta, Iÿ skonËeno ñ vym·h·nÌ ukonËeno, Iÿ skonËeno ñ podat ûalobu (proces probÌh· v agendÏ nalÈzacÌho ¯ÌzenÌ)
  */

  ----------------HU/HR----------------------
  execute immediate ('truncate table update_person_status_lp drop storage');

  insert into update_person_status_lp (id_person,
                                       skp_person_status,
                                       dtime_valid_from)
    with akt_stav as
     (
          --vybereme aktualni stavy ze vsech rizeni
          select e.id_person,
                  e.skp_epr_status as status,
                  e.dtime_valid_from,
                  e.dtime_valid_to,
                  'epr' as rizeni
            from epr_status_data e
           where flag_deleted = 'N'
             and dtime_valid_to > sysdate
      union
          select ex.id_person,
                  ex.skp_exe_status as status,
                  ex.dtime_valid_from + 1 / (60 * 60 * 24),
                  ex.dtime_valid_to,
                  'exe' as rizeni
            from exe_status_data ex
           where flag_deleted = 'N'
             and dtime_valid_to > sysdate
      union
          select ded.id_person,
                  ded.skp_ded_status as status,
                  ded.dtime_valid_from,
                  ded.dtime_valid_to,
                  'ded' as rizeni
            from ded_status_data ded
           where flag_deleted = 'N'
             and dtime_valid_to > sysdate
      union
          select trest.id_person,
                  trest.skp_trest_status as status,
                  trest.dtime_valid_from,
                  trest.dtime_valid_to,
                  'trest' as rizeni
            from trest_status_data trest
           where flag_deleted = 'N'
             and dtime_valid_to > sysdate
      union
          select id_person,
                  skp_ins_status as status,
                  dtime_valid_from,
                  dtime_valid_to,
                  'ins' as rizeni
            from ins_status_data ins
           where flag_deleted = 'N'
             and dtime_valid_to > SYSDATE
      )
    
    select a.id_person, pps.skp_person_status, a.dtime_valid_from
      from akt_stav a
      join (select max(dtime_valid_from) as d, id_person
              from akt_stav
             group by id_person) aa
        on aa.id_person = a.id_person
       and aa.d = a.dtime_valid_from --najdeme akt. stav
      left join person_status_data ps
        on ps.id_person = a.id_person
       and ps.flag_valid = 'Y'
       and ps.flag_deleted = 'N'
       and ps.flag_collected = 'N'
      join person_status_data pps
        on pps.id_person = a.id_person
       and pps.flag_valid = 'Y'
       and pps.flag_deleted = 'N'
       and pps.flag_collected = 'Y'
       and pps.skp_person_status in (1, 2) -- aby to byli klienti HU nebo HYREF
     where ((rizeni = 'epr' and status in (63, 65)) or
           (rizeni = 'exe' and status in (20)) or
           (rizeni = 'ded' and status in (11, 12, 16, 15)) or
           (rizeni = 'trest' and status in (22)) or
           (rizeni = 'ins' and status in (4, 7, 15, 14))) --najdeme stav ukoncujici vymahani
       and ps.id_person is null --nema vymahani ukonceno
    ;

  --v tabulce uokncime interval aktivniho vymahani
  update person_status_data a
     set a.dtime_valid_to    =
         (select distinct dtime_valid_from
            from update_person_status_lp b
           where a.id_person = b.id_person),
         a.flag_valid         = 'N',
         a.dtime_modification = sysdate
   where exists (select 1
            from update_person_status_lp h
           where h.id_person = a.id_person
             and h.skp_person_status = a.skp_person_status)
     and a.flag_valid = 'Y';

  --vlozime novy interval, ve kterem neni vymahani
  insert into person_status_data
    select rownum +
           (select max(skp_person_status_data) from person_status_data),
           id_person,
           skp_person_status,
           'N',
           dtime_valid_from + 1 / (60 * 60 * 24),
           DATE'3000-01-01',
           sysdate,
           'Y',
           'N',
           'sys'
      from update_person_status_lp;

  ----------------DEBET----------------------
  --nyni update doplacenych pro debetare, podivame se nejdriv na ty, co jsou collected a doplatili

  execute immediate ('truncate table update_person_status_dv drop storage');

  insert into update_person_status_dv (id_person)
    select e.id_person
      from person_status_data e
      join (select sum(amt_balance) amt_min, id_person
              from ucet
             group by id_person
            having sum(amt_balance) >= 0) a
        on a.id_person = e.id_person
     where e.skp_person_status = 3
       and e.flag_deleted = 'N'
       and e.flag_collected = 'Y'
       and flag_valid = 'Y';

  -- u debetaru, kteri jiz nemaji dluh, zmenime platnost intervalu vymahani
  update person_status_data
     set dtime_valid_to     = trunc(sysdate) - 1 / (60 * 60 * 24),
         dtime_modification = sysdate,
         flag_valid         = 'N'
   where id_person in (select * from update_person_status_dv)
     and flag_valid = 'Y';

  ---- u debetaru, kteri jiz nemaji dluh, nahrajeme novy radek o tom, ze jiz nejsou vymahani
-- MM: 19.5.2017 ˙prava - mnoûenÌ v tabulce  person_status_data RISK-617, p¯id·n distinct a with konstrukce

 insert into person_status_data (skp_person_status_data,
                                 id_person,
                                 skp_person_status,
                                 flag_collected,
                                 dtime_valid_from,
                                 dtime_valid_to,
                                 dtime_modification,
                                 flag_valid,
                                 flag_deleted,
                                 employee)
    with base as (
    select distinct 
           id_person,
           skp_person_status,
           'N' as flag_collected ,
           trunc(sysdate) as dtime_valid_from ,
           DATE'3000-01-01' as dtime_valid_to ,
           sysdate as dtime_modification ,
           'Y' as flag_valid ,
           'N' as flag_deleted ,
           'sys' as employee 
      --select count(*)
      from person_status_data e
     where e.id_person in (select a.id_person from update_person_status_dv a)
     )
     
     select rownum +
           (select max(skp_person_status_data) from person_status_data) as skp_person_status_data ,
            b.* 
     from base b;

  --u debetaru musime udelat i zmenu dopl - > vymahany, za jakych podminek?
  --podivame se do lcs tabulek a pokud bude klient v late po tom, co bylo u nas vymahani ukonceno

  --po nasazeni 1.3.4. LCS je potreba zkontrolvat, jestli tento select stale odpovida!!

execute immediate ('truncate table update_person_status_dn drop storage');

  insert into update_person_status_dn (id_person)

    with lcs_late as
     (select distinct sr.id_person
        from owner_int.v_lcs_deal d
        left join owner_int.v_lcs_classifiable c
          on c.t_deal_1 = d.t_deal_id
        left join owner_int.v_lcs_milestone m
          on m.t_milestone_id = c.t_milestone_1
        join ucet sr
          on sr.cislo_uctu = d.contractnr
       where d.dtype = 'DEBIT'
         and m.name like '%Late%'
         and c.ctype = 'DEAL'
         and amt_balance < 0 --podminka, aby to byli klienti v Late v LCS a v Jez. meli na uctu zaporny zustatek
      )
    select e.id_person
      from person_status_data e
      join lcs_late l
        on l.id_person = e.id_person
     where flag_valid = 'Y'
       and flag_collected = 'N'
       and skp_person_status = 3
       and flag_deleted = 'N'
    
    ;
  --ukoncime interval not collected u debetaru
  update person_status_data
     set dtime_valid_to     = trunc(sysdate) - 1 / (60 * 60 * 24),
         flag_valid         = 'N',
         dtime_modification = sysdate
   where id_person in
         (select id_person from update_person_status_dn)
     and flag_valid = 'Y';

 -- MM: 19.5.2017 ˙prava - mnoûenÌ v tabulce  person_status_data RISK-617, p¯id·n distinct a with konstrukce
    insert into person_status_data
      with base as (
    select distinct
           id_person,
           3 as skp_person_status,
           'Y' as flag_collected,
           trunc(sysdate) as dtime_valid_from,
           DATE'3000-01-01' as dtime_valid_to,
           sysdate as dtime_modification,
           'Y' as flag_valid ,
           'N' as flag_deleted ,
           'sys' as employee
      from update_person_status_dn)
      select rownum +
           (select max(skp_person_status_data) from person_status_data) as skp_person_status_data , b.* 
       from base b;

commit;

end;


-------------------------------------------------------------------------------------
-- Generov·nÌ obecn˝ch upozornÏnÌ, kterÈ navz·jem p¯Ìliö nesouvisÌ.
-------------------------------------------------------------------------------------
procedure insert_notices IS
BEGIN


--docasna kontrola na ukonceni RS, viz RISK-280
execute immediate ('truncate table update_contr_term drop storage');

insert into update_contr_term (text_contract_number,
                               dtime_contract_termination,
                               id_person)
select distinct t.TEXT_CONTRACT_NUMBER,
                t.DTIME_CONTRACT_TERMINATION,
                p.id_person
  from owner_dwh.f_contract_at t
  join APP_COLL_JEZEVCIK_TEST.person p
    on p.skp_person = t.skp_person and p.flag_valid = 'Y'
  join owner_dwh.f_loan_balance_at l
    on l.skp_person = p.skp_person
   and l.FLAG_DELETED = 'N'
   and l.FLAG_VALID = 'Y'
  left join notice n
    on n.id_person = p.ID_PERSON
   and n.text_notice like '%Klient si zruöil RS%'
 where l.amt_payoff > 0
   and t.DTIME_CONTRACT_TERMINATION < sysdate
   AND T.FLAG_DELETED = 'N'
   and t.DTIME_CONTRACT_TERMINATION + 14 > sysdate
   and n.id_person is null;


insert into notice t (t.skp_notice, t.id_person, t.text_notice, t.dtime_notice_created, t.dtime_notice_deadline, t.code_status, t.employee)
  select (select max(skp_notice) from notice) + rownum,
         id_person,
         'Klient si zruöil RS dne ' || DTIME_CONTRACT_TERMINATION || '.',
         sysdate,
         trunc(sysdate),
         'a',
         'sys'
    from update_contr_term;
  execute immediate ('truncate table update_contr_term drop storage');

commit;



EXECUTE IMMEDIATE 'truncate table APP_COLL_JEZEVCIK_TEST.terminated_bu_notice drop storage';

  insert INTO 
  APP_COLL_JEZEVCIK_TEST.terminated_bu_notice (ID_PERSON,
                                               DATE_BANK_ACCOUNT_TERMINATION,
                                               TEXT_NOTICE)
    with payoff_base as
     ( -- Baze zesplatnenych klientu a debetaru v jezevcikovi
      select a.id_person, c.name_person_status
      -- select count(*), count(distinct a.id_person)
        from APP_COLL_JEZEVCIK_TEST.person a
        join APP_COLL_JEZEVCIK_TEST.person_status_data b
          on a.id_person = b.id_person
         and b.flag_valid = 'Y' and a.flag_valid = 'Y'
        join APP_COLL_JEZEVCIK_TEST.person_status c
          on b.skp_person_status = c.skp_person_status
       where a.flag_principal_paid_obs = 'N'
         and a.flag_interest_paid_obs = 'N')
    
    ,
    loan_base as
     ( -- Napojim se na aktivni uvery
      select a.id_person,
              a.name_person_status,
              l.SKP_LOAN,
              l.SKP_BANK_ACCOUNT as BANK_ACCOUNT_LOAN
        from payoff_base a
        JOIN APP_COLL_JEZEVCIK_TEST.person prso 
          ON prso.id_person = a.id_person
        join owner_dwh.f_loan_at l
          on prso.skp_person = l.SKP_PERSON
         and a.name_person_status in ('HU', 'HYREF')
         and l.DATE_LOAN_UTILIZATION <> date
       '3000-01-01' /*nacerpany*/
         and l.DATE_LOAN_CLOSE = date '3000-01-01' -- nechci ukoncene uvery
        join owner_dwh.cl_loan_status cl
          on l.SKP_LOAN_STATUS = cl.SKP_LOAN_STATUS
         and cl.NAME_LOAN_STATUS = 'AktivnÌ')
    
    ,
    loan_account_terminations as
     ( -- O aktivnich uverech zhora uvazuji, jestli byl BU ze ktereho jej splaceli zrusen
      select dba.SKP_BANK_ACCOUNT,
              lb.ID_PERSON,
              lb.name_person_status,
              lb.skp_loan,
              dba.DATE_BANK_ACCOUNT_TERMINATION,
              dba.TEXT_BANK_ACCOUNT_NUMBER
        from loan_base lb
        join owner_dwh.DC_BANK_ACCOUNT dba
          on dba.SKP_BANK_ACCOUNT = lb.BANK_ACCOUNT_LOAN -- chci jen ty ucty, ktere jsou navazany na splatku loanu
         and dba.SKP_BANK_ACCOUNT_TYPE = 198 -- BU
         and dba.DATE_BANK_ACCOUNT_TERMINATION <> date
       '3000-01-01'
         and dba.CODE_ACCOUNT_TERMIN_REASON in
             ('CLIENTNOTICE', 'CLIENTWITHDRAWAL'))
    
    ,
    debet_base as
     ( -- Ucty, na kterych jsou klienti v debetu
      select a.id_person,
              a.name_person_status,
              bt.SKP_BANK_ACCOUNT as BANK_ACCOUNT_DEBET,
              bt.AMT_BALANCE_AVAILABLE,
              bt.AMT_GOVERNMENT_HOLD_SUM
        from payoff_base a
        JOIN APP_COLL_JEZEVCIK_TEST.person prso 
          ON prso.id_person = a.id_person
        join owner_dwh.F_ACCOUNT_BALANCE_TD bt
          on prso.skp_person = bt.skp_person
         and a.name_person_status = 'DEBET'
         and bt.SKP_BANK_ACCOUNT_TYPE = 198 -- BU
         and trunc(sysdate) between bt.DATE_BALANCE_VALID_FROM and
             bt.DATE_BALANCE_VALID_TO
         and bt.AMT_BALANCE_AVAILABLE + bt.AMT_GOVERNMENT_HOLD_SUM < 0)
    
    ,
    bank_account_terminations as
     ( -- Zrusene debetni ucty
      select dba.SKP_BANK_ACCOUNT,
              db.ID_PERSON,
              db.name_person_status,
              null as skp_loan,
              dba.DATE_BANK_ACCOUNT_TERMINATION,
              dba.TEXT_BANK_ACCOUNT_NUMBER
        from debet_base db
        join owner_dwh.DC_BANK_ACCOUNT dba
          on dba.SKP_BANK_ACCOUNT = db.BANK_ACCOUNT_DEBET -- chci jen ty ucty, ktere jsou navazany na debetni ucet
         and dba.SKP_BANK_ACCOUNT_TYPE = 198 -- BU
         and dba.DATE_BANK_ACCOUNT_TERMINATION <> date
       '3000-01-01'
         and dba.CODE_ACCOUNT_TERMIN_REASON in
             ('CLIENTNOTICE', 'CLIENTWITHDRAWAL'))
    
    ,
    prepared as
     ( -- Spojim dohromady
      select *
        from loan_account_terminations
      union all
      select *
        from bank_account_terminations)
        
        
    -- Na zaver vypisu ve forme poznamky
    select distinct id_person,
                    date_bank_account_termination,
                    case
                      when name_person_status = 'HU' then
                       'Klient ' || trunc(date_bank_account_termination) ||
                       ' zruöil sv˘j B⁄ ËÌslo: ' || text_bank_account_number ||
                       ', ze kterÈho spl·cel HU'
                      when name_person_status = 'HYREF' then
                       'Klient ' || trunc(date_bank_account_termination) ||
                       ' zruöil sv˘j B⁄ ËÌslo: ' || text_bank_account_number ||
                       ', ze kterÈho spl·cel HR'
                      when name_person_status = 'DEBET' then
                       'Klient ' || trunc(date_bank_account_termination) ||
                       ' zruöil sv˘j B⁄ ËÌslo: ' || text_bank_account_number ||
                       ', na kterÈm byl v debetu'
                    end as text_notice
      from prepared
     where trunc(date_bank_account_termination) >= date '2016-10-31' -- Toto je datum, od ktereho chci vest upozorneni v jezevcikovi, predtim bylo manualne vylito do excelu
    ;

  -- Vlozim do tabulek s upozornenimi
  insert into APP_COLL_JEZEVCIK_TEST.notice t (t.skp_notice, t.id_person, t.text_notice, t.dtime_notice_created, t.dtime_notice_deadline, t.code_status, t.employee)
    select rownum +
           (select max(skp_notice) from APP_COLL_JEZEVCIK_TEST.notice) as skp_notice,
           a.id_person,
           a.text_notice,
           sysdate as dtime_notice_created,
           trunc(sysdate) as dtime_notice_deadline,
           'a' as code_status,
           'sys' as employee
      from APP_COLL_JEZEVCIK_TEST.terminated_bu_notice a
      left join APP_COLL_JEZEVCIK_TEST.notice old
        on a.id_person = old.id_person
       and old.text_notice LIKE '%' || a.text_notice || '%'  -- vymahaci si obcas pripisou poznamku do textu UPO (na zacatek nebo konec) -- Jez.v.5.3
     where old.id_person is null -- vezmu jenom ty zaznamy, kde ten zaznam jeste nebyl
    ;

  COMMIT;
  
  -- RISK-523  -- nove upozorneni - insolvence
  -- navazuje RISK-784
  -- vytvo¯enÌ novÈho upozornÏnÌ, kterÈ bude vyskakovat pro zruseni konkurzu a odejmuti osvobozeni
  
  
  --RISK-985
  insert into APP_COLL_JEZEVCIK_TEST.notice t (t.skp_notice, t.id_person, t.text_notice, t.dtime_notice_created, t.dtime_notice_deadline, t.code_status, t.employee)
    select rownum +
           (select max(skp_notice) from APP_COLL_JEZEVCIK_TEST.notice) as skp_notice,
           prso.id_person,
           CASE ci.code_insolvency_case_item_type 
             WHEN '95' THEN 'ZruöenÌ konkurzu. Publikov·no ' || to_char(trunc(ci.DTIME_PUBLICATION), 'dd.mm.yyyy')
             WHEN '179' THEN 'OdejmutÌ osvobozenÌ. Publikov·no ' || to_char(trunc(ci.DTIME_PUBLICATION), 'dd.mm.yyyy')
             WHEN '177' THEN 'SplnÏnÌ oddluûenÌ. Publikov·no ' || to_char(trunc(ci.DTIME_PUBLICATION), 'dd.mm.yyyy')
           END AS text_notice,
           SYSDATE as dtime_notice_created,
           trunc(SYSDATE) as dtime_notice_deadline,
           'a' as code_status,
           'sys' as employee
      from APP_COLL_JEZEVCIK_TEST.INSOL_LIST il
      JOIN APP_COLL_JEZEVCIK_TEST.person prso 
        ON prso.skp_person = il.skp_person
      JOIN owner_dwh.dc_insolvency_case ic
        ON ic.TEXT_CASE_IDENT_NUMBER = il.ic_spisova_znacka
      JOIN owner_dwh.f_insolvency_case_item_at ci
        ON ci.SKP_INSOLVENCY_CASE = ic.SKP_INSOLVENCY_CASE
      left join APP_COLL_JEZEVCIK_TEST.notice old
        on prso.id_person = old.id_person
       and old.text_notice LIKE CASE ci.code_insolvency_case_item_type 
                               WHEN '95' THEN 'ZruöenÌ konkurzu%'
                               WHEN '179' THEN 'OdejmutÌ osvobozenÌ%'
                               WHEN '177' THEN 'SplnÏnÌ oddluûenÌ%'
                             END
     where ci.code_insolvency_case_item_type in ('95','179') 
       AND ci.DTIME_PUBLICATION > '21.7.2017' -- Toto je datum, od ktereho chci vest upozorneni v jezevcikovi, predtim bylo manualne vylito do excelu
       AND ci.FLAG_DELETED = 'N'
       AND old.id_person is null -- vezmu jenom ty zaznamy, kde ten zaznam jeste nebyl
    ;
    
    COMMIT;
    
    
    -- RISK-857 - SonÌk: automatickÈ vytvo¯enÌ karty klienta v trestu
    -- klienti kterym se zmenil stav datove zpravy a pritom nejsou v Jezevcikovi, se pri nejblizsim loadu dat nacitaji do person, a pak nasledujici select vlozi stav
    -- (viz trigger sonik_dm_status_insert)
    INSERT INTO APP_COLL_JEZEVCIK_TEST.trest_status_data (skp_trest_status_data, skp_trest_status, id_person, dtime_valid_from, dtime_valid_to, text_note,
                                                                 flag_deleted, dtime_modification, employee)
    SELECT (SELECT MAX(skp_trest_status_data) FROM APP_COLL_JEZEVCIK_TEST.trest_status_data) + ROWNUM,
            ad.skp_trest_status,
            ad.id_person,
            ad.dtime_status_valid_from, -- a nebo ad.dtime_inserted ?
            LAG (ad.dtime_status_valid_from - 1/(24*60*60), 1, DATE'3000-01-01') OVER (PARTITION BY ad.id_person ORDER BY ad.dtime_status_valid_from DESC) AS valid_to,
            '',
            'N',
            SYSDATE,
            'sys'
    FROM APP_COLL_JEZEVCIK_TEST.trest_status_addition ad
    JOIN APP_COLL_JEZEVCIK_TEST.person p
      ON p.id_person = ad.id_person
    JOIN APP_COLL_JEZEVCIK_TEST.cuid_addition c
      ON c.cuid = p.cuid
    WHERE ad.flag_addition = 'N'
    ;
    
    -- nastavime flag_addition
    UPDATE APP_COLL_JEZEVCIK_TEST.trest_status_addition adt
    SET flag_addition = 'Y'
    WHERE EXISTS 
            ( SELECT ad.skp_trest_status_addition    
                FROM APP_COLL_JEZEVCIK_TEST.trest_status_addition ad
                JOIN APP_COLL_JEZEVCIK_TEST.person p
                  ON p.id_person = ad.id_person
                JOIN APP_COLL_JEZEVCIK_TEST.cuid_addition c
                  ON c.cuid = p.cuid
                WHERE ad.flag_addition = 'N'
                  AND adt.skp_trest_status_addition = ad.skp_trest_status_addition
             )
    ;
    
    COMMIT;




    -- FA -- RISK-722  - SonÌk - propis zpÏtnÏ do Trestu p¯i aut. upozornÏnÌ p¯i zmÏnÏ stavu
    execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.clients_trest_help');

INSERT INTO clients_trest_help (skp_data_message,
                                id_person,
                                name_data_message_status,
                                dtime_valid_from,
                                dtime_valid_to,
                                data_message_number)
      SELECT dm.skp_data_message
           , tr.id_person
           , tt.name_data_message_status
           , t.dtime_valid_from
           , t.dtime_valid_to
           , dm.data_message_number
      from APP_COLL_JEZEVCIK_TEST.SONIK_DATA_MESSAGE_STATUS t
      JOIN APP_COLL_JEZEVCIK_TEST.sonik_cl_data_message_status tt
        ON t.skp_data_message_status = tt.skp_data_message_status
        
      JOIN APP_COLL_JEZEVCIK_TEST.sonik_data_message dm  -- skp_data_message
        ON t.skp_data_message = dm.skp_data_message
      JOIN APP_COLL_JEZEVCIK_TEST.sonik_confirm_pair son -- PAIR
        ON son.skp_data_message = dm.skp_data_message
      JOIN APP_COLL_JEZEVCIK_TEST.person p
        ON p.skp_person = son.skp_person AND p.flag_valid = 'Y'
      JOIN APP_COLL_JEZEVCIK_TEST.v_trest tr      -- id_person v trestu
        ON tr.id_person = p.id_person
      LEFT JOIN APP_COLL_JEZEVCIK_TEST.trest_note tn
             ON tn.id_person = p.id_person
             AND tn.text_note LIKE '%' || dm.data_message_number || '%'
             AND tn.text_note LIKE '%' || tt.name_data_message_status || '%'
             AND tn.text_note LIKE '%p¯id·n nov˝ stav%'

      WHERE tn.id_person IS NULL -- vkladame iba nove zaznamy
        AND nvl(tr.dtime_valid_to, DATE'3000-01-01') > SYSDATE
        AND son.name_confirmation_status = 'Potvrzeno'     
        AND (t.flag_valid = 'Y'
            OR t.dtime_valid_from > SYSDATE - 14)       -- podminka kvuli klientum, kteri v momente zadani stavu jeste nejsou v jezevcikovi (budou zitra). Tech dalsich 13 dni je Bulharska rezerva.

        AND tt.name_data_message_status IN ('P¯ipojenÌ se ökodou', 'VyrozumÏnÌ - prostudov·nÌ spisu', 'Zah·jenÌ trestnÌho stÌh·nÌ', 'P¯Ìpad odloûen', 'Pod·na ûaloba'
                                            ,'VyrozumÏnÌ o pod·nÌ obûaloby', 'N·vrh na potrest·nÌ', 'ÿÌzenÌ zastaveno', 'PodmÌnÏnÈ zastavenÌ Tÿ', 'VyrozumenÌ - nenÌ T»'
                                            , 'OsvÏdËenÌ/NeosvÏdËenÌ')
    ;

    ------ 1) poznamka pridan novy stav

    INSERT INTO APP_COLL_JEZEVCIK_TEST.trest_note (skp_trest_note, skp_trest_note_type, id_person, skp_trest_status_data
                                                          , text_note, dtime_note_created, employee)
      SELECT rownum + (select max(skp_trest_note) from trest_note)
             , 3 -- Poznamka
             , a.id_person
             , NULL -- ziadny stav
             , 'SONÕK: (ËÌslo DZ:' || a.data_message_number || ') p¯id·n nov˝ stav: ' || a.name_data_message_status || '. Platnost od ' || a.dtime_valid_from || 
                 CASE WHEN a.dtime_valid_to = to_date('1.1.3000') THEN '.'
                      ELSE ' do ' || a.dtime_valid_to || '.'
                 END
             , SYSDATE
             , 'sys'
      from APP_COLL_JEZEVCIK_TEST.Clients_Trest_Help a
    ;


    ------ 2) poznamka nove upozorneni

    INSERT INTO APP_COLL_JEZEVCIK_TEST.trest_note (skp_trest_note, skp_trest_note_type, id_person, skp_trest_status_data, text_note, dtime_note_created, employee)
      SELECT rownum + (select max(skp_trest_note) from trest_note)
             , 4 -- Upozorneni
             , a.id_person
             , NULL -- ziadny stav
             , 'SONÕK: (ËÌslo DZ:' || a.data_message_number || ') p¯id·no upozornÏnÌ: ' || a.name_data_message_status || '. Platnost od ' || a.dtime_valid_from || 
                 CASE WHEN a.dtime_valid_to = to_date('1.1.3000') THEN ''
                      ELSE ' do ' || a.dtime_valid_to || ''
                 END
                 || ' (deadline: ' || TRUNC(SYSDATE) || ').'
             , SYSDATE
             , 'sys'
      from APP_COLL_JEZEVCIK_TEST.Clients_Trest_Help a
    ;


    ------ 3) nove upozorneni

    INSERT INTO APP_COLL_JEZEVCIK_TEST.notice (skp_notice, id_person, text_notice, dtime_notice_created, dtime_notice_deadline, code_status, employee)
    SELECT ROWNUM + (SELECT MAX(skp_notice) FROM APP_COLL_JEZEVCIK_TEST.notice)
         , a.id_person
         -- SonÌk: (ËÌslo DZ:495983481) p¯id·n nov˝ stav: P¯ipojenÌ se ökodou. Platnost od 14.09.2017 15:24:58 (deadline:  14.9.2017).
         , 'SONÕK: (ËÌslo DZ:' || a.data_message_number || ') p¯id·n nov˝ stav: ' || a.name_data_message_status || '. Platnost od ' || a.dtime_valid_from || 
             CASE WHEN a.dtime_valid_to = to_date('1.1.3000') THEN ''
                  ELSE ' do ' || a.dtime_valid_to || ''
             END
             || ' (deadline: ' || TRUNC(SYSDATE) || ').'
         , SYSDATE
         , trunc(SYSDATE)
         , 'a'
         , 'sys'
    FROM APP_COLL_JEZEVCIK_TEST.Clients_Trest_Help a
    ;
    
    COMMIT;
    -- delete FROM clients_trest_help;
    
    
    
    
    -- FA -- RISK-705  - V˝öe ökody v SonÌku do pozn·mek JezevËÌka v Trestu
    -- N·sleduje po RISK-857 a RISK-722, jinak se pozn·mka s v˝öÌ ökody propsala aû den po p¯id·nÌ do Trestu (tedy dva dny po zad·nÌ v SonÌku)
    
     INSERT INTO APP_COLL_JEZEVCIK_TEST.trest_note (skp_trest_note, skp_trest_note_type, id_person, /*skp_trest_status_data,*/ text_note, dtime_note_created, employee)
      SELECT rownum + (select max(skp_trest_note) from trest_note)
           , 3 -- Poznamka
           , p.id_person
           --, NULL -- ziadny stav
           , 'SONÕK: ' || sd.text_note || ' »Ìslo datovÈ zpr·vy: ' || sdm.data_message_number || '. V˝öka ökody: ' || sd.amt_damage || ' KË. Platnost od ' || nvl(sd.dtime_modified, sd.dtime_inserted) -- platnost od sysdate?
           , SYSDATE
           , 'sys'
      FROM APP_COLL_JEZEVCIK_TEST.sonik_damage sd
      JOIN APP_COLL_JEZEVCIK_TEST.sonik_data_message sdm
        ON sdm.skp_data_message = sd.skp_data_message
      JOIN APP_COLL_JEZEVCIK_TEST.person p
        ON p.skp_person = sd.skp_person AND p.flag_valid = 'Y'
      JOIN APP_COLL_JEZEVCIK_TEST.v_trest tr
        ON tr.id_person = p.id_person
      LEFT JOIN APP_COLL_JEZEVCIK_TEST.trest_note tn           -- klic je kombinace id_person, cislo datove zpravy a vyska skody
             ON tn.id_person = p.id_person
             AND tn.text_note LIKE '%' || sdm.data_message_number || '%'
             AND tn.text_note LIKE '%V˝öka ökody: ' || sd.amt_damage || '%'
      WHERE nvl(tr.dtime_valid_to, DATE'3000-01-01') > SYSDATE
        AND tn.id_person IS NULL                                     -- vkladame iba nove zaznamy
      ;
    
    COMMIT;
    
    
    -- FA -- RISK-713  - Kontrola nov˝ch klient˘ v Trestu jestli jsou v SonÌku
    -- pozor, prekryva sa to s upozornenim na vysku skody (RISK-705). Preto najprv generovat vysky skody (su specifickejsie) a az potom toto.
    INSERT INTO APP_COLL_JEZEVCIK_TEST.trest_note (skp_trest_note, skp_trest_note_type, id_person, /*skp_trest_status_data,*/ text_note, dtime_note_created, employee)
      SELECT rownum + (select max(skp_trest_note) from trest_note)
           , 3 -- Poznamka
           , tr.id_person
           --, NULL -- ziadny stav
           , 'SONÕK: Klient se nach·zÌ v SonÌku, ËÌslo datovÈ zpr·vy: ' || dm.data_message_number || '. Platnost od ' || son.dtime_inserted
           , SYSDATE
           , 'sys'
      from APP_COLL_JEZEVCIK_TEST.v_trest tr
      JOIN APP_COLL_JEZEVCIK_TEST.person p
        ON p.id_person = tr.id_person AND p.flag_valid = 'Y'
      JOIN APP_COLL_JEZEVCIK_TEST.sonik_confirm_pair son
        ON p.skp_person = son.skp_person
      JOIN APP_COLL_JEZEVCIK_TEST.sonik_data_message dm
        ON son.skp_data_message = dm.skp_data_message AND dm.skp_data_message > 0 -- kv˘li fake datovce pro klienty z JezevËÌka, kter· m· ËÌslo -1
      LEFT JOIN APP_COLL_JEZEVCIK_TEST.trest_note tn            -- klic je kombinace id_person a cislo datove zpravy
             ON tn.id_person = p.id_person
             AND tn.text_note LIKE '%' || dm.data_message_number || '%'
      WHERE tn.id_person IS NULL                                      -- vkladame iba nove zaznamy
        AND nvl(tr.dtime_valid_to, DATE'3000-01-01') > SYSDATE
        AND son.name_confirmation_status = 'Potvrzeno'
      ;

    
    -- RISK-776
    -- automatickÈ upozornÏnÌ po zad·nÌ stavu ParkoviötÏ
    INSERT INTO APP_COLL_JEZEVCIK_TEST.notice (skp_notice, id_person, text_notice, dtime_notice_created, dtime_notice_deadline, code_status, employee)
      SELECT ROWNUM + (SELECT MAX(skp_notice) FROM APP_COLL_JEZEVCIK_TEST.notice)
           , kl.id_person
           , 'ParkoviötÏ: ' || kl.name || '. Zaparkov·n od ' || kl.dtime_valid_from || '.'
           , SYSDATE
           , CASE WHEN to_char(kl.dtime_valid_from, 'MM') < 7 THEN to_date('1.7.' || to_char(kl.dtime_valid_from /*SYSDATE*/, 'YYYY'), 'fmdd.mm.yyyy')
                  ELSE to_date('1.1.' || to_char(add_months(kl.dtime_valid_from /*SYSDATE*/,6), 'YYYY'), 'fmdd.mm.yyyy')
             END -- upo chceme 1.1. nebo 1.7. - co je blize
           , 'a'
           , 'sys'
      FROM ( -- platni klienti s platnym stavem Parkoviste kteri jeste nemaji upozorneni
            SELECT p.cuid, p.name, sd.id_person, sd.dtime_valid_from 
              FROM APP_COLL_JEZEVCIK_TEST.exe_status_data sd
              JOIN APP_COLL_JEZEVCIK_TEST.exe_status a 
                ON a.skp_exe_status = sd.skp_exe_status
              JOIN APP_COLL_JEZEVCIK_TEST.person p
                ON p.id_person = sd.id_person
              LEFT JOIN APP_COLL_JEZEVCIK_TEST.notice n
                     ON n.id_person = sd.id_person
                     AND n.text_notice LIKE '%ParkoviötÏ%'
                     AND n.text_notice LIKE '%Zaparkov·n od%'
                     AND n.text_notice LIKE '%' || sd.dtime_valid_from || '%'
              WHERE a.name_exe_status LIKE 'ParkoviötÏ' 
                AND a.dtime_valid_to > SYSDATE
                AND sd.flag_deleted = 'N'
                AND sd.dtime_valid_to > SYSDATE
                AND p.flag_valid = 'Y'
                AND n.id_person IS NULL                               -- pouze ty co jeste upozorneni nemaji
            ) kl -- klienti v parkovisti
    ;
    
    COMMIT;

    -- RISK-778 (Evidov·nÌ öance na kartÏ klienta)
    -- V p¯ÌpadÏ, kdy je na klientovi uloûen jeden ze 3 uveden˝ch stav˘ a z·roveÚ neexistuje na klientovi zadan· öance s typem SK urËen˝ soudem,
    --   vyskoËÌ v jezevËÌkovi upozornÏnÌ (klidnÏ aû na druh˝ den) s pozn·mkou "u klienta nenÌ zaevidovan˝ SK"
    INSERT INTO APP_COLL_JEZEVCIK_TEST.notice (skp_notice, id_person, text_notice, dtime_notice_created, dtime_notice_deadline, code_status, employee)
      (SELECT ROWNUM + (SELECT MAX(skp_notice) FROM APP_COLL_JEZEVCIK_TEST.notice),
              sd.id_person, 
              'u klienta nenÌ zaevidovan˝ SK', 
              SYSDATE, 
              trunc(SYSDATE), 
              'a', 
              'sys'
          FROM APP_COLL_JEZEVCIK_TEST.epr_status_data sd
          JOIN APP_COLL_JEZEVCIK_TEST.epr_status st 
            ON sd.skp_epr_status = st.skp_epr_status
          LEFT JOIN (select *
                       from APP_COLL_JEZEVCIK_TEST.chance_import 
                       where TYP_SK = 'SK urËen˝ soudem' 
                         and flag_deleted = 'N'
                     ) chance
                 ON chance.id_person = sd.id_person
          WHERE st.name_epr_status IN ('PM rozsudek ñ SK s doloûkou', 'PM Rozsudek - SK Ë·steËnÏ zamÌtnuto s doloûkou', 'PM UsnesenÌ o schv·lenÌ smÌru s doloûkou')
            AND st.dtime_valid_to > SYSDATE
            AND sd.dtime_valid_to > SYSDATE
            AND sd.flag_deleted = 'N'
            AND chance.id_person IS NULL
            AND NOT EXISTS (SELECT * -- RISK-930, nechceme aby upozorneni vyskakovalo opakovane
                            FROM APP_COLL_JEZEVCIK_TEST.notice x
                            WHERE x.id_person = sd.id_person
                              AND x.text_notice LIKE '%u klienta nenÌ zaevidovan˝ SK%'
                              AND x.code_status <> 's'
                              AND x.dtime_notice_created > SYSDATE - 60
                           )
          )
    ;
    
    COMMIT;


end;



-------------------------------------------------------------------------------------
-- cÌlem je pro klienty vhodnÈ k ûalobÏ nebo insolvenËnÌ p¯ihl·öce mÌt kategorizaci debetu na jednotlivÈ poloûky;
-- epr_debety_all
-------------------------------------------------------------------------------------
procedure update_epr_debety IS

dpd_pridaj CONSTANT INT NOT NULL := 3; -- Bulhar -- Pripocita sa k DPD. Cakal by som 2(1 den zpozdeni dat na DWH, 1 extra den kvoli nerovnosti), ale najlepsie vysledky mam pri hodnote 3.


begin

/*
Zadavatel: FA
Datum vytvo¯enÌ: 26.7.2017
⁄Ëel: NaËÌtavanie dat pre vkl·danie debetu do EPR öablony.
RISK-391

-- pokud je klient vhodny k EPR a zaroven k IP, v EPR se nebude debet generovat spravne (pouze do data pred upadkem). Nicmene je to cajk, pokud ceka na prihlaseni tak se EPR nepodava

*/


execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.epr_vhodni_zaloba_help drop storage');
execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.ip_vhodni_klienti_help drop storage');
execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.epr_debety_trans drop storage');
execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.epr_debety_zustatky drop storage');
execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.epr_debety_tarif_help drop storage');
execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.epr_debety_tarify drop storage');
execute immediate ('truncate table APP_COLL_JEZEVCIK_TEST.epr_debety_all drop storage');


-- klientska zakladna  - k 26.7.2017  120 klientov
-- potrebujeme vsetky data iba ku klientom ktori su aktualne vhodni k zalobe

-- klienti vhodni k zalobe EPR (nizsie su insolvencni klienti)

INSERT INTO APP_COLL_JEZEVCIK_TEST.epr_vhodni_zaloba_help (id_person, skp_bank_account, dpd_b, den_pred_upadkem)
SELECT DISTINCT p.id_person
     , u.skp_bank_account
     , u.cnt_days_overdue_tol_b AS dpd_b
     , SYSDATE                                           -- stlpec kvuli insolvencim, datum DO ktereho chceme brat transakce.
FROM APP_COLL_JEZEVCIK_TEST.person p
JOIN APP_COLL_JEZEVCIK_TEST.ucet u
  ON u.id_person = p.id_person
LEFT JOIN APP_COLL_JEZEVCIK_TEST.ins_client_list i
  ON i.skp_person = p.skp_person
  WHERE p.date_death IS NULL
  AND p.flag_valid = 'Y'
  AND u.amt_balance < 0                                  -- disponibilnÌ z˘statek oËiötÏn˝ o exekuËnÌ blokace
  AND u.typ_uctu = 'B⁄'
  AND u.code_currency = 'CZK'
  AND i.dtime_modified IS NULL
  AND EXISTS (SELECT * FROM AP_COLL.SR_032_LATE_STAT1 b
              WHERE b.stav = 'Vhodn˝'
                AND p.skp_person = b.skp_person
              )
;


-- klienti vhodni k prihlaseni do insolvencniho rizeni
-- column amt_balance_pred_upadkem se vyuziva (pouze) v jezevcikovi v objekte Insolvency (function ZjistiDebetPredUpadkem())


--RISK-985
INSERT INTO APP_COLL_JEZEVCIK_TEST.ip_vhodni_klienti_help (id_person, skp_bank_account, dpd_b, den_pred_upadkem, amt_balance_pred_upadkem)
SELECT DISTINCT p.id_person
     , u.skp_bank_account
     , u.cnt_days_overdue_tol_b AS dpd_b
     , info.datum_upadku - 1 AS den_pred_upadkem
     , bal.AMT_BALANCE AS amt_balance_pred_upadkem
FROM APP_COLL_JEZEVCIK_TEST.person p
JOIN APP_COLL_JEZEVCIK_TEST.ucet u
  ON u.id_person = p.id_person
JOIN APP_COLL_JEZEVCIK_TEST.insol_list ins
  ON ins.skp_person = p.skp_person
JOIN APP_COLL_JEZEVCIK_TEST.INS_INFO info
  ON info.id_person = p.id_person
LEFT JOIN owner_dwh.f_account_balance_pd bal
       ON bal.skp_person = p.skp_person
       AND bal.SKP_BANK_ACCOUNT = u.skp_bank_account
       AND bal.DATE_EFFECTIVE = (info.datum_upadku - 1)
WHERE p.date_death IS NULL
  AND p.flag_valid = 'Y'
  AND u.amt_balance < 0                                                                  -- disponibilnÌ z˘statek oËiötÏn˝ o exekuËnÌ blokace
  AND u.typ_uctu = 'B⁄'
  AND u.code_currency = 'CZK'
  AND ins.ic_stav_rizeni IN ('ODDLUéENÕ', 'KONKURS', '⁄PADEK')
  AND info.datum_upadku IS NOT NULL
  AND EXISTS (SELECT st.name_ins_status, std.id_person                              -- vyhodime tych co maju historicky stav 'prihlaska podana'
                FROM APP_COLL_JEZEVCIK_TEST.ins_status_data std
                JOIN APP_COLL_JEZEVCIK_TEST.ins_status st
                  ON std.skp_ins_status = st.skp_ins_status
               WHERE st.name_ins_status = '»ek· na p¯ihl·öenÌ'
                 AND st.dtime_valid_to > SYSDATE
                 AND std.flag_deleted = 'N'
                 AND std.dtime_valid_to > SYSDATE
                 AND p.id_person = std.id_person
              )
;




  /* ------------------------------------  */
  /*          epr_debety_trans             */
  /* ------------------------------------  */


-- vsetky debetne transakcie ku klientom vhodnym k zalobe
-- v tabulke budu vsetky transakcie odkedy je v debetu, potom zvlast sa budeme dotazovat podla kategorie (upominky, sms, ...)
-- tarify a storna tarifov sem nedavam, tie riesi zvlast tabulka nizsie.


INSERT INTO APP_COLL_JEZEVCIK_TEST.epr_debety_trans (ID_PERSON,
                                                     SKP_BANK_ACCOUNT,
                                                     SKF_REALIZED_TRANSACTION,
                                                     AMT_TRANSACTION,
                                                     NAME_TRANSACTION_TYPE,
                                                     DTIME_TRANSACTION_ACCOUNTED,
                                                     DTIME_TRANSACTION_GENERATION,
                                                     DATE_TRANSACTION_DUE,
                                                     FLAG_TRANSACTION_DEBIT,
                                                     DPD_B,
                                                     SKP_FEE,
                                                     SKF_REALIZED_TRANS_CANCELLED,
                                                     SKF_REALIZED_TRANS_CANCELLING,
                                                     TEXT_NOTE_PART1,
                                                     FLAG_VLOZIT_DO_EPR)
SELECT vh.id_person
     , vh.skp_bank_account
     , t.SKF_REALIZED_TRANSACTION
     , t.AMT_TRANSACTION
     , tt.NAME_TRANSACTION_TYPE
     , t.DTIME_TRANSACTION_ACCOUNTED         --dtime tento je primarny pre nas do textace v jez.
     , t.DTIME_TRANSACTION_GENERATION
     , t.DATE_TRANSACTION_DUE                -- pridano v ramci RISK-734
     , t.FLAG_TRANSACTION_DEBIT         
     , vh.dpd_b
     , t.SKP_FEE
     , t.SKF_REALIZED_TRANS_CANCELLED
     , t.SKF_REALIZED_TRANS_CANCELLING
     , t.TEXT_NOTE_PART1
     , 'Y'                                    -- FLAG_vlozit_do_EPR
FROM (SELECT id_person, skp_bank_account, dpd_b, den_pred_upadkem FROM APP_COLL_JEZEVCIK_TEST.ip_vhodni_klienti_help
      UNION
      SELECT id_person, skp_bank_account, dpd_b, den_pred_upadkem FROM APP_COLL_JEZEVCIK_TEST.epr_vhodni_zaloba_help
     ) vh
JOIN APP_COLL_JEZEVCIK_TEST.person prso 
  ON prso.id_person = vh.id_person
JOIN owner_dwh.f_realized_transaction_tt t
  ON t.skp_person = prso.skp_person
  AND t.SKP_BANK_ACCOUNT = vh.skp_bank_account
JOIN owner_dwh.cl_transaction_type tt
  ON tt.SKP_TRANSACTION_TYPE = t.SKP_TRANSACTION_TYPE
  AND trunc(t.DTIME_TRANSACTION_GENERATION) > (trunc(SYSDATE) - vh.dpd_b - dpd_pridaj)        -- bulhar - nechceme transakcie z doby ked to bol dobry klient --XXXX malo by stacit 2
  AND t.DTIME_TRANSACTION_GENERATION <= trunc(vh.den_pred_upadkem)
  AND t.FLAG_DELETED = 'N'
  AND lower(tt.name_TRANSACTION_TYPE) NOT LIKE ('%tarif%')
  AND lower(t.text_note_part1) NOT LIKE '%vracÌme v·m poplatek za jeho vedenÌ%'
;




  /* ------------------------------------  */
  /*     epr_debety_zustatky       */
  /* ------------------------------------  */



-- ku kazdemu uctu (SKP_BANK_ACCOUNT) a klientovi vhodnemu pre EPR alebo IP mam jeho posledni kladny zustatek


INSERT INTO APP_COLL_JEZEVCIK_TEST.epr_debety_zustatky (SKP_BANK_ACCOUNT,
                                                        ID_PERSON,
                                                        POSLEDNI_KLADNY_ZUSTATEK,
                                                        DATE_BALANCE_VALID_FROM) -- ~ 6 min

SELECT ab.skp_bank_account
     , prso.id_person
     , ab.AMT_BALANCE POSLEDNI_KLADNY_ZUSTATEK
     , ab.DATE_PERIOD AS DATE_BALANCE_VALID_FROM
    FROM owner_dwh.f_account_balance_pd ab
    JOIN APP_COLL_JEZEVCIK_TEST.person prso 
      ON prso.skp_person = ab.skp_person
    JOIN (SELECT id_person, skp_bank_account, dpd_b FROM APP_COLL_JEZEVCIK_TEST.ip_vhodni_klienti_help
          UNION
          SELECT id_person, skp_bank_account, dpd_b FROM APP_COLL_JEZEVCIK_TEST.epr_vhodni_zaloba_help
         ) vh
      ON vh.skp_bank_account = ab.skp_bank_account
      AND prso.skp_person = ab.SKP_PERSON
    WHERE ab.DATE_PERIOD = (trunc(SYSDATE) - vh.dpd_b - dpd_pridaj) -- bulhar (pri hodnote 2 su zostatky vsetky zaporne, pre 3 su vsetky NEzaporne)
;






  /* ------------------------------------  */
  /*         epr_debety_tarify             */
  /* ------------------------------------  */
  


INSERT INTO APP_COLL_JEZEVCIK_TEST.epr_debety_tarif_help (ID_PERSON,
                                                          SKP_BANK_ACCOUNT,
                                                          SKF_REALIZED_TRANSACTION,
                                                          SKF_REQUIRED_TRANSACTION,
                                                          SKP_TRANSACTION_TYPE,
                                                          SKP_TRANSACTION_STATUS,
                                                          SKP_TRANSACTION_CLASS,
                                                          SKP_FEE,
                                                          SKP_CURRENCY,
                                                          SKF_REALIZED_TRANS_CANCELLED,
                                                          SKF_REALIZED_TRANS_CANCELLING,
                                                          TEXT_NOTE_PART1,
                                                          AMT_TRANSACTION,
                                                          DATE_EFFECTIVE,
                                                          DATE_TRANSACTION_DUE,
                                                          DTIME_TRANSACTION_GENERATION,
                                                          DTIME_REQ_TRANSACT_GENERATION,
                                                          DTIME_TRANSACTION_ACCOUNTED,
                                                          DTIME_TRANS_STATUS_VALID_FROM,
                                                          DATE_INTEREST_CALCULATION,
                                                          FLAG_TRANSACTION_DEBIT,
                                                          FLAG_CLIENT_ACCOUNT_TRN,
                                                          AMT_BALANCE)

WITH 

epr_clients AS (

      SELECT DISTINCT skp_bank_account
                    , id_person
                    , dpd_b
                    , trunc(SYSDATE) - dpd_b - dpd_pridaj + 1 as date_first_day_dpd
                    , den_pred_upadkem
      --FROM APP_COLL_JEZEVCIK_TEST.epr_vhodni_zaloba_help
      FROM (SELECT id_person, skp_bank_account, dpd_b, den_pred_upadkem FROM APP_COLL_JEZEVCIK_TEST.ip_vhodni_klienti_help
            UNION
            SELECT id_person, skp_bank_account, dpd_b, den_pred_upadkem FROM APP_COLL_JEZEVCIK_TEST.epr_vhodni_zaloba_help
           )
  )


--, transakce as ( -- mm_debety_odp_1day_trans3

select ec.id_person 
     , ec.skp_bank_account
     , rt.SKF_REALIZED_TRANSACTION
     , rt.SKF_REQUIRED_TRANSACTION 
     , rt.skp_transaction_type
     , rt.SKP_TRANSACTION_STATUS
     , rt.SKP_TRANSACTION_CLASS
     , rt.skp_fee
     , rt.SKP_CURRENCY 
     , rt.SKF_REALIZED_TRANS_CANCELLED
     , rt.SKF_REALIZED_TRANS_CANCELLING 
     , rt.TEXT_NOTE_PART1
     , rt.AMT_TRANSACTION
     , rt.DATE_EFFECTIVE
     , rt.DATE_TRANSACTION_DUE
     , rt.DTIME_TRANSACTION_GENERATION
     , rt.DTIME_REQ_TRANSACT_GENERATION
     , rt.DTIME_TRANSACTION_ACCOUNTED
     , rt.DTIME_TRANS_STATUS_VALID_FROM
     , rt.DATE_INTEREST_CALCULATION
     , rt.FLAG_TRANSACTION_DEBIT
     , rt.FLAG_CLIENT_ACCOUNT_TRN
     , bt.AMT_BALANCE

  from epr_clients ec
  join owner_dwh.f_realized_transaction_tt rt 
    on ec.skp_bank_account = rt.skp_bank_account
  join OWNER_DWH.F_ACCOUNT_BALANCE_TD bt 
    on bt.SKP_BANK_ACCOUNT = ec.skp_bank_account
   and trunc(bt.DATE_BALANCE_VALID_FROM) = trunc(rt.DTIME_TRANSACTION_ACCOUNTED)
 where rt.FLAG_DELETED = 'N' 
   and rt.FLAG_TRANSACTION_DEBIT = 'Y' 
   and rt.FLAG_CLIENT_ACCOUNT_TRN = 'N'
   and rt.SKP_FEE not in (1173,1174,1175,1176/*UpomÌnka*/)
   and rt.DTIME_TRANSACTION_ACCOUNTED >= trunc(ec.date_first_day_dpd,'month')
   AND rt.DTIME_TRANSACTION_ACCOUNTED <= trunc(ec.den_pred_upadkem, 'month')
;



INSERT INTO APP_COLL_JEZEVCIK_TEST.epr_debety_tarify (SKP_BANK_ACCOUNT,
                                                      ID_PERSON,
                                                      POPLATEK,
                                                      DTIME_TRANSACTION_ACCOUNTED,
                                                      MONTH_TRANSACTION_ACCOUNTED,
                                                      AMT_TRANSACTION,
                                                      AMT_TRANSACTION_CZK,
                                                      SKF_REALIZED_TRANSACTION)

WITH debety_odp_trans AS (

select t.id_person,
       t.skp_bank_account,
       t.skf_realized_transaction,
       t.skf_required_transaction,
       t.skp_transaction_type,
       t.skp_transaction_status,
       t.skp_transaction_class,
       t.skp_fee,
       t.skp_currency,
       t.skf_realized_trans_cancelled,
       t.skf_realized_trans_cancelling,
       t.text_note_part1,
       t.amt_transaction,
       t.date_effective,
       t.date_transaction_due,
       t.dtime_transaction_generation,
       t.dtime_req_transact_generation,
       t.dtime_transaction_accounted,
       t.dtime_trans_status_valid_from,
       t.date_interest_calculation,
       t.flag_transaction_debit,
       t.flag_client_account_trn,
       t.amt_balance
     , tt.NAME_TRANSACTION_TYPE
     , tt.NAME_TRANSACTION_GROUP 
     , tt.NAME_TRANSACTION_CLASS
     , fee.NAME_FEE 
     , ts.NAME_TRANSACTION_STATUS
     , ts.NAME_TRANSACTION_WFSTATUS 

  from APP_COLL_JEZEVCIK_TEST.epr_debety_tarif_help t
  join owner_dwh.cl_transaction_type tt 
    on tt.SKP_TRANSACTION_TYPE = t.SKP_TRANSACTION_TYPE
  join owner_dwh.cl_transaction_status ts 
    on t.skp_transaction_status = ts.SKP_TRANSACTION_STATUS
  join owner_dwh.dc_fee fee 
    on t.skp_fee = fee.SKP_FEE
)


, base as (

-- Tady vyberu vsechny ucty, ktere v danych mesicich mely strzeny poplatek (Velk˝ tarif)
select tt.skp_bank_account
     , tt.id_person
     , case when tt.name_fee = ('MÏsÌËnÌ vedenÌ tarifu') then 'Tarif'
       end as poplatek
     , tt.dtime_transaction_accounted
     , trunc(tt.dtime_transaction_accounted,'MONTH') as month_transaction_accounted
     , tt.amt_transaction
     , tt.amt_transaction * nvl(cr.rate_exchange_middle, 1/*CZK*/) as amt_transaction_CZK
     , tt.skf_realized_transaction
     
  from debety_odp_trans tt
  left join owner_dwh.f_currency_rate_pd cr 
    on tt.SKP_CURRENCY = cr.skp_currency_source 
   and cr.date_rate_effective = trunc(tt.dtime_transaction_accounted)
 where tt.name_fee in ('MÏsÌËnÌ vedenÌ tarifu')
)

, odpustky1 as (

select tt.skp_bank_account
     , tt.skf_realized_transaction
     , tt.text_note_part1
     , tt.dtime_transaction_accounted
     , case when tt.text_note_part1 like '% - %' then instr(tt.text_note_part1,'-') end as pozice_pomlcky

  from debety_odp_trans tt 
 where tt.text_note_part1 like '%vracÌme%' 
   and tt.text_note_part1 like '%poplatek%'

)

, odpustky2 as (

select a.skp_bank_account
     , a.skf_realized_transaction 
     , a.text_note_part1
     , a.dtime_transaction_accounted

,case when pozice_pomlcky is not null and substr(a.text_note_part1,1,pozice_pomlcky-2) like '%leden%' then 1 
      when pozice_pomlcky is not null and substr(a.text_note_part1,1,pozice_pomlcky-2) like '%˙nor%' then 2 
      when pozice_pomlcky is not null and substr(a.text_note_part1,1,pozice_pomlcky-2) like '%b¯ezen%' then 3 
      when pozice_pomlcky is not null and substr(a.text_note_part1,1,pozice_pomlcky-2) like '%duben%' then 4 
      when pozice_pomlcky is not null and substr(a.text_note_part1,1,pozice_pomlcky-2) like '%kvÏten%' then 5 
      when pozice_pomlcky is not null and substr(a.text_note_part1,1,pozice_pomlcky-2) like '%Ëerven%' then 6 
      when pozice_pomlcky is not null and substr(a.text_note_part1,1,pozice_pomlcky-2) like '%Ëervenec%' then 7
      when pozice_pomlcky is not null and substr(a.text_note_part1,1,pozice_pomlcky-2) like '%srpen%' then 8
      when pozice_pomlcky is not null and substr(a.text_note_part1,1,pozice_pomlcky-2) like '%z·¯Ì%' then 9 
      when pozice_pomlcky is not null and substr(a.text_note_part1,1,pozice_pomlcky-2) like '%¯Ìjen%' then 10 
      when pozice_pomlcky is not null and substr(a.text_note_part1,1,pozice_pomlcky-2) like '%listopad%' then 11 
      when pozice_pomlcky is not null and substr(a.text_note_part1,1,pozice_pomlcky-2) like '%prosinec%' then 12 
      
      when pozice_pomlcky is null and a.text_note_part1 like '%leden%' then 1 
      when pozice_pomlcky is null and a.text_note_part1 like '%˙nor%' then 2 
      when pozice_pomlcky is null and a.text_note_part1 like '%b¯ezen%' then 3 
      when pozice_pomlcky is null and a.text_note_part1 like '%duben%' then 4 
      when pozice_pomlcky is null and a.text_note_part1 like '%kvÏten%' then 5 
      when pozice_pomlcky is null and a.text_note_part1 like '%Ëerven%' then 6 
      when pozice_pomlcky is null and a.text_note_part1 like '%Ëervenec%' then 7
      when pozice_pomlcky is null and a.text_note_part1 like '%srpen%' then 8
      when pozice_pomlcky is null and a.text_note_part1 like '%z·¯Ì%' then 9 
      when pozice_pomlcky is null and a.text_note_part1 like '%¯Ìjen%' then 10 
      when pozice_pomlcky is null and a.text_note_part1 like '%listopad%' then 11 
      when pozice_pomlcky is null and a.text_note_part1 like '%prosinec%' then 12     
        
      else EXTRACT(MONTH FROM a.dtime_transaction_accounted)   

 end as month_beginning

,case when pozice_pomlcky is not null and substr(a.text_note_part1,1,pozice_pomlcky-2) like '%2012%' then 2012 
      when pozice_pomlcky is not null and substr(a.text_note_part1,1,pozice_pomlcky-2) like '%2013%' then 2013 
      when pozice_pomlcky is not null and substr(a.text_note_part1,1,pozice_pomlcky-2) like '%2014%' then 2014 
      when pozice_pomlcky is not null and substr(a.text_note_part1,1,pozice_pomlcky-2) like '%2015%' then 2015 
      when pozice_pomlcky is not null and substr(a.text_note_part1,1,pozice_pomlcky-2) like '%2016%' then 2016 
      when pozice_pomlcky is not null and substr(a.text_note_part1,1,pozice_pomlcky-2) like '%2017%' then 2017 
        
      when pozice_pomlcky is null and a.text_note_part1 like '%2012%' then 2012 
      when pozice_pomlcky is null and a.text_note_part1 like '%2013%' then 2013 
      when pozice_pomlcky is null and a.text_note_part1 like '%2014%' then 2014 
      when pozice_pomlcky is null and a.text_note_part1 like '%2015%' then 2015 
      when pozice_pomlcky is null and a.text_note_part1 like '%2016%' then 2016 
      when pozice_pomlcky is null and a.text_note_part1 like '%2017%' then 2017 
        
      else EXTRACT(YEAR FROM a.dtime_transaction_accounted)   
      
 end as year_beginning

,case when pozice_pomlcky is not null and substr(a.text_note_part1,pozice_pomlcky,20) like '%leden%' then 1 
      when pozice_pomlcky is not null and substr(a.text_note_part1,pozice_pomlcky,20) like '%˙nor%' then 2 
      when pozice_pomlcky is not null and substr(a.text_note_part1,pozice_pomlcky,20) like '%b¯ezen%' then 3 
      when pozice_pomlcky is not null and substr(a.text_note_part1,pozice_pomlcky,20) like '%duben%' then 4 
      when pozice_pomlcky is not null and substr(a.text_note_part1,pozice_pomlcky,20) like '%kvÏten%' then 5 
      when pozice_pomlcky is not null and substr(a.text_note_part1,pozice_pomlcky,20) like '%Ëerven%' then 6 
      when pozice_pomlcky is not null and substr(a.text_note_part1,pozice_pomlcky,20) like '%Ëervenec%' then 7
      when pozice_pomlcky is not null and substr(a.text_note_part1,pozice_pomlcky,20) like '%srpen%' then 8
      when pozice_pomlcky is not null and substr(a.text_note_part1,pozice_pomlcky,20) like '%z·¯Ì%' then 9 
      when pozice_pomlcky is not null and substr(a.text_note_part1,pozice_pomlcky,20) like '%¯Ìjen%' then 10 
      when pozice_pomlcky is not null and substr(a.text_note_part1,pozice_pomlcky,20) like '%listopad%' then 11 
      when pozice_pomlcky is not null and substr(a.text_note_part1,pozice_pomlcky,20) like '%prosinec%' then 12 
        
      when pozice_pomlcky is null and a.text_note_part1 like '%leden%' then 1 
      when pozice_pomlcky is null and a.text_note_part1 like '%˙nor%' then 2 
      when pozice_pomlcky is null and a.text_note_part1 like '%b¯ezen%' then 3 
      when pozice_pomlcky is null and a.text_note_part1 like '%duben%' then 4 
      when pozice_pomlcky is null and a.text_note_part1 like '%kvÏten%' then 5 
      when pozice_pomlcky is null and a.text_note_part1 like '%Ëerven%' then 6 
      when pozice_pomlcky is null and a.text_note_part1 like '%Ëervenec%' then 7
      when pozice_pomlcky is null and a.text_note_part1 like '%srpen%' then 8
      when pozice_pomlcky is null and a.text_note_part1 like '%z·¯Ì%' then 9 
      when pozice_pomlcky is null and a.text_note_part1 like '%¯Ìjen%' then 10 
      when pozice_pomlcky is null and a.text_note_part1 like '%listopad%' then 11 
      when pozice_pomlcky is null and a.text_note_part1 like '%prosinec%' then 12  
        
      else EXTRACT(MONTH FROM a.dtime_transaction_accounted)   
      
 end as month_ending

,case when pozice_pomlcky is not null and substr(a.text_note_part1,pozice_pomlcky,20) like '%2012%' then 2012 
      when pozice_pomlcky is not null and substr(a.text_note_part1,pozice_pomlcky,20) like '%2013%' then 2013 
      when pozice_pomlcky is not null and substr(a.text_note_part1,pozice_pomlcky,20) like '%2014%' then 2014 
      when pozice_pomlcky is not null and substr(a.text_note_part1,pozice_pomlcky,20) like '%2015%' then 2015 
      when pozice_pomlcky is not null and substr(a.text_note_part1,pozice_pomlcky,20) like '%2016%' then 2016 
      when pozice_pomlcky is not null and substr(a.text_note_part1,pozice_pomlcky,20) like '%2017%' then 2017 
        
      when pozice_pomlcky is null and a.text_note_part1 like '%2012%' then 2012 
      when pozice_pomlcky is null and a.text_note_part1 like '%2013%' then 2013 
      when pozice_pomlcky is null and a.text_note_part1 like '%2014%' then 2014 
      when pozice_pomlcky is null and a.text_note_part1 like '%2015%' then 2015 
      when pozice_pomlcky is null and a.text_note_part1 like '%2016%' then 2016 
      when pozice_pomlcky is null and a.text_note_part1 like '%2017%' then 2017 
      
      else EXTRACT(YEAR FROM a.dtime_transaction_accounted) 
      
 end as year_ending
 
from odpustky1 a
)

, odpustky3 as (

  -- Ted mam sestrojene mesice, kdy byl klientum poplatek odpusten
  select a.skp_bank_account
       , a.skf_realized_transaction
       , a.text_note_part1
       , to_date(month_beginning||'/'||year_beginning,'MM/YYYY') as date_odpustek_from
       , last_day(to_date(month_ending||'/'||year_ending,'MM/YYYY')) as date_odpustek_to
  from odpustky2 a
  )



-- FINALNI SELECT 
--zostanu mi iba transakcie ktore to nenaslo v odpustkoch, cize tie NEodpustene.
select a.*
  from base a
  /*JOIN APP_COLL_JEZEVCIK_TEST.epr_debety_tarif_help tt
    on a.skf_realized_transaction = tt.skf_realized_transaction*/
  left join odpustky3 o 
    on a.skp_bank_account = o.skp_bank_account 
   and a.month_transaction_accounted between o.date_odpustek_from and o.date_odpustek_to
 where o.skp_bank_account is null -- nechci mesice, kdy byl poplatek jiz odpusten
;




     /*  --------------------------------------------  */
     /*    PAROVANIE STORNA TRANSAKCII (MIMO TARIF)    */
     /*  --------------------------------------------  */





/*
logika parovania:
 > cielom je u transakcii ktore sa navzajom neguju (poplatok a jeho vratenie) nastavit hodnotu flag_vlozit_do_epr = 'N' (v tabulke epr_debit_transactions)
 > v tabulke f_realized_transaction na dwh je nepochopitelna logika atributov skf_..._cancelled a skf_..._cancelling
 > ALE, niektore storna transakcii maju rovnake skf_..._cancelled ako povodna transakcia (ta ktoru rusili)
    - tie som naparoval ako prve a nastavil flag
 > U ostatnych som rozdelil skp_bank_account podla poctu storno transakcii
    - tam kde bola iba 1 na klienta som pohladal TRN s MENSIM dtime_transaction_accounted a rovnakym amt_balance. V pripade, ze taka sa nasla jedina, naparoval som ich (flag).
    - detto pre storna s roznym amt_balance u 1 klienta (pokial sa kazde vyskytlo prave raz)
    - tam kde sa u klienta opakovalo nejake storno viackrat, sme to bohuzial M nedokazali vymysliet.. Nizsie je dlhy skript k tomu (velmi podobny tomu prvemu, ale nefunguje uplne..
    - tyka sa cca 15 klientov, takze tych chcem hodit do nejakej black_sheep tabulky a v jezevcikovi to kontrolovat -> chybova hlaska

 > tarify tento skript neriesi, na to je zvlast tabulka epr_debety_tarify
*/


---- roztriedenie skp_bank_account podla toho, kolko storno transakci maju.
/*
SELECT a.skp_bank_account
     , a.skp_person
     , COUNT (*) pocet_storno_trn
FROM APP_COLL_JEZEVCIK_TEST.epr_debety_trans a
WHERE a.name_transaction_type NOT LIKE '%tarif%'
      AND lower(a.name_transaction_type) NOT LIKE '%daÚ z ˙roku%'
      AND lower(a.text_note_part1) NOT LIKE '%vracÌme v·m poplatek za jeho vedenÌ.%'
      AND a.flag_vlozit_do_epr = 'Y'                    -- robime iba s tymi TRN co este nie su naparovane
      AND a.skp_fee = -1
      AND a.flag_transaction_debit = 'Y'
      ----AND a.skf_realized_trans_cancelled > 0        -- toto je asi tiez zmysluplna podmienka, ale nemam to overene (nerozumiem atributu ani po znacnom usili)
      AND a.amt_transaction < 0
GROUP BY a.skp_bank_account, a.skp_person
--HAVING COUNT (*) != 1
;
*/

--- u storna transakci, ktore maju vo dvojici rovnake skf_realized_trans_cancelled, nastavime flag_vlozit_do_epr = 'N'
UPDATE APP_COLL_JEZEVCIK_TEST.epr_debety_trans edt
SET edt.flag_vlozit_do_epr = 'N'
WHERE EXISTS (
  WITH storno AS (

    SELECT MIN(skf_realized_transaction) skf_realized_transaction_min
         , MAX(skf_realized_transaction) skf_realized_transaction_max
    FROM APP_COLL_JEZEVCIK_TEST.epr_debety_trans a
    WHERE a.name_transaction_type NOT LIKE '%tarif%'
      AND lower(a.text_note_part1) NOT LIKE '%vracÌme v·m poplatek za jeho vedenÌ.%'
      AND a.skf_realized_trans_cancelled > 0
    GROUP BY a.skf_realized_trans_cancelled
    HAVING COUNT (DISTINCT a.skf_realized_transaction ) = 2
  )
  
  SELECT * 
  FROM storno 
  WHERE edt.skf_realized_transaction = storno.skf_realized_transaction_min
     OR edt.skf_realized_transaction = storno.skf_realized_transaction_max
)
;



------ parovanie storna u TRN, ktore sa nedaju podla skf_..._cancelled
--- naparuje sa iba vtedy, ak u daneho klienta a daneho storna najde jedinu TRN s opacnou vyskou a skorsim datumom

/*
do klienti_base ziskam klientov, ktori maju prave 1 storno (extra kvoli group by) pre nejaku vysku TRN
do stornovane_trn k tymto klientom najdem storna transakcii
do povodne_aj_storno_trn ku kazdej stornovanej najdem PRAVE JEDNU povodnu
*/
UPDATE APP_COLL_JEZEVCIK_TEST.epr_debety_trans edt
SET edt.flag_vlozit_do_epr = 'N'
WHERE EXISTS (
  WITH klienti_base AS (
  
    SELECT a.id_person
       , a.skp_bank_account
       , a.amt_transaction
       , COUNT (*) pocet_storno_trn
    FROM APP_COLL_JEZEVCIK_TEST.epr_debety_trans a
    WHERE a.name_transaction_type NOT LIKE '%tarif%'
      AND lower(a.name_transaction_type) NOT LIKE '%daÚ z ˙roku%'
      AND lower(a.text_note_part1) NOT LIKE '%vracÌme v·m poplatek za jeho vedenÌ.%'
      AND a.flag_vlozit_do_epr = 'Y'                    -- robime iba s tymi TRN co este nie su naparovane
      AND a.flag_transaction_debit = 'Y'
      AND a.skp_fee = -1
      AND a.amt_transaction < 0
    GROUP BY a.id_person, a.skp_bank_account, a.amt_transaction
    HAVING COUNT (*) = 1                               ---- vyberam tych ktori maju iba 1 storno (pri viacerych je potrebne parovat rucne(jezevcik zasteka))
  )
  
  , stornovane_trn AS (
  
    SELECT edt.*
    FROM APP_COLL_JEZEVCIK_TEST.epr_debety_trans edt
    JOIN klienti_base kb
      ON kb.skp_bank_account = edt.skp_bank_account
      AND kb.id_person = edt.id_person
      AND kb.amt_transaction = edt.amt_transaction
    WHERE edt.name_transaction_type NOT LIKE '%tarif%'
      AND lower(edt.name_transaction_type) NOT LIKE '%daÚ z ˙roku%'
      AND lower(edt.text_note_part1) NOT LIKE '%vracÌme v·m poplatek za jeho vedenÌ.%'
      AND edt.flag_vlozit_do_epr = 'Y'                    -- robime iba s tymi TRN co este nie su naparovane
      AND edt.flag_transaction_debit = 'Y'
      AND edt.skp_fee = -1
      AND edt.amt_transaction < 0
  )
  
  , povodne_aj_storno_trn AS (

    SELECT max(a.skf_realized_transaction) skf_realized_transaction_A
         , s.skf_realized_transaction skf_realized_transaction_B
    FROM APP_COLL_JEZEVCIK_TEST.epr_debety_trans a
    JOIN stornovane_trn s
      ON s.id_person = a.id_person
      AND s.skp_bank_account = a.skp_bank_account
      AND s.amt_transaction = - a.amt_transaction
      AND s.dtime_transaction_accounted >= a.dtime_transaction_accounted -- niektore storna maju rovnaky dtime ako original trn (zopar SMS)
    WHERE a.flag_vlozit_do_epr = 'Y'
      AND a.flag_transaction_debit = 'Y'
      AND a.name_transaction_type NOT LIKE '%tarif%'
      AND lower(a.name_transaction_type) NOT LIKE '%daÚ z ˙roku%'
      AND lower(a.text_note_part1) NOT LIKE '%vracÌme v·m poplatek za jeho vedenÌ.%'
    GROUP BY s.skf_realized_transaction
    HAVING COUNT(a.skf_realized_transaction) = 1                 --- zostanu iba storna, ktore maju prave 1 povodnu transakciu
  )
  
  SELECT * 
  FROM povodne_aj_storno_trn s
  WHERE edt.skf_realized_transaction = s.skf_realized_transaction_A
     OR edt.skf_realized_transaction = s.skf_realized_transaction_B
)
;

-- validity check
/*
SELECT COUNT(*) parne_cislo
     , SUM(a.amt_transaction) nula
FROM APP_COLL_JEZEVCIK_TEST.epr_debety_trans a
WHERE a.flag_vlozit_do_epr = 'N'
;
*/











     /*  --------------------------------------------  */
     /*                epr_debety_all                  */
     /*  --------------------------------------------  */







--     = tabulka z ktorej bude citat excel

-- do tejto tabulky ziskame vsetky NENAPAROVANE transakcie, aj posledni kladny zustatek, aj pripadne prichodzie transakcie - pomocou union all
-- potrebujeme k nim vysku, datum a kategoriu.
-- posledne 2 stlpce hovoria ci dany klient ma nejaku nezrovnalost. Idealne budu rovne 0, ak nie tak jezevcik zasteka ze sa nedotiahla nejaka transakcia alebo ostalo storno.
--  > obidva sa nacitavaju zvlast az nizsie.

INSERT INTO APP_COLL_JEZEVCIK_TEST.epr_debety_all (id_person, SKP_BANK_ACCOUNT, DTIME_TRANSACTION_ACCOUNTED, DATE_TRANSACTION_DUE, AMT_TRANSACTION, 
KATEGORIE, FLAG_TRANSACTION_DEBIT, NEVYSVETLENA_CASTKA, NENAPAROVANE_STORNA, SUMA_PRICHOZICH_TRN)
SELECT edt.id_person
     , edt.skp_bank_account
     , edt.dtime_transaction_accounted
     , trunc(edt.dtime_transaction_accounted) AS DATE_TRANSACTION_DUE
     , edt.amt_transaction_czk AS amt_transaction
     , 'tarif' kategorie
     , NULL flag_transaction_debit
     , NULL
     , NULL
     , NULL
FROM APP_COLL_JEZEVCIK_TEST.epr_debety_tarify edt

UNION ALL

SELECT z.id_person
     , z.skp_bank_account
     , z.date_balance_valid_from
     , trunc(z.date_balance_valid_from) AS DATE_TRANSACTION_DUE
     , z.posledni_kladny_zustatek AS amt_transaction
     , 'posledni kladny zustatek' AS  kategorie
     , NULL AS flag_transaction_debit
     , NULL
     , NULL
     , NULL
FROM APP_COLL_JEZEVCIK_TEST.epr_debety_zustatky z

UNION ALL -- union ALL, pretoze nechceme vyhodit duplikaty

-- nasledujuci select hodi vsetky polozky do 1 tabulky tak aby sa v nej jednoduchsie hladalo. 

SELECT edt.id_person
     , edt.skp_bank_account
     , edt.dtime_transaction_accounted
     , edt.DATE_TRANSACTION_DUE                             -- RISK-734
     , edt.amt_transaction
     , CASE WHEN LOWER(edt.name_transaction_type) LIKE ('%n·klady na vym·h%')
                 AND LOWER(edt.text_note_part1) not LIKE ('%upomÌnk%')
                 AND LOWER(edt.text_note_part1) not LIKE ('%storno%')
                 AND edt.skp_fee > 0
              THEN 'n·klady na vym·h·nÌ'                                          --- naklady na vymahani bez upominek
            
            WHEN LOWER(edt.text_note_part1) LIKE ('%upomÌnk%')
                 AND edt.skp_fee > 0
              THEN 'upomÌnka'
                
            WHEN LOWER(f.NAME_FEE) LIKE ('%sms%')
                 AND f.skp_fee > 0
              THEN 'sms'
            
            WHEN LOWER(f.NAME_FEE) LIKE ('%v˝pis%')
                 AND f.skp_fee > 0
              THEN 'v˝pis'
            
            WHEN LOWER(f.NAME_FEE) LIKE ('%v˝bÏr hotovosti%')
                 AND edt.skp_fee > 0
              THEN 'poplatek v˝bÏr'                                   -- POPLATKY za vyber hotovosti  -- nutno odlisit od "vyber u sazky"
            
            WHEN LOWER(f.NAME_FEE) LIKE ('%sazk%')
                 AND edt.skp_fee > 0
              THEN 'poplatek sazka'
            
            WHEN LOWER(edt.name_transaction_type) LIKE ('%˙hrada poji%')
              THEN 'pojiötÏnÌ'
              
            ELSE edt.name_transaction_type                     -- okrem ineho tu budu vybery (nie poplatky za ne)
            
       END kategorie
     , edt.flag_transaction_debit
     , NULL
     , NULL
     , NULL
FROM APP_COLL_JEZEVCIK_TEST.epr_debety_trans edt
JOIN owner_dwh.dc_fee f
  ON f.skp_fee = edt.skp_fee
WHERE edt.flag_vlozit_do_epr = 'Y'
;






/*
naplnime stlpec nevysvetlena_castka pre kazde id_person
 > idealne nula, to znamena ze suma transakcii a zostatku odpoveda vyske debetu
 > ak ine nez nula, jezevcik zasteka, niekde sme stratili transakciu a treba dohladat.
*/



update APP_COLL_JEZEVCIK_TEST.epr_debety_all
SET nevysvetlena_castka = 
(
  SELECT nevysvetlena_castka
  FROM (SELECT id_person
             , skp_bank_account
             , suma_transakci + debetni_zustatek AS nevysvetlena_castka -- tento select iba kvoli suctu stlpcov
        FROM (SELECT a.id_person
                   , a.skp_bank_account
                   , SUM(CASE WHEN a.flag_transaction_debit = 'Y' THEN a.amt_transaction
                              WHEN a.flag_transaction_debit = 'N' THEN - a.amt_transaction
                              WHEN a.flag_transaction_debit IS NULL AND a.kategorie = 'posledni kladny zustatek' THEN - a.amt_transaction
                              WHEN a.flag_transaction_debit IS NULL AND a.kategorie = 'tarif' THEN a.amt_transaction
                              END
                          ) AS suma_transakci
                   , NVL(MIN(ins.amt_balance), MIN(u.amt_balance)) AS debetni_zustatek  -- k insolvencnim klientum zustatek z dne pred upadkem, ostatni klienti aktualni
                     
              FROM APP_COLL_JEZEVCIK_TEST.epr_debety_all a
              JOIN APP_COLL_JEZEVCIK_TEST.ucet u
                ON u.id_person = a.id_person
                AND u.skp_bank_account = a.skp_bank_account
              LEFT JOIN ( SELECT prso.id_person, a.skp_bank_account, i.den_pred_upadkem, a.AMT_BALANCE  -- zustatky k insolvencnim klientum
                          FROM APP_COLL_JEZEVCIK_TEST.ip_vhodni_klienti_help i
                          JOIN APP_COLL_JEZEVCIK_TEST.person prso 
                            ON prso.id_person = i.id_person
                          JOIN  owner_dwh.f_account_balance_pd a 
                            ON a.skp_person = prso.skp_person 
                            AND a.SKP_BANK_ACCOUNT = i.skp_bank_account
                            AND a.DATE_EFFECTIVE = i.den_pred_upadkem
                        ) ins
                     ON ins.id_person = a.id_person
                     AND ins.skp_bank_account = a.skp_bank_account
                     
              WHERE u.amt_balance < 0
                AND u.typ_uctu = 'B⁄'
                AND u.code_currency = 'CZK'
              GROUP BY a.id_person, a.skp_bank_account
             )
        ) base 
  WHERE APP_COLL_JEZEVCIK_TEST.epr_debety_all.id_person = base.id_person
    AND APP_COLL_JEZEVCIK_TEST.epr_debety_all.skp_bank_account = base.skp_bank_account
)
;

-- validity check
/*
SELECT a.nevysvetlena_castka, COUNT(DISTINCT a.id_person)
FROM APP_COLL_JEZEVCIK_TEST.epr_debety_all a
GROUP BY a.nevysvetlena_castka
ORDER BY COUNT(DISTINCT a.id_person) DESC
;
-- pocet klientov u ktorych vyhodi chybovu hlasku
SELECT COUNT (DISTINCT id_person) FROM APP_COLL_JEZEVCIK_TEST.epr_debety_all a WHERE a.nevysvetlena_castka != 0
;
*/







/*
nenaparovane_storna
 > naplnime stlpec pre kazde id_person
 > idealne nula, pri inej hodnote bude jezevcik stekat
 > Hladame stornovane transakcie, ktore sa nepodarilo naparovat

*/

UPDATE APP_COLL_JEZEVCIK_TEST.epr_debety_all a
SET a.nenaparovane_storna = 
(
  SELECT base.nenaparovane_storna
  FROM (SELECT edt.id_person
             , edt.skp_bank_account
             , sum(edt.amt_transaction) nenaparovane_storna
        FROM APP_COLL_JEZEVCIK_TEST.epr_debety_trans edt
        WHERE edt.flag_transaction_debit = 'Y'
          AND edt.amt_transaction < 0                        -- neni nutna podminka
          AND edt.skp_fee < 0                                -- neni nutna podminka
          AND edt.flag_vlozit_do_epr = 'Y'                   -- chceme pouze nenaparovane
          AND edt.skf_realized_trans_cancelled > 0
        GROUP BY edt.id_person, edt.skp_bank_account
       ) base
  WHERE base.id_person = a.id_person
    AND base.skp_bank_account = a.skp_bank_account
)
;
-- este nastavime 0 miesto NULLs
UPDATE APP_COLL_JEZEVCIK_TEST.epr_debety_all a
SET a.nenaparovane_storna = 0
WHERE a.nenaparovane_storna IS NULL;






------ naplnime stlpec suma_prichozich_trn
---- pozor musi sa plnit az PO naplneni stlpcu nevysvetlena_castka (pretoze nizsie mazeme niektore trn, nesedeli by cisla vobec)

UPDATE APP_COLL_JEZEVCIK_TEST.epr_debety_all
SET suma_prichozich_trn = 
(
    SELECT suma_prichozich_trn
    FROM (SELECT a.id_person
               , a.skp_bank_account
               , nvl(SUM(CASE WHEN a.flag_transaction_debit = 'Y' AND a.amt_transaction < 0 AND a.kategorie = 'DaÚ z ˙roku' -- tieto TRN nizsie odstranime z tabulky
                                THEN - a.amt_transaction
                              WHEN a.flag_transaction_debit = 'N' AND a.amt_transaction > 0                                 -- tieto TRN nizsie odstranime z tabulky
                                THEN a.amt_transaction
                              END
                         ), 0) AS suma_prichozich_trn
          FROM APP_COLL_JEZEVCIK_TEST.epr_debety_all a
          GROUP BY a.id_person, a.skp_bank_account
          ) base
    WHERE epr_debety_all.id_person = base.id_person
      AND epr_debety_all.skp_bank_account = base.skp_bank_account
)
;

DELETE FROM APP_COLL_JEZEVCIK_TEST.epr_debety_all a
WHERE (a.flag_transaction_debit = 'N' AND a.amt_transaction > 0)                                -- odstranime TRN, ktore sme vyssie pridali do suma_prichozich_trn
   OR (a.flag_transaction_debit = 'Y' AND a.amt_transaction < 0 AND a.kategorie = 'DaÚ z ˙roku')
;

-- DELETE FROM APP_COLL_JEZEVCIK_TEST.epr_vhodni_zaloba_help; 

COMMIT;

END;

-------------------------------------------------------------------------------------
-- VLASTIKU - DOPLN DOKUMENTACI PROSIM!!!!!
-------------------------------------------------------------------------------------


PROCEDURE HEIRS IS


  -- Local variables here -- used when generating notices
  int_skp_notice INTEGER;
  int_deadline INTEGER := 7; -- deadline na upozornÏnÌ doplacenÌ vöemi dÏdici
  int_skp_note INTEGER;



  int_timeshift_status_sync INTEGER := 3; -- pro jistotu vÌce neû jeden den, kdyby nebÏûel load -- pro inci·lnÌ load dej 10000
  skp_heir_status INTEGER;
  nflag_insert_heir_status INTEGER;
  -- Local variables here
    -- kurzor pro proch·zenÌ klient˘
  CURSOR STAVY IS
  
  WITH base AS ( -- na zaË·tku nahraju inici·lnÌm loadem, pak se bude sledovat jen ned·vnÈ zmÏny
               SELECT e.name_epr_status AS NAME_STATUS, e.dtime_valid_from ,e.dtime_valid_to, 1000 + ROWNUM + nvl(max_order_item-1000,0) AS order_item -- divn˝ konstrukt, hled·m maxim·lnÌ sesynchronizovan˝ z danÈho ¯ÌzenÌ a k nÏmu naËÌt·m d·le
               FROM APP_COLL_JEZEVCIK_TEST.epr_status e
               CROSS JOIN (SELECT MAX(h.order_item) AS max_order_item FROM APP_COLL_JEZEVCIK_TEST.Heir_Status h WHERE h.order_item BETWEEN 1000 AND 1999)
               WHERE e.name_epr_status NOT LIKE '%öance%'/*dÏdicovÈ majÌ svoje öance*/ AND dtime_modification >= SYSDATE - int_timeshift_status_sync -- chci sledovat jen ned·vnÈ - ne d·vno ukonËenÈ stavy
               
               UNION
               SELECT i.name_ins_status AS NAME_STATUS, i.dtime_valid_from, i.dtime_valid_to, 2000 + ROWNUM+ nvl(max_order_item-2000,0) AS order_item 
               FROM APP_COLL_JEZEVCIK_TEST.ins_status i 
               CROSS JOIN (SELECT MAX(h.order_item) AS max_order_item FROM APP_COLL_JEZEVCIK_TEST.Heir_Status h WHERE h.order_item BETWEEN 2000 AND 2999)
               WHERE dtime_modification >= SYSDATE - int_timeshift_status_sync -- chci sledovat jen ned·vnÈ - ne d·vno ukonËenÈ stavy
               
               UNION 
               SELECT ex.name_exe_status AS NAME_STATUS, ex.dtime_valid_from, ex.dtime_valid_to, 3000 + ROWNUM + nvl(max_order_item-3000,0) AS order_item 
               FROM APP_COLL_JEZEVCIK_TEST.exe_status ex 
               CROSS JOIN (SELECT MAX(h.order_item) AS max_order_item FROM APP_COLL_JEZEVCIK_TEST.Heir_Status h WHERE h.order_item BETWEEN 3000 AND 3999)
               WHERE dtime_modification >= SYSDATE - int_timeshift_status_sync -- chci sledovat jen ned·vnÈ - ne d·vno ukonËenÈ stavy
               
               UNION
               SELECT d.name_ded_status AS NAME_STATUS, d.dtime_valid_from, d.dtime_valid_to, 4000 + ROWNUM + nvl(max_order_item-4000,0) AS order_item 
               FROM APP_COLL_JEZEVCIK_TEST.ded_status d 
               CROSS JOIN (SELECT MAX(h.order_item) AS max_order_item FROM APP_COLL_JEZEVCIK_TEST.Heir_Status h WHERE h.order_item BETWEEN 4000 AND 4999)
               WHERE dtime_modification >= SYSDATE - int_timeshift_status_sync -- chci sledovat jen ned·vnÈ - ne d·vno ukonËenÈ stavy
               )
               , prepared AS (
               SELECT b.*, row_number() OVER (PARTITION BY name_status ORDER BY order_item) AS rn
               FROM base b
               ) 
               -- Pokud je tam nÏjak˝ stav se stejn˝m jmÈnem vÌcekr·t, tak vyberu podle prioritnÌho ¯ÌzenÌ
               SELECT NAME_STATUS, dtime_valid_from ,dtime_valid_to, order_item
               FROM prepared 
               WHERE rn = 1;



BEGIN

--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/


-- Update tabulky PERSON - t˝k· se pouze dÏdic˘

-- PouûÌv· se p¯i p¯id·nÌ dÏdice, kter˝ je klientem. V momentÏ p¯id·nÌ nem·m k dispozici ËÌslo smlouvy, dot·hnu jej aû druh˝ den v procedu¯e
UPDATE APP_COLL_JEZEVCIK_TEST.person p
SET p.TEXT_CONTRACT_NUMBER = (SELECT c.TEXT_CONTRACT_NUMBER
                              from owner_dwh.f_contract_at c
                              JOIN owner_dwh.cl_contract_status CS ON c.SKP_CONTRACT_STATUS = cs.SKP_CONTRACT_STATUS
                              WHERE c.skp_person = p.skp_person AND p.skp_person > 0 AND p.flag_valid = 'Y'
                                AND c.CODE_CONTRACT_TYPE = 'OSOBA' AND cs.NAME_CONTRACT_STATUS = '⁄Ëinn·')
WHERE EXISTS (SELECT c.TEXT_CONTRACT_NUMBER
                              from owner_dwh.f_contract_at c
                              JOIN owner_dwh.cl_contract_status CS ON c.SKP_CONTRACT_STATUS = cs.SKP_CONTRACT_STATUS
                              WHERE c.skp_person = p.skp_person AND p.skp_person > 0 AND p.flag_valid = 'Y'
                                AND c.CODE_CONTRACT_TYPE = 'OSOBA' AND cs.NAME_CONTRACT_STATUS = '⁄Ëinn·')
AND p.text_contract_number IS NULL
;

--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/

-- Spl·cenÌ dÏdic˘:

-- SledujÌ se platby dÏdic˘ na BÏûn˝ nebo vnit¯nÌ ˙Ëet zem¯elÈho podle variabilnÌho symbolu v tabulce HEIRS

-- KÛd do procedury
MERGE INTO APP_COLL_JEZEVCIK_TEST.heir_payment l
      USING (-- B·ze ˙Ët˘
              WITH Ded_Account_Base AS 
              ( -- potom m˘ûu vyhodit odkazy na tabulky PERSON a DEDICTVI, aû to bude otestovane a bude fungovat. Staci si nechat SKP_HEIR, VS a skp_bank_account
                  SELECT h.skp_heir, h.variable_symbol, p.id_person/*zemrely*/, p.date_death, u.SKP_BANK_ACCOUNT, u.CISLO_UCTU, u.TYP_UCTU
                  -- select count(*), count(distinct p.id_person)
                  from APP_COLL_JEZEVCIK_TEST.person p
                  JOIN APP_COLL_JEZEVCIK_TEST.heirs h ON p.id_person = h.id_person_dead
                  JOIN APP_COLL_JEZEVCIK_TEST.ucet u ON p.id_person = u.id_person
                  /*JOIN owner_dwh.f_bank_account_ad bad ON bad.SKP_BANK_ACCOUNT = u.skp_bank_account
                  JOIN owner_dwh.cl_bank_account_type ba ON ba.SKP_BANK_ACCOUNT_TYPE = bad.SKP_BANK_ACCOUNT_TYPE*/
                  WHERE p.flag_valid = 'Y' AND p.date_death IS NOT NULL AND u.typ_uctu IN ('B⁄','E⁄','JPK')
              )

              -- Sleduji p¯ÌchozÌ transakce na tyto ˙Ëty
              /*select * from dba_part_key_columns where owner = 'OWNER_DWH' and name = 'FT_REALIZED_TRANSACTION_TT';*/

              SELECT b.skp_heir, b.skp_bank_account ,rt.SKF_REALIZED_TRANSACTION, rt.AMT_TRANSACTION, rt.DTIME_TRANSACTION_ACCOUNTED, rt.TEXT_VARIABLE_SYMBOL
              --, rt.CODE_CURRENCY -- Asi musÌ b˝t vûdy v korun·ch ûe?
              from  Ded_Account_Base b
              JOIN owner_dwh.f_realized_transaction_tt rt ON b.skp_bank_account = rt.SKP_BANK_ACCOUNT
                                                          AND rt.FLAG_TRANSACTION_DEBIT = 'N' -- chci kreditnÌ transakci, tedy ûe p¯iöel na tento ˙Ëet
                                                          AND rt.TEXT_VARIABLE_SYMBOL = b.variable_symbol -- kaûd˝ dÏdic bude mÌt sv˘j, podle toho rozdÏlÌm p¯ÌchozÌ transakce
                                                          AND rt.flag_deleted = 'N'
                                                          AND rt.DTIME_TRANSACTION_ACCOUNTED >= b.date_death -- chceme platby pouze po ˙mrtÌ, ty p¯edtÌm nechceme zapoËÌt·vat
              WHERE rt.DTIME_TRANSACTION_GENERATION > SYSDATE - 30 -- kv˘li partitionu na tabulku. RadÏji napoËtu vÌce neû poslednÌ den a potom zmerguju p¯es SKF transaction
              ) e ON (e.SKF_REALIZED_TRANSACTION = l.SKF_REALIZED_TRANSACTION) -- podmÌnky napojenÌ

      WHEN NOT MATCHED THEN -- pokud nenÌ z·znam v tabulce table, doplnÌ se

INSERT (l.skp_heir,l.SKF_REALIZED_TRANSACTION, l.skp_bank_account, l.AMT_TRANSACTION, l.DTIME_TRANSACTION_ACCOUNTED, l.TEXT_VARIABLE_SYMBOL)
VALUES (e.skp_heir,e.SKF_REALIZED_TRANSACTION, e.skp_bank_account, e.AMT_TRANSACTION, e.DTIME_TRANSACTION_ACCOUNTED, e.TEXT_VARIABLE_SYMBOL)
;


-- Potom n·sleduje kus kÛdu na tabulku HEIR, kde je ukazatel, kolik uû klient zaplatil, kter˝ by vûdy po insertu do tÈto tabulky updatoval onu hodnotu.

UPDATE APP_COLL_JEZEVCIK_TEST.heirs h
SET h.amt_receivable_paid = (SELECT SUM(hp.amt_transaction)
                     FROM APP_COLL_JEZEVCIK_TEST.heir_payment hp
                     WHERE h.skp_heir = hp.skp_heir
                     GROUP BY h.skp_heir)
WHERE EXISTS (SELECT SUM(hp.amt_transaction)
                     FROM APP_COLL_JEZEVCIK_TEST.heir_payment hp
                     WHERE h.skp_heir = hp.skp_heir
                     GROUP BY h.skp_heir);


-- Potom m· vyskoËit v Jezevcikovi upozorneni, pokud bude mÌt klient doplaceno... tedy APP_COLL_JEZEVCIK_TEST.heirs.amt_receivable_paid >= APP_COLL_JEZEVCIK_TEST.heirs.amt_receivable

-- Tady bych to videl na trigger se zmenou na vysi_pohledavky nebo zaplacenou pohledavku, kter˝ vygeneruje upozorneni na id_person_heir v NOTICE

--TRIGGER HEIR_AFTER_UPDATE ¯eöÌ individu·lnÌ doplacenÌ jednÌm dÏdicem
--  jakmile bude mÌt nÏkter˝ z dÏdic˘ doplaceno, zobrazÌ se upozornÏnÌ v jezevËÌkovi - na dÏdici


--  upozornÏnÌ takÈ vyskoËÌ v p¯ÌpadÏ, kdy majÌ vöichni dÏdicovÈ doplaceno -- to by mÏlo vyskoËit na zem¯elÈm
-- P¯edpoklad: celkov· suma se doplatÌ, pokud doplatÌ kaûd˝ dÏdic svoji Ë·st
-- Vöichni dÏdicovÈ doplatili, pot¯ebuji novÈ podmÌnky - nelze mÏnit triggerem, protoûe se pot¯ebuji odk·zat na mÏnÏnou tabulku a to pad·... Takûe musÌ jÌt do dennÌ procedury




-- Vöichni dÏdicovÈ doplatili, pot¯ebuji novÈ podmÌnky - nelze mÏnit triggerem, protoûe se pot¯ebuji odk·zat na mÏnÏnou tabulku a to pad·... Takûe musÌ jÌt do dennÌ procedury


SELECT MAX(skp_notice) INTO int_skp_notice FROM APP_COLL_JEZEVCIK_TEST.NOTICE;

MERGE INTO APP_COLL_JEZEVCIK_TEST.NOTICE n
      USING (
            SELECT h.*, int_skp_notice + ROWNUM AS skp_notice FROM (
              SELECT h.id_person_dead, SUM(nvl(h.amt_receivable,0)) AS total_receivable
              FROM APP_COLL_JEZEVCIK_TEST.HEIRS h
              WHERE h.amt_receivable > 0 -- P¯id·na podmÌnka, protoûe kdyû nevÌ kolik dluûÌ, tak dajÌ 0 a potom tam vyskakuje zbyteËnÏ upozornÏnÌ
              GROUP BY h.id_person_dead
              HAVING SUM(nvl(h.amt_receivable_paid,0)) >= SUM(nvl(h.amt_receivable,0))
              ) h
            ) h
      ON (h.id_person_dead = n.id_person AND n.text_notice LIKE 'Vöichni dÏdicovÈ doplatili svoje pohled·vky%')
      WHEN NOT MATCHED THEN
INSERT (n.skp_notice, n.id_person, n.text_notice, n.dtime_notice_created, n.dtime_notice_deadline, n.code_status, n.employee, n.skp_process)
VALUES (h.skp_notice, h.id_person_dead, 'Vöichni dÏdicovÈ doplatili svoje pohled·vky: '|| h.total_receivable || ' KË',SYSDATE, trunc(SYSDATE)+int_deadline,'a','sys',1)
   ;

-- JeötÏ pozn·mku o tom samÈm na zem¯elÈho

SELECT MAX(skp_ded_note) INTO int_skp_note FROM APP_COLL_JEZEVCIK_TEST.DED_NOTE;

MERGE INTO APP_COLL_JEZEVCIK_TEST.DED_NOTE n
      USING (
            SELECT h.*, int_skp_note + ROWNUM AS skp_note FROM (
              SELECT h.id_person_dead, SUM(nvl(h.amt_receivable,0)) AS total_receivable
              FROM APP_COLL_JEZEVCIK_TEST.HEIRS h
              WHERE h.amt_receivable > 0 -- P¯id·na podmÌnka, protoûe kdyû nevÌ kolik dluûÌ, tak dajÌ 0 a potom tam vyskakuje zbyteËnÏ upozornÏnÌ
              GROUP BY h.id_person_dead
              HAVING SUM(nvl(h.amt_receivable_paid,0)) >= SUM(nvl(h.amt_receivable,0))
              ) h
            ) h
      ON (h.id_person_dead = n.id_person AND n.text_note LIKE 'Vöichni dÏdicovÈ doplatili svoje pohled·vky%')
      WHEN NOT MATCHED THEN
INSERT (n.skp_ded_note, n.skp_ded_note_type ,n.id_person, n.text_note, n.dtime_note_created, n.employee)
VALUES (h.skp_note, 4 /*UpozornÏnÌ*/ ,h.id_person_dead, 'Vöichni dÏdicovÈ doplatili svoje pohled·vky: '|| h.total_receivable || ' KË',SYSDATE, 'sys')
   ;



-- Synchronizace stav˘ HEIR_STATUS

 FOR J IN STAVY LOOP
      
      --DBMS_OUTPUT.PUT_LINE('I got here:'||J.NAME_STATUS||' is the new value');  
  
      --Pokud tam stav jiû je, a novÏ jiû nenÌ platn˝ tak zneplatnÌm
      
      UPDATE APP_COLL_JEZEVCIK_TEST.Heir_Status h
      SET h.DTIME_VALID_TO = J.DTIME_VALID_TO, h.dtime_modification = SYSDATE
      WHERE h.NAME_HEIR_STATUS = J.NAME_STATUS -- stav tam jiû je
        AND j.DTIME_VALID_TO < SYSDATE -- byl zneplatnÏn
      ;
      
      
      
      --Pokud tam stav jeötÏ nenÌ tak p¯id·m
      
     select MIN(nvl2(h.name_heir_status,0,1)) INTO nflag_insert_heir_status -- pokud alespoÚ jedna nula tak nemÏnÌm - to znamen·, ûe tam stav uû byl
     from DUAL d -- d·v·m sem dual aby byl vûdy v˝sledkem nÏjak˝ ¯·dek
     LEFT JOIN APP_COLL_JEZEVCIK_TEST.Heir_Status h ON h.name_heir_status = J.NAME_STATUS
     ; 
     
     IF nflag_insert_heir_status = 1 THEN
     
             -- zÌsk·nÌ dalöÌho id pro novÈ z·znamy do tabulky upozornÏnÌ
      select max(skp_heir_status) into skp_heir_status from APP_COLL_JEZEVCIK_TEST.Heir_Status;
      if skp_heir_status is null then
        skp_heir_status := 1;
      else
        skp_heir_status := skp_heir_status + 1 ;
      END IF;

      -- vloûenÌ stavu
      insert into APP_COLL_JEZEVCIK_TEST.Heir_Status h (h.SKP_HEIR_STATUS, h.NAME_HEIR_STATUS, h.DTIME_VALID_FROM, h.DTIME_VALID_TO, h.DTIME_MODIFICATION, h.EMPLOYEE, h.SKP_PROCESS, h.ORDER_ITEM)
      values (skp_heir_status, J.NAME_STATUS, J.DTIME_VALID_FROM, J.DTIME_VALID_TO, SYSDATE, 'sys',6, J.ORDER_ITEM);
        
      --DBMS_OUTPUT.PUT_LINE('I got here:'||:J.NAME_STATUS || skp_heir_status  ||' is the new value');  
      
     END IF;

  END LOOP;





COMMIT;

EXCEPTION

--pokud procedura spadne, poöle se mail
  WHEN OTHERS THEN
  BEGIN
      SEND_MAIL('Martin.Macicek@airbank.cz', --email na kter˝ posÌl·m zpr·vu
                       'HEIRS failed', --p¯edmÏt zpr·vy
                       'Spadla procedura HEIRS, na n·sledujÌcÌ chybu:
                         '||translate(SQLERRM,CHR(13)||CHR(10),'//')) --text zpr·vy
                         ;

      SEND_MAIL('Filip.Antol@airbank.cz', --email na kter˝ posÌl·m zpr·vu
                       'HEIRS failed', --p¯edmÏt zpr·vy
                       'Spadla procedura HEIRS, na n·sledujÌcÌ chybu:
                         '||translate(SQLERRM,CHR(13)||CHR(10),'//')) --text zpr·vy
                         ;

  END;

end HEIRS;




-------------------------------------------------------------------------------------
-- Procedura generuje podkladovu tabulku pre cast V (tzv. petit) exekucniho navrhu
-- hlavni vystup je tabulka EXE_NAVRH_CAST_V s intervalmi.
-------------------------------------------------------------------------------------


PROCEDURE SABLONA_EXEKUCNI_NAVRH IS
  
  /*
  vytvoril Filip Antol, 7.5.2018
  RISK-977
  
  Do intervalov z reportu PM pohledavka vlozi nove podla toho, ci evidujeme nejake prichozi platby od klienta
  hlavni vystup je tabulka EXE_NAVRH_CAST_V s intervalmi.
  hlavni vystup je tabulka EXE_NAVRH_CAST_V s intervalmi.
  Tabule EXE_NAVRH_CAST_V kopiruje strukturu AP_COLL.PB_040_IMPORT_ALL. Predovsetkym z toho dovodu, ze casti I a V exekucniho navrhu jezevcik generuje rovnakym kodom.
  Nove intervaly su vkladane iterativne, v kurzore splatky_grouped je na to podmienka.
  Ak sa tabulka kompletne premaze, nicomu by to nemalo vadit.
  */


 
  -----  promenne k vytvareni intervalu pro HUcka

  strTypHU                           VARCHAR2(2) := 'HU';
  pocet_nalezenych_intervalu         NUMBER;
  pocitadlo_vlozenych_intervalu      NUMBER := 0;
  pocitadlo_failed_intervalu         NUMBER := 0;
  prichozi_platba                    V_EXE_NAVRH_UVERY_GROUPED%ROWTYPE;  -- promenna pro cely radek
  aktualni_interval                  EXE_NAVRH_CAST_V%ROWTYPE;           -- promenna pro cely radek

  CURSOR splatky_grouped IS
      SELECT a.ID_PERSON
           , a.DATUM_PLATBY_TRUNC
           , a.UVER_CISLO
           , a.amt_payment_jistina
           , a.amt_payment_urok
      FROM V_EXE_NAVRH_UVERY_GROUPED a
      WHERE a.DATUM_PLATBY_TRUNC > (SELECT MAX(b.OU_DTIME_VALID_FROM) FROM APP_COLL_JEZEVCIK_TEST.EXE_NAVRH_CAST_V B WHERE b.ID_PERSON = a.ID_PERSON)  --  Dulezite aby se neduplikovali platby, ktere byly vlozeny vcera
      ORDER BY a.DATUM_PLATBY_TRUNC ASC  --  !!! MUSI BYT SERAZENO !!!  -- jadro procedury je iterativni, bez serazeni se kompletne rozhazou intervaly kde byla vic nez jedina platba
      ;


  ---    promenne k novemu (new) zaznamu, ktery bude vlozen do tabulky intervalu
  new_jistina               exe_navrh_cast_V.JISTINA%TYPE;
  new_ku                    exe_navrh_cast_V.KU%TYPE;
  new_celkem                exe_navrh_cast_V.CELKEM%TYPE;
  -- ... su ...
  new_su_castka             exe_navrh_cast_V.SU_CASTKA%TYPE;
  new_su_dtime_valid_from   exe_navrh_cast_V.SU_DTIME_VALID_FROM%TYPE;
  new_su_dtime_valid_to     exe_navrh_cast_V.SU_DTIME_VALID_TO%TYPE;


BEGIN
    
    <<promazani_tabulky>>
    BEGIN
        
        DELETE FROM app_coll_jezevcik_test.exe_navrh_cast_V;
        INSERT INTO app_coll_jezevcik_test.exe_navrh_cast_V 
        SELECT p.ID_PERSON, /*a.CUID,*/  a.EPR, a.TYP, a.HU, a.CELKEM, a.JISTINA, a.KU, a.OU_SAZBA, a.OU_CASTKA,
                            a.OU_DTIME_VALID_FROM, a.OU_DTIME_VALID_TO, a.SU_SAZBA, a.SU_CASTKA, a.SU_DTIME_VALID_FROM,
                            a.SU_DTIME_VALID_TO, a.INTERVAL, a.BANK_ACCOUNT_NUMBER, a.DEBET, a.SOP,
                            a.TRAVEL_EXPENSES, a.OTHER_EXPENSES, a.NOTE, a.SPECIAL
        FROM ap_coll.PB_040_IMPORT_ALL a
        JOIN app_coll_jezevcik_test.person p ON a.cuid = p.cuid
        ;    
        
    END promazani_tabulky;

  
    --/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
    --/--/--/--/    update SOPu
    

    <<update_SOPu>>
    BEGIN
      
        MERGE INTO EXE_NAVRH_CAST_V a
        USING (
                  SELECT a.ID_PERSON
                       , SUM(a.amt_paid) AS amt_paid_sum
                  FROM V_EXE_NAVRH_SOPY_GROUPED a 
                  GROUP BY a.ID_PERSON 
                  HAVING SUM(a.amt_paid) > 0
              ) s
        ON (a.id_person = s.id_person)
        WHEN MATCHED THEN
        UPDATE
        SET a.SOP = a.sop - s.amt_paid_sum
        WHERE a.typ = 'SOP'
              AND a.sop IS NOT NULL
        ;
       
        
        -- kontrola:
          /* SELECT a.sop, a.* FROM EXE_NAVRH_CAST_V a WHERE a.sop IS NOT NULL AND a.ID_PERSON IN
          (
                SELECT a.ID_PERSON
                FROM V_EXE_NAVRH_SOPY_GROUPED a 
                GROUP BY a.ID_PERSON 
                HAVING SUM(a.amt_paid) > 0
          )
          ORDER BY a.ID_PERSON
          ;*/
 
    END update_SOPu;


    --/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/--/
    --/--/--/--/  update intervalov
    
    <<update_intervalov>>
    BEGIN
        open splatky_grouped;
          
          <<pro_kazdou_splatku>>
          LOOP

              fetch splatky_grouped INTO prichozi_platba;
              exit when splatky_grouped%notfound;

              ---      zjistit pocet intervalu v tabulce, ktore pokryvaju datum platby
              
              SELECT COUNT(id_person) INTO pocet_nalezenych_intervalu
              FROM app_coll_jezevcik_test.exe_navrh_cast_V a
              WHERE a.ID_PERSON = prichozi_platba.ID_PERSON                            -- stejna podminka jako nize v UPDATE aktualniho zaznamu
                AND a.HU = prichozi_platba.UVER_CISLO
                AND a.TYP = strTypHU
                AND prichozi_platba.DATUM_PLATBY_TRUNC BETWEEN a.OU_DTIME_VALID_FROM AND a.OU_DTIME_VALID_TO
              ;
              
              IF pocet_nalezenych_intervalu = 1 THEN

                  ---      vzdycky by mel byt prave 1
                  
                  SELECT */*a.EPR,
                         a.JISTINA,
                         a.KU,
                         a.INTERVAL,
                         a.OU_SAZBA, a.OU_CASTKA, a.OU_DTIME_VALID_FROM, a.OU_DTIME_VALID_TO,
                         a.SU_SAZBA, a.SU_CASTKA, a.SU_DTIME_VALID_FROM, a.SU_DTIME_VALID_TO*/
                  INTO  aktualni_interval
                  FROM app_coll_jezevcik_test.exe_navrh_cast_V a
                  WHERE a.ID_PERSON = prichozi_platba.ID_PERSON                            -- stejna podminka jako nize v UPDATE aktualniho zaznamu
                    AND a.HU = prichozi_platba.UVER_CISLO
                    AND a.TYP = strTypHU
                    AND prichozi_platba.DATUM_PLATBY_TRUNC BETWEEN a.OU_DTIME_VALID_FROM AND a.OU_DTIME_VALID_TO
                  ;

                  ---   overeni ze mame spravny zaznam
                  /*check/assert: interval = max(interval) u tohto klienta*/
                  /*check/assert: OU_DTIME_VALID_TO = 1.1.3000*/
                  /*check/assert: SU_DTIME_VALID_TO = 1.1.3000 or NULL*/



                  ------     update aktualniho zaznamu

                  IF prichozi_platba.DATUM_PLATBY_TRUNC > aktualni_interval.SU_DTIME_VALID_FROM THEN

                    ---- NEnachazime se v prvnich 10 dnech od data zesplatneni, kdy od klienta nechceme sankcni uroky
                    
                    UPDATE app_coll_jezevcik_test.exe_navrh_cast_V a
                    SET a.OU_DTIME_VALID_TO = prichozi_platba.DATUM_PLATBY_TRUNC - 1
                      , a.SU_DTIME_VALID_TO = prichozi_platba.DATUM_PLATBY_TRUNC - 1            -- muzeme, protoze plati ze prich_datum_platby > akt_su_dtime_valid_from
                      /*, a.CELKEM = ??*/
                      /*, a.ku = ??*/
                    WHERE a.ID_PERSON = prichozi_platba.ID_PERSON                           -- stejna podminka jako v select ... into akt_...
                      AND a.HU = prichozi_platba.UVER_CISLO
                      AND a.TYP = strTypHU
                      AND prichozi_platba.DATUM_PLATBY_TRUNC BETWEEN a.OU_DTIME_VALID_FROM AND a.OU_DTIME_VALID_TO
                    ;

                  ELSE

                    ---- nachazime se v prvnich 10 dnech od data zesplatneni, kdy od klienta NEchceme sankcni uroky
                    
                    UPDATE app_coll_jezevcik_test.exe_navrh_cast_V a
                    SET a.OU_DTIME_VALID_TO = prichozi_platba.DATUM_PLATBY_TRUNC - 1
                      , a.SU_DTIME_VALID_FROM = NULL
                      , a.SU_DTIME_VALID_TO = NULL
                      , a.SU_CASTKA = NULL
                      /*, a.CELKEM = ??*/
                      /*, a.ku = ??*/
                    WHERE a.ID_PERSON = prichozi_platba.ID_PERSON                           -- stejna podminka jako v select ... into akt_...
                      AND a.HU = prichozi_platba.UVER_CISLO
                      AND a.TYP = strTypHU
                      AND prichozi_platba.DATUM_PLATBY_TRUNC BETWEEN a.OU_DTIME_VALID_FROM AND a.OU_DTIME_VALID_TO
                    ;

                  END IF;


                  ------      vypocet novych hodnot


                  new_jistina := aktualni_interval.JISTINA - prichozi_platba.amt_payment_jistina;
                  new_ku :=      aktualni_interval.KU - prichozi_platba.amt_payment_urok;
                  new_celkem :=  new_jistina + new_ku;



                  -- ... hodnoty new_  resime pouze pro sankcni urok, pro pripad kdy by platba od klienta byla v prvnich 10 dnech od data zesplatneni ...

                  IF prichozi_platba.DATUM_PLATBY_TRUNC > aktualni_interval.SU_DTIME_VALID_FROM THEN

                    -- NEnachazime se v prvnich 10 dnech od data zesplatneni, kdy od klienta chceme sankcni uroky
                    new_su_castka := new_jistina;
                    new_su_dtime_valid_from := prichozi_platba.DATUM_PLATBY_TRUNC /*+ 1*/;
                    new_su_dtime_valid_to := DATE'3000-01-01';

                  ELSE

                    -- nachazime se v prvnich 10 dnech od data zesplatneni, kdy od klienta NEchceme sankcni uroky
                    new_su_castka := NULL;
                    new_su_dtime_valid_from := NULL;
                    new_su_dtime_valid_to := NULL;

                  END IF;




                  ------      vlozeni noveho intervalu


                  INSERT INTO app_coll_jezevcik_test.exe_navrh_cast_V (ID_PERSON, EPR, TYP, HU, CELKEM, JISTINA, KU, OU_SAZBA, OU_CASTKA, OU_DTIME_VALID_FROM, OU_DTIME_VALID_TO,
                                                                       SU_SAZBA, SU_CASTKA, SU_DTIME_VALID_FROM, SU_DTIME_VALID_TO, INTERVAL
                                                                       /*,BANK_ACCOUNT_NUMBER, DEBET, SOP, TRAVEL_EXPENSES, OTHER_EXPENSES, NOTE , SPECIAL*/)
                  VALUES (prichozi_platba.ID_PERSON, aktualni_interval.epr, strTypHU, prichozi_platba.UVER_CISLO
                          , new_celkem
                          , new_jistina
                          , new_ku
                          , aktualni_interval.OU_SAZBA
                          , new_jistina
                          , prichozi_platba.DATUM_PLATBY_TRUNC /*+ 1*/
                          , DATE'3000-01-01'
                          , aktualni_interval.SU_SAZBA
                          , new_su_castka
                          , new_su_dtime_valid_from
                          , new_su_dtime_valid_to
                          , aktualni_interval.INTERVAL + 1
                          )
                  ;


                  --dbms_output.put_line('done, id_person: ' || prich_id_pers );
                  pocitadlo_vlozenych_intervalu := pocitadlo_vlozenych_intervalu + 1;


              ELSIF pocet_nalezenych_intervalu > 1 THEN

                  dbms_output.put_line('vic nez 1 interval, id_person: ' || prichozi_platba.ID_PERSON );
                  pocitadlo_failed_intervalu := pocitadlo_failed_intervalu + 1;

              ELSE

                  dbms_output.put_line('nenalezen zadny interval, id_person: ' || prichozi_platba.ID_PERSON );
                  pocitadlo_failed_intervalu := pocitadlo_failed_intervalu + 1;

              END IF; -- pocet_nalezenych_intervalu = 1

          end LOOP pro_kazdou_splatku;
        close splatky_grouped;

        dbms_output.put_line('pocitadlo_vlozenych_intervalu: ' || pocitadlo_vlozenych_intervalu );
        dbms_output.put_line('pocitadlo_failed_intervalu: ' || pocitadlo_failed_intervalu );
        
        COMMIT;
        
    END update_intervalov;
    
END SABLONA_EXEKUCNI_NAVRH;






-------------------------------------------------------------------------------------
-- Kontroly
-- > jestli se v nektere tabulce nemnozi zaznamy prilis rychle (stalo se)
-- > jestli jsou zapnute vsechny triggery
-- > jestli jsou skompilovane vsechny objekty na hristi
-------------------------------------------------------------------------------------
procedure QUALITY_CHECK is

  -- Created on 22.5.2017 by FANTOL2174AB 
  
  pocet INTEGER;
  seznam VARCHAR2(10000);
  trigger1 NUMBER := 50;        -- maximalny pocet novych radku
  trigger2 NUMBER := 0.05;     -- maximalny relativny pocet novych radku
  min_pocet_radku NUMBER := 30; -- ignoruje male tabulky
  
  skp INTEGER;
  current_table VARCHAR2(100);
  CURSOR all_tables_cur IS
    SELECT table_name FROM all_tables WHERE owner IN (nazev_aktualniho_hriste) AND TABLE_NAME != 'NROW_COUNT';
    
  nezkompilovane_objekty_pocet INTEGER;
  nezkompilovane_objekty_seznam NVARCHAR2(5000);
  vypnute_tigry_pocet INTEGER;
  vypnute_tigry_seznam NVARCHAR2(5000);
    
BEGIN
  
  OPEN all_tables_cur;
  LOOP
    FETCH all_tables_cur INTO current_table;
    EXIT WHEN all_tables_cur%NOTFOUND;
        
    SELECT max(a.skp_nrow_count) INTO skp 
    FROM APP_COLL_JEZEVCIK_TEST.nrow_count a;
    IF skp IS NULL THEN
      skp := 1;
    ELSE 
      skp := skp+1;
    END IF;
    
    EXECUTE IMMEDIATE 'INSERT INTO APP_COLL_JEZEVCIK_TEST.nrow_count
                         SELECT ' || skp || ' AS SKP_nrow_count,
                         ''' || current_table || ''' as table_name,
                         TRUNC(SYSDATE) AS date_period,
                         COUNT(*) AS nrow
                       FROM APP_COLL_JEZEVCIK_TEST.'|| current_table ||'
                       HAVING COUNT(*) > '|| min_pocet_radku ||'
                       '; 
  END LOOP;
  COMMIT;
  
  /* puvodni verze, NUM_ROWS vsak obsahovalo neaktualni data
  INSERT INTO APP_COLL_JEZEVCIK_TEST.nrow_count
  SELECT APP_COLL_JEZEVCIK_TEST.skp_nrow_count_seq.nextval AS SKP_nrow_count,
         TABLE_NAME,
         TRUNC(SYSDATE) AS date_period,
         (NUM_ROWS) AS nrow                -- XX num_rows davalo neaktualne pocty riadkov
  FROM ALL_TABLES
  WHERE  OWNER in (nazev_aktualniho_hriste)
    AND num_rows > trigger1;
  COMMIT;*/  -- povodna verzia, num_rows ale zlyhalo kvoli neaktualnym statistikam
  
  
  SELECT COUNT(*), replace(to_char(wm_concat(table_name)), ',', CHR(13)||chr(10)) INTO pocet, seznam -- skaredy zapis aby bola kazda tabulka na novom riadku
  FROM (
    SELECT a.* ,
    LAG (a.nrow,1) OVER (PARTITION BY a.table_name ORDER BY a.date_period) AS nrow_previous
    FROM APP_COLL_JEZEVCIK_TEST.nrow_count a) b
  WHERE date_period = TRUNC(SYSDATE)
  AND (nrow - nrow_previous) > trigger1
  AND (nrow - nrow_previous) > nrow_previous * trigger2;

  --dbms_output.put_line(pocet);
  --dbms_output.put_line(seznam);

  IF (pocet >= 1) THEN
      BEGIN
        SEND_MAIL(recipient => 'Filip.Antol@airbank.cz',
                     subj => 'Upozorneni - mnoho novych radku v Jezevcikovi',
                     mail => 'Prilis mnoho novych radku v ' || pocet || ' tabulkach, je potreba zkontrolovat tabulky: ' || CHR(13)||CHR(10) || seznam);
      END;
  END IF;



  --    CAST II - kontrola ci su vsetky objekty na hristi skompilovane
  
  
  SELECT COUNT(*) INTO nezkompilovane_objekty_pocet
      FROM   user_objects a
      WHERE  STATUS <> 'VALID'
  ;
  
  IF nezkompilovane_objekty_pocet > 0 THEN
      
      SELECT replace(to_char(wm_concat(a.OBJECT_NAME)), ',', CHR(13)||chr(10)) INTO nezkompilovane_objekty_seznam
          FROM   user_objects a
          WHERE  STATUS <> 'VALID'
      ;
      
      BEGIN
        SEND_MAIL(recipient => 'Filip.Antol@airbank.cz',
                     subj => 'Upozorneni - nezkompilovane objekty v Jezevcikovi',
                     mail => 'Bylo nalezeno ' || nezkompilovane_objekty_pocet || ' nezkompilovanych objektu: ' || CHR(13)||CHR(10) || nezkompilovane_objekty_seznam);
      END;
      
      BEGIN
        SEND_MAIL(recipient => 'Martin.Macicek@airbank.cz',
                     subj => 'Upozorneni - nezkompilovane objekty v Jezevcikovi',
                     mail => 'Bylo nalezeno ' || nezkompilovane_objekty_pocet || ' nezkompilovanych objektu: ' || CHR(13)||CHR(10) || nezkompilovane_objekty_seznam);
      END;
      
      BEGIN
        SEND_MAIL(recipient => 'Vlastimil.Severa@airbank.cz',
                     subj => 'Upozorneni - nezkompilovane objekty v Jezevcikovi',
                     mail => 'Bylo nalezeno ' || nezkompilovane_objekty_pocet || ' nezkompilovanych objektu: ' || CHR(13)||CHR(10) || nezkompilovane_objekty_seznam);
      END;
      
  END IF;
  
  
  
  --    CAST III - kontrola ci su vsetky triggers ENABLED
  
  
  SELECT COUNT(*) INTO vypnute_tigry_pocet
      FROM user_triggers a 
      WHERE a.status <> 'ENABLED'
  ;
  
  IF vypnute_tigry_pocet > 0 THEN
    
      SELECT replace(to_char(wm_concat(a.trigger_name)), ',', CHR(13)||chr(10)) INTO vypnute_tigry_seznam
          FROM user_triggers a 
          WHERE a.status <> 'ENABLED'
      ;
      
      BEGIN
        SEND_MAIL(recipient => 'Filip.Antol@airbank.cz',
               subj => 'Upozorneni - vypnute tigry v Jezevcikovi',
               mail => 'Bylo nalezeno ' || vypnute_tigry_pocet || ' vypnutych tigru: ' || CHR(13)||CHR(10) || vypnute_tigry_seznam);
      END;
      
      BEGIN
        SEND_MAIL(recipient => 'Martin.Macicek@airbank.cz',
               subj => 'Upozorneni - vypnute tigry v Jezevcikovi',
               mail => 'Bylo nalezeno ' || vypnute_tigry_pocet || ' vypnutych tigru: ' || CHR(13)||CHR(10) || vypnute_tigry_seznam);
      END;
    
      BEGIN
        SEND_MAIL(recipient => 'Vlastimil.Severa@airbank.cz',
               subj => 'Upozorneni - vypnute tygri v Jezevcikovi',
               mail => 'Bilo nalezeno ' || vypnute_tigry_pocet || ' vipnutich tygru: ' || CHR(13)||CHR(10) || vypnute_tigry_seznam);
      END;
    
  END IF;


EXCEPTION

--pokud procedura spadne, poöle se mail
  WHEN OTHERS THEN
  BEGIN
      SEND_MAIL('Filip.Antol@airbank.cz', --email na kter˝ posÌl·m zpr·vu
                       'jezevcik_quality_check', --p¯edmÏt zpr·vy
                       'Spadla procedura jezevcik_quality_check, na n·sledujÌcÌ chybu:
                         '||translate(SQLERRM,CHR(13)||CHR(10),'//')) --text zpr·vy
                         ;
  END;
  
  BEGIN
      SEND_MAIL('Martin.Macicek@airbank.cz', --email na kter˝ posÌl·m zpr·vu
                       'jezevcik_quality_check', --p¯edmÏt zpr·vy
                       'Spadla procedura jezevcik_quality_check, na n·sledujÌcÌ chybu:
                         '||translate(SQLERRM,CHR(13)||CHR(10),'//')) --text zpr·vy
                         ;
  END;

END;



-------------------------------------------------------------------------------------
-- Tato procedura hlida, zda se podarilo nacist data do Jezevicka a posle mail.
-------------------------------------------------------------------------------------
PROCEDURE CHECK_JEZ_SUCCEEDED is

trvani_loadu number;

BEGIN
  begin
    
  select round(24*(sysdate - max(startime)),2) into trvani_loadu from chain_tracking where trunc(startime) = trunc(sysdate);
  
      SEND_MAIL(recipient => 'Vlastimil.Severa@airbank.cz',
                   subj => 'Data do Jezevcika se podarilo nacist',
                   mail => 'Vse je OK, load trval celkem '|| trvani_loadu ||' hodin');
      SEND_MAIL(recipient => 'Filip.Antol@airbank.cz',
                   subj => 'Data do Jezevcika se podarilo nacist',
                   mail => 'Vse je OK, load trval celkem '|| trvani_loadu ||' hodin');
      SEND_MAIL(recipient => 'Martin.Macicek@airbank.cz',
                   subj => 'Data do Jezevcika se podarilo nacist',
                   mail => 'Vse je OK, load trval celkem '|| trvani_loadu ||' hodin');        
END;


update chain_tracking
   set endtime = sysdate, duration = sysdate - startime, result = 'SUCCEDED'
 where trunc(startime) = trunc(sysdate)
   and endtime is NULL
   and chain_name = 'CHAIN_D1'
   AND startime = (SELECT max(startime) from chain_tracking where trunc(startime) = trunc(sysdate));


insert into CHAIN_tracking_step
SELECT job_name, job_subname, status,  
                  TO_DATE(TO_CHAR(actual_start_date, 'YYYY-MON-DD HH24:MI:SS'), 
                          'YYYY-MON-DD HH24:MI:SS') start_date,   extract(hour from run_duration) * 60 +
                  extract(minute from run_duration) +
                  extract(minute from run_duration) / 60 duration_in_min
FROM USER_SCHEDULER_JOB_RUN_DETAILS
where JOB_SUBNAME IS NOT NULL
and 
                  trunc(TO_DATE(TO_CHAR(actual_start_date, 'YYYY-MON-DD HH24:MI:SS'), 
                          'YYYY-MON-DD HH24:MI:SS')) = trunc(sysdate)
;

commit;

end;


-------------------------------------------------------------------------------------
-- Tato procedura hlida, zda se podarilo nacist data do Jezevicka a posle mail.
-------------------------------------------------------------------------------------
PROCEDURE CHECK_JEZ_FAILED is

BEGIN
  begin
      SEND_MAIL(recipient => 'Vlastimil.Severa@airbank.cz',
                   subj => 'Upozorneni - nepodarilo se nacist vcerejsi data do Jezevcika',
                   mail => 'Je potreba zkontrolovat a opravyt PCKG_JEZEVCIK'); -- ovÏ¯Ìme, ûe VlastÌk Ëte maily.
      SEND_MAIL(recipient => 'Filip.Antol@airbank.cz',
                   subj => 'Upozorneni - nepodarilo se nacist vcerejsi data do Jezevcika',
                   mail => 'Je potreba zkontrolovat a opravit PCKG_JEZEVCIK');
      SEND_MAIL(recipient => 'Martin.Macicek@airbank.cz',
                   subj => 'Upozorneni - nepodarilo se nacist vcerejsi data do Jezevcika',
                   mail => 'Je potreba zkontrolovat a opravit PCKG_JEZEVCIK');       


                   
update chain_tracking
   set endtime = sysdate, duration = sysdate - startime, RESULT = 'FAILED'
 where trunc(startime) = trunc(sysdate)
   and endtime is null
   and chain_name = 'CHAIN_D1'
   AND startime = (SELECT max(startime) from chain_tracking where trunc(startime) = trunc(sysdate));

insert into CHAIN_tracking_step
SELECT job_name, job_subname, status,  
                  TO_DATE(TO_CHAR(actual_start_date, 'YYYY-MON-DD HH24:MI:SS'), 
                          'YYYY-MON-DD HH24:MI:SS') start_date,   extract(hour from run_duration) * 60 +
                  extract(minute from run_duration) +
                  extract(minute from run_duration) / 60 duration_in_min
FROM USER_SCHEDULER_JOB_RUN_DETAILS
where JOB_SUBNAME IS NOT NULL and 
                  trunc(TO_DATE(TO_CHAR(actual_start_date, 'YYYY-MON-DD HH24:MI:SS'), 
                          'YYYY-MON-DD HH24:MI:SS')) = trunc(sysdate)
;

commit;


END;
end;





end PCKG_JEZEVCIK;
/
