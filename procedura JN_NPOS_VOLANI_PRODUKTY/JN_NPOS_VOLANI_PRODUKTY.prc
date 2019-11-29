create or replace procedure JN_NPOS_VOLANI_PRODUKTY

       (     -- p_schema varchar2 := 'ap_sales'
     -- , p_date date := '31.1.2019'
        p_date date := trunc(sysdate-1)
     -- , p_date date := trunc(add_months(last_day(sysdate-1),-1))    
                  )

authid current_user

is

begin

if user = 'AP_SALES'
  then
    
 begin
   /* -- email posilame z procedury na CRM
      SEND_MAIL(recipient => 'vendula.ondraskova@airbank.cz', 
                   subj => 'Poboèkový reporting - start', 
                   mail => 'Právì se zaèal poèítat POS reporting 2018/2019.'); 


     SEND_MAIL(recipient => 'petr.boril@airbank.cz', 
                   subj => 'Poboèkový reporting - start', 
                   mail => 'Právì se zaèal poèítat POS reporting 2018/2019.'); 
   */                
     SEND_MAIL(recipient => 'petra.hottowa@airbank.cz',
                   subj => 'Poboèkový reporting - start', 
                   mail => 'Právì se zaèal poèítat POS reporting 2018/2019.');                    
                   
     SEND_MAIL(recipient => 'tomas.kostrhoun@airbank.cz',
                   subj => 'Poboèkový reporting - start', 
                   mail => 'Právì se zaèal poèítat POS reporting 2018/2019.'); 

  END;

end if;

----------------------------------------
--VOLÁNÍ ÚVÌRY
----------------------------------------

--naèerpané úvìry daný mìsíc
insert into JN_NPOS_VOLANI_UVERY_1_GTT
select fl.SKP_LOAN
       , fl.SKP_APPLICATION
       , fa.SKP_APPLICATION_PARENT as parent_skp_application
       , case when fa1.SKP_APPLICATION_PARENT= -1 then fa.skp_application else fa1.SKP_APPLICATION_PARENT end as skp_application_subject
       , fl.skp_contract
       , fl.skp_person
       , fl.SKP_PRODUCT 
       , p.TEXT_PRODUCT_DESCRIPTION
       , fl.skp_completion_subject
       , fl.CODE_LOAN_NUMBER
       , fl.DTIME_INSERTED
       , fl.DATE_LOAN_OPEN
       , fl.DATE_LOAN_UTILIZATION
       , to_char(fl.DATE_LOAN_UTILIZATION, 'YYYY/MM') as mesic_nacerpani
       , fl.AMT_LOAN_LIMIT
       , s.NAME_LOAN_STATUS
       , fa1.TEXT_APPLICATION_NUMBER
from owner_dwh.f_loan_at fl
join owner_dwh.dc_product p on p.SKP_PRODUCT = fl.SKP_PRODUCT
join owner_dwh.cl_loan_status s on s.SKP_LOAN_STATUS=fl.SKP_LOAN_STATUS 
join owner_dwh.f_application_at fa on fa.SKP_APPLICATION = fl.SKP_APPLICATION
join owner_dwh.f_application_at fa1 on fa1.SKP_APPLICATION = fa.SKP_APPLICATION
where trunc(date_loan_utilization) >  TO_CHAR( ADD_MONTHS(LAST_DAY(to_char(p_date,'DD.MM.YYYY')), -1))and  date_loan_utilization < '1.1.3000'
;

--datum dokonèení žádosti - ze strany klienta, neuvažujeme hypotéky
insert into JN_NPOS_VOLANI_UVERY_2_GTT 
select s.skp_application 
       , trunc(min(att.DTIME_VALID_FROM)) as datum_dokonceni_klient  ---z max zmìnìno na min 28.3.2019
       , att.SKP_EMPLOYEE_LOGGEDIN as skp_employee --DP-546, pøidáno 31.10.2019
from JN_NPOS_VOLANI_UVERY_1_GTT s
join owner_dwh.f_application_tt  att on s.skp_application = att.SKP_APPLICATION
join owner_ext.ext_pos_employee e on att.SKP_EMPLOYEE_LOGGEDIN= e.skp_employee --DP-546, pøidáno 31.10.2019
where code_last_page in ('waitForPreScoringResultTask', 'waitForLoanScoringResultTask', 'waitForScoringResultTask')--in ('generateDocumentationTask', 'finalizationTask')--DP-546, pøidáno 31.10.2019
      and s.text_product_description not like '%Hypoték%'
      and att.dtime_valid_from  < = att.dtime_valid_to
group by s.skp_application, att.SKP_EMPLOYEE_LOGGEDIN
;

--doplním datum podpisu tam, kde nedohledám datum dokonèení žádosti ze strany klienta
insert into JN_NPOS_VOLANI_UVERY_2_GTT 
with pom as(
select p.skp_application, p.skp_loan 
from JN_NPOS_VOLANI_UVERY_1_GTT p
where not exists (select * from  JN_NPOS_VOLANI_UVERY_2_GTT z where z.skp_application = p.skp_application)
and p.text_product_description not like '%Hypo%')
, pom1 as (select o.*
       , CODE_COMPLETION_SUBJECT_TYPE
       , skp_signature
from JN_NPOS_VOLANI_UVERY_1_GTT o
join (--kompletace kvùli podpisu
    select pc.id_source
           , pc.skp_completion_subject 
           , pc.id_source_application as id_appl_loan
           , pc.skp_completion_subject as skp_signature
           ,pc.code_completion_subject_type
    from OWNER_DWH.f_completion_subject_at pc
    where pc.code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING', 'RHY_LOAN_COMMITMENT') and pc.skp_completion_parent ='-1'
      union
      --úvìr + BU
    select pc.id_source
           , pc.skp_completion_subject 
           , pc.id_source_application as id_appl_loan
           , pc1.skp_completion_subject as skp_signature
           ,pc.code_completion_subject_type
    from OWNER_DWH.f_completion_subject_at pc
    join OWNER_DWH.f_completion_subject_at  pc1 on (pc1.skp_completion_subject = pc.skp_completion_parent and pc1.skp_completion_parent ='-1')
    where pc.code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING','RHY_LOAN_COMMITMENT') and pc.skp_completion_parent <>'-1'
      union
      --úvìr + BU + RS
    select pc.id_source
           , pc.skp_completion_subject 
           , pc.id_source_application as id_appl_loan
           , pc1.skp_completion_parent as skp_signature
           , pc.code_completion_subject_type
    from OWNER_DWH.f_completion_subject_at pc
    join OWNER_DWH.f_completion_subject_at pc1 on (pc1.skp_completion_subject = pc.SKP_COMPLETION_PARENT and pc1.skp_completion_parent <>'-1')
    where pc.code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING','RHY_LOAN_COMMITMENT') and pc.SKP_COMPLETION_PARENT <>'-1'
 ) k on k.id_appl_loan=o.text_application_number
  )
select po.skp_application
       , trunc(ct.DTIME_SIGNATURE) as datum_dokonceni_klient 
       , ct.SKP_EMPLOYEE
from pom po
join pom1 p1 on po.skp_application = p1.skp_application
join owner_dwh.f_completion_signature_tt ct on ct.SKP_COMPLETION_SUBJECT = p1.SKP_SIGNATURE and trunc(ct.DTIME_SIGNATURE_VALID_TO) = '1.1.3000'
join owner_ext.ext_pos_employee e on e.skp_employee = ct.SKP_EMPLOYEE
;


--naèerpané pùjèky klienta, propojené na kampaò - hovor
insert into JN_NPOS_VOLANI_UVERY_3_GTT
select distinct t.skp_loan
       , t.skp_application
       , t.parent_skp_application
       , t.skp_application_subject
       , t.skp_contract
       , t.skp_person
       , t.skp_product
       , t.amt_loan_limit
       , t.text_product_description
       , t.text_application_number
       , t.skp_completion_subject
       , t.code_loan_number
       , t.dtime_inserted
       , t.date_loan_open
       , t.date_loan_utilization
     --  , csa.CODE_COMPLETION_SUBJECT_TYPE
       , e.skp_branch --zmena ee.skp_branch DP-546
    /* , case when tt.SKP_EMPLOYEE_FINISHED = 28074 then 30522 
              when tt.SKP_EMPLOYEE_FINISHED = 29260 then 35408 
              when tt.SKP_EMPLOYEE_FINISHED = 3347 then 36732  
                else tt.SKP_EMPLOYEE_FINISHED end as skp_employee --k.skp_finish_employee, DP-546*/ --nahrazeno dz.skp_employee DP - 599, 18.11.2019
       , dz.skp_employee
       , t.mesic_nacerpani
       , t.name_loan_status
       , k.campaign_name
       , trunc(a.DTIME_CONTACT_realized) as datum_kontaktu --k.datum_kontaktu DP-546 -- D-599, nahrazeno 18.11.2019 tt.DTIME_CONTACT_FINISHED
       , trunc(a.DTIME_CONTACT_realized) as dtime_planned_call_finished --k.dtime_planned_call_finished DP-546 -- D-599, nahrazeno 18.11.2019 tt.DTIME_CONTACT_FINISHED
       , t.date_loan_utilization as datum_nacerpani_pujcky--k.datum_nacerpani_pujcky DP-546
       , dz.datum_dokonceni_klient
from JN_NPOS_VOLANI_UVERY_1_GTT  t
left join JN_NPOS_VOLANI_UVERY_2_GTT dz on dz.skp_application = t.skp_application
join owner_dwh.f_contact_at a on a.SKP_PERSON = t.skp_person --DP-546, pøidáno 31.10.2019
left join owner_dwh.f_contact_finish_tt tt on a.SKP_CONTACT = tt.skp_contact --DP-546, pøidáno 31.10.2019
and dz.skp_employee = case when tt.SKP_EMPLOYEE_FINISHED = 28074 then 30522 
                           when tt.SKP_EMPLOYEE_FINISHED = 29260 then 35408 
                           when tt.SKP_EMPLOYEE_FINISHED = 3347 then 36732  
                                else tt.SKP_EMPLOYEE_FINISHED end --DP-546, pøidáno 31.10.2019
join owner_ext.ext_pos_employee e on e.skp_employee = dz.skp_employee --DP-546, pøidáno 31.10.2019
left join ap_crm.JN_NPOS_KAMP_VOLANI_2018 k on t.skp_person = k.skp_person  --DP-546, pøidáno 31.10.2019
--and k.skp_finished_employee = tt.SKP_EMPLOYEE_FINISHED 
and k.datum_kontaktu = trunc(a.DTIME_CONTACT_realized)
--join ap_crm.JN_NPOS_KAMP_VOLANI_2018 k on t.skp_loan = k.skp_loan
/*join ap_crm.JN_NPOS_KAMP_VOLANI_2018 k on t.skp_person = k.skp_person
join owner_ext.ext_pos_employee ee on ee.skp_employee = case when k.skp_finished_employee = 28074 then 30522 
                                                             when k.skp_finished_employee = 29260 then 35408 
                                                             when k.skp_finished_employee = 3347 then 36732  
                                                                  else k.skp_finished_employee end*/
--join owner_dwh.f_completion_subject_at csa on csa.SKP_COMPLETION_SUBJECT = t.skp_completion_subject 
where trunc(a.DTIME_CONTACT_realized) = dz.datum_dokonceni_klient --opraveno 18..11, DP-599 nahrazeno tt.dtime_contact_finished
and a.skp_contact_channel = 1091 -- hovor
and a.skp_contact_process_type = 1007 --kampaò
--where k.datum_kontaktu = dz.datum_dokonceni_klient 
;

--wa 25.3.2019 --není ani jeden úvìr, který by nemìl vyplnìné skp_employee

--kompletace, všechny možnosti s èím mohl jít úvìr
insert into  JN_NPOS_VOLANI_UVERY_4_GTT 
select o.*
       , CODE_COMPLETION_SUBJECT_TYPE
       , skp_signature
from JN_NPOS_VOLANI_UVERY_3_GTT o
join (--kompletace kvùli podpisu
    select pc.id_source
           , pc.skp_completion_subject 
           , pc.id_source_application as id_appl_loan
           , pc.skp_completion_subject as skp_signature
           ,pc.code_completion_subject_type
    from OWNER_DWH.f_completion_subject_at pc
    where pc.code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING', 'RHY_LOAN_COMMITMENT') and pc.skp_completion_parent ='-1'
      union
      --úvìr + BU
    select pc.id_source
           , pc.skp_completion_subject 
           , pc.id_source_application as id_appl_loan
           , pc1.skp_completion_subject as skp_signature
           ,pc.code_completion_subject_type
    from OWNER_DWH.f_completion_subject_at pc
    join OWNER_DWH.f_completion_subject_at  pc1 on (pc1.skp_completion_subject = pc.skp_completion_parent and pc1.skp_completion_parent ='-1')
    where pc.code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING','RHY_LOAN_COMMITMENT') and pc.skp_completion_parent <>'-1'
      union
      --úvìr + BU + RS
    select pc.id_source
           , pc.skp_completion_subject 
           , pc.id_source_application as id_appl_loan
           , pc1.skp_completion_parent as skp_signature
           , pc.code_completion_subject_type
    from OWNER_DWH.f_completion_subject_at pc
    join OWNER_DWH.f_completion_subject_at pc1 on (pc1.skp_completion_subject = pc.SKP_COMPLETION_PARENT and pc1.skp_completion_parent <>'-1')
    where pc.code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING','RHY_LOAN_COMMITMENT') and pc.SKP_COMPLETION_PARENT <>'-1'
 ) k on k.id_appl_loan=o.text_application_number
;  

---k podpisu došlo na poboèce, objem tedy už bude mít v plánech 

insert into  JN_NPOS_VOLANI_UVERY_5_GTT
select n.skp_loan
       , n.amt_loan_limit
       , n.text_product_description
       , n.date_loan_utilization
       , n.mesic_nacerpani
       , n.CODE_COMPLETION_SUBJECT_TYPE 
       , ct.skp_completion_subject
       , ct.SKP_BRANCH_OFFICE
       , case when ct.SKP_EMPLOYEE = 28074 then 30522 
              when ct.skp_employee = 29260 then 35408 
              when ct.skp_employee = 3347 then 36732    
                else ct.SKP_EMPLOYEE end as SKP_EMPLOYEE
       , sch.CODE_SIGNATURE_CHANNEL
       , sch.NAME_SIGNATURE_CHANNEL
       , n.code_loan_number
       , n.name_loan_status
       , trunc(ct.DTIME_SIGNATURE) as dtime_signature
from JN_NPOS_VOLANI_UVERY_4_GTT n
join owner_dwh.f_completion_signature_tt  ct on ct.SKP_COMPLETION_SUBJECT= n.SKP_SIGNATURE
join owner_dwh.dc_signature_channel sch on sch.SKP_SIGNATURE_CHANNEL=ct.SKP_CHANNEL_SIGNATURE_DELIVERY
and code_signature_channel = 'PO' and skp_branch_office != '-1'
and  n.mesic_nacerpani = to_char(p_date, 'YYYY/MM')
and ct.DTIME_SIGNATURE_VALID_TO = '1.1.3000'
;


--duplicitnì
insert into JN_NPOS_VOLANI_UVERY_6_GTT 
--podepsali
select tk.skp_loan
       , tk.amt_loan_limit
       , tk.text_product_description
       , tk.code_completion_subject_type
       , tk.skp_branch_office
       , tk.skp_employee
       , tk.mesic_nacerpani
       , tk.date_loan_utilization
       , tk.name_loan_status
      , 'podpis' as typ
      , b.name_region
from  JN_NPOS_VOLANI_UVERY_5_GTT tk
join owner_ext.ext_pos_branch b on b.skp_branch = tk.skp_branch_office  
where  tk.code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING')

union 
--obvolali
select v.skp_loan
       , v.amt_loan_limit
       , v.text_product_description
       , v.CODE_COMPLETION_SUBJECT_TYPE
       , v.skp_branch
       , v.skp_employee
       , v.mesic_nacerpani
       , v.date_loan_utilization
       , v.name_loan_status
       ,'volani' as typ
       , b.name_region
from  JN_NPOS_VOLANI_UVERY_4_GTT v
join owner_ext.ext_pos_branch b on b.skp_branch = v.skp_branch
where v.code_completion_subject_type in ('CREATE_LOAN','CREATE_REFINANCING')
;


insert into  JN_NPOS_VOLANI_UVERY_7_GTT 
WITH pom AS (
--duplicity
SELECT a.SKP_LOAN
      ,a.skp_branch_office
FROM JN_NPOS_VOLANI_UVERY_6_GTT   a 
WHERE a.skp_loan IN (
SELECT a.skp_loan FROM JN_NPOS_VOLANI_UVERY_6_GTT  a GROUP BY a.skp_loan HAVING COUNT(*)>1)
GROUP BY a.SKP_LOAN
      , a.skp_branch_office      
)
--první pøijdou s podpisem
SELECT a.skp_loan
       , a.amt_loan_limit
       , a.text_product_description
       , a.code_completion_subject_type
       , a.skp_branch_office
       , a.skp_employee
       , a.mesic_nacerpani
       , a.date_loan_utilization
       , a.name_loan_status
       , a.name_region 
       , 'PODPIS' as vstup
FROM JN_NPOS_VOLANI_UVERY_6_GTT  a
JOIN pom p ON p.SKP_LOAN = a.SKP_LOAN AND p.skp_branch_office = a.skp_branch_office and a.typ = 'podpis' 
where mesic_nacerpani = to_char(p_date, 'YYYY/MM')
;

insert into JN_NPOS_VOLANI_UVERY_7_GTT 
WITH pom AS (
--duplicity
SELECT a.SKP_LOAN
      ,a.skp_branch_office
FROM JN_NPOS_VOLANI_UVERY_6_GTT  a 
WHERE a.skp_loan IN (
SELECT a.skp_loan FROM JN_NPOS_VOLANI_UVERY_6_GTT  a GROUP BY a.skp_loan HAVING COUNT(*)>1)
GROUP BY a.SKP_LOAN
      , a.skp_branch_office      
)
SELECT a.skp_loan
       , a.amt_loan_limit
       , a.text_product_description
       , a.code_completion_subject_type
       , a.skp_branch_office
       , a.skp_employee
       , a.mesic_nacerpani
       , a.date_loan_utilization
       , a.name_loan_status  
       , a.name_region 
       ,'DUPLICITA_VOLANI' as vstup
FROM JN_NPOS_VOLANI_UVERY_6_GTT  a
JOIN pom p ON p.SKP_LOAN = a.SKP_LOAN AND p.skp_branch_office != a.skp_branch_office and a.typ != 'podpis' 
where mesic_nacerpani = to_char(p_date, 'YYYY/MM')
;

--18
insert into JN_NPOS_VOLANI_UVERY_7_GTT
WITH pom AS (
--duplicity
SELECT a.SKP_LOAN
      ,a.skp_branch_office
FROM JN_NPOS_VOLANI_UVERY_6_GTT  a 
WHERE a.skp_loan IN (
SELECT a.skp_loan FROM JN_NPOS_VOLANI_UVERY_6_GTT a GROUP BY a.skp_loan HAVING COUNT(*)>1)
GROUP BY a.SKP_LOAN
      , a.skp_branch_office      
)
SELECT a.skp_loan
       , a.amt_loan_limit
       , a.text_product_description
       , a.code_completion_subject_type
       , a.skp_branch_office
       , a.skp_employee
       , a.mesic_nacerpani
       , a.date_loan_utilization
       , a.name_loan_status
       , a.name_region  
       , 'VOLANI' as vstup
FROM JN_NPOS_VOLANI_UVERY_6_GTT a 
where not exists (select * from pom p where a.skp_loan = p.skp_loan and p.skp_branch_office = a.skp_branch_office)
and mesic_nacerpani = to_char(p_date, 'YYYY/MM') 
;

---ještì dohledám duplicity s podpisem
insert into  JN_NPOS_VOLANI_UVERY_7a_GTT
select skp_loan from JN_NPOS_VOLANI_UVERY_7_GTT  having count(skp_loan)> 1 group by skp_loan
;
--upravím podpis na duplicita podpis, shoduje se s duplicita_volani
--upravuji název podpis na duplicita_podpis
UPDATE 
      (select p.vstup 
        from JN_NPOS_VOLANI_UVERY_7_GTT   p
        where p.skp_loan in (select u.skp_loan from  JN_NPOS_VOLANI_UVERY_7a_GTT u)
        and p.vstup = 'PODPIS')
SET VSTUP = 'DUPLICITA_PODPIS'
WHERE VSTUP = 'PODPIS';


insert into JN_NPOS_VOLANI_UVERY_8_GTT
select u.mesic_nacerpani
       , b.name_region
       , b.type_branch_office as typ_pobocky_volala
       , b.skp_branch as skp_pobocky_volala
       , b.name_branch as  nazev_pobocky_volala
       , e.name_employee as zamestnanec_volal
       , e.skp_employee as skp_zam_volal
       , p.campaign_name as nazev_kampanì
       , u.skp_loan
       , u.date_loan_utilization
       , u.AMT_LOAN_LIMIT
       , u.vstup
       , case when u.vstup in ('PODPIS', 'DUPLICITA_PODPIS', 'DUPLICITA_VOLANI') then 'POS'
                   else ch.code_signature_channel end as kanal_podpisu --DP-457
                -- else 'jinde' end as kanal_podpisu
       , case when u.vstup in ('PODPIS', 'DUPLICITA_PODPIS', 'DUPLICITA_VOLANI') then b1.name_branch
                   end as pobocka_podpisu
       , case when u.vstup in ('PODPIS','DUPLICITA_PODPIS','DUPLICITA_VOLANI') then b1.skp_branch
                   end as skp_pobocky_podpisu
       , case when u.vstup in ('PODPIS', 'DUPLICITA_PODPIS','DUPLICITA_VOLANI')then e1.name_employee_opposite
              end as podepsal_asistent
       , case when u.vstup in  ('PODPIS', 'DUPLICITA_PODPIS','DUPLICITA_VOLANI') then e1.skp_employee
               end as skp_asist_podepsal
 
from JN_NPOS_VOLANI_UVERY_7_GTT u        
left join JN_NPOS_VOLANI_UVERY_3_GTT p on u.skp_loan = p.skp_loan ---použiji jen pro zijištìní názvu kampanì  a údaje ze které poboèky se volalo
left join JN_NPOS_VOLANI_UVERY_5_GTT pp on pp.skp_loan = u.skp_loan --kde se podepsalo
left join owner_ext.ext_pos_branch b  on b.skp_branch = p.skp_branch
left join owner_ext.ext_pos_employee e on e.skp_employee = p.skp_employee 
left join owner_ext.ext_pos_employee e1 on e1.skp_employee = pp.skp_employee
left join owner_ext.ext_pos_branch b1 on b1.skp_branch = pp.skp_branch_office
left join owner_dwh.f_loan_at lo on lo.SKP_loan = u.skp_loan
left join owner_dwh.f_completion_signature_tt sg on lo.SKP_COMPLETION_SUBJECT = sg.SKP_COMPLETION_SUBJECT
left join owner_dwh.dc_signature_channel ch on ch.SKP_SIGNATURE_CHANNEL = sg.SKP_CHANNEL_SIGNATURE_DELIVERY
;


insert into  JN_NPOS_VOLANI_UVERY_9_GTT 
select 'uvery' as produkt_volani
       , v.mesic_nacerpani as obdobi
       , v.name_region
       , v.typ_pobocky_volala
       , v.nazev_pobocky_volala
       , v.zamestnanec_volal
       , v.nazev_kampanì
       , v.skp_loan
       , v.date_loan_utilization as datum_nacerpani
       , v.amt_loan_limit as nacerpany_objem
       , v.vstup
       , v.kanal_podpisu
       , v.pobocka_podpisu
       , v.podepsal_asistent
       , case when vstup in ('PODPIS', 'DUPLICITA_PODPIS') then pobocka_podpisu
              when vstup in ('VOLANI', 'DUPLICITA_VOLANI') then v.nazev_pobocky_volala end as pocita_se_do_planu
from JN_NPOS_VOLANI_UVERY_8_GTT v
group by 'uvery' 
       , v.mesic_nacerpani 
       , v.name_region
       , v.typ_pobocky_volala
       , v.nazev_pobocky_volala
       , v.zamestnanec_volal
       , v.nazev_kampanì
       , v.skp_loan
       , v.date_loan_utilization 
       , v.amt_loan_limit 
       , v.vstup
       , v.kanal_podpisu
       , v.pobocka_podpisu
       , v.podepsal_asistent
       , case when vstup in ('PODPIS', 'DUPLICITA_PODPIS') then pobocka_podpisu
              when vstup in ('VOLANI', 'DUPLICITA_VOLANI') then v.nazev_pobocky_volala end
;


delete from JN_NPOS_VOLANI_UVERY_FINAL  WHERE obdobi =  to_char(p_date, 'YYYY/MM') -- vymazu aktualni datum
;


--upraveno 19.1.2019 pøidány dva sloupce
insert into JN_NPOS_VOLANI_UVERY_FINAL 
select u9.*
       ,  u3.text_product_description as popis_produktu
       , case when u3.text_product_description like 'Konsolid%' then /*round((le.amt_refinanced_sum_internal*0.25),0) +*/ case when le.AMT_LOAN_UPSELL_USED is null then re.amt_loan_upsell_used 
       else le.AMT_LOAN_UPSELL_USED end + case when le.amt_refinanced_sum_external is null then 0 else  le.amt_refinanced_sum_external end --DP-458 - nechci calc, ale used -- DP-457 zapoznámkované interky --DP - 624, prozatimní oprava, dokud nebude opraveno AS
           when u3.text_product_description like 'Standardní%' then u3.amt_loan_limit 
              when  u3.text_product_description like '%ypot%' then round(u3.amt_loan_limit *0.10,0)
                end as objem_do_odmen 
       , case when le.AMT_LOAN_UPSELL_USED is null then re.amt_loan_upsell_used else le.AMT_LOAN_UPSELL_USED end as AMT_LOAN_UPSELL_USED    --- pøidáno 19.1.2019 --DP - 624, prozatimní oprava, dokud nebude opraveno AS
       , p.id_cuid                   --- pøidáno 19.1.2019
       , u3.datum_kontaktu           --- pøidáno 29.3.2019    
       , u3.datum_dokonceni_klient   --- pøidáno 29.3.2019    
       , trunc(cs.dtime_signature) as datum_podpisu --- pøidáno 29.3.2019 
       , case when le.amt_refinanced_sum_internal is null then 0 else le.amt_refinanced_sum_internal end as amt_refinanced_sum_internal ---pøidáno 18.9.2019 DP-457 --DP - 624, prozatimní oprava, dokud nebude opraveno AS
       , case when le.amt_refinanced_sum_external is null then 0 else le.amt_refinanced_sum_external end as amt_refinanced_sum_external ---pøidáno 18.9.2019 DP-457 --DP - 624, prozatimní oprava, dokud nebude opraveno AS
from  JN_NPOS_VOLANI_UVERY_9_GTT u9
join JN_NPOS_VOLANI_UVERY_3_GTT u3 on u3.skp_loan = u9.skp_loan
left join ap_risk.mm_v_cons_int_ext_loan le on le.SKP_LOAN = u9.skp_loan
left join owner_dwh.f_application_refinance_at re on re.SKP_APPLICATION = u3.skp_application
left join JN_NPOS_VOLANI_UVERY_4_GTT u4 on u4.skp_loan = u9.skp_loan
join owner_dwh.f_completion_signature_tt cs on cs.SKP_COMPLETION_SUBJECT = u4.skp_completion_subject 
left join owner_dwh.dc_person p on p.SKP_PERSON = u3.skp_person 
WHERE NOT EXISTS (SELECT b.obdobi from JN_NPOS_VOLANI_UVERY_FINAL b WHERE u9.obdobi = b.obdobi) -- pojistka, aby se neduplikovaly zaznamy
;

----------------------------------------
--VOLÁNÍ POJIŠTÌNÍ
----------------------------------------

--uzavøené pojistky 
--podepsané pojištìní v POS

insert into JN_NPOS_VOLANI_POJISTENI_1_GTT 
select distinct k.skp_signature
       , k.id_appl_poj
       , k.skp_person
       , s.code_signature_channel
       , tt.skp_branch_office
       , case when tt.SKP_EMPLOYEE = 28074 then 30522 
              when tt.SKP_EMPLOYEE = 20077 then 30878
              when tt.skp_employee = 29260 then 35408 
              when tt.skp_employee = 3347 then 36732   
              else tt.skp_employee end as skp_employee
       , trunc(tt.DTIME_SIGNATURE) as dtime_signature
       , tt.dtime_signature_valid_from 
       , k.skp_application
from owner_dwh.f_completion_signature_tt tt
join owner_dwh.dc_signature_channel s on s.skp_signature_channel = tt.skp_channel_signature_delivery
join ( ---- samostatna poj
        select distinct pc.id_source
            ,pc.SKP_PERSON
            ,pc.skp_completion_subject
            ,pc.id_source_application        as id_appl_poj
            ,pc.skp_completion_subject       as skp_signature
            ,pc.code_completion_subject_type
            ,pc.skp_application
        from owner_dwh.f_completion_subject_at pc
        where pc.code_completion_subject_type in('INSURANCE_ENROLMENT')
             and pc.skp_completion_parent = '-1'
                   union
      ----poj s RS nebo s CL
      select distinct pc.id_source
            ,pc.SKP_PERSON
            ,pc.skp_completion_subject
            ,pc.id_source_application        as id_appl_poj
            ,pc1.skp_completion_subject      as skp_signature
            ,pc.code_completion_subject_type
            ,pc.skp_application
        from owner_dwh.f_completion_subject_at pc
        join owner_dwh.f_completion_subject_at pc1 on (pc1.skp_completion_subject =pc.skp_completion_parent and pc1.skp_completion_parent = '-1')
       where pc.code_completion_subject_type in ('INSURANCE_ENROLMENT')
            and pc.skp_completion_parent <> '-1'
                union
       ---poj s CL a RS - DP-305 - od R38 mùže jít pojištìní spoleènì s pùjèkou, která zase mùže jít spoleènì s RS
       select pc.id_source
           , pc.SKP_PERSON
           , pc.skp_completion_subject 
           , pc.id_source_application as id_appl_loan
           , pc1.skp_completion_parent as skp_signature
           , pc.code_completion_subject_type
           , pc.skp_application
       from OWNER_DWH.f_completion_subject_at pc
       join OWNER_DWH.f_completion_subject_at pc1 on (pc1.skp_completion_subject = pc.SKP_COMPLETION_PARENT and pc1.skp_completion_parent <>'-1')
       where pc.code_completion_subject_type in ('INSURANCE_ENROLMENT') 
           and pc.SKP_COMPLETION_PARENT <>'-1'
                 union
       ---poj s CL, úètem a RS - DP-305 - od R38 mùže jít pojištìní spoleènì s pùjèkou, která zase mùže jít spoleènì s RS a s tou zase mùže jít i BÚ
      select pc.id_source
           , pc.SKP_PERSON
           , pc.skp_completion_subject 
           , pc.id_source_application as id_appl_loan
           , pc2.skp_completion_parent as skp_signature
           , pc.code_completion_subject_type
           , pc.skp_application
       from OWNER_DWH.f_completion_subject_at pc
       join OWNER_DWH.f_completion_subject_at pc1 on (pc1.skp_completion_subject = pc.SKP_COMPLETION_PARENT and pc1.skp_completion_parent <>'-1')
       join OWNER_DWH.f_completion_subject_at pc2 on (pc2.skp_completion_subject = pc1.SKP_COMPLETION_PARENT and pc2.skp_completion_parent <>'-1')
       where pc.code_completion_subject_type in ('INSURANCE_ENROLMENT') 
       and pc.SKP_COMPLETION_PARENT <>'-1'
       and pc1.SKP_COMPLETION_PARENT <>'-1'
       ) k   on k.skp_signature = tt.skp_completion_subject
where --s.code_signature_channel ='PO' and tt.SKP_BRANCH_OFFICE<>-1 and
       tt.DTIME_SIGNATURE_VALID_TO='1.1.3000'
and trunc(dtime_signature) > TO_CHAR(ADD_MONTHS(LAST_DAY(p_date), -1))
;

insert into JN_NPOS_VOLANI_POJISTENI_2_GTT 
select ic.SKP_INSURANCE_CONTRACT
       , ic.SKP_PERSON
       , pp.skp_completion_subject
         , ic.dtime_contract_valid_from
       , ic.dtime_contract_valid_to
      -- , ic.TEXT_CONTRACT_NUMBER
       , to_char(pp.dtime_signature, 'YYYY/MM') as obdobi_pojisteni
       , pp.id_appl_poj
       , pp.code_signature_channel
       , pp.skp_branch_office
       , pp.skp_employee
       , trunc(pp.dtime_signature) as dtime_signature
       , pp.dtime_signature_valid_from
from JN_NPOS_VOLANI_POJISTENI_1_GTT pp
join owner_dwh.f_insurance_contract_at ic on ic.SKP_APPLICATION = pp.skp_application --nahrazeno 15.10.2018, upraveno pøímé napojení pøes skp, místo data
--join owner_dwh.f_insurance_contract_at ic on pp.skp_person = ic.SKP_PERSON  and ic.dtime_contract_valid_from = trunc(pp.dtime_signature_valid_from) + 1   --doèasná úprava kvùli chybì v DWH po pøechodu na letní èas
--join owner_dwh.f_insurance_contract_at ic on pp.skp_person = ic.SKP_PERSON  and ic.dtime_contract_valid_from = pp.dtime_signature + 1
--join owner_dwh.f_completion_subj_status_tt fcs on fcs.SKP_COMPLETION_SUBJECT = ic.SKP_COMPLETION_SUBJECT and code_completion_subj_status = 'COMPLETED'
--join owner_dwh.f_application_at fa on fa.ID_SOURCE = pp.id_appl_poj ---and fa.SKP_APPLICATION_STATUS = 1 -- schválená, upraveno 7.3. na žádost Michal M. a Pavly C.
where  trunc(pp.dtime_signature) < = to_char(p_date,'DD.MM.YYYY') 
     ---  and ic.dtime_contract_valid_to = '1.1.3000'
        and trunc(ic.dtime_contract_valid_to) > trunc(ic.dtime_contract_valid_from)
group by ic.SKP_INSURANCE_CONTRACT
       , ic.SKP_PERSON
       , pp.skp_completion_subject
         , ic.dtime_contract_valid_from
       , ic.dtime_contract_valid_to
      -- , ic.TEXT_CONTRACT_NUMBER
       , to_char(pp.dtime_signature, 'YYYY/MM') 
       , pp.id_appl_poj
       , pp.code_signature_channel
       , pp.skp_branch_office
       , pp.skp_employee
       , trunc(pp.dtime_signature)
       , pp.dtime_signature_valid_from
;

insert into JN_NPOS_VOLANI_POJISTE_2_1_GTT 
select p.skp_insurance_contract,
       p.skp_person,
       p.skp_completion_subject,
       p.dtime_contract_valid_from,
       p.dtime_contract_valid_to,
       p.obdobi_pojisteni,
       p.id_appl_poj,
       p.code_signature_channel,
       p.skp_branch_office,
       s.skp_employee,
       p.dtime_signature,
       p.dtime_signature_valid_from 
from JN_NPOS_VOLANI_POJISTENI_2_GTT p 
join owner_dwh.f_completion_signature_tt s on s.SKP_COMPLETION_SUBJECT = p.skp_completion_subject and s.skp_employee != -1
where p.code_signature_channel = 'PO' and p.skp_branch_office != - 1 and p.skp_employee = -1
;

--2182
insert into JN_NPOS_VOLANI_POJISTE_2_1_GTT
select * from JN_NPOS_VOLANI_POJISTENI_2_GTT po
where not exists (select * from JN_NPOS_VOLANI_POJISTE_2_1_GTT poj where po.skp_completion_subject = poj.skp_completion_subject)
;

insert into JN_NPOS_VOLANI_POJISTENI_3_GTT
WITH pom AS (
SELECT a.SKP_PERSON
       , a.skp_completion_subject
       , a.dtime_contract_valid_from
       , MAX(a.dtime_contract_valid_to)  AS max_dtime_valid_to
FROM JN_NPOS_VOLANI_POJISTE_2_1_GTT a 
WHERE a.skp_completion_subject IN (
SELECT a.skp_completion_subject FROM JN_NPOS_VOLANI_POJISTE_2_1_GTT a GROUP BY a.skp_completion_subject HAVING COUNT(*)>1)
GROUP BY a.SKP_PERSON
      , a.skp_completion_subject
      ,a.dtime_contract_valid_from
)
SELECT a.* FROM JN_NPOS_VOLANI_POJISTE_2_1_GTT a
JOIN pom p ON p.SKP_PERSON = a.SKP_PERSON AND p.dtime_contract_valid_from = a.dtime_contract_valid_from AND p.max_dtime_valid_to = a.dtime_contract_valid_to 
UNION
SELECT a.* FROM JN_NPOS_VOLANI_POJISTE_2_1_GTT a 
LEFT JOIN pom p ON a.skp_completion_subject = p.skp_completion_subject
WHERE p.skp_completion_subject IS NULL
;



insert into JN_NPOS_VOLANI_POJISTENI_4_GTT  
select pp.skp_person
       , pp.skp_completion_subject
       , trunc(min(tt.DTIME_VALID_FROM)) as datum_dokonceni_zadosti --- z max zmìnìno na min 28.3.2019
       , tt.SKP_EMPLOYEE_LOGGEDIN as skp_employee
from JN_NPOS_VOLANI_POJISTENI_3_GTT pp
join owner_dwh.f_completion_subject_at cs on cs.id_source_application = pp.id_appl_poj
join owner_dwh.f_application_at fa on fa.id_source = cs.id_source_application
join owner_dwh.f_application_tt tt on tt.skp_application = fa.skp_application
join owner_ext.ext_pos_employee e on tt.SKP_EMPLOYEE_LOGGEDIN = e.skp_employee
where tt.code_last_page in( 'generateDocumentationTask', 'finalizationTask')
and tt.dtime_valid_from  < = tt.dtime_valid_to
group by pp.skp_person, pp.skp_completion_subject, tt.SKP_EMPLOYEE_LOGGEDIN
;

--kde nedohledám potøebnou code_last_page tak vezmu datum podpisu
--1
insert into JN_NPOS_VOLANI_POJISTENI_4_GTT
select p.skp_person
       , p.skp_completion_subject
       , trunc(p.dtime_signature) 
       , p.skp_employee
from JN_NPOS_VOLANI_POJISTENI_3_GTT p
join owner_ext.ext_pos_employee e on p.skp_employee = e.skp_employee
where not exists (select * from  JN_NPOS_VOLANI_POJISTENI_4_GTT z where z.skp_completion_subject = p.skp_completion_subject)
;

insert into  JN_NPOS_VOLANI_POJISTENI_5_GTT 
select to_char(pd.dtime_signature, 'YYYY/MM') as obdobi
       , bk.name_region as region_volani
       , bk.type_branch_office as typ_volajici_pobocky
       , e.skp_branch as skp_pobocky_volala
       , e.name_branch as poboèka_ktera_volala
       , tt.SKP_EMPLOYEE_FINISHED as skp_asistent_volal--k.skp_finished_employee as skp_asistent_volal -- zmìnìno DP-546, 28.11.2019
       , e.name_employee as kdo_volal--k.finished_employee_lastname ||' '|| k.finished_employee_firstname as kdo_volal -- zmìnìno DP-546, 28.11.2019
       , k.campaign_name
       , pd.skp_insurance_contract 
       , pd.dtime_signature
       , pd.code_signature_channel
       , b.skp_branch as skp_pobocky_podepsala
       , b.name_branch as podepsala_poboèka
       , ee.skp_employee as skp_asistent_podepsal
       , ee.name_employee_opposite as podepsal_asistent
       , pd.skp_person
       , dz.datum_dokonceni_zadosti
       , trunc(tt.DTIME_CONTACT_FINISHED) as datum_kontaktu--k.datum_kontaktu -- zmìnìno DP-546, 28.11.2019
       , k.dtime_planned_call_finished 
       , pd.dtime_contract_valid_from
       , pd.id_appl_poj
       , pd.skp_completion_subject
from JN_NPOS_VOLANI_POJISTENI_3_GTT pd
join JN_NPOS_VOLANI_POJISTENI_4_GTT dz on dz.skp_completion_subject = pd.skp_completion_subject
join owner_dwh.f_contact_at a on a.SKP_PERSON = pd.skp_person-- pøidáno DP-546, 28.11.2019
join owner_dwh.f_contact_finish_tt tt on a.SKP_CONTACT = tt.skp_contact-- pøidáno DP-546, 28.11.2019
left join ap_CRM.JN_NPOS_KAMP_VOLANI_2018 k on tt.SKP_PERSON = k.skp_person and k.skp_finished_employee = tt.SKP_EMPLOYEE_FINISHED and trunc(k.dtime_planned_call_finished ) = trunc(tt.DTIME_CONTACT_FINISHED)
--join ap_CRM.JN_NPOS_KAMP_VOLANI_2018 k on pd.skp_person = k.skp_person -- zmìnìno DP-546, 28.11.2019
join JN_NPOS_VOLANI_POJISTE_2_1_GTT i on i.skp_insurance_contract = pd.skp_insurance_contract
left join owner_ext.ext_pos_branch b on b.skp_branch = i.skp_branch_office
left join owner_ext.ext_pos_employee ee on ee.skp_employee = i.skp_employee  
join owner_ext.ext_pos_employee e on e.skp_employee = tt.SKP_EMPLOYEE_FINISHED --k.skp_finished_employee -- zmìnìno DP-546, 28.11.2019
left join owner_ext.ext_pos_branch bk on bk.skp_branch = e.skp_branch   
where dz.datum_dokonceni_zadosti = trunc(tt.DTIME_CONTACT_FINISHED)--where dz.datum_dokonceni_zadosti = k.datum_kontaktu -- zmìnìno DP-546, 28.11.2019
and a.skp_contact_channel = 1091 -- hovor -- pøidáno DP-546, 28.11.2019
and a.skp_contact_process_type = 1007 --kampaò -- pøidáno DP-546, 28.11.2019
;

---podepsané na poboèce, již vstoupili do plánu z kanálu podpisu

insert into JN_NPOS_VOLANI_POJISTENI_6_GTT 
select * from JN_NPOS_VOLANI_POJISTENI_5_GTT 
where code_signature_channel = 'PO' 
      and skp_pobocky_podepsala != '-1'
      and obdobi = to_char(p_date, 'YYYY/MM')
;

insert into JN_NPOS_VOLANI_POJISTENI_7_GTT 
--podepsali
---podpisu došlo na poboèce, objem tedy už bude mít v plánech 
select tk.skp_insurance_contract
       , tk.skp_completion_subject
       , tk.region_volani
       , 'INSURANCE' as text_product_description
       , tk.obdobi
       , 'podpis' as typ
       , tk.skp_pobocky_podepsala
       , tk.skp_asistent_podepsal
       , tk.skp_asistent_volal
       , tk.skp_pobocky_volala
from  JN_NPOS_VOLANI_POJISTENI_6_GTT  tk

union 
--obvolali
--k podpisu došlo jinde, nejsou v plánech
select tk.skp_insurance_contract
       , tk.skp_completion_subject
       , tk.region_volani
       , 'INSURANCE' as text_product_description
       , tk.obdobi
       , 'volani' as typ
       , tk.skp_pobocky_podepsala
       , tk.skp_asistent_podepsal
       , tk.skp_asistent_volal
       , tk.skp_pobocky_volala
from  JN_NPOS_VOLANI_POJISTENI_5_GTT   tk
;

insert into  JN_NPOS_VOLANI_POJISTENI_8_GTT
WITH pom AS (
--duplicity
SELECT a.skp_insurance_contract
       , a.skp_completion_subject
       , a.skp_pobocky_podepsala
FROM JN_NPOS_VOLANI_POJISTENI_7_GTT   a 
WHERE a.skp_insurance_contract IN (
SELECT a.skp_insurance_contract FROM JN_NPOS_VOLANI_POJISTENI_7_GTT  a GROUP BY a.skp_insurance_contract HAVING COUNT(*)>1)
GROUP BY a.skp_insurance_contract
       , a.skp_completion_subject
      , a.skp_pobocky_podepsala      
)
--první pøijdou s podpisem
SELECT a.skp_insurance_contract
       , a.skp_completion_subject
       , a.text_product_description
       , a.skp_pobocky_podepsala
       , a.skp_asistent_podepsal
       , a.obdobi
       , a.region_volani
       , 'PODPIS' as vstup
FROM JN_NPOS_VOLANI_POJISTENI_7_GTT  a
JOIN pom p ON p.Skp_Insurance_Contract = a.Skp_Insurance_Contract AND p.skp_pobocky_podepsala = a.skp_pobocky_podepsala and a.typ = 'podpis'  ---23
where obdobi = to_char(p_date, 'YYYY/MM')
;

insert into JN_NPOS_VOLANI_POJISTENI_8_GTT 
WITH pom AS (
--duplicity
SELECT  a.skp_insurance_contract
       , a.skp_completion_subject
       , a.skp_pobocky_podepsala
FROM JN_NPOS_VOLANI_POJISTENI_7_GTT  a 
WHERE a.skp_insurance_contract IN (
SELECT a.skp_insurance_contract FROM JN_NPOS_VOLANI_POJISTENI_7_GTT  a GROUP BY a.skp_insurance_contract  HAVING COUNT(*)>1)
GROUP BY a.skp_insurance_contract
       , a.skp_completion_subject
       , a.skp_pobocky_podepsala      
)
SELECT  a.skp_insurance_contract
       , a.skp_completion_subject
       , a.text_product_description
       , a.skp_pobocky_podepsala
       , a.skp_asistent_podepsal
       , a.obdobi
       , a.region_volani
       , 'DUPLICITA_VOLANI' as vstup
FROM JN_NPOS_VOLANI_POJISTENI_7_GTT  a
JOIN pom p ON p.Skp_Insurance_Contract = a.Skp_Insurance_Contract AND p.skp_pobocky_podepsala != a.skp_pobocky_podepsala and a.typ != 'podpis' --4
where obdobi = to_char(p_date, 'YYYY/MM')
;

--18
insert into JN_NPOS_VOLANI_POJISTENI_8_GTT
WITH pom AS (
--duplicity
SELECT  a.skp_insurance_contract
       , a.skp_completion_subject
       , a.skp_pobocky_podepsala
FROM JN_NPOS_VOLANI_POJISTENI_7_GTT  a 
WHERE a.skp_insurance_contract IN (
SELECT a.skp_insurance_contract FROM JN_NPOS_VOLANI_POJISTENI_7_GTT  a GROUP BY a.skp_insurance_contract  HAVING COUNT(*)>1)
GROUP BY a.skp_insurance_contract
       , a.skp_completion_subject
       , a.skp_pobocky_podepsala      
)
SELECT a.skp_insurance_contract
       , a.skp_completion_subject
       , a.text_product_description
       , a.skp_pobocky_podepsala
       , a.skp_asistent_podepsal
       , a.obdobi
       , a.region_volani
       , 'VOLANI' as vstup
FROM JN_NPOS_VOLANI_POJISTENI_7_GTT  a 
where not exists (select * from pom p where a.skp_insurance_contract = p.skp_insurance_contract and p.skp_pobocky_podepsala = a.skp_pobocky_podepsala)--11
and obdobi = to_char(p_date, 'YYYY/MM') 
;

---ještì dohledám duplicity s podpisem
insert into  JN_NPOS_VOLANI_POJISTENI_9_GTT
select skp_insurance_contract from JN_NPOS_VOLANI_POJISTENI_8_GTT  having count(skp_insurance_contract)> 1 group by skp_insurance_contract
;

--upravím podpis na duplicita podpis, shoduje se s duplicita_volani
--upravuji název podpis na duplicita_podpis
UPDATE 
      (select p.vstup 
        from JN_NPOS_VOLANI_POJISTENI_8_GTT   p
        where p.skp_insurance_contract in (select u.skp_insurance_contract from  JN_NPOS_VOLANI_POJISTENI_9_GTT u)
        and p.vstup = 'PODPIS')
SET VSTUP = 'DUPLICITA_PODPIS'
WHERE VSTUP = 'PODPIS'
;

insert into JN_NPOS_VOLANI_POJISTEN_10_GTT
select 'Pojištìní' as produkt_pojisteni
       , u.obdobi
       , b.name_region
       , b.type_branch_office as typ_pobocky_volala
       , b.skp_branch as skp_pobocky_volala
       , b.name_branch as  nazev_pobocky_volala
       , e.name_employee as zamestnanec_volal
       , e.skp_employee as skp_zam_volal
       , p.campaign_name as nazev_kampanì
       , u.skp_insurance_contract
       , pp.dtime_signature
       , u.vstup
     --  , pp.code_signature_channel
       , case when u.vstup in ('PODPIS', 'DUPLICITA_PODPIS', 'DUPLICITA_VOLANI') then 'POS'
                   else ch.code_signature_channel end as kanal_podpisu --DP-457
                   --else 'jinde' end as kanal_podpisu 
       , case when u.vstup in ('PODPIS', 'DUPLICITA_PODPIS', 'DUPLICITA_VOLANI') then b1.name_branch
                   end as pobocka_podpisu
       , case when u.vstup in ('PODPIS','DUPLICITA_PODPIS','DUPLICITA_VOLANI') then b1.skp_branch
                   end as skp_pobocky_podpisu
       , case when u.vstup in ('PODPIS', 'DUPLICITA_PODPIS','DUPLICITA_VOLANI')then e1.name_employee_opposite
              end as podepsal_asistent
       , case when u.vstup in  ('PODPIS', 'DUPLICITA_PODPIS','DUPLICITA_VOLANI') then e1.skp_employee
               end as skp_asist_podepsal
       , p.datum_dokonceni_zadosti
       , p.datum_kontaktu
       , p.id_appl_poj
from JN_NPOS_VOLANI_POJISTENI_8_GTT u        
left join JN_NPOS_VOLANI_POJISTENI_5_GTT p on u.skp_insurance_contract = p.skp_insurance_contract ---použiji jen pro zijištìní názvu kampanì  a údaje ze které poboèky se volalo
left join JN_NPOS_VOLANI_POJISTENI_1_GTT pp on pp.skp_completion_subject = u.skp_completion_subject --kde se podepsalo
left join owner_ext.ext_pos_branch b  on b.skp_branch = p.skp_pobocky_volala
left join owner_ext.ext_pos_employee e on e.skp_employee = p.skp_asistent_volal
left join owner_ext.ext_pos_employee e1 on e1.skp_employee = pp.skp_employee
left join owner_ext.ext_pos_branch b1 on b1.skp_branch = pp.skp_branch_office
left join owner_dwh.f_completion_signature_tt sg on u.SKP_COMPLETION_SUBJECT = sg.SKP_COMPLETION_SUBJECT
left join owner_dwh.dc_signature_channel ch on ch.SKP_SIGNATURE_CHANNEL = sg.SKP_CHANNEL_SIGNATURE_DELIVERY
;

insert into JN_NPOS_VOLANI_POJISTEN_11_GTT
select  v.produkt_pojisteni as produkt_volani
       , v.obdobi
       , v.name_region as region
       , v.typ_pobocky_volala
       , v.nazev_pobocky_volala
       , v.zamestnanec_volal
       , v.nazev_kampane
       , v.skp_insurance_contract as skp_insurance
       , v.dtime_signature as datum_podpisu
       , v.datum_dokonceni_zadosti
       , v.datum_kontaktu
       , v.id_appl_poj 
       , v.vstup
       , v.kanal_podpisu
       , v.pobocka_podpisu
       , v.podepsal_asistent
       , case when vstup in ('PODPIS', 'DUPLICITA_PODPIS') then pobocka_podpisu
              when vstup in ('VOLANI', 'DUPLICITA_VOLANI') then v.nazev_pobocky_volala end as pocita_se_do_planu
from JN_NPOS_VOLANI_POJISTEN_10_GTT v
group by v.produkt_pojisteni 
       , v.obdobi
       , v.name_region 
       , v.typ_pobocky_volala
       , v.nazev_pobocky_volala
       , v.zamestnanec_volal
       , v.nazev_kampane
       , v.skp_insurance_contract
       , v.dtime_signature 
       , v.datum_dokonceni_zadosti
       , v.datum_kontaktu
       , v.id_appl_poj 
       , v.vstup
       , v.kanal_podpisu
       , v.pobocka_podpisu
       , v.podepsal_asistent
       , case when vstup in ('PODPIS', 'DUPLICITA_PODPIS') then pobocka_podpisu
              when vstup in ('VOLANI', 'DUPLICITA_VOLANI') then v.nazev_pobocky_volala end
;


delete from JN_NPOS_VOLANI_POJISTENI_FINAL  WHERE obdobi =  to_char(p_date, 'YYYY/MM') -- vymazu aktualni datum
;

insert into JN_NPOS_VOLANI_POJISTENI_FINAL 
select a.*, p.id_cuid 
from JN_NPOS_VOLANI_POJISTEN_11_GTT a
join owner_dwh.f_insurance_contract_at i on i.SKP_INSURANCE_CONTRACT = a.skp_insurance 
join owner_dwh.dc_person p on p.SKP_PERSON = i.skp_person 
WHERE NOT EXISTS (SELECT b.obdobi from JN_NPOS_VOLANI_POJISTENI_FINAL b WHERE a.obdobi = b.obdobi) -- pojistka, aby se neduplikovaly zaznamy
;

----------------------------------------
--VOLÁNÍ KONTOKORENT
----------------------------------------

---zkompletovaný KTK = zapnutý kontokorent


insert into JN_NPOS_VOLANI_KTK_1_GTT 
select sa.SKP_COMPLETION_SUBJECT
       , a.SKP_OVERDRAFT
       , sa.SKP_PERSON
       , sa.id_source_application as id_appl_ktk
       , sa.SKP_APPLICATION
       , sa.SKP_COMPLETION_SUBJECT as skp_signature
       , sa.SKP_COMPLETION_PARENT
       , ct.code_COMPLETION_SUBJECT_TYPE
       , cs.CODE_COMPLETION_SUBJECT_STATUS
       , sa.DTIME_COMPLETION_FINISHED
       , a.DATE_OVERDRAFT_UTILIZATION
       , case when tt.SKP_EMPLOYEE = 3347 then 36732 else tt.skp_employee end as skp_employee
       , tt.SKP_BRANCH_OFFICE
       , trunc(tt.DTIME_SIGNATURE) as dtime_signature
       , tt.dtime_signature_valid_from 
       , to_char (sa.DTIME_COMPLETION_FINISHED, 'YYYY/MM') as obdobi
       , cos.NAME_OVERDRAFT_STATUS
       , so.CODE_SIGNATURE_CHANNEL
from owner_dwh.f_completion_subject_at sa 
join owner_dwh.cl_completion_subject_status cs on cs.SKP_COMPLETION_SUBJECT_STATUS = sa.SKP_COMPLETION_SUBJECT_STATUS 
join owner_dwh.f_overdraft_at a on a.SKP_COMPLETION_SUBJECT = sa.skp_completion_subject 
join owner_dwh.cl_completion_subject_type ct on ct.SKP_COMPLETION_SUBJECT_TYPE = sa.SKP_COMPLETION_SUBJECT_TYPE 
join owner_dwh.f_completion_signature_tt tt on tt.SKP_COMPLETION_SUBJECT = sa.SKP_COMPLETION_SUBJECT and trunc(tt.dtime_signature) != '1.1.3000' and tt.DTIME_SIGNATURE_VALID_TO = '1.1.3000'
left join owner_dwh.dc_signature_channel so on so.SKP_SIGNATURE_CHANNEL = tt.SKP_CHANNEL_SIGNATURE_DELIVERY 
join owner_dwh.cl_overdraft_status cos on cos.SKP_OVERDRAFT_STATUS = a.SKP_OVERDRAFT_STATUS
where cs.CODE_COMPLETION_SUBJECT_STATUS = 'COMPLETED'
and to_char(sa.DTIME_COMPLETION_FINISHED, 'YYYY/MM') =  to_char(p_date,'YYYY/MM')
;

--wa 25.3.2019
insert into JN_NPOS_VOLANI_KTK_1_1_GTT
select k.skp_completion_subject,
       k.skp_overdraft,
       k.skp_person,
       k.id_appl_ktk,
       k.skp_application,
       k.skp_signature,
       k.skp_completion_parent,
       k.code_completion_subject_type,
       k.code_completion_subject_status,
       k.dtime_completion_finished,
       k.date_overdraft_utilization,
       s.skp_employee,
       k.skp_branch_office,
       k.dtime_signature,
       k.dtime_signature_valid_from,
       k.obdobi,
       k.name_overdraft_status,
       k.code_signature_channel 
from  JN_NPOS_VOLANI_KTK_1_GTT k
join owner_dwh.f_completion_signature_tt s on s.SKP_COMPLETION_SUBJECT = k.skp_completion_subject and s.skp_employee != -1 
where k.skp_branch_office != -1 and k.skp_employee = -1
;

--wa 25.3.2019
insert into  JN_NPOS_VOLANI_KTK_1_1_GTT 
select * from  JN_NPOS_VOLANI_KTK_1_GTT  po
where not exists (select * from  JN_NPOS_VOLANI_KTK_1_1_GTT  poj where po.skp_completion_subject = poj.skp_completion_subject)
;



--- øeší duplicity kontokorentù, klienti si založí v jednom mìsíci více KTK jeden ukonèí a další si založí

insert into JN_NPOS_VOLANi_KTK_2_GTT 
with pom as (
select skp_person, count(*) from JN_NPOS_VOLANi_KTK_1_1_GTT  group by skp_person having count(*)> 1    )
select v.* from JN_NPOS_VOLANi_KTK_1_1_GTT  v join pom p on v.skp_person = p.skp_person and NAME_OVERDRAFT_STATUS = 'Aktivní' 
union 
select v1.* from JN_NPOS_VOLANi_KTK_1_1_GTT v1 where not exists (select * from  pom p1 where  v1.skp_person = p1.skp_person )
;

---datum dokonèení žádosti ze strany klienta 
---1371
insert into JN_NPOS_VOLANi_KTK_3_GTT
select vk.skp_application 
       , vk.skp_overdraft
       , vk.skp_completion_subject
       , trunc(min(att.DTIME_VALID_FROM)) as datum_dokonceni_klient -----z max zmìnìno na min 28.3.2019
       , att.SKP_EMPLOYEE_LOGGEDIN as skp_employee --DP - 546, pøidáno 27.11.2019
from  JN_NPOS_VOLANi_KTK_2_GTT vk
join owner_dwh.f_application_tt  att on vk.skp_application = att.SKP_APPLICATION
join owner_ext.ext_pos_employee e on e.skp_employee = att.SKP_EMPLOYEE_LOGGEDIN
where code_last_page in ('waitForPreScoringResultTask', 'waitForLoanScoringResultTask', 'waitForScoringResultTask')--('generateDocumentationTask', 'finalizationTask') --DP - 546, zmìnìno 27.11.2019
and att.dtime_valid_from  < = att.dtime_valid_to
group by vk.skp_application, vk.skp_overdraft, vk.skp_completion_subject, att.skp_employee_loggedin
;

--kde nedohledám potøebnou code_last_page tak vezmu datum podpisu
--0

insert into  JN_NPOS_VOLANi_KTK_3_GTT
select p.skp_application
       , p.skp_overdraft
       , p.skp_completion_subject
       , trunc(p.dtime_signature) 
       , p.skp_employee
from JN_NPOS_VOLANi_KTK_2_GTT  p
join owner_ext.ext_pos_employee e on e.skp_employee = p.skp_employee
where not exists (select * from  JN_NPOS_VOLANI_KTK_3_GTT z where z.skp_completion_subject = p.skp_completion_subject)
;


---propojení kontaktu z kampanì a KTK

insert into  JN_NPOS_VOLANi_KTK_4_GTT
select vk.obdobi
       , vk.skp_completion_subject
       , vk.skp_person
       , vk.skp_overdraft
       , vk.id_appl_ktk
       , vk.skp_application
       , vk.skp_signature
       , vk.code_completion_subject_type
       , vk.code_completion_subject_status
       , vk.dtime_signature
       , vk.dtime_signature_valid_from
       , vk.dtime_completion_finished
       , vk.date_overdraft_utilization
       , vk.code_signature_channel as kanal_podpisu
       , vk.skp_employee as skp_empl_podpis
       , ee.name_employee_opposite as zamestnanec_podepsal
       , vk.skp_branch_office as skp_pobocky_podepsala
       , b.name_branch as pobocka_podepsala
       , b.name_region as region_podepsal
       , v.datum_dokonceni_klient
       , kam.skf_message
       , kam.dtime_created
       , kam.dtime_planned_call_finished
       , trunc(tt.DTIME_CONTACT_FINISHED) as datum_kontaktu--kam.datum_kontaktu zmìnìno DP - 546 27.11.2019
       , tt.SKP_EMPLOYEE_FINISHED as skp_finished_employee --kam.skp_finished_employee zmìnìno DP - 546 27.11.2019
       , ee1.name_employee as zamestnanec_volal--kam.finished_employee_lastname ||''|| kam.finished_employee_firstname as zamestnanec_volal zmìnìno DP - 546 27.11.2019
       , b1.skp_branch as skp_branch_volajici
       , b1.name_branch as pobocka_volala
       , b1.name_region as region_volal
       , kam.queue_code
       , kam.campaign_name
       , kam.campaign_code
from JN_NPOS_VOLANi_KTK_2_GTT vk
join JN_NPOS_VOLANi_KTK_3_GTT v on v.skp_application = vk.skp_application
join owner_dwh.f_contact_at a on a.SKP_PERSON = vk.skp_person -- pøidáno DP - 546 27.11.2019
join owner_dwh.f_contact_finish_tt tt on a.SKP_CONTACT = tt.skp_contact --pøidáno DP - 546 27.11.2019
left join ap_crm.JN_NPOS_KAMP_VOLANI_2018  kam on kam.skp_person = vk.skp_person and v.skp_employee = kam.skp_finished_employee and trunc(kam.dtime_planned_call_finished) = v.datum_dokonceni_klient--zmìnìno DP - 546 27.11.2019
left join owner_ext.ext_pos_employee ee on ee.skp_employee = vk.skp_employee
left join owner_ext.ext_pos_branch b  on b.skp_branch = vk.skp_branch_office
join owner_ext.ext_pos_employee ee1 on ee1.skp_employee = tt.SKP_EMPLOYEE_FINISHED -- kam.skp_finished_employee --zmìnìno DP - 546 27.11.2019
join owner_ext.ext_pos_branch b1 on ee1.skp_branch = b1.skp_branch --zmìnìno DP - 546 27.11.2019
where v.datum_dokonceni_klient = trunc(tt.DTIME_CONTACT_FINISHED)--kam.datum_kontaktu zmìnìno DP - 546 27.11.2019
and a.skp_contact_channel = 1091 -- hovor pøidáno DP - 546 27.11.2019
and a.skp_contact_process_type = 1007 --kampaò pøidáno DP - 546 27.11.2019
;


insert into JN_NPOS_VOLANi_KTK_5_GTT
select * 
from JN_NPOS_VOLANi_KTK_4_GTT
where kanal_podpisu = 'PO' 
      and skp_pobocky_podepsala != '-1'
      and obdobi = to_char(p_date, 'YYYY/MM')
;

--podepsali
---k podpisu došlo na poboèce, pøiøazen dle podpisu
insert into JN_NPOS_VOLANI_KTK_6_GTT 
select tk.skp_completion_subject
       , tk.skp_overdraft
       , tk.skp_person
       , tk.skp_application
       , tk.skp_signature
       , tk.id_appl_ktk
       , 'KTK' as text_product_description
       , tk.obdobi
       , 'podpis' as typ
       , tk.region_podepsal
       , tk.skp_pobocky_podepsala
       , tk.skp_empl_podpis as skp_asistent_podepsal
       , tk.region_volal
       , tk.skp_branch_volajici as skp_pobocka_volala
       , tk.skp_finished_employee as skp_asistent_volal
from   JN_NPOS_VOLANI_KTK_5_GTT  tk

union 
--obvolali
--k podpisu došlo jinde, nejsou v plánech
select  tk.skp_completion_subject
        , tk.skp_overdraft
       , tk.skp_person
       , tk.skp_application
       , tk.skp_signature
       , tk.id_appl_ktk
       , 'KTK' as text_product_description
       , tk.obdobi
       , 'volani' as typ
       , tk.region_podepsal
       , tk.skp_pobocky_podepsala
       , tk.skp_empl_podpis as skp_asistent_podepsal
       , tk.region_volal
       , tk.skp_branch_volajici as skp_pobocka_volala
       , tk.skp_finished_employee as skp_asistent_volal
from   JN_NPOS_VOLANI_KTK_4_GTT    tk
;

insert into JN_NPOS_VOLANI_KTK_7_GTT
WITH pom AS (
--duplicity
SELECT a.skp_overdraft
       , a.skp_completion_subject
       , a.skp_pobocky_podepsala
FROM JN_NPOS_VOLANI_KTK_6_GTT  a 
WHERE a.skp_overdraft IN (
SELECT a.skp_overdraft FROM JN_NPOS_VOLANI_KTK_6_GTT  a GROUP BY a.skp_overdraft HAVING COUNT(*)>1)
GROUP BY a.skp_overdraft
       , a.skp_completion_subject
      , a.skp_pobocky_podepsala      
)
--první pøijdou s podpisem
SELECT a.skp_overdraft
       , a.skp_completion_subject
       , a.text_product_description
       , a.skp_pobocky_podepsala
       , a.skp_asistent_podepsal
       , a.obdobi
       , a.region_volal
       , 'PODPIS' as vstup
FROM JN_NPOS_VOLANI_KTK_6_GTT   a
JOIN pom p ON p.Skp_overdraft = a.Skp_Overdraft AND p.skp_pobocky_podepsala= a.skp_pobocky_podepsala and a.typ = 'podpis'  ---23
where obdobi = to_char(p_date, 'YYYY/MM')
;

insert into JN_NPOS_VOLANI_KTK_7_GTT  
WITH pom AS (
--duplicity
SELECT  a.skp_overdraft
       , a.skp_completion_subject
       , a.skp_pobocky_podepsala
       , a.typ
FROM JN_NPOS_VOLANI_KTK_6_GTT   a 
WHERE a.skp_overdraft IN (
SELECT a.skp_overdraft FROM JN_NPOS_VOLANI_KTK_6_GTT  a GROUP BY a.skp_overdraft  HAVING COUNT(*)>1)
GROUP BY a.skp_overdraft
       , a.skp_completion_subject
       , a.skp_pobocky_podepsala 
       , a.typ     
)
SELECT  a.skp_overdraft
       , a.skp_completion_subject
       , a.text_product_description
       , a.skp_pobocky_podepsala
       , a.skp_asistent_podepsal
       , a.obdobi
       , a.region_volal
       , 'DUPLICITA_VOLANI' as vstup
FROM JN_NPOS_VOLANI_KTK_6_GTT   a
JOIN pom p ON p.Skp_overdraft = a.Skp_overdraft AND p.skp_pobocky_podepsala != a.skp_pobocka_volala and p.typ != 'podpis' 
where obdobi = to_char(sysdate-1, 'YYYY/MM')
group by a.skp_overdraft
       , a.skp_completion_subject
       , a.text_product_description
       , a.skp_pobocky_podepsala
       , a.skp_asistent_podepsal
       , a.obdobi
       , a.region_volal
       , 'DUPLICITA_VOLANI' 
;


insert into JN_NPOS_VOLANI_KTK_7_GTT 
WITH pom AS (
--duplicity
SELECT  a.skp_overdraft
       , a.skp_completion_subject
       , a.skp_pobocky_podepsala
FROM JN_NPOS_VOLANI_KTK_6_GTT  a 
WHERE a.skp_overdraft IN (
SELECT a.skp_overdraft FROM JN_NPOS_VOLANI_KTK_6_GTT a GROUP BY a.skp_overdraft  HAVING COUNT(*)>1)
GROUP BY a.skp_overdraft
       , a.skp_completion_subject
       , a.skp_pobocky_podepsala      
)
SELECT a.skp_overdraft
       , a.skp_completion_subject
       , a.text_product_description
       , a.skp_pobocky_podepsala
       , a.skp_asistent_podepsal
       , a.obdobi
       , a.region_volal
       , 'VOLANI' as vstup
FROM JN_NPOS_VOLANI_KTK_6_GTT a 
where not exists (select * from pom p where a.skp_overdraft = p.skp_overdraft and p.skp_pobocky_podepsala = a.skp_pobocky_podepsala)
and obdobi = to_char(p_date, 'YYYY/MM') 
;

---ještì dohledám duplicity s podpisem
insert into JN_NPOS_VOLANI_KTK_8_GTT 
select skp_overdraft from JN_NPOS_VOLANI_KTK_7_GTT   having count(skp_overdraft)> 1 group by skp_overdraft
;


--upravím podpis na duplicita podpis, shoduje se s duplicita_volani
--upravuji název podpis na duplicita_podpis
UPDATE 
      (select p.vstup 
        from JN_NPOS_VOLANI_KTK_7_GTT   p
        where p.skp_overdraft in (select u.skp_overdraft from  JN_NPOS_VOLANI_KTK_8_GTT u)
        and p.vstup = 'PODPIS')
SET VSTUP = 'DUPLICITA_PODPIS'
WHERE VSTUP = 'PODPIS'
;

insert into JN_NPOS_VOLANI_KTK_9_GTT 
select 'KTK' as produkt
       , u.obdobi
       , b.name_region as region_volal
       , b.type_branch_office as typ_pobocky_volala
       , b.skp_branch as skp_pobocky_volala
       , b.name_branch as  nazev_pobocky_volala
       , e.skp_employee as skp_asistent_volal
       , e.name_employee_opposite as asistent_volal
       , p.campaign_name as nazev_kampanì
       , pp.dtime_signature
       , u.vstup
     --  , pp.code_signature_channel
       , case when u.vstup in ('PODPIS', 'DUPLICITA_PODPIS', 'DUPLICITA_VOLANI') then 'POS'
                   else 'jinde' end as kanal_podpisu
       , case when u.vstup in ('PODPIS', 'DUPLICITA_PODPIS', 'DUPLICITA_VOLANI') then b1.name_branch
                   end as pobocka_podpisu
       , case when u.vstup in ('PODPIS','DUPLICITA_PODPIS','DUPLICITA_VOLANI') then b1.skp_branch
                   end as skp_pobocky_podpisu
       , case when u.vstup in ('PODPIS', 'DUPLICITA_PODPIS','DUPLICITA_VOLANI')then e1.name_employee_opposite
              end as podepsal_asistent
       , case when u.vstup in  ('PODPIS', 'DUPLICITA_PODPIS','DUPLICITA_VOLANI') then e1.skp_employee
               end as skp_asist_podepsal
       , p.kanal_podpisu as kod_kanalu_podpisu
       , p.datum_dokonceni_klient
       , p.datum_kontaktu
       , u.skp_overdraft
       , p.skp_completion_subject
       , p.skp_person
       , p.id_appl_ktk
from JN_NPOS_VOLANI_KTK_7_GTT u        
left join JN_NPOS_VOLANI_KTK_4_GTT p on u.skp_overdraft = p.skp_overdraft ---použiji jen pro zijištìní názvu kampanì  a údaje ze které poboèky se volalo
left join JN_NPOS_VOLANI_KTK_1_1_GTT pp on pp.skp_completion_subject = u.skp_completion_subject --kde se podepsalo
left join owner_ext.ext_pos_branch b  on b.skp_branch = p.skp_branch_volajici
left join owner_ext.ext_pos_employee e on e.skp_employee = p.skp_finished_employee
left join owner_ext.ext_pos_employee e1 on e1.skp_employee = pp.skp_employee
left join owner_ext.ext_pos_branch b1 on b1.skp_branch = pp.skp_branch_office
;


insert into  JN_NPOS_VOLANI_KTK_10_GTT 
select  v.produkt as produkt_volani
       , v.obdobi
       , v.region_volal 
       , v.typ_pobocky_volala
       , v.nazev_pobocky_volala
       , v.asistent_volal
       , v.nazev_kampanì
       , v.skp_overdraft 
       , v.dtime_signature as datum_podpisu
       , v.datum_dokonceni_klient
       , v.datum_kontaktu
       , v.id_appl_ktk
       , v.vstup
       , v.kod_kanalu_podpisu as kanal_pospisu
       , v.pobocka_podpisu
       , v.podepsal_asistent
       , case when vstup in ('PODPIS', 'DUPLICITA_PODPIS') then v.pobocka_podpisu
              when vstup in ('VOLANI', 'DUPLICITA_VOLANI') then v.nazev_pobocky_volala end as pocita_se_do_planu
from JN_NPOS_VOLANI_KTK_9_GTT v
group by  v.produkt 
       , v.obdobi
       , v.region_volal 
       , v.typ_pobocky_volala
       , v.nazev_pobocky_volala
       , v.asistent_volal
       , v.nazev_kampanì
       , v.skp_overdraft 
       , v.dtime_signature 
       , v.datum_dokonceni_klient
       , v.datum_kontaktu
       , v.id_appl_ktk
       , v.vstup
       , v.kod_kanalu_podpisu 
       , v.pobocka_podpisu
       , v.podepsal_asistent
       , case when vstup in ('PODPIS', 'DUPLICITA_PODPIS') then v.pobocka_podpisu
              when vstup in ('VOLANI', 'DUPLICITA_VOLANI') then v.nazev_pobocky_volala end
;



delete from JN_NPOS_VOLANI_KTK_FINAL  WHERE obdobi =  to_char(p_date, 'YYYY/MM') -- vymazu aktualni datum
;

insert into JN_NPOS_VOLANI_KTK_FINAL 
select a.*
       , d.id_cuid 
from JN_NPOS_VOLANI_KTK_10_GTT a
join owner_dwh.f_overdraft_at f on a.SKP_OVERDRAFT = f.skp_overdraft 
join owner_dwh.dc_person d on d.skp_person = f.skp_person  
WHERE NOT EXISTS (SELECT b.obdobi from JN_NPOS_VOLANI_KTK_FINAL b WHERE a.obdobi = b.obdobi) -- pojistka, aby se neduplikovaly zaznamy
;

-- doplnìno dle DP-340, již poèítány také mobility z kampaòového volání nasazeno 31.7.2019
----------------------------------------
--MOBILITY
----------------------------------------

--podepsané žádosti o mobilitu
Insert into MKL_NPOS_VOLANI_MOB_1_GTT
select distinct k.skp_completion_subject
       , k.id_appl_mob
       , k.skp_person
       , s.code_signature_channel
       , tt.skp_branch_office
       , case when tt.skp_employee = 32612 then 34295 
              when tt.skp_employee = 3347 then 36732
                   else tt.skp_employee end as skp_employee
       , trunc(tt.DTIME_SIGNATURE) as dtime_signature
       , tt.dtime_signature_valid_from
       , tt.dtime_signature_valid_to
       , k.skp_application
       , k.code_completion_subject_type
       , k.skp_signature
       , to_char (dtime_signature, 'YYYY/MM') as obdobi
from owner_dwh.f_completion_signature_tt tt
join owner_dwh.dc_signature_channel s on s.skp_signature_channel = tt.skp_channel_signature_delivery
join ( ---- samostatna mob
        select distinct pc.id_source
            ,pc.SKP_PERSON
            ,pc.skp_completion_subject
            ,pc.id_source_application        as id_appl_mob
            ,pc.skp_completion_subject       as skp_signature
            ,pc.code_completion_subject_type
            ,pc.SKP_APPLICATION
        from owner_dwh.f_completion_subject_at pc
        where pc.code_completion_subject_type in('CREATE_MOBILITY')
             and pc.skp_completion_parent = '-1'
              union
      ----mob s RS
      select distinct pc.id_source
            ,pc.SKP_PERSON
            ,pc.skp_completion_subject
            ,pc.id_source_application        as id_appl_mob
            ,pc1.skp_completion_subject      as skp_signature
            ,pc.code_completion_subject_type
            ,pc.SKP_APPLICATION
        from owner_dwh.f_completion_subject_at pc
        join owner_dwh.f_completion_subject_at pc1 on (pc1.skp_completion_subject =pc.skp_completion_parent and pc1.skp_completion_parent = '-1')
       where pc.code_completion_subject_type in ('CREATE_MOBILITY')
            and pc.skp_completion_parent <> '-1'
       ) k   on k.skp_signature = tt.skp_completion_subject
where tt.DTIME_SIGNATURE_VALID_TO='1.1.3000'
and trunc(dtime_signature) > TO_CHAR(ADD_MONTHS(LAST_DAY(p_date), -1))
and trunc(dtime_signature) ! = '1.1.3000'
;


/*
Create table MKL_MOB1_1 as
select k.skp_completion_subject,
       k.id_appl_mob,
       k.skp_person,
       k.code_signature_channel,
       k.skp_branch_office,
       s.skp_employee,
       k.dtime_signature,
       k.dtime_signature_valid_from,
       k.dtime_signature_valid_to,
       k.skp_application,
       k.code_completion_subject_type,
       k.skp_signature,
       k.obdobi
from  MKL_MOB1 k
join owner_dwh.f_completion_signature_tt s on s.SKP_COMPLETION_SUBJECT = k.skp_completion_subject and s.skp_employee != -1 
where k.skp_branch_office != -1 and k.skp_employee = -1
;

SELECT * FROM owner_dwh.f_completion_signature_tt
where skp_employee <> -1


insert into  MKL_MOB1_1
select * from  MKL_MOB1  po
where not exists (select * from  MKL_MOB1_1  poj where po.skp_completion_subject = poj.skp_completion_subject)
;*/


--øešení duplicit mobilit
insert into MKL_NPOS_VOLANI_MOB_2_GTT
with pom as (
select skp_person, count(*) from MKL_NPOS_VOLANI_MOB_1_GTT  group by skp_person having count(*)> 1    )
select v.* from MKL_NPOS_VOLANI_MOB_1_GTT v join pom p on v.skp_person = p.skp_person
union 
select v1.* from MKL_NPOS_VOLANI_MOB_1_GTT v1 where not exists (select * from  pom p1 where  v1.skp_person = p1.skp_person )
;

--dokonèené klientem
insert into MKL_NPOS_VOLANI_MOB_3_GTT
select vk.skp_application 
       , vk.skp_completion_subject
       , vk.skp_person
       , trunc(min(att.DTIME_VALID_FROM)) as datum_dokonceni_klient -----z max zmìnìno na min 28.3.2019
       , case when code_last_page in ('generalContractSummaryTask', 'finalizationTask')  then att.SKP_EMPLOYEE_LOGGEDIN else att.SKP_EMPLOYEE_LOGGEDIN end as skp_employee
from  MKL_NPOS_VOLANI_MOB_2_GTT vk
join owner_dwh.f_application_tt  att on vk.skp_application = att.SKP_APPLICATION
join owner_ext.ext_pos_employee e on e.skp_employee = att.SKP_EMPLOYEE_LOGGEDIN
where --code_last_page in ('generalContractSummaryTask', 'finalizationTask') 
att.dtime_valid_from  < = att.dtime_valid_to
--and att.dtime_valid_from < = att.DTIME_VALID_TO
group by vk.skp_application, vk.skp_completion_subject, vk.skp_person, code_last_page, att.SKP_EMPLOYEE_LOGGEDIN
;

--kde nedohledám potøebnou code_last_page tak vezmu datum podpisu
--0

insert into MKL_NPOS_VOLANI_MOB_3_GTT
select p.skp_application 
       , p.skp_completion_subject
       , p.skp_person
       , trunc(p.dtime_signature) as datum_dokonceni_klient
       , p.skp_employee
from MKL_NPOS_VOLANI_MOB_2_GTT p
join owner_ext.ext_pos_employee e on e.skp_employee = p.skp_employee
where not exists (select * from  MKL_NPOS_VOLANI_MOB_3_GTT z where z.skp_completion_subject = p.skp_completion_subject)
;

---propojení kontaktu z kampanì a MOB

insert into  MKL_NPOS_VOLANI_MOB_4_GTT 
select vk.obdobi
       , vk.skp_completion_subject
       , vk.skp_person
       , vk.id_appl_mob
       , vk.skp_application
       , vk.dtime_signature
       , vk.dtime_signature_valid_from
       , vk.code_signature_channel as kanal_podpisu
       , vk.skp_employee as skp_empl_podpis
       , ee.name_employee_opposite as zamestnanec_podepsal
       , vk.skp_branch_office as skp_pobocky_podepsala
       , b.name_branch as pobocka_podepsala
       , b.name_region as region_podepsal
       , v.datum_dokonceni_klient
       , kam.skf_message
       , kam.dtime_created
       , kam.dtime_planned_call_finished
       , trunc(tt.DTIME_CONTACT_FINISHED) as datum_kontaktu--kam.datum_kontaktu zmìnìno 28.11.2019, DP-546
       , tt.SKP_EMPLOYEE_FINISHED as skp_finished_employee--kam.skp_finished_employee zmìnìno 28.11.2019, DP-546
       , ee1.name_employee_opposite as zamestnanec_volal--kam.finished_employee_lastname ||' '|| kam.finished_employee_firstname as zamestnanec_volal zmìnìno 28.11.2019, DP-546
       , b1.skp_branch as skp_branch_volajici
       , b1.name_branch as pobocka_volala
       , b1.name_region as region_volal
       , kam.queue_code
       , kam.campaign_name
       , kam.campaign_code
from MKL_NPOS_VOLANI_MOB_2_GTT vk
join MKL_NPOS_VOLANI_MOB_3_GTT v on v.skp_application = vk.skp_application
join owner_dwh.f_contact_at a on a.SKP_PERSON = vk.skp_person 
join owner_dwh.f_contact_finish_tt tt on a.SKP_CONTACT = tt.skp_contact and to_char(tt.DTIME_CONTACT_FINISHED, 'YYYY/MM') >= to_char (p_date, 'YYYY/MM')-- pøidáno 28.11.2019, DP-546
left join ap_crm.JN_NPOS_KAMP_VOLANI_2018  kam on kam.skp_person = tt.skp_person and kam.skp_finished_employee = tt.SKP_EMPLOYEE_FINISHED and trunc(tt.DTIME_CONTACT_FINISHED) = trunc(kam.dtime_planned_call_finished)--pøidáno 28.11.2019, DP-546
--join ap_crm.JN_NPOS_KAMP_VOLANI_2018  kam on kam.skp_person = vk.skp_person and v.datum_dokonceni_klient = kam.datum_kontaktu zmìnìno 28.11.2019, DP-546
left join owner_ext.ext_pos_employee ee on ee.skp_employee = vk.skp_employee 
left join owner_ext.ext_pos_branch b  on b.skp_branch = vk.skp_branch_office
join owner_ext.ext_pos_employee ee1 on ee1.skp_employee = tt.SKP_EMPLOYEE_FINISHED--kam.skp_finished_employee zmìnìno 28.11.2019, DP-546
left join owner_ext.ext_pos_branch b1 on ee1.skp_branch = b1.skp_branch 
where v.datum_dokonceni_klient = trunc(tt.DTIME_CONTACT_FINISHED)--kam.datum_kontaktu zmìnìno 28.11.2019, DP-546
and a.skp_contact_channel = 1091 -- hovor pøidáno 28.11.2019, DP-546
and a.skp_contact_process_type = 1007 --kampaò pøidáno 28.11.2019, DP-546
;

---podepsané na poboèce, již vstoupili do plánu z kanálu podpisu
insert into MKL_NPOS_VOLANI_MOB_5_GTT 
select * 
from MKL_NPOS_VOLANI_MOB_4_GTT
where kanal_podpisu = 'PO' 
      and skp_pobocky_podepsala != '-1'
      and obdobi = to_char(p_date, 'YYYY/MM')
;

--podepsali
---k podpisu došlo na poboèce, pøiøazen dle podpisu
insert into MKL_NPOS_VOLANI_MOB_6_GTT 
select tk.skp_completion_subject
       , tk.skp_person
       , tk.skp_application
       , tk.id_appl_mob
       , 'MOBILITA' as text_product_description
       , tk.obdobi
       , 'podpis' as typ
       , tk.region_podepsal
       , tk.skp_pobocky_podepsala
       , tk.skp_empl_podpis as skp_asistent_podepsal
       , tk.region_volal
       , tk.skp_branch_volajici as skp_pobocka_volala
       , tk.skp_finished_employee as skp_asistent_volal
from   MKL_NPOS_VOLANI_MOB_5_GTT  tk

union 
--obvolali
--k podpisu došlo jinde, nejsou v plánech
select  tk.skp_completion_subject
       , tk.skp_person
       , tk.skp_application
       , tk.id_appl_mob
       , 'MOBILITA' as text_product_description
       , tk.obdobi
       , 'volani' as typ
       , tk.region_podepsal
       , tk.skp_pobocky_podepsala
       , tk.skp_empl_podpis as skp_asistent_podepsal
       , tk.region_volal
       , tk.skp_branch_volajici as skp_pobocka_volala
       , tk.skp_finished_employee as skp_asistent_volal
from   MKL_NPOS_VOLANI_MOB_4_GTT    tk
;

insert into MKL_NPOS_VOLANI_MOB_7_GTT 
WITH pom AS (
--duplicity
SELECT a.skp_completion_subject
     , a.skp_pobocky_podepsala
FROM MKL_NPOS_VOLANI_MOB_6_GTT  a 
WHERE a.skp_completion_subject IN (
SELECT a.skp_completion_subject FROM MKL_NPOS_VOLANI_MOB_6_GTT  a GROUP BY a.skp_completion_subject HAVING COUNT(*)>1)
GROUP BY a.skp_completion_subject
       , a.skp_pobocky_podepsala      
)
--první pøijdou s podpisem
SELECT a.skp_completion_subject
     , a.text_product_description
     , a.skp_pobocky_podepsala
     , a.skp_asistent_podepsal
     , a.obdobi
     , a.region_volal
     , 'PODPIS' as vstup
FROM MKL_NPOS_VOLANI_MOB_6_GTT   a
JOIN pom p ON p.skp_completion_subject = a.skp_completion_subject AND p.skp_pobocky_podepsala= a.skp_pobocky_podepsala and a.typ = 'podpis'  ---23
where obdobi = to_char(p_date, 'YYYY/MM')
;

insert into MKL_NPOS_VOLANI_MOB_7_GTT  
WITH pom AS (
--duplicity
SELECT a.skp_completion_subject
       , a.skp_pobocky_podepsala
       , a.typ
FROM MKL_NPOS_VOLANI_MOB_6_GTT   a 
WHERE a.skp_completion_subject IN (
SELECT a.skp_completion_subject FROM MKL_NPOS_VOLANI_MOB_6_GTT  a GROUP BY a.skp_completion_subject  HAVING COUNT(*)>1)
GROUP BY a.skp_completion_subject
       , a.skp_pobocky_podepsala 
       , a.typ     
)
SELECT  a.skp_completion_subject
       , a.text_product_description
       , a.skp_pobocky_podepsala
       , a.skp_asistent_podepsal
       , a.obdobi
       , a.region_volal
       , 'DUPLICITA_VOLANI' as vstup
FROM MKL_NPOS_VOLANI_MOB_6_GTT   a
JOIN pom p ON p.skp_completion_subject = a.skp_completion_subject AND p.skp_pobocky_podepsala != a.skp_pobocka_volala and p.typ != 'podpis' 
where obdobi = to_char(p_date, 'YYYY/MM')
group by a.skp_completion_subject
       , a.text_product_description
       , a.skp_pobocky_podepsala
       , a.skp_asistent_podepsal
       , a.obdobi
       , a.region_volal
       , 'DUPLICITA_VOLANI' 
;

insert into MKL_NPOS_VOLANI_MOB_7_GTT 
WITH pom AS (
--duplicity
SELECT a.skp_completion_subject
       , a.skp_pobocky_podepsala
FROM MKL_NPOS_VOLANI_MOB_6_GTT  a 
WHERE a.skp_completion_subject IN (
SELECT a.skp_completion_subject FROM MKL_NPOS_VOLANI_MOB_6_GTT a GROUP BY a.skp_completion_subject  HAVING COUNT(*)>1)
GROUP BY a.skp_completion_subject
       , a.skp_pobocky_podepsala      
)
SELECT  a.skp_completion_subject
       , a.text_product_description
       , a.skp_pobocky_podepsala
       , a.skp_asistent_podepsal
       , a.obdobi
       , a.region_volal
       , 'VOLANI' as vstup
FROM MKL_NPOS_VOLANI_MOB_6_GTT a 
where not exists (select * from pom p where a.skp_completion_subject = p.skp_completion_subject and p.skp_pobocky_podepsala = a.skp_pobocky_podepsala)
and obdobi = to_char(p_date, 'YYYY/MM') 
;

---ještì dohledám duplicity s podpisem
insert into MKL_NPOS_VOLANI_MOB_8_GTT 
select skp_completion_subject from MKL_NPOS_VOLANI_MOB_7_GTT   having count(skp_completion_subject)> 1 group by skp_completion_subject
;


--upravím podpis na duplicita podpis, shoduje se s duplicita_volani
--upravuji název podpis na duplicita_podpis
UPDATE 
      (select p.vstup 
        from MKL_NPOS_VOLANI_MOB_7_GTT   p
        where p.skp_completion_subject in (select u.skp_completion_subject from  MKL_NPOS_VOLANI_MOB_8_GTT u)
        and p.vstup = 'PODPIS')
SET VSTUP = 'DUPLICITA_PODPIS'
WHERE VSTUP = 'PODPIS'
;

insert into MKL_NPOS_VOLANI_MOB_9_GTT 
select 'MOBILITA' as produkt
       , u.obdobi
       , b.name_region as region_volal
       , b.type_branch_office as typ_pobocky_volala
       , b.skp_branch as skp_pobocky_volala
       , b.name_branch as  nazev_pobocky_volala
       , e.skp_employee as skp_asistent_volal
       , e.name_employee_opposite as asistent_volal
       , p.campaign_name as nazev_kampanì
       , pp.dtime_signature
       , u.vstup
     --  , pp.code_signature_channel
       , case when u.vstup in ('PODPIS', 'DUPLICITA_PODPIS', 'DUPLICITA_VOLANI') then 'POS'
                   else 'jinde' end as kanal_podpisu
       , case when u.vstup in ('PODPIS', 'DUPLICITA_PODPIS', 'DUPLICITA_VOLANI') then b1.name_branch
                   end as pobocka_podpisu
       , case when u.vstup in ('PODPIS','DUPLICITA_PODPIS','DUPLICITA_VOLANI') then b1.skp_branch
                   end as skp_pobocky_podpisu
       , case when u.vstup in ('PODPIS', 'DUPLICITA_PODPIS','DUPLICITA_VOLANI')then e1.name_employee_opposite
              end as podepsal_asistent
       , case when u.vstup in  ('PODPIS', 'DUPLICITA_PODPIS','DUPLICITA_VOLANI') then e1.skp_employee
               end as skp_asist_podepsal
       , p.kanal_podpisu as kod_kanalu_podpisu
       , p.datum_dokonceni_klient
       , p.datum_kontaktu
       , p.skp_completion_subject
       , p.skp_person
       , p.id_appl_mob
from MKL_NPOS_VOLANI_MOB_7_GTT u        
left join MKL_NPOS_VOLANI_MOB_4_GTT p on u.skp_completion_subject = p.skp_completion_subject ---použiji jen pro zijištìní názvu kampanì  a údaje ze které poboèky se volalo
left join MKL_NPOS_VOLANI_MOB_2_GTT pp on pp.skp_completion_subject = u.skp_completion_subject --kde se podepsalo
left join owner_ext.ext_pos_branch b  on b.skp_branch = p.skp_branch_volajici
left join owner_ext.ext_pos_employee e on e.skp_employee = p.skp_finished_employee
left join owner_ext.ext_pos_employee e1 on e1.skp_employee = pp.skp_employee
left join owner_ext.ext_pos_branch b1 on b1.skp_branch = pp.skp_branch_office
;


insert into MKL_NPOS_VOLANI_MOB_10_GTT
select  v.produkt as produkt_volani
       , v.obdobi
       , v.region_volal 
       , v.typ_pobocky_volala
       , v.nazev_pobocky_volala
       , v.asistent_volal
       , v.nazev_kampanì
       , v.dtime_signature as datum_podpisu
       , v.datum_dokonceni_klient
       , v.datum_kontaktu
       , v.id_appl_mob
       , v.vstup
       , v.kod_kanalu_podpisu as kanal_pospisu
       , v.pobocka_podpisu
       , v.podepsal_asistent
       , v.skp_person
       , case when vstup in ('PODPIS', 'DUPLICITA_PODPIS') then v.pobocka_podpisu
              when vstup in ('VOLANI', 'DUPLICITA_VOLANI') then v.nazev_pobocky_volala end as pocita_se_do_planu
from MKL_NPOS_VOLANI_MOB_9_GTT v
group by  v.produkt 
       , v.obdobi
       , v.region_volal 
       , v.typ_pobocky_volala
       , v.nazev_pobocky_volala
       , v.asistent_volal
       , v.nazev_kampanì
       , v.dtime_signature 
       , v.datum_dokonceni_klient
       , v.datum_kontaktu
       , v.id_appl_mob
       , v.vstup
       , v.kod_kanalu_podpisu 
       , v.pobocka_podpisu
       , v.podepsal_asistent
       , v.skp_person
       , case when vstup in ('PODPIS', 'DUPLICITA_PODPIS') then v.pobocka_podpisu
              when vstup in ('VOLANI', 'DUPLICITA_VOLANI') then v.nazev_pobocky_volala end
;


delete from MKL_NPOS_VOLANI_MOB_FINAL WHERE obdobi =  to_char(p_date, 'YYYY/MM') -- vymazu aktualni datum
;

insert into MKL_NPOS_VOLANI_MOB_FINAL
select a.*
       , d.id_cuid 
from MKL_NPOS_VOLANI_MOB_10_GTT a
join owner_dwh.dc_person d on d.skp_person = a.skp_person  
WHERE NOT EXISTS (SELECT b.obdobi from MKL_NPOS_VOLANI_MOB_FINAL b WHERE a.obdobi = b.obdobi) -- pojistka, aby se neduplikovaly zaznamy
;




----------------------------------------
--VOLÁNÍ DATUM AKTUALIZACE REPORTU
----------------------------------------

--vysledna tabule k datu nápoètu
delete from  datum_reporty_final where object_name = 'JN_NPOS_VOLANI_PRODUKTY';
       
---- datum aktualizace

insert into DATUM_REPORTY_FINAL 
select user, 'JN_NPOS_VOLANI_PRODUKTY' as PROCEDURA, sysdate as datum_aktualizace from dual
;

commit;

end JN_NPOS_VOLANI_PRODUKTY;
/
