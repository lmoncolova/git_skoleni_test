-- Mail od Martina 16.7.2018
/*
Ahoj,

Vašek teï dìlal do reportu k hypotékám stavy na chytré rezervì.
Možná se ten kus kódu bude hodit do splátkáèù na HYPO v JEZ.
*/


select l.SKP_APPLICATION, td.AMT_BALANCE as amt_otb_3m_after_util
from owner_dwh.f_loan_at l
join owner_dwh.f_account_loan_rel_at rel on rel.SKP_LOAN = l.SKP_LOAN
join owner_dwh.f_Account_Balance_Td td on td.SKP_BANK_ACCOUNT = rel.SKP_BANK_ACCOUNT
                                      and add_months(l.DATE_LOAN_UTILIZATION,3) between td.DATE_BALANCE_VALID_FROM and td.DATE_BALANCE_VALID_TO
                                      and td.DATE_BALANCE_VALID_FROM >= date'2015-3-13' -- kvùli partitionu, ale moc nefunguje. Bìží pùl minuty
join owner_dwh.cl_bank_account_type cl on cl.SKP_BANK_ACCOUNT_TYPE = td.SKP_BANK_ACCOUNT_TYPE
                                      and cl.NAME_BANK_ACCOUNT_TYPE = 'Úèet pøedplatku (OTB)'
where l.CODE_LOAN_NUMBER like 'HR%'
  and l.DATE_INSTALMENT_FIRST <= l.DATE_INSTALMENT_LAST
