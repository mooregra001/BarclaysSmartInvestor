-- qupp_050_AdminFee
-- Exported from Access on 2025-10-01 18:29:17

UPDATE tblTransactionsModified SET tblTransactionsModified.ModifiedDetails = "AdminFee"
WHERE ((([tblTransactionsModified].ModifiedDetails) Is Null) And (([tblTransactionsModified].Details) Like "*fee on investment holdings*" Or ([tblTransactionsModified].Details) Like "*Admin Fee*" Or ([tblTransactionsModified].Details) Like "*AdminFee*" Or ([tblTransactionsModified].Details) Like "*fee payment for SIPP Account*" Or ([tblTransactionsModified].Details) Like "*fee payment for Third Party SIPP Account*")) Or ((([tblTransactionsModified].Details) Like "*SMT 133262 transfer*")) Or ((([tblTransactionsModified].Details) Like "*Administration fee*")) Or ((([tblTransactionsModified].Details) Like "*Outstanding administration fee*"));

