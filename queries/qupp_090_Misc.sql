-- qupp_090_Misc
-- Exported from Access on 2025-10-01 18:29:17

UPDATE tblTransactionsModified INNER JOIN tblMapMiscTransactions ON InStr(tblTransactionsModified.Details, tblMapMiscTransactions.Details) > 0 SET tblTransactionsModified.ModifiedDetails = tblMapMiscTransactions.ModifiedDetails
WHERE tblTransactionsModified.ModifiedDetails IS NULL;

