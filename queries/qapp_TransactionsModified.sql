-- qapp_TransactionsModified
-- Exported from Access on 2025-10-01 18:29:16

INSERT INTO tblTransactionsModified ( [Date], Account, Details, [Paid In], Withdrawn )
SELECT tblTransactions.Date, tblTransactions.Account, tblTransactions.Details, tblTransactions.[Paid In], tblTransactions.Withdrawn
FROM tblTransactions;

