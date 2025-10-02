-- qapp_TransactionsModifiedRogueTrades
-- Exported from Access on 2025-10-01 18:29:16

INSERT INTO tblTransactionsModified ( [Date], Details, [Paid In], Withdrawn, Quantity, Price )
SELECT tblMapRogueTrades.Date, tblMapRogueTrades.Details, tblMapRogueTrades.[Paid In], tblMapRogueTrades.Withdrawn, tblMapRogueTrades.Quantity, tblMapRogueTrades.Price
FROM tblMapRogueTrades;

