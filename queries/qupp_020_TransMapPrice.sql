-- qupp_020_TransMapPrice
-- Exported from Access on 2025-10-01 18:29:16

UPDATE tblTransactionsModified SET tblTransactionsModified.Price = IIf([tblTransactionsModified]!Quantity>0,-[tblTransactionsModified]!Withdrawn/[tblTransactionsModified]!Quantity,[tblTransactionsModified]![Paid In]/-[tblTransactionsModified]!Quantity);

