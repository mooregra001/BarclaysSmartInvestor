-- qupp_033_StripBuySell
-- Exported from Access on 2025-10-01 18:29:16

UPDATE tblTransactionsModified SET tblTransactionsModified.ModifiedDetails = StripTransactionFee ([Details])
WHERE ((([tblTransactionsModified].Details) Like "*transaction fee*"));

