-- qupp_032_StripAfterQuantity
-- Exported from Access on 2025-10-01 18:29:16

UPDATE tblTransactionsModified SET tblTransactionsModified.ModifiedDetails = StripAfterQuantity ([Details])
WHERE ((([tblTransactionsModified].Details) Like "*Bought*")) Or ((([tblTransactionsModified].Details) Like "*Sold*"));

