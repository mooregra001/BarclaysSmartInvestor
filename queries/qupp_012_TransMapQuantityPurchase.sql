-- qupp_012_TransMapQuantityPurchase
-- Exported from Access on 2025-10-01 18:29:16

UPDATE tblTransactionsModified SET tblTransactionsModified.Quantity = PopulateQuantityFromPurchase ([Details])
WHERE [Details] LIKE '*Purchase of *';

