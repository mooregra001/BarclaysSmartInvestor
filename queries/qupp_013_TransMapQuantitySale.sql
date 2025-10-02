-- qupp_013_TransMapQuantitySale
-- Exported from Access on 2025-10-01 18:29:16

UPDATE tblTransactionsModified SET tblTransactionsModified.Quantity = PopulateQuantityFromSale ([Details])
WHERE [Details] LIKE '*Sale of *';

