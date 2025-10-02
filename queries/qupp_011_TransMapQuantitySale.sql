-- qupp_011_TransMapQuantitySale
-- Exported from Access on 2025-10-01 18:29:16

UPDATE tblTransactionsModified SET Quantity = ExtractQuantityFromSell([Details])
WHERE [Details] LIKE '*Sold*';

