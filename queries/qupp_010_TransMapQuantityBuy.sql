-- qupp_010_TransMapQuantityBuy
-- Exported from Access on 2025-10-01 18:29:16

UPDATE tblTransactionsModified SET Quantity = ExtractQuantityFromBuy([Details])
WHERE [Details] LIKE '*Bought*';

