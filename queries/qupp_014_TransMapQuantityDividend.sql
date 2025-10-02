-- qupp_014_TransMapQuantityDividend
-- Exported from Access on 2025-10-01 18:29:16

UPDATE tblTransactionsModified SET tblTransactionsModified.Quantity = ExtractQuantityDividendReinvestment ([Details])
WHERE [Details] LIKE '*Automatic dividend reinvest - purchase*';

