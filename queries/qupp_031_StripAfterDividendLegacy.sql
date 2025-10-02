-- qupp_031_StripAfterDividendLegacy
-- Exported from Access on 2025-10-01 18:29:16

UPDATE tblTransactionsModified SET tblTransactionsModified.ModifiedDetails = StripAfterDividendLegacy ([Details])
WHERE ((([tblTransactionsModified].Details) Like "*Dividend on*"));

