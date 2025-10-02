-- qupp_031_StripAfterDividend
-- Exported from Access on 2025-10-01 18:29:16

UPDATE tblTransactionsModified SET tblTransactionsModified.ModifiedDetails = StripAfterDividend ([Details])
WHERE ((([tblTransactionsModified].Details) Like "*Dividend:*"));

