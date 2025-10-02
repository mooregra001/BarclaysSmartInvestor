-- qupp_036_StripAfterAutomaticLegacy
-- Exported from Access on 2025-10-01 18:29:17

UPDATE tblTransactionsModified SET tblTransactionsModified.ModifiedDetails = StripAfterAutomaticLegacy ([Details])
WHERE ((([tblTransactionsModified].Details) Like "*Automatic dividend reinvest - purchase *"));

