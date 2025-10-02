-- qupp_035_StripAfterSaleLegacy
-- Exported from Access on 2025-10-01 18:29:17

UPDATE tblTransactionsModified SET tblTransactionsModified.ModifiedDetails = StripAfterSaleLegacy ([Details])
WHERE ((([tblTransactionsModified].Details) Like "*Sale of*"));

