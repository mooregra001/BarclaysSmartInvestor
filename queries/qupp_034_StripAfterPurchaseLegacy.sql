-- qupp_034_StripAfterPurchaseLegacy
-- Exported from Access on 2025-10-01 18:29:16

UPDATE tblTransactionsModified SET tblTransactionsModified.ModifiedDetails = StripAfterPurchaseLegacy ([Details])
WHERE ((([tblTransactionsModified].Details) Like "*Purchase of*"));

