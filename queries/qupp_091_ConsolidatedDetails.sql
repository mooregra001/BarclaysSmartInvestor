-- qupp_091_ConsolidatedDetails
-- Exported from Access on 2025-10-01 18:29:17

UPDATE tblTransactionsModified INNER JOIN tblMapConsolidatedDetails ON tblTransactionsModified.ModifiedDetails = tblMapConsolidatedDetails.ModifiedDetails SET tblTransactionsModified.ConsolidatedDetails = tblMapConsolidatedDetails.ConsolidatedDetails
WHERE (((tblTransactionsModified.ConsolidatedDetails) Is Null));

