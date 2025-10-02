-- qryRec
-- Exported from Access on 2025-10-01 18:29:16

SELECT tblInvestmentsConsolidated.ConsolidatedInvestments, tblTransactionsModified.ConsolidatedDetails, Sum(tblTransactionsModified.Quantity) AS SumOfQuantity, Sum(tblTransactionsModified.[Paid In]) AS [SumOfPaid In], Sum(tblTransactionsModified.Withdrawn) AS SumOfWithdrawn, tblInvestmentsConsolidated.[Value (£)], Sum(Nz([tblTransactionsModified]![Paid In])+Nz([tblTransactionsModified]![Withdrawn]))+[tblInvestmentsConsolidated]![Value (£)] AS PnL, tblInvestmentsConsolidated.[Quantity Held], [Quantity Held]-[SumOfQuantity] AS PositionDifference
FROM tblInvestmentsConsolidated RIGHT JOIN tblTransactionsModified ON tblInvestmentsConsolidated.[ConsolidatedInvestments] = tblTransactionsModified.ConsolidatedDetails
GROUP BY tblInvestmentsConsolidated.ConsolidatedInvestments, tblTransactionsModified.ConsolidatedDetails, tblInvestmentsConsolidated.[Value (£)], tblInvestmentsConsolidated.[Quantity Held]
HAVING (((Sum(tblTransactionsModified.Quantity))<>0));

