-- qryRecInvestments
-- Exported from Access on 2025-10-01 18:29:16

SELECT tblTransactionsModified.ConsolidatedDetails, tblInvestmentsConsolidated.ConsolidatedInvestments, Sum(tblTransactionsModified.Quantity) AS SumOfQuantity, Sum(tblTransactionsModified.[Paid In]) AS [SumOfPaid In], Sum(tblTransactionsModified.Withdrawn) AS SumOfWithdrawn, tblInvestmentsConsolidated.[Value (£)], Sum(Nz([tblTransactionsModified]![Paid In])+Nz([tblTransactionsModified]![Withdrawn]))+[tblInvestmentsConsolidated].[Value (£)] AS PnL, tblInvestmentsConsolidated.[Quantity Held], [SumOfQuantity]-[Quantity Held] AS [Qty Diff]
FROM tblInvestmentsConsolidated LEFT JOIN tblTransactionsModified ON tblInvestmentsConsolidated.[ConsolidatedInvestments] = tblTransactionsModified.ConsolidatedDetails
GROUP BY tblTransactionsModified.ConsolidatedDetails, tblInvestmentsConsolidated.ConsolidatedInvestments, tblInvestmentsConsolidated.[Value (£)], tblInvestmentsConsolidated.[Quantity Held]
HAVING (((Sum(tblTransactionsModified.Quantity))<>0));

