-- qryPnL
-- Exported from Access on 2025-10-01 18:29:16

SELECT tblTransactionsModified.ConsolidatedDetails, Sum(tblTransactionsModified.Quantity) AS SumOfQuantity, tblInvestmentsConsolidated.[Quantity Held], [SumOfQuantity]-[Quantity Held] AS [Qty Diff], Sum(tblTransactionsModified.[Paid In]) AS [SumOfPaid In], Sum(tblTransactionsModified.Withdrawn) AS SumOfWithdrawn, tblInvestmentsConsolidated.[Value (£)], Sum(Nz([tblTransactionsModified]![Paid In])+Nz([tblTransactionsModified]![Withdrawn]))+Nz([tblInvestmentsConsolidated]![Value (£)]) AS PnL
FROM tblTransactionsModified LEFT JOIN tblInvestmentsConsolidated ON tblTransactionsModified.ConsolidatedDetails = tblInvestmentsConsolidated.[ConsolidatedInvestments]
GROUP BY tblTransactionsModified.ConsolidatedDetails, tblInvestmentsConsolidated.[Quantity Held], tblInvestmentsConsolidated.[Value (£)];

