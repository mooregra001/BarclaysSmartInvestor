-- qapp_InvestmentsConsolidated
-- Exported from Access on 2025-10-01 18:29:16

INSERT INTO tblInvestmentsConsolidated ( ConsolidatedInvestments, Investment, Identifier, [Quantity Held], [Last Price], [Last Price CCY], [Value], [Value CCY], [FX Rate], [Last Price (p)], [Value (£)], [Book Cost], [Book Cost CCY], [Average FX Rate], [Book Cost (£)], [% Change] )
SELECT tblMapConsolidatedDetails.ConsolidatedDetails, tblInvestments.Investment, tblInvestments.Identifier, tblInvestments.[Quantity Held], tblInvestments.[Last Price], tblInvestments.[Last Price CCY], tblInvestments.Value, tblInvestments.[Value CCY], tblInvestments.[FX Rate], tblInvestments.[Last Price (p)], tblInvestments.[Value (£)], tblInvestments.[Book Cost], tblInvestments.[Book Cost CCY], tblInvestments.[Average FX Rate], tblInvestments.[Book Cost (£)], tblInvestments.[% Change]
FROM tblInvestments LEFT JOIN tblMapConsolidatedDetails ON tblInvestments.Investment = tblMapConsolidatedDetails.ModifiedDetails;

