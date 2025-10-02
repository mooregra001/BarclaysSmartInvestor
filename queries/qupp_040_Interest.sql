-- qupp_040_Interest
-- Exported from Access on 2025-10-01 18:29:17

UPDATE tblTransactionsModified SET tblTransactionsModified.ModifiedDetails = "Interest"
WHERE (
        (
            ([tblTransactionsModified].ModifiedDetails) IS NULL
        )
        AND (
            ([tblTransactionsModified].Details) LIKE "*Interest*"
        )
    );

