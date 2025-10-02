-- qmak_ImportLog
-- Exported from Access on 2025-10-01 18:29:16

CREATE TABLE tblImportLog (
    LogID AUTOINCREMENT PRIMARY KEY,
    TableName TEXT(255),
    Status TEXT(50),
    Details LONGTEXT,
    ImportDate DATETIME
);
