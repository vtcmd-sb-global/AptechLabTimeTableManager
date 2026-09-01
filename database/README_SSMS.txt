monthly_bsr database – SSMS / SQL Server notes
==============================================

Application database file (SQLite):
  database/monthly_bsr.db

Convenience binary copy (same content as the .db):
  database/monthly_bsr.bak

SQL Server import script (run inside SSMS):
  database/monthly_bsr.sql

Why a native SQL Server .bak cannot be produced here
----------------------------------------------------
SQL Server .bak files are a proprietary binary format that can only be
produced by Microsoft SQL Server itself (BACKUP DATABASE ... TO DISK).
This application uses SQLite for its local store; a true SQL Server
backup file cannot be generated from the Python/SQLite environment.

How to open the data in SSMS
----------------------------
1. In SSMS create an empty database, e.g. monthly_bsr.
2. Open monthly_bsr.sql and execute it against that database.
3. All tables, columns and current rows will be created.

Alternatively you can inspect monthly_bsr.db (or the .bak copy) with
DB Browser for SQLite, DBeaver, or any SQLite-compatible tool.
