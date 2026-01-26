This folder contains SQL scripts to create the local SQL Server Express database
for the SQL Svr VBA API project.

> **Note:** Update `DB_CONN_STRING` if your SQL Server instance name is different (for example, not `.\SQLEXPRESS`).
> 
If your SQL Server instance name is different from `.\SQLEXPRESS`, update the `DB_CONN_STRING` constant in the VBA config module accordingly.

## Database Setup

1. Run the scripts in `/database` in this order:
   - 01_create_database.sql  
   - 02_create_tables.sql  
   - 03_create_procedures.sql  
   - 04_seed_sample_data.sql  

2. Configure the connection string in VBA:

```vb
Public Const DB_CONN_STRING As String = _
    "Provider=MSOLEDBSQL;" & _
    "Server=.\SQLEXPRESS;" & _
    "Database=MGA_RiskLab;" & _
    "Trusted_Connection=Yes;"

