This folder contains SQL scripts to create the local SQL Server Express database
for the SQL Svr VBA API project.

> **Note:** Update `DB_CONN_STRING` if your SQL Server instance name is different (for example, not `.\SQLEXPRESS`).
> 
If your SQL Server instance name is different from `.\SQLEXPRESS`, update the `DB_CONN_STRING` constant in the VBA config module accordingly.

# Database Setup (SQL Server Express)

This project uses SQL Server Express locally.

## Prerequisites
- SQL Server Express installed
- SQL Server Management Studio (SSMS)

## Setup Steps

1. Open SSMS and connect to your local instance:
   - Server name: .\SQLEXPRESS

2. Run the scripts in this order:

   1. 01_create_database.sql  
   2. 02_create_tables.sql  
   3. 03_create_procedures.sql  
   4. 04_seed_sample_data.sql  

3. Verify the database:
   - Database name: MGA_RiskLab
   - Tables created under dbo schema

## Connection String (VBA)

Use Windows Authentication:

Provider=MSOLEDBSQL;
Server=.\SQLEXPRESS;
Database=MGA_RiskLab;
Trusted_Connection=Yes;
This is gold for reviewers.
________________________________________
Step 4) Put the connection string in ONE config place
In your VBA project:
Create a module:
Public Const DB_CONN_STRING As String = _
    "Provider=MSOLEDBSQL;" & _
    "Server=.\SQLEXPRESS;" & _
    "Database=MGA_RiskLab;" & _
    "Trusted_Connection=Yes;"
Then all code uses this constant.

# Note:  Update DB_CONN_STRING if your SQL instance name is different.
________________________________________
Step 5) What NOT to put in Git
Do NOT commit:
•	.mdf or .ldf files
•	Actual production data
•	Credentials
•	Machine-specific paths
Add to .gitignore:
*.mdf
*.ldf
*.bak
________________________________________
Option #2 - One-click Database setup script

Run setup_all.sql to create the full database.

Which will execute these scripts:
:r 01_create_database.sql
:r 02_create_tables.sql
:r 03_create_procedures.sql
:r 04_seed_sample_data.sql



