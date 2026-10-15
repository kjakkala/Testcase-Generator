# Internal QA Dashboard Database Deployment Package

This package creates the database schema for an internal QA dashboard that supports:

- user authentication and roles
- per-user integration credentials for TestRail, Jira, GitHub, and AI providers
- project access control
- generated test plan storage
- audit logging

## Files

- deploy.sql - full deployment script
- 01_create_database.sql - creates the database
- 02_create_tables.sql - creates the schema and tables

## Prerequisites

- Microsoft SQL Server 2019 or newer
- SQL login with permission to create a database and tables
- A target database name (default: QA_TEST_PLAN)

## Deploy

Run the script in SQL Server Management Studio (SSMS) or Azure Data Studio:

1. Open SSMS
2. Connect to the SQL Server instance
3. Open deploy.sql
4. Update the database name if needed
5. Execute

Example:

```sql
USE master;
GO

CREATE DATABASE QA_TEST_PLAN;
GO
```

## Table overview

### Users
Stores application users and roles.

### UserIntegrations
Stores each user's API credentials for external systems, such as TestRail, Jira, GitHub, or AI providers.

### UserProjects
Controls which projects a user can access.

### GeneratedPlans
Stores generated test plan drafts or final exports.

### AuditLogs
Tracks actions for security and troubleshooting.

## Security model

This package is designed for the following security pattern:

- the web app authenticates the user
- the server loads only that user's credentials from SQL Server
- API calls to TestRail, Jira, GitHub, and AI providers happen server-side
- browser clients do not receive raw tokens
- credentials are stored encrypted or service-side protected

## Recommended production security

- Use HTTPS between users and the web app
- Use TLS between web server and SQL Server
- Use a least-privilege SQL login
- Store secrets encrypted at rest or via SQL Server encryption features
- Restrict DB access to the web server only
- Use Windows authentication or a dedicated service account when possible
- Keep the web app on an internal private network or VPN

## Suggested connection pattern

Web server -> SQL Server: private network / VPN / TLS

Browser -> Web app: HTTPS

## Notes

This package provides the database layer. The web app still needs to authenticate users and read only the active user's credentials from the database at runtime.
