USE master;
GO

IF DB_ID(N'QA_TEST_PLAN') IS NULL
BEGIN
    CREATE DATABASE QA_TEST_PLAN;
END;
GO

USE QA_TEST_PLAN;
GO

IF OBJECT_ID(N'dbo.Users', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Users (
        UserId UNIQUEIDENTIFIER NOT NULL CONSTRAINT PK_Users PRIMARY KEY DEFAULT NEWID(),
        Email NVARCHAR(255) NOT NULL,
        DisplayName NVARCHAR(200) NULL,
        PasswordHash NVARCHAR(255) NULL,
        RoleName NVARCHAR(50) NOT NULL CONSTRAINT DF_Users_Role DEFAULT 'QA',
        IsActive BIT NOT NULL CONSTRAINT DF_Users_IsActive DEFAULT 1,
        CreatedUtc DATETIME2 NOT NULL CONSTRAINT DF_Users_CreatedUtc DEFAULT SYSUTCDATETIME(),
        LastLoginUtc DATETIME2 NULL,
        CONSTRAINT UQ_Users_Email UNIQUE (Email)
    );
END;
GO

IF OBJECT_ID(N'dbo.UserIntegrations', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.UserIntegrations (
        IntegrationId UNIQUEIDENTIFIER NOT NULL CONSTRAINT PK_UserIntegrations PRIMARY KEY DEFAULT NEWID(),
        UserId UNIQUEIDENTIFIER NOT NULL,
        Provider NVARCHAR(30) NOT NULL,
        BaseUrl NVARCHAR(500) NULL,
        Username NVARCHAR(255) NULL,
        SecretValue NVARCHAR(500) NULL,
        SecretType NVARCHAR(50) NULL,
        IsEnabled BIT NOT NULL CONSTRAINT DF_UserIntegrations_IsEnabled DEFAULT 1,
        CreatedUtc DATETIME2 NOT NULL CONSTRAINT DF_UserIntegrations_CreatedUtc DEFAULT SYSUTCDATETIME(),
        UpdatedUtc DATETIME2 NOT NULL CONSTRAINT DF_UserIntegrations_UpdatedUtc DEFAULT SYSUTCDATETIME(),
        CONSTRAINT FK_UserIntegrations_Users FOREIGN KEY (UserId) REFERENCES dbo.Users (UserId) ON DELETE CASCADE,
        CONSTRAINT CK_UserIntegrations_Provider CHECK (Provider IN ('TestRail','Jira','GitHub','AI','OpenAI'))
    );
END;
GO

IF OBJECT_ID(N'dbo.UserProjects', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.UserProjects (
        UserProjectId UNIQUEIDENTIFIER NOT NULL CONSTRAINT PK_UserProjects PRIMARY KEY DEFAULT NEWID(),
        UserId UNIQUEIDENTIFIER NOT NULL,
        ProjectKey NVARCHAR(100) NOT NULL,
        ProjectName NVARCHAR(200) NULL,
        IsAllowed BIT NOT NULL CONSTRAINT DF_UserProjects_IsAllowed DEFAULT 1,
        CreatedUtc DATETIME2 NOT NULL CONSTRAINT DF_UserProjects_CreatedUtc DEFAULT SYSUTCDATETIME(),
        CONSTRAINT FK_UserProjects_Users FOREIGN KEY (UserId) REFERENCES dbo.Users (UserId) ON DELETE CASCADE
    );
END;
GO

IF OBJECT_ID(N'dbo.GeneratedPlans', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.GeneratedPlans (
        PlanId UNIQUEIDENTIFIER NOT NULL CONSTRAINT PK_GeneratedPlans PRIMARY KEY DEFAULT NEWID(),
        UserId UNIQUEIDENTIFIER NOT NULL,
        Title NVARCHAR(250) NOT NULL,
        Status NVARCHAR(50) NOT NULL CONSTRAINT DF_GeneratedPlans_Status DEFAULT 'Draft',
        PlanJson NVARCHAR(MAX) NULL,
        CreatedUtc DATETIME2 NOT NULL CONSTRAINT DF_GeneratedPlans_CreatedUtc DEFAULT SYSUTCDATETIME(),
        UpdatedUtc DATETIME2 NOT NULL CONSTRAINT DF_GeneratedPlans_UpdatedUtc DEFAULT SYSUTCDATETIME(),
        CONSTRAINT FK_GeneratedPlans_Users FOREIGN KEY (UserId) REFERENCES dbo.Users (UserId) ON DELETE CASCADE
    );
END;
GO

IF OBJECT_ID(N'dbo.AuditLogs', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.AuditLogs (
        AuditLogId UNIQUEIDENTIFIER NOT NULL CONSTRAINT PK_AuditLogs PRIMARY KEY DEFAULT NEWID(),
        UserId UNIQUEIDENTIFIER NULL,
        ActionName NVARCHAR(100) NOT NULL,
        EntityType NVARCHAR(100) NULL,
        EntityId NVARCHAR(200) NULL,
        LogDetails NVARCHAR(MAX) NULL,
        CreatedUtc DATETIME2 NOT NULL CONSTRAINT DF_AuditLogs_CreatedUtc DEFAULT SYSUTCDATETIME(),
        CONSTRAINT FK_AuditLogs_Users FOREIGN KEY (UserId) REFERENCES dbo.Users (UserId)
    );
END;
GO

CREATE UNIQUE INDEX IX_Users_Email ON dbo.Users(Email);
CREATE INDEX IX_UserIntegrations_UserId ON dbo.UserIntegrations(UserId);
CREATE INDEX IX_UserProjects_UserId ON dbo.UserProjects(UserId);
CREATE INDEX IX_GeneratedPlans_UserId ON dbo.GeneratedPlans(UserId);
CREATE INDEX IX_AuditLogs_UserId ON dbo.AuditLogs(UserId);
GO

PRINT 'Database schema deployed successfully.';
GO
