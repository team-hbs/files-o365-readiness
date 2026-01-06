$databaseName = GetConfig('DatabaseName')

# -------------------------------------------------
# Batch
# -------------------------------------------------
$query = "
    USE [$databaseName]

    IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Batch' AND schema_id = SCHEMA_ID('dbo'))
    BEGIN
        CREATE TABLE [dbo].[Batch](
            [Id] [int] IDENTITY(1,1) NOT NULL,
            [RunDate] [datetime] NULL,
            [Server] [nvarchar](max) NULL,
            [BatchNumber] [int] NULL,
            [Status] [int] NULL,
            [Wave] [nvarchar](max) NULL,
            [CutoffDate] [datetime] NULL
        ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
    END
"
SqlQueryInsert($query)

# -------------------------------------------------
# Event
# -------------------------------------------------
$query = "
    USE [$databaseName]

    IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Event' AND schema_id = SCHEMA_ID('dbo'))
    BEGIN
        CREATE TABLE [dbo].[Event](
            [Id] [int] IDENTITY(1,1) NOT NULL,
            [OwnerId] [int] NULL,
            [EventType] [nvarchar](max) NULL,
            [EventDate] [datetime] NULL
        ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
    END
"
SqlQueryInsert($query)

# -------------------------------------------------
# GlobalConfig
# -------------------------------------------------
$query = "
    USE [$databaseName]

    IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'GlobalConfig' AND schema_id = SCHEMA_ID('dbo'))
    BEGIN
        CREATE TABLE [dbo].[GlobalConfig](
            [Id] [int] IDENTITY(1,1) NOT NULL,
            [Key] [nvarchar](max) NULL,
            [Value] [nvarchar](max) NULL,
            [Server] [nvarchar](max) NULL
        ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
    END
"
SqlQueryInsert($query)

# -------------------------------------------------
# MigrationQueue
# -------------------------------------------------
$query = "
    USE [$databaseName]

    IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'MigrationQueue' AND schema_id = SCHEMA_ID('dbo'))
    BEGIN
        CREATE TABLE [dbo].[MigrationQueue](
            [OwnerId] [int] NOT NULL,
            [Created] [datetime] NULL,
            [Server] [nvarchar](max) NULL,
            [BatchNumber] [int] NULL,
            CONSTRAINT [PK_MigrationQueue] PRIMARY KEY CLUSTERED ([OwnerId] ASC)
        ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
    END
"
SqlQueryInsert($query)

# -------------------------------------------------
# ScanFile
# -------------------------------------------------
$query = "
    USE [$databaseName]

    IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ScanFile' AND schema_id = SCHEMA_ID('dbo'))
    BEGIN
        CREATE TABLE [dbo].[ScanFile](
            [Id] [int] IDENTITY(1,1) NOT NULL,
            [FileName] [nvarchar](max) NULL,
            [Location] [nvarchar](max) NULL,
            [Created] [datetime] NULL,
            [Modified] [datetime] NULL,
            [Author] [nvarchar](max) NULL,
            [Extension] [nvarchar](max) NULL,
            [Folder01] [nvarchar](max) NULL,
            [Folder02] [nvarchar](max) NULL,
            [Folder03] [nvarchar](max) NULL,
            [Folder04] [nvarchar](max) NULL,
            [Folder05] [nvarchar](max) NULL,
            [Folder06] [nvarchar](max) NULL,
            [Folder07] [nvarchar](max) NULL,
            [Folder08] [nvarchar](max) NULL,
            [Folder09] [nvarchar](max) NULL,
            [Folder10] [nvarchar](max) NULL,
            [Folder11] [nvarchar](max) NULL,
            [Folder12] [nvarchar](max) NULL,
            [Folder13] [nvarchar](max) NULL,
            [Folder14] [nvarchar](max) NULL,
            [Folder15] [nvarchar](max) NULL,
            [Folder16] [nvarchar](max) NULL,
            [Folder17] [nvarchar](max) NULL,
            [Folder18] [nvarchar](max) NULL,
            [Folder19] [nvarchar](max) NULL,
            [Folder20] [nvarchar](max) NULL,
            [Normalized] [bit] NULL,
            [OwnerId] [int] NULL,
            [Ignore] [bit] NULL,
            [Path] [nvarchar](max) NULL,
            [FolderDepth] [int] NULL,
            [ParentFolder] [nvarchar](max) NULL,
            [RelativeFolder] [nvarchar](max) NULL,
            [HasMacro] [bit] NULL,
            [HasLink] [bit] NULL,
            [HasPattern] [bit] NULL,
            [OfficeOpen] [bit] NULL,
            [PathLength] [int] NULL,
            [Error] [nvarchar](max) NULL,
            [Size] [float] NULL,
            [ScanCreatedDate] [datetime] NULL,
            CONSTRAINT [PK_ScanFile] PRIMARY KEY CLUSTERED ([Id] ASC)
        ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
    END
"
SqlQueryInsert($query)

# -------------------------------------------------
# ScanJob
# -------------------------------------------------
$query = "
    USE [$databaseName]

    IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ScanJob' AND schema_id = SCHEMA_ID('dbo'))
    BEGIN
        CREATE TABLE [dbo].[ScanJob](
            [Id] [int] IDENTITY(1,1) NOT NULL,
            [OwnerId] [int] NOT NULL,
            [Email] [nvarchar](max) NULL,
            [FileCountDisk] [int] NULL,
            [FileCountCrawl] [int] NULL,
            [MacroCount] [int] NULL,
            [LinkCount] [int] NULL,
            [OfficeConversion] [bit] NULL,
            [OfficeCleanup] [bit] NULL,
            [Extensions] [nvarchar](max) NULL,
            [FileSizeDisk] [float] NULL,
            [FileSizeCrawl] [float] NULL,
            [ErrorCount] [int] NULL,
            [OfficeErrorCount] [int] NULL,
            [OldOfficeCount] [int] NULL,
            [PathLengthCount] [int] NULL,
            [Migration] [bit] NULL,
            [ApiNeedToMigrateCount] [int] NULL,
            [ApiCompletedCount] [int] NULL,
            [ApiErrorCount] [int] NULL,
            [ApiReportPath] [nvarchar](max) NULL,
            [NoAccessCount] [int] NULL,
            [CreatedDate] [datetime] NULL,
            [SpmtProcess] [int] NULL,
            [SpmtMessage] [nvarchar](max) NULL,
            CONSTRAINT [PK_ScanJob] PRIMARY KEY CLUSTERED ([OwnerId] ASC)
        ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
    END
"
SqlQueryInsert($query)

# -------------------------------------------------
# MigrationJob
# -------------------------------------------------
$query = "
    USE [$databaseName]

    IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'MigrationJob' AND schema_id = SCHEMA_ID('dbo'))
    BEGIN
        CREATE TABLE [dbo].[MigrationJob](
            [Id] [int] IDENTITY(1,1) NOT NULL,
            [OwnerId] [int] NULL,
            [Email] [nvarchar](max) NULL,
            [FileCountDisk] [int] NULL,
            [FileCountCrawl] [int] NULL,
            [MacroCount] [int] NULL,
            [OfficeConversion] [bit] NULL,
            [OfficeCleanup] [bit] NULL,
            [Extensions] [nvarchar](max) NULL,
            [FileSizeDisk] [float] NULL,
            [FileSizeCrawl] [float] NULL,
            [ErrorCount] [int] NULL,
            [OfficeErrorCount] [int] NULL,
            [OldOfficeCount] [int] NULL,
            [PathLengthCount] [int] NULL,
            [Migration] [bit] NULL,
            [ApiNeedToMigrateCount] [int] NULL,
            [ApiCompletedCount] [int] NULL,
            [ApiErrorCount] [int] NULL,
            [ApiReportPath] [nvarchar](max) NULL,
            [NoAccessCount] [int] NULL,
            [CreatedDate] [datetime] NULL,
            [SpmtProcess] [int] NULL,
            [SpmtMessage] [nvarchar](max) NULL,
            [UserReportPath] [nvarchar](max) NULL
        ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
    END
"
SqlQueryInsert($query)

# -------------------------------------------------
# MigrationFile
# -------------------------------------------------
$query = "
    USE [$databaseName]

    IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'MigrationFile' AND schema_id = SCHEMA_ID('dbo'))
    BEGIN
        CREATE TABLE [dbo].[MigrationFile](
            [Id] [int] IDENTITY(1,1) NOT NULL,
            [FileName] [nvarchar](max) NULL,
            [Location] [nvarchar](max) NULL,
            [Created] [datetime] NULL,
            [Modified] [datetime] NULL,
            [Author] [nvarchar](max) NULL,
            [Extension] [nvarchar](max) NULL,
            [Folder01] [nvarchar](max) NULL,
            [Folder02] [nvarchar](max) NULL,
            [Folder03] [nvarchar](max) NULL,
            [Folder04] [nvarchar](max) NULL,
            [Folder05] [nvarchar](max) NULL,
            [Folder06] [nvarchar](max) NULL,
            [Folder07] [nvarchar](max) NULL,
            [Folder08] [nvarchar](max) NULL,
            [Folder09] [nvarchar](max) NULL,
            [Folder10] [nvarchar](max) NULL,
            [Folder11] [nvarchar](max) NULL,
            [Folder12] [nvarchar](max) NULL,
            [Folder13] [nvarchar](max) NULL,
            [Folder14] [nvarchar](max) NULL,
            [Folder15] [nvarchar](max) NULL,
            [Folder16] [nvarchar](max) NULL,
            [Folder17] [nvarchar](max) NULL,
            [Folder18] [nvarchar](max) NULL,
            [Folder19] [nvarchar](max) NULL,
            [Folder20] [nvarchar](max) NULL,
            [Normalized] [bit] NULL,
            [OwnerId] [int] NULL,
            [Ignore] [bit] NULL,
            [Path] [nvarchar](max) NULL,
            [FolderDepth] [int] NULL,
            [ParentFolder] [nvarchar](max) NULL,
            [RelativeFolder] [nvarchar](max) NULL,
            [HasMacro] [bit] NULL,
            [OfficeOpen] [bit] NULL,
            [PathLength] [int] NULL,
            [Error] [nvarchar](max) NULL,
            [Size] [float] NULL,
            [ScanCreatedDate] [datetime] NULL
        ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
    END
"
SqlQueryInsert($query)

# -------------------------------------------------
# Source
# -------------------------------------------------
$query = "
    USE [$databaseName]

    IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Source' AND schema_id = SCHEMA_ID('dbo'))
    BEGIN
        CREATE TABLE [dbo].[Source](
            [Id] [int] IDENTITY(1,1) NOT NULL,
            [SamAccountName] [nvarchar](max) NULL,
            [LastLogonDate] [datetime] NULL,
            [BatchNumber] [int] NULL,
            [BatchDate] [datetime] NULL,
            [Building] [nvarchar](max) NULL,
            [Title] [nvarchar](max) NULL,
            [AllDayEvening] [nvarchar](max) NULL,
            [Role] [nvarchar](max) NULL,
            [ADHomeDirectory] [nvarchar](max) NULL,
            [Counter] [float] NULL,
            [ADMigrateGroupMember] [bit] NULL,
            [UserPrincipalName] [nvarchar](max) NULL,
            [OneDriveUrl] [nvarchar](max) NULL,
            [TaskJson] [nvarchar](max) NULL,
            [DestinationLibrary] [nvarchar](max) NULL,
            [DestinationFolder] [nvarchar](max) NULL,
            [FileCount] [float] NULL,
            [FileSize] [float] NULL,
            [MigrationType] [nvarchar](max) NULL,
            [SourceSiteUrl] [nvarchar](max) NULL,
            [DestinationSiteName] [nvarchar](max) NULL,
            [SourceLibrary] [nvarchar](max) NULL,
            [SourceTeamName] [nvarchar](max) NULL,
            [DestinationTeamName] [nvarchar](max) NULL,
            [Validated] [bit] NULL,
            [SourceTeamId] [nvarchar](max) NULL,
            [SourceEmail] [nvarchar](max) NULL,
            [Geo] [nvarchar](max) NULL,
            [SourceFolder] [nvarchar](max) NULL,
            [Wave] [nvarchar](max) NULL
        ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
    END
"
SqlQueryInsert($query)

# -------------------------------------------------
# ScanLink
# -------------------------------------------------
$query = "
    USE [$databaseName]

    IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ScanLink' AND schema_id = SCHEMA_ID('dbo'))
    BEGIN
        CREATE TABLE [dbo].[ScanLink](
            [Id] [int] IDENTITY(1,1) NOT NULL,
            [OwnerId] [int] NULL,
            [FileId] [int] NULL,
            [Url] [nvarchar](max) NULL,
            NULL
        ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
    END
"
SqlQueryInsert($query)

# -------------------------------------------------
# ScanMatches
# -------------------------------------------------
$query = "
    USE [$databaseName]

    IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ScanMatches' AND schema_id = SCHEMA_ID('dbo'))
    BEGIN
        CREATE TABLE [dbo].[ScanMatches](
            [Id] [int] IDENTITY(1,1) NOT NULL,
            [OwnerId] [int] NULL,
            [FileId] [int] NULL,
            [Match] [nvarchar](max) NULL,
            [PatternName] [nvarchar](max) NULL,
            NULL
        ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
    END
"
SqlQueryInsert($query)

# -------------------------------------------------
# ScanPatterns
# -------------------------------------------------
$query = "
    USE [$databaseName]

    IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ScanPatterns' AND schema_id = SCHEMA_ID('dbo'))
    BEGIN
        CREATE TABLE [dbo].[ScanPatterns](
            [Id] [int] IDENTITY(1,1) NOT NULL,
            [Name] [nvarchar](max) NULL,
            [Pattern] [nvarchar](max) NULL
        ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
    END
"
SqlQueryInsert($query)
