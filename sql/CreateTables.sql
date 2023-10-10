USE [Migration]
GO

/****** Object:  Table [dbo].[MigrationQueue]    Script Date: 11/14/2022 2:20:55 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE TABLE [dbo].[Batch](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[RunDate] [datetime] NULL,
	[Server] [nvarchar](max) NULL,
	[BatchNumber] [int] NULL,
	[Status] [int] NULL,
	[Wave] [nvarchar](max) NULL,
	[CutoffDate] [datetime] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Event](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[OwnerId] [int] NULL,
	[EventType] [nvarchar](max) NULL,
	[EventDate] [datetime] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[GlobalConfig](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[Key] [nvarchar](max) NULL,
	[Value] [nvarchar](max) NULL,
	[Server] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

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
GO

CREATE TABLE [dbo].[MigrationQueue](
	[OwnerId] [int] NOT NULL,
	[Created] [datetime] NULL,
	[Server] [nvarchar](max) NULL,
	[BatchNumber] [int] NULL,
 CONSTRAINT [PK_MigrationQueue] PRIMARY KEY CLUSTERED 
(
	[OwnerId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


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
	[OfficeOpen] [bit] NULL,
	[PathLength] [int] NULL,
	[Error] [nvarchar](max) NULL,
	[Size] [float] NULL,
	[ScanCreatedDate] [datetime] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[ScanJob](
	[OwnerId] [int] NULL,
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
	[SpmtMessage] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


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
GO


