$databaseName = GetConfig('DatabaseName')

$query = "  USE [" + $databaseName  + "]

            /****** Object:  Table [dbo].[Batch]    Script Date: 8/4/2020 11:47:51 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

            CREATE TABLE [dbo].[Batch](
			    [Id] [int] IDENTITY(1,1) NOT NULL,
                [RunDate] [datetime] NULL,
                [Server] [nvarchar](max) NULL,
				[BatchNumber] [int] NULL,
				[Status] [int] NULL
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)


$query = "  USE [" + $databaseName  + "]

            /****** Object:  Table [dbo].[Config]    Script Date: 8/4/2020 11:47:51 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

            CREATE TABLE [dbo].[Batch](
			    [Id] [int] IDENTITY(1,1) NOT NULL,
                [RunDate] [datetime] NULL,
                [Server] [nvarchar](max) NULL,
				[BatchNumber] [int] NULL,
				[Status] [int] NULL
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)

SqlQueryInsert($query)

$query = "  USE [" + $databaseName  + "]

            /****** Object:  Table [dbo].[GlobalConfig]    Script Date: 2/24/2021 11:22:00 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

            CREATE TABLE [dbo].[GlobalConfig](
			    [Id] [int] IDENTITY(1,1) NOT NULL,
                [Key] [nvarchar](max) NULL,
                [Value] [nvarchar](max) NULL,
                [Server] [nvarchar](max) NULL,
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)

$query = "  USE [" + $databaseName  + "]

            /****** Object:  Table [dbo].[Event]    Script Date: 2/24/2021 11:22:00 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

            CREATE TABLE [dbo].[Event](
			    [Id] [int] IDENTITY(1,1) NOT NULL,
                [OwnerId] [int]  NULL,
                [EventType] [nvarchar](max) NULL,
				[EventDate] [datetime] NULL
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)

$query = "  USE [" + $databaseName  + "]

            /****** Object:  Table [dbo].[Source]    Script Date: 8/4/2020 11:47:51 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

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
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)

$query = "  USE [" + $databaseName  + "]

            /****** Object:  Table [dbo].[ScanFile]    Script Date: 8/10/2020 10:25:53 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

            CREATE TABLE [dbo].[ScanFile](
                [Id] [int] IDENTITY(1,1) NOT NULL,
                [FileName] [nvarchar](max) NULL,
                [Location] [nvarchar](max) NULL,
                [Created] [float] NULL,
                [Modified] [float] NULL,
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
                [Ignore] [bit] NULL,
				[ScanCreatedDate] [float] NULL
            PRIMARY KEY CLUSTERED 
            (
                [Id] ASC
            )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)

$query = "  USE [" + $databaseName  + "]

            /****** Object:  Table [dbo].[ScanJob]    Script Date: 8/11/2020 10:22:11 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

            CREATE TABLE [dbo].[ScanJob](
			    [Id] [int] IDENTITY(1,1) NOT NULL,
                [OwnerId] [int] NOT NULL,
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
                [CreatedDate] [float] NULL
            PRIMARY KEY CLUSTERED 
            (
                [OwnerId] ASC
            )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)

$query = "  USE [" + $databaseName + "]

            /****** Object:  Table [dbo].[ScanFile]    Script Date: 8/4/2020 11:51:33 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

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
				[ScanCreatedDate] [float] NULL
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)

$query = "  USE [" + $databaseName + "]

            /****** Object:  Table [dbo].[MigrationFile]    Script Date: 8/4/2020 11:51:33 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

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
                [Size] [float] NULL
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)

$query = "  USE [" + $databaseName + "]

            /****** Object:  Table [dbo].[ScanJob]    Script Date: 8/4/2020 11:52:09 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

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
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)

$query = "  USE [" + $databaseName + "]

            /****** Object:  Table [dbo].[MigrationJob]    Script Date: 8/4/2020 11:52:09 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

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
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)