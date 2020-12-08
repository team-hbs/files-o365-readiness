$query = "  USE [FileToOneDrive]

            /****** Object:  Table [dbo].[Config]    Script Date: 8/11/2020 10:28:35 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

            CREATE TABLE [dbo].[Config](
                [Id] [int] IDENTITY(1,1) NOT NULL,
                [mode] [nvarchar](max) NULL,
                [source] [nvarchar](max) NULL,
                [report] [nvarchar](max) NULL,
                [notifications] [nvarchar](max) NULL,
                [email] [nvarchar](max) NULL,
            PRIMARY KEY CLUSTERED 
            (
                [Id] ASC
            )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)

$query = "  USE [FileToOneDrive]

            /****** Object:  Table [dbo].[Files_Batch_Users]    Script Date: 8/4/2020 11:47:51 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

            CREATE TABLE [dbo].[Files_Batch_Users](
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
                [TaskJson] [nvarchar](max) NULL
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)

$query = "  USE [FileToOneDrive]

            /****** Object:  Table [dbo].[Files_OneDrive]    Script Date: 8/10/2020 10:25:53 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

            CREATE TABLE [dbo].[Files_OneDrive](
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
                [OfficeOpen] [bit] NULL,
                [PathLength] [int] NULL,
                [Error] [nvarchar](max) NULL,
                [Size] [float] NULL,
                [Ignore] [bit] NULL
            PRIMARY KEY CLUSTERED 
            (
                [Id] ASC
            )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)

$query = "  USE [FileToOneDrive]

            /****** Object:  Table [dbo].[Files_Users]    Script Date: 8/11/2020 10:22:11 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

            CREATE TABLE [dbo].[Files_Users](
                [OwnerId] [int] NOT NULL,
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
                [CreatedDate] [float] NULL
            PRIMARY KEY CLUSTERED 
            (
                [OwnerId] ASC
            )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)

$query = "  USE [FileToOneDrive]

            /****** Object:  Table [dbo].[Files_OneDrive_Conversion]    Script Date: 8/4/2020 11:51:33 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

            CREATE TABLE [dbo].[Files_OneDrive_Conversion](
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

$query = "  USE [FileToOneDrive]

            /****** Object:  Table [dbo].[Files_Users_Conversion]    Script Date: 8/4/2020 11:52:09 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

            CREATE TABLE [dbo].[Files_Users_Conversion](
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
                [SpmtMessage] [nvarchar](max) NULL
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)