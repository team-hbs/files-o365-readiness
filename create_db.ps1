$query = "  USE [FileToOneDrive]
            

            /****** Object:  Table [dbo].[Config]    Script Date: 7/21/2020 8:15:29 AM ******/
            SET ANSI_NULLS ON
            

            SET QUOTED_IDENTIFIER ON
            

            CREATE TABLE [dbo].[Config](
                [Id] [int] IDENTITY(1,1) NOT NULL,
                [source] [text] NULL,
                [mode] [varchar](max) NULL,
                [report] [varchar](max) NULL,
                [notifications] [varchar](max) NULL,
                [email] [varchar](max) NULL,
            PRIMARY KEY CLUSTERED 
            (
                [Id] ASC
            )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)

$query = "  USE [FileToOneDrive]

            /****** Object:  Table [dbo].[Files_Batch_Users]    Script Date: 7/21/2020 8:16:08 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

            CREATE TABLE [dbo].[Files_Batch_Users](
                [Id] [int] IDENTITY(1,1) NOT NULL,
                [BatchNumber] [int] NULL,
                [ADHomeDirectory] [varchar](max) NULL,
                [Counter] [int] NULL,
                [OneDriveUrl] [int] NULL,
                [SamAccountName] [varchar](max) NULL,
            PRIMARY KEY CLUSTERED 
            (
                [Id] ASC
            )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)

$query = "  USE [FileToOneDrive]

            /****** Object:  Table [dbo].[Files_OneDrive]    Script Date: 7/21/2020 8:16:42 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

            CREATE TABLE [dbo].[Files_OneDrive](
                [Id] [int] IDENTITY(1,1) NOT NULL,
                [FileName] [text] NULL,
                [Author] [text] NULL,
                [Extension] [varchar](255) NULL,
                [Folder01] [text] NULL,
                [Folder02] [text] NULL,
                [Folder03] [text] NULL,
                [Folder04] [text] NULL,
                [Folder05] [text] NULL,
                [Folder06] [text] NULL,
                [Folder07] [text] NULL,
                [Folder08] [text] NULL,
                [Folder09] [text] NULL,
                [Folder10] [text] NULL,
                [Folder11] [text] NULL,
                [Folder12] [text] NULL,
                [Folder13] [text] NULL,
                [Folder14] [text] NULL,
                [Folder15] [text] NULL,
                [Folder16] [text] NULL,
                [Folder17] [text] NULL,
                [Folder18] [text] NULL,
                [Folder19] [text] NULL,
                [Folder20] [text] NULL,
                [OwnerId] [int] NULL,
                [Ignore] [varchar](255) NULL,
                [Path] [text] NULL,
                [FolderDepth] [int] NULL,
                [ParentFolder] [text] NULL,
                [RelativeFolder] [text] NULL,
                [HasMacro] [int] NULL,
                [OfficeOpen] [int] NULL,
                [PathLength] [int] NULL,
                [Error] [varchar](max) NULL,
                [Size] [numeric](38, 5) NULL,
                [Created] [numeric](38, 5) NULL,
                [Modified] [numeric](38, 5) NULL,
            PRIMARY KEY CLUSTERED 
            (
                [Id] ASC
            )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)

$query = "  USE [FileToOneDrive]

            /****** Object:  Table [dbo].[Files_Users]    Script Date: 7/21/2020 8:17:41 AM ******/
            SET ANSI_NULLS ON

            SET QUOTED_IDENTIFIER ON

            CREATE TABLE [dbo].[Files_Users](
                [OwnerId] [int] NOT NULL,
                [FileCountDisk] [int] NULL,
                [FileCountCrawl] [int] NULL,
                [MacroCount] [int] NULL,
                [OfficeConversion] [int] NULL,
                [OfficeCleanup] [int] NULL,
                [Extensions] [text] NULL,
                [ErrorCount] [int] NULL,
                [OfficeErrorCount] [int] NULL,
                [OldOfficeCount] [int] NULL,
                [PathLengthCount] [int] NULL,
                [Migration] [int] NULL,
                [ApiNeedToMigrateCount] [int] NULL,
                [ApiCompletedCount] [int] NULL,
                [ApiErrorCount] [int] NULL,
                [ApiReportPath] [int] NULL,
                [NoAccessCount] [int] NULL,
                [CreatedDate] [int] NULL,
                [FileSizeDisk] [numeric](38, 5) NULL,
                [FileSizeCrawl] [numeric](38, 5) NULL,
            PRIMARY KEY CLUSTERED 
            (
                [OwnerId] ASC
            )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

SqlQueryInsert($query)