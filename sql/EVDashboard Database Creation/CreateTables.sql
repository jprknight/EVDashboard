USE [EVDashboard]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[EnabledUsers](
	EnabledUserID INT IDENTITY(1,1) PRIMARY KEY,
	RecordCreateTimestamp datetime DEFAULT GETDATE() NOT NULL,
	DisplayName varchar(max) NOT NULL,
	ExchangeServer varchar(max) NOT NULL,
	EVDatabase varchar(max) NOT NULL,
	NumItemsMailbox int NOT NULL,
	NumItemsArchive int NOT NULL,
	MailboxSizeMB int NOT NULL,
	ArchiveSizeMB int NOT NULL,
	TotalSizeMB int NOT NULL,
	ArchiveCreated datetime NOT NULL,
	ArchiveUpdated datetime NOT NULL,
	ExchangeState varchar(max) NOT NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF



USE [EVDashboard]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[DisabledUsers](
	DisabledUserID INT IDENTITY(1,1) PRIMARY KEY,
	RecordCreateTimestamp datetime DEFAULT GETDATE() NOT NULL,
	Mailbox varchar(max) NOT NULL,
	WarningMB INT NOT NULL,
	SendMB INT NOT NULL,
	ReceiveMB INT NOT NULL,
	MailboxSizeMB INT NOT NULL,
	NumItemsMailbox INT NOT NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF



USE [EVDashboard]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[ArchiveStates](
	ArchiveStatesID INT IDENTITY(1,1) PRIMARY KEY,
	RecordCreateTimestamp datetime DEFAULT GETDATE() NOT NULL,
	NumMailboxes INT NOT NULL,
	ArchiveState INT NOT NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF



USE [EVDashboard]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[DiskUsageInfo](
	DiskUsageInfoID INT IDENTITY(1,1) PRIMARY KEY,
	RecordCreateTimestamp datetime DEFAULT GETDATE() NOT NULL,
	EVServer varchar(max) NOT NULL,
	EVRole varchar(max) NOT NULL,
	Drive varchar(max) NOT NULL,
	VolumeName varchar(max) NOT NULL,
	SizeGB varchar(max) NOT NULL,
	UsedGB varchar(max) NOT NULL,
	FreeGB varchar(max) NOT NULL,
	PercentageFree varchar(max) NOT NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF



USE [EVDashboard]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[HourlyArchivingRates](
	HourlyArchivingRatesID INT IDENTITY(1,1) PRIMARY KEY,
	RecordCreateTimestamp datetime DEFAULT GETDATE() NOT NULL,
	EVServer varchar(max) NOT NULL,
	EVRole varchar(max) NOT NULL,
	DBServer varchar(max) NOT NULL,
	EVDatabase varchar(max) NOT NULL,
	ArchivePeriod varchar(max) NOT NULL,
	MsgCount INT NOT NULL,
	AverageSize INT NOT NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF



USE [EVDashboard]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[OverSendLimit](
	OverSendLimitID INT IDENTITY(1,1) PRIMARY KEY,
	RecordCreateTimestamp datetime DEFAULT GETDATE() NOT NULL,
	Mailbox varchar(max) NOT NULL,
	OverSendMB INT NOT NULL,
	WarningMB INT NOT NULL,
	SendMB INT NOT NULL,
	ReceiveMB INT NOT NULL,
	MailboxSizeMB INT NOT NULL,
	NumItemsMailbox INT NOT NULL,
	ExchangeServer varchar(max) NOT NULL,
	LastArchived datetime NOT NULL,
	ExchangeState varchar(max) NOT NULL,
	PolicyUsed varchar(max) NOT NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF



USE [EVDashboard]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[OverWarningLimit](
	OverWarningLimitID INT IDENTITY(1,1) PRIMARY KEY,
	RecordCreateTimestamp datetime DEFAULT GETDATE() NOT NULL,
	Mailbox varchar(max) NOT NULL,
	OverWarningMB INT NOT NULL,
	WarningMB INT NOT NULL,
	SendMB INT NOT NULL,
	ReceiveMB INT NOT NULL,
	MailboxSizeMB INT NOT NULL,
	NumItemsMailbox INT NOT NULL,
	ExchangeServer varchar(max) NOT NULL,
	LastArchived datetime NOT NULL,
	ExchangeState varchar(max) NOT NULL,
	PolicyUsed varchar(max) NOT NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF



USE [EVDashboard]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[OverReceiveLimit](
	OverReceiveLimitID INT IDENTITY(1,1) PRIMARY KEY,
	RecordCreateTimestamp datetime DEFAULT GETDATE() NOT NULL,
	Mailbox varchar(max) NOT NULL,
	OverReceiveMB INT NOT NULL,
	WarningMB INT NOT NULL,
	SendMB INT NOT NULL,
	ReceiveMB INT NOT NULL,
	MailboxSizeMB INT NOT NULL,
	NumItemsMailbox INT NOT NULL,
	ExchangeServer varchar(max) NOT NULL,
	LastArchived datetime NOT NULL,
	ExchangeState varchar(max) NOT NULL,
	PolicyUsed varchar(max) NOT NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF



USE [EVDashboard]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[ExchangeMailboxes](
	ExchangeMailboxesID INT IDENTITY(1,1) PRIMARY KEY,
	RecordCreateTimestamp datetime DEFAULT GETDATE() NOT NULL,
	ExchangeServer varchar(max) NOT NULL,
	Mailstore varchar(max) NOT NULL,
	MbxCount INT NOT NULL,
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF



USE [EVDashboard]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[ExchangeTransactionLogs](
	ExchangeTransactionLogsID INT IDENTITY(1,1) PRIMARY KEY,
	RecordCreateTimestamp datetime DEFAULT GETDATE() NOT NULL,
	ExchangeServer varchar(max) NOT NULL,
	StorageGroupName varchar(max) NOT NULL,
	LogFilePath varchar(max) NOT NULL,
	LogFilePrefix varchar(max) NOT NULL,
	NumFilesInDirectory int NOT NULL,
	DiskSpaceUsedMB int NOT NULL,
	OldestLogFileInDirectory varchar(max) NOT NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF



USE [EVDashboard]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[ServiceAlerts](
	ServiceAlertsID INT IDENTITY(1,1) PRIMARY KEY,
	RecordCreateTimestamp datetime DEFAULT GETDATE() NOT NULL,
	DNSAlias varchar(max) NOT NULL,
	Server varchar(max) NOT NULL,
	Role varchar(max) NOT NULL,
	Service varchar(max) NOT NULL,
	Status varchar(max) NOT NULL,
	ShowHide varchar(max),
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF




USE [EVDashboard]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[OrganisationArchivingHistory](
	OrganisationArchivingHistoryID INT IDENTITY(1,1) PRIMARY KEY,
	RecordCreateTimestamp datetime DEFAULT GETDATE() NOT NULL,
	ArchivedDate varchar(max) NOT NULL,
	DNSAlias varchar(max) NOT NULL,
	EVServer varchar(max) NOT NULL,
	EVDatabase varchar(max) NOT NULL,
	RecordCount INT NOT NULL,
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF



USE [EVDashboard]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[MSMQMessageCountAlerts](
	MSMQMessageCountAlertsID INT IDENTITY(1,1) PRIMARY KEY,
	RecordCreateTimestamp datetime DEFAULT GETDATE() NOT NULL,
	EVServer varchar(max) NOT NULL,
	MSMQPath varchar(max) NOT NULL,
	RecordCount varchar(max) NOT NULL,
	FileSize varchar(max) NOT NULL,
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF



USE [EVDashboard]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[EnabledMailboxes](
	EnabledMailboxesID INT IDENTITY(1,1) PRIMARY KEY,
	RecordCreateTimestamp datetime DEFAULT GETDATE() NOT NULL,
	Mailbox varchar(max) NOT NULL,
	ExchangeServer varchar(max) NOT NULL,
	EVServer varchar(max) NULL,
	DBServer varchar(max) NULL,
	EVDatabase varchar(max) NULL,
	NumItemsMailbox INT NOT NULL,
	MailboxSizeMB INT NOT NULL,
	ExchangeState INT NOT NULL,
	DefaultVaultID varchar(max) NOT NULL,
	NumItemsArchive INT NULL,
	ArchiveSizeMB INT NULL,
	TotalSizeMB INT NULL,
	ArchiveCreated varchar(max) NULL,
	ArchiveUpdated varchar(max) NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF



USE [EVDashboard]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[ArchivingActivity](
	ArchivingActivityID INT IDENTITY(1,1) PRIMARY KEY,
	RecordCreateTimestamp datetime DEFAULT GETDATE() NOT NULL,
	Mailbox varchar(max) NOT NULL,
	ExchangeServer varchar(max) NOT NULL,
	EVServer varchar(max) NOT NULL,
	DBServer varchar(max) NOT NULL,
	EVDatabase varchar(max) NOT NULL,
	NumItemsArchived INT NULL,
	root_rootidentity varchar(max) NOT NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF


USE [EVDashboard]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[UsageReportByArchive](
	UsageReportByarchiveID INT IDENTITY(1,1) PRIMARY KEY,
	RecordCreateTimestamp datetime DEFAULT GETDATE() NOT NULL,
	Mailbox varchar(max) NOT NULL,
	ArchivePeriod varchar(max) NULL,
	ItemCount varchar(max) NULL,
	EMEDefaultVaultID varchar(max) NOT NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF




USE [EVDashboard]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[EVServerEventLogs](
	EVServerEventLogsID INT IDENTITY(1,1) PRIMARY KEY,
	RecordCreateTimestamp datetime DEFAULT GETDATE() NOT NULL,
	EVServer varchar(max) NOT NULL,
	EventCode INT NOT NULL,
	TimeWritten varchar(max) NULL,
	EventType varchar(max) NOT NULL,
	Message varchar(max) NOT NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF




USE [EVDashboard]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[ExchangeServerEventLogs](
	ExchangeServerEventLogsID INT IDENTITY(1,1) PRIMARY KEY,
	RecordCreateTimestamp datetime DEFAULT GETDATE() NOT NULL,
	ExchangeServer varchar(max) NOT NULL,
	EventCode INT NOT NULL,
	TimeWritten varchar(max) NULL,
	EventType varchar(max) NOT NULL,
	Message varchar(max) NOT NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF




USE [EVDashboard]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[SISReport](
	SISReportID INT IDENTITY(1,1) PRIMARY KEY,
	RecordCreateTimestamp datetime DEFAULT GETDATE() NOT NULL,
	ArchiveName varchar(max) NOT NULL,
	ItemsShared INT NOT NULL,
	ItemsArchived INT NOT NULL,
	ArchiveSizeMB INT NOT NULL,
	EVDatabase varchar(max) NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF