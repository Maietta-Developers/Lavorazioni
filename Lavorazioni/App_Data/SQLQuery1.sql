CREATE TABLE [dbo].[Appointments] (
        [UniqueID] [int] IDENTITY (1, 1) NOT NULL ,
        [Type] [int] NULL ,
        [StartDate] [smalldatetime] NULL ,
        [EndDate] [smalldatetime] NULL ,
        [AllDay] [bit] NULL ,
        [Subject] [nvarchar] (50) NULL ,
        [Location] [nvarchar] (50) NULL ,
        [Description] [nvarchar](max) NULL ,
        [Status] [int] NULL ,
        [Label] [int] NULL ,
        [ResourceID] [int] NULL ,
        [ResourceIDs] [nvarchar](max) NULL ,
        [ReminderInfo] [nvarchar](max) NULL ,
        [RecurrenceInfo] [nvarchar](max) NULL ,
        [CustomField1] [nvarchar](max) NULL 
CONSTRAINT [PK_Appointments] PRIMARY KEY CLUSTERED
(
        [UniqueID] ASC
)
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

CREATE TABLE [dbo].[Resources] (
        [UniqueID] [int] IDENTITY (1, 1) NOT NULL ,
        [ResourceID] [int] NOT NULL ,
        [ResourceName] [nvarchar] (50) NULL ,
        [Color] [int] NULL ,
        [Image] [image] NULL ,
        [CustomField1] [nvarchar](max) NULL 
CONSTRAINT [PK_Resources] PRIMARY KEY CLUSTERED
(
        [UniqueID] ASC
)
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO