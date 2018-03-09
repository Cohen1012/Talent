USE [is_recruit_db]
GO

/****** Object:  Table [dbo].[Project_Experience]    Script Date: 2018/3/8 ¤U¤È 04:02:14 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Project_Experience](
	[Interview_Id] [int] NOT NULL,
	[Company] [nvarchar](50) NOT NULL,
	[Project_Name] [nvarchar](50) NOT NULL,
	[OS] [varchar](10) NULL,
	[Database] [varchar](20) NULL,
	[Position] [nvarchar](10) NULL,
	[Language] [varchar](20) NULL,
	[Tools] [nvarchar](20) NULL,
	[Description] [nvarchar](max) NULL,
	[Start_End_Date] [nvarchar](20) NULL,
 CONSTRAINT [PK_Project_Experience] PRIMARY KEY CLUSTERED 
(
	[Interview_Id] ASC,
	[Company] ASC,
	[Project_Name] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO



CREATE TABLE [dbo].[Files](
	[Interview_Id] [int] NOT NULL,
	[File_Path] [nvarchar](255) NOT NULL,
	[Belong] [nvarchar](20) NOT NULL,
 CONSTRAINT [PK_Files] PRIMARY KEY CLUSTERED 
(
	[Interview_Id] ASC,
	[File_Path] ASC,
	[Belong] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO


CREATE TABLE [dbo].[Interview_Info](
	[Vacancies] [nvarchar](20) NULL,
	[Interview_Date] [date] NULL,
	[Interview_Id] [int] IDENTITY(1,1) NOT NULL,
	[Contact_Id] [int] NOT NULL,
	[Name] [nvarchar](10) NULL,
	[Sex] [nvarchar](1) NULL,
	[Birthday] [date] NULL,
	[Married] [nvarchar](2) NULL,
	[Mail] [varchar](50) NULL,
	[Adress] [nvarchar](200) NULL,
	[CellPhone] [varchar](50) NULL,
	[Image] [nvarchar](50) NULL,
	[Expertise_Language] [nvarchar](max) NULL,
	[Expertise_Tools] [nvarchar](max) NULL,
	[Expertise_Devops] [nvarchar](max) NULL,
	[Expertise_OS] [nvarchar](max) NULL,
	[Expertise_BigData] [nvarchar](max) NULL,
	[Expertise_DataBase] [nvarchar](max) NULL,
	[Expertise_Certification] [nvarchar](max) NULL,
	[IsStudy] [nvarchar](max) NULL,
	[IsService] [nvarchar](10) NULL,
	[Relatives_Relationship] [nvarchar](10) NULL,
	[Relatives_Name] [nvarchar](10) NULL,
	[Care_Work] [nvarchar](max) NULL,
	[Hope_Salary] [nvarchar](20) NULL,
	[When_Report] [nvarchar](20) NULL,
	[Advantage] [nvarchar](max) NULL,
	[Disadvantages] [nvarchar](max) NULL,
	[Hobby] [nvarchar](50) NULL,
	[Attract_Reason] [nvarchar](max) NULL,
	[Future_Goal] [nvarchar](max) NULL,
	[Hope_Supervisor] [nvarchar](max) NULL,
	[Hope_Promise] [nvarchar](max) NULL,
	[Introduction] [nvarchar](max) NULL,
	[Appointment] [nvarchar](6) NULL,
	[Results_Remark] [nvarchar](max) NULL,
	[During_Service] [nvarchar](20) NULL,
	[Exemption_Reason] [nvarchar](50) NULL,
	[Urgent_Contact_Person] [nvarchar](10) NULL,
	[Urgent_Relationship] [nvarchar](10) NULL,
	[Urgent_CellPhone] [varchar](10) NULL,
	[Education] [nvarchar](max) NULL,
	[Language] [nvarchar](max) NULL,
	[Work_Experience] [nvarchar](max) NULL,
	[Expertise_Tools_Framwork] [nvarchar](100) NULL,
 CONSTRAINT [PK_Interview_Info] PRIMARY KEY CLUSTERED 
(
	[Interview_Id] ASC,
	[Contact_Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO



CREATE TABLE [dbo].[Interview_Comments](
	[Interviewer] [nvarchar](10) NULL,
	[Result] [nvarchar](max) NULL,
	[Interview_Id] [int] NOT NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


CREATE TABLE [dbo].[Contact_Status_List](
	[Name] [nvarchar](20) NULL
) ON [PRIMARY]
GO


CREATE TABLE [dbo].[Code](
	[Code_Id] [varchar](20) NOT NULL,
	[Contact_Id] [int] NOT NULL,
	[Id] [int] IDENTITY(1,1) NOT NULL,
 CONSTRAINT [PK_Code_1] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO



CREATE TABLE [dbo].[Contact_Info](
	[Name] [nvarchar](20) NULL,
	[Sex] [nvarchar](1) NULL,
	[Mail] [varchar](50) NULL,
	[CellPhone] [varchar](50) NULL,
	[UpdateTime] [datetime] NULL,
	[Cooperation_Mode] [nvarchar](2) NULL,
	[Status] [nvarchar](2) NULL,
	[Place] [nvarchar](200) NULL,
	[Skill] [nvarchar](200) NULL,
	[Contact_Id] [int] IDENTITY(1,1) NOT NULL,
	[Year] [varchar](4) NULL,
 CONSTRAINT [PK_Contact_Info] PRIMARY KEY CLUSTERED 
(
	[Contact_Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO



CREATE TABLE [dbo].[Contact_Situation](
	[Contact_Status] [nvarchar](10) NULL,
	[Remarks] [text] NULL,
	[Contact_Date] [datetime] NULL,
	[Contact_Id] [int] NOT NULL,
	[Contact_status_Id] [int] IDENTITY(1,1) NOT NULL,
 CONSTRAINT [PK_Contact_Situation_1] PRIMARY KEY CLUSTERED 
(
	[Contact_status_Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


CREATE TABLE [dbo].[Member](
	[Account] [varchar](50) NOT NULL,
	[Password] [varchar](50) NULL,
	[States] [nvarchar](6) NULL,
 CONSTRAINT [PK_Member_1] PRIMARY KEY CLUSTERED 
(
	[Account] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

