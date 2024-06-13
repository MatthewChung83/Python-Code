USE [UIS]
GO

/****** Object:  Table [dbo].[FJUD_Default_AD]    Script Date: 2022/3/6 下午 04:22:04 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[FJUD_Default_AD](
	[court] [nvarchar](4000) NULL,
	[court_word] [nvarchar](4000) NULL,
	[space] [nvarchar](4000) NULL,
	[event] [nvarchar](4000) NULL,
	[ann_date] [nvarchar](4000) NULL,
	[link] [nvarchar](4000) NULL,
	[qtype] [nvarchar](4000) NULL,
	[s_parse] [nvarchar](4000) NULL,
	[e_parse] [nvarchar](4000) NULL,
	[main_sta] [nvarchar](4000) NULL,
	[claimant] [nvarchar](4000) NULL,
	[opposite] [nvarchar](4000) NULL,
	[stakeholder] [nvarchar](4000) NULL,
	[protester] [nvarchar](4000) NULL,
	[plaintiff] [nvarchar](4000) NULL,
	[defendant] [nvarchar](4000) NULL,
	[dissenter] [nvarchar](4000) NULL,
	[article] [nvarchar](4000) NULL,
	[insertdate] [datetime] NULL
) ON [PRIMARY]
GO


USE [UIS]
GO

/****** Object:  Table [dbo].[FJUD_Default_AD_detail]    Script Date: 2022/3/6 下午 04:22:14 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[FJUD_Default_AD_detail](
	[link] [nvarchar](4000) NULL,
	[character] [nvarchar](4000) NULL,
	[name] [nvarchar](4000) NULL,
	[insertdate] [datetime] NULL
) ON [PRIMARY]
GO