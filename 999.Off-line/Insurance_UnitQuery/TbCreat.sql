USE [UIS]
GO

/****** Object:  Table [dbo].[test_case]    Script Date: 2022/1/25 ¤W¤È 06:40:24 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Insurance_UnitQuery](
	[CaseI] [int] NULL,
	[Name] [varchar](50) NULL,
	[ID] [varchar](50) NULL,
	[EmpI] [int] NULL,
	[PersonI] [int] NULL,
	[Type] [varchar](50) NULL,
	[AddressI] [int] NULL,
	[Response] [nvarchar](4000) NULL,
	[Status] [varchar](50) NULL,
	[insertdate] [datetime] NULL
) ON [PRIMARY]

GO

delete from [dbo].[Insurance_UnitQuery]

select * from [dbo].[Insurance_UnitQuery]


