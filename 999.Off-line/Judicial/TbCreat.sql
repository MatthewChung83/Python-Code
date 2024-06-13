/* 法文主檔 */
USE [UIS]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[judicial](
        keyword [nvarchar](4000) NULL,
        court [nvarchar](4000) NULL,
        judicial_no [nvarchar](4000) NULL,
        recipient [nvarchar](4000) NULL,
        area [nvarchar](4000) NULL,
        doc_type [nvarchar](4000) NULL,
        publishing_date [nvarchar](4000) NULL,
        judicial_type [nvarchar](4000) NULL,
        judicial_title [nvarchar](4000) NULL,
        judicial_doc_date [nvarchar](4000) NULL,
        judicial_doc_no [nvarchar](4000) NULL,
        judicial_doc_subject [nvarchar](4000) NULL,
        judicial_doc_basis [nvarchar](4000) NULL,
        judicial_doc_anncm [nvarchar](4000) NULL,
        judicial_doc_pdate [nvarchar](4000) NULL,
        judicial_doc_udate [nvarchar](4000) NULL,
        judicial_doc_unit [nvarchar](4000) NULL,
        judicial_content [nvarchar](4000) NULL,
        insertdate [datetime] NULL,
		names  [nvarchar](4000) NULL
) ON [PRIMARY]

GO


/* 債務人檔 */
USE [UIS]
GO

/****** Object:  Table [dbo].[judicial]    Script Date: 2022/1/16 下午 10:34:08 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[judicial_person](
        judicial_no [nvarchar](4000) NULL,
        recipient [nvarchar](4000) NULL,
        publishing_date [nvarchar](4000) NULL,
        judicial_doc_date [nvarchar](4000) NULL,
        judicial_doc_no [nvarchar](4000) NULL,
        judicial_doc_subject [nvarchar](4000) NULL,
		debtors [nvarchar](4000) NULL,
		debtor [nvarchar](4000) NULL,
        insertdate [datetime] NULL,
) ON [PRIMARY]
GO


/*
drop table judicial_person
drop table judicial

delete from judicial_person
delete from judicial
*/