USE [yamaha_it]
GO

/****** Object:  Table [dbo].[tbl_loan_sale]    Script Date: 13/02/2015 8:57:11 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[tbl_loan_sale](
	[saleID] [int] IDENTITY(1,1) NOT NULL,
	[saleAccountCode] [varchar](50) NOT NULL,
	[saleModelNo] [varchar](50) NOT NULL,
	[saleSerialNo] [varchar](50) NULL,
	[saleOrderNo] [int] NULL,
	[saleOrderLine] [int] NULL,
	[saleQty] [smallint] NULL,
	[saleConnote] [varchar](20) NULL,
	[saleDealerCode] [varchar](9) NOT NULL,
	[saleLogisticsConfirmation] [smallint] NOT NULL,
	[saleLogisticsConfirmationDate] [datetime] NULL,
	[saleLogisticsConfirmationBy] [varchar](50) NULL,
	[saleDateCreated] [datetime] NOT NULL,
	[saleCreatedBy] [varchar](50) NOT NULL,
	[saleDateModified] [datetime] NULL,
	[saleModifiedBy] [varchar](50) NULL,
	[saleStatus] [smallint] NOT NULL,
	[saleComments] [varchar](500) NULL,
 CONSTRAINT [PK_tbl_loan_sale] PRIMARY KEY CLUSTERED 
(
	[saleID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

ALTER TABLE [dbo].[tbl_loan_sale] ADD  CONSTRAINT [DF_tbl_loan_sale_saleQty]  DEFAULT (1) FOR [saleQty]
GO

ALTER TABLE [dbo].[tbl_loan_sale] ADD  CONSTRAINT [DF_tbl_loan_sale_saleLogisticsConfirmation]  DEFAULT (0) FOR [saleLogisticsConfirmation]
GO

ALTER TABLE [dbo].[tbl_loan_sale] ADD  CONSTRAINT [DF_tbl_loan_sale_saleDateCreated]  DEFAULT (getdate()) FOR [saleDateCreated]
GO

ALTER TABLE [dbo].[tbl_loan_sale] ADD  CONSTRAINT [DF_tbl_loan_sale_saleStatus]  DEFAULT (1) FOR [saleStatus]
GO