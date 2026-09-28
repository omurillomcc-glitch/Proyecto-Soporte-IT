USE [SANTODOMINGO]
GO

/****** Object:  Table [dbo].[TN_INDICES_FRAGMENTADOS]    Script Date: 27/12/2019 00:32:42 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[TN_INDICES_FRAGMENTADOS](
	[FECHA] [date] NULL,
	[QUERY] [nvarchar](283) NULL,
	[FRAGMENTACION] [float] NULL,
	[BD] [sysname] NOT NULL,
	[INDICE] [sysname] NULL,
	[TABLA] [sysname] NOT NULL,
	[ORIGEN] [varchar](1) NULL
) ON [PRIMARY]
GO


