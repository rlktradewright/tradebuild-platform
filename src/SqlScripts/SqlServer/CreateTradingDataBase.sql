USE [Trading]
GO
/****** Object:  Table [dbo].[Exchange]    Script Date: 06/19/2014 22:53:45 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Exchange](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Name] [char](20) NOT NULL,
	[TimeZoneID] [int] NULL,
	[Notes] [text] NULL,
 CONSTRAINT [PK_Exchange] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY],
 CONSTRAINT [IX_Exchange_Name] UNIQUE CLUSTERED 
(
	[Name] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[BarData]    Script Date: 06/19/2014 22:53:45 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[BarData](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[PeriodDataID] [int] NOT NULL,
	[BarType] [tinyint] NOT NULL,
	[OpenPrice] [float] NOT NULL,
	[HighPrice] [float] NOT NULL,
	[LowPrice] [float] NOT NULL,
	[ClosePrice] [float] NOT NULL,
	[TickVolume] [int] NULL,
 CONSTRAINT [PK_BarData] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY],
 CONSTRAINT [IX_BarData] UNIQUE CLUSTERED 
(
	[PeriodDataID] ASC,
	[BarType] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[TimeZone]    Script Date: 06/19/2014 22:53:45 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[TimeZone](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[CanonicalId] [int] NULL,
	[Name] [varchar](255) NOT NULL,
 CONSTRAINT [PK_TimeZone] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[TickDataFormat]    Script Date: 06/19/2014 22:53:45 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[TickDataFormat](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Name] [varchar](255) NOT NULL,
 CONSTRAINT [PK_InstrumentData] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY],
 CONSTRAINT [IX_TickDataFormat_Name] UNIQUE CLUSTERED 
(
	[Name] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[TickData]    Script Date: 06/19/2014 22:53:45 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TickData](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[InstrumentId] [int] NOT NULL,
	[TickDataFormatId] [int] NOT NULL,
	[DateTime] [datetime] NOT NULL,
	[BasePrice] [float] NOT NULL,
	[TickSize] [float] NULL,
	[Data] [image] NULL,
 CONSTRAINT [PK_TickData] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY],
 CONSTRAINT [IX_TickData_InstrumentId_DateTime] UNIQUE CLUSTERED 
(
	[InstrumentId] ASC,
	[DateTime] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PeriodData]    Script Date: 06/19/2014 22:53:45 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PeriodData](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[InstrumentID] [int] NOT NULL,
	[BarLengthMinutes] [int] NOT NULL,
	[DateTime] [datetime] NOT NULL,
	[Volume] [int] NOT NULL,
	[OpenInterest] [int] NOT NULL,
 CONSTRAINT [PK_PeriodData] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY],
 CONSTRAINT [IX_PeriodData] UNIQUE CLUSTERED 
(
	[InstrumentID] ASC,
	[BarLengthMinutes] ASC,
	[DateTime] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[InstrumentLocalSymbol]    Script Date: 06/19/2014 22:53:45 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[InstrumentLocalSymbol](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[ProviderKey] [char](20) NOT NULL,
	[InstrumentID] [int] NOT NULL,
	[LocalSymbol] [char](20) NOT NULL,
 CONSTRAINT [PK_ContractLocalSymbol] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY],
 CONSTRAINT [IX_ContractLocalSymbol_ProviderContractID] UNIQUE NONCLUSTERED 
(
	[ProviderKey] ASC,
	[InstrumentID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[InstrumentClass]    Script Date: 06/19/2014 22:53:45 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[InstrumentClass](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[ExchangeID] [int] NULL,
	[Name] [varchar](100) NOT NULL,
	[TickSize] [decimal](18, 8) NOT NULL,
	[TickValue] [decimal](18, 8) NOT NULL,
	[Currency] [char](3) NOT NULL,
	[InstrumentCategoryID] [int] NOT NULL,
	[SessionStartTime] [datetime] NOT NULL,
	[SessionEndTime] [datetime] NOT NULL,
	[DaysBeforeExpiryToSwitch] [int] NULL,
	[Notes] [text] NULL,
 CONSTRAINT [PK_InstrumentClass] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY],
 CONSTRAINT [IX_InstrumentClassExchangeName] UNIQUE CLUSTERED 
(
	[ExchangeID] ASC,
	[Name] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[InstrumentCategory]    Script Date: 06/19/2014 22:53:45 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[InstrumentCategory](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Name] [char](10) NOT NULL,
 CONSTRAINT [PK_InstrumentCategory] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY],
 CONSTRAINT [IX_InstrumentCategory_Name] UNIQUE CLUSTERED 
(
	[Name] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Instrument]    Script Date: 06/19/2014 22:53:45 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Instrument](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[InstrumentClassID] [int] NOT NULL,
	[Name] [varchar](100) NOT NULL,
	[ShortName] [char](20) NOT NULL,
	[Symbol] [char](10) NOT NULL,
	[ExpiryDate] [datetime] NULL,
	[StrikePrice] [money] NULL,
	[OptRight] [char](1) NULL,
	[Notes] [text] NULL,
	[TickSize] [decimal](18, 8) NULL,
	[TickValue] [decimal](18, 8) NULL,
	[Currency] [char](3) NULL,
 CONSTRAINT [PK_Instrument] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY],
 CONSTRAINT [IX_Instrument_InstrumentClassId_name] UNIQUE NONCLUSTERED 
(
	[InstrumentClassID] ASC,
	[Name] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY],
 CONSTRAINT [IX_Instrument_ShortName] UNIQUE NONCLUSTERED 
(
	[InstrumentClassID] ASC,
	[ShortName] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY],
 CONSTRAINT [IX_Instrument_Specifier_Fields] UNIQUE NONCLUSTERED 
(
	[InstrumentClassID] ASC,
	[Symbol] ASC,
	[ExpiryDate] ASC,
	[StrikePrice] ASC,
	[OptRight] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  UserDefinedFunction [dbo].[HasTickData]    Script Date: 06/19/2014 22:53:46 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE FUNCTION [dbo].[HasTickData]
(
	-- Add the parameters for the function here
	@InstrumentId int
)
RETURNS bit
AS
BEGIN
	RETURN (select cast(1 as bit) where exists ( select * from tickdata where instrumentid=@InstrumentId))

END
GO
/****** Object:  UserDefinedFunction [dbo].[HasBarData]    Script Date: 06/19/2014 22:53:46 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE FUNCTION [dbo].[HasBarData]
(
	-- Add the parameters for the function here
	@InstrumentId int
)
RETURNS bit
AS
BEGIN
	RETURN (select cast(1 as bit) where exists ( select * from perioddata where instrumentid=@InstrumentId))

END
GO
/****** Object:  StoredProcedure [dbo].[FetchTickData]    Script Date: 11/01/2018 12:38:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[FetchTickData]
	(
		@InstrumentID int,
		@From datetime,
		@To datetime,
		@SessionStart time = '00:00:00',
		@SessionEnd time = '00:00:00'
	)
As
	set nocount on 
	SET DATEFIRST 1

	select [DateTime], 
			TickDataFormat.Name as Version,
			TickSize,
			BasePrice, 
			Data 
	from TickData inner join TickDataFormat
		on TickDataFormat.id = TickData.TickDataFormatId 
	where InstrumentID = @InstrumentID and
			[DateTime] >= @From and 
			[DateTime] < @To and
			CASE WHEN @SessionStart < @SessionEnd THEN
				CASE WHEN DATEPART(dw, [DateTime]) BETWEEN 1 AND 5 and 
						  CAST([DateTime] AS TIME) >= @SessionStart and 
						  CAST([DateTime] AS TIME) < @SessionEnd THEN
					1
				ELSE
					0
				END 
			WHEN @SessionStart > @SessionEnd THEN
				CASE WHEN (DATEPART(dw, [DateTime]) NOT BETWEEN 5 AND 6 and 
							 CAST([DateTime] AS TIME) >= @SessionStart) OR 
						  (DATEPART(dw, [DateTime]) BETWEEN 1 AND 5 and
							 CAST([DateTime] AS TIME) < @SessionEnd) THEN
					1
				ELSE
					0
				END
			ELSE
				1
			END = 1


	return 
GO
/****** Object:  StoredProcedure [dbo].[FetchBarData]    Script Date: 02/01/2018 15:51:21 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[FetchBarData]
	(
		@InstrumentID int,
		@BarType int,
		@BarLength smallint,
		@NumberRequired int = 500,
		@From datetime = '1900/01/01',
		@To datetime = '1900/01/01',
		@SessionStart time = '00:00:00',
		@SessionEnd time = '00:00:00',
		@StartAtFromDate int = 0,
		@Ascending int = 0
	)
As

	SET NOCOUNT ON
	SET DATEFIRST 1

	IF ISNULL(@To, '1900/01/01') = '1900/01/01'
	BEGIN
		SET @To = GETDATE()
	END

	select * from
		(select	top (@NumberRequired) PeriodData.[DateTime], 
			BarData.BarType,
			PeriodData.BarLengthMinutes,  
			BarData.OpenPrice as OpenPrice,
			BarData.HighPrice as HighPrice,
			BarData.LowPrice as LowPrice,
			BarData.ClosePrice as ClosePrice,
			case @BarType when 0 then PeriodData.Volume else 0 end as Volume,
			case when BarData.TickVolume IS NULL then 0 else BarData.TickVolume end as TickVolume,
			case @BarType when 0 then PeriodData.OpenInterest else 0 end as OpenInterest
		from PeriodData inner join BarData 
			on PeriodData.ID=BarData.PeriodDataID
		where PeriodData.InstrumentID=@InstrumentID and
			PeriodData.BarLengthMinutes=@BarLength and
			BarData.BarType=@BarType  and
			PeriodData.[DateTime] >= ISNULL(@From, '1900/01/01') and 
			PeriodData.[DateTime] < @To and
			CASE WHEN @SessionStart < @SessionEnd THEN
				CASE WHEN DATEPART(dw, PeriodData.[DateTime]) BETWEEN 1 AND 5 and 
						  CAST(PeriodData.[DateTime] AS TIME) >= @SessionStart and 
						  CAST(PeriodData.[DateTime] AS TIME) < @SessionEnd THEN
					1
				ELSE
					0
				END 
			WHEN @SessionStart > @SessionEnd THEN
				CASE WHEN (DATEPART(dw, PeriodData.[DateTime]) NOT BETWEEN 5 AND 6 and 
							 CAST(PeriodData.[DateTime] AS TIME) >= @SessionStart) OR 
						  (DATEPART(dw, PeriodData.[DateTime]) BETWEEN 1 AND 5 and
							 CAST(PeriodData.[DateTime] AS TIME) < @SessionEnd) THEN
					1
				ELSE
					0
				END
			ELSE
				1
			END = 1
		order by CASE WHEN @StartAtFromDate = 1 THEN PeriodData.[dateTime] END ASC,
				 CASE WHEN @StartAtFromDate = 0 THEN PeriodData.[dateTime] END DESC) as Bars
	order by CASE WHEN @Ascending = 1 THEN Bars.[dateTime] END ASC,
			 CASE WHEN @Ascending = 0 THEN Bars.[dateTime] END DESC


	return 
GO
/****** Object:  View [dbo].[vinstrumentclasses]    Script Date: 06/19/2014 22:53:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[vinstrumentclasses]
AS
SELECT     dbo.InstrumentClass.ID, dbo.InstrumentClass.Name, dbo.InstrumentCategory.Name AS Category, dbo.Exchange.ID AS ExchangeID, 
                      dbo.Exchange.Name AS Exchange, dbo.InstrumentClass.Currency, dbo.InstrumentClass.TickSize, dbo.InstrumentClass.TickValue, 
                      dbo.InstrumentClass.SessionStartTime, dbo.InstrumentClass.SessionEndTime, dbo.TimeZone.Name AS TimeZoneName, 
                      dbo.InstrumentClass.DaysBeforeExpiryToSwitch, dbo.InstrumentClass.Notes, dbo.InstrumentClass.InstrumentCategoryID
FROM         dbo.InstrumentCategory INNER JOIN
                      dbo.InstrumentClass ON dbo.InstrumentCategory.ID = dbo.InstrumentClass.InstrumentCategoryID INNER JOIN
                      dbo.Exchange ON dbo.Exchange.ID = dbo.InstrumentClass.ExchangeID LEFT OUTER JOIN
                      dbo.TimeZone ON dbo.Exchange.TimeZoneID = dbo.TimeZone.ID
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "Exchange"
            Begin Extent = 
               Top = 199
               Left = 103
               Bottom = 314
               Right = 255
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "InstrumentCategory"
            Begin Extent = 
               Top = 7
               Left = 625
               Bottom = 92
               Right = 777
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "InstrumentClass"
            Begin Extent = 
               Top = 11
               Left = 347
               Bottom = 126
               Right = 557
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "TimeZone"
            Begin Extent = 
               Top = 209
               Left = 481
               Bottom = 309
               Right = 633
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vinstrumentclasses'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vinstrumentclasses'
GO
/****** Object:  View [dbo].[vexchanges]    Script Date: 06/19/2014 22:53:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[vexchanges]
AS
SELECT     dbo.Exchange.ID, dbo.Exchange.Name, dbo.Exchange.Notes, dbo.Exchange.TimeZoneID, (CASE WHEN timezone.Name IS NULL 
                      THEN tz.Name ELSE timezone.Name END) AS TimeZoneName
FROM         dbo.Exchange LEFT OUTER JOIN
                      dbo.TimeZone AS tz ON dbo.Exchange.TimeZoneID = tz.ID LEFT OUTER JOIN
                      dbo.TimeZone ON dbo.TimeZone.ID = tz.CanonicalId
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "Exchange"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 121
               Right = 190
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tz"
            Begin Extent = 
               Top = 6
               Left = 228
               Bottom = 106
               Right = 380
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "TimeZone"
            Begin Extent = 
               Top = 6
               Left = 418
               Bottom = 106
               Right = 570
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vexchanges'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vexchanges'
GO
/****** Object:  StoredProcedure [dbo].[WriteTickData]    Script Date: 06/19/2014 22:53:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[WriteTickData]
	(
		@InstrumentID int,
		@DataVersion varchar(255),
		@DateAndTime datetime,
		@BasePrice float,
		@TickSize as float,
		@Data image
	)
As
	SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
	set  NOCOUNT on

	declare @TickDataFormatID as int
	declare @TickDataID as int
	
	begin transaction

	select @TickDataFormatID = ID 
	from TickDataFormat 
	where Name=@DataVersion

	if @TickDataFormatID is NULL 
		begin
		insert into TickDataFormat (Name)
			values (@DataVersion)
		set @TickDataFormatID=@@IDENTITY
		end
	
	set @TickDataID=(select ID from TickData with (UPDLOCK)
				where InstrumentID=@InstrumentID AND
					[DateTime]=@DateAndTime)

	if @TickDataID is null
		begin
		insert into TickData (InstrumentID,
				TickDataFormatID,
				[DateTime], 
				BasePrice,
				TickSize, 
				Data)
			values (@InstrumentID, 
				@TickDataFormatID,
				@DateAndTime, 
				@BasePrice,
				@TickSize,
				@Data)
		end
	else
		begin
		update TickData set 
				BasePrice=@BasePrice,
				TickSize = @TickSize,
				Data=@Data
			where ID=@TickDataID
		end
	commit transaction
	return
GO
/****** Object:  StoredProcedure [dbo].[WriteBarData]    Script Date: 06/19/2014 22:53:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[WriteBarData]
	(
		@InstrumentID int,
		@BarType int,
		@BarLength int,
		@DateAndTime datetime,
		@OpenPrice float,
		@HighPrice float,
		@LowPrice float,
		@ClosePrice float,
		@Volume int,
		@TickVolume int,
		@OpenInterest int
	)
As
	SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
	set nocount on

	declare @PeriodDataID as int
	declare @BarDataID as int

	begin transaction
	
	set @PeriodDataID=(select ID from PeriodData with (UPDLOCK) 
					where InstrumentID=@InstrumentID and
						BarLengthMinutes=@BarLength and
						[DateTime]=@DateAndTime)

	if @PeriodDataID is null 
		begin
		insert into PeriodData (InstrumentID, 
				BarlengthMinutes,
				[DateTime], 
				Volume,
				OpenInterest)
			values (@InstrumentID, 
				@BarLength,
				@DateAndTime,
				case @BarType when 0 then @Volume else 0 end,
				case @BarType when 0 then @OpenInterest else 0 end)
		set @PeriodDataID=@@IDENTITY
		end
	else
		begin
		update PeriodData set 
				Volume=case @BarType when 0 then @Volume else Volume end,
				OpenInterest=case @BarType when 0 then @OpenInterest else OpenInterest end
			where ID=@PeriodDataID
		end
	
	set @BarDataID=(select ID from BarData with (UPDLOCK)
				where PeriodDataID=@PeriodDataID AND
					BarType=@BarType)

	if @BarDataID is null
		begin
		insert into BarData (PeriodDataID,
				BarType, 
				OpenPrice,
				HighPrice,
				LowPrice,
				ClosePrice,
				TickVolume)
			values (@PeriodDataID, 
				@BarType, 
				@OpenPrice,
				@HighPrice,
				@LowPrice,
				@ClosePrice, 
				@TickVolume)
		end
	else
		begin
		update BarData set 
				OpenPrice=@OpenPrice,
				HighPrice=@HighPrice,
				LowPrice=@LowPrice,
				ClosePrice=@ClosePrice,
				TickVolume=@TickVolume
			where ID=@BarDataID
		end

	commit

	return
GO
/****** Object:  View [dbo].[vtimezones]    Script Date: 06/19/2014 22:53:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[vtimezones]
AS
SELECT     tz.ID, tz.Name, (CASE WHEN timezone.Name IS NULL THEN tz.Name ELSE timezone.Name END) AS CanonicalName, tz.CanonicalId
FROM         dbo.TimeZone AS tz LEFT OUTER JOIN
                      dbo.TimeZone ON dbo.TimeZone.ID = tz.CanonicalId
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "tz"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 106
               Right = 190
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "TimeZone"
            Begin Extent = 
               Top = 6
               Left = 228
               Bottom = 106
               Right = 380
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vtimezones'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vtimezones'
GO
/****** Object:  View [dbo].[vInstrumentDetails]    Script Date: 06/19/2014 22:53:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[vInstrumentDetails]
AS
SELECT     dbo.Instrument.ID, dbo.InstrumentClass.Name AS InstrumentClassName, dbo.Instrument.Name, dbo.Instrument.ShortName, dbo.Instrument.Symbol, 
                      dbo.InstrumentCategory.Name AS Category, dbo.Instrument.ExpiryDate, CONVERT(varchar(8), dbo.Instrument.ExpiryDate, 112) AS ExpiryMonth, 
                      dbo.Instrument.StrikePrice, dbo.Instrument.OptRight, dbo.Exchange.Name AS Exchange, (CASE WHEN NOT instrument.Currency IS NULL 
                      THEN instrument.Currency ELSE instrumentclass.Currency END) AS EffectiveCurrency, dbo.Instrument.Currency, 
                      (CASE WHEN NOT instrument.TickSize IS NULL THEN instrument.TickSize ELSE instrumentclass.TickSize END) AS EffectiveTickSize, 
                      dbo.Instrument.TickSize, (CASE WHEN NOT instrument.TickValue IS NULL THEN instrument.TickValue ELSE instrumentclass.TickValue END) 
                      AS EffectiveTickValue, dbo.Instrument.TickValue, dbo.InstrumentClass.SessionStartTime, dbo.InstrumentClass.SessionEndTime, 
                      (CASE WHEN TimeZone.CanonicalId IS NULL THEN dbo.TimeZone.Name ELSE tz.Name END) AS TimezoneName, 
                      dbo.InstrumentClass.DaysBeforeExpiryToSwitch, dbo.Instrument.InstrumentClassID, dbo.InstrumentClass.InstrumentCategoryID, 
                      dbo.HasBarData(dbo.Instrument.ID) AS HasBarData, dbo.HasTickData(dbo.Instrument.ID) AS HasTickData, dbo.Instrument.Notes
FROM         dbo.Exchange INNER JOIN
                      dbo.InstrumentCategory INNER JOIN
                      dbo.InstrumentClass ON dbo.InstrumentCategory.ID = dbo.InstrumentClass.InstrumentCategoryID ON 
                      dbo.Exchange.ID = dbo.InstrumentClass.ExchangeID INNER JOIN
                      dbo.Instrument ON dbo.InstrumentClass.ID = dbo.Instrument.InstrumentClassID LEFT OUTER JOIN
                      dbo.TimeZone ON dbo.Exchange.TimeZoneID = dbo.TimeZone.ID LEFT OUTER JOIN
                      dbo.TimeZone AS tz ON dbo.TimeZone.CanonicalId = tz.ID
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "Exchange"
            Begin Extent = 
               Top = 20
               Left = 579
               Bottom = 148
               Right = 731
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "InstrumentCategory"
            Begin Extent = 
               Top = 293
               Left = 508
               Bottom = 378
               Right = 698
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "InstrumentClass"
            Begin Extent = 
               Top = 11
               Left = 228
               Bottom = 280
               Right = 438
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "Instrument"
            Begin Extent = 
               Top = 82
               Left = 8
               Bottom = 329
               Right = 178
            End
            DisplayFlags = 280
            TopColumn = 8
         End
         Begin Table = "TimeZone"
            Begin Extent = 
               Top = 24
               Left = 824
               Bottom = 141
               Right = 976
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tz"
            Begin Extent = 
               Top = 6
               Left = 1014
               Bottom = 106
               Right = 1166
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 3' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vInstrumentDetails'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'030
         Alias = 900
         Table = 2565
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vInstrumentDetails'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vInstrumentDetails'
GO
/****** Object:  Default [DF_InstrumentClass_TickSize]    Script Date: 06/19/2014 22:53:45 ******/
ALTER TABLE [dbo].[InstrumentClass] ADD  CONSTRAINT [DF_InstrumentClass_TickSize]  DEFAULT ((0)) FOR [TickSize]
GO
/****** Object:  Default [DF_InstrumentClass_TickValue]    Script Date: 06/19/2014 22:53:45 ******/
ALTER TABLE [dbo].[InstrumentClass] ADD  CONSTRAINT [DF_InstrumentClass_TickValue]  DEFAULT ((0)) FOR [TickValue]
GO
/****** Object:  Default [DF_InstrumentClass_Currency]    Script Date: 06/19/2014 22:53:45 ******/
ALTER TABLE [dbo].[InstrumentClass] ADD  CONSTRAINT [DF_InstrumentClass_Currency]  DEFAULT ('') FOR [Currency]
GO
/****** Object:  Default [DF_InstrumentClass_CategoryID]    Script Date: 06/19/2014 22:53:45 ******/
ALTER TABLE [dbo].[InstrumentClass] ADD  CONSTRAINT [DF_InstrumentClass_CategoryID]  DEFAULT ((0)) FOR [InstrumentCategoryID]
GO
/****** Object:  Default [DF_InstrumentClass_SessionStartTime]    Script Date: 06/19/2014 22:53:45 ******/
ALTER TABLE [dbo].[InstrumentClass] ADD  CONSTRAINT [DF_InstrumentClass_SessionStartTime]  DEFAULT ('00:00') FOR [SessionStartTime]
GO
/****** Object:  Default [DF_InstrumentClass_SessionEndTime]    Script Date: 06/19/2014 22:53:45 ******/
ALTER TABLE [dbo].[InstrumentClass] ADD  CONSTRAINT [DF_InstrumentClass_SessionEndTime]  DEFAULT ('00:00') FOR [SessionEndTime]
GO
