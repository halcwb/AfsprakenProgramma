USE [UMCU_WKZ_AP_Test]
GO
/****** Object:  UserDefinedFunction [dbo].[GetLastStandardPatientHospNum]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date, ,>
-- Description:	<Description, ,>
-- =============================================
CREATE FUNCTION [dbo].[GetLastStandardPatientHospNum] 
(
)
RETURNS NVARCHAR(50)
AS
BEGIN
	-- Declare the return variable here
	DECLARE @hospNum as NVARCHAR(50)

	-- Add the T-SQL statements to compute the return value here
	SET @hospNum = (SELECT TOP 1 p.HospitalNumber FROM GetStandardPatients() p ORDER BY p.HospitalNumber DESC)

	-- Return the result of the function
	RETURN @hospNum

END
GO
/****** Object:  UserDefinedFunction [dbo].[GetLatestConfigMedContVersionForDepartment]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date, ,>
-- Description:	<Description, ,>
-- =============================================
CREATE FUNCTION [dbo].[GetLatestConfigMedContVersionForDepartment] 
(
	-- Add the parameters for the function here
	@dep NVARCHAR(60)
)
RETURNS int
AS
BEGIN
	-- Declare the return variable here
	DECLARE @versionID int

	-- Add the T-SQL statements to compute the return value here
	SET @versionID = (
	SELECT TOP 1 [VersionID] FROM dbo.GetConfigMedContVersionsForDepartment(@dep)
	ORDER BY [VersionID] DESC )

	-- Return the result of the function
	RETURN @versionID

END
GO
/****** Object:  UserDefinedFunction [dbo].[GetLatestConfigMedDiscDoseVersion]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date, ,>
-- Description:	<Description, ,>
-- =============================================
CREATE FUNCTION [dbo].[GetLatestConfigMedDiscDoseVersion] 
(
)
RETURNS INT
AS
BEGIN
	-- Declare the return variable here
	DECLARE @versionID INT

	-- Add the T-SQL statements to compute the return value here
	SET @versionID = (
	SELECT TOP 1 [VersionID] FROM dbo.GetConfigMedDiscDoseVersions()
	ORDER BY [VersionID] DESC )

	-- Return the result of the function
	RETURN @versionID

END
GO
/****** Object:  UserDefinedFunction [dbo].[GetLatestConfigMedDiscSolutionVersion]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date, ,>
-- Description:	<Description, ,>
-- =============================================
CREATE FUNCTION [dbo].[GetLatestConfigMedDiscSolutionVersion] 
(
	-- Add the parameters for the function here
)
RETURNS INT
AS
BEGIN
	-- Declare the return variable here
	DECLARE @versionID INT

	-- Add the T-SQL statements to compute the return value here
	SET @versionID = (
	SELECT TOP 1 [VersionID] FROM dbo.GetConfigMedDiscSolutionVersions()
	ORDER BY [VersionID] DESC )

	-- Return the result of the function
	RETURN @versionID

END
GO
/****** Object:  UserDefinedFunction [dbo].[GetLatestConfigMedDiscSolutionVersionForDepartment]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date, ,>
-- Description:	<Description, ,>
-- =============================================
CREATE FUNCTION [dbo].[GetLatestConfigMedDiscSolutionVersionForDepartment] 
(
	-- Add the parameters for the function here
	@dep NVARCHAR(60)
)
RETURNS INT
AS
BEGIN
	-- Declare the return variable here
	DECLARE @versionID INT

	-- Add the T-SQL statements to compute the return value here
	SET @versionID = (
	SELECT TOP 1 [VersionID] FROM dbo.GetConfigMedDiscSolutionVersionsForDepartment(@dep)
	ORDER BY [VersionID] DESC )

	-- Return the result of the function
	RETURN @versionID

END
GO
/****** Object:  UserDefinedFunction [dbo].[GetLatestConfigMedDiscVersion]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date, ,>
-- Description:	<Description, ,>
-- =============================================
CREATE FUNCTION [dbo].[GetLatestConfigMedDiscVersion] 
(
)
RETURNS INT
AS
BEGIN
	-- Declare the return variable here
	DECLARE @versionID INT

	-- Add the T-SQL statements to compute the return value here
	SET @versionID = (
	SELECT TOP 1 [VersionID] FROM dbo.GetConfigMedDiscVersions()
	ORDER BY [VersionID] DESC )

	-- Return the result of the function
	RETURN @versionID

END
GO
/****** Object:  UserDefinedFunction [dbo].[GetLatestConfigParEntVersion]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date, ,>
-- Description:	<Description, ,>
-- =============================================
CREATE FUNCTION [dbo].[GetLatestConfigParEntVersion] 
(
)
RETURNS INT
AS
BEGIN
	-- Declare the return variable here
	DECLARE @versionID INT

	-- Add the T-SQL statements to compute the return value here
	SET @versionID = (
	SELECT TOP 1 [VersionID] FROM dbo.GetConfigParEntVersions()
	ORDER BY [VersionID] DESC )

	-- Return the result of the function
	RETURN @versionID

END
GO
/****** Object:  UserDefinedFunction [dbo].[GetLatestPrescriptionVersionForHospitalNumber]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date, ,>
-- Description:	<Description, ,>
-- =============================================
CREATE FUNCTION [dbo].[GetLatestPrescriptionVersionForHospitalNumber] 
(
	-- Add the parameters for the function here
	@hospitalNumber NVARCHAR(50)
)
RETURNS INT
AS
BEGIN
	-- Declare the return variable here
	DECLARE @versionID INT

	-- Add the T-SQL statements to compute the return value here
	SET @versionID = (
	SELECT TOP 1 [VersionID] FROM dbo.GetPrescriptionVersionsForHospitalNumber(@hospitalNumber)
	ORDER BY [VersionID] DESC )

	-- Return the result of the function
	RETURN @versionID

END
GO
/****** Object:  UserDefinedFunction [dbo].[GetPrescriberPIN]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date, ,>
-- Description:	<Description, ,>
-- =============================================
CREATE FUNCTION [dbo].[GetPrescriberPIN] 
(
	-- Add the parameters for the function here
	@prescr NVARCHAR(255)
)
RETURNS INT
AS
BEGIN
	-- Declare the return variable here
	DECLARE @PIN INT

	-- Add the T-SQL statements to compute the return value here
	SET @PIN = (SELECT p.PIN FROM Prescriber p WHERE p.Prescriber = @prescr)

	-- Return the result of the function
	RETURN @PIN

END
GO
/****** Object:  Table [dbo].[ConfigMedCont]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ConfigMedCont](
	[VersionID] [int] NOT NULL,
	[VersionUTC] [datetime] NOT NULL,
	[VersionDate] [datetime] NOT NULL,
	[Department] [nvarchar](60) NOT NULL,
	[Generic] [nvarchar](300) NOT NULL,
	[GenericUnit] [nvarchar](50) NULL,
	[GenericQuantity] [float] NULL,
	[GenericVolume] [float] NULL,
	[SolutionVolume] [float] NULL,
	[Solution_2_6_Quantity] [float] NULL,
	[Solution_2_6_Volume] [float] NULL,
	[Solution_6_11_Quantity] [float] NULL,
	[Solution_6_11_Volume] [float] NULL,
	[Solution_11_40_Quantity] [float] NULL,
	[Solution_11_40_Volume] [float] NULL,
	[Solution_40_Quantity] [float] NULL,
	[Solution_40_Volume] [float] NULL,
	[MinConcentration] [float] NULL,
	[MaxConcentration] [float] NULL,
	[Solution] [nvarchar](300) NULL,
	[SolutionRequired] [bit] NULL,
	[DripQuantity] [float] NULL,
	[DoseUnit] [nvarchar](50) NULL,
	[MinDose] [float] NULL,
	[MaxDose] [float] NULL,
	[AbsMaxDose] [float] NULL,
	[DoseAdvice] [nvarchar](max) NULL,
	[Product] [nvarchar](max) NULL,
	[ShelfLife] [float] NULL,
	[ShelfCondition] [nvarchar](50) NULL,
	[PreparationText] [nvarchar](max) NULL,
	[Signed] [bit] NULL,
	[DilutionText] [nvarchar](max) NULL,
 CONSTRAINT [PK_ConfigMedCont_1] PRIMARY KEY CLUSTERED 
(
	[VersionID] ASC,
	[Department] ASC,
	[Generic] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ConfigMedDisc]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ConfigMedDisc](
	[VersionID] [int] NOT NULL,
	[VersionUTC] [datetime] NOT NULL,
	[VersionDate] [datetime] NOT NULL,
	[GPK] [int] NOT NULL,
	[ATC] [nvarchar](50) NOT NULL,
	[MainGroup] [nvarchar](300) NULL,
	[SubGroup] [nvarchar](300) NULL,
	[Generic] [nvarchar](300) NOT NULL,
	[Product] [nvarchar](300) NOT NULL,
	[Label] [nvarchar](300) NOT NULL,
	[Shape] [nvarchar](150) NOT NULL,
	[Routes] [nvarchar](300) NOT NULL,
	[GenericQuantity] [float] NOT NULL,
	[GenericUnit] [nvarchar](50) NOT NULL,
	[MultipleQuantity] [float] NULL,
	[MultipleUnit] [nvarchar](50) NULL,
	[Indications] [nvarchar](max) NULL,
	[HasSolutions] [bit] NOT NULL,
	[IsActive] [bit] NOT NULL,
 CONSTRAINT [PK_ConfigMedDisc] PRIMARY KEY CLUSTERED 
(
	[VersionID] ASC,
	[GPK] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ConfigMedDiscDose]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ConfigMedDiscDose](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[VersionID] [int] NOT NULL,
	[VersionUTC] [datetime] NOT NULL,
	[VersionDate] [datetime] NOT NULL,
	[Department] [nvarchar](60) NULL,
	[Generic] [nvarchar](300) NOT NULL,
	[Shape] [nvarchar](150) NOT NULL,
	[Route] [nvarchar](60) NOT NULL,
	[Indication] [nvarchar](500) NOT NULL,
	[Gender] [nvarchar](50) NULL,
	[MinAge] [float] NULL,
	[MaxAge] [float] NULL,
	[MinWeight] [float] NULL,
	[MaxWeight] [float] NULL,
	[MinGestAge] [int] NULL,
	[MaxGestAge] [int] NULL,
	[Frequencies] [nvarchar](500) NULL,
	[DoseUnit] [nvarchar](50) NULL,
	[NormDose] [float] NULL,
	[MinDose] [float] NULL,
	[MaxDose] [float] NULL,
	[MaxPerDose] [float] NULL,
	[StartDose] [float] NULL,
	[IsDosePerKg] [bit] NULL,
	[IsDosePerM2] [bit] NULL,
	[AbsMaxDose] [float] NULL,
 CONSTRAINT [PK_ConfigMedDiscDose] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ConfigMedDiscSolution]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ConfigMedDiscSolution](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[VersionID] [int] NOT NULL,
	[VersionUTC] [datetime] NOT NULL,
	[VersionDate] [datetime] NOT NULL,
	[Department] [nvarchar](60) NOT NULL,
	[Generic] [nvarchar](300) NOT NULL,
	[Shape] [nvarchar](150) NOT NULL,
	[MinGenericQuantity] [float] NULL,
	[MaxGenericQuantity] [float] NULL,
	[Solutions] [nvarchar](500) NULL,
	[SolutionVolume] [float] NULL,
	[MinConc] [float] NULL,
	[MaxConc] [float] NULL,
	[MinInfusionTime] [int] NULL,
 CONSTRAINT [PK_ConfigMedDiscSolution] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ConfigMedTallMan]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ConfigMedTallMan](
	[Generic] [nvarchar](300) NOT NULL,
	[Tallman] [nvarchar](300) NOT NULL,
 CONSTRAINT [PK_ConfigMedTallMan] PRIMARY KEY CLUSTERED 
(
	[Generic] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ConfigParEnt]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ConfigParEnt](
	[VersionID] [int] NOT NULL,
	[VersionUTC] [datetime] NOT NULL,
	[VersionDate] [datetime] NOT NULL,
	[Name] [nvarchar](300) NOT NULL,
	[Energy] [float] NULL,
	[Protein] [float] NULL,
	[Carbohydrate] [float] NULL,
	[Lipid] [float] NULL,
	[Sodium] [float] NULL,
	[Potassium] [float] NULL,
	[Calcium] [float] NULL,
	[Phosphor] [float] NULL,
	[Magnesium] [float] NULL,
	[Iron] [float] NULL,
	[VitD] [float] NULL,
	[Chloride] [float] NULL,
	[Product] [nvarchar](max) NULL,
	[Signed] [bit] NOT NULL,
 CONSTRAINT [PK_ConfigParEnt_1] PRIMARY KEY CLUSTERED 
(
	[VersionID] ASC,
	[Name] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Log]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Log](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[Prescriber] [nvarchar](255) NOT NULL,
	[HospitalNumber] [nvarchar](50) NULL,
	[VersionID] [int] NOT NULL,
	[VersionUTC] [datetime] NOT NULL,
	[VersionDate] [datetime] NOT NULL,
	[Table] [nvarchar](500) NULL,
	[Text] [nvarchar](max) NOT NULL,
 CONSTRAINT [PK_Log_1] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Patient]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Patient](
	[HospitalNumber] [nvarchar](50) NOT NULL,
	[BirthDate] [date] NULL,
	[FirstName] [nvarchar](50) NULL,
	[LastName] [nvarchar](50) NULL,
	[Gender] [nchar](10) NULL,
	[GestWeeks] [int] NULL,
	[GestDays] [int] NULL,
	[BirthWeight] [float] NULL,
 CONSTRAINT [PK_Patient] PRIMARY KEY CLUSTERED 
(
	[HospitalNumber] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Prescriber]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Prescriber](
	[Prescriber] [nvarchar](255) NOT NULL,
	[LastName] [nvarchar](100) NULL,
	[FirstName] [nvarchar](100) NULL,
	[Role] [nchar](10) NULL,
	[PIN] [int] NULL,
 CONSTRAINT [PK_Prescriber] PRIMARY KEY CLUSTERED 
(
	[Prescriber] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PrescriptionData]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PrescriptionData](
	[HospitalNumber] [nvarchar](50) NOT NULL,
	[VersionID] [int] NOT NULL,
	[VersionUTC] [datetime] NOT NULL,
	[VersionDate] [datetime] NOT NULL,
	[Prescriber] [nvarchar](255) NOT NULL,
	[Signed] [bit] NOT NULL,
	[Parameter] [nvarchar](300) NOT NULL,
	[Data] [nvarchar](max) NOT NULL,
 CONSTRAINT [PK_PrescriptionData_1] PRIMARY KEY CLUSTERED 
(
	[HospitalNumber] ASC,
	[VersionID] ASC,
	[Parameter] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PrescriptionText]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PrescriptionText](
	[HospitalNumber] [nvarchar](50) NOT NULL,
	[VersionID] [int] NOT NULL,
	[VersionUTC] [datetime] NOT NULL,
	[VersionDate] [datetime] NOT NULL,
	[Prescriber] [nvarchar](255) NOT NULL,
	[Signed] [bit] NOT NULL,
	[Parameter] [nvarchar](300) NOT NULL,
	[Text] [nvarchar](max) NOT NULL,
 CONSTRAINT [PK_PrescriptionText_1] PRIMARY KEY CLUSTERED 
(
	[HospitalNumber] ASC,
	[VersionID] ASC,
	[Parameter] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  UserDefinedFunction [dbo].[GetConfigMedDiscDoseLatest]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetConfigMedDiscDoseLatest] 
(	
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT 
	      [Id]
		, [VersionID]
		, [VersionUTC]
		, [VersionDate]
		, [Generic]
		, [Shape]
		, [Route]
		, [Indication]
		, [Gender]
		, [MinAge]
		, [MaxAge]
		, [MinWeight]
		, [MaxWeight]
		, [MinGestAge]
		, [MaxGestAge]
		, [Frequencies]
		, [DoseUnit]
		, [NormDose]
		, [MinDose]
		, [MaxDose]
		, [MaxPerDose]
		, [StartDose]
		, [IsDosePerKg]
		, [IsDosePerM2]
		, [AbsMaxDose]
		FROM [dbo].[ConfigMedDiscDose] md
			  WHERE
				md.[VersionID] = [dbo].GetlatestConfigMedDiscDoseVersion()
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetConfigMedDiscLatest]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO





-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetConfigMedDiscLatest] 
(	
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
SELECT 
	md.[VersionID]
	, md.[VersionUTC]
	, md.[VersionDate]
	, [GPK]
	, [ATC]
	, [MainGroup]
	, [SubGroup]
	, md.[Generic]
	, tm.[Tallman] 
	, [Product]
	, [Label]
	, md.[Shape]
	, [Routes]
	, [GenericQuantity]
	, [GenericUnit]
	, [MultipleQuantity]
	, [MultipleUnit]
	, [Indications]
	, [HasSolutions]
	, [IsActive]
		
	, d.[Route]
	, d.Indication
	, d.Gender
	, d.MinAge
	, d.MaxAge
	, d.MinWeight
	, d.MaxWeight
	, d.MinGestAge
	, d.MaxGestAge
	, d.Frequencies
	, d.NormDose
	, d.MinDose
	, d.MaxDose
	, d.AbsMaxDose
	, d.MaxPerDose
	, d.IsDosePerKg
	, d.IsDosePerM2

FROM [dbo].[ConfigMedDisc] md
	LEFT JOIN [dbo].[GetConfigMedDiscDoseLatest]() d ON 
		d.[VersionID] = md.[VersionID]
		AND d.[Generic] = md.[Generic] 
		AND d.[Shape] = md.[Shape] 
		AND CHARINDEX(d.[Route], md.[Routes]) > 0
	LEFT JOIN [dbo].[ConfigMedTallMan] tm ON tm.[Generic] = REPLACE(REPLACE(md.[Generic], ' ', '-'), '/', '+')

WHERE
	md.[VersionID] = [dbo].GetlatestConfigMedDiscVersion()
)
GO
/****** Object:  View [dbo].[PatientPrescriberPrescriptionTexts]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[PatientPrescriberPrescriptionTexts]
AS
SELECT        dbo.PrescriptionText.HospitalNumber, dbo.PrescriptionText.Prescriber, dbo.Prescriber.LastName AS PrescriberLastName, dbo.Prescriber.FirstName AS PrescriberFirstName, dbo.Patient.BirthDate, 
                         dbo.Patient.FirstName AS PatientFirstName, dbo.Patient.LastName AS PatientLastName, dbo.Patient.Gender, dbo.Patient.GestWeeks, dbo.Patient.GestDays, dbo.Patient.BirthWeight, 
                         dbo.PrescriptionText.Parameter, dbo.PrescriptionText.Text, dbo.PrescriptionText.Signed, dbo.PrescriptionText.VersionID, dbo.PrescriptionText.VersionUTC, dbo.PrescriptionText.VersionDate
FROM            dbo.PrescriptionText INNER JOIN
                         dbo.Prescriber ON dbo.PrescriptionText.Prescriber = dbo.Prescriber.Prescriber INNER JOIN
                         dbo.Patient ON dbo.PrescriptionText.HospitalNumber = dbo.Patient.HospitalNumber
GO
/****** Object:  UserDefinedFunction [dbo].[GetPrescriptionTextLatest]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetPrescriptionTextLatest] 
(	
	-- Add the parameters for the function here
	  @hospitalNumber NVARCHAR(50)
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
SELECT [HospitalNumber]
      ,[VersionID]
      ,[VersionUTC]
      ,[VersionDate]
      ,[Prescriber]
      ,[PrescriberLastName]
      ,[PrescriberFirstName]
      ,[BirthDate]
      ,[PatientFirstName]
      ,[PatientLastName]
      ,[Gender]
      ,[GestWeeks]
      ,[GestDays]
      ,[BirthWeight]
      ,[Parameter]
      ,[Text]
      ,[Signed]
  FROM [dbo].[PatientPrescriberPrescriptionTexts] pt
	  WHERE
		pt.[HospitalNumber] = @hospitalNumber AND 
		pt.[VersionID] = dbo.GetLatestPrescriptionVersionForHospitalNumber(@hospitalNumber)
)
GO
/****** Object:  View [dbo].[PrescriptionDataPrescriberVersions]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[PrescriptionDataPrescriberVersions]
AS
SELECT        dbo.PrescriptionData.HospitalNumber, dbo.PrescriptionData.Prescriber, dbo.Prescriber.LastName, dbo.Prescriber.FirstName, dbo.PrescriptionData.VersionID, dbo.PrescriptionData.VersionUTC, 
                         dbo.PrescriptionData.VersionDate
FROM            dbo.PrescriptionData INNER JOIN
                         dbo.Prescriber ON dbo.PrescriptionData.Prescriber = dbo.Prescriber.Prescriber
GO
/****** Object:  UserDefinedFunction [dbo].[GetPrescriptionPrescribersForHospitalNumber]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetPrescriptionPrescribersForHospitalNumber]
(	
	-- Add the parameters for the function here
	@hospitalNumber NVARCHAR(50)
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT DISTINCT 
		  [VersionID]
		, [VersionUTC]
		, [VersionDate]
		, [Prescriber]
		, [LastName]
		, [FirstName]
	  FROM [dbo].[PrescriptionDataPrescriberVersions] pd
	  WHERE pd.[HospitalNumber] = @hospitalNumber
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetConfigMedContForDepartmentLatest]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetConfigMedContForDepartmentLatest] 
(	
	-- Add the parameters for the function here
	  @dep NVARCHAR(60)
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT [VersionID]
		  , [VersionUTC]
		  , [VersionDate]
		  , [Department]
		  , [Generic]
		  , [GenericUnit]
		  , [GenericQuantity]
		  , [GenericVolume]
		  , [SolutionVolume]
		  , [Solution_2_6_Quantity]
		  , [Solution_2_6_Volume]
		  , [Solution_6_11_Quantity]
		  , [Solution_6_11_Volume]
		  , [Solution_11_40_Quantity]
		  , [Solution_11_40_Volume]
		  , [Solution_40_Quantity]
		  , [Solution_40_Volume]
		  , [MinConcentration]
		  , [MaxConcentration]
		  , [Solution]
		  , [SolutionRequired]
		  , [DripQuantity]
		  , [DoseUnit]
		  , [MinDose]
		  , [MaxDose]
		  , [AbsMaxDose]
		  , [DoseAdvice]
		  , [Product]
		  , [ShelfLife]
		  , [ShelfCondition]
		  , [PreparationText]
		  , [Signed]
		  , [DilutionText]
	  FROM [dbo].[ConfigMedCont] cmc
	  WHERE
		cmc.[Department] = @dep AND 
		cmc.[VersionID] = dbo.GetLatestConfigMedContVersionForDepartment(@dep)
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetConfigMedContForDepartmentWithVersion]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetConfigMedContForDepartmentWithVersion] 
(	
	-- Add the parameters for the function here
	  @dep NVARCHAR(60)
	  , @versionID INT
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT [VersionID]
		  , [VersionUTC]
		  , [VersionDate]
		  , [Department]
		  , [Generic]
		  , [GenericUnit]
		  , [GenericQuantity]
		  , [GenericVolume]
		  , [SolutionVolume]
		  , [Solution_2_6_Quantity]
		  , [Solution_2_6_Volume]
		  , [Solution_6_11_Quantity]
		  , [Solution_6_11_Volume]
		  , [Solution_11_40_Quantity]
		  , [Solution_11_40_Volume]
		  , [Solution_40_Quantity]
		  , [Solution_40_Volume]
		  , [MinConcentration]
		  , [MaxConcentration]
		  , [Solution]
		  , [SolutionRequired]
		  , [DripQuantity]
		  , [DoseUnit]
		  , [MinDose]
		  , [MaxDose]
		  , [AbsMaxDose]
		  , [DoseAdvice]
		  , [Product]
		  , [ShelfLife]
		  , [ShelfCondition]
		  , [PreparationText]
		  , [Signed]
		  , [DilutionText]
	  FROM [dbo].[ConfigMedCont] cmc
	  WHERE
		cmc.[Department] = @dep AND 
		cmc.[VersionID] = @versionID
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetConfigMedContVersionsForDepartment]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetConfigMedContVersionsForDepartment]
(	
	-- Add the parameters for the function here
	@dep NVARCHAR(60)
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
SELECT DISTINCT [VersionID], [VersionUTC], [VersionDate]
  FROM [dbo].[ConfigMedCont] cmc
  WHERE cmc.Department = @dep)
GO
/****** Object:  UserDefinedFunction [dbo].[GetConfigMedDiscDoseForVersion]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO





-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetConfigMedDiscDoseForVersion] 
(
	@versionID INT	
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT  [Id]
		  , [VersionID]
		  , [VersionUTC]
		  , [VersionDate]
		  , [Generic]
		  , [Shape]
		  , [Route]
		  , [Indication]
		  , [Gender]
		  , [MinAge]
		  , [MaxAge]
		  , [MinWeight]
		  , [MaxWeight]
		  , [MinGestAge]
		  , [MaxGestAge]
		  , [Frequencies]
		  , [DoseUnit]
		  , [NormDose]
		  , [MinDose]
		  , [MaxDose]
		  , [MaxPerDose]
		  , [StartDose]
		  , [IsDosePerKg]
		  , [IsDosePerM2]
		  , [AbsMaxDose]
		  FROM [dbo].[ConfigMedDiscDose] md
			  WHERE
				md.[VersionID] = @versionID
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetConfigMedDiscDoseVersions]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO





-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetConfigMedDiscDoseVersions]
(	
	-- Add the parameters for the function here
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT DISTINCT [VersionID], [VersionUTC], [VersionDate]
	  FROM [dbo].ConfigMedDiscDose
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetConfigMedDiscForVersion]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO





-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetConfigMedDiscForVersion] 
(	
	@versionID INT
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
SELECT 
	md.[VersionID]
	, md.[VersionUTC]
	, md.[VersionDate]
	, [GPK]
	, [ATC]
	, [MainGroup]
	, [SubGroup]
	, md.[Generic]
	, tm.[Tallman] 
	, [Product]
	, [Label]
	, md.[Shape]
	, [Routes]
	, [GenericQuantity]
	, [GenericUnit]
	, [MultipleQuantity]
	, [MultipleUnit]
	, [Indications]
	, [HasSolutions]
	, [IsActive]
		
	, d.[Route]
	, d.Indication
	, d.Gender
	, d.MinAge
	, d.MaxAge
	, d.MinWeight
	, d.MaxWeight
	, d.MinGestAge
	, d.MaxGestAge
	, d.Frequencies
	, d.NormDose
	, d.MinDose
	, d.MaxDose
	, d.AbsMaxDose
	, d.MaxPerDose
	, d.IsDosePerKg
	, d.IsDosePerM2

FROM [dbo].[ConfigMedDisc] md
	LEFT JOIN [dbo].[ConfigMedDiscDose] d ON 
		d.[VersionID] = md.[VersionID]
		AND d.[Generic] = md.[Generic] 
		AND d.[Shape] = md.[Shape] 
		AND CHARINDEX(d.[Route], md.[Routes]) > 0
	LEFT JOIN [dbo].[ConfigMedTallMan] tm ON tm.[Generic] = REPLACE(REPLACE(md.[Generic], ' ', '-'), '/', '+')
WHERE
	md.[VersionID] = @versionID
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetConfigMedDiscSolutionForDepartmentLatest]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetConfigMedDiscSolutionForDepartmentLatest] 
(	
	-- Add the parameters for the function here
	  @dep NVARCHAR(60)
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT [Id]
		, [VersionID]
		, [VersionUTC]
		, [VersionDate]
		, [Department]
		, [Generic]
		, [Shape]
		, [Solutions]
		, [SolutionVolume]
		, [MinConc]
		, [MaxConc]
		, [MinInfusionTime]
	  FROM [dbo].[ConfigMedDiscSolution] mds
	  WHERE
		mds.[Department] = @dep AND 
		mds.[VersionID] = dbo.GetLatestConfigMedDiscSolutionVersionForDepartment(@dep)
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetConfigMedDiscSolutionForDepartmentWithVersion]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO





-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetConfigMedDiscSolutionForDepartmentWithVersion] 
(	
	-- Add the parameters for the function here
	  @versionID INT
	, @dep NVARCHAR(60)
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT 
	      [Id]
		, [VersionID]
		, [VersionUTC]
		, [VersionDate]
		, [Department]
		, [Generic]
		, [Shape]
		, [Solutions]
		, [SolutionVolume]
		, [MinConc]
		, [MaxConc]
		, [MinInfusionTime]
	  FROM [dbo].[ConfigMedDiscSolution] mds
	  WHERE
		mds.[Department] = @dep AND 
		mds.[VersionID] = @versionID
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetConfigMedDiscSolutionForGenericAndShapeLatest]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetConfigMedDiscSolutionForGenericAndShapeLatest] 
(	
	-- Add the parameters for the function here
	    @generic NVARCHAR(300)
	  , @shape NVARCHAR(150)
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT [Id]
		, [VersionID]
		, [VersionUTC]
		, [VersionDate]
		, [Department]
		, [Generic]
		, [Shape]
		, [MinGenericQuantity]
		, [MaxGenericQuantity]
		, [Solutions]	
		, [SolutionVolume]
		, [MinConc]
		, [MaxConc]
		, [MinInfusionTime]
	  FROM [dbo].[ConfigMedDiscSolution] mds
	  WHERE
		mds.[Generic] = @generic AND
		mds.[Shape] = @shape AND
		mds.[VersionID] = dbo.GetLatestConfigMedDiscSolutionVersion()
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetConfigMedDiscSolutionVersions]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO





-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetConfigMedDiscSolutionVersions]
(	
	-- Add the parameters for the function here
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT DISTINCT [VersionID], [VersionUTC], [VersionDate]
	  FROM [dbo].[ConfigMedDiscSolution] mds
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetConfigMedDiscSolutionVersionsForDepartment]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetConfigMedDiscSolutionVersionsForDepartment]
(	
	-- Add the parameters for the function here
	@dep NVARCHAR(60)
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT DISTINCT [VersionID], [VersionUTC], [VersionDate]
	  FROM [dbo].[ConfigMedDiscSolution] mds
	  WHERE mds.Department = @dep)
GO
/****** Object:  UserDefinedFunction [dbo].[GetConfigMedDiscVersions]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetConfigMedDiscVersions]
(	
	-- Add the parameters for the function here
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT DISTINCT [VersionID], [VersionUTC], [VersionDate]
	  FROM [dbo].ConfigMedDisc
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetConfigParEntForVersion]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetConfigParEntForVersion] 
(
	@versionID AS INT	
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT [VersionID]
		  , [VersionUTC]
		  , [VersionDate]
		  , [Name]
		  , [Energy]
		  , [Protein]
		  , [Carbohydrate]
		  , [Lipid]
		  , [Sodium]
		  , [Potassium]
		  , [Calcium]
		  , [Phosphor]
		  , [Magnesium]
		  , [Iron]
		  , [VitD]
		  , [Chloride]
		  , [Product]
		  , [Signed]
	  FROM [dbo].[ConfigParEnt] cpe
	  WHERE
		cpe.[VersionID] = @versionID
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetConfigParEntLatest]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetConfigParEntLatest] 
(	
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT [VersionID]
		  , [VersionUTC]
		  , [VersionDate]
		  , [Name]
		  , [Energy]
		  , [Protein]
		  , [Carbohydrate]
		  , [Lipid]
		  , [Sodium]
		  , [Potassium]
		  , [Calcium]
		  , [Phosphor]
		  , [Magnesium]
		  , [Iron]
		  , [VitD]
		  , [Chloride]
		  , [Product]
		  , [Signed]
	  FROM [dbo].[ConfigParEnt] cpe
	  WHERE
		cpe.[VersionID] = dbo.GetLatestConfigParEntVersion()
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetConfigParEntVersions]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetConfigParEntVersions]
(	
	-- Add the parameters for the function here
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT DISTINCT [VersionID], [VersionUTC], [VersionDate]
	  FROM [dbo].ConfigParEnt
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetPatients]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetPatients]
(	
	@hospitalNumber NVARCHAR(50)
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT [HospitalNumber]
		  ,[BirthDate]
		  ,[FirstName]
		  ,[LastName]
		  ,[Gender]
		  ,[GestWeeks]
		  ,[GestDays]
		  ,[BirthWeight]
	  FROM [dbo].[Patient] 
	  WHERE
		[HospitalNumber] = @hospitalNumber OR @hospitalNumber = ''
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetPrescribers]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetPrescribers]
(	
	@login NVARCHAR(50)
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT [Prescriber]
		  ,[LastName]
		  ,[FirstName]
		  ,[Role]
		  ,[PIN]
	  FROM [dbo].[Prescriber]
	  WHERE
		[Prescriber] = @login OR @login = ''
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetPrescriptionDataForVersion]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetPrescriptionDataForVersion] 
(	
	-- Add the parameters for the function here
	  @hospitalNumber NVARCHAR(50)
	  , @versionID INT
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT 
	      [HospitalNumber]
		, [VersionID]
		, [VersionUTC]
		, [VersionDate]
		, [Prescriber]
		, [Signed]
		, [Parameter]
		, [Data]
	  FROM [dbo].[PrescriptionData] pd
	  WHERE
		pd.[HospitalNumber] = @hospitalNumber AND 
		pd.[VersionID] = @versionID
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetPrescriptionDataLatest]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetPrescriptionDataLatest] 
(	
	-- Add the parameters for the function here
	  @hospitalNumber NVARCHAR(50)
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT [HospitalNumber]
		  ,[VersionID]
		  ,[VersionUTC]
		  ,[VersionDate]
		  ,[Prescriber]
		  ,[Signed]
		  ,[Parameter]
		  ,[Data]
	  FROM [dbo].[PrescriptionData] pd
	  WHERE
		pd.[HospitalNumber] = @hospitalNumber AND 
		pd.[VersionID] = dbo.GetLatestPrescriptionVersionForHospitalNumber(@hospitalNumber)
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetPrescriptionTextForVersion]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetPrescriptionTextForVersion] 
(	
	-- Add the parameters for the function here
	  @hospitalNumber NVARCHAR(50)
	, @versionID INT 
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT 
	      [HospitalNumber]
		, [VersionID]
		, [VersionUTC]
		, [VersionDate]
		, [Prescriber]
		, [Signed]
		, [Parameter]
		, [Text]
	  FROM [dbo].[PrescriptionText] pt
	  WHERE
		pt.[HospitalNumber] = @hospitalNumber AND 
		pt.[VersionID] = @versionID
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetPrescriptionVersionsForHospitalNumber]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetPrescriptionVersionsForHospitalNumber]
(	
	-- Add the parameters for the function here
	@hospitalNumber NVARCHAR(50)
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT DISTINCT [VersionID], [VersionUTC], [VersionDate]
	  FROM [dbo].[PrescriptionData] pd
	  WHERE pd.[HospitalNumber] = @hospitalNumber
)
GO
/****** Object:  UserDefinedFunction [dbo].[GetStandardPatients]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE FUNCTION [dbo].[GetStandardPatients]
(	
)
RETURNS TABLE 
AS
RETURN 
(
	-- Add the SELECT statement with parameter references here
	SELECT [HospitalNumber]
		  ,[BirthDate]
		  ,[FirstName]
		  ,[LastName]
		  ,[Gender]
		  ,[GestWeeks]
		  ,[GestDays]
		  ,[BirthWeight]
	  FROM [dbo].[Patient] 
	  WHERE
		[HospitalNumber] LIKE 'standaard_%'
)
GO
/****** Object:  View [dbo].[LogView]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[LogView]
AS
SELECT        dbo.[Log].Id, dbo.[Log].Prescriber, dbo.Prescriber.LastName, dbo.Prescriber.FirstName, dbo.[Log].HospitalNumber, dbo.Patient.FirstName AS PatientFirstname, dbo.Patient.LastName AS PatientLastName, 
                         dbo.Patient.BirthDate, dbo.[Log].Text, dbo.[Log].VersionID, dbo.[Log].VersionUTC, dbo.[Log].VersionDate, dbo.[Log].[Table]
FROM            dbo.[Log] LEFT OUTER JOIN
                         dbo.Prescriber ON dbo.[Log].Prescriber = dbo.Prescriber.Prescriber LEFT OUTER JOIN
                         dbo.Patient ON dbo.[Log].HospitalNumber = dbo.Patient.HospitalNumber
GO
ALTER TABLE [dbo].[PrescriptionData]  WITH CHECK ADD  CONSTRAINT [FK_PrescriptionData_HospitalNumber_Patient] FOREIGN KEY([HospitalNumber])
REFERENCES [dbo].[Patient] ([HospitalNumber])
GO
ALTER TABLE [dbo].[PrescriptionData] CHECK CONSTRAINT [FK_PrescriptionData_HospitalNumber_Patient]
GO
ALTER TABLE [dbo].[PrescriptionData]  WITH CHECK ADD  CONSTRAINT [FK_PrescriptionData_Prescriber_Prescriber] FOREIGN KEY([Prescriber])
REFERENCES [dbo].[Prescriber] ([Prescriber])
GO
ALTER TABLE [dbo].[PrescriptionData] CHECK CONSTRAINT [FK_PrescriptionData_Prescriber_Prescriber]
GO
ALTER TABLE [dbo].[PrescriptionText]  WITH CHECK ADD  CONSTRAINT [FK_PrescriptionText_HospitalNumber_Patient] FOREIGN KEY([HospitalNumber])
REFERENCES [dbo].[Patient] ([HospitalNumber])
GO
ALTER TABLE [dbo].[PrescriptionText] CHECK CONSTRAINT [FK_PrescriptionText_HospitalNumber_Patient]
GO
ALTER TABLE [dbo].[PrescriptionText]  WITH CHECK ADD  CONSTRAINT [FK_PrescriptionText_Prescriber_Prescriber] FOREIGN KEY([Prescriber])
REFERENCES [dbo].[Prescriber] ([Prescriber])
GO
ALTER TABLE [dbo].[PrescriptionText] CHECK CONSTRAINT [FK_PrescriptionText_Prescriber_Prescriber]
GO
/****** Object:  StoredProcedure [dbo].[AddForeignKey]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[AddForeignKey] 
	@table AS VARCHAR(50),
	@key AS VARCHAR(50),
	@references AS VARCHAR(50)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	DECLARE @sql VARCHAR(MAX)
	DECLARE @name VARCHAR(255)

	SET @name = 'FK_' + @table + '_' + @key + '_' + @references

    -- Insert statements for procedure here
	SET @sql = 'ALTER TABLE [dbo].[' + @table + ']  WITH CHECK ADD  CONSTRAINT [' + @name + '] FOREIGN KEY([' + @key + '])'
	SET @sql = @sql + ' ' + 'REFERENCES [dbo].[' + @references + '] ([' + @key + '])'

	SET @sql = @sql + ' ' + 'ALTER TABLE [dbo].[' + @table + '] CHECK CONSTRAINT [' + @name + ']'

	PRINT @sql

	EXEC (@sql)
END
GO
/****** Object:  StoredProcedure [dbo].[ClearDatabase]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[ClearDatabase] (@dbname VARCHAR(50), @delLog BIT)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	IF @delLog = 1
	BEGIN
	 EXEC dbo.TruncateDatabaseTable @dbname, 'Log'
	END

	EXEC dbo.TruncateDatabaseTable @dbname, 'ConfigMedTallMan'
	EXEC dbo.TruncateDatabaseTable @dbname, 'ConfigMedCont'
	EXEC dbo.TruncateDatabaseTable @dbname, 'ConfigMedDisc'
	EXEC dbo.TruncateDatabaseTable @dbname, 'ConfigMedDiscDose'
	EXEC dbo.TruncateDatabaseTable @dbname, 'ConfigMedDiscSolution'
	EXEC dbo.TruncateDatabaseTable @dbname, 'ConfigParEnt'
	EXEC dbo.TruncateDatabaseTable @dbname, 'PrescriptionData'
	EXEC dbo.TruncateDatabaseTable @dbname, 'PrescriptionText'

	EXEC dbo.DropForeignKey 'PrescriptionText', 'HospitalNumber', 'Patient'
	EXEC dbo.DropForeignKey 'PrescriptionData', 'HospitalNumber', 'Patient'

	EXEC dbo.TruncateDatabaseTable @dbname, 'Patient'

	EXEC dbo.AddForeignKey 'PrescriptionText', 'HospitalNumber', 'Patient'
	EXEC dbo.AddForeignKey 'PrescriptionData', 'HospitalNumber', 'Patient'

	EXEC dbo.DropForeignKey 'PrescriptionData', 'Prescriber', 'Prescriber'
	EXEC dbo.DropForeignKey 'PrescriptionText', 'Prescriber', 'Prescriber'
	
	EXEC dbo.TruncateDatabaseTable @dbname, 'Prescriber'

	EXEC dbo.AddForeignKey 'PrescriptionData', 'Prescriber', 'Prescriber'
	EXEC dbo.AddForeignKey 'PrescriptionText', 'Prescriber', 'Prescriber'

END
GO
/****** Object:  StoredProcedure [dbo].[DropForeignKey]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[DropForeignKey]
	@table AS VARCHAR(50),
	@key AS VARCHAR(50),
	@references AS VARCHAR(50)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here

	DECLARE @sql VARCHAR(MAX)
	DECLARE @name VARCHAR(255)

	SET @name = 'FK_' + @table + '_' + @key + '_' + @references


    -- Insert statements for procedure here
	SET @sql = 'ALTER TABLE [dbo].['+ @table +']'
	SET @sql = @sql + ' DROP CONSTRAINT [' + @name +']'

	PRINT @sql

	IF EXISTS (	SELECT
		* 
		FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS 
		WHERE CONSTRAINT_NAME = @name)

	EXEC (@sql)
END
GO
/****** Object:  StoredProcedure [dbo].[InsertConfigMedCont]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[InsertConfigMedCont] 
	-- Add the parameters for the stored procedure here
             @versionID INT
           , @versionUTC DATETIME
           , @versionDate DATETIME
           , @department NVARCHAR(60)
           , @generic NVARCHAR(300)
           , @genericUnit NVARCHAR(50)
           , @genericQuantity FLOAT
           , @genericVolume FLOAT
           , @solutionVolume FLOAT
           , @solution_2_6_Quantity FLOAT
           , @solution_2_6_Volume FLOAT
           , @solution_6_11_Quantity FLOAT
           , @solution_6_11_Volume FLOAT
           , @solution_11_40_Quantity FLOAT
           , @solution_11_40_Volume FLOAT
           , @solution_40_Quantity FLOAT
           , @solution_40_Volume FLOAT
           , @minConcentration FLOAT
           , @maxConcentration FLOAT
           , @solution NVARCHAR(300)
           , @solutionRequired BIT
           , @dripQuantity FLOAT
           , @doseUnit NVARCHAR(50)
           , @minDose FLOAT
           , @maxDose FLOAT
           , @absMaxDose FLOAT
           , @doseAdvice NVARCHAR(MAX)
           , @product NVARCHAR(MAX)
           , @shelfLife FLOAT
           , @shelfCondition NVARCHAR(50)
           , @preparationText NVARCHAR(MAX)
           , @signed BIT
		   , @dilutionText NVARCHAR(MAX)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
INSERT INTO [dbo].[ConfigMedCont]
           ([VersionID]
           ,[VersionUTC]
           ,[VersionDate]
           ,[Department]
           ,[Generic]
           ,[GenericUnit]
           ,[GenericQuantity]
           ,[GenericVolume]
           ,[SolutionVolume]
           ,[Solution_2_6_Quantity]
           ,[Solution_2_6_Volume]
           ,[Solution_6_11_Quantity]
           ,[Solution_6_11_Volume]
           ,[Solution_11_40_Quantity]
           ,[Solution_11_40_Volume]
           ,[Solution_40_Quantity]
           ,[Solution_40_Volume]
           ,[MinConcentration]
           ,[MaxConcentration]
           ,[Solution]
		   ,[SolutionRequired]
           ,[DripQuantity]
           ,[DoseUnit]
           ,[MinDose]
           ,[MaxDose]
           ,[AbsMaxDose]
           ,[DoseAdvice]
           ,[Product]
           ,[ShelfLife]
           ,[ShelfCondition]
           ,[PreparationText]
           ,[Signed]
		   ,[DilutionText])
     VALUES
           ( @versionID
		   , @versionUTC
		   , @versionDate 
           , @department 
           , @generic 
           , @genericUnit 
           , @genericQuantity 
           , @genericVolume 
           , @solutionVolume 
           , @solution_2_6_Quantity 
           , @solution_2_6_Volume 
           , @solution_6_11_Quantity 
           , @solution_6_11_Volume 
           , @solution_11_40_Quantity 
           , @solution_11_40_Volume 
           , @solution_40_Quantity 
           , @solution_40_Volume 
           , @minConcentration 
           , @maxConcentration 
           , @solution 
		   , @solutionRequired
           , @dripQuantity 
           , @doseUnit 
           , @minDose 
           , @maxDose 
           , @absMaxDose 
           , @doseAdvice 
           , @product 
           , @shelfLife 
           , @shelfCondition 
           , @preparationText 
           , @signed
		   , @dilutionText)
END
GO
/****** Object:  StoredProcedure [dbo].[InsertConfigMedDisc]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO





-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[InsertConfigMedDisc] 
	-- Add the parameters for the stored procedure here
             @versionID INT
           , @versionUTC DATETIME
           , @versionDate DATETIME
           , @GPK int
           , @ATC nvarchar(10)
           , @MainGroup nvarchar(300)
           , @SubGroup nvarchar(300)
           , @Generic nvarchar(300)
           , @Product nvarchar(300)
           , @Label nvarchar(300)
           , @Shape nvarchar(150)
           , @Routes nvarchar(300)
           , @GenericQuantity float
           , @GenericUnit nvarchar(50)
           , @MultipleQuantity float
           , @MultipleUnit nvarchar(50)
           , @Indications nvarchar(max)
           , @HasSolutions bit
           , @IsActive bit
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here

INSERT INTO [dbo].[ConfigMedDisc]
           ([VersionID]
           ,[VersionUTC]
           ,[VersionDate]
           ,[GPK]
           ,[ATC]
           ,[MainGroup]
           ,[SubGroup]
           ,[Generic]
           ,[Product]
           ,[Label]
           ,[Shape]
           ,[Routes]
           ,[GenericQuantity]
           ,[GenericUnit]
           ,[MultipleQuantity]
           ,[MultipleUnit]
           ,[Indications]
		   ,[HasSolutions]
           ,[IsActive])
     VALUES
           ( @VersionID
		   , @versionUTC
		   , @versionDate 
           , @GPK 
           , @ATC
           , @MainGroup 
           , @SubGroup 
           , @Generic 
           , @Product 
           , @Label 
           , @Shape 
           , @Routes 
           , @GenericQuantity 
           , @GenericUnit 
           , @MultipleQuantity 
           , @MultipleUnit 
           , @Indications
		   , @HasSolutions
           , @IsActive )
END
GO
/****** Object:  StoredProcedure [dbo].[InsertConfigMedDiscDose]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO





-- =============================================
-- Author:		<Author,,Name
-- Create date: <Create Date,,
-- Description:	<Description,,
-- =============================================
CREATE PROCEDURE [dbo].[InsertConfigMedDiscDose] 
	-- Add the parameters for the stored procedure here
             @VersionID int
		   , @VersionUTC datetime
		   , @VersionDate datetime
		   , @Department nvarchar(60)
           , @Generic nvarchar(300)
           , @Shape nvarchar(150)
           , @Route nvarchar(60)
           , @Indication nvarchar(500)
           , @Gender nvarchar(50)
           , @MinAge float
           , @MaxAge float
           , @MinWeight float
           , @MaxWeight float
           , @MinGestAge int
           , @MaxGestAge int
           , @Frequencies nvarchar(500)
           , @DoseUnit nvarchar(50)
           , @NormDose float
           , @MinDose float
           , @MaxDose float
           , @MaxPerDose float
           , @StartDose float
           , @IsDosePerKg bit
           , @IsDosePerM2 bit
           , @AbsMaxDose float
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here

INSERT INTO [dbo].[ConfigMedDiscDose]
           ([VersionID]
           ,[VersionUTC]
           ,[VersionDate]
		   ,[Department]
           ,[Generic]
           ,[Shape]
           ,[Route]
           ,[Indication]
           ,[Gender]
           ,[MinAge]
           ,[MaxAge]
           ,[MinWeight]
           ,[MaxWeight]
           ,[MinGestAge]
           ,[MaxGestAge]
           ,[Frequencies]
           ,[DoseUnit]
           ,[NormDose]
           ,[MinDose]
           ,[MaxDose]
           ,[MaxPerDose]
           ,[StartDose]
           ,[IsDosePerKg]
           ,[IsDosePerM2]
           ,[AbsMaxDose])
     VALUES
           ( @VersionID
		   , @VersionUTC
		   , @VersionDate 
		   , @Department
           , @Generic 
           , @Shape 
           , @Route 
           , @Indication 
           , @Gender 
           , @MinAge 
           , @MaxAge 
           , @MinWeight 
           , @MaxWeight 
           , @MinGestAge 
           , @MaxGestAge 
           , @Frequencies
           , @DoseUnit 
           , @NormDose 
           , @MinDose 
           , @MaxDose 
           , @MaxPerDose 
           , @StartDose 
           , @IsDosePerKg 
           , @IsDosePerM2 
           , @AbsMaxDose )
END
GO
/****** Object:  StoredProcedure [dbo].[InsertConfigMedDiscSolution]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO






-- =============================================
-- Author:		<Author,,Name
-- Create date: <Create Date,,
-- Description:	<Description,,
-- =============================================
CREATE PROCEDURE [dbo].[InsertConfigMedDiscSolution] 
	-- Add the parameters for the stored procedure here
             @VersionID int
		   , @VersionUTC datetime
		   , @VersionDate datetime
           , @Department nvarchar(60)
           , @Generic nvarchar(300)
           , @Shape nvarchar(150)
           , @MinGenericQuantity float
           , @MaxGenericQuantity float
           , @Solutions nvarchar(150)
           , @SolutionVolume float
           , @MinConc float
           , @MaxConc float
           , @MinInfusionTime int

AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here

INSERT INTO [dbo].[ConfigMedDiscSolution]
           ([VersionID]
           ,[VersionUTC]
           ,[VersionDate]
           ,[Department]
           ,[Generic]
           ,[Shape]
		   ,[MinGenericQuantity]
		   ,[MaxGenericQuantity]
           ,[Solutions]
           ,[SolutionVolume]
           ,[MinConc]
           ,[MaxConc]
           ,[MinInfusionTime])
     VALUES
           ( @VersionID
		   , @VersionUTC
		   , @VersionDate 
           , @Department 
           , @Generic 
           , @Shape 
		   , @MinGenericQuantity
		   , @MaxGenericQuantity
           , @Solutions
           , @SolutionVolume 
           , @MinConc 
           , @MaxConc 
           , @MinInfusionTime)

END
GO
/****** Object:  StoredProcedure [dbo].[InsertConfigParEnt]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[InsertConfigParEnt] 
	-- Add the parameters for the stored procedure here
      @versionID INT
	, @versionUTC DATETIME
	, @versionDate DATETIME
	, @name NVARCHAR(300)
	, @energy FLOAT
	, @protein FLOAT
	, @carbohydrate FLOAT
	, @lipid FLOAT
	, @sodium FLOAT
	, @potassium FLOAT
	, @calcium FLOAT
	, @phosphor FLOAT
	, @magnesium FLOAT
	, @iron FLOAT
	, @vitD FLOAT
	, @chloride FLOAT
	, @product NVARCHAR(MAX)
	, @signed BIT
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	INSERT INTO [dbo].[ConfigParEnt]
			([VersionID]
			,[VersionUTC]
			,[VersionDate]
			,[Name]
			,[Energy]
			,[Protein]
			,[Carbohydrate]
			,[Lipid]
			,[Sodium]
			,[Potassium]
			,[Calcium]
			,[Phosphor]
			,[Magnesium]
			,[Iron]
			,[VitD]
			,[Chloride]
			,[Product]
			,[Signed])
		VALUES
			( @versionID
			, @versionUTC
			, @versionDate 
			, @name
			, @energy
			, @protein 
			, @carbohydrate 
			, @lipid 
			, @sodium 
			, @potassium 
			, @calcium 
			, @phosphor 
			, @magnesium 
			, @iron 
			, @vitD 
			, @chloride 
			, @product
			, @signed )
END
GO
/****** Object:  StoredProcedure [dbo].[InsertLog]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[InsertLog]
	-- Add the parameters for the stored procedure here
		@prescriber NVARCHAR(255)
		, @hospitalNumber NVARCHAR(50)
		, @versionID INT
		, @versionUTC DATETIME
		, @versionDate DATETIME
		, @table nvarchar(500)
		, @text NVARCHAR(MAX)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	INSERT INTO [dbo].[Log]
			   ( [Prescriber]
			   , [HospitalNumber]
			   , [VersionID]
			   , [VersionUTC]
			   , [VersionDate]
			   , [Table]
			   , [Text])
		 VALUES
			   ( @prescriber
			   , @hospitalNumber
			   , @VersionID
			   , @versionUTC
			   , @versionDate 
			   , @table
			   , @text)
END
GO
/****** Object:  StoredProcedure [dbo].[InsertOrUpdateConfigMedTallMan]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE  [dbo].[InsertOrUpdateConfigMedTallMan] 
	-- Add the parameters for the stored procedure here
	  @generic AS NVARCHAR(300)
	, @tallMan AS NVARCHAR(300)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	IF EXISTS(SELECT * FROM ConfigMedTallMan tm WHERE tm.Generic = @generic)
		UPDATE ConfigMedTallMan
		SET Tallman = @tallMan
		WHERE Generic = @generic
	ELSE
		INSERT INTO ConfigMedTallMan (Generic, TallMan)
		VALUES (@generic, @tallMan)

END
GO
/****** Object:  StoredProcedure [dbo].[InsertPatient]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[InsertPatient]
	-- Add the parameters for the stored procedure here
             @hospitalNumber NVARCHAR(50)
           , @birthDate date
           , @firstName NVARCHAR(50)
           , @lastName NVARCHAR(50)
           , @gender NCHAR(10)
           , @gestWeeks int
           , @gestDays int
           , @birthWeight float
AS

BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here

	INSERT INTO [dbo].[Patient]
			   ([HospitalNumber]
			   ,[BirthDate]
			   ,[FirstName]
			   ,[LastName]
			   ,[Gender]
			   ,[GestWeeks]
			   ,[GestDays]
			   ,[BirthWeight])
		 VALUES
			   (@hospitalNumber
			   ,@birthDate
			   ,@firstName
			   ,@lastName
			   ,@gender
			   ,@gestWeeks
			   ,@gestDays
			   ,@birthWeight)
END
GO
/****** Object:  StoredProcedure [dbo].[InsertPrescriber]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[InsertPrescriber]
	-- Add the parameters for the stored procedure here
            @prescriber NVARCHAR(255)
           ,@lastName NVARCHAR(100)
           ,@firstName NVARCHAR(100)
           ,@role NCHAR(10)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
INSERT INTO [dbo].[Prescriber]
           ([Prescriber]
           ,[LastName]
           ,[FirstName]
           ,[Role])
     VALUES
            (@prescriber
            ,@lastName
            ,@firstName
            ,@role)
END
GO
/****** Object:  StoredProcedure [dbo].[InsertPrescriptionData]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[InsertPrescriptionData]
	-- Add the parameters for the stored procedure here
	  @hospitalNumber NVARCHAR(50)
	, @versionID INT
	, @versionUTC DATETIME
	, @versionDate DATETIME
	, @prescriber NVARCHAR(255)
	, @signed bit
	, @parameter NVARCHAR(300)
	, @data NVARCHAR(MAX)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	INSERT INTO [dbo].[PrescriptionData]
			   ( [HospitalNumber]
			   , [VersionID]
			   , [VersionUTC]
			   , [VersionDate]
			   , [Prescriber]
			   , [Signed]
			   , [Parameter]
			   , [Data])
		 VALUES
			   ( @hospitalNumber
			   , @versionID
			   , @versionUTC
			   , @versionDate 
			   , @prescriber
			   , @signed
			   , @parameter
			   , @data )
END
GO
/****** Object:  StoredProcedure [dbo].[InsertPrescriptionText]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[InsertPrescriptionText]
	-- Add the parameters for the stored procedure here
	  @hospitalNumber NVARCHAR(50)
	, @versionID INT
	, @versionUTC DATETIME
	, @versionDate DATETIME
	, @prescriber NVARCHAR(255)
	, @signed bit
	, @parameter NVARCHAR(300)
	, @text NVARCHAR(MAX)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
INSERT INTO [dbo].[PrescriptionText]
		( [HospitalNumber]
		, [VersionID]
		, [VersionUTC]
		, [VersionDate]
		, [Prescriber]
		, [Signed]
		, [Parameter]
		, [Text])
     VALUES
			   ( @hospitalNumber
			   , @versionID
			   , @versionUTC
			   , @versionDate 
			   , @prescriber
			   , @signed
			   , @parameter
			   , @text )
END
GO
/****** Object:  StoredProcedure [dbo].[TruncateDatabaseTable]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[TruncateDatabaseTable]  
	@dbname VARCHAR(50), @tablename VARCHAR(50) 
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	DECLARE @sql VARCHAR(200)

	SET @sql = 'USE ' + @dbname + ' TRUNCATE TABLE ' + 'dbo.[' + RTRIM(@tablename) + ']'
	PRINT 'running TruncateDataBaseTable'
	PRINT @sql

	EXECUTE (@sql)
END
GO
/****** Object:  StoredProcedure [dbo].[UpdatePatient]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[UpdatePatient] 
	-- Add the parameters for the stored procedure here
             @hospitalNumber NVARCHAR(50)
           , @birthDate DATE
           , @firstName NVARCHAR(50)
           , @lastName NVARCHAR(50)
           , @gender NCHAR(10)
           , @gestWeeks INT
           , @gestDays INT
           , @birthWeight FLOAT
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	UPDATE [dbo].[Patient] 
	   SET [BirthDate] = @birthDate
		  ,[FirstName] = @firstName
		  ,[LastName] = @lastName
		  ,[Gender] = @gender
		  ,[GestWeeks] = @gestWeeks
		  ,[GestDays] = @gestDays
		  ,[BirthWeight] = @birthWeight
	 WHERE [HospitalNumber] = @hospitalNumber
END
GO
/****** Object:  StoredProcedure [dbo].[UpdatePrescriber]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[UpdatePrescriber] 
	-- Add the parameters for the stored procedure here
            @prescriber NVARCHAR(255)
           ,@lastName NVARCHAR(100)
           ,@firstName NVARCHAR(100)
           ,@role NCHAR(10)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	UPDATE [dbo].[Prescriber]
	   SET [Prescriber] = @prescriber
		  ,[LastName] = @lastName
		  ,[FirstName] = @firstName
		  ,[Role] = @role
	 WHERE [Prescriber] = @prescriber
END
GO
/****** Object:  StoredProcedure [dbo].[UpdatePrescriberPIN]    Script Date: 17-10-2019 9:30:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[UpdatePrescriberPIN] 
	-- Add the parameters for the stored procedure here
            @prescriber NVARCHAR(255)
           ,@pin int
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	UPDATE [dbo].[Prescriber]
	   SET [Prescriber] = @prescriber
		  ,[PIN] = @pin
	 WHERE [Prescriber] = @prescriber
END
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[35] 4[32] 2[14] 3) )"
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
         Begin Table = "Log"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 236
               Right = 215
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "Prescriber"
            Begin Extent = 
               Top = 6
               Left = 253
               Bottom = 195
               Right = 423
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "Patient"
            Begin Extent = 
               Top = 6
               Left = 461
               Bottom = 278
               Right = 638
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
         Column = 2625
         Alias = 3030
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
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'LogView'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'LogView'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[41] 4[39] 2[2] 3) )"
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
         Begin Table = "PrescriptionText"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 175
               Right = 215
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "Prescriber"
            Begin Extent = 
               Top = 6
               Left = 253
               Bottom = 182
               Right = 423
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "Patient"
            Begin Extent = 
               Top = 6
               Left = 461
               Bottom = 221
               Right = 638
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
         Column = 2790
         Alias = 2820
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
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'PatientPrescriberPrescriptionTexts'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'PatientPrescriberPrescriptionTexts'
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
         Begin Table = "PrescriptionData"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 200
               Right = 215
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "Prescriber"
            Begin Extent = 
               Top = 6
               Left = 253
               Bottom = 178
               Right = 423
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
         Column = 4530
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
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'PrescriptionDataPrescriberVersions'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'PrescriptionDataPrescriberVersions'
GO
