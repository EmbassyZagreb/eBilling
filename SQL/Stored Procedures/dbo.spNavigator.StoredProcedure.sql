/****** Object:  StoredProcedure [dbo].[spNavigator]    Script Date: 12/02/2014 15:00:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[spNavigator] 
	(@EmpID Varchar(10)='',
	 @LoginID Varchar(50)='',
	 @MobilePhone Varchar(30),
	 @sPeriod varchar(6)=Null,
	 @ePeriod varchar(6)=Null,
	 @GraphHeight Int)
AS

Declare @MaxCellPhoneBillRp numeric(18,2),
		@MaxAccumulatedDebt numeric(18,2)

DECLARE @AccuTable TABLE(
Period varchar(6) NOT NULL,
AccumulatedDebt numeric(18,2) NOT NULL)

Select  @MaxCellPhoneBillRp = Max (CellPhoneBillRp) 
FROM         dbo.vwMonthlyBilling
WHERE (EmpID=@EmpID or @EmpID='') and (LoginID=@LoginID or @LoginID='') and MobilePhone=@MobilePhone and YearP+MonthP>=@sPeriod and YearP+MonthP<=@ePeriod

INSERT INTO @AccuTable
SELECT  YearP+MonthP,
		(ISNULL((SELECT     SUM(CellPhonePrsBillRp)
                FROM         dbo.vwMonthlyBilling AS b
                WHERE     (YearP + MonthP <= a.YearP + a.MonthP)
							AND (ProgressId = 4 OR ProgressId = 5) AND (MobilePhone = @MobilePhone) AND (LoginID=@LoginID or @LoginID='') AND (EmpID=@EmpID or @EmpID='')), 0))
FROM         dbo.vwMonthlyBilling a
WHERE (EmpID=@EmpID or @EmpID='') and (LoginID=@LoginID or @LoginID='') and MobilePhone=@MobilePhone and YearP+MonthP>=@sPeriod and YearP+MonthP<=@ePeriod
ORDER BY YearP+MonthP


Select  @MaxAccumulatedDebt = Max (AccumulatedDebt) 
FROM         @AccuTable


SELECT  ISNULL((CASE WHEN A.CellPhoneBillRp - A.CellPhonePrsBillRp > 0 THEN A.CellPhoneBillRp - A.CellPhonePrsBillRp ELSE 0 END),0),
		ISNULL((CASE WHEN A.CellPhoneBillRp - A.CellPhonePrsBillRp > 0 THEN A.CellPhonePrsBillRp ELSE A.CellPhoneBillRp END),0),
		ISNULL(B.AccumulatedDebt, 0) AS AccumulatedDebt,
		@GraphHeight*ISNULL((CASE WHEN A.CellPhoneBillRp - A.CellPhonePrsBillRp > 0 THEN A.CellPhoneBillRp - A.CellPhonePrsBillRp ELSE 0 END),0)/ISNULL(NULLIF(@MaxCellPhoneBillRp,0),1) AS HeightOfficial, 
		@GraphHeight*ISNULL((CASE WHEN A.CellPhoneBillRp - A.CellPhonePrsBillRp > 0 THEN A.CellPhonePrsBillRp ELSE A.CellPhoneBillRp END),0)/ISNULL(NULLIF(@MaxCellPhoneBillRp,0),1) AS HeightPersonal,
		@GraphHeight*ISNULL(B.AccumulatedDebt, 0)/ISNULL(NULLIF(@MaxAccumulatedDebt,0),1) AS HeightAccumulatedDebt,		
		A.MonthP,
		A.YearP,
		A.ProgressId,
		A.EmpName,
		ISNULL(A.CellPhoneBillRp, 0),
		A.ProgressDesc,
		A.AgencyFundingDesc,
		A.EmailAddress,
		A.SupervisorEmail,
		A.Notes,
		A.SupervisorRemark,
		A.Office,
		A.EmpID,
		A.ProgressId,
		A.FiscalStripNonVAT
FROM         dbo.vwMonthlyBilling A
Inner Join @AccuTable B on (A.YearP+A.MonthP = B.Period)
WHERE (A.EmpID=@EmpID or @EmpID='') and (A.LoginID=@LoginID or @LoginID='') and A.MobilePhone=@MobilePhone and A.YearP+A.MonthP>=@sPeriod and A.YearP+A.MonthP<=@ePeriod
ORDER BY A.YearP+A.MonthP
GO
