/****** Object:  StoredProcedure [dbo].[spNavigatorNotification]    Script Date: 12/02/2014 15:00:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[spNavigatorNotification]
	(@sPeriod varchar(6)=Null,
	 @ePeriod varchar(6)=Null,
	 @GraphHeight Int)
AS

Declare @MaxNrOfBills Int

SELECT @MaxNrOfBills = MAX(Y.Num ) From
(Select  Count(YearP+MonthP) As Num 
FROM         dbo.vwMonthlyBilling
WHERE YearP+MonthP>=@sPeriod and YearP+MonthP<=@ePeriod
Group By YearP+MonthP) As Y






SELECT  sum(case when SendMailStatusDesc = 'Not Sent' then 1 else 0 end) As NotSent,
		sum(case when SendMailStatusDesc = 'Sent' then 1 else 0 end) As Sent,
		sum(case when SendMailStatusDesc = 'ReSent' then 1 else 0 end) As ReSent,
		@GraphHeight*sum(case when SendMailStatusDesc = 'Not Sent' then 1 else 0 end)/ISNULL(@MaxNrOfBills,1) AS HeightNotSent, 
		@GraphHeight*sum(case when SendMailStatusDesc = 'Sent' then 1 else 0 end)/ISNULL(@MaxNrOfBills,1) AS HeightSent,
		@GraphHeight*sum(case when SendMailStatusDesc = 'ReSent' then 1 else 0 end)/ISNULL(@MaxNrOfBills,1) AS HeightReSent,	
		A.MonthP,
		A.YearP
FROM         dbo.vwMonthlyBilling A
WHERE A.YearP+A.MonthP>=@sPeriod and A.YearP+A.MonthP<=@ePeriod
Group By A.MonthP, A.YearP
ORDER BY A.YearP+A.MonthP
GO
