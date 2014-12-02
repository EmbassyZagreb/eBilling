/****** Object:  View [dbo].[vwShuttleBus]    Script Date: 12/02/2014 15:00:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE View [dbo].[vwShuttleBus]
As
/*
Select A.ShuttleID, A.MonthP, A.YearP, A.EmpID, B.LastName+' '+B.FirstName AS EmpName, B.Agency, B.Office
	, A.TransportDate, A.EventType, A.QtyPerson 
From ShuttleBill A
Inner Join vwPhoneCustomerList B on (A.EmpID=B.EmpID)
*/

Select A.ShuttleID, A.MonthP, A.YearP, A.EmpID, B.EmpName, B.Agency, B.Office
From ShuttleBill A
Inner Join vwPhoneCustomerList B on (A.EmpID=B.EmpID)
GO
