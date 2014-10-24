/****** Object:  StoredProcedure [dbo].[spSendNotificationUpdate]    Script Date: 08/01/2014 13:31:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Procedure [dbo].[spSendNotificationUpdate] 
		 @EmpID Varchar(10)
		,@MonthP Varchar(2) 
		,@YearP Varchar(4)
As		
UPdate MonthlyBilling Set SendMailStatusID=Case When SendMailStatusID<3 Then SendMailStatusID+1 Else 3 End
	, SendMailDate=Case When SendMailStatusID=1 Then GETDATE() Else SendMailDate End
Where EmpID=@EmpID And MonthP=@MonthP And YearP=@YearP
GO
