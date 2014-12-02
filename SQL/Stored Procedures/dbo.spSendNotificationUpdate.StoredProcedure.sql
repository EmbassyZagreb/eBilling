/****** Object:  StoredProcedure [dbo].[spSendNotificationUpdate]    Script Date: 12/02/2014 15:00:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[spSendNotificationUpdate] 
		 @PhoneNumber Varchar(30)
		,@MonthP Varchar(2) 
		,@YearP Varchar(4)
As		
UPdate MonthlyBilling Set SendMailStatusID=Case When SendMailStatusID<3 Then SendMailStatusID+1 Else 3 End
	, SendMailDate=Case When SendMailStatusID=1 Then GETDATE() Else SendMailDate End
Where PhoneNumber=@PhoneNumber And MonthP=@MonthP And YearP=@YearP
GO
