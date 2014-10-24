/****** Object:  StoredProcedure [dbo].[spSendNotificationGroupUpdate]    Script Date: 08/01/2014 13:31:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[spSendNotificationGroupUpdate] 
		 @EmpID Varchar(10)
As		
UPdate MonthlyBilling Set SendMailStatusID=Case When SendMailStatusID<3 Then SendMailStatusID+1 Else 3 End
	, SendMailDate=Case When SendMailStatusID=1 Then GETDATE() Else SendMailDate End
Where EmpID=@EmpID
GO
