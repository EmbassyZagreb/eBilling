/****** Object:  StoredProcedure [dbo].[spSendNotificationGroupUpdate]    Script Date: 12/02/2014 15:00:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
CREATE Procedure [dbo].[spSendNotificationGroupUpdate] 
		 @EmpID Varchar(10)
As		
UPdate MonthlyBilling Set SendMailStatusID=Case When SendMailStatusID<3 Then SendMailStatusID+1 Else 3 End
	, SendMailDate=Case When SendMailStatusID=1 Then GETDATE() Else SendMailDate End
Where EmpID=@EmpID
GO
