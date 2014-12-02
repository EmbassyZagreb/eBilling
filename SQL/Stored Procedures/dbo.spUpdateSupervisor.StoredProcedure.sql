/****** Object:  StoredProcedure [dbo].[spUpdateSupervisor]    Script Date: 12/02/2014 15:00:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
CREATE procedure [dbo].[spUpdateSupervisor]
		@PrevSpvID varchar(10)=Null
		,@CurSpvID varchar(10)=Null
		,@UserName varchar(100)=Null
As
--Update Employee
Update MsEmployee Set SupervisorId=@CurSpvID, CreateBy=@UserName
Where SupervisorId=@PrevSpvID

--Update supervisor's email in MonthlyBilling table where status is pending/awaiting for supervisor apporval
Declare @PrevEmail varchar(100)
		,@CurEmail varchar(100)
		
		
Select @PrevEmail=EmailAddress from vwPhoneCustomerList Where EmpID=@PrevSpvID
Select @CurEmail=EmailAddress from vwPhoneCustomerList Where EmpID=@CurSpvID
Update MonthlyBilling Set SupervisorEmail=@CurEmail
Where SupervisorEmail=@PrevEmail And ProgressID < 4

--Update UpdateSupervisor table
Insert UpdateSupervisor(PrevSpvID, CurSpvID, CreatedBy)
Values (@PrevSpvID, @CurSpvID, @UserName)
GO
