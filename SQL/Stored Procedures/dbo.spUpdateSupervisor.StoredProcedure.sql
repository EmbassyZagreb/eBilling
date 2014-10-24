/****** Object:  StoredProcedure [dbo].[spUpdateSupervisor]    Script Date: 08/01/2014 13:31:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[spUpdateSupervisor]
		@PrevSpvID varchar(10)=Null
		,@CurSpvID varchar(10)=Null
		,@UserName varchar(100)=Null
As
--Update Employee
Update MsEmployee Set SupervisorId=@CurSpvID, CreateBy=@UserName
Where SupervisorId=@PrevSpvID

--Update Non Employee
Declare @PrevEmail varchar(50)
		,@CurEmail varchar(50)
		
/*		
Select @PrevEmail=EmailAddress from vwPhoneCustomerList Where EmpID=@PrevSpvID
Select @CurEmail=EmailAddress from vwPhoneCustomerList Where EmpID=@CurSpvID
Update MsNonEmployee Set Email=@CurEmail
Where Email=@PrevEmail
*/

Insert UpdateSupervisor(PrevSpvID, CurSpvID, CreatedBy)
Values (@PrevSpvID, @CurSpvID, @UserName)
GO
