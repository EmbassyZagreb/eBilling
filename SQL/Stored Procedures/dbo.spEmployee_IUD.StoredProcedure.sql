/****** Object:  StoredProcedure [dbo].[spEmployee_IUD]    Script Date: 08/01/2014 13:31:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[spEmployee_IUD] 
	(@Mode Varchar(1),
	 @EmpID varchar(10)=Null,
	 @EmpName varchar(50)=Null,
	 @AgencyID int=0,
	 @Post varchar(100)=Null,
     @EmpType varchar(10)=Null,
	 @Agency varchar(50)=Null,
     @OfficeSection varchar(50)=Null,
	 @WorkingTitle varchar(255)=Null,
	 @EmailAddress varchar(255)=Null,
     @AlternateEmail varchar(255)=Null,
	 @SupervisorId varchar(10)=Null,
	 @LoginID varchar(50)=Null,
	 @Remark varchar(100)=Null,
	 @Status varchar(1)=Null,
	 @UserName varchar(50)=Null
	)
AS
If @Mode='I'
Begin
	Select @EmpID=dbo.fnGetEmpID()
	Insert MsEmployee(EmpID, EmpName, AgencyID, Post, EmpType, Agency, OfficeSection, WorkingTitle, EmailAddress, AlternateEmail, SupervisorId, LoginID, Remark, [Status], CreateBy) 
	Values(@EmpID, @EmpName, @AgencyID, @Post, @EmpType, @Agency, @OfficeSection, @WorkingTitle, @EmailAddress, @AlternateEmail, @SupervisorId, @LoginID, @Remark, @Status, @UserName)
End
Else If @Mode='E'
Begin
	Update MsEmployee Set EmpName=@EmpName,
						  EmailAddress=@EmailAddress,
						  AgencyID=@AgencyID,
						  Post=@Post,
						  EmpType=@EmpType,
						  Agency=@Agency,
						  OfficeSection=@OfficeSection,
						  WorkingTitle=@WorkingTitle,
						  AlternateEmail=@AlternateEmail,
						  SupervisorId=@SupervisorId,
						  LoginID=@LoginID,
						  Remark = @Remark,
						  [Status]=@Status,
						  UpdateBy=@UserName, 
						  UpdateDate=Getdate()
	Where EmpID = @EmpID
End
Else If @Mode='D'
Begin
	Update MsEmployee Set Status='D' 
	Where EmpID = @EmpID
End
GO
