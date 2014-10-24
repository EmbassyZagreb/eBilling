/****** Object:  StoredProcedure [dbo].[spCellPhoneNumber_IUD]    Script Date: 08/01/2014 13:31:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[spCellPhoneNumber_IUD] 
	(@Mode Varchar(1),
	 @ID Int=0,
	 @PhoneNumber varchar(30)=Null,
	 @PhoneType varchar(1)=Null,
	 @EmpId varchar(10)=Null,
	 @AlternateEmail varchar(100)=Null,
	 @Remark varchar(500)=Null,
	 @BillFlag varchar(1)=Null,
	 @Discontinued varchar(1)=Null,
	 @DiscontinuedDate varchar(20)=Null,
	 @OwnerID varchar(10)=Null,
	 @UserName varchar(50)=Null
	)
AS
Set @DiscontinuedDate=NullIf(@DiscontinuedDate,'')

If @Mode='I'
Begin
	Insert MsCellPhoneNumber(PhoneNumber, PhoneType, EmpID, AlternateEmail, Remark, BillFlag, Discontinued, DiscontinuedDate, OwnerID, CreateBy) 
	Values(@PhoneNumber,@PhoneType,@EmpId, @AlternateEmail, @Remark, @BillFlag, @Discontinued, @DiscontinuedDate, @OwnerID, @UserName)
End
Else If @Mode='E'
Begin
	Update MsCellPhoneNumber Set PhoneNumber = @PhoneNumber, PhoneType = @PhoneType, EmpId=@EmpId, AlternateEmail=@AlternateEmail, Remark=@Remark
		, BillFlag=@BillFlag, Discontinued=@Discontinued, DiscontinuedDate=@DiscontinuedDate, OwnerID=@OwnerID, UpdateBy=@UserName, UpdateDate=Getdate()
	Where [ID] = @ID
End
Else If @Mode='D'
Begin
	Delete MsCellPhoneNumber Where [ID] = @ID
End
GO
