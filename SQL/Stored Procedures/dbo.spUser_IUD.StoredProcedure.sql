/****** Object:  StoredProcedure [dbo].[spUser_IUD]    Script Date: 12/02/2014 15:00:24 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
CREATE PROCEDURE [dbo].[spUser_IUD] 
	(@Mode Varchar(1),
	 @LoginID varchar(50)=Null,
	 @RoleID varchar(10)=Null)
AS
If @Mode='I'
Begin
	Insert Users(LoginID,RoleID) 
	Values(@LoginID,@RoleID)
End
Else If @Mode='E'
Begin
	Update Users Set RoleID = @RoleID Where LoginID = @LoginID
End
Else If @Mode='D'
Begin
	Delete Users Where LoginID=@LoginID
End
GO
