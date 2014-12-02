/****** Object:  UserDefinedFunction [dbo].[fnGetEmpID]    Script Date: 12/02/2014 15:00:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
CREATE Function [dbo].[fnGetEmpID]()
Returns varchar(10)
As
Begin
	Declare @ID int
			,@NewEmpID varchar(10)

	Select @ID=isNull(Max(convert(int,LEFT(EmpID,LEN(EmpID)-1))),0) from MsEmployee

	Set @NewEmpID=Convert(varchar(10),@ID+1)+'W'

	Return @NewEmpID
End
GO
