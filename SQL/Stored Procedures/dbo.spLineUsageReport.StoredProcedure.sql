/****** Object:  StoredProcedure [dbo].[spLineUsageReport]    Script Date: 12/02/2014 15:00:18 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
CREATE Procedure [dbo].[spLineUsageReport] 
	@StartDate varchar(10)=Null
	,@EndDate varchar(10)=Null
	,@Post varchar(20)=Null
	,@Agency varchar(100)=Null
	,@Office varchar(20)=Null
	,@EmpID varchar(10)=Null
	,@PhoneType varchar(20)=Null
	,@Orderby varchar(20)=Null
	,@Sortby varchar(5)=Null	
As
	Declare @Command Varchar(1000)

	Set @Command='Select PhoneType, PhoneNumber, Post, EmpID, EmpName, Agency, Office 
			, Sum(isNull(CallDurationSecond,0)) As CallDurationSecond, Sum(isNull(Cost,0)) As Cost
		      From vwLineUsages 
		Where DialedDateTime>='''+@StartDate+''' And DialedDateTime<='''+@EndDate+''' 
		And (Post = '''+ @Post + ''' or '''+@Post+''' =''A'')  
		And (Agency = '''+ @Agency + ''' or '''+@Agency+''' =''A'') 
		And (Office = '''+ @Office + ''' or '''+@Office+''' =''A'') 
		And (EmpID = '''+ @EmpID + ''' or '''+@EmpID+''' =''A'') 
		And (PhoneType = '''+ @PhoneType + ''' or '''+@PhoneType+''' =''A'') 
		Group By PhoneType, PhoneNumber, Post, EmpID, EmpName, Agency, Office
		order by '+@Orderby+' '+@Sortby

	Exec(@Command)
	print @Command
GO
