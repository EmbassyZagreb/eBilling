/****** Object:  StoredProcedure [dbo].[spLongDistanceCallsReport]    Script Date: 12/02/2014 15:00:19 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
CREATE Procedure [dbo].[spLongDistanceCallsReport] 
	@Type varchar(1)=Null
	,@StartDate varchar(10)=Null
	,@EndDate varchar(10)=Null
	,@Post varchar(20)=Null
	,@Agency varchar(100)=Null
	,@Office varchar(20)=Null
	,@EmpID varchar(10)=Null
	,@PhoneType varchar(20)=Null
	,@CallType varchar(10)=Null
	,@Orderby varchar(20)=Null
	,@Sortby varchar(5)=Null	
As
Declare @Command Varchar(1000)
If @Type='1'
Begin
	Set @Command='Select * From vwLongDistanceCall 
		Where DialedDateTime>='''+@StartDate+''' And DialedDateTime<='''+@EndDate+''' 
		And (Post = '''+ @Post + ''' or '''+@Post+''' =''A'')  
		And (Agency = '''+ @Agency + ''' or '''+@Agency+''' =''A'') 
		And (Office = '''+ @Office + ''' or '''+@Office+''' =''A'') 
		And (EmpID = '''+ @EmpID + ''' or '''+@EmpID+''' =''A'') 
		And (PhoneType = '''+ @PhoneType + ''' or '''+@PhoneType+''' =''A'') 
		And (isPersonal='''+@CallType+''' or '''+@CallType+'''=''A'' )
		order by '+@Orderby+' '+@Sortby
End
If @Type='2'
Begin
	Set @Command='Select isNull(sum(CallDurationSecond),0) As TotalDuration, isNull(sum(Cost),0) As TotalCost from vwLongDistanceCall 
		Where DialedDateTime>='''+@StartDate+''' And DialedDateTime<='''+@EndDate+''' 
		And (Post = '''+ @Post + ''' or '''+@Post+''' =''A'')  
		And (Agency = '''+ @Agency + ''' or '''+@Agency+''' =''A'') 
		And (Office = '''+ @Office + ''' or '''+@Office+''' =''A'') 
		And (EmpID = '''+ @EmpID + ''' or '''+@EmpID+''' =''A'') 
		And (PhoneType = '''+ @PhoneType + ''' or '''+@PhoneType+''' =''A'') 
		And (isPersonal='''+@CallType+''' or '''+@CallType+'''=''A'' )'
End
Exec(@Command)
--print @Command
GO
