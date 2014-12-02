/****** Object:  StoredProcedure [dbo].[spLineNoUsageReport]    Script Date: 12/02/2014 15:00:18 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
CREATE Procedure [dbo].[spLineNoUsageReport] 
	@StartDate varchar(10)=Null
	,@EndDate varchar(10)=Null
	,@Post varchar(20)=Null
	,@Agency varchar(100)=Null
	,@Office varchar(20)=Null
	,@PhoneType varchar(20)=Null
	,@Orderby varchar(20)=Null
	,@Sortby varchar(5)=Null	
As
	Declare @Command Varchar(1000)

If @PhoneType='Office Phone'
Begin
	Set @Command='Select * from vwOfficePhoneNumberList Where PhoneNumber not in 
		(Select Distinct Extension From BillingDt Where Cost>0 And DialedDateTime>='''+@StartDate+''' And DialedDateTime<='''+@EndDate+''')  
		And (Post = '''+ @Post + ''' or '''+@Post+''' =''A'') 
		And (Agency = '''+ @Agency + ''' or '''+@Agency+''' =''A'') 
		And (Office = '''+ @Office + ''' or '''+@Office+''' =''A'')
		order by '+@Orderby+' '+@Sortby
End
Else If @PhoneType='Cell Phone'
Begin
	Set @Command='Select * from vwCellPhoneNumberList Where PhoneNumber not in 
		(Select Distinct PhoneNumber from CellPhoneDt Where Cost>0 And DialedDateTime>='''+@StartDate+''' And DialedDateTime<='''+@EndDate+''') 
		And (Post = '''+ @Post + ''' or '''+@Post+''' =''A'') 
		And (Agency = '''+ @Agency + ''' or '''+@Agency+''' =''A'') 
		And (Office = '''+ @Office + ''' or '''+@Office+''' =''A'') 
		order by '+@Orderby+' '+@Sortby
End
Else If @PhoneType='Home Phone'
Begin
	Set @Command='Select * from vwHomePhoneNumberList Where PhoneNumber not in 
		(Select Distinct PhoneNumber From HomePhoneDt Where Cost>0 And DialedDateTime>='''+@StartDate+''' And DialedDateTime<='''+@EndDate+''')  
		And (Post = '''+ @Post + ''' or '''+@Post+''' =''A'') 
		And (Agency = '''+ @Agency + ''' or '''+@Agency+''' =''A'') 
		And (Office = '''+ @Office + ''' or '''+@Office+''' =''A'') 
		order by '+@Orderby+' '+@Sortby
End
	Exec(@Command)

	print @Command
GO
