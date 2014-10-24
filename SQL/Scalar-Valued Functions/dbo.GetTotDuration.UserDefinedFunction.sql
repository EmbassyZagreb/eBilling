/****** Object:  UserDefinedFunction [dbo].[GetTotDuration]    Script Date: 08/01/2014 13:34:10 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Function [dbo].[GetTotDuration]
(
	@Time varchar(20)
)Returns Int
As
Begin
	Declare @Hour int
		,@Minute int
		,@Second int
		,@posX tinyint
		,@lenX tinyint
		,@TotDuration Int
	Set @TotDuration=0
	Set @posX = CharIndex(':',@Time)
	Set @lenX = Len(@Time)
	If @posX >1 
	Begin
		Set @Hour=Substring(@Time,1,@posX-1)*3600
--		print @Hour
		Set @Time = Substring(@Time,@posX+1,@lenX-@posX)
		Set @posX = CharIndex(':',@Time)
--		print @Time
		Set @Minute=Substring(@Time,1,@posX-1)*60
--		print @Minute
		Set @Time = Substring(@Time,@posX+1,@lenX-@posX)
		--print @Time
		Set @Second=Substring(@Time,1,@posX-1)
--		print @Second
		Set @TotDuration=@Hour+@Minute+@Second
	End 
	Else
	Begin
		Set @TotDuration=@Time
	End
	Return @TotDuration
End
GO
