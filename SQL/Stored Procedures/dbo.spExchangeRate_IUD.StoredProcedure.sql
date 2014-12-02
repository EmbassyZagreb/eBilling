/****** Object:  StoredProcedure [dbo].[spExchangeRate_IUD]    Script Date: 12/02/2014 15:00:14 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
CREATE PROCEDURE [dbo].[spExchangeRate_IUD] 
	(@Mode Varchar(1),
	 @ExchangeID int=Null,
	 @ExchangeMonth varchar(2)=Null,
	 @ExchangeYear varchar(4)=Null,
	 @ExchangeRate float=Null)
AS
If @Mode='I'
Begin
	Insert ExchangeRate(ExchangeRate, ExchangeMonth, ExchangeYear) 
	Values(@ExchangeRate, @ExchangeMonth, @ExchangeYear)
End
Else If @Mode='E'
Begin
	Update ExchangeRate Set ExchangeMonth = @ExchangeMonth, ExchangeYear=@ExchangeYear
	Where ExchangeID = @ExchangeID
End
Else If @Mode='D'
Begin
	Delete ExchangeRate Where ExchangeID = @ExchangeID
End
GO
