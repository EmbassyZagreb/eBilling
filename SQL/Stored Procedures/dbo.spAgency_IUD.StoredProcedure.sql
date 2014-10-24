USE [dev-eBilling]
GO
/****** Object:  StoredProcedure [dbo].[spAgency_IUD]    Script Date: 08/01/2014 13:31:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[spAgency_IUD]
	(@Mode varchar(1),
	 @AgencyId Int=0,
	 @AgencyFundingCode varchar (10)=Null,
	 @AgencyDesc varchar (50)=Null,
	 @FiscalStripVAT varchar(200)=Null,
	 @FiscalStripNonVAT varchar(200)=Null,
	 @Disabled varchar(1)=Null
	)
AS 
If @Mode='I'
Begin
	INSERT INTO AgencyFunding(AgencyFundingCode, AgencyDesc, FiscalStripVAT, FiscalStripNonVAT, [Disabled])
	VALUES(@AgencyFundingCode, @AgencyDesc, @FiscalStripVAT, @FiscalStripNonVAT, @Disabled)
End
Else If @Mode='U'
Begin
	UPDATE AgencyFunding Set AgencyFundingCode=@AgencyFundingCode, AgencyDesc=@AgencyDesc, FiscalStripVAT=@FiscalStripVAT, FiscalStripNonVAT=@FiscalStripNonVAT, [Disabled]=@Disabled
	WHERE AgencyId=@AgencyId
End
Else If @Mode='D'
Begin
	DELETE AgencyFunding WHERE AgencyId=@AgencyId
End
GO
