/****** Object:  StoredProcedure [dbo].[spImportDataFromVIP]    Script Date: 07/28/2014 12:45:38 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		Name
-- Create date: 
-- Description:	
-- =============================================
CREATE PROCEDURE [dbo].[spImportDataFromVIP] 
	-- Add the parameters for the stored procedure here
	@p1 int = 0, 
	@p2 int = 0
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	SELECT @p1, @p2
END
GO
