/****** Object:  View [dbo].[vwUnknownCellphoneBill]    Script Date: 08/01/2014 13:35:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE View [dbo].[vwUnknownCellphoneBill]
As
/*
Select     ID AS CellphoneID, MonthP, YearP, PhoneNumber, B.EmpName, B.Office, ISNULL(PreviousBalance, 0) AS PreviousBalance, ISNULL(Payment, 0) AS Payment, 
                      ISNULL(Adjustment, 0) AS Adjustment, ISNULL(BalanceDue, 0) AS BalanceDue, ISNULL(SubscriptionFee, 0) AS SubscriptionFee, ISNULL(LocalCall, 0) 
                      AS LocalCall, ISNULL(Interlocal, 0) AS Interlocal, ISNULL(IDD, 0) AS IDD, ISNULL(SMS, 0) AS SMS, ISNULL(MobileData, 0) AS MobileData, ISNULL(IRS, 
                      0) AS IRS, ISNULL(IRL, 0) AS IRL, ISNULL(Prepaid, 0) AS Prepaid, ISNULL(FARIDA, 0) AS FARIDA, ISNULL(MobileBanking, 0) AS MobileBanking, 
                      ISNULL(DetailedCallRecord, 0) AS DetailedCallRecord, ISNULL(Internet, 0) AS Internet, ISNULL(FLASHABO, 0) + ISNULL(BLBERRYABO, 0) 
                      AS DataRoam, ISNULL(MinUsage, 0) AS MinUsage, ISNULL(SubTotal, 0) AS SubTotal, ISNULL(PPN, 0) AS PPN, ISNULL(StampFee, 0) AS StampFee, 
                      ISNULL(CurrentBalance, 0) AS CurrentBalance, ISNULL(Total, 0) AS Total, ISNULL(MonthB, '') AS MonthB, ISNULL(YearB, '') AS YearB
From dbo.CellPhone A
Left Join vwPhoneCustomerList B on (A.PhoneNumber=B.MobilePhone)
Where PhoneNumber not in (Select PhoneNumber from MsCellPhoneNumber)
*/
Select ID AS CellphoneID, A.MonthP, A.YearP, A.PhoneNumber, B.EmpName, B.Office, isNULL(A.TOTBILLAMOUNT,0) As CurrentBalance
From CellPhoneHd A
Left Join vwPhoneCustomerList B on (A.PhoneNumber=B.MobilePhone)
Where PhoneNumber not in (Select PhoneNumber from MsCellPhoneNumber)
GO
