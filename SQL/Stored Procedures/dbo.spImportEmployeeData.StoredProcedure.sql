/****** Object:  StoredProcedure [dbo].[spImportEmployeeData]    Script Date: 08/01/2014 13:31:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[spImportEmployeeData]
As

--Delete MsEmployee
--Insert AMER Employee
Insert MsEmployee(EmpID, EmpName, Post, EmpType, Agency, OfficeSection, WorkingTitle, EmailAddress, Status, CreateBy)
Select Distinct Convert(varchar(10),A.Amer_Emp_Key_Num)+'A' As EmpID, Amer_Emp_Last_TXT+', '+Amer_Emp_First_TXT As EmpName
	, ISNULL(F.fldPostDesc,'') , 'AMER', G.REF_AGNCY_ABBR_TXT, D.Ref_Post_Section_Abbr_txt as OfficeLocation, C.AMER_POS_WRKG_TITLE_TXT As WorkingTitle
	, E.AMER_EMP_ADDR_PERSONAL_EMAIL, B.AMER_EMP_POS_STATUS_IND, 'Admin'
From PassPS.dbo.Amer_Emp A
Inner Join PassPS.dbo.AMER_EMP_POS B on (A.Amer_Emp_Key_Num=B.Amer_Emp_Key_Num)
inner join PassPS.dbo.AMER_POS C on (B.Amer_Pos_Key_num=C.Amer_Pos_Key_num)
Inner join PassPS.dbo.REF_POST_SECTION D on (C.REF_POST_SECTION_KEY_NUM=D.REF_POST_SECTION_KEY_NUM)
Left Join PassPS.dbo.AMER_EMP_ADDR E on (A.AMER_EMP_KEY_NUM=E.AMER_EMP_KEY_NUM And REF_AMER_EMP_ADDR_KEY_NUM=13)
Left Join PASSISIS.dbo.tblPost F on (A.REF_POST_KEY_NUM=F.fldPostKey)
Left Join PASSPS.dbo.REF_AGNCY G on (C.REF_AGNCY_KEY_NUM=G.REF_AGNCY_KEY_NUM)
Where B.AMER_EMP_POS_STATUS_IND='C' And Convert(varchar(10),A.Amer_Emp_Key_Num)+'A' not in
(Select EmpID From MsEmployee)

--Insert Local Employee
Insert MsEmployee(EmpID, EmpName, Post, EmpType, Agency, OfficeSection, WorkingTitle, EmailAddress, Status, CreateBy)
Select Distinct Convert(varchar(10),A.FN_Emp_Key_Num)+'L' As EmpID, A.FN_EMP_LIST_NAME_TXT,ISNULL(F.fldPostDesc,'')
	, 'LES', G.REF_AGNCY_ABBR_TXT, D.Ref_Post_Section_Abbr_txt as OfficeLocation, C.FN_POS_WRKG_TITLE_TXT As WorkingTitle
	, FN_EMP_ADDR_PERSONAL_EMAIL, B.FN_EMP_POS_STATUS_IND, 'Admin'
From PassPS.dbo.FN_EMP A
Inner JOin PassPS.dbo.FN_EMP_POS B on (A.FN_Emp_Key_Num=B.FN_Emp_Key_Num)
inner join PassPS.dbo.FN_POS C on (B.FN_Pos_Key_num=C.FN_Pos_Key_num)
Inner join PassPS.dbo.REF_POST_SECTION D on (C.REF_POST_SECTION_KEY_NUM=D.REF_POST_SECTION_KEY_NUM)
Left Join PassPS.dbo.FN_EMP_ADDR E on (A.FN_Emp_Key_Num=E.FN_Emp_Key_Num And REF_FN_EMP_ADDR_KEY_NUM=5)
Left Join PASSISIS.dbo.tblPost F on (A.REF_POST_KEY_NUM=F.fldPostKey)
Left Join PASSPS.dbo.REF_AGNCY G on (C.REF_AGNCY_KEY_NUM=G.REF_AGNCY_KEY_NUM)
Where B.FN_EMP_POS_STATUS_IND='C' And Convert(varchar(10),A.FN_Emp_Key_Num)+'L' not in
(Select EmpID From MsEmployee)
GO
