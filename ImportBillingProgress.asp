<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="upload.asp" -->
<!--#include file="connect.inc" -->
<!-- #include file = "config.asp" -->

<%
On Error Resume Next


'------------------------------------------------

Dim strFolder, bolUpload, strMessage, FileExcel
Dim httpref, lngFileSize
Dim strIncludes, strExcludes, ExcelCon, conStr, rsData, strInsert, StrValue, strExecute
Dim strsql
Dim MonthP_
Dim YearP_


%>
<%

' Create the FileUploader
Dim Uploader, File
Set Uploader = New FileUploader

' This starts the upload process
Uploader.Upload()

MonthP_ = Uploader.Form("MonthList")
YearP_ = Uploader.Form("YearList")

'response.write MonthP_
'response.write YearP_

'******************************************
' Use [FileUploader object].Form to access 
' additional form variables submitted with
' the file upload(s). (used below)
'******************************************

' Check if any files were uploaded
If Uploader.Files.Count = 0 Then
	strMessage = "No file entered."
Else
	' Loop through the uploaded files
	For Each File In Uploader.Files.Items		

		bolUpload = false		

		'Response.Write lngMaxSize
		'Response.End 

		if lngFileSize = 0 then
			bolUpload = true
		else		
			if File.FileSize > lngFileSize then
				bolUpload = false
				strMessage = "File too large"
			else
				bolUpload = true
			end if
		end if

		if bolUpload = true then				
		    'Check to see if file extensions are excluded
		    If strExcludes <> "" Then
				If ValidFileExtension(File.FileName, strExcludes) Then
		            strMessage = "It is not allowed to upload a file containing a [." & GetFileExtension(File.FileName) & "] extension"
					bolUpload = false
				End If
			End If
			'Check to see if file extensions are included
			If strIncludes <> "" Then
				If InValidFileExtension(File.FileName, strIncludes) Then
					strMessage = "It is not allowed to upload a file containing a [." & GetFileExtension(File.FileName) & "] extension"
					bolUpload = false
				End If
			End If			
		end if
		
'		response.write strFolder

		if bolUpload = true then
'			File.SaveToDisk strFolder ' Save the file			
'			strMessage =  "File Uploaded: " & File.FileName
'			strMessage =  "Upload completed !"

'			strsql = "Exec spUpdateFileName " & HSId & "," & VId & ",'" & PicType  & "','" & File.FileName & "'"
			'Response.Write 	strsql
'			SDTCon.execute strsql


			'strMessage = strMessage & "Size: " & File.FileSize & " bytes<br>"
			'strMessage = strMessage & "Type: " & File.ContentType & "<br><br>"			
		end if

		FileExcel =  strFolder + "\"+File.FileName
		'response.write FileExcel
		
		
			conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FileExcel & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"
			Set ExcelCon = CreateObject("ADODB.Connection")
			'response.write strCon 
			ExcelCon.Open conStr

			set rsData = server.createobject("adodb.recordset") 
			set rsData= ExcelCon.execute("select top 1 * from [CellPhoneHd$]") 
			strInsert="Insert Cellphone(MonthP, YearP, PhoneNumber, PreviousBalance, Payment, Adjustment, BalanceDue, SubscriptionFee, LocalCall, Interlocal, IDD, SMS, MobileData, SMSBanking, SMSINTER, ROAMGPRS, IRS, IRL, Prepaid, FARIDA, MobileBanking, DetailedCallRecord, Internet, GPRS, FLASHABO, MMS, BLBERRYABO, MinUsage, SubTotal, PPN, StampFee, CurrentBalance, Total) "
			do while not rsData.eof  
	'			StrValue = "Values('" & rsData("MonthP") & "','" & rsData("YearP") & "','" & rsData("PhoneNumber") & "'," & rsData("PreviousBalance") & "," & rsData("Payment") & "," & rsData("Adjustment") & "," & rsData("BalanceDue") & "," & rsData("SubscriptionFee") & "," & rsData("LocalCall") & "," & rsData("Interlocal") & "," & rsData("IDD") & "," & rsData("SMS") & "," & rsData("MobileData") & "," & rsData("SMSBanking") & "," & rsData("SMSINTER") & "," & rsData("ROAMGPRS") & "," & rsData("IRS") & "," & rsData("IRL") & "," & rsData("Prepaid") & "," & rsData("FARIDA") & "," & rsData("MobileBanking") & "," & rsData("DetailedCallRecord") & "," & rsData("Internet") & "," & rsData("GPRS") & "," & rsData("FLASHABO") & "," & rsData("MMS") & "," & rsData("BLBERRYABO") & "," & rsData("MinUsage")  & "," & rsData("SubTotal")  & "," & rsData("PPN")  & "," & rsData("StampFee") & "," & rsData("CurrentBalance")& "," & rsData("Tota") & "') " 
				StrValue = "Values('" & rsData("MonthP") & "','" & rsData("YearP") & "','"
				strExecute = strInsert & StrValue 
				response.write strExecute
				'response.write rsData("Phonenumber") 
				'response.write rsData("PreviousBalance") & "<br>"
				rsData.movenext
			loop
		
		If Err.Number <> 0 Then
			response.write Err.Description
		end if
		

	Next		
End If


'--------------------------------------------
' ValidFileExtension()
' You give a list of file extensions that are allowed to be uploaded.
' Purpose:  Checks if the file extension is allowed
' Inputs:   strFileName -- the filename
'           strFileExtension -- the fileextensions not allowed
' Returns:  boolean
' Gives False if the file extension is NOT allowed
'--------------------------------------------
Function ValidFileExtension(strFileName, strFileExtensions)

    Dim arrExtension
    Dim strFileExtension
    Dim i
    
    strFileExtension = UCase(GetFileExtension(strFileName))
    
    arrExtension = Split(UCase(strFileExtensions), ";")
    
    For i = 0 To UBound(arrExtension)
        
        'Check to see if a "dot" exists
        If Left(arrExtension(i), 1) = "." Then
            arrExtension(i) = Replace(arrExtension(i), ".", vbNullString)
        End If
        
        'Check to see if FileExtension is allowed
        If arrExtension(i) = strFileExtension Then
            ValidFileExtension = True
            Exit Function
        End If
        
    Next
    
    ValidFileExtension = False

End Function

'--------------------------------------------
' InValidFileExtension()
' You give a list of file extensions that are not allowed.
' Purpose:  Checks if the file extension is not allowed
' Inputs:   strFileName -- the filename
'           strFileExtension -- the fileextensions that are allowed
' Returns:  boolean
' Gives False if the file extension is NOT allowed
'--------------------------------------------
Function InValidFileExtension(strFileName, strFileExtensions)

    Dim arrExtension
    Dim strFileExtension
    Dim i
        
    strFileExtension = UCase(GetFileExtension(strFileName))
    
    'Response.Write "filename : " & strFileName & "<br>"
    'Response.Write "file extension : " & strFileExtension & "<br>"    
    'Response.Write strFileExtensions & "<br>"
    'Response.End 
    
    arrExtension = Split(UCase(strFileExtensions), ";")
    
    For i = 0 To UBound(arrExtension)
        
        'Check to see if a "dot" exists
        If Left(arrExtension(i), 1) = "." Then
            arrExtension(i) = Replace(arrExtension(i), ".", vbNullString)
        End If
        
        'Check to see if FileExtension is not allowed
        If arrExtension(i) = strFileExtension Then
            InValidFileExtension = False
            Exit Function
        End If
        
    Next
    
    InValidFileExtension = True

End Function

'--------------------------------------------
' GetFileExtension()
' Purpose:  Returns the extension of a filename
' Inputs:   strFileName     -- string containing the filename
'           varContent      -- variant containing the filedata
' Outputs:  a string containing the fileextension
'--------------------------------------------
Function GetFileExtension(strFileName)

    GetFileExtension = Mid(strFileName, InStrRev(strFileName, ".") + 1)
    
End Function

%>

