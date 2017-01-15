<!--#include file="mail.asp"-->
<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: inclogin.asp
	 

'

'@DESCRIPTION: login the user, send forgotton password, change user information

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws  and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO
'----------------------------------------------------
' Writes a temp record into customer table
' Returns the customer ID, writes to a cookie
'----------------------------------------------------
Function getCustomerID(sEmail,sPassword)
	Dim rsWrite, iCustID, sCookieDie
	
	Set rsWrite = Server.CreateObject("ADODB.RecordSet")
		rsWrite.CursorLocation = adUseClient
		rsWrite.Open "sfCustomers", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
		rsWrite.AddNew
		rsWrite.Fields("custFirstName")= "Saved Cart Customer"
		rsWrite.Fields("custEmail") = sEmail
		rsWrite.Fields("custPasswd") = sPassword
		rsWrite.Update
		iCustID = rsWrite.Fields("custID")	

	' Write to cookie	
	sCookieDie= Date() + 730
	Response.Cookies("sfCustomer").Expires = sCookieDie
	Response.Cookies("sfCustomer")("custID") = iCustID	
		
	' Return CustID	
	getCustomerID = iCustID	
	closeobj(rsWrite)
End Function

'-------------------------------------------------------
' Updates CustID in Saved Cart
'-------------------------------------------------------
Sub UpdateCustID(iCustID)
	Dim sLocalSQL, rsUpdateCust
	
	sLocalSQL = "SELECT	odrdtsvdCustID FROM sfSavedOrderDetails WHERE odrdtsvdCustID = 0 AND odrdtsvdSessionID = " & Session("SessionID")
		
	Set rsUpdateCust = Server.CreateObject("ADODB.RecordSet")
		rsUpdateCust.Open sLocalSQL, cnn, adOpenDynamic,adLockOptimistic, adCmdText
		If rsUpdateCust.EOF And rsUpdateCust.BOF Then Response.Redirect "abandon.asp"  
		rsUpdateCust.Fields("odrdtsvdCustID") = iCustID
		rsUpdateCust.Update	
	closeobj(rsUpdateCust)	
	
	If vDebug = 1 Then Response.Write "<p> UpdateCustID SQL = " & sLocalSQL & " <br><font color=""red""> Successful Update</font>"
	
End Sub

'--------------------------------------------------------
' Sends password associated with email address, returns success or failure
'--------------------------------------------------------
Function SendPassword(sEmail)	
	Dim sLocalSQL, sPasswd, rsGetPasswd, bSuccess, sErrorDescription, sInfo
	
	sLocalSQL = "SELECT custPasswd FROM sfCustomers WHERE custEmail = '" & sEmail & "'"
	
	Set rsGetPasswd = Server.CreateObject("ADODB.RecordSet")
	rsGetPasswd.Open sLocalSQL,cnn,adOpenForwardOnly,adLockReadOnly,adCmdText				 
	
	If rsGetPasswd.BOF Or rsGetPasswd.EOF  Or sEmail = "" Then
		sErrorDescription = "<br>No password found for email " & sEmail
		bSuccess = 0
	Else	
		sPasswd = trim(rsGetPasswd.Fields("custPasswd"))
		sInfo = sEmail & "|" & sPasswd
		Call createMail("FPWD",sInfo)		
		bSuccess = 1
	End If
	
	SendPassword = bSuccess
	closeobj(rsGetPasswd)
End Function

'--------------------------------------------------------
' Check if customer has filled out form or is a saved cart customer
' returns a 1 for SvdCartCustomer, 0 for customers who've already filled out the form
'--------------------------------------------------------
Function getSvdCartCustomer(iCookieID,sReturnType)
	Dim sLocalSQL, sReturn, rsSvdCust
	
	sLocalSQL = "SELECT custFirstName, custEmail FROM sfCustomers WHERE custID =" & iCookieID
	
	Set rsSvdCust = Server.CreateObject("ADODB.RecordSet")
	rsSvdCust.Open sLocalSQL,cnn,adOpenForwardOnly,adLockReadOnly,adCmdText				 
	
	If rsSvdCust.BOF Or rsSvdCust.EOF Then
		If vDebug = 1 Then Response.Write "<br>No record found for customer :" & iCookieID
		sReturn = -1
		' redirect to error.asp	-- to be done
	Else		
			If sReturnType = "boolean" Then
					If rsSvdCust.Fields("custFirstName")= "Saved Cart Customer" Then
						sReturn = 1
					Else 
						sReturn = 0
					End If				
			ElseIf sReturnType = "email" Then			
					If rsSvdCust.Fields("custFirstName") = "Saved Cart Customer" Then
						sReturn = rsSvdCust.Fields("custEmail")			
					Else
						sReturn = -1
					End If	
			Else
				sReturn = -1
			' End return type If	
			End If		
		
	' End recordset if	
	End If	
	
	closeobj(rsSvdCust)
	getSvdCartCustomer = sReturn
End Function


'-------------------------------------------------------------
' Updates password for a customer record
'-------------------------------------------------------------
Sub UpdatePassword(iCustID,sNewPassword)
	Dim sLocalSQL, rsCust
	
	sLocalSQL = "SELECT custPasswd FROM sfCustomers WHERE custID = " & iCustID
	Set rsCust = Server.CreateObject("ADODB.RecordSet")
		rsCust.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
		rsCust.Fields("custPasswd") = sNewPassword
		rsCust.Update
	
	closeobj(rsCust)	
End Sub
%>








