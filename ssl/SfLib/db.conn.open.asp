<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: db.conn.open.asp
	 


'@DESCRIPTION: Establish an open ado connection 

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws  and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO
	
	' Variable Declarations
	Dim cnn, DSN_Name, SSLSession

	' Object Creation
	Set cnn=Server.CreateObject("ADODB.Connection")
	
	If Session("DSN_Name") = "" Then
	
		'***********DSN NAME*********************
		DSN_Name = "YOUR DSN NAME"
		'****************************************
		Session("DSN_Name") = trim(DSN_Name)
	Else
		DSN_Name = Session("DSN_Name")	
	End If

	If Session("SSLSession") = "" Then	
		Session("GeneratedKey") = generateKey()
		Session("SessionID") = Request.Form("SessionID")
		Session("sfREFERER") = Request.Form("REFERER")
		Session("sfHTTP_REFERER") = Request.Form("HTTP_REFERER")
		Session("sfREMOTE_ADDRESS") = Request.Form("REMOTE_ADDRESS")
		
		If Request.Form("LoggedSessionID") <> "" AND Request.Form("LoggedSessionID") <> Request.Cookies("EndSession") Then
			Response.Cookies(Session("GeneratedKey") & "sfOrder")("SessionID") = Session("SessionID")
			Response.Cookies(Session("GeneratedKey") & "sfOrder").Expires = Date() + 1
		End If				
		
		If Request.Form("CustID") <> "" Then
			Session("CustID") = Trim(Request.Form("CustID"))
		ElseIf Request.Cookies("sfCustomer")("CustID") <> "" Then
			Session("CustID") = Request.Cookies("sfCustomer")("CustID")
		End IF		
		
		Session("SSLSession") = "1"		
	End If
	cnn.open DSN_Name	
	
	'-------------------------------------------------------------------
	' Function that releases an object
	'-------------------------------------------------------------------
	Sub closeObj(objItem)
		On Error Resume Next
		objItem.Close
		Set objItem=nothing	
		On Error GoTo 0
	End Sub

	'-------------------------------------------------------------------
	' Generates Random Key
	'-------------------------------------------------------------------
	Function generateKey
		Dim Random_Number_Min, Random_Number_Max, iKey
		Randomize
		Random_Number_Min = 1000000
		Random_Number_Max = 9999999
		iKey = Int(((Random_Number_Max-Random_Number_Min+1) * Rnd) + Random_Number_Min)
		generateKey = iKey
	End Function
%>








