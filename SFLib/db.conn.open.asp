<%

	'@BEGINVERSIONINFO

	'@APPVERSION: 50.4011.0.2

	'@FILENAME: db.conn.open.asp
	 

	'

	'@DESCRIPTION: Opens Database Connection

	'@STARTCOPYRIGHT
	'The contents of this file is protected under the United States
	'copyright laws as an unpublished work, and is confidential and proprietary to
	'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
	'expressed written permission of LaGarde, Incorporated is expressly prohibited.
	'
	'(c) Copyright 2000 by LaGarde, Incorporated.  All rights reserved.
	'@ENDCOPYRIGHT

	'@ENDVERSIONINFO
	
	' Variable Declarations
	Dim cnn, DSN_Name

	' Object Creation
	Set cnn=Server.CreateObject("ADODB.Connection")
	
	DSN_Name = Session("DSN_Name")	

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


%>



