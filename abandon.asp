<%@ Language=VBScript %>
<%	option explicit 
	Response.Buffer = True
	%>
<!--#include file="SFLib/db.conn.open.asp"-->
<%
Response.Buffer = true
Session.Abandon 

'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: abandon.asp
 
'

'@DESCRIPTION: abandons session

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO

Dim rsAdmin, sPath

Set rsAdmin = Server.CreateObject("ADODB.Recordset")
rsAdmin.Open "sfAdmin", cnn, 0, 1, &H0002
sPath = rsAdmin.Fields("adminDomainName")
rsAdmin.Close
cnn.Close
Set rsAdmin = Nothing
Set cnn = Nothing

Response.Redirect(sPath)
%>
<html>
<head>
<link rel="stylesheet" href="sfCSS.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>SF Abandons Session</title>


</head>
<body>
</body>
</html>



