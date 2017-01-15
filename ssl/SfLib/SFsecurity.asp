<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: processor.asp
	 

'

'@DESCRIPTION: Enforces Login Security for all Admin Files

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
'Original File supplied by Michael Hamilton, Suite500.net
'@ENDCOPYRIGHT

'@ENDVERSIONINFO

	If Request.ServerVariables("LOGON_USER") = "" Then
		Dim fso, drv, drvPath
		drvPath = left(Request.ServerVariables("PATH_TRANSLATED"),2)  ' get the drive ???
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set drv = fso.GetDrive(fso.GetDriveName(drvPath))
		if drv.filesystem = "NTFS" then
			response.clear
			response.addheader "WWW-Authenticate","Basic Realm="&chr(34)&"StoreFront 5.0 Administration http://www.storefront.net"&chr(34)
			response.addheader "WWW-Authenticate","NTLM Realm="&chr(34)&"StoreFront 5.0 Administration http://www.storefront.net"&chr(34)
			Response.Status = "401 Access Denied"

response.write "<head>"
response.write "<link rel=""stylesheet"" href=""sfCSS.css"" type=""text/css"">"
response.write "<style>"
response.write "a:link                  {font:8pt/11pt verdana; color:FF0000}"
response.write "a:visited               {font:8pt/11pt verdana; color:#4e4e4e}"
response.write "</style>"
response.write "<META NAME=""ROBOTS"" CONTENT=""NOINDEX"">"
response.write "<title>You are not authorized to view this page</title>"
response.write "<META HTTP-EQUIV=""Content-Type"" Content=""text-html; charset=Windows-1252"">"
response.write "</head>"
response.write "<script> "
response.write "function Homepage(){"
response.write "<!--"
response.write "// in real bits, urls get returned to our script like this:"
response.write "// res://shdocvw.dll/http_404.htm#http://www.DocURL.com/bar.htm "

response.write "	//For testing use DocURL = ""res://shdocvw.dll/http_404.htm#https://www.microsoft.com/bar.htm"""
response.write "	DocURL=document.URL;"
	
response.write "	//this is where the http or https will be, as found by searching for :// but skipping the res://"
response.write "	protocolIndex=DocURL.indexOf(""://"",4);"
	
response.write "	//this finds the ending slash for the domain server "
response.write "	serverIndex=DocURL.indexOf(""/"",protocolIndex + 3);"

response.write "	//for the href, we need a valid URL to the domain. We search for the # symbol to find the begining "
response.write "	//of the true URL, and add 1 to skip it - this is the BeginURL value. We use serverIndex as the end marker."
response.write "	//urlresult=DocURL.substring(protocolIndex - 4,serverIndex);"
response.write "	BeginURL=DocURL.indexOf(""#"",1) + 1;"
response.write "	urlresult=DocURL.substring(BeginURL,serverIndex);"
		
response.write "	//for display, we need to skip after http://, and go to the next slash"
response.write "	displayresult=DocURL.substring(protocolIndex + 3 ,serverIndex);"
response.write "	document.write('<A HREF=""' + urlresult + '"">' + displayresult + ""</a>"");"
response.write "}"
response.write "//-->"
response.write "</script>"
response.write "<body bgcolor=""FFFFFF"">"
response.write "<table width=""410"" cellpadding=""3"" cellspacing=""5"">"
response.write "  <tr>    "
response.write "    <td align=""left"" valign=""middle"" width=""360"">"
response.write "	<h1 style=""COLOR:000000; FONT: 13pt/15pt verdana""><!--Problem-->" & Request.ServerVariables("AUTH_TYPE")
response.write "StoreFront 5.0 Authorization Error<br><br>You are not authorized to view this page</h1>"
response.write "    </td>"
response.write "  </tr>"

response.write "  <tr>"
response.write "    <td width=""400"" colspan=""2"">"
response.write "	<font style=""COLOR:000000; FONT: 8pt/11pt verdana"">You do not have permission to view this directory or page using the credentials you supplied.</font></td>"
response.write "  </tr>"
response.write "  <tr>"
response.write "    <td width=""400"" colspan=""2"">"
response.write "	<font style=""COLOR:000000; FONT: 8pt/11pt verdana"">"

response.write "	<hr color=""#C0C0C0"" noshade>"
	
response.write "   <p>Please try the following:</p>"
response.write "<ul>"
response.write "<li>Click the <a href=""javascript:location.reload()"">Refresh</a> button to try again with different credentials.</li>"
response.write "<li>If you believe you should be able to view this directory or page, please contact the Web site administrator by using the e-mail address or phone number listed on the "
response.write "	<script> "
response.write "	<!--"
response.write "	if (!((window.navigator.userAgent.indexOf(""MSIE"") > 0) && (window.navigator.appVersion.charAt(0) == ""2"")))"
response.write "	{"
response.write "		Homepage();"
response.write "	}"
response.write "	//-->"
response.write "	</script>"
response.write "	home page.</li>"
response.write "</ul>"
response.write "    <h2 style=""font:8pt/11pt verdana; color:000000"">HTTP 401.2 - Unauthorized: Logon failed<br></h2>"
response.write "	<hr color=""#C0C0C0"" noshade>"
response.write "	<p>Technical Information (for support personnel)</p>"
response.write "	<ul>"
response.write "	<li>Background:<br>"
response.write "	This is caused by a providing incorrect credentials."
response.write "<p>"
response.write "<li>More information:<br>"
response.write "<a href=""http://www.storefront.net"" target=""_blank"">StoreFront</a>"
response.write "</li>"

response.write "</ul> "
response.write "	</font></td>"
response.write "  </tr>"
response.write "</table>"
response.write "</body>"
response.write "</html>"

	Response.End		
End If
End If
Response.Buffer = True


%>









