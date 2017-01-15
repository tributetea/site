<%
	Response.Buffer = True
	%>
	<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: error_trap.asp
	 


'@DESCRIPTION: Traps asp errors

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO
%>

<%
'@BEGINCODE

	Sub CheckForError()

		Dim strErrorNumber
		Dim strErrorDescription
			
		If Err.number = 0 Then
			Exit Sub
		End If
		
		strErrorNumber = Err.number 
		strErrorDescription = Err.description

Response.write "<html>"
Response.write "<head>"
Response.write "<link rel=""stylesheet"" href=""../sfCSS.css"" type=""text/css"">"
Response.write "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"">"
Response.write "<title>SF Error Trapping Page</title>"
Response.write "</head>"

Response.write "<body link=""" & C_LINK & """ vlink=""" & C_VLINK & """ alink=""" & C_ALINK & """>"

Response.write "<table border=""0"" cellpadding=""1"" cellspacing=""0"" class=""tdbackgrnd"" width=""" & C_WIDTH & """ align=""center"">"
Response.write "  <tr>"
Response.write "    <td>"
Response.write "	  <table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""3"">"
Response.write "	    <tr>"
Response.write "          <td align=""center"" class=""tdTopBanner"">" & C_STORENAME
Response.write "	      </td>"
Response.write "	    </tr>"
Response.write "	    <tr>"
Response.write "	      <td class=""tdContent"">"
Response.write "		    <p>StoreFront has encountered an error with the page you are attempting to view.</p>"
Response.write "		    <p>Please contact the webmaster of the site at <b><a href=""mailto:" & strEmailAddress & "?SUBJECT=StoreFront Error"">" & strEmailAddress& "</a></b> and tell them that you've experienced an error in <b>" & strPageName& "</b>.<br><br>"
Response.write "		    <b>The error number was:</b>" & strErrorNumber & "<br><br>"
Response.write "		    <b>The error message was:</b> " & strErrorDescription & "<br><br>"
Response.write "		    </p>"
Response.write "		    <p align=""center"">We apologize for the inconvenience.<br></p>"
Response.write "	      </td>"
	      
Response.write "                  <!--#include file=""footer.txt""-->"
Response.write "	      </table>"
Response.write "        </td>"
Response.write "      </tr>"
Response.write "    </table>"

Response.write "  </body>"
Response.write "</html>"

	End Sub

'@ENDCODE
%>








