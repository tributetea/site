<%@ LANGUAGE = VBScript %>
<%
' ********************************************************
' System: Generic Component Checker 
' Create Date: 11/28/2001
' Author: Suite500.Net
' Description: Determines if COM components are installed
' ********************************************************
' --------------------------------------------------------------------------------------
' -			COPYRIGHT NOTICE
' -
' -	The contents of this file is protected under the United States
' -	copyright laws as an unpublished work, and is proprietary to Suite500.Net
' -
' -     Its use or disclosure in whole is not permitted.
' -
' -     (c) Copyright 2001 by Suite500.net  All rights reserved.
' -
' --------------------------------------------------------------------------------------

set xmlHttp = server.CreateObject ("Microsoft.XMLHTTP")
xmlHttp.open "POST", "http://support.storefront.net/ComDiag/LoadDiag.asp", false

xmlHttp.setRequestHeader "XML","XML" 
xmlHttp.send()

set diagxml = server.CreateObject ("msxml.domdocument")
diagxml.async = false



if diagxml.loadxml(xmlHttp.responseText) then 
	

	set diags = diagxml.documentelement
	response.write "<center><h2><B>" & diags.selectsinglenode("Titles").text & "</h2></b>"
	response.write "<h3><B>Version " & diags.selectsinglenode("Version").text & "</h3></b></center>"
else
	response.write "Could not Find the COM Diag Configuration File<br>"
	response.end
end if
%>

<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
	<TR>
		<TD VALIGN="TOP">
<B>ServerName: </B><%=Request.ServerVariables("SERVER_NAME")%><BR>
<B>Server Type: </B><%=Request.ServerVariables("Server_software")%><BR>
<B>ServerProtocol: </B><%=Request.ServerVariables("SERVER_PROTOCOL")%><BR>
<B>PathInfo: </B><% = Request.ServerVariables("PATH_INFO")%><BR>
<B>PathTranslated: </B><%=Request.ServerVariables("PATH_TRANSLATED")%><BR>
<B>Shared Hosting: </B>This website is site # <%=Request.ServerVariables("INSTANCE_ID")%> on the server<BR>
		</TD>
		<TD VALIGN="TOP">
<FONT SIZE="3"><B>Script Engine</B><BR>
<B>Type: </B><% = ScriptEngine%><BR>
<B>Version: </B><%=ScriptEngineMajorVersion()%>.<%=ScriptEngineMinorVersion()%><BR>
<B>Build: </B><%=ScriptEngineBuildVersion()%><BR>
		</TD>
		</TR>
</TABLE>
<%
set configxml = server.CreateObject ("msxml.domdocument")
configxml.async = false

for each diag in diags.selectnodes("list")
	thefile = diag.selectsinglenode("location").text
	
	xmlHttp.open "GET", thefile, false
	xmlHttp.send()
	
	if configxml.loadxml(xmlHttp.responseText) then 
		set coms = configxml.documentelement
		response.write "<hr><center><h3><B>" & diag.selectsinglenode("titles").text & "</h3></b></center>"
		process
	else
		response.write "Could not Find the Configuration File<br>"
'		response.end
end if
%>
<br>
<%
next

set diagxml = nothing
set configxml = nothing
set xmlhttp = nothing

Function process

on error resume next
for each com in coms.selectnodes("com")
	err.clear
	set theobject = server.createobject(com.selectsinglenode("CreateUsing").text)
	If Err.Number <> 0 Then
		%><font color = "black"><%
      		Response.Write(com.selectsinglenode("Description").text & " is " & "<font color=" & chr(34) & "red" & chr(34) & "> <b>not</b></font> installed")
   	Else
		%><font color = "green"><%
		if com.selectsinglenode("Version").text = "" then
			Response.Write(com.selectsinglenode("Description").text & "<B> is installed</B>")
		else
			version = ""
			version = theobject.VERSION 			
			Response.Write(com.selectsinglenode("Description").text & " Version: " & VERSION & "<B> is installed</B>")
		end if%></font><%
		
	end if
	If com.selectsinglenode("URL").text <> "" then
		response.write ("<font size=" & chr(34) & "2" & chr(34) & ">   , this vendors url is <A HREF=" & chr(34) & com.selectsinglenode("URL").text & chr(34) & "target=" & chr(34) & "_new" & chr(34) & ">" & com.selectsinglenode("URL").text & "</A></font><br>")
	else
		response.write ("<br>")
	end if

theobject.close
next

set theobject = nothing

end function
%>
