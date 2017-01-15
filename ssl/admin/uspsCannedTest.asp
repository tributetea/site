<%@ Language=VBScript %>

<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: uspscannedtest.asp
	 

'

'@DESCRIPTION:   Required for Account Activation with USPS On-Line Tools

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO	
%>
<!--#include file="../SFLib/db.conn.open.asp"-->
<!--#include file="../SFLib/adovbs.inc"-->
<%

dim dTest,fTest
 dtest = DomesticCanned

if isnumeric(dtest) then
  Response.Write "Domestic Canned Test Passed!, Results = " & dtest
 else
  Response.Write dtest
 end if 
 Response.Write "<br>"
ftest = InterNationalCanned
 if isnumeric(ftest) then
  Response.Write "International Canned Test Passed!, Results = " & ftest
 else
  Response.Write ftest
 end if  

%>


<%
function DomesticCanned()
dim Rst,uspsUsername,uspsPassword
dim sResults, Currnode,ParseLogin,ParseUsername
set httpconn = server.createobject("httpcom.chttpcom")


Set rst = Server.CreateObject("ADODB.RecordSet")
	sSQL = "SELECT adminUsPsUserName,adminUsPsPassword From sfAdmin"
		
	rst.Open sSQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText


    if Rst.EOF =false and Rst.BOF =false then

	If inStr(rst.Fields("adminUsPsUserName"),",") Then

		ParseUsername = Split(rst.Fields("adminUsPsUserName"),",")
	
		uspsUsername = ParseUsername(0)
	
	
	Else
	
	uspsUsername = rst.Fields("adminUsPsUserName")
	End If

	If InStr(rst.Fields("adminUsPsPassword"),",") Then
     	ParseLogin = Split(rst.Fields("adminUsPsPassword"),",")
       	uspsPassword	= ParseLogin(0)
    Else
    	uspsPassword = rst.Fields("adminUsPsPassword")
    End If
    '  uspsUsername = Rst.Fields("adminUsPsUserName")
    '  uspsPassword = Rst.Fields("adminUsPsPassword")
      
   else
   
      DomesticCanned ="Failed No User Name or Pass word"
     Response.End 
     exit function
   end if  
   
  if isnull(uspsUsername) or uspsUsername= "" or isnull(uspsPassword) or uspsPassword = "" then  
     DomesticCanned ="Failed No User Name or Pass word"
      Response.End 
     exit function
  end if  


set xmlDoc = Server.CreateObject("MSXML.DOMDocument")
    
Set RequestLevel = xmlDoc.createElement("RateRequest")
Set RequestLevel = xmlDoc.createElement("RateRequest")
RequestLevel.setAttribute "USERID", uspsUsername

RequestLevel.setAttribute "PASSWORD", uspsPassword
'" '  uspsPassword

Set PackageLevel = xmlDoc.createElement("Package")
PackageLevel.setAttribute "ID", "0"
Set PackageElementLevel = xmlDoc.createElement("Service")
Set t = xmlDoc.createTextNode("EXPRESS")                   'Request.Form("ServiceType"))
PackageElementLevel.appendChild (t)
Call PackageLevel.appendChild(PackageElementLevel)
Set PackageElementLevel = xmlDoc.createElement("ZipOrigination")
Set t = xmlDoc.createTextNode("20770")                            'Request.Form("txtOriginZip"))
PackageElementLevel.appendChild (t)
Call PackageLevel.appendChild(PackageElementLevel)
Set PackageElementLevel = xmlDoc.createElement("ZipDestination")
Set t = xmlDoc.createTextNode("20852")                               '     Request.Form("txtDestZip"))
PackageElementLevel.appendChild (t)
Call PackageLevel.appendChild(PackageElementLevel)
Set PackageElementLevel = xmlDoc.createElement("Pounds")
Set t = xmlDoc.createTextNode("10")                                    '  Request.Form("txtPounds"))
PackageElementLevel.appendChild (t)
Call PackageLevel.appendChild(PackageElementLevel)
Set PackageElementLevel = xmlDoc.createElement("Ounces")
Set t = xmlDoc.createTextNode("0")    ')                               'Request.Form("txtOunces"))
PackageElementLevel.appendChild (t)
Call PackageLevel.appendChild(PackageElementLevel)
Set PackageElementLevel = xmlDoc.createElement("Container")
Set t = xmlDoc.createTextNode("None")                            'Request.Form("selContainerDesc"))
PackageElementLevel.appendChild (t)
Call PackageLevel.appendChild(PackageElementLevel)
Set PackageElementLevel = xmlDoc.createElement("Size")
Set t = xmlDoc.createTextNode("REGULAR")                    'Request.Form("selSize")
PackageElementLevel.appendChild (t)
Call PackageLevel.appendChild(PackageElementLevel)
Set PackageElementLevel = xmlDoc.createElement("Machinable")
Set t = xmlDoc.createTextNode("")                    'Request.Form("selSize")
PackageElementLevel.appendChild (t)

Call PackageLevel.appendChild(PackageElementLevel)
Call RequestLevel.appendChild(PackageLevel)

Call xmlDoc.appendChild(RequestLevel)
'Call xmlDoc.save(Server.MapPath("\sample2.xml"))


srvr = "testing.shippingapis.com" 'You must modify the Server name
Port = 80
Path = "/ShippingAPITest.dll" 'You must modify the Path name
query = "" '"?API=Rate&XML=" & xmlDoc.xml 'xmldoc.xmloc.xml
 msg = "API=Rate&XML=" & URL_Encode(xmlDoc.xml)
'You must modify the API name
'Response.Write srvr & path & msg

'Response.End
contentType = "" ' "application/x-www-form-urlencoded"
proxyServer = "" 'You must modify proxy server or leave
'blank if no proxy
proxyPort = 0 'Modify the proxy port, if necessary. Leave blank if no
'proxy.

Dim resp 
resp = ""
'on error resume next
'respcode = "beforE"
'Response.Write  respcode
sRespcode = httpConn.GetResponse(srvr, Port, Path, query, msg, contentType, proxyServer, proxyPort, resp)
If Err.Number <> 0 Then
        If Err.number = 424 Then
		  DomesticCanned = "The USPS component is not properly installed."
		Else
			DomesticCanned = "Failed"
		End If
Else

DomesticCanned = getprice(resp) 


End If
    Set xmlDoc = Nothing
    Set httpConn = Nothing
    Set httpConn = Nothing
    set rsAdmin = nothing
end function




Function URL_Encode(s)
    
 URL_SAFE_CHARS = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890$-_.+!*'(),"

    outStr = ""
    For i = 1 To Len(s)
        c = Right(Left(s, i), 1) 'set c equal to ith char
        res = InStr(URL_SAFE_CHARS, c)
        If res = 0 Then
            'escape unsafe characters as %xx where xx is the
            'hexadecimal value of the character's ASCII code.
            outStr = outStr & "%" & Right("00" & CStr(Hex(Asc(c))), 2)
        Else
            outStr = outStr & c
        End If
    Next
    
    URL_Encode = outStr
    
End Function
Function GetPrice(vXml)

On error resume next
Dim sNew,posit

	posit = instr(vxml,"<Postage>")
	If posit > 0 then
		snew = mid(vxml,posit,len(vxml)- posit )	
		posit = instr(snew,"</Postage>")
		snew = mid(snew,10,posit-10)
		getprice = snew 
	Else  
		getprice = vxml
	End if
End Function


Function InterNationalCanned()
	Dim	sLen,uPounds,uOunces,httpconn,xmlDoc,srvr,Port,Path,query,msg,contentType,proxyServer,proxyPort,sRespcode,resp,RequestLevel,PackageLevel,PackageElementLevel,t
	dim ParseUsername,ParseLogin
On error resume next
    Set rst = Server.CreateObject("ADODB.RecordSet")
	sSQL = "SELECT adminUsPsUserName,adminUsPsPassword From sfAdmin"
		
	rst.Open sSQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
    if Rst.EOF =false and Rst.BOF =false then
	If inStr(rst.Fields("adminUsPsUserName"),",") Then

		ParseUsername = Split(rst.Fields("adminUsPsUserName"),",")
	
		uspsUsername = ParseUsername(0)
	
	
	Else
	
	uspsUsername = rst.Fields("adminUsPsUserName")
	End If

	If InStr(rst.Fields("adminUsPsPassword"),",") Then
     	ParseLogin = Split(rst.Fields("adminUsPsPassword"),",")
       	uspsPassword	= ParseLogin(0)
    Else
    	uspsPassword = rst.Fields("adminUsPsPassword")
    End If

'      uspsUsername = Rst.Fields("adminUsPsUserName")
 '     uspsPassword = Rst.Fields("adminUsPsPassword")
     ' Response.Write uspsUsername
   else
     InterNationalCanned = "UserName or Password missing"
      Response.End 
     exit function 
   end if  
  if isnull(uspsUsername) or uspsUsername= "" or isnull(uspsPassword) or uspsPassword = "" then  
      InterNationalCanned = "UserName or Password missing"
      Response.End 
     exit function 
   end if  
  
'  http://SERVERNAME/ShippingAPITest.dll?API=IntlRate&XML=<IntlRateRequest
'USERID="xxxxxxxx" PASSWORD="xxxxxxxx"><Package ID="0"><Pounds>
'2</Pounds><Ounces>0</Ounces><MailType>Letters or Letter Packages
'</MailType><Country>Albania</Country></Package></IntlRateRequest>

	Set xmlDoc = Server.CreateObject("MSXML.DOMDocument")                     
	Set RequestLevel = xmlDoc.createElement("IntlRateRequest")  
	Set RequestLevel = xmlDoc.createElement("IntlRateRequest")
		RequestLevel.setAttribute "USERID", uspsUsername
		RequestLevel.setAttribute "PASSWORD", uspsPassword
	Set RateLevel = xmlDoc.createElement("Package")
		RateLevel.setAttribute "ID", "0"
	Set RateElementLevel = xmlDoc.createElement("Pounds")
	Set t = xmlDoc.createTextNode(2)
		RateElementLevel.appendChild (t)
		Call RateLevel.appendChild(RateElementLevel)
	Set RateElementLevel = xmlDoc.createElement("Ounces")
	Set t = xmlDoc.createTextNode(0)
		RateElementLevel.appendChild (t)
		Call RateLevel.appendChild(RateElementLevel)
	Set RateElementLevel = xmlDoc.createElement("MailType")
	Set t = xmlDoc.createTextNode("Letters or Letter Packages")
		RateElementLevel.appendChild (t)
	Call RateLevel.appendChild(RateElementLevel)
	Set RateElementLevel = xmlDoc.createElement("Country")
	Set t = xmlDoc.createTextNode("Albania")
		RateElementLevel.appendChild (t)
	Call RateLevel.appendChild(RateElementLevel)
	Call RequestLevel.appendChild(RateLevel)
	Call xmlDoc.appendChild(RequestLevel)
	'Call xmlDoc.save(Server.MapPath("/Sendme.xml"))
  ' Response.Write xmlDoc.xml
  ' Response.End 
	Set httpconn = server.createobject("httpcom.chttpcom")	

		srvr = "testing.shippingapis.com" 'You must modify the Server name
		Port = 80
		Path = "/ShippingAPITest.dll" 'You must modify the Path name
		query = "" '"?API=Rate&XML=" & xmlDoc.xml 'xmldoc.xmloc.xml
		msg = "API=IntlRate&XML=" & URL_Encode(xmlDoc.xml)
		contentType =  "application/x-www-form-urlencoded"
		proxyServer = "" 'You must modify proxy server or leave
		proxyPort = 0 'Modify the proxy port, if necessary. Leave blank if no proxy

	resp = ""

	On Error Resume Next
		Err.Number = 0
		sRespcode = httpConn.GetResponse(srvr, Port, Path, query, msg, contentType, proxyServer, proxyPort, resp)
	If Err.Number <> 0 Then
	   
	    If Err.number = 424 Then
			InterNationalCanned = "The USPS component is not properly installed."
		Else
			InterNationalCanned ="Failed"
		End If
	    Set xmlDoc = Nothing
	    Set httpConn = Nothing
	    
	    Set httpConn = Nothing
	    set rst = nothing
		
		Exit Function
Else
  ' Response.Write resp
   
	InterNationalCanned = GetPrice(resp)

End If

        Set xmlDoc = Nothing
    Set httpConn = Nothing
    Set httpConn = Nothing
    set rst = nothing		
	
End Function
%>
































