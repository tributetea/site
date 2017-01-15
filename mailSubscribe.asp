<%@ Language=VBScript %>
<%	option explicit 
	Response.Buffer = True
	%>
<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file="sfLib/incDesign.asp"-->
<!--#include file="SFLib/incGeneral.asp"-->
<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: mailsubscribe.asp
 
'

'@DESCRIPTION: Allows Customer to Subscribe to Merchant Mailing List

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
	Dim rsCust, sFirstName, sLastName,sPassword,sConfirmPass,sEmail, sSql
	sEmail         = Trim(Request.Form("emailadd"))
	If  semail <> "" then	
		sFirstName		= Trim(Request.Form("fname"))
		sLastName		= Trim(Request.Form("lname"))
		sPassword		= Trim(Request.Form("password"))
		sConfirmPass   = Trim(Request.Form("passwordconfirm"))
		
	
		Set rsCust = Server.CreateObject("ADODB.RecordSet")			
	sSql="Select * From sfCustomers where custEmail = " & "'" & sEmail & "'"
		rsCust.Open  sSql , cnn, adOpenKeyset, adLockOptimistic, adCmdText
   if rsCust.EOF =false and rsCust.BOF =false then
	   
	   	    	'rsCust.Fields("custFirstName")		= sfirstname
		    	'rsCust.Fields("custLastName")		= slastName
			    rsCust.Fields("custPasswd")	        = spassword
			    rsCust.Fields("custIsSubscribed")	= 1
		        rsCust.Update 
                  
     
         
    else
    closeObj(rsCust)
     	 sSql = " Select * from sfCustomers " 
     	
     			rsCust.Open sSql , cnn, adOpenKeyset, adLockOptimistic, adCmdText

            rsCust.AddNew 
               rsCust.Fields("custFirstName")		= sfirstname
		    	rsCust.Fields("custLastName")		= slastName
			    rsCust.Fields("custPasswd")	        = spassword
			    rsCust.Fields("custEmail")	= sEmail
			     rsCust.Fields("custLastAccess") = Date()
			    rsCust.Fields("custIsSubscribed")	= 1
		        rsCust.Update
		end if         
		closeObj(rsCust)
	Else
		closeObj(rsCust)
		sfirstname		=""
		slastname = ""
		spassword      ="" 
		sconfirmpass=""
		sEmail =""
	End If	
%>
<html>
<head>

<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%= C_STORENAME %>-Join Mailing List</title>


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function form1_onsubmit() {
			                
                            if (form1.fname.value == "") {
                            alert("Must Fill in User Name");
                            form1.fname.focus()
                            return false; }
                           if (form1.lName.value == "") {
                            alert("Must Fill in Last Name");
                            form1.lName.focus()
                            return false; }
                           if (form1.password.value == "") {
                            alert("Must supply a password");
                            form1.password.focus()
                            return false; }
                            if (form1.password2.value == "") {
                            alert("Must supply a matching password confirmation");
                            form1.password2.focus()
                            return false; } 
                          if (form1.password.value != form1.password2.value) {
                            alert("Password and Password Confirmation Did Not Match");
                            form1.password.focus()
                            return false; }
                           if (form1.emailadd.value == "") {
                            alert("Must Fill in Email Address");
                            form1.emailadd.focus()
                            return false; }  
                           
                       }
//-->
</SCRIPT>
<!--Header Begin -->
<link rel="stylesheet" href="sfCSS.css" type="text/css">
</head>

<body  bgproperties="fixed" link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">
<form method="post" action="mailsubscribe.asp"  id="form1" name="form1" LANGUAGE=javascript onsubmit="return form1_onsubmit()">
	<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="<%= C_WIDTH %>" align="center">
      <tr>
        <td>
  
            <table width="100%" border="0" cellspacing="1" cellpadding="3">
              <tr>
			          
            <td align="middle"  class="tdTopBanner"> 
              <%If C_BNRBKGRND = "" Then%>
              <%Else%>
              <img src="buttons/tt_blue.gif" border="0" width="275" height="36"> 
              <%End If%>
            </td>
              </tr>
<!--Header End -->
              <tr>
	            <td align="middle"   class="tdMiddleTopBanner">Mailing List</td>        
              </tr>
              <tr>
                <td class="tdBottomTopBanner">Complete
                  the form below to subscribe to our mailing list.&nbsp;
                  Subscribers will be able to receive store newsletters, sale
                  announcements and other mailings of interest.</td>
              </tr>
              <tr>
                <td class="tdContent2" width="100%" nowrap>
                  <table border="0" width="100%" cellpadding="4" cellspacing="0">        
                    <% If sEmail <> "" Then %>
                    <tr>
                      <td width="100%" colspan="2" align="center" height="90" valign="center">
			            <table width="60%" cellpadding="1" cellspacing="0" class="tdbackgrnd">
			              <tr><td width="100%">
				              <table cellpadding="5" cellspacing="0" class="tdContent" width="100%">
				                <tr>
                              <td width="100%" class="tdContent3" align="center">
                                <font class='MSFont30'><b>You have been added 
                                to the Mailing List.</b></font></td>
                            </tr>
				              </table>
			                </td></tr>	
			            </table>
                      </td>
                    </tr>
		            <% End If %>
                    
                    <tr>
                      <td width="40%" align="right">First Name:</td>
                      <td width="60%"><input type="text" value="<%= sFirstname %>" name="fname" size="40" style="<%= C_FORMDESIGN %>"></td>
                    </tr>
                    <tr>
                      <td width="40%" align="right" valign="top">Last Name :</td>
                      <td width="60%"><input type="text" value="<%= slastname %>" name="lName" size="40" style="<%= C_FORMDESIGN %>">
                      </td>
                    </tr>
                    <tr>
                      <td width="40%" align="right">Password:</td>
                      <td width="60%"><input type="password" value="<%= spassword %>" name="password" size="40" style="<%= C_FORMDESIGN %>"></td>
                    </tr>
                    <tr>
                      <td width="40%" align="right">Confirm Password:</td>
                      <td width="60%"><input type="password" value="<%= sconfirmpass %>" name="password2" size="40" style="<%= C_FORMDESIGN %>"></td>
                    </tr>
                    <tr>
                      <td width="40%" align="right" valign="top">E-Mail Address:</td>
                      <td width="60%"><input type="text" value="<%= sEmail %>" name="emailadd" size="40" style="<%= C_FORMDESIGN %>">
                      </td>
                    </tr>
                    <tr>
                      <td width="100%" align="center" valign="top" colspan="2"><input type="image" name="Submit" border="0" src="<%= C_BTN18 %>" ></td>
                    </tr>     
                  </table>
	          </td>
	        </tr>
<!--Footer begin-->
              <!--#include file="footer.txt"-->
          </table>
	          </td>
	        </tr>
          </table>
</form>
</body>
</html>
<!--Footer End-->
<%
closeObj(cnn)
%>



