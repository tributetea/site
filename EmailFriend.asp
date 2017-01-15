<%@ Language=VBScript %>
<%	option explicit 
	Response.Buffer = True
	%>
<!--#include file="SFLib/incDesign.asp"-->
<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file="SFLib/mail.asp"-->
<!--#include file="SFLib/incGeneral.asp"-->
	<%
		
	'@BEGINVERSIONINFO


	'@APPVERSION: 50.4011.0.2

	'@FILENAME: emailfriend.asp
 

	'

	'@DESCRIPTION: Send Email of product to a friend

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

<HTML>

<HEAD>
<SCRIPT LANGUAGE=javascript>
<!--
function chkEMail() {
if(frmEmailfriend.txtFriend2.value != ""){  
  var emailadd = frmEmailfriend.txtFriend2.value
  var stemp1 = emailadd.indexOf("@");
  var stemp2 = emailadd.indexOf(".");  
    if((stemp1 < 1 )||(stemp2 < 1))
     {
       alert("Must Use a valid friend 2 email Address");
        frmEmailfriend.txtFriend2.focus();
      }
       
 // }
  
  }
}
//-->
</SCRIPT>
<SCRIPT language=javascript src="SFLib/sfCheckErrors.js"></SCRIPT>

<meta http-equiv="Pragma" content="no-cache">

<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%= C_STORENAME %>-SF Email a Friend Popup Window</title>





<%
Dim sClosing
sClosing = ""

If Request.Form("btnEmailSubmit.x") <> "" Then
	Dim sInformation 
	sInformation = Request.Form("txtFriend") & "|" & Request.Form("txtEmail") & "|" & Request.Form("txtMessage")& "|" & Request.Form("prodID") & "|" & Request.Form("txtSubject") & "|" & Request.Form("txtFriend2")
	'Response.Write sInformation
	Call createMail ("EmailFriend",sInformation)
	sClosing = "onload=""javascript:window.close()"""
End If
%>

<!--Header Begin Modified (Body tag is modified)-->
<link rel="stylesheet" href="sfCSS.css" type="text/css">
</HEAD>
<body <% If Request.Form("btnEmailSubmit.x") <> "" Then response.write sClosing %>  link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">

<form method="post" name="frmEmailfriend" onsubmit="javascript:this.txtFriend2.optional=true;return sfCheck(this)">  
<table border="0" cellpadding="1" cellspacing="0"  class="tdbackgrnd" width="100%" align="center">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="3">
        <tr>
          <td align="middle"  class="tdTopBanner"><%= C_STORENAME %></td>
        </tr>
<!--Header End -->
        <tr>
          <td class="tdContent2">
<table border=0 width="100%">
    <tr>
      <td colspan=2 nowrap><b>Use this form to e-mail information about this product to others.</b>
        &nbsp;<br>
        <br>
      </td>
    </tr>
    <tr>
      <td nowrap width="20%">To
        E-Mail Address:</td>
      <td nowrap><input type=text size=40 name="txtFriend" title="Your Firend's Email Address"></td>
    </tr>
    <tr>
      <td nowrap width="20%">CC
        E-Mail Address:</td>
      <td nowrap><input type=text size=40 name="txtFriend2"   onfocusout="chkEMail()"></td>
    </tr>
    <tr>
      <td nowrap width="20%">From
        E-Mail Address:</td>
      <td nowrap><input type=text size=40 name="txtEmail" title="Your Email Address"></td>
    </tr>
    <tr>
      <td nowrap width="20%">Subject:</td>
      <td nowrap><input type=text size=50 name="txtSubject" title="Email Subject"></td>
    </tr>
    <tr>
      <td colspan=2 nowrap>Message:</td>
    </tr>
    <tr>
      <td colspan=2 align=center nowrap><textarea cols=70 rows=7 name="txtMessage" title="Message">Hello, <%=vbcrlf%>This link leads to <%= C_STORENAME%> where you'll find this product.</textarea></td>
    </tr>
    <tr>
      <td colspan=2 align=center nowrap><input type=image border="0" src="<%= C_BTN18 %>" name="btnEmailSubmit"></td>
    </tr>
  <input type=hidden name=prodID value="<%= Request.QueryString("ProdID") %>">
</table>
<!--Footer Begin -->
            <tr>
              <td class="tdFooter"><p align="center"><b><a href="javascript:window.close()">Close</a></b></td>
                </tr>
                </table>
              </td>
            </tr>
            </table>
</form>
</BODY>
</html>
<!--Footer End -->

<%
closeObj(cnn)
%>



