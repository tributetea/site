<%@ Language=VBScript %>
<%	option explicit 
	Response.Buffer = True

%>

<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="SFLib/incProcOrder.asp"-->
<!--#include file="error_trap.asp"-->
<!--#include file="SFLib/incDesign.asp"-->
<!--#include file="SFLib/incGeneral.asp"-->
<!--#include file="SFLib/ADOVBS.inc"-->
<!--#include file="SFLib/incAE.asp"-->

<%
    Const vDebug = 0
	'@BEGINVERSIONINFO

	'@APPVERSION: 50.4011.0.2

	'@FILENAME: process_order.asp
		 

	
	'@DESCRIPTION: Gathers Information For order.

	'@STARTCOPYRIGHT
	'The contents of this file is protected under the United States
	'copyright laws  and is confidential and proprietary to
	'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
	'expressed written permission of LaGarde, Incorporated is expressly prohibited.
	'
	'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
	'@ENDCOPYRIGHT

	'@ENDVERSIONINFO
dim tmpOrderQty
If Application("AppName") = "StoreFrontAE" Then 'SFAE
	tmpOrderQty=GetTotalOrderQTY
else
	tmpOrderQty=GetTotalOrderQTYSE
end if
if isnumeric(tmpOrderQty)=false or tmpOrderQty <= 0 then
'If Session("SessionID") = "" Then 'SFUPDATE
	Session.Abandon
	Response.Redirect(C_HomePath & "search.asp")		
End If
If Application("AppName") = "StoreFrontAE" Then 'SFAE
	Confirm_CheckCartAndRedirect 
End IF

Dim sSql,sEmail,iOrderID,rsMyOrders, sPassword, iAuthenticate, bLoggedIn, sCondition, sPaymentMethod, sPaymentList
Dim sProdID,iQuantity,iShip,sTotalPrice,strProdID,intQuantity,rstAdmin,sTotalCost,blnFree, iFreeShip
' initially false
bLoggedIn = false
iShip = 0
sSQL = "SELECT * FROM sfTmpOrderDetails WHERE odrdttmpShipping >= 1 AND odrdttmpSessionID = " & Session("SessionID")
Set rsMyOrders = Server.CreateObject("ADODB.RecordSet")
Set rstAdmin = Server.CreateObject("ADODB.RecordSet")
rsMyOrders.Open sSQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
rstAdmin.Open "sfAdmin", cnn, adOpenForwardOnly, adLockReadOnly, adCmdTable


If Request.QueryString("Persist") = 1 then '#245
   sTotalPrice= Session("persistTotalPrice")
else
  sTotalPrice =Request.Form("TotalPrice")
end if  

Session("persistTotalPrice") = sTotalPrice
If NOT rsMyOrders.EOF = True AND NOT rsMyOrders.EOF = true Then
rsMyOrders.MoveFirst 
iShip = 0
Do While NOT rsMyOrders.EOF
 iShip = iShip + rsMyOrders.Fields("odrdttmpShipping")
  rsMyOrders.MoveNext 
loop
End If

blnFree =False
if iShip <> 0 then
iFreeShip = rstAdmin.Fields("adminFreeShippingIsActive") 
  if (cdbl(sTotalPrice) > cdbl(rstAdmin.Fields("adminFreeShippingAmount")) AND iFreeShip = "1")  then
    blnFree =True
  end if
elseif iShip = 0 then
   blnFree =True
     else
  blnFree =False
     end if  
rsMyOrders.Close 
rstAdmin.Close 
 set rstAdmin =nothing
 set rsMyOrders = nothing    

'-------------------------------------------------------
' See if session is repeating, if so, give new id to use
'-------------------------------------------------------
If Session("SessionID") = Request.Cookies("EndSession") Then	
	bLoggedIn = false
End If	

'-------------------------------------------------------
' Check if custID exists 
'-------------------------------------------------------
iCustID = Session("custID")
If iCustID <> "" Then
	 Dim bCustIdExists
	   	bCustIdExists = CheckCustomerExists(iCustID)
    	If bCustIdExists = false Then
    		Response.Cookies("sfCustomer")("custID") = ""
	   		Response.Cookies("sfCustomer").Expires = NOW()
	   	Else
			Response.Cookies("sfCustomer")("custID") = iCustID
			Response.Cookies("sfCustomer").Expires = Date() + 730
		End If
End If	
	
If Request.Cookies(Session("GeneratedKey") & "sfOrder")("SessionID") = Session("SessionID") AND Request.Cookies(Session("GeneratedKey") & "sfOrder")("SessionID") <> ""  AND bCustIdExists Then
	bLoggedIn = true
End If

' Get Payment List
	sPaymentList = getPaymentMethods()

'-------------------------------------------------------
' If login button is depressed
'-------------------------------------------------------
If Trim(Request.Form("btnLogin.x")) <> "" Then
	
	sEmail			= Trim(Request.Form("Email"))
	sPassword		= Trim(Request.Form("Passwd"))
	
	' Authenticate
	iAuthenticate	= customerAuth(sEmail,sPassword,"loose")
		
	If iAuthenticate > 0 Then
		If Request.Cookies("sfCustomer")("custID") <> "" AND iAuthenticate <> Request.Cookies("sfCustomer")("custID")  Then
			Dim bSvdCartCust
			bSvdCartCust = CheckSavedCartCustomer(Request.Cookies("sfCustomer")("custID"))
			'Response.write "Saved Cart Cust?" & Request.Cookies("sfCustomer")("custID") & "False?" & bSvdCartCust 
			If bSvdCartCust Then
				' Delete SvdCartCustomer Row
				Call DeleteCustRow(Request.Cookies("sfCustomer")("custID"))
				' See if saved cart has any remaining saved
				Call setUpdateSavedCartCustID(iAuthenticate,Request.Cookies("sfCustomer")("custID"))
			End If
		End If	
		Response.Cookies(Session("GeneratedKey") & "sfOrder")("SessionID") = Session("SessionID")
		Response.Cookies(Session("GeneratedKey") & "sfOrder").Expires = Date() + 1
		Response.Cookies("sfCustomer")("custID") = iAuthenticate
		Response.Cookies("sfCustomer").Expires = Date() + 730
		Session("custID") = iAuthenticate
		bLoggedIn = true
		iCustID = iAuthenticate
	Else 	
		If customerAuth(sEmail,sPassword,"loosest") > 0 Then
			sCondition = "EmailMatch"   
			Response.Cookies("sfCustomer").Expires = Now()
		Else 
			sCondition = "WrongCombination"
			Response.Cookies("sfCustomer").Expires = Now()		
		End If			
	End If			
End If		
%>

<html>
<head>
	
<script language="javascript">
<!--
function clearShipping(form)
 {
for (var i=0; i < form.length; i++) 
   {
	 e = form.elements[i];
	 if (e.name.indexOf("Ship") == 0) 
	 {
			e.value = "";
	 }
   } 
}
function validate_Me(frm)
{
 var bmain_is_good = sfCheck(frm);
 if(bmain_is_good == true)
   {
    var bshipping_is_good = sfCheckPlus(frm);
    return bshipping_is_good;
   }
     
 else
  {
  return false;
  } 
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<SCRIPT language="javascript" src="SFLib/sfCheckErrors.js"></SCRIPT>


<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>TRIBUTE TEA CHECKOUT - 1. Customer Information</title>


<!--Header Begin -->
<link rel="stylesheet" href="sfCSS.css" type="text/css">
</head>


<body  bgproperties="fixed"  link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" text="#666666">
<tr>
    <td>
      
    <table width="618" border="0" cellspacing="1" cellpadding="2" align="center">
      <tr> 
        <td valign="top" align="middle" height="46" width="612"  class="tdTopBanner"> 
          <div align="center"><img src="buttons/tt_blue.gif" width="275" height="40"> 
          </div>
        </td>
      </tr>
      <tr> 
        <td height="62"  valign="center" align="center"  class="tdContent2"><font size="1"><img src="ssl.gif" width="612" height="25"><br>
          <br>
          <img src="step1.gif" width="540" height="12"></font></td>
      </tr>
      <!--Header End -->
      <tr valign="center"> 
        <td class="tdContent2" align="center" valign="center"> 
          <% If Not bLoggedIn Then %>
          <form action="process_order.asp" method="post" name="frmEmail">
            <input type="hidden" name="FreeShip" value="<%= blnFree %>">
            <input type="hidden" name="FreeShip" value="<%= blnFree %>">
            <table border="0" width="97%" cellpadding="5" cellspacing="1">
              <tr> 
                <td width="50%" align="center" class="tdContent2" valign="center"> 
                  <table border="0" class="tdBottomTopBanner2" width="97%" cellpadding="3" cellspacing="1">
                    <tr> 
                      <td width="100%" align="center" class="tdBottomTopBanner"><font face="Arial, Helvetica, sans-serif"><b>Returning 
                        Customer Login</b></font></td>
                    </tr>
                    <tr> 
                      <td align="center" valign="center" class="tdContent2"> 
                        <table border="0" width="97%" cellpadding="2">
                          <tr> 
                            <td width="32%" align="right"><b> <font face="Arial, Helvetica, sans-serif">e-mail:</font></b></td>
                            <td width="68%"> 
                              <input type="text" size="25" name="Email"  title="E-Mail Address" style="<%= C_FORMDESIGN %>" maxlength="50">
                            </td>
                          </tr>
                          <tr> 
                            <td width="32%" align="right"><b><font face="Arial, Helvetica, sans-serif">Password:</font></b></td>
                            <td width="68%"> 
                              <input type="password" size="25" name="Passwd" title="Password" style="<%= C_FORMDESIGN %>" maxlength="50">
                            </td>
                          </tr>
                          <tr> 
                            <td align="middle" colspan="2"> 
                              <div align="center"> 
                                <input Type="image" src="buttons/submit.gif" name="btnLogin" border="0" width="92" height="22">
                              </div>
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                    <tr> 
                      <td valign="top" height="13"> 
                        <div align="center"><a href="password.asp?status=fpwd"><font face="Arial, Helvetica, sans-serif" size="2">FORGOT 
                          PASSWORD?</font></a> <font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2">|</font> 
                          <font face="Arial, Helvetica, sans-serif" size="2"><a href="password.asp?status=change">CHANGE 
                          LOGIN</a></font><font face="Arial, Helvetica, sans-serif"><a href="password.asp?status=change"> 
                          </a></font></div>
                      </td>
                    </tr>
                  </table>
                </td>
                <td width="50%" class="tdContent2"> 
                  <center>
                    <font class="Error"><b> 
                    <% If sCondition = "EmailMatch" or sCondition = "WrongCombination" Then %>
                    <font color="#990000" face="Arial, Helvetica, sans-serif">Login 
                    Failed</font> <font face="Arial, Helvetica, sans-serif"> 
                    <% Else %>
                    <font color="#990000">Login Directions</font></font> <font face="Arial, Helvetica, sans-serif"> 
                    <% End If %>
                    </font></b></font> 
                    <hr width="90%" noshade size="1">
                  </center>
                  <font face="Arial, Helvetica, sans-serif"> 
                  <% If sCondition = "EmailMatch" Then %>
                  Your combination was wrong, but an e-mail match was found. Please 
                  login with the correct password or if you wish to open a new 
                  account, you must choose a new password. 
                  <% ElseIf sCondition = "WrongCombination" Then %>
                  Your e-mail and password combination is incorrect. Try again. 
                  <% Else %>
                  Please use your e-mail address and password to log in and retrieve 
                  your customer information.</font> 
                  <% End If %>
                </td>
              </tr>
            </table>
            <input type="hidden" name="TotalPrice" value="<%= sTotalPrice %>">
            <input type="hidden" name="FreeShip" value="<%= blnFree %>">
            <input type="hidden" name="bShip" value="<%= iShip %>">
          </form>
          <% End If %>
          <!--#include file="orderform.asp"-->
      <tr> 
        <td height="63" valign="top" bgcolor="#003366"> 
          <div align="center"><a href="http://www.tributetea.com/index.html"><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">HOME</font></a><font color="#FFFFFF" face="Arial, Helvetica, sans-serif"> 
            | <a href="http://www.tributetea.com/teas.asp">TEAS</a> | <a href="http://www.tributetea.com/herbs.asp">HERBS</a> 
            | <a href="http://www.tributetea.com/teaware.asp">TEAWARE</a> | <a href="http://www.tributetea.com/incense.asp">INCENSE</a> 
            | <a href="http://www.tributetea.com/books.asp">BOOKS &amp; MUSIC</a> 
            | <a href="http://www.tributetea.com//sale.asp">SPECIALS</a><br>
            <a href="http://www.tributetea.com/teaschool.asp">TEA SCHOOL</a> | 
            <a href="http://www.tributetea.com/news.asp">NEWS</a> | <a href="http://www.tributetea.com/about.asp">ABOUT 
            TT</a> | <a href="http://www.tributetea.com/faq.asp">Q &amp; A</a> 
            | <a href="http://www.tributetea.com/wholesale.asp">WHOLESALE</a> 
            | <a href="http://www.tributetea.com/contact.asp">CONTACT</a> | <a href="http://www.tributetea.com/search.asp">SEARCH</a></font><font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><font size="1" onClick="MM_openBrWindow('privacy_security.html','security','scrollbars=yes,width=425,height=425')"><b><font size="1"><font color="#CC0000" onClick="MM_openBrWindow('file:///C|/tribute_tea/ssl/privacy_security.html','security','scrollbars=yes,width=425,height=425')"><br>
            <br>
            </font><font color="#666600"><font color="#990000"><b><font size="-1" color="#FF0000">&copy; 
            2002-6 TT. All rights reserved.</font></b></font></font></font></b></font></font></div>
        </td>
      </tr>
    </table>
</body>

    </html>