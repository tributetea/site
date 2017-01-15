<%@ Language=VBScript %>
<%	option explicit 
Response.Buffer = True
%>
<!--#include file="../SfLib/incDesign.asp"-->
<!--#include file="../SFLib/db.conn.open.asp"-->
<!--#include file="../SFLib/adovbs.inc"-->
<!--#include file="../SFLib/incGeneral.asp"-->
<!--#include file="../SFLib/SFSecurity.asp"-->
<%
'--------------------------------------------------------------------
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: sfreports6.asp
	 

'

'@DESCRIPTION:   web reporting tool

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws  and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO	

%>
	<html>
	<head>

	<title>SF Reports Page</title>
	<!--Header Begin -->
<link rel="stylesheet" href="../sfCSS.css" type="text/css">
</head>

<body bgproperties="static" link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">

                
<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="<%= C_WIDTH %>" align="center">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="3">
        <tr>
          <td align="middle"  class="tdTopBanner"><%If C_BNRBKGRND = "" Then%><%= C_STORENAME %><%Else%><img src="../<%= C_BNRBKGRND %>" border="0"><%End If%></td>
        </tr>
<!--Header End -->
<%

If Request.Form("btnSubmit.x") <> "" Then
	Dim sOrderId, sSQL, rsOrders, sFirstName, sLastName, iCounter, sBgColor, sFontFace, sFontColor, sFontSize

	sOrderId = Request.Form("OrderID")
	sFirstName = Request.Form("FirstName")
	sLastName = Request.Form("LastName")

	sSQL = "Select custID, custFirstName, custLastName, custMiddleInitial, orderID, orderCustID, orderDate, orderGrandTotal " _
		   & "FROM sfCustomers INNER JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId WHERE orderID LIKE '%" & sOrderId & "%'" _
		   & " AND custFirstName LIKE '%" & sFirstName & "%' AND custLastName LIKE '%" & sLastName & "%' and sfOrders.orderIsComplete = 1 Order By orderID"  
	Set rsOrders = Server.CreateObject("ADODB.RecordSet")
	rsOrders.Open sSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText
	%>

	    <form method="post" id="form1" name="form1">
	    <tr>
		<td align="middle" class="tdMiddleTopBanner">Transaction Details</td>        
	    </tr>
	    <tr>
		<td class="tdBottomTopBanner"><b>Instructions: </b>This reporting tool will allow you to retrieve a detailed report for a single order.  Enter the order ID if you know it, or enter the customer's first or last name.  All matches will be displayed, and you will be able to select the one you are looking for.</td>    
	    </tr>
	    <tr>
	    <td class="tdContent2" width="100%" nowrap>
	        <table border="0" width="100%" cellpadding="4" cellspacing="0">
	        <tr>
	        <td width="15%" align="center" class="tdContentBar">Order Number</td>        
			<td width="25%" align="center" class="tdContentBar">Date</td>        
			<td width="35%" class="tdContentBar">Customer Name</td>        
			<td width="25%" align="center" class="tdContentBar">Order Total</td>        
	        </tr>
	        <%
	        If rsOrders.EOF Then
	        %>
				<tr>
				<td colspan="4" align="center" class="tdContent2"><font class="MSFont30">There Were no Orders for your Search Criteria</font></td>
				</tr>
	        <%
	        Else
				iCounter = 1
				Do While Not rsOrders.EOF 
					If iCounter mod 2 = 0 Then
						sBgColor = C_ALTBGCOLOR1
						sFontFace = C_ALTFONTFACE1
						sFontColor = C_ALTFONTCOLOR1
						sFontSize = C_ALTFONTSIZE1
					Else
						sBgColor = C_ALTBGCOLOR2
						sFontFace = C_ALTFONTFACE2
						sFontColor = C_ALTFONTCOLOR2
						sFontSize = C_ALTFONTSIZE2
					End If
	        %>
					<tr>
					<td align="center" valign="top" bgcolor="<%= sBgColor %>"><font face="<%= sFontFace %>" color="<%= sFontColor %>" SIZE="<%= sFontSize %>"><a href="sfReports1.asp?OrderID=<%= rsOrders.Fields("orderID") %>"><%= rsOrders.Fields("orderID") %></a></font></td>
					<td align="center" valign="top" bgcolor="<%= sBgColor %>"><font face="<%= sFontFace %>" color="<%= sFontColor %>" SIZE="<%= sFontSize %>"><%= rsOrders.Fields("orderDate")%></font></td>
					<td bgcolor="<%= sBgColor %>"><font face="<%= sFontFace %>" color="<%= sFontColor %>" SIZE="<%= sFontSize %>"><%= rsOrders.Fields("custFirstName") %>&nbsp;<%= rsOrders.Fields("custMiddleInitial") %>&nbsp;<%= rsOrders.Fields("custLastName") %></font></td>
					<td align="center" valign="top" bgcolor="<%= sBgColor %>"><font face="<%= sFontFace %>" color="<%= sFontColor %>" SIZE="<%= sFontSize %>"><%= FormatCurrency(rsOrders.Fields("orderGrandTotal")) %></font></td>
					</tr>
	        <%
					iCounter = iCounter + 1
					rsOrders.MoveNext 
				Loop
			End If
	        %>
	        <tr>
	        <td width="100%" align="center" valign="top" colspan="3"></td>
	        </tr>
	        </table>
	    </td>
	    </tr></form>
	<%
	rsOrders.Close 
	Set rsOrders = nothing
	cnn.Close
	Set cnn = nothing
	%>

<% Else %>

	    <tr>
		<td align="middle" class="tdMiddleTopBanner">
        Retrieve Order</td>        
	    </tr>
	    <tr>
		<td class="tdBottomTopBanner"><b>Instructions: </b>The following orders match your criteria.  To view the specifics of a single order, click that record's order ID.</td>    
	    </tr>
	    <tr>
	    <td class="tdContent2" width="100%" nowrap>
	        <form method="post" name="frm1">
	        <table border="0" width="100%" cellpadding="4" cellspacing="0">
	        <tr>
	        <td align="right">Order ID</td>
	        <td align="left">
            <input type="text" style="<%= C_FORMDESIGN %>" name="OrderID"></td>
	        </tr>
	        <tr>
	        <td align="right">First Name</td>
	        <td align="left">
            <input type="text" style="<%= C_FORMDESIGN %>" name="FirstName"></td>
	        </tr>
	        <tr>
	        <td align="right">Last Name</td>
	        <td align="left">
            <input type="text" style="<%= C_FORMDESIGN %>" name="LastName"></td>
	        </tr>
	        <tr>
	        <td width="100%" align="center" valign="top" colspan="2"></td>
	        </tr>
	        <tr>
	        <td colspan="2" width="100%" align="center" valign="top"><input type="image" name="btnSubmit" border="0" src="../<%= C_BTN18 %>" alt="Submit" WIDTH="108" HEIGHT="21"></td>
	        </tr>
	        </table>
		</form>
	    </td>
	    </tr>
	  
<% End If %>


     <!--Footer begin-->
                <tr>
		<td class="tdFooter"><p align="center"><font class="Footer"><b><a href="sfReports.asp">Reports</a> | <a href="<%= C_HOMEPATH %>"><%= C_STORENAME %></a></b></font></p></td>
	    </tr>
              </table>
            </td>
          </tr>
        </table>
       </body>

    </html>
<!--Footer End-->






