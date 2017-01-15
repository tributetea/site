<%@ Language=VBScript %>
<%	option explicit 
Response.Buffer=True
%>
<!--#include file="../SfLib/SFsecurity.asp"-->
<!--#include file="../SFLib/db.conn.open.asp"-->
<!--#include file="../SFLib/adovbs.inc"-->
<%
'--------------------------------------------------------------------
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: printEchechk.asp
	 

'

'@DESCRIPTION:   prints an echeck

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws  and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO	


Dim iOrder,sSQL, rsEcheck
iOrder = Request.QueryString("orderid")

Set rsEcheck = Server.CreateObject("ADODB.Recordset")
sSQL = "SELECT custFirstName, custLastName, custAddr1, custAddr2, custCity, custState, custCountry, custZip, orderCheckAcctNumber, orderCheckNumber, orderBankName, orderRoutingNumber," _
	& " orderDate, orderGrandTotal FROM sfCustomers INNER JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId WHERE orderID = " & iOrder
rsEcheck.Open sSQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>
<HTML>
<HEAD>

<TITLE>Print ECheck</TITLE>
<!--Header Begin Modified (No store name)-->
<link rel="stylesheet" href="../sfCSS.css" type="text/css">
</head>

<body bgproperties="static" link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">

                
<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="<%= C_WIDTH %>" align="center">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="3">
        
<!--Header End -->
		<tr>
		<td align="center"><font class='MSFont21'>Payment Authorized by Account Holder. Indemnification Agreement Provided by Depositor.</font></td>
		</tr>	
		<tr><td>
			<table border="0" cellpadding="4" cellspacing="0" width="100%" align="center">
			<tr>
				<td align="left"><font class='MSFont21'><b><%= rsEcheck.Fields("custFirstName") & " " & rsEcheck.Fields("custLastName") %></b></font></td>
				<td align="right"><font class='MSFont21'><b>Check Number: </b><%= rsEcheck.Fields("orderCheckNumber") %></font></td>
			</tr>
			<tr>
				<td align="left" colspan="2"><font class='MSFont21'><%= rsEcheck.Fields("custAddr1") %>
				<br><%= rsEcheck.Fields("custAddr2") %></font></td>
			
			<tr>
				<td align="left" ><font class='MSFont21'><%= rsEcheck.Fields("custCity")%>&nbsp;<%= rsEcheck.Fields("custState") %>, <%= rsEcheck.Fields("custZip") %></font></td>
				<td align="right"><font class='MSFont21'><b>Date: </b><u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u></font></td>
			</tr>
			<tr>	
				<td align="left" colspan="2"><font class='MSFont21'><%= rsEcheck.Fields("custCountry") %></font></td>						
			</tr>
			<tr>
				<td align="left" colspan="2" height="10"></td>
			</tr>
			<tr>
			<td align="left"><b><font class='MSFont21'>Pay the amount of : </font></b></td><td align="right"><b><font class='MSFont21'><%= FormatCurrency(rsEcheck.Fields("orderGrandTotal"))%>&nbsp;&nbsp;&nbsp;</font></b>
			    
			</td></tr>
<td valign="top" colspan="2"><hr size="1" width="100%" color="#445566" align="center"></td>
</tr>
			<tr>
				<td width="50%" height="20">&nbsp;&nbsp;&nbsp;<font class='MSFont21'><b><%= rsEcheck.Fields("orderBankName") %></b></font></td>
				<td align="right"><font class="MSFont21">									    
				Electronically Signed By: <b><%= rsEcheck.Fields("custFirstName") & " " & rsEcheck.Fields("custLastName") %>&nbsp;&nbsp;&nbsp;</b></font>
				<hr size="1" width="100%"></td>
			</tr>	
			<tr>
			<td colspan="2" align="left"><b><font face="MICR 013 BT" size="5"><b>A<%= rsEcheck.Fields("orderRoutingNumber") %>A<%= rsEcheck.Fields("orderCheckAcctNumber") %>B</b></font></b></td>
			</tr>
			</table>
	       </td></tr>
	      <!--Footer begin Modified (no links)-->
                
              </table>
            </td>
          </tr>
        </table>
       </body>

    </html>
<!--Footer End-->












