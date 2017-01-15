<%@ Language=VBScript %>
<%
option explicit
Response.Buffer = True
%>
<!--#include file="../SFLib/sfSecurity.asp"-->
<!--#include file="../SFLib/incDesign.asp"-->
<!--#include file="../SFLib/adovbs.inc"-->
<!--#include file="../SFLib/db.conn.open.asp"-->
<!--#include file="../SFLib/incGeneral.asp"-->
<%
'--------------------------------------------------------------------
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: sfreports2.asp
	 
'Access Version
'

'@DESCRIPTION:   web reporting tool

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws as an unpublished work, and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO	

%>
<HTML>
<HEAD>

<title>SF Reports Page</title>


<%
Dim sStartDate, sEndDate, rsSales, sSQL, sTotalNet, sTotalSTax, sTotalCTax,sTotalShipping, sGrandTotal, arrSales, sHandling, i

sStartDate = MakeUSDate(Request.QueryString("startDate"))
sEndDate = MakeUSDate(Request.QueryString("endDate"))

Set rsSales = Server.CreateObject("ADODB.RecordSet")
sSQL = "SELECT orderAmount, orderSTax, orderCTax, orderShippingAmount, orderHandling, orderGrandTotal FROM sfOrders WHERE orderDate BETWEEN #" & sStartDate & "# AND #" & sEndDate & "#  and orderIsComplete = 1"
rsSales.Open sSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText

If Not (rsSales.EOF and rsSales.BOF) Then arrSales = rsSales.GetRows 

sTotalNet = 0 
sTotalSTax = 0
sTotalCTax = 0
sTotalShipping = 0
sHandling = 0 
sGrandTotal = 0

rsSales.Close 
Set rsSales= Nothing

If isArray(arrSales) Then
	For i=0 to uBound(arrSales, 2)
		If arrSales(0, i) <> "" Then sTotalNet = sTotalNet + cDbl(arrSales(0, i))
		If arrSales(1, i) <> "" Then sTotalSTax = sTotalSTax + cDbl(arrSales(1, i))
		If arrSales(2, i) <> "" Then sTotalCTax = sTotalCTax + cDbl(arrSales(2, i))
		If arrSales(3, i) <> "" Then sTotalShipping = sTotalShipping + cDbl(arrSales(3, i))
		If arrSales(4, i) <> "" Then sHandling = sHandling + cDbl(arrSales(4, i))
		If arrSales(5, i) <> "" Then sGrandTotal = sGrandTotal + cDbl(arrSales(5, i))
	Next
End If 

%>
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
    <tr>
	<td align="middle" class="tdMiddleTopBanner">Sales Summary</td>        
    </tr>
    <tr>
	<td class="tdBottomTopBanner">The total
      of all orders placed within the date range you specified are listed below.</td>    
    </tr>
    <tr>
    <td class="tdContent2" width="100%" nowrap>
        <table border="0" width="100%" cellpadding="4" cellspacing="0">
        
        <% If isArray(arrSales) Then %>
        <tr>
		<td width="100%" class="tdContentBar" colspan="4">Report for <%= sStartDate %> to <%= sEndDate %></td>        
        </tr>
        <tr>
        <td width="75%" align="right" valign="top">Net Sales:</td>
        <td width="25%" align="left" valign="top"><%= FormatCurrency(sTotalNet) %></td>
        </tr>  
        <tr>
        <td width="75%" align="right" valign="top">Total State/Province Tax:</td>
        <td width="25%" align="left" valign="top"><%= FormatCurrency(sTotalSTax) %></td>
        </tr> 
        <tr>
        <td width="75%" align="right" valign="top">Total Country Tax:</td>
        <td width="25%" align="left" valign="top"><%= FormatCurrency(sTotalCTax) %></td>
        </tr>  
        <tr>
        <td width="75%" align="right" valign="top">Total Shipping:</td>
        <td width="25%" align="left" valign="top"><%= FormatCurrency(sTotalShipping) %></td>
        </tr>                 
        <tr>
        <td width="75%" align="right" valign="top">Total Handling:</td>
        <td width="25%" align="left" valign="top"><%= FormatCurrency(sHandling) %></td>
        </tr>   
        <tr>
        <td width="75%" align="right" valign="top">Total Sales:</td>
        <td width="25%" align="left" valign="top"><%= FormatCurrency(sGrandTotal) %></td>
        </tr>
        <% Else %>
        <tr>
				<td colspan=4 align="center" class="tdContent2"><font class="Error">There Were No Sales Between <%= sStartDate %> And <%= sEndDate %></font></td>
				</tr>
        <% End If %>
        <tr>
        <td width="100%" align="center" valign="top" colspan="4"></td>
        </tr>
        </table>
    </td>
    </tr>
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



