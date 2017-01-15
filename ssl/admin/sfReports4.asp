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

'@FILENAME: sfreports4.asp
	
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
Dim sStartDate, sEndDate, rsSales, sSQL, sTotalNet, sTotalSTax, sTotalCTax, sTotalShipping, sGrandTotal, arrSales, sHandling, i, rsPartners, j, sAffiliate, rsAff, sFilter

sStartDate = MakeUSDate(Request.QueryString("startDate"))
sEndDate = MakeUSDate(Request.QueryString("endDate"))
sAffiliate = Request.QueryString("Affiliate")

If sAffiliate = "" Then
	Set rsSales = Server.CreateObject("ADODB.RecordSet")
	sSQL = "SELECT orderAmount, orderSTax, orderCTax, orderShippingAmount, orderHandling, orderGrandTotal, orderTradingPartner FROM sfOrders WHERE (orderDate BETWEEN #" & sStartDate & "# AND #" & sEndDate & "#) AND (orderTradingPartner IS NOT NULL) and orderIsComplete = 1 and orderTradingPartner in (select affName from sfAffiliates)"
	rsSales.Open sSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText

	If Not (rsSales.EOF and rsSales.BOF) Then arrSales = rsSales.GetRows 

	closeObj(rsSales)
	
	sSQL = "SELECT DISTINCT orderTradingPartner FROM sfOrders WHERE orderTradingPartner <> '' and orderIsComplete = 1 and orderTradingPartner in (select affName from sfAffiliates)"
	Set rsPartners = Server.CreateObject("ADODB.Recordset")
	rsPartners.Open sSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText
	%>


	    <tr>
		<td align="middle" class="tdMiddleTopBanner">Affiliate Sales Summary</td>        
	    </tr>
	    <tr>
		<td class="tdBottomTopBanner">Referral sales for all affiliate partners are listed below.  Chose an <B>Affiliate ID</B> to view that affiliate's transaction detail.</td>    
	    </tr>
	    <tr>
	    <td class="tdContent2" width="100%" nowrap>
	        <table border="0" width="100%" cellpadding="4" cellspacing="0">
	        
	        <% 
	        If isArray(arrSales) AND Not (rsPartners.EOF And rsPartners.BOF) Then 
	        %>
	        <tr>
			<td width="100%" class="tdContentBar" colspan="4">Report for <%= sStartDate %> to <%= sEndDate %></td>        
	        </tr>
	   
	        <%
				Do While Not rsPartners.EOF
						sTotalNet = 0 
						sTotalSTax = 0
						sTotalCTax = 0
						sTotalShipping = 0
						sHandling = 0 
						sGrandTotal = 0

						For i=0 to uBound(arrSales, 2)
							If rsPartners.Fields("orderTradingPartner") = arrSales(6, i) Then			
								If arrSales(0, i) <> "" Then sTotalNet = sTotalNet + cDbl(arrSales(0, i))
								If arrSales(1, i) <> "" Then sTotalSTax = sTotalSTax + cDbl(arrSales(1, i))
								If arrSales(2, i) <> "" Then sTotalCTax = sTotalCTax + cDbl(arrSales(2, i))
								If arrSales(3, i) <> "" Then sTotalShipping = sTotalShipping + cDbl(arrSales(3, i))
								If arrSales(4, i) <> "" Then sHandling = sHandling + cDbl(arrSales(4, i))
								If arrSales(5, i) <> "" Then sGrandTotal = sGrandTotal + cDbl(arrSales(5, i))
							End If
						Next
	        %>
	        <tr>
	        <td>
	        <table border="0" width="100%" cellpadding="4" cellspacing="0">
	        <tr>
	        <td width="100%" colspan=2 align="left" valign="top"><b>Affiliate Partner:&nbsp;<A href="sfReports4.asp?Affiliate=<%= rsPartners.Fields("orderTradingPartner") %>"><%= rsPartners.Fields("orderTradingPartner") %></a></b></td>
	        </tr>  
	        <tr>
	        <td width="75%" align="right" valign="top">Net Sales:</td>
	        <td width="25%" align="left" valign="top"><%= FormatCurrency(sTotalNet) %></td>
	        </tr>  
	        <tr>
	        <td width="75%" align="right" valign="top">Total State/Providence Tax:</td>
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
	        </table>
	        </td>
	        </tr>
	        <%
					rsPartners.MoveNext 
				Loop 
				closeObj(rsPartners)
	        Else 
	        %>
							<tr>
				<td colspan=4 align="center" class="tdContent"><font class="Error">There Were No Affiliate Sales Between <%= sStartDate %> And <%= sEndDate %></font></td>
				</tr>

			<% End If %>
			<tr>
			<td width="100%" align="center" valign="top" colspan="4"></td>
			</tr>
			</table>
			
    </td>
    </tr>
<%
Else
	Set rsAff = Server.CreateObject("ADODB.Recordset")
	rsAff.Open "sfAffiliates", cnn, adOpenStatic, adLockReadOnly, adCmdTable
	sFilter = "affName = '" & sAffiliate & "'"
	rsAff.Filter = sFilter	
%>

	    <tr>
		<td align="middle" class="tdMiddleTopBanner">Affiliate Information</td>        
	    </tr>
	    <tr>
		<td class="tdBottomTopBanner"><b>Instructions: </b>The information for the affiliate partner you selected is listed below.  To modify this information, return to the main menu and click on <B>Affiliate Partner Administration</b>.</td>    
	    </tr>
	    <tr>
	    <td class="tdContent2" width="100%" nowrap>
	    <table width="100%">
	    <%If Not (rsAff.EOF And rsAff.BOF) Then%>
	    <tr>
	    <td align="left" width="20%" nowrap>Affiliate ID:</td>
	    <td align="left" width="80%" nowrap><b><%= rsAff.Fields("affName") %></b></td>
	    </tr>
	    <tr>
	    <td align="left" width="20%" nowrap>Company:</td>
	    <td align="left" width="80%" nowrap><b><%= rsAff.Fields("affCompany") %></b></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap>Address Line 1:</td>
	    <td align="left" width="80%" nowrap><b><%= rsAff.Fields("affAddress1") %></b></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap>Address Line 2:</td>
	    <td align="left" width="80%" nowrap><b><%= rsAff.Fields("affAddress2") %></b></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap>City:</td>
	    <td align="left" width="80%" nowrap><b><%= rsAff.Fields("affCity") %></b></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap>State:</td>
	    <td align="left" width="80%" nowrap><b><%= rsAff.Fields("affState") %></b></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap>Zip/Postal Code:</td>
	    <td align="left" width="80%" nowrap><b><%= rsAff.Fields("affZip") %></b></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap>Country:</td>
	    <td align="left" width="80%" nowrap><b><%= rsAff.Fields("affCountry") %></b></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap>Phone Number:</td>
	    <td align="left" width="80%" nowrap><b><%= rsAff.Fields("affPhone") %></b></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap>Fax Number:</td>
	    <td align="left" width="80%" nowrap><b><%= rsAff.Fields("affFAX") %></b></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap>Email:</td>
	    <td align="left" width="80%" nowrap><a href="mailto:<%= rsAff.Fields("affEmail") %>"><%= rsAff.Fields("affEmail") %></a></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap>Notes:</td>
	    <td align="left" width="80%" nowrap><b><%= rsAff.Fields("affNotes") %></b></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap>Web Site:</td>
	    <td align="left" width="80%" nowrap><a href="<%= rsAff.Fields("affHttpAddr") %>"><%= rsAff.Fields("affHttpAddr") %></a></td>
	    </tr>
	    <% Else %>
	    <tr>
	    <td align="center" width=100%"><font class="Content_Small">There is no information available for <%= sAffiliate %></font></td>
	    </tr>	    
	    <% End If %>
	    </table>
		
    </td>
    </tr>
<%
	closeObj(rsAff)
End If
closeObj(cnn)
%>


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









