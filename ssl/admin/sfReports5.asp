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

<!--#include file="incAdmin.asp"-->
<%
'--------------------------------------------------------------------
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.3

'@FILENAME: sfreports5.asp
	 
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
'Modified 12/6/01 
'Storefront Ref#'s: 251 djp
%>
<html>
<head>

<SCRIPT LANGUAGE=javascript>
<!--
function check_Input()
{
 var sID = frmProductSummary.txtProdID.value
 var cboID = frmProductSummary.sltProdID.value

  if(sID == "" && cboID == "")
  {
   alert("Please enter a product id, or select a product from the dropdown");
   frmProductSummary.txtProdID.focus();
   return false;
  } 
  else
  {
   return true;
  }
}

//-->
</SCRIPT>

<title>SF Reports Page</title>
<!--Header Begin -->
<link rel="stylesheet" href="../../sfCSS.css" type="text/css">
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
	Dim sStartDate, sEndDate, rsSales, sSQL, sTotalNet, sTotalSTax, sTotalCTax, sTotalShipping, sGrandTotal, arrSales, sHandling, i, sProdId
    Dim Itot 
    Dim iTotalSold 'As Integer
	sStartDate = MakeUSDate(Request.Form("startDate"))
	sEndDate = MakeUSDate(Request.Form("endDate"))
	sProdId = Request.Form("txtProdID")
	If sProdId = "" Then sProdId = Request.Form("sltProdId")
	
	Set rsSales = Server.CreateObject("ADODB.RecordSet")
	 sSql = "SELECT sfProducts.prodName, sfProducts.prodID, sfProducts.prodPrice,sfOrderDetails.odrdtQuantity FROM" _
   & " (sfOrders INNER JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) INNER JOIN sfProducts ON sfOrderDetails.odrdtProductID = sfProducts.prodID" _
   & " WHERE ((orderDate BETWEEN #" & sStartDate & "# AND #" & sEndDate & "#) AND odrdtProductID = '" & sProdId & "') and sfOrders.orderIsComplete = 1"

	
	
	
'	sSQL = "SELECT orderAmount, orderSTax, orderCTax, orderShippingAmount, orderHandling, orderGrandTotal FROM sfOrders INNER JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId " _
'	 & "WHERE ((orderDate BETWEEN #" & sStartDate & "# AND #" & sEndDate & "#) AND odrdtProductID = '" & sProdId & "') and sfOrders.orderIsComplete = 1"
	rsSales.Open sSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText

	If Not (rsSales.EOF and rsSales.BOF) Then arrSales = rsSales.GetRows 

	sTotalNet = 0 
	iTotalSold = 0
	'sTotalSTax = 0
	'sTotalCTax = 0
	'sTotalShipping = 0
	'sHandling = 0 
	'sGrandTotal = 0

	rsSales.Close 
	Set rsSales= Nothing
  ITot = 0
	If isArray(arrSales) Then
		For i=0 to uBound(arrSales, 2)
			If arrSales(3, i) <> "" Then 
			 'sTotalNet = sTotalNet + cDbl(arrSales(2, i))
			 iTotalSold = iTotalSold + CInt(arrSales(3, I))
			end if
			
			'If arrSales(1, i) <> "" Then sTotalSTax = sTotalSTax + cDbl(arrSales(1, i))
			'If arrSales(2, i) <> "" Then sTotalCTax = sTotalCTax + cDbl(arrSales(2, i))
			'If arrSales(3, i) <> "" Then sTotalShipping = sTotalShipping + cDbl(arrSales(3, i))
			'If arrSales(4, i) <> "" Then sHandling = sHandling + cDbl(arrSales(4, i))
			'If arrSales(5, i) <> "" Then sGrandTotal = sGrandTotal + cDbl(arrSales(5, i))
		Next
		sTotalNet = CDbl(arrSales(2, 0)) * iTotalSold
		ITot = iTotalSold
	End If 

	%>
	
	    <tr>
		<td align="middle" class="tdMiddleTopBanner">Sales Summary</td>        
	    </tr>

	    <tr>
		<td class="tdBottomTopBanner">Total sales for the product item selected are shown below.</td>    
	    </tr>
	    <tr>
	    <td class="tdContent2" width="100%" nowrap>
	        <table border="0" width="100%" cellpadding="4" cellspacing="0">
	        <tr>
			<td width="100%" class="tdContentBar" colspan="4">Report for Product <%= sProdId %> from <%= sStartDate %> to <%= sEndDate %></td>        
	        </tr>
	        <% If isArray(arrSales) Then %>
	        <tr>
	        <td width="75%" align="right" valign="top">Total Sold:</td>
	        <td width="25%" align="left" valign="top"><%= iTot %></td>
	        </tr>  
	        <tr>
	        <td width="75%" align="right" valign="top">Net Sales:</td>
	        <td width="25%" align="left" valign="top"><%= FormatCurrency(sTotalNet) %></td>
	        </tr>
	        <% Else %>
	        <tr>
	        <td colspan="2" align="center"><font class="Content_Small">No Sales Reported for Product <%= sProdId%> from <%= sStartDate %> to <%= sEndDate %></font></td>
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
	Dim objRS, sProdList
	Set objRS = getProductList()
	Do While Not objRS.EOF 		
		sProdList = sProdList & "<option value=""" & objRS("prodID") & """>" &  objRs("prodName") & "</option>"
		objRS.MoveNext
	Loop
	closeobj(objRS)
%>
	    <form method="post" name="frmProductSummary" onSubmit="javasrcipt: return check_Input();">
	    <tr>
		<td align="middle" class="tdMiddleTopBanner">StoreFront Reports</td>        
	    </tr>
	    <tr>
		<td class="tdBottomTopBanner">Enter the Product ID of the item in the <B>Product ID</B> field or select a product item from the <B>Product</B> drop-down box to view the total sales for the selected item within the date range specified.</td>    
	    </tr>
	    <tr>
	    <td class="Content2" width="100%" nowrap>
	        <table border="0" width="100%" cellpadding="4" cellspacing="0">
	        <tr>
			<td colspan="2" width="100%" class="tdContentBar">Create Product Report</td>        
	        </tr>
            <tr>
            <td width="50%" align="right">Product ID:</td>
            <td width="50%"><input name="txtProdID" style="<%= C_FORMDESIGN %>" size="25"></td>
            </tr>
            <tr>
            <td width="50%" align="right">Product :</td>
            <td width="50%"><select name="sltProdID" style="<%= C_FORMDESIGN %>" size="1"><option></option><%= sProdList %></select></td>
            </tr>
	        <tr>
	        <td colspan="2" width="100%" align="center" valign="top"><input type="image" name="btnSubmit" border="0" src="../<%= C_BTN18 %>" alt="Submit" WIDTH="108" HEIGHT="21"></td>
	        </tr>
	        </table><input Type="hidden" name="startDate" value="<%= Request.QueryString("startDate") %>">
	<input Type="hidden" name="endDate" value="<%= Request.QueryString("endDate") %>">
	    </td>
	    </tr>
	</form>
		
	
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







