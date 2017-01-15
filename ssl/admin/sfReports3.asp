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

'@FILENAME: sfreports3.asp
	 
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
<title>SF Reports Page</title>
<HEAD>


<%
Dim sStartDate, sEndDate, rsTrans, sSQL, sOrderId

sStartDate = MakeUSDate(Request.QueryString("startDate"))
sEndDate = MakeUSDate(Request.QueryString("endDate"))
sOrderId = Request.QueryString("OrderId")

Set rsTrans = Server.CreateObject("ADODB.RecordSet")

If sOrderId = "" Then
	sSQL = "SELECT orderID, orderDate, trnsrspID, trnsrspOrderId, trnsrspCustTransNo, trnsrspMerchTransNo, trnsrspAVSCode, trnsrspAUXMsg, trnsrspActionCode, trnsrspRetrievalCode, trnsrspAuthNo, trnsrspErrorMsg, trnsrspErrorLocation, trnsrspSuccess " _ 
		& "FROM sfOrders INNER JOIN sfTransactionResponse ON sfOrders.orderID = sfTransactionResponse.trnsrspOrderId WHERE orderDate BETWEEN #" & sStartDate & "# AND #" & sEndDate & "#  and sfOrders.orderIsComplete = 1 ORDER BY orderID"
Else
	sSQL = "SELECT orderID, orderDate, trnsrspID, trnsrspOrderId, trnsrspCustTransNo, trnsrspMerchTransNo, trnsrspAVSCode, trnsrspAUXMsg, trnsrspActionCode, trnsrspRetrievalCode, trnsrspAuthNo, trnsrspErrorMsg, trnsrspErrorLocation, trnsrspSuccess " _ 
		& "FROM sfOrders INNER JOIN sfTransactionResponse ON sfOrders.orderID = sfTransactionResponse.trnsrspOrderId WHERE orderID = " & sOrderId & " and sfOrders.orderIsComplete = 1 ORDER BY orderID"
End If
rsTrans.Open sSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText
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
	<td align="middle" class="tdMiddleTopBanner">Transaction Services</td>        
    </tr>
    <tr>
	<td class="tdBottomTopBanner">Transaction information for each sale within the date range specified is listed below.  This report contains information returned by the payment processing service including error or authorization codes.</td>    
    </tr>
    <tr>
    <td class="tdContent2" width="100%" nowrap>
        <table border="0" width="100%" cellpadding="4" cellspacing="0">
<%if rsTrans.EOF and rsTrans.BOF then%>
   <tr>
				<td colspan=4 align="center" class="tdContent2"><font class="Error">There Were No Transactions Between <%= sStartDate %> And <%= sEndDate %></font></td>
				</tr>
<%else%>				

        <tr>
		<td width="100%" class="tdContentBar" colspan="4">Report for <%= sStartDate %> to <%= sEndDate %></td>        
        </tr>
<%
Do While Not rsTrans.EOF
%>
	<tr>
	<td colspan=4 width="90%"><hr width="100%"></td>
	</tr>
	<tr>
	<td colspan=4>
	<table border="1" width="90%" align="center">
	<tr>
	<td align="Left">Order ID:&nbsp;<%= rsTrans.Fields("orderID") %></td>
	<td align="Left">Order Date:&nbsp;<%= rsTrans.Fields("orderDate") %></td>
	</tr>
	<tr>
	<td align="Left">Authorization #:&nbsp;<%= rsTrans.Fields("trnsrspAuthNo") %></td>
	<td align="Left">Success:&nbsp;<%= rsTrans.Fields("trnsrspSuccess") %></td>
	</tr>
	<tr>
	<td align="Left">Customer Transaction #:&nbsp;<%= rsTrans.Fields("trnsrspCustTransNo") %></td>
	<td align="Left">Merchant Transaction #:&nbsp;<%= rsTrans.Fields("trnsrspMerchTransNo") %></td>
	</tr>
	<tr>
	<td align="Left">AVS Code:&nbsp;<%= rsTrans.Fields("trnsrspAVSCode") %></td>
	<td align="Left">AUX Message:&nbsp;<%= rsTrans.Fields("trnsrspAUXMsg") %></td>
	</tr>
	<tr>
	<td align="Left">Action Code:&nbsp;<%= rsTrans.Fields("trnsrspActionCode") %></td>
	<td align="Left">Retrieval Code:&nbsp;<%= rsTrans.Fields("trnsrspRetrievalCode") %></td>
	</tr>
	<tr>
	<td align="Left">Error Message:&nbsp;<%= rsTrans.Fields("trnsrspErrorMsg") %></td>
	<td align="Left">Error Location:&nbsp;<%= rsTrans.Fields("trnsrspErrorLocation") %></td>
	</tr>
	
	</table>
</td>  </tr>
<%
	rsTrans.MoveNext  
Loop 
closeObj(rsTrans)
closeObj(cnn)
%>
        <tr>
        <td width="100%" align="center" valign="top" colspan="4"></td>
        </tr>
<%end if%>				        
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









