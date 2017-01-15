
<%@ Language=VBScript %>
<%	option explicit 
%>
<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file="SFLib/incGeneral.asp"-->
<!--#include file="sfLib/incDesign.asp"-->
<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: 

'@FILEVERSION: 



'@DESCRIPTION: 

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO


Dim sProdId

sProdId = Request.QueryString("sProdId")


%>

<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<title><%= C_STORENAME %>-SF StockInfo</title>
<!--Header Begin -->
<link rel="stylesheet" href="sfCSS.css" type="text/css">
</head>
<body >
<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="100%" align="center">
  <tr>
  <td>
		<table width="100%" border="0" cellspacing="1" cellpadding="3">
			<tr><td align="middle" class="tdTopBanner"><%= C_STORENAME %></td></tr>
<!--Header End -->
			<tr><td align="center" class="tdMiddleTopBanner"><center>Stock Information: <%=GetProductName(sProdid)%></center></td><tr>
			
		</table>
     
		<table border="1" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="100%" align="center">
		
		<tr>
		<td align = "middle" class="tdContent2"><b>Product</B></td>
		<td align = "middle" class="tdContent2"><b>In Stock</b></td>
		
		<%ShowInventory%>
<tr><td colspan=2 class="tdFooter" align=center><B><a href="javascript:window.close()">Close</a></B></td></tr>
<!--Footer Begin -->
            </table>
            
            </body>
         </html>
<!--Footer End -->

<%
'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------

Sub ShowInventory
Dim sql
Dim rst
Dim i
dim sAttName
	sql = "Select * FROM sfInventory WHERE invenProdid= '" & sProdID & "'"
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
	If rst.recordcount <= 0 then
		Response.Write "Sorry, no stock information is available on this product."
	Else
	        
		rst.MoveFirst 
		For I = 1  to rst.RecordCount
		
		If trim(rst("invenAttName")) = "" then  
			sAttName = GetProductName(sProdId)
		else
			sAttName = trim(rst("invenAttName"))
		End If

			response.write "<tr>"
			response.write "<td align = ""middle"" class=""tdContent"">"
			response.write sAttName & "</td>"
		
			response.write "<td align = ""middle"" class=""tdContent"">"
			response.write rst("invenInStock")			
			response.write "</td> "
         
			response.write "</tr>"

		If not rst.EOF then rst.MoveNext 
		Next 
			
	end if

	CloseObj (rst)

End Sub


'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------

Function GetProductName (strProductID)
dim sql
dim rst

	
	
	Set rst = Server.CreateObject("ADODB.RecordSet")		

	sql = "Select * FROM sfProducts WHERE ProdID= '" & strProductID & "'" 

	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	

	If rst.recordcount >0 then
		GetProductName = rst("ProdName")
	else
		GetProductName= ""
	end If

	CloseObj (rst)

End function





cnn.Close
Set cnn = Nothing




%>



