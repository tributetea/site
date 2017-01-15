<%@ Language=VBScript %>
<%	option explicit 
	Response.Buffer = True
	%>
<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file="SFLib/incGeneral.asp"-->
<!--#include file="sfLib/incDesign.asp"-->
<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME:  



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
<!--Header Begin -->
<link rel="stylesheet" href="sfCSS.css" type="text/css">
</head>
<body link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">
<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="100%" align="center">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="3">
        <tr><td align="middle" class="tdTopBanner"><%= C_STORENAME %></td> </tr>
<!--Header End -->
    	<tr><td align="center" class="tdMiddleTopBanner"><center>Volume Pricing Information: <%=GetProductName(sProdid)%></center></td><tr>
		</table>
        </tr>
        <tr> <td class="tdContent2">
            <hr noshade color="#000000" size="1" width="90%">     <p align="center">
            <b>
            <%ShowMTPrices%>
            </b>
            <BR>
            <BR>
            <b>
            
            <hr noshade color="#000000" size="1" width="90%">
      </b></td>        
            </tr>
<!--Footer Begin Modified (only one closing table tag)-->
            <tr>
              <td class="tdFooter"></td>
                </tr>
                </table>
            <CEnter> <a href="javascript:window.close()">Close</a> </center>
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

Sub ShowMTPrices
Dim sql
Dim rst
Dim i

	sql = "Select * FROM sfMTPrices WHERE mtprodid= '" & sProdID & "' ORDER By mtIndex ASC"
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
	If rst.recordcount <= 0 then
		Response.Write "Sorry, no volume discount is available on this product."
	Else
		rst.MoveFirst 
		For I = 1  to rst.RecordCount
		    If rst("mtType") = "Amount" Then	
				Response.Write "<BR> Buy " & rst("mtQUANTITY") & " or more and save " & FormatCurrency(rst("mtvalue")) & " on each item. " 
			else
				Response.Write "<BR> Buy " & rst("mtQUANTITY") & " or more and save " & rst("mtvalue") & "% on each item. " 
			End If
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
