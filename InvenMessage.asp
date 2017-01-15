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
Dim sMessage

'sMessage = "One or more products inventory has been depleted and your cart has been modified by the system. Please review your cart to verify the changes."
sMessage = "We're sorry but one or more items in your cart are no longer in " & _ 
			"stock. We have automatically updated your cart by removing the unavailable items. " & _
			"Please review your cart to verify changes."
%>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<!--Header Begin -->
<link rel="stylesheet" href="sfCSS.css" type="text/css">
</head>
<body>
<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="100%" align="center">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="3">
        <tr>
          <td align="middle"  class="tdTopBanner"><%= C_STORENAME %></td>
    
        </tr>
<!--Header End -->
        <tr>
          <td class="tdContent2">
            <hr noshade color="#000000" size="1" width="90%">
            <p align="center"><font class="Content_Large"><b>Stock Depleted!</b></font><br>
               <%=sMessage%>  <BR>
            <b><a href="javascript:window.close()">Close</a>
            <hr noshade color="#000000" size="1" width="90%">
            </b>
	        </td>        
            </tr>
<!--Footer Begin -->
            <tr>
              <td class="tdFooter"></td>
                </tr>
                </table>
              </td>
            </tr>
            </table>
            </body>
         </html>
<!--Footer End -->
<%cnn.Close
Set cnn = Nothing
%>



