<%@ Language=VBScript %>
<%	option explicit 
	Response.Buffer = True
%>
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="SFLib/incGeneral.asp"-->
<!--#include file="SFLib/incDesign.asp"-->
<%		
	'@BEGINVERSIONINFO

	'@APPVERSION: 50.4011.0.2

	'@FILENAME: unsubscribe.asp 
 
	

	'@DESCRIPTION: Unsubscribes Customers

	'@STARTCOPYRIGHT
	'The contents of this file is protected under the United States
	'copyright laws and is confidential and proprietary to
	'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
	'expressed written permission of LaGarde, Incorporated is expressly prohibited.
	'
	'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
	'@ENDCOPYRIGHT

	'@ENDVERSIONINFO

%>
<html>


<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%= C_STORENAME %>-SF Unsubscribe to Mailing List Page</title>


<%
Dim sEmail, rs, sSQL

sEmail = trim(Request.QueryString("email"))

sSQL = "SELECT custID, custEmail, custIsSubscribed FROM sfCustomers WHERE custEmail = '" & sEmail & "'"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText
If Not (rs.EOF And rs.BOF) Then
	rs.Fields("custIsSubscribed") = 0
	rs.Update 
End If
closeObj(rs)
closeObj(cnn)

If sEmail = "" Then sEmail = "Nothing"

%>
<!--Header Begin -->
<link rel="stylesheet" href="sfCSS.css" type="text/css">
</head>
<body bgproperties="fixed"  link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">

<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="<%= C_WIDTH %>" align="center">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="3">
        <tr>
          <td align="middle"  class="tdTopBanner"><%If C_BNRBKGRND = "" Then%><%= C_STORENAME %><%Else%><img src="<%= C_BNRBKGRND %>" border="0"><%End If%></td>
	    </tr>	
<!--Header End -->         
        <tr>
          <td align="center" class="tdMiddleTopBanner">Unsubscribe</td>
        </tr>
 	    <tr>
          <td class="tdContent">        
		    <table border="0" cellpadding="0" cellspacing="5" width="100%">
              <tr>
			    <td width="75%" align="center"><b><%= sEmail %> has been removed from the mailing list of <%= C_STORENAME %></b></td>
            
              </tr>
            </table>
	      </td>
        </tr>
<!--Footer begin-->
                <!--#include file="footer.txt"-->
              </table>
            </td>
          </tr>
        </table>
       </body>

    </html>
<!--Footer End-->

