<%	option explicit 
	Response.Buffer = True
%>
<!--#include file="SFLib/db.conn.open.asp"-->

<!--#include file="sfLib/incDesign.asp"-->
<!--#include file="sfLib/incText.asp"-->
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file="SFLib/incGeneral.asp"-->
<%		
	'@BEGINVERSIONINFO

	'@APPVERSION: 50.4011.0.2

	'@FILENAME: default.asp
 

	'

	'@DESCRIPTION: Search Page
	
	'@STARTCOPYRIGHT
	'The contents of this file is protected under the United States
	'copyright laws and is confidential and proprietary to
	'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
	'expressed written permission of LaGarde, Incorporated is expressly prohibited.
	'
	'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
	'@ENDCOPYRIGHT

	'@ENDVERSIONINFO
	
	On Error Resume Next
	dim rsAdmin, sStoreName,sSSLpath,sDomain,sSql
	'Response.Write Session("CurrencyISO")
%>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">

<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%= C_STORENAME %>-SF Search Engine Page</title>



<%
set rsAdmin = server.CreateObject("adodb.recordset")	
sSQL = "Select adminSSLPath, adminDomainName,adminStoreName from sfAdmin"
			rsAdmin.Open sSQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
			sSSLpath = trim(rsadmin.Fields("adminSSLPath"))
			sDomain =trim(rsadmin.Fields("adminDomainName"))
			sStoreName = rsadmin.Fields("adminStoreName")
			closeObj(rsAdmin)

%>

<!--Header Begin -->
<link rel="stylesheet" href="sfcss.css" type="text/css">
</head>

<body  bgproperties="fixed" link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">

<table border="0" cellpadding="1" cellspacing="0"  class="tdbackgrnd" width="<%= C_WIDTH %>" align="center">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="5">
        <tr>
          <td align="middle"  class="tdTopBanner"><%If C_BNRBKGRND = "" Then%><%= C_STORENAME %><%Else%><img src="<%= C_BNRBKGRND %>" border="0"><%End If%></td>
	    </tr>	
        <tr>
          <td align="center"  class="tdMiddleTopBanner">Welcome</td>
        </tr>
    
        <%
    if sSSLpath = "http://www.yourdomain.com/ssl/process_order.asp" or sDomain ="http://www.yourdomain.com(do not specify a page)" then
    %>
        <tr>
          <td bgcolor="<%= C_BGCOLOR3 %>" class="Error">Error!  It appears you need to set the domain, or the SSL directory for your store!
          </td>
	    </tr>
	    <%
	else
	%>
	    <tr>
          <td class="tdContent2"><h1>Welcome to your StoreFront Web Store!</h1>
      <p>The StoreFront Web Creation Wizard has successfully created a new
      StoreFront web on your local machine. Following are links to important
      files in your StoreFront web:</p>
           <blockquote>
              <blockquote>
              <p><a href="search.asp">Search</a> To Search your Store for Products.</p>
              <p><a href="advancedsearch.asp">Advanced Search</a> For Advanced Product Search Features.</p>
              <p><a href="newproduct.asp">New Products</a> To View Newly Added Items in your Store.</p>
              <p><a href="salespage.asp">Sale Items</a> To View all Products on Sale.</p>
              <p><a href="affiliate.asp">Affiliates</a> For Affiliate Partners Sign-up.</p>
              </blockquote>
            </blockquote>
            <p>To begin building your web store, use the Store Builder tool to
      complete the following:</p>
            <blockquote>
              <blockquote>
                <p>Configure the StoreFront Web</p>
              <p>Build the Product Database</p>
              <p>Create Product Pages</p>
              </blockquote>
            </blockquote>
            <p>If you have any questions, see the StoreFront Help files that are
      located in the StoreFront toolbar, visit the online <a href="http://support.storefront.net/Resources/kbase/kbsearch.asp">
      StoreFront Knowledge Base </a> ,
      or <a href="mailto:support@storefront.net">e-mail StoreFront Support</a>.<br>
                </p>
            <p>&nbsp;</p>
            </td>
	        </tr>
           <!--#include file="footer.txt"-->            
       <%
	end if 
	%>	        
<!--Footer begin  Modified for this page (footer.txt is above)-->
	            
            </table>
          </td>
        </tr>
      </table>
    </body>
  </html>
<!--Footer End-->
<%
closeObj(cnn)
%>





