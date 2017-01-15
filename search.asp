<%@ Language=VBScript %>
<%	option explicit 
	Response.Buffer = True
%>
<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="SFLib/incSearchResult.asp"-->
<!--#include file="sfLib/incDesign.asp"-->
<!--#include file="sfLib/incText.asp"-->
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file="SFLib/incGeneral.asp"-->
<%		
	'@BEGINVERSIONINFO

	'@APPVERSION: 50.4011.0.2

	'@FILENAME: search.asp
 

	

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

	If Trim(Request.QueryString("referer")) <> "" Then
		Session("TradingPartnerID") = Request.QueryString("referer")
	End If	

	Session("HttpReferer") = Request.ServerVariables("HTTP_REFERER") 
%>
<html>

<head>

<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%= C_STORENAME %>-SF Search Engine Page</title>


<!--Header Begin -->
<link rel="stylesheet" href="sfCSS.css" type="text/css">
</head>

<body bgproperties="fixed"  link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="<%= C_WIDTH %>" align="center">
  <tr> 
    <td width="1130"> 
      <table width="100%" border="0" cellspacing="1" cellpadding="3">
        <tr> 
          <td align="middle"  class="tdTopBanner"> 
            <%If C_BNRBKGRND = "" Then%>
            <img src="buttons/tt_blue.gif" border="0" width="275" height="36"> 
            <%Else%>
            <%End If%>
          </td>
        </tr>
        <!--Header End -->
        <tr> 
          <td align="center" class="tdMiddleTopBanner">Search Store</td>
        </tr>
        <tr> 
          <td class="tdBottomTopBanner2"> Please input the word(s) that you would 
            like to search for in our product database. For additional control 
            you may choose to search on &quot;All Words&quot; or &quot;Any Words&quot; 
            or for the &quot;Exact Phrase.&quot;&nbsp; For additional search options 
            use <i>Advanced Search</i>. </td>
        </tr>
        <tr> 
          <td class="tdContent2"> 
            <form method="get" name="searchForm" action="search_results.asp">
              <table border="0" cellpadding="0" cellspacing="5" width="100%">
                <tr> 
                  <td colspan="2" width="75%" align="right"><b>Search</b>&nbsp;&nbsp; 
                    <input type="text" style="<%= C_FORMDESIGN %>" name="txtsearchParamTxt" size="20">
                    &nbsp;&nbsp;<b>In</b>&nbsp;&nbsp; 
                    <select size="1" name="txtsearchParamCat" style="<%= C_FORMDESIGN %>">
                      <option value="ALL">All <%= C_CategoryNameP %></option>
                      <%= getCategoryList(0) %> 
                    </select>
                  </td>
                  <td width="25%" align="left"> 
                    <input type="image" name="btnSearch" src="<%= C_BTN01 %>" alt="Search" border="0">
                  </td>
                </tr>
                <tr> 
                  <td width="100%" colspan="3" align="center"><font class="Content_Small"> 
                    <p align="center"> 
                      <input type="radio" value="ALL" checked name="txtsearchParamType">
                      <b>ALL</b> Words 
                      <input type="radio" name="txtsearchParamType" value="ANY">
                      <b>ANY</b> Words 
                      <input type="radio" name="txtsearchParamType" value="Exact">
                      Exact Phrase | <a href="advancedsearch.asp"> Advanced Search</a>
                    </font> </td>
                </tr>
              </table>
              <input type="hidden" name="iLevel" value="1">
              <input type="hidden" name="txtsearchParamMan" value="ALL">
              <input type="hidden" name="txtsearchParamVen" value="ALL">
              <input type="hidden" name="txtFromSearch" value="fromSearch">
            </form>
          </td>
        </tr>
        <!--Footer begin-->
        <!--#include file="foot.txt"-->
      </table>
    </td>
  </tr>
  <tr> 
    <td height="25" valign="top" bgcolor="#FFFFFF">
      <div align="center"> <br>
        If you have been directed to this page by mistake or during checkout it 
        may be because the security settings on your computer are too high or 
        you are running privacy software that is blocking our shopping cart. To 
        prevent this from happening please set your computrer's security and privacy 
        settings to &quot;medium&quot; and temporarily disable any privacy software. 
        You will also need to have your &quot;cookies&quot; turned on. This allows 
        the shopping cart to recognize who you are and process your order. (For 
        more information on cookies please refer to <a href="http://support.microsoft.com/default.aspx?kbid=196955"><font color="#0000FF">Microsoft</font></a>.)<br>
        <br>
        If you are still having trouble please contact us by e-mail (<a href="mailto:contact@tributetea.com"><font color="#0000FF">contact@tributetea.com</font></a>).<br>
        If you send us a message letting us know when and where to call we'll 
        happy to call you back to take a phone order.<br>
      </div>
    </td>
  </tr>
  <tr>
    <td height="50"></td>
  </tr>
</table>
  
  
  </body>
</html>
<!--Footer End-->


<%
closeObj(cnn)
%>





