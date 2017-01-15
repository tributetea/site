<%@ Language=VBScript %>
<%	option explicit 
	Response.Buffer = True
	server.ScriptTimeout = 900
%>
<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file="SFLib/incSearchResult.asp"-->
<!--#include file="SFLib/incGeneral.asp"-->
<!--#include file="sfLib/incDesign.asp"-->
<!--#include file="sfLib/incText.asp"-->
<!--#include file="sflib/incAE.asp"--> 

<%	

	'@BEGINVERSIONINFO

	'@APPVERSION: 50.4011.0.3

	'@FILENAME: salespage.asp
 
	

	'@DESCRIPTION: Displayes all products on sale

	'@STARTCOPYRIGHT
	'The contents of this file is protected under the United States
	'copyright laws and is confidential and proprietary to
	'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
	'expressed written permission of LaGarde, Incorporated is expressly prohibited.
	'
	'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
	'@ENDCOPYRIGHT

	'@ENDVERSIONINFO
    'log ref #186 djp
    'modified 10-29-01 
	' Constant Declarations
	const varDebug  = 0		'DeBug Setting
	'const iDesign	= 3		'Layout Selection
	dim iDesign
	iDesign	= C_DesignType		'Layout Selection
	
	const iPageSize	= 10	'unchangable Page size 
	const iMaxRecords = 0   'Maximum amount of records returned, 0 is no maximum
	
	Dim txtsearchParamTxt, txtsearchParamType, txtsearchParamCat, txtFromSearch,  txtsearchParamMan
	Dim txtCatName, txtsearchParamVen, txtImagePath, txtOutput, txtDateAddedStart
	Dim txtDateAddedEnd, txtPriceStart, txtPriceEnd, txtSale, SQL, sAmount
	Dim iAttCounter, irsSearchAttRecordCount, iAttDetailCounter, irsSearchAttDetailRecordCount
	Dim iPage, iRec, iNumOfPages, iDesignCounter, iVarPageSize, iSearchRecordCount, icounter
	Dim rsCat, rsSearch, rsSearchAtt, rsSearchAttDetail, rsCatImage, arrAttDetail, arrProduct, arrAtt, CurrencyISO
	
	iDesignCounter = 2
	iVarPageSize = iPageSize ' Records Per Page
	txtsearchParamCat = "ALL"
	
%>

<%

	CurrencyISO = getCurrencyISO(Session.LCID)
	If iConverion = 1 Then Response.Write "<script language=""JavaScript"" src=""http://www.oanda.com/cgi-bin/fxcommerce/fxcommerce.js?user=" & sUserName & """></script>"

	Set rsSearch = Server.CreateObject("ADODB.RecordSet")
	' -------------------------------------------
	' RecordSet Paging Setup --------------------
	' -------------------------------------------
	
	If Application("AppName")="StoreFrontAE" then	
    ' SQL = "SELECT sfProducts.prodID, sfProducts.prodName, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodPrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice," _
    ' & " sfProducts.prodDescription, sfProducts.prodAttrNum, sfProducts.prodCategoryId, sfProducts.prodShortDescription, sfSub_Categories.CatHierarchy" _
    ' & " FROM (sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) INNER JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID" _
    ' & " WHERE prodSaleIsActive = 1 AND prodEnabledIsActive = 1"
      SQL = "SELECT ProdID, prodName, prodImageSmallPath, prodLink, prodPrice, prodSaleIsActive, prodSalePrice, ProdID, prodDescription, prodAttrNum, prodCategoryId, prodShortDescription " _
			& "FROM sfProducts  WHERE prodSaleIsActive = 1 AND prodEnabledIsActive = 1" 
   
   
   ' Response.Write "1"
   Else  
     SQL = "SELECT ProdID, prodName, prodImageSmallPath, prodLink, prodPrice, prodSaleIsActive, prodSalePrice, catName, prodDescription, prodAttrNum, prodCategoryId, prodShortDescription " _
			& "FROM sfProducts INNER JOIN sfCategories ON sfProducts.prodCategoryId = sfCategories.catID WHERE prodSaleIsActive = 1 AND prodEnabledIsActive = 1" 
  ' Response.Write "2"
   End if
	 
  ' Response.Write SQL

		 
	If varDebug = 1 Then Response.Write SQL & "<br><br>"
	With rsSearch
		.CursorLocation = adUseClient
		.CacheSize = iVarPageSize
		.MaxRecords = iMaxRecords
		.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText 
		.PageSize = iVarPageSize
	End With
	
	' Determine the page user is requesting
	If Request.QueryString("PAGE") = "" Then
		iPage = 1
	Else
		iPage = CInt(Request.QueryString("PAGE"))
		' Protect against out of range pages, in case
		' of a user specified page number
		If iPage < 1 Then
			iPage = 1
		Else
			If iPage > rsSearch.PageCount Then
				iPage = rsSearch.PageCount
			Else
				iPage = CInt(Request.QueryString("PAGE"))
			End If
		End If
	End If
	
	iSearchRecordCount = rsSearch.RecordCount
	iNumOfPages = Int(iSearchRecordCount / iPageSize)
	
	If CInt(iNumOfPages+1) = CInt(iPage) Then iVarPageSize = rsSearch.RecordCount - (iNumOfPages * iPageSize) 

	If NOT rsSearch.EOF Then rsSearch.AbsolutePage = CInt(iPage)
	
	' Create Attribute Record Sets for product on page
	SQL = getAttributeSQL(rsSearch, iVarPageSize, iPage)
	If varDebug = 1 Then Response.Write SQL & "<br><br>"	
	Set rsSearchAtt = Server.CreateObject("ADODB.RecordSet")
	If SQL <> "" Then 
		rsSearchAtt.Open SQL, cnn, adOpenKeyset, adLockReadOnly, adCmdText
		SQL = getAttributeDetailSQL(rsSearchAtt)
		If varDebug = 1 Then Response.Write SQL & "<br><br>"
		If SQL <> "" Then
			Set rsSearchAttDetail = Server.CreateObject("ADODB.RecordSet")
			rsSearchAttDetail.Open SQL, cnn, adOpenKeyset, adLockReadOnly, adCmdText
		End If 
	End If 
	
	If txtsearchParamCat = "ALL" Then
		txtCatName = "All " & C_CategoryNameP
	Else
		If Not rsSearch.EOF Then
		    if Application("AppName")="StoreFrontAE" then	
		    	txtCatName =GetFullPath(rsSearch.Fields("CatHierarchy"),1,"ALL")
		    
		    else
		      
		      txtCatName =rsSearch.Fields("CatNAME")
		   end if 
		Else
			Set rsCat = Server.CreateObject("ADODB.RecordSet")
			SQL = getCategorySQL(txtsearchParamCat)
			rsCat.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
			txtCatName = rsCat.Fields("catName")
			closeObj(rsCat)
		End If 
	End If
	If txtsearchParamTxt = "" Then txtsearchParamTxt = "*"
	'Corrects Number of Pages if there is overflow less then records per page 
	If iSearchRecordCount mod iPageSize <> 0 Then iNumOfPages = iNumOfPages + 1
                     	
%>
<html>

<head>
<SCRIPT language="javascript" src="SFLib/incae.js"></SCRIPT>
<SCRIPT language="javascript" src="SFLib/sfCheckErrors.js"></SCRIPT>
<SCRIPT language="javascript" src="SFLib/sfEmailFriend.js"></SCRIPT>

<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%= C_STORENAME %>-SF Product Sale Page</title>


<!--Header Begin -->
<link rel="stylesheet" href="sfCSS.css" type="text/css">
</head>

<body  bgproperties="fixed"  link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">

<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="<%= C_WIDTH %>" align="center">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="2">
        <tr>
          <td align="middle"  class="tdTopBanner"> 
            <%If C_BNRBKGRND = "" Then%>
            <%Else%>
            <img src="buttons/tt_blue.gif" border="0" width="275" height="36"> 
            <%End If%>
          </td>
        </tr>
<!--Header End -->
        <tr>
          <td>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
		      <tr>
		        <td align="left" class="tdMiddleTopBanner" width = "20%">&nbsp;</td>
		        <td align="center" class="tdMiddleTopBanner"><center>Sales Page</center></td>
		        <td align="right" class="tdMiddleTopBanner" width="20%"><a href="<%= C_HomePath %>order.asp"><img src="buttons/checkout.gif" border="0" align="right" valign="top" alt="Check Out" width="107" height="22"></a></td>
		      </tr>
	        </table>
	      </td>
        </tr>
        <tr>
    
          <!-- ### SEARCH RESULT QUERY OUTPUT ::: BEGIN ### -->    
          <td align="middle" class="tdBottomTopBanner">Search for: <b><%= txtsearchParamTxt %></b> in <%= C_CategoryNameS %>: <b><%= txtCatName %></b><%
    %>
            <br><b><%= iSearchRecordCount %></b> Records Returned <b><%= iVarPageSize %></b> Records Displayed.<br><a href="search.asp"><b>New Search</b></a>|<a href="advancedsearch.asp"><b>Advanced Search</b></a></td>
          <!-- ### SEARCH RESULT QUERY OUTPUT ::: END ### -->            
        </tr>
        <tr>
          <td class="tdContent2">
            <table border="0" width="100%">
              <tr>
		        <td width="100%"><p align="center"><font class="Content_Small">
                  <%
        If iNumOfPages <> 1 And iNumOfPages <> 0 Then Response.Write bottomPaging(iPage, iPageSize, iSearchRecordCount, iNumOfPages, "SalesPage")
		%>
                  </font>
                  <hr noshade color="#000000" size="1" width="90%">
                  </td>
                </tr>
		      </table>
              <%
	' -------------------------------------------
    ' Empty Search Results ----------------------
    ' -------------------------------------------
    If rsSearch.EOF Then%>
<table border=0 width=100&#37;>
<tr>
<td><center>There are currently no sale items available from inventory</center></td></tr>
<tr>
<td width="100%"; colspan="2"><hr noshade color="#000000" size="1" width="90%">
</td>
</tr>
</table>
<%
    Else
        ' -------------------------------------------
        ' SEARCH RESULT PRODUCT OUTPUT ::: BEGIN ----
        ' -------------------------------------------
        'create arrays for display
        arrProduct = rsSearch.GetRows(iVarPageSize)
        If Not rsSearchAtt.EOF Then 
			arrAtt = rsSearchAtt.GetRows
			irsSearchAttRecordCount = rsSearchAtt.RecordCount-1
			If Not rsSearchAttDetail.EOF Then 
				arrAttDetail = rsSearchAttDetail.GetRows
				irsSearchAttDetailRecordCount = rsSearchAttDetail.RecordCount-1 
			End If
		End If
				
		dim bDuplicate,iDupRec
		
        For iRec = 0 to iVarPageSize-1 
           SearchResults_CheckInventoryTracked 'SFAE  b2   
		    'sfAE
		     If iRec > (iVarPageSize-1) then EXIT FOR 'SFAE b2 
             
               
               
                ' Set Default Image if none specified for product
               ' If arrProduct(2, iRec) = "" Then
               '     Set rsCatImage = Server.CreateObject("ADODB.RecordSet")
               '     SQL = "SELECT catImage FROM sfCategories WHERE catID = " & arrProduct(9, iRec)
               '     rsCatImage.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
               '     txtImagePath = rsCatImage.Fields("catImage")
               '     closeObj(rsCatImage) 
               ' Else
                    txtImagePath = arrProduct(2, iRec)
               ' End If
      			icounter = 1 
				       

%>		
              <form method="post" name="<%= arrProduct(0, iRec)%>" action="<%= C_HomePath %>addproduct.asp" onSubmit="this.QUANTITY.quantityBox=true;return sfCheck(this);">            
                <input TYPE="hidden" NAME="PRODUCT_ID" VALUE="<%= arrProduct(0, iRec)%>">         
                <table border="0" width="100%" class="tdContent" >
                  <tr>
                    <% If iDesign = "1" Then %>            
                    <td width="30%" align="center"><a href="<%= replace(arrProduct(3, iRec)," ","+") %>"><%If txtImagePath <> "" Then%><img border="1" src="<%= txtImagePath %>"><%ElseIf Trim(arrProduct(3, iRec)) <> "" Then %>Link<%End If%></a><br>
                    </td>
                    <% ElseIf iDesign = "3" And (iDesignCounter / 2) = Int(iDesignCounter / 2) Then%>
		            <td width="30%" align="center"><a href="<%= replace(arrProduct(3, iRec)," ","+") %>"><%If txtImagePath <> "" Then%><img border="1" src="<%= txtImagePath %>"><%ElseIf Trim(arrProduct(3, iRec)) <> "" Then %>Link<%End If%></a><br>
                    </td>
                    <% End If %>
                    <td width="70%" valign="top">
                      <b><%= C_ProductID %>:</b>&nbsp;<%= arrProduct(0, iRec) %>&nbsp;&nbsp;&nbsp;
                      <b><%= C_CategoryNameS %>:</b>&nbsp;<%
                      if Application("AppName")="StoreFrontAE" then	
                        Response.Write  GetFullPath(arrProduct(7, iRec),0,"ALL") 
                      else  
                        Response.Write arrProduct(7, iRec)
                      end if  
                       %><br>
                      <b><font class='Content_Large'><%= arrProduct(1, iRec) %></font></b><br>
                      <b><%= C_Description %>:</b>&nbsp;<%
                      
                        If arrProduct(11, iRec) <> "" Then
                       
                         Response.Write  arrProduct(11, iRec) 
                        Else
                          Response.Write  arrProduct(8, iRec) 
                        End If
                      
                        
                       %><br>
                      <b><%= C_Price %>:</b>&nbsp;
			          <% If iConverion = 1 Then
					If arrProduct(5, iRec) = 1 Then 
							Response.Write "<i><strike><script>document.write(""" & FormatCurrency(arrProduct(4, iRec)) & " = ("" + OANDAconvert(" & trim(arrProduct(4, iRec)) & "," & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script></strike></i><br>"
							Response.Write "<font color=#FF0000><b>" & C_SPrice & ": <script>document.write(""" & FormatCurrency(arrProduct(6, iRec)) & " = ("" + OANDAconvert(" & trim(arrProduct(6, iRec)) & ", " & chr(34) & CurrencyISO & chr(34) & ")+ "")"");</script></b></font><br>"
							Response.Write "<font color=#FF0000><i>" & C_YSave & " <script>document.write(""" & FormatCurrency(CDbl(arrProduct(4, iRec))-CDbl(arrProduct(6, iRec))) & " = ("" + OANDAconvert(" & trim(CDbl(arrProduct(4, iRec))-CDbl(arrProduct(6, iRec))) & ", " & chr(34) & CurrencyISO & chr(34) & ")+ "")"");</script></i></font><br>"
					Else
							Response.Write "<script>document.write(""" & FormatCurrency(arrProduct(4, iRec)) & " = ("" + OANDAconvert(" & trim(arrProduct(4, iRec)) & ", " & chr(34) & CurrencyISO & chr(34) & ")+ "")"");</script>"
					End If 
			   Else
					If arrProduct(5, iRec) = 1 Then 
							Response.Write "<i><strike>" & FormatCurrency(arrProduct(4, iRec)) & "</strike></i><br>"
							Response.Write "<font color=#FF0000><b>" & C_SPrice & ": " & FormatCurrency(arrProduct(6, iRec)) & "</b></font><br>"
							Response.Write "<font color=#FF0000><i>" & C_YSave & " " & FormatCurrency(CDbl(arrProduct(4, iRec))-CDbl(arrProduct(6, iRec))) & "</i></font><br>"
					Else
							Response.Write FormatCurrency(arrProduct(4, iRec))
					End If 
			   End If
			%>
				  <% SearchResults_GetProductInventory arrProduct(0, iRec) 'SFAE %>								
				  <% SearchResults_ShowMTPricesLink arrProduct(0, iRec) 'SFAE%> 

			          <br>
                      <table border="0" align="center">
                        <%
                            ' -------------------------------------------
                            ' SEARCH RESULT ATTRIBUTE OUTPUT ::: BEGIN --
                            ' -------------------------------------------
                            If irsSearchAttRecordCount <> "" Then
								For iAttCounter = 0 to irsSearchAttRecordCount
									If arrProduct(0, iRec) = arrAtt(2, iAttCounter) Then
%>                    
                        <tr>                
                          <td align="right"><%= arrAtt(1, iAttCounter) %></td>
                          <td><select size="1" name="attr<%= icounter %>" style="<%= C_FORMDESIGN %>">
                              <%
										For iAttDetailCounter = 0 to irsSearchAttDetailRecordCount
											If arrAtt(0, iAttCounter) = arrAttDetail(1, iAttDetailCounter) Then
														sAmount = ""
												Select Case arrAttDetail(4, iAttDetailCounter)
													Case 1 
														sAmount = " (add " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
													Case 2 
														sAmount = " (subtract " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
												End Select
												Response.Write "<option value=" & arrAttDetail(0, iAttDetailCounter) & ">" & arrAttDetail(2, iAttDetailCounter) & sAmount & "</option>"
											End If
										Next
%>
                            </select></td>
                        </tr>
                        <%  
									icounter = icounter + 1
									End If 
								Next
							End If 
                            ' -------------------------------------------
                            ' SEARCH RESULT ATTRIBUTE OUTPUT ::: END
                            ' -------------------------------------------
                     
%>
                     </table>
                      <p align="center"><center><%= C_Quantity %>: <input style="<%= C_FORMDESIGN %>"  type="text" name="QUANTITY" title="Quantity" size="3"><br>
                      <%SearchResults_GetGiftWrap arrProduct(0, iRec) 'SFAE%>
                        <input type="image" name="AddProduct" border="0" src="buttons/addtocart3.gif" alt="Add To Cart" width="107" height="28">
                        <br>
                      <% If iSaveCartActive = 1 Then%>
                        <input type="image" name="SaveCart" border="0" src="buttons/savetocart3.gif" alt="Save To Cart" width="138" height="22">
                        <%
        End if
        If iEmailActive = 1 Then
        %>
                        <a href="javascript:emailFriend('<%= server.urlencode(arrProduct(0, iRec)) %>')"><img border="0" src="buttons/email.gif" alt="Email a Friend" width="151" height="22"></a> 
                        <% End If %>
                      </CENTER></P></td>
                    <%  If iDesign = "2" Then %>            
		            <td width="30%" align="center"><a href="<%= replace(arrProduct(3, iRec)," ","+") %>"><%If txtImagePath <> "" Then%><img border="1" src="<%= txtImagePath %>"><%ElseIf Trim(arrProduct(3, iRec)) <> "" Then %>Link<%End If%></a><br>
                    </td>
                    <%  ElseIf iDesign = "3" And (iDesignCounter / 2) <> Int(iDesignCounter / 2) Then %>
		            <td width="30%" align="center"><a href="<%= replace(arrProduct(3, iRec)," ","+") %>"><%If txtImagePath <> "" Then%><img border="1" src="<%= txtImagePath %>"><%ElseIf Trim(arrProduct(3, iRec)) <> "" Then %>Link<%End If%></a><br>
                    </td>
                    <%  End If %>
                  </tr>
                  <tr>
                    <td width="100%" colspan="2"><hr noshade color="#000000" size="1" width="90%">
                    </td>
                  </tr>    
                   
                 </table></form>
				 <%
				Response.Flush
            If iDesign = "3" Then iDesignCounter = iDesignCounter + 1
         
        Next
        ' -------------------------------------------
        ' SEARCH RESULT PRODUCT OUTPUT ::: END ------
        ' -------------------------------------------
    End If 
%>        
        
             <table border="0" width="100%">
                    <tr>
                      <td width="100%"><p align="center"><font class="Content_Small">
                        <%
                ' -------------------------------------------   
                ' SEARCH RESULT PAGING OUTPUT ::: BEGIN -----
                ' -------------------------------------------
If iNumOfPages <> 1 And iNumOfPages <> 0 Then Response.Write bottomPaging(iPage, iPageSize, iSearchRecordCount, iNumOfPages, "SalesPage")
                ' -------------------------------------------   
                ' SEARCH RESULT PAGING OUTPUT ::: END -------
                ' -------------------------------------------
%>
                        </font></p>
                      </td>
                    </tr>
                  </table>
    
<!--Footer begin-->
                  <!--#include file="footer.txt"-->
                </table>
              </td>
            </tr>
          </table>
        </body>
</html>
<!--Footer End-->
        <%
Function GetFullPath(Vdata,justMain,subCatID) 
Dim sSql ,X
Dim iCatId,sCriteria
Dim sFirst
Dim rst,rsCat,rsSubCat
Dim arrTemp ,bMain


If subCatID = "ALL" Then
		 sSql = "SELECT sfSubCatDetail.ProdID, sfSub_Categories.CatHierarchy" _
		 & " FROM sfSubCatDetail INNER JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID" _
		 & "  Where sfSubCatDetail.ProdID = '" & vData & "'"
		Set rst = Server.CreateObject("ADODB.RecordSet")
		rst.open sSql, cnn,adOpenStatic,adLockReadOnly ,1 
        ' Response.Write ssql

		If rst.eof = false then 
		 sCriteria = rst("CatHierarchy")
		else
		 GetFullPath = "No Category"
		 rst.close
		 set rst = nothing
		 exit Function
		End if
		rst.close
		set rst = nothing
Else
 sCriteria = vData
End if

bMain = false
 if left(sCriteria,4)= "none" then
  bMain = True
  arrTemp = split(sCriteria,"-")
  sCriteria = arrtemp(1)
 elseif sCriteria = "" then
   GetFullPath = "" 
  exit function
 elseif instr(sCriteria,"-") = 0  then
    sCriteria = sCriteria 
 end if 
  arrTemp = split(sCriteria,"-")
 Set rsCat = Server.CreateObject("ADODB.RecordSet")
 Set rsSubCat = Server.CreateObject("ADODB.RecordSet")
  rsSubCat.Open "sfSub_Categories",cnn,adOpenStatic,adLockReadOnly ,adcmdtable 
   For X = 0 To UBound(arrTemp)
     rsSubCat.Requery
     if arrTemp(X)<> "" then
      rsSubCat.Find "SubCatId = " & CInt(arrTemp(X))
      GetFullPath = GetFullPath & rsSubCat("SubCatName") & "-"
     end if
   Next
  sSql  = "Select catName From sfCategories Where catId =" & rsSubCat("subcatCategoryId")   
 rsCat.Open sSql,cnn,adOpenStatic,adLockReadOnly ,adcmdText
 if justmain = 1 then
    GetFullPath = rsCat("catName")
 else 
   On error Resume next
   if bMain = True Then
      GetFullPath = rsCat("catName")
   else
     GetFullPath = rsCat("catName") & "-" &  Left(GetFullPath, Len(GetFullPath) - 1)
   end if 
 end if
 Set rsCat = Nothing
 Set rsSubCat = Nothing
 Exit Function
End Function
        
	' Object Cleanup
	closeObj(rsSearch)
	closeObj(rsSearchAtt)
	closeObj(rsSearchAttDetail)
	closeObj(cnn)
%>



















