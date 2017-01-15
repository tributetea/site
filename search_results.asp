<%@ Language=VBScript %>

<%
	option explicit
	Response.Buffer = True
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

	'@FILENAME: search_results.asp
 

	

	'@DESCRIPTION: Displays search results

	'@STARTCOPYRIGHT
	'The contents of this file is protected under the United States
	'copyright laws and is confidential and proprietary to
	'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
	'expressed written permission of LaGarde, Incorporated is expressly prohibited.
	'
	'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
	'@ENDCOPYRIGHT

	'@ENDVERSIONINFO

'Modified 11/20/01 
'Storefront Ref#'s: 128 'JF
'Storefront Ref#'s: 219 'DP
	' Constant Declarations
	const varDebug  = 0		'DeBug Setting
	const iPageSize	= 10 'Records Per Page 
	const iMaxRecords = 0   'Maximum amount of records returned, 0 is no maximum
	
	Dim txtsearchParamTxt, txtsearchParamType, txtsearchParamCat, txtFromSearch,  txtsearchParamMan
	Dim txtCatName, txtsearchParamVen, txtImagePath, txtOutput, txtDateAddedStart
	Dim txtDateAddedEnd, txtPriceStart, txtPriceEnd, txtSale, SQL, sAmount, rsCatImage
	Dim iAttCounter, irsSearchAttRecordCount, iAttDetailCounter, irsSearchAttDetailRecordCount
	Dim iPage, iRec, iNumOfPages, iDesignCounter, iVarPageSize, iSearchRecordCount, icounter, iDesign
	Dim rsCat, rsSearch, rsSearchAtt, rsSearchAttDetail, arrAttDetail, arrProduct, arrAtt, rsManufacturer, rsVendor
	dim CurrencyISO,sSubCat,sALLSUB,X,sMainCat ,iLevel,sNextLevel
		'sfAE
		dim bDuplicate,iDupRec
    
	 
	iDesign	= C_DesignType		'Layout Selection
	iDesignCounter = 2
	iVarPageSize = iPageSize ' Records Per Page
	
	txtFromSearch = Trim(Request.Form("txtFromSearch"))
    sSubCat = Request.item("subcat")
    if  sSubCat = "" then
     
         sSubCat = Request.item("txtsearchParamCat")
        
    end if

    sALLSUB = Request.item("txtsearchParamCat")
    iLevel  = Request.item("iLevel")
    if ilevel = 2 and  sALLSUB = "ALL" then
     sSubCat = Request.item("subcat")
    end if 
   ' Requests the variables depending on how the page is entered
	If txtFromSearch = "fromSearch" Then
	  txtsearchParamTxt	= trim(Replace(Replace(Request.Form("txtsearchParamTxt"), "'", "''"), "*", ""))
	  txtsearchParamType	= trim(Request.Form("txtsearchParamType"))
	   if Ilevel = 2 and sALLSUB = "ALL" then
	     txtsearchParamCat	= sSubCat
	     Ilevel = 1 
	   else 
    	 txtsearchParamCat	= trim(Request.QueryString("txtsearchParamCat"))
       end if
	  txtsearchParamMan	= trim(Request.Form("txtsearchParamMan"))
	  txtsearchParamVen	= trim(Request.Form("txtsearchParamVen"))
	  txtDateAddedStart	= MakeUSDate(trim(Request.Form("txtDateAddedStart")))
	  txtDateAddedEnd 	= MakeUSDate(trim(Request.Form("txtDateAddedEnd")))
	  txtPriceStart		= trim(Request.Form("txtPriceStart"))
	  txtPriceEnd 		= trim(Request.Form("txtPriceEnd"))
	  txtSale			= trim(Request.Form("txtSale"))
	Else
	  txtsearchParamTxt	= trim(Replace(Replace(Request.QueryString("txtsearchParamTxt"), "'", "''"), "*", ""))
	  txtsearchParamType	= trim(Request.QueryString("txtsearchParamType"))
	   if Ilevel = 2 and sALLSUB = "ALL" then
	     txtsearchParamCat	= sSubCat
	     Ilevel = 1 
	   else 
    	 txtsearchParamCat	= trim(Request.QueryString("txtsearchParamCat"))
	   end if
	  txtsearchParamMan	= trim(Request.QueryString("txtsearchParamMan"))
	  txtsearchParamVen	= trim(Request.QueryString("txtsearchParamVen"))
	  txtDateAddedStart	= MakeUSDate(trim(Request.QueryString("txtDateAddedStart")))
	  txtDateAddedEnd 	= MakeUSDate(trim(Request.QueryString("txtDateAddedEnd")))
	  txtPriceStart		= trim(Request.QueryString("txtPriceStart"))
	  txtPriceEnd 		= trim(Request.QueryString("txtPriceEnd"))
	  txtSale			= trim(Request.QueryString("txtSale"))
	End If 

CurrencyISO = getCurrencyISO(Session.LCID)
If iConverion = 1 Then Response.Write  "<script language=""JavaScript"" src=""http://www.oanda.com/cgi-bin/fxcommerce/fxcommerce.js?user=" & sUserName & """></script>"
	Set rsSearch = Server.CreateObject("ADODB.RecordSet")
	' -------------------------------------------
	' RecordSet Paging Setup --------------------
	' -------------------------------------------
	if Application("AppName")="StoreFrontAE" then
	   dim iSubCat
	   iSubCat = sSubCat
	   SQL = getProductSQLAE(txtsearchParamType, txtsearchParamTxt, txtsearchParamCat, txtsearchParamMan, _
	   txtsearchParamVen, txtDateAddedStart, txtDateAddedEnd, txtPriceStart, txtPriceEnd, txtSale,iSubCat,iLevel)

	   if txtsearchParamCat <> "ALL" then
	      sNextLevel = getSubCategoryList(ilevel,sSubcat)
           if trim(snextlevel) <> ""  then
            iLevel = Ilevel + 1
	       end if
	   End if   
	 
	else	
    	 SQL = getProductSQL(txtsearchParamType, txtsearchParamTxt, txtsearchParamCat, txtsearchParamMan, _
	     txtsearchParamVen, txtDateAddedStart, txtDateAddedEnd, txtPriceStart, txtPriceEnd, txtSale)
	     
	end if
	  
	If varDebug = 1 Then Response.Write SQL & "<br><br>"
	With rsSearch
		.CursorLocation = adUseClient
		.CacheSize = iVarPageSize
		.MaxRecords = iMaxRecords
		.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText 
		.PageSize = iVarPageSize
	End With
'Response.Write SQL
'Response.end
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
	'create arrays for display
        'arrProduct = rsSearch.GetRows(iVarPageSize)
 if rsSearch.BOF and rsSearch.EOF then
 iSearchRecordCount=0
 else
     
         arrProduct = rsSearch.GetRows()
      
    iSearchRecordCount=ubound(arrProduct,2) + 1 
	iNumOfPages = Int(iSearchRecordCount / iPageSize)
end if	
	If CInt(iNumOfPages+1) = CInt(iPage) Then iVarPageSize = iSearchRecordCount - (iNumOfPages * iPageSize) 
    'Response.Write "<BR>iVarPageSize " & iVarPageSize & "<BR>iSearchRecordCount - (iNumOfPages * iPageSize)" 
	'Response.Write "<BR>"  & iSearchRecordCount &  "-" & "(" & iNumOfPages & " * " & iPageSize & ") = " & iSearchRecordCount - (iNumOfPages * iPageSize) 
	'Corrects Number of Pages if there is overflow less then records per page 
	
If iSearchRecordCount mod iPageSize <> 0 Then iNumOfPages = iNumOfPages + 1                    	

If rsSearch.bof=false and rsSearch.eof=true then
  rsSearch.movefirst
end if
	
If NOT rsSearch.EOF Then rsSearch.AbsolutePage = CInt(iPage)
	
	' Create Attribute Record Sets for product on page
	SQL = getAttributeSQL(rsSearch, iVarPageSize, iPage)

		
	If varDebug = 1 Then Response.Write SQL & "<br><br>"	
	Set rsSearchAtt = Server.CreateObject("ADODB.RecordSet")
	If SQL <> "" Then 
		rsSearchAtt.Open SQL, cnn, adOpenKeyset, adLockReadOnly, adCmdText
		SQL = getAttributeDetailSQL(rsSearchAtt)
		If varDebug = 1 Then Response.Write  SQL & "<br><br>"
		If SQL <> "" Then
			Set rsSearchAttDetail = Server.CreateObject("ADODB.RecordSet")
			rsSearchAttDetail.Open SQL, cnn, adOpenKeyset, adLockReadOnly, adCmdText
		End If 
	End If 
	
	If txtsearchParamCat = "ALL" Then
		txtCatName = "All " & C_CategoryNameP
	Else
        If Not rsSearch.EOF Then
		  if Application("AppName")<> "StoreFrontAE" then  
	     	txtCatName = rsSearch.Fields("catName")
		  else
		      Dim arrTemp
		      on error resume next
		      if txtsearchParamCat = "ALL" then
				arrTemp = GetFullPath(rsSearch.Fields("CatHierarchy"),1,iSubCat)
	     	  else
               	arrTemp = GetFullPath(rsSearch.Fields("CatHierarchy"),1,iSubCat)
 	     	  end if  
		          txtCatName = arrtemp
           end if
		Else
		   if Application("AppName")<> "StoreFrontAE" then  
	     	set rsCat = Server.CreateObject("ADODB.RecordSet")
			SQL = getCategorySQL(txtsearchParamCat)
			rsCat.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
			txtCatName = rsCat.Fields("catName")
			closeObj(rsCat)
		   else
		      on error resume next
		      txtCatName =GetFullPath(Request.Item("txtCatName"),1,iSubCat) 
           end if      
		   
		
		End If 
    End If
	If txtsearchParamTxt = "" Then txtsearchParamTxt = "*"

%>
<html>

<head>
<SCRIPT language="javascript" src="SFLib/incae.js"></SCRIPT>
<SCRIPT language="javascript" src="SFLib/sfCheckErrors.js"></SCRIPT>
<SCRIPT language="javascript" src="SFLib/sfEmailFriend.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
					<!--
					function drillmore(vId)
				 	  {
				 	    if(vId != "Drill me")
					    {
					     document.drilMe.subcat.value =vId
					     document.drilMe.submit() 
					    }
					  }
					//-->
    			</SCRIPT>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%= C_STORENAME %>-SF Search Engine Output Page</title>


<!--Header Begin -->
<link rel="stylesheet" href="sfCSS.css" type="text/css">
</head>
<body bgproperties="fixed"  link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">

<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd"width="<%= C_WIDTH %>" align="center">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="2">
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
          <td>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
		      <tr>
		        <td align="left" class="tdMiddleTopBanner" width = "20%">&nbsp;</td>
		        <td align="center" class="tdMiddleTopBanner"><center>Search Results</center></td>
		        <td align="right" class="tdMiddleTopBanner" width="20%"><a href="<%= C_HomePath %>order.asp"><img src="buttons/ocheckout.gif" border="0" align="right" valign="top" alt="Check Out" width="101" height="19"></a></td>
		      </tr>
	        </table>
	      </td>
        </tr>
        <tr>
          <!-- ### SEARCH RESULT QUERY OUTPUT ::: BEGIN ### -->    
          <td align="middle" class="tdBottomTopBanner">Search for: <b><%= txtsearchParamTxt %></b> in <%= C_CategoryNameS %>: <b><%= txtCatName %></b><%
    If txtsearchParamMan <> "ALL" Then
		SQL = "SELECT mfgName FROM sfManufacturers WHERE mfgID = " & txtsearchParamMan
		Set rsManufacturer = Server.CreateObject("ADODB.RecordSet")
		rsManufacturer.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
		Response.Write ", " & C_ManufacturerNameS & ": <b>" & rsManufacturer.Fields("mfgName") & "</b>"
		closeObj(rsManufacturer)
	End If 
	If txtsearchParamVen <> "ALL" Then
		SQL = "SELECT vendName FROM sfVendors WHERE vendID = " & txtsearchParamVen
		Set rsVendor = Server.CreateObject("ADODB.RecordSet")
		rsVendor.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
		Response.Write ", " & C_VendorNameS & ": <b>" & rsVendor.Fields("vendName") & "</b>"
		closeObj(rsVendor)
	End If 

    %>
            <br><b><%= iSearchRecordCount %></b> Records Returned <b><%=iVarPageSize%></b> Records Displayed.<br><a href="search.asp"><b>New Search</b></a>|<a href="advancedsearch.asp"><b>Advanced Search</b></a></td>
          <!-- ### SEARCH RESULT QUERY OUTPUT ::: END ### -->            
        </tr>
           <%If trim(sNextLevel) <> ""   Then%> 
           
        <tr>

         <td align="middle" class="tdBottomTopBanner">
           <form name=Drilldown>
           <select size="1" name="subcat" style="<%= C_FORMDESIGN %>" onchange="drillmore(this.options[this.options.selectedIndex].value)" >
             <option value="Drill me">Refine Search</option><%= sNextLevel %>
           </select>
        </form>
         </td>
        
        </tr>
        
          <%End if%>

        
        <tr>
          <td class="tdContent2">
            <table border="0" width="100%">
              <tr>
	   <td width="100%"><p align="center"><font class="Content_Small">
           
       <%
        If iNumOfPages <> 1 And iNumOfPages <> 0 Then  Response.Write bottomPaging(iPage, iPageSize, iSearchRecordCount, iNumOfPages, "Search")
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
    If rsSearch.EOF Then
%>    
<table border=0 width=100&#37;>
<tr>
<td><center><font class='Content_Large'>Sorry, No Matching Records Returned!</font></center></td></tr>
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
 
        arrProduct = rsSearch.GetRows(iVarPageSize)

        If Not rsSearchAtt.EOF Then 
			arrAtt = rsSearchAtt.GetRows
			irsSearchAttRecordCount = rsSearchAtt.RecordCount-1
			If Not rsSearchAttDetail.EOF Then 
				arrAttDetail = rsSearchAttDetail.GetRows
				irsSearchAttDetailRecordCount = rsSearchAttDetail.RecordCount-1 
			End If
		End If
	   
        For iRec = 0 to iVarPageSize - 1 
			   SearchResults_CheckInventoryTracked 'SFAE  b2   
		    'sfAE
		       If iRec > (iVarPageSize-1) then EXIT FOR 'SFAE b2 
			
                ' Set Default Image if none specified for product
              
                If arrProduct(2, iRec) <> "" Then
                    txtImagePath = arrProduct(2, iRec)
                Else
                    txtImagePath = ""
                End If
      			icounter = 1 

%>				
			  

                 <form method="post" name="<%= arrProduct(0, iRec)%>" action="<%= C_HomePath %>addproduct.asp" onSubmit="this.QUANTITY.quantityBox=true;return sfCheck(this);">                
                <table border="0" width="100%" class="tdContent2" >

                  <input TYPE="hidden" NAME="PRODUCT_ID" VALUE="<%= arrProduct(0, iRec)%>">         
                
                  <tr>

                    <% If iDesign = "1" Then %>            
                    <td width="30%" align="center"><% If Trim(arrProduct(3, iRec)) <> "" Then %><a href="<%= replace(arrProduct(3, iRec)," ","+") %>"><% End If %><%If txtImagePath <> "" Then%><img border="1" src="<%= txtImagePath %>"><%ElseIf Trim(arrProduct(3, iRec)) <> "" Then %>Link <%End If%><% If Trim(arrProduct(3, iRec)) <> "" Then %></a><% End If %><br>
                    </td>
                    <% ElseIf iDesign = "3" And (iDesignCounter / 2) = Int(iDesignCounter / 2) Then%>
		            <td width="30%" align="center"><% If Trim(arrProduct(3, iRec)) <> "" Then %><a href="<%= replace(arrProduct(3, iRec)," ","+") %>"><% End If %><%If txtImagePath <> "" Then%><img border="1" src="<%= txtImagePath %>"><%ElseIf Trim(arrProduct(3, iRec)) <> "" Then %>Link<%End If%><% If Trim(arrProduct(3, iRec)) <> "" Then %></a><% End If %><br>
                    </td>
                    <% End If %>
                    
                  <td width="70%" valign="top"> <b><font class='Content_Large'><%= arrProduct(1, iRec) %></font></b><br>
                      <b><%= C_Description %>:</b>&nbsp;<%If arrProduct(11, iRec) <> "" Then%><%= arrProduct(11, iRec) %><%Else%><%= arrProduct(8, iRec) %><%End If%><br>
                      <b><%= C_Price %>:</b>&nbsp;
			          <% If iConverion = 1 Then
					If arrProduct(5, iRec) = 1 Then 
							Response.Write "<i><strike><script>document.write(""" & FormatCurrency(arrProduct(4, iRec)) & " = ("" + OANDAconvert(" & trim(arrProduct(4, iRec)) & "," & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script></strike></i><br>"
							Response.Write "<font color=#FF0000><b>" & C_SPrice & ": <script>document.write(""" & FormatCurrency(arrProduct(6, iRec)) & " = ("" + OANDAconvert(" & trim(arrProduct(6, iRec)) & ", " & chr(34) & CurrencyISO & chr(34) & ")+ "")"");</script></b></font><br>"
							Response.Write "<font color=#FF0000><i>" & C_YSave & ": <script>document.write(""" & FormatCurrency(CDbl(arrProduct(4, iRec))-CDbl(arrProduct(6, iRec))) & " = ("" + OANDAconvert(" & trim(CDbl(arrProduct(4, iRec))-CDbl(arrProduct(6, iRec))) & ", " & chr(34) & CurrencyISO & chr(34) & ")+ "")"");</script></i></font><br>"
					Else
							Response.Write "<script>document.write(""" & FormatCurrency(arrProduct(4, iRec)) & " = ("" + OANDAconvert(" & trim(arrProduct(4, iRec)) & ", " & chr(34) & CurrencyISO & chr(34) & ")+ "")"");</script>"
					End If 
			   Else
					If arrProduct(5, iRec) = 1 Then 
							Response.Write "<i><strike>" & FormatCurrency(arrProduct(4, iRec)) & "</strike></i><br>"
							Response.Write "<font color=#FF0000><b>" & C_SPrice & ": " & FormatCurrency(arrProduct(6, iRec)) & "</b></font><br>"
							Response.Write "<font color=#FF0000><i>" & C_YSave & ": " & FormatCurrency(CDbl(arrProduct(4, iRec))-CDbl(arrProduct(6, iRec))) & "</i></font><br>"
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
                      <p align="center"><center><%= C_Quantity %>:<input style="<%= C_FORMDESIGN %>"  type="text" name="QUANTITY" title="Quantity" size="3" value="1" ><br>
                      <%SearchResults_GetGiftWrap arrProduct(0, iRec) 'SFAE%>
                        <input type="image" name="AddProduct" border="0" src="buttons/addtocart3.gif" alt="Add To Cart" width="107" height="28">
                        <br>
                      <% If iSaveCartActive = 1 Then
                      if Application("AppName")="StoreFrontAE" then%>
                      <input type="image" name="SaveCart" border="0" src="<%= C_BTN02 %>" alt="Add to Wish List">
                      <%else%>
                      <input type="image" name="SaveCart" border="0" src="<%= C_BTN02 %>" alt="Save To Cart">
                      <%end if%>
                      <%
        End if
        If iEmailActive = 1 Then
        %>
                      <a href="javascript:emailFriend('<%= server.urlencode(arrProduct(0, iRec)) %>')"><img border="0" src="<%= C_BTN24 %>" alt="Email a Friend"></a>
                      <% End If %>
                      </center></p>
                    </td>
                    <%  If iDesign = "2" Then %>            
		            <td width="30%" align="center"><a href="<%= replace(arrProduct(3, iRec)," ","+") %>"><%If txtImagePath <> "" Then%><img border="1" src="<%= txtImagePath %>"><%ElseIf Trim(arrProduct(3, iRec)) <> "" Then %>Link<%End If%></a><br>
                    </td>
                    <%  ElseIf iDesign = "3" And (iDesignCounter / 2) <> Int(iDesignCounter / 2) Then %>
		            <td width="30%" align="center"><% If Trim(arrProduct(3, iRec)) <> "" Then %><a href="<%= replace(arrProduct(3, iRec)," ","+") %>"><% End If %><%If txtImagePath <> "" Then%><img border="1" src="<%= txtImagePath %>"><%ElseIf Trim(arrProduct(3, iRec)) <> "" Then %>Link<%End If%><% If Trim(arrProduct(3, iRec)) <> "" Then %></a><% End If %><br>
                    </td>
                 
                    <%  End If %>
                  </tr>
                  <tr>
                    <td width="100%" colspan="2"><hr noshade color="#000000" size="1" width="90%">
                    </td>
                  </tr>    
</table></form>
                <%             
'				Response.Write "</table></form>"
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
If iNumOfPages <> 1 And iNumOfPages <> 0 Then Response.Write bottomPaging(iPage, iPageSize, iSearchRecordCount, iNumOfPages, "Search")
                ' -------------------------------------------   
                ' SEARCH RESULT PAGING OUTPUT ::: END -------
                ' -------------------------------------------
%>
                        </font></p>
                      </td>
                    </tr>
                  </table>
    
<!--Footer begin-->
                  <!--#include file="foot.txt"-->
                </table>
              </td>
            </tr>
          </table>
<%           if Application("AppName")="StoreFrontAE" then 
                  %>
                    <form method="get" action="search_results.asp"  name="drilMe"> 
                     <input type="hidden" name="txtsearchParamType" value="<%= txtsearchParamType%>">
                    <input type="hidden" name="txtsearchParamCat" value="<%= txtsearchParamCat%>">
                    <input type="hidden" name="txtsearchParamMan" value="<%= txtsearchParamMan%>">
                    <input type="hidden" name="txtsearchParamVen" value="<%= txtsearchParamVen%>">
                    <input type="hidden" name="txtDateAddedStart" value="<%= txtDateAddedStart%>">
                    <input type="hidden" name="txtPriceEnd" value="<%= txtPriceEnd%>">
                    <input type="hidden" name="txtPriceStart" value="<%= txtPriceStart%>">
                    <input type="hidden" name="txtSale" value="<%= txtSale%>">
                    <input type="hidden" name="txtsearchParamTxt" value="<%=txtsearchParamTxt%>">
                    <input type="hidden" name="txtFromSearch" value="">
                    <input type="hidden" name="subcat" value="">
                     <input type="hidden" name="iLevel" value="<%= iLevel%>">   
                     <input type="hidden" name="txtCatName" value="<%= iLevel%>">   
                   </form>
               

          <%end if  %>

       </body>

    </html>
<!--Footer End-->

        <%
	' Object Cleanup
	closeObj(rsSearch)
	closeObj(rsSearchAtt)
	closeObj(rsSearchAttDetail)
	closeObj(cnn)

%>






















