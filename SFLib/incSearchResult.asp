<%
'--------------------------------------------------------------------
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.3

'@FILENAME: incSearCHrESULTS.asp
	 
'Access Version

'@DESCRIPTION:   functions to return search results

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws as an unpublished work, and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO

'Modified 11/20/01 
'Storefront Ref#'s: 131 'JF
	
Function getManufacturersList(iValue)
	' Variable Declarations
	Dim rsManufacturersList
	Dim sList
	Dim intId
	
	' Object Creation and Query
	Set rsManufacturersList = Server.CreateObject("ADODB.RecordSet")
	rsManufacturersList.Open "select mfgID, mfgName from sfManufacturers ORDER BY mfgName", cnn, adOpenForwardOnly, adLockOptimistic, adCmdText	'adCmdTable
	sList = ""
	If iValue = "" Then
		Do While Not rsManufacturersList.EOF
			sList = sList & "<OPTION value=" & rsManufacturersList.Fields("mfgID") & ">" & rsManufacturersList.Fields("mfgName") & "</OPTION>"
			rsManufacturersList.MoveNext
		Loop
	Else
		Do While Not rsManufacturersList.EOF
			intId = trim(rsManufacturersList.Fields("mfgID"))
			If iValue = intId Then
				slist = slist & "<OPTION value=" & intId & " selected>" & rsManufacturersList.Fields("mfgName") & "</OPTION>"
			Else
				slist = slist & "<OPTION value=" & intId & ">" & rsManufacturersList.Fields("mfgName") & "</OPTION>"
			End If 
			rsManufacturersList.MoveNext
		Loop				
	End If
	'object cleanup
	rsManufacturersList.Close 
	Set rsManufacturersList = nothing
	
	'return value
	getManufacturersList = sList
End Function

Function getVendorList(iValue)

	' Variable Declarations
	Dim rsVendorList
	Dim sList
	Dim intId
	
	' Object Creation and Query
	Set rsVendorList = Server.CreateObject("ADODB.RecordSet")
	rsVendorList.Open "select vendID,vendName from sfVendors ORDER BY vendName", cnn, adOpenForwardOnly, adLockOptimistic, adCmdText	'adCmdTable
	sList = ""
	If iValue = "" Then
		Do While Not rsVendorList.EOF
			sList = sList & "<OPTION value=" & rsVendorList.Fields("vendID") & ">" & rsVendorList.Fields("vendName") & "</OPTION>"
			rsVendorList.MoveNext
		Loop
	Else
		Do While Not rsVendorList.EOF
			intId = trim(rsVendorList.Fields("vendID"))
			If iValue = intId Then
				sList = sList & "<OPTION value=" & intId & " selected>" & rsVendorList.Fields("vendName") & "</OPTION>"
			Else
				sList = sList & "<OPTION value=" & intId & ">" & rsVendorList.Fields("vendName") & "</OPTION>"
			End If 
			rsVendorList.MoveNext
		Loop				
	End If
	
	'object cleanup
	rsVendorList.Close
	Set rsVendorList = nothing 
	
	'return value
	getVendorList = sList
End Function

'--------------------------------------------------------------------
' Function : getCategoryList
' This returns the category list in HTML format for dropdown box.
'--------------------------------------------------------------------	
Function getCategoryList(iValue)
	
	' Variable Declarations
	Dim rsCategoryList
	Dim sList
	Dim intId
	
	' Object Creation and Query
	Set rsCategoryList = Server.CreateObject("ADODB.RecordSet")
	'rsCategoryList.Open "SELECT DISTINCT catID, catName  FROM sfCategories INNER JOIN sfProducts ON sfCategories.catID = sfProducts.prodCategoryId ORDER BY CatName", cnn, adOpenForwardOnly, adLockOptimistic, adCmdText

	rsCategoryList.Open "SELECT DISTINCT catID, catName  FROM sfCategories Order By CatName", cnn, adOpenForwardOnly, adLockOptimistic, adCmdText '#319
	sList = ""
	If iValue = "" Then
		Do While Not rsCategoryList.EOF
			sList = sList & "<OPTION value=" & rsCategoryList.Fields("catID") & ">" & rsCategoryList.Fields("catName") & "</OPTION>"
			rsCategoryList.MoveNext
		Loop
	Else
		Do While Not rsCategoryList.EOF
			intId = trim(rsCategoryList.Fields("catID"))
			If iValue = intId Then
				sList = sList & "<OPTION value=" & intId & " selected>" & rsCategoryList.Fields("catName") & "</OPTION>"
			Else
				sList = sList & "<OPTION value=" & intId & ">" & rsCategoryList.Fields("catName") & "</OPTION>"
			End If 
			rsCategoryList.MoveNext
		Loop				
	End If

	' Object Cleanup
	rsCategoryList.Close
	Set rsCategoryList = nothing 

	' Return Value	
	getCategoryList = sList
End Function
Function getSubCategoryList(ilevel,subcatID)
	'on error resume next
	' Variable Declarations
	Dim rsSubCategoryList
	Dim sList,iLen
	dim sSQl,sHierarchy
	dim MainCatID
'Response.Write subcatID
 	
if instr(subcatid, "-") > 0 then
  getSubCategoryList =""
  exit function
end if  
	
if ilevel = 1 And subCatID <> "ALL" then 
 MainCatID = setSubcatId(subCatId,"subcatCategoryId","subcatCategoryId")
elseif ilevel > 1 And subCatID <> "ALL" then
 MainCatID = setSubcatId(subCatId,"subCatId","subcatCategoryId")
end if  
'subCatID = setSubcatId(subCatId,"subcatCategoryId")
if  subCatID <> "ALL" then 
	 sHierarchy = getCatHierarchy(subcatID)
	 iLen = len(sHierarchy)
	if Ilevel = 1 then
	 sSQl = "SELECT Distinct subcatCategoryId, subcatID ,SubcatName,Bottom  FROM sfSub_Categories Where Depth = " & iLevel & " And subcatCategoryId = " & MainCatID
	 
	else
	 sSQl = "SELECT Distinct subcatCategoryId, subcatID ,SubcatName,Bottom  FROM sfSub_Categories Where Depth = " & iLevel & " And subcatCategoryId = " & MainCatID & _
	 " AND LEFT(CatHierarchy," & iLen & ") = '" & sHierarchy & "'"
    end if
	Set rsSubCategoryList = Server.CreateObject("ADODB.RecordSet")
		rsSubCategoryList.Open sSQL , cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
	
	if rsSubCategoryList.EOF = false and rsSubCategoryList.BOF =false then
		'sList = sList & "<OPTION value=" & "Drill" & " selected>" & "Drill Down" & "</OPTION>"
		Do While Not rsSubCategoryList.EOF
		   if rsSubCategoryList.Fields("Bottom") = 1 then
		 	sList = sList & "<OPTION value=" & rsSubCategoryList.Fields("subcatID") & "-bottom>" & rsSubCategoryList.Fields("subcatName") & "</OPTION>"
		   else    
			sList = sList & "<OPTION value=" & rsSubCategoryList.Fields("subcatID") & ">" & rsSubCategoryList.Fields("subcatName") & "</OPTION>"
		   end if 	
			rsSubCategoryList.MoveNext                                   
		Loop
		
	    getSubCategoryList = sList
	else
	 getSubCategoryList =""
	end if 
else
     sSQl = "SELECT DISTINCT catID, catName  FROM sfCategories " 
   'sSQl ="SELECT sfCategories.catName, sfSub_Categories.subcatName, sfSub_Categories.subcatID " _
'    & "FROM sfCategories RIGHT JOIN sfSub_Categories ON sfCategories.catID = sfSub_Categories.subcatCategoryId " _
 '   & " Where sfSub_Categories.Depth = " & iLevel
 Set rsSubCategoryList = Server.CreateObject("ADODB.RecordSet")
		rsSubCategoryList.Open sSQL , cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
		if rsSubCategoryList.EOF = false and rsSubCategoryList.BOF =false then
		Do While Not rsSubCategoryList.EOF
		sList = sList & "<OPTION value=" & rsSubCategoryList.Fields("catID") & ">" & rsSubCategoryList.Fields("catName") & "</OPTION>"
		rsSubCategoryList.MoveNext                                   
		Loop
		
	    getSubCategoryList = sList
	else
	 getSubCategoryList =""
	end if   
    
 end if 
	'Response.Write sSql
	' Object Cleanup
	rsSubCategoryList.Close
	Set rsSubCategoryList = nothing 
	' Return Value	
	
End Function


Function getProductSQLAE(searchParamType, searchParamTxt, searchParamCat, searchParamMan, searchParamVen, DateAddedStart, DateAddedEnd, PriceStart, PriceEnd, Sale,subCatID,Ilevel)
if instr(subcatid, "bottom") > 0 then
 subcatid = left(subcatID,instr(subcatId,"-")-1)
end if

Dim upperLim, SQL, counter, txtArray
	searchParamTxt = Replace(searchParamTxt, "*", "")
sSubcat =  subCatID 


if iLevel = 1 and subCatID <> "ALL"   then 
  subCatId = setSubcatId(subCatId,"subcatCategoryId","subcatCategoryId")

end if

if iLevel = 1 then sALLSUB = "ALL"
if subCatID = "ALL" Then
'Response.Write "1  "
'SQL = " SELECT sfProducts.ProdID, sfProducts.prodName, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodPrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice,sfSub_Categories.CatHierarchy," _
'& "sfProducts.prodDescription, sfProducts.prodAttrNum, sfProducts.prodCategoryId, sfProducts.prodShortDescription" _
'& " FROM (sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) INNER JOIN sfSub_Categories ON sfSubCatdetail.subcatCategoryId = sfSub_Categories.SubcatID WHERE "
SQL = "SELECT ProdID, prodName, prodImageSmallPath, prodLink, prodPrice, prodSaleIsActive, prodSalePrice, ProdID, prodDescription, prodAttrNum, prodCategoryId, prodShortDescription " _
			& "FROM sfProducts WHERE "


elseif instr(subCatID,";") > 0  Then
' Response.Write "2  "
 ' 
   SQL = " SELECT sfProducts.ProdID, sfProducts.prodName, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodPrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice,sfSub_Categories.CatHierarchy," _
& "sfProducts.prodDescription, sfProducts.prodAttrNum, sfProducts.prodCategoryId, sfProducts.prodShortDescription" _
& " FROM (sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) INNER JOIN sfSub_Categories ON sfSubCatdetail.subcatCategoryId = sfSub_Categories.SubcatID  "
SQL = SQL & " WHERE (sfSubCatDetail.subcatCategoryId IN (Select subcatCategoryId From sfSubCatDetail Where subcatCategoryId = " & GetSubCatIDs(sSubCat) & ")) AND " 
else

  SQL = " SELECT sfProducts.ProdID, sfProducts.prodName, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodPrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice,sfSub_Categories.CatHierarchy," _
& "sfProducts.prodDescription, sfProducts.prodAttrNum, sfProducts.prodCategoryId, sfProducts.prodShortDescription" _
& " FROM (sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) INNER JOIN sfSub_Categories ON sfSubCatdetail.subcatCategoryId = sfSub_Categories.SubcatID "
   if sALLSUB <> "ALL" Then
 '   Response.Write "3-1  "
     
    SQL = SQL & " WHERE (sfSubCatDetail.subcatCategoryId IN (Select subcatCategoryId From sfSubCatDetail Where subcatCategoryId = " & GetSubCatIDs(sSubCat) & ")) AND " 
   else
'   Response.Write "3-2  "
   
   SQL = SQL & " WHERE sfSubCatDetail.subcatCategoryId IN (Select subcatID From sfSub_Categories Where subcatCategoryId= " & sSubCat & ") AND " 
   end if
end if 

if searchParamTxt <> "" Then 
    If searchParamType = "ALL" Then 
	'	Response.Write " A  <BR>"
		txtArray = split(searchParamTxt, " ")
		upperLim = Ubound(txtArray)
		
		If searchParamTxt <> "" Then
			For counter=0 to (upperLim-1)
				SQL = SQL &  "  (sfProducts.prodName LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodShortDescription LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodID LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodDescription LIKE '%" & txtArray(counter) & "%') AND "
			Next
			SQL = SQL & " (sfProducts.prodName LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodShortDescription LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodID LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodDescription LIKE '%" & txtArray(counter) & "%') AND " 
		     
		End If
			 
	'	Response.Write " A-1  <BR>"	
	Elseif searchParamType = "ANY" Then
	    'Response.Write " B  <BR>"
		txtArray = split(searchParamTxt, " ")
		upperLim = Ubound(txtArray)
		SQL=SQL & "("
		If searchParamTxt <> "" Then
			For counter=0 to (upperLim-1)
				SQL = SQL &  " (sfProducts.prodName LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodShortDescription LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodID LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodDescription LIKE '%" & txtArray(counter) & "%') OR "
			Next
			SQL = SQL & " (sfProducts.prodName LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodShortDescription LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodID LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodDescription LIKE '%" & txtArray(counter) & "%')"
		End If 
		SQL=SQL & ")"
	   'Response.Write " B-1  <BR>"	
	    SQL = SQL & " And "		
	Elseif searchParamType = "Exact" Then
	  ' Response.Write " C  <BR>"
	   if searchParamTxt <> "" Then SQL = SQL & " (sfProducts.prodName LIKE '%" & searchParamTxt & "%' OR sfProducts.prodShortDescription LIKE '%" & searchParamTxt & "%' OR sfProducts.prodID ='" & searchParamTxt & "' OR sfProducts.prodDescription LIKE '%" & searchParamTxt & "%') AND "
	    
	else 
	 'Response.Write " D  <BR>"
	 SQL = SQL & " WHERE "
	End If
end if	
	'If searchParamCat = "ALL" Then
'		If searchParamTxt = "" Then
'			SQL = SQL & " prodCategoryId > 0"
'		Else 
'			SQL = SQL & " AND prodCategoryId > 0"
'		End If 
'	Else
'		If searchParamTxt = "" Then 
'			SQL = SQL & " prodCategoryId = " & searchParamCat
'		Else
'			SQL = SQL & " AND prodCategoryId = " & searchParamCat
'		End If 
'	End If 

	If searchParamMan = "ALL" Then 
		SQL = SQL &  "  sfProducts.prodManufacturerId > 0"
	Else
		SQL = SQL & "  sfProducts.prodManufacturerId = " & searchParamMan
	End If 
	If searchParamVen = "ALL" Then 
		SQL = SQL & " AND sfProducts.prodVendorId > 0"
	Else
		SQL = SQL & " AND sfProducts.prodVendorId = " & searchParamVen
	End If	
	
    If  DateAddedEnd <> ""  then
     DateAddedEnd = dateAdd("d",1,DateAddedEnd)
    end if  
    
   
    If DateAddedStart <> "" And DateAddedEnd <> "" Then SQL = SQL & " AND sfProducts.prodDateAdded BETWEEN #" & CDate(DateAddedStart) & "# AND #" & CDate(DateAddedEnd) & "# "
	If DateAddedStart <> "" And DateAddedEnd = "" Then SQL = SQL & " AND sfProducts.prodDateAdded > #" & CDate(DateAddedStart) & "# " 
	If DateAddedStart = "" And DateAddedEnd <> "" Then SQL = SQL & " AND sfProducts.prodDateAdded < #" & CDate(DateAddedEnd) & "# "  
	If PriceStart <> "" And PriceEnd <> "" Then SQL = SQL & " AND (CDbl(sfProducts.prodPrice) BETWEEN " & CDbl(PriceStart) & " AND " & CDbl(PriceEnd) & " or (sfProducts.prodSaleIsActive=1 and CDbl(sfProducts.prodSalePrice) BETWEEN " & CDbl(PriceStart) & " AND " & CDbl(PriceEnd) & ")) " 
	If PriceStart <> "" And PriceEnd = "" Then SQL = SQL & " AND (CDbl(sfProducts.prodPrice) >= " & CDbl(PriceStart) & " or (sfProducts.prodSaleIsActive=1 and CDbl(sfProducts.prodSalePrice) >= " & CDbl(PriceStart) & ")) " 
	If PriceStart = "" And PriceEnd <> "" Then SQL = SQL & " AND (CDbl(sfProducts.prodPrice) <= " & CDbl(PriceEnd) & "  or (sfProducts.prodSaleIsActive=1 and CDbl(sfProducts.prodSalePrice) <= " & CDbl(PriceEnd) & ")) " 
	If Sale <> "" Then SQL = SQL & " AND sfProducts.prodSaleIsActive = 1 "
	SQL = SQL & " AND sfProducts.prodEnabledIsActive = 1 "
	
	getProductSQLAE = SQL
End Function
Function getProductSQL(searchParamType, searchParamTxt, searchParamCat, searchParamMan, searchParamVen, DateAddedStart, DateAddedEnd, PriceStart, PriceEnd, Sale)
	Dim upperLim, SQL, counter, txtArray
	searchParamTxt = Replace(searchParamTxt, "*", "") 
	SQL = "SELECT ProdID, prodName, prodImageSmallPath, prodLink, prodPrice, prodSaleIsActive, prodSalePrice, catName, prodDescription, prodAttrNum, prodCategoryId, prodShortDescription " _
			& "FROM sfProducts INNER JOIN sfCategories ON sfProducts.prodCategoryId = sfCategories.catID WHERE "
	'create where statement
	If searchParamType = "ALL" Then
		txtArray = split(searchParamTxt, " ")
		upperLim = Ubound(txtArray)
		If searchParamTxt <> "" Then
			For counter=0 to (upperLim-1)
				SQL = SQL &  " (prodName LIKE '%" & txtArray(counter) & "%' OR prodShortDescription LIKE '%" & txtArray(counter) & "%' OR prodID LIKE '%" & txtArray(counter) & "%' OR prodDescription LIKE '%" & txtArray(counter) & "%') AND "
			Next
			SQL = SQL & " (prodName LIKE '%" & txtArray(counter) & "%' OR prodShortDescription LIKE '%" & txtArray(counter) & "%' OR prodID LIKE '%" & txtArray(counter) & "%' OR prodDescription LIKE '%" & txtArray(counter) & "%') " 
		End If
		 	
	Elseif searchParamType = "ANY" Then
		txtArray = split(searchParamTxt, " ")
		upperLim = Ubound(txtArray)
		SQL=SQL & "("
		If searchParamTxt <> "" Then
			For counter=0 to (upperLim-1)
				SQL = SQL &  " (prodName LIKE '%" & txtArray(counter) & "%' OR prodShortDescription LIKE '%" & txtArray(counter) & "%' OR prodID LIKE '%" & txtArray(counter) & "%' OR prodDescription LIKE '%" & txtArray(counter) & "%') OR "
			Next
			SQL = SQL & " (prodName LIKE '%" & txtArray(counter) & "%' OR prodShortDescription LIKE '%" & txtArray(counter) & "%' OR prodID LIKE '%" & txtArray(counter) & "%' OR prodDescription LIKE '%" & txtArray(counter) & "%')"
		End If 
			SQL=SQL & ")"
	Elseif searchParamType = "Exact" Then
			If searchParamTxt <> "" Then SQL = SQL & " (prodName LIKE '%" & searchParamTxt & "%' OR prodShortDescription LIKE '%" & searchParamTxt & "%' OR prodID ='" & searchParamTxt & "' OR prodDescription LIKE '%" & searchParamTxt & "%') "
	End If
	
	If searchParamCat = "ALL" Then
		If searchParamTxt = "" Then
			SQL = SQL & " prodCategoryId > 0"
		Else 
			SQL = SQL & " AND prodCategoryId > 0"
		End If 
	Else
		If searchParamTxt = "" Then 
			SQL = SQL & " prodCategoryId = " & searchParamCat
		Else
			SQL = SQL & " AND prodCategoryId = " & searchParamCat
		End If 
	End If 
	If searchParamMan = "ALL" Then 
		SQL = SQL &  " AND prodManufacturerId > 0"
	Else
		SQL = SQL & " AND prodManufacturerId = " & searchParamMan
	End If 
	If searchParamVen = "ALL" Then 
		SQL = SQL & " AND prodVendorId > 0"
	Else
		SQL = SQL & " AND prodVendorId = " & searchParamVen
	End If	
	If  DateAddedEnd <> ""  then
     DateAddedEnd = dateAdd("d",1,DateAddedEnd)
    end if  
    
	
	If DateAddedStart <> "" And DateAddedEnd <> "" Then SQL = SQL & " AND prodDateAdded BETWEEN #" & CDate(DateAddedStart) & "# AND #" & CDate(DateAddedEnd) & "# "
	If DateAddedStart <> "" And DateAddedEnd = "" Then SQL = SQL & " AND prodDateAdded > #" & CDate(DateAddedStart) & "# " 
	If DateAddedStart = "" And DateAddedEnd <> "" Then SQL = SQL & " AND prodDateAdded < #" & CDate(DateAddedEnd) & "# "  
	'djp log 201
	If PriceStart <> "" And PriceEnd <> "" Then SQL = SQL & " AND (CDbl(sfProducts.prodPrice) BETWEEN " & CDbl(PriceStart) & " AND " & CDbl(PriceEnd) & " or (sfProducts.prodSaleIsActive=1 and CDbl(sfProducts.prodSalePrice) BETWEEN " & CDbl(PriceStart) & " AND " & CDbl(PriceEnd) & ")) " 
	If PriceStart <> "" And PriceEnd = "" Then SQL = SQL & " AND (CDbl(sfProducts.prodPrice) >= " & CDbl(PriceStart) & " or (sfProducts.prodSaleIsActive=1 and CDbl(sfProducts.prodSalePrice) >= " & CDbl(PriceStart) & ")) " 
	If PriceStart = "" And PriceEnd <> "" Then SQL = SQL & " AND (CDbl(sfProducts.prodPrice) <= " & CDbl(PriceEnd) & "  or (sfProducts.prodSaleIsActive=1 and CDbl(sfProducts.prodSalePrice) <= " & CDbl(PriceEnd) & ")) " 

	
	If Sale <> "" Then SQL = SQL & " AND prodSaleIsActive = 1 "
	SQL = SQL & " AND prodEnabledIsActive = 1 "
	getProductSQL = SQL
End Function

Function getAttributeSQL(rsSearch, iPageSize, iPage)
	Dim counter, rs, SQL
	If Not rsSearch.EOF Then
		' Clone rsSearch so it is not manipulated by the Function
		Set rs = Server.CreateObject("ADODB.RecordSet")
		Set rs = rsSearch.Clone 
	
		
	       rs.AbsolutePosition  = rsSearch.AbsolutePosition   
		SQL = "SELECT attrID, attrName, attrProdID FROM sfAttributes WHERE "
		
		For counter = 1 to iPageSize
			SQL = SQL & "attrProdId = '" & rs.Fields("prodID") & "' OR "
			rs.MoveNext
		Next
		closeObj(rs)
		SQL = Mid(SQL, 1, len(SQL)-3)
		getAttributeSQL = SQL & " ORDER BY attrName"
	Else
		getAttributeSQL = ""
	End If 
End Function

Function getAttributeDetailSQL(rs)
	If Not rs.EOF Then
		Dim SQL
		SQL = "SELECT attrdtID, attrdtAttributeId, attrdtName, attrdtPrice, attrdtType, attrdtOrder FROM sfAttributeDetail WHERE "
		Do While Not rs.EOF
			SQL = SQL & " attrdtAttributeId = " & rs.Fields("attrID") & " OR "
			rs.MoveNext
		Loop
		rs.MoveFirst
		SQL = Mid(SQL, 1, len(SQL)-3)
		getAttributeDetailSQL = SQL & " ORDER BY attrdtOrder"
	Else
		getAttributeDetailSQL = ""
	End If 
End Function

Function getCategorySQL(txtCategory)
	getCategorySQL = "SELECT catName FROM sfCategories WHERE catID = " & txtCategory
End Function

Function bottomPaging(iPage, iPageSize, iSearchRecordCount, iNumOfPages, sFromPage)
	Dim txtPage, icounter, output, iStart, iEnd, sLink,iLoop
	output = ""
	
	if Application("AppName")<> "StoreFrontAE" then
	txtPage = "&txtsearchParamTxt=" & Server.URLEncode(txtsearchParamTxt) & "&txtsearchParamType=" & txtsearchParamType & "&txtsearchParamCat=" _
         & txtsearchParamCat & "&txtsearchParamMan=" & txtsearchParamMan & "&txtsearchParamVen=" & txtsearchParamVen _
         & "&txtDateAddedStart=" & txtDateAddedStart & "&txtDateAddedEnd=" & txtDateAddedEnd _
         & "&txtPriceStart=" & txtPriceStart & "&txtPriceEnd=" & txtPriceEnd & "&txtSale=" & txtSale
	
	else
		txtPage =""
		  For iLoop = 1 to Request.QueryString.Count 
		    If lcase(Request.QueryString.Key(iLoop)) <> "page" then  
		     txtPage = txtPage &  "&" & Request.QueryString.Key(iLoop) & "=" & Request.QueryString.Item(iLoop)  
		    End if 
		  Next
		 txtPage = replace(txtPage," ","+")
	end if
	 
	If sFromPage = "SalesPage" Then
		sLink = "<a href=salespage.asp?PAGE="
		txtPage = ""
	ElseIf sFromPage = "NewProducts" Then
		sLink = "<a href=newproduct.asp?PAGE="
		txtPage = ""
	Else
		sLink = "<a href=search_results.asp?PAGE="
	End If 
	
	If iPage <> "1" Then
	    output = output &  "<font color=black>&lt;&lt; " & sLink & iPage - 1 & txtPage &">Previous</a> | "
	Else 
	    output = output &  "<font color=silver>&lt;&lt; Previous</font> | "
	End If
	'Two cases, less than ten pages or more than ten pages total                
	If iNumOfPages > 10 Then 'Four cases inbeded 
		'First case, first ten pages
		If iPage <= 10 Then
			If iPage <> 10 Then
				For icounter = 1 to 9
				    If iCounter = CInt(iPage) Then
				        output = output &  iCounter & " | "
				    Else
				        output = output & sLink & icounter & txtPage & ">" & icounter & "</a> | "
				    End If                      
				Next
					output = output &  sLink & icounter & txtPage & ">" & icounter & "...</a> | "
			Else
				If iNumOfPages < 20 Then
					For icounter = 10 to iNumOfPages
					    If iCounter = CInt(iPage) Then
					        output = output &  iCounter & " | "
					    Else
					        output = output &  sLink & icounter & txtPage & ">" & icounter & "</a> | "
					    End If                      
					Next
				Else
					For icounter = 10 to 19
					    If iCounter = CInt(iPage) Then
					        output = output &  iCounter & " | "
					    Else
					        output = output &  sLink & icounter & txtPage & ">" & icounter & "</a> | "
					    End If                      
					Next
						output = output &  sLink & icounter & txtPage & ">" & icounter & "...</a> | "
				End If 
			End If 
		'rare case when the number of pages is divisable to records per page
		ElseIf iPage <= (iNumOfPages - (iNumOfPages mod 10)) AND iPage > iNumOfPages-iPageSize AND iNumOfPages mod iPageSize = 0 Then  
			For icounter = iNumOfPages-9 to iNumOfPages
			    If iCounter = CInt(iPage) Then
			        output = output &  iCounter & " | "
			    Else
			        output = output &  sLink & icounter & txtPage & ">" & icounter & "</a> | "
			    End If                      
			Next
		'Case for the inbetween areas ie 10-20 20-30... 
		ElseIf iPage < (iNumOfPages - (iNumOfPages mod 10)) Then
			If iPage mod 10 = 0 Then
				iStart = iPage
				iEnd = iPage + 9
			Else
				iStart = (iPage - (iPage mod 10))
				iEnd = iStart + 9
			End If  
			For icounter = iStart to iEnd
			    If iCounter = CInt(iPage) Then
			        output = output &  iCounter & " | "
			    Else
			        output = output &  sLink & icounter & txtPage & ">" & icounter & "</a> | "
			    End If                      
			Next
			output = output &  sLink & icounter & txtPage & ">" & icounter & "...</a> | "
		'Case when last few pages is less then ten
		Else
			For icounter = (iPage - (iPage mod 10)) to iNumOfPages
			    If iCounter = CInt(iPage) Then
			        output = output &  iCounter & " | "
			    Else
			        output = output &  sLink & icounter & txtPage & ">" & icounter & "</a> | "
			    End If                      
			Next
		End If
	'If total number of pages is less than ten
	Else             
		For icounter = 1 to iNumOfPages
		    If icounter = CInt(iPage) Then
		        output = output &  iCounter & " | "
		    Else
		        output = output &  sLink & icounter & txtPage & ">" & icounter & "</a> | "
		    End If                      
		Next
	End If 
	                
	If CInt(iNumOfPages) <> CInt(iPage) Then 
	    output = output &  sLink & iPage + 1 & txtPage &">Next</a><font color=black> &gt;&gt;"
	Else
	    output = output &  "<font color=silver>Next &gt;&gt;</font>"
	End If 
bottomPaging = output
End Function

Function GetSubCatIDs(vID)
dim rstSubCat,sSQL ,sHierarchy,iLen
dim tempID,sCriteria
Set rstSubCat = Server.CreateObject("ADODB.RecordSet")

if Instr(vId,";")> 0 then
 sSql = "Select SubCatId  From sfSub_Categories Where subcatCategoryId = " & Cint(left(vID,len(instr(vid,";")-1))) _
 & " AND left(CatHierarchy,4)= '" & "none" & "'"
 rstSubCat.Open sSql, cnn,adOpenStatic ,adLockReadOnly , adCmdText	
 
  vID = rstSubCat("SubCatID")
   rstSubCat.Close 
end if
sSql = "Select CatHierarchy,hasprods From sfSub_Categories Where SubcatID = " & vID
rstSubCat.Open sSql, cnn,adOpenStatic ,adLockReadOnly , adCmdText	
if rstSubCat.EOF =true and rstSubCat.BOF = true then
else
  if rstSubCat("hasprods") = 1 then
   GetSubCatIDs =vID
  else 
   sHierarchy = rstSubCat("CatHierarchy")
   iLen = len(sHierarchy)
   sSql = "Select SubCatID,CatHierarchy,hasprods From sfSub_Categories Where left(CatHierarchy," & iLen & ") = '" & sHierarchy & "' AND Hasprods = 1 AND Depth > 0"
   rstSubCat.Close
   rstSubCat.Open sSql, cnn,adOpenStatic ,adLockReadOnly , adCmdText	
     if rstSubCat.EOF =true and rstSubCat.BOF = true then
      GetSubCatIDs = vID
     else
        tempID =""
        sCriteria = " OR subcatCategoryId ="
       while rstSubCat.EOF =false
         tempID = tempID & rstSubCat("SubCatID") & sCriteria  
         rstSubCat.MoveNext 
       wend     
        tempID = left(tempID,len(tempID) - len(sCriteria))
        GetSubCatIDs = tempID 
     end if 
  end if 
end if
if GetSubCatIDs ="" then
 GetSubCatIDs = vID
end if

End Function

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

function setSubcatId(iCatId,sCriteria,returnField)
dim rst,sSql
Set rst = Server.CreateObject("ADODB.RecordSet")
' sSql = "Select subCatID,subcatCategoryId from sfSub_Categories where subcatCategoryId = "  & iCatid
 sSql = "Select subCatID,subcatCategoryId from sfSub_Categories where " & sCriteria & " = "  & iCatid
 rst.Open ssql,cnn,3,3,1
 if rst.eof = true and rst.eof = true then
   setSubcatId= icatId
 else  
  setSubcatId = rst(returnField) 
 end if
 rst.Close
 set rst = nothing
end function
function getCatHierarchy(vID)
dim rst,sSql
Set rst = Server.CreateObject("ADODB.RecordSet")
 sSql = "Select CatHierarchy from sfSub_Categories where subcatID = "  & vID
 rst.Open ssql,cnn,3,3,1
 if rst.eof = true then
   getCatHierarchy ="vID"
 else
  getCatHierarchy = rst("CatHierarchy") 
 end if
 rst.Close
 set rst = nothing
end function

%>



