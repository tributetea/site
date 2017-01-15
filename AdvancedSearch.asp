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

'@FILENAME: advancedsearch.asp
 
'

'@DESCRIPTION: Product Search Page

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO
 dim sSubCategories
 dim FrontPage_Form1 

%>

<html>

<head>
<script language="javascript" src="SFLib/sfCheckErrors.js"></script>
<script language="javascript">
/******************************************************************
   convert_date()
   
   Function to convert supplied dates to format - dd/mm/yyyy.
	Valid input dates = 
		ddmmyy, ddmmmyy, ddmmyyyy, ddmmmyyyy,
		d/m/yy, dd/m/yy, d/mm/yy, dd/mm/yy, d/mmm/yy, dd/mmm/yy,
		d/m/yyyy, dd/m/yyyy, d/mm/yyyy, dd/mm/yyyy, d/mmm/yyyy, dd/mmm/yyyy
	Valid date seperators =
		'-','.','/',' ',':','_',','
		
	Calls convert_month()
			invalid_date()
			validate_date()
			validate_year()
     
   Author: Simon Kneafsey 
   Email: simonkneafsey@hotmail.com
   WebSite: www.simonkneafsey.co.uk
   Date Created: 4/9/00
   
   Notes: Please feel free to use/edit this script.  If you do please keep my comments and details 
   intact and notify me via a quick Email to the address above.  Enjoy!
*******************************************************************/

function convert_date(field1)
{
var fLength = field1.value.length; // Length of supplied field in characters.
var divider_values = new Array ('-','.','/',' ',':','_',','); // Array to hold permitted date seperators.  Add in '\' value
var array_elements = 7; // Number of elements in the array - divider_values.
var day1 = new String(null); // day value holder
var month1 = new String(null); // month value holder
var year1 = new String(null); // year value holder
var divider1 = null; // divider holder
var outdate1 = null; // formatted date to send back to calling field holder
var counter1 = 0; // counter for divider looping 
var divider_holder = new Array ('0','0','0'); // array to hold positions of dividers in dates
var s = String(field1.value); // supplied date value variable

//If field is empty do nothing
if ( fLength == 0 ) {
   return true;
}

// Deal with today or now
if ( field1.value.toUpperCase() == 'NOW' || field1.value.toUpperCase() == 'TODAY' ) {
   
	var newDate1 = new Date();
	
  		if (navigator.appName == "Netscape") {
    		var myYear1 = newDate1.getYear() + 1900;
  		}
  		else {
  			var myYear1 =newDate1.getYear();
  		}
  
	var myMonth1 = newDate1.getMonth()+1;  
	var myDay1 = newDate1.getDate();
	field1.value = myDay1 + "/" + myMonth1 + "/" + myYear1;
	fLength = field1.value.length;//re-evaluate string length.
	s = String(field1.value)//re-evaluate the string value.
}

//Check the date is the required length
if ( fLength != 0 && (fLength < 6 || fLength > 11) ) {
	invalid_date(field1);
	return false;   
	}

// Find position and type of divider in the date
for ( var i=0; i<3; i++ ) {
	for ( var x=0; x<array_elements; x++ ) {
		if ( s.indexOf(divider_values[x], counter1) != -1 ) {
			divider1 = divider_values[x];
			divider_holder[i] = s.indexOf(divider_values[x], counter1);
		   //alert(i + " divider1 = " + divider_holder[i]);
			counter1 = divider_holder[i] + 1;
			//alert(i + " counter1 = " + counter1);
			break;
		}
 	}
 }

// if element 2 is not 0 then more than 2 dividers have been found so date is invalid.
if ( divider_holder[2] != 0 ) {
   invalid_date(field1);
	return false;   
}

// See if no dividers are present in the date string.
if ( divider_holder[0] == 0 && divider_holder[1] == 0 ) { 
   
		//continue processing
		if ( fLength == 6 ) {//ddmmyy
   		//day1 = field1.value.substring(0,2);
     	//	month1 = field1.value.substring(2,4);
     	month1 = field1.value.substring(0,2);
     	day1 = field1.value.substring(2,4);
  			year1 = field1.value.substring(4,6);
  			if ( (year1 = validate_year(year1)) == false ) {
   			invalid_date(field1);
				return false; 
				}
			}
			
		else if ( fLength == 7 ) {//mmmddy
   		//day1 = field1.value.substring(0,2);
  		//	month1 = field1.value.substring(2,5);
  		 month1= field1.value.substring(0,3);
  			day1 = field1.value.substring(3,5);
  			year1 = field1.value.substring(5,7);
  			if ( (month1 = convert_month(month1)) == false ) {
   			invalid_date(field1);
				return false; 
				}
  			if ( (year1 = validate_year(year1)) == false ) {
   			invalid_date(field1);
				return false; 
				}
			}
		else if ( fLength == 8 ) {//mmddyyyy
   		//day1 = field1.value.substring(0,2);
  		//	month1 = field1.value.substring(2,4);
  		 month1= field1.value.substring(0,2);
  			day1 = field1.value.substring(2,4);
  			year1 = field1.value.substring(4,8);
			}
		else if ( fLength == 9 ) {//mmmddyyyy
   		//day1 = field1.value.substring(0,2);
  		//	month1 = field1.value.substring(2,5);
  		month1 = field1.value.substring(0,3);
  		day1 = field1.value.substring(3,5);
  			year1 = field1.value.substring(5,9);
  			if ( (month1 = convert_month(month1)) == false ) {
   			invalid_date(field1);
				return false; 
				}
			}
		
		if ( (outdate1 = validate_date(day1,month1,year1)) == false ) {
   		alert("The value " + field1.value + " is not a vaild date.\n\r" +  
			"Please enter a valid date in the format mm/dd/yyyy");
			field1.focus();
			field1.select();
			return false;
			}

		field1.value = outdate1;
		return true;// All OK
		}
		
// 2 dividers are present so continue to process	
if ( divider_holder[0] != 0 && divider_holder[1] != 0 ) { 	
  	//day1 = field1.value.substring(0, divider_holder[0]);
  	//month1 = field1.value.substring(divider_holder[0] + 1, divider_holder[1]);
  	 month1= field1.value.substring(0, divider_holder[0]);
  	day1 = field1.value.substring(divider_holder[0] + 1, divider_holder[1]);
  	//alert(month1);
  	year1 = field1.value.substring(divider_holder[1] + 1, field1.value.length);
	}

if ( isNaN(day1) && isNaN(year1) ) { // Check day and year are numeric
	invalid_date(field1);
	return false;  
   }

if ( day1.length == 1 ) { //Make d day dd
   day1 = '0' + day1;  
}

if ( month1.length == 1 ) {//Make m month mm
	month1 = '0' + month1;   
}

if ( year1.length == 2 ) {//Make yy year yyyy
   if ( (year1 = validate_year(year1)) == false ) {
   	invalid_date(field1);
		return false;  
		}
}

if ( month1.length == 3 || month1.length == 4 ) {//Make mmm month mm
   if ( (month1 = convert_month(month1)) == false) {
   	alert("month1" + month1);
   	invalid_date(field1);
   	return false;  
   }
}

// Date components are OK
if ( (day1.length == 2 || month1.length == 2 || year1.length == 4) == false) {
   invalid_date(field1);
   return false;
}

//Validate the date
if ( (outdate1 = validate_date(day1, month1, year1)) == false ) {
   alert("The value " + field1.value + " is not a vaild date.\n\r" +  
	"Please enter a valid date in the format mm/dd/yyyy");
	
	field1.focus();
	field1.select();

	return false;
}

// Redisplay the date in dd/mm/yyyy format
field1.value = outdate1;
return true;//All is well

}
/******************************************************************
   convert_month()
   
   Function to convert mmm month to mm month 
   
   Called by convert_date()    
   
   Author: Simon Kneafsey 
   Date Created: 4/9/00
   Email: simonkneafsey@hotmail.com
   WebSite: www.simonkneafsey.co.uk
   
   Notes:P lease feel free to use/edit this script.  If you do please keep my comments and details 
   intact and notify me via a quick Email to the address above.  Enjoy!
*******************************************************************/
function convert_month(monthIn) {

var month_values = new Array ("JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC");

monthIn = monthIn.toUpperCase(); 

if ( monthIn.length == 3 ) {
	for ( var i=0; i<12; i++ ) 
		{
   	if ( monthIn == month_values[i] ) 
   		{
			monthIn = i + 1;
			if ( i != 10 && i != 11 && i != 12 ) 
				{
   			monthIn = '0' + monthIn;
				}
			return monthIn;
			}
		}
	}

else if ( monthIn.length == 4 && monthIn == 'SEPT') {
   monthIn = '09';
   return monthIn;
	}
	
else {
	return false;
	} 
}
/******************************************************************
   invalid_date()
   
   If an entered date is deemed to be invalid, invali
   d_date() is called to display a warning message to
   the user.  Also returns focus to the date  in que
   stion and selects the date for edit.
        
   Called by convert_date()
   
   Author: Simon Kneafsey
   Date Created: 4/9/00
   Email: simonkneafsey@hotmail.com
   WebSite: www.simonkneafsey.co.uk
   
   Notes: Please feel free to use/edit this script.  If you do please keep my comments and details 
   intact and notify me via a quick Email to the address above.  Enjoy!
*******************************************************************/
function invalid_date(inField) 
{
alert("The value " + inField.value + " is not in a vaild date format.\n\r" + 
        "Please enter date in the format mm/dd/yyyy");
inField.focus();
inField.select();
return true   
}
/******************************************************************
   validate_date()
   
   Validates date output from convert_date().  Checks
   day is valid for month, leap years, month !> 12,.
   
   Author: Simon Kneafsey
   Date Created: 4/9/00
   Email: simonkneafsey@hotmail.com
   WebSite: www.simonkneafsey.co.uk
   
   Notes: Please feel free to use/edit this script.  If you do please keep my comments and details 
   intact and notify me via a quick Email to the address above.  Enjoy!
*******************************************************************/
function validate_date(day2, month2, year2)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   
{                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       
var DayArray = new Array(31,28,31,30,31,30,31,31,30,31,30,31);                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
var MonthArray = new Array("01","02","03","04","05","06","07","08","09","10","11","12");                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              
var inpDate = month2 + day2 + year2;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    
var filter=/^[0-9]{2}[0-9]{2}[0-9]{4}$/;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          

//Check mmddyyyy date supplied
if (! filter.test(inpDate))                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           
  {                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
  return false;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    
  }                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     
/* Check Valid Month */                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               
filter=/01|02|03|04|05|06|07|08|09|10|11|12/ ;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        
if (! filter.test(month2))                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         
  {                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               
  return false;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      
  }                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         
/* Check For Leap Year */                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            
var N = Number(year2);                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   
if ( ( N%4==0 && N%100 !=0 ) || ( N%400==0 ) )                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        
  	{                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     
   DayArray[1]=29;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     
  	}                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       
/* Check for valid days for month */                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  
for(var ctr=0; ctr<=11; ctr++)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        
  	{                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     
   if (MonthArray[ctr]==month2)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      
   	{                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    
      if (day2<= DayArray[ctr] && day2 >0 )                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       
        {
        inpDate = month2 + '/' + day2 + '/' + year2;       
        return inpDate;
        }                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 
      else                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             
        {                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
        return false;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                
        }                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                
   	}                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   
   }                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                
}
/******************************************************************
   validate_year()
   
   converts yy years to yyyy
   Uses a hinge date of 10
        < 10 = 20yy 
        => 10 = 19yy.
         
   Called by convert_date() before validate_date().
      
   Author: Simon Kneafsey 
   Date Created: 4/9/00
   Email: simonkneafsey@hotmail.com
   WebSite: www.simonkneafsey.co.uk
   
   Notes: Please feel free to use/edit this script.  If you do please keep my comments and details 
   intact and notify me via a quick Email to the address above.  Enjoy!
*******************************************************************/
function validate_year(inYear) 
{
if ( inYear < 10 ) 
	{
   inYear = "20" + inYear;
   return inYear;
	}
else if ( inYear >= 10 )
	{
   inYear = "19" + inYear;
   return inYear;
	}
else 
	{
	return false;
	}   
}
</script>

<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%= C_STORENAME %>-SF Advanced Search Engine Page</title>



<!--Header Begin -->
<link rel="stylesheet" href="sfCSS.css" type="text/css">
</head>
<%
 
%>
<body bgproperties="static" link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">

                <form method="get" action="search_results.asp"  name="FrontPage_Form1"> 
<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="<%= C_WIDTH %>" align="center">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="3">
        <tr>
            <td align="middle"  class="tdTopBanner"> 
              <%If C_BNRBKGRND = "" Then%>
              <%Else%>
              <img src="buttons/tt_blue.gif" border="0" width="275" height="36"> 
              <%End If%>
            </td>
        </tr>
<!--Header End -->        
            <td align="center"   class="tdMiddleTopBanner"> <font size="4"><font face="Arial, Helvetica, sans-serif">Advanced 
              Search</font></font></td>
        <tr>
            <td class="tdBottomTopBanner">Use the options below to perform a more 
              selective search of our product database. You can choose to search 
              only by items that have been added within a certain time range, 
              or search within a specific price range. 
          <tr>
              <td class="tdContent2">
                  
                  <table border="0" width="100%" cellpadding="4">
                    <tr>
                      
                  <td width="100%"> 
                    <div align="center"><b>Enter Keyword(s):</b><br>
                      &nbsp;&nbsp;&nbsp;&nbsp; 
                      <input style="<%= C_FORMDESIGN %>" name="txtsearchParamTxt" size="40" optional=true>
                    </div>
                    <p align="center"><b>Search using:</b><br>
                      &nbsp;&nbsp;&nbsp;&nbsp; 
                      <select size="1" style="<%= C_FORMDESIGN %>" name="txtsearchParamType">
                        <option selected value="ALL">All of the Keywords</option>
                        <option value="ANY">Any of the Keywords</option>
                        <option value="Exact">Exact Phrase</option>
                      </select>
                    </p>
                    <%If C_CategoryIsActive <> 0 Then%>
                    <p align="center"><b>Select a <%= C_CategoryNameS %>:</b><br>
                      &nbsp;&nbsp;&nbsp;&nbsp; 
                      <select style="<%= C_FORMDESIGN %>" size="1" name="txtsearchParamCat">
                        <option value="ALL">All <%= C_CategoryNameP %></option>
                        <%= getCategoryList(0) %>
                      </select>
                    </p>
                    <%Else%>
                    <div align="center">
                      <input type="hidden" name= "txtsearchParamCat" value="ALL">
                      <%End If           
			            If C_MFGIsActive <> 0 Then%>
                    </div>
                    <p align="center"><b>Select a <%= C_ManufacturerNameS %>:</b><br>
                      &nbsp;&nbsp;&nbsp;&nbsp; 
                      <select style="<%= C_FORMDESIGN %>" size="1" name="txtsearchParamMan">
                        <option value="ALL">All <%= C_ManufacturerNameP %></option>
                        <%= getManufacturersList(0) %>
                      </select>
                    </p>
                    <%Else%>
                    <div align="center">
                      <input type="hidden" name= "txtsearchParamMan" value="ALL">
                      <%End If%>
                      <%If C_VendorIsActive <> 0 Then%>
                    </div>
                    <p align="center"><b>Select a <%= C_VendorNameS %>:</b><br>
                      &nbsp;&nbsp;&nbsp;&nbsp; 
                      <select style="<%= C_FORMDESIGN %>" size="1" name="txtsearchParamVen">
                        <option value="ALL">All <%= C_VendorNameP %></option>
                        <%= getVendorList(0) %>
                      </select>
                    </p>
                    <%Else%>
                    <div align="center">
                      <input type="hidden" name= "txtsearchParamVen" value="ALL">
                      <%End If%>
                      <%If C_AddedIsActive <> 0 Then%>
                    </div>
                    <p align="center"><b>Added to Inventory Between:</b><br>
                      &nbsp;&nbsp;&nbsp;&nbsp;
                      <input optional=true type="text" style="<%= C_FORMDESIGN %>" name="txtDateAddedStart" size="8" onblur="javascript:convert_date(this)">
                      <b>And</b> 
                      <input optional=true type="text" style="<%= C_FORMDESIGN %>" name="txtDateAddedEnd" size="8" onblur="javascript:convert_date(this)">
                    </p>
                    <%End If%>
                    <%If C_PriceIsActive <> 0 Then%>
                    <p align="center"><b>Price Between:</b><br>
                      &nbsp;&nbsp;&nbsp;&nbsp;
                      <input optional=true number=true style="<%= C_FORMDESIGN %>" type="text" name="txtPriceStart" size="8">
                      <b>To</b> 
                      <input optional=true number=true style="<%= C_FORMDESIGN %>" type="text" name="txtPriceEnd" size="8">
                    </p>
                    <%End If%>
                    <%If C_SaleIsActive <> 0 Then%>
                    <p align="center">&nbsp;&nbsp;&nbsp;
                      <input optional=true type="checkbox" name="txtSale" value="1">
                      <b>Only Sale Items</b> 
                      <%End If%>
                    <p align="center"> 
                      <input type="image" name="btnSearch" src="buttons/search.gif" alt="Search" border="0" width="92" height="22">
                    </p>
			            </td>
                      </tr>
                    </table>
                    <input type="hidden" name="txtFromSearch" value="fromSearch">
                    <input type="hidden" name="iLevel" value="1">

	            </td>
<!--Footer begin-->
                <!--#include file="footer.txt"-->
              </table>
            </td>
          </tr>
        </table>
  
  </form>  
      
      </body>

    </html>
<!--Footer End-->
<%
 closeObj(cnn)     
%>




