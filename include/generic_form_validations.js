<!--
/*******************************************************************************************
GENERIC FORM VALIDATIONS
*******************************************************************************************/
var dtCh= "/";
var minYear=2010;
var maxYear=2020;

function isInteger(s){
	var i;
    for (i = 0; i < s.length; i++){   
        // Check that current character is number.
        var c = s.charAt(i);
        if (((c < "0") || (c > "9"))) return false;
    }
    // All characters are numbers.
    return true;
}

function stripCharsInBag(s, bag){
	var i;
    var returnString = "";
    // Search through string's characters one by one.
    // If character is not in bag, append to returnString.
    for (i = 0; i < s.length; i++){   
        var c = s.charAt(i);
        if (bag.indexOf(c) == -1) returnString += c;
    }
    return returnString;
}

function daysInFebruary (year){
	// February has 29 days in any year evenly divisible by four,
    // EXCEPT for centurial years which are not also divisible by 400.
    return (((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0))) ? 29 : 28 );
}
function DaysArray(n) {
	for (var i = 1; i <= n; i++) {
		this[i] = 31
		if (i==4 || i==6 || i==9 || i==11) {this[i] = 30}
		if (i==2) {this[i] = 29}
   } 
   return this;
}

function isDate(dtStr){
	var daysInMonth = DaysArray(12)
	var pos1=dtStr.indexOf(dtCh)
	var pos2=dtStr.indexOf(dtCh,pos1+1)
	var strDay=dtStr.substring(0,pos1)
	var strMonth=dtStr.substring(pos1+1,pos2)
	var strYear=dtStr.substring(pos2+1)
	strYr=strYear
	if (strDay.charAt(0)=="0" && strDay.length>1) strDay=strDay.substring(1)
	if (strMonth.charAt(0)=="0" && strMonth.length>1) strMonth=strMonth.substring(1)
	for (var i = 1; i <= 3; i++) {
		if (strYr.charAt(0)=="0" && strYr.length>1) strYr=strYr.substring(1)
	}
	month=parseInt(strMonth)
	day=parseInt(strDay)
	year=parseInt(strYr)
	if (pos1==-1 || pos2==-1){
		alert("The date format should be: DD/MM/YYYY")
		return false
	}
	if (strMonth.length<1 || month<1 || month>12){
		alert("Please enter a valid month")
		return false
	}
	if (strDay.length<1 || day<1 || day>31 || (month==2 && day>daysInFebruary(year)) || day > daysInMonth[month]){
		alert("Please enter a valid day")
		return false
	}
	if (strYear.length != 4 || year==0 || year<minYear || year>maxYear){
		alert("Please enter a valid 4 digit year between "+minYear+" and "+maxYear)
		return false
	}
	if (dtStr.indexOf(dtCh,pos2+1)!=-1 || isInteger(stripCharsInBag(dtStr, dtCh))==false){
		alert("Please enter a valid date")
		return false
	}
	return true;
}

function validateEmptyField(fld,errormsg) {
    var error = "";
 
    if (fld.value.length == 0) {
        fld.style.background = 'Yellow'; 
        error = "- " + errormsg + " is empty.\n"
    } else {
        fld.style.background = 'White';
    }
    return error;  
}

function validateSpecialCharacters(fld,errormsg) {
	var error = "";
	
	var iChars = "@#$%^&*+=[]\\\;{}|\<>";
    for (var i = 0; i < fld.value.length; i++) {
    	if (iChars.indexOf(fld.value.charAt(i)) != -1) {
			fld.style.background = 'Yellow';
            error = "- Special characters are not allowed in " + errormsg + ". Please remove them. \n";                
        }
    }
	return error;  
}

function validateCheckBox(fld,errormsg) {
    var error = "";
 	if (!fld.checked) {   
        fld.style.background = 'Yellow'; 
		error = "- You must accept our " + errormsg + ".\n"
    } else {
        fld.style.background = 'White';
    }
    return error;  
}

function trim(s)
{
  return s.replace(/^\s+|\s+$/, '');
}

function validateEmail(fld) {
    var error="";
    var tfld = trim(fld.value);                        // value of field with whitespace trimmed off
    var emailFilter = /^[^@]+@[^@.]+\.[^@]*\w\w$/ ;
    var illegalChars= /[\(\)\<\>\,\;\:\\\"\[\]]/ ;
   
    if (fld.value == "") {
        fld.style.background = 'Yellow';
        error = "- Email address is empty.\n";
    } else if (!emailFilter.test(tfld)) {              //test email for illegal characters
        fld.style.background = 'Yellow';
        error = "- Please enter a valid email address.\n";
    } else if (fld.value.match(illegalChars)) {
        fld.style.background = 'Yellow';
        error = "- Email address contains illegal characters.\n";
    } else {
        fld.style.background = 'White';
    }
    return error;
}

function validateNumeric(fld,errormsg){
	var error = "";
	
	if (isNaN(parseInt(fld.value))){
		fld.style.background = 'Yellow'; 
		error = "- Only numbers are allowed in " + errormsg + ". \n";
	} else {
        fld.style.background = 'White';
    }
    return error;
}

function validateDate(fld,errormsg) {
    var error = "";
	
	var today = new Date();
	today.setDate(today.getDate()+2);
		
	var dayfield = fld.value.split("/")[0];
	var monthfield = fld.value.split("/")[1];
	var yearfield = fld.value.split("/")[2];
	
    if (isDate(fld.value) == false) {
        fld.style.background = 'Yellow'; 
        error = "- " + errormsg + " must be in DD/MM/YYYY.\n"
    } else {
        fld.style.background = 'White';
    }
    return error;  
}

function validateFinalDeadline(fld) {
    var error = "";
	
	var today = new Date();
	
	var dayfield = fld.value.split("/")[0];
	var monthfield = fld.value.split("/")[1];
	var yearfield = fld.value.split("/")[2];
	var deadline = new Date(yearfield, monthfield-1, dayfield);
	
	var firstdayfield = form_add_project.txtFirstDeadline.value.split("/")[0];
	var firstmonthfield = form_add_project.txtFirstDeadline.value.split("/")[1];
	var firstyearfield = form_add_project.txtFirstDeadline.value.split("/")[2];
	var firstdeadline = new Date(firstyearfield, firstmonthfield-1, firstdayfield);
	
	firstdeadline.setDate(firstdeadline.getDate()+1);
	
    if (isDate(fld.value) == false) {
        fld.style.background = 'Yellow'; 
        error = "- Printing / publishing deadline must be in DD/MM/YYYY.\n"
	} else if (deadline < today) {
        fld.style.background = 'Yellow';
        error = "- Printing / publishing deadline must be greater than todays date.\n";
	} else if (deadline < firstdeadline) {
        fld.style.background = 'Yellow';
        error = "- Printing / publishing deadline must be greater than first draft deadline.\n";			
    } else {
        fld.style.background = 'White';
    }
    return error;
}

function validateFirstDeadline(fld) {
    var error = "";
	
	var today = new Date();
	
	var dayfield = fld.value.split("/")[0];
	var monthfield = fld.value.split("/")[1];
	var yearfield = fld.value.split("/")[2];
	var first_deadline = new Date(yearfield, monthfield-1, dayfield);
	
    if (isDate(fld.value) == false) {
        fld.style.background = 'Yellow'; 
        error = "- First draft deadline must be in DD/MM/YYYY.\n"
	} else if (first_deadline < today) {
        fld.style.background = 'Yellow';
        error = "- First draft deadline must be greater than todays date.\n";		
    } else {
        fld.style.background = 'White';
    }
    return error;
}

function validateDeadline(fld) {
    var error = "";
	
	var today = new Date();
	
	var dayfield = fld.value.split("/")[0];
	var monthfield = fld.value.split("/")[1];
	var yearfield = fld.value.split("/")[2];
	var deadline = new Date(yearfield, monthfield-1, dayfield);
	
    if (isDate(fld.value) == false) {
        fld.style.background = 'Yellow'; 
        error = "- Printing / publishing deadline must be in DD/MM/YYYY.\n"
	} else if (deadline < today) {
        fld.style.background = 'Yellow';
        error = "- Printing / publishing deadline must be greater than todays date.\n";		
    } else {
        fld.style.background = 'White';
    }
    return error;
}

function completeProgress(form) {
	if (form.cboProgress.value == "100") {
		alert("100%!");
		form.cboStatus.value == "0";
	} else {	
		form.cboStatus.value == "";
	}
}

/*******************************************************************************************
LIMIT TEXT AREA
*******************************************************************************************/
 
function limitText(limitField, limitCount, limitNum) {
	if (limitField.value.length > limitNum) {
		limitField.value = limitField.value.substring(0, limitNum);
	} else {
		limitCount.value = limitNum - limitField.value.length;
	}
}

/*******************************************************************************************
BACK BUTTON
*******************************************************************************************/
function goBack()  {
  window.history.back()
}
//-->