<!--
/*******************************************************************************************
COMMON VALIDATIONS
*******************************************************************************************/

function validateWords(fld) {
	var error = "";
	
	    var iChars = "@#$%^&*+=[]\\\;/{}|\<>'";
        for (var i = 0; i < fld.value.length; i++) {
                if (iChars.indexOf(fld.value.charAt(i)) != -1) {
					fld.style.background = 'Yellow';
                	error = "- Special characters are not allowed. Please remove them. \n";                
        		}
        }
	return error;  
}


/*******************************************************************************************
DAMAGED ITEM FORM - add_damage.asp & update_damage.asp
*******************************************************************************************/

function validateDamagedItem(fld) {
    var error = "";
 
    if (fld.value.length == 0) {
        fld.style.background = 'Yellow'; 
        error = "- Damaged Item has not been filled in.\n"
    } else {
        fld.style.background = 'White';
    }
    return error;  
}

function validateSerialNo(fld) {
    var error = "";
 
    if (fld.value.length == 0) {
        fld.style.background = 'Yellow'; 
        error = "- Serial No has not been filled in.\n"
    } else {
        fld.style.background = 'White';
    }
    return error;  
}


function validateDamagedItemFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	reason += validateDamagedItem(theForm.txtDamagedItem); 
	reason += validateSerialNo(theForm.txtSerialNo);    	
	
  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);
    	
		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Submit';
		
		return true;
    }
}

/*******************************************************************************************

*******************************************************************************************/


//-->