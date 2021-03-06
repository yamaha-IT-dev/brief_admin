<!--#include file="../include/connection_it.asp " -->
<!--#include file="class/clsLoanTransfer.asp " -->
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<!--[if lt IE 9]>
	<script src="//oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
	<script src="//oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
<![endif]-->
<title>New Loan Transfer</title>
<link rel="stylesheet" href="bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" href="css/header.css">
<link rel="stylesheet" href="//maxcdn.bootstrapcdn.com/font-awesome/4.3.0/css/font-awesome.min.css">
<link rel="stylesheet" href="//ajax.googleapis.com/ajax/libs/jqueryui/1.11.2/themes/smoothness/jquery-ui.css">
<script src="//code.jquery.com/jquery.js"></script>
<script src="//ajax.googleapis.com/ajax/libs/jqueryui/1.11.2/jquery-ui.min.js"></script>
<script src="bootstrap/js/bootstrap.js"></script>
<script src="../include/generic_form_validations.js"></script>
<script>
$(function() {
	var availableTags = 
	["7AG001 Adam Goodwin","7AY000 Alex Yabsley","7AD000 Alf Demasi","7BE001 Bradley Eustace","7CH001 Chris Herring","7CT001 Chris Thew","7CJ000 Craig Jones","7CS001 Craig Stack","7DH000 Damien Henderson","7DH002 Damien Henderson Road Show","7DU000 Dion Durante","7DC001 Dominica","7EO000 Eric Ong","7EM001 Euan McInnes","7GL001 Gavin Liu","7GN001 George Nasr","7GL000 Gavin Liu","7JW001 Jaclyn Williams","7JG000 Jamie Goff","7AUDW0 Jamie Goff AUDW","7LB000 Leon Blaher","7MC001 Mark Condon","7MT000 Mat Taylor","7MD000 Matt Dawkins","7ML003 Matt Livingstone","7MH000 Mick Hughes","7MIRG0 Miroslaw Galcynski","7NB001 Nathan Biggin","7PW000 Paul Wheeler","7BP000 Peter Beveridge","7RW001 Russell Wykes","7SB000 Sam Bertolini","9SLA01 Drew Morrow","7SMC01 Shaun McMahon","7STA00 Stephen Allan","7SL001 Steve Legg","7SVR01 Steven Vranch","7TM000 Terry McMahon","7YME01 Altona","7YME02 Balwyn","7YME09 Baulkham Hills","7YME05 Carnegie","7YME10 Chatswood","7YME07 Glen Waverley","7YME12 Morley","7YME08 Strathmore", "7RW555 AV Marketing Room", "7SMC55 Product Review"];
	
	$( "#txtRecipient" ).autocomplete({
		source: availableTags
	});
});
</script>
<script>
function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	reason += validateEmptyField(theForm.txtAccountCode,"Account Code");
	reason += validateSpecialCharacters(theForm.txtAccountCode,"Account Code");
	
	reason += validateEmptyField(theForm.txtModelNo,"Model No");
	//reason += validateSpecialCharacters(theForm.txtModelNo,"Model No");
	
	reason += validateEmptyField(theForm.txtSerialNo,"Serial No");
	reason += validateSpecialCharacters(theForm.txtSerialNo,"Serial No");
	
	reason += validateNumeric(theForm.txtQty,"Qty");
	//reason += validateSpecialCharacters(theForm.txtQty,"Qty");
	
	reason += validateEmptyField(theForm.txtRecipient,"Recipient");

  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);    	
		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Add';
		return true;
    }
}
</script>
<%
sub main	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		if trim(request("Action")) = "Add" then
			dim traAccountCode, traModelNo, traSerialNo, traQty, traRecipient, traCreatedBy
			
			traAccountCode 	= Server.HTMLEncode(Replace(Trim(Request("txtAccountCode")),"'","''"))
			traModelNo 		= Server.HTMLEncode(Replace(Trim(Request("txtModelNo")),"'","''"))
			traSerialNo 	= Server.HTMLEncode(Replace(Trim(Request("txtSerialNo")),"'","''"))
			traQty 			= Server.HTMLEncode(Replace(Trim(Request("txtQty")),"'","''"))
			traRecipient 	= Server.HTMLEncode(Left(Trim(Request("txtRecipient")),6))
			traCreatedBy 	= Trim(session("logged_username"))
						
			'response.write (traAccountCode & "," & traModelNo & "," & traSerialNo & "," & traQty & "," & traRecipient & "," & traCreatedBy & "," & session("newOrderNo") & "," & session("newOrderLine"))
			call addTransfer(traAccountCode, traModelNo, traSerialNo, session("newOrderNo"), session("newOrderLine"), traQty, traRecipient, traCreatedBy)
			'call addLoanTransfer(traAccountCode, traModelNo, traSerialNo, traQty, traRecipient, traCreatedBy)
		end if
	end if
end sub

call main
%>
</head>
<body>
<div class="blog-masthead">
    <div class="container">
    	<nav class="blog-nav">
          <a class="blog-nav-item" href="loan_summary.asp"><i class="fa fa-home fa-lg"></i></a>
          <a class="blog-nav-item active">Transfer</a>
          <a class="blog-nav-item" href="loan-sale.asp">Sale</a>
        </nav>
	</div>
</div>
<div class="container"> 
  <br>
  <ol class="breadcrumb">
    <li><a href="loan-transfer.asp">Loan Transfer</a></li>
    <li class="active">New Transfer</li>
  </ol>
  <h1 class="page-header"><i class="fa fa-plus"></i> New Loan Transfer</h1>
  <h3>Order no: <%= session("newOrderNo") & "-" & session("newOrderLine")%></h3>
  <form action="" method="post" name="form_add_sales" id="form_add_sales" onsubmit="return validateFormOnSubmit(this)">
    <div class="form-group">
      <label for="cboProduct">Account code<font color="red">*</font>:</label>
      <input type="text" class="form-control" name="txtAccountCode" id="txtAccountCode" maxlength="6" size="6" value="<%= session("newAccountCode") %>" placeholder="Account" pattern=".{6,}" required title="6 characters minimum" />
    </div>
    <div class="form-group">
      <label for="cboProduct">Model no<font color="red">*</font>:</label>
      <input type="text" class="form-control" name="txtModelNo" id="txtModelNo" maxlength="20" size="20" value="<%= session("newModelNo") %>" placeholder="Model no" pattern=".{2,}" required title="2 characters minimum" />
    </div>
    <div class="form-group">
      <label for="txtSerialNo">Serial no<font color="red">*</font>:</label>
      <input type="text" class="form-control" name="txtSerialNo" id="txtSerialNo" maxlength="11" size="10" value="<%= session("newSerialNo") %>" placeholder="Serial no" pattern=".{4,}" required title="4 characters minimum" />
    </div>
    <div class="form-group">
      <label for="txtQty">Qty<font color="red">*</font>:</label>
      <input type="text" class="form-control" name="txtQty" id="txtQty" maxlength="2" size="2" placeholder="Qty" value="1" required />
    </div> 
    <div class="form-group">
      <label for="txtRecipient">Recipient's Account code<font color="red">*</font>:</label>
      <input type="text" class="form-control" name="txtRecipient" id="txtRecipient" maxlength="6" placeholder="Recipient" pattern=".{6,}" required />
    </div> 
    <!--<div class="form-group"> 
      <label for="cboRecipient">Recipient<font color="red">*</font>:</label>
      <select name="cboRecipient" id="cboRecipient" class="form-control">
          <option value="damienh">Damien</option>
          <option value="georgen">George</option>
          <option value="russellw">Russell</option>
          <option value="shaunm">Shaun</option>                    
        </select>
    </div>-->
    <div class="form-group">
      <input type="hidden" name="Action" />
      <input type="submit" name="submit" id="submit" class="btn btn-default" value="Transfer" />
    </div>
  </form>    
</div>
</body>
</html>