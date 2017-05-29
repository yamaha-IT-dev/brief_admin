<!--#include file="../include/connection.asp " -->
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
<script src="//code.jquery.com/jquery.js"></script>
<script src="bootstrap/js/bootstrap.js"></script>
<script src="../include/generic_form_validations.js"></script>
<script>
function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	reason += validateEmptyField(theForm.txtAccountCode,"Account Code");
	reason += validateSpecialCharacters(theForm.txtAccountCode,"Account Code");
	
	reason += validateEmptyField(theForm.txtModelNo,"Model No");
	reason += validateSpecialCharacters(theForm.txtModelNo,"Model No");
	
	reason += validateEmptyField(theForm.txtSerialNo,"Serial No");
	reason += validateSpecialCharacters(theForm.txtSerialNo,"Serial No");
	
	reason += validateEmptyField(theForm.txtConnote,"Connote");
	reason += validateSpecialCharacters(theForm.txtConnote,"Connote");
	
	reason += validateEmptyField(theForm.cboRecipient,"Recipient");

  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);    	
		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Update';
		return true;
    }
}
</script>
<%
sub main
	dim intID
	intID = Trim(Request("id"))
	call getLoanTransfer(intID)	
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		if trim(request("Action")) = "Update" then
			dim traAccountCode, traModelNo, traSerialNo, traConnote, traRecipient, traCreatedBy
			
			traAccountCode 	= Server.HTMLEncode(Replace(Trim(Request("txtAccountCode")),"'","''"))
			traModelNo 		= Server.HTMLEncode(Replace(Trim(Request("txtModelNo")),"'","''"))
			traSerialNo 	= Server.HTMLEncode(Replace(Trim(Request("txtSerialNo")),"'","''"))
			traConnote 		= Server.HTMLEncode(Replace(Trim(Request("txtConnote")),"'","''"))
			traRecipient 	= Server.HTMLEncode(Replace(Trim(Request("cboRecipient")),"'","''"))
			traCreatedBy 	= Trim(session("logged_username"))
			
			'response.write ("SUBMIT!")
			'response.write (traAccountCode & "," & traModelNo & "," & traSerialNo & "," & traConnote & "," & traRecipient & "," & traCreatedBy)
			call updateLoanTransfer(traAccountCode, traModelNo, traSerialNo, traConnote, traRecipient, traCreatedBy)
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
          <a class="blog-nav-item active" href="loan-transfer.asp">Transfer</a>
        </nav>
	</div>
</div>
<div class="container"> 
  <br>
  <ol class="breadcrumb">
    <li><a href="loan-transfer.asp">Loan Transfer</a></li>
    <li class="active">Update Transfer</li>
  </ol>
  <h1 class="page-header"><i class="fa fa-pencil-square-o"></i> Update Loan Transfer ID = <%= Request("id") %></h1>
  <form action="" method="post" name="form_add_sales" id="form_add_sales" onsubmit="return validateFormOnSubmit(this)">
    <div class="form-group">
      <label for="cboProduct">Account code<font color="red">*</font>:</label>
      <input type="text" class="form-control" name="txtAccountCode" id="txtAccountCode" maxlength="6" size="6" value="<%= session("traAccountCode") %>" placeholder="Account" pattern=".{6,}" required title="6 characters minimum" />
    </div>
    <div class="form-group">
      <label for="cboProduct">Model no<font color="red">*</font>:</label>
      <input type="text" class="form-control" name="txtModelNo" id="txtModelNo" maxlength="20" size="20" value="<%= session("traModelNo") %>" placeholder="Model no" pattern=".{2,}" required title="2 characters minimum" />
    </div>
    <div class="form-group">
      <label for="txtSerialNo">Serial no<font color="red">*</font>:</label>
      <input type="text" class="form-control" name="txtSerialNo" id="txtSerialNo" maxlength="11" size="10" value="<%= session("traSerialNo") %>" placeholder="Serial no" pattern=".{4,}" required title="4 characters minimum" />
    </div>
    <div class="form-group">
      <label for="txtConnote">Con-note<font color="red">*</font>:</label>
      <input type="text" class="form-control" name="txtConnote" id="txtConnote" maxlength="30" size="30" value="<%= session("traConnote") %>" placeholder="Connote" pattern=".{4,}" required title="4 characters minimum" />
    </div>
    <div class="form-group">
      <label for="cboRecipient">Recipient<font color="red">*</font>:</label>
      <select name="cboRecipient" id="cboRecipient" class="form-control">
        <option value="damienh">Damien</option>
        <option value="georgen">George</option>
        <option value="russellw">Russell</option>
        <option value="shaunm">Shaun</option>
      </select>
    </div>
    <div class="form-group">
      <input type="hidden" name="Action" />
      <input type="submit" name="submit" id="submit" class="btn btn-default" value="Update" />
    </div>
  </form>
</div>
</body>
</html>