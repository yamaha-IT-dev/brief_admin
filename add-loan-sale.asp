<!--#include file="../include/connection_it.asp " -->
<!--#include file="class/clsLoanSale.asp " -->
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
<title>New Loan Sale</title>
<link rel="stylesheet" href="bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" href="css/header.css">
<link rel="stylesheet" href="//maxcdn.bootstrapcdn.com/font-awesome/4.3.0/css/font-awesome.min.css">
<link rel="stylesheet" href="//ajax.googleapis.com/ajax/libs/jqueryui/1.11.2/themes/smoothness/jquery-ui.css">
<script src="//code.jquery.com/jquery.js"></script>
<script src="//ajax.googleapis.com/ajax/libs/jqueryui/1.11.2/jquery-ui.min.js"></script>
<script src="bootstrap/js/bootstrap.js"></script>
<script src="../include/generic_form_validations.js"></script>
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
	
	reason += validateEmptyField(theForm.txtDealerCode,"Recipient");

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
			dim saleAccountCode, saleModelNo, saleSerialNo, saleQty, saleDealerCode, salePurchaseOrderNo, saleCreatedBy
			
			saleAccountCode = Server.HTMLEncode(Replace(Trim(Request("txtAccountCode")),"'","''"))
			saleModelNo 	= Server.HTMLEncode(Replace(Trim(Request("txtModelNo")),"'","''"))
			saleSerialNo 	= Server.HTMLEncode(Replace(Trim(Request("txtSerialNo")),"'","''"))
			saleQty 		= Server.HTMLEncode(Replace(Trim(Request("txtQty")),"'","''"))
			saleDealerCode 	= Server.HTMLEncode(Left(Trim(Request("txtDealerCode")),6))
			salePurchaseOrderNo = Server.HTMLEncode(Replace(Trim(Request("txtPurchaseOrderNo")),"'","''"))
			saleCreatedBy 	= Trim(session("logged_username"))
						
			'response.write (saleAccountCode & "," & traModelNo & "," & traSerialNo & "," & traQty & "," & traRecipient & "," & traCreatedBy & "," & session("newOrderNo") & "," & session("newOrderLine"))
			call addSale(saleAccountCode, saleModelNo, saleSerialNo, session("newOrderNo"), session("newOrderLine"), saleQty, saleDealerCode, salePurchaseOrderNo, saleCreatedBy)			
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
          <a class="blog-nav-item" href="loan-transfer.asp">Transfer</a>
          <a class="blog-nav-item active">Sale</a>
        </nav>
	</div>
</div>
<div class="container"> 
  <br>
  <ol class="breadcrumb">
    <li><a href="loan-sale.asp">Loan Sale</a></li>
    <li class="active">New Sale</li>
  </ol>
  <h1 class="page-header"><i class="fa fa-cart-plus"></i> New Loan Sale</h1>
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
      <label for="txtDealerCode">Dealer code<font color="red">*</font>:</label>
      <input type="text" class="form-control" name="txtDealerCode" id="txtDealerCode" maxlength="9"  placeholder="Dealer Code" pattern=".{4,}" required />
    </div>
    <div class="form-group">
      <label for="txtDealerCode">Sales Order no<font color="red">*</font>:</label>
      <input type="text" class="form-control" name="txtPurchaseOrderNo" id="txtPurchaseOrderNo" maxlength="12" placeholder="Sales Order No" pattern=".{4,}" required />
    </div>  
    <div class="form-group">
      <input type="hidden" name="Action" />
      <input type="submit" name="submit" id="submit" class="btn btn-default" value="Submit" />
    </div>
  </form>    
</div>
</body>
</html>