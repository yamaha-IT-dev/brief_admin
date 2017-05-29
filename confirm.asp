<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Confirmation</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
</head>
<body style="margin: 10px 10px 10px 10px">
<h2>The loan item has been successfully submitted. (<%= session("loan_account_code") %>)</h2>
<p><a href="loan_user.asp?account=<%= session("loan_account_code") %>">Go back to previous page.</a></p>
</body>
</html>