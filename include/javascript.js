<!--
/*******************************************************************************************
PRINT FRIENDLY
*******************************************************************************************/

function PrintThisPage() 
{ 
   var sOption="toolbar=yes,location=no,directories=yes,menubar=yes,"; 
       sOption+="scrollbars=yes,width=900,height=900,left=100,top=25"; 

   var sWinHTML = document.getElementById('contentstart').innerHTML; 
   
   var winprint=window.open("","",sOption); 
       winprint.document.open(); 
       winprint.document.write('<html><title>LOGISTICS Shipment Report</title><LINK href=include/stylesheet.css rel=Stylesheet><body>'); 
       winprint.document.write(sWinHTML);          
       winprint.document.write('</body></html>'); 
       winprint.document.close(); 
       winprint.focus(); 
}

//-->