<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
</head>

<body>
<h1>default.asp</h1>
default<br />
<br />

YOU ENTERED <% = Request.Form("Name") %><br />
<br />
<h1>PICK YOUR DESTINATION</h1>

         <form action="page2.asp" method="POST">
         <input type="text" name="Name">
         <input type="submit" value="Submit">
         </form>
</body>
</html>
