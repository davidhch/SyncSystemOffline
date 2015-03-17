<!--#include virtual="\_work\bugs\_private\config.asp"-->
<%If bSuperAdminAccess Then%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<link rel="stylesheet" href="theme_w3c.css" type="text/css">
<style>
imput, select{
font-size:10px;
}
</style>
</head>

<body>
<form id="form1" name="form1" method="post" action="test.asp">
  <table border="0" cellspacing="0" cellpadding="5">
    <tr>
      <td>Section:</td>
      <td><select name="form_Section">
	<option value="">SELECT AN OPTION</option>
	<option value="admin">admin</option>
	<option value="dashboard">dashboard</option>
	<option value="development">development</option>
	<option value="ecom">ecom</option>
	<option value="editing">editing</option>
	<option value="file library">file library</option>
	<option value="ftaf">ftaf</option>
	<option value="help system">help system</option>
	<option value="import">import</option>
	<option value="order entry">order entry</option>
	<option value="report">report</option>
	<option value="sending">sending</option>
	<option value="subscriber">subscriber</option>
	 <option value="superadmin">superadmin</option>	
	<option value="all">ALL</option>
      </select></td>
    </tr>
    <tr>
      <td>NewZapp:</td>
      <td><select name="form_NewZapp">
          <option value="0">SELECT AN OPTION</option>
          <option value="1" selected>Yes</option>
          <option value="0">No</option>
      </select></td>
    </tr>
    <tr>
      <td>EditSite:</td>
      <td><select name="form_EditSite">
          <option value="0">SELECT AN OPTION</option>
          <option value="1">Yes</option>
          <option value="0" selected>No</option>
      </select></td>
    </tr>
    <tr>
      <td>Date:</td>
      <td><input name="form_ReportedDate" type="text" value="<%response.write now()%>"/></td>
    </tr>
    <tr>
      <td>Reported By:</td>
      <td><select name="form_ReportedBy">
          <option value="">SELECT AN OPTION</option>
          <option value="aw">Annette West</option>
		  <option value="dmh">Darren Hepburn</option>
		  <option value="dc">Darren Curnow</option>
		  <option value="dhh">David Hazzard</option>
		  <option value="gm">Gemma Mitchell</option>
		  <option value="jm">Justin Mortimore</option>
		  <option value="mm">Mandy Mitchell</option>
		  <option value="mn">Mike Nicol</option>
		  <option value="rb">Rachael Buchan</option>
		  <option value="rc">Richard Curnow</option>
		  <option value="tm">Trevor Munday</option>
		  <option value="cust">Customer</option>
		  <option value="oth">other</option>
      </select></td>
    </tr>
    <tr>
      <td>Type:</td>
      <td><select name="form_Type">
          
          <option value="1">Task</option>
          <option value="2">Wish List</option>
          <option value="3" selected="selected">Bug</option>
		  <option value="4">Support</option>
      </select></td>
    </tr>
    <tr>
      <td>Notes:</td>
      <td><textarea name="form_Notes" cols="100" rows="10" /></textarea></td>
    </tr>
	 <tr>
      <td>PRIORITY:</td>
      <td><input name="form_Priority" type="text" value=""/> 0-5 , where 0 is most urgent.</td>
    </tr>
	 <tr>
	   <td>CID</td>
	   <td><input name="form_CID" type="text" value=""/></td>
    </tr>
  </table>
  <input name="" type="submit" />
</form>
</body>
</html>
<%Else%>
	ERROR
<%End If%>