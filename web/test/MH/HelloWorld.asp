<%@ Language=VBScript %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<%
cars = Request.Form("cars")
%>
<head>


<%
sub vbproc(num1,num2)
response.write(num1*num2)
end sub
%>
<%
Response.Buffer =true
%>
<%
Response.ExpiresAbsolute=#July 15, 2008 11:15:00#
%>
<script language="javascript" runat="server" type="text/javascript">

function jsproc(num1,num2)
{
response.write(num1*num2)
}

function message()
{
alert("This alert box was called with the onload event");
}

</script>


<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Hello World Test</title>
</head>

<body onload="message()">
<%
response.write("Hello World!")
%>

<%
response.write("<h2>You can use HTML tags to format the text!</h2>")
%>

<%
response.write("<p style='color:#0000ff'>This text is styled with the style attribute!</p>")
%>

<%
dim name
name = "Donald Duck"
response.write("My name is: " & name & "<br />")

%>

<br />
<script language="javascript">

var myarray = new Array();

function array(ProductName)
{	
	myarray[0] = ProductName;	
	
	document.getElementById('CurrentOrders').innerHTML=myarray[0];
}

</script>


<br />
<div id="CurrentOrders">
	
</div>
<br />

<input type="text" id="ProductName" onclick="this.value='';"/>
<br />
<br />
<input type="button" id="btnAddOrder" value="Add Orders" onclick="array(document.getElementById('ProductName').value"/>
<br />


<%
Dim famname(5),i
famname(0) = "Jan Egil"
famname(1) = "Tove"
famname(2) = "Hege"
famname(3) = "Stale"
famname(4) = "Kai Jim"
famname(5) = "Borge"

for i = 0 to 5
	response.write(famname(i) & "<br />")
next
%>

<%
for i=1 to 6
   response.write("<h" & i & ">Header " & i & "</h" & i & ">")
next
%>

<%
dim h
h =hour(now())

response.write("<p>" & now())
response.write("</p>")
if h<12 then
	response.write("Good Morning!")
else
	response.write("Good Day!")
end if
%>

<script language="javascript">
var d = new Date()
var h = d.getHours()

response.write("<p>")
response.write(d)
response.write("</p>")

if h < 12
	response.write("Good Morning!")
else
	response.write("Good Day!")
</script>
<br />
Today's date is: <%response.write(date())%>.
<br />
The server's local time is: <%response.write(time())%>.

<p>
VBScripts function <b>WeekdayName</b> is used to get a weekday:
</p>
<%
response.write(WeekdayName(1))
response.write("<br />")
response.write(WeekdayName(2))
%>

<p>Abbreviated name of a weekday: <p>
<%
response.write(WeekdayName(1,true))
response.write("<br />")
response.write(WeekdayName(2, true))
%>

<p>Current weekday:</p>
<%
response.write(WeekdayName(weekday(date)))
response.write("<br />")
response.write(WeekdayName(weekday(date), true))
%>

<p>VBScripts function <b>MonthName</b> is used to get a month: </p>
<%
response.write(MonthName(1))
response.write("<br />")
response.write(MonthName(2))
%>
<p>Abbreviated name of a month:</p>
<%
response.write(MonthName(1,true))
response.write("<br />")
response.write(MonthName(2, true))
%>

<p>Current Month:</p>
<%
response.write(MonthName(month(date)))
response.write("<br />")
response.write(MonthName(month(date),true))
%>

Today it is
<% response.write(weekdayname(weekday(date)))%>,
<br />
and the month is
<% response.write(MonthName(month(date)))%>

<p>Countdown to year 3000:</p>

<p>
<%millennium=cdate("1/1/3000 00:00:00")%>
It is
<%response.write(datediff("yyyy",Now(),millennium))%>
years to year 3000!
<br />
It is
<%response.write(Datediff("m",now(),millennium))%>
months to year 3000!
<br />
It is
<%response.write(datediff("ww",now(),millennium))%>
weeks to year 3000!
<br />
It is
<%response.write(datediff("d", now(), millennium))%>
days to year 3000!
<br />
It is
<%response.write(datediff("h", now(), millennium))%>
hours to year 3000!
<br />
It is
<%response.write(datediff("n", now(), millennium))%>
minutes to year 3000!
<br />
It is
<%response.write(datediff("s", now(), millennium))%>
seconds to year 3000!
</p>
<br />

<%
response.write(dateadd("d",30,date()))
%>

<p>
Syntax for DateAdd: DateAdd(interval,number,date). You can use <b>DateAdd</b> for example to calculate a date 30 days from today.
</p>
<br />

<%
response.write(formatdatetime(date(),vbgeneraldate))
response.write("<br />")
response.write(FormatDateTime(date(),vblongdate))
response.write("<br />")
response.write(FormatDateTime(date(),vbshortdate))
response.write("<br />")
response.write(FormatDateTime(now(),vblongtime))
response.write("<br />")
response.write(FormatDateTime(now(),vbshorttime))
%>
<p>
Syntax for FormatDateTime: FormatDateTime(date,namedformat).
</p>

<br />

<%
somedate = "10/30/99"
response.write("Is the following a date, 10/30/99")
response.write("<br />")
response.write(IsDate(somedate))
%>

<br />

<%
name = "Bill Gates"
response.write(ucase(name))
response.write("<br />")
response.write(lcase(name))
%>
<br />

<%

name = " W3Schools "
response.write("visit" & name & "now <br />")
response.write("visit" & trim(name) & "now<br />")
response.write("visit" & ltrim(name) & "now <br />")
response.write("visit" & rtrim(name) & "now")
%>
<br />

<%
sometext = "Hello Everyone!"
response.write(strReverse(sometext))
%>

<br />

<%
i = 48.66776677
j = 48.3333333
response.write(Round(i))
response.write("<br />")
response.write(Round(j))
%>

<br />

<%
randomize()
response.write(rnd())
%>

<br />

<%
sometext = "Welcome to this web"
response.write(Left(sometext, 5))
response.write("<br />")
response.write(Right(sometext,5))
%>

<br />

<%
sometext = "Welcome to this web!!"
response.write(Replace(sometext, "Web", "Page"))
%>

<br />

<%
sometext = "Welcome to this Web!!"
response.write(Mid(sometext, 9,2))
%>
<br />

<p>
You can call a procedure like this:
</p>
<p>
Result1: <%call vbproc(3,4)%>
</p>
<p>
Or, like this:
</p>
<p>
Result2: <% call vbproc(3,4)%>
</p>
<p>
Result3: <% call jsproc(3,4) %>
</p>

<form action="HelloWorld.asp" method="get">
Your name: <input type="text" name="fname" size="20" />
<input type="submit" value="Submit" />
</form>
<%
dim fname
fname=Request.QueryString("fname")
if fname <> "" then
	response.write("Hello " & fname & "! <br />")
	response.write("How are you today?")
end if
%>

<form action="HelloWorld.asp" method="post">
Your name: <input type="text" name="fname" size="20" />
<Input type="submit" value="Submit" />
</form>
<%
fname = Request.form("fname")
if fname <> "" then
	response.write("Hello " & fname & "!<br />")
	response.write("How are you today?")
end if
%>

<br />

<form action="HelloWorld.asp" method="post">
<p>Please select your favourite car:</p>

<input type="radio" name="cars"
<%if cars="Volvo" then Response.write("Checked")%>
value = "Volvo" />Volvo</input>
<br />
<input type="radio" name="cars"
<%if cars="Saab" then Response.Write("checked")%>
value="Saab">Saab</input>
<br />
<input type="radio" name="cars"
<%if cars="BMW" then Response.Write("checked")%>
value="BMW">BMW</input>
<br /><br />
<input type="submit" value="Submit" />
</form>
<%
if cars <> "" then
	response.write("<p>Your favourite car is: " & cars & "</p>")
end if
%>

<br />

<%
dim numvisits
response.cookies("NumVisits").Expires=date+365
numvisits=request.cookies("NumVisits")

if numvisits="" then
	response.cookies("NumVisits")=1
	response.write("Welcome! This is the first time you are visiting this web page.")
else
	response.cookies("NumVisits")=numvisits+1
	response.write("you have visited this ")
	response.write("Web page " & numvisits)
	if numvisits = 1 then
		response.write " time before!"
	else
		response.write " times before!"
	end if
end if
%>

<br />

<%
if request.form("select") <> "" then
	response.redirect(request.form("select"))
end if
%>

<form action="HelloWorld.asp" method="post">

<input type="radio" name="select" 
value="www.google.co.uk">
Server Example<br />

<input type="radio" name="select" 
value="www.msn.co.uk">
Text Example<br /><br />
<input type="submit" value="Go!">
</form>

<br />

<%
randomize()
r=rnd()
if r>0.5 then
	response.write("<a href='http://www.w3schols.com'>W3Schools.com!</a>")
else
	response.write("<a href='http://www.refsnesdata.no'>Refsnesdata.no!</a>")
end if
%>

<p>
This example demonstrates a link, each time you load the page, it will display 
one of two links: W3Schools.com! OR Refsnesdata.no! There is a 50% chance for 
each of them.
</p>

<br />

<p> 
This text will be sent to your browser when my response buffer is flushed. 
</p>
<% 'Controlling the buffer
Response.flush
%>
<br />

<p>This is some text I want to send to the user.</p>
<p>No, I changed my mind. I want to clear the text.</p>
<% 'Clearing the buffer
Response.Clear
%>

<br />

<p>I am writing some text. This text will never be<br />
<% ' End the script in the middle of processing
'Response.End
%>
finished! It's too late to write more!</p>
<br />

<p>This page will be refreshed with each access!</p>

<br />

<p>This page will expire on July 15, 2008, 11:15:00!</p>

<br />

<%
If response.IsClientConnected=true then
	Response.Write("The user is still connected!")
else
	Response.Write("The user is not connected!")
end if
%>

<br />

<a href="HelloWorld.asp?color=green">Example</a>
<%
Response.Write(Request.QueryString)
%>

<form action="HelloWorld.asp" method="get">
First name: <input type="text" name="fname" /><br />
last name: <input type="text" name="lname" /><br />
<input type="submit" value="Submit" />
</form>
<%
response.write(Request.QueryString)
%>
<br />
<p>
<b>You are browsing this site with:</b>
<%Response.Write(Request.ServerVariables("http_user_agent"))%>
</p>
<p>
<b>Your IP address is:</b>
<%Response.Write(Request.ServerVariables("remote_addr"))%>
</p>
<p>
<b>The DNS lookup of the IP address is:</b>
<%Response.Write(Request.ServerVariables("remote_host"))%>
</p>
<p>
<b>The method used to call the page:</b>
<%Response.Write(Request.ServerVariables("request_method"))%>
</p>
<p>
<b>The server's domain name:</b>
<%Response.Write(Request.ServerVariables("server_name"))%>
</p>
<p>
<b>The server's port:</b>
<%Response.Write(Request.ServerVariables("server_port"))%>
</p>
<p>
<b>The server's software:</b>
<%Response.Write(Request.ServerVariables("server_software"))%>
</p>
<br />

<p>
All possible server variables:
</p>
<%
For Each Item in Request.ServerVariables
	Response.Write(Item & "<br />")
next
%>

<br />

<form action="HelloWorld.asp" method="post">
Please type something:
<input type="text" name="txt"><br /><br />
<input type="submit" value="Submit">
</form>

<%
If Request.Form("txt")<>"" Then
   Response.Write("You submitted: ")
   Response.Write(Request.Form)
   Response.Write("<br /><br />")
   Response.Write("Total bytes: ")
   Response.Write(Request.Totalbytes)
End If
%>
<br />

<%
Response.write(Session.SessionID)
%>
<br />

<%
Response.write("<p>")
response.write("Default Timeout is: " & Session.Timeout & " minutes.")
response.write("</p>")

Session.timeout=30

response.write("<p>")
response.write("Timeout is now: " & Session.Timeout & " minutes.")
response.write("</p>")
%>

<br />
<%
Set fs = server.createobject("Scripting.FileSystemObject")
Set rs = fs.getFile(Server.MapPath("HelloWorld.asp"))
modified = rs.DateLastModified
%>
This file was last modified on: <%response.write(modified)
set rs = nothing
set fs = nothing
%>
<br />

<p>This is the text in the text file:</p>
<%
Set fs=Server.CreateObject("Scripting.FileSystemObject")

Set f=fs.OpenTextFile(Server.MapPath("testread.txt"), 1)
Response.Write(f.ReadAll)
f.Close

Set f=Nothing
Set fs=Nothing
%>
<link rel="stylesheet" href="http://system.newzapp.co.uk/css/import.css" type="text/css">
<link rel="stylesheet" href="http://system.newzapp.co.uk/css/theme.css" type="text/css">

<div id="divContent" style="border:0;">
  <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td width="50%" valign="top" style="border-bottom:1px solid #7f7f7f; border-right:1px solid #7f7f7f; padding:20px; padding-bottom:5px;"><h2>Welcome to Campaign Reports</h2>
        <br>
        From here you can view and monitor your Campaigns. To view a Campaign click on the Campaigns folder in the navigation panel, this will display all of the Email Campaigns you have sent. <br>
        <br>
        <ul>
          <li>To start managing your reports click on the Campaign in the Navigation panel.</li>
        </ul>
		<br />
<div id="lastUpdated"></div></td>
      <td valign="top" style="border-bottom:1px solid #7f7f7f; border-right:0px; border-left:1px solid #ffffff;padding:20px;"><h2>Send Summary</h2>
        <br>
        The table below contains details of the total emails sent from your account per month.. <br>
        <br>
        <div id="subscriberStats"></div>
        <div id="subscriberStatsUnsubscribed"></div>
        
       <div id="divStats" style="width:auto; overflow:auto; height:150px;">
	   <table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr bgcolor="#A2B0BF">
				<th>Month</th>
				<th>Number Sent</th>
			</tr>
			<tr>
				<td><strong>July 2008 </strong></td>
				<td align="right" id="nMonth1">0</td>
			</tr>
			<tr>
				<td><strong>June 2007 </strong></td>
				<td align="right" id="nMonth2">60</td>
			</tr>
		</table>
	   </div></td>
    </tr>
    <tr>
      <td colspan="2" valign="top" style="border-bottom: 0; border-right:0px; border-top:1px solid #ffffff; padding:20px; padding-bottom:0px;"><h2>Campaign History & Status</h2>
        <br>
        <div id="div_Imports"> Loading Campaign Data <img src="/images/16time.gif" align="absmiddle"> </div>
        <div id="returned" style="font-family:Arial, Helvetica, sans-serif; font-size:11px; padding:0px; display:none;">
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <th>Campaign</th>
              <th>Send Date</th>
              <th>Status</th>
              <th>Selected</th>
              <th>Sent</th>
              <th>Opened</th>
             <th>Clicked</th>
            </tr>
            <tr><td><font color="#990000"><strong>NO NAME SPECIFIED</strong></font></td><td>15/07/2008 13:55:44</td><td align="center"><span id="status_99826"><img src="/interfaceimages/campaign_sendingcomplete.gif" align="absmiddle"></span></td><td align="right" id="selected_99826">-</td><td align="right" id="sent_99826">-</td><td align="right" id="opened_99826">-</td><td align="right" id="clicked_99826">-</td></tr><tr><td><font color="#990000"><strong>NO NAME SPECIFIED</strong></font></td><td>15/07/2008 13:55:29</td><td align="center"><span id="status_99825"><img src="/interfaceimages/campaign_sendingcomplete.gif" align="absmiddle"></span></td><td align="right" id="selected_99825">-</td><td align="right" id="sent_99825">-</td><td align="right" id="opened_99825">-</td><td align="right" id="clicked_99825">-</td></tr><tr><td><font color="#990000"><strong>NO NAME SPECIFIED</strong></font></td><td>15/07/2008 13:55:13</td><td align="center"><span id="status_99824"><img src="/interfaceimages/campaign_sendingcomplete.gif" align="absmiddle"></span></td><td align="right" id="selected_99824">-</td><td align="right" id="sent_99824">-</td><td align="right" id="opened_99824">-</td><td align="right" id="clicked_99824">-</td></tr><tr><td><font color="#990000"><strong>NO NAME SPECIFIED</strong></font></td><td>15/07/2008 13:54:51</td><td align="center"><span id="status_99823"><img src="/interfaceimages/campaign_sendingcomplete.gif" align="absmiddle"></span></td><td align="right" id="selected_99823">-</td><td align="right" id="sent_99823">-</td><td align="right" id="opened_99823">-</td><td align="right" id="clicked_99823">-</td></tr>
          </table>
        </div>
		<div id="updateDate"></div>
		</td>
    </tr>
  </table>
  </div>
</div>
<br />

<div class="reportstoolbar">
<table  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td nowrap>Current Selection: </td>
    <td><LABEL FOR="currentView1">
      <input name="currentView" type="radio" id="currentView1" onClick="setCurrentView(this.value);" value="0" checked="checked" > 
	  
	Campaign Overview </LABEL>
	<%
	If bAdvancedReporting_v002d = True Then
	%>
	<LABEL FOR="currentView2"><input id="currentView2" name="currentView" type="radio" value="1" onClick="setCurrentView(this.value);" >
	Compare Campaigns</LABEL>
	<%
	Else
	%>
	<LABEL FOR="currentView2" disabled=true><input id="currentView2" name="currentView" type="radio" value="1" onClick="setCurrentView(this.value);" >
	Compare Campaigns</LABEL>
	<%
	End If
	%>
	
</td>

    <!-- <td><img src="/interfaceimages/breaker1.gif" width="0" height="22" align="absMiddle" /></td>-->
    <td><div class="btn" onMouseDown="NewClass(this, 'btnDown');" onMouseUp="NewClass(this, 'btn');" onMouseOver="NewClass(this, 'btnOver');" onMouseOut="NewClass(this, 'btn');"><a href="javascript:exportCampaignData();"><img src="/interfaceimages/exportreport.gif" width="133" height="22" title="Export Campaign Data"></a></div></td>
	<!-- 
	<%If bCustomFields and 1=2 Then%>
	<td><img src="/interfaceimages/breaker1.gif" width="9" height="22" align="absMiddle" /></td>
	<td width="22"><a href="javascript: getJSHTML();">showJSHTML</a></td>
	<%
	End If
	%>
	<% If bBounceMgmt Then %>
		<td><img src="/interfaceimages/breaker1.gif" width="9" height="22" align="absMiddle" /></td>
	<td width="22"><a href="javascript:unsubscribeBounces('Hard');"><img src="/interfaceimages/b_unsubscribeHardBounces.gif" width="150" height="20" title="Export Campaign Data"></a></td>
	<td><img src="/interfaceimages/breaker1.gif" width="9" height="22" align="absMiddle" /></td>
	<td width="22"><a href="javascript:unsubscribeBounces('Soft');"><img src="/interfaceimages/b_unsubscribeSoftBounces.gif" width="150" height="20" title="Export Campaign Data"></a></td>
	<% End If %>
	-->
  </tr>
</table>
</div>

<br />

<%
'when the text file was created 
dim fs, f
set fs=server.CreateObject("Scripting.FileSystemObject")
set f=fs.getfile(Server.MapPath("testread.txt"))
response.Write("The file testread.txt was created on: " & f.datecreated)
set f =nothing
set fs=nothing
%>

<br />

<%
'when the text file was last modified
dim fs2, f2
set fs2=server.CreateObject("Scripting.filesystemobject")
set f2=fs2.getfile(server.MapPath("testread.txt"))
response.write("The file testread.txt was last modified on: " & f2.datelastmodified)
set f2=nothing
set fs2=nothing
%>

<br />

<%
dim fs3,f3
set fs3=server.CreateObject("Scripting.filesystemobject")
set f3=fs3.getfile(server.MapPath("testread.txt"))
response.Write("The file testread.txt was last accessed on: " & f3.datelastaccessed)
set f3=nothing
set fs3=nothing
%>

<br />

<%
dim fs4, f4
set fs4=server.CreateObject("Scripting.filesystemobject")
set f4=fs4.getfile(server.MapPath("testread.txt"))
response.Write("The attributes of the file testread.txt are: " & f4.attributes)
set f4=nothing
set fs4=nothing
%>

<br />

<%
dim d,a,s
set d=server.CreateObject("Scripting.Dictionary")
d.add "n", "Norway"
d.add "i", "Italy"
response.Write("<p>The values of the items are:</p>")
a=d.items
for i = 0 to d.count -1
	s = s & a(i) & "<br />"
next
response.Write(s)
set d=nothing
%>

<br />

<%
dim d2,a2,s2
set d2=server.CreateObject("Scripting.Dictionary")
d2.add "n", "Norway"
d2.add "i", "Italy"
response.Write("<p>The values of the keys are:</p>")
a2=d2.keys
for i = 0 to d2.count -1
	s2 = s2 & a2(i) & "<br />"
next
response.Write(s2)
set d2=nothing
%>

<br />

<%
dim d3
set d3=server.CreateObject("Scripting.Dictionary")
d3.Add "n", "Norway"
d3.Add "i", "Italy"
response.Write("The value of item n is: " & d3.item("n"))
set d3=nothing
%>

<br />

<%
dim d4
set d4=Server.CreateObject("Scripting.Dictionary")
d4.Add "n", "Norway"
d4.Add "i", "Italy"
d4.Key("i") = "it"
Response.Write("The key i has been set to it, and the value is: " & d4.Item("it"))
set d4=nothing
%>

<br />

<%
dim d5,a5,s5
set d5=server.CreateObject("Scripting.Dictionary")
d5.add "n", "Norway"
d5.add "i", "Italy"
response.Write("The number of key/item pairs is: " & d5.count)
set d5=nothing
%>
<br />
<script type="text/javascript">
document.write("Hello World!");
document.write("<h1>This is a header</h1>");
var x=5+5;
document.write(x);

function disp_confirm()
{
var r=confirm("Press a button");
	if(r==true)
	{
		document.write("You pressed OK!");
	}
	else
	{
		document.write("You pressed cancel!");
	}
}

var i=0;
for(i=0;i<=10;i++)
{
document.write("The number i is " + i);
document.write("<br />");
}
document.write("<br />");
var j=0;
while(j<=10)
{
document.write("The number j is " + j);
document.write("<br />");
j++;
}
document.write("<br />");
var x;
var mycars = new Array();
mycars[0] = "Saab";
mycars[1] = "Volvo";
mycars[2] = "BMW";

for (x in mycars)
{
document.write(mycars[x] + "<br />");
}

</script>

<br />
<script type="text/javascript">

personObj = new Object();
personObj.firstname="John";
personObj.lastname="Doe";
personObj.age=50;
personObj.eyecolor="blue";
//document.write(personObj.firstname);
personObj.eat="eat";

function person(firstname, lastname, age, eyecolor)
{
this.firstname=firstname;
this.lastname=lastname;
this.age=age;
this.eyecolor=eyecolor;

this.newlastname=newlastname;
}

myfather=new person("john", "Doe",50,"blue");
myMother=new person("Sally","Rally",48,"green");
myMother.newlastname("Doe");
document.write(myMother.lastname);

function newlastname(new_lastname)
{
this.lastname=new_lastname;
}
</script>
<br />

<script type="text/javascript">

TeamObj = new Object();
TeamObj.Club="Man Utd";
TeamObj.Division="Prem";
TeamObj.pts=50;
TeamObj.Position=5;

function TeamStats(Club, Division,pts, Position)
{
	this.Club = Club;
	this.Division = Division;
	this.pts = pts;
	this.Position = Position;
	this.newPosition=newPosition;
}
team1 = new TeamStats("Man Utd","Prem",100,5);

document.write(team1.Club+ ", " +team1.Division+ ", "+team1.pts+ ", "+team1.Position);
function newPosition(new_Position)
{
	this.Position=new_Position;
}

</script>
<br />

<script type="text/javascript">
TodayObj = new Object();
TodayObj.LastDate = new Date();
document.write(TodayObj.LastDate);

</script>

<br />

<script type="text/javascript">
nameObj = new Object();
nameObj.name = "Mike";

function myTest(name)
{
this.name = name;
}
document.write(nameObj.name);

</script>

<br />

<script type="text/javascript">
var mything = new myobject();
function myobject()
{
this.containedValue = 0;
this.othercontainedValue = 0;
this.anothercontainedValue = 0;
}
myobject.prototype.newContainedValue = 5;
</script>

<%

Dim objCC
set objCC = new CDBInterface
objCC.Init arrCustDetails

objCC.NewSQL
objCC.AddFieldToQuery ORDERDETAILS_PRODUCTID
objCC.AddFieldToQuery ORDERDETAILS_QUANTITY
objCC.AddWhere Field1, "=", 1,"number", "and"

objCC.QueryDatabase


%>

</body>
</html>







