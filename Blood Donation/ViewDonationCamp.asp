  <% @language="vbscript"%>
  <html>
    <head>
      <!--Import Google Icon Font-->
      <link href="http://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
      <!--Import materialize.css-->
      <link type="text/css" rel="stylesheet" href="css/materialize.min.css"  media="screen,projection"/>

      <!--Let browser know website is optimized for mobile-->
      <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
	
    </head>

    <body bgcolor="" border="5">
      <!--Import jQuery before materialize.js-->
      <script type="text/javascript" src="https://code.jquery.com/jquery-2.1.1.min.js"></script>
      <script type="text/javascript" src="js/materialize.min.js"></script>
	 <!-- Navbar goes here --> 
	&nbsp&nbsp&nbsp<img src="photos/blood.jpg" height="150"   align="middle" border="5"></img>
  <nav>
    <div class="nav-wrapper">
      <a href="#" class="brand-logo"><font face="Cooper Std Black" color="black">&nbsp&nbspRedDonor</font></img></a>
      <ul id="nav-mobile" class="right hide-on-med-and-down">
        <li><a href="home.html">Home</a></li>
        <li><a href="">Contact Us</a></li>
        <li><a href="collapsible.html">About Us</a></li>
      </ul>
    </div>
  </nav>
 </body>
 </html>
<% 
	dim conn,res
	set conn=Server.CreateObject("ADODB.Connection")
	conn.provider="Microsoft.Jet.OLEDB.4.0"
	conn.open "C:\Inetpub\wwwroot\Blood Donation\Database\DonationCamp.mdb"
	set res=Server.CreateObject("ADODB.Recordset")
	res.open "DonationCamp",conn,0,3,2
%>
<html>
<body bgcolor="#ffcccc">
<table border="3">
<tr>
	<th>Campfor</th>
	<th>Organisation</th>
	<th>CampVenue</th>
	<th>CampTime</th>
	<th>Date</th>
	<th>Month</th>
	<th>Year</th>
	<th>PhoneNo</th>
	<th>Email</th>
</tr>
<%
	do while not res.EOF
	Response.write("<tr><td>"&res("Campfor")&"</td>")
	Response.write("<td>"&res("Organisation")&"</td>")
	Response.write("<td>"&res("CampVenue")&"</td>")
	Response.write("<td>"&res("CampTime")&"</td>")
	Response.write("<td>"&res("Date")&"</td>")
	Response.write("<td>"&res("Month")&"</td>")
	Response.write("<td>"&res("Year")&"</td>")
	Response.write("<td>"&res("PhoneNo")&"</td>")
	Response.write("<td>"&res("Email")&"</td></tr>")
	res.movenext
	loop
%>
</table> 
</body>
</html>