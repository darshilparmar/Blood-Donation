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
	conn.open "C:\Inetpub\wwwroot\Blood Donation\Database\Donor.mdb"
	set res=Server.CreateObject("ADODB.Recordset")
	res.open "Donor",conn,0,3,2
%>
<html>
<body bgcolor="#ffcccc">
<table border="3">
<tr>
	<th>Name</th>
	<th>Email</th>
	<th>MobileNo</th>
	<th>Gender</th>
	<th>BloodGroup</th>
	<th>RHFactor</th>
	<th>Country</th>
	<th>City</th>
	<th>Pin</th>
	<th>Bloodtime</th>
	<th>RegisteredAs</th>
</tr>
<%
	do while not res.EOF
	Response.write("<tr><td>"&res("Name")&"</td>")
	Response.write("<td>"&res("Email")&"</td>")
	Response.write("<td>"&res("MobileNo")&"</td>")
	Response.write("<td>"&res("Gender")&"</td>")
	Response.write("<td>"&res("BloodGroup")&"</td>")
	Response.write("<td>"&res("RHFactor")&"</td>")
	Response.write("<td>"&res("Country")&"</td>")
	Response.write("<td>"&res("City")&"</td>")
	Response.write("<td>"&res("Pin")&"</td>")
	Response.write("<td>"&res("Bloodtime")&"</td>")
	Response.write("<td>"&res("RegisteredAs")&"</td></tr>")
	res.movenext
	loop
%>
</table> 
</body>
</html>

