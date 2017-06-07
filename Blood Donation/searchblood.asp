<% @language="vbscript"%>
<% Option explicit
	dim conn,res
	set conn=Server.CreateObject("ADODB.Connection")
	conn.provider="Microsoft.Jet.OLEDB.4.0"
	conn.open "C:\Inetpub\wwwroot\Blood Donation\Database\Newuser.mdb"
	set res=Server.CreateObject("ADODB.Recordset")
	res.open "Newuser",conn,0,3,2
%>
<html>
<body>
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

