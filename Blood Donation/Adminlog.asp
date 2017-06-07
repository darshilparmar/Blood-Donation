<% @language="vbscript"%>
<% Option explicit
	dim conn,res
	set conn=Server.CreateObject("ADODB.Connection")
	conn.provider="Microsoft.Jet.OLEDB.4.0"
	conn.open "C:\Inetpub\wwwroot\Blood Donation\Database\Login.mdb"
	set res=Server.CreateObject("ADODB.Recordset")
	res.open "Login",conn,0,3,2
	dim Name,Password
	Name=Request.Form("username")
	Password=Request.Form("password")
	Do While not res.EOF
	If(Name=res("Name") AND Password=res("Password")) then
	Response.Redirect("Home.html")
	Else
	Response.write("Invalid username or password")
	response.redirect("login.html")
	End If
	res.MoveNext
	Loop
%>
