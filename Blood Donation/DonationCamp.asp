<% @language="vbscript"%>
<% Option explicit
	dim conn,res
	set conn=Server.CreateObject("ADODB.Connection")
	conn.provider="Microsoft.Jet.OLEDB.4.0"
	conn.open "C:\Inetpub\wwwroot\Blood Donation\Database\DonationCamp.mdb"
	set res=Server.CreateObject("ADODB.Recordset")
	res.open "DonationCamp",conn,0,3,2
	
	res.addnew
	res("Campfor")=Request.Form("campfor")
	res("Organisation")=Request.Form("org")
	res("Address")=Request.Form("address")
	res("Country")=Request.Form("country")
	res("State")=Request.Form("state")
	res("City")=Request.Form("city")
	res("Pin")=Request.Form("pin")
	res("CampVenue")=Request.Form("campvenue")
	res("CampTime")=Request.Form("camptime")
	res("Date")=Request.Form("date")
	res("Month")=Request.Form("month")
	res("Year")=Request.Form("year")
	res("PhoneNo")=Request.Form("phoneno")
	res("Email")=Request.Form("email")
	res.update
	res.movenext
	response.write("Signup Success")
	response.redirect("home.html")
%>


