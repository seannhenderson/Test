<% const pagename="Event Registration" %>
<!--#include file="includes/hdr.asp"-->
<%
If Request.QueryString("view") <> "" Then
	Response.Write("<table>")
	Set connAtn=Server.CreateObject("ADODB.Connection")
	connReg.Provider="Microsoft.Jet.OLEDB.4.0"
	connReg.Open(Server.Mappath("db/test.mdb"))
	sql="SELECT FirstName,LastName,City,GuestCnt,EventDesc FROM tblEventReg ORDER BY EventDesc,City,LastName,FirstName"
	Set rs1=Server.CreateObject("ADODB.recordset")
	'ResultSetObject.Open SQL, ConnectionObject
	rs.Open sql, connReg
	Do Until rs.EOF
		Response.Write("<tr>")
			For each x in rs.Fields
			   Response.Write("<td>" & x.value & "</td>")
			Next
		Response.Write("<tr>")
    rs.MoveNext
	Loop
	rs.Close
	connReg.Close
	'set connReg = Nothing
	Response.Write("</table>")
Else
	dim regFlag
	dim sql
	reg = Request.Form("")
	regSubmt = Request.Form("submt")
	If regSubmt<>"" Then
		dim regFName
		dim regLName
		dim regEmail
		dim rsCnt
		dim ra
		regFName = Request.Form("FirstName")
		regLName = Request.Form("LastName")
		regEmail = Request.Form("Email")
		regPhone = Request.Form("Phone")
		regStreet1 = Request.Form("Street1")
		regStreet2 = Request.Form("Street2")
		regCity = Request.Form("City")
		regStateProv = Request.Form("StateProv")
		regPostalCode = Request.Form("PostalCode")
		regCountry = Request.Form("Country")
		regGuestCnt = Request.Form("GuestCnt")
		If regFName<>"" Then
			If regLName<>"" Then
				If regEmail<>"" Then
					'If regPword<>"" Then
						Set connReg=Server.CreateObject("ADODB.Connection")
						connReg.Provider="Microsoft.Jet.OLEDB.4.0"
						connReg.Open(Server.Mappath("db/test.mdb"))
						sql="SELECT count(*) as Cnt FROM tblEventReg WHERE Email like '" & regEmail & "' and StatusCode not in ('X','D')"
						Set rs1=Server.CreateObject("ADODB.recordset")
						'ResultSetObject.Open SQL, ConnectionObject
						rs1.Open sql, connReg
						rsCnt=rs1.Fields("Cnt")
						If rsCnt=0 Then
							err=0
							sql="INSERT INTO tblEventReg (FirstName, LastName, Email, Phone, Street1, Street2, City, StateProv, PostalCode, Country, GuestCnt) VALUES ('" & regFname & "', '" & regLname & "', '" & regEmail & "', '" & regPhone & "', '" & regStreet1 & "', '" & regStreet2 & "', '" & regCity & "', '" & regStateProv & "', '" & regPostalCode & "', '" & regCountry & "', " & regGuestCnt & ")"
							'connection.Method input, varToStoreRecordsAffectedCount, OptionalOptions
							on error resume next
							set ra=connReg.Execute(sql)
							If err<>0 Then
								Response.Write("Action cannot be performed! Error ID " & err & "<br><br>")
							Else
								Set regFlag=1
							End If
						Else
							Response.Write("<br>Account exists!<br>")
						End If
						regX=0
						rs1.close
						connReg.close
						'set connReg = Nothing
					'Else
					'	Response.Write("<h3>No Password!</h3>")
					'End If
				Else
					Response.Write("<h3>Email Required!</h3>")
				End If
			Else
				Response.Write("<h3>Last Name Required!</h3>")
			End If
		Else
			Response.Write("<h3>First Name Required!</h3>")
		End If
	End If
	If regFlag=1 Then 
%>
	<p>Registered!</p>
	<p>Click <a href="index.asp">here</a> to continue</p>
<% 
	Else 
%>
	<form id="regform" name="regform" METHOD="POST" ACTION="registration.asp" class="regformClass">
	<fieldset>
		<legend>Register for Event:</legend>
		<label>First:</label>
			<input type="text" name="FirstName" value="" class="regformFirstNameFldClass"><br>
		<label>Last:</label>
			<input type="text" name="LastName" value="" class="regformLastNameFldClass"><br>
		<label>Email:</label>
			<input type="text" name="Email" value="" class="regformEmailFldClass"><br>
		<label>Phone:</label>
			<input type="text" name="Phone" value="" class="regformPhoneFldClass"><br>
		<label>Street Line 1:</label>
			<input type="text" name="Street1" value="" class="regformStreet1FldClass"><br>
		<label>Street Line 2:</label>
			<input type="text" name="Street2" value="" class="regformStreet2FldClass"><br>
		<label>City:</label>
			<input type="text" name="City" value="" class="regformCityFldClass"><br>
		<label>State/Province:</label>
			<input type="text" name="StateProv" value="" class="regformStateProvFldClass"><br>
		<label>Postal Code:</label>
			<input type="text" name="PostalCode" value="" class="regformPostalCodeFldClass"><br>
		<label>Country:</label>
			<input type="text" name="Country" value="" class="regformCountryFldClass"><br>
		<label>Guests:</label>
			<input type="text" name="GuestCnt" value="1" class="regformGuestCntFldClass"><br>
		<!--<label>Password:</label>
			<input type="password" name="rPassword" value="" class="regformPwordFldClass"><br>
		-->
		<input type="submit" value="Register" class="regformSubmitBtnClass">
		<input type="hidden" name="submt" value="1">
	</fieldset>
	</form>
<% 
	End If 
End If
%>
<hr>
<p>Click <a href="registation.asp?view=1">here</a> to view attendees</p>
<br><br>
<!--#include file="includes/ftr.asp"-->