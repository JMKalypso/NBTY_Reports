<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Select what report to generate</title>
<link rel="stylesheet" href="pagestyle.css" media="screen, projection" />
<link rel="stylesheet" href="slimpicker.css" media="screen, projection" />
<script src="mootools-1.2.4-core-yc.js"></script>
<script src="mootools-1.2.4.4-more-yc.js"></script>
<script src="slimpicker.js"></script>

<script type="text/javascript">
	function checkforblank() {
		var errormessage = "";
		var reportType = document.getElementById('reports');
		if (reportType.options[reportType.selectedIndex].text == "") {
			errormessage += "Please, select a valid option for report type.\n";
			document.getElementById('reports').style.borderColor="red";	
		} else {
			document.getElementById('reports').style.borderColor="";
		}
		if (document.getElementById('email').value == "") {
			errormessage += "Please, enter an e-mail address.\n";
			document.getElementById('email').style.borderColor="red";	
		} else {
			document.getElementById('email').style.borderColor="";
		}
		if (document.getElementById('fromDate').value == "") {
			errormessage += "Please, enter starting date.\n";
			document.getElementById('fromDate').style.borderColor="red";	
		} else {
			document.getElementById('fromDate').style.borderColor="";
		}
		if (document.getElementById('toDate').value == "") {
			errormessage += "Please, enter end date.\n";
			document.getElementById('toDate').style.borderColor="red";	
		} else {
			document.getElementById('toDate').style.borderColor="";
		}
		if (errormessage != "") {
			alert(errormessage);
			return false;
		}
	}
</script>

</head>
<body>
	<form name="inputForm" action="URLPXServlet" method="POST" onsubmit="return checkforblank()">
		<table border="0">
			<tbody>
				<tr>
					<td>Select what report to generate: </td>
					<td>
						<select name="reports" id="reports">
							<option></option>
							<option>MBR Changes</option>
							<option>CU-Bulk Report</option>
							<option>CU Changes</option>
						</select>
					</td>
				</tr>
				<tr>
					<td>From Date: (mm/dd/yyyy)</td>
					<td>
						<input id="fromDate" name="fromDate" type="text" class="slimpicker" value="${fromDate}" />
					</td>
				</tr>
				<tr>
					<td>To Date: (mm/dd/yyyy)</td>
					<td>
						<input id="toDate" name="toDate" type="text" class="slimpicker" value="${toDate}" />
					</td>
				</tr>
				<tr>
					<td>Send to e-mail:</td>
					<td>
						<input id="email" name="email" type="text" value="${email}" size="35"/>
						${errorMessage}
					</td>
				</tr>
			</tbody>
		</table>
		<input type="reset" value="Clear" name="clear" />
		<input type="submit" value="Submit" name="submit" />
	</form>
	<script>
	
	$$('input.slimpicker').each( function(el){
		var picker = new SlimPicker(el);
	});
	
	</script>
</body>
</html>