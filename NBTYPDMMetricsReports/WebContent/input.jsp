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
</head>
<body>
	<form name="inputForm" action="URLPXServlet" method="POST">
		<table border="0">
			<tbody>
				<tr>
					<td>Select what report to generate: </td>
					<td>
						<select name="reports">
							<option>MBR Changes</option>
							<option>CU-Bulk Report</option>
							<option>CU Changes</option>
						</select>
					</td>
				</tr>
				<tr>
					<td>From Date: (mm/dd/yyyy)</td>
					<td>
						<input id="fromDate" name="fromDate" type="text" class="slimpicker" autocomplete="off" value="" />
					</td>
				</tr>
				<tr>
					<td>To Date: (mm/dd/yyyy)</td>
					<td>
						<input id="toDate" name="toDate" type="text" class="slimpicker" autocomplete="off" value="" />
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