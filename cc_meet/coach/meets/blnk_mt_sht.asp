<%@ Language=VBScript %>
<%
Option Explicit

Dim i, j
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>GSE� Blank CC Meet Sheet</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="keywords" content="running, nordic skiing, cross-country, mountain biking, road races, snowshoe, race, timing, ">
<meta name="description" content="A Fitness Event Timing Service specializing in road racing, nordic ski events, showshoe events, mountain bike events, and high school and college cross-country meet timing.">
<meta name="postinfo" content="/scripts/postinfo.asp">
<meta name="resource" content="document">
<meta name="distribution" content="global">

<script type="text/javascript" src="../../../misc/vira.js"></script>
<link rel="stylesheet" type="text/css" href="../../../misc/vira.css">

<style type="text/css">
<!--
td{
	color:#000000;
	font-size:10pt;
	padding:5px;
	border:1px solid #000000
	}
-->
</style>

</head>
<body style="margin:5px;background-color:#ffffff;background-image:none">
<a href="javascript:window.print()">Print</a><br />
	<table style="border:1px solid #000000">
		<tr>
			<td class="table_head" style="padding-left:100px;text-align:left;color:#000000" colspan="9">
				<br />
				Meet Sheet for:________________________________________________________<br /><br />
				<span style=" font-size:8pt;color:#000000;font-weight:normal">(A customized version of this form is available for your 
				team by logging into your account at www.gopherstateevents.com)</span>
			</td>
		</tr>
		<tr>
			<td style="font-weight:bold;border:none" colspan="2">
				Race Name:
			</td>
			<td style="border:none">
				&nbsp;
			</td>
			<td style="text-align:right;font-weight:bold;border:none" colspan="2">
				Race Time: 
			</td>
			<td style="border:none" colspan="5">
				&nbsp;
			</td>
		</tr>
		<tr>
			<td style="font-weight:bold;padding:2px;width:10px" rowspan="2" valign="bottom">
				No.
			</td>
			<td style="font-weight:bold;padding:2px;width:150px;" rowspan="2" valign="bottom">
				Name
			</td>
			<td style="font-weight:bold;padding:2px" rowspan="2" valign="bottom">
				Gr
			</td>
			<td style="font-weight:bold;padding-left:10px;padding-right:10px" rowspan="2" valign="bottom">
				Pl
			</td>
			<td style="font-weight:bold;padding-left:25px;padding-right:25px" rowspan="2" valign="bottom">
				Time
			</td>
			<td style="font-weight:bold;padding:2px;text-align:center" colspan="3">
				Splits
			</td>
			<td style="text-align:center;font-weight:bold;padding:2px;width:200px" rowspan="2" valign="bottom">
				Comments
			</td>
		</tr>
		<tr>
			<td style="font-weight:bold;padding-left:10px;padding-right:10px" nowrap="nowrap">
				Split 1
			</td>
			<td style="font-weight:bold;padding-left:10px;padding-right:10px" nowrap="nowrap">
				Split 2
			</td>
			<td style="font-weight:bold;padding-left:10px;padding-right:10px" nowrap="nowrap">
				Split 3
			</td>
		</tr>
		<%For i = 0 To 24%>
			<tr>
				<td style="text-align:right;"><%=i + 1%>)</td>
				<%For j = 0 To 7%>
					<td>&nbsp;</td>
				<%Next%>
			</tr>
		<%Next%>
	</table>
</body>
</html>
