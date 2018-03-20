<% Option Explicit %>
<html>
<head>
<title>ASP Page FTW</title>
</head>
<body bgcolor="white" text="black">
<h1>ASP Page FTW</h1>
<ul>
<li><a href='default.htm'>Static HTML Page</a></li>
</ul>
<h1>Some Code</h1>
<ul>
<%
 Dim i

 For i = 1 to 10
  Response.Write("<li>Hello " & i & "</li>")
 Next
%>
</ul>
</body>
</html>