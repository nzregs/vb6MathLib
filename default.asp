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
<table border="1" cellspacing="0" cellpadding="2">
  <thead>
    <th width=100>x</th>
    <th width=100>y</th>
    <th width=200>x+y</th>
  </thead>
<%
Response.Buffer = true
Dim objMathLib

Set objMathLib = Server.CreateObject( "MathLibNZRegs.Addition" )
Dim x 
Dim y 
Dim answer
For y = 14 to 100 step 3
x = 10
answer = objMathLib.Add(CInt(x),CInt(y))
Response.Write("<tr><td align=center>"& x &"</td><td align=center>"& y & "</td><td align=center>" &  answer & "</td></tr>")
Next
%>
</table>
<table border="1" cellspacing="0" cellpadding="2">
    <thead>
      <th width=100>x</th>
      <th width=100>y</th>
      <th width=200>x*y</th>
    </thead>
  <%
  Response.Buffer = true
  For y = 14 to 100 step 3
  x = 10
  answer = objMathLib.Multiply(CInt(x),CInt(y))
  Response.Write("<tr><td align=center>"& x &"</td><td align=center>"& y & "</td><td align=center>" &  answer & "</td></tr>")
  Next
  %>
  </table>
</ul>
</body>
</html>