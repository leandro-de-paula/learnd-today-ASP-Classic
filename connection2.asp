<%
Conn_SQL = "UID=root; PWD=root; DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=forms"
 Set Connect = Server.CreateObject("ADODB.Connection")
 Connect.Open Conn_SQL
%>