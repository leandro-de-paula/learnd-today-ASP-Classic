<% 
'Criando o objeto de conexão
Set conn = Server.CreateObject("ADODB.Connection") 
 
'Abrindo uma conexão com o banco de dados - [IMPORTANTE] altere os dados abaixo com as informações de sua base de dados
'conn.Open("DRIVER={MySQL ODBC 5.3 ANSI Driver};SERVER=localhost;DATABASE=forms;USER=root;PASSWORD=root;")

conn.Open("Driver={MySQL ODBC 5.3 Unicode Driver};Server=127.0.0.1;Database=forms;UID=root;PWD=root;")
 
response.write "banco conectado" 
 
'Fechando a conexão com o banco de dados
conn.Close()
 
'Destruindo o objeto
Set conn = Nothing
%>

