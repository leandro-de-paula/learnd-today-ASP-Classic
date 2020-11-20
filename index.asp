<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="Windows 1252">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tuturial ASP</title>
</head>
<body>
<%

'Declarar uma variável
dim name
name="Leandro"

    response.write("<p style='color:#0000ff'> Hello World: " & name &  "</p>" & "<br>")


'criar arrays
dim testeName(5), i
testeName(0) = "Leandro Novamente"
testeName(1) = "Evellyn Fontenele"
testeName(2) = "Julia"
testeName(3) = "Hege"
testeName(4) = "Stale"
testeName(5) = "Borge"

For i = 0 To 5
     response.write(testeName(i) & "<br>")
Next


'Percorrendo os cabeçalhos HTML
dim x
For x = 1 To 6 
    response.write("<h" & x & ">Consegui " & x & "</h" & x & ">" )
Next


'Saudação baseado em hora usando VBScript
dim head
h = hour(now())

response.write("<p>" & now())
response.write("</p>")

If h<12 Then
    ' true
    response.write("Good Morning!")
Else
    ' false
    response.write("Good day!")
End if



'Saudação baseado em hora usando JavaScript dentro de 
'VBScript tem problams de rederização em codigos iniciados em VBScript
response.write("<br>")
response.write("<br>")


'criar e alterar uma variável
dim firstname
firstname = "Evellyn"
response.write(firstname)
response.write("<br>")
firstname = "Caroline"
response.write(firstname)


'Inserir um valor de variável em um texto
dim nome
nome = "Evellyn Caroline"
response.write("<h1>")
response.write(nome)
response.write("</h1>")

response.write("<h3>" & nome & "</h3>")




%>
</body>
</html>