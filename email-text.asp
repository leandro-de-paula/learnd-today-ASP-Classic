<%
Set myMail = CreateObject("CDO.Message")
myMail.Subject = "Sending email with CDO"
myMail.From = "leandrodepaula.ti@gmail.com"
myMail.To = "leandrodepaula.ti@gmail.com"
myMail.TextBody = "This is a message."
myMail.Send
set myMail = nothing
%>