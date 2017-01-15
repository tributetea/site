<%
For Each x In Request.Form
 message=message & x & ": " & Request.Form(x) & CHR(10)
Next

set smtp = Server.CreateObject("Bamboo.SMTP")
' You only need to change the smtp.Rcpt ans smpt.from part to your email address
smtp.Server = "mail.<domain name>"
smtp.Rcpt = "YOUR FORM WILL BE SEND TO THE E-MAIL ADDRESS WRITTEN HERE"
smtp.From = "THIS SHOULD BE YOUR E-MAIL ADDRESS"
smtp.FromName = Request.ServerVariables("HTTP_REFERER")
smtp.Subject = "Your web form - " & Request.ServerVariables("HTTP_REFERER")
smtp.Message = message
on error resume next
smtp.Send
if err then
 response.Write err.Description 
else
 response.Write ("Thank you for your submission.... Your message has been delivered successfully.")
end if
set smtp = Nothing
%>