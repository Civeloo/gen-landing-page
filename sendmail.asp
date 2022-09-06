<%@LANGUAGE="VBSCRIPT" %>
<!-- Cambiar "cuenta@dominio.com" por una cuenta de correo de su dominio -->
<!--METADATA TYPE="TypeLib" FILE="C:\WINDOWS\system32\cdosys.dll" -->
<h1>PRONTO NOS PONDREMOS EN CONTACTO CON USTED...</H1>
<!-- Formulario para completar con los datos -->
 <!-- <form action="sendmail.asp" method="POST">
	E-mail destinatario: <input type="text" name="email" width="50"></input><br/>
	<input type="submit" value="Enviar e-mail" /><input type="hidden" name="enviar" value="1"/>
</form>  -->
<!-- Fin Formulario para completar con los datos -->

<%
' Se verifica que los datos han sido enviados desde el formulario, para la validaci�n con el SMTP
If Request("enviar") = 1 Then	
		' Se crean los objetos necesarios para el env�o del correo
		Set oMail = Server.CreateObject("CDO.Message") 
		Set iConf = Server.CreateObject("CDO.Configuration") 
		Set Flds = iConf.Fields 
		
		' Se configuran los parametros necesarios para el env�o
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost" 
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10 
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Update
		' Se asignan las propiedades de configuraci�n al objeto
		Set oMail.Configuration = iConf 
	
		' Destinatario del correo
		oMail.To = "consultas@civeloo.com"
		' Remitente del correo
		oMail.From = "no-reply@civeloo.com"
		' Subject o asunto
		oMail.Subject = Request("email")
		' Cuerpo del mensaje
		oMail.TextBody = Request("name") +" tel: "+ Request("phone") + "  mensaje: "+ Request("message")
		' Se env�a el correo
		oMail.Send
		' Se destruyen los objetos
		Set iConf = Nothing 
		Set Flds = Nothing

End If
%>