<%@ LANGUAGE=VBScript %>

<!--#INCLUDE Virtual="itdap/Helpdesk/library/Funciones.asp"-->

<script language="VBScript" runat="Server" type="text/vbscript">

Dim vBackToURL,oRSrole

	If Not IsEmpty(Request.Form) Then

		If Len(Request.Form("BackToURL")) = 0 Then
			vBackToURL = "main.asp"
		Else
			vBackToURL = Request.Form("BackToURL")
		End If

		Select Case Ucase(Request.Form("Accion"))
		Case "LOGIN"		'IniciarSesion
			IniciarSesion Request.Form("IDGrupo"),Request.Form("Password"), Request.Form("IDCiudad"), lcase(Request.Form("Usuario"))

		Case "LOGOUT"		'CerrarSesion
			CerrarSesion

		Case "CHANGE-PASSWORD"		'Cambiar contraseña
			CambiarContrasenia Request.Form("Usuario"),Request.Form("Password"), Request.Form("PasswordNew")

		Case else
			Response.Redirect "mensaje.asp?Msg=Error en la aplicación.&BackToURL=" & vBackToURL
		End Select

		Response.Redirect vBackToURL
	Else
		Response.Redirect "mensaje.asp?Msg=Error en la aplicación.&BackToURL=" & vBackToURL
	End If

Function CerrarSesion



	Session("IDGrupo") = 0
	Session("NombreUsuario") = ""
	Session("IDU") = 0
	Session("FullName") = ""
	Session("Login") = 0
	Session("Password") = ""
	Session("IDCiudad") = 0
	Session("IDidioma") = 0


End Function

Function IniciarSesion(IDGrupo, Password, IDCiudad, Usuario)
On error Resume Next
If Session("Login") <> 0 Then
		CerrarSesion
End If

Dim oRS__Login
Dim oRS__Password, SP_Login

set oRS = Server.CreateObject("ADODB.Recordset")

oRS.ActiveConnection = MM_HelpDesk_STRING

SP_Login = "{call dbo.SPU_Usuarios_Login('" + Replace(Usuario, "'", "''") + "','" + Replace(Password, "'", "''") +   "')}"
oRS.Source = SP_Login
oRS.Open()


	If oRS.EOF = True Then


		IniciarSesion = 3
		IDGrupo = ""
	  Response.redirect("default.asp?IniciarSesion=3")
	Else


		If (oRS.Fields.Item("Estado").Value) = 1 Then
		   IniciarSesion = 2		'Usuario Cancelado
		   IDGrupo = oRs("IDGrupo")
		   Response.redirect("default.asp?IniciarSesion=2")
		Else

				IDGrupo = oRs("IDGrupo")
				IniciarSesion = 1		'Login OK

				Session("IDGrupo") = IDGrupo
				Session("NombreUsuario") = Usuario
				Session("IDU") = oRs("IDUsuario")
				Session("FullName") = oRs("Nombre") + " " + oRs("Apellido")
				Session("IDCiudad") = oRs("IDCiudad")
				'Session("Password") = oRs("Password")
			End If
	End If

	oRS.Close
	Set oRS = Nothing

	Session("Login") = IniciarSesion

End Function

Function CambiarContrasenia(Usuario, Password, PasswordNew)


set oRS = Server.CreateObject("ADODB.Recordset")
oRS.ActiveConnection = MM_HelpDesk_STRING
oRS.Source = "{call dbo.SPU_Usuarios_LE('" + Replace(Usuario, "'", "''") + "','" + Replace(Password, "'", "''") + "','" + "" + "')}"
oRS.Open()

IF oRs.EOF <> true then
	Usuario = oRS.Fields.Item("Login").Value
	Passwordnew = Request.Form ("PasswordNew")
else IDUsuario = ""
	Passwordnew = ""
end If


	If oRS.EOF = True Then
		CambiarContrasenia = 0		'Usuario inválido
	Else
		If oRS.Fields.Item("Estado").Value = 1 Then
		   CambiarContrasenia = 0		'Usuario Cancelado
		Else
			If Password <> oRS.Fields.Item("Password").Value Then
				CambiarContrasenia = 0		'Bad Password
			Else
				redim vParam(1)

				vParam(0) = Usuario
				vParam(1) = PasswordNew



			Dim ST
			Set oConn = Server.CreateObject ("ADODB.Connection")
			oConn.Open = MM_HelpDesk_STRING

			ST = "EXEC SPU_UsuariosPassword_M " & Usuario & "," & PasswordNew
			Oconn.Execute ST
			oConn.Close
			Set ocon = nothing

			CambiarContrasenia = 1

			End If
		End If
	End If

	Session("CambiarContrasenia") = CambiarContrasenia

	Set oRS = Nothing

End Function

</script>
