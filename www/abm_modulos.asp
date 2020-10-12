<% 
If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 

If (Session("Login") <> 1) Then
	If vDataForm("lang") = "esp" Then
		Response.Redirect ("mensaje.asp?Msg=Usted no esta logueado en el sistema.&backtourl=default.asp")
	Else
		Response.Redirect ("mensaje.asp?Msg=You are not logged to the system.&backtourl=default.asp")
	End If
End If 
%>

<form name="agregarmodulo">
                  <p>
                  <input type="text" name="NombreModulo" class="formtable" size="10">
                  <input type="hidden" name="accion" value="A">
                  <input type="button" value="Agregar" name="Agregar" class="FormButton" onclick="javascript: submitfrmmodulo();">
                  </p>
                </form>