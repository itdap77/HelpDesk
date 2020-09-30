<!--#INCLUDE file="library/Funciones.asp"-->

<script language="VBScript" runat="Server" type="text/vbscript">
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

dim str

set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_HelpDesk_STRING

select case vDataForm("accion")
	case "A"	    'Agregar
			
		 str  = "{call dbo.SPU_CIUDAD_A ('" + 	replace(vDataForm("newciudadalta"), "'", "''") + "','" +_												
												replace(vDataForm("telefonoalta"), "'", "''") + "','" + _
												replace(vDataForm("internoalta"), "'", "''") + "'," +_	
												replace(vDataForm("paisalta"), "'", "''")    + ",'" +_
												replace(vDataForm("sigla"),"'","''") + "'" +_
												")}"

	case "M" 	'Modificar

		 str = "{call dbo.SPU_CIUDAD_M ('" + replace(vDataForm("ciudadmod"), "'", "''") + "','" + _
											replace(vDataForm("telefonomod"), "'", "''") + "','" + _
											replace(vDataForm("internomod"), "'", "''") + "'," + _
											replace(vDataForm("paismod"), "'", "''") + "," + _
											replace(vDataForm("idciudad"), "'", "''") _
											+ ")}"
											
										
	case "B"	    'Borrar

	
		 str = "{call dbo.SPU_CIUDAD_B ('" + replace(vDataForm("ciudadesbaja"), "'", "''") + "')}"


end select

rs.Source = str
rs.open

redir = "abm_ciudad.asp?lang=" & vDataForm("lang")
Response.Redirect (redir)

</script>