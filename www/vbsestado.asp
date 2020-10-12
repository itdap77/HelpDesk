<% @ LANGUAGE=VBScript %>
<!--#INCLUDE file="library/Funciones.asp"-->

<script language="VBScript" runat="Server" type="text/vbscript">
ValidSession()

If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 

dim str

set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_HelpDesk_STRING

If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 

select case vDataForm("accion")
	case "A"	'Agregar
			
		str = "{call dbo.SPU_ESTADO_A ('" + replace(vDataForm("alta_estado"), "'", "''") + "')}"


	case "M"	'Modificar

		str = "{call dbo.SPU_ESTADO_M ('" + replace(vDataForm("modificacion_estado"), "'", "''") + "'," + _
		replace(vDataForm("estados_modificacion"), "'", "''") + ")}"

	case "B"	'Borrar

		str = "{call dbo.SPU_ESTADO_B ('" + replace(vDataForm("estados_baja"), "'", "''") + "')}"
		
end select

rs.Source = str
rs.Open()
Response.Redirect ("abm_estados.asp")		

</script>