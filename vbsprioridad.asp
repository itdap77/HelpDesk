<% @ LANGUAGE=VBScript %>
<!--#INCLUDE file="library/Funciones.asp"-->

<script language="VBScript" runat="Server" type="text/vbscript">
ValidSession()

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
			
		str = "{call dbo.SPU_PRIORIDAD_A ('" + replace(vDataForm("alta_prioridad"), "'", "''") + "')}"

	case "M"	'Modificar

		str = "{call dbo.SPU_PRIORIDAD_M ('" + replace(vDataForm("modificacion_prioridad"), "'", "''") + "'," + _
		replace(vDataForm("prioridad_modificacion"), "'", "''") + ")}"

	case "B"	'Borrar

		str = "{call dbo.SPU_PRIORIDAD_B ('" + replace(vDataForm("prioridad_baja"), "'", "''") + "')}"

end select

rs.Source = str
rs.Open()
Response.Redirect ("abm_prioridad.asp")
'Response.Write "es:" & str
</script>