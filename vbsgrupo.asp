<% @ LANGUAGE=VBScript %>
<!--#INCLUDE file="library/Funciones.asp"-->

<script language="VBScript" runat="Server" type="text/vbscript">
ValidSession ()

If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 


dim str

set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_HelpDesk_STRING

select case vDataForm("accion")
	case "A"	'Agregar
			
		str = "{call dbo.SPU_GRUPO_A ('" + replace(vDataForm("alta_grupo"), "'", "''") + "')}"

	case "M"	'Modificar

		str = "{call dbo.SPU_GRUPO_M ('" + replace(vDataForm("modificacion_grupo"), "'", "''") + "'," + _
		replace(vDataForm("grupos_modificacion"), "'", "''") + ")}"


	case "B"	'Borrar

		str = "{call dbo.SPU_GRUPO_B ('" + replace(vDataForm("grupos_baja"), "'", "''") + "')}"

		
end select

rs.Source = str
rs.Open()
Response.Redirect ("abm_grupo.asp")
'Response.Write "es:" & str
</script>