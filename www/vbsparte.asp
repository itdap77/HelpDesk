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
			
		str = "{call dbo.SPU_PARTE_A ('" + replace(vDataForm("alta_parte"), "'", "''") + "','" + _
								   replace(vDataForm("sigla"), "'", "''") + "')}"

	case "M"	'Modificar

		str = "{call dbo.SPU_PARTE_M ('" + replace(vDataForm("modificacion_parte"), "'", "''") + "'," + _
		replace(vDataForm("partes_modificacion"), "'", "''") + ")}"

	case "B"	'Borrar

		str = "{call dbo.SPU_PARTE_B ('" + replace(vDataForm("partes_baja"), "'", "''") + "')}"
		
end select

rs.Source = str
rs.Open()
Response.Redirect ("abm_parte.asp")

</script>