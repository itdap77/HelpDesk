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
			
		str = "{call dbo.SPU_PAIS_A ('" + replace(vDataForm("alta_pais"), "'", "''") + "','" + _
		replace(vDataForm ("sigla"),"'", "''") + "'" + _
		 ")}"

	case "M"	'Modificar

		str = "{call dbo.SPU_PAIS_M ('" + replace(vDataForm("paismod"), "'", "''") + "'," + _
			replace(vDataForm("idpais"), "'", "''") + ",'" + _
			replace(vDataForm ("sigla_modificacion"),"'", "''") + _
			"')}"

	case "B"	'Borrar

		str = "{call dbo.SPU_PAIS_B (" + replace(vDataForm("paises_baja"), "'", "''") + ")}"

end select


rs.Source = str
rs.Open()

redir = "abm_pais.asp?lang=" & vDataForm("lang")
Response.Redirect (redir)

</script>