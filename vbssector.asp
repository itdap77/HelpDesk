<% @ LANGUAGE=VBScript %>
<!--#INCLUDE file="library/Funciones.asp"-->

<script language="VBScript" runat="Server" type="text/vbscript">
ValidSession()

dim Str

set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_HelpDesk_STRING


If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 


select case vDataForm("accion")
	case "A"	'Agregar
	
		str = "{call dbo.SPU_SECTOR_A ('" + replace(vDataForm("newsectoralta"), "'", "''") + "'," + _
										replace(vDataForm("ciudadalta"), "'", "''") + ",'" + _
										replace(vDataForm("sigla"), "'", "''") + "'" + _
									")}"

	case "M"	'Modificar

		str = "{call dbo.SPU_SECTOR_M ('" + replace(vDataForm("sectoresmod"), "'", "''") + "'," + _
										replace(vDataForm("sectormod"), "'", "''") + ")}"

	case "B"	'Borrar

		str = "{call dbo.SPU_SECTOR_B ('" + replace(vDataForm("sectoresbaja"), "'", "''") + "')}"

end select

rs.Source = str
rs.Open()
Response.Redirect ("abm_sector.asp")


</script>