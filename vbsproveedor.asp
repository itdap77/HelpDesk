<% @ LANGUAGE=VBScript %>
<!--#INCLUDE file="library/Funciones.asp"-->

<script language="VBScript" runat="Server" type="text/vbscript">
ValidSession()

dim Str, FechaAlta

set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_HelpDesk_STRING

If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 


select case vDataForm("accion")
case "A"	'Agregar
		
	Str= "{call dbo.SPU_Proveedor_A (" + _
	"'" & replace(vDataForm("Proveedor"), "'", "''") & "'," + _
	"'" &  replace(vDataForm("Telefono"), "'", "''") & "'," + _
	"'" &  replace(vDataForm("Interno"), "'", "''") & "'," + _
	"'" &  replace(vDataForm("Contacto"), "'", "''") & "'," + _
	"'" &  replace(vDataForm("Email"), "'", "''") & "'," + _
	"'" &  replace(vDataForm("PaginaWeb"), "'", "''") & "'," + _
	"'" &  replace(vDataForm("Direccion"), "'", "''") & "'," + _
	 replace(vDataForm("IdPais"), "'", "''") & "," + _
	 replace(vDataForm("IdCiudad"), "'", "''") & "," + _
	"'" &  replace(vDataForm("CodigoPostal"), "'", "''") & "'" + _
	")}"

	rs.Source = Str
'response.Write(str)	
rs.Open()
Response.Redirect ("abm_alta_Proveedor.asp")	
	
case "M"	'Modificar

	Str= "{call dbo.SPU_Proveedor_M (" + _
	"'" & replace(vDataForm("IdProveedor"), "'", "''") & "'," + _
	"'" & replace(vDataForm("Proveedormod"), "'", "''") & "'," + _
	"'" & replace(vDataForm("Telefono"), "'", "''") & "'," + _
	"'" & replace(vDataForm("Interno"), "'", "''") & "'," + _
	"'" & replace(vDataForm("Contacto"), "'", "''") & "'," + _
	"'" & replace(vDataForm("Email"), "'", "''") & "'," + _
	"'" & replace(vDataForm("PaginaWeb"), "'", "''") & "'," + _
	"'" & replace(vDataForm("Direccion"), "'", "''") & "'," + _
	replace(vDataForm("IdPais"), "'", "''") & "," + _
	replace(vDataForm("IdCiudad"), "'", "''") & "," + _
	"'" & replace(vDataForm("CodigoPostal"), "'", "''") & "'" + _
	")}"

	rs.Source = Str
'response.Write(str)
	rs.Open()
	Response.Redirect ("abm_modificacion_Proveedor.asp")

case "B"	'Borrar
	Str= "{call dbo.SPU_Proveedor_B ('" + replace(vDataForm("IDProveedor"), "'", "''") + "')}"

	rs.Source = Str
	rs.Open()
	Response.Redirect ("abm_baja_Proveedor.asp")

end select

rs.close()

</script>