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
	Dim FechaAlta 
	
	FechaAlta = getdate()
	
	Str= "{call dbo.SPU_Permisos_A ('" + _
	replace(vDataForm("IDModulo"), "'", "''") & "," + _
	replace(vDataForm("IDUsuario"), "'", "''") & "," + _
	replace(vDataForm("Tipo"), "'", "''") & ",'" + _
	replace(vDataForm("Permisos"), "'", "''") & "'," + _
	replace(vDataForm("Estado"), "'", "''") & ",'" + _
	FechaAlta & "','" + _        
	FechaAlta & "'," + _	       
	replace(vDataForm("IDUsuarioMod"), "'", "''") & ",'" + _
	FechaAlta "'" + _	
	"')}"

	rs.Source = Str
	rs.Open()
		
case "M"	'Modificar

	Str= "{call dbo.SPU_Permisos_M ('" + _
	replace(vDataForm("IDModulo"), "'", "''") & "," + _
	replace(vDataForm("IDUsuario"), "'", "''") & "," + _
	replace(vDataForm("Tipo"), "'", "''") & ",'" + _
	replace(vDataForm("Permisos"), "'", "''") & "'," + _
	replace(vDataForm("Estado"), "'", "''") & ",'" + _
	replace(vDataForm("FechaEstado"), "'", "''") & ",'" + _ 	'Fecha Estado
	replace(vDataForm("IDUsuarioMod"), "'", "''") & ",'" + _
	replace(vDataForm("FechaModificacion"), "'", "''") & ",'" + _	'Fecha Modificacion
	"')}"

	rs.Source = Str
	rs.Open()

case "B"	'Borrar
	Str= "{call dbo.SPU_Permisos_B ('" + replace(vDataForm("IDPermisos"), "'", "''") + "')}"

	rs.Source = Str
	rs.Open()
end select

Response.Redirect ("permisos.asp")
rs.close()
</script>