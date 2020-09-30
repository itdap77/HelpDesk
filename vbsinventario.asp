<% @ LANGUAGE=VBScript %>
<!--#INCLUDE file="library/Funciones.asp"-->

<script language="VBScript" runat="Server" type="text/vbscript">
ValidSession()

dim Str, FechaAlta
dim Sector

set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_HelpDesk_STRING


If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 


If vDataForm("IdSector") <> 0 and vDataForm("Idciudad") <> 0 and vDataForm("Idparte") <> 0 Then 
	set Rs_Idpc = Server.CreateObject("ADODB.Recordset")
	Rs_Idpc.ActiveConnection = MM_HelpDesk_STRING
	Rs_Idpc.Source = "{call dbo.SPU_Inventario_ID (" + request.form ("IDParte") + "," + request.form ("idciudad") + "," + request.form ("IDSector") + ")}"
	Rs_Idpc.Open()
End If

select case vDataForm("accion")
case "A"	'Agregar
		
	Str= "{call dbo.SPU_inventario_A (" + _
	"'" &  cstr(Rs_idpc.fields.item("idpc").value) & "'," + _
	replace(vDataForm("IdParte"), "'", "''") & "," & _
	"'" & replace(vDataForm("marca"), "'", "''") & "'," + _
	"'" &  replace(vDataForm("nroserie"), "'", "''") & "'," + _
	"'" &  replace(vDataForm("modelo"), "'", "''") & "'," + _
	replace(vDataForm("idproveedor"), "'", "''") & "," + _
	"'" &  replace(vDataForm("soporte"), "'", "''") & "'," + _
	"'" &  replace(vDataForm("garantia"), "'", "''") & "'," + _
	"'" &  replace(vDataForm("fechavencimiento"), "'", "''") & "'," + _
	 replace(vDataForm("idsector"), "'", "''") & "," + _
	 replace(vDataForm("Inventario"), "'", "''") & "," + _
	 replace(vDataForm("IdCiudad"), "'", "''") + _
	")}"

	'vBackToUrl = "abm_alta_inventario.asp"

	Response.Redirect ("abm_alta_inventario.asp")
	
case "M"	'Modificar
If vDataForm("cambioid") = "1" Then 

	Str= "{call dbo.SPU_inventario_M (" + _
	replace(vDataForm("idinventario"), "'", "''") & "," + _
	"'" &  cstr(Rs_idpc.fields.item("idpc").value) & "'," + _
	replace(vDataForm("IdParte"), "'", "''") & "," + _
	"'" & replace(vDataForm("marca"), "'", "''") & "'," + _
	"'" &  replace(vDataForm("nroserie"), "'", "''") & "'," + _
	"'" &  replace(vDataForm("modelo"), "'", "''") & "'," + _
	replace(vDataForm("idproveedor"), "'", "''") & "," + _
	"'" &  replace(vDataForm("soporte"), "'", "''") & "'," + _
	"'" &  replace(vDataForm("garantia"), "'", "''") & "'," + _
	"'" &  replace(vDataForm("fechavencimiento"), "'", "''") & "'," + _
	 replace(vDataForm("idsector"), "'", "''") & "," + _
	 replace(vDataForm("inventario"), "'", "''") & "," + _
	 replace(vDataForm("IdCiudad"), "'", "''") + _
	")}"

else
     	Str= "{call dbo.SPU_inventario_M (" + _
	replace(vDataForm("idinventario"), "'", "''") & ",'" + _
	replace(vDataForm("idpc"), "'", "''") & "'," + _ 
	replace(vDataForm("IdParte"), "'", "''") & "," + _
	"'" & replace(vDataForm("marca"), "'", "''") & "'," + _
	"'" &  replace(vDataForm("nroserie"), "'", "''") & "'," + _
	"'" &  replace(vDataForm("modelo"), "'", "''") & "'," + _
	replace(vDataForm("idproveedor"), "'", "''") & "," + _
	"'" &  replace(vDataForm("soporte"), "'", "''") & "'," + _
	"'" &  replace(vDataForm("garantia"), "'", "''") & "'," + _
	"'" &  replace(vDataForm("fechavencimiento"), "'", "''") & "'," + _
	 replace(vDataForm("idsector"), "'", "''") & "," + _
	 replace(vDataForm("inventario"), "'", "''") & "," + _
	 replace(vDataForm("IdCiudad"), "'", "''") + _
	")}"
		
	'vBackToUrl = "abm_modificacion_inventario.asp"

	Response.Redirect ("abm_modificacion_inventario.asp")


'	@IDInventario as int,
'@IdPc  varchar(8),
'@IDParte int,
'@Marca varchar(30), 
'@NroSerie varchar(50), 
'@Modelo  varchar(50), 
'@IdProveedor int,
'@Soporte varchar(20), 
'@Garantia varchar(20), 
'@FechaVencimiento smalldatetime,
'@IDsector int, 
'@Inventario as int,
'@IDCiudad int

end if

	

case "B"	'Borrar


	Str= "{call dbo.SPU_inventario_B ('" + replace(vDataForm("inventario"), "'", "''") + "')}"
	vBackToUrl = "abm_baja_inventario.asp"
	
end select


	rs.Source = Str
	rs.Open()
	
'	RESPONSE.WRITE STR
'	Response.Write "cambioid=" & vdataForm("cambioid")
	
	
'	Response.Redirect "mtMensaje.asp?Msg=Error en la aplicación.&BackToURL=" & vBackToURL 

	Response.Redirect ("abm_baja_inventario.asp")


rs.Close()

</script>