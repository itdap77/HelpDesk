<% @ language=VBScript %>
<!--#INCLUDE file="library/Funciones.asp"-->

<script type="text/vbscript" language="VBScript" RUNAT="Server">
ValidSession()

dim str

set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_HelpDesk_STRING


If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If



if vDataForm("estimacion_a") = "" Then
        estimacion_a = "0"
Else
        estimacion_a = vDataForm("estimacion_a") 
End If

if vDataForm("estimacion_m") = "" Then
        estimacion_m = "0"
        Else
        estimacion_a = vDataForm("estimacion_m") 
End If

select case vDataForm("accion_problema")
	case "A"	'Agregar
	
		str  = "{call dbo.SPU_Problema_A ('" + _
		        replace(vDataForm("detalleproblema_a"), "'", "''") + "','" + _							
				replace(vDataForm("problema_a"), "'", "''") + "'," + _
				estimacion_a + ",'" + _
				replace(vDataForm("categoria_a"), "'", "''")  + "','" + _	
				replace(vDataForm("problemdetail_a"), "'", "''")  + "'" + _	
					
				")}"
			
	case "M"	'Modificar

		str = "{call dbo.SPU_Problema_M (" + _
											replace(vDataForm("idproblema_m"), "'", "''") + ",'" & _
											replace(vDataForm("detalleproblema_m"), "'", "''") + "'," & _
											replace(vDataForm("categoria_m"), "'", "''") + "," & _
											estimacion_m  + ",'" & _
											replace(vDataForm("problemdetail_m"), "'", "''")  & _
										    "')}"
									
		
	case "B"	'Borrar

		str = "{call dbo.SPU_Problema_B (" & replace(vDataForm("problemab"), "'", "''") & ")}"
		
end select

rs.Source = str
rs.Open()


redir = "abm_problemas.asp?lang=" & vDataForm("lang")
Response.Redirect (redir)

</script>