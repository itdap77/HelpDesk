<% @ LANGUAGE=VBScript %> 

<!--#INCLUDE file="library/Funciones.asp"-->


<script language="VBScript" runat="Server" type="text/vbscript">
ValidSession()
    
If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 


If lcase(Session("NombreUsuario")) <> "demo" Then 'Run only if different for the demo account
        dim Str,FechaMod

        set rs = Server.CreateObject("ADODB.Recordset")
        rs.ActiveConnection = MM_HelpDesk_STRING


        FechaMod = vDataForm("anio") & "/" & vDataForm("mes") & "/" & vDataForm("dia") 

        Str = "{call dbo.SPU_USUARIOS_M ('" + replace(vDataForm("login"), "'", "''") + "','" + _
									          replace(vDataForm("nombre"), "'", "''") + "','" + _
									          replace(vDataForm("apellido"), "'", "''") + "','" + _	
									          replace(vDataForm("direccion"), "'", "''") + "','" + _	
									          vDataForm("cp") + "','" + _	
									          replace(vDataForm("tep"), "'", "''") + "','" + _	
									          replace(vDataForm("mail"), "'", "''") + "','" + _	
									          FechaMod + "','" + _	
									          replace(vDataForm("te"), "'", "''") + "','" + _	
									          replace(vDataForm("interno"), "'", "''") + "','" + _	
									          replace(vDataForm("legajo"), "'", "''") + "'," + _	
									          vDataForm("paismod") + "," + _	
									          vDataForm("ciudadmod") + ",'" + _	
									          replace(vDataForm("password"), "'", "''") + "'," + _	
									          vDataForm("IDGrupo") + "," + _	
									          vDataForm("IDSector") + "," + _	
  									        vDataForm("Estado") + ",'" + _										  
									          replace(vDataForm("loginmod"), "'", "''") + "','"  + _
									          vDataForm("empresa") + "'" + _										  
									          ")}"
			
								
        rs.Source = Str 
        rs.Open()
End If

redir = "modificacion_usuario_propio.asp?lang=" & vDataForm("lang")
Response.Redirect (redir)



</script>