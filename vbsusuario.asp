<% @ LANGUAGE=VBScript %>
<!--#INCLUDE file="library/Funciones.asp"-->

<script language="VBScript" runat="Server" type="text/vbscript">
ValidSession()

dim str, Cadena,fechamod
dim Sector, Grupo, Pais, Ciudad, pc

set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_HelpDesk_STRING

If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 

If lcase(Session("NombreUsuario")) <> "demo" Then 'Run only if different for the demo account

        select case vDataForm("accion")
        case "A"	'Agregar
			
		        If vDataForm("login") = "" Then
		           ' Salida rapida si no se indico el login...
		           Response.Redirect ("alta_usuario.asp")
		        End If

		        Cadena = "'" & request.form("dia") & "/" & request.form("mes") & "/" & request.form("anio") & "'"

		        ' Cargar Valores por Default.
		        Sector = vDataForm("sector")
        '		If CInt(Sector) = 0 Then
        '		   Sector = "1"
        '		End If

		        Grupo = vDataForm("grupo")
		        If Grupo = "----" Then
		           Grupo = "10"	' Usuarios.
		        End If

		        Pais = replace(vDataForm("pais"), "'", "''")
		        If CInt(Pais) = 0 Then
		           Pais = "1"   ' Argentina.
		        End If

		        Ciudad = replace(vDataForm("ciudad"), "'", "''")
		        If CInt(Ciudad) = 0 Then
		           Ciudad = "15"  ' Buenos Aires.
		        End If


		        str = "{call dbo.SPU_USUARIOS_A ('" + replace(vDataForm("login"), "'", "''") + "','" + _
											          replace(vDataForm("nombre"), "'", "''") + "','" + _
											          replace(vDataForm("apellido"), "'", "''") + "','" + _	
											          replace(vDataForm("direccion"), "'", "''") + "','" + _	
											          replace(vDataForm("codigopostal"), "'", "''") + "','" + _	
											          replace(vDataForm("telefono_particular"), "'", "''") + "','" + _	
											          replace(vDataForm("mail"), "'", "''") + "'," + _	
											          Cadena + ",'" + _
											          replace(vDataForm("telefono_laboral"), "'", "''") + "','" + _		
											          replace(vDataForm("interno"), "'", "''") + "','" + _		
											          replace(vDataForm("legajo"), "'", "''") + "'," + _		
											          Pais + "," + _	
											          Ciudad + ",'" + _	
											          replace(vDataForm("password"), "'", "''") + "'," + _	
											          Grupo + "," + _	
											          Sector + ",'" + _	
											          replace(vDataForm("pc"), "'", "''")  + "','" + _
											          replace(vDataForm("empresa"), "'", "''")  + "'" + _	
											           ")}"


		        rs.Source = str
		        rs.Open()
	
	        redir = "alta_usuario.asp?lang=" & vDataForm("lang")
            Response.Redirect (redir)

			

        case "M"	'Modificar

		        FechaMod = vDataForm("anio") & "/" & vDataForm("mes") & "/" & vDataForm("dia") 

		        If vDataForm("login") = "" Then
		           ' Invalid data to save!, skip this request..
		        Else

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
											          vDataForm("sector") + "," + _	
		  									          vDataForm("Estado") + ",'" + _
											          replace(vDataForm("loginmod"), "'", "''")  + "','" + _
											          replace(vDataForm("empresa"), "'", "''")  + "'" + _	
											           ")}"
		          rs.Source = Str 
		          rs.Open()
		        End If


        redir = "modificacion_usuario.asp?lang=" & vDataForm("lang")
        Response.Redirect (redir)

        case "B"	'Borrar

		        str = "{call dbo.SPU_USUARIOS_B ('" + replace(vDataForm("usuarios"), "'", "''") + "')}"
		        rs.Source = str
		        rs.Open()

        redir = "baja_usuario.asp?lang=" & vDataForm("lang")
        Response.Redirect (redir)

		
		
		
        end select
Else
        redir = "main.asp?lang=" & vDataForm("lang")
        Response.Redirect (redir)

End If


</script>