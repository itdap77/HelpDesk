<!--#INCLUDE file="Funciones.asp"-->

<%
Response.Expires = 0 
%>

<html>
<head>

<!--- STYLE SHEETS --->
<link rel="STYLESHEET" type="text/css" href="library/Styles.css" />
<link rel="STYLESHEET" type="text/css" href="library/NavArea.css" />
<link rel="STYLESHEET" type="text/css" href="library/MenuArea.css" />
<link rel="STYLESHEET" type="text/css" href="library/FormArea.css" />
<link rel="icon" href="favicon.png" type="image/png" />

<!--- LIBRARIAS DE FUNCIONES --->

<script type="text/javascript" language="javascript" src="library/script.js"></script>

<title>Helpdesk - ITDap Worldwide Solutions</title>
<%
ValidSession() 

If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 


%>

<!--- DECLARACION DEL SUBMENU --->
<% if vDataForm ("lang") = "esp" Then %>
<xml id=hpflyout>
<menu site="ITDap" subsite="Homepage">
	<submenu handle="usuarios">
	<item href="alta_usuario.asp?lang=esp" label="Alta / Add" />
	<item href="modificacion_usuario.asp?lang=esp" label="Modificacion / Modify" />
	<item href="baja_usuario.asp?lang=esp" label="Baja / Delete" />
	</submenu>

</menu>
</xml>
<%else %>
<xml id=hpflyout>
<menu site="ITDap" subsite="Homepage">
	<submenu handle="usuarios">
	<item href="alta_usuario.asp?lang=eng" label="Alta / Add" />
	<item href="modificacion_usuario.asp?lang=eng" label="Modificacion / Modify" />
	<item href="baja_usuario.asp?lang=eng" label="Baja / Delete" />
	</submenu>

</menu>
</xml>

<%End if %>
<!---  item href="abm_permisos.asp" label="Asignacion de Permisos" / --->

<!--- VARIABLES DEL DOCUMENTO --->

</head>

<body topmargin="0" leftmargin="0" link="#000000" vlink="#000000" alink="#000000" fieldOnFocus="<%=Request.Form("FieldOnFocus")%>" style="behavior: url('library/Document.htc');">

<%
Select Case vDataForm("lang")  
    Case "esp" 
	    Session("IDidioma") = 1
	    
    Case "eng" 
	    Session("IDidioma") = 2
	   
	Case Else
	   Session("IDidioma") = 2
End Select

Set Rs_LenguajeOpen = Server.CreateObject("ADODB.Recordset")
Rs_LenguajeOpen.ActiveConnection = MM_HelpDesk_STRING
Rs_LenguajeOpen.Source = "{call dbo.SPU_Lenguaje_V(" + cstr(Session("IDidioma")) + ",'pageopen.asp')}"
Rs_LenguajeOpen.Open()

Set Rs_Lenguaje = Server.CreateObject("ADODB.Recordset")
Rs_Lenguaje.ActiveConnection = MM_HelpDesk_STRING
Rs_Lenguaje.Source = "{call dbo.SPU_Lenguaje_V(" + cstr(Session("IDidioma")) + ",'" + GetPageName() + "')}"
Rs_Lenguaje.Open()

Set Rs_LenguajeClose = Server.CreateObject("ADODB.Recordset")
Rs_LenguajeClose.ActiveConnection = MM_HelpDesk_STRING
Rs_LenguajeClose.Source = "{call dbo.SPU_Lenguaje_V(" + cstr(Session("IDidioma")) + ",'" + "footernav.asp" + "')}"
Rs_LenguajeClose.Open()

%>

<!--- TABLA PRINCIPAL --->
<table id="MainTable" border="0"  class="defaultContent" cellspacing="0" cellpadding="0">
<tr>
	
<!--- AREA DE NAVEGACION --->

  <td class="NavArea" width="143"> 
    <table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr>
        <td class="Titulo" style="text-align:center;">
        	<% If len(vDataForm) = 0 Then 
        		
        					Response.Write ("<a href=" + nameofpage + "?" + "lang=esp><img src=""images/flag-esp.png""></a>")
        					Response.Write ("&nbsp;")
        					Response.Write ("<a href=" + nameofpage + "?" + "lang=eng><img src=""images/flag-eng.png""></a>")
        			Else
        					If vDataForm("lang") <> "" Then
        							
        							If vDataForm ("Accion") <> "" Then
				        					Response.Write ("<a href=" + nameofpage + "?Accion=" + request.QueryString ("Accion") + "&" + "lang=esp" + "><img src=""images/flag-esp.png""></a>")
				        					Response.Write ("&nbsp;")
				        					Response.Write ("<a href=" + nameofpage + "?Accion=" + request.QueryString ("Accion") + "&" +  "lang=eng"  + "><img src=""images/flag-eng.png""></a>")
			        				Else
			        						If vDataForm ("IDTicket") <> "" Then
			        				                Response.Write ("<a href=" + nameofpage + "?IDTicket=" + request.QueryString ("IDTicket") + "&" + "lang=esp" + "><img src=""images/flag-esp.png""></a>")
				        					        Response.Write ("&nbsp;")
				        					        Response.Write ("<a href=" + nameofpage + "?IDTicket=" + request.QueryString ("IDTicket") + "&" +  "lang=eng"  + "><img src=""images/flag-eng.png""></a>")
			        				        Else
			        				            Response.Write ("<a href=" + nameofpage + "?" + "lang=esp" + "><img src=""images/flag-esp.png""></a>")
				        					    Response.Write ("&nbsp;")
				        					    Response.Write ("<a href=" + nameofpage + "?" + "lang=eng"  + "><img src=""images/flag-eng.png""></a>")
			        				        End If
				        					    
			        				End If
			        				
	
		        					
		        					
		        						        					
	        				Else
	        						
	        						Response.Write ("<a href=" + nameofpage + "?" + request.QueryString + "&" + "lang=esp><img src=""images/flag-esp.png""></a>")
		        					Response.Write ("&nbsp;")
		        					Response.Write ("<a href=" + nameofpage + "?" + request.QueryString + "&" + "lang=eng><img src=""images/flag-eng.png""></a>")
	        				End If
	        				
        			End If
        	%>
        	
        	</td>
 </tr>
 </table>

 <!--- LOGO EMPRESA --->
 <table border="0" width="100%" cellspacing="0" cellpadding="0" style="margin:0px;BORDER-BOTTOM: gray 1px solid;">
 	<tr><td height="20px">&nbsp;</td></tr>
 <tr>
        <td style="text-align:center;height:50px;" > <a href="http://www.itdap.com"><img SRC="images/logo.gif" width="130" alt="helpdesk - ITDap Worldwide Solutions"/> </a>
        </td>
 </tr>
 </table>
 
 <!--- ESTADO DE SESION --->
 <table class="NavAreaSesion" width="100%" cellpadding="2" cellspacing="0" border="0">
 <tr><td>
  <table cellpadding="0" cellspacing="0" border="0">
  <tr><td class="NavAreaHeading" style="font: 11px Verdana, Arial, Helvetica; color: black">    
   <%=FormatDateTime(Date,vbShortDate) & " - " & FormatDateTime(Time,vbShortTime) & "Hs."%>
  </td></tr>
  <tr><td class="NavAreaHeading" style="font: 11px Verdana, Arial, Helvetica; color: white">
  
        <%If Session("Login") = 1 Then%>
                        <b><% Call ReadLang(Rs_Lenguajeopen,25) %>&nbsp;</b><%=Session("NombreUsuario")%>
        <%Else%>
           <b><% Call ReadLang(Rs_Lenguajeopen,31) %></b>
        <% 
        response.redirect( "adios.asp")
        %>
        <%End If%>

  </td></tr></table>
 </td></tr></table>
    
 <!--- BOTON DE INICIO/CIERRE DE SESION --->
 <table class="NavAreaSesion" width="100%" cellpadding="2" cellspacing="0" border="0">
 <tr><td>
  <table cellpadding="0" cellspacing="0" border="0">
  <tr><td class="NavAreaLink" style="border-color: #FFDD88;">
   <form style="margin:0" method="post" action="vbsLogin.asp" id="NavSesion" name="NavSesion">
    
    
    <input type="hidden" name="Accion" value="LOGOUT" />

        <% 
   
            
        If Session("Login") = 1 Then  
        				
						        Response.Write ("<a href=""Javascript:NavSesion.submit();"">")
						        Call ReadLang(Rs_Lenguajeopen,7) 
						        Response.Write ("</a>")
        			
        									
         Else 
            
						        Response.Write ("<a href=""Javascript:NavSesion.submit();"">")
						        Call ReadLang(Rs_Lenguajeopen,7) 
						        Response.Write ("</a>")
        						
        End If 
        
        
        If vDataForm("lang") = "esp" Then
                Response.Write "<input type=""hidden"" name=""BackToURL"" value=""adios.asp?lang=esp"" />"
        Else
                Response.Write "<input type=""hidden"" name=""BackToURL"" value=""adios.asp?lang=eng"" />"
        End If
        
        %>




   </form>
  </td></tr></table>
 </td></tr></table>

    <!--- INICIO --->
    <!--- AREA DE MENU --->
    <table class="NavAreaMenu" width="100%" cellpadding="2" cellspacing="0" border="0">
 				<tr>
 						<td>
		 					<table cellpadding="0" cellspacing="0" border="0">
		 					<tr>
		                      <td class="NavAreaHeading"> 
		                        <div align="center"><% Call ReadLang(Rs_Lenguajeopen,11) %></div>
        		                
		                      </td>
		                    </tr>
		                   </table>
		          <table cellpadding="0" cellspacing="0" border="0" width="100%">
		            <tr> 
		            	<% If vDataForm("lang") = "esp" or session ("IDIdioma") = 1 Then
	        							Response.write ("<td class=""NavAreaLink""><a href=""IngresoTickets.asp?lang=esp"">&nbsp;&nbsp;")
	        							
			        				
		        				Else
		        						Response.write ("<td class=""NavAreaLink""><a href=""IngresoTickets.asp?lang=eng"">&nbsp;&nbsp;")
		        						
		        				End If 
						        				
						     Call ReadLang(Rs_Lenguajeopen,8) %></a></td>
		            </tr>
		          </table>
		          
		          <% if session("IDGrupo") = 1 or session("IDGrupo") = 4 then 'Los asignados son solo para los admins y los tecnicos
		          			Response.write ("<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" >")
										Response.write ("            <tr>")
										
										
						          	If vDataForm = "" Then 
						          		Response.write ("              <td class=""NavAreaLink""><a href=""vertickets.asp?Accion=A"" > &nbsp;&nbsp;")
				        					
					        			Else
					        					If vDataForm("lang") = "esp" Then
					        							Response.write ("              <td class=""NavAreaLink""><a href=""vertickets.asp?Accion=A&lang=esp"" > &nbsp;&nbsp;")
							        				
						        				Else
						        						Response.write ("              <td class=""NavAreaLink""><a href=""vertickets.asp?Accion=A&lang=eng"" > &nbsp;&nbsp;")
						        				End If
					        			End If
				        		
				        		
				        		Call ReadLang(Rs_Lenguajeopen,9) 
										Response.write ("</a></td>")
										Response.write ("            </tr>")
										Response.write ("          </table>")
								End If
        	%>
        	
        	
		      
								
		          <% If Session("IDGrupo") = 1 Then
		          	
											If vDataForm = "" Then 
						          		Response.Write ("<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%""><tr><td class=""NavAreaLink"" ><A HREF=""vertickets.asp?Accion=P""> &nbsp;&nbsp;")
						          		Call ReadLang(Rs_Lenguajeopen,33) 
		          						Response.write ("</A></td></tr></table>")
				        					Response.Write ("<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%""><tr><td class=""NavAreaLink"" ><A HREF=""vertickets.asp?Accion=T""> &nbsp;&nbsp;")
					          			Call ReadLang(Rs_Lenguajeopen,10) 
					          			Response.write ("</A></td></tr></table>")
		          			
					        			Else
					        					If vDataForm("lang") = "esp" Then
					        							Response.Write ("<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%""><tr><td class=""NavAreaLink"" ><A HREF=""vertickets.asp?Accion=P&lang=esp""> &nbsp;&nbsp;")
									          		Call ReadLang(Rs_Lenguajeopen,33) 
					          						Response.write ("</A></td></tr></table>")
							        					Response.Write ("<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%""><tr><td class=""NavAreaLink"" ><A HREF=""vertickets.asp?Accion=T&lang=esp""> &nbsp;&nbsp;")
								          			Call ReadLang(Rs_Lenguajeopen,10) 
								          			Response.write ("</A></td></tr></table>")
							        				
						        				Else
						        						Response.Write ("<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%""><tr><td class=""NavAreaLink"" ><A HREF=""vertickets.asp?Accion=P&lang=eng""> &nbsp;&nbsp;")
									          		Call ReadLang(Rs_Lenguajeopen,33) 
					          						Response.write ("</A></td></tr></table>")
							        					Response.Write ("<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%""><tr><td class=""NavAreaLink"" ><A HREF=""vertickets.asp?Accion=T&lang=eng""> &nbsp;&nbsp;")
								          			Call ReadLang(Rs_Lenguajeopen,10) 
								          			Response.write ("</A></td></tr></table>")
						        				End If
					        			End If
		          		
		          Else	
		          		If vDataForm = "" Then 
		          			Response.Write ("<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%""><tr><td class=""NavAreaLink"" ><A HREF=""vertickets.asp?Accion=P""> &nbsp;&nbsp;")
		          			Call ReadLang(Rs_Lenguajeopen,33)
		          			Response.Write ("</A></td></tr></table>")
		          		Else
		          			If vDataForm("lang") = "esp" Then
		          				      			Response.Write ("<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%""><tr><td class=""NavAreaLink"" ><A HREF=""vertickets.asp?Accion=P?lang=esp""> &nbsp;&nbsp;")
									          			Call ReadLang(Rs_Lenguajeopen,33)
									          			Response.Write ("</A></td></tr></table>")
		          			Else
		          				      			Response.Write ("<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%""><tr><td class=""NavAreaLink"" ><A HREF=""vertickets.asp?Accion=P&lang=eng""> &nbsp;&nbsp;")
		          										Call ReadLang(Rs_Lenguajeopen,33)
		         											Response.Write ("</A></td></tr></table>")
		          			End If
		          		End IF
							End If %>
							
			          
							</td>
      		</tr>
      </table>  

 <!--- AREA DE USUARIO LOGEADO --->
 <table class="NavAreaMenu" width="100%" cellpadding="2" cellspacing="0" border="0">
 			<tr>
 				<td>
					 <table cellpadding="0" cellspacing="0" border="0"><tr><td class="NavAreaHeading">
								 <% Call ReadLang(Rs_Lenguajeopen,12) %>
								 </td></tr>
					 </table>
					 <table cellpadding="0" cellspacing="0" border="0"><tr><td class="NavAreaLink">
								 <% If vDataForm("lang") = "esp" or session ("IDIdioma") = 1 Then
	        							Response.write ("<a href=""modificacion_usuario_propio.asp?lang=esp""> &nbsp;")
			        				
		        				Else
		        						Response.write ("<a href=""modificacion_usuario_propio.asp?lang=eng""> &nbsp;")
		        				End If 
		        				
		        				Call ReadLang(Rs_Lenguajeopen,13) 
								Response.Write "</a>"
								
								 %>
								 
								 </td></tr>
					 </table>
 					</td>
 				</tr>
 	</table> 

<%If Session("IDGrupo") = 1 or lcase(Session("NombreUsuario")) = "demo" Then%>
 <!--- AREA DE MANTENIMIENTO --->
 <table width="100%" cellpadding="2" cellspacing="0" border="0">
 		<tr>
 				<td>
			 		<table class="NavAreaHeading" width="100%" border="0" cellspacing="0" cellpadding="0">
			      <tr>
			        <td><% Call ReadLang(Rs_Lenguajeopen,14) %></td>
			      </tr>
			    </table>
					 <table class=flyoutMenu menudata="#hpflyout" border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" width="100%" id="AutoNumber1">
				       <tr>
				         <td class=flyoutLink handle="usuarios" ><a href="#"><% Call ReadLang(Rs_LenguajeOpen,15) %></a> </td>
				       </tr>
				       <% If vDataForm("lang") = "esp" or session ("IDIdioma") = 1 Then %>
				               <tr>
				                 <td class=flyoutLink  ><a href="abm_ciudad.asp?lang=esp"><% Call ReadLang(Rs_Lenguajeopen,16) %></a> </td>
				               </tr>
				               <tr>
				                 <td class=flyoutLink  ><a href="abm_pais.asp?lang=esp"><% Call ReadLang(Rs_Lenguajeopen,17) %></a> </td>
				                 </tr>
				               <tr>
				                 <td class=flyoutLink> <a href="abm_problemas.asp?lang=esp"><% Call ReadLang(Rs_Lenguajeopen,18) %></a> </td>
				               </tr>
				               <tr>
				       <% Else %>
				       
				               <tr>
				                 <td class=flyoutLink  ><a href="abm_ciudad.asp?lang=eng"><% Call ReadLang(Rs_Lenguajeopen,16) %></a> </td>
				               </tr>
				               <tr>
				                 <td class=flyoutLink  ><a href="abm_pais.asp?lang=eng"><% Call ReadLang(Rs_Lenguajeopen,17) %></a> </td>
				                 </tr>
				               <tr>
				                 <td class=flyoutLink> <a href="abm_problemas.asp?lang=eng"><% Call ReadLang(Rs_Lenguajeopen,18) %></a> </td>
				               </tr>
				               <tr>
				       <%End If %>
				
				       </tr>
				     </table>
				   </td>
 			</tr>
</table>
 <%Else %>
<%End If %>

 </td>
  <td style="BORDER-BOTTOM: gray 1px solid" width="613"> 
 
    <!--- AREA DE CONTENIDO --->