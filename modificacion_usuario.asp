<%@ LANGUAGE=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
<%
ValidSession()
If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 

'ValidUserAction "ABM", "ABM"

set Rs_Ciudad = Server.CreateObject("ADODB.Recordset")
Rs_Ciudad.ActiveConnection = MM_HelpDesk_STRING
Rs_Ciudad.Source = "{call dbo.SPU_Ciudad_VE}"
Rs_Ciudad.Open()


If Request.Form ("ciudad") <> 0 Then

	set Rs_IDUsuario = Server.CreateObject("ADODB.Recordset")
	Rs_IDUsuario.ActiveConnection = MM_HelpDesk_STRING
	Rs_IDUsuario.Source = "{call dbo.SPU_Usuario_IDXCiudad(" + Replace(Request.Form("ciudad"),"'","''") + ")}"
	Rs_IDUsuario.Open()
	
	if Request.Form("usuarios") <> 0 then
		set Rs_Dia = Server.CreateObject("ADODB.Recordset")
		Rs_Dia.ActiveConnection = MM_HelpDesk_STRING
		Rs_Dia.Source = "{call dbo.SPU_Dia_V}"
		Rs_Dia.Open()

		set Rs_Mes = Server.CreateObject("ADODB.Recordset")
		Rs_Mes.ActiveConnection = MM_HelpDesk_STRING
		Rs_Mes.Source = "{call dbo.SPU_Mes_V}"
		Rs_Mes.Open()

		set Rs_Anio = Server.CreateObject("ADODB.Recordset")
		Rs_Anio.ActiveConnection = MM_HelpDesk_STRING
		Rs_Anio.Source = "{call dbo.SPU_Anio_V}"
		Rs_Anio.Open()
		
		set Rs_Pais = Server.CreateObject("ADODB.Recordset")
		Rs_Pais.ActiveConnection = MM_HelpDesk_STRING
		Rs_Pais.Source = "{call dbo.SPU_Pais_VE(" & 0 & "," & 1 & ")}"
		Rs_Pais.Open()
		
		set RsGrupo = Server.CreateObject("ADODB.Recordset")
		RsGrupo.ActiveConnection = MM_HelpDesk_STRING
		RsGrupo.Source = "{call dbo.SPU_Grupo_VE(" & 0 & "," & 1 & ")}"
		RsGrupo.Open()
		
		set Rs_Usuario = Server.CreateObject("ADODB.Recordset")
		Rs_Usuario.ActiveConnection = MM_HelpDesk_STRING
		Rs_Usuario.Source = "{call dbo.SPU_Usuarios_L_Full(" + Replace(Request.Form("usuarios"),"'","''") + ")}"
		Rs_Usuario.Open()
		
		if Request.Form("paismod") then
			set Rs_Ciudadxpais = Server.CreateObject("ADODB.Recordset")
			Rs_Ciudadxpais.ActiveConnection = MM_HelpDesk_STRING
			Rs_Ciudadxpais.Source = "{call dbo.SPU_IDCiudadXPais_L(" + Request.Form("paismod") + ")}"
			Rs_Ciudadxpais.Open()
		else
			if not Rs_Usuario.EOF then
					set Rs_Ciudadxpais = Server.CreateObject("ADODB.Recordset")
					Rs_Ciudadxpais.ActiveConnection = MM_HelpDesk_STRING
					Rs_Ciudadxpais.Source = "{call dbo.SPU_IDCiudadXPais_L(" & Rs_usuario.Fields.item("idpais").value & ")}"
					Rs_Ciudadxpais.Open()
			end if
		end if

		
	End If
End if



%>

<script  type="text/javascript" language="JavaScript">
var lang = getQuerystring("lang");
    
function reload() {
        frmmodifusuario.submit();
    }

function reloadc()
{
if (frmmodifusuario.usuarios)
    {
    frmmodifusuario.usuarios.value = 0;
    }
frmmodifusuario.submit();
} 

function reloadu()
{
if (frmmodifusuario.Paismod)
	{
	frmmodifusuario.Paismod.value = 0;
	}
frmmodifusuario.submit();
}

function submitform() {
var langtext_m;

    if (lang == 'esp') {
        langtext_m = 'Esta seguro que quiere modificar el usuario ?';
       
    }
    else {
        langtext_m = 'Are you sure that you want to modify the user ?';
        
    }
    
    var validresult;
    validresult =   MM_validateForm('nombre', '', 'R', 'apellido', '', 'R', 'password', '', 'R', 'loginmod', '', 'R', 'paismod a', '', 'R', 'ciudadmod', '', 'R', 'te', '', 'R', 'empresa', '', 'R', 'mail', '', 'R', 'telefono_laboral', '', 'R');


    if (validresult == 0) 
   {if (confirm(langtext_m)){
     frmmodifusuario.action = "vbsusuario.asp";
     frmmodifusuario.submit();
     }
     }
}

</script>

      <!--- DEFINICION DE IFRAMES UTILIZADOS --->
<form name="frmmodifusuario" method="post" action="" >
         <input type="hidden" name="accion" value="M"/>
          <table width="100%" border="0" cellspacing="0" cellpadding="0" class="TopMenuArea">
            <tr> 
              <td class="Titulo"><% Call ReadLang(Rs_Lenguaje,95) %></td>
            </tr>
          </table>
      
            <table class="ContentArea" border="0" width="100%" height="100%" >
              <tr>       
                <td >
                
                <table  width="100%" >
                  <tr> 
                    <td>

                            <table class="FormBoxHeader" border="0" >
			            <tr><td width="5"><img src="images/gripgray.gif" alt="">
			            </td>
			            <td><% Call ReadLang(Rs_Lenguaje,96) %>
			            </td></tr>
			            </table> 
                            <table class="FormBoxBody" border="0" cellspacing="0" cellpadding="0" >
                      <tr><td>
                              <table border="0" width="100%" cellspacing="0" cellpadding="2" class="FormTable" >
                                <tr> 
                                  <td width="16%">&nbsp; </td>
                                  <td width="37%">&nbsp; </td>
                                  <td width="13%">&nbsp;</td>
                                  <td>&nbsp;</td>
                                </tr>
                                <tr> 
                                  <td width="18%"><% Call ReadLang(Rs_Lenguaje,97) %></td>
                                  <td width="37%"> 
                                    <select name="ciudad" class="FormTable" onchange="javascript:reloadc();">
                                      <option value="0">----</option>
                        		            <%                  
															            While (NOT Rs_Ciudad.EOF)
																            %>
																            <option value="<%=(Rs_Ciudad.Fields.Item("IDCiudad").Value)%>"
																            <%
																            if (CStr(request.form("ciudad")) = CStr(Rs_Ciudad.Fields.Item("IDCiudad").Value)) then 
																	            Response.Write("SELECTED")
																            Else 
																	            Response.Write("")
																            end if
																            %>><%=(Rs_Ciudad.Fields.Item("Ciudad").Value)%></option>
																            <%
																            Rs_Ciudad.MoveNext()
															            Wend
															            %>
                                  </select>
                                  </td>
                                  <td width="13%"><% Call ReadLang(Rs_Lenguaje,113) %></td>
                                  <td> 
                                    <select name="usuarios" class="FormTable" onChange="javascript:reloadu();">
                                      <option selected value="0">----</option>
																	            <% If lcase(Session("NombreUsuario")) <> "demo" Then 'Run only if different for the demo account
																	            Response.Write (Request.Form("ciudad") )
																	            If Request.Form("ciudad")<>0  Then
            																	
																		            If Request.Form("usuarios")<>0 then
																					            While (NOT Rs_IDUsuario.EOF)
																										            Response.Write  ("<option value=""" )
																										            Response.Write  (Rs_IDUsuario.Fields.Item("IDUsuario").Value)
																										            Response.Write  (""" ")
																										            if (CStr(Request.form("usuarios")) = CStr(Rs_IDUsuario.Fields.Item("IDUsuario").Value)) then 
																										            Response.Write("SELECTED") 
																										            else Response.Write("")
																										            end if
																										            Response.Write  (">")
																										            Response.Write  (Rs_IDUsuario.Fields.Item("login").Value)
																										            Response.Write  ("</option>")
																						              Rs_IDUsuario.MoveNext()
																					            Wend
																		            Else
																			            While (NOT Rs_IDUsuario.EOF)
																				            Response.Write (" <option value=""" )
																				            Response.Write (Rs_IDUsuario.Fields.Item("IDUsuario").Value & """> ")
																				            Response.Write (Rs_IDUsuario.Fields.Item("login").Value )
																				            Response.Write ("</option>")
																				            Rs_IDUsuario.MoveNext()
																			            Wend
																		            End if
            																	
																	            End If

                                                                                    End if
																	            %>
                                    </select>
                                  </td>
                                </tr>
                                <tr><td width="20%" colspan="6" style="color:red;text-align:center;font-weight:bold;">
                          <%If lcase(Session("NombreUsuario")) = "demo" Then 'Run only if is the demo account %>
                          <% Call ReadLang(Rs_Lenguaje,203) %>
                           <%End If%>
                        </td></tr>
                                </table>
                                </table>
                                
				            <% if request.form("ciudad") <> 0 and request.form("usuarios") <> 0 then %>
				            <table class="formboxheader">
						            <tr>
						            <td><img alt="" src="images/gripgray.gif" />
						            </td>
						            <td><% Call ReadLang(Rs_Lenguaje,98) %></td>
						            </tr>
				            </table>

                            <table class="formboxbody" border="0">
                                <tr> 
                                  <td width="18%"> 
                                    <p><% Call ReadLang(Rs_Lenguaje,99) %></p>
                                  </td>
                                  <td width="37%"><%
                                 
																	             If    (request.form("Ciudad")<>0) Then 
  																 	                    Response.Write (" <input type=""text"" name=""nombre"" maxlength=""20"" class=""formtable"" value=""")
																			            If  Request.Form ("usuarios") <> 0 Then
																				             Response.Write (Rs_Usuario.fields.item("Nombre").value ) 
																			            End if
  																 	                    Response.Write (""">" )
																	             Else
										  	    							            Response.Write ("<input type=""text"" name=""nombre"" maxlength=""20"" class=""formtable""  value=""")
																			            If  Request.Form ("usuarios") <> 0 Then
																				             Response.Write (Rs_Usuario.fields.item("Nombre").value ) 
																			            Else																			
	 																			             Response.Write Request.Form ("nombre") 
																			            End if
  																 	                    Response.Write (""">" )
																	            End If
            						
																            %></td>
                                  <td width="13%"><% Call ReadLang(Rs_Lenguaje,114) %></td>
                                  <td><%
																	             If    (Len(Request.Form("apellido"))=0) Then 
  																 	                    Response.Write (" <input type=""text"" name=""apellido"" maxlength=""30"" class=""formtable"" value=""")
																			            If  Request.Form ("usuarios") <> 0 Then
																				             Response.Write (Rs_Usuario.fields.item("Apellido").value ) 
																			            End if
  																 	                    Response.Write (""">" )
																	             Else
										  	    							            Response.Write ("<input type=""text"" name=""apellido"" maxlength=""30"" class=""formtable""  value=""")
																			            If  Request.Form ("usuarios") <> 0 Then
																				             Response.Write (Rs_Usuario.fields.item("Apellido").value ) 
																			            Else																			
	 																			             Response.Write Request.Form ("apellido") 
																			            End if
  																 	                    Response.Write (""">" )
																	            End If
																            %></td>
                                </tr>
                                <tr> 
                                  <td width="18%"><% Call ReadLang(Rs_Lenguaje,100) %></td>
                                  <td width="37%"><%
																	             If    (Len(Request.Form("login"))=0) Then 
  																 	                    Response.Write (" <input type=""text"" name=""loginmod"" maxlength=""15"" class=""formtable"" readonly value=""")
																			            If  Request.Form ("usuarios") <> 0 Then
																				             Response.Write (Rs_Usuario.fields.item("Login").value ) 
																			            End if
  																 	                    Response.Write (""">" )
  																 	                    Response.Write (" <input type=""hidden"" name=""login"" maxlength=""15"" class=""formtable"" readonly value=""")
																			            If  Request.Form ("usuarios") <> 0 Then
																				             Response.Write (Rs_Usuario.fields.item("Login").value ) 
																			            End if
  																 	                    Response.Write (""">" )
																	             Else
										  	    							            Response.Write ("<input type=""text"" name=""loginmod"" maxlength=""15"" class=""formtable"" readonly value=""")
																			            If  Request.Form ("usuarios") <> 0 Then
																				             Response.Write (Rs_Usuario.fields.item("Login").value ) 
																			            Else																			
	 																			             Response.Write Request.Form ("login") 
																			            End if
  																 	                    Response.Write (""">" )
  																 	                    Response.Write ("<input type=""hidden"" name=""login"" maxlength=""15"" class=""formtable"" readonly value=""")
																			            If  Request.Form ("usuarios") <> 0 Then
																				             Response.Write (Rs_Usuario.fields.item("Login").value ) 
																			            Else																			
	 																			             Response.Write Request.Form ("login") 
																			            End if
  																 	                    Response.Write (""">" )

																	            End If
																            %></td>
                                  <td width="13%"><% Call ReadLang(Rs_Lenguaje,101) %></td>
                                  <td><%
																	             If    (Len(Request.Form("password"))=0) Then 
  																 	                    Response.Write (" <input type=""password"" name=""password"" maxlength=""30"" class=""formtable"" value=""")
																			            If  Request.Form ("usuarios") <> 0 Then
																				             'Response.Write (Rs_Usuario.fields.item("Password").value ) 
																			            End if
  																 	                    Response.Write (""">" )
              																 	        
																	             Else
										  	    							            Response.Write ("<input type=""password"" name=""password"" maxlength=""30"" class=""formtable""  value=""")
																			            If  Request.Form ("usuarios") <> 0 Then
																				             Response.Write (Rs_Usuario.fields.item("Password").value ) 
																			            Else																			
	 																			             Response.Write Request.Form ("password") 
																			            End if
  																 	                    Response.Write (""">" )
																	            End If
																            %></td>
                                </tr>
                                <tr> 
                                  <td width="18%"><% Call ReadLang(Rs_Lenguaje,102) %></td>
                                  <td width="37%"><%
			 	                    Response.Write (" <textarea maxlength=""50"" rows=""2"" cols=""35"" maxlength=""50"" name=""direccion"" class=""formtable"">")
						 		            Response.Write (RS_Usuario.fields.item("Direccion").value ) 
			 	                    Response.Write ("</textarea>")  %></td>
                                  <td width="13%"><% Call ReadLang(Rs_Lenguaje,103) %></td>
                                  <td>
                      		            <Select Name="Paismod" class="FormTable" onchange="javascript:reload();">
            												                   
									                <% if Request.Form ("paismod") = 0 then
									                                While (NOT Rs_Pais.EOF)%>
																	                <option value=<%=Rs_Pais.Fields.Item("IDPais").Value%>
																	                <% if (CStr(Rs_usuario("idpais")) = CStr(Rs_Pais.Fields.Item("IDPais").Value)) then%>
																	                 SELECTED
																	                <% end if %>>
																	                <%=Rs_Pais.Fields.Item("Pais").Value %></option>
    																					
																	                  <% Rs_Pais.MoveNext()
																                Wend %>
									                <% else
    	
                                  While (NOT Rs_Pais.EOF)%>
											                <option value=<%=Rs_Pais.Fields.Item("IDPais").Value%>
											                <% if (CStr(Request.Form("paismod")) = CStr(Rs_Pais.Fields.Item("IDPais").Value)) then%>
											                 SELECTED
											                <% end if %>>
											                <%=Rs_Pais.Fields.Item("Pais").Value %>
									                </option>
									                 <% Rs_Pais.MoveNext()
											                Wend
											                end if				
									                 %>
					                </Select>
                                  	
                      	            </td>
                                </tr>
                                <tr> 
                                  <td width="18%"><% Call ReadLang(Rs_Lenguaje,104) %></td>
                                  <td width="37%"><%
												             If    (Len(Request.Form("tep"))=0) Then 
											 	                    Response.Write (" <input type=""text"" name=""tep"" class=""formtable"" maxlength=""30"" value=""")
														            If  Request.Form ("usuarios") <> 0 Then
															             Response.Write (Rs_Usuario.fields.item("Telefonop").value ) 
														            End if
											 	                    Response.Write (""">" )
												             Else
					  	    							            Response.Write ("<input type=""text"" name=""tep"" class=""formtable""  maxlength=""30"" value=""")
														            If  Request.Form ("usuarios") <> 0 Then
															             Response.Write (Rs_Usuario.fields.item("Telefonop").value ) 
														            Else																			
															             Response.Write Request.Form ("tep") 
														            End if
											 	                    Response.Write (""">" )
												            End If
																            %></td>
                                  <td width="13%"><% Call ReadLang(Rs_Lenguaje,105) %></td>
                                  <td>
                      	            <select size="1" name="ciudadmod" class="FormTable" onchange="javascript:reload();">
                                                    
								            <% 
								            if Request.Form ("ciudadmod") = 0 then
										            Rs_Ciudadxpais.movefirst()
										            While (NOT Rs_Ciudadxpais.EOF)%>
													            <option value=<%=Rs_Ciudadxpais.Fields.Item("IDCiudad").Value%>
															            <%if (CStr(Rs_usuario("idciudad")) = CStr(Rs_Ciudadxpais.Fields.Item("IDCiudad").Value)) then%> 
																             SELECTED
															            <%end if%>
															            ><%=Rs_Ciudadxpais.Fields.Item("Ciudad").Value%>
													            </option>
											              <%Rs_Ciudadxpais.MoveNext()
										            Wend
													
								            else
											
								            Rs_Ciudadxpais.movefirst()
										            While (NOT Rs_Ciudadxpais.EOF)%>
													            <option value=<%=Rs_Ciudadxpais.Fields.Item("IDCiudad").Value%>
															            <%if CStr(Request.Form("Ciudadmod")) = CStr(Rs_Ciudadxpais.Fields.Item("IDCiudad").Value) then%> 
																             SELECTED
															            <%end if%>
															            ><%=Rs_Ciudadxpais.Fields.Item("Ciudad").Value%>
													            </option>
											              <%Rs_Ciudadxpais.MoveNext()
										            Wend
											
								            end if
								            %>
								            </select>
                                  	
                                  	
                                  </td>
                                </tr>
                                <tr> 
                                  <td width="18%"><% Call ReadLang(Rs_Lenguaje,106) %></td>
                                  <td width="37%">
                                  <% If Request.Form ("usuarios") <> 0 Then %>
                                  <select size="1" name="dia" class="FormTable">
                                      <option value="0">Dia</option>
                                      
                                     <% While NOT Rs_Dia.EOF%>
													            <option value="<% =Rs_Dia.Fields.Item("Dia").Value%>"
													            <% if (cstr(DAY(Rs_Usuario.Fields.Item("FechaNacimiento").Value)) = CStr(Rs_Dia.Fields.Item("Dia").Value)) then Response.Write("SELECTED") : Response.Write("") %>><%=Rs_Dia.Fields.Item("Dia").Value %>
													            </option>
                                      <%
                                      Rs_Dia.MoveNext()
                                      Wend
                                      %>
                                    </select>
                                    <select size="1" name="mes" class="FormTable">
                                      <option value="0">Mes</option>
                                      <% While NOT Rs_Mes.EOF%>
													            <option value="<% =Rs_Mes.Fields.Item("Mes").Value%>"
													            <% if (CStr(MONTH(Rs_Usuario.Fields.Item("FechaNacimiento").Value)) = CStr(Rs_Mes.Fields.Item("Mes").Value)) then Response.Write("SELECTED") : Response.Write("") %>><%=Rs_Mes.Fields.Item("Mes").Value %>
													            </option>
                                      <%Rs_Mes.MoveNext()
                                      Wend %>
                                    </select>
                                    <select size="1" name="anio" class="FormTable">
                                      <option value="0">Año</option>

                                      <% While NOT Rs_Anio.EOF%>
		                                      <option value="<% =Rs_Anio.Fields.Item("Anio").Value%>"
														              <% If (CStr(Year(Rs_Usuario.Fields.Item("FechaNacimiento").Value)) = CStr(Rs_Anio.Fields.Item("Anio").Value)) then Response.Write("SELECTED") : Response.Write("") %>><%=Rs_Anio.Fields.Item("Anio").Value %>
														              </option>
		                                      <% Rs_Anio.MoveNext()
                                      
                          		            Wend %>
                                                                
                                      <% Else
				                                      Response.Write ("<select size=""1"" name=""dia"" class=""FormTable""> ")
				                                      Response.Write ("<option value=""0"">Dia</option></select>")
				                                      Response.Write ("<select size=""1"" name=""mes"" class=""FormTable""> ")
				                                      Response.Write ("<option value=""0"">mes</option></select>")
				                                      Response.Write ("<select size=""1"" name=""Anio"" class=""FormTable""> ")
				                                      Response.Write ("<option value=""0"">Anio</option></select>")
                                      
                                      End if %>
                                    </select>
                                  </td>
                                  <td width="13%"><% Call ReadLang(Rs_Lenguaje,107) %></td>
                                  <td>
                    				            <%
						            Response.Write (" <input type=""text"" name=""cp"" class=""formtable"" maxlength=""10"" value=""")
            						
										             If (Len(Request.Form("cp"))=0) Then 
													 	     
															            If  Request.Form ("usuarios") <> 0 Then
																             Response.Write (Rs_Usuario.fields.item("CodigoPostal").value ) 
															            End if
																		
													 	        
										             Else
						  	    															  	    							
														            If  Request.Form ("usuarios") <> 0 Then
															 	            Response.Write (Rs_Usuario.fields.item("CodigoPostal").value ) 
														            Else																			
															 	            Response.Write Request.Form ("cp") 
														            End if
															 	  
															 	  
															 	  
										            End If
													
										            Response.Write (""">" )
									            %>
                                    </td>
                                </tr>
                                <tr> 
                                  <td width="18%"><% Call ReadLang(Rs_Lenguaje,108) %></td>
                                  <td width="37%">	<%
																	             If    (Len(Request.Form("te"))=0) Then 
  																 	                    Response.Write (" <input type=""text"" name=""te"" class=""formtable"" maxlength=""20"" value=""")
																			            If  Request.Form ("usuarios") <> 0 Then
																				             Response.Write (Rs_Usuario.fields.item("Telefono").value ) 
																			            End if
  																 	                    Response.Write (""">" )
																	             Else
										  	    							            Response.Write ("<input type=""text"" name=""te"" class=""formtable""  maxlength=""20"" value=""")
																			            If  Request.Form ("usuarios") <> 0 Then
																				             Response.Write (Rs_Usuario.fields.item("Telefono").value ) 
																			            Else																			
	 																			             Response.Write Request.Form ("te") 
																			            End if
  																 	                    Response.Write (""">" )
																	            End If
																            %>
											            </td>
                                  <td width="13%"><% Call ReadLang(Rs_Lenguaje,109) %></td>
			                                  <td>	<%
																	             If    (Len(Request.Form("mail"))=0) Then 
  																 	                    Response.Write (" <input type=""text"" name=""mail"" size=""35"" maxlenght=""50"" class=""formtable"" onBlur=""MM_validateForm('mail','','NisEmail');return document.MM_returnValue"" value=""")
																			            If  Request.Form ("usuarios") <> 0 Then
																				             Response.Write (Rs_Usuario.fields.item("Mail").value ) 
																			            End if
  																 	                    Response.Write (""">" )
																	             Else
										  	    							            Response.Write ("<input type=""text"" name=""mail"" class=""formtable"" maxlength=""50"" onBlur=""MM_validateForm('mail','','NisEmail');return document.MM_returnValue"" value=""")
																			            If  Request.Form ("usuarios") <> 0 Then
																				             Response.Write (Rs_Usuario.fields.item("Mail").value ) 
																			            Else																			
	 																			             Response.Write Request.Form ("mail") 
																			            End if
  																 	                    Response.Write (""">" )
																	            End If
																            %>
            			                      	
			                      	            <!--<select size="1" name="sectormod" class="FormTable">
			                                                    <option value="0">----</option>
            			 
																		            If  Request.Form ("usuarios") <> 0 Then
																						            While (NOT RS_Sector.EOF)
            																																				
																											            Response.Write  ("<option value=""" )
																											            Response.Write  (RS_Sector.Fields.Item("IDSector").Value)
																											            Response.Write  (""" ")
																											            if (CStr(Rs_Usuario.Fields.Item("IDSector").Value) = CStr(RS_Sector.Fields.Item("IDSector").Value)) then 
																											            Response.Write("SELECTED") 
																											            else Response.Write("")
																											            end if
																											            Response.Write  (">")
																											            Response.Write  (RS_Sector.Fields.Item("Sector").Value)
																											            Response.Write  ("</option>")& vbCrLf
            																		
																							              RS_Sector.MoveNext()
																						            Wend
																		            End If
			            </select>
			            -->
			            </td>
                                </tr>
                                	
                      	            <!-- <select name="pc" size="1" class="FormTable" id="pc">
                                    <option value="0">----</option>
                                     
													            If  Request.Form ("usuarios") <> 0 Then
																	            While (NOT Rs_Inventario.EOF)
            																															
																						            Response.Write  ("<option value=""" )
																						            Response.Write  (Rs_Inventario.Fields.Item("IDInventario").Value)
																						            Response.Write  (""" ")
																						            if isnull(Rs_Usuario.Fields.Item("IDInventario").value) then
																						               Response.Write("")
																						            else
																						              if (CStr(Rs_Usuario.Fields.Item("IDInventario").Value) = CStr(Rs_Inventario.Fields.Item("IDInventario").Value)) then 
																						               Response.Write("SELECTED") 
																						               else Response.Write("")
																						               end if
																						            end if
																						            Response.Write  (">")
																						            Response.Write  (Rs_Inventario.Fields.Item("IDPc").Value)
																						            Response.Write  ("</option>")& vbCrLf
            													
																		              Rs_Inventario.MoveNext()
																	            Wend
													            End If
            														
                                  </select>
                                  -->
                                  
                                  	
                                  	
                                  
                                <tr> 
                                  <td width="18%"><% Call ReadLang(Rs_Lenguaje,110) %></td>
                                  <td width="37%"><%
												             If    (Len(Request.Form("interno"))=0) Then 
											 	                    Response.Write (" <input type=""text"" name=""interno"" maxlength=""10"" class=""formtable"" value=""")
														            If  Request.Form ("usuarios") <> 0 Then
															             Response.Write (Rs_Usuario.fields.item("Interno").value ) 
														            End if
											 	                    Response.Write (""">" )
												             Else
					  	    							            Response.Write ("<input type=""text"" name=""interno"" maxlength=""10"" class=""formtable""  value=""")
														            If  Request.Form ("usuarios") <> 0 Then
															             Response.Write (Rs_Usuario.fields.item("Interno").value ) 
														            Else																			
															             Response.Write Request.Form ("interno") 
														            End if
											 	                    Response.Write (""">" )
												            End If
																            %></td>
                                  <td width="13%"><% Call ReadLang(Rs_Lenguaje,111) %></td>
                                  <td><select size="1" name="idgrupo" class="FormTable">
                                    <% 
								            If  Request.Form ("usuarios") <> 0 Then
									            While (NOT RsGrupo.EOF)
																										
														            Response.Write  ("<option value=""" )
														            Response.Write  (RsGrupo.Fields.Item("IDGrupo").Value)
														            Response.Write  (""" ")
														            if (CStr(Rs_Usuario.Fields.Item("IDGrupo").Value) = CStr(RsGrupo.Fields.Item("IDGrupo").Value)) then 
														            Response.Write("SELECTED") 
														            else Response.Write("")
														            end if
														            Response.Write  (">")
														            Response.Write  (RsGrupo.Fields.Item("Grupo").Value)
														            Response.Write  ("</option>")& vbCrLf
								
								
										              RsGrupo.MoveNext()
									            Wend
								            End If
												            %>
												            </select>

            </td>
                                </tr>
                                <tr> 
                                  <td width="18%"><% Call ReadLang(Rs_Lenguaje,112) %></td>
                                  <td width="37%">
			                      	            <Select name="Estado" class="FormTable">
                                  <option value="<% 
							                                  If request.form ("usuarios") <> 0 Then 
							                      		            If rs_usuario.fields.item("Estado").value = 1 then 
							                      			            Response.Write "1"  
							                      			            Response.Write """" 
							                      			            Response.Write " Selected>" 
							                      			            Response.Write "Si"
							                      			            Response.Write "</option>"
							                      			            Response.Write "<option value=""0"">No"
							                    		            Else               		       			
																					            Response.Write "0"  
							                      			            Response.Write """" 
							                      			            Response.Write " Selected>" 
							                      			            Response.Write "No"	
							                      			            Response.Write "</option>"
							                      			            Response.Write "<option value=""1"">Si"
            							                      												  
							                      		            End If
							                                  Else
							                                    Response.Write """"		
							               					            Response.Write "</option>"
							                                  End If
												 				            %>">
                                  </option>
                                  </select>
			                      	            <!--
																	             If    (Len(Request.Form("legajo"))=0) Then 
  																 	                    Response.Write (" <input type=""text"" name=""legajo"" class=""formtable"" value=""")
																			            If  Request.Form ("usuarios") <> 0 Then
																				             Response.Write (Rs_Usuario.fields.item("Legajo").value ) 
																			            End if
  																 	                    Response.Write (""">" )
																	             Else
										  	    							            Response.Write ("<input type=""text"" name=""legajo"" class=""formtable""  value=""")
																			            If  Request.Form ("usuarios") <> 0 Then
																				             Response.Write (Rs_Usuario.fields.item("Legajo").value ) 
																			            Else																			
	 																			             Response.Write Request.Form ("legajo") 
																			            End if
  																 	                    Response.Write (""">" )
																	            End If
																            -->
            																
																            </td>
                                  <td width="13%">&nbsp;</td>
                                  <td>
                      	            <input name="sector" type="hidden" value="NULL">
			                      	            <input name="idinventario" type="hidden" value="NULL" />
			                      	            <input name="legajo" type="hidden" value="NULL" />
            			                      	
            			                      	
                                  </td>
                                </tr>
                                                    <tr>
                                  <td colspan="4">&nbsp;</td>
                                </tr>
                                <tr> 
                                  <td colspan="4"> 
                                      <p align="center"> 
                          	            <% 	Response.Write ("<input type=""button"" name=""Button"" value=""")
				                    	            Call ReadLang(Rs_Lenguaje,115) 
				                    	            Response.Write (""" class=""formbutton"" onclick=""javascript: submitform();"">")
		                    	            %>&nbsp;
		                    	            <% 	Response.Write ("<input type=""reset"" name=""Limpiar"" value=""")
				                    	            Call ReadLang(Rs_Lenguaje,116) 
				                    	            Response.Write (""" class=""formbutton"" >")
		                    	            %>
            														      
                                            </td>
                                        
                                        <tr>
                                        <td colspan="4">&nbsp;
                                        </td>
                                        </tr>
                                      </table>
                             <% 
                             else 'If no values selected return an empty table to avoid footer width go out boundaries.
                             Response.Write("<table width='100%'><tr><td>&nbsp;</td></tr></table>")
                             End If
                              %>
                    </td>
                    </tr>
                </table><!-- closging inner contentarea table -->
            </td>         
        </tr>
        </table> <!-- closging contentarea table -->
      
      
</form>  
      <!--#INCLUDE file="library/PageClose.asp"-->
<%
Rs_Ciudad.Close()

If Request.Form ("ciudad") <> 0 Then
	Rs_IDUsuario.Close()
	If Request.Form ("usuarios") <> 0 Then
		Rs_Usuario.Close()
		Rs_anio.Close()
		RS_mes.Close()
		Rs_dia.Close()
		Rs_Pais.Close()
		RsGrupo.Close()
	End if
End if

if Request.Form("inventario") then
	RS_Sector.Close()
end if


%>