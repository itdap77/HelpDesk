<%@ LANGUAGE=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
<%
ValidSession()

If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 

dim usuario


if (vDataForm ("RO") = 1) Then
	Usuario = vDataForm("Usuario")
	
Else
	Usuario = Session("IDU")	
End If

set RS_Ciudad = Server.CreateObject("ADODB.Recordset")
RS_Ciudad.ActiveConnection = MM_HelpDesk_STRING
RS_Ciudad.Source = "{call dbo.SPU_Ciudad_VE}"
RS_Ciudad.Open()


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


set RS_Usuario = Server.CreateObject("ADODB.Recordset")
RS_Usuario.ActiveConnection = MM_HelpDesk_STRING
RS_Usuario.Source = "{call dbo.SPU_Usuarios_L_Full(" + Replace(usuario,"'","''") + ")}"
RS_Usuario.Open()


set Rs_Pais = Server.CreateObject("ADODB.Recordset")
Rs_Pais.ActiveConnection = MM_HelpDesk_STRING
Rs_Pais.Source = "{call dbo.SPU_Pais_VE(" & 0 & "," & 1 & ")}"
Rs_Pais.Open()

if Request.Form("paismod") <> 0 then
			set Rs_Ciudadxpais = Server.CreateObject("ADODB.Recordset")
			Rs_Ciudadxpais.ActiveConnection = MM_HelpDesk_STRING
			Rs_Ciudadxpais.Source = "{call dbo.SPU_IDCiudadXPais_L(" + Request.Form("paismod") + ")}"
			Rs_Ciudadxpais.Open()
else
			if not Rs_Usuario.EOF <> 0 then
					set Rs_Ciudadxpais = Server.CreateObject("ADODB.Recordset")
					Rs_Ciudadxpais.ActiveConnection = MM_HelpDesk_STRING
					Rs_Ciudadxpais.Source = "{call dbo.SPU_IDCiudadXPais_L(" & Rs_usuario.Fields.item("idpais").value & ")}"
					Rs_Ciudadxpais.Open()
			end if
end if

%>

<script type="text/javascript" >
var validresult;
    
function reload()
{
frmmodifusuario.submit();
}

function submitform() {
    validresult = MM_validateForm('nombre', '', 'R', 'apellido', '', 'R', 'password', '', 'R', 'loginmod', '', 'R', 'paismod a', '', 'R', 'ciudadmod', '', 'R', 'te', '', 'R', 'empresa', '', 'R');

    if (validresult == 0) {
        frmmodifusuario.action = "vbsmodificacionusuariopropio.asp";
        frmmodifusuario.submit();
    }
}

</script>

      <!--- DEFINICION DE IFRAMES UTILIZADOS --->

          <table width="100%" border="0" class="TopMenuArea" cellspacing="0" cellpadding="0">
            <tr> 
              <td class="Titulo"><% Call ReadLang(Rs_Lenguaje,55) %></td>
            </tr>
          </table>
          <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" class="ContentArea">
            <tr>
              <td>
                <table width="100%" border="0" class="FormBoxHeader">
                <tr>
                  <td width="5"><img src="images/gripgray.gif" width="10" height="13"></td>
                  <td><% Call ReadLang(Rs_Lenguaje,38) %></td>
                </tr>
              </table>
                <form name="frmmodifusuario" method="post" action="">
				          <%
				 		 	        Response.Write (" <input type=""hidden"" name=""IdGrupo"" value=""")
									Response.Write (RS_Usuario.fields.item("IdGrupo").value ) 
				  					Response.Write (""">" )

                                    Response.Write (" <input type=""hidden"" name=""lang"" value=""")
									Response.Write (vDataForm("lang") ) 
				  					Response.Write (""">" )
									%>
          
                  <table border="0" width="100%"  class="FormBoxBody" height="300">
                    <tr>
                      <td width="20%" colspan="4" style="color:red;text-align:center;font-weight:bold;">
                          &nbsp;</td>
                           
                    </tr>
                    <tr> 
                      <td width="13%"> 
                        <p><% Call ReadLang(Rs_Lenguaje,39) %></p>
                      </td>
                      <td width="37%">
                      								<%
																 	        Response.Write ("<input type=""text"" name=""nombre"" class=""formtable"" onBlur=""MM_validateForm('nombre','','R');return document.MM_returnValue"" value=")
						 												    	Response.Write (RS_Usuario.fields.item("Nombre").value ) 
																 	        Response.Write (">" )
																			%>
											</td>
                      <td width="13%"><% Call ReadLang(Rs_Lenguaje,40) %></td>
                      <td width="37%">
                      								<%
			  																	Response.Write (" <input type=""text"" name=""apellido"" class=""formtable"" onBlur=""MM_validateForm('apellido','','R');return document.MM_returnValue"" value=""")
																					Response.Write (RS_Usuario.fields.item("Apellido").value ) 
			  												 	        Response.Write (""">" )
																			%>
											</td>
                    </tr>
                    <tr> 
                      <td width="13%"><% Call ReadLang(Rs_Lenguaje,41) %></td>
		                  <td width="37%">
		                  								<% 	Response.Write (" <input type=""hidden"" name=""login""  maxlength=""15"" class=""formtable"" onBlur=""MM_validateForm('login','','R');return document.MM_returnValue"" readonly value=""")
	                    										Response.Write (RS_Usuario.fields.item("Login").value ) 
																 	        Response.Write (""">" )
																 	        Response.Write (" <input type=""text"" name=""loginmod""  maxlength=""15"" class=""formtable"" onBlur=""MM_validateForm('loginmod','','R');return document.MM_returnValue"" readonly value=""")
																					Response.Write (RS_Usuario.fields.item("Login").value ) 
														 	        		Response.Write (""">" )
																			%>
											</td>
                      <td width="13%"><% Call ReadLang(Rs_Lenguaje,42) %></td>
                      <td width="37%">
                      								<%
																 	        Response.Write (" <input type=""password"" name=""password""  maxlength=""50"" class=""formtable"" onBlur=""MM_validateForm('password','','R');return document.MM_returnValue"" value=""")
																					'Response.Write (RS_Usuario.fields.item("Password").value ) 
																 	        Response.Write (""">" )
																			%>
											</td>
                    </tr>
                    <tr> 
                      <td width="13%"><% Call ReadLang(Rs_Lenguaje,43) %></td>
                      <td width="37%">
                      								<%
																 	        Response.Write (" <textarea maxlength=""50"" rows=""2"" cols=""35"" name=""direccion"" class=""formtable"">")
																			 		Response.Write (RS_Usuario.fields.item("Direccion").value ) 
																 	        Response.Write ("</textarea>")
																			%>
											</td>
                      <td width="13%"><% Call ReadLang(Rs_Lenguaje,48) %></td>
                      <td width="37%"><Select Name="Paismod" class="FormTable" onchange="javascript:reload();">
												                   
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
                      <td width="13%"><% Call ReadLang(Rs_Lenguaje,44) %></td>
                      <td width="37%">
                      								<%
  																 	     Response.Write (" <input type=""text"" name=""tep""  maxlength=""30"" class=""formtable"" value=""")
																				 Response.Write (RS_Usuario.fields.item("Telefonop").value ) 
  																 	     Response.Write (""">" )
																			%>
											</td>
                      <td width="13%"><% Call ReadLang(Rs_Lenguaje,49) %></td>
                      <td width="37%">
                      	
			                      	<select size="1" name="ciudadmod" class="FormTable" onchange="javascript:reload();">
			                               
																			<% 
																			if Request.Form ("ciudadmod") = 0 then
																					Rs_Ciudadxpais.movefirst()
																					While (NOT Rs_Ciudadxpais.EOF)%>
																								<option value=<%=Rs_Ciudadxpais.Fields.Item("IDCiudad").Value%>
																										<%if (CStr(Rs_usuario("idciudad")) = CStr(Rs_Ciudadxpais.Fields.Item("IDCiudad").Value)) then%> 
																											 SELECTED
																										<% end if %>
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
                      <td width="13%"><% Call ReadLang(Rs_Lenguaje,45) %></td>
                      <td width="37%">
                                       <select size="1" name="dia" class="FormTable">
								                         <% While NOT Rs_Dia.EOF %>
																					<option value="<% =Rs_Dia.Fields.Item("Dia").Value%>"
																					<% if (CStr(DAY(Rs_Usuario.Fields.Item("FechaNacimiento").Value)) = CStr(Rs_Dia.Fields.Item("Dia").Value)) then Response.Write("SELECTED") : Response.Write("") %>><%=Rs_Dia.Fields.Item("Dia").Value %>
																					</option>
								                          <%
								                          Rs_Dia.MoveNext()
								                          Wend
								                          %>
								                        </select>
								                        
								                        <select size="1" name="mes" class="FormTable">
								                          
								                          <% While NOT Rs_Mes.EOF %>
																						<option value="<% =Rs_Mes.Fields.Item("Mes").Value%>"
																						<% if (CStr(MONTH(Rs_Usuario.Fields.Item("FechaNacimiento").Value)) = CStr(Rs_Mes.Fields.Item("Mes").Value)) then Response.Write("SELECTED") : Response.Write("") %>><%=Rs_Mes.Fields.Item("Mes").Value %>
																						</option>
								                          <%Rs_Mes.MoveNext()
								                          Wend %>
								                        </select>
								                        
								                        <select size="1" name="anio" class="FormTable">
								                          <%While NOT Rs_Anio.EOF%>
								                          <option value="<% =Rs_Anio.Fields.Item("Anio").Value%>"
																				  <% If (CStr(Year(Rs_Usuario.Fields.Item("FechaNacimiento").Value)) = CStr(Rs_Anio.Fields.Item("Anio").Value)) then Response.Write("SELECTED") : Response.Write("") %>><%=Rs_Anio.Fields.Item("Anio").Value %>
																				  </option>
								                          <%Rs_Anio.MoveNext()
								                          
								                          Wend %>
								                                                    
								                        </select> 
                      </td>
                      <td width="13%"><% Call ReadLang(Rs_Lenguaje,46) %></td>
                      <td width="37%">
                      								<%
																 	        Response.Write (" <input type=""text"" name=""cp""  maxlength=""10"" class=""formtable"" size=""10"" value=""")
																		 			Response.Write (RS_Usuario.fields.item("CodigoPostal").value ) 
															 	        	Response.Write (""">" )
																			%>
                      								
											</td>
                    </tr>
                    <tr>
                    	<td width="13%"><% Call ReadLang(Rs_Lenguaje,51) %></td>
                      <td width="37%"><%
  																 	        Response.Write (" <input type=""hidden"" name=""legajo"" class=""formtable"" value=""")
																				 		Response.Write (RS_Usuario.fields.item("Legajo").value ) 
  																 	        Response.Write (""">" )
  																 	        Response.Write (" <input type=""text"" name=""empresa""  maxlength=""50"" class=""formtable"" value=""")
																				 		Response.Write (RS_Usuario.fields.item("empresa").value ) 
  																 	        Response.Write (""">" )
																			%>
																			
                      <td width="13%"><% Call ReadLang(Rs_Lenguaje,47) %></td>
                      <td width="37%"><%
  																 	     Response.Write (" <input type=""text"" name=""mail"" size=""30""  maxlength=""50"" class=""formtable"" onBlur=""MM_validateForm('mail','','RisEmail');return document.MM_returnValue"" value=""")
																				 Response.Write (RS_Usuario.fields.item("Mail").value ) 
  																 	     Response.Write (""">" )
																			%>
                      		
											</td>
											
                      
                      	
                      	<!-- <select size="1" name="sectormod" class="FormTable">
                                        <option value="0">----</option>
																				 RS_SECTOR.MoveFIRST()
																							'if Request.Form ("sectormod") <> 0 Then
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
																							'End if
																				
																				
																			</select>
																			-->
											</td>
                    </tr>
                    <tr> 
                      <td width="13%"><% Call ReadLang(Rs_Lenguaje,90) %></td>
                      <td width="37%"><%
  																 	        Response.Write (" <input type=""text"" name=""te""  maxlength=""30"" class=""formtable"" value=""")
																				 		Response.Write (RS_Usuario.fields.item("Telefono").value ) 
  																 	        Response.Write (""">" )
																%></td>
                      <td width="13%"><% Call ReadLang(Rs_Lenguaje,52) %></td>
                      <td width="37%"><%
  																 	        Response.Write (" <input type=""text"" name=""interno""  maxlength=""10"" class=""formtable"" value=""")
																				 		Response.Write (RS_Usuario.fields.item("Interno").value ) 
  																 	        Response.Write (""">" )
																%></td>
                    </tr>
                    <tr> 
                      <td width="13%"></td>
                      <td width="37%">
                      									<% Response.Write (" <input type=""hidden"" name=""IDSector"" class=""formtable"" value=""")
																					 Response.Write ("NULL" ) 
																					 Response.Write (""">" )
									  										%>
											</td>
                      <td width="13%">
									                      <% Response.Write (" <input type=""hidden"" name=""Estado"" class=""formtable"" value=""")
																					 Response.Write (Rs_Usuario.fields.item("Estado").value ) 
																					 Response.Write (""">" )
									  										%>
				  						</td>
                      <td width="37%">&nbsp;</td>
                    </tr>
                    <td width="20%" colspan="4" style="color:red;text-align:center;font-weight:bold;">
                          <%If lcase(Session("NombreUsuario")) = "demo" Then 'Run only if is the demo account %>
                          <% Call ReadLang(Rs_Lenguaje,201) %>
                           <%End If%>
                        </td>
                    </tr>
                    <tr> 
                      <td colspan="4"> 
                  									<% If vDataForm ("RO") <> 1 then %>
																						<div align="center"> 
																						  <p align="center"> 
							<%If lcase(Session("NombreUsuario")) <> "demo" Then 'Run only if is the demo account %>															  	
																						  		<% 	Response.Write ("<input type=""button"" name=""Button"" value=""")
																                    	Call ReadLang(Rs_Lenguaje,53) 
																                    	Response.Write (""" class=""formbutton"" onclick=""javascript: submitform();"">")
														                    	%>&nbsp;
														                    	<% 	Response.Write ("<input type=""reset"" name=""Limpiar"" value=""")
																                    	Call ReadLang(Rs_Lenguaje,54) 
																                    	Response.Write (""" class=""formbutton"" >")
														                    	%>
																						</div>
																		<% End If %>
                          <%End If %>
                      </td>
                    </tr>
                  </table>
           
          </form></td>
            </tr>
      </table>
        <!--#INCLUDE file="library/PageClose.asp"-->
<%
RS_Ciudad.Close()
RS_Usuario.Close()
Rs_Pais.Close()
if Request.Form ("ciudadmod") then
	Rs_Sector.Close()
end if
%>