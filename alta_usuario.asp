<%@ LANGUAGE=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
<%

ValidSession()


'ValidUserAction "ABM", "ABM"

If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 


set RsDia = Server.CreateObject("ADODB.Recordset")
RsDia.ActiveConnection = MM_HelpDesk_STRING
RsDia.Source = "{call dbo.SPU_Dia_V}"
RsDia.Open()

set RsMes = Server.CreateObject("ADODB.Recordset")
RsMes.ActiveConnection = MM_HelpDesk_STRING
RsMes.Source = "{call dbo.SPU_Mes_V}"
RsMes.Open()

set RsAnio = Server.CreateObject("ADODB.Recordset")
RsAnio.ActiveConnection = MM_HelpDesk_STRING
RsAnio.Source = "{call dbo.SPU_Anio_V}"
RsAnio.Open()

set RS_Grupo = Server.CreateObject("ADODB.Recordset")
RS_Grupo.ActiveConnection = MM_HelpDesk_STRING
RS_Grupo.Source = "{call dbo.SPU_Grupo_VE (" & 0 & "," & 1 & ")}"
RS_Grupo.Open()

set RS_Pais = Server.CreateObject("ADODB.Recordset")
RS_Pais.ActiveConnection = MM_HelpDesk_STRING
RS_Pais.Source = "{call dbo.SPU_Pais_VE(" & 0 & "," & 1 & ")}"
RS_Pais.Open()

If Request.Form ("Pais") <> 0 Then
	set RS_CiudadxPais = Server.CreateObject("ADODB.Recordset")
	RS_CiudadxPais.ActiveConnection = MM_HelpDesk_STRING
	RS_CiudadxPais.Source = "{call dbo.SPU_IDCiudadXPais_L(" + Replace(Request.Form("Pais"),"'","''") + ")}"
	RS_CiudadxPais.Open()
End if

If Request.Form ("Ciudad") <> 0 Then
	set RS_SectorXCiudad = Server.CreateObject("ADODB.Recordset")
	RS_SectorXCiudad.ActiveConnection = MM_HelpDesk_STRING
	RS_SectorXCiudad.Source = "{call dbo.SPU_IDSectorXCiudad_L(" + Replace(Request.Form("Ciudad"),"'","''") + ")}"
	RS_SectorXCiudad.Open()
	
	set RS_IDCiudad = Server.CreateObject("ADODB.Recordset")
	RS_IDCiudad.ActiveConnection = MM_HelpDesk_STRING
	RS_IDCiudad.Source = "{call dbo.SPU_IDPCiudad_L(" + Replace(Request.Form("Ciudad"),"'","''") + ")}"
	RS_IDCiudad.Open()
End if


%>

<script type="text/javascript">

function submitform()
{
    validresult = MM_validateForm('nombre', '', 'R', 'apellido', '', 'R', 'password', '', 'R', 'login', '', 'R', 'pais a', '', 'R', 'ciudad', '', 'R', 'te', '', 'R', 'empresa', '', 'R', 'mail', '', 'R', 'telefono_laboral', '', 'R');

    if (validresult == 0) {
		frmaltausuario.action = "vbsusuario.asp" ;
		frmaltausuario.submit();
	}
}

</script>
      <!--- DEFINICION DE IFRAMES UTILIZADOS --->
          <table border="0" class="TopMenuArea" cellspacing="0" cellpadding="0">
            <tr> 
              <td class="Titulo"><% Call ReadLang(Rs_Lenguaje,77) %></td>
            </tr>
          </table>
          <form name="frmaltausuario" method="post" action="">
          <input type="hidden" name="accion" value="A">
		 
			 <table class="ContentArea" border="0" width="100%" cellspacing="0" cellpadding="0" height="100%">
              <tr> 
                <td width="100%">
								<% if vDataForm("return") = "-1" Then 
											If vDataForm("lang") = "esp" or session("ididioma") = 1 Then
												Response.Redirect ("mensaje.asp?Msg=El login proporcionado ya existe. El usuario no fue creado.&BackToURL=alta_usuario.asp")
											Else
												Response.Redirect ("mensaje.asp?Msg=Login exists. The user was not created.&backtourl=alta_usuario.asp")
											End If
											
									Else 
											If vDataForm("return") = "1" Then
													If vDataForm("lang") = "esp" or session("ididioma") = 1  Then
														Response.Redirect ("mensaje.asp?Msg=El usuario fue creado satisfactoriamente.&BackToURL=alta_usuario.asp")
													Else
														Response.Redirect ("mensaje.asp?Msg=User created succesfully.&backtourl=alta_usuario.asp")
													End If
											End If
									End If
										%>
                <table border="0" class="FormBoxHeader" >
					<tr><td width="10"><img src="images/gripgray.gif" alt="" />
					</td>
					<td><% Call ReadLang(Rs_Lenguaje,78) %>
					</tr></td>
                </table>
              <table border="0" width="100%" cellspacing="0" cellpadding="0" class="FormBoxBody" >
              <tr><td> 
                  <table border="0" width="100%" cellspacing="0" cellpadding="2" class="FormTable" >
										<tr><td>
										    </td>
										</tr>
								    <tr> 
                      <td width="14%"> 
                        <% Call ReadLang(Rs_Lenguaje,79) %>(*)
                      </td>
                      <td width="36%"> 
				                        <%
										 If    (Len(Request.Form("nombre"))=0) Then 
				  						        Response.Write (" <input type=""text"" maxlength=""20"" name=""nombre"" size=""20"" class=""formtable"" value="""  )
															Response.Write (""" class=""FormTable""> ")
										 Else
															Response.Write ("<input type=""text"" name=""nombre"" maxlength=""20"" size=""20"" class=""formtable""  value=""" )
															Response.Write Request.Form("nombre") 
															Response.Write (""" class=""FormTable""> ")
										End If
									%>
                      </td>
                      <td width="14%"><% Call ReadLang(Rs_Lenguaje,80) %>(*)</td>
                      <td width="36%"> 
		                        <%
								 If    (Len(Request.Form("apellido"))=0) Then 
		  						        	Response.Write (" <input type=""text"" name=""apellido"" maxlength=""30"" size=""20"" class=""formtable"" value="""  )
														Response.Write (""" class=""FormTable""> ")
								 Else
														Response.Write ("<input type=""text"" name=""apellido"" size=""20"" maxlength=""30"" class=""formtable""  value=""" )
														Response.Write Request.Form("apellido") 
														Response.Write (""" class=""FormTable""> ")
								End If
								%>
					             </td>
					          </tr>
                      <td width="14%"><% Call ReadLang(Rs_Lenguaje,81) %>(*)</td>
                      <td width="36%"> 
		                        <%
								 If    (Len(Request.Form("login"))=0) Then 
			  						        Response.Write (" <input type=""text"" name=""login"" size=""20"" maxlength=""15"" class=""formtable"" value="""  )
														Response.Write (""" class=""FormTable""> ")
								 Else
														Response.Write ("<input type=""text"" name=""login"" size=""20"" maxlength=""15"" class=""formtable""  value=""" )
														Response.Write Request.Form("login") 
														Response.Write (""" class=""FormTable""> ")
								End If
								%>
                      </td>
                      <td width="14%"><% Call ReadLang(Rs_Lenguaje,82) %>(*)</td>
                      <td width="36%"> 
		                        <%
								         If    (Len(Request.Form("password"))=0) Then 
		  						        	        Response.Write (" <input type=""password"" name=""password"" maxlength=""50"" size=""20"" value="""  )
														        Response.Write (""" class=""FormTable""> ")
								         Else
														        Response.Write ("<input type=""password"" name=""password"" maxlength=""50"" size=""20""  value=""" )
														        Response.Write Request.Form("password") 
														        Response.Write (""" class=""FormTable""> ")
								        End If
						        %>
                      </td>
                    </tr>
                    <tr> 
                      <td width="14%"><% Call ReadLang(Rs_Lenguaje,83) %></td>
                      <td width="36%"> 
		                        <%
								 If    (Len(Request.Form("direccion"))=0) Then 
			  						        Response.Write (" <textarea type=""text"" name=""direccion"" maxlength=""50"" rows=""2"" cols=""35""  class=""formtable"">"  )
														Response.Write ("</textarea>")
								 Else
														Response.Write ("<textarea type=""text"" name=""direccion""  rows=""2"" cols=""35"" maxlength=""50"" class=""formtable"">" )
														Response.Write Request.Form("direccion") 
														Response.Write ("</textarea>")
								End If
								%>
                      </td>
                      <td width="14%"><% Call ReadLang(Rs_Lenguaje,89) %>(*)</td>
                      <td width="36%"> 
                      	
                      	<select size="1" name="pais" class="FormTable" onchange="JavaScript:reloadpage(frmaltausuario);">
                          <option value="0">----</option>
		                        <%
							    While (NOT RS_Pais.EOF)
							    %>
							     <option value="<% =RS_Pais.Fields.Item("IDPais").Value %>"
								<% if (CStr(request.form("pais")) = CStr(RS_Pais.Fields.Item("IDPais").Value)) then Response.Write("SELECTED") : Response.Write("") %> > 
								 <% =RS_Pais.Fields.Item("Pais").Value %>
							</option>
						      <%               
									  RS_Pais.MoveNext()
										Wend
									%>
                        </select>
	                       
                    </td>
                  </tr>
                  <tr> 
                    <td width="14%"><% Call ReadLang(Rs_Lenguaje,92) %></td>
                    <td width="36%"> 
		                      <%
							 If    (Len(Request.Form("telefono_particular"))=0) Then 
		  						        Response.Write (" <input type=""text"" name=""telefono_particular"" maxlength=""30"" size=""15"" class=""formtable"" value="""  )
													Response.Write (""" class=""FormTable""> ")
							 Else
													Response.Write ("<input type=""text"" name=""telefono_particular"" maxlength=""30"" size=""15"" class=""formtable""  value=""" )
													Response.Write Request.Form("telefono_particular") 
													Response.Write (""" class=""FormTable""> ")
							End If
							%>
                    </td>
                    <td width="14%"><% Call ReadLang(Rs_Lenguaje,88) %>(*)</td>
                    <td width="36%"> <select size="1" name="ciudad" class="FormTable" onchange="JavaScript:reloadpage(frmaltausuario);">
                          
                          <% If  Request.Form ("pais") <> 0 Then
								    While (NOT RS_CiudadxPais.EOF)
								    %>
                          <option value="<%=(RS_CiudadxPais.Fields.Item("IDCiudad").Value)%>" <%if (CStr(request.form("ciudad")) = CStr(RS_CiudadxPais.Fields.Item("IDCiudad").Value)) then Response.Write("SELECTED") : Response.Write("")%> ><%=(RS_CiudadxPais.Fields.Item("Ciudad").Value)%></option>
                          <%
								      RS_CiudadxPais.MoveNext()
								    Wend
    								
					      End If
					      %>
                        </select>
		                      
                    </td>
                  </tr>
                  <tr> 
                      <td width="14%"><% Call ReadLang(Rs_Lenguaje,87) %></td>
                      <td width="36%"> 
                      
		                         <select size="1" name="dia" class="FormTable" >
		                          
		                          
		                          
		                          <% 

		                                While (NOT RsDia.EOF) %>
										    <option value="<% =RsDia.Fields.Item("Dia").Value%>"
									        <% if (CStr(request.form("dia")) = CStr(RsDia.Fields.Item("Dia").Value)) then Response.Write("SELECTED") : Response.Write("") %> > 
			                                  <% =RsDia.Fields.Item("Dia").Value %>
			                                  </option>
			                                  <% RsDia.MoveNext()
									          Wend 
									          
									            If (RsDia.EOF) Then
											        Response.Write "<option selected value=""1"">" 
											        Response.Write ReadLang(Rs_Lenguaje,174) 
											        Response.Write "</option>"
										        End if
										
										%>
			                                </select>
        			                        
        			                        
			                                <select size="1" name="mes" class="FormTable">
        			                          
			                                  <%
			                                  
										        While (NOT RsMes.EOF)
										            %>
										            <option value="<% =RsMes.Fields.Item("Mes").Value%>"
										            <% if (CStr(request.form("mes")) = CStr(RsMes.Fields.Item("Mes").Value)) then Response.Write("SELECTED") : Response.Write("") %> > 
				                                      <% =RsMes.Fields.Item("Mes").Value %>
				                                      </option>
				                                      <%               
										            RsMes.MoveNext()
										        Wend
												
												if (RsMes.EOF) Then
										            Response.Write "<option selected value=""1"">" 
										            Response.Write ReadLang(Rs_Lenguaje,175) 
										            Response.Write "</option>"
										        End if
        													
																
															  %>
								 
		                        </select>
		                        <select size="1" name="anio" class="FormTable">
		                          
		                          <%
                                    
    										
										    While (NOT RsAnio.EOF)
										    %>
										                              <option value="<% =RsAnio.Fields.Item("Anio").Value%>"
										    <% if (CStr(request.form("anio")) = CStr(RsAnio.Fields.Item("Anio").Value)) then Response.Write("SELECTED") : Response.Write("") %> > 
										                              <% =RsAnio.Fields.Item("Anio").Value %>
										                              </option>
										                              <%               
										      RsAnio.MoveNext()
										      
										      
										    Wend
										    
										     If (RsAnio.EOF) Then
										        Response.Write "<option selected value=""1900"">" 
										        Response.Write ReadLang(Rs_Lenguaje,176) 
										        Response.Write "</option>"
										    End if
										    
									      %>
                                    </select>
                        
                      </td>
                      <td width="14%"><% Call ReadLang(Rs_Lenguaje,84) %></td>
                      <td width="36%">  <%
													 If    (len(Request.Form("codigopostal")) = 0) Then 
								  						        Response.Write (" <input type=""text"" name=""codigopostal"" maxlength=""10"" size=""10"" class=""formtable"" value="""  )
																			Response.Write (""" class=""FormTable""> ")
													 Else
																			Response.Write ("<input type=""text"" name=""codigopostal"" maxlength=""10"" size=""10"" class=""formtable""  value=""" )
																			Response.Write Request.Form("codigopostal") 
																			Response.Write (""" class=""FormTable""> ")
													End If
													%>
                        
                      </td>
                    </tr>
                    <tr> 
                      <td width="14%"><% Call ReadLang(Rs_Lenguaje,91) %>
                      	</td>
                      <td width="36%"> 
                      <% If (Len(Request.Form("empresa"))=0) Then 
					          Response.Write (" <input type=""text"" name=""empresa"" maxlength=""50"" size=""30"" class=""formtable"" value="""  )
							        Response.Write (""" class=""FormTable""> ")
			                 Else
						                Response.Write ("<input type=""text"" name=""empresa"" maxlength=""50"" size=""30"" class=""formtable""  value=""" )
						                Response.Write vDataForm("empresa") 
						                Response.Write (""" class=""FormTable""> ")
			                End If
					%>
                        								
                        								<input type="hidden" value="null" name="sector" />
                      </td>
                      <td width="14%"><% Call ReadLang(Rs_Lenguaje,86) %>(*)</td>
                      <td width="36%"> 
                      	
                      	<%
													 If    (Len(Request.Form("mail"))=0) Then 
							  						        Response.Write (" <input type=""text"" name=""mail"" size=""20"" maxlength=""50"" class=""FormTable"" onBlur=""MM_validateForm('mail','','RisEmail');return document.MM_returnValue"" value="""  )
																		Response.Write (""" class=""FormTable""> ")
													 Else
																		Response.Write ("<input type=""text"" name=""mail"" size=""20"" maxlength=""50"" class=""FormTable"" onBlur=""MM_validateForm('mail','','RisEmail');return document.MM_returnValue"" value=""" )
																		Response.Write Request.Form("mail") 
																		Response.Write (""" class=""FormTable""> ")
													End If
													%>
                      	
                      	
                        
                   <!--     <select size="1" name="sector" class="FormTable" >
                          <option selected value="0">----</option>
                          
														If Request.Form("ciudad")<>0  Then
												
															While (NOT RS_SectorXCiudad.EOF)
															
														Response.Write (" <option value=""" )
														Response.Write (RS_SectorXCiudad.Fields.Item("IDSector").Value & """> ")
														Response.Write (RS_SectorXCiudad.Fields.Item("Sector").Value )
														Response.Write ("</option>")
																							
															  RS_SectorXCiudad.MoveNext()
															Wend
															If (RS_SectorXCiudad.CursorType > 0) Then
															 RS_SectorXCiudad.MoveFirst
															Else
															 RS_SectorXCiudad.Requery
															End If
														End If
														
                        </select>
                        -->
                      </td>
                    </tr>
                    <tr> 
                      <td width="14%"><% Call ReadLang(Rs_Lenguaje,85) %></td>
                      <td width="36%">
                      	 <%
												 If    (Len(Request.Form("telefono_laboral"))=0) Then 
							  						        Response.Write (" <input type=""text"" name=""telefono_laboral"" size=""15"" maxlength=""30"" class=""formtable"" value="""  )
																		Response.Write (""" class=""FormTable""> ")
												 Else
																		Response.Write ("<input type=""text"" name=""telefono_laboral"" size=""15"" maxlength=""30"" class=""formtable""  value=""" )
																		Response.Write Request.Form("telefono_laboral") 
																		Response.Write (""" class=""FormTable""> ")
												End If
												%>
												
												
                      	
                       <!-- <select size="1" name="pc" class="FormTable">
                          <option selected value="0">----</option>
											         
													If Request.Form("ciudad")<>0  Then
											
														While (NOT RS_IDCiudad.EOF)
														
													Response.Write (" <option value=""" )
													Response.Write (RS_IDCiudad.Fields.Item("IDPC").Value & """> ")
													Response.Write (RS_IDCiudad.Fields.Item("IDPC").Value )
													Response.Write ("</option>")
																						
														  RS_IDCiudad.MoveNext()
														Wend
														
													End If
													
                        </select>
                        -->
                      </td>
                      <td width="14%"><% Call ReadLang(Rs_Lenguaje,93) %></td>
                      <td width="36%"><%
												 If    (Len(Request.Form("interno"))=0) Then 
						  						        Response.Write (" <input type=""text"" name=""interno"" maxlength=""10"" size=""10"" class=""formtable"" value="""  )
																	Response.Write (""" class=""FormTable""> ")
												 Else
																	Response.Write ("<input type=""text"" name=""interno"" maxlength=""10"" size=""10"" class=""formtable""  value=""" )
																	Response.Write Request.Form("interno") 
																	Response.Write (""" class=""FormTable""> ")
												End If
												%>
                       
                      </td>
                    </tr>
                    <tr> 
                      <td width="14%"><% Call ReadLang(Rs_Lenguaje,94) %></td>
                      <td width="36%"><select size="1" name="grupo" class="FormTable">
                          
                        <%
												While (NOT RS_Grupo.EOF)
															If RS_Grupo.Fields.Item("IDGrupo").Value = 10 Then 
																
									                  Response.Write ("<option value=")
									                  Response.Write (RS_Grupo.Fields.Item("IDGrupo").Value)
									                  Response.Write (" selected>")
									                  Response.Write (RS_Grupo.Fields.Item("Grupo").Value)
									                  Response.Write ("</option>")
									              Else 
									                  Response.Write ("<option value=")
									                  Response.Write (RS_Grupo.Fields.Item("IDGrupo").Value)
									                  Response.Write (">")
									                  Response.Write (RS_Grupo.Fields.Item("Grupo").Value)
									                  Response.Write ("</option>")
						                        
						                   End If
												  RS_Grupo.MoveNext()
												Wend
												%>
                        </select>
                        
                      </td>
                      <td width="14%"></td>
                      <td width="36%"><input type="hidden" value="NULL" name="pc"/>
                        
                      </td>
                    </tr>
                    <tr> 
                      <td width="14%"></td>
                      <td width="36%"><input type="hidden" value="NULL" name="legajo"/>
                      	<!--
                        
												 If    (Len(Request.Form("legajo"))=0) Then 
						  						        Response.Write (" <input type=""text"" name=""legajo"" size=""10"" class=""formtable"" value="""  )
														Response.Write (""" class=""FormTable""> ")
												 Else
														Response.Write ("<input type=""text"" name=""legajo"" size=""10"" class=""formtable""  value=""" )
														Response.Write Request.Form("legajo") 
														Response.Write (""" class=""FormTable""> ")
												End If
												-->
                      </td>
                      <td width="14%">&nbsp;</td>
                      <td width="36%">&nbsp; </td>
                    </tr>
                    <tr> 
                       <td width="20%" colspan="4" style="color:red;text-align:center;font-weight:bold;">
                          <%If lcase(Session("NombreUsuario")) = "demo" Then 'Run only if is the demo account %>
                          <% Call ReadLang(Rs_Lenguaje,202) %>
                           <%End If%>
                        </td>
                    <tr> 
                      <td colspan="4"> 
                        
                          <p align="center"> 
                              <%If lcase(Session("NombreUsuario")) <> "demo" Then 'Run only if is the demo account %>
                          		<% 	Response.Write ("<input type=""button"" name=""Button"" value=""")
																                    	Call ReadLang(Rs_Lenguaje,117) 
																                    	Response.Write (""" class=""formbutton"" onclick=""javascript: submitform();"">")
														                    	%>&nbsp;
														                    	<% 	Response.Write ("<input type=""reset"" name=""Limpiar"" value=""")
																                    	Call ReadLang(Rs_Lenguaje,118) 
																                    	Response.Write (""" class=""formbutton"" >")
														                    	%>
                            <%end If %>
                       
                      </td>
                    </tr>
                  </table> <!-- Form table -->                          
	</td>
	</tr>
	</table>  <!-- Form Box Body -->
   </tr></td>
       </table> <!-- content area -->
          </form>


      <!--#INCLUDE file="library/PageClose.asp"-->
<%
RS_Grupo.Close()
RS_Pais.Close()
If Request.Form ("pais") <> 0 Then
   RS_CiudadxPais.Close()
End if
If Request.Form ("ciudad") <> 0 Then
   RS_SectorXCiudad.Close()
   RS_IDCiudad.Close()
End if
RsDia.Close()
RsMes.Close()
RsAnio.Close()
%>