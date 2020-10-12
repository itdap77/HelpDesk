<%@ LANGUAGE=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
<%

If (Session("Login") = 1) Then
else
	Response.Redirect "Mensaje.asp?Msg=No esta logueado en el sistema.&BackToURL=default.asp"
End If

ValidUserAction "ABM", "ABM"
	
	set Rs_Proveedor = Server.CreateObject("ADODB.Recordset")
	Rs_Proveedor.ActiveConnection = MM_HelpDesk_STRING
	Rs_Proveedor.Source = "{call dbo.SPU_Proveedores_VE}"
	Rs_Proveedor.Open()

if request.Form("idproveedor") <> 0 then
	set Rs_Proveedor_ID = Server.CreateObject("ADODB.Recordset")
	Rs_Proveedor_ID.ActiveConnection = MM_HelpDesk_STRING
	Rs_Proveedor_ID.Source = "{call dbo.SPU_Proveedores_Id(" & request.Form("idproveedor") & ")}"
	Rs_Proveedor_ID.Open()
	

	set Rs_Pais = Server.CreateObject("ADODB.Recordset")
	Rs_Pais.ActiveConnection = MM_HelpDesk_STRING
	Rs_Pais.Source = "{call dbo.SPU_Pais_VE(" & 0 & "," & 1 & ")}"
	Rs_Pais.Open()

	set Rs_Ciudad = Server.CreateObject("ADODB.Recordset")
	Rs_Ciudad.ActiveConnection = MM_HelpDesk_STRING
	Rs_Ciudad.Source = "{call dbo.SPU_Ciudad_VE}"
	Rs_Ciudad.Open()
	
	if Request.Form("idpais") then 'modificado
			set Rs_Ciudadxpais = Server.CreateObject("ADODB.Recordset")
			Rs_Ciudadxpais.ActiveConnection = MM_HelpDesk_STRING
			Rs_Ciudadxpais.Source = "{call dbo.SPU_IDCiudadXPais_L(" + Request.Form("idpais") + ")}"
			Rs_Ciudadxpais.Open()
		else
			if not Rs_Proveedor_id.EOF then
					set Rs_Ciudadxpais = Server.CreateObject("ADODB.Recordset")
					Rs_Ciudadxpais.ActiveConnection = MM_HelpDesk_STRING
					Rs_Ciudadxpais.Source = "{call dbo.SPU_IDCiudadXPais_L(" & Rs_Proveedor_id.Fields.item("idpais").value & ")}"
					Rs_Ciudadxpais.Open()
			end if
		end if
	
end if


%>

<script language="JavaScript">
function reload()
{
	frmmodificacionproveedor.submit();
}
function modificaproveedor()
{
	
	frmmodificacionproveedor.action = "vbsproveedor.asp";
	frmmodificacionproveedor.submit();
}
</script>

        <table class="TopMenuArea" id="PageBorder" border="0" cellspacing="0" cellpadding="0" >
          <tr> 
            <td class="Titulo">Modificacion de Proveedores</td>
          </tr>
        </table>

          <form name="frmmodificacionproveedor" method="POST" action="">
			<input name="accion" value="M" type="hidden">
            <table class="ContentArea" border="0" cellspacing="0" cellpadding="0" height=100% width=100% >
              <tr> 
                <td width="100%"> 
                  <table width="100%" height="8%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td>&nbsp;</td>
                    </tr>
                  </table>
                  <table border="0" width="100%" class="FormboxHeader" style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0" >
                <tr><td width="10"><img src="images/gripgray.gif">
                </td>
                <td>Seleccione un Proveedor.</td>
                </tr>
                </table>
                  <table border="0" width="100%" class="FormboxBody">
                    <tr> 
                      <td width="100%">
                       
                  <table border="0" width="100%" class="FormTable" style="border-collapse: collapse" bordercolor="#111111" cellpadding="2" cellspacing="0">
                    <tr> 
                      <td colspan="4">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td width="45%"> 
                        <div align="right">Proveedor:</div>
                      </td>
                      <td colspan="3"> 
                        <select name="idproveedor" class="FormTable" id="idproveedor" onChange="javascript:reload();">
                          <option value="0">----</option>
                          <%                  
					While (NOT Rs_Proveedor.EOF)
					%>
                          <option value="<%=(Rs_Proveedor.Fields.Item("IDProveedores").Value)%>"
                        <%
                        if (CStr(request.Form("idproveedor")) = CStr(Rs_Proveedor.Fields.Item("IDProveedores").Value)) then 
                          Response.Write("SELECTED")
                          dim vPais
	                      vPais = Rs_Proveedor.Fields.Item("IDPais").Value
                        Else 
                          Response.Write("")
                        end if
                        %>><%=(Rs_Proveedor.Fields.Item("Proveedor").Value)%></option>
                          <%
					Rs_Proveedor.MoveNext()
					Wend
					%>
                        </select>
                      </td>
                    </tr>
                    <tr> 
                      <td colspan="4">&nbsp;</td>
                    </tr>
                  </table>
					</table>
		
<% if Request.Form("idproveedor") Then %>
<table class="formboxHeader">
<tr>
<td><img src="images/gripgray.gif">
</td>
<td>Actualize los datos del proveedor que desee.
</td>
</tr>
</table>

            <table class="formboxbody">
              <tr> 
                <td colspan="4">&nbsp;</td>
              <tr> 
                <td width="15%">Nuevo Proveedor(*):</td>
                <td width="28%"> 
                  <%
						 If    (Len(Request.Form("proveedormod"))=0) Then 
  					 	    Response.Write (" <input type=""text"" name=""proveedormod"" class=""formtable"" onBlur=""MM_validateForm('proveedormod','','R');return document.MM_returnValue"" value=""")
							If  request.Form("idproveedor") <> 0 Then
								 Response.Write (Rs_Proveedor_Id.fields.item("proveedor").value ) 
							End if
  					 	    Response.Write (""">" )
						 Else
							Response.Write ("<input type=""text"" name=""proveedormod"" class=""formtable""  onBlur=""MM_validateForm('proveedormod','','R');return document.MM_returnValue"" value=""")
							If  request.Form("idproveedor") <> 0 Then
								 Response.Write (Rs_Proveedor_Id.fields.item("proveedor").value ) 
							Else																			
	 							 Response.Write Request.Form ("proveedormod") 
							End if
  					 	    Response.Write (""">" )
						End If
					  %>
                </td>
                <td width="15%">Direccion:</td>
                <td > 
                  <%
							 If    (Len(Request.Form("direccion"))=0) Then 
  							        Response.Write (" <input type=""text"" name=""direccion"" class=""formtable"" value=""")
									If  request.Form("idproveedor") <> 0 Then
										 Response.Write (Rs_Proveedor_Id.fields.item("direccion").value ) 
									End if
  							        Response.Write (""">" )
							 Else
									Response.Write ("<input type=""text"" name=""direccion"" class=""formtable""  value=""")
									If  request.Form("idproveedor") <> 0 Then
										 Response.Write (Rs_Proveedor_Id.fields.item("direccion").value ) 
									Else																			
	 									 Response.Write Request.Form ("direccion") 
									End if
  							        Response.Write (""">" )
							End If
							%>
              <tr> 
                <td width="15%">Codigo Postal:</td>
                <td width="28%"> 
                  <%
						 If    (Len(Request.Form("codigopostal"))=0) Then 
  						        Response.Write (" <input type=""text"" name=""codigopostal"" class=""formtable"" value=""")
								If  request.Form("idproveedor") <> 0 Then
									 Response.Write (Rs_Proveedor_Id.fields.item("codigopostal").value ) 
								End if
  						        Response.Write (""">" )
						 Else
								Response.Write ("<input type=""text"" name=""codigopostal"" class=""formtable""  value=""")
								If  Request.Form ("codigopostal") <> 0 Then
									 Response.Write (Rs_Proveedor_Id.fields.item("codigopostal").value ) 
								Else																			
	 								 Response.Write Request.Form ("codigopostal") 
								End if
  						        Response.Write (""">" )
						End If
						%>
                </td>
                <td width="15%">Telefono:</td>
                <td > 
                  <%
						 If    (Len(Request.Form("telefono"))=0) Then 
  						        Response.Write (" <input type=""text"" name=""telefono"" class=""formtable"" value=""")
								If  request.Form("idproveedor") <> 0 Then
									 Response.Write (Rs_Proveedor_Id.fields.item("telefono").value ) 
								End if
  						        Response.Write (""">" )
						 Else
								Response.Write ("<input type=""text"" name=""telefono"" class=""formtable""  value=""")
								If  request.Form("idproveedor") <> 0 Then
									 Response.Write (Rs_Proveedor_Id.fields.item("telefono").value ) 
								Else																			
	 								 Response.Write Request.Form ("telefono") 
								End if
  						        Response.Write (""">" )
						End If
						%>
              <tr> 
                <td width="15%">Pagina web:</td>
                <td width="28%"> 
                  <%
						 If    (Len(Request.Form("paginaweb"))=0) Then 
  					 	        Response.Write (" <input type=""text"" name=""paginaweb"" class=""formtable"" value=""")
								If  request.Form("idproveedor") <> 0 Then
									 Response.Write (Rs_Proveedor_Id.fields.item("paginaweb").value ) 
								End if
  					 	        Response.Write (""">" )
						 Else
								Response.Write ("<input type=""text"" name=""paginaweb"" class=""formtable""  value=""")
								If  Request.Form ("cp") <> 0 Then
									 Response.Write (Rs_Proveedor_Id.fields.item("paginaweb").value ) 
								Else																			
	 								 Response.Write Request.Form ("paginaweb") 
								End if
  					 	        Response.Write (""">" )
						End If
					  %>
                </td>
                <td width="15%">Interno:</td>
                <td width="42%"> 
                  <%
						 If    (Len(Request.Form("interno"))=0) Then 
  						        Response.Write (" <input type=""text"" name=""interno"" class=""formtable"" value=""")
								If  request.Form("idproveedor") <> 0 Then
									 Response.Write (Rs_Proveedor_Id.fields.item("interno").value ) 
								End if
  						        Response.Write (""">" )
						 Else
								Response.Write ("<input type=""text"" name=""interno"" class=""formtable""  value=""")
								If  request.Form("idproveedor") <> 0 Then
									 Response.Write (Rs_Proveedor_Id.fields.item("interno").value ) 
								Else																			
	 								 Response.Write Request.Form ("interno") 
								End if
  						        Response.Write (""">" )
						End If
					  %>
                </td>
              </tr>
              <tr> <td width="15%">Pais:</td>
                <td ><Select Name="idpais" class="FormTable" onBlur="MM_validateForm('idpais','','R');return document.MM_returnValue" onchange="javascript:reload();">
                         <option value="0">----</option>
<% if Request.Form ("idpais") = 0 then
                         While (NOT Rs_Pais.EOF)%>
								<option value=<%=Rs_Pais.Fields.Item("IDPais").Value%>
								<%if (CStr(Rs_proveedor_id("idpais")) = CStr(Rs_Pais.Fields.Item("IDPais").Value)) then%>
								 SELECTED
								<%end if%>>
								<%=Rs_Pais.Fields.Item("Pais").Value%></option>
									
								  <%Rs_Pais.MoveNext()
							Wend%>
<%else

                          While (NOT Rs_Pais.EOF)%>
								<option value=<%=Rs_Pais.Fields.Item("IDPais").Value%>
								<%if (CStr(Request.Form("idpais")) = CStr(Rs_Pais.Fields.Item("IDPais").Value)) then%>
								 SELECTED
								<%end if%>>
								<%=Rs_Pais.Fields.Item("Pais").Value%></option>
									
								  <%Rs_Pais.MoveNext()
							Wend
end if				
			%>

					
							
						</Select>
                </td>
                <td width="15%">Ciudad:</td>
                <td width="28%"> 
                  <select size="1" name="idciudad" class="FormTable" onBlur="MM_validateForm('idciudad','','R');return document.MM_returnValue" onchange="javascript:reload();">
                                        <option value="0">----</option>
				<% 'if Request.Form("idpais") <> 0 then
				if Request.Form ("idciudad") = 0 then
						Rs_Ciudadxpais.movefirst()
						While (NOT Rs_Ciudadxpais.EOF)%>
									<option value=<%=Rs_Ciudadxpais.Fields.Item("IDCiudad").Value%>
											<%if (CStr(Rs_proveedor_id("idciudad")) = CStr(Rs_Ciudadxpais.Fields.Item("IDCiudad").Value)) then%> 
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
											<%if CStr(Request.Form("idciudad")) = CStr(Rs_Ciudadxpais.Fields.Item("IDCiudad").Value) then%> 
												 SELECTED
											<%end if%>
											><%=Rs_Ciudadxpais.Fields.Item("Ciudad").Value%>
									</option>
							  <%Rs_Ciudadxpais.MoveNext()
						Wend
				
				end if
				%></select>
                </td>
                
              </tr>
              <tr> 
                <td width="15%">Contacto:</td>
                <td width="28%"> 
                  <%
						 If    (Len(Request.Form("contacto"))=0) Then 
  						    Response.Write (" <input type=""text"" name=""contacto"" class=""formtable"" value=""")
							If  request.Form("idproveedor") <> 0 Then
								 Response.Write (Rs_Proveedor_Id.fields.item("contacto").value ) 
							End if
  						    Response.Write (""">" )
						 Else
							Response.Write ("<input type=""text"" name=""contacto"" class=""formtable""  value=""")
							If  request.Form("idproveedor") <> 0 Then
								 Response.Write (Rs_Proveedor_Id.fields.item("contacto").value ) 
							Else																			
	 							 Response.Write Request.Form ("contacto") 
							End if
  						    Response.Write (""">" )
						End If
					  %>
                </td>
                <td width="15%" > E-Mail:</td>
                <td > 
                  <%
						 If    (Len(Request.Form("email"))=0) Then 
  						    Response.Write (" <input type=""text"" name=""email"" class=""formtable"" onBlur=""MM_validateForm('email','','NisEmail');return document.MM_returnValue"" value=""")
							If  request.Form("idproveedor") <> 0 Then
								 Response.Write (Rs_Proveedor_Id.fields.item("email").value ) 
							End if
  						    Response.Write (""">" )
						 Else
							Response.Write ("<input type=""text"" name=""email"" class=""formtable""  onBlur=""MM_validateForm('email','','NisEmail');return document.MM_returnValue"" value=""")
							If  request.Form("idproveedor") <> 0 Then
								 Response.Write (Rs_Proveedor_Id.fields.item("email").value ) 
							Else																			
	 							 Response.Write Request.Form ("email") 
							End if
  						    Response.Write (""">" )
						End If
					  %>
                </td>
              </tr>
              <tr> 
                <td colspan="4">&nbsp;</td>
              </tr>
              <tr> 
                <td width="100%" colspan="4"> 
                  <p align="center"> 
                    <input type="button" value="Modificar" name="Ingresar" class="FormButton" onClick="modificaproveedor();">
                    <input type="reset" value="Limpiar" name="Limpiar" class="FormButton">
                </td>
              </tr>
            </table>
                  
               
            <% end if%>
                  </td>
                    </tr>
                  </table> 
          </form>
       
      <!--#INCLUDE file="library/PageClose.asp"-->
