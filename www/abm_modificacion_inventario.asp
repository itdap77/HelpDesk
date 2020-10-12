<%@ LANGUAGE=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
<%

If (Session("Login") = 1) Then
else
	Response.Redirect "Mensaje.asp?Msg=No esta logueado en el sistema.&BackToURL=default.asp"
End If

ValidUserAction "ABM", "ABM"
If request.form("ciudades") <> 0 then
if request.form("idinventario") <> 0 then
	set Rs_Pais = Server.CreateObject("ADODB.Recordset")
	Rs_Pais.ActiveConnection = MM_HelpDesk_STRING
	Rs_Pais.Source = "{call dbo.SPU_Pais_VE(" & 0 & "," & 1 & ")}"
	Rs_Pais.Open()
	
	set Rs_Inventario = Server.CreateObject("ADODB.Recordset")
	Rs_Inventario.ActiveConnection = MM_HelpDesk_STRING
	Rs_Inventario.Source = "{call SPU_inventario_xID (" & request.form("idinventario") & ")}"
	Rs_Inventario.Open()
	
	if not rs_inventario.BOF then
		
		set Rs_Inventario_V = Server.CreateObject("ADODB.Recordset")
		Rs_Inventario_V.ActiveConnection = MM_HelpDesk_STRING
		Rs_Inventario_V.Source = "{call SPU_IDInventarioXCiudad_L(" & rs_inventario("idciudad") & ")}"
		Rs_Inventario_V.Open()

	end if
	
	set Rs_Sector = Server.CreateObject("ADODB.Recordset")
	Rs_Sector.ActiveConnection = MM_HelpDesk_STRING
	Rs_Sector.Source = "{call dbo.SPU_IDSectorXCiudad_L (" 
		If request.form("idciudad") <> 0 Then		
			Rs_sector.Source = Rs_sector.Source & request.Form("idciudad") & ")}"
		else
			Rs_sector.Source = Rs_sector.Source & request.Form("ciudades") & ")}"
		end if
	Rs_Sector.Open()
	
end if
END IF

If request.form("ciudades") <> 0 then
		set Rs_InvCiu = Server.CreateObject("ADODB.Recordset")
		Rs_InvCiu.ActiveConnection = MM_HelpDesk_STRING
		Rs_InvCiu.Source = "{call dbo.SPU_IDInventarioXCiudad_L(" & Replace(Request.Form("ciudades"),"'","''") & ")}"
		Rs_InvCiu.Open()
End If


set Rs_Proveedor = Server.CreateObject("ADODB.Recordset")
Rs_Proveedor.ActiveConnection = MM_HelpDesk_STRING
Rs_Proveedor.Source = "{call dbo.SPU_Proveedores_VE}"
Rs_Proveedor.Open()


set Rs_Parte = Server.CreateObject("ADODB.Recordset")
Rs_Parte.ActiveConnection = MM_HelpDesk_STRING
Rs_Parte.Source = "{call dbo.SPU_Parte_VE}"
Rs_Parte.Open()

set Rs_Ciudad = Server.CreateObject("ADODB.Recordset")
Rs_Ciudad.ActiveConnection = MM_HelpDesk_STRING
Rs_Ciudad.Source = "{call dbo.SPU_Ciudad_VE}"
Rs_Ciudad.Open()

%>

<script language="JavaScript">

function reload()
{
frmmodinventario.submit();
}

function reloadc()
{

if (document.frmmodinventario.idinventario)
	{
	 frmmodinventario.idinventario.value = "0";
	}
 frmmodinventario.submit();
}

function reloadi()
{
	frmmodinventario.estadosector.value = "1";
	frmmodinventario.submit();
}

function reloads()//recarga sector y le asigna el cambio de idpc
{
frmmodinventario.cambioid.value = "1";
frmmodinventario.submit();

}
function reloadp()//recarga parte y le asigna el cambio de idpc
{
frmmodinventario.cambioid.value = 1;
frmmodinventario.submit();

}
function reloadci() //recarga id ciudad y le asigna el cambio de idpc
{
frmmodinventario.cambioid.value = 1;
frmmodinventario.submit();

}

function submitfrm()
{
	iErr = 0;
	if (frmmodinventario.marca.value == "")
	{
		alert ('El campo marca es obligatorio.!!');
		iErr = 1;
	}
	if (frmmodinventario.modelo.value == "")
	{
		alert ('El campo modelo es obligatorio.!!');
		iErr = 1;
	}
	if (frmmodinventario.idparte.value == "0")
	{
		alert ('El campo parte es obligatorio.!!');
		iErr = 1;
	}
	if (frmmodinventario.idciudad.value == "0")
	{
		alert ('El campo ciudad es obligatorio.!!');
		iErr = 1;
	}
	if (frmmodinventario.idsector.value == "0")
	{
		alert ('El campo sector es obligatorio.!!');
		iErr = 1;
	}
	
		if (iErr == 0)
	{	
		
		frmmodinventario.action="vbsinventario.asp";
		frmmodinventario.submit();
	}
}

</script>

    <table class="TopMenuArea" id="PageBorder" border="0" cellspacing="0" cellpadding="0" >
      <tr> 
        <td class="Titulo">Modificacion de Inventario<%'=Request.Form("cambioid")%></td>
      </tr>
    </table>

          <form name="frmmodinventario" method="POST" action="">

<input name="accion" value="M" type="hidden"> <!-- Variable de accion para agregar, modificar, o borrar segun vbsinventario.asp-->
<input type="hidden" name="estadosector" value="0">

<table class="ContentArea" border="0" cellspacing="0" cellpadding="0" height=100% width=100% >
              <tr> 
                <td width="100%"> 
                  <table width="100%" height="8%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td>&nbsp;</td>
                    </tr>
				  </table>

              <table width="100%" border="0" class="FormBoxHeader">
                <tr>
                  <td width="5">
                    <p><img src="images/gripgray.gif" width="10" height="13"></p>
                  </td>
                  <td>Seleccione el inventario a modificar.</td>
                </tr>
              </table>
              <table width="100%" border="0" cellpadding="2" cellspacing="0" class="FormBoxBody">
                <tr>
                  <td width="50%">&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                  <td><div align="right">Seleccione la Ciudad:</div>
                  </td>
                  <td width="70%">
                    <select name="ciudades" class="FormTable" onchange="javascript:reloadc();">
                      <option value="0">----</option>
                    <%
						While (NOT Rs_Ciudad.EOF)
					%>
							<option value="<%=(Rs_Ciudad.Fields.Item("idciudad").Value)%>"
						<%	If (CStr(request.form("ciudades")) = CStr(Rs_Ciudad.Fields.Item("idciudad").Value)) then 
								Response.Write("SELECTED") 
						    End If
						    Response.Write("")%>><%=(Rs_Ciudad.Fields.Item("Ciudad").Value)
						%>	</option>
						<%
							Rs_Ciudad.MoveNext()
						Wend
				
					%>

                    </select>
                  </td>
                </tr>
                <tr>
                  <td><div align="right">Seleccione el Equipo:</div>
                  </td>
                  <td width="70%">
                    <select name="idinventario" class="FormTable" onchange="javascript:reloadi();">
                      <option value="0">----</option>
                    <%
					If Request.Form("ciudades")<>0  Then
						While (NOT Rs_InvCiu.EOF)
						%>
                         <option value="<%=(Rs_InvCiu.Fields.Item("IDInventario").Value)%>"
						<%If (CStr(request.form("Idinventario")) = CStr(Rs_InvCiu.Fields.Item("IDInventario").Value)) then 
							  Response.Write("SELECTED")
							 vIdpc = Rs_InvCiu.Fields.Item("IDPc").Value
 						  End If	
                        Response.Write("")%>><%=(Rs_InvCiu.Fields.Item("IDPc").Value)
                        %></option>
                        <%
							Rs_InvCiu.MoveNext()
						Wend
					End If
					%>
                    </select>
                  </td>
                </tr>
                <tr>
                  <td colspan="2">&nbsp;</td>
                </tr>

              </table>
<%If request.form("ciudades") <> 0 then
 If request.form("idinventario") <> 0 then%> 
 
<input type="hidden" name="idpc" value="<%if not rs_inventario.bof then Response.write (rs_inventario("idpc"))end if%>">

<%if Request.Form ("cambioid") = 0 then %>
<input type="hidden" name="cambioid" value="0">
<% else %>
<input type="hidden" name="cambioid" value="1">  
<% end if%>					  		  
		
              <table width="100%" border="0" class="FormBoxHeader">
                <tr>
                  <td width="5">
                    <p><img src="images/gripgray.gif" width="10" height="13"></p>
                  </td>
                  <td>Ingrese la actualizacion de datos.</td>
                </tr>
              </table>
              <table border="0" width="100%" class="FormboxBody">
                <tr> 
                  <td width="100%">
                   <table border="0" width="100%" class="FormTable" style="border-collapse: collapse" bordercolor="#111111" cellpadding="2" cellspacing="0">
                    
					 <tr>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				    <td>&nbsp;</td>
                  <td >&nbsp;</td>
                </tr>
                <tr> 
                      <td width="15%">Marca(*):</td>
                      <td width="28%"> 
					  	<% 
					  		If request.form("marca") <> "" then 
								Response.Write ("<input name=""marca"" type=""text"" class=""FormTable"" size=""20"" maxlength=""30"" value=""")
									If Request.Form("idinventario") <> 0 Then
										Response.Write(Rs_Inventario.Fields.item("marca").value)
									Else
										If request.form("idinventario") <> 0 then
											Response.Write (request.Form("marca"))
										Else
										   response.write ("")
										End if
									 End If
							Response.Write (""">")
							Else 
            					Response.Write ("<input name=""marca"" type=""text"" class=""FormTable"" size=""20"" maxlength=""30"" value=""")
								If request.form("idinventario") <> 0 then
									Response.Write(Rs_Inventario.Fields.item("marca").value)
								End If	            		      
            				Response.Write (""">")
							End If
							
							 %>
                      </td>
                      <td width="15%">Modelo(*):</td>
                      <td >
                      	<% 
                      	if request.form("modelo") <> "" then
							Response.Write ("<input name=""modelo"" type=""text"" class=""FormTable"" size=""20"" maxlength=""50"" value=""")
								if request.form("idinventario")<> 0 then
									Response.Write(Rs_Inventario.Fields.item("modelo").value)
								Else
									if request.form("idinventario") <> 0 then
										Response.Write (request.Form("modelo"))
									else 
										response.write ("")
									end if
								End If	
						Response.Write (""">")
						else 
            				Response.Write ("<input name=""modelo"" type=""text"" class=""FormTable"" size=""20"" maxlength=""50"" value=""")
							if request.form("idinventario")<> 0  then
								Response.Write(Rs_Inventario.Fields.item("modelo").value)
							End If	            		      
            			Response.Write (""">")
						end if 
							%>
                      </td>
                </tr>
              </table>
                   
              <table width="100%" border="0" cellpadding="2" cellspacing="0" class="FormTable">
                    <tr> 
                      <td width="15%">Nro. Serie:</td>
                      <td width="28%"> 
					  	<%
					  	if request.form("nroserie") <> "" then
		                  Response.Write ("<input name=""nroserie"" type=""text"" class=""FormTable"" size=""20"" maxlength=""50"" value=""")
							if request.form("idinventario")<> 0 then
								Response.Write(Rs_Inventario.Fields.item("NroSerie").value)
							Else
								if request.form("idinventario") <> 0 then
								            Response.Write (request.Form("nroserie"))
								else 
									Response.write("")				
								end if
							End If
						  Response.Write (""">")
						else 
            		     Response.Write ("<input name=""nroserie"" type=""text"" class=""FormTable"" size=""20"" maxlength=""50"" value=""")
							if request.form("idinventario")<> 0  then
							 Response.Write(Rs_Inventario.Fields.item("NroSerie").value)
							End If	            		      
            		     Response.Write (""">")
						end if 
						 %>
                      </td>

                      <td width="15%">Proveedor:</td>
                      <td ><select name="idproveedor" size="1" class="FormTable">
							<option value="0">----</option>
								<%	if request.form("idproveedor") <> 0 then 'Si entra es por q ya tenia un valor en el formulario										
										While (NOT Rs_Proveedor.EOF) %>
											<option value="<%=(Rs_Proveedor.Fields.Item("IDProveedores").Value)%>" 
								<%		if (cstr(request.form("idproveedor")) = cstr(rs_proveedor.Fields.Item("IDProveedores").Value)) then Response.Write("SELECTED") : Response.Write("")%>><%=(Rs_Proveedor.Fields.Item("Proveedor").Value)%></option>
								<%		Rs_Proveedor.MoveNext()
										Wend
									Else
										While (NOT Rs_Proveedor.EOF) %>
											<option value="<%=(Rs_Proveedor.Fields.Item("IDProveedores").Value)%>" 
								<%			if (cstr(rs_proveedor.Fields.Item("IDProveedores").Value) = cstr(Rs_Inventario.Fields.Item("IDProveedor").Value)) then Response.Write("SELECTED") : Response.Write("")%>><%=(Rs_Proveedor.Fields.Item("Proveedor").Value)%></option>
								<%		Rs_Proveedor.MoveNext()
										Wend
									End if
									
								%>
                                </select>
					</td>
                    </tr>
                  </table>
                  <table width="100%" border="0" cellpadding="2" cellspacing="0" class="FormTable">
                    <tr> 

                      <td width="15%">Soporte:</td>
                      <td width="28%"> 
                 <%         
							If request.form("soporte") <> "" then 
								Response.Write ("<input name=""soporte"" type=""text"" class=""FormTable"" size=""20"" maxlength=""30"" value=""")
									If Request.Form("idinventario") <> 0 Then
										Response.Write(Rs_Inventario.Fields.item("soporte").value)
									Else
										If request.form("idinventario") <> 0 then
											Response.Write (request.Form("soporte"))
										Else
										   response.write ("")
										End if
									 End If
							Response.Write (""">")
							Else 
            					Response.Write ("<input name=""soporte"" type=""text"" class=""FormTable"" size=""20"" maxlength=""30"" value=""")
								If request.form("idinventario") <> 0 then
									Response.Write(Rs_Inventario.Fields.item("soporte").value)
								End If	            		      
            				Response.Write (""">")
							End If 
							%>
					  			
                        </td>

                      <td width="15%">Garantia:</td>
                      <td width="42%"> 
                      <%  
					  		If request.form("garantia") <> "" then 
								Response.Write ("<input name=""garantia"" type=""text"" class=""FormTable"" size=""20"" maxlength=""30"" value=""")
									If Request.Form("idinventario") <> 0 Then
										Response.Write(Rs_Inventario.Fields.item("garantia").value)
									Else
										If request.form("idinventario") <> 0 then
											Response.Write (request.Form("garantia"))
										Else
										   response.write ("")
										End if
									 End If
							Response.Write (""">")
							Else 
            					Response.Write ("<input name=""garantia"" type=""text"" class=""FormTable"" size=""20"" maxlength=""30"" value=""")
								If request.form("idinventario") <> 0 then
									Response.Write(Rs_Inventario.Fields.item("garantia").value)
								End If	            		      
            				Response.Write (""">")
							End If 
							 %>
                       
                      </td>
                    </tr>
                  </table>
                  <table width="100%" border="0" cellpadding="2" cellspacing="0" class="FormTable">
                    <tr>

                      <td>Fecha Vencimiento:</td>
                      <td>
					  			<%  
					  		If request.form("fechavencimiento") <> "" then 
								Response.Write ("<input name=""fechavencimiento"" type=""text"" class=""FormTable"" size=""20"" maxlength=""30"" onBlur=""checkdate(this)"" value=""")
									If Request.Form("idinventario") <> 0 Then
										Response.Write(Rs_Inventario.Fields.item("fechavencimiento").value)
									Else
										If request.form("idinventario") <> 0 then
											Response.Write (request.Form("fechavencimiento"))
										Else
										   response.write ("")
										End if
									 End If
							Response.Write (""">")
							Else 
            					Response.Write ("<input name=""fechavencimiento"" type=""text"" class=""FormTable"" size=""20"" maxlength=""30"" onBlur=""checkdate(this)"" value=""")
								If request.form("idinventario") <> 0 then
									Response.Write(Rs_Inventario.Fields.item("fechavencimiento").value)
								End If	            		      
            				Response.Write (""">")
							End If 
							 %>
					    </td>

                      <td>Parte(*):</td>
                      <td ><select name="idparte" size="1" class="FormTable" onchange="reloadp();">
                       		<option value="0">----</option>
							<% if Request.Form ("idparte") <> 0 then 
									rs_parte.MoveFirst()
									While (NOT Rs_Parte.EOF) 'con request.form idparte%>
										<option value="<%=(Rs_Parte.Fields.Item("IDParte").Value)%>" <% if (CStr(request.form("IDParte")) = CStr(Rs_Parte.Fields.Item("idparte").Value)) then Response.Write("SELECTED"):Response.Write("")%> >
										<%=(Rs_Parte.Fields.Item("Parte").Value)%>
										</option>
										<%  Rs_Parte.MoveNext()
									Wend
								else 'con recordset
									rs_parte.MoveFirst()
									While (NOT Rs_Parte.EOF) %>
										<option value="<%=(Rs_Parte.Fields.Item("IDParte").Value)%>" <% if (CStr(Rs_Inventario.fields.item("IDParte").value) = CStr(Rs_Parte.Fields.Item("IDParte").Value)) then Response.Write("SELECTED"):Response.Write("")%> >
										<%=(Rs_Parte.Fields.Item("Parte").Value)%>
										</option>
										<%  Rs_Parte.MoveNext()
									wend
								end if
							%>
                      	</select></td>
                    </tr>
                    <tr> 

                      <td width="15%">Ciudad(*):</td>
                      <td width="28%">
					 <select name="idciudad" class="FormTable" onchange="JavaScrip:reloadci();">
                          <option value="0">----</option>
                      <% rs_ciudad.movefirst()
							if Request.Form ("idciudad") <> 0 then
								While (NOT Rs_Ciudad.EOF)%> 
										<option value="<%=(Rs_Ciudad.Fields.Item("idciudad").Value)%>"
									<%	if (CStr(Request.Form("idciudad")) = CStr(Rs_Ciudad.Fields.Item("idciudad").Value)) then Response.Write("SELECTED"): Response.Write("") %>><%=(Rs_Ciudad.Fields.Item("Ciudad").Value)%></option>
									<%
									Rs_Ciudad.MoveNext()
								Wend
							else
								While (NOT Rs_Ciudad.EOF)%> 
										<option value="<%=(Rs_Ciudad.Fields.Item("idciudad").Value)%>"
									<%	if (CStr(Rs_inventario.fields.item("idciudad").value) = CStr(Rs_Ciudad.Fields.Item("idciudad").Value)) then Response.Write("SELECTED"): Response.Write("") %>><%=(Rs_Ciudad.Fields.Item("Ciudad").Value)%></option>
									<%
									Rs_Ciudad.MoveNext()
								Wend
							end if
						%>
						
                      </select>
                      			  		 						</td>

                      <td width="15%">Sector(*): </td>
                       <td><select name="idsector" size="1" class="FormTable" onchange="javascript:reloads();">
							<option value="0">----</option>
									
						<%
									Rs_Sector.MoveFirst()

									While (NOT Rs_Sector.EOF)%>
										<option value="<%=(Rs_Sector.Fields.Item("idsector").Value)%>" 
											<%	if cStr(Rs_Sector.Fields.Item("idsector").Value) =  Request.Form("idsector") and request.form("estadosector") = "0" then %> SELECTED
												<%  
													
												else 
													if CStr(Rs_Inventario.Fields.Item("idsector").value) = CStr(Rs_Sector.Fields.Item("idsector").Value) and request.form("estadosector") = "1"  then %> SELECTED
												<%  
													end if
												end if%>>
										<%=(Rs_Sector.Fields.Item("Sector").Value)%></option>
										<%  Rs_Sector.MoveNext()
									Wend%>
                        </select>

</td>
                    </tr>
                  </table>
                  <table width="100%" border="0" cellpadding="2" cellspacing="0" class="FormTable">
                    <tr> 

                      <td width="15%">Depende de:</td>
                      <td width="28%">                        
					  <select name="Inventario" size="1" class="FormTable" >
                          <option value="0">----</option>
                 <% 
					Rs_Inventario_V.MoveFirst()
                 
							While (NOT Rs_Inventario_V.EOF) 
								 %>
                          <option value="<%=(Rs_Inventario_V.Fields.Item("IDInventario").Value)%>"
                          <% if  Rs_Inventario_V.Fields.item("idinventario").value = Rs_inventario.Fields.item("inventario").value then
								 Response.Write ("SELECTED")
							 else
								Response.Write("")
							 end if
                          %>>
                          <%Response.Write (Rs_Inventario_V.Fields.Item("IDPc").Value)
							%>
                          </option>
                          <%
						  Rs_Inventario_V.MoveNext()
							Wend
						
						%>
                      </select>
					</td>

                      <td width="15%">&nbsp;</td>
                      <td>&nbsp;
						</td>
                    </tr>
                  </table>
                  <table width="100%" border="0" cellpadding="2" cellspacing="0" class="FormTable">
                    <tr>
                      <td>&nbsp;</td>
                    </tr>
                    <tr> 
                      <td width="100%"> 
                        <p align="center"> 
                          <input type="button" value="Modificar" name="Modificar" class="FormButton" onClick="JavaScript:submitfrm()">
                          <input type="reset" value="Limpiar" name="Limpiar" class="FormButton"></p>
  
                      </td>
                    </tr>
                 </table>
<% 
End If
End If
%>
		   </table>
          </table>
       </form> 
          <!--#INCLUDE file="library/PageClose.asp"-->

<% 'Cierre de Recordsets
if request.Form("IdPais") <> 0 then
Rs_Ciudad.Close()
end if

if request.form("idinventario") then
	Rs_Pais.close()
	Rs_Inventario.close()

	Rs_Sector.Close()
end if

If request.form("ciudades") <> 0 then

		Rs_InvCiu.Close()
End If


Rs_Proveedor.Close()

Rs_Parte.Close()

Rs_Ciudad.Close()

%>