<%@ LANGUAGE=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
<%

If (Session("Login") = 1) Then
else
	Response.Redirect "Mensaje.asp?Msg=No esta logueado en el sistema.&BackToURL=default.asp"
End If
%>
<%

ValidUserAction "ABM", "ABM"

set Rs_Pais = Server.CreateObject("ADODB.Recordset")
Rs_Pais.ActiveConnection = MM_HelpDesk_STRING
Rs_Pais.Source = "{call dbo.SPU_Pais_V}"
Rs_Pais.Open()

set Rs_Proveedor = Server.CreateObject("ADODB.Recordset")
Rs_Proveedor.ActiveConnection = MM_HelpDesk_STRING
Rs_Proveedor.Source = "{call dbo.SPU_Proveedores_VE}"
Rs_Proveedor.Open()

set Rs_Parte = Server.CreateObject("ADODB.Recordset")
Rs_Parte.ActiveConnection = MM_HelpDesk_STRING
Rs_Parte.Source = "{call dbo.SPU_Parte_V}"
Rs_Parte.Open()

if request.form ("idpais") <> 0 then
	set Rs_Ciudad = Server.CreateObject("ADODB.Recordset")
	Rs_Ciudad.ActiveConnection = MM_HelpDesk_STRING
	Rs_Ciudad.Source = "{call dbo.SPU_idciudadXPais_L (" + request.Form("idpais") + ")}"
	Rs_Ciudad.Open()
End if

If request.form("idciudad") <> 0 then
	set Rs_Sector = Server.CreateObject("ADODB.Recordset")
	Rs_Sector.ActiveConnection = MM_HelpDesk_STRING
	Rs_Sector.Source = "{call dbo.SPU_IDSectorXCiudad_L (" & request.Form("idciudad") & ")}"
	Rs_Sector.Open()

	set Rs_inventario = Server.CreateObject("ADODB.Recordset")
	Rs_inventario.ActiveConnection = MM_HelpDesk_STRING
	Rs_inventario.Source = "{call dbo.SPU_IDinventarioXCiudad_L (" & Request.Form ("idciudad") & ")}"
	Rs_inventario.Open()
End If

If request.form("IdSector") <> 0 Then
	set Rs_inventario_Id = Server.CreateObject("ADODB.Recordset")
	Rs_inventario_Id.ActiveConnection = MM_HelpDesk_STRING
	Rs_inventario_Id.Source = "{call dbo.SPU_inventario_ID (" + request.form ("idparte") + "," + request.form ("idciudad") + "," + request.form ("IDSector") + ")}"
	Rs_inventario_Id.Open()
End If


%>

<script language="JavaScript">
function submitfrm()
{
	iErr = 0;
	if (frmaltaproveedor.marca.value == "")
	{
		alert ('El campo marca es obligatorio.!!');
		iErr = 1;
	
	}
	if (frmaltaproveedor.modelo.value == "")
	{
		alert ('El campo modelo es obligatorio.!!');
		iErr = 1;
	}
	if (frmaltaproveedor.idparte.value == "0")
	{
		alert ('El campo parte es obligatorio.!!');
		iErr = 1;
	}
	if (frmaltaproveedor.idciudad.value == "0")
	{
		alert ('El campo ciudad es obligatorio.!!');
		iErr = 1;
	}
	if (frmaltaproveedor.inventario.value.length == 0)
	{
		frmaltaproveedor.inventario.value = "0";
	}
	
	
	if (iErr == 0)
	{
		frmaltaproveedor.action="vbsinventario.asp";
		frmaltaproveedor.submit();
	}
}

</script>

          <table class="TopMenuArea" id="PageBorder" border="0" cellspacing="0" cellpadding="0" >
            <tr> 
              <td class="Titulo">Alta de inventario</td>
            </tr>
          </table>
          <form name="frmaltaproveedor" method="POST" action="">
<input name="accion" value="A" type="hidden">
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
                <td>Ingrese los datos correspondientes al nuevo producto.</td>
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
					  			<% if request.form("marca") <> "" then 
		                          Response.Write ("<input name=""marca"" type=""text"" class=""FormTable"" size=""20"" maxlength=""30"" value=""")
								  Response.Write (request.Form("marca"))
								  Response.Write (""">")
							     else 
            		          	  Response.Write ("<input name=""marca"" type=""text"" class=""FormTable"" size=""20"" maxlength=""30"" value="""">")
								 end if %>
                      </td>
                      <td width="15%">Modelo(*):</td>
                      <td > 			<% if request.form("modelo") <> "" then
		                          Response.Write ("<input name=""modelo"" type=""text"" class=""FormTable"" size=""20"" maxlength=""50"" value=""")
								  Response.Write (request.Form("modelo"))
								  Response.Write (""">")
							     else 
            		          	  Response.Write ("<input name=""modelo"" type=""text"" class=""FormTable"" size=""20"" maxlength=""50"" value="""">")
								 end if %>
                    
                      </td>
                    </tr>
                  </table>
                   
              <table width="100%" border="0" cellpadding="2" cellspacing="0" class="FormTable">
                    <tr> 
                      <td width="15%">Nro. Serie:</td>
                      <td width="28%"> 
					  			<% if request.form("nroserie") <> "" then
		                          Response.Write ("<input name=""nroserie"" type=""text"" class=""FormTable"" size=""20"" maxlength=""50"" value=""")
								  Response.Write (request.Form("nroserie"))
								  Response.Write (""">")
							     else 
            		          	  Response.Write ("<input name=""nroserie"" type=""text"" class=""FormTable"" size=""20"" maxlength=""50"" value="""">")
								 end if %>
                 
                      </td>
                      <td width="15%">Proveedor:</td>
                      <td ><select name="idproveedor" size="1" class="FormTable" id="idproveedor">
                                  <option value="0">----</option>
                                  <% While (NOT Rs_Proveedor.EOF) %>
                                  <option value="<%=(Rs_Proveedor.Fields.Item("IDProveedores").Value)%>" <% if (CStr(request.form("idproveedor")) = CStr(Rs_Proveedor.Fields.Item("IDProveedores").Value)) then Response.Write("SELECTED") : Response.Write("")%> > <%=(Rs_Proveedor.Fields.Item("Proveedor").Value)%> </option>
                                  <%  Rs_Proveedor.MoveNext()
							Wend
							If (Rs_Proveedor.CursorType > 0) Then
							  Rs_Proveedor.MoveFirst
							Else
							  Rs_Proveedor.Requery
							End If
							%>
                                </select>
					</td>
                    </tr>
                  </table>
                  <table width="100%" border="0" cellpadding="2" cellspacing="0" class="FormTable">
                    <tr> 
                      <td width="15%">Soporte:</td>
                      <td width="28%"> 
					  			<% if request.form("soporte") <> "" then 
		                          Response.Write ("<input name=""soporte"" type=""text"" class=""FormTable"" size=""20"" maxlength=""20"" value=""")
								  Response.Write (request.Form("soporte"))
								  Response.Write (""">")
							     else 
            		          	  Response.Write ("<input name=""soporte"" type=""text"" class=""FormTable"" size=""20"" maxlength=""20"" value="""">")
								 end if %>
                        </td>
                      <td width="15%">Garantia:</td>
                      <td width="42%"> 
					  <% if request.form("garantia") <> "" then 
		                          Response.Write ("<input name=""garantia"" type=""text"" class=""FormTable"" size=""20"" maxlength=""20"" value=""")
								  Response.Write (request.Form("garantia"))
								  Response.Write (""">")
							     else 
            		          	  Response.Write ("<input name=""garantia"" type=""text"" class=""FormTable"" size=""20"" maxlength=""20"" value="""">")
								 end if %>
                       
                      </td>
                    </tr>
                  </table>
                  <table width="100%" border="0" cellpadding="2" cellspacing="0" class="FormTable">
                    <tr>
                      <td>Fecha Vencimiento:</td>
                      <td>
					  			<% if request.form("fechavencimiento") <> "" then 
		                          Response.Write ("<input name=""fechavencimiento"" type=""text"" class=""FormTable"" size=""20"" maxlength=""10"" onBlur=""checkdate(this)"" value=""")
								  Response.Write (request.Form("fechavencimiento"))
								  Response.Write (""">")
							     else 
            		          	  Response.Write ("<input name=""fechavencimiento"" type=""text"" class=""FormTable"" size=""20"" maxlength=""10"" onBlur=""checkdate(this)"" value="""">")
								 end if %>
					    </td>
                      <td>Parte(*):</td>
                      <td ><select name="idparte" size="1" class="FormTable">
                       		<option value="0">----</option>
							<% While (NOT Rs_Parte.EOF) %>
							<option value="<%=(Rs_Parte.Fields.Item("idparte").Value)%>" <% if (CStr(request.form("idparte")) = CStr(Rs_Parte.Fields.Item("idparte").Value)) then Response.Write("SELECTED") : Response.Write("")%> >
							<%=(Rs_Parte.Fields.Item("Parte").Value)%>
							</option>
							<%  Rs_Parte.MoveNext()
							Wend
							If (Rs_Parte.CursorType > 0) Then
							  Rs_Parte.MoveFirst
							Else
							  Rs_Parte.Requery
							End If
							%>
                      	</select></td>
                    </tr>
                    <tr> 
                      <td width="15%">Pais(*):</td>
                      <td width="28%">
					  <select name="idpais" size="1" class="FormTable" onchange="JavaScript:reloadpage(frmaltaproveedor);">
                        <option value="0">----</option>
                        <% While (NOT Rs_Pais.EOF) %>
                        <option value="<%=(Rs_Pais.Fields.Item("idpais").Value)%>" <% if (CStr(request.form("idpais")) = CStr(Rs_Pais.Fields.Item("idpais").Value)) then Response.Write("SELECTED") : Response.Write("")%> ><%=(Rs_Pais.Fields.Item("Pais").Value)%></option>
                        <% Rs_Pais.MoveNext()
						Wend
						If (Rs_Pais.CursorType > 0) Then
						  Rs_Pais.MoveFirst
						Else
						  Rs_Pais.Requery
						End If
						%>
                      </select>
						</td>
                      <td width="15%">Ciudad(*):</td>
                      <td ><select name="idciudad" id="idciudad" size="1" class="FormTable" onchange="JavaScript:reloadpage(frmaltaproveedor);">
                        <option value="0">----</option>
                        <%
						if request.Form("idpais") <> 0 then
						While (NOT Rs_Ciudad.EOF)
						%>
                        <option value="<%=(Rs_Ciudad.Fields.Item("idciudad").Value)%>" <% if (CStr(request.form("idciudad")) = CStr(Rs_Ciudad.Fields.Item("idciudad").Value)) then Response.Write("SELECTED") : Response.Write("")%> ><%=(Rs_Ciudad.Fields.Item("Ciudad").Value)%></option>
                        <%
						  Rs_Ciudad.MoveNext()
						Wend
						If (Rs_Ciudad.CursorType > 0) Then
						  Rs_Ciudad.MoveFirst
						Else
						  Rs_Ciudad.Requery
						End If
						End if
						%>
                      </select></td>
                    </tr>
                  </table>
                  <table width="100%" border="0" cellpadding="2" cellspacing="0" class="FormTable">
                    <tr> 
                      <td width="15%">Depende de:</td>
                      <td width="28%">                        
					  <select name="inventario" size="1" class="FormTable" >
                          <option value="0">----</option>
                        <% 
                        if Request.Form ("idciudad") <> 0 then
							While (NOT Rs_inventario.EOF) %>
								<option value="<%=(Rs_inventario.Fields.Item("IDinventario").Value)%>" <% if (CStr(request.form("inventario")) = CStr(Rs_inventario.Fields.Item("IdPc").Value)) then Response.Write("SELECTED") : Response.Write("")%> ><%=(Rs_inventario.Fields.Item("IDPc").Value)%></option>
								<%
								Rs_inventario.MoveNext()
							Wend
						end if
						%>
                      </select>
					</td>
                      <td width="15%" >Sector(*): </td>
                      <td > 
                        <select name="idsector" size="1" class="FormTable" onChange="JavaScript:reloadpage(frmaltaproveedor)">
                          <option value="0">----</option>
                          <% If request.form ("idciudad") <> 0 then
								While (NOT Rs_Sector.EOF)	%>
							  <option value="<%=(Rs_Sector.Fields.Item("idsector").Value)%>" <% if (CStr(request.form("idsector")) = CStr(Rs_Sector.Fields.Item("idsector").Value)) then Response.Write("SELECTED") : Response.Write("")%> ><%=(Rs_Sector.Fields.Item("Sector").Value)%></option>
						  <% Rs_Sector.MoveNext()
							Wend
							If (Rs_Sector.CursorType > 0) Then
							  Rs_Sector.MoveFirst
							Else
							  Rs_Sector.Requery
							End If
						End If
						%>
                        </select></td>
                    </tr>
                  </table>
                  <table width="100%" border="0" cellpadding="2" cellspacing="0" class="FormTable">
                    <tr>
                      <td>&nbsp;</td>
                    </tr>
                    <tr> 
                      <td width="100%"> 
                        <p align="center"> 
                          <input type="button" value="Ingresar" name="Ingresar" class="FormButton" onClick="JavaScript:submitfrm()">
                          <input type="reset" value="Limpiar" name="Limpiar" class="FormButton"></p>
                      </td>
                    </tr>
                  </table>
                  
                  </td>
                    </tr>
                  </table>

		<% 	  If request.form("idciudad") <> 0 and request.form("IDSector") <> 0 and request.form("idparte") <> 0 then %>
		<table width="70%" border="0" cellpadding="0" cellspacing="0" class="FormBoxHeader">
                    <tr>
                      <td width="5"><img src="images/gripgray.gif" width="10" height="13"></td>
                      <td>
                    <div align="center">Id de pc Recomendado</div>
                      </td>
                    </tr>
                  </table>
                  <table width="70%" border="0" cellpadding="0" cellspacing="0" class="FormBoxBody">
                    <tr>
                      <td align="center"><font size="4" ><% 
					  response.Write ("<input name=""IDPC"" type=""hidden"" value=""")
					  response.write (Rs_inventario_Id.fields.item("IDPC").value)
  					  response.Write (""">")
					  response.Write (Rs_inventario_Id.fields.item("IDPC").value) %> </font>
                      </td>
                    </tr>
                  </table>
		<%End If%>
		</td>
              </tr>
            </table>
          </form>
       
          <!--#INCLUDE file="library/PageClose.asp"-->
<%
Rs_Pais.Close()
%>
<%
if request.Form("idpais") <> 0 then
Rs_Ciudad.Close()
end if
%>