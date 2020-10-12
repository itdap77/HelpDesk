<%@ LANGUAGE=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
<%

If (Session("Login") = 1) Then
else
	Response.Redirect "Mensaje.asp?Msg=No esta logueado en el sistema.&BackToURL=default.asp"
End If

ValidUserAction "ABM", "ABM"

set Rs_ciudad = Server.CreateObject("ADODB.Recordset")
Rs_ciudad.ActiveConnection = MM_HelpDesk_STRING
Rs_ciudad.Source = "{call dbo.SPU_Ciudad_VE}"
Rs_ciudad.Open()


If Request.Form ("ciudadesbaja") <> 0 Then
	set Rs_Sectorbaja = Server.CreateObject("ADODB.Recordset")
	Rs_Sectorbaja.ActiveConnection = MM_HelpDesk_STRING
	Rs_Sectorbaja.Source = "{call dbo.SPU_IDSectorXCiudad_L(" + Replace(Request.Form("ciudadesbaja"),"'","''") + ")}"
	Rs_Sectorbaja.Open()
End if

If Request.Form ("ciudadesmod") <> 0 Then
	set Rs_Sectormod = Server.CreateObject("ADODB.Recordset")
	Rs_Sectormod.ActiveConnection = MM_HelpDesk_STRING
	Rs_Sectormod.Source = "{call dbo.SPU_IDSectorXCiudad_L(" + Replace(Request.Form("ciudadesmod"),"'","''") + ")}"
	Rs_Sectormod.Open()
End if

%>

<script languague="javascript" language="JavaScript">
function submitfrmalta()
{
altasector.action = "vbssector.asp" ;
altasector.submit();
}

function reload()
{
//alert (bajasector.ciudadesbaja.value);
bajasector.submit();
}

function reloadc()
{

bajasector.submit();
}

function reload2()
{
modificacionsector.submit();
}

function submitfrmbaja()
{
	if (confirm("Esta Seguro que desea borrar el Sector Seleccionado?"))
	{
		bajasector.action = "vbssector.asp";
		bajasector.submit();
	}
}

function submitfrmmodif()
{
modificacionsector.action = "vbssector.asp";
modificacionsector.submit();
}
</script>

      <div align="center"> 
        <center>

          <table class="TopMenuArea" id="PageBorder" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td class="Titulo">ABM de Sectores</td>
            </tr>
          </table>

          <table class="ContentArea" border="0" width="100%" cellspacing="0" cellpadding="0" height="100%">
            <tr> 
              <td width="100%">
<table height="8%" border="0" cellspacing="0" cellpadding="0">
<tr><td>&nbsp;
</td></tr>
</table>

                        <form name="altasector" method="POST" action="">
<input name="accion" type="hidden" id="accion" value="A">
                          <table width="100%" border="0" class="FormBoxHeader">
                            <tr>
                              <td width="1"><img src="images/gripgray.gif" width="10" height="13"></td>
                              
                      <td>Alta de Sectores</td>
                            </tr>
                          </table>
                          
                  <table width="100%" border="0" cellpadding="0" cellspacing="2" class="FormBoxBody">
                    <tr> 
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr> 
                      <td> &nbsp;&nbsp;&nbsp;Nombre del nuevo Sector(*):</td>
                      <td> 
                        <input type="text" name="newsectoralta" size="20" class="FormTable" maxlength="20">
                      </td>
                      <td>A que Ciudad Pertenece(*):</td>
                      <td> 
                        <select size="1" name="ciudadalta" class="FormTable">
                          <option value="0">----</option>
                          <%While (NOT Rs_ciudad.EOF)%>
                          <option value="<%=(Rs_ciudad.Fields.Item("IDCiudad").Value)%>"><%=(Rs_ciudad.Fields.Item("Ciudad").Value)%></option>
                          <%Rs_ciudad.MoveNext()
							Wend
						  %>
                        </select>
                      </td>
                    </tr>
                    <tr> 
                      <td>&nbsp;&nbsp;&nbsp;Sigla(*):</td>
                      <td colspan="3"> 
                        <input type="text" name="sigla" size="3" class="FormTable" maxlength="2">
                      </td>
                    </tr>
                    <tr> 
                      <td height="23" colspan="4">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td height="23" colspan="4"> 
                        <div align="center"> 
                          <input type="button" value="Ingresar" name="Ingresar" class="FormButton" onClick="javascript: submitfrmalta();">
                          <input type="reset" name="Limpiar2" value="Limpiar" class="FormButton">
                        </div>
                      </td>
                    </tr>
                  </table>
                  </form>
     <!----------------- B A J A   D E   S E C T O R   -------------->
                        <form name="bajasector" action="" method="POST">
							<input name="accion" type="hidden" id="accion" value="B">
							<table height="1%" border="0" cellspacing="0" cellpadding="0">
								<tr><td>&nbsp;
								</td></tr>
							</table>
                  <table width="100%" border="0" class="FormBoxHeader">
                    <tr>
                      <td width="5"><img src="images/gripgray.gif" width="10" height="13"></td>
                      <td> 
                        Baja
                      </td>
                    </tr>
                  </table>
                          
                  <table width="100%" border="0" class="FormBoxBody" cellspacing="2">
                    <tr> 
                      <td width="24%">&nbsp;</td>
                      <td width="25%">&nbsp;</td>
                      <td width="26%">&nbsp;</td>
                      <td width="25%">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td width="24%">&nbsp;&nbsp;&nbsp;Seleccione la Ciudad:</td>
                      <td width="25%"> 
                        <select name="ciudadesbaja" class="FormTable" onChange="javascript: reload();">
                          <option value="0">----</option>
							<%
							Rs_Ciudad.MoveFirst()
							While (NOT Rs_ciudad.EOF)
							%>
							<option value="<% =Rs_ciudad.Fields.Item("IDCiudad").Value%>"
							<% if (CStr(request.form("ciudadesbaja")) = CStr(Rs_ciudad.Fields.Item("IDCiudad").Value)) then Response.Write("SELECTED") : Response.Write("") %> > 
							<% =Rs_ciudad.Fields.Item("Ciudad").Value %>
							</option>
							<%             
							Rs_ciudad.MoveNext()
							Wend
							Rs_ciudad.MoveFirst()
							%>
                        </select>
                      </td>
                      <td width="26%">Seleccione el Sector:</td>
                      <td width="25%"> 
                        <select name="sectoresbaja" class="FormTable">
                          <option value="0">----</option>
                          <%
							If Request.Form("ciudadesbaja") <> 0  Then

								While (NOT Rs_Sectorbaja.EOF)
										
									Response.Write (" <option value=""" )
									Response.Write (Rs_Sectorbaja.Fields.Item("IDSector").Value & """> ")
									Response.Write (Rs_Sectorbaja.Fields.Item("Sector").Value )
									Response.Write ("</option>")
																		
								 Rs_Sectorbaja.MoveNext()
								Wend
										
							End If
						   %>
                        </select>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="4">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td colspan="4"> 
                        <div align="center"> 
                          <input type="button" value="Eliminar" name="Eliminar" class="FormButton" onClick="javascript: submitfrmbaja();">
                        </div>
                      </td>
                    </tr>
                  </table>
                          </form>
   <!----------------- F I N   B A J A   D E   S E C T O R   -------------->
                          <table height="1%" border="0" cellspacing="0" cellpadding="0">
<tr><td>&nbsp;
</td></tr>
</table>
<!----------------- M O D I F I C A C I O N   D E   S E C T O R   -------------->
                     <form name="modificacionsector" action="" method="POST">
<input name="accion" id="accion" value="M" type="hidden">
                  <table width="100%" border="0" class="FormBoxHeader">
                    <tr>
                      <td width="5"><img src="images/gripgray.gif" width="10" height="13"></td>
                      <td> 
                        <p>Modificacion</p>
                      </td>
                    </tr>
                  </table>
                                         
                <table border="0" width="100%" class="FormBoxBody">
                  <tr> 
                    <td width="23%">&nbsp;</td>
                    <td width="26%">&nbsp;</td>
                    <td width="26%">&nbsp;</td>
                    <td width="25%">&nbsp;</td>
                  </tr>
                  <tr> 
                    <td width="23%">&nbsp;&nbsp;&nbsp;Seleccione la Ciudad:</td>
                    <td width="26%"> 
                      <select size="1" name="ciudadesmod" class="FormTable" onChange="javascript: reload2();">
                        <option value=0>----</option>
                        <%
While (NOT Rs_ciudad.EOF)
%>
                        <option value="<% =Rs_ciudad.Fields.Item("IDCiudad").Value%>"
<% if (CStr(request.form("ciudadesmod")) = CStr(Rs_ciudad.Fields.Item("IDCiudad").Value)) then Response.Write("SELECTED") : Response.Write("") %> > 
                        <% =Rs_ciudad.Fields.Item("Ciudad").Value %>
                        </option>
                        <%               
  Rs_ciudad.MoveNext()
Wend
%>
                      </select>
                    </td>
                    <td width="26%">Seleccione el Sector:</td>
                    <td width="25%"> 
                      <select size="1" name="sectormod" class="FormTable">
                        <option value=0>----</option>
                        <%
If Request.Form("ciudadesmod")<> 0  Then

			While (NOT Rs_Sectormod.EOF)
			
		Response.Write (" <option value=""" )
		Response.Write (Rs_Sectormod.Fields.Item("IDSector").Value & """> ")
		Response.Write (Rs_Sectormod.Fields.Item("Sector").Value )
		Response.Write ("</option>")
											
			  Rs_Sectormod.MoveNext()
			Wend
			If (Rs_Sectormod.CursorType > 0) Then
			 Rs_Sectormod.MoveFirst
			Else
			 Rs_Sectormod.Requery
			End If
End If
%>
                      </select>
                    </td>
                  </tr>
                  <tr> 
                    <td width="23%">&nbsp;&nbsp;&nbsp;Nombre del nuevo &nbsp;&nbsp;&nbsp;Sector:</td>
                    <td width="26%"> 
                      <input type="text" name="sectoresmod" class="formtable" value="" size="20" maxlength="20">
                    </td>
                    <td width="26%">&nbsp;</td>
                    <td width="25%">&nbsp;</td>
                  </tr>
                  <tr> 
                    <td width="23%">&nbsp;</td>
                    <td width="26%">&nbsp;</td>
                    <td width="26%">&nbsp;</td>
                    <td width="25%">&nbsp;</td>
                  </tr>
                  <tr> 
                    <td colspan="4"> 
                      <div align="center"> 
                        <input type="button" value="Modificar" name="Modificar" class="FormButton" onClick="javascript: submitfrmmodif();">

                      </div>
                    </td>
                  </tr>
                </table>
</form>
 </table>

             
        </center>
      </div>
      <!--#INCLUDE file="library/PageClose.asp"-->
<%
Rs_ciudad.Close()
If Request.Form ("ciudadesbaja") <> 0 Then
Rs_Sectorbaja.Close()
End if
If Request.Form ("ciudadesmod") <> 0 Then
Rs_Sectormod.Close()
End if
%>