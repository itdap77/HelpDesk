<%@ LANGUAGE=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
<%

If (Session("Login") = 1) Then
else
	Response.Redirect "Mensaje.asp?Msg=No esta logueado en el sistema.&BackToURL=default.asp"
End If
ValidUserAction "ABM", "ABM"

set Rs_Estado = Server.CreateObject("ADODB.Recordset")
Rs_Estado.ActiveConnection = MM_HelpDesk_STRING
Rs_Estado.Source = "{call dbo.SPU_Estado_VE}"
Rs_Estado.Open()

%>


<script languague="javascript" language="JavaScript">
function submitfrmalta()
{
altaestado.action = "vbsestado.asp" ;
altaestado.submit();

}

function submitfrmbaja()
{
bajaestado.action = "vbsestado.asp" ;
bajaestado.submit();

}

function submitfrmmodif()
{
modificacionestado.action = "vbsestado.asp" ;
modificacionestado.submit();

}
</script>
    <table class="TopMenuArea" id="PageBorder" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td class="Titulo">ABM de Estados de Tickets</td>
            </tr>
          </table>
          <table class="ContentArea" border="0" width="100%" cellspacing="0" cellpadding="0" height="100%">
            <tr> 
              <td width="100%"> 
                                     
                        <form name="altaestado" method="POST" action="">
                        <input type="hidden" name="accion" value="A">
              <table width="100%" border="0" height="8%">
                <tr> 
                  <td>&nbsp; </td>
                </tr>
              </table>
              <table width="100%" border="0" class="FormBoxHeader">
                            <tr>
                              <td width="5"><img src="images/gripgray.gif" width="10" height="13"></td>
                              <td>Alta</td>
                            </tr>
                          </table>
                          <table width="100%" border="0" class="FormBoxBody">
                            <tr> 
                              <td colspan="2">&nbsp;</td>
                            </tr>
                            <tr> 
                              <td width="40%">&nbsp;&nbsp;&nbsp;&nbsp;Nombre del 
                                nuevo Estado:</td>
                              <td> 
                                <input type="text" name="alta_estado" class="FormTable" size="20" maxlength="20">
                              </td>
                            </tr>
                            <tr> 
                              <td colspan="2">&nbsp;</td>
                            </tr>
                            <tr> 
                              <td colspan="2"> 
                                <div align="center"> 
                                  <input type="button" value="Ingresar" name="Ingresar" class="FormButton" onClick="javascript: submitfrmalta();">
                                  <input type="reset" name="Limpiar2" value="Limpiar" class="FormButton">
                                </div>
                              </td>
                            </tr>
                          </table>
                          <table width="100%" border="0" height="1%">
                            <tr> 
                              <td>&nbsp; </td>
                            </tr>
                          </table>
                        </form>
                        <form name="bajaestado" action="" method="POST">
                        <input type="hidden" name="accion" value="B">
              <table width="100%" border="0" class="FormBoxHeader">
                <tr> 
                  <td width="5"><img src="images/gripgray.gif" width="10" height="13"></td>
                  <td>Baja</td>
                </tr>
              </table>
              <table width="100%" border="0" class="FormBoxBody">
                            <tr> 
                              <td colspan="2">&nbsp;</td>
                            </tr>
                            <tr> 
                              <td width="40%">&nbsp;&nbsp;&nbsp;&nbsp;Seleccione 
                                el Estado:</td>
                              <td> 
                                <select name="estados_baja" class="FormTable">
                                  <Option value="0">----</option>
                                  <%
									While (NOT Rs_Estado.EOF)
									%>
									<option value="<%=(Rs_Estado.Fields.Item("Estado").Value)%>"><%=(Rs_Estado.Fields.Item("Estado").Value)%></option>
									<%
									  Rs_Estado.MoveNext()
									Wend

									%>
									                                </select>
                              </td>
                            </tr>
                            <tr> 
                              <td colspan="2">&nbsp; </td>
                            </tr>
                            <tr> 
                              <td colspan="2"> 
                                <div align="center"> 
                                  <input type="button" value="Eliminar" name="Eliminar" class="FormButton" onClick="javascript: submitfrmbaja();">
                                </div>
                              </td>
                            </tr>
                          </table>
                          <table width="100%" border="0" height="1%">
                            <tr> 
                              <td>&nbsp; </td>
                            </tr>
                          </table>
                        </form>
                        <form name="modificacionestado" action="" method="POST">
                        <input type="hidden" name="accion" value="M">
                          <table width="100%" border="0" class="FormBoxHeader">
                            <tr>
                              <td width="5"><img src="images/gripgray.gif" width="10" height="13"></td>
                              
                  <td>Modificacion</td>
                            </tr>
                          </table>
                          
              <table width="100%" border="0" class="FormBoxBody">
                <tr> 
                  <td colspan="2">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="30%">&nbsp;&nbsp;&nbsp;&nbsp;Seleccione el Estado:</td>
                  <td width="70%"> 
                    <select name="estados_modificacion" class="FormTable">
                      <Option selected value="0">----</option>
                      <%
	                                If  NOT Rs_Estado.BOF Then
										Rs_Estado.MoveFirst()
									End if
								While (NOT Rs_Estado.EOF)
								%>
                      <option value="<%=(Rs_Estado.Fields.Item("IDEstado").Value)%>"><%=(Rs_Estado.Fields.Item("Estado").Value)%></option>
                      <%
								  Rs_Estado.MoveNext()
								Wend

								%>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td width="30%">&nbsp;&nbsp;&nbsp;&nbsp;Nombre del nuevo Grupo:</td>
                  <td width="70%"> 
                    <input type="text" name="modificacion_estado" class="FormTable" size="20" maxlength="20">
                  </td>
                </tr>
                <tr> 
                  <td colspan="2">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="2"> 
                    <div align="center"> 
                      <input type="button" value="Modificar" name="Modificar" class="FormButton" onClick="javascript: submitfrmmodif();">
                      <input type="reset" name="Limpiar" value="Limpiar" class="FormButton">
                    </div>
                  </td>
                </tr>
              </table>
                          </form>
                      </td></tr></table>
      <!--#INCLUDE file="library/PageClose.asp"-->
<%
Rs_Estado.Close()
%>