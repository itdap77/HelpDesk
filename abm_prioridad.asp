<%@ LANGUAGE=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
<%

If (Session("Login") = 1) Then
else
	Response.Redirect "Mensaje.asp?Msg=No esta logueado en el sistema.&BackToURL=default.asp"
End If

ValidUserAction "ABM", "ABM"

set Rs_Prioridad = Server.CreateObject("ADODB.Recordset")
Rs_Prioridad.ActiveConnection = MM_HelpDesk_STRING
Rs_Prioridad.Source = "{call dbo.SPU_Prioridad_VE}"
Rs_Prioridad.Open()

%>


<script languague="javascript" language="JavaScript">
function submitfrmalta()
{
altaprioridad.accion.value = "A";
altaprioridad.action = "vbsprioridad.asp" ;
altaprioridad.submit();

}

function submitfrmbaja()
{
if (confirm("Esta Seguro de borrar la prioridad seleccionada?"));
{	bajaprioridad.accion.value = "B";
	bajaprioridad.action = "vbsprioridad.asp" ;
	bajaprioridad.submit();
}
}

function submitfrmmodif()
{
modificacionprioridad.accion.value = "M";
modificacionprioridad.action = "vbsprioridad.asp" ;
modificacionprioridad.submit();

}
</script>

      <div align="center"> 
        <center>
          <table class="TopMenuArea" id="PageBorder" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td class="Titulo">ABM de Prioridades</td>
            </tr>
          </table>
          <table class="ContentArea" border="0" width="100%" cellspacing="0" cellpadding="0" height="100%">
            <tr> 
              <td width="100%"> 
               
                        <form name="altaprioridad" method="POST" action="">
                        <input type="hidden" name="accion" value="">
                  <table width="100%" border="0" height="8%">
                    <tr> 
                      <td>&nbsp; </td>
                    </tr>
                  </table>
                  <table width="100%" border="0" class="FormBoxHeader">
                            <tr>
                              <td width="5"><img src="images/gripgray.gif" width="10" height="13"></td>
                              <td> 
                                <p>Alta</p>
                              </td>
                            </tr>
                          </table>
                          <table width="100%" border="0" class="FormBoxBody">
                            <tr> 
                              <td colspan="2">

                      </td>
                            </tr>
                            <tr> 
                              <td width="40%">&nbsp;&nbsp;&nbsp;&nbsp;Nombre de 
                                la nueva Prioridad:</td>
                              <td> 
                                
                        <input type="text" name="alta_prioridad" class="FormTable" size="20" maxlength="10">
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
                        <form name="bajaprioridad" action="" method="POST">
                        <input type="hidden" name="accion" value="">
                          <table width="100%" border="0" class="FormBoxHeader">
                            <tr>
                              <td width="5"><img src="images/gripgray.gif" width="10" height="13"></td>
                              <td> 
                                <p>Baja</p>
                              </td>
                            </tr>
                          </table>
                          <table width="100%" border="0" class="FormBoxBody">
                            <tr> 
                              <td colspan="2" height="19">&nbsp;</td>
                            </tr>
                            <tr> 
                              <td width="40%">&nbsp;&nbsp;&nbsp;&nbsp;Seleccione 
                                la Prioridad:</td>
                              <td> 
                                <select name="prioridad_baja" class="FormTable">
                                  <Option selected value="----">----</option>
                                  <%
While (NOT Rs_Prioridad.EOF)
%>
                                  <option value="<%=(Rs_Prioridad.Fields.Item("Prioridad").Value)%>"><%=(Rs_Prioridad.Fields.Item("Prioridad").Value)%></option>
                                  <%
  Rs_Prioridad.MoveNext()
Wend
If (Rs_Prioridad.CursorType > 0) Then
  Rs_Prioridad.MoveFirst
Else
  Rs_Prioridad.Requery
End If
%>
                                </select>
                              </td>
                            </tr>
                            <tr> 
                              <td colspan="2">&nbsp;</td>
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
                        <form name="modificacionprioridad" action="" method="POST">
                        <input type="hidden" name="accion" value="">
                          <table width="100%" border="0" class="FormBoxHeader">
                            <tr>
                              <td width="5"><img src="images/gripgray.gif" width="10" height="13"></td>
                              <td> 
                                <p>Modificacion</p>
                              </td>
                            </tr>
                          </table>
                          <table width="100%" border="0" class="FormBoxBody">
                            <tr> 
                              <td colspan="2">&nbsp;</td>
                            </tr>
                            <tr> 
                              <td width="40%">&nbsp;&nbsp;&nbsp;&nbsp;Seleccione 
                                la Prioridad:</td>
                              <td> 
                                <select name="prioridad_modificacion" class="FormTable">
                                  <Option selected value="----">----</option>
                                  <%
While (NOT Rs_Prioridad.EOF)
%>
                                  <option value="<%=(Rs_Prioridad.Fields.Item("IDPrioridad").Value)%>"><%=(Rs_Prioridad.Fields.Item("Prioridad").Value)%></option>
                                  <%
  Rs_Prioridad.MoveNext()
Wend
If (Rs_Prioridad.CursorType > 0) Then
  Rs_Prioridad.MoveFirst
Else
  Rs_Prioridad.Requery
End If
%>
                                </select>
                              </td>
                            </tr>
                            <tr> 
                              <td width="40%">&nbsp;&nbsp;&nbsp;&nbsp;Nombre de 
                                la nueva Prioridad:</td>
                              <td> 
                                
                        <input type="text" name="modificacion_prioridad" class="FormTable" size="20" maxlength="10">
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
                      
          </table>
        </center>
      </div>
      <!--#INCLUDE file="library/PageClose.asp"-->
<%
Rs_Prioridad.Close()
%>