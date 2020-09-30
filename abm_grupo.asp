<%@ LANGUAGE=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
<%

If (Session("Login") = 1) Then
else
	Response.Redirect "Mensaje.asp?Msg=No esta logueado en el sistema.&BackToURL=default.asp"
End If
ValidUserAction "ABM", "ABM"

set Rs_Grupo = Server.CreateObject("ADODB.Recordset")
Rs_Grupo.ActiveConnection = MM_HelpDesk_STRING
Rs_Grupo.Source = "{call dbo.SPU_Grupo_VE(" & 1 & "," & 0 & ")}"
Rs_Grupo.Open()

%>


<script languague="javascript" language="JavaScript">
function submitfrmalta()
{
altagrupo.action = "vbsgrupo.asp" ;
altagrupo.submit();

}

function submitfrmbaja()
{
if (confirm("Esta Seguro de borrar el grupo seleccionado?"))
{
	bajagrupo.action = "vbsgrupo.asp" ;
	bajagrupo.submit();
}
}

function submitfrmmodif()
{
modificaciongrupo.action = "vbsgrupo.asp" ;
modificaciongrupo.submit();

}
</script>

      <div align="center"> 
        <center>
         <table class="TopMenuArea" id="PageBorder" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td class="Titulo">ABM de Grupos de usuarios</td>
            </tr>
          </table>
          <table class="ContentArea" border="0" width="100%" cellspacing="0" cellpadding="0" height="100%">
            <tr> 
              <td width="100%"> 
                
                        <table width="100%" border="0" height="1%">
                          <tr> 
                            <td>&nbsp; </td>
                          </tr>
                        </table>
                        <form name="altagrupo" method="POST" action="">
                        <input name="accion" type="hidden" value="A">
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
                              <td colspan="3">&nbsp;</td>
                            </tr>
                            <tr> 
                              <td width="40%">&nbsp;&nbsp;&nbsp;&nbsp;Nombre del 
                                nuevo Grupo:</td>
                              <td colspan="2"> 
                                <input type="text" name="alta_grupo" class="FormTable" size="20" maxlength="20">
                              </td>
                            </tr>
                            <tr> 
                              <td colspan="3">&nbsp;</td>
                            </tr>
                            <tr> 
                              <td colspan="3"> 
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
                        <form name="bajagrupo" action="" method="POST">
                        <input name="accion" type="hidden" value="B">
                          <table width="100%" border="0" class="FormBoxHeader">
                            <tr> 
                              <td width="5"><img src="images/gripgray.gif" width="10" height="13"></td>
                              <td> 
                                <p>Baja</p>
                              </td>
                            </tr>
                          </table>
                          <table border="0" class="FormBoxBody" width="100%" cellpadding="0" cellspacing="2">
                            <tr> 
                              <td colspan="3">&nbsp;</td>
                            </tr>
                            <tr> 
                              <td width="40%">&nbsp;&nbsp;&nbsp;&nbsp;Seleccione 
                                el Grupo:</td>
                              <td colspan="2"> 
                                <select name="grupos_baja" class="FormTable">
                                  <Option selected value="----">----</option>
                                  <%
While (NOT Rs_Grupo.EOF)
%>
                                  <option value="<%=(Rs_Grupo.Fields.Item("Grupo").Value)%>"><%=(Rs_Grupo.Fields.Item("Grupo").Value)%></option>
                                  <%
  Rs_Grupo.MoveNext()
Wend
If (Rs_Grupo.CursorType > 0) Then
  Rs_Grupo.MoveFirst
Else
  Rs_Grupo.Requery
End If
%>
                                </select>
                              </td>
                            </tr>
                            <tr> 
                              <td colspan="3">&nbsp;</td>
                            </tr>
                            <tr> 
                              <td colspan="3"> 
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
                        <form name="modificaciongrupo" action="" method="POST">
                        <input name="accion" type="hidden" value="M">
                          <table width="100%" border="0" class="FormBoxHeader">
                            <tr> 
                              <td width="5"><img src="images/gripgray.gif" width="10" height="13"></td>
                              <td> 
                                <p>Modificacion</p>
                              </td>
                            </tr>
                          </table>
                          <table border="0" class="FormBoxBody">
                            <tr> 
                              <td width="40%">&nbsp;</td>
                              <td>&nbsp;</td>
                            </tr>
                            <tr> 
                              <td>&nbsp;&nbsp;&nbsp;&nbsp;Seleccione el Grupo:</td>
                              <td> 
                                <select name="grupos_modificacion" class="FormTable">
                                  <Option selected value="----">----</option>
                                  <%
While (NOT Rs_Grupo.EOF)
%>
                                  <option value="<%=(Rs_Grupo.Fields.Item("IDGrupo").Value)%>"><%=(Rs_Grupo.Fields.Item("Grupo").Value)%></option>
                                  <%
  Rs_Grupo.MoveNext()
Wend
If (Rs_Grupo.CursorType > 0) Then
  Rs_Grupo.MoveFirst
Else
  Rs_Grupo.Requery
End If
%>
                                </select>
                              </td>
                            </tr>
                            <tr> 
                              <td>&nbsp;&nbsp;&nbsp;&nbsp;Nombre del nuevo Grupo:</td>
                              <td> 
                                <input type="text" name="modificacion_grupo" class="FormTable" size="20" maxlength="20">
                              </td>
                            </tr>
                            <tr> 
                              <td>&nbsp;</td>
                              <td>&nbsp;</td>
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
Rs_Grupo.Close()
%>