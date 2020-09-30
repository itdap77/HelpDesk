<%@ LANGUAGE=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
<%
'ValidUserAction "ABM", "ABM"

set Rs_Ciudad = Server.CreateObject("ADODB.Recordset")
Rs_Ciudad.ActiveConnection = MM_HelpDesk_STRING
Rs_Ciudad.Source = "{call dbo.SPU_Ciudad_VE}"
Rs_Ciudad.Open()

If Request.Form ("ciudades") <> 0 Then
	set Rs_Usuario = Server.CreateObject("ADODB.Recordset")
	Rs_Usuario.ActiveConnection = MM_HelpDesk_STRING
	Rs_Usuario.Source = "{call dbo.SPU_IDCUsuario_L(" + Replace(Request.Form("ciudades"),"'","''") + ")}"
	Rs_Usuario.Open()
End if
%>

      <script type="text/javascript">
function reload()
{
bajausuario.submit();
}

function submitform()
{
if (confirm ("Esta Seguro de borrar el usuario seleccionado??"))
	{	bajausuario.action = "vbsusuario.asp" ;
		bajausuario.submit();
	}
}
</script>
      <div align="center"> 
        <center>
          <table class="TopMenuArea" id="PageBorder" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td class="Titulo">Baja de Usuarios</td>
            </tr>
          </table>
          <table class="ContentArea" border="0" width="100%" cellspacing="0" cellpadding="0" height="100%">
            <tr> 
              <td width="100%"> 
                <table border="0" width="100%" cellspacing="0" cellpadding="0" height="300">
                  <tr> 
                    <td></td>
                  </tr>
                  <tr> 
                    <td colspan="4"> 
                      <div align="center"> 
                        <form name="bajausuario" action="" method="POST">
                        <input type="hidden" name="accion" value="B">
                    
<table  class="formboxheader">
                            <tr> 
                              <td width="5"><img src="images/gripgray.gif" width="10" height="13"></td>
                              <td>Seleccione el usuario a borrar.</td>
                            </tr>

                          </table>
                          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="FormBoxBody">
                            <tr>
                              <td>&nbsp;</td>
                              <td>&nbsp;</td>
                            </tr>
                            <tr> 
                              <td width="50%"><div align="right">Seleccione la Ciudad:</div></td>
                              <td width="70%"> 
                                <select name="ciudades" class="FormTable" onchange="javascript:reload();">
                                  <option value="0">----</option>
                        <%
While (NOT Rs_Ciudad.EOF)
%>
                        <option value="<%=(Rs_Ciudad.Fields.Item("IDCiudad").Value)%>"<%if (CStr(request.form("ciudades")) = CStr(Rs_Ciudad.Fields.Item("IDCiudad").Value)) then Response.Write("SELECTED") : Response.Write("")%>><%=(Rs_Ciudad.Fields.Item("Ciudad").Value)%></option>
                        <%
  Rs_Ciudad.MoveNext()
Wend
If (Rs_Ciudad.CursorType > 0) Then
  Rs_Ciudad.MoveFirst
Else
  Rs_Ciudad.Requery
End If
%>

                                </select>
                              </td>
                            </tr>
                            <tr> 
                              <td width="50%"><div align="right">Seleccione el Usuario:</div></td>
                              <td width="70%"> 
                                <select name="usuarios" class="FormTable">
                                  
                                  <option selected value="0">----</option>
<%
If lcase(Session("NombreUsuario")) <> "demo" Then 
        If Request.Form("ciudades")<>0  Then

			        While (NOT Rs_Usuario.EOF)
			
		        Response.Write (" <option value=""" )
		        Response.Write (Rs_Usuario.Fields.Item("login").Value & """> ")
		        Response.Write (Rs_Usuario.Fields.Item("login").Value )
		        Response.Write ("</option>")
											
			          Rs_Usuario.MoveNext()
			        Wend
			        If (Rs_Usuario.CursorType > 0) Then
			         Rs_Usuario.MoveFirst
			        Else
			         Rs_Usuario.Requery
			        End If
        End If
End If
%>
                                </select>
                              </td>
                            </tr>
                            <tr>
                              <td colspan="2">&nbsp;</td>
                            </tr>
                            <tr>
                              
                                  <% If lcase(Session("NombreUsuario")) <> "demo" Then %>
                                      <td colspan="2"><div align="center">
                                      <input type="button" value="Eliminar" name="Eliminar" class="formbutton" onClick="javascript: submitform();">
                                  <%Else%>
                                          <td width="20%" colspan="6" style="color:red;text-align:center;font-weight:bold;">
                                          <% Call ReadLang(Rs_Lenguaje,204) %>
                                  <%End If%>
                        </td>
                                  
                              </div></td>
                            </tr>
                          </table>
                        </form>
                      </div>
                    </td>
                  </tr>
                </table>
          </table>
        </center>
      </div>
      <!--#INCLUDE file="library/PageClose.asp"-->
<%
Rs_Ciudad.Close()
If Request.Form ("ciudades") <> 0 Then
Rs_Usuario.Close()
End if
%>