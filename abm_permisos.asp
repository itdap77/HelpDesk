<%@ LANGUAGE=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
<%

If (Session("Login") = 1) Then
else
	Response.Redirect "Mensaje.asp?Msg=No esta logueado en el sistema.&BackToURL=default.asp"
End If
ValidUserAction "ABM", "ABM"

set RS_Grupo = Server.CreateObject("ADODB.Recordset")
RS_Grupo.ActiveConnection = MM_HelpDesk_STRING
RS_Grupo.Source = "{call dbo.SPU_Grupo_V}"
RS_Grupo.Open()

set RS_Usuario = Server.CreateObject("ADODB.Recordset")
RS_Usuario.ActiveConnection = MM_HelpDesk_STRING
RS_Usuario.Source = "{call dbo.SPU_Usuarios_V}"
RS_Usuario.Open()

set RS_Modulo = Server.CreateObject("ADODB.Recordset")
RS_Modulo.ActiveConnection = MM_HelpDesk_STRING
RS_Modulo.Source = "{call dbo.SPU_Modulo_V}"
RS_Modulo.Open()


%>
<script Language="JavaScript">

function reload()
{
AgregarPermisos.submit();
}

function submitfrmmodulo()
{
agregarmodulo.action = "vbspermisos.asp" ;
agregarmodulo.submit();

}

</script>

<body>

<div align="center">
  <center>
  <table class="TopMenuArea" id="PageBorder" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td class="Titulo">ABM de Permisos</td>
    </tr>
  </table>
  <table class="ContentArea" border="0" width="100%" cellspacing="0" cellpadding="0" height="100%">
    <tr>
      <td width="100%"><table width="100%" height="9%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="9%">&nbsp;</td>
        </tr>
      </table>
        <table width="100%" border="0" class="FormBoxHeader">
        <tr>
          <td width="5"><img src="images/gripgray.gif" width="10" height="13"></td>
          <td><font size="2"><b>Asignacion de Permisos</b></font></td>
        </tr>
      </table>
        <table width="102%" border="0" class="FormBoxBody">
          <tr>
            <td width="30%">&nbsp;&nbsp;&nbsp;&nbsp;Seleccione el Grupo</td>
            <td width="30%">
              <select name="IDGrupo" class="FormTable" onChange="javascript:reload();">
                <option value="0">----</option>
                <% While (NOT RS_Grupo.EOF)%>
                <option value="<%=(RS_Grupo.Fields.Item("IDGrupo").Value)%>"<%if (CStr(request.form("IDGrupo")) = CStr(RS_Grupo.Fields.Item("IDGrupo").Value)) then 
					Response.Write("SELECTED") 
				Else 
					Response.Write("")
				End If
				%>>
                <%Response.Write(RS_Grupo.Fields.Item("Grupo").Value )%>
                </option>
                <%
			    RS_Grupo.MoveNext()
				Wend
				If (RS_Grupo.CursorType > 0) Then
				  RS_Grupo.MoveFirst
				Else
			 	 RS_Grupo.Requery
				End If
				
				

%>
              </select>
            </td>
            <td>ó el Usuario:
                <select name="IDUsuario" class="FormTable" onChange="javascript:reload();">
                  <option value="0">----</option>
                  <% While (NOT RS_Usuario.EOF)%>
                  <option value="<%=(RS_Usuario.Fields.Item("IDUsuario").Value)%>"<%if (CStr(request.form("IDUsuario")) = CStr(RS_Usuario.Fields.Item("IDUsuario").Value)) then 
					Response.Write("SELECTED") 
				Else 
					Response.Write("")
				End If
				%>>
                  <%Response.Write(RS_Usuario.Fields.Item("Nombre").Value )%>
                  </option>
                  <%
			    RS_Usuario.MoveNext()
				Wend
				If (RS_Usuario.CursorType > 0) Then
				  RS_Usuario.MoveFirst
				Else
			 	 RS_Usuario.Requery
				End If
				
				

%>
                </select>
            </td>
          </tr>
          <tr>
            <td>&nbsp;&nbsp;&nbsp; Seleccione el Modulo</td>
            <td>
              <select name="IDModulo" class="FormTable" onChange="javascript:reload();">
                <option value="0">----</option>
                <%
While (NOT RS_Modulo.EOF)
%>
                <option value="<%=(RS_Modulo.Fields.Item("IDModulo").Value)%>" <%	if (CStr(request.form("IDModulo")) = CStr(RS_Modulo.Fields.Item("IDModulo").Value)) then 
		Response.Write("SELECTED") 
	Else 
		Response.Write("")
	End If
%>>
                <%Response.Write(RS_Modulo.Fields.Item("Nombre").Value )%>
                </option>
                <%
  RS_Modulo.MoveNext()
Wend
If (RS_Modulo.CursorType > 0) Then
  RS_Modulo.MoveFirst
Else
 RS_Modulo.Requery
End If
%>
              </select>
            </td>
            <td>&nbsp; </td>
          </tr>
          <tr>
            <td> &nbsp;&nbsp;&nbsp; Permisos</td>
            <td><select  name="Permisos" class="FormTable">
                <option value="A">Alta </option>
                <option value="M">Modificacion </option>
                <option value="B">Borrado </option>
              </select>
            </td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td>&nbsp;&nbsp;&nbsp; Estado</td>
            <td><Select name="Estado" class="FormTable">
                <option value="<% 
                      If request.form ("usuarios") <> 0 Then 
                      				If rs_usuario.fields.item("Estado").value = 1 then 
                      					Response.Write "1"  
                      					Response.Write """" 
                      					Response.Write " Selected>" 
                      					Response.Write "Si"
                      					Response.Write "</option>"
                      					Response.Write "<option value=""0"">No"
                      					
                      					
                      		           Else               		       			
										Response.Write "0"  
                      					Response.Write """" 
                      					Response.Write " Selected>" 
                      					Response.Write "No"	
                      					Response.Write "</option>"
                      					Response.Write "<option value=""1"">Si"
                      												  
                      				   End If
                      	End If
									%>
                      </option>
                      </select>
                      </td>
                    </tr>
                    <tr> 
                      <td colspan="3">&nbsp;
              </select>
            </td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td>&nbsp;&nbsp;&nbsp; Fecha Estado</td>
            <td><input type="text" name="FechaEstado" size="20">
            </td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td>&nbsp;&nbsp;&nbsp; IDusuariomod</td>
            <td>
              <input type="text" name="IDUsuarioMod" size="20">
            </td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td colspan="3">              <div align="center">
                <input type="button" value="Modificar" name="Enviar" class="formbutton" onclick="javascript: submitfrmmodif();">
                <input type="reset" name="Limpiar" value="Limpiar" class="formbutton">
              </div>
            </td>
          </tr>
        </table>
</td>
    </tr>
  </table>
  </center>
</div>
<!--#INCLUDE file="library/PageClose.asp"-->

<%
RS_Grupo.Close()
%>