<%@ LANGUAGE=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
<%

If (Session("Login") = 1) Then
else
	Response.Redirect "Mensaje.asp?Msg=No esta logueado en el sistema.&BackToURL=default.asp"
End If

ValidUserAction "ABM", "ABM"
%>
<%
set Rs_Ciudad = Server.CreateObject("ADODB.Recordset")
Rs_Ciudad.ActiveConnection = MM_HelpDesk_STRING
Rs_Ciudad.Source = "{call dbo.SPU_Ciudad_VE}"
Rs_Ciudad.Open()

If Request.Form ("ciudades") <> 0 Then
set Rs_Proveedores = Server.CreateObject("ADODB.Recordset")
Rs_Proveedores.ActiveConnection = MM_HelpDesk_STRING
Rs_Proveedores.Source = "{call dbo.SPU_IDCProveedor_L(" + Replace(Request.Form("ciudades"),"'","''") + ")}"
Rs_Proveedores.Open()
End if
%>

      <script languague="javascript" language="JavaScript">
function reload()
{
bajaproveedor.submit();
}

function submitform()
{
	if (confirm ("Esta Seguro de borrar el inventario seleccionado??"))
	{	bajaproveedor.action = "vbsproveedor.asp" ;
		bajaproveedor.submit();
	}

}
</script>
      <div align="center"> 
        <center>
          <table class="TopMenuArea" id="PageBorder" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td class="Titulo">Baja de Proveedores</td>
            </tr>
          </table>
          <table class="ContentArea" border="0" width="100%" cellspacing="0" cellpadding="0" height="100%">
            <tr> 
              <td width="100%"> 
                <table border="0" width="100%" cellspacing="0" cellpadding="0" height="100%">
                  <tr> 
                    <td></td>
                  </tr>
                  <tr> 
                    <td colspan="4"> 
                      <div align="center"> 
                        <form name="bajaproveedor" action="" method="POST">
<input type="hidden" name="accion" value="B">
                          <table width="100%" height="8%" border="0">
                            <tr>
                              <td align="center">&nbsp; </td>
                            </tr>
                          </table>
                          <table width="100%" border="0" class="FormBoxHeader">
                            <tr>
                              <td width="5" align="center"><img src="images/gripgray.gif" width="10" height="13"></td>
                              <td align="center"><div align="left">Seleccione el proveedor a elinimar</div></td>
                            </tr>
                          </table>
                          <table width="100%" border="0" class="FormBoxBody">
                            <tr>
                              <td>&nbsp;</td>
                              <td>&nbsp;</td>
                            </tr>
                            <tr> 
                              <td width="30%">Seleccione la Ciudad:</td>
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
                              <td width="30%">Seleccione el Proveedor:</td>
                              <td width="70%"> 
                                <select name="idproveedor" class="FormTable" id="idproveedor">
                                  <option selected value="0">----</option>
<%
If Request.Form("ciudades")<>0  Then

			While (NOT Rs_Proveedores.EOF)
			
		Response.Write (" <option value=""" )
		Response.Write (Rs_Proveedores.Fields.Item("IDProveedores").Value & """> ")
		Response.Write (Rs_Proveedores.Fields.Item("Proveedor").Value )
		Response.Write ("</option>")
											
			  Rs_Proveedores.MoveNext()
			Wend
			If (Rs_Proveedores.CursorType > 0) Then
			 Rs_Proveedores.MoveFirst
			Else
			 Rs_Proveedores.Requery
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
                              <td colspan="2"><div align="center">
                                <input type="button" value="Eliminar" name="Enviar" class="FormButton" onClick="javascript: submitform();">
                                <input type="reset" name="Limpiar" value="Limpiar" class="FormButton">
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
Rs_Proveedores.Close()
End if
%>