<%@ LANGUAGE=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
<%

If (Session("Login") = 1) Then
else
	Response.Redirect "Mensaje.asp?Msg=No esta logueado en el sistema.&BackToURL=default.asp"
End If

ValidUserAction "ABM", "ABM"

set Rs_Ciudad = Server.CreateObject("ADODB.Recordset")
Rs_Ciudad.ActiveConnection = MM_HelpDesk_STRING
Rs_Ciudad.Source = "{call dbo.SPU_Ciudad_VE}"
Rs_Ciudad.Open()

If Request.Form ("ciudades") <> 0 Then
set Rs_Inventario_Ciudad = Server.CreateObject("ADODB.Recordset")
Rs_Inventario_Ciudad.ActiveConnection = MM_HelpDesk_STRING
Rs_Inventario_Ciudad.Source = "{call dbo.SPU_IDInventarioXCiudad_L(" + Replace(Request.Form("ciudades"),"'","''") + ")}"
Rs_Inventario_Ciudad.Open()
End if
%>

<script languague="javascript" language="JavaScript">
function reload()
{
bajainventario.submit();
}

function submitform()
{
	if (confirm ("Esta Seguro de borrar el inventario seleccionado??"))
	{	bajainventario.action = "vbsinventario.asp" ;
		bajainventario.submit();
	}


}
</script>
    
     
	  <table class="TopMenuArea" id="PageBorder" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td class="Titulo">Baja de Inventario</td>
            </tr>
      </table>

              
<table class="ContentArea" border="0" width="100%" cellspacing="0" cellpadding="0" height="100%">
  <form name="bajainventario" action="" method="POST">
  <input type="hidden" name="accion" value="B">
            <tr> 
              <td width="100%"> <table width="100%" height="8%" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td>
                <p>&nbsp;</p>
              </td>
            </tr>
          </table>     
              <table width="100%" border="0" class="FormBoxHeader">
                <tr>
                  <td width="5">
                    <p><img src="images/gripgray.gif" width="10" height="13"></p>
                  </td>
                  <td>Seleccione el inventario a eliminar.</td>
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
                  <td><div align="right">Seleccione el Equipo:</div>
                  </td>
                  <td width="70%">
                    <select name="inventario" class="FormTable">
                      <option selected value="0">----</option>
                    <%
					If Request.Form("ciudades")<>0  Then
						While (NOT Rs_Inventario_Ciudad.EOF)
							Response.Write (" <option value=""" )
							Response.Write (Rs_Inventario_Ciudad.Fields.Item("IDInventario").Value & """> ")
							Response.Write (Rs_Inventario_Ciudad.Fields.Item("IDPc").Value )
							Response.Write ("</option>")
						    Rs_Inventario_Ciudad.MoveNext()
						Wend
						If (Rs_Inventario_Ciudad.CursorType > 0) Then
						 Rs_Inventario_Ciudad.MoveFirst
						Else
						 Rs_Inventario_Ciudad.Requery
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
                      <input type="button" value="Eliminar" name="Enviar" class="FormButton" onclick="javascript: submitform();">
                      <input type="reset" name="Limpiar" value="Limpiar" class="FormButton">
                    </div>
                  </td>
                </tr>

              </table>
              </form>          
    </table>
      <!--#INCLUDE file="library/PageClose.asp"-->
<%
Rs_Ciudad.Close()
If Request.Form ("ciudades") <> 0 Then
  Rs_Inventario_Ciudad.Close()
End if
%>